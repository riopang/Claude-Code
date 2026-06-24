#!/usr/bin/env python3
import os, sys, shutil, subprocess, json, queue, threading, time
from pathlib import Path
from datetime import datetime
from flask import Flask, render_template, request, jsonify, Response, stream_with_context

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 200MB

BASE        = Path(__file__).parent
UPLOADS     = BASE / "uploads"
ARCHIVE     = Path.home() / "Desktop" / "Property Archive"
UPLOADS.mkdir(exist_ok=True)
ARCHIVE.mkdir(exist_ok=True)

# ── SSE progress queue ────────────────────────────────────────────────────────
_progress_queue: queue.Queue = queue.Queue()

def emit(msg, type="log"):
    _progress_queue.put(json.dumps({"msg": msg, "type": type}))

# ── ffmpeg helper ─────────────────────────────────────────────────────────────
def get_ffmpeg():
    try:
        r = subprocess.run(["ffmpeg", "-version"], capture_output=True)
        if r.returncode == 0:
            return "ffmpeg"
    except FileNotFoundError:
        pass
    try:
        import imageio_ffmpeg
        return imageio_ffmpeg.get_ffmpeg_exe()
    except ImportError:
        return None

# ── Photo enhancement ─────────────────────────────────────────────────────────
def enhance_photos(src_files, out_dir):
    from PIL import Image, ImageEnhance, ImageFilter
    try:
        import pillow_heif
        pillow_heif.register_heif_opener()
    except ImportError:
        pass

    out_dir.mkdir(parents=True, exist_ok=True)
    enhanced = []
    for i, src in enumerate(src_files, 1):
        emit(f"Enhancing [{i}/{len(src_files)}] {Path(src).name}…")
        try:
            img = Image.open(src).convert("RGB")

            # Upscale to 2K
            w, h = img.size
            tw, th = 2560, 1440
            if w < tw or h < th:
                ratio = max(tw / w, th / h)
                img = img.resize((int(w * ratio), int(h * ratio)), Image.LANCZOS)

            # Enhance
            img = ImageEnhance.Brightness(img).enhance(1.05)
            img = ImageEnhance.Contrast(img).enhance(1.10)
            img = ImageEnhance.Color(img).enhance(1.10)
            img = ImageEnhance.Sharpness(img).enhance(1.30)
            img = img.filter(ImageFilter.UnsharpMask(radius=1, percent=120, threshold=3))

            # Crop to 16:9
            w2, h2 = img.size
            if abs(w2 / h2 - 16 / 9) > 0.05:
                if w2 / h2 > 16 / 9:
                    nw = int(h2 * 16 / 9)
                    img = img.crop(((w2 - nw) // 2, 0, (w2 - nw) // 2 + nw, h2))
                else:
                    nh = int(w2 * 9 / 16)
                    img = img.crop((0, (h2 - nh) // 2, w2, (h2 - nh) // 2 + nh))

            out = out_dir / (Path(src).stem + "_edited.jpg")
            img.save(out, "JPEG", quality=92, optimize=True)
            enhanced.append(out)
            emit(f"✓ {Path(src).name}", "success")
        except Exception as e:
            emit(f"✗ {Path(src).name}: {e}", "error")
    return enhanced

# ── Ken Burns reel ────────────────────────────────────────────────────────────
def make_reel(photo, out_path, duration=5):
    ff = get_ffmpeg()
    if not ff:
        return False
    cmd = [
        ff, "-y", "-loop", "1", "-i", str(photo),
        "-vf", (
            f"scale=2560:1440:force_original_aspect_ratio=increase,"
            f"crop=2560:1440,"
            f"zoompan=z='min(zoom+0.0015,1.5)':x='iw/2-(iw/zoom/2)':y='ih/2-(ih/zoom/2)'"
            f":d={duration*25}:s=1280x720:fps=25,fps=25"
        ),
        "-t", str(duration), "-c:v", "libx264", "-pix_fmt", "yuv420p", "-preset", "fast",
        str(out_path)
    ]
    r = subprocess.run(cmd, capture_output=True)
    return r.returncode == 0

def stitch_reels(reel_files, out_path):
    ff = get_ffmpeg()
    if not ff or not reel_files:
        return False
    concat = out_path.parent / "_concat.txt"
    with open(concat, "w") as f:
        for r in sorted(reel_files):
            f.write(f"file '{r}'\n")
    cmd = [ff, "-y", "-f", "concat", "-safe", "0", "-i", str(concat), "-c", "copy", str(out_path)]
    result = subprocess.run(cmd, capture_output=True)
    concat.unlink(missing_ok=True)
    return result.returncode == 0

# ── Main workflow ─────────────────────────────────────────────────────────────
def run_workflow(data):
    try:
        emit("── Starting workflow ──", "heading")

        address     = data["address"]
        postal      = data["postal"]
        prop_type   = data["prop_type"]
        tenure      = data["tenure"]
        floor_area  = data["floor_area"]
        bedrooms    = data["bedrooms"]
        bathrooms   = data["bathrooms"]
        price       = data["price"]
        usps        = data["usps"]
        do_reels    = data.get("do_reels", False)
        do_agentnet = data.get("do_agentnet", False)

        # Job folder
        safe = address.replace("/", "-").replace(" ", "_")[:55]
        ts   = datetime.now().strftime("%Y%m%d_%H%M")
        job  = ARCHIVE / f"{safe}_{ts}"
        job.mkdir(parents=True, exist_ok=True)
        emit(f"Job folder: {job.name}")

        # Save details
        details_txt = job / "listing_details.txt"
        with open(details_txt, "w") as f:
            f.write("PROPERTY LISTING DETAILS\n" + "=" * 40 + "\n")
            f.write(f"Address     : {address}, S({postal})\n")
            f.write(f"Type        : {prop_type}\n")
            f.write(f"Tenure      : {tenure}\n")
            f.write(f"Floor Area  : {floor_area} sqft\n")
            f.write(f"Bedrooms    : {bedrooms}\n")
            f.write(f"Bathrooms   : {bathrooms}\n")
            f.write(f"Price       : {price}\n")
            f.write(f"USPs:\n")
            for u in usps:
                f.write(f"  • {u}\n")
            f.write(f"\nGenerated: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n")

        # Collect uploaded photos
        src_files = sorted(UPLOADS.glob("*"))
        src_files = [f for f in src_files if f.suffix.lower() in (".jpg",".jpeg",".png",".heic",".webp")]

        if not src_files:
            emit("⚠️ No photos found in uploads. Skipping photo steps.", "warn")
        else:
            # Enhance
            emit(f"\n── Enhancing {len(src_files)} photo(s) ──", "heading")
            edited_dir  = job / "edited_photos"
            enhanced    = enhance_photos(src_files, edited_dir)
            emit(f"✅ {len(enhanced)} photos enhanced", "success")

            # Reels
            final_video = None
            if do_reels and enhanced:
                emit(f"\n── Creating {len(enhanced)} reel(s) ──", "heading")
                reels_dir = job / "reels"
                reels_dir.mkdir(exist_ok=True)
                reel_files = []
                for i, ph in enumerate(enhanced, 1):
                    out = reels_dir / (ph.stem.replace("_edited","") + "_reel.mp4")
                    emit(f"Rendering reel [{i}/{len(enhanced)}] {ph.name}…")
                    ok = make_reel(ph, out)
                    if ok:
                        reel_files.append(out)
                        emit(f"✓ {out.name}", "success")
                    else:
                        emit(f"✗ Failed: {ph.name}", "error")

                if reel_files:
                    emit("\n── Stitching final video ──", "heading")
                    safe_addr   = address.replace("/","-").replace(" ","_")[:50]
                    final_video = job / f"{safe_addr}_final.mp4"
                    ok = stitch_reels(reel_files, final_video)
                    if ok:
                        emit(f"✅ Final video ready: {final_video.name}", "success")
                    else:
                        emit("✗ Stitching failed", "error")
                        final_video = None

            # Archive originals
            emit("\n── Archiving originals ──", "heading")
            orig_dir = job / "original_photos"
            orig_dir.mkdir(exist_ok=True)
            for f in src_files:
                shutil.move(str(f), orig_dir / f.name)
            emit(f"✅ {len(src_files)} originals archived", "success")

        # AgentNet
        if do_agentnet:
            emit("\n── Opening AgentNet portal ──", "heading")
            subprocess.run(["open", "https://agentnet.propertyguru.com.sg"])
            emit("✅ AgentNet opened in your browser", "success")
            emit(f"   Address     : {address}, S({postal})")
            emit(f"   Type        : {prop_type}")
            emit(f"   Tenure      : {tenure}")
            emit(f"   Floor Area  : {floor_area} sqft")
            emit(f"   Bedrooms    : {bedrooms}")
            emit(f"   Bathrooms   : {bathrooms}")
            emit(f"   Price       : {price}")
            for u in usps:
                emit(f"   • {u}")

        emit("\n── All done! ──", "heading")
        emit(f"📁 Saved to: Property Archive → {job.name}", "success")
        emit("DONE", "done")

    except Exception as e:
        emit(f"❌ Error: {e}", "error")
        emit("DONE", "done")

# ── Routes ────────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload():
    # Clear previous uploads
    for f in UPLOADS.iterdir():
        f.unlink()
    files = request.files.getlist("photos")
    saved = []
    for f in files:
        if f.filename:
            dest = UPLOADS / f.filename
            f.save(dest)
            saved.append(f.filename)
    return jsonify({"saved": saved})

@app.route("/run", methods=["POST"])
def run():
    data = request.json
    # Clear queue
    while not _progress_queue.empty():
        _progress_queue.get_nowait()
    t = threading.Thread(target=run_workflow, args=(data,), daemon=True)
    t.start()
    return jsonify({"status": "started"})

@app.route("/progress")
def progress():
    def generate():
        while True:
            try:
                item = _progress_queue.get(timeout=30)
                yield f"data: {item}\n\n"
                if json.loads(item).get("type") == "done":
                    break
            except queue.Empty:
                yield "data: {\"msg\":\".\",\"type\":\"ping\"}\n\n"
    return Response(stream_with_context(generate()), mimetype="text/event-stream")

@app.route("/archive")
def list_archive():
    jobs = sorted(ARCHIVE.iterdir(), reverse=True) if ARCHIVE.exists() else []
    return jsonify([j.name for j in jobs if j.is_dir()])

if __name__ == "__main__":
    import webbrowser
    print("\n  Property Dashboard running at http://localhost:5050\n")
    threading.Timer(1.0, lambda: webbrowser.open("http://localhost:5050")).start()
    app.run(port=5050, debug=False)
