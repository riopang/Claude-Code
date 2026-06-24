#!/usr/bin/env python3
"""Property Dashboard — stdlib-only HTTP server (no Flask required)."""
import os, sys, json, queue, threading, shutil, subprocess, cgi, email.parser
from pathlib import Path
from datetime import datetime
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse

BASE     = Path(globals().get("__file__", "/tmp/prop_dash_app.py")).parent
TMPL     = Path("/tmp/prop_dash_templates/index.html")
UPLOADS  = Path("/tmp/prop_dashboard_uploads")
ARCHIVE  = Path.home() / "Desktop" / "Property Archive"
UPLOADS.mkdir(exist_ok=True)
ARCHIVE.mkdir(exist_ok=True)

_q: queue.Queue = queue.Queue()
_confirm_event = threading.Event()
_current_job: dict = {}   # holds job folder + listing data for review

def emit(msg, t="log"):
    _q.put(json.dumps({"msg": msg, "type": t}))

# ── ffmpeg ────────────────────────────────────────────────────────────────────
def get_ffmpeg():
    try:
        r = subprocess.run(["ffmpeg", "-version"], capture_output=True)
        if r.returncode == 0: return "ffmpeg"
    except FileNotFoundError: pass
    try:
        import imageio_ffmpeg; return imageio_ffmpeg.get_ffmpeg_exe()
    except ImportError: return None

# ── Photo enhancement ─────────────────────────────────────────────────────────
def enhance_photos(src_files, out_dir):
    try:
        from PIL import Image, ImageEnhance, ImageFilter
    except ImportError:
        subprocess.run([sys.executable, "-m", "pip", "install", "Pillow"], capture_output=True)
        from PIL import Image, ImageEnhance, ImageFilter
    try:
        import pillow_heif; pillow_heif.register_heif_opener()
    except ImportError: pass

    out_dir.mkdir(parents=True, exist_ok=True)
    enhanced = []
    for i, src in enumerate(src_files, 1):
        emit(f"Enhancing [{i}/{len(src_files)}] {Path(src).name}…")
        try:
            img = Image.open(src).convert("RGB")
            w, h = img.size
            tw, th = 2560, 1440
            if w < tw or h < th:
                ratio = max(tw / w, th / h)
                img = img.resize((int(w * ratio), int(h * ratio)), Image.LANCZOS)
            img = ImageEnhance.Brightness(img).enhance(1.05)
            img = ImageEnhance.Contrast(img).enhance(1.10)
            img = ImageEnhance.Color(img).enhance(1.10)
            img = ImageEnhance.Sharpness(img).enhance(1.30)
            img = img.filter(ImageFilter.UnsharpMask(radius=1, percent=120, threshold=3))
            w2, h2 = img.size
            if abs(w2/h2 - 16/9) > 0.05:
                if w2/h2 > 16/9:
                    nw = int(h2*16/9); img = img.crop(((w2-nw)//2, 0, (w2-nw)//2+nw, h2))
                else:
                    nh = int(w2*9/16); img = img.crop((0, (h2-nh)//2, w2, (h2-nh)//2+nh))
            out = out_dir / (Path(src).stem + "_edited.jpg")
            img.save(out, "JPEG", quality=92, optimize=True)
            enhanced.append(out)
            emit(f"✓ {Path(src).name}", "success")
        except Exception as e:
            emit(f"✗ {Path(src).name}: {e}", "error")
    return enhanced

def make_reel(photo, out_path, duration=5):
    ff = get_ffmpeg()
    if not ff: return False
    cmd = [ff, "-y", "-loop", "1", "-i", str(photo),
           "-vf", (f"scale=2560:1440:force_original_aspect_ratio=increase,crop=2560:1440,"
                   f"zoompan=z='min(zoom+0.0015,1.5)':x='iw/2-(iw/zoom/2)':y='ih/2-(ih/zoom/2)'"
                   f":d={duration*25}:s=1280x720:fps=25,fps=25"),
           "-t", str(duration), "-c:v", "libx264", "-pix_fmt", "yuv420p", "-preset", "fast", str(out_path)]
    return subprocess.run(cmd, capture_output=True).returncode == 0

def stitch_reels(reel_files, out_path):
    ff = get_ffmpeg()
    if not ff or not reel_files: return False
    concat = out_path.parent / "_concat.txt"
    with open(concat, "w") as f:
        for r in sorted(reel_files): f.write(f"file '{r}'\n")
    cmd = [ff, "-y", "-f", "concat", "-safe", "0", "-i", str(concat), "-c", "copy", str(out_path)]
    ok = subprocess.run(cmd, capture_output=True).returncode == 0
    concat.unlink(missing_ok=True)
    return ok

# ── Workflow ──────────────────────────────────────────────────────────────────
def run_workflow(data):
    try:
        emit("── Starting workflow ──", "heading")
        address, postal    = data["address"], data["postal"]
        prop_type, tenure  = data["prop_type"], data["tenure"]
        floor_area         = data["floor_area"]
        bedrooms, baths    = data["bedrooms"], data["bathrooms"]
        listing_type, price, usps = data.get("listing_type","For Sale"), data["price"], data["usps"]
        do_reels          = data.get("do_reels", False)
        do_agentnet       = data.get("do_agentnet", False)

        safe = address.replace("/","-").replace(" ","_")[:55]
        ts   = datetime.now().strftime("%Y%m%d_%H%M")
        job  = ARCHIVE / f"{safe}_{ts}"
        job.mkdir(parents=True, exist_ok=True)
        emit(f"Job folder: {job.name}")

        with open(job / "listing_details.txt", "w") as f:
            f.write("PROPERTY LISTING DETAILS\n" + "="*40 + "\n")
            f.write(f"Address     : {address}, S({postal})\n")
            f.write(f"Listing     : {listing_type}\n")
            f.write(f"Type        : {prop_type}\n")
            f.write(f"Tenure      : {tenure}\n")
            f.write(f"Floor Area  : {floor_area} sqft\n")
            f.write(f"Bedrooms    : {bedrooms}\n")
            f.write(f"Bathrooms   : {baths}\n")
            f.write(f"Price       : {price}\n")
            f.write("USPs:\n")
            for u in usps: f.write(f"  • {u}\n")
            f.write(f"\nGenerated: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n")

        src_files = sorted([f for f in UPLOADS.iterdir()
                            if f.suffix.lower() in (".jpg",".jpeg",".png",".heic",".webp")])
        if not src_files:
            emit("⚠️ No photos found. Skipping photo steps.", "warn")
        else:
            emit(f"\n── Enhancing {len(src_files)} photo(s) ──", "heading")
            enhanced = enhance_photos(src_files, job / "edited_photos")
            emit(f"✅ {len(enhanced)} photos enhanced", "success")

            final_video = None
            if do_reels and enhanced:
                emit(f"\n── Creating {len(enhanced)} reel(s) ──", "heading")
                reels_dir = job / "reels"; reels_dir.mkdir(exist_ok=True)
                reel_files = []
                for i, ph in enumerate(enhanced, 1):
                    out = reels_dir / (ph.stem.replace("_edited","") + "_reel.mp4")
                    emit(f"Rendering reel [{i}/{len(enhanced)}] {ph.name}…")
                    if make_reel(ph, out):
                        reel_files.append(out); emit(f"✓ {out.name}", "success")
                    else:
                        emit(f"✗ Failed: {ph.name}", "error")
                if reel_files:
                    emit("\n── Stitching final video ──", "heading")
                    final_video = job / f"{safe}_final.mp4"
                    if stitch_reels(reel_files, final_video):
                        emit(f"✅ Final video ready: {final_video.name}", "success")
                    else:
                        emit("✗ Stitching failed", "error"); final_video = None

            emit("\n── Archiving originals ──", "heading")
            orig_dir = job / "original_photos"; orig_dir.mkdir(exist_ok=True)
            for f in src_files: shutil.move(str(f), orig_dir / f.name)
            emit(f"✅ {len(src_files)} originals archived", "success")

        if do_agentnet:
            emit("\n── Review your listing before posting ──", "heading")
            # Collect edited photo paths for the review panel
            edited_dir = job / "edited_photos"
            photo_names = [f.name for f in sorted(edited_dir.iterdir()) if f.suffix.lower() in (".jpg",".jpeg",".png")] if edited_dir.exists() else []
            _current_job.update({
                "job_path": str(job),
                "photos": photo_names,
                "details": {
                    "address": f"{address}, S({postal})",
                    "listing_type": listing_type,
                    "prop_type": prop_type,
                    "tenure": tenure,
                    "floor_area": floor_area,
                    "bedrooms": bedrooms,
                    "bathrooms": baths,
                    "price": price,
                    "usps": usps,
                }
            })
            emit("SHOW_REVIEW", "review")
            _confirm_event.clear()
            _confirm_event.wait()
            subprocess.run(["open", "https://agentnet.propertyguru.com.sg"])
            emit("✅ AgentNet opened in your browser", "success")
            emit(f"   Address : {address}, S({postal})")
            emit(f"   Type    : {prop_type}  |  Tenure: {tenure}")
            emit(f"   Area    : {floor_area} sqft  |  {bedrooms} bed / {baths} bath")
            emit(f"   Price   : {price}")
            for u in usps: emit(f"   • {u}")

        emit("\n── All done! ──", "heading")
        emit(f"📁 Saved to: Property Archive → {job.name}", "success")
        emit("DONE", "done")
    except Exception as e:
        emit(f"❌ Error: {e}", "error")
        emit("DONE", "done")

# ── HTTP Handler ──────────────────────────────────────────────────────────────
class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a): pass  # silence access log

    def send_json(self, obj, status=200):
        body = json.dumps(obj).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", len(body))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET,POST,OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_GET(self):
        path = urlparse(self.path).path

        if path == "/":
            html = TMPL.read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", len(html))
            self.end_headers()
            self.wfile.write(html)

        elif path == "/review_data":
            self.send_json(_current_job)

        elif path.startswith("/photo/"):
            fname    = path[len("/photo/"):]
            job_path = _current_job.get("job_path", "")
            fpath    = Path(job_path) / "edited_photos" / fname
            if fpath.exists():
                data = fpath.read_bytes()
                self.send_response(200)
                self.send_header("Content-Type", "image/jpeg")
                self.send_header("Content-Length", len(data))
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(data)
            else:
                self.send_response(404); self.end_headers()

        elif path == "/archive":
            jobs = sorted(ARCHIVE.iterdir(), reverse=True) if ARCHIVE.exists() else []
            self.send_json([j.name for j in jobs if j.is_dir()])

        elif path == "/progress":
            self.send_response(200)
            self.send_header("Content-Type", "text/event-stream")
            self.send_header("Cache-Control", "no-cache")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            try:
                while True:
                    try:
                        item = _q.get(timeout=25)
                        self.wfile.write(f"data: {item}\n\n".encode())
                        self.wfile.flush()
                        if json.loads(item).get("type") == "done": break
                    except queue.Empty:
                        self.wfile.write(b"data: {\"msg\":\".\",\"type\":\"ping\"}\n\n")
                        self.wfile.flush()
            except (BrokenPipeError, ConnectionResetError): pass

        else:
            self.send_response(404); self.end_headers()

    def do_POST(self):
        path = urlparse(self.path).path

        if path == "/upload":
            ct = self.headers.get("Content-Type", "")
            length = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(length)

            # Parse multipart
            msg = email.parser.BytesParser().parsebytes(
                f"Content-Type: {ct}\r\n\r\n".encode() + body)
            saved = []
            # Clear previous uploads
            for f in UPLOADS.iterdir(): f.unlink(missing_ok=True)
            for part in msg.get_payload():
                if hasattr(part, 'get_filename') and part.get_filename():
                    fname = part.get_filename()
                    dest  = UPLOADS / fname
                    dest.write_bytes(part.get_payload(decode=True))
                    saved.append(fname)
            self.send_json({"saved": saved})

        elif path == "/confirm":
            _confirm_event.set()
            self.send_json({"status": "confirmed"})

        elif path == "/run":
            length = int(self.headers.get("Content-Length", 0))
            data   = json.loads(self.rfile.read(length))
            while not _q.empty(): _q.get_nowait()
            threading.Thread(target=run_workflow, args=(data,), daemon=True).start()
            self.send_json({"status": "started"})

        else:
            self.send_response(404); self.end_headers()

# ── Entry point ───────────────────────────────────────────────────────────────
import socketserver

class ThreadedHTTPServer(socketserver.ThreadingMixIn, HTTPServer):
    daemon_threads = True

if __name__ == "__main__":
    import webbrowser
    server = ThreadedHTTPServer(("127.0.0.1", 5050), Handler)
    print("up", flush=True)
    threading.Timer(1.0, lambda: webbrowser.open("http://localhost:5050")).start()
    server.serve_forever()
