"""
Door-knock leave-behind flyer — A4 portrait, navy + gold palette
Single PPTX slide sized 8.27" x 11.69"
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import lxml.etree as etree
from pptx.oxml.ns import qn

# ─── COLOURS ──────────────────────────────────────────────────────────────────
NAVY    = RGBColor(0x1B, 0x2A, 0x4A)
GOLD    = RGBColor(0xC9, 0xA8, 0x4C)
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT   = RGBColor(0xF4, 0xF6, 0xFA)
LTGOLD  = RGBColor(0xFD, 0xF6, 0xE3)
MIDGRAY = RGBColor(0x64, 0x74, 0x8B)
DARK2   = RGBColor(0x16, 0x23, 0x38)
LTBLUE  = RGBColor(0xB0, 0xBE, 0xD4)
RULE    = RGBColor(0xD9, 0xE2, 0xEF)

W = 8.27   # A4 width inches
H = 11.69  # A4 height inches

def add_rect(slide, x, y, w, h, fill, line=None, lw=Pt(0.75)):
    s = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line: s.line.color.rgb = line; s.line.width = lw
    else: s.line.fill.background()
    return s

def add_oval(slide, x, y, w, h, fill, line=None):
    s = slide.shapes.add_shape(9, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line: s.line.color.rgb = line; s.line.width = Pt(1.5)
    else: s.line.fill.background()
    return s

def tb(slide, x, y, w, h, text, size=11, bold=False, italic=False,
       color=None, align=PP_ALIGN.LEFT, wrap=True):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    box.word_wrap = wrap
    tf = box.text_frame; tf.word_wrap = wrap
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run(); r.text = text
    r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic
    r.font.name = "Calibri"
    if color: r.font.color.rgb = color
    return box

def set_bg(slide, color):
    bg = slide.background; fill = bg.fill
    fill.solid(); fill.fore_color.rgb = color

def build():
    prs = Presentation()
    prs.slide_width  = Inches(W)
    prs.slide_height = Inches(H)
    blank = prs.slide_layouts[6]
    s = prs.slides.add_slide(blank)
    set_bg(s, WHITE)

    # ── HEADER BAND ────────────────────────────────────────────────────────────
    add_rect(s, 0, 0, W, 2.55, NAVY)

    # Gold left accent bar
    add_rect(s, 0, 0, 0.28, 2.55, GOLD)

    # Headline
    tb(s, 0.45, 0.25, 6.8, 0.72,
       "Is now the right time\nto sell your landed home?",
       size=26, bold=True, color=WHITE)

    # Sub
    tb(s, 0.45, 1.05, 6.5, 0.4,
       "Singapore's landed market is performing strongly — and serious buyers are actively looking in your estate.",
       size=11, italic=True, color=GOLD)

    # Agent placeholder box (top right)
    add_rect(s, 6.45, 0.22, 1.62, 2.05, DARK2, RGBColor(0x2E,0x41,0x66), Pt(0.75))
    tb(s, 6.5, 0.3,  1.52, 0.35, "ERA SINGAPORE", size=9, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
    add_rect(s, 6.62, 0.65, 1.2, 0.04, GOLD)   # thin rule
    tb(s, 6.5, 0.74, 1.52, 0.32, "[Agent Name]",     size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    tb(s, 6.5, 1.05, 1.52, 0.26, "[CEA Reg. No.]",   size=8.5, color=LTBLUE, align=PP_ALIGN.CENTER)
    tb(s, 6.5, 1.30, 1.52, 0.26, "[Mobile No.]",     size=9.5, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
    tb(s, 6.5, 1.56, 1.52, 0.26, "[Email]",          size=8.5, color=LTBLUE, align=PP_ALIGN.CENTER)
    tb(s, 6.5, 1.82, 1.52, 0.3,  "[WhatsApp QR]",    size=8, italic=True, color=MIDGRAY, align=PP_ALIGN.CENTER)

    # ── 4 STAT BOXES ───────────────────────────────────────────────────────────
    stats = [
        ("+6.7%",   "landed price growth\nyear-on-year (2025)"),
        ("+73%",    "URA landed price index\nQ1 2019 → Q1 2026"),
        ("~75,000", "landed homes in SG\n(only 5–6% of all housing)"),
        ("+46%",    "ultra-wealthy buyer count\nforecast by 2031"),
    ]
    bw = (W - 1.0) / 4
    for i, (val, lbl) in enumerate(stats):
        x = 0.5 + i * bw
        add_rect(s, x, 2.68, bw - 0.12, 1.12, LTGOLD, GOLD, Pt(1))
        tb(s, x+0.06, 2.75, bw-0.24, 0.48, val,
           size=22, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
        tb(s, x+0.06, 3.22, bw-0.24, 0.5, lbl,
           size=8.5, color=MIDGRAY, align=PP_ALIGN.CENTER)

    # ── WHY NOW ────────────────────────────────────────────────────────────────
    tb(s, 0.5, 3.98, 4.8, 0.28, "WHY SELLERS ARE MOVING NOW",
       size=8.5, bold=True, color=GOLD)
    add_rect(s, 0.5, 4.28, 4.8, 2.62, LIGHT, RULE, Pt(0.5))

    reasons = [
        ("Prices up 6.7% YoY",         "Strong market recovery after seasonal Q1 dip — Q2/Q3 is peak demand."),
        ("Declining interest rates",    "SORA falling → buyer affordability improving → larger qualified buyer pool."),
        ("Wealthiest buyer base ever",  "Singapore UHNWI count up 55% in 5 years. Forecast: +46% more by 2031."),
        ("Supply stays scarce",         "Landed stock grew only 12% in 25 years vs 228% for non-landed condos."),
        ("60% ABSD for foreigners",     "Landed stays exclusive to locals/PRs — buyers compete in a motivated pool."),
    ]
    for i, (title, desc) in enumerate(reasons):
        y = 4.38 + i * 0.48
        add_oval(s, 0.62, y+0.06, 0.26, 0.26, NAVY)
        tb(s, 0.62, y+0.06, 0.26, 0.26, str(i+1), size=8, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
        tb(s, 0.98, y+0.04, 1.4, 0.22, title, size=9.5, bold=True, color=NAVY)
        tb(s, 0.98, y+0.25, 3.2, 0.2,  desc,  size=8.5, color=MIDGRAY)

    # ── MEDIAN PRICES ──────────────────────────────────────────────────────────
    tb(s, 5.5, 3.98, 2.58, 0.28, "MEDIAN PRICES (1Q 2026)",
       size=8.5, bold=True, color=GOLD)
    add_rect(s, 5.5, 4.28, 2.58, 2.62, LIGHT, RULE, Pt(0.5))

    prices = [
        ("OCR Terrace",    "$4.44M"),
        ("OCR Semi-D",     "$5.75M"),
        ("RCR Detached",   "$10.9M"),
        ("CCR Detached",   "$16.8M"),
    ]
    for i, (label, price) in enumerate(prices):
        y = 4.38 + i * 0.6
        add_rect(s, 5.6, y, 2.36, 0.5, WHITE, RULE, Pt(0.5))
        tb(s, 5.68, y+0.05, 1.3, 0.24, label, size=9.5, color=NAVY)
        tb(s, 6.85, y+0.05, 1.0, 0.38, price, size=14, bold=True, color=GOLD, align=PP_ALIGN.RIGHT)

    tb(s, 5.5, 6.72, 2.58, 0.18, "Source: ERA Research 1Q 2026 Landed Report",
       size=7.5, italic=True, color=MIDGRAY)

    # ── WHAT YOUR HOME COULD BE WORTH ──────────────────────────────────────────
    add_rect(s, 0, 7.05, W, 1.32, NAVY)
    add_rect(s, 0, 7.05, 0.28, 1.32, GOLD)
    tb(s, 0.45, 7.1, 5.5, 0.3, "WHAT YOUR HOME COULD BE WORTH",
       size=9, bold=True, color=GOLD)
    tb(s, 0.45, 7.4, 7.5, 0.68,
       "Based on recent comparable transactions in your estate, similar homes have transacted at "
       "$[X]M–$[Y]M. After settling your CPF refund and any outstanding mortgage, most sellers in "
       "this estate are walking away with $[Z]M or more in cash proceeds.",
       size=10.5, color=WHITE)

    # ── 3-STEP PROCESS ─────────────────────────────────────────────────────────
    tb(s, 0.5, 8.55, 7.5, 0.28, "HOW IT WORKS — THREE SIMPLE STEPS",
       size=8.5, bold=True, color=GOLD)

    steps = [
        ("1", "Free property assessment",   "I visit, assess condition, pull the latest comps and give you an honest, no-obligation valuation."),
        ("2", "Marketing strategy meeting", "We agree on the right price and approach. I'll show you exactly who the buyers are and how I'll reach them."),
        ("3", "Sign listing & launch",      "Photography within 5 days. Listing live within 7 days. Active buyer outreach starts immediately."),
    ]
    sw = (W - 1.0) / 3
    for i, (num, title, desc) in enumerate(steps):
        x = 0.5 + i * sw
        add_rect(s, x, 8.88, sw - 0.15, 1.55, LTGOLD, GOLD, Pt(0.75))
        add_oval(s, x + (sw-0.15)/2 - 0.2, 8.95, 0.4, 0.4, NAVY)
        tb(s, x + (sw-0.15)/2 - 0.2, 8.95, 0.4, 0.4, num,
           size=12, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
        tb(s, x+0.1, 9.42, sw-0.35, 0.3, title,
           size=9.5, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
        tb(s, x+0.1, 9.72, sw-0.35, 0.6, desc,
           size=8.5, color=MIDGRAY, align=PP_ALIGN.CENTER)

    # ── FOOTER BAND ────────────────────────────────────────────────────────────
    add_rect(s, 0, 10.58, W, 1.11, DARK2)
    add_rect(s, 0, 10.58, 0.28, 1.11, GOLD)

    tb(s, 0.45, 10.65, 3.5, 0.3,  "[Agent Name]",
       size=13, bold=True, color=WHITE)
    tb(s, 0.45, 10.96, 3.5, 0.26, "ERA Singapore  |  CEA Reg. No. [XXXXX]",
       size=9, color=LTBLUE)

    tb(s, 4.1, 10.65, 2.3, 0.26, "[Mobile / WhatsApp]", size=10.5, bold=True, color=GOLD)
    tb(s, 4.1, 10.94, 2.3, 0.26, "[Email Address]",     size=10,   color=LTBLUE)

    # Disclaimer
    tb(s, 0.45, 11.38, W-0.9, 0.22,
       "Data sourced from ERA Research, URA REALIS and Knight Frank (2026). This flyer is for informational purposes only and does not constitute a formal valuation.",
       size=6.5, italic=True, color=MIDGRAY)

    out = "/Users/riopang/Desktop/Claude Code/Doorknock Leave-Behind Flyer.pptx"
    prs.save(out)
    print(f"Saved → {out}")

build()
