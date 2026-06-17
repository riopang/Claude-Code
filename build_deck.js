const pptxgen = require("pptxgenjs");
const path = require("path");

// ─── COLORS ───────────────────────────────────────────────────────────────────
const NAVY    = "1B2A4A";
const GOLD    = "C9A84C";
const WHITE   = "FFFFFF";
const LIGHT   = "F4F6FA";
const MIDGRAY = "64748B";
const LTGOLD  = "FDF6E3";
const PLHOLD  = "E8EDF5";   // placeholder box bg
const RULE    = "D9E2EF";   // subtle divider color

// ─── HELPERS ──────────────────────────────────────────────────────────────────
const mkShadow = () => ({ type: "outer", color: "000000", blur: 8, offset: 2, angle: 45, opacity: 0.10 });

function addSlideHeader(slide, label, title, darkMode = false) {
  const labelColor  = darkMode ? GOLD   : GOLD;
  const titleColor  = darkMode ? WHITE  : NAVY;
  const bgColor     = darkMode ? NAVY   : WHITE;

  // Label (small caps category)
  slide.addText(label.toUpperCase(), {
    x: 0.5, y: 0.28, w: 9, h: 0.22,
    fontFace: "Calibri", fontSize: 9, bold: true,
    color: labelColor, charSpacing: 3, align: "left", margin: 0,
  });

  // Title
  slide.addText(title, {
    x: 0.5, y: 0.52, w: 9, h: 0.72,
    fontFace: "Calibri", fontSize: 28, bold: true,
    color: titleColor, align: "left", margin: 0,
  });
}

function statBox(slide, val, label, x, y, w = 2.7, h = 1.15) {
  slide.addShape("rect", {
    x, y, w, h,
    fill: { color: LTGOLD },
    line: { color: GOLD, width: 1.5 },
    shadow: mkShadow(),
  });
  slide.addText(val, {
    x: x + 0.12, y: y + 0.06, w: w - 0.24, h: 0.6,
    fontFace: "Calibri", fontSize: 28, bold: true, color: NAVY, align: "center", margin: 0,
  });
  slide.addText(label, {
    x: x + 0.12, y: y + 0.62, w: w - 0.24, h: 0.45,
    fontFace: "Calibri", fontSize: 10, color: MIDGRAY, align: "center", margin: 0,
  });
}

function bulletList(slide, items, x, y, w, h, fontSize = 13) {
  const arr = items.map((t, i) => ({
    text: t,
    options: { bullet: true, breakLine: i < items.length - 1, fontSize, color: NAVY, fontFace: "Calibri", paraSpaceAfter: 4 },
  }));
  slide.addText(arr, { x, y, w, h, valign: "top", margin: 4 });
}

// ─── PRESENTATION ─────────────────────────────────────────────────────────────
async function buildDeck() {
  const pres = new pptxgen();
  pres.layout  = "LAYOUT_16x9";
  pres.title   = "Landed Property Seller Pitch Deck — Singapore 2026";
  pres.author  = "ERA Singapore";

  // ══════════════════════════════════════════════════════════════════════
  // SLIDE 1 — COVER
  // ══════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: NAVY };

    // Gold accent block (left panel)
    s.addShape("rect", { x: 0, y: 0, w: 0.35, h: 5.625, fill: { color: GOLD }, line: { color: GOLD } });

    // Headline
    s.addText("Thinking of Selling\nYour Landed Home?", {
      x: 0.65, y: 0.9, w: 7.5, h: 2.0,
      fontFace: "Calibri", fontSize: 40, bold: true, color: WHITE,
      align: "left", margin: 0,
    });

    // Subheadline
    s.addText("Your trusted landed property specialist in Singapore.", {
      x: 0.65, y: 2.95, w: 7.2, h: 0.45,
      fontFace: "Calibri", fontSize: 16, italic: true, color: GOLD,
      align: "left", margin: 0,
    });

    // Divider
    s.addShape("rect", { x: 0.65, y: 3.48, w: 4.5, h: 0.03, fill: { color: GOLD }, line: { color: GOLD } });

    // Tagline
    s.addText("Data-driven.  Results-focused.  Landed specialist.", {
      x: 0.65, y: 3.6, w: 7.0, h: 0.35,
      fontFace: "Calibri", fontSize: 13, charSpacing: 1, color: "B0BED4",
      align: "left", margin: 0,
    });

    // Agent info block
    s.addShape("rect", { x: 0.65, y: 4.12, w: 5.2, h: 1.2,
      fill: { color: "162338" }, line: { color: "2E4166" }, shadow: mkShadow() });
    s.addText([
      { text: "[INSERT: Agent Name]",     options: { breakLine: true, bold: true, fontSize: 14, color: WHITE } },
      { text: "[INSERT: CEA Reg. No.]",   options: { breakLine: true, fontSize: 11, color: "B0BED4" } },
      { text: "[INSERT: Mobile Number]",  options: { breakLine: true, fontSize: 11, color: GOLD } },
      { text: "[INSERT: Email Address]",  options: { fontSize: 11, color: GOLD } },
    ], { x: 0.88, y: 4.22, w: 4.8, h: 1.0, fontFace: "Calibri", margin: 0 });

    // ERA badge placeholder
    s.addShape("rect", { x: 8.0, y: 4.1, w: 1.6, h: 1.2,
      fill: { color: "162338" }, line: { color: "2E4166" } });
    s.addText("ERA\nSINGAPORE", {
      x: 8.0, y: 4.1, w: 1.6, h: 1.2,
      fontFace: "Calibri", fontSize: 14, bold: true, color: GOLD,
      align: "center", valign: "middle", margin: 0,
    });

    s.addNotes("Cover slide. Fill in agent name, CEA registration, mobile and email before presenting.");
  }

  // ══════════════════════════════════════════════════════════════════════
  // SLIDE 2 — SCARCITY: Singapore Landed Market
  // ══════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    addSlideHeader(s, "Market Context", "Singapore Landed — A Structurally Scarce Asset");

    // 3 stat boxes
    statBox(s, "~75,000",  "landed homes in Singapore\n(~5–6% of total housing stock)", 0.5,  1.45, 2.85, 1.35);
    statBox(s, "+12%",     "supply growth over 25 years\n(vs. +228% for non-landed condos)", 3.57, 1.45, 2.85, 1.35);
    statBox(s, "+73%",     "URA landed price index\nQ1 2019 → Q1 2026", 6.65, 1.45, 2.85, 1.35);

    // Body
    bulletList(s, [
      "Supply grew from 67,229 units (2000) to 75,338 units (2025) — just 12% over 25 years.",
      "Over the same period, private non-landed condos grew 228%: 114,532 → 375,612 units.",
      "The government releases very little new land for landed housing — scarcity is permanent.",
      "Buyers recognise this: landed homes are a defensive asset that cannot scale to meet demand.",
      "Result: prices have compounded steadily even through multiple rounds of cooling measures.",
    ], 0.5, 2.93, 9.0, 2.3);

    s.addNotes("Key message: landed supply is structurally capped. Non-landed condos flooded the market; landed did not. This is why landed prices are resilient.");
  }

  // ══════════════════════════════════════════════════════════════════════
  // SLIDE 3 — 2026 MARKET SNAPSHOT
  // ══════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    addSlideHeader(s, "Market Data", "2026 Landed Market Snapshot");

    // Stat callouts (top row)
    statBox(s, "+6.7%",  "YoY price growth\n(landed, 2025 full year)", 0.5,  1.45, 2.2, 1.1);
    statBox(s, "+3.4%",  "QoQ surge in Q4 2025\n(strongest quarter)", 2.82, 1.45, 2.2, 1.1);
    statBox(s, "+15%",   "transaction value growth\n(2025 vs 2024 YoY)", 5.14, 1.45, 2.2, 1.1);
    statBox(s, "+5–7%",  "ERA 2026 price forecast\n(1,750–1,950 transactions)", 7.30, 1.45, 2.2, 1.1);

    // Median price table
    const header = { fill: { color: NAVY }, color: WHITE, bold: true, fontSize: 12 };
    const row1   = { fill: { color: LIGHT }, color: NAVY, fontSize: 12 };
    const row2   = { fill: { color: WHITE }, color: NAVY, fontSize: 12 };

    s.addTable([
      [
        { text: "Property Type",         options: { ...header } },
        { text: "Region",                options: { ...header } },
        { text: "Median Price (2026)",    options: { ...header } },
      ],
      [
        { text: "Terrace (Intermediate)", options: { ...row1 } },
        { text: "OCR",                    options: { ...row1 } },
        { text: "$4.44M",                 options: { ...row1, bold: true } },
      ],
      [
        { text: "Semi-Detached",          options: { ...row2 } },
        { text: "OCR",                    options: { ...row2 } },
        { text: "$5.75M",                 options: { ...row2, bold: true } },
      ],
      [
        { text: "Detached",               options: { ...row1 } },
        { text: "RCR",                    options: { ...row1 } },
        { text: "$10.9M",                 options: { ...row1, bold: true } },
      ],
      [
        { text: "Detached (Bungalow)",    options: { ...row2 } },
        { text: "CCR",                    options: { ...row2 } },
        { text: "$16.825M",               options: { ...row2, bold: true } },
      ],
    ], {
      x: 0.5, y: 2.68, w: 9.0, h: 2.6,
      colW: [3.5, 2.0, 3.5],
      border: { pt: 1, color: RULE },
    });

    s.addNotes("Source: ERA Research 1Q 2026 Landed Report. Q1 2026 showed a -0.4% QoQ seasonal dip — use this to frame now as a buying window for buyers, and an optimal listing window for sellers before Q2/Q3 demand returns.");
  }

  // ══════════════════════════════════════════════════════════════════════
  // SLIDE 4 — WHY SELL NOW
  // ══════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    addSlideHeader(s, "Market Timing", "Why Now Is the Right Time to Sell");

    // Gold highlight box
    s.addShape("rect", { x: 0.5, y: 1.42, w: 9.0, h: 0.72,
      fill: { color: LTGOLD }, line: { color: GOLD, width: 1.5 } });
    s.addText("Q1 seasonal dip = optimal listing window before Q2/Q3 demand surge returns", {
      x: 0.65, y: 1.50, w: 8.7, h: 0.54,
      fontFace: "Calibri", fontSize: 13, bold: true, italic: true, color: NAVY,
      align: "left", margin: 0,
    });

    // Two columns of bullets
    const leftPts = [
      "SORA is declining → buyer affordability improving → larger qualified buyer pool",
      "Singapore ultra-wealthy (US$30M+) grew 55% in 5 years: 4,642 → 7,171 in 2026",
      "Knight Frank forecasts UHNWI count to grow a further 46% to 10,497 by 2031",
    ];
    const rightPts = [
      "25,000–30,000 new citizenships granted annually → steady UHNWI demand pipeline",
      "60% ABSD for foreigners keeps landed as a local/PR exclusive — buyers compete in a restricted, motivated pool",
      "2025 landed transaction value was up 15% YoY — momentum is building",
    ];

    bulletList(s, leftPts,  0.5,  2.28, 4.45, 2.95);
    bulletList(s, rightPts, 5.05, 2.28, 4.45, 2.95);

    s.addNotes("Key message for data-savvy owners: the confluence of SORA easing, UHNWI wealth growth, and citizenship-driven demand means the qualified buyer pool is growing — but the supply of landed homes is not.");
  }

  // ══════════════════════════════════════════════════════════════════════
  // SLIDE 5 — WHO IS BUYING
  // ══════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    addSlideHeader(s, "Buyer Profiling", "Who Is Buying Landed in 2026?");

    // Big stat left
    s.addShape("rect", { x: 0.5, y: 1.42, w: 3.0, h: 3.8,
      fill: { color: NAVY }, shadow: mkShadow() });
    s.addText("57.9%", {
      x: 0.5, y: 1.9, w: 3.0, h: 1.0,
      fontFace: "Calibri", fontSize: 44, bold: true, color: GOLD,
      align: "center", margin: 0,
    });
    s.addText("of all landed transactions\nare priced above $5M", {
      x: 0.55, y: 2.95, w: 2.9, h: 0.7,
      fontFace: "Calibri", fontSize: 12, color: WHITE,
      align: "center", margin: 0,
    });
    s.addText("Only 2.1%\nare HDB upgraders", {
      x: 0.55, y: 3.72, w: 2.9, h: 0.65,
      fontFace: "Calibri", fontSize: 11, italic: true, color: "B0BED4",
      align: "center", margin: 0,
    });

    // Right: buyer segments
    s.addText("Primary Buyer Segments", {
      x: 3.8, y: 1.42, w: 5.7, h: 0.35,
      fontFace: "Calibri", fontSize: 14, bold: true, color: NAVY, margin: 0,
    });

    const segs = [
      ["Multi-generational family wealth",  "Upgrading within the landed tier; wealth preservation mindset."],
      ["UHNWI / New Citizens",              "High-net-worth PRs and new citizens; eligible pool expanding +46% by 2031."],
      ["Condo upgraders",                   "2026 new launches skew to 1–2 bed units — landlocked condo owners look to landed."],
      ["Developer / redevelopment buyers",  "Buying for knock-down-rebuild; broad buyer pool expands your exit options."],
    ];

    segs.forEach(([title, desc], i) => {
      const y = 1.88 + i * 0.84;
      s.addShape("rect", { x: 3.8, y, w: 5.7, h: 0.72,
        fill: { color: i % 2 === 0 ? LIGHT : WHITE },
        line: { color: RULE, width: 0.5 },
      });
      s.addText(title, {
        x: 4.0, y: y + 0.06, w: 5.3, h: 0.26,
        fontFace: "Calibri", fontSize: 12, bold: true, color: NAVY, margin: 0,
      });
      s.addText(desc, {
        x: 4.0, y: y + 0.34, w: 5.3, h: 0.32,
        fontFace: "Calibri", fontSize: 10, color: MIDGRAY, margin: 0,
      });
    });

    s.addNotes("Key insight: the buyer pool is smaller but wealthier and more determined. ABSD restrictions for foreigners actually protect local landed sellers by keeping competition high among an eligible pool that cannot grow via immigration.");
  }

  // ══════════════════════════════════════════════════════════════════════
  // SLIDE 6 — WHAT DRIVES YOUR HOME'S VALUE
  // ══════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    addSlideHeader(s, "Valuation", "What Drives Your Home's Value");

    const drivers = [
      ["01", "Condition of Home",         "Age, renovation quality, quality of fittings and finishes"],
      ["02", "Facing & Views",            "Landmark views, park/greenery frontage, orientation (N-S preferred)"],
      ["03", "Parking Availability",      "Number of car porch spaces, covered vs uncovered"],
      ["04", "MRT & Connectivity",        "Walking distance to MRT, bus interchange, expressway access"],
      ["05", "School Catchment",          "Primary school within 1km (P1 registration priority)"],
      ["06", "Comparable PSF Transactions","Recent transacted prices on your street and estate"],
      ["07", "Neighbourhood Upgrading",   "HIP, road widening, new amenities — signals area trajectory"],
    ];

    drivers.forEach(([num, title, desc], i) => {
      const col = i < 4 ? 0 : 1;
      const row = i < 4 ? i : i - 4;
      const x = col === 0 ? 0.5 : 5.1;
      const y = 1.42 + row * 1.02;

      s.addShape("rect", { x, y, w: 4.45, h: 0.88,
        fill: { color: col === 0 ? LIGHT : LTGOLD },
        line: { color: RULE, width: 0.5 },
        shadow: mkShadow(),
      });
      s.addText(num, {
        x: x + 0.12, y: y + 0.08, w: 0.45, h: 0.35,
        fontFace: "Calibri", fontSize: 16, bold: true, color: GOLD, margin: 0,
      });
      s.addText(title, {
        x: x + 0.62, y: y + 0.08, w: 3.7, h: 0.3,
        fontFace: "Calibri", fontSize: 12, bold: true, color: NAVY, margin: 0,
      });
      s.addText(desc, {
        x: x + 0.62, y: y + 0.42, w: 3.7, h: 0.38,
        fontFace: "Calibri", fontSize: 10, color: MIDGRAY, margin: 0,
      });
    });

    // Bonus note (slide 7 slot in right col)
    const bx = 5.1, by = 1.42 + 3 * 1.02;
    s.addShape("rect", { x: bx, y: by, w: 4.45, h: 0.88,
      fill: { color: NAVY }, line: { color: NAVY } });
    s.addText("+ BONUS", {
      x: bx + 0.12, y: by + 0.08, w: 0.8, h: 0.28,
      fontFace: "Calibri", fontSize: 9, bold: true, color: GOLD, charSpacing: 2, margin: 0,
    });
    s.addText("Road Line Category & Redevelopment Potential", {
      x: bx + 0.12, y: by + 0.36, w: 4.2, h: 0.42,
      fontFace: "Calibri", fontSize: 11, color: WHITE, margin: 0,
    });

    s.addNotes("Walk the owner through each driver and ask them to score their own property. This builds trust and positions you as the advisor.");
  }

  // ══════════════════════════════════════════════════════════════════════
  // SLIDE 7 — PSF BENCHMARKS BY ZONE
  // ══════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    addSlideHeader(s, "Pricing Intelligence", "Current PSF Benchmarks by Zone (2026)");

    const zones = [
      { zone: "OCR",  type: "Terrace (Intermediate)", psf: "$1,100 – $1,400",  note: "Outside Central Region"  },
      { zone: "OCR",  type: "Semi-Detached",           psf: "$1,200 – $1,600",  note: "Outside Central Region"  },
      { zone: "RCR",  type: "Terrace / Semi-D",        psf: "$1,500 – $2,000",  note: "Rest of Central Region"  },
      { zone: "CCR",  type: "Detached / GCB",          psf: "$2,000 – $3,500+", note: "Core Central Region"     },
    ];

    zones.forEach(({ zone, type, psf, note }, i) => {
      const x = 0.5 + i * 2.35;
      s.addShape("rect", { x, y: 1.42, w: 2.15, h: 2.55,
        fill: { color: i === 3 ? NAVY : LIGHT },
        line: { color: i === 3 ? GOLD : RULE, width: 1 },
        shadow: mkShadow(),
      });
      s.addText(zone, {
        x: x + 0.08, y: 1.52, w: 1.98, h: 0.36,
        fontFace: "Calibri", fontSize: 22, bold: true,
        color: i === 3 ? GOLD : NAVY, align: "center", margin: 0,
      });
      s.addText(type, {
        x: x + 0.08, y: 1.92, w: 1.98, h: 0.45,
        fontFace: "Calibri", fontSize: 11, color: i === 3 ? "B0BED4" : MIDGRAY,
        align: "center", margin: 0,
      });
      s.addShape("rect", { x: x + 0.15, y: 2.44, w: 1.85, h: 0.03,
        fill: { color: i === 3 ? GOLD : GOLD }, line: { color: i === 3 ? GOLD : GOLD } });
      s.addText(psf, {
        x: x + 0.08, y: 2.52, w: 1.98, h: 0.52,
        fontFace: "Calibri", fontSize: 14, bold: true,
        color: i === 3 ? WHITE : NAVY, align: "center", margin: 0,
      });
      s.addText("psf (land)", {
        x: x + 0.08, y: 3.06, w: 1.98, h: 0.28,
        fontFace: "Calibri", fontSize: 9, color: i === 3 ? "B0BED4" : MIDGRAY,
        align: "center", margin: 0,
      });
    });

    // Premiums
    s.addShape("rect", { x: 0.5, y: 4.1, w: 9.0, h: 0.8,
      fill: { color: LTGOLD }, line: { color: GOLD, width: 1 } });
    s.addText([
      { text: "Freehold premium: ", options: { bold: true } },
      { text: "+10–20% over 99-year leasehold     ", options: {} },
      { text: "Corner terrace premium: ", options: { bold: true } },
      { text: "+15–25% over intermediate terrace     ", options: {} },
      { text: "Redevelopment potential: ", options: { bold: true } },
      { text: "adds a further 5–15% to buyer appetite", options: {} },
    ], {
      x: 0.7, y: 4.15, w: 8.6, h: 0.65,
      fontFace: "Calibri", fontSize: 11, color: NAVY, align: "left", margin: 0,
    });

    // Placeholder
    s.addText("[INSERT: Your property's zone, PSF range, and indicative value — to be filled in during consultation]", {
      x: 0.5, y: 4.96, w: 9.0, h: 0.35,
      fontFace: "Calibri", fontSize: 9.5, italic: true, color: MIDGRAY,
      align: "center", margin: 0,
    });

    s.addNotes("PSF figures are land PSF (not built-up/GFA). Use URA REALIS or PropertyGuru to pull the last 6 months of transactions on the same street before presenting.");
  }

  // ══════════════════════════════════════════════════════════════════════
  // SLIDE 8 — NET PROCEEDS
  // ══════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    addSlideHeader(s, "Financial Planning", "Your Estimated Net Proceeds");

    const rows = [
      ["Estimated Sale Price",                              "$[AMOUNT]",  false, false],
      ["Less: Agent Commission (2% + 9% GST)",              "−$[AMOUNT]", false, true ],
      ["Less: CPF Refund + Accrued Interest (if applicable)","−$[AMOUNT]", false, true ],
      ["Less: Outstanding Mortgage (if any)",               "−$[AMOUNT]", false, true ],
      ["Less: Seller's Stamp Duty (SSD — if owned <3 yrs)", "−$[AMOUNT]", false, true ],
      ["Estimated Cash Proceeds",                           "$[AMOUNT]",  true,  false],
    ];

    const tableData = rows.map(([label, val, isTotal, isDeduct]) => [
      { text: label, options: {
        fill:  { color: isTotal ? NAVY : isDeduct ? WHITE : LTGOLD },
        color: isTotal ? WHITE : NAVY, bold: isTotal, fontSize: 12,
      }},
      { text: val, options: {
        fill:  { color: isTotal ? NAVY : isDeduct ? WHITE : LTGOLD },
        color: isTotal ? GOLD  : isDeduct ? "C0392B" : NAVY,
        bold: isTotal, fontSize: 12, align: "right",
      }},
    ]);

    s.addTable(tableData, {
      x: 1.0, y: 1.45, w: 8.0, h: 3.5,
      colW: [5.8, 2.2],
      border: { pt: 1, color: RULE },
    });

    s.addShape("rect", { x: 1.0, y: 5.05, w: 8.0, h: 0.35,
      fill: { color: PLHOLD }, line: { color: RULE } });
    s.addText("Note: BSD (Buyer's Stamp Duty) is paid by the buyer, not the seller.  SSD applies only if property held <3 years (Tier 1: 12%, Tier 2: 8%, Tier 3: 4%).", {
      x: 1.1, y: 5.09, w: 7.8, h: 0.27,
      fontFace: "Calibri", fontSize: 9, color: MIDGRAY, margin: 0,
    });

    s.addNotes("Fill in the estimated sale price before the meeting. Walk the owner through each deduction line. The CPF refund is often the biggest surprise — they must refund CPF OA principal + accrued interest at prevailing CPF OA rate.");
  }

  // ══════════════════════════════════════════════════════════════════════
  // SLIDE 9 — SELLING JOURNEY
  // ══════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    addSlideHeader(s, "Process", "The Selling Journey — What to Expect");

    const stages = [
      { num: "01", title: "Pre-Launch",          weeks: "2 – 4 weeks",       bullets: ["Staging & property prep", "Professional photography", "Drone / aerial video", "Valuation & pricing strategy"] },
      { num: "02", title: "Active Marketing",    weeks: "4 – 8 weeks",       bullets: ["PropertyGuru listing", "ERA network activation", "Buyer outreach & door-knock", "Viewings & feedback"] },
      { num: "03", title: "OTP & Negotiation",   weeks: "1 – 2 weeks",       bullets: ["Offer received", "Option to Purchase issued", "14-day exercise period", "Terms negotiated"] },
      { num: "04", title: "Completion",          weeks: "8 – 12 weeks",      bullets: ["Conveyancing by solicitors", "CPF refund processed", "Loan redemption", "Key handover"] },
    ];

    // Connector line
    s.addShape("rect", { x: 0.5, y: 2.22, w: 9.0, h: 0.04,
      fill: { color: GOLD }, line: { color: GOLD } });

    stages.forEach(({ num, title, weeks, bullets }, i) => {
      const x = 0.5 + i * 2.35;

      // Dot on timeline
      s.addShape("ellipse", { x: x + 0.75, y: 2.02, w: 0.44, h: 0.44,
        fill: { color: NAVY }, line: { color: GOLD, width: 2 } });
      s.addText(num, {
        x: x + 0.75, y: 2.02, w: 0.44, h: 0.44,
        fontFace: "Calibri", fontSize: 9, bold: true, color: GOLD,
        align: "center", valign: "middle", margin: 0,
      });

      // Stage card
      s.addShape("rect", { x, y: 2.6, w: 2.15, h: 2.7,
        fill: { color: i % 2 === 0 ? LIGHT : LTGOLD },
        line: { color: RULE, width: 0.5 },
        shadow: mkShadow(),
      });
      s.addText(title, {
        x: x + 0.1, y: 2.68, w: 1.95, h: 0.35,
        fontFace: "Calibri", fontSize: 13, bold: true, color: NAVY, margin: 0,
      });
      s.addText(weeks, {
        x: x + 0.1, y: 3.04, w: 1.95, h: 0.26,
        fontFace: "Calibri", fontSize: 9.5, italic: true, color: GOLD, margin: 0,
      });
      bulletList(s, bullets, x + 0.1, 3.36, 1.95, 1.8, 10);
    });

    // Total timeline note
    s.addText("Typical total timeline: 3 – 6 months from signed listing to key handover.", {
      x: 0.5, y: 5.27, w: 9.0, h: 0.25,
      fontFace: "Calibri", fontSize: 10, italic: true, color: MIDGRAY, align: "center", margin: 0,
    });

    s.addNotes("Walk through this timeline to set realistic expectations. Most sellers underestimate how long completion takes after OTP exercise — it's 8–12 weeks for solicitors to complete.");
  }

  // ══════════════════════════════════════════════════════════════════════
  // SLIDE 10 — MARKETING YOUR HOME
  // ══════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    addSlideHeader(s, "Marketing Strategy", "How I Will Market Your Home");

    // Left column: photo/media
    s.addShape("rect", { x: 0.5, y: 1.42, w: 4.45, h: 3.9,
      fill: { color: NAVY }, shadow: mkShadow() });
    s.addText("PHOTOGRAPHY\n& MEDIA", {
      x: 0.65, y: 1.55, w: 4.1, h: 0.55,
      fontFace: "Calibri", fontSize: 14, bold: true, color: GOLD,
      charSpacing: 1, align: "left", margin: 0,
    });

    const mediaItems = [
      "360° interior photography with gimbal rig",
      "Architectural highlights: staircase, marble, pool, lifestyle spaces",
      "Drone/aerial video showcasing land area and surroundings",
      "Professional video walkthrough for online listings",
      "Staged photography — room-by-room with natural lighting",
    ];
    bulletList(s, mediaItems, 0.65, 2.18, 4.1, 2.9, 11);

    // Right column: distribution
    s.addShape("rect", { x: 5.05, y: 1.42, w: 4.45, h: 3.9,
      fill: { color: LIGHT }, shadow: mkShadow() });
    s.addText("DISTRIBUTION\n& OUTREACH", {
      x: 5.2, y: 1.55, w: 4.1, h: 0.55,
      fontFace: "Calibri", fontSize: 14, bold: true, color: NAVY,
      charSpacing: 1, align: "left", margin: 0,
    });

    const distItems = [
      "PropertyGuru Premium + ERA PropTrack listing",
      "ERA Singapore network: 7,000+ active agents",
      "ERA Asia-Pacific network for UHNWI & overseas PR buyers",
      "Targeted door-knock outreach to direct buyers in your estate",
      "Instagram & Facebook social media campaign with hashtag strategy",
      "Weekly seller report: viewings count, enquiry quality, market feedback",
    ];
    bulletList(s, distItems, 5.2, 2.18, 4.1, 2.9, 11);

    s.addNotes("Emphasise the ERA international network — UHNWI buyers often come through referrals from ERA's regional offices. This is a differentiator over smaller agencies.");
  }

  // ══════════════════════════════════════════════════════════════════════
  // SLIDE 11 — TECHNICAL ADVISORY
  // ══════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    addSlideHeader(s, "Specialist Knowledge", "Technical Advisory — Why a Specialist Matters");

    s.addShape("rect", { x: 0.5, y: 1.42, w: 9.0, h: 0.62,
      fill: { color: LTGOLD }, line: { color: GOLD, width: 1.2 } });
    s.addText("Serious buyers for landed homes — especially redevelopment buyers — ask technical questions. An agent who can answer them on the spot wins the listing and closes faster.", {
      x: 0.65, y: 1.5, w: 8.7, h: 0.48,
      fontFace: "Calibri", fontSize: 11, italic: true, color: NAVY,
      align: "left", margin: 0,
    });

    const leftItems = [
      ["Road Line Category (CAT 1–5)", "Determines required front setback; CAT 1–2 roads may require land take-back, affecting buildable area."],
      ["SIP / DIP / WSP", "Sewerage Information Plan, Drainage Interpretation Plan, Water Service Plan — critical checks before redevelopment."],
      ["Tree Conservation Areas (TCA)", "South/Eastern zones; individual tree removal costs $150–$600/tree; affects plot use."],
    ];
    const rightItems = [
      ["Redevelopment Cost Guide", "Rebuild from scratch: $800,000+, 1.5–2.5 years. Reconstruction: 1–2 years. A&A: 6 months+."],
      ["Pool & Lift Costs for Buyers", "Concrete pool: $150K–$300K | Fibreglass: $30K–$60K. Home lift: $30K–$100K+ depending on type."],
      ["Envelope Control & Zoning", "2-storey vs 3-storey rebuild potential; gross plot ratio; URA planning parameters."],
    ];

    [leftItems, rightItems].forEach((items, col) => {
      items.forEach(([title, desc], row) => {
        const x = col === 0 ? 0.5 : 5.1;
        const y = 2.18 + row * 1.08;
        s.addShape("rect", { x, y, w: 4.45, h: 0.95,
          fill: { color: col === 0 ? LIGHT : LTGOLD },
          line: { color: RULE, width: 0.5 },
        });
        s.addText(title, {
          x: x + 0.15, y: y + 0.09, w: 4.15, h: 0.28,
          fontFace: "Calibri", fontSize: 11.5, bold: true, color: NAVY, margin: 0,
        });
        s.addText(desc, {
          x: x + 0.15, y: y + 0.42, w: 4.15, h: 0.46,
          fontFace: "Calibri", fontSize: 10, color: MIDGRAY, margin: 0,
        });
      });
    });

    s.addNotes("This slide demonstrates specialist credibility. Memorise these points — the moment you start talking about road line categories and SIP checks, data-savvy owners know you're the real deal.");
  }

  // ══════════════════════════════════════════════════════════════════════
  // SLIDE 12 — YOUR NEXT MOVE (S.W.A.P)
  // ══════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    addSlideHeader(s, "Wealth Planning", "What Happens After You Sell — Your Next Move");

    // SWAP cards
    const cards = [
      { letter: "S", word: "Savings",           desc: "Unlock home equity tied up in your property. Redeploy as liquid capital or down-payment." },
      { letter: "W", word: "Work Progression",  desc: "As income grows, so does your borrowing capacity — upgrade to the next tier of property." },
      { letter: "A", word: "Appreciation",      desc: "Strategic property swaps have historically generated $400K–$500K in net gains within 5–7 years." },
      { letter: "P", word: "Principal Paydown", desc: "Each mortgage payment builds equity. At sale, these gains compound into your next purchase." },
    ];

    cards.forEach(({ letter, word, desc }, i) => {
      const x = 0.5 + i * 2.35;
      s.addShape("rect", { x, y: 1.42, w: 2.15, h: 2.65,
        fill: { color: i % 2 === 0 ? NAVY : LTGOLD },
        line: { color: i % 2 === 0 ? GOLD  : GOLD, width: 1 },
        shadow: mkShadow(),
      });
      s.addText(letter, {
        x: x + 0.08, y: 1.52, w: 1.98, h: 0.88,
        fontFace: "Calibri", fontSize: 52, bold: true,
        color: GOLD, align: "center", margin: 0,
      });
      s.addText(word, {
        x: x + 0.08, y: 2.42, w: 1.98, h: 0.36,
        fontFace: "Calibri", fontSize: 13, bold: true,
        color: i % 2 === 0 ? WHITE : NAVY, align: "center", margin: 0,
      });
      s.addText(desc, {
        x: x + 0.1, y: 2.82, w: 1.95, h: 1.18,
        fontFace: "Calibri", fontSize: 10,
        color: i % 2 === 0 ? "B0BED4" : MIDGRAY, align: "left", margin: 4,
      });
    });

    // Options row
    const opts = [
      ["RIGHT-SIZE",  "Sell your landed, move to a condo — unlock $2M+ cash, reduce maintenance burden."],
      ["UPGRADE",     "Proceeds from your landed can fund a GCB or prime district detached (D10/D11)."],
      ["INVEST",      "Deploy proceeds into a portfolio of income-generating properties for retirement."],
    ];
    opts.forEach(([label, desc], i) => {
      const x = 0.5 + i * 3.1;
      s.addShape("rect", { x, y: 4.22, w: 2.9, h: 1.08,
        fill: { color: LIGHT }, line: { color: RULE } });
      s.addText(label, {
        x: x + 0.12, y: 4.3, w: 2.65, h: 0.28,
        fontFace: "Calibri", fontSize: 11, bold: true, color: GOLD,
        charSpacing: 1, margin: 0,
      });
      s.addText(desc, {
        x: x + 0.12, y: 4.62, w: 2.65, h: 0.6,
        fontFace: "Calibri", fontSize: 10, color: NAVY, margin: 0,
      });
    });

    s.addNotes("S.W.A.P framework from #KND ONE Lesson. Use the property-swap case study: e.g. Botanique at Bartley → Parc Esta generated ~$400K–$500K net gain over 5 years. Ask the owner: 'What is your next move?'");
  }

  // ══════════════════════════════════════════════════════════════════════
  // SLIDE 13 — TRACK RECORD
  // ══════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    addSlideHeader(s, "My Credentials", "My Track Record in Landed Property");

    // Recent transactions placeholder
    s.addText("RECENT TRANSACTIONS", {
      x: 0.5, y: 1.45, w: 5.5, h: 0.3,
      fontFace: "Calibri", fontSize: 10, bold: true, color: GOLD,
      charSpacing: 2, margin: 0,
    });

    const tHeader = { fill: { color: NAVY }, color: WHITE, bold: true, fontSize: 11 };
    const tRow1   = { fill: { color: LIGHT }, color: NAVY, fontSize: 11 };
    const tRow2   = { fill: { color: WHITE }, color: NAVY, fontSize: 11 };

    s.addTable([
      [
        { text: "Address / District", options: { ...tHeader } },
        { text: "Type",               options: { ...tHeader } },
        { text: "Price",              options: { ...tHeader } },
        { text: "DOM",                options: { ...tHeader } },
      ],
      [
        { text: "[INSERT: Address]",  options: { ...tRow1 } },
        { text: "[Type]",             options: { ...tRow1 } },
        { text: "$[X]M",              options: { ...tRow1 } },
        { text: "[X] days",           options: { ...tRow1 } },
      ],
      [
        { text: "[INSERT: Address]",  options: { ...tRow2 } },
        { text: "[Type]",             options: { ...tRow2 } },
        { text: "$[X]M",              options: { ...tRow2 } },
        { text: "[X] days",           options: { ...tRow2 } },
      ],
      [
        { text: "[INSERT: Address]",  options: { ...tRow1 } },
        { text: "[Type]",             options: { ...tRow1 } },
        { text: "$[X]M",              options: { ...tRow1 } },
        { text: "[X] days",           options: { ...tRow1 } },
      ],
    ], {
      x: 0.5, y: 1.82, w: 5.5, h: 2.1,
      colW: [2.2, 1.2, 1.2, 0.9],
      border: { pt: 1, color: RULE },
    });

    // Certifications
    s.addText("CREDENTIALS", {
      x: 0.5, y: 4.05, w: 5.5, h: 0.28,
      fontFace: "Calibri", fontSize: 10, bold: true, color: GOLD, charSpacing: 2, margin: 0,
    });
    bulletList(s, [
      "CEA Registered Salesperson (RES)",
      "ERA Landed Expert Series — all 5 modules completed",
      "ERA PropNett Network member",
      "Districts served: [INSERT districts]",
    ], 0.5, 4.38, 5.4, 1.0, 11);

    // Testimonial boxes
    s.addText("CLIENT TESTIMONIALS", {
      x: 6.2, y: 1.45, w: 3.3, h: 0.3,
      fontFace: "Calibri", fontSize: 10, bold: true, color: GOLD, charSpacing: 2, margin: 0,
    });

    [0, 1].forEach(i => {
      const y = 1.82 + i * 1.75;
      s.addShape("rect", { x: 6.2, y, w: 3.3, h: 1.55,
        fill: { color: PLHOLD }, line: { color: RULE },
        shadow: mkShadow(),
      });
      s.addText(`"[INSERT: Client testimonial quote — 1–2 sentences about the experience of selling through you.]"`, {
        x: 6.35, y: y + 0.12, w: 3.0, h: 0.95,
        fontFace: "Calibri", fontSize: 10, italic: true, color: NAVY, margin: 0,
      });
      s.addText("— [Client Name], [Property Type], [District]", {
        x: 6.35, y: y + 1.12, w: 3.0, h: 0.3,
        fontFace: "Calibri", fontSize: 9.5, bold: true, color: MIDGRAY, margin: 0,
      });
    });

    s.addNotes("Fill in your 3 most recent landed transactions before each presentation. Show actual DOM (days on market) — faster-than-average DOM is a powerful proof point for sellers.");
  }

  // ══════════════════════════════════════════════════════════════════════
  // SLIDE 14 — 7-POINT PROMISE
  // ══════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    addSlideHeader(s, "My Commitment", "My 7-Point Marketing Promise to You");

    const promises = [
      ["01", "Photography & Drone Video",         "Professional shoot within 5 days of appointment. 360°, gimbal interior, drone aerial, lifestyle video."],
      ["02", "Listing Goes Live in 7 Days",        "PropertyGuru Premium listing + ERA PropTrack activated within 1 week of signing."],
      ["03", "10+ Qualified Buyer Contacts/Week",  "Minimum 10 targeted outreach contacts to qualified buyers per week during active campaign."],
      ["04", "Weekly Written Seller Update",       "Viewings count, enquiry quality score, market feedback, recommended next steps — every week."],
      ["05", "Technical Advisory On-Demand",       "Road line, SIP/DIP, redevelopment Q&A for buyer technical questions — answered on the spot."],
      ["06", "Full Negotiation Representation",    "I represent your interests at OTP stage — pricing, conditions, timeline, deposit structure."],
      ["07", "After-Sale Support",                 "CPF refund coordination, key handover checklist, referrals for solicitors and movers."],
    ];

    promises.forEach(([num, title, desc], i) => {
      const col = i < 4 ? 0 : 1;
      const row = i < 4 ? i : i - 4;
      const x   = col === 0 ? 0.5 : 5.1;
      const y   = 1.42 + row * 1.02;

      s.addShape("rect", { x, y, w: 4.45, h: 0.88,
        fill: { color: col === 0 ? LIGHT : LTGOLD },
        line: { color: RULE, width: 0.5 },
        shadow: mkShadow(),
      });
      // Number circle
      s.addShape("ellipse", { x: x + 0.12, y: y + 0.2, w: 0.42, h: 0.42,
        fill: { color: NAVY }, line: { color: NAVY } });
      s.addText(num, {
        x: x + 0.12, y: y + 0.2, w: 0.42, h: 0.42,
        fontFace: "Calibri", fontSize: 10, bold: true, color: GOLD,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText(title, {
        x: x + 0.65, y: y + 0.08, w: 3.7, h: 0.3,
        fontFace: "Calibri", fontSize: 11.5, bold: true, color: NAVY, margin: 0,
      });
      s.addText(desc, {
        x: x + 0.65, y: y + 0.42, w: 3.7, h: 0.38,
        fontFace: "Calibri", fontSize: 9.5, color: MIDGRAY, margin: 0,
      });
    });

    s.addNotes("Present this as a signed commitment if you can — print it as a one-pager and leave it with the owner. Specific numbers (5 days, 7 days, 10 contacts/week) make the promise credible.");
  }

  // ══════════════════════════════════════════════════════════════════════
  // SLIDE 15 — FAQ
  // ══════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    addSlideHeader(s, "Common Questions", "Frequently Asked Questions");

    const faqs = [
      {
        q: "Do I need to move out before the property is marketed?",
        a: "No — we can market with you in residence. Professional staging and photography work around your schedule. Occupied properties often photograph better with furniture.",
      },
      {
        q: "How long will it realistically take to sell?",
        a: "Well-priced landed homes in active districts typically receive offers within 4–8 weeks. Total timeline from listing to keys: 3–6 months including legal completion.",
      },
      {
        q: "What if a buyer wants to redevelop my property?",
        a: "This broadens your buyer pool — redevelopment buyers are often cash-rich and motivated. I can advise on redevelopment potential, BCA submission process, and pricing strategy.",
      },
      {
        q: "What happens to my CPF monies when I sell?",
        a: "CPF OA principal used for the property + accrued interest (at CPF OA rate, currently 2.5% p.a.) must be fully refunded to your CPF OA upon sale. I will walk you through the exact calculation.",
      },
    ];

    faqs.forEach(({ q, a }, i) => {
      const col = i < 2 ? 0 : 1;
      const row = i < 2 ? i : i - 2;
      const x = col === 0 ? 0.5 : 5.1;
      const y = 1.42 + row * 1.98;

      s.addShape("rect", { x, y, w: 4.45, h: 1.85,
        fill: { color: col === 0 ? LIGHT : LTGOLD },
        line: { color: RULE, width: 0.5 },
        shadow: mkShadow(),
      });
      // Q label
      s.addShape("ellipse", { x: x + 0.12, y: y + 0.16, w: 0.38, h: 0.38,
        fill: { color: NAVY }, line: { color: NAVY } });
      s.addText("Q", {
        x: x + 0.12, y: y + 0.16, w: 0.38, h: 0.38,
        fontFace: "Calibri", fontSize: 11, bold: true, color: GOLD,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText(q, {
        x: x + 0.62, y: y + 0.1, w: 3.72, h: 0.54,
        fontFace: "Calibri", fontSize: 11, bold: true, color: NAVY, margin: 0,
      });
      // A label
      s.addShape("ellipse", { x: x + 0.12, y: y + 0.78, w: 0.38, h: 0.38,
        fill: { color: GOLD }, line: { color: GOLD } });
      s.addText("A", {
        x: x + 0.12, y: y + 0.78, w: 0.38, h: 0.38,
        fontFace: "Calibri", fontSize: 11, bold: true, color: NAVY,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText(a, {
        x: x + 0.62, y: y + 0.72, w: 3.72, h: 1.0,
        fontFace: "Calibri", fontSize: 10, color: MIDGRAY, margin: 0,
      });
    });

    s.addNotes("Use this slide to pre-empt objections. Print a version with your answers filled in — it positions you as someone who has thought through their concerns in advance.");
  }

  // ══════════════════════════════════════════════════════════════════════
  // SLIDE 16 — NEXT STEPS + BACK COVER
  // ══════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: NAVY };

    // Headline
    s.addText("Let's Get Started", {
      x: 0.5, y: 0.3, w: 9.0, h: 0.65,
      fontFace: "Calibri", fontSize: 34, bold: true, color: WHITE,
      align: "center", margin: 0,
    });
    s.addText("Three simple steps to finding out what your home is worth today.", {
      x: 0.5, y: 0.98, w: 9.0, h: 0.32,
      fontFace: "Calibri", fontSize: 13, italic: true, color: GOLD,
      align: "center", margin: 0,
    });

    // Steps
    const steps = [
      { n: "1", title: "Free Property Assessment",      desc: "I visit your home, assess condition and redevelopment potential, run recent comps, and give you an honest indicative valuation. No strings attached." },
      { n: "2", title: "Marketing Strategy Meeting",    desc: "We agree on the right price, launch timing, and marketing approach. I share the comparable transactions and buyer profiles most relevant to your home." },
      { n: "3", title: "Sign & Launch",                 desc: "We sign the exclusive listing agreement and I go to work — photography within 5 days, listing live within 7 days, buyer outreach begins immediately." },
    ];

    steps.forEach(({ n, title, desc }, i) => {
      const x = 0.5 + i * 3.1;
      s.addShape("rect", { x, y: 1.45, w: 2.9, h: 2.65,
        fill: { color: "162338" },
        line: { color: GOLD, width: 1.2 },
        shadow: mkShadow(),
      });
      s.addShape("ellipse", { x: x + 1.13, y: 1.55, w: 0.64, h: 0.64,
        fill: { color: GOLD }, line: { color: GOLD } });
      s.addText(n, {
        x: x + 1.13, y: 1.55, w: 0.64, h: 0.64,
        fontFace: "Calibri", fontSize: 22, bold: true, color: NAVY,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText(title, {
        x: x + 0.12, y: 2.28, w: 2.66, h: 0.44,
        fontFace: "Calibri", fontSize: 12, bold: true, color: WHITE,
        align: "center", margin: 0,
      });
      s.addText(desc, {
        x: x + 0.15, y: 2.76, w: 2.6, h: 1.25,
        fontFace: "Calibri", fontSize: 10, color: "B0BED4",
        align: "left", margin: 4,
      });
    });

    // Contact row
    s.addShape("rect", { x: 0.5, y: 4.25, w: 9.0, h: 1.1,
      fill: { color: "0F1D32" }, line: { color: "2E4166" } });

    s.addText("[INSERT: Agent Name]", {
      x: 0.75, y: 4.38, w: 3.2, h: 0.38,
      fontFace: "Calibri", fontSize: 15, bold: true, color: WHITE, margin: 0,
    });
    s.addText("[INSERT: CEA Reg. No.] | ERA Singapore", {
      x: 0.75, y: 4.78, w: 3.2, h: 0.28,
      fontFace: "Calibri", fontSize: 10, color: "B0BED4", margin: 0,
    });

    s.addText("[INSERT: Mobile / WhatsApp]", {
      x: 4.0, y: 4.38, w: 2.6, h: 0.32,
      fontFace: "Calibri", fontSize: 12, color: GOLD, margin: 0,
    });
    s.addText("[INSERT: Email Address]", {
      x: 4.0, y: 4.72, w: 2.6, h: 0.32,
      fontFace: "Calibri", fontSize: 12, color: GOLD, margin: 0,
    });

    s.addShape("rect", { x: 7.5, y: 4.3, w: 1.7, h: 1.0,
      fill: { color: "162338" }, line: { color: "2E4166" } });
    s.addText("ERA\nSINGAPORE", {
      x: 7.5, y: 4.3, w: 1.7, h: 1.0,
      fontFace: "Calibri", fontSize: 14, bold: true, color: GOLD,
      align: "center", valign: "middle", margin: 0,
    });

    s.addNotes("Close on this slide. Ask: 'When would be a good time for me to come by for the property assessment? It takes about 30 minutes and there is absolutely no obligation.' Silence after asking — let them answer.");
  }

  // ─── WRITE ──────────────────────────────────────────────────────────────────
  const outPath = "/Users/riopang/Desktop/Claude Code/Landed Seller Pitch Deck.pptx";
  await pres.writeFile({ fileName: outPath });
  console.log("Done →", outPath);
}

buildDeck().catch(err => { console.error(err); process.exit(1); });
