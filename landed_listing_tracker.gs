// ============================================================
//  Landed Buyer Listing Tracker — Google Apps Script
//  Paste this into Extensions > Apps Script > Code.gs
//  Then run: onOpen() or use the "Listing Tracker" menu
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Listing Tracker')
    .addItem('1. Setup / Reset Sheets', 'setupListingTracker')
    .addItem('2. Add New Listing Row', 'addListingRow')
    .addSeparator()
    .addItem('How to Add Photos', 'showPhotoInstructions')
    .toUi();
}

// ============================================================
//  MAIN SETUP
// ============================================================

function setupListingTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Remove existing sheets if re-running setup
  ['Listings', 'Internal'].forEach(name => {
    const s = ss.getSheetByName(name);
    if (s) ss.deleteSheet(s);
  });

  const listingSheet  = ss.insertSheet('Listings', 0);
  const internalSheet = ss.insertSheet('Internal', 1);

  setupListingsSheet(listingSheet);
  setupInternalSheet(internalSheet);

  // Land back on the Listings tab
  ss.setActiveSheet(listingSheet);

  SpreadsheetApp.getUi().alert(
    'Setup complete!\n\n' +
    'Share the LISTINGS tab only — go to the bottom sheet tab, right-click → "Copy link to this sheet". ' +
    'Your buyer sees Listings. Only you can see Internal.'
  );
}

// ============================================================
//  LISTINGS SHEET  (client-facing — no listing URLs here)
// ============================================================

function setupListingsSheet(sheet) {
  const DARK_NAVY  = '#1a1a2e';
  const WHITE      = '#ffffff';

  const headers = [
    'Ref', 'Type', 'Area', 'Street / Vicinity',
    'Asking Price', 'Land (sqft)', 'Built-up (sqft)', 'Land PSF',
    'Tenure', 'Bed', 'Bath', 'Carparks',
    'Key Highlights', "Agent's Take", 'Status',
    'Photo 1', 'Photo 2', 'Photo 3'
  ];

  // ── Headers ──────────────────────────────────────────────
  const hdr = sheet.getRange(1, 1, 1, headers.length);
  hdr.setValues([headers])
     .setBackground(DARK_NAVY)
     .setFontColor(WHITE)
     .setFontWeight('bold')
     .setFontSize(11)
     .setVerticalAlignment('middle')
     .setHorizontalAlignment('center');
  sheet.setRowHeight(1, 42);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);

  // ── Column widths ─────────────────────────────────────────
  const widths = [60, 120, 100, 160, 130, 100, 110, 90, 90, 50, 50, 80, 220, 220, 120, 185, 185, 185];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // ── Sample data ───────────────────────────────────────────
  const rows = [
    ['L001', 'Inter Terrace', 'Serangoon',   'Lorong Chuan area',  2800000, 1600, 2200, '', 'Freehold',  4, 3, 2, 'Renovated, good facing, quiet street',       'Strong buy — priced 5% below recent transacted',   'Shortlisted',     '', '', ''],
    ['L002', 'Semi-D',        'Bukit Timah', 'Coronation Road area', 5500000, 3500, 4800, '', '999-year',  5, 5, 3, 'Original condition, large land, top school belt', 'Worth viewing — land size excellent for the price', 'Available',       '', '', ''],
    ['L003', 'Detached',      'Holland Road','Sixth Ave area',      9800000, 6200, 7500, '', 'Freehold',  6, 6, 4, 'Modern reno, pool, smart home',              'Premium but fully turnkey — no reno costs',         'Under Offer',     '', '', ''],
  ];
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  // PSF formula (col 8 = H) — price per sqft of land
  for (let r = 2; r <= 101; r++) {
    sheet.getRange(r, 8).setFormula(`=IF(F${r}="","",E${r}/F${r})`);
  }

  // ── Number formats ────────────────────────────────────────
  sheet.getRange('E2:E101').setNumberFormat('"S$"#,##0');
  sheet.getRange('F2:G101').setNumberFormat('#,##0 "sqft"');
  sheet.getRange('H2:H101').setNumberFormat('"S$"#,##0 "psf"');

  // ── Dropdowns ─────────────────────────────────────────────
  setDropdown(sheet, 'B2:B101', ['Inter Terrace', 'Corner Terrace', 'Semi-D', 'Detached', 'Cluster House', 'Good Class Bungalow']);
  setDropdown(sheet, 'I2:I101', ['Freehold', '999-year', '99-year']);
  setDropdown(sheet, 'O2:O101', ['Available', 'Shortlisted', 'Viewing Arranged', 'Under Offer', 'Closed (Not Suitable)', 'Sold']);

  // ── Status conditional formatting ────────────────────────
  const rules = [
    { text: 'Shortlisted',         bg: '#c8e6c9', fg: '#1b5e20' },
    { text: 'Available',           bg: '#e3f2fd', fg: '#0d47a1' },
    { text: 'Viewing Arranged',    bg: '#f3e5f5', fg: '#4a148c' },
    { text: 'Under Offer',         bg: '#fff3e0', fg: '#e65100' },
    { text: 'Closed (Not Suitable)', bg: '#fce4ec', fg: '#880e4f' },
    { text: 'Sold',                bg: '#f5f5f5', fg: '#9e9e9e' },
  ];
  const statusRange = [sheet.getRange('O2:O101')];
  sheet.setConditionalFormatRules(rules.map(r =>
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(r.text)
      .setBackground(r.bg).setFontColor(r.fg)
      .setRanges(statusRange).build()
  ));

  // ── Row banding & height ──────────────────────────────────
  sheet.getRange(2, 1, 100, headers.length)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);

  for (let i = 2; i <= 30; i++) sheet.setRowHeight(i, 130);

  // ── Text wrap for notes columns ───────────────────────────
  sheet.getRange('M2:N101').setWrap(true).setVerticalAlignment('top');

  // ── Align photo cells ─────────────────────────────────────
  sheet.getRange('P2:R101').setVerticalAlignment('middle').setHorizontalAlignment('center');

  // ── Column notes (hover hints) ────────────────────────────
  sheet.getRange('P1').setNote(
    'How to add a photo:\n' +
    '1. Upload image to Google Drive\n' +
    '2. Right-click → Get link → "Anyone with the link"\n' +
    '3. Copy the File ID (long string between /d/ and /view)\n' +
    '4. In this cell paste:\n' +
    '=IMAGE("https://drive.google.com/uc?export=view&id=YOUR_FILE_ID")'
  );
  sheet.getRange('Q1').setNote('Photo 2 — same formula as Photo 1');
  sheet.getRange('R1').setNote('Photo 3 — optional');
}

// ============================================================
//  INTERNAL SHEET  (you only — never share this tab URL)
// ============================================================

function setupInternalSheet(sheet) {
  // Warning banner row
  sheet.insertRowBefore(1);
  const warning = sheet.getRange(1, 1, 1, 8);
  warning.merge()
    .setValue('⚠️  INTERNAL USE ONLY — DO NOT SHARE THIS TAB URL WITH BUYERS')
    .setBackground('#ffcdd2').setFontColor('#b71c1c')
    .setFontWeight('bold').setFontSize(12)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);

  const headers = ['Ref', 'Full Address', 'Listing URL', 'Seller Agent', 'Agent Contact', 'Commission (%)', 'Notes'];
  const hdr = sheet.getRange(2, 1, 1, headers.length);
  hdr.setValues([headers])
     .setBackground('#b71c1c').setFontColor('#ffffff')
     .setFontWeight('bold').setFontSize(11)
     .setVerticalAlignment('middle').setHorizontalAlignment('center');
  sheet.setRowHeight(2, 40);
  sheet.setFrozenRows(2);

  const widths = [60, 220, 300, 150, 140, 110, 280];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  const rows = [
    ['L001', '12 Lorong Chuan, S556789',    'https://propertyguru.com.sg/listing/...', 'David Tan', '9123 4567', 1.5, 'Owner flexible, open to $2.65M. Call before arranging viewing.'],
    ['L002', '45 Coronation Rd, S269475',   'https://99.co/singapore/sale/...',        'Sarah Lim', '9876 5432', 1.0, 'Estate sale — 3 siblings involved, decisions take time.'],
    ['L003', '8 Sixth Avenue, S276479',     'https://edgeprop.sg/property/...',        'James Wong','8765 4321', 1.0, 'Firm on price. Developer interest — act fast if buyer is serious.'],
  ];
  sheet.getRange(3, 1, rows.length, headers.length).setValues(rows);
  sheet.getRange('G3:G101').setWrap(true).setVerticalAlignment('top');

  sheet.protect().setDescription('Internal — listing URLs and agent contacts').setWarningOnly(true);
}

// ============================================================
//  ADD NEW LISTING ROW
// ============================================================

function addListingRow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ls = ss.getSheetByName('Listings');
  const is = ss.getSheetByName('Internal');

  if (!ls || !is) {
    SpreadsheetApp.getUi().alert('Run Setup first from the Listing Tracker menu.');
    return;
  }

  const dataRows = ls.getLastRow() - 1;        // subtract header
  const ref = 'L' + String(dataRows + 1).padStart(3, '0');
  const newRow = ls.getLastRow() + 1;

  ls.appendRow([ref, '', '', '', '', '', '', '', '', '', '', '', '', '', 'Available', '', '', '']);
  ls.getRange(newRow, 8).setFormula(`=IF(F${newRow}="","",E${newRow}/F${newRow})`);
  ls.setRowHeight(newRow, 130);

  is.appendRow([ref, '', '', '', '', '', '']);

  ss.setActiveSheet(ls);
  ls.setActiveRange(ls.getRange(newRow, 1));

  SpreadsheetApp.getUi().alert('New listing row added: ' + ref + '\n\nFill in the Listings tab for your buyer, then add the URL and agent details in the Internal tab.');
}

// ============================================================
//  PHOTO INSTRUCTIONS POPUP
// ============================================================

function showPhotoInstructions() {
  SpreadsheetApp.getUi().alert(
    'How to add photos to a listing',
    'Step 1 — Upload the property photo to Google Drive.\n\n' +
    'Step 2 — Right-click the file → "Get link" → change access to "Anyone with the link can view".\n\n' +
    'Step 3 — Copy the File ID. It\'s the long string in the URL between "/d/" and "/view"\n' +
    '   Example URL: drive.google.com/file/d/1aBcDeFgHiJkLmNoPqRsTuV/view\n' +
    '   File ID: 1aBcDeFgHiJkLmNoPqRsTuV\n\n' +
    'Step 4 — In the Photo 1 / 2 / 3 cell, type this formula:\n' +
    '=IMAGE("https://drive.google.com/uc?export=view&id=PASTE_FILE_ID_HERE")\n\n' +
    'The photo will display inside the cell. Row height is pre-set to 130px.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ============================================================
//  HELPER
// ============================================================

function setDropdown(sheet, range, values) {
  sheet.getRange(range).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(values).build()
  );
}
