// ===== SETTINGS =====
const SOURCE_SHEET_NAME  = "";       // sheet name; "" = currently open sheet
const HEADER_ROW         = 1;
const FIRST_DATA_ROW     = 2;
const COL_SRC_FIRST_DATE = 17;       // Q (date columns are searched from here to the right)

// Columns copied as is (values, colors, dropdowns, checkboxes, links)
const COPY_FIRST_COL = 1;            // A
const COPY_LAST_COL  = 15;           // O

// Target table
const TARGET_SPREADSHEET_URL = "";  // "" = same file; or URL / ID of another spreadsheet
const TARGET_SHEET_PREFIX    = "Sorted"; // target sheet name (created automatically if missing)
const TARGET_PER_SOURCE      = true;     // true = "Sorted - <source name>", one target per source sheet
const COL_OUT_FIRST_DONOR    = 17;  // Q
const COL_OUT_STEP           = 2;   // donor in every 2nd column, its error goes right next to it
const DONOR_HEADER           = "Ready Link"; // becomes "Ready Link 1", "Ready Link 2"...
const ERROR_HEADER           = "Error";      // becomes "Error 1", "Error 2"...
const DONOR_COL_WIDTH        = 220;
const ERROR_COL_WIDTH        = 100;

const SKIP_ROWS_WITHOUT_DONORS = true; // true = skip rows with no links

// ===== MENU =====
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu("Link Collector")
      .addItem("📋 Collect Links", "collectLinks")
      .addToUi();
  } catch (e) {}
}

// ===== MAIN =====
function collectLinks() {
  const ui = SpreadsheetApp.getUi();
  const srcSS = SpreadsheetApp.getActiveSpreadsheet();
  const src = SOURCE_SHEET_NAME ? srcSS.getSheetByName(SOURCE_SHEET_NAME) : srcSS.getActiveSheet();
  if (!src) { ui.alert("Source sheet not found: " + SOURCE_SHEET_NAME); return; }

  const TARGET_SHEET_NAME = TARGET_PER_SOURCE
    ? TARGET_SHEET_PREFIX + " - " + src.getName()
    : TARGET_SHEET_PREFIX;

  let tgtSS;
  try {
    tgtSS = openTarget_(srcSS);
  } catch (e) {
    ui.alert("Can't open target table. Check the URL and access.");
    return;
  }
  const sameFile = tgtSS.getId() === srcSS.getId();

  let tgt = tgtSS.getSheetByName(TARGET_SHEET_NAME);
  if (tgt && sameFile && tgt.getSheetId() === src.getSheetId()) {
    ui.alert("Target sheet is the same as source. Change TARGET_SHEET_PREFIX.");
    return;
  }

  const lastRow = src.getLastRow();
  const lastCol = src.getLastColumn();
  const numRows = lastRow - FIRST_DATA_ROW + 1;
  if (numRows < 1 || lastCol < COL_SRC_FIRST_DATE) {
    ui.alert("No data to collect.");
    return;
  }

  const width = lastCol - COL_SRC_FIRST_DATE + 1;
  const headers = src.getRange(HEADER_ROW, COL_SRC_FIRST_DATE, 1, width).getDisplayValues()[0];
  const dateIdx = [];
  const isDate = [];
  for (let c = 0; c < headers.length; c++) {
    isDate[c] = headerHasDate_(headers[c]);
    if (isDate[c]) dateIdx.push(c);
  }
  if (dateIdx.length === 0) {
    ui.alert("No date columns found from column " + colLetter_(COL_SRC_FIRST_DATE) + ".");
    return;
  }

  const block = src.getRange(FIRST_DATA_ROW, COL_SRC_FIRST_DATE, numRows, width);
  const rich = block.getRichTextValues();
  const forms = block.getFormulas();
  const disp = block.getDisplayValues();
  const bgs = block.getBackgrounds();

  const out = [];
  let maxDonors = 0, totalDonors = 0, totalErrors = 0, skipped = 0;
  for (let i = 0; i < numRows; i++) {
    const donors = [];
    for (let k = 0; k < dateIdx.length; k++) {
      const c = dateIdx[k];
      const url = extractUrl_(rich[i][c], forms[i][c], disp[i][c]);
      if (!url) continue;

      let err = "";
      if (c + 1 < width && !isDate[c + 1]) {
        const t = String(disp[i][c + 1] || "").trim();
        if (t && !isHttp_(t)) err = t;
      }
      if (err) totalErrors++;
      donors.push({ url: url, bg: bgs[i][c], err: err });
    }
    if (donors.length === 0 && SKIP_ROWS_WITHOUT_DONORS) { skipped++; continue; }

    out.push({ srcRow: FIRST_DATA_ROW + i, donors: donors });
    totalDonors += donors.length;
    if (donors.length > maxDonors) maxDonors = donors.length;
  }

  if (totalDonors === 0) { ui.alert("No links found in date columns."); return; }

  const confirm = ui.alert(
    "Confirm",
    "Date columns: " + dateIdx.length +
    "\nRows: " + out.length +
    "\nLinks: " + totalDonors +
    "\nErrors: " + totalErrors +
    "\nMax links per row: " + maxDonors +
    (skipped ? "\nSkipped empty rows: " + skipped : "") +
    "\n\n" + (tgt ? "Sheet '" + TARGET_SHEET_NAME + "' will be cleared and refilled." : "Sheet '" + TARGET_SHEET_NAME + "' will be created.") +
    " Continue?",
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  if (!tgt) tgt = tgtSS.insertSheet(TARGET_SHEET_NAME);

  const lastDonorCol = COL_OUT_FIRST_DONOR + (maxDonors - 1) * COL_OUT_STEP;
  const neededRows = FIRST_DATA_ROW + out.length - 1;
  const neededCols = Math.max(COPY_LAST_COL, lastDonorCol + 1);
  if (tgt.getMaxRows() < neededRows) tgt.insertRowsAfter(tgt.getMaxRows(), neededRows - tgt.getMaxRows());
  if (tgt.getMaxColumns() < neededCols) tgt.insertColumnsAfter(tgt.getMaxColumns(), neededCols - tgt.getMaxColumns());

  tgt.getRange(1, 1, tgt.getMaxRows(), tgt.getMaxColumns()).clear().clearDataValidations();

  copyColumns_(src, tgt, tgtSS, sameFile, out);
  for (let c = COPY_FIRST_COL; c <= COPY_LAST_COL; c++) {
    tgt.setColumnWidth(c, src.getColumnWidth(c));
  }

  for (let k = 0; k < maxDonors; k++) {
    const col = COL_OUT_FIRST_DONOR + k * COL_OUT_STEP;
    const vals = out.map(function (r) { return [r.donors[k] ? r.donors[k].url : ""]; });
    const colors = out.map(function (r) { return [r.donors[k] ? bgOrNull_(r.donors[k].bg) : null]; });
    const errs = out.map(function (r) { return [r.donors[k] ? r.donors[k].err : ""]; });

    tgt.getRange(FIRST_DATA_ROW, col, out.length, 1)
      .setNumberFormat("@")
      .setValues(vals)
      .setBackgrounds(colors);

    tgt.getRange(FIRST_DATA_ROW, col + 1, out.length, 1)
      .setNumberFormat("@")
      .setValues(errs)
      .setHorizontalAlignment("left");

    const firstHeader = tgt.getRange(HEADER_ROW, COPY_FIRST_COL);
    const hDonor = tgt.getRange(HEADER_ROW, col);
    const hErr = tgt.getRange(HEADER_ROW, col + 1);
    firstHeader.copyTo(hDonor, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    firstHeader.copyTo(hErr, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    hDonor.setValue(DONOR_HEADER + " " + (k + 1));
    hErr.setValue(ERROR_HEADER + " " + (k + 1));

    tgt.setColumnWidth(col, DONOR_COL_WIDTH);
    tgt.setColumnWidth(col + 1, ERROR_COL_WIDTH);
  }

  SpreadsheetApp.flush();
  ui.alert("Done! Rows: " + out.length + ", links: " + totalDonors + ", errors: " + totalErrors + ".");
}

// ===== HELPERS =====
function copyColumns_(src, tgt, tgtSS, sameFile, out) {
  let from = src, temp = null;
  if (!sameFile) {
    temp = src.copyTo(tgtSS);
    from = temp;
  }
  try {
    const nCols = COPY_LAST_COL - COPY_FIRST_COL + 1;

    from.getRange(HEADER_ROW, COPY_FIRST_COL, 1, nCols)
      .copyTo(tgt.getRange(HEADER_ROW, COPY_FIRST_COL));

    let i = 0;
    while (i < out.length) {
      let j = i;
      while (j + 1 < out.length && out[j + 1].srcRow === out[j].srcRow + 1) j++;
      from.getRange(out[i].srcRow, COPY_FIRST_COL, j - i + 1, nCols)
        .copyTo(tgt.getRange(FIRST_DATA_ROW + i, COPY_FIRST_COL));
      i = j + 1;
    }
  } finally {
    if (temp) tgtSS.deleteSheet(temp);
  }
}

function bgOrNull_(color) {
  const s = String(color || "").toLowerCase();
  return (!s || s === "#ffffff") ? null : s;
}

function openTarget_(srcSS) {
  const t = String(TARGET_SPREADSHEET_URL || "").trim();
  if (!t) return srcSS;
  return /^https?:\/\//i.test(t) ? SpreadsheetApp.openByUrl(t) : SpreadsheetApp.openById(t);
}

function extractUrl_(richCell, formula, display) {
  if (richCell) {
    let u = richCell.getLinkUrl();
    if (isHttp_(u)) return u.trim();
    const runs = richCell.getRuns();
    for (let k = 0; k < runs.length; k++) {
      u = runs[k].getLinkUrl();
      if (isHttp_(u)) return u.trim();
    }
  }
  const m = String(formula || "").match(/^=\s*HYPERLINK\(\s*"([^"]+)"/i);
  if (m && isHttp_(m[1])) return m[1].trim();
  const t = String(display || "").trim();
  return isHttp_(t) ? t : "";
}

function isHttp_(s) {
  return /^https?:\/\//i.test(String(s || "").trim());
}

function headerHasDate_(header) {
  const parts = String(header || "").split("/");
  for (let i = 0; i < parts.length; i++) {
    const m = parts[i].trim().match(/^(\d{1,2})[.\-](\d{1,2})(?:[.\-]\d{2,4})?$/);
    if (m) {
      const d = parseInt(m[1], 10), mo = parseInt(m[2], 10);
      if (d >= 1 && d <= 31 && mo >= 1 && mo <= 12) return true;
    }
  }
  return false;
}

function colLetter_(num) {
  let s = "";
  while (num > 0) {
    const mod = (num - 1) % 26;
    s = String.fromCharCode(65 + mod) + s;
    num = Math.floor((num - 1) / 26);
  }
  return s;
}