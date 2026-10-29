/**
 * Keeps the header of every year sheet (rows 1-3) as it is: the label and day
 * numbers of row 1 (not the seller's name in B1, which is typed by hand), the title of row 2, and the labels of row 3 ("Барааны нэрс",
 * "анхны үлдэгдэл", "өртөг үнэ", "зарах үнэ", "Ор", "Зарлага", "тооллого", ...)
 * together with the formulas that live there (AL3:AN3, AO3:AS3).
 *
 * Rows 1-3 are protected, so another editor cannot change them at all. The owner
 * can - by a slip of the hand as well. So the header of each sheet is remembered
 * (document property HEADER_KEY + sheet id), and a header cell that was edited
 * or emptied gets its remembered content back:
 *   - at once, after an edit that touches rows 1-3 (instant-sync.gs)
 *   - every hour, by the background job (year-sync.gs)
 *
 * setupYear remembers the headers at its end (step 15), after it has put dates,
 * titles and formulas in place. To change a header on purpose: change it, then
 * run rememberHeaders before the hourly job runs - or simply run setupYear.
 * DRY_RUN (config.gs) makes everything report only. Needs config.gs.
 */

const HEADER_KEY = 'header:'; // + sheet id: JSON rows 1-3 of that sheet
const HEADER_ROWS = 3;
const HEADER_MAX_COLUMNS = LAYOUT.summaryFirst + TOTAL_COUNT - 1; // A..AS
// Cells of the header that are typed by hand and therefore never put back.
const HEADER_INPUT_CELLS = [
  { row: 1, column: 1 }, // A1: which columns are shown (view-columns.gs)
  { row: 1, column: 2 }, // B1: the seller's name
];
// Labels every sheet must have; an empty cell gets the text when the header is remembered.
const HEADER_REQUIRED_LABELS = [{ row: 3, column: 2, text: 'анхны үлдэгдэл ' }]; // B3

function rememberHeaders() {
  const lines = runRememberHeaders(currentBook());
  Logger.log('%s\n%s', modeLabel(), lines.join('\n'));
}

function restoreHeaders() {
  const lines = runRestoreHeaders(currentBook());
  Logger.log('%s\n%s', modeLabel(), lines.length > 0 ? lines.join('\n') : 'гарчгийн мөрүүд хэвээр байна');
}

// ---------- pure helpers (no Spreadsheet calls) ----------

function isDateValue(value) {
  return Object.prototype.toString.call(value) === '[object Date]';
}

// A cell as it is kept in the note: its formula when it has one, otherwise its
// value. A date becomes { date: milliseconds }, because JSON has no dates.
function encodeHeaderCell(value, formula) {
  if (formula !== '') {
    return formula;
  }
  return isDateValue(value) ? { date: value.getTime() } : value;
}

function decodeHeaderCell(cell) {
  return cell !== null && typeof cell === 'object' ? new Date(cell.date) : cell;
}

function encodeHeader(values, formulas) {
  return values.map((row, r) => row.map((value, c) => encodeHeaderCell(value, formulas[r][c])));
}

function isSameHeader(a, b) {
  return JSON.stringify(a) === JSON.stringify(b);
}

// The remembered header, cut or padded to the width the sheet has now.
function fitHeader(saved, width) {
  return saved.map((row) => Array.from({ length: width }, (_, c) => (c < row.length ? row[c] : '')));
}

// Before the header is remembered anew: a cell that is empty now but was not
// gets its old content back, so a deleted label is never remembered as "empty".
function fillEmptiedCells(current, saved) {
  return current.map((row, r) => row.map((cell, c) => (cell === '' ? saved[r][c] : cell)));
}

function fillRequiredLabels(header) {
  const labelAt = (r, c) => HEADER_REQUIRED_LABELS.find((label) => label.row === r + 1 && label.column === c + 1);
  return header.map((row, r) =>
    row.map((cell, c) => (cell === '' && labelAt(r, c) !== undefined ? labelAt(r, c).text : cell))
  );
}

// The header to aim for: as remembered, except the hand-typed cells, which stay as they are now.
function keepInputCells(header, current) {
  const isInput = (r, c) => HEADER_INPUT_CELLS.some((cell) => cell.row === r + 1 && cell.column === c + 1);
  return header.map((row, r) => row.map((cell, c) => (isInput(r, c) ? current[r][c] : cell)));
}

// ---------- spreadsheet access ----------

function headerRange(sheet) {
  return sheet.getRange(1, 1, HEADER_ROWS, Math.min(sheet.getMaxColumns(), HEADER_MAX_COLUMNS));
}

function readHeader(sheet) {
  const range = headerRange(sheet);
  return encodeHeader(range.getValues(), range.getFormulas());
}

function readSavedHeader(sheet) {
  const saved = bookProperties().getProperty(HEADER_KEY + sheet.getSheetId());
  try {
    const parsed = saved ? JSON.parse(saved) : null;
    return Array.isArray(parsed) && parsed.length === HEADER_ROWS ? parsed : null;
  } catch (error) {
    return null; // unreadable note: the header is remembered again by setupYear
  }
}

function writeHeader(sheet, header) {
  // setValues stores a text that starts with "=" as a formula.
  headerRange(sheet).setValues(header.map((row) => row.map(decodeHeaderCell)));
}

function listHeaderSheets(spreadsheet) {
  return spreadsheet.getSheets().filter((sheet) => sheetSlot(sheet.getName()) !== null);
}

// Puts an edited or emptied header cell back. Returns a log line, or null when
// the header is as remembered (or was never remembered).
function restoreHeader(sheet) {
  const saved = readSavedHeader(sheet);
  if (saved === null) {
    return null;
  }
  const current = readHeader(sheet);
  const wanted = keepInputCells(fitHeader(saved, current[0].length), current);
  if (isSameHeader(current, wanted)) {
    return null;
  }
  if (!DRY_RUN) {
    writeHeader(sheet, wanted);
  }
  return `${sheet.getName()}: гарчгийн мөрийн өөрчлөгдсөн бичвэрийг ${DRY_RUN ? 'сэргээнэ' : 'сэргээлээ'}`;
}

function runRestoreHeaders(spreadsheet) {
  return listHeaderSheets(spreadsheet)
    .map(restoreHeader)
    .filter((line) => line !== null);
}

// Remembers the header of one sheet as it is now; a label that was deleted since
// the last time is put back first. Returns true when the note changed.
function rememberHeader(sheet) {
  const saved = readSavedHeader(sheet);
  const current = readHeader(sheet);
  const header = fillRequiredLabels(
    saved === null
      ? current
      : keepInputCells(fillEmptiedCells(current, fitHeader(saved, current[0].length)), current)
  );
  if (saved !== null && isSameHeader(saved, header)) {
    return false;
  }
  if (!DRY_RUN) {
    if (!isSameHeader(current, header)) {
      writeHeader(sheet, header);
    }
    bookProperties().setProperty(
      HEADER_KEY + sheet.getSheetId(),
      JSON.stringify(header)
    );
  }
  return true;
}

function runRememberHeaders(spreadsheet) {
  const changed = listHeaderSheets(spreadsheet).filter(rememberHeader).length;
  return [`гарчгийн мөрүүд (1-3): ${changed} sheet-ийнхийг ${DRY_RUN ? 'санана' : 'санаж авлаа'}`];
}
