/**
 * Trims the empty rows at the bottom of every sheet down to its row limit:
 * the rowLimit of its shop (config.gs: My home 250, Соронзон 400), or
 * TRIM_DEFAULT_ROW_LIMIT for any other sheet.
 *
 * Google Sheets has no setting that forbids more rows: inserting a row simply
 * grows the sheet. What can be done is removing the unused empty rows (they
 * count toward the 10 million cell limit and slow the file down). A sheet is
 * never cut above its last used row: it keeps max(limit, lastRow + spare).
 *
 * A limit of 0 switches it off for those sheets.
 *
 * Usage: run trimSheetsToLimit (DRY_RUN in config.gs decides whether it only
 * reports). The background job (year-sync.gs) runs it too, and instant-sync.gs
 * right after a row is inserted.
 */

const TRIM_DEFAULT_ROW_LIMIT = 400;
const TRIM_SPARE_ROWS = 20;
const TRIM_OFF = 0; // a row limit of 0 = never remove rows

// ---------- pure helpers (no Spreadsheet calls) ----------

function rowLimitFor(sheetName) {
  const slot = sheetSlot(sheetName);
  if (slot === null) {
    return TRIM_DEFAULT_ROW_LIMIT;
  }
  return CHAINS.find((chain) => chain.key === slot.chainKey).rowLimit;
}

// The last row that holds something a person can see. values: all cells of the sheet.
function lastFilledRow(values) {
  for (let index = values.length - 1; index >= 0; index -= 1) {
    if (values[index].some((value) => value !== '')) {
      return index + 1;
    }
  }
  return 0;
}

// ---------- spreadsheet access ----------

// sheet.getLastRow() is not used: the array formulas of AL:AN (calc-formulas.gs)
// fill every row down to the bottom with an empty text, and Google Sheets counts
// those as used rows, so the sheet would look full and never be cut.
function findLastDataRow(sheet) {
  const rows = sheet.getMaxRows();
  const columns = sheet.getMaxColumns();
  return rows > 0 && columns > 0 ? lastFilledRow(sheet.getRange(1, 1, rows, columns).getValues()) : 0;
}

function trimSheetsToLimit() {
  const lines = runTrim(currentBook());
  Logger.log('%s\n%s', modeLabel(), lines.join('\n'));
}

function runTrim(spreadsheet) {
  return spreadsheet
    .getSheets()
    .filter((sheet) => !sheet.getName().startsWith('_'))
    .map(trimSheet);
}

function trimSheet(sheet) {
  const name = sheet.getName();
  if (isStockSheet(sheet)) {
    return `${name}: тооллогын хуудас - мөр хасахгүй`;
  }
  const limit = rowLimitFor(name);
  if (limit === TRIM_OFF) {
    return `${name}: мөр хасахгүй (хязгаар 0)`;
  }
  const lastDataRow = findLastDataRow(sheet);
  const rowsBefore = sheet.getMaxRows();
  const keepRows = Math.max(limit, lastDataRow + TRIM_SPARE_ROWS);
  const extraRows = rowsBefore - keepRows;
  if (extraRows <= 0) {
    return `${name}: ${rowsBefore} мөр, хасах зүйлгүй`;
  }
  if (!DRY_RUN) {
    sheet.deleteRows(keepRows + 1, extraRows);
  }
  return `${name}: ${rowsBefore} -> ${keepRows} мөр (сүүлийн өгөгдөлтэй мөр ${lastDataRow})`;
}
