/**
 * Protects the header rows (1-3) on every sheet so columns cannot be
 * deleted or inserted by accident. Rows 4+ (product rows, columns A:AJ and
 * the rest) stay fully editable: rows can be added and deleted freely.
 *
 * Why this works: deleting or inserting a column crosses rows 1-3, which are
 * protected across the full sheet width. Deleting rows 4+ touches no
 * protected cell, so it is never blocked. A sheet that holds protected cells
 * also cannot be deleted by anyone but the file owner.
 *
 * Two cells of the header are left open on the year sheets, so the person doing
 * the data entry can use them: A1, the list that shows or hides columns, and B1,
 * the seller's name. Row 1 still cannot be deleted (the cells to their right are
 * protected), and neither can columns A and B (rows 2-3 are).
 *
 * Usage: run protectAllSheets (DRY_RUN in config.gs decides whether it only
 * reports). Sheets that are already protected are skipped. Run
 * removeAllProtections to undo. Only protections created by this script are
 * ever touched.
 */

const PROTECTION_DESCRIPTION = 'Гарчгийн мөр (1-3) - багана устгахаас хамгаалсан';
const FIRST_OPEN_ROW = 4;
// Open cells of row 1 on the year sheets: A1 the list that shows or hides columns
// (view-columns.gs), B1 the seller's name.
const PROTECTION_OPEN_ROW = 1;
const PROTECTION_OPEN_COLUMNS = [1, 2];

// false = hard lock: only the owner can edit rows 1-3 and nobody else can
//         delete or insert columns. The person doing the data entry must be
//         a separate editor, not the file owner (owners are never restricted).
// true  = Google shows an English "edit anyway?" popup instead of blocking.
const WARNING_ONLY = false;

function protectAllSheets() {
  const lines = runProtectHeaders(currentBook());
  Logger.log('%s\n%s', modeLabel(), lines.join('\n'));
}

function removeAllProtections() {
  const sheets = currentBook().getSheets();
  const removed = sheets.reduce((total, sheet) => total + removeOwnProtections(sheet), 0);
  Logger.log('Энэ script-ийн тавьсан %s хамгаалалтыг авлаа', removed);
}

function runProtectHeaders(spreadsheet) {
  return spreadsheet
    .getSheets()
    .filter((sheet) => !sheet.getName().startsWith('_'))
    .map(protectHeaderRows);
}

// ---------- pure helpers (no Spreadsheet calls) ----------

// The parts of rows 1-3 to protect, as { row, column, rows, columns }. A year sheet
// keeps the open cells free: row 1 is protected in the stretches between them.
function headerProtectionRanges(isYearSheet, maxColumns) {
  const headerRows = FIRST_OPEN_ROW - 1;
  const open = PROTECTION_OPEN_COLUMNS.filter((column) => column < maxColumns);
  if (!isYearSheet || open.length === 0) {
    return [{ row: 1, column: 1, rows: headerRows, columns: maxColumns }];
  }
  const row = PROTECTION_OPEN_ROW;
  const starts = [1, ...open.map((column) => column + 1)];
  const ends = [...open.map((column) => column - 1), maxColumns];
  const stretches = starts
    .map((first, index) => ({ row, column: first, rows: 1, columns: ends[index] - first + 1 }))
    .filter((area) => area.columns > 0);
  return [...stretches, { row: row + 1, column: 1, rows: headerRows - row, columns: maxColumns }];
}

function isSameArea(range, area) {
  return (
    range.getRow() === area.row &&
    range.getColumn() === area.column &&
    range.getNumRows() === area.rows &&
    range.getNumColumns() === area.columns
  );
}

// ---------- spreadsheet access ----------

function wantedHeaderAreas(sheet) {
  return headerProtectionRanges(sheetSlot(sheet.getName()) !== null, sheet.getMaxColumns());
}

function protectHeaderRows(sheet) {
  const name = sheet.getName();
  if (isStockSheet(sheet)) {
    return `${name}: тооллогын хуудас - бүхэлдээ түгжээтэй, алгаслаа`;
  }
  if (isHeaderProtected(sheet)) {
    return `${name}: аль хэдийн хамгаалагдсан`;
  }
  const areas = wantedHeaderAreas(sheet);
  if (!DRY_RUN) {
    removeOwnProtections(sheet);
    areas.forEach((area) => {
      const protection = sheet
        .getRange(area.row, area.column, area.rows, area.columns)
        .protect()
        .setDescription(PROTECTION_DESCRIPTION);
      protection.setWarningOnly(WARNING_ONLY);
      if (!WARNING_ONLY) {
        restrictToOwner(protection);
      }
    });
  }
  const open = areas.length > 1 ? ' (A1 харах сонголт, B1 худалдагчийн нэр нээлттэй)' : '';
  return `${name}: 1-3-р мөрийг ${DRY_RUN ? 'хамгаална' : 'хамгааллаа'}${open}`;
}

// True when this script's protections cover exactly the wanted parts, no more, no less.
function isHeaderProtected(sheet) {
  const own = findOwnHeaderProtections(sheet);
  const areas = wantedHeaderAreas(sheet);
  return (
    own.length === areas.length &&
    areas.every((area) =>
      own.some(
        (protection) =>
          isSameArea(protection.getRange(), area) && protection.isWarningOnly() === WARNING_ONLY
      )
    )
  );
}

function restrictToOwner(protection) {
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
}

function findOwnHeaderProtections(sheet) {
  return sheet
    .getProtections(SpreadsheetApp.ProtectionType.RANGE)
    .filter((protection) => protection.getDescription() === PROTECTION_DESCRIPTION);
}

function removeOwnProtections(sheet) {
  const own = findOwnHeaderProtections(sheet);
  own.forEach((protection) => protection.remove());
  return own.length;
}
