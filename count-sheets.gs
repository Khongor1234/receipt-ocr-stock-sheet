/**
 * Тооллого хийх хуудас: "Соронзон тооллого" and "My home тооллого" (profile.gs COUNT_SHEETS).
 *
 * For counting on the 15th and the last day, on paper or in the sheet: column A is
 * the product list of the shop's sheet of today (categories coloured), column B
 * "Тооллого" is empty for the count. No expected numbers are shown.
 *   - The names are written by this script and locked: Mom cannot change them.
 *     Column B is the only place she can type.
 *   - A count belongs to its product, not to its row: when a product is added or
 *     removed the list is written again and every typed count goes with its name.
 *   - Counts are never cleared by the script. Before a new count, select B and
 *     press Delete.
 * The list follows a product change within seconds (instant-sync.gs, after the
 * product sync) and at the latest with the hourly job; on the 16th and the 1st it
 * moves to the new half-month sheet the same way.
 *
 * installStockSheets (stock-sheets.gs) makes these sheets too.
 * Needs config.gs, profile.gs, stock-sheets.gs, year-sync.gs (positionOfDate).
 */

const COUNT_HEADER_ROW = 2;
const COUNT_FIRST_ROW = 3;
const COUNT_SHOP_FIRST_ROW = 4; // first product row of a shop sheet
const COUNT_SHOP_COLUMNS = 4; // A name, B opening, C cost, D price
const COUNT_COLUMN_COUNT = 2; // A name, B count
const COUNT_HEADERS = ['Барааны нэр', 'Тооллого'];
const COUNT_PROTECTION_DESCRIPTION = 'Тооллого хийх хуудас - нэр түгжээтэй, зөвхөн B бичнэ';
const COUNT_CATEGORY_COLOR = '#fce5cd';
const COUNT_PLAIN_COLOR = '#ffffff';
const COUNT_INPUT_COLOR = '#ffffff'; // white: a number written on a black-and-white print stays readable

function isCountSheetName(name) {
  return COUNT_SHEETS.some((count) => count.name === String(name));
}

// ---------- pure helpers (no Spreadsheet calls) ----------

function countKey(name) {
  return String(name).trim().toLowerCase();
}

// Rows of a shop sheet (A..D from row 4) -> [{ name, isCategory }]. A product has a
// number in B, C or D, like the AL:AN formulas; a named row without one is a category.
function readCountList(rows) {
  return rows
    .filter((row) => String(row[0]).trim() !== '')
    .map((row) => ({
      name: row[0],
      isCategory: !row.slice(1, COUNT_SHOP_COLUMNS).some((value) => typeof value === 'number'),
    }));
}

// The table to write: names with the count typed earlier for the same product.
function buildCountTable(list, oldRows) {
  const typed = new Map(
    oldRows
      .filter((row) => String(row[0]).trim() !== '' && row[1] !== '')
      .map((row) => [countKey(row[0]), row[1]])
  );
  return list.map((item) => [item.name, item.isCategory ? '' : typed.get(countKey(item.name)) ?? '']);
}

function isCountSheetTouched(count, position, scopes) {
  if (scopes === undefined) {
    return true;
  }
  return scopes.some(
    (scope) => scope === undefined || (scope.chainKey === count.chainKey && scope.fromPosition <= position)
  );
}

function countTitle(chain, tabName) {
  return `${chain.label} тооллого - "${tabName}" хуудасны бараа (B баганад тоолсон тоо)`;
}

// true when the sheet already shows exactly this table and nothing below it.
function isTableShown(table, oldRows) {
  const isSameRow = (row, index) => row[0] === oldRows[index][0] && row[1] === oldRows[index][1];
  return (
    oldRows.length >= table.length &&
    table.every(isSameRow) &&
    oldRows.slice(table.length).every((row) => row[0] === '' && row[1] === '')
  );
}

// ---------- spreadsheet access ----------

// Writes the count sheets' lists again where they changed. Returns log lines.
// scopes (optional): the changes just synced, as instant-sync.gs lists them
// ({ chainKey, fromPosition }, undefined = every sheet). Only a count sheet whose
// shop changed on today's sheet or before is read; without scopes, all are.
function refreshCountSheets(spreadsheet, scopes) {
  const position = positionOfDate(new Date());
  return COUNT_SHEETS.filter((count) => isCountSheetTouched(count, position, scopes))
    .map((count) => refreshCountSheet(spreadsheet, count, position))
    .filter((line) => line !== null);
}

function refreshCountSheet(spreadsheet, count, position) {
  const sheet = spreadsheet.getSheetByName(count.name);
  const chain = CHAINS.find((candidate) => candidate.key === count.chainKey);
  if (sheet === null || !chain) {
    return null; // not made yet: installStockSheets makes it
  }
  const tabName = chainSheetName(chain, position);
  const source = spreadsheet.getSheetByName(tabName);
  if (source === null) {
    return `${count.name}: өнөөдрийн "${tabName}" sheet олдсонгүй`;
  }
  const list = readCountList(readShopRows(source));
  const oldRows = readCountRows(sheet);
  const table = buildCountTable(list, oldRows);
  const title = countTitle(chain, tabName);
  if (sheet.getRange(1, 1).getValue() === title && isTableShown(table, oldRows)) {
    return null;
  }
  if (DRY_RUN) {
    return `${count.name}: ${list.length} мөрийг "${tabName}"-ээс дахин бичнэ (тоолсон тоо бараатайгаа үлдэнэ)`;
  }
  writeCountTable(sheet, title, list, table, oldRows.length);
  return `${count.name}: ${list.length} мөрийг "${tabName}"-ээс дахин бичлээ`;
}

function readShopRows(source) {
  const height = findLastDataRow(source) - COUNT_SHOP_FIRST_ROW + 1;
  return height > 0 ? source.getRange(COUNT_SHOP_FIRST_ROW, 1, height, COUNT_SHOP_COLUMNS).getValues() : [];
}

function readCountRows(sheet) {
  const height = sheet.getMaxRows() - COUNT_FIRST_ROW + 1;
  return height > 0 ? sheet.getRange(COUNT_FIRST_ROW, 1, height, COUNT_COLUMN_COUNT).getValues() : [];
}

function writeCountTable(sheet, title, list, table, oldHeight) {
  const missingRows = list.length + COUNT_FIRST_ROW - 1 - sheet.getMaxRows();
  if (missingRows > 0) {
    sheet.insertRowsAfter(sheet.getMaxRows(), missingRows);
    unprotectCountColumn(sheet);
  }
  sheet.getRange(1, 1).setValue(title);
  if (oldHeight > 0) {
    sheet.getRange(COUNT_FIRST_ROW, 1, oldHeight, COUNT_COLUMN_COUNT).clearContent().setFontWeight('normal');
    sheet.getRange(COUNT_FIRST_ROW, 1, oldHeight, 1).setBackground(COUNT_PLAIN_COLOR);
    sheet.getRange(COUNT_FIRST_ROW, 2, oldHeight, 1).setBackground(COUNT_INPUT_COLOR);
  }
  if (table.length === 0) {
    return;
  }
  sheet.getRange(COUNT_FIRST_ROW, 1, table.length, COUNT_COLUMN_COUNT).setValues(table);
  list.forEach((item, index) => {
    if (item.isCategory) {
      sheet
        .getRange(COUNT_FIRST_ROW + index, 1, 1, COUNT_COLUMN_COUNT)
        .setBackground(COUNT_CATEGORY_COLOR)
        .setFontWeight('bold');
    }
  });
}

// Makes an empty count sheet (or empties an old one): headers, widths, lock.
function prepareCountSheet(sheet) {
  clearStockSheet(sheet);
  fitStockGrid(sheet, COUNT_COLUMN_COUNT);
  sheet.getRange(COUNT_HEADER_ROW, 1, 1, COUNT_COLUMN_COUNT).setValues([COUNT_HEADERS]);
  sheet.setFrozenRows(COUNT_HEADER_ROW);
  sheet.setColumnWidth(1, STOCK_NAME_WIDTH);
  sheet.setColumnWidth(2, STOCK_COUNT_WIDTH);
  sheet.getRange(1, 1).setFontWeight('bold').setFontSize(12);
  sheet
    .getRange(COUNT_HEADER_ROW, 1, 1, COUNT_COLUMN_COUNT)
    .setFontWeight('bold')
    .setBackground(STOCK_HEADER_COLOR)
    .setHorizontalAlignment('center');
  sheet
    .getRange(COUNT_FIRST_ROW, 2, sheet.getMaxRows() - COUNT_FIRST_ROW + 1, 1)
    .setBackground(COUNT_INPUT_COLOR)
    .setHorizontalAlignment('center');
  const protection = sheet.protect().setDescription(COUNT_PROTECTION_DESCRIPTION);
  protection.setWarningOnly(false);
  restrictToOwner(protection);
  unprotectCountColumn(sheet);
}

function unprotectCountColumn(sheet) {
  const open = sheet.getRange(COUNT_FIRST_ROW, 2, sheet.getMaxRows() - COUNT_FIRST_ROW + 1, 1);
  findCountProtection(sheet).forEach((protection) => protection.setUnprotectedRanges([open]));
}

function findCountProtection(sheet) {
  return sheet
    .getProtections(SpreadsheetApp.ProtectionType.SHEET)
    .filter((protection) => protection.getDescription() === COUNT_PROTECTION_DESCRIPTION);
}

function installCountSheet(spreadsheet, count) {
  const chain = CHAINS.find((candidate) => candidate.key === count.chainKey);
  if (!chain) {
    return `${count.name}: "${count.chainKey}" дэлгүүр profile.gs CHAINS-д алга - алгаслаа`;
  }
  const existing = spreadsheet.getSheetByName(count.name);
  if (existing !== null && findCountProtection(existing).length > 0) {
    return `${count.name}: аль хэдийн бэлэн`;
  }
  const action = existing === null ? 'шинээр үүсгэж' : 'цэвэрлээд';
  if (DRY_RUN) {
    return `${count.name}: ${action} ${chain.label}-ийн тооллого хийх хуудас болгоно (нэр түгжээтэй, B-д бичнэ)`;
  }
  prepareCountSheet(existing || spreadsheet.insertSheet(count.name));
  return `${count.name}: ${action} ${chain.label}-ийн тооллого хийх хуудас болголоо`;
}
