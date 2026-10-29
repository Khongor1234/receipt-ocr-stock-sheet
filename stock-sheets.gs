/**
 * Тооллогын хуудас: Sheet100 (Соронзон) and Sheet100* (My home), profile.gs STOCK_SHEETS.
 *
 * Each one lists every product of its shop and how many are in the shop TODAY: on
 * 10 October it shows the 10 October stock. It is a formula, so it follows every
 * entry at once and moves to the next half-month sheet by itself:
 *   today's sheet   the shop's sheet of today's half-month (Sor/2026/10/15, ...)
 *   stock           opening balance B + income of the days up to today + transfer in
 *                   - expense of the days up to today - transfer out; where the
 *                   sheet before was not counted at all, its final balance (AL)
 *                   stands in for the empty B
 *   categories      a row with no opening balance, cost or price (like the AL:AN
 *                   formulas) shows its name only, bold and coloured
 * The transfer columns (T, AJ) have no day, so they count as soon as they are typed.
 * TODAY() follows the time zone of the file (File > Settings).
 *
 * Column C "Тооллого" stays empty for a count on paper or typed by hand. It is not
 * tied to the list: a typed number stays in its cell when the list moves.
 *
 * The sheet is protected except column C: only the file owner can change the
 * list, Mom's account cannot type in A:B, delete rows or delete the sheet. The
 * other scripts (colours, dividers, row trimming, header protection) leave these
 * sheets and the count sheets (count-sheets.gs) alone.
 *
 * installStockSheets: makes both sheets and both count sheets (setupYear does it
 * as well). It clears what is on them first. Running it again is safe.
 * Needs config.gs and profile.gs.
 */

const STOCK_HEADER_ROW = 2;
const STOCK_FIRST_ROW = 3;
const STOCK_COLUMN_COUNT = 3; // A name, B stock, C count by hand
const STOCK_HAND_COUNT_COLUMN = 3;
const STOCK_MIN_ROWS = 500; // more than the longest shop sheet, so the list always fits
const STOCK_HEADERS = ['Барааны нэр', 'Үлдэгдэл', 'Тооллого'];
const STOCK_PROTECTION_DESCRIPTION = 'Тооллогын хуудас - зөвхөн харах (C-д тооллого бичиж болно)';
const STOCK_INPUT_COLOR = '#ffffff'; // white: a number written on a black-and-white print stays readable
const STOCK_CATEGORY_COLOR = '#fce5cd';
const STOCK_HEADER_COLOR = '#cfe2f3';
const STOCK_ZERO_COLOR = '#999999';
const STOCK_NEGATIVE_COLOR = '#cc0000';
const STOCK_NAME_WIDTH = 280;
const STOCK_COUNT_WIDTH = 110;

function installStockSheets() {
  const lines = runInstallStockSheets(currentBook());
  Logger.log('%s\n%s', modeLabel(), lines.join('\n'));
}

// The stock sheets and the count sheets: the other scripts leave both alone.
function isStockSheetName(name) {
  return STOCK_SHEETS.some((stock) => stock.name === String(name)) || isCountSheetName(name);
}

function isStockSheet(sheet) {
  return isStockSheetName(sheet.getName());
}

// ---------- pure helpers (no Spreadsheet calls) ----------

// Formula text of today's tab name of a shop, e.g. "Sor/2026/10/31". t = today.
function todayTabExpression(chain) {
  const start = `"${chain.tabPrefix}/"&YEAR(t)&"/"&MONTH(t)`;
  if (PERIODS_PER_MONTH !== 2) {
    return start;
  }
  return `${start}&"/"&IF(DAY(t)<=${FIRST_HALF_LAST_DAY},${FIRST_HALF_LAST_DAY},DAY(EOMONTH(t,0)))`;
}

// Formula text of the tab name of the period before today's, e.g. "Sor/2026/10/15".
function previousTabExpression(chain) {
  const monthEnd = 'EOMONTH(t,-1)';
  const lastMonth = `"${chain.tabPrefix}/"&YEAR(${monthEnd})&"/"&MONTH(${monthEnd})`;
  if (PERIODS_PER_MONTH !== 2) {
    return lastMonth;
  }
  const firstHalf = `"${chain.tabPrefix}/"&YEAR(t)&"/"&MONTH(t)&"/${FIRST_HALF_LAST_DAY}"`;
  return `IF(DAY(t)<=${FIRST_HALF_LAST_DAY},${lastMonth}&"/"&DAY(${monthEnd}),${firstHalf})`;
}

// How many day columns of today's sheet are already past (today included). Day 31
// is typed in the last day column, so the count never goes beyond DAY_COLUMNS.
function todayDayCountExpression() {
  if (PERIODS_PER_MONTH !== 2) {
    return `MIN(${DAY_COLUMNS},DAY(t))`;
  }
  return `IF(DAY(t)<=${FIRST_HALF_LAST_DAY},DAY(t),MIN(${DAY_COLUMNS},DAY(t)-${FIRST_HALF_LAST_DAY}))`;
}

function stockColumns(first, last) {
  return `"${columnLetter(first)}4:${columnLetter(last)}"`;
}

// The three formulas of a stock sheet: title (A1), today's sheet name (B1), list (A3).
// Opening balance: B of today's sheet. Where the sheet before was not counted at all
// (its count column is empty), B stays empty, so that sheet's final balance (AL) of
// the same product is used instead. Where it was counted, an empty B means none.
function buildStockFormulas(chain) {
  const today = `t,TODAY(),tab,${todayTabExpression(chain)}`;
  const number = (name) => `IF(ISNUMBER(${name}),${name},0)`;
  const read = (name, columns) => `${name},INDIRECT(s&${columns})`;
  const previous = `p,"'"&${previousTabExpression(chain)}&"'!"`;
  const counted = `counted,IFERROR(COUNT(INDIRECT(p&${stockColumns(LAYOUT.count, LAYOUT.count)}))>0,TRUE)`;
  const carried = `carried,IFERROR(VLOOKUP(nm,INDIRECT(p&${stockColumns(1, LAYOUT.balance)}),${LAYOUT.balance},FALSE),0)`;
  const list = [
    `=ARRAYFORMULA(LET(${today},s,"'"&tab&"'!",k,${todayDayCountExpression()}`,
    read('nm', stockColumns(1, 1)),
    read('opn', stockColumns(2, 2)),
    read('prices', stockColumns(2, 4)),
    read('inc', stockColumns(LAYOUT.incomeFirst, LAYOUT.transferIn - 1)),
    read('tin', stockColumns(LAYOUT.transferIn, LAYOUT.transferIn)),
    read('outc', stockColumns(LAYOUT.expenseFirst, LAYOUT.transferOut - 1)),
    read('tout', stockColumns(LAYOUT.transferOut, LAYOUT.expenseLast)), // the transfer out and the gifts
    previous,
    counted,
    carried,
    `start,IF(ISNUMBER(opn),opn,IF(counted,0,${number('carried')}))`,
    `upto,--(SEQUENCE(${DAY_COLUMNS})<=k)`,
    `isp,BYROW(prices,LAMBDA(r,COUNT(r)>0))`,
    `bal,start+MMULT(${number('inc')},upto)+${number('tin')}-MMULT(${number('outc')},upto)-${LAYOUT.gift ? `BYROW(${number('tout')},LAMBDA(r,SUM(r)))` : number('tout')}`,
    `IFERROR(FILTER(HSTACK(nm,IF(isp,bal,"")),LEN(TRIM(nm))>0),"Өнөөдрийн sheet олдсонгүй: "&tab)))`,
  ].join(',');
  return {
    title: `=LET(t,TODAY(),"${chain.label}: "&TEXT(t,"yyyy.mm.dd")&" өдрийн байдлаарх үлдэгдэл")`,
    source: `=LET(${today},"эх sheet: "&tab)`,
    list,
  };
}

// ---------- spreadsheet access ----------

function runInstallStockSheets(spreadsheet) {
  if (STOCK_SHEETS.length === 0 && COUNT_SHEETS.length === 0) {
    return ['тооллогын хуудас энэ файлд тохируулаагүй (STOCK_SHEETS, COUNT_SHEETS хоосон)'];
  }
  return [
    ...STOCK_SHEETS.map((stock) => installStockSheet(spreadsheet, stock)),
    ...COUNT_SHEETS.map((count) => installCountSheet(spreadsheet, count)),
    ...refreshCountSheets(spreadsheet),
  ];
}

function installStockSheet(spreadsheet, stock) {
  const chain = CHAINS.find((candidate) => candidate.key === stock.chainKey);
  if (!chain) {
    return `${stock.name}: "${stock.chainKey}" дэлгүүр profile.gs CHAINS-д алга - алгаслаа`;
  }
  const existing = spreadsheet.getSheetByName(stock.name);
  const formulas = buildStockFormulas(chain);
  if (existing !== null && isStockSheetInstalled(existing, formulas)) {
    return `${stock.name}: ${chain.label}-ийн тооллогын хуудас аль хэдийн бэлэн`;
  }
  const action = existing === null ? 'шинээр үүсгэж' : 'одоо байгааг цэвэрлээд';
  if (DRY_RUN) {
    return `${stock.name}: ${action} ${chain.label}-ийн өнөөдрийн үлдэгдлийн хуудас болгоно, бүхэлд нь түгжинэ`;
  }
  const sheet = existing || spreadsheet.insertSheet(stock.name);
  writeStockSheet(sheet, formulas);
  return `${stock.name}: ${action} ${chain.label}-ийн өнөөдрийн үлдэгдлийн хуудас болголоо, бүхэлд нь түгжлээ`;
}

function isStockSheetInstalled(sheet, formulas) {
  const isProtected = sheet
    .getProtections(SpreadsheetApp.ProtectionType.SHEET)
    .some((protection) => protection.getDescription() === STOCK_PROTECTION_DESCRIPTION);
  return (
    isProtected &&
    sheet.getRange(1, 1).getFormula() === formulas.title &&
    sheet.getRange(1, 2).getFormula() === formulas.source &&
    sheet.getRange(STOCK_HEADER_ROW, STOCK_HAND_COUNT_COLUMN).getValue() === STOCK_HEADERS[2] &&
    sheet.getRange(STOCK_FIRST_ROW, 1).getFormula() === formulas.list
  );
}

function writeStockSheet(sheet, formulas) {
  clearStockSheet(sheet);
  fitStockGrid(sheet, STOCK_COLUMN_COUNT);
  sheet.getRange(1, 1, 1, 2).setFormulas([[formulas.title, formulas.source]]);
  sheet.getRange(STOCK_HEADER_ROW, 1, 1, STOCK_COLUMN_COUNT).setValues([STOCK_HEADERS]);
  sheet.getRange(STOCK_FIRST_ROW, 1).setFormula(formulas.list);
  styleStockSheet(sheet);
  protectStockSheet(sheet);
}

// Everything earlier scripts or hands put on the sheet goes: values, colours,
// protections, filters, rules.
function clearStockSheet(sheet) {
  sheet
    .getProtections(SpreadsheetApp.ProtectionType.RANGE)
    .concat(sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET))
    .filter((protection) => protection.canEdit())
    .forEach((protection) => protection.remove());
  const filter = sheet.getFilter();
  if (filter !== null) {
    filter.remove();
  }
  const all = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
  all.breakApart();
  all.clearDataValidations();
  sheet.clearConditionalFormatRules();
  sheet.clear();
}

function fitStockGrid(sheet, columns) {
  const extraColumns = sheet.getMaxColumns() - columns;
  if (extraColumns > 0) {
    sheet.deleteColumns(columns + 1, extraColumns);
  }
  if (extraColumns < 0) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), -extraColumns);
  }
  const missingRows = STOCK_MIN_ROWS - sheet.getMaxRows();
  if (missingRows > 0) {
    sheet.insertRowsAfter(sheet.getMaxRows(), missingRows);
  }
}

function styleStockSheet(sheet) {
  sheet.setFrozenRows(STOCK_HEADER_ROW);
  sheet.setColumnWidth(1, STOCK_NAME_WIDTH);
  sheet.setColumnWidth(2, STOCK_COUNT_WIDTH);
  sheet.setColumnWidth(STOCK_HAND_COUNT_COLUMN, STOCK_COUNT_WIDTH);
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
  sheet.getRange(1, 1).setFontSize(12);
  sheet
    .getRange(STOCK_HEADER_ROW, 1, 1, STOCK_COLUMN_COUNT)
    .setFontWeight('bold')
    .setBackground(STOCK_HEADER_COLOR)
    .setHorizontalAlignment('center');
  const body = sheet.getRange(STOCK_FIRST_ROW, 1, sheet.getMaxRows() - STOCK_FIRST_ROW + 1, STOCK_COLUMN_COUNT);
  sheet.getRange(STOCK_FIRST_ROW, 2, body.getNumRows(), 1).setNumberFormat('0').setHorizontalAlignment('center');
  stockHandCountRange(sheet).setBackground(STOCK_INPUT_COLOR).setHorizontalAlignment('center');
  const first = `$A${STOCK_FIRST_ROW}`;
  const count = `$B${STOCK_FIRST_ROW}`;
  const rule = (formula) => SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(formula).setRanges([body]);
  sheet.setConditionalFormatRules([
    rule(`=AND(LEN(TRIM(${first}))>0,${count}="")`).setBold(true).setBackground(STOCK_CATEGORY_COLOR).build(),
    rule(`=AND(ISNUMBER(${count}),${count}<0)`).setFontColor(STOCK_NEGATIVE_COLOR).setBold(true).build(),
    rule(`=AND(ISNUMBER(${count}),${count}=0)`).setFontColor(STOCK_ZERO_COLOR).build(),
  ]);
}

function protectStockSheet(sheet) {
  const protection = sheet.protect().setDescription(STOCK_PROTECTION_DESCRIPTION);
  protection.setWarningOnly(false);
  restrictToOwner(protection);
  protection.setUnprotectedRanges([stockHandCountRange(sheet)]);
}

function stockHandCountRange(sheet) {
  return sheet.getRange(STOCK_FIRST_ROW, STOCK_HAND_COUNT_COLUMN, sheet.getMaxRows() - STOCK_FIRST_ROW + 1, 1);
}
