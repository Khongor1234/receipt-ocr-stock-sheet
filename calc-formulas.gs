/**
 * Replaces the per-row formulas in AL ("эцсийн үлдэгдэл"), AM ("үнэлгээ") and
 * AN ("борлуулалтын тоо") with ONE array formula per column, written in the
 * protected header row 3 and filling every row below it.
 *
 * Why: per-row formulas are missing on rows that were added later (139 rows
 * without AL, 1039 without AN in the 2026 export) and cannot be protected without
 * making product rows undeletable. An array formula in the protected header:
 *   - covers new rows automatically (insert a row, the result appears)
 *   - cannot be deleted (header row 3 is protected, columns cannot be deleted)
 *   - leaves product rows free to add and delete
 * A value typed into one of the filled cells blocks the array of that column (its
 * header shows #REF!). repairCalcColumns removes such a value and the column fills
 * again: at once through instant-sync.gs, and every hour by the background job.
 *
 * Same rules as the old formulas:
 *   AL = sum(E:T) - sum(U:AJ) + B        AN = sum(U:AI)
 * (with the gift column, config.gs GIFT_COLUMN, the gifts are part of the expense sum, not of "sold")
 *   AM = "зөв" when AK (count) equals AL, otherwise "буруу"
 * A row is a product row when at least one of B, C, D holds a number; divider and
 * empty rows stay blank.
 *
 * The two shops' layout (LAYOUT.paired, config.gs) has four columns instead: the final
 * balance AM, the check AN, and in place of "sold" the OTHER shop's count AO and final balance
 * AP (the owner, 2026-10-28: the count first; with the gift column AP and AQ), read from its sheet
 * of the same half-month (My/2026/10/15 <-> Sor/2026/10/15) by product
 * name (case and spaces do not matter). A product the other shop does not have, or did not
 * count, stays blank. So a product one over in one shop and one short in the other shows at a
 * glance (a transfer that was not written). Like the other formula columns they are never
 * typed in: a typed value is taken out again (repairCalcColumns).
 *
 * Hand-typed AL values (no formula) are replaced by the computed value; the log
 * counts them per sheet so a stale number is never replaced silently.
 *
 * Usage: setupYear installs the formulas (step 5). By hand: installCalcFormulas,
 * with DRY_RUN = true (config.gs) first. Needs year-sync.gs (isEntrySheet).
 */

const CALC_HEADER_ROW = 3;
const CALC_FIRST_ROW = 4;
const CALC_FIRST_COL = LAYOUT.balance; // AL (AM = 39, AN = 40); the shops' layout: AM
const CALC_COLUMN_COUNT = LAYOUT.paired ? 4 : 3; // the shops' layout: AM:AP
const CALC_DEFAULT_HEADERS = ['эцсийн үлдэгдэл ', 'үнэлгээ ', 'борлуулалтын тоо'];
// The other shop's two columns of the shops' layout: their headers name that shop; the count first.
const calcPairedHeaders = (otherLabel) => [`${otherLabel} тооллого`, `${otherLabel} эцсийн үлдэгдэл`];
const CALC_MATCH_TEXT = 'зөв'; // AM: the count equals the computed balance
const CALC_MISMATCH_TEXT = 'буруу';
const CALC_HEADER_IN_FORMULA = /^=ARRAYFORMULA\(\{"((?:[^"]|"")*)";/;

function installCalcFormulas() {
  const lines = runInstallCalcFormulas(currentBook());
  Logger.log('%s\n%s', modeLabel(), lines.join('\n'));
}

// The other shop's two columns the other way round on every shop sheet (the owner, 2026-10-28: the
// count AP, then the balance AQ): the four formulas are written anew where they differ, and the
// headers remembered again at once, so the hourly restore (header-guard.gs) keeps the new ones.
// Run once from the editor; DRY_RUN (config.gs) reports only.
function swapOtherShopColumns() {
  const book = currentBook();
  const lines = [...runInstallCalcFormulas(book), ...runRememberHeaders(book)];
  Logger.log('%s\n%s', modeLabel(), lines.join('\n'));
  return lines;
}

function listCalcSheets(spreadsheet) {
  return spreadsheet
    .getSheets()
    .filter((sheet) => !sheet.getName().startsWith('_') && !isStockSheet(sheet))
    .filter(isEntrySheet);
}

function runInstallCalcFormulas(spreadsheet) {
  return listCalcSheets(spreadsheet).map(installOnSheet);
}

// ---------- pure helpers (no Spreadsheet calls) ----------

function quoteText(text) {
  return `"${String(text).replace(/"/g, '""')}"`;
}

// 'Sor/2026/10/15' as a formula writes a tab name.
const quoteSheetName = (name) => `'${String(name).replace(/'/g, "''")}'`;

// The other shop's value of each product row: its column `column` on tab otherTab, found by
// the name in A with case and spaces left out; blank when the name is not there or the cell is.
function otherShopLookup(otherTab, column) {
  const tab = quoteSheetName(otherTab);
  const key = (range) => `REGEXREPLACE(LOWER(${range}),"\\s","")`;
  const letter = columnLetter(column);
  const values = `IF(LEN(${tab}!${letter}4:${letter}),${tab}!${letter}4:${letter},"")`;
  return `IFERROR(VLOOKUP(${key('A4:A')},{${key(`${tab}!A4:A`)},${values}},2,FALSE),"")`;
}

// headers: one per column. otherTab: the other shop's tab of the same half-month (the shops'
// layout only).
function buildCalcFormulas(headers, otherTab) {
  const isProduct = 'BYROW(B4:D,LAMBDA(r,COUNT(r)>0))';
  const open = (first, last) => `BYROW(${columnLetter(first)}4:${columnLetter(last)},LAMBDA(r,SUM(r)))`;
  const income = open(LAYOUT.incomeFirst, LAYOUT.transferIn); // E4:T
  const expense = open(LAYOUT.expenseFirst, LAYOUT.expenseLast); // U4:AJ (with the gift column: to AL)
  const sold = open(LAYOUT.expenseFirst, LAYOUT.transferOut - 1); // U4:AI
  const countColumn = columnLetter(LAYOUT.count);
  const balanceColumn = columnLetter(LAYOUT.balance);
  const opening = 'BYROW(B4:B,LAMBDA(r,N(r)))';
  const shared = [
    `=ARRAYFORMULA({${quoteText(headers[0])};LET(isp,${isProduct},inc,${income},outc,${expense},opn,${opening},IF(isp,inc-outc+opn,""))})`,
    `=ARRAYFORMULA({${quoteText(headers[1])};IF(${isProduct},IF(${countColumn}4:${countColumn}=${balanceColumn}4:${balanceColumn},${quoteText(CALC_MATCH_TEXT)},${quoteText(CALC_MISMATCH_TEXT)}),"")})`,
  ];
  if (!LAYOUT.paired) {
    return [...shared, `=ARRAYFORMULA({${quoteText(headers[2])};IF(${isProduct},${sold},"")})`];
  }
  const other = (index, column) => `=ARRAYFORMULA({${quoteText(headers[index])};IF(${isProduct},${otherShopLookup(otherTab, column)},"")})`;
  return [...shared, other(2, LAYOUT.count), other(3, LAYOUT.balance)];
}

// The other shop of a year sheet and its tab of the same period: { label, tab }, or null for a
// sheet outside the chains or a workbook with one shop.
function otherShopOf(sheetName) {
  const slot = sheetSlot(sheetName);
  if (!slot || CHAINS.length !== 2) {
    return null;
  }
  const other = CHAINS.find((chain) => chain.key !== slot.chainKey);
  return { label: other.label, tab: chainSheetName(other, slot.position) };
}

// The header text lives inside the array formula once it is installed. Reading it
// from there keeps the name even while the cell shows an error such as #REF!.
function pickHeader(formula, value, fallback) {
  const installed = String(formula).match(CALC_HEADER_IN_FORMULA);
  if (installed) {
    return installed[1].replace(/""/g, '"');
  }
  const text = String(value);
  return formula !== '' || text.trim() === '' || text.startsWith('#') ? fallback : text;
}

function countHandTyped(values, formulas) {
  return values.filter((value, index) => value !== '' && formulas[index] === '').length;
}

// ---------- spreadsheet access ----------

function calcBodyRange(sheet) {
  const rows = sheet.getMaxRows() - CALC_FIRST_ROW + 1;
  return rows > 0 ? sheet.getRange(CALC_FIRST_ROW, CALC_FIRST_COL, rows, CALC_COLUMN_COUNT) : null;
}

// "AL3:AN3" (the shops' layout: "AM3:AP3"), for the log.
const calcHeaderCells = () => `${columnLetter(CALC_FIRST_COL)}3:${columnLetter(CALC_FIRST_COL + CALC_COLUMN_COUNT - 1)}3`;

function installOnSheet(sheet) {
  const name = sheet.getName();
  const other = LAYOUT.paired ? otherShopOf(name) : null;
  if (LAYOUT.paired && !other) {
    return `${name}: нөгөө дэлгүүрийн хуудас танигдсангүй, томьёо тавиагүй`;
  }
  const headerRange = sheet.getRange(CALC_HEADER_ROW, CALC_FIRST_COL, 1, CALC_COLUMN_COUNT);
  const currentFormulas = headerRange.getFormulas()[0];
  // The other shop's two headers always name that shop (the column held "sold" before).
  const defaults = other ? [...CALC_DEFAULT_HEADERS.slice(0, 2), ...calcPairedHeaders(other.label)] : CALC_DEFAULT_HEADERS;
  const headers = headerRange
    .getValues()[0]
    .map((value, index) => (index >= 2 && other ? defaults[index] : pickHeader(currentFormulas[index], value, defaults[index])));
  const wanted = buildCalcFormulas(headers, other && other.tab);
  const body = calcBodyRange(sheet);
  if (currentFormulas.every((formula, index) => formula === wanted[index])) {
    return repairCalcColumns(sheet) || `${name}: томьёо аль хэдийн тавигдсан`;
  }
  const alColumn = sheet.getRange(CALC_FIRST_ROW, CALC_FIRST_COL, body.getNumRows(), 1);
  const handTyped = countHandTyped(
    alColumn.getValues().map((cells) => cells[0]),
    alColumn.getFormulas().map((cells) => cells[0])
  );
  if (!DRY_RUN) {
    body.clearContent();
    headerRange.setFormulas([wanted]);
  }
  return `${name}: ${calcHeaderCells()}-д томьёо ${DRY_RUN ? 'тавина' : 'тавилаа'} (гараар бичсэн ${handTyped} ${columnLetter(CALC_FIRST_COL)} утгыг томьёо орлоно)`;
}

// A value typed into AL:AN blocks the array formula of that column: its header
// cell shows an error and the whole column goes empty. While a column is blocked
// nothing is filled in, so everything found below the header was typed by hand;
// it is removed and the formula fills the column again. Returns a log line, or
// null when the columns are healthy.
function repairCalcColumns(sheet) {
  const body = calcBodyRange(sheet);
  if (body === null) {
    return null;
  }
  const headers = sheet.getRange(CALC_HEADER_ROW, CALC_FIRST_COL, 1, CALC_COLUMN_COUNT).getDisplayValues()[0];
  const blocked = headers
    .map((text, index) => ({ index, isBlocked: String(text).startsWith('#') }))
    .filter((column) => column.isBlocked);
  if (blocked.length === 0) {
    return null;
  }
  if (!DRY_RUN) {
    blocked.forEach((column) =>
      sheet.getRange(CALC_FIRST_ROW, CALC_FIRST_COL + column.index, body.getNumRows(), 1).clearContent()
    );
  }
  return `${sheet.getName()}: ${columnLetter(CALC_FIRST_COL)}:${columnLetter(CALC_FIRST_COL + CALC_COLUMN_COUNT - 1)}-д гараар бичсэн утгыг ${DRY_RUN ? 'арилгана' : 'арилгаж томьёог сэргээлээ'} (${blocked.length} багана)`;
}

function runRepairCalcColumns(spreadsheet) {
  return findChains(spreadsheet)
    .flatMap(({ sheets }) => sheets.map(repairCalcColumns))
    .filter((line) => line !== null);
}
