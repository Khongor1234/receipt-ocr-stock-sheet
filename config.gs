/**
 * Shared helpers for every script in this project: the tab names, the period a tab
 * stands for, and the column layout of a sheet. The values that differ from one
 * workbook to the other (DRY_RUN, the shops, how long a period is) are in
 * profile.gs, which is loaded BEFORE this file.
 *
 * A sheet is recognised by its tab name: shop / year / month, and for a workbook
 * with two sheets a month also the last day of the period.
 *   half-month sheets   My/2026/1/15, My/2026/1/31, ... Sor/2026/12/31
 *   month sheets        Tos/2026/7, Tos/2026/8, ...
 * The old names (Sheet1..Sheet24 = My home, Sheet31..Sheet54 = Соронзон, also the
 * misspelt "Sheet 9", "Seet 41") are still understood where a shop has them;
 * setupYear renames them. Any other sheet (Sheet100, Sheet200, ...) is never part
 * of a chain.
 *
 * The columns of a sheet, for DAY_COLUMNS = 15 (half-month) and 30 (month):
 *   A name, B opening balance, C cost, D price
 *   income days from E, then the transfer-in column (T / AI)
 *   expense days, then the transfer-out column (AJ / BN)
 *   count, final balance, check, sold (AK AL AM AN / BO BP BQ BR)
 *   the five totals (AO..AS / BS..BW)
 * Тосгон (SHELF_COLUMN in its profile, 2026-10-22) has the shelf price in E too, so its days
 * start at F; it keeps its "sold" column.
 * The two shops' workbook (PAIRED_SHOPS in profile.gs) has two more columns:
 *   A..D as above, E the shelf price (typed by hand: what the shop really asks)
 *   income days from F, transfer in U, expense days from V, transfer out AK
 *   count AL, final balance AM, check AN, then in place of "sold" the other shop's final
 *   balance AO and count AP (array formulas, calc-formulas.gs), the five totals AQ..AU
 * and, with GIFT_COLUMN (profile.gs), the goods given away as gifts in a column of their own
 * right after the transfer out (AL "Бэлэг"): not sold, so the sales leave them out, but they
 * leave the shop (the final balance takes them off) and cost money (the sales profit takes them
 * off at cost). Everything from the count on is one column further right then:
 *   count AM, final balance AN, check AO, the other shop AP AQ, the totals AR..AV
 */

const MONTHS_PER_YEAR = 12;
let CHAIN_PERIOD_COUNT = CHAIN_MONTHS * PERIODS_PER_MONTH; // set again by useBook (year-books.gs)
const SHEET_NAME_PATTERN = /^s[het]{2,4}\s*(\d+)\s*$/i; // old: "Sheet19", "Sheet 9", "Seet 41"
const TAB_NAME_PATTERN = /^([a-z]+)\/(\d{4})\/(\d{1,2})(?:\/(\d{1,2}))?$/i; // "My/2026/10/15", "Tos/2026/7"
const FIRST_HALF_LAST_DAY = 15;
const FIRST_DAY_COLUMN = 5; // E (F in the two shops' workbook, after the shelf price)
const SHELF_PRICE_COLUMN = 5; // E, the two shops' workbook only
const TOTAL_COUNT = 5; // the five "Нийт ..." totals
const SHELF_PRICE_HEADER = 'яг харагдах үнэ'; // row 3 of the shelf price column: marks the new layout
const GIFT_HEADER = 'Бэлэг'; // row 3 of the gift column (GIFT_COLUMN): marks a sheet that has it
const LAYOUT_HEADER_ROW = 3;

function modeLabel() {
  return DRY_RUN ? 'ТУРШИЛТ - юу ч өөрчлөгдөөгүй' : 'ХИЙГДЛЭЭ';
}

// ---------- pure helpers (no Spreadsheet calls) ----------

// Column numbers (1 = A) of everything after the product name, for a block of `days` day columns.
// paired: the two shops' layout (a shelf price after D; the other shop's balance and count in
// place of "sold"). gift: a gift column after the transfer out (the paired layout only).
// shelf: the shelf price after D without the rest of the pairing (Тосгон); the paired layout has it.
// expenseLast: the last column of the expense block (the gift column, or the transfer out).
// A field a layout does not have is null.
function buildColumnLayout(days, { paired = false, gift = false, shelf = paired } = {}) {
  const incomeFirst = shelf ? SHELF_PRICE_COLUMN + 1 : FIRST_DAY_COLUMN;
  const transferIn = incomeFirst + days;
  const expenseFirst = transferIn + 1;
  const transferOut = expenseFirst + days;
  const giftColumn = paired && gift ? transferOut + 1 : null;
  const expenseLast = giftColumn || transferOut;
  return {
    days,
    paired,
    shelfPrice: shelf ? SHELF_PRICE_COLUMN : null,
    incomeFirst,
    transferIn,
    expenseFirst,
    transferOut,
    gift: giftColumn,
    expenseLast,
    count: expenseLast + 1,
    balance: expenseLast + 2,
    check: expenseLast + 3,
    sold: paired ? null : expenseLast + 4,
    // the other shop's count before its balance (the owner, 2026-10-28; the other way round before)
    otherCount: paired ? expenseLast + 4 : null,
    otherBalance: paired ? expenseLast + 5 : null,
    summaryFirst: expenseLast + (paired ? 6 : 5),
  };
}

const LAYOUT = buildColumnLayout(DAY_COLUMNS, { paired: PAIRED_SHOPS, gift: PAIRED_SHOPS && GIFT_COLUMN, shelf: SHELF_COLUMN });
// The paired layout as it was before the gift column (layout-migration.gs moves from it).
const PAIRED_BASE_LAYOUT = buildColumnLayout(DAY_COLUMNS, { paired: true });

// Where a column of a shop sheet of the old layout is in the paired layout: one to the right
// from E on (the shelf price), one more after the old "sold" column (the other shop's count).
// Used by the move (layout-migration.gs) and by the OCR's cell references (ocr-layout.gs).
function pairedColumnOf(oldColumn) {
  return oldColumn + (oldColumn >= SHELF_PRICE_COLUMN ? 1 : 0) + (oldColumn > buildColumnLayout(DAY_COLUMNS).sold ? 1 : 0);
}

// Where a column of a paired sheet without the gift column is once it has one: one to the right
// after the transfer out. Used by the gift move (layout-migration.gs) and the OCR (ocr-layout.gs).
function giftColumnOf(oldColumn) {
  return oldColumn + (oldColumn > PAIRED_BASE_LAYOUT.transferOut ? 1 : 0);
}

// A shop sheet's layout from its own header row 3: the shelf price header in column E marks the
// paired layout (a sheet not moved yet keeps the old one); then the gift header right after the
// transfer out marks the gift column. headerE, headerGift: those cells' text.
const shopLayoutOf = (headerE, days, headerGift = '') => {
  const paired = String(headerE).trim() === SHELF_PRICE_HEADER;
  return buildColumnLayout(days, { paired, gift: paired && String(headerGift).trim() === GIFT_HEADER });
};
// A Тосгон month sheet's layout from its own E3: the shelf price header marks a moved sheet.
const tosgonLayoutOf = (headerE) => buildColumnLayout(30, { shelf: String(headerE).trim() === SHELF_PRICE_HEADER });
// The column of row 3 that shopLayoutOf reads as headerGift, for a sheet with `days` day columns.
const giftHeaderColumn = (days) => buildColumnLayout(days, { paired: true }).transferOut + 1;

function daysInMonth(year, monthNumber) {
  return new Date(Date.UTC(year, monthNumber, 0)).getUTCDate();
}

function columnLetter(column) {
  let letters = '';
  for (let rest = column; rest > 0; rest = Math.floor((rest - 1) / 26)) {
    letters = String.fromCharCode(65 + ((rest - 1) % 26)) + letters;
  }
  return letters;
}

// A link that opens a workbook's tab with one cell chosen: a note names the cell it speaks of this
// way (the owner, 2026-10-26: "not the photo's link, the spreadsheet's").
const sheetCellLink = (bookId, gid, a1) => `https://docs.google.com/spreadsheets/d/${bookId}/edit#gid=${gid}&range=${a1}`;

function sheetNumber(name) {
  const match = String(name).trim().match(SHEET_NAME_PATTERN);
  return match ? Number(match[1]) : null;
}

function slotFromOldName(name) {
  const number = sheetNumber(name);
  if (number === null) {
    return null;
  }
  const chain = CHAINS.find(
    (candidate) =>
      candidate.firstNumber !== null &&
      number >= candidate.firstNumber &&
      number < candidate.firstNumber + CHAIN_PERIOD_COUNT
  );
  return chain ? { chainKey: chain.key, position: number - chain.firstNumber } : null;
}

// Period number counted from the first month of the chain: 0 = the first month (its
// first half where a month has two sheets). half: 0 or 1.
function periodPosition(year, month, half) {
  const monthIndex = (year - CHAIN_FIRST_MONTH.year) * MONTHS_PER_YEAR + (month - CHAIN_FIRST_MONTH.month);
  return monthIndex * PERIODS_PER_MONTH + half;
}

// The year, month and half (0 or 1) of a period number.
// A position before the chain (negative: last year's file, year-books.gs) counts back from it.
function periodMonth(position) {
  const monthIndex = Math.floor(position / PERIODS_PER_MONTH);
  const months = CHAIN_FIRST_MONTH.month - 1 + monthIndex;
  return {
    year: CHAIN_FIRST_MONTH.year + Math.floor(months / MONTHS_PER_YEAR),
    month: (((months % MONTHS_PER_YEAR) + MONTHS_PER_YEAR) % MONTHS_PER_YEAR) + 1,
    half: ((position % PERIODS_PER_MONTH) + PERIODS_PER_MONTH) % PERIODS_PER_MONTH,
  };
}

function slotFromTabName(name) {
  const match = String(name).trim().match(TAB_NAME_PATTERN);
  if (!match) {
    return null;
  }
  const chain = CHAINS.find((candidate) => candidate.tabPrefix.toLowerCase() === match[1].toLowerCase());
  const month = Number(match[3]);
  const hasDay = match[4] !== undefined;
  if (!chain || month < 1 || month > MONTHS_PER_YEAR || hasDay !== (PERIODS_PER_MONTH === 2)) {
    return null;
  }
  const half = hasDay && Number(match[4]) > FIRST_HALF_LAST_DAY ? 1 : 0;
  const position = periodPosition(Number(match[2]), month, half);
  return position < 0 || position >= CHAIN_PERIOD_COUNT ? null : { chainKey: chain.key, position };
}

// Which business and which period (0 = the first one of the chain) a tab is.
function sheetSlot(name) {
  return slotFromTabName(name) || slotFromOldName(name);
}

function periodLastDay(position) {
  const { year, month, half } = periodMonth(position);
  return PERIODS_PER_MONTH === 2 && half === 0 ? FIRST_HALF_LAST_DAY : daysInMonth(year, month);
}

// Tab name of a period: shop / year / month (/ last day), e.g. "Sor/2026/10/31", "Tos/2026/7".
function chainSheetName(chain, position) {
  const { year, month } = periodMonth(position);
  const parts = [chain.tabPrefix, year, month];
  return (PERIODS_PER_MONTH === 2 ? [...parts, periodLastDay(position)] : parts).join('/');
}

function canonicalSheetName(name) {
  const slot = sheetSlot(name);
  if (slot === null) {
    return String(name);
  }
  return chainSheetName(CHAINS.find((chain) => chain.key === slot.chainKey), slot.position);
}

// ---------- counts between the period ends (from the 10/31 sheets on) ----------
// A count made on another day than the period end (Тосгон's 10/23) is not the sheet's count: it
// gets a block of its own at the right edge of its sheet, after the five totals, three columns
// a counted day:
//   "Тооллого 10/23"  the count (the OCR writes it; only the owner and the shop account change it)
//   "Үлдэгдэл 10/23"  the books at the end of that day: opening + income - expense up to it
//   "Зөрүү 10/23"     the count minus that balance
// and A1 offers "Тооллого 10/23": the days up to the 23rd with that block (view-columns.gs).
// A1's choices of such a sheet are a row of the hidden sheet VIEW_LIST_SHEET (A the tab name,
// B.. the choices), so a newly counted day is one more cell there. The OCR's project reads this
// file too (build.sh), so the two write the same.
const SPOT_COUNTS_FROM = '2026-10-31'; // the first period end whose sheet takes them
const SPOT_BLOCK_WIDTH = 3;
const SPOT_LABELS = ['Тооллого', 'Үлдэгдэл', 'Зөрүү'];
const SPOT_DAY_PATTERN = /^Тооллого (\d{1,2})\/(\d{1,2})$/;
const VIEW_LIST_SHEET = '_харах';
const VIEW_LIST_WIDTH = 60; // choices a tab can have
const VIEW_ALL = 'Бүгд';
const VIEW_INCOME_ONLY = 'Зөвхөн орлого';
const VIEW_EXPENSE_ONLY = 'Зөвхөн зарлага';
const VIEW_PRINT = 'Хэвлэх (нийт дүнгүй)';
const VIEW_WRONG_ONLY = 'Зөвхөн буруу';
const VIEW_CHOICES = [VIEW_ALL, VIEW_INCOME_ONLY, VIEW_EXPENSE_ONLY, VIEW_PRINT, VIEW_WRONG_ONLY];
const VIEW_HALVES = ['1–15', '16–31']; // a month sheet (Тосгон): the days of one half only

const spotChoice = (month, day) => `${SPOT_LABELS[0]} ${month}/${day}`;
const spotHeaders = (month, day) => SPOT_LABELS.map((label) => `${label} ${month}/${day}`);
// The month and day of "Тооллого 10/23" (a choice, or a block's first header); null for other text.
function spotDayOf(text) {
  const match = String(text).trim().match(SPOT_DAY_PATTERN);
  return match ? { month: Number(match[1]), day: Number(match[2]) } : null;
}
// The first column of the blocks: right after the five totals.
const spotFirstColumn = (layout) => layout.summaryFirst + TOTAL_COUNT;
// Whether the sheet of the period that ends on this day ("2026-10-31") takes them.
const takesSpotCounts = (periodEndIso) => periodEndIso >= SPOT_COUNTS_FROM;
// A1's choices on such a sheet before any day is counted: a month sheet adds its halves.
const viewBaseChoices = (layout) => [...VIEW_CHOICES, ...(layout.days > FIRST_HALF_LAST_DAY ? VIEW_HALVES : [])];

// A tab's row of choices: the base ones, then the counted days by date, each once; plus `extra`.
function viewListValues(current, base, extra = '') {
  const counted = [...current, extra].map((value) => String(value).trim()).filter((value) => spotDayOf(value) !== null);
  const order = (text) => { const { month, day } = spotDayOf(text); return month * 100 + day; };
  return [...base, ...[...new Set(counted)].sort((a, b) => order(a) - order(b))];
}

// ---------- spreadsheet access ----------

// The cells of a tab's choices in VIEW_LIST_SHEET (the sheet is made, hidden, when missing),
// holding `base` and the counted days listed so far, plus `extra` when given. Writes; the caller
// minds DRY_RUN.
function viewListRange(spreadsheet, tabName, base, extra = '') {
  let list = spreadsheet.getSheetByName(VIEW_LIST_SHEET);
  if (!list) {
    list = spreadsheet.insertSheet(VIEW_LIST_SHEET);
    list.hideSheet();
  }
  if (list.getMaxColumns() < VIEW_LIST_WIDTH + 1) {
    list.insertColumnsAfter(list.getMaxColumns(), VIEW_LIST_WIDTH + 1 - list.getMaxColumns());
  }
  const last = list.getLastRow();
  const names = last > 0 ? list.getRange(1, 1, last, 1).getValues().map(([name]) => String(name).trim()) : [];
  const found = names.indexOf(tabName);
  const row = found >= 0 ? found + 1 : last + 1;
  const range = list.getRange(row, 2, 1, VIEW_LIST_WIDTH);
  const current = found >= 0 ? range.getValues()[0] : [];
  const wanted = viewListValues(current, base, extra);
  const padded = Array.from({ length: VIEW_LIST_WIDTH }, (_, index) => (index < wanted.length ? wanted[index] : ''));
  if (found < 0) {
    list.getRange(row, 1).setValue(tabName);
  }
  if (JSON.stringify(current) !== JSON.stringify(padded)) {
    range.setValues([padded]);
  }
  return range;
}

// A1 of a sheet chooses from its row of VIEW_LIST_SHEET.
function setViewListRule(sheet, listRange) {
  sheet.getRange(1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInRange(listRange, true).setAllowInvalid(false).build()
  );
}

function isInScope(slot, scope) {
  if (!scope) {
    return true;
  }
  const isBeforeEnd = scope.toPosition === undefined || slot.position <= scope.toPosition;
  return slot.chainKey === scope.chainKey && slot.position >= scope.fromPosition && isBeforeEnd;
}

// Entry sheets of each business, ordered by period. Needs isEntrySheet (year-sync.gs).
// scope (optional): { chainKey, fromPosition, toPosition? } limits the result to one
// business and to the periods from..to, so a change made in October does not read January.
function findChains(spreadsheet, scope) {
  const entries = spreadsheet
    .getSheets()
    .map((sheet) => ({ sheet, slot: sheetSlot(sheet.getName()) }))
    .filter((entry) => entry.slot !== null && isInScope(entry.slot, scope))
    .filter((entry) => isEntrySheet(entry.sheet));
  return CHAINS.map((chain) => ({
    label: chain.key,
    sheets: entries
      .filter((entry) => entry.slot.chainKey === chain.key)
      .sort((a, b) => a.slot.position - b.slot.position)
      .map((entry) => entry.sheet),
  })).filter((chain) => chain.sheets.length > 0);
}
