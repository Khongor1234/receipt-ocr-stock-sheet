/**
 * The settings of THIS workbook. Everything that differs from one workbook to the
 * other lives here; config.gs and every other file only read it. This file is for
 * the main workbook "бараа орлого зарлага" (two shops, one sheet per half-month).
 * The Тосгон workbook (one warehouse, one sheet per month) has its own file,
 * profiles/tosgon.gs; `sh build.sh tosgon` puts it in the build in place of this one.
 *
 * DRY_RUN is the ONE switch for all scripts:
 *   true  = nothing is changed, every function only writes what it WOULD do to
 *           the log (Execution log)
 *   false = changes are really made
 */

const DRY_RUN = true;

// Day columns of one block: the income block and the expense block have this many
// each. 15 = a half-month sheet (day 31 is entered in the last day column of the
// second half), 30 = a month sheet (day 31 is entered in the "30" column).
const DAY_COLUMNS = 15;
// How many sheets one month has: 2 = half-month sheets, 1 = one sheet a month.
const PERIODS_PER_MONTH = 2;
// true: the two shops' layout (config.gs buildColumnLayout): a hand-typed shelf price after D,
// and the other shop's final balance and count in place of the "sold" column.
const PAIRED_SHOPS = true;
// true: a hand-typed shelf price "яг харагдах үнэ" in E (config.gs buildColumnLayout); it comes
// with the paired layout here.
const SHELF_COLUMN = PAIRED_SHOPS;
// true: a "Бэлэг" column after the transfer out (config.gs buildColumnLayout): the goods given
// away, taken off the final balance and, at cost, off the sales profit (not counted as sold).
const GIFT_COLUMN = true;
// The kind of workbook this project serves, and the one it is bound to: its year, first month and
// months. The workbooks of the other years are January to December (year-books.gs).
const BOOK_KIND = 'main';
const HOME_CHAIN = { year: 2026, first: { year: 2026, month: 1 }, months: 12 };
// The first month of every chain of sheets and how many months it runs: those of the workbook a
// run works on (useBook, year-books.gs), the home workbook's until then.
let CHAIN_FIRST_MONTH = HOME_CHAIN.first;
let CHAIN_MONTHS = HOME_CHAIN.months;
// The yearly summary sheet "Sheet200" and its charts belong to this workbook.
const GUARD_EXTRA_SHEETS = ['Sheet200'];
// true: the opening balance follows the count of the sheet before, and where nothing
// was counted, that sheet's computed final balance (opening + income - expense).
// false: it follows the count only, and stays free to type where nothing was counted.
const CARRY_FROM_BALANCE = false;
// Sheets whose count column (AK) holds balances copied by hand when the half-month was closed, not
// a count (Соронзон was not counted 2026-04-01 to 08-15): the next sheet's opening balance follows
// the computed final balance, so goods typed in after the copy are carried too.
// Empty for now (the owner, 2026-10-23): Соронзон 4/15-8/15 already have count = final balance,
// except four products whose negative balance is to stay where it is, not be carried on. The
// list was: 'Sor/2026/4/15', 'Sor/2026/4/30', 'Sor/2026/5/15', 'Sor/2026/5/31', 'Sor/2026/6/15',
// 'Sor/2026/6/30', 'Sor/2026/7/15', 'Sor/2026/7/31', 'Sor/2026/8/15'.
const UNCOUNTED_SHEETS = [];

// tabPrefix: start of the tab name. firstNumber: the old name ("Sheet1", "Sheet31", ...)
// the year was built with; old names are still understood and get renamed
// (null = this shop never had such names).
// rowLimit: the sheets of this shop are kept at this many rows (trim-rows.gs).
//           0 = the rows of this shop are never removed.
const CHAINS = [
  { key: 'my home', label: 'My home', tabPrefix: 'My', firstNumber: 1, rowLimit: 250 },
  { key: 'соронзон', label: 'Соронзон', tabPrefix: 'Sor', firstNumber: 31, rowLimit: 400 },
];
const PERIOD_BUSINESS_LABELS = { 'my home': 'My home', 'соронзон': 'Соронзон' };
// Тооллогын хуудас (stock-sheets.gs): each tab shows its shop's stock of TODAY and is
// locked for everyone except the file owner. chainKey: a key of CHAINS.
const STOCK_SHEETS = [
  { name: 'Sheet100', chainKey: 'соронзон' },
  { name: 'Sheet100*', chainKey: 'my home' },
];
// Тооллого хийх хуудас (count-sheets.gs): only the product names (locked) and an empty
// count column for the 15th / last-day count. chainKey: a key of CHAINS.
const COUNT_SHEETS = [
  { name: 'Соронзон тооллого', chainKey: 'соронзон' },
  { name: 'My home тооллого', chainKey: 'my home' },
];

// Categories that must exist. Each one sits directly above the first product whose
// name starts with firstProductPrefix. Its name is written back if it was erased,
// and the whole row is made again if it was deleted.
const DIVIDER_REQUIRED_CATEGORIES = [
  { chainKey: 'соронзон', name: 'Плетик', firstProductPrefix: 'delux 50-60', color: '#b45f06' },
];

// Steps of setupYear that only this workbook needs (none here): the function gets the
// spreadsheet and the deadline and returns steps like the ones in year-setup.gs.
function extraSetupSteps(spreadsheet, deadline) {
  return [];
}

// shop key (profile.gs CHAINS): [[spelling to keep, [other spellings, as keys]], ...]
const NAME_FIXES = {
  'my home': [
    ['8037 ус', ['8037 us', '8037 гы', '8037us']],
    ['Aucma 341l', ['acuma 341l']],
    ['Haier 10-858', ['haier10-858']],
    ['Meiling 145l', ['meiling145l']],
    ['Meiling 201l', ['miling 201l']],
    ['Pnasonica 160l', ['panasonica 160l']],
    ['TGL 201L', ['tgl 201']],
    ['Tоль 250', ['толь 250', 'толь250']],
    ['motor 4800', ['мотор 4800']],
  ],
  'соронзон': [
    ['Acuma 196l', ['aucma 196l']],
    ['Cmidea 80', ['cmidea80', 'cmidea80l']],
    ['Haer 2028', ['haier 2028']],
    ['Haier kond', ['haier kon']],
    ['Meiling 145', ['meiling 145 saar']],
    ['Meiling 202l', ['meiling 202l saar', 'meiling 202l саарал']],
    ['Meiling 231l', ['meiling231l']],
    ['Meiling 254l', ['meiling254', 'meiling254l']],
    ['Meiling 90l', ['meiling 90']],
    ['Pana чийгшүүлэгч', ['pana чийгшүүлэгч жижиг']],
    ['Panasonica 10 авто', ['panasonica 10 abto']],
    ['Panasonica 13.5l', ['panasonica 13,5l', 'panasonica 13.5']],
    ['Panasonica 210', ['panasonica суурь 210']],
    ['Panasonica 4l', ['panasonica4l']],
    ['Pilipis 3000p', ['pilipis 3000']],
    ['Ronshin 1.200', ['ronshin 1200']],
    ['Table', ['taube']],
    ['Xinix 150l', ['xinix 150']],
    ['cx 50l', ['lc 50l']],
    ['kiston 27.5l', ['kitson 27.5']],
    ['Камин 990', ['kamin 990']],
    ['Ус буцагагч 1819', ['ус буцалгагч 1819']],
    ['Ус буцалгагч 1705', ['ус буцал1705']],
    ['Цагаан тень', ['урт цагаан тень']],
    ['маск 15', ['mack 15']],
    ['маск 20', ['mask 20']],
    ['маск 25', ['mack 25']],
    ['маск 27', ['mack 27']],
    ['маск 30', ['mack 30']],
    ['машины хөлдөөгч 40l', ['машины хөл 40l']],
    ['машины хөлдөөгч 50Л', ['машин хөл 50l', 'машины хөл 50л']],
    ['сав 70', ['сав70']],
  ],
};
