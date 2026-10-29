// Checks of year-books.gs and year-rollover.gs: the new year's workbook made from December, the
// registry, and the scripts working on the right year's file. Usage: node tests/rollover_test.js <dir>
const fs = require("fs"), assert = require("assert");
const { Spreadsheet, makeRunnerFor } = require("./fake_book.js");

const dir = process.argv[2];
const FIRST_FILES = ["profile.gs", "config.gs"];
const gsFiles = fs.readdirSync(dir).filter((f) => f.endsWith(".gs"));
const loadOrder = [...FIRST_FILES.filter((f) => gsFiles.includes(f)), ...gsFiles.filter((f) => !FIRST_FILES.includes(f))];
const code = loadOrder.map((f) => fs.readFileSync(`${dir}/${f}`, "utf8")).join("\n");
let n = 0;
const eq = (a, b, m) => { assert.deepStrictEqual(a, b, m); n++; };
const ok = (c, m) => { assert(c, m); n++; };

const BAND = "#f6b26b";
const band = (s, row) => { for (let c = 1; c <= 36; c++) s.cell(row, c).bg = BAND; };
// A December sheet: a category and 22 priced products with income and expense.
function decemberSheet(book, name) {
  const s = book.insertSheet(name);
  s.getRange(2, 1).setValue(name);
  s.getRange(3, 5).setValue("яг харагдах үнэ");
  s.getRange(4, 1).setValue("Хөргөгч"); band(s, 4);
  Array.from({ length: 22 }, (_, i) => `Haier ${100 + i}l`).forEach((label, i) => {
    s.getRange(5 + i, 1).setValue(label);
    s.getRange(5 + i, 2, 1, 4).setValues([[3, 500, 700 + i, ""]]);
    s.getRange(5 + i, 7).setValue(2); // income
    s.getRange(5 + i, 24).setValue(1); // expense
  });
  return s;
}

function world() {
  const home = new Spreadsheet(); home.name = "бараа орлого зарлага 2026";
  const registry = new Spreadsheet(); registry.name = "Он бүрийн файл";
  const tab = registry.insertSheet("Файлууд");
  tab.getRange(1, 1, 1, 4).setValues([["Он", "Үндсэн", "Тосгон", "Эхлэх"]]);
  tab.getRange(2, 1, 1, 4).setValues([[2026, home.getId(), "tos-2026", "2026-01-01"]]);
  decemberSheet(home, "My/2026/12/31");
  decemberSheet(home, "Sor/2026/12/31");
  home.insertSheet("Бараа").getRange(4, 1, 1, 2).setValues([["Haier 100l", "Хоёулаа"]]);
  home.insertSheet("Sheet200").getRange(1, 1).setValue("сар бүрийн дүн");
  const state = {
    triggers: [], mails: [], books: { [home.getId()]: home, [registry.getId()]: registry },
    docProps: { pairedLayoutReady: "1", giftColumnReady: "1" },
    shares: { [home.getId()]: { editors: ["mom@example.com"], viewers: [] } },
  };
  const withRegistry = (c) => c.replace(/const REGISTRY_ID = '[^']*';/, `const REGISTRY_ID = '${registry.getId()}';`);
  const runner = makeRunnerFor(withRegistry(code))(home, state);
  return { home, registry, tab, state, runner, run: (expr) => runner(false, expr) };
}
const at = (y, m, d) => `new Date(${y}, ${m - 1}, ${d})`;
const rollover = (y, m, d) => `runYearRollover(${at(y, m, d)}, Date.now() + 60000)`;

// ---------- pure rules ----------
{
  const { run } = world();
  eq(run(`[${at(2026, 12, 19)}, ${at(2026, 12, 20)}, ${at(2026, 12, 31)}, ${at(2027, 1, 5)}, ${at(2027, 2, 1)}].map(rolloverPhase)`).value,
     [null, "prepare", "prepare", "carry", null], "prepared from 20 December, carried into in January");
  eq(run(`[newYearOf(${at(2026, 12, 20)}), newYearOf(${at(2027, 1, 3)})]`).value, [2027, 2027], "the year being made");
  const registry = `parseRegistry([[2026, "a", "t", "2026-01-01"], [2027, "b", "u", "2027-01-01"]])`;
  eq(run(`[entryForDay(${registry}, "2026-12-31").main, entryForDay(${registry}, "2027-01-01").main]`).value, ["a", "b"], "31 December is still 2026's file, 1 January 2027's");
  eq(run(`[chainSpanOf(2026).first, chainSpanOf(2027).first, chainSpanOf(2027).months]`).value, [{ year: 2026, month: 1 }, { year: 2027, month: 1 }, 12], "every new year runs January to December");
}

// ---------- 20 December: the new file ----------
const { home, tab, state, run, runner } = world();
run(rollover(2026, 12, 20));
const newId = JSON.parse(state.props["rollover:main:2027"]).id;
const book = state.books[newId];
ok(book && book.getName() === "бараа орлого зарлага 2027", "the 2027 workbook is made");
eq(book.getSheets().map((s) => s.getName()).sort(), ["Sheet200", "Sor/2027/1/15", "My/2027/1/15", "Бараа"].sort(), "December's sheets as January's, the list and Sheet200 copied; the empty Sheet1 gone");
const jan = book.getSheetByName("Sor/2027/1/15");
eq([jan.cell(5, 1).v, jan.cell(5, 4).v, jan.cell(4, 1).v, jan.cell(4, 1).bg], ["Haier 100l", 700, "Хөргөгч", BAND], "with December's products, prices and categories");
eq([jan.cell(5, 2).v, jan.cell(5, 7).v, jan.cell(5, 24).v], ["", "", ""], "income, expense and the opening balance emptied");
eq([tab.cell(3, 1).v, tab.cell(3, 2).v, tab.cell(3, 4).v], [2027, newId, "2027-01-01"], "the registry knows the 2027 file from 1 January");
ok(state.shares[newId].editors.includes("mom@example.com") && state.shares[newId].viewers.some((e) => e.startsWith("sheet-reader@")), "the same people can edit it; the read-only account can read it");
ok(state.props.setupBook === newId && state.triggers.some((t) => t.getHandlerFunction() === "setupYear"), "setupYear is started on the new file");
ok(state.docProps[`${newId}:pairedLayoutReady`] === "1" && state.docProps[`${newId}:giftColumnReady`] === "1", "the new file's layout flags are its own");
eq(JSON.parse(state.props["rollover:main:2027"]).step, "wait", "and the rollover waits for the setup");

// ---------- setup done, the owner mailed ----------
run(`markYearSetupDone(SpreadsheetApp.openById("${newId}"))`);
run(rollover(2026, 12, 21));
ok(state.mails.length === 1 && state.mails[0][1].includes("2027") && JSON.parse(state.props["rollover:main:2027"]).step === "done", "the owner gets the link once the file is ready");

// ---------- 25 December: a new product and a new price in December ----------
const dec = home.getSheetByName("Sor/2026/12/31");
dec.getRange(27, 1).setValue("Haier 105.5l"); dec.getRange(27, 2, 1, 4).setValues([[1, 400, 650, ""]]);
dec.cell(6, 4).v = 999; // Haier 101l's price
run(rollover(2026, 12, 25));
const labels = Array.from({ length: 24 }, (_, i) => jan.cell(5 + i, 1).v);
ok(labels.indexOf("Haier 105.5l") === labels.indexOf("Haier 105l") + 1, "December's new product reaches January, in its place in name order");
eq(jan.cell(5 + labels.indexOf("Haier 101l"), 4).v, 999, "and a changed price too");

// ---------- 3 January: December's count typed late ----------
const countColumn = run("LAYOUT.count").value;
dec.cell(5, countColumn).v = 7; // Haier 100l counted on 31 December
run(rollover(2027, 1, 3));
eq(jan.cell(5, 2).v, 7, "the count of 31 December is January's opening balance");

// ---------- the scripts work on the right year's file ----------
eq(run(`[sheetSlot("Sor/2027/1/15"), positionOfDate(${at(2027, 1, 5)})]`).value, [null, 24], "a run on the 2026 file: January 2027 is past its end");
eq(run(`(() => { useBook(SpreadsheetApp.openById("${newId}")); return [sheetSlot("Sor/2027/1/15").position, positionOfDate(${at(2027, 1, 5)}), CHAIN_PERIOD_COUNT]; })()`).value,
   [0, 0, 24], "a run on the 2027 file reads its tabs as the 2027 chain");
eq(run(`booksToMaintain(${at(2027, 1, 5)}).map((b) => b.getName())`).value, ["бараа орлого зарлага 2027", "бараа орлого зарлага 2026"], "in January the hourly job looks after both years");
eq(run(`booksToMaintain(${at(2027, 2, 1)}).map((b) => b.getName())`).value, ["бараа орлого зарлага 2027"], "from February only the new one");
eq(run(`[targetBook({ source: SpreadsheetApp.openById("${newId}") }).getName(), CHAIN_FIRST_MONTH.year]`).value, ["бараа орлого зарлага 2027", 2027], "an edit of the 2027 file works on 2027");
ok(runner(true, "1") && state.docProps.pairedLayoutReady === "1" && state.docProps[`${newId}:priceLinkReady`] === undefined, "one file's flags never count for another");
{
  const plain = makeRunnerFor(code.replace(/const REGISTRY_ID = '[^']*';/, "const REGISTRY_ID = '';"))(new Spreadsheet(), { triggers: [], mails: [] });
  eq(plain(false, `[REGISTRY_ID, runYearRollover(${at(2026, 12, 20)}, Date.now())]`).value,
     ["", ["шинэ он: бүртгэл (REGISTRY_ID) алга эсвэл хоосон тул хийсэнгүй"]], "without a registry nothing moves: today's files as always");
  const unreadable = makeRunnerFor(code)(new Spreadsheet(), { triggers: [], mails: [] });
  eq(unreadable(false, `[readRegistry().length, runYearRollover(${at(2026, 12, 20)}, Date.now()).length]`).value, [0, 1], "a registry the run cannot open: today's files as always");
}
{
  const dry = world();
  const lines = dry.runner(true, rollover(2026, 12, 20)).value;
  ok(lines.some((line) => line.includes('"бараа орлого зарлага 2027" файл үүсгэнэ')) && Object.keys(dry.state.books).length === 2 && !(dry.state.props || {})["rollover:main:2027"], "a trial run makes nothing");
}

console.log(`ALL ${n} ROLLOVER CHECKS PASSED`);
