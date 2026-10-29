// Checks of the two shops' layout (PAIRED_SHOPS, config.gs): the shelf price after D, the other
// shop's balance and count in place of "sold", and the one-time move of the sheets
// (layout-migration.gs). Usage: node tests/paired_test.js <project dir>
const fs = require("fs"), assert = require("assert");
const { Spreadsheet, makeRunnerFor } = require("./fake_book.js");

const dir = process.argv[2];
const FIRST_FILES = ["profile.gs", "config.gs"];
const gsFiles = fs.readdirSync(dir).filter((f) => f.endsWith(".gs"));
const loadOrder = [...FIRST_FILES.filter((f) => gsFiles.includes(f)), ...gsFiles.filter((f) => !FIRST_FILES.includes(f))];
const shipped = [...loadOrder, "profiles/sheet-rows.gs"].map((f) => fs.readFileSync(`${dir}/${f}`, "utf8")).join("\n");
assert(shipped.includes("const PAIRED_SHOPS = true;"), "the main workbook ships with the paired layout");
// These checks are the move to the paired layout as it was made (2026-10), before the gift
// column; the gift column has its own checks (tests/gift_test.js).
assert(shipped.includes("const GIFT_COLUMN = true;"), "the main workbook ships with the gift column");
const code = shipped.replace("const GIFT_COLUMN = true;", "const GIFT_COLUMN = false;");
const makeRunner = makeRunnerFor(code);
let n = 0;
const ok = (c, m) => { assert(c, m); n++; };
const eq = (a, b, m) => { assert.strictEqual(JSON.stringify(a), JSON.stringify(b), m); n++; };

// ---------- the layout and every constant that follows it ----------
{
  const run = makeRunner(new Spreadsheet(), { triggers: [], mails: [] });
  const v = (e) => run(true, e).value;
  eq(v("LAYOUT"), { days: 15, paired: true, shelfPrice: 5, incomeFirst: 6, transferIn: 21, expenseFirst: 22, transferOut: 37, gift: null, expenseLast: 37, count: 38, balance: 39, check: 40, sold: null, otherCount: 41, otherBalance: 42, summaryFirst: 43 }, "E shelf price, F..T income, U in, V..AJ expense, AK out, AL count, AM balance, AN check, AO/AP the other shop's count and balance (2026-10-28), AQ.. totals");
  eq(v("[LAYOUT.shelfPrice, LAYOUT.incomeFirst, LAYOUT.transferIn, LAYOUT.expenseFirst, LAYOUT.transferOut, LAYOUT.count, LAYOUT.balance, LAYOUT.check, LAYOUT.otherBalance, LAYOUT.otherCount, LAYOUT.summaryFirst].map(columnLetter)"), ["E", "F", "U", "V", "AK", "AL", "AM", "AN", "AP", "AO", "AQ"], "the letters");
  eq(v("[CALC_FIRST_COL, CALC_COLUMN_COUNT, SUMMARY_FIRST_COL, SETUP_ENTRY_FIRST_COL, SETUP_ENTRY_LAST_COL, SYNC_COUNT_COL, INSTANT_COUNT_COLUMN, MONEY_PRICE_COL_COUNT, HEADER_MAX_COLUMNS]"), [39, 4, 43, 6, 38, 38, 38, 3, 47], "every column constant follows the layout");
  eq(v("INSTANT_WATCHED_COLUMNS"), [1, 3, 4, 5], "a typed shelf price is carried at once, like the price");
  eq(v("VIEW_PRICE_COLUMNS"), { first: 3, count: 3 }, "the print view hides C..E");
  eq(v("DIVIDER_GRID_ZONES"), [{ firstColumn: 1, width: 5 }, { firstColumn: 38, width: 5 }], "the grid covers A..E and AL..AP");
  eq(v("PRICE_FIELDS.map((f) => [f.label, f.column])"), [["өртөг", 3], ["үнэ", 4], ["яг харагдах үнэ", 5]], "the shelf price is carried forward like cost and price");
  eq(v("SHEET200_TOTAL_CELLS"), ["$AT$3", "$AU$3"], "Sheet200 reads the sales totals from AT3, AU3");
  eq(v("[SYNC_SHELF_PRICE_COL, SYNC_DESCRIBE_LAST_COL]"), [5, 5], "a product added to a later sheet takes its shelf price along");
  ok(v("SUMMARY_LINES[3].formula") === "=SUMPRODUCT(D4:D,BYROW(V4:AJ,LAMBDA(r,SUM(r))))", "Нийт борлуулалт: the expense days V..AJ (not AK, the transfer out)");
  ok(v("SUMMARY_LINES[4].formula") === "=SUMPRODUCT(D4:D,BYROW(V4:AJ,LAMBDA(r,SUM(r))))-SUMPRODUCT(C4:C,BYROW(V4:AJ,LAMBDA(r,SUM(r))))", "Нийт борлуулалтын ашиг");
  ok(v("SUMMARY_LINES[2].formula") === "=SUMPRODUCT(D4:D,AL4:AL)", "Нийт эцсийн үлдэгдэл: the count AL");
  eq(v(`[otherShopOf("My/2026/10/15"), otherShopOf("Sor/2026/2/28"), otherShopOf("Sheet100")]`), [{ label: "Соронзон", tab: "Sor/2026/10/15" }, { label: "My home", tab: "My/2026/2/28" }, null], "the other shop's tab of the same half-month");
  const f = v(`buildCalcFormulas(["эцсийн үлдэгдэл ", "үнэлгээ ", ...calcPairedHeaders("Соронзон")], "Sor/2026/10/15")`);
  ok(f.length === 4, "four formula columns");
  ok(f[0].includes("BYROW(F4:U,") && f[0].includes("BYROW(V4:AK,"), "the balance: income F..U (with the transfer in) minus expense V..AK");
  ok(f[1].includes("AL4:AL=AM4:AM"), "the check: count AL against balance AM");
  ok(f[2].startsWith('=ARRAYFORMULA({"Соронзон тооллого";') && f[2].includes("'Sor/2026/10/15'!AL4:AL") && f[2].includes("IF(LEN('Sor/2026/10/15'!AL4:AL),"), "the other shop's count first (the owner, 2026-10-28): AL of its tab; not counted stays blank (not 0)");
  ok(f[3].startsWith('=ARRAYFORMULA({"Соронзон эцсийн үлдэгдэл";') && f[3].includes("'Sor/2026/10/15'!AM4:AM") && f[3].includes("VLOOKUP(REGEXREPLACE(LOWER(A4:A),\"\\s\",\"\")"), "then its balance: AM of its tab, found by the name without case and spaces");
  ok(f.slice(2).every((x) => x.includes("IFERROR(") && x.includes("IF(BYROW(B4:D,")), "a product the other shop does not have is blank; dividers stay blank");
  eq(v("[1, 4, 5, 20, 36, 37, 38, 39, 40, 41, 45].map(pairedColumnOf)"), [1, 4, 6, 21, 37, 38, 39, 40, 41, 43, 47], "an old column's place after the move");
  eq(v(`movePairedColors({ "haier 12": { "1": "#f00", "7": "#0f0", "41": "#00f" } })`), { "haier 12": { "1": "#f00", "8": "#0f0", "43": "#00f" } }, "the owner's colours move with their columns");
  eq(v(`findPriceChanges({ "a": [1, 2] }, { "a": [1, 2, 5] })`), [], "a shelf price the saved state does not know yet is no change (baseline)");
  eq(v(`findPriceChanges({ "a": [1, 2, 5] }, { "a": [1, 2, 6] }).map((c) => [c.index, c.from, c.to])`), [[2, 5, 6]], "a changed shelf price is a change");
}

// ---------- the move: a workbook of the old layout ----------
const OLD = { income: 5, inT: 20, expense: 21, outT: 36, count: 37, balance: 38, check: 39, sold: 40, total: 41 };
function oldSheet(ss, name, title, rows) {
  const s = ss.insertSheet(name);
  s.maxCols = 45;
  s.getRange(1, OLD.income).setValue(1);
  s.getRange(2, 1).setValue(title);
  [["Барааны нэрс", 1], ["анхны үлдэгдэл ", 2], ["өртөг үнэ", 3], ["зарах үнэ", 4], ["Орлого ", OLD.income], ["Зарлага", OLD.expense], ["тооллого", OLD.count]].forEach(([text, c]) => s.getRange(3, c).setValue(text));
  s.getRange(3, OLD.balance).setFormula("=ARRAYFORMULA({\"эцсийн үлдэгдэл \";E4:T})");
  s.getRange(3, OLD.check).setFormula("=ARRAYFORMULA({\"үнэлгээ \";AK4:AK})");
  s.getRange(3, OLD.sold).setFormula("=ARRAYFORMULA({\"борлуулалтын тоо\";U4:AI})");
  ["Нийт анхны үлдэгдэл ашиг", "Нийт анхны үлдэгдэл", "Нийт эцсийн үлдэгдэл", "Нийт борлуулалт", "Нийт борлуулалтын ашиг"].forEach((label, i) => { s.getRange(1, OLD.total + i).setValue(label); s.getRange(3, OLD.total + i).setFormula(`=OLD${i}`); });
  rows.forEach(([label, opening, income, expense, count], i) => {
    const r = 4 + i;
    s.getRange(r, 1).setValue(label); s.getRange(r, 2).setValue(opening); s.getRange(r, 3).setValue(100); s.getRange(r, 4).setValue(150);
    if (income) s.getRange(r, OLD.income).setValue(income);
    if (expense) s.getRange(r, OLD.expense).setValue(expense);
    if (count !== null) s.getRange(r, OLD.count).setValue(count);
  });
  return s;
}
const products = (shift) => Array.from({ length: 22 }, (_, i) => [`Haier ${i + 1}`, 2, i === 0 ? 1 : 0, i === 0 ? 1 : 0, i === 0 ? 2 + shift : null]);
{
  const ss = new Spreadsheet();
  const my = oldSheet(ss, "My/2026/10/15", "My home 2026 10 сарын 1-15", products(1));
  const sor = oldSheet(ss, "Sor/2026/10/15", "Соронзон 2026 10 сарын 1-15", products(-1));
  const state = { triggers: [], mails: [], docProps: { [`approvedColors:${my.getSheetId()}`]: JSON.stringify({ "haier 1": { "7": "#ffff00" } }), [`header:${my.getSheetId()}`]: "[]" } };
  const run = makeRunner(ss, state);

  // before the move every layout-dependent entry point waits
  const before = run(false, "(runBackgroundJobs(), onSheetEdit({ range: ss.getSheetByName('My/2026/10/15').getRange(4, 3) }), isPairedLayoutReady())");
  ok(before.value === false && before.logs.includes("migrateToPairedLayout"), "until the move the hourly job does nothing and says why");
  ok(my.getRange(3, 5).getValue() === "Орлого ", "and nothing was changed");

  const dry = run(true, "runPairedMigration(ss, Date.now() + 60000)");
  ok(dry.value.length === 3 && dry.value[2].includes("2 хуудсанд багана нэмж"), "a dry run only reports");
  ok(my.getMaxColumns() === 45 && my.getRange(3, 5).getValue() === "Орлого ", "a dry run changes nothing");

  const real = run(false, "runPairedMigration(ss, Date.now() + 60000)");
  ok(real.value[real.value.length - 1].startsWith("Бүх 2 хуудас"), "the move finishes");
  eq(state.docProps.pairedLayoutReady, "1", "and marks the workbook ready");
  ok(my.getMaxColumns() === 47, "two columns more");
  eq([my.getRange(3, 4).getValue(), my.getRange(3, 5).getValue(), my.getRange(3, 6).getValue(), my.getRange(3, 22).getValue(), my.getRange(3, 38).getValue()], ["зарах үнэ", "яг харагдах үнэ", "Орлого ", "Зарлага", "тооллого"], "row 3: D price, E shelf price, F income, V expense, AL count");
  eq([my.getRange(4, 6).getValue(), my.getRange(4, 22).getValue(), my.getRange(4, 38).getValue(), my.getRange(4, 5).getValue()], [1, 1, 3, ""], "the numbers moved one column right; E is empty for mom to type");
  ok(my.getRange(1, 5).getValue() === "" && my.getRange(1, 6).getValue() === 1, "row 1: the day numbers moved too, E1 stays empty");
  const header = my.getRange(3, 39, 1, 4).getFormulas()[0];
  ok(header[0].includes("BYROW(F4:U,") && header[1].includes("AL4:AL=AM4:AM"), "the balance and the check sit in AM, AN");
  ok(header[2].includes("'Sor/2026/10/15'!AL4:AL") && header[3].includes("'Sor/2026/10/15'!AM4:AM"), "My home's AO, AP read Соронзон's count and balance (the count first, 2026-10-28)");
  { // a sheet still with the balance first (before 2026-10-28): swapOtherShopColumns writes the count first and remembers the header anew
    const was = my.getRange(3, 41, 1, 2).getFormulas()[0];
    my.getRange(3, 41, 1, 2).setFormulas([[was[1], was[0]]]);
    run(false, "runRememberHeaders(ss)");
    const dryLines = run(true, "swapOtherShopColumns()").value;
    ok(my.getRange(3, 41).getFormula() === was[1] && dryLines.some((line) => line.includes("тавина")), "DRY_RUN only says what it would write");
    run(false, "swapOtherShopColumns()");
    const now = my.getRange(3, 41, 1, 2).getFormulas()[0];
    ok(now[0].startsWith('=ARRAYFORMULA({"Соронзон тооллого";') && now[1].startsWith('=ARRAYFORMULA({"Соронзон эцсийн үлдэгдэл";'), "AO the other shop's count, AP its balance");
    ok(JSON.parse(state.docProps[`header:${my.getSheetId()}`])[2][40] === now[0], "the remembered header holds the new formulas, so the hourly restore keeps them");
    const restored = run(false, "runRestoreHeaders(ss)").value;
    ok(my.getRange(3, 41).getFormula() === now[0] && restored.length === 0, "and the restore changes nothing");
  }
  ok(sor.getRange(3, 41).getFormula().includes("'My/2026/10/15'!AL4:AL"), "and Соронзон's AO reads My home's count");
  ok(my.getRange(3, 43).getFormula() === "=SUMPRODUCT(B4:B,D4:D)-SUMPRODUCT(B4:B,C4:C)" && my.getRange(3, 46).getFormula() === "=SUMPRODUCT(D4:D,BYROW(V4:AJ,LAMBDA(r,SUM(r))))", "the totals sit in AQ..AU, sales summed from the expense days");
  eq(my.getRange(1, 43, 1, 5).getValues()[0], ["Нийт анхны үлдэгдэл ашиг", "Нийт анхны үлдэгдэл", "Нийт эцсийн үлдэгдэл", "Нийт борлуулалт", "Нийт борлуулалтын ашиг"], "with their labels");
  eq(JSON.parse(state.docProps[`approvedColors:${my.getSheetId()}`]), { "haier 1": { "8": "#ffff00" } }, "the owner's colour of the old G is kept for the new H");
  const savedHeader = JSON.parse(state.docProps[`header:${my.getSheetId()}`]);
  ok(Array.isArray(savedHeader) && savedHeader.length === 3, "the header is remembered again from the new columns");
  ok(my.getRange(1, 5).getValue() === "", "the old remembered header did not fill the new E1");

  ok(state.docProps[`pairedFinished:${my.getSheetId()}`] === "1" && state.docProps[`pairedFinished:${sor.getSheetId()}`] === "1", "each finished sheet is marked");
  my.getRange(3, 41).setFormula("=MARK"); // a finished sheet is not done again by a later run
  const again = run(false, "runPairedMigration(ss, Date.now() + 60000)");
  ok(my.getMaxColumns() === 47 && again.value[again.value.length - 1].startsWith("Бүх 2 хуудас"), "a second run inserts nothing more");
  ok(my.getRange(3, 41).getFormula() === "=MARK", "and skips the sheets already finished");
  ok(run(false, "isPairedLayoutReady()").value === true, "from now on the scripts work again");
  state.docProps[`dividerBorders:${my.getSheetId()}`] = "1";
  run(false, "redrawTableBorders()");
  ok(state.docProps.pairedBordersReset === "1" && state.docProps[`dividerBorders:${my.getSheetId()}`] === "1+5,38+5", "redrawTableBorders: the grid is drawn again by the divider pass, which marks the sheet with its grid");
  ok(my.getRange(5, 42).sheet.cell(5, 42).border === true && my.getRange(5, 5).sheet.cell(5, 5).border === true, "the grid now boxes the new columns E and AP");
}

// ---------- a run cut off by the time limit ----------
{
  const ss = new Spreadsheet();
  oldSheet(ss, "My/2026/10/15", "My home 2026 10 сарын 1-15", products(0));
  oldSheet(ss, "Sor/2026/10/15", "Соронзон 2026 10 сарын 1-15", products(0));
  const state = { triggers: [], mails: [] };
  const run = makeRunner(ss, state);
  const cut = run(false, "runPairedMigration(ss, Date.now() - 1)");
  ok(cut.value[cut.value.length - 1].includes("дахин ажиллуулна уу") && state.docProps === undefined, "out of time: nothing is marked ready");
  const rest = run(false, "runPairedMigration(ss, Date.now() + 60000)");
  ok(rest.value[rest.value.length - 1].startsWith("Бүх 2 хуудас") && ss.getSheetByName("Sor/2026/10/15").getMaxColumns() === 47, "the next run finishes it");
}

// ---------- a run stopped halfway (the first column in, the second not) goes on where it stopped ----------
{
  const ss = new Spreadsheet();
  const half = oldSheet(ss, "My/2026/10/15", "My home 2026 10 сарын 1-15", products(0));
  oldSheet(ss, "Sor/2026/10/15", "Соронзон 2026 10 сарын 1-15", products(0));
  half.insertColumnsAfter(4, 1); // stopped right after the first insert
  const run = makeRunner(ss, { triggers: [], mails: [] });
  const done = run(false, "runPairedMigration(ss, Date.now() + 60000)");
  ok(done.value[done.value.length - 1].startsWith("Бүх 2 хуудас"), "it finishes");
  eq([half.getMaxColumns(), half.getRange(3, 5).getValue(), half.getRange(1, 6).getValue(), half.getRange(3, 6).getValue(), half.getRange(1, 43).getValue()], [47, "яг харагдах үнэ", 1, "Орлого ", "Нийт анхны үлдэгдэл ашиг"], "no column twice: days from F, totals from AQ");
}

// ---------- a year sheet with only a few priced products is moved too ----------
{
  const ss = new Spreadsheet();
  oldSheet(ss, "My/2026/10/15", "My home 2026 10 сарын 1-15", products(0));
  oldSheet(ss, "Sor/2026/10/15", "Соронзон 2026 10 сарын 1-15", products(0));
  const few = oldSheet(ss, "My/2026/12/31", "My home 2026 12 сарын 16-31", products(0).slice(0, 2));
  const run = makeRunner(ss, { triggers: [], mails: [] });
  const done = run(false, "runPairedMigration(ss, Date.now() + 60000)");
  ok(done.value[done.value.length - 1].startsWith("Бүх 3 хуудас") && few.getRange(3, 5).getValue() === "яг харагдах үнэ" && few.getMaxColumns() === 47, "a sheet the hourly job would not count yet still gets the new columns");
}

// ---------- a shop sheet's layout is read from its own header ----------
{
  const run = makeRunner(new Spreadsheet(), { triggers: [], mails: [] });
  const v = (e) => run(true, e).value;
  const row = ["Haier 1", 2, 100, 150, 160, 1, ...Array(14).fill(""), 0, 1, ...Array(14).fill(""), 0, 3, 2];
  const parsed = v(`parseSheetRows([${JSON.stringify(row)}], buildColumnLayout(15, { paired: true }))[0]`);
  eq([parsed.price, parsed.shelfPrice, parsed.income[0], parsed.expense[0], parsed.count, parsed.balance], [150, 160, 1, 1, 3, 2], "a paired row: price D, shelf price E, income from F, expense from V, count AL, balance AM");
  eq(v(`parseSheetRows([${JSON.stringify(row)}], buildColumnLayout(15))[0].shelfPrice`), "", "the old layout has no shelf price");
}

// ---------- counts between the period ends on a shop's second-half sheet (Sor/2026/10/31) ----------
{
  const run = makeRunnerFor(shipped)(new Spreadsheet(), { triggers: [], mails: [] }); const v = (e) => run(true, e).value;
  eq([v("spotFirstColumn(LAYOUT)"), v("viewBaseChoices(LAYOUT)").length], [49, 5], "the blocks start at AW after the totals AR..AV; a half-month sheet has no halves to choose");
  const flags = v(`viewColumnFlags('Тооллого 10/23', LAYOUT, 16, [{ first: 49, month: 10, day: 23 }], 51, true)`);
  const span = (a, b) => Array.from({ length: b - a + 1 }, (_, i) => a + i);
  eq(flags.map((f, i) => [i + 1, f]).filter(([, f]) => f === true).map(([c]) => c), [3, 4, 5, ...span(6, 13), 21, ...span(22, 29), 37, 38, 49, 50, 51], "\"Тооллого 10/23\" on the 16-31 sheet: days 16-23 (F..M, V..AC), the transfers and gifts, that day's block");
}

console.log(`ALL ${n} PAIRED CHECKS PASSED`);
