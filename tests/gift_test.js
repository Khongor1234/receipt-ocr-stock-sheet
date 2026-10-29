// Checks of the gift column (GIFT_COLUMN, config.gs): a "Бэлэг" column after the transfer out,
// taken off the final balance and, at cost, off the sales profit, and the one-time move of the
// paired sheets (gift-migration.gs). Usage: node tests/gift_test.js <project dir>
const fs = require("fs"), assert = require("assert");
const { Spreadsheet, makeRunnerFor } = require("./fake_book.js");

const dir = process.argv[2];
const FIRST_FILES = ["profile.gs", "config.gs"];
const gsFiles = fs.readdirSync(dir).filter((f) => f.endsWith(".gs"));
const loadOrder = [...FIRST_FILES.filter((f) => gsFiles.includes(f)), ...gsFiles.filter((f) => !FIRST_FILES.includes(f))];
const code = [...loadOrder, "profiles/sheet-rows.gs"].map((f) => fs.readFileSync(`${dir}/${f}`, "utf8")).join("\n");
assert(code.includes("const GIFT_COLUMN = true;"), "the main workbook ships with the gift column");
const makeRunner = makeRunnerFor(code);
const makePairedRunner = makeRunnerFor(code.replace("const GIFT_COLUMN = true;", "const GIFT_COLUMN = false;")); // the code as it was before
let n = 0;
const ok = (c, m) => { assert(c, m); n++; };
const eq = (a, b, m) => { assert.strictEqual(JSON.stringify(a), JSON.stringify(b), m); n++; };

// ---------- the layout and every constant that follows it ----------
{
  const run = makeRunner(new Spreadsheet(), { triggers: [], mails: [] });
  const v = (e) => run(true, e).value;
  eq(v("[LAYOUT.transferOut, LAYOUT.gift, LAYOUT.expenseLast, LAYOUT.count, LAYOUT.balance, LAYOUT.check, LAYOUT.otherBalance, LAYOUT.otherCount, LAYOUT.summaryFirst].map(columnLetter)"),
    ["AK", "AL", "AL", "AM", "AN", "AO", "AQ", "AP", "AR"], "AK transfer out, AL gifts, then count AM, balance AN, check AO, the other shop's count AP and balance AQ, totals from AR");
  eq(v("[CALC_FIRST_COL, SUMMARY_FIRST_COL, SETUP_ENTRY_LAST_COL, SYNC_COUNT_COL, SYNC_MOVEMENT_LAST_COL, INSTANT_COUNT_COLUMN, DIVIDER_BODY_LAST_COLUMN]"), [40, 44, 39, 39, 38, 39, 38], "every column constant follows the gift column; the gift column is typed like the days");
  eq(v("VIEW_EXPENSE_COLUMNS"), { first: 22, count: 17 }, "\"Зөвхөн орлого\" hides the expense days, the transfer out and the gifts");
  ok(v("SUMMARY_LINES[3].formula") === "=SUMPRODUCT(D4:D,BYROW(V4:AJ,LAMBDA(r,SUM(r))))", "Нийт борлуулалт: the expense days only (gifts are not sold)");
  ok(v("SUMMARY_LINES[4].formula") === "=SUMPRODUCT(D4:D,BYROW(V4:AJ,LAMBDA(r,SUM(r))))-SUMPRODUCT(C4:C,BYROW(V4:AJ,LAMBDA(r,SUM(r))))-SUMPRODUCT(C4:C,AL4:AL)", "Нийт борлуулалтын ашиг: minus the gifts at cost");
  ok(v("SUMMARY_LINES[2].formula") === "=SUMPRODUCT(D4:D,AM4:AM)", "Нийт эцсийн үлдэгдэл: the count AM");
  const f = v(`buildCalcFormulas(["эцсийн үлдэгдэл ", "үнэлгээ ", ...calcPairedHeaders("Соронзон")], "Sor/2026/10/15")`);
  ok(f[0].includes("BYROW(F4:U,") && f[0].includes("BYROW(V4:AL,"), "the balance: income minus the expense days, the transfer out AND the gifts");
  ok(f[1].includes("AM4:AM=AN4:AN") && f[2].includes("'Sor/2026/10/15'!AM4:AM") && f[3].includes("'Sor/2026/10/15'!AN4:AN"), "the check and the other shop's columns read the new count AM and balance AN (the count first)");
  ok(v("buildStockFormulas(CHAINS[0]).list").includes('"AK4:AL"') && v("buildStockFormulas(CHAINS[0]).list").includes("BYROW(IF(ISNUMBER(tout),tout,0),LAMBDA(r,SUM(r)))"), "today's stock takes off the transfer out and the gifts");
  eq(v("[1, 37, 38, 42, 47].map(giftColumnOf)"), [1, 37, 39, 43, 48], "a paired column's place after the gift column");
  eq(v(`moveGiftColors({ "haier 1": { "8": "#f00", "38": "#0f0" } })`), { "haier 1": { "8": "#f00", "39": "#0f0" } }, "the owner's colours move with their columns");
  eq(v(`[shopLayoutOf("яг харагдах үнэ", 15, "Бэлэг").count, shopLayoutOf("яг харагдах үнэ", 15, "тооллого").count, shopLayoutOf("Орлого", 15, "Бэлэг").count, giftHeaderColumn(15)]`), [39, 38, 37, 38], "a sheet's own row 3 tells whether it has the gift column (AL3)");
}

// ---------- the move: a workbook of the paired layout ----------
const P = { shelf: 5, income: 6, expense: 22, outT: 37, count: 38 };
function pairedSheet(ss, name, title, rows) {
  const s = ss.insertSheet(name);
  s.maxCols = 47;
  s.getRange(1, P.income).setValue(1);
  s.getRange(2, 1).setValue(title);
  [["Барааны нэрс", 1], ["анхны үлдэгдэл ", 2], ["өртөг үнэ", 3], ["зарах үнэ", 4], ["яг харагдах үнэ", P.shelf], ["Орлого ", P.income], ["Зарлага", P.expense], ["Шилжүүлэг", P.outT], ["тооллого", P.count]].forEach(([text, c]) => s.getRange(3, c).setValue(text));
  rows.forEach(([label, opening, expense, count], i) => {
    const r = 4 + i;
    s.getRange(r, 1).setValue(label); s.getRange(r, 2).setValue(opening); s.getRange(r, 3).setValue(100); s.getRange(r, 4).setValue(150);
    if (expense) s.getRange(r, P.expense).setValue(expense);
    if (count !== null) s.getRange(r, P.count).setValue(count);
  });
  return s;
}
const products = () => Array.from({ length: 22 }, (_, i) => [`Haier ${i + 1}`, 2, i === 0 ? 1 : 0, i === 0 ? 1 : null]);
{
  const ss = new Spreadsheet();
  const my = pairedSheet(ss, "My/2026/10/15", "My home 2026 10 сарын 1-15", products());
  const sor = pairedSheet(ss, "Sor/2026/10/15", "Соронзон 2026 10 сарын 1-15", products());
  const state = { triggers: [], mails: [], docProps: { pairedLayoutReady: "1", [`approvedColors:${my.getSheetId()}`]: JSON.stringify({ "haier 1": { "38": "#ffff00" } }), [`header:${my.getSheetId()}`]: "[]" } };
  // the paired sheets get their formulas and totals from the code as it was (no gift column)
  const paired = makePairedRunner(ss, state);
  [my, sor].forEach((sheet) => paired(false, `finishPairedSheet(ss.getSheetByName("${sheet.getName()}"))`));
  ok(my.getRange(3, 39).getFormula().includes("BYROW(V4:AK,"), "the start: a paired sheet, balance in AM");
  const run = makeRunner(ss, state);

  const before = run(false, "(runBackgroundJobs(), isPairedLayoutReady())");
  ok(before.value === false && before.logs.includes("migrateToGiftColumn"), "until the gift column is in, the hourly job does nothing and says why");

  const dry = run(true, "runGiftMigration(ss, Date.now() + 60000)");
  ok(dry.value.length === 3 && dry.value[0].includes("AK-ийн дараа \"Бэлэг\"") && my.getMaxColumns() === 47, "a dry run only reports");

  const real = run(false, "runGiftMigration(ss, Date.now() + 60000)");
  ok(real.value[real.value.length - 1].startsWith("Бүх 2 хуудсанд"), "the move finishes");
  eq(state.docProps.giftColumnReady, "1", "and marks the workbook ready");
  ok(my.getMaxColumns() === 48 && sor.getMaxColumns() === 48, "one column more");
  eq([my.getRange(3, 37).getValue(), my.getRange(3, 38).getValue(), my.getRange(3, 39).getValue()], ["Шилжүүлэг", "Бэлэг", "тооллого"], "AL3 says Бэлэг, the count moved to AM");
  eq([my.getRange(4, 22).getValue(), my.getRange(4, 38).getValue(), my.getRange(4, 39).getValue()], [1, "", 1], "the numbers after the transfer out moved one right; the gift column is empty for mom");
  const header = my.getRange(3, 40, 1, 4).getFormulas()[0];
  ok(header[0].includes("BYROW(V4:AL,") && header[1].includes("AM4:AM=AN4:AN"), "the balance (with the gifts) and the check sit in AN, AO");
  ok(header[2].includes("'Sor/2026/10/15'!AM4:AM") && header[3].includes("'Sor/2026/10/15'!AN4:AN"), "the other shop's columns AP, AQ read Соронзон's new count and balance (the count first, 2026-10-28)");
  ok(my.getRange(3, 48).getFormula().endsWith("-SUMPRODUCT(C4:C,AL4:AL)"), "the sales profit in AV takes the gifts off at cost");
  eq(my.getRange(1, 44, 1, 5).getValues()[0], ["Нийт анхны үлдэгдэл ашиг", "Нийт анхны үлдэгдэл", "Нийт эцсийн үлдэгдэл", "Нийт борлуулалт", "Нийт борлуулалтын ашиг"], "the totals' labels sit in AR..AV");
  eq(JSON.parse(state.docProps[`approvedColors:${my.getSheetId()}`]), { "haier 1": { "39": "#ffff00" } }, "the owner's colour of the count moved with it");
  ok(Array.isArray(JSON.parse(state.docProps[`header:${my.getSheetId()}`])), "the header is remembered again from the new columns");
  ok(state.docProps[`giftFinished:${my.getSheetId()}`] === "1" && state.docProps[`giftFinished:${sor.getSheetId()}`] === "1", "each finished sheet is marked");
  const again = run(false, "runGiftMigration(ss, Date.now() + 60000)");
  ok(my.getMaxColumns() === 48 && again.value[again.value.length - 1].startsWith("Бүх 2 хуудсанд"), "a second run inserts nothing more");
  ok(run(false, "isPairedLayoutReady()").value === true, "from now on the scripts work again");
}

// ---------- not before the paired move; a run cut off by the time limit ----------
{
  const ss = new Spreadsheet();
  pairedSheet(ss, "My/2026/10/15", "My home 2026 10 сарын 1-15", products());
  pairedSheet(ss, "Sor/2026/10/15", "Соронзон 2026 10 сарын 1-15", products());
  const state = { triggers: [], mails: [] };
  const run = makeRunner(ss, state);
  ok(run(false, "runGiftMigration(ss, Date.now() + 60000)").value[0].startsWith("ЗОГСЛОО"), "the gift column waits for the paired move");
  state.docProps = { pairedLayoutReady: "1" };
  const cut = run(false, "runGiftMigration(ss, Date.now() - 1)");
  ok(cut.value[cut.value.length - 1].includes("дахин ажиллуулна уу") && state.docProps.giftColumnReady === undefined, "out of time: nothing is marked ready");
  const rest = run(false, "runGiftMigration(ss, Date.now() + 60000)");
  ok(rest.value[rest.value.length - 1].startsWith("Бүх 2 хуудсанд") && ss.getSheetByName("Sor/2026/10/15").getMaxColumns() === 48, "the next run finishes it");
}

// ---------- a row read by the site and the OCR ----------
{
  const run = makeRunner(new Spreadsheet(), { triggers: [], mails: [] });
  const row = ["Haier 1", 2, 100, 150, 160, 1, ...Array(14).fill(""), "", 1, ...Array(14).fill(""), 0, 1, 3, 2];
  const parsed = run(true, `parseSheetRows([${JSON.stringify(row)}], shopLayoutOf("яг харагдах үнэ", 15, "Бэлэг"))[0]`).value;
  eq([parsed.expense[0], parsed.count, parsed.balance], [1, 3, 2], "a row with the gift column: count AM, balance AN");
}

// ---------- the gift column's own colour (the owner's choice, 2026-10-21) ----------
{
  const ss = new Spreadsheet();
  const sheet = ss.insertSheet("Sor/2026/10/31");
  sheet.getRange(3, 5).setValue("яг харагдах үнэ"); sheet.getRange(3, 38).setValue("Бэлэг");
  for (let r = 1; r <= 3; r++) { sheet.cell(r, 37).bg = "#fff2cc"; sheet.cell(r, 38).bg = "#fff2cc"; }
  sheet.getRange(4, 1).setValue("Хөргөгч"); for (let c = 1; c <= 45; c++) sheet.cell(4, c).bg = "#f6b26b";
  for (let i = 0; i < 22; i++) { sheet.getRange(5 + i, 1).setValue(`Haier ${i}`); sheet.getRange(5 + i, 2, 1, 3).setValues([[1, 100, 150]]); sheet.cell(5 + i, 37).bg = "#fce5cd"; sheet.cell(5 + i, 38).bg = "#fce5cd"; }
  sheet.cell(6, 38).bg = "#ffe599"; // a gift the OCR wrote
  const run = makeRunner(ss, { triggers: [], mails: [], docProps: { pairedLayoutReady: "1", giftColumnReady: "1" } });
  const dry = run(true, "runTintGiftColumn(ss)").value;
  ok(dry.some((l) => l.includes("AL 24 нүд будагдана")) && sheet.cell(5, 38).bg === "#fce5cd", "the trial run counts the cells and paints nothing");
  run(false, "runTintGiftColumn(ss)");
  eq([sheet.cell(2, 38).bg, sheet.cell(5, 38).bg, sheet.cell(26, 38).bg, sheet.cell(5, 37).bg], ["#c9e4c1", "#e2efda", "#e2efda", "#fce5cd"], "AL gets its own tint, AK keeps its colour");
  eq([sheet.cell(4, 38).bg, sheet.cell(6, 38).bg], ["#f6b26b", "#ffe599"], "the category row and the OCR's cell keep theirs");
  eq(run(true, "findColumnColors(ss.getSheetByName('Sor/2026/10/31').getRange(5, 1, 22, 40).getBackgrounds(), Array.from({ length: 22 }, (_, i) => i))[37]").value, "#e2efda", "dividers.gs takes the tint as the column's own colour, so it stays");
}

console.log(`ALL ${n} GIFT CHECKS PASSED`);
