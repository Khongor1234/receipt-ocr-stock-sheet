// Fake in-memory spreadsheet + checks for the whole script set. Usage: node setup_test.js <project dir>
const fs = require("fs"), vm = require("vm"), assert = require("assert");
const { WHITE, cell, Protection, Range, Sheet, Spreadsheet: BaseSpreadsheet, makeRunnerFor } = require("./fake_book.js");
const dim = (m) => new Date(Date.UTC(2026, m, 0)).getUTCDate();
// New tab name of the sheet that used to be called "Sheet<n>".
const N = (n) => { const [prefix, first] = n >= 31 ? ["Sor", 31] : ["My", 1]; const pos = n - first; const m = Math.floor(pos / 2) + 1; return `${prefix}/2026/${m}/${pos % 2 ? dim(m) : 15}`; };
const isYearNumber = (n) => (n >= 1 && n <= 24) || (n >= 31 && n <= 54);
class Spreadsheet extends BaseSpreadsheet {
  // Test convenience only: "Sheet19" also finds the sheet once it is renamed to "My/2026/10/15".
  getSheets() { return this.sheets.slice(); } getSheetByName(n) { const m = /^Sheet(\d+)$/.exec(n); return this.sheets.find(s => s.name === n) || (m && isYearNumber(Number(m[1])) && this.sheets.find(s => s.name === N(Number(m[1])))) || null; }
  insertSheet(n) { const s = new Sheet(this, n); this.sheets.push(s); return s; }
  deleteSheet(s) { this.sheets = this.sheets.filter(x => x !== s); }
  setActiveSheet(s) { this.active = s; } getActiveSheet() { return this.active; } moveActiveSheet(pos) { this.sheets = this.sheets.filter(x => x !== this.active); this.sheets.splice(pos - 1, 0, this.active); }
  getOwner() { return { getEmail: () => "owner@example.com" }; }
  getSpreadsheetTimeZone() { return "Asia/Ulaanbaatar"; } getName() { return "test"; } getUrl() { return "url"; }
  logical() { return this.sheets.map(s => s.logical()).join("\n"); }
}

const band = (s, row, color) => { for (let c = 1; c <= 35; c++) s.cell(row, c).bg = color; };
function makeSheet(ss, name, title, summaryRow) {
  const s = ss.insertSheet(name); s.getRange(2, 1).setValue(title);
  const put = (label, row) => { s.getRange(row, 1).setValue(label); [5, 100, 150].forEach((v, i) => s.getRange(row, 2 + i).setValue(v)); s.getRange(row, 7).setValue(3); s.getRange(row, 37).setValue(8); };
  s.getRange(4, 1).setValue("Cat A"); band(s, 4, "#f6b26b");
  for (let i = 1; i <= 12; i++) put(`A ${i}`, 4 + i);
  s.getRange(17, 1).setValue("Cat B"); band(s, 17, "#0b5394");
  for (let i = 1; i <= 12; i++) put(`B ${i}`, 17 + i);
  if (summaryRow) {
    ["Нийт анхны үлдэгдэл ашиг", "Нийт анхны үлдэгдэл ", "Нийт эцсийн үлдэгдэл   ", "Нийт борлуулалт ", "Нийт борлуулалтын ашиг"].forEach((label, i) => { s.getRange(summaryRow + i, 41).setValue(label); s.cell(summaryRow + i, 42).v = 1000 + i; s.cell(summaryRow + i, 42).f = "=OLD"; s.cell(summaryRow + i, 41).border = true; s.cell(summaryRow + i, 42).border = true; });
    for (let c = 41; c <= 45; c++) { s.cell(8, c).bg = "#ea9999"; s.cell(9, c).bg = "#ea9999"; }
  }
  return s;
}
function buildBook() {
  const ss = new Spreadsheet();
  makeSheet(ss, "Sheet100", "Соронзон 2026 9 сарын 30", 6); makeSheet(ss, "Sheet56", "My home 2026 9 сарын 30", 6);
  const s200 = ss.insertSheet("Sheet200");
  for (let r = 6; r <= 9; r++) for (let c = 3; c <= 14; c++) s200.cell(r, c).f = `=SUM(Sheet${c}!$AP$8,Sheet${c}!$AP$8)`;
  s200.cell(10, 3).f = "=sum(C6+C8)";
  const names = { 9: "Sheet 9", 13: "Sheet 13", 39: "Sheet 39", 41: "Seet 41" };
  for (let i = 1; i <= 18; i++) { makeSheet(ss, names[i] || `Sheet${i}`, `My home 2026 x ${i}`, i % 2 ? 5 : 6); makeSheet(ss, names[30 + i] || `Sheet${30 + i}`, `Соронзон 2026 x ${i}`, i === 18 ? 0 : 6); }
  ss.insertSheet("Sheet55"); ss.insertSheet("Sheet54");
  return ss;
}
const dir = process.argv[2];
// profile.gs and config.gs first, like the build: the other files read LAYOUT while they load.
const FIRST_FILES = ["profile.gs", "config.gs"];
const gsFiles = fs.readdirSync(dir).filter(f => f.endsWith(".gs"));
const loadOrder = [...FIRST_FILES.filter(f => gsFiles.includes(f)), ...gsFiles.filter(f => !FIRST_FILES.includes(f))];
const shippedCode = loadOrder.map(f => fs.readFileSync(`${dir}/${f}`, "utf8")).join("\n");
assert(shippedCode.includes("const DRY_RUN = true;"), "DRY_RUN must ship as true");
// These checks are written for the single-block layout (A..D, days from E, "sold" in AN); every
// script reads its columns from LAYOUT, so they run with PAIRED_SHOPS = false. The two shops'
// layout (shelf price, other shop's columns) has its own checks: tests/paired_test.js.
assert(shippedCode.includes("const PAIRED_SHOPS = true;"), "the main workbook ships with the paired layout");
// The fake shops' sheets have the real names: the real list of uncounted sheets (profile.gs) is
// left out, a check below puts one in.
assert(/const UNCOUNTED_SHEETS = \[[^\]]*\];/.test(shippedCode), "the main workbook lists its uncounted sheets");
const code = shippedCode.replace("const PAIRED_SHOPS = true;", "const PAIRED_SHOPS = false;").replace(/const UNCOUNTED_SHEETS = \[[^\]]*\];/, "const UNCOUNTED_SHEETS = [];");
const makeRunner = makeRunnerFor(code);
let n = 0; const ok = (c, m) => { assert(c, m); n++; };
const eq = (a, b, m) => { assert.strictEqual(JSON.stringify(a), JSON.stringify(b), m); n++; };

// ---------- pure helpers ----------
{
  const run = makeRunner(new Spreadsheet(), { triggers: [], mails: [] });
  const v = (e) => run(true, e).value;
  eq(v(`["Sheet 9","Seet 41","Sheet13","sheet 40","Sheet200","_sync","Тайлан"].map(canonicalSheetName)`), ["My/2026/5/15", "Sor/2026/6/15", "My/2026/7/15", "Sor/2026/5/31", "Sheet200", "_sync", "Тайлан"], "old names map to shop/year/month/last day");
  eq(v(`sheetSlot("My/2026/10/15")`), { chainKey: "my home", position: 18 }, "My/2026/10/15 = October 1-15");
  eq(v(`sheetSlot("Sor/2026/2/28")`), { chainKey: "соронзон", position: 3 }, "Sor/2026/2/28 = February 16-28");
  eq(v(`sheetSlot(" sor/2026/12/31 ")`), { chainKey: "соронзон", position: 23 }, "case and spaces do not matter");
  eq(v(`["My/2025/1/15","My/2026/13/15","Xx/2026/1/15","My/2026/1","Sheet200"].map(sheetSlot)`), [null, null, null, null, null], "other year, bad month, unknown shop are not year sheets");
  eq(v(`["my/2026/2/29","Sor/2026/4/16","My/2026/1/15"].map(canonicalSheetName)`), ["My/2026/2/28", "Sor/2026/4/30", "My/2026/1/15"], "a slightly wrong new name is corrected to the right last day");
  eq(v(`[0,1,3,19,23].map(p => chainSheetName(CHAINS[1], p))`), ["Sor/2026/1/15", "Sor/2026/1/31", "Sor/2026/2/28", "Sor/2026/10/31", "Sor/2026/12/31"], "names follow the calendar");
  eq(v(`sheetSlot("Sheet1")`), { chainKey: "my home", position: 0 }, "Sheet1 slot");
  eq(v(`sheetSlot("Sheet 24")`), { chainKey: "my home", position: 23 }, "Sheet24 slot");
  eq(v(`sheetSlot("Seet 41")`), { chainKey: "соронзон", position: 10 }, "Seet 41 slot");
  eq(v(`sheetSlot("Sheet54")`), { chainKey: "соронзон", position: 23 }, "Sheet54 slot");
  eq(v(`["Sheet25","Sheet30","Sheet55","Sheet56","Sheet100","Sheet200","Sheet0"].map(sheetSlot)`), [null, null, null, null, null, null, null], "non-chain sheets have no slot");
  const f = v("buildSheet200Formulas()");
  ok(f.length === 4 && f.every(r => r.length === 12), "Sheet200 block is 4 x 12");
  ok(f[0][0] === "=SUM('My/2026/1/15'!$AR$3,'My/2026/1/31'!$AR$3)" && f[1][0] === "=SUM('My/2026/1/15'!$AS$3,'My/2026/1/31'!$AS$3)", "My home January, names quoted");
  ok(f[0][11] === "=SUM('My/2026/12/15'!$AR$3,'My/2026/12/31'!$AR$3)" && f[2][0] === "=SUM('Sor/2026/1/15'!$AR$3,'Sor/2026/1/31'!$AR$3)" && f[3][11] === "=SUM('Sor/2026/12/15'!$AS$3,'Sor/2026/12/31'!$AS$3)", "December and Соронзон");
  ok(v("expectedSheetNames()").length === 53 && v("expectedSheetNames()").includes("Sheet100*") && v("expectedSheetNames()").includes("My home тооллого"), "48 year sheets + Sheet200 + both stock sheets + both count sheets expected");
  eq(v(`findMissingSheets(["Sheet 1","My/2026/1/31","Sheet200"],["My/2026/1/15","My/2026/1/31","My/2026/2/15","Sheet200"])`), ["My/2026/2/15"], "missing sheet found; an old or misspelt name counts as present");
  eq(v(`describePeriod(19,2026)`), { year: 2026, month: 10, isSecondHalf: true, firstDay: 16, lastDay: 31, monthDays: 31 }, "position 19 = October 16-31");
  eq(v(`buildDayHeaders(describePeriod(0,2026))`), [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15], "January 1-15 day numbers");
  eq(v(`buildDayHeaders(describePeriod(3,2026)).slice(11)`), [27, 28, "", ""], "February has no 29, 30");
  ok(v(`buildTitle("My home",describePeriod(0,2026))`) === "My home 2026 1 сарын 1-15", "title");
}

// ---------- dividers written as a coloured name only (My home) ----------
{
  const run = makeRunner(new Spreadsheet(), { triggers: [], mails: [] }); const v = (e) => run(true, e).value;
  const aOnly = JSON.stringify(["#e06666", ...Array(34).fill("#ffffff")]); const blank = (label, extra = {}) => JSON.stringify(Array.from({ length: 37 }, (_, i) => i === 0 ? label : (extra[i] ?? "")));
  ok(v(`isDividerRow(${aOnly}, ${blank("Хөргөгч")})`) === true, "coloured known category name with an empty row is a divider");
  ok(v(`isDividerRow(${aOnly}, ${blank("  бусад  (Жижиг бараа) ")})`) === true, "spacing and case of the name do not matter");
  ok(v(`isDividerRow(${aOnly}, ${blank("Frestech 200l")})`) === false, "coloured PRODUCT with no numbers yet is not a divider (must stay editable)");
  ok(v(`isDividerRow(${aOnly}, ${blank("Хөргөгч", { 3: 500000 })})`) === false, "a row with a price is a product");
  ok(v(`isDividerRow(${JSON.stringify(Array(35).fill("#ffffff"))}, ${blank("Хөргөгч")})`) === false, "uncoloured name is not a divider");
  ok(v(`isDividerRow(${aOnly})`) === false && v(`isDividerRow(${JSON.stringify(Array(35).fill("#0b5394"))})`) === true, "without values only the full band counts");
  const names = ["Sheet100", "Sheet200", "Sheet1", "Sheet31", "Sheet2", "Sheet32", "Sheet49", "Sheet18", "Sheet19", "Sheet55", "_sync"];
  eq(v(`desiredTabOrder(${JSON.stringify(names)})`), ["Sheet100", "Sheet200", "Sheet31", "Sheet1", "Sheet32", "Sheet2", "Sheet18", "Sheet49", "Sheet19", "Sheet55", "_sync"], "pairs: Соронзон then My home for every period; other sheets keep their side");
  const blocks = (c) => c.replace("const TAB_ORDER_MODE = 'pairs';", "const TAB_ORDER_MODE = 'blocks';");
  eq(run(true, `desiredTabOrder(${JSON.stringify(names)})`, blocks).value, ["Sheet100", "Sheet200", "Sheet31", "Sheet32", "Sheet49", "Sheet1", "Sheet2", "Sheet18", "Sheet19", "Sheet55", "_sync"], "blocks: all Соронзон, then all My home");
  eq(v(`planTabMoves(desiredTabOrder(${JSON.stringify(names)}))`), [], "nothing to move when the order is right");
}

// ---------- a colour inherited from a divider does not make a row a divider ----------
{
  const run = makeRunner(new Spreadsheet(), { triggers: [], mails: [] }); const v = (e) => run(true, e).value;
  const bandBg = (c) => JSON.stringify(Array(35).fill(c)); const vals = (label, extra = {}) => JSON.stringify(Array.from({ length: 37 }, (_, i) => i === 0 ? label : (extra[i] ?? "")));
  ok(v(`isDividerRow(${bandBg("#00ff00")}, ${vals("Haier 5105", { 1: 2, 2: 100, 3: 150 })})`) === false, "a fully coloured row WITH numbers is a product (it inherited the colour)");
  ok(v(`isDividerRow(${bandBg("#00ff00")}, ${vals("Зурагт")})`) === true && v(`isDividerRow(${bandBg("#00ff00")}, ${vals("Миний ангилал")})`) === true, "a coloured row with a name and no numbers is a divider");
  ok(v(`isDividerRow(${bandBg("#00ff00")}, ${vals("")})`) === false, "a coloured row with nothing in it is not a divider");
  const info = (row, color, label, opts = {}) => `describeRow(${row}, ${opts.aOnly ? JSON.stringify([color, ...Array(34).fill("#ffffff")]) : bandBg(color)}, ${vals(label, opts.nums ? { 2: 100, 3: 150 } : {})})`;
  const classify = (rows, known = []) => { const r = v(`classifyRows([${rows.join(",")}], ${JSON.stringify(known)})`); return { dividers: r.dividers.map(x => x.row), inherited: r.inherited.map(x => x.row) }; };
  eq(classify([info(10, "#0f0", "Haier 5105", { nums: true }), info(11, "#0f0", "чемодан", { nums: true }), info(12, "#0f0", "Хятад жижиг бараа"), info(13, "#ffffff", "x", { nums: true })]), { dividers: [12], inherited: [10, 11] }, "coloured products above a divider: colour is inherited, only the category is a divider");
  eq(classify([info(20, "#0f0", "Taube"), info(21, "#0f0", "хөлдөөгч")]), { dividers: [21], inherited: [20] }, "a product name without numbers yet, touching a named category of the same colour, is not a divider");
  eq(classify([info(30, "#0f0", "Хөргөгч"), info(31, "#0f0", "Хөлдөөгч")]), { dividers: [30, 31], inherited: [] }, "two named categories next to each other are both dividers");
  eq(classify([info(40, "#0f0", "Миний ангилал"), info(41, "#ffffff", "x", { nums: true })]), { dividers: [40], inherited: [] }, "an unlisted name on a band that touches no category is a divider");
  eq(classify([info(50, "#0f0", "Миний ангилал"), info(51, "#0f0", "")], [50]), { dividers: [50], inherited: [51] }, "an empty row inserted next to an already protected divider is not a divider");
  eq(classify([info(60, "#0f0", "Миний ангилал"), info(61, "#00f", "Өөр ангилал")], [60]), { dividers: [60, 61], inherited: [] }, "a named neighbour with a different colour is a new divider");
  eq(classify([info(70, "#0f0", "Aucma 441l", { aOnly: true, nums: true }), info(71, "#0f0", "Хөлдөөгч", { aOnly: true })]), { dividers: [71], inherited: [70] }, "the screenshot case: green product cell above the green category cell");
  eq(classify([info(80, "#f00", "онцлох бараа", { aOnly: true, nums: true }), info(81, "#0f0", "Хөлдөөгч", { aOnly: true })]), { dividers: [81], inherited: [] }, "a product highlighted in another colour keeps its colour");
  eq(classify([info(107, "#ffffff", "a", { nums: true }), info(108, "#fce", "b", { nums: true }), info(109, "#fce", "c", { nums: true }), info(110, "#fce", ""), info(111, "#ffffff", "d", { nums: true })], [110]), { dividers: [], inherited: [108, 109, 110] }, "the second screenshot: products coloured across the row and the empty coloured row under them all lose the colour; nothing is a divider");
  eq(classify([info(120, "#ffffff", "a", { nums: true }), info(121, "#fce", ""), info(122, "#ffffff", "b", { nums: true })]), { dividers: [121], inherited: [] }, "an empty coloured row standing on its own is a divider (the unnamed separator of the Соронзон sheets)");
  eq(classify([info(150, "#00f", "будсан бараа", { nums: true }), info(151, "#00f", "Жагсаалтад байхгүй ангилал")], [151]), { dividers: [151], inherited: [150] }, "a named divider keeps its place even when a coloured product sits right beside it");
  eq(classify([info(125, "#0f0", "хөлдөөгч"), info(126, "#ffffff", "x", { nums: true }), info(127, "#b45", ""), info(128, "#ffffff", "y", { nums: true }), info(129, "#0ff", "Хөргүүр")]), { dividers: [125, 127, 129], inherited: [] }, "the real layout: category, products, unnamed brown separator, products, category");
  eq(classify([info(130, "#fce", "онцолсон", { nums: true }), info(131, "#ffffff", "b", { nums: true })]), { dividers: [], inherited: [130] }, "a product coloured across the whole row loses the colour even with no divider near");
  eq(classify([info(140, "#fce", "онцолсон", { aOnly: true, nums: true }), info(141, "#ffffff", "b", { nums: true })]), { dividers: [], inherited: [] }, "a product highlighted in column A only keeps its highlight");
}

// ---------- dry run changes nothing ----------
const ss = buildBook(); const state = { triggers: [], mails: [] }; const run = makeRunner(ss, state);
{
  const before = ss.logical();
  const r = run(true, "setupYear()");
  ok(ss.logical() === before && ss.sheets.length === 41, "dry run: no cell, sheet or protection changed");
  ok(state.triggers.length === 0, "dry run: no trigger installed");
  ok(r.logs.includes('"Sheet 9" -> "My/2026/5/15"') && r.logs.includes('"Seet 41" -> "Sor/2026/6/15"') && r.logs.includes('"Sheet1" -> "My/2026/1/15"'), "dry run reports renames");
  ok(r.logs.includes("My/2026/10/15: Sheet18-ээс хуулж үүсгэнэ") && r.logs.includes("Sor/2026/12/31: "), "dry run reports the sheets it would create");
  ok(r.logs.includes("ТУРШИЛТ"), "dry run says so");
}

// ---------- real run ----------
const first = run(false, "setupYear()");
ok(first.logs.includes("БҮГД ДУУСЛАА"), "real run finishes:\n" + first.logs.slice(-600));
const names = ss.sheets.map(s => s.name);
ok([1, 9, 13, 39, 41, 48].map(N).every(x => names.includes(x)) && !names.some(x => /^(Sheet |Seet)/.test(x)) && !names.some(x => { const m = /^Sheet(\d+)$/.exec(x); return m && isYearNumber(Number(m[1])) && x !== "Sheet54"; }), "every year sheet is renamed to shop/year/month/last day");
for (const [from, label] of [[1, "My home"], [31, "Соронзон"]]) {
  ok(Array.from({ length: 24 }, (_, i) => N(from + i)).every(x => names.includes(x)), `${label}: 24 sheets exist`);
}
{
  const year = names.filter(x => /^(My|Sor)\/2026\//.test(x));
  const expected = Array.from({ length: 24 }, (_, i) => [N(31 + i), N(1 + i)]).flat();
  eq(expected.slice(0, 4).concat(expected.slice(-2)), ["Sor/2026/1/15", "My/2026/1/15", "Sor/2026/1/31", "My/2026/1/31", "Sor/2026/12/31", "My/2026/12/31"], "(the expected order, spelled out)");
  eq(year, expected, "tabs: Соронзон then My home, period after period, January to December");
  eq(names.slice(0, 3), ["Sheet100", "Sheet56", "Sheet200"], "sheets in front of the year sheets stay in front");
  ok(names.indexOf("Sheet55") > names.indexOf(N(24)), "other sheets go after the year sheets");
}
ok(names.includes("Sheet54") && names.includes("Sheet55") && ss.sheets.filter(s => !s.name.startsWith("_")).length === 41 + 12 + 3 && names.includes("Sheet100*") && names.includes("Соронзон тооллого"), "12 sheets added + Sheet100* + the two count sheets; the empty Sheet54 / Sheet55 are left alone");
const s19 = ss.getSheetByName("Sheet19"), s18 = ss.getSheetByName("Sheet18"), s24 = ss.getSheetByName("Sheet24"), s54 = ss.getSheetByName(N(54));
ok(s19.cell(5, 1).v === "A 1" && s19.cell(5, 3).v === 100 && s19.cell(5, 4).v === 150, "new sheet keeps names, cost, price");
ok(s19.cell(5, 7).v === "" && s19.cell(5, 37).v === "", "new sheet: movements and count emptied");
ok(s19.cell(5, 2).v === 8, "Sheet19 opening balance = Sheet18 count (carried)");
ok(s24.cell(5, 2).v === "" && s24.cell(5, 37).v === "", "later new sheets wait for a count");
ok(s18.cell(5, 7).v === 3 && s18.cell(5, 37).v === 8 && s18.cell(5, 2).v === 8 && ss.getSheetByName("Sheet1").cell(5, 2).v === 5, "source sheet: movements and count untouched, opening balance = the count before it (the first sheet keeps its own)");
ok(s19.cell(4, 1).bg === "#f6b26b" && s19.protections.filter(p => p.desc.startsWith("Ангиллын")).length === 2, "new sheet has both dividers, protected");
const yearSheets = ss.sheets.filter(s => /^(My|Sor)\/2026\//.test(s.name));
ok(yearSheets.length === 48, "48 year sheets");
{
  const head = (s) => s.protections.filter(p => p.desc.startsWith("Гарчгийн мөр"));
  const covers = (s, r, c) => head(s).some(p => r >= p.range.r && r < p.range.r + p.range.h && c >= p.range.c && c < p.range.c + p.range.w && p.warning === false && p.editors.length === 0);
  const cells = (s) => [1, 2, 3].flatMap(r => Array.from({ length: s.maxCols }, (_, i) => [r, i + 1]));
  ok(yearSheets.every(s => head(s).length === 2 && cells(s).every(([r, c]) => (r === 1 && c <= 2) ? !covers(s, r, c) : covers(s, r, c))), "rows 1-3 hard-locked full width on every year sheet, except A1 (view list) and B1 (seller) (after AS was added)");
  const s200p = ss.getSheetByName("Sheet200"); ok(head(s200p).length === 1 && cells(s200p).every(([r, c]) => covers(s200p, r, c)), "a sheet that is not a year sheet keeps its whole header locked");
  eq(run(true, "headerProtectionRanges(true, 45)").value, [{ row: 1, column: 3, rows: 1, columns: 43 }, { row: 2, column: 1, rows: 2, columns: 45 }], "the header is protected around the open cells");
}
ok(yearSheets.every(s => s.maxCols >= 45 && s.cell(1, 44).v === "Нийт борлуулалт" && s.cell(3, 44).f === "=SUMPRODUCT(D4:D,AN4:AN)" && s.cell(3, 45).f.startsWith("=SUMPRODUCT(D4:D,AN4:AN)-")), "totals in AO1:AS3 on every year sheet (also where they were missing: Sheet48)");
ok(yearSheets.every(s => { for (let r = 4; r <= 30; r++) if (s.cell(r, 41).v !== "" || s.cell(r, 42).f !== "") return false; return true; }), "old AO:AP block removed from product rows");
ok(yearSheets.every(s => [41, 42, 45].every(c => [5, 6, 8, 9, 10, 300].every(r => s.cell(r, c).bg === "#ffffff" && !s.cell(r, c).border))), "left-over boxes and colours under the totals are gone");
ok(yearSheets.every(s => s.cell(4, 41).bg === "#f6b26b" && s.cell(4, 45).bg === "#f6b26b" && s.cell(17, 43).bg === "#0b5394"), "divider bands run on across the totals columns");
ok(yearSheets.every(s => s.cell(1, 41).bg === "#9fc5e8" && s.cell(3, 41).bg === "#deebf7" && s.cell(1, 44).bg === "#93c47d" && s.cell(3, 44).bg === "#d9ead3" && s.cell(1, 45).bg === "#ffd966" && s.cell(3, 45).bold && s.cell(1, 45).border), "each total has its own colour: label strong, value light, bold, boxed");
ok(new Set([41, 42, 43, 44, 45].map(c => ss.getSheetByName("Sheet1").cell(1, c).bg)).size === 5, "five different label colours");
ok(yearSheets.every(s => s.cell(5, 38).bg === "#ffffff" && s.cell(5, 1).bg === "#ffffff" && s.cell(2, 41).bg === "#ffffff"), "product columns and row 2 are not recoloured");
ok(yearSheets.every(s => [41, 42, 43, 44, 45].every(col => s.cell(3, col).nf === '#,##0 "₮"')), "totals shown as tugrik on every year sheet");
ok(yearSheets.every(s => s.cell(5, 3).nf === '#,##0 "₮"' && s.cell(s.maxRows, 4).nf === '#,##0 "₮"' && s.cell(5, 2).nf === "" && s.cell(5, 7).nf === "" && s.cell(5, 37).nf === ""), "cost and price are tugrik, quantities are not");
ok(ss.getSheetByName("Sheet200").cell(6, 3).nf === '#,##0 "₮"' && ss.getSheetByName("Sheet200").cell(11, 14).nf === '#,##0 "₮"' && ss.getSheetByName("Sheet200").cell(12, 3).nf === "0%", "Sheet200 rows 6-11 are tugrik, the profit-share row is a percentage");
ok(ss.getSheetByName("Sheet56").cell(5, 4).nf === "", "non-year sheets are not reformatted");
ok(yearSheets.every(s => s.cell(3, 38).f.startsWith("=ARRAYFORMULA(") && s.cell(3, 40).f.includes("BYROW(U4:AI")), "AL:AN array formulas everywhere");
ok(ss.getSheetByName("Sheet1").cell(2, 1).v === "My home 2026 1 сарын 1-15" && ss.getSheetByName("Sheet1").cell(1, 5).v === 1 && ss.getSheetByName("Sheet1").cell(1, 19).v === 15 && ss.getSheetByName("Sheet1").cell(1, 21).v === 1 && ss.getSheetByName("Sheet1").cell(1, 35).v === 15, "Sheet1: title and days 1-15 over income and expense");
ok(s19.cell(2, 1).v === "My home 2026 10 сарын 1-15" && s19.cell(1, 1).v === "Бүгд" && s19.cell(1, 2).v === "" && s19.cell(1, 3).v === "← Худалдагчийн нэр", "Sheet19 = October 1-15; A1 is the view list, B1 waits for a seller's name, C1 is its label");
ok(ss.getSheetByName("Sheet20").cell(1, 20).v === "" && ss.getSheetByName("Sheet20").cell(1, 36).v === "" && ss.getSheetByName("Sheet20").cell(1, 19).v === 30 && ss.getSheetByName("Sheet20").cell(2, 1).v === "My home 2026 10 сарын 16-31", "October 16-31: T and AJ (transfers) get no day number, last day column stays 30");
ok(ss.getSheetByName("Sheet22").cell(2, 1).v === "My home 2026 11 сарын 16-30", "November 16-30 title");
ok(s54.cell(2, 1).v === "Соронзон 2026 12 сарын 16-31" && s54.cell(1, 5).v === 16 && s54.cell(1, 19).v === 30, "Sheet54 = Соронзон December 16-31");
ok(ss.getSheetByName("Sheet56").cell(2, 1).v === "My home 2026 9 сарын 30" && ss.getSheetByName("Sheet56").cell(6, 41).v !== "", "Sheet56 is not treated as a year sheet");
{
  // Тооллогын хуудас: Sheet100 = Соронзон, Sheet100* = My home, rebuilt and locked.
  const stock = [["Sheet100", "Sor/", "Соронзон: "], ["Sheet100*", "My/", "My home: "]].map(([name, prefix, label]) => ({ s: ss.getSheetByName(name), prefix, label }));
  ok(stock.every(({ s }) => s !== null && s.maxCols === 3 && s.maxRows >= 500), "both stock sheets exist with 3 columns and room for the longest shop sheet");
  ok(stock.every(({ s, prefix, label }) => s.cell(1, 1).f.includes(label) && s.cell(1, 2).f.includes(`"${prefix}"&YEAR(t)`) && s.cell(3, 1).f.startsWith("=ARRAYFORMULA(LET(t,TODAY()") && s.cell(3, 1).f.includes(`tab,"${prefix}"&YEAR(t)`)), "each stock sheet reads its own shop's sheet of today");
  ok(stock.every(({ s }) => s.cell(2, 1).v === "Барааны нэр" && s.cell(2, 2).v === "Үлдэгдэл" && s.cell(2, 3).v === "Тооллого" && s.cell(5, 4).v === "" && s.cell(6, 41).v === ""), "old content is gone, headers written, C is the empty count column");
  ok(stock.every(({ s }) => s.protections.length === 1 && s.protections[0].isSheet && s.protections[0].editors.length === 0 && !s.protections[0].warning), "the whole sheet is locked for everyone but the owner (no divider or header protections left)");
  ok(stock.every(({ s }) => { const open = s.protections[0].getUnprotectedRanges(); return open.length === 1 && open[0].c === 3 && open[0].r === 3 && open[0].w === 1 && open[0].getLastRow() === s.maxRows; }), "only column C (from row 3) stays open for a count typed by hand");
  ok(stock.every(({ s }) => s.frozenRows === 2 && s.cfRules.length === 3), "header frozen, category / negative / zero rules set");
}
const s200 = ss.getSheetByName("Sheet200");
ok(s200.cell(6, 3).f === "=SUM('My/2026/1/15'!$AR$3,'My/2026/1/31'!$AR$3)" && s200.cell(9, 14).f === "=SUM('Sor/2026/12/15'!$AS$3,'Sor/2026/12/31'!$AS$3)" && s200.cell(10, 3).f === "=sum(C6+C8)", "Sheet200 rewritten with the new names, rows 10+ untouched");
ok(ss.sheets.filter(s => s.name.startsWith("My/")).every(s => s.maxRows === 250) && ss.sheets.filter(s => s.name.startsWith("Sor/")).every(s => s.maxRows === 400), "My home sheets are cut to 250 rows, Соронзон sheets to 400");
ok(ss.sheets.filter(s => !/^(My|Sor)\//.test(s.name) && !s.name.startsWith("_") && !s.name.startsWith("Sheet100") && !s.name.endsWith(" тооллого")).every(s => s.maxRows === 400), "other sheets keep the general limit of 400 (stock and count sheets are never trimmed)");
{
  const limits = makeRunner(new Spreadsheet(), { triggers: [], mails: [] })(true, `["My/2026/3/15","Sor/2026/3/15","Sheet7","Sheet40","Sheet200","Тэмдэглэл"].map(rowLimitFor)`).value;
  eq(limits, [250, 400, 250, 400, 400, 400], "the limit follows the shop, under the new and the old tab names");
}
{
  // Like Google Sheets: the array formulas of AL:AN leave an empty text in every row down to the bottom.
  const book = new Spreadsheet(); const sh = new Sheet(book, "My/2026/3/15"); book.sheets.push(sh);
  for (let r = 1; r <= 100; r++) sh.cell(r, 1).v = "бараа " + r;
  for (let r = 4; r <= 1000; r++) { sh.cell(r, 38).f = "(spill)"; sh.cell(r, 38).v = ""; }
  ok(sh.getLastRow() === 1000, "(the fake sheet looks full to getLastRow, like the real one)");
  const line = makeRunner(book, { triggers: [], mails: [] })(false, `trimSheet(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("My/2026/3/15"))`).value;
  ok(sh.maxRows === 250 && /1000 -> 250/.test(line), "rows that hold only an empty formula result are cut");
  const offBook = new Spreadsheet(); const offSheet = new Sheet(offBook, "My/2026/3/15"); offBook.sheets.push(offSheet); offSheet.cell(10, 1).v = "бараа";
  const offLine = makeRunner(offBook, { triggers: [], mails: [] })(false, `trimSheet(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("My/2026/3/15"))`, (c) => c.replace("rowLimit: 250", "rowLimit: 0")).value;
  ok(offSheet.maxRows === 1000 && offLine.includes("мөр хасахгүй"), "rowLimit 0: no row of that shop is removed");
  const empty = makeRunner(book, { triggers: [], mails: [] })(true, `[lastFilledRow([]), lastFilledRow([["", ""], ["", 0], ["", ""]])]`).value;
  eq(empty, [0, 2], "lastFilledRow: an empty sheet is 0, a zero counts as data");
}
ok(ss.getSheetByName("_sync") && ss.getSheetByName("_sync").getLastRow() === 48, "baseline saved for 48 sheets");
ok(state.triggers.map(t => t.getHandlerFunction()).sort().join() === "onInstantChange,onSheetEdit,onSheetStructureChange,runBackgroundJobs", "four triggers installed: instant (edit + change), guard, hourly");

// ---------- charts on Sheet200 follow the numbers by themselves ----------
{
  const s200c = ss.getSheetByName("Sheet200"); const graph = ss.getSheetByName("_graph");
  const titles = s200c.getCharts().map(c => c.spec.options.title);
  eq(titles, ["Сарын борлуулалт ба ашиг"], "one chart: monthly sales and profit of each shop");
  ok(graph && graph.hidden, "the helper sheet is hidden");
  const sales = s200c.getCharts()[0].spec;
  ok(sales.ranges[0].sheet === graph && sales.ranges[0].c === 1 && sales.ranges[0].h === 13 && sales.ranges[0].w === 5 && sales.type === "COLUMN" && sales.position[0] === 15, "the chart reads header + 12 months and sits under the table");
  ok(graph.cell(2, 1).f === "=TRANSPOSE('Sheet200'!C5:N5)" && graph.cell(2, 2).f === "=ARRAYFORMULA(IFERROR(TRANSPOSE('Sheet200'!C6:N6),0))" && graph.cell(2, 3).f.includes("C7:N7") && graph.cell(2, 4).f.includes("C8:N8") && graph.cell(2, 5).f.includes("C9:N9") && graph.cell(1, 3).v === "My home ашиг" && graph.cell(1, 4).v === "Соронзон борлуулалт", "the monthly block is formulas on Sheet200: no script run is needed for a new month");
  ok(graph.cell(2, 2).nf === '#,##0 "₮"' && graph.cell(13, 5).nf === '#,##0 "₮"' && sales.options.colors.length === 4 && sales.options.vAxis.format === '#,##0 "₮"', "the amounts are in tugrik, so the axis is too");
  ok(s200c.cell(6, 15).f === "=SUM(C6:N6)" && s200c.cell(11, 15).f === "=SUM(C11:N11)" && s200c.cell(12, 15).f === "=IFERROR(O11/O10,0)" && s200c.cell(6, 15).nf === '#,##0 "₮"' && s200c.cell(12, 15).nf === "0%", "the yearly total column sums whatever months are entered so far");
  ok(s200c.cell(12, 2).v.startsWith("Ашгийн хувь") && s200c.cell(12, 3).nf === "0%" && s200c.cell(10, 3).f === "=sum(C6+C8)", "the profit-share row gets a name and a percent format; existing formulas stay");
  const bg = (r, c) => s200c.cell(r, c).bg;
  ok(bg(5, 1) === "#d9d9d9" && bg(6, 1) === "#cfe2f3" && bg(7, 15) === "#cfe2f3" && bg(8, 3) === "#f4cccc" && bg(9, 15) === "#f4cccc" && bg(10, 2) === "#d9ead3" && bg(12, 1) === "#fff2cc" && bg(4, 1) === "#ffffff", "the table is coloured by shop across A..O");
  run(false, 'reconcileSheet(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet200"))');
  ok(bg(6, 1) === "#cfe2f3" && bg(8, 3) === "#f4cccc", "the colour clean-up leaves the table of Sheet200 alone");
  const old = (title) => ({ spec: { options: { title } }, getOptions: () => ({ get: () => title }) });
  s200c.charts = [old("Сарын борлуулалт"), old("Ашгийн хувь (ашиг / борлуулалт)")]; s200c.cell(6, 15).f = "=МИНИЙ"; const again = run(false, "installCharts()");
  eq(s200c.getCharts().map(c => c.spec.options.title), ["Сарын борлуулалт ба ашиг"], "old charts are removed and the kept one is drawn again");
  ok(again.logs.includes("хуучин 2 графикийг") && s200c.cell(6, 15).f === "=МИНИЙ", "a total typed by hand is left alone");
  s200c.cell(6, 15).f = "=SUM(C6:N6)";
}

// ---------- header rows 1-3: an edited or deleted text comes back ----------
{
  const sh = ss.getSheetByName(N(33)); const fire = (r, c, h = 1, w = 1) => { state.editEvent = { range: sh.getRange(r, c, h, w) }; const out = run(false, "onSheetEdit(editEvent)"); state.editEvent = undefined; return out; };
  const isDate = (x) => Object.prototype.toString.call(x) === "[object Date]";
  const original = [[3, 1], [3, 2], [3, 3], [3, 4], [1, 1], [1, 2], [1, 3], [1, 39]].map(([r, c]) => [r, c, sh.cell(r, c).v]); const savedNote = state.docProps["header:" + sh.getSheetId()];
  sh.cell(3, 1).v = "Барааны нэрс"; sh.cell(3, 2).v = "анхны үлдэгдэл"; sh.cell(3, 3).v = "өртөг үнэ"; sh.cell(3, 4).v = "зарах үнэ"; sh.cell(1, 39).v = new Date(Date.UTC(2026, 1, 1)); run(false, "rememberHeaders()");
  const title = sh.cell(2, 1).v, label = sh.cell(3, 1).v, day = sh.cell(1, 5).v, formula = sh.cell(3, 38).f, start = sh.cell(1, 39).v;
  ok(typeof title === "string" && title !== "" && formula.startsWith("=ARRAYFORMULA") && isDate(start), "(the header has a title, a formula and a date)");
  sh.cell(2, 1).v = "өөр гарчиг"; fire(2, 1); ok(sh.cell(2, 1).v === title, "an edited title comes back at once");
  sh.cell(3, 1).v = ""; sh.cell(1, 5).v = ""; sh.cell(3, 38).f = ""; sh.cell(3, 38).v = ""; sh.cell(1, 39).v = ""; sh.cell(1, 3).v = ""; sh.cell(1, 2).v = "Оюун"; fire(1, 1, 3, 40);
  ok(sh.cell(3, 1).v === label && sh.cell(1, 5).v === day && sh.cell(3, 38).f === formula, "deleted labels, day numbers and formulas come back");
  ok(isDate(sh.cell(1, 39).v) && sh.cell(1, 39).v.getTime() === start.getTime(), "a deleted date comes back as a date");
  ok(sh.cell(1, 3).v === "← Худалдагчийн нэр" && sh.cell(1, 2).v === "Оюун", "the seller label comes back; a typed seller name is kept");
  sh.cell(1, 2).v = "Болд"; fire(1, 2); ok(sh.cell(1, 2).v === "Болд", "the seller name may be changed at any time");
  ok(s200.cell(13, 2).v === "My home худалдагч" && s200.cell(14, 2).v === "Соронзон худалдагч" && s200.cell(14, 4).f === `=IFERROR(TEXTJOIN(", ",TRUE,UNIQUE({'Sor/2026/2/15'!$B$1;'Sor/2026/2/28'!$B$1})),"")` && s200.cell(13, 14).f.includes("'My/2026/12/31'!$B$1") && s200.cell(13, 3).bg === "#cfe2f3" && s200.cell(14, 3).bg === "#f4cccc", "Sheet200 names the seller of each month, one line per shop");
  {
    const hidden = () => [...(sh.hiddenCols || [])].sort((x, y) => x - y); const span = (a, b) => Array.from({ length: b - a + 1 }, (_, i) => a + i);
    ok(ss.getSheetByName(N(40)).cell(1, 1).v === "Бүгд" && ss.getSheetByName(N(40)).validations.some(v => v.r === 1 && v.c === 1 && v.rule.list.length === 5 && v.rule.allowInvalid === false), "setup put the view list in A1");
    const pick = (value) => { sh.cell(1, 1).v = value; fire(1, 1); };
    pick("Зөвхөн зарлага"); eq(hidden(), [3, 4, ...span(5, 20)], "\"Зөвхөн зарлага\" hides the income columns E..T and the prices C..D");
    pick("Зөвхөн орлого"); eq(hidden(), [3, 4, ...span(21, 36)], "\"Зөвхөн орлого\" hides the expense columns U..AJ and the prices C..D, and shows the income again");
    ok(sh.cell(1, 1).v === "Зөвхөн орлого", "the choice is not put back by the header guard");
    pick("Бүгд"); eq(hidden(), [], "\"Бүгд\" shows every column");
    pick("Хэвлэх (нийт дүнгүй)"); eq(hidden(), [3, 4, ...span(41, 45)], "\"Хэвлэх\" hides cost and price C..D and the totals AO..AS; the list in A1 stays in view");
    pick("Зөвхөн зарлага"); eq(hidden(), [3, 4, ...span(5, 20)], "leaving the print view shows the totals again (the prices stay hidden in the expense view)");
    pick("Зөвхөн буруу"); ok(JSON.stringify(hidden()) === JSON.stringify(span(41, 45)) && sh.filter && sh.filter.range.r === 3 && sh.filter.range.c === 1 && sh.filter.range.w === sh.maxCols && sh.filter.criteria[39].values[0] === "буруу", "\"Зөвхөн буруу\" hides the totals AO..AS and filters row 3 down on AM = буруу");
    pick("Зөвхөн буруу"); ok(sh.filter && sh.filter.criteria[39], "choosing it again sets the filter anew");
    pick("Зөвхөн зарлага"); ok(!sh.filter && !hidden().includes(41), "another choice removes that filter and shows the totals again");
    { const own = sh.getRange(3, 1, 10, 5).createFilter(); pick("Бүгд"); ok(sh.filter === own, "a filter made by hand is left alone by the other choices"); pick("Зөвхөн буруу"); ok(sh.filter !== own && sh.filter.criteria[39], "but gives way to \"Зөвхөн буруу\""); pick("Зөвхөн зарлага"); }
    pick(""); eq(hidden(), [], "an emptied choice shows every column");
    sh.cell(1, 1).v = "Бүгд";
  }
  const before = ss.logical(); sh.reads = 0; fire(5, 5); ok(ss.logical() === before, "an edit below the header changes nothing in it");
  sh.cell(3, 2).v = ""; ok(run(true, "runRestoreHeaders(SpreadsheetApp.getActiveSpreadsheet())").value.length === 1 && sh.cell(3, 2).v === "", "test mode only reports the header to restore");
  run(false, "runBackgroundJobs()"); ok(sh.cell(3, 2).v !== "", "the hourly job restores a header too");
  eq(run(true, "[taskOrder(4, 0), taskOrder(4, 3), taskOrder(1, 0)]").value, [[0, 1, 2, 3], [3, 0, 1, 2], [0]], "the hourly tasks are taken in order, starting where the last run stopped");
  { const noTime = (c) => c.replace("const BACKGROUND_TIME_BUDGET_MS = 2.5 * 60 * 1000;", "const BACKGROUND_TIME_BUDGET_MS = -1;");
    const label = sh.cell(3, 2).v; sh.cell(3, 2).v = ""; state.props.backgroundNext = "0";
    const first = run(false, "runBackgroundJobs()", noTime);
    ok(first.logs.includes("ажлыг дараагийн удаад үлдээлээ") && state.props.backgroundNext === "1" && sh.cell(3, 2).v === "", "a run that is out of time does one task and leaves the rest for the next run");
    let runs = 1; while (sh.cell(3, 2).v === "" && runs < 20) { run(false, "runBackgroundJobs()", noTime); runs += 1; }
    ok(sh.cell(3, 2).v === label && runs === 10 && state.props.backgroundNext === "10", "the headers task is reached on the 10th short run, the count sheets come next");
    run(false, "runBackgroundJobs()", noTime); run(false, "runBackgroundJobs()", noTime); run(false, "runBackgroundJobs()", noTime); ok(state.props.backgroundNext === "0", "after the last tasks (name order, count sheets, the shops' transfers) it starts over");
  }
  sh.cell(3, 4).v = ""; sh.cell(3, 3).v = "шинэ нэр"; run(false, "rememberHeaders()");
  ok(sh.cell(3, 4).v !== "" && sh.cell(3, 3).v === "шинэ нэр", "remembering anew keeps a deliberate change and still puts back an emptied cell");
  sh.cell(3, 3).v = "өртөг үнэ"; run(false, "rememberHeaders()");
  const other = ss.getSheetByName("Sheet200"); const keep = other.cell(5, 3).v; other.cell(1, 1).v = "миний тэмдэглэл"; state.editEvent = { range: other.getRange(1, 1) }; run(false, "onSheetEdit(editEvent)"); state.editEvent = undefined;
  ok(other.cell(1, 1).v === "миний тэмдэглэл", "sheets that are not year sheets are left alone"); other.cell(1, 1).v = "";
  original.forEach(([r, c, value]) => { sh.cell(r, c).v = value; }); state.docProps["header:" + sh.getSheetId()] = savedNote;
}

// ---------- idempotence ----------
{
  const before = ss.logical();
  const r = run(false, "setupYear()");
  ok(ss.logical() === before, "second real run changes nothing");
  ok(state.triggers.length === 4 && r.logs.includes("БҮГД ДУУСЛАА"), "triggers not duplicated");
}

// ---------- style and clean-up are added to sheets whose totals were installed by an older version ----------
{
  const s1 = ss.getSheetByName("Sheet1"); for (let col = 41; col <= 45; col++) { s1.cell(3, col).nf = ""; s1.cell(1, col).bg = "#ffffff"; s1.cell(8, col).bg = "#ea9999"; }
  const dry = run(true, "setupYear()"); ok(s1.cell(3, 41).nf === "" && s1.cell(8, 41).bg === "#ea9999" && dry.logs.includes("My/2026/1/15: хийх ажил: нийт дүнг өнгөөр ялгах") && dry.logs.includes("My/2026/1/15: шинэ зааглагч 0, авсан хамгаалалт 0"), "dry run only reports the missing style and clean-up");
  run(false, "setupYear()"); ok(s1.cell(3, 45).nf === '#,##0 "₮"' && s1.cell(1, 41).bg === "#9fc5e8" && s1.cell(8, 41).bg === "#ffffff" && s1.cell(3, 44).f === "=SUMPRODUCT(D4:D,AN4:AN)", "style and clean-up applied, formulas untouched");
  const again = run(false, "setupYear()"); ok(again.logs.includes("My/2026/1/15: нийт дүн аль хэдийн тавигдсан"), "nothing left to do on the next run");
}

// ---------- an old 31 in T1/AJ1 is removed; idle background job writes nothing ----------
{
  const sync = ss.getSheetByName("_sync"); const clearsBefore = sync.clears; const before = ss.logical();
  run(false, "runBackgroundJobs()"); run(false, "runBackgroundJobs()");
  ok(sync.clears === clearsBefore && ss.logical() === before, "background job with nothing to do does not write to the file");
}

// ---------- a dragged tab is put back; a coloured-name category gets protected ----------
{
  const moved = ss.getSheetByName("Sheet40"); ss.setActiveSheet(moved); ss.moveActiveSheet(1);
  const s3 = ss.getSheetByName("Sheet3"); s3.insertRowAfter(10); s3.getRange(11, 1).setValue("Хөргөгч"); s3.cell(11, 1).bg = "#e06666";
  s3.insertRowAfter(12); s3.getRange(13, 1).setValue("Шинэ будсан бараа"); s3.cell(13, 1).bg = "#e06666";
  run(false, "runBackgroundJobs()");
  const order = ss.sheets.map(x => x.name);
  ok(order.indexOf(N(40)) === order.indexOf(N(9)) + 1 && order[0] === "Sheet100", "background job puts a dragged tab back");
  const locked = s3.protections.filter(p => p.desc.startsWith("Ангиллын")).map(p => s3.cell(p.range.r, 1).v);
  ok(locked.includes("Хөргөгч") && !locked.includes("Шинэ будсан бараа"), "coloured category name is protected, coloured new product is left editable");
  ok([2, 20, 36, 40, s3.maxCols].every(col => s3.cell(11, col).bg === "#e06666") && s3.cell(13, 2).bg === "#ffffff" && s3.cell(12, 2).bg === "#ffffff", "category colour is extended across its whole row, other rows keep theirs");
  const again = run(false, "protectDividers()"); ok(again.logs.includes("бүх зааглагч хамгаалагдсан"), "extended divider is a full band now: nothing left to do");
  const s5 = ss.getSheetByName("Sheet5"); s5.insertRowAfter(8); s5.getRange(9, 1).setValue("Плитка"); s5.cell(9, 1).bg = "#ea9999";
  const dry = run(true, "protectDividers()"); ok(s5.cell(9, 5).bg === "#ffffff" && dry.logs.includes("My/2026/3/15:") && dry.logs.includes("өнгийг зассан нүд:"), "dry run only reports the colour it would extend");
  const s4 = ss.getSheetByName("Sheet4"); const find = (sh, label) => { for (let r = 1; r <= sh.rows.length; r++) if (sh.cell(r, 1).v === label) return r; return null; };
  const cat4 = find(s4, "Хөргөгч");
  ok(find(s4, "Шинэ будсан бараа") !== null && cat4 !== null && s4.cell(cat4, 1).bg === "#e06666" && s4.cell(cat4, 30).bg === "#e06666" && s4.protections.some(p => p.range.r === cat4 && p.desc.startsWith("Ангиллын")), "the new product is carried forward as a product; the new category is carried to the other sheets as a category (coloured, protected)");
}

// ---------- a changed price or cost is carried to later sheets ----------
{
  const P = (n) => ss.getSheetByName(`Sheet${n}`); const rowOf = (sh, label) => { for (let r = 1; r <= sh.rows.length; r++) if (sh.cell(r, 1).v === label) return r; return null; };
  const price = (n, label) => P(n).cell(rowOf(P(n), label), 4).v; const cost = (n, label) => P(n).cell(rowOf(P(n), label), 3).v;
  const setPrice = (n, label, v) => { P(n).cell(rowOf(P(n), label), 4).v = v; }; const setCost = (n, label, v) => { P(n).cell(rowOf(P(n), label), 3).v = v; };
  ok(ss.getSheetByName("_prices") && ss.getSheetByName("_prices").hidden && ss.getSheetByName("_prices").getLastRow() === 48, "setup saved the price baseline for 48 sheets");
  eq(run(true, `findPriceChanges({a:[1,2],b:[3,4],gone:[9,9]},{a:[1,5],b:[3,4],fresh:[7,7]})`).value, [{ key: "a", index: 1, from: 2, to: 5 }], "only products known before and now, only the changed field");
  setPrice(19, "A 5", 175000);
  const dry = run(true, "syncPricesForward()"); ok(price(20, "A 5") === 150 && dry.logs.includes("ТУРШИЛТ") && dry.logs.includes('"a 5" үнэ 150 -> 175000, дараагийн 5 sheet-д'), "dry run reports the price change and writes nothing");
  run(false, "runBackgroundJobs()");
  ok([20, 21, 22, 23, 24].every(n => price(n, "A 5") === 175000), "price changed on October 1-15 reaches every later sheet of that shop");
  ok(price(18, "A 5") === 150 && price(1, "A 5") === 150, "earlier sheets keep the old price");
  ok(ss.getSheetByName("Sheet49").cell(rowOf(ss.getSheetByName("Sheet49"), "A 5"), 4).v === 150, "the other shop is untouched");
  ok(cost(20, "A 5") === 100 && price(20, "A 6") === 150, "cost of that product and other products are untouched");
  const before = ss.logical(); run(false, "runBackgroundJobs()"); ok(ss.logical() === before, "second run: nothing more to carry, nothing written");
  setPrice(22, "A 7", 999); run(false, "syncPricesForward()");
  setPrice(20, "A 7", 160); run(false, "syncPricesForward()");
  ok(price(21, "A 7") === 160 && price(22, "A 7") === 999 && price(23, "A 7") === 999, "a later sheet with its own price keeps it, and the change stops there");
  setCost(19, "B 2", 120); setCost(21, "B 2", 130); run(false, "syncPricesForward()");
  ok(cost(20, "B 2") === 120 && cost(21, "B 2") === 130 && cost(24, "B 2") === 130, "two changes at once: each is carried from its own sheet onward");
  P(23).deleteRow(rowOf(P(23), "B 9")); setPrice(21, "B 9", 180); run(false, "syncPricesForward()");
  ok(price(22, "B 9") === 180 && price(24, "B 9") === 180, "a sheet without the product is skipped, later sheets still get the price");
}

// ---------- an edit of a name, cost or price is carried at once ----------
{
  const P = (n) => ss.getSheetByName(`Sheet${n}`); const rowOf = (sh, label) => { for (let r = 1; r <= sh.rows.length; r++) if (sh.cell(r, 1).v === label) return r; return null; };
  const edit = (sheet, row, col, w = 1) => { state.editEvent = { range: sheet.getRange(row, col, 1, w) }; const r = run(false, "onSheetEdit(editEvent)"); state.editEvent = undefined; return r; };
  const v = (e) => run(true, e).value;
  ok(v(`isWatchedEdit({sheetName:"My/2026/2/28",lastRow:9,firstColumn:4,lastColumn:4})`) === true && v(`isWatchedEdit({sheetName:"My/2026/2/28",lastRow:9,firstColumn:2,lastColumn:6})`) === true, "price edit, or a paste covering C and D, is watched");
  ok(v(`isWatchedEdit({sheetName:"My/2026/2/28",lastRow:9,firstColumn:5,lastColumn:36})`) === false && v(`isWatchedEdit({sheetName:"My/2026/2/28",lastRow:3,firstColumn:1,lastColumn:4})`) === false && v(`isWatchedEdit({sheetName:"Sheet200",lastRow:9,firstColumn:4,lastColumn:4})`) === false, "daily entries, header rows and other sheets are ignored");
  const r4 = rowOf(P(4), "B 4"); P(4).cell(r4, 4).v = 222000;
  const before = ss.logical(); edit(P(4), r4, 9); ok(ss.logical() === before, "an edit in a day column starts nothing");
  edit(P(4), r4, 4);
  ok([5, 6, 12, 24].every(n => P(n).cell(rowOf(P(n), "B 4"), 4).v === 222000) && P(3).cell(rowOf(P(3), "B 4"), 4).v === 150, "price edited on My/2026/2/28 is on every later sheet at once, earlier ones unchanged");
  const last = rowOf(P(6), "A 12"); P(6).insertRowAfter(last); P(6).getRange(last + 1, 1).setValue("Шууд нэмсэн"); edit(P(6), last + 1, 1);
  ok([7, 20].every(n => rowOf(P(n), "Шууд нэмсэн") !== null) && rowOf(P(5), "Шууд нэмсэн") === null, "a product typed into column A appears on the later sheets at once");
  P(6).cell(last + 1, 3).v = 5000; P(6).cell(last + 1, 4).v = 7000; edit(P(6), last + 1, 3, 2);
  ok(P(20).cell(rowOf(P(20), "Шууд нэмсэн"), 3).v === 5000 && P(20).cell(rowOf(P(20), "Шууд нэмсэн"), 4).v === 7000, "its cost and price, typed afterwards, follow");
  ok(state.props.instantPending === undefined, "no note is left behind");
}

// ---------- a count (AK) becomes the next sheet's opening balance (B), at once, and stays editable ----------
{
  const P = (n) => ss.getSheetByName(`Sheet${n}`); const rowOf = (sh, label) => { for (let r = 1; r <= sh.rows.length; r++) if (sh.cell(r, 1).v === label) return r; return null; };
  const v = (e) => run(true, e).value;
  eq([v(`decideOpening("", 7)`), v(`decideOpening("", 0)`), v(`decideOpening(4, 9)`), v(`decideOpening(7, "")`), v(`decideOpening(9, 9)`), v(`decideOpening("", "abc")`), v(`decideOpening("", "")`)],
     ["write", "write", "write", "same", "same", "same", "same"], "the opening balance is the count, once there is one");
  const count = (n, label) => P(n).cell(rowOf(P(n), label), 37); const opening = (n, label) => P(n).cell(rowOf(P(n), label), 2);
  const typeCount = (n, label, value) => { count(n, label).v = value; state.editEvent = { range: P(n).getRange(rowOf(P(n), label), 37) }; const r = run(false, "onSheetEdit(editEvent)"); state.editEvent = undefined; return r; };
  ok(opening(21, "A 4").v === "", "(a future sheet has no opening balance yet)");
  ss.sheets.forEach(x => { x.reads = 0; });
  typeCount(20, "A 4", 7);
  ok(opening(21, "A 4").v === 7, "a count typed on one sheet is the next sheet's opening balance at once");
  eq(ss.sheets.filter(x => x.reads > 0 && !x.name.startsWith("_")).map(x => x.name), [20, 21].map(N), "only that sheet and the next one are read");
  ok(opening(22, "A 4").v === "" && opening(20, "A 4").v !== 7, "sheets further on, and the sheet itself, are not touched");
  typeCount(20, "A 4", 9); ok(opening(21, "A 4").v === 9, "a corrected count corrects the opening balance");
  const typeOpening = (n, label, value) => { opening(n, label).v = value; state.editEvent = { range: P(n).getRange(rowOf(P(n), label), 2) }; const r = run(false, "onSheetEdit(editEvent)"); state.editEvent = undefined; return r; };
  ss.sheets.forEach(x => { x.reads = 0; });
  const fixed = typeOpening(21, "A 4", 4);
  ok(opening(21, "A 4").v === 9 && fixed.logs.includes("A 4 4 → 9"), "an opening balance typed by hand is put back to the count at once, and logged");
  eq(ss.sheets.filter(x => x.reads > 0 && !x.name.startsWith("_")).map(x => x.name), [20, 21].map(N), "only the sheet before and that sheet are read");
  typeCount(20, "A 4", ""); ok(opening(21, "A 4").v === 9, "a count that is removed leaves the opening balance as it was");
  typeOpening(21, "A 4", 3); ok(opening(21, "A 4").v === 3, "while nothing is counted on the sheet before, the opening balance can be typed");
  typeCount(20, "A 4", 0); ok(opening(21, "A 4").v === 0, "a count of 0 is carried too");
  { const sh = P(21); const row = rowOf(sh, "A 4"); sh.insertRowAfter(row); sh.cell(row + 1, 1).v = "Шинэ ирсэн бараа"; sh.cell(row + 1, 3).v = 10; sh.cell(row + 1, 4).v = 20;
    typeOpening(21, "Шинэ ирсэн бараа", 6); ok(opening(21, "Шинэ ирсэн бараа").v === 6, "a product that is new on a sheet keeps the opening balance typed for it"); }
  typeOpening(1, "A 4", 77); ok(opening(1, "A 4").v === 77, "the first sheet of the year has no sheet before it: its opening balance is free");
  ok(opening(2, "A 4").v === 8 && count(1, "A 4").v === 8, "setup made old opening balances equal to the count before them");
  run(false, "runBackgroundJobs()"); ok(opening(21, "A 4").v === 0 && opening(21, "Шинэ ирсэн бараа").v === 6, "the hourly job agrees with what the instant run did");
  const before = ss.logical(); run(false, "runBackgroundJobs()"); ok(ss.logical() === before, "and a second hourly run writes nothing");
  count(22, "B 3").v = 6; const dry = run(true, "carryOpeningBalances()"); ok(opening(23, "B 3").v === "" && dry.logs.includes("ТУРШИЛТ") && dry.logs.includes("тооллогоос авсан 1"), "dry run reports and writes nothing");
  run(false, "carryOpeningBalances()"); ok(opening(23, "B 3").v === 6, "the menu item carries it by hand");
  // a sheet whose count column holds balances copied by hand, not a count (Соронзон 4/15-8/15): the
  // next opening balance follows the sheet's computed balance (opening + income - expense)
  eq([v(`carryValue({count: 6, opening: 4, net: 3}, true)`), v(`carryValue({count: "", opening: "", net: 0}, true)`), v(`carryValue({count: 6, opening: 4, net: 3})`)], [7, "", 6],
     "an uncounted sheet carries its balance; a counted one its count");
  const uncounted = makeRunnerFor(code.replace(/const UNCOUNTED_SHEETS = \[[^\]]*\];/, `const UNCOUNTED_SHEETS = ["${N(22)}"];`))(ss, state);
  const row22 = rowOf(P(22), "B 3");
  opening(22, "B 3").v = 4; P(22).cell(row22, 5).v = 3;
  uncounted(false, "carryOpeningBalances()");
  ok(opening(23, "B 3").v === 7, "the sheet after an uncounted one opens with its balance, not with the copied count");
  P(22).cell(row22, 5).v = 5; state.editEvent = { range: P(22).getRange(row22, 5) }; uncounted(false, "onSheetEdit(editEvent)"); state.editEvent = undefined;
  ok(opening(23, "B 3").v === 9, "an income typed on an uncounted sheet reaches the next opening balance at once");
  P(22).cell(row22, 5).v = ""; opening(22, "B 3").v = ""; run(false, "carryOpeningBalances()");
}

// ---------- an instant sync reads only the changed shop, from the changed sheet onward ----------
{
  const P = (n) => ss.getSheetByName(`Sheet${n}`); const rowOf = (sh, label) => { for (let r = 1; r <= sh.rows.length; r++) if (sh.cell(r, 1).v === label) return r; return null; };
  const v = (e) => run(true, e).value;
  eq(v(`mergeScope(mergeScope(mergeScope({}, {chainKey:"my home",fromPosition:18}), {chainKey:"my home",fromPosition:6}), {chainKey:"соронзон",fromPosition:20})`), { "my home": 6, "соронзон": 20 }, "pending changes are remembered per shop, from the earliest sheet");
  eq(v(`listScopes({"my home":6,"*":0})`), [null], "a change of unknown place means every sheet");
  eq(v(`mergeScope(mergeScope({}, {chainKey:"my home",fromPosition:18}, true), {chainKey:"my home",fromPosition:6}, false)`), { "my home": 6, "#products": true }, "a pending product sync is not forgotten when a price edit follows");
  eq(v(`listScopes({"соронзон":22,"#products":true})`), [{ chainKey: "соронзон", fromPosition: 22 }], "the product flag is not a shop");
  {
    const sh = ss.getSheetByName(N(53)); let row = null; for (let r = 4; r <= sh.rows.length; r++) if (typeof sh.cell(r, 4).v === "number" && sh.cell(r, 1).v !== "") { row = r; break; }
    const label = sh.cell(row, 1).v; const next = ss.getSheetByName(N(54)); let nextRow = null; for (let r = 4; r <= next.rows.length; r++) if (next.cell(r, 1).v === label) nextRow = r;
    const oldPrice = sh.cell(row, 4).v; const fire = (oldValue) => { state.editEvent = { range: sh.getRange(row, 4), oldValue }; run(false, "onSheetEdit(editEvent)"); state.editEvent = undefined; };
    const prices = ss.getSheetByName("_prices"); const products = ss.getSheetByName("_sync");
    sh.cell(row, 4).v = oldPrice + 10000; prices.reads = 0; products.reads = 0; fire(String(oldPrice));
    ok(nextRow !== null && next.cell(nextRow, 4).v === oldPrice + 10000, "a typed price reaches December 16-31 at once");
    ok(prices.reads === 0 && products.reads === 0, "one typed price cell is carried without reading the saved prices or products");
    next.cell(nextRow, 4).v = 777; sh.cell(row, 4).v = oldPrice + 20000; fire(oldPrice + 10000);
    ok(next.cell(nextRow, 4).v === 777, "a later sheet with its own price keeps it");
    next.cell(nextRow, 4).v = oldPrice + 20000; run(false, "syncPricesForward()"); // (the hourly job saves the prices as they are now)
    sh.cell(row, 4).v = oldPrice + 30000; prices.reads = 0; products.reads = 0; fire(undefined);
    ok(next.cell(nextRow, 4).v === oldPrice + 30000 && prices.reads > 0 && products.reads === 0, "without the old value (a paste) the saved prices are compared, still without the product sync");
    sh.cell(row, 4).v = oldPrice; next.cell(nextRow, 4).v = oldPrice; run(false, "syncPricesForward()");
    sh.cell(row, 4).v = oldPrice; state.editEvent = { range: sh.getRange(row, 4) }; run(false, "onSheetEdit(editEvent)"); state.editEvent = undefined;
  }
  eq(v(`findChains(ss, {chainKey:"my home",fromPosition:21}).map(c => c.sheets.map(x => x.getName()))`), [["My/2026/11/30", "My/2026/12/15", "My/2026/12/31"]], "scope: one shop, from one period onward");
  P(3).cell(rowOf(P(3), "B 7"), 4).v = 333;
  ss.sheets.forEach(x => { x.reads = 0; });
  const r = rowOf(P(21), "B 8"); P(21).cell(r, 4).v = 444; state.editEvent = { range: P(21).getRange(r, 4) }; run(false, "onSheetEdit(editEvent)"); state.editEvent = undefined;
  ok([22, 23, 24].every(n => P(n).cell(rowOf(P(n), "B 8"), 4).v === 444), "a price changed in November reaches the sheets up to December");
  const readSheets = ss.sheets.filter(x => x.reads > 0 && !x.name.startsWith("_")).map(x => x.name);
  eq(readSheets, [21, 22, 23, 24].map(N), "only those 4 sheets were read, not all 48");
  ok(P(4).cell(rowOf(P(4), "B 7"), 4).v === 150, "(a change waiting on an earlier sheet is not touched by it)");
  run(false, "runBackgroundJobs()"); ok(P(4).cell(rowOf(P(4), "B 7"), 4).v === 333 && P(20).cell(rowOf(P(20), "B 7"), 4).v === 333, "the hourly job carries whatever an instant run did not look at");
}

// ---------- deleted row, new divider, moved tab, inserted rows: handled at once ----------
{
  const P = (n) => ss.getSheetByName(`Sheet${n}`); const rowOf = (sh, label) => { for (let r = 1; r <= sh.rows.length; r++) if (sh.cell(r, 1).v === label) return r; return null; };
  const change = (type, active) => { ss.active = active || null; state.editEvent = { changeType: type, source: ss }; const r = run(false, "onInstantChange(editEvent)"); state.editEvent = undefined; return r; };
  const before = ss.logical(); change("EDIT", P(20)); ok(ss.logical() === before, "a typed value is not handled twice (On edit does it)");
  P(20).deleteRow(rowOf(P(20), "B 11")); change("REMOVE_ROW", P(20));
  ok([21, 22, 24].every(n => rowOf(P(n), "B 11") === null) && rowOf(P(19), "B 11") !== null && rowOf(P(21), "B 10") !== null, "a product deleted on one sheet leaves the later sheets at once");
  const at = rowOf(P(22), "B 5"); P(22).insertRowAfter(at); P(22).getRange(at + 1, 1).setValue("Шинэ ангилал"); for (let c = 1; c <= 35; c++) P(22).cell(at + 1, c).bg = "#00ffff";
  const lockedBefore = P(22).protections.length; change("FORMAT", P(22));
  ok(P(22).protections.length === lockedBefore + 1 && P(22).protections.some(p => P(22).cell(p.range.r, 1).v === "Шинэ ангилал"), "a newly coloured divider is protected at once");
  {
    const newRow = rowOf(P(22), "Шинэ ангилал"); const productRow = rowOf(P(22), "B 5");
    ok(ss.sheets.filter(x => /^(My|Sor)\//.test(x.name)).every(x => x.borderCalls >= 1), "setup removed the cell borders of the dividers on every year sheet");
    for (let c = 5; c <= 36; c++) P(22).cell(newRow, c).border = true; P(22).cell(productRow, 1).border = true; change("FORMAT", P(22));
    ok(P(22).cell(newRow, 5).border === false && P(22).cell(newRow, 36).border === false && P(22).cell(productRow, 1).border === true, "borders drawn on a divider are removed at once; a product row keeps its own");
    ok(P(22).cell(productRow, 20).border === true && P(22).cell(productRow, 8).border !== true, "setup left one line in the daily columns: after T");
    P(22).cell(productRow, 8).border = true; run(false, 'clearBodyBorders(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("' + P(22).name + '"))');
    ok(P(22).cell(productRow, 8).border === false && P(22).cell(productRow, 20).border === true && P(22).cell(productRow, 1).border === true, "a box left on a single cell of E..AJ is removed; the T line and the A..D grid stay");
    {
      const sh = P(22); const last = (() => { let l = 0; for (let r = 4; r <= sh.rows.length; r++) if (String(sh.cell(r, 1).v).trim() !== "") l = r; return l; })();
      const fresh = last + 1; ok([1, 4, 37, 40, 20].every(c => sh.cell(last, c).border === true) && [1, 4, 37, 40, 20].every(c => sh.cell(fresh, c).border !== true && sh.cell(sh.maxRows, c).border !== true), "setup drew the grid down to the last product and nothing below it");
      sh.getRange(fresh, 1).setValue("Шинэ бараа хүрээтэй"); state.editEvent = { range: sh.getRange(fresh, 1) }; run(false, "onSheetEdit(editEvent)"); state.editEvent = undefined;
      ok([1, 4, 37, 40, 20].every(c => sh.cell(fresh, c).border === true) && sh.cell(fresh, 8).border !== true && sh.cell(fresh + 1, 1).border !== true, "a product typed under the last one gets the grid at once: A..D, AK..AN and the T line");
      ok([23, 24].every(n => { const r = rowOf(P(n), "Шинэ бараа хүрээтэй"); return r !== null; }), "(and it is carried to the later sheets as before)");
      sh.getRange(fresh + 1, 1).setValue("Түр бараа"); state.editEvent = { range: sh.getRange(fresh + 1, 1) }; run(false, "onSheetEdit(editEvent)");
      ok(sh.cell(fresh + 1, 1).border === true && sh.cell(fresh + 1, 20).border === true, "(a second product under the last one has the grid)");
      sh.getRange(fresh + 1, 1).setValue(""); state.editEvent = { range: sh.getRange(fresh + 1, 1) }; run(false, "onSheetEdit(editEvent)"); state.editEvent = undefined;
      ok([1, 4, 37, 40, 20].every(c => sh.cell(fresh + 1, c).border === false) && sh.cell(fresh, 1).border === true && sh.cell(fresh, 20).border === true, "the last product cleared: its empty row loses the grid at once, the row above keeps it");
      sh.cell(fresh, 2).border = false; sh.cell(fresh + 3, 2).border = true; sh.cell(fresh + 3, 38).border = true; run(false, 'clearBodyBorders(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("' + sh.name + '"))');
      ok(sh.cell(fresh, 2).border === true && sh.cell(productRow, 3).border === true && sh.cell(fresh + 3, 2).border === false && sh.cell(fresh + 3, 38).border === false, "putting the borders in order: grid on every row down to the last product, none below");
    }
    const calls = P(22).borderCalls; run(false, 'reconcileSheet(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("' + P(22).name + '"))');
    ok(P(22).borderCalls === calls, "a sheet in order is not touched again by the hourly check");
  }
  while (P(22).maxRows <= 250) P(22).insertRowAfter(P(22).getLastRow() + 5);
  ok(P(22).maxRows === 251, "(inserted rows made the My home sheet 251 rows)"); change("INSERT_ROW", P(22)); ok(P(22).maxRows === 250, "an inserted row: the sheet is cut back to its limit at once");
  ss.setActiveSheet(P(45)); ss.moveActiveSheet(2); change("OTHER", P(45));
  const order = ss.sheets.map(x => x.name); ok(order.indexOf(N(45)) === order.indexOf(N(14)) + 1, "a dragged tab is back in place at once");
  {
    const victim = P(46); const realName = victim.name; const where = ss.sheets.indexOf(victim);
    victim.setName("ээжийн өгсөн нэр"); ss.setActiveSheet(victim); ss.moveActiveSheet(1); const renamed = change("OTHER", victim);
    ok(victim.name === realName && ss.sheets.indexOf(victim) === where && renamed.logs.includes("сэргээлээ"), "a renamed and dragged year tab gets its name and its place back at once");
    const other = ss.getSheetByName("Sheet200"); other.setName("Жилийн дүн"); change("OTHER", other);
    ok(other.name === "Жилийн дүн", "a sheet that is not a year sheet may be renamed"); other.setName("Sheet200");
    eq(run(true, `planTabRenames([{id:"1",name:"My/2026/1/15"},{id:"2",name:"хуучин"},{id:"3",name:"Sor/2026/1/15"}], {"2":"Sor/2026/1/15","9":"My/2026/2/15"})`).value, { renames: [], names: { "1": "My/2026/1/15", "3": "Sor/2026/1/15" } }, "a name already in use is not given twice; a deleted sheet is forgotten");
  }
  const s8 = P(8); s8.insertRowAfter(9); s8.getRange(10, 1).setValue("Зурагт"); s8.cell(10, 1).bg = "#93c47d"; state.editEvent = { range: s8.getRange(10, 1) }; run(false, "onSheetEdit(editEvent)"); state.editEvent = undefined;
  ok(s8.protections.some(p => s8.cell(p.range.r, 1).v === "Зурагт") && s8.cell(10, 30).bg === "#93c47d" && rowOf(P(9), "Зурагт") === null, "typing a category name into a coloured cell: protected and extended at once, not copied as a product");
  const idle = ss.logical(); change("FORMAT", ss.getSheetByName("_sync")); change("FORMAT", null); ok(ss.logical() === idle, "changes on hidden helper sheets are ignored");
}

// ---------- rows inserted next to a divider: colour removed, never locked ----------
{
  const P = (n) => ss.getSheetByName(`Sheet${n}`); const rowOf = (sh, label) => { for (let r = 1; r <= sh.rows.length; r++) if (sh.cell(r, 1).v === label) return r; return null; };
  const change = (type, active) => { ss.active = active; state.editEvent = { changeType: type, source: ss }; run(false, "onInstantChange(editEvent)"); state.editEvent = undefined; };
  const locked = (sh) => sh.protections.filter(p => p.desc.startsWith("Ангиллын")).map(p => p.range.r);
  const s11 = P(11); const catB = rowOf(s11, "Cat B"); const lockedBefore = locked(s11).length;
  s11.insertRowAfter(catB); ok(s11.cell(catB + 1, 1).bg === "#0b5394" && s11.cell(catB + 1, 30).bg === "#0b5394", "(an inserted row inherits the whole blue band)");
  change("INSERT_ROW", s11);
  ok(s11.cell(catB + 1, 1).bg === "#ffffff" && s11.cell(catB + 1, 30).bg === "#ffffff" && s11.cell(catB, 30).bg === "#0b5394", "the inherited band is removed at once, the divider keeps its colour");
  ok(locked(s11).length === lockedBefore && !locked(s11).includes(catB + 1) && s11.maxRows === 250, "the new row is not locked; the My home sheet stays at 250 rows");
  s11.insertRowAfter(catB - 1); for (let c = 1; c <= 35; c++) s11.cell(catB, c).bg = "#0b5394"; s11.getRange(catB, 1).setValue("Өнгө өвлөсөн бараа"); s11.getRange(catB, 3, 1, 2).setValues([[10, 20]]);
  state.editEvent = { range: s11.getRange(catB, 1, 1, 4) }; run(false, "onSheetEdit(editEvent)"); state.editEvent = undefined;
  ok(s11.cell(catB, 2).bg === "#ffffff" && !locked(s11).includes(catB) && rowOf(P(12), "Өнгө өвлөсөн бараа") !== null, "a coloured row that got a name and prices is a product: colour removed, not locked, carried forward");
  const s13 = P(13); const at = rowOf(s13, "A 6"); s13.insertRowAfter(at); s13.getRange(at + 1, 1).setValue("Хөлдөөгч"); s13.cell(at + 1, 1).bg = "#6aa84f"; s13.cell(at, 1).bg = "#6aa84f";
  change("FORMAT", s13);
  ok(s13.cell(at, 1).bg === "#ffffff" && s13.cell(at, 4).v === 150 && s13.cell(at + 1, 30).bg === "#6aa84f" && locked(s13).includes(at + 1) && !locked(s13).includes(at), "green product cell above the green category cell: product is plain again, category extended and locked");
}

// ---------- AL:AN: no "invalid" marks; a typed value is removed and the formula comes back ----------
{
  const P = (n) => ss.getSheetByName(`Sheet${n}`);
  const onFormulas = (x) => x.validations.filter(v => v.r >= 4 && v.c + v.w > 38);
  ok(ss.sheets.every(x => onFormulas(x).length === 0), "no validation rule is put on the formula columns (it painted every result red)");
  const s6 = P(6);
  const productRow = (sh, label) => { for (let r = 1; r <= sh.rows.length; r++) if (sh.cell(r, 1).v === label) return r; return null; };
  const r6 = productRow(s6, "A 2"); s6.cell(r6, 39).v = "миний бичсэн"; s6.cell(3, 39).v = "#REF!";
  state.editEvent = { range: s6.getRange(r6, 39) }; run(false, "onSheetEdit(editEvent)"); state.editEvent = undefined;
  ok(s6.cell(r6, 39).v === "" && s6.cell(3, 39).f.startsWith("=ARRAYFORMULA(") && s6.cell(r6, 4).v === 150, "typing into AM: the typed value is removed at once, the formula and the row's data stay");
  const s7 = P(7); const r7 = productRow(s7, "B 2"); s7.cell(r7, 38).v = 5; s7.cell(3, 38).v = "#REF!"; run(false, "runBackgroundJobs()");
  ok(s7.cell(r7, 38).v === "", "the hourly job repairs a blocked column too");
  const before = ss.logical(); state.editEvent = { range: P(8).getRange(productRow(P(8), "A 2"), 37) }; run(false, "onSheetEdit(editEvent)"); state.editEvent = undefined; ok(ss.logical() === before, "typing a count in AK touches nothing");
}

// ---------- colours: dividers coloured across, only the owner's colours stay ----------
{
  const P = (n) => ss.getSheetByName(`Sheet${n}`); const rowOf = (sh, label) => { for (let r = 1; r <= sh.rows.length; r++) if (sh.cell(r, 1).v === label) return r; return null; };
  const v = (e) => run(true, e).value;
  eq(v(`findColumnColors([["#ffffff","#fff2cc"],["#ff0000","#fff2cc"],["#ffffff","#fff2cc"],["#0000ff","#0000ff"]], [0,1,2])`), ["#ffffff", "#fff2cc"], "the usual colour of a column is what most product rows have");
  const infos = `[{row:4,key:"cat",color:"#0000ff",hasLabel:true},{row:5,key:"a",color:"#ffffff",hasLabel:true},{row:6,key:"b",color:"#ffffff",hasLabel:true},{row:7,key:"",color:"#ffffff",hasLabel:false}]`;
  const bgs = `[["#0000ff","#ffffff","#0000ff"],["#ffffff","#ff0000","#fff2cc"],["#ffffff","#ffffff","#fff2cc"],["#00ff00","#ffffff","#fff2cc"]]`;
  eq(v(`buildWantedColors(${infos}, ${bgs}, [4], {})`), [["#0000ff", "#0000ff", "#0000ff"], ["#ffffff", "#ffffff", "#fff2cc"], ["#ffffff", "#ffffff", "#fff2cc"], ["#ffffff", "#ffffff", "#fff2cc"]], "divider: gap filled; product and empty rows: stray colours gone, the column tint stays");
  eq(v(`buildWantedColors(${infos}, ${bgs}, [4], {a:{2:"#ff0000"}})`)[1], ["#ffffff", "#ff0000", "#fff2cc"], "a cell the owner coloured is kept");
  eq(v(`collectOwnColors(${infos}, ${bgs}, [4])`), { a: { 2: "#ff0000" } }, "what differs from the usual colour on a product row is what the owner coloured");
  {
    // the grey of a count filled with the balance by hand is no mark any more (the owner, 2026-10-24)
    const C = v("LAYOUT.count"), GRAY = "#d0e0e3";
    const row = (at, color) => Array.from({ length: C + 1 }, (_, i) => (i === at - 1 ? color : "#ffffff"));
    const infos2 = `[{row:4,key:"a",color:"#ffffff",hasLabel:true},{row:5,key:"b",color:"#ffffff",hasLabel:true},{row:6,key:"c",color:"#ffffff",hasLabel:true}]`;
    const bgs2 = JSON.stringify([row(C, GRAY), row(C, "#ffffff"), row(2, GRAY)]);
    ok(!v("DIVIDER_KEPT_COLORS").includes(GRAY), "the colour keeper no longer keeps the grey");
    const wanted = v(`buildWantedColors(${infos2}, ${bgs2}, [], {a:{${C}:"${GRAY}"},c:{2:"${GRAY}"}})`);
    eq([wanted[0][C - 1], wanted[2][1]], ["#ffffff", GRAY], "the grey goes from the count column even where it was approved; in another column it is an owner's colour like any other");
    eq(v(`collectOwnColors(${infos2}, ${bgs2}, [])`), { c: { 2: GRAY } }, "nor is it remembered as the owner's colour in the count column");
  }
  const act = (who, type, sheet) => { state.actor = who; ss.active = sheet; state.editEvent = { changeType: type, source: ss }; const r = run(false, "onInstantChange(editEvent)"); state.editEvent = undefined; state.actor = ""; return r; };
  const s14 = P(14); const cat = rowOf(s14, "Cat B"); const a3 = rowOf(s14, "A 3"); const a5 = rowOf(s14, "A 5");
  for (let r = 5; r <= s14.getLastRow(); r++) if (r !== cat) { s14.cell(r, 20).bg = "#fff2cc"; s14.cell(r, 36).bg = "#fce5cd"; }
  s14.cell(cat, 2).bg = "#ffffff"; s14.cell(cat, 36).bg = "#fce5cd"; s14.cell(a3, 2).bg = "#ff0000"; s14.cell(a5, 1).bg = "#00ff00";
  act("", "FORMAT", s14);
  ok(s14.cell(cat, 2).bg === "#0b5394" && s14.cell(cat, 20).bg === "#0b5394" && s14.cell(cat, 36).bg === "#0b5394" && s14.cell(cat, s14.maxCols).bg === "#0b5394", "the divider is one colour across the whole row: no white gap in B, no tint in T / AJ");
  ok(s14.cell(a3, 2).bg === "#ffffff" && s14.cell(a5, 1).bg === "#ffffff", "cells coloured by someone who is not the owner are plain again");
  ok(s14.cell(a3, 20).bg === "#fff2cc" && s14.cell(a3, 36).bg === "#fce5cd" && s14.cell(a3, 4).v === 150, "the yellow and peach columns keep their tint; numbers are untouched");
  s14.cell(a3, 2).bg = "#ff0000"; act("owner@example.com", "FORMAT", s14);
  ok(s14.cell(a3, 2).bg === "#ff0000", "a cell coloured by the owner stays");
  s14.cell(a5, 3).bg = "#00ff00"; act("", "FORMAT", s14);
  ok(s14.cell(a5, 3).bg === "#ffffff" && s14.cell(a3, 2).bg === "#ff0000", "afterwards another person's colour goes, the owner's stays");
  s14.insertRowAfter(5); run(false, "runBackgroundJobs()");
  ok(s14.cell(rowOf(s14, "A 3"), 2).bg === "#ff0000" && s14.cell(a3, 2).bg === "#ffffff", "the owner's colour belongs to the product: it moves with the row");
  s14.cell(rowOf(s14, "A 3"), 2).bg = "#ffffff"; act("owner@example.com", "FORMAT", s14); s14.cell(rowOf(s14, "A 3"), 2).bg = "#ff0000"; act("", "FORMAT", s14);
  ok(s14.cell(rowOf(s14, "A 3"), 2).bg === "#ffffff", "once the owner removes a colour it is no longer kept");
  const a5now = rowOf(s14, "A 5"); ss.active = s14; s14.cell(a5now, 5).bg = "#123456"; state.actor = "mom@example.com"; run(false, "approveActiveSheetColors()"); state.actor = ""; act("", "FORMAT", s14);
  ok(s14.cell(a5now, 5).bg === "#ffffff", "the menu item does not let anyone but the owner approve colours");
  const before = ss.logical(); act("", "FORMAT", s14); run(false, "runBackgroundJobs()"); ok(ss.logical() === before, "a sheet that is in order is not written to");
}

// ---------- a required category ("Плетик") is put back when its name or its row is gone ----------
{
  const P = (n) => ss.getSheetByName(`Sheet${n}`); const rowOf = (sh, label) => { for (let r = 1; r <= sh.rows.length; r++) if (sh.cell(r, 1).v === label) return r; return null; };
  const v = (e) => run(true, e).value;
  const cat = `{chainKey:"соронзон",name:"Плетик",firstProductPrefix:"delux 50-60",color:"#b45f06"}`;
  const row = (r, key, o = {}) => `{row:${r},key:${JSON.stringify(key)},hasLabel:${key !== ""},isCandidate:${!!o.divider},isEmptyBand:${!!o.emptyBand},isEmpty:${key === "" && !o.nums}}`;
  eq(v(`planRequiredCategory([${row(9, "xinix 240l")},${row(10, "", { emptyBand: true })},${row(11, "delux 50-60 4 sh")}], ${cat})`), { action: "name", row: 10 }, "the coloured row is there but its name was erased");
  eq(v(`planRequiredCategory([${row(9, "xinix 240l")},${row(10, "delux 50-60 4")}], ${cat})`), { action: "insert", beforeRow: 10 }, "the whole row was deleted");
  eq(v(`planRequiredCategory([${row(9, "xinix")},${row(10, "")},${row(11, "delux 50-60 4")}], ${cat})`), { action: "name-and-color", row: 10 }, "an empty uncoloured row stands there");
  eq(v(`planRequiredCategory([${row(9, "xinix")},${row(10, "плетик")},${row(11, "delux 50-60 4")}], ${cat})`), { action: "color", row: 10 }, "the name is there without its colour");
  eq([v(`planRequiredCategory([${row(10, "плетик", { divider: true })},${row(11, "delux 50-60 4")}], ${cat})`), v(`planRequiredCategory([${row(10, "a")},${row(11, "b")}], ${cat})`)], [null, null], "nothing to do when it is in order, or when the product is not on the sheet");
  const addStoves = (sh, withRow) => { const last = sh.getLastRow(); if (withRow) for (let c = 1; c <= 35; c++) sh.cell(last + 1, c).bg = "#b45f06"; const r = last + (withRow ? 2 : 1); sh.getRange(r, 1).setValue("Delux 50-60 4 sh"); sh.getRange(r, 3, 1, 2).setValues([[100, 150]]); return r; };
  const s42 = P(42); addStoves(s42, true); const s43 = P(43); const lone = addStoves(s43, false); const my = P(10); addStoves(my, false);
  const dry = run(true, "protectDividers()"); ok(rowOf(s42, "Плетик") === null && dry.logs.includes('"Плетик" ангиллын нэр, өнгийг сэргээх') && dry.logs.includes('"Плетик" ангиллын мөрийг дахин үүсгэх'), "dry run reports both kinds of repair and writes nothing");
  run(false, "runBackgroundJobs()");
  const r42 = rowOf(s42, "Плетик"), r43 = rowOf(s43, "Плетик");
  ok(r42 !== null && r42 === rowOf(s42, "Delux 50-60 4 sh") - 1 && s42.cell(r42, 2).bg === "#b45f06" && s42.cell(r42, s42.maxCols).bg === "#b45f06", "the unnamed brown row gets its name back and is coloured across");
  ok(r43 === lone && rowOf(s43, "Delux 50-60 4 sh") === lone + 1 && s43.cell(lone + 1, 4).v === 150 && s43.cell(r43, 20).bg === "#b45f06", "where the row was deleted it is made again above the product; the product is untouched");
  ok([s42, s43].every(sh => sh.protections.some(p => p.desc.startsWith("Ангиллын") && sh.cell(p.range.r, 1).v === "Плетик")), "the category row is protected: nobody but the owner can delete it or erase its name");
  ok(rowOf(my, "Плетик") === null, "the rule belongs to Соронзон only");
  s42.cell(r42, 1).v = ""; run(false, "runBackgroundJobs()"); ok(s42.cell(r42, 1).v === "Плетик", "an erased name is written back");
  const before = ss.logical(); run(false, "runBackgroundJobs()"); ok(ss.logical() === before, "nothing more to do afterwards");
}

// ---------- every sheet of a shop has the same categories ----------
{
  const P = (n) => ss.getSheetByName(`Sheet${n}`); const rowOf = (sh, label) => { for (let r = 1; r <= sh.rows.length; r++) if (sh.cell(r, 1).v === label) return r; return null; };
  const v = (e) => run(true, e).value;
  const L = (cats, rows) => `{categories:${JSON.stringify(cats)},rowOfProduct:${JSON.stringify(rows)},infoAbove:{}}`;
  const full = L([{ key: "a", label: "A", color: "#111", productKeys: ["p1", "p2"] }, { key: "b", label: "B", color: "#222", productKeys: ["p3", "p4"] }], { p1: 5, p2: 6, p3: 8, p4: 9 });
  const lacking = L([{ key: "a", label: "A", color: "#111", productKeys: ["p1", "p2", "p4"] }], { p1: 5, p2: 6, p4: 7 });
  eq(v(`collectSharedCategories([${full},${full},${full},${lacking}], []).map(c => [c.key, c.sheetCount, c.label, c.color])`), [["a", 4, "A", "#111"], ["b", 3, "B", "#222"]], "categories found on enough sheets are the shared ones");
  eq(v(`collectSharedCategories([${full},${lacking}], []).map(c => c.key)`), ["a"].slice(0, 0), "a row only one or two sheets treat as a category is not spread");
  eq(v(`planMissingCategory([${full},${full},${full},${lacking}], 3, {key:"b"})`), { action: "insert", beforeRow: 7 }, "a missing category goes above the first of its products that is on the sheet");
  eq(v(`planMissingCategory([${full},${L([], { x: 5 })}], 1, {key:"b"})`), null, "none of its products on the sheet: nothing is added");
  const s15 = P(15); const s16 = P(16); const gone = rowOf(s15, "Cat B"); s15.deleteRow(gone);
  const s44 = P(44); s44.deleteRow(rowOf(s44, "Cat A"));
  ok(rowOf(s15, "Cat B") === null && rowOf(s44, "Cat A") === null, "(two category rows are deleted, as the owner could)");
  const dry = run(true, "runCategorySync(ss)").value; ok(dry.some(line => line.includes(N(15)) && line.includes('"Cat B"')) && dry.some(line => line.includes(N(44)) && line.includes('"Cat A"')) && rowOf(s15, "Cat B") === null, "dry run reports both and writes nothing");
  run(false, "runBackgroundJobs()");
  const back = rowOf(s15, "Cat B"); const first = rowOf(s15, "B 1");
  ok(back !== null && back === first - 1 && s15.cell(back, 1).bg === "#0b5394" && s15.cell(back, s15.maxCols).bg === "#0b5394", "the missing category is back above its first product, in its usual colour across the row");
  ok(s15.protections.some(p => p.desc.startsWith("Ангиллын") && p.range.r === back) && s15.cell(first, 4).v === 150, "it is protected again; the products are untouched");
  ok(rowOf(s44, "Cat A") === rowOf(s44, "A 1") - 1 && s44.cell(rowOf(s44, "Cat A"), 1).bg === "#f6b26b", "the same for the other shop, with that shop's colour");
  ok(rowOf(s16, "Шинэ ангилал") === null && rowOf(P(23), "Шинэ ангилал") === null, "a category that exists on a single sheet is not copied everywhere");
  const before = ss.logical(); run(false, "runBackgroundJobs()"); ok(ss.logical() === before, "nothing more to do afterwards");
}

// ---------- "зөв" / "буруу" in AM, and the label of B3 ----------
{
  const v = (e) => run(true, e).value;
  const am = v(`buildCalcFormulas(["a","b","c"])[1]`);
  ok(am.includes('IF(AK4:AK=AL4:AL,"зөв","буруу")') && !am.includes('"ok"') && !am.includes("FALSE"), "AM says зөв / буруу");
  const year = ss.sheets.filter(x => /^(My|Sor)\//.test(x.name));
  ok(year.length === 48 && year.every(x => x.cell(3, 39).f.includes('"зөв","буруу"')), "every year sheet has the new AM formula");
  eq(v(`fillRequiredLabels([["a","",""],["","",""],["n","","c"]])`), [["a", "", ""], ["", "", ""], ["n", "анхны үлдэгдэл ", "c"]], "an empty B3 gets its label, nothing else is touched");
  eq(v(`fillRequiredLabels([["a"],["b"],["n","миний нэр"]])`)[2], ["n", "миний нэр"], "a B3 that has a text keeps it");
  ok(year.every(x => String(x.cell(3, 2).v).trim() !== ""), "after setup no year sheet has an empty B3");
}

// ---------- same name twice; the name check ----------
{
  const v = (e) => run(true, e).value;
  const pp = (list) => JSON.stringify(list.map(([key, price, cost]) => ({ key, price: price ?? "", cost: cost ?? "" })));
  eq(v(`pairProducts(${pp([["a", 100], ["b", 5]])}, ${pp([["b", 9], ["c", 1], ["a", 999]])})`), [1, -1, 0], "a name that stands once is paired whatever its price; a new name has no pair");
  eq(v(`pairProducts(${pp([["g", 110000], ["g", 120000]])}, ${pp([["g", 125000], ["g", 110000]])})`), [1, 0], "same-named rows are paired by the nearest price, not by their order");
  eq(v(`pairProducts(${pp([["g", 110000], ["g", 120000]])}, ${pp([["g", 170000]])})`), [1], "one of two is gone: the remaining row continues the nearer one");
  eq(v(`pairProducts(${pp([["g"], ["g"]])}, ${pp([["g"], ["g"]])})`), [0, 1], "without prices they go by their order");
  eq(v(`pairProducts(${pp([["g", null, 50], ["g", null, 80]])}, ${pp([["g", null, 82], ["g", null, 49]])})`), [1, 0], "the cost tells them apart when no sale price is written");
  const P = (n) => ss.getSheetByName(N(n));
  const twin = (sh, count) => { sh.insertRowAfter(6); sh.cell(7, 1).v = sh.cell(6, 1).v; sh.cell(7, 3).v = 1; sh.cell(7, 4).v = 2; sh.cell(7, 37).v = count; };
  const a = P(5), b = P(6); twin(a, 99); twin(b, ""); a.cell(6, 37).v = 41;
  run(false, "carryOpeningBalances()");
  ok(b.cell(6, 2).v === 41 && b.cell(7, 2).v === 99, "two rows with one name each get the count of their own row");
  { const [p6, p7] = [b.cell(6, 4).v, b.cell(7, 4).v]; b.cell(6, 4).v = a.cell(7, 4).v; b.cell(7, 4).v = a.cell(6, 4).v; run(false, "carryOpeningBalances()");
    ok(b.cell(6, 2).v === 99 && b.cell(7, 2).v === 41, "when the two rows stand the other way round, each still follows the row with its own price"); b.cell(6, 4).v = p6; b.cell(7, 4).v = p7; run(false, "carryOpeningBalances()"); }
  eq(v(`compareNames([{key:"a",label:"A"},{key:"b",label:"B"}], [{key:"a",label:"A"},{key:"c",label:"C "},{key:"c",label:"C"}])`), { gone: ["B"], fresh: ["C"], repeated: ["C"] }, "names gone, new and repeated are found");
  const before = ss.logical(); const report = run(false, "runCheckNames(SpreadsheetApp.getActiveSpreadsheet())").value.join("\n");
  ok(ss.logical() === before && report.includes(b.name) && report.includes("ижил нэртэй хэд хэдэн мөр"), "the name check only reads and names the sheet with the twin rows");
}

// ---------- a corrected name keeps its row; a repeated name warns; the hourly name watch ----------
{
  const v = (e) => run(true, e).value;
  const P = (n) => ss.getSheetByName(N(n)); const rowOf = (sh, label) => { for (let r = 1; r <= sh.rows.length; r++) if (sh.cell(r, 1).v === label) return r; return null; };
  const retype = (sh, row, value, oldValue) => { sh.cell(row, 1).v = value; state.editEvent = { range: sh.getRange(row, 1), oldValue }; const r = run(false, "onSheetEdit(editEvent)"); state.editEvent = undefined; return r; };
  const src = P(40), later = P(41), last = P(54);
  const row = rowOf(src, "B 9"); const laterRow = rowOf(later, "B 9"); later.cell(laterRow, 9).v = 5; later.cell(laterRow, 37).v = 3;
  const rows = later.rows.length; const out = retype(src, row, "B 9 шинэ", "B 9");
  ok(rowOf(later, "B 9") === null && rowOf(later, "B 9 шинэ") === laterRow && later.cell(laterRow, 9).v === 5 && later.cell(laterRow, 37).v === 3 && later.rows.length === rows, "a corrected name is renamed on the next sheet: same row, numbers kept");
  ok(rowOf(last, "B 9 шинэ") !== null && rowOf(last, "B 9") === null && out.logs.includes("солигдлоо"), "and on every later sheet, with a message");
  ok(rowOf(P(39), "B 9") !== null, "earlier sheets keep the old name");
  { const fresh = src.rows.length + 1; src.cell(fresh, 3).v = 10; src.cell(fresh, 4).v = 20; const dup = retype(src, fresh, "B 11", undefined);
    ok(src.cell(fresh, 1).v === "B 11 (2)" && dup.logs.includes("гэж нэрлэлээ") && rowOf(src, "B 11") !== fresh, "a typed name that is already on the sheet gets a number");
    ok(rowOf(later, "B 11 (2)") !== null && rowOf(last, "B 11 (2)") !== null, "and is carried to the later sheets under that name, as a new product");
    src.cell(fresh + 1, 3).v = 10; src.cell(fresh + 1, 4).v = 20; retype(src, fresh + 1, "b  11", undefined); ok(src.cell(fresh + 1, 1).v === "b  11 (3)", "the next one gets the next number"); }
  eq([v(`freeName("A", ["b"])`), v(`freeName("A ", ["a"])`), v(`freeName("A", ["a", "a (2)"])`)], ["A", "A (2)", "A (3)"], "a free name stays, a taken one gets the first free number");
  { const pr = (list) => JSON.stringify(list.map(([key, price]) => ({ key, label: key.toUpperCase(), price })));
    eq(v(`planNumbering([], [], ${pr([["g", 110], ["x", 5], ["g", 120]])})`), ["G", "X", "G (2)"], "on the first sheet the second row of a name gets a number");
    eq(v(`planNumbering(${pr([["g", 110], ["x", 5], ["g", 120]])}, ["G", "X", "G (2)"], ${pr([["g", 125], ["g", 110], ["x", 5], ["g", 300]])})`), ["G (2)", "G", "X", "G (3)"], "on the next sheet each row keeps the name of the row it continues, by price; a further row gets the next number");
    eq(v(`planNumbering(${pr([["g (2)", 120], ["g", 110]])}, ["G (2)", "G"], ${pr([["g", 110], ["g (2)", 120]])})`), ["G", "G (2)"], "names that are numbered already stay as they are"); }
  eq([v(`positionOfDate(new Date(2026, 0, 15))`), v(`positionOfDate(new Date(2026, 0, 16))`), v(`positionOfDate(new Date(2026, 11, 31))`)], [0, 1, 23], "a date gives its half-month");
  ok(v(`hashText("a")`) !== v(`hashText("b")`) && v(`hashText("a")`) === v(`hashText("a")`), "the fingerprint tells two reports apart");
  run(false, "runBackgroundJobs()"); state.mails = []; state.props.nameWatch = undefined;
  const far = P(53); far.cell(rowOf(far, "A 2"), 1).v = "A 2 өөр";
  eq(run(false, "runNameWatch(SpreadsheetApp.getActiveSpreadsheet(), new Date(2026, 0, 1))").value.length, 1, "the hourly watch reports a name that differs on a later sheet");
  ok(state.mails.length === 1 && state.mails[0][1] === "Барааны нэрийн шалгалт" && state.mails[0][2].includes("A 2 өөр"), "and e-mails the owner");
  run(false, "runNameWatch(SpreadsheetApp.getActiveSpreadsheet(), new Date(2026, 0, 1))"); ok(state.mails.length === 1, "the same finding is mailed once");
  eq(run(false, "runNameWatch(SpreadsheetApp.getActiveSpreadsheet(), new Date(2026, 11, 20))").value, ["нэрийн зөрүү арилсан"], "sheets that are over are not watched");
  far.cell(rowOf(far, "A 2 өөр"), 1).v = "A 2"; state.props.nameWatch = undefined;
  run(false, "checkProductNames()"); ok(state.mails.length === 2 && state.mails[1][1] === "Барааны нэрийн шалгалт", "the check by hand mails the whole report");
  run(false, "runBackgroundJobs()"); state.mails = [];
}

// ---------- one product, one name on every sheet ----------
{
  const v = (e) => run(true, e).value;
  const P = (n) => ss.getSheetByName(N(n)); const rowOf = (sh, label) => { for (let r = 1; r <= sh.rows.length; r++) if (sh.cell(r, 1).v === label) return r; return null; };
  eq(v(`buildNameMap([["Meiling 201l", ["miling  201L", "meiling201l"]]])`), { "miling 201l": "Meiling 201l", "meiling201l": "Meiling 201l" }, "every other spelling points at the one to keep");
  eq(v(`planRenames([{key:"miling 201l",label:"Miling 201l",row:5},{key:"x",label:"X",row:6}], {"miling 201l":"Meiling 201l"})`), { renames: [{ row: 5, from: "Miling 201l", to: "Meiling 201l" }], clashes: [] }, "a row with another spelling gets the kept one");
  eq(v(`planRenames([{key:"miling 201l",label:"Miling 201l",row:5},{key:"meiling 201l",label:"Meiling 201l",row:6}], {"miling 201l":"Meiling 201l"})`), { renames: [], clashes: ["Miling 201l"] }, "not when the kept spelling is on the sheet as well");
  run(false, "runBackgroundJobs()");
  const old3 = P(33), old4 = P(34), other = P(3);
  const r3 = rowOf(old3, "B 7"), r4 = rowOf(old4, "B 7"); old3.cell(r3, 1).v = "B7 huuchin"; old4.cell(r4, 1).v = "b 7 ХУУЧИН"; old3.cell(r3, 37).v = 21; old4.cell(r4, 2).v = 4;
  other.cell(rowOf(other, "B 7"), 1).v = "B7 huuchin";
  const fixes = `{"соронзон": [["B 7", ["b7 huuchin", "b 7 хуучин"]]]}`;
  const rowsBefore = ss.sheets.map(x => x.rows.length).join();
  const dry = run(true, `runUnifyNames(SpreadsheetApp.getActiveSpreadsheet(), ${fixes})`).value;
  ok(dry[0].includes("2 нүд солигдоно") && old3.cell(r3, 1).v === "B7 huuchin", "test mode only lists the names");
  const done = run(false, `runUnifyNames(SpreadsheetApp.getActiveSpreadsheet(), ${fixes})`).value;
  ok(done[0].includes("2 нүд солигдлоо") && old3.cell(r3, 1).v === "B 7" && old4.cell(r4, 1).v === "B 7", "both spellings become the kept name");
  ok(other.cell(rowOf(other, "B7 huuchin"), 1).v === "B7 huuchin", "the other shop has its own table");
  const sync = run(false, "runBackgroundJobs()");
  ok(ss.sheets.map(x => x.rows.length).join() === rowsBefore.replace(/^/, "") && rowOf(P(35), "B7 huuchin") === null && !sync.logs.includes("+1"), "the syncs add and remove no row afterwards");
  ok(old4.cell(r4, 2).v === 21, "the opening balance of the renamed product now follows the count before it");
  other.cell(rowOf(other, "B7 huuchin"), 1).v = "B 7"; run(false, "runUnifyNames(SpreadsheetApp.getActiveSpreadsheet(), {})"); run(false, "saveListsAnew(SpreadsheetApp.getActiveSpreadsheet())"); state.mails = [];
  ok(v("Object.keys(NAME_FIXES).join()") === "my home,соронзон" && v("NAME_FIXES['my home'].length") > 5, "the table of the real file is there");
}

// ---------- no two rows of one name: the whole year is numbered ----------
{
  const P = (n) => ss.getSheetByName(N(n)); const rowOf = (sh, label) => { for (let r = 1; r <= sh.rows.length; r++) if (sh.cell(r, 1).v === label) return r; return null; };
  run(false, "runNumberRepeated(SpreadsheetApp.getActiveSpreadsheet())"); run(false, "runBackgroundJobs()");
  const a = P(45), b = P(46), c = P(47);
  const twin = (sh, price, count) => { const r = rowOf(sh, "B 4"); sh.insertRowAfter(r); sh.cell(r + 1, 1).v = "B 4"; sh.cell(r + 1, 3).v = 1; sh.cell(r + 1, 4).v = price; sh.cell(r + 1, 37).v = count; return r + 1; };
  const ra = twin(a, 900, 7); const base = a.cell(ra - 1, 4).v;
  { const r = rowOf(b, "B 4"); b.insertRowBefore(r); b.cell(r, 1).v = "B 4"; b.cell(r, 3).v = 1; b.cell(r, 4).v = 950; }  // on the next sheet the dearer one stands first
  twin(c, 900, "");
  const dry = run(true, "runNumberRepeated(SpreadsheetApp.getActiveSpreadsheet())").value;
  ok(dry[0].includes("3 нүд солигдоно") && a.cell(ra, 1).v === "B 4", "test mode only lists the rows to number");
  const done = run(false, "runNumberRepeated(SpreadsheetApp.getActiveSpreadsheet())").value;
  ok(done[0].includes("3 нүд солигдлоо") && a.cell(ra, 1).v === "B 4 (2)" && a.cell(ra - 1, 1).v === "B 4" && a.cell(ra - 1, 4).v === base, "the second row of a name gets (2), the first keeps the name");
  ok(b.cell(rowOf(b, "B 4 (2)"), 4).v === 950 && c.cell(rowOf(c, "B 4 (2)"), 4).v === 900, "on later sheets the same product has the same number, wherever it stands");
  eq(run(false, "runNumberRepeated(SpreadsheetApp.getActiveSpreadsheet())").value, [], "a second run has nothing to do");
  { const rule = a.validations.find(x => x.r === 4 && x.c === 1 && x.w === 1);
    ok(rule && rule.h === a.maxRows - 3 - 0 || rule && rule.h > 200, "setup put a rule on column A from row 4 down");
    ok(rule.rule.list.includes("SUMPRODUCT") && rule.rule.allowInvalid === false && rule.rule.help.includes("аль хэдийн байна"), "it refuses a name that stands twice, with a message");
    ok(run(false, `installNameRule(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("${a.name}"))`).value === false, "and is not put there twice"); }
  eq(run(true, `planRenames([{key:"v1",label:"V1",row:5},{key:"v2",label:"V2",row:6}], {"v1":"Kept","v2":"Kept"})`).value, { renames: [{ row: 5, from: "V1", to: "Kept" }], clashes: ["V2"] }, "two spellings of one product on one sheet: only the first takes the kept name");
  const sync = run(false, "runBackgroundJobs()"); ok(sync.logs.includes("Нэмэгдсэн: 0, хасагдсан: 0") && rowOf(P(48), "B 4 (2)") === null && b.cell(rowOf(b, "B 4 (2)"), 2).v === 7, "afterwards the syncs add and remove no row, and the count follows the numbered row");
  state.mails = [];
}

// ---------- sheets are kept at the row limit ----------
{
  const s33 = ss.getSheetByName("Sheet33"); for (let i = 0; i < 7; i++) s33.insertRowAfter(20);
  ok(s33.maxRows === 407, "inserting rows grows the sheet");
  run(false, "runBackgroundJobs()");
  ok(s33.maxRows === 400 && s33.cell(5, 1).v === "A 1" && state.hours === 1, "hourly background job cuts it back to 400 without touching data");
}

// ---------- background job + guard ----------
{
  const s1 = ss.getSheetByName("Sheet1");
  const findRow = (s, label) => { for (let r = 1; r <= s.rows.length; r++) if (s.cell(r, 1).v === label) return r; return null; };
  const last = findRow(s1, "A 12");
  s1.insertRowAfter(last); s1.getRange(last + 1, 1).setValue("Шинэ бараа"); s1.getRange(last + 1, 4).setValue(999);
  s1.insertRowAfter(findRow(s1, "B 12")); const dv = findRow(s1, "B 12") + 1; s1.getRange(dv, 1).setValue("Cat C"); band(s1, dv, "#00ff00");
  run(false, "runBackgroundJobs()");
  ok([2, 12, 19, 24].every(i => { const s = ss.getSheetByName(`Sheet${i}`); const r = findRow(s, "Шинэ бараа"); return r > findRow(s, "A 12") && r < findRow(s, "Cat B") && s.cell(r, 4).v === 999; }), "new product carried to old AND new sheets of the same business, inside its own category");
  ok(findRow(ss.getSheetByName("Sheet31"), "Шинэ бараа") === null, "other business untouched");
  ok(s1.protections.filter(p => p.desc.startsWith("Ангиллын")).some(p => s1.cell(p.range.r, 1).v === "Cat C"), "background job protects the new divider");
  run(false, `onSheetStructureChange({changeType:"EDIT"})`); ok(state.mails.length === 0, "guard ignores ordinary edits");
  run(false, `onSheetStructureChange({changeType:"REMOVE_GRID"})`); ok(state.mails.length === 0, "guard silent while all sheets exist");
  ss.deleteSheet(ss.getSheetByName("Sheet7"));
  run(false, `onSheetStructureChange({changeType:"REMOVE_GRID"})`);
  ok(state.mails.length === 1 && state.mails[0][0] === "owner@example.com" && state.mails[0][2].includes("My/2026/4/15"), "guard e-mails the owner the missing sheet");
  const r = run(false, "setupYear()");
  ok(ss.getSheetByName("Sheet7") && ss.getSheetByName("Sheet7").cell(2, 1).v === "My home 2026 4 сарын 1-15" && r.logs.includes("My/2026/4/15: My/2026/3/31-ээс хуулж үүсгэлээ"), "setupYear re-creates a lost sheet in the right slot");
}

// ---------- a part that finds the lock taken makes a new appointment ----------
{
  const ssL = buildBook(); const stL = { triggers: [], mails: [] }; const runL = makeRunner(ssL, stL);
  const busy = (c) => c.replace("if (!lock.tryLock(SYNC_LOCK_WAIT_MS)) {", "if (true) {");
  const before = ssL.logical(); const r = runL(false, "setupYear()", busy);
  ok(ssL.logical() === before && r.logs.includes("1 минутын дараа өөрөө дахин эхэлнэ") && stL.triggers.map(t => t.getHandlerFunction()).join() === "setupYear" && stL.delays[stL.delays.length - 1] === 60000, "a setup part that finds the lock taken changes nothing and restarts itself in a minute");
  const dry = { triggers: [], mails: [] }; makeRunner(buildBook(), dry)(true, "setupYear()", busy);
  ok(dry.triggers.length === 0, "in test mode a locked-out setup schedules nothing");
  const done = runL(false, "setupYear()"); ok(done.logs.includes("БҮГД ДУУСЛАА") && !stL.triggers.some(t => t.getHandlerFunction() === "setupYear"), "and the next part, with the lock free, finishes the setup");
}

// ---------- one click: unfinished parts restart themselves ----------
{
  const ss3 = buildBook(); const st = { triggers: [], mails: [] }; const run3 = makeRunner(ss3, st);
  const noTime = (c) => c.replace("const SETUP_TIME_BUDGET_MS = 4 * 60 * 1000;", "const SETUP_TIME_BUDGET_MS = -1;");
  const handlers = () => st.triggers.map(t => t.getHandlerFunction()).sort().join();
  const r1 = run3(false, "setupYear()", noTime);
  ok(handlers() === "setupYear" && st.delays[st.delays.length - 1] === 60000 && r1.logs.includes("өөрөө үргэлжилнэ"), "unfinished part schedules exactly one continuation a minute later");
  ok(st.delays[0] === 420000 && st.props.setupRunCount === "1" && st.mails.length === 0, "a fallback is set while a part runs; no e-mail yet");
  run3(false, "setupYear()", noTime); ok(handlers() === "setupYear" && st.props.setupRunCount === "2", "second unfinished part: still one continuation");
  const r3 = run3(false, "setupYear()");
  ok(r3.logs.includes("БҮГД ДУУСЛАА") && !handlers().includes("setupYear") && st.props.setupRunCount === undefined, "finished part removes its trigger and counter");
  ok(st.mails.length === 1 && st.mails[0][2].includes("БҮГД ДУУСЛАА"), "owner is e-mailed when the automatic parts finish");
  st.props.setupRunCount = "14"; const r4 = run3(false, "setupYear()", noTime);
  ok(!handlers().includes("setupYear") && r4.logs.includes("гараар дахин"), "gives up restarting after the maximum number of parts");
  const st5 = { triggers: [], mails: [] }; makeRunner(buildBook(), st5)(true, "setupYear()", noTime);
  ok(st5.triggers.length === 0 && Object.keys(st5.props || {}).length === 0, "dry run never schedules or saves anything");
}

// ---------- a continuing part resumes at the step where the last one stopped ----------
{
  const ss4 = buildBook(); const st = { triggers: [], mails: [] }; const run4 = makeRunner(ss4, st);
  const first = run4(false, "setupYear()"); ok(first.logs.includes(" 1. Sheet-ийн нэр") && first.logs.includes("БҮГД ДУУСЛАА") && st.props.setupNextStep === undefined, "a finished setup leaves no resume point");
  st.props = { setupRunCount: "1", setupNextStep: "dev|11|0" };
  const resumed = run4(false, "setupYear()");
  ok(!resumed.logs.includes(" 1. Sheet-ийн нэр") && !resumed.logs.includes(" 11. Ангиллыг") && resumed.logs.includes(" 12. Зааглагчийн") && resumed.logs.includes("зогссон газраас үргэлжилж байна") && resumed.logs.includes("БҮГД ДУУСЛАА"), "the next part starts at step 12, not at step 1");
  st.props = { setupRunCount: "1", setupNextStep: "older-code|9|0" };
  ok(run4(false, "setupYear()").logs.includes(" 1. Sheet-ийн нэр"), "a resume point written by an older copy of the code is ignored: all steps run again");
  st.props = { setupNextStep: "dev|9|0" };
  ok(run4(false, "setupYear()").logs.includes(" 1. Sheet-ийн нэр"), "a run started by hand after a finished setup checks everything from step 1");
  const noTime = (c) => c.replace("const SETUP_TIME_BUDGET_MS = 4 * 60 * 1000;", "const SETUP_TIME_BUDGET_MS = -1;");
  st.props = { setupRunCount: "1", setupNextStep: "dev|9|7" }; run4(false, "setupYear()", noTime);
  ok(st.props.setupNextStep === "dev|9|7" && st.props.setupRunCount === "2", "a part that runs out of time keeps the resume point for the next one");
}

// ---------- a long step stops between two sheets when the time is used up ----------
{
  const r = makeRunner(new Spreadsheet(), { triggers: [], mails: [] });
  eq(r(true, `forEachUntil([1,2,3], Date.now() + 60000, (n) => n === 2 ? null : "s" + n)`).value, { lines: ["s1", "s3"], complete: true }, "all items done in time");
  eq(r(true, `forEachUntil([1,2,3], Date.now() - 1, (n) => "s" + n)`).value, { lines: ["хугацаа дууслаа - дараагийн хэсэгт үргэлжилнэ"], complete: false, nextItem: 0 }, "no time left: nothing is started, the step is reported unfinished");
  eq(r(true, `forEachUntil([1,2,3,4], Date.now() + 60000, (n) => "s" + n, 2)`).value, { lines: ["s3", "s4"], complete: true }, "a step that continues starts at the item where it stopped");
  eq(r(true, `forEachUntil([1,2,3,4], Date.now() - 1, (n) => "s" + n, 3)`).value.nextItem, 3, "and reports that same item again if there is still no time");
  eq(r(true, `forEachUntil([], Date.now() + 60000, (n) => n)`).value, { lines: ["хийх зүйлгүй"], complete: true }, "nothing to do is still a finished step");
  const ss5 = buildBook(); const st = { triggers: [], mails: [] }; const run5 = makeRunner(ss5, st);
  const slow = (c) => c.replace("function installSummary(sheet) {", "function installSummary(sheet) { ss.summaryCalls = (ss.summaryCalls || 0) + 1; if (ss.summaryCalls === 3) { const end = Date.now() + 60; while (Date.now() < end) {} }").replace("const SETUP_TIME_BUDGET_MS = 4 * 60 * 1000;", "const SETUP_TIME_BUDGET_MS = 40;");
  const parts = [run5(false, "setupYear()", slow).logs];
  while (!parts[parts.length - 1].includes("БҮГД ДУУСЛАА") && parts.length < 400) { st.props.setupRunCount = "1"; parts.push(run5(false, "setupYear()", slow).logs); }
  ok(parts.length > 1 && parts.length < 400 && parts[parts.length - 1].includes("БҮГД ДУУСЛАА"), "a setup cut into short parts still finishes (" + parts.length + " parts)");
  ok(parts.some(log => log.includes("хугацаа дууслаа - дараагийн хэсэгт үргэлжилнэ") && log.includes("ДУУСААГҮЙ")), "a long step stops between two sheets instead of running past the limit");
  ok(ss5.sheets.filter(x => /^(My|Sor)\/2026\//.test(x.name)).every(x => x.cell(3, 44).f === "=SUMPRODUCT(D4:D,AN4:AN)" && x.cell(3, 38).f.startsWith("=ARRAYFORMULA(") && x.protections.some(p => p.range.r === 1)), "and every sheet got its totals, formulas and protection");
}

// ---------- "is this a product sheet" is asked once per sheet in a run ----------
{
  const ss6 = buildBook(); const run6 = makeRunner(ss6, { triggers: [], mails: [] });
  run6(false, "setupYear()"); ss6.sheets.forEach(x => { x.reads = 0; });
  const inChains = run6(false, "(() => { let found = 0; for (let i = 0; i < 5; i++) found = findChains(ss).reduce((n, c) => n + c.sheets.length, 0); return found; })()").value;
  ok(inChains === 48, "48 product sheets are found, asked five times over");
  ok(ss6.sheets.filter(x => /^(My|Sor)\/2026\//.test(x.name)).every(x => x.reads === 1), "each sheet is read once for that question, however often it is asked");
}

// ---------- name collision stops ----------
{
  const ss2 = buildBook(); const note = ss2.insertSheet("Sor/2026/12/31"); note.cell(1, 1).v = "тэмдэглэл";
  const st2 = { triggers: [], mails: [] }; const r = makeRunner(ss2, st2)(false, "setupYear()");
  ok(st2.triggers.length === 0, "a blocked setup does not restart itself");
  ok(r.logs.includes("ЗОГСЛОО") && r.logs.includes("ДУУСААГҮЙ") && note.cell(1, 1).v === "тэмдэглэл" && ss2.sheets.includes(note) && ss2.sheets.some(x => x.name === "Sheet54"), "a non-empty sheet holding a needed name is never deleted or renamed; setup stops");
}

// ---------- тооллогын хуудас (stock-sheets.gs) ----------
{
  const ss7 = buildBook(); const st7 = { triggers: [], mails: [] }; const run7 = makeRunner(ss7, st7);
  const before = ss7.logical();
  const dry = run7(true, "runInstallStockSheets(ss)").value;
  ok(ss7.logical() === before && ss7.getSheetByName("Sheet100*") === null, "DRY_RUN: nothing changes, Sheet100* is not made");
  ok(dry[0].includes("Sheet100: одоо байгааг цэвэрлээд") && dry[1].includes("Sheet100*: шинээр үүсгэж"), "DRY_RUN log says what would happen to each sheet");
  run7(false, "runInstallStockSheets(ss)");
  const again = run7(false, "runInstallStockSheets(ss)").value;
  ok(again.filter(line => !line.includes("олдсонгүй")).every(line => line.includes("аль хэдийн бэлэн")), "a second run changes nothing");
  const stock = ss7.getSheetByName("Sheet100");
  stock.cell(10, 1).bg = "#ff0000";
  ok(run7(false, "reconcileSheet(ss.getSheetByName('Sheet100'))").value === null && stock.cell(10, 1).bg === "#ff0000" && stock.protections.length === 1, "the colour / divider pass leaves a stock sheet alone");
  ok(run7(false, "trimSheet(ss.getSheetByName('Sheet100*'))").value.includes("мөр хасахгүй") && run7(false, "protectHeaderRows(ss.getSheetByName('Sheet100'))").value.includes("алгаслаа"), "row trimming and header protection skip the stock sheets");
  ok(!run7(false, "listCalcSheets(ss).map(x => x.getName())").value.some(name => name.startsWith("Sheet100")), "no AL:AN formulas on the stock sheets");
  const f = run7(true, "[buildStockFormulas(CHAINS[1]).list, buildStockFormulas(CHAINS[0]).list]").value;
  ok(f[0].includes(`IF(DAY(t)<=15,"Sor/"&YEAR(EOMONTH(t,-1))&"/"&MONTH(EOMONTH(t,-1))&"/"&DAY(EOMONTH(t,-1)),"Sor/"&YEAR(t)&"/"&MONTH(t)&"/15")`), "the sheet before: last month's 2nd half on days 1-15, this month's 1st half after");
  ok(f[0].includes("k,IF(DAY(t)<=15,DAY(t),MIN(15,DAY(t)-15))") && f[0].includes("upto,--(SEQUENCE(15)<=k)"), "only the day columns up to today count; day 31 stays in the last column");
  ok(f[0].includes('INDIRECT(s&"E4:S")') && f[0].includes('INDIRECT(s&"T4:T")') && f[0].includes('INDIRECT(s&"U4:AI")') && f[0].includes('INDIRECT(s&"AJ4:AJ")'), "income E:S, transfer in T, expense U:AI, transfer out AJ");
  ok(f[0].includes('COUNT(INDIRECT(p&"AK4:AK"))>0') && f[0].includes('VLOOKUP(nm,INDIRECT(p&"A4:AL"),38,FALSE)'), "an empty opening balance falls back to the final balance only where the sheet before was not counted");
  ok(f[1].includes('tab,"My/"') && !f[1].includes('"Sor/'), "Sheet100* formula reads My home only");
  const log = run7(false, "setupYear()").logs;
  ok(log.includes("16. Тооллогын хуудас") && log.includes("17. Автомат ажил"), "setupYear has the stock-sheet step before automation");
}

// ---------- a script writes no name the list lacks (the owner, 2026-10-25) ----------
{
  const ssL = buildBook(); const stL = { triggers: [], mails: [] }; const runL = makeRunner(ssL, stL);
  runL(false, "setupYear()");
  runL(false, "runSync(ss)");
  const rowOf = (sh, label) => { for (let r = 1; r <= sh.rows.length; r++) if (sh.cell(r, 1).v === label) return r; return null; };
  const has = (sh, label) => rowOf(sh, label) !== null;
  eq(runL(true, "listNameKeys(ss)").value, null, "no list yet: no name is refused");
  const list = ssL.insertSheet("Бараа");
  [["Cat A", ""], ["Listed new", "Соронзон"]].forEach(([name, shop], i) => { list.cell(4 + i, 1).v = name; list.cell(4 + i, 2).v = shop; });
  const sep = ssL.getSheetByName("Sor/2026/9/30");
  const later = ["Sor/2026/10/15", "Sor/2026/10/31", "Sor/2026/12/31"].map((name) => ssL.getSheetByName(name));
  const r = rowOf(sep, "A 3");
  sep.insertRowAfter(r); sep.cell(r + 1, 1).v = "Listed new"; sep.cell(r + 1, 4).v = 70;
  sep.insertRowAfter(r + 1); sep.cell(r + 2, 1).v = "Off list Z"; sep.cell(r + 2, 4).v = 80;
  const report = runL(false, "runSync(ss).lines.join('\\n')").value;
  ok(later.every((sh) => has(sh, "Listed new")), "a new product on the list goes on to the later sheets");
  ok(later.every((sh) => !has(sh, "Off list Z")) && has(sep, "Off list Z"), "a name the list lacks is not written anywhere by the script; the typed row stays");
  ok(report.includes('"Off list Z"') && report.includes("жагсаалтад алга"), "and the log says why");
  list.cell(4, 1).v = "#REF!"; list.cell(5, 1).v = "";
  eq(runL(true, "listNameKeys(ss)").value, null, "a copy whose IMPORTRANGE failed judges nothing (no list read is no reason to refuse)");
}

// ---------- a sheet's products carried into the later sheets that lack them (the owner, 2026-10-28) ----------
// Tos/2026/11 and 12 held about 145 of October's 290 products: the hourly sync carries only what it saw
// change, so a product a later sheet never had stays missing. carryProductsForward puts them in.
{
  const ssC = buildBook(); const stC = { triggers: [], mails: [] }; const runC = makeRunner(ssC, stC);
  runC(false, "setupYear()");
  runC(false, "runSync(ss)");
  const rowOf = (sh, label) => { for (let r = 1; r <= sh.rows.length; r++) if (sh.cell(r, 1).v === label) return r; return null; };
  const has = (sh, label) => rowOf(sh, label) !== null;
  const sep = ssC.getSheetByName("Sor/2026/9/30");
  const later = ["Sor/2026/10/15", "Sor/2026/10/31", "Sor/2026/12/31"].map((name) => ssC.getSheetByName(name));
  later.forEach((sh) => { sh.deleteRow(rowOf(sh, "A 3")); sh.deleteRow(rowOf(sh, "B 5")); }); // made from a shorter list
  const dec = later[2];
  dec.insertRowAfter(rowOf(dec, "B 12")); dec.cell(rowOf(dec, "B 12") + 1, 1).v = "Only later";
  sep.cell(rowOf(sep, "A 3"), 3).v = 111; sep.cell(rowOf(sep, "A 3"), 4).v = 222;
  const list = ssC.insertSheet("Бараа");
  ["Cat A", ...Array.from({ length: 12 }, (_, i) => `A ${i + 1}`), "Cat B", "Only later", ...Array.from({ length: 12 }, (_, i) => `B ${i + 1}`).filter((name) => name !== "B 5")]
    .forEach((name, i) => { list.cell(4 + i, 1).v = name; list.cell(4 + i, 2).v = "Соронзон"; });
  const carry = (dry, deadline = "Date.now() + 1e9") => runC(dry, `runCarryProducts(ss, periodPosition(2026, 9, 1), ${deadline}).join('\\n')`);
  const before = ssC.logical();
  const dry = carry(true);
  ok(dry.error === null || dry.error === undefined, "the carry runs: " + dry.error);
  ok(ssC.logical() === before && /Sor\/2026\/10\/15: 1 бараа нэмнэ \(туршилт\): A 3/.test(dry.value), "DRY_RUN changes nothing and names what each later sheet lacks");
  const live = carry(false).value;
  ok(later.every((sh) => has(sh, "A 3")), "a product of the sheet a later sheet lacks is put in there");
  ok(later.every((sh) => sh.cell(rowOf(sh, "A 3"), 3).v === 111 && sh.cell(rowOf(sh, "A 3"), 4).v === 222), "with its cost and price");
  ok(later.every((sh) => rowOf(sh, "A 3") > rowOf(sh, "Cat A") && rowOf(sh, "A 3") < rowOf(sh, "Cat B")), "in its category");
  ok(later.every((sh) => sh.cell(rowOf(sh, "A 3"), 7).v === ""), "and none of the month's numbers");
  ok(later.every((sh) => !has(sh, "B 5")) && /"B 5"/.test(live) && /жагсаалтад алга/.test(live), "a name the list lacks is not written, only named");
  ok(has(dec, "Only later"), "nothing is taken away from a later sheet");
  ok(runC(true, "loadSnapshots(ss)").value["Sor/2026/12/31"].includes("a 3"), "the sync's saved lists know the new rows, so it does not carry them again");
  const after = ssC.logical();
  carry(false);
  ok(ssC.logical() === after, "a second run adds nothing");
  dec.deleteRow(rowOf(dec, "A 4"));
  const late = carry(false, "0").value;
  ok(!has(dec, "A 4") && /цаг дууслаа/.test(late), "past the deadline nothing more is put in, and the log says to run it again");
  carry(false);
  ok(has(dec, "A 4"), "the next run puts it in");
}

// ---------- one product list for both shops (product-list.gs) ----------
{
  const ssP = buildBook(); const stP = { triggers: [], mails: [] }; const runP = makeRunner(ssP, stP);
  runP(false, "setupYear()");
  const vP = (e) => runP(true, e).value;
  const shopSheets = (prefix) => ssP.sheets.filter(x => x.name.startsWith(prefix + "/2026/"));
  const rowOf = (sh, label) => { for (let r = 1; r <= sh.rows.length; r++) if (sh.cell(r, 1).v === label) return r; return null; };
  const has = (sh, label) => rowOf(sh, label) !== null;
  // pure: the owner's yes / no table
  const table = [
    ["", "", "A2 typo", 0, "A 2", 0, "", "A 2", "тийм", "", ""],
    ["", "", "B 3", 0, "Bee 3", 0, "", "Bee 3", "тийм", " B 3 new ", ""],
    ["", "", "A 5", 0, "A 6", 0, "", "A 6", "үгүй", "", ""],
  ];
  const plan = vP(`planNameDecisions(${JSON.stringify(table)})`);
  eq(plan.undecided, [], "every row decided");
  eq(plan.fixes["my home"], [["A 2", ["a2 typo"]], ["B 3 new", ["b 3"]]], "My home: its spelling -> the chosen name (H, or J when filled)");
  eq(plan.fixes["соронзон"], [["B 3 new", ["bee 3"]]], "Соронзон: a name that already is the chosen one needs no fix");
  eq(vP(`planNameDecisions(${JSON.stringify([table[0], ["", "", "X", 0, "Y", 0, "", "X", "", "", ""]])}).undecided`), [4], "an empty answer is reported with its row number");
  ok(vP(`planNameDecisions(${JSON.stringify([table[0], ["", "", "X", 0, "Y", 0, "", "A 2", "тийм", "", ""]])}).problems[0]`).includes("A 2"), "two pairs given one name is a problem");
  ok(vP(`planNameDecisions(${JSON.stringify([["", "", "X", 0, "Y", 0, "", "X", "магадгүй", "", ""]])}).problems[0]`).includes("магадгүй"), "an answer that is not тийм / үгүй is a problem");
  // pure: the list is built from the current sheets of both shops
  const list = vP(`planProductList(
    { categories: [{ key: "cat a", label: "Cat A" }], products: [{ key: "a 2", label: "A 2", categoryKey: "cat a" }, { key: "bee 3", label: "Bee 3", categoryKey: "cat a" }] },
    { categories: [{ key: "cat a", label: "  Cat   A " }, { key: "плитка", label: "Плитка" }], products: [{ key: "a2 typo", label: "A2 typo", categoryKey: "cat a" }, { key: "b 3", label: "B 3", categoryKey: "cat a" }, { key: "only", label: "Only My", categoryKey: "плитка" }] },
    { "соронзон": { "bee 3": "B 3 new" }, "my home": { "a2 typo": "A 2", "b 3": "B 3 new" } })`);
  eq(list, [["Cat A", ""], ["A 2", "Хоёулаа"], ["B 3 new", "Хоёулаа"], ["Плитка", ""], ["Only My", "My home"]], "categories as heading rows, one row per product, which shop has it");
  eq(vP(`planListEdit("A 7", "A 7 renamed", ["Cat A", "A 7", "A 8"])`), { rename: true, from: "a 7", to: "A 7 renamed" }, "a product renamed in the list is renamed in the shops");
  ok(!vP(`planListEdit("A 7", "a  8", ["Cat A", "A 7", "A 8"])`).rename, "a name already on the list is not taken by another product");
  ok(!vP(`planListEdit("Хөргөгч", "Хөргөгч 2", ["Хөргөгч", "A 7"])`).rename && !vP(`planListEdit("", "New", ["A 7"])`).rename, "a category heading or a new row renames nothing");

  // the real flow on the fake book
  shopSheets("My").forEach(sh => { const r = rowOf(sh, "A 2"); if (r) sh.cell(r, 1).v = "A2 typo"; });
  shopSheets("Sor").forEach(sh => { const r = rowOf(sh, "B 3"); if (r) sh.cell(r, 1).v = "Bee 3"; });
  runP(false, "saveListsAnew(ss)");
  const tab = ssP.insertSheet("_Нэр тулгалт"); tab.cell(2, 1).v = "№";
  table.forEach((row, i) => row.forEach((value, j) => { tab.cell(3 + i, 1 + j).v = value; }));
  const sep = "new Date(2026, 8, 20)";
  const current = (prefix) => ssP.getSheetByName(`${prefix}/2026/9/30`);
  { const my = current("My"); const r = rowOf(my, "A 12"); my.insertRowAfter(r); my.cell(r + 1, 1).v = "Only My"; my.cell(r + 1, 4).v = 50; }
  const before = ssP.logical();
  const dry = runP(true, `runSetupProductList(ss, ${sep})`).value.join("\n");
  ok(ssP.logical() === before, "DRY_RUN: nothing changes");
  ok(dry.includes("нэр нэгтгэх") && dry.includes("Бараа") && dry.includes("dropdown"), "DRY_RUN log names every step");
  tab.cell(5, 9).v = "";
  const refused = runP(false, `runSetupProductList(ss, ${sep})`).value.join("\n");
  ok(refused.includes("шийдээгүй") && refused.includes("5") && ssP.getSheetByName("Бараа") === null && has(current("My"), "A2 typo"), "an undecided row stops everything");
  tab.cell(5, 9).v = "үгүй";
  const done = runP(false, `runSetupProductList(ss, ${sep})`).value.join("\n");
  ok(shopSheets("My").every(sh => !has(sh, "A2 typo") && !has(sh, "B 3")) && has(current("My"), "A 2") && has(current("My"), "B 3 new"), "My home gets the chosen names on every sheet of the year");
  ok(shopSheets("Sor").every(sh => !has(sh, "Bee 3")) && has(current("Sor"), "B 3 new"), "Соронзон too");
  const master = ssP.getSheetByName("Бараа");
  const masterNames = []; for (let r = 4; r <= master.rows.length; r++) if (master.cell(r, 1).v !== "") masterNames.push(master.cell(r, 1).v);
  ok(masterNames.includes("Cat A") && masterNames.includes("A 2") && masterNames.includes("B 3 new") && masterNames.includes("Only My") && !masterNames.includes("A2 typo"), "the list holds the chosen names");
  eq(masterNames.length, new Set(masterNames.map(x => x.toLowerCase())).size, "each name once");
  ok(master.cell(rowOf(master, "Only My"), 2).v === "My home" && master.cell(rowOf(master, "A 1"), 2).v === "Хоёулаа", "the list says which shop has the product");
  const rule = current("Sor").validations.filter(x => x.r === 4 && x.c === 1).pop(); // the newest rule wins, as in Sheets
  ok(rule && rule.rule.list.getSheet().getName() === "Бараа" && rule.rule.list.getRow() === 4 && rule.rule.dropdown === true, "column A of the year sheets is a dropdown of the list");
  ok(rule.rule.allowInvalid === !vP("PRODUCT_LIST_STRICT"), "a name not on the list is refused only in strict mode");
  ok(done.includes("ХИЙГДЛЭЭ"), "the log says it was done");
  ok(runP(false, `installNameRule(ss.getSheetByName("${current("My").name}"))`).value === false, "the dropdown is not put there twice");
  { const values = () => JSON.stringify(master.rows.map(row => [row[1] && row[1].v, row[2] && row[2].v]).filter(([a, b]) => a || b));
    const snap = values(); runP(false, "runBackgroundJobs()");
    ok(values() === snap && master.rows.every(row => Object.values(row).every(x => x.f === "")), "the hourly jobs leave the names of the list alone, no formulas");
    const height = () => current("My").validations.filter(x => x.r === 4 && x.c === 1).pop().rule.list.getNumRows();
    const was = height(); master.insertRowsAfter(master.maxRows, 40); for (let r = master.getLastRow() + 1; r <= 420; r++) master.cell(r, 1).v = `Шинэ ${r}`;
    runP(false, "runBackgroundJobs()");
    ok(master.maxRows === 440 && height() === 437 && height() > was, "a list grown past the row limit reaches the dropdown within the hourly job"); }
  master.cell(4, 3).v = "тэмдэглэл";
  runP(false, `runSetupProductList(ss, ${sep})`);
  ok(master.cell(4, 3).v === "тэмдэглэл" && ssP.sheets.filter(x => x.name === "Бараа").length === 1, "a second run leaves the existing list alone");
  // a name corrected in the list follows in both shops
  const r7 = rowOf(master, "A 7"); master.cell(r7, 1).v = "A 7 renamed";
  stP.editEvent = { range: master.getRange(r7, 1), oldValue: "A 7", value: "A 7 renamed" };
  runP(false, "onSheetEdit(editEvent)");
  ok([...shopSheets("My"), ...shopSheets("Sor")].every(sh => !has(sh, "A 7")) && has(current("My"), "A 7 renamed") && has(current("Sor"), "A 7 renamed"), "renamed in the list -> renamed on every sheet of both shops");
  { const log = ssP.getSheetByName("_Нэр солилт");
    ok(log !== null && log.cell(2, 2).v === "a 7" && log.cell(2, 3).v === "A 7 renamed" && /^\d{4}-\d{2}-\d{2}/.test(log.cell(2, 1).v), "the rename is written to the rename log, for Тосгон to follow");
    eq(runP(false, "runListRenamesFromSource(ss)").value, [], "the workbook that owns the list follows no other list");
    ok(vP("PRODUCT_LIST_STRICT") === true, "hard rule: a name that is not on the list is refused"); }
  const r8 = rowOf(master, "A 8"); master.cell(r8, 1).v = "A 9";
  stP.editEvent = { range: master.getRange(r8, 1), oldValue: "A 8", value: "A 9" };
  runP(false, "onSheetEdit(editEvent)");
  ok(has(current("My"), "A 8") && has(current("My"), "A 9"), "a name already on the list renames nothing in the shops");
  master.cell(r8, 1).v = "A 8"; // put the list back as it was

  // hourly: names on the sheets from this half-month on that the list does not have
  eq(vP(`planListWatch([
      { shop: "my home", products: [{ label: "A 1" }, { label: "Zzz" }, { label: "A 1 (2)" }, { label: "Zzz" }], categories: [{ label: " Cat  A" }, { label: "Cat Z" }] },
      { shop: "соронзон", products: [{ label: "a  1" }], categories: [] }], ["a 1", "cat a"])`),
    ["My home\n  жагсаалтад байхгүй бараа: Zzz\n  жагсаалтад байхгүй ангилал: Cat Z\n  давхардаж дугаарлагдсан нэр: A 1 (2)"],
    "per shop: products, categories and numbered names the list lacks, each once; a shop with none is not listed");
  eq(vP(`planListWatch([{ shop: "соронзон", products: [{ label: "A 1 (2)" }], categories: [] }], ["a 1 (2)"])`), [], "a numbered name on the list is fine");
  {
    const sor = current("Sor"); const r = rowOf(sor, "A 3"); sor.insertRowAfter(r); sor.cell(r + 1, 1).v = "Шинэ бараа X"; sor.cell(r + 1, 4).v = 70;
    stP.mails = []; stP.props = {};
    const dryW = runP(true, `runListWatch(ss, ${sep})`).value;
    ok(dryW.length === 1 && dryW[0].includes("1 дэлгүүрт") && stP.mails.length === 0, "DRY_RUN: says a mail would go, sends none");
    const sent = runP(false, `runListWatch(ss, ${sep})`).value;
    ok(sent[0].includes("явууллаа") && stP.mails.length === 1 && stP.mails[0][1] === "Бараа жагсаалтад байхгүй нэр" && stP.mails[0][2].includes("Соронзон\n  жагсаалтад байхгүй бараа: Шинэ бараа X"), "the owner gets one mail with the names, by shop");
    eq(runP(false, `runListWatch(ss, ${sep})`).value, [], "the same finding is not mailed again");
    ok(stP.mails.length === 1, "so no second mail");
    eq(runP(false, `runListWatch(ss, new Date(2026, 9, 20))`).value.length, 1, "sheets before this half-month do not count: the finding on 9/30 is gone in late October");
    ok(runP(false, `runListWatch(ss, ${sep})`).value[0].includes("явууллаа"), "it comes back when it is there again");
    sor.deleteRow(r + 1);
    ok(runP(false, `runListWatch(ss, ${sep})`).value[0].includes("арилсан") && !("listWatch" in stP.props), "once fixed, the watch says so and forgets the finding");
    const noList = buildBook(); const stN = { triggers: [], mails: [] };
    eq(makeRunner(noList, stN)(false, `runListWatch(ss, ${sep})`).value, [], "no list yet: nothing to watch");
    ok(vP("BACKGROUND_TASKS.some(task => String(task).includes('runListWatch'))"), "the hourly job runs the watch");
  }

  // Тосгон: its names against the list, as a yes / no table (tosgon-names.gs)
  eq(vP(`["Haeir 103L", "Haier  103l", "Smart10.9l", "хөнгөн цагаан 11Л"].map(matchKey)`), ["haier103", "haier103", "smart10.9", "xөhгөhцaгaah11"], "one key for spellings that differ only in spaces, case, l and known typos");
  eq(vP(`[pairLevel("Haeir 103l", "Haier 103l").level, pairLevel("8037 ус", "Ус буцалгагч 8037").level, pairLevel("Haier 11l", "Haier 110-118").level, pairLevel("Haier 11l", "Meiling 11l").level, pairLevel("Smart 7.2l", "Smart 7.6l").level]`), [3, 2, 1, 0, 0], "same key 3, same numbers and a shared word 2, a number that starts the other 1, otherwise 0");
  const tPlan = vP(`planTosgonPairs(
    [{ name: "A 1", category: "Cat A", price: 10 }, { name: "Haeir 103l", category: "Cat A", price: 20 }, { name: "Haier  103 L", category: "Cat A", price: 20 }, { name: "Tosgon only", category: "Cat B", price: "" }],
    [{ name: "A 1", shops: "Хоёулаа" }, { name: "Haier 103l", shops: "Соронзон" }])`);
  eq(tPlan.same, 1, "a name already on the list needs no row");
  eq(tPlan.rows.map(r => [r[2], r[4], r[5], r[7]]), [["Haeir 103l", "Haier 103l", "Соронзон", "Haier 103l"], ["Haier  103 L", "Haier 103l", "Соронзон", "Haier 103l"], ["Tosgon only", "", "", "Tosgon only"]],
    "the list's name suggested; a second Тосгон spelling of the same list name is shown too (both тийм is then refused); a name with no pair last, as a new product");
  eq(vP(`[spellingDifference("Panasonica 10 abto", "Panasonica 10 авто"), spellingDifference("Haier 90 - 197", "Haier 90-197"), spellingDifference("Cmidea 360L", "Cmedia 360l"), spellingDifference("8037 ус", "Ус буцалгагч 8037")]`), ["латин / кирилл үсэг: abto → авто", "", "Cmidea → Cmedia", "— → буцалгагч"], "the words that differ, as written; spaces, case and l are not a difference, Latin for Cyrillic is said");
  eq(vP(`[spellingSimilarity("Panasonica 160l", "Pnasonica 160l"), spellingSimilarity("Aucma 341", "Aucma 341L"), spellingSimilarity("12 хөлдөөгч 120", "Хөлдөөгч 120л")].map(x => Math.round(x * 100))`), [92, 100, 85], "% alike: the share of letters that need no change");
  ok(vP(`describeTosgonPair("Haeir 103l", "Haier 103l", MATCH_LEVEL_LABELS[3])`).includes("75% төстэй · Haeir → Haier"), "G: why, how alike in %, what differs");
  eq(vP(`[tosgonRowTone(MATCH_LEVEL_LABELS[3] + "\\n100%"), tosgonRowTone(MATCH_CONTESTED_LABEL), tosgonRowTone("")]`), ["#d9ead3", "#f4cccc", "#ffffff"], "the row colour follows the kind G starts with");
  ok(tPlan.rows[1][6].includes("Давхар") && tPlan.rows[2][6].includes("олдсонгүй") && tPlan.rows.map(r => r[0]).join() === "1,2,3", "each row says why; rows are numbered");
  {
    const tos9 = ssP.insertSheet("Tos/2026/9"); tos9.cell(3, 1).v = "Барааны нэрс";
    tos9.cell(4, 1).v = "Cat A"; band(tos9, 4, "#f6b26b");
    [["A 1", 5], ["A  2l", 6], ["B 3new", 7], ["Tosgon only", 8]].forEach(([name, price], i) => { tos9.cell(5 + i, 1).v = name; tos9.cell(5 + i, 2).v = 1; tos9.cell(5 + i, 4).v = price; });
    ssP.insertSheet("Tos/2026/10").cell(3, 1).v = "Барааны нэрс"; // made, nothing typed yet
    const tos11 = ssP.insertSheet("Tos/2026/11"); tos11.cell(5, 1).v = "Future"; tos11.cell(5, 4).v = 1;
    const tosBefore = [tos9.logical(), tos11.logical()].join();
    const oct = "new Date(2026, 9, 19)";
    const before = ssP.logical();
    const dryT = runP(true, `runTosgonNameTable(ss, ss, ${oct})`).value.join("\n");
    ok(ssP.logical() === before && dryT.includes("Tos/2026/9") && dryT.includes("_Тосгон нэр тулгалт"), "DRY_RUN: the newest month with names up to today is read, nothing is written");
    runP(false, `runTosgonNameTable(ss, ss, ${oct})`);
    const tTab = ssP.getSheetByName("_Тосгон нэр тулгалт");
    const tRows = []; for (let r = 3; r <= tTab.rows.length; r++) if (tTab.cell(r, 3).v !== "") tRows.push([tTab.cell(r, 3).v, tTab.cell(r, 5).v]);
    eq(tRows, [["A  2l", "A 2"], ["B 3new", "B 3 new"], ["Tosgon only", ""]], "the table pairs the Тосгон names with the list's; a future month is not read");
    ok(tTab.cell(2, 9).v.includes("тийм") && tTab.validations.some(v => v.c === 9 && v.r === 3), "same columns as the shops' table, with a тийм / үгүй list in I");
    ok([tos9.logical(), tos11.logical()].join() === tosBefore, "the Тосгон sheets are only read");
    tTab.cell(3, 9).v = "тийм";
    ok(runP(false, `runTosgonNameTable(ss, ss, ${oct})`).value.join("\n").includes("хөндсөнгүй") && tTab.cell(3, 9).v === "тийм", "an existing table with answers is never overwritten");
    ok(tTab.cell(3, 1).bg === "#d9ead3" && tTab.cell(5, 1).bg === "#efefef" && tTab.cell(3, 7).v.includes("% төстэй") && tTab.cell(5, 7).v === "Жагсаалтад олдсонгүй - шинэ бараа уу?", "rows are coloured by how sure the pair is; G says how alike and what differs");
    ok(tTab.cfRules.length === 2 && tTab.cfRules[0].formula === '=$I3="тийм"' && tTab.cell(3, 4).nf === "#,##0 ₮" && tTab.merges.length === 1, "тийм / үгүй colour I, the price reads as money, the title spans the table");
    // a table made before the colours: plain G, the price as text; restyle keeps every answer
    tTab.cell(3, 7).v = "Үсгийн алдаа / үгийн дараалал"; tTab.cell(3, 4).v = "1200000"; tTab.cell(3, 10).v = "A 2 зөв"; tTab.cell(3, 11).v = "тэмдэглэл";
    [3, 4, 5].forEach(r => { for (let c = 1; c <= 11; c++) tTab.cell(r, c).bg = ""; });
    const answersBefore = [3, 4, 5].map(r => [9, 10, 11].map(c => tTab.cell(r, c).v)).join();
    const dryR = runP(true, "runRestyleTosgonTable(ss)").value.join("\n");
    ok(dryR.includes("3 мөрийг") && tTab.cell(3, 1).bg === "", "DRY_RUN restyle only reports");
    runP(false, "runRestyleTosgonTable(ss)");
    ok([3, 4, 5].map(r => [9, 10, 11].map(c => tTab.cell(r, c).v)).join() === answersBefore, "restyle never touches I, J, K");
    ok(tTab.cell(3, 4).v === 1200000 && tTab.cell(3, 1).bg !== "" && tTab.cell(3, 7).v.includes("% төстэй") && tTab.cell(5, 1).bg === "#efefef", "restyle: the price becomes a number, G gets the % and the difference, rows get their colour");
    eq(makeRunner(buildBook(), { triggers: [], mails: [] })(false, "runRestyleTosgonTable(ss)").value[0].startsWith("ЗОГСЛОО"), true, "no table: restyle stops");
  }
}

// ---------- тооллого хийх хуудас (count-sheets.gs) ----------
{
  const ss8 = buildBook(); const st8 = { triggers: [], mails: [] }; const run8 = makeRunner(ss8, st8);
  run8(false, "setupYear()");
  const today = run8(true, "chainSheetName(CHAINS[1], positionOfDate(new Date()))").value;
  const shop = ss8.getSheetByName(today), cs = ss8.getSheetByName("Соронзон тооллого");
  ok(cs !== null && ss8.getSheetByName("My home тооллого") !== null, "setupYear makes both count sheets");
  ok(cs.cell(1, 1).v.includes(`"${today}"`) && cs.cell(2, 1).v === "Барааны нэр" && cs.cell(2, 2).v === "Тооллого", "title names today's shop sheet; headers name and count");
  ok(cs.cell(3, 1).v === "Cat A" && cs.cell(3, 1).bold && cs.cell(4, 1).v === "A 1" && cs.cell(4, 2).v === "" && cs.cell(16, 1).v === "Cat B", "names in the shop's order, categories bold, counts empty, no expected numbers");
  const lock = cs.protections.filter(p => p.isSheet);
  ok(lock.length === 1 && lock[0].editors.length === 0 && lock[0].getUnprotectedRanges().length === 1 && lock[0].getUnprotectedRanges()[0].c === 2 && lock[0].getUnprotectedRanges()[0].r === 3, "names locked; only column B (from row 3) is open for the count");
  cs.cell(4, 2).v = 7; cs.cell(6, 2).v = 2;
  shop.insertRowBefore(5); shop.cell(5, 1).v = "Шинэ бараа"; shop.cell(5, 3).v = 100;
  const dry = run8(true, "refreshCountSheets(ss)").value;
  ok(dry.length === 1 && dry[0].includes("дахин бичнэ") && cs.cell(5, 1).v === "A 2", "DRY_RUN only reports the new list");
  run8(false, "refreshCountSheets(ss)");
  ok(cs.cell(4, 1).v === "Шинэ бараа" && cs.cell(4, 2).v === "" && cs.cell(5, 1).v === "A 1" && cs.cell(5, 2).v === 7 && cs.cell(6, 1).v === "A 2" && cs.cell(6, 2).v === "" && cs.cell(7, 1).v === "A 3" && cs.cell(7, 2).v === 2, "a product added on today's sheet appears; typed counts stay with their product");
  shop.deleteRow(6); // "A 1"
  run8(false, "refreshCountSheets(ss)");
  ok(cs.cell(4, 1).v === "Шинэ бараа" && cs.cell(5, 1).v === "A 2" && cs.cell(6, 1).v === "A 3" && cs.cell(6, 2).v === 2 && ![3, 4, 5, 6, 7].some(r => cs.cell(r, 2).v === 7), "a removed product takes its count with it, the others keep theirs");
  ok(run8(false, "refreshCountSheets(ss)").value.length === 0, "nothing changed: nothing is written");
  eq(run8(true, `[isCountSheetTouched(COUNT_SHEETS[0], 19, [{ chainKey: "my home", fromPosition: 0 }]), isCountSheetTouched(COUNT_SHEETS[0], 19, [{ chainKey: "соронзон", fromPosition: 21 }]), isCountSheetTouched(COUNT_SHEETS[0], 19, [{ chainKey: "соронзон", fromPosition: 19 }]), isCountSheetTouched(COUNT_SHEETS[0], 19, [undefined]), isCountSheetTouched(COUNT_SHEETS[0], 19, [])]`).value, [false, false, true, true, false], "after a sync only the changed shop's count sheet is read, and only when today's sheet or one before it changed");
  ok(run8(false, "reconcileSheet(ss.getSheetByName('Соронзон тооллого'))").value === null && run8(false, "trimSheet(ss.getSheetByName('My home тооллого'))").value.includes("мөр хасахгүй"), "colours and row trimming leave the count sheets alone");
}
// ---------- the OCR's yellow / green cells stay (dividers.gs) ----------
{
  const ssC = buildBook(); const runC = makeRunner(ssC, { triggers: [], mails: [] });
  runC(false, "setupYear()");
  const sh = ssC.getSheetByName("My/2026/10/31");
  const rowOf = (s, label) => { for (let r = 1; r <= s.rows.length; r++) if (s.cell(r, 1).v === label) return r; return null; };
  const aRows = Array.from({ length: 12 }, (_, i) => rowOf(sh, `A ${i + 1}`));
  aRows.forEach((r) => { sh.cell(r, 6).bg = "#ffe599"; }); // a day column where the OCR wrote many rows
  sh.cell(aRows[0], 7).bg = "#d9ead3"; sh.cell(aRows[1], 8).bg = "#ff0000";
  // T: a transfer column tinted #fff2cc; the OCR wrote one cell yellow and emptied another (no colour)
  const productRows = [...aRows, ...Array.from({ length: 12 }, (_, i) => rowOf(sh, `B ${i + 1}`))];
  productRows.forEach((r) => { sh.cell(r, 20).bg = "#fff2cc"; });
  sh.cell(aRows[2], 20).bg = "#ffe599"; sh.cell(aRows[3], 20).bg = WHITE;
  const b1 = rowOf(sh, "B 1");
  runC(false, "runBackgroundJobs()");
  ok(sh.cell(aRows[0], 6).bg === "#ffe599" && sh.cell(aRows[0], 7).bg === "#d9ead3" && sh.cell(aRows[2], 20).bg === "#ffe599", "the OCR's written (yellow) and approved (green) cells keep their colour");
  ok(sh.cell(b1, 6).bg === WHITE, "the other cells of that column keep the column's colour");
  ok(sh.cell(aRows[3], 20).bg === "#fff2cc", "a cell the OCR emptied gets its column's tint back (T, AJ)");
  ok(sh.cell(aRows[1], 8).bg === WHITE, "any other stray colour is still removed");
}

// ---------- the hourly sync inserts and deletes rows on a sheet that shows "Зөвхөн буруу" (a filter) ----------
{
  const t = ss.insertSheet("Filter test");
  t.getRange(4, 1).setValue("Хөргөгч"); for (let c = 1; c <= 36; c++) t.cell(4, c).bg = "#6aa84f";
  t.getRange(5, 1, 1, 4).setValues([["Old one", 1, 100, 120]]);
  t.getRange(6, 1, 1, 4).setValues([["Z last", 1, 100, 120]]);
  t.getRange(3, 1, 10, 40).createFilter();
  const result = run(false, `applyChange(ss.getSheetByName("Filter test"), [{ key: "new one", label: "New one", cost: 50, price: 60, shelfPrice: "", categoryKey: "хөргөгч" }], new Set(["old one"]))`).value;
  ok(result.added.join() === "new one" && result.removed.join() === "old one", "the product is added and the empty one removed although the sheet has a filter");
  ok(t.getFilter() !== null && [5, 6].map((r) => t.cell(r, 1).v).join() === "New one,Z last", "the filter is put back, the rows are in place");
  ss.deleteSheet(t);
}

console.log(`ALL ${n} SETUP CHECKS PASSED`);
