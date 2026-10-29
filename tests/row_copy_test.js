// Checks of row-copies.gs: the notes and script colours a new row copied from its template row.
// Usage: node tests/row_copy_test.js <project dir>
const fs = require("fs"), assert = require("assert");
const { Spreadsheet, makeRunnerFor } = require("./fake_book.js");

const dir = process.argv[2];
const FIRST_FILES = ["profile.gs", "config.gs"];
const gsFiles = fs.readdirSync(dir).filter((f) => f.endsWith(".gs"));
const loadOrder = [...FIRST_FILES.filter((f) => gsFiles.includes(f)), ...gsFiles.filter((f) => !FIRST_FILES.includes(f))];
const makeRunner = makeRunnerFor(loadOrder.map((f) => fs.readFileSync(`${dir}/${f}`, "utf8")).join("\n"));
let n = 0;
const eq = (a, b, m) => { assert.deepStrictEqual(a, b, m); n++; };

const BAND = "#f6b26b", WHITE = "#ffffff", TINT = "#fff2cc";
const BACKFILL = "Дэлгүүрийн орлогоос: Соронзон 2";
const OCR = "https://docs.google.com/spreadsheets/d/b/edit#gid=0&range=A2:M2&ocr=abc.1 = 1 (2026-06-02)";
// A Тосгон month as the owner saw it (2026-10-26): AP56 the backfill's 2, AP16..AP18 empty with its
// note (purple, copied with the row); G42 the OCR's 1, G30 and G31 empty, pink, with its note; I20
// emptied by hand, its note its own; column 36 tinted on every row.
function world() {
  const ss = new Spreadsheet();
  const s = ss.insertSheet("Tos/2026/10");
  s.getRange(4, 1).setValue("Хөргөгч"); for (let c = 1; c <= 60; c++) s.cell(4, c).bg = BAND;
  for (let r = 5; r <= 60; r++) { s.getRange(r, 1).setValue(`F${r}`); s.cell(r, 36).bg = TINT; }
  Object.assign(s.cell(56, 42), { v: 2, note: BACKFILL, color: "#8e24aa" });
  [16, 17, 18].forEach((r) => Object.assign(s.cell(r, 42), { note: BACKFILL, color: "#8e24aa" }));
  Object.assign(s.cell(42, 7), { v: 1, note: OCR, bg: "#ffe599", color: "#1155cc" });
  [30, 31].forEach((r) => Object.assign(s.cell(r, 7), { note: OCR, bg: "#ff99dd", color: "#1155cc" }));
  Object.assign(s.cell(20, 9), { note: "https://docs.google.com/spreadsheets/d/b/edit#gid=0&range=A3:M3&ocr=abc.2 = 3" });
  Object.assign(s.cell(30, 36), { note: "Шилжүүлэг (автомат): My home-ийн AK-аас хуулсан (1)", bg: "#f8bbd0" }); // a copy on the tinted column
  Object.assign(s.cell(44, 36), { v: 1, note: "Шилжүүлэг (автомат): My home-ийн AK-аас хуулсан (1)", bg: "#f8bbd0" });
  // the category's last row, the template of a new row: its own numbers, notes and colours
  Object.assign(s.cell(60, 42), { v: 1, note: "Дэлгүүрийн орлогоос: My home 1", color: "#8e24aa" });
  Object.assign(s.cell(60, 7), { v: 4, note: "https://docs.google.com/spreadsheets/d/b/edit#gid=0&range=A9:M9&ocr=abc.9 = 4", bg: "#ff99dd", color: "#1155cc" });
  const shop = ss.insertSheet("Бараа"); // not a month tab: never gone through
  Object.assign(shop.cell(5, 3), { note: BACKFILL });
  shop.cell(6, 3).v = 1; shop.cell(6, 3).note = BACKFILL;
  return { ss, s, shop, run: makeRunner(ss, { triggers: [], mails: [] }) };
}
const look = (s, r, c) => { const x = s.cell(r, c); return [x.v, x.note || "", x.bg, x.color || "#000000"]; };

// ---------- which cells are copies ----------
{
  const { run } = world();
  const found = run(true, `findCopiedNotes([["", 2], ["", ""], [1, ""]], [["a", "a"], ["b", ""], ["b", "c"]])`).value;
  eq(found, { copies: [{ row: 5, column: 1 }], own: [{ row: 4, column: 1, note: "a" }, { row: 6, column: 2, note: "c" }] },
     "an empty cell is a copy when a cell with a number in its own column has the very note; another column's does not count");
}

// ---------- DRY_RUN lists, changes nothing ----------
{
  const { s, run } = world();
  const result = run(true, "runClearCopiedNotes([ss])").value;
  eq(result[0], "Хуулагдсан тэмдэглэл нийт 6 нүдэнд", "six copies found");
  eq(look(s, 16, 42), ["", BACKFILL, WHITE, "#8e24aa"], "and nothing changed in DRY_RUN");
  eq(result.some((line) => line.includes("I20") && line.includes("хөндөөгүй")), true, "the hand-emptied cell with its own note is listed");
}

// ---------- the copies put right ----------
{
  const { s, shop, run } = world();
  const result = run(false, "runClearCopiedNotes([ss, ss])").value;
  eq(result[0], "Хуулагдсан тэмдэглэл нийт 6 нүдэнд", "a workbook given twice is gone through once");
  eq([16, 17, 18].map((r) => look(s, r, 42)), Array(3).fill(["", "", WHITE, "#000000"]), "the backfill's copied notes go, its purple with them");
  eq([30, 31].map((r) => look(s, r, 7)), Array(2).fill(["", "", WHITE, "#000000"]), "the OCR's copied notes go and the pink back to the column's white");
  eq(look(s, 30, 36), ["", "", TINT, "#000000"], "a copied transfer pink goes back to the column's tint");
  eq([look(s, 56, 42), look(s, 42, 7), look(s, 44, 36)], [[2, BACKFILL, WHITE, "#8e24aa"], [1, OCR, "#ffe599", "#1155cc"], [1, "Шилжүүлэг (автомат): My home-ийн AK-аас хуулсан (1)", "#f8bbd0", "#000000"]],
     "the cells with the numbers keep their notes and colours");
  eq(look(s, 20, 9)[1].startsWith("https://"), true, "a hand-emptied cell keeps its own note");
  eq(shop.cell(5, 3).note, BACKFILL, "a tab that is not a month is not gone through");
  eq(run(false, "runClearCopiedNotes([ss])").value[0], "Хуулагдсан тэмдэглэл нийт 0 нүдэнд", "a second run finds nothing");
}

// ---------- a new row from a template row (year-sync.gs insertProduct) ----------
{
  const { ss, s, run } = world();
  const result = run(false, `(() => { const sheet = ss.getSheetByName("Tos/2026/10"); insertProduct(sheet, readModel(sheet), { key: "f56a", label: "F60a", categoryKey: "хөргөгч", cost: 1, price: 2 }); return readModel(sheet).products.find((p) => p.label === "F60a").row; })()`).value;
  eq(result, 61, "the new row goes in under the category's last row, its template");
  eq([look(s, result, 42), look(s, result, 7)], [["", "", WHITE, "#000000"], ["", "", WHITE, "#000000"]], "the new row has no note and no script colour of its template");
  eq(look(s, result, 36)[2], TINT, "a column's own tint stays on the new row");
  eq([look(s, 60, 42), look(s, 60, 7)[2]], [[1, "Дэлгүүрийн орлогоос: My home 1", WHITE, "#8e24aa"], "#ff99dd"], "and the template keeps its own");
}

console.log(`ALL ${n} ROW COPY CHECKS PASSED`);
