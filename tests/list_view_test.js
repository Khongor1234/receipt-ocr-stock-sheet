// Checks of the price view in the product list "Бараа" (product-list.gs clearListViewEdit): C..E
// below row 4 are filled by ONE formula in C4, so anything typed or pasted there is cleared at once.
// Usage: node tests/list_view_test.js <project dir>
const fs = require("fs"), assert = require("assert");
const { Spreadsheet, makeRunnerFor } = require("./fake_book.js");

const dir = process.argv[2];
const FIRST_FILES = ["profile.gs", "config.gs"];
const gsFiles = fs.readdirSync(dir).filter((f) => f.endsWith(".gs"));
const loadOrder = [...FIRST_FILES.filter((f) => gsFiles.includes(f)), ...gsFiles.filter((f) => !FIRST_FILES.includes(f))];
const makeRunner = makeRunnerFor(loadOrder.map((f) => fs.readFileSync(`${dir}/${f}`, "utf8")).join("\n"));
let n = 0;
const eq = (a, b, m) => { assert.deepStrictEqual(a, b, m); n++; };

const VIEW = "=ARRAYFORMULA(VLOOKUP(...))";
const SPILLED = 110000; // a number the C4 formula shows (the fake keeps it as a plain value)
// The list as it stands live (2026-10-27): A name, B shop, C..E the price view of C4.
function world({ copy = false } = {}) {
  const ss = new Spreadsheet();
  const list = ss.insertSheet("Бараа");
  list.getRange(1, 1).setValue("Соронзон орон");
  list.getRange(3, 1, 1, 5).setValues([["Барааны нэр", "Аль дэлгүүрт", "өртөг үнэ", "зарах үнэ", "яг харагдах үнэ"]]);
  if (copy) list.getRange(4, 1).setFormula('=IMPORTRANGE("x","Бараа!A4:B")');
  else list.getRange(4, 1).setValue("Хятад жижиг бараа");
  list.getRange(4, 3).setFormula(VIEW);
  ["304 7l", "304 8l", "304 9l", "304 10l"].forEach((name, at) => list.getRange(5 + at, 1, 1, 2).setValues([[name, "Хоёулаа"]]));
  list.cell(6, 3).v = SPILLED;
  return { ss, list, run: makeRunner(ss, { triggers: [], mails: [] }) };
}
// runProductListEdit with an edit event of the range (row, column, rows, columns) and the typed value.
const edit = (run, [r, c, h = 1, w = 1], value) =>
  run(false, `runProductListEdit({ range: ss.getSheetByName("Бараа").getRange(${r}, ${c}, ${h}, ${w}), value: ${JSON.stringify(value)} })`);
const said = (logs) => logs.includes("Үнийг энд бичихгүй");

// ---------- a price typed in the view is cleared, with a message ----------
{
  const { list, run } = world();
  list.cell(5, 4).v = 150000;
  const { logs } = edit(run, [5, 4], "150000");
  eq(list.cell(5, 4).v, "", "the typed price is cleared, so the C4 formula fills the column again");
  eq(list.cell(4, 3).f, VIEW, "the view formula itself is untouched");
  eq(said(logs), true, "the message says prices go on the shop sheet");
}

// ---------- a pasted row: only its C..E part is cleared ----------
{
  const { list, run } = world();
  list.getRange(7, 1, 1, 6).setValues([["Шинэ бараа", "Соронзон", 1, 2, 3, "тэмдэглэл"]]);
  const { logs } = edit(run, [7, 1, 1, 6], undefined);
  eq(list.getRange(7, 1, 1, 6).getValues()[0], ["Шинэ бараа", "Соронзон", "", "", "", "тэмдэглэл"], "name, shop and F stay; C..E cleared");
  eq(said(logs), true, "and the message is shown");
}

// ---------- nothing happens outside the view or when a cell was only emptied ----------
{
  const { list, run } = world();
  const before = list.logical();
  eq(said(edit(run, [6, 3], undefined).logs), false, "Delete on a cell the formula fills: no message");
  eq(list.cell(6, 3).v, SPILLED, "and nothing cleared");
  list.cell(5, 2).v = "Соронзон";
  eq(said(edit(run, [5, 2], "Соронзон").logs), false, "a shop chosen in B is not the view");
  list.cell(5, 6).v = 7;
  eq(said(edit(run, [5, 6], "7").logs), false, "column F is not the view");
  eq(list.cell(5, 6).v, 7, "and keeps its value");
  list.cell(3, 4).v = "үнэ";
  eq(said(edit(run, [3, 4], "үнэ").logs), false, "the header row is not the view");
  eq(list.cell(3, 4).v, "үнэ", "and keeps its value");
  list.cell(4, 4).v = 9;
  eq(said(edit(run, [4, 4], "9").logs), false, "row 4 (the formula's own row, owner only) is not the view");
  eq(list.cell(4, 4).v, 9, "and keeps its value");
  list.cell(3, 4).v = "зарах үнэ"; list.cell(5, 2).v = "Хоёулаа"; list.cell(5, 6).v = ""; list.cell(4, 4).v = "";
  eq(list.logical(), before, "nothing else changed");
}

// ---------- a copy of the list (Тосгон) is left alone ----------
{
  const { list, run } = world({ copy: true });
  list.cell(5, 3).v = 5;
  eq(said(edit(run, [5, 3], "5").logs), false, "no message in a copy");
  eq(list.cell(5, 3).v, 5, "and nothing cleared");
}

console.log(`list_view_test: ${n} checks passed`);
