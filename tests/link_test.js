// Checks of price-link.gs: one product, one price in Соронзон, My home and Тосгон, and Соронзон's
// new products in Тосгон. Usage: node tests/link_test.js <project dir>
const fs = require("fs"), assert = require("assert");
const { Spreadsheet, makeRunnerFor } = require("./fake_book.js");

const dir = process.argv[2];
const FIRST_FILES = ["profile.gs", "config.gs"];
const gsFiles = fs.readdirSync(dir).filter((f) => f.endsWith(".gs"));
const loadOrder = [...FIRST_FILES.filter((f) => gsFiles.includes(f)), ...gsFiles.filter((f) => !FIRST_FILES.includes(f))];
const shipped = [...loadOrder, "profiles/sheet-rows.gs"].map((f) => fs.readFileSync(`${dir}/${f}`, "utf8")).join("\n");
// The price link does not touch the gift column; these sheets are built without it (paired layout).
const code = shipped.replace("const GIFT_COLUMN = true;", "const GIFT_COLUMN = false;");
const makeRunner = makeRunnerFor(code);
let n = 0;
const ok = (c, m) => { assert(c, m); n++; };
const eq = (a, b, m) => { assert.deepStrictEqual(a, b, m); n++; };
const withTosgon = (c) => `${c}\nSpreadsheetApp.openById = () => extra;`; // TOSGON_BOOK_ID opens the fake Тосгон book

const band = (s, row) => { for (let c = 1; c <= 36; c++) s.cell(row, c).bg = "#f6b26b"; };
// A shop sheet in the paired layout: A name, B opening, C cost, D price, E shelf price.
function shopSheet(ss, name, price = 150) {
  const s = ss.insertSheet(name);
  s.getRange(2, 1).setValue(name);
  const put = (row, label, d) => { s.getRange(row, 1).setValue(label); s.getRange(row, 2, 1, 4).setValues([[2, 100, d, d + 50]]); };
  s.getRange(4, 1).setValue("Хөргөгч"); band(s, 4);
  for (let i = 1; i <= 12; i++) put(4 + i, `F${i}`, price);
  s.getRange(17, 1).setValue("Хятад жижиг бараа"); band(s, 17);
  for (let i = 1; i <= 10; i++) put(17 + i, `K${i}`, 30);
  return s;
}
// A Тосгон month sheet: A name, B opening, C cost, D price, E the first income day; once moved
// (shelf, layout-migration.gs, 2026-10-22) E is the shelf price "яг харагдах үнэ" (E3) and day 1 is F.
function tosgonSheet(book, name, price = 150, count = 10, shelf = false) {
  const s = book.insertSheet(name);
  if (shelf) s.getRange(3, 5).setValue("яг харагдах үнэ");
  s.getRange(4, 1).setValue("Хөргөгч"); band(s, 4);
  for (let i = 1; i <= count; i++) { s.getRange(4 + i, 1).setValue(`F${i}`); s.getRange(4 + i, 2, 1, 4).setValues([[5, 100, price, shelf ? price + 50 : ""]]); }
  return s;
}
const rowOf = (s, label) => s.rows.findIndex((r) => r[1] && r[1].v === label) + 1;
const valueOf = (s, label, column) => s.cell(rowOf(s, label), column).v;

function world() {
  const ss = new Spreadsheet();
  ["Sor/2026/9/30", "Sor/2026/10/15", "Sor/2026/10/31", "Sor/2026/11/15", "My/2026/10/15", "My/2026/10/31"].forEach((name) => shopSheet(ss, name));
  const tos = new Spreadsheet();
  ["Tos/2026/9", "Tos/2026/10", "Tos/2026/11"].forEach((name) => tosgonSheet(tos, name));
  const state = { triggers: [], mails: [], extra: tos, docProps: { pairedLayoutReady: "1", priceLinkReady: "1" } };
  const run = makeRunner(ss, state);
  const edit = (sheet, row, column, oldValue) => { state.editEvent = { range: sheet.getRange(row, column), oldValue }; const r = run(false, "onSheetEdit(editEvent)", withTosgon); state.editEvent = undefined; return r; };
  return { ss, tos, state, run, edit };
}

// ---------- pure rules ----------
{
  const { run } = world();
  eq(run(true, `["хятад жижиг бараа", "хятад жижиг бараа будаа агшаагч", "солонгос жижиг бараа", "хөргөгч", null].map(isShopOnlyCategory)`).value,
     [true, true, true, false, false], "the small-goods categories (and every one that starts so) stay in the shops");
  eq(run(true, `[tosgonTabOf(periodPosition(2026, 10, 0)), tosgonTabOf(periodPosition(2026, 10, 1)), otherChainOf("соронзон").label, otherChainOf("my home").label]`).value,
     ["Tos/2026/10", "Tos/2026/10", "My home", "Соронзон"], "both halves of a month go to Тосгон's sheet of that month; the other shop");
}

// ---------- a price typed in one shop ----------
{
  const { ss, tos, edit } = world();
  const sor = ss.getSheetByName("Sor/2026/10/15"), my15 = ss.getSheetByName("My/2026/10/15"), my31 = ss.getSheetByName("My/2026/10/31");
  my31.cell(rowOf(my31, "F3"), 4).v = 999; // My home's own price on its later sheet: kept
  sor.cell(rowOf(sor, "F1"), 4).v = 170; edit(sor, rowOf(sor, "F1"), 4, 150);
  eq([valueOf(my15, "F1", 4), valueOf(my31, "F1", 4)], [170, 170], "the price reaches My home's sheet of the same half-month and its later ones");
  eq(["Tos/2026/9", "Tos/2026/10", "Tos/2026/11"].map((t) => valueOf(tos.getSheetByName(t), "F1", 4)), [150, 170, 170], "and Тосгон's sheet of that month and later, never the months before");
  sor.cell(rowOf(sor, "F3"), 4).v = 160; edit(sor, rowOf(sor, "F3"), 4, 150);
  eq([valueOf(my15, "F3", 4), valueOf(my31, "F3", 4)], [160, 999], "a later sheet with its own price keeps it, as in price-sync.gs");
  sor.cell(rowOf(sor, "F2"), 3).v = 120; edit(sor, rowOf(sor, "F2"), 3, 100);
  eq([valueOf(my15, "F2", 3), valueOf(tos.getSheetByName("Tos/2026/10"), "F2", 3)], [120, 120], "the cost the same way");
  sor.cell(rowOf(sor, "F2"), 5).v = 222; edit(sor, rowOf(sor, "F2"), 5, 200);
  eq([valueOf(my15, "F2", 5), valueOf(tos.getSheetByName("Tos/2026/10"), "F2", 5)], [222, ""], "the shelf price (E) only in the other shop: Тосгон's E is a day column");
  my15.cell(rowOf(my15, "F4"), 4).v = 175; edit(my15, rowOf(my15, "F4"), 4, 150);
  eq([valueOf(sor, "F4", 4), valueOf(ss.getSheetByName("Sor/2026/11/15"), "F4", 4), valueOf(tos.getSheetByName("Tos/2026/10"), "F4", 4)], [175, 175, 175], "the other way round: My home -> Соронзон and Тосгон");
}

// ---------- not before setupPriceLink ----------
{
  const { ss, tos, state, edit } = world();
  delete state.docProps.priceLinkReady;
  const sor = ss.getSheetByName("Sor/2026/10/15");
  sor.cell(rowOf(sor, "F1"), 4).v = 170; edit(sor, rowOf(sor, "F1"), 4, 150);
  eq([valueOf(ss.getSheetByName("My/2026/10/15"), "F1", 4), valueOf(tos.getSheetByName("Tos/2026/10"), "F1", 4)], [150, 150], "until it is turned on, a price stays in its own shop");
}

// ---------- a product added in Соронзон ----------
{
  const { ss, tos, edit } = world();
  const sor = ss.getSheetByName("Sor/2026/10/15");
  sor.insertRowAfter(16); sor.getRange(17, 1).setValue("F new"); sor.getRange(17, 2, 1, 4).setValues([[1, 300, 400, 450]]);
  edit(sor, 17, 1);
  const t10 = tos.getSheetByName("Tos/2026/10");
  ok(rowOf(t10, "F new") > rowOf(t10, "F10") && rowOf(t10, "F10") > 0, "a new product of Соронзон is added to Тосгон's sheet of that month, under its category");
  eq([valueOf(t10, "F new", 3), valueOf(t10, "F new", 4), valueOf(t10, "F new", 5)], [300, 400, ""], "with its cost and price; Тосгон's E (a day) stays empty");
  ok(rowOf(t10, "K1") === 0 && rowOf(t10, "F11") > 0, "the small goods never go to Тосгон; the other products Тосгон lacked came along");
  const my = ss.getSheetByName("My/2026/10/15");
  my.insertRowAfter(16); my.getRange(17, 1).setValue("M new"); my.getRange(17, 2, 1, 4).setValues([[1, 300, 400, 450]]);
  edit(my, 17, 1);
  ok(rowOf(t10, "M new") === 0, "a product added in My home does not go to Тосгон");
}

// ---------- the one-time alignment ----------
{
  const { ss, tos, state, run } = world();
  delete state.docProps.priceLinkReady;
  ["Sor/2026/9/30", "Sor/2026/10/31", "Sor/2026/11/15"].forEach((name) => { const s = ss.getSheetByName(name); s.cell(rowOf(s, "F1"), 4).v = 180; });
  const preview = run(false, `runPriceAlignment(SpreadsheetApp.getActiveSpreadsheet(), extra, new Date(2026, 9, 20), false)`, withTosgon).value;
  eq([preview.cells, preview.added, preview.tab, valueOf(tos.getSheetByName("Tos/2026/10"), "F1", 4)], [1, ["F11", "F12"], "Tos/2026/10", 150],
     "the preview lists October's price (from Соронзон's second half) and the products to add, and writes nothing");
  const done = run(false, `runPriceAlignment(SpreadsheetApp.getActiveSpreadsheet(), extra, new Date(2026, 9, 20), true)`, withTosgon).value;
  eq(["Tos/2026/9", "Tos/2026/10"].map((t) => valueOf(tos.getSheetByName(t), "F1", 4)), [150, 180], "October takes Соронзон's October price; September is left as it was");
  ok(done.cells === 1 && rowOf(tos.getSheetByName("Tos/2026/10"), "F12") > 0, "and the missing products are added to the current month");
  run(false, "setupPriceLink()", withTosgon);
  eq(state.docProps.priceLinkReady, "1", "setupPriceLink turns the linking on");
}

// ---------- only the list's names go to Тосгон (2026-10-22) ----------
{
  const { ss, tos, state, run } = world();
  delete state.docProps.priceLinkReady;
  const list = ss.insertSheet("Бараа");
  list.getRange(4, 1, 13, 2).setValues([["Хөргөгч", ""], ...Array.from({ length: 11 }, (_, i) => [`F${i + 1}`, "Хоёулаа"]), ["Хятад жижиг бараа", ""]]);
  const result = run(false, `runPriceAlignment(SpreadsheetApp.getActiveSpreadsheet(), extra, new Date(2026, 9, 20), true)`, withTosgon).value;
  const t10 = tos.getSheetByName("Tos/2026/10");
  ok(rowOf(t10, "F11") > 0 && rowOf(t10, "F12") === 0, "a name on the list is added to Тосгон, one off the list is not");
  ok(result.notes.some((note) => note.includes('"F12"') && note.includes("жагсаалтад алга")), "and the one left out is reported");
  const again = run(false, `runPriceAlignment(SpreadsheetApp.getActiveSpreadsheet(), extra, new Date(2026, 9, 20), true)`, withTosgon).value;
  eq(again.added, [], "a rerun adds nothing twice");
  run(false, "setupPriceLink()", withTosgon);
  eq(state.docProps.priceLinkReady, "1", "and setupPriceLink turns the link on");
}

// ---------- only the missing products, in the list's spelling (2026-10-22) ----------
{
  const { ss, tos, state, run } = world();
  delete state.docProps.priceLinkReady;
  const list = ss.insertSheet("Бараа");
  list.getRange(4, 1, 13, 2).setValues([["Хөргөгч", ""], ...Array.from({ length: 10 }, (_, i) => [`F${i + 1}`, "Хоёулаа"]), ["F 12", "Соронзон"], ["Хятад жижиг бараа", ""]]);
  const sor = ss.getSheetByName("Sor/2026/10/31");
  sor.cell(rowOf(sor, "F12"), 1).v = "F  12"; // typed with two spaces
  sor.cell(rowOf(sor, "F1"), 4).v = 180;
  const result = run(false, `runMissingToTosgon(SpreadsheetApp.getActiveSpreadsheet(), extra, new Date(2026, 9, 20), true)`, withTosgon).value;
  const t10 = tos.getSheetByName("Tos/2026/10");
  ok(rowOf(t10, "F 12") > 0 && rowOf(t10, "F  12") === 0, "a name typed with extra spaces goes to Тосгон in the list's spelling");
  eq(valueOf(t10, "F1", 4), 150, "only the missing products: no price is changed");
  eq(result.added, ["F 12"], "the report names what was added");
  ok(state.docProps.priceLinkReady === undefined, "and the link is not turned on by it");
}

// ---------- Тосгон with the shelf price in E (2026-10-22) ----------
{
  const { ss, tos, run, edit } = world();
  tos.deleteSheet(tos.getSheetByName("Tos/2026/10"));
  const moved = tosgonSheet(tos, "Tos/2026/10", 150, 10, true); // the moved sheet; Tos/2026/11 is not moved yet
  const notMoved = tos.getSheetByName("Tos/2026/11");
  const sor = ss.getSheetByName("Sor/2026/10/15");
  sor.cell(rowOf(sor, "F2"), 5).v = 230; edit(sor, rowOf(sor, "F2"), 5, 200);
  eq([valueOf(moved, "F2", 5), valueOf(notMoved, "F2", 5)], [230, ""], "the shelf price reaches a moved Тосгон sheet, never the day 1 column of one not moved");
  const sor31 = ss.getSheetByName("Sor/2026/10/31");
  sor31.cell(rowOf(sor31, "F3"), 5).v = 260;
  run(false, `runPriceAlignment(SpreadsheetApp.getActiveSpreadsheet(), extra, new Date(2026, 9, 20), true)`, withTosgon);
  eq([valueOf(moved, "F3", 5), valueOf(notMoved, "F3", 5)], [260, ""], "the alignment brings Соронзон's shelf price to Тосгон's moved months");
}
{
  const { ss, tos, run } = world();
  tos.deleteSheet(tos.getSheetByName("Tos/2026/10"));
  const moved = tosgonSheet(tos, "Tos/2026/10", 150, 10, true);
  run(false, `runMissingToTosgon(SpreadsheetApp.getActiveSpreadsheet(), extra, new Date(2026, 9, 20), true)`, withTosgon);
  eq([valueOf(moved, "F11", 3), valueOf(moved, "F11", 4), valueOf(moved, "F11", 5), valueOf(moved, "F11", 6)], [100, 150, 200, ""], "a product added to a moved Тосгон sheet gets its shelf price in E; day 1 (F) stays empty");
}

// ---------- a product added to Тосгон takes no other product's note or the script's colours (2026-10-26) ----------
{
  const { tos, run } = world();
  const t10 = tos.getSheetByName("Tos/2026/10");
  const last = rowOf(t10, "F10"); // the category's last product row: the template of a new row
  Object.assign(t10.cell(last, 10), { v: 2, note: "Дэлгүүрийн орлогоос: Соронзон 2", color: "#8e24aa" });
  Object.assign(t10.cell(last, 12), { v: 3, note: "https://docs.google.com/spreadsheets/d/b/edit#gid=0&range=A2:M2&ocr=abc.1 = 3", bg: "#ff99dd", color: "#1155cc" });
  run(false, `runMissingToTosgon(SpreadsheetApp.getActiveSpreadsheet(), extra, new Date(2026, 9, 20), true)`, withTosgon);
  const added = rowOf(t10, "F11");
  const marks = (row, column) => { const x = t10.cell(row, column); return [x.v, x.note || "", x.bg, x.color || "#000000"]; };
  eq([marks(added, 10), marks(added, 12)], [["", "", "#ffffff", "#000000"], ["", "", "#ffffff", "#000000"]],
     "the new row has none of the template's notes, OCR colours or script fonts (Tos/2026/10!AP, 2026-10-26)");
  eq([marks(last, 10), marks(last, 12)], [[2, "Дэлгүүрийн орлогоос: Соронзон 2", "#ffffff", "#8e24aa"], [3, "https://docs.google.com/spreadsheets/d/b/edit#gid=0&range=A2:M2&ocr=abc.1 = 3", "#ff99dd", "#1155cc"]],
     "and the template row keeps its own");
}

console.log(`ALL ${n} LINK CHECKS PASSED`);
