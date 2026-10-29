// Checks of official-names.gs: the factory's brand spelling and the name order of the rows.
// Usage: node tests/names_order_test.js <project dir>
const fs = require("fs"), assert = require("assert");
const { Spreadsheet, makeRunnerFor } = require("./fake_book.js");

const dir = process.argv[2];
const FIRST_FILES = ["profile.gs", "config.gs"];
const gsFiles = fs.readdirSync(dir).filter((f) => f.endsWith(".gs"));
const loadOrder = [...FIRST_FILES.filter((f) => gsFiles.includes(f)), ...gsFiles.filter((f) => !FIRST_FILES.includes(f))];
const code = loadOrder.map((f) => fs.readFileSync(`${dir}/${f}`, "utf8")).join("\n");
const makeRunner = makeRunnerFor(code);
let n = 0;
const eq = (a, b, m) => { assert.deepStrictEqual(a, b, m); n++; };
const ok = (c, m) => { assert(c, m); n++; };

const BAND = "#f6b26b";
const band = (s, row) => { for (let c = 1; c <= 36; c++) s.cell(row, c).bg = BAND; };
// A shop sheet (paired layout header in E3): a category, its products in the given order, then fillers
// so the sheet counts as an entry sheet (20 priced rows).
function shopSheet(ss, name, products) {
  const s = ss.insertSheet(name);
  s.getRange(2, 1).setValue(name);
  s.getRange(3, 5).setValue("яг харагдах үнэ");
  s.getRange(4, 1).setValue("Хөргөгч"); band(s, 4);
  products.forEach((label, i) => { s.getRange(5 + i, 1).setValue(label); s.getRange(5 + i, 2, 1, 4).setValues([[i + 1, 100, 200 + i, ""]]); });
  const next = 5 + products.length;
  s.getRange(next, 1).setValue("Жижиг бараа"); band(s, next);
  for (let i = 1; i <= 20; i++) { s.getRange(next + i, 1).setValue(`Z ${i}`); s.getRange(next + i, 2, 1, 4).setValues([[1, 10, 20, ""]]); }
  return s;
}
const labels = (s, from, count) => Array.from({ length: count }, (_, i) => s.cell(from + i, 1).v);

// ---------- pure rules ----------
{
  const run = makeRunner(new Spreadsheet(), { triggers: [], mails: [] });
  eq(run(true, `["Cmidea 723l", "Panasonica 55", "TGL 201l", "Cmidea16", "Индүү аkira 888", "TGL -12OL", "xinix 240l", "Ус буцалгагч 8037", "Haier  kond", "Pa nasonica smart", "Skyworch 70"].map(officialName)`).value,
     ["Midea 723l", "Panasonic 55", "TCL 201l", "Midea 16", "Индүү Akira 888", "TCL 120L", "XINGX 240l", "Ус буцалгагч 8037", "Haier kond", "Panasonic smart", "Skyworth 70"],
     "brands as their factory writes them; a glued number gets its space; Mongolian words stay");
  eq(run(true, `["тогтоогч дунд", "солонгос плетик", "tb том", "0,8хоол", "xinix 50L"].map(officialName)`).value,
     ["Тогтоогч дунд", "Солонгос плетик", "Tb том", "0,8хоол", "XINGX 50L"], "every name starts with a capital letter");
  const plan = run(true, `planOfficialRenames(["Panasonica 50l", "Panasonic 50l", "Cmidea 195l", "Haier 300l"])`).value;
  eq(plan.renames, [{ from: "Cmidea 195l", to: "Midea 195l" }], "a name already right is not renamed");
  ok(plan.skipped.length === 1 && plan.skipped[0].includes("Panasonica 50l"), "a new name another product has is left and reported");
  const sorted = run(true, `["Meiling 112l", "Ус буцалгагч 8037", "Meiling 86l", "Tоль 250", "Acewise 117l (2)", "Haier 300l", "Толь 200", "Acewise 117l", "8811 ус", "304 7l", "аkira 888"].sort(compareProductNames)`).value;
  eq(sorted, ["304 7l", "Acewise 117l", "Acewise 117l (2)", "аkira 888", "Haier 300l", "Meiling 86l", "Meiling 112l", "8811 ус", "Толь 200", "Tоль 250", "Ус буцалгагч 8037"],
     "English names first, then Mongolian; letters by alphabet, numbers by value");
  eq(run(true, `planRowOrder([{ label: "B", movable: true }, { label: "A", movable: true }, { label: "Cat", movable: false }, { label: "D 10", movable: true }, { label: "D 9", movable: true }])`).value,
     [1, 0, 2, 4, 3], "each run between headings is ordered on its own; a heading stays");
  eq(run(true, `[22, 23, 3, 6, 7, 12, 21].map(isOrderNight)`).value, [true, true, true, true, false, false, false], "rows move only at night");
  eq(run(true, `findInsertAfterRow({ categories: [{ key: "хөргөгч", row: 4 }], products: [{ label: "Haier 100l", row: 5, categoryKey: "хөргөгч" }, { label: "Haier 300l", row: 6, categoryKey: "хөргөгч" }] }, "хөргөгч", "Haier 200l")`).value,
     5, "a product the sync copies goes in at its place in name order");
  eq(run(true, `findInsertAfterRow({ categories: [{ key: "хөргөгч", row: 4 }], products: [{ label: "Haier 100l", row: 5, categoryKey: "хөргөгч" }] }, "хөргөгч", "Aucma 50l")`).value,
     4, "a product that comes first goes right under its heading");
}

// ---------- tidyProductNames on a workbook ----------
function world() {
  const ss = new Spreadsheet();
  const list = ss.insertSheet("Бараа");
  list.getRange(4, 1, 8, 2).setValues([["Хөргөгч", ""], ["Meiling 112l", "Хоёулаа"], ["Cmidea 195l", "Соронзон"], ["Meiling 86l", "Хоёулаа"], ["Panasonica 50l", "My home"], ["хөргөгч жижиг", "Соронзон"], ["Жижиг бараа", ""], ["Z 1", "Соронзон"]]);
  const sor = shopSheet(ss, "Sor/2026/10/15", ["Meiling 112l", "Cmidea 195l", "Meiling 86l", "хөргөгч жижиг"]);
  const my = shopSheet(ss, "My/2026/10/15", ["Panasonica 50l", "Meiling 112l", "Meiling 86l"]);
  sor.cell(5, 7).v = 4; sor.cell(5, 7).bg = "#ffff00"; sor.cell(5, 7).note = "OCR"; // Meiling 112l's income cell
  const table = ss.insertSheet("_Тосгон нэр тулгалт");
  table.getRange(3, 5).setValue("Cmidea 195l"); table.getRange(3, 8).setValue("Cmidea 195l"); table.getRange(3, 10).setValue("Panasonica 50l");
  const state = { triggers: [], mails: [], docProps: { pairedLayoutReady: "1" } };
  return { ss, list, sor, my, table, state, run: makeRunner(ss, state) };
}
{
  const { list, sor, run } = world();
  const before = sor.logical();
  const dry = run(true, "runTidyProductNames(SpreadsheetApp.getActiveSpreadsheet())").value.lines;
  ok(dry.some((line) => line.includes("Cmidea 195l -> Midea 195l")) && dry.some((line) => line.includes("Sor/2026/10/15: ")), "the trial run lists the renames and the sheets to reorder");
  ok(sor.logical() === before && list.cell(6, 1).v === "Cmidea 195l", "and changes nothing");
}
{
  const { list, sor, my, table, ss, run } = world();
  run(false, "runTidyProductNames(SpreadsheetApp.getActiveSpreadsheet())");
  eq(labels(list, 5, 5), ["Meiling 86l", "Meiling 112l", "Midea 195l", "Panasonic 50l", "Хөргөгч жижиг"], "the list: factory names with a capital, English then Mongolian, in name order");
  eq([list.cell(4, 1).v, list.cell(10, 1).v], ["Хөргөгч", "Жижиг бараа"], "its headings stay");
  eq(labels(sor, 5, 4), ["Meiling 86l", "Meiling 112l", "Midea 195l", "Хөргөгч жижиг"], "Соронзон's sheet: renamed (a capital too) and in order");
  eq(labels(my, 5, 3), ["Meiling 86l", "Meiling 112l", "Panasonic 50l"], "My home's sheet too");
  const row112 = 6;
  eq([sor.cell(row112, 2).v, sor.cell(row112, 4).v, sor.cell(row112, 7).v, sor.cell(row112, 7).bg, sor.cell(row112, 7).note], [1, 200, 4, "#ffff00", "OCR"],
     "a row moves with its numbers, colour and note");
  eq([sor.cell(4, 1).v, sor.cell(9, 1).v, sor.cell(4, 1).bg], ["Хөргөгч", "Жижиг бараа", BAND], "the category rows stay in place");
  eq([table.cell(3, 5).v, table.cell(3, 8).v, table.cell(3, 10).v], ["Midea 195l", "Midea 195l", "Panasonic 50l"], "the Тосгон name table follows the renames");
  const log = ss.getSheetByName("_Нэр солилт");
  ok(log !== null && log.cell(2, 2).v !== "" && log.cell(2, 3).v !== "", "the renames are logged for Тосгон's own sheets");
  const again = run(false, "runTidyProductNames(SpreadsheetApp.getActiveSpreadsheet())").value.lines;
  ok(again[0].includes("албан ёсны нэр: 0 бараа") && again.some((line) => line.startsWith("нэрийн дараалал: 0 хуудас")), "a second run finds nothing to do");
}
{
  const { sor, run } = world();
  sor.cell(6, 9).f = "=G6*2";
  const lines = run(false, "runNameOrder(SpreadsheetApp.getActiveSpreadsheet(), { list: null, sheets: [SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sor/2026/10/15')] })").value;
  ok(lines.some((line) => line.includes("томьёотой")) && sor.cell(5, 1).v === "Meiling 112l", "a sheet with a formula pointing at a cell is not reordered");
}

{ // a count block after the totals (Тооллого 10/23, Тосгон's BY:CA) moves with its row (2026-10-26)
  const { sor, run } = world();
  const F = run(true, "spotFirstColumn(LAYOUT)").value;
  sor.cell(3, F).v = "Тооллого 10/23"; sor.cell(3, F + 1).v = "Үлдэгдэл 10/23"; sor.cell(3, F + 2).f = '=ARRAYFORMULA({"Зөрүү 10/23";G4:G})';
  sor.cell(5, F).v = 2; sor.cell(5, F).note = "Meiling 112l count"; sor.cell(5, F).bg = "#ffe599"; sor.cell(5, F + 1).v = 1; // Meiling 112l
  sor.cell(7, F).v = 7; sor.cell(7, F).note = "Meiling 86l count"; // Meiling 86l
  run(false, "runNameOrder(SpreadsheetApp.getActiveSpreadsheet(), { list: null, sheets: [SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sor/2026/10/15')] })");
  const row = (name) => labels(sor, 1, 12).indexOf(name) + 1;
  eq([row("Cmidea 195l"), row("Meiling 86l"), row("Meiling 112l")], [5, 6, 7], "the rows are put in name order");
  eq([sor.cell(7, F).v, sor.cell(7, F).note, sor.cell(7, F).bg, sor.cell(7, F + 1).v], [2, "Meiling 112l count", "#ffe599", 1], "the block's count, its note and colour and the balance go with their product");
  eq([sor.cell(6, F).v, sor.cell(6, F).note, sor.cell(6, F + 1).v || ""], [7, "Meiling 86l count", ""], "the other product's count goes with its own row");
  eq([sor.cell(5, F).v || "", sor.cell(5, F).note || ""], ["", ""], "and a product without a count has none");
  ok(String(sor.cell(3, F + 2).f).startsWith("=ARRAYFORMULA"), "the difference column, one formula of the header row, is left alone");
}

{ // cut short after every piece of work (Google's 6 minutes): the next runs go on from there
  const { list, sor, my, state, run } = world();
  const short = (c) => c.replace("const TIDY_TIME_BUDGET_MS = 4 * 60 * 1000;", "const TIDY_TIME_BUDGET_MS = -1;");
  const first = run(false, "runTidyProductNames(SpreadsheetApp.getActiveSpreadsheet(), Date.now() - 1)", short).value;
  ok(!first.done && list.cell(6, 1).v === "Midea 195l" && labels(sor, 5, 4).includes("Cmidea 195l"), "a part stops after the list: the shops still have the old name");
  ok(JSON.parse(state.props.tidyNamesState).renames.length === 3, "and the planned renames are kept");
  let runs = 1;
  while (state.props.tidyNamesState && runs < 20) { run(false, "runTidyProductNames(SpreadsheetApp.getActiveSpreadsheet(), Date.now() - 1)", short); runs += 1; }
  eq([labels(sor, 5, 4), labels(my, 5, 3)], [["Meiling 86l", "Meiling 112l", "Midea 195l", "Хөргөгч жижиг"], ["Meiling 86l", "Meiling 112l", "Panasonic 50l"]], "the later parts rename and order every sheet all the same");
  ok(runs > 3 && !state.props.tidyNamesState, "in several parts, and the plan is gone once done");
  run(false, "tidyProductNames()", (c) => short(c).replace("const TIDY_STATE_KEY = 'tidyNamesState';", "const TIDY_STATE_KEY = 'tidyNamesState2';"));
  ok(state.triggers.some((t) => t.getHandlerFunction() === "tidyProductNames"), "an unfinished part starts the next one by itself");
}

{ // the live case of 2026-10-21: the list renamed, the first sheet renamed, the run stopped at the second
  const { list, sor, my, state, run } = world();
  state.props = state.props || {};
  state.docProps.giftColumnReady = "1";
  const short = (c) => c.replace("const TIDY_TIME_BUDGET_MS = 4 * 60 * 1000;", "const TIDY_TIME_BUDGET_MS = -1;");
  run(false, "runTidyProductNames(SpreadsheetApp.getActiveSpreadsheet(), Date.now() - 1)", short); // the list step
  my.cell(5, 1).v = "Panasonic 50l"; // the first sheet got its new name before the stop
  ok(run(false, "[runSync(SpreadsheetApp.getActiveSpreadsheet()).lines[0], runPriceSync(SpreadsheetApp.getActiveSpreadsheet()).cells, runBookJobs(SpreadsheetApp.getActiveSpreadsheet(), Date.now())[0]]").value.every((v) => v === 0 || String(v).includes("нэр засвар")),
     "while a tidy is under way, the product sync, the price sync and the hourly tasks wait");
  list.cell(labels(list, 1, 12).indexOf("Midea 195l") + 1, 1).v = "Cmidea 195l"; // a new name the list does not have (here: put back by hand)
  while (state.props.tidyNamesState) { run(false, "runTidyProductNames(SpreadsheetApp.getActiveSpreadsheet(), Date.now() + 60000)"); }
  eq([labels(my, 5, 3), labels(sor, 5, 4)], [["Meiling 86l", "Meiling 112l", "Panasonic 50l"], ["Cmidea 195l", "Meiling 86l", "Meiling 112l", "Хөргөгч жижиг"]],
     "a rerun goes on from there; a new name the list lacks is never written into a shop sheet");
  ok(!run(false, "isTidyInProgress()").value, "and the syncs run again once it is done");
}

{ // column A's rules are lifted for a move and put back; one sheet's failure does not stop the others
  const { sor, my, state, run } = world();
  sor.getRange(5, 1, 4, 1).setDataValidation("only list names");
  run(false, "runNameOrder(SpreadsheetApp.getActiveSpreadsheet(), { list: null, sheets: [SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sor/2026/10/15')] })");
  ok(labels(sor, 5, 3).join() === "Cmidea 195l,Meiling 86l,Meiling 112l" && sor.getRange(5, 1, 4, 1).getDataValidations().every(([rule]) => rule !== null), "after the move every name cell has its rule again");
  my.getRange = () => { throw new Error("broken sheet"); };
  const lines = run(false, "(() => { const lines = []; onEachSheet(lines, (s) => s.getRange(1, 1))(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('My/2026/10/15')); return lines; })()").value;
  ok(lines.length === 1 && lines[0].includes("алгаслаа"), "a sheet that fails is reported and skipped");
}

{ // two names that are one product: Соронзон's "TGL -12OL" is My home's "TCL 120L"
  const { list, sor, my, state, ss, run } = world();
  sor.cell(6, 1).v = "TGL -12OL"; my.cell(5, 1).v = "TCL 120L";
  list.getRange(6, 1, 1, 2).setValues([["TGL -12OL", "Соронзон"]]); list.getRange(8, 1, 1, 2).setValues([["TCL 120L", "My home"]]);
  run(false, "runMergeNames(SpreadsheetApp.getActiveSpreadsheet(), MERGE_NAMES)");
  eq([sor.cell(6, 1).v, my.cell(5, 1).v], ["TCL 120L", "TCL 120L"], "Соронзон's rows take the other name");
  const names = labels(list, 4, 8);
  ok(!names.includes("TGL -12OL") && names.filter((n) => n === "TCL 120L").length === 1 && list.cell(4 + names.indexOf("TCL 120L"), 2).v === "Хоёулаа", "one list row, for both shops");
  ok(ss.getSheetByName("_Нэр солилт").cell(2, 3).v === "TCL 120L", "and Тосгон follows through the rename log");
}

{ // the renames logged for Тосгон are carried in one pass; a name renamed twice goes to its last name
  const { run } = world();
  eq(run(true, `chainRenames([["cmidea 195l", "Midea 195l"], ["midea 195l", "Midea 195L"], ["panasonica 50l", "Panasonic 50l"]])`).value,
     [["Midea 195L", ["cmidea 195l"]], ["Panasonic 50l", ["panasonica 50l"]]], "chains are followed to the end");
}

{ // the next part is set up before the work, and nothing is left once everything is done
  const { state, run } = world();
  run(false, "tidyProductNames()");
  ok(!state.triggers.some((t) => t.getHandlerFunction() === "tidyProductNames") && !(state.props || {}).tidyNamesState, "a finished tidy leaves no trigger and no plan");
  const started = run(false, "(() => { const before = ScriptApp.getProjectTriggers().length; ScriptApp.newTrigger(TIDY_HANDLER).timeBased().after(TIDY_SAFETY_AFTER_MS).create(); return ScriptApp.getProjectTriggers().length - before; })()").value;
  ok(started === 1 && state.delays.includes(8 * 60 * 1000), "the safety part comes 8 minutes later, after Google's 6");
}

{ // a sheet Google stopped the run on is skipped next time; a tidy past its renames can be stopped
  const { sor, my, state, run } = world();
  run(false, "runTidyProductNames(SpreadsheetApp.getActiveSpreadsheet(), Date.now() + 60000)");
  ok(!state.props.tidyNamesState, "(a normal run first)");
  const stuck = { step: "order", next: 0, trying: 0, renames: [], skipped: [], cells: 0, moved: 0 };
  state.props.tidyNamesState = JSON.stringify(stuck);
  my.cell(5, 1).v = "Zz last"; my.cell(5, 1).v = "Meiling 86l";
  const lines = run(false, "runTidyProductNames(SpreadsheetApp.getActiveSpreadsheet(), Date.now() + 60000)").value.lines;
  ok(lines.some((l) => l.includes("зогссон тул алгаслаа")) && !state.props.tidyNamesState, "the sheet it stopped on is skipped and the run ends");
  state.props.tidyNamesState = JSON.stringify({ ...stuck, step: "shops", trying: null });
  run(false, "stopTidyProductNames()");
  ok(Boolean(state.props.tidyNamesState), "a tidy still renaming is not stopped");
  state.props.tidyNamesState = JSON.stringify({ ...stuck, trying: null });
  run(false, "stopTidyProductNames()");
  ok(!state.props.tidyNamesState, "a tidy that only sorts can be stopped: the syncs run again");
}

{ // names a sheet wrote with other spaces, capitals or the old brand spelling get the list's spelling (2026-10-22)
  const { list, sor, my, run } = world();
  list.getRange(5, 1, 4, 1).setValues([["Meiling 112l"], ["Midea 195l"], ["Meiling 86l"], ["Panasonic 50l"]]);
  sor.cell(5, 1).v = "Meiling  112l "; sor.cell(6, 1).v = "Cmidea 195l"; sor.cell(7, 1).v = "meiling 86l";
  my.cell(5, 1).v = "Panasonica 50l";
  const before = sor.logical();
  const dry = run(true, "runListSpelling(SpreadsheetApp.getActiveSpreadsheet())").value;
  ok(dry.some((line) => line.includes("4 нүд")) && sor.logical() === before, "the trial run counts the cells and writes nothing");
  run(false, "runListSpelling(SpreadsheetApp.getActiveSpreadsheet())");
  eq(labels(sor, 5, 4), ["Meiling 112l", "Midea 195l", "Meiling 86l", "хөргөгч жижиг"], "spaces, capitals and the brand follow the list; a name off the list stays");
  eq([labels(my, 5, 1)[0], sor.cell(5, 2).v, sor.cell(5, 7).v], ["Panasonic 50l", 1, 4], "only the name cell changes");
  my.getRange(9, 1).setValue("Panasonic 50l"); my.cell(5, 1).v = "Panasonica 50l";
  const clash = run(false, "runListSpelling(SpreadsheetApp.getActiveSpreadsheet())").value;
  ok(my.cell(5, 1).v === "Panasonica 50l" && clash.some((line) => line.includes("Panasonica 50l")), "a name the sheet already has elsewhere is left and reported");
}

{ // a sheet in the "Зөвхөн буруу" view: its filter hides rows, and Google lifts the rule only on the rows it shows (2026-10-22)
  const { sor, run } = world();
  const filter = sor.getRange(3, 1, sor.getMaxRows() - 2, 40).createFilter().setColumnFilterCriteria(39, { getCriteriaValues: () => ["буруу"] });
  const filterAtWrite = [];
  const getRange = sor.getRange.bind(sor);
  sor.getRange = (...args) => { const range = getRange(...args); const setValues = range.setValues.bind(range); range.setValues = (values) => { filterAtWrite.push(sor.getFilter()); return setValues(values); }; return range; };
  run(false, "runTidyProductNames(SpreadsheetApp.getActiveSpreadsheet())");
  eq(labels(sor, 5, 4), ["Meiling 86l", "Meiling 112l", "Midea 195l", "Хөргөгч жижиг"], "the sheet is put in order");
  ok(filterAtWrite.length > 0 && filterAtWrite.every((f) => f === null), "with the filter off while the rows are written");
  const back = sor.getFilter();
  ok(back !== null && back !== filter && back.getRange().getRow() === 3 && String(back.getColumnFilterCriteria(39).getCriteriaValues()[0]) === "буруу", "and the same filter is back afterwards");
}

console.log(`ALL ${n} NAME ORDER CHECKS PASSED`);
