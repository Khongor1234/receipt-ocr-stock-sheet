// Checks of the transfer link between the two shops (transfer-link.gs): one shop's transfer out
// (AK) is the other shop's transfer in (U) of the same half-month, a difference is marked pink
// for mom, and a count that is short in one shop and over in the other (from 2026/10/31 on) gets a
// lavender hint, no number. Usage: node tests/transfer_test.js <project dir>
const fs = require("fs"), assert = require("assert");
const { Spreadsheet, makeRunnerFor } = require("./fake_book.js");

const dir = process.argv[2];
const FIRST_FILES = ["profile.gs", "config.gs"];
const gsFiles = fs.readdirSync(dir).filter((f) => f.endsWith(".gs"));
const loadOrder = [...FIRST_FILES.filter((f) => gsFiles.includes(f)), ...gsFiles.filter((f) => !FIRST_FILES.includes(f))];
const code = loadOrder.map((f) => fs.readFileSync(`${dir}/${f}`, "utf8")).join("\n");
const makeRunner = makeRunnerFor(code);
let n = 0;
const ok = (c, m) => { assert(c, m); n++; };
const eq = (a, b, m) => { assert.strictEqual(JSON.stringify(a), JSON.stringify(b), m); n++; };

// The gift layout: U transfer in, AK transfer out, AL gifts, AM count, AN balance.
const C = { opening: 2, in: 21, out: 37, count: 39, balance: 40 };
const TINT = "#fff2cc"; // the transfer columns' usual colour
const AUTO = "#f8bbd0", CONFLICT = "#f06292", NEG = "#d5a6bd", HINT = "#d9d2e9";

// rows: [name, { in, out, count, balance, inNote, outNote }]; the balance stands for the formula's result
function shopSheet(ss, name, rows) {
  const s = ss.insertSheet(name);
  s.maxCols = 48;
  [["Барааны нэрс", 1], ["анхны үлдэгдэл ", 2], ["яг харагдах үнэ", 5], ["Бэлэг", 38], ["тооллого", C.count], ["эцсийн үлдэгдэл ", C.balance]].forEach(([text, c]) => s.getRange(3, c).setValue(text));
  rows.forEach(([label, v], i) => {
    const r = 4 + i;
    s.getRange(r, 1).setValue(label);
    s.getRange(r, C.opening).setValue(5);
    [C.in, C.out].forEach((c) => s.getRange(r, c).setBackground(TINT));
    if (v.in !== undefined) s.getRange(r, C.in).setValue(v.in);
    if (v.out !== undefined) s.getRange(r, C.out).setValue(v.out);
    if (v.inNote) s.getRange(r, C.in).setNotes([[v.inNote]]);
    if (v.outNote) s.getRange(r, C.out).setNotes([[v.outNote]]);
    if (v.count !== undefined) s.getRange(r, C.count).setValue(v.count);
    s.getRange(r, C.balance).setValue(v.balance === undefined ? 5 : v.balance);
  });
  return s;
}
const letter = (column) => String.fromCharCode(64 + Math.floor((column - 1) / 26)).replace("@", "") + String.fromCharCode(65 + ((column - 1) % 26));
const rowOf = (s, label) => s.rows.findIndex((r) => r[1] && r[1].v === label) + 1;
const cellOf = (s, label, column) => {
  const row = s.rows.findIndex((r) => r[1] && r[1].v === label) + 1;
  const x = s.cell(row, column);
  return { value: x.v, bg: x.bg, note: x.note || "" };
};

// ---------- the pure rules ----------
{
  const run = makeRunner(new Spreadsheet(), { triggers: [], mails: [] });
  const v = (e) => run(true, e).value;
  eq(v("[TRANSFER_AUTO_COLOR, TRANSFER_CONFLICT_COLOR]"), [AUTO, CONFLICT], "two pinks: written by the script, and two shops that disagree");
  ok(v("DIVIDER_TRANSFER_COLORS").includes(AUTO) && v("DIVIDER_TRANSFER_COLORS").includes(CONFLICT), "the colour keeper knows both pinks");
  const kept = v(`buildWantedColors([{ row: 4, hasLabel: true, key: "a", color: "#ffffff" }, { row: 5, hasLabel: true, key: "b", color: "#ffffff" }], [["#ffffff", "${AUTO}"], ["#ffffff", "#fff2cc"]], [], {})`);
  eq(kept[0][1], AUTO, "the hourly colour tidy-up keeps a pink transfer cell");
  eq(v(`parseTransferCell(3, "", "#fff2cc")`), { amount: 3, typed: 3, mirrored: false, matched: 0, ocr: false }, "a typed transfer");
  eq(v(`parseTransferCell(2, "Шилжүүлэг (автомат): My home-ийн AK-аас хуулсан", "${AUTO}")`), { amount: 2, typed: null, mirrored: true, matched: 0, ocr: false }, "a copied transfer is not typed");
  eq(v(`parseTransferCell(4, "Шилжүүлэг (автомат): тооллогоор тааруулсан +1", "${AUTO}")`), { amount: 4, typed: 3, mirrored: false, matched: 1, ocr: false }, "a part an earlier version matched is not typed");
  eq(v(`parseTransferCell(2, "Шилжүүлэг (автомат): хасах үлдэгдлээр тааруулсан +1", "${NEG}")`), { amount: 2, typed: 1, mirrored: false, matched: 1, ocr: false }, "nor a part matched by a negative balance");
  eq(v(`parseTransferCell("", "Шилжүүлэг (санал): тооллогоор My home-аас Соронзон руу 1 шилжсэн байж магадгүй", "${HINT}")`), { amount: 0, typed: null, mirrored: false, matched: 0, ocr: false }, "a hint holds no number");
  ok(v("TRANSFER_HINT_COLOR") === HINT && v("DIVIDER_KEPT_COLORS").includes(HINT) && v("DIVIDER_KEPT_COLORS").includes(NEG), "the colour keeper knows the hint's lavender (and the old violet-pink)");
  eq([v("isHintPeriod(sheetSlot('Sor/2026/10/15').position)"), v("isHintPeriod(sheetSlot('Sor/2026/10/31').position)"), v("isHintPeriod(sheetSlot('Sor/2026/4/15').position)")], [false, true, false], "no hints up to 2026/10/15 (the owner, 2026-10-24)");
  eq(v(`parseTransferCell("", "", "#fff2cc")`), { amount: 0, typed: null, mirrored: false, matched: 0, ocr: false }, "an empty cell");
  eq(v(`parseTransferCell(2, "OCR: IMG_1 [abc#1] = 2", "#ffe599")`).ocr, true, "a cell the receipt OCR wrote");
  eq(v(`parseTransferCell(2, "Зураг: https://drive.google.com/file/d/abc/view#1 = 2 (2026-10-03)", "#ffe599")`).ocr, true, "... also in its note of 2026-10-25 (the photo's link only)");
  eq(v(`parseTransferCell(2, "https://docs.google.com/spreadsheets/d/bk/edit#gid=7&range=A30:M30&ocr=abc.1 = 2 (2026-10-23)", "#ffe599")`).ocr, true, "... and in its note of 2026-10-26 (the link to the line's row)");
}

// ---------- one half-month of both shops ----------
const state = { triggers: [], mails: [], docProps: { pairedLayoutReady: "1", giftColumnReady: "1" } };
const ss = new Spreadsheet();
const my = shopSheet(ss, "My/2026/10/15", [
  ["Copied", { out: 2 }],
  ["Same", { in: 1 }],
  ["Differ", { out: 3 }],
  ["Matched", { count: 4, balance: 5 }],
  ["By OCR", { out: 2, outNote: "OCR: IMG_1 [abc#1] = 2" }],
  ["Only here", { out: 2 }],
  ["From Тосгон", { in: 4 }],
  ["Not counted", { count: 1, balance: 3 }],
]);
const sor = shopSheet(ss, "Sor/2026/10/15", [
  ["Copied", {}],
  ["Same", { out: 1 }],
  ["Differ", { in: 2 }],
  ["Matched", { count: 3, balance: 2 }],
  ["By OCR", {}],
  ["From Тосгон", {}],
  ["Not counted", { balance: 1 }],
]);
const run = makeRunner(ss, state);
const position = run(true, `sheetSlot("My/2026/10/15").position`).value;

{
  const dry = run(true, `runTransferLink(ss, [${position}], { hint: true }, Date.now() + 60000)`);
  ok(dry.value.join("\n").includes("хуулна") && cellOf(sor, "Copied", C.in).value === "", "a dry run only reports");
}

const real = run(false, `runTransferLink(ss, [${position}], { hint: true }, Date.now() + 60000)`).value.join("\n");
{
  const x = cellOf(sor, "Copied", C.in);
  eq([x.value, x.bg], [2, AUTO], "My home's transfer out is Соронзон's transfer in, pink");
  ok(x.note.includes("My home") && x.note.includes("хуулсан"), "the note says where it came from");
  eq(x.note.slice(x.note.indexOf("https://")), `https://docs.google.com/spreadsheets/d/${ss.getId()}/edit#gid=${my.getSheetId()}&range=${letter(C.out)}${rowOf(my, "Copied")}`,
     "and links to My home's transfer out cell (2026-10-26)");
  eq(cellOf(my, "Copied", C.out).value, 2, "the typed side stays");
  eq([cellOf(my, "Same", C.in).value, cellOf(my, "Same", C.in).bg, cellOf(sor, "Same", C.out).bg], [1, TINT, TINT], "two sides that agree are left alone");
  const a = cellOf(my, "Differ", C.out), b = cellOf(sor, "Differ", C.in);
  eq([a.value, b.value, a.bg, b.bg], [3, 2, CONFLICT, CONFLICT], "two sides that disagree keep their numbers and turn pink for mom");
  ok(a.note.includes("3") && a.note.includes("2") && a.note.split(" Соронзон: https://")[0] === b.note.split(" My home: https://")[0], "both notes give both numbers");
  ok(a.note.endsWith(`range=${letter(C.in)}${rowOf(sor, "Differ")}`) && b.note.endsWith(`range=${letter(C.out)}${rowOf(my, "Differ")}`), "each links to the other shop's cell");
  const mOut = cellOf(my, "Matched", C.out), sIn = cellOf(sor, "Matched", C.in);
  eq([mOut.value, sIn.value, mOut.bg, sIn.bg, mOut.note, sIn.note], ["", "", TINT, TINT, "", ""], "1 short in My home and 1 over in Соронзон on 10/15: no number is added, and no hint before 10/31");
  const ocrOut = cellOf(my, "By OCR", C.out), copied = cellOf(sor, "By OCR", C.in);
  eq([ocrOut.value, ocrOut.note, ocrOut.bg], [2, "OCR: IMG_1 [abc#1] = 2", TINT], "a cell the receipt OCR wrote is never touched");
  eq([copied.value, copied.bg], [2, AUTO], "the other side follows the OCR's number");
  const tosgon = cellOf(my, "From Тосгон", C.in);
  eq([tosgon.value, tosgon.bg, cellOf(sor, "From Тосгон", C.out).value], [4, TINT, ""], "a transfer in with no transfer out on the other side may come from Тосгон: it is not copied back");
  const lone = cellOf(my, "Only here", C.out);
  eq([lone.value, lone.bg], [2, CONFLICT], "a transfer to a product the other shop does not list is marked for mom");
  ok(lone.note.includes("алга"), "the note says the product is missing there");
  eq([cellOf(my, "Not counted", C.out).value, cellOf(sor, "Not counted", C.in).value], ["", ""], "no match while one shop has not counted");
  ok(real.includes("хуулсан 2") && !real.includes("санал") && real.includes("зөрүүтэй 2"), "the log counts what was done");
}

{
  const again = run(false, `runTransferLink(ss, [${position}], { hint: true }, Date.now() + 60000)`).value.join("\n");
  ok(!again.includes("хуулсан") && !again.includes("санал"), "a second run changes nothing");
  eq([cellOf(my, "Matched", C.out).value, cellOf(sor, "Matched", C.in).value], ["", ""], "still no number");
}

// the typed side is changed, then cleared
my.getRange(4, C.out).setValue(1);
run(false, `runTransferLink(ss, [${position}], { hint: true }, Date.now() + 60000)`);
eq(cellOf(sor, "Copied", C.in).value, 1, "a changed transfer out is copied again");
my.getRange(4, C.out).setValue("");
run(false, `runTransferLink(ss, [${position}], { hint: true }, Date.now() + 60000)`);
{
  const x = cellOf(sor, "Copied", C.in);
  eq([x.value, x.bg, x.note], ["", TINT, ""], "a cleared transfer out clears the copy and its pink");
}

// someone types over a copied cell: it is a typed number now, so the two sides disagree
my.getRange(4, C.out).setValue(2);
run(false, `runTransferLink(ss, [${position}], { hint: true }, Date.now() + 60000)`);
sor.getRange(4, C.in).setValue(5);
run(false, `runTransferLink(ss, [${position}], { hint: true }, Date.now() + 60000)`);
eq([cellOf(sor, "Copied", C.in).value, cellOf(sor, "Copied", C.in).bg, cellOf(my, "Copied", C.out).value], [5, CONFLICT, 2], "a number typed over a copy is not overwritten; both turn pink");

// the disagreement is settled by mom: the pink goes
my.getRange(4, C.out).setValue(5);
run(false, `runTransferLink(ss, [${position}], { hint: true }, Date.now() + 60000)`);
eq([cellOf(sor, "Copied", C.in).bg, cellOf(my, "Copied", C.out).bg, cellOf(my, "Copied", C.out).note], [TINT, TINT, ""], "once both sides agree the pink and the note go");

// mirror only (the one-time copy before the uncounted counts are filled again)
{
  const ss2 = new Spreadsheet();
  const a = shopSheet(ss2, "My/2026/10/15", [["Matched", { count: 4, balance: 5 }], ["Copied", { out: 1 }]]);
  const b = shopSheet(ss2, "Sor/2026/10/15", [["Matched", { count: 3, balance: 2 }], ["Copied", {}]]);
  const run2 = makeRunner(ss2, { triggers: [], mails: [], docProps: { pairedLayoutReady: "1", giftColumnReady: "1" } });
  run2(false, `runTransferLink(ss, [${position}], { hint: false }, Date.now() + 60000)`);
  eq([cellOf(a, "Matched", C.out).value, cellOf(b, "Copied", C.in).value], ["", 1], "without matching only the copies are made");
}

// a person's edit of a copied cell drops the copy mark at once (onSheetEdit)
{
  const run3 = makeRunner(ss, state);
  sor.getRange(4, C.in).setNotes([["Шилжүүлэг (автомат): My home-ийн AK-аас хуулсан"]]);
  run3(false, `forgetTransferMarks(ss.getSheetByName("Sor/2026/10/15").getRange(4, ${C.in}))`);
  eq(cellOf(sor, "Copied", C.in).note, "", "the copy mark is gone, so the number counts as typed");
}

// a typed transfer out is copied at once by the edit trigger (onSheetEdit)
{
  const ss4 = new Spreadsheet();
  const a = shopSheet(ss4, "My/2026/10/15", [["Typed", { out: 3 }]]);
  const b = shopSheet(ss4, "Sor/2026/10/15", [["Typed", {}]]);
  const state4 = { triggers: [], mails: [], docProps: { pairedLayoutReady: "1", giftColumnReady: "1" } };
  state4.editEvent = { range: a.getRange(4, C.out), value: 3, oldValue: undefined, source: ss4 };
  makeRunner(ss4, state4)(false, "onSheetEdit(editEvent)");
  eq([cellOf(b, "Typed", C.in).value, cellOf(b, "Typed", C.in).bg], [3, AUTO], "typing 3 in My home's AK puts 3 in Соронзон's U at once");
}

// an earlier version copied a transfer in back into the other shop's transfer out: undone
{
  const ss6 = new Spreadsheet();
  const a = shopSheet(ss6, "My/2026/10/15", [["Haier", { out: 14, outNote: "Шилжүүлэг (автомат): Соронзон-ийн U-аас хуулсан (14)" }], ["Other", {}]]);
  const b = shopSheet(ss6, "Sor/2026/10/15", [["Haier", { in: 14 }], ["Other", {}]]);
  a.getRange(4, C.out).setBackground(AUTO);
  makeRunner(ss6, { triggers: [], mails: [], docProps: {} })(false, `runTransferLinkAll(ss, { hint: false })`);
  eq([cellOf(a, "Haier", C.out).value, cellOf(a, "Haier", C.out).note, cellOf(a, "Haier", C.out).bg, cellOf(b, "Haier", C.in).value], ["", "", TINT, 14], "the wrong copy goes, the typed transfer in stays");
}

// a sheet with the "Зөвхөн буруу" filter: Apps Script writes no note or colour into a row the
// filter hides, so the filter is taken off while writing and put back as it was
{
  const ss7 = new Spreadsheet();
  const a = shopSheet(ss7, "My/2026/10/15", [["Sent", { out: 2 }], ["Other", {}]]);
  const b = shopSheet(ss7, "Sor/2026/10/15", [["Sent", {}], ["Other", {}]]);
  b.getRange(3, 1, 10, 48).createFilter().setColumnFilterCriteria(41, "буруу");
  const before = b.getFilter();
  makeRunner(ss7, { triggers: [], mails: [], docProps: {} })(false, `runTransferLinkAll(ss, { hint: false })`);
  const after = b.getFilter();
  eq([cellOf(b, "Sent", C.in).value, cellOf(b, "Sent", C.in).bg], [2, AUTO], "the copy is written");
  ok(after !== null && after !== before && after.getColumnFilterCriteria(41) === "буруу" && after.getRange().getRow() === 3, "the filter was lifted and put back with its criteria");
  ok(a.getFilter() === null, "a sheet without a filter gets none");
}

// the whole year (by hand): a transfer out to a product the other shop lacks is listed, not marked
{
  const ss5 = new Spreadsheet();
  const a = shopSheet(ss5, "My/2026/10/15", [["Only here", { out: 2 }], ["Both", {}]]);
  shopSheet(ss5, "Sor/2026/10/15", [["Both", {}]]);
  const log = makeRunner(ss5, { triggers: [], mails: [], docProps: {} })(false, `runTransferLinkAll(ss, { hint: true })`).value.join("\n");
  eq([cellOf(a, "Only here", C.out).bg, cellOf(a, "Only here", C.out).note], [TINT, ""], "the old sheets are not marked");
  // a mark an earlier version left on an old sheet goes
  a.getRange(4, C.out).setNotes([["Шилжүүлэг зөрүүтэй: Соронзон-д \"Only here\" бараа алга. Тэнд мөр нэмэх эсвэл энэ тоог засна уу."]]).setBackground(CONFLICT);
  makeRunner(ss5, { triggers: [], mails: [], docProps: {} })(false, `runTransferLinkAll(ss, { hint: true })`);
  eq([cellOf(a, "Only here", C.out).value, cellOf(a, "Only here", C.out).bg, cellOf(a, "Only here", C.out).note], [2, TINT, ""], "an old missing-row mark is cleared, the number stays");
  ok(log.includes("мөргүй 1") && log.includes("Only here") && log.includes("Бүх хагас сар"), "but the log lists it");
}

// from 2026/10/31 on (the owner, 2026-10-24): a count short in one shop and over in the other gets a
// lavender hint on both sides, and no number. A negative balance is no reason for anything any more.
// A transfer an earlier version matched loses its added part on every half-month.
{
  const OLD = "Шилжүүлэг (автомат): тооллогоор тааруулсан +2";
  const OLD_NEG = "Шилжүүлэг (автомат): хасах үлдэгдлээр тааруулсан +1";
  const ss8 = new Spreadsheet();
  const a = shopSheet(ss8, "My/2026/10/31", [
    ["Hinted", { count: 4, balance: 5 }],
    ["Negative", { count: 1, balance: 3 }],
    ["Old match", { out: 3, outNote: OLD }],
    ["Old negative", { in: 1, inNote: OLD_NEG }],
  ]);
  const b = shopSheet(ss8, "Sor/2026/10/31", [
    ["Hinted", { count: 3, balance: 2 }],
    ["Negative", { balance: -1 }],
    ["Old match", { in: 2, inNote: OLD }],
    ["Old negative", { out: 1, outNote: OLD_NEG }],
  ]);
  [[a, 6, C.out], [b, 6, C.in]].forEach(([s, r, c]) => s.getRange(r, c).setBackground(AUTO));
  [[a, 7, C.in], [b, 7, C.out]].forEach(([s, r, c]) => s.getRange(r, c).setBackground(NEG));
  const run8 = makeRunner(ss8, { triggers: [], mails: [], docProps: {} });
  const pos = run8(true, `sheetSlot("My/2026/10/31").position`).value;
  const log = run8(false, `runTransferLink(ss, [${pos}], { hint: true }, Date.now() + 60000)`).value.join("\n");
  const out = cellOf(a, "Hinted", C.out), into = cellOf(b, "Hinted", C.in);
  eq([out.value, into.value, out.bg, into.bg], ["", "", HINT, HINT], "1 short in My home and 1 over in Соронзон: both cells lavender, no number");
  ok(out.note.includes("санал") && out.note.includes("1") && into.note === out.note, "the note says what may have moved");
  const neg = [cellOf(a, "Negative", C.out), cellOf(b, "Negative", C.in)];
  eq(neg.map((x) => [x.value, x.bg, x.note]), [["", TINT, ""], ["", TINT, ""]], "a negative balance gets nothing");
  const oldOut = cellOf(a, "Old match", C.out), oldIn = cellOf(b, "Old match", C.in);
  eq([oldOut.value, oldOut.bg, oldOut.note], [1, TINT, ""], "an old match loses its +2: the typed 1 stays");
  eq([oldIn.value, oldIn.bg], [1, AUTO], "and the other side follows it as a plain copy");
  ok(!oldIn.note.includes("тааруулсан"), "the match line is gone");
  const oldNeg = [cellOf(a, "Old negative", C.in), cellOf(b, "Old negative", C.out)];
  eq(oldNeg.map((x) => [x.value, x.bg, x.note]), [["", TINT, ""], ["", TINT, ""]], "an old violet-pink match goes entirely");
  ok(log.includes("санал тавьсан 1") && log.includes("тааруулсныг арилгасан 2"), "the log counts the hints and the removed matches");
  // the balances follow the removed matches
  a.getRange(6, C.balance).setValue(7);
  b.getRange(6, C.balance).setValue(4);
  const again = run8(false, `runTransferLink(ss, [${pos}], { hint: true }, Date.now() + 60000)`).value.join("\n");
  ok(!again.includes("санал") && !again.includes("арилгасан") && !again.includes("хуулсан"), "a second run changes nothing");
  eq([cellOf(a, "Hinted", C.out).bg, cellOf(b, "Hinted", C.in).bg], [HINT, HINT], "the hint stays while the counts disagree");
  // mom writes the transfer by hand: the hint goes
  a.getRange(4, C.out).setValue(1);
  a.getRange(4, C.balance).setValue(4);
  b.getRange(4, C.balance).setValue(2);
  run8(false, `runTransferLink(ss, [${pos}], { hint: true }, Date.now() + 60000)`);
  const typed = cellOf(a, "Hinted", C.out), copy = cellOf(b, "Hinted", C.in);
  eq([typed.value, typed.bg, typed.note], [1, TINT, ""], "the typed transfer is plain");
  eq([copy.value, copy.bg], [1, AUTO], "and copied pink as usual, without the hint");
  ok(!copy.note.includes("санал"), "the hint line is gone");
}

// a one-off by hand: removeCountMatches takes every old match out of the whole year at once,
// without the rest of the link (no copies, no marks)
{
  const ss9 = new Spreadsheet();
  const a = shopSheet(ss9, "My/2026/4/30", [
    ["Matched", { out: 3, outNote: "Шилжүүлэг (автомат): тооллогоор тааруулсан +3" }],
    ["Typed", { out: 2 }],
    ["Copy and match", { in: 3, inNote: "Шилжүүлэг (автомат): Соронзон-ийн AK-аас хуулсан (1)\nШилжүүлэг (автомат): тооллогоор тааруулсан +2" }],
  ]);
  const b = shopSheet(ss9, "Sor/2026/8/31", [
    ["Negative", { in: 1, inNote: "Шилжүүлэг (автомат): хасах үлдэгдлээр тааруулсан +1" }],
    ["Other", { in: 1, inNote: "mom's note" }],
  ]);
  a.getRange(4, C.out).setBackground(AUTO);
  a.getRange(6, C.in).setBackground(AUTO);
  b.getRange(4, C.in).setBackground(NEG);
  const run9 = makeRunner(ss9, { triggers: [], mails: [], docProps: {} });
  const dry = run9(true, "runRemoveCountMatches(ss)").value.join("\n");
  ok(dry.includes("My/2026/4/30") && cellOf(a, "Matched", C.out).value === 3, "a dry run only reports");
  const log = run9(false, "runRemoveCountMatches(ss)").value.join("\n");
  const m = cellOf(a, "Matched", C.out);
  eq([m.value, m.bg, m.note], ["", TINT, ""], "a cell that held only a match is emptied");
  eq([cellOf(a, "Typed", C.out).value, cellOf(a, "Typed", C.out).bg], [2, TINT], "a typed cell is left alone");
  const both = cellOf(a, "Copy and match", C.in);
  eq([both.value, both.bg, both.note], [1, AUTO, "Шилжүүлэг (автомат): Соронзон-ийн AK-аас хуулсан (1)"], "a copy keeps its copied part, its line and its pink");
  const n = cellOf(b, "Negative", C.in);
  eq([n.value, n.bg, n.note], ["", TINT, ""], "a violet-pink match goes");
  eq([cellOf(b, "Other", C.in).value, cellOf(b, "Other", C.in).note], [1, "mom's note"], "mom's own note stays");
  ok(log.includes("My/2026/4/30: 2") && log.includes("Sor/2026/8/31: 1"), "the log counts the cells per sheet");
  ok(run9(false, "runRemoveCountMatches(ss)").value.join("\n").includes("олдсонгүй"), "a second run finds none");
}

// the half-months before 2026/4/15 keep their old matches (the owner, 2026-10-24): neither the
// one-off nor the link (an edit, linkShopTransfers) takes them out
{
  const OLD = "Шилжүүлэг (автомат): тооллогоор тааруулсан +1";
  const ss10 = new Spreadsheet();
  const a = shopSheet(ss10, "My/2026/3/15", [["Kept", { in: 1, inNote: OLD }], ["Typed", { out: 2 }]]);
  const b = shopSheet(ss10, "Sor/2026/3/15", [["Kept", { out: 1, outNote: OLD }], ["Typed", {}]]);
  a.getRange(4, C.in).setBackground(AUTO);
  b.getRange(4, C.out).setBackground(AUTO);
  const run10 = makeRunner(ss10, { triggers: [], mails: [], docProps: {} });
  const v = (e) => run10(true, e).value;
  eq([v("keepsOldMatches(sheetSlot('Sor/2026/3/31').position)"), v("keepsOldMatches(sheetSlot('Sor/2026/4/15').position)")], [true, false], "3/31 and before keep them, 4/15 on not");
  ok(run10(false, "runRemoveCountMatches(ss)").value.join("\n").includes("олдсонгүй"), "the one-off leaves 3/15 alone");
  const pos = v(`sheetSlot("My/2026/3/15").position`);
  run10(false, `runTransferLink(ss, [${pos}], { hint: true }, Date.now() + 60000)`);
  const kept = [cellOf(a, "Kept", C.in), cellOf(b, "Kept", C.out)];
  eq(kept.map((x) => [x.value, x.bg, x.note]), [[1, AUTO, OLD], [1, AUTO, OLD]], "the link leaves the old match of 3/15 as it is");
  eq([cellOf(b, "Typed", C.in).value, cellOf(b, "Typed", C.in).bg], [2, AUTO], "and still copies the other transfers there");
}

// the colour tidy-up (dividers.gs) on a sheet with the "Зөвхөн буруу" filter: the rows the filter
// hides get their colours too (the grey of My/2026/9/15 stayed there, 2026-10-24)
{
  const GREY = "#d0e0e3";
  const ss12 = new Spreadsheet();
  const a = shopSheet(ss12, "My/2026/9/15", [["Hidden", { count: 5, balance: 5 }], ["Shown", { count: 1, balance: 5 }], ["Other", {}]]);
  a.getRange(4, 41).setValue("зөв");
  a.getRange(5, 41).setValue("буруу");
  a.getRange(4, C.count).setBackground(GREY);
  a.getRange(3, 1, 10, 48).createFilter().setColumnFilterCriteria(41, "буруу");
  makeRunner(ss12, { triggers: [], mails: [], docProps: {} })(false, `reconcileSheet(ss.getSheetByName("My/2026/9/15"))`);
  ok(cellOf(a, "Hidden", C.count).bg !== GREY, "the grey of a row the filter hides goes");
  const after = a.getFilter();
  ok(after !== null && after.getColumnFilterCriteria(41) === "буруу", "the filter is put back as it was");
}

console.log(`transfer_test: ${n} checks passed`);
