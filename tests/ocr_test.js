// Checks of the receipt OCR: the pure decisions, then whole runs against a fake Drive and a fake
// Gemini. Usage: node tests/ocr_test.js <project dir> [--bundle]
const fs = require("fs"), vm = require("vm"), assert = require("assert"), crypto = require("crypto"), zlib = require("zlib");

const dir = process.argv[2];
const FILES = ["profile.gs", "config.gs", "year-books.gs", "ocr/ocr-config.gs", "ocr/ocr-validate.gs", "ocr/ocr-review.gs", "ocr/ocr-gemini.gs", "ocr/ocr-catalog.gs", "ocr/ocr-files.gs", "ocr/ocr-sheet.gs", "ocr/ocr-post.gs", "ocr/ocr-setup.gs", "ocr/ocr-watch.gs", "ocr/ocr-mirror.gs", "ocr/ocr-orphans.gs", "ocr/ocr-spot.gs", "ocr/ocr-instant.gs", "ocr/ocr-split.gs", "ocr/ocr-reread.gs", "ocr/ocr-photo-links.gs", "ocr/ocr-check.gs", "ocr/ocr-elsewhere.gs", "ocr/ocr-reconcile.gs", "ocr/ocr-layout.gs", "ocr/ocr-calendar.gs", "ocr/ocr-suggest.gs", "ocr/ocr-backfill.gs", "ocr/ocr-report.gs", "ocr/ocr-count-read.gs", "ocr/ocr-count-report.gs", "ocr/ocr-twins.gs", "ocr/ocr-main.gs", "dividers.gs", "row-copies.gs", "ocr/ocr-new-product.gs", "ocr/ocr-relink.gs"];
const isBundle = process.argv[3] === "--bundle";
// The registry of the years' files (year-books.gs) is left out here: the fake Drive has none.
const shippedCode = (isBundle ? fs.readFileSync(`${dir}/neg-file/ocr-all-in-one.gs`, "utf8") : FILES.map((f) => fs.readFileSync(`${dir}/${f}`, "utf8")).join("\n"))
  .replace(/const REGISTRY_ID = '[^']*';/, "const REGISTRY_ID = '';");
assert(shippedCode.includes("const DRY_RUN = true;"), "DRY_RUN must ship as true");
// The fake workbooks below use the shops' single-block layout (days from E); the OCR reads every
// column from LAYOUT, so these checks run with PAIRED_SHOPS = false. The paired layout (days from
// F, after the shelf price) is checked at the end with its own fake workbook.
assert(shippedCode.includes("const PAIRED_SHOPS = true;"), "the shops ship with the paired layout");
const code = shippedCode.replace("const PAIRED_SHOPS = true;", "const PAIRED_SHOPS = false;").replace("const OCR_TOSGON_SHELF = true;", "const OCR_TOSGON_SHELF = false;");
// The paired checks below are the move as it was made, before the gift column (GIFT_COLUMN);
// the gift column's own checks use the shipped code as it is (giftCode).
assert(shippedCode.includes("const GIFT_COLUMN = true;"), "the shops ship with the gift column");
// Тосгон's shelf price in E (2026-10-22): until its own checks at the end, Тосгон has its old
// columns (days from E) in every fake workbook here.
assert(shippedCode.includes("const OCR_TOSGON_SHELF = true;"), "the OCR ships with Тосгон's shelf price column");
const tosgonOld = (text) => text.replace("const OCR_TOSGON_SHELF = true;", "const OCR_TOSGON_SHELF = false;");
const pairedCode = tosgonOld(shippedCode.replace("const GIFT_COLUMN = true;", "const GIFT_COLUMN = false;"));
const giftCode = tosgonOld(shippedCode);
const shelfCode = shippedCode;
assert(code.includes("const OCR_POST_FROM = '2026-10-20';"), "the tests set the start day through this line");
let n = 0;
const ok = (cond, msg) => { assert(cond, msg); n++; };
const eq = (a, b, msg) => { assert.deepStrictEqual(a, b, msg); n++; };
const plain = (value) => JSON.parse(JSON.stringify(value)); // values made inside the vm

// ---------- fake Drive ----------
const iter = (list) => { let i = 0; return { hasNext: () => i < list.length, next: () => list[i++] }; };
let nextId = 0;
const trashBin = []; // Drive still finds a trashed file by its id
class File {
  constructor(name, mime, bytes, created, owner = "me") { Object.assign(this, { id: `f${++nextId}`, name, mime, bytes: Buffer.from(bytes), created, owner, parent: null, editors: [] }); }
  getParents() { return iter(this.parent ? [this.parent] : []); }
  getName() { return this.name; } setName(name) { this.name = name; return this; } getMimeType() { return this.mime; } getId() { return this.id; }
  getUrl() { return `https://drive.google.com/file/d/${this.id}/view`; } getDateCreated() { return new Date(Date.UTC(2026, 9, 10) + this.created * 3600000); }
  getLastUpdated() { return this.getDateCreated(); }
  setTrashed(isTrashed) { this.trashed = isTrashed; if (isTrashed && this.parent) { this.parent.files = this.parent.files.filter((f) => f !== this); trashBin.push(this); } return this; }
  isTrashed() { return Boolean(this.trashed); }
  getEditors() { return this.editors.map((email) => ({ getEmail: () => email })); }
  addEditor(email) { this.editors.push(email); return this; } removeEditor(email) { this.editors = this.editors.filter((e) => e !== email); return this; }
  getBlob() { const bytes = this.bytes; return { getBytes: () => [...bytes].map((b) => (b > 127 ? b - 256 : b)) }; }
  moveTo(folder) { if (this.owner !== "me") throw new Error("only the owner may move"); if (this.parent) this.parent.files = this.parent.files.filter((f) => f !== this); folder.add(this); }
  makeCopy(name, folder) { return folder.add(new File(name, this.mime, this.bytes, this.created)); }
}
class Folder {
  constructor(name, parent = null) { Object.assign(this, { id: `d${++nextId}`, name, parent, folders: [], files: [], viewers: [] }); }
  getViewers() { return this.viewers.map((email) => ({ getEmail: () => email })); } addViewer(email) { this.viewers.push(email); return this; }
  getName() { return this.name; } getId() { return this.id; } getParents() { return iter(this.parent ? [this.parent] : []); }
  getFolders() { return iter(this.folders.slice()); } getFiles() { return iter(this.files.slice()); }
  getFoldersByName(name) { return iter(this.folders.filter((f) => f.name === name)); }
  getFilesByName(name) { return iter(this.files.filter((f) => f.name === name)); }
  createFolder(name) { const f = new Folder(name, this); this.folders.push(f); return f; }
  allFiles() { return [...this.files, ...this.folders.flatMap((f) => f.allFiles())]; }
  add(file) { file.parent = this; this.files.push(file); return file; }
  path(...names) { return names.reduce((folder, name) => folder.folders.find((f) => f.name === name) || folder.createFolder(name), this); }
  find(...names) { return names.reduce((folder, name) => folder && folder.folders.find((f) => f.name === name), this) || null; }
}

// ---------- fake spreadsheets ----------
class Protection {
  constructor(sheet, type) { Object.assign(this, { sheet, type, description: "", editors: [], unprotected: [] }); }
  setDescription(d) { this.description = d; return this; } getDescription() { return this.description; }
  setUnprotectedRanges(r) { this.unprotected = r; return this; }
  addEditor(email) { if (!this.editors.includes(email)) this.editors.push(email); return this; }
  removeEditors(users) { const gone = users.map((u) => u.getEmail()); this.editors = this.editors.filter((e) => !gone.includes(e)); return this; }
  getEditors() { return [...this.editors, "someone@else"].map((email) => ({ getEmail: () => email })); }
  canDomainEdit() { return false; } remove() { this.sheet.protections = this.sheet.protections.filter((p) => p !== this); }
}
class Sheet {
  constructor(cells = []) { Object.assign(this, { cells, frozen: 0, rules: [], images: [], heights: {}, notes: {}, fonts: {}, backgrounds: [], validations: [], protections: [], hidden: [] }); }
  insertImage(blob, column, row) {
    const sheet = this;
    const image = { blob, column, row, w: 750, h: 1000, remove() { sheet.images = sheet.images.filter((other) => other !== this); }, getWidth() { return this.w; }, getHeight() { return this.h; }, setWidth(w) { this.w = w; return this; }, setHeight(h) { this.h = h; return this; },
      getAnchorCell() { return { getRow: () => this.row, getColumn: () => this.column }; }, setAnchorCell(range) { this.row = range.r; this.column = range.c; return this; } };
    this.images.push(image); return image;
  }
  getImages() { return this.images.slice(); }
  setRowHeights(row, count, height) { for (let r = row; r < row + count; r++) this.heights[r] = height; return this; }
  getParent() { return this.parent; } getName() { return this.name; } setName(n) { this.name = n; return this; } getSheetId() { this.sid = this.sid || ++nextId; return this.sid; }
  getLastColumn() { return Math.max(0, ...Array.from(this.cells, (row) => (row || []).length)); }
  getMaxColumns() { return Math.max(this.maxCols || 0, this.getLastColumn()); }
  setRowHeight(row, height) { this.heights[row] = height; return this; }
  getRowHeight(row) { return this.heights[row] || 21; }
  insertRowsAfter() { return this; }
  getFilter() { return null; }
  // Like Google Sheets: the rows of `range` before row `dest` (counted as before the move), all they hold with them.
  moveRows(range, dest) {
    const height = this.getLastRow();
    const order = Array.from({ length: height }, (_, i) => i + 1).filter((r) => r < range.r || r >= range.r + range.h);
    const moving = Array.from({ length: range.h }, (_, i) => range.r + i);
    const at = order.filter((r) => r < dest).length;
    const final = [...order.slice(0, at), ...moving, ...order.slice(at)];
    const to = Object.fromEntries(final.map((r, i) => [r, i + 1]));
    const where = (r) => to[r] || r;
    this.cells = final.map((r) => this.cells[r - 1] || []);
    const moveKeys = (map) => Object.fromEntries(Object.entries(map).map(([k, v]) => { const [r, c] = k.split(",").map(Number); return [`${where(r)},${c}`, v]; }));
    this.notes = moveKeys(this.notes); this.fonts = moveKeys(this.fonts);
    this.backgrounds = this.backgrounds.map((b) => (b.h === 1 ? { ...b, r: where(b.r) } : b));
    return this;
  }
  // Like Google Sheets: a row after `row`, everything below one row down (values, notes, fonts, colours, lists).
  insertRowAfter(row) {
    this.cells.splice(row, 0, []);
    const down = (r) => (r > row ? r + 1 : r);
    const moveKeys = (map) => Object.fromEntries(Object.entries(map).map(([k, v]) => { const [r, c] = k.split(",").map(Number); return [`${down(r)},${c}`, v]; }));
    this.notes = moveKeys(this.notes); this.fonts = moveKeys(this.fonts);
    this.backgrounds = this.backgrounds.map((b) => ({ ...b, r: down(b.r) }));
    this.validations = this.validations.map((v) => ({ ...v, r: down(v.r) }));
    this.inserted = [...(this.inserted || []), row + 1];
    return this;
  }
  setColumnWidths() { return this; }
  setColumnWidth(column, width) { this.widths = { ...(this.widths || {}), [column]: width }; return this; }
  getColumnWidth(column) { return (this.widths || {})[column] || 100; }
  getMaxRows() { return 1000; } setFrozenRows(n) { this.frozen = n; }
  getLastRow() { return this.cells.length; } clear() { this.cells = []; }
  // Like Google Sheets: everything right of `col` moves n columns further (values, notes, fonts, colours).
  insertColumnsAfter(col, n) {
    this.maxCols = this.getMaxColumns() + n;
    const move = (c) => (c > col ? c + n : c);
    this.cells.forEach((row) => { if (row && row.length > col) row.splice(col, 0, ...Array(n).fill("")); });
    const moveKeys = (map) => Object.fromEntries(Object.entries(map).map(([k, v]) => { const [r, c] = k.split(",").map(Number); return [`${r},${move(c)}`, v]; }));
    this.notes = moveKeys(this.notes); this.fonts = moveKeys(this.fonts);
    this.backgrounds = this.backgrounds.map((b) => ({ ...b, c: move(b.c) }));
    return this;
  }
  hideColumns(c, count = 1) { for (let i = 0; i < count; i++) if (!this.hidden.includes(c + i)) this.hidden.push(c + i); }
  showColumns(c, count = 1) { this.hidden = this.hidden.filter((h) => h < c || h >= c + count); }
  protect() { const p = new Protection(this, "SHEET"); this.protections.push(p); return p; }
  hideSheet() { this.isHidden = true; }
  getProtections(type) { return this.protections.filter((p) => p.type === type); }
  getConditionalFormatRules() { return this.rules.slice(); } setConditionalFormatRules(r) { this.rules = r; }
  value(r, c) { return (this.cells[r - 1] || [])[c - 1] ?? ""; }
  getRange(r, c, h = 1, w = 1) {
    const sheet = this, key = `${r},${c}`;
    return {
      r, c, h, w,
      setValues(values) { assert(values.length === h && values.every((row) => row.length === w), "setValues size"); values.forEach((row, i) => { sheet.cells[r - 1 + i] = sheet.cells[r - 1 + i] || []; row.forEach((v, j) => { sheet.cells[r - 1 + i][c - 1 + j] = v; }); }); return this; },
      getValues() { return Array.from({ length: h }, (_, i) => Array.from({ length: w }, (_, j) => sheet.value(r + i, c + j))); },
      getValue() { return sheet.value(r, c); },
      setValue(v) { return this.setValues([[v]]); },
      setFormula(f) { return this.setValues(Array.from({ length: h }, () => Array.from({ length: w }, () => f))); },
      setFormulas(fs) { return this.setValues(fs); },
      getNote() { return sheet.notes[key] || ""; }, setNote(n) { sheet.notes[key] = n; return this; },
      getNotes() { return Array.from({ length: h }, (_, i) => Array.from({ length: w }, (_, j) => sheet.notes[`${r + i},${c + j}`] || "")); },
      getFontColors() { return Array.from({ length: h }, (_, i) => Array.from({ length: w }, (_, j) => sheet.fonts[`${r + i},${c + j}`] || null)); },
      setFontColors(colors) { colors.forEach((row, i) => row.forEach((color, j) => { if (color) sheet.fonts[`${r + i},${c + j}`] = color; else delete sheet.fonts[`${r + i},${c + j}`]; })); return this; },
      setNotes(notes) { notes.forEach((row, i) => row.forEach((n, j) => { sheet.notes[`${r + i},${c + j}`] = n; })); return this; },
      setNumberFormat(format) { sheet.formats = [...(sheet.formats || []), { c, w, format }]; return this; },
      getNumberFormat() { const found = (sheet.formats || []).filter((f) => c >= f.c && c < f.c + f.w).pop(); return found ? found.format : "General"; },
      setFontColor(color) { sheet.fonts[key] = color; return this; },
      setBackground(color) { sheet.backgrounds.push({ r, c, h, w, color }); return this; },
      getDataValidations() { return Array.from({ length: h }, (_, i) => Array.from({ length: w }, (_, j) => (sheet.validations.find((v) => v.r <= r + i && v.c === c + j) || {}).rule || null)); },
      setDataValidations(rules) { sheet.validationSets = (sheet.validationSets || 0) + 1; return this; },
      setBackgrounds(colors) { colors.forEach((row, i) => row.forEach((color, j) => sheet.backgrounds.push({ r: r + i, c: c + j, h: 1, w: 1, color }))); return this; },
      getBackground() { const found = sheet.backgrounds.filter((b) => r >= b.r && r < b.r + b.h && c >= b.c && c < b.c + b.w).pop(); return found && found.color ? found.color : "#ffffff"; },
      getBackgrounds() { return Array.from({ length: h }, (_, i) => Array.from({ length: w }, (_, j) => sheet.getRange(r + i, c + j).getBackground())); },
      setDataValidation(rule) { sheet.validations.push({ r, c, h, w, rule }); sheet.validationWrites = (sheet.validationWrites || 0) + 1; return this; },
      clearDataValidations() { sheet.validations = sheet.validations.filter((v) => v.c < c || v.c >= c + w); return this; },
      // values, notes, font colours and backgrounds, cell by cell (a template row's formats)
      copyTo(dest) { for (let i = 0; i < h; i++) for (let j = 0; j < w; j++) { const target = sheet.getRange(dest.r + i, dest.c + j); target.setValue(sheet.value(r + i, c + j)); target.setNote(sheet.notes[`${r + i},${c + j}`] || ""); const font = sheet.fonts[`${r + i},${c + j}`]; if (font) target.setFontColor(font); target.setBackground(sheet.getRange(r + i, c + j).getBackground()); } return this; },
      clearContent() { for (let i = 0; i < h; i++) for (let j = 0; j < w; j++) if (sheet.cells[r - 1 + i]) sheet.cells[r - 1 + i][c - 1 + j] = ""; return this; },
      setFontWeight() { return this; }, setVerticalAlignment() { return this; }, setWrap() { return this; },
      setRichTextValue(v) { (sheet.rich = sheet.rich || {})[key] = v; return this.setValue(v.text); }, clearNote() { for (let i = 0; i < h; i++) for (let j = 0; j < w; j++) sheet.notes[`${r + i},${c + j}`] = ""; return this; },
      protect() { const p = new Protection(sheet, "RANGE"); p.range = this; sheet.protections.push(p); return p; },
    };
  }
}
class Book {
  constructor(id, name, sheets) { Object.assign(this, { id, name, order: [] }); Object.entries(sheets).forEach(([n, sheet]) => this.add(n, sheet)); }
  add(n, sheet, index = this.order.length) { sheet.name = n; sheet.parent = this; this.order.splice(index, 0, sheet); return sheet; }
  getName() { return this.name; }
  get sheets() { return Object.fromEntries(this.order.map((sheet) => [sheet.name, sheet])); }
  getId() { return this.id; } getSheets() { return this.order.slice(); } getSheetByName(n) { return this.order.find((sheet) => sheet.name === n) || null; }
  getSpreadsheetTimeZone() { return this.timeZone || "UTC"; }
  insertSheet(name, index) { return this.add(name, new Sheet(), index); }
  deleteSheet(sheet) { this.order = this.order.filter((other) => other !== sheet); }
}
const CATALOG_ID = "YOUR_PRODUCT_LIST_SPREADSHEET_ID";
const HEAD = [["2026/3/1"], ["Соронзон 2026"], ["Барааны нэрс"]];
const SHOP_HEAD = [["2026/3/1"], ["Соронзон 2026"], ["Барааны нэрс", "анхны үлдэгдэл", "өртөг үнэ", "зарах үнэ", "яг харагдах үнэ"]];
function catalogBook() {
  // Sheet100 / Sheet100*: names only (a 30 September copy); the current sheets: names and prices
  const names = (list) => new Sheet([...HEAD, ...list.map((x) => [x])]);
  return new Book(CATALOG_ID, "бараа", {
    "Sheet100": names(["Хөргөгч ", "Haier 300l", "Meiling  501l", "Sevenstar 11", "Panasonica 195l"]),
    "Sheet100*": names(["Haier 90-197", "Haier 300l", "xinix 240l"]),
    "Sor/2026/10/31": new Sheet([...HEAD, ["Haier 300l", 2, 900000, 1200000], ["Haier 400l", 1, 1100000, 1500000], ["Хөргөгч"]]),
    "My/2026/10/31": new Sheet([...HEAD, ["Haier 90-197", 1, 300000, 450000]]),
    // the first half of October: where lines of 1-15 October are written
    "Sor/2026/10/15": new Sheet([...HEAD, ["Meiling 501l"], ["Haier 400l"], ["Haier 300l"], ["xinix 240l"]]),
    "My/2026/10/15": new Sheet([...HEAD, ["Haier 300l"], ["Meiling 501l"]]),
  });
}

// ---------- the world: Drive, Gemini, properties ----------
const GEMINI = (result) => ({ code: 200, text: JSON.stringify({ candidates: [{ content: { parts: [{ text: "thinking", thought: true }, { text: JSON.stringify(result) }] } }] }) });
const item = (name, matched_name, qty, unit_price, confidence = 0.95) => ({ name, matched_name, qty, unit_price, confidence });
const receipt = (date, items, extra = {}) => ({ document_type: "shiljuuleg", document_title: "ЗАРЛАГЫН ПАДААН", date, items, issues: [], overall_confidence: 0.9, ...extra });

const FOLDERS = {
  "Соронзон": ["орлого", "зарлага", "Соронзоноос My home руу", "Соронзоноос Тосгон руу (буцаалт)", "тооллого"],
  "My home": ["орлого", "зарлага", "My home-оос Соронзон руу", "My home-оос Тосгон руу (буцаалт)", "тооллого"],
  "Тосгон": ["орлого", "Тосгоноос Соронзон руу", "Тосгоноос My home руу", "тооллого"],
};
const TOSGON_ID = "YOUR_TOSGON_BOOK_ID";
const BP = 68; // Тосгон's computed balance column
function tosgonBook() {
  // rows 4.. : Meiling 501l, Haier 300l (balance 5), Haier 400l, xinix 240l
  const tab = new Sheet([...HEAD, ["Meiling 501l"], ["Haier 300l"], ["Haier 400l"], ["xinix 240l "]]);
  tab.cells[4][BP - 1] = 5;
  return new Book(TOSGON_ID, "Тосгон 2026", { "Tos/2026/10": tab });
}
function makeWorld() {
  const root = new Folder("Баримт");
  Object.entries(FOLDERS).forEach(([place, names]) => names.forEach((name) => root.path(place, name)));
  root.path("Бусад");
  return { root, books: { [CATALOG_ID]: catalogBook(), [TOSGON_ID]: tosgonBook() }, docs: {}, created: {}, answers: {}, props: {}, triggers: [], geminiCalls: [], queries: [], thumbnails: [], today: "2026-10-17" };
}
const tos = (world) => world.books[TOSGON_ID].sheets["Tos/2026/10"];
const put = (world, path, name, bytes, created, answer, owner = "mom") => {
  const file = world.root.path(...path).add(new File(name, "image/heif", bytes, created, owner));
  if (answer) world.answers[bytes] = answer;
  return file;
};

// A workbook as a Drive file: its full copy is the check copy (ocr-check.gs).
const workbookFile = (world, id) => world.books[id] && {
  getId: () => id, getName: () => world.books[id].name, getParents: () => iter([]), isTrashed: () => Boolean(world.books[id].trashed),
  makeCopy: (name) => {
    const copyId = `${id}-check`;
    const sheets = Object.fromEntries(world.books[id].order.map((sheet) => [sheet.name, Object.assign(new Sheet(JSON.parse(JSON.stringify(sheet.cells))), { notes: { ...sheet.notes } })]));
    world.books[copyId] = new Book(copyId, name, sheets);
    return { getId: () => copyId };
  },
};
// A Blob of Utilities (newBlob, gzip, ungzip): its bytes as Apps Script gives them (signed).
const fakeBlob = (buffer) => ({ buffer, getBytes: () => [...buffer].map((b) => (b > 127 ? b - 256 : b)), getDataAsString: () => buffer.toString("utf8") });
// The background a cell has now (the last one set on a range holding it).
const bgOf = (sheet, r, c) => { const found = sheet.backgrounds.filter((b) => r >= b.r && r < b.r + b.h && c >= b.c && c < b.c + b.w).pop(); return found ? found.color : null; };
function run(world, dry, expr, source = code) {
  const logs = [];
  const ctx = {
    JSON, Date, Math, Set, Map, Number, String, Array, Object, Error,
    Logger: { log: (...a) => logs.push(a.join(" ")) },
    MimeType: { GOOGLE_SHEETS: "application/vnd.google-apps.spreadsheet" },
    Utilities: {
      DigestAlgorithm: { MD5: "MD5" },
      computeDigest: (alg, bytes) => [...crypto.createHash("md5").update(Buffer.from(bytes.map((b) => (b + 256) % 256))).digest()].map((b) => (b > 127 ? b - 256 : b)),
      base64Encode: (bytes) => Buffer.from(bytes.map((b) => (b + 256) % 256)).toString("base64"),
      base64Decode: (text) => [...Buffer.from(text, "base64")].map((b) => (b > 127 ? b - 256 : b)),
      newBlob: (data) => fakeBlob(Buffer.from(typeof data === "string" ? data : data.map((b) => (b + 256) % 256))),
      gzip: (blob) => fakeBlob(zlib.gzipSync(blob.buffer)),
      ungzip: (blob) => fakeBlob(zlib.gunzipSync(blob.buffer)),
      formatDate: (d, tz, fmt) => { if (fmt.includes("'T'")) return d.toISOString().slice(0, 19); if (tz === "Asia/Tokyo") return new Date(d.getTime() + 9 * 3600000).toISOString().slice(0, 10); const isNow = Math.abs(d - Date.now()) < 60000; const day = isNow ? world.today : d.toISOString().slice(0, 10); return fmt === "yyyy-MM-dd" ? day : `${day} 10:00`; },
    },
    UrlFetchApp: { fetch: (url, opts) => {
      if (url.startsWith("https://drive.google.com/thumbnail")) {
        world.thumbnails.push({ url, auth: opts.headers.Authorization });
        const ok = !world.noThumbnail;
        return { getResponseCode: () => (ok ? 200 : 404), getHeaders: () => ({ "Content-Type": ok ? "image/jpeg" : "text/html" }), getBlob: () => ({ jpeg: url }) };
      }
      if (url.includes("/externaldata/addimportrangepermissions")) {
        world.imports = [...(world.imports || []), url];
        return { getResponseCode: () => (world.importRefused ? 403 : 200), getContentText: () => "" };
      }
      const body = JSON.parse(opts.payload);
      if (!body.contents[0].parts[0].inline_data) { // a new product's category (ocr-new-product.gs)
        const ask = { prompt: body.contents[0].parts[0].text, schema: body.generationConfig.responseSchema, key: opts.headers["x-goog-api-key"] };
        world.categoryCalls = [...(world.categoryCalls || []), ask];
        const answer = typeof world.categoryAnswer === "function" ? world.categoryAnswer(ask) : world.categoryAnswer || {};
        const { code, text } = answer.code ? answer : GEMINI(answer);
        return { getResponseCode: () => code, getContentText: () => text };
      }
      const bytes = Buffer.from(body.contents[0].parts[0].inline_data.data, "base64").toString();
      world.geminiCalls.push({ url, bytes, key: opts.headers["x-goog-api-key"], prompt: body.contents[0].parts[1].text });
      const answer = world.answers[bytes];
      const { code, text } = answer && answer.code ? answer : GEMINI(answer || {});
      return { getResponseCode: () => code, getContentText: () => text };
    } },
    CacheService: { getScriptCache: () => ({ get: (k) => (world.cache = world.cache || {})[k] ?? null, put: (k, v, seconds) => { assert(typeof v === "string" && v.length <= 100 * 1024 && seconds > 0 && seconds <= 21600, `cache ${k}: a string under 100 KB, kept up to 6 hours`); world.cache[k] = v; }, remove: (k) => { delete (world.cache = world.cache || {})[k]; } }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => world.props[k] ?? null, setProperty: (k, v) => { assert(typeof v === "string" && v.length < 9000, `property ${k} too long`); world.props[k] = v; }, deleteProperty: (k) => { delete world.props[k]; }, getProperties: () => ({ ...world.props }) }) },
    Session: { getEffectiveUser: () => ({ getEmail: () => "owner@example.org" }) },
    DocumentApp: {
      ParagraphHeading: { TITLE: "TITLE", HEADING2: "HEADING2" },
      openById: (id) => world.docs[id],
      create: (name) => {
        const id = `doc${Object.keys(world.docs).length}`;
        const paragraph = (text, list) => ({ text, list, setHeading(h) { this.heading = h; return this; }, setLinkUrl(u) { this.url = u; return this; } });
        const doc = { id, name, paragraphs: [], saved: 0, getId: () => id, saveAndClose() { this.saved++; },
          getBody() { const d = this; return { clear() { d.paragraphs = []; }, appendParagraph(t) { const x = paragraph(t, false); d.paragraphs.push(x); return x; }, appendListItem(t) { const x = paragraph(t, true); d.paragraphs.push(x); return x; } }; } };
        world.docs[id] = doc;
        world.created = { ...world.created, [id]: Object.assign(new File(name, "application/vnd.google-apps.document", "", 0, "me"), { id }) };
        return doc;
      },
    },
    LockService: { getScriptLock: () => ({ tryLock: () => !world.locked, releaseLock() {} }) },
    ScriptApp: {
      getOAuthToken: () => "token-1",
      getProjectTriggers: () => world.triggers.slice(), deleteTrigger: (t) => { world.triggers = world.triggers.filter((x) => x !== t); },
      newTrigger: (handler) => { const b = { timeBased: () => b, everyHours: (h) => { b.hours = h; return b; }, after: (ms) => { b.after = ms; return b; }, forSpreadsheet: (book) => { b.book = book; return b; }, onEdit: () => b, create: () => world.triggers.push({ getHandlerFunction: () => handler, getTriggerSourceId: () => (b.book ? b.book.getId() : null), hours: b.hours, book: b.book }) }; return b; },
    },
    DriveApp: {
      getFolderById: () => world.root,
      getFileById: (id) => { const file = (world.created || {})[id] || world.root.allFiles().find((f) => f.id === id) || trashBin.find((f) => f.id === id) || workbookFile(world, id); if (!file) throw new Error(`no file ${id}`); return file; },
      searchFiles: (query) => { world.queries.push(query); return iter(world.root.allFiles().filter((f) => query.includes(`'${f.mime}'`))); },
    },
    SpreadsheetApp: {
      openById: (id) => world.books[id],
      create: (name) => {
        const id = `book${Object.keys(world.books).length}`;
        world.books[id] = new Book(id, name, { Sheet1: new Sheet() });
        world.created = { ...world.created, [id]: Object.assign(new File(name, "application/vnd.google-apps.spreadsheet", "", 0, "me"), { id }) };
        return world.books[id];
      },
      ProtectionType: { SHEET: "SHEET", RANGE: "RANGE" },
      newRichTextValue: () => { const v = { text: "", styles: [] }; const b = { setText: (t) => { v.text = t; return b; }, setTextStyle: (start, end, style) => { v.styles.push({ start, end, ...style }); return b; }, build: () => v }; return b; },
      newTextStyle: () => { const st = {}; const b = { setForegroundColor: (c) => { st.color = c; return b; }, setBold: (x) => { st.bold = x; return b; }, setFontSize: (n) => { st.size = n; return b; }, build: () => st }; return b; },
      newDataValidation: () => { const rule = {}; const b = { requireValueInList: (list) => { rule.list = list; return b; }, requireValueInRange: (range) => { rule.range = range; return b; }, requireDate: () => { rule.date = true; return b; }, setAllowInvalid: (x) => { rule.allowInvalid = x; return b; }, build: () => rule }; return b; },
      newConditionalFormatRule: () => { const rule = {}; const b = { whenFormulaSatisfied: (f) => { rule.formula = f; return b; }, setBackground: (c) => { rule.color = c; return b; }, setRanges: (r) => { rule.ranges = r; return b; }, build: () => rule }; return b; },
    },
  };
  vm.createContext(ctx);
  vm.runInContext(source.replace("const DRY_RUN = true;", `const DRY_RUN = ${dry};`).replace("const OCR_POST_FROM = '2026-10-20';", `const OCR_POST_FROM = '${world.postFrom || "2026-01-01"}';`).replace("const OCR_TOSGON_HAND_UNTIL = '2026-05-15';", `const OCR_TOSGON_HAND_UNTIL = '${world.handUntil || "2025-12-31"}';`), ctx);
  let value, error = null;
  try { value = vm.runInContext(expr, ctx); } catch (e) { error = e; }
  return { value, error, logs: logs.join("\n"), ctx };
}
const pure = run(makeWorld(), true, "0").ctx;
// Баримт/OCR/<month>/OCR <place> <month>, and its tab of a day
// The month spreadsheet of a place and kind (Баримт/OCR/2026-10/Орлого/OCR Соронзон орлого 2026-10);
// without a kind, the first kind's spreadsheet that has the day's tab.
const KINDS = [["Орлого", "орлого"], ["Зарлага", "зарлага"], ["Тооллого", "тооллого"], ["Шилжүүлэг", "шилжүүлэг"], ["Бусад", "бусад"]];
const kindBook = (world, place, month, [folder, word]) => { const dir = world.root.find("Зураг уншуулах", month, folder) || world.root.find("OCR", month, folder); const file = dir && dir.files.find((f) => f.name === `OCR ${place} ${word} ${month}`); return file ? { file, book: world.books[file.id] } : null; };
const monthBook = (world, place, iso, word) => {
  const kinds = word ? KINDS.filter(([, w]) => w === word) : KINDS;
  const found = kinds.map((kind) => kindBook(world, place, iso.slice(0, 7), kind)).filter(Boolean);
  return (iso.length > 7 && !word ? found.find((entry) => entry.book.getSheetByName(iso.slice(8))) : found[0]) || null;
};
const monthTab = (world, place, iso, word) => { const found = monthBook(world, place, iso, word); return found ? found.book.getSheetByName(iso.slice(8)) : null; };
// a workbook note's link to a line's row in its day tab (the owner, 2026-10-26)
const rowLink = (tab, row, id) => { const [file, line] = id.split("#"); return `https://docs.google.com/spreadsheets/d/${tab.getParent().getId()}/edit#gid=${tab.getSheetId()}&range=A${row}:N${row}&ocr=${file}${line ? `.${line}` : ""}`; };
// the rows of a day tab, up to its last receipt (the legend in column O beside the first rows is none)
const tabRows = (world, place, iso, word) => {
  const tab = monthTab(world, place, iso, word);
  if (!tab) {
    return null;
  }
  const isLine = (row) => (row || []).slice(0, header.length).some((v) => v !== "" && v !== undefined && v !== null);
  return tab.cells.slice(0, rowsOf(tab)).filter((row, index) => index === 0 || isLine(row)); // the blank rows under a photo are no lines
};
// how many rows a day tab has up to its last receipt, the header too
const rowsOf = (tab) => Math.max(1, tab.cells.map((row) => (row || []).slice(0, header.length).some((v) => v !== "" && v !== undefined && v !== null)).lastIndexOf(true) + 1);

// ---------- pure decisions ----------
eq(plain(pure.ocrFolderMeaning(["Соронзон", "зарлага"])), { label: "Соронзон / зарлага", kind: "zarlaga", moves: [{ place: "Соронзон", side: "out" }] }, "shop expense folder");
eq(pure.ocrFolderMeaning(["My home", "орлого"]).moves[0].side, "in", "shop income folder");
// one spreadsheet per place, kind and month (2026-10-25)
{
  const P = (expr) => plain(vm.runInContext(expr, pure));
  eq(P(`[["My home", "орлого"], ["My home", "зарлага"], ["Соронзон", "тооллого"], ["Тосгон", "Тосгоноос Соронзон руу"], ["My home", "My home-оос Тосгон руу (буцаалт)"], ["My home", "миний хавтас"]].map((path) => ocrBookKind(path).folder)`),
     ["Орлого", "Зарлага", "Тооллого", "Шилжүүлэг", "Шилжүүлэг", "Бусад"], "a folder's kind: a return is a transfer, a folder of mom's own is \"Бусад\"");
  eq(P(`[ocrMonthBookName("My home", OCR_BOOK_KINDS.zarlaga, "2026-10"), ocrMonthBookOf("OCR My home зарлага 2026-10"), ocrMonthBookOf("OCR My home 2026-10"), ocrMonthBookOf("OCR Дархан орлого 2026-10")]`),
     ["OCR My home зарлага 2026-10", { place: "My home", kind: { folder: "Зарлага", word: "зарлага" }, month: "2026-10" }, null, null], "the name says place, kind and month; the old name of every kind is not one");
  eq(P(`[ocrOldMonthBookOf("OCR My home 2026-10"), ocrOldMonthBookOf("OCR My home зарлага 2026-10"), ocrOldMonthBookOf("OCR Соронзон 2026-10 (хуучин)")]`),
     [{ place: "My home", month: "2026-10" }, null, null], "an old spreadsheet of every kind is told apart, and once put away it is not one any more");
}
eq(plain(pure.ocrFolderMeaning(["Тосгон", "Тосгоноос Соронзон руу"])), { label: "Тосгон / Тосгоноос Соронзон руу", kind: "shiljuuleg", moves: [{ place: "Тосгон", side: "out", other: "Соронзон" }, { place: "Соронзон", side: "in", other: "Тосгон" }] }, "a transfer sits in the giver's folder");
eq(plain(pure.ocrFolderMeaning(["My home", "My home-оос Тосгон руу (буцаалт)"])), { label: "My home / My home-оос Тосгон руу (буцаалт)", kind: "butsaalt", moves: [{ place: "My home", side: "out", other: "Тосгон" }, { place: "Тосгон", side: "in", other: "My home" }] }, "a return: My home gives back to Тосгон");
eq(pure.ocrFolderMeaning(["Соронзон", "Тосгоноос Соронзон руу"]).moves.length, 0, "a transfer folder under the wrong place is not trusted");
eq(plain(pure.ocrTransferPlaces("My home-оос Соронзон руу")), { from: "My home", to: "Соронзон" }, "My home-оос");
eq(plain(pure.ocrTransferPlaces("Соронзоноос Тосгон руу (буцаалт)")), { from: "Соронзон", to: "Тосгон" }, "Соронзоноос … (буцаалт)");
eq(pure.ocrFolderMeaning(["Бусад", "x"]).moves.length, 0, "unknown folder gives no direction");
// a Тосгон issue slip put in a shop's income folder: one slip, two sides (the owner, 2026-10-28):
// Тосгон's expense and the shop's income, as if it lay in "Тосгоноос <shop> руу"; no note. Until
// then (2026-10-23) it was noted and only the shop's income was written.
{
  const soIn = pure.ocrFolderMeaning(["Соронзон", "орлого", "2026-04", "02"]);
  const slip = (result, folder = soIn) => pure.ocrSlipFolder(folder, result);
  const reasonsOf = (result, folder = soIn) => plain(pure.ocrDocumentCheck({ items: [item("a", null, 1, 1)], ...result }, slip(result, folder), "2026-04-02", "2026-04-02")).reasons;
  const fromTosgon = { label: "Соронзон / орлого (Тосгоноос)", kind: "shiljuuleg", moves: [{ place: "Тосгон", side: "out", other: "Соронзон" }, { place: "Соронзон", side: "in", other: "Тосгон" }] };
  eq(plain(slip({ document_type: "zarlaga", header_sender: "Тос", header_receiver: "Соронзонд" })), fromTosgon,
     "a slip Тосгон sent, in Соронзон's income folder, is Тосгон's expense and Соронзон's income");
  eq(plain(pure.ocrFolderMeaning(["Соронзон", "орлого (Тосгоноос)"])), fromTosgon, "its rows' folder label says so, and means so again when read back");
  eq([pure.ocrFolderMeaning(["Тосгон", "орлого (Тосгоноос)"]).moves.length, pure.ocrFolderMeaning(["Соронзон", "зарлага (Тосгоноос)"]).moves.length], [0, 0], "only a shop's income");
  eq(reasonsOf({ document_type: "zarlaga", header_sender: "Тос", header_receiver: "Соронзонд" }), [], "and it is not noted as in the wrong folder");
  eq(reasonsOf({ document_type: "zarlaga", header_sender: null }), ["хавтас буруу байж магадгүй: илгээгч тодорхойгүй → Тосгоноос Соронзон?"],
     "an issue slip without a sender is still noted, in a few words: it may be a supplier's");
  eq(plain(slip({ document_type: "zarlaga", header_sender: null })), plain(soIn), "and stays the shop's income only");
  eq(reasonsOf({ document_type: "orlogo", header_sender: "Номин ХХК" }), [], "a supplier's receipt is fine");
  eq(plain(slip({ document_type: "zarlaga", header_sender: "Тос" }, pure.ocrFolderMeaning(["Соронзон", "зарлага"]))), plain(pure.ocrFolderMeaning(["Соронзон", "зарлага"])), "only an income folder");
  eq(plain(slip({ document_type: "zarlaga", header_sender: "Тос" }, pure.ocrFolderMeaning(["Тосгон", "орлого"]))), plain(pure.ocrFolderMeaning(["Тосгон", "орлого"])), "not Тосгон's own");
  eq([slip({ document_type: "shiljuuleg", header_sender: "ТОСГОН агуулах" }, pure.ocrFolderMeaning(["My home", "орлого"])).label, reasonsOf({ document_type: "shiljuuleg", header_sender: "ТОСГОН агуулах" }, pure.ocrFolderMeaning(["My home", "орлого"]))],
     ["My home / орлого (Тосгоноос)", []], "a transfer slip from Тосгон in My home's income folder: the same");
  // mom's own names for the places (2026-10-28): "Төв" on the organisation line is Тосгон
  eq(slip({ document_type: "zarlaga", header_sender: "Төв", header_receiver: "Ачит" }, pure.ocrFolderMeaning(["My home", "орлого"])).label, "My home / орлого (Тосгоноос)", "a slip whose header says Төв was sent by Тосгон");
  eq(slip({ document_type: "zarlaga", header_sender: "төв агуулах" }).label, "Соронзон / орлого (Тосгоноос)", "Төв as the first word, in any case");
  eq(slip({ document_type: "orlogo", header_sender: "Төвийн Номин ХХК" }).label, "Соронзон / орлого", "a supplier whose name only starts with the letters is not Тосгон");
  eq(slip({ document_type: "orlogo", header_sender: "Хүнсний төв" }).label, "Соронзон / орлого", "nor one with төв inside its name");
  // a row read before 2026-10-28 (noted, the shop's income only) takes the new meaning in the review
  const oldRow = { "Хавтас (ээж хийсэн)": "My home / орлого", "Шалгах шалтгаан": "хавтас буруу байж магадгүй: Тосгоноос My home руу хавтсанд байх ёстой юм шиг (илгээгч: Тосгон); үнэ байхгүй", "Төлөв": "Шалгах" };
  eq(plain(pure.ocrTosgonSentRecord(oldRow)), { ...oldRow, "Хавтас (ээж хийсэн)": "My home / орлого (Тосгоноос)", "Шалгах шалтгаан": "", "Төлөв": "OK" }, "an old noted row: the new folder label, the note and the missing price gone (a transfer needs none): OK");
  eq(plain(pure.ocrTosgonSentRecord({ ...oldRow, "Шалгах шалтгаан": `${oldRow["Шалгах шалтгаан"]}; тоо зөрсөн (Midea 142 ≠ Midea 143l)` })), { ...oldRow, "Хавтас (ээж хийсэн)": "My home / орлого (Тосгоноос)", "Шалгах шалтгаан": "тоо зөрсөн (Midea 142 ≠ Midea 143l)" }, "any other reason stays: still Шалгах");
  // the same slip read fresh in its transfer folder asks no price
  {
    const sent = { document_type: "shiljuuleg", header_sender: "Тосгон", header_receiver: "Ачит", items: [item("Haier 300л", "Haier 300l", 2, null)] };
    const folder = slip(sent, pure.ocrFolderMeaning(["My home", "орлого", "2026-10", "23"]));
    const rows = plain(pure.ocrRows(sent, { name: "IMG_5880.HEIC", id: "p" }, folder, pure.ocrDocumentCheck(sent, folder, "2026-10-23", "2026-10-23"), new Set(["Haier 300l"]), ""));
    const col = (title) => vm.runInContext(`OCR_HEADER.indexOf(${JSON.stringify(title)})`, pure);
    eq([rows[0][col("Төлөв")], rows[0][col("Шалгах шалтгаан")], rows[0][col("Хавтас (ээж хийсэн)")]], ["OK", "", "My home / орлого (Тосгоноос)"], "a Тосгон slip without prices in My home's income folder is OK (yellow)");
  }
  eq(pure.ocrTosgonSentRecord({ ...oldRow, "Шалгах шалтгаан": "хавтас буруу байж магадгүй: Тосгоноос My home руу хавтсанд байх ёстой юм шиг (илгээгч: Төв)" })["Төлөв"], "OK", "nothing left to check: OK");
  eq([pure.ocrTosgonSentRecord({ ...oldRow, "Шалгах шалтгаан": "хавтас буруу байж магадгүй: зарлагын баримт, илгээгч тодорхойгүй (Тосгоноос ирсэн бол Тосгоноос My home руу хавтсанд)" }), pure.ocrTosgonSentRecord({ ...oldRow, "Шалгах шалтгаан": "үнэ байхгүй" })],
     [null, null], "a row without a sender, or without the note, is left as it is");
  // the photo also lies on the other side: a copy in the other place's folder of the same day
  eq(plain([pure.ocrMirrorPaths(["Тосгон", "Тосгоноос Соронзон руу", "2026-10", "02"], null), pure.ocrMirrorPaths(["My home", "орлого", "2026-10", "23"], "Тосгон"), pure.ocrMirrorPaths(["Тосгон", "Тосгоноос My home руу", "2026-10"], null)]),
     [[["Соронзон", "орлого", "2026-10", "02"]], [["Тосгон", "Тосгоноос My home руу", "2026-10", "23"]], [["My home", "орлого", "2026-10"]]],
     "a Тосгон transfer photo goes into the shop's income folder of its day; a slip Тосгон sent, in a shop's income folder, into Тосгон's folder to that shop");
  eq(plain([pure.ocrMirrorPaths(["My home", "орлого", "2026-10", "23"], null), pure.ocrMirrorPaths(["Тосгон", "орлого", "2026-10", "23"], "Тосгон"), pure.ocrMirrorPaths(["My home", "My home-оос Тосгон руу (буцаалт)", "2026-10", "23"], "Тосгон"), pure.ocrMirrorPaths(["Тосгон", "Тосгоноос Соронзон руу"], null)]),
     [[], [], [], []], "nothing for a supplier's slip, Тосгон's own income, a return, or a photo outside a month folder");
  ok(pure.ocrPrompt([]).includes('"Төв", "Тос", "Тов", "Склад", "Склд", "Скл" = Тосгон, "Ачит", "Myhome" = My home'), "the prompt names mom's names for the places");
  // "Склд" (склад, the warehouse) is Тосгон too (the owner, 2026-10-29: read as "Солонгос" until then)
  eq(["Склд", "Скл", "склад", "Склд агуулах"].map((sender) => slip({ document_type: "zarlaga", header_sender: sender }, pure.ocrFolderMeaning(["My home", "орлого"])).label),
     Array(4).fill("My home / орлого (Тосгоноос)"), "a slip whose header says Склд was sent by Тосгон");
  eq(slip({ document_type: "orlogo", header_sender: "Скляр ХХК" }).label, "Соронзон / орлого", "a supplier whose name only starts with the letters is not");
  ok(pure.ocrPrompt([]).includes('"Байгууллагын нэр" нь илгээгч'), "and which header line of the issue slip is the sender");
  ok(pure.ocrPrompt([]).includes("Тосгоны зарлага, тэр дэлгүүрийн орлого хоёулаа"), "and that a slip Тосгон issued to a shop is both, so its sender and receiver are always written (the owner, 2026-10-28)");
  eq(plain(pure.ocrCalendarDay([
    { "Хавтас (ээж хийсэн)": "Соронзон / орлого", "ID": "p1#1", "Бараа": "Haier 300l", "Шалгах шалтгаан": "хавтас буруу байж магадгүй: ...", "Төлөв": "Шалгах" },
    { "Хавтас (ээж хийсэн)": "Соронзон / орлого", "ID": "p1#2", "Бараа": "Haier 400l", "Шалгах шалтгаан": "хавтас буруу байж магадгүй: ...", "Төлөв": "Шалгах" },
    { "Хавтас (ээж хийсэн)": "Соронзон / орлого", "ID": "p2#1", "Бараа": "Haier 400l", "Шалгах шалтгаан": "", "Төлөв": "OK" },
  ]))["орлого"][6], 1, "the calendar's 7th number: the photos that may be in the wrong folder (one photo, not its lines)");
}
eq(plain(pure.ocrFolderMeaning(["Тосгон", "тооллого", "2026-10", "31"])), { label: "Тосгон / тооллого", kind: "toollogo", moves: [{ place: "Тосгон", side: "count" }] }, "a count folder");
// a slip from the city ("Хотоос" on the organisation line) in a shop's income folder (the owner,
// 2026-10-28): the goods came from the city to Тосгон and on to the shop, so the one slip is Тосгон's
// income, Тосгон's expense and the shop's income; its photo is kept in all three places' folders
{
  const soIn = pure.ocrFolderMeaning(["Соронзон", "орлого", "2026-10", "24"]);
  const city = { document_type: "zarlaga", header_sender: "Хотоос", header_receiver: "Соронзон" };
  const fromCity = { label: "Соронзон / орлого (Хотоос)", kind: "shiljuuleg", moves: [{ place: "Тосгон", side: "in" }, { place: "Тосгон", side: "out", other: "Соронзон" }, { place: "Соронзон", side: "in", other: "Тосгон" }] };
  eq(plain(pure.ocrSlipFolder(soIn, city)), fromCity, "a slip from the city in Соронзон's income folder is Тосгон's income, Тосгон's expense and Соронзон's income");
  eq(plain(pure.ocrFolderMeaning(["Соронзон", "орлого (Хотоос)"])), fromCity, "its rows' folder label says so, and means so again when read back");
  eq([pure.ocrSlipFolder(pure.ocrFolderMeaning(["My home", "орлого"]), { document_type: "zarlaga", header_sender: "Хот" }).label, pure.ocrSlipFolder(soIn, { document_type: "zarlaga", header_sender: "хотын агуулах" }).label],
     ["My home / орлого (Хотоос)", "Соронзон / орлого (Хотоос)"], "My home the same; Хот in any form");
  eq([pure.ocrFolderMeaning(["Тосгон", "орлого (Хотоос)"]).moves.length, pure.ocrSlipFolder(pure.ocrFolderMeaning(["Тосгон", "орлого"]), city).label, pure.ocrSlipFolder(pure.ocrFolderMeaning(["Соронзон", "зарлага"]), city).label],
     [0, "Тосгон / орлого", "Соронзон / зарлага"], "only a shop's income folder: Тосгон's own income from the city stays its income");
  eq(plain(pure.ocrDocumentCheck({ items: [item("a", null, 1, null)], ...city }, pure.ocrSlipFolder(soIn, city), "2026-10-24", "2026-10-24")).reasons, [], "not noted as in the wrong folder");
  eq([pure.ocrSlipSender(city), pure.ocrSlipSender({ header_sender: "Тос" }), pure.ocrSlipSender({ header_sender: "Номин ХХК" }), pure.ocrSlipSender(null)], ["Хот", "Тосгон", null, null], "who sent a slip, as the rule needs it");
  eq(plain(pure.ocrMirrorPaths(["Соронзон", "орлого", "2026-10", "24"], "Хот")), [["Тосгон", "Тосгоноос Соронзон руу", "2026-10", "24"], ["Тосгон", "орлого", "2026-10", "24"]],
     "its photo is copied into Тосгон's folder to the shop and Тосгон's income folder of the day: three places in all");
  ok(pure.ocrPrompt([]).includes("Тосгоны орлого, Тосгоны зарлага, тэр дэлгүүрийн орлого гурвуулаа"), "the prompt says a slip from the city to a shop is all three, so its sender is written as Хот");
}


const place = (p, side, iso) => pure.ocrSheetPlace({ place: p, side }, iso);
eq(place("Соронзон", "in", "2026-10-01"), "Sor/2026/10/15 · орлого · E", "first half, income, day 1");
eq(place("My home", "out", "2026-09-16"), "My/2026/9/30 · зарлага · U", "second half starts again at the first day column");
eq(place("Соронзон", "out", "2026-10-31"), "Sor/2026/10/31 · зарлага · AI", "day 31 goes into the last day column");
eq(place("Тосгон", "in", "2026-10-01"), "Tos/2026/10 · орлого · E", "Тосгон month sheet");
eq(place("Тосгон", "out", "2026-10-31"), "Tos/2026/10 · зарлага · BM", "Тосгон day 31 = its 30 column");
eq(place("Соронзон", "in", "2025-12-01"), "", "outside the workbook's year: no place");
{ // the registry of the years' workbooks (year-books.gs): a receipt goes into the file of its own year
  const years = "[{ year: 2026, main: 'main26', tosgon: 'tos26', from: '2026-01-01' }, { year: 2027, main: 'main27', tosgon: 'tos27', from: '2027-01-01' }]";
  const y = run(makeWorld(), true, "0", code.replace("function readRegistry() {", `function readRegistry() { return ${years}; }\nfunction readRegistryOfDrive() {`)).ctx;
  ok(y.ocrSheetPlace({ place: "Соронзон", side: "in" }, "2027-01-02").startsWith("Sor/2027/1/15 · орлого"), "a 2027 receipt goes into the 2027 sheet");
  eq([y.ocrBookIdFor("Соронзон", "Sor/2026/12/31"), y.ocrBookIdFor("Соронзон", "Sor/2027/1/15"), y.ocrBookIdFor("Тосгон", "Tos/2027/1")], ["main26", "main27", "tos27"], "into that year's workbook");
  ok(y.ocrSheetPlace({ place: "Соронзон", side: "in" }, "2026-10-01").startsWith("Sor/2026/10/15 · орлого"), "and a 2026 receipt still into 2026's");
}
eq(place("Соронзон", "count", "2026-10-31"), "Sor/2026/10/31 · тооллого · AK", "a shop count goes to the count column of that half-month");
eq(place("My home", "count", "2026-10-15"), "My/2026/10/15 · тооллого · AK", "also at the end of the first half");
eq(place("My home", "count", "2026-10-03"), "шалгах тооллого (workbook-д бичихгүй)", "a count on another day is a check, never written");
eq([place("Тосгон", "count", "2026-09-15"), place("Тосгон", "count", "2026-02-28")], ["шалгах тооллого (workbook-д бичихгүй)", "Tos/2026/2 · тооллого · BO"], "Тосгон closes its month only");
eq([place("Тосгон", "count", "2026-10-15"), place("Соронзон", "count", "2026-10-23"), place("Соронзон", "count", "2026-10-10")], ["Tos/2026/10 · Тооллого 10/15", "Sor/2026/10/31 · Тооллого 10/23", "шалгах тооллого (workbook-д бичихгүй)"], "from the 10/31 sheets on a count between the ends goes into its day's block");
eq(place("Тосгон", "count", "2026-10-31"), "Tos/2026/10 · тооллого · BO", "Тосгон's count column");
const moved = (folderName, iso) => pure.ocrFolderMeaning(folderName.split(" / ")).moves.map((move) => pure.ocrSheetPlace(move, iso));
eq(plain(moved("Соронзон / Соронзоноос My home руу", "2026-10-03")), ["Sor/2026/10/15 · шилжүүлэг · AJ", "My/2026/10/15 · шилжүүлэг · T"], "between the two shops: the transfer columns, not sales");
eq(plain(moved("My home / My home-оос Тосгон руу (буцаалт)", "2026-10-20")), ["My/2026/10/31 · шилжүүлэг · AJ", "Tos/2026/10 · орлого · X"], "a return: out of the shop's transfer column (not a sale), into Тосгон's day");
eq(plain(moved("Тосгон / Тосгоноос My home руу", "2026-10-20")), ["Tos/2026/10 · зарлага · BC", "My/2026/10/31 · орлого · I"], "from Тосгон: the shop's income of that day");

ok(!pure.ocrBrandMatches("Aucma 195", "Panasonica 195l"), "a different brand is refused");
ok(pure.ocrBrandMatches("SESTAR 11", "Sevenstar 11"), "a misspelt brand is fine");
ok(pure.ocrBrandMatches("XNX 240", "xinix 240l"), "short misspelling");
ok(pure.ocrBrandMatches("Cmedia 418l", "Cmidea 418l"), "swapped letters");
ok(pure.ocrBrandMatches("441l", "Aucma 441L"), "only a number: the number decides");

const catalog = new Set(["Haier 300l", "Sevenstar 11"]);
eq(plain(pure.ocrItemCheck(item("Haier 300", "Haier 300l", 1, 430), catalog, true).reasons), [], "a good line has no reason");
eq(plain(pure.ocrItemCheck(item("Haier 300", "Haier 300", 1, 430), catalog, true)), { matched: "Хамааралгүй", reasons: ["нэр жагсаалтад таараагүй"] }, "a name not exactly in the list becomes Хамааралгүй, for mom to pick");
eq(plain(pure.ocrItemCheck(item("Haier 300", "Хамааралгүй", 1, 430), catalog, true)), { matched: "Хамааралгүй", reasons: ["жагсаалтад тохирох бараа алга"] }, "Gemini found no product: Хамааралгүй");
eq(plain(pure.ocrItemCheck(item("Haier 300", "Haier 300l", null, null, 0.45), catalog, true)), { matched: "Haier 300l", reasons: ["тоо байхгүй", "үнэ байхгүй", "итгэл бага (0.45)"] }, "count, price, and a Gemini confidence below 0.7 (the owner, 2026-10-25; none 10/23-10/25)");
eq([0.7, 0.69, undefined].map((confidence) => plain(pure.ocrItemCheck({ ...item("Haier 300", "Haier 300l", 1, 430), confidence }, catalog, true).reasons)), [[], ["итгэл бага (0.69)"], []], "0.7 itself is sure enough; a line without a confidence is not held for it");
// The owner's criteria (2026-10-22), measured on 220 lines: Gemini's own confidence is 0.85-0.95 for
// nearly every line, also for "Aucma 200 l" taken as "Aucma 205l". A line goes in when the numbers of
// the name match, the receipt price is within ±10% of the list price (±50% until 2026-10-22) and the
// brand matches; otherwise it waits for mom with the reason, its product still filled in (2026-10-25:
// always one product, mom corrects a wrong one).
{
  const wide = new Set(["Haier 300l", "Aucma 205l", "Haier 90-197", "Smart 7.2l", "Haier 100l", "Hisense 200l", "Ronshin 1.200"]);
  const prices = new Map([["Haier 300l", 1030000]]); // the price the customer sees
  const check = (it) => plain(pure.ocrItemCheck(it, wide, true, prices));
  const low = check(item("Haier 300", "Haier 300l", 1, 1000, 0.45));
  eq([low.matched, low.reasons], ["Haier 300l", ["итгэл бага (0.45)"]], "a low confidence makes it a Шалгах line; the pick stays (the owner, 2026-10-25)");
  const size = check(item("Aucma 200 l", "Aucma 205l", 1, 500));
  eq([size.matched, size.reasons], ["Aucma 205l", ["тоо зөрсөн (Aucma 200 l ≠ Aucma 205l)"]], "a size or model number that differs: the pick stays, mom checks it");
  eq(check(item("Hisense 201", "Hisense 200l", 1, 500)).reasons, ["тоо зөрсөн (Hisense 201 ≠ Hisense 200l)"], "201 is not 200");
  eq([check(item("Haier 9.0-197", "Haier 90-197", 1, 400)).matched, check(item("Smart 72", "Smart 7.2l", 1, 300)).matched, check(item("Ronshin 1200", "Ronshin 1.200", 1, 2400)).matched],
     ["Haier 90-197", "Smart 7.2l", "Ronshin 1.200"], "9.0 is 90, 72 is 7.2, 1200 is 1.200: only the dots and commas differ");
  eq(check(item("Haier 100-628", "Haier 100l", 1, 480)).matched, "Haier 100l", "a model code the receipt adds is fine: the list name's numbers are all there");
  const cheap = check(item("Haier 300", "Haier 300l", 1, 430));
  eq([cheap.matched, cheap.reasons], ["Haier 300l", ["үнэ зөрсөн (баримт 430,000₮, жагсаалт 1,030,000₮)"]], "a receipt price more than 10% off the price the customer sees: mom checks");
  eq(check(item("Haier 300", "Haier 300l", 1, 950)).reasons, [], "within ±10% it goes in (a sale price is a little lower)");
  eq(check(item("Haier 300", "Haier 300l", 1, 900)).reasons.length, 1, "12% off is too far (2026-10-22: ±10%)");
  eq(check(item("Haier 300", "Haier 300l", 1, null)).reasons, ["үнэ байхгүй"], "no receipt price: no price check, as before");
}
eq(plain(pure.ocrItemCheck(item("Haier 300", "Haier 300l", 1, null), catalog, false).reasons), [], "a transfer or a return needs no price");

const folderZ = pure.ocrFolderMeaning(["Соронзон", "зарлага", "2026-10", "01"]);
eq(folderZ.label, "Соронзон / зарлага", "the label is the place and the folder, not the day");
eq(pure.ocrFolderDate(["Соронзон", "зарлага", "2026-10", "01"]), "2026-10-01", "the day folder gives the date");
eq([pure.ocrFolderDate(["Соронзон", "зарлага"]), pure.ocrFolderDate(["Соронзон", "зарлага", "2026-10"]), pure.ocrFolderDate(["Соронзон", "зарлага", "2026-02", "30"])], [null, null, null], "no date outside a real day folder");
eq(plain(pure.ocrUpcomingMonths("2026-12-05")), ["2026-12", "2027-01"], "this month and the next, also over the year end");
eq(pure.ocrMonthDays("2026-02").length, 28, "a month gets as many day folders as it has days");
const chk = (result, folderDate, folder = folderZ) => plain(pure.ocrDocumentCheck(result, folder, folderDate, "2026-10-15"));
const zarlaga = (date) => receipt(date, [item("a", null, 1, 1)], { document_type: "zarlaga" });
eq(chk(zarlaga("2026-10-09"), "2026-10-01"), { date: "2026-10-01", reasons: [] }, "the day is the folder mom chose; the date on the receipt is not used or compared");
eq(chk(zarlaga(null), "2026-10-01"), { date: "2026-10-01", reasons: [] }, "a receipt without a date is no problem in a day folder");
eq(chk(zarlaga("2026-10-03"), null), { date: "2026-10-03", reasons: ["өдрийн хавтсанд хийгээгүй, баримт дээрх огноогоор (2026-10-03)"] }, "outside a day folder: the date written on the receipt (the owner, 2026-10-28), marked");
eq(chk(zarlaga(null), null), { date: "2026-10-15", reasons: ["өдрийн хавтсанд хийгээгүй, оруулсан өдрөөр (2026-10-15)"] }, "outside a day folder, no date on the receipt: the upload day, marked");
eq([chk(zarlaga("2020-10-03"), null).date, chk(zarlaga("2026-09-03"), null).date, chk(zarlaga("2026-10-32"), null).date, chk(zarlaga("18"), null).date], ["2026-10-03", "2026-10-15", "2026-10-15", "2026-10-15"],
   "the year on the slip is not trusted (\"20__ он\"); another month, an impossible day or a bare number falls back to the upload day");
eq(plain(pure.ocrDocumentCheck(zarlaga("2026-09-03"), folderZ, null, "2026-10-15", "2026-09")).date, "2026-09-03", "a loose photo in last month's folder: the receipt's day in that month");
eq([pure.ocrFolderMonth(["Соронзон", "зарлага", "2026-10"]), pure.ocrFolderMonth(["Соронзон", "зарлага", "2026-10", "01"]), pure.ocrFolderMonth(["Соронзон", "зарлага"])], ["2026-10", "2026-10", null], "the month folder of a path");
ok(chk(receipt("2026-10-01", [item("a", null, 1, 1)], { document_type: "other" }), "2026-10-01").reasons.includes("маягтгүй хуудас, гараар шалгана"), "a notebook page is marked");
ok(chk(receipt("2026-10-01", []), "2026-10-01").reasons.includes("хоосон баримт"), "a blank receipt is marked");
ok(chk(receipt("2026-10-01", [item("a", null, 1, 1)], { document_type: "butsaalt" }), "2026-10-01").reasons.some((r) => r.includes("шилжүүлэг/буцаалтын баримт мэт")), "a return slip in a sale folder is marked");
eq(chk(receipt("2026-10-01", [item("a", null, 1, 1)], { document_type: "butsaalt" }), "2026-10-01", pure.ocrFolderMeaning(["My home", "My home-оос Тосгон руу (буцаалт)", "2026-10", "01"])).reasons, [], "a return slip in the return folder is fine");
const countFolder = pure.ocrFolderMeaning(["Соронзон", "тооллого", "2026-10", "31"]);
eq(chk(receipt("2026-10-31", [item("a", null, 3, null)], { document_type: "toollogo" }), "2026-10-31", countFolder).reasons, [], "a count sheet in the count folder is fine");
eq([chk(zarlaga("2026-10-31"), "2026-10-31", countFolder).reasons, chk(receipt("2026-10-31", [item("a", null, 3, null)], { document_type: "other" }), "2026-10-31", countFolder).reasons], [[], []],
   "in the count folder the folder says it is a count, whatever Gemini calls the page (2026-10-25: Тосгон's 10/23 pages came back as \"other\")");
ok(chk(receipt("2026-10-31", [item("a", null, 3, null)], { document_type: "toollogo" }), "2026-10-31").reasons.some((r) => r.startsWith("тооллогын хуудас мэт")), "a count sheet in a sale folder is marked");

const header = plain(vm.runInContext("OCR_HEADER", pure));
const rows = pure.ocrRows(receipt("2026-10-01", [item("Haier 300", "Haier 300l", 1, 430), item("??", null, 1, 5, 0.5)]), { name: "IMG_1", url: "u", id: "IMG_1-id" },
  pure.ocrFolderMeaning(["Тосгон", "Тосгоноос Соронзон руу"]), { date: "2026-10-01", reasons: [] }, catalog, "2026-10-17 10:00");
ok(rows.every((row) => row.length === header.length), "every row has every column");
eq(rows.map((row) => row[header.indexOf("ID")]), ["IMG_1-id#1", "IMG_1-id#2"], "each row has the photo's id and its line number");
ok(rows.every((row) => ["Харах", "Ээж: Зөв/Буруу", "Бичсэн"].every((title) => row[header.indexOf(title)] === "")), "the checks and the written cell start empty");
eq([rows[0][header.indexOf("Тоо")], rows[0][header.indexOf("Тоо (уншсан)")]], [1, 1], "\"Тоо\" starts as the count read, which is also kept");
eq(["Бараа (уншсан)", "Нэгж үнэ", "Огноо", "Холбоос", "Гарчиг", "Төрөл", "Уншсан цаг"].filter((title) => header.includes(title)), [], "what was read as text, the receipt's price, date, link, title, kind, confidence and read time are not kept (the owner, 2026-10-25)");
eq(header.slice(0, 4), ["Харах", "Бараа", "Тоо", "Ээж: Зөв/Буруу"], "mom's columns first, the script's after them");
eq(["Ээж бичсэн", "Зарах үнэ / Тосгон өртөг", "Хаана орох"].filter((title) => header.includes(title)), [], "what mom typed in the workbook, the price the customer sees and where a line goes are gone too (the owner, 2026-10-25: nothing reads them)");
eq(rows.map((row) => [row[header.indexOf("Бараа")], row[header.indexOf("Бараа (тааруулсан)")]]), [["Haier 300l", "Haier 300l"], ["Хамааралгүй", "Хамааралгүй"]], "\"Бараа\" starts as the product Gemini picked, which is also kept; never empty");
eq([rows[0][header.indexOf("Төлөв")], rows[1][header.indexOf("Төлөв")]], ["OK", "Шалгах"], "status per line");
// A notebook page of goods bought in China (2026-10-23, Тосгон/орлого/2026-05/03/Document.HEIC):
// "15x426" is 15 pieces at 426 yuan, a rough cost price, not 426,000₮.
{
  const yuan = pure.ocrRows(receipt("2026-05-03", [item("Пл 20л", "Haier 300l", 15, 426)], { document_type: "other", price_currency: "CNY" }), { name: "Document.HEIC", url: "u", id: "y-id" },
    pure.ocrFolderMeaning(["Тосгон", "орлого"]), { date: "2026-05-03", reasons: [] }, catalog, "2026-10-23 10:00", [{ name: "Haier 300l", price: 1030000 }]);
  eq(yuan[0][header.indexOf("Тоо")], 15, "\"15x426\" from China: 15 pieces");
  ok(!String(yuan[0][header.indexOf("Шалгах шалтгаан")]).includes("үнэ зөрсөн"), "a yuan cost price is not compared with the price the customer sees");
  const tugrik = pure.ocrRows(receipt("2026-05-03", [item("Haier 300", "Haier 300l", 1, 430)], { price_currency: "MNT" }), { name: "IMG_T", url: "u", id: "t-id" },
    pure.ocrFolderMeaning(["Тосгон", "орлого"]), { date: "2026-05-03", reasons: [] }, catalog, "2026-10-23 10:00", [{ name: "Haier 300l", price: 1030000 }]);
  ok(String(tugrik[0][header.indexOf("Шалгах шалтгаан")]).includes("үнэ зөрсөн (баримт 430,000₮"), "a tögrög receipt is in thousands as before");
}

eq(pure.ocrParseResponse(200, GEMINI({ items: [] }).text).ok, true, "the answer after the thinking part is read");
eq(pure.ocrParseResponse(401, "bad key").fatal, true, "a wrong key stops the run");
eq(pure.ocrParseResponse(429, "busy").fatal, false, "a busy Gemini is not fatal");
eq(pure.ocrParseResponse(200, "<html>").ok, false, "garbage is a failed read");
eq(pure.ocrMonthFolderName("2026-10-01"), "2026-10", "month folder name");

// ---------- the name list ----------
eq(plain(pure.ocrCatalogTabs("2026-10-17")), ["Бараа", "My/2026/10/31", "Sor/2026/10/31", "Sheet100", "Sheet100*"], "the product list first, then the current half-month sheets, then the old name lists");
eq(plain(pure.ocrCatalogTabs("2026-10-03")).slice(1, 3), ["My/2026/10/15", "Sor/2026/10/15"], "first half of the month");
eq(plain(pure.ocrCatalogTabs("2025-12-20")), ["Бараа", "Sheet100", "Sheet100*"], "outside the workbook's year: only the name lists");
{
  const withList = makeWorld();
  const list = new Sheet([...HEAD, ["Угаалгын машин", ""], ["Haier 300L", "Хоёулаа"], ["Meiling 501 l", "Соронзон"], ["Зөвхөн жагсаалтад", "My home"]]);
  withList.books[CATALOG_ID].add("Бараа", list, 0);
  const fromList = run(withList, true, "loadOcrCatalog('2026-10-17')").value;
  ok(fromList.lines.includes("Haier 300L | 1200000₮") && !fromList.names.includes("Haier 300l"), "the list's spelling wins over the shop sheets; the price still comes from them");
  ok(fromList.names.includes("Зөвхөн жагсаалтад") && !fromList.names.includes("Угаалгын машин"), "every product of the list is known; its category headings are not products");
  ok(fromList.names.includes("Meiling 501 l") && !fromList.names.includes("Meiling 501l"), "a spelling of another tab is not added next to the list's");
  ok(!fromList.set.has("Haier 400l") && !fromList.set.has("Sevenstar 11"), "with a list, a name only a sheet holds is not offered: names come from the list alone (the owner, 2026-10-25)");
  const tosTab = withList.books[TOSGON_ID].sheets["Tos/2026/10"];
  tosTab.cells.push(["Kamin 2200", 1, 1360000, 1900000, 7]);
  ok(!run(withList, true, "loadOcrCatalog('2026-10-17')").value.set.has("Kamin 2200"), "nor one only Тосгон's sheet holds");
}
{ // the list now shows one shop's stock, cost, price and shelf price in C..F (a formula view): D is the cost (2026-10-22)
  const withView = makeWorld();
  const list = new Sheet([...HEAD, ["Хөргөгч", ""], ["Haier 300l", "Хоёулаа", 2, 900000, 1200000, 1250000], ["Haier 400l", "Соронзон", 1, 1100000, 1500000, 1600000]]);
  withView.books[CATALOG_ID].add("Бараа", list, 0);
  const viewed = run(withView, true, "loadOcrCatalog('2026-10-17')").value;
  ok(viewed.lines.includes("Haier 300l | 1200000₮") && viewed.lines.includes("Haier 400l | 1500000₮") && !viewed.lines.some((line) => line.includes("900000₮")),
     "the list's own columns are not the price: the selling price comes from the shop sheets");
  ok(!viewed.names.some((name) => name.trim().toLowerCase() === "хөргөгч"), "a category heading of a shop sheet or Sheet100 is not a product either");
}
{ // the price sent is the one the customer sees: the shops' E; without one, D as it is (2026-10-22, no more +3%)
  const shelf = makeWorld();
  // the shops' layout: E "яг харагдах үнэ", the price the customer sees
  shelf.books[CATALOG_ID].add("Sor/2026/10/31", new Sheet([...SHOP_HEAD, ["Haier 300l", 2, 900000, 1165000, 1210000], ["Haier 400l", 1, 1100000, 1456000, ""]]), 0);
  shelf.books[CATALOG_ID].order = shelf.books[CATALOG_ID].order.filter((sheet, i) => i === 0 || sheet.name !== "Sor/2026/10/31");
  const shown = run(shelf, true, "loadOcrCatalog('2026-10-17')").value;
  ok(shown.lines.includes("Haier 300l | 1210000₮"), "E of the shop sheet, not D");
  ok(shown.lines.includes("Haier 400l | 1456000₮"), "a row with E empty: D as it is");
  const tosOnly = makeWorld();
  const tosTab = tosOnly.books[TOSGON_ID].sheets["Tos/2026/10"];
  tosTab.cells.push(["Kamin 2200", 1, 1360000, 1900000, 7]);
  ok(run(tosOnly, true, "loadOcrCatalog('2026-10-17')").value.lines.includes("Kamin 2200 | 1900000₮"), "Тосгон's E is a day: its D");
}
eq(plain(pure.ocrCatalogEntries([["Haier 300l", 1, 2, 1200000], ["Meiling  501l ", "", "", ""], ["Haier 300l"], ["", "", "", 5], ["Meiling 501l", "", "", 900000]])),
  [{ name: "Haier 300l", price: 1200000 }, { name: "Meiling 501l", price: 900000 }], "each name once (cleaned), the first price found");
const loaded = run(makeWorld(), true, "loadOcrCatalog('2026-10-17')").value;
ok(loaded.set.has("Haier 400l") && !loaded.names.includes(""), "a product of this month's sheet is known though Sheet100 does not have it yet");
ok(loaded.lines.includes("Haier 300l | 1200000₮") && loaded.lines.includes("Sevenstar 11"), "names go to Gemini with the selling price where there is one");
const missingTab = makeWorld(); missingTab.books[CATALOG_ID].order = missingTab.books[CATALOG_ID].order.filter((sheet) => sheet.name !== "My/2026/10/31");
ok(/"My\/2026\/10\/31" sheet олдсонгүй/.test(run(missingTab, true, "loadOcrCatalog('2026-10-17')").logs), "a missing sheet is skipped with a note");
const body = pure.ocrRequestBody("AAAA", "image/heic", pure.ocrPrompt(["Haier 300l | 1200000₮"]), ["Haier 300l", "Sevenstar 11", "Haier 300l"]);
const itemSchema = body.generationConfig.responseSchema.properties.items.items;
eq([plain(itemSchema.properties.matched_name.enum), itemSchema.properties.matched_name.nullable, itemSchema.required.includes("matched_name")], [["Haier 300l", "Sevenstar 11", "Хамааралгүй"], undefined, true],
   "Gemini must pick one name of the list or Хамааралгүй for every line, never none (the owner, 2026-10-23)");
ok(/ЗААВАЛ ганц нэр сонго/.test(body.contents[0].parts[1].text) && /Panu\/Pana\/Пану = Panasonic/.test(body.contents[0].parts[1].text) && /"8\.0" = 80l/.test(body.contents[0].parts[1].text),
   "the prompt says so, with the writers' misspellings and abbreviations");
eq([pure.ocrParseResponse(400, "Invalid JSON payload: response_schema enum too long").fatal, pure.ocrParseResponse(400, "Unable to process input image").fatal], [true, false],
   "a refused schema stops the run (every photo would fail), a bad photo does not");
ok(itemSchema.required.includes("qty") && !itemSchema.properties.qty.nullable && /"qty".*ЗААВАЛ/.test(body.contents[0].parts[1].text), "the count is always asked for: never left empty");
ok(body.generationConfig.responseSchema.properties.price_currency.enum.join() === "MNT,CNY" && /15x426/.test(body.contents[0].parts[1].text) && /юань/.test(body.contents[0].parts[1].text),
  "Gemini is told how \"15x426\" reads and says whether the prices are tögrög or yuan");
ok(body.generationConfig.thinkingConfig.thinkingBudget === 1024 && body.contents[0].parts[1].text.endsWith("Haier 300l | 1200000₮"), "limited thinking and the list with prices in the prompt");

// ---------- a whole run ----------
const world = makeWorld();
const DAY = (place, folder, day) => [place, folder, "2026-10", day];
const TO_SO_1 = DAY("Тосгон", "Тосгоноос Соронзон руу", "01"), SO_IN_2 = DAY("Соронзон", "орлого", "02"), MY_OUT_1 = DAY("My home", "зарлага", "01");
world.props.GEMINI_KEY = "test-key";

const dry = run(world, true, "processNewReceipts()");
eq(dry.error, null, "the dry run works");
ok(/Соронзон \/ орлого \/ 2026-10: өдрийн хавтсууд үүснэ/.test(dry.logs) && /сар бэлдсэн 28/.test(dry.logs), "it lists the month folders it would make (14 folders x 2 months)");
eq(world.root.find("Соронзон", "орлого", "2026-10"), null, "DRY_RUN makes no folder");

run(world, false, "processNewReceipts()");
const octIn = world.root.find("Соронзон", "орлого", "2026-10");
ok(octIn && octIn.folders.length === 31 && world.root.find("Тосгон", "Тосгоноос My home руу", "2026-11").folders.length === 30, "this and next month get a folder for every day, in every folder");
ok(world.root.find("Соронзон", "орлого", "2026-10", "01") && !world.root.find("Тосгон", "зарлага"), "only the folders that exist get months");
const before = JSON.stringify(world.root.find("My home", "орлого").folders.map((f) => [f.name, f.folders.length]));
run(world, false, "processNewReceipts()");
eq(JSON.stringify(world.root.find("My home", "орлого").folders.map((f) => [f.name, f.folders.length])), before, "a second run makes nothing new");
world.root.find("My home", "орлого", "2026-11").folders.pop();
run(world, false, "processNewReceipts()");
eq(world.root.find("My home", "орлого", "2026-11").folders.length, 29, "the same day the folders are not looked at again");
delete world.props["ocr-folders-checked"]; // the next day
run(world, false, "processNewReceipts()");
eq(world.root.find("My home", "орлого", "2026-11").folders.length, 30, "a month left half-made is completed");

put(world, TO_SO_1, "IMG_1.HEIC", "photo-1", 1, receipt("2026-10-01", [item("Meiling 501", "Meiling 501l", 1, 1280), item("XNX 240", "xinix 240l", 1, 708, 0.45)]));
put(world, TO_SO_1, "IMG_2.HEIC", "photo-2", 2, receipt(null, [item("HAIER 400", "Haier 400l", 1, 1500)]));
put(world, MY_OUT_1, "IMG_3.HEIC", "photo-1", 3); // the same photo as IMG_1, uploaded again
put(world, SO_IN_2, "IMG_4.HEIC", "photo-4", 4, receipt("2026-10-05", [item("Haier 300", "Haier 300l", 1, 1150)], { document_type: "orlogo" }));
put(world, SO_IN_2, "IMG_5.HEIC", "photo-5", 5, { code: 503, text: "overloaded" });
put(world, ["Соронзон", "зарлага"], "IMG_6.HEIC", "photo-6", 6, receipt(null, [item("Haier 300", "Haier 300l", 1, 1180)]));
world.root.path(...SO_IN_2).add(new File("notes.txt", "text/plain", "x", 7));
world.root.path("Бусад").add(new File("IMG_9.HEIC", "image/heif", "photo-9", 8));
const snapshot = () => JSON.stringify(world.root.allFiles().filter((f) => f.mime === "image/heif").map((f) => [f.name, f.parent.id]));
const photosBefore = snapshot();

const dry2 = run(world, true, "processNewReceipts()");
eq(world.geminiCalls.length, 0, "DRY_RUN does not call Gemini (a photo is read only once, for real)");
ok(/Тосгон\/Тосгоноос Соронзон руу\/2026-10\/01\/IMG_1\.HEIC|IMG_1\.HEIC \(Тосгон \/ Тосгоноос Соронзон руу \/ 2026-10 \/ 01\) -> Gemini-ээр уншина/.test(dry2.logs) && /уншаагүй \(туршилт\) 6/.test(dry2.logs), "the dry run lists the six photos it would read");
ok(!/IMG_9|notes\.txt/.test(dry2.logs), "photos outside the place folders and non-photo files are left alone");
ok(world.queries.filter((q) => q.includes("image/")).every((q) => /modifiedDate > '\d{4}-\d{2}-\d{2}T/.test(q) && q.includes("trashed = false")), "photos are found with one search over the recent ones");

const live = run(world, false, "processNewReceipts()");
eq(live.error, null, "the real run works");
eq(world.geminiCalls.map((c) => c.bytes), ["photo-1", "photo-2", "photo-4", "photo-5", "photo-6"], "each new photo asked once, the duplicate never");
ok(world.geminiCalls.every((c) => c.key === "test-key" && c.url.includes("gemini-3.8-flash") && c.prompt.includes("Haier 400l | 1500000₮")), "the key, the model, the list with prices");
const copyIds = Object.keys(world.props).filter((key) => key.startsWith("ocr-mirror:")).map((key) => world.props[key]);
const copiesNow = world.root.allFiles().filter((f) => copyIds.includes(f.id));
eq(JSON.stringify(JSON.parse(snapshot()).filter(([, folder], index) => !copiesNow.some((f) => f.parent.id === folder && f === world.root.allFiles().filter((x) => x.mime === "image/heif")[index]))), photosBefore, "no photo is moved (mom owns them)");
eq(copiesNow.map((f) => [f.name, f.parent.parent.parent.parent.name, f.parent.parent.parent.name, f.parent.parent.name, f.parent.name]).sort(),
   [["IMG_1.HEIC", "Соронзон", "орлого", "2026-10", "01"], ["IMG_2.HEIC", "Соронзон", "орлого", "2026-10", "01"]],
   "Тосгон's transfer photos to Соронзон are copied into Соронзон's income folder of their day (the owner, 2026-10-28)");
const dayRows = tabRows(world, "Тосгон", "2026-10-01");
eq(dayRows[0].slice(0, header.length), ["Ээж", ...header.slice(1)], "the day's tab of the place's month spreadsheet starts with the header, the view choice in A1");
const PHOTO = header.indexOf("Баримт");
eq(dayRows.slice(1).map((r) => r[PHOTO]), ["IMG_1.HEIC", "IMG_1.HEIC", "IMG_2.HEIC"], "receipts of the same day folder follow each other in one spreadsheet");
{ // Gemini's confidence kept in the hidden "Итгэл"; below OCR_REVIEW_CONFIDENCE the line is Шалгах (the owner, 2026-10-25)
  const at = (title) => header.indexOf(title);
  const rows = dayRows.slice(1);
  ok(header.includes("Итгэл") && !plain(vm.runInContext("OCR_VISIBLE_COLUMNS", pure)).includes("Итгэл"), "\"Итгэл\" is a column of its own, hidden from mom");
  ok(rows.every((r) => typeof r[at("Итгэл")] === "number"), "every line keeps the confidence Gemini gave");
  eq(rows.map((r) => [r[at("Итгэл")] < 0.7, r[at("Төлөв")], String(r[at("Шалгах шалтгаан")]).includes(`итгэл бага (${r[at("Итгэл")]})`)]),
     rows.map((r) => (r[at("Итгэл")] < 0.7 ? [true, "Шалгах", true] : [false, "OK", false])), "a line under 0.7 (XNX 240, 0.45) waits for mom, saying why; the others go in OK");
  ok(rows.some((r) => r[at("Итгэл")] < 0.7), "the test has such a line");
}
const daySheet = monthTab(world, "Тосгон", "2026-10-01");
ok(world.root.find("OCR") === null, "the month spreadsheets' folder is named \"Зураг уншуулах\" (the owner, 2026-10-25), no \"OCR\" folder is made");
const octBooks = world.root.find("Зураг уншуулах", "2026-10");
eq([octBooks.files.length, Object.fromEntries(octBooks.folders.map((f) => [f.name, f.files.map((x) => x.name).sort()]))],
   [0, { "Шилжүүлэг": ["OCR Тосгон шилжүүлэг 2026-10"], "Зарлага": ["OCR My home зарлага 2026-10", "OCR Соронзон зарлага 2026-10"], "Орлого": ["OCR Соронзон орлого 2026-10"] }],
   "one spreadsheet per place, kind and month, each in its kind's folder under Баримт/Зураг уншуулах/2026-10");
eq(["орлого", "зарлага"].map((word) => monthBook(world, "Соронзон", "2026-10", word).book.getSheets().map((sheet) => sheet.getName())), [["02", "_нэрс", "Тулгалт"], ["10", "_нэрс", "Тулгалт"]],
   "a tab per day with photos of that kind; the hidden name list and the checks after them");
ok(!world.root.find(...TO_SO_1).files.some((f) => f.mime.includes("spreadsheet")), "nothing is put next to mom's photos");
const col = (title) => header.indexOf(title) + 1;
const letter = (title) => vm.runInContext(`columnLetter(${col(title)})`, pure);
ok(!header.some((title) => title.startsWith("Ажилчин")), "no staff columns: only mom checks");
eq(header.slice(header.indexOf("Бараа"), header.indexOf("Бичсэн") + 1), ["Бараа", "Тоо", "Ээж: Зөв/Буруу", "Төлөв", "Бичсэн"], "what it goes in as and mom's mark, then the script's");
eq(header.slice(-3), ["ID", "Зураг", "Шалгах шалтгаан"], "the photo column after the last of the script's, and why a line is pink right after the photo (the owner, 2026-10-28: she sees the photo and the reason side by side)");
{ // the lists on the lines only: no choice on a blank row under a photo or below the last line (the owner, 2026-10-26)
  const covered = (title) => daySheet.validations.filter((v) => v.c === col(title) && v.r >= 2).flatMap((v) => Array.from({ length: v.h }, (_, i) => v.r + i)).sort((x, y) => x - y);
  const lineRows = plain(run(world, true, `ocrReadRecords(SpreadsheetApp.openById("${monthBook(world, "Тосгон", "2026-10-01", "шилжүүлэг").book.getId()}").getSheetByName("01")).map((r) => r._row)`).value);
  eq([covered("Ээж: Зөв/Буруу").join(), covered("Бараа").join()], [lineRows.join(), lineRows.join()], "Зөв/Буруу and the product list on each line, none on the blank rows");
  ok(daySheet.validations.filter((v) => v.c === col("Ээж: Зөв/Буруу") && v.r >= 2).every((v) => plain(v.rule.list).join() === "Зөв,Буруу"), "mom chooses Зөв/Буруу from a list");
  eq([...new Set(daySheet.validations.filter((v) => v.r === 1).map((v) => JSON.stringify([v.c, plain(v.rule.list)])))], [JSON.stringify([1, ["Ээж", "Бүгд", "Бүх мөр Зөв"]])], "and the view in A1, with the whole day's Зөв");
}
eq([daySheet.value(1, 1), daySheet.value(1, 2)], ["Ээж", "Бараа"], "A1 holds the view, mom's by default; the product next to it");
{ // "Бараа" as wide as its longest name (the owner, 2026-10-28), between the bounds
  const longest = Math.max(...plain(run(world, true, `ocrReadRecords(SpreadsheetApp.openById("${daySheet.getParent().getId()}").getSheetByName("${daySheet.getName()}")).map((r) => String(r["Бараа"]).length)`).value), 5);
  eq(daySheet.getColumnWidth(col("Бараа")), Math.min(320, Math.max(100, Math.ceil(longest * 7.5 + 28))), "the product column fits its longest name");
  eq(plain(vm.runInContext(`[ocrNameWidth([]), ocrNameWidth(["Ширэм экстрем 180"]), ocrNameWidth(["x".repeat(80)])]`, pure)), [100, 156, 320], "never narrower than the title, never wider than 320px");
  const fitted = daySheet.getColumnWidth(col("Бараа"));
  daySheet.setColumnWidth(col("Бараа"), 100); // a tab set up before the fit was part of the setup
  const dry = run(world, true, "fitOcrNameColumnsNow()");
  eq([daySheet.getColumnWidth(col("Бараа")), /тааруулна \(туршилт\)/.test(dry.logs)], [100, true], "DRY_RUN only says so");
  run(world, false, "fitOcrNameColumnsNow()");
  eq(daySheet.getColumnWidth(col("Бараа")), fitted, "fitOcrNameColumnsNow fits every existing tab's product column at once");
}
{ // "Шалгах шалтгаан" as wide as its longest reason (the owner, 2026-10-28: mom reads why a line is pink without opening the cell), between the bounds
  const longest = Math.max(...plain(run(world, true, `ocrReadRecords(SpreadsheetApp.openById("${daySheet.getParent().getId()}").getSheetByName("${daySheet.getName()}")).map((r) => String(r["Шалгах шалтгаан"]).length)`).value), 15);
  eq(daySheet.getColumnWidth(col("Шалгах шалтгаан")), Math.min(700, Math.max(160, Math.ceil(longest * 7 + 8))), "the reason column fits its longest reason");
  eq(plain(vm.runInContext(`[ocrReasonWidth([]), ocrReasonWidth(["нэр жагсаалтад таараагүй"]), ocrReasonWidth(["x".repeat(300)])]`, pure)), [160, 176, 700], "never narrower than 160px, never wider than 700px (a screen beside the photo)");
  const fittedReason = daySheet.getColumnWidth(col("Шалгах шалтгаан"));
  daySheet.setColumnWidth(col("Шалгах шалтгаан"), 100);
  run(world, false, "fitOcrNameColumnsNow()");
  eq(daySheet.getColumnWidth(col("Шалгах шалтгаан")), fittedReason, "fitOcrNameColumnsNow fits the reason column too");
}
eq([1, 2, 30, 100].flatMap((row) => [1, col("Бараа"), col("Тоо"), col("Ээж: Зөв/Буруу")].map((c) => daySheet.getRange(row, c).getBackground())).filter((color) => color !== "#ffffff"), [], "no purple: mom's columns and the view choice have no colour of their own (the owner, 2026-10-26)");
eq(plain(daySheet.rules.map((r) => [r.formula, r.color])), [[`=AND($${letter("Бараа")}2="Хамааралгүй",OR($${letter("Бараа (тааруулсан)")}2<>"Хамааралгүй",$${letter("Ээж: Зөв/Буруу")}2<>""))`, "#efefef"], [`=$${letter("Ээж: Зөв/Буруу")}2="Буруу"`, "#f4cccc"], [`=OR($${letter("Ээж: Зөв/Буруу")}2="Зөв",$${letter("Ээж: Зөв/Буруу")}2="Бэлэг")`, "#d9ead3"], [`=AND($${letter("Ээж: Зөв/Буруу")}2="",OR(AND($${letter("Бараа")}2<>"",LOWER(SUBSTITUTE($${letter("Бараа")}2," ",""))<>LOWER(SUBSTITUTE($${letter("Бараа (тааруулсан)")}2," ",""))),AND($${letter("Тоо (уншсан)")}2<>"",$${letter("Тоо")}2&""<>$${letter("Тоо (уншсан)")}2&"")))`, "#d9ead3"], [`=AND($${letter("Төлөв")}2="Шалгах",$${letter("Ээж: Зөв/Буруу")}2="")`, "#ff99dd"], [`=AND($${letter("Төлөв")}2="OK",$${letter("Ээж: Зөв/Буруу")}2="")`, "#ffe599"]], "a row turns grey when it is Хамааралгүй, red when mom said Буруу, green when Зөв or Бэлэг or when she changed its product or count (her approval, as in the workbook), pink while a Шалгах line waits for her, yellow when the AI is sure: the same yellow as its cell in the workbook (the owner, 2026-10-25)");
{ // the colour legend in column O, right of the reasons in N: a sentence a cell on its own colour (the owner, 2026-10-25)
  const O = header.length + 1;
  const column = Array.from({ length: 8 }, (_, i) => daySheet.value(i + 1, O));
  eq(column[0], "Өнгө:", "the legend's title in O1, by the header");
  eq(column.slice(1, 6).map((text) => text.split(":")[0]), ["Шар мөр", "Ягаан мөр", "Ногоон мөр", "Улаан мөр", "Саарал мөр"], "then a colour a cell, O2 to O6");
  const bgAt = (row) => { const hit = daySheet.backgrounds.filter((b) => b.c === O && b.r === row && b.h === 1); return hit.length ? hit[hit.length - 1].color : undefined; };
  eq([2, 3, 4, 5, 6].map(bgAt), ["#ffe599", "#ff99dd", "#d9ead3", "#f4cccc", "#efefef"], "each cell on the colour it explains");
  ok(column[6].startsWith("Ягаан болох шалтгаан") && (column[7] === "" || column[7] === undefined), "why a line is pink in O7, nothing below");
  ok(column[6].includes("\"Шалгах шалтгаан\" баганад"), "and that each pink line says its own reason in \"Шалгах шалтгаан\" (the owner, 2026-10-28)");
  ok(!column.join(" ").includes("Нил ягаан"), "mom's purple columns are not explained (the owner: not needed)");
  ok(column[1].includes("Үндсэн хүснэгтэд ч шар"), "and that the workbook's cells mean the same");
  ok(daySheet.hidden.every((c) => c < O), "it is never hidden");
  eq(tabRows(world, "Тосгон", "2026-10-01").length, 4, "the legend's rows are no receipts: the tab still has its 3 lines");
  // the legend of before (S1, or one colour a cell from S1 to Z1) is taken away
  const old = new Sheet([["Ээж", ...header.slice(1), "", "", "", "", "Өнгө:\n■ Шар мөр: …", "Нил ягаан: …", "Ягаан мөр: …", "x", "y", "z", "w", "v"]]);
  old.getRange(1, 21).setBackground("#ff99dd").setNote("Ягаан болох шалтгаан: …");
  vm.runInContext("ocrWriteLegend", pure)(old);
  ok(Array.from({ length: 8 }, (_, i) => old.value(1, 19 + i)).every((v) => v === "") && old.getRange(1, 21).getBackground() === "#ffffff" && old.notes["1,21"] === "",
     "S1:Z1 are emptied, their colours and notes taken away");
  eq(old.value(1, O), "Өнгө:", "and O holds the legend");
  // something else in O is left as it is
  const busy = new Sheet([["Ээж", ...header.slice(1), "миний тэмдэглэл"]]);
  vm.runInContext("ocrWriteLegend", pure)(busy);
  eq([busy.value(1, O), busy.value(2, O) || ""], ["миний тэмдэглэл", ""], "a cell of someone's own in O: no legend written over it");
}
{ // the photos in M, as wide as they are, the reason after them; the photos of before, in N (until 2026-10-28) or O (until 2026-10-25), moved there
  const M = header.indexOf("Зураг") + 1;
  eq(M, header.length - 1, "the photo column is the one before the last: the reason follows it");
  ok(daySheet.images.length > 0 && daySheet.images.every((image) => image.column === M), "a new photo stands in M, after the script's columns");
  const old = new Sheet([["Ээж", ...header.slice(1)], ["x"], ["y"], ["z"]]);
  old.insertImage({}, M + 1, 2); old.insertImage({}, M + 2, 4);
  const widths = [];
  old.setColumnWidth = (column, width) => { widths.push([column, width]); return old; };
  vm.runInContext("ocrPlacePhotos", pure)(old);
  eq([old.images.map((image) => image.column), widths.find(([c]) => c === M)], [[M, M], [M, 428]], "a photo of before is moved from N or O to M, and M is made as wide as a photo (420 px) and a little");
}
const shown = header.map((title, i) => i + 1).filter((c) => !daySheet.hidden.includes(c));
eq(shown, plain(vm.runInContext("OCR_VISIBLE_COLUMNS", pure)).map(col), "mom's view: her columns; the rest hidden");
eq(shown.map((c) => header[c - 1]), ["Харах", "Бараа", "Тоо", "Ээж: Зөв/Буруу", "Зураг", "Шалгах шалтгаан"], "mom's view: the product, the count, her mark, the photo and why a line is pink (the owner, 2026-10-28: she sees it at a glance); the rest only under \"Бүгд\"");
ok(daySheet.frozen === 1 && !daySheet.hidden.includes(header.indexOf("Зураг") + 1), "frozen header, the photo column shown");
const momHidden = daySheet.hidden.slice().sort((x, y) => x - y);
const viewEdit = (row, column, value) => {
  daySheet.cells[row - 1][column - 1] = value;
  const bookId = monthBook(world, "Тосгон", "2026-10-01").file.id;
  const result = run(world, false, `onOcrViewEdit({ range: { getRow: () => ${row}, getColumn: () => ${column}, getLastColumn: () => ${column}, getValue: () => ${JSON.stringify(value)}, getSheet: () => SpreadsheetApp.openById("${bookId}").getSheetByName("01") } })`);
  assert.equal(result.error, null);
};
viewEdit(1, 1, "Бүгд");
eq(daySheet.hidden.filter((c) => c <= header.length), [], "\"Бүгд\" in A1 shows every column at once");
viewEdit(1, 1, "Ээж");
eq(daySheet.hidden.slice().sort((x, y) => x - y), momHidden, "\"Ээж\" hides them again");
viewEdit(2, 1, "Бүгд");
eq(daySheet.hidden.slice().sort((x, y) => x - y), momHidden, "an edit elsewhere changes no view");
daySheet.cells[1][0] = "";
{ // every row as tall as any, the photo over as many as it needs, the next receipt below it (the owner, 2026-10-26)
  const rowsFor = Math.ceil((560 + 12) / 21); // 28
  eq(daySheet.images.map((i) => [i.row, i.column, i.w, i.h]), [[2, header.indexOf("Зураг") + 1, 420, 560], [2 + rowsFor, header.indexOf("Зураг") + 1, 420, 560]], "each receipt's photo sits right of its own first row, 420 px wide; the next receipt starts below the photo");
  eq(Object.values(daySheet.heights).filter((h) => h !== 21), [], "no row is made taller");
  const lines = daySheet.cells.slice(1, 2 + rowsFor - 1).map((row) => (row || []).slice(0, header.length).some((v) => v !== "" && v !== undefined));
  eq([lines.filter(Boolean).length, lines.slice(2).every((x) => !x)], [2, true], "the first receipt's 2 lines, then blank rows under its photo");
  eq(run(world, true, `ocrReadRecords(SpreadsheetApp.openById("${monthBook(world, "Тосгон", "2026-10-01", "шилжүүлэг").book.getId()}").getSheetByName("01")).map((r) => r._row).join()`).value, [2, 3, 2 + rowsFor].join(),
     "the blank rows are no records; each record keeps its row");
}
ok(world.thumbnails.every((t) => /sz=s1000$/.test(t.url) && t.auth === "Bearer token-1"), "the photo comes from Drive as a JPEG, with the script's own sign-in");
const reason = (rowsOf, name) => rowsOf.filter((r) => r[PHOTO] === name).map((r) => r[header.indexOf("Шалгах шалтгаан")]).join(" | ");
const in2 = tabRows(world, "Соронзон", "2026-10-02").slice(1);
eq(in2.filter((r) => r[PHOTO] === "IMG_4.HEIC").map((r) => ["2026-10-02", r[header.indexOf("Төлөв")]]), [["2026-10-02", "OK"]], "a receipt dated 5 October in the 2 October folder stays on the 2nd, no remark");
ok(reason(in2, "IMG_5.HEIC").startsWith("Gemini уншиж чадсангүй: HTTP 503"), "a failed read is not tried again but marked for a person");
const dup = tabRows(world, "My home", "2026-10-01").slice(1);
ok(reason(dup, "IMG_3.HEIC").startsWith("давхардсан, өмнө уншсан: Тосгон/Тосгоноос Соронзон руу/2026-10/01/IMG_1.HEIC"), "the duplicate says where the first copy is");
const looseRows = tabRows(world, "Соронзон", "2026-10-10").slice(1);
ok(reason(looseRows, "IMG_6.HEIC").startsWith("өдрийн хавтсанд хийгээгүй, оруулсан өдрөөр (2026-10-10)"), "a photo put outside a day folder is read, named after its upload day, its sheet beside it, and marked");
ok(!Object.keys(world.props).some((k) => k.startsWith("ocr-result:")), "kept answers are dropped once written");

const again = run(world, false, "processNewReceipts()");
eq([again.error, world.geminiCalls.length], [null, 5], "a second run asks Gemini nothing");
ok(!/IMG_\d\.HEIC .*->/.test(again.logs), "and reports no photo again");

put(world, DAY("My home", "орлого", "03"), "IMG_2 copy.HEIC", "photo-2", 9);
run(world, false, "processNewReceipts()");
eq(world.geminiCalls.length, 5, "the same photo uploaded again is not read again");
ok(tabRows(world, "My home", "2026-10-03").some((r) => r[PHOTO] === "IMG_2 copy.HEIC"), "it is listed as a duplicate in its own day");

// a run cut off after Gemini answered: the kept answer is used, no second call
put(world, DAY("Соронзон", "орлого", "04"), "IMG_7.HEIC", "photo-7", 10);
world.props[`ocr-result:${crypto.createHash("md5").update("photo-7").digest("hex")}`] = JSON.stringify(receipt("2026-10-04", [item("Haier 300", "Haier 300l", 2, 1150)], { document_type: "orlogo" }));
run(world, false, "processNewReceipts()");
ok(world.geminiCalls.length === 5 && tabRows(world, "Соронзон", "2026-10-04"), "the kept answer is written without asking again");

// Drive gives no JPEG: the rows are still written, the log says so
world.noThumbnail = true;
put(world, DAY("Соронзон", "орлого", "05"), "IMG_10.HEIC", "photo-10", 12, receipt("2026-10-05", [item("Haier 300", "Haier 300l", 1, 1150)], { document_type: "orlogo" }));
const noPhoto = run(world, false, "processNewReceipts()");
ok(tabRows(world, "Соронзон", "2026-10-05").some((r) => r[PHOTO] === "IMG_10.HEIC") && /IMG_10\.HEIC: зургийг sheet-д оруулж чадсангүй/.test(noPhoto.logs), "without a preview the rows still stand");
world.noThumbnail = false;
eq(monthBook(world, "Соронзон", "2026-10", "орлого").book.getSheets().map((sheet) => sheet.getName()), ["02", "04", "05", "_нэрс", "Тулгалт"], "a later day's tab goes in day order");

{ // the two shops' sales spreadsheets: the receipt's price in O, ours in P, the legend in Q (the owner, 2026-10-27)
  const [O, P, Q] = [header.length + 1, header.length + 2, header.length + 3];
  put(world, DAY("Соронзон", "зарлага", "06"), "IMG_91.HEIC", "photo-91", 13,
    receipt("2026-10-06", [item("Haier 300", "Haier 300l", 1, 771), item("юу ч биш", "Хамааралгүй", 1, null)], { document_type: "zarlaga" }));
  eq(run(world, false, "processNewReceipts()").error, null, "a sales slip is read");
  const sale = monthTab(world, "Соронзон", "2026-10-06", "зарлага");
  eq([sale.value(1, O), sale.value(1, P), sale.value(1, Q)], ["Баримтын үнэ", "Манай үнэ", "Өнгө:"], "O and P hold the prices, the legend moved on to Q");
  eq([sale.value(2, O), sale.value(3, O)], [771000, ""], "the receipt's price in full tögrög, nothing where none was read");
  const formula = `=IFERROR(VLOOKUP(INDEX($B:$B,ROW()),'_нэрс'!$C:$D,2,FALSE),"")`;
  eq([sale.value(2, P), sale.value(3, P)], [formula, formula], "ours: the price beside the line's product on the name tab, so it follows mom's change");
  const names = sale.getParent().getSheetByName("_нэрс").cells.filter((row) => row && row[2] === "Haier 300l");
  ok(names.length === 1 && typeof names[0][3] === "number" && names[0][3] > 0, "the name tab has the price the customer sees beside each name");
  eq(tabRows(world, "Соронзон", "2026-10-06", "зарлага").length, 3, "the price cells are no receipts: the tab has its 2 lines");
  const income = monthTab(world, "Соронзон", "2026-10-05", "орлого");
  eq([income.value(1, O), income.value(1, P) || "", income.value(2, P) || ""], ["Өнгө:", "", ""], "every other spreadsheet keeps the legend in O and has no prices");
  // a sales tab of before: the legend in O goes to Q, O and P get the prices
  const before = new Sheet([["Ээж", ...header.slice(1), "", "Өнгө:"], ["", "Haier 300l", 1, "", "OK", "", "", "", "", "", "", "", "x#1"]]);
  new Book("b1", "OCR My home зарлага 2026-10", { "01": before });
  vm.runInContext("(sheet) => { ocrWriteLegend(sheet); ocrWritePriceColumns(sheet); }", pure)(before);
  eq([before.value(1, O), before.value(1, P), before.value(1, Q), before.value(2, O) || "", before.value(2, P)], ["Баримтын үнэ", "Манай үнэ", "Өнгө:", "", formula],
     "the old legend in O is taken away, the prices put in, ours on the existing line");
  // Тосгон's income: the yuan read beside Тосгон's cost of the month (the owner, 2026-10-27)
  const tos = new Sheet([["Ээж", ...header.slice(1), "", "Өнгө:"], ["", "Haier 200l", 30, "", "Шалгах", "", "", "", "", "", "", "", "y#1"]]);
  const tosBook = new Book("b2", "OCR Тосгон орлого 2026-10", { "17": tos });
  vm.runInContext("(sheet, book) => { ocrWriteLegend(sheet); ocrWritePriceColumns(sheet); ocrSetCosts(book, [['Haier 200l', 450000]]); }", pure)(tos, tosBook);
  eq([tos.value(1, O), tos.value(1, P), tos.value(1, Q), tos.value(2, P)], ["Баримтын үнэ", "Өртөг үнэ", "Өнгө:", `=IFERROR(VLOOKUP(INDEX($B:$B,ROW()),'_өртөг'!$A:$B,2,FALSE),"")`],
     "Тосгон's income: P is Тосгон's cost, looked up on its hidden cost tab");
  eq(tosBook.getSheetByName("_өртөг").cells[0].slice(0, 2), ["Haier 200l", 450000], "the cost tab holds name and cost");
  eq(plain(vm.runInContext("ocrReadPrices", pure)({ price_currency: "CNY", items: [item("Haier 200", "Haier 200l", 30, 850), item("x", "Хамааралгүй", 1, null)] })), ["850 юань", ""],
     "a yuan price is written as it was read, its currency named");
  const shopIncome = new Sheet([["Ээж", ...header.slice(1)]]);
  new Book("b3", "OCR Соронзон орлого 2026-10", { "17": shopIncome });
  vm.runInContext("(sheet) => { ocrWriteLegend(sheet); ocrWritePriceColumns(sheet); }", pure)(shopIncome);
  eq([shopIncome.value(1, O), shopIncome.value(1, P) || ""], ["Өнгө:", ""], "a shop's income spreadsheet has no prices");
}

// a count sheet: counted numbers, no price needed, into the count column
put(world, ["Тосгон", "тооллого", "2026-10", "31"], "IMG_11.HEIC", "photo-11", 13, receipt(null, [item("Haier 300", "Haier 300l", 4, null), item("Meiling 501", "Meiling 501l", 0, null)], { document_type: "toollogo" }));
run(world, false, "processNewReceipts()");
const countRows = tabRows(world, "Тосгон", "2026-10-31").slice(1);
eq(countRows.map((r) => [r[header.indexOf("Бараа")], r[header.indexOf("Тоо")], r[header.indexOf("Төлөв")]]),
  [["Haier 300l", 4, "OK"], ["Meiling 501l", 0, "OK"]], "counted numbers (also 0) go to Тосгон's count column, no price needed");

// a wrong key stops the run before the photo counts as read
put(world, DAY("Соронзон", "зарлага", "05"), "IMG_8.HEIC", "photo-8", 11, { code: 401, text: "API key not valid" });
const bad = run(world, false, "processNewReceipts()");
ok(bad.error && /HTTP 401/.test(bad.error.message), "a wrong key stops the run with the reason");
ok(!Object.values(world.props).some((v) => String(v).includes("IMG_8")), "the photo is not marked as read");

delete world.props.GEMINI_KEY;
ok(/GEMINI_KEY/.test(String(run(world, false, "processNewReceipts()").error)), "no key: a clear error");

// the hourly trigger
const viewTriggers = world.triggers.filter((t) => t.getHandlerFunction() === "onOcrViewEdit");
const ocrBookIds = world.root.find("Зураг уншуулах", "2026-10").folders.flatMap((folder) => folder.files.map((f) => f.id)).sort();
eq(viewTriggers.map((t) => t.getTriggerSourceId()).sort(), ocrBookIds, "each month spreadsheet has one view trigger, made by the review");
world.triggers = world.triggers.filter((t) => !viewTriggers.includes(t));
run(world, true, "installOcrTrigger()");
eq(world.triggers.length, 0, "DRY_RUN makes no trigger");
run(world, false, "installOcrTrigger()"); run(world, false, "installOcrTrigger()");
eq(world.triggers.map((t) => [t.getHandlerFunction(), t.hours || null, Boolean(t.book)]), [["processNewReceipts", 1, false], ["onCountReportEdit", null, true]], "one hourly trigger and the count report's edit trigger, also after installing twice");

// ---------- the two checks, as decisions ----------
const rec = (values) => pure.ocrRecord(header.map((title) => values[title] ?? ""));
const okLine = { "Бараа (уншсан)": "Haier 300", "Бараа": "Haier 300l", "Бараа (тааруулсан)": "Haier 300l", "Тоо": 1, "Тоо (уншсан)": 1, "Төлөв": "OK" };
const reviewLine = { ...okLine, "Төлөв": "Шалгах" };
const decide = (values) => plain(pure.ocrPostDecision(rec(values)));
eq(decide(okLine), { post: true, name: "Haier 300l", qty: 1, approved: false, check: false }, "a line with a product is written at once, not yet approved (yellow: the AI is sure)");
eq(decide(reviewLine), { post: true, name: "Haier 300l", qty: 1, approved: false, check: true }, "... also a Шалгах line: every line with a name goes in, marked to be checked (pink)");
eq(decide({ ...reviewLine, "Ээж: Зөв/Буруу": "Зөв" }), { post: true, name: "Haier 300l", qty: 1, approved: true, check: false }, "mom's Зөв approves it (green)");
eq(decide({ ...reviewLine, "Тоо": 2 }), { post: true, name: "Haier 300l", qty: 2, approved: true, check: false }, "mom writing another count in \"Тоо\" approves it too");
eq(decide({ ...reviewLine, "Бараа": " Haier  400l", "Тоо": "2" }), { post: true, name: "Haier 400l", qty: 2, approved: true, check: false }, "... and so does picking another product in \"Бараа\"");
eq(decide({ "Тоо (уншсан)": 1, "Тоо": 1, "Төлөв": "Шалгах" }), { post: false, wait: "ээж" }, "a line read with no product (before 2026-10-25) waits for mom to pick one");
eq(decide({ "Тоо (уншсан)": 1, "Тоо": 1, "Төлөв": "Шалгах", "Бараа": "Haier 300l" }), { post: true, name: "Haier 300l", qty: 1, approved: true, check: false }, "... and is written, approved, once she has");
eq([decide({ ...reviewLine, "Бараа": "Хамааралгүй", "Бараа (тааруулсан)": "Хамааралгүй" }), plain(pure.ocrReviewState(rec({ ...reviewLine, "Бараа": "Хамааралгүй", "Бараа (тааруулсан)": "Хамааралгүй" }))).needsMom], [{ post: false, wait: "" }, true],
   "Хамааралгүй the script picked is not written, and still waits for mom to look"); 
eq(decide({ ...reviewLine, "Бараа": "Haier 400l", "Ээж: Зөв/Буруу": "Буруу" }), { post: false, wait: "" }, "her Буруу still wins");
eq(decide({ ...okLine, "Ээж: Зөв/Буруу": "Буруу" }), { post: false, wait: "" }, "mom's Буруу: never written");
eq(decide({ "Төлөв": "Шалгах", "Ээж: Зөв/Буруу": "Зөв" }), { post: false, wait: "" }, "a row without a product (a duplicate, a failed read) is not written");
ok(pure.ocrValueProblem({ name: "", qty: 1 }).startsWith("бараа сонгоогүй") && pure.ocrValueProblem({ name: "x", qty: NaN }).startsWith("тоо буруу") && pure.ocrValueProblem({ name: "x", qty: 0 }) === "", "a name and a number are needed (0 is a number)");
ok(pure.ocrValueProblem({ name: "", qty: 1 }).includes("\"Бараа\"") && pure.ocrValueProblem({ name: "x", qty: "a" }).includes("\"Тоо\""), "the problem says which of mom's columns to fill");
const state = (values) => plain(pure.ocrReviewState(rec(values)));
eq([state({}), state({ "Ээж: Зөв/Буруу": "Зөв" }).needsMom, state({ ...reviewLine, "Бараа": "Haier 400l" }).needsMom], [{ mom: "", chosen: false, needsMom: true }, false, false], "mom is the only one who still has to look; her choice of product counts as looked at");
const notRelevant = { "Бараа (уншсан)": "Никел тоо", "Тоо": 1, "Тоо (уншсан)": 1, "Төлөв": "Шалгах", "Бараа": "Хамааралгүй" };
eq([decide(notRelevant), decide({ ...okLine, "Бараа": "Хамааралгүй", "Ээж: Зөв/Буруу": "Зөв" })], [{ post: false, wait: "" }, { post: false, wait: "" }], "\"Хамааралгүй\" (a bicycle, never counted): never written, not even with Зөв");
eq(state(notRelevant).needsMom, false, "... and it counts as looked at");
const notRelevantSummary = plain(pure.ocrSheetSummary({ name: "17" }, [rec(notRelevant)]));
eq([notRelevantSummary.mom, notRelevantSummary.failed, notRelevantSummary.waiting, notRelevantSummary.check], [0, [], [], []], "... so neither mom's report nor the count report shows it as waiting or missing");
const folderOf = (photo, values) => rec({ ...values, "Хавтас (ээж хийсэн)": "2026-10-17 / зарлага", "ID": `${photo}#1` });
const calendarDay = plain(pure.ocrCalendarDay([
  folderOf("P1", reviewLine), folderOf("P2", okLine), folderOf("P3", { ...reviewLine, "Ээж: Зөв/Буруу": "Зөв" }),
  folderOf("P4", notRelevant), folderOf("P5", { "Төлөв": "Шалгах" }),
]));
eq(calendarDay, { "зарлага": [5, 4, 2, 0, 1, 1, 0] }, "the calendar's sixth number: Шалгах lines mom has not checked (an OK line only waits for her; Зөв and Хамааралгүй are done)");
const legacy = plain(pure.ocrUpgradeRecord({ "Ажилчин: Зөв/Буруу": "Буруу", "Ажилчин: зөв бараа": "Haier 400l", "Ажилчин: зөв тоо": 2, "Ажилчин: тайлбар": "хайрцаг", "Ээж: Зөв/Буруу": "", "Ээж: тайлбар": "дараа" }));
eq([legacy["Бараа"], legacy["Тоо"], legacy["Шалгах шалтгаан"]], ["Haier 400l", 2, "тайлбар: дараа · ажилчин Буруу: хайрцаг"], "an old row: the staff's correction becomes mom's, their mark goes into her note");
const upgraded = plain(pure.ocrUpgradeRecord({ "Бараа (Sheet100)": "C", "Ажилчин: зөв бараа": "A", "Ээж: зөв бараа": "B", "Ээж: тайлбар": "" }));
eq([upgraded["Бараа"], upgraded["Бараа (тааруулсан)"]], ["B", "C"], "mom's own correction wins and becomes \"Бараа\"; the matched name is kept");
const plainOld = plain(pure.ocrUpgradeRecord({ "Бараа (Sheet100)": "C", "Ээж: зөв бараа": "" }));
eq([plainOld["Бараа"], plainOld["Бараа (тааруулсан)"]], ["C", "C"], "without a correction \"Бараа\" is the matched name: nothing counts as mom's choice");
const idRec = rec({ ...okLine, "Баримт": "IMG", "ID": "x#1" });
const change = (current, note, qty, isCount = false, approved = false, isCheck = false, needsCheck = false) => {
  const result = plain(pure.ocrCellChange(current, note, idRec, qty, isCount, "", approved, isCheck, null, needsCheck));
  delete result.entries;
  return result;
};
const both = "OCR: B [y#1] = 1\nOCR: IMG [x#1] = 2";
eq(change("", "", 2), { action: "write", value: 2, note: "Зураг: https://drive.google.com/file/d/x/view#1 = 2", color: "#ffe599" }, "an empty cell is written yellow, its note is the photo's link with the line and its count (the owner, 2026-10-25)");
eq(change(1, "OCR: B [y#1] = 1", 2), { action: "write", value: 3, note: "Зураг: https://drive.google.com/file/d/y/view#1 = 1\nЗураг: https://drive.google.com/file/d/x/view#1 = 2", color: "#ffe599" }, "a second receipt of that day adds to it");
eq(change(3, both, 2).action, "skip", "a line already in the cell is not written again");
eq(change(3, both, 5), { action: "write", value: 6, note: "Зураг: https://drive.google.com/file/d/y/view#1 = 1\nЗураг: https://drive.google.com/file/d/x/view#1 = 5", color: "#ffe599" }, "a corrected count replaces its own part only");
eq(change(3, both, null), { action: "write", value: 1, note: "Зураг: https://drive.google.com/file/d/y/view#1 = 1", color: "#ffe599" }, "a line taken out leaves the others");
eq(change(2, "OCR: IMG [x#1] = 2", null), { action: "write", value: "", note: "", color: null }, "the last line taken out empties the cell and its colour");
eq(change(2, "OCR: IMG [x#1] = 2", 2, false, true), { action: "write", value: 2, note: "Зураг: https://drive.google.com/file/d/x/view#1 = 2 ✓", color: "#d9ead3" }, "mom's approval marks its line and turns the cell green");
eq(change(3, "OCR: B [y#1] = 1\nOCR: IMG [x#1] = 2 ✓", 2, false, true).action, "skip", "an approved line is read back from the note");
eq(change(1, "OCR: B [y#1] = 1", 2, false, true).color, "#ffe599", "a cell stays yellow while one of its lines is not approved");
eq(change("", "", 2, false, false, false, true), { action: "write", value: 2, note: "Зураг: https://drive.google.com/file/d/x/view#1 = 2 ⚠", color: "#ff99dd" }, "a Шалгах line mom has not marked is written pink, its note line marked ⚠ (the owner, 2026-10-25)");
eq(change(1, "OCR: B [y#1] = 1 ✓", 2, false, false, false, true).color, "#ff99dd", "one unchecked Шалгах line makes the whole cell pink: the worst state wins");
eq(change(1, "OCR: B [y#1] = 1 ⚠", 2).color, "#ff99dd", "... also one read back from the note");
eq(change(2, "OCR: IMG [x#1] = 2 ⚠", 2, false, false, false, true).action, "skip", "a marked line is read back from the note");
eq(change(2, "OCR: IMG [x#1] = 2", 2, false, false, false, true).color, "#ff99dd", "a cell written before the pink is recoloured at the next review");
eq(change(2, "OCR: IMG [x#1] = 2 ⚠", 2, false, true).color, "#d9ead3", "mom's mark takes the pink away: green");
eq(vm.runInContext('ocrWrittenLine("Tos/2026/10!BY5", { qty: 1, name: "Haier 300l", approved: false, check: true })', pure), "Tos/2026/10!BY5 ← 1 · Haier 300l ⚠", "\"Бичсэн\" marks a line still to be checked");
eq(plain(pure.ocrParseWrittenLine("Tos/2026/10!BY5 ← 1 · Haier 300l ⚠")).name, "Haier 300l", "and the mark is not part of the name read back");
eq(change(3, "", 2).reason, "нүдэнд гараар бичсэн 3 байна", "a person's number is never written over");
eq(change(4, "OCR: B [y#1] = 1", 2).reason, "нүдийг гараар өөрчилсөн (4, тэмдэглэлээр 1)", "nor a cell changed by hand after the script wrote it");
eq(change(4, "OCR: B [y#1] = 4", 2, true).action, "refuse", "two counts are not added up");
eq(change(5, "мамын тэмдэглэл", 2, false, false, true), { action: "write", value: 2, note: "Зураг: https://drive.google.com/file/d/x/view#1 = 2", color: "#ffe599" }, "a check copy's cell holds the receipts only: what was copied from the hand-typed cell is replaced");
{ // a note line is the photo's link, its line, count and day (the owner, 2026-10-25); the notes written before still read
  const line = "Зураг: https://drive.google.com/file/d/abc/view#2 = 1 (2026-10-03)";
  const old = "OCR: IMG [abc#2] = 1 @2026-10-03\nЗураг: https://drive.google.com/file/d/abc/view";
  eq(run(makeWorld(), true, 'ocrNoteLine({ id: "abc#2", qty: 1, day: "2026-10-03", approved: false, elsewhere: null })').value, line, "no \"OCR:\", file name or ID: the link only, with the line's count and day");
  eq(run(makeWorld(), true, 'ocrNoteLine({ id: "abc#2", qty: 0, day: "", approved: true, check: false, elsewhere: { day: 12, cell: "N214" } })').value,
    "Зураг: https://drive.google.com/file/d/abc/view#2 = 0 ✓ (ээж 12-нд гараар: N214)", "an approved line moved to mom's day keeps both marks");
  const read = (note) => plain(pure.ocrNoteEntries(note)).map((e) => [e.id, e.qty, e.day, e.approved, e.check, e.elsewhere]);
  eq(read(line), [["abc#2", 1, "2026-10-03", false, false, null]], "the new line reads back");
  eq(read("Зураг: https://drive.google.com/file/d/abc/view#2 = 0 ✓ (ээж 12-нд гараар: N214)\nЗураг: https://drive.google.com/file/d/d-e_f/view#7 = 3 ⚠"),
    [["abc#2", 0, "", true, false, { day: 12, cell: "N214" }], ["d-e_f#7", 3, "", false, true, null]], "... with its marks, two to a cell");
  eq(read(old), [["abc#2", 1, "2026-10-03", false, false, null]], "an old note reads the same; its bare link line is not a line of its own");
  eq([pure.ocrIsHandCell(1, line), pure.ocrIsHandCell(1, old), pure.ocrIsHandCell(1, "OCR: IMG [abc#2] = 1 @2026-10-03")], [false, false, false], "a cell with a note of either kind is the script's");
  eq(pure.ocrIsHandCell(1, line + "\nээж: 1 нь буцаалт"), true, "a line of mom's own makes it hand-typed");
  eq(change(2, "OCR: IMG [x#1] = 2", 2).action, "skip", "an old note is not rewritten by a line whose row is not known (addSheetLinksToOcrNotes does it once)");
  // the line's row known: the note links to it (the owner, 2026-10-26)
  const at = { book: "bk", gid: 7, row: 30 };
  const atChange = (current, note, qty) => plain(pure.ocrCellChange(current, note, { ...rec({ ...okLine, "ID": "x#1" }), _at: at }, qty, false, "2026-10-23"));
  const sheetLine = "https://docs.google.com/spreadsheets/d/bk/edit#gid=7&range=A30:N30&ocr=x.1 = 2 (2026-10-23)";
  eq(atChange("", "", 2).note, sheetLine, "no \"Зураг:\": the link opens the line's row in its day tab, the ID kept after \"&ocr=\"");
  eq(read(sheetLine + " ⚠\nЗураг: https://drive.google.com/file/d/y/view#2 = 1"), [["x#1", 2, "2026-10-23", false, true, null], ["y#2", 1, "", false, false, null]], "it reads back, next to a line of the photo's kind");
  eq(plain(pure.ocrNoteEntries(sheetLine))[0].at, at, "with the row it links to");
  eq([atChange(2, sheetLine, 2).action, pure.ocrIsHandCell(2, sheetLine)], ["skip", false], "a line linked to its row is not written again; the cell is the script's");
  eq(atChange(2, "Зураг: https://drive.google.com/file/d/x/view#1 = 2 (2026-10-23)", 2).note, sheetLine, "a line of the photo's kind is linked to its row when written again");
  eq(atChange(2, sheetLine.replace(/30/g, "31"), 2).note, sheetLine, "... and so is a line whose row moved");
  eq(change(1, "OCR: B [y#1] = 1\nЗураг: https://drive.google.com/file/d/y/view", 2).note, "Зураг: https://drive.google.com/file/d/y/view#1 = 1\nЗураг: https://drive.google.com/file/d/x/view#1 = 2", "a cell written again is all in the new way");
}
eq([pure.ocrIsHandCell(3, ""), pure.ocrIsHandCell("", "x"), pure.ocrIsHandCell(2, "OCR: IMG [x#1] = 2"), pure.ocrIsHandCell("", "")], [true, true, false, false], "which cells were typed by hand");
const writtenRec = { ...okLine, "Бичсэн": "Tos/2026/10!AJ5 ← 1 · Haier 300l" };
eq([pure.ocrWrittenState(rec(okLine)), pure.ocrWrittenState(rec(writtenRec))], ["new", "done"], "written or not");
eq(pure.ocrWrittenState(rec({ ...okLine, "Бичсэн": pure.ocrFailedText("x", rec(okLine)) })), "failed", "a failed line is not tried again as it is");
eq(pure.ocrWrittenState(rec({ ...okLine, "Ээж: Зөв/Буруу": "Зөв", "Бичсэн": pure.ocrFailedText("x", rec(okLine)) })), "retry", "... but once someone changed its checks");
eq(pure.ocrWrittenState(rec({ ...okLine, "Бичсэн": "— workbook-оос хассан (ээж Буруу гэсэн)" })), "removed", "a line mom took out");
eq(pure.ocrWrittenState(rec({ ...okLine, "Бараа": "Хамааралгүй", "Бичсэн": "— workbook-оос хассан (Хамааралгүй)" })), "removed", "... also one she made \"Хамааралгүй\"");
eq(plain(vm.runInContext('ocrWrittenCells("✗ x [k:a]\\nTos/2026/10!AJ5 ← 1 · Haier 300l")', pure)), ["Tos/2026/10!AJ5"], "a failed line still knows where it was written before, to take it out later");
eq(plain(pure.ocrIdsFromLinks(["https://drive.google.com/file/d/AB-1/view", "https://drive.google.com/file/d/AB-1/view", "https://x?id=CD"])), [["AB-1#1"], ["AB-1#2"], ["CD#1"]], "rows of an old sheet get IDs from their links");
const momOf = (raw) => run({ ...makeWorld(), props: { OCR_REVIEWERS: raw } }, true, "ocrMomEmail(PropertiesService.getScriptProperties())").value;
eq([momOf(JSON.stringify({ "Соронзон": "so@gmail.com", "Ээж": "mom@gmail.com" })), momOf(JSON.stringify({ "Ээж": "bad" })), momOf("{")], ["mom@gmail.com", "", ""], "only mom's address is read; a staff address is ignored, a broken value gives none");
eq([momOf(JSON.stringify({ "Ээж": "mom@gmail.com, own@gmail.com" })), momOf(JSON.stringify({ "Ээж": ["mom@gmail.com", "bad", "own@gmail.com", "mom@gmail.com"] }))], ["mom@gmail.com,own@gmail.com", "mom@gmail.com,own@gmail.com"],
   "mom may have several accounts, by commas or as a list (the owner, 2026-10-27)");

// ---------- mom's report, as text ----------
const summary = (extra) => ({ id: "s1", url: "https://sheet/1", name: "OCR 2026-10-01", place: "Соронзон", label: "Соронзон / зарлага / 2026-10 / 01", rows: 3, mom: 0, failed: [], ...extra });
const sums = [summary({ mom: 3 }), summary({ id: "s2", url: "https://sheet/2", name: "OCR 2026-09-30", place: "My home", label: "My home / зарлага / 2026-09 / 30", mom: 1,
  failed: [{ product: "xinix 240l", reason: "бараа алга" }] })];
const isOpen = (extra) => vm.runInContext(`ocrSummaryIsOpen(${JSON.stringify(summary(extra))})`, pure);
ok(isOpen({ mom: 1 }) && !isOpen({}), "a day is open while mom has lines to check");
eq(["2026-10-15", "2026-10-31", "2026-11-30", "2026-02-28", "2026-10-17", "2026-10-30", "2026-10-16"].map((iso) => pure.ocrIsReportDay(iso)), [true, true, true, true, false, false, false], "the report comes out only on the 15th and the month's last day");
eq(["2026-10-15", "2026-10-31", "2026-10-03"].map((iso) => pure.ocrHalfStart(iso)), ["2026-10-01", "2026-10-16", "2026-10-01"], "it covers its half-month: 1-15 or 16 to the end");
const periodProps = { getProperties: () => ({ "ocr-day:2026-10-15": JSON.stringify({ photos: 9, counts: [{ place: "Тосгон", day: "2026-10-15" }] }), "ocr-day:2026-10-16": JSON.stringify({ photos: 2, rows: 4, counts: [{ place: "Тосгон", day: "2026-10-20" }] }), "ocr-day:2026-10-31": JSON.stringify({ photos: 3, rows: 1, counts: [{ place: "Тосгон", day: "2026-10-20" }, { place: "Соронзон", day: "2026-10-31" }] }) }) };
periodProps.getProperty = (key) => periodProps.getProperties()[key] || null;
eq(plain(pure.ocrPeriodStats(periodProps, "2026-10-16", "2026-10-31")), { photos: 5, rows: 5, ok: 0, written: 0, corrected: 0, failed: 0, counts: [{ place: "Тосгон", day: "2026-10-20" }, { place: "Соронзон", day: "2026-10-31" }] }, "the numbers of the half-month's days added up, each count once, the day before left out");
eq(plain(pure.ocrKeptCounts(periodProps, "2026-10-16")), [{ place: "Тосгон", day: "2026-10-20" }, { place: "Соронзон", day: "2026-10-31" }], "the counts of the half-month only");
const report = plain(pure.ocrReportLines("2026-10-01", "2026-10-15", "14:00", { photos: 2, rows: 5, ok: 3, written: 3, corrected: 1, failed: 1, counts: [] }, sums, [{ text: "Тооллогын тайлан Тосгон 2026-10-15: дутсан бараа алга", url: "https://doc/1" }]));
const reportText = report.map((line) => line.text).join("\n");
ok(report[0].style === "title" && report[0].text === "Баримтын тайлан 2026-10-01 – 2026-10-15 (14:00-ийн байдлаар)" && reportText.includes("Энэ хагас сар") && reportText.includes("Уншсан зураг 2, мөр 5 (шууд OK 3).") && reportText.includes("Workbook-д бичсэн мөр 3, засвар орсон 1, бичиж чадаагүй 1."), "the report starts with the half-month's numbers");
const at = (start) => report.findIndex((line) => line.text.startsWith(start));
ok(reportText.includes("Ээжийн шалгах мөр (нийт 4)") && at("2026-09-30 · My home / зарлага · 1 мөр шалгах") < at("2026-10-01 · Соронзон / зарлага · 3 мөр шалгах"), "what waits for mom, oldest first");
ok(!reportText.includes("Ажилчд") && !reportText.includes("санал зөрсөн"), "no staff sections any more");
ok(at("xinix 240l: бараа алга") > 0 && at("Тооллогын тайлан Тосгон 2026-10-15") > 0 && !reportText.includes("гараар засах"), "what could not be written, the count reports; nothing to fix by hand");
ok(report.filter((line) => line.style === "item").every((line) => line.url), "every line links to its sheet or report");

// ---------- the count report ----------
const closing = { place: "Тосгон", day: "2026-10-31", tab: "Tos/2026/10", isClosing: true,
  counted: [{ name: "Haier 300l", counted: 4 }, { name: "Meiling 501l", counted: 2 }, { name: "xinix 240l", counted: 3 }],
  balances: [{ name: "Haier 300l", computed: 5 }, { name: "Meiling 501l", computed: 2 }, { name: "xinix 240l", computed: 1 }, { name: "Haier 400l", computed: 2 }, { name: "Aucma", computed: 0 }] };
const waitingFor = (summaries, place = "Тосгон") => pure.ocrWaitingByProduct(summaries, place);
const tableOf = (count, waiting = waitingFor([]), elsewhere = new Map()) => plain(pure.ocrCountTable(count, waiting, elsewhere));
const closingTable = tableOf(closing);
eq(closingTable.summary, "Тосгон, 2026-10-31 (хаалтын тооллого, Tos/2026/10): тоолсон бараа 3, таарсан 1, дутсан 1 (1 ширхэг), илүү 1 (2 ширхэг).", "how many counted, agree, short, over");
eq(closingTable.conclusion, "Дүгнэлт: 3 бараанаас 1 нь дутсан байна (нийт 1 ширхэг).", "it says plainly how many are short");
eq(closingTable.rows.map((row) => row.slice(0, 4)), [["Haier 300l", 4, 5, -1], ["xinix 240l", 3, 1, 2], ["Meiling 501l", 2, 2, 0]], "a row per product: short first, then over, then the ones that agree");
ok(closingTable.rows[0][4].startsWith("Workbook-д ороогүй баримт олдсонгүй, тиймээс бараа үнэхээр дутсан байна.") && closingTable.rows[1][4].startsWith("Орлого эсвэл буцаалт бүртгэгдээгүй") && closingTable.rows[2][4] === "Таарсан", "each with its likely reason");
eq(closingTable.uncounted, [["Haier 400l", 2]], "a closing count lists what nobody counted (with a balance)");
const spotTable = tableOf({ ...closing, isClosing: false, day: "2026-10-08" });
ok(spotTable.uncounted.length === 0 && spotTable.summary.includes("шалгах тооллого, workbook-д бичээгүй"), "a check count does not list what it did not count");
const waited = tableOf(closing, waitingFor([{ waiting: [{ product: "Haier 300l", places: ["Тосгон", "Соронзон"] }, { product: "Haier 300l", places: ["My home"] }] }]));
ok(waited.rows[0][4].startsWith("Энэ барааны 1 мөр баримт шалгалт хүлээж workbook-д ороогүй байна"), "a shortage with a receipt of that place still waiting points to it first");
eq(tableOf({ ...closing, counted: [] }).conclusion, "Дүгнэлт: энэ өдөр тооллого ороогүй байна.", "a day without a count says so");
const shopCount = (place, counted, computed) => ({ place, day: "2026-10-31", tab: "t", isClosing: false, counted: [{ name: "Haier 300l", counted }], balances: [{ name: "Haier 300l", computed }] });
const elsewhere = pure.ocrDiffsElsewhere([{ place: "Соронзон", items: [{ name: "Haier 300l", counted: 6, computed: 5 }] }, { place: "Тосгон", items: [{ name: "Haier 300l", counted: 5, computed: 5 }] }]);
ok(tableOf(shopCount("My home", 2, 3), waitingFor([], "My home"), elsewhere).rows[0][4].startsWith("Соронзонд 1 ширхэг илүү байна: My home-оос Соронзон руу шилжүүлэг бүртгэгдээгүй"), "My home 1 short and Соронзон 1 over: an unwritten transfer is named first");
const myShort = pure.ocrDiffsElsewhere([{ place: "My home", items: [{ name: "Haier 300l", counted: 2, computed: 3 }] }]);
ok(tableOf(shopCount("Соронзон", 6, 5), waitingFor([], "Соронзон"), myShort).rows[0][4].startsWith("My home-д 1 ширхэг дутсан байна: My home-оос Соронзон руу шилжүүлэг бүртгэгдээгүй"), "... and the same from Соронзон's side");

// ---------- a whole run with the checks ----------
const rw = makeWorld();
rw.props.GEMINI_KEY = "k";
rw.props.OCR_REVIEWERS = JSON.stringify({ "Соронзон": "so@gmail.com", "My home": ["mh1@gmail.com", "mh2@gmail.com"] }); // staff addresses of the old version: ignored
const T1 = DAY("Тосгон", "Тосгоноос Соронзон руу", "01"), M3 = DAY("My home", "My home-оос Соронзон руу", "03"), C31 = ["Тосгон", "тооллого", "2026-10", "31"];
const AJ = 36, BO = 67, E = 5, T = 20;
tos(rw).cells[5][AJ - 1] = 3; // Haier 400l, 1 October: typed by hand
tos(rw).cells[4][1] = 5; // Haier 300l: opening balance 5
const photoA = put(rw, T1, "A.HEIC", "a", 1, receipt(null, [item("Meiling 501", "Meiling 501l", 1, 1280), item("XNX 240", "xinix 240l", 1, 708, 0.7), item("HAIER 400", "Haier 400l", 1, 1500)]));
put(rw, M3, "B.HEIC", "b", 2, receipt(null, [item("Haier 300", "Haier 300l", 2, null)]));
put(rw, C31, "C.HEIC", "c", 3, receipt(null, [item("Haier 300", "Haier 300l", 4, null)], { document_type: "toollogo" }));
const first = run(rw, false, "processNewReceipts()");
eq(first.error, null, "the run with the checks works");
// a place's month: its kinds' days put together (kept one per kind: ocr-cal:Тосгон|2026-10|шилжүүлэг)
const calendarOf = (world, place) => {
  const keys = Object.keys(world.props).filter((key) => key.startsWith(`ocr-cal:${place}|2026-10|`));
  return keys.length === 0 ? null : keys.reduce((days, key) => { Object.entries(JSON.parse(world.props[key])).forEach(([day, folders]) => { days[day] = { ...(days[day] || {}), ...folders }; }); return days; }, {});
};
ok(!(`ocr-cal:Тосгон|2026-10` in rw.props) && "ocr-cal:Тосгон|2026-10|тооллого" in rw.props && "ocr-cal:Тосгон|2026-10|шилжүүлэг" in rw.props, "the calendar is kept per kind");
const tosDay = calendarOf(rw, "Тосгон");
eq([Object.keys(tosDay).sort(), tosDay["01"]["Тосгоноос Соронзон руу"].slice(0, 3), tosDay["01"]["Тосгоноос Соронзон руу"][4], Object.keys(tosDay["31"])], [["01", "31"], [1, 3, 3], 0, ["тооллого"]], "the calendar: per day and folder the photos, the lines and the lines mom has not marked");
ok(tosDay["01"]["Тосгоноос Соронзон руу"][3] >= 1 && calendarOf(rw, "My home")["03"]["My home-оос Соронзон руу"][0] === 1, "... and the lines already in the workbook");
rw.props.OCR_FEED_TOKEN = "secret-token";
const feed = (params) => plain(run(rw, true, `ocrFeedAnswer(${JSON.stringify(params)}, PropertiesService.getScriptProperties(), new Date())`).value);
const answer = feed({ token: "secret-token", mode: "receipts", month: "2026-10" });
eq([answer.ok, answer.month, answer.months, Object.keys(answer.places), answer.places["Тосгон"]["01"]["Тосгоноос Соронзон руу"][0]], [true, "2026-10", ["2026-10"], ["Тосгон", "Соронзон", "My home"], 1], "the site's web app answers a month of the calendar");
eq([feed({ token: "wrong", mode: "receipts" }), feed({ mode: "receipts" }), feed({ token: "secret-token", mode: "stock" }).error], [{ ok: false, error: "forbidden" }, { ok: false, error: "forbidden" }, "unknown mode"], "only with the token, only this answer");
eq(Object.keys(feed({ token: "secret-token", mode: "receipts", month: "2026-11" }).places["Соронзон"]), [], "a month without photos is empty");
ok(answer.checks["Тосгон"]["01"].some((c) => c.name === "Meiling 501l" && c.status === "filled" && c.hand === null && c.receipt === 1 && c.sheet === "Tos/2026/10" && c.cell === "AJ4"), "a cell the receipts filled in is \"filled\"");
// the site turns a day's numbers into a link to that day's tab (2026-10-22)
{
  const tosBook = monthBook(rw, "Тосгон", "2026-10", "шилжүүлэг");
  const tosLinks = answer.links["Тосгон"];
  eq([Object.keys(tosLinks).sort(), tosLinks["шилжүүлэг"].id, Object.keys(tosLinks["шилжүүлэг"].days), tosLinks["шилжүүлэг"].days["01"], Object.keys(tosLinks["тооллого"].days)],
     [["тооллого", "шилжүүлэг"], tosBook.file.id, ["01"], tosBook.book.getSheetByName("01").getSheetId(), ["31"]],
     "the feed has each place's month spreadsheet of every kind and the tab of every day with rows");
  ok(Number.isInteger(tosLinks["тооллого"].days["31"]) && tosLinks["тооллого"].days["31"] >= 0, "a tab is its whole-number gid");
  eq([Object.keys(answer.links), feed({ token: "secret-token", mode: "receipts", month: "2026-11" }).links["Соронзон"]], [["Тосгон", "Соронзон", "My home"], null], "every place is there; a month without a spreadsheet has none");
  ok(!JSON.stringify(run(rw, true, "ocrFeedAnswer({}, PropertiesService.getScriptProperties(), new Date())").value).includes("links"), "and without the token no link is given");
}
delete rw.props.OCR_FEED_TOKEN;
const sorHalf = rw.books[CATALOG_ID].sheets["Sor/2026/10/15"], myHalf = rw.books[CATALOG_ID].sheets["My/2026/10/15"];
eq([tos(rw).value(4, AJ), tos(rw).fonts[`4,${AJ}`], bgOf(tos(rw), 4, AJ), sorHalf.value(4, E)], [1, "#1155cc", "#ffe599", 1], "a transfer line goes out of Тосгон and into Соронзон at once, blue on yellow");
eq(tos(rw).notes[`4,${AJ}`], `${rowLink(monthTab(rw, "Тосгон", "2026-10-01", "шилжүүлэг"), 2, `${photoA.id}#1`)} = 1 (2026-10-01)`, "the cell's note links to the line's row in its day tab, with its count and day; no \"Зураг:\" (the owner, 2026-10-26)");
eq([myHalf.value(4, AJ), sorHalf.value(6, T)], [2, 2], "My home -> Соронзон: My home's transfer out (AJ) and Соронзон's transfer in (T) of that half-month");
eq([tos(rw).value(7, AJ), tos(rw).value(6, AJ)], [1, 3], "a line with a name goes in; a person's number stays");
const t1 = monthTab(rw, "Тосгон", "2026-10-01");
const cell = (sheet, row, title) => sheet.value(row, header.indexOf(title) + 1);
eq(cell(t1, 2, "Бичсэн"), "Tos/2026/10!AJ4 ← 1 · Meiling 501l\nSor/2026/10/15!E4 ← 1 · Meiling 501l", "the row says where it was written");
const tosCheck = () => rw.books[`${TOSGON_ID}-check`].sheets["Tos/2026/10"];
eq(cell(t1, 4, "Бичсэн"), "check:Tos/2026/10!AJ6 ← 1 · Haier 400l (гараар 3)\nSor/2026/10/15!E5 ← 1 · Haier 400l", "a line for a hand-typed cell goes into the check copy, and says what was typed");
eq([tosCheck().value(6, AJ), bgOf(tosCheck(), 6, AJ), tos(rw).value(6, AJ), bgOf(tos(rw), 6, AJ), sorHalf.value(5, E)], [1, "#ffe599", 3, "#ff99dd", 1], "the copy's cell holds the receipt, the hand-typed one stays and turns pink (they differ: check it); its other place is written as usual");
eq(tos(rw).notes[`6,${AJ}`], "Баримтаар: 1", "the light red cell's note says what the receipts make");
ok(rw.props[`ocr-check-copy:${TOSGON_ID}`] === `${TOSGON_ID}-check` && rw.books[`${TOSGON_ID}-check`].name === "Тосгон 2026 (check)", "the check copy is made once and kept");
const t1Book = monthBook(rw, "Тосгон", "2026-10-01").book;
ok(t1Book.sheets["_нэрс"].isHidden && t1Book.sheets["_нэрс"].cells.map((r) => r[2]).includes("Haier 400l") && t1.validations.some((v) => v.c === header.indexOf("Бараа") + 1 && v.rule.range && v.rule.range.c === 1 && v.rule.range.h === 1000), "mom chooses a right name from a list (hidden tab, its whole column A), never types it");
eq(tos(rw).value(5, BO), 4, "a count goes into Тосгон's count column");
eq(plain(JSON.parse(rw.props["ocr-day:2026-10-17"]).counts), [{ place: "Тосгон", day: "2026-10-31" }], "the count is kept for mom's daily report");
const reportSheetOf = (world) => world.books[world.root.find("Тайлан").files.find((f) => f.name === "Тооллогын тайлан").id].getSheets()[0];
const countSheet = reportSheetOf(rw);
eq([countSheet.value(1, 1), countSheet.value(1, 3), countSheet.value(1, 4)], ["Өдөр:", "Газар:", "Соронзон"], "the count report: a day and a place to choose");
ok(countSheet.protections[0].unprotected.map((r) => [r.r, r.c]).join() === "1,2,1,4", "only the day and the place may be changed in it");
countSheet.cells[0][1] = "2026-10-31"; countSheet.cells[0][3] = "Тосгон";
run(rw, false, "onCountReportEdit({ range: { getRow: () => 1, getColumn: () => 4 } })");
eq(countSheet.cells[6].slice(0, 4), ["Haier 300l", 4, 5, -1], "choosing a day and a place shows that day's count at once");
ok(String(countSheet.value(2, 1)).startsWith("Тосгон, 2026-10-31 (хаалтын тооллого, Tos/2026/10)") && String(countSheet.value(3, 1)).startsWith("Дүгнэлт: 1 бараанаас 1 нь дутсан"), "with what it is and its conclusion");
const guard = (sheet) => sheet.protections.map((p) => [p.description, p.editors.join(","), p.unprotected.map((r) => [r.c, r.w])]);
eq(plain(guard(t1)), [["OCR хамгаалалт: -", "owner@example.org", []]], "Тосгон's sheets: only the owner account changes anything");
const m3 = monthTab(rw, "My home", "2026-10-03");
eq(plain(guard(m3)), [["OCR хамгаалалт: -", "owner@example.org", []]], "My home's sheets too: the staff addresses give no access any more");
const fileOf = (place, iso) => monthBook(rw, place, iso).file;
eq([plain(fileOf("My home", "2026-10-03").editors), plain(fileOf("Тосгон", "2026-10-01").editors)], [[], []], "nobody but the owner is an editor");
eq(["Соронзон", "My home", "Тосгон", "Зураг уншуулах", "Тайлан"].map((name) => rw.root.find(name).viewers.length), [0, 0, 0, 0, 0], "no folder is shared with the staff");
ok(!["Соронзон", "My home", "Тосгон"].some((place) => rw.root.find(place).files.some((f) => f.name === "Шалгах даалгавар")), "no staff task lists");
ok(!(rw.root.find("Тайлан", "2026-10")?.files || []).some((f) => f.name.startsWith("Тайлан ")) && /өнөөдөр \(2026-10-17\) гаргасангүй/.test(first.logs), "no report on the 17th: it comes out on the 15th and the month's last day");
eq(rw.root.find("Тайлан").files.map((f) => f.name), ["Тооллогын тайлан"], "the count report is still there to look at any day");
rw.today = "2026-10-31";
run(rw, false, "processNewReceipts()");
const reportDoc = () => rw.docs[rw.root.find("Тайлан", "2026-10").files.find((f) => f.name === "Тайлан 2026-10-31").id];
const docText = () => reportDoc().paragraphs.map((p) => p.text).join("\n");
ok(reportDoc().paragraphs[0].text.startsWith("Баримтын тайлан 2026-10-16 – 2026-10-31") && /Уншсан зураг 3,/.test(docText()), "on the month's last day: the report of 16-31, the 17th's photos counted in");
ok(reportDoc().paragraphs[0].heading === "TITLE" &&/Ээжийн шалгах мөр \(нийт 5\)/.test(docText()) && /Тосгон 2026-10-31: 1 бараанаас 1 нь дутсан/.test(docText()) && /Гараар бичсэн тоотой зөрсөн баримт[^\n]*\n[^\n]*Haier 400l: Tos\/2026\/10!AJ6 ← 1 · Haier 400l \(гараар 3\)/.test(docText()), "mom's report is a Doc in Баримт/Тайлан/2026-10, with the last counts linking to the count report");

const cellsNow = () => JSON.stringify([tos(rw).cells, sorHalf.cells, myHalf.cells, tos(rw).notes]);
const afterFirst = cellsNow();
eq([run(rw, false, "processNewReceipts()").error, cellsNow()], [null, afterFirst], "a second run writes nothing twice");

// a second receipt of the same day and product adds to it
put(rw, T1, "D.HEIC", "d", 4, receipt(null, [item("Meiling 501", "Meiling 501l", 2, 1280)]));
run(rw, false, "processNewReceipts()");
eq([tos(rw).value(4, AJ), tos(rw).notes[`4,${AJ}`].split("\n").filter((line) => line.startsWith("https://docs.google.com/spreadsheets/d/")).length, sorHalf.value(4, E)], [3, 2, 3], "two receipts of a day add up in one cell, both in its note");

// mom checks: the doubtful line with its right count, Зөв on the line a person had typed, Буруу on a written one
const setCell = (sheet, row, title, value) => { sheet.cells[row - 1][header.indexOf(title)] = value; };
setCell(t1, 3, "Ээж: Зөв/Буруу", "Зөв"); setCell(t1, 3, "Тоо", 2);
setCell(t1, 4, "Ээж: Зөв/Буруу", "Зөв");
setCell(t1, 2, "Ээж: Зөв/Буруу", "Буруу");
const beforeMarks = cellsNow();
const dryMarks = run(rw, true, "processNewReceipts()");
ok(cellsNow() === beforeMarks && /Tos\/2026\/10!AJ7 ← 2 · xinix 240l/.test(dryMarks.logs), "DRY_RUN only says what it would write");
const momSection = (text) => text.slice(text.indexOf("Ээжийн шалгах мөр"), text.indexOf("Бичиж чадаагүй мөр"));
run(rw, false, "processNewReceipts()");
eq(momSection(dryMarks.logs), momSection(docText()), "DRY_RUN's report shows the checks as they stand now, as the real run then writes them");
eq([tos(rw).value(7, AJ), sorHalf.value(7, E), bgOf(tos(rw), 7, AJ), /= 2 \(2026-10-01\) ✓$/m.test(tos(rw).notes[`7,${AJ}`])], [2, 2, "#d9ead3", true], "mom's corrected count replaces the one read, approved: green, ✓ in the note");
eq([tos(rw).value(4, AJ), /^https:\S+&ocr=\S+\.1 = 2 \(2026-10-01\)$/m.test(tos(rw).notes[`4,${AJ}`]), sorHalf.value(4, E), cell(t1, 2, "Бичсэн")], [2, true, 2, "— workbook-оос хассан (ээж Буруу гэсэн)"], "mom's Буруу takes the line out of both cells by itself; the other receipt stays");
eq([tos(rw).value(6, AJ), sorHalf.value(5, E), tosCheck().value(6, AJ), bgOf(tosCheck(), 6, AJ), bgOf(sorHalf, 5, E)], [3, 1, 1, "#d9ead3", "#d9ead3"], "the typed cell still stays; mom's Зөв turns the line's cells green, in the copy too");

// a written line corrected afterwards: mom picks another product and count
setCell(m3, 2, "Бараа", "Meiling 501l"); setCell(m3, 2, "Тоо", 1); setCell(m3, 2, "Ээж: Зөв/Буруу", "Зөв");
run(rw, false, "processNewReceipts()");
eq([myHalf.value(4, AJ), myHalf.value(5, AJ), sorHalf.value(6, T), sorHalf.value(4, T)], ["", 1, "", 1], "the line moves out of Haier 300l's row into Meiling 501l's, in both shops, with the new count");
eq(cell(m3, 2, "Бичсэн"), "My/2026/10/15!AJ5 ← 1 · Meiling 501l ✓\nSor/2026/10/15!T4 ← 1 · Meiling 501l ✓", "the row says where it is now (✓: approved)");
ok(/засвар орсон 3/.test(docText()), "mom's report counts the corrections (the corrected count, the line taken out, the line moved; not her approvals)");
// a written line made "Хамааралгүй" afterwards leaves both shops; picked again, it comes back
const writtenBefore = cell(m3, 2, "Бичсэн");
setCell(m3, 2, "Бараа", "Хамааралгүй");
run(rw, false, "processNewReceipts()");
eq([myHalf.value(5, AJ), sorHalf.value(4, T), cell(m3, 2, "Бичсэн")], ["", "", "— workbook-оос хассан (Хамааралгүй)"], "\"Хамааралгүй\" takes a written line out of every workbook");
setCell(m3, 2, "Бараа", "Meiling 501l");
run(rw, false, "processNewReceipts()");
eq([myHalf.value(5, AJ), sorHalf.value(4, T), cell(m3, 2, "Бичсэн")], [1, 1, writtenBefore], "... and picking the product again writes it back");
{ // mom's Зөв goes into the workbook at once, by the spreadsheet's On edit trigger (the owner, 2026-10-26)
  const markEdit = (sheet, row, first, last) => run(rw, false, `onOcrViewEdit({ range: { getRow: () => ${row}, getColumn: () => ${first}, getLastColumn: () => ${last}, getValue: () => "", getSheet: () => SpreadsheetApp.openById("${sheet.getParent().getId()}").getSheetByName("${sheet.getName()}") } })`);
  const mark = header.indexOf("Ээж: Зөв/Буруу") + 1;
  setCell(m3, 2, "Тоо", 2);
  eq(markEdit(m3, 2, mark - 1, mark - 1).error, null);
  eq([myHalf.value(5, AJ), sorHalf.value(4, T)], [1, 1], "an edit of another column writes nothing at once");
  eq(markEdit(m3, 2, mark, mark).error, null);
  eq([myHalf.value(5, AJ), sorHalf.value(4, T), cell(m3, 2, "Бичсэн")], [2, 2, writtenBefore.replace(/← 1/g, "← 2")], "Зөв (or a new count under it) is written into both shops at once, without waiting for the hourly run");
  eq(markEdit(m3, 1, mark, mark).error, null, "the header row writes nothing");
  setCell(m3, 2, "Тоо", 1);
  eq(markEdit(m3, 2, 1, header.length).error, null);
  eq([myHalf.value(5, AJ), sorHalf.value(4, T), cell(m3, 2, "Бичсэн")], [1, 1, writtenBefore], "a paste over the whole line counts as a mark too");
}

// mom's own account (the shop account): she may mark her columns, nothing that was read
rw.props.OCR_REVIEWERS = JSON.stringify({ "Соронзон": "so@gmail.com", "My home": ["mh1@gmail.com", "mh2@gmail.com"], "Ээж": "mom@gmail.com" });
fileOf("My home", "2026-10-03").editors.push("mh1@gmail.com"); // marked under the old version
run(rw, false, "processNewReceipts()");
const momCol = header.indexOf("Ээж: Зөв/Буруу") + 1;
const nameCol = header.indexOf("Бараа") + 1;
const momOnly = ["OCR хамгаалалт: ээж", "owner@example.org,mom@gmail.com", []];
eq(plain(guard(t1)), [["OCR хамгаалалт: mom@gmail.com", "owner@example.org", [[1, 1], [nameCol, 2], [momCol, 1]]], momOnly, momOnly, momOnly], "mom's account changes only the view choice, \"Бараа\", \"Тоо\" and her mark");
eq([plain(fileOf("Тосгон", "2026-10-01").editors), plain(fileOf("My home", "2026-10-03").editors), m3.protections[0].description], [["mom@gmail.com"], ["mom@gmail.com"], "OCR хамгаалалт: mom@gmail.com"], "mom's account is the only editor of every place's sheets; a staff member loses the access of the old version");
{ // mom with two accounts (the owner, 2026-10-27: her personal account beside the shop account)
  rw.props.OCR_REVIEWERS = JSON.stringify({ "Ээж": "mom@gmail.com, own@gmail.com" });
  run(rw, false, "processNewReceipts()");
  const both = ["OCR хамгаалалт: ээж", "owner@example.org,mom@gmail.com,own@gmail.com", []];
  eq(plain(guard(t1)), [["OCR хамгаалалт: mom@gmail.com,own@gmail.com", "owner@example.org", [[1, 1], [nameCol, 2], [momCol, 1]]], both, both, both], "both of mom's accounts change her columns, nothing else");
  eq(plain(fileOf("Тосгон", "2026-10-01").editors).sort(), ["mom@gmail.com", "own@gmail.com"], "and both are editors of the sheets");
  rw.props.OCR_REVIEWERS = JSON.stringify({ "Соронзон": "so@gmail.com", "My home": ["mh1@gmail.com", "mh2@gmail.com"], "Ээж": "mom@gmail.com" });
  run(rw, false, "processNewReceipts()");
  eq(plain(fileOf("Тосгон", "2026-10-01").editors), ["mom@gmail.com"], "an account taken out of the list loses the access again");
}

// a day spreadsheet of the first version (next to the photo, old checkbox column): moved into the month spreadsheet
const oldFolder = rw.root.path("Тосгон", "Тосгоноос My home руу", "2026-10", "05");
const oldPhoto = oldFolder.add(new File("OLD.HEIC", "image/heif", "old", 5));
rw.props[`ocr-done:${oldPhoto.id}`] = "read before";
const oldFile = oldFolder.add(new File("OCR 2026-10-05", "application/vnd.google-apps.spreadsheet", "", 5, "me"));
const oldHeader = ["Баримт", "Холбоос", "Хавтас (ээж хийсэн)", "Гарчиг", "Төрөл", "Огноо", "Бараа (уншсан)", "Бараа (Sheet100)", "Тоо", "Нэгж үнэ", "Итгэл", "Хаана орох", "Төлөв", "Шалгах шалтгаан", "Уншсан цаг", "Шалгасан"];
const oldRow = oldHeader.map((title) => ({ "Баримт": "OLD.HEIC", "Холбоос": oldPhoto.getUrl(), "Хавтас (ээж хийсэн)": "Тосгон / Тосгоноос My home руу", "Огноо": new Date(Date.UTC(2026, 9, 4, 15)),
  "Бараа (уншсан)": "Meiling 501", "Бараа (Sheet100)": "Meiling 501l", "Тоо": 1, "Төлөв": "OK", "Шалгасан": true }[title] ?? ""));
rw.books[oldFile.id] = new Book(oldFile.id, oldFile.name, { Sheet1: new Sheet([oldHeader, oldRow]) });
run(rw, false, "processNewReceipts()");
const movedTab = monthTab(rw, "Тосгон", "2026-10-05");
eq([oldFile.trashed, cell(movedTab, 2, "Баримт"), cell(movedTab, 2, "Ээж: Зөв/Буруу"), cell(movedTab, 2, "ID"), movedTab.getName()], [true, "OLD.HEIC", "", `${oldPhoto.id}#1`, "05"], "an old day spreadsheet's rows move into the month spreadsheet with IDs, the old file goes to the trash");
eq(movedTab.images.map((image) => [image.row, image.column]), [[2, header.indexOf("Зураг") + 1]], "with the photo beside its rows");
eq([tos(rw).value(4, AJ + 4), myHalf.value(5, E + 4)], [1, 1], "and its OK line is written");
run(rw, false, "processNewReceipts()");
eq(movedTab.cells.filter((row) => row.slice(0, header.length).some((v) => v !== "" && v !== undefined)).length, 2, "a later run moves nothing twice");

// a month tab of the version with the staff's columns: moved into the new layout by its titles
const OLD_HEADER = ["Баримт", "Холбоос", "Хавтас (ээж хийсэн)", "Гарчиг", "Төрөл", "Огноо", "Бараа (уншсан)", "Бараа (Sheet100)", "Тоо", "Нэгж үнэ", "Итгэл", "Хаана орох", "Төлөв", "Шалгах шалтгаан", "Уншсан цаг",
  "Ажилчин: Зөв/Буруу", "Ажилчин: зөв бараа", "Ажилчин: зөв тоо", "Ажилчин: тайлбар", "Ээж: Зөв/Буруу", "Ээж: тайлбар", "Бичсэн", "ID"];
const old7Row = OLD_HEADER.map((title) => ({ "Баримт": "OLD7.HEIC", "Бараа (Sheet100)": "Haier 300l", "Тоо": 1, "Төлөв": "Шалгах", "Ажилчин: Зөв/Буруу": "Буруу", "Ажилчин: зөв тоо": 3,
  "Ээж: Зөв/Буруу": "Зөв", "Бичсэн": "Tos/2026/10!AJ9 ← 3 · Haier 300l", "ID": "p7#1" }[title] ?? ""));
const old7 = new Sheet([OLD_HEADER, old7Row]);
old7.hidden = [OLD_HEADER.length]; old7.insertImage({}, OLD_HEADER.length + 2, 2); old7.protect().setDescription("OCR хамгаалалт: so@gmail.com");
rw.books.old7 = new Book("old7", "OCR Соронзон 2026-10", { "07": old7 });
const old7Expr = (fn) => `${fn}(SpreadsheetApp.openById("old7").getSheetByName("07"))`;
const readOld = plain(run(rw, true, old7Expr("ocrReadRecords")).value)[0];
eq([readOld["Бараа (Sheet100)"], readOld["Тоо"], readOld["Тоо (уншсан)"], readOld["Ээж: Зөв/Буруу"]], ["Haier 300l", 3, 1, "Зөв"], "an old tab is read by its titles, the staff's count as mom's");
eq([run(rw, true, old7Expr("ocrUpgradeTab")).value, old7.cells[0][15]], [false, "Ажилчин: Зөв/Буруу"], "DRY_RUN leaves the old tab as it is");
eq(run(rw, false, old7Expr("ocrUpgradeTab")).value, true, "the old tab is moved into the new layout");
eq([old7.cells[0].slice(0, header.length), old7.cells[0].every((v, j) => j < header.length || j === header.length || v === "" || v === undefined)], [["Ээж", ...header.slice(1)], true], "with the new titles (the view choice in A1), the old ones beyond them emptied (O1 is the legend's title)");
eq([cell(old7, 2, "Бараа"), cell(old7, 2, "Бараа (тааруулсан)")], ["Haier 300l", "Haier 300l"], "the matched name is both \"Бараа\" and the kept match");
eq([cell(old7, 2, "Баримт"), cell(old7, 2, "Тоо"), cell(old7, 2, "Ээж: Зөв/Буруу"), cell(old7, 2, "Шалгах шалтгаан"), cell(old7, 2, "Бичсэн"), cell(old7, 2, "ID")],
  ["OLD7.HEIC", 3, "Зөв", "тайлбар: ажилчин Буруу", "Tos/2026/10!AJ9 ← 3 · Haier 300l", "p7#1"], "every value under its own title, the staff's count in \"Тоо\", their mark kept in the reason");
eq([old7.images[0].column, header.map((t, i) => i + 1).filter((c) => !old7.hidden.includes(c)).length, old7.protections.length], [header.indexOf("Зураг") + 1, plain(vm.runInContext("OCR_VISIBLE_COLUMNS", pure)).length, 0], "the photo moves beside the new columns, only mom's view is shown, the old protection goes (made anew by the review)");
eq(run(rw, false, old7Expr("ocrUpgradeTab")).value, false, "a tab in the new layout is left alone");

// ---------- check counts: any day but the last of a sheet ----------
const notesInOut = "OCR: x [a#1] = 1 @2026-10-02\nOCR: y [b#1] = 2 @2026-10-12";
eq([pure.ocrTransferUpTo(3, notesInOut, "2026-10-05"), pure.ocrTransferUpTo(5, notesInOut, "2026-10-05"), pure.ocrTransferUpTo(4, "", "2026-10-05")], [1, 3, 4], "a transfer cell up to a day: its dated lines, plus what was typed by hand");
const layout15 = plain(vm.runInContext("LAYOUT", pure));
const shopRow = Array(40).fill("");
shopRow[1] = 5; shopRow[4] = 2; shopRow[7] = 1; shopRow[20] = 1; shopRow[29] = 4;
eq(pure.ocrBalanceAsOf(shopRow, { in: "", out: "" }, layout15, 4, "2026-10-05"), 7, "opening + income of days 1-5 - expense of days 1-5 (day 10's sale left out)");
const spotRecs = [okLine, okLine, reviewLine].map((values) => rec({ ...values, "Хавтас (ээж хийсэн)": "Соронзон / тооллого" }));
eq(plain(pure.ocrCountedLines(spotRecs)), [{ name: "Haier 300l", counted: 3 }], "every counted line with a product counts, one sum per product");

const sw = makeWorld();
sw.props.GEMINI_KEY = "k";
const swSor = sw.books[CATALOG_ID].sheets["Sor/2026/10/15"], swMy = sw.books[CATALOG_ID].sheets["My/2026/10/15"];
// Соронзон's Haier 300l (row 6): opening 5, +2 on the 1st, -1 on the 2nd, -4 on the 10th, transfers in 1 (2nd) + 2 (12th)
Object.entries({ 2: 5, 5: 2, 21: 1, 30: 4, 20: 3 }).forEach(([column, value]) => { swSor.cells[5][column - 1] = value; });
swSor.notes[`6,20`] = notesInOut;
swMy.cells[3][1] = 2; // My home's Haier 300l (row 4): opening 2
put(sw, ["Соронзон", "тооллого", "2026-10", "03"], "S.HEIC", "s", 1, receipt(null, [item("Haier 300", "Haier 300l", 6, null)], { document_type: "toollogo" }));
put(sw, ["My home", "тооллого", "2026-10", "03"], "M.HEIC", "m", 2, receipt(null, [item("Haier 300", "Haier 300l", 3, null)], { document_type: "toollogo" }));
eq(run(sw, false, "processNewReceipts()").error, null, "a run with check counts works");
const spotTab = monthTab(sw, "Соронзон", "2026-10-03");
eq([swSor.value(6, 37), swMy.value(4, 37), cell(spotTab, 2, "Бичсэн")], ["", "", ""], "a check count is not written into the count column");
const spotSheet = reportSheetOf(sw);
spotSheet.cells[0][1] = "2026-10-03"; spotSheet.cells[0][3] = "Соронзон";
run(sw, false, "onCountReportEdit({ range: { getRow: () => 1, getColumn: () => 2 } })");
eq(spotSheet.cells[6].slice(0, 4), ["Haier 300l", 6, 7, -1], "a check count is compared with the books at the end of its day (the 10th's sale and the 12th's transfer left out)");
ok(spotSheet.cells[6][4].startsWith("My home-д 1 ширхэг илүү байна: Соронзоноос My home руу шилжүүлэг бүртгэгдээгүй") && String(spotSheet.value(2, 1)).includes("шалгах тооллого"), "two shops' checks of the same day point at a transfer");
run(sw, false, "onCountReportEdit({ range: { getRow: () => 5, getColumn: () => 2 } })");
eq(spotSheet.cells[6][0], "Haier 300l", "an edit elsewhere in the report changes nothing");

// ---------- counts between the period ends: a block of their own (from the 10/31 sheets on) ----------
{
  const shipped = { run: (expr) => { const c = vm.createContext({}); vm.runInContext(shippedCode + `\nthis.__v = (${expr});`, c); return c.__v; } };
  eq([shipped.run("spotFirstColumn(LAYOUT)"), shipped.run("spotFirstColumn(OCR_TOSGON.layout)")], [49, 77], "the blocks start at AW in the shops' sheets and at BY in Тосгон's");
  const xw = makeWorld();
  xw.props.GEMINI_KEY = "k";
  xw.today = "2026-10-24";
  const book = xw.books[CATALOG_ID], tbook = xw.books[TOSGON_ID];
  const xSor = book.sheets["Sor/2026/10/31"], xTos = tbook.sheets["Tos/2026/10"];
  const F = run(xw, true, "spotFirstColumn(LAYOUT)").value, TF = run(xw, true, "spotFirstColumn(OCR_TOSGON.layout)").value;
  const countColumn = run(xw, true, "LAYOUT.count").value;
  put(xw, ["Соронзон", "тооллого", "2026-10", "23"], "C1.HEIC", "c1", 1, receipt(null, [item("Haier 300", "Haier 300l", 6, null)], { document_type: "toollogo" }));
  put(xw, ["Тосгон", "тооллого", "2026-10", "23"], "T1.HEIC", "t1", 2, receipt(null, [item("Haier 300", "Haier 300l", 4, null)], { document_type: "toollogo" }));
  put(xw, ["Тосгон", "тооллого", "2026-10", "23"], "T2.HEIC", "t2", 3, receipt(null, [item("Haier 300", "Haier 300l", 1, null)], { document_type: "toollogo" }));
  eq(run(xw, false, "processNewReceipts()").error, null, "a run with counts of the 23rd works");
  const sorTab = monthTab(xw, "Соронзон", "2026-10-23");
  eq([xSor.value(3, F), xSor.value(3, F + 1)], ["Тооллого 10/23", "Үлдэгдэл 10/23"], "the 23rd gets a block of its own on Sor/2026/10/31, after the totals");
  ok(String(xSor.value(3, F + 2)).startsWith('=ARRAYFORMULA({"Зөрүү 10/23";'), "its third column is the count minus the balance, as one formula");
  eq([xSor.value(4, F), xSor.value(4, countColumn)], [6, ""], "the count goes into the block, not into the sheet's count column");
  ok(String(xSor.notes[`4,${F}`]).includes("(2026-10-23)") && String(xSor.notes[`4,${F}`]).startsWith("https://docs.google.com/spreadsheets/d/"), "with the usual note");
  ok(cell(sorTab, 2, "Бичсэн").startsWith(`Sor/2026/10/31!${run(xw, true, `columnLetter(${F})`).value}4`), "\"Бичсэн\" names the cell");
  eq([xSor.value(4, F + 1), xSor.value(5, F + 1), xSor.value(6, F + 1)], [2, 1, ""], "the balance column holds each product's books at the end of the 23rd (no balance on a heading row)");
  ok((xSor.formats || []).some((f) => f.c === F && f.w === 3 && f.format === "0"), "the block's numbers are plain numbers (not the ₮ of the totals beside it)");
  const guard = xSor.protections.find((p) => p.description === "OCR хамгаалалт: дундын тооллого");
  ok(guard && guard.range.c === F && guard.range.w === 3 && guard.editors.includes("owner@example.org") && guard.editors.includes("shop-account@example.com") && !guard.editors.includes("someone@else"), "the block is changed by the owner and the shop account only");
  const list = book.sheets["_харах"];
  const listRow = list.cells.find((row) => row[0] === "Sor/2026/10/31");
  ok(list.isHidden && listRow && listRow.includes("Тооллого 10/23") && listRow.includes("Бүгд") && !listRow.includes("1–15"), "A1's list of the sheet gets the day (a half-month sheet has no halves)");
  const rule = xSor.validations.filter((v) => v.r === 1 && v.c === 1).pop();
  ok(rule && rule.rule.range && rule.rule.range.r === list.cells.indexOf(listRow) + 1, "A1 chooses from that list");
  eq([xTos.value(3, TF), xTos.value(5, TF)], ["Тооллого 10/23", 5], "Тосгон's two pages of the 23rd add up in its block of Tos/2026/10");
  const tosRow = tbook.sheets["_харах"].cells.find((row) => row[0] === "Tos/2026/10");
  ok(tosRow.includes("1–15") && tosRow.includes("Тооллого 10/23"), "a month sheet's list has its halves too");
  // a second run changes nothing and makes no second block
  run(xw, false, "processNewReceipts()");
  eq([xSor.value(3, F + 3), xSor.value(4, F), xSor.protections.filter((p) => p.description === "OCR хамгаалалт: дундын тооллого").length], ["", 6, 1], "a second run makes no second block");
  // the books change: the balance follows within the hour
  xSor.cells[3][1] = 3;
  run(xw, false, "processNewReceipts()");
  eq(xSor.value(4, F + 1), 3, "the balance column is worked out again every run");
  // mom says the line is wrong: it comes out of the block
  sorTab.cells[1][header.indexOf("Ээж: Зөв/Буруу")] = "Буруу";
  run(xw, false, "processNewReceipts()");
  eq([xSor.value(4, F), xSor.notes[`4,${F}`] || ""], ["", ""], "\"Буруу\" takes the count out of the block");
}

// ---------- the likeliest products of a line nothing was matched to ----------
const suggestEntries = [{ name: "Haier 300l", price: 760000 }, { name: "Haier 200l", price: 570000 }, { name: "Haier 103l", price: 520000 }, { name: "Meiling 501l", price: 1200000 }, { name: "Сууръ", price: 10000 }, { name: "Smart 6.5l", price: 280000 }];
const suggest = (read, price) => plain(pure.ocrSuggestNames(read, price, suggestEntries));
eq(suggest("HAIER 300", 760), ["Haier 300l", "Haier 200l", "Haier 103l"], "the same brand and number first, then the brand's others");
eq(suggest("Smart 6.5", 280)[0], "Smart 6.5l", "a size in the name counts");
eq([suggest("", 100), suggest("zzzz qqq", null)], [[], []], "nothing read, or nothing alike: no suggestion");
const pots = [{ name: "1 хүний тогоо", price: 65000 }, { name: "солонгос будаа агш 5Л", price: 125000 }, { name: "Гахай 5l", price: 70000 }, { name: "Haier 300l", price: 760000 }];
eq([plain(pure.ocrSuggestNames("16 тохоо", 240, pots))[0], plain(pure.ocrSuggestNames("Co 61A 5л", 130, pots))[0]], ["1 хүний тогоо", "солонгос будаа агш 5Л"], "any word of a name counts (\"тохоо\" ~ \"тогоо\"), Cyrillic and Latin look-alikes are one (5л = 5l), and the price helps");
const unmatched = pure.ocrRows(receipt("2026-10-01", [item("HAIER 300", null, 1, 760)]), { name: "IMG_S", url: "u", id: "s-id" },
  pure.ocrFolderMeaning(["Соронзон", "зарлага"]), { date: "2026-10-01", reasons: [] }, new Set(suggestEntries.map((e) => e.name)), "2026-10-17 10:00", suggestEntries);
ok(String(unmatched[0][header.indexOf("Шалгах шалтгаан")]).endsWith("Магадгүй: Haier 300l, Haier 200l, Haier 103l") && unmatched[0][header.indexOf("Бараа")] === "Хамааралгүй", "a line Gemini found no product for is Хамааралгүй and says what it likely is");
// a tab read before 2026-10-25, still in its old layout: its lines with no product get their suggestion
const oldLine = { "Бараа (уншсан)": "haier 300", "Нэгж үнэ": 760, "Тоо": 1, "Бараа": "", "Төлөв": "Шалгах", "Шалгах шалтгаан": "нэр жагсаалтад таараагүй" };
eq(pure.ocrSuggestedReason(oldLine, suggestEntries), "нэр жагсаалтад таараагүй; Магадгүй: Haier 300l, Haier 200l, Haier 103l", "a line read before gets its suggestion at the review");
eq([pure.ocrSuggestedReason({ ...oldLine, "Бараа": "Haier 300l" }, suggestEntries), pure.ocrSuggestedReason({ ...oldLine, "Шалгах шалтгаан": "Магадгүй: X" }, suggestEntries)], [null, null], "not once mom has picked one, nor twice");

const lastLayout = plain(pure.ocrUpgradeRecord({ "Бараа": "Haier 400l", "Бараа (тааруулсан)": "Haier 300l", "Тоо": 1, "Ээж: зөв тоо": "", "Ээж: тайлбар": "", "Тоо (уншсан)": "" }));
eq([lastLayout["Бараа"], lastLayout["Тоо"], lastLayout["Тоо (уншсан)"], lastLayout["Шалгах шалтгаан"]], ["Haier 400l", 1, 1, ""], "a tab of the layout of 2026-10-19: mom's chosen product stays, the count read is kept");
const lastTyped = plain(pure.ocrUpgradeRecord({ "Бараа": "Haier 300l", "Бараа (тааруулсан)": "Haier 300l", "Тоо": 1, "Ээж: зөв тоо": 3, "Ээж: тайлбар": "хайрцагтай", "Тоо (уншсан)": "" }));
eq([lastTyped["Тоо"], lastTyped["Тоо (уншсан)"], lastTyped["Шалгах шалтгаан"], pure.ocrIsChosenByMom(lastTyped)], [3, 1, "тайлбар: хайрцагтай", true], "... and mom's count becomes \"Тоо\": her change, so her approval");

// ---------- what mom sees of a line nothing was matched to ----------
const hw = makeWorld();
hw.props.GEMINI_KEY = "k";
put(hw, DAY("My home", "зарлага", "05"), "H.HEIC", "h", 1, receipt(null, [item("HAIER 300", null, 1, 1200), item("Meiling 501", "Meiling 501l", 2, 130.0)], { document_type: "zarlaga" }));
eq(run(hw, false, "processNewReceipts()").error, null, "a run with a line nothing was matched to works");
const h5 = monthTab(hw, "My home", "2026-10-05");
const nameColumn = header.indexOf("Бараа") + 1;
eq([cell(h5, 2, "Бараа"), String(cell(h5, 2, "Шалгах шалтгаан")).includes("Магадгүй: Haier 300l"), h5.notes[`2,${nameColumn}`] || "", h5.notes[`3,${nameColumn}`] || ""], ["Хамааралгүй", true, "", ""],
   "a line Gemini found no product for: Хамааралгүй, the likeliest products in its reason; no note (\"Бараа\" is never empty)");
ok(["Тоо", "Тоо (уншсан)"].every((title) => h5.formats.some((f) => f.c <= header.indexOf(title) + 1 && header.indexOf(title) + 1 < f.c + f.w)), "the number columns get number formats (a moved column keeps no date format)");
const namesTab = monthBook(hw, "My home", "2026-10-05").book.sheets["_нэрс"];
ok(namesTab.cells.some((row) => row[2] === "Haier 300l") && !namesTab.cells.some((row) => row[2] === "Хамааралгүй"), "the script writes every source's names in C");
const listFormula = namesTab.cells[0][0];
ok(listFormula.startsWith('=LET(names,ARRAYFORMULA(TRIM({C1:C;IFERROR(QUERY(IMPORTRANGE("') && listFormula.includes(`","Бараа!A4:B"),"select Col1 where Col2 is not null",0),"")})),{"Хамааралгүй";UNIQUE(FILTER(names,names<>"",names<>"Хамааралгүй"))})`),
   "mom's list (A) is Хамааралгүй, then C and \"Бараа\" read live: a product added there is in her list within seconds (the owner, 2026-10-25)");
ok((hw.imports || []).some((url) => url.includes(`/d/${monthBook(hw, "My home", "2026-10-05").file.id}/externaldata/addimportrangepermissions?donorDocId=${CATALOG_ID}`)) && hw.props[`ocr-import-ok:${monthBook(hw, "My home", "2026-10-05").file.id}`] === CATALOG_ID,
   "the import is allowed once per month spreadsheet, as the owner's click would");
{
  const asked = hw.imports.length;
  hw.props["ocr-book-setup:" + monthBook(hw, "My home", "2026-10-05").file.id] = "old";
  run(hw, false, "processNewReceipts()");
  eq(hw.imports.length, asked, "not asked again once allowed");
  hw.importRefused = true;
  delete hw.props[`ocr-import-ok:${monthBook(hw, "My home", "2026-10-05").file.id}`];
  hw.props["ocr-book-setup:" + monthBook(hw, "My home", "2026-10-05").file.id] = "old";
  const refused = run(hw, false, "processNewReceipts()");
  ok(refused.error === null && /Allow access/.test(refused.logs) && namesTab.cells.some((row) => row[2] === "Haier 300l"), "a refused import says what to click; the list the script wrote still stands");
}

// ---------- Тосгон's expense before OCR_POST_FROM, from the shops' typed income ----------
const bw = makeWorld();
bw.postFrom = "2026-10-20";
const shopTab = (world, name) => world.books[CATALOG_ID].sheets[name];
const setShop = (name, row, day, qty, half = 0) => { shopTab(bw, name).cells[row - 1][E - 1 + day - (half ? 16 : 1)] = qty; };
setShop("Sor/2026/10/15", 6, 3, 2); // Haier 300l, 3 October
setShop("My/2026/10/15", 4, 3, 1); // Haier 300l, 3 October
setShop("My/2026/10/15", 5, 5, 1); // Meiling 501l, 5 October: Тосгон's cell typed by hand already
shopTab(bw, "My/2026/10/15").cells.push(["Залгуур 45"]);
setShop("My/2026/10/15", 6, 2, 3); // a small product, not on Тосгон's sheet
setShop("Sor/2026/10/31", 5, 19, 1, 1); // Haier 400l, 19 October
setShop("Sor/2026/10/31", 5, 21, 1, 1); // ... and 21 October: from OCR_POST_FROM on the OCR writes it
tos(bw).cells[3][AJ + 4 - 1] = 4; // Meiling 501l, 5 October, typed by hand
const expenseOf = (row, day) => tos(bw).value(row, AJ + day - 1);
const preview = run(bw, false, "previewTosgonBackfill()");
ok(expenseOf(5, 3) === "" && /Tos\/2026\/10: 2 нүд, 4 ширхэг бичнэ \(урьдчилж харав/.test(preview.logs), "the preview only says what it would write");
const backfill = run(bw, false, "runTosgonBackfill()");
// the link to a shop's income cell: the tab, the row, the day's column (2026-10-26)
const shopLink = (name, row, day, half = 0) =>
  `https://docs.google.com/spreadsheets/d/${CATALOG_ID}/edit#gid=${shopTab(bw, name).getSheetId()}&range=${run(bw, true, `columnLetter(LAYOUT.incomeFirst + ${day - (half ? 16 : 1)})`).value}${row}`;
eq([expenseOf(5, 3), tos(bw).notes[`5,${AJ + 2}`], tos(bw).fonts[`5,${AJ + 2}`]],
   [3, `Дэлгүүрийн орлогоос:\n${shopLink("My/2026/10/15", 4, 3)} = 1 (My home)\n${shopLink("Sor/2026/10/15", 6, 3)} = 2 (Соронзон)`, "#8e24aa"],
   "a big product's shop income of a day becomes Тосгон's expense of that day, purple, its note linking to the shops' cells");
eq([bgOf(tos(bw), 5, AJ + 2), bgOf(tos(bw), 4, AJ + 4)], ["#ffe599", null], "a filled cell is yellow, like any number the system is sure of (the owner, 2026-10-28); a hand-typed one keeps its colour");
eq([expenseOf(6, 19), expenseOf(6, 21)], [1, ""], "the days from OCR_POST_FROM on are left to the OCR");
eq(expenseOf(4, 5), 4, "a cell typed by hand stays as it is");
ok(/1 нүдийг хөндсөнгүй: Meiling 501l 5-нд 4 \(дэлгүүрт 1\)/.test(backfill.logs) && /жижиг барааны 1 мөр алгассан/.test(backfill.logs), "... and is reported, like the small goods left out");
ok(/Tos\/2026\/9: "Tos\/2026\/9" sheet алга/.test(backfill.logs), "a month without a Тосгон sheet is said so");
const backfillAgain = run(bw, false, "runTosgonBackfill()");
ok(expenseOf(5, 3) === 3 && /Tos\/2026\/10: 2 нүд, 4 ширхэг бичлээ/.test(backfillAgain.logs), "run again it writes the same, nothing twice");
setShop("Sor/2026/10/15", 6, 3, "");
run(bw, false, "runTosgonBackfill()");
eq([expenseOf(5, 3), tos(bw).notes[`5,${AJ + 2}`]], [1, `Дэлгүүрийн орлогоос:\n${shopLink("My/2026/10/15", 4, 3)} = 1 (My home)`], "a shop's income taken back is taken out of Тосгон's expense too");
setShop("Sor/2026/10/31", 5, 19, "", 1);
run(bw, false, "runTosgonBackfill()");
eq([expenseOf(6, 19), bgOf(tos(bw), 6, AJ + 18)], ["", null], "a cell emptied so loses its yellow too");
setShop("Sor/2026/10/31", 5, 19, 1, 1);
run(bw, false, "runTosgonBackfill()");
eq([expenseOf(6, 19), bgOf(tos(bw), 6, AJ + 18)], [1, "#ffe599"], "and gets it back with its number");
{ // the cells filled before 2026-10-28 (purple, no colour) are coloured once, their numbers left alone
  tos(bw).getRange(5, AJ + 2).setBackground(null);
  tos(bw).getRange(6, AJ + 18).setBackground("#d9ead3"); // a colour someone gave it is kept
  const dry = run(bw, true, "colorTosgonBackfillCells()");
  eq([dry.error, bgOf(tos(bw), 5, AJ + 2), /Tos\/2026\/10: 1 нүд шар болно \(туршилт\)/.test(dry.logs)], [null, null, true], "DRY_RUN only counts them");
  const done = run(bw, false, "colorTosgonBackfillCells()");
  eq([done.error, bgOf(tos(bw), 5, AJ + 2), bgOf(tos(bw), 6, AJ + 18), expenseOf(5, 3), bgOf(tos(bw), 4, AJ + 4)], [null, "#ffe599", "#d9ead3", 1, null],
     "a backfill cell without a colour turns yellow; one with another colour, and a hand-typed one, stay; no number changes");
  tos(bw).getRange(6, AJ + 18).setBackground("#ffe599");
}
{
  // a note of before 2026-10-26 gets the links; the number is never changed
  tos(bw).notes[`5,${AJ + 2}`] = "Дэлгүүрийн орлогоос: My home 1";
  tos(bw).cells[5][AJ + 6 - 1] = 7; tos(bw).notes[`6,${AJ + 6}`] = "Дэлгүүрийн орлогоос: Соронзон 7"; // no shop income behind it any more
  tos(bw).notes[`7,${AJ + 2}`] = "Дэлгүүрийн орлогоос: My home 1"; // an empty copy of a new row
  const dry = run(bw, true, "linkTosgonBackfillNotes()");
  ok(tos(bw).notes[`5,${AJ + 2}`] === "Дэлгүүрийн орлогоос: My home 1" && /Tos\/2026\/10: 1 нүдний тэмдэглэл холбоостой болно \(туршилт\)/.test(dry.logs), "DRY_RUN only says what it would link");
  const linked = run(bw, false, "linkTosgonBackfillNotes()");
  eq([expenseOf(5, 3), tos(bw).notes[`5,${AJ + 2}`]], [1, `Дэлгүүрийн орлогоос:\n${shopLink("My/2026/10/15", 4, 3)} = 1 (My home)`], "an old note gets the link to the shop's cell");
  eq([expenseOf(6, 7), tos(bw).notes[`6,${AJ + 6}`]], [7, "Дэлгүүрийн орлогоос: Соронзон 7"], "a number the shops' income no longer makes is left as it is ...");
  ok(/таарахгүй 1 нүдийг хөндсөнгүй: [A-Z]+6 = 7 \(дэлгүүрт 0\)/.test(linked.logs) && /хоосон нүдэнд үлдсэн 1 тэмдэглэл/.test(linked.logs), "... and listed, the empty copy counted for clearCopiedNotes");
  ok(/Tos\/2026\/10: 0 нүдний/.test(run(bw, false, "linkTosgonBackfillNotes()").logs), "a second run links nothing");
  tos(bw).cells[5][AJ + 6 - 1] = ""; delete tos(bw).notes[`6,${AJ + 6}`]; delete tos(bw).notes[`7,${AJ + 2}`];
}
setShop("My/2026/10/15", 4, 3, "");
run(bw, false, "runTosgonBackfill()");
eq([expenseOf(5, 3), tos(bw).notes[`5,${AJ + 2}`] || "", tos(bw).fonts[`5,${AJ + 2}`] || null], ["", "", null], "... and its cell is emptied when nothing is left");

// up to OCR_TOSGON_HAND_UNTIL Тосгон typed its own expense, often on another day: only what is left over
const hb = makeWorld();
hb.postFrom = "2026-10-20"; hb.handUntil = "2026-10-10";
const setHb = (name, row, day, qty, half = 0) => { shopTab(hb, name).cells[row - 1][E - 1 + day - (half ? 16 : 1)] = qty; };
setHb("Sor/2026/10/15", 6, 3, 2); setHb("My/2026/10/15", 4, 3, 1); // Haier 300l: 3 in the shops on 3 October
setHb("Sor/2026/10/31", 5, 19, 1, 1); // Haier 400l on 19 October: after Тосгон stopped typing
tos(hb).cells[4][AJ + 7 - 2] = 2; // Тосгон typed Haier 300l on 7 October: 2 of the 3
const hbOut = (row, day) => tos(hb).value(row, AJ + day - 1);
const trimRun = run(hb, false, "runTosgonBackfill()");
eq([hbOut(5, 3), hbOut(5, 7), hbOut(6, 19)], [1, 2, 1], "a hand-typed day leaves only the rest (3 - 2) to the backfill; days after Тосгон stopped typing are taken whole");
ok(/2-г хассан/.test(tos(hb).notes[`5,${AJ + 2}`]) && /давхардах 2 ширхгийг оруулаагүй/.test(trimRun.logs), "the note and the log say how much was left out");
tos(hb).cells[4][AJ + 7 - 2] = 3; // Тосгон had typed all 3
run(hb, false, "runTosgonBackfill()");
eq([hbOut(5, 3), tos(hb).notes[`5,${AJ + 2}`] || ""], ["", ""], "run again: when the hand-typed numbers cover it all, our cell is emptied");
eq(plain(run(hb, true, "[ocrHandLastDay(9), ocrHandLastDay(10), ocrHandLastDay(11)]").value), [30, 10, 0], "the hand-typed days of each month");

// ---------- every line goes in; a hand-typed cell gets the check copy ----------
const cw = makeWorld();
cw.props.GEMINI_KEY = "k";
const myHalfCw = cw.books[CATALOG_ID].sheets["My/2026/10/15"];
const W3 = 23; // My home's expense column of 3 October
myHalfCw.cells[3][W3 - 1] = 2; // Haier 300l, typed by hand: the receipt makes the same
myHalfCw.cells[4][W3 - 1] = 3; // Meiling 501l, typed by hand: the receipt makes 1
put(cw, DAY("My home", "зарлага", "03"), "K.HEIC", "k", 1, receipt(null, [item("Haier 300", "Haier 300l", 2, 1200, 0.7), item("Meiling 501", "Meiling 501l", 1, 1280)], { document_type: "zarlaga" }));
put(cw, ["My home", "зарлага", "2025-12", "05"], "OLD25.HEIC", "o", 2, receipt(null, [item("Haier 300", "Haier 300l", 1, 1200)], { document_type: "zarlaga" }));
put(cw, DAY("My home", "зарлага", "03"), "L.HEIC", "l", 3, receipt(null, [item("Haier 400", "Haier 400l", 1, 1500)], { document_type: "zarlaga" })); // a name the half-month sheet has no row for
myHalfCw.cells.push(["Panasonica 195l"]);
myHalfCw.cells[5][W3 - 1] = 1; // typed by hand, no receipt for it
const dryCheck = run(cw, true, "processNewReceipts()");
ok(!cw.books[`${CATALOG_ID}-check`] && dryCheck.error === null, "DRY_RUN makes no check copy");
eq(run(cw, false, "processNewReceipts()").error, null, "a run with hand-typed cells works");
const k3 = monthTab(cw, "My home", "2026-10-03");
const myCheck = () => cw.books[`${CATALOG_ID}-check`].sheets["My/2026/10/15"];
eq([cell(k3, 2, "Бичсэн"), cell(k3, 3, "Бичсэн")], ["check:My/2026/10/15!W4 ← 2 · Haier 300l (гараар бичсэнтэй таарсан)", "check:My/2026/10/15!W5 ← 1 · Meiling 501l (гараар 3)"], "each line says whether the receipts make what was typed");
eq([myHalfCw.value(4, W3), bgOf(myHalfCw, 4, W3), myHalfCw.value(5, W3), bgOf(myHalfCw, 5, W3)], [2, "#d9ead3", 3, "#ff99dd"], "the hand-typed cells stay: green when the receipts make the same, pink when they do not (the owner, 2026-10-28: three colours only)");
eq([myHalfCw.notes[`4,${W3}`] || "", myHalfCw.notes[`5,${W3}`]], ["", "Баримтаар: 1"], "only the light red one gets a note");
{
  const table = monthBook(cw, "My home", "2026-10-03").book.getSheetByName("Тулгалт");
  eq(plain(table.cells.slice(0, 5)), [
    ["Өдөр", "Бараа", "Ээж бичсэн", "Баримтаар", "Төлөв", "Нүд", "Код"],
    ["03", "Haier 300l", 2, 2, "таарсан", "My/2026/10/15!W4", "match"],
    ["03", "Meiling 501l", 3, 1, "зөрсөн", "My/2026/10/15!W5", "diff"],
    ["03", "Haier 400l", "", 1, "нэр хуудсанд алга", "My/2026/10/15", "noname"],
    ["03", "Panasonica 195l", 1, "", "баримт алга", "My/2026/10/15!W6", "noreceipt"],
  ], "the month spreadsheet's Тулгалт tab lists every day's comparison, mom's rows without a receipt too");
  ok(!table.isHidden, "and anyone opening the spreadsheet sees it");
}
{
  const writesBefore = myHalfCw.backgrounds.length, notesBefore = JSON.stringify(myHalfCw.notes);
  run(cw, false, "processNewReceipts()");
  eq([myHalfCw.backgrounds.length, JSON.stringify(myHalfCw.notes)], [writesBefore, notesBefore], "a run with nothing changed writes no hand-typed cell again");
  const k3Validations = k3.validationWrites;
  run(cw, false, "processNewReceipts()");
  eq(k3.validationWrites, k3Validations, "nor makes the month spreadsheet's name list and rules again");
  cw.props["ocr-book-setup:" + monthBook(cw, "My home", "2026-10-03").file.id] = "old";
  run(cw, false, "processNewReceipts()");
  ok(k3.validationWrites > k3Validations, "they are made again once what they come from changed");
  const cut = plain(run(cw, false, `(() => {
    const props = PropertiesService.getScriptProperties();
    const root = DriveApp.getFolderById(OCR_ROOT_FOLDER_ID);
    const entry = findOcrBooks(root, null).find((found) => found.place === "My home" && found.month === "2026-10");
    const key = ocrCalendarKey(entry.place, entry.month, entry.kind);
    const kept = props.getProperty(key);
    props.deleteProperty(key);
    const result = reviewOcrBook(entry, new Map(), { props, root, mom: "", catalog: { entries: [] }, touched: new Map() }, 0);
    const after = props.getProperty(key);
    props.setProperty(key, kept);
    return [result.isFinished, result.written, after];
  })()`).value);
  eq(cut, [false, 0, null], "past the deadline a month spreadsheet stops before its next day and keeps no half calendar");
  const lateSetup = plain(run(cw, false, `(() => {
    const props = PropertiesService.getScriptProperties();
    const root = DriveApp.getFolderById(OCR_ROOT_FOLDER_ID);
    const entry = findOcrBooks(root, null).find((found) => found.place === "My home" && found.month === "2026-10");
    const key = OCR_BOOK_SETUP_PREFIX + entry.file.getId();
    props.setProperty(key, "old");
    const result = reviewOcrBook(entry, new Map(), { props, root, mom: "", catalog: { entries: [] }, touched: new Map() }, 0);
    return [result.isFinished, props.getProperty(key)];
  })()`).value);
  eq(lateSetup, [false, "old"], "past the deadline a month spreadsheet is not set up either: the setup waits for the next run");
  const inTime = plain(run(cw, true, `(() => {
    const done = [];
    return [ocrEachInTime([1, 2], (item) => done.push(item), Infinity), ocrEachInTime([3], (item) => done.push(item), 0), done];
  })()`).value);
  eq(inTime, [true, false, [1, 2]], "a step is taken only while there is time, and the caller learns whether all were taken");
  const halfRun = run(cw, false, `(() => {
    const props = PropertiesService.getScriptProperties();
    const root = DriveApp.getFolderById(OCR_ROOT_FOLDER_ID);
    const entry = findOcrBooks(root, null).find((found) => found.place === "My home" && found.month === "2026-10");
    const book = SpreadsheetApp.openById(entry.file.getId());
    ["27", "28"].forEach((day) => ocrDayTab(book, day)); // three day tabs
    const key = OCR_BOOK_SETUP_PREFIX + entry.file.getId();
    const progressKey = OCR_BOOK_SETUP_PROGRESS_PREFIX + entry.file.getId();
    const context = { props, root, mom: "", catalog: { entries: [] }, touched: new Map() };
    const fitted = [];
    const realFit = ocrFitColumns;
    ocrFitColumns = (tab) => { fitted.push(tab.getName()); return realFit(tab); };
    props.setProperty(key, "old");
    const realNow = Date.now;
    let checks = 0;
    Date.now = () => (checks++ < 1 ? 0 : 2); // the first tab in time, the next one late
    let first;
    try {
      first = ocrSetUpBook(book, entry, context, 1);
    } finally {
      Date.now = realNow;
    }
    const progress = props.getProperty(progressKey);
    const cutState = [first, props.getProperty(key), fitted.length, progress ? JSON.parse(progress).tabs.length : "none"];
    fitted.length = 0;
    const second = ocrSetUpBook(book, entry, context, Infinity);
    ocrFitColumns = realFit;
    const tabCount = ocrDayTabs(book).length;
    return [...cutState, second, props.getProperty(key) !== "old", props.getProperty(progressKey), fitted.length === tabCount - 1, tabCount > 1];
  })()`);
  if (halfRun.error) throw halfRun.error;
  const halfSetup = plain(halfRun.value);
  eq(halfSetup, [false, "old", 1, 1, true, true, null, true, true], "a setup cut by the time keeps the tabs it did; the next run sets up only the rest and then marks the spreadsheet set up");
  const order = plain(run(cw, true, `(() => {
    const book = (id, at) => ({ file: { getId: () => id, getLastUpdated: () => new Date(at) } });
    const ids = (cutId) => ocrReviewOrder([book("a", 3), book("b", 1), book("cut", 5)], cutId).map((entry) => entry.file.getId());
    return [ids(null), ids("cut")];
  })()`).value);
  eq(order, [["b", "a", "cut"], ["cut", "b", "a"]], "the month spreadsheet a pass stopped in goes first the next time, the others longest unchanged first");
}
cw.props.OCR_FEED_TOKEN = "t";
const checksOf = () => plain(run(cw, true, `ocrFeedAnswer({ token: "t", mode: "receipts", month: "2026-10" }, PropertiesService.getScriptProperties(), new Date())`).value).checks;
const my3 = { sheet: "My/2026/10/15" };
eq(checksOf()["My home"]["03"], [
  { name: "Haier 300l", hand: 2, receipt: 2, status: "match", ...my3, cell: "W4" },
  { name: "Meiling 501l", hand: 3, receipt: 1, status: "diff", ...my3, cell: "W5" },
  { name: "Haier 400l", hand: null, receipt: 1, status: "noname", ...my3, cell: "" },
  { name: "Panasonica 195l", hand: 1, receipt: null, status: "noreceipt", ...my3, cell: "W6" },
], "the feed compares, per day, mom's hand-typed cells with the receipts: the same, another number, a name the sheet has not, a number without a receipt");
{
  const myBook = monthBook(cw, "My home", "2026-10-03").book;
  myBook.insertSheet("_тулгалт").cells = [[JSON.stringify({ day: "03", checks: [] })]]; // the hidden tab of the morning's version
  myBook.deleteSheet(myBook.getSheetByName("Тулгалт"));
  eq(Object.keys(checksOf()["My home"]), ["03"], "the feed still reads the hidden tab of the version before");
  run(cw, false, "processNewReceipts()");
  eq([myBook.getSheetByName("_тулгалт"), checksOf()["My home"]["03"].length], [null, 4], "the next review takes it away and the Тулгалт tab is read again");
}
{ // the feed reads each spreadsheet's Тулгалт kept in Script properties (the owner, 2026-10-29: a doGet opening
  // the month's 8 spreadsheets took 14 s, the site waits 12 and showed the answer of before: days mom checked stayed blue)
  const myId = monthBook(cw, "My home", "2026-10-03").book.getId();
  const key = `ocr-checks:${myId}`;
  ok(cw.props[key] !== undefined, "a review keeps the spreadsheet's Тулгалт beside it");
  const closed = run(cw, true, `(() => { const real = SpreadsheetApp.openById; SpreadsheetApp.openById = (id) => { if (id === "${myId}") throw new Error("opened"); return real(id); }; try { return ocrReadChecks("${myId}"); } finally { SpreadsheetApp.openById = real; } })()`);
  eq([closed.error, plain(closed.value)["03"].length], [null, 4], "the feed reads it without opening the spreadsheet");
  delete cw.props[key];
  eq([plain(run(cw, false, `ocrReadChecks("${myId}")`).value)["03"].length, cw.props[key] !== undefined], [4, true], "none kept yet: read from the tab once, then kept");
}

eq([checksOf()["Соронзон"], Object.keys(checksOf())], [{}, ["Тосгон", "Соронзон", "My home"]], "a place without photos that month has no checks");
ok(!("checks" in plain(run(cw, true, `ocrFeedAnswer({}, PropertiesService.getScriptProperties(), new Date())`).value)), "and without the token none are given");
const longName = "x".repeat(120);
eq(plain(run(cw, true, `(() => { const kept = ocrLimitChecks([{ name: ${JSON.stringify(longName)} }, ...Array.from({ length: 250 }, () => ({ name: "a" }))]); return [kept[0].name.length, kept.length]; })()`).value), [80, 200], "a name is cut at 80 characters, a day at 200 rows");
eq([myCheck().value(4, W3), myCheck().value(5, W3), bgOf(myCheck(), 5, W3)], [2, 1, "#ffe599"], "the check copy holds what the receipts say");
setCell(k3, 2, "Ээж: Зөв/Буруу", "Зөв");
run(cw, false, "processNewReceipts()");
eq([bgOf(myHalfCw, 4, W3), bgOf(myCheck(), 4, W3), cell(k3, 2, "Бичсэн")], ["#d9ead3", "#d9ead3", "check:My/2026/10/15!W4 ← 2 · Haier 300l ✓ (гараар бичсэнтэй таарсан)"], "mom's Зөв turns the copy green; her own cell stays green");
myHalfCw.cells[4][W3 - 1] = 1; // mom puts the number right
run(cw, false, "processNewReceipts()");
eq([myHalfCw.value(5, W3), bgOf(myHalfCw, 5, W3), myHalfCw.notes[`5,${W3}`]], [1, "#d9ead3", ""], "once mom's number is the receipts' it turns green and the note line goes");
myHalfCw.getRange(5, W3).setBackground("#c9daf8"); // the light blue before 2026-10-28
run(cw, false, "processNewReceipts()");
eq(bgOf(myHalfCw, 5, W3), "#d9ead3", "a cell still in the old light blue turns green");
myHalfCw.cells[4][W3 - 1] = 4;
myHalfCw.notes[`5,${W3}`] = "ээж: дараа шалгана";
run(cw, false, "processNewReceipts()");
eq([bgOf(myHalfCw, 5, W3), myHalfCw.notes[`5,${W3}`]], ["#ff99dd", "ээж: дараа шалгана\nБаримтаар: 1"], "mom's own note lines stay next to the receipts' number");
["#f4c7c3", "#e6b8af"].forEach((old) => { // the light reds before 2026-10-23 and before 2026-10-28
  myHalfCw.getRange(5, W3).setBackground(old);
  run(cw, false, "processNewReceipts()");
  eq(bgOf(myHalfCw, 5, W3), "#ff99dd", `a cell still in an old light red (${old}) turns pink`);
});
setCell(k3, 3, "Ээж: Зөв/Буруу", "Буруу");
run(cw, false, "processNewReceipts()");
eq([myHalfCw.value(5, W3), bgOf(myHalfCw, 5, W3), myHalfCw.notes[`5,${W3}`]], [4, null, "ээж: дараа шалгана"], "with no receipt left her cell has no colour and no receipt line");
{ // a row put in above in the original only: the check copy's rows are no longer the original's
  myHalfCw.cells.splice(3, 0, ["Aucma 1"]);
  myHalfCw.notes = Object.fromEntries(Object.entries(myHalfCw.notes).map(([key, note]) => { const [r, c] = key.split(",").map(Number); return [`${r >= 4 ? r + 1 : r},${c}`, note]; }));
  myHalfCw.backgrounds = myHalfCw.backgrounds.map((b) => (b.r >= 4 ? { ...b, r: b.r + 1 } : b));
  run(cw, false, "processNewReceipts()");
  const haier = checksOf()["My home"]["03"].filter((c) => c.name === "Haier 300l");
  eq(plain(haier), [{ name: "Haier 300l", hand: 2, receipt: 2, status: "match", sheet: "My/2026/10/15", cell: "W5" }], "a check names mom's own cell, not the copy's, and that cell is not also \"no receipt\"");
}
eq([pure.ocrHandNote("", 2), pure.ocrHandNote("a\nБаримтаар: 2", 3), pure.ocrHandNote("Баримтаар: 2", null)], ["Баримтаар: 2", "a\nБаримтаар: 3", ""], "the receipts' line of a hand-typed cell's note is put in, changed or taken out; the rest stays");
ok(pure.ocrIsHandCell(3, "Баримтаар: 1") && pure.ocrIsHandCell("", "Баримтаар: 1"), "a cell with the receipts' line is still a hand-typed one");
const colourAfter = (start, color) => run(cw, false, `(() => { const r = SpreadsheetApp.openById("${CATALOG_ID}").getSheetByName("My/2026/10/15").getRange(9, 20); r.setBackground(${JSON.stringify(start)}); ocrSetCellColor(r, ${JSON.stringify(color)}); return r.getBackground(); })()`).value;
eq([colourAfter("#fff2cc", null), colourAfter("#ffe599", null), colourAfter("#fff2cc", "#d9ead3")], ["#fff2cc", "#ffffff", "#d9ead3"], "taking a colour away clears only the OCR's own: a transfer column keeps its yellow");
const old25 = monthTab(cw, "My home", "2025-12-05");
eq(cell(old25, 2, "Бичсэн"), "— 2026 оноос өөр оны баримт: workbook-д бичихгүй", "a receipt of another year is not written");
eq([pure.ocrWrittenState(rec({ ...okLine, "Бичсэн": "≈ My/2026/10/15!U22: workbook-д 1 — таарсан" })), pure.ocrWrittenState(rec({ ...okLine, "Бичсэн": "✗ нүдэнд гараар бичсэн 1 байна [k:||Haier 300l|1]" }))], ["compare", "retry"], "a line only compared, or refused, under the rules before is written now");

// ---------- the owner's rules for Тосгон's income (2026-10-23) ----------
{
  const ow = makeWorld();
  ow.props.GEMINI_KEY = "k";
  const tt = tos(ow);
  tt.cells[3][2] = 1000000; tt.cells[4][2] = 600000; // C, the cost: Meiling 501l, Haier 300l
  const IN = (day) => DAY("Тосгон", "орлого", day);
  const china = (lines) => receipt(null, lines, { document_type: "orlogo", price_currency: "CNY" });
  put(ow, IN("05"), "S1.HEIC", "s1", 1, receipt(null, [item("Meiling 501", "Meiling 501l", 1, 1280)], { document_type: "zarlaga" }));
  put(ow, IN("06"), "P1.HEIC", "p1", 2, china([item("Meiling 501", "Meiling 501l", 2, 150), item("Haier 300", "Haier 300l", 1, 90)]));
  put(ow, IN("08"), "P2.HEIC", "p2", 3, china([item("Meiling 501", "Meiling 501l", 2, 150), item("Haier 300", "Haier 300l", 1, 90), item("Haier 400", "Haier 400l", 1, 80)]));
  put(ow, IN("09"), "P3.HEIC", "p3", 4, china([item("Meiling 501", "Meiling 501l", 2, 150)]));
  eq(run(ow, false, "processNewReceipts()").error, null, "a run with the owner's Тосгон rules works");
  const s5 = monthTab(ow, "Тосгон", "2026-10-05");
  eq([rowsOf(s5), cell(s5, 2, "Шалгах шалтгаан"), cell(s5, 2, "Бараа"), cell(s5, 2, "Бичсэн")], [2, "Алгассан: зарлагын баримт орлогын хавтаст байсан", "", ""],
     "an issue slip in Тосгон's income folder is not read into lines: one row says it was skipped");
  const [p1, p2, p3] = ["06", "08", "09"].map((day) => monthTab(ow, "Тосгон", `2026-10-${day}`));
  ok(String(cell(p2, 2, "Шалгах шалтгаан")).includes("давхар зураг: P1.HEIC (06)") && cell(p2, 2, "Төлөв") === "Шалгах", "a photo read almost like an earlier one is marked as the same slip");
  eq([cell(p1, 2, "Бичсэн") !== "", cell(p2, 2, "Бичсэн"), cell(p3, 2, "Шалгах шалтгаан").includes("давхар")], [true, "", false], "the first is written, the twin waits; one line alike is not a twin");
  setCell(p2, 2, "Ээж: Зөв/Буруу", "Зөв");
  run(ow, false, "processNewReceipts()");
  ok(String(cell(p2, 2, "Бичсэн")).startsWith("Tos/2026/10!"), "once mom says Зөв the twin's line is written");
}

// ---------- mom typed the receipt's goods on another day: the receipt's day gets 0, its note says where ----------
{
  const ew = makeWorld();
  ew.props.GEMINI_KEY = "k";
  ew.props.OCR_FEED_TOKEN = "t";
  const my = ew.books[CATALOG_ID].sheets["My/2026/10/15"];
  const W = 23, Y = 25, AA = 27; // My home's expense of 3, 5 and 7 October
  my.cells[3][AA - 1] = 2; // Haier 300l: mom typed 2 on the 7th
  const e3 = put(ew, DAY("My home", "зарлага", "03"), "E3.HEIC", "e3", 1, receipt(null, [item("Haier 300", "Haier 300l", 2, 1200)], { document_type: "zarlaga" }));
  eq(run(ew, false, "processNewReceipts()").error, null, "a run with mom's number on another day works");
  const t3 = monthTab(ew, "My home", "2026-10-03");
  eq([my.value(4, W), my.notes[`4,${W}`], my.value(4, AA), my.notes[`4,${AA}`] || ""], [0, `${rowLink(t3, 2, `${e3.id}#1`)} = 0 (2026-10-03) (ээж 7-нд гараар: AA4)`, 2, ""], "the receipt's day gets 0 and its note names mom's cell; hers stays");
  eq(cell(t3, 2, "Бичсэн"), "My/2026/10/15!W4 ← 2 · Haier 300l (ээж 7-нд гараар бичсэн: AA4, энд 0)", "the line says so too");
  const feedChecks = () => plain(run(ew, true, `ocrFeedAnswer({ token: "t", mode: "receipts", month: "2026-10" }, PropertiesService.getScriptProperties(), new Date())`).value).checks["My home"];
  eq(feedChecks()["03"], [{ name: "Haier 300l", hand: 2, receipt: 2, status: "otherday", sheet: "My/2026/10/15", cell: "AA4" }], "the feed shows it as typed on another day, with mom's cell");
  put(ew, DAY("My home", "зарлага", "05"), "E5.HEIC", "e5", 2, receipt(null, [item("Haier 300", "Haier 300l", 2, 1200)], { document_type: "zarlaga" }));
  run(ew, false, "processNewReceipts()");
  eq([my.value(4, W), my.value(4, Y)], [0, 2], "mom's one number answers one receipt only: the next one is written");
  my.cells[3][AA - 1] = "";
  run(ew, false, "processNewReceipts()");
  eq([my.value(4, W), my.notes[`4,${W}`], cell(t3, 2, "Бичсэн")], [2, `${rowLink(t3, 2, `${e3.id}#1`)} = 2 (2026-10-03)`, "My/2026/10/15!W4 ← 2 · Haier 300l"], "once mom's number is gone the receipt's is written");
}

// ---------- the shops' paired layout: a shelf price after D, every shop column one further right ----------
{
  const p = run(makeWorld(), true, "0", pairedCode).ctx;
  const moveText = (text) => vm.runInContext(`ocrMovePairedText(${JSON.stringify(text)})`, p);
  eq([
    moveText("Sor/2026/10/15!U22 ← 2 · Haier 300l ✓"),
    moveText("check:My/2026/10/15!W4 ← 2 · Haier 300l (гараар 2)"),
    moveText("Tos/2026/10!AJ57 ← 1 · Haier 300l"),
    moveText("✗ нүд алга [k:a]\nSor/2026/10/15!E4 ← 1 · Meiling 501l"),
    moveText("≈ My/2026/10/15!AK5: workbook-д 1 — таарсан"),
  ], [
    "Sor/2026/10/15!V22 ← 2 · Haier 300l ✓",
    "check:My/2026/10/15!X4 ← 2 · Haier 300l (гараар 2)",
    "Tos/2026/10!AJ57 ← 1 · Haier 300l",
    "✗ нүд алга [k:a]\nSor/2026/10/15!F4 ← 1 · Meiling 501l",
    "≈ My/2026/10/15!AL5: workbook-д 1 — таарсан",
  ], "every shop cell named in Бичсэн moves one column right (the count AK -> AL); Тосгон's stay");
  eq([p.ocrSheetPlace({ place: "Соронзон", side: "in" }, "2026-10-01"), p.ocrSheetPlace({ place: "My home", side: "out" }, "2026-10-16"), p.ocrSheetPlace({ place: "Тосгон", side: "in" }, "2026-10-01")],
    ["Sor/2026/10/15 · орлого · F", "My/2026/10/31 · зарлага · V", "Tos/2026/10 · орлого · E"], "the paired layout: shop income from F, expense from V; Тосгон unchanged");

  // a line written while the shops' workbook had the old layout
  const w = makeWorld();
  w.props.GEMINI_KEY = "test-key";
  const SOR = () => w.books[CATALOG_ID].sheets["Sor/2026/10/15"];
  put(w, DAY("Соронзон", "орлого", "02"), "IMG_P1.HEIC", "photo-p1", 1, receipt("2026-10-02", [item("Haier 300", "Haier 300l", 1, 1150)], { document_type: "orlogo" }));
  run(w, false, "processNewReceipts()");
  const t2 = monthTab(w, "Соронзон", "2026-10-02");
  const rowOf = (tab, name) => tab.cells.findIndex((r, i) => i > 0 && r[header.indexOf("Бараа")] === name) + 1;
  const pr = rowOf(t2, "Haier 300l");
  eq([cell(t2, pr, "Бичсэн").split(" ← ")[0], SOR().value(6, 6)], ["Sor/2026/10/15!F6", 1], "old layout: day 2 of the income is F");

  // the new OCR code before the shops' workbook moved: nothing is read or written
  put(w, DAY("Соронзон", "орлого", "03"), "IMG_P2.HEIC", "photo-p2", 2, receipt("2026-10-03", [item("Haier 400", "Haier 400l", 1, 1450)], { document_type: "orlogo" }));
  const calls = w.geminiCalls.length;
  const stopped = run(w, false, "processNewReceipts()", pairedCode);
  ok(stopped.error === null && stopped.value === null && /хараахан шилжээгүй/.test(stopped.logs) && w.geminiCalls.length === calls && !w.props["ocr-paired-layout"], "the shops' workbook not moved yet: the run stops at once");

  // the shops' workbook moves (layout-migration.gs); the check copy still has the old columns
  const copyId = `${CATALOG_ID}-check`;
  w.books[copyId] = new Book(copyId, "бараа (check)", { "Sor/2026/10/15": new Sheet(JSON.parse(JSON.stringify(SOR().cells))) });
  w.props[`ocr-check-copy:${CATALOG_ID}`] = copyId;
  const movePaired = (book) => book.getSheets().filter((s) => /^(Sor|My)\//.test(s.getName())).forEach((s) => { s.insertColumnsAfter(4, 1); s.getRange(3, 5).setValue("яг харагдах үнэ"); s.insertColumnsAfter(41, 1); });
  movePaired(w.books[CATALOG_ID]);
  eq(SOR().value(6, 7), 1, "the written cell moved with its column (F -> G)");

  const moved = run(w, false, "processNewReceipts()", pairedCode);
  ok(moved.error === null && /шинэ бүтэцтэй боллоо/.test(moved.logs) && w.props["ocr-paired-layout"] === "1", "the first run after the move moves the OCR's own references once");
  eq(cell(t2, pr, "Бичсэн").split(" ← ")[0], "Sor/2026/10/15!G6", "Бичсэн names the moved cell");
  eq([w.books[copyId].sheets["Sor/2026/10/15"].value(3, 5), w.books[copyId].sheets["Sor/2026/10/15"].value(6, 7)], ["яг харагдах үнэ", 1], "the check copy got the same columns");
  const t3 = monthTab(w, "Соронзон", "2026-10-03");
  const r3 = rowOf(t3, "Haier 400l");
  eq([cell(t3, r3, "Бичсэн").split(" ← ")[0], SOR().value(5, 8)], ["Sor/2026/10/15!H5", 1], "a new line goes to the new column: day 3 is H");

  // mom corrects the old line: the moved cell changes, nothing lands in the old column
  setCell(t2, pr, "Тоо", 3);
  run(w, false, "processNewReceipts()", pairedCode);
  eq([SOR().value(6, 7), SOR().value(6, 6), SOR().value(6, 5)], [3, "", ""], "the correction reaches G6; F6 and the shelf price E6 stay empty");
  const again = run(w, false, "processNewReceipts()", pairedCode);
  ok(!/шинэ бүтэцтэй боллоо/.test(again.logs) && cell(t2, pr, "Бичсэн").startsWith("Sor/2026/10/15!G6"), "the references are moved only once");

  // the gift column (the workbook's gift-migration.gs): a column after the transfer out AK
  const g = run(makeWorld(), true, "0", giftCode).ctx;
  const moveGift = (text) => vm.runInContext(`ocrMoveGiftText(${JSON.stringify(text)})`, g);
  eq([moveGift("≈ My/2026/10/15!AL5: workbook-д 1 — таарсан"), moveGift("Sor/2026/10/15!AK22 ← 1 · Haier 300l"), moveGift("Tos/2026/10!BO5 ← 1 · Haier 300l")],
    ["≈ My/2026/10/15!AM5: workbook-д 1 — таарсан", "Sor/2026/10/15!AK22 ← 1 · Haier 300l", "Tos/2026/10!BO5 ← 1 · Haier 300l"], "the count AL moves to AM; the transfer out AK and Тосгон's cells stay");
  const beforeGift = run(w, false, "processNewReceipts()", giftCode);
  ok(beforeGift.value === null && /"Бэлэг" багана хараахан/.test(beforeGift.logs) && !w.props["ocr-gift-column"], "the shops' workbook without the gift column yet: the run stops at once");
  w.books[CATALOG_ID].getSheets().filter((sh) => /^(Sor|My)\//.test(sh.getName())).forEach((sh) => { sh.insertColumnsAfter(37, 1); sh.getRange(3, 38).setValue("Бэлэг"); });
  const withGift = run(w, false, "processNewReceipts()", giftCode);
  ok(withGift.error === null && /"Бэлэг" багана нэмэгдлээ/.test(withGift.logs) && w.props["ocr-gift-column"] === "1", "the first run after the gift move moves the OCR's references once");
  eq([w.books[copyId].sheets["Sor/2026/10/15"].value(3, 38), cell(t2, pr, "Бичсэн").split(" ← ")[0]], ["Бэлэг", "Sor/2026/10/15!G6"], "the check copy got the gift column; an income cell keeps its address");
  { // mom's Бэлэг on a shop's sale (the owner, 2026-10-28): approved, into the gift column AL instead of the day's sales column
    eq([g.ocrSheetPlace({ place: "Соронзон", side: "out", gift: true }, "2026-10-06"), g.ocrSheetPlace({ place: "Соронзон", side: "out" }, "2026-10-06")],
      ["Sor/2026/10/15 · бэлэг · AL", "Sor/2026/10/15 · зарлага · AA"], "a gift goes into AL, a sale into its day");
    const moves = (folder, mark) => plain(vm.runInContext(`ocrPostMoves({ 'Хавтас (ээж хийсэн)': ${JSON.stringify(folder)}, 'Ээж: Зөв/Буруу': ${JSON.stringify(mark)} }, '2026-10-06')`, g));
    eq([moves("Соронзон / зарлага", "Бэлэг"), moves("Соронзон / зарлага", "Зөв"), moves("Соронзон / орлого", "Бэлэг")],
      [[{ place: "Соронзон", side: "out", gift: true }], [{ place: "Соронзон", side: "out" }], [{ place: "Соронзон", side: "in" }]], "only a shop's sale can be a gift");
    ok(vm.runInContext(`ocrIsApproved({ "Ээж: Зөв/Буруу": "Бэлэг" }) && !ocrIsApproved({ "Ээж: Зөв/Буруу": "" })`, g), "Бэлэг is mom's approval");
    const choices = (name) => plain(vm.runInContext(`ocrMarkChoices({ getParent: () => ({ getName: () => ${JSON.stringify(name)} }) })`, g)).join();
    eq(["OCR Соронзон зарлага 2026-10", "OCR My home зарлага 2026-10", "OCR Соронзон орлого 2026-10", "OCR Тосгон зарлага 2026-10"].map(choices),
      ["Зөв,Буруу,Бэлэг", "Зөв,Буруу,Бэлэг", "Зөв,Буруу", "Зөв,Буруу"], "the Бэлэг choice in the two shops' sales spreadsheets only");
    put(w, DAY("Соронзон", "зарлага", "06"), "IMG_G1.HEIC", "photo-g1", 6, receipt("2026-10-06", [item("Haier 300", "Haier 300l", 1, 795)]));
    run(w, false, "processNewReceipts()", giftCode);
    const s6 = monthTab(w, "Соронзон", "2026-10-06", "зарлага");
    const sold = cell(s6, 2, "Бичсэн").split(" ← ")[0];
    ok(/^Sor\/2026\/10\/15!AA\d+$/.test(sold), `a sale goes into its day (${sold})`);
    s6.cells[1][header.indexOf("Ээж: Зөв/Буруу")] = "Бэлэг";
    run(w, false, "processNewReceipts()", giftCode);
    const given = cell(s6, 2, "Бичсэн");
    const row = Number(sold.match(/\d+$/)[0]);
    eq([given, SOR().value(row, 38), SOR().value(row, 27)], [`Sor/2026/10/15!AL${row} ← 1 · Haier 300l ✓`, 1, ""], "marked Бэлэг it moves out of the day into the gift column, approved");
  }
  // Тосгон gets the shelf price in E (its layout-migration.gs, 2026-10-22): its days move one right
  const TOS = () => w.books[TOSGON_ID].sheets["Tos/2026/10"];
  // the fake sheets share HEAD's rows: the shops' moves above wrote their headers into Тосгон's row 3 too
  TOS().cells = JSON.parse(JSON.stringify(TOS().cells)); TOS().cells[2] = ["Барааны нэрс"];
  put(w, DAY("Тосгон", "Тосгоноос Соронзон руу", "04"), "IMG_T1.HEIC", "photo-t1", 3, receipt("2026-10-04", [item("Haier 300", "Haier 300l", 1, 1150)]));
  run(w, false, "processNewReceipts()", giftCode);
  const t4 = monthTab(w, "Тосгон", "2026-10-04");
  const tr = rowOf(t4, "Haier 300l");
  const oldCell = cell(t4, tr, "Бичсэн").split("\n").find((line) => line.startsWith("Tos/")).split(" ← ")[0];
  const [, oldLetters, tosRow] = oldCell.match(/!([A-Z]+)(\d+)$/);
  const letterNumber = (letters) => letters.split("").reduce((total, letter) => total * 26 + letter.charCodeAt(0) - 64, 0);
  const oldColumn = letterNumber(oldLetters);
  eq([oldColumn, TOS().value(Number(tosRow), oldColumn)], [39, 1], "old Тосгон layout: the transfer out of day 4 is AM");
  const tosCalls = w.geminiCalls.length;
  const waiting = run(w, false, "processNewReceipts()", shelfCode);
  ok(waiting.value === null && /Тосгоны файл .*хараахан шилжээгүй/.test(waiting.logs) && w.geminiCalls.length === tosCalls && !w.props["ocr-tosgon-shelf"], "Тосгон not moved yet: the run stops at once");
  const tosCopyId = `${TOSGON_ID}-check`;
  w.books[tosCopyId] = new Book(tosCopyId, "Тосгон (check)", { "Tos/2026/10": new Sheet(JSON.parse(JSON.stringify(TOS().cells))) });
  w.props[`ocr-check-copy:${TOSGON_ID}`] = tosCopyId;
  TOS().insertColumnsAfter(4, 1); TOS().getRange(3, 5).setValue("яг харагдах үнэ");
  const shelfMoved = run(w, false, "processNewReceipts()", shelfCode);
  ok(shelfMoved.error === null && /Тосгоны файлд "яг харагдах үнэ" нэмэгдлээ/.test(shelfMoved.logs) && w.props["ocr-tosgon-shelf"] === "1", "the first run after Тосгон's move moves its references once");
  eq(cell(t4, tr, "Бичсэн").split("\n").find((line) => line.startsWith("Tos/")).split(" ← ")[0], `Tos/2026/10!AN${tosRow}`, "Бичсэн names Тосгон's moved cell (AM -> AN)");
  eq([w.books[tosCopyId].sheets["Tos/2026/10"].value(3, 5), w.books[tosCopyId].sheets["Tos/2026/10"].value(Number(tosRow), 40)], ["яг харагдах үнэ", 1], "Тосгон's check copy got the column too");
  eq(cell(t2, pr, "Бичсэн").split(" ← ")[0], "Sor/2026/10/15!G6", "a shop cell keeps its address");
  put(w, DAY("Тосгон", "Тосгоноос Соронзон руу", "05"), "IMG_T2.HEIC", "photo-t2", 4, receipt("2026-10-05", [item("Haier 300", "Haier 300l", 1, 1150)]));
  run(w, false, "processNewReceipts()", shelfCode);
  const t5 = monthTab(w, "Тосгон", "2026-10-05");
  eq(cell(t5, rowOf(t5, "Haier 300l"), "Бичсэн").split("\n").find((line) => line.startsWith("Tos/")).split(" ← ")[0], `Tos/2026/10!AO${tosRow}`, "a new Тосгон line goes to the new column: day 5 out is AO");
  eq(TOS().value(Number(tosRow), 5), "", "Тосгон's shelf price E is never written by the OCR");
}

// a slip Тосгон sent, filed in Соронзон's income folder: one slip, two sides (the owner, 2026-10-28):
// Тосгон's expense and Соронзон's income at once, no note; the photo is copied into Тосгон's folder
// to Соронзон, and the copy is never read. Until 2026-10-28 it was noted and only Соронзон's side
// written: such a row takes the new meaning in the next review and gets Тосгон's side.
{
  const slipResult = receipt(null, [item("Meiling 501", "Meiling 501l", 1, 1280)], { document_type: "zarlaga", document_title: "ЗАРЛАГЫН БАРИМТ", header_sender: "Тос", header_receiver: "Соронзонд" });
  const setUp = () => {
    const sw = makeWorld();
    sw.props.GEMINI_KEY = "test-key";
    put(sw, DAY("Соронзон", "орлого", "02"), "S.HEIC", "s", 9, slipResult);
    const so2 = () => monthTab(sw, "Соронзон", "2026-10-02");
    const sr = () => so2().cells.findIndex((r, i) => i > 0 && r[header.indexOf("Баримт")] === "S.HEIC") + 1;
    const at = (title) => so2().value(sr(), header.indexOf(title) + 1);
    const sorOf = sw.books[CATALOG_ID].sheets["Sor/2026/10/15"];
    const misfiled = () => Object.values(calendarOf(sw, "Соронзон")["02"]).reduce((total, numbers) => total + numbers[6], 0);
    const copies = () => (sw.root.find("Тосгон", "Тосгоноос Соронзон руу", "2026-10", "02") || { files: [] }).files.filter((f) => f.name === "S.HEIC");
    return { sw, so2, at, sorOf, misfiled, copies };
  };
  { // read now
    const { sw, so2, at, sorOf, misfiled, copies } = setUp();
    eq(run(sw, false, "processNewReceipts()").error, null, "a run reads the slip");
    eq([at("Хавтас (ээж хийсэн)"), at("Шалгах шалтгаан"), at("Төлөв")], ["Соронзон / орлого (Тосгоноос)", "", "OK"], "the row says the slip came from Тосгон; nothing to check");
    eq([tos(sw).value(4, AJ + 1), sorOf.value(4, E + 1), misfiled()], [1, 1, 0], "Тосгон's expense of the 2nd and Соронзон's income of the 2nd, once each; the calendar counts no misfiled photo");
    eq(copies().length, 1, "the photo is copied into Тосгон's folder to Соронзон, the same day");
    const rowsBefore = so2().cells.length;
    run(sw, false, "processNewReceipts()");
    eq([copies().length, so2().cells.length, tos(sw).value(4, AJ + 1), sorOf.value(4, E + 1), monthTab(sw, "Тосгон", "2026-10-02")], [1, rowsBefore, 1, 1, null],
       "the next run neither reads the copy nor copies again: no new rows, no count twice");
  }
  { // read by the version before: noted, Соронзон's side only; the review adds Тосгон's
    const { sw, at, sorOf, misfiled, copies } = setUp();
    const before = code.replace("function ocrSlipFolder(folder, result) {", "function ocrSlipFolder(folder, result) { return folder;")
      .replace("function ocrMirrorPhoto(file, path, sender, context) {", "function ocrMirrorPhoto(file, path, sender, context) { return [];")
      .replace("function ocrTosgonSentRows(sheet, where, records, context) {", "function ocrTosgonSentRows(sheet, where, records, context) { return records;");
    ok(before !== code && run(sw, false, "processNewReceipts()", before.replace(/function ocrMisfiledReason\(result, folder\) \{/, "function ocrMisfiledReason(result, folder) { if (folder.kind === 'orlogo' && ocrSentByTosgon(result)) return `хавтас буруу байж магадгүй: Тосгоноос ${folder.moves[0].place} руу хавтсанд байх ёстой юм шиг (илгээгч: ${result.header_sender})`;")).error === null, "the version before reads it");
    eq([at("Хавтас (ээж хийсэн)"), at("Шалгах шалтгаан"), tos(sw).value(4, AJ + 1), sorOf.value(4, E + 1), misfiled(), copies().length], ["Соронзон / орлого", "хавтас буруу байж магадгүй: Тосгоноос Соронзон руу хавтсанд байх ёстой юм шиг (илгээгч: Тос)", "", 1, 1, 0], "as it stood: noted, Соронзон's income only");
    const dry = run(sw, true, "processNewReceipts()");
    eq([dry.error, at("Хавтас (ээж хийсэн)"), copies().length, /Тосгоноос ирсэн баримт: Тосгоны зарлага, дэлгүүрийн орлого/.test(dry.logs)], [null, "Соронзон / орлого", 0, true], "DRY_RUN only says so");
    Object.keys(sw.props).filter((key) => key.startsWith("ocr-book-setup:")).forEach((key) => { sw.props[key] = "old"; });
    eq(run(sw, false, "processNewReceipts()").error, null, "the next review");
    eq([at("Хавтас (ээж хийсэн)"), at("Шалгах шалтгаан"), at("Төлөв"), tos(sw).value(4, AJ + 1), sorOf.value(4, E + 1), misfiled(), copies().length], ["Соронзон / орлого (Тосгоноос)", "", "OK", 1, 1, 0, 1],
       "the row takes the new meaning: Тосгон's expense is added, Соронзон's income stays once, the note goes, the photo is copied");
    run(sw, false, "processNewReceipts()");
    eq([tos(sw).value(4, AJ + 1), sorOf.value(4, E + 1), copies().length], [1, 1, 1], "and once only");
  }
}

// a slip from the city, filed in Соронзон's income folder (the owner, 2026-10-28): Тосгон's income,
// Тосгон's expense and Соронзон's income of its day, once each; the photo is copied into Тосгон's
// income folder and its folder to Соронзон, and the copies are never read
{
  const cw3 = makeWorld();
  cw3.props.GEMINI_KEY = "test-key";
  put(cw3, DAY("Соронзон", "орлого", "02"), "C.HEIC", "c", 9, receipt(null, [item("Meiling 501", "Meiling 501l", 1, null)], { document_type: "zarlaga", document_title: "ЗАРЛАГЫН БАРИМТ", header_sender: "Хотоос", header_receiver: "Соронзон" }));
  eq(run(cw3, false, "processNewReceipts()").error, null, "a run reads the slip");
  const so2 = monthTab(cw3, "Соронзон", "2026-10-02");
  const at = (title) => so2.value(so2.cells.findIndex((r, i) => i > 0 && r[header.indexOf("Баримт")] === "C.HEIC") + 1, header.indexOf(title) + 1);
  const tosIn = vm.runInContext('ocrDayColumn(OCR_TOSGON.layout, "in", 1)', pure);
  const sorOf = cw3.books[CATALOG_ID].sheets["Sor/2026/10/15"];
  const copiesIn = (...folders) => (cw3.root.find("Тосгон", ...folders, "2026-10", "02") || { files: [] }).files.filter((f) => f.name === "C.HEIC").length;
  const state = () => [tos(cw3).value(4, tosIn), tos(cw3).value(4, AJ + 1), sorOf.value(4, E + 1), copiesIn("орлого"), copiesIn("Тосгоноос Соронзон руу")];
  eq([at("Хавтас (ээж хийсэн)"), at("Шалгах шалтгаан"), at("Төлөв")], ["Соронзон / орлого (Хотоос)", "", "OK"], "the row says the slip came from the city; nothing to check");
  eq(state(), [1, 1, 1, 1, 1], "Тосгон's income, Тосгон's expense and Соронзон's income of the 2nd once each; the photo lies in Тосгон's income folder and its folder to Соронзон too");
  const rowsBefore = so2.cells.length;
  run(cw3, false, "processNewReceipts()");
  eq([...state(), so2.cells.length, monthTab(cw3, "Тосгон", "2026-10-02")], [1, 1, 1, 1, 1, rowsBefore, null], "the next run neither reads the copies nor copies again: nothing twice");
}

// a Тосгон transfer photo is copied into the shop's income folder of its day (the owner, 2026-10-28),
// and the copy is never read: the transfer is written once, from Тосгон's folder
{
  const tw = makeWorld();
  tw.props.GEMINI_KEY = "test-key";
  put(tw, DAY("Тосгон", "Тосгоноос My home руу", "03"), "T.HEIC", "t", 9, receipt(null, [item("Meiling 501", "Meiling 501l", 2, 1280)]));
  eq(run(tw, false, "processNewReceipts()").error, null, "a run reads the transfer photo");
  const copies = () => (tw.root.find("My home", "орлого", "2026-10", "03") || { files: [] }).files.filter((f) => f.name === "T.HEIC");
  const myOf = tw.books[CATALOG_ID].sheets["My/2026/10/15"];
  const myRow = myOf.cells.findIndex((r) => r && r[0] === "Meiling 501l") + 1;
  eq([copies().length, tos(tw).value(4, AJ + 2), myOf.value(myRow, E + 2)], [1, 2, 2], "the photo lies in My home's income folder of the 3rd too; Тосгон's expense and My home's income once each");
  run(tw, false, "processNewReceipts()");
  eq([copies().length, tos(tw).value(4, AJ + 2), myOf.value(myRow, E + 2), monthTab(tw, "My home", "2026-10-03")], [1, 2, 2, null], "the copy is not read: no My home income spreadsheet, nothing twice");
}

// a photo outside a day folder is read as the date on its slip; once mom puts it into another day's
// folder, its rows, its photo and its workbook count follow (the owner, 2026-10-28)
{
  const lw = makeWorld();
  lw.props.GEMINI_KEY = "test-key";
  const loose = put(lw, ["Соронзон", "зарлага", "2026-10"], "L.HEIC", "loose", 9, receipt("2026-10-07", [item("Haier 300", "Haier 300l", 1, 1150)], { document_type: "zarlaga" }));
  run(lw, false, "processNewReceipts()");
  const nameOf = (tab, name) => tab.cells.findIndex((r, i) => i > 0 && r && r[header.indexOf("Бараа")] === name) + 1;
  const sor = lw.books[CATALOG_ID].sheets["Sor/2026/10/15"];
  const hr = sor.cells.findIndex((r) => r[0] === "Haier 300l") + 1;
  const colOf = (iso) => vm.runInContext(`ocrColumnNumber("${pure.ocrSheetPlace({ place: "Соронзон", side: "out" }, iso).split(" · ")[2]}")`, pure);
  const t7 = monthTab(lw, "Соронзон", "2026-10-07", "зарлага");
  const r7 = nameOf(t7, "Haier 300l");
  eq([monthTab(lw, "Соронзон", "2026-10-10", "зарлага"), cell(t7, r7, "Шалгах шалтгаан"), cell(t7, r7, "Төлөв"), cell(t7, r7, "Бичсэн").split("!")[1].replace(/\d+ ←.*/s, ""), sor.value(hr, colOf("2026-10-07"))],
     [null, "өдрийн хавтсанд хийгээгүй, баримт дээрх огноогоор (2026-10-07)", "Шалгах", pure.ocrSheetPlace({ place: "Соронзон", side: "out" }, "2026-10-07").split(" · ")[2], 1],
     "a loose slip dated the 7th goes under the 7th (not the day it came in), marked, and is counted on the 7th");
  t7.getRange(r7, header.indexOf("Ээж: Зөв/Буруу") + 1).setValue(""); // nothing of mom's yet
  loose.parent.files = loose.parent.files.filter((f) => f !== loose); lw.root.path("Соронзон", "зарлага", "2026-10", "09").add(loose); // mom puts it into the 9th
  const moving = run(lw, false, "processNewReceipts()");
  const t9 = monthTab(lw, "Соронзон", "2026-10-09", "зарлага");
  const r9 = t9 ? nameOf(t9, "Haier 300l") : 0;
  eq([moving.error, nameOf(t7, "Haier 300l"), t7.cells.filter((r) => r[PHOTO] === "L.HEIC").length, t7.images.length, r9 > 0, t9 && t9.images.length],
     [null, 0, 0, 0, true, 1], "its rows and photo leave the 7th's tab for the 9th's");
  eq([cell(t9, r9, "Шалгах шалтгаан"), cell(t9, r9, "Төлөв"), cell(t9, r9, "ID"), cell(t9, r9, "Бичсэн").split(" ← ")[0], sor.value(hr, colOf("2026-10-07")), sor.value(hr, colOf("2026-10-09"))],
     ["", "OK", `${loose.getId()}#1`, `Sor/2026/10/15!${pure.ocrSheetPlace({ place: "Соронзон", side: "out" }, "2026-10-09").split(" · ")[2]}${hr}`, "", 1],
     "the note is gone, the ID stays, and the count moves from the 7th's column to the 9th's");
  ok(/зургийг Соронзон \/ зарлага \/ 2026-10 \/ 09 руу зөөсөн/.test(moving.logs) && lw.props[`ocr-done:${loose.getId()}`].endsWith(" / 09"), "the log says so; the photo is remembered under its new day");
  const again = run(lw, false, "processNewReceipts()");
  eq([again.error, nameOf(t9, "Haier 300l"), t9.images.length, sor.value(hr, colOf("2026-10-09"))], [null, r9, 1, 1], "a run later changes nothing");
  // into another kind's folder: the rows go to that kind's spreadsheet, their folder with them, the count to the income side
  loose.parent.files = loose.parent.files.filter((f) => f !== loose); lw.root.path("Соронзон", "орлого", "2026-10", "09").add(loose);
  const kinds = run(lw, false, "processNewReceipts()");
  const i9 = monthTab(lw, "Соронзон", "2026-10-09", "орлого");
  const ri = i9 ? nameOf(i9, "Haier 300l") : 0;
  const inCol = vm.runInContext(`ocrColumnNumber("${pure.ocrSheetPlace({ place: "Соронзон", side: "in" }, "2026-10-09").split(" · ")[2]}")`, pure);
  eq([kinds.error, nameOf(t9, "Haier 300l"), t9.images.length, ri > 0, cell(i9, ri, "Хавтас (ээж хийсэн)"), cell(i9, ri, "Төлөв"), cell(i9, ri, "Бичсэн").split(" ← ")[0]],
     [null, 0, 0, true, "Соронзон / орлого", "OK", `Sor/2026/10/15!${pure.ocrSheetPlace({ place: "Соронзон", side: "in" }, "2026-10-09").split(" · ")[2]}${hr}`], "moved into the орлого folder: the rows leave the зарлага spreadsheet for the орлого one, as income");
  eq([sor.value(hr, colOf("2026-10-09")), sor.value(hr, inCol)], ["", 1], "the sale of the 9th is taken out, the income of the 9th written");
  // loose in another month's folder: that month's spreadsheet, under the day it came in there (or the 1st), noted
  loose.parent.files = loose.parent.files.filter((f) => f !== loose); lw.root.path("Соронзон", "орлого", "2026-09").add(loose);
  const months = run(lw, false, "processNewReceipts()");
  const sep = monthTab(lw, "Соронзон", "2026-09-01", "орлого");
  const rs = sep ? nameOf(sep, "Haier 300l") : 0;
  eq([months.error, i9 ? nameOf(i9, "Haier 300l") : 0, rs > 0, sep && cell(sep, rs, "Шалгах шалтгаан").split(";")[0], sep && cell(sep, rs, "Төлөв"), lw.props[`ocr-done:${loose.getId()}`]],
     [null, 0, true, "өдрийн хавтсанд хийгээгүй, 2026-09-01-ээр (өдөр тодорхойгүй)", "Шалгах", "Зураг уншуулах/2026-09/Орлого/OCR Соронзон орлого 2026-09 / 01"],
     "put loose into September's folder (it came in in October): September's spreadsheet, the 1st, marked");
  // a photo of another place, or gone: its rows stay
  const other = put(lw, DAY("My home", "зарлага", "03"), "O.HEIC", "other", 9, receipt("2026-10-03", [item("Haier 300", "Haier 300l", 1, 1150)], { document_type: "zarlaga" }));
  run(lw, false, "processNewReceipts()");
  const m3 = monthTab(lw, "My home", "2026-10-03", "зарлага");
  const rm = nameOf(m3, "Haier 300l");
  ok(rm > 0, "a My home sale of the 3rd is read");
  other.parent.files = other.parent.files.filter((f) => f !== other); lw.root.path("Соронзон", "зарлага", "2026-10", "03").add(other);
  run(lw, false, "processNewReceipts()");
  eq([nameOf(m3, "Haier 300l"), cell(m3, rm, "Хавтас (ээж хийсэн)")], [rm, "My home / зарлага"], "moved to the other shop's folder: left where it is (another place)");
  // deleted by mom after it was read (the owner, 2026-10-28): its rows, photo and the counts the script
  // wrote leave; its slip is forgotten, so one put back is read again
  const twin = put(lw, DAY("My home", "зарлага", "03"), "O 2.HEIC", "other", 9);
  run(lw, false, "processNewReceipts()");
  const twinRow = () => m3.cells.findIndex((r, i) => i > 0 && r && r[PHOTO] === "O 2.HEIC") + 1;
  ok(twinRow() > 0 && cell(m3, twinRow(), "Шалгах шалтгаан").startsWith("давхардсан, өмнө уншсан:"), "a second copy of the slip is noted as a duplicate");
  const my = lw.books[CATALOG_ID].sheets["My/2026/10/15"];
  const [, myLetters, myRow] = cell(m3, rm, "Бичсэн").split(" ← ")[0].match(/!([A-Z]+)(\d+)$/);
  const myColumn = vm.runInContext(`ocrColumnNumber("${myLetters}")`, pure);
  eq(my.value(Number(myRow), myColumn), 1, "the sale of the 3rd is in My home's sheet");
  twin.setTrashed(true);
  run(lw, false, "processNewReceipts()");
  eq([twinRow(), nameOf(m3, "Haier 300l"), m3.images.length, lw.props[`ocr-done:${twin.getId()}`], Object.values(lw.props).includes("My home/зарлага/2026-10/03/O.HEIC")],
     [0, rm, 1, undefined, true], "the duplicate copy deleted: its note row goes, the slip stays known");
  other.setTrashed(true);
  // nothing in the spreadsheet changed: the pass finds it by the photo (ocr-watch.gs)
  const watch = (now) => plain(run(lw, false, `ocrWithLeftPhotos([], { props: PropertiesService.getScriptProperties(), root: DriveApp.getFolderById(OCR_ROOT_FOLDER_ID) }, ${now}).map((entry) => entry.file.getName())`).value);
  eq([watch("new Date(2026, 9, 28)"), watch("new Date(2026, 10, 30)"), watch("new Date(2026, 11, 1)")],
     [["OCR My home зарлага 2026-10"], ["OCR My home зарлага 2026-10"], []], "a deleted photo's spreadsheet is looked at in its month and the next, not later");
  const gone = run(lw, false, "processNewReceipts()");
  eq([gone.error, nameOf(m3, "Haier 300l"), m3.cells.filter((r) => r && r[PHOTO] === "O.HEIC").length, m3.images.length, my.value(Number(myRow), myColumn)],
     [null, 0, 0, 0, ""], "the photo deleted: its rows and photo leave the tab, the count leaves the workbook");
  eq([lw.props[`ocr-done:${other.getId()}`], Object.values(lw.props).includes("My home/зарлага/2026-10/03/O.HEIC")], [undefined, false], "the photo and its slip are forgotten");
  ok(new RegExp(`${other.getId()} зургийг устгасан тул мөрүүд нь хасагдлаа`).test(gone.logs), "the log says so");
  lw.root.path("My home", "зарлага", "2026-10", "03").add(other); other.setTrashed(false); // mom puts it back
  run(lw, false, "processNewReceipts()");
  const back = nameOf(m3, "Haier 300l");
  eq([back > 0, m3.images.length, my.value(Number(myRow), myColumn), cell(m3, back, "ID")], [true, 1, 1, `${other.getId()}#1`], "put back from the trash: read again as new, counted again");
}

// ---------- cells the OCR leaves alone (the owner, 2026-10-23: Sor/2026/4/30 F95, F101, F134) ----------
{
  eq(plain(run(makeWorld(), true, "OCR_HANDS_OFF_CELLS").value), ["Sor/2026/4/30!F95", "Sor/2026/4/30!F101", "Sor/2026/4/30!F134"], "the three cells set to 0 by hand are left alone");
  const ho = makeWorld();
  ho.props.GEMINI_KEY = "k";
  const half = ho.books[CATALOG_ID].sheets["My/2026/10/15"];
  const W3 = 23;
  half.cells[3][W3 - 1] = 2; // Haier 300l, typed by hand
  half.cells[4][W3 - 1] = 0; // Meiling 501l: set to 0 by hand, with a note of its own
  half.notes[`5,${W3}`] = "OCR: K.HEIC [k#2] = 1 @2026-10-03\nээж 5-нд гараар оруулсан, 0 болгосон";
  put(ho, DAY("My home", "зарлага", "03"), "K.HEIC", "k", 1, receipt(null, [item("Haier 300", "Haier 300l", 2, 1200), item("Meiling 501", "Meiling 501l", 1, 1280)], { document_type: "zarlaga" }));
  const handsOff = (source) => source.replace(/const OCR_HANDS_OFF_CELLS = \[[^\]]*\];/, 'const OCR_HANDS_OFF_CELLS = ["My/2026/10/15!W5"];');
  eq(run(ho, false, "processNewReceipts()", handsOff(code)).error, null, "a run with a hands-off cell works");
  const k = monthTab(ho, "My home", "2026-10-03");
  eq([half.value(5, W3), half.notes[`5,${W3}`], bgOf(half, 5, W3)], [0, "OCR: K.HEIC [k#2] = 1 @2026-10-03\nээж 5-нд гараар оруулсан, 0 болгосон", null], "the hands-off cell keeps its number, note and colour");
  eq(cell(k, 3, "Бичсэн"), "", "its receipt line is left waiting, not written anywhere");
  ok(cell(k, 2, "Бичсэн").startsWith("check:My/2026/10/15!W4"), "the other line of the receipt is written as usual");
}

// ---------- the shops' sheets of 4/1-8/15 out of the OCR's reach (the owner, 2026-10-29) ----------
{
  eq(["Sor/2026/4/15!G192", "My/2026/8/15!E4", "check:Sor/2026/6/30!T259", "Sor/2026/3/31!E4", "Sor/2026/8/31!E4", "Tos/2026/5!E4", "Sor/2026/10/15!E4", "Sheet100!A1"].map((c) => pure.ocrIsHandsOffCell(c)),
     [true, true, true, false, false, false, false, false], "a shop's half-month whose last day is in 4/1-8/15 is out of reach, its check copy too; Тосгон's month and the other half-months are not");
  // the first half of October out of reach here, as April to August is in the shipped code
  const october = (source) => source.replace(/const OCR_HANDS_OFF_PERIOD = \{[^}]*\};/, "const OCR_HANDS_OFF_PERIOD = { from: '2026-10-01', until: '2026-10-15' };");
  const hp = makeWorld();
  hp.props.GEMINI_KEY = "k";
  const half = hp.books[CATALOG_ID].sheets["My/2026/10/15"];
  put(hp, DAY("My home", "зарлага", "03"), "H.HEIC", "h", 1, receipt(null, [item("Haier 300", "Haier 300l", 1, 1200)], { document_type: "zarlaga" }));
  run(hp, false, "processNewReceipts()");
  const h = monthTab(hp, "My home", "2026-10-03");
  const [, letters, row] = cell(h, 2, "Бичсэн").match(/^My\/2026\/10\/15!([A-Z]+)(\d+)/);
  const column = run(hp, true, `ocrColumnNumber("${letters}")`).value;
  eq(half.value(Number(row), column), 1, "written while the half-month was not out of reach");
  half.cells[row - 1][column - 1] = ""; half.notes[`${row},${column}`] = ""; // the owner takes the OCR's number out
  put(hp, DAY("My home", "зарлага", "05"), "M.HEIC", "m", 2, receipt(null, [item("Meiling 501", "Meiling 501l", 1, 1280)], { document_type: "zarlaga" }));
  put(hp, DAY("My home", "зарлага", "20"), "L.HEIC", "l", 3, receipt(null, [item("Haier 90-197", "Haier 90-197", 1, 450)], { document_type: "zarlaga" }));
  const out = run(hp, false, "processNewReceipts()", october(code));
  const text = run(hp, true, "OCR_HANDS_OFF_TEXT", october(code)).value;
  eq([out.error, text], [null, "— 2026-10-01 – 2026-10-15: дэлгүүрүүдийн workbook-д OCR бичихгүй"], "a run with the period works; the text names it");
  eq([cell(h, 2, "Бичсэн"), half.value(Number(row), column), half.notes[`${row},${column}`]], [text, "", ""], "a line written there before: its cell left as the owner made it, \"Бичсэн\" says why");
  const m = monthTab(hp, "My home", "2026-10-05");
  eq([cell(m, 2, "Бичсэн"), half.value(5, column + 2), half.notes[`5,${column + 2}`] || ""], [text, "", ""], "a new line of those days is not written");
  ok(cell(monthTab(hp, "My home", "2026-10-20"), 2, "Бичсэн").startsWith("My/2026/10/31!"), "a line of the next half-month is written as usual");
  eq(pure.ocrWrittenState(rec({ ...okLine, "Бичсэн": run(hp, true, "OCR_HANDS_OFF_TEXT").value })), "hands off", "the shipped text is read as such");
  const notes = JSON.stringify(half.notes);
  run(hp, false, "processNewReceipts()", october(code));
  eq([cell(h, 2, "Бичсэн"), JSON.stringify(half.notes), half.value(Number(row), column)], [text, notes, ""], "run again: nothing changes");
}

// ---------- "Бичсэн" and Тулгалт named again after rows are put into a workbook (the owner, 2026-10-29) ----------
{
  const rw = makeWorld();
  rw.props.GEMINI_KEY = "k";
  const half = rw.books[CATALOG_ID].sheets["My/2026/10/15"];
  eq(plain(run(rw, true, 'ocrMonthTabs("2026-10")').value),
     [{ bookId: TOSGON_ID, tab: "Tos/2026/10" }, { bookId: CATALOG_ID, tab: "Sor/2026/10/15" }, { bookId: CATALOG_ID, tab: "Sor/2026/10/31" }, { bookId: CATALOG_ID, tab: "My/2026/10/15" }, { bookId: CATALOG_ID, tab: "My/2026/10/31" }],
     "a month's workbook tabs: Тосгон's month, each shop's two half-months");
  put(rw, DAY("My home", "зарлага", "03"), "R.HEIC", "r", 1, receipt(null, [item("Haier 300", "Haier 300l", 1, 1200)], { document_type: "zarlaga" }));
  run(rw, false, "processNewReceipts()");
  const r = monthTab(rw, "My home", "2026-10-03");
  const [, letters] = cell(r, 2, "Бичсэн").match(/^My\/2026\/10\/15!([A-Z]+)4 /);
  const queued = () => (rw.props["ocr-relink-queue"] ? JSON.parse(rw.props["ocr-relink-queue"]).months : []);
  const goesOn = () => rw.triggers.filter((t) => t.getHandlerFunction() === "reviewMovedOcrBooksNow").length;
  const watch = (dry = false) => run(rw, dry, "ocrWatchWorkbookRows(PropertiesService.getScriptProperties())");
  ok(rw.props["ocr-relink-rows:2026-10"] !== undefined && queued().length === 0 && goesOn() === 0, "the pass keeps the month's rows; the first time there is nothing to compare with");
  const putRowAbove = (name) => { // a row put in above Haier 300l, by hand or by another script
    half.cells.splice(3, 0, [name]);
    const down = (key) => { const [row, column] = key.split(",").map(Number); return `${row >= 4 ? row + 1 : row},${column}`; };
    half.notes = Object.fromEntries(Object.entries(half.notes).map(([key, note]) => [down(key), note]));
    half.fonts = Object.fromEntries(Object.entries(half.fonts || {}).map(([key, font]) => [down(key), font]));
    half.backgrounds = half.backgrounds.map((b) => (b.r >= 4 ? { ...b, r: b.r + 1 } : b));
  };
  putRowAbove("Aucma 1");
  const watched = watch();
  eq([watched.error, plain(watched.value), queued(), goesOn()], [null, ["2026-10"], ["2026-10"], 1], "a row put in: the month is queued, a run a minute later set");
  eq(cell(r, 2, "Бичсэн").split(" ")[0], `My/2026/10/15!${letters}4`, "till then \"Бичсэн\" names the row before");
  const relinked = run(rw, false, "reviewMovedOcrBooksNow()");
  eq([relinked.error, cell(r, 2, "Бичсэн").split(" ")[0], half.value(5, run(rw, true, `ocrColumnNumber("${letters}")`).value), plain(relinked.value).left, queued(), goesOn()],
     [null, `My/2026/10/15!${letters}5`, 1, 0, [], 0], "the queued run names the product's row now; nothing left, it stops");
  const checks = monthBook(rw, "My home", "2026-10-03").book.getSheetByName("Тулгалт");
  ok(checks.cells.some((row) => row.includes(`My/2026/10/15!${letters}5`)) && !checks.cells.some((row) => row.includes(`My/2026/10/15!${letters}4`)), "Тулгалт names it too");
  run(rw, false, "processNewReceipts()");
  eq([queued(), goesOn()], [[], 0], "the pass after: the rows are as kept, nothing queued");
  putRowAbove("Aucma 2");
  const kept = rw.props["ocr-relink-rows:2026-10"];
  eq([watch(true).error, rw.props["ocr-relink-rows:2026-10"], queued(), goesOn()], [null, kept, [], 0], "DRY_RUN: nothing kept or queued");
  run(rw, false, "processNewReceipts()");
  eq([queued(), goesOn()], [["2026-10"], 1], "every pass looks: rows put in are found by the next one");
  rw.locked = true; // the hourly pass holds the lock
  rw.triggers = [];
  eq([run(rw, false, "reviewMovedOcrBooksNow()").error, goesOn(), queued()], [null, 1, ["2026-10"]], "the lock taken: it tries again a minute later");
  rw.locked = false;
  // a new product put in by the queued review itself moves the rows again: its month starts over
  const restarted = run(rw, false, "(() => { const real = reviewOcrBook; let n = 0; reviewOcrBook = (...a) => { if (n++ === 0) ocrRelinkSoon(PropertiesService.getScriptProperties(), ['2026-10']); return real(...a); }; try { return reviewMovedOcrBooksNow(); } finally { reviewOcrBook = real; } })()");
  const books = run(rw, true, "findOcrBooks(DriveApp.getFolderById(OCR_ROOT_FOLDER_ID), null).filter((entry) => entry.month === '2026-10').length").value;
  eq([restarted.error, plain(restarted.value).reviewed, queued(), goesOn()], [null, books + 1, [], 0], "rows moved during the run: the month's spreadsheets reviewed again, then it stops");
  // a line waiting for mom keeps "Бичсэн" as it was; Тулгалт still names its product's row now (2026-10-29: 3 such lines)
  const waiting = run(rw, true, `ocrWrittenChecks([{ "Бичсэн": "✗ нүдэнд гараар бичсэн 1 байна [k:v3||Haier 300l|1]\\nMy/2026/10/15!${letters}4 ← 1 · Haier 300l" }], new Map())`);
  eq([waiting.error, plain(waiting.value).map((check) => check.cell)], [null, [`${letters}6`]], "Тулгалт finds a written cell by its product, whatever \"Бичсэн\" says");
}

// ---------- a pass out of time leaves the count report to the next one (2026-10-29: two passes ran past Google's 6 minutes) ----------
{
  const tw = makeWorld();
  tw.props.GEMINI_KEY = "k";
  const late = run(tw, false, "runOcrPass(0)");
  ok(late.error === null && late.logs.includes("цаг дууссан тул дараагийн ажлаар") && /OCR \(.*\): .*\(\d+ сек\)/.test(late.logs), `out of time: no report, the seconds logged: ${late.logs}`);
  ok(!run(tw, false, "processNewReceipts()").logs.includes("цаг дууссан"), "in time: the report as always");
}

// ---------- a hand-typed cell's line is not written again when the check copy's rows differ (2026-10-29:
// "check:…!V269" against the original's V261: 26 of 77 lines of a review rewritten, 20 s, nothing changed) ----------
{
  const hw = makeWorld();
  hw.props.GEMINI_KEY = "k";
  const half = hw.books[CATALOG_ID].sheets["My/2026/10/15"];
  const W = 23; // My home's expense column of 3 October
  half.cells[3][W - 1] = 2; // Haier 300l, typed by hand: the receipt makes the same
  put(hw, DAY("My home", "зарлага", "03"), "H.HEIC", "h", 1, receipt(null, [item("Haier 300", "Haier 300l", 2, 1200, 0.7)], { document_type: "zarlaga" }));
  eq(run(hw, false, "processNewReceipts()").error, null, "the line goes into the check copy");
  const tab = monthTab(hw, "My home", "2026-10-03");
  eq(cell(tab, 2, "Бичсэн"), "check:My/2026/10/15!W4 ← 2 · Haier 300l (гараар бичсэнтэй таарсан)", "it names the copy's row");
  // a row put into the workbook above Haier 300l after the copy was made: the copy keeps its rows
  half.cells.splice(3, 0, ["Aucma 1"]);
  const down = (key) => { const [row, column] = key.split(",").map(Number); return `${row >= 4 ? row + 1 : row},${column}`; };
  half.notes = Object.fromEntries(Object.entries(half.notes).map(([key, note]) => [down(key), note]));
  half.backgrounds = half.backgrounds.map((b) => (b.r >= 4 ? { ...b, r: b.r + 1 } : b));
  const book = monthBook(hw, "My home", "2026-10-03");
  const again = run(hw, false, `(() => { let applied = 0; const real = ocrApplyCellChanges; ocrApplyCellChanges = (...a) => { applied += 1; return real(...a); }; try { reviewOcrBook({ file: DriveApp.getFileById(${JSON.stringify(book.file.id)}), ...ocrMonthBookOf(${JSON.stringify(book.file.name)}) }, new Map(), ocrEditorContext(PropertiesService.getScriptProperties()), Infinity); } finally { ocrApplyCellChanges = real; } return applied; })()`);
  eq([again.error, again.value, cell(tab, 2, "Бичсэн")], [null, 0, "check:My/2026/10/15!W4 ← 2 · Haier 300l (гараар бичсэнтэй таарсан)"], "nothing changed: no cell written, the line as it was");
}

// ---------- the year registry read once a run (2026-10-29: one review read it 467 times from CacheService, 13 s) ----------
{
  const gw = makeWorld();
  const withRegistry = code.replace("const REGISTRY_ID = '';", "const REGISTRY_ID = 'registry';"); // the tests run without one
  const counted = (expr) => run(gw, true, `(() => { const real = CacheService; let reads = 0; CacheService = { getScriptCache: () => { const cache = real.getScriptCache(); return { ...cache, get: (key) => { reads += 1; return cache.get(key); } }; } }; ${expr}; return reads; })()`, withRegistry);
  const once = counted("readRegistry(); readRegistry(); ocrBookIdFor('Соронзон', 'Sor/2026/10/15'); ocrHasBooksFor(2027)");
  eq([once.error, once.value], [null, 1], "the registry is read once a run, however often it is asked for");
  eq(counted("readRegistry()").value, 1, "and again by the next run");
}

// ---------- the month and day folders looked at once a day, not every hour (2026-10-29: 13 s a pass) ----------
{
  const fw = makeWorld();
  fw.props.GEMINI_KEY = "k";
  const today = run(fw, true, "ocrStamp(new Date(), 'yyyy-MM-dd')").value;
  fw.props["ocr-folders-checked"] = today;
  const checked = run(fw, false, "processNewReceipts()");
  ok(checked.error === null && /сар бэлдсэн 0/.test(checked.logs) && fw.root.find("Соронзон", "орлого", "2026-10") === null, "looked at today: no folder looked at again");
  delete fw.props["ocr-folders-checked"];
  const fresh = run(fw, false, "processNewReceipts()");
  ok(/сар бэлдсэн 28/.test(fresh.logs) && fw.props["ocr-folders-checked"] === today, "not yet today: the folders made, today kept");
  fw.props["ocr-folders-checked"] = "2026-01-01";
  eq(run(fw, true, "processNewReceipts()").error, null, "DRY_RUN looks too");
  eq(fw.props["ocr-folders-checked"], "2026-01-01", "but keeps nothing");
}

// ---------- every pass logs where its time went (2026-10-29: passes with no new photo took 280-294 s) ----------
{
  const sw = makeWorld();
  sw.props.GEMINI_KEY = "k";
  put(sw, DAY("My home", "зарлага", "03"), "S.HEIC", "s", 1, receipt(null, [item("Haier 300", "Haier 300l", 1, 1200)], { document_type: "zarlaga" }));
  const pass = run(sw, false, "processNewReceipts()");
  ok(pass.error === null && /OCR бэлтгэл: хавтас \d+ сек, барааны жагсаалт \d+ сек, зураг хайх \d+ сек/.test(pass.logs), `the start of a pass, by step: ${pass.logs}`);
  ok(/Шалгалтын бэлтгэл: шинэ барааны ангилал \d+ сек, хуучин файл \d+ сек, өөрчлөгдсөн файл хайх \d+ сек \(\d+ файл\)/.test(pass.logs), "the start of the review, by step");
  ok(/Шалгасан: .+ \d{4}-\d{2} \d+ сек/.test(pass.logs), "each spreadsheet reviewed, with its seconds");
  ok(/ \d{4}-\d{2}: бэлтгэл \d+ сек, \d+ өдөр \d+ сек, хадгалах \d+ сек/.test(pass.logs), "and where its seconds went: set-up, the day tabs, keeping");
  ok(/Шалгалт \(.*\): .*\(\d+ сек; ажлын номын мөр \d+ сек\)/.test(pass.logs), "the review's whole time and the workbook rows' look");
}

// ---------- the notes written before rewritten as the link to the line's row (the owner, 2026-10-25, 2026-10-26) ----------
{
  const pl = makeWorld();
  const half = pl.books[CATALOG_ID].sheets["My/2026/10/15"];
  half.cells[3][22] = 2; half.notes["4,23"] = "OCR: K.HEIC [k#1] = 2 @2026-10-03";
  half.cells[4][22] = 3; half.notes["5,23"] = "мамын тэмдэглэл";
  half.cells.push(["Panasonica 195l"]); half.cells[5][22] = 1; half.notes["6,23"] = "OCR: L.HEIC [l#1] = 1 @2026-10-03\nЗураг: https://drive.google.com/file/d/l/view";
  const dry = run(pl, true, "addSheetLinksToOcrNotes()");
  ok(dry.error === null && half.notes["4,23"] === "OCR: K.HEIC [k#1] = 2 @2026-10-03" && /2 нүд/.test(dry.logs), "DRY_RUN only counts the notes it would change");
  eq(run(pl, false, "addSheetLinksToOcrNotes()").error, null, "rewriting the notes works");
  eq([half.notes["4,23"], half.notes["5,23"], half.notes["6,23"], half.value(4, 23)],
     ["Зураг: https://drive.google.com/file/d/k/view#1 = 2 (2026-10-03)", "мамын тэмдэглэл", "Зураг: https://drive.google.com/file/d/l/view#1 = 1 (2026-10-03)", 2],
     "an old OCR note, with or without its link line, becomes a link with the count (the photo's, its row not found); a note of mom's stays; the number stays");
  const again = run(pl, false, "addSheetLinksToOcrNotes()");
  ok(/0 нүд/.test(again.logs), "run again: nothing left to change");
  ok(pl.props["ocr-notes-sheet-linked"] === undefined && /Бүх tab дууслаа/.test(again.logs), "a run that got through every tab forgets where it was");
  // a run cut off by the time: the next one skips the tabs done
  half.notes["4,23"] = "OCR: K.HEIC [k#1] = 2 @2026-10-03";
  pl.props["ocr-notes-sheet-linked"] = JSON.stringify([`${CATALOG_ID}|My/2026/10/15`]);
  const resumed = run(pl, false, "addSheetLinksToOcrNotes()");
  eq([half.notes["4,23"], /0 нүд/.test(resumed.logs), pl.props["ocr-notes-sheet-linked"]], ["OCR: K.HEIC [k#1] = 2 @2026-10-03", true, undefined], "a tab done before is not read again");
}

// ---------- a note of the photo's kind linked to the line's row: by the rewrite, by a review (2026-10-26) ----------
{
  const sw = makeWorld();
  sw.props.GEMINI_KEY = "k";
  const e3 = put(sw, DAY("My home", "зарлага", "03"), "E3.HEIC", "e3", 1, receipt(null, [item("Haier 300", "Haier 300l", 2, 1200)], { document_type: "zarlaga" }));
  eq(run(sw, false, "processNewReceipts()").error, null, "a receipt is written");
  const my = sw.books[CATALOG_ID].sheets["My/2026/10/15"];
  const t3 = monthTab(sw, "My home", "2026-10-03");
  const linked = `${rowLink(t3, 2, `${e3.id}#1`)} = 2 (2026-10-03)`;
  const photo = `Зураг: https://drive.google.com/file/d/${e3.id}/view#1 = 2 (2026-10-03)`;
  eq(my.notes["4,23"], linked, "a new line's note links to its row");
  my.notes["4,23"] = photo;
  eq(run(sw, false, "addSheetLinksToOcrNotes()").error, null, "the rewrite works");
  eq([my.notes["4,23"], my.value(4, 23)], [linked, 2], "a note of the photo's kind is linked to the row its ID is in; the number stays");
  my.notes["4,23"] = photo;
  eq(run(sw, false, `reviewOcrBookNow("${t3.getParent().getId()}")`).error, null, "a review of the month spreadsheet works");
  eq(my.notes["4,23"], linked, "a review links a note of the photo's kind to the line's row too");
}

// ---------- a line left in a cell its "Бичсэн" does not name is taken out (the owner, 2026-10-26) ----------
{
  const ow = makeWorld();
  ow.props.GEMINI_KEY = "k";
  ow.today = "2026-10-24";
  const oTos = ow.books[TOSGON_ID].sheets["Tos/2026/10"];
  put(ow, ["Тосгон", "тооллого", "2026-10", "23"], "T1.HEIC", "t1", 2, receipt(null, [item("Haier 300", "Haier 300l", 4, null), item("Haier 400", "Haier 400l", 2, null)], { document_type: "toollogo" }));
  eq(run(ow, false, "processNewReceipts()").error, null, "a count of two products is written");
  const t23 = monthTab(ow, "Тосгон", "2026-10-23");
  const TF = run(ow, true, `ocrColumnNumber("${cell(t23, 2, "Бичсэн").match(/!([A-Z]+)\d+ /)[1]}")`).value; // the block of the 23rd
  const rowOf = (name) => oTos.cells.findIndex((r) => r[0] === name) + 1;
  const [r300, r400, r501, rX] = ["Haier 300l", "Haier 400l", "Meiling 501l", "xinix 240l "].map(rowOf);
  const line300 = oTos.notes[`${r300},${TF}`], line400 = oTos.notes[`${r400},${TF}`];
  eq([oTos.value(r300, TF), oTos.value(r400, TF)], [4, 2], "each in its own row of the block");
  // the rows were put in order without the block: Haier 300l's line stands in other rows too
  oTos.cells[r501 - 1][TF - 1] = 4; oTos.notes[`${r501},${TF}`] = line300;
  oTos.cells[r400 - 1][TF - 1] = 6; oTos.notes[`${r400},${TF}`] = `${line400}\n${line300}`;
  oTos.cells[rX - 1][TF - 1] = 9; oTos.notes[`${rX},${TF}`] = line300; // changed by hand: 9, not the note's 4
  const bookId = t23.getParent().getId();
  const dry = run(ow, true, `reviewOcrBookNow("${bookId}")`);
  ok(dry.error === null && /өөр нүдэнд бичигдсэн мөрийг хаслаа.*туршилт/.test(dry.logs) && oTos.value(r501, TF) === 4, "DRY_RUN names the stray lines and changes nothing");
  eq(run(ow, false, `reviewOcrBookNow("${bookId}")`).error, null, "the review works");
  eq([oTos.value(r501, TF), oTos.notes[`${r501},${TF}`] || "", bgOf(oTos, r501, TF)], ["", "", null], "a cell holding only a stray line is emptied, its colour too");
  eq([oTos.value(r400, TF), oTos.notes[`${r400},${TF}`]], [2, line400], "a stray line beside a right one is taken out: the cell is the right one's again");
  eq([oTos.value(r300, TF), oTos.notes[`${r300},${TF}`]], [4, line300], "the line stays where its \"Бичсэн\" says");
  eq([oTos.value(rX, TF), oTos.notes[`${rX},${TF}`]], [9, line300], "a cell changed by hand is left as it is");
  ok(!/өөр нүдэнд бичигдсэн/.test(run(ow, false, `reviewOcrBookNow("${bookId}")`).logs), "a second review finds nothing more");
  // the only copy of a line, in another cell than its "Бичсэн" names (an old month: the row moved
  // with its cell, the line waits on a failed write and is not written again): it stays
  const rOther = rowOf("Meiling 501l");
  setCell(t23, 3, "Бичсэн", `✗ огноо алга [k:${vm.runInContext("OCR_WRITE_RULES", pure)}||Haier 400l|2]\nTos/2026/10!${run(ow, true, `columnLetter(${TF})`).value}${r400} ← 2 · Haier 400l`);
  oTos.cells[r400 - 1][TF - 1] = ""; oTos.notes[`${r400},${TF}`] = "";
  oTos.cells[rOther - 1][TF - 1] = 2; oTos.notes[`${rOther},${TF}`] = line400;
  run(ow, false, `reviewOcrBookNow("${bookId}")`);
  eq([oTos.value(rOther, TF), oTos.notes[`${rOther},${TF}`], oTos.value(r400, TF) || ""], [2, line400, ""], "a line that is nowhere else is never taken out");
  oTos.notes[`${r400},${TF}`] = line400; // the cell "Бичсэн" names was emptied by hand, its note left (My/2026/9/30!F101)
  run(ow, false, `reviewOcrBookNow("${bookId}")`);
  eq([oTos.value(rOther, TF), oTos.notes[`${rOther},${TF}`]], [2, line400], "nor one whose named cell was changed by hand: that cell holds no sure line");
}

// ---------- mom's product has no row on the sheet: the line still leaves the wrong product's cell (the owner, 2026-10-28) ----------
// Gemini read "Sharp 102l", mom chose "Sharp 100l" (Зөв), which Tos/2026/10 has no row for: the write
// failed and, until then, left the line in Sharp 102l's cell (pink) for good.
{
  const fw = makeWorld();
  fw.props.GEMINI_KEY = "k";
  fw.today = "2026-10-24";
  const fTos = fw.books[TOSGON_ID].sheets["Tos/2026/10"];
  put(fw, ["Тосгон", "тооллого", "2026-10", "23"], "F1.HEIC", "f1", 2, receipt(null, [item("Haier 300", "Haier 300l", 4, null)], { document_type: "toollogo" }));
  eq(run(fw, false, "processNewReceipts()").error, null, "a count line is written");
  const t23 = monthTab(fw, "Тосгон", "2026-10-23");
  const written = cell(t23, 2, "Бичсэн");
  const TF = run(fw, true, `ocrColumnNumber("${written.match(/!([A-Z]+)\d+ /)[1]}")`).value;
  const r300 = fTos.cells.findIndex((r) => r[0] === "Haier 300l") + 1;
  const note300 = fTos.notes[`${r300},${TF}`];
  eq(fTos.value(r300, TF), 4, "in Haier 300l's row");
  // since 2026-10-29 a product mom approved is put in (ocr-new-product.gs): here the list spells it
  // "Sharp 100", so it is not, and the line waits for mom's pick
  fw.books[CATALOG_ID].add("Бараа", new Sheet([[], [], [], ["Барааны нэр", "Аль дэлгүүрт"], ["хөлдөөгч"], ["Sharp 100", "Тосгон"]]));
  setCell(t23, 2, "Бараа", "Sharp 100l");
  setCell(t23, 2, "Ээж: Зөв/Буруу", "Зөв");
  const bookId = t23.getParent().getId();
  eq(run(fw, false, `reviewOcrBookNow("${bookId}")`).error, null, "the review works");
  const after = cell(t23, 2, "Бичсэн");
  eq([fTos.value(r300, TF) || "", fTos.notes[`${r300},${TF}`] || "", /^✗ "Sharp 100l" Tos\/2026\/10 дээр алга/.test(after), /Tos\/2026\/10!/.test(after)], ["", "", true, false],
     "the line leaves Haier 300l's cell though Sharp 100l has no row: the wrong product counts nothing, and the line waits for its row");
  // the same line failed under the rules before, its old cell still holding it (the real tab 23)
  fTos.cells[r300 - 1][TF - 1] = 4; fTos.notes[`${r300},${TF}`] = note300;
  setCell(t23, 2, "Бичсэн", `✗ "Sharp 100l" Tos/2026/10 дээр алга [k:v2|Зөв|Sharp 100l|4]\n${written}`);
  eq(run(fw, false, `reviewOcrBookNow("${bookId}")`).error, null, "the next review");
  eq([fTos.value(r300, TF) || "", fTos.notes[`${r300},${TF}`] || "", /Tos\/2026\/10!/.test(cell(t23, 2, "Бичсэн"))], ["", "", false], "a line failed under the rules before is tried again and leaves the wrong cell too");
  const quiet = run(fw, false, `reviewOcrBookNow("${bookId}")`);
  eq([fTos.value(r300, TF) || "", /мөр 2:/.test(quiet.logs)], ["", false], "and nothing comes back; a review that still finds no row writes and says nothing");
  // once the sheet has a row for mom's product, the next review writes the line there
  const rNew = fTos.cells.findIndex((r) => r[0] === "xinix 240l ") + 1;
  fTos.cells[rNew - 1][0] = "Sharp 100l";
  eq(run(fw, false, `reviewOcrBookNow("${bookId}")`).error, null, "a review after the row was added");
  eq([fTos.value(rNew, TF), cell(t23, 2, "Бичсэн").startsWith("✗"), fTos.value(r300, TF) || ""], [4, false, ""], "the line is written into Sharp 100l's new row, and only there");
}

// ---------- a tab of the 2026-10-24 layout, moved into the short one (2026-10-25) ----------
{
  const uw = makeWorld();
  const LAYOUT_1024 = ["Ээж", "Баримт", "Холбоос", "Хавтас (ээж хийсэн)", "Гарчиг", "Төрөл", "Огноо", "Бараа (уншсан)", "Бараа", "Тоо", "Нэгж үнэ", "Зарах үнэ / Тосгон өртөг",
    "Ээж: Зөв/Буруу", "Төлөв", "Шалгах шалтгаан", "Бичсэн", "Ээж бичсэн", "Бараа (тааруулсан)", "Тоо (уншсан)", "Итгэл", "Хаана орох", "Уншсан цаг", "ID"];
  const old = (values) => LAYOUT_1024.map((title) => values[title] ?? "");
  const tab = new Sheet([LAYOUT_1024,
    old({ "Баримт": "P.HEIC", "Хавтас (ээж хийсэн)": "Тосгон / тооллого", "Огноо": "2026-10-23", "Бараа (уншсан)": "Panasonic 250", "Бараа": "Panasonic 250l", "Тоо": 1, "Нэгж үнэ": 1100000,
      "Ээж: Зөв/Буруу": "Зөв", "Төлөв": "Шалгах", "Бичсэн": "Tos/2026/10!BY117 ← 1 · Panasonic 250l", "Бараа (тааруулсан)": "Panasonic 250l", "Тоо (уншсан)": 1, "Итгэл": 0.9, "ID": "p#1" }),
    old({ "Баримт": "P.HEIC", "Хавтас (ээж хийсэн)": "Тосгон / тооллого", "Бараа (уншсан)": "Ha", "Бараа": "Хамааралгүй", "Тоо": 0, "Бараа (тааруулсан)": "Ha", "Тоо (уншсан)": 1, "ID": "p#2" })]);
  tab.insertImage("img", LAYOUT_1024.length + 2, 2);
  uw.books.u = new Book("u", "OCR Тосгон тооллого 2026-10", { "23": tab });
  eq(run(uw, false, "ocrUpgradeTab(SpreadsheetApp.openById('u').getSheetByName('23'))").value, true, "a tab of the 2026-10-24 layout is moved");
  const at = (r, title) => tab.value(r, header.indexOf(title) + 1);
  eq(tab.cells[0].slice(0, header.length), ["Ээж", ...header.slice(1)], "into the short layout, mom's view kept");
  eq(["Бараа", "Тоо", "Ээж: Зөв/Буруу", "Бичсэн", "Бараа (тааруулсан)", "Тоо (уншсан)", "Хавтас (ээж хийсэн)", "Баримт", "ID"].map((title) => at(2, title)),
     ["Panasonic 250l", 1, "Зөв", "Tos/2026/10!BY117 ← 1 · Panasonic 250l", "Panasonic 250l", 1, "Тосгон / тооллого", "P.HEIC", "p#1"], "mom's mark, product and count, where it was written and the ID stay");
  eq(["Бараа", "Тоо", "Бараа (тааруулсан)"].map((title) => at(3, title)), ["Хамааралгүй", 0, "Ha"], "mom's Хамааралгүй and her 0 stay");
  const isLeftOver = (v, j) => j >= header.length && j !== header.length && v !== "" && v !== undefined; // O (index header.length) is the legend
  eq([tab.images.map((image) => image.column), tab.cells.every((row) => !row.some(isLeftOver))], [[header.indexOf("Зураг") + 1], true],
     "the photo moves beside the new last column, into N; nothing of the removed columns is left");
}

// ---------- a day tab read again with the current rules (2026-10-25) ----------
{
  const rw2 = makeWorld();
  rw2.props.GEMINI_KEY = "k";
  const countDay = ["Тосгон", "тооллого", "2026-10", "23"];
  const a = put(rw2, countDay, "A.jpeg", "photo-a", 1, receipt(null, [item("Haier 400", "Haier 400l", 1, null), item("Panasonic 4L", "Haier 300l", 16, null), item("Haier 300", "Haier 300l", 2, null)], { document_type: "other" }));
  const b = put(rw2, countDay, "B.jpeg", "photo-b", 2, { code: 503, text: "busy" });
  const line = (photo, n, values) => header.map((title) => ({ "Хавтас (ээж хийсэн)": "Тосгон / тооллого", "Баримт": photo.name, "Төлөв": "Шалгах", "ID": `${photo.id}#${n}`, ...values }[title] ?? ""));
  const tab = new Sheet([["Ээж", ...header.slice(1)],
    line(a, 1, { "Бараа": "Haier 400l", "Тоо": 1, "Ээж: Зөв/Буруу": "Зөв", "Бараа (тааруулсан)": "Haier 400l", "Тоо (уншсан)": 1, "Бичсэн": "Tos/2026/10!BY117 ← 1 · Haier 400l ✓" }),
    line(a, 2, { "Бараа": "Sevenstar 11", "Тоо": 16, "Ээж: Зөв/Буруу": "Зөв", "Бараа (тааруулсан)": "Sevenstar 11", "Тоо (уншсан)": 16 }),
    line(a, 3, { "Бараа": "Haier 300l", "Тоо": 3, "Бараа (тааруулсан)": "Haier 300l", "Тоо (уншсан)": 2, "Бичсэн": "Tos/2026/10!BY9 ← 3 · Haier 300l ✓" }),
    line(a, 4, { "Бараа": "Haier 200l", "Тоо": 1, "Бараа (тааруулсан)": "Haier 200l", "Тоо (уншсан)": 1, "Бичсэн": "Tos/2026/10!BY5 ← 1 · Haier 200l" }),
    line(b, 1, { "Бараа": "Meiling 501l", "Тоо": 5, "Бараа (тааруулсан)": "Meiling 501l", "Тоо (уншсан)": 5, "Бичсэн": "Tos/2026/10!BY7 ← 5 · Meiling 501l" })]);
  const book = rw2.root.path("OCR", "2026-10", "Тооллого").add(new File("OCR Тосгон тооллого 2026-10", "application/vnd.google-apps.spreadsheet", "", 1, "me"));
  rw2.books[book.id] = new Book(book.id, book.name, { "23": tab });
  const dry = run(rw2, true, `rereadOcrDay("${book.id}", "23")`);
  eq([plain(dry.value), rw2.geminiCalls.length, rowsOf(tab)], [{ photos: 2, read: 0 }, 0, 6], "DRY_RUN says what it would read, asks Gemini nothing and changes nothing");
  const live = run(rw2, false, `rereadOcrDay("${book.id}", "23")`);
  eq([live.error, plain(live.value), rw2.geminiCalls.map((c) => c.bytes)], [null, { photos: 2, read: 1, rows: 5 }, ["photo-a", "photo-b"]], "each photo of the tab is asked once more");
  ok(/A\.jpeg: \{"document_type":"other"/.test(live.logs), "Gemini's answer goes to the log");
  const lineRows = tab.cells.map((row, index) => [row, index + 1]).filter(([row], index) => index > 0 && (row || []).slice(0, header.length).some((v) => v !== "" && v !== undefined)).map(([, number]) => number);
  const at = (n, title) => tab.value(lineRows[n - 1], header.indexOf(title) + 1); // the n-th line, blank rows skipped
  eq([1, 2, 3, 4, 5].map((r) => [at(r, "ID"), at(r, "Бараа"), at(r, "Тоо"), at(r, "Ээж: Зөв/Буруу")]), [
    [`${a.id}#1`, "Haier 400l", 1, "Зөв"],
    [`${a.id}#2`, "Sevenstar 11", 16, "Зөв"],
    [`${a.id}#3`, "Haier 300l", 3, ""],
    [`${a.id}#4`, "Хамааралгүй", 1, ""],
    [`${b.id}#1`, "Meiling 501l", 5, ""],
  ], "a line mom said Зөв to keeps her product, count and mark whatever the new reading says (the owner, 2026-10-25); her own count stays; a line no longer read is Хамааралгүй; a photo Gemini could not read keeps its rows");
  eq([1, 3, 4].map((r) => at(r, "Бичсэн")), ["Tos/2026/10!BY117 ← 1 · Haier 400l ✓", "Tos/2026/10!BY9 ← 3 · Haier 300l ✓", "Tos/2026/10!BY5 ← 1 · Haier 200l"],
     "where each ID was written goes with it, so the next review moves or takes out its cell entry");
  eq([at(1, "Шалгах шалтгаан"), String(at(4, "Шалгах шалтгаан")).startsWith("дахин уншихад энэ мөр гараагүй")], ["", true], "the count folder is trusted now; a line gone says so");
  eq([tab.images.map((image) => image.row), lineRows], [[2, 30], [2, 3, 4, 5, 30]], "each photo beside its lines again, the next photo's below the last one's (blank rows between)");
  eq(plain(vm.runInContext(`ocrCarryOver({ "Бараа": "Haier 300l", "Тоо": 1, "Ээж: Зөв/Буруу": "" }, { "Бараа": "", "Тоо": 1, "Тоо (уншсан)": 1, "Ээж: Зөв/Буруу": "Зөв", "Бичсэн": "" })`, pure))["Бараа"], "Haier 300l",
     "a Зөв on a line that had no product keeps no empty product: the new reading's goes in");
}

// ---------- the tabs whose slips were read before "Төв" was known, read again (the owner, 2026-10-28) ----------
{
  const sw = makeWorld();
  sw.props.GEMINI_KEY = "k";
  const unknown = "хавтас буруу байж магадгүй: зарлагын баримт, илгээгч тодорхойгүй (Тосгоноос ирсэн бол Тосгоноос Соронзон руу хавтсанд); үнэ байхгүй";
  const day16 = ["Соронзон", "орлого", "2026-10", "16"];
  const day17 = ["Соронзон", "орлого", "2026-10", "17"];
  const s = put(sw, day16, "S.jpeg", "photo-s", 1, receipt(null, [item("Haier 300", "Haier 300l", 2, null)], { document_type: "zarlaga", header_sender: "Тов", header_receiver: "Соронзон" }));
  const t = put(sw, day17, "T.jpeg", "photo-t", 2, receipt(null, [item("Haier 400", "Haier 400l", 1, null)]));
  const line = (photo, values) => header.map((title) => ({ "Хавтас (ээж хийсэн)": "Соронзон / орлого", "Баримт": photo.name, "Төлөв": "Шалгах", "ID": `${photo.id}#1`, ...values }[title] ?? ""));
  const tab16 = new Sheet([["Ээж", ...header.slice(1)], line(s, { "Бараа": "Haier 300l", "Тоо": 2, "Бараа (тааруулсан)": "Haier 300l", "Тоо (уншсан)": 2, "Шалгах шалтгаан": unknown })]);
  const tab17 = new Sheet([["Ээж", ...header.slice(1)], line(t, { "Бараа": "Haier 400l", "Тоо": 1, "Бараа (тааруулсан)": "Haier 400l", "Тоо (уншсан)": 1, "Төлөв": "OK" })]);
  const book = sw.root.path("OCR", "2026-10", "Орлого").add(new File("OCR Соронзон орлого 2026-10", "application/vnd.google-apps.spreadsheet", "", 1, "me"));
  sw.books[book.id] = new Book(book.id, book.name, { "16": tab16, "17": tab17 });
  const dry = run(sw, true, `rereadUnknownSenderDays("${book.id}")`);
  eq([plain(dry.value), sw.geminiCalls.length], [{ days: ["16"], left: [] }, 0], "DRY_RUN names the tabs whose sender was unknown and asks Gemini nothing");
  const live = run(sw, false, `rereadUnknownSenderDays("${book.id}")`);
  eq([live.error, plain(live.value), sw.geminiCalls.map((c) => c.bytes)], [null, { days: ["16"], left: [] }, ["photo-s"]], "only the tab whose sender was unknown is read again");
  const at = (title) => tab16.value(2, header.indexOf(title) + 1);
  eq([at("Хавтас (ээж хийсэн)"), at("Шалгах шалтгаан"), at("Төлөв")], ["Соронзон / орлого (Тосгоноос)", "", "OK"], "a slip from \"Тов\" now counts as Тосгон's transfer: no folder or price reason, OK");
  eq(plain(run(sw, false, `rereadUnknownSenderDays("${book.id}")`).value), { days: [], left: [] }, "a second run finds nothing left to read");
}
{
  const hw = makeWorld();
  hw.props.GEMINI_KEY = "k";
  const unknown = "хавтас буруу байж магадгүй: зарлагын баримт, илгээгч тодорхойгүй (Тосгоноос ирсэн бол Тосгоноос My home руу хавтсанд); үнэ байхгүй";
  const h = put(hw, ["My home", "орлого", "2026-10", "24"], "H.jpeg", "photo-h", 1, receipt(null, [item("Haier 300", "Haier 300l", 2, null)], { document_type: "zarlaga", header_sender: "Хотоос", header_receiver: "My home" }));
  const row = header.map((title) => ({ "Хавтас (ээж хийсэн)": "My home / орлого", "Баримт": h.name, "Төлөв": "Шалгах", "ID": `${h.id}#1`, "Бараа": "Haier 300l", "Тоо": 2, "Бараа (тааруулсан)": "Haier 300l", "Тоо (уншсан)": 2, "Шалгах шалтгаан": unknown }[title] ?? ""));
  const tab24 = new Sheet([["Ээж", ...header.slice(1)], row]);
  const book = hw.root.path("OCR", "2026-10", "Орлого").add(new File("OCR My home орлого 2026-10", "application/vnd.google-apps.spreadsheet", "", 1, "me"));
  hw.books[book.id] = new Book(book.id, book.name, { "24": tab24 });
  const live = run(hw, false, `rereadUnknownSenderDays("${book.id}")`);
  const at = (title) => tab24.value(2, header.indexOf(title) + 1);
  eq([live.error, at("Хавтас (ээж хийсэн)"), at("Шалгах шалтгаан"), at("Төлөв")], [null, "My home / орлого (Хотоос)", "", "OK"], "a slip from the city read before the rule (noted, sender unknown) takes the three sides when read again");
  ok(/rereadMyHomeIncome10/.test(code) && /OCR_MY_HOME_INCOME_1023\.bookId\)/.test(code.slice(code.indexOf("function rereadMyHomeIncome10"))), "My home's 2026-10 income has its own re-read too");
}

// ---------- every tab of a month read under rules since changed, read again (the owner, 2026-10-28) ----------
{
  const mw = makeWorld();
  mw.props.GEMINI_KEY = "k";
  const unknown = "хавтас буруу байж магадгүй: зарлагын баримт, илгээгч тодорхойгүй (Тосгоноос ирсэн бол Тосгоноос Соронзон руу хавтсанд); үнэ байхгүй";
  const s = put(mw, ["Соронзон", "орлого", "2026-10", "16"], "S.jpeg", "photo-s", 1, receipt(null, [item("Haier 300", "Haier 300l", 2, null)], { document_type: "zarlaga", header_sender: "Төв" }));
  const t = put(mw, ["Соронзон", "орлого", "2026-10", "17"], "T.jpeg", "photo-t", 2, receipt(null, [item("Haier 400", "Haier 400l", 1, null)]));
  const z = put(mw, ["My home", "зарлага", "2026-10", "01"], "Z.jpeg", "photo-z", 3, receipt(null, [item("Haier 200", "Haier 200l", 1, null)], { document_type: "zarlaga" }));
  const line = (photo, folder, values) => header.map((title) => ({ "Хавтас (ээж хийсэн)": folder, "Баримт": photo.name, "Төлөв": "Шалгах", "ID": `${photo.id}#1`, ...values }[title] ?? ""));
  const sheetOf = (...rows) => new Sheet([["Ээж", ...header.slice(1)], ...rows]);
  const income16 = sheetOf(line(s, "Соронзон / орлого", { "Бараа": "Haier 300l", "Тоо": 2, "Бараа (тааруулсан)": "Haier 300l", "Тоо (уншсан)": 2, "Шалгах шалтгаан": unknown }));
  const income17 = sheetOf(line(t, "Соронзон / орлого", { "Бараа": "Haier 400l", "Тоо": 1, "Бараа (тааруулсан)": "Haier 400l", "Тоо (уншсан)": 1, "Шалгах шалтгаан": "брэнд зөрсөн (Haier 4 ≠ Haier 400l)" }));
  const sales01 = sheetOf(line(z, "My home / зарлага", { "Бараа": "Haier 200l", "Тоо": 1, "Бараа (тааруулсан)": "Haier 200l", "Тоо (уншсан)": 1, "Шалгах шалтгаан": "нэр жагсаалтад таараагүй; итгэл бага" }));
  const book = (folder, name, tabs) => {
    const file = mw.root.path("OCR", "2026-10", folder).add(new File(name, "application/vnd.google-apps.spreadsheet", "", 1, "me"));
    mw.books[file.id] = new Book(file.id, name, tabs);
    return file;
  };
  book("Орлого", "OCR Соронзон орлого 2026-10", { "16": income16, "17": income17 });
  book("Зарлага", "OCR My home зарлага 2026-10", { "01": sales01 });
  book("Орлого", "OCR Соронзон орлого 2026-09", { "16": sheetOf(line(s, "Соронзон / орлого", { "Шалгах шалтгаан": unknown })) });
  const tabsOf = (value) => plain(value);
  const stale = ["OCR My home зарлага 2026-10 / 01", "OCR Соронзон орлого 2026-10 / 16"];
  const dry = run(mw, true, `rereadStaleMonth("2026-10")`);
  eq([tabsOf(dry.value), mw.geminiCalls.length, mw.triggers.length], [{ read: [], left: stale }, 0, 0],
     "DRY_RUN names the month's tabs whose reasons came from rules since changed (the sender unknown, a confidence without its number), not a tab of mom's only, nor another month");
  const first = run(mw, false, `rereadStaleMonth("2026-10", 1)`);
  eq([first.error, tabsOf(first.value), mw.geminiCalls.map((c) => c.bytes)], [null, { read: [stale[0]], left: [stale[1]] }, ["photo-z"]], "out of time after one tab, the rest is left");
  eq(mw.triggers.map((trigger) => trigger.getHandlerFunction()), ["rereadOctober"], "and a run of its own is set to go on with them");
  const second = run(mw, false, `rereadStaleMonth("2026-10")`);
  eq([second.error, tabsOf(second.value), mw.geminiCalls.map((c) => c.bytes)], [null, { read: [stale[1]], left: [] }, ["photo-z", "photo-s"]], "the next run reads the rest only");
  eq(mw.triggers.length, 0, "and, nothing left, takes its run away");
  const reason = (tab) => tab.value(2, header.indexOf("Шалгах шалтгаан") + 1);
  ok(!/илгээгч тодорхойгүй/.test(reason(income16)) && !/итгэл бага(?! \()/.test(reason(sales01)), "the old reasons are gone");
  eq(tabsOf(run(mw, false, `rereadStaleMonth("2026-10")`).value), { read: [], left: [] }, "a tab read once is not read again, whatever it still says");
}

// ---------- one month spreadsheet reviewed now, out of the pass's order (the owner, 2026-10-25) ----------
{
  const nw = makeWorld();
  const oldHeader = header.filter((title) => title !== "Итгэл"); // the 12 columns before "Итгэл"
  const p = put(nw, ["Тосгон", "тооллого", "2026-10", "23"], "P.jpeg", "photo-p", 1, receipt(null, [item("Haier 300", "Haier 300l", 1, null)]));
  const row = (values) => oldHeader.map((title) => values[title] ?? "");
  const tab = new Sheet([["Ээж", ...oldHeader.slice(1)],
    row({ "Бараа": "Haier 300l", "Тоо": 1, "Төлөв": "Шалгах", "Хавтас (ээж хийсэн)": "Тосгон / тооллого", "Баримт": "P.jpeg", "Бараа (тааруулсан)": "Haier 300l", "Тоо (уншсан)": 1, "ID": `${p.id}#1` })]);
  const book = nw.root.path("Зураг уншуулах", "2026-10", "Тооллого").add(new File("OCR Тосгон тооллого 2026-10", "application/vnd.google-apps.spreadsheet", "", 1, "me"));
  nw.books[book.id] = new Book(book.id, book.name, { "23": tab });
  const before = JSON.stringify(tab.cells);
  run(nw, true, `reviewOcrBookNow("${book.id}")`);
  eq(JSON.stringify(tab.cells), before, "DRY_RUN changes nothing");
  const done = run(nw, false, `reviewOcrBookNow("${book.id}")`);
  eq(done.error, null, "the review of one month spreadsheet runs from the editor");
  eq(tab.cells[0].slice(0, header.length), ["Ээж", ...header.slice(1)], "its tabs are moved into the current columns (\"Итгэл\" among them)");
  eq(tab.value(2, header.indexOf("ID") + 1), `${p.id}#1`, "every line keeps its ID");
  eq([tab.value(1, header.length + 1), tab.value(1, 19) || ""], ["Өнгө:", ""], "the colour legend is written in O, not in S1");
  ok(/Тосгон тооллого 2026-10/.test(done.logs), "the log names the spreadsheet");
  eq(run(nw, true, `reviewOcrBookNow("${nw.root.allFiles().find((f) => f.name === "P.jpeg").id}")`).error !== null, true, "a file that is no month spreadsheet is refused");
}

// ---------- a tab of before: tall rows made of the default height, room left under each photo (the owner, 2026-10-26) ----------
{
  const N = header.indexOf("Зураг") + 1; // the photo column (M since 2026-10-28)
  const line = (photo, n, name) => header.map((title) => ({ "Бараа": name, "Тоо": 1, "Төлөв": "OK", "Баримт": photo, "ID": `${photo}#${n}` }[title] ?? ""));
  // receipt A: one line, its row as tall as the photo; receipt B: two lines, the last one tall
  const tab = new Sheet([["Ээж", ...header.slice(1)], line("A", 1, "Haier 200l"), line("B", 1, "Midea 300l"), line("B", 2, "Hisense 200l")]);
  tab.notes[`2,${header.indexOf("Бараа") + 1}`] = "ээжийн тэмдэглэл";
  tab.insertImage("a", N, 2).setHeight(560); tab.insertImage("b", N, 3).setHeight(300);
  tab.setRowHeight(2, 572); tab.setRowHeight(4, 291);
  vm.runInContext("ocrSpreadPhotos", pure)(tab);
  const rowsA = Math.ceil(572 / 21); // 28
  eq([tab.value(2, 2), tab.value(2 + rowsA, 2), tab.value(3 + rowsA, 2)], ["Haier 200l", "Midea 300l", "Hisense 200l"], "the lines move down below the photo before them");
  eq(tab.images.map((image) => image.row), [2, 2 + rowsA], "each photo goes with its receipt");
  eq(Object.values(tab.heights).every((h) => h === 21), true, "every row as tall as any");
  eq(tab.notes[`2,${header.indexOf("Бараа") + 1}`], "ээжийн тэмдэглэл", "a cell's note stays with its line");
  eq(tab.validationSets, 1, "the lists of the columns are put back after the lines moved (a name the list lacks now is moved, not refused)");
  eq(vm.runInContext("ocrReadRecords", pure)(tab).map((r) => r._row).join(), [2, 2 + rowsA, 3 + rowsA].join(), "the records are read with their new rows");
  const again = JSON.stringify(tab.cells);
  vm.runInContext("ocrSpreadPhotos", pure)(tab);
  eq(JSON.stringify(tab.cells), again, "a tab done already is left as it is");
  // a column written back by records: the blank rows between them stay blank
  const records = vm.runInContext("ocrReadRecords", pure)(tab);
  vm.runInContext("ocrRewriteColumn", pure)(tab, "Шалгах шалтгаан", records, records.map((r) => ({ ...r, "Шалгах шалтгаан": `x${r._row}` })));
  const reasonAt = (row) => tab.value(row, header.indexOf("Шалгах шалтгаан") + 1);
  eq([reasonAt(2), reasonAt(3), reasonAt(2 + rowsA), reasonAt(3 + rowsA)], ["x2", "", `x${2 + rowsA}`, `x${3 + rowsA}`], "each value goes into its record's row");
}

// ---------- every month spreadsheet set up again now: photos in N, legend in O (the owner, 2026-10-26) ----------
{
  const sw = makeWorld();
  sw.props.GEMINI_KEY = "k";
  put(sw, ["Тосгон", "тооллого", "2026-10", "23"], "P.jpeg", "photo-p", 1, receipt(null, [item("Haier 300", "Haier 300l", 1, null)]));
  eq(run(sw, false, "processNewReceipts()").error, null, "a month spreadsheet made by a run");
  const tab = monthTab(sw, "Тосгон", "2026-10-23", "тооллого");
  const N = header.indexOf("Зураг") + 1, O = header.length + 1; // N: the photo column (M since 2026-10-28)
  // as a tab of 2026-10-25 stood: the photo in O, the legend in S1, set up already by that version
  tab.images.forEach((image) => { image.column = O; });
  for (let r = 1; r <= 7; r++) { tab.cells[r - 1] = tab.cells[r - 1] || []; tab.cells[r - 1][O - 1] = ""; }
  tab.cells[0][18] = "Өнгө:\n■ Шар мөр: …";
  const B = header.indexOf("Бараа") + 1, D = header.indexOf("Ээж: Зөв/Буруу") + 1;
  tab.getRange(1, B, 1000, 3).setBackground("#d9d2e9"); // the purple of before
  tab.getRange(2, D, 999, 1).setDataValidation({ list: ["Зөв", "Буруу"] }); // the list of before, on every row
  Object.keys(sw.props).filter((key) => key.startsWith("ocr-book-setup:")).forEach((key) => { sw.props[key] = "old"; });
  const rowsBefore = JSON.stringify(tabRows(sw, "Тосгон", "2026-10-23", "тооллого").map((row) => row.slice(0, header.length)));
  const dry = run(sw, true, "setUpAllOcrBooksNow()");
  eq([dry.error, tab.images.every((image) => image.column === O), /шинэ бүтэц рүү оруулна/.test(dry.logs)], [null, true, true], "DRY_RUN only says it would");
  const done = run(sw, false, "setUpAllOcrBooksNow()");
  eq([done.error, plain(done.value).left], [null, 0], "every month spreadsheet is gone through in one run here");
  eq([tab.images.every((image) => image.column === N), tab.value(1, O), tab.value(1, 19) || ""], [true, "Өнгө:", ""], "the photo moved to N, the legend in O, S1 emptied");
  eq([tab.getRange(1, B).getBackground(), tab.getRange(50, D).getBackground()], ["#ffffff", "#ffffff"], "the purple is taken away");
  eq(tab.validations.filter((v) => v.c === D).map((v) => [v.r, v.h]), [[2, 1]], "Зөв/Буруу only on the line, not on the rows below");
  eq(JSON.stringify(tabRows(sw, "Тосгон", "2026-10-23", "тооллого").map((row) => row.slice(0, header.length))), rowsBefore, "no receipt line changed");
  ok(Object.keys(sw.props).filter((key) => key.startsWith("ocr-book-setup:")).every((key) => sw.props[key] !== "old"), "each one is marked set up, so the next run skips it");
}

// ---------- the month spreadsheets of every kind, split by kind (2026-10-25) ----------
{
  const sw = makeWorld();
  const photoOf = (folder, name) => sw.root.path("Соронзон", folder, "2026-09", "03").add(new File(name, "image/heif", name, 1, "mom"));
  const [p1, p2, p3, p4] = [photoOf("зарлага", "Z.HEIC"), photoOf("орлого", "O.HEIC"), photoOf("Соронзоноос My home руу", "S.HEIC"), photoOf("тооллого", "T.HEIC")];
  const p5 = sw.root.path("Соронзон", "зарлага", "2026-09", "05").add(new File("Z5.HEIC", "image/heif", "Z5.HEIC", 1, "mom"));
  const line = (photo, folder, values = {}) => header.map((title) => ({ "Баримт": photo.name, "Хавтас (ээж хийсэн)": `Соронзон / ${folder}`, "Бараа": "Haier 300l", "Тоо": 1, "Төлөв": "OK", "ID": `${photo.id}#1`, ...values }[title] ?? ""));
  const oldFile = sw.root.path("OCR", "2026-09").add(new File("OCR Соронзон 2026-09", "application/vnd.google-apps.spreadsheet", "", 1, "me"));
  const day3 = [header, line(p1, "зарлага", { "Ээж: Зөв/Буруу": "Зөв", "Бичсэн": "Sor/2026/9/15!Z5 ← 1 · Haier 300l ✓" }), line(p2, "орлого", { "Ээж: Зөв/Буруу": "Буруу", "Бараа": "Хамааралгүй", "Тоо": 0, "Бараа (тааруулсан)": "Ha", "Тоо (уншсан)": 1 }), line(p3, "Соронзоноос My home руу"), line(p4, "тооллого"), header.map(() => "")];
  sw.books[oldFile.id] = new Book(oldFile.id, oldFile.name, { "03": new Sheet(day3), "05": new Sheet([header, line(p5, "зарлага")]), "_нэрс": new Sheet([["Haier 300l"]]) });
  Object.assign(sw.props, { [`ocr-open:${oldFile.id}:03`]: "{}", "ocr-open:other:01": "{}", [`ocr-book-setup:${oldFile.id}`]: "x", "ocr-cal:Соронзон|2026-09": "{}", "ocr-cal-link:Соронзон|2026-09": "{}" });
  sw.triggers.push({ getHandlerFunction: () => "onOcrViewEdit", getTriggerSourceId: () => oldFile.id }, { getHandlerFunction: () => "onOcrViewEdit", getTriggerSourceId: () => "other" });

  const dry = run(sw, true, "splitOcrBooksByKind()");
  ok(/03: 1 мөр -> Зарлага\/OCR Соронзон зарлага 2026-09 \(туршилт\)/.test(dry.logs) && /"OCR Соронзон 2026-09 \(хуучин\)" болж OCR\/2026-09\/хуучин руу орно/.test(dry.logs), "DRY_RUN tells where every line would go");
  ok(sw.root.find("Зураг уншуулах") === null, "a Drive whose folder is still named \"OCR\" is used as it is: no second folder is made");
  eq([oldFile.name, sw.root.find("OCR", "2026-09").folders.length, Object.keys(sw.props).filter((key) => key.startsWith("ocr-split:")).length], ["OCR Соронзон 2026-09", 0, 0], "and changes nothing");

  // a line already copied by a run cut off before (same ID) is not copied twice
  run(sw, false, `ocrAppendRows(ocrDayTab(ocrMonthBook(DriveApp.getFolderById(OCR_ROOT_FOLDER_ID), "Соронзон", OCR_BOOK_KINDS.zarlaga, "2026-09").book, "03"), [ocrRow({ "Баримт": "Z.HEIC", "ID": "${p1.id}#1", "Ээж: Зөв/Буруу": "Зөв" })])`);
  const live = run(sw, false, "splitOcrBooksByKind()");
  eq([live.error, plain(live.value)], [null, { done: 1, left: 0 }], "the live run splits the old spreadsheet");
  const ids = (word, day) => (tabRows(sw, "Соронзон", `2026-09-${day}`, word) || []).slice(1).map((row) => row[header.indexOf("ID")]);
  eq([ids("зарлага", "03"), ids("зарлага", "05"), ids("орлого", "03"), ids("шилжүүлэг", "03"), ids("тооллого", "03")], [[`${p1.id}#1`], [`${p5.id}#1`], [`${p2.id}#1`], [`${p3.id}#1`], [`${p4.id}#1`]],
     "every line into its kind's spreadsheet, the same day's tab; one already there not twice; the empty row left out");
  const t3 = monthTab(sw, "Соронзон", "2026-09-03", "орлого");
  eq([cell(t3, 2, "Хавтас (ээж хийсэн)"), t3.images.map((image) => image.row)], ["Соронзон / орлого", [2]], "with its photo beside it again");
  eq(["Ээж: Зөв/Буруу", "Бараа", "Тоо", "Бараа (тааруулсан)", "Тоо (уншсан)"].map((title) => cell(t3, 2, title)), ["Буруу", "Хамааралгүй", 0, "Ha", 1], "mom's mark, name and count go with the line");
  ok(sw.root.find("OCR", "2026-09", "Шилжүүлэг").files.some((f) => f.name === "OCR Соронзон шилжүүлэг 2026-09"), "a transfer goes to the Шилжүүлэг folder");
  const z5 = monthTab(sw, "Соронзон", "2026-09-05", "зарлага");
  eq([cell(z5, 2, "Баримт"), cell(z5, 2, "Төлөв")], ["Z5.HEIC", "OK"], "the read and its status go with it");
  const oldMoved = sw.root.find("OCR", "2026-09", "хуучин");
  eq([oldFile.name, oldFile.parent === oldMoved, Boolean(sw.books[oldFile.id])], ["OCR Соронзон 2026-09 (хуучин)", true, true], "the old spreadsheet is renamed and put into OCR/2026-09/хуучин, its id kept");
  eq(Object.keys(sw.props).filter((key) => key.includes(oldFile.id) || key.endsWith("|2026-09")).sort(), [], "what was kept about it is let go");
  ok("ocr-open:other:01" in sw.props && sw.triggers.map((t) => t.getTriggerSourceId()).join() === "other", "other spreadsheets keep theirs; its view trigger goes");
  const again = run(sw, false, "splitOcrBooksByKind()");
  eq([plain(again.value), ids("зарлага", "03").length], [{ done: 0, left: 0 }, 1], "a second run finds nothing left");
}
{
  // the site's calendar: a place's kinds put together; checks of all its spreadsheets
  const fw = makeWorld();
  Object.assign(fw.props, {
    OCR_FEED_TOKEN: "t",
    "ocr-cal:Соронзон|2026-10|орлого": JSON.stringify({ "02": { "орлого": [1, 2, 0, 2, 0, 0, 0] } }),
    "ocr-cal:Соронзон|2026-10|зарлага": JSON.stringify({ "02": { "зарлага": [3, 4, 1, 3, 0, 0, 0] }, "05": { "зарлага": [1, 1, 0, 1, 0, 0, 0] } }),
    "ocr-cal-link:Соронзон|2026-10|орлого": JSON.stringify({ id: "o1", days: { "02": 7 } }),
    "ocr-cal-link:Соронзон|2026-10|зарлага": JSON.stringify({ id: "z1", days: { "02": 8, "05": 9 } }),
  });
  const answer = plain(run(fw, true, `ocrFeedAnswer({ token: "t", mode: "receipts", month: "2026-10" }, PropertiesService.getScriptProperties(), new Date())`).value);
  eq([answer.places["Соронзон"], answer.links["Соронзон"], answer.places["Тосгон"], answer.links["Тосгон"], answer.months],
     [{ "02": { "зарлага": [3, 4, 1, 3, 0, 0, 0], "орлого": [1, 2, 0, 2, 0, 0, 0] }, "05": { "зарлага": [1, 1, 0, 1, 0, 0, 0] } }, { "зарлага": { id: "z1", days: { "02": 8, "05": 9 } }, "орлого": { id: "o1", days: { "02": 7 } } }, {}, null, ["2026-10"]],
     "the days of a place's kinds are put together for the site (their folders differ)");
  eq(plain(run(fw, true, `ocrMergeChecks([{ "02": [1] }, { "02": [2], "05": [3] }])`).value), { "02": [1, 2], "05": [3] }, "and so are their checks");
  // the web app keeps a month's answer in the script cache, so a visit does not open every month spreadsheet
  const cached = (token = "t") => plain(run(fw, true, `ocrFeedCachedAnswer({ token: "${token}", mode: "receipts", month: "2026-10" }, PropertiesService.getScriptProperties(), CacheService.getScriptCache(), new Date())`).value);
  eq(cached().places["Соронзон"]["05"], { "зарлага": [1, 1, 0, 1, 0, 0, 0] }, "the first visit of a month is answered from the properties and the spreadsheets");
  fw.props["ocr-cal:Соронзон|2026-10|зарлага"] = JSON.stringify({ "02": { "зарлага": [3, 4, 1, 3, 0, 0, 0] }, "05": { "зарлага": [1, 1, 0, 1, 0, 0, 0] }, "06": { "зарлага": [2, 2, 0, 0, 0, 0, 0] } });
  eq("06" in cached().places["Соронзон"], false, "the next visit within its time gets the kept answer, nothing is read again");
  eq(cached("x"), { ok: false, error: "forbidden" }, "a wrong token is refused, kept answer or not");
  fw.props["ocr-last-scan"] = "2026-10-28T01:00:00.000Z";
  eq([("06" in cached().places["Соронзон"]), cached().updated], [true, "2026-10-28T01:00:00.000Z"], "a review pass (a new last-scan stamp) makes the next visit read again");
  eq(plain(run(fw, true, `ocrFeedCachedAnswer({ token: "t", mode: "receipts", month: "2026-09" }, PropertiesService.getScriptProperties(), CacheService.getScriptCache(), new Date())`).value).places["Соронзон"], {}, "each month is kept apart");
}

// ---------- "Бүх мөр Зөв" in A1: a whole day marked Зөв at once (the owner, 2026-10-29) ----------
{
  const sw = makeWorld();
  const photo = sw.root.path("Соронзон", "зарлага", "2026-10", "22").add(new File("Z.HEIC", "image/heif", "Z.HEIC", 1, "mom"));
  const line = (n, values) => header.map((title) => ({ "Хавтас (ээж хийсэн)": "Соронзон / зарлага", "Баримт": photo.name, "Төлөв": "OK", "ID": `${photo.id}#${n}`, ...values }[title] ?? ""));
  const tab = new Sheet([["Бүгд", ...header.slice(1)],
    line(1, { "Бараа": "Haier 300l", "Тоо": 1, "Бараа (тааруулсан)": "Haier 300l", "Тоо (уншсан)": 1 }),
    line(2, { "Бараа": "Haier 400l", "Тоо": 2, "Бараа (тааруулсан)": "Haier 400l", "Тоо (уншсан)": 2, "Төлөв": "Шалгах", "Шалгах шалтгаан": "итгэл бага (0.5)" }),
    line(3, { "Бараа": "Meiling 501l", "Тоо": 1, "Бараа (тааруулсан)": "Meiling 501l", "Тоо (уншсан)": 1, "Ээж: Зөв/Буруу": "Буруу" }),
    line(4, { "Бараа": "Haier 300l", "Тоо": 3, "Бараа (тааруулсан)": "Haier 300l", "Тоо (уншсан)": 1 }), // mom changed the count: approved already
    header.map(() => ""), // a blank row under the photo
  ]);
  const file = sw.root.path("OCR", "2026-10", "Зарлага").add(new File("OCR Соронзон зарлага 2026-10", "application/vnd.google-apps.spreadsheet", "", 1, "me"));
  sw.books[file.id] = new Book(file.id, file.name, { "22": tab });
  const D = header.indexOf("Ээж: Зөв/Буруу") + 1;
  const marks = () => [2, 3, 4, 5, 6].map((r) => tab.value(r, D));
  const pickInA1 = (dry, old) => {
    tab.cells[0][0] = "Бүх мөр Зөв";
    return run(sw, dry, `onOcrViewEdit({ value: "Бүх мөр Зөв", oldValue: ${JSON.stringify(old)}, range: { getRow: () => 1, getColumn: () => 1, getLastColumn: () => 1, getValue: () => "Бүх мөр Зөв", getSheet: () => SpreadsheetApp.openById("${file.id}").getSheetByName("22") } })`);
  };
  eq(plain(vm.runInContext("OCR_VIEW_CHOICES", pure)), ["Ээж", "Бүгд", "Бүх мөр Зөв"], "A1's list: the two views, then the whole day's Зөв");
  const dry = pickInA1(true, "Бүгд");
  eq([dry.error, marks(), tab.value(1, 1)], [null, ["", "", "Буруу", "", ""], "Бүгд"], "DRY_RUN marks nothing; A1 shows its view again");
  ok(/Бүх мөр Зөв.*2 мөр/.test(dry.logs), "DRY_RUN says how many lines it would mark");
  const live = pickInA1(false, "Бүгд");
  eq([live.error, marks(), tab.value(1, 1)], [null, ["Зөв", "Зөв", "Буруу", "", ""], "Бүгд"],
     "every line mom had not looked at gets Зөв, the doubtful one too; her Буруу, her changed count and the blank row stay; A1 shows the view it showed");
  eq(tab.hidden.filter((c) => c <= header.length), [], "the view is still \"Бүгд\"");
  ok(/2 мөр/.test(live.logs) && /ээжийн тэмдгээр бичсэн/.test(live.logs), "the marked tab goes into the workbooks at once, as a single Зөв does");
  eq(plain(run(sw, true, `ocrCalendarDay(ocrReadRecords(SpreadsheetApp.openById("${file.id}").getSheetByName("22")))`).value)["зарлага"].slice(2, 3), [0], "nothing waits for mom: the day can turn green");
  // the site sees it at once, not at the next hourly pass (the owner, 2026-10-29)
  const kept = JSON.parse(sw.props["ocr-cal:Соронзон|2026-10|зарлага"] || "{}");
  eq([kept["22"] && kept["22"]["зарлага"][2], kept["22"] && kept["22"]["зарлага"][5], Boolean(sw.props["ocr-cal-edit"])], [0, 0, true], "the day's numbers are kept right after the mark: no line waits for mom, none for checking");
  sw.props.OCR_FEED_TOKEN = "t";
  const feed = () => plain(run(sw, true, `ocrFeedCachedAnswer({ token: "t", mode: "receipts", month: "2026-10" }, PropertiesService.getScriptProperties(), CacheService.getScriptCache(), new Date())`).value).places["Соронзон"]["22"];
  eq(feed()["зарлага"][2], 0, "the site's answer has them");
  sw.props["ocr-cal:Соронзон|2026-10|зарлага"] = JSON.stringify({ "22": { "зарлага": [1, 4, 2, 0, 0, 1, 0] } });
  eq(feed()["зарлага"][2], 0, "a visit within the cache time gets the kept answer ...");
  sw.props["ocr-cal-edit"] = "later";
  eq(feed()["зарлага"][2], 2, "... until mom's next mark stamps a new cache key");
  pickInA1(false, undefined);
  eq(tab.value(1, 1), "Ээж", "no view known before (a paste): mom's view");
}

// ---------- a tab of the layout before 2026-10-28 and mom's instant marks (2026-10-29) ----------
// The layout before: the reason in F, "Бичсэн" in G, no photo column. Mom's Зөв wrote such a tab at
// once by the new columns: "Бичсэн" into F (the old reason) and the reason into N (no title).
{
  const OLD = ["Ээж", "Бараа", "Тоо", "Ээж: Зөв/Буруу", "Төлөв", "Шалгах шалтгаан", "Бичсэн", "Хавтас (ээж хийсэн)", "Баримт", "Бараа (тааруулсан)", "Тоо (уншсан)", "Итгэл", "ID"];
  const oldLine = (values, n) => [...OLD.map((title) => ({ "Хавтас (ээж хийсэн)": "Соронзон / зарлага", "Баримт": "Z.HEIC", "Төлөв": "OK", "ID": `z1#${n}`, ...values }[title] ?? "")), values.N ?? ""];
  const at = (sheet, r, title) => sheet.value(r, header.indexOf(title) + 1);
  { // the damage repaired when the tab moves into the new layout
    const sw = makeWorld();
    const tab = new Sheet([[...OLD],
      oldLine({ "Бараа": "Haier 300l", "Тоо": 1, "Ээж: Зөв/Буруу": "Зөв", "Шалгах шалтгаан": "Sor/2026/10/31!AB271 ← 1 · Haier 300l ✓", "Бичсэн": "Sor/2026/10/31!AB271 ← 1 · Haier 300l" }, 1),
      oldLine({ "Бараа": "26-55", "Тоо": 1, "Ээж: Зөв/Буруу": "Бэлэг", "Төлөв": "Шалгах", "Шалгах шалтгаан": "✗ \"26-55\" Sor/2026/10/31 дээр алга [k:v2|Бэлэг|26-55|1]", "Бичсэн": "✗ \"26-55\" Sor/2026/10/31 дээр алга [k:v2||26-55|1]", N: "үнэ байхгүй; бичиж чадсангүй: \"26-55\" Sor/2026/10/31 дээр алга" }, 2),
      oldLine({ "Бараа": "аяга", "Тоо": 1, "Ээж: Зөв/Буруу": "Бэлэг", "Төлөв": "Шалгах", "Шалгах шалтгаан": "✗ \"аяга\" Sor/2026/10/31 дээр алга [k:v2|Бэлэг|аяга|1]", N: "✗ \"аяга\" Sor/2026/10/31 дээр алга [k:v2|Бэлэг|аяга|1]" }, 3),
      oldLine({ "Бараа": "Haier 400l", "Тоо": 2, "Төлөв": "Шалгах", "Шалгах шалтгаан": "итгэл бага (0.6)", "Бичсэн": "Sor/2026/10/31!AB12 ← 2 · Haier 400l ⚠" }, 4),
    ]);
    const file = sw.root.path("OCR", "2026-10", "Зарлага").add(new File("OCR Соронзон зарлага 2026-10", "application/vnd.google-apps.spreadsheet", "", 1, "me"));
    sw.books[file.id] = new Book(file.id, file.name, { "22": tab });
    const moved = run(sw, false, `ocrUpgradeTab(SpreadsheetApp.openById("${file.id}").getSheetByName("22"))`);
    eq([moved.error, plain(moved.value), tab.cells[0].slice(1, header.length)], [null, true, header.slice(1)], "the tab moves into the new layout");
    eq([2, 3, 4, 5].map((r) => [at(tab, r, "Бичсэн"), at(tab, r, "Шалгах шалтгаан")]), [
      ["Sor/2026/10/31!AB271 ← 1 · Haier 300l ✓", ""],
      ["✗ \"26-55\" Sor/2026/10/31 дээр алга [k:v2|Бэлэг|26-55|1]", "үнэ байхгүй; бичиж чадсангүй: \"26-55\" Sor/2026/10/31 дээр алга"],
      ["✗ \"аяга\" Sor/2026/10/31 дээр алга [k:v2|Бэлэг|аяга|1]", ""],
      ["Sor/2026/10/31!AB12 ← 2 · Haier 400l ⚠", "итгэл бага (0.6)"],
    ], "\"Бичсэн\" written into the old reason column is \"Бичсэн\" again, the reason in N its reason (none when N got a \"Бичсэн\" too); an untouched line keeps both");
    eq([2, 3, 4, 5].map((r) => [at(tab, r, "Ээж: Зөв/Буруу"), at(tab, r, "ID")]), [["Зөв", "z1#1"], ["Бэлэг", "z1#2"], ["Бэлэг", "z1#3"], ["", "z1#4"]], "mom's marks and the IDs stay");
  }
  { // mom's mark on such a tab: it moves into the new layout first, then it is written
    const sw = makeWorld();
    const tab = new Sheet([[...OLD], oldLine({ "Бараа": "Haier 400l", "Тоо": 2, "Бараа (тааруулсан)": "Haier 400l", "Тоо (уншсан)": 2, "Ээж: Зөв/Буруу": "Зөв", "Төлөв": "Шалгах", "Шалгах шалтгаан": "итгэл бага (0.6)" }, 1)]);
    const file = sw.root.path("OCR", "2026-10", "Зарлага").add(new File("OCR Соронзон зарлага 2026-10", "application/vnd.google-apps.spreadsheet", "", 1, "me"));
    sw.books[file.id] = new Book(file.id, file.name, { "22": tab });
    const D = OLD.indexOf("Ээж: Зөв/Буруу") + 1;
    const edit = run(sw, false, `onOcrViewEdit({ range: { getRow: () => 2, getColumn: () => ${D}, getLastColumn: () => ${D}, getValue: () => "Зөв", getSheet: () => SpreadsheetApp.openById("${file.id}").getSheetByName("22") } })`);
    eq([edit.error, tab.cells[0].slice(1, header.length)], [null, header.slice(1)], "the tab is in the new layout");
    ok(at(tab, 2, "Бичсэн") !== "" && at(tab, 2, "Шалгах шалтгаан").startsWith("итгэл бага (0.6)"), "\"Бичсэн\" and the reason each in its own column");
    eq([at(tab, 2, "Хавтас (ээж хийсэн)"), at(tab, 2, "ID")], ["Соронзон / зарлага", "z1#1"], "the other columns too");
    ok(tab.protections.some((p) => p.description.startsWith("OCR хамгаалалт")), "and it is protected again (the move took the old protection away)");
  }
  { // one spreadsheet reviewed now from the editor (the owner, 2026-10-29: not at the next hourly pass)
    const sw = makeWorld();
    const tab = new Sheet([[...OLD], oldLine({ "Бараа": "Haier 300l", "Тоо": 1, "Ээж: Зөв/Буруу": "Зөв", "Шалгах шалтгаан": "Sor/2026/10/31!AB271 ← 1 · Haier 300l ✓", "Бичсэн": "Sor/2026/10/31!AB271 ← 1 · Haier 300l" }, 1)]);
    const file = sw.root.path("OCR", "2026-10", "Зарлага").add(new File("OCR Соронзон зарлага 2026-10", "application/vnd.google-apps.spreadsheet", "", 1, "me"));
    sw.books[file.id] = new Book(file.id, file.name, { "22": tab });
    sw.props["ocr-cal-edit"] = "before";
    const now = run(sw, false, `reviewOcrBookNow("${file.id}")`);
    eq([now.error, plain(now.value).isFinished, tab.cells[0].slice(1, header.length), /^Sor\/2026\/10\/31!\w+ ← 1 · Haier 300l ✓$/.test(at(tab, 2, "Бичсэн")), at(tab, 2, "Шалгах шалтгаан")], [null, true, header.slice(1), true, ""],
       "every tab moved and put right, the spreadsheet reviewed (the line in this fake workbook's Haier 300l row)");
    eq([JSON.parse(sw.props["ocr-cal:Соронзон|2026-10|зарлага"])["22"]["зарлага"][2], sw.props["ocr-cal-edit"] !== "before"], [0, true], "its days' numbers kept, with a new stamp in the site's cache key: the site shows them at once");
    ok(/reviewOcrBookNow\(OCR_SORONZON_SALES_2026_10\)/.test(code), "Соронзон's sales of October have their own function to run from the editor");
  }
}


// ---------- the reasons in a few words, as wide as mom's screen beside the photo (the owner, 2026-10-29) ----------
{
  const short = (text) => vm.runInContext(`ocrShortenReasons(${JSON.stringify(text)})`, pure);
  eq(short("хавтас буруу байж магадгүй: зарлагын баримт, илгээгч тодорхойгүй (Тосгоноос ирсэн бол Тосгоноос My home руу хавтсанд); үнэ байхгүй"),
     "хавтас буруу байж магадгүй: илгээгч тодорхойгүй → Тосгоноос My home?; үнэ байхгүй", "the old wrong-folder note in a few words, the rest as it was");
  eq(short("үнэ байхгүй; бичиж чадсангүй: \"Panasonic 9.0l\" My/2026/10/31 дээр алга"), "үнэ байхгүй; бичиж чадсангүй: бараа ажлын номд алга",
     "a product with no row: the name is in Бараа already, the sheet in Бичсэн");
  eq(short("бичиж чадсангүй: check: \"Panasonic 203\" Sor/2026/10/31 дээр алга"), "бичиж чадсангүй: check файлд мөр алга", "one the check copy lacks");
  eq(short("бичиж чадсангүй: \"a\" My/2026/10/31 дээр алга; бичиж чадсангүй: \"b\" My/2026/10/15 дээр алга"), "бичиж чадсангүй: бараа ажлын номд алга", "two that say the same now: once");
  const other = "итгэл бага (0.6); бичиж чадсангүй: Sor/2026/10/31!V262: нүдэнд гараар бичсэн 1 байна";
  eq([short(other), short(""), short(short("үнэ байхгүй; бичиж чадсангүй: \"x\" My/2026/10/31 дээр алга"))], [other, "", "үнэ байхгүй; бичиж чадсангүй: бараа ажлын номд алга"], "any other reason as it is; twice is once");
  { // a tab's old long reasons put right by its review, and a failed line says so in a few words, once
    const sw = makeWorld();
    const row = (values, n) => header.map((title) => ({ "Хавтас (ээж хийсэн)": "Соронзон / зарлага", "Баримт": "Z.HEIC", "Төлөв": "Шалгах", "ID": `z1#${n}`, ...values }[title] ?? ""));
    const tab = new Sheet([[...header],
      row({ "Бараа": "26-55", "Тоо": 1, "Ээж: Зөв/Буруу": "Зөв", "Бичсэн": "✗ \"26-55\" Sor/2026/10/31 дээр алга [k:v3|Зөв|26-55|1]", "Шалгах шалтгаан": "үнэ байхгүй; бичиж чадсангүй: \"26-55\" Sor/2026/10/31 дээр алга" }, 1),
      // not approved (mom neither marked it nor chose the product): no new product (ocr-new-product.gs)
      row({ "Бараа": "аяга", "Тоо": 2, "Бараа (тааруулсан)": "аяга", "Тоо (уншсан)": 2, "Шалгах шалтгаан": "үнэ байхгүй" }, 2),
    ]);
    const file = sw.root.path("OCR", "2026-10", "Зарлага").add(new File("OCR Соронзон зарлага 2026-10", "application/vnd.google-apps.spreadsheet", "", 1, "me"));
    sw.books[file.id] = new Book(file.id, file.name, { "22": tab });
    const now = run(sw, false, `reviewOcrBookNow("${file.id}")`);
    const at = (r, title) => tab.value(r, header.indexOf(title) + 1);
    eq([now.error, at(2, "Шалгах шалтгаан"), at(3, "Шалгах шалтгаан")], [null, "үнэ байхгүй; бичиж чадсангүй: бараа ажлын номд алга", "үнэ байхгүй; бичиж чадсангүй: бараа ажлын номд алга"],
       "the old reason shortened, the new failure in the same few words, never twice");
  }
}

// ---------- a product the workbook has no row for, put in once mom approved its line (the owner, 2026-10-29) ----------
{
  const pick = (expr) => plain(vm.runInContext(expr, pure));
  eq(pick(`[ocrLooseKey("Meiling 254"), ocrLooseKey("Meiling 254l"), ocrLooseKey("Panasonic 9.0l"), ocrLooseKey("xinix 240l") === ocrLooseKey("XINGX 240l")]`),
     ["meiling254", "meiling254", "panasonic90", false], "the same product spelt a little differently: no case, spaces, dots or litre mark");
  eq(pick(`[ocrListLabel("Соронзон", "My home"), ocrListLabel("Хоёулаа", "My home"), ocrListLabel("My home", "My home"), ocrListLabel("", "Тосгон")]`),
     ["Хоёулаа", "Хоёулаа", "My home", "Тосгон"], "\"Аль дэлгүүрт\": a second place makes it Хоёулаа");
  const cats = (labels) => labels.map((label, i) => ({ key: label.trim().replace(/\s+/g, " ").toLowerCase(), label, row: i + 4 }));
  eq(pick(`[ocrSmallGoodsKey(${JSON.stringify(cats(["Хөргөгч", "Бусад (Жижиг бараа)", "Индүү"]))}), ocrSmallGoodsKey(${JSON.stringify(cats(["Зурагт", "Хятад жижиг бараа"]))}), ocrSmallGoodsKey(${JSON.stringify(cats(["Зурагт"]))})]`),
     ["бусад (жижиг бараа)", "хятад жижиг бараа", null], "no category fits: the small goods'");
  eq(pick(`[ocrTabsFrom("My home", "2026-10-22"), ocrTabsFrom("Соронзон", "2026-12-03"), ocrTabsFrom("Тосгон", "2026-11-05")]`),
     [["My/2026/10/31", "My/2026/11/15", "My/2026/11/30", "My/2026/12/15", "My/2026/12/31"], ["Sor/2026/12/15", "Sor/2026/12/31"], ["Tos/2026/11", "Tos/2026/12"]],
     "the line's tab and every later one of the year");

  const sw = makeWorld();
  sw.props.GEMINI_KEY = "test-key";
  const book = sw.books[CATALOG_ID];
  book.add("Бараа", new Sheet([[], ["Соронзон орон"], ["тайлбар"], ["Барааны нэр", "Аль дэлгүүрт", "өртөг үнэ", "зарах үнэ", "яг харагдах үнэ"],
    ["Хятад жижиг бараа"], ["Залгуур 40", "Соронзон", 28000, 40000, 40000], ["Залгуур 45", "Хоёулаа", 27000, 45000, 45000],
    ["Хөргөгч"], ["Haier 90-197", "Хоёулаа", 300000, 450000, 460000], ["Meiling 254l", "Хоёулаа", 625000, 860000, 890000], ["Akira 666 индүү", "Хоёулаа", 37000, 50000, 55000]]));
  // a category: a coloured name in A with nothing else on its row (dividers.gs)
  const shopTab = () => {
    const tab = new Sheet([...HEAD, ["Хөргөгч"], ["Haier 90-197", 1, 300000, 450000], ["Бусад (Жижиг бараа)"], ["Залгуур 45", 2, 27000, 45000], ["Индүү"], ["Akira 666 индүү", 4, 37000, 50000]]);
    [4, 6, 8].forEach((r) => tab.getRange(r, 1).setBackground("#fce5cd"));
    tab.getRange(7, 1).setNote("Залгуур 45-ын тэмдэглэл");
    return tab;
  };
  book.order = book.order.filter((tab) => tab.name !== "My/2026/10/31");
  const oct = book.add("My/2026/10/31", shopTab());
  const nov = book.add("My/2026/11/15", shopTab());
  const line = (values, n) => header.map((title) => ({ "Хавтас (ээж хийсэн)": "My home / орлого", "Баримт": "IMG_5879.HEIC", "Төлөв": "Шалгах", "ID": `n1#${n}`, ...values }[title] ?? ""));
  const tab = new Sheet([[...header], line({ "Бараа": "Akira 666 индүү", "Тоо": 3, "Бараа (тааруулсан)": "Akira 666 индүү", "Тоо (уншсан)": 3, "Ээж: Зөв/Буруу": "Зөв" }, 1)]);
  const file = sw.root.path("OCR", "2026-10", "Орлого").add(new File("OCR My home орлого 2026-10", "application/vnd.google-apps.spreadsheet", "", 1, "me"));
  sw.books[file.id] = new Book(file.id, file.name, { "22": tab });
  // a line by its ID: the review spreads the photos, so the rows move
  const at = (n, title) => tab.value(tab.cells.findIndex((row) => row[header.indexOf("ID")] === `n1#${n}`) + 1, header.indexOf(title) + 1);
  const review = (dry = false) => run(sw, dry, `reviewOcrBookNow("${file.id}")`);
  eq(review().error, null, "a line of a product the tab has");
  const akira = at(1, "Бичсэн").match(/^My\/2026\/10\/31!([A-Z]+)9 ← 3 · Akira 666 индүү ✓$/);
  ok(akira, `written into its row: ${at(1, "Бичсэн")}`);
  const col = akira[1].split("").reduce((t, l) => t * 26 + l.charCodeAt(0) - 64, 0);

  // mom approves a product the tab has no row for, one the list spells otherwise, one she did not mark
  tab.cells.push(line({ "Бараа": "Залгуур 40", "Тоо": 5, "Бараа (тааруулсан)": "Залгуур 40", "Тоо (уншсан)": 5, "Ээж: Зөв/Буруу": "Зөв" }, 2));
  tab.cells.push(line({ "Бараа": "Meiling 254", "Тоо": 1, "Бараа (тааруулсан)": "Meiling 254", "Тоо (уншсан)": 1, "Ээж: Зөв/Буруу": "Зөв" }, 3));
  tab.cells.push(line({ "Бараа": "Шинэ тогоо", "Тоо": 2, "Бараа (тааруулсан)": "Шинэ тогоо", "Тоо (уншсан)": 2 }, 4));
  sw.categoryAnswer = { sheet_category: "Бусад (Жижиг бараа)", list_category: "Хятад жижиг бараа" };
  const dry = review(true);
  eq([dry.error, (sw.categoryCalls || []).length, oct.getLastRow()], [null, 0, 9], "DRY_RUN: no question, no row");
  ok(dry.logs.split("\n").some((l) => l.includes("Залгуур 40") && l.includes("шинэ бараа")), "only logged");
  // past the run's deadline: putting a product's rows into every later tab takes seconds a tab (a pass
  // and a queued review ran past Google's 6 minutes, 2026-10-29), so the line waits for the next run
  const late = run(sw, false, `(() => { const book = SpreadsheetApp.openById("${file.id}"); return reviewOcrTab(book.getSheetByName("22"), { file: DriveApp.getFileById("${file.id}"), ...ocrMonthBookOf(book.getName()) }, { workbooks: new Map(), context: ocrEditorContext(PropertiesService.getScriptProperties()), timeZone: "UTC", deadline: 0 }); })()`);
  eq([late.error, plain(late.value).deferred, (sw.categoryCalls || []).length, oct.getLastRow(), at(2, "Бичсэн"), at(3, "Бичсэн")], [null, 2, 0, 9, "", ""],
     "out of time: the lines that need a new product's rows wait, nothing asked or put in");
  eq(review().error, null, "the review");
  eq([JSON.parse(sw.props["ocr-relink-queue"]).months, sw.triggers.filter((t) => t.getHandlerFunction() === "reviewMovedOcrBooksNow").length], [["2026-10", "2026-11"], 1],
     "rows put in: their months' spreadsheets are queued to name their cells again, a minute later");
  eq(sw.categoryCalls.length, 1, "Gemini asked once, for the one approved new product");
  ok(sw.categoryCalls[0].prompt.includes("Залгуур 40") && ["Хөргөгч", "Бусад (Жижиг бараа)", "Индүү"].every((c) => sw.categoryCalls[0].prompt.includes(c)), "with the sheet's own categories");
  eq(oct.cells.slice(3).map((row) => row[0]), ["Хөргөгч", "Haier 90-197", "Бусад (Жижиг бараа)", "Залгуур 45", "Залгуур 40", "Индүү", "Akira 666 индүү"], "under the category Gemini chose, at its end");
  eq([1, 2, 3, 4].map((c) => oct.value(8, c)), ["Залгуур 40", 0, 28000, 40000], "opening balance 0, the list's cost and price");
  eq([oct.value(9, 1), oct.getRange(8, 1).getNote(), oct.getRange(9, 1).getBackground()], ["Индүү", "", "#fce5cd"], "the rows below one down; the template's note not copied");
  eq(nov.cells.slice(3).map((row) => row[0]), ["Хөргөгч", "Haier 90-197", "Бусад (Жижиг бараа)", "Залгуур 45", "Залгуур 40", "Индүү", "Akira 666 индүү"], "the later tab of the year too");
  eq([nov.value(8, 2), nov.value(8, 3), nov.value(8, 4)], ["", 28000, 40000], "its opening balance left to the carry");
  const list = book.sheets["Бараа"];
  eq([list.value(6, 1), list.value(6, 2), list.getLastRow()], ["Залгуур 40", "Хоёулаа", 11], "the list: My home has it too now");
  ok(new RegExp(`^My/2026/10/31![A-Z]+8 ← 5 · Залгуур 40 ✓$`).test(at(2, "Бичсэн")), `the line written into the new row: ${at(2, "Бичсэн")}`);
  ok(at(3, "Шалгах шалтгаан").includes("жагсаалтад \"Meiling 254l\" гэж бий") && !oct.cells.some((row) => row[0] === "Meiling 254"), "a name the list spells otherwise: nothing put in, mom picks it");
  ok(at(4, "Шалгах шалтгаан").includes("бараа ажлын номд алга") && !oct.cells.some((row) => row[0] === "Шинэ тогоо"), "a line mom did not approve: nothing put in");
  eq([review().error, sw.categoryCalls.length, oct.getLastRow()], [null, 1, 10], "the next review: nothing asked or put in again");
  eq([at(1, "Бичсэн"), oct.value(10, col), oct.value(9, col)], [`My/2026/10/31!${akira[1]}10 ← 3 · Akira 666 индүү ✓`, 3, ""], "a line whose row moved down is found again by its product, not counted twice");

  // Gemini finds no category, or does not answer: the small goods'
  tab.cells.push(line({ "Бараа": "Гэрэл 5", "Тоо": 1, "Бараа (тааруулсан)": "Гэрэл 5", "Тоо (уншсан)": 1, "Ээж: Зөв/Буруу": "Зөв" }, 5));
  sw.categoryAnswer = { code: 500, text: "error" };
  eq(review().error, null, "Gemini down");
  eq(oct.cells.slice(3).map((row) => row[0]), ["Хөргөгч", "Haier 90-197", "Бусад (Жижиг бараа)", "Залгуур 45", "Залгуур 40", "Гэрэл 5", "Индүү", "Akira 666 индүү"], "into the small goods");
  eq([list.value(8, 1), list.value(8, 2), list.getLastRow()], ["Гэрэл 5", "My home"].concat(12), "a product new to the list: under the list's small goods heading, at its end");

  // mom drags a new product's row under another category: the later tabs follow (the owner, 2026-10-29)
  eq(pick(`[ocrCategoryPlan(["a", "b", "a", "a"], ["a", "a", "a", "a"]), ocrCategoryPlan(["a", "b", "b", "b"], ["a", "b", "b", "b"]), ocrCategoryPlan(["a", "b", null, "c"], ["a", "a", "a", "a"])]`), [
    { moves: [{ index: 2, to: "b" }, { index: 3, to: "b" }], recorded: ["a", "b", "b", "b"] },
    { moves: [], recorded: ["a", "b", "b", "b"] },
    { moves: [], recorded: ["a", "b", "a", "c"] },
  ], "a tab mom moved it in leads the tabs after it; the tabs before stay; a tab she moved too is hers");
  const rowOf = (sheet, name) => sheet.cells.findIndex((row) => row[0] === name) + 1;
  oct.moveRows(oct.getRange(rowOf(oct, "Залгуур 40"), 1, 1, 1), rowOf(oct, "Akira 666 индүү") + 1);
  eq(oct.cells.slice(3).map((row) => row[0]), ["Хөргөгч", "Haier 90-197", "Бусад (Жижиг бараа)", "Залгуур 45", "Гэрэл 5", "Индүү", "Akira 666 индүү", "Залгуур 40"], "mom moved it under Индүү in October");
  const dryFollow = run(sw, true, "followCategoryMovesNow()");
  eq([dryFollow.error, nov.cells.slice(3).map((row) => row[0]).indexOf("Залгуур 40")], [null, 4], "DRY_RUN: November as it was");
  const follow = run(sw, false, "followCategoryMovesNow()");
  eq([follow.error, nov.cells.slice(3).map((row) => row[0])], [null, ["Хөргөгч", "Haier 90-197", "Бусад (Жижиг бараа)", "Залгуур 45", "Гэрэл 5", "Индүү", "Akira 666 индүү", "Залгуур 40"]], "November follows, its numbers with it");
  eq([nov.value(rowOf(nov, "Залгуур 40"), 3), nov.value(rowOf(nov, "Залгуур 40"), 4)], [28000, 40000], "the row moved whole");
  const novBefore = JSON.stringify(nov.cells);
  eq([run(sw, false, "followCategoryMovesNow()").error, JSON.stringify(nov.cells) === novBefore, rowOf(oct, "Залгуур 40")], [null, true, 11], "the next time: nothing moves back");

  // a product the list has under a heading the sheet has too goes there, Gemini not asked (2026-10-29:
  // My home's Panasonic 9.0l, a washing machine in the list, was put among the small goods)
  const asked = sw.categoryCalls.length;
  tab.cells.push(line({ "Бараа": "Meiling 254l", "Тоо": 1, "Бараа (тааруулсан)": "Meiling 254", "Тоо (уншсан)": 1, "Ээж: Зөв/Буруу": "Зөв" }, 6));
  eq([review().error, sw.categoryCalls.length, oct.cells.slice(3, 7).map((row) => row[0])], [null, asked, ["Хөргөгч", "Haier 90-197", "Meiling 254l", "Бусад (Жижиг бараа)"]],
     "under the list's own heading, the sheet's category of that name");
  ok(pick(`ocrCategoryPrompt("Залгуур 40", ["Бусад (Жижиг бараа)"], [], "Хятад жижиг бараа")`).includes("\"Хятад жижиг бараа\" ангилалд байдаг"), "a heading the sheet lacks is told to Gemini");

  // the price matters: a dear product is no small goods (the owner, 2026-10-29)
  eq(pick(`[ocrPriceFitKey([{ key: "хөргөгч", price: 450000 }, { key: "бусад (жижиг бараа)", price: 45000 }, { key: "индүү", price: 50000 }], 1200000), ocrPriceFitKey([{ key: "хөргөгч", price: 450000 }, { key: "индүү", price: 50000 }], 60000), ocrPriceFitKey([{ key: "бусад (жижиг бараа)", price: 45000 }], 900000)]`),
     ["хөргөгч", "индүү", null], "the category whose usual price is nearest, never the small goods'");
  ok(pick(`ocrCategoryPrompt("Шинэ 500", ["Хөргөгч", "Бусад (Жижиг бараа)"], [], null, 1200000, { "Хөргөгч": 450000, "Бусад (Жижиг бараа)": 45000 })`).match(/1,200,000₮[\s\S]*"Хөргөгч" ~450,000₮[\s\S]*жижиг бараа биш/), "the prompt gives the price, each category's usual one and the rule");
  const PRICE_COL = pick("OCR_READ_PRICE_COLUMN");
  tab.cells.push(line({ "Бараа": "Шинэ хөргөгч 500", "Тоо": 1, "Бараа (тааруулсан)": "Шинэ хөргөгч", "Тоо (уншсан)": 1, "Ээж: Зөв/Буруу": "Зөв" }, 7));
  tab.cells[tab.cells.length - 1][PRICE_COL - 1] = 1200000; // the receipt price read
  sw.categoryAnswer = { sheet_category: "Бусад (Жижиг бараа)", list_category: "Хятад жижиг бараа" };
  eq(review().error, null, "a dear new product Gemini put among the small goods");
  ok(sw.categoryCalls[sw.categoryCalls.length - 1].prompt.includes("1,200,000₮"), "Gemini is told its receipt price");
  eq(oct.cells.slice(3, 8).map((row) => row[0]), ["Хөргөгч", "Haier 90-197", "Meiling 254l", "Шинэ хөргөгч 500", "Бусад (Жижиг бараа)"], "it goes where its price belongs");
}

// ---------- every month spreadsheet set up, run after run by itself until all are (the owner, 2026-10-29) ----------
{
  const sw = makeWorld();
  sw.props.GEMINI_KEY = "k";
  put(sw, ["Тосгон", "тооллого", "2026-10", "23"], "G.jpeg", "photo-g", 1, receipt(null, [item("Haier 300", "Haier 300l", 1, null)]));
  eq(run(sw, false, "processNewReceipts()").error, null, "a month spreadsheet made by a run");
  Object.keys(sw.props).filter((key) => key.startsWith("ocr-book-setup:")).forEach((key) => { sw.props[key] = "old"; });
  const goesOn = () => sw.triggers.filter((t) => t.getHandlerFunction() === "setUpAllOcrBooksNow").length;
  const cut = run(sw, false, "(() => { const real = Date.now; let n = 0; Date.now = () => (n++ < 2 ? real() : real() + 3600000); try { return setUpAllOcrBooksNow(); } finally { Date.now = real; } })()");
  eq([cut.error, plain(cut.value).left > 0, goesOn()], [null, true, 1], "out of time: it runs again by itself a minute later");
  sw.triggers = sw.triggers.filter((t) => t.getHandlerFunction() !== "setUpAllOcrBooksNow"); // the run a minute later came
  sw.locked = true; // the hourly pass holds the lock
  eq([run(sw, false, "setUpAllOcrBooksNow()").error, goesOn()], [null, 1], "the lock taken: it tries again a minute later");
  sw.locked = false;
  const done = run(sw, false, "setUpAllOcrBooksNow()");
  eq([done.error, plain(done.value).left, goesOn()], [null, 0, 0], "all set up: it stops");
}

console.log(`ALL ${n} OCR CHECKS PASSED${isBundle ? " (bundle)" : ""}`);
