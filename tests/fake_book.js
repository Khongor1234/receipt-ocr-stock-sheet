// An in-memory spreadsheet (just enough of the Apps Script API) and a runner that evaluates the
// project's .gs files against it. Shared by main_test.js and tosgon_test.js.
const vm = require("vm"), assert = require("assert");
const WHITE = "#ffffff";
const cell = () => ({ v: "", bg: WHITE, f: "", nf: "" });
class Protection {
  constructor(range) { this.range = range; this.desc = ""; this.warning = true; this.editors = ["mom"]; }
  setDescription(d) { this.desc = d; return this; } getDescription() { return this.desc; }
  setWarningOnly(w) { this.warning = w; return this; } isWarningOnly() { return this.warning; }
  getEditors() { return this.editors; }
  removeEditors(e) { this.editors = this.editors.filter(x => !e.includes(x)); return this; }
  canDomainEdit() { return false; } setDomainEdit() {} getRange() { return this.range; } canEdit() { return true; } setUnprotectedRanges(r) { this.unprotected = r; return this; } getUnprotectedRanges() { return this.unprotected || []; }
  remove() { this.range.sheet.protections = this.range.sheet.protections.filter(p => p !== this); }
}
class Range {
  constructor(sheet, r, c, h, w) { Object.assign(this, { sheet, r, c, h, w }); }
  each(fn) { for (let i = 0; i < this.h; i++) for (let j = 0; j < this.w; j++) fn(this.sheet.cell(this.r + i, this.c + j), i, j); }
  map(fn) { const out = []; for (let i = 0; i < this.h; i++) { out.push([]); for (let j = 0; j < this.w; j++) out[i].push(fn(this.sheet.cell(this.r + i, this.c + j))); } return out; }
  getValues() { this.sheet.reads = (this.sheet.reads || 0) + 1; return this.map(x => x.v); } getValue() { return this.getValues()[0][0]; }
  getBackgrounds() { return this.map(x => x.bg); } getBackground() { return this.sheet.cell(this.r, this.c).bg; } getFormulas() { return this.map(x => x.f); }
  createFilter() { const sheet = this.sheet; const filter = { range: this, criteria: {}, getRange() { return filter.range; }, setColumnFilterCriteria(c, v) { filter.criteria[c] = v; return filter; }, getColumnFilterCriteria(c) { return filter.criteria[c] || null; }, remove() { sheet.filter = null; } }; assert(!sheet.filter, "a sheet holds one filter"); sheet.filter = filter; return filter; }
  getRow() { return this.r; } getNumRows() { return this.h; } getNumColumns() { return this.w; }
  getSheet() { return this.sheet; } getColumn() { return this.c; } getLastColumn() { return this.c + this.w - 1; } getLastRow() { return this.r + this.h - 1; }
  protect() { const p = new Protection(this); this.sheet.protections.push(p); return p; }
  breakApart() { return this; } merge() { this.sheet.merges = (this.sheet.merges || []).concat([[this.r, this.c, this.h, this.w]]); return this; } setFontColor() { return this; }
  setValue(v) { this.each(x => { x.v = v; x.f = ""; }); return this; }
  // Like Apps Script: a text that starts with "=" is stored as a formula.
  setValues(a) { this.each((x, i, j) => { const isFormula = typeof a[i][j] === "string" && a[i][j].startsWith("="); x.f = isFormula ? a[i][j] : ""; x.v = isFormula ? "(тооцоолсон)" : a[i][j]; }); return this; }
  setFormula(f) { this.each(x => { x.f = f; x.v = "(тооцоолсон)"; }); return this; } getFormula() { return this.sheet.cell(this.r, this.c).f; }
  setFormulas(a) { this.each((x, i, j) => { x.f = a[i][j]; x.v = "(тооцоолсон)"; }); return this; }
  getDisplayValues() { return this.map(x => String(x.v)); }
  getDataValidations() { return this.map(() => null).map((row, i) => row.map((_, j) => this.sheet.validations.some(v => this.r + i >= v.r && this.r + i < v.r + v.h && this.c + j >= v.c && this.c + j < v.c + v.w) ? "rule" : null)); }
  getDataValidation() { const v = this.sheet.validations.filter(x => this.r >= x.r && this.r < x.r + x.h && this.c >= x.c && this.c < x.c + x.w).pop(); return v ? v.rule : null; } // the newest rule of a cell wins
  setDataValidations(a) { this.each((x, i, j) => { if (a[i][j]) this.sheet.validations.push({ r: this.r + i, c: this.c + j, h: 1, w: 1, rule: a[i][j] }); }); return this; }
  clearDataValidations() { const inside = (v) => v.r >= this.r && v.r + v.h <= this.r + this.h && v.c >= this.c && v.c + v.w <= this.c + this.w; this.sheet.validations = this.sheet.validations.filter(v => !inside(v)); return this; }
  getNumberFormats() { return this.map(x => x.nf || ""); } setNumberFormat(f) { this.each(x => { x.nf = f; }); return this; }
  setNumberFormats(a) { this.each((x, i, j) => { x.nf = a[i][j]; }); return this; }
  getFontColors() { return this.map(x => x.color || "#000000"); } setFontColors(a) { this.each((x, i, j) => { x.color = a[i][j]; }); return this; }
  getNotes() { return this.map(x => x.note || ""); } setNotes(a) { this.each((x, i, j) => { x.note = a[i][j]; }); return this; } clearNote() { this.each(x => { x.note = ""; }); return this; }
  setBackground(c) { this.each(x => { x.bg = c; }); return this; }
  // like Apps Script (seen 2026-10-22/24): no colour goes into a row a filter hides
  setBackgrounds(a) { assert(a.length === this.h && a[0].length === this.w, "setBackgrounds size"); this.each((x, i, j) => { if (!this.sheet.isFilteredOut(this.r + i)) x.bg = a[i][j]; }); return this; }
  setBorder(...sides) { const on = sides.some(v => v === true); this.each(x => { x.border = on; }); return this; }
  setFontWeight(w) { this.each(x => { x.bold = w === "bold"; }); return this; } setFontSize() { return this; } setHorizontalAlignment() { return this; } setVerticalAlignment() { return this; }
  setWrap() { return this; } setDataValidation(rule) { this.sheet.validations = this.sheet.validations.filter(v => !(v.r >= this.r && v.r + v.h <= this.r + this.h && v.c >= this.c && v.c + v.w <= this.c + this.w)); this.sheet.validations.push({ r: this.r, c: this.c, h: this.h, w: this.w, rule }); return this; }
  clearContent() { this.each(x => { x.v = ""; x.f = ""; }); return this; }
  // Like Google Sheets: a copy (also PASTE_FORMAT) takes the cells' notes and font colours too.
  copyTo(dest) { this.each((x, i, j) => { const d = dest.sheet.cell(dest.r + i, dest.c + j); d.v = x.v; d.bg = x.bg; d.f = x.f; d.note = x.note; d.color = x.color; }); }
}
class Sheet {
  constructor(ss, name) { this.ss = ss; this.name = name; this.rows = []; this.hidden = false; this.protections = []; this.validations = []; this.maxRows = 1000; this.maxCols = 42; }
  cell(r, c) { while (this.rows.length < r) this.rows.push({}); return (this.rows[r - 1][c] = this.rows[r - 1][c] || cell()); }
  isFilteredOut(r) {
    const f = this.filter;
    if (!f || r <= f.range.r || r >= f.range.r + f.range.h) return false;
    const wanted = (crit) => String(crit && crit.getCriteriaValues ? crit.getCriteriaValues()[0] : crit);
    return Object.entries(f.criteria).some(([c, crit]) => String(((this.rows[r - 1] || {})[Number(c)] || {}).v ?? "") !== wanted(crit));
  }
  getName() { return this.name; } getParent() { return this.ss; } setName(n) { this.name = n; return this; } getFilter() { return this.filter || null; }
  getMaxColumns() { return this.maxCols; } getMaxRows() { return this.maxRows; }
  insertColumnsAfter(col, n) { this.maxCols += n; this.rows.forEach(row => { Object.keys(row).map(Number).filter(c => c > col).sort((x, y) => y - x).forEach(c => { row[c + n] = row[c]; delete row[c]; }); const left = row[col]; if (left) for (let i = 1; i <= n; i++) row[col + i] = { v: "", bg: left.bg, f: "", nf: left.nf || "" }; }); }
  deleteColumns(first, n) { this.maxCols -= n; this.rows.forEach(row => { Object.keys(row).map(Number).filter(c => c >= first).sort((x, y) => x - y).forEach(c => { const v = row[c]; delete row[c]; if (c >= first + n) row[c - n] = v; }); }); } getColumnWidth() { return 100; } setColumnWidths() {}
  getLastColumn() { let last = 0; this.rows.forEach(row => Object.entries(row).forEach(([c, x]) => { if ((x.v !== "" || x.f !== "") && Number(c) > last) last = Number(c); })); return last; }
  getSheetId() { return this.id || (this.id = ++Sheet.nextId); }
  getCharts() { return (this.charts || []).slice(); } removeChart(chart) { this.charts = (this.charts || []).filter(c => c !== chart); } insertChart(chart) { this.charts = [...(this.charts || []), chart]; }
  newChart() { const spec = { options: {}, ranges: [] }; const b = { setChartType: (t) => { spec.type = t; return b; }, addRange: (r) => { spec.ranges.push(r); return b; }, setNumHeaders: (n) => { spec.headers = n; return b; }, setOption: (k, v) => { spec.options[k] = v; return b; }, setPosition: (row, col) => { spec.position = [row, col]; return b; }, build: () => ({ spec, getOptions: () => ({ get: (k) => spec.options[k] }) }) }; return b; }
  getProtections(type) { return this.protections.filter(p => (type === "SHEET") === Boolean(p.isSheet)); }
  protect() { const p = new Protection(this.getRange(1, 1, this.maxRows, this.maxCols)); p.isSheet = true; this.protections.push(p); return p; }
  clearConditionalFormatRules() { this.cfRules = []; } setConditionalFormatRules(rules) { this.cfRules = rules; } getConditionalFormatRules() { return (this.cfRules || []).slice(); }
  setFrozenRows(n) { this.frozenRows = n; } setRowHeight() { return this; } setColumnWidth() { return this; } insertRowsAfter(r, n) { this.maxRows += n; } getIndex() { return this.ss.sheets.indexOf(this) + 1; }
  getLastRow() { let last = 0; this.rows.forEach((row, i) => { if (Object.values(row).some(x => x.v !== "" || x.f !== "")) last = i + 1; }); return last; }
  getRange(r, c, h = 1, w = 1) { return new Range(this, r, c, h, w); }
  hideColumns(first, count) { this.hiddenCols = this.hiddenCols || new Set(); for (let c = first; c < first + count; c++) this.hiddenCols.add(c); }
  showColumns(first, count) { this.hiddenCols = this.hiddenCols || new Set(); for (let c = first; c < first + count; c++) this.hiddenCols.delete(c); }
  getRangeList(list) { const ranges = list.map(a1 => { const m = a1.match(/^R(\d+)C(\d+):R(\d+)C(\d+)$/); assert(m, "getRangeList notation " + a1); return this.getRange(+m[1], +m[2], +m[3] - m[1] + 1, +m[4] - m[2] + 1); }); return { setBorder: (...sides) => { this.borderCalls = (this.borderCalls || 0) + 1; ranges.forEach(r => r.each(x => { if (sides[4] === false) x.border = false; })); } }; }
  insertRowAfter(r) { assert(!this.filter, "This operation is not supported on a range with a filtered out row."); const src = this.rows[r - 1] || {}; const copy = {}; Object.entries(src).forEach(([c, x]) => { copy[c] = { v: "", bg: x.bg, f: "" }; }); this.rows.splice(r, 0, copy); this.maxRows++; this.shiftProtections(r + 1, 1); }
  // Like Google Sheets: a protected range moves with its row; it goes when its row is deleted.
  shiftProtections(fromRow, by) { this.protections.forEach(p => { if (p.range.h === 1 && p.range.r >= fromRow) p.range = new Range(this, p.range.r + by, p.range.c, 1, p.range.w); }); }
  dropProtections(r, n) { this.protections = this.protections.filter(p => !(p.range.h === 1 && p.range.r >= r && p.range.r < r + n)); }
  deleteRow(r) { this.deleteRows(r, 1); } deleteRows(r, n) { assert(!this.filter, "This operation is not supported on a range with a filtered out row."); this.rows.splice(r - 1, n); this.maxRows -= n; this.dropProtections(r, n); this.shiftProtections(r + n, -n); }
  insertRowBefore(r) { const src = this.rows[r - 1] || {}; const copy = {}; Object.entries(src).forEach(([c, x]) => { copy[c] = { v: "", bg: x.bg, f: "" }; }); this.rows.splice(r - 1, 0, copy); this.maxRows++; this.shiftProtections(r, 1); }
  hideSheet() { this.hidden = true; } clear() { this.rows = []; this.clears = (this.clears || 0) + 1; }
  copyTo(ss) { const s = new Sheet(ss, "Copy of " + this.name); s.rows = JSON.parse(JSON.stringify(this.rows)); s.maxRows = this.maxRows; s.maxCols = this.maxCols; ss.sheets.push(s); return s; }
  logical() { return JSON.stringify([this.name, this.maxRows, this.protections.length, this.rows.map(row => Object.entries(row).filter(([, x]) => x.v !== "" || x.f !== "" || x.bg !== WHITE).map(([c, x]) => [c, String(x.v), x.f, x.bg])).filter(row => row.length > 0)]); }
}
class Spreadsheet {
  constructor() { this.sheets = []; this.active = null; }
  getSheets() { return this.sheets.slice(); } getSheetByName(n) { return this.sheets.find(s => s.name === n) || null; }
  insertSheet(n) { const s = new Sheet(this, n); this.sheets.push(s); return s; }
  deleteSheet(s) { this.sheets = this.sheets.filter(x => x !== s); }
  setActiveSheet(s) { this.active = s; } getActiveSheet() { return this.active; } moveActiveSheet(pos) { this.sheets = this.sheets.filter(x => x !== this.active); this.sheets.splice(pos - 1, 0, this.active); }
  getOwner() { return { getEmail: () => "owner@example.com" }; }
  getSpreadsheetTimeZone() { return "Asia/Ulaanbaatar"; } getName() { return this.name || "test"; } getUrl() { return "url"; }
  getId() { return this.id || (this.id = "book-" + (++Spreadsheet.nextId)); } rename(n) { this.name = n; }
  logical() { return this.sheets.map(s => s.logical()).join("\n"); }
}
Sheet.nextId = 0;
Spreadsheet.nextId = 0;

// Returns makeRunner(ss, state) for the given code: run(dry, expr, patch) evaluates expr in a
// fresh context where DRY_RUN is `dry` and returns { value, logs }.
function makeRunnerFor(code) {
  const makeRunner = (ss, state) => (dry, expr, patch = (c) => c) => {
    const logs = [];
    const trigger = (handler, source = null) => ({ getHandlerFunction: () => handler, getTriggerSourceId: () => source });
    const builder = (handler) => { let source = null; const b = { timeBased: () => b, after: (ms) => { state.delays = [...(state.delays || []), ms]; return b; }, everyMinutes: () => { throw new Error('quota: use everyHours'); }, everyHours: (h) => { state.hours = h; return b; }, forSpreadsheet: (s) => { source = typeof s === "string" ? s : s.getId(); return b; }, onChange: () => b, onEdit: () => b, create: () => { state.triggers.push(trigger(handler, source)); } }; return b; };
    const ctx = { ss, extra: state.extra, JSON, Date, editEvent: state.editEvent, Charts: { ChartType: { COLUMN: "COLUMN", LINE: "LINE" } }, Logger: { log: (...a) => logs.push(a.join(" ")) },
      DriveApp: { getFileById: (id) => { const share = (state.shares = state.shares || {}); share[id] = share[id] || { editors: [], viewers: [] }; const users = (list) => list.map((email) => ({ getEmail: () => email })); return { getParents: () => { let left = 1; return { hasNext: () => left > 0, next: () => { left -= 1; return { getId: () => "folder" }; } }; }, moveTo: () => {}, getEditors: () => users(share[id].editors), getViewers: () => users(share[id].viewers), addEditor: (e) => { share[id].editors.push(e); }, addViewer: (e) => { share[id].viewers.push(e); } }; } },
      SpreadsheetApp: { openById: (id) => (state.books && state.books[id]) || state.extra, create: (name) => { const book = new Spreadsheet(); book.name = name; book.insertSheet("Sheet1"); (state.books = state.books || {})[book.getId()] = book; return book; }, newFilterCriteria: () => { const c = { values: [], getCriteriaValues: () => c.values }; const b = { whenTextEqualTo: (text) => { c.values = [text]; return b; }, build: () => c }; return b; }, newDataValidation: () => { const rule = { getCriteriaValues: () => [rule.list, true], getAllowInvalid: () => rule.allowInvalid }; const b = { requireValueInList: (list) => { rule.list = list; return b; }, requireValueInRange: (range, show) => { rule.list = range; rule.dropdown = show; return b; }, requireFormulaSatisfied: (f) => { rule.list = f; return b; }, requireDate: () => { rule.date = true; return b; }, requireNumberGreaterThan: (n) => { rule.min = n; return b; }, setHelpText: (text) => { rule.help = text; return b; }, setAllowInvalid: (allow) => { rule.allowInvalid = allow; return b; }, build: () => rule }; return b; }, getActiveSpreadsheet: () => ss, ProtectionType: { RANGE: "RANGE", SHEET: "SHEET" }, CopyPasteType: { PASTE_FORMAT: "PASTE_FORMAT" },
        newConditionalFormatRule: () => { const rule = {}; const b = { whenFormulaSatisfied: (f) => { rule.formula = f; return b; }, setRanges: (r) => { rule.ranges = r; return b; }, setBold: () => b, setBackground: (c) => { rule.bg = c; return b; }, setFontColor: (c) => { rule.color = c; return b; }, build: () => rule }; return b; }, getUi() { throw new Error("no ui"); }, flush() {},
         },
      Utilities: { parseDate: (iso) => "date:" + iso, getUuid: () => require("crypto").randomUUID(), formatDate: (date) => date.toISOString().slice(0, 16).replace("T", " ") },
      CacheService: { getScriptCache: () => ({ get: (k) => (state.cache = state.cache || {})[k] || null, put: (k, v) => { (state.cache = state.cache || {})[k] = v; }, remove: (k) => { delete (state.cache = state.cache || {})[k]; } }) },
      ContentService: { createTextOutput: (text) => ({ text, setMimeType() { return this; } }), MimeType: { JSON: "JSON" } },
      LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) },
      ScriptApp: { getProjectTriggers: () => state.triggers.slice(), deleteTrigger: (t) => { state.triggers = state.triggers.filter(x => x !== t); }, newTrigger: builder, getService: () => ({ getUrl: () => state.serviceUrl || "" }) },
      PropertiesService: { getDocumentProperties: () => ({ getProperty: (k) => (state.docProps = state.docProps || {})[k] ?? null, setProperty: (k, v) => { (state.docProps = state.docProps || {})[k] = v; }, deleteProperty: (k) => { delete (state.docProps = state.docProps || {})[k]; } }), getScriptProperties: () => ({ getProperty: (k) => (state.props = state.props || {})[k] || null, setProperty: (k, v) => { (state.props = state.props || {})[k] = v; }, deleteProperty: (k) => { delete (state.props = state.props || {})[k]; } }) },
      MailApp: { sendEmail: (...a) => state.mails.push(a) }, Session: { getEffectiveUser: () => ({ getEmail: () => "owner@example.com" }), getActiveUser: () => ({ getEmail: () => state.actor || "" }) } };
    vm.createContext(ctx);
    vm.runInContext(patch(code.replace("const DRY_RUN = true;", `const DRY_RUN = ${dry};`)), ctx);
    const out = vm.runInContext(`JSON.stringify(${expr})`, ctx);
    return { value: out === undefined ? null : JSON.parse(out), logs: logs.join("\n") };
  };
  return makeRunner;
}

module.exports = { WHITE, cell, Protection, Range, Sheet, Spreadsheet, makeRunnerFor };
