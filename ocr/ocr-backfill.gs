/**
 * Receipt OCR - Тосгон's expense of the days before OCR_POST_FROM, from what the shops typed by
 * hand. The big goods (those on Тосгон's sheet) always reach a shop through Тосгон (the owner's
 * rule; small goods come straight from suppliers), so a shop's income of such a product on a day
 * is Тосгон's expense of that day. From then on the OCR writes both sides of a transfer itself.
 *
 * Only empty expense cells of Тосгон are filled, purple on yellow (the system is sure of them; the
 * owner, 2026-10-28: three colours only; a colour a cell was given stays), each with a note that links to the shops'
 * income cells it comes from (since 2026-10-26, the owner: the spreadsheet's link, as the OCR's
 * notes): "Дэлгүүрийн орлогоос:" and a line a cell, "https://docs.google.com/.../edit#gid=7&range=H9 = 2 (Соронзон)".
 * A note of before ("Дэлгүүрийн орлогоос: My home 2, Соронзон 1") is still the backfill's own;
 * linkTosgonBackfillNotes() rewrites those in the new way once, the numbers left as they are
 * (a cell whose number is no longer the shops' income is listed, not changed). A cell typed by hand or written by the OCR is
 * left as it is and reported. Run again, it only puts its own cells right (and empties one whose
 * shop income is gone). Up to OCR_TOSGON_HAND_UNTIL Тосгон typed its own expense, often on another
 * day than the shop's income, so for those days a product only gets what its hand-typed numbers
 * leave over (shop income - hand-typed, earliest days first); the rest would be counted twice.
 *   previewTosgonBackfill(): what it would write, in the log; runTosgonBackfill(): writes it;
 *   colorTosgonBackfillCells(): the cells filled before 2026-10-28 turn yellow, their numbers left alone.
 */

const OCR_BACKFILL_MARK = 'Дэлгүүрийн орлогоос:';
const OCR_BACKFILL_FONT_COLOR = '#8e24aa';
const OCR_BACKFILL_YEAR = 2026; // the year the shops' workbook holds
const OCR_BACKFILL_LIST_LIMIT = 10;

// lines: [{ key, day, place, qty, link }] of the shops' income; tosgonRows: product key -> row of
// Тосгон's tab. The cells to fill: "row|day" -> { row, day, total, parts: { place: qty }, sources:
// [{ place, qty, link }] (a shop cell each) }, and how many lines were left out as small goods (not
// on Тосгон's sheet).
function ocrBackfillCells(lines, tosgonRows) {
  const cells = new Map();
  let small = 0;
  lines.forEach((line) => {
    const row = tosgonRows.get(line.key);
    if (!row) {
      small += 1;
      return;
    }
    const id = `${row}|${line.day}`;
    const cell = cells.get(id) || { row, day: line.day, total: 0, parts: {}, sources: [] };
    const same = cell.sources.find((source) => source.link === line.link && source.place === line.place);
    const sources = same
      ? cell.sources.map((source) => (source === same ? { ...source, qty: source.qty + line.qty } : source))
      : [...cell.sources, { place: line.place, qty: line.qty, link: line.link }];
    cells.set(id, { ...cell, total: cell.total + line.qty, parts: { ...cell.parts, [line.place]: (cell.parts[line.place] || 0) + line.qty }, sources });
  });
  return { cells, small };
}

// A filled cell's note: its shop cells' links, by shop; the shops' names only when no link is known.
function ocrBackfillNote(cell) {
  const sources = (cell.sources || []).filter((source) => source.link);
  if (sources.length === 0) {
    return `${OCR_BACKFILL_MARK} ${Object.keys(cell.parts).sort().map((place) => `${place} ${cell.parts[place]}`).join(', ')}`;
  }
  const byPlace = [...sources].sort((a, b) => a.place.localeCompare(b.place));
  return [OCR_BACKFILL_MARK, ...byPlace.map((source) => `${source.link} = ${source.qty} (${source.place})`)].join('\n');
}

const ocrIsBackfillNote = (note) => ocrText(note).startsWith(OCR_BACKFILL_MARK);

// A backfill cell's colour: yellow while it holds a number and had no colour, none once emptied
// (only the yellow goes); any other colour it was given stays.
const ocrIsNoColor = (color) => ['', '#ffffff', 'null'].includes(String(color || '').toLowerCase());
function ocrBackfillBackground(current, hasNumber) {
  if (hasNumber) {
    return ocrIsNoColor(current) ? OCR_WRITTEN_COLOR : current;
  }
  return String(current || '').toLowerCase() === OCR_WRITTEN_COLOR ? null : current;
}

// Backgrounds to write back: a cell read as white had no colour (Sheets reads none as #ffffff), and
// writing white would give it one.
const ocrBackgroundsToWrite = (backgrounds) => backgrounds.map((row) => row.map((color) => (ocrIsNoColor(color) ? null : color)));

// The last day of `month` (of OCR_BACKFILL_YEAR) on which Тосгон typed its own expense; 0: none.
function ocrHandLastDay(month) {
  const [year, lastMonth, lastDay] = OCR_TOSGON_HAND_UNTIL.split('-').map(Number);
  if (year > OCR_BACKFILL_YEAR || (year === OCR_BACKFILL_YEAR && month < lastMonth)) {
    return daysInMonth(OCR_BACKFILL_YEAR, month);
  }
  return year === OCR_BACKFILL_YEAR && month === lastMonth ? lastDay : 0;
}

// Cuts the cells of the hand-typed days to what they leave over. handTotals: row -> qty Тосгон
// typed on days up to handLastDay; isWritable(cell): the cell is empty or one of ours (a cell typed
// by hand is never written, so it takes no share). Returns { cells: a new Map, cut: qty left out };
// a cut cell keeps { ..., total: what is left, cut }.
function ocrTrimBackfill(cells, handTotals, handLastDay, isWritable) {
  const handDays = [...cells.values()].filter((cell) => cell.day <= handLastDay);
  if (handDays.length === 0) {
    return { cells, cut: 0 };
  }
  const shopTotals = handDays.reduce((totals, cell) => totals.set(cell.row, (totals.get(cell.row) || 0) + cell.total), new Map());
  const left = new Map([...shopTotals].map(([row, total]) => [row, Math.max(0, total - (handTotals.get(row) || 0))]));
  const trimmed = new Map(cells);
  let cut = 0;
  handDays
    .filter(isWritable)
    .sort((a, b) => a.day - b.day)
    .forEach((cell) => {
      const take = Math.min(cell.total, left.get(cell.row));
      left.set(cell.row, left.get(cell.row) - take);
      cut += cell.total - take;
      const id = `${cell.row}|${cell.day}`;
      if (take === 0) {
        trimmed.delete(id);
      } else if (take < cell.total) {
        trimmed.set(id, { ...cell, total: take, cut: cell.total - take });
      }
    });
  return { cells: trimmed, cut };
}

// Qty Тосгон typed (not ours) on each row up to handLastDay. values / notes: the expense block.
function ocrHandTotals(values, notes, handLastDay) {
  return values.reduce((totals, row, r) => {
    const sum = row
      .slice(0, handLastDay)
      .reduce((total, value, d) => total + (typeof value === 'number' && value > 0 && !ocrIsBackfillNote(notes[r][d]) ? value : 0), 0);
    return sum > 0 ? totals.set(r + OCR_BOOK_FIRST_ROW, sum) : totals;
  }, new Map());
}

// A shop tab's income lines [{ key, day, place, qty, link }] of the days up to lastDay. values: the
// rows from OCR_BOOK_FIRST_ROW, columns A to the last income day; firstDay: the day of the first
// income column; tab: { book, gid } of the tab, for the link to each line's cell (none without it).
// The sheet's last day column holds day 30 and 31, so a day is at most OCR_MONTH_DAYS.
function ocrShopIncomeLines(values, place, firstDay, lastDay, tab = null) {
  return values.flatMap((row, r) => {
    const key = ocrNameKey(row[0]);
    if (key === '') {
      return [];
    }
    const linkOf = (index) => (tab ? sheetCellLink(tab.book, tab.gid, `${columnLetter(LAYOUT.incomeFirst + index)}${OCR_BOOK_FIRST_ROW + r}`) : null);
    return row.slice(LAYOUT.incomeFirst - 1)
      .map((qty, index) => ({ key, day: Math.min(firstDay + index, OCR_MONTH_DAYS), place, qty, link: linkOf(index) }))
      .filter((line) => typeof line.qty === 'number' && line.qty > 0 && line.day <= lastDay);
  });
}

// The income lines of both shops in a month, up to lastDay.
function ocrShopIncomeOfMonth(month, lastDay, workbooks) {
  return CHAINS.flatMap((chain) => [0, 1].flatMap((half) => {
    const firstDay = half === 0 ? 1 : FIRST_HALF_LAST_DAY + 1;
    const tab = ocrWorkbookTab(workbooks, OCR_CATALOG_SPREADSHEET_ID, chainSheetName(chain, periodPosition(OCR_BACKFILL_YEAR, month, half)));
    if (tab.error || firstDay > lastDay) {
      return [];
    }
    const count = tab.sheet.getLastRow() - OCR_BOOK_FIRST_ROW + 1;
    const width = LAYOUT.incomeFirst + DAY_COLUMNS - 1;
    const values = count > 0 ? tab.sheet.getRange(OCR_BOOK_FIRST_ROW, 1, count, width).getValues() : [];
    return ocrShopIncomeLines(values, chain.label, firstDay, lastDay, { book: tab.sheet.getParent().getId(), gid: tab.sheet.getSheetId() });
  }));
}

// What the shops' income makes of one month of Тосгон, read once for writing or linking:
// { tab, missing } | { tab, small, empty: true } | { tab, sheet, block, values, notes, colors, cells, cut, small }.
function ocrBackfillPlan(month, lastDay, workbooks) {
  const name = `${OCR_TOSGON.tabPrefix}/${OCR_BACKFILL_YEAR}/${month}`;
  const tosgon = ocrWorkbookTab(workbooks, OCR_TOSGON_SPREADSHEET_ID, name);
  if (tosgon.error) {
    return { tab: name, missing: tosgon.error };
  }
  const shopCells = ocrBackfillCells(ocrShopIncomeOfMonth(month, lastDay, workbooks), tosgon.rows);
  const rowCount = tosgon.sheet.getLastRow() - OCR_BOOK_FIRST_ROW + 1;
  if (rowCount < 1) {
    return { tab: name, small: shopCells.small, empty: true };
  }
  const block = tosgon.sheet.getRange(OCR_BOOK_FIRST_ROW, OCR_TOSGON.layout.expenseFirst, rowCount, OCR_MONTH_DAYS);
  const values = block.getValues();
  const notes = block.getNotes();
  const colors = block.getFontColors();
  const backgrounds = block.getBackgrounds();
  const handLastDay = ocrHandLastDay(month);
  const isWritable = (cell) => {
    const [r, d] = [cell.row - OCR_BOOK_FIRST_ROW, cell.day - 1];
    return ocrIsBackfillNote(notes[r][d]) || (ocrText(values[r][d]) === '' && ocrText(notes[r][d]) === '');
  };
  const { cells, cut } = ocrTrimBackfill(shopCells.cells, ocrHandTotals(values, notes, handLastDay), handLastDay, isWritable);
  return { tab: name, sheet: tosgon.sheet, block, values, notes, colors, backgrounds, cells, cut, small: shopCells.small };
}

const ocrBackfillCellNote = (cell) =>
  ocrBackfillNote(cell) + (cell.cut ? ` (Тосгонд гараар бичсэнтэй давхардах ${cell.cut}-г хассан)` : '');

// One month of Тосгон: { tab, cells, qty, small, kept: [text] } or { tab, missing }.
function ocrBackfillMonth(month, lastDay, workbooks, isWrite) {
  const plan = ocrBackfillPlan(month, lastDay, workbooks);
  if (plan.missing || plan.empty) {
    return plan.missing ? plan : { tab: plan.tab, cells: 0, qty: 0, small: plan.small, kept: [], cut: 0 };
  }
  const { values, notes, colors, backgrounds, cells } = plan;
  // own cells whose shop income is gone are emptied
  notes.forEach((noteRow, r) => noteRow.forEach((note, d) => {
    if (ocrIsBackfillNote(note) && !cells.has(`${r + OCR_BOOK_FIRST_ROW}|${d + 1}`)) {
      values[r][d] = '';
      notes[r][d] = '';
      colors[r][d] = null;
      backgrounds[r][d] = ocrBackfillBackground(backgrounds[r][d], false);
    }
  }));
  const kept = [];
  let written = 0;
  let qty = 0;
  [...cells.values()].forEach((cell) => {
    const r = cell.row - OCR_BOOK_FIRST_ROW;
    const d = cell.day - 1;
    const isOwn = ocrIsBackfillNote(notes[r][d]);
    if (!isOwn && (ocrText(values[r][d]) !== '' || ocrText(notes[r][d]) !== '')) {
      kept.push(`${ocrText(plan.sheet.getRange(cell.row, 1).getValue())} ${cell.day}-нд ${values[r][d]} (дэлгүүрт ${cell.total})`);
      return;
    }
    values[r][d] = cell.total;
    notes[r][d] = ocrBackfillCellNote(cell);
    colors[r][d] = OCR_BACKFILL_FONT_COLOR;
    backgrounds[r][d] = ocrBackfillBackground(backgrounds[r][d], true);
    written += 1;
    qty += cell.total;
  });
  if (isWrite) {
    plan.block.setValues(values).setNotes(notes).setFontColors(colors).setBackgrounds(ocrBackgroundsToWrite(backgrounds));
  }
  return { tab: plan.tab, cells: written, qty, small: plan.small, kept, cut: plan.cut };
}

// One month's own notes in the new way, the numbers left alone: { tab, linked, left: [text], empty }
// or { tab, missing }. A cell whose number is not what the shops' income makes now is listed
// (left), and an empty one with the backfill's note (a copy of a new row, row-copies.gs) is counted.
function ocrLinkBackfillMonth(month, lastDay, workbooks, isWrite) {
  const plan = ocrBackfillPlan(month, lastDay, workbooks);
  if (plan.missing || plan.empty) {
    return plan.missing ? plan : { tab: plan.tab, linked: 0, left: [], empty: 0 };
  }
  const { values, notes, cells } = plan;
  const left = [];
  let linked = 0;
  let empty = 0;
  notes.forEach((noteRow, r) => noteRow.forEach((note, d) => {
    if (!ocrIsBackfillNote(note)) {
      return;
    }
    const cell = cells.get(`${r + OCR_BOOK_FIRST_ROW}|${d + 1}`);
    if (ocrText(values[r][d]) === '') {
      empty += 1;
      return;
    }
    if (!cell || cell.total !== values[r][d]) {
      left.push(`${columnLetter(OCR_TOSGON.layout.expenseFirst + d)}${r + OCR_BOOK_FIRST_ROW} = ${values[r][d]} (дэлгүүрт ${cell ? cell.total : 0})`);
      return;
    }
    const wanted = ocrBackfillCellNote(cell);
    if (wanted !== note) {
      notes[r][d] = wanted;
      linked += 1;
    }
  }));
  if (isWrite && linked > 0) {
    plan.block.setNotes(notes);
  }
  return { tab: plan.tab, linked, left, empty };
}

// The months of OCR_BACKFILL_YEAR before OCR_POST_FROM, with the last day of each to take.
function ocrBackfillMonths() {
  const [year, month, day] = OCR_POST_FROM.split('-').map(Number);
  const lastMonth = year > OCR_BACKFILL_YEAR ? MONTHS_PER_YEAR : month;
  return Array.from({ length: lastMonth }, (_, index) => index + 1)
    .map((m) => ({ month: m, lastDay: year === OCR_BACKFILL_YEAR && m === month ? day - 1 : daysInMonth(OCR_BACKFILL_YEAR, m) }))
    .filter((entry) => entry.lastDay >= 1);
}

function ocrBackfillTosgon(isWrite) {
  const workbooks = new Map();
  const results = ocrBackfillMonths().map(({ month, lastDay }) => ocrBackfillMonth(month, lastDay, workbooks, isWrite));
  const verb = isWrite ? 'бичлээ' : 'бичнэ (урьдчилж харав, юу ч өөрчлөөгүй)';
  results.forEach((result) => {
    if (result.missing) {
      Logger.log(`${result.tab}: ${result.missing}`);
      return;
    }
    Logger.log(`${result.tab}: ${result.cells} нүд, ${result.qty} ширхэг ${verb}; Тосгонд байхгүй жижиг барааны ${result.small} мөр алгассан` +
      (result.cut > 0 ? `; Тосгонд гараар бичсэнтэй давхардах ${result.cut} ширхгийг оруулаагүй` : '') +
      (result.kept.length > 0 ? `; гараар/OCR-аар бичсэн ${result.kept.length} нүдийг хөндсөнгүй: ${result.kept.slice(0, OCR_BACKFILL_LIST_LIMIT).join('; ')}` : ''));
  });
  return results;
}

function previewTosgonBackfill() {
  return ocrBackfillTosgon(false);
}

// Run by hand once (2026-10-26): the backfill's notes of before get the links to the shop cells.
function linkTosgonBackfillNotes() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(OCR_LOCK_WAIT_MS)) {
    Logger.log('OCR ажиллаж байна, хэсэг хүлээгээд дахин Run хийнэ үү.');
    return null;
  }
  try {
    const workbooks = new Map();
    const results = ocrBackfillMonths().map(({ month, lastDay }) => ocrLinkBackfillMonth(month, lastDay, workbooks, !DRY_RUN));
    results.forEach((result) => Logger.log(result.missing ? `${result.tab}: ${result.missing}` :
      `${result.tab}: ${result.linked} нүдний тэмдэглэл ${DRY_RUN ? 'холбоостой болно (туршилт)' : 'холбоостой боллоо'}` +
      (result.left.length > 0 ? `; тоо нь дэлгүүрийн орлоготой таарахгүй ${result.left.length} нүдийг хөндсөнгүй: ${result.left.slice(0, OCR_BACKFILL_LIST_LIMIT).join('; ')}` : '') +
      (result.empty > 0 ? `; хоосон нүдэнд үлдсэн ${result.empty} тэмдэглэл (үндсэн хүснэгтийн clearCopiedNotes арилгана)` : '')));
    return results;
  } finally {
    lock.releaseLock();
  }
}

// Run by hand once (2026-10-28): the backfill's cells of before (purple, no colour) turn yellow,
// as every number the system is sure of; their numbers and notes are left as they are.
function colorTosgonBackfillCells() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(OCR_LOCK_WAIT_MS)) {
    Logger.log('OCR ажиллаж байна, хэсэг хүлээгээд дахин Run хийнэ үү.');
    return null;
  }
  try {
    const workbooks = new Map();
    return ocrBackfillMonths().map(({ month }) => {
      const result = ocrColorBackfillMonth(month, workbooks, !DRY_RUN);
      Logger.log(result.missing ? `${result.tab}: ${result.missing}` : `${result.tab}: ${result.colored} нүд шар ${DRY_RUN ? 'болно (туршилт)' : 'боллоо'}`);
      return result;
    });
  } finally {
    lock.releaseLock();
  }
}

// One month of Тосгон: { tab, colored } or { tab, missing }.
function ocrColorBackfillMonth(month, workbooks, isWrite) {
  const name = `${OCR_TOSGON.tabPrefix}/${OCR_BACKFILL_YEAR}/${month}`;
  const tosgon = ocrWorkbookTab(workbooks, OCR_TOSGON_SPREADSHEET_ID, name);
  if (tosgon.error) {
    return { tab: name, missing: tosgon.error };
  }
  const rowCount = tosgon.sheet.getLastRow() - OCR_BOOK_FIRST_ROW + 1;
  if (rowCount < 1) {
    return { tab: name, colored: 0 };
  }
  const block = tosgon.sheet.getRange(OCR_BOOK_FIRST_ROW, OCR_TOSGON.layout.expenseFirst, rowCount, OCR_MONTH_DAYS);
  const values = block.getValues();
  const notes = block.getNotes();
  const before = block.getBackgrounds();
  const after = before.map((row, r) => row.map((color, d) =>
    (ocrIsBackfillNote(notes[r][d]) && ocrText(values[r][d]) !== '' ? ocrBackfillBackground(color, true) : color)));
  const colored = after.flat().filter((color, at) => color !== before.flat()[at]).length;
  if (isWrite && colored > 0) {
    block.setBackgrounds(ocrBackgroundsToWrite(after));
  }
  return { tab: name, colored };
}

function runTosgonBackfill() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(OCR_LOCK_WAIT_MS)) {
    Logger.log('OCR ажиллаж байна, хэсэг хүлээгээд дахин Run хийнэ үү.');
    return null;
  }
  try {
    return ocrBackfillTosgon(!DRY_RUN);
  } finally {
    lock.releaseLock();
  }
}
