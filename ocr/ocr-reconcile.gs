/**
 * Receipt OCR - for the price-watch site ("Баримт" tab, 2026-10-23): per day, the workbook cells
 * of that day compared with the receipts, one row per cell:
 *   match      mom typed it and the receipts make the same number (the cell is light blue)
 *   diff       mom typed another number (light red, its note says "Баримтаар: n")
 *   filled     mom had typed nothing: the receipts' number was written (yellow, green once approved)
 *   otherday   mom typed it on another day: the receipt's day holds 0 (ocr-elsewhere.gs); the cell
 *              named is hers
 *   noname     the receipt's product has no row on that sheet, so nothing was written
 *   noreceipt  mom typed a number on a day with photos and no receipt line is for it
 * Read from the day tab's "Бичсэн" and the day columns of the place's own folders (a hand-typed
 * cell a receipt of another place's spreadsheet is for counts as noreceipt here). Kept on the
 * month spreadsheet's tab OCR_CHECKS_SHEET, one row a cell, for anyone to read (Script properties
 * are too small for them), and read from there by the web app (ocrFeedAnswer).
 */

const OCR_CHECKS_SHEET = 'Тулгалт';
const OCR_OLD_CHECKS_SHEET = '_тулгалт'; // the hidden JSON tab of 2026-10-23 morning, taken away
const OCR_CHECKS_HEADER = ['Өдөр', 'Бараа', 'Ээж бичсэн', 'Баримтаар', 'Төлөв', 'Нүд', 'Код'];
const OCR_CHECK_LABELS = {
  match: 'таарсан', diff: 'зөрсөн', filled: 'систем нөхсөн', otherday: 'өөр өдөрт бичсэн', noname: 'нэр хуудсанд алга', noreceipt: 'баримт алга',
};
const OCR_CHECKS_DAY_LIMIT = 200;
const OCR_CHECKS_NAME_LIMIT = 80;
const OCR_WRITTEN_CHECK_LINE = /^(check:)?(\S+)!([A-Z]+\d+) ← (-?[\d.]+) · (.+)$/;
const OCR_NO_ROW_REASON = /^"(.+)" (\S+) дээр алга$/; // ocrFindCell's reason

// [{ name, hand, receipt, status, sheet, cell }] of one day tab of entry ({ place, month }).
function ocrDayChecks(records, entry, day, workbooks) {
  const written = ocrWrittenChecks(records, workbooks);
  const known = new Set(written.map((check) => `${check.sheet}!${check.cell}`));
  return ocrLimitChecks([
    ...written,
    ...records.flatMap(ocrNoRowChecks),
    ...ocrNoReceiptChecks(records, entry.place, `${entry.month}-${day}`, known, workbooks),
  ]);
}

const ocrLimitChecks = (checks) =>
  checks.slice(0, OCR_CHECKS_DAY_LIMIT).map((check) => ({ ...check, name: check.name.slice(0, OCR_CHECKS_NAME_LIMIT) }));

// The cells named in "Бичсэн" (a failed line's cells written before too), the receipts' lines of
// one cell added up. A check copy's cell is named by mom's own cell: the product's row of the
// original, as rows may have been added there since the copy was made.
function ocrWrittenChecks(records, workbooks) {
  const cells = new Map();
  records.flatMap((record) => ocrWrittenLines(record['Бичсэн'])).forEach((line) => {
    const parsed = ocrOriginalCell(ocrParseWrittenLine(line), workbooks);
    const key = `${parsed.sheet}!${parsed.cell}`;
    const before = cells.get(key);
    cells.set(key, before ? { ...parsed, receipt: before.receipt + parsed.receipt } : parsed);
  });
  return [...cells.values()].map((check) => ({
    name: check.name,
    hand: check.status === 'match' || check.status === 'otherday' ? check.receipt : check.hand,
    receipt: check.receipt,
    status: check.status,
    sheet: check.sheet,
    cell: check.cell,
  }));
}

// Every cell by its product's row now, not only a check copy's (the owner, 2026-10-29: a line waiting
// for mom keeps its "Бичсэн", and rows put in above moved its product down).
function ocrOriginalCell(parsed, workbooks) {
  const place = ocrPlaceOfTab(parsed.sheet);
  const tab = place && OCR_POST_BOOKS[place] ? ocrWorkbookTab(workbooks, ocrBookIdFor(place, parsed.sheet), parsed.sheet) : { error: true };
  const row = tab.error ? null : tab.rows.get(ocrNameKey(parsed.name));
  return row ? { ...parsed, cell: parsed.cell.replace(/\d+$/, String(row)) } : parsed;
}

// "check:My/2026/10/15!W5 ← 1 · Meiling 501l ✓ (гараар 3)" -> its cell, number, name and status.
function ocrParseWrittenLine(line) {
  const elsewhere = ocrElsewhereOfLine(line);
  if (elsewhere) {
    const [, , sheet, , qty, rest] = line.replace(OCR_ELSEWHERE_LINE, '').match(OCR_WRITTEN_CHECK_LINE);
    const name = ocrWithoutMark(rest);
    return { name, sheet, cell: elsewhere.cell, receipt: Number(qty), status: 'otherday', hand: Number(qty) };
  }
  const [, check, sheet, cell, qty, rest] = line.match(OCR_WRITTEN_CHECK_LINE);
  const status = check ? (rest.match(/ \(([^()]*)\)$/) || [])[1] || '' : '';
  const name = ocrWithoutMark(status ? rest.slice(0, rest.length - status.length - 3) : rest);
  const typed = status.replace(/^гараар /, '');
  return {
    name,
    sheet,
    cell,
    receipt: Number(qty),
    status: !check ? 'filled' : status === OCR_HAND_MATCH_TEXT ? 'match' : 'diff',
    hand: !check || typed === 'хоосон' ? null : Number.isNaN(Number(typed)) ? typed : Number(typed),
  };
}

// A line not written because its product has no row on the sheet.
function ocrNoRowChecks(record) {
  const decision = ocrPostDecision(record);
  if (!decision.post || !ocrText(record['Бичсэн']).startsWith(OCR_POST_FAILED)) {
    return [];
  }
  return ocrFailedReason(record['Бичсэн']).split('; ')
    .map((reason) => reason.match(OCR_NO_ROW_REASON))
    .filter(Boolean)
    .map(([, name, sheet]) => ({ name, hand: null, receipt: decision.qty, status: 'noname', sheet, cell: '' }));
}

// The hand-typed cells of the day's columns of the place itself (income, expense) that no line of
// the tab is for.
function ocrNoReceiptChecks(records, place, iso, known, workbooks) {
  const columns = new Map();
  records.forEach((record) => {
    ocrFolderMeaning(ocrText(record['Хавтас (ээж хийсэн)']).split(' / ')).moves
      .filter((move) => move.place === place && OCR_POST_BOOKS[place])
      .map((move) => ocrTarget(move, iso))
      .filter((target) => target && ![target.layout.transferIn, target.layout.transferOut, target.layout.count].includes(target.column))
      .forEach((target) => columns.set(`${target.tab}|${target.column}`, target));
  });
  return [...columns.values()].flatMap((target) => {
    const grid = ocrTabGrid(workbooks, ocrBookIdFor(place, target.tab), target.tab);
    if (!grid) {
      return [];
    }
    const letters = columnLetter(target.column);
    return grid.values.flatMap((row, index) => {
      const value = row[target.column - 1];
      const note = grid.notes[index][target.column - 1];
      const cell = `${letters}${OCR_BOOK_FIRST_ROW + index}`;
      if (ocrText(value) === '' || !ocrIsHandCell(value, note) || known.has(`${target.tab}!${cell}`) ||
        ocrTakenElsewhere(grid.notes[index], null).has(cell)) {
        return [];
      }
      return [{ name: ocrText(row[0]), hand: value, receipt: null, status: 'noreceipt', sheet: target.tab, cell }];
    });
  });
}

// A workbook tab's product rows, values, notes and colours, read once per run (cells written in the
// run are not in it); null when it can not be read.
function ocrTabGrid(workbooks, bookId, tabName) {
  const key = `grid|${bookId}|${tabName}`;
  if (!workbooks.has(key)) {
    const tab = ocrWorkbookTab(workbooks, bookId, tabName);
    const count = tab.error ? 0 : tab.sheet.getLastRow() - OCR_BOOK_FIRST_ROW + 1;
    const range = count > 0 ? tab.sheet.getRange(OCR_BOOK_FIRST_ROW, 1, count, tab.sheet.getLastColumn()) : null;
    workbooks.set(key, range ? { values: range.getValues(), notes: range.getNotes(), backgrounds: range.getBackgrounds() } : null);
  }
  return workbooks.get(key);
}

// results: [{ day, checks }] of one month spreadsheet { place, month }, onto its Тулгалт tab
// (written only when changed; the hidden tab before it is taken away).
function ocrKeepChecks(book, entry, results) {
  const rows = results
    .filter((result) => result.checks.length > 0)
    .sort((a, b) => a.day.localeCompare(b.day))
    .flatMap((result) => result.checks.map((check) => ocrCheckRow(result.day, check)));
  if (DRY_RUN) {
    Logger.log(`Тулгалт ${entry.place} ${entry.kind.word} ${entry.month}: ${rows.length} мөр (туршилт, хадгалаагүй)`);
    return;
  }
  ocrKeepFeedChecks(book.getId(), rows);
  const old = book.getSheetByName(OCR_OLD_CHECKS_SHEET);
  if (old) {
    book.deleteSheet(old);
  }
  let tab = book.getSheetByName(OCR_CHECKS_SHEET);
  if (!tab && rows.length === 0) {
    return;
  }
  if (!tab) {
    tab = book.insertSheet(OCR_CHECKS_SHEET, book.getSheets().length);
    ocrOnlyOwner(tab.protect().setDescription(`${OCR_GUARD_PREFIX}: тулгалт`), Session.getEffectiveUser().getEmail());
  }
  const wanted = [OCR_CHECKS_HEADER, ...rows];
  const kept = tab.getLastRow() > 0 ? tab.getRange(1, 1, tab.getLastRow(), OCR_CHECKS_HEADER.length).getValues() : [];
  if (JSON.stringify(kept) === JSON.stringify(wanted)) {
    return;
  }
  tab.clear();
  tab.getRange(1, 1, wanted.length, 1).setNumberFormat('@'); // "02" stays a day, not the number 2
  tab.getRange(1, 1, wanted.length, OCR_CHECKS_HEADER.length).setValues(wanted);
  tab.setFrozenRows(1);
}

const ocrCheckRow = (day, check) => [
  day, check.name, check.hand === null ? '' : check.hand, check.receipt === null ? '' : check.receipt,
  OCR_CHECK_LABELS[check.status], check.cell ? `${check.sheet}!${check.cell}` : check.sheet, check.status,
];

// A month spreadsheet's Тулгалт rows kept in Script properties too, gzipped, so the feed (doGet) does
// not open it (the owner, 2026-10-29: opening the month's 8 spreadsheets took a doGet 10-16 s, the
// site gives up at 12 s and shows the answer of before, so the days mom had checked stayed blue).
const OCR_CHECKS_PROPERTY_PREFIX = 'ocr-checks:'; // + month spreadsheet id
function ocrKeepFeedChecks(bookId, rows) {
  const props = PropertiesService.getScriptProperties();
  const packed = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(JSON.stringify(rows))).getBytes());
  try {
    props.setProperty(OCR_CHECKS_PROPERTY_PREFIX + bookId, packed);
  } catch (error) {
    props.deleteProperty(OCR_CHECKS_PROPERTY_PREFIX + bookId); // too big for a property: the feed reads the tab
    Logger.log(`Тулгалт (${bookId}) Script properties-д багтсангүй: ${error.message}`);
  }
}

// The rows ocrKeepFeedChecks kept; null when none are.
function ocrKeptFeedChecks(bookId) {
  const packed = PropertiesService.getScriptProperties().getProperty(OCR_CHECKS_PROPERTY_PREFIX + bookId);
  if (!packed) {
    return null;
  }
  try {
    return JSON.parse(Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(packed), 'application/x-gzip')).getDataAsString());
  } catch (error) {
    Logger.log(`Тулгалт (${bookId}) уншигдсангүй: ${error.message}, табаас уншина`);
    return null;
  }
}

// The checks of a month spreadsheet for the feed: { "DD": [check] }; {} when there are none. Read from
// the tab (and kept, ocrKeepFeedChecks) only when none were kept.
function ocrReadChecks(bookId) {
  const kept = ocrKeptFeedChecks(bookId);
  if (kept) {
    return ocrFeedChecks(kept);
  }
  try {
    const book = SpreadsheetApp.openById(bookId);
    const tab = book.getSheetByName(OCR_CHECKS_SHEET);
    if (tab && tab.getLastRow() > 1) {
      const rows = tab.getRange(2, 1, tab.getLastRow() - 1, OCR_CHECKS_HEADER.length).getValues();
      if (!DRY_RUN) {
        ocrKeepFeedChecks(bookId, rows);
      }
      return ocrFeedChecks(rows);
    }
    const old = book.getSheetByName(OCR_OLD_CHECKS_SHEET); // not looked at again since
    return old && old.getLastRow() > 0 ? ocrOldFeedChecks(old.getRange(1, 1, old.getLastRow(), 1).getValues()) : {};
  } catch (error) {
    Logger.log(`Тулгалт уншигдсангүй (${bookId}): ${error.message}`);
    return {};
  }
}

// Тулгалт rows -> { "DD": [{ name, hand, receipt, status, sheet, cell }] }.
function ocrFeedChecks(rows) {
  const days = {};
  rows.filter((row) => ocrText(row[0]) !== '').forEach(([day, name, hand, receipt, , place, status]) => {
    const key = ocrText(day).padStart(2, '0');
    const [sheet, cell = ''] = ocrText(place).split('!');
    days[key] = [...(days[key] || []), { name: ocrText(name), hand: hand === '' ? null : hand, receipt: receipt === '' ? null : receipt, status: ocrText(status), sheet, cell }];
  });
  return Object.fromEntries(Object.entries(days).map(([day, checks]) => [day, ocrLimitChecks(checks)]));
}

const ocrOldFeedChecks = (rows) =>
  Object.fromEntries(rows.filter(([text]) => ocrText(text) !== '').map(([text]) => JSON.parse(text)).map((kept) => [kept.day, ocrLimitChecks(kept.checks)]));
