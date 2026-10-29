/**
 * Receipt OCR - a count, read for the count report (ocr-count-report.gs). A count is the "тооллого"
 * lines of a place's day tab ("OCR Соронзон 2026-10" / "08"), as they would be written (OK, or
 * checked by mom), compared with the books at the end of that day:
 *   opening balance (B) + income of the days up to it - expense of the days up to it
 *   + transfers in up to it - transfers out up to it
 * so a receipt of that day or before counts even when it came in late, and nothing of a later day
 * ever does. The transfer columns (T, AJ) have no day columns, so their notes give the day of each
 * line ("... = 2 @2026-10-03", ocr-post.gs); a number typed there by hand counts as before.
 * Count after closing, or put a count made before opening in the day before.
 * Only on the last day of its sheet (shop: 15th, month end; Тосгон: month end) is a count also
 * written into the count column, which is carried to the next sheet (ocrIsPeriodEnd).
 */

const ocrNumber = (value) => (typeof value === 'number' ? value : 0);

// A transfer cell's part up to `day`: what the note says by day, plus what was typed by hand.
function ocrTransferUpTo(value, note, day) {
  const entries = ocrNoteEntries(note) || [];
  const noted = entries.reduce((total, entry) => total + entry.qty, 0);
  const upTo = entries.filter((entry) => entry.day === '' || entry.day <= day).reduce((total, entry) => total + entry.qty, 0);
  return ocrNumber(value) - noted + upTo;
}

// A product row's balance at the end of the day whose day column is `offset` from the first one.
function ocrBalanceAsOf(row, notes, layout, offset, day) {
  const sum = (first) => row.slice(first - 1, first + offset).reduce((total, value) => total + ocrNumber(value), 0);
  // gifts (the shops' gift column) carry no day: all of the sheet's are taken off
  const gifts = layout.gift ? ocrNumber(row[layout.gift - 1]) : 0;
  return ocrNumber(row[1]) + sum(layout.incomeFirst) - sum(layout.expenseFirst)
    + ocrTransferUpTo(row[layout.transferIn - 1], notes.in, day) - ocrTransferUpTo(row[layout.transferOut - 1], notes.out, day) - gifts;
}

// The count lines of a day tab that may be counted (as they would be written), one per product:
// [{ name, counted }].
function ocrCountedLines(records) {
  const counted = new Map();
  records
    .filter((record) => ocrFolderMeaning(ocrText(record['Хавтас (ээж хийсэн)']).split(' / ')).kind === 'toollogo')
    .map((record) => ocrPostDecision(record))
    .filter((decision) => decision.post && ocrValueProblem(decision) === '')
    .forEach(({ name, qty }) => {
      const known = counted.get(ocrNameKey(name));
      counted.set(ocrNameKey(name), { name: known ? known.name : name, counted: (known ? known.counted : 0) + qty });
    });
  return [...counted.values()];
}

// The month spreadsheet of a place's counts, or null (never made here).
function ocrFindMonthBook(root, place, month) {
  const kind = OCR_BOOK_KINDS.toollogo;
  const folder = [month, kind.folder].reduce((parent, name) => parent && ocrChildFolder(parent, name), ocrSheetsFolder(root, false));
  const files = folder ? folder.getFilesByName(ocrMonthBookName(place, kind, month)) : null;
  return files && files.hasNext() ? SpreadsheetApp.openById(files.next().getId()) : null;
}

// The count of a place on a day: { counted: [{ name, counted }], balances: [{ name, computed }],
// isClosing, tab } (balances: every product of the workbook sheet at the end of that day), or null
// when there is no workbook sheet for that day.
function ocrReadCountAt(root, place, day) {
  const target = ocrTarget({ place, side: 'count' }, day);
  const sheet = target ? SpreadsheetApp.openById(ocrBookIdFor(place, target.tab)).getSheetByName(target.tab) : null;
  if (!sheet) {
    return null;
  }
  const book = ocrFindMonthBook(root, place, ocrMonthFolderName(day));
  const dayTab = book ? book.getSheetByName(day.slice(8)) : null;
  const records = dayTab ? ocrReadRecords(dayTab) : [];
  const { layout } = target;
  const offset = ocrTarget({ place, side: 'in' }, day).column - layout.incomeFirst;
  const height = sheet.getLastRow() - OCR_BOOK_FIRST_ROW + 1;
  const rows = height > 0 ? sheet.getRange(OCR_BOOK_FIRST_ROW, 1, height, layout.expenseLast).getValues() : [];
  const notesOf = (column) => (height > 0 ? sheet.getRange(OCR_BOOK_FIRST_ROW, column, height, 1).getNotes().map(([note]) => note) : []);
  const [notesIn, notesOut] = [notesOf(layout.transferIn), notesOf(layout.transferOut)];
  const balances = rows
    .map((row, index) => ({ name: ocrCleanName(ocrText(row[0])), row, index }))
    .filter(({ name, row }) => name !== '' && [1, 2, 3].some((column) => typeof row[column] === 'number'))
    .map(({ name, row, index }) => ({ name, computed: ocrBalanceAsOf(row, { in: notesIn[index], out: notesOut[index] }, layout, offset, day) }));
  return { counted: ocrCountedLines(records), balances, isClosing: ocrIsPeriodEnd(place, day), tab: target.tab };
}

// [{ name, counted, computed }] of a count: each counted product with its balance that day.
function ocrCountItems(count) {
  const balances = new Map(count.balances.map((entry) => [ocrNameKey(entry.name), entry.computed]));
  return count.counted.map((entry) => ({ ...entry, computed: balances.get(ocrNameKey(entry.name)) || 0 }));
}
