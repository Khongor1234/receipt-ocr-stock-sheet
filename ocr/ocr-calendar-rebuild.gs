/**
 * Receipt OCR - counts the price-watch site's calendar ("Баримт" tab) again for older months.
 *
 * The review pass only looks at month spreadsheets changed in the last OCR_LOOKBACK_DAYS, so a
 * month older than that (2026-04, say) never reaches the calendar and the site shows 0 photos.
 * rebuildReceiptCalendar reads such a month's spreadsheets (read only: nothing is written to a
 * sheet, the workbook or the OCR state) and keeps their numbers where ocrFeedAnswer finds them.
 *
 * Run from the Apps Script editor: rebuildReceiptCalendarApril (or rebuildReceiptCalendarAll).
 */

const OCR_REBUILD_MONTHS = ['2026-01', '2026-02', '2026-03', '2026-04', '2026-05', '2026-06', '2026-07', '2026-08', '2026-09'];

// The month spreadsheets of `months` ('YYYY-MM'), whatever their age: [{ file, place, month }].
function findOcrBooksOfMonths(root, months) {
  return findOcrBooks(root, null).filter((entry) => months.includes(entry.month));
}

function rebuildReceiptCalendar(months) {
  const props = PropertiesService.getScriptProperties();
  const root = DriveApp.getFolderById(OCR_ROOT_FOLDER_ID);
  const books = findOcrBooksOfMonths(root, months);
  books.forEach((entry) => {
    const book = SpreadsheetApp.openById(entry.file.getId());
    const results = ocrDayTabs(book).map((tab) => ({
      day: tab.getName(),
      gid: tab.getSheetId(),
      calendar: ocrCalendarDay(ocrReadRecords(tab)),
    }));
    ocrKeepCalendar(props, entry, results);
    Logger.log(`${entry.place} ${entry.kind.word} ${entry.month}: ${results.filter((result) => Object.keys(result.calendar).length > 0).length} өдөр`);
  });
  Logger.log(`Календарь дахин тоолсон (${modeLabel()}): ${books.length} сарын spreadsheet, сар: ${months.join(', ')}.`);
  return books.length;
}

function rebuildReceiptCalendarApril() {
  return rebuildReceiptCalendar(['2026-04']);
}

function rebuildReceiptCalendarAll() {
  return rebuildReceiptCalendar(OCR_REBUILD_MONTHS);
}
