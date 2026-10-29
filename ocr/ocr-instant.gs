/**
 * Receipt OCR - mom's mark written at once (the owner, 2026-10-26: "Зөв" on a count line goes into
 * the workbook right away, not at the next hourly run). The month spreadsheet's On edit trigger
 * (onOcrViewEdit, ocr-sheet.gs) calls it when "Ээж: Зөв/Буруу" of a day tab changes: that one tab is
 * reviewed as the hourly pass does (reviewOcrTab), so Буруу takes a line out at once too. While the
 * pass runs the edit is left to it: the spreadsheet changed, so the next pass looks at it.
 */

const OCR_INSTANT_LOCK_WAIT_MS = 20000;

// An edit of mom's mark on a line (a paste over several columns too), not of the header.
function ocrIsMarkEdit(range) {
  const mark = ocrColumn('Ээж: Зөв/Буруу');
  return range.getRow() >= 2 && range.getColumn() <= mark && mark <= range.getLastColumn();
}

// { written, failed, ... } of the tab, or null when it was left to the hourly pass.
function ocrWriteTabNow(sheet) {
  const book = sheet.getParent();
  const month = ocrMonthBookOf(book.getName());
  if (!month) {
    return null;
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(OCR_INSTANT_LOCK_WAIT_MS)) {
    Logger.log(`${book.getName()} / ${sheet.getName()}: OCR ажиллаж байна, ээжийн тэмдэг дараагийн ажлаар бичигдэнэ.`);
    return null;
  }
  try {
    const props = PropertiesService.getScriptProperties();
    const today = ocrStamp(new Date(), 'yyyy-MM-dd');
    const context = {
      props, today, catalog: loadOcrCatalog(today), root: DriveApp.getFolderById(OCR_ROOT_FOLDER_ID), touched: new Map(), mom: ocrMomEmail(props),
    };
    const entry = { file: DriveApp.getFileById(book.getId()), ...month };
    ocrUpgradeTabNow(sheet, context);
    const result = reviewOcrTab(sheet, entry, { workbooks: new Map(), context, timeZone: book.getSpreadsheetTimeZone() });
    Logger.log(`${book.getName()} / ${sheet.getName()} (${modeLabel()}): ээжийн тэмдгээр бичсэн ${result.written}, бичиж чадаагүй ${result.failed}.`);
    ocrKeepCalendarDay(props, entry, result);
    return result;
  } finally {
    lock.releaseLock();
  }
}

// A tab still in a layout before (the hourly pass moves every tab, but a pass cut by the time leaves
// some) is moved before it is written: written by the new columns it was not (2026-10-28..29: mom's
// Зөв put "Бичсэн" into the old reason column). The move took its protection away: made again.
function ocrUpgradeTabNow(sheet, context) {
  if (ocrUpgradeTab(sheet)) {
    ocrGuardTab(sheet, ocrGuardDescription(context.mom), context.mom, Session.getEffectiveUser().getEmail());
  }
}

// The site's numbers of the one tab just written, kept at once (the owner, 2026-10-29: the day turns
// green right after mom's Зөв, not at the next hourly pass, which keeps every tab's: ocrKeepCalendar),
// and a new stamp in the site's cache key (ocrFeedCachedAnswer), so the next visit reads them.
function ocrKeepCalendarDay(props, entry, result) {
  if (DRY_RUN) {
    return;
  }
  const key = ocrCalendarKey(entry.place, entry.month, entry.kind);
  const { [result.day]: before, ...others } = JSON.parse(props.getProperty(key) || '{}');
  const kept = Object.keys(result.calendar).length > 0 ? { ...others, [result.day]: result.calendar } : others;
  props.setProperty(key, JSON.stringify(kept));
  props.setProperty(OCR_CALENDAR_EDIT_PROPERTY, new Date().toISOString());
}

// "Бүх мөр Зөв" in A1 (OCR_MARK_DAY_OK): A1 shows the view it showed (view) again first, so the tab
// never keeps the choice; then every line mom has not looked at (ocrReviewState: no mark of hers, no
// product or count she changed) gets Зөв, the doubtful (pink) ones too: she checked the whole day in
// the workbook. Her Буруу and Бэлэг stay. The tab is then written as a single Зөв writes it.
function ocrMarkDayOk(sheet, view) {
  ocrViewCell(sheet).setValue(view);
  ocrApplyView(sheet, view);
  const rows = ocrReadRecords(sheet).filter((record) => ocrIsProductLine(record) && ocrReviewState(record).needsMom).map((record) => record._row);
  Logger.log(`${sheet.getParent().getName()} / ${sheet.getName()}: ${OCR_MARK_DAY_OK}: ${rows.length} мөр${DRY_RUN ? ' (туршилт, тэмдэглээгүй)' : ''}`);
  if (DRY_RUN || rows.length === 0) {
    return;
  }
  const mark = ocrColumn('Ээж: Зөв/Буруу');
  ocrRuns(rows).forEach(([first, count]) => sheet.getRange(first, mark, count, 1).setValues(Array.from({ length: count }, () => [OCR_VERDICT_OK])));
  ocrFitColumns(sheet);
  ocrWriteTabNow(sheet);
}
