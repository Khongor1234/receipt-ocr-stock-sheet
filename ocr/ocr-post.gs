/**
 * Receipt OCR - the review pass, after the photos of a run were read. Every month spreadsheet that
 * changed (new rows, a mark of mom) is moved into the current layout and protected (ocr-sheet.gs),
 * the lines of its day tabs that may be written go into the workbook of their place
 * (ocrPostDecision, ocr-review.gs), and what is left to do in each day is kept for mom's report
 * (ocr-report.gs).
 *
 * A line is corrected in its day spreadsheet only (a right name or count, mom's Буруу), and the
 * correction is carried into the workbook. A person may still type into the workbook where a receipt
 * is missing (a transfer nobody photographed); the script never touches such a cell.
 * A workbook cell the script writes is blue on yellow (green once mom approved every line in it)
 * and its note lists the lines in it, one per line: the link to the line's row in its month
 * spreadsheet (its ID after "&ocr=", "<file id>.<line>"), its count and day ("✓": approved, "⚠": a
 * Шалгах line mom has not marked; the owner, 2026-10-26):
 *   "https://docs.google.com/spreadsheets/d/<book>/edit#gid=<tab>&range=A30:M30&ocr=<file id>.13 = 2 (2026-10-03) ✓"
 * and the cell is always the sum of them. So a line is never in a cell twice, a corrected count
 * replaces its own entry, a changed name takes the entry out of the old product's row and puts it
 * into the new one, and mom's Буруу takes it out. The row's "Бичсэн" says where it is now.
 * A cell whose number does not agree with its note (typed or changed by hand) is never touched:
 * the line goes into the same cell of the workbook's check copy instead (ocr-check.gs), and the
 * hand-typed cell turns yellow when the receipts make its number. A count is never added to
 * another count.
 */

const OCR_DAY_MS = 24 * 3600 * 1000;

// Where this pass starts: where a cut-off pass stopped; else the last pass (a little before, as
// Drive's search is late); else - first pass, mom's address or the columns changed - OCR_LOOKBACK_DAYS ago.
function ocrScanStart(props, applied, now) {
  const from = props.getProperty(OCR_SCAN_FROM_PROPERTY);
  if (from) {
    return new Date(from);
  }
  const last = props.getProperty(OCR_REVIEWERS_APPLIED_PROPERTY) === applied ? props.getProperty(OCR_LAST_SCAN_PROPERTY) : null;
  return last ? new Date(new Date(last).getTime() - OCR_SCAN_MARGIN_MS) : new Date(now.getTime() - OCR_LOOKBACK_DAYS * OCR_DAY_MS);
}

function runReviewPass(context, deadline) {
  const { props, root } = context;
  context.mom = ocrMomEmail(props);
  const startedAt = new Date();
  const lap = ocrStopwatch();
  ocrFollowCategoryMoves(props, deadline); // a new product's row mom moved: the later tabs too (ocr-new-product.gs)
  const movesSeconds = lap();
  // a new address or a new layout: every recent month spreadsheet is protected and moved again
  const applied = JSON.stringify({ mom: context.mom, header: OCR_HEADER, calendar: OCR_CALENDAR_VERSION });
  // the spreadsheets of the first version are moved into the month spreadsheets first
  findOldDaySheets(root).forEach((entry) => {
    if (DRY_RUN) {
      Logger.log(`${entry.path.join('/')}/${entry.file.getName()}: сарын spreadsheet руу шилжүүлнэ (туршилт)`);
      return;
    }
    const moved = ocrMigrateDaySheet(root, entry);
    context.touched.set(moved.file.getId(), moved);
  });
  const oldSeconds = lap();
  // the ones written in this run first, then the one the last pass stopped in (the books it went
  // through were changed by it, so would come before it again), then the others, longest unchanged first
  // and those of a photo mom deleted or moved since, which changes no spreadsheet (ocr-watch.gs)
  const changed = ocrReviewOrder(
    ocrWithLeftPhotos(findOcrBooks(root, ocrScanStart(props, applied, startedAt)), context, startedAt)
      .filter((entry) => !context.touched.has(entry.file.getId())),
    props.getProperty(OCR_REVIEW_CUT_PROPERTY)
  );
  // a spreadsheet first written during the pass (rows moved into it, ocrRefileDays) is added at the end
  const books = [...context.touched.values(), ...changed];
  Logger.log(`Шалгалтын бэлтгэл: шинэ барааны ангилал ${movesSeconds} сек, хуучин файл ${oldSeconds} сек, өөрчлөгдсөн файл хайх ${lap()} сек (${books.length} файл).`);
  const workbooks = new Map();
  const counts = { sheets: 0, written: 0, failed: 0, left: 0 };
  let reachedUpdate = null;
  for (let index = 0; index < books.length; index++) {
    if (Date.now() > deadline) {
      counts.left = books.length - index;
      break;
    }
    const { file } = books[index];
    const updatedBefore = context.touched.has(file.getId()) ? reachedUpdate : file.getLastUpdated(); // before this pass writes to it
    const result = reviewOcrBook(books[index], workbooks, context, deadline);
    Logger.log(`Шалгасан: ${books[index].place} ${books[index].kind.word} ${books[index].month} ${lap()} сек`);
    counts.sheets += 1;
    counts.written += result.written;
    counts.failed += result.failed;
    if (!result.isFinished) {
      counts.left = books.length - index; // this one again from its first day, the done days quickly
      if (!DRY_RUN) {
        props.setProperty(OCR_REVIEW_CUT_PROPERTY, file.getId());
      }
      break;
    }
    if (!DRY_RUN && props.getProperty(OCR_REVIEW_CUT_PROPERTY) === file.getId()) {
      props.deleteProperty(OCR_REVIEW_CUT_PROPERTY);
    }
    reachedUpdate = updatedBefore;
    [...context.touched.values()].filter((entry) => !books.some((known) => known.file.getId() === entry.file.getId())).forEach((entry) => books.push(entry));
  }
  if (!DRY_RUN) {
    ocrRememberScan(props, counts.left === 0, startedAt, reachedUpdate, applied);
  }
  lap();
  if (Date.now() <= deadline) {
    ocrWatchWorkbookRows(props); // rows put into a workbook, this pass's new products too (ocr-relink.gs)
  }
  const rowsSeconds = lap();
  const seconds = Math.round((Date.now() - startedAt.getTime()) / 1000);
  Logger.log(`Шалгалт (${modeLabel()}): хуудас ${counts.sheets}, workbook-д бичсэн мөр ${counts.written}, бичиж чадаагүй ${counts.failed}, дараагийн удаад ${counts.left} (${seconds} сек; ажлын номын мөр ${rowsSeconds} сек).`);
  return counts;
}

// Month spreadsheets to look at: the one cut off last time (cutId) first, then the longest unchanged.
const ocrReviewOrder = (entries, cutId) =>
  [...entries].sort((a, b) => (b.file.getId() === cutId) - (a.file.getId() === cutId) || a.file.getLastUpdated() - b.file.getLastUpdated());

// A finished pass: the next one starts from now. A cut-off one: from the last spreadsheet it reached.
function ocrRememberScan(props, isFinished, startedAt, reachedUpdate, applied) {
  if (!isFinished) {
    if (reachedUpdate) {
      props.setProperty(OCR_SCAN_FROM_PROPERTY, reachedUpdate.toISOString());
    }
    return;
  }
  props.deleteProperty(OCR_SCAN_FROM_PROPERTY);
  props.setProperty(OCR_LAST_SCAN_PROPERTY, startedAt.toISOString());
  props.setProperty(OCR_REVIEWERS_APPLIED_PROPERTY, applied);
}

// One month spreadsheet { file, place, month }: its tabs into the current layout, protect it,
// then every day tab until the deadline: { written, failed, isFinished }. One cut off keeps the
// calendar and checks it had.
function reviewOcrBook(entry, workbooks, context, deadline) {
  const lap = ocrStopwatch();
  const book = SpreadsheetApp.openById(entry.file.getId());
  const timeZone = book.getSpreadsheetTimeZone();
  // each step stops at the deadline (the owner, 2026-10-28: a 30-day spreadsheet's setup ran past
  // Google's 6 minutes and the run was killed); the next run goes on where this one stopped
  const isPrepared = ocrEachInTime(ocrDayTabs(book), ocrUpgradeTab, deadline) // only logs in DRY_RUN
    && (DRY_RUN || ocrSetUpBook(book, entry, context, deadline))
    && ocrRefileDays(book, entry, context, workbooks, timeZone, deadline);
  if (!isPrepared) {
    Logger.log(`${entry.place} ${entry.kind.word} ${entry.month}: бүтэц, тохиргоо хийж байгаад цаг дууслаа, дараагийн удаад үргэлжилнэ.`);
    return { written: 0, failed: 0, isFinished: false };
  }
  const tabs = ocrDayTabs(book);
  const setUpSeconds = lap();
  const results = [];
  for (const tab of tabs) {
    if (Date.now() > deadline) {
      break;
    }
    results.push(reviewOcrTab(tab, entry, { workbooks, context, timeZone, deadline }));
  }
  const tabsSeconds = lap();
  // a line left for a new product's rows (out of time) brings the spreadsheet back first next time
  const isFinished = results.length === tabs.length && results.every((result) => !result.deferred);
  if (isFinished) {
    ocrKeepCalendar(context.props, entry, results);
    ocrKeepChecks(book, entry, results);
    Logger.log(`  ${entry.place} ${entry.kind.word} ${entry.month}: бэлтгэл ${setUpSeconds} сек, ${results.length} өдөр ${tabsSeconds} сек, хадгалах ${lap()} сек`);
  } else {
    Logger.log(`${entry.place} ${entry.kind.word} ${entry.month}: ${results.length}/${tabs.length} өдөр шалгаад цаг дууслаа, дараагийн удаад үргэлжилнэ.`);
  }
  const total = results.reduce((sum, result) => ({ written: sum.written + result.written, failed: sum.failed + result.failed }), { written: 0, failed: 0 });
  return { ...total, isFinished };
}

// One month spreadsheet reviewed now, out of the pass's order (the owner, 2026-10-25: "OCR Тосгон
// тооллого 2026-10" waited behind 14 others for the new colours and columns): its tabs into the
// current layout, set up, and every day tab reviewed, as runReviewPass does. Run reviewTosgonCount202610
// from the editor; reviewOcrBookNow(id) for another one.
function reviewTosgonCount202610() {
  return reviewOcrBookNow(OCR_TOSGON_COUNT_1023.bookId);
}

// OCR Соронзон зарлага 2026-10 (the owner, 2026-10-29): its tabs 16, 17, 22 and 24 had "Бичсэн" in the
// old reason column (ocrRepairInstantWrites); run reviewSoronzonSales202610 to put them right and show
// its days on the site now, not after the next hourly pass.
const OCR_SORONZON_SALES_2026_10 = 'YOUR_SORONZON_SALES_BOOK_ID';
function reviewSoronzonSales202610() {
  return reviewOcrBookNow(OCR_SORONZON_SALES_2026_10);
}

// OCR My home орлого 2026-10 (the owner, 2026-10-29): its slips read again as Тосгон's ("Склд") and its
// new products put in, now rather than when the passes get to it after the setup of every spreadsheet.
const OCR_MY_HOME_INCOME_2026_10 = 'YOUR_MY_HOME_INCOME_BOOK_ID';
function reviewMyHomeIncome202610() {
  return reviewOcrBookNow(OCR_MY_HOME_INCOME_2026_10);
}

function reviewOcrBookNow(bookId) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(OCR_LOCK_WAIT_MS)) {
    Logger.log('OCR ажиллаж байна, хэсэг хүлээгээд дахин Run хийнэ үү.');
    return null;
  }
  try {
    return ocrReviewBookNow(bookId, Date.now() + OCR_TIME_BUDGET_MS);
  } finally {
    lock.releaseLock();
  }
}

function ocrReviewBookNow(bookId, deadline) {
  const file = DriveApp.getFileById(bookId);
  const month = ocrMonthBookOf(file.getName());
  if (!month) {
    throw new Error(`"${file.getName()}" сарын OCR файл биш`);
  }
  const result = reviewOcrBook({ file, ...month }, new Map(), ocrEditorContext(PropertiesService.getScriptProperties()), deadline);
  Logger.log(`${file.getName()} (${modeLabel()}): бичсэн ${result.written}, бичиж чадаагүй ${result.failed}${result.isFinished ? '' : ', цаг дууслаа: дахин Run хийнэ үү'}.`);
  return result;
}

// What a review run out of the hourly pass needs (reviewOcrBook's context).
function ocrEditorContext(props) {
  const today = ocrStamp(new Date(), 'yyyy-MM-dd');
  return {
    props, today, catalog: loadOcrCatalog(today), root: DriveApp.getFolderById(OCR_ROOT_FOLDER_ID), touched: new Map(), mom: ocrMomEmail(props),
  };
}

// Every month spreadsheet set up again now (the owner, 2026-10-26: the photos in N and the legend in
// O in all of them, not only in the ones a pass looks at). Run from the editor. One set up already
// is skipped (ocrSetUpBook), so a run cut off by the time goes on where it stopped when run again.
// Only the setup: no line is written into the workbooks.
function setUpAllOcrBooksNow() {
  return ocrRunGoingOn(OCR_SET_UP_ALL_HANDLER, ocrSetUpAllBooks); // 19 spreadsheets took an hour of Runs (2026-10-29)
}

// work(deadline) -> { left } under the script lock, run again a minute later while something is left.
function ocrRunGoingOn(handler, work) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(OCR_LOCK_WAIT_MS)) {
    Logger.log('OCR ажиллаж байна, минутын дараа өөрөө дахин оролдоно.');
    ocrGoOn(handler, !DRY_RUN); // the run a minute later found the hourly pass at work
    return null;
  }
  try {
    // set before the work, so a run Google stops at its 6 minutes still goes on (one took 365 s, 2026-10-29)
    ocrGoOn(handler, !DRY_RUN);
    const result = work(Date.now() + OCR_TIME_BUDGET_MS);
    ocrGoOn(handler, !DRY_RUN && result.left > 0);
    return result;
  } finally {
    lock.releaseLock();
  }
}

// A run of handler a minute later while one is wanted, none once not.
const OCR_SET_UP_ALL_HANDLER = 'setUpAllOcrBooksNow';
const OCR_GO_ON_MS = 60 * 1000;
function ocrGoOn(handler, isWanted) {
  ScriptApp.getProjectTriggers()
    .filter((trigger) => trigger.getHandlerFunction() === handler)
    .forEach((trigger) => ScriptApp.deleteTrigger(trigger));
  if (isWanted) {
    ScriptApp.newTrigger(handler).timeBased().after(OCR_GO_ON_MS).create();
  }
}

// Every day tab's "Бараа" and "Шалгах шалтгаан" columns fitted to their texts now, in every month
// spreadsheet (the owner, 2026-10-28: the tabs set up before the fit was part of the setup kept their
// old width). Run from the editor; a set-up spreadsheet is not set up again for this.
function fitOcrNameColumnsNow() {
  const books = findOcrBooks(DriveApp.getFolderById(OCR_ROOT_FOLDER_ID), null);
  if (DRY_RUN) {
    Logger.log(`"Бараа", "Шалгах шалтгаан" багана: ${books.length} файлын өдрийн табуудад текстийн уртаар тааруулна (туршилт).`);
    return books.length;
  }
  books.forEach((entry) => ocrDayTabs(SpreadsheetApp.openById(entry.file.getId())).forEach(ocrFitColumns));
  Logger.log(`"Бараа", "Шалгах шалтгаан" багана: ${books.length} файлын өдрийн табуудад текстийн уртаар тааруулав.`);
  return books.length;
}

// { done, left }: the month spreadsheets gone through, and those left for the next run.
function ocrSetUpAllBooks(deadline) {
  const props = PropertiesService.getScriptProperties();
  const today = ocrStamp(new Date(), 'yyyy-MM-dd');
  const context = { props, today, catalog: loadOcrCatalog(today), mom: ocrMomEmail(props) };
  // the oldest month first: the newest are set up last, so they keep their view triggers (OCR_VIEW_TRIGGERS_MAX)
  const books = findOcrBooks(DriveApp.getFolderById(OCR_ROOT_FOLDER_ID), null).sort((a, b) => a.month.localeCompare(b.month));
  for (let index = 0; index < books.length; index++) {
    if (Date.now() > deadline) {
      Logger.log(`Сарын OCR файл (${modeLabel()}): ${index}/${books.length} шинэ бүтэцтэй боллоо, цаг дууслаа: дахин Run хийнэ үү.`);
      return { done: index, left: books.length - index };
    }
    const book = SpreadsheetApp.openById(books[index].file.getId());
    const isDone = ocrEachInTime(ocrDayTabs(book), ocrUpgradeTab, deadline) // only logs in DRY_RUN
      && (DRY_RUN || ocrSetUpBook(book, books[index], context, deadline));
    if (!isDone) {
      Logger.log(`Сарын OCR файл (${modeLabel()}): ${index}/${books.length} шинэ бүтэцтэй боллоо, цаг дууслаа: дахин Run хийнэ үү.`);
      return { done: index, left: books.length - index };
    }
  }
  Logger.log(`Сарын OCR файл (${modeLabel()}): бүх ${books.length} файл ${DRY_RUN ? 'шинэ бүтэц рүү оруулна' : 'шинэ бүтэцтэй боллоо'} (зураг M, шалтгаан N, тайлбар O).`);
  return { done: books.length, left: 0 };
}

// IMPORTRANGE of the workbook into a month spreadsheet (ocrNameListFormula) needs the owner's "Allow
// access" once per spreadsheet: given here as the click would give it, once. When Google refuses, the
// list stays what the script wrote until the owner clicks "Allow access" on the hidden name tab.
function ocrAllowImport(bookId, mainId, props) {
  const key = `${OCR_IMPORT_ALLOWED_PREFIX}${bookId}`;
  if (props.getProperty(key) === mainId) {
    return;
  }
  const url = `https://docs.google.com/spreadsheets/d/${bookId}/externaldata/addimportrangepermissions?donorDocId=${mainId}`;
  const response = UrlFetchApp.fetch(url, { method: 'post', headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` }, muteHttpExceptions: true });
  if (response.getResponseCode() === 200) {
    props.setProperty(key, mainId);
    return;
  }
  Logger.log(`${bookId}: "Бараа"-г шууд уншихыг Google зөвшөөрсөнгүй (HTTP ${response.getResponseCode()}); _нэрс tab дээр "Allow access" дарна уу`);
}

// A short fingerprint of a text (FNV-1a, 32 bits), to tell whether it changed.
function ocrTextHash(text) {
  let hash = 0x811c9dc5;
  for (let index = 0; index < text.length; index++) {
    hash = Math.imul(hash ^ text.charCodeAt(index), 0x01000193) >>> 0;
  }
  return hash.toString(16);
}

// One day tab: write what may be written, keep what is left to do. Read by its titles, so a tab
// still in the old layout (DRY_RUN) is read right; it is written to only in the current layout.
function reviewOcrTab(sheet, entry, { workbooks, context, timeZone, deadline = Infinity }) {
  const day = sheet.getName();
  const where = `${entry.file.getName()} / ${day}`;
  const result = { written: 0, failed: 0, deferred: 0 };
  // the date of every line is the day of its tab (there is no date column since 2026-10-25)
  // _at: the line's row, which its cells' notes link to
  const dated = ocrShortenTabReasons(sheet, ocrReadRecords(sheet)).map((record, index) => ({
    ...record,
    'Огноо': `${entry.month}-${day}`,
    _at: { book: entry.file.getId(), gid: sheet.getSheetId(), row: ocrRowOf(record, index) },
  }));
  const refiled = ocrTosgonSentRows(sheet, where, ocrRefileMoved(sheet, where, dated, context, timeZone), context);
  const written = refiled.map((record, index) => {
    const outcome = ocrWriteRecord(record, workbooks, timeZone, deadline);
    if (outcome && outcome.deferred) {
      result.deferred += 1;
    }
    if (!outcome || outcome.deferred) {
      return record;
    }
    Logger.log(`  ${where} мөр ${ocrRowOf(record, index)}: ${outcome.updates['Бичсэн']}`);
    const updated = { ...record, ...outcome.updates };
    if (!DRY_RUN) {
      Object.keys(outcome.updates).forEach((title) => sheet.getRange(ocrRowOf(record, index), ocrColumn(title)).setValue(updated[title]));
    }
    if (!outcome.skipped) {
      result[outcome.ok ? 'written' : 'failed'] += 1;
      ocrAddDayStats(context, { written: outcome.ok && !outcome.corrected ? 1 : 0, corrected: outcome.corrected ? 1 : 0, failed: outcome.ok ? 0 : 1 });
    }
    return updated;
  });
  ocrSweepOrphans(written, workbooks, where); // a line left in a cell its "Бичсэн" does not name (ocr-orphans.gs)
  const records = ocrApplySuggestions(sheet, where, written, context.catalog.entries || []);
  // a count in this tab: listed in mom's half-month report, shown in the count report (ocr-count-report.gs)
  if (records.some((record) => ocrFolderMeaning(ocrText(record['Хавтас (ээж хийсэн)']).split(' / ')).kind === 'toollogo')) {
    ocrAddDayStats(context, { counts: [{ place: entry.place, day: `${entry.month}-${day}` }] });
  }
  const info = {
    id: `${entry.file.getId()}:${day}`,
    url: `${entry.file.getUrl()}#gid=${sheet.getSheetId()}`,
    name: `${OCR_DAY_SPREADSHEET_PREFIX}${entry.month}-${day}`,
    place: entry.place,
    label: `${entry.place} / ${entry.kind.folder}`,
  };
  ocrKeepSummary(context.props, ocrSheetSummary(info, records));
  return { ...result, day, gid: sheet.getSheetId(), calendar: ocrCalendarDay(records), checks: ocrDayChecks(records, entry, day, workbooks) };
}

// A tab's reasons in a few words (ocrShortenReasons): the long ones written before rewritten, and the
// column fitted to them again. The records as they then stand.
function ocrShortenTabReasons(sheet, records) {
  const shortened = records.map((record) => {
    const text = ocrShortenReasons(record[OCR_REASON_COLUMN]);
    return text === ocrText(record[OCR_REASON_COLUMN]) ? record : { ...record, [OCR_REASON_COLUMN]: text };
  });
  if (DRY_RUN || shortened.every((record, index) => record === records[index])) {
    return shortened;
  }
  ocrRewriteColumn(sheet, OCR_REASON_COLUMN, records, shortened);
  ocrFitReasonColumn(sheet);
  return shortened;
}

// A photo mom moved after it was read (the owner, 2026-10-28: a loose photo read as the day it came in
// or the date on it, then put into its day's folder; a slip put into another month's or kind's folder):
// its rows go to the day tab of the spreadsheet its folder names now, its photo and receipt prices with
// them, the folder on the rows that of now. That tab's review then moves the workbook cells: "Бичсэн"
// names the old day's, so they are taken out and the new day's written. A photo of another place now,
// or outside a month folder is left; one outside a day folder in the same month stays where it is
// (nothing says its day), in another month it goes under the day it came in, noted (ocrPhotoMoves). A
// photo mom put in the trash after it was read (the owner, 2026-10-28: Соронзон/зарлага/2026-10's
// IMG_5866, read loose as the 26th, then deleted; its rows and counts stayed and the site showed the
// 26th) has its rows and photo taken out of the tab and its lines out of the workbook cells
// (ocrForgetPhoto). Each photo's folder is looked at once a run (context.located).
function ocrRefileDays(book, entry, context, workbooks, timeZone, deadline = Infinity) {
  const located = context.located || (context.located = new Map());
  const spotOf = (id) => {
    if (!located.has(id)) {
      located.set(id, ocrPhotoSpot(id, context.root));
    }
    return located.get(id);
  };
  return ocrEachInTime(ocrDayTabs(book), (tab) => ocrRefileTab(tab, book, entry, { context, workbooks, timeZone, spotOf }), deadline);
}

// One day tab: the rows of its moved photos into their tabs of now, those of its deleted photos out.
function ocrRefileTab(tab, book, entry, { context, workbooks, timeZone, spotOf }) {
  const records = ocrReadRecords(tab);
  const moves = ocrPhotoMoves(records, entry, tab.getName(), spotOf);
  const gone = ocrGonePhotos(records, spotOf);
  if (moves.length === 0 && gone.length === 0) {
    return;
  }
  const here = `${entry.place} ${entry.kind.word} ${entry.month} / ${tab.getName()}`;
  moves.forEach((move) => Logger.log(`  ${here}: ${move.photoId} зургийг ${move.spot.label} / ${move.spot.month} / ${move.day}${move.spot.day ? '' : ' (өдрийн хавтасгүй)'} руу зөөсөн тул мөрүүд нь тийшээ ${DRY_RUN ? 'шилжинэ' : 'шилжлээ'}`));
  gone.forEach((photoId) => Logger.log(`  ${here}: ${photoId} зургийг устгасан тул мөрүүд нь ${DRY_RUN ? 'хасагдана' : 'хасагдлаа'}`));
  if (DRY_RUN) {
    return;
  }
  const prices = ocrHasPrices(tab) ? records.map((record) => tab.getRange(record._row, OCR_READ_PRICE_COLUMN).getValue()) : records.map(() => '');
  moves.forEach((move) => {
    const target = ocrRefileTarget(book, entry, move, context);
    const mine = records.map((record, index) => [record, index]).filter(([record]) => ocrPhotoIdOf(record) === move.photoId);
    const first = ocrAppendRows(target, mine.map(([record]) => ocrRow(ocrRefiledRecord(record, move))), mine.map(([, index]) => prices[index]));
    ocrAddPhoto(target, first, mine.length, DriveApp.getFileById(move.photoId));
    context.props.setProperty(OCR_DONE_PREFIX + move.photoId, `${OCR_SHEETS_FOLDER}/${move.spot.month}/${move.spot.kind.folder}/${ocrMonthBookName(move.spot.place, move.spot.kind, move.spot.month)} / ${move.day}`);
  });
  gone.forEach((photoId) => ocrForgetPhoto(photoId, records.filter((record) => ocrPhotoIdOf(record) === photoId), `${entry.month}-${tab.getName()}`, here, context.props, workbooks, timeZone));
  const left = new Set([...moves.map((move) => move.photoId), ...gone]);
  const kept = records.map((record, index) => [record, index]).filter(([record]) => !left.has(ocrPhotoIdOf(record)));
  ocrRewriteTab(tab, kept.map(([record]) => ocrRow(record)), kept.map(([, index]) => prices[index]));
}

// The photos of a tab that are in the trash now, one id each.
const ocrGonePhotos = (records, spotOf) =>
  [...new Set(records.map(ocrPhotoIdOf).filter(Boolean))].filter((photoId) => (spotOf(photoId) || {}).gone);

// A deleted photo's lines leave the workbook cells their "Бичсэн" names, as a line mom marked Буруу
// does (ocrWriteRecord; a cell changed by hand keeps its number, logged); the photo is forgotten, and
// its slip too unless these rows only noted it as a copy of another photo's, so the slip put back or
// sent again is read as new. iso: the tab's day. rows: the photo's records.
function ocrForgetPhoto(photoId, rows, iso, here, props, workbooks, timeZone) {
  rows.forEach((record) => {
    const outcome = ocrWriteRecord({ ...record, 'Ээж: Зөв/Буруу': OCR_VERDICT_WRONG, 'Огноо': iso }, workbooks, timeZone);
    if (outcome && !outcome.ok) {
      Logger.log(`  ${here}: ${record['ID']} мөрийн тоог буцааж чадсангүй (${ocrFailedReason(outcome.updates['Бичсэн'])})`);
    }
  });
  props.deleteProperty(OCR_DONE_PREFIX + photoId);
  if (rows.some((record) => !ocrText(record['Шалгах шалтгаан']).startsWith(OCR_DUPLICATE_REASON))) {
    ocrForgetSlip(photoId, props);
  }
}

// The slip of a photo in the trash (still readable there) is no longer known as filed.
function ocrForgetSlip(photoId, props) {
  try {
    props.deleteProperty(OCR_SEEN_PREFIX + ocrMd5(DriveApp.getFileById(photoId).getBlob()));
  } catch (error) {
    Logger.log(`  ${photoId}: устгасан зургийг уншиж чадсангүй, дахин оруулбал давхардсан гэж тэмдэглэгдэнэ (${error.message})`);
  }
}

// The day tab the moved rows go to: of this spreadsheet, or of the one the photo's kind and month name
// now (made when there is none); that one is looked at by the review too (context.touched).
function ocrRefileTarget(book, entry, move, context) {
  if (move.isSameBook) {
    return ocrDayTab(book, move.day);
  }
  const { book: other, file } = ocrMonthBook(context.root, move.spot.place, move.spot.kind, move.spot.month);
  if (context.touched && !context.touched.has(file.getId())) {
    context.touched.set(file.getId(), { file, place: move.spot.place, kind: move.spot.kind, month: move.spot.month });
  }
  return ocrDayTab(other, move.day);
}

// Where a photo lies now: { place, kind, month, day ('DD' or null outside a day folder), label, uploadedOn };
// { gone: true } in the trash; null when it is gone for good (not told from a Drive error: left), outside
// Баримт/<place>/<folder>/<month> or in a folder of mom's own. paths: as ocrLocate's.
function ocrPhotoSpot(photoId, root, paths = null) {
  try {
    const file = DriveApp.getFileById(photoId);
    if (file.isTrashed()) {
      return { gone: true };
    }
    const place = ocrLocate(file, root.getId(), paths);
    const month = place && ocrFolderMonth(place.path);
    const folder = place && ocrFolderMeaning(place.path);
    if (!month || folder.moves.length === 0) {
      return null;
    }
    const date = ocrFolderDate(place.path);
    return { place: place.path[0], kind: ocrBookKind(place.path), month, day: date ? date.slice(8) : null, label: folder.label, uploadedOn: ocrStamp(file.getDateCreated(), 'yyyy-MM-dd') };
  } catch (error) {
    return null;
  }
}

// The photos of a tab that lie elsewhere now: [{ photoId, spot, day, note, isSameBook }], one per photo.
// spotOf(photoId): ocrPhotoSpot. A photo outside a day folder in this spreadsheet's month stays; in
// another month it goes under the day it came in (that month's 1st when it came in another month), noted.
function ocrPhotoMoves(records, entry, tabDay, spotOf) {
  const photoIds = [...new Set(records.map(ocrPhotoIdOf).filter(Boolean))];
  return photoIds.flatMap((photoId) => {
    const spot = spotOf(photoId);
    if (!spot || spot.gone || spot.place !== entry.place) {
      return [];
    }
    const isSameBook = spot.kind.word === entry.kind.word && spot.month === entry.month;
    if (isSameBook && (!spot.day || spot.day === tabDay)) {
      return [];
    }
    const isUploadedThen = spot.uploadedOn.startsWith(spot.month);
    const day = spot.day || (isUploadedThen ? spot.uploadedOn.slice(8) : '01');
    const note = spot.day ? '' : `${OCR_NO_DAY_REASON}, ${isUploadedThen ? `оруулсан өдрөөр (${spot.uploadedOn})` : `${spot.month}-${day}-ээр (өдөр тодорхойгүй)`}`;
    return [{ photoId, spot, day, note, isSameBook }];
  });
}

// A moved row as it goes into its new tab: the folder of now, the "өдрийн хавтсанд хийгээгүй" note gone
// (or the new one, outside a day folder still) and, with another folder, the "хавтас буруу" note too;
// nothing left to check: OK.
function ocrRefiledRecord(record, move) {
  // a slip Тосгон or the city sent keeps its "(Тосгоноос)" or "(Хотоос)" while its photo stays in the shop's income folder
  const recordLabel = ocrText(record['Хавтас (ээж хийсэн)']);
  const label = OCR_SLIP_MARKS.some((mark) => recordLabel === `${move.spot.label}${mark}`) ? recordLabel : move.spot.label;
  const isOtherFolder = label !== recordLabel;
  const kept = ocrText(record['Шалгах шалтгаан']).split('; ')
    .filter((reason) => reason !== '' && !reason.startsWith(OCR_NO_DAY_REASON) && !(isOtherFolder && reason.startsWith(OCR_MISFILED_REASON)));
  const reasons = [...(move.note ? [move.note] : []), ...kept];
  return { ...record, 'Хавтас (ээж хийсэн)': label, 'Шалгах шалтгаан': reasons.join('; '), 'Төлөв': reasons.length === 0 ? OCR_STATUS_OK : OCR_STATUS_REVIEW };
}

// A photo noted as maybe in the wrong folder (ocrMisfiledReason) and since moved by mom: its rows take
// the new folder, so ocrWriteRecord writes the sides that were missing (a line already in a cell has
// its own note entry there and is not counted twice). The note goes; the rest of the reasons stay.
function ocrRefileMoved(sheet, where, records, context, timeZone) {
  const located = new Map();
  return records.map((record, index) => {
    if (!ocrText(record['Шалгах шалтгаан']).includes(OCR_MISFILED_REASON)) {
      return record;
    }
    const fileId = ocrText(record['ID']).split('#')[0];
    if (!located.has(fileId)) {
      located.set(fileId, ocrLocateById(fileId, context.root));
    }
    const place = located.get(fileId);
    const folder = place ? ocrFolderMeaning(place.path) : null;
    if (!folder || folder.moves.length === 0 || folder.label === ocrText(record['Хавтас (ээж хийсэн)'])) {
      return record;
    }
    const reasons = ocrText(record['Шалгах шалтгаан']).split('; ').filter((reason) => reason !== '' && !reason.startsWith(OCR_MISFILED_REASON));
    const updates = {
      'Хавтас (ээж хийсэн)': folder.label,
      'Шалгах шалтгаан': reasons.join('; '),
      ...(reasons.length === 0 ? { 'Төлөв': OCR_STATUS_OK } : {}),
    };
    Logger.log(`  ${where} мөр ${ocrRowOf(record, index)}: зургийг "${folder.label}" хавтас руу зөөсөн тул тэр хавтсаар бичнэ`);
    if (!DRY_RUN) {
      Object.keys(updates).forEach((title) => sheet.getRange(ocrRowOf(record, index), ocrColumn(title)).setValue(updates[title]));
    }
    return { ...record, ...updates };
  });
}

// Where a photo lies now ({ folder, path }, ocrLocate), or null when it is gone or not under Баримт.
function ocrLocateById(fileId, root) {
  try {
    return ocrLocate(DriveApp.getFileById(fileId), root.getId());
  } catch (error) {
    return null;
  }
}

// 'YYYY-MM-DD' from the Огноо cell; '' when there is none. Sheets turns the text into a date at
// midnight of the spreadsheet's time zone, so it is read back in that zone (not OCR_TIME_ZONE: a
// sheet made in Tokyo time would otherwise give the day before).
function ocrIsoDate(value, timeZone) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, timeZone, 'yyyy-MM-dd');
  }
  return ocrIsRealDate(ocrText(value)) ? ocrText(value) : '';
}

// The moves of a line into the workbooks there are (OCR_POST_BOOKS), from the folder it came from;
// a check count (not on the last day of its sheet) is written only from the 10/31 sheets on, into
// its day's block (ocr-spot.gs).
// A shop's sale mom marked Бэлэг goes into the gift column instead (ocrTargetInYear).
const ocrPostMoves = (record, iso) =>
  ocrFolderMeaning(ocrText(record['Хавтас (ээж хийсэн)']).split(' / ')).moves
    .filter((move) => OCR_POST_BOOKS[move.place] && !(iso && ocrIsSpotCount(move, iso) && !ocrIsSpotBlockCount(move, iso)))
    .map((move) => (ocrIsGift(record) && ocrIsShopSale(move) ? { ...move, gift: true } : move));

// A move into a shop's sheet out of the OCR's reach (OCR_HANDS_OFF_PERIOD).
const ocrIsHandsOffMove = (move, iso) => {
  const target = iso ? ocrTarget(move, iso) : null;
  return Boolean(target) && ocrIsHandsOffTab(target.tab);
};

// "Бичсэн" without the check prefix and the hand-typed status of its cells, to tell whether a
// line changed without reading the cells.
const ocrWrittenCore = (text) =>
  ocrText(text).split('\n').map((line) => line.replace(/ \([^()]*\)$/, '').replace(OCR_CHECK_PREFIX, '')).join('\n');

// null when nothing is to be done now; { deferred: true } when the line needs a new product's rows
// and the run is past its deadline; else { ok, corrected, skipped?, updates: { title: value } }.
function ocrWriteRecord(record, workbooks, timeZone, deadline = Infinity) {
  const state = ocrWrittenState(record);
  // a line that failed only for a missing row is tried at every review: the row may be there now
  if ((state === 'failed' && !ocrIsNoRowFailure(record['Бичсэн'])) || state === 'other year') {
    return null;
  }
  const decision = ocrPostDecision(record);
  // the cells "Бичсэн" names; one of a sheet out of reach (OCR_HANDS_OFF_PERIOD) is never touched again
  const named = ['compare', 'hands off'].includes(state) ? [] : ocrCellsNow(record['Бичсэн'], workbooks);
  const previous = named.filter((cell) => !ocrIsHandsOffCell(cell));
  const isWrong = ocrText(record['Ээж: Зөв/Буруу']) === OCR_VERDICT_WRONG;
  const isTakenOut = isWrong || ocrIsNotRelevant(record);
  const iso = ocrIsoDate(record['Огноо'], timeZone);
  if (decision.post && iso !== '' && !ocrHasBooksFor(Number(iso.slice(0, 4))) && previous.length === 0) {
    return { ok: true, corrected: false, skipped: true, updates: { 'Бичсэн': OCR_OTHER_YEAR_TEXT } };
  }
  const allMoves = decision.post ? ocrPostMoves(record, iso) : [];
  const moves = allMoves.filter((move) => !ocrIsHandsOffMove(move, iso));
  // every cell of the line in a sheet out of reach: "Бичсэн" says so, nothing is written or taken out
  if (moves.length === 0 && previous.length === 0 && (allMoves.length > 0 || named.length > 0)) {
    return ocrText(record['Бичсэн']) === OCR_HANDS_OFF_TEXT ? null : { ok: true, corrected: false, skipped: true, updates: { 'Бичсэн': OCR_HANDS_OFF_TEXT } };
  }
  // waiting for a name, or nothing to write: what is in the workbook stays as it is
  if (moves.length === 0 && !(isTakenOut && previous.length > 0)) {
    return null;
  }
  const fail = (reason) => ocrFailure(record, reason);
  const problem = decision.post ? ocrValueProblem(decision) || (iso === '' ? 'огноо алга' : '') : '';
  if (problem) {
    return fail(problem);
  }
  const found = moves.map((move) => ocrFindOrPutIn(move, iso, decision, workbooks, record, deadline)); // a new product put in (ocr-new-product.gs)
  if (found.some((target) => target.deferred)) {
    return { deferred: true };
  }
  if ([...found.filter((target) => target.ok).map((target) => target.cell), ...previous].some((cell) => OCR_HANDS_OFF_CELLS.includes(cell))) {
    return null;
  }
  // nothing changed since it was written (a hand-typed cell: nor mom's number there)
  const lines = ocrWrittenLines(record['Бичсэн']);
  if (found.every((target) => target.ok) && state === 'done' &&
    ocrWrittenCore(record['Бичсэн']) === found.map((target, index) => ocrWrittenLine(ocrLineCell(lines[index], target, workbooks), decision)).join('\n') &&
    ocrHandLinesCurrent(record, found, workbooks, decision) && ocrLinksCurrent(record, found, workbooks)) {
    return null;
  }
  const wanted = found
    .map((target) => (target.ok && ocrIsHandTarget(target) ? ocrCheckTarget(target, workbooks) : target))
    .map((target) => ocrWithElsewhere(target, decision, record['ID'], workbooks));
  const dropped = previous.filter((cell) => !wanted.some((target) => target.cell === cell)).map((cell) => ocrFindWrittenCell(cell, iso, workbooks));
  const targets = [...wanted.map((target) => ({ target, qty: target.elsewhere ? 0 : decision.qty })), ...dropped.map((target) => ({ target, qty: null }))];
  const changeOf = ({ target, qty }) => ({
    target,
    change: ocrCellChange(target.range.getValue(), target.range.getNote(), record, qty, target.isCount, iso, Boolean(decision.approved), Boolean(target.isCheck), target.elsewhere || null, Boolean(decision.check)),
  });
  const lost = targets.filter(({ target }) => !target.ok).map(({ target }) => target.reason);
  if (lost.length > 0) {
    // the cells of a product the line is no longer still give it up, though its new one has no row
    // yet (the owner, 2026-10-28: Sharp 102l stayed counted, pink, after mom chose Sharp 100l)
    const failure = ocrFailure(record, lost.join('; '), ocrTakeOut(dropped.map((target) => ({ target, qty: null })), changeOf));
    return failure.updates['Бичсэн'] === ocrText(record['Бичсэн']) ? null : failure; // still no row: nothing new
  }
  const changes = targets.map(changeOf);
  const refused = changes.filter(({ change }) => change.action === 'refuse').map(({ target, change }) => `${target.cell}: ${change.reason}`);
  if (refused.length > 0) {
    return fail(refused.join('; '));
  }
  const after = ocrApplyCellChanges(changes);
  const text = decision.post
    ? wanted.map((target) => ocrWrittenLine(target.cell, decision, target.isCheck ? ocrHandStatus(target, after.get(target.cell)) : ocrElsewhereStatus(target.elsewhere))).join('\n')
    : isWrong ? OCR_REMOVED_TEXT : OCR_NOT_RELEVANT_REMOVED_TEXT;
  if (text === ocrText(record['Бичсэн'])) {
    return null;
  }
  // only mom's approval or the hand-typed status changed: no new line, no correction
  const plainText = (value) => ocrWrittenCore(value).replace(new RegExp(` (${OCR_APPROVED_MARK}|${OCR_MUST_CHECK_MARK})`, 'g'), '');
  const isSameLine = previous.length > 0 && plainText(text) === plainText(record['Бичсэн']);
  return { ok: true, corrected: previous.length > 0 && !isSameLine, skipped: isSameLine, updates: { 'Бичсэн': text } };
}

// The cells "Бичсэн" names, each where its product's row is now: a row put in above it (a new product,
// ocr-new-product.gs; the year's sync; by hand) moved it down, and the cell named is another product's
// (the owner, 2026-10-29). A check cell, or a product its tab does not have, stays as named.
// The cell a line of "Бичсэн" names for a found target (ocrWrittenCore's form): a hand-typed cell's line
// names the check copy's row of its product, which is not the original's once rows were put into the
// workbook after the copy was made (2026-10-29: V269 for V261, so 26 of a review's 77 lines were
// written again every hour, nothing changed). Read from the tabs read this run (ocrTabGrid).
function ocrLineCell(line, target, workbooks) {
  if (!String(line || '').startsWith(OCR_CHECK_PREFIX)) {
    return target.cell;
  }
  const bookId = ocrBookIdFor(target.place, target.tabName);
  const copyId = ocrCheckBookId(bookId, workbooks);
  const original = ocrTabGrid(workbooks, bookId, target.tabName);
  const copy = copyId ? ocrWorkbookTab(workbooks, copyId, target.tabName) : null;
  const name = original && original.values[target.row - OCR_BOOK_FIRST_ROW] ? ocrText(original.values[target.row - OCR_BOOK_FIRST_ROW][0]) : '';
  const row = copy && !copy.error && name !== '' ? copy.rows.get(ocrNameKey(name)) : null;
  return row ? `${target.tabName}!${columnLetter(target.column)}${row}` : target.cell;
}

function ocrCellsNow(written, workbooks) {
  return ocrWrittenLines(written).map((line) => {
    const cell = line.match(OCR_WRITTEN_LINE)[1];
    const parts = cell.match(/^(.+)!([A-Z]+)(\d+)$/);
    const place = parts && !cell.startsWith(OCR_CHECK_PREFIX) ? ocrPlaceOfTab(parts[1]) : null;
    if (!place || !OCR_POST_BOOKS[place]) {
      return cell;
    }
    const [, tabName, letters, row] = parts;
    const tab = ocrWorkbookTab(workbooks, ocrBookIdFor(place, tabName), tabName);
    const rowNow = tab.error ? null : ocrWrittenNames(line).map((name) => tab.rows.get(ocrNameKey(name))).find(Boolean);
    return rowNow && rowNow !== Number(row) ? `${tabName}!${letters}${rowNow}` : cell;
  });
}

// The product of a "Бичсэн" line ("…!F12 ← 2 · Haier 300l ✓ (гараар 1)"): the name whole first (it may
// end in "(2)"), then without the status in brackets.
function ocrWrittenNames(line) {
  const rest = line.replace(/^\S+ ← \S+ · /, '');
  return [...new Set([ocrWithoutMark(rest), ocrWithoutMark(rest.replace(/ \([^()]*\)$/, ''))])];
}

// A failed line's reason names a product with no row on its sheet (OCR_NO_ROW_REASON, ocr-reconcile.gs).
const ocrIsNoRowFailure = (written) =>
  ocrFailedReason(written).split('; ').some((reason) => OCR_NO_ROW_REASON.test(reason.replace(/^check: /, '')));

// The changes made: written cells as they are now, the hand-typed originals of check cells marked;
// a Map of cell -> { value, entries } after them.
function ocrApplyCellChanges(changes) {
  changes.filter(({ change }) => change.action === 'write').forEach(({ target, change }) => ocrApplyChange(target.range, change));
  const after = new Map(changes.map(({ target, change }) => [target.cell, change.action === 'write' ? change : ocrCellAfter(target)]));
  changes.filter(({ target }) => target.isCheck).forEach(({ target }) => ocrMarkHandCell(target, after.get(target.cell)));
  return after;
}

// A failed line out of the cells it no longer belongs in ({ target, qty: null }; changeOf gives the
// change): the cells it left, or none when one of them can not be found or refuses.
function ocrTakeOut(leaving, changeOf) {
  if (leaving.some(({ target }) => !target.ok)) {
    return [];
  }
  const changes = leaving.map(changeOf);
  if (changes.some(({ change }) => change.action === 'refuse')) {
    return [];
  }
  ocrApplyCellChanges(changes);
  return changes.map(({ target }) => target.cell);
}

// The line could not be written; it says why and waits for a change of its checks. Nothing was
// changed in any workbook but the cells it left (takenOut, ocrTakeOut), no longer named in "Бичсэн".
function ocrFailure(record, reason, takenOut = []) {
  const earlier = ocrText(record['Шалгах шалтгаан']);
  return {
    ok: false,
    corrected: false,
    updates: {
      'Бичсэн': ocrFailedText(reason, record, takenOut),
      'Төлөв': OCR_STATUS_REVIEW,
      'Шалгах шалтгаан': ocrShortenReasons([earlier, `${OCR_FAILED_PREFIX}${reason}`].filter(Boolean).join('; ')), // in a few words, once
    },
  };
}

// The lines of a cell's note: [{ id, qty, day, approved, check, elsewhere, at }], or null when it
// holds anything else. elsewhere ({ day, cell }): mom typed the line's goods on another day of the row,
// so the line holds 0 here (ocrHandElsewhere). The day matters in a transfer column (T, AJ), which has no day columns: a check count
// compares with the transfers up to its own day only. at ({ book, gid, row }): the line's row in its
// month spreadsheet, which the line links to (since 2026-10-26). The notes written before read too:
// the photo's link ("Зураг: https://drive.google.com/file/d/<file id>/view#13 = 2 (2026-10-03)",
// 2026-10-25) and "OCR: IMG.HEIC [<ID>] = 2 @2026-10-03 ✓" with its bare link line under it.
const OCR_NOTE_TAIL = '( ✓| ⚠)?(?: \\(ээж (\\d+)-нд гараар: ([A-Z]+\\d+)\\))?$';
const OCR_SHEET_LINK_PREFIX = 'https://docs.google.com/spreadsheets/d/';
const OCR_SHEET_NOTE_LINE = new RegExp(`^https://docs\\.google\\.com/spreadsheets/d/([^/\\s]+)/edit#gid=(\\d+)&range=A(\\d+):[A-Z]+\\d+&ocr=([^\\s.]+)(?:\\.(\\d+))? = (-?[\\d.]+)(?: \\((\\d{4}-\\d{2}-\\d{2})\\))?${OCR_NOTE_TAIL}`);
const OCR_NOTE_LINE = new RegExp(`^Зураг: https://drive\\.google\\.com/file/d/([^/\\s#]+)/view(#\\d+)? = (-?[\\d.]+)(?: \\((\\d{4}-\\d{2}-\\d{2})\\))?${OCR_NOTE_TAIL}`);
const OCR_OLD_NOTE_LINE = new RegExp(`^OCR: .* \\[([^\\]]+)\\]() = (-?[\\d.]+)(?: @(\\d{4}-\\d{2}-\\d{2}))?${OCR_NOTE_TAIL}`);
function ocrNoteEntries(note) {
  const lines = ocrText(note).split('\n').filter((line) => line.trim() !== '' && !ocrIsPhotoLinkLine(line));
  const entries = lines.map(ocrNoteEntry);
  return entries.every(Boolean) ? entries : null;
}

// One line of a note, of any of its kinds; null when it is not the script's.
function ocrNoteEntry(line) {
  const linked = line.match(OCR_SHEET_NOTE_LINE);
  // the parts of a line of the photo's kind, a line of the sheet's kind put in their order
  const match = linked
    ? [line, linked[4], linked[5] ? `#${linked[5]}` : '', ...linked.slice(6)]
    : line.match(OCR_NOTE_LINE) || line.match(OCR_OLD_NOTE_LINE);
  return match ? {
    id: match[1] + (match[2] || ''), qty: Number(match[3]), day: match[4] || '', approved: match[5] === ` ${OCR_APPROVED_MARK}`,
    check: match[5] === ` ${OCR_MUST_CHECK_MARK}`,
    elsewhere: match[7] ? { day: Number(match[6]), cell: match[7] } : null,
    at: linked ? { book: linked[1], gid: Number(linked[2]), row: Number(linked[3]) } : null,
  } : null;
}

// The link to a line's row: "abc#13" in tab 0, row 30 ->
// "https://docs.google.com/spreadsheets/d/<book>/edit#gid=0&range=A30:M30&ocr=abc.13"; Sheets opens
// the tab with the row chosen, and "&ocr=" keeps the line's ID (a row moves; the ID stays).
const ocrRowLink = (at, id) => {
  const [fileId, line] = String(id).split('#');
  return `${OCR_SHEET_LINK_PREFIX}${at.book}/edit#gid=${at.gid}&range=A${at.row}:${columnLetter(OCR_HEADER.length)}${at.row}&ocr=${fileId}${line ? `.${line}` : ''}`;
};

// "abc#13" -> "https://docs.google.com/spreadsheets/d/<book>/edit#gid=0&range=A30:M30&ocr=abc.13 = 2 (2026-10-03) ✓";
// a line whose row is not known yet keeps the photo's link ("Зураг: https://drive.google.com/file/d/abc/view#13 = 2 ...").
const ocrNoteLine = (entry) => {
  const [fileId, line] = String(entry.id).split('#');
  const link = entry.at ? ocrRowLink(entry.at, entry.id) : `${OCR_PHOTO_LINK_PREFIX}${fileId}/view${line ? `#${line}` : ''}`;
  return `${link} = ${entry.qty}${entry.day ? ` (${entry.day})` : ''}${ocrLineMark(entry)}` +
    (entry.elsewhere ? ` (ээж ${entry.elsewhere.day}-нд гараар: ${entry.elsewhere.cell})` : '');
};

// Whether the notes of a line's cells link to its row as it is now: a row moves (ocrSpreadPhotos, a
// reread) and a note of 2026-10-25 links to the photo; read from each tab once per run (ocrTabGrid).
function ocrLinksCurrent(record, found, workbooks) {
  if (!record._at) {
    return true;
  }
  const link = `${ocrRowLink(record._at, record['ID'])} = `;
  return ocrWrittenLines(record['Бичсэн']).every((line, index) => {
    const target = found[index];
    const isCheck = line.startsWith(OCR_CHECK_PREFIX);
    const bookId = ocrBookIdFor(target.place, target.tabName);
    const gridId = isCheck ? ocrCheckBookId(bookId, workbooks) : bookId;
    const grid = gridId ? ocrTabGrid(workbooks, gridId, target.tabName) : null;
    const row = isCheck ? Number(ocrParseWrittenLine(line).cell.replace(/^[A-Z]+/, '')) : target.row;
    const note = grid ? (grid.notes[row - OCR_BOOK_FIRST_ROW] || [])[target.column - 1] : undefined;
    return note === undefined || ocrText(note).split('\n').some((noteLine) => noteLine.startsWith(link));
  });
}
const ocrNoteSum = (entries) => entries.reduce((total, entry) => total + entry.qty, 0);
// The background of a cell holding these lines, its worst line's (the owner, 2026-10-25): green when
// mom approved them all, pink while one is a Шалгах line she has not marked, else yellow (the AI is sure).
function ocrCellColor(entries) {
  if (entries.length === 0) {
    return null;
  }
  if (entries.every((entry) => entry.approved)) {
    return OCR_APPROVED_CELL_COLOR;
  }
  return entries.some((entry) => entry.check) ? OCR_CHECK_COLOR : OCR_WRITTEN_COLOR;
}

// A cell typed or changed by hand: a number, or a note, that is not the script's lines.
function ocrIsHandCell(current, note) {
  const entries = ocrNoteEntries(note);
  const isEmpty = current === '' || current === null;
  return entries === null || (!isEmpty && entries.length === 0) || (isEmpty ? 0 : current) !== ocrNoteSum(entries);
}

// What to do with a workbook cell so the line holds qty in it (null: the line is taken out):
// { action: 'write', value, note, color, entries } | { action: 'skip' } | { action: 'refuse', reason }.
// A cell of the check copy (isCheck) holds the script's lines only: what was copied into it from
// the hand-typed original is replaced. elsewhere: mom's cell of the goods on another day (qty is 0).
// check: the line is a Шалгах line mom has not marked (ocrPostDecision); a line written before the
// pink (2026-10-25) is rewritten with its mark, so its cell gets the colour at the next review.
function ocrCellChange(current, note, record, qty, isCount, day = '', approved = false, isCheck = false, elsewhere = null, check = false) {
  const read = ocrNoteEntries(note);
  const entries = read === null && isCheck ? [] : read;
  const isEmpty = current === '' || current === null;
  if (!isCheck && (entries === null || (!isEmpty && entries.length === 0))) {
    return { action: 'refuse', reason: `нүдэнд гараар бичсэн ${current === '' ? 'тэмдэглэл' : current} байна` };
  }
  if (!isCheck && (isEmpty ? 0 : current) !== ocrNoteSum(entries)) {
    return { action: 'refuse', reason: `нүдийг гараар өөрчилсөн (${current === '' ? 'хоосон' : current}, тэмдэглэлээр ${ocrNoteSum(entries)})` };
  }
  const id = record['ID'];
  const others = entries.filter((entry) => entry.id !== id);
  const mine = entries.find((entry) => entry.id === id);
  const at = record._at || (mine && mine.at) || null;
  const isSame = isCheck ? current === (entries.length > 0 ? ocrNoteSum(entries) : '') : true;
  if (isSame && (qty === null ? !mine : mine && mine.qty === qty && mine.day === day && mine.approved === approved && mine.check === (check && !approved) &&
    JSON.stringify(mine.elsewhere) === JSON.stringify(elsewhere) && JSON.stringify(mine.at) === JSON.stringify(at))) {
    return { action: 'skip' };
  }
  if (isCount && qty !== null && others.length > 0) {
    return { action: 'refuse', reason: `тооллого өөр баримтаас бичигдсэн (${current})` };
  }
  const next = qty === null ? others : [...others, { id, qty, day, approved, check: check && !approved, elsewhere, at }];
  return {
    action: 'write',
    value: next.length > 0 ? ocrNoteSum(next) : '',
    note: next.map(ocrNoteLine).join('\n'),
    color: ocrCellColor(next),
    entries: next,
  };
}

function ocrApplyChange(range, change) {
  if (DRY_RUN) {
    return;
  }
  range.setValue(change.value).setNote(change.note).setFontColor(change.value === '' ? null : OCR_POSTED_FONT_COLOR);
  ocrSetCellColor(range, change.color);
}

// A cell's OCR colour; null takes away only a colour the OCR gave (the sheet's own stays).
function ocrSetCellColor(range, color) {
  if (color) {
    range.setBackground(color);
  } else if (OCR_CELL_COLORS.includes(String(range.getBackground()).toLowerCase())) {
    range.setBackground(null);
  }
}

// A workbook cell: { ok, cell: 'Tos/2026/10!AJ57', range, isCount, isDay, place, tabName, row,
// column } | { ok: false, reason }. A cell of a count block (ocr-spot.gs) is neither: the lines of
// a day's pages add up there.
function ocrCellAt(place, tabName, row, column, layout, iso, workbooks) {
  const tab = ocrWorkbookTab(workbooks, ocrBookIdFor(place, tabName), tabName);
  if (tab.error) {
    return { ok: false, reason: tab.error };
  }
  const isCount = column === layout.count;
  const isDay = column < layout.summaryFirst && ![layout.transferIn, layout.transferOut, layout.gift, layout.count].includes(column);
  return { ok: true, cell: `${tabName}!${columnLetter(column)}${row}`, range: tab.sheet.getRange(row, column), isCount, isDay, place, tabName, row, column };
}

// The cell of a line: the product's row, the column of its move and day.
function ocrFindCell(move, iso, name, workbooks) {
  const target = ocrTarget(move, iso);
  if (!target) {
    return { ok: false, reason: `${iso} workbook-ийн хугацаанд ороогүй` };
  }
  const tab = ocrWorkbookTab(workbooks, ocrBookIdFor(move.place, target.tab), target.tab);
  if (tab.error) {
    return { ok: false, reason: tab.error };
  }
  const row = tab.rows.get(ocrNameKey(name));
  if (!row) {
    return { ok: false, reason: `"${name}" ${target.tab} дээр алга` };
  }
  const column = ocrIsSpotBlockCount(move, iso) ? ocrSpotBlock(tab.sheet, ocrBookIdFor(move.place, target.tab), target.layout, iso) : target.column;
  return ocrCellAt(move.place, target.tab, row, column, target.layout, iso, workbooks);
}

// The place of a workbook tab ("Sor/2026/10/15" -> Соронзон, "Tos/2026/10" -> Тосгон); undefined if none.
function ocrPlaceOfTab(tabName) {
  const prefix = tabName.split('/')[0];
  const chain = CHAINS.find((candidate) => candidate.tabPrefix === prefix);
  return prefix === OCR_TOSGON.tabPrefix ? OCR_TOSGON.place : chain && chain.label;
}

// A cell named in "Бичсэн" ("Sor/2026/10/15!E4", "check:Sor/2026/10/15!U22"), to take a line out of it.
function ocrFindWrittenCell(cell, iso, workbooks) {
  if (cell.startsWith(OCR_CHECK_PREFIX)) {
    return ocrFindCheckCell(cell, iso, workbooks);
  }
  const [, tabName, letters, row] = cell.match(/^(.+)!([A-Z]+)(\d+)$/);
  const place = ocrPlaceOfTab(tabName);
  if (!place || !OCR_POST_BOOKS[place]) {
    return { ok: false, reason: `${cell}: аль workbook болох нь танигдсангүй` };
  }
  const layout = place === OCR_TOSGON.place ? OCR_TOSGON.layout : LAYOUT;
  return ocrCellAt(place, tabName, Number(row), ocrColumnNumber(letters), layout, iso, workbooks);
}

// "AJ" -> 36
const ocrColumnNumber = (letters) => letters.split('').reduce((total, letter) => total * 26 + letter.charCodeAt(0) - 64, 0);

// A tab of a workbook with its product rows (name -> row), read once a run.
function ocrWorkbookTab(workbooks, bookId, tabName) {
  const key = `${bookId}|${tabName}`;
  if (!workbooks.has(key)) {
    workbooks.set(key, ocrReadTab(bookId, tabName));
  }
  return workbooks.get(key);
}

function ocrReadTab(bookId, tabName) {
  let sheet;
  try {
    sheet = SpreadsheetApp.openById(bookId).getSheetByName(tabName);
  } catch (error) {
    return { error: `workbook нээгдсэнгүй (${error.message})` };
  }
  if (!sheet) {
    return { error: `"${tabName}" sheet алга` };
  }
  const count = sheet.getLastRow() - OCR_BOOK_FIRST_ROW + 1;
  const names = count > 0 ? sheet.getRange(OCR_BOOK_FIRST_ROW, 1, count, 1).getValues() : [];
  const rows = new Map();
  names.forEach(([name], index) => {
    const key = ocrNameKey(name);
    if (key !== '' && !rows.has(key)) {
      rows.set(key, OCR_BOOK_FIRST_ROW + index);
    }
  });
  return { sheet, rows };
}
