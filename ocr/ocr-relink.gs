/**
 * Receipt OCR - "Бичсэн" and Тулгалт named again after a workbook's rows move (the owner, 2026-10-29:
 * after rows were put in above them - new products, Тосгон's "Жижиг бараа" divider - 133 of
 * October's "Бичсэн" cells and 121 of its Тулгалт cells named another product's row, and the site's
 * links opened that one). A month spreadsheet is reviewed when it changed itself, which a row put
 * into the workbook does not do. So every review pass keeps a fingerprint of the names of each
 * month's workbook tabs (OCR_RELINK_ROWS_PREFIX); when one changed, that month is queued
 * (OCR_RELINK_QUEUE_PROPERTY), as it is when a new product's rows are put in (ocrPutProductInNow).
 * reviewMovedOcrBooksNow then reviews the month's spreadsheets one after the other, a run a minute
 * after the other (ocrGoOn) until none is left: a review names each line's cells where its product's
 * row is now (ocrCellsNow) and makes Тулгалт from them again (ocrKeepChecks).
 */

const OCR_RELINK_ROWS_PREFIX = 'ocr-relink-rows:'; // + 'YYYY-MM' -> the fingerprint of that month's workbook tabs' names
const OCR_RELINK_QUEUE_PROPERTY = 'ocr-relink-queue'; // { months: ['YYYY-MM'], left: [spreadsheet ids of months[0] still to review] | null }
const OCR_RELINK_HANDLER = 'reviewMovedOcrBooksNow';
const OCR_RELINK_DAYS = ['01', '16']; // a day of each half-month

// A month's workbook tabs: [{ bookId, tab }] (Тосгон's month, each shop's two half-months).
function ocrMonthTabs(month) {
  const tabs = new Map();
  OCR_PLACES.forEach((place) => OCR_RELINK_DAYS.forEach((day) => {
    const target = ocrTarget({ place, side: 'in' }, `${month}-${day}`);
    const bookId = target ? ocrBookIdFor(place, target.tab) : null;
    if (bookId) {
      tabs.set(`${bookId}|${target.tab}`, { bookId, tab: target.tab });
    }
  }));
  return [...tabs.values()];
}

// The names of a month's workbook tabs, as a fingerprint: a row put in, moved or taken out changes
// it. opened: the workbooks opened in this run (bookId -> Spreadsheet).
function ocrMonthRowsPrint(month, opened = new Map()) {
  const names = ocrMonthTabs(month).map(({ bookId, tab }) => {
    if (!opened.has(bookId)) {
      opened.set(bookId, SpreadsheetApp.openById(bookId));
    }
    const sheet = opened.get(bookId).getSheetByName(tab);
    const height = sheet ? sheet.getLastRow() - OCR_BOOK_FIRST_ROW + 1 : 0;
    return [tab, height > 0 ? sheet.getRange(OCR_BOOK_FIRST_ROW, 1, height, 1).getValues().map(([name]) => ocrNameKey(name)) : []];
  });
  return ocrTextHash(JSON.stringify(names));
}

// The months that have month spreadsheets, by their names only (no folder walk: a pass looks every time).
function ocrBookMonths() {
  const months = new Set();
  ocrIterate(DriveApp.searchFiles(ocrSheetsQuery(null)), (file) => {
    const book = ocrMonthBookOf(file.getName());
    if (book) {
      months.add(book.month);
    }
  });
  return [...months].sort();
}

// Run by every review pass (runReviewPass): the months whose workbook rows changed since the last
// look are queued. -> those months
function ocrWatchWorkbookRows(props) {
  const opened = new Map();
  const changed = ocrBookMonths().filter((month) => {
    const key = OCR_RELINK_ROWS_PREFIX + month;
    const before = props.getProperty(key);
    let print;
    try {
      print = ocrMonthRowsPrint(month, opened);
    } catch (error) {
      Logger.log(`Ажлын номын мөр (${month}) уншигдсангүй: ${error.message}`);
      return false;
    }
    if (!DRY_RUN) {
      props.setProperty(key, print);
    }
    return before !== null && before !== print;
  });
  if (changed.length > 0) {
    Logger.log(`Ажлын номын мөр шилжсэн сар (${modeLabel()}): ${changed.join(', ')}: тэдгээрийн OCR файлуудыг дахин шалгана.`);
    ocrRelinkSoon(props, changed);
  }
  return changed;
}

function ocrRelinkState(props) {
  try {
    const state = JSON.parse(props.getProperty(OCR_RELINK_QUEUE_PROPERTY) || 'null');
    return state && Array.isArray(state.months) ? { months: state.months, left: Array.isArray(state.left) ? state.left : null } : { months: [], left: null };
  } catch (error) {
    Logger.log(`Дахин шалгах дараалал уншигдсангүй (${error.message}): хоосноос эхэлнэ`);
    return { months: [], left: null };
  }
}

// Months put in the queue, and its run set for a minute later. A month already begun starts over:
// the spreadsheets reviewed before the move name the rows before it.
function ocrRelinkSoon(props, months) {
  if (DRY_RUN) {
    return;
  }
  const state = ocrRelinkState(props);
  const isRestarted = months.includes(state.months[0]);
  props.setProperty(OCR_RELINK_QUEUE_PROPERTY, JSON.stringify({ months: [...new Set([...state.months, ...months])], left: isRestarted ? null : state.left }));
  ocrGoOn(OCR_RELINK_HANDLER, true);
}

// The queued months' spreadsheets reviewed, a run a minute after the other until none is left.
// Run by its trigger; from the editor too.
function reviewMovedOcrBooksNow() {
  return ocrRunGoingOn(OCR_RELINK_HANDLER, ocrReviewMovedBooks);
}

// -> { reviewed, left }: the spreadsheets reviewed, the months left for the next run.
function ocrReviewMovedBooks(deadline) {
  const props = PropertiesService.getScriptProperties();
  const context = ocrEditorContext(props);
  const workbooks = new Map();
  let books = null; // every month spreadsheet, listed once a run
  let state = ocrRelinkState(props);
  // kept after every step: a new product put in by a review queues its month again (ocrRelinkSoon)
  const keep = (next) => {
    state = next;
    if (!DRY_RUN) {
      props.setProperty(OCR_RELINK_QUEUE_PROPERTY, JSON.stringify(next));
    }
  };
  let reviewed = 0;
  while (state.months.length > 0 && Date.now() <= deadline) {
    const month = state.months[0];
    books = books || findOcrBooks(context.root, null);
    const left = state.left || books.filter((entry) => entry.month === month).map((entry) => entry.file.getId());
    if (left.length === 0) {
      if (!DRY_RUN) {
        props.setProperty(OCR_RELINK_ROWS_PREFIX + month, ocrMonthRowsPrint(month)); // as reviewed: the next pass does not queue it again
      }
      keep({ months: state.months.slice(1), left: null });
      continue;
    }
    keep({ ...state, left });
    const entry = books.find((book) => book.file.getId() === left[0]);
    const result = entry ? reviewOcrBook(entry, workbooks, context, deadline) : { isFinished: true }; // a spreadsheet gone: skipped
    if (!result.isFinished) {
      break;
    }
    reviewed += 1;
    const queued = DRY_RUN ? state : ocrRelinkState(props);
    keep(queued.left === null ? queued : { ...queued, left: left.slice(1) }); // started over during the review: from its first again
  }
  if (!DRY_RUN && state.months.length === 0) {
    props.deleteProperty(OCR_RELINK_QUEUE_PROPERTY);
  }
  Logger.log(`Мөр шилжсэний дараах шалгалт (${modeLabel()}): ${reviewed} файл${state.months.length > 0 ? `, үлдсэн сар: ${state.months.join(', ')} (минутын дараа үргэлжилнэ)` : ', бүгд дууслаа'}.`);
  return { reviewed, left: state.months.length };
}

// October's month spreadsheets named again now (the owner, 2026-10-29: the rows put into the
// workbooks that day): run from the editor; it goes on by itself until all are done.
function relinkOcrBooks202610() {
  ocrRelinkSoon(PropertiesService.getScriptProperties(), ['2026-10']);
  return reviewMovedOcrBooksNow();
}
