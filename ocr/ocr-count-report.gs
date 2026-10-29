/**
 * Receipt OCR - the count report: ONE spreadsheet, Баримт/Тайлан/Тооллогын тайлан. Mom (or the
 * owner) chooses a day (B1) and a place (D1); below comes that day's count (ocr-count-read.gs), a
 * row per counted product: counted, in the books, the difference and its likely reason, the
 * short ones first and red, the over ones orange. It is written again at once when B1 or D1
 * changes (onCountReportEdit, installed by installOcrTrigger) and on every hourly run, so a late
 * receipt of that day or before shows; nothing of a later day ever does.
 * The likely reason, most telling first:
 *   1. another place counted the same product the other way that day (My home 1 short, Соронзон 1
 *      over): a transfer between them was not written, or written for the wrong place
 *   2. lines of that product wait in a day tab (not checked, or not written)
 *   3. none of these: the product is really short (sold without a receipt, written on another day
 *      or place, or counted wrong)
 * A closing count (the last day of its sheet) also lists the products nobody counted.
 */

const OCR_COUNT_REPORT_NAME = 'Тооллогын тайлан';
const OCR_COUNT_TABLE_ROW = 6; // its header; rows below
const OCR_COUNT_HEADER = ['Бараа', 'Тоолсон', 'Бүртгэлээр', 'Зөрүү', 'Байж болох шалтгаан'];
const OCR_COUNT_EDIT_HANDLER = 'onCountReportEdit';
const OCR_SHORT_COLOR = '#f4cccc';
const OCR_OVER_COLOR = '#fce5cd';
const OCR_FROM_NAMES = { 'Соронзон': 'Соронзоноос', 'Тосгон': 'Тосгоноос', 'My home': 'My home-оос' };
const OCR_AT_NAMES = { 'Соронзон': 'Соронзонд', 'Тосгон': 'Тосгонд', 'My home': 'My home-д' };
const ocrFrom = (place) => OCR_FROM_NAMES[place] || `${place}-аас`;
const ocrAt = (place) => OCR_AT_NAMES[place] || `${place}-д`;

// ---------- pure ----------

// How many lines of each product wait in a day tab for that place: Map name key -> count.
function ocrWaitingByProduct(summaries, place) {
  const waiting = new Map();
  summaries.flatMap((summary) => summary.waiting || [])
    .filter((line) => line.places.includes(place))
    .forEach((line) => waiting.set(ocrNameKey(line.product), (waiting.get(ocrNameKey(line.product)) || 0) + 1));
  return waiting;
}

// The other places' differences: Map name key -> [{ place, diff }] (diff = counted - computed).
function ocrDiffsElsewhere(others) {
  const diffs = new Map();
  others.forEach(({ place, items }) =>
    items
      .filter((entry) => entry.counted !== entry.computed)
      .forEach((entry) => diffs.set(ocrNameKey(entry.name), [...(diffs.get(ocrNameKey(entry.name)) || []), { place, diff: entry.counted - entry.computed }]))
  );
  return diffs;
}

// The reason when another place counted the same product the other way; '' when none did.
function ocrTransferReason(place, isShort, elsewhere) {
  const other = (elsewhere || []).find((entry) => (isShort ? entry.diff > 0 : entry.diff < 0));
  if (!other) {
    return '';
  }
  const [from, to] = isShort ? [place, other.place] : [other.place, place];
  return `${ocrAt(other.place)} ${Math.abs(other.diff)} ширхэг ${isShort ? 'илүү' : 'дутсан'} байна: ${ocrFrom(from)} ${to} руу шилжүүлэг бүртгэгдээгүй эсвэл буруу газарт бичигдсэн байж магадгүй, шилжүүлгийн баримтыг шалгана уу (баримт байхгүй бол workbook-д гараар нэмнэ).`;
}

function ocrCountReason(waitingCount, isShort) {
  if (waitingCount > 0) {
    return `Энэ барааны ${waitingCount} мөр баримт шалгалт хүлээж workbook-д ороогүй байна: эхлээд тэдгээрийг шалгана уу, зөрүү арилж магадгүй.`;
  }
  return isShort
    ? 'Workbook-д ороогүй баримт олдсонгүй, тиймээс бараа үнэхээр дутсан байна. Баримтгүй зарагдсан, өөр өдөр эсвэл газарт бичигдсэн, эсвэл буруу тоолсон байж болно: дахин тоолж шалгана уу.'
    : 'Орлого эсвэл буцаалт бүртгэгдээгүй, эсвэл зарлага давхар бичигдсэн байж болно.';
}

// A count as text and table: { summary, conclusion, rows: [[name, counted, computed, diff,
// reason]] (short first, then over, then the ones that agree), uncounted: [[name, computed]] }.
function ocrCountTable(count, waiting, elsewhere) {
  const items = ocrCountItems(count);
  const short = items.filter((entry) => entry.counted < entry.computed);
  const over = items.filter((entry) => entry.counted > entry.computed);
  const pieces = (list) => list.reduce((total, entry) => total + Math.abs(entry.counted - entry.computed), 0);
  const reason = (entry) => {
    if (entry.counted === entry.computed) {
      return 'Таарсан';
    }
    const isShort = entry.counted < entry.computed;
    return ocrTransferReason(count.place, isShort, elsewhere.get(ocrNameKey(entry.name))) || ocrCountReason(waiting.get(ocrNameKey(entry.name)) || 0, isShort);
  };
  const gap = (entry) => entry.counted - entry.computed;
  const rows = [...short.sort((a, b) => gap(a) - gap(b)), ...over.sort((a, b) => gap(b) - gap(a)), ...items.filter((entry) => gap(entry) === 0)]
    .map((entry) => [entry.name, entry.counted, entry.computed, gap(entry), reason(entry)]);
  const countedKeys = new Set(items.map((entry) => ocrNameKey(entry.name)));
  const uncounted = count.isClosing
    ? count.balances.filter((entry) => entry.computed > 0 && !countedKeys.has(ocrNameKey(entry.name))).map((entry) => [entry.name, entry.computed])
    : [];
  let conclusion = items.length === 0 ? 'Дүгнэлт: энэ өдөр тооллого ороогүй байна.' : 'Дүгнэлт: бүх тоо бүртгэлтэй таарч байна, дутсан бараа алга.';
  if (short.length > 0) {
    conclusion = `Дүгнэлт: ${items.length} бараанаас ${short.length} нь дутсан байна (нийт ${pieces(short)} ширхэг).`;
  } else if (over.length > 0) {
    conclusion = `Дүгнэлт: дутсан бараа алга, ${over.length} бараа илүү гарсан.`;
  }
  const kind = count.isClosing ? `хаалтын тооллого, ${count.tab}` : 'шалгах тооллого, workbook-д бичээгүй';
  const summary = `${count.place}, ${count.day} (${kind}): тоолсон бараа ${items.length}, таарсан ${items.length - short.length - over.length}, дутсан ${short.length} (${pieces(short)} ширхэг), илүү ${over.length} (${pieces(over)} ширхэг).`;
  return { summary, conclusion, rows, uncounted };
}

// ---------- the spreadsheet ----------

// Баримт/Тайлан/Тооллогын тайлан, made the first time: the day (B1) and the place (D1) to choose,
// the rest only the owner's; mom's own account may choose too. null in DRY_RUN until it is made.
function ocrCountReportBook(root, mom) {
  const folder = DRY_RUN ? ocrChildFolder(root, OCR_REPORT_FOLDER) : ocrEnsurePath(root, [OCR_REPORT_FOLDER]);
  if (!folder) {
    return null;
  }
  const files = folder.getFilesByName(OCR_COUNT_REPORT_NAME);
  if (files.hasNext()) {
    const found = files.next();
    if (!DRY_RUN) { // an account of mom's added later may choose too
      const editors = found.getEditors().map((user) => user.getEmail());
      ocrMomList(mom).filter((email) => !editors.includes(email)).forEach((email) => found.addEditor(email));
    }
    return SpreadsheetApp.openById(found.getId());
  }
  if (DRY_RUN) {
    return null;
  }
  const book = SpreadsheetApp.create(OCR_COUNT_REPORT_NAME);
  const file = DriveApp.getFileById(book.getId());
  file.moveTo(folder);
  const sheet = book.getSheets()[0];
  sheet.getRange(1, 1, 1, 4).setValues([['Өдөр:', '', 'Газар:', OCR_PLACES[1]]]).setFontWeight('bold');
  sheet.getRange(1, 2).setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build());
  sheet.getRange(1, 4).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(OCR_PLACES, true).setAllowInvalid(false).build());
  const diff = columnLetter(OCR_COUNT_HEADER.indexOf('Зөрүү') + 1);
  const first = OCR_COUNT_TABLE_ROW + 1;
  const body = sheet.getRange(first, 1, sheet.getMaxRows() - OCR_COUNT_TABLE_ROW, OCR_COUNT_HEADER.length);
  const rule = (formula, color) => SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(formula).setBackground(color).setRanges([body]).build();
  sheet.setConditionalFormatRules([
    rule(`=AND(ISNUMBER($${diff}${first}),$${diff}${first}<0)`, OCR_SHORT_COLOR),
    rule(`=AND(ISNUMBER($${diff}${first}),$${diff}${first}>0)`, OCR_OVER_COLOR),
  ]);
  sheet.setFrozenRows(OCR_COUNT_TABLE_ROW);
  const protection = sheet.protect().setDescription(`${OCR_GUARD_PREFIX}: тооллогын тайлан`);
  protection.setUnprotectedRanges([sheet.getRange(1, 2), sheet.getRange(1, 4)]);
  ocrOnlyOwner(protection, Session.getEffectiveUser().getEmail());
  ocrMomList(mom).forEach((email) => file.addEditor(email));
  return book;
}

// Writes the chosen day's count of the chosen place (today and Соронзон when nothing is chosen).
function refreshCountReport(root, props, today) {
  const book = ocrCountReportBook(root, ocrMomEmail(props));
  const sheet = book ? book.getSheets()[0] : null;
  const chosen = (sheet && ocrIsoDate(sheet.getRange(1, 2).getValue(), book.getSpreadsheetTimeZone())) || today;
  const typed = sheet ? ocrText(sheet.getRange(1, 4).getValue()) : '';
  const place = OCR_PLACES.includes(typed) ? typed : OCR_PLACES[1];
  const count = ocrReadCountAt(root, place, chosen);
  const others = OCR_PLACES.filter((other) => other !== place && OCR_POST_BOOKS[other])
    .map((other) => ({ place: other, read: ocrReadCountAt(root, other, chosen) }))
    .filter(({ read }) => read && read.counted.length > 0)
    .map(({ place: other, read }) => ({ place: other, items: ocrCountItems(read) }));
  const table = count
    ? ocrCountTable({ ...count, place, day: chosen }, ocrWaitingByProduct(ocrOpenSummaries(props), place), ocrDiffsElsewhere(others))
    : { summary: `${place}, ${chosen}: workbook-д энэ өдрийн sheet алга.`, conclusion: '', rows: [], uncounted: [] };
  if (DRY_RUN || !sheet) {
    Logger.log(`Тооллогын тайлан (туршилт, бичээгүй): ${table.summary} ${table.conclusion}`);
    return { ...table, url: '' };
  }
  ocrWriteCountTable(sheet, table);
  return { ...table, url: DriveApp.getFileById(book.getId()).getUrl() };
}

function ocrWriteCountTable(sheet, table) {
  const width = OCR_COUNT_HEADER.length;
  const blank = (text) => [text, ...Array(width - 1).fill('')];
  sheet.getRange(2, 1, sheet.getMaxRows() - 1, width).clearContent();
  sheet.getRange(2, 1, 2, width).setValues([blank(table.summary), blank(table.conclusion)]);
  const uncounted = table.uncounted.length > 0
    ? [blank(''), blank(`Тоолоогүй бараа (бүртгэлээр үлдэгдэлтэй ${table.uncounted.length})`), ...table.uncounted.map(([name, computed]) => [name, '', computed, '', 'тоологдоогүй'])]
    : [];
  const rows = [OCR_COUNT_HEADER, ...table.rows, ...uncounted];
  sheet.getRange(OCR_COUNT_TABLE_ROW, 1, rows.length, width).setValues(rows);
  sheet.getRange(OCR_COUNT_TABLE_ROW, 1, 1, width).setFontWeight('bold');
}

// Installable onEdit of the count report: a new day or place is shown at once.
function onCountReportEdit(event) {
  const range = event && event.range;
  if (!range || range.getRow() !== 1 || ![2, 4].includes(range.getColumn())) {
    return;
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(OCR_LOCK_WAIT_MS * 6)) {
    return;
  }
  try {
    refreshCountReport(DriveApp.getFolderById(OCR_ROOT_FOLDER_ID), PropertiesService.getScriptProperties(), ocrStamp(new Date(), 'yyyy-MM-dd'));
  } finally {
    lock.releaseLock();
  }
}

// The edit trigger of the count report (the report is made when it is not there yet).
function installCountReportTrigger() {
  const book = ocrCountReportBook(DriveApp.getFolderById(OCR_ROOT_FOLDER_ID), ocrMomEmail(PropertiesService.getScriptProperties()));
  ScriptApp.getProjectTriggers()
    .filter((trigger) => trigger.getHandlerFunction() === OCR_COUNT_EDIT_HANDLER)
    .forEach((trigger) => ScriptApp.deleteTrigger(trigger));
  ScriptApp.newTrigger(OCR_COUNT_EDIT_HANDLER).forSpreadsheet(book).onEdit().create();
}
