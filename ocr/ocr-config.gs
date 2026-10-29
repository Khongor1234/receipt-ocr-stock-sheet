/**
 * Receipt OCR - settings.
 *
 * Mom puts a photo of a receipt in the folder of the place that GIVES the goods:
 *   Баримт/Соронзон/{орлого, зарлага, Соронзоноос My home руу, Соронзоноос Тосгон руу (буцаалт)}
 *   Баримт/My home/{орлого, зарлага, My home-оос Соронзон руу, My home-оос Тосгон руу (буцаалт)}
 *   Баримт/Тосгон/{орлого, Тосгоноос Соронзон руу, Тосгоноос My home руу}
 * ("орлого" is the one exception: goods a supplier brings into that place.) Each place also has a
 * "тооллого" folder for photos of its physical count; those lines go to the count column.
 * Inside each of them she opens the month, then the day, and puts the photo there; the script makes
 * these folders ahead (this month and the next):
 *   Соронзон/зарлага/2026-10/01/IMG_5603.HEIC     mom's photo; it is never moved
 * processNewReceipts (ocr-main.gs) reads it once with Gemini and writes one row per product line
 * into the tab "01" of Баримт/Зураг уншуулах/2026-10/Зарлага/OCR Соронзон зарлага 2026-10 (one spreadsheet per
 * place, kind and month), the photo beside its rows.
 * The day is always the folder mom chose (a photo put outside a day folder: the day it was uploaded);
 * the date written on the receipt is not used. A photo outside a day folder, a blank or notebook
 * page, a duplicate: their rows are marked "Шалгах" with the reason.
 * Then mom (only her) marks the lines Зөв/Буруу, with the right name or count when wrong
 * (ocr-sheet.gs). A line goes into the workbook of its place (ocr-post.gs): an OK line at once, a
 * "Шалгах" line once mom said Зөв. Mom gets a report on the 15th and the month's last day (ocr-report.gs). The inventory workbook is
 * only read for names.
 * Photos are never moved (Drive lets only a file's owner move it, and mom owns her uploads), so the
 * script runs in the owner's account (shared by mom and the owner); which photos were read is kept
 * in Script properties.
 *
 * The bundle is profile.gs + config.gs (tab names, column layout) + ocr/*.gs: sh build.sh ocr.
 * Setup in a NEW standalone Apps Script project (script.google.com > New project), in the shared account:
 *   1. paste neg-file/ocr-all-in-one.gs
 *   2. Project Settings > Script properties: GEMINI_KEY = the Gemini API key (never in this file)
 *   3. run processNewReceipts with DRY_RUN = true and read the log: which folders would be made and
 *      which photos would be read (Gemini is NOT called, nothing changes)
 *   4. DRY_RUN = false, run processNewReceipts once, then installOcrTrigger once (every hour)
 *   5. Script properties: OCR_REVIEWERS (mom's Gmail, see below)
 */

const OCR_ROOT_FOLDER_ID = 'YOUR_RECEIPTS_ROOT_FOLDER_ID'; // Баримт
// What was read goes into ONE spreadsheet per place, kind and month, a tab per day, in a folder of
// its kind (2026-10-25: from one spreadsheet of every kind nobody could tell income from expense):
// Баримт/Зураг уншуулах/2026-10/Орлого/OCR Соронзон орлого 2026-10, tabs "01" ... "31" (only the days with photos).
const OCR_DAY_SPREADSHEET_PREFIX = 'OCR ';
// The folder of the month spreadsheets, in Mongolian (the owner, 2026-10-25). Until the folder in
// Drive is renamed it is found by its old name; a new one is made by the new name only (ocrSheetsFolder).
const OCR_SHEETS_FOLDER = 'Зураг уншуулах';
const OCR_OLD_SHEETS_FOLDER = 'OCR';
// The kinds, by the kind of mom's folder (ocrFolderMeaning): the folder under the month and the word
// in the file name. A return is a transfer too.
const OCR_BOOK_KINDS = {
  orlogo: { folder: 'Орлого', word: 'орлого' },
  zarlaga: { folder: 'Зарлага', word: 'зарлага' },
  toollogo: { folder: 'Тооллого', word: 'тооллого' },
  shiljuuleg: { folder: 'Шилжүүлэг', word: 'шилжүүлэг' },
  butsaalt: { folder: 'Шилжүүлэг', word: 'шилжүүлэг' },
};
const OCR_OTHER_KIND = { folder: 'Бусад', word: 'бусад' }; // a folder of mom's own making
// Where the month spreadsheets of every kind made before go once their lines are split (ocr-split.gs):
// Баримт/Зураг уншуулах/2026-10/хуучин/OCR Соронзон 2026-10 (хуучин), their ids kept.
const OCR_OLD_BOOKS_FOLDER = 'хуучин';
const OCR_OLD_BOOK_MARK = ' (хуучин)';
// The spreadsheets of the first version, one per day next to the photos ("OCR 2026-10-01"): their
// rows are moved into the month spreadsheets and the old files put in the trash (ocrMigrateDaySheet).
const OCR_OLD_DAY_SPREADSHEET = /^OCR \d{4}-\d{2}-\d{2}$/;
const OCR_RETURN_MARK = '(буцаалт)'; // a transfer folder whose name ends so holds returns
const OCR_NO_DAY_REASON = 'өдрийн хавтсанд хийгээгүй'; // a photo outside a day folder (ocrDocumentCheck); the note goes once mom puts it in one (ocrRefileDays)
const OCR_MISFILED_REASON = 'хавтас буруу байж магадгүй'; // a slip that seems to belong in another folder (ocrMisfiledReason)
// A slip Тосгон sent, in a shop's income folder, is Тосгон's expense and the shop's income at once
// (the owner, 2026-10-28): its rows' folder label is "<shop> / орлого (Тосгоноос)" (ocrSlipFolder).
const OCR_FROM_TOSGON_MARK = ' (Тосгоноос)';
// A slip from the city ("Хотоос" on its organisation line) in a shop's income folder (the owner,
// 2026-10-28): the goods came from the city to Тосгон and on to the shop, so it is Тосгон's income,
// Тосгон's expense and the shop's income at once; its rows say "<shop> / орлого (Хотоос)".
const OCR_CITY = 'Хот';
const OCR_FROM_CITY_MARK = ' (Хотоос)';
// The note such a slip had until then; its rows take the new label in the review (ocrTosgonSentRecord).
const OCR_SENT_BY_TOSGON_NOTE = /^хавтас буруу байж магадгүй: Тосгоноос (.+) руу хавтсанд байх ёстой юм шиг \(илгээгч: /;
// A Тосгон or city slip's photo also lies on the other sides (ocr-mirror.gs): + the photo's id -> its
// first copy's id, + "<photo id>#2" -> its second (a city slip's copy in Тосгон's income folder).
const OCR_MIRROR_PREFIX = 'ocr-mirror:';
const OCR_MIRROR_DONE = 'хуулбар: '; // ocr-done:<copy id> -> "хуулбар: <photo id>": the copy is never read
const OCR_MONTH_FOLDER = /^\d{4}-\d{2}$/; // "2026-10"
const OCR_DAY_FOLDER = /^\d{2}$/; // "01" .. "31"
const OCR_TIME_ZONE = 'Asia/Ulaanbaatar';

// The product names a receipt line may be matched to: column A of these tabs (from row 4), together
// with the current half-month sheet of each shop, which also gives the selling price (ocr-catalog.gs).
const OCR_CATALOG_SPREADSHEET_ID = 'YOUR_PRODUCT_LIST_SPREADSHEET_ID';
const OCR_CATALOG_TABS = ['Sheet100', 'Sheet100*'];
const OCR_CATALOG_FIRST_ROW = 4;
// The one product list of both shops (product-list.gs of the workbook): read before every other tab,
// so its spelling wins. Its category headings have no shop in column B and are not products.
const OCR_PRODUCT_LIST_TAB = 'Бараа';
const OCR_PRODUCT_LIST_SHOP_INDEX = 1; // column B: which shop has the product

const OCR_MODEL = 'gemini-3.8-flash';
// With free thinking one photo took 40-80 s; UrlFetchApp gives up at about 60 s. 1024 keeps it at 6-16 s.
const OCR_THINKING_BUDGET = 1024;
const OCR_PRICE_TOLERANCE = 0.1; // a receipt price more than 10% off the price the customer sees is another product (a sale is less)
const OCR_BRAND_SIMILARITY = 0.6; // "Aucma 195" must not become "Panasonica 195l"
const OCR_TIME_BUDGET_MS = 4.5 * 60 * 1000; // a run stops taking new photos after this; Google stops it at 6 min
const OCR_TRIGGER_HOURS = 1;
const OCR_LOOKBACK_DAYS = 62; // photos changed longer ago than this are not looked for any more
const OCR_MONTHS_AHEAD = 1; // day folders are made for this month and this many months after it

// A photo is sent to Gemini ONCE (the owner's rule: a wrong reading is fixed by a person, not by
// reading again). So: a failed call is not tried again, the photo is filed with a "Шалгах" row; the answer is
// kept (OCR_RESULT_PREFIX) until the photo is filed, so a run cut off in between does not ask
// again; DRY_RUN does not call Gemini at all; a photo already filed (same md5) is a duplicate.
const OCR_KEY_PROPERTY = 'GEMINI_KEY';
const OCR_SEEN_PREFIX = 'ocr-seen:'; // + md5 of a photo already filed -> where it went
const OCR_DUPLICATE_REASON = 'давхардсан, өмнө уншсан'; // + ": <where>" on the one row of a second copy
const OCR_RESULT_PREFIX = 'ocr-result:'; // + md5 -> Gemini's answer, until the photo is filed
// + file id -> where its rows were written. Photos stay where mom put them, so this is how a run
// knows which ones it already read.
const OCR_DONE_PREFIX = 'ocr-done:';

const OCR_PLACES = ['Тосгон', 'Соронзон', 'My home'];
// The names mom writes for the places in a slip's header (the owner, 2026-10-28: "Төв" on the
// "Байгууллагын нэр" line is Тосгон, "Ачит" on the buyer's line is My home). The prompt tells
// Gemini (ocr-gemini.gs); a sender read as one of them still counts as that place (ocrSentByTosgon).
const OCR_PLACE_ALIASES = { 'Тосгон': ['Төв', 'Тос', 'Тов', 'Склад', 'Склд', 'Скл'], 'My home': ['Ачит', 'Myhome'] }; // Склд: the warehouse (the owner, 2026-10-29)
// Тосгон's shelf price "яг харагдах үнэ" in E (2026-10-22, its layout-migration.gs): its days start at F.
const OCR_TOSGON_SHELF = true;
const OCR_TOSGON = { place: 'Тосгон', tabPrefix: 'Tos', layout: buildColumnLayout(30, { shelf: OCR_TOSGON_SHELF }) };
// What Gemini accepts; Drive calls an iPhone photo image/heif.
const OCR_MIME_TYPES = {
  'image/jpeg': 'image/jpeg',
  'image/png': 'image/png',
  'image/webp': 'image/webp',
  'image/heic': 'image/heic',
  'image/heif': 'image/heic',
  'application/pdf': 'application/pdf',
};

// The photo column (M): a title only, its cells stay empty, the photo lies over them (the owner,
// 2026-10-28: the photo among the columns, the reason right after it).
const OCR_PHOTO_COLUMN = 'Зураг';
const OCR_REASON_COLUMN = 'Шалгах шалтгаан';
const OCR_HEADER = [
  'Харах', // A1 holds the view choice (OCR_VIEWS); the rows below stay empty
  'Бараа', 'Тоо', 'Ээж: Зөв/Буруу',
  // the script's, hidden in mom's view
  'Төлөв', 'Бичсэн', 'Хавтас (ээж хийсэн)', 'Баримт', 'Бараа (тааруулсан)', 'Тоо (уншсан)', 'Итгэл', 'ID',
  // the photo, then why a line is pink, in mom's view (the owner, 2026-10-28: she sees the photo and
  // the reason side by side without opening "Бүгд"; up to then the reason stood after "Төлөв",
  // hidden, and the photo right after the last column)
  OCR_PHOTO_COLUMN, OCR_REASON_COLUMN,
];
// Mom sees the product, the count and her mark, the photo beside them; nothing else (the owner,
// 2026-10-25: what was read as text, the receipt's price, title and kind, the date (the tab is the
// day), the link, the price the customer sees, where a line goes and what mom typed in the workbook
// are gone). Gemini's confidence is back in "Итгэл" (2026-10-25), hidden: a line under
// OCR_REVIEW_CONFIDENCE is Шалгах, and the owner compares it with mom's marks to tune the threshold. "Бараа" and "Тоо" are what the line goes in as: the script puts
// the product Gemini picked (always one of the list, or Хамааралгүй) and the count it read there (also
// kept in "Бараа (тааруулсан)" and "Тоо (уншсан)"); mom picks another product from the list or writes
// another count, and that change is her approval (no Зөв needed). Otherwise she marks Зөв/Буруу.
// Her columns are the only ones she may change besides the view choice in A1 (ocr-sheet.gs protects
// the rest); they have no colour of their own (the owner, 2026-10-26: purple was taken away). The rest is the script's: "Төлөв" and "Шалгах шалтгаан" (why a line waits),
// "Бичсэн" (where the line was written in the workbook, to move or take it out again), "Хавтас"
// (mom's folder of the photo: where the line goes), "Баримт" (the photo's file name, for the cell
// notes) and "ID" (photo file id # line number, also in the workbook cell's note).
const OCR_NAME_COLUMN = 'Бараа';
// A line Gemini is less sure of than this is Шалгах, beside the mismatch rules (the owner, 2026-10-25;
// no threshold 2026-10-23..25). It never changes the product Gemini picked.
const OCR_REVIEW_CONFIDENCE = 0.7;
const OCR_QTY_COLUMN = 'Тоо';
// The owner's rules for Тосгон's income (2026-10-23): an issue slip in its income folder is not read
// into lines; a photo read almost like an earlier one of its folder is the same slip twice, and waits
// for mom's Зөв before it is written (ocr-twins.gs).
const OCR_SKIPPED_SLIP_TEXT = 'Алгассан: зарлагын баримт орлогын хавтаст байсан';
const OCR_TWIN_REASON = 'давхар зураг';
const OCR_TWIN_SHARE = 0.6; // of both photos' lines (product and count) alike
const OCR_TWIN_MIN_LINES = 2; // one line alike is no twin: the same product is often sold twice
const OCR_RECEIPT_THOUSANDS = 10000; // a receipt price under this is written in thousands ("130" = 130,000₮)
const OCR_MOM_MARKS = ['Ээж: Зөв/Буруу'];
const OCR_MOM_COLUMNS = [OCR_NAME_COLUMN, OCR_QTY_COLUMN, ...OCR_MOM_MARKS];
// The view choice in A1: mom's view shows what was read, the product, count and prices, and her
// mark; "Бүгд" shows every column (also why a line is not in the workbook yet). Changing it shows
// or hides the columns at once (onOcrViewEdit, ocr-sheet.gs).
const OCR_VIEW_MOM = 'Ээж';
const OCR_VIEW_ALL = 'Бүгд';
const OCR_VIEWS = [OCR_VIEW_MOM, OCR_VIEW_ALL];
// The third choice in A1 (the owner, 2026-10-29): mom checked the whole day in the workbook, so
// every line she has not looked at yet gets Зөв at once, not one Зөв a line; A1 then shows the view
// it showed before (onOcrViewEdit, ocrMarkDayOk). It is no view: the day turns green on the site.
const OCR_MARK_DAY_OK = 'Бүх мөр Зөв';
const OCR_VIEW_CHOICES = [...OCR_VIEWS, OCR_MARK_DAY_OK];
// Mom's view is the product, the count and her mark, then the photo and why a line is pink (the
// owner, 2026-10-25: she checks at a glance; 2026-10-28: the reason too, as wide as it needs,
// ocrFitReasonColumn).
const OCR_VISIBLE_COLUMNS = ['Харах', OCR_NAME_COLUMN, OCR_QTY_COLUMN, ...OCR_MOM_MARKS, OCR_PHOTO_COLUMN, OCR_REASON_COLUMN];
const OCR_VIEW_HANDLER = 'onOcrViewEdit';
const OCR_VIEW_TRIGGERS_MAX = 16; // the newest month spreadsheets keep their view trigger (a script may have 20 triggers; a month has up to 11 spreadsheets)
// The columns of the versions before (2026-10-19): an old tab is moved into OCR_HEADER by its
// titles (ocrUpgradeRecord, ocr-review.gs): the staff's correction becoming mom's, the matched name
// and mom's corrected name becoming "Бараа (тааруулсан)" and "Бараа".
const OCR_OLD_STAFF_COLUMNS = { mark: 'Ажилчин: Зөв/Буруу', name: 'Ажилчин: зөв бараа', qty: 'Ажилчин: зөв тоо', note: 'Ажилчин: тайлбар' };
const OCR_OLD_NAME_COLUMNS = { matched: 'Бараа (Sheet100)', corrected: 'Ээж: зөв бараа' };
// The columns of the version before 2026-10-20: mom's count and note (her count becomes "Тоо").
const OCR_OLD_MOM_COLUMNS = { qty: 'Ээж: зөв тоо', note: 'Ээж: тайлбар' };
const OCR_MOM_COLOR = '#d9d2e9'; // mom's columns until 2026-10-26: taken away (ocrClearMomColor)
const OCR_AGREED_COLOR = '#d9ead3'; // mom said Зөв
const OCR_WRONG_COLOR = '#f4cccc'; // mom said Буруу
// Every colour the system puts on means one thing (the owner, 2026-10-25): pink = must be checked,
// in the day tab (a Шалгах row mom has not marked) and in the workbook (a cell with such a line);
// yellow = the AI is sure, in the day tab (an OK row) and in the workbook alike. Pink here
// is none of the transfer link's pinks (#f8bbd0, #f06292, #d5a6bd, dividers.gs).
const OCR_CHECK_COLOR = '#ff99dd';
const OCR_VERDICT_OK = 'Зөв';
const OCR_VERDICT_WRONG = 'Буруу';
// A sale given away (the owner, 2026-10-28): mom's third choice in the two shops' sales spreadsheets.
// The line is approved and goes into the shop's gift column (LAYOUT.gift, AL "Бэлэг") in place of its
// day's sales column.
const OCR_VERDICT_GIFT = 'Бэлэг';
// The first choice of "Бараа": goods the shops never count (a bicycle, say). Such a line is never
// written to a workbook and waits for no one. Its "price" on the name list is no number, so the
// sale price column stays empty.
const OCR_NOT_RELEVANT = 'Хамааралгүй';
const OCR_NOT_RELEVANT_COLOR = '#efefef'; // grey row

// The Gmail address never goes in this file (it is kept in git): Script property OCR_REVIEWERS,
// JSON {"Ээж": "...@gmail.com"} - the account mom checks with (the shop account), when it is not the
// owner account; several accounts as "a@gmail.com, b@gmail.com" or a list (the owner, 2026-10-27:
// her personal account too) (the owner account runs the script and owns Баримт). Then mom changes only her
// columns, never what was read; without it mom uses the owner account, which Google lets change
// everything. Other keys (the staff addresses of the version before) are ignored.
const OCR_REVIEWERS_PROPERTY = 'OCR_REVIEWERS';
const OCR_MOM_KEY = 'Ээж';

// The workbook each place's lines are written into. A transfer goes into both places: e.g.
// "My home-оос Соронзон руу" is My home's transfer out (AJ) and Соронзон's transfer in (T) of that
// half-month sheet. A place left out here is not written to.
const OCR_TOSGON_SPREADSHEET_ID = 'YOUR_TOSGON_BOOK_ID'; // Тосгон 2026
const OCR_POST_BOOKS = { 'Тосгон': OCR_TOSGON_SPREADSHEET_ID, 'Соронзон': OCR_CATALOG_SPREADSHEET_ID, 'My home': OCR_CATALOG_SPREADSHEET_ID };
// The workbook of a place for a tab ("Sor/2027/1/15", "Tos/2027/1"): that year's file in the
// registry (year-books.gs), else the one above. A place left out of OCR_POST_BOOKS stays out.
function ocrBookIdFor(place, tabName) {
  if (!OCR_POST_BOOKS[place]) {
    return undefined;
  }
  const year = Number(String(tabName).split('/')[1]);
  return bookIdOfYear(year, place === OCR_TOSGON.place ? 'tosgon' : 'main') || OCR_POST_BOOKS[place];
}
// A year whose receipts have workbooks to go into: the home year, or one the registry knows.
const ocrHasBooksFor = (year) => year === OCR_WORKBOOK_YEAR || bookIdOfYear(year, 'main') !== null;
const OCR_BOOK_FIRST_ROW = 4; // product rows start here, name in column A
const OCR_POSTED_FONT_COLOR = '#1155cc'; // a number the script wrote is blue
const OCR_POST_FAILED = '✗';
// Every line with a product is written at once, whatever its day (in OCR_WORKBOOK_YEAR, the year
// the workbooks hold): its cell turns yellow, green once mom approved every line in it. A cell
// mom had typed by hand stays as it is: the line goes into the same cell of the workbook's full
// copy ("… (check)", made once), so the two files can be compared side by side; the hand-typed
// cell turns yellow too when the receipts make the same number.
const OCR_WORKBOOK_YEAR = 2026;
// Workbook cells the OCR never writes, colours or notes (the owner, 2026-10-23): mom typed these
// goods on the 22nd (L95, L101, L134), so the receipt's cells of the 16th were set to 0 by hand,
// with a note line of their own. A receipt line that would touch one of them is left waiting.
const OCR_HANDS_OFF_CELLS = ['Sor/2026/4/30!F95', 'Sor/2026/4/30!F101', 'Sor/2026/4/30!F134'];
// The shops' sheets the OCR never writes, colours or notes again (the owner, 2026-10-29): Соронзон's
// and My home's half-months of 4/1 to 8/15 were typed by hand, and what the OCR had written there was
// taken out (its numbers, notes and colours; mom's numbers stayed). A line of those days says so in
// "Бичсэн" (OCR_HANDS_OFF_TEXT) and is not written. A half-month is out when its last day is in the
// period, its check copy's too; Тосгон's months are not.
const OCR_HANDS_OFF_PERIOD = { from: '2026-04-01', until: '2026-08-15' };
const OCR_HANDS_OFF_TEXT = `— ${OCR_HANDS_OFF_PERIOD.from} – ${OCR_HANDS_OFF_PERIOD.until}: дэлгүүрүүдийн workbook-д OCR бичихгүй`;
// "Sor/2026/5/31" (a shop's half-month, by its last day) in OCR_HANDS_OFF_PERIOD
function ocrIsHandsOffTab(tabName) {
  const parts = String(tabName).split('/');
  if (parts.length !== 4 || !parts.slice(1).every((part) => /^\d+$/.test(part))) {
    return false;
  }
  const [, year, month, day] = parts;
  const last = `${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`;
  return last >= OCR_HANDS_OFF_PERIOD.from && last <= OCR_HANDS_OFF_PERIOD.until;
}
// "Sor/2026/5/31!T189", "check:Sor/2026/5/31!T189"
function ocrIsHandsOffCell(cell) {
  return ocrIsHandsOffTab(String(cell).replace(OCR_CHECK_PREFIX, '').split('!')[0]);
}
const OCR_OTHER_YEAR_TEXT = `— ${OCR_WORKBOOK_YEAR} оноос өөр оны баримт: workbook-д бичихгүй`;
const OCR_WRITTEN_COLOR = '#ffe599'; // yellow: written by the system, the AI sure of it (the transfer columns are a paler #fff2cc)
const OCR_APPROVED_CELL_COLOR = '#d9ead3'; // green: mom approved every line in the cell
// Mom's hand-typed cell: green when the receipts make her number (checked: both say so), pink when
// they make another (its note then says "Баримтаар: n": check it), no colour when no receipt is for
// it. Three colours only in the workbooks (the owner, 2026-10-28): pink = check, yellow = the system
// is sure, green = checked and right; until then a hand-typed cell was light blue or light red.
const OCR_HAND_MATCH_COLOR = OCR_APPROVED_CELL_COLOR;
const OCR_HAND_DIFF_COLOR = OCR_CHECK_COLOR;
// The light blue and light reds of before (#f4c7c3 until 2026-10-23, #e6b8af until 2026-10-28): a
// cell still in one is given the new colour by the next review (ocrHandLinesCurrent sees it is not
// the wanted one), so they are still the OCR's own.
const OCR_HAND_OLD_COLORS = ['#c9daf8', '#e6b8af', '#f4c7c3'];
const OCR_HAND_NOTE_MARK = 'Баримтаар:';
// A cell's own colour (a transfer column's, say) is never taken away: only these are cleared.
const OCR_CELL_COLORS = [OCR_WRITTEN_COLOR, OCR_CHECK_COLOR, OCR_APPROVED_CELL_COLOR, ...OCR_HAND_OLD_COLORS];
const OCR_APPROVED_MARK = '✓'; // on a note line and a "Бичсэн" line of an approved line
const OCR_MUST_CHECK_MARK = '⚠'; // ... of a Шалгах line mom has not marked yet (its cell is pink)
const OCR_CHECK_PREFIX = 'check:'; // "check:Sor/2026/10/15!U22": a cell of the check copy
const OCR_CHECK_SUFFIX = ' (check)';
const OCR_CHECK_COPY_PROPERTY = 'ocr-check-copy:'; // + workbook id -> the id of its check copy
// Shop income before this day is copied into Тосгон's expense by ocr-backfill.gs (the days the
// shops typed by hand before the OCR was used).
const OCR_POST_FROM = '2026-10-20';
// Тосгон typed its own expense by hand up to this day (often on another day than the shop's
// income), so ocr-backfill.gs only adds what those hand-typed numbers leave over.
const OCR_TOSGON_HAND_UNTIL = '2026-05-15';
const OCR_COMPARE_MARK = '≈'; // "Бичсэн" of a line of the version before (compared only): written now
const OCR_LAST_SCAN_PROPERTY = 'ocr-last-scan';
// Mom's address the sheets were last protected for; when it changes, every recent month
// spreadsheet is looked at again (not only the changed ones).
const OCR_REVIEWERS_APPLIED_PROPERTY = 'ocr-reviewers-applied';
const OCR_SCAN_FROM_PROPERTY = 'ocr-scan-from'; // where a pass cut off by time goes on
const OCR_REVIEW_CUT_PROPERTY = 'ocr-review-cut'; // the month spreadsheet a pass stopped in: first the next time
const OCR_BOOK_SETUP_PREFIX = 'ocr-book-setup:'; // + month spreadsheet id -> what its setup was made from (ocrSetUpBook)
const OCR_BOOK_SETUP_PROGRESS_PREFIX = 'ocr-book-setup-tabs:'; // + month spreadsheet id -> { wanted, tabs, isGuarded } of a setup cut by the time
const OCR_BOOK_SETUP_VERSION = 13; // change it with a change of ocrSetUpBook's parts: every month spreadsheet is set up again
const OCR_IMPORT_ALLOWED_PREFIX = 'ocr-import-ok:'; // + month spreadsheet id -> the workbook id its name list may import (ocrAllowImport)
const OCR_SCAN_MARGIN_MS = 30 * 60 * 1000; // Drive's search sees a change some minutes late
const OCR_OPEN_PREFIX = 'ocr-open:'; // + spreadsheet id -> what is still to do in it (ocr-report.gs)
const OCR_DAY_PREFIX = 'ocr-day:'; // + YYYY-MM-DD -> what happened that day, for the report
const OCR_DAY_KEEP_DAYS = 16; // a whole half-month (16-31), for the report of its last day
const OCR_REPORT_RESERVE_MS = 60 * 1000; // time kept for writing and the report

// A hidden tab in each day spreadsheet with the product names, so a name is chosen from a list and
// is always written the same way.
const OCR_NAMES_SHEET = '_нэрс';
const OCR_REPORT_FOLDER = 'Тайлан'; // Баримт/Тайлан/2026-10/Тайлан 2026-10-18, for mom
const OCR_REPORT_PREFIX = 'Тайлан ';

// The photo is shown to the right of its own rows, to compare with them. Drive gives a JPEG of any
// photo (also HEIC, which Sheets can not show); Sheets takes images up to 1 million pixels, so the
// longest side is fetched at 1000 px. Every row stays as tall as any (the owner, 2026-10-26): the
// photo lies over as many rows as it needs, blank ones below a short receipt, and the next
// receipt's rows and photo start below it (ocrNextFreeRow).
const OCR_IMAGE_SIZE = 1000;
const OCR_IMAGE_WIDTH = 420; // px on the sheet
// M (OCR_PHOTO_COLUMN), as wide as the photo, the reason after it (the owner, 2026-10-28); before,
// the photo stood in N, right after the last column (2026-10-25..28), and in O before that
const OCR_IMAGE_COLUMN = OCR_HEADER.indexOf(OCR_PHOTO_COLUMN) + 1;
const OCR_IMAGE_COLUMN_WIDTH = OCR_IMAGE_WIDTH + 8; // px
const OCR_OLD_IMAGE_COLUMNS = [OCR_HEADER.length, OCR_HEADER.length + 1]; // N, O
const OCR_ROW_HEIGHT = 21; // Sheets' default row height
const OCR_IMAGE_GAP = 12; // px between one receipt's photo and the next
const OCR_STATUS_OK = 'OK';
const OCR_STATUS_REVIEW = 'Шалгах';
