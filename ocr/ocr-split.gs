/**
 * Receipt OCR - the month spreadsheets of every kind made before 2026-10-25 ("OCR Соронзон 2026-10")
 * split into the spreadsheets of their kinds (ocrMonthBook): by the file alone nobody could tell
 * income from expense. Every line of a day tab goes, by its folder (Хавтас), into the same day tab
 * of its kind's spreadsheet, its photo beside it again; mom's marks, "Бичсэн" and the ID go with it,
 * so the workbook cells (their notes name the ID) stay as they are and nothing is written twice.
 * A line already there (same ID) is not added again and a day done is remembered (OCR_SPLIT_PREFIX),
 * so a run cut off half-way does no harm. The old spreadsheet is then renamed "... (хуучин)" and
 * moved into Баримт/Зураг уншуулах/<month>/хуучин, its id kept; what the script kept about it (report lines,
 * calendar, setup, view trigger) is let go. Newest month first.
 * Run splitOcrBooksByKind from the editor until its log says nothing is left.
 */

const OCR_SPLIT_PREFIX = 'ocr-split:'; // + old spreadsheet id + ':' + day -> that day is split
const OCR_SPLIT_BUDGET_MS = 5 * 60 * 1000;

// ---------- pure helpers (no Drive calls) ----------

// The lines of one old day tab by kind, in the order the kinds first come: [{ kind, records }].
// An empty row (no ID and no photo) is left out.
function ocrSplitByKind(records) {
  return records
    .filter((record) => ocrText(record['ID']) !== '' || ocrText(record['Баримт']) !== '')
    .reduce((groups, record) => {
      const kind = ocrBookKind(ocrText(record['Хавтас (ээж хийсэн)']).split(' / '));
      const at = groups.findIndex((group) => group.kind.word === kind.word);
      return at < 0
        ? [...groups, { kind, records: [record] }]
        : groups.map((group, index) => (index === at ? { ...group, records: [...group.records, record] } : group));
    }, []);
}

// Newest month first, then by place.
const ocrSplitOrder = (entries) => [...entries].sort((a, b) => b.month.localeCompare(a.month) || a.place.localeCompare(b.place));

// ---------- Drive access ----------

function splitOcrBooksByKind() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(OCR_LOCK_WAIT_MS)) {
    Logger.log('OCR ажиллаж байна, хэсэг хүлээгээд дахин Run хийнэ үү.');
    return null;
  }
  try {
    return ocrSplitOldBooks(Date.now() + OCR_SPLIT_BUDGET_MS);
  } finally {
    lock.releaseLock();
  }
}

function ocrSplitOldBooks(deadline) {
  const props = PropertiesService.getScriptProperties();
  const root = DriveApp.getFolderById(OCR_ROOT_FOLDER_ID);
  const books = ocrSplitOrder(findOldMonthBooks(root));
  let done = 0;
  while (done < books.length && Date.now() <= deadline && ocrSplitBook(root, books[done], props, deadline)) {
    done += 1;
  }
  const left = books.length - done;
  Logger.log(`Төрлөөр салгах (${modeLabel()}): ${done} файл дууссан, ${left} үлдсэн` +
    (left > 0 ? '; splitOcrBooksByKind-ийг дахин ажиллуулна уу.' : '.'));
  return { done, left };
}

// The month spreadsheets of every kind still under Баримт: [{ file, place, month }].
function findOldMonthBooks(root) {
  const found = [];
  ocrIterate(DriveApp.searchFiles(ocrSheetsQuery(null)), (file) => {
    const old = ocrOldMonthBookOf(file.getName());
    if (old && ocrIsUnder(file, root.getId())) {
      found.push({ file, ...old });
    }
  });
  return found;
}

// One old spreadsheet, day by day until the deadline; true once every day is split and it is put away.
function ocrSplitBook(root, entry, props, deadline) {
  const id = entry.file.getId();
  const tabs = ocrDayTabs(SpreadsheetApp.openById(id));
  for (const tab of tabs) {
    const key = `${OCR_SPLIT_PREFIX}${id}:${tab.getName()}`;
    if (props.getProperty(key) !== null) {
      continue;
    }
    if (Date.now() > deadline) {
      return false;
    }
    ocrSplitByKind(ocrReadRecords(tab)).forEach(({ kind, records }) => ocrCopyLines(root, entry, kind, tab.getName(), records));
    if (!DRY_RUN) {
      props.setProperty(key, '1');
    }
  }
  ocrRetireOldBook(root, entry, props);
  return true;
}

// The lines of one kind and day into its spreadsheet, each photo beside its lines again.
function ocrCopyLines(root, entry, kind, day, records) {
  const name = ocrMonthBookName(entry.place, kind, entry.month);
  if (DRY_RUN) {
    Logger.log(`  ${entry.file.getName()} / ${day}: ${records.length} мөр -> ${kind.folder}/${name} (туршилт)`);
    return;
  }
  const { book } = ocrMonthBook(root, entry.place, kind, entry.month);
  const tab = ocrDayTab(book, day);
  const idColumn = OCR_HEADER.indexOf('ID');
  const last = ocrLastRecordRow(tab);
  const known = new Set(last > 1 ? tab.getRange(2, idColumn + 1, last - 1, 1).getValues().map(([value]) => ocrText(value)) : []);
  const isNew = (record) => ocrText(record['ID']) === '' || !known.has(ocrText(record['ID']));
  const rows = records.filter(isNew).map((record) => ocrRow(record));
  ocrGroupByPhoto(rows).forEach((group) => {
    const first = ocrAppendRows(tab, group.rows);
    try {
      ocrAddPhoto(tab, first, group.rows.length, DriveApp.getFileById(group.photoId));
    } catch (error) {
      Logger.log(`  ${group.photoId}: зураг олдсонгүй (${error.message})`);
    }
  });
  Logger.log(`  ${entry.file.getName()} / ${day}: ${rows.length} мөр -> ${kind.folder}/${name}`);
}

// The split spreadsheet renamed and put away (its id kept), and what was kept about it let go.
function ocrRetireOldBook(root, entry, props) {
  const id = entry.file.getId();
  const sheets = ocrSheetsFolder(root, false); // its real name: "OCR" until renamed
  const where = `${sheets ? sheets.getName() : OCR_SHEETS_FOLDER}/${entry.month}/${OCR_OLD_BOOKS_FOLDER}`;
  if (DRY_RUN) {
    Logger.log(`${entry.file.getName()}: "${entry.file.getName()}${OCR_OLD_BOOK_MARK}" болж ${where} руу орно (туршилт)`);
    return;
  }
  entry.file.setName(`${entry.file.getName()}${OCR_OLD_BOOK_MARK}`);
  entry.file.moveTo(ocrEnsurePath(ocrSheetsFolder(root), [entry.month, OCR_OLD_BOOKS_FOLDER]));
  const kept = Object.keys(props.getProperties()).filter((key) =>
    key.startsWith(`${OCR_SPLIT_PREFIX}${id}:`) || key.startsWith(`${OCR_OPEN_PREFIX}${id}:`) || key === OCR_BOOK_SETUP_PREFIX + id);
  [...kept, ocrCalendarKey(entry.place, entry.month), ocrCalendarLinkKey(entry.place, entry.month)].forEach((key) => props.deleteProperty(key));
  if (props.getProperty(OCR_REVIEW_CUT_PROPERTY) === id) {
    props.deleteProperty(OCR_REVIEW_CUT_PROPERTY);
  }
  ScriptApp.getProjectTriggers()
    .filter((trigger) => trigger.getHandlerFunction() === OCR_VIEW_HANDLER && trigger.getTriggerSourceId() === id)
    .forEach((trigger) => ScriptApp.deleteTrigger(trigger));
  Logger.log(`${entry.file.getName()}: бүх өдөр салсан, ${where} руу орлоо`);
}
