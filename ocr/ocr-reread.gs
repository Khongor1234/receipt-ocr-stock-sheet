/**
 * Receipt OCR - one day tab read again with the current rules (the owner, 2026-10-25: Тосгон's count of
 * 10/23 had been read before every line had to get a product). Every photo of the tab is asked once
 * more and its lines take the place of the tab's rows. A line keeps its ID (photo # line number), so
 * the workbook cells (their notes name the ID) are put right by the next review:
 *   - the old row's "Бичсэн" goes with the new row of the same ID: a changed product or count moves
 *     its entry, as a correction of mom's would
 *   - an old line the new reading does not have any more becomes Хамааралгүй: its entry is taken out
 *   - a line mom approved (her Зөв, or a product or count she chose herself) keeps her product, count
 *     and mark whatever the new reading says (the owner, 2026-10-25: hers is right); on any other line
 *     mom's Буруу stays only while it still reads the same product and count
 * A photo Gemini could not read keeps its old rows. Run rereadTosgonCount1023 from the editor, or
 * rereadSoronzonIncome10 (rereadMyHomeIncome10) for every tab of Соронзон's (My home's) 2026-10 income
 * whose sender was unknown: a slip from Тосгон ("Тов") or the city ("Хотоос") takes its sides then.
 */

const OCR_REREAD_GONE_REASON = 'дахин уншихад энэ мөр гараагүй';
const OCR_TOSGON_COUNT_1023 = { bookId: 'YOUR_TOSGON_COUNT_BOOK_ID', day: '23' };
// OCR My home орлого 2026-10, tab 23 (the owner, 2026-10-28: a slip whose header says "Төв" and "Ачит").
const OCR_MY_HOME_INCOME_1023 = { bookId: 'YOUR_MY_HOME_INCOME_BOOK_ID', day: '23' };

// OCR Соронзон орлого 2026-10 (the owner, 2026-10-28): its slips from Тосгон say "Төв" and were read
// before that name was known, so their lines say the sender is unknown.
const OCR_SORONZON_INCOME_10 = 'YOUR_SORONZON_INCOME_BOOK_ID';
const OCR_UNKNOWN_SENDER_MARK = 'илгээгч тодорхойгүй'; // in the misfiled reason (ocrMisfiledReason)

// ---------- pure helpers (no Drive calls) ----------

// A reason only rules since changed wrote (the owner, 2026-10-28: all of October without old
// problems): the sender unknown (mom's "Төв", "Хот" read since), a Тосгон slip noted as misfiled (it
// counts on both sides since), a low confidence without its number (the reading before 2026-10-23).
const OCR_STALE_REASONS = [/илгээгч тодорхойгүй/, /руу хавтсанд байх ёстой юм шиг/, /итгэл бага(?! \()/];
const ocrIsStaleRecord = (record) => OCR_STALE_REASONS.some((pattern) => pattern.test(ocrText(record['Шалгах шалтгаан'])));
const OCR_STALE_DONE_PREFIX = 'ocr-stale-done:'; // + month: the tabs read again ("<book id>/<day>"), JSON
const OCR_STALE_HANDLER = 'rereadOctober'; // the run that goes on when one ran out of time
const OCR_STALE_GO_ON_MS = 60 * 1000;
const OCR_REREAD_PHOTO_MS = 30 * 1000; // about one photo's reading and writing: a tab is begun only with time for all of it
const ocrStaleLabel = (tab) => `${tab.name} / ${tab.day}`;

const ocrHasUnknownSender = (records) => records.some((record) => ocrText(record['Шалгах шалтгаан']).includes(OCR_UNKNOWN_SENDER_MARK));

const ocrPhotoIdOf = (record) => ocrText(record['ID']).split('#')[0];

// The rows of a tab read again: fresh rows (ocrRows) of the photos in `fresh` (their ids), the old
// rows of the other photos as they are, each photo's lines together in the order the photos came.
function ocrMergeReread(oldRecords, freshRecords, fresh) {
  const old = new Map(oldRecords.filter((record) => ocrText(record['ID']) !== '').map((record) => [ocrText(record['ID']), record]));
  const freshIds = new Set(freshRecords.map((record) => ocrText(record['ID'])));
  const photos = [...new Set([...oldRecords, ...freshRecords].map(ocrPhotoIdOf).filter(Boolean))];
  return photos.flatMap((photo) => {
    const before = oldRecords.filter((record) => ocrPhotoIdOf(record) === photo);
    if (!fresh.has(photo)) {
      return before;
    }
    const kept = freshRecords.filter((record) => ocrPhotoIdOf(record) === photo).map((record) => ocrCarryOver(record, old.get(ocrText(record['ID']))));
    const gone = before.filter((record) => ocrText(record['ID']) !== '' && !freshIds.has(ocrText(record['ID']))).map(ocrGoneLine);
    return [...kept, ...gone];
  });
}

// A fresh line with what the old line of its ID had: where it was written, and mom's say.
function ocrCarryOver(record, old) {
  if (!old) {
    return record;
  }
  // a mark on a line that had no product says nothing of one: the new reading's product goes in
  const isMine = ocrIsApproved(old) && ocrText(old[OCR_NAME_COLUMN]) !== '';
  const isSame = ocrNameKey(old[OCR_NAME_COLUMN]) === ocrNameKey(record[OCR_NAME_COLUMN]) && ocrText(old[OCR_QTY_COLUMN]) === ocrText(record[OCR_QTY_COLUMN]);
  return {
    ...record,
    'Бичсэн': old['Бичсэн'],
    ...(isMine ? { [OCR_NAME_COLUMN]: old[OCR_NAME_COLUMN], [OCR_QTY_COLUMN]: old[OCR_QTY_COLUMN] } : {}),
    ...(isMine || isSame ? { 'Ээж: Зөв/Буруу': old['Ээж: Зөв/Буруу'] } : {}),
  };
}

// An old line the new reading does not have: Хамааралгүй, so the review takes it out of the workbook.
const ocrGoneLine = (record) => ({
  ...record,
  [OCR_NAME_COLUMN]: OCR_NOT_RELEVANT,
  'Төлөв': OCR_STATUS_REVIEW,
  'Шалгах шалтгаан': [OCR_REREAD_GONE_REASON, ocrText(record['Шалгах шалтгаан'])].filter(Boolean).join('; '),
});

// ---------- Drive access ----------

function rereadTosgonCount1023() {
  return rereadOcrDay(OCR_TOSGON_COUNT_1023.bookId, OCR_TOSGON_COUNT_1023.day);
}

function rereadMyHomeIncome1023() {
  return rereadOcrDay(OCR_MY_HOME_INCOME_1023.bookId, OCR_MY_HOME_INCOME_1023.day);
}

function rereadSoronzonIncome10() {
  return rereadUnknownSenderDays(OCR_SORONZON_INCOME_10);
}

// The same for My home's 2026-10 income (the owner, 2026-10-28: its slips from the city were read as
// sender unknown too; read again, they count as Тосгон's income, Тосгон's expense and My home's income).
function rereadMyHomeIncome10() {
  return rereadUnknownSenderDays(OCR_MY_HOME_INCOME_1023.bookId);
}

// Every day tab of the book with a line whose sender was unknown, read again (ocrRereadDay); a run
// stops taking new tabs after OCR_TIME_BUDGET_MS and names the ones left: Run it again for them.
function rereadUnknownSenderDays(bookId) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(OCR_LOCK_WAIT_MS)) {
    Logger.log('OCR ажиллаж байна, хэсэг хүлээгээд дахин Run хийнэ үү.');
    return null;
  }
  try {
    const started = Date.now();
    const days = ocrDayTabs(SpreadsheetApp.openById(bookId))
      .filter((tab) => ocrHasUnknownSender(ocrReadRecords(tab)))
      .map((tab) => tab.getName());
    const done = [];
    for (const day of days) {
      if (Date.now() - started >= OCR_TIME_BUDGET_MS) {
        break;
      }
      ocrRereadDay(bookId, day);
      done.push(day);
    }
    const left = days.slice(done.length);
    Logger.log(`Илгээгч тодорхойгүй өдрүүд: ${days.join(', ') || 'алга'}${left.length ? `; цаг дууслаа, дахин Run хийхэд ${left.join(', ')}` : ''}`);
    return { days: done, left };
  } finally {
    lock.releaseLock();
  }
}

function rereadOctober() {
  return rereadStaleMonth('2026-10');
}

// Every day tab of the month's spreadsheets with a line of an old reason (OCR_STALE_REASONS), read
// again once (ocrRereadDay; mom's marks stay). Out of time, it sets a run of its own a minute later
// for the rest, and takes it away when nothing is left. budgetMs: for the tests.
function rereadStaleMonth(month, budgetMs = OCR_TIME_BUDGET_MS) {
  const started = Date.now();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(OCR_LOCK_WAIT_MS)) {
    Logger.log('OCR ажиллаж байна, хэсэг хүлээгээд дахин Run хийнэ үү.');
    return null;
  }
  try {
    const props = PropertiesService.getScriptProperties();
    const doneKey = OCR_STALE_DONE_PREFIX + month;
    const done = JSON.parse(props.getProperty(doneKey) || '[]');
    const tabs = ocrStaleTabs(month, done);
    if (DRY_RUN) {
      Logger.log(`${month}: хуучин шалтгаантай ${tabs.length} өдөр (туршилт, Gemini дуудахгүй): ${tabs.map(ocrStaleLabel).join(', ') || 'алга'}`);
      return { read: [], left: tabs.map(ocrStaleLabel) };
    }
    ocrSetGoOn(tabs.length > 0); // before the reading: a run Google stops still goes on
    const read = ocrReadStaleTabs(tabs, started, budgetMs, (tab) => {
      props.setProperty(doneKey, JSON.stringify([...JSON.parse(props.getProperty(doneKey) || '[]'), `${tab.bookId}/${tab.day}`]));
    });
    const left = tabs.slice(read.length);
    ocrSetGoOn(left.length > 0);
    Logger.log(`${month}: ${read.length} өдөр дахин уншлаа${left.length ? `; үлдсэн ${left.length} өдрийг минутын дараа өөрөө үргэлжлүүлнэ: ${left.map(ocrStaleLabel).join(', ')}` : ', бүгд дууслаа'}`);
    return { read: read.map(ocrStaleLabel), left: left.map(ocrStaleLabel) };
  } finally {
    lock.releaseLock();
  }
}

// The month's day tabs with an old reason, not read again yet: [{ bookId, name, day, photos }].
function ocrStaleTabs(month, done) {
  const root = DriveApp.getFolderById(OCR_ROOT_FOLDER_ID);
  return findOcrBooks(root, null)
    .filter((entry) => entry.month === month)
    .flatMap((entry) => ocrDayTabs(SpreadsheetApp.openById(entry.file.getId())).map((tab) => ({ entry, tab })))
    .filter(({ entry, tab }) => !done.includes(`${entry.file.getId()}/${tab.getName()}`))
    .map(({ entry, tab }) => ({ entry, tab, records: ocrReadRecords(tab) }))
    .filter(({ records }) => records.some(ocrIsStaleRecord))
    .map(({ entry, tab, records }) => ({ bookId: entry.file.getId(), name: entry.file.getName(), day: tab.getName(), photos: new Set(records.map(ocrPhotoIdOf)).size }))
    .sort((a, b) => ocrStaleLabel(a).localeCompare(ocrStaleLabel(b)));
}

// The tabs read again while each fits in the time (the first always); onRead(tab) after each.
function ocrReadStaleTabs(tabs, started, budgetMs, onRead) {
  const read = [];
  for (const tab of tabs) {
    if (read.length > 0 && Date.now() - started + tab.photos * OCR_REREAD_PHOTO_MS > budgetMs) {
      break;
    }
    ocrRereadDay(tab.bookId, tab.day);
    onRead(tab);
    read.push(tab);
  }
  return read;
}

// The run that goes on with the rest a minute later: set when wanted, else taken away.
function ocrSetGoOn(isWanted) {
  ScriptApp.getProjectTriggers()
    .filter((trigger) => trigger.getHandlerFunction() === OCR_STALE_HANDLER)
    .forEach((trigger) => ScriptApp.deleteTrigger(trigger));
  if (isWanted) {
    ScriptApp.newTrigger(OCR_STALE_HANDLER).timeBased().after(OCR_STALE_GO_ON_MS).create();
  }
}

function rereadOcrDay(bookId, day) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(OCR_LOCK_WAIT_MS)) {
    Logger.log('OCR ажиллаж байна, хэсэг хүлээгээд дахин Run хийнэ үү.');
    return null;
  }
  try {
    return ocrRereadDay(bookId, day);
  } finally {
    lock.releaseLock();
  }
}

function ocrRereadDay(bookId, day) {
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty(OCR_KEY_PROPERTY);
  const root = DriveApp.getFolderById(OCR_ROOT_FOLDER_ID);
  const tab = SpreadsheetApp.openById(bookId).getSheetByName(day);
  if (!tab) {
    throw new Error(`${bookId}: "${day}" tab алга`);
  }
  if (!DRY_RUN) {
    ocrUpgradeTab(tab); // into the current columns first: the new rows are written in them
  }
  const oldRecords = ocrReadRecords(tab);
  const photoIds = [...new Set(oldRecords.map(ocrPhotoIdOf).filter(Boolean))];
  if (DRY_RUN) {
    Logger.log(`${day}: ${photoIds.length} зураг, ${oldRecords.length} мөрийг дахин уншина (туршилт, Gemini дуудахгүй)`);
    return { photos: photoIds.length, read: 0 };
  }
  if (!key) {
    throw new Error('Script properties-д GEMINI_KEY алга (Project Settings > Script properties).');
  }
  const catalog = loadOcrCatalog(ocrStamp(new Date(), 'yyyy-MM-dd'));
  const readAt = ocrStamp(new Date(), 'yyyy-MM-dd HH:mm');
  const fresh = new Set();
  const freshRows = photoIds.flatMap((id) => {
    const rows = ocrRereadPhoto(id, root, key, catalog, readAt);
    if (rows) {
      fresh.add(id);
    }
    return rows || [];
  });
  const merged = ocrMergeReread(oldRecords, freshRows.map(ocrRecord), fresh).map(ocrRow);
  ocrRewriteTab(tab, merged);
  Logger.log(`${day}: ${fresh.size}/${photoIds.length} зураг дахин уншиж ${merged.length} мөр бичлээ (өмнө ${oldRecords.length})`);
  return { photos: photoIds.length, read: fresh.size, rows: merged.length };
}

// One photo's rows read again (ocrRows), or null when it can not be read; its answer goes to the log.
function ocrRereadPhoto(id, root, key, catalog, readAt) {
  const file = DriveApp.getFileById(id);
  const place = ocrLocate(file, root.getId());
  if (!place) {
    Logger.log(`  ${file.getName()}: Баримт дотор олдсонгүй, хуучнаар нь үлдээв`);
    return null;
  }
  const read = readReceiptWithGemini(file.getBlob(), OCR_MIME_TYPES[file.getMimeType()], key, catalog);
  if (!read.ok) {
    Logger.log(`  ${file.getName()}: Gemini уншиж чадсангүй (${read.error}), хуучнаар нь үлдээв`);
    return null;
  }
  Logger.log(`  ${file.getName()}: ${JSON.stringify(read.result)}`);
  // a slip Тосгон sent counts on both sides, as in a first reading (handleOcrPhoto, 2026-10-28)
  const folder = ocrSlipFolder(ocrFolderMeaning(place.path), read.result);
  const folderDate = ocrFolderDate(place.path);
  const check = ocrDocumentCheck(read.result, folder, folderDate, ocrStamp(file.getDateCreated(), 'yyyy-MM-dd'), ocrFolderMonth(place.path));
  ocrMirrorPhoto(file, place.path, ocrSlipSender(read.result), { props: PropertiesService.getScriptProperties(), root });
  return ocrRows(read.result, { name: file.getName(), url: file.getUrl(), id }, folder, check, catalog.set, readAt, catalog.entries);
}

// The tab's rows written again, each photo beside its lines; prices: the receipt price of each row
// ('' for none), kept when given (a line read again has none).
function ocrRewriteTab(tab, rows, prices = []) {
  const last = ocrLastRecordRow(tab);
  if (last > 1) {
    tab.getRange(2, 1, last - 1, OCR_HEADER.length).clearContent(); // not the legend beside them
    if (ocrHasPrices(tab)) {
      tab.getRange(2, OCR_READ_PRICE_COLUMN, last - 1, OCR_PRICE_COLUMNS).clearContent(); // the prices move with their lines
    }
    tab.setRowHeights(2, last - 1, OCR_ROW_HEIGHT);
  }
  tab.getImages().forEach((image) => image.remove());
  ocrGroupByPhoto(rows).reduce(({ first, index }, group) => {
    ocrEnsureRows(tab, first + group.rows.length - 1);
    tab.getRange(first, 1, group.rows.length, OCR_HEADER.length).setValues(group.rows).setVerticalAlignment('top');
    ocrWritePrices(tab, first, group.rows.length, prices.slice(index, index + group.rows.length));
    const next = { index: index + group.rows.length };
    try {
      return { ...next, first: first + ocrAddPhoto(tab, first, group.rows.length, DriveApp.getFileById(group.photoId)) }; // blank rows under a short one
    } catch (error) {
      Logger.log(`  ${group.photoId}: зураг олдсонгүй (${error.message})`);
      return { ...next, first: first + group.rows.length };
    }
  }, { first: 2, index: 0 });
  ocrSetLineChoices(tab); // the lists where the lines are now
  ocrFitColumns(tab);
}
