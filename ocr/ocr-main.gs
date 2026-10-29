/**
 * Receipt OCR - the run. processNewReceipts makes the month and day folders ahead, reads every new
 * photo once and writes its rows into the day's tab of its place's month spreadsheet
 * (ocr-config.gs), then goes over the month spreadsheets that changed: protects them, writes the lines that may be written into
 * the workbook (ocr-post.gs), renews mom's report (ocr-report.gs) and the balances of the count
 * blocks (ocr-spot.gs).
 * installOcrTrigger runs it every hour.
 */

const OCR_LOCK_WAIT_MS = 5000;
const OCR_HANDLER = 'processNewReceipts';
const OCR_FOLDERS_CHECKED_PROPERTY = 'ocr-folders-checked'; // the day ('yyyy-MM-dd') the folders were last looked at

function processNewReceipts() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(OCR_LOCK_WAIT_MS)) {
    Logger.log('OCR: өмнөх ажил дуусаагүй байна, энэ удаад алгаслаа.');
    return null;
  }
  try {
    return runOcrPass(Date.now() + OCR_TIME_BUDGET_MS);
  } finally {
    lock.releaseLock();
  }
}

const ocrStamp = (date, format) => Utilities.formatDate(date, OCR_TIME_ZONE, format);

// lap() -> the whole seconds since the last lap (the first: since the start), for the logs of a pass.
function ocrStopwatch() {
  let last = Date.now();
  return () => {
    const now = Date.now();
    const seconds = Math.round((now - last) / 1000);
    last = now;
    return seconds;
  };
}

// This month's and the next month's day folders in every inbox.
function prepareOcrFolders(root, today) {
  const missing = ocrMissingMonths(root, ocrUpcomingMonths(today));
  missing.forEach(({ inbox, month }) => {
    Logger.log(`${inbox.path.join(' / ')} / ${month}: өдрийн хавтсууд ${DRY_RUN ? 'үүснэ' : 'үүсгэв'}`);
    if (!DRY_RUN) {
      ocrMakeMonth(inbox, month);
    }
  });
  return missing.length;
}

function runOcrPass(deadline) {
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty(OCR_KEY_PROPERTY);
  if (!key && !DRY_RUN) {
    throw new Error('Script properties-д GEMINI_KEY алга (Project Settings > Script properties).');
  }
  const now = new Date();
  const lap = ocrStopwatch();
  const today = ocrStamp(now, 'yyyy-MM-dd');
  const root = DriveApp.getFolderById(OCR_ROOT_FOLDER_ID);
  if (!ocrPrepareShopLayout(props, root) || !ocrPrepareGiftColumn(props, root) || !ocrPrepareTosgonShelf(props, root)) {
    return null; // the shops' or Тосгон's workbook still has the old columns (ocr-layout.gs)
  }
  // the folders are made a month ahead, so once a day is enough (looking took 13 s a pass, 2026-10-29)
  const isFoldersChecked = props.getProperty(OCR_FOLDERS_CHECKED_PROPERTY) === today;
  const months = isFoldersChecked ? 0 : prepareOcrFolders(root, today);
  if (!isFoldersChecked && !DRY_RUN) {
    props.setProperty(OCR_FOLDERS_CHECKED_PROPERTY, today);
  }
  const folderSeconds = lap();
  const catalog = loadOcrCatalog(today);
  const catalogSeconds = lap();
  const photos = findInboxPhotos(root, (id) => props.getProperty(OCR_DONE_PREFIX + id) !== null);
  Logger.log(`OCR бэлтгэл: хавтас ${folderSeconds} сек, барааны жагсаалт ${catalogSeconds} сек, зураг хайх ${lap()} сек.`);
  // touched: the month spreadsheets written in this run, looked at by the review pass at once
  const context = { props, key, catalog, today, root, touched: new Map(), readAt: ocrStamp(now, 'yyyy-MM-dd HH:mm') };
  const counts = { read: 0, duplicate: 0, failed: 0, waiting: 0, left: 0 };
  for (let index = 0; index < photos.length; index++) {
    if (Date.now() > deadline - OCR_REPORT_RESERVE_MS) {
      counts.left = photos.length - index;
      break;
    }
    counts[handleOcrPhoto(photos[index], context)] += 1;
  }
  // the seconds since the start: a pass with no new photo took 39 s to here at 4:42 and 210 s at 6:42 (2026-10-29)
  const seconds = () => Math.round((Date.now() - now.getTime()) / 1000);
  Logger.log(
    `OCR (${modeLabel()}): сар бэлдсэн ${months}, барааны нэр ${catalog.names.length}, шинэ зураг ${photos.length}; уншсан ${counts.read}, ` +
      `давхардсан ${counts.duplicate}, уншиж чадаагүй ${counts.failed}, уншаагүй (туршилт) ${counts.waiting}, дараагийн удаад ${counts.left} (${seconds()} сек).`
  );
  runReviewPass(context, deadline);
  // out of time, the count report waits for the next pass: Google stops a run at 6 minutes, and two
  // passes were stopped there (2026-10-29), before they kept where they had got to
  if (Date.now() > deadline) {
    Logger.log(`Тайлан: цаг дууссан тул дараагийн ажлаар шинэчлэгдэнэ (${seconds()} сек).`);
    return counts;
  }
  writeOcrOutputs(context);
  const blocks = ocrRefreshSpotBalances(deadline);
  if (blocks > 0) {
    Logger.log(`Дундын тооллого: ${blocks} блокийн үлдэгдэл ${DRY_RUN ? 'тооцоолно' : 'шинэчлэв'}.`);
  }
  return counts;
}

// 'read' | 'duplicate' | 'failed' | 'waiting' (DRY_RUN: would be read)
function handleOcrPhoto(photo, context) {
  const { file } = photo;
  const folder = ocrFolderMeaning(photo.path);
  const folderDate = ocrFolderDate(photo.path);
  const info = { name: file.getName(), url: file.getUrl(), id: file.getId() };
  const uploadedOn = ocrStamp(file.getDateCreated(), 'yyyy-MM-dd');
  const blob = file.getBlob();
  const md5 = ocrMd5(blob);
  const seen = context.props.getProperty(OCR_SEEN_PREFIX + md5);
  if (seen) {
    const rows = [ocrNoteRow(info, folder, `${OCR_DUPLICATE_REASON}: ${seen}`, context.readAt)];
    return fileOcrRows(photo, context, { md5, date: folderDate || uploadedOn, outcome: 'duplicate', rows });
  }
  if (DRY_RUN) {
    Logger.log(`${info.name} (${photo.path.join(' / ')}) -> Gemini-ээр уншина`);
    return 'waiting';
  }
  const read = ocrReadOnce(blob, file.getMimeType(), md5, context);
  if (!read.ok) {
    const rows = [ocrNoteRow(info, folder, `Gemini уншиж чадсангүй: ${read.error}`, context.readAt)];
    return fileOcrRows(photo, context, { md5, date: folderDate || uploadedOn, outcome: 'failed', rows });
  }
  if (ocrIsSkippedSlip(read.result, folder)) {
    const rows = [ocrNoteRow(info, folder, OCR_SKIPPED_SLIP_TEXT, context.readAt)];
    return fileOcrRows(photo, context, { md5, date: folderDate || uploadedOn, outcome: 'read', rows });
  }
  const slipFolder = ocrSlipFolder(folder, read.result); // a slip Тосгон sent: both sides (2026-10-28)
  const check = ocrDocumentCheck(read.result, slipFolder, folderDate, uploadedOn, ocrFolderMonth(photo.path));
  const rows = ocrRows(read.result, info, slipFolder, check, context.catalog.set, context.readAt, context.catalog.entries);
  const outcome = fileOcrRows(photo, context, { md5, date: check.date, outcome: 'read', rows, prices: ocrReadPrices(read.result) });
  ocrMirrorPhoto(file, photo.path, ocrSlipSender(read.result), context);
  return outcome;
}

// Gemini's answer for this photo: the kept one when a run was cut off after asking, else the
// one and only call. A wrong key stops the run before the photo counts as read.
function ocrReadOnce(blob, mimeType, md5, context) {
  const kept = context.props.getProperty(OCR_RESULT_PREFIX + md5);
  if (kept) {
    return { ok: true, result: JSON.parse(kept) };
  }
  const read = readReceiptWithGemini(blob, OCR_MIME_TYPES[mimeType], context.key, context.catalog);
  if (read.fatal) {
    throw new Error(`Gemini: ${read.error}`);
  }
  if (read.ok) {
    try {
      context.props.setProperty(OCR_RESULT_PREFIX + md5, JSON.stringify(read.result));
    } catch (error) {
      Logger.log(`хариуг түр хадгалж чадсангүй (${error.message}); шууд бичнэ`);
    }
  }
  return read;
}

function ocrRowSummary(rows) {
  const status = OCR_HEADER.indexOf('Төлөв');
  const ok = rows.filter((row) => row[status] === OCR_STATUS_OK).length;
  return `${rows.length} мөр (${ok} OK, ${rows.length - ok} шалгах)`;
}

// plan: { md5, date ('YYYY-MM-DD', its day), outcome, rows }. The rows go into that day's tab of
// the month spreadsheet of the photo's place and kind (ocrMonthBook); the photo stays where it is.
function fileOcrRows(photo, context, plan) {
  const place = photo.path[0];
  const kind = ocrBookKind(photo.path);
  const month = ocrMonthFolderName(plan.date);
  const day = plan.date.slice(8);
  const sheetName = `${ocrMonthBookName(place, kind, month)} / ${day}`;
  const where = [...photo.path, photo.file.getName()].join('/');
  Logger.log(`${where} -> "${sheetName}": ${ocrRowSummary(plan.rows)}`);
  if (DRY_RUN) {
    return plan.outcome;
  }
  const { book, file } = ocrMonthBook(context.root, place, kind, month);
  const tab = ocrDayTab(book, day);
  const rows = plan.outcome === 'read' ? ocrMarkTwin(book, plan.rows, day, context) : plan.rows;
  const first = ocrAppendRows(tab, rows, plan.prices);
  ocrAddPhoto(tab, first, rows.length, photo.file);
  context.touched.set(book.getId(), { file, place, kind, month });
  const status = OCR_HEADER.indexOf('Төлөв');
  ocrAddDayStats(context, { photos: 1, rows: plan.rows.length, ok: plan.rows.filter((row) => row[status] === OCR_STATUS_OK).length });
  if (plan.outcome !== 'duplicate') {
    context.props.setProperty(OCR_SEEN_PREFIX + plan.md5, where);
  }
  context.props.setProperty(OCR_DONE_PREFIX + photo.file.getId(), `${OCR_SHEETS_FOLDER}/${month}/${kind.folder}/${sheetName}`);
  context.props.deleteProperty(OCR_RESULT_PREFIX + plan.md5);
  return plan.outcome;
}

// The photo next to its rows; when Drive can not give it, the rows still stand (with the link).
// Returns the rows the receipt takes: its lines, or more when its photo is taller.
function ocrAddPhoto(sheet, first, count, file) {
  try {
    const blob = ocrPhotoPreview(file);
    if (blob) {
      return Math.max(count, ocrPhotoRows(ocrShowPhoto(sheet, first, blob)));
    }
    Logger.log(`  ${file.getName()}: зургийг sheet-д оруулж чадсангүй (Drive JPEG өгсөнгүй), холбоосоор нээнэ`);
  } catch (error) {
    Logger.log(`  ${file.getName()}: зургийг sheet-д оруулж чадсангүй (${error.message}), холбоосоор нээнэ`);
  }
  return count;
}

function installOcrTrigger() {
  const existing = ScriptApp.getProjectTriggers().filter((trigger) => trigger.getHandlerFunction() === OCR_HANDLER);
  if (DRY_RUN) {
    Logger.log(`${modeLabel()}: ${OCR_HANDLER} ${OCR_TRIGGER_HOURS} цаг тутам ажиллахаар тохируулна (одоо ${existing.length} trigger).`);
    return;
  }
  existing.forEach((trigger) => ScriptApp.deleteTrigger(trigger));
  ScriptApp.newTrigger(OCR_HANDLER).timeBased().everyHours(OCR_TRIGGER_HOURS).create();
  installCountReportTrigger();
  Logger.log(`OCR: ${OCR_HANDLER} ${OCR_TRIGGER_HOURS} цаг тутам ажиллана.`);
}
