/**
 * Receipt OCR - a slip Тосгон issued to a shop is Тосгон's expense and the shop's income at once (the
 * owner, 2026-10-28: one slip, two sides); one from the city to a shop is Тосгон's income too (three
 * sides). Its lines are written on every side once, from where mom put the photo (ocrSlipFolder,
 * ocr-validate.gs), and the photo is copied to the other sides, so it lies in every place's folder:
 *   Тосгон/Тосгоноос My home руу/2026-10/23/IMG.HEIC  ->  My home/орлого/2026-10/23/IMG.HEIC
 *   My home/орлого/2026-10/23/IMG.HEIC (sent by Тосгон)  ->  Тосгон/Тосгоноос My home руу/2026-10/23/IMG.HEIC
 *   Соронзон/орлого/2026-10/24/IMG.HEIC (from the city)  ->  Тосгон/Тосгоноос Соронзон руу/2026-10/24/IMG.HEIC
 *                                                          and Тосгон/орлого/2026-10/24/IMG.HEIC
 * A copy is marked as read (ocr-done:<copy id>), so it is never read again and nothing is counted twice;
 * a photo is copied once to each place (OCR_MIRROR_PREFIX). The rows read before 2026-10-28 from a
 * Тосгон slip (noted as maybe misfiled, the shop's income only) take the new meaning in the review
 * (ocrTosgonSentRows); those of a city slip, when read again (rereadUnknownSenderDays, ocr-reread.gs).
 */

// The copies of a Тосгон or city slip's photo on its other sides (ocrMirrorPaths; sender: ocrSlipSender).
// A Drive failure is logged: the lines are written all the same.
function ocrMirrorPhoto(file, path, sender, context) {
  return ocrMirrorPaths(path, sender)
    .map((target, index) => ocrMirrorTo(file, target, `${OCR_MIRROR_PREFIX}${file.getId()}${index === 0 ? '' : `#${index + 1}`}`, context))
    .filter(Boolean);
}

// One copy of the photo in the target folder, once (key: where its id is kept); null when none is made.
function ocrMirrorTo(file, target, key, context) {
  if (context.props.getProperty(key)) {
    return null;
  }
  if (DRY_RUN) {
    Logger.log(`  ${file.getName()}: зургийг "${target.join(' / ')}" хавтсанд хуулна (туршилт)`);
    return null;
  }
  try {
    const copy = file.makeCopy(file.getName(), ocrEnsurePath(context.root, ocrMirrorFolderNames(context.root, target)));
    context.props.setProperty(OCR_DONE_PREFIX + copy.getId(), `${OCR_MIRROR_DONE}${file.getId()}`);
    context.props.setProperty(key, copy.getId());
    Logger.log(`  ${file.getName()}: зургийг "${target.join(' / ')}" хавтсанд хуулав`);
    return copy;
  } catch (error) {
    Logger.log(`  ${file.getName()}: зургийг "${target.join(' / ')}" хавтсанд хуулж чадсангүй (${error.message})`);
    return null;
  }
}

// The folder names of a copy's place: Тосгон's folder to a shop as mom named it ("Тосгоноос My home
// руу", whatever its spelling, not a return), else as ocrMirrorPaths has it (Тосгон's "орлого").
function ocrMirrorFolderNames(root, target) {
  const [place, sub, ...rest] = target;
  const meaning = ocrFolderMeaning([place, sub]);
  const placeFolder = place === OCR_TOSGON.place && meaning.kind === 'shiljuuleg' ? ocrChildFolder(root, place) : null;
  if (!placeFolder) {
    return target;
  }
  const shop = meaning.moves[1].place;
  const names = [];
  ocrIterate(placeFolder.getFolders(), (folder) => names.push(folder.getName()));
  const found = names.find((name) => !name.includes(OCR_RETURN_MARK) && ocrFolderMeaning([place, name]).kind === 'shiljuuleg' && ocrFolderMeaning([place, name]).moves[1].place === shop);
  return [place, found || sub, ...rest];
}

// The review of a day tab: rows read before 2026-10-28 from a slip Тосгон sent (ocrTosgonSentRecord)
// take the new label, lose the note and so get Тосгон's side written; their photo is copied once.
function ocrTosgonSentRows(sheet, where, records, context) {
  const photos = new Set();
  return records.map((record, index) => {
    const updated = ocrTosgonSentRecord(record);
    if (!updated) {
      return record;
    }
    Logger.log(`  ${where} мөр ${ocrRowOf(record, index)}: Тосгоноос ирсэн баримт: Тосгоны зарлага, дэлгүүрийн орлого хоёуланд нь бичнэ`);
    const photoId = ocrText(record['ID']).split('#')[0];
    if (!photos.has(photoId)) {
      photos.add(photoId);
      ocrMirrorById(photoId, context);
    }
    if (!DRY_RUN) {
      ['Хавтас (ээж хийсэн)', 'Шалгах шалтгаан', 'Төлөв']
        .filter((title) => updated[title] !== record[title])
        .forEach((title) => sheet.getRange(ocrRowOf(record, index), ocrColumn(title)).setValue(updated[title]));
    }
    return updated;
  });
}

// ocrMirrorPhoto for a photo known by its id, where it lies now (a slip Тосгон sent).
function ocrMirrorById(photoId, context) {
  const place = photoId ? ocrLocateById(photoId, context.root) : null;
  if (place) {
    ocrMirrorPhoto(DriveApp.getFileById(photoId), place.path, OCR_TOSGON.place, context);
  }
}
