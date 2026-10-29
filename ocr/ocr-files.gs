/**
 * Receipt OCR - Drive: make the month and day folders ahead, find the photos, make the month
 * spreadsheets (one per place, kind and month) and their day tabs and add rows. Photos are never moved (only their owner may move
 * them). Every
 * function here changes Drive only when called; ocr-main.gs calls them only with DRY_RUN = false.
 */

function ocrChildFolder(parent, name) {
  const folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : null;
}

function ocrEnsurePath(root, names) {
  return names.reduce((folder, name) => ocrChildFolder(folder, name) || folder.createFolder(name), root);
}

// The folder of the month spreadsheets under Баримт: "Зураг уншуулах", or "OCR" while it still has its
// old name; made (by the new name) only when isMade and neither is there, else null.
function ocrSheetsFolder(root, isMade = true) {
  const found = ocrChildFolder(root, OCR_SHEETS_FOLDER) || ocrChildFolder(root, OCR_OLD_SHEETS_FOLDER);
  return found || (isMade ? root.createFolder(OCR_SHEETS_FOLDER) : null);
}

const ocrIterate = (iterator, visit) => {
  while (iterator.hasNext()) {
    visit(iterator.next());
  }
};

// The folders mom puts photos in: every folder directly under Соронзон, My home and Тосгон.
function ocrInboxFolders(root) {
  const inboxes = [];
  OCR_PLACES.forEach((place) => {
    const placeFolder = ocrChildFolder(root, place);
    if (placeFolder) {
      ocrIterate(placeFolder.getFolders(), (folder) => inboxes.push({ folder, path: [place, folder.getName()] }));
    }
  });
  return inboxes;
}

function ocrCountFolders(folder) {
  let count = 0;
  ocrIterate(folder.getFolders(), () => { count += 1; });
  return count;
}

// Months not ready in some inbox (no month folder, or not all its days): [{ inbox, month }].
function ocrMissingMonths(root, months) {
  return ocrInboxFolders(root).flatMap((inbox) =>
    months
      .filter((month) => {
        const monthFolder = ocrChildFolder(inbox.folder, month);
        return !monthFolder || ocrCountFolders(monthFolder) < ocrMonthDays(month).length;
      })
      .map((month) => ({ inbox, month }))
  );
}

// The month folder with a folder for each of its days; what exists already is kept, so a run cut
// off half-way is completed by the next one.
function ocrMakeMonth(inbox, month) {
  const monthFolder = ocrEnsurePath(inbox.folder, [month]);
  ocrMonthDays(month).forEach((day) => ocrEnsurePath(monthFolder, [day]));
}

// The folder a file lies in and the folder names from Баримт down to it, or null when the file is
// not under Баримт/<place>/<folder>. paths: folder id -> its names (or null), kept across files when
// many are looked at (ocrBooksOfLeftPhotos).
function ocrLocate(file, rootId, paths = null) {
  const parents = file.getParents();
  if (!parents.hasNext()) {
    return null;
  }
  const folder = parents.next();
  const id = folder.getId();
  const path = paths && paths.has(id) ? paths.get(id) : ocrFolderPath(folder, rootId);
  if (paths) {
    paths.set(id, path);
  }
  return path ? { folder, path } : null;
}

// The folder names from Баримт down to a folder, or null when it is not under Баримт/<place>/<folder>.
function ocrFolderPath(folder, rootId) {
  const names = [];
  for (let current = folder; current; ) {
    if (current.getId() === rootId) {
      return names.length >= 2 && OCR_PLACES.includes(names[0]) ? names : null;
    }
    names.unshift(current.getName());
    const up = current.getParents();
    current = up.hasNext() ? up.next() : null;
  }
  return null;
}

// The photos not read yet, oldest first. One Drive search for the images changed in the last
// OCR_LOOKBACK_DAYS (walking every day folder every hour would not fit in a run after a few months);
// only a new one is traced up to its folder (normally Соронзон/зарлага/2026-10/01).
function findInboxPhotos(root, isDone) {
  const since = Utilities.formatDate(new Date(Date.now() - OCR_LOOKBACK_DAYS * 24 * 3600 * 1000), 'UTC', "yyyy-MM-dd'T'HH:mm:ss");
  const types = Object.keys(OCR_MIME_TYPES).map((type) => `mimeType = '${type}'`).join(' or ');
  const photos = [];
  ocrIterate(DriveApp.searchFiles(`modifiedDate > '${since}' and trashed = false and (${types})`), (file) => {
    if (isDone(file.getId())) {
      return;
    }
    const place = ocrLocate(file, root.getId());
    if (place) {
      photos.push({ file, ...place });
    }
  });
  return photos.sort((a, b) => a.file.getDateCreated() - b.file.getDateCreated());
}

// The kind of the spreadsheet a photo's lines go to, by the folder mom put it in ([place, folder]).
const ocrBookKind = (path) => OCR_BOOK_KINDS[ocrFolderMeaning(path).kind] || OCR_OTHER_KIND;
const ocrKindOfWord = (word) => [...Object.values(OCR_BOOK_KINDS), OCR_OTHER_KIND].find((kind) => kind.word === word) || null;

const ocrMonthBookName = (place, kind, month) => `${OCR_DAY_SPREADSHEET_PREFIX}${place} ${kind.word} ${month}`;

// "OCR Соронзон орлого 2026-10" -> { place, kind, month }; null for any other name.
function ocrMonthBookOf(name) {
  const match = String(name).match(/^OCR (.+) (\S+) (\d{4}-\d{2})$/);
  const kind = match && ocrKindOfWord(match[2]);
  return kind && OCR_PLACES.includes(match[1]) ? { place: match[1], kind, month: match[3] } : null;
}

// "OCR Соронзон 2026-10", a month spreadsheet of every kind made before 2026-10-25 -> { place, month }.
function ocrOldMonthBookOf(name) {
  const match = String(name).match(/^OCR (.+) (\d{4}-\d{2})$/);
  return match && OCR_PLACES.includes(match[1]) ? { place: match[1], month: match[2] } : null;
}

function ocrIsUnder(file, rootId) {
  for (let parents = file.getParents(); parents.hasNext(); ) {
    let folder = parents.next();
    while (folder) {
      if (folder.getId() === rootId) {
        return true;
      }
      const up = folder.getParents();
      folder = up.hasNext() ? up.next() : null;
    }
  }
  return false;
}

const ocrSheetsQuery = (since) =>
  `title contains '${OCR_DAY_SPREADSHEET_PREFIX.trim()}' and mimeType = '${MimeType.GOOGLE_SHEETS}' and trashed = false` +
  (since ? ` and modifiedDate > '${Utilities.formatDate(since, 'UTC', "yyyy-MM-dd'T'HH:mm:ss")}'` : '');

// The month spreadsheets under Баримт changed after `since`: [{ file, place, kind, month }].
function findOcrBooks(root, since) {
  const found = [];
  ocrIterate(DriveApp.searchFiles(ocrSheetsQuery(since)), (file) => {
    const month = ocrMonthBookOf(file.getName());
    if (month && ocrIsUnder(file, root.getId())) {
      found.push({ file, ...month });
    }
  });
  return found;
}

// The day spreadsheets of the first version that are still there: [{ file, folder, path }].
function findOldDaySheets(root) {
  const found = [];
  ocrIterate(DriveApp.searchFiles(ocrSheetsQuery(null)), (file) => {
    const place = OCR_OLD_DAY_SPREADSHEET.test(file.getName()) ? ocrLocate(file, root.getId()) : null;
    if (place) {
      found.push({ file, ...place });
    }
  });
  return found;
}

// Баримт/Зураг уншуулах/<month>/<Kind>/OCR <place> <kind> <month>, made there if it is not yet: { file, book }.
// Who may change it is set by the review pass of the same run (ocr-sheet.gs).
function ocrMonthBook(root, place, kind, month) {
  const folder = ocrEnsurePath(ocrSheetsFolder(root), [month, kind.folder]);
  const name = ocrMonthBookName(place, kind, month);
  const files = folder.getFilesByName(name);
  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
      return { file, book: SpreadsheetApp.openById(file.getId()) };
    }
  }
  const book = SpreadsheetApp.create(name);
  const file = DriveApp.getFileById(book.getId());
  file.moveTo(folder);
  return { file, book };
}

const ocrDayTabs = (book) => book.getSheets().filter((sheet) => OCR_DAY_FOLDER.test(sheet.getName()));

// The tab of a day ("01"), made and prepared when missing, in the order of the days; the first,
// empty tab of a new spreadsheet becomes it. An existing tab of an older layout is moved into the
// current one first, so new rows go under the right titles.
function ocrDayTab(book, day) {
  const existing = book.getSheetByName(day);
  if (existing) {
    ocrUpgradeTab(existing);
    return existing;
  }
  const sheets = book.getSheets();
  const isBlank = sheets.length === 1 && ocrDayTabs(book).length === 0 && sheets[0].getLastRow() === 0;
  const tab = isBlank ? sheets[0].setName(day) : book.insertSheet(day, ocrDayTabs(book).filter((sheet) => sheet.getName() < day).length);
  ocrPrepareSheet(tab);
  return tab;
}

// Rows go under the last one and under its photo, their prices beside them (ocrWritePrices; readPrices:
// one a row, from ocrReadPrices). Returns the number of the first new row.
function ocrAppendRows(sheet, rows, readPrices = []) {
  const first = ocrNextFreeRow(sheet);
  ocrEnsureRows(sheet, first + rows.length - 1);
  sheet.getRange(first, 1, rows.length, OCR_HEADER.length).setValues(rows).setVerticalAlignment('top');
  ocrWritePrices(sheet, first, rows.length, readPrices);
  ocrSetLineChoices(sheet, first, rows.length);
  ocrFitColumns(sheet);
  return first;
}

// The rows a photo lies over, at the default row height, with the gap before the next one.
const ocrPhotoRows = (image) => Math.ceil((image.getHeight() + OCR_IMAGE_GAP) / OCR_ROW_HEIGHT);

// The first row of the next receipt: below the last line and below the last photo.
function ocrNextFreeRow(sheet) {
  const photoEnds = sheet.getImages()
    .filter((image) => image.getAnchorCell().getColumn() === OCR_IMAGE_COLUMN)
    .map((image) => image.getAnchorCell().getRow() + ocrPhotoRows(image));
  return Math.max(ocrLastRecordRow(sheet) + 1, ...photoEnds);
}

function ocrEnsureRows(sheet, last) {
  if (sheet.getMaxRows() < last) {
    sheet.insertRowsAfter(sheet.getMaxRows(), last - sheet.getMaxRows());
  }
}

// A JPEG of the photo from Drive (null when Drive has none), at most OCR_IMAGE_SIZE px on its longest side.
function ocrPhotoPreview(file) {
  const response = UrlFetchApp.fetch(`https://drive.google.com/thumbnail?id=${file.getId()}&sz=s${OCR_IMAGE_SIZE}`, {
    headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
    muteHttpExceptions: true,
  });
  const isImage = response.getResponseCode() === 200 && /^image\//.test(String(response.getHeaders()['Content-Type'] || ''));
  return isImage ? response.getBlob() : null;
}

// Puts the photo to the right of the receipt's first row; the rows keep their height. The image.
function ocrShowPhoto(sheet, first, blob) {
  const image = sheet.insertImage(blob, OCR_IMAGE_COLUMN, first);
  const height = Math.round((image.getHeight() * OCR_IMAGE_WIDTH) / image.getWidth());
  return image.setWidth(OCR_IMAGE_WIDTH).setHeight(height);
}

function ocrMd5(blob) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, blob.getBytes())
    .map((byte) => ((byte + 256) % 256).toString(16).padStart(2, '0'))
    .join('');
}
