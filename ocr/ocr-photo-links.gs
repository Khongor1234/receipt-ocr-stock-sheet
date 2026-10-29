/**
 * Receipt OCR - the link in a workbook cell's note (the owner, 2026-10-23). Since 2026-10-26 each
 * line of the note links to the line's row in its month spreadsheet, with the line's count after it
 * (ocrNoteLine): "https://docs.google.com/spreadsheets/d/<book>/edit#gid=0&range=A30:M30&ocr=<file id>.13 = 5 (2026-10-23)".
 * A note written before links to the photo ("Зураг: https://drive.google.com/file/d/<file id>/view#13 = 5 ...",
 * 2026-10-25) or has an "OCR: ..." line and a bare link line under it; the note's reader
 * (ocrNoteEntries) reads all of them and skips the bare links.
 * addSheetLinksToOcrNotes rewrites the notes written before in the new way once (run by hand from
 * the editor; DRY_RUN only counts): the row of each line is found by its ID in the month
 * spreadsheets. A run cut off by the time goes on where it stopped when run again. A note with a
 * line of its own (mom's, ocrHandNote's) is left as it is, and so are OCR_HANDS_OFF_CELLS; a line
 * whose row is not found keeps the photo's link.
 */

const OCR_PHOTO_LINK_PREFIX = 'Зураг: https://drive.google.com/file/d/';
const OCR_LINKED_TAB = /^(Sor|My|Tos)\/\d{4}\//;

// the bare link line under an "OCR: ..." line of a note written before 2026-10-25
const ocrIsPhotoLinkLine = (line) => /^Зураг: https:\/\/drive\.google\.com\/file\/d\/[^/\s]+\/view$/.test(line);

const OCR_NOTES_DONE_PROPERTY = 'ocr-notes-sheet-linked'; // the tabs done ("<file id>|<tab>"), while not all are
const OCR_NOTES_BUDGET_MS = 4.5 * 60 * 1000; // Google stops a run at 6 minutes: run it again, it goes on

function addSheetLinksToOcrNotes() {
  const props = PropertiesService.getScriptProperties();
  const deadline = Date.now() + OCR_NOTES_BUDGET_MS;
  const books = [...new Set(Object.values(OCR_POST_BOOKS))];
  const copies = books.map((id) => props.getProperty(OCR_CHECK_COPY_PROPERTY + id)).filter(Boolean);
  const done = new Set(JSON.parse(props.getProperty(OCR_NOTES_DONE_PROPERTY) || '[]'));
  const rows = ocrLineRows();
  Logger.log(`Сарын spreadsheet-үүдээс ${rows.size} мөрийн байрлал олдлоо`);
  let changed = 0;
  for (const id of [...books, ...copies]) {
    const result = ocrLinkBookNotes(id, done, deadline, rows);
    changed += result.changed;
    if (!DRY_RUN) {
      props.setProperty(OCR_NOTES_DONE_PROPERTY, JSON.stringify([...done]));
    }
    if (!result.finished) {
      Logger.log(`${modeLabel()}: тэмдэглэл энэ удаа ${changed} нүд ${DRY_RUN ? 'олдлоо' : 'шинэчлэгдлээ'}. Цаг дууссан тул дуусаагүй: дахин ажиллуулахад үлдсэнээсээ үргэлжилнэ.`);
      return;
    }
  }
  props.deleteProperty(OCR_NOTES_DONE_PROPERTY);
  Logger.log(`${modeLabel()}: тэмдэглэл ${changed} нүд ${DRY_RUN ? 'олдлоо' : 'spreadsheet-ийн мөрийн холбоостой боллоо'}. Бүх tab дууслаа.`);
}

// The row of every line of the month spreadsheets: ID -> { book, gid, row }.
function ocrLineRows() {
  const rows = new Map();
  findOcrBooks(DriveApp.getFolderById(OCR_ROOT_FOLDER_ID), null).forEach((entry) => {
    const book = SpreadsheetApp.openById(entry.file.getId());
    ocrDayTabs(book).forEach((sheet) => {
      const last = ocrLastRecordRow(sheet);
      if (last < 2) {
        return;
      }
      sheet.getRange(2, ocrColumn('ID'), last - 1, 1).getValues().forEach(([id], index) => {
        if (ocrText(id) !== '' && !rows.has(ocrText(id))) {
          rows.set(ocrText(id), { book: entry.file.getId(), gid: sheet.getSheetId(), row: index + 2 });
        }
      });
    });
  });
  return rows;
}

// The notes of one workbook's tabs not done yet, until the deadline: { changed, finished }.
function ocrLinkBookNotes(bookId, done, deadline, rows) {
  let book;
  try {
    book = SpreadsheetApp.openById(bookId);
  } catch (error) {
    Logger.log(`${bookId}: нээгдсэнгүй (${error.message})`);
    return { changed: 0, finished: true };
  }
  let changed = 0;
  for (const sheet of book.getSheets().filter((tab) => OCR_LINKED_TAB.test(tab.getName()))) {
    const key = `${bookId}|${sheet.getName()}`;
    if (done.has(key)) {
      continue;
    }
    if (Date.now() > deadline) {
      return { changed, finished: false };
    }
    changed += ocrLinkSheetNotes(sheet, rows);
    done.add(key);
  }
  return { changed, finished: true };
}

function ocrLinkSheetNotes(sheet, rows) {
  const count = sheet.getLastRow() - OCR_BOOK_FIRST_ROW + 1;
  const columns = sheet.getLastColumn();
  if (count < 1 || columns < 1) {
    return 0;
  }
  const range = sheet.getRange(OCR_BOOK_FIRST_ROW, 1, count, columns);
  const notes = range.getNotes();
  let changed = 0;
  const linked = notes.map((row, r) => row.map((note, c) => {
    const cell = `${sheet.getName()}!${columnLetter(c + 1)}${r + OCR_BOOK_FIRST_ROW}`;
    const entries = OCR_HANDS_OFF_CELLS.includes(cell) ? null : ocrNoteEntries(note);
    const wanted = entries && entries.length > 0
      ? entries.map((entry) => ocrNoteLine({ ...entry, at: rows.get(entry.id) || entry.at })).join('\n')
      : note;
    if (wanted !== note) {
      changed++;
    }
    return wanted;
  }));
  if (changed > 0 && !DRY_RUN) {
    range.setNotes(linked);
  }
  return changed;
}
