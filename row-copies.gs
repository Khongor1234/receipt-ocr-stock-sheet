/**
 * A product row put in under its category takes its formats from a product row of that category
 * (year-sync.gs insertProduct, price-link.gs insertTosgonProduct, tosgon-link.gs insertShopProduct).
 * Until 2026-10-26 it took that row's notes and the colours of its written cells with them, and only
 * the numbers were cleared: the year's sheets had 253 empty cells holding another product's note
 * (the owner saw "Дэлгүүрийн орлогоос: Соронзон 2" on 22 empty cells of Tos/2026/10!AP, every one
 * copied from AP56, the last row of the category), 28 of them pink or yellow like a written cell.
 *   clearCopiedMarks: a new row loses the notes it got that way, and the script's colours (the
 *     OCR's and the transfer link's backgrounds, the blue and purple of the numbers they write) go
 *     back to the column's.
 *   clearCopiedNotes(): run by hand once: in every month tab of this workbook and of Тосгон's, an
 *     empty cell whose note is the very note of a cell with a number in the same column is such a
 *     copy, and is put right the same way. A note on an empty cell no other cell has (a number
 *     emptied by hand) stays and is listed. DRY_RUN only lists.
 * Needs dividers.gs (DIVIDER_KEPT_COLORS, findColumnColors, withFilterLifted).
 */

const ROW_COPY_FIRST_ROW = 4; // below the three header rows
// the colours of the numbers the receipt OCR writes (OCR_POSTED_FONT_COLOR, ocr/ocr-config.gs) and
// of Тосгон's expense it filled from the shops' income (OCR_BACKFILL_FONT_COLOR, ocr/ocr-backfill.gs)
const ROW_COPY_SCRIPT_FONTS = ['#1155cc', '#8e24aa'];
const ROW_COPY_FONT = '#000000';
const ROW_COPY_TAB = /^[^/]+\/\d{4}\/\d{1,2}(\/\d{1,2})?$/; // "Sor/2026/10/31", "Tos/2026/10"
const ROW_COPY_LIST_LIMIT = 20;

const isScriptBackground = (color) => DIVIDER_KEPT_COLORS.includes(String(color).toLowerCase());
const isScriptFont = (color) => ROW_COPY_SCRIPT_FONTS.includes(String(color).toLowerCase());

// The usual background of each of a sheet's first `width` columns on its rows (dividers.gs).
function rowCopyColumnColors(sheet, width) {
  const height = sheet.getLastRow() - ROW_COPY_FIRST_ROW + 1;
  if (height <= 0) {
    return Array.from({ length: width }, () => DIVIDER_WHITE);
  }
  const backgrounds = sheet.getRange(ROW_COPY_FIRST_ROW, 1, height, width).getBackgrounds();
  return findColumnColors(backgrounds, backgrounds.map((_, index) => index));
}

// After a row's formats were copied from a template row: its notes and the script's colours go.
function clearCopiedMarks(sheet, row, width) {
  const range = sheet.getRange(row, 1, 1, width);
  range.clearNote();
  const [backgrounds] = range.getBackgrounds();
  const [fonts] = range.getFontColors();
  if (backgrounds.some(isScriptBackground)) {
    const usual = rowCopyColumnColors(sheet, width);
    range.setBackgrounds([backgrounds.map((color, index) => (isScriptBackground(color) ? usual[index] : color))]);
  }
  if (fonts.some(isScriptFont)) {
    range.setFontColors([fonts.map((color) => (isScriptFont(color) ? ROW_COPY_FONT : color))]);
  }
}

// The cells of one tab's grid (from ROW_COPY_FIRST_ROW) that hold a copied note: { copies: [{ row,
// column }], own: [{ row, column, note }] }, own being the empty cells with a note no other cell has.
function findCopiedNotes(values, notes) {
  const isEmpty = (value) => value === '' || value === null;
  const held = new Set();
  values.forEach((row, r) => row.forEach((value, c) => {
    if (!isEmpty(value) && notes[r][c] !== '') {
      held.add(`${c}|${notes[r][c]}`);
    }
  }));
  const copies = [];
  const own = [];
  values.forEach((row, r) => row.forEach((value, c) => {
    const note = notes[r][c];
    if (!isEmpty(value) || note === '') {
      return;
    }
    const cell = { row: r + ROW_COPY_FIRST_ROW, column: c + 1 };
    if (held.has(`${c}|${note}`)) {
      copies.push(cell);
    } else {
      own.push({ ...cell, note });
    }
  }));
  return { copies, own };
}

// One tab: its copies put right (not in DRY_RUN). -> { copies: ["AP16", ...], own: ["I182 (note)", ...] }
function clearSheetCopies(sheet) {
  const height = sheet.getLastRow() - ROW_COPY_FIRST_ROW + 1;
  const width = sheet.getLastColumn();
  if (height <= 0 || width <= 0) {
    return { copies: [], own: [] };
  }
  const range = sheet.getRange(ROW_COPY_FIRST_ROW, 1, height, width);
  const notes = range.getNotes();
  const { copies, own } = findCopiedNotes(range.getValues(), notes);
  const a1 = ({ row, column }) => `${columnLetter(column)}${row}`;
  if (copies.length > 0 && !DRY_RUN) {
    const backgrounds = range.getBackgrounds();
    const fonts = range.getFontColors();
    const usual = findColumnColors(backgrounds, backgrounds.map((_, index) => index));
    const isCopy = new Set(copies.map(({ row, column }) => `${row - ROW_COPY_FIRST_ROW}|${column - 1}`));
    const fix = (grid, change) => grid.map((row, r) => row.map((value, c) => (isCopy.has(`${r}|${c}`) ? change(value, c) : value)));
    withFilterLifted(sheet, () => {
      range.setNotes(fix(notes, () => ''));
      range.setBackgrounds(fix(backgrounds, (color, c) => (isScriptBackground(color) ? usual[c] : color)));
      range.setFontColors(fix(fonts, (color) => (isScriptFont(color) ? ROW_COPY_FONT : color)));
    });
  }
  return { copies: copies.map(a1), own: own.map((cell) => `${a1(cell)} (${cell.note.split('\n')[0].slice(0, 60)})`) };
}

function clearCopiedNotes() {
  withScriptLock(() => Logger.log('%s\n%s', modeLabel(),
    runClearCopiedNotes([currentBook(), partnerBook(currentBook(), 'tosgon', TOSGON_BOOK_ID)]).join('\n')));
}

function runClearCopiedNotes(books) {
  const lines = [];
  let total = 0;
  const unique = books.filter((book, index) => book && books.findIndex((other) => other && other.getId() === book.getId()) === index);
  unique.forEach((book) => book.getSheets()
    .filter((sheet) => ROW_COPY_TAB.test(sheet.getName()))
    .forEach((sheet) => {
      const { copies, own } = clearSheetCopies(sheet);
      total += copies.length;
      if (copies.length > 0) {
        const shown = copies.slice(0, ROW_COPY_LIST_LIMIT).join(' ') + (copies.length > ROW_COPY_LIST_LIMIT ? ' …' : '');
        lines.push(`${sheet.getName()}: хуулагдсан тэмдэглэл ${copies.length} ${DRY_RUN ? 'арилгана' : 'арилгасан'}: ${shown}`);
      }
      if (own.length > 0) {
        lines.push(`${sheet.getName()}: хоосон нүдэнд өөрийн тэмдэглэлтэй ${own.length} (хөндөөгүй): ${own.join('; ')}`);
      }
    }));
  return [`Хуулагдсан тэмдэглэл нийт ${total} нүдэнд`, ...lines];
}
