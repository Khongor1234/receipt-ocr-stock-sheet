/**
 * Receipt OCR - the counts made between the period ends (the owner, 2026-10-24). From the 10/31
 * sheets on (takesSpotCounts, config.gs) a count of another day than the sheet's last (Тосгон's
 * 10/23) goes into a block of its own at the right edge of its sheet, after the five totals
 * (spotFirstColumn), three columns a counted day:
 *   "Тооллого 10/23"  the counted number: the day's count lines, written as any other line
 *                     (ocrWriteRecord, with the usual note; the pages of a day add up)
 *   "Үлдэгдэл 10/23"  the books at the end of that day (ocrBalanceAsOf), worked out again every run
 *   "Зөрүү 10/23"     the count minus that balance, one formula in its header cell
 * The block is made the first time its day is written: its columns are protected so that only the
 * owner and the shop account (OCR_SPOT_EDITOR) change them, and A1's list of the sheet gets
 * "Тооллого 10/23" (viewListRange, config.gs; view-columns.gs shows the days up to it).
 * The count column itself, and so the next sheet's opening balance, is never touched.
 */

const OCR_SPOT_EDITOR = 'shop-account@example.com';
const OCR_SPOT_GUARD = `${OCR_GUARD_PREFIX}: дундын тооллого`;
const OCR_SPOT_NUMBER_FORMAT = '0';

// ---------- pure helpers (no Spreadsheet calls) ----------

// Where the block of a day is among the headers of row 3 from the first block column on: its
// offset, found or the first free block (blocks are SPOT_BLOCK_WIDTH apart). -> { offset, isNew }
function ocrSpotOffset(headers, month, day) {
  const found = headers.findIndex((text) => {
    const spot = spotDayOf(text);
    return spot !== null && spot.month === month && spot.day === day;
  });
  if (found >= 0) {
    return { offset: found, isNew: false };
  }
  const isFree = (offset) => headers.slice(offset, offset + SPOT_BLOCK_WIDTH).every((text) => String(text).trim() === '');
  let offset = 0;
  while (!isFree(offset)) {
    offset += SPOT_BLOCK_WIDTH;
  }
  return { offset, isNew: true };
}

// The "Зөрүү" column's formula: count - balance on the product rows (B..D hold a number, as the
// sheet's own balance formula), blank where nothing was counted.
function ocrSpotDiffFormula(header, countColumn) {
  const count = columnLetter(countColumn);
  const balance = columnLetter(countColumn + 1);
  return `=ARRAYFORMULA({"${header}";IF(BYROW(B4:D,LAMBDA(r,COUNT(r)>0)),IF(LEN(${count}4:${count}),${count}4:${count}-${balance}4:${balance},""),"")})`;
}

// 2026, 10, 23 -> "2026-10-23"
const ocrIsoOf = (year, month, day) => `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;

// "Sor/2026/10/31" -> "2026-10-31", "Tos/2026/10" -> "2026-10-31"; '' for any other tab.
function ocrTabPeriodEnd(name, place) {
  const parts = String(name).split('/').slice(1).map(Number);
  if (parts.some((part) => !Number.isInteger(part) || part <= 0)) {
    return '';
  }
  const [year, month, day] = parts;
  if (place === OCR_TOSGON.place) {
    return parts.length === 2 && month <= 12 ? ocrIsoOf(year, month, daysInMonth(year, month)) : '';
  }
  return parts.length === 3 && month <= 12 ? ocrIsoOf(year, month, day) : '';
}

// ---------- spreadsheet access ----------

function ocrSpotHeaders(sheet, first) {
  const width = sheet.getMaxColumns() - first + 1;
  return width > 0 ? sheet.getRange(LAYOUT_HEADER_ROW, first, 1, width).getValues()[0] : [];
}

// The "Тооллого" column of the block of `iso` on a workbook sheet, made when missing (DRY_RUN:
// where it would be made).
function ocrSpotBlock(sheet, bookId, layout, iso) {
  const [, month, day] = iso.split('-').map(Number);
  const first = spotFirstColumn(layout);
  const { offset, isNew } = ocrSpotOffset(ocrSpotHeaders(sheet, first), month, day);
  if (isNew && !DRY_RUN) {
    ocrMakeSpotBlock(sheet, bookId, layout, first + offset, month, day);
  }
  return first + offset;
}

function ocrMakeSpotBlock(sheet, bookId, layout, column, month, day) {
  const missing = column + SPOT_BLOCK_WIDTH - 1 - sheet.getMaxColumns();
  if (missing > 0) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), missing);
  }
  // new columns take the format of the totals beside them (₮)
  sheet.getRange(OCR_BOOK_FIRST_ROW, column, sheet.getMaxRows() - OCR_BOOK_FIRST_ROW + 1, SPOT_BLOCK_WIDTH).setNumberFormat(OCR_SPOT_NUMBER_FORMAT);
  const [countHeader, balanceHeader, diffHeader] = spotHeaders(month, day);
  sheet.getRange(LAYOUT_HEADER_ROW, column, 1, 2).setValues([[countHeader, balanceHeader]]);
  sheet.getRange(LAYOUT_HEADER_ROW, column + 2).setFormula(ocrSpotDiffFormula(diffHeader, column));
  const protection = sheet.getRange(1, column, sheet.getMaxRows(), SPOT_BLOCK_WIDTH).protect().setDescription(OCR_SPOT_GUARD);
  ocrOnlyOwner(protection, Session.getEffectiveUser().getEmail());
  protection.addEditor(OCR_SPOT_EDITOR);
  const book = SpreadsheetApp.openById(bookId);
  setViewListRule(sheet, viewListRange(book, sheet.getName(), viewBaseChoices(layout), spotChoice(month, day)));
  Logger.log(`${sheet.getName()}: ${spotChoice(month, day)}-ны блок ${columnLetter(column)}-д үүслээ`);
}

// Every run: the "Үлдэгдэл" column of every block on the sheets that take them, from the books as
// they stand now. -> the number of blocks looked at
function ocrRefreshSpotBalances(deadline) {
  return ocrSpotSheets().reduce((total, entry) => (Date.now() > deadline ? total : total + ocrRefreshSheetBalances(entry)), 0);
}

// The workbook sheets that may hold blocks: { sheet, place, layout, year }, those whose period
// ends on SPOT_COUNTS_FROM or later, in the places' workbooks and this year's (year-books.gs).
function ocrSpotSheets() {
  const year = new Date().getFullYear();
  const ids = [...new Set(Object.keys(OCR_POST_BOOKS).flatMap((place) => [OCR_POST_BOOKS[place], ocrBookIdFor(place, `-/${year}`)]))].filter(Boolean);
  return ids.flatMap((id) => {
    let book;
    try {
      book = SpreadsheetApp.openById(id);
    } catch (error) {
      Logger.log(`${id}: нээгдсэнгүй (${error.message})`);
      return [];
    }
    return book.getSheets()
      .map((sheet) => ({ sheet, place: ocrPlaceOfTab(sheet.getName()) }))
      .filter(({ sheet, place }) => place && ocrBookIdFor(place, sheet.getName()) === id)
      .map((entry) => ({ ...entry, end: ocrTabPeriodEnd(entry.sheet.getName(), entry.place) }))
      .filter(({ end }) => end !== '' && takesSpotCounts(end))
      .map(({ sheet, place, end }) => ({ sheet, place, layout: place === OCR_TOSGON.place ? OCR_TOSGON.layout : LAYOUT, year: Number(end.slice(0, 4)) }));
  });
}

function ocrRefreshSheetBalances({ sheet, place, layout, year }) {
  const first = spotFirstColumn(layout);
  const blocks = ocrSpotHeaders(sheet, first)
    .map((text, index) => ({ column: first + index, spot: spotDayOf(text) }))
    .filter((block) => block.spot !== null);
  const height = sheet.getLastRow() - OCR_BOOK_FIRST_ROW + 1;
  if (blocks.length === 0 || height <= 0) {
    return 0;
  }
  const rows = sheet.getRange(OCR_BOOK_FIRST_ROW, 1, height, layout.expenseLast).getValues();
  const notesOf = (column) => sheet.getRange(OCR_BOOK_FIRST_ROW, column, height, 1).getNotes().map(([note]) => note);
  const [notesIn, notesOut] = [notesOf(layout.transferIn), notesOf(layout.transferOut)];
  blocks.forEach(({ column, spot }) => {
    const iso = ocrIsoOf(year, spot.month, spot.day);
    const target = ocrTarget({ place, side: 'in' }, iso);
    if (!target || target.tab !== sheet.getName()) {
      return; // a header of a day outside this sheet: left as it is
    }
    const offset = target.column - layout.incomeFirst;
    const wanted = rows.map((row, index) => [
      [1, 2, 3].some((at) => typeof row[at] === 'number')
        ? ocrBalanceAsOf(row, { in: notesIn[index], out: notesOut[index] }, layout, offset, iso)
        : '',
    ]);
    const range = sheet.getRange(OCR_BOOK_FIRST_ROW, column + 1, height, 1);
    if (!DRY_RUN && JSON.stringify(range.getValues()) !== JSON.stringify(wanted)) {
      range.setValues(wanted);
    }
  });
  return blocks.length;
}
