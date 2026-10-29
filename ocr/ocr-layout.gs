/**
 * Receipt OCR - the shops' workbook moving to the paired layout (its layout-migration.gs: a shelf
 * price after D, the other shop's count after the old "sold" column, so every day column is one
 * further right). The OCR counts the shops' columns from LAYOUT (config.gs), already the new way,
 * so at the start of every run (ocrPrepareShopLayout):
 *   - while the shops' workbook still has the old layout the run stops at once: a line would go
 *     one column off
 *   - once it has moved, every shop cell named in "Бичсэн" ("Sor/2026/10/15!U22", also in a
 *     "check:" cell) moves with its column, and the check copy of the workbook (ocr-check.gs)
 *     gets the same two columns. Done once: Script property OCR_PAIRED_DONE_PROPERTY.
 * Тосгон's cells ("Tos/2026/10!AJ57") stay as they are: its workbook keeps its layout.
 *
 * The gift column (config.gs GIFT_COLUMN, the workbook's gift-migration.gs) is the same once
 * more (ocrPrepareGiftColumn): until the shops' workbook has it the run stops; then the
 * shop cells after the transfer out (the count) move one column right, in "Бичсэн" and in the
 * check copy. Done once: Script property OCR_GIFT_DONE_PROPERTY.
 *
 * Тосгон gets the shelf price in E too (2026-10-22, OCR_TOSGON_SHELF; its layout-migration.gs), the
 * same once more (ocrPrepareTosgonShelf): until Тосгон's workbook has it the run stops; then every
 * Тосгон cell from E on ("Tos/2026/10!AJ57") moves one column right, in "Бичсэн" and in Тосгон's
 * check copy. Done once: Script property OCR_TOSGON_SHELF_DONE_PROPERTY.
 */

const OCR_PAIRED_DONE_PROPERTY = 'ocr-paired-layout';
const OCR_PAIRED_SINCE = new Date(2026, 0, 1); // the OCR month spreadsheets looked through
const OCR_PAIRED_HEADER_ROW = 3;
const OCR_GIFT_DONE_PROPERTY = 'ocr-gift-column';
const OCR_TOSGON_SHELF_DONE_PROPERTY = 'ocr-tosgon-shelf';

// ---------- pure helpers (no Spreadsheet calls) ----------

// A shop cell in a text: "Sor/2026/10/15!U22" (the tab, the column letters, the row).
const ocrShopCellPattern = () => new RegExp(`((?:${CHAINS.map((chain) => chain.tabPrefix).join('|')})\\/\\d{4}\\/\\d{1,2}\\/\\d{1,2}!)([A-Z]+)(\\d+)`, 'g');

// A "Бичсэн" text with every shop cell moved to its column of the paired layout.
const ocrMovePairedText = (text) =>
  String(text).replace(ocrShopCellPattern(), (all, tab, letters, row) => `${tab}${columnLetter(pairedColumnOf(ocrColumnNumber(letters)))}${row}`);

// A "Бичсэн" text with every shop cell moved to its column once the gift column is in.
const ocrMoveGiftText = (text) =>
  String(text).replace(ocrShopCellPattern(), (all, tab, letters, row) => `${tab}${columnLetter(giftColumnOf(ocrColumnNumber(letters)))}${row}`);

// A "Бичсэн" text with every Тосгон cell from E on moved one column right (the shelf price in E).
const ocrMoveTosgonShelfText = (text) =>
  String(text).replace(new RegExp(`(${OCR_TOSGON.tabPrefix}\\/\\d{4}\\/\\d{1,2}!)([A-Z]+)(\\d+)`, 'g'), (all, tab, letters, row) => {
    const column = ocrColumnNumber(letters);
    return `${tab}${columnLetter(column + (column >= SHELF_PRICE_COLUMN ? 1 : 0))}${row}`;
  });

// ---------- spreadsheet access ----------

// Does the shops' workbook have the paired layout (row 3 of column E of a year sheet)?
function ocrShopsArePaired() {
  const book = SpreadsheetApp.openById(OCR_CATALOG_SPREADSHEET_ID);
  const tab = book.getSheets().find((sheet) => sheetSlot(sheet.getName()) !== null);
  return Boolean(tab) && shopLayoutOf(tab.getRange(OCR_PAIRED_HEADER_ROW, SHELF_PRICE_COLUMN).getValue(), DAY_COLUMNS).paired;
}

// true when the run may go on (see the top of the file).
function ocrPrepareShopLayout(props, root) {
  if (!LAYOUT.paired || props.getProperty(OCR_PAIRED_DONE_PROPERTY) === '1') {
    return true;
  }
  if (!ocrShopsArePaired()) {
    Logger.log('OCR: дэлгүүрийн файл шинэ бүтэц рүү (яг харагдах үнэ) хараахан шилжээгүй тул энэ удаад юу ч уншиж, бичсэнгүй. Тэнд migrateToPairedLayout-ийг ажиллуулна уу.');
    return false;
  }
  const cells = ocrMoveReferences(root, ocrMovePairedText);
  const tabs = ocrMoveCheckCopy(props);
  Logger.log(`OCR: дэлгүүрийн файл шинэ бүтэцтэй боллоо: "Бичсэн" ${cells} нүдийн хаягийг ${DRY_RUN ? 'шилжүүлнэ' : 'шилжүүллээ'}, check хуулбарын ${tabs} хуудсанд багана ${DRY_RUN ? 'нэмнэ' : 'нэмлээ'}.`);
  if (!DRY_RUN) {
    props.setProperty(OCR_PAIRED_DONE_PROPERTY, '1');
  }
  return true;
}

// Is a Тосгон month sheet moved: the shelf price header in E3?
const ocrTosgonSheetHasShelf = (sheet) => String(sheet.getRange(OCR_PAIRED_HEADER_ROW, SHELF_PRICE_COLUMN).getValue()).trim() === SHELF_PRICE_HEADER;

// Does this year's Тосгон workbook have the shelf price column (row 3 of E of a month sheet)?
function ocrTosgonHasShelf() {
  const year = Number(ocrStamp(new Date(), 'yyyy'));
  const book = SpreadsheetApp.openById(bookIdOfYear(year, 'tosgon') || OCR_TOSGON_SPREADSHEET_ID);
  const tab = book.getSheets().find((sheet) => sheet.getName().startsWith(`${OCR_TOSGON.tabPrefix}/`));
  return Boolean(tab) && ocrTosgonSheetHasShelf(tab);
}

// true when the run may go on (see the top of the file): the same as ocrPrepareShopLayout, for
// Тосгон's shelf price column.
function ocrPrepareTosgonShelf(props, root) {
  if (!OCR_TOSGON.layout.shelfPrice || props.getProperty(OCR_TOSGON_SHELF_DONE_PROPERTY) === '1') {
    return true;
  }
  if (!ocrTosgonHasShelf()) {
    Logger.log('OCR: Тосгоны файл "яг харагдах үнэ" (E) багана руу хараахан шилжээгүй тул энэ удаад юу ч уншиж, бичсэнгүй. Тэнд migrateToPairedLayout-ийг ажиллуулна уу.');
    return false;
  }
  const cells = ocrMoveReferences(root, ocrMoveTosgonShelfText);
  const tabs = ocrTosgonShelfCheckCopy(props);
  Logger.log(`OCR: Тосгоны файлд "яг харагдах үнэ" нэмэгдлээ: "Бичсэн" ${cells} нүдийн хаягийг ${DRY_RUN ? 'шилжүүлнэ' : 'шилжүүллээ'}, check хуулбарын ${tabs} хуудсанд багана ${DRY_RUN ? 'нэмнэ' : 'нэмлээ'}.`);
  if (!DRY_RUN) {
    props.setProperty(OCR_TOSGON_SHELF_DONE_PROPERTY, '1');
  }
  return true;
}

// The check copy of Тосгон's workbook: the shelf price column in every month sheet without it.
// Returns how many sheets.
function ocrTosgonShelfCheckCopy(props) {
  const year = Number(ocrStamp(new Date(), 'yyyy'));
  const id = props.getProperty(OCR_CHECK_COPY_PROPERTY + (bookIdOfYear(year, 'tosgon') || OCR_TOSGON_SPREADSHEET_ID));
  if (!id || !ocrFileIsThere(id)) {
    return 0;
  }
  const sheets = SpreadsheetApp.openById(id)
    .getSheets()
    .filter((sheet) => sheet.getName().startsWith(`${OCR_TOSGON.tabPrefix}/`))
    .filter((sheet) => !ocrTosgonSheetHasShelf(sheet));
  if (!DRY_RUN) {
    sheets.forEach((sheet) => {
      sheet.insertColumnsAfter(SHELF_PRICE_COLUMN - 1, 1);
      sheet.getRange(OCR_PAIRED_HEADER_ROW, SHELF_PRICE_COLUMN).setValue(SHELF_PRICE_HEADER);
    });
  }
  return sheets.length;
}

// Does the shops' workbook have the gift column (row 3 right after the transfer out)?
function ocrShopsHaveGift() {
  const book = SpreadsheetApp.openById(OCR_CATALOG_SPREADSHEET_ID);
  const tab = book.getSheets().find((sheet) => sheetSlot(sheet.getName()) !== null);
  return Boolean(tab) && String(tab.getRange(OCR_PAIRED_HEADER_ROW, giftHeaderColumn(DAY_COLUMNS)).getValue()).trim() === GIFT_HEADER;
}

// true when the run may go on (see the top of the file): the same as ocrPrepareShopLayout, for
// the gift column.
function ocrPrepareGiftColumn(props, root) {
  if (!LAYOUT.gift || props.getProperty(OCR_GIFT_DONE_PROPERTY) === '1') {
    return true;
  }
  if (!ocrShopsHaveGift()) {
    Logger.log('OCR: дэлгүүрийн файлд "Бэлэг" багана хараахан нэмэгдээгүй тул энэ удаад юу ч уншиж, бичсэнгүй. Тэнд migrateToGiftColumn-ийг ажиллуулна уу.');
    return false;
  }
  const cells = ocrMoveReferences(root, ocrMoveGiftText);
  const tabs = ocrGiftCheckCopy(props);
  Logger.log(`OCR: дэлгүүрийн файлд "Бэлэг" багана нэмэгдлээ: "Бичсэн" ${cells} нүдийн хаягийг ${DRY_RUN ? 'шилжүүлнэ' : 'шилжүүллээ'}, check хуулбарын ${tabs} хуудсанд багана ${DRY_RUN ? 'нэмнэ' : 'нэмлээ'}.`);
  if (!DRY_RUN) {
    props.setProperty(OCR_GIFT_DONE_PROPERTY, '1');
  }
  return true;
}

// The check copy of the shops' workbook: the gift column in every year sheet that has none.
// Returns how many sheets.
function ocrGiftCheckCopy(props) {
  const id = props.getProperty(OCR_CHECK_COPY_PROPERTY + OCR_CATALOG_SPREADSHEET_ID);
  if (!id || !ocrFileIsThere(id)) {
    return 0;
  }
  const column = giftHeaderColumn(DAY_COLUMNS);
  const sheets = SpreadsheetApp.openById(id)
    .getSheets()
    .filter((sheet) => sheetSlot(sheet.getName()) !== null)
    .filter((sheet) => String(sheet.getRange(OCR_PAIRED_HEADER_ROW, column).getValue()).trim() !== GIFT_HEADER);
  if (!DRY_RUN) {
    sheets.forEach((sheet) => {
      sheet.insertColumnsAfter(column - 1, 1);
      sheet.getRange(OCR_PAIRED_HEADER_ROW, column).setValue(GIFT_HEADER);
    });
  }
  return sheets.length;
}

// Every "Бичсэн" of the OCR month spreadsheets, moved by move(text). Returns how many cells changed.
function ocrMoveReferences(root, move) {
  let changed = 0;
  findOcrBooks(root, OCR_PAIRED_SINCE).forEach(({ file }) => {
    ocrDayTabs(SpreadsheetApp.openById(file.getId())).forEach((tab) => {
      const height = tab.getLastRow() - 1;
      const column = tab.getRange(1, 1, 1, Math.max(tab.getLastColumn(), 1)).getValues()[0].indexOf('Бичсэн') + 1;
      if (height <= 0 || column === 0) {
        return;
      }
      const range = tab.getRange(2, column, height, 1);
      const before = range.getValues();
      const after = before.map(([text]) => [text === '' ? '' : move(text)]);
      const count = after.filter(([text], index) => text !== before[index][0]).length;
      if (count > 0 && !DRY_RUN) {
        range.setValues(after);
      }
      changed += count;
    });
  });
  return changed;
}

// The check copy of the shops' workbook: the same two columns in every year sheet not moved yet.
// Returns how many sheets.
function ocrMoveCheckCopy(props) {
  const id = props.getProperty(OCR_CHECK_COPY_PROPERTY + OCR_CATALOG_SPREADSHEET_ID);
  if (!id || !ocrFileIsThere(id)) {
    return 0;
  }
  const paired = buildColumnLayout(DAY_COLUMNS, { paired: true });
  const sheets = SpreadsheetApp.openById(id)
    .getSheets()
    .filter((sheet) => sheetSlot(sheet.getName()) !== null)
    .filter((sheet) => !shopLayoutOf(sheet.getRange(OCR_PAIRED_HEADER_ROW, SHELF_PRICE_COLUMN).getValue(), DAY_COLUMNS).paired);
  if (!DRY_RUN) {
    sheets.forEach((sheet) => {
      sheet.insertColumnsAfter(SHELF_PRICE_COLUMN - 1, 1);
      sheet.getRange(OCR_PAIRED_HEADER_ROW, SHELF_PRICE_COLUMN).setValue(SHELF_PRICE_HEADER);
      sheet.insertColumnsAfter(paired.otherBalance, 1);
    });
  }
  return sheets.length;
}
