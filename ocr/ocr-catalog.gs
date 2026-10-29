/**
 * Receipt OCR - the product names a receipt line may be matched to, with their selling price as a
 * hint for Gemini. Read from the inventory workbook (never changed):
 *   - "Бараа": the one product list of both shops, read first so its spelling wins; only its names
 *     (C..F there show one shop's stock, cost and prices, a view: D is the cost, not the price)
 *   - the sheets of the current half-month (e.g. Sor/2026/10/31, My/2026/10/31): the newest names
 *     and prices, so a product added this month is known without anyone updating a list
 *   - Sheet100 (Соронзон) and Sheet100* (My home): the owner's name lists (a 30 September copy)
 *   - this month's sheet of the Тосгон workbook (Tos/2026/10): its names too
 * Once the list exists, only its names are offered (the owner, 2026-10-25): the other tabs give prices.
 * A name found in several places is listed once, with the price of the first place that has one.
 * The price is the one the customer sees: E "яг харагдах үнэ" of the shops' sheets; D (зарах үнэ)
 * where E is empty, and on Тосгон's sheet, whose E is a day.
 */

const OCR_CATALOG_PRICE_COLUMN = 4; // D, the selling price
const OCR_CATALOG_WIDTH = SHELF_PRICE_COLUMN; // A..E

const ocrCleanName = (value) => String(value).replace(/\s+/g, ' ').trim();
// How a product is told apart: no case, no spaces ("Haier 300L" = "haier300l"), so the same
// product spelt a little differently in two workbooks is one product.
const ocrNameKey = (name) => String(name === undefined || name === null ? '' : name).toLowerCase().replace(/\s+/g, '');

// The tabs to read for "YYYY-MM-DD": the product list first (its spelling wins), then the current
// sheet of each shop (names added since, and the prices), then the old name lists.
function ocrCatalogTabs(today) {
  const [year, month, day] = today.split('-').map(Number);
  const current = !ocrHasBooksFor(year) ? [] : withChainYear(year, () => {
    const position = periodPosition(year, month, day > FIRST_HALF_LAST_DAY ? 1 : 0);
    return position >= 0 && position < CHAIN_PERIOD_COUNT ? CHAINS.map((chain) => chainSheetName(chain, position)) : [];
  });
  return [OCR_PRODUCT_LIST_TAB, ...current, ...OCR_CATALOG_TABS];
}

// rows: [name, opening, cost, price] per product row, tabs in order -> [{ name, price }], each
// product once (ocrNameKey), with the spelling and the price of the first tab that has them.
function ocrCatalogEntries(rows) {
  const entries = new Map();
  rows.forEach(([rawName, , , rawPrice]) => {
    const name = ocrCleanName(rawName === undefined ? '' : rawName);
    const price = typeof rawPrice === 'number' && rawPrice > 0 ? rawPrice : null;
    if (name === '') {
      return;
    }
    const known = entries.get(ocrNameKey(name));
    if (!known) {
      entries.set(ocrNameKey(name), { name, price });
    } else if (known.price === null && price !== null) {
      known.price = price;
    }
  });
  return [...entries.values()];
}

// How a name is shown to Gemini: "Haier 300l | 1200000₮" (the price only as a hint).
const ocrCatalogLine = (entry) => (entry.price === null ? entry.name : `${entry.name} | ${entry.price}₮`);

// The name rows [name, opening, cost, price] of some tabs of a workbook; a missing tab is noted.
function ocrCatalogRows(book, tabs) {
  return tabs.flatMap((tab) => {
    const sheet = book.getSheetByName(tab);
    if (!sheet) {
      Logger.log(`Барааны нэр: "${tab}" sheet олдсонгүй, алгаслаа.`);
      return [];
    }
    const count = sheet.getLastRow() - OCR_CATALOG_FIRST_ROW + 1;
    const rows = count > 0 ? sheet.getRange(OCR_CATALOG_FIRST_ROW, 1, count, OCR_CATALOG_WIDTH).getValues() : [];
    // In the product list a category heading is the row without a shop.
    if (tab === OCR_PRODUCT_LIST_TAB) {
      return rows.filter((row) => ocrCleanName(row[OCR_PRODUCT_LIST_SHOP_INDEX] ?? '') !== '').map(([name]) => [name, '', '', '']);
    }
    const hasShelf = ocrCleanName(sheet.getRange(LAYOUT_HEADER_ROW, SHELF_PRICE_COLUMN).getValue() ?? '') === SHELF_PRICE_HEADER;
    return rows.map((row) => [row[0], '', '', ocrShownPrice(row, hasShelf)]);
  });
}

// The price a customer sees on a row [name, opening, cost, price, shelf]: E when the sheet has the
// shelf price column and E is filled, else D.
function ocrShownPrice(row, hasShelf) {
  const shelf = row[SHELF_PRICE_COLUMN - 1];
  return hasShelf && typeof shelf === 'number' && shelf > 0 ? shelf : row[OCR_CATALOG_PRICE_COLUMN - 1];
}

// The category headings of the product list ("Хөргөгч"): rows of a shop sheet or Sheet100 with
// such a name are headings there too, not products.
function ocrListHeadingKeys(book) {
  const list = book.getSheetByName(OCR_PRODUCT_LIST_TAB);
  const count = list ? list.getLastRow() - OCR_CATALOG_FIRST_ROW + 1 : 0;
  const rows = count > 0 ? list.getRange(OCR_CATALOG_FIRST_ROW, 1, count, OCR_PRODUCT_LIST_SHOP_INDEX + 1).getValues() : [];
  return new Set(rows.filter((row) => ocrCleanName(row[OCR_PRODUCT_LIST_SHOP_INDEX] ?? '') === '').map(([name]) => ocrNameKey(ocrCleanName(name ?? ''))));
}

// The products of the list (rows with a shop), or null while the workbook has none: then the
// sheets' names are all there is. With a list, a name is offered only when the list has it (the
// owner's rule, 2026-10-25); the other tabs give the prices only.
function ocrListProductKeys(book) {
  const list = book.getSheetByName(OCR_PRODUCT_LIST_TAB);
  const count = list ? list.getLastRow() - OCR_CATALOG_FIRST_ROW + 1 : 0;
  const rows = count > 0 ? list.getRange(OCR_CATALOG_FIRST_ROW, 1, count, OCR_PRODUCT_LIST_SHOP_INDEX + 1).getValues() : [];
  const keys = rows.filter((row) => ocrCleanName(row[OCR_PRODUCT_LIST_SHOP_INDEX] ?? '') !== '').map(([name]) => ocrNameKey(ocrCleanName(name ?? '')));
  return keys.length > 0 ? new Set(keys) : null;
}

// The id of a year's inventory workbook, the one the product list "Бараа" is in.
const ocrCatalogBookId = (year) => bookIdOfYear(year, 'main') || OCR_CATALOG_SPREADSHEET_ID;

function loadOcrCatalog(today) {
  const [year, month] = today.split('-').map(Number);
  const main = SpreadsheetApp.openById(ocrCatalogBookId(year));
  const rows = [
    ...ocrCatalogRows(main, ocrCatalogTabs(today)),
    ...ocrCatalogRows(SpreadsheetApp.openById(bookIdOfYear(year, 'tosgon') || OCR_TOSGON_SPREADSHEET_ID), [`${OCR_TOSGON.tabPrefix}/${year}/${month}`]),
  ];
  const headings = ocrListHeadingKeys(main);
  const listed = ocrListProductKeys(main);
  const entries = ocrCatalogEntries(rows).filter(
    (entry) => !headings.has(ocrNameKey(entry.name)) && (listed === null || listed.has(ocrNameKey(entry.name)))
  );
  if (entries.length === 0) {
    throw new Error('Барааны нэрсийн жагсаалт хоосон байна: "Бараа", Sheet100 эсвэл энэ хагас сарын sheet-ийг шалгана уу.');
  }
  const names = entries.map((entry) => entry.name);
  return { names, entries, lines: entries.map(ocrCatalogLine), set: new Set(names) };
}
