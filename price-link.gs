/**
 * One product, one price in all three places, and Соронзон's new products in Тосгон too (the
 * owner's rules, 2026-10-20):
 *   - the cost (C) and the price (D) are the same in Соронзон, My home and Тосгон; the shelf price
 *     (E) is the same in the two shops (Тосгон has none). One typed in a shop is written at once
 *     into the other shop's sheet of the same half-month and into Тосгон's sheet of that month,
 *     then carried to their later sheets as long as those still hold the value the first one had
 *     (the rule of price-sync.gs). One typed in Тосгон reaches both shops (tosgon-link.gs).
 *   - a product added in Соронзон is added to Тосгон's sheet of that month, unless its category is
 *     one of LINK_SHOP_ONLY_CATEGORIES (small goods Тосгон never holds); Тосгон's own background
 *     job carries it to its later months.
 *   - previewPriceLink (report only) / setupPriceLink: the one-time alignment. Тосгон's sheets of
 *     this month and the months after get Соронзон's cost and price of the same month (its second
 *     half); every product of Соронзон that Тосгон's current month lacks is added. The months
 *     before stay as they are (their profit is already counted). setupPriceLink then turns the
 *     linking on (LINK_READY_KEY); until then edits are not linked.
 * Products are matched by name (normalizeKey): a name Тосгон spells another way is not found, so
 * the Тосгон name table (tosgon-names.gs) is applied first.
 * Main workbook only (build.sh); Тосгон's side is tosgon-link.gs. Needs config.gs, year-sync.gs
 * (readModel, findInsertAfterRow, pickTemplateRow), link-common.gs (writeLinked, isShopOnlyCategory),
 * tosgon-names.gs (TOSGON_BOOK_ID).
 */

const LINK_READY_KEY = 'priceLinkReady'; // document property: '1' once setupPriceLink ran
const LINK_TOSGON_PREFIX = 'Tos';
// Тосгон's month sheets: the shelf price in E once moved (2026-10-22, config.gs tosgonLayoutOf).
const tosgonSheetLayout = (sheet) => tosgonLayoutOf(sheet.getRange(LAYOUT_HEADER_ROW, SHELF_PRICE_COLUMN).getValue());
const LINK_SOURCE_CHAIN = 'соронзон'; // the shop whose new products go to Тосгон
const LINK_FIELDS = [
  { label: 'өртөг', column: 3, tosgon: true },
  { label: 'үнэ', column: 4, tosgon: true },
  ...(LAYOUT.shelfPrice ? [{ label: 'яг харагдах үнэ', column: LAYOUT.shelfPrice, tosgon: true }] : []),
];
const LINK_REPORT_LIMIT = 40;
// One edit links at most this many rows (an edit has 30 seconds); a bigger paste is left to setupPriceLink.
const LINK_MAX_ROWS = 50;

// ---------- pure helpers (no Spreadsheet calls) ----------

const otherChainOf = (chainKey) => CHAINS.find((chain) => chain.key !== chainKey);
const linkFieldOf = (column) => LINK_FIELDS.find((field) => field.column === column) || null;

// "Tos/2026/10": Тосгон's sheet of the month of a period position.
function tosgonTabOf(position) {
  const { year, month } = periodMonth(position);
  return `${LINK_TOSGON_PREFIX}/${year}/${month}`;
}

// ---------- writing a linked value ----------

// The month sheets of Тосгон from a period's month to the end of the year, those that exist.
function tosgonSheetsFrom(tosgonBook, position) {
  const { year, month } = periodMonth(position);
  return Array.from({ length: MONTHS_PER_YEAR - month + 1 }, (_, step) => tosgonBook.getSheetByName(`${LINK_TOSGON_PREFIX}/${year}/${month + step}`))
    .filter(Boolean);
}

// A Тосгон sheet takes a field when it has its column: the shelf price only once moved (on a sheet
// not moved yet E is the first income day).
const tosgonHasField = (sheet, field) => field.column !== SHELF_PRICE_COLUMN || tosgonSheetLayout(sheet).shelfPrice !== null;

// One price cell of a shop sheet into the other shop and Тосгон.
function linkPriceCell(spreadsheet, tosgonBook, sheet, row, column) {
  const slot = sheetSlot(sheet.getName());
  const field = linkFieldOf(column);
  const label = row > SYNC_HEADER_ROWS ? String(sheet.getRange(row, SYNC_NAME_COL).getValue()).trim() : '';
  if (slot === null || !field || label === '') {
    return [];
  }
  const key = normalizeKey(label);
  const value = sheet.getRange(row, column).getValue();
  const other = otherChainOf(slot.chainKey);
  const shopSheets = findChains(spreadsheet, { chainKey: other.key, fromPosition: slot.position }).flatMap((chain) => chain.sheets);
  const places = [[other.label, writeLinked(shopSheets, key, column, value)]];
  if (field.tosgon && tosgonBook) {
    const tosgonSheets = tosgonSheetsFrom(tosgonBook, slot.position).filter((tosSheet) => tosgonHasField(tosSheet, field));
    places.push(['Тосгон', writeLinked(tosgonSheets, key, column, value)]);
  }
  return places
    .filter(([, result]) => result.cells > 0)
    .map(([place, result]) => `${sheet.getName()}: "${label}" ${field.label} ${value} -> ${place} (${result.cells} sheet)`);
}

// ---------- Соронзон's new products into Тосгон ----------

// Adds to a Тосгон month sheet the products of a Соронзон sheet it lacks (shop-only categories
// excepted); write false only lists them. Only names of the list "Бараа" go: Тосгон's column A takes
// no other (its dropdown), so a name off the list is reported instead (2026-10-22: "Panasonica 65").
// A product Тосгон already has is never added again, so a rerun adds only what is still missing.
// Returns the names added (or to add); notes: what was left out, and why.
function addMissingToTosgon(sorSheet, tosgonSheet, write, notes = []) {
  const tosgonKeys = new Set(readModel(tosgonSheet).products.map((product) => product.key));
  const list = sorSheet.getParent().getSheetByName(PRODUCT_LIST_TAB);
  // key -> the list's spelling: Тосгон's dropdown takes "Katro p-24", not "Katro  p-24" as typed
  const listNames = list === null ? null : new Map(readListNames(list).map(({ name }) => [normalizeKey(name), name]));
  const missing = readModel(sorSheet).products
    .filter((product) => !isShopOnlyCategory(product.categoryKey) && !tosgonKeys.has(product.key))
    .map((product) => (listNames !== null && listNames.has(product.key) ? { ...product, label: listNames.get(product.key) } : product));
  const offList = missing.filter((product) => listNames !== null && !listNames.has(product.key));
  offList.forEach((product) => notes.push(`"${product.label}": "${PRODUCT_LIST_TAB}" жагсаалтад алга тул Тосгонд нэмсэнгүй`));
  const added = missing.filter((product) => !offList.includes(product)).filter((product) => {
    if (!write) {
      return true;
    }
    try {
      insertTosgonProduct(tosgonSheet, product);
      SpreadsheetApp.flush(); // a refused row shows here, with its own name
      return true;
    } catch (error) {
      notes.push(`"${product.label}": Тосгонд нэмэгдсэнгүй (${error.message})`);
      return false;
    }
  });
  return added.map((product) => product.label);
}

// A row in Тосгон's layout: under its category when Тосгон has it, else after the last product;
// only the name, cost and price are written, and the shelf price on a moved sheet (E; not moved: a day).
function insertTosgonProduct(sheet, product) {
  const model = readModel(sheet);
  const afterRow = findInsertAfterRow(model, product.categoryKey, product.label); // in name order (official-names.gs)
  const templateRow = pickTemplateRow(model, product.categoryKey);
  sheet.insertRowAfter(afterRow);
  const newRow = afterRow + 1;
  const layout = tosgonSheetLayout(sheet);
  const width = layout.count;
  if (templateRow !== null) {
    const source = templateRow > afterRow ? templateRow + 1 : templateRow;
    sheet.getRange(source, 1, 1, width).copyTo(sheet.getRange(newRow, 1, 1, width));
    clearCopiedMarks(sheet, newRow, width); // the template's notes are no product's of this row (row-copies.gs)
  }
  sheet.getRange(newRow, 1, 1, width).clearContent();
  sheet.getRange(newRow, SYNC_NAME_COL).setValue(product.label);
  sheet.getRange(newRow, 3, 1, 2).setValues([[product.cost, product.price]]);
  if (layout.shelfPrice !== null) {
    sheet.getRange(newRow, layout.shelfPrice).setValue(product.shelfPrice === undefined ? '' : product.shelfPrice);
  }
}

// ---------- the one-time alignment ----------

// The cost and price cells of a Тосгон sheet that differ from a Соронзон sheet: [{ row, column, label, from, to, name }].
function priceDifferences(sorSheet, tosSheet) {
  const tosRows = new Map(readModel(tosSheet).products.map((product) => [product.key, product]));
  return readModel(sorSheet).products.flatMap((product) => {
    const there = tosRows.get(product.key);
    if (!there) {
      return [];
    }
    const shelf = tosgonSheetLayout(tosSheet).shelfPrice !== null && LAYOUT.shelfPrice ? [['яг харагдах үнэ', SHELF_PRICE_COLUMN, product.shelfPrice, there.shelfPrice]] : [];
    return [['өртөг', 3, product.cost, there.cost], ['үнэ', 4, product.price, there.price], ...shelf]
      .filter(([, , wanted, current]) => wanted !== undefined && wanted !== '' && !isSamePrice(wanted, current))
      .map(([label, column, wanted, current]) => ({ row: there.row, column, label, from: current, to: wanted, name: product.label }));
  });
}

// Тосгон's sheets of this month and later get Соронзон's cost and price of the same month (its
// second half); Тосгон's current month gets the products of Соронзон's current sheet it lacks.
// write false: the same report, nothing written.
function runPriceAlignment(spreadsheet, tosgonBook, today, write) {
  const sor = CHAINS.find((chain) => chain.key === LINK_SOURCE_CHAIN);
  const year = today.getFullYear();
  const lines = [];
  let cells = 0;
  for (let month = today.getMonth() + 1; month <= MONTHS_PER_YEAR; month += 1) {
    const sorSheet = spreadsheet.getSheetByName(chainSheetName(sor, periodPosition(year, month, 1)));
    const tosSheet = tosgonBook.getSheetByName(`${LINK_TOSGON_PREFIX}/${year}/${month}`);
    if (!sorSheet || !tosSheet) {
      continue;
    }
    priceDifferences(sorSheet, tosSheet).forEach((change) => {
      if (write) {
        tosSheet.getRange(change.row, change.column).setValue(change.to);
      }
      cells += 1;
      lines.push(`${tosSheet.getName()}: "${change.name}" ${change.label} ${change.from} -> ${change.to}`);
    });
  }
  return { ...runMissingToTosgon(spreadsheet, tosgonBook, today, write), cells, lines };
}

// Only the products of Соронзон's current sheet that Тосгон's current month lacks; no price is touched.
function runMissingToTosgon(spreadsheet, tosgonBook, today, write) {
  const sor = CHAINS.find((chain) => chain.key === LINK_SOURCE_CHAIN);
  const now = periodPosition(today.getFullYear(), today.getMonth() + 1, today.getDate() > FIRST_HALF_LAST_DAY ? 1 : 0);
  const sorNow = spreadsheet.getSheetByName(chainSheetName(sor, now));
  const tosNow = tosgonBook.getSheetByName(tosgonTabOf(now));
  const notes = [];
  const added = sorNow && tosNow ? addMissingToTosgon(sorNow, tosNow, write, notes) : [];
  return { cells: 0, lines: [], added, notes, tab: tosNow ? tosNow.getName() : '', write };
}

function formatAlignment(result) {
  const mode = result.write ? '' : '[ТУРШИЛТ - юу ч өөрчлөгдөөгүй] ';
  return [
    `${mode}Тосгоны үнэ Соронзонтой: ${result.cells} нүд; ${result.tab || 'Тосгоны энэ сарын sheet'}-д нэмэх бараа: ${result.added.length}`,
    ...result.lines.slice(0, LINK_REPORT_LIMIT),
    ...(result.lines.length > LINK_REPORT_LIMIT ? [`  ... бас ${result.lines.length - LINK_REPORT_LIMIT}`] : []),
    ...(result.added.length > 0 ? [`Тосгонд нэмэх: ${result.added.join(', ')}`] : []),
    ...(result.notes || []),
  ];
}

// ---------- run by hand, and from onSheetEdit ----------

function previewPriceLink() {
  const result = runPriceAlignment(currentBook(), partnerBook(currentBook(), 'tosgon', TOSGON_BOOK_ID), new Date(), false);
  Logger.log(formatAlignment(result).join('\n'));
}

// Run by hand: Тосгон gets the products it lacks, the prices are left as they are.
function addMissingProductsToTosgon() {
  withScriptLock(() => {
    const spreadsheet = currentBook();
    const result = runMissingToTosgon(spreadsheet, partnerBook(spreadsheet, 'tosgon', TOSGON_BOOK_ID), new Date(), !DRY_RUN);
    Logger.log(formatAlignment(result).join('\n'));
  });
}

function setupPriceLink() {
  withScriptLock(() => {
    const spreadsheet = currentBook();
    const result = runPriceAlignment(spreadsheet, partnerBook(spreadsheet, 'tosgon', TOSGON_BOOK_ID), new Date(), !DRY_RUN);
    if (!DRY_RUN) {
      bookProperties().setProperty(LINK_READY_KEY, '1');
    }
    Logger.log([...formatAlignment(result), DRY_RUN ? '' : 'Үнийн холбоос асаалттай: дэлгүүрт бичсэн үнэ нөгөө дэлгүүр, Тосгонд шууд орно.'].join('\n'));
  });
}

const isPriceLinkReady = () => bookProperties().getProperty(LINK_READY_KEY) === '1';

// onSheetEdit: the price cells of an edited range, and a name typed on a Соронзон sheet.
// A failure is logged and never stops the edit's own sync.
function linkEdit(range) {
  if (!isPriceLinkReady()) {
    return;
  }
  try {
    const sheet = range.getSheet();
    const slot = sheetSlot(sheet.getName());
    if (slot === null) {
      return;
    }
    const spreadsheet = sheet.getParent();
    let tosgonBook = null;
    const tosgon = () => (tosgonBook = tosgonBook || partnerBook(spreadsheet, 'tosgon', TOSGON_BOOK_ID)); // the same year's Тосгон
    const fields = LINK_FIELDS.filter((field) => field.column >= range.getColumn() && field.column <= range.getLastColumn());
    const lines = [];
    LINK_NAMES.clear();
    const lastRow = Math.min(range.getLastRow(), range.getRow() + LINK_MAX_ROWS - 1);
    if (fields.length > 0 && lastRow < range.getLastRow()) {
      lines.push(`Үнийн холбоос: эхний ${LINK_MAX_ROWS} мөрийг л холбов, үлдсэнийг setupPriceLink-ээр тааруулна.`);
    }
    for (let row = range.getRow(); row <= lastRow; row += 1) {
      fields.forEach((field) => lines.push(...linkPriceCell(spreadsheet, field.tosgon ? tosgon() : null, sheet, row, field.column)));
    }
    if (range.getColumn() <= SYNC_NAME_COL && slot.chainKey === LINK_SOURCE_CHAIN) {
      LINK_NAMES.clear(); // the rows of Тосгон move as products are added
      const tosSheet = tosgon().getSheetByName(tosgonTabOf(slot.position));
      const added = tosSheet ? addMissingToTosgon(sheet, tosSheet, !DRY_RUN) : [];
      if (added.length > 0) {
        lines.push(`Тосгонд нэмсэн (${tosSheet.getName()}): ${added.join(', ')}`);
      }
    }
    if (lines.length > 0) {
      Logger.log(lines.join('\n'));
    }
  } catch (error) {
    Logger.log(`Үнийн холбоос: ${error.message}`);
  }
}
