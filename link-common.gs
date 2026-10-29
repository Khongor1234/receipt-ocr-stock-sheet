/**
 * What the shops' side (price-link.gs, main workbook) and Тосгон's side (tosgon-link.gs) of
 * "one product, one price" share: the categories only the shops hold, and writing one linked
 * value into a run of sheets. Needs config.gs, year-sync.gs, price-sync.gs (isSamePrice, findProductRow).
 */

// Small goods only the shops hold; every category whose name starts so ("хятад жижиг бараа будаа агшаагч" too).
const LINK_SHOP_ONLY_CATEGORIES = ['хятад жижиг бараа', 'солонгос жижиг бараа'];

const isShopOnlyCategory = (categoryKey) => LINK_SHOP_ONLY_CATEGORIES.some((prefix) => String(categoryKey || '').startsWith(prefix));

// The names of a sheet are read once per edit: one paste may link many rows into the same sheets.
const LINK_NAMES = new Map();
function productRowOf(sheet, key) {
  const firstRow = SYNC_HEADER_ROWS + 1;
  const id = `${sheet.getParent().getId ? sheet.getParent().getId() : ''}|${sheet.getName()}`;
  if (!LINK_NAMES.has(id)) {
    const height = sheet.getLastRow() - SYNC_HEADER_ROWS;
    LINK_NAMES.set(id, height > 0 ? sheet.getRange(firstRow, SYNC_NAME_COL, height, 1).getValues() : []);
  }
  return findProductRow(LINK_NAMES.get(id), key, firstRow);
}

// The first sheet gets the value; each later one too while it still holds the first one's old value.
// { cells: how many were written, found: whether the first sheet has the product }.
function writeLinked(sheets, key, column, value) {
  const [first, ...later] = sheets;
  const firstRow = first ? productRowOf(first, key) : null;
  if (firstRow === null) {
    return { cells: 0, found: false };
  }
  const old = first.getRange(firstRow, column).getValue();
  if (isSamePrice(old, value)) {
    return { cells: 0, found: true };
  }
  if (!DRY_RUN) {
    first.getRange(firstRow, column).setValue(value);
  }
  let cells = 1;
  for (const sheet of later) {
    const row = productRowOf(sheet, key);
    if (row === null) {
      continue;
    }
    const cell = sheet.getRange(row, column);
    if (!isSamePrice(cell.getValue(), old)) {
      break; // this sheet has its own value
    }
    if (!DRY_RUN) {
      cell.setValue(value);
    }
    cells += 1;
  }
  return { cells, found: true };
}

