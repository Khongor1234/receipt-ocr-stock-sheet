/**
 * Protects the full-width category divider rows ("Угаалгын машин", "Хөргөгч", ...)
 * so they cannot be deleted by accident. Product rows stay free to add and delete.
 *
 * A divider row is a coloured row with a NAME and without numbers (no balance,
 * price, movement or count in B..AK). Two forms exist in the file:
 *   1. a full-width band: column A and a spread of probe columns across the row
 *      all share the same non-white background (Соронзон sheets)
 *   2. a coloured cell in column A holding one of DIVIDER_CATEGORY_NAMES ("Хөргөгч",
 *      "Плитка", ... on the My home sheets); its colour is extended across the row
 *
 * Inherited colour. A row inserted next to a coloured row gets that colour from
 * Google Sheets, so a coloured row is not automatically a divider (classifyRows):
 *   - a row with numbers is always a product; if it is coloured across the whole
 *     row, the colour is removed
 *   - a named row without numbers that touches a category or a protected divider of
 *     the same colour is a new product row too
 *   - a product cell or an empty row of the same colour touching one of those loses
 *     the colour as well
 * Such rows are never protected. An empty coloured row standing on its own is a
 * divider too.
 *
 * Required categories (DIVIDER_REQUIRED_CATEGORIES): a category whose name was
 * erased, or whose row was deleted, is put back above its first product.
 * Same categories everywhere (runCategorySync): a category that the other sheets
 * of the shop have and one sheet lacks is made again on that sheet.
 *
 * Colours. Google Sheets cannot forbid colouring to someone who may type, so the
 * colours are put back instead (reconcileSheet, at once through instant-sync.gs):
 *   - a divider is coloured across the whole row, without gaps
 *   - every other row gets the usual colour of each column (white, or the tint of
 *     a whole column such as T and AJ)
 *   - a cell the OWNER coloured stays: when the owner changes a colour, the colours
 *     of that sheet are remembered (approveSheetColors) under the product's name
 * So a colour put on by anyone else disappears within seconds. To make a new
 * category: type its name first, then colour the row; right next to another
 * category use a name from DIVIDER_CATEGORY_NAMES or a different colour.
 *
 * Deleting a row that contains a protected cell is refused for everyone except the
 * file owner, and protections move with their row when rows are inserted/deleted,
 * so this only needs to run again after a NEW divider has been created.
 *
 * isDividerRow / DIVIDER_SCAN_WIDTH are also used by year-sync.gs, keep both files.
 *
 * Usage: run protectDividers (DRY_RUN in config.gs decides whether it only
 * reports). If the time budget is reached the log says so; just run it again,
 * finished rows are skipped. The background job (year-sync.gs) runs it too, so a
 * newly made divider is protected within the hour.
 * removeDividerProtections undoes it (only protections created by this script).
 */

const DIVIDER_FIRST_ROW = 4;
const DIVIDER_SCAN_WIDTH = LAYOUT.transferOut - 1; // A..AI
const DIVIDER_VALUE_WIDTH = LAYOUT.count; // A..AK
const DIVIDER_APPROVED_KEY = 'approvedColors:'; // + sheet id: the cells the owner coloured
const DIVIDER_BORDERS_KEY = 'dividerBorders:'; // + sheet id: the stray cell borders of the sheet are gone
// The daily columns E..AJ have no cell borders of their own, only one line after
// T that parts income from expense. (The line after AJ belongs to AJ/AK and stays.)
const DIVIDER_BODY_FIRST_COLUMN = LAYOUT.incomeFirst; // E
const DIVIDER_BODY_LAST_COLUMN = LAYOUT.expenseLast; // AJ (with the gift column: AL)
const DIVIDER_SEPARATOR_COLUMN = LAYOUT.transferIn; // T
// Every row down to the last product gets the grid of the table: A..D and AK..AN (the shops'
// layout: A..E with the shelf price, AL..AP with the other shop's two columns).
const DIVIDER_GRID_ZONES = [
  { firstColumn: 1, width: LAYOUT.incomeFirst - 1 }, // A..D
  { firstColumn: LAYOUT.count, width: LAYOUT.summaryFirst - LAYOUT.count }, // AK..AN
];
// What the border mark of a sheet holds: the grid it was drawn with ("1+5,67+4" = A..E and
// BO..BR). A sheet marked with another grid gets it drawn again once: the Тосгон sheets were
// marked "1" with A..D before the shelf price came in E (2026-10-22), so E had no boxes.
const DIVIDER_GRID_MARK = DIVIDER_GRID_ZONES.map((zone) => `${zone.firstColumn}+${zone.width}`).join(',');
// Categories that must exist. Each one sits directly above the first product whose
// name starts with firstProductPrefix. Its name is written back if it was erased,
// and the whole row is made again if it was deleted.
// A category is carried to the other sheets of its shop when its name is one of
// DIVIDER_CATEGORY_NAMES or when at least this many sheets have it. A row that only
// one or two sheets treat as a category is more likely an accident.
const DIVIDER_SHARED_MIN_SHEETS = 3;
// Category names that are written as a coloured cell in column A only (lower case).
// Add a name here when a new category of that kind is created.
const DIVIDER_CATEGORY_NAMES = [
  'угаалгын машин',
  'хөргөгч',
  'хөлдөөгч',
  'хөргүүр',
  'плитка',
  'плетик',
  'индүү',
  'зурагт',
  'будаа агшаагч',
  'жижиг бараа',
  'бусад (жижиг бараа)',
  'хятад жижиг бараа',
  'хятад жижиг бараа будаа агшаагч',
  'солонгос жижиг бараа',
];
// A F K P V AA AF AI: the name, every fifth day column of both blocks, the last day column
const DIVIDER_PROBE_COLUMNS = buildProbeColumns();

function buildProbeColumns() {
  const everyFifth = (from, to) => Array.from({ length: Math.floor((to - from) / 5) + 1 }, (_, step) => from + step * 5);
  return [
    1,
    ...everyFifth(LAYOUT.incomeFirst + 1, LAYOUT.transferIn - 1),
    ...everyFifth(LAYOUT.expenseFirst + 1, LAYOUT.transferOut - 1),
    LAYOUT.transferOut - 1,
  ];
}
const DIVIDER_WHITE = '#ffffff';
// The receipt OCR (its own Apps Script project) colours the cells it writes: yellow
// OCR_WRITTEN_COLOR (the AI is sure), pink OCR_CHECK_COLOR (a line still to be checked, 2026-10-25),
// green OCR_APPROVED_CELL_COLOR, and mom's hand-typed cells light blue
// OCR_HAND_MATCH_COLOR / light red OCR_HAND_DIFF_COLOR (ocr/ocr-config.gs). Kept as they are.
// (#fff2cc is NOT one of them: it is the tint of the transfer columns T and AJ.)
// The mismatch red was #f4c7c3, too close to a category band (#f4cccc): the OCR moves to #e6b8af,
// the old one is kept until its cells are recoloured.
const DIVIDER_OCR_COLORS = ['#ffe599', '#ff99dd', '#d9ead3', '#c9daf8', '#e6b8af', '#f4c7c3'];
// A count left empty and filled with the sheet's balance by hand (2026-10-22, My home 8/31-9/30,
// note "Тооллого бичигдээгүй байсан") was greyed. Dropped 2026-10-24 (the owner: the counts were
// made, only not typed in): the grey goes from the count column, even where it was approved.
const DIVIDER_COUNT_FILLED_COLOR = '#d0e0e3';
// The transfer link between the shops (transfer-link.gs) colours the transfers it writes pink and
// the ones the two shops disagree on a darker pink. Kept as they are, like the OCR's colours.
// The violet-pink #d5a6bd marked the transfers it once added for a negative balance (dropped
// 2026-10-24); kept so that its last cells go back to the column's colour. A hint (a count short in
// one shop and over in the other, no number written) is lavender #d9d2e9.
const DIVIDER_TRANSFER_COLORS = ['#f8bbd0', '#f06292', '#d5a6bd', '#d9d2e9'];
const DIVIDER_KEPT_COLORS = [...DIVIDER_OCR_COLORS, ...DIVIDER_TRANSFER_COLORS];
// The count column's old grey: never wanted there (column numbers from 1).
const isDroppedCountGrey = (column, color) => column === LAYOUT.count && color === DIVIDER_COUNT_FILLED_COLOR;
const DIVIDER_DESCRIPTION = 'Ангиллын зааглагч - устгахаас хамгаалсан';
const DIVIDER_TIME_BUDGET_MS = 5 * 60 * 1000;

function protectDividers() {
  const { lines } = runProtectDividers(
    currentBook(),
    Date.now() + DIVIDER_TIME_BUDGET_MS
  );
  Logger.log('%s\n%s', modeLabel(), lines.join('\n'));
}

function runProtectDividers(spreadsheet, deadline) {
  const sheets = spreadsheet.getSheets().filter((sheet) => !sheet.getName().startsWith('_'));
  const lines = [];
  for (const sheet of sheets) {
    if (Date.now() > deadline) {
      lines.push('хугацаа дууслаа - дахин ажиллуулбал үргэлжилнэ');
      return { lines, complete: false };
    }
    const line = reconcileSheet(sheet);
    if (line !== null) {
      lines.push(line);
    }
  }
  return { lines: lines.length > 0 ? lines : ['бүх зааглагч хамгаалагдсан байна'], complete: true };
}

function removeDividerProtections() {
  const sheets = currentBook().getSheets();
  const removed = sheets.reduce((total, sheet) => {
    const own = findOwnProtections(sheet);
    own.forEach((protection) => protection.remove());
    return total + own.length;
  }, 0);
  Logger.log('Removed %s divider protections', removed);
}

// ---------- pure helpers (no Spreadsheet calls) ----------

function hasBandColor(rowBackgrounds) {
  const first = rowBackgrounds[0];
  return (
    Boolean(first) &&
    first !== DIVIDER_WHITE &&
    DIVIDER_PROBE_COLUMNS.every((column) => rowBackgrounds[column - 1] === first)
  );
}

// Nothing but (maybe) a name: no balance, price, movement or count in B..AK.
function hasNoNumbers(rowValues) {
  return rowValues.slice(1).every((value) => value === '' || value === null);
}

function isCategoryName(label) {
  return DIVIDER_CATEGORY_NAMES.includes(String(label).trim().replace(/\s+/g, ' ').toLowerCase());
}

// A coloured cell with a known category name and nothing else in the row.
function isNamedCategoryRow(rowBackgrounds, rowValues) {
  const first = rowBackgrounds[0];
  return (
    Boolean(first) && first !== DIVIDER_WHITE && isCategoryName(rowValues[0]) && hasNoNumbers(rowValues)
  );
}

// A divider is a coloured row that has a NAME and no numbers. A row that carries
// numbers is always a product, whatever its colour: rows inserted next to a divider
// inherit its colour. rowValues (A..AK) is optional; without it only the colour can
// be judged.
function isDividerRow(rowBackgrounds, rowValues) {
  if (!Array.isArray(rowValues)) {
    return hasBandColor(rowBackgrounds);
  }
  if (!hasNoNumbers(rowValues) || String(rowValues[0]).trim() === '') {
    return false;
  }
  return hasBandColor(rowBackgrounds) || isNamedCategoryRow(rowBackgrounds, rowValues);
}

function describeRow(row, rowBackgrounds, rowValues) {
  const hasLabel = String(rowValues[0]).trim() !== '';
  const isFullBand = hasBandColor(rowBackgrounds);
  const isNumberless = hasNoNumbers(rowValues);
  return {
    row,
    key: String(rowValues[0]).trim().replace(/\s+/g, ' ').toLowerCase(),
    label: String(rowValues[0]).trim(),
    color: rowBackgrounds[0],
    isFullBand,
    hasLabel,
    isNamed: isNamedCategoryRow(rowBackgrounds, rowValues),
    isCandidate: isDividerRow(rowBackgrounds, rowValues),
    isBandProduct: isFullBand && !isNumberless, // a product coloured across the whole row
    isEmptyBand: isFullBand && isNumberless && !hasLabel, // coloured, but nothing in it
    isEmpty: isNumberless && !hasLabel,
  };
}

// Rows that touch an anchor row (directly, or through each other) and have the
// same colour in column A: they got that colour by being inserted next to it.
function findSameColorNeighbours(rowInfos, anchorRows, isEligible) {
  const byRow = new Map(rowInfos.map((info) => [info.row, info]));
  const anchors = new Set(anchorRows);
  const found = new Set();
  anchorRows.forEach((anchorRow) => {
    const anchor = byRow.get(anchorRow);
    if (!anchor || !anchor.color || anchor.color === DIVIDER_WHITE) {
      return;
    }
    [-1, 1].forEach((step) => {
      let info = byRow.get(anchorRow + step);
      while (info && !anchors.has(info.row) && info.color === anchor.color && isEligible(info)) {
        found.add(info.row);
        info = byRow.get(info.row + step);
      }
    });
  });
  return [...found].sort((a, b) => a - b);
}

// Splits the rows into real dividers and rows that only inherited a colour.
// A divider is a coloured, numberless row: named (band, or a category name in a
// coloured cell), or an EMPTY band standing on its own (the file has such unnamed
// separators). It is not a divider, but an inherited colour, when it
//   - is not a category / not yet protected, and touches a category or a protected
//     divider of the same colour (a new product row whose numbers are not typed yet)
//   - touches a product that is coloured across the whole row in the same colour
// "inherited" lists every row whose colour came from a neighbour; it is for the
// report, the colours themselves are put in order by buildWantedColors.
// knownRows: rows that are already protected as dividers.
function classifyRows(rowInfos, knownRows) {
  const known = new Set(knownRows);
  const rowsOf = (infos) => infos.map((info) => info.row);
  const isCandidate = (info) => info.isCandidate || info.isEmptyBand;
  const bandProducts = rowsOf(rowInfos.filter((info) => info.isBandProduct));
  // An EMPTY coloured row that touches a product coloured across the row in the same
  // colour belongs to that accident, even if an earlier run protected it. A row with
  // a name is never judged by this.
  const besideProducts = new Set(
    findSameColorNeighbours(rowInfos, bandProducts, (info) => info.isEmptyBand)
  );
  const isAnchor = (info) =>
    isCandidate(info) && !besideProducts.has(info.row) && (info.isNamed || known.has(info.row));
  const inheritedCandidates = new Set([
    ...besideProducts,
    ...findSameColorNeighbours(
      rowInfos,
      rowsOf(rowInfos.filter(isAnchor)),
      (info) => isCandidate(info) && !isAnchor(info)
    ),
  ]);
  const dividers = rowInfos.filter((info) => isCandidate(info) && !inheritedCandidates.has(info.row));
  const dividerRows = new Set(rowsOf(dividers));
  const certain = [...bandProducts, ...inheritedCandidates];
  const touching = findSameColorNeighbours(
    rowInfos,
    [...dividerRows, ...certain],
    (info) => !dividerRows.has(info.row)
  );
  const inheritedRows = new Set([...certain, ...touching]);
  return { dividers, inherited: rowInfos.filter((info) => inheritedRows.has(info.row)) };
}

// What has to be done so that a required category sits above its first product:
//   null                                    - it is there (or the product is not on this sheet)
//   { action: 'name', row }                 - the row is there, coloured, but its name is gone
//   { action: 'name-and-color', row }       - an empty uncoloured row stands in its place
//   { action: 'color', row }                - the name is there without its colour
//   { action: 'insert', beforeRow }         - the row was deleted
function planRequiredCategory(rowInfos, category) {
  const wantedKey = category.name.trim().toLowerCase();
  const productIndex = rowInfos.findIndex(
    (info) => info.hasLabel && info.key.startsWith(category.firstProductPrefix)
  );
  if (productIndex === -1) {
    return null;
  }
  const above = productIndex > 0 ? rowInfos[productIndex - 1] : null;
  if (above === null) {
    return { action: 'insert', beforeRow: rowInfos[productIndex].row };
  }
  if (above.key === wantedKey) {
    return above.isCandidate ? null : { action: 'color', row: above.row };
  }
  if (above.isEmptyBand) {
    return { action: 'name', row: above.row };
  }
  if (above.isEmpty) {
    return { action: 'name-and-color', row: above.row };
  }
  return { action: 'insert', beforeRow: rowInfos[productIndex].row };
}

// One sheet as a list of categories, each with the products under it, in order.
//   categories: [{ key, label, color, productKeys }]
//   rowOfProduct: { productKey: row }      infoAbove: { row: the row info above it }
function buildLayout(rowInfos, dividerRows) {
  const isDivider = new Set(dividerRows);
  const start = { categories: [], rowOfProduct: {}, infoAbove: {}, previous: null };
  const layout = rowInfos.reduce((all, info) => {
    const infoAbove = all.previous ? { ...all.infoAbove, [info.row]: all.previous } : all.infoAbove;
    const base = { ...all, infoAbove, previous: info };
    if (isDivider.has(info.row)) {
      return info.hasLabel
        ? { ...base, categories: [...all.categories, { key: info.key, label: info.label, color: info.color, productKeys: [] }] }
        : base;
    }
    if (!info.hasLabel) {
      return base;
    }
    const rowOfProduct = info.key in all.rowOfProduct ? all.rowOfProduct : { ...all.rowOfProduct, [info.key]: info.row };
    if (all.categories.length === 0) {
      return { ...base, rowOfProduct };
    }
    const last = all.categories[all.categories.length - 1];
    const categories = [...all.categories.slice(0, -1), { ...last, productKeys: [...last.productKeys, info.key] }];
    return { ...base, rowOfProduct, categories };
  }, start);
  return { categories: layout.categories, rowOfProduct: layout.rowOfProduct, infoAbove: layout.infoAbove };
}

function mostCommon(values) {
  const counts = values.reduce((all, value) => ({ ...all, [value]: (all[value] || 0) + 1 }), {});
  return Object.entries(counts).sort((a, b) => b[1] - a[1])[0][0];
}

// The categories every sheet of one shop should have: those found on enough sheets
// (or carrying a known category name), with their usual spelling and colour.
function collectSharedCategories(layouts, skippedKeys) {
  const found = layouts.flatMap((layout) => layout.categories);
  const keys = [...new Set(found.map((category) => category.key))];
  return keys
    .filter((key) => !skippedKeys.includes(key))
    .map((key) => {
      const same = found.filter((category) => category.key === key);
      return {
        key,
        sheetCount: same.length,
        label: mostCommon(same.map((category) => category.label)),
        color: mostCommon(same.map((category) => category.color)),
      };
    })
    .filter((category) => category.sheetCount >= DIVIDER_SHARED_MIN_SHEETS || isCategoryName(category.key));
}

// Where a missing category belongs on sheet `index`: above the first of its
// products, as listed on the nearest sheet that has the category (an earlier sheet
// is preferred). null when none of those products is on this sheet.
function planMissingCategory(layouts, index, category) {
  const byDistance = layouts
    .map((layout, position) => ({ layout, position }))
    .filter(({ position }) => position !== index)
    .sort((a, b) => Math.abs(a.position - index) - Math.abs(b.position - index) || a.position - b.position);
  const reference = byDistance
    .map(({ layout }) => layout.categories.find((candidate) => candidate.key === category.key))
    .find((candidate) => candidate !== undefined);
  if (!reference) {
    return null;
  }
  const rows = reference.productKeys
    .map((key) => layouts[index].rowOfProduct[key])
    .filter((row) => row !== undefined);
  if (rows.length === 0) {
    return null;
  }
  const beforeRow = Math.min(...rows);
  const above = layouts[index].infoAbove[beforeRow];
  if (above && above.isEmpty) {
    return { action: 'name-and-color', row: above.row };
  }
  return { action: 'insert', beforeRow };
}

function reconcileRows(existingRows, dividerRows) {
  const existing = new Set(existingRows);
  const wanted = new Set(dividerRows);
  return {
    add: dividerRows.filter((row) => !existing.has(row)),
    remove: existingRows.filter((row) => !wanted.has(row)),
  };
}

// The usual colour of every column: the colour most product rows have there
// (white, or the tint of a whole column such as T and AJ).
function findColumnColors(backgrounds, productIndexes) {
  const width = backgrounds.length > 0 ? backgrounds[0].length : 0;
  return Array.from({ length: width }, (_, column) => {
    const counts = productIndexes.reduce((all, index) => {
      const color = backgrounds[index][column];
      return { ...all, [color]: (all[color] || 0) + 1 };
    }, {});
    // On a tie white wins: a column is only tinted when most of its rows are.
    const ranked = Object.entries(counts).sort(
      (a, b) => b[1] - a[1] || (a[0] === DIVIDER_WHITE ? -1 : b[0] === DIVIDER_WHITE ? 1 : 0)
    );
    return ranked.length > 0 ? ranked[0][0] : DIVIDER_WHITE;
  });
}

// The colours the rows SHOULD have:
//   divider row  - its own colour across the whole row, no gaps
//   other rows   - the usual colour of each column, plus the cells the owner coloured
//                  and the cells the receipt OCR or the transfer link coloured (DIVIDER_KEPT_COLORS)
// approved: { productKey: { column: colour } }
function buildWantedColors(rowInfos, backgrounds, dividerRows, approved) {
  const isDivider = new Set(dividerRows);
  const productIndexes = rowInfos
    .map((info, index) => ({ info, index }))
    .filter(({ info }) => info.hasLabel && !isDivider.has(info.row))
    .map(({ index }) => index);
  const columnColors = findColumnColors(backgrounds, productIndexes);
  return rowInfos.map((info, index) => {
    if (isDivider.has(info.row)) {
      return columnColors.map(() => info.color);
    }
    const own = (info.hasLabel && approved[info.key]) || {};
    return columnColors.map((color, column) => {
      const current = backgrounds[index][column];
      const ownColor = isDroppedCountGrey(column + 1, own[column + 1]) ? undefined : own[column + 1];
      return DIVIDER_KEPT_COLORS.includes(current) ? current : ownColor || color;
    });
  });
}

function countDifferentCells(actual, wanted) {
  return actual.reduce(
    (total, row, index) => total + row.filter((color, column) => color !== wanted[index][column]).length,
    0
  );
}

// The cells of product rows whose colour is not the usual one of their column:
// what the owner coloured on purpose. Returned as { productKey: { column: colour } }.
function collectOwnColors(rowInfos, backgrounds, dividerRows) {
  const isDivider = new Set(dividerRows);
  const plain = buildWantedColors(rowInfos, backgrounds, dividerRows, {});
  return rowInfos.reduce((all, info, index) => {
    if (!info.hasLabel || isDivider.has(info.row)) {
      return all;
    }
    const cells = backgrounds[index].reduce(
      (found, color, column) =>
        color === plain[index][column] || isDroppedCountGrey(column + 1, color) ? found : { ...found, [column + 1]: color },
      {}
    );
    return Object.keys(cells).length > 0 ? { ...all, [info.key]: cells } : all;
  }, {});
}

// ---------- spreadsheet access ----------

// Reads rows 4.. once: values A..AK and the colours of the whole width.
function readRows(sheet) {
  const height = sheet.getLastRow() - DIVIDER_FIRST_ROW + 1;
  if (height <= 0) {
    return { rowInfos: [], backgrounds: [] };
  }
  const values = sheet.getRange(DIVIDER_FIRST_ROW, 1, height, DIVIDER_VALUE_WIDTH).getValues();
  const backgrounds = sheet
    .getRange(DIVIDER_FIRST_ROW, 1, height, Math.max(sheet.getMaxColumns(), DIVIDER_SCAN_WIDTH))
    .getBackgrounds();
  const rowInfos = backgrounds.map((colors, index) =>
    describeRow(DIVIDER_FIRST_ROW + index, colors.slice(0, DIVIDER_SCAN_WIDTH), values[index])
  );
  return { rowInfos, backgrounds };
}

function scanSheet(sheet, own) {
  const { rowInfos, backgrounds } = readRows(sheet);
  const knownRows = own.map((protection) => protection.getRange().getRow());
  return { rowInfos, backgrounds, ...classifyRows(rowInfos, knownRows) };
}

// Every real divider row as { row, color, isFullBand, ... }.
function findDividers(sheet) {
  return scanSheet(sheet, findOwnProtections(sheet)).dividers;
}

function findDividerRows(sheet) {
  return findDividers(sheet).map((divider) => divider.row);
}

// ---------- colours: only the owner's colours stay ----------

function approvedColorsKey(sheet) {
  return DIVIDER_APPROVED_KEY + sheet.getSheetId();
}

function readApprovedColors(sheet) {
  const saved = bookProperties().getProperty(approvedColorsKey(sheet));
  try {
    const parsed = saved ? JSON.parse(saved) : {};
    return parsed !== null && typeof parsed === 'object' ? parsed : {};
  } catch (error) {
    return {};
  }
}

// Is the person who caused this run the owner of the file? Google only tells a
// script who the user is when that user is the owner of the trigger, so for anyone
// else the name comes back empty, and the answer is "no".
function isOwnerActing() {
  try {
    const actor = Session.getActiveUser().getEmail();
    const owner = SpreadsheetApp.getActiveSpreadsheet().getOwner();
    return actor !== '' && owner !== null && actor === owner.getEmail();
  } catch (error) {
    return false;
  }
}

// Remembers the colours the sheet has now as "coloured by the owner". Run when the
// owner changes a colour (instant-sync.gs) or from the menu.
function approveSheetColors(sheet) {
  const own = findOwnProtections(sheet);
  const { rowInfos, backgrounds, dividers } = scanSheet(sheet, own);
  const colors = collectOwnColors(rowInfos, backgrounds, dividers.map((divider) => divider.row));
  if (!DRY_RUN) {
    bookProperties().setProperty(approvedColorsKey(sheet), JSON.stringify(colors));
  }
  return Object.keys(colors).length;
}

function approveActiveSheetColors() {
  if (!isOwnerActing()) {
    notify('Өнгийг зөвхөн файлын эзэн баталж чадна.');
    return;
  }
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const count = approveSheetColors(sheet);
  notify(`"${sheet.getName()}": ${count} барааны мөрийн өнгийг баталлаа. Эдгээр өнгө арилахгүй.`);
}

function findOwnProtections(sheet) {
  return sheet
    .getProtections(SpreadsheetApp.ProtectionType.RANGE)
    .filter((protection) => protection.getDescription() === DIVIDER_DESCRIPTION);
}

function requiredCategoriesOf(sheet) {
  const slot = sheetSlot(sheet.getName());
  return slot === null
    ? []
    : DIVIDER_REQUIRED_CATEGORIES.filter((category) => category.chainKey === slot.chainKey);
}

// Puts back the required categories of this sheet. Returns what was (or would be) done.
function restoreRequiredCategories(sheet, rowInfos) {
  return requiredCategoriesOf(sheet)
    .map((category) => ({ category, plan: planRequiredCategory(rowInfos, category) }))
    .filter(({ plan }) => plan !== null)
    .map(({ category, plan }) => {
      if (!DRY_RUN) {
        applyRequiredCategory(sheet, category, plan);
      }
      return `"${category.name}" ангиллын ${plan.action === 'insert' ? 'мөрийг дахин үүсгэх' : 'нэр, өнгийг сэргээх'}`;
    });
}

// Makes every sheet of a shop carry the same categories: one that is missing on a
// sheet (its row was deleted) is made again above its first product, with the
// colour it has on the other sheets. Returns the lines for the log.
function runCategorySync(spreadsheet) {
  return findChains(spreadsheet).flatMap(({ label: chainKey, sheets }) => {
    const required = DIVIDER_REQUIRED_CATEGORIES.filter((category) => category.chainKey === chainKey);
    const scans = sheets.map((sheet) => readSheetLayout(sheet));
    const layouts = scans.map((scan) => scan.layout);
    // A category with a rule of its own is put back by that rule (reconcileSheet).
    const shared = collectSharedCategories(
      layouts,
      required.map((category) => category.name.trim().toLowerCase())
    );
    return sheets.flatMap((sheet, index) => {
      const present = new Set(layouts[index].categories.map((category) => category.key));
      const missing = shared.filter((category) => !present.has(category.key));
      const plans = missing.map((category) => ({ category, plan: planMissingCategory(layouts, index, category) }));
      // From the bottom up, so that an inserted row does not move the rows still to do.
      const doable = plans
        .filter(({ plan }) => plan !== null)
        .sort((a, b) => (b.plan.beforeRow || b.plan.row) - (a.plan.beforeRow || a.plan.row));
      if (!DRY_RUN) {
        doable.forEach(({ category, plan }) =>
          applyRequiredCategory(sheet, { name: category.label, color: category.color }, plan)
        );
      }
      return plans.map(({ category, plan }) =>
        plan === null
          ? `${sheet.getName()}: "${category.label}" ангилал алга, гэхдээ түүний бараа энэ sheet-д байхгүй тул нэмээгүй`
          : `${sheet.getName()}: алга болсон "${category.label}" ангиллыг ${DRY_RUN ? 'нөхнө' : 'нөхлөө'}`
      );
    });
  });
}

function readSheetLayout(sheet) {
  const scan = scanSheet(sheet, findOwnProtections(sheet));
  return { scan, layout: buildLayout(scan.rowInfos, scan.dividers.map((divider) => divider.row)) };
}

function applyRequiredCategory(sheet, category, plan) {
  if (plan.action === 'insert') {
    sheet.insertRowBefore(plan.beforeRow);
  }
  const row = plan.action === 'insert' ? plan.beforeRow : plan.row;
  const range = sheet.getRange(row, 1, 1, sheet.getMaxColumns());
  if (plan.action === 'insert') {
    range.clearContent();
  }
  if (plan.action !== 'name') {
    range.setBackground(category.color);
  }
  sheet.getRange(row, 1).setValue(category.name);
}

// Brings one sheet in order: required categories there, dividers protected and
// coloured across the whole row, every other colour removed unless the owner put
// it there. Returns a log line, or null when the sheet is already in order.
function reconcileSheet(sheet) {
  const name = sheet.getName();
  if (name === CHART_SOURCE_SHEET || isStockSheet(sheet)) {
    return null; // the yearly summary and the stock sheets have their own colours, no dividers
  }
  const firstScan = scanSheet(sheet, findOwnProtections(sheet));
  const restored = restoreRequiredCategories(sheet, firstScan.rowInfos);
  // A restored category changes rows and names, so the sheet is read again.
  const own = findOwnProtections(sheet);
  const { rowInfos, backgrounds, dividers } =
    restored.length > 0 && !DRY_RUN ? scanSheet(sheet, own) : firstScan;
  const dividerRows = dividers.map((divider) => divider.row);
  const { add, remove } = reconcileRows(
    own.map((protection) => protection.getRange().getRow()),
    dividerRows
  );
  const wanted = buildWantedColors(rowInfos, backgrounds, dividerRows, readApprovedColors(sheet));
  const recolored = countDifferentCells(backgrounds, wanted);
  // Borders cannot be read by a script, so they are removed once per sheet, and
  // after that only on a new divider (and after every colour change, instant-sync.gs).
  const isFirstBorderPass = !hasCleanDividerBorders(sheet);
  const isOldGrid = !isFirstBorderPass && !hasCurrentGrid(sheet);
  const borderRows = isFirstBorderPass || isOldGrid ? dividerRows : add;
  const changes = add.length + remove.length + recolored + restored.length + borderRows.length;
  if (changes === 0 && !isFirstBorderPass && !isOldGrid) {
    return null;
  }
  if (!DRY_RUN) {
    own
      .filter((protection) => remove.includes(protection.getRange().getRow()))
      .forEach((protection) => protection.remove());
    if (recolored > 0) {
      withFilterLifted(sheet, () => sheet.getRange(DIVIDER_FIRST_ROW, 1, wanted.length, wanted[0].length).setBackgrounds(wanted));
    }
    add.forEach((row) => protectRow(sheet, row));
    if (isFirstBorderPass) {
      clearBodyBorders(sheet);
    } else if (isOldGrid) {
      drawTableGrid(sheet); // the daily columns are left as they are
    }
    clearBordersOnRows(sheet, borderRows);
  }
  const strayText = isFirstBorderPass && sheetSlot(name) !== null ? ', E:AJ-ийн илүү хүрээг арилгах' : '';
  const gridText = isOldGrid && sheetSlot(name) !== null ? `, хүснэгтийн хүрээг дахин зурлаа (${gridZonesText()})` : '';
  const restoredText = `${strayText}${gridText}${restored.length > 0 ? `, ${restored.join(', ')}` : ''}`;
  return (
    `${name}: шинэ зааглагч ${add.length}, авсан хамгаалалт ${remove.length} ` +
    `(хэвээр: ${own.length - remove.length}), өнгийг зассан нүд: ${recolored}, ` +
    `хүрээг арилгасан зааглагч: ${borderRows.length}${restoredText}`
  );
}

// Apps Script writes the values but no note or colour into a row a filter hides (seen on the
// sheets with "Зөвхөн буруу": the transfers 2026-10-22, this colour tidy-up 2026-10-24), so the
// filter is taken off while writing and put back with the same criteria. Also used by
// transfer-link.gs.
function withFilterLifted(sheet, write) {
  const filter = sheet.getFilter();
  if (filter === null) {
    write();
    return;
  }
  const range = filter.getRange();
  const criteria = Array.from({ length: range.getNumColumns() }, (_, offset) => range.getColumn() + offset)
    .map((column) => [column, filter.getColumnFilterCriteria(column)])
    .filter(([, criterion]) => criterion !== null);
  filter.remove();
  try {
    write();
  } finally {
    const restored = range.createFilter();
    criteria.forEach(([column, criterion]) => restored.setColumnFilterCriteria(column, criterion));
  }
}

// ---------- borders: a divider is one plain band ----------

function hasCleanDividerBorders(sheet) {
  const saved = bookProperties().getProperty(DIVIDER_BORDERS_KEY + sheet.getSheetId());
  return saved !== null && saved !== undefined;
}

function hasCurrentGrid(sheet) {
  return bookProperties().getProperty(DIVIDER_BORDERS_KEY + sheet.getSheetId()) === DIVIDER_GRID_MARK;
}

// "A:E, BO:BR" for the log.
const gridZonesText = () =>
  DIVIDER_GRID_ZONES.map((zone) => `${columnLetter(zone.firstColumn)}:${columnLetter(zone.firstColumn + zone.width - 1)}`).join(', ');

// Removes the lines between the cells of the given divider rows. The lines above
// and below a divider belong to the product rows next to it and are left alone.
function clearBordersOnRows(sheet, rows) {
  if (rows.length > 0) {
    const lastColumn = sheet.getMaxColumns();
    sheet
      .getRangeList(rows.map((row) => `R${row}C1:R${row}C${lastColumn}`))
      .setBorder(null, null, null, null, false, null);
  }
  bookProperties().setProperty(DIVIDER_BORDERS_KEY + sheet.getSheetId(), DIVIDER_GRID_MARK);
}

// The grid of the table on the given rows: boxes on A..D and AK..AN, and the line
// after T that parts income from expense.
function drawRowGrid(sheet, firstRow, rowCount) {
  DIVIDER_GRID_ZONES.forEach((zone) => {
    sheet
      .getRange(firstRow, zone.firstColumn, rowCount, zone.width)
      .setBorder(true, true, true, true, true, true);
  });
  sheet
    .getRange(firstRow, DIVIDER_SEPARATOR_COLUMN, rowCount, 1)
    .setBorder(null, null, null, true, null, null);
}

// The last row with a name in column A; 0 when there is none.
function findLastNamedRow(sheet) {
  const rows = sheet.getMaxRows() - DIVIDER_FIRST_ROW + 1;
  if (rows <= 0) {
    return 0;
  }
  const names = sheet.getRange(DIVIDER_FIRST_ROW, 1, rows, 1).getValues();
  const last = names.map((cells) => String(cells[0]).trim() !== '').lastIndexOf(true);
  return last === -1 ? 0 : DIVIDER_FIRST_ROW + last;
}

// The empty rows under the last product carry no borders at all.
function clearGridBelow(sheet, firstEmptyRow) {
  const rowCount = sheet.getMaxRows() - firstEmptyRow + 1;
  if (rowCount <= 0) {
    return;
  }
  DIVIDER_GRID_ZONES.forEach((zone) => {
    // The top edge is the bottom line of the last product: left as it is.
    sheet
      .getRange(firstEmptyRow, zone.firstColumn, rowCount, zone.width)
      .setBorder(null, false, false, false, false, false);
  });
  sheet
    .getRange(firstEmptyRow, DIVIDER_SEPARATOR_COLUMN, rowCount, 1)
    .setBorder(null, null, null, false, null, null);
}

// Puts the borders of a year sheet in order: no cell borders in the daily columns
// (boxes left on single cells, lines under a divider), the grid of the table on
// every row down to the last product, and nothing below it. A product typed
// under the last one gets its grid from drawEditedRowBorders.
function clearBodyBorders(sheet) {
  const rows = sheet.getMaxRows() - DIVIDER_FIRST_ROW + 1;
  if (sheetSlot(sheet.getName()) === null || rows <= 0) {
    return;
  }
  const width = DIVIDER_BODY_LAST_COLUMN - DIVIDER_BODY_FIRST_COLUMN + 1;
  sheet
    .getRange(DIVIDER_FIRST_ROW, DIVIDER_BODY_FIRST_COLUMN, rows, width)
    .setBorder(null, null, null, null, false, false);
  drawTableGrid(sheet);
}

// The grid of the table on every row down to the last product, and nothing below it.
function drawTableGrid(sheet) {
  if (sheetSlot(sheet.getName()) === null || sheet.getMaxRows() < DIVIDER_FIRST_ROW) {
    return;
  }
  const lastNamedRow = findLastNamedRow(sheet);
  if (lastNamedRow >= DIVIDER_FIRST_ROW) {
    drawRowGrid(sheet, DIVIDER_FIRST_ROW, lastNamedRow - DIVIDER_FIRST_ROW + 1);
  }
  clearGridBelow(sheet, Math.max(lastNamedRow, DIVIDER_FIRST_ROW - 1) + 1);
}

// Column A was edited. A name typed under the last product: that row gets the grid
// of the table at once (a divider keeps its plain band). The last product cleared:
// its row, now empty at the bottom, loses the grid at once.
function drawEditedRowBorders(sheet, firstRow, lastRow) {
  const from = Math.max(firstRow, DIVIDER_FIRST_ROW);
  if (DRY_RUN || sheetSlot(sheet.getName()) === null || lastRow < from) {
    return;
  }
  const lastNamedRow = findLastNamedRow(sheet);
  const to = Math.min(lastRow, lastNamedRow);
  if (to >= from) {
    drawRowGrid(sheet, from, to - from + 1);
    clearBordersOnRows(
      sheet,
      findDividerRows(sheet).filter((row) => row >= from && row <= to)
    );
  }
  clearGridBelow(sheet, Math.max(lastNamedRow, DIVIDER_FIRST_ROW - 1) + 1);
}

// After a formatting change. Borders cannot be read, so they are simply put in
// order again: always on the dividers, and in the daily columns unless the owner
// made the change (the owner may draw borders there).
function cleanBordersAfterFormat(sheet, isOwner) {
  if (DRY_RUN || sheet.getName() === CHART_SOURCE_SHEET) {
    return;
  }
  if (!isOwner) {
    clearBodyBorders(sheet);
  }
  clearBordersOnRows(sheet, findDividerRows(sheet));
}

function protectRow(sheet, row) {
  const protection = sheet
    .getRange(row, 1, 1, sheet.getMaxColumns())
    .protect()
    .setDescription(DIVIDER_DESCRIPTION);
  protection.setWarningOnly(false);
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
}
