/**
 * Carries product changes forward through the year.
 *
 * Every business ("My home", "Соронзон") has a chain of half-month sheets in tab
 * order. When a product is added to, or removed from, one sheet, the same change
 * is applied to every LATER sheet of the same business:
 *   - added   -> a new row (name, cost, price, formats of the block) is inserted
 *                into the same category block of each later sheet that lacks it;
 *                AL:AN fill themselves from the array formulas (calc-formulas.gs)
 *   - removed -> the row is deleted from each later sheet, unless that row has
 *                income/expense entries (E:AJ): history is never destroyed, the
 *                row is kept and counted in the report
 *
 * How it knows what changed: after each run the product list of every sheet is
 * saved in a hidden sheet "_sync". The next run compares the sheet to its saved
 * list; only that difference is propagated. The very first run only saves the
 * lists (baseline) and changes nothing.
 *
 * Products are matched by name (trimmed, lower-case, single spaces), never by
 * row number, because rows are not aligned between sheets. Category divider
 * rows are recognised with isDividerRow from dividers.gs (keep both files).
 *
 * carryOpeningBalances copies the "тооллого" (AK) of a sheet into the "анхны
 * үлдэгдэл" (B) of the next sheet of the same business, and keeps it that way: as
 * soon as a product has a count, the opening balance on the next sheet IS that
 * count, and another number typed there is put back (at once through
 * instant-sync.gs, and every hour). The log lists what was replaced. Free to type:
 * the opening balance of a product that is not on the sheet before (a new
 * product), of the first sheet of the year, and while nothing is counted yet.
 *
 * Usage: DRY_RUN (config.gs) decides whether anything is really changed.
 *   syncProductsForward   - run by hand (menu "Жилийн sheet")
 *   carryOpeningBalances  - count -> next sheet's opening balance (menu too; at once
 *                           through instant-sync.gs)
 *   runBackgroundJobs     - sync, prices (price-sync.gs), opening balances, new dividers, tab order and the
 *                           row limit of each shop, run by the trigger
 *   installAutoSync       - run once (setupYear does it): every SYNC_TRIGGER_HOURS hour(s)
 *   removeAutoSync        - removes that trigger
 * Needs config.gs (findChains) and dividers.gs.
 */

const SYNC_HEADER_ROWS = 3;
const SYNC_NAME_COL = 1; // A
const SYNC_OPENING_COL = 2; // B
const SYNC_COST_COL = 3; // C
const SYNC_PRICE_COL = 4; // D
const SYNC_SHELF_PRICE_COL = LAYOUT.shelfPrice; // E in the shops' layout (typed by hand), else null
const SYNC_DESCRIBE_LAST_COL = SYNC_SHELF_PRICE_COL || SYNC_PRICE_COL; // A..D, or A..E
const SYNC_MOVEMENT_FIRST_COL = LAYOUT.incomeFirst; // E
const SYNC_MOVEMENT_LAST_COL = LAYOUT.expenseLast; // AJ (with the gift column: AL)
const SYNC_COUNT_COL = LAYOUT.count; // AK
const SYNC_TEMPLATE_LAST_COL = LAYOUT.count; // AK; AL:AN are filled by the array formulas of calc-formulas.gs
const SYNC_MIN_PRICED_PRODUCTS = 20; // sheets with fewer priced products are not entry sheets
const SYNC_STATE_SHEET = '_sync';
const CHECK_NAMES_MAIL_SUBJECT = 'Барааны нэрийн шалгалт';
const NAME_WATCH_KEY = 'nameWatch'; // script property: fingerprint of the report already mailed
const CARRY_MAX_LOG_LINES = 100; // replaced opening balances listed in the log
const SYNC_TIME_BUDGET_MS = 4.5 * 60 * 1000;
const SYNC_LOCK_WAIT_MS = 30 * 1000;
// Hourly: Google allows triggers roughly 90 minutes of run time per day and stops
// a run at 6 minutes. So one hourly run starts no new task after
// BACKGROUND_TIME_BUDGET_MS (24 runs x 2.5 minutes = 60 minutes a day, the rest is
// left for the instant triggers), and the next run goes on with the task after it.
const SYNC_TRIGGER_HOURS = 1;
const BACKGROUND_TIME_BUDGET_MS = 2.5 * 60 * 1000;
const BACKGROUND_NEXT_KEY = 'backgroundNext'; // script property: the task the next run starts with
// Each task: (spreadsheet, deadline) => log lines. Dividers come before the product
// sync: a product row that inherited a divider's colour must be cleaned before the
// product lists are read.
const BACKGROUND_TASKS = [
  (spreadsheet, deadline) => [
    ...runCategorySync(spreadsheet),
    ...runProtectDividers(spreadsheet, deadline).lines,
  ],
  (spreadsheet, deadline) => [formatSyncReport(runSync(spreadsheet, deadline))],
  (spreadsheet) => [runCarry(spreadsheet)],
  (spreadsheet) => [formatPriceReport(runPriceSync(spreadsheet))],
  (spreadsheet) => runArrangeTabs(spreadsheet),
  (spreadsheet) => runTrim(spreadsheet).filter((line) => line.includes('->')),
  (spreadsheet) => runRepairCalcColumns(spreadsheet),
  (spreadsheet) => [
    ...runNumberRepeated(spreadsheet, positionOfDate(new Date())),
    ...[runInstallNameRules(spreadsheet)].filter((line) => !line.endsWith(': 0 sheet-д тавьлаа')), // the list grew
  ],
  (spreadsheet) => [
    ...runNameWatch(spreadsheet),
    ...runListRenamesFromSource(spreadsheet), // product-list.gs: a copy of the list follows its renames
    ...runListWatch(spreadsheet),
  ],
  (spreadsheet) => runRestoreHeaders(spreadsheet),
  // official-names.gs (the main workbook only): typed rows to their place in name order, at night
  (spreadsheet) => (typeof runNightlyNameOrder === 'function' ? runNightlyNameOrder(spreadsheet) : []),
  (spreadsheet) => refreshCountSheets(spreadsheet), // count-sheets.gs: a new half-month, a missed change
  // transfer-link.gs (the main workbook only): the shops' transfers on both sides, counts matched
  (spreadsheet) => (typeof runRecentTransferLink === 'function' ? runRecentTransferLink(spreadsheet) : []),
];
const SYNC_BACKGROUND_HANDLER = 'runBackgroundJobs';
const SYNC_TRIGGER_HANDLERS = [SYNC_BACKGROUND_HANDLER];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Жилийн sheet')
    .addItem('Шинэ, хасагдсан барааг дараагийн sheet-үүд рүү дамжуулах', 'syncProductsForward')
    .addItem('Үнэ, өртгийн өөрчлөлтийг дараагийн sheet-үүд рүү дамжуулах', 'syncPricesForward')
    .addItem('Тооллогыг дараагийн sheet-ийн анхны үлдэгдэл рүү дамжуулах', 'carryOpeningBalances')
    .addItem('Энэ sheet-ийн өнгийг батлах (зөвхөн эзэн)', 'approveActiveSheetColors')
    .addToUi();
}

function syncProductsForward() {
  withScriptLock(() => notify(formatSyncReport(runSync(currentBook()))));
}

function carryOpeningBalances() {
  withScriptLock(() => notify(runCarry(currentBook())));
}

// Runs every SYNC_TRIGGER_HOURS hour(s) as the file owner (installAutoSync), for today's
// workbook and, in January, last year's too (booksToMaintain, year-books.gs); then the new
// year's files are prepared when it is time (year-rollover.gs).
function runBackgroundJobs() {
  withScriptLock(() => {
    const deadline = Date.now() + BACKGROUND_TIME_BUDGET_MS;
    const lines = booksToMaintain(new Date()).flatMap((book) => runBookJobs(useBook(book), deadline));
    if (typeof runYearRollover === 'function') {
      lines.push(...runYearRollover(new Date(), deadline));
    }
    Logger.log(lines.join('\n'));
  });
}

// The hourly tasks of one workbook, in order, starting with the one the last run did not get
// to; no new task starts once the time budget is used up.
function runBookJobs(spreadsheet, deadline) {
  if (!requirePairedLayout('Цагийн ажил')) {
    return [];
  }
  if (typeof isTidyInProgress === 'function' && isTidyInProgress()) {
    return [`${spreadsheet.getName()}: ${TIDY_BUSY_TEXT}`]; // official-names.gs finishes first, then the hourly tasks
  }
  const order = taskOrder(BACKGROUND_TASKS.length, readBackgroundStart());
  const lines = [];
  let done = 0;
  while (done < order.length && (done === 0 || Date.now() <= deadline)) {
    lines.push(...BACKGROUND_TASKS[order[done]](spreadsheet, deadline));
    done += 1;
  }
  const left = order.length - done;
  saveBackgroundStart(left === 0 ? 0 : order[done]);
  if (left > 0) {
    lines.push(`хугацаа хэмнэхээр ${left} ажлыг дараагийн удаад үлдээлээ`);
  }
  return lines;
}

function readBackgroundStart() {
  const saved = Number(bookScriptProperties().getProperty(BACKGROUND_NEXT_KEY));
  return Number.isInteger(saved) && saved >= 0 && saved < BACKGROUND_TASKS.length ? saved : 0;
}

function saveBackgroundStart(index) {
  if (!DRY_RUN) {
    bookScriptProperties().setProperty(BACKGROUND_NEXT_KEY, String(index));
  }
}

function installAutoSync() {
  removeAutoSync();
  ScriptApp.newTrigger(SYNC_BACKGROUND_HANDLER)
    .timeBased()
    .everyHours(SYNC_TRIGGER_HOURS)
    .create();
}

function removeAutoSync() {
  ScriptApp.getProjectTriggers()
    .filter((trigger) => SYNC_TRIGGER_HANDLERS.includes(trigger.getHandlerFunction()))
    .forEach((trigger) => ScriptApp.deleteTrigger(trigger));
}

// ---------- pure helpers (no Spreadsheet calls) ----------

// The task numbers 0..count-1 in the order a run does them: from `start`, then around.
function taskOrder(count, start) {
  return Array.from({ length: count }, (_, offset) => (start + offset) % count);
}

function normalizeKey(label) {
  return String(label).trim().replace(/\s+/g, ' ').toLowerCase();
}

function diffKeys(previousKeys, currentKeys) {
  const previous = new Set(previousKeys);
  const current = new Set(currentKeys);
  return {
    added: currentKeys.filter((key) => !previous.has(key)),
    removed: previousKeys.filter((key) => !current.has(key)),
  };
}

// The row a new product goes after: in its category at its place in name order (official-names.gs),
// at the end of the category without a label; at the very end without a category.
function findInsertAfterRow(model, categoryKey, label) {
  const category = model.categories.find((entry) => entry.key === categoryKey);
  if (!category) {
    return Math.max(SYNC_HEADER_ROWS, ...model.products.map((product) => product.row));
  }
  const block = model.products.filter((product) => product.categoryKey === categoryKey);
  if (label === undefined || block.length === 0 || typeof rowBeforeInOrder !== 'function') { // Тосгон: no official-names.gs
    return Math.max(category.row, ...block.map((product) => product.row));
  }
  const before = rowBeforeInOrder(block, label);
  return before === null ? Math.min(...block.map((product) => product.row)) - 1 : before;
}

function pickTemplateRow(model, categoryKey) {
  const sameBlock = model.products.filter((product) => product.categoryKey === categoryKey);
  const pool = sameBlock.length > 0 ? sameBlock : model.products;
  return pool.length > 0 ? pool[pool.length - 1].row : null;
}

function applyKeyChange(snapshot, { add, remove }) {
  const removing = new Set(remove);
  return [...new Set([...snapshot, ...add])].filter((key) => !removing.has(key));
}

// ---------- reading the spreadsheet ----------

function readModel(sheet) {
  const height = sheet.getLastRow() - SYNC_HEADER_ROWS;
  if (height <= 0) {
    return { products: [], categories: [] };
  }
  const firstRow = SYNC_HEADER_ROWS + 1;
  const labels = sheet.getRange(firstRow, SYNC_NAME_COL, height, DIVIDER_VALUE_WIDTH).getValues();
  const colors = sheet.getRange(firstRow, 1, height, DIVIDER_SCAN_WIDTH).getBackgrounds();
  const products = [];
  const categories = [];
  let currentCategory = null;
  labels.forEach((cells, index) => {
    const label = String(cells[0]).trim();
    if (label === '') {
      return;
    }
    const row = firstRow + index;
    if (isDividerRow(colors[index], cells)) {
      currentCategory = { key: normalizeKey(label), label, row };
      categories.push(currentCategory);
      return;
    }
    products.push({
      key: normalizeKey(label),
      label,
      row,
      categoryKey: currentCategory ? currentCategory.key : null,
      opening: cells[SYNC_OPENING_COL - 1],
      cost: cells[SYNC_COST_COL - 1],
      price: cells[SYNC_PRICE_COL - 1],
      shelfPrice: SYNC_SHELF_PRICE_COL ? cells[SYNC_SHELF_PRICE_COL - 1] : '',
      count: cells[SYNC_COUNT_COL - 1],
      net: netMovement(cells),
    });
  });
  return { products, categories };
}

// Income minus expense of one product row (cells = A..count of the row). Text and empty cells count as 0.
function netMovement(cells) {
  const total = (first, last) =>
    cells.slice(first - 1, last).reduce((sum, value) => sum + (typeof value === 'number' ? value : 0), 0);
  return total(LAYOUT.incomeFirst, LAYOUT.transferIn) - total(LAYOUT.expenseFirst, LAYOUT.expenseLast);
}

// Reading a sheet to answer this takes a moment and the answer is asked for many
// times in one run, so it is remembered per sheet for the length of the run.
const ENTRY_SHEET_CACHE = new Map();

function isEntrySheet(sheet) {
  const id = `${sheet.getParent().getId()}:${sheet.getSheetId()}`; // two workbooks may share a sheet id
  if (!ENTRY_SHEET_CACHE.has(id)) {
    ENTRY_SHEET_CACHE.set(id, hasEnoughPricedProducts(sheet));
  }
  return ENTRY_SHEET_CACHE.get(id);
}

function hasEnoughPricedProducts(sheet) {
  const height = sheet.getLastRow() - SYNC_HEADER_ROWS;
  if (height <= 0) {
    return false;
  }
  const prices = sheet.getRange(SYNC_HEADER_ROWS + 1, SYNC_PRICE_COL, height, 1).getValues();
  return prices.filter(([value]) => typeof value === 'number').length >= SYNC_MIN_PRICED_PRODUCTS;
}

function hasMovement(sheet, row) {
  const width = SYNC_MOVEMENT_LAST_COL - SYNC_MOVEMENT_FIRST_COL + 1;
  return sheet
    .getRange(row, SYNC_MOVEMENT_FIRST_COL, 1, width)
    .getValues()[0]
    .some((value) => value !== '');
}

// ---------- snapshots ----------

function loadSnapshots(spreadsheet, stateSheetName = SYNC_STATE_SHEET) {
  const stateSheet = spreadsheet.getSheetByName(stateSheetName);
  if (!stateSheet || stateSheet.getLastRow() === 0) {
    return {};
  }
  return stateSheet
    .getRange(1, 1, stateSheet.getLastRow(), 2)
    .getValues()
    .reduce((all, [name, json]) => ({ ...all, [name]: JSON.parse(json) }), {});
}

function saveSnapshots(spreadsheet, snapshots, stateSheetName = SYNC_STATE_SHEET) {
  const stateSheet =
    spreadsheet.getSheetByName(stateSheetName) || spreadsheet.insertSheet(stateSheetName);
  stateSheet.clear();
  const rows = Object.entries(snapshots).map(([name, keys]) => [name, JSON.stringify(keys)]);
  if (rows.length > 0) {
    stateSheet.getRange(1, 1, rows.length, 2).setValues(rows);
  }
  stateSheet.hideSheet();
}

// ---------- the sync ----------

function runSync(spreadsheet, deadline = Date.now() + SYNC_TIME_BUDGET_MS, scope) {
  if (typeof isTidyInProgress === 'function' && isTidyInProgress()) {
    // official-names.gs is renaming sheet by sheet: a half-renamed year would look like removed and added products
    return { lines: [TIDY_BUSY_TEXT], totals: { added: 0, removed: 0, kept: 0 } };
  }
  let snapshots = loadSnapshots(spreadsheet);
  const savedState = JSON.stringify(snapshots);
  let totals = { added: 0, removed: 0, kept: 0 };
  const lines = [];

  findChains(spreadsheet, scope).forEach(({ label, sheets }) => {
    lines.push(`${label}: ${sheets.map((sheet) => sheet.getName()).join(' > ')}`);
    sheets.forEach((sheet, index) => {
      const name = sheet.getName();
      const keys = readModel(sheet).products.map((product) => product.key);
      if (!snapshots[name]) {
        snapshots = { ...snapshots, [name]: keys };
        lines.push(`  ${name}: барааны жагсаалт хадгалагдлаа`);
        return;
      }
      const change = diffKeys(snapshots[name], keys);
      if (change.added.length === 0 && change.removed.length === 0) {
        return;
      }
      lines.push(`  ${name}: +${change.added.length} -${change.removed.length}`);
      const result = propagate(sheet, sheets.slice(index + 1), change, deadline);
      if (result.offList.length > 0) {
        lines.push(`  ${result.offList.map((label) => `"${label}"`).join(', ')}: "${PRODUCT_LIST_TAB}" жагсаалтад алга тул дараагийн sheet-үүдэд нэмсэнгүй`);
      }
      totals = result.perTarget.reduce(addTotals, totals);
      if (!DRY_RUN) {
        snapshots = updateSnapshots(snapshots, result.perTarget);
        if (result.complete) {
          snapshots = { ...snapshots, [name]: keys };
        } else {
          lines.push('  хугацаа дууслаа - үлдсэн нь дараагийн удаа үргэлжилнэ');
        }
      }
    });
  });

  // Written only when something changed: every write to the file wakes other
  // triggers (the web-site sync), and this runs every few minutes.
  if (JSON.stringify(snapshots) !== savedState) {
    saveSnapshots(spreadsheet, snapshots);
  }
  return { lines, totals };
}

// A new product goes on only when its name is on the list "Бараа" (listNameKeys, product-list.gs):
// one the list lacks stays on the sheet it was typed in and is reported (offList). Removals are not
// judged by the list: a list that could not be read must never take rows away.
function propagate(source, targets, change, deadline) {
  const sourceModel = readModel(source);
  const addedKeys = new Set(change.added);
  const listed = listNameKeys(source.getParent());
  const added = sourceModel.products.filter((product) => addedKeys.has(product.key));
  const offList = added.filter((product) => listed !== null && !listed.has(product.key)).map((product) => product.label);
  const items = added
    .filter((product) => listed === null || listed.has(product.key))
    .map((product) => describeProduct(source, product));
  const removedKeys = new Set(change.removed);
  const perTarget = [];
  for (const target of targets) {
    if (Date.now() > deadline) {
      return { complete: false, perTarget, offList };
    }
    perTarget.push({ name: target.getName(), ...applyChange(target, items, removedKeys) });
  }
  return { complete: true, perTarget, offList };
}

function describeProduct(source, product) {
  const [label, , cost, price, shelfPrice = ''] = source.getRange(product.row, 1, 1, SYNC_DESCRIBE_LAST_COL).getValues()[0];
  return { key: product.key, label: String(label).trim(), cost, price, shelfPrice, categoryKey: product.categoryKey };
}

function applyChange(target, items, removedKeys) {
  const model = readModel(target);
  const present = new Set(model.products.map((product) => product.key));
  const toAdd = items.filter((item) => !present.has(item.key));
  const toRemove = model.products.filter((product) => removedKeys.has(product.key));
  const deletable = toRemove.filter((product) => !hasMovement(target, product.row));
  // A sheet showing "Зөвхөн буруу" (view-columns.gs) has a filter, and Google refuses to insert or
  // delete rows under it ("not supported on a range with a filtered out row"): the hourly job failed
  // on every run while one was set (2026-10-25). The filter is taken off and put back the same.
  if (!DRY_RUN && (deletable.length > 0 || toAdd.length > 0)) {
    withoutFilter(target, () => {
      [...deletable].sort((a, b) => b.row - a.row).forEach((product) => target.deleteRow(product.row));
      toAdd.forEach((item) => insertProduct(target, readModel(target), item));
    });
  }
  return {
    added: toAdd.map((item) => item.key),
    alreadyThere: items.filter((item) => present.has(item.key)).map((item) => item.key),
    removed: deletable.map((product) => product.key),
    kept: toRemove.length - deletable.length,
  };
}

function insertProduct(sheet, model, item) {
  const afterRow = findInsertAfterRow(model, item.categoryKey, item.label);
  const templateRow = pickTemplateRow(model, item.categoryKey);
  sheet.insertRowAfter(afterRow);
  const newRow = afterRow + 1;
  if (templateRow !== null) {
    const shiftedTemplate = templateRow > afterRow ? templateRow + 1 : templateRow;
    sheet
      .getRange(shiftedTemplate, 1, 1, SYNC_TEMPLATE_LAST_COL)
      .copyTo(sheet.getRange(newRow, 1, 1, SYNC_TEMPLATE_LAST_COL));
    clearCopiedMarks(sheet, newRow, SYNC_TEMPLATE_LAST_COL); // the template's notes are no product's of this row (row-copies.gs)
  }
  sheet.getRange(newRow, 1, 1, SYNC_COUNT_COL).clearContent();
  sheet.getRange(newRow, SYNC_NAME_COL).setValue(item.label);
  sheet.getRange(newRow, SYNC_COST_COL, 1, 2).setValues([[item.cost, item.price]]);
  if (SYNC_SHELF_PRICE_COL) {
    sheet.getRange(newRow, SYNC_SHELF_PRICE_COL).setValue(item.shelfPrice === undefined ? '' : item.shelfPrice);
  }
}

function updateSnapshots(snapshots, perTarget) {
  return perTarget.reduce((all, result) => {
    if (!all[result.name]) {
      return all;
    }
    const next = applyKeyChange(all[result.name], {
      add: [...result.added, ...result.alreadyThere],
      remove: result.removed,
    });
    return { ...all, [result.name]: next };
  }, snapshots);
}

function addTotals(totals, result) {
  return {
    added: totals.added + result.added.length,
    removed: totals.removed + result.removed.length,
    kept: totals.kept + result.kept,
  };
}

function formatSyncReport({ lines, totals }) {
  const mode = DRY_RUN ? '[ТУРШИЛТ - юу ч өөрчлөгдөөгүй] ' : '';
  const summary =
    `${mode}Нэмэгдсэн: ${totals.added}, хасагдсан: ${totals.removed}, ` +
    `хөдөлгөөнтэй тул хасаагүй: ${totals.kept}`;
  Logger.log([summary, ...lines].join('\n'));
  return summary;
}

// ---------- opening balances ----------

// What the next sheet's opening balance has to be, given the product on the sheet
// before: its count when it was counted. Where it was not counted, a workbook with
// CARRY_FROM_BALANCE (profile.gs) takes the computed final balance (opening + income -
// expense); a product with no opening balance and no movement stays empty. Anywhere
// else '' = nothing to carry. isUncounted: the sheet before is one of UNCOUNTED_SHEETS, its
// count column copied balances, so its balance is carried whatever that column holds.
function carryValue(previous, isUncounted = false) {
  if (!isUncounted && typeof previous.count === 'number') {
    return previous.count;
  }
  if (!CARRY_FROM_BALANCE && !isUncounted) {
    return '';
  }
  const balance = (typeof previous.opening === 'number' ? previous.opening : 0) + previous.net;
  return balance === 0 && previous.opening === '' ? '' : balance;
}

// What to do with one opening balance (B), given the count (AK) of the same
// product on the sheet before.
//   'write' - the count is a number and B is empty or holds another number
//   'same'  - nothing to do: B equals the count, or nothing was counted yet
function decideOpening(opening, count) {
  return typeof count === 'number' && opening !== count ? 'write' : 'same';
}

// scope (optional) limits the run to one business and a few periods, see findChains.
function runCarry(spreadsheet, scope) {
  let totals = { written: 0, replaced: [] };
  findChains(spreadsheet, scope).forEach(({ sheets }) => {
    const models = sheets.map((sheet) => ({ sheet, name: sheet.getName(), ...readModel(sheet) }));
    // Each sheet is carried into with the openings of the sheet before AS THEY ARE NOW, so a
    // change of one month reaches every later month in the same run.
    models.slice(1).reduce((previous, current) => {
      const result = carryInto(previous.products, current, UNCOUNTED_SHEETS.includes(previous.name));
      totals = {
        written: totals.written + result.written,
        replaced: [...totals.replaced, ...result.replaced],
      };
      return { name: current.name, products: result.products };
    }, models[0]);
  });
  const mode = DRY_RUN ? '[ТУРШИЛТ - юу ч өөрчлөгдөөгүй] ' : '';
  const summary =
    `${mode}Анхны үлдэгдэл: тооллогоос авсан ${totals.written}, ` +
    `үүнээс гараар бичсэнийг зассан ${totals.replaced.length}`;
  const shown = totals.replaced.slice(0, CARRY_MAX_LOG_LINES);
  const hidden = totals.replaced.length - shown.length;
  Logger.log([summary, ...shown, ...(hidden > 0 ? [`... өөр ${hidden} мөр`] : [])].join('\n'));
  return summary;
}

// A price that tells same-named rows apart: the sale price, or the cost when no
// sale price is written. null when the row has neither.
function tellingPrice(product) {
  if (typeof product.price === 'number') {
    return product.price;
  }
  return typeof product.cost === 'number' ? product.cost : null;
}

// Which row of the sheet before each row of this sheet continues: the index into
// previous, or -1 for a product that is new here. A name that stands once on both
// sheets is simply paired. Rows that share a name are paired by the nearest price
// (a product keeps about the same price from one half-month to the next, while
// two different products of one name differ in price); rows without a price, and
// equal distances, go by their order on the sheet.
function pairProducts(previous, current) {
  const indexesByKey = (products) =>
    products.reduce((all, product, at) => ({ ...all, [product.key]: [...(all[product.key] || []), at] }), {});
  const before = indexesByKey(previous);
  const now = indexesByKey(current);
  const pairs = current.map(() => -1);
  Object.keys(now).forEach((key) => {
    const candidates = (before[key] || []).flatMap((from, fromOrder) =>
      now[key].map((to, toOrder) => {
        const a = tellingPrice(previous[from]);
        const b = tellingPrice(current[to]);
        return {
          from,
          to,
          distance: a === null || b === null ? Number.MAX_SAFE_INTEGER : Math.abs(a - b),
          orderGap: Math.abs(fromOrder - toOrder),
        };
      })
    );
    const taken = new Set();
    [...candidates]
      .sort((x, y) => x.distance - y.distance || x.orderGap - y.orderGap)
      .forEach((candidate) => {
        if (pairs[candidate.to] === -1 && !taken.has(candidate.from)) {
          pairs[candidate.to] = candidate.from;
          taken.add(candidate.from);
        }
      });
  });
  return pairs;
}

// previous: the products of the sheet before. A product that is not on the sheet
// before is new here: its opening balance is typed by hand and left alone.
function carryInto(previous, current, isUncounted = false) {
  let written = 0;
  const replaced = [];
  const pairs = pairProducts(previous, current.products);
  const products = current.products.map((product, at) => {
    if (pairs[at] === -1) {
      return product;
    }
    const carried = carryValue(previous[pairs[at]], isUncounted);
    if (decideOpening(product.opening, carried) !== 'write') {
      return product;
    }
    written += 1;
    if (product.opening !== '') {
      replaced.push(`${current.name}: ${product.label} ${product.opening} → ${carried}`);
    }
    if (!DRY_RUN) {
      current.sheet.getRange(product.row, SYNC_OPENING_COL).setValue(carried);
    }
    return { ...product, opening: carried };
  });
  return { written, replaced, products };
}

// ---------- name check (reads only) ----------

// What differs between the product names of a sheet and the sheet before it.
// previous / current: [{ key, label }]. A name that differs by one letter is a
// different product for every sync, so its opening balance is not carried.
function compareNames(previous, current) {
  const keysOf = (products) => products.reduce((all, product) => ({ ...all, [product.key]: true }), {});
  const before = keysOf(previous);
  const now = keysOf(current);
  const counts = current.reduce((all, product) => ({ ...all, [product.key]: (all[product.key] || 0) + 1 }), {});
  const labels = (products) => [...new Set(products.map((product) => product.label.trim()))];
  return {
    gone: labels(previous.filter((product) => !(product.key in now))),
    fresh: labels(current.filter((product) => !(product.key in before))),
    repeated: labels(current.filter((product) => counts[product.key] > 1)),
  };
}

function runCheckNames(spreadsheet, fromPosition = 0) {
  const lines = findNameDifferences(spreadsheet, fromPosition);
  return lines.length > 0 ? lines : ['бүх sheet-ийн барааны нэр өмнөх sheet-тэйгээ таарч байна'];
}

// fromPosition: only the periods from this one on.
function findNameDifferences(spreadsheet, fromPosition) {
  const lines = [];
  findChains(spreadsheet).forEach(({ sheets }) => {
    const models = sheets.map((sheet) => ({ name: sheet.getName(), products: readModel(sheet).products }));
    models.forEach((current, index) => {
      const diff = compareNames(index > 0 ? models[index - 1].products : current.products, current.products);
      const parts = [
        diff.gone.length > 0 ? `өмнөх sheet-д байсан, энд алга: ${diff.gone.join(', ')}` : '',
        diff.fresh.length > 0 ? `энд шинэ: ${diff.fresh.join(', ')}` : '',
        diff.repeated.length > 0 ? `ижил нэртэй хэд хэдэн мөр: ${diff.repeated.join(', ')}` : '',
      ].filter((part) => part !== '');
      if (parts.length > 0 && sheetSlot(current.name).position >= fromPosition) {
        lines.push(`${current.name}\n  ${parts.join('\n  ')}`);
      }
    });
  });
  return lines;
}

// A short fingerprint of a text, to see whether a report is the one already sent.
function hashText(text) {
  return String(
    [...text].reduce((hash, letter) => (hash * 31 + letter.codePointAt(0)) % 2147483647, text.length)
  );
}

// Hourly: the sheets after the current half-month must hold the same names as the
// sheet before them (the syncs see to that). When they do not, or a name is there
// twice, the owner gets an e-mail - once per finding, not every hour.
function runNameWatch(spreadsheet, today = new Date()) {
  const lines = findNameDifferences(spreadsheet, positionOfDate(today) + 1);
  const properties = bookScriptProperties(spreadsheet);
  const fingerprint = lines.length > 0 ? hashText(lines.join('\n')) : '';
  if (fingerprint === (properties.getProperty(NAME_WATCH_KEY) || '')) {
    return [];
  }
  if (DRY_RUN) {
    return [`нэрийн зөрүү ${lines.length} sheet дээр: email явуулна`];
  }
  if (lines.length === 0) {
    properties.deleteProperty(NAME_WATCH_KEY);
    return ['нэрийн зөрүү арилсан'];
  }
  MailApp.sendEmail(
    Session.getEffectiveUser().getEmail(),
    CHECK_NAMES_MAIL_SUBJECT,
    `"${spreadsheet.getName()}"\n${spreadsheet.getUrl()}\n\n${lines.join('\n')}`
  );
  properties.setProperty(NAME_WATCH_KEY, fingerprint);
  return [`нэрийн зөрүү ${lines.length} sheet дээр: email явууллаа`];
}

// The period a date falls in: 0 = the first sheet of the chain (January 1-15 for the main
// workbook). A date before the chain gives a negative number, one after it a number past the end.
function positionOfDate(date) {
  const half = PERIODS_PER_MONTH === 2 && date.getDate() > FIRST_HALF_LAST_DAY ? 1 : 0;
  return periodPosition(date.getFullYear(), date.getMonth() + 1, half);
}

// Run by hand: lists, sheet by sheet, the names that do not match the sheet before.
// The whole year is too long for the log, so it goes to the owner by e-mail; the
// log shows the sheets from the current half-month on, the ones that still matter.
function checkProductNames() {
  const spreadsheet = currentBook();
  MailApp.sendEmail(
    Session.getEffectiveUser().getEmail(),
    CHECK_NAMES_MAIL_SUBJECT,
    `"${spreadsheet.getName()}"\n${spreadsheet.getUrl()}\n\n${runCheckNames(spreadsheet).join('\n')}`
  );
  const recent = runCheckNames(spreadsheet, positionOfDate(new Date()));
  Logger.log(['Бүтэн жагсаалтыг email-ээр явууллаа. Одоогийн sheet-ээс хойш:', ...recent].join('\n'));
}

// ---------- plumbing ----------

// Returns false when another run holds the lock and the action was not done.
function withScriptLock(action) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(SYNC_LOCK_WAIT_MS)) {
    notify('Өөр шинэчлэлт явагдаж байна. Түр хүлээгээд дахин оролдоно уу.');
    return false;
  }
  try {
    action();
  } finally {
    lock.releaseLock();
  }
  return true;
}

function notify(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (error) {
    Logger.log(message); // time-driven runs have no UI
  }
}
