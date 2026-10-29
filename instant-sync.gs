/**
 * Carries a change to the later sheets right away, instead of waiting for the
 * hourly background job. Two triggers do it:
 *
 * "On edit" -> onSheetEdit, after every typed or pasted value. A name corrected in
 * the product list ("Бараа") is corrected in both shops (product-list.gs).
 * Otherwise it acts only when the edit is in a product row of a year sheet and touches
 *   column A      a product was added             -> product sync (year-sync.gs);
 *                 one name was corrected          -> the same row is renamed on the
 *                 later sheets, its numbers stay (renameOnLaterSheets); a name
 *                 that is already on the sheet gets a number, "Haier 12 (2)":
 *                 no two rows of one name (name-fixes.gs)
 *                 a category name was typed       -> that sheet's dividers;
 *                 the row gets the borders of the table
 *   column C / D  a cost or price was changed     -> price sync (price-sync.gs)
 *   AK            a count was typed or corrected  -> the next sheet's opening
 *                 balance (B) follows (year-sync.gs)
 *   column B      an opening balance was typed    -> put back to the count of
 *                 the sheet before, when that product was counted there
 *   AL / AM / AN  something was typed over a formula column -> it is removed and
 *                 the formula fills the column again (calc-formulas.gs)
 *   rows 1-3      a title or label was changed or emptied -> it is put back
 *                 (header-guard.gs)
 *   A1            the list of views -> the income, expense, price or total
 *                 columns are hidden or shown (view-columns.gs)
 *
 * "On change" -> onInstantChange, after a change that is not a typed value:
 *   a row was deleted        -> product sync (the product leaves the later sheets)
 *   a row was inserted       -> a colour inherited from a divider is removed and
 *                               that sheet is cut back to the row limit
 *   a colour was changed     -> the owner's colours are remembered, anyone else's
 *                               are removed, dividers are protected, coloured
 *                               across the whole row and lose their cell borders
 *   a tab was moved or added -> the tab order is restored
 *
 * To be quick, only the business that was changed is read, from the changed sheet
 * onward: a price changed in October reads the 6 sheets up to December, not all 48.
 *
 * Several quick changes in a row do not start several product syncs: a change
 * that arrives while one is working only leaves a note (INSTANT_PENDING_KEY), and
 * the running one goes round again. Whatever is missed is done by the hourly job.
 *
 * installInstantSync is run once (setupYear does it); removeInstantSync undoes it.
 * Needs config.gs, year-sync.gs, price-sync.gs, dividers.gs, tab-order.gs, trim-rows.gs.
 */

const INSTANT_HANDLER = 'onSheetEdit';
const INSTANT_CHANGE_HANDLER = 'onInstantChange';
// A name, C cost, D price (the shops' layout: also E, the shelf price)
const INSTANT_WATCHED_COLUMNS = [1, 3, 4, ...(LAYOUT.shelfPrice ? [LAYOUT.shelfPrice] : [])];
const INSTANT_NAME_COLUMN = 1;
const INSTANT_COUNT_COLUMN = LAYOUT.count; // AK
const INSTANT_OPENING_COLUMN = 2; // B
const INSTANT_FIRST_ROW = 4;
const INSTANT_PENDING_KEY = 'instantPending'; // JSON: which business changed, from which period
const INSTANT_ALL_CHAINS = '*';
const INSTANT_PRODUCTS_FLAG = '#products';
const INSTANT_LOCK_WAIT_MS = 1000;
const INSTANT_MAX_ROUNDS = 5;
const INSTANT_TOAST_TITLE = 'Жилийн sheet';
const INSTANT_TOAST_SECONDS = 8;
// What each kind of "On change" event starts. Typed values ("EDIT") are handled by onSheetEdit.
const INSTANT_CHANGE_TASKS = {
  REMOVE_ROW: 'sync',
  INSERT_ROW: 'trim',
  FORMAT: 'dividers',
  OTHER: 'tabs',
  INSERT_GRID: 'tabs',
};

// The edit and change triggers of `spreadsheet` (the workbook of this run unless given: a year's
// file holds no script, so this project listens to it; year-books.gs).
function installInstantSync(spreadsheet = currentBook()) {
  const installed = triggersOf(spreadsheet).map((trigger) => trigger.getHandlerFunction());
  if (!installed.includes(INSTANT_HANDLER)) {
    ScriptApp.newTrigger(INSTANT_HANDLER).forSpreadsheet(spreadsheet).onEdit().create();
  }
  if (!installed.includes(INSTANT_CHANGE_HANDLER)) {
    ScriptApp.newTrigger(INSTANT_CHANGE_HANDLER).forSpreadsheet(spreadsheet).onChange().create();
  }
}

function removeInstantSync() {
  ScriptApp.getProjectTriggers()
    .filter((trigger) => [INSTANT_HANDLER, INSTANT_CHANGE_HANDLER].includes(trigger.getHandlerFunction()))
    .forEach((trigger) => ScriptApp.deleteTrigger(trigger));
}

function onSheetEdit(event) {
  if (!event || !event.range) {
    return;
  }
  targetBook(event); // year-books.gs: the edited workbook, whatever its year
  if (!isPairedLayoutReady()) {
    return; // while the columns are being moved (layout-migration.gs) nothing is done
  }
  if (isProductListEdit(event)) {
    runProductListEdit(event); // a name corrected in the product list follows in both shops
    return;
  }
  const edit = describeEdit(event.range);
  if (isHeaderEdit(edit)) {
    restoreHeader(event.range.getSheet()); // a title or label was changed or emptied
  }
  if (isViewEdit(edit)) {
    applyView(event.range.getSheet()); // A1: show or hide columns
  }
  if (isCalcEdit(edit)) {
    repairCalcColumns(event.range.getSheet()); // something was typed over a formula column
  }
  if (isCountEdit(edit)) {
    runInstantCarry(edit.sheetName, 0); // this sheet's count -> the next sheet's B
  }
  if (isOpeningEdit(edit)) {
    runInstantCarry(edit.sheetName, -1); // B typed by hand: the count before it wins
  }
  if ((CARRY_FROM_BALANCE || UNCOUNTED_SHEETS.includes(edit.sheetName)) && isMovementEdit(edit)) {
    runInstantCarry(edit.sheetName, 0); // income or expense changed the final balance
  }
  if (typeof isTransferEdit === 'function' && isTransferEdit(edit)) {
    runEditedTransferLink(event.range); // transfer-link.gs (main workbook): the other shop's side at once
  }
  if (!isWatchedEdit(edit)) {
    return;
  }
  if (typeof linkEdit === 'function') {
    linkEdit(event.range); // price-link.gs (main workbook only): the price in the other shop and Тосгон, a new product in Тосгон
  }
  if (typeof tosgonLinkEdit === 'function') {
    tosgonLinkEdit(event.range); // tosgon-link.gs (Тосгон only): the price in both shops, a new product in Соронзон
  }
  if (runQuickPriceCarry(event)) {
    return; // one cost or price cell: done the quick way
  }
  // Only a changed name (column A) can add, rename or remove a product. A changed
  // cost or price needs the price sync alone, which is much less to read.
  const isNameEdit = edit.firstColumn <= INSTANT_NAME_COLUMN;
  const sheet = isNameEdit ? event.range.getSheet() : null;
  if (isNameEdit) {
    runQuickRename(event, edit); // before the sync, which would remove and add the row
  }
  runPendingSync(scopeOf(edit.sheetName), isNameEdit, () => {
    if (sheet !== null) {
      reconcileSheet(sheet); // a coloured cell that just got a category name
      drawEditedRowBorders(sheet, event.range.getRow(), event.range.getLastRow());
    }
  });
}

function onInstantChange(event) {
  const task = event ? INSTANT_CHANGE_TASKS[event.changeType] : undefined;
  if (task === undefined) {
    return;
  }
  targetBook(event); // year-books.gs: the changed workbook, whatever its year
  if (!isPairedLayoutReady()) {
    return;
  }
  if (task === 'sync') {
    const active = event.source ? event.source.getActiveSheet() : null;
    runPendingSync(active ? scopeOf(active.getName()) : null, true, () => {});
    return;
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(INSTANT_LOCK_WAIT_MS)) {
    return; // something else is working; the hourly job covers this
  }
  try {
    const spreadsheet = currentBook();
    const sheet = event.source ? event.source.getActiveSheet() : null;
    runInstantTask(task, spreadsheet, sheet);
  } finally {
    lock.releaseLock();
  }
}

function runInstantTask(task, spreadsheet, sheet) {
  if (task === 'tabs') {
    Logger.log(runArrangeTabs(spreadsheet).join('\n'));
    return;
  }
  if (sheet === null || sheet.getName().startsWith('_') || isStockSheet(sheet)) {
    return;
  }
  // The owner's colouring is remembered before the colours are put in order, so it
  // stays; anyone else's is removed. An inserted row may have inherited a divider's
  // colour, so this is done after an insert too.
  const isOwner = task === 'dividers' && isOwnerActing();
  if (isOwner && sheetSlot(sheet.getName()) !== null) {
    approveSheetColors(sheet);
  }
  reconcileSheet(sheet);
  if (task === 'dividers') {
    cleanBordersAfterFormat(sheet, isOwner); // borders are formatting too, and cannot be read
  }
  if (task === 'trim') {
    trimSheet(sheet);
  }
}

// One typed cost or price cell is carried without reading the whole business.
// Returns false when the edit is something else and the full sync has to run:
// also when Google does not say what the cell held before (an empty cell, a paste).
function runQuickPriceCarry(event) {
  if (event.oldValue === undefined) {
    return false;
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(INSTANT_LOCK_WAIT_MS)) {
    return false;
  }
  try {
    const spreadsheet = currentBook();
    const prices = carryEditedPrice(spreadsheet, event.range, event.oldValue);
    if (prices === null) {
      return false;
    }
    formatPriceReport(prices);
    showPriceToast(spreadsheet, prices);
    return true;
  } finally {
    lock.releaseLock();
  }
}

// One name cell was typed over: the product keeps its row and its numbers on the
// later sheets and only gets the new name there. Does nothing when Google does not
// say what the cell held before (a new row, a paste): then the product sync decides.
function runQuickRename(event, edit) {
  const isOneCell = edit.firstRow === edit.lastRow && edit.firstColumn === edit.lastColumn;
  if (!isOneCell || edit.firstRow < INSTANT_FIRST_ROW) {
    return;
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(INSTANT_LOCK_WAIT_MS)) {
    return; // something else is working; the product sync handles the edit
  }
  try {
    const spreadsheet = currentBook();
    const products = readModel(event.range.getSheet()).products;
    const edited = products.find((product) => product.row === edit.firstRow);
    if (edited === undefined) {
      return; // an emptied cell or a category row
    }
    const messages = [];
    const taken = products.filter((product) => product.row !== edited.row).map((product) => product.key);
    const named = { ...edited, label: freeName(edited.label, taken) };
    if (named.label !== edited.label) {
      // No two rows of one name: the typed one gets a number, "Haier 12 (2)".
      if (!DRY_RUN) {
        event.range.setValue(named.label);
      }
      messages.push(`"${edited.label}" нэртэй бараа аль хэдийн байгаа тул "${named.label}" гэж нэрлэлээ.`);
    }
    const renamedProduct = { ...named, key: normalizeKey(named.label) };
    const oldKey = event.oldValue === undefined ? '' : normalizeKey(event.oldValue);
    if (oldKey !== '' && oldKey !== renamedProduct.key) {
      const renamed = renameOnLaterSheets(spreadsheet, sheetSlot(edit.sheetName), oldKey, renamedProduct);
      if (renamed > 0) {
        messages.push(
          DRY_RUN
            ? 'ТУРШИЛТЫН горим (DRY_RUN = true): нэрийн өөрчлөлт дараагийн sheet-д ДАМЖААГҮЙ'
            : `Барааны нэр дараагийн ${renamed} sheet дээр солигдлоо, тоонууд хэвээр`
        );
      }
    }
    showToast(spreadsheet, messages.join(' '));
  } finally {
    lock.releaseLock();
  }
}

// Returns how many later sheets got (or, in a dry run, would get) the new name.
// A sheet that already has the new name is left alone: no second row of that name.
function renameOnLaterSheets(spreadsheet, slot, oldKey, edited) {
  const scope = { chainKey: slot.chainKey, fromPosition: slot.position + 1 };
  return findChains(spreadsheet, scope)
    .flatMap(({ sheets }) => sheets)
    .filter((sheet) => {
      const products = readModel(sheet).products;
      const old = products.find((product) => product.key === oldKey);
      if (old === undefined || products.some((product) => product.key === edited.key)) {
        return false;
      }
      if (!DRY_RUN) {
        sheet.getRange(old.row, INSTANT_NAME_COLUMN).setValue(edited.label);
      }
      return true;
    }).length;
}

function showToast(spreadsheet, message) {
  if (message === '') {
    return;
  }
  try {
    spreadsheet.toast(message, INSTANT_TOAST_TITLE, INSTANT_TOAST_SECONDS);
  } catch (error) {
    Logger.log(message); // no window to show it in
  }
}

// Product and price sync, at most one at a time; changes that arrive meanwhile are
// picked up by the one that is running. Only the business that was changed is read,
// from the changed sheet onward (scope); null means every sheet. needsProducts:
// the list of products may have changed, so the product sync runs before the prices.
function runPendingSync(scope, needsProducts, beforeSync) {
  const properties = bookScriptProperties(); // year-books.gs: the pending work of this workbook
  properties.setProperty(
    INSTANT_PENDING_KEY,
    JSON.stringify(mergeScope(readPendingScopes(properties), scope, needsProducts))
  );
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(INSTANT_LOCK_WAIT_MS)) {
    return; // a run is working; it sees the note and goes round again
  }
  try {
    const spreadsheet = currentBook();
    const synced = [];
    beforeSync();
    for (let round = 0; round < INSTANT_MAX_ROUNDS; round += 1) {
      const pending = readPendingScopes(properties);
      if (listScopes(pending).length === 0) {
        break;
      }
      synced.push(...listScopes(pending));
      properties.deleteProperty(INSTANT_PENDING_KEY);
      listScopes(pending).forEach((pendingScope) => {
        if (pending[INSTANT_PRODUCTS_FLAG] === true) {
          formatSyncReport(runSync(spreadsheet, undefined, pendingScope));
        }
        const prices = runPriceSync(spreadsheet, pendingScope);
        formatPriceReport(prices);
        showPriceToast(spreadsheet, prices);
      });
    }
    Logger.log(refreshCountSheets(spreadsheet, synced).join('\n')); // the count sheets follow the product list
  } finally {
    lock.releaseLock();
  }
}

// A short note in the corner of the sheet, so a carried price is seen at once -
// and so is a price that was NOT carried because the script is in test mode.
function showPriceToast(spreadsheet, prices) {
  if (prices.lines.length === 0) {
    return;
  }
  const message = DRY_RUN
    ? 'ТУРШИЛТЫН горим (DRY_RUN = true): үнийн өөрчлөлт дараагийн sheet-д ДАМЖААГҮЙ'
    : `Үнэ, өртөг дараагийн sheet-үүдийн ${prices.cells} нүдэнд шинэчлэгдлээ`;
  try {
    spreadsheet.toast(message, INSTANT_TOAST_TITLE, INSTANT_TOAST_SECONDS);
  } catch (error) {
    Logger.log(message); // no window to show it in
  }
}

function readPendingScopes(properties) {
  const saved = properties.getProperty(INSTANT_PENDING_KEY);
  try {
    const parsed = saved ? JSON.parse(saved) : {};
    return parsed !== null && typeof parsed === 'object' ? parsed : {};
  } catch (error) {
    return { [INSTANT_ALL_CHAINS]: 0, [INSTANT_PRODUCTS_FLAG]: true }; // a note in an older form: do everything once
  }
}

// ---------- pure helpers (no Spreadsheet calls) ----------

function scopeOf(sheetName) {
  const slot = sheetSlot(sheetName);
  return slot === null ? null : { chainKey: slot.chainKey, fromPosition: slot.position };
}

// pending: { chainKey: earliest changed period }, or { '*': 0 } for "every sheet",
// plus { '#products': true } when the product sync is needed too.
function mergeScope(pending, scope, needsProducts) {
  const key = scope ? scope.chainKey : INSTANT_ALL_CHAINS;
  const position = scope ? scope.fromPosition : 0;
  const earliest = pending[key] === undefined ? position : Math.min(pending[key], position);
  const flag = needsProducts || pending[INSTANT_PRODUCTS_FLAG] === true;
  return { ...pending, [key]: earliest, ...(flag ? { [INSTANT_PRODUCTS_FLAG]: true } : {}) };
}

// The scopes to run; undefined stands for "every sheet" and replaces all others.
function listScopes(pending) {
  if (pending[INSTANT_ALL_CHAINS] !== undefined) {
    return [undefined];
  }
  return Object.entries(pending)
    .filter(([key]) => key !== INSTANT_PRODUCTS_FLAG)
    .map(([chainKey, fromPosition]) => ({ chainKey, fromPosition }));
}

// edit: { sheetName, firstRow, lastRow, firstColumn, lastColumn }
function isWatchedEdit(edit) {
  return (
    sheetSlot(edit.sheetName) !== null &&
    edit.lastRow >= INSTANT_FIRST_ROW &&
    INSTANT_WATCHED_COLUMNS.some((column) => column >= edit.firstColumn && column <= edit.lastColumn)
  );
}

// A count (AK) was typed or corrected: the next sheet's opening balance follows.
// An opening balance (B) was typed: it is put back to the count of the sheet before.
// Two sheets are read: the edited one and the one after it (offset 0), or the one
// before it and the edited one (offset -1).
function runInstantCarry(sheetName, offset) {
  const slot = sheetSlot(sheetName);
  const fromPosition = slot.position + offset;
  if (fromPosition < 0) {
    return; // the first sheet of the year has no sheet before it
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(INSTANT_LOCK_WAIT_MS)) {
    return; // something else is working; the hourly job covers this
  }
  try {
    runCarry(currentBook(), {
      chainKey: slot.chainKey,
      fromPosition,
      // a balance reaches every later sheet
      toPosition: CARRY_FROM_BALANCE || UNCOUNTED_SHEETS.includes(sheetName) ? undefined : fromPosition + 1,
    });
  } finally {
    lock.releaseLock();
  }
}

function isCountEdit(edit) {
  return (
    sheetSlot(edit.sheetName) !== null &&
    edit.lastRow >= INSTANT_FIRST_ROW &&
    edit.firstColumn <= INSTANT_COUNT_COLUMN &&
    edit.lastColumn >= INSTANT_COUNT_COLUMN
  );
}

// An edit that reaches into the income or expense columns of a product row.
function isMovementEdit(edit) {
  return (
    sheetSlot(edit.sheetName) !== null &&
    edit.lastRow >= INSTANT_FIRST_ROW &&
    edit.firstColumn <= LAYOUT.expenseLast &&
    edit.lastColumn >= LAYOUT.incomeFirst
  );
}

function isOpeningEdit(edit) {
  return (
    sheetSlot(edit.sheetName) !== null &&
    edit.lastRow >= INSTANT_FIRST_ROW &&
    edit.firstColumn <= INSTANT_OPENING_COLUMN &&
    edit.lastColumn >= INSTANT_OPENING_COLUMN
  );
}

// An edit that touches the header rows 1-3 of a year sheet.
function isHeaderEdit(edit) {
  return sheetSlot(edit.sheetName) !== null && edit.firstRow <= HEADER_ROWS;
}

// An edit that reaches into the formula columns AL:AN below their header.
function isCalcEdit(edit) {
  return (
    sheetSlot(edit.sheetName) !== null &&
    edit.lastRow >= INSTANT_FIRST_ROW &&
    edit.lastColumn >= CALC_FIRST_COL &&
    edit.firstColumn < CALC_FIRST_COL + CALC_COLUMN_COUNT
  );
}

// ---------- spreadsheet access ----------

function describeEdit(range) {
  return {
    sheetName: range.getSheet().getName(),
    firstRow: range.getRow(),
    lastRow: range.getLastRow(),
    firstColumn: range.getColumn(),
    lastColumn: range.getLastColumn(),
  };
}
