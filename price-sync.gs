/**
 * Carries a changed cost (C) or price (D) forward through the year.
 *
 * When the cost or price of a product is changed on one sheet, the same product
 * gets the new value on every LATER sheet of the same business, as long as that
 * later sheet still holds the OLD value. A later sheet whose value is already
 * different (the price was changed there on purpose, or is being changed there
 * right now) keeps its own value, and the change stops at that sheet.
 *   Example: "Haier 12" costs 550,000 on the sheets of October to December. The
 *   price is set to 570,000 on October 1-15 -> October 16-31, November and
 *   December become 570,000 too. Earlier sheets are never touched.
 *
 * How it knows what changed: the cost and price of every product are saved in a
 * hidden sheet "_prices" after each run; the next run compares against it. The
 * very first run only saves them (baseline) and changes nothing.
 *
 * One cell typed by hand is carried the quick way (carryEditedPrice): the edit
 * itself says what the cell held before, so the saved prices are not needed.
 *
 * Products are matched by name, like in year-sync.gs.
 *   syncPricesForward - run by hand (menu "Жилийн sheet"); the background job runs
 *                       it every hour. DRY_RUN (config.gs) makes it report only.
 * Needs config.gs, year-sync.gs and dividers.gs.
 */

const PRICE_STATE_SHEET = '_prices';
const PRICE_FIELDS = [
  { label: 'өртөг', column: 3 },
  { label: 'үнэ', column: 4 },
  // the shops' layout: the hand-typed shelf price is carried the same way
  ...(LAYOUT.shelfPrice ? [{ label: 'яг харагдах үнэ', column: LAYOUT.shelfPrice }] : []),
];

function syncPricesForward() {
  withScriptLock(() => notify(formatPriceReport(runPriceSync(currentBook()))));
}

// ---------- pure helpers (no Spreadsheet calls) ----------

// saved / current: { productKey: [cost, price(, shelf price)] }. Only products present in both.
// A value the saved state does not have yet (saved before the shelf price existed) is no change:
// the next run saves it as the baseline.
function findPriceChanges(saved, current) {
  return Object.keys(current)
    .filter((key) => Array.isArray(saved[key]))
    .flatMap((key) =>
      PRICE_FIELDS.map((field, index) => ({
        key,
        index,
        from: saved[key][index],
        to: current[key][index],
      }))
    )
    .filter((change) => change.from !== undefined && change.from !== change.to);
}

function changeId(change) {
  return `${change.index}|${change.key}`;
}

function withValue(values, key, index, value) {
  const pair = values[key].map((old, position) => (position === index ? value : old));
  return { ...values, [key]: pair };
}

// ---------- spreadsheet access ----------

function readPrices(sheet) {
  return readModel(sheet).products.reduce(
    (all, product) => ({
      rows: { ...all.rows, [product.key]: product.row },
      values: { ...all.values, [product.key]: [product.cost, product.price, ...(LAYOUT.shelfPrice ? [product.shelfPrice] : [])] },
    }),
    { rows: {}, values: {} }
  );
}

function runPriceSync(spreadsheet, scope) {
  if (typeof isTidyInProgress === 'function' && isTidyInProgress()) {
    return { lines: [TIDY_BUSY_TEXT], cells: 0 }; // official-names.gs: names are being changed sheet by sheet
  }
  const saved = loadSnapshots(spreadsheet, PRICE_STATE_SHEET);
  const lines = [];
  let cells = 0;
  let next = saved;

  findChains(spreadsheet, scope).forEach(({ sheets }) => {
    const states = sheets.map((sheet) => ({ sheet, name: sheet.getName(), ...readPrices(sheet) }));
    const changes = states.map((state) =>
      saved[state.name] ? findPriceChanges(saved[state.name], state.values) : []
    );
    const ownChanges = changes.map((list) => new Set(list.map(changeId)));
    let values = states.map((state) => state.values);

    changes.forEach((list, index) => {
      list.forEach((change) => {
        const carried = carryPrice(states, values, ownChanges, index, change);
        values = carried.values;
        cells += carried.count;
        lines.push(
          `${states[index].name}: "${change.key}" ${PRICE_FIELDS[change.index].label} ` +
            `${change.from} -> ${change.to}, дараагийн ${carried.count} sheet-д`
        );
      });
    });
    states.forEach((state, index) => {
      next = { ...next, [state.name]: values[index] };
    });
  });

  // Nothing is saved in a dry run, so the change is still seen by the real run.
  if (!DRY_RUN && JSON.stringify(next) !== JSON.stringify(saved)) {
    saveSnapshots(spreadsheet, next, PRICE_STATE_SHEET);
  }
  return { lines, cells };
}

// Writes one change into the later sheets until one of them has its own value.
function carryPrice(states, values, ownChanges, sourceIndex, change) {
  let nextValues = values;
  let count = 0;
  for (let index = sourceIndex + 1; index < states.length; index += 1) {
    const pair = nextValues[index][change.key];
    if (!pair) {
      continue; // the product is not on this sheet
    }
    if (ownChanges[index].has(changeId(change)) || pair[change.index] !== change.from) {
      break;
    }
    if (!DRY_RUN) {
      states[index].sheet
        .getRange(states[index].rows[change.key], PRICE_FIELDS[change.index].column)
        .setValue(change.to);
    }
    nextValues = nextValues.map((sheetValues, position) =>
      position === index ? withValue(sheetValues, change.key, change.index, change.to) : sheetValues
    );
    count += 1;
  }
  return { values: nextValues, count };
}

// ---------- the quick way: one cost or price cell typed by hand ----------

function isSamePrice(a, b) {
  const isEmpty = (value) => value === '' || value === null || value === undefined;
  if (isEmpty(a) || isEmpty(b)) {
    return isEmpty(a) && isEmpty(b);
  }
  return Number(a) === Number(b) || String(a) === String(b);
}

// The row of a product on a sheet, by name; null when it is not there. names: column A
// from the first product row down. The last one wins, like in readPrices.
function findProductRow(names, key, firstRow) {
  const index = names.map((cells) => normalizeKey(cells[0])).lastIndexOf(key);
  return index === -1 ? null : firstRow + index;
}

// "On edit" tells what the cell held before, so nothing has to be compared with
// the saved prices: only column A and one cell of each later sheet are read. The
// rule is the same as in runPriceSync. Returns the report, or null when the edit
// is not one cost/price cell of a product on a year sheet. The saved prices are
// left as they were; the next full run sees the new values and saves them.
function carryEditedPrice(spreadsheet, range, oldValue) {
  const sheet = range.getSheet();
  const slot = sheetSlot(sheet.getName());
  const field = PRICE_FIELDS.find((candidate) => candidate.column === range.getColumn());
  const isOneCell = range.getNumRows() === 1 && range.getNumColumns() === 1;
  if (slot === null || !field || !isOneCell || range.getRow() <= SYNC_HEADER_ROWS) {
    return null;
  }
  const label = String(sheet.getRange(range.getRow(), SYNC_NAME_COL).getValue()).trim();
  if (label === '') {
    return null;
  }
  const key = normalizeKey(label);
  const newValue = range.getValue();
  const laterSheets = findChains(spreadsheet, { chainKey: slot.chainKey, fromPosition: slot.position + 1 })
    .flatMap((chain) => chain.sheets);
  const firstRow = SYNC_HEADER_ROWS + 1;
  let cells = 0;
  for (const later of laterSheets) {
    const height = later.getLastRow() - SYNC_HEADER_ROWS;
    const names = height > 0 ? later.getRange(firstRow, SYNC_NAME_COL, height, 1).getValues() : [];
    const row = findProductRow(names, key, firstRow);
    if (row === null) {
      continue; // the product is not on this sheet
    }
    const cell = later.getRange(row, field.column);
    if (!isSamePrice(cell.getValue(), oldValue)) {
      break; // this sheet has its own value: the change stops here
    }
    if (!DRY_RUN) {
      cell.setValue(newValue);
    }
    cells += 1;
  }
  const line = `${sheet.getName()}: "${key}" ${field.label} ${oldValue} -> ${newValue}, дараагийн ${cells} sheet-д`;
  return { lines: cells > 0 ? [line] : [], cells };
}

function formatPriceReport({ lines, cells }) {
  const mode = DRY_RUN ? '[ТУРШИЛТ - юу ч өөрчлөгдөөгүй] ' : '';
  const summary = `${mode}Үнэ, өртгийн өөрчлөлт: ${lines.length}, дараагийн sheet-үүдэд шинэчилсэн нүд: ${cells}`;
  Logger.log([summary, ...lines].join('\n'));
  return summary;
}
