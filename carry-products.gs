/**
 * The products of a sheet carried at once into the later sheets of its chain that lack them (the
 * owner, 2026-10-28: Tos/2026/11 and Tos/2026/12 held about 145 of October's 290 products).
 *
 * The hourly sync (year-sync.gs runSync) carries only what it saw added or removed since its last
 * look ("_sync"): a product a later sheet never had stays missing there for good, and its count does
 * not reach the next month's opening balance. Here every product of the current sheet that a later
 * sheet lacks is put into its category there, as the sync puts in a new one (insertProduct: name,
 * cost, price, shelf price, the row's formats; none of the month's numbers). Nothing is taken away. A
 * name the list "Бараа" lacks is not written, only named (the owner's rule, 2026-10-25). The sync's
 * saved lists learn the new rows, so it does not carry them a second time.
 *
 * Putting rows in one by one is slow: a run stops at the deadline and says how many are left, and
 * the next run goes on with them.
 *   carryProductsForward(): from the sheet of today on, in every chain (DRY_RUN only lists).
 * Needs config.gs, year-sync.gs, product-list.gs (listNameKeys), official-names.gs (withoutFilter).
 */

const CARRY_PRODUCTS_TIME_BUDGET_MS = 4.5 * 60 * 1000;
const CARRY_PRODUCTS_LIST_LIMIT = 30;

function carryProductsForward() {
  withScriptLock(() => {
    const deadline = Date.now() + CARRY_PRODUCTS_TIME_BUDGET_MS;
    notify(runCarryProducts(currentBook(), positionOfDate(new Date()), deadline).join('\n'));
  });
}

// The log lines of one run. position: the sheet the products come from (config.gs periodPosition).
function runCarryProducts(spreadsheet, position, deadline) {
  const listed = listNameKeys(spreadsheet);
  const saved = loadSnapshots(spreadsheet);
  const results = findChains(spreadsheet).map((chain) => carryChainProducts(chain, position, listed, deadline));
  const snapshots = updateSnapshots(saved, results.flatMap((result) => result.perTarget));
  if (!DRY_RUN && JSON.stringify(snapshots) !== JSON.stringify(saved)) {
    saveSnapshots(spreadsheet, snapshots);
  }
  return [modeLabel(), ...results.flatMap((result) => result.lines)];
}

// One chain: { lines, perTarget } (perTarget as propagate gives it, for updateSnapshots).
function carryChainProducts({ label, sheets }, position, listed, deadline) {
  const at = sheets.findIndex((sheet) => sheetSlot(sheet.getName()).position === position);
  if (at === -1) {
    return { lines: [`${label}: энэ үеийн sheet алга`], perTarget: [] };
  }
  const source = sheets[at];
  const products = readModel(source).products;
  const isListed = (product) => listed === null || listed.has(product.key);
  const items = products.filter(isListed).map((product) => describeProduct(source, product));
  const offList = products.filter((product) => !isListed(product)).map((product) => `"${product.label}"`);
  const carried = sheets.slice(at + 1).map((target) => ({ target, ...carryProductsInto(target, items, deadline) }));
  const lines = [
    `${label}: ${source.getName()}-ийн ${items.length} бараа дараагийн sheet-үүдэд`,
    ...carried.filter(({ missing }) => missing.length > 0).map(carryProductsLine),
    ...(offList.length > 0 ? [`  ${offList.join(', ')}: "${PRODUCT_LIST_TAB}" жагсаалтад алга тул нэмсэнгүй`] : []),
  ];
  const perTarget = carried.map(({ target, added }) => ({ name: target.getName(), added, alreadyThere: [], removed: [] }));
  return { lines, perTarget };
}

// The items a target lacks, put in until the deadline: { missing: [item], added: [key] }.
function carryProductsInto(target, items, deadline) {
  const present = new Set(readModel(target).products.map((product) => product.key));
  const missing = items.filter((item) => !present.has(item.key));
  if (DRY_RUN || missing.length === 0) {
    return { missing, added: [] };
  }
  const added = [];
  withoutFilter(target, () => {
    for (const item of missing) {
      if (Date.now() > deadline) {
        return;
      }
      insertProduct(target, readModel(target), item);
      added.push(item.key);
    }
  });
  return { missing, added };
}

function carryProductsLine({ target, missing, added }) {
  const names = missing.slice(0, CARRY_PRODUCTS_LIST_LIMIT).map((item) => item.label).join(', ');
  const more = missing.length > CARRY_PRODUCTS_LIST_LIMIT ? ` ... бас ${missing.length - CARRY_PRODUCTS_LIST_LIMIT}` : '';
  if (DRY_RUN) {
    return `  ${target.getName()}: ${missing.length} бараа нэмнэ (туршилт): ${names}${more}`;
  }
  const left = missing.length - added.length;
  return `  ${target.getName()}: ${added.length} бараа нэмлээ` +
    (left > 0 ? `; цаг дууслаа, ${left} нь үлдлээ: дахин Run хийнэ үү` : `: ${names}${more}`);
}
