/**
 * ONE product list for both shops, so one product has one name in Соронзон and in My home.
 *
 * Until now each shop typed its own names, so the same product is "Haeir 103l" in
 * one shop and "Haier 103l" in the other, "8037 ус" here and "Ус буцалгагч 8037"
 * there. The scripts know a product by its name only, so for them those are
 * different products.
 *
 * The owner first marks the pairs in the tab DECISION_TAB ("_Нэр тулгалт": My home
 * name, Соронзон name, suggested name, тийм / үгүй, corrected name). Then
 * setupProductList does, in one go:
 *   1. the confirmed pairs get the chosen name on every sheet of the year, in both
 *      shops (runUnifyNames of name-fixes.gs, which also saves _sync / _prices anew);
 *   2. the tab PRODUCT_LIST_TAB ("Бараа") is made from the current half-month of
 *      both shops: the category names as heading rows, each product once, and which
 *      shop has it. An existing list is never overwritten;
 *   3. column A of every year sheet becomes a dropdown of that list (installNameRule
 *      in name-fixes.gs), instead of the "no name twice" rule.
 * It stops before changing anything while a row of the table has no answer.
 *
 * Afterwards:
 *   - a new product is added as one row of "Бараа", then picked from the dropdown
 *     in the shop's sheet; the dropdown reads the whole column, so it is there at once;
 *   - a name corrected in "Бараа" is corrected on every sheet of both shops
 *     (runProductListEdit, called by the On edit trigger of instant-sync.gs);
 *   - PRODUCT_LIST_STRICT = true (the owner's hard rule, 2026-10-20): a name that is not
 *     on the list is refused; false would only give it a red corner. A new category
 *     is first added to the list as a heading row. A repeated name is still numbered
 *     "X (2)" by the script (name-fixes.gs).
 *
 * Тосгон (its own workbook) uses the same list: a protected tab "Бараа" there is a copy
 * made by IMPORTRANGE (installListCopy, tosgon-names.gs), so its dropdown is the same
 * list. A list filled by a formula is a copy: typing in it renames nothing. A name
 * corrected in the real list is written to LIST_RENAME_LOG_TAB, and the workbook with
 * the copy carries it into its own sheets every hour (runListRenamesFromSource).
 * Needs config.gs, year-sync.gs, name-fixes.gs, dividers.gs, instant-sync.gs.
 */

const PRODUCT_LIST_TAB = 'Бараа';
const PRODUCT_LIST_FIRST_ROW = 4; // rows 1-3: title, how to use it, column names
const PRODUCT_LIST_STRICT = true;
// The workbook that holds the real list (бараа орлого зарлага 2026); other workbooks hold a copy.
const PRODUCT_LIST_SOURCE_ID = 'YOUR_PRODUCT_LIST_SPREADSHEET_ID';
const LIST_RENAME_LOG_TAB = '_Нэр солилт'; // date, old key, new name: one row per rename made in the list
const LIST_RENAMES_DONE_KEY = 'listRenamesDone'; // document property: log rows already carried here
const PRODUCT_LIST_LOCK_WAIT_MS = 30 * 1000;
const PRODUCT_LIST_NAME_WIDTH = 260;
const PRODUCT_LIST_SHOP_WIDTH = 110;
const PRODUCT_LIST_HELP =
  '"Бараа" tab-ын жагсаалтаас сонгоно уу. Шинэ бараа бол эхлээд "Бараа" tab-д нэг мөр нэмнэ.';
const PRODUCT_LIST_HEADER = [
  ['Барааны жагсаалт: хоёр дэлгүүрт нэг бараа нэг нэртэй', ''],
  [
    'Шинэ бараа ирвэл ЭНД нэг мөр нэмээд дэлгүүрийн sheet-д жагсаалтаас сонгоно. ' +
      'Нэрийг энд засвал хоёр дэлгүүрийн бүх sheet дээр солигдоно.',
    '',
  ],
  ['Барааны нэр', 'Аль дэлгүүрт'],
];
const MY_HOME_KEY = 'my home';
const SORONZON_KEY = 'соронзон';
const TOSGON_KEY = 'тосгон';
const SHOP_LABELS = { [MY_HOME_KEY]: 'My home', [SORONZON_KEY]: 'Соронзон', [TOSGON_KEY]: 'Тосгон' };
const BOTH_SHOPS_LABEL = 'Хоёулаа';

const DECISION_TAB = '_Нэр тулгалт';
const DECISION_FIRST_ROW = 3;
const DECISION_WIDTH = 11; // A..K
const DECISION_COLUMNS = { my: 3, sor: 5, suggested: 8, answer: 9, corrected: 10 }; // C E H I J
const DECISION_YES = 'тийм';
const DECISION_NO = 'үгүй';

const LIST_WATCH_KEY = 'listWatch'; // script property: fingerprint of the report already mailed
const LIST_WATCH_MAIL_SUBJECT = 'Бараа жагсаалтад байхгүй нэр';
const LIST_WATCH_HINT = 'Эдгээрийг "Бараа" tab-д нэмэх, эсвэл дэлгүүрийн sheet-д жагсаалтаас сонгож засна уу.';
const NUMBERED_NAME = /\(\d+\)$/; // "Haier 12 (2)": a repeated name the script numbered
// C..E of the list: cost, price and shelf price of the shop chosen in A1, all filled by ONE
// formula in C4 (made by hand live, 2026-10-27). Nothing is typed there.
const LIST_VIEW_FIRST_COLUMN = 3;
const LIST_VIEW_LAST_COLUMN = 5;
const LIST_VIEW_MESSAGE =
  'Үнийг энд бичихгүй: дэлгүүрийн sheet-ийн C (өртөг үнэ), D (зарах үнэ), E (яг харагдах үнэ)-д бичнэ үү. Энд бичсэнийг арилгалаа.';

function setupProductList() {
  withScriptLock(() => {
    const lines = runSetupProductList(currentBook(), new Date());
    Logger.log(lines.join('\n'));
  });
}

// ---------- pure helpers (no Spreadsheet calls) ----------

function cleanLabel(value) {
  return String(value === undefined || value === null ? '' : value).trim().replace(/\s+/g, ' ');
}

// rows: the table from DECISION_FIRST_ROW down (DECISION_WIDTH cells each).
// Returns { fixes: { shop key: [[chosen name, [key of that shop's spelling]]] },
//           undecided: [row numbers], problems: [messages] }.
function planNameDecisions(rows) {
  const cellOf = (row, column) => cleanLabel(row[column - 1]);
  const filled = rows
    .map((row, at) => ({ row, line: DECISION_FIRST_ROW + at }))
    .filter(({ row }) => cellOf(row, DECISION_COLUMNS.my) !== '' || cellOf(row, DECISION_COLUMNS.sor) !== '');
  const answerOf = (row) => cellOf(row, DECISION_COLUMNS.answer).toLowerCase();
  const undecided = filled.filter(({ row }) => answerOf(row) === '').map(({ line }) => line);
  const odd = filled
    .filter(({ row }) => !['', DECISION_YES, DECISION_NO].includes(answerOf(row)))
    .map(({ row, line }) => `${line}-р мөр: "${cellOf(row, DECISION_COLUMNS.answer)}" - "тийм" эсвэл "үгүй" гэж бичнэ`);
  const pairs = filled
    .filter(({ row }) => answerOf(row) === DECISION_YES)
    .map(({ row, line }) => ({
      line,
      my: cellOf(row, DECISION_COLUMNS.my),
      sor: cellOf(row, DECISION_COLUMNS.sor),
      name: cellOf(row, DECISION_COLUMNS.corrected) || cellOf(row, DECISION_COLUMNS.suggested),
    }));
  const nameless = pairs.filter((pair) => pair.name === '').map((pair) => `${pair.line}-р мөр: нэр хоосон`);
  const named = pairs.filter((pair) => pair.name !== '');
  const usesOf = (pair) => named.filter((other) => normalizeKey(other.name) === normalizeKey(pair.name)).length;
  const twice = named
    .filter((pair) => usesOf(pair) > 1)
    .map((pair) => `${pair.line}-р мөр: "${pair.name}" нэрийг өөр хос бас авсан`);
  const fixesFor = (side) =>
    named
      .filter((pair) => pair[side] !== '' && normalizeKey(pair[side]) !== normalizeKey(pair.name))
      .map((pair) => [pair.name, [normalizeKey(pair[side])]]);
  return {
    fixes: { [MY_HOME_KEY]: fixesFor('my'), [SORONZON_KEY]: fixesFor('sor') },
    chosen: named.map((pair) => pair.name),
    undecided,
    problems: [...odd, ...nameless, ...twice],
  };
}

// { shop key: { key of the shop's spelling: chosen name } }
function renameMaps(fixes) {
  return Object.fromEntries(
    Object.entries(fixes).map(([shop, list]) => [shop, Object.fromEntries(list.map(([name, [key]]) => [key, name]))])
  );
}

// The products of one shop's sheet (readModel), each with its chosen name and category.
function listEntries(model, renames, shop) {
  const categoryLabels = Object.fromEntries(model.categories.map((category) => [category.key, cleanLabel(category.label)]));
  return model.products.map((product) => {
    const name = renames[product.key] || cleanLabel(product.label);
    const categoryLabel = product.categoryKey === null ? '' : categoryLabels[product.categoryKey] || '';
    return { key: normalizeKey(name), name, categoryKey: normalizeKey(categoryLabel), categoryLabel, shop };
  });
}

function shopsLabel(shops) {
  return new Set(shops).size > 1 ? BOTH_SHOPS_LABEL : SHOP_LABELS[shops[0]];
}

// Rows of the list ([name, shops]; a category heading has no shops), Соронзон's
// order first, My home's other products after, under their category.
function planProductList(sorModel, myModel, renames) {
  const entries = [
    ...listEntries(sorModel, renames[SORONZON_KEY] || {}, SORONZON_KEY),
    ...listEntries(myModel, renames[MY_HOME_KEY] || {}, MY_HOME_KEY),
  ];
  const products = entries.reduce((all, entry) => {
    const known = all.find((product) => product.key === entry.key);
    return known === undefined
      ? [...all, { ...entry, shops: [entry.shop] }]
      : all.map((product) => (product === known ? { ...product, shops: [...product.shops, entry.shop] } : product));
  }, []);
  const categories = products.reduce(
    (all, product) =>
      all.some((category) => category.key === product.categoryKey)
        ? all
        : [...all, { key: product.categoryKey, label: product.categoryLabel }],
    []
  );
  return categories.flatMap((category) => [
    ...(category.label === '' ? [] : [[category.label, '']]),
    ...products
      .filter((product) => product.categoryKey === category.key)
      .map((product) => [product.name, shopsLabel(product.shops)]),
  ]);
}

// A name typed over in the list. otherNames: the names of the list's other rows.
// Returns { rename: true, from: old key, to: new name } or { rename: false, reason }.
function planListEdit(oldValue, newValue, otherNames) {
  const from = normalizeKey(cleanLabel(oldValue));
  const to = cleanLabel(newValue);
  if (from === '' || to === '' || from === normalizeKey(to) || isCategoryName(oldValue)) {
    return { rename: false, reason: '' }; // a new row, an emptied row, a category heading
  }
  if (otherNames.some((name) => normalizeKey(name) === normalizeKey(to))) {
    return { rename: false, reason: `"${to}" нэр жагсаалтад аль хэдийн байна - дэлгүүрүүдэд солиогүй.` };
  }
  return { rename: true, from, to };
}

// The labels whose key is not in onList, each key once (first spelling).
function labelsOffList(labels, onList) {
  return labels
    .map(cleanLabel)
    .filter((label) => label !== '' && !onList.has(normalizeKey(label)))
    .reduce((kept, label) => (kept.some((other) => normalizeKey(other) === normalizeKey(label)) ? kept : [...kept, label]), []);
}

// shops: [{ shop: chain key, products: [{ label }], categories: [{ label }] }] (readModel of
// the sheets that count); listKeys: normalizeKey of every name of the list.
// One block per shop that has something the list lacks; [] when all is on the list.
function planListWatch(shops, listKeys) {
  const onList = new Set(listKeys);
  return shops.flatMap(({ shop, products, categories }) => {
    const names = labelsOffList(products.map((product) => product.label), onList);
    const parts = [
      ['жагсаалтад байхгүй бараа', names.filter((name) => !NUMBERED_NAME.test(name))],
      ['жагсаалтад байхгүй ангилал', labelsOffList(categories.map((category) => category.label), onList)],
      ['давхардаж дугаарлагдсан нэр', names.filter((name) => NUMBERED_NAME.test(name))],
    ]
      .filter(([, list]) => list.length > 0)
      .map(([title, list]) => `  ${title}: ${list.join(', ')}`);
    return parts.length > 0 ? [`${SHOP_LABELS[shop] || shop}\n${parts.join('\n')}`] : [];
  });
}

// ---------- spreadsheet access ----------

// Hourly: the sheets of both shops from the half-month of `today` on should hold only
// names of the list. What they hold besides goes to the owner by e-mail - once per
// finding, like runNameWatch (year-sync.gs). Nothing to watch while there is no list.
function runListWatch(spreadsheet, today = new Date()) {
  if (productListRange(spreadsheet) === null) {
    return [];
  }
  const listKeys = readListNames(spreadsheet.getSheetByName(PRODUCT_LIST_TAB)).map(({ name }) => normalizeKey(name));
  const fromPosition = positionOfDate(today);
  const shops = findChains(spreadsheet).map(({ label, sheets }) => {
    const models = sheets.filter((sheet) => sheetSlot(sheet.getName()).position >= fromPosition).map(readModel);
    return {
      shop: label,
      products: models.flatMap((model) => model.products),
      categories: models.flatMap((model) => model.categories),
    };
  });
  const lines = planListWatch(shops, listKeys);
  const properties = bookScriptProperties(spreadsheet);
  const fingerprint = lines.length > 0 ? hashText(lines.join('\n')) : '';
  if (fingerprint === (properties.getProperty(LIST_WATCH_KEY) || '')) {
    return [];
  }
  if (DRY_RUN) {
    return [`жагсаалтад байхгүй нэр ${lines.length} дэлгүүрт: email явуулна`];
  }
  if (lines.length === 0) {
    properties.deleteProperty(LIST_WATCH_KEY);
    return ['жагсаалтад байхгүй нэр арилсан'];
  }
  MailApp.sendEmail(
    Session.getEffectiveUser().getEmail(),
    LIST_WATCH_MAIL_SUBJECT,
    `"${spreadsheet.getName()}"\n${spreadsheet.getUrl()}\n\n${lines.join('\n')}\n\n${LIST_WATCH_HINT}`
  );
  properties.setProperty(LIST_WATCH_KEY, fingerprint);
  return [`жагсаалтад байхгүй нэр ${lines.length} дэлгүүрт: email явууллаа`];
}

// The names of the list (column A from PRODUCT_LIST_FIRST_ROW down), for the dropdown.
// null while there is no list.
function productListRange(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(PRODUCT_LIST_TAB);
  if (sheet === null || sheet.getLastRow() < PRODUCT_LIST_FIRST_ROW) {
    return null;
  }
  return sheet.getRange(PRODUCT_LIST_FIRST_ROW, 1, sheet.getMaxRows() - PRODUCT_LIST_FIRST_ROW + 1, 1);
}

function readDecisionRows(sheet) {
  const height = sheet.getLastRow() - DECISION_FIRST_ROW + 1;
  return height > 0 ? sheet.getRange(DECISION_FIRST_ROW, 1, height, DECISION_WIDTH).getValues() : [];
}

function readListNames(sheet) {
  const height = sheet.getLastRow() - PRODUCT_LIST_FIRST_ROW + 1;
  if (height <= 0) {
    return [];
  }
  return sheet
    .getRange(PRODUCT_LIST_FIRST_ROW, 1, height, 1)
    .getValues()
    .map(([name], at) => ({ row: PRODUCT_LIST_FIRST_ROW + at, name: cleanLabel(name) }))
    .filter(({ name }) => name !== '');
}

// The owner's rule (2026-10-25): a script writes no product name the list lacks, as Sheets
// refuses one typed (PRODUCT_LIST_STRICT). -> the normalizeKey of every name of the list, or null
// while there is no list to judge by: no list tab yet, or a copy whose IMPORTRANGE failed ("#REF!").
function listNameKeys(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(PRODUCT_LIST_TAB);
  const names = sheet === null ? [] : readListNames(sheet).map(({ name }) => name);
  if (names.length === 0 || names.some((name) => name.startsWith('#'))) {
    return null;
  }
  return new Set(names.map(normalizeKey));
}

// The sheet of the half-month of `today` of each shop, read. { models, missing: [tab names] }
function readCurrentModels(spreadsheet, today) {
  const position = positionOfDate(today);
  const found = CHAINS.map((chain) => {
    const name = chainSheetName(chain, position);
    return { chain, name, sheet: spreadsheet.getSheetByName(name) };
  });
  return {
    models: Object.fromEntries(
      found.filter(({ sheet }) => sheet !== null).map(({ chain, sheet }) => [chain.key, readModel(sheet)])
    ),
    missing: found.filter(({ sheet }) => sheet === null).map(({ name }) => name),
  };
}

function runSetupProductList(spreadsheet, today) {
  const table = spreadsheet.getSheetByName(DECISION_TAB);
  if (table === null) {
    return [`ЗОГСЛОО: "${DECISION_TAB}" tab олдсонгүй`];
  }
  const plan = planNameDecisions(readDecisionRows(table));
  if (plan.undecided.length > 0 || plan.problems.length > 0) {
    return [
      `ЗОГСЛОО, юу ч өөрчлөөгүй: "${DECISION_TAB}" tab-ыг засна уу`,
      ...(plan.undecided.length > 0 ? [`шийдээгүй мөр (тийм/үгүй хоосон): ${plan.undecided.join(', ')}`] : []),
      ...plan.problems,
    ];
  }
  const current = readCurrentModels(spreadsheet, today);
  if (current.missing.length > 0 || !current.models[SORONZON_KEY] || !current.models[MY_HOME_KEY]) {
    return [`ЗОГСЛОО: энэ хагас сарын sheet олдсонгүй: ${current.missing.join(', ') || 'Соронзон / My home'}`];
  }
  const rows = planProductList(current.models[SORONZON_KEY], current.models[MY_HOME_KEY], renameMaps(plan.fixes));
  return [
    modeLabel(),
    ...runUnifyNames(spreadsheet, plan.fixes),
    writeProductList(spreadsheet, rows),
    runInstallNameRules(spreadsheet),
  ];
}

// Makes the list tab. An existing list is left as it is: it may hold the owner's edits.
function writeProductList(spreadsheet, rows) {
  const existing = spreadsheet.getSheetByName(PRODUCT_LIST_TAB);
  if (existing !== null && existing.getLastRow() >= PRODUCT_LIST_FIRST_ROW) {
    return `"${PRODUCT_LIST_TAB}" tab аль хэдийн байна - хөндсөнгүй`;
  }
  const headings = rows.filter(([, shops]) => shops === '').length;
  const summary = `"${PRODUCT_LIST_TAB}" tab: ${rows.length - headings} бараа, ${headings} ангилал`;
  if (DRY_RUN || rows.length === 0) {
    return `${summary} - ${DRY_RUN ? 'үүсгэнэ' : 'бараа олдсонгүй, үүсгэсэнгүй'}`;
  }
  const sheet = existing || spreadsheet.insertSheet(PRODUCT_LIST_TAB, 0);
  sheet.getRange(1, 1, PRODUCT_LIST_HEADER.length, 2).setValues(PRODUCT_LIST_HEADER);
  sheet.getRange(PRODUCT_LIST_HEADER.length, 1, 1, 2).setFontWeight('bold');
  const body = sheet.getRange(PRODUCT_LIST_FIRST_ROW, 1, rows.length, 2);
  body.setNumberFormat('@'); // "26-55" stays a name, not a date
  body.setValues(rows);
  rows.forEach(([, shops], at) => {
    if (shops === '') {
      sheet.getRange(PRODUCT_LIST_FIRST_ROW + at, 1).setFontWeight('bold');
    }
  });
  sheet.setFrozenRows(PRODUCT_LIST_HEADER.length);
  sheet.setColumnWidth(1, PRODUCT_LIST_NAME_WIDTH);
  sheet.setColumnWidth(2, PRODUCT_LIST_SHOP_WIDTH);
  return `${summary} - үүсгэлээ`;
}

// Column A of every year sheet: the dropdown of the list (or the "no name twice" rule
// while there is no list).
function runInstallNameRules(spreadsheet) {
  const sheets = findChains(spreadsheet).flatMap((chain) => chain.sheets);
  if (DRY_RUN) {
    return `A баганын dropdown ("${PRODUCT_LIST_TAB}" жагсаалтаас): ${sheets.length} sheet-д тавина`;
  }
  const changed = sheets.filter((sheet) => installNameRule(sheet)).length;
  return `A баганын dropdown ("${PRODUCT_LIST_TAB}" жагсаалтаас): ${changed} sheet-д тавьлаа`;
}

function isProductListEdit(event) {
  return event.range.getSheet().getName() === PRODUCT_LIST_TAB;
}

// A list whose first name cell is a formula (IMPORTRANGE) is a copy of the real one.
// A copy of the list: filled by IMPORTRANGE (the home Тосгон), or by value with the copy's
// title in A1 (another year's Тосгон, year-rollover.gs refreshListValues).
function isListCopy(sheet) {
  return (
    sheet !== null &&
    (sheet.getRange(PRODUCT_LIST_FIRST_ROW, 1).getFormula() !== '' || sheet.getRange(1, 1).getValue() === LIST_COPY_TITLE)
  );
}

// Something typed or pasted over the price view (C..E below the formula's row) is cleared at
// once: a value there stops the C4 formula from filling the column (#REF!), and a price belongs
// on the shop sheet. One emptied cell (Delete) is no typing. -> true when something was cleared
function clearListViewEdit(event) {
  const range = event.range;
  const isOneCell = range.getNumRows() === 1 && range.getNumColumns() === 1;
  if (isOneCell && (event.value === undefined || event.value === '')) {
    return false;
  }
  const firstRow = Math.max(range.getRow(), PRODUCT_LIST_FIRST_ROW + 1);
  const firstColumn = Math.max(range.getColumn(), LIST_VIEW_FIRST_COLUMN);
  const lastColumn = Math.min(range.getLastColumn(), LIST_VIEW_LAST_COLUMN);
  if (firstRow > range.getLastRow() || firstColumn > lastColumn) {
    return false;
  }
  const sheet = range.getSheet();
  if (isListCopy(sheet)) {
    return false;
  }
  sheet.getRange(firstRow, firstColumn, range.getLastRow() - firstRow + 1, lastColumn - firstColumn + 1).clearContent();
  showToast(sheet.getParent(), LIST_VIEW_MESSAGE);
  return true;
}

// One name typed over in the list: the same product gets the new name on every
// sheet of both shops, and the rename is logged for the workbooks with a copy.
// A new row, a category heading, a name the list already has or a copy of the
// list changes nothing.
function runProductListEdit(event) {
  if (clearListViewEdit(event)) {
    return;
  }
  const range = event.range;
  const isOneCell = range.getNumRows() === 1 && range.getNumColumns() === 1;
  if (!isOneCell || range.getColumn() !== 1 || range.getRow() < PRODUCT_LIST_FIRST_ROW) {
    return;
  }
  const sheet = range.getSheet();
  if (isListCopy(sheet)) {
    return;
  }
  const spreadsheet = sheet.getParent();
  const others = readListNames(sheet)
    .filter(({ row }) => row !== range.getRow())
    .map(({ name }) => name);
  const plan = planListEdit(event.oldValue, range.getValue(), others);
  if (!plan.rename) {
    showToast(spreadsheet, plan.reason);
    return;
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(PRODUCT_LIST_LOCK_WAIT_MS)) {
    showToast(spreadsheet, `Өөр ажил явагдаж байсан тул "${plan.to}" нэр дэлгүүрүүдэд солигдсонгүй. Нэрийг дахин бичнэ үү.`);
    return;
  }
  try {
    const fixes = Object.fromEntries(CHAINS.map((chain) => [chain.key, [[plan.to, [plan.from]]]]));
    const lines = runUnifyNames(spreadsheet, fixes);
    logListRename(spreadsheet, plan.from, plan.to);
    Logger.log('%s\n%s', modeLabel(), lines.join('\n'));
    showToast(spreadsheet, `"${plan.to}": ${lines[0]}`);
  } finally {
    lock.releaseLock();
  }
}

function logListRename(spreadsheet, fromKey, toName) {
  if (DRY_RUN) {
    return;
  }
  const log =
    spreadsheet.getSheetByName(LIST_RENAME_LOG_TAB) ||
    (() => {
      const made = spreadsheet.insertSheet(LIST_RENAME_LOG_TAB);
      made.getRange(1, 1, 1, 3).setValues([['Огноо', 'Хуучин нэр (түлхүүр)', 'Шинэ нэр']]);
      return made;
    })();
  const stamp = Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm');
  const row = log.getLastRow() + 1;
  log.getRange(row, 1, 1, 3).setNumberFormat('@').setValues([[stamp, fromKey, toName]]);
}

// The renames logged in the source workbook that this workbook (with a copy of the
// list) has not carried yet, carried now into its own sheets, one by one (a name may
// be renamed twice). source: the workbook of the real list.
function runListRenames(spreadsheet, source) {
  const log = source.getSheetByName(LIST_RENAME_LOG_TAB);
  const height = log === null ? 0 : log.getLastRow() - 1;
  if (height <= 0) {
    return [];
  }
  const rows = log.getRange(2, 1, height, 3).getValues().filter(([, from, to]) => cleanLabel(from) !== '' && cleanLabel(to) !== '');
  const properties = bookProperties();
  const done = Number(properties.getProperty(LIST_RENAMES_DONE_KEY) || 0);
  if (rows.length <= done) {
    return [];
  }
  // All waiting renames in ONE pass over the sheets (runUnifyNames reads every sheet and saves the
  // lists anew once): 185 renames one by one did not fit in the hourly job (2026-10-21).
  const fixes = chainRenames(rows.slice(done).map(([, from, to]) => [normalizeKey(from), cleanLabel(to)]));
  const lines = [
    `жагсаалтын нэр солилт: ${rows.length - done} нэр`,
    ...runUnifyNames(spreadsheet, Object.fromEntries(CHAINS.map((chain) => [chain.key, fixes]))),
  ];
  if (!DRY_RUN) {
    properties.setProperty(LIST_RENAMES_DONE_KEY, String(rows.length));
  }
  return lines;
}

// Renames in the order they were made ([old key, new name]) -> runUnifyNames' fixes
// ([[final name, [old keys]]]): a name renamed twice goes straight to its last name.
function chainRenames(renames) {
  const finalOf = renames.reduce((names, [from, to]) => {
    const updated = Object.fromEntries(Object.entries(names).map(([key, name]) => [key, normalizeKey(name) === from ? to : name]));
    return { ...updated, [from]: to };
  }, {});
  const byName = Object.entries(finalOf).reduce((all, [key, name]) => ({ ...all, [name]: [...(all[name] || []), key] }), {});
  return Object.entries(byName).map(([name, keys]) => [name, keys.filter((key) => key !== normalizeKey(name))]).filter(([, keys]) => keys.length > 0);
}

// Hourly: only a workbook with a copy of the list follows the renames of the real one.
function runListRenamesFromSource(spreadsheet) {
  if (!isListCopy(spreadsheet.getSheetByName(PRODUCT_LIST_TAB))) {
    return [];
  }
  return runListRenames(spreadsheet, partnerBook(spreadsheet, 'main', PRODUCT_LIST_SOURCE_ID)); // the same year's list
}
