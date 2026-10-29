/**
 * The factory's spelling of every brand, and the products of every category in name order, in
 * both shops: the English (Latin) names first, then the Mongolian (Cyrillic) ones, each group by
 * letters and then numbers ("Meiling 86l" before "Meiling 112l").
 *
 * The owner's rule (2026-10-21): a product's name is written the way its factory writes it
 * ("Cmidea 723l" -> "Midea 723l", "Panasonica 55" -> "Panasonic 55", "TGL 201l" -> "TCL 201l"),
 * every name starts with a capital letter ("тогтоогч дунд" -> "Тогтоогч дунд"), and the rows of
 * a category stand in name order, a new product too.
 *
 * tidyProductNames does it once, in order:
 *   1. every name of the list "Бараа" whose brand word is misspelt gets the factory's spelling and
 *      a capital first letter:
 *      in the list, on every sheet of the year of both shops (runUnifyNames, name-fixes.gs),
 *      in the Тосгон name table (_Тосгон нэр тулгалт: the list names in E, H, J), and in the
 *      rename log, so Тосгон's own sheets follow within the hour (runListRenames, product-list.gs);
 *      a new name another product already has is left alone and reported;
 *   2. the product rows of every category are put in name order, in the list and on every sheet
 *      of the year of both shops. A row moves with everything typed in it (values, colours,
 *      notes, number formats) up to the count column, and with its cells of the count blocks
 *      after the totals ("Тооллого 10/23", "Үлдэгдэл 10/23"); the computed columns are array
 *      formulas of row 3 and follow by themselves.
 * Afterwards the order stays:
 *   - a product the sync copies to the later sheets goes in at its place (findInsertAfterRow,
 *     year-sync.gs);
 *   - a row typed anywhere by hand, in a shop sheet or in the list, is moved to its place by the
 *     hourly job between ORDER_NIGHT_FROM_HOUR and ORDER_NIGHT_TO_HOUR only, so no row moves
 *     while someone is typing in it.
 * A sheet with a formula that points at a cell is not reordered (the formula would point at the
 * wrong row); the log names it.
 *
 * Usage: tidyProductNames with DRY_RUN = true (config.gs) and read the log, then with false.
 * Needs config.gs, year-sync.gs, name-fixes.gs, product-list.gs, dividers.gs, tosgon-names.gs.
 */

// A brand word (letters only, lower case, Latin look-alikes folded) -> the factory's spelling.
const OFFICIAL_BRAND_WORDS = {
  cmidea: 'Midea', cmedia: 'Midea', midea: 'Midea',
  panasonica: 'Panasonic', panasonico: 'Panasonic', pnasonica: 'Panasonic', pana: 'Panasonic', panasonic: 'Panasonic',
  tgl: 'TCL', tcl: 'TCL',
  xinix: 'XINGX', xingx: 'XINGX',
  skyworch: 'Skyworth', skywortch: 'Skyworth', skyworth: 'Skyworth',
  mieiling: 'Meiling', miling: 'Meiling', meiling: 'Meiling',
  acuma: 'Aucma', aucma: 'Aucma',
  haeir: 'Haier', haier: 'Haier',
  asewise: 'Acewise', acewise: 'Acewise',
  dawoo: 'Daewoo', daewoo: 'Daewoo',
  pilipis: 'Philips', plipis: 'Philips', philips: 'Philips',
  gallanz: 'Galanz', galanz: 'Galanz',
  ronshin: 'Ronshen', ronshen: 'Ronshen',
  akira: 'Akira', konka: 'Konka', hisense: 'Hisense', sharp: 'Sharp', hyundai: 'Hyundai',
  frestech: 'Frestech', royalstar: 'Royalstar', sevenstar: 'Sevenstar',
};
// Whole names a word rule cannot mend (a split brand, a letter O typed for a zero).
const OFFICIAL_NAME_FIXES = {
  'pa nasonica smart': 'Panasonic smart',
  'tgl -12ol': 'TCL 120L',
  'royalstar 1ol': 'Royalstar 10l',
  'smart 6 0l': 'Smart 6.0l',
};
const TOSGON_TABLE_NAME_COLUMNS = [5, 8, 10]; // E list name, H suggested name, J corrected name
const TOSGON_TABLE_FIRST_ROW = 3;
const ORDER_NIGHT_FROM_HOUR = 22; // the hourly job reorders typed rows from 22:00 ...
const ORDER_NIGHT_TO_HOUR = 7; // ... to 06:59 (the workbook's time zone)
// Mongolian Cyrillic in its alphabet's order.
const MONGOLIAN_ALPHABET = 'абвгдеёжзийклмноөпрстуүфхцчшщъыьэюя';
// Cyrillic letters typed inside a Latin name ("аkira") read as the Latin letter they look like,
// and Latin ones inside a Mongolian name ("Tоль") as the Cyrillic letter.
const LATIN_LOOKALIKES = { а: 'a', е: 'e', о: 'o', р: 'p', с: 'c', х: 'x', у: 'y', к: 'k' };
const CYRILLIC_LOOKALIKES = { a: 'а', e: 'е', o: 'о', p: 'р', c: 'с', x: 'х', y: 'у', k: 'к', t: 'т', m: 'м', h: 'н', b: 'в' };
const ENGLISH_GROUP = 0;
const MONGOLIAN_GROUP = 1;
const CELL_REFERENCE = /\$?[A-Z]{1,3}\$?\d+/;
// tidyProductNames goes in steps and keeps its plan in a script property until every step is done,
// so a run Google stops at its 6 minutes, or the time budget, goes on from where it was:
//   list    the list's names (the renames planned from the list, kept in TIDY_STATE_KEY)
//   shops   every year sheet of both shops, sheet by sheet (next: the sheet to do)
//   log     the rename log for Тосгон's own sheets, and the Тосгон name table
//   order   the name order: the list, then sheet by sheet
const TIDY_STATE_KEY = 'tidyNamesState';
const TIDY_HANDLER = 'tidyProductNames';
const TIDY_TIME_BUDGET_MS = 4 * 60 * 1000;
const TIDY_CONTINUE_AFTER_MS = 60 * 1000;
const TIDY_SAFETY_AFTER_MS = 8 * 60 * 1000; // a part Google stopped (6 minutes) is followed all the same
const TIDY_STEPS = ['list', 'shops', 'log', 'order', 'done'];
const TIDY_BUSY_TEXT = 'нэр засвар (tidyProductNames) дуусаагүй байгаа тул алгаслаа';

// While a tidy is under way the syncs wait (year-sync.gs runSync, runBookJobs; price-sync.gs):
// a year renamed halfway would look to them like products removed and others added.
const isTidyInProgress = () => PropertiesService.getScriptProperties().getProperty(TIDY_STATE_KEY) !== null;

// Does as much as TIDY_TIME_BUDGET_MS allows and, when work is left, starts itself again a minute
// later (a one-off trigger) until everything is done; the log says "БҮГД ДУУСЛАА" then.
// The next part is set up BEFORE the work (TIDY_SAFETY_AFTER_MS, after Google's 6 minutes): a part
// Google stops never gets to its end, and the chain would end with it (2026-10-21 23:55).
function tidyProductNames() {
  withScriptLock(() => {
    clearTidyTriggers();
    if (!DRY_RUN) {
      ScriptApp.newTrigger(TIDY_HANDLER).timeBased().after(TIDY_SAFETY_AFTER_MS).create();
    }
    const result = runTidyProductNames(currentBook(), Date.now() + TIDY_TIME_BUDGET_MS);
    clearTidyTriggers();
    if (!result.done && !DRY_RUN) {
      ScriptApp.newTrigger(TIDY_HANDLER).timeBased().after(TIDY_CONTINUE_AFTER_MS).create();
    }
    Logger.log('%s\n%s', modeLabel(), result.lines.join('\n'));
  });
}

// Two names the owner says are one product (2026-10-21: Соронзон's "TGL -12OL" is My home's
// "TCL 120L"): the first becomes the second on every sheet of both shops, the rename is logged for
// Тосгон, and in the list the first row goes while the second is marked for both shops.
const MERGE_NAMES = [['TGL -12OL', 'TCL 120L']];
const BOTH_SHOPS = 'Хоёулаа';

function mergeProductNames() {
  withScriptLock(() => Logger.log('%s\n%s', modeLabel(), runMergeNames(currentBook(), MERGE_NAMES).join('\n')));
}

function runMergeNames(spreadsheet, pairs) {
  if (isTidyInProgress()) {
    return [TIDY_BUSY_TEXT];
  }
  const fixes = pairs.map(([from, to]) => [to, [normalizeKey(from)]]);
  const lines = runUnifyNames(spreadsheet, Object.fromEntries(CHAINS.map((chain) => [chain.key, fixes])));
  pairs.forEach(([from, to]) => logListRename(spreadsheet, normalizeKey(from), to));
  const list = spreadsheet.getSheetByName(PRODUCT_LIST_TAB);
  if (list === null) {
    return lines;
  }
  const rows = readListNames(list);
  const rowOf = (name) => rows.find((entry) => normalizeKey(entry.name) === normalizeKey(name)) || null;
  const gone = pairs
    .map(([from, to]) => ({ from: rowOf(from), to: rowOf(to) }))
    .filter(({ from, to }) => from !== null && to !== null);
  gone.forEach(({ from, to }) => {
    lines.push(`"${PRODUCT_LIST_TAB}": "${from.name}" мөр ${DRY_RUN ? 'хасагдана' : 'хасагдлаа'}, "${to.name}" ${BOTH_SHOPS}`);
  });
  if (!DRY_RUN) {
    gone.forEach(({ to }) => list.getRange(to.row, 2).setValue(BOTH_SHOPS));
    [...gone].sort((a, b) => b.from.row - a.from.row).forEach(({ from }) => list.deleteRow(from.row));
  }
  return lines;
}

// Ends a tidy that is past its renames (the "order" step) without sorting the rest: the syncs and
// the hourly job go on again. A tidy still renaming is not stopped (half a year renamed).
function stopTidyProductNames() {
  withScriptLock(() => {
    const properties = PropertiesService.getScriptProperties();
    const saved = properties.getProperty(TIDY_STATE_KEY);
    const state = saved ? JSON.parse(saved) : null;
    if (state !== null && TIDY_STEPS.indexOf(state.step) < TIDY_STEPS.indexOf('order')) {
      Logger.log(`ЗОГСООГҮЙ: нэр солих алхам ("${state.step}") дуусаагүй байна. tidyProductNames-ийг үргэлжлүүлнэ үү.`);
      return;
    }
    if (!DRY_RUN) {
      properties.deleteProperty(TIDY_STATE_KEY);
      clearTidyTriggers();
    }
    Logger.log(state === null ? 'нэр засвар явагдаагүй байна' : `${modeLabel()}: нэр засвар зогслоо (эрэмбэлэлт ${state.next}-р sheet дээр), sync дахин ажиллана`);
  });
}

function clearTidyTriggers() {
  ScriptApp.getProjectTriggers()
    .filter((trigger) => trigger.getHandlerFunction() === TIDY_HANDLER)
    .forEach((trigger) => ScriptApp.deleteTrigger(trigger));
}

// ---------- pure helpers (no Spreadsheet calls) ----------

const foldLookalikes = (text, table = LATIN_LOOKALIKES) => [...String(text).toLowerCase()].map((ch) => table[ch] || ch).join('');

// One word with the factory's brand spelling: "Cmidea16" -> "Midea 16", "аkira" -> "Akira".
function officialWord(word) {
  const match = foldLookalikes(word).match(/^([a-z]+)(.*)$/);
  if (!match || !OFFICIAL_BRAND_WORDS[match[1]]) {
    return word;
  }
  const rest = word.slice(match[1].length);
  return OFFICIAL_BRAND_WORDS[match[1]] + (/^\d/.test(rest) ? ` ${rest}` : rest);
}

// The name as the factory writes it, its first letter a capital (the owner's rule, 2026-10-21:
// "тогтоогч дунд" -> "Тогтоогч дунд"); the rest of a name stays as it is (spaces tidied).
function officialName(name) {
  const clean = cleanLabel(name);
  const fixed = OFFICIAL_NAME_FIXES[clean.toLowerCase()] || clean.split(' ').map(officialWord).join(' ');
  return fixed.charAt(0).toUpperCase() + fixed.slice(1);
}

// names: the list's product names. -> { renames: [{ from, to }], skipped: [messages] }; a new
// name that another product already has (or that two old names would share) is skipped.
function planOfficialRenames(names) {
  const taken = new Set(names.map(normalizeKey));
  return names.reduce(
    (plan, from) => {
      const to = officialName(from);
      if (to === cleanLabel(from)) {
        return plan;
      }
      if (normalizeKey(to) !== normalizeKey(from) && taken.has(normalizeKey(to))) {
        return { ...plan, skipped: [...plan.skipped, `"${from}" -> "${to}": ийм нэртэй бараа аль хэдийн байгаа тул хэвээр`] };
      }
      taken.add(normalizeKey(to));
      return { ...plan, renames: [...plan.renames, { from, to }] };
    },
    { renames: [], skipped: [] }
  );
}

// English when the name has more Latin letters than Cyrillic ones ("аkira 888", "304 7l"),
// Mongolian otherwise ("Tоль 250", "Ус буцалгагч 8037").
function nameGroup(name) {
  const text = String(name).toLowerCase();
  const latin = (text.match(/[a-z]/g) || []).length;
  const cyrillic = (text.match(/[а-яёөү]/g) || []).length;
  return cyrillic > latin ? MONGOLIAN_GROUP : ENGLISH_GROUP;
}

// "Meiling 112l" -> ["meiling ", 112, "l"]: text in lower case (look-alikes folded the way of the
// name's group), numbers as numbers.
function nameOrderParts(name) {
  const lower = cleanLabel(name).toLowerCase();
  const text = foldLookalikes(lower, nameGroup(name) === ENGLISH_GROUP ? LATIN_LOOKALIKES : CYRILLIC_LOOKALIKES);
  return (text.match(/\d+(?:[.,]\d+)?|\D+/g) || []).map((part) => (/^\d/.test(part) ? Number(part.replace(',', '.')) : part));
}

// Rank of a character: a space, then Latin a-z, then Mongolian Cyrillic in alphabet order, then
// anything else.
function charRank(ch) {
  if (ch === ' ') {
    return 0;
  }
  if (/[a-z]/.test(ch)) {
    return 100 + ch.charCodeAt(0);
  }
  const mongolian = MONGOLIAN_ALPHABET.indexOf(ch);
  return mongolian >= 0 ? 300 + mongolian : 1000 + ch.charCodeAt(0);
}

function compareText(a, b) {
  for (let i = 0; i < Math.min(a.length, b.length); i += 1) {
    const gap = charRank(a[i]) - charRank(b[i]);
    if (gap !== 0) {
      return gap;
    }
  }
  return a.length - b.length;
}

// Name order: the English names before the Mongolian ones; inside a group, letters in alphabet
// order and numbers by their value ("86" before "112"), a number before letters.
function compareProductNames(a, b) {
  const groupGap = nameGroup(a) - nameGroup(b);
  if (groupGap !== 0) {
    return groupGap;
  }
  const [left, right] = [nameOrderParts(a), nameOrderParts(b)];
  for (let i = 0; i < Math.min(left.length, right.length); i += 1) {
    const [x, y] = [left[i], right[i]];
    const gap =
      typeof x === 'number' && typeof y === 'number' ? x - y
        : typeof x === 'number' ? -1
        : typeof y === 'number' ? 1
        : compareText(x, y);
    if (gap !== 0) {
      return gap;
    }
  }
  return left.length - right.length;
}

// rows: [{ label, movable }] top to bottom. Each run of movable rows is put in name order; an
// unmovable row (a category heading, an empty row) stays where it is. -> the new order as
// indexes into rows (the identity when nothing moves).
function planRowOrder(rows) {
  const order = rows.map((_, index) => index);
  let start = 0;
  while (start < rows.length) {
    if (!rows[start].movable) {
      start += 1;
      continue;
    }
    let end = start;
    while (end < rows.length && rows[end].movable) {
      end += 1;
    }
    const run = order.slice(start, end).sort((x, y) => compareProductNames(rows[x].label, rows[y].label) || x - y);
    order.splice(start, run.length, ...run);
    start = end;
  }
  return order;
}

const isMoved = (order) => order.some((index, at) => index !== at);


// Where a product goes in its category: after the last product whose name comes before it.
// products: that category's products ({ label, row }) in row order; null: before them all.
function rowBeforeInOrder(products, label) {
  const before = products.filter((product) => compareProductNames(product.label, label) <= 0);
  return before.length === 0 ? null : before[before.length - 1].row;
}

function isOrderNight(hour) {
  return hour >= ORDER_NIGHT_FROM_HOUR || hour < ORDER_NIGHT_TO_HOUR;
}

// ---------- runs (Spreadsheet calls) ----------

// One part of the work, from the saved step on. -> { done, lines }
function runTidyProductNames(spreadsheet, deadline = Date.now() + TIDY_TIME_BUDGET_MS) {
  const found = spreadsheet.getSheetByName(PRODUCT_LIST_TAB);
  const list = found === null || isListCopy(found) ? null : found; // Тосгон's list is a copy: it follows the main one
  if (list === null && BOOK_KIND !== 'tosgon') {
    return { done: true, lines: [`"${PRODUCT_LIST_TAB}" tab олдсонгүй (эсвэл хуулбар байна): юу ч өөрчлөөгүй`] };
  }
  const sheets = yearShopSheets(spreadsheet);
  let state = readTidyState() || (list ? startTidyState(list) : startSheetsTidyState(sheets));
  const isNew = state.next === 0 && (state.step === 'list' || (list === null && state.step === 'shops' && state.cells === 0));
  const lines = isNew ? planLines(state) : [`үргэлжилж байна: "${state.step}" алхмаас (${state.renames.length} нэр)`];
  do {
    const result = TIDY_WORK[state.step](spreadsheet, list, sheets, state, deadline);
    lines.push(...result.lines);
    state = result.state;
    saveTidyState(state);
  } while (state.step !== 'done' && Date.now() <= deadline);
  const done = state.step === 'done';
  if (done && !DRY_RUN) {
    PropertiesService.getScriptProperties().deleteProperty(TIDY_STATE_KEY);
  }
  const left = DRY_RUN ? `ТУРШИЛТ: хугацаа хүрсэнгүй, "${state.step}" алхмаас цааш харагдсангүй` : `үлдсэнийг 1 минутын дараа өөрөө үргэлжлүүлнэ ("${state.step}" алхам)`;
  lines.push(done ? (DRY_RUN ? 'ТУРШИЛТ: дээрх нь хийгдэнэ' : 'БҮГД ДУУСЛАА') : left);
  return { done, lines };
}

function planLines(state) {
  return [
    `албан ёсны нэр: ${state.renames.length} бараа ${DRY_RUN ? 'солигдоно' : 'солигдоно (эхэллээ)'}`,
    ...state.renames.map(({ from, to }) => `  ${from} -> ${to}`),
    ...state.skipped,
  ];
}

function readTidyState() {
  if (DRY_RUN) {
    return null;
  }
  const saved = PropertiesService.getScriptProperties().getProperty(TIDY_STATE_KEY);
  return saved ? JSON.parse(saved) : null;
}

function saveTidyState(state) {
  if (!DRY_RUN) {
    PropertiesService.getScriptProperties().setProperty(TIDY_STATE_KEY, JSON.stringify(state));
  }
}

// The plan, made once from the list as it is before anything is renamed.
function startTidyState(list) {
  const rows = readListNames(list).filter(({ row }) => cleanLabel(list.getRange(row, 2).getValue()) !== '');
  const plan = planOfficialRenames(rows.map(({ name }) => name));
  return { step: 'list', next: 0, renames: plan.renames, skipped: plan.skipped, cells: 0, moved: 0 };
}

// Тосгон (its list is a copy of the main one): the plan from the names on its own sheets, from the shops step on.
function startSheetsTidyState(sheets) {
  const names = [...new Set(sheets.flatMap((sheet) => readModel(sheet).products.map((product) => product.label)))];
  const plan = planOfficialRenames(names);
  return { step: 'shops', next: 0, renames: plan.renames, skipped: plan.skipped, cells: 0, moved: 0 };
}

const nextTidyStep = (state, extra = {}) => ({ ...state, ...extra, step: TIDY_STEPS[TIDY_STEPS.indexOf(state.step) + 1], next: 0, trying: null });

// Runs work(item) for the items from state.next on while the time lasts (at least one), and
// saves the place after each one (saved(next)), so a run Google stops goes on from the next item.
// -> { next, finished }
// A sheet a part was working on when Google stopped it (state.trying) is not tried again: it is
// reported (skipped) and the next one goes on, so one heavy sheet can never hold the run forever.
function eachUntil(items, state, deadline, work, saved = () => {}, skipped = () => {}) {
  let next = state.next;
  let slowest = 0; // a sheet starts only when one as slow as the slowest so far still fits
  while (next < items.length && (next === state.next || Date.now() + slowest <= deadline)) {
    if (state.trying === next) {
      skipped(items[next]);
    } else {
      saved(next, next); // "trying this one": a stop from here on skips it next time
      const started = Date.now();
      work(items[next]);
      slowest = Math.max(slowest, Date.now() - started);
    }
    next += 1;
    saved(next, null);
  }
  return { next, finished: next >= items.length };
}

const skippedLine = (lines) => (sheet) => lines.push(`${sheet.getName()}: өмнөх удаа энэ sheet дээр зогссон тул алгаслаа (дахин tidyProductNames ажиллуулахад дахин оролдоно)`);

// work(sheet), and a failure on one sheet is reported in lines instead of stopping the run.
function onEachSheet(lines, work) {
  return (sheet) => {
    try {
      work(sheet);
    } catch (error) {
      const where = String(error.stack || '').split('\n').slice(1, 3).map((line) => line.trim()).join(' ← ');
      lines.push(`${sheet.getName()}: алдаа, алгаслаа (${error.message}) ${where}`);
    }
  };
}

const TIDY_WORK = {
  list: (spreadsheet, list, sheets, state) => {
    const newName = new Map(state.renames.map(({ from, to }) => [normalizeKey(from), to]));
    const height = Math.max(list.getLastRow() - PRODUCT_LIST_FIRST_ROW + 1, 0);
    if (height > 0 && !DRY_RUN) {
      const range = list.getRange(PRODUCT_LIST_FIRST_ROW, 1, height, 1);
      range.setValues(range.getValues().map(([name]) => [newName.get(normalizeKey(name)) || name]));
    }
    return { state: nextTidyStep(state), lines: [`"${PRODUCT_LIST_TAB}": ${state.renames.length} нэр ${DRY_RUN ? 'солигдоно' : 'солигдлоо'}`] };
  },
  shops: (spreadsheet, list, sheets, state, deadline) => {
    // Column A takes only names of the list (its dropdown refuses any other): a new name the list
    // does not have is left as it was and reported.
    const listSheet = spreadsheet.getSheetByName(PRODUCT_LIST_TAB);
    const listKeys = listSheet === null ? null : new Set(readListNames(listSheet).map(({ name }) => normalizeKey(name)));
    const allowed = state.renames.filter(({ to }) => listKeys === null || listKeys.has(normalizeKey(to)));
    const nameMap = buildNameMap(allowed.map(({ from, to }) => [to, [normalizeKey(from)]]));
    const lines = state.next === 0
      ? state.renames.filter((rename) => !allowed.includes(rename)).map(({ from, to }) => `"${from}" -> "${to}": жагсаалтад алга тул дэлгүүрүүдэд хэвээр`)
      : [];
    const run = eachUntil(sheets, state, deadline, onEachSheet(lines, (sheet) => {
      const count = renameOnSheet(sheet, nameMap, lines);
      if (count > 0) {
        lines.push(`${sheet.getName()}: ${count} нэр`);
      }
      state = { ...state, cells: state.cells + count };
    }), (next, trying) => saveTidyState({ ...state, next, trying }), skippedLine(lines));
    if (!run.finished) {
      return { state: { ...state, next: run.next, trying: null }, lines };
    }
    if (!DRY_RUN) {
      saveListsAnew(spreadsheet);
    }
    return { state: nextTidyStep(state), lines: [...lines, `дэлгүүрүүд: ${state.cells} нүд ${DRY_RUN ? 'солигдоно' : 'солигдлоо'}`] };
  },
  log: (spreadsheet, list, sheets, state) => ({
    state: nextTidyStep(state),
    lines: list === null ? [] : [logListRenames(spreadsheet, state.renames), runTosgonTableRenames(spreadsheet, state.renames)],
  }),
  order: (spreadsheet, list, sheets, state, deadline) => {
    const lines = [];
    if (state.next === 0 && list !== null) {
      onEachSheet(lines, (sheet) => lines.push(orderListRows(sheet)))(list);
    }
    const run = eachUntil(sheets, state, deadline, onEachSheet(lines, (sheet) => {
      const line = orderShopRows(sheet);
      if (line !== '') {
        lines.push(line);
        state = { ...state, moved: state.moved + 1 };
      }
    }), (next, trying) => saveTidyState({ ...state, next, trying }), skippedLine(lines));
    const shown = lines.filter((line) => line !== '');
    if (!run.finished) {
      return { state: { ...state, next: run.next, trying: null }, lines: shown };
    }
    if (!DRY_RUN && state.moved > 0) {
      saveListsAnew(spreadsheet);
    }
    return { state: nextTidyStep(state), lines: [...shown, `нэрийн дараалал: ${state.moved} хуудас ${DRY_RUN ? 'эрэмбэлэгдэнэ' : 'эрэмбэлэгдлээ'}`] };
  },
};

// The planned renames on one sheet, cell by cell: only the renamed cells are written, so a name
// elsewhere in column A that its dropdown would refuse is never written again; a new name the sheet
// already has elsewhere is left and reported (planRenames, name-fixes.gs). -> how many cells
function renameOnSheet(sheet, nameMap, lines) {
  const plan = planRenames(readModel(sheet).products, nameMap);
  plan.clashes.forEach((name) => lines.push(`${sheet.getName()}: "${name}" хэвээр (шинэ нэртэй мөр энэ sheet дээр бас байна)`));
  if (!DRY_RUN) {
    plan.renames.forEach((rename) => sheet.getRange(rename.row, SYNC_NAME_COL).setValue(rename.to));
    SpreadsheetApp.flush(); // a refused write shows here, with this sheet, not at the next one
  }
  return plan.renames.length;
}

// The renames in the log for the workbooks with a copy of the list, in one write (product-list.gs).
function logListRenames(spreadsheet, renames) {
  if (renames.length === 0) {
    return 'нэр солилтын бүртгэл: бичих зүйл алга';
  }
  if (DRY_RUN) {
    return `нэр солилтын бүртгэл: ${renames.length} мөр бичигдэнэ`;
  }
  const log =
    spreadsheet.getSheetByName(LIST_RENAME_LOG_TAB) ||
    (() => {
      const made = spreadsheet.insertSheet(LIST_RENAME_LOG_TAB);
      made.getRange(1, 1, 1, 3).setValues([['Огноо', 'Хуучин нэр (түлхүүр)', 'Шинэ нэр']]);
      return made;
    })();
  const stamp = Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm');
  log.getRange(log.getLastRow() + 1, 1, renames.length, 3)
    .setNumberFormat('@')
    .setValues(renames.map(({ from, to }) => [stamp, normalizeKey(from), to]));
  return `нэр солилтын бүртгэл: ${renames.length} мөр бичигдлээ`;
}

// The list names in the Тосгон name table follow the renames, so its answers still name list rows.
function runTosgonTableRenames(spreadsheet, renames) {
  const table = spreadsheet.getSheetByName(TOSGON_DECISION_TAB);
  const height = table === null ? 0 : table.getLastRow() - TOSGON_TABLE_FIRST_ROW + 1;
  if (height <= 0) {
    return `${TOSGON_DECISION_TAB}: хүснэгт алга`;
  }
  const newName = Object.fromEntries(renames.map(({ from, to }) => [normalizeKey(from), to]));
  let cells = 0;
  TOSGON_TABLE_NAME_COLUMNS.forEach((column) => {
    const range = table.getRange(TOSGON_TABLE_FIRST_ROW, column, height, 1);
    const values = range.getValues();
    const renamed = values.map(([value]) => [newName[normalizeKey(value)] || value]);
    cells += renamed.filter(([value], at) => value !== values[at][0]).length;
    if (!DRY_RUN) {
      range.setValues(renamed);
    }
  });
  return `${TOSGON_DECISION_TAB}: ${cells} нүд ${DRY_RUN ? 'солигдоно' : 'солигдлоо'}`;
}

// The entry sheets of both shops from the given half-month on (0: the whole year).
function yearShopSheets(spreadsheet, fromPosition = 0) {
  return CHAINS.flatMap((chain) => findChains(spreadsheet, { chainKey: chain.key, fromPosition }).flatMap((found) => found.sheets));
}

// Puts the products of every category in name order in the list and the given sheets.
function runNameOrder(spreadsheet, { list, sheets }) {
  const lines = [...(list === null ? [] : [orderListRows(list)]), ...sheets.map(orderShopRows)].filter((line) => line !== '');
  const moved = lines.filter((line) => line.includes(' мөр ')).length;
  if (!DRY_RUN && moved > 0) {
    saveListsAnew(spreadsheet);
  }
  return [`нэрийн дараалал: ${moved} хуудас ${DRY_RUN ? 'эрэмбэлэгдэнэ' : 'эрэмбэлэгдлээ'}`, ...lines];
}

// ---------- a sheet's names in the list's spelling ----------

// The list's spelling of a sheet's name: the same name in other spaces or capitals, or with the
// brand's old spelling ("Panasonica 65" -> "Panasonic 65"); undefined for a name off the list.
function listSpellingOf(label, spellings) {
  return spellings.get(normalizeKey(label)) || spellings.get(normalizeKey(officialName(label)));
}

// Every year sheet of both shops: a name the list spells otherwise gets the list's spelling, cell by
// cell (renameOnSheet); a name off the list, or one the sheet already has elsewhere, stays.
function runListSpelling(spreadsheet) {
  const list = spreadsheet.getSheetByName(PRODUCT_LIST_TAB);
  if (list === null || isListCopy(list)) {
    return [`"${PRODUCT_LIST_TAB}" жагсаалт энэ файлд алга (эсвэл хуулбар), юу ч хийсэнгүй`];
  }
  const spellings = new Map(readListNames(list).map(({ name }) => [normalizeKey(name), name]));
  const lines = [];
  let cells = 0;
  yearShopSheets(spreadsheet).forEach(onEachSheet(lines, (sheet) => {
    const nameMap = {};
    readModel(sheet).products.forEach((product) => {
      const spelling = listSpellingOf(product.label, spellings);
      if (spelling !== undefined && spelling !== product.label) {
        nameMap[product.key] = spelling;
      }
    });
    const count = renameOnSheet(sheet, nameMap, lines);
    if (count > 0) {
      lines.push(`${sheet.getName()}: ${count} нэр`);
    }
    cells += count;
  }));
  return [`${modeLabel()}: жагсаалтын бичлэгээр ${cells} нүд ${DRY_RUN ? 'засна' : 'заслаа'}`, ...lines];
}

function matchListSpelling() {
  withScriptLock(() => Logger.log(runListSpelling(currentBook()).join('\n')));
}

// Hourly, at night only: the list and the sheets from the current half-month on.
function runNightlyNameOrder(spreadsheet, now = new Date()) {
  const hour = Number(Utilities.formatDate(now, spreadsheet.getSpreadsheetTimeZone(), 'H'));
  if (!isOrderNight(hour)) {
    return [];
  }
  const list = spreadsheet.getSheetByName(PRODUCT_LIST_TAB);
  const lines = runNameOrder(spreadsheet, {
    list: list === null || isListCopy(list) ? null : list,
    sheets: yearShopSheets(spreadsheet, positionOfDate(now)),
  });
  return lines.length > 1 ? lines : [];
}

function orderListRows(list) {
  const height = list.getLastRow() - PRODUCT_LIST_FIRST_ROW + 1;
  if (height <= 1) {
    return '';
  }
  const range = list.getRange(PRODUCT_LIST_FIRST_ROW, 1, height, 2);
  const rows = range.getValues().map(([name, shops]) => ({ label: cleanLabel(name), movable: cleanLabel(name) !== '' && cleanLabel(shops) !== '' }));
  return moveRows(list, range, planRowOrder(rows));
}

function orderShopRows(sheet) {
  const height = sheet.getLastRow() - SYNC_HEADER_ROWS;
  if (height <= 1) {
    return '';
  }
  const range = sheet.getRange(SYNC_HEADER_ROWS + 1, 1, height, shopCountColumn(sheet));
  const values = range.getValues();
  const colors = sheet.getRange(SYNC_HEADER_ROWS + 1, 1, height, DIVIDER_SCAN_WIDTH).getBackgrounds();
  const rows = values.map((cells, at) => {
    const label = cleanLabel(cells[0]);
    return { label, movable: label !== '' && !isDividerRow(colors[at], cells.slice(0, DIVIDER_VALUE_WIDTH)) };
  });
  return moveRows(sheet, range, planRowOrder(rows), values, spotCountRanges(sheet, range));
}

// The layout of this sheet (its own: a sheet may be before the gift column).
function shopSheetLayout(sheet) {
  const header = (column) => sheet.getRange(LAYOUT_HEADER_ROW, column).getValue();
  return shopLayoutOf(header(SHELF_PRICE_COLUMN), DAY_COLUMNS, header(giftHeaderColumn(DAY_COLUMNS)));
}

const shopCountColumn = (sheet) => shopSheetLayout(sheet).count;

// The typed columns of the count blocks after the totals ("Тооллого 10/23" and its "Үлдэгдэл"),
// over the rows of `range`: they move with their rows too (2026-10-26: Tos/2026/10's BY stayed
// where it was when the rows above it were put in order, so its counts stood by other products).
// The block's third column is one array formula of the header row and follows by itself.
function spotCountRanges(sheet, range) {
  // looked for after the count column: Тосгон's sheets have their blocks one column before
  // spotFirstColumn of the shops' layout (2026-10-26: BY, not BZ)
  const first = shopSheetLayout(sheet).count + 1;
  const width = Math.max(sheet.getMaxColumns(), sheet.getLastColumn()) - first + 1;
  if (width <= 0) {
    return [];
  }
  return sheet.getRange(LAYOUT_HEADER_ROW, first, 1, width).getValues()[0]
    .map((header, offset) => (spotDayOf(header) === null ? null : first + offset))
    .filter((column) => column !== null)
    .map((column) => sheet.getRange(range.getRow(), column, range.getNumRows(), SPOT_BLOCK_WIDTH - 1));
}

// work() with the sheet's filter taken off and put back the same afterwards. A filter (the
// "Зөвхөн буруу" view, view-columns.gs) hides rows, and Google lifts a rule only on the rows it
// shows, while the values reach the hidden rows too: 13 sheets were refused (2026-10-22).
function withoutFilter(sheet, work) {
  const filter = sheet.getFilter();
  if (filter === null) {
    return work();
  }
  const range = filter.getRange();
  const criteria = [];
  for (let column = range.getColumn(); column < range.getColumn() + range.getNumColumns(); column += 1) {
    const criterion = filter.getColumnFilterCriteria(column);
    if (criterion !== null) {
      criteria.push([column, criterion]);
    }
  }
  filter.remove();
  try {
    return work();
  } finally {
    const again = sheet.getRange(range.getRow(), range.getColumn(), range.getNumRows(), range.getNumColumns()).createFilter();
    criteria.forEach(([column, criterion]) => again.setColumnFilterCriteria(column, criterion));
  }
}

// Writes the rows of `range` in `order`, each with its values, colours, notes and number formats;
// the same rows of `extra` (ranges of other columns over the same rows) move with them.
function moveRows(sheet, range, order, values = range.getValues(), extra = []) {
  if (!isMoved(order)) {
    return '';
  }
  const formulas = range.getFormulas();
  const extraFormulas = extra.map((part) => part.getFormulas());
  if ([formulas, ...extraFormulas].some((grid) => grid.some((row) => row.some((formula) => CELL_REFERENCE.test(formula))))) {
    return `${sheet.getName()}: нүд заасан томьёотой тул эрэмбэлээгүй`;
  }
  const moved = order.filter((index, at) => index !== at).length;
  if (!DRY_RUN) {
    const pick = (grid) => order.map((index) => grid[index]);
    const read = (part, partValues, partFormulas) => ({
      part,
      cells: partValues.map((row, r) => row.map((value, c) => partFormulas[r][c] || value)),
      backgrounds: part.getBackgrounds(),
      fontColors: part.getFontColors(),
      notes: part.getNotes(),
      formats: part.getNumberFormats(),
    });
    const parts = [read(range, values, formulas), ...extra.map((part, index) => read(part, part.getValues(), extraFormulas[index]))];
    // Column A's rule (the list's dropdown, or the old "no name twice") checks each cell as it is
    // written, and halfway through a move a name stands twice: the rule is lifted from the first
    // moved row to the sheet's bottom and put back as ONE rule (the list's dropdown; a sheet that
    // still had the old rule gets the dropdown too). Per-cell rules of a whole sheet are slow to
    // write back (2 minutes for a Соронзон sheet, 2026-10-21), one rule is not.
    withoutFilter(sheet, () => {
      const names = sheet.getRange(range.getRow(), SYNC_NAME_COL, sheet.getMaxRows() - range.getRow() + 1, 1);
      const hadRule = names.getDataValidations().some(([rule]) => rule !== null);
      names.clearDataValidations();
      SpreadsheetApp.flush();
      try {
        parts.forEach(({ part, cells, backgrounds, fontColors, notes, formats }) => {
          part.setValues(pick(cells));
          part.setBackgrounds(pick(backgrounds));
          part.setFontColors(pick(fontColors));
          part.setNotes(pick(notes));
          part.setNumberFormats(pick(formats));
        });
        SpreadsheetApp.flush();
      } finally {
        if (hadRule && sheet.getName() !== PRODUCT_LIST_TAB) {
          names.setDataValidation(buildNameRule(sheet.getParent()));
        }
        SpreadsheetApp.flush();
      }
    });
  }
  return `${sheet.getName()}: ${moved} мөр байраа ${DRY_RUN ? 'солино' : 'сольлоо'}`;
}
