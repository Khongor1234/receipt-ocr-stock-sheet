/**
 * The new year's workbook, made by itself from December (the owner's rule, 2026-10-21): on
 * 20 December the files of the coming year ("бараа орлого зарлага 2027", "Тосгон 2027") are made
 * from December's products, categories and prices, and from 1 January every system works on them
 * (year-books.gs). The old files stay as they are, never locked.
 *
 * Each project makes its own kind (BOOK_KIND): the main project the shops' workbook, Тосгон's
 * project Тосгон's. The hourly job (runBackgroundJobs) calls runYearRollover, which goes on from
 * where it stopped, so it can stop at any point and is safe to run again:
 *   from 20 December:
 *     create   a new spreadsheet in the old one's folder
 *     seed     the last sheet of each chain (Sor/2026/12/31, My/2026/12/31; Tos/2026/12) copied
 *              as the first one of the new year (Sor/2027/1/15 ...; Tos/2027/1) with its income,
 *              expense, count and opening balance emptied (copyAsNextPeriod's rule), the product
 *              list and Sheet200 copied, the layout flags of the old file set for the new one
 *     share    the old file's editors and viewers, and READER_ACCOUNTS (a read-only service account), get the same
 *     register the year's row of the registry gets this file's id, from 1 January
 *     setup    setupYear on the new file (year-setup.gs, in parts by itself): the other sheets,
 *              headers, formulas, protections, totals, charts, its edit triggers
 *     mail     the owner gets the links
 *   until 31 December, every hour: products added in December and changed prices reach the new
 *     year's first sheet (syncIntoNewYear);
 *   in January, every hour: the counts of 31 December become the new year's opening balances,
 *     also when they are typed late (carryIntoNewYear).
 * Тосгон's new file gets its product list as values from the new main file (no IMPORTRANGE, so
 * nobody has to allow access); the hourly job keeps it up to date (refreshListValues).
 *
 * rehearseNewYear does the same now with files called "... (туршилт)", not registered, so the
 * result can be looked at before December. Needs year-books.gs, year-setup.gs, year-sync.gs.
 */

const ROLLOVER_MONTH = 11; // December (0-based)
const ROLLOVER_DAY = 20;
const ROLLOVER_STEPS = ['create', 'seed', 'share', 'register', 'setup', 'wait', 'mail', 'done'];
const ROLLOVER_STATE_KEY = 'rollover'; // script property "rollover:<kind>:<year>[:rehearsal]": { id, step }
const SETUP_BOOK_KEY = 'setupBook'; // script property: the workbook setupYear works on, when not today's
const YEAR_SETUP_DONE_KEY = 'yearSetupDone'; // that workbook's property once setupYear finished
const BOOK_TITLES = { main: 'бараа орлого зарлага', tosgon: 'Тосгон' };
const REHEARSAL_SUFFIX = ' (туршилт)';
const READER_ACCOUNTS = ['sheet-reader@YOUR_GCP_PROJECT.iam.gserviceaccount.com']; // a daily build that only reads the files
const COPIED_TABS = ['Бараа', 'Sheet200'];
const LAYOUT_FLAGS = ['pairedLayoutReady', 'giftColumnReady', 'priceLinkReady'];
const ROLLOVER_MAIL_SUBJECT = 'Шинэ оны файл';
const ROLLOVER_REHEARSAL_BUDGET_MS = 4 * 60 * 1000;

function rehearseNewYear() {
  withScriptLock(() => {
    const today = new Date();
    const day = new Date(today.getFullYear(), ROLLOVER_MONTH, ROLLOVER_DAY);
    const lines = runYearRollover(day, Date.now() + ROLLOVER_REHEARSAL_BUDGET_MS, { rehearsal: true });
    Logger.log('%s\n%s', modeLabel(), lines.join('\n'));
  });
}

// ---------- pure helpers (no Spreadsheet calls) ----------

// 'prepare' from 20 to 31 December, 'carry' in January, null otherwise.
function rolloverPhase(date) {
  if (date.getMonth() === ROLLOVER_MONTH && date.getDate() >= ROLLOVER_DAY) {
    return 'prepare';
  }
  return date.getMonth() === 0 ? 'carry' : null;
}

// The year whose files are being made or carried into.
const newYearOf = (date) => (date.getMonth() === ROLLOVER_MONTH ? date.getFullYear() + 1 : date.getFullYear());

const bookTitle = (kind, year, rehearsal) => `${BOOK_TITLES[kind]} ${year}${rehearsal ? REHEARSAL_SUFFIX : ''}`;

const nextStep = (step) => ROLLOVER_STEPS[Math.min(ROLLOVER_STEPS.indexOf(step) + 1, ROLLOVER_STEPS.length - 1)];

// ---------- the run ----------

function runYearRollover(today, deadline, { rehearsal = false } = {}) {
  const phase = rehearsal ? 'prepare' : rolloverPhase(today);
  if (phase === null) {
    return [];
  }
  if (readRegistry().length === 0 && !rehearsal) {
    return ['шинэ он: бүртгэл (REGISTRY_ID) алга эсвэл хоосон тул хийсэнгүй'];
  }
  const year = newYearOf(today);
  const old = rehearsal ? SpreadsheetApp.getActiveSpreadsheet() : bookForDate(new Date(year - 1, ROLLOVER_MONTH, 31));
  const state = readRolloverState(year, rehearsal);
  const lines = state.step === 'done' ? [] : prepareNewBook(old, year, state, deadline, rehearsal);
  const now = readRolloverState(year, rehearsal);
  if (now.id && ROLLOVER_STEPS.indexOf(now.step) > ROLLOVER_STEPS.indexOf('seed')) {
    const book = SpreadsheetApp.openById(now.id);
    lines.push(...(phase === 'prepare' ? syncIntoNewYear(old, book) : carryIntoNewYear(old, book)));
    lines.push(...refreshListValues(book));
  }
  useBook(old);
  return lines.map((line) => `шинэ он ${year}: ${line}`);
}

function stateKey(year, rehearsal) {
  return [ROLLOVER_STATE_KEY, BOOK_KIND, year, ...(rehearsal ? ['rehearsal'] : [])].join(':');
}

function readRolloverState(year, rehearsal) {
  const saved = PropertiesService.getScriptProperties().getProperty(stateKey(year, rehearsal));
  return saved ? JSON.parse(saved) : { id: null, step: 'create' };
}

function saveRolloverState(year, rehearsal, state) {
  if (!DRY_RUN) {
    PropertiesService.getScriptProperties().setProperty(stateKey(year, rehearsal), JSON.stringify(state));
  }
}

// Does the steps from the saved one on, while the time lasts; each one saved when done.
function prepareNewBook(old, year, start, deadline, rehearsal) {
  const lines = [];
  let state = start;
  while (state.step !== 'done' && Date.now() <= deadline) {
    const result = ROLLOVER_WORK[state.step](old, year, state, rehearsal);
    lines.push(...result.lines);
    if (!result.done) {
      break;
    }
    state = { ...state, ...(result.state || {}), step: nextStep(state.step) };
    saveRolloverState(year, rehearsal, state);
    if (DRY_RUN) {
      lines.push(`(туршилт: "${state.step}" алхмаас цааш үргэлжлэхгүй)`);
      break;
    }
  }
  return lines;
}

const ROLLOVER_WORK = {
  create: (old, year, state, rehearsal) => {
    const title = bookTitle(BOOK_KIND, year, rehearsal);
    if (DRY_RUN) {
      return { done: true, lines: [`"${title}" файл үүсгэнэ`] };
    }
    const book = SpreadsheetApp.create(title);
    const folders = DriveApp.getFileById(old.getId()).getParents();
    if (folders.hasNext()) {
      DriveApp.getFileById(book.getId()).moveTo(folders.next());
    }
    return { done: true, state: { id: book.getId() }, lines: [`"${title}" файл үүслээ: ${book.getUrl()}`] };
  },
  seed: (old, year, state) => ({ done: true, lines: seedNewBook(old, SpreadsheetApp.openById(state.id)) }),
  share: (old, year, state) => {
    const file = DriveApp.getFileById(state.id);
    const source = DriveApp.getFileById(old.getId());
    const emails = (users) => users.map((user) => user.getEmail()).filter((email) => email !== '');
    const editors = emails(source.getEditors());
    const viewers = [...emails(source.getViewers()), ...READER_ACCOUNTS];
    if (!DRY_RUN) {
      editors.forEach((email) => file.addEditor(email));
      viewers.forEach((email) => file.addViewer(email));
    }
    return { done: true, lines: [`эрх: ${editors.length} засагч, ${viewers.length} уншигч`] };
  },
  register: (old, year, state, rehearsal) => ({
    done: true,
    lines: [rehearsal ? 'бүртгэлд оруулаагүй (туршилт)' : registerBook(year, state.id)],
  }),
  setup: (old, year, state) => {
    if (!DRY_RUN) {
      PropertiesService.getScriptProperties().setProperty(SETUP_BOOK_KEY, state.id);
      scheduleSetupRun(SETUP_CONTINUE_AFTER_MS);
    }
    return { done: true, lines: ['setupYear эхэллээ (хэсэг хэсгээр өөрөө үргэлжилнэ)'] };
  },
  wait: (old, year, state) => {
    const isDone = bookProperties(SpreadsheetApp.openById(state.id)).getProperty(YEAR_SETUP_DONE_KEY) === '1';
    return { done: isDone, lines: [isDone ? 'setupYear дууслаа' : 'setupYear үргэлжилж байна'] };
  },
  mail: (old, year, state, rehearsal) => {
    const book = SpreadsheetApp.openById(state.id);
    if (!DRY_RUN) {
      const when = rehearsal ? ' (туршилт, бүртгэлд ороогүй)' : `, ${year}-01-01-ээс бүх систем үүн дээр ажиллана`;
      MailApp.sendEmail(
        Session.getEffectiveUser().getEmail(),
        `${ROLLOVER_MAIL_SUBJECT}: ${book.getName()}`,
        `"${book.getName()}" бэлэн боллоо${when}.\n${book.getUrl()}\nӨмнөх файл: ${old.getUrl()}`
      );
    }
    return { done: true, lines: ['эзэнд имэйл явууллаа'] };
  },
};

// The first sheet of each chain from the last one of the old year, then the copied tabs and flags.
function seedNewBook(old, book) {
  useBook(old);
  const sources = CHAINS.map((chain) => old.getSheetByName(chainSheetName(chain, CHAIN_PERIOD_COUNT - 1)));
  useBook(book);
  const lines = CHAINS.map((chain, at) => {
    const name = chainSheetName(chain, 0);
    if (sources[at] === null) {
      return `${chain.label}: өмнөх оны сүүлийн sheet алга`;
    }
    if (book.getSheetByName(name) === null && !DRY_RUN) {
      copyAsFirstPeriod(book, sources[at], name);
    }
    return `${chain.label}: ${sources[at].getName()} -> ${name}`;
  });
  COPIED_TABS.filter((tab) => old.getSheetByName(tab) !== null && book.getSheetByName(tab) === null).forEach((tab) => {
    if (!DRY_RUN) {
      old.getSheetByName(tab).copyTo(book).setName(tab);
    }
    lines.push(`"${tab}" хуулагдлаа`);
  });
  if (!DRY_RUN) {
    // The empty first tab of a new spreadsheet: "Sheet1" would be read as an old tab name (config.gs).
    book.getSheets().filter((sheet) => /^(Sheet1|Хуудас1)$/.test(sheet.getName())).forEach((sheet) => book.deleteSheet(sheet));
    const from = bookProperties(old);
    const to = bookProperties(book);
    LAYOUT_FLAGS.filter((flag) => from.getProperty(flag) === '1').forEach((flag) => to.setProperty(flag, '1'));
  }
  return lines;
}

// copyAsNextPeriod's rule (year-setup.gs) into another workbook: the same products, categories,
// colours and prices; income, expense, count and the opening balance emptied.
function copyAsFirstPeriod(book, source, name) {
  const copy = source.copyTo(book).setName(name);
  const bodyRows = copy.getMaxRows() - SETUP_FIRST_BODY_ROW + 1;
  if (bodyRows > 0) {
    copy.getRange(SETUP_FIRST_BODY_ROW, SETUP_ENTRY_FIRST_COL, bodyRows, SETUP_ENTRY_LAST_COL - SETUP_ENTRY_FIRST_COL + 1).clearContent();
    copy.getRange(SETUP_FIRST_BODY_ROW, SETUP_OPENING_COL, bodyRows, 1).clearContent();
  }
  return copy;
}

// The year's registry row gets this project's kind of file, used from 1 January of that year.
function registerBook(year, id) {
  const column = { main: 2, tosgon: 3 }[BOOK_KIND];
  const sheet = SpreadsheetApp.openById(REGISTRY_ID).getSheetByName(REGISTRY_TAB);
  const height = Math.max(sheet.getLastRow() - 1, 0);
  const years = height === 0 ? [] : sheet.getRange(2, 1, height, 1).getValues().map(([value]) => Number(value));
  const row = years.includes(year) ? years.indexOf(year) + 2 : sheet.getLastRow() + 1;
  if (!DRY_RUN) {
    sheet.getRange(row, 1).setValue(year);
    sheet.getRange(row, column).setValue(id);
    sheet.getRange(row, 4).setNumberFormat('@').setValue(`${year}-01-01`);
    CacheService.getScriptCache().remove(REGISTRY_CACHE_KEY);
    registryOfRun = null;
  }
  return `бүртгэл: ${year} оны ${BOOK_KIND} файл ${DRY_RUN ? 'бичигдэнэ' : 'бичигдлээ'}`;
}

// The last sheet of each chain in the old year's workbook and the first one in the new year's.
function chainEnds(old, book) {
  useBook(old);
  const last = CHAINS.map((chain) => old.getSheetByName(chainSheetName(chain, CHAIN_PERIOD_COUNT - 1)));
  useBook(book);
  return CHAINS.map((chain, at) => ({ chain, last: last[at], first: book.getSheetByName(chainSheetName(chain, 0)) }))
    .filter(({ last: from, first }) => from !== null && first !== null);
}

// Until 31 December: December's new products and changed prices reach the new year's first sheet.
function syncIntoNewYear(old, book) {
  return chainEnds(old, book).map(({ chain, last, first }) => {
    const present = new Set(readModel(first).products.map((product) => product.key));
    const items = readModel(last).products
      .filter((product) => !present.has(product.key))
      .map((product) => describeProduct(last, product));
    if (!DRY_RUN && items.length > 0) {
      applyChange(first, items, new Set());
    }
    return `${chain.label}: ${items.length} шинэ бараа, ${copyPrices(last, first)} үнэ -> ${first.getName()}`;
  });
}

// Cost, price and shelf price of the old year's last sheet where the new first sheet differs.
function copyPrices(from, to) {
  const byKey = new Map(readModel(from).products.map((product) => [product.key, product]));
  const fields = [
    ['cost', SYNC_COST_COL],
    ['price', SYNC_PRICE_COL],
    ...(SYNC_SHELF_PRICE_COL ? [['shelfPrice', SYNC_SHELF_PRICE_COL]] : []),
  ];
  let changed = 0;
  readModel(to).products.forEach((product) => {
    const source = byKey.get(product.key);
    fields
      .filter(([field]) => source && source[field] !== '' && source[field] !== product[field])
      .forEach(([field, column]) => {
        changed += 1;
        if (!DRY_RUN) {
          to.getRange(product.row, column).setValue(source[field]);
        }
      });
  });
  return changed;
}

// January: the old year's last count (or balance, CARRY_FROM_BALANCE) is the new year's opening.
function carryIntoNewYear(old, book) {
  return chainEnds(old, book).map(({ chain, last, first }) => {
    const result = carryInto(readModel(last).products, { sheet: first, name: first.getName(), ...readModel(first) });
    return `${chain.label}: анхны үлдэгдэл ${result.written} -> ${first.getName()}`;
  });
}

// A Тосгон workbook of another year holds the product list as values (A1 = LIST_COPY_TITLE): they
// follow the same year's main list every hour. The home Тосгон keeps its IMPORTRANGE copy.
function refreshListValues(book) {
  if (BOOK_KIND !== 'tosgon' || isHomeBook(book)) {
    return [];
  }
  const source = partnerBook(book, 'main', PRODUCT_LIST_SOURCE_ID).getSheetByName(PRODUCT_LIST_TAB);
  if (source === null) {
    return ['үндсэн файлын "Бараа" алга'];
  }
  const height = Math.max(source.getLastRow() - PRODUCT_LIST_FIRST_ROW + 1, 0);
  const values = height === 0 ? [] : source.getRange(PRODUCT_LIST_FIRST_ROW, 1, height, 2).getValues();
  if (!DRY_RUN) {
    const list = book.getSheetByName(PRODUCT_LIST_TAB) || book.insertSheet(PRODUCT_LIST_TAB, 0);
    list.getRange(1, 1).setValue(LIST_COPY_TITLE);
    const old = Math.max(list.getLastRow() - PRODUCT_LIST_FIRST_ROW + 1, 0);
    if (old > 0) {
      list.getRange(PRODUCT_LIST_FIRST_ROW, 1, old, 2).clearContent();
    }
    if (values.length > 0) {
      list.getRange(PRODUCT_LIST_FIRST_ROW, 1, values.length, 2).setValues(values);
    }
  }
  return [`"${PRODUCT_LIST_TAB}" жагсаалт ${values.length} мөр (утгаар)`];
}
