/**
 * One set of files per year, and which one the scripts work on.
 *
 * The owner's rule (2026-10-21): every year has its own workbooks ("бараа орлого зарлага 2027",
 * "Тосгон 2027"), made from December's products and prices (year-rollover.gs), and every system
 * moves to them on 1 January by itself. The scripts stay in the projects of the 2026 workbooks:
 * the files of the other years hold no script, and these projects open them by id.
 *
 * The registry, a spreadsheet of its own (REGISTRY_ID, tab REGISTRY_TAB), has one row per year:
 *   A year, B the main workbook's id, C Тосгон's id, D the first day its files are used (YYYY-MM-DD).
 * The file of a date is the one of the row with the latest first day on or before it. Without a
 * registry (or a row for the date) the scripts keep to the files they always used: the workbook
 * the project is bound to, and the ids written in the code.
 *
 * Each run works on one workbook at a time (useBook): the year of that workbook decides the tab
 * names (CHAIN_FIRST_MONTH, CHAIN_MONTHS in profile.gs), and its state is kept apart from the
 * other years' (bookProperties: the keys of another year's file start with its id).
 *   - an edit: the workbook edited (event.source);
 *   - anything else: the workbook of today (targetBook), and in January the hourly job also
 *     looks after last year's (booksToMaintain), where December's counts arrive late.
 * Needs profile.gs (BOOK_KIND, HOME_CHAIN), config.gs.
 */

const REGISTRY_ID = 'YOUR_YEAR_REGISTRY_SPREADSHEET_ID'; // "Он бүрийн файл (бүртгэл)"; '' = no registry
const REGISTRY_TAB = 'Файлууд';
const REGISTRY_CACHE_KEY = 'yearBooks';
const REGISTRY_CACHE_SECONDS = 10 * 60;
const MAINTAIN_LAST_YEAR_UNTIL_DAY = 31; // January: last year's workbook is still looked after
const YEAR_IN_NAME = /(\d{4})\s*(?:\(|$)/; // "бараа орлого зарлага 2027", "Тосгон 2027 (туршилт)"

let ACTIVE_BOOK = null; // the workbook this run works on (useBook)

// ---------- pure helpers (no Spreadsheet calls) ----------

// Registry rows ([year, mainId, tosgonId, firstDay]) -> [{ year, main, tosgon, from }], oldest first.
function parseRegistry(rows) {
  return rows
    .map(([year, main, tosgon, from]) => ({
      year: Number(year),
      main: String(main || '').trim(),
      tosgon: String(tosgon || '').trim(),
      from: from instanceof Date ? isoDay(from) : String(from || '').trim(),
    }))
    .filter((entry) => Number.isInteger(entry.year) && entry.year > 2000 && /^\d{4}-\d{2}-\d{2}$/.test(entry.from))
    .sort((a, b) => (a.from < b.from ? -1 : 1));
}

function isoDay(date) {
  const pad = (n) => String(n).padStart(2, '0');
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}

// The registry entry of a day ("YYYY-MM-DD"): the latest one that has started. null: none.
function entryForDay(entries, day) {
  const started = entries.filter((entry) => entry.from <= day);
  return started.length === 0 ? null : started[started.length - 1];
}

// The months a workbook of `year` holds: the home workbook keeps its own span (Тосгон 2026 starts
// in October 2025); every later year is January to December.
function chainSpanOf(year) {
  return year === HOME_CHAIN.year ? HOME_CHAIN : { year, first: { year, month: 1 }, months: MONTHS_PER_YEAR };
}

// The year a workbook is for: its registry row, else the year in its name, else the home year.
function bookYearOf(entries, id, name) {
  const entry = entries.find((candidate) => candidate.main === id || candidate.tosgon === id);
  if (entry) {
    return entry.year;
  }
  const match = String(name || '').match(YEAR_IN_NAME);
  return match ? Number(match[1]) : HOME_CHAIN.year;
}

// ---------- the registry ----------

// The registry as this run first read it (registerBook forgets it): one receipt review asked for it
// 467 times, each a CacheService read, 13 s of its 47 (2026-10-29).
let registryOfRun = null;

function readRegistry() {
  if (registryOfRun === null) {
    registryOfRun = readRegistryNow();
  }
  return registryOfRun;
}

function readRegistryNow() {
  if (REGISTRY_ID === '') {
    return [];
  }
  const cache = CacheService.getScriptCache();
  const cached = cache.get(REGISTRY_CACHE_KEY);
  if (cached) {
    return JSON.parse(cached);
  }
  let sheet = null;
  try {
    const book = SpreadsheetApp.openById(REGISTRY_ID);
    sheet = book ? book.getSheetByName(REGISTRY_TAB) : null;
  } catch (error) {
    Logger.log(`Он бүрийн файлын бүртгэл уншигдсангүй (${error.message}): энэ удаа ердийн файлууд дээр ажиллана`);
    return []; // not cached: the next run tries again
  }
  const height = sheet === null ? 0 : sheet.getLastRow() - 1;
  const entries = height <= 0 ? [] : parseRegistry(sheet.getRange(2, 1, height, 4).getValues());
  cache.put(REGISTRY_CACHE_KEY, JSON.stringify(entries), REGISTRY_CACHE_SECONDS);
  return entries;
}

// The id of a kind ("main" | "tosgon") of workbook for a date; null: the registry has none.
function bookIdForDate(date, kind = BOOK_KIND) {
  const entry = entryForDay(readRegistry(), isoDay(date));
  return entry && entry[kind] ? entry[kind] : null;
}

// The workbook of a kind for a date: from the registry, else the one the project is bound to
// (the kind this project serves) or the id written in the code (fallbackId, the other kind).
function bookForDate(date, kind = BOOK_KIND, fallbackId = null) {
  const id = bookIdForDate(date, kind);
  if (id !== null) {
    return SpreadsheetApp.openById(id);
  }
  return fallbackId === null ? SpreadsheetApp.getActiveSpreadsheet() : SpreadsheetApp.openById(fallbackId);
}

// The other kind's workbook of the same year as `book` (the main workbook's Тосгон, and back);
// fallbackId: the one written in the code, for a year the registry does not have.
function partnerBook(book, kind, fallbackId) {
  const entries = readRegistry();
  const year = bookYearOf(entries, book.getId(), book.getName());
  const entry = entries.find((candidate) => candidate.year === year);
  return SpreadsheetApp.openById(entry && entry[kind] ? entry[kind] : fallbackId);
}

// ---------- the workbook of this run ----------

function useChainYear(year) {
  const span = chainSpanOf(year);
  CHAIN_FIRST_MONTH = span.first;
  CHAIN_MONTHS = span.months;
  CHAIN_PERIOD_COUNT = CHAIN_MONTHS * PERIODS_PER_MONTH;
}

// Works on `book` from now on: its year decides the tab names.
function useBook(book) {
  useChainYear(bookYearOf(readRegistry(), book.getId(), book.getName()));
  ACTIVE_BOOK = book;
  return book;
}

// fn() with the tab names of `year`'s workbook (a receipt of that year), the run's own after it.
function withChainYear(year, fn) {
  const saved = [CHAIN_FIRST_MONTH, CHAIN_MONTHS, CHAIN_PERIOD_COUNT];
  useChainYear(year);
  try {
    return fn();
  } finally {
    [CHAIN_FIRST_MONTH, CHAIN_MONTHS, CHAIN_PERIOD_COUNT] = saved;
  }
}

// The registry's id of a kind of workbook for a year; null: the registry has none.
function bookIdOfYear(year, kind) {
  const entry = readRegistry().find((candidate) => candidate.year === year);
  return entry && entry[kind] ? entry[kind] : null;
}

// The workbook an entry point works on: the edited one, else today's.
function targetBook(event) {
  return useBook(event && event.source ? event.source : bookForDate(new Date()));
}

function currentBook() {
  return ACTIVE_BOOK || targetBook();
}

// The workbooks the hourly job looks after on `date`: today's, and in January last year's too.
function booksToMaintain(date) {
  const today = bookForDate(date);
  if (date.getMonth() !== 0 || date.getDate() > MAINTAIN_LAST_YEAR_UNTIL_DAY) {
    return [today];
  }
  const lastId = bookIdForDate(new Date(date.getFullYear() - 1, 11, 31));
  return lastId === null || lastId === today.getId() ? [today] : [today, SpreadsheetApp.openById(lastId)];
}

// The project's triggers that listen to `book` (edit, change), not the time-based ones.
function triggersOf(book) {
  return ScriptApp.getProjectTriggers().filter((trigger) => trigger.getTriggerSourceId() === book.getId());
}

// ---------- state kept per workbook ----------

const isHomeBook = (book) => {
  const container = SpreadsheetApp.getActiveSpreadsheet();
  return container === null || book.getId() === container.getId();
};

// A property store whose keys belong to `book`: the home workbook's keys as they always were, the
// other years' with the file's id in front, so a sheet id or a flag of one file never counts for
// another. store: PropertiesService.getDocumentProperties() or getScriptProperties().
function storeOfBook(store, book) {
  const key = (name) => (isHomeBook(book) ? name : `${book.getId()}:${name}`);
  return {
    getProperty: (name) => store.getProperty(key(name)),
    setProperty: (name, value) => store.setProperty(key(name), value),
    deleteProperty: (name) => store.deleteProperty(key(name)),
  };
}

const bookProperties = (book = currentBook()) => storeOfBook(PropertiesService.getDocumentProperties(), book);
const bookScriptProperties = (book = currentBook()) => storeOfBook(PropertiesService.getScriptProperties(), book);
