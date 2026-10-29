/**
 * Setting up a month spreadsheet (protection, name list, view trigger, formats, rules, Тосгон's
 * costs), a tab at a time until the deadline. A 30-day spreadsheet's setup takes minutes: on
 * 2026-10-28 it ran past Google's 6 minutes and the whole run was killed. A setup cut by the time
 * keeps the tabs it did (OCR_BOOK_SETUP_PROGRESS_PREFIX), and the next run does only the rest.
 */

// step(item) for each item while there is time: true when every item was done.
function ocrEachInTime(items, step, deadline) {
  for (const item of items) {
    if (Date.now() > deadline) {
      return false;
    }
    step(item);
  }
  return true;
}

// A month spreadsheet's protection, name list, view trigger, formats and rules (and Тосгон's costs),
// made again only when what they come from changed (mom's address, the layout, the names, a new
// day tab, Тосгон's costs of the month). Made on every review until 2026-10-23, they took about a
// minute a spreadsheet and changed the file, so every spreadsheet looked at came up again in the
// next pass. true: the spreadsheet is set up; false: the time ran out, the next run goes on.
function ocrSetUpBook(book, entry, context, deadline = Infinity) {
  const { props } = context;
  const key = OCR_BOOK_SETUP_PREFIX + entry.file.getId();
  const progressKey = OCR_BOOK_SETUP_PROGRESS_PREFIX + entry.file.getId();
  const priced = ocrPricedBookOf(book);
  const costs = priced && priced.place === OCR_TOSGON.place ? ocrTosgonCosts(entry.month) : null;
  const tabs = ocrDayTabs(book);
  const wanted = ocrTextHash(JSON.stringify([OCR_BOOK_SETUP_VERSION, context.mom, OCR_HEADER, context.catalog.entries, tabs.map((tab) => tab.getName()), costs]));
  if (props.getProperty(key) === wanted) {
    return true;
  }
  const saved = ocrSetUpProgress(props.getProperty(progressKey), wanted);
  if (saved.tabs.length === 0) {
    ocrSetUpBookWide(book, entry, context, costs);
  }
  const description = ocrGuardDescription(context.mom);
  const owner = Session.getEffectiveUser().getEmail();
  let progress = saved;
  const isDone = ocrEachInTime(tabs.filter((tab) => !saved.tabs.includes(tab.getName())), (tab) => {
    const isGuarded = ocrGuardTab(tab, description, context.mom, owner);
    ocrSetUpTab(tab);
    progress = { ...progress, tabs: [...progress.tabs, tab.getName()], isGuarded: progress.isGuarded || isGuarded };
  }, deadline);
  if (!isDone) {
    props.setProperty(progressKey, JSON.stringify(progress));
    return false;
  }
  if (progress.isGuarded) {
    ocrShareWith(DriveApp.getFileById(book.getId()), ocrMomList(context.mom), owner);
  }
  props.deleteProperty(progressKey);
  props.setProperty(key, wanted);
  return true;
}

// What a cut setup of the same wanted state already did: { wanted, tabs, isGuarded }.
function ocrSetUpProgress(text, wanted) {
  const fresh = { wanted, tabs: [], isGuarded: false };
  try {
    const saved = text ? JSON.parse(text) : null;
    return saved && saved.wanted === wanted && Array.isArray(saved.tabs) ? { ...fresh, tabs: saved.tabs, isGuarded: Boolean(saved.isGuarded) } : fresh;
  } catch (error) {
    Logger.log(`OCR тохиргооны явц уншигдсангүй (${error.message}): эхнээс нь хийнэ`);
    return fresh;
  }
}

// What a month spreadsheet gets once, not per tab.
function ocrSetUpBookWide(book, entry, context, costs) {
  if (costs) {
    ocrSetCosts(book, costs);
  }
  const mainId = ocrCatalogBookId(Number(entry.month.slice(0, 4)));
  ocrAllowImport(book.getId(), mainId, context.props);
  ocrSetNameList(book, context.catalog.entries, mainId);
  ocrEnsureViewTrigger(book);
}

// One day tab's formats, rules, photos, lists, widths, legend and prices.
function ocrSetUpTab(tab) {
  ocrFormatColumns(tab);
  tab.setConditionalFormatRules(ocrRowRules(tab)); // a tab made before a rule was added gets it too
  ocrPlacePhotos(tab); // the photos in N (OCR_BOOK_SETUP_VERSION 6), in M from 11 ...
  ocrSpreadPhotos(tab); // ... over rows of the default height, room left below them (7) ...
  ocrClearMomColor(tab); // ... no purple, the lists on the lines only (8) ...
  ocrSetLineChoices(tab);
  ocrSetViewChoices(tab); // ... "Бүх мөр Зөв" in A1's list (12) ...
  ocrShortenTabReasons(tab, ocrReadRecords(tab)); // ... the reasons in a few words (13) ...
  ocrFitColumns(tab); // ... "Бараа" as wide as its names, "Шалгах шалтгаан" as its reasons (the owner, 2026-10-28; 11: the reason after the photo, in mom's view) ...
  ocrWriteLegend(tab); // ... and the legend of the colours beside them (in Q from 9) ...
  ocrWritePriceColumns(tab); // ... after the receipt price and ours (or Тосгон's cost) in O and P (9)
}
