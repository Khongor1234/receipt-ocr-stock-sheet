/**
 * Receipt OCR - the month calendar of the price-watch site ("Баримт" tab): for every day, place
 * and folder, how many photos came in, how many lines were read, how many wait for mom, how
 * many are in the workbook and how many of mom's lines the OCR doubted (Шалгах).
 *
 * Every month spreadsheet the review pass looks at is counted again (ocrCalendarDay per day tab)
 * and kept in Script properties, one value per place, month and kind:
 *   ocr-cal:Соронзон|2026-10|зарлага = {"01":{"зарлага":[photos, lines, mom, written, unread, review, misfiled]}, ...}
 * The month spreadsheet and the tab of each day are kept next to it, for the site's links:
 *   ocr-cal-link:Соронзон|2026-10|зарлага = {"id":"<spreadsheet id>","days":{"01":123456}}
 * A web app (doGet) answers that month as JSON, each place's kinds put together (their folders differ),
 * with each place's comparison of mom's hand-typed cells with the receipts (checks, read from the
 * month spreadsheets: ocr-reconcile.gs); it is called by the site's server only, with a secret token
 * (OCR_FEED_TOKEN_PROPERTY, made by installReceiptFeed). A month's answer is kept in the script cache
 * (ocrFeedCachedAnswer) until the next review pass, so a visit does not open every month spreadsheet:
 * that took seconds, and the site gave up waiting now and then ("feed timeout").
 */

const OCR_CALENDAR_PREFIX = 'ocr-cal:';
const OCR_CALENDAR_VERSION = 6; // a change makes the next run count every recent month spreadsheet again
const OCR_FEED_TOKEN_PROPERTY = 'OCR_FEED_TOKEN';
const OCR_FEED_MODE = 'receipts';
const OCR_MONTH_KEY = /^\d{4}-\d{2}$/;
const OCR_FEED_CACHE_PREFIX = 'ocr-feed:';
const OCR_CALENDAR_EDIT_PROPERTY = 'ocr-cal-edit'; // when mom's mark last changed a day's numbers (ocrKeepCalendarDay): part of the cache key
const OCR_FEED_CACHE_SECONDS = 10 * 60; // the Тулгалт tabs change outside the review pass too (mom's edits): read again this often at most

// One kind's key; a place's month is every key starting with ocrCalendarKey(place, month) + '|'.
const ocrCalendarKey = (place, month, kind) => `${OCR_CALENDAR_PREFIX}${place}|${month}${kind ? `|${kind.word}` : ''}`;
const OCR_CALENDAR_LINK_PREFIX = 'ocr-cal-link:';
const ocrCalendarLinkKey = (place, month, kind) => `${OCR_CALENDAR_LINK_PREFIX}${place}|${month}${kind ? `|${kind.word}` : ''}`;
const ocrPhotoOf = (record) => ocrText(record['ID']).split('#')[0] || ocrText(record['Баримт']);

// One day tab's numbers per folder: { "зарлага": [photos, lines, mom, written, unread, review, misfiled] },
// where lines are product lines, mom the lines she has not looked at (a Хамааралгүй line counts as
// looked at), written the lines in the workbook, unread the photos without a single product line
// (a failed reading, a duplicate), review those of mom's lines the OCR marked Шалгах. The site
// colours a day blue while review > 0, yellow while only mom > 0, green when both are 0. misfiled: the
// photos noted as maybe in the wrong folder (ocrMisfiledReason); the site marks them ⚠ (2026-10-23).
function ocrCalendarDay(records) {
  const folders = new Map();
  records.forEach((record) => {
    const folder = ocrText(record['Хавтас (ээж хийсэн)']).split(' / ')[1] || '?';
    folders.set(folder, [...(folders.get(folder) || []), record]);
  });
  return Object.fromEntries(
    [...folders].map(([folder, rows]) => {
      const products = rows.filter(ocrIsProductLine);
      const forMom = products.filter((record) => ocrReviewState(record).needsMom);
      const photos = new Set(rows.map(ocrPhotoOf));
      const read = new Set(products.map(ocrPhotoOf));
      return [folder, [
        photos.size,
        products.length,
        forMom.length,
        products.filter((record) => ocrWrittenState(record) === 'done').length,
        [...photos].filter((photo) => !read.has(photo)).length,
        forMom.filter((record) => ocrText(record['Төлөв']) === OCR_STATUS_REVIEW).length,
        new Set(rows.filter((record) => ocrText(record['Шалгах шалтгаан']).includes(OCR_MISFILED_REASON)).map(ocrPhotoOf)).size,
      ]];
    })
  );
}

// results: [{ day: "01", gid, calendar: ocrCalendarDay(...) }] of one month spreadsheet; a day
// without rows is left out of both the numbers and the links.
function ocrKeepCalendar(props, entry, results) {
  const shown = results
    .filter((result) => Object.keys(result.calendar).length > 0)
    .sort((a, b) => a.day.localeCompare(b.day));
  const kept = Object.fromEntries(shown.map((result) => [result.day, result.calendar]));
  const links = { id: entry.file.getId(), days: Object.fromEntries(shown.map((result) => [result.day, result.gid])) };
  if (DRY_RUN) {
    Logger.log(`Календарь ${entry.place} ${entry.kind.word} ${entry.month}: ${Object.keys(kept).length} өдөр (туршилт, хадгалаагүй)`);
    return;
  }
  props.setProperty(ocrCalendarKey(entry.place, entry.month, entry.kind), JSON.stringify(kept));
  props.setProperty(ocrCalendarLinkKey(entry.place, entry.month, entry.kind), JSON.stringify(links));
  props.setProperty(OCR_CALENDAR_EDIT_PROPERTY, new Date().toISOString()); // a spreadsheet reviewed from the editor (reviewOcrBookNow) is no pass: the site reads it at once all the same
}

// The values of a place's month, one per kind: [[word, value]] (all: the Script properties).
const ocrKindValues = (all, base) => Object.keys(all).filter((key) => key.startsWith(`${base}|`)).sort().map((key) => [key.slice(base.length + 1), JSON.parse(all[key])]);

// A place's month: the days of its kinds put together ({ "01": { folder: numbers } }).
const ocrPlaceCalendar = (all, place, month) => ocrKindValues(all, ocrCalendarKey(place, month)).reduce(
  (days, [, kept]) => Object.entries(kept).reduce((merged, [day, folders]) => ({ ...merged, [day]: { ...(merged[day] || {}), ...folders } }), days),
  {}
);

// checks of several spreadsheets ({ "01": [check] }) put together.
const ocrMergeChecks = (lists) => lists.reduce(
  (days, checks) => Object.entries(checks).reduce((merged, [day, list]) => ({ ...merged, [day]: [...(merged[day] || []), ...list] }), days),
  {}
);

// The answer of the web app. params: the query (token, mode, month 'YYYY-MM'; this month if none).
function ocrFeedAnswer(params, props, now) {
  const token = props.getProperty(OCR_FEED_TOKEN_PROPERTY);
  if (!token || ocrText(params.token) !== token) {
    return { ok: false, error: 'forbidden' };
  }
  if (params.mode !== OCR_FEED_MODE) {
    return { ok: false, error: 'unknown mode' };
  }
  const month = ocrFeedMonth(params, now);
  const all = props.getProperties();
  const months = [...new Set(Object.keys(all).filter((key) => key.startsWith(OCR_CALENDAR_PREFIX)).map((key) => key.split('|')[1]))].sort();
  const places = Object.fromEntries(OCR_PLACES.map((place) => [place, ocrPlaceCalendar(all, place, month)]));
  // each place's spreadsheets by kind ({ "орлого": { id, days } }), null when it has none that month
  const links = Object.fromEntries(OCR_PLACES.map((place) => {
    const kinds = ocrKindValues(all, ocrCalendarLinkKey(place, month));
    return [place, kinds.length > 0 ? Object.fromEntries(kinds) : null];
  }));
  const checks = Object.fromEntries(OCR_PLACES.map((place) => [place, ocrMergeChecks(Object.values(links[place] || {}).map((link) => ocrReadChecks(link.id)))]));
  return { ok: true, month, months, places, links, checks, updated: all[OCR_LAST_SCAN_PROPERTY] || '' };
}

// The month asked for ('YYYY-MM'), this month if none.
const ocrFeedMonth = (params, now) => (OCR_MONTH_KEY.test(ocrText(params.month)) ? ocrText(params.month) : Utilities.formatDate(now, OCR_TIME_ZONE, 'yyyy-MM'));

// ocrFeedAnswer, with a month's answer kept in the script cache: the key carries the last review
// pass's stamp, so a pass makes the next visit read again; a refusal is answered at once and never kept.
function ocrFeedCachedAnswer(params, props, cache, now) {
  const token = props.getProperty(OCR_FEED_TOKEN_PROPERTY);
  if (!token || ocrText(params.token) !== token || params.mode !== OCR_FEED_MODE) {
    return ocrFeedAnswer(params, props, now);
  }
  const key = `${OCR_FEED_CACHE_PREFIX}${ocrFeedMonth(params, now)}|${props.getProperty(OCR_LAST_SCAN_PROPERTY) || ''}|${props.getProperty(OCR_CALENDAR_EDIT_PROPERTY) || ''}`;
  const kept = cache.get(key);
  if (kept) {
    return JSON.parse(kept);
  }
  const answer = ocrFeedAnswer(params, props, now);
  if (answer.ok) {
    try {
      cache.put(key, JSON.stringify(answer), OCR_FEED_CACHE_SECONDS);
    } catch (error) {
      Logger.log(`Сайтын хариу cache-д багтсангүй (${key}): ${error.message}`); // over 100 KB: answered without the cache
    }
  }
  return answer;
}

function doGet(event) {
  const answer = ocrFeedCachedAnswer((event && event.parameter) || {}, PropertiesService.getScriptProperties(), CacheService.getScriptCache(), new Date());
  return ContentService.createTextOutput(JSON.stringify(answer)).setMimeType(ContentService.MimeType.JSON);
}

// Run once: makes the site's token (kept if there is one) and says how to publish the web app.
function installReceiptFeed() {
  const props = PropertiesService.getScriptProperties();
  const existing = props.getProperty(OCR_FEED_TOKEN_PROPERTY);
  if (DRY_RUN) {
    Logger.log(`${modeLabel()}: сайтын token ${existing ? 'аль хэдийн байгаа' : 'шинээр үүснэ'}.`);
    return;
  }
  const token = existing || (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  props.setProperty(OCR_FEED_TOKEN_PROPERTY, token);
  Logger.log(`Сайтын token (RECEIPTS_FEED_TOKEN): ${token}`);
  Logger.log('Дараа нь: Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone), /exec хаягийг RECEIPTS_FEED_URL болгоно.');
}
