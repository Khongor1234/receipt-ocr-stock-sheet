/**
 * Receipt OCR - what mom sees of the checks: "Баримт/Тайлан/2026-10/Тайлан 2026-10-15", a Google
 * Doc in Mongolian, only on the last day of a half-month (the 15th and the month's last day, like
 * the workbook's sheets): what was read and written (or corrected) in that half-month, what waits
 * for her, what could not be written, and its counts (the count report, ocr-count-report.gs, shows
 * a chosen day in full, on any day). On those days it is written again on every run, so it ends
 * with the day's last run. What is left to do per day spreadsheet is kept in Script properties
 * (OCR_OPEN_PREFIX, dropped once nothing is left); what happened on a day too (OCR_DAY_PREFIX, kept
 * OCR_DAY_KEEP_DAYS days).
 */

// counts: the counts read that day, { place, day } (the day of the count, not of the reading)
const OCR_EMPTY_STATS = { photos: 0, rows: 0, ok: 0, written: 0, corrected: 0, failed: 0, counts: [] };
const ocrCountKey = (count) => `${count.place}|${count.day}`;

// ---------- kept state ----------

// DRY_RUN keeps this run's summaries here instead of in Script properties, so its report shows
// the checks as they stand now (and the kept ones of the sheets it did not look at).
const ocrDrySummaries = new Map();

function ocrKeepSummary(props, summary) {
  if (DRY_RUN) {
    ocrDrySummaries.set(summary.id, summary);
    return;
  }
  if (ocrSummaryIsOpen(summary)) {
    props.setProperty(OCR_OPEN_PREFIX + summary.id, JSON.stringify(summary));
  } else {
    props.deleteProperty(OCR_OPEN_PREFIX + summary.id);
  }
}

function ocrOpenSummaries(props) {
  const kept = Object.entries(props.getProperties())
    .filter(([key]) => key.startsWith(OCR_OPEN_PREFIX))
    .map(([, value]) => JSON.parse(value))
    .filter((summary) => !ocrDrySummaries.has(summary.id));
  return [...kept, ...[...ocrDrySummaries.values()].filter(ocrSummaryIsOpen)];
}

const ocrDayStats = (props, day) => ({ ...OCR_EMPTY_STATS, ...JSON.parse(props.getProperty(OCR_DAY_PREFIX + day) || '{}') });

// Adds to today's numbers (delta: some of photos, rows, ok, written, corrected, failed, counts).
function ocrAddDayStats(context, delta) {
  if (DRY_RUN) {
    return;
  }
  const stats = ocrDayStats(context.props, context.today);
  const counts = [...stats.counts, ...(delta.counts || [])];
  const merged = Object.fromEntries(
    Object.keys(OCR_EMPTY_STATS).map((key) =>
      key === 'counts'
        ? [key, counts.filter((count, index) => counts.findIndex((other) => ocrCountKey(other) === ocrCountKey(count)) === index)]
        : [key, stats[key] + (delta[key] || 0)]
    )
  );
  context.props.setProperty(OCR_DAY_PREFIX + context.today, JSON.stringify(merged));
}

const ocrUniqueCounts = (counts) => counts.filter((count, index) => counts.findIndex((other) => ocrCountKey(other) === ocrCountKey(count)) === index);

// The kept days from `from` to `to` ('YYYY-MM-DD', both included), as Script property keys.
const ocrDayKeys = (props, from, to = '9999-12-31') => Object.keys(props.getProperties())
  .filter((key) => key.startsWith(OCR_DAY_PREFIX))
  .filter((key) => key.slice(OCR_DAY_PREFIX.length) >= from && key.slice(OCR_DAY_PREFIX.length) <= to);

// The numbers of the days from `from` to `to` added up, each count once.
function ocrPeriodStats(props, from, to) {
  const days = ocrDayKeys(props, from, to).sort().map((key) => ocrDayStats(props, key.slice(OCR_DAY_PREFIX.length)));
  return Object.fromEntries(
    Object.keys(OCR_EMPTY_STATS).map((key) =>
      key === 'counts'
        ? [key, ocrUniqueCounts(days.flatMap((day) => day.counts))]
        : [key, days.reduce((total, day) => total + day[key], 0)]
    )
  );
}

// The count sheets read from `from` on, each once, the newest first.
function ocrKeptCounts(props, from = '') {
  const counts = ocrDayKeys(props, from)
    .sort()
    .reverse()
    .flatMap((key) => JSON.parse(props.getProperty(key)).counts || []);
  return ocrUniqueCounts(counts);
}

function ocrDropOldDays(props, today) {
  const oldest = new Date(new Date(`${today}T00:00:00Z`).getTime() - OCR_DAY_KEEP_DAYS * OCR_DAY_MS).toISOString().slice(0, 10);
  Object.keys(props.getProperties())
    .filter((key) => key.startsWith(OCR_DAY_PREFIX) && key.slice(OCR_DAY_PREFIX.length) < oldest)
    .forEach((key) => props.deleteProperty(key));
}

// ---------- pure: what is shown ----------

const ocrSummaryDay = (summary) => summary.name.slice(OCR_DAY_SPREADSHEET_PREFIX.length);
const ocrByDay = (summaries) => summaries.slice().sort((a, b) => (ocrSummaryDay(a) + a.label).localeCompare(ocrSummaryDay(b) + b.label));
const ocrFolderOf = (summary) => summary.label.split(' / ').slice(0, 2).join(' / ');

// The report comes out on the last day of a half-month: the 15th and the month's last day.
function ocrIsReportDay(iso) {
  const [year, month, day] = iso.split('-').map(Number);
  return day === FIRST_HALF_LAST_DAY || day === daysInMonth(year, month);
}

// The first day of the half-month of 'YYYY-MM-DD': the 1st or the 16th.
function ocrHalfStart(iso) {
  const isSecondHalf = Number(iso.slice(8)) > FIRST_HALF_LAST_DAY;
  return `${iso.slice(0, 8)}${isSecondHalf ? FIRST_HALF_LAST_DAY + 1 : '01'}`;
}

// Mom's report of the half-month from..today as lines: { style: 'title' | 'heading' | 'text' |
// 'item', text, url? }. countReports: [{ text, url }], the counts of the half-month (url: the count report).
function ocrReportLines(from, today, time, stats, summaries, countReports) {
  const title = (text) => ({ style: 'title', text });
  const heading = (text) => ({ style: 'heading', text });
  const text = (value) => ({ style: 'text', text: value });
  const item = (value, url) => ({ style: 'item', text: value, url });
  const sheetLine = (summary, count, what) => item([ocrSummaryDay(summary), ocrFolderOf(summary), `${count} мөр ${what}`].filter(Boolean).join(' · '), summary.url);
  const sorted = ocrByDay(summaries);
  const sum = (key) => sorted.reduce((total, summary) => total + summary[key], 0);
  const section = (name, lines, empty) => [heading(name), ...(lines.length > 0 ? lines : [text(empty)])];

  const mom = sorted.filter((summary) => summary.mom > 0).map((summary) => sheetLine(summary, summary.mom, 'шалгах'));
  const failed = sorted.flatMap((summary) => summary.failed.map((line) => item(`${line.product || 'нэргүй мөр'}: ${line.reason}`, summary.url)));
  const counts = countReports.map((report) => item(report.text, report.url));
  const differences = sorted.flatMap((summary) => (summary.check || []).map((line) => item(`${ocrSummaryDay(summary)} · ${line.product}: ${line.text}`, summary.url)));

  return [
    title(`Баримтын тайлан ${from} – ${today} (${time}-ийн байдлаар)`),
    heading('Энэ хагас сар'),
    text(`Уншсан зураг ${stats.photos}, мөр ${stats.rows} (шууд OK ${stats.ok}).`),
    text(`Workbook-д бичсэн мөр ${stats.written}, засвар орсон ${stats.corrected}, бичиж чадаагүй ${stats.failed}.`),
    ...section(`Ээжийн шалгах мөр (нийт ${sum('mom')})`, mom, 'Шалгах мөр алга.'),
    ...section('Бичиж чадаагүй мөр', failed, 'Алга.'),
    ...section('Гараар бичсэн тоотой зөрсөн баримт (check файлд бичсэн; шалгаад Зөв эсвэл Буруу гэж тэмдэглэнэ)', differences, 'Зөрсөн мөр алга.'),
    ...section('Энэ хагас сарын тооллогууд (Тооллогын тайлан дээр өдрөө сонгож дэлгэрэнгүйг харна)', counts, 'Энэ хагас сард тооллого ороогүй.'),
  ];
}

// ---------- written to Drive ----------

// A Google Doc of lines in Баримт/Тайлан/<month>, written anew; its link ('' in DRY_RUN).
function ocrWriteDoc(root, today, name, lines) {
  if (DRY_RUN) {
    Logger.log(`${name} (туршилт, бичээгүй):\n${lines.map((line) => line.text).join('\n')}`);
    return '';
  }
  const folder = ocrEnsurePath(root, [OCR_REPORT_FOLDER, ocrMonthFolderName(today)]);
  const files = folder.getFilesByName(name);
  let doc;
  if (files.hasNext()) {
    doc = DocumentApp.openById(files.next().getId());
  } else {
    doc = DocumentApp.create(name);
    DriveApp.getFileById(doc.getId()).moveTo(folder);
  }
  const body = doc.getBody();
  body.clear();
  const headings = { title: DocumentApp.ParagraphHeading.TITLE, heading: DocumentApp.ParagraphHeading.HEADING2 };
  lines.forEach((line) => {
    const paragraph = line.style === 'item' ? body.appendListItem(line.text) : body.appendParagraph(line.text);
    if (headings[line.style]) {
      paragraph.setHeading(headings[line.style]);
    }
    if (line.url) {
      paragraph.setLinkUrl(line.url);
    }
  });
  doc.saveAndClose();
  return DriveApp.getFileById(doc.getId()).getUrl();
}

// The count report on every run (mom chooses a day in it); mom's report, from what is kept, only
// on the last day of a half-month.
function writeOcrOutputs(context) {
  const { props, root, today, readAt } = context;
  // the count report keeps the day and place chosen in it
  const report = refreshCountReport(root, props, today);
  if (!DRY_RUN) {
    ocrDropOldDays(props, today);
  }
  if (!ocrIsReportDay(today)) {
    Logger.log(`Тайлан: 15-нд болон сарын сүүлийн өдөр гарна, өнөөдөр (${today}) гаргасангүй.`);
    return;
  }
  const from = ocrHalfStart(today);
  const summaries = ocrOpenSummaries(props);
  const stats = ocrPeriodStats(props, from, today);
  // the counts of the half-month with their results as they stand now (a late receipt of that day changes them)
  const countReports = ocrKeptCounts(props, from).map((count) => {
    const read = ocrReadCountAt(root, count.place, count.day);
    const table = read ? ocrCountTable({ ...read, ...count }, ocrWaitingByProduct(summaries, count.place), new Map()) : null;
    return { text: `${count.place} ${count.day}: ${table ? table.conclusion.replace(/^Дүгнэлт: /, '') : 'workbook-д sheet алга'}`, url: report.url };
  });
  ocrWriteDoc(root, today, OCR_REPORT_PREFIX + today, ocrReportLines(from, today, readAt.slice(11), stats, summaries, countReports));
}
