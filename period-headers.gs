/**
 * Fills the period information of every half-month sheet from its tab name
 * (sheetSlot in config.gs): My/2026/1/15 = January 1-15, My/2026/1/31 = January
 * 16-31, My/2026/2/15 = February 1-15, ... the same for Sor/....
 *
 * Per sheet it writes
 *   A1             the list that shows or hides columns (view-columns.gs)
 *   B1 / C1        the seller's name is typed in B1 and shown per month on
 *                  Sheet200 (charts.gs); C1 holds its label. A name typed
 *                  in B1 is never touched.
 *   E1:S1, U1:AI1  the day numbers (1-15 or 16-30); days beyond the end of the
 *                  month (29/30 of February) are left empty. Day 31 has no column
 *                  of its own: it is entered in the "30" column, as before.
 *   T1 / AJ1       not written (transfers between the two stores, not a day)
 *   A2             title, e.g. "My home 2026 2 сарын 16-28"
 * Nothing else is touched. Rows 1-2 are protected, so run this as the file owner
 * (by hand, or from an installable trigger), not from Mom's account.
 *
 * Usage: run fillPeriodHeaders (DRY_RUN in config.gs decides whether it only
 * reports: start and end date, old -> new title per sheet). setupYear runs it for
 * the whole year. Needs config.gs (findChains, sheetSlot).
 */

const PERIOD_DAY_COLUMNS = DAY_COLUMNS;
const PERIOD_FIRST_DAYS_START_COL = LAYOUT.incomeFirst; // E (E:S)
const PERIOD_SECOND_DAYS_START_COL = LAYOUT.expenseFirst; // U (U:AI)
const PERIOD_SELLER_LABEL = '← Худалдагчийн нэр';
const PERIOD_SELLER_LABEL_COL = 3; // C1
const PERIOD_SELLER_NAME_COL = 2; // B1: typed by hand

function fillPeriodHeaders() {
  const lines = runFillPeriodHeaders(currentBook());
  Logger.log('%s\n%s', modeLabel(), lines.join('\n'));
}

function runFillPeriodHeaders(spreadsheet) {
  const timeZone = spreadsheet.getSpreadsheetTimeZone();
  return findChains(spreadsheet).flatMap(({ label, sheets }) =>
    sheets.map((sheet) =>
      fillSheetPeriod(sheet, label, sheetSlot(sheet.getName()).position, timeZone)
    )
  );
}

// ---------- pure helpers (no Spreadsheet calls) ----------

function describePeriod(position) {
  const { year, month, half } = periodMonth(position);
  const monthDays = daysInMonth(year, month);
  const isSecondHalf = PERIODS_PER_MONTH === 2 && half === 1;
  return {
    year,
    month,
    isSecondHalf,
    firstDay: isSecondHalf ? FIRST_HALF_LAST_DAY + 1 : 1,
    lastDay: PERIODS_PER_MONTH === 1 || isSecondHalf ? monthDays : FIRST_HALF_LAST_DAY,
    monthDays,
  };
}

function buildDayHeaders(period) {
  return Array.from({ length: PERIOD_DAY_COLUMNS }, (_, index) => {
    const day = period.firstDay + index;
    return day <= period.lastDay ? day : '';
  });
}

function buildTitle(businessLabel, period) {
  return `${businessLabel} ${period.year} ${period.month} сарын ${period.firstDay}-${period.lastDay}`;
}

function toIsoDate(year, month, day) {
  const pad = (value) => String(value).padStart(2, '0');
  return `${year}-${pad(month)}-${pad(day)}`;
}

// ---------- spreadsheet access ----------

function fillSheetPeriod(sheet, chainKey, position, timeZone) {
  const name = sheet.getName();
  const period = describePeriod(position);
  const title = buildTitle(PERIOD_BUSINESS_LABELS[chainKey] || chainKey, period);
  const days = buildDayHeaders(period);
  const firstIso = toIsoDate(period.year, period.month, period.firstDay);
  const lastIso = toIsoDate(period.year, period.month, period.lastDay);
  const before = String(sheet.getRange(2, 1).getValue()).trim();
  if (!DRY_RUN) {
    sheet.getRange(1, PERIOD_SELLER_LABEL_COL).setValue(PERIOD_SELLER_LABEL);
    sheet.getRange(1, PERIOD_FIRST_DAYS_START_COL, 1, PERIOD_DAY_COLUMNS).setValues([days]);
    sheet.getRange(1, PERIOD_SECOND_DAYS_START_COL, 1, PERIOD_DAY_COLUMNS).setValues([days]);
    sheet.getRange(2, 1).setValue(title);
    installViewControl(sheet); // A1: the list that shows or hides columns
    installNameRule(sheet); // column A: a name that is already on the sheet is refused
  }
  const change = before === title ? 'гарчиг хэвээр' : `"${before}" -> "${title}"`;
  return `${name}: ${firstIso}-нд эхэлж ${lastIso}-нд дуусна | ${change}`;
}
