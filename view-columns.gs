/**
 * Lets the person doing the data entry hide the income or the expense columns.
 *
 * Rows 1-3 are protected, and Google Sheets does not let another editor hide a
 * column that holds a protected cell. So the columns are hidden by this script
 * instead (it runs as the owner): cell A1 of every year sheet is a small list,
 *   Бүгд            every column is shown
 *   Зөвхөн орлого   the expense columns U..AJ are hidden, and the prices C..D (C..E)
 *   Зөвхөн зарлага  the income columns E..T are hidden, and the prices C..D (C..E)
 *   Хэвлэх (нийт дүнгүй)  for printing: cost and price C..D and the totals
 *                   AO..AS are hidden
 *   Зөвхөн буруу    only the products whose check (AM) says "буруу" are shown:
 *                   a filter on row 3 hides every other row, and the totals
 *                   (AQ..AU on the paired layout) are hidden. Choosing anything
 *                   else removes that filter again. The filter is set when the
 *                   choice is made; a row that turns "зөв" afterwards stays until
 *                   the choice is made again. The rows are not sorted by price:
 *                   see зөвхөн-буруу-үнээр-эрэмбэлэх-боломжгүй-2026-10-23.md.
 * and the "On edit" trigger (instant-sync.gs) applies the choice within seconds.
 * The list sits in A1 because columns A and B are never hidden, so it can always
 * be reached to show the columns again. A1 is left open by protect-sheets.gs and
 * is not put back by header-guard.gs. A hidden column is hidden for everyone who
 * has the file open, until the choice is changed.
 *
 * From the 10/31 sheets on (SPOT_COUNTS_FROM, config.gs) a sheet also takes the counts made
 * between its period ends, a block of three columns a counted day after the totals, and its
 * list is a row of the hidden sheet "_харах", so the OCR adds a counted day by adding a cell:
 *   1–15, 16–31     (a month sheet: Тосгон) the days of that half only
 *   Тооллого 10/23  the days up to the 23rd, that day's block; the period-end count, balance,
 *                   check and totals and the later days are hidden
 * "Бүгд" shows the blocks too; the other choices hide them.
 *
 * setupYear installs the list on every year sheet (period-headers.gs).
 * DRY_RUN (config.gs) makes it report only. Needs config.gs and period-headers.gs.
 */

const VIEW_ROW = 1;
const VIEW_CHOICE_COLUMN = 1; // A1
const VIEW_FILTER_HEADER_ROW = 3; // the labels the filter hangs on
const VIEW_CHECK_COLUMN = LAYOUT.check; // AM: "зөв" / "буруу" (calc-formulas.gs)
const VIEW_INCOME_COLUMNS = { first: LAYOUT.incomeFirst, count: LAYOUT.days + 1 }; // E..T (T: transfer in)
const VIEW_EXPENSE_COLUMNS = { first: LAYOUT.expenseFirst, count: LAYOUT.expenseLast - LAYOUT.expenseFirst + 1 }; // U..AJ (AJ: transfer out; then the gift column)
const VIEW_PRICE_COLUMNS = { first: 3, count: LAYOUT.incomeFirst - 3 }; // C..D (cost and price; the shops' layout: C..E with the shelf price)
const VIEW_TOTAL_COLUMNS = { first: LAYOUT.summaryFirst, count: TOTAL_COUNT }; // the five totals: AO..AS, AQ..AU on the paired layout

// ---------- pure helpers (no Spreadsheet calls) ----------

// Which blocks are shown for a choice. Anything unknown shows everything.
function planView(choice) {
  const text = String(choice).trim();
  return {
    showsIncome: text !== VIEW_EXPENSE_ONLY,
    showsExpense: text !== VIEW_INCOME_ONLY,
    showsPrices: ![VIEW_PRINT, VIEW_INCOME_ONLY, VIEW_EXPENSE_ONLY].includes(text),
    showsTotals: text !== VIEW_PRINT && text !== VIEW_WRONG_ONLY,
    showsWrongOnly: text === VIEW_WRONG_ONLY,
  };
}

// The day columns a choice shows, as offsets from the first one: { from, to }. firstDay: the
// sheet's first day (16 on a second-half sheet).
function viewDayRange(choice, days, firstDay) {
  const text = String(choice).trim();
  const half = VIEW_HALVES.indexOf(text);
  if (half >= 0 && days > FIRST_HALF_LAST_DAY) {
    return half === 0 ? { from: 0, to: FIRST_HALF_LAST_DAY - 1 } : { from: FIRST_HALF_LAST_DAY, to: days - 1 };
  }
  const spot = spotDayOf(text);
  return { from: 0, to: spot === null ? days - 1 : Math.min(days - 1, spot.day - firstDay) };
}

// Whether each column 1..maxColumns is shown (true), hidden (false) or left as it is (null).
// blocks: the counted days of the sheet, [{ first, month, day }]. isSpotSheet: a sheet that takes
// them; on the others the count..sold columns and what lies after the totals are left as they are.
function viewColumnFlags(choice, layout, firstDay, blocks, maxColumns, isSpotSheet) {
  const plan = planView(choice);
  const spot = spotDayOf(choice);
  const range = viewDayRange(choice, layout.days, firstDay);
  const inRange = (offset) => offset >= range.from && offset <= range.to;
  const isAll = String(choice).trim() === VIEW_ALL;
  const flagOf = (column) => {
    if (column < VIEW_PRICE_COLUMNS.first) return null;
    if (column < layout.incomeFirst) return plan.showsPrices;
    if (column < layout.transferIn) return plan.showsIncome && inRange(column - layout.incomeFirst);
    if (column === layout.transferIn) return plan.showsIncome;
    if (column < layout.transferOut) return plan.showsExpense && inRange(column - layout.expenseFirst);
    if (column <= layout.expenseLast) return plan.showsExpense;
    if (column < layout.summaryFirst) return isSpotSheet ? spot === null : null;
    if (column < spotFirstColumn(layout)) return plan.showsTotals && spot === null;
    if (!isSpotSheet) return null;
    const block = blocks.find((candidate) => column >= candidate.first && column < candidate.first + SPOT_BLOCK_WIDTH);
    if (!block) return spot === null;
    return spot === null ? isAll : block.month === spot.month && block.day === spot.day;
  };
  return Array.from({ length: maxColumns }, (_, index) => flagOf(index + 1));
}

// Runs of neighbouring columns that are shown or hidden alike: [{ first, count, isShown }].
function viewRuns(flags) {
  return flags.reduce((runs, isShown, index) => {
    const last = runs[runs.length - 1];
    if (isShown === null) return runs;
    if (last && last.isShown === isShown && last.first + last.count === index + 1) {
      return [...runs.slice(0, -1), { ...last, count: last.count + 1 }];
    }
    return [...runs, { first: index + 1, count: 1, isShown }];
  }, []);
}

// edit: { sheetName, firstRow, firstColumn, lastColumn } (describeEdit in instant-sync.gs)
function isViewEdit(edit) {
  return (
    sheetSlot(edit.sheetName) !== null &&
    edit.firstRow === VIEW_ROW &&
    edit.firstColumn <= VIEW_CHOICE_COLUMN &&
    edit.lastColumn >= VIEW_CHOICE_COLUMN
  );
}

// ---------- spreadsheet access ----------

// The period of a year sheet (describePeriod, period-headers.gs) and whether it takes spot counts.
function viewPeriodOf(sheet) {
  const slot = sheetSlot(sheet.getName());
  if (slot === null) {
    return null;
  }
  const period = describePeriod(slot.position);
  return { ...period, isSpotSheet: takesSpotCounts(toIsoDate(period.year, period.month, period.lastDay)) };
}

// The counted days of a sheet: the blocks whose first header (row 3) reads "Тооллого 10/23".
function spotBlocksOf(sheet) {
  const first = spotFirstColumn(LAYOUT);
  const width = sheet.getMaxColumns() - first + 1;
  if (width <= 0) {
    return [];
  }
  return sheet
    .getRange(LAYOUT_HEADER_ROW, first, 1, width)
    .getValues()[0]
    .map((text, index) => ({ first: first + index, day: spotDayOf(text) }))
    .filter((block) => block.day !== null)
    .map((block) => ({ first: block.first, month: block.day.month, day: block.day.day }));
}

// Shows and hides the columns as A1 says.
function applyView(sheet) {
  const period = viewPeriodOf(sheet);
  if (DRY_RUN || period === null) {
    return;
  }
  const choice = sheet.getRange(VIEW_ROW, VIEW_CHOICE_COLUMN).getValue();
  const blocks = period.isSpotSheet ? spotBlocksOf(sheet) : [];
  viewRuns(viewColumnFlags(choice, LAYOUT, period.firstDay, blocks, sheet.getMaxColumns(), period.isSpotSheet)).forEach((run) =>
    run.isShown ? sheet.showColumns(run.first, run.count) : sheet.hideColumns(run.first, run.count)
  );
  applyWrongFilter(sheet, planView(choice).showsWrongOnly);
}

// True for the filter this script sets; a filter someone made by hand is another one.
function isWrongFilter(filter) {
  const criteria = filter.getColumnFilterCriteria(VIEW_CHECK_COLUMN);
  return criteria !== null && String(criteria.getCriteriaValues()[0]) === CALC_MISMATCH_TEXT;
}

// A sheet holds one filter at most. Ours is removed when another view is chosen;
// a filter made by hand is left alone unless "Зөвхөн буруу" needs its place.
function applyWrongFilter(sheet, isWanted) {
  const existing = sheet.getFilter();
  if (existing !== null && (isWanted || isWrongFilter(existing))) {
    existing.remove();
  }
  if (!isWanted || sheet.getMaxColumns() < VIEW_CHECK_COLUMN) {
    return;
  }
  const rows = sheet.getMaxRows() - VIEW_FILTER_HEADER_ROW + 1;
  sheet
    .getRange(VIEW_FILTER_HEADER_ROW, 1, rows, sheet.getMaxColumns())
    .createFilter()
    .setColumnFilterCriteria(
      VIEW_CHECK_COLUMN,
      SpreadsheetApp.newFilterCriteria().whenTextEqualTo(CALC_MISMATCH_TEXT).build()
    );
}

// True when A1 offers exactly the choices of this version.
function hasCurrentList(rule) {
  return rule !== null && JSON.stringify(rule.getCriteriaValues()[0]) === JSON.stringify(VIEW_CHOICES);
}

// Puts the list in place. Returns true when something was missing (a sheet that takes spot
// counts: always, its row of "_харах" is put in order each time).
function installViewControl(sheet) {
  const period = viewPeriodOf(sheet);
  if (period !== null && period.isSpotSheet) {
    return installViewList(sheet);
  }
  const choice = sheet.getRange(VIEW_ROW, VIEW_CHOICE_COLUMN);
  const hasChoice = VIEW_CHOICES.includes(String(choice.getValue()).trim());
  const hasList = hasCurrentList(choice.getDataValidation());
  if (hasChoice && hasList) {
    return false;
  }
  if (!DRY_RUN) {
    if (!hasList) {
      choice.setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(VIEW_CHOICES, true)
          .setAllowInvalid(false)
          .build()
      );
    }
    if (!hasChoice) {
      choice.setValue(VIEW_ALL);
    }
  }
  return true;
}

// A1 of a sheet that takes spot counts: its row of "_харах" (the counted days listed are kept).
function installViewList(sheet) {
  if (DRY_RUN) {
    return true;
  }
  const range = viewListRange(sheet.getParent(), sheet.getName(), viewBaseChoices(LAYOUT));
  setViewListRule(sheet, range);
  const choices = range.getValues()[0].map((value) => String(value).trim()).filter((value) => value !== '');
  const choice = sheet.getRange(VIEW_ROW, VIEW_CHOICE_COLUMN);
  if (!choices.includes(String(choice.getValue()).trim())) {
    choice.setValue(VIEW_ALL);
  }
  return true;
}
