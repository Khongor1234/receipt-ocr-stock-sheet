/**
 * setupYear - prepares the whole year in one run.
 *
 * Steps, in this order (each one skips what is already done, so running it
 * again is always safe):
 *   1. tab names        "Sheet19" -> "My/2026/10/15", "Seet 41" -> "Sor/2026/6/15"
 *   2. missing sheets   every business gets all 24 half-month sheets; a missing
 *                       one is a copy of the sheet before it with the income,
 *                       expense, count and opening balance emptied
 *   3. tab order        tab-order.gs
 *   4. totals           the five "Нийт ..." totals go to AO1:AS1 (labels) and
 *                       AO3:AS3 (values), the same cells on every sheet, each in
 *                       its own colour; old boxes and colours below are removed
 *   5. AL:AN formulas   calc-formulas.gs
 *   6. dates and title  period-headers.gs
 *   7. Sheet200         monthly sales and profit read the fixed total cells
 *   8. money format     cost and price (C, D), the totals and Sheet200 show "₮"
 *   9. charts           charts.gs: monthly sales, profit and profit share on Sheet200
 *  10. header rows      protect-sheets.gs (also makes the sheets undeletable for
 *                       every account except the owner)
 *  11. same categories  a category missing on one sheet is made again (dividers.gs)
 *  12. dividers         dividers.gs: protection, colours
 *  13. empty rows       trim-rows.gs
 *  14. opening balance  B is the previous sheet's count wherever one was counted
 *  15. product lists    baseline (products, cost, price) for the carry-forward sync
 *  16. stock sheets     Sheet100 / Sheet100*: today's stock, locked (stock-sheets.gs)
 *  17. automation       hourly background job, instant carry-forward on edit
 *                       (instant-sync.gs), deleted-sheet e-mail
 *
 * Usage: run setupYear with DRY_RUN = true (config.gs) and read the log. Then set
 * DRY_RUN = false and run it ONCE. One part works for about 4 minutes; if it is
 * not finished it starts itself again a minute later (a one-off trigger) until
 * everything is done, then e-mails the owner "БҮГД ДУУСЛАА". Progress of the
 * automatic parts is under Executions in the Apps Script editor.
 * Run it from the OWNER account: rows 1-3 are protected for everyone else.
 * There is no undo: work on a copy first, or restore from File > Version history.
 *
 * Needs every other .gs file of this project.
 */

const SETUP_TIME_BUDGET_MS = 4 * 60 * 1000;
const SETUP_HEAVY_STEP_RESERVE_MS = 3 * 60 * 1000; // a heavy step starts in the first minute of a part
const SETUP_HANDLER = 'setupYear';
const SETUP_CONTINUE_AFTER_MS = 60 * 1000; // next part starts one minute after a part ends
const SETUP_FALLBACK_AFTER_MS = 7 * 60 * 1000; // in case Google stops a part at its 6-minute limit
const SETUP_MAX_AUTO_RUNS = 15;
const SETUP_RUN_COUNT_KEY = 'setupRunCount';
const SETUP_NEXT_STEP_KEY = 'setupNextStep'; // "<version>|<step not finished yet>|<sheet in that step>"
// Stamped by the build. A part that was started by an older copy of the code does
// not resume in the middle: the new code goes through every step again.
const SETUP_VERSION = 'dev';
const SETUP_MAIL_SUBJECT = 'Тайлангийн файлын бэлтгэл';
const SETUP_FIRST_BODY_ROW = 4;
const SETUP_OPENING_COL = 2; // B
const SETUP_ENTRY_FIRST_COL = LAYOUT.incomeFirst; // E (the shops' layout: F; the shelf price E stays like D)
const SETUP_ENTRY_LAST_COL = LAYOUT.count; // AK

const SUMMARY_FIRST_COL = LAYOUT.summaryFirst; // AO
const SUMMARY_OLD_VALUE_COL = LAYOUT.summaryFirst + 1; // AP, where the totals used to be (rows 4+)
const SUMMARY_LABEL_ROW = 1;
const SUMMARY_VALUE_ROW = 3;
const SUMMARY_LABEL_PREFIX = 'Нийт';
const SUMMARY_COUNT_RANGE = `${columnLetter(LAYOUT.count)}4:${columnLetter(LAYOUT.count)}`; // AK4:AK
// The units sold per row: the "sold" column AN4:AN, or where a layout has none (the shops' layout
// shows the other shop there) the expense days summed per row, V4:AJ without the transfer out.
const SUMMARY_SOLD_RANGE = LAYOUT.sold
  ? `${columnLetter(LAYOUT.sold)}4:${columnLetter(LAYOUT.sold)}`
  : `BYROW(${columnLetter(LAYOUT.expenseFirst)}4:${columnLetter(LAYOUT.transferOut - 1)},LAMBDA(r,SUM(r)))`;
// Gifts (the gift column, config.gs GIFT_COLUMN) bring no money in but their goods were paid for:
// the sales profit takes them off at cost.
const SUMMARY_GIFT_COST = LAYOUT.gift ? `-SUMPRODUCT(C4:C,${columnLetter(LAYOUT.gift)}4:${columnLetter(LAYOUT.gift)})` : '';
// One colour per total so each can be told apart at a glance: label cell (row 1)
// in the stronger tone, value cell (row 3) in the lighter one.
const SUMMARY_LINES = [
  {
    label: 'Нийт анхны үлдэгдэл ашиг',
    formula: '=SUMPRODUCT(B4:B,D4:D)-SUMPRODUCT(B4:B,C4:C)',
    labelColor: '#9fc5e8',
    valueColor: '#deebf7',
  },
  {
    label: 'Нийт анхны үлдэгдэл',
    formula: '=SUMPRODUCT(D4:D,B4:B)',
    labelColor: '#6fa8dc',
    valueColor: '#cfe2f3',
  },
  {
    label: 'Нийт эцсийн үлдэгдэл',
    formula: `=SUMPRODUCT(D4:D,${SUMMARY_COUNT_RANGE})`,
    labelColor: '#b4a7d6',
    valueColor: '#e6e0f3',
  },
  {
    label: 'Нийт борлуулалт',
    formula: `=SUMPRODUCT(D4:D,${SUMMARY_SOLD_RANGE})`,
    labelColor: '#93c47d',
    valueColor: '#d9ead3',
  },
  {
    label: 'Нийт борлуулалтын ашиг',
    formula: `=SUMPRODUCT(D4:D,${SUMMARY_SOLD_RANGE})-SUMPRODUCT(C4:C,${SUMMARY_SOLD_RANGE})${SUMMARY_GIFT_COST}`,
    labelColor: '#ffd966',
    valueColor: '#fff2cc',
  },
];
const SUMMARY_VALUE_FONT_SIZE = 12;
const SUMMARY_COLUMN_WIDTH = 150;
const SIDE_AREA_CLEAN_KEY = 'sideAreaClean:'; // + sheet id, set once the old borders are removed

const MONEY_FORMAT = '#,##0 "₮"'; // 1,234,567 ₮
const MONEY_PRICE_FIRST_COL = 3; // C (cost) and D (price)
const MONEY_PRICE_COL_COUNT = LAYOUT.shelfPrice ? 3 : 2; // the shops' layout: also E, the shelf price

const SHEET200_NAME = 'Sheet200';
const SHEET200_FIRST_ROW = 6;
const SHEET200_FIRST_COL = 3; // C = January ... N = December
const SHEET200_MONTHS = 12;
const SHEET200_MONEY_LAST_ROW = 11; // rows 6-11 are amounts, row 12 is a ratio
const SHEET200_REFERENCE_PATTERN = /(sheet\s*\d+|\/\d{4}\/\d{1,2}\/\d{1,2}')!/i; // old or new tab name
// sales, sales profit ($AR$3, $AS$3)
const SHEET200_TOTAL_CELLS = [3, 4].map((offset) => `$${columnLetter(LAYOUT.summaryFirst + offset)}$3`);

function setupYear() {
  useBook(setupTargetBook());
  if (!requirePairedLayout('setupYear')) {
    return;
  }
  const hasRun = withScriptLock(() => {
    const spreadsheet = currentBook();
    const runNumber = beginSetupRun();
    const deadline = Date.now() + SETUP_TIME_BUDGET_MS;
    const steps = buildSetupSteps(spreadsheet, deadline);
    const resume = readResumePoint(runNumber, steps.length);
    if (resume.step > 0 || resume.item > 0) {
      Logger.log('Өмнөх хэсгийн зогссон газраас үргэлжилж байна: %s', steps[resume.step].title);
    }
    let isComplete = true;
    let isBlocked = false;
    for (let index = resume.step; index < steps.length; index += 1) {
      // A step that cannot stop halfway starts only early in a part (or as its first
      // step), so that it has the rest of Google's 6 minutes to itself.
      const isFirstOfPart = index === resume.step;
      const startBy = steps[index].isHeavy && !isFirstOfPart ? deadline - SETUP_HEAVY_STEP_RESERVE_MS : deadline;
      if (Date.now() > startBy && !(steps[index].isHeavy && isFirstOfPart)) {
        isComplete = false;
        break;
      }
      const result = steps[index].run(index === resume.step ? resume.item : 0);
      Logger.log('== %s\n%s', steps[index].title, result.lines.join('\n'));
      if (!result.complete) {
        isComplete = false;
        isBlocked = result.blocked === true;
        saveResumePoint(index, result.nextItem || 0);
        break;
      }
      saveResumePoint(index + 1, 0);
    }
    const willContinue = !DRY_RUN && !isComplete && !isBlocked && runNumber < SETUP_MAX_AUTO_RUNS;
    const outcome = describeSetupOutcome(isComplete, isBlocked, willContinue);
    Logger.log(outcome);
    finishSetupRun(spreadsheet, outcome, willContinue, runNumber);
    if (isComplete && !DRY_RUN) {
      markYearSetupDone(spreadsheet);
    }
  });
  if (!hasRun) {
    retrySetupLater();
  }
}

// The workbook a setup works on: the new year's file that year-rollover.gs is preparing
// (SETUP_BOOK_KEY), otherwise the workbook of today (year-books.gs).
function setupTargetBook() {
  const id = PropertiesService.getScriptProperties().getProperty(SETUP_BOOK_KEY);
  return id ? SpreadsheetApp.openById(id) : targetBook();
}

// A finished setup of a new year's file: year-rollover.gs goes on to its next step.
function markYearSetupDone(spreadsheet) {
  const properties = PropertiesService.getScriptProperties();
  if (properties.getProperty(SETUP_BOOK_KEY) === spreadsheet.getId()) {
    properties.deleteProperty(SETUP_BOOK_KEY);
    bookProperties(spreadsheet).setProperty(YEAR_SETUP_DONE_KEY, '1');
  }
}

// The hourly job (or an edit) held the lock, so this part did nothing. Without a
// new appointment an unfinished setup would stop here for good.
function retrySetupLater() {
  if (DRY_RUN) {
    return;
  }
  clearSetupTriggers();
  scheduleSetupRun(SETUP_CONTINUE_AFTER_MS);
  Logger.log('Өөр ажил явагдаж байсан тул энэ хэсэг ажиллаагүй. 1 минутын дараа өөрөө дахин эхэлнэ.');
}

// A part that continues an unfinished setup starts at the step, and inside that step
// at the sheet, where the part before it stopped, instead of checking everything
// that is finished again.
function readResumePoint(runNumber, stepCount) {
  const start = { step: 0, item: 0 };
  if (DRY_RUN || runNumber <= 1) {
    return start;
  }
  const saved = String(bookScriptProperties().getProperty(SETUP_NEXT_STEP_KEY) || '');
  const [version, stepText, itemText] = saved.split('|');
  const step = Number(stepText);
  const item = Number(itemText);
  const isValid =
    version === SETUP_VERSION &&
    Number.isInteger(step) && step >= 0 && step < stepCount &&
    Number.isInteger(item) && item >= 0;
  return isValid ? { step, item } : start;
}

function saveResumePoint(step, item) {
  if (!DRY_RUN) {
    bookScriptProperties().setProperty(SETUP_NEXT_STEP_KEY, `${SETUP_VERSION}|${step}|${item}`);
  }
}

function describeSetupOutcome(isComplete, isBlocked, willContinue) {
  if (DRY_RUN) {
    return 'ТУРШИЛТ - юу ч өөрчлөгдөөгүй. DRY_RUN = false болгоод setupYear-ийг дахин ажиллуулна.';
  }
  if (isComplete) {
    return 'БҮГД ДУУСЛАА';
  }
  if (isBlocked) {
    return 'ДУУСААГҮЙ - дээрх ЗОГСЛОО мөрийг засаад setupYear-ийг дахин ажиллуулна';
  }
  return willContinue
    ? 'ДУУСААГҮЙ - 1 минутын дараа өөрөө үргэлжилнэ, юу ч дарах хэрэггүй. Дуусахад email ирнэ.'
    : 'ДУУСААГҮЙ - автоматаар олон удаа ажиллаад дуусаагүй. setupYear-ийг гараар дахин ажиллуулна.';
}

// ---------- self-continuation: one click, the script restarts itself until done ----------

function clearSetupTriggers() {
  ScriptApp.getProjectTriggers()
    .filter((trigger) => trigger.getHandlerFunction() === SETUP_HANDLER)
    .forEach((trigger) => ScriptApp.deleteTrigger(trigger));
}

function scheduleSetupRun(delayMs) {
  ScriptApp.newTrigger(SETUP_HANDLER).timeBased().after(delayMs).create();
}

function beginSetupRun() {
  clearSetupTriggers();
  if (DRY_RUN) {
    return 0;
  }
  const properties = bookScriptProperties();
  const runNumber = Number(properties.getProperty(SETUP_RUN_COUNT_KEY) || 0) + 1;
  properties.setProperty(SETUP_RUN_COUNT_KEY, String(runNumber));
  if (runNumber < SETUP_MAX_AUTO_RUNS) {
    scheduleSetupRun(SETUP_FALLBACK_AFTER_MS);
  }
  return runNumber;
}

function finishSetupRun(spreadsheet, outcome, willContinue, runNumber) {
  if (DRY_RUN) {
    return;
  }
  clearSetupTriggers();
  if (willContinue) {
    scheduleSetupRun(SETUP_CONTINUE_AFTER_MS);
    return;
  }
  const properties = bookScriptProperties();
  properties.deleteProperty(SETUP_RUN_COUNT_KEY);
  properties.deleteProperty(SETUP_NEXT_STEP_KEY);
  if (runNumber > 1) {
    // Parts started by the trigger have no screen, so the owner is told by e-mail.
    MailApp.sendEmail(
      Session.getEffectiveUser().getEmail(),
      SETUP_MAIL_SUBJECT,
      `"${spreadsheet.getName()}": ${outcome}\n${spreadsheet.getUrl()}`
    );
  }
}

// Does the work item by item (sheet by sheet) and stops when the time of this part
// is used up. It then reports the item it stopped at (nextItem), and the next part
// goes on from exactly there (startItem) instead of looking at the finished sheets
// again.
function forEachUntil(items, deadline, work, startItem = 0) {
  const lines = [];
  for (let index = startItem; index < items.length; index += 1) {
    if (Date.now() > deadline) {
      return {
        lines: [...lines, 'хугацаа дууслаа - дараагийн хэсэгт үргэлжилнэ'],
        complete: false,
        nextItem: index,
      };
    }
    const line = work(items[index]);
    if (line !== null && line !== undefined) {
      lines.push(line);
    }
  }
  return { lines: lines.length > 0 ? lines : ['хийх зүйлгүй'], complete: true };
}

// Every step is { title, run(startItem) }. Steps that go sheet by sheet use startItem
// to continue where the part before stopped; the others ignore it.
function buildSetupSteps(spreadsheet, deadline) {
  const done = (lines) => ({ lines, complete: true });
  const visibleSheets = () => spreadsheet.getSheets().filter((sheet) => !sheet.getName().startsWith('_'));
  const chainSheets = () => findChains(spreadsheet).flatMap(({ sheets }) => sheets);
  const eachSheet = (listSheets, work) => (startItem) =>
    forEachUntil(listSheets(), deadline, work, startItem);
  return [
    { title: '1. Sheet-ийн нэр', run: () => done(normalizeSheetNames(spreadsheet)) },
    ...extraSetupSteps(spreadsheet, deadline),
    { title: '2. Дутуу sheet үүсгэх', run: () => createMissingSheets(spreadsheet, deadline) },
    { title: '3. Табын дараалал', run: () => done(runArrangeTabs(spreadsheet)) },
    { title: `4. Нийт дүн (${columnLetter(SUMMARY_FIRST_COL)}1:${columnLetter(SUMMARY_FIRST_COL + TOTAL_COUNT - 1)}3)`, run: eachSheet(chainSheets, installSummary) },
    { title: `5. ${columnLetter(CALC_FIRST_COL)}:${columnLetter(CALC_FIRST_COL + CALC_COLUMN_COUNT - 1)} томьёо`, run: eachSheet(() => listCalcSheets(spreadsheet), installOnSheet) },
    { title: '6. Огноо, гарчиг', run: (startItem) => fillHeadersUntil(spreadsheet, deadline, startItem) },
    { title: '7. Sheet200', run: () => done([rewriteSheet200(spreadsheet)]) },
    { title: '8. Мөнгөн дүнг ₮ хэлбэртэй болгох', run: (startItem) => formatMoneyUntil(spreadsheet, deadline, startItem) },
    { title: '9. Sheet200-ийн график', run: () => done(describeCharts(spreadsheet)) },
    { title: '10. Гарчгийн мөрийн хамгаалалт', run: eachSheet(visibleSheets, protectHeaderRows) },
    { title: '11. Ангиллыг бүх sheet-д адил болгох', isHeavy: true, run: () => done(describeCategorySync(spreadsheet)) },
    { title: '12. Зааглагчийн хамгаалалт', run: eachSheet(visibleSheets, reconcileSheet) },
    { title: '13. Илүү хоосон мөр', run: eachSheet(visibleSheets, trimSheet) },
    { title: '14. Анхны үлдэгдэл', isHeavy: true, run: () => done([runCarry(spreadsheet)]) },
    { title: '15. Барааны жагсаалт', isHeavy: true, run: () => done(saveSyncBaseline(spreadsheet, deadline)) },
    { title: '16. Тооллогын хуудас', run: () => done(runInstallStockSheets(spreadsheet)) },
    { title: '17. Автомат ажил', run: () => done(installAutomation()) },
  ];
}

function fillHeadersUntil(spreadsheet, deadline, startItem) {
  const timeZone = spreadsheet.getSpreadsheetTimeZone();
  const entries = findChains(spreadsheet).flatMap(({ label, sheets }) =>
    sheets.map((sheet) => ({ sheet, label }))
  );
  return forEachUntil(
    entries,
    deadline,
    ({ sheet, label }) => fillSheetPeriod(sheet, label, sheetSlot(sheet.getName()).position, timeZone),
    startItem
  );
}

function formatMoneyUntil(spreadsheet, deadline, startItem) {
  const sheets = findChains(spreadsheet).flatMap(({ sheets: chainSheets }) => chainSheets);
  const result = forEachUntil(sheets, deadline, formatPriceColumns, startItem);
  if (!result.complete) {
    return result;
  }
  const summaryLine = formatSheet200Money(spreadsheet);
  return summaryLine === null ? result : { lines: [...result.lines, summaryLine], complete: true };
}

function describeCategorySync(spreadsheet) {
  const lines = runCategorySync(spreadsheet);
  return lines.length > 0 ? lines : ['бүх sheet ижил ангилалтай байна'];
}

function describeCharts(spreadsheet) {
  const lines = runInstallCharts(spreadsheet);
  return lines.length > 0 ? lines : ['графикууд байгаа'];
}

// ---------- pure helpers (no Spreadsheet calls) ----------

function buildSheet200Formulas() {
  return CHAINS.flatMap((chain) =>
    SHEET200_TOTAL_CELLS.map((cell) =>
      Array.from({ length: SHEET200_MONTHS }, (_, month) => {
        const firstHalf = chainSheetName(chain, month * 2);
        const secondHalf = chainSheetName(chain, month * 2 + 1);
        return `=SUM('${firstHalf}'!${cell},'${secondHalf}'!${cell})`;
      })
    )
  );
}

function normalizeLabel(label) {
  return String(label).trim().replace(/\s+/g, ' ');
}

function describeTotalChanges(oldTotals, newValues) {
  return SUMMARY_LINES.map((line, index) => {
    const old = oldTotals[line.label];
    const current = newValues[index];
    if (old === undefined) {
      return `${line.label}: (өмнө байгаагүй) -> ${current}`;
    }
    return old === current ? null : `${line.label}: ${old} -> ${current}`;
  }).filter((text) => text !== null);
}

// ---------- step 1: tab names ----------

function normalizeSheetNames(spreadsheet) {
  const lines = spreadsheet
    .getSheets()
    // Only real product sheets: an empty or unrelated sheet that happens to be
    // called "Sheet54" is left alone.
    .filter((sheet) => sheetSlot(sheet.getName()) !== null && isEntrySheet(sheet))
    .map((sheet) => renameToCanonical(spreadsheet, sheet))
    .filter((line) => line !== null);
  return lines.length > 0 ? lines : ['бүх нэр зөв байна'];
}

function renameToCanonical(spreadsheet, sheet) {
  const name = sheet.getName();
  const wanted = canonicalSheetName(name);
  if (wanted === name) {
    return null;
  }
  if (spreadsheet.getSheetByName(wanted)) {
    return `"${name}": нэрийг ЗАСААГҮЙ, "${wanted}" нэртэй sheet аль хэдийн байна`;
  }
  if (!DRY_RUN) {
    sheet.setName(wanted);
  }
  return `"${name}" -> "${wanted}"`;
}

// ---------- step 2: missing sheets ----------

function createMissingSheets(spreadsheet, deadline) {
  const existing = findChains(spreadsheet);
  const lines = [];
  for (const chain of CHAINS) {
    const found = existing.find((entry) => entry.label === chain.key);
    if (!found) {
      lines.push(`${chain.label}: sheet олдсонгүй, хуулах эх байхгүй`);
      continue;
    }
    const result = fillChain(spreadsheet, chain, found.sheets, deadline);
    lines.push(...result.lines);
    if (!result.complete) {
      return { lines, complete: false, blocked: result.blocked === true };
    }
  }
  return { lines, complete: true };
}

function fillChain(spreadsheet, chain, sheets, deadline) {
  const byPosition = sheets.reduce(
    (all, sheet) => ({ ...all, [sheetSlot(sheet.getName()).position]: sheet }),
    {}
  );
  const lines = [];
  let previous = null;
  for (let position = 0; position < CHAIN_PERIOD_COUNT; position += 1) {
    if (byPosition[position]) {
      previous = byPosition[position];
      continue;
    }
    const name = chainSheetName(chain, position);
    if (previous === null) {
      lines.push(`${name}: БАЙХГҮЙ, өмнө нь хуулах sheet алга`);
      continue;
    }
    if (Date.now() > deadline) {
      lines.push('хугацаа дууслаа - setupYear-ийг дахин ажиллуулна');
      return { lines, complete: false };
    }
    const occupant = spreadsheet.getSheetByName(name);
    if (occupant && occupant.getLastRow() > 0) {
      lines.push(`ЗОГСЛОО: "${name}" нэртэй sheet байгаа боловч барааны sheet биш. Нэрийг нь солих эсвэл устгана уу.`);
      return { lines, complete: false, blocked: true };
    }
    lines.push(`${name}: ${previous.getName()}-ээс хуулж ${DRY_RUN ? 'үүсгэнэ' : 'үүсгэлээ'}`);
    if (DRY_RUN) {
      continue;
    }
    if (occupant) {
      spreadsheet.deleteSheet(occupant); // empty placeholder with the name we need
    }
    previous = copyAsNextPeriod(spreadsheet, previous, name);
  }
  return { lines: lines.length > 0 ? lines : [`${chain.label}: ${CHAIN_PERIOD_COUNT} sheet бүгд байна`], complete: true };
}

function copyAsNextPeriod(spreadsheet, source, name) {
  const copy = source.copyTo(spreadsheet).setName(name);
  spreadsheet.setActiveSheet(copy);
  spreadsheet.moveActiveSheet(source.getIndex() + 1);
  const bodyRows = copy.getMaxRows() - SETUP_FIRST_BODY_ROW + 1;
  if (bodyRows > 0) {
    const entryWidth = SETUP_ENTRY_LAST_COL - SETUP_ENTRY_FIRST_COL + 1;
    copy.getRange(SETUP_FIRST_BODY_ROW, SETUP_ENTRY_FIRST_COL, bodyRows, entryWidth).clearContent();
    copy.getRange(SETUP_FIRST_BODY_ROW, SETUP_OPENING_COL, bodyRows, 1).clearContent();
  }
  return copy;
}

// ---------- step 4: totals ----------

function findOldSummaryRows(sheet) {
  const height = sheet.getLastRow() - SETUP_FIRST_BODY_ROW + 1;
  if (height <= 0 || sheet.getMaxColumns() < SUMMARY_OLD_VALUE_COL) {
    return [];
  }
  return sheet
    .getRange(SETUP_FIRST_BODY_ROW, SUMMARY_FIRST_COL, height, 2)
    .getValues()
    .map(([label, value], index) => ({ label: normalizeLabel(label), value, row: SETUP_FIRST_BODY_ROW + index }))
    .filter((entry) => entry.label.startsWith(SUMMARY_LABEL_PREFIX));
}

function hasSummaryInstalled(sheet) {
  const width = SUMMARY_LINES.length;
  if (sheet.getMaxColumns() < SUMMARY_FIRST_COL + width - 1) {
    return false;
  }
  return sheet
    .getRange(SUMMARY_VALUE_ROW, SUMMARY_FIRST_COL, 1, width)
    .getFormulas()[0]
    .every((formula, index) => formula === SUMMARY_LINES[index].formula);
}

function summaryValueRange(sheet) {
  return sheet.getRange(SUMMARY_VALUE_ROW, SUMMARY_FIRST_COL, 1, SUMMARY_LINES.length);
}

function summaryLabelRange(sheet) {
  return sheet.getRange(SUMMARY_LABEL_ROW, SUMMARY_FIRST_COL, 1, SUMMARY_LINES.length);
}

function hasSummaryStyle(sheet) {
  const sameColors = (range, key) =>
    range.getBackgrounds()[0].every((color, index) => color === SUMMARY_LINES[index][key]);
  return (
    sameColors(summaryLabelRange(sheet), 'labelColor') &&
    sameColors(summaryValueRange(sheet), 'valueColor') &&
    summaryValueRange(sheet)
      .getNumberFormats()[0]
      .every((format) => format === MONEY_FORMAT)
  );
}

function styleSummary(sheet) {
  summaryLabelRange(sheet)
    .setBackgrounds([SUMMARY_LINES.map((line) => line.labelColor)])
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBorder(true, true, true, true, true, true);
  summaryValueRange(sheet)
    .setBackgrounds([SUMMARY_LINES.map((line) => line.valueColor)])
    .setNumberFormat(MONEY_FORMAT)
    .setFontWeight('bold')
    .setFontSize(SUMMARY_VALUE_FONT_SIZE)
    .setHorizontalAlignment('right')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true);
  sheet.setColumnWidths(SUMMARY_FIRST_COL, SUMMARY_LINES.length, SUMMARY_COLUMN_WIDTH);
}

// ---------- the area under the totals (AO4 and below): left-over boxes ----------
// The colours there are put in order together with the rest of the sheet
// (reconcileSheet in dividers.gs); only the old borders are removed here, once.

function sideAreaRange(sheet) {
  const rows = sheet.getMaxRows() - SETUP_FIRST_BODY_ROW + 1;
  const columns = sheet.getMaxColumns() - SUMMARY_FIRST_COL + 1;
  return rows > 0 && columns > 0
    ? sheet.getRange(SETUP_FIRST_BODY_ROW, SUMMARY_FIRST_COL, rows, columns)
    : null;
}

function isSideAreaClean(sheet) {
  const marker = bookProperties().getProperty(
    SIDE_AREA_CLEAN_KEY + sheet.getSheetId()
  );
  return marker !== null;
}

function cleanSideArea(sheet) {
  const range = sideAreaRange(sheet);
  if (range !== null) {
    // The left edge is the right edge of the AN column and the top edge belongs to
    // the totals, so those two are left as they are (null).
    range.setBorder(null, null, false, false, false, false);
  }
  bookProperties().setProperty(SIDE_AREA_CLEAN_KEY + sheet.getSheetId(), '1');
}

function installSummary(sheet) {
  const name = sheet.getName();
  const oldRows = findOldSummaryRows(sheet);
  const needsInstall = !hasSummaryInstalled(sheet) || oldRows.length > 0;
  const needsStyle = needsInstall || !hasSummaryStyle(sheet);
  const needsCleanup = needsInstall || !isSideAreaClean(sheet);
  if (!needsInstall && !needsStyle && !needsCleanup) {
    return `${name}: нийт дүн аль хэдийн тавигдсан`;
  }
  const oldTotals = oldRows.reduce((all, entry) => ({ ...all, [entry.label]: entry.value }), {});
  const tasks = [
    needsInstall ? '5 нийт дүнг AO1:AS3-д тавих' : null,
    needsStyle ? 'нийт дүнг өнгөөр ялгах' : null,
    needsCleanup ? 'доорх хуучин хүрээг арилгах' : null,
  ].filter((task) => task !== null);
  if (DRY_RUN) {
    return `${name}: хийх ажил: ${tasks.join(', ')}`;
  }
  let changes = [];
  if (needsInstall) {
    writeSummary(sheet, oldRows);
    SpreadsheetApp.flush();
    changes = describeTotalChanges(oldTotals, summaryValueRange(sheet).getValues()[0]);
  }
  if (needsStyle) {
    styleSummary(sheet);
  }
  if (needsCleanup) {
    cleanSideArea(sheet);
  }
  const numbers = !needsInstall ? '' : changes.length > 0 ? ' | ' + changes.join(' | ') : ' | бүх тоо хэвээр';
  return `${name}: хийгдлээ: ${tasks.join(', ')}${numbers}`;
}

function writeSummary(sheet, oldRows) {
  const width = SUMMARY_LINES.length;
  const lastCol = SUMMARY_FIRST_COL + width - 1;
  if (sheet.getMaxColumns() < lastCol) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), lastCol - sheet.getMaxColumns());
  }
  oldRows.forEach((entry) => sheet.getRange(entry.row, SUMMARY_FIRST_COL, 1, 2).clearContent());
  summaryLabelRange(sheet).setValues([SUMMARY_LINES.map((line) => line.label)]);
  summaryValueRange(sheet).setFormulas([SUMMARY_LINES.map((line) => line.formula)]);
}

// ---------- step 7: Sheet200 ----------

function rewriteSheet200(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(SHEET200_NAME);
  if (!sheet) {
    return `${SHEET200_NAME}: олдсонгүй (алгаслаа)`;
  }
  const wanted = buildSheet200Formulas();
  const range = sheet.getRange(SHEET200_FIRST_ROW, SHEET200_FIRST_COL, wanted.length, SHEET200_MONTHS);
  const current = range.getFormulas();
  if (JSON.stringify(current) === JSON.stringify(wanted)) {
    return `${SHEET200_NAME}: аль хэдийн шинэчлэгдсэн`;
  }
  const looksRight = current.every((row) => SHEET200_REFERENCE_PATTERN.test(row[0]));
  if (!looksRight) {
    return `${SHEET200_NAME}: АЛГАСЛАА - C6:C9-д sheet-үүдийн нийлбэрийн томьёо (=SUM(...!...)) алга`;
  }
  if (!DRY_RUN) {
    range.setFormulas(wanted);
  }
  return `${SHEET200_NAME}: C6:N9-ийг sheet бүрийн AR3 (борлуулалт), AS3 (ашиг)-аас уншдаг ${DRY_RUN ? 'болгоно' : 'болголоо'}`;
}

// ---------- step 8: money format ----------

function applyMoneyFormat(range) {
  const isDone = range
    .getNumberFormats()
    .every((formats) => formats.every((format) => format === MONEY_FORMAT));
  if (isDone) {
    return false;
  }
  if (!DRY_RUN) {
    range.setNumberFormat(MONEY_FORMAT);
  }
  return true;
}

function formatPriceColumns(sheet) {
  const rows = sheet.getMaxRows() - SETUP_FIRST_BODY_ROW + 1;
  if (rows <= 0) {
    return null;
  }
  const range = sheet.getRange(SETUP_FIRST_BODY_ROW, MONEY_PRICE_FIRST_COL, rows, MONEY_PRICE_COL_COUNT);
  return applyMoneyFormat(range)
    ? `${sheet.getName()}: өртөг, үнийг (C, D${LAYOUT.shelfPrice ? ', E' : ''}) ₮ хэлбэртэй ${DRY_RUN ? 'болгоно' : 'болголоо'}`
    : null;
}

function formatSheet200Money(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(SHEET200_NAME);
  if (!sheet) {
    return null;
  }
  const lastCol = Math.max(sheet.getLastColumn(), SHEET200_FIRST_COL + SHEET200_MONTHS - 1);
  const range = sheet.getRange(
    SHEET200_FIRST_ROW,
    SHEET200_FIRST_COL,
    SHEET200_MONEY_LAST_ROW - SHEET200_FIRST_ROW + 1,
    lastCol - SHEET200_FIRST_COL + 1
  );
  return applyMoneyFormat(range)
    ? `${SHEET200_NAME}: 6-11-р мөрийн дүнг ₮ хэлбэртэй ${DRY_RUN ? 'болгоно' : 'болголоо'}`
    : null;
}

// ---------- steps 15 and 17 ----------

function saveSyncBaseline(spreadsheet, deadline) {
  if (DRY_RUN) {
    return ['sheet бүрийн барааны жагсаалт, үнэ, гарчгийн мөрийг хадгална'];
  }
  return [
    formatSyncReport(runSync(spreadsheet, deadline)),
    formatPriceReport(runPriceSync(spreadsheet)),
    ...runRememberHeaders(spreadsheet),
  ];
}

function installAutomation() {
  if (DRY_RUN) {
    return ['цаг тутмын автомат ажил, шууд дамжуулалт, sheet устгагдвал ирэх email-ийг суулгана'];
  }
  installAutoSync();
  installSheetGuard();
  installInstantSync();
  return ['цаг тутмын автомат ажил, шууд дамжуулалт, sheet устгагдвал ирэх email суулгагдлаа'];
}
