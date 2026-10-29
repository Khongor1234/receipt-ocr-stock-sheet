/**
 * Moves the two shops' workbook to the paired layout (config.gs buildColumnLayout, PAIRED_SHOPS
 * in profile.gs), once (Тосгон, SHELF_COLUMN in its profile: only the first column, 2026-10-22):
 *   - a new column after D, "яг харагдах үнэ": the price the shop really asks, typed by hand
 *     (mom rounds the price + 3% to an easy number, so it is not a formula); every day,
 *     transfer and count column moves one to the right
 *   - a new column after the old "sold" column (AN): the old "sold" column and the new one show
 *     the OTHER shop's final balance and count (calc-formulas.gs); the totals move to AQ..AU and
 *     "Нийт борлуулалт" is summed from the expense days instead of the "sold" column
 * Inserting the columns (not rewriting them) lets Google move everything that points at the
 * cells: formulas, Sheet200's references to the totals, notes, colours, validations, protections.
 * What the scripts remember by column number is moved here: the owner's approved colours
 * (dividers.gs). The remembered header of each sheet is taken again (header-guard.gs).
 *
 * Two passes: first the columns of EVERY year sheet are inserted, then the formulas are put in.
 * The other shop's columns point at the other shop's sheet; were they put in while that sheet
 * still had the old layout, its own column insert would push the reference one column too far.
 * A sheet's columns are moved when row 3 of column E holds SHELF_PRICE_HEADER; the second pass
 * only writes what differs, so a run cut off by Google's time limit is simply started again.
 * Until both passes are done, the hourly job, the edit handlers and setupYear do nothing
 * (isPairedLayoutReady): their code already counts columns the new way.
 *   migrateToPairedLayout(): with DRY_RUN = true it only reports what it would do.
 * The receipt OCR (its own project) moves its cell references and its check copy with
 * ocrMigrateToPairedLayout (ocr-check.gs). Needs config.gs, year-setup.gs, calc-formulas.gs,
 * dividers.gs, header-guard.gs, view-columns.gs, year-sync.gs.
 */

const PAIRED_LAYOUT_READY_KEY = 'pairedLayoutReady'; // document property: '1' once every sheet is moved
const PAIRED_FINISHED_KEY = 'pairedFinished:'; // + sheet id: '1' once pass 2 has done that sheet
const PAIRED_BORDERS_RESET_KEY = 'pairedBordersReset'; // '1' once redrawTableBorders took the border marks away
const PAIRED_MIGRATION_TIME_BUDGET_MS = 4 * 60 * 1000;
const PAIRED_HEADER_ROW = 3;

// ---------- pure helpers (no Spreadsheet calls) ----------

// Where a column of a sheet not moved yet is once moved: the shops' two columns (pairedColumnOf),
// or Тосгон's one, the shelf price in E.
const movedColumnOf = (column) => (LAYOUT.paired ? pairedColumnOf(column) : column + (column >= SHELF_PRICE_COLUMN ? 1 : 0));

// The owner's approved colours { productKey: { column: color } } with the columns moved.
function movePairedColors(saved) {
  return Object.fromEntries(
    Object.entries(saved).map(([key, cells]) => [
      key,
      Object.fromEntries(Object.entries(cells).map(([column, color]) => [String(movedColumnOf(Number(column))), color])),
    ])
  );
}

// ---------- spreadsheet access ----------

// Also waits for the gift column (gift-migration.gs) where this workbook has one.
function isPairedLayoutReady() {
  const isPaired = !LAYOUT.shelfPrice || bookProperties().getProperty(PAIRED_LAYOUT_READY_KEY) === '1';
  return isPaired && isGiftColumnReady();
}

// For an entry point that works with the columns: true when it may go on, else it logs why not.
function requirePairedLayout(what) {
  if (isPairedLayoutReady()) {
    return true;
  }
  Logger.log(`${what}: хуудсууд шинэ бүтэц рүү хараахан шилжээгүй тул юу ч хийсэнгүй. migrateToPairedLayout${LAYOUT.gift ? ', дараа нь migrateToGiftColumn' : ''}-ийг ажиллуулна уу.`);
  return false;
}

const isSheetPaired = (sheet) => String(sheet.getRange(PAIRED_HEADER_ROW, SHELF_PRICE_COLUMN).getValue()).trim() === SHELF_PRICE_HEADER;

function migrateToPairedLayout() {
  if (!LAYOUT.shelfPrice) {
    Logger.log('Энэ файл хоёр дэлгүүрийн бүтэцтэй биш: шилжүүлэх юм алга.');
    return;
  }
  withScriptLock(() => {
    const lines = runPairedMigration(currentBook(), Date.now() + PAIRED_MIGRATION_TIME_BUDGET_MS);
    Logger.log('%s\n%s', modeLabel(), lines.join('\n'));
  });
}

function runPairedMigration(spreadsheet, deadline) {
  // every year sheet, also one with few priced products yet (findChains would leave it out)
  const sheets = spreadsheet.getSheets().filter((sheet) => sheetSlot(sheet.getName()) !== null);
  const toMove = sheets.filter((sheet) => !isSheetPaired(sheet));
  if (DRY_RUN) {
    return [...toMove.map(describePairedMove), `${toMove.length} хуудсанд багана нэмж, дараа нь ${sheets.length} хуудсанд томьёо тавина (туршилт).`];
  }
  const lines = [];
  let left = 0;
  toMove.forEach((sheet) => {
    if (Date.now() > deadline) {
      left += 1;
      return;
    }
    lines.push(movePairedColumns(sheet));
  });
  if (left > 0) {
    return [...lines, `${left} хуудсанд багана нэмэх үлдлээ: migrateToPairedLayout-ийг дахин ажиллуулна уу.`];
  }
  // a sheet finished in an earlier run is skipped, so a long second pass gets through in parts
  const props = bookProperties();
  const finishedKey = (sheet) => PAIRED_FINISHED_KEY + sheet.getSheetId();
  let unfinished = 0;
  sheets.filter((sheet) => props.getProperty(finishedKey(sheet)) !== '1').forEach((sheet) => {
    if (Date.now() > deadline) {
      unfinished += 1;
      return;
    }
    lines.push(...finishPairedSheet(sheet));
    props.setProperty(finishedKey(sheet), '1');
  });
  if (unfinished > 0) {
    return [...lines, `${unfinished} хуудсанд томьёо тавих үлдлээ: migrateToPairedLayout-ийг дахин ажиллуулна уу.`];
  }
  bookProperties().setProperty(PAIRED_LAYOUT_READY_KEY, '1');
  return [...lines, `Бүх ${sheets.length} хуудас шинэ бүтэцтэй боллоо. Цагийн ажил, шууд засвар дахин ажиллана.`];
}

const describePairedMove = (sheet) =>
  LAYOUT.paired
    ? `${sheet.getName()}: D-ийн дараа "${SHELF_PRICE_HEADER}", AN-ийн дараа нэг багана нэмж, ${columnLetter(LAYOUT.otherBalance)}:${columnLetter(LAYOUT.otherCount)}-д нөгөө дэлгүүрийг харуулна`
    : `${sheet.getName()}: D-ийн дараа "${SHELF_PRICE_HEADER}" багана нэмж, өдрүүд F-ээс эхэлнэ`;

// Has the first column (the shelf price) been inserted? Row 1 holds the day numbers: E1 in the
// old layout, F1 once a column is in front of them.
const hasShelfColumn = (sheet) => {
  const [e1, f1] = sheet.getRange(1, SHELF_PRICE_COLUMN, 1, 2).getValues()[0];
  return typeof e1 !== 'number' && typeof f1 === 'number';
};

// Has the second column (the other shop's count) been inserted, after the first? The totals'
// labels in row 1 start at AQ then, at AP before.
const hasOtherCountColumn = (sheet) => !String(sheet.getRange(1, LAYOUT.otherCount).getValue()).trim().startsWith(SUMMARY_LABEL_PREFIX);

// Pass 1: the two new columns of one sheet of the old layout. Each step looks whether it is
// done already, so a run stopped halfway goes on where it stopped; the E3 header, the mark of a
// moved sheet, is written last. (Formats are not copied from D: row 2's title is merged across
// the columns and a paste into it fails. Pass 2 gives E the money format.) Returns a log line.
function movePairedColumns(sheet) {
  const name = sheet.getName();
  const priceColumn = SHELF_PRICE_COLUMN - 1; // D
  if (!hasShelfColumn(sheet)) {
    sheet.insertColumnsAfter(priceColumn, 1);
  }
  if (LAYOUT.paired && !hasOtherCountColumn(sheet)) {
    sheet.insertColumnsAfter(LAYOUT.otherBalance, 1); // the old "sold" column is AO now; AP is new
  }
  sheet.setColumnWidth(SHELF_PRICE_COLUMN, sheet.getColumnWidth(priceColumn));
  const props = bookProperties();
  const colorsKey = approvedColorsKey(sheet);
  const savedColors = props.getProperty(colorsKey);
  if (savedColors) {
    props.setProperty(colorsKey, JSON.stringify(movePairedColors(JSON.parse(savedColors))));
  }
  // The remembered header has the old columns: it would fill the new empty E1 with the old E1
  // (a day number). It is taken again from the sheet in pass 2.
  props.deleteProperty(HEADER_KEY + sheet.getSheetId());
  sheet.getRange(PAIRED_HEADER_ROW, SHELF_PRICE_COLUMN).setValue(SHELF_PRICE_HEADER).setBackground(sheet.getRange(PAIRED_HEADER_ROW, priceColumn).getBackground());
  return LAYOUT.paired ? `${name}: E "${SHELF_PRICE_HEADER}", ${columnLetter(LAYOUT.otherCount)} багана нэмлээ` : `${name}: E "${SHELF_PRICE_HEADER}" багана нэмлээ`;
}

// After the move: the table's grid (boxes on A..E and AL..AP, dividers.gs DIVIDER_GRID_ZONES) is
// drawn again on every year sheet. The divider pass draws it once per sheet (its border mark);
// the marks are taken away once (PAIRED_BORDERS_RESET_KEY), so a run cut off by the time limit
// simply goes on where it stopped.
function redrawTableBorders() {
  const props = bookProperties();
  const spreadsheet = currentBook();
  if (!DRY_RUN && props.getProperty(PAIRED_BORDERS_RESET_KEY) !== '1') {
    spreadsheet.getSheets()
      .filter((sheet) => sheetSlot(sheet.getName()) !== null)
      .forEach((sheet) => props.deleteProperty(DIVIDER_BORDERS_KEY + sheet.getSheetId()));
    props.setProperty(PAIRED_BORDERS_RESET_KEY, '1');
  }
  protectDividers();
}

// Pass 2, once every sheet has its new columns: the formulas, the totals, the money format, the
// shown columns and the remembered header. Writes only what differs. Returns log lines.
function finishPairedSheet(sheet) {
  const lines = [installOnSheet(sheet), installSummary(sheet), formatPriceColumns(sheet)].filter((line) => line);
  applyView(sheet);
  rememberHeader(sheet);
  return lines;
}
