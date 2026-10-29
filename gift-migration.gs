/**
 * Adds the gift column (config.gs GIFT_COLUMN) to every year sheet of the shops' workbook, once:
 * a column "Бэлэг" right after the transfer out (AK; the new column is AL), so the count, the final
 * balance, the check, the other shop's columns and the totals move one to the right. Mom types a
 * gift there (not in the expense days): it leaves the shop (the final balance takes it off) but is
 * not sold (the sales leave it out) and its goods were paid for (the sales profit takes it off at
 * cost, year-setup.gs).
 *
 * Inserting the column (not rewriting) lets Google move everything that points at the cells:
 * formulas, Sheet200's references to the totals, notes, colours, validations, protections. What
 * the scripts remember by column number is moved here, like in the paired move
 * (layout-migration.gs): the owner's approved colours (dividers.gs) and the remembered header
 * (header-guard.gs, taken again). Two passes: first the column goes into EVERY year sheet, then
 * the formulas are put in (the other shop's columns read the other shop's sheet, so not before
 * that sheet has its column too), then the stock sheets (their INDIRECT text does not move); the
 * table borders are drawn again by the divider pass. A sheet has the column when row 3 right after
 * the transfer out says GIFT_HEADER; a run cut off by the time limit is simply started again.
 * Until it is done the hourly job, the edit handlers and setupYear do nothing
 * (isPairedLayoutReady, layout-migration.gs).
 *   migrateToGiftColumn(): with DRY_RUN = true it only reports what it would do.
 * The receipt OCR moves its cell references and its check copy by itself (ocr-layout.gs).
 * Needs config.gs, layout-migration.gs, year-setup.gs, calc-formulas.gs, dividers.gs,
 * header-guard.gs, view-columns.gs, stock-sheets.gs, year-sync.gs.
 */

const GIFT_HEADER_FONT_SIZE = 9; // the owner's choice (2026-10-21), as set on the 48 sheets of 2026
// The gift column's own tint, a little apart from the transfer out before it (the owner's choice,
// 2026-10-21): rows 1-3 and the product rows. Category rows keep their colour; a cell the receipt
// OCR coloured (DIVIDER_OCR_COLORS) keeps it. Not one of the OCR's colours, so dividers.gs takes it
// as the column's own colour (findColumnColors) and never clears it.
const GIFT_HEADER_COLOR = '#c9e4c1';
const GIFT_BODY_COLOR = '#e2efda';
const GIFT_READY_KEY = 'giftColumnReady'; // document property: '1' once every sheet has the column
const GIFT_FINISHED_KEY = 'giftFinished:'; // + sheet id: '1' once pass 2 has done that sheet
const GIFT_COLUMN_WIDTH = 60;
const GIFT_MIGRATION_TIME_BUDGET_MS = 4 * 60 * 1000;

// ---------- pure helpers (no Spreadsheet calls) ----------

// The owner's approved colours { productKey: { column: color } } with the columns moved.
function moveGiftColors(saved) {
  return Object.fromEntries(
    Object.entries(saved).map(([key, cells]) => [
      key,
      Object.fromEntries(Object.entries(cells).map(([column, color]) => [String(giftColumnOf(Number(column))), color])),
    ])
  );
}

// ---------- spreadsheet access ----------

const isGiftColumnReady = () => !LAYOUT.gift || bookProperties().getProperty(GIFT_READY_KEY) === '1';

const sheetHasGiftColumn = (sheet) =>
  String(sheet.getRange(LAYOUT_HEADER_ROW, giftHeaderColumn(DAY_COLUMNS)).getValue()).trim() === GIFT_HEADER;

function migrateToGiftColumn() {
  if (!LAYOUT.gift) {
    Logger.log('Энэ файлд бэлгийн багана байхгүй (GIFT_COLUMN): шилжүүлэх юм алга.');
    return;
  }
  withScriptLock(() => {
    const lines = runGiftMigration(currentBook(), Date.now() + GIFT_MIGRATION_TIME_BUDGET_MS);
    Logger.log('%s\n%s', modeLabel(), lines.join('\n'));
  });
}

function runGiftMigration(spreadsheet, deadline) {
  const props = bookProperties();
  if (props.getProperty(PAIRED_LAYOUT_READY_KEY) !== '1') {
    return ['ЗОГСЛОО: эхлээд migrateToPairedLayout (яг харагдах үнэ) дуусах ёстой'];
  }
  const sheets = spreadsheet.getSheets().filter((sheet) => sheetSlot(sheet.getName()) !== null);
  const toMove = sheets.filter((sheet) => !sheetHasGiftColumn(sheet));
  const out = columnLetter(PAIRED_BASE_LAYOUT.transferOut);
  if (DRY_RUN) {
    return [
      ...toMove.map((sheet) => `${sheet.getName()}: ${out}-ийн дараа "${GIFT_HEADER}" багана нэмнэ`),
      `${toMove.length} хуудсанд багана нэмж, дараа нь ${sheets.length} хуудсанд томьёо тавина (туршилт).`,
    ];
  }
  const lines = [];
  let left = 0;
  toMove.forEach((sheet) => {
    if (Date.now() > deadline) {
      left += 1;
      return;
    }
    lines.push(insertGiftColumn(sheet));
  });
  if (left > 0) {
    return [...lines, `${left} хуудсанд багана нэмэх үлдлээ: migrateToGiftColumn-ийг дахин ажиллуулна уу.`];
  }
  const finishedKey = (sheet) => GIFT_FINISHED_KEY + sheet.getSheetId();
  let unfinished = 0;
  sheets.filter((sheet) => props.getProperty(finishedKey(sheet)) !== '1').forEach((sheet) => {
    if (Date.now() > deadline) {
      unfinished += 1;
      return;
    }
    lines.push(...finishGiftSheet(sheet));
    props.deleteProperty(DIVIDER_BORDERS_KEY + sheet.getSheetId()); // the divider pass draws the grid again
    props.setProperty(finishedKey(sheet), '1');
  });
  if (unfinished > 0) {
    return [...lines, `${unfinished} хуудсанд томьёо тавих үлдлээ: migrateToGiftColumn-ийг дахин ажиллуулна уу.`];
  }
  props.setProperty(GIFT_READY_KEY, '1');
  return [
    ...lines,
    ...runInstallStockSheets(spreadsheet),
    `Бүх ${sheets.length} хуудсанд "${GIFT_HEADER}" багана нэмэгдлээ. Хүрээг protectDividers (цагийн ажил ч) дахин зурна.`,
  ];
}

// Pass 1: the gift column of one sheet. The header, the mark of a moved sheet, is written at once.
function insertGiftColumn(sheet) {
  const out = PAIRED_BASE_LAYOUT.transferOut;
  const gift = out + 1;
  sheet.insertColumnsAfter(out, 1);
  sheet
    .getRange(LAYOUT_HEADER_ROW, gift)
    .setValue(GIFT_HEADER)
    .setFontWeight('bold')
    .setFontSize(GIFT_HEADER_FONT_SIZE);
  sheet.setColumnWidth(gift, GIFT_COLUMN_WIDTH);
  tintGiftColumnOn(sheet, gift);
  const props = bookProperties();
  const colorsKey = approvedColorsKey(sheet);
  const savedColors = props.getProperty(colorsKey);
  if (savedColors) {
    props.setProperty(colorsKey, JSON.stringify(moveGiftColors(JSON.parse(savedColors))));
  }
  // The remembered header has the old columns: it is taken again from the sheet in pass 2.
  props.deleteProperty(HEADER_KEY + sheet.getSheetId());
  return `${sheet.getName()}: ${columnLetter(gift)} "${GIFT_HEADER}" багана нэмлээ`;
}

// Pass 2, once every sheet has its column: the formulas, the totals, the shown columns and the
// remembered header. Writes only what differs. Returns log lines.
function finishGiftSheet(sheet) {
  const lines = [installOnSheet(sheet), installSummary(sheet)].filter((line) => line);
  applyView(sheet);
  rememberHeader(sheet);
  return lines;
}

// ---------- the gift column's own colour ----------

// Run by hand once (DRY_RUN = true first): every year sheet of both shops gets the gift column's tint.
function tintGiftColumn() {
  withScriptLock(() => Logger.log('%s\n%s', modeLabel(), runTintGiftColumn(currentBook()).join('\n')));
}

function runTintGiftColumn(spreadsheet) {
  const sheets = findChains(spreadsheet).flatMap((chain) => chain.sheets);
  const lines = sheets.map((sheet) => {
    const header = (column) => sheet.getRange(LAYOUT_HEADER_ROW, column).getValue();
    const layout = shopLayoutOf(header(SHELF_PRICE_COLUMN), DAY_COLUMNS, header(giftHeaderColumn(DAY_COLUMNS)));
    if (!layout.gift) {
      return `${sheet.getName()}: "${GIFT_HEADER}" багана алга`;
    }
    const cells = tintGiftColumnOn(sheet, layout.gift);
    return `${sheet.getName()}: ${columnLetter(layout.gift)} ${cells} нүд ${DRY_RUN ? 'будагдана' : 'будагдлаа'}`;
  });
  return [`"${GIFT_HEADER}" баганын өнгө: ${sheets.length} sheet`, ...lines];
}

// Rows 1-3 and every product row of the column get the tint; category rows and the OCR's cells keep
// theirs. -> how many cells change
function tintGiftColumnOn(sheet, column) {
  const height = sheet.getLastRow() - DIVIDER_FIRST_ROW + 1;
  const values = height > 0 ? sheet.getRange(DIVIDER_FIRST_ROW, 1, height, DIVIDER_VALUE_WIDTH).getValues() : [];
  const rowColors = height > 0 ? sheet.getRange(DIVIDER_FIRST_ROW, 1, height, DIVIDER_SCAN_WIDTH).getBackgrounds() : [];
  const current = height > 0 ? sheet.getRange(DIVIDER_FIRST_ROW, column, height, 1).getBackgrounds() : [];
  const wanted = current.map(([color], at) => {
    const keep = isDividerRow(rowColors[at], values[at]) || DIVIDER_KEPT_COLORS.includes(String(color).toLowerCase());
    return [keep ? color : GIFT_BODY_COLOR];
  });
  const header = sheet.getRange(1, column, DIVIDER_FIRST_ROW - 1, 1);
  const changed =
    header.getBackgrounds().filter(([color]) => color !== GIFT_HEADER_COLOR).length +
    wanted.filter(([color], at) => color !== current[at][0]).length;
  if (!DRY_RUN && changed > 0) {
    header.setBackground(GIFT_HEADER_COLOR);
    if (height > 0) {
      sheet.getRange(DIVIDER_FIRST_ROW, column, height, 1).setBackgrounds(wanted);
    }
  }
  return changed;
}
