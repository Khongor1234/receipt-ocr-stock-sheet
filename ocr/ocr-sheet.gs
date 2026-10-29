/**
 * Receipt OCR - the day spreadsheet as mom sees it: the view choice in A1 (her view: what was
 * read, her columns and why a line is not in the workbook yet; "Бүгд": every column), the
 * product list in "Бараа" and a Зөв/Буруу list on the lines only (ocrSetLineChoices), rows turning
 * yellow, pink, green, red or grey (the legend in column O says which is what), and who may change
 * what:
 *   - the owner account: everything
 *   - mom's account (the "Ээж" address of OCR_REVIEWERS, the shop account): only her columns; it is
 *     the only editor of the month spreadsheets besides the owner
 *   - nobody else (the staff of the version before lose their access)
 * A tab of the version before (with the staff's columns) is moved into the new layout by its
 * titles (ocrUpgradeTab). Called from ocr-files.gs / ocr-post.gs, only with DRY_RUN = false.
 */

const OCR_GUARD_PREFIX = 'OCR хамгаалалт';

const ocrIsEmail = (value) => /^[^@\s]+@[^@\s]+$/.test(String(value));

const ocrColumn = (title) => OCR_HEADER.indexOf(title) + 1;

// Some neighbouring columns from row `first` to the bottom of the sheet.
const ocrColumnsRange = (sheet, titles, first) => sheet.getRange(first, ocrColumn(titles[0]), sheet.getMaxRows() - first + 1, titles.length);
// Mom's columns ("Бараа" and "Тоо" side by side, her mark further right) from row `first` down.
const ocrMomRanges = (sheet, first) => [ocrColumnsRange(sheet, [OCR_NAME_COLUMN, OCR_QTY_COLUMN], first), ocrColumnsRange(sheet, OCR_MOM_MARKS, first)];
const ocrViewCell = (sheet) => sheet.getRange(1, ocrColumn('Харах'));
const ocrViewOf = (value) => (OCR_VIEWS.includes(ocrText(value)) ? ocrText(value) : OCR_VIEW_MOM);
// The header as written: the view choice in A1, the titles after it.
const ocrHeaderRow = (view) => [view, ...OCR_HEADER.slice(1)];
// A tab in the current layout: A1 one of its choices (the whole day's Зөв stands there a moment),
// the other titles those of OCR_HEADER.
const ocrIsCurrentLayout = (titles) =>
  OCR_VIEW_CHOICES.includes(ocrText(titles[0])) && titles.slice(1, OCR_HEADER.length).map(ocrText).join('|') === OCR_HEADER.slice(1).join('|');

// The list in A1: the views and the whole day's Зөв (OCR_VIEW_CHOICES).
function ocrSetViewChoices(sheet) {
  const choices = SpreadsheetApp.newDataValidation().requireValueInList(OCR_VIEW_CHOICES, true).setAllowInvalid(false).build();
  ocrViewCell(sheet).setDataValidation(choices);
}

function ocrRowRules(sheet) {
  const body = sheet.getRange(2, 1, sheet.getMaxRows() - 1, OCR_HEADER.length);
  const mom = columnLetter(ocrColumn('Ээж: Зөв/Буруу'));
  const status = columnLetter(ocrColumn('Төлөв'));
  const name = columnLetter(ocrColumn(OCR_NAME_COLUMN));
  const picked = columnLetter(ocrColumn('Бараа (тааруулсан)'));
  const qty = columnLetter(ocrColumn(OCR_QTY_COLUMN));
  const read = columnLetter(ocrColumn('Тоо (уншсан)'));
  const rule = (formula, color) => SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(formula).setBackground(color).setRanges([body]).build();
  return [
    // first: it wins over the others; Хамааралгүй the script picked stays pink until mom looked at it
    rule(`=AND($${name}2="${OCR_NOT_RELEVANT}",OR($${picked}2<>"${OCR_NOT_RELEVANT}",$${mom}2<>""))`, OCR_NOT_RELEVANT_COLOR),
    rule(`=$${mom}2="${OCR_VERDICT_WRONG}"`, OCR_WRONG_COLOR),
    rule(`=OR($${mom}2="${OCR_VERDICT_OK}",$${mom}2="${OCR_VERDICT_GIFT}")`, OCR_AGREED_COLOR),
    // mom changed the product or the count: her approval (ocrIsChosenByMom), green as its workbook cell
    rule(`=AND($${mom}2="",OR(AND($${name}2<>"",LOWER(SUBSTITUTE($${name}2," ",""))<>LOWER(SUBSTITUTE($${picked}2," ",""))),` +
      `AND($${read}2<>"",$${qty}2&""<>$${read}2&"")))`, OCR_AGREED_COLOR),
    rule(`=AND($${status}2="${OCR_STATUS_REVIEW}",$${mom}2="")`, OCR_CHECK_COLOR), // waits for mom
    // the AI is sure: yellow, as its cell in the workbook (the owner, 2026-10-25: one colour, one meaning)
    rule(`=AND($${status}2="${OCR_STATUS_OK}",$${mom}2="")`, OCR_WRITTEN_COLOR),
  ];
}

// Right of the reasons, in some spreadsheets only (the owner, 2026-10-27): O the price read from the
// receipt, P our price of the line's product to compare it with (a formula on "Бараа", so it follows
// mom's change). The two shops' sales: the price the customer sees, from the name tab
// (ocrSetNameList); Тосгон's income (notebook pages, prices in yuan): Тосгон's cost of the month, C of
// Tos/YYYY/M, from its hidden cost tab (ocrSetCosts). Neither is a receipt column: the rows are
// counted in A..M.
const OCR_READ_PRICE_TITLE = 'Баримтын үнэ';
const OCR_COSTS_SHEET = '_өртөг'; // a Тосгон income spreadsheet's hidden tab: name, cost
const OCR_PRICED_BOOKS = [
  { place: 'Соронзон', kind: OCR_BOOK_KINDS.zarlaga, title: 'Манай үнэ', lookup: `'${OCR_NAMES_SHEET}'!$C:$D` },
  { place: 'My home', kind: OCR_BOOK_KINDS.zarlaga, title: 'Манай үнэ', lookup: `'${OCR_NAMES_SHEET}'!$C:$D` },
  { place: OCR_TOSGON.place, kind: OCR_BOOK_KINDS.orlogo, title: 'Өртөг үнэ', lookup: `'${OCR_COSTS_SHEET}'!$A:$B` },
];
const OCR_PRICE_TITLES = [OCR_READ_PRICE_TITLE, ...new Set(OCR_PRICED_BOOKS.map((priced) => priced.title))];
const OCR_READ_PRICE_COLUMN = OCR_HEADER.length + 1; // O
const OCR_OUR_PRICE_COLUMN = OCR_HEADER.length + 2; // P
const OCR_PRICE_COLUMNS = 2; // O and P
const OCR_PRICE_FORMAT = '#,##0';
const OCR_PRICE_WIDTH = 90; // px

// The price columns of a month spreadsheet (one of OCR_PRICED_BOOKS), or null: it has none.
function ocrPricedBookOf(book) {
  const month = book && ocrMonthBookOf(book.getName());
  return (month && OCR_PRICED_BOOKS.find((priced) => priced.kind === month.kind && priced.place === month.place)) || null;
}
const ocrPricesOf = (sheet) => ocrPricedBookOf(sheet.getParent());
const ocrHasPrices = (sheet) => ocrPricesOf(sheet) !== null;

// What the colours mean, in column O right of the reasons (Q, right of the prices, where there are
// prices; the owner, 2026-10-25): the title in row 1, then a sentence a cell, each on its own colour (every colour the system puts on means one thing), and
// why a line turns pink. The cells overflow to the right, so a photo's tall row or a short one shows
// them the same. They are no receipt row: the rows are counted in the receipt columns only
// (ocrLastRecordRow). [colour, meaning]
const OCR_LEGEND_COLUMN = OCR_HEADER.length + 1; // O
const OCR_LEGEND_COLUMN_PRICED = OCR_HEADER.length + 3; // Q
const ocrLegendColumn = (sheet) => (ocrHasPrices(sheet) ? OCR_LEGEND_COLUMN_PRICED : OCR_LEGEND_COLUMN);
const OCR_LEGEND_WIDTH = 160; // px, the rest overflows into the empty columns
const OCR_LEGEND_TITLE = 'Өнгө:';
const OCR_LEGEND = [
  [OCR_WRITTEN_COLOR, 'Шар мөр: AI итгэлтэй, шалгах шаардлагагүй. Үндсэн хүснэгтэд ч шар.'],
  [OCR_CHECK_COLOR, 'Ягаан мөр: заавал шалгана (AI эргэлзсэн). Зургийг харж Зөв/Буруу гэх эсвэл бараа, тоог засна. Үндсэн хүснэгтэд ч ягаан.'],
  [OCR_AGREED_COLOR, 'Ногоон мөр: шалгасан (ээж Зөв гэсэн эсвэл бараа, тоог зассан). Үндсэн хүснэгтэд ч ногоон.'],
  [OCR_WRONG_COLOR, 'Улаан мөр: ээж Буруу гэсэн, систем бичсэн тоогоо үндсэн хүснэгтээс хасна.'],
  [OCR_NOT_RELEVANT_COLOR, 'Саарал мөр: хамааралгүй бараа, үндсэн хүснэгтэд бичигдэхгүй.'],
];
const OCR_LEGEND_REASONS =
  'Ягаан болох шалтгаан: мөр бүрийнхийг "Шалгах шалтгаан" баганад бичсэн: нэр жагсаалтад таараагүй, брэнд зөрсөн, хэмжээ/тоо зөрсөн (211 ≠ 210 гэх мэт), ' +
  'үнэ 10%-иас их зөрсөн, тоо эсвэл үнэ уншигдаагүй, AI-ийн итгэл 0.7-оос бага, давхар зураг, хавтас буруу.';
// The legend of before in row 1 from S (one cell, or a cell a colour in S1:Z1): emptied.
const OCR_OLD_LEGEND_COLUMN = 19;
const OCR_OLD_LEGEND_CELLS = 8;
const OCR_OLD_LEGEND_WIDTH = 100; // px, the width those columns had before

// The legend in O1:O7 (Q1:Q7 beside the prices), unless something else is there; in a tab with
// prices the one of before in O is taken away first.
function ocrWriteLegend(sheet) {
  ocrClearOldLegend(sheet);
  const column = ocrLegendColumn(sheet);
  if (column !== OCR_LEGEND_COLUMN) {
    ocrClearLegendAt(sheet, OCR_LEGEND_COLUMN);
  }
  const height = OCR_LEGEND.length + 2;
  if (sheet.getMaxColumns() < column) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), column - sheet.getMaxColumns());
  }
  const range = sheet.getRange(1, column, height, 1);
  const wanted = [OCR_LEGEND_TITLE, ...OCR_LEGEND.map(([, meaning]) => meaning), OCR_LEGEND_REASONS];
  const isOurs = range.getValues().every(([value], index) => ocrText(value) === '' || ocrText(value) === wanted[index]);
  if (!isOurs) {
    return;
  }
  sheet.setColumnWidth(column, OCR_LEGEND_WIDTH);
  range.setValues(wanted.map((text) => [text])).setBackgrounds([[null], ...OCR_LEGEND.map(([color]) => [color]), [null]])
    .setWrap(false).setVerticalAlignment('top');
  sheet.getRange(1, column).setFontWeight('bold');
}

// The legend of before in a column (O of a tab with prices until 2026-10-27), only when it is ours: its title in row 1
// and nothing but its own sentences below.
function ocrClearLegendAt(sheet, column) {
  const height = OCR_LEGEND.length + 2;
  if (sheet.getMaxColumns() < column) {
    return;
  }
  const range = sheet.getRange(1, column, height, 1);
  const values = range.getValues().map(([value]) => ocrText(value));
  const ours = [OCR_LEGEND_TITLE, ...OCR_LEGEND.map(([, meaning]) => meaning), OCR_LEGEND_REASONS];
  if (values[0] !== OCR_LEGEND_TITLE || !values.every((value, index) => value === '' || value === ours[index])) {
    return;
  }
  range.clearContent().setBackground(null).setFontWeight('normal');
}

// The price columns O and P of a tab with prices: their titles, the number format, and our price on
// every line (ocrWriteOurPrices); a line read before 2026-10-27 has no receipt price.
function ocrWritePriceColumns(sheet) {
  const priced = ocrPricesOf(sheet);
  if (!priced) {
    return;
  }
  if (sheet.getMaxColumns() < OCR_OUR_PRICE_COLUMN) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), OCR_OUR_PRICE_COLUMN - sheet.getMaxColumns());
  }
  sheet.getRange(1, OCR_READ_PRICE_COLUMN, 1, OCR_PRICE_COLUMNS).setValues([[OCR_READ_PRICE_TITLE, priced.title]]).setFontWeight('bold');
  sheet.setColumnWidths(OCR_READ_PRICE_COLUMN, OCR_PRICE_COLUMNS, OCR_PRICE_WIDTH);
  if (sheet.getMaxRows() > 1) {
    sheet.getRange(2, OCR_READ_PRICE_COLUMN, sheet.getMaxRows() - 1, OCR_PRICE_COLUMNS).setNumberFormat(OCR_PRICE_FORMAT);
  }
  ocrRuns(ocrReadRecords(sheet).map((record) => record._row)).forEach(([first, count]) => ocrWriteOurPrices(sheet, priced, first, count));
}

// Our price of some lines: the number beside the line's product on the spreadsheet's hidden tab,
// empty for Хамааралгүй or a name without one. INDEX(…,ROW()) keeps the formula right when the line moves.
const ocrOurPriceFormula = (priced) => `=IFERROR(VLOOKUP(INDEX($B:$B,ROW()),${priced.lookup},2,FALSE),"")`;
function ocrWriteOurPrices(sheet, priced, first, count) {
  sheet.getRange(first, OCR_OUR_PRICE_COLUMN, count, 1).setFormulas(Array.from({ length: count }, () => [ocrOurPriceFormula(priced)]));
}

// O and P of lines just written in a tab with prices: the receipt prices (ocrReadPrices; '' where
// none) and the formula.
function ocrWritePrices(sheet, first, count, readPrices = []) {
  const priced = ocrPricesOf(sheet);
  if (count === 0 || !priced) {
    return;
  }
  sheet.getRange(first, OCR_READ_PRICE_COLUMN, count, 1)
    .setValues(Array.from({ length: count }, (_, index) => [readPrices[index] === undefined ? '' : readPrices[index]]));
  ocrWriteOurPrices(sheet, priced, first, count);
}

// Тосгон's costs of a month ("2026-10" -> [[name, cost]] from C of Tos/2026/10), for "Өртөг үнэ" of
// its income lines; [] when the tab can not be read.
function ocrTosgonCosts(month) {
  const [year, monthNumber] = month.split('-').map(Number);
  const tabName = `${OCR_TOSGON.tabPrefix}/${year}/${monthNumber}`;
  try {
    const sheet = SpreadsheetApp.openById(ocrBookIdFor(OCR_TOSGON.place, tabName)).getSheetByName(tabName);
    const count = sheet ? sheet.getLastRow() - OCR_BOOK_FIRST_ROW + 1 : 0;
    const rows = count > 0 ? sheet.getRange(OCR_BOOK_FIRST_ROW, 1, count, 3).getValues() : [];
    return rows.filter(([name, , cost]) => ocrText(name) !== '' && typeof cost === 'number').map(([name, , cost]) => [ocrText(name), cost]);
  } catch (error) {
    Logger.log(`${tabName}: өртөг уншигдсангүй (${error.message})`);
    return [];
  }
}

// The hidden cost tab of a Тосгон income spreadsheet (OCR_COSTS_SHEET), written again whole.
function ocrSetCosts(book, costs) {
  let tab = book.getSheetByName(OCR_COSTS_SHEET);
  if (!tab) {
    tab = book.insertSheet(OCR_COSTS_SHEET, book.getSheets().length);
    tab.hideSheet();
    ocrOnlyOwner(tab.protect().setDescription(`${OCR_GUARD_PREFIX}: Тосгоны өртөг`), Session.getEffectiveUser().getEmail());
  }
  tab.clear();
  if (costs.length > 0) {
    tab.getRange(1, 1, costs.length, 2).setValues(costs);
  }
}

// The legend of before (S1, or S1:Z1), where it is still there.
function ocrClearOldLegend(sheet) {
  if (sheet.getMaxColumns() < OCR_OLD_LEGEND_COLUMN) {
    return;
  }
  const first = ocrText(sheet.getRange(1, OCR_OLD_LEGEND_COLUMN).getValue());
  if (first !== '' && !first.startsWith(OCR_LEGEND_TITLE)) {
    return;
  }
  const cells = Math.min(OCR_OLD_LEGEND_CELLS, sheet.getMaxColumns() - OCR_OLD_LEGEND_COLUMN + 1);
  sheet.getRange(1, OCR_OLD_LEGEND_COLUMN, 1, cells).clearContent().clearNote().setBackground(null);
  sheet.setColumnWidths(OCR_OLD_LEGEND_COLUMN, cells, OCR_OLD_LEGEND_WIDTH);
}

// "Бараа" as wide as its longest name (the owner, 2026-10-28: the names were cut off), between the
// bounds; the list's arrow takes OCR_NAME_ARROW_WIDTH of it. Fitted with the reason column
// (ocrFitColumns) when lines come in, when the tab is set up and when mom picks another name.
const OCR_NAME_CHAR_WIDTH = 7.5; // px a letter, Arial 10
const OCR_NAME_ARROW_WIDTH = 28; // px
const OCR_NAME_MIN_WIDTH = 100; // px
const OCR_NAME_MAX_WIDTH = 320; // px
function ocrNameWidth(names) {
  const longest = Math.max(OCR_NAME_COLUMN.length, ...names.map((name) => ocrText(name).length));
  return Math.min(OCR_NAME_MAX_WIDTH, Math.max(OCR_NAME_MIN_WIDTH, Math.ceil(longest * OCR_NAME_CHAR_WIDTH + OCR_NAME_ARROW_WIDTH)));
}
const ocrFitNameColumn = (sheet) => ocrFitColumn(sheet, OCR_NAME_COLUMN, ocrNameWidth);

// "Шалгах шалтгаан" as wide as its longest reason (the owner, 2026-10-28: mom reads why a line is
// pink without opening the cell), between the bounds; a reason longer than the widest is cut at
// the edge (the row stays one line tall: the photos lie over rows of the default height).
const OCR_REASON_CHAR_WIDTH = 7; // px a letter, Arial 10 (mostly lower-case Cyrillic)
const OCR_REASON_PADDING = 8; // px
const OCR_REASON_MIN_WIDTH = 160; // px
const OCR_REASON_MAX_WIDTH = 700; // px, about 100 letters: beside the photo on mom's screen (the owner, 2026-10-29; 1200 until then)
function ocrReasonWidth(reasons) {
  const longest = Math.max(OCR_REASON_COLUMN.length, ...reasons.map((reason) => ocrText(reason).length));
  return Math.min(OCR_REASON_MAX_WIDTH, Math.max(OCR_REASON_MIN_WIDTH, Math.ceil(longest * OCR_REASON_CHAR_WIDTH + OCR_REASON_PADDING)));
}
const ocrFitReasonColumn = (sheet) => ocrFitColumn(sheet, OCR_REASON_COLUMN, ocrReasonWidth);

// Both fitted columns at once: when lines come in, when the tab is set up and when mom marks a line.
function ocrFitColumns(sheet) {
  ocrFitNameColumn(sheet);
  ocrFitReasonColumn(sheet);
}

// A column as wide as widthOf(its texts on the lines) says, only when it is not so already.
function ocrFitColumn(sheet, title, widthOf) {
  const column = ocrColumn(title);
  const last = ocrLastRecordRow(sheet);
  const texts = last >= 2 ? sheet.getRange(2, column, last - 1, 1).getValues().map(([text]) => text) : [];
  const width = widthOf(texts);
  if (sheet.getColumnWidth(column) !== width) {
    sheet.setColumnWidth(column, width);
  }
}

// The photos in M, as wide as they are; those of before, in N or O, moved there (the owner,
// 2026-10-28; 2026-10-25).
function ocrPlacePhotos(sheet) {
  sheet.getImages()
    .filter((image) => OCR_OLD_IMAGE_COLUMNS.includes(image.getAnchorCell().getColumn()))
    .forEach((image) => image.setAnchorCell(sheet.getRange(image.getAnchorCell().getRow(), OCR_IMAGE_COLUMN)));
  sheet.setColumnWidth(OCR_IMAGE_COLUMN, OCR_IMAGE_COLUMN_WIDTH);
}

// A tab of before 2026-10-26: the last row of a receipt was as tall as its photo. Every row is made
// as tall as any, and the receipts move down to make room for their photos (blank rows below a
// short one), their notes and photos with them (the owner, 2026-10-26). Once: a tab whose rows are
// all of the default height is left as it is.
function ocrSpreadPhotos(sheet) {
  const last = ocrLastRecordRow(sheet);
  if (last < 2) {
    return;
  }
  const isTall = Array.from({ length: last - 1 }, (_, index) => sheet.getRowHeight(index + 2)).some((height) => height > OCR_ROW_HEIGHT);
  if (!isTall) {
    return;
  }
  const block = sheet.getRange(2, 1, last - 1, OCR_HEADER.length);
  const [values, notes] = [block.getValues(), block.getNotes()];
  const photos = new Map(sheet.getImages()
    .filter((image) => image.getAnchorCell().getColumn() === OCR_IMAGE_COLUMN)
    .map((image) => [image.getAnchorCell().getRow(), image]));
  // a receipt: from row 2 or a photo's row to the next photo's row
  const starts = [...new Set([2, ...[...photos.keys()].filter((row) => row <= last)])].sort((a, b) => a - b);
  const moves = starts.map((start, index) => {
    const end = (starts[index + 1] || last + 1) - 1;
    return { start, lines: end - start + 1, image: photos.get(start) };
  });
  const placed = moves.reduce((acc, move) => [...acc, { ...move, to: acc.length === 0 ? 2 : acc[acc.length - 1].next }].map((item, index, all) => (index === all.length - 1
    ? { ...item, next: item.to + Math.max(item.lines, item.image ? ocrPhotoRows(item.image) : 0) } : item)), []);
  const height = Math.max(last, placed[placed.length - 1].next - 1) - 1;
  const empty = () => Array.from({ length: height }, () => Array(OCR_HEADER.length).fill(''));
  const [newValues, newNotes] = [empty(), empty()];
  placed.forEach(({ start, lines, to }) => {
    for (let offset = 0; offset < lines; offset++) {
      newValues[to - 2 + offset] = values[start - 2 + offset];
      newNotes[to - 2 + offset] = notes[start - 2 + offset];
    }
  });
  ocrEnsureRows(sheet, height + 1);
  // Sheets checks a value written as if typed: a name the list has no more (or never had) would be
  // refused in "Бараа". The lists are taken off while the lines move, then put back the same.
  const range = sheet.getRange(2, 1, height, OCR_HEADER.length);
  const rules = range.getDataValidations();
  range.clearDataValidations();
  range.setValues(newValues).setNotes(newNotes).setVerticalAlignment('top');
  range.setDataValidations(rules);
  sheet.setRowHeights(2, height, OCR_ROW_HEIGHT);
  placed.filter(({ image }) => image).forEach(({ image, to }) => image.setAnchorCell(sheet.getRange(to, OCR_IMAGE_COLUMN)));
  Logger.log(`  ${sheet.getName()}: мөр бүр ердийн өндөртэй боллоо, ${placed.length} баримт зурагтаа багтав`);
}

// The last row of the receipts: the last with something in their columns (the legend beside the
// first rows is not one). 1 (the header) when there is none.
function ocrLastRecordRow(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) {
    return last;
  }
  const rows = sheet.getRange(1, 1, last, OCR_HEADER.length).getValues();
  const filled = rows.map((row) => row.some((value) => value !== '' && value !== null));
  return filled.lastIndexOf(true) + 1;
}

// Mom's columns and the view choice lose the purple they had until 2026-10-26 (the owner: no colour
// but the rows' own, ocrRowRules).
function ocrClearMomColor(sheet) {
  [ocrViewCell(sheet), ...ocrMomRanges(sheet, 1)].forEach((range) => range.setBackground(null));
}

// The lists to choose from ("Бараа": the product list, "Ээж: Зөв/Буруу": ocrMarkChoices) on the
// lines of the receipts only: a blank row under a photo, or below the last line, has none (the
// owner, 2026-10-26). Made again for the whole tab, or for some new lines (first, count).
function ocrSetLineChoices(sheet, first = null, count = 0) {
  const names = sheet.getParent().getSheetByName(OCR_NAMES_SHEET);
  const rules = [
    [OCR_NAME_COLUMN, names ? SpreadsheetApp.newDataValidation().requireValueInRange(names.getRange(1, 1, names.getMaxRows(), 1), true).setAllowInvalid(false).build() : null],
    [OCR_MOM_MARKS[0], SpreadsheetApp.newDataValidation().requireValueInList(ocrMarkChoices(sheet), true).setAllowInvalid(false).build()],
  ];
  const rows = first === null ? ocrReadRecords(sheet).map((record) => record._row) : Array.from({ length: count }, (_, index) => first + index);
  if (first === null && sheet.getMaxRows() > 1) {
    rules.forEach(([title]) => ocrColumnsRange(sheet, [title], 2).clearDataValidations());
  }
  ocrRuns(rows).forEach(([start, length]) => rules
    .filter(([, rule]) => rule)
    .forEach(([title, rule]) => sheet.getRange(start, ocrColumn(title), length, 1).setDataValidation(rule)));
}

// Mom's marks: Зөв, Буруу, and Бэлэг in the two shops' sales spreadsheets (the owner, 2026-10-28).
function ocrMarkChoices(sheet) {
  const month = ocrMonthBookOf(sheet.getParent().getName());
  const isShopSales = Boolean(month) && month.kind === OCR_BOOK_KINDS.zarlaga && month.place !== OCR_TOSGON.place;
  return isShopSales ? [OCR_VERDICT_OK, OCR_VERDICT_WRONG, OCR_VERDICT_GIFT] : [OCR_VERDICT_OK, OCR_VERDICT_WRONG];
}

// Row numbers (ascending) as runs of neighbours: [[first, how many], ...].
const ocrRuns = (rows) => rows.reduce((runs, row) => {
  const last = runs[runs.length - 1];
  return last && last[0] + last[1] === row ? [...runs.slice(0, -1), [last[0], last[1] + 1]] : [...runs, [row, 1]];
}, []);

// The header columns a view does not show, as runs [first column, how many]; "Бүгд" shows all.
function ocrHiddenRuns(view) {
  return OCR_HEADER.reduce((runs, title, index) => {
    if (view === OCR_VIEW_ALL || OCR_VISIBLE_COLUMNS.includes(title)) {
      return runs;
    }
    const last = runs[runs.length - 1];
    return last && last[0] + last[1] === index + 1 ? [...runs.slice(0, -1), [last[0], last[1] + 1]] : [...runs, [index + 1, 1]];
  }, []);
}

function ocrPrepareSheet(sheet) {
  const view = ocrViewOf(ocrViewCell(sheet).getValue());
  sheet.getRange(1, 1, 1, OCR_HEADER.length).setValues([ocrHeaderRow(view)]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  ocrClearMomColor(sheet);
  ocrSetViewChoices(sheet);
  sheet.setConditionalFormatRules(ocrRowRules(sheet));
  ocrFormatColumns(sheet);
  ocrPlacePhotos(sheet);
  ocrSetLineChoices(sheet);
  ocrFitColumns(sheet);
  ocrWriteLegend(sheet);
  ocrWritePriceColumns(sheet);
  ocrApplyView(sheet, view);
}

// The number formats of the columns that hold numbers (a column moved by a new layout keeps the
// format of the one that stood there before, e.g. a date). Only when not set yet.
const OCR_NUMBER_FORMATS = [[[OCR_QTY_COLUMN], '0.##'], [['Тоо (уншсан)'], '0.##']];
function ocrFormatColumns(sheet) {
  if (sheet.getRange(2, ocrColumn('Тоо (уншсан)')).getNumberFormat() === '0.##') {
    return;
  }
  OCR_NUMBER_FORMATS.forEach(([titles, format]) => ocrColumnsRange(sheet, titles, 2).setNumberFormat(format));
}

// Shows the columns of a view (OCR_VIEWS), hides the others.
function ocrApplyView(sheet, view) {
  sheet.showColumns(1, OCR_HEADER.length); // also what an older layout hid
  ocrHiddenRuns(view).forEach(([first, count]) => sheet.hideColumns(first, count));
}

// The On edit trigger of a month spreadsheet: a new choice in A1 of a day tab shows its view at
// once; mom's Зөв/Буруу on a line writes the tab into the workbooks at once (ocr-instant.gs), and
// so does "Бүх мөр Зөв" in A1 after giving every line she has not looked at its Зөв.
function onOcrViewEdit(event) {
  const range = event && event.range;
  if (!range || !OCR_DAY_FOLDER.test(range.getSheet().getName())) {
    return;
  }
  if (range.getRow() === 1 && range.getColumn() === ocrColumn('Харах') && ocrText(range.getValue()) === OCR_MARK_DAY_OK) {
    ocrMarkDayOk(range.getSheet(), ocrViewOf(event.oldValue));
  } else if (range.getRow() === 1 && range.getColumn() === ocrColumn('Харах')) {
    ocrApplyView(range.getSheet(), ocrViewOf(range.getValue()));
  } else if (ocrIsMarkEdit(range)) {
    ocrFitColumns(range.getSheet());
    ocrWriteTabNow(range.getSheet());
  }
}

// A month spreadsheet gets its view trigger once; only the newest OCR_VIEW_TRIGGERS_MAX keep one.
function ocrEnsureViewTrigger(book) {
  const views = ScriptApp.getProjectTriggers().filter((trigger) => trigger.getHandlerFunction() === OCR_VIEW_HANDLER);
  if (views.some((trigger) => trigger.getTriggerSourceId() === book.getId())) {
    return;
  }
  views.slice(0, Math.max(0, views.length - OCR_VIEW_TRIGGERS_MAX + 1)).forEach((trigger) => ScriptApp.deleteTrigger(trigger));
  ScriptApp.newTrigger(OCR_VIEW_HANDLER).forSpreadsheet(book).onEdit().create();
}

// The rows of a day tab as records, by the tab's own titles: a tab of the version before is read
// right too (ocrUpgradeRecord), before or without being moved into the new layout. A blank row
// (room for a photo, ocrNextFreeRow) is none; each record keeps its row number in _row (ocrRowOf).
function ocrReadRecords(sheet) {
  const last = ocrLastRecordRow(sheet);
  const width = sheet.getLastColumn();
  if (last < 2 || width === 0) {
    return [];
  }
  const [titles, ...rows] = sheet.getRange(1, 1, last, width).getValues();
  const columns = titles.map((title, index) => [ocrText(title), index]).filter(([title]) => title !== '' && title !== OCR_LEGEND_TITLE && !OCR_PRICE_TITLES.includes(title));
  const blank = Object.fromEntries(OCR_HEADER.map((title) => [title, '']));
  return rows
    .map((row, index) => ({ row, number: index + 2 }))
    .filter(({ row }) => columns.some(([, index]) => row[index] !== '' && row[index] !== null))
    .map(({ row, number }) => ({ ...ocrUpgradeRecord({ ...blank, ...Object.fromEntries(columns.map(([title, index]) => [title, row[index]])) }), _row: number }));
}

// The row of a record read by ocrReadRecords (by its place in the list for one made otherwise).
const ocrRowOf = (record, index) => record._row || index + 2;

// A tab of the version before (other titles than OCR_HEADER): its rows are written again in the
// new order, its photos moved beside the new columns, its old protections removed (the review
// makes them anew) and it is prepared like a new tab. true: it was moved.
function ocrUpgradeTab(sheet) {
  const width = sheet.getLastColumn();
  const titles = width > 0 ? sheet.getRange(1, 1, 1, width).getValues()[0].map(ocrText) : [];
  if (titles.length === 0 || ocrIsCurrentLayout(titles)) {
    return false;
  }
  if (DRY_RUN) {
    Logger.log(`  ${sheet.getName()}: шинэ бүтэц рүү шилжүүлнэ (туршилт)`);
    return false;
  }
  const records = ocrRepairInstantWrites(sheet, titles, ocrReadRecords(sheet));
  const oldImageColumn = titles.length + 2;
  // only the old titles' columns are emptied: the legend and the prices right of them stay
  const beyond = titles.findIndex((title) => title === '' || title === OCR_LEGEND_TITLE || OCR_PRICE_TITLES.includes(title));
  const area = sheet.getRange(1, 1, sheet.getMaxRows(), Math.max(beyond === -1 ? width : beyond, OCR_HEADER.length));
  area.clearContent();
  area.clearDataValidations();
  area.setBackground(null);
  // every line at its own row again: its photo and prices stay beside it (blank rows under a photo stay blank)
  ocrRuns(records.map((record) => record._row)).forEach(([first, count]) => {
    const rows = records.filter((record) => record._row >= first && record._row < first + count).map(ocrRow);
    sheet.getRange(first, 1, count, OCR_HEADER.length).setValues(rows);
  });
  sheet.getImages()
    .filter((image) => [oldImageColumn, ...OCR_OLD_IMAGE_COLUMNS].includes(image.getAnchorCell().getColumn()))
    .forEach((image) => image.setAnchorCell(sheet.getRange(image.getAnchorCell().getRow(), OCR_IMAGE_COLUMN)));
  [...sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET), ...sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)]
    .filter((protection) => protection.getDescription().startsWith(OCR_GUARD_PREFIX))
    .forEach((protection) => protection.remove());
  ocrPrepareSheet(sheet);
  Logger.log(`  ${sheet.getName()}: шинэ бүтэц рүү шилжлээ (${records.length} мөр)`);
  return true;
}

// The lines of a tab of the layout before 2026-10-28 (titles: its own) with the damage of mom's
// instant marks put right (ocrRepairInstantWrite): only where "Бичсэн"'s new column holds the old
// reason's title and the new reason column stands beside the old titles.
function ocrRepairInstantWrites(sheet, titles, records) {
  const reasonAt = ocrColumn(OCR_REASON_COLUMN);
  const isBefore = ocrText(titles[ocrColumn('Бичсэн') - 1]) === OCR_REASON_COLUMN && ocrText(titles[reasonAt - 1]) === '';
  if (!isBefore || records.length === 0) {
    return records;
  }
  const last = Math.max(...records.map((record) => record._row));
  const beside = sheet.getLastColumn() >= reasonAt ? sheet.getRange(1, reasonAt, last, 1).getValues() : [];
  const repaired = records.map((record) => ocrRepairInstantWrite(record, (beside[record._row - 1] || [''])[0]));
  const count = repaired.filter((record, index) => record !== records[index]).length;
  if (count > 0) {
    Logger.log(`  ${sheet.getName()}: ээжийн тэмдгээр буруу баганад бичигдсэн ${count} мөрийг засна`);
  }
  return repaired;
}

// The file id in a Drive link ("…/file/d/<id>/view", "…?id=<id>"); the link itself otherwise.
function ocrFileIdFromUrl(url) {
  const match = String(url).match(/\/d\/([\w-]+)|[?&]id=([\w-]+)/);
  return match ? match[1] || match[2] : String(url);
}

// IDs for rows written before there was an ID column: the photo of the link, line by line.
function ocrIdsFromLinks(links) {
  const lines = {};
  return links.map((link) => {
    const id = ocrFileIdFromUrl(link);
    lines[id] = (lines[id] || 0) + 1;
    return [ocrRowId(id, lines[id] - 1)];
  });
}

// A day spreadsheet of the first version (next to the photos, maybe with the old checkbox column):
// its rows, by their titles, go into the day's tab of the month spreadsheet, each photo beside its
// rows again, and the old file goes to the trash. A row already there (same ID) is not added again,
// so a run cut off half-way does no harm. Returns the month spreadsheet { file, place, month }.
function ocrMigrateDaySheet(root, entry) {
  const old = SpreadsheetApp.openById(entry.file.getId()).getSheets()[0];
  const iso = entry.file.getName().slice(OCR_DAY_SPREADSHEET_PREFIX.length);
  const [place, kind, month, day] = [entry.path[0], ocrBookKind(entry.path), ocrMonthFolderName(iso), iso.slice(8)];
  const last = old.getLastRow();
  const titles = last > 0 ? old.getRange(1, 1, 1, old.getLastColumn()).getValues()[0].map(ocrText) : [];
  const records = last > 1 ? old.getRange(2, 1, last - 1, titles.length).getValues().map((row) => Object.fromEntries(titles.map((title, index) => [title, row[index]]))) : [];
  const ids = ocrIdsFromLinks(records.map((record) => record['Холбоос']));
  const rows = records.map((record, index) => ocrRow({ ...ocrUpgradeRecord(record), 'Огноо': iso, 'ID': ocrText(record['ID']) || ids[index][0] }));
  const { book, file } = ocrMonthBook(root, place, kind, month);
  const tab = ocrDayTab(book, day);
  const idColumn = OCR_HEADER.indexOf('ID');
  const lastRecord = ocrLastRecordRow(tab);
  const known = lastRecord > 1 ? tab.getRange(2, idColumn + 1, lastRecord - 1, 1).getValues().map(([id]) => ocrText(id)) : [];
  ocrGroupByPhoto(rows.filter((row) => !known.includes(row[idColumn]))).forEach((group) => {
    const first = ocrAppendRows(tab, group.rows);
    try {
      ocrAddPhoto(tab, first, group.rows.length, DriveApp.getFileById(group.photoId));
    } catch (error) {
      Logger.log(`  ${group.photoId}: зураг олдсонгүй (${error.message})`);
    }
  });
  entry.file.setTrashed(true);
  Logger.log(`${entry.path.join('/')}/${entry.file.getName()}: ${rows.length} мөр "${ocrMonthBookName(place, kind, month)} / ${day}" руу шилжив, хуучин файл хогийн саванд`);
  return { file, place, kind, month };
}

// Rows next to each other with the same photo (the file id before "#" of their ID).
function ocrGroupByPhoto(rows) {
  const idColumn = OCR_HEADER.indexOf('ID');
  return rows.reduce((groups, row) => {
    const photoId = ocrText(row[idColumn]).split('#')[0];
    const lastGroup = groups[groups.length - 1];
    return lastGroup && lastGroup.photoId === photoId
      ? [...groups.slice(0, -1), { photoId, rows: [...lastGroup.rows, row] }]
      : [...groups, { photoId, rows: [row] }];
  }, []);
}

function ocrOnlyOwner(protection, owner) {
  protection.addEditor(owner);
  protection.removeEditors(protection.getEditors().filter((user) => user.getEmail() !== owner));
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
}

// Mom's own accounts, when she has some besides the owner account: the "Ээж" addresses of
// OCR_REVIEWERS, joined by commas ('' when mom uses the owner account). Only then can the sheets keep
// her out of what was read: Google always lets the owner account change everything.
function ocrMomEmail(props) {
  try {
    const value = JSON.parse(props.getProperty(OCR_REVIEWERS_PROPERTY) || '{}')[OCR_MOM_KEY] || '';
    const emails = (Array.isArray(value) ? value : String(value).split(',')).map((email) => String(email).trim()).filter(ocrIsEmail);
    return [...new Set(emails)].join(',');
  } catch (error) {
    return '';
  }
}

// Mom's addresses one by one (ocrMomEmail gives them joined).
const ocrMomList = (mom) => (mom ? mom.split(',') : []);

// Who changes what, in every day tab: the owner everything; mom (her own account) her columns -
// and nothing that was read. A tab is done again when her address changed (it is in the
// protection's description) or the tab was moved into a new layout (its protection was removed).
// Mom's accounts are made the file's editors once a tab's protection was made (ocrShareWith).
// The protection's description: mom's addresses in it make a new address redo it.
const ocrGuardDescription = (mom) => `${OCR_GUARD_PREFIX}: ${mom || '-'}`;

// One day tab's protection (in ocrSetUpBook); true: it was (re)made.
function ocrGuardTab(sheet, description, mom, owner) {
  const existing = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  if (existing.some((protection) => protection.getDescription() === description)) {
    return false;
  }
  existing.forEach((protection) => protection.remove());
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)
    .filter((protection) => protection.getDescription().startsWith(OCR_GUARD_PREFIX))
    .forEach((protection) => protection.remove());
  const momRanges = mom ? [ocrViewCell(sheet), ...ocrMomRanges(sheet, 2)] : [];
  const protection = sheet.protect().setDescription(description);
  protection.setUnprotectedRanges(momRanges);
  ocrOnlyOwner(protection, owner);
  // mom's columns are hers and the owner's only
  momRanges.forEach((range) => {
    const own = range.protect().setDescription(`${OCR_GUARD_PREFIX}: ээж`);
    ocrOnlyOwner(own, owner);
    ocrMomList(mom).forEach((email) => own.addEditor(email));
  });
  return true;
}

// Mom's accounts become the file's only editors besides the owner (a staff member of the version
// before, or a replaced address, loses it). An editor through a parent folder can not be removed
// here: logged.
function ocrShareWith(file, editorsWanted, owner) {
  const editors = file.getEditors().map((user) => user.getEmail());
  editors
    .filter((email) => email !== owner && !editorsWanted.includes(email))
    .forEach((email) => {
      try {
        file.removeEditor(email);
      } catch (error) {
        Logger.log(`  ${file.getName()}: ${email}-ийн эрхийг хасаж чадсангүй (${error.message})`);
      }
    });
  editorsWanted.filter((email) => !editors.includes(email)).forEach((email) => file.addEditor(email));
}

const OCR_NAME_LIST_ROOM = 300; // rows free under the names for products added to "Бараа" later

// The list mom picks from (column A of OCR_NAMES_SHEET): Хамааралгүй first, then the names the
// script wrote in C (every source, ocr-catalog.gs) and the product list "Бараа" of the workbook read
// live (IMPORTRANGE), so a product added there is in her list within seconds, not at the next run
// (the owner, 2026-10-25). A heading row of "Бараа" (no shop in B) is left out; while the import is
// not allowed yet, the list is what the script wrote.
function ocrNameListFormula(mainId) {
  const live = `IFERROR(QUERY(IMPORTRANGE("${mainId}","${OCR_PRODUCT_LIST_TAB}!A${OCR_CATALOG_FIRST_ROW}:B"),"select Col1 where Col2 is not null",0),"")`;
  // ARRAYFORMULA: inside LET a plain TRIM gives the first name only
  return `=LET(names,ARRAYFORMULA(TRIM({C1:C;${live}})),{"${OCR_NOT_RELEVANT}";UNIQUE(FILTER(names,names<>"",names<>"${OCR_NOT_RELEVANT}"))})`;
}

// The product names on the hidden tab OCR_NAMES_SHEET (ocrNameListFormula); mom's name column of
// every day tab chooses from them only (ocrSetLineChoices, in ocrSetUpBook: a name typed another way
// is refused), so a name is always written the same way. entries: [{ name, price }] (ocr-catalog.gs); mainId: the workbook of "Бараа".
function ocrSetNameList(book, entries, mainId) {
  if (entries.length === 0) {
    return;
  }
  const names = entries.map((entry) => entry.name).filter((name) => name !== OCR_NOT_RELEVANT);
  let list = book.getSheetByName(OCR_NAMES_SHEET);
  if (!list) {
    list = book.insertSheet(OCR_NAMES_SHEET, book.getSheets().length);
    list.hideSheet();
    ocrOnlyOwner(list.protect().setDescription(`${OCR_GUARD_PREFIX}: барааны нэрс`), Session.getEffectiveUser().getEmail());
  }
  const missingRows = names.length + OCR_NAME_LIST_ROOM - list.getMaxRows();
  if (missingRows > 0) {
    list.insertRowsAfter(list.getMaxRows(), missingRows);
  }
  list.clear();
  // D: the price the customer sees beside each name, for "Манай үнэ" (ocrWriteOurPrices)
  const prices = new Map(entries.map((entry) => [entry.name, typeof entry.price === 'number' ? entry.price : '']));
  list.getRange(1, 3, names.length, 2).setValues(names.map((name) => [name, prices.get(name) ?? '']));
  list.getRange(1, 1).setFormula(ocrNameListFormula(mainId));
}
