/**
 * Receipt OCR - the check copy of a workbook. A cell mom had typed by hand is never changed; a
 * receipt line for it goes into the same cell of the workbook's full copy, "<name> (check)", made
 * once in the same folder (its id kept in Script properties), so the two files can be compared
 * side by side:
 *   - the copy's cell holds the receipts' lines only (what was copied from the original is
 *     replaced), yellow, green once mom approved them all
 *   - the original hand-typed cell keeps mom's number: light blue when the receipts make it,
 *     light red when they make another (a "Баримтаар: n" line added to its note), no colour and
 *     no such line when no receipt is for it
 * The copy's row of a product is found by its name, as rows may be added to the original later.
 */

// The copy of a workbook, made once; null in DRY_RUN before it exists.
function ocrCheckBookId(bookId, workbooks) {
  const key = `check|${bookId}`;
  if (workbooks.has(key)) {
    return workbooks.get(key);
  }
  const props = PropertiesService.getScriptProperties();
  const known = props.getProperty(OCR_CHECK_COPY_PROPERTY + bookId);
  let id = known && ocrFileIsThere(known) ? known : null;
  if (!id && !DRY_RUN) {
    const original = DriveApp.getFileById(bookId);
    const parents = original.getParents();
    const name = original.getName() + OCR_CHECK_SUFFIX;
    const copy = parents.hasNext() ? original.makeCopy(name, parents.next()) : original.makeCopy(name);
    id = copy.getId();
    props.setProperty(OCR_CHECK_COPY_PROPERTY + bookId, id);
    Logger.log(`"${name}" үүсгэлээ: гараар бичсэн нүдний баримтыг тэнд бичнэ.`);
  }
  workbooks.set(key, id);
  return id;
}

function ocrFileIsThere(id) {
  try {
    return !DriveApp.getFileById(id).isTrashed();
  } catch (error) {
    return false;
  }
}

const ocrIsHandTarget = (target) => ocrIsHandCell(target.range.getValue(), target.range.getNote());

// The cell of the check copy for a hand-typed workbook cell (ocrCellAt's target):
// { ...target, cell: 'check:…', range (the copy's), isCheck, hand: { range, value } }.
function ocrCheckTarget(target, workbooks) {
  const copyId = ocrCheckBookId(ocrBookIdFor(target.place, target.tabName), workbooks);
  if (!copyId) {
    return { ok: false, reason: `${target.cell}: гараар бичсэн нүд, check файл хараахан үүсээгүй (туршилт)` };
  }
  const copy = ocrWorkbookTab(workbooks, copyId, target.tabName);
  if (copy.error) {
    return { ok: false, reason: `check: ${copy.error}` };
  }
  const original = ocrWorkbookTab(workbooks, ocrBookIdFor(target.place, target.tabName), target.tabName);
  const name = ocrText(original.sheet.getRange(target.row, 1).getValue());
  const row = copy.rows.get(ocrNameKey(name));
  if (!row) {
    return { ok: false, reason: `check: "${name}" ${target.tabName} дээр алга` };
  }
  return {
    ...target,
    cell: `${OCR_CHECK_PREFIX}${target.tabName}!${columnLetter(target.column)}${row}`,
    range: copy.sheet.getRange(row, target.column),
    isCheck: true,
    hand: { range: target.range, value: target.range.getValue() },
  };
}

// A cell of the check copy named in "Бичсэн", with its original cell when the product is still there.
function ocrFindCheckCell(cell, iso, workbooks) {
  const [, tabName, letters, row] = cell.slice(OCR_CHECK_PREFIX.length).match(/^(.+)!([A-Z]+)(\d+)$/);
  const place = ocrPlaceOfTab(tabName);
  const copyId = place && OCR_POST_BOOKS[place] ? ocrCheckBookId(ocrBookIdFor(place, tabName), workbooks) : null;
  const copy = copyId ? ocrWorkbookTab(workbooks, copyId, tabName) : { error: `${cell}: check файл олдсонгүй` };
  if (copy.error) {
    return { ok: false, reason: copy.error };
  }
  const column = ocrColumnNumber(letters);
  const layout = place === OCR_TOSGON.place ? OCR_TOSGON.layout : LAYOUT;
  const name = ocrText(copy.sheet.getRange(Number(row), 1).getValue());
  const original = ocrWorkbookTab(workbooks, ocrBookIdFor(place, tabName), tabName);
  const originalRow = original.error ? null : original.rows.get(ocrNameKey(name));
  const range = originalRow ? original.sheet.getRange(originalRow, column) : null;
  return {
    ok: true,
    cell,
    range: copy.sheet.getRange(Number(row), column),
    isCount: column === layout.count,
    isCheck: true,
    hand: range ? { range, value: range.getValue() } : null,
  };
}

// A cell as it stands, as a change would leave it: { value, entries }.
function ocrCellAfter(target) {
  return { value: target.range.getValue(), entries: ocrNoteEntries(target.range.getNote()) || [] };
}

// The original hand-typed cell: light blue when the receipts make its number, light red with the
// receipts' number in its note when they make another, plain when no receipt is for it.
function ocrMarkHandCell(target, after) {
  if (DRY_RUN || !target.hand) {
    return;
  }
  const range = target.hand.range;
  const note = range.getNote();
  const look = ocrHandLook(target.hand.value, note, after.value);
  if (String(range.getBackground()).toLowerCase() !== look.color) {
    ocrSetCellColor(range, look.color);
  }
  if (look.note !== ocrText(note)) {
    range.setNote(look.note);
  }
}

// "(гараар бичсэнтэй таарсан)" or "(гараар 2)" after a check cell's line in "Бичсэн".
const ocrHandStatus = (target, after) => (target.hand ? ocrHandLook(target.hand.value, '', after.value).status : '');

// What a hand-typed cell shows for the receipts' number of its check cell ('': no receipt):
// { color, note, status }, status being what follows its line in "Бичсэн".
function ocrHandLook(handValue, note, receiptValue) {
  const hasReceipt = ocrText(receiptValue) !== '';
  const isMatch = hasReceipt && receiptValue === handValue;
  return {
    color: !hasReceipt ? null : isMatch ? OCR_HAND_MATCH_COLOR : OCR_HAND_DIFF_COLOR,
    note: ocrHandNote(note, hasReceipt && !isMatch ? receiptValue : null),
    status: isMatch ? OCR_HAND_MATCH_TEXT : `гараар ${ocrText(handValue) === '' ? 'хоосон' : handValue}`,
  };
}

// A hand-typed cell's note with its "Баримтаар: n" line put in, changed, or taken out (null); mom's
// own lines stay, so the cell still counts as hand-typed.
function ocrHandNote(note, receiptValue) {
  const own = ocrText(note).split('\n').filter((line) => line.trim() !== '' && !line.startsWith(OCR_HAND_NOTE_MARK));
  const line = receiptValue === null ? [] : [`${OCR_HAND_NOTE_MARK} ${receiptValue}`];
  return [...own, ...line].join('\n');
}

// Whether every hand-typed cell of a written line already shows what its check cell holds (colour,
// note, the status in "Бичсэн"), and a cell written by the line still has mom's number on another
// day as it had (ocrHandElsewhere). Read from each tab once per run (ocrTabGrid), so a line is gone
// through again only when mom changed her cells or the receipts changed. found: the line's original
// cells, in the order of its "Бичсэн" lines.
function ocrHandLinesCurrent(record, found, workbooks, decision) {
  return ocrWrittenLines(record['Бичсэн']).every((line, index) => {
    if (!line.startsWith(OCR_CHECK_PREFIX)) {
      const wanted = ocrWithElsewhere(found[index], decision, record['ID'], workbooks).elsewhere || null;
      return JSON.stringify(wanted) === JSON.stringify(ocrElsewhereOfLine(line));
    }
    const target = found[index];
    const bookId = ocrBookIdFor(target.place, target.tabName);
    const copyId = ocrCheckBookId(bookId, workbooks);
    const original = ocrTabGrid(workbooks, bookId, target.tabName);
    const copy = copyId ? ocrTabGrid(workbooks, copyId, target.tabName) : null;
    const copyRow = Number(ocrParseWrittenLine(line).cell.replace(/^[A-Z]+/, ''));
    const at = (grid, part, row) => ((grid && grid[part][row - OCR_BOOK_FIRST_ROW]) || [])[target.column - 1];
    if (!original || !copy || at(copy, 'values', copyRow) === undefined) {
      return false;
    }
    const hand = at(original, 'values', target.row);
    const note = at(original, 'notes', target.row);
    const color = String(at(original, 'backgrounds', target.row)).toLowerCase();
    const look = ocrHandLook(hand, note, at(copy, 'values', copyRow));
    return ocrIsHandCell(hand, note) &&
      (look.color ? color === look.color : !OCR_CELL_COLORS.includes(color)) &&
      look.note === ocrText(note) &&
      line.endsWith(` (${look.status})`);
  });
}
