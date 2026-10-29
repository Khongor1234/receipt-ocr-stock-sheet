/**
 * Receipt OCR - a line left in a workbook cell that its "Бичсэн" does not name (the owner,
 * 2026-10-26: Tos/2026/10's count of 10/23 had 17 lines in 15 cells of other products, 84 pieces,
 * after the rows above them were put in name order without the count block). Every review of a day
 * tab goes through the workbook columns its lines are written in and takes out of a cell each line
 * of the tab that is not written there by its "Бичсэн" while the cell "Бичсэн" names holds it too:
 * a second copy only, and only while that cell is the script's (not changed by hand). A line whose
 * "Бичсэн" names a cell without it (a row moved with its cell in an old month nobody reviewed since,
 * 2026-10-26: 64 such cells) stays: it is the only one. The cell
 * keeps its other lines and is their sum again, its colour theirs. A cell typed or changed by hand
 * is left as it is, and so are OCR_HANDS_OFF_CELLS. The check copy is not gone through.
 */

function ocrSweepOrphans(records, workbooks, where) {
  const homes = new Map(records
    .filter((record) => ocrText(record['ID']) !== '')
    .map((record) => [ocrText(record['ID']), new Set(ocrWrittenCells(record['Бичсэн']))]));
  const isAtHome = ocrLineAtHome(homes, workbooks);
  return [...ocrWrittenColumns(records).values()]
    .reduce((removed, column) => removed + ocrSweepColumn(column, homes, isAtHome, workbooks, where), 0);
}

// Whether a line is in a cell its "Бичсэн" names, read as the cells are now (once a line).
function ocrLineAtHome(homes, workbooks) {
  const known = new Map();
  return (id) => {
    if (!known.has(id)) {
      known.set(id, [...homes.get(id)].filter((cell) => !cell.startsWith(OCR_CHECK_PREFIX)).some((cell) => {
        const range = ocrHomeRange(cell, workbooks);
        if (range === null) {
          return false;
        }
        const [value, note] = [range.getValue(), range.getNote()];
        // a cell changed by hand holds no sure line (2026-10-26: My/2026/9/30!F101 emptied, its note left)
        return !ocrIsHandCell(value, note) && (ocrNoteEntries(note) || []).some((entry) => entry.id === id);
      }));
    }
    return known.get(id);
  };
}

// "Tos/2026/10!BY126" -> its range; null when the workbook or tab is not there.
function ocrHomeRange(cell, workbooks) {
  const [, tabName, letters, row] = cell.match(/^(.+)!([A-Z]+)(\d+)$/);
  const place = ocrPlaceOfTab(tabName);
  const tab = place && OCR_POST_BOOKS[place] ? ocrWorkbookTab(workbooks, ocrBookIdFor(place, tabName), tabName) : { error: true };
  return tab.error ? null : tab.sheet.getRange(Number(row), ocrColumnNumber(letters));
}

// The workbook columns the lines are written in: "<book>|<tab>|<letters>" -> { bookId, tabName, column }.
function ocrWrittenColumns(records) {
  const columns = new Map();
  records
    .flatMap((record) => ocrWrittenCells(record['Бичсэн']))
    .filter((cell) => !cell.startsWith(OCR_CHECK_PREFIX))
    .forEach((cell) => {
      const [, tabName, letters] = cell.match(/^(.+)!([A-Z]+)(\d+)$/);
      const place = ocrPlaceOfTab(tabName);
      if (place && OCR_POST_BOOKS[place]) {
        const bookId = ocrBookIdFor(place, tabName);
        columns.set(`${bookId}|${tabName}|${letters}`, { bookId, tabName, column: ocrColumnNumber(letters) });
      }
    });
  return columns;
}

// One column: the cells with a line of the tab its "Бичсэн" does not name; the number of lines taken out.
function ocrSweepColumn({ bookId, tabName, column }, homes, isAtHome, workbooks, where) {
  const grid = ocrTabGrid(workbooks, bookId, tabName);
  if (!grid) {
    return 0;
  }
  const isElsewhere = (cell) => (entry) => homes.has(entry.id) && !homes.get(entry.id).has(cell);
  const isStray = (cell) => (entry) => isElsewhere(cell)(entry) && isAtHome(entry.id);
  return grid.notes.reduce((removed, row, index) => {
    const cell = `${tabName}!${columnLetter(column)}${index + OCR_BOOK_FIRST_ROW}`;
    const read = ocrNoteEntries(row[column - 1]);
    if (!read || !read.some(isElsewhere(cell)) || OCR_HANDS_OFF_CELLS.includes(cell) || !read.some(isStray(cell))) {
      return removed;
    }
    // the grid was read at the start of the run: the cell as it is now decides
    const range = ocrWorkbookTab(workbooks, bookId, tabName).sheet.getRange(index + OCR_BOOK_FIRST_ROW, column);
    const [value, note] = [range.getValue(), range.getNote()];
    const entries = ocrNoteEntries(note);
    const strays = entries ? entries.filter(isStray(cell)) : [];
    if (strays.length === 0 || ocrIsHandCell(value, note)) {
      return removed;
    }
    const next = entries.filter((entry) => !strays.includes(entry));
    ocrApplyChange(range, {
      action: 'write',
      value: next.length > 0 ? ocrNoteSum(next) : '',
      note: next.map(ocrNoteLine).join('\n'),
      color: ocrCellColor(next),
      entries: next,
    });
    Logger.log(`  ${where}: ${cell}-ээс өөр нүдэнд бичигдсэн мөрийг хаслаа (${strays.map((entry) => `${entry.id} = ${entry.qty}`).join(', ')})${DRY_RUN ? ' (туршилт)' : ''}`);
    return removed + strays.length;
  }, 0);
}
