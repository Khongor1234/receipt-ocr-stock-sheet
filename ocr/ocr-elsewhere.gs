/**
 * Receipt OCR - goods mom typed on another day (2026-10-23). A receipt's day cell is empty, but in
 * the same row and block (income, expense) of the sheet mom typed the same number by hand on
 * another day (the receipt of the 16th typed on the 22nd): the receipt's day cell gets 0, its note
 * line names mom's cell ("(ээж 22-нд гараар: L95)"), so the goods are not counted twice and her
 * number stays where she put it. One hand-typed number answers one receipt line only; the nearest
 * day is taken. Once mom's number is gone (or another), the receipt's number is written again.
 */

const OCR_ELSEWHERE_LINE = / \(ээж (\d+)-нд гараар бичсэн: ([A-Z]+\d+), энд 0\)$/;

// The target with .elsewhere ({ day, cell }) when mom typed its goods on another day; else as it is.
function ocrWithElsewhere(target, decision, id, workbooks) {
  if (!target.ok || target.isCheck || !target.isDay || !decision.post || !(decision.qty > 0)) {
    return target;
  }
  const elsewhere = ocrHandElsewhere(target, decision.qty, id, workbooks);
  return elsewhere ? { ...target, elsewhere } : target;
}

// Mom's hand-typed cell of the same number on another day of the target's row and block, not taken
// by another receipt line (its note, or earlier in this run); the nearest day, the earlier one on a tie.
function ocrHandElsewhere(target, qty, id, workbooks) {
  const layout = target.place === OCR_TOSGON.place ? OCR_TOSGON.layout : LAYOUT;
  const first = target.column >= layout.expenseFirst ? layout.expenseFirst : layout.incomeFirst;
  const grid = ocrTabGrid(workbooks, ocrBookIdFor(target.place, target.tabName), target.tabName);
  const index = target.row - OCR_BOOK_FIRST_ROW;
  if (!grid || !grid.values[index]) {
    return null;
  }
  const taken = ocrTakenElsewhere(grid.notes[index], id);
  const candidates = Array.from({ length: layout.days }, (_, offset) => first + offset)
    .filter((column) => column !== target.column)
    .map((column) => ({ column, cell: `${columnLetter(column)}${target.row}` }))
    .filter(({ column, cell }) => {
      const value = grid.values[index][column - 1];
      const note = grid.notes[index][column - 1];
      const color = String(grid.backgrounds[index][column - 1]).toLowerCase();
      const runTaken = workbooks.get(`elsewhere|${target.tabName}!${cell}`);
      return value === qty && ocrIsHandCell(value, note) && !OCR_CELL_COLORS.includes(color) &&
        !taken.has(cell) && (!runTaken || runTaken === id);
    })
    .sort((a, b) => Math.abs(a.column - target.column) - Math.abs(b.column - target.column) || a.column - b.column);
  if (candidates.length === 0) {
    return null;
  }
  const { column, cell } = candidates[0];
  workbooks.set(`elsewhere|${target.tabName}!${cell}`, id);
  return { day: ocrDayOfColumn(target.tabName, column - first), cell };
}

// The cells of a row other receipt lines already answer with mom's number (their note lines).
function ocrTakenElsewhere(rowNotes, id) {
  return new Set(rowNotes
    .flatMap((note) => ocrNoteEntries(note) || [])
    .filter((entry) => entry.elsewhere && entry.id !== id)
    .map((entry) => entry.elsewhere.cell));
}

// The day of a block's column offset: "Sor/2026/4/30" starts on the 16th, "Sor/2026/4/15" and
// "Tos/2026/4" on the 1st.
function ocrDayOfColumn(tabName, offset) {
  const parts = tabName.split('/');
  const isSecondHalf = parts.length > 3 && Number(parts[3]) > FIRST_HALF_LAST_DAY;
  return (isSecondHalf ? FIRST_HALF_LAST_DAY + 1 : 1) + offset;
}

// "(ээж 22-нд гараар бичсэн: L95, энд 0)" after a line in "Бичсэн"; '' for none.
const ocrElsewhereStatus = (elsewhere) => (elsewhere ? `ээж ${elsewhere.day}-нд гараар бичсэн: ${elsewhere.cell}, энд 0` : '');

// The { day, cell } a "Бичсэн" line names as mom's; null for none.
function ocrElsewhereOfLine(line) {
  const match = line.match(OCR_ELSEWHERE_LINE);
  return match ? { day: Number(match[1]), cell: match[2] } : null;
}
