/**
 * The transfers between the two shops (the owner's rule, 2026-10-22): what one shop gives the
 * other is written once and stands on both sheets of that half-month. One shop's transfer out
 * (AK) is the other shop's transfer in (U); goods for Тосгон never go in AK.
 *   - A transfer out typed in one shop is copied to the other shop's transfer in, pink
 *     (TRANSFER_AUTO_COLOR), with a note saying where it came from. A changed or cleared number
 *     follows. Not the other way: a transfer in may hold goods from Тосгон (seen in the 2026
 *     sheets), so a U with no AK on the other side is left as it is.
 *   - Two typed numbers that differ are both left as they are and turn a darker pink
 *     (TRANSFER_CONFLICT_COLOR) with a note giving both: mom decides which is right.
 *   - A product short in one shop's count and over in the other's (same half-month) may have been
 *     moved without being written: both cells turn lavender (TRANSFER_HINT_COLOR) with a note, and
 *     no number is written (the owner, 2026-10-24: the counts were made, only not typed in, so the
 *     script takes nothing off). Only from 2026/10/31 on (TRANSFER_HINT_FROM). It is worked out
 *     again on every run, so it goes once the counts or a typed transfer agree.
 *   - Until 2026-10-24 such a product got a pink transfer added ("тооллогоор тааруулсан +n"), and
 *     a negative balance too ("хасах үлдэгдлээр тааруулсан +n"). Every run takes such an added part
 *     out again; removeCountMatches() takes them out of all the half-months at once. The half-months
 *     before 2026/4/15 keep theirs (TRANSFER_MATCHES_KEPT_UNTIL, the owner, 2026-10-24): a product
 *     with one there is left as it is.
 * A cell the receipt OCR wrote (its note has "OCR:", or since 2026-10-25 the photo's link
 * "Зураг: https://drive.google.com/file/d/...") is never changed; it counts as typed.
 * A number typed over a copy counts as typed (the copy's note keeps the copied number, and
 * onSheetEdit drops the copy mark at once: forgetTransferMarks).
 * A transfer out to a product the other shop does not list is marked too, on the recent
 * half-months only: the old sheets have many names of the time, so there they are only listed.
 *   linkShopTransfers():   every half-month of the year, with the hints
 *   mirrorShopTransfers(): every half-month, copies only (before the uncounted counts are filled)
 * The hourly job does the current half-month and the two before it; an edit of U or AK does
 * its own half-month at once. Both pinks are kept by the colour tidy-up (dividers.gs).
 * Needs config.gs, dividers.gs, year-sync.gs (normalizeKey, positionOfDate), instant-sync.gs.
 */

const TRANSFER_AUTO_COLOR = DIVIDER_TRANSFER_COLORS[0];
const TRANSFER_CONFLICT_COLOR = DIVIDER_TRANSFER_COLORS[1];
const TRANSFER_HINT_COLOR = DIVIDER_TRANSFER_COLORS[3];
const TRANSFER_NOTE_MARK = 'Шилжүүлэг (автомат)';
const TRANSFER_HINT_MARK = 'Шилжүүлэг (санал)';
const TRANSFER_HINT_FROM = { year: 2026, month: 10, half: 1 }; // 2026/10/31: none before (the owner, 2026-10-24)
const TRANSFER_MATCHES_KEPT_UNTIL = { year: 2026, month: 3, half: 1 }; // 2026/3/31: old matches stay up to here
// what an earlier version wrote on the transfers it added (read only to take them out again)
const TRANSFER_MATCH_TEXT = 'тооллогоор тааруулсан';
const TRANSFER_NEGATIVE_TEXT = 'хасах үлдэгдлээр тааруулсан';
const TRANSFER_CONFLICT_MARK = 'Шилжүүлэг зөрүүтэй';
const TRANSFER_OCR_MARKS = ['OCR:', 'Зураг: https://drive.google.com/file/d/', '&ocr='];
const TRANSFER_FIRST_ROW = 4; // below the three header rows
const TRANSFER_RECENT_PERIODS = 3; // hourly: this half-month and the two before
const TRANSFER_TIME_BUDGET_MS = 4.5 * 60 * 1000;
const TRANSFER_NEXT_KEY = 'transferLinkNext'; // document property: where a cut-off run goes on

// ---------- pure helpers (no Spreadsheet calls) ----------

const transferNumber = (value) => (typeof value === 'number' && Number.isFinite(value) ? value : 0);
const isTransferMarkLine = (line) =>
  [TRANSFER_NOTE_MARK, TRANSFER_CONFLICT_MARK, TRANSFER_HINT_MARK].some((mark) => line.startsWith(mark));
const isMatchLine = (line) => line.includes(TRANSFER_MATCH_TEXT) || line.includes(TRANSFER_NEGATIVE_TEXT);
const otherNoteLines = (note) => String(note || '').split('\n').filter((line) => line.trim() !== '' && !isTransferMarkLine(line));

// A transfer cell: { amount, typed (null: nothing a person or the OCR typed), mirrored, matched, ocr }
function parseTransferCell(value, note, color) {
  const text = String(note || '');
  const amount = transferNumber(value);
  const ocr = TRANSFER_OCR_MARKS.some((mark) => text.includes(mark));
  const matchedText = text.match(new RegExp(`(?:${TRANSFER_MATCH_TEXT}|${TRANSFER_NEGATIVE_TEXT}) \\+(\\d+)`));
  const matched = ocr || !matchedText ? 0 : Number(matchedText[1]);
  const copied = ocr ? null : text.match(/-аас хуулсан(?: \((\d+)\))?/);
  // a copy whose number was typed over is a typed number now
  const isUntouchedCopy = copied !== null && (copied[1] === undefined || Number(copied[1]) === amount - matched);
  const typedAmount = amount - matched;
  return {
    amount,
    typed: isUntouchedCopy || typedAmount === 0 ? null : typedAmount,
    mirrored: isUntouchedCopy,
    matched,
    ocr,
  };
}

// The new state of one transfer between the shops: out (the giver's AK), into (the taker's U),
// each a parsed cell or null when that shop has no row of the product; gaps: count - balance of
// the giver and the taker (null: not counted). An old match's added part is never kept.
// -> { conflict, missing, outAmount, inAmount, hint (how many may have moved unwritten), copyIn }
function planTransfer(out, into, gaps, hint) {
  if (!out || !into) {
    const missing = Boolean(out) && out.typed !== null;
    const keep = (cell) => (cell && cell.typed !== null ? cell.typed : 0); // a copy or a match of a vanished row goes
    return { missing, conflict: false, outAmount: keep(out), inAmount: keep(into), hint: 0, copyIn: false };
  }
  if (out.typed !== null && into.typed !== null && out.typed !== into.typed) {
    return { conflict: true, missing: false, outAmount: out.amount, inAmount: into.amount, hint: 0, copyIn: false };
  }
  const linked = out.typed !== null ? out.typed : 0; // what went from one shop to the other
  const fromElsewhere = out.typed === null && into.typed !== null ? into.typed : 0; // e.g. from Тосгон
  const canHint = hint && !out.ocr && !into.ocr && gaps.out !== null && gaps.into !== null;
  // the gaps as they will be once the cells hold these numbers (the out side lowers the giver's balance)
  const giver = canHint ? gaps.out + (linked - out.amount) : 0;
  const taker = canHint ? gaps.into - (linked + fromElsewhere - into.amount) : 0;
  return {
    conflict: false,
    missing: false,
    outAmount: linked,
    inAmount: linked + fromElsewhere,
    hint: giver < 0 && taker > 0 ? Math.min(-giver, taker) : 0,
    copyIn: into.typed === null && linked > 0,
  };
}

// True for a half-month whose old matches stay: up to TRANSFER_MATCHES_KEPT_UNTIL.
function keepsOldMatches(position) {
  const { year, month, half } = TRANSFER_MATCHES_KEPT_UNTIL;
  return position <= periodPosition(year, month, half);
}

// True for a half-month that gets the hints: from TRANSFER_HINT_FROM on.
function isHintPeriod(position) {
  const { year, month, half } = TRANSFER_HINT_FROM;
  return position >= periodPosition(year, month, half);
}

function transferNote(previousNote, lines) {
  return [...otherNoteLines(previousNote), ...lines].join('\n');
}

// The colour a column has on most product rows, leaving out the pinks and the OCR's colours.
function usualTransferColor(backgrounds) {
  const special = [...DIVIDER_TRANSFER_COLORS, ...DIVIDER_OCR_COLORS];
  const counts = backgrounds
    .map(([color]) => String(color).toLowerCase())
    .filter((color) => !special.includes(color))
    .reduce((tally, color) => ({ ...tally, [color]: (tally[color] || 0) + 1 }), {});
  const ranked = Object.entries(counts).sort((a, b) => b[1] - a[1]);
  return ranked.length > 0 ? ranked[0][0] : DIVIDER_WHITE;
}

// ---------- spreadsheet access ----------

function linkShopTransfers() {
  withScriptLock(() => Logger.log('%s\n%s', modeLabel(), runTransferLinkAll(currentBook(), { hint: true }).join('\n')));
}

function mirrorShopTransfers() {
  withScriptLock(() => Logger.log('%s\n%s', modeLabel(), runTransferLinkAll(currentBook(), { hint: false }).join('\n')));
}

// Run by hand once (the owner, 2026-10-24): the parts an earlier version added to the transfers
// ("тооллогоор тааруулсан +n", "хасах үлдэгдлээр тааруулсан +n") come out of every shop sheet of
// the year; nothing else of the link is done.
function removeCountMatches() {
  withScriptLock(() => Logger.log('%s\n%s', modeLabel(), runRemoveCountMatches(currentBook()).join('\n')));
}

function runRemoveCountMatches(spreadsheet) {
  const lines = spreadsheet.getSheets()
    .filter((sheet) => {
      const slot = sheetSlot(sheet.getName());
      return slot !== null && !keepsOldMatches(slot.position) && sheet.getLastRow() >= TRANSFER_FIRST_ROW;
    })
    .map(removeSheetMatches)
    .filter((line) => line !== '');
  const verb = DRY_RUN ? 'арилгана' : 'арилгасан';
  return lines.length > 0 ? [`Тааруулсан шилжүүлэг ${verb} (хуудас: нүд):`, ...lines] : ['Тааруулсан шилжүүлэг олдсонгүй.'];
}

// -> "<sheet>: <cells>", or '' when the sheet holds no match
function removeSheetMatches(sheet) {
  const book = readTransferSheet(sheet);
  const change = { in: new Map(), out: new Map() };
  ['in', 'out'].forEach((side) => {
    const column = book.columns[side];
    column.values.forEach(([value], index) => {
      const note = String(column.notes[index][0] || '');
      const cell = parseTransferCell(value, note, column.backgrounds[index][0]);
      if (cell.matched > 0) {
        const lines = note.split('\n').filter((line) => isTransferMarkLine(line) && !isMatchLine(line));
        change[side].set(index, { amount: cell.amount - cell.matched, lines, color: lines.length > 0 ? TRANSFER_AUTO_COLOR : null });
      }
    });
  });
  const cells = change.in.size + change.out.size;
  if (cells === 0) {
    return '';
  }
  writeTransferChanges(book, change);
  return `${sheet.getName()}: ${cells}`;
}

// Every half-month of the year, from where a cut-off run stopped.
function runTransferLinkAll(spreadsheet, options) {
  const props = bookProperties();
  const start = Number(props.getProperty(TRANSFER_NEXT_KEY)) || 0;
  const positions = Array.from({ length: CHAIN_PERIOD_COUNT - start }, (_, index) => start + index);
  const lines = runTransferLink(spreadsheet, positions, { ...options, markMissing: false }, Date.now() + TRANSFER_TIME_BUDGET_MS);
  const left = lines.left;
  if (!DRY_RUN) {
    if (left === null) {
      props.deleteProperty(TRANSFER_NEXT_KEY);
    } else {
      props.setProperty(TRANSFER_NEXT_KEY, String(left));
    }
  }
  const done = left === null ? 'Бүх хагас сар шалгагдлаа.' : 'Хугацаа дууслаа: дахин ажиллуулбал үргэлжилнэ.';
  return [...lines, done];
}

// Hourly: the current half-month and the two before it.
function runRecentTransferLink(spreadsheet) {
  const now = positionOfDate(new Date());
  const positions = Array.from({ length: TRANSFER_RECENT_PERIODS }, (_, back) => now - back)
    .filter((position) => position >= 0 && position < CHAIN_PERIOD_COUNT)
    .reverse();
  return runTransferLink(spreadsheet, positions, { hint: true, markMissing: true }, Date.now() + TRANSFER_TIME_BUDGET_MS);
}

const isTransferEdit = (edit) =>
  sheetSlot(edit.sheetName) !== null &&
  edit.lastRow >= TRANSFER_FIRST_ROW &&
  [LAYOUT.transferIn, LAYOUT.transferOut].some((column) => edit.firstColumn <= column && edit.lastColumn >= column);

// An edit of U or AK in a shop sheet: that half-month at once. A copy typed over loses its mark.
function runEditedTransferLink(range) {
  const slot = sheetSlot(range.getSheet().getName());
  if (slot === null || CHAINS.length !== 2) {
    return;
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(INSTANT_LOCK_WAIT_MS)) {
    return; // the hourly job covers it
  }
  try {
    forgetTransferMarks(range);
    Logger.log(runTransferLink(currentBook(), [slot.position], { hint: true, markMissing: true }, Date.now() + TRANSFER_TIME_BUDGET_MS).join('\n'));
  } finally {
    lock.releaseLock();
  }
}

// The edited U and AK cells lose their copy mark: what a person typed there is typed.
function forgetTransferMarks(range) {
  const sheet = range.getSheet();
  [LAYOUT.transferIn, LAYOUT.transferOut]
    .filter((column) => column >= range.getColumn() && column <= range.getLastColumn())
    .forEach((column) => {
      const cells = sheet.getRange(range.getRow(), column, range.getNumRows(), 1);
      const notes = cells.getNotes();
      const cleaned = notes.map(([note]) => [String(note || '').split('\n').filter((line) => !line.startsWith(TRANSFER_NOTE_MARK)).join('\n')]);
      if (!DRY_RUN && cleaned.some(([note], index) => note !== notes[index][0])) {
        cells.setNotes(cleaned);
      }
    });
}

// Links the half-months of `positions` in order until the deadline. The returned lines carry
// .left: the first position not done (null: all done).
function runTransferLink(spreadsheet, positions, options, deadline) {
  const lines = [];
  let left = null;
  positions.some((position) => {
    if (Date.now() > deadline) {
      left = position;
      return true;
    }
    lines.push(...linkTransferPeriod(spreadsheet, position, options));
    return false;
  });
  lines.left = left;
  return lines;
}

// One shop sheet's rows, by product key, with the transfer columns as they are now.
function readTransferSheet(sheet) {
  const height = sheet.getLastRow() - TRANSFER_FIRST_ROW + 1;
  if (height <= 0) {
    return { sheet, keys: new Map(), columns: {} };
  }
  const values = sheet.getRange(TRANSFER_FIRST_ROW, 1, height, LAYOUT.balance).getValues();
  const column = (number) => {
    const range = sheet.getRange(TRANSFER_FIRST_ROW, number, height, 1);
    const backgrounds = range.getBackgrounds();
    return { range, values: values.map((row) => [row[number - 1]]), notes: range.getNotes(), backgrounds, usual: usualTransferColor(backgrounds) };
  };
  const columns = { in: column(LAYOUT.transferIn), out: column(LAYOUT.transferOut) };
  const keys = new Map();
  values.forEach((row, index) => {
    const key = normalizeKey(row[SYNC_NAME_COL - 1]);
    if (key !== '' && !keys.has(key)) {
      const gap = countGap(row[LAYOUT.count - 1], row[LAYOUT.balance - 1]);
      keys.set(key, { index, label: String(row[SYNC_NAME_COL - 1]).trim(), gap });
    }
  });
  return { sheet, keys, columns };
}

// count - balance, or null when not counted
function countGap(count, balance) {
  return typeof count === 'number' && typeof balance === 'number' ? count - balance : null;
}

function parsedCell(book, side, key) {
  const row = book.keys.get(key);
  if (!row) {
    return null;
  }
  const column = book.columns[side];
  return parseTransferCell(column.values[row.index][0], column.notes[row.index][0], column.backgrounds[row.index][0]);
}

function hasTransferMark(book, side, key) {
  const row = book.keys.get(key);
  return Boolean(row) && String(book.columns[side].notes[row.index][0] || '').split('\n').some(isTransferMarkLine);
}

// Both shops' sheets of one half-month. -> log lines
// options: { hint (the count hints, from TRANSFER_HINT_FROM on), markMissing (mark a transfer out to
// a product the other shop lacks; when false it is only listed) }
function linkTransferPeriod(spreadsheet, position, { hint, markMissing = true }) {
  const sheets = CHAINS.map((chain) => spreadsheet.getSheetByName(chainSheetName(chain, position)));
  if (CHAINS.length !== 2 || sheets.some((sheet) => !sheet)) {
    return [];
  }
  const books = sheets.map(readTransferSheet);
  if (books.some((book) => book.keys.size === 0)) {
    return [];
  }
  const shops = CHAINS.map((chain) => chain.label);
  // running gaps: a transfer written for one direction changes what the other direction sees
  const gaps = books.map((book) => new Map([...book.keys].map(([key, row]) => [key, row.gap])));
  const changes = books.map(() => ({ in: new Map(), out: new Map() }));
  const tally = { copied: 0, hinted: 0, unmatched: 0, conflicts: 0, missing: [] };
  const hints = Boolean(hint) && isHintPeriod(position);
  const keepsMatches = keepsOldMatches(position);
  const keys = [...new Set([...books[0].keys.keys(), ...books[1].keys.keys()])];
  const isQuiet = (cell) => !cell || (cell.amount === 0 && !cell.mirrored && cell.matched === 0);
  [[0, 1], [1, 0]].forEach(([giver, taker]) => {
    keys.forEach((key) => {
      const out = parsedCell(books[giver], 'out', key);
      const into = parsedCell(books[taker], 'in', key);
      if (keepsMatches && [out, into].some((cell) => cell && cell.matched > 0)) {
        return; // an old match before 2026/4/15 stays as it is
      }
      const gapOf = (book) => (gaps[book].has(key) ? gaps[book].get(key) : null);
      const hasMark = hasTransferMark(books[giver], 'out', key) || hasTransferMark(books[taker], 'in', key);
      const mayHint = hints && out && into && gapOf(giver) !== null && gapOf(taker) !== null && gapOf(giver) < 0 && gapOf(taker) > 0;
      if (isQuiet(out) && isQuiet(into) && !hasMark && !mayHint) {
        return;
      }
      const plan = planTransfer(out, into, { out: gapOf(giver), into: gapOf(taker) }, hints);
      const label = (books[giver].keys.get(key) || books[taker].keys.get(key)).label;
      applyTransferPlan({ books, changes, gaps, tally, shops, giver, taker, key, label, out, into, plan, markMissing });
    });
  });
  books.forEach((book, index) => writeTransferChanges(book, changes[index]));
  const parts = [
    tally.copied > 0 ? `${DRY_RUN ? 'хуулна' : 'хуулсан'} ${tally.copied}` : '',
    tally.hinted > 0 ? `${DRY_RUN ? 'санал тавина' : 'санал тавьсан'} ${tally.hinted} (тоо нэмээгүй)` : '',
    tally.unmatched > 0 ? `тааруулсныг ${DRY_RUN ? 'арилгана' : 'арилгасан'} ${tally.unmatched}` : '',
    tally.conflicts > 0 ? `зөрүүтэй ${tally.conflicts} (ягаан, ээж шийднэ)` : '',
    tally.missing.length > 0 ? `нөгөө дэлгүүрт мөргүй ${tally.missing.length} (тэмдэглээгүй): ${tally.missing.join(', ')}` : '',
  ].filter((part) => part !== '');
  return parts.length > 0 ? [`${sheets.map((sheet) => sheet.getName()).join(' ↔ ')}: ${parts.join(', ')}`] : [];
}

// The link to one side's transfer cell of a product (config.gs sheetCellLink), for the other side's note.
function transferCellLink(book, side, key) {
  const row = book.keys.get(key).index + TRANSFER_FIRST_ROW;
  const column = side === 'in' ? LAYOUT.transferIn : LAYOUT.transferOut;
  return sheetCellLink(book.sheet.getParent().getId(), book.sheet.getSheetId(), `${columnLetter(column)}${row}`);
}

// Puts one plan into the pending changes of both sides and moves the running gaps. A copy's note and
// a disagreement's notes link to the other shop's cell (2026-10-26).
function applyTransferPlan({ books, changes, gaps, tally, shops, giver, taker, key, label, out, into, plan, markMissing }) {
  const set = (book, side, next) => {
    const row = books[book].keys.get(key);
    if (row) {
      changes[book][side].set(row.index, next);
    }
  };
  if (!out || !into) {
    if (plan.missing && !markMissing) {
      tally.missing.push(`${shops[giver]} "${label}" ${out.typed}`);
      set(giver, 'out', { amount: out.amount, lines: [], color: null }); // only listed: a mark left by an earlier run goes
      return;
    }
    const [book, side, amount] = out ? [giver, 'out', plan.outAmount] : [taker, 'in', plan.inAmount];
    const note = `${TRANSFER_CONFLICT_MARK}: ${shops[taker]}-д "${label}" бараа алга. Тэнд мөр нэмэх эсвэл энэ тоог засна уу.`;
    set(book, side, { amount, lines: plan.missing ? [note] : [], color: plan.missing ? TRANSFER_CONFLICT_COLOR : null });
    tally.conflicts += plan.missing ? 1 : 0;
    return;
  }
  if (plan.conflict) {
    const note = `${TRANSFER_CONFLICT_MARK}: ${shops[giver]} гаргасан ${out.typed}, ${shops[taker]} авсан ${into.typed}. Аль нэгийг нь засна уу.`;
    set(giver, 'out', { amount: plan.outAmount, lines: [`${note} ${shops[taker]}: ${transferCellLink(books[taker], 'in', key)}`], color: TRANSFER_CONFLICT_COLOR });
    set(taker, 'in', { amount: plan.inAmount, lines: [`${note} ${shops[giver]}: ${transferCellLink(books[giver], 'out', key)}`], color: TRANSFER_CONFLICT_COLOR });
    tally.conflicts += 1;
    return;
  }
  const hintLine = `${TRANSFER_HINT_MARK}: тооллогоор ${shops[giver]}-аас ${shops[taker]} руу ${plan.hint} шилжсэн байж магадгүй. Тоонд нэмээгүй: зөв бол гараар бичнэ үү.`;
  const hintLines = plan.hint > 0 ? [hintLine] : [];
  const copyLines = plan.copyIn
    ? [`${TRANSFER_NOTE_MARK}: ${shops[giver]}-ийн ${columnLetter(LAYOUT.transferOut)}-аас хуулсан (${plan.outAmount}): ${transferCellLink(books[giver], 'out', key)}`]
    : [];
  const colorOf = (lines) => (hintLines.length > 0 ? TRANSFER_HINT_COLOR : lines.length > 0 ? TRANSFER_AUTO_COLOR : null);
  set(giver, 'out', { amount: plan.outAmount, lines: hintLines, color: colorOf(hintLines) });
  set(taker, 'in', { amount: plan.inAmount, lines: [...copyLines, ...hintLines], color: colorOf(copyLines) });
  if (plan.copyIn && into.amount - into.matched !== plan.inAmount) {
    tally.copied += 1;
  }
  const hasHintLine = (book, side) =>
    String(books[book].columns[side].notes[books[book].keys.get(key).index][0] || '').split('\n').includes(hintLine);
  const hadHint = hasHintLine(giver, 'out') && hasHintLine(taker, 'in');
  tally.hinted += plan.hint > 0 && !hadHint ? 1 : 0;
  tally.unmatched += out.matched > 0 || into.matched > 0 ? 1 : 0;
  // the giver's balance falls by what goes out, the taker's rises by what comes in
  const move = (book, delta) => {
    if (gaps[book].get(key) !== null && gaps[book].has(key)) {
      gaps[book].set(key, gaps[book].get(key) + delta);
    }
  };
  move(giver, plan.outAmount - out.amount);
  move(taker, -(plan.inAmount - into.amount));
}

// Writes the changed cells of one sheet, a whole column at a time. A pink that is no longer
// wanted goes back to the column's usual colour; a colour that is not ours stays.
function writeTransferChanges(book, change) {
  ['in', 'out'].forEach((side) => {
    const pending = change[side];
    if (pending.size === 0) {
      return;
    }
    const column = book.columns[side];
    const next = column.values.map(([value], index) => {
      const wanted = pending.get(index);
      const note = column.notes[index][0] || '';
      const color = column.backgrounds[index][0];
      if (!wanted) {
        return { value, note, color };
      }
      const isOurColor = DIVIDER_TRANSFER_COLORS.includes(String(color).toLowerCase());
      return {
        value: wanted.amount === 0 ? '' : wanted.amount,
        note: transferNote(note, wanted.lines),
        color: wanted.color || (isOurColor ? column.usual : color),
      };
    });
    const isSame = next.every((cell, index) =>
      cell.value === column.values[index][0] && cell.note === (column.notes[index][0] || '') && cell.color === column.backgrounds[index][0]);
    if (DRY_RUN || isSame) {
      return;
    }
    withFilterLifted(book.sheet, () => {
      column.range.setValues(next.map((cell) => [cell.value]));
      column.range.setNotes(next.map((cell) => [cell.note]));
      column.range.setBackgrounds(next.map((cell) => [cell.color]));
    });
  });
}
