/**
 * Receipt OCR - mom's check of a line, as pure decisions (tested in node):
 *   - mom picks the product in "Бараа" (another than the script matched, or one where nothing was
 *     matched) or writes another count in "Тоо": that is her approval of the line
 *   - or she marks it Зөв or Буруу: Зөв = the line as it stands may go into the workbook, Буруу =
 *     it may not
 *   - whether a line is written into the workbook now, and with which name and count
 *   - what is left to do in a day spreadsheet, for mom's report (ocr-report.gs)
 */

const ocrText = (value) => String(value === undefined || value === null ? '' : value).trim();

// A row of the day spreadsheet as { title: value }.
function ocrRecord(row) {
  return Object.fromEntries(OCR_HEADER.map((title, index) => [title, row[index] === undefined ? '' : row[index]]));
}

// A row of a tab of a version before, read by its titles. With the staff's columns: their
// correction becomes mom's where she wrote none. With "Бараа (Sheet100)" and "Ээж: зөв бараа": the
// matched name is kept as "Бараа (тааруулсан)" and the name the line goes in as (the corrected
// one, else the matched one) becomes "Бараа". With "Ээж: зөв тоо": the count read is kept as
// "Тоо (уншсан)" and mom's count (else the count read) becomes "Тоо"; her note and the staff's mark
// go into "Шалгах шалтгаан".
function ocrUpgradeRecord(record) {
  const isOld = [OCR_OLD_NAME_COLUMNS.matched, OCR_OLD_STAFF_COLUMNS.mark, OCR_OLD_MOM_COLUMNS.qty].some((title) => title in record);
  if (!isOld) {
    return record;
  }
  const old = OCR_OLD_STAFF_COLUMNS;
  const mine = (title, theirs) => (ocrText(record[title]) !== '' ? record[title] : record[theirs] ?? '');
  const mark = ocrText(record[old.mark]);
  const staffNote = [mark === '' ? '' : `ажилчин ${mark}`, ocrText(record[old.note])].filter(Boolean).join(': ');
  const matched = ocrText(record[OCR_OLD_NAME_COLUMNS.matched] ?? record['Бараа (тааруулсан)']);
  const corrected = ocrText(mine(OCR_OLD_NAME_COLUMNS.corrected, old.name));
  const read = ocrText(record['Тоо (уншсан)']) !== '' ? record['Тоо (уншсан)'] : record[OCR_QTY_COLUMN];
  const typed = ocrTypedQty(mine(OCR_OLD_MOM_COLUMNS.qty, old.qty));
  const notes = [ocrText(record[OCR_OLD_MOM_COLUMNS.note]), staffNote].filter(Boolean).join(' · ');
  return {
    ...record,
    [OCR_NAME_COLUMN]: corrected || (OCR_OLD_NAME_COLUMNS.matched in record ? matched : record[OCR_NAME_COLUMN] ?? ''),
    'Бараа (тааруулсан)': matched,
    [OCR_QTY_COLUMN]: typed === null || Number.isNaN(typed) ? read : typed,
    'Тоо (уншсан)': read,
    'Шалгах шалтгаан': [ocrText(record['Шалгах шалтгаан']), notes && `тайлбар: ${notes}`].filter(Boolean).join('; '),
  };
}

// A "Бичсэн" text (ocrWrittenState): the cells a line went into, why it could not, or that it was
// taken out. No reason looks like one.
const ocrLooksWritten = (value) => {
  const text = ocrText(value);
  return text.startsWith(OCR_POST_FAILED) || text.startsWith(OCR_COMPARE_MARK) || /![A-Z]+\d+ ← /.test(text) ||
    [OCR_REMOVED_TEXT, OCR_NOT_RELEVANT_REMOVED_TEXT, OCR_OTHER_YEAR_TEXT].includes(text);
};

// A line of a tab of the layout before 2026-10-28 (the reason in F, "Бичсэн" in G) that mom's mark
// wrote at once by the new columns (ocrWriteTabNow did not move the tab first, until 2026-10-29):
// "Бичсэн" went into F, under the old reason's title, the reason into N beside the old titles
// (beside: what N holds). Both go back to their own titles; a reason that is a "Бичсэн" too (the line
// written again over the first damage) is lost: none.
function ocrRepairInstantWrite(record, beside) {
  if (!ocrLooksWritten(record[OCR_REASON_COLUMN])) {
    return record;
  }
  return { ...record, 'Бичсэн': record[OCR_REASON_COLUMN], [OCR_REASON_COLUMN]: ocrLooksWritten(beside) ? '' : ocrText(beside) };
}

// The reasons in a few words (the owner, 2026-10-29: mom reads them beside the photo, on one line of
// her screen). Each "; " part once; the long ones written before put the short way:
//   - the wrong-folder note of before 2026-10-29 (ocrMisfiledNote)
//   - a line that could not be written for a product with no row (OCR_NO_ROW_REASON): its name is in
//     "Бараа" already and the sheet in "Бичсэн", so only what is wrong; any other failure as it is
const OCR_FAILED_PREFIX = 'бичиж чадсангүй: ';
const OCR_OLD_MISFILED_NOTE = /^хавтас буруу байж магадгүй: зарлагын баримт, илгээгч тодорхойгүй \(Тосгоноос ирсэн бол Тосгоноос (.+) руу хавтсанд\)$/;
const OCR_CHECK_REASON_PREFIX = 'check: ';
function ocrFailureNote(reason) {
  const isCheck = reason.startsWith(OCR_CHECK_REASON_PREFIX);
  if (!OCR_NO_ROW_REASON.test(isCheck ? reason.slice(OCR_CHECK_REASON_PREFIX.length) : reason)) {
    return reason;
  }
  return isCheck ? 'check файлд мөр алга' : 'бараа ажлын номд алга';
}
function ocrShortReason(reason) {
  const misfiled = reason.match(OCR_OLD_MISFILED_NOTE);
  if (misfiled) {
    return ocrMisfiledNote(misfiled[1]);
  }
  return reason.startsWith(OCR_FAILED_PREFIX) ? OCR_FAILED_PREFIX + ocrFailureNote(reason.slice(OCR_FAILED_PREFIX.length)) : reason;
}
const ocrShortenReasons = (text) => [...new Set(ocrText(text).split('; ').filter((reason) => reason !== '').map(ocrShortReason))].join('; ');

// The count mom wrote: a number, null when none was written, NaN when it is not a number.
function ocrTypedQty(value) {
  if (typeof value === 'number') {
    return value;
  }
  const text = ocrText(value).replace(',', '.');
  if (text === '') {
    return null;
  }
  return /^\d+(\.\d+)?$/.test(text) ? Number(text) : NaN;
}

// Mom changed the line herself: "Бараа" holds another name than the script matched, or "Тоо"
// another count than was read.
function ocrIsChosenByMom(record) {
  const isOtherName = ocrText(record[OCR_NAME_COLUMN]) !== '' && ocrNameKey(record[OCR_NAME_COLUMN]) !== ocrNameKey(record['Бараа (тааруулсан)']);
  const read = ocrText(record['Тоо (уншсан)']);
  const isOtherCount = read !== '' && ocrText(record[OCR_QTY_COLUMN]) !== read;
  return isOtherName || isOtherCount;
}

function ocrReviewState(record) {
  const mom = ocrText(record['Ээж: Зөв/Буруу']);
  const chosen = ocrIsChosenByMom(record);
  return { mom, chosen, needsMom: mom === '' && !chosen };
}

// The name and count to write: "Бараа" and "Тоо" (mom's or what was read and matched).
function ocrLineValues(record) {
  const qty = ocrTypedQty(record[OCR_QTY_COLUMN]);
  return {
    name: ocrCleanName(ocrText(record[OCR_NAME_COLUMN])),
    qty: qty === null ? '' : qty,
  };
}

const ocrIsNotRelevant = (record) => ocrText(record[OCR_NAME_COLUMN]) === OCR_NOT_RELEVANT;

// A line of a receipt: it has a product, or a count was read for it (a line read before Gemini had to
// pick a product may have none, 2026-10-25); a row of a photo that was not read has neither.
const ocrIsProductLine = (record) =>
  ['Бараа (уншсан)', OCR_NAME_COLUMN, 'Бараа (тааруулсан)'].some((title) => ocrText(record[title]) !== '') || typeof record['Тоо (уншсан)'] === 'number';

// Mom approved the line: her Зөв or Бэлэг, or her change of "Бараа" or "Тоо". Its cell turns green.
const ocrIsApproved = (record) => [OCR_VERDICT_OK, OCR_VERDICT_GIFT].includes(ocrText(record['Ээж: Зөв/Буруу'])) || ocrIsChosenByMom(record);

// Mom said the line was given away (OCR_VERDICT_GIFT): approved, written into the gift column.
const ocrIsGift = (record) => ocrText(record['Ээж: Зөв/Буруу']) === OCR_VERDICT_GIFT;

// A line of a photo read almost like an earlier one (ocr-twins.gs): written only once mom approves it.
const ocrIsTwin = (record) => ocrText(record['Шалгах шалтгаан']).includes(`${OCR_TWIN_REASON}:`);

// Whether the line is written now: { post: true, name, qty, approved } | { post: false, wait }
// where wait is '' (nothing to write: no product, mom said Буруу or chose Хамааралгүй) or 'ээж' (no product name
// yet: it waits for her to pick one). A line with a name is written at once, yellow until mom
// approves it.
function ocrPostDecision(record) {
  const state = ocrReviewState(record);
  if (!ocrIsProductLine(record) || state.mom === OCR_VERDICT_WRONG || ocrIsNotRelevant(record)) {
    return { post: false, wait: '' };
  }
  const values = ocrLineValues(record);
  if (values.name === '' || (ocrIsTwin(record) && !ocrIsApproved(record))) {
    return { post: false, wait: 'ээж' };
  }
  const approved = ocrIsApproved(record);
  // check: a Шалгах line mom has not marked; its cell is pink until she does (2026-10-25)
  return { post: true, ...values, approved, check: !approved && ocrText(record['Төлөв']) === OCR_STATUS_REVIEW };
}

// Why a name and count can not be written; '' when they can.
function ocrValueProblem(values) {
  if (values.name === '') {
    return 'бараа сонгоогүй ("Бараа"-д жагсаалтаас сонгоно)';
  }
  if (typeof values.qty !== 'number' || !(values.qty >= 0)) {
    return 'тоо буруу ("Тоо"-д тоогоор бичнэ)';
  }
  return '';
}

// "Бичсэн" of a row, one line per workbook cell: "Tos/2026/10!AJ57 ← 2 · Haier 300l ✓", where ✓
// marks mom's approval and ⚠ a Шалгах line she has not marked yet, and a cell of the check copy (the original cell was typed by hand) is
// "check:Sor/2026/10/15!U22 ← 1 · Sevenstar 11 (гараар 2)" or "(гараар бичсэнтэй таарсан)".
// A failed write: "✗ reason [k:...]" on the first line (tried again once the checks of the line
// changed), then the cells written before, so a later correction can still take them out.
// Taken out of the workbook again (mom said Буруу): OCR_REMOVED_TEXT.
const OCR_WRITTEN_LINE = /^(\S+![A-Z]+\d+) ← /;
const OCR_REMOVED_TEXT = '— workbook-оос хассан (ээж Буруу гэсэн)';
const OCR_NOT_RELEVANT_REMOVED_TEXT = `— workbook-оос хассан (${OCR_NOT_RELEVANT})`;
const OCR_HAND_MATCH_TEXT = 'гараар бичсэнтэй таарсан';
const ocrWrittenLine = (cell, values, status = '') =>
  `${cell} ← ${values.qty} · ${values.name}${ocrLineMark(values)}${status ? ` (${status})` : ''}`;
// The mark after a line's name, in "Бичсэн" and in a cell's note: ✓ approved, ⚠ still to be checked.
const ocrLineMark = (line) => (line.approved ? ` ${OCR_APPROVED_MARK}` : line.check ? ` ${OCR_MUST_CHECK_MARK}` : '');
// A name read back without its mark.
const ocrWithoutMark = (text) => text.replace(new RegExp(` (${OCR_APPROVED_MARK}|${OCR_MUST_CHECK_MARK})$`), '');
const ocrWrittenLines = (written) => ocrText(written).split('\n').filter((line) => OCR_WRITTEN_LINE.test(line));
const ocrWrittenCells = (written) => ocrWrittenLines(written).map((line) => line.match(OCR_WRITTEN_LINE)[1]);

// The checks of a line, and the version of the writing rules: a line refused under the rules
// before (2026-10-20: a hand-typed cell refused it; v3, 2026-10-28: a failed line left the cell of
// the product it no longer is) is tried again under these.
const OCR_WRITE_RULES = 'v3';
function ocrFingerprint(record) {
  return [OCR_WRITE_RULES, ...['Ээж: Зөв/Буруу', OCR_NAME_COLUMN, OCR_QTY_COLUMN].map((title) => ocrText(record[title]))].join('|');
}
function ocrFailedText(reason, record, takenOut = []) {
  const kept = ocrWrittenLines(record['Бичсэн']).filter((line) => !takenOut.includes(line.match(OCR_WRITTEN_LINE)[1]));
  return [`${OCR_POST_FAILED} ${reason} [k:${ocrFingerprint(record)}]`, ...kept].join('\n');
}
const ocrFailedReason = (written) => ocrText(written).split('\n')[0].replace(/^✗\s*/, '').replace(/\s*\[k:[^\]]*\]$/, '');

// 'new' | 'done' | 'removed' | 'failed' (waits for a change of its checks) | 'retry' | 'compare'
// (a line of the version before, only compared with the workbook: written like a new one) |
// 'other year' (not written) | 'hands off' (its sheet out of the OCR's reach, OCR_HANDS_OFF_PERIOD:
// looked at again like a new line, as mom may change its folder or day)
function ocrWrittenState(record) {
  const written = ocrText(record['Бичсэн']);
  if (written === '') {
    return 'new';
  }
  if (written === OCR_HANDS_OFF_TEXT) {
    return 'hands off';
  }
  if (written === OCR_REMOVED_TEXT || written === OCR_NOT_RELEVANT_REMOVED_TEXT) {
    return 'removed';
  }
  if (written === OCR_OTHER_YEAR_TEXT) {
    return 'other year';
  }
  if (written.startsWith(OCR_COMPARE_MARK)) {
    return 'compare';
  }
  if (!written.startsWith(OCR_POST_FAILED)) {
    return 'done';
  }
  return written.split('\n')[0].endsWith(`[k:${ocrFingerprint(record)}]`) ? 'failed' : 'retry';
}

const OCR_SUMMARY_LIST_LIMIT = 15;

// What is left to do in one day spreadsheet. sheet: { id, url, name, place, label }.
function ocrSheetSummary(sheet, records) {
  const states = records.map((record) => ({ record, state: ocrReviewState(record), written: ocrWrittenState(record) }));
  const product = (record) => ocrLineValues(record).name || ocrText(record['Бараа (уншсан)']);
  const failed = states
    .filter(({ written }) => written === 'failed')
    .map(({ record }) => ({ product: product(record), reason: ocrFailedReason(record['Бичсэн']) }));
  // lines not in the workbook yet, for the count report: a shortage may only be a line not written
  const waiting = states
    .filter(({ record, written }) => written === 'failed' || (!['done', 'removed', 'other year', 'hands off'].includes(written) && ocrPostDecision(record).wait !== ''))
    .map(({ record }) => ({ product: product(record), places: ocrFolderMeaning(ocrText(record['Хавтас (ээж хийсэн)']).split(' / ')).moves.map((move) => move.place) }))
    .filter((line) => line.product !== '');
  const count = (test) => states.filter(test).length;
  return {
    ...sheet,
    rows: records.length,
    mom: count(({ state }) => state.needsMom),
    failed: failed.slice(0, OCR_SUMMARY_LIST_LIMIT),
    waiting: waiting.slice(0, OCR_SUMMARY_LIST_LIMIT),
    check: ocrHandDifferences(records, product).slice(0, OCR_SUMMARY_LIST_LIMIT),
  };
}

// The lines written into the check copy whose receipts make another number than mom typed, and
// that she has not approved yet: [{ product, text }] for her report.
function ocrHandDifferences(records, product) {
  return records
    .filter((record) => !ocrIsApproved(record) && ocrText(record['Ээж: Зөв/Буруу']) === '')
    .map((record) => ({
      product: product(record),
      text: ocrWrittenLines(record['Бичсэн'])
        .filter((line) => line.startsWith(OCR_CHECK_PREFIX) && /\(гараар \S+\)$/.test(line))
        .map((line) => line.slice(OCR_CHECK_PREFIX.length))
        .join('; '),
    }))
    .filter((line) => line.text !== '');
}

const ocrSummaryIsOpen = (summary) => summary.mom + summary.failed.length + (summary.check || []).length > 0;
