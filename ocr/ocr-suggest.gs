/**
 * Receipt OCR - for a line no product was matched to, the products it most likely is, so mom sees
 * them before she picks one in "Бараа" (where typing a few letters also narrows the list): the
 * note of that "Бараа" cell, and in "Шалгах шалтгаан":
 *   "нэр жагсаалтад таараагүй; Магадгүй: Haier 300l, Haier 200l, Haier 103l"
 * A product is scored by how alike the names are, its brand, the numbers in its name (a size,
 * a model) and its price. Lines read before this existed get theirs at the next review.
 */

const OCR_SUGGEST_COUNT = 3;
const OCR_SUGGEST_MIN_SCORE = 0.3;
const OCR_SUGGEST_MARK = 'Магадгүй:';
const OCR_SUGGEST_WEIGHTS = { name: 0.35, words: 0.3, numbers: 0.25, price: 0.1 };
// Cyrillic letters that look like Latin ones, so "Со 5л" and "Co 5l" are alike.
const OCR_LOOKALIKES = { а: 'a', в: 'b', е: 'e', к: 'k', м: 'm', н: 'h', о: 'o', р: 'p', с: 'c', т: 't', у: 'y', х: 'x', л: 'l' };
const ocrLatinLike = (text) => String(text || '').toLowerCase().replace(/[авекмнорстухл]/g, (letter) => OCR_LOOKALIKES[letter]);

const ocrNumbersIn = (text) => (String(text || '').match(/\d+(?:[.,]\d+)?/g) || []).map((number) => number.replace(',', '.'));

// 0..1: how likely the read name (and price) is this catalog entry { name, price }. Words are
// compared each with each (the telling word of a Mongolian name is not always its first).
function ocrSuggestScore(read, readPrice, entry) {
  const words = ocrWords(ocrLatinLike(read));
  const theirWords = ocrWords(ocrLatinLike(entry.name));
  const wordScore = words.length > 0 && theirWords.length > 0
    ? Math.max(...words.flatMap((word) => theirWords.map((their) => ocrSimilarity(word, their))))
    : 0;
  const numbers = ocrNumbersIn(read);
  const theirs = ocrNumbersIn(entry.name);
  const numberScore = numbers.length > 0 && theirs.length > 0
    ? numbers.filter((number) => theirs.includes(number)).length / Math.max(numbers.length, theirs.length)
    : 0;
  const price = typeof readPrice === 'number' && readPrice > 0 ? ocrFullPrice(readPrice) : 0;
  const priceScore = entry.price && price > 0 ? Math.max(0, 1 - Math.abs(entry.price - price) / entry.price) : 0;
  return OCR_SUGGEST_WEIGHTS.name * ocrSimilarity(ocrNameKey(ocrLatinLike(read)), ocrNameKey(ocrLatinLike(entry.name))) +
    OCR_SUGGEST_WEIGHTS.words * wordScore +
    OCR_SUGGEST_WEIGHTS.numbers * numberScore +
    OCR_SUGGEST_WEIGHTS.price * priceScore;
}

// The names of the likeliest entries, best first (none for a line with nothing read).
function ocrSuggestNames(read, readPrice, entries) {
  if (ocrText(read) === '') {
    return [];
  }
  return entries
    .map((entry) => ({ name: entry.name, score: ocrSuggestScore(read, readPrice, entry) }))
    .filter((candidate) => candidate.score >= OCR_SUGGEST_MIN_SCORE)
    .sort((a, b) => b.score - a.score || a.name.localeCompare(b.name))
    .slice(0, OCR_SUGGEST_COUNT)
    .map((candidate) => candidate.name);
}

const ocrSuggestText = (names) => (names.length > 0 ? `${OCR_SUGGEST_MARK} ${names.join(', ')}` : '');

// The reason of a line with no product in "Бараа" and no suggestion yet, with its suggestion;
// null when nothing changes.
function ocrSuggestedReason(record, entries) {
  const reason = ocrText(record['Шалгах шалтгаан']);
  if (ocrText(record[OCR_NAME_COLUMN]) !== '' || reason.includes(OCR_SUGGEST_MARK)) {
    return null;
  }
  const text = ocrSuggestText(ocrSuggestNames(record['Бараа (уншсан)'], record['Нэгж үнэ'], entries));
  return text === '' ? null : [reason, text].filter(Boolean).join('; ');
}

// The suggestion part of a reason ("Магадгүй: A, B, C"); '' when it has none.
const ocrSuggestionOf = (reason) => {
  const text = ocrText(reason);
  return text.includes(OCR_SUGGEST_MARK) ? text.slice(text.indexOf(OCR_SUGGEST_MARK)).split(';')[0].trim() : '';
};

// One column of a day tab written again when a value of it changed; each record in its own row
// (ocrRowOf), the blank rows between them as they are.
function ocrRewriteColumn(sheet, title, before, after) {
  if (after.length === 0 || after.every((record, index) => record[title] === before[index][title])) {
    return;
  }
  const range = ocrRecordsColumn(sheet, title, after);
  const values = range.getValues();
  after.forEach((record, index) => { values[ocrRowOf(record, index) - 2][0] = record[title]; });
  range.setValues(values);
}

// A column of a day tab from row 2 down to the last of these records.
const ocrRecordsColumn = (sheet, title, records) =>
  sheet.getRange(2, ocrColumn(title), Math.max(...records.map((record, index) => ocrRowOf(record, index))) - 1, 1);

// What the review puts right in a day tab for mom: a line with no product (read before Gemini had to
// pick one, 2026-10-25) gets its likeliest ones, in "Шалгах шалтгаан" and as the note of its "Бараа"
// cell (what mom sees in her view). The records as they then stand.
function ocrApplySuggestions(sheet, where, records, entries) {
  const updated = records.map((record) => {
    const reason = ocrSuggestedReason(record, entries);
    return reason === null ? record : { ...record, 'Шалгах шалтгаан': reason };
  });
  updated.forEach((record, index) => {
    if (record['Шалгах шалтгаан'] !== records[index]['Шалгах шалтгаан']) {
      Logger.log(`  ${where} мөр ${ocrRowOf(record, index)}: ${ocrSuggestionOf(record['Шалгах шалтгаан'])}`);
    }
  });
  if (DRY_RUN || updated.length === 0) {
    return updated;
  }
  ocrRewriteColumn(sheet, 'Шалгах шалтгаан', records, updated);
  const names = ocrRecordsColumn(sheet, OCR_NAME_COLUMN, updated);
  const existing = names.getNotes();
  const notes = existing.map(([note]) => [note]);
  updated.forEach((record, index) => {
    notes[ocrRowOf(record, index) - 2][0] = ocrText(record[OCR_NAME_COLUMN]) === '' ? ocrSuggestionOf(record['Шалгах шалтгаан']) : '';
  });
  if (notes.some((note, index) => note[0] !== existing[index][0])) {
    names.setNotes(notes);
  }
  return updated;
}
