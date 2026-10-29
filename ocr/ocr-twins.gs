/**
 * Receipt OCR - one slip photographed twice (the owner, 2026-10-23). A new photo whose lines
 * (product and count) are mostly those of an earlier photo of the same folder in the month
 * spreadsheet is marked "давхар зураг: <earlier photo> (<its day>)" and waits: its lines are written
 * into the workbook only once mom approves them (ocrPostDecision), so the goods are not counted twice.
 * At least OCR_TWIN_MIN_LINES lines and OCR_TWIN_SHARE of both photos' lines must be alike.
 */

// The rows of a photo, marked when an earlier photo of the month spreadsheet is its twin. day: "DD".
function ocrMarkTwin(book, rows, day, context) {
  const photos = ocrBookPhotos(book, context);
  const records = rows.map((row) => Object.fromEntries(OCR_HEADER.map((title, index) => [title, row[index]])));
  const keys = records.map(ocrTwinKey).filter(Boolean);
  const folder = ocrText(records[0]['Хавтас (ээж хийсэн)']);
  const twin = photos.find((photo) => photo.folder === folder && ocrIsTwinOf(keys, photo.keys));
  photos.push({ name: ocrText(records[0]['Баримт']), day, folder, keys });
  if (!twin) {
    return rows;
  }
  const mark = `${OCR_TWIN_REASON}: ${twin.name} (${twin.day})`;
  const reason = OCR_HEADER.indexOf('Шалгах шалтгаан');
  const status = OCR_HEADER.indexOf('Төлөв');
  Logger.log(`  ${mark}`);
  return rows.map((row) => row.map((value, index) => {
    if (index === reason) {
      return [ocrText(value), mark].filter(Boolean).join('; ');
    }
    return index === status ? OCR_STATUS_REVIEW : value;
  }));
}

// "meiling501l|2": a line's product (mom's or the matched one, else what was read) and count.
function ocrTwinKey(record) {
  const name = ocrText(record[OCR_NAME_COLUMN]) || ocrText(record['Бараа (уншсан)']);
  return name ? `${ocrNameKey(name)}|${ocrText(record[OCR_QTY_COLUMN])}` : null;
}

function ocrIsTwinOf(keys, earlier) {
  const left = new Map();
  earlier.forEach((key) => left.set(key, (left.get(key) || 0) + 1));
  const shared = keys.filter((key) => {
    const count = left.get(key) || 0;
    left.set(key, count - 1);
    return count > 0;
  }).length;
  return shared >= OCR_TWIN_MIN_LINES && shared >= OCR_TWIN_SHARE * keys.length && shared >= OCR_TWIN_SHARE * earlier.length;
}

// The photos already in a month spreadsheet: [{ name, day, folder, keys }], read once a run.
function ocrBookPhotos(book, context) {
  context.twins = context.twins || new Map();
  if (!context.twins.has(book.getId())) {
    const photos = new Map();
    ocrDayTabs(book).forEach((tab) => ocrReadRecords(tab).forEach((record) => {
      const name = ocrText(record['Баримт']);
      if (!photos.has(name)) {
        photos.set(name, { name, day: tab.getName(), folder: ocrText(record['Хавтас (ээж хийсэн)']), keys: [] });
      }
      const key = ocrTwinKey(record);
      if (key) {
        photos.get(name).keys.push(key);
      }
    }));
    context.twins.set(book.getId(), [...photos.values()]);
  }
  return context.twins.get(book.getId());
}
