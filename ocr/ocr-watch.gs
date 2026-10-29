/**
 * Receipt OCR - the photos mom deleted or moved after they were read. A review pass looks at the
 * month spreadsheets changed since the last one; deleting or moving a photo changes none (the owner,
 * 2026-10-28: My home's IMG_5873 pair, read loose as the 26th, then deleted; nothing else in
 * "OCR My home зарлага 2026-10" changed, so its rows and count stayed and the site showed the 26th).
 * So every pass looks where the photos read into this month's and last month's spreadsheets lie now,
 * and the spreadsheets of those deleted or put into another day, kind or month are reviewed with the
 * changed ones; their review moves or takes out the rows (ocrRefileDays). A photo of another place
 * now, outside a month folder, or loose in its own month stays where it is, as ocrPhotoMoves has it.
 */

const OCR_WATCH_MONTHS = 2; // this month and the one before
// "Зураг уншуулах/2026-10/Зарлага/OCR My home зарлага 2026-10 / 26" (ocr-done:<photo id>, fileOcrRows)
const OCR_DONE_SPOT = /^[^/]+\/(\d{4}-\d{2})\/[^/]+\/(OCR .+) \/ (\d{2})$/;

// The months watched: 'YYYY-MM', this one first.
const ocrWatchMonths = (now) => Array.from({ length: OCR_WATCH_MONTHS }, (_, back) => {
  const month = new Date(now.getFullYear(), now.getMonth() - back, 1);
  return `${month.getFullYear()}-${String(month.getMonth() + 1).padStart(2, '0')}`;
});

// The photos read into a month spreadsheet of the watched months: [{ photoId, book, place, day }].
function ocrWatchedPhotos(props, now) {
  const months = ocrWatchMonths(now);
  return Object.entries(props.getProperties())
    .filter(([key]) => key.startsWith(OCR_DONE_PREFIX))
    .map(([key, value]) => [key.slice(OCR_DONE_PREFIX.length), String(value).match(OCR_DONE_SPOT)])
    .filter(([, match]) => match && months.includes(match[1]) && ocrMonthBookOf(match[2]))
    .map(([photoId, match]) => ({ photoId, book: match[2], place: ocrMonthBookOf(match[2]).place, day: match[3] }));
}

// Whether a photo's rows are to leave the tab they are in: spot (ocrPhotoSpot) of a photo deleted,
// or of one in another day's folder or another spreadsheet's of its place.
function ocrPhotoLeft(spot, filed) {
  if (!spot) {
    return false;
  }
  if (spot.gone) {
    return true;
  }
  if (spot.place !== filed.place) {
    return false;
  }
  return ocrMonthBookName(spot.place, spot.kind, spot.month) !== filed.book || (spot.day !== null && spot.day !== filed.day);
}

// The names of the month spreadsheets with a photo deleted or moved; each photo's spot is kept for
// the review (context.located).
function ocrBooksOfLeftPhotos(context, now) {
  const located = context.located || (context.located = new Map());
  const paths = new Map();
  const left = ocrWatchedPhotos(context.props, now).filter((filed) => {
    if (!located.has(filed.photoId)) {
      located.set(filed.photoId, ocrPhotoSpot(filed.photoId, context.root, paths));
    }
    return ocrPhotoLeft(located.get(filed.photoId), filed);
  });
  return new Set(left.map((filed) => filed.book));
}

// entries (findOcrBooks) and the month spreadsheets of the photos deleted or moved, once each.
function ocrWithLeftPhotos(entries, context, now) {
  const names = ocrBooksOfLeftPhotos(context, now);
  const known = new Set(entries.map((entry) => entry.file.getName()));
  const missing = [...names].filter((name) => !known.has(name));
  if (missing.length === 0) {
    return entries;
  }
  Logger.log(`Зураг устгасан/зөөсөн тул шалгана: ${missing.join(', ')}`);
  return [...entries, ...findOcrBooks(context.root, null).filter((entry) => missing.includes(entry.file.getName()))];
}
