/**
 * Receipt OCR - pure decisions (no Drive, no Gemini, so they are tested in node):
 *   - what the folder mom chose says (which place gets the goods, which gives them)
 *   - which product lines can be trusted (name from Sheet100, same brand, count, price, confidence)
 *   - where a line belongs in the inventory workbook (tab, block, day column)
 *   - the rows written to the day spreadsheet
 */

const OCR_INCOME_FOLDER = 'орлого';
const OCR_NO_PRICE_REASON = 'үнэ байхгүй'; // an income line without its price (ocrItemCheck)
const OCR_KIND_BY_FOLDER = { [OCR_INCOME_FOLDER]: 'orlogo', 'зарлага': 'zarlaga', 'тооллого': 'toollogo' };
// Which part of the sheet a folder's lines go to: income days, expense days, or the count column.
const OCR_SIDE_BY_KIND = { orlogo: 'in', zarlaga: 'out', toollogo: 'count' };
const OCR_BLOCK_BY_SIDE = { in: 'орлого', out: 'зарлага', count: 'тооллого' };
const OCR_TRANSFER_TYPES = ['shiljuuleg', 'butsaalt'];
const OCR_MONTH_DAYS = 30; // a sheet's last day column also takes day 31

// "Тосгоноос Соронзон руу" -> { from: 'Тосгон', to: 'Соронзон' }; anything else -> null.
function ocrTransferPlaces(folderName) {
  const name = String(folderName).trim();
  const from = OCR_PLACES.find((place) => name.startsWith(place));
  if (!from) {
    return null;
  }
  const to = OCR_PLACES.find((place) => place !== from && name.slice(from.length).includes(place));
  return to ? { from, to } : null;
}

// path: [place, folder, month?, day?], e.g. ['Соронзон', 'зарлага', '2026-10', '01'] or
// ['My home', 'My home-оос Тосгон руу (буцаалт)', ...]. A transfer or a return sits in the folder of
// the place that gives the goods. moves: the place that gets the goods (side 'in') and the place
// that gives them ('out'), each with the other place of the transfer.
function ocrFolderMeaning(path) {
  const [place, sub = ''] = path;
  const label = path.slice(0, 2).join(' / ');
  const unknown = { label, kind: null, moves: [] };
  if (!OCR_PLACES.includes(place)) {
    return unknown;
  }
  const kind = OCR_KIND_BY_FOLDER[sub];
  if (kind) {
    return { label, kind, moves: [{ place, side: OCR_SIDE_BY_KIND[kind] }] };
  }
  if (sub === `${OCR_INCOME_FOLDER}${OCR_FROM_TOSGON_MARK}`) {
    return place === OCR_TOSGON.place ? unknown : { label, kind: 'shiljuuleg', moves: ocrTransferMoves(OCR_TOSGON.place, place) };
  }
  if (sub === `${OCR_INCOME_FOLDER}${OCR_FROM_CITY_MARK}`) {
    return place === OCR_TOSGON.place ? unknown : { label, kind: 'shiljuuleg', moves: [{ place: OCR_TOSGON.place, side: 'in' }, ...ocrTransferMoves(OCR_TOSGON.place, place)] };
  }
  const places = ocrTransferPlaces(sub);
  if (!places || places.from !== place) {
    return unknown;
  }
  return { label, kind: sub.includes(OCR_RETURN_MARK) ? 'butsaalt' : 'shiljuuleg', moves: ocrTransferMoves(places.from, places.to) };
}

// The two sides of goods going from one place to another.
const ocrTransferMoves = (from, to) => [{ place: from, side: 'out', other: to }, { place: to, side: 'in', other: from }];

// The folder a slip counts as: one Тосгон sent, in a shop's income folder, is Тосгон's expense and the
// shop's income at once (the owner, 2026-10-28: one slip, two sides), as if it lay in "Тосгоноос <shop>
// руу"; its rows say "<shop> / орлого (Тосгоноос)". One from the city there is Тосгон's income too
// (three sides, "<shop> / орлого (Хотоос)"). Any other slip counts as its folder.
function ocrSlipFolder(folder, result) {
  const [move] = folder.moves;
  const sender = ocrSlipSender(result);
  if (folder.kind !== 'orlogo' || !move || move.place === OCR_TOSGON.place || !sender) {
    return folder;
  }
  const mark = sender === OCR_CITY ? OCR_FROM_CITY_MARK : OCR_FROM_TOSGON_MARK;
  return ocrFolderMeaning([move.place, `${OCR_INCOME_FOLDER}${mark}`]);
}

// The marks of a shop's income label whose slip counts on more sides (ocrSlipFolder).
const OCR_SLIP_MARKS = [OCR_FROM_TOSGON_MARK, OCR_FROM_CITY_MARK];

// A row read before 2026-10-28 from such a slip (noted as maybe misfiled, written as the shop's income
// only): with the new label and without the note (OK when nothing else is left), so the review writes
// Тосгон's side too; null for any other row. Its "үнэ байхгүй" goes too: a transfer needs no price, as
// a fresh reading of the slip would say (the owner, 2026-10-28: such a slip is yellow).
function ocrTosgonSentRecord(record) {
  const reasons = ocrText(record['Шалгах шалтгаан']).split('; ').filter((reason) => reason !== '');
  const note = reasons.map((reason) => reason.match(OCR_SENT_BY_TOSGON_NOTE)).find(Boolean);
  if (!note || ocrText(record['Хавтас (ээж хийсэн)']) !== `${note[1]} / ${OCR_INCOME_FOLDER}`) {
    return null;
  }
  const kept = reasons.filter((reason) => !OCR_SENT_BY_TOSGON_NOTE.test(reason) && reason !== OCR_NO_PRICE_REASON);
  return {
    ...record,
    'Хавтас (ээж хийсэн)': `${note[1]} / ${OCR_INCOME_FOLDER}${OCR_FROM_TOSGON_MARK}`,
    'Шалгах шалтгаан': kept.join('; '),
    ...(kept.length === 0 ? { 'Төлөв': OCR_STATUS_OK } : {}),
  };
}

// Where the copies of a Тосгон or city slip's photo go (the owner, 2026-10-28: the photo on every side):
// a photo in "Тосгоноос <shop> руу" into the shop's income folder; a slip Тосгон sent (sender
// "Тосгон", ocrSlipSender) in a shop's income folder into "Тосгоноос <shop> руу"; one from the city
// ("Хот") there into "Тосгоноос <shop> руу" and Тосгон's income folder; the same month and day.
// []: no copy (a return, a supplier's slip, a photo outside a month folder).
function ocrMirrorPaths(path, sender) {
  const [place, sub, month, day] = path;
  if (!OCR_MONTH_FOLDER.test(month || '')) {
    return [];
  }
  const rest = day ? [month, day] : [month];
  const folder = ocrFolderMeaning([place, sub]);
  if (folder.kind === 'shiljuuleg' && place === OCR_TOSGON.place) {
    return [[folder.moves[1].place, OCR_INCOME_FOLDER, ...rest]];
  }
  if (folder.kind !== 'orlogo' || place === OCR_TOSGON.place || !sender) {
    return [];
  }
  const toShop = [OCR_TOSGON.place, `Тосгоноос ${place} руу`, ...rest];
  return sender === OCR_CITY ? [toShop, [OCR_TOSGON.place, OCR_INCOME_FOLDER, ...rest]] : [toShop];
}

// Does the receipt's header name Тосгон as the sender: "Тос", "ТОСГОН агуулах", or one of mom's names
// for it as the first word ("Төв", "Төв агуулах"; OCR_PLACE_ALIASES), not inside another name.
const ocrPlaceAlias = (place) => new RegExp(`^(?:${OCR_PLACE_ALIASES[place].join('|')})(?!\\p{L})`, 'iu');
function ocrSentByTosgon(result) {
  const sender = String((result && result.header_sender) || '').trim();
  return /тос/i.test(sender) || ocrPlaceAlias(OCR_TOSGON.place).test(sender);
}

// Does the header name the city as the sender: "Хот", "Хотоос", "хотын агуулах" (the first word).
const ocrSentByCity = (result) => /^хот/iu.test(String((result && result.header_sender) || '').trim());

// Who sent a slip, as the slip rules need it: OCR_CITY, Тосгон, or null (a supplier, or no sender).
function ocrSlipSender(result) {
  if (ocrSentByCity(result)) {
    return OCR_CITY;
  }
  return ocrSentByTosgon(result) ? OCR_TOSGON.place : null;
}

// A slip that may belong in another folder (2026-10-23): an issue slip with
// no sender, in a shop's income folder (it may be a supplier's). It is written as filed and only noted:
// the owner asks, then mom moves the photo and the next run reads its new folder (ocrRefileMoved,
// ocr-post.gs). One Тосгон or the city sent is no longer noted: it counts on every side (ocrSlipFolder).
function ocrMisfiledReason(result, folder) {
  const [move] = folder.moves;
  if (folder.kind !== 'orlogo' || !move || move.place === OCR_TOSGON.place || ocrSlipSender(result)) {
    return null;
  }
  return result.document_type === 'zarlaga' ? ocrMisfiledNote(move.place) : null;
}

// The note in a few words (the owner, 2026-10-29: the reasons as wide as the screen beside the photo):
// "→ Тосгоноос <shop>?" is the folder it may belong in, if Тосгон sent it.
const ocrMisfiledNote = (place) => `${OCR_MISFILED_REASON}: илгээгч тодорхойгүй → Тосгоноос ${place}?`;

function ocrIsRealDate(iso) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(iso))) {
    return false;
  }
  const [year, month, day] = iso.split('-').map(Number);
  return month >= 1 && month <= 12 && day >= 1 && day <= daysInMonth(year, month);
}

// The column of a day in a layout: its income or expense day column, or the count column.
function ocrDayColumn(layout, side, dayIndex) {
  if (side === 'count') {
    return layout.count;
  }
  return (side === 'in' ? layout.incomeFirst : layout.expenseFirst) + dayIndex;
}

// A shop's transfer column (T in, AJ out; not sales) for goods between the two shops and for a
// return to Тосгон; goods from Тосгон are the shop's income of that day. null: a day column.
function ocrShopTransferColumn(move) {
  if (!move.other) {
    return null;
  }
  if (move.side === 'out') {
    return LAYOUT.transferOut;
  }
  return move.other === OCR_TOSGON.place ? null : LAYOUT.transferIn;
}

const OCR_TRANSFER_BLOCK = 'шилжүүлэг';
const OCR_GIFT_BLOCK = 'бэлэг';

// A shop's own sale (not a transfer, not Тосгон's): the only line mom may mark Бэлэг.
const ocrIsShopSale = (move) => move.side === 'out' && !move.other && move.place !== OCR_TOSGON.place;

// Where a line of that day belongs: { tab, block, column } or null outside the workbook's year.
// Тосгон: its month sheet, days in the day columns (it only gives to and gets from the shops).
function ocrTarget(move, iso) {
  const year = Number(iso.slice(0, 4));
  if (move.place !== OCR_TOSGON.place && !ocrHasBooksFor(year)) {
    return null; // a shop's year without a workbook
  }
  return withChainYear(year, () => ocrTargetInYear(move, iso)); // that year's tab names (year-books.gs)
}

function ocrTargetInYear(move, iso) {
  const [year, month, day] = iso.split('-').map(Number);
  const block = OCR_BLOCK_BY_SIDE[move.side];
  if (move.place === OCR_TOSGON.place) {
    const column = ocrDayColumn(OCR_TOSGON.layout, move.side, Math.min(day, OCR_MONTH_DAYS) - 1);
    return { tab: `${OCR_TOSGON.tabPrefix}/${year}/${month}`, block, column, layout: OCR_TOSGON.layout };
  }
  const chain = CHAINS.find((candidate) => candidate.label === move.place);
  const half = day > FIRST_HALF_LAST_DAY ? 1 : 0;
  const position = periodPosition(year, month, half);
  if (!chain || position < 0 || position >= CHAIN_PERIOD_COUNT) {
    return null;
  }
  const tab = chainSheetName(chain, position);
  if (move.gift) {
    return LAYOUT.gift ? { tab, block: OCR_GIFT_BLOCK, column: LAYOUT.gift, layout: LAYOUT } : null;
  }
  const transfer = ocrShopTransferColumn(move);
  if (transfer !== null) {
    return { tab, block: OCR_TRANSFER_BLOCK, column: transfer, layout: LAYOUT };
  }
  const offset = half === 0 ? day - 1 : Math.min(day, OCR_MONTH_DAYS) - FIRST_HALF_LAST_DAY - 1;
  return { tab, block, column: ocrDayColumn(LAYOUT, move.side, offset), layout: LAYOUT };
}

// A count is written into the count column (and carried to the next sheet's opening balance) only
// on the last day of its sheet: the 15th or the month's last day for a shop, the month's last day
// for Тосгон. A count on any other day is a check: compared with the books as they stood that day
// (ocr-count-report.gs); from the 10/31 sheets on it is also written into a block of its own
// (ocrIsSpotBlockCount), never into the count column.
function ocrIsPeriodEnd(place, iso) {
  const [year, month, day] = iso.split('-').map(Number);
  const last = daysInMonth(year, month);
  return day === last || (place !== OCR_TOSGON.place && day === FIRST_HALF_LAST_DAY);
}

const ocrIsSpotCount = (move, iso) => move.side === 'count' && !ocrIsPeriodEnd(move.place, iso);
const OCR_SPOT_PLACE = 'шалгах тооллого (workbook-д бичихгүй)';

// The last day of the sheet a day belongs to: the 15th or the month's end for a shop, the month's
// end for Тосгон ("2026-10-23" -> "2026-10-31").
function ocrPeriodEndIso(place, iso) {
  const [year, month, day] = iso.split('-').map(Number);
  const last = place !== OCR_TOSGON.place && day <= FIRST_HALF_LAST_DAY ? FIRST_HALF_LAST_DAY : daysInMonth(year, month);
  return `${iso.slice(0, 8)}${String(last).padStart(2, '0')}`;
}

// A count between the period ends whose sheet takes such counts (takesSpotCounts, config.gs, from
// the 10/31 sheets on): written into a block of its own at the sheet's right edge (ocr-spot.gs).
// Before that it is a check only, never written.
const ocrIsSpotBlockCount = (move, iso) => ocrIsSpotCount(move, iso) && takesSpotCounts(ocrPeriodEndIso(move.place, iso));

// "Sor/2026/10/15 · орлого · E", "Tos/2026/10 · тооллого · BO", "Sor/2026/10/31 · Тооллого 10/23";
// '' outside the workbook's year.
function ocrSheetPlace(move, iso) {
  if (ocrIsSpotCount(move, iso) && !ocrIsSpotBlockCount(move, iso)) {
    return OCR_SPOT_PLACE;
  }
  const target = ocrTarget(move, iso);
  if (target && ocrIsSpotBlockCount(move, iso)) {
    const [, month, day] = iso.split('-').map(Number);
    return `${target.tab} · ${spotChoice(month, day)}`;
  }
  return target ? `${target.tab} · ${target.block} · ${columnLetter(target.column)}` : '';
}

function ocrEditDistance(a, b) {
  let previous = Array.from({ length: b.length + 1 }, (_, index) => index);
  for (let i = 1; i <= a.length; i++) {
    const current = [i];
    for (let j = 1; j <= b.length; j++) {
      current.push(Math.min(previous[j] + 1, current[j - 1] + 1, previous[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1)));
    }
    previous = current;
  }
  return previous[b.length];
}

const ocrSimilarity = (a, b) => 1 - ocrEditDistance(a, b) / Math.max(a.length, b.length, 1);
// Words of two letters or more: a unit like "l" or "ш" is not a brand.
const ocrWords = (text) => (String(text || '').toLowerCase().match(/\p{L}+/gu) || []).filter((word) => word.length > 1);

// The brand of the matched name must be one of the words that were read ("Aucma 195" may not
// become "Panasonica 195l"). Without words on either side (e.g. "441l") the number decides.
function ocrBrandMatches(readName, matchedName) {
  const read = ocrWords(readName);
  const [brand] = ocrWords(matchedName);
  if (read.length === 0 || brand === undefined) {
    return true;
  }
  return Math.max(...read.map((word) => ocrSimilarity(word, brand))) >= OCR_BRAND_SIMILARITY;
}

// The number groups of a name, without the dots and commas inside them: "Haier 9.0-197" -> ["90", "197"].
const ocrNumberGroups = (text) => (String(text || '').match(/\d+(?:[.,]\d+)*/g) || []).map((group) => group.replace(/[.,]/g, ''));

// Every number of the matched name must be among the numbers that were read: "Aucma 200 l" may not
// become "Aucma 205l", but a model code the receipt adds ("Haier 100-628" for "Haier 100l") is fine.
function ocrNumbersMatch(readName, matchedName) {
  const read = new Set(ocrNumberGroups(readName));
  return ocrNumberGroups(matchedName).every((group) => read.has(group));
}

const ocrMoney = (value) => `${String(Math.round(value)).replace(/\B(?=(\d{3})+(?!\d))/g, ',')}₮`;

// Why the receipt price does not fit the price the customer sees, or null when it fits or a price is missing.
function ocrPriceMismatch(unitPrice, listPrice) {
  if (typeof unitPrice !== 'number' || unitPrice <= 0 || typeof listPrice !== 'number' || listPrice <= 0) {
    return null;
  }
  const paid = ocrFullPrice(unitPrice);
  const list = listPrice;
  return Math.abs(paid - list) <= list * OCR_PRICE_TOLERANCE ? null : `үнэ зөрсөн (баримт ${ocrMoney(paid)}, жагсаалт ${ocrMoney(list)})`;
}

// Why the matched name may not be used, or null. Each failure leaves the line for mom to pick.
function ocrMatchProblem(item, matched, catalog, prices) {
  if (matched === null || !catalog.has(matched)) {
    return 'нэр жагсаалтад таараагүй';
  }
  if (!ocrBrandMatches(item.name, matched)) {
    return `брэнд зөрсөн (${item.name} ≠ ${matched})`;
  }
  if (!ocrNumbersMatch(item.name, matched)) {
    return `тоо зөрсөн (${item.name} ≠ ${matched})`;
  }
  return ocrPriceMismatch(item.unit_price, prices.get(matched));
}

// The product of a line is always one name of the list, OCR_NOT_RELEVANT when none fits (the owner,
// 2026-10-23: no threshold, mom corrects a wrong pick); a name not on the list becomes OCR_NOT_RELEVANT.
// A doubtful pick (brand, size or price not as read) stays and only gives its reason, so mom checks it.
// catalog: a Set of the Sheet100 names. prices: Map name -> the price the customer sees, may be empty.
function ocrItemCheck(item, catalog, needsPrice, prices = new Map()) {
  const picked = ocrText(item.matched_name);
  const matched = catalog.has(picked) ? picked : OCR_NOT_RELEVANT;
  const problem = matched === OCR_NOT_RELEVANT
    ? (picked === OCR_NOT_RELEVANT ? 'жагсаалтад тохирох бараа алга' : 'нэр жагсаалтад таараагүй')
    : ocrMatchProblem(item, matched, catalog, prices);
  const reasons = [
    problem,
    typeof item.qty !== 'number' ? 'тоо байхгүй' : null,
    needsPrice && typeof item.unit_price !== 'number' ? OCR_NO_PRICE_REASON : null,
    typeof item.confidence === 'number' && item.confidence < OCR_REVIEW_CONFIDENCE ? `итгэл бага (${item.confidence})` : null,
  ].filter(Boolean);
  return { matched, reasons };
}

// The day mom chose: ['…', '…', '2026-10', '01'] -> '2026-10-01'; null outside a day folder.
function ocrFolderDate(path) {
  const [, , month = '', day = ''] = path;
  const iso = `${month}-${day}`;
  return path.length === 4 && OCR_MONTH_FOLDER.test(month) && OCR_DAY_FOLDER.test(day) && ocrIsRealDate(iso) ? iso : null;
}

// "2026-12-05" -> ['2026-12', '2027-01'] (this month and OCR_MONTHS_AHEAD after it).
function ocrUpcomingMonths(today) {
  const [year, month] = today.split('-').map(Number);
  return Array.from({ length: OCR_MONTHS_AHEAD + 1 }, (_, index) => {
    const months = month - 1 + index;
    return `${year + Math.floor(months / MONTHS_PER_YEAR)}-${String((months % MONTHS_PER_YEAR) + 1).padStart(2, '0')}`;
  });
}

// The day folders of a month: "2026-10" -> ['01', ..., '31'].
function ocrMonthDays(month) {
  const [year, monthNumber] = month.split('-').map(Number);
  return Array.from({ length: daysInMonth(year, monthNumber) }, (_, index) => String(index + 1).padStart(2, '0'));
}

// The day a receipt outside a day folder is read as: the date written on it, when Gemini read one in
// the folder's month (the owner, 2026-10-28: a slip dated the 18th, left loose, was counted on the day
// it came in); its year is not trusted ("20__ он"). null: none, or another month.
function ocrWrittenDay(value, month) {
  const match = /^\d{4}-(\d{2})-(\d{2})$/.exec(ocrText(value));
  const iso = match && match[1] === month.slice(5) ? `${month}-${match[2]}` : '';
  return ocrIsRealDate(iso) ? iso : null;
}

// The month folder of a photo's path ('2026-10'), or null outside one.
function ocrFolderMonth(path) {
  const month = path[2] || '';
  return OCR_MONTH_FOLDER.test(month) ? month : null;
}

// What is wrong with the receipt as a whole, and its day. The day is where mom put the photo (the
// day folder); outside one, the date written on the receipt (ocrWrittenDay), else the day it was
// uploaded; a photo in a day folder keeps its day whatever the receipt says.
function ocrDocumentCheck(result, folder, folderDate, uploadedOn, folderMonth = null) {
  const written = folderDate ? null : ocrWrittenDay(result.date, folderMonth || uploadedOn.slice(0, 7));
  const reasons = folderDate ? [] : [written ? `${OCR_NO_DAY_REASON}, баримт дээрх огноогоор (${written})` : `${OCR_NO_DAY_REASON}, оруулсан өдрөөр (${uploadedOn})`];
  const items = Array.isArray(result.items) ? result.items : [];
  const isCount = folder.kind === 'toollogo'; // a count is a page of its own: its folder says what it is
  if (result.document_type === 'other' && !isCount) {
    reasons.push('маягтгүй хуудас, гараар шалгана');
  }
  if (result.document_type === 'blank' || items.length === 0) {
    reasons.push('хоосон баримт');
  }
  if (folder.moves.length === 0) {
    reasons.push('хавтас танигдсангүй');
  } else if (ocrMisfiledReason(result, folder)) {
    reasons.push(ocrMisfiledReason(result, folder));
  } else if (!OCR_TRANSFER_TYPES.includes(folder.kind) && OCR_TRANSFER_TYPES.includes(result.document_type)) {
    reasons.push(`шилжүүлэг/буцаалтын баримт мэт, хавтас нь ${folder.label}`);
  } else if (!isCount && result.document_type === 'toollogo') {
    reasons.push(`тооллогын хуудас мэт, хавтас нь ${folder.label}`);
  }
  return { date: folderDate || written || uploadedOn, reasons };
}

const ocrNumberCell = (value) => (typeof value === 'number' ? value : '');
// A receipt price in full tögrög: the receipts write thousands ("430" or "130.0" = 430,000 / 130,000₮).
// A notebook page of goods bought in China ("15x426"): its prices are rough cost prices in yuan, so
// no price check takes them for thousands of tögrög.
const OCR_CURRENCY_YUAN = 'CNY';
const ocrFullPrice = (value) => (typeof value === 'number' && value > 0 && value < OCR_RECEIPT_THOUSANDS ? value * 1000 : value);

function ocrIsTosgonIncome(folderLabel) {
  const folder = ocrFolderMeaning(ocrText(folderLabel).split(' / '));
  return folder.kind === 'orlogo' && folder.moves[0].place === OCR_TOSGON.place;
}

// An issue slip in Тосгон's income folder (the owner, 2026-10-23): it is not read into lines.
const ocrIsSkippedSlip = (result, folder) => result.document_type === 'zarlaga' && ocrIsTosgonIncome(folder.label);

// A row in OCR_HEADER order from named values; mom's columns start empty.
function ocrRow(values) {
  return OCR_HEADER.map((title) => (values[title] === undefined ? '' : values[title]));
}

// The ID of a photo's n-th row: "<file id>#1", ... It stays with the row and goes into the note of
// the workbook cell the row is written to, so a row is never written twice.
const ocrRowId = (fileId, index) => `${fileId}#${index + 1}`;

// The catalog's list prices by name; names without a price are left out.
const ocrPriceMap = (entries) => new Map(entries.filter((entry) => typeof entry.price === 'number').map((entry) => [entry.name, entry.price]));

// One row a product line (one row for a receipt without lines). photo: { name, url, id }.
// entries: the catalog with prices, for the likeliest products of a line Gemini found none for (ocr-suggest.gs).
// What was read as text, the receipt's price, title and kind are not kept: the photo is beside the
// line (the owner, 2026-10-25). Gemini's confidence is, in the hidden "Итгэл".
function ocrRows(result, photo, folder, check, catalog, readAt, entries = []) {
  const base = { 'Баримт': photo.name, 'Хавтас (ээж хийсэн)': folder.label };
  const items = Array.isArray(result.items) && result.items.length > 0 ? result.items : [null];
  const isYuan = result.price_currency === OCR_CURRENCY_YUAN;
  return items.map((item, index) => {
    const id = ocrRowId(photo.id, index);
    if (item === null) {
      return ocrRow({ ...base, 'ID': id, 'Төлөв': OCR_STATUS_REVIEW, 'Шалгах шалтгаан': check.reasons.join('; ') });
    }
    // a yuan cost price is not compared with the price the customer sees, nor used to suggest a product
    const { matched, reasons } = ocrItemCheck(item, catalog, folder.kind === 'orlogo', isYuan ? new Map() : ocrPriceMap(entries));
    const isProduct = matched !== OCR_NOT_RELEVANT;
    const suggestion = isProduct ? '' : ocrSuggestText(ocrSuggestNames(item.name, isYuan ? null : item.unit_price, entries));
    const all = [...check.reasons, ...reasons, suggestion].filter(Boolean);
    return ocrRow({
      ...base,
      'ID': id,
      [OCR_NAME_COLUMN]: matched,
      'Бараа (тааруулсан)': matched,
      [OCR_QTY_COLUMN]: ocrNumberCell(item.qty),
      'Тоо (уншсан)': ocrNumberCell(item.qty),
      'Итгэл': typeof item.confidence === 'number' ? item.confidence : '',
      'Төлөв': all.length > 0 ? OCR_STATUS_REVIEW : OCR_STATUS_OK,
      'Шалгах шалтгаан': all.join('; '),
    });
  });
}

// The receipt price of each row ocrRows makes, for "Баримтын үнэ" (the owner, 2026-10-27): full tögrög,
// a yuan cost price with its currency named, '' where none was read.
function ocrReadPrices(result) {
  const items = Array.isArray(result.items) && result.items.length > 0 ? result.items : [null];
  return items.map((item) => {
    if (!item || typeof item.unit_price !== 'number') {
      return '';
    }
    return result.price_currency === OCR_CURRENCY_YUAN ? `${item.unit_price} юань` : ocrFullPrice(item.unit_price);
  });
}

// A row for a photo that was not read (duplicate, Gemini failed).
function ocrNoteRow(photo, folder, reason, readAt) {
  return ocrRow({
    'Баримт': photo.name,
    'Хавтас (ээж хийсэн)': folder.label,
    'Төлөв': OCR_STATUS_REVIEW,
    'Шалгах шалтгаан': reason,
    'ID': ocrRowId(photo.id, 0),
  });
}

// The month folder of a day: "2026-10-01" -> "2026-10".
function ocrMonthFolderName(iso) {
  return iso.slice(0, 7);
}
