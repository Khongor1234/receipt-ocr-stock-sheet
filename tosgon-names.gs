/**
 * Тосгон (the warehouse, its own workbook) against the one product list "Бараа".
 *
 * Тосгон ships to both shops, so its names should be the list's names too. This
 * only READS the Тосгон workbook: it takes its newest month sheet that holds names
 * (up to today), pairs every name the list does not have with the list name that
 * looks like the same product, and writes the pairs as a yes / no table into
 * TOSGON_DECISION_TAB of this workbook, in the columns of "_Нэр тулгалт"
 * (product-list.gs): the owner marks I = тийм / үгүй and, where the suggestion is
 * wrong, writes the right name in J. Names with no pair come last: new products the
 * list may need. The table never renames anything; an existing table is never overwritten.
 *
 * How names are compared (pure helpers below): a key without case, spaces, "l"
 * after a number, look-alike Cyrillic letters and known brand typos ("Haeir" =
 * "Haier"). Same key -> 3; a shared word and the same numbers -> 2 ("8037 ус" /
 * "Ус буцалгагч 8037"); a shared word and a number that begins the other -> 1
 * ("Haier 11l" / "Haier 110-118"). Each list name is suggested for one Тосгон name;
 * another spelling of it is shown as "Давхар".
 *
 * Then, in the Тосгон workbook's own Apps Script, applyTosgonNames makes the rule hard
 * there too: the тийм answers rename every Тосгон month sheet, a тийм name the list
 * lacks is added to the list (column B "Тосгон"), a protected copy of the list
 * ("Бараа", IMPORTRANGE) is made in Тосгон, and column A of the month sheets becomes
 * a dropdown of that copy. It stops before changing anything while a row has no answer.
 *
 * Usage: makeTosgonNameTable in the main workbook (DRY_RUN = true first, then false);
 * restyleTosgonNameTable colours a table made before the colours, keeping its answers;
 * the owner answers the table; applyTosgonNames in the Тосгон workbook (the same way).
 * Needs config.gs, year-sync.gs, name-fixes.gs, product-list.gs.
 */

const TOSGON_BOOK_ID = 'YOUR_TOSGON_BOOK_ID'; // Тосгон 2026
const TOSGON_TAB_PATTERN = /^Tos\/(\d{4})\/(\d{1,2})$/;
const TOSGON_DECISION_TAB = '_Тосгон нэр тулгалт';
const TOSGON_TABLE_HEADER = [
  'Тосгоны нэр "Бараа" жагсаалттай таарахгүй байгаа нь. I баганад "тийм" эсвэл "үгүй" сонгоно уу. ' +
    'Санал болгосон нэр (H) буруу бол J баганад зөв нэрийг бичнэ. Энэ tab-д бичихэд юу ч өөрчлөгдөхгүй.\n' +
    'Өнгө: ногоон = зөвхөн бичлэг өөр, шар = үсгийн алдаа, улбар шар = эргэлзээтэй, улаан = давхар, саарал = жагсаалтад байхгүй. ' +
    'G багана: хэр төстэй (%) ба ялгаатай хэсэг (Тосгон → Бараа).',
  [
    '№',
    'Ангилал (Тосгон)',
    'Тосгон дахь нэр',
    'Тосгон зарах үнэ',
    '"Бараа" жагсаалтын нэр',
    'Аль дэлгүүрт',
    'Юугаараа ялгаатай',
    'Санал болгох нэг нэр',
    'Ижил бараа юу? (тийм/үгүй)',
    'Зөв нэр (санал буруу бол)',
    'Тайлбар',
  ],
];
const TOSGON_ANSWER_COLUMN = 9; // I
const MATCH_LEVEL_LABELS = {
  3: 'Зөвхөн бичлэг өөр (зай, том/жижиг үсэг, l)',
  2: 'Үсгийн алдаа / үгийн дараалал',
  1: 'Эргэлзээтэй - нэр төстэй',
};
const MATCH_NO_PAIR_LABEL = 'Жагсаалтад олдсонгүй - шинэ бараа уу?';
const MATCH_NO_PAIR_NOTE = 'тийм = энэ нэрээр "Бараа" жагсаалтад нэмнэ';
const MATCH_CONTESTED_LABEL = 'Давхар: энэ нэрийг Тосгоны өөр мөр бас авсан - нэгийг нь л "тийм"';
// Row colours by how sure the pair is (the kinds above), like the shops' table "_Нэр тулгалт".
const TOSGON_ROW_TONES = [
  [MATCH_LEVEL_LABELS[3], '#d9ead3'],
  [MATCH_LEVEL_LABELS[2], '#fff2cc'],
  [MATCH_LEVEL_LABELS[1], '#fce5cd'],
  [MATCH_CONTESTED_LABEL, '#f4cccc'],
  [MATCH_NO_PAIR_LABEL, '#efefef'],
];
const TOSGON_TITLE_TONE = '#fff2cc';
const TOSGON_HEADER_TONE = '#d9d9d9';
const TOSGON_ANSWER_TONES = { [DECISION_YES]: '#b6d7a8', [DECISION_NO]: '#ea9999' };
const TOSGON_CORRECTED_TONE = '#cfe2f3';
const TOSGON_COLUMN_WIDTHS = [40, 130, 170, 100, 170, 90, 300, 170, 120, 160, 220]; // A..K
const TOSGON_TITLE_HEIGHT = 56;
const TOSGON_PRICE_FORMAT = '#,##0 ₮';
const TOSGON_COLUMNS = { name: 3, price: 4, listName: 5, kind: 7, suggested: 8 }; // C D E G H
const TOSGON_LIST_LABEL = 'Тосгон'; // column B of a list row added from Тосгон
const LIST_COPY_TITLE = 'Үндсэн файлын "Бараа" жагсаалтын хуулбар (IMPORTRANGE). Энд засахгүй: нэрийг үндсэн файлд засна.';
const MATCH_LOOKALIKES = { а: 'a', е: 'e', о: 'o', с: 'c', р: 'p', х: 'x', к: 'k', м: 'm', т: 't', н: 'h', в: 'b', л: 'l' };
const MATCH_BRAND_TYPOS = {
  haeir: 'haier',
  acweis: 'acewise',
  acrwise: 'acewise',
  asewise: 'acewise',
  acuma: 'aucma',
  cmedia: 'cmidea',
  panansonico: 'panasonico',
  pnasonica: 'panasonica',
  hinsensin: 'hisense',
  hinsinsen: 'hisense',
  mieiling: 'meiling',
  dawoo: 'daewoo',
  dewoo: 'daewoo',
  plipis: 'pilipis',
  skyworch: 'skyworth',
  skywortch: 'skyworth',
};
const MATCH_UNIT = 'l';
const MATCH_MIN_WORD = 2;

function makeTosgonNameTable() {
  const lines = runTosgonNameTable(currentBook(), partnerBook(currentBook(), 'tosgon', TOSGON_BOOK_ID), new Date());
  Logger.log('%s\n%s', modeLabel(), lines.join('\n'));
}

// ---------- pure helpers (no Spreadsheet calls) ----------

const isNumberToken = (token) => /^\d/.test(token);

// Words and numbers of a name, lower case, look-alike letters and brand typos made one.
function matchTokens(name) {
  const text = cleanLabel(name).toLowerCase().replace(NUMBERED_NAME, '');
  return (text.match(/\p{L}+|\d+(?:[.,]\d+)?/gu) || []).map((token) => {
    if (isNumberToken(token)) {
      return token.replace(',', '.');
    }
    const latin = [...token].map((letter) => MATCH_LOOKALIKES[letter] || letter).join('');
    return MATCH_BRAND_TYPOS[latin] || latin;
  });
}

// The tokens without an "l" that only gives the unit of the number before it.
function meaningfulTokens(name) {
  const tokens = matchTokens(name);
  return tokens.filter((token, at) => !(token === MATCH_UNIT && at > 0 && isNumberToken(tokens[at - 1])));
}

function matchKey(name) {
  return meaningfulTokens(name).join('');
}

// How likely two names are one product: { level: 0..3, kind }.
function pairLevel(a, b) {
  if (matchKey(a) === matchKey(b)) {
    return { level: 3, kind: MATCH_LEVEL_LABELS[3] };
  }
  const words = (name) => new Set(meaningfulTokens(name).filter((token) => !isNumberToken(token) && token.length >= MATCH_MIN_WORD));
  const numbers = (name) => meaningfulTokens(name).filter(isNumberToken);
  const wordsOfB = words(b);
  if (![...words(a)].some((word) => wordsOfB.has(word))) {
    return { level: 0, kind: '' };
  }
  const [na, nb] = [numbers(a), numbers(b)];
  if (na.length > 0 && na.join('|') === nb.join('|')) {
    return { level: 2, kind: MATCH_LEVEL_LABELS[2] };
  }
  const [first, second] = [na[0] || '', nb[0] || ''];
  const begins = first !== '' && second !== '' && first !== second && (first.startsWith(second) || second.startsWith(first));
  return begins ? { level: 1, kind: MATCH_LEVEL_LABELS[1] } : { level: 0, kind: '' };
}

// Dice similarity of the letter pairs of two strings, 0..1.
function diceSimilarity(ka, kb) {
  const pairs = (key) => [...key].slice(1).map((letter, at) => key[at] + letter);
  const [pa, pb] = [pairs(ka), pairs(kb)];
  if (pa.length + pb.length === 0) {
    return 0;
  }
  const shared = pa.reduce(
    ({ count, rest }, pair) =>
      rest.includes(pair) ? { count: count + 1, rest: rest.filter((_, at) => at !== rest.indexOf(pair)) } : { count, rest },
    { count: 0, rest: pb }
  ).count;
  return (2 * shared) / (pa.length + pb.length);
}

// Similarity of the match keys: the tie-break between pairs.
const keySimilarity = (a, b) => diceSimilarity(matchKey(a), matchKey(b));

const lookalikeLatin = (token) => [...token.toLowerCase()].map((letter) => MATCH_LOOKALIKES[letter] || letter).join('');

// The words and numbers of a name as written (case kept), without an "l" / "л" after a number.
function writtenTokens(name) {
  const tokens = cleanLabel(name).replace(NUMBERED_NAME, '').match(/\p{L}+|\d+(?:[.,]\d+)?/gu) || [];
  return tokens.filter((token, at) => !(lookalikeLatin(token) === MATCH_UNIT && at > 0 && isNumberToken(tokens[at - 1])));
}

// Letters to change, add or drop to turn one string into the other (Levenshtein).
function editDistance(a, b) {
  const first = [...Array(b.length + 1).keys()];
  return [...a].reduce(
    (above, letter, i) =>
      [...b].reduce((row, other, j) => [...row, Math.min(above[j + 1] + 1, row[j] + 1, above[j] + (letter === other ? 0 : 1))], [i + 1]),
    first
  )[b.length];
}

// How alike two names look, 0..1: the share of letters that need no change. Look-alike
// letters count as one ("abto" = "авто"), a typo does not ("Pnasonica" / "Panasonica").
function spellingSimilarity(a, b) {
  const [ka, kb] = [a, b].map((name) => writtenTokens(name).map(lookalikeLatin).join(''));
  const longest = Math.max(ka.length, kb.length);
  return longest === 0 ? 0 : 1 - editDistance(ka, kb) / longest;
}

// The words of a that b lacks; same(token) is what counts as the same word.
function tokensMissing(ta, tb, same) {
  return ta.reduce(
    ({ kept, left }, token) => {
      const at = left.indexOf(same(token));
      return at >= 0 ? { kept, left: left.filter((_, i) => i !== at) } : { kept: [...kept, token], left };
    },
    { kept: [], left: tb.map(same) }
  ).kept;
}

// What the eye sees differ: "Cmidea → Cmedia", "латин / кирилл үсэг: abto → авто";
// '' when only spaces, case or the "l" differ.
function spellingDifference(a, b) {
  const [ta, tb] = [writtenTokens(a), writtenTokens(b)];
  const shown = (same) => {
    const [onlyA, onlyB] = [tokensMissing(ta, tb, same), tokensMissing(tb, ta, same)];
    return onlyA.length + onlyB.length === 0 ? '' : `${onlyA.join(' ') || '—'} → ${onlyB.join(' ') || '—'}`;
  };
  const byLetters = shown(lookalikeLatin);
  if (byLetters !== '') {
    return byLetters;
  }
  const byScript = shown((token) => token.toLowerCase());
  return byScript === '' ? '' : `латин / кирилл үсэг: ${byScript}`;
}

// Column G of a pair: why, how alike (%) and what differs.
function describeTosgonPair(tosgonName, listName, kind) {
  const difference = spellingDifference(tosgonName, listName);
  const percent = Math.round(spellingSimilarity(tosgonName, listName) * 100);
  return `${kind}\n${percent}% төстэй · ${difference || 'зөвхөн зай / том жижиг үсэг'}`;
}

// The row colour for column G's text (the kind it starts with); white for anything else.
function tosgonRowTone(kindText) {
  const found = TOSGON_ROW_TONES.find(([kind]) => String(kindText).startsWith(kind));
  return found ? found[1] : '#ffffff';
}

const isBetterPair = (pair, other) => pair.level > other.level || (pair.level === other.level && pair.score > other.score);

// tosgon: [{ name, category, price }]; list: [{ name, shops }] (products of "Бараа").
// Returns { rows: the table rows (A..K), same: names already on the list }.
function planTosgonPairs(tosgon, list) {
  const listKeys = new Set(list.map((entry) => normalizeKey(entry.name)));
  const open = tosgon.filter((product) => !listKeys.has(normalizeKey(product.name)));
  const best = open.map((product) =>
    list.reduce((top, entry) => {
      const { level, kind } = pairLevel(product.name, entry.name);
      const pair = { product, entry, level, kind, score: keySimilarity(product.name, entry.name) };
      return level > 0 && (top === null || isBetterPair(pair, top)) ? pair : top;
    }, null)
  );
  const pairs = best
    .filter((pair) => pair !== null)
    .filter((pair, _, all) => !all.some((other) => other !== pair && other.entry === pair.entry && isBetterPair(other, pair)))
    .filter((pair, at, all) => all.findIndex((other) => other.entry === pair.entry) === at) // a tie: the first one
    .sort((a, b) => b.level - a.level);
  // Another Тосгон spelling of a list name that a better one took: shown too, so both
  // can be seen; "тийм" on both is refused (planNameDecisions: one name, two pairs).
  const contested = best.filter((pair) => pair !== null && !pairs.includes(pair));
  const rowOf = ({ product, entry }, kind) => [
    product.category,
    product.name,
    product.price,
    entry.name,
    entry.shops,
    describeTosgonPair(product.name, entry.name, kind),
    entry.name,
    '',
    '',
    '',
  ];
  const newRows = open
    .filter((_, at) => best[at] === null)
    .map((product) => [product.category, product.name, product.price, '', '', MATCH_NO_PAIR_LABEL, cleanLabel(product.name), '', '', MATCH_NO_PAIR_NOTE]);
  return {
    rows: [...pairs.map((pair) => rowOf(pair, pair.kind)), ...contested.map((pair) => rowOf(pair, MATCH_CONTESTED_LABEL)), ...newRows].map(
      (row, at) => [at + 1, ...row]
    ),
    same: tosgon.length - open.length,
  };
}

// The Тосгон month tabs up to the month of `today`, newest first.
function tosgonTabsUpTo(names, today) {
  const limit = today.getFullYear() * 100 + today.getMonth() + 1;
  return names
    .map((name) => ({ name, match: TOSGON_TAB_PATTERN.exec(name) }))
    .filter(({ match }) => match !== null)
    .map(({ name, match }) => ({ name, order: Number(match[1]) * 100 + Number(match[2]) }))
    .filter(({ order }) => order <= limit)
    .sort((a, b) => b.order - a.order)
    .map(({ name }) => name);
}

// ---------- spreadsheet access ----------

// The products of "Бараа" (headings have no shop) as [{ name, shops }].
function readListProducts(sheet) {
  const height = sheet.getLastRow() - PRODUCT_LIST_FIRST_ROW + 1;
  if (height <= 0) {
    return [];
  }
  return sheet
    .getRange(PRODUCT_LIST_FIRST_ROW, 1, height, 2)
    .getValues()
    .map(([name, shops]) => ({ name: cleanLabel(name), shops: cleanLabel(shops) }))
    .filter(({ name, shops }) => name !== '' && shops !== '');
}

// The newest month of the Тосгон workbook (up to today) that holds products.
function readTosgonProducts(book, today) {
  const found = tosgonTabsUpTo(book.getSheets().map((sheet) => sheet.getName()), today)
    .map((name) => ({ name, model: readModel(book.getSheetByName(name)) }))
    .find(({ model }) => model.products.length > 0);
  if (found === undefined) {
    return null;
  }
  const categories = Object.fromEntries(found.model.categories.map((category) => [category.key, cleanLabel(category.label)]));
  return {
    tab: found.name,
    products: found.model.products.map((product) => ({
      name: product.label, // as it stands in the Тосгон sheet, so it can be found there
      category: categories[product.categoryKey] || '',
      price: typeof product.price === 'number' ? product.price : '',
    })),
  };
}

// spreadsheet: this workbook (the list, the table); book: the Тосгон workbook (only read).
function runTosgonNameTable(spreadsheet, book, today) {
  const listSheet = spreadsheet.getSheetByName(PRODUCT_LIST_TAB);
  if (listSheet === null || readListProducts(listSheet).length === 0) {
    return [`ЗОГСЛОО: "${PRODUCT_LIST_TAB}" жагсаалт алга - эхлээд setupProductList`];
  }
  const tosgon = readTosgonProducts(book, today);
  if (tosgon === null) {
    return ['ЗОГСЛОО: Тосгоны файлд нэртэй сарын sheet олдсонгүй'];
  }
  const plan = planTosgonPairs(tosgon.products, readListProducts(listSheet));
  const newCount = plan.rows.filter((row) => row[6] === MATCH_NO_PAIR_LABEL).length;
  const summary =
    `${tosgon.tab}: ${tosgon.products.length} бараа, жагсаалттай ижил ${plan.same}, ` +
    `хос ${plan.rows.length - newCount}, жагсаалтад олдоогүй ${newCount}`;
  const existing = spreadsheet.getSheetByName(TOSGON_DECISION_TAB);
  if (existing !== null && existing.getLastRow() >= DECISION_FIRST_ROW) {
    return [summary, `"${TOSGON_DECISION_TAB}" tab аль хэдийн байна - хөндсөнгүй (дахин гаргах бол устгаад ажиллуулна)`];
  }
  if (DRY_RUN || plan.rows.length === 0) {
    return [summary, `"${TOSGON_DECISION_TAB}" tab: ${plan.rows.length} мөр ${DRY_RUN ? 'бичнэ' : '- бичих зүйлгүй'}`];
  }
  writeTosgonTable(existing || spreadsheet.insertSheet(TOSGON_DECISION_TAB), plan.rows);
  return [summary, `"${TOSGON_DECISION_TAB}" tab: ${plan.rows.length} мөр бичлээ`];
}

function writeTosgonTable(sheet, rows) {
  const [title, header] = TOSGON_TABLE_HEADER;
  sheet.getRange(1, 1).setValue(title);
  sheet.getRange(2, 1, 1, header.length).setValues([header]);
  const body = sheet.getRange(DECISION_FIRST_ROW, 1, rows.length, header.length);
  body.setNumberFormat('@'); // "26-55" stays a name, not a date
  sheet.getRange(DECISION_FIRST_ROW, TOSGON_COLUMNS.price, rows.length, 1).setNumberFormat(TOSGON_PRICE_FORMAT);
  body.setValues(rows);
  sheet
    .getRange(DECISION_FIRST_ROW, TOSGON_ANSWER_COLUMN, rows.length, 1)
    .setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList([DECISION_YES, DECISION_NO], true).setAllowInvalid(false).build()
    );
  styleTosgonTable(sheet, rows.map((row) => row[TOSGON_COLUMNS.kind - 1]));
}

// Colours, widths and formats of the table; kinds: column G of each row, top down.
function styleTosgonTable(sheet, kinds) {
  const width = TOSGON_TABLE_HEADER[1].length;
  const count = kinds.length;
  sheet.getRange(1, 1, 1, width).merge().setBackground(TOSGON_TITLE_TONE).setFontWeight('bold').setWrap(true).setVerticalAlignment('middle');
  sheet.setRowHeight(1, TOSGON_TITLE_HEIGHT);
  sheet.getRange(2, 1, 1, width).setBackground(TOSGON_HEADER_TONE).setFontWeight('bold').setWrap(true).setVerticalAlignment('middle');
  TOSGON_COLUMN_WIDTHS.forEach((pixels, at) => sheet.setColumnWidth(at + 1, pixels));
  sheet.setFrozenRows(2);
  if (count === 0) {
    return;
  }
  const body = sheet.getRange(DECISION_FIRST_ROW, 1, count, width);
  body.setBackgrounds(kinds.map((kind) => new Array(width).fill(tosgonRowTone(kind)))).setVerticalAlignment('middle');
  body.setBorder(true, true, true, true, true, true);
  sheet.getRange(DECISION_FIRST_ROW, TOSGON_COLUMNS.kind, count, 1).setWrap(true);
  sheet.getRange(DECISION_FIRST_ROW, TOSGON_COLUMNS.suggested, count, 1).setFontWeight('bold');
  sheet.getRange(DECISION_FIRST_ROW, TOSGON_COLUMNS.price, count, 1).setNumberFormat(TOSGON_PRICE_FORMAT);
  const corrected = sheet.getRange(DECISION_FIRST_ROW, DECISION_COLUMNS.corrected, count, 1);
  corrected.setBackground(TOSGON_CORRECTED_TONE);
  const answers = sheet.getRange(DECISION_FIRST_ROW, TOSGON_ANSWER_COLUMN, count, 1);
  const column = columnLetter(TOSGON_ANSWER_COLUMN);
  sheet.setConditionalFormatRules(
    Object.entries(TOSGON_ANSWER_TONES).map(([answer, tone]) =>
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=$${column}${DECISION_FIRST_ROW}="${answer}"`)
        .setBackground(tone)
        .setRanges([answers])
        .build()
    )
  );
}

// An answered table made before the colours: rewrites only the title, G (why / % / what
// differs) and D (the price as a number), then colours it. The answers (I, J, K) stay.
function restyleTosgonNameTable() {
  Logger.log('%s\n%s', modeLabel(), runRestyleTosgonTable(currentBook()).join('\n'));
}

// The kind a row had: from its names again, or the contested / no pair mark it was given.
function kindOfTosgonRow(name, listName, oldKind) {
  if (listName === '') {
    return MATCH_NO_PAIR_LABEL;
  }
  if (String(oldKind).startsWith(MATCH_CONTESTED_LABEL)) {
    return MATCH_CONTESTED_LABEL;
  }
  const { level, kind } = pairLevel(name, listName);
  return level > 0 ? kind : String(oldKind).split('\n')[0];
}

function runRestyleTosgonTable(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(TOSGON_DECISION_TAB);
  const height = sheet === null ? 0 : sheet.getLastRow() - DECISION_FIRST_ROW + 1;
  if (height <= 0) {
    return [`ЗОГСЛОО: "${TOSGON_DECISION_TAB}" tab алга эсвэл хоосон - эхлээд makeTosgonNameTable`];
  }
  const rows = sheet.getRange(DECISION_FIRST_ROW, 1, height, DECISION_WIDTH).getValues();
  const cell = (row, column) => cleanLabel(row[column - 1]);
  const kinds = rows.map((row) => {
    if (cell(row, TOSGON_COLUMNS.name) === '') {
      return '';
    }
    const kind = kindOfTosgonRow(cell(row, TOSGON_COLUMNS.name), cell(row, TOSGON_COLUMNS.listName), row[TOSGON_COLUMNS.kind - 1]);
    return kind === MATCH_NO_PAIR_LABEL ? kind : describeTosgonPair(cell(row, TOSGON_COLUMNS.name), cell(row, TOSGON_COLUMNS.listName), kind);
  });
  const prices = rows.map((row) => {
    const text = cell(row, TOSGON_COLUMNS.price).replace(/[\s,₮]/g, '');
    return [text !== '' && Number.isFinite(Number(text)) ? Number(text) : row[TOSGON_COLUMNS.price - 1]];
  });
  const summary = `"${TOSGON_DECISION_TAB}": ${kinds.filter((kind) => kind !== '').length} мөрийг өнгөтэй болгоно (I, J, K-д хүрэхгүй)`;
  if (DRY_RUN) {
    return [summary, ...kinds.filter((kind) => kind !== '').slice(0, 5).map((kind) => kind.replace('\n', ' | '))];
  }
  sheet.getRange(1, 1).setValue(TOSGON_TABLE_HEADER[0]);
  sheet.getRange(DECISION_FIRST_ROW, TOSGON_COLUMNS.kind, height, 1).setValues(kinds.map((kind, at) => [kind || rows[at][TOSGON_COLUMNS.kind - 1]]));
  sheet.getRange(DECISION_FIRST_ROW, TOSGON_COLUMNS.price, height, 1).setNumberFormat(TOSGON_PRICE_FORMAT).setValues(prices);
  styleTosgonTable(sheet, kinds);
  return [summary.replace('болгоно', 'болголоо')];
}

// Run in the Тосгон workbook ("Тосгон 2026"), not in the main one.
function applyTosgonNames() {
  withScriptLock(() => {
    const lines = runApplyTosgonNames(currentBook(), partnerBook(currentBook(), 'main', PRODUCT_LIST_SOURCE_ID));
    Logger.log(lines.join('\n'));
  });
}

// book: the Тосгон workbook; main: the workbook with the list and the answered table (only
// its list gets new rows).
function runApplyTosgonNames(book, main) {
  if (!CHAINS.some((chain) => chain.key === TOSGON_KEY)) {
    return ['ЗОГСЛОО: applyTosgonNames-ийг "Тосгон 2026" файлын Apps Script-д ажиллуулна'];
  }
  const table = main.getSheetByName(TOSGON_DECISION_TAB);
  const list = main.getSheetByName(PRODUCT_LIST_TAB);
  if (table === null || list === null) {
    return [`ЗОГСЛОО: үндсэн файлд "${TOSGON_DECISION_TAB}" эсвэл "${PRODUCT_LIST_TAB}" tab алга`];
  }
  const plan = planNameDecisions(readDecisionRows(table));
  if (plan.undecided.length > 0 || plan.problems.length > 0) {
    return [
      `ЗОГСЛОО, юу ч өөрчлөөгүй: "${TOSGON_DECISION_TAB}" tab-ыг засна уу`,
      ...(plan.undecided.length > 0 ? [`шийдээгүй мөр (тийм/үгүй хоосон): ${plan.undecided.join(', ')}`] : []),
      ...plan.problems,
    ];
  }
  const onList = new Set(readListNames(list).map(({ name }) => normalizeKey(name)));
  const added = plan.chosen.filter((name) => !onList.has(normalizeKey(name)));
  return [
    modeLabel(),
    ...runUnifyNames(book, { [TOSGON_KEY]: plan.fixes[MY_HOME_KEY] }), // column C holds the Тосгон name
    addToProductList(list, added),
    installListCopy(book),
    runInstallNameRules(book),
  ];
}

// New names at the end of the real list, marked as Тосгон's.
function addToProductList(list, names) {
  if (names.length === 0) {
    return `"${PRODUCT_LIST_TAB}" жагсаалтад нэмэх шинэ нэр алга`;
  }
  const quoted = names.map((name) => `"${name}"`).join(', ');
  if (DRY_RUN) {
    return `"${PRODUCT_LIST_TAB}" жагсаалтад нэмнэ (${TOSGON_LIST_LABEL}): ${quoted}`;
  }
  const first = list.getLastRow() + 1;
  const missing = first + names.length - 1 - list.getMaxRows();
  if (missing > 0) {
    list.insertRowsAfter(list.getMaxRows(), missing);
  }
  list
    .getRange(first, 1, names.length, 2)
    .setNumberFormat('@')
    .setValues(names.map((name) => [name, TOSGON_LIST_LABEL]));
  return `"${PRODUCT_LIST_TAB}" жагсаалтад нэмлээ (${TOSGON_LIST_LABEL}): ${quoted}`;
}

// A protected tab "Бараа" whose names come from the real list by IMPORTRANGE: the
// dropdown of this workbook reads it. A tab of that name that is not a copy is left alone.
function installListCopy(book) {
  const existing = book.getSheetByName(PRODUCT_LIST_TAB);
  if (isListCopy(existing)) {
    return `"${PRODUCT_LIST_TAB}" хуулбар аль хэдийн байна`;
  }
  if (existing !== null) {
    return `"${PRODUCT_LIST_TAB}" tab байгаа ч хуулбар биш - хөндсөнгүй`;
  }
  if (DRY_RUN) {
    return `"${PRODUCT_LIST_TAB}" хуулбар (IMPORTRANGE) үүсгэнэ`;
  }
  const sheet = book.insertSheet(PRODUCT_LIST_TAB, 0);
  sheet.getRange(1, 1).setValue(LIST_COPY_TITLE);
  sheet.getRange(3, 1, 1, 2).setValues([PRODUCT_LIST_HEADER[2]]).setFontWeight('bold');
  sheet
    .getRange(PRODUCT_LIST_FIRST_ROW, 1)
    .setFormula(`=IMPORTRANGE("${PRODUCT_LIST_SOURCE_ID}","'${PRODUCT_LIST_TAB}'!A${PRODUCT_LIST_FIRST_ROW}:B")`);
  const protection = sheet.protect().setDescription(LIST_COPY_TITLE);
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
  return (
    `"${PRODUCT_LIST_TAB}" хуулбар үүсгэлээ. Тосгоны файлын "${PRODUCT_LIST_TAB}" tab-ын A4 нүдэн дээр ` +
    'нэг удаа "Allow access" дарна уу, тэгэхгүй бол dropdown хоосон байна'
  );
}
