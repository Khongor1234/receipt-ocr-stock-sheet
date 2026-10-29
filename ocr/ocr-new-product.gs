/**
 * Receipt OCR - a product the workbook has no row for, put in once mom approved its line (the owner,
 * 2026-10-29: a new product, or one never written there before, that mom marks Зөв goes into the
 * sheet under its category, Gemini choosing the category; the opening balance 0; its price from the
 * product list, empty when the list has none; put in in October, it is in November and December too):
 *   - the product list "Бараа" of the year's workbook: a product it lacks gets a row at the end of the
 *     heading Gemini picks among the list's headings, "Аль дэлгүүрт" the place; one it has gets the
 *     place added (Хоёулаа once a second place has it)
 *   - the line's tab and every later tab of the place in the year (ocrTabsFrom): a row at the end of
 *     the category Gemini picks among that sheet's own (a coloured divider, dividers.gs isDividerRow),
 *     with the formats of the category's last product (row-copies.gs clearCopiedMarks), the opening
 *     balance 0 on the line's tab (a later one gets its own carried by the workbook), the cost, price
 *     and shelf price the list shows for the shops (Тосгон's prices are others: empty)
 *   - no category fits, or Gemini does not answer: the small goods' ("Бусад (Жижиг бараа)", ...)
 *   - the price matters (the owner, 2026-10-29: a dear product is no small goods): Gemini is told the
 *     product's price (the list's, else the receipt's) and each category's usual one, and a product
 *     dearer than OCR_SMALL_GOODS_MAX_PRICE goes to the category whose usual price is nearest, not to
 *     the small goods', whatever Gemini said
 *   - a name the list has only in another spelling ("Meiling 254" / "Meiling 254l"): nothing is put
 *     in, the reason names the list's spelling and mom picks it
 * A line mom has not approved waits as before ("бараа ажлын номд алга"). DRY_RUN only logs.
 * The rows below the new one move down: "Бичсэн" of a line written there is found again by its
 * product (ocrCellsNow, ocr-post.gs). Needs dividers.gs and row-copies.gs in the bundle.
 *
 * Gemini's category can be wrong: mom drags the row under the right one, on any tab, and every pass
 * (ocrFollowCategoryMoves) moves the product's row on the tabs after that one the same way (the owner,
 * 2026-10-29: the workbook's own sync carries new and removed products only, not a move). Each product
 * put in keeps the category of its row on every tab (OCR_NEW_PRODUCT_PREFIX), so a tab whose row is
 * under another one now is a tab mom moved it in. The product list keeps its heading.
 */

const OCR_LIST_HEADER_NAME = 'Барааны нэр'; // the list's row of column titles
const OCR_LIST_WIDTH = 5; // A..E: name, place, then the shop's cost, price and shelf price (a view)
const OCR_BOTH_SHOPS_LABEL = 'Хоёулаа';
const OCR_NO_CATEGORY = 'тохирох ангилал алга';
const OCR_SMALL_GOODS_KEYS = ['бусад (жижиг бараа)', 'жижиг бараа']; // else any "... жижиг бараа", else "бусад ..."
const OCR_CATEGORY_THINKING_BUDGET = 256;
const OCR_SMALL_GOODS_MAX_PRICE = 200000; // ₮: a product dearer than this is no small goods
const OCR_NEW_PRODUCT_PREFIX = 'ocr-new-product:'; // + place|name key: { place, name, iso, categories } (a key per tab of ocrTabsFrom)

// ---------- pure helpers (no Drive calls) ----------

// The same product spelt a little differently: no case, spaces, dots, commas or dashes, no litre mark.
const ocrLooseKey = (name) => ocrNameKey(name).replace(/[.,-]/g, '').replace(/[lл]$/, '');
// A category as year-sync.gs and dividers.gs tell them apart: trimmed, single spaces, lower case.
const ocrCategoryKey = (label) => String(label).trim().replace(/\s+/g, ' ').toLowerCase();

// "Аль дэлгүүрт" once the place has the product too.
const ocrListLabel = (current, place) => (ocrText(current) === '' || ocrText(current) === place ? place : OCR_BOTH_SHOPS_LABEL);

const ocrIsSmallGoodsKey = (key) => key.includes('жижиг бараа') || key.startsWith('бусад');

// The category whose usual price ({ key, price }) is nearest a product's, small goods' left out; null
// when none has a price.
function ocrPriceFitKey(categories, price) {
  const priced = categories.filter((category) => typeof category.price === 'number' && category.price > 0 && !ocrIsSmallGoodsKey(category.key));
  const distance = (category) => Math.abs(Math.log(price / category.price));
  return priced.length > 0 ? priced.reduce((best, category) => (distance(category) < distance(best) ? category : best)).key : null;
}

// Each category's usual price, the middle one of its products' ([{ key, label, price }], price null without).
function ocrCategoryUsualPrices(model) {
  return model.categories.map((category) => {
    const prices = model.products.filter((product) => product.categoryKey === category.key && product.price !== null).map((product) => product.price).sort((a, b) => a - b);
    return { key: category.key, label: category.label, price: prices.length > 0 ? prices[Math.floor(prices.length / 2)] : null };
  });
}

// The category a product goes to: Gemini's (else the small goods'), unless the product is dear and that
// is the small goods' category: then the one whose usual price is nearest.
function ocrChooseCategory(answerLabel, model, price) {
  const usual = ocrCategoryUsualPrices(model);
  const chosen = answerLabel ? ocrCategoryKey(answerLabel) : ocrSmallGoodsKey(model.categories);
  const isDear = typeof price === 'number' && price > OCR_SMALL_GOODS_MAX_PRICE;
  return isDear && (chosen === null || ocrIsSmallGoodsKey(chosen)) ? ocrPriceFitKey(usual, price) || chosen : chosen;
}

// The small goods' category among some ({ key }), null when there is none.
function ocrSmallGoodsKey(categories) {
  const keys = categories.map((category) => category.key);
  return OCR_SMALL_GOODS_KEYS.find((key) => keys.includes(key)) ||
    keys.find((key) => key.includes('жижиг бараа')) || keys.find((key) => key.startsWith('бусад')) || null;
}

// The line's tab and every later one of its place in the year: a shop's half-months, Тосгон's months.
function ocrTabsFrom(place, iso) {
  const [year, month, day] = iso.split('-').map(Number);
  if (place === OCR_TOSGON.place) {
    return Array.from({ length: MONTHS_PER_YEAR - month + 1 }, (_, index) => `${OCR_TOSGON.tabPrefix}/${year}/${month + index}`);
  }
  return withChainYear(year, () => {
    const chain = CHAINS.find((candidate) => candidate.label === place);
    const first = periodPosition(year, month, day > FIRST_HALF_LAST_DAY ? 1 : 0);
    return chain ? Array.from({ length: Math.max(0, CHAIN_PERIOD_COUNT - first) }, (_, index) => chainSheetName(chain, first + index)) : [];
  });
}

// Where a new product's row goes: after the last product of its category (the category's own row
// while it has none), after the last product without a category; its formats from that last product.
function ocrNewRowPlace(model, categoryKey) {
  const category = model.categories.find((entry) => entry.key === categoryKey);
  const block = model.products.filter((product) => product.categoryKey === categoryKey);
  const pool = block.length > 0 ? block : model.products;
  const templateRow = pool.length > 0 ? pool[pool.length - 1].row : null;
  const afterRow = category
    ? Math.max(category.row, ...block.map((product) => product.row))
    : Math.max(OCR_BOOK_FIRST_ROW - 1, ...model.products.map((product) => product.row));
  return { afterRow, templateRow };
}

// The last row of a heading's block in the list (the heading itself while empty); the list's last row
// without the heading.
function ocrListBlockEnd(rows, headingKey) {
  const start = rows.findIndex((entry) => entry.isHeading && ocrCategoryKey(entry.name) === headingKey);
  if (start < 0) {
    return Math.max(...rows.map((entry) => entry.row));
  }
  const next = rows.findIndex((entry, index) => index > start && entry.isHeading);
  return Math.max(...rows.slice(start, next < 0 ? rows.length : next).map((entry) => entry.row));
}

function ocrCategoryPrompt(name, sheetLabels, listLabels, listHeading = null, price = null, usualPrices = {}) {
  const shown = (label) => (typeof usualPrices[label] === 'number' ? `"${label}" ~${ocrMoney(usualPrices[label])}` : `"${label}"`);
  return [
    `Монголын цахилгаан бараа, гэр ахуйн барааны дэлгүүрт шинэ бараа ирлээ: "${name}". Нэрээр нь (брэнд, загвар, литр, хэмжээ) ямар бараа болохыг таамаглаад ангилалд нь хуваарил.`,
    ...(typeof price === 'number' ? [`- Үнэ: ${ocrMoney(price)}.`] : []),
    ...(listHeading ? [`- Манай барааны жагсаалтад энэ бараа "${listHeading}" ангилалд байдаг.`] : []),
    `- "sheet_category": ${sheetLabels.map(shown).join(', ')} -аас ЯГ нэгийг ("~" ард нь тэр ангиллын ердийн зарах үнэ).`,
    `- Үнэ чухал: үнэтэй бараа (${ocrMoney(OCR_SMALL_GOODS_MAX_PRICE)}-өөс дээш) жижиг бараа биш, ердийн үнэ нь ойр ангиллыг сонго.`,
    listLabels.length > 0 ? `- "list_category": ${listLabels.map((label) => `"${label}"`).join(', ')} -аас ЯГ нэгийг.` : `- "list_category": "${OCR_NO_CATEGORY}".`,
    `Тохирох ангилал огт байхгүй бол "${OCR_NO_CATEGORY}".`,
  ].join('\n');
}

const ocrCategorySchema = (sheetLabels, listLabels) => ({
  type: 'OBJECT',
  properties: {
    sheet_category: { type: 'STRING', enum: [...sheetLabels, OCR_NO_CATEGORY] },
    list_category: { type: 'STRING', enum: [...listLabels, OCR_NO_CATEGORY] },
  },
  required: ['sheet_category', 'list_category'],
});

// { sheet, list }: the labels Gemini chose, null where it chose none or did not answer.
function ocrParseCategory(code, text, sheetLabels, listLabels) {
  const none = { sheet: null, list: null };
  if (code !== 200) {
    return none;
  }
  try {
    const parts = JSON.parse(text).candidates[0].content.parts.filter((part) => typeof part.text === 'string' && !part.thought);
    const answer = JSON.parse(parts[parts.length - 1].text);
    const known = (value, labels) => (labels.includes(value) ? value : null);
    return { sheet: known(answer.sheet_category, sheetLabels), list: known(answer.list_category, listLabels) };
  } catch (error) {
    return none;
  }
}

// The categories of a product's row on its tabs (null: no row) and the ones kept since: a tab whose
// row is under another category than kept was moved there by mom, and the tabs after it follow it
// until one she moved too. -> { moves: [{ index, to }], recorded } (recorded: as they are once moved)
function ocrCategoryPlan(categories, kept) {
  const { moves, recorded } = categories.reduce((plan, category, index) => {
    if (category === null) {
      return { ...plan, recorded: [...plan.recorded, kept[index] ?? null] };
    }
    if (category !== kept[index]) {
      return { ...plan, current: category, recorded: [...plan.recorded, category] };
    }
    if (plan.current !== null && category !== plan.current) {
      return { ...plan, moves: [...plan.moves, { index, to: plan.current }], recorded: [...plan.recorded, plan.current] };
    }
    return { ...plan, recorded: [...plan.recorded, category] };
  }, { current: null, moves: [], recorded: [] });
  return { moves, recorded };
}

// The heading a product's row of the list is under, null above the first one.
function ocrListHeadingOf(rows, row) {
  const headings = rows.filter((entry) => entry.isHeading && entry.row < row);
  return headings.length > 0 ? headings[headings.length - 1].name : null;
}

// ---------- Drive access ----------

// The cell of a line; a product its tab has no row for is put in first once mom approved the line.
function ocrFindOrPutIn(move, iso, decision, workbooks, record, deadline = Infinity) {
  const target = ocrFindCell(move, iso, decision.name, workbooks);
  if (target.ok || !decision.approved || !OCR_NO_ROW_REASON.test(target.reason)) {
    return target;
  }
  if (Date.now() > deadline) {
    return { ...target, deferred: true }; // its rows take seconds a tab: the next run puts them in
  }
  const put = ocrPutProductIn(move, iso, decision.name, workbooks, ocrReceiptPriceOf(record));
  if (put.ok) {
    return ocrFindCell(move, iso, decision.name, workbooks);
  }
  return put.reason ? { ...target, reason: `${target.reason}; ${put.reason}` } : target;
}

// { ok: true } once the product's rows are in; { ok: false, reason } ('' when nothing more is to say).
function ocrPutProductIn(move, iso, name, workbooks, receiptPrice = null) {
  try {
    return ocrPutProductInNow(move, iso, name, workbooks, receiptPrice);
  } catch (error) {
    Logger.log(`  шинэ бараа "${name}" (${move.place}): нэмж чадсангүй: ${error.message}`);
    return { ok: false, reason: `шинэ бараа нэмж чадсангүй (${error.message})` };
  }
}

function ocrPutProductInNow(move, iso, name, workbooks, receiptPrice) {
  const target = ocrTarget(move, iso);
  if (!target) {
    return { ok: false, reason: '' };
  }
  const list = ocrReadList(SpreadsheetApp.openById(ocrCatalogBookId(Number(iso.slice(0, 4)))));
  const key = ocrNameKey(name);
  const products = list ? list.rows.filter((entry) => !entry.isHeading) : [];
  const listed = products.find((entry) => entry.key === key) || null;
  const twin = listed ? null : products.find((entry) => ocrLooseKey(entry.name) === ocrLooseKey(name));
  if (twin) {
    return { ok: false, reason: `жагсаалтад "${twin.name}" гэж бий: тэрийг сонгоно уу` };
  }
  const bookId = ocrBookIdFor(move.place, target.tab);
  const book = SpreadsheetApp.openById(bookId);
  const tabs = ocrTabsFrom(move.place, iso).map((tabName) => book.getSheetByName(tabName)).filter(Boolean);
  if (tabs.length === 0 || tabs[0].getName() !== target.tab) {
    return { ok: false, reason: '' };
  }
  if (DRY_RUN) {
    Logger.log(`  шинэ бараа "${name}" (${move.place}): ${tabs.map((sheet) => sheet.getName()).join(', ')}-д мөр нэмнэ (туршилт)`);
    return { ok: false, reason: '' };
  }
  const model = ocrReadTabModel(tabs[0]);
  const listHeadings = list && !listed ? [...new Set(list.rows.filter((entry) => entry.isHeading).map((entry) => entry.name))] : [];
  // the list's heading of a product it has, when the sheet has a category of that name: no question
  const listHeading = listed ? ocrListHeadingOf(list.rows, listed.row) : null;
  const sameCategory = listHeading ? model.categories.find((category) => category.key === ocrCategoryKey(listHeading)) : null;
  const listPrice = listed && typeof listed.prices[1] === 'number' && listed.prices[1] > 0 ? listed.prices[1] : null;
  const price = listPrice || receiptPrice;
  const usualPrices = Object.fromEntries(ocrCategoryUsualPrices(model).filter((usual) => usual.price !== null).map((usual) => [usual.label, usual.price]));
  const answer = sameCategory
    ? { sheet: sameCategory.label, list: null }
    : ocrAskCategory(name, [...new Set(model.categories.map((category) => category.label))], listHeadings, { listHeading, price, usualPrices });
  if (list) {
    ocrKeepInList(list, listed, name, move.place, answer.list);
  }
  const chosenKey = sameCategory ? sameCategory.key : ocrChooseCategory(answer.sheet, model, price);
  const isShop = move.place !== OCR_TOSGON.place;
  const prices = isShop && listed ? listed.prices.map((price) => (typeof price === 'number' ? price : '')) : [];
  const width = target.layout.shelfPrice || OCR_CATALOG_PRICE_COLUMN; // A..D, or A..E with the shelf price
  const categories = tabs.map((sheet, index) => {
    const tabModel = index === 0 ? model : ocrReadTabModel(sheet);
    const there = tabModel.products.find((product) => product.key === key);
    if (there) {
      return [sheet.getName(), there.categoryKey];
    }
    const categoryKey = tabModel.categories.some((category) => category.key === chosenKey) ? chosenKey : ocrSmallGoodsKey(tabModel.categories);
    const values = [name, index === 0 ? 0 : '', ...prices, '', '', ''].slice(0, isShop ? width : 2);
    const row = ocrInsertProductRow(sheet, tabModel, categoryKey, values, target.layout.count);
    Logger.log(`  шинэ бараа "${name}": ${sheet.getName()} мөр ${row}, ангилал "${categoryKey || 'алга'}"`);
    workbooks.delete(`${bookId}|${sheet.getName()}`);
    return [sheet.getName(), categoryKey];
  });
  const byTab = new Map(categories);
  const props = PropertiesService.getScriptProperties();
  props.setProperty(`${OCR_NEW_PRODUCT_PREFIX}${move.place}|${key}`,
    JSON.stringify({ place: move.place, name, iso, categories: ocrTabsFrom(move.place, iso).map((tabName) => byTab.get(tabName) ?? null) }));
  // the rows below moved down: the months' "Бичсэн" and Тулгалт are named again soon (ocr-relink.gs)
  ocrRelinkSoon(props, [...new Set(tabs.map((sheet) => sheet.getName().split('/').slice(1, 3).map((part) => part.padStart(2, '0')).join('-')))]);
  return { ok: true };
}

// Mom's moves of the products put in, carried to the tabs after (see above): run from the editor, and
// by every review pass (runReviewPass). -> the rows moved
function followCategoryMovesNow() {
  return ocrFollowCategoryMoves(PropertiesService.getScriptProperties(), Date.now() + OCR_TIME_BUDGET_MS);
}

function ocrFollowCategoryMoves(props, deadline) {
  const kept = Object.entries(props.getProperties()).filter(([key]) => key.startsWith(OCR_NEW_PRODUCT_PREFIX));
  const models = new Map(); // "book|tab" -> { sheet, model }, read once a run
  let moved = 0;
  for (const [key, value] of kept) {
    if (Date.now() > deadline) {
      break;
    }
    try {
      const entry = JSON.parse(value);
      const result = ocrFollowProduct(entry, models);
      moved += result.moved;
      if (!DRY_RUN && JSON.stringify(result.recorded) !== JSON.stringify(entry.categories)) {
        props.setProperty(key, JSON.stringify({ ...entry, categories: result.recorded }));
      }
    } catch (error) {
      Logger.log(`Шинэ барааны ангилал (${key}): ${error.message}`);
    }
  }
  if (moved > 0) {
    Logger.log(`Шинэ барааны ангилал (${modeLabel()}): ээжийн зөөснөөр ${moved} мөр ${DRY_RUN ? 'зөөнө' : 'зөөлөө'}`);
  }
  return moved;
}

// One product: its row on the tabs after mom's move moved under her category. -> { moved, recorded }
function ocrFollowProduct(entry, models) {
  const tabNames = ocrTabsFrom(entry.place, entry.iso);
  const bookId = ocrBookIdFor(entry.place, tabNames[0]);
  const key = ocrNameKey(entry.name);
  const read = (tabName) => {
    const id = `${bookId}|${tabName}`;
    if (!models.has(id)) {
      const sheet = SpreadsheetApp.openById(bookId).getSheetByName(tabName);
      models.set(id, sheet ? { sheet, model: ocrReadTabModel(sheet) } : null);
    }
    return models.get(id);
  };
  const found = tabNames.map((tabName) => {
    const tab = read(tabName);
    const product = tab ? tab.model.products.find((entryOfTab) => entryOfTab.key === key) : null;
    return product ? { ...tab, tabName, product } : null;
  });
  const plan = ocrCategoryPlan(found.map((tab) => (tab ? tab.product.categoryKey : null)), entry.categories || []);
  const done = plan.moves.filter(({ index, to }) => {
    const tab = found[index];
    if (!tab.model.categories.some((category) => category.key === to)) {
      return false; // that tab has no such category: it stays
    }
    Logger.log(`  "${entry.name}": ${tab.tabName}-д "${to}" ангилал руу${DRY_RUN ? ' зөөнө (туршилт)' : ''}`);
    if (!DRY_RUN) {
      ocrMoveProductRow(tab.sheet, tab.model, tab.product.row, to);
      models.delete(`${bookId}|${tab.tabName}`);
    }
    return true;
  });
  const recorded = plan.recorded.map((category, index) =>
    (plan.moves.some((move) => move.index === index) && !done.some((move) => move.index === index) ? found[index].product.categoryKey : category));
  return { moved: done.length, recorded };
}

// A product's row moved to the end of a category, all it holds with it.
function ocrMoveProductRow(sheet, model, row, categoryKey) {
  const { afterRow } = ocrNewRowPlace({ ...model, products: model.products.filter((product) => product.row !== row) }, categoryKey);
  withFilterLifted(sheet, () => sheet.moveRows(sheet.getRange(row, 1), afterRow + 1)); // counted as before the move
}

// The receipt price read for a line (O of its row in the month spreadsheet), null without one.
function ocrReceiptPriceOf(record) {
  const at = record && record._at;
  if (!at) {
    return null;
  }
  try {
    const sheet = SpreadsheetApp.openById(at.book).getSheets().find((tab) => tab.getSheetId() === at.gid);
    const price = sheet ? sheet.getRange(at.row, OCR_READ_PRICE_COLUMN).getValue() : null;
    return typeof price === 'number' && price > 0 ? price : null;
  } catch (error) {
    return null;
  }
}

// The product list of a year's workbook: { sheet, rows: [{ row, name, key, place, prices, isHeading }] },
// null without one. A heading is a row without a place.
function ocrReadList(book) {
  const sheet = book.getSheetByName(OCR_PRODUCT_LIST_TAB);
  const count = sheet ? sheet.getLastRow() - OCR_CATALOG_FIRST_ROW + 1 : 0;
  if (count <= 0) {
    return null;
  }
  const rows = sheet.getRange(OCR_CATALOG_FIRST_ROW, 1, count, OCR_LIST_WIDTH).getValues()
    .map((values, index) => ({ row: OCR_CATALOG_FIRST_ROW + index, name: ocrCleanName(values[0] ?? ''), place: ocrText(values[1]), prices: values.slice(2) }))
    .filter((entry) => entry.name !== '' && entry.name !== OCR_LIST_HEADER_NAME)
    .map((entry) => ({ ...entry, key: ocrNameKey(entry.name), isHeading: entry.place === '' }));
  return { sheet, rows };
}

// The list after the product went in: its row, or the place added to the row it has.
function ocrKeepInList(list, listed, name, place, heading) {
  if (listed) {
    const label = ocrListLabel(listed.place, place);
    if (label !== listed.place) {
      list.sheet.getRange(listed.row, OCR_PRODUCT_LIST_SHOP_INDEX + 1).setValue(label);
    }
    return;
  }
  const headings = list.rows.filter((entry) => entry.isHeading).map((entry) => ({ key: ocrCategoryKey(entry.name) }));
  const after = ocrListBlockEnd(list.rows, heading ? ocrCategoryKey(heading) : ocrSmallGoodsKey(headings));
  list.sheet.insertRowAfter(after);
  list.sheet.getRange(after + 1, 1, 1, 2).setValues([[name, place]]);
  Logger.log(`  шинэ бараа "${name}": "${OCR_PRODUCT_LIST_TAB}" жагсаалтын мөр ${after + 1}`);
}

// A workbook tab's categories (coloured dividers, as year-sync.gs reads them) and products.
function ocrReadTabModel(sheet) {
  const height = sheet.getLastRow() - OCR_BOOK_FIRST_ROW + 1;
  if (height <= 0) {
    return { products: [], categories: [] };
  }
  const values = sheet.getRange(OCR_BOOK_FIRST_ROW, 1, height, DIVIDER_VALUE_WIDTH).getValues();
  const colors = sheet.getRange(OCR_BOOK_FIRST_ROW, 1, height, DIVIDER_SCAN_WIDTH).getBackgrounds();
  return values.reduce((model, cells, index) => {
    const label = ocrText(cells[0]);
    if (label === '') {
      return model;
    }
    const row = OCR_BOOK_FIRST_ROW + index;
    if (isDividerRow(colors[index], cells)) {
      return { ...model, categories: [...model.categories, { key: ocrCategoryKey(label), label, row }] };
    }
    const category = model.categories[model.categories.length - 1];
    const price = cells[OCR_CATALOG_PRICE_COLUMN - 1];
    return { ...model, products: [...model.products, { key: ocrNameKey(label), row, categoryKey: category ? category.key : null, price: typeof price === 'number' && price > 0 ? price : null }] };
  }, { products: [], categories: [] });
}

// A row put in under a category (year-sync.gs insertProduct does the same): the formats of the
// category's last product, none of its numbers or notes. -> the new row's number
function ocrInsertProductRow(sheet, model, categoryKey, values, width) {
  const { afterRow, templateRow } = ocrNewRowPlace(model, categoryKey);
  const row = afterRow + 1;
  withFilterLifted(sheet, () => {
    sheet.insertRowAfter(afterRow);
    if (templateRow !== null) {
      const from = templateRow > afterRow ? templateRow + 1 : templateRow;
      sheet.getRange(from, 1, 1, width).copyTo(sheet.getRange(row, 1, 1, width));
      clearCopiedMarks(sheet, row, width);
    }
    sheet.getRange(row, 1, 1, width).clearContent();
    sheet.getRange(row, 1, 1, values.length).setValues([values]);
  });
  return row;
}

// The categories Gemini chooses for a product: { sheet, list } labels, null where none.
function ocrAskCategory(name, sheetLabels, listLabels, { listHeading = null, price = null, usualPrices = {} } = {}) {
  const key = PropertiesService.getScriptProperties().getProperty(OCR_KEY_PROPERTY);
  if (!key || sheetLabels.length === 0) {
    return { sheet: null, list: null };
  }
  const body = {
    contents: [{ role: 'user', parts: [{ text: ocrCategoryPrompt(name, sheetLabels, listLabels, listHeading, price, usualPrices) }] }],
    generationConfig: {
      temperature: 0,
      thinkingConfig: { thinkingBudget: OCR_CATEGORY_THINKING_BUDGET },
      responseMimeType: 'application/json',
      responseSchema: ocrCategorySchema(sheetLabels, listLabels),
    },
  };
  try {
    const response = UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1beta/models/${OCR_MODEL}:generateContent`, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': key },
      payload: JSON.stringify(body),
      muteHttpExceptions: true,
    });
    const answer = ocrParseCategory(response.getResponseCode(), response.getContentText(), sheetLabels, listLabels);
    Logger.log(`  шинэ бараа "${name}": Gemini-ийн ангилал "${answer.sheet || OCR_NO_CATEGORY}"`);
    return answer;
  } catch (error) {
    Logger.log(`  шинэ бараа "${name}": Gemini хариулсангүй (${error.message}), жижиг бараанд`);
    return { sheet: null, list: null };
  }
}
