/**
 * Gives one product ONE name on every sheet of the year.
 *
 * Before the syncs existed the product list was typed again for every half-month,
 * so the same product is spelled differently from sheet to sheet ("Miling 201l" /
 * "Meiling 201l", "8037 гы" / "8037 us" / "8037 ус"). The scripts know a product by
 * its name only: a differently spelled row is another product for them, and its
 * opening balance is not carried from the count before it.
 *
 * NAME_FIXES lists, per shop, the spelling to keep (the one of the newest sheet,
 * 2026-10-13) and the other spellings found in the year; the owner confirmed the
 * doubtful ones. unifyNames writes the kept spelling over the others on every
 * sheet, then saves the product and price lists anew (_sync, _prices), so the
 * syncs see no product "removed" and "added". A sheet that already holds the kept
 * spelling as well is left alone and reported: there the two are different rows.
 *
 * The hard rule: NO TWO ROWS OF ONE NAME on a sheet. Every sync knows a product by
 * its name, so two rows called "гадах гэрэл" cannot be told apart. The second row
 * of a name gets a number - "гадах гэрэл (2)", then "(3)" - and keeps that same
 * name on every later sheet: rows are followed from sheet to sheet by the nearest
 * price (pairProducts in year-sync.gs). unifyNames numbers the whole year after
 * the spellings.
 * From then on a name that is already on the sheet cannot be typed at all: column
 * A (row 4 down) has a rule that refuses it with a message (installNameRule, put
 * on every sheet by setupYear). What still gets past the rule - a pasted block
 * can carry its own rule - is numbered at once when it is one cell
 * (instant-sync.gs), and by the hourly job otherwise.
 * Once the product list "Бараа" exists (product-list.gs) column A holds the dropdown
 * of that list instead; a repeated name is then only numbered, not refused.
 *
 * Usage: run unifyNames with DRY_RUN = true (config.gs) and read the log, then
 * with false. Afterwards run carryOpeningBalances (or wait for the hourly job):
 * the opening balances of the renamed products follow the counts before them.
 * Needs config.gs, year-sync.gs, price-sync.gs.
 */

const NAME_FIRST_NUMBER = 2; // the first row keeps the plain name
const NAME_FIRST_ROW = 4;
// True while the name of the row stands once in A4:A, whatever its capitals and spaces.
const NAME_RULE_FORMULA = '=SUMPRODUCT(--(TRIM(LOWER($A$4:$A))=TRIM(LOWER($A4))))<=1';
const NAME_RULE_MESSAGE =
  'Ийм нэртэй бараа энэ sheet дээр аль хэдийн байна. Өөр нэр бичнэ үү (жишээ нь ард нь 2 гэж нэмнэ).';

function unifyNames() {
  withScriptLock(() => {
    const spreadsheet = currentBook();
    // In a dry run the numbering is planned on the names as they are now.
    const lines = [...runUnifyNames(spreadsheet), ...runNumberRepeated(spreadsheet)];
    Logger.log('%s\n%s', modeLabel(), lines.join('\n'));
  });
}

// ---------- pure helpers (no Spreadsheet calls) ----------

// { key of another spelling: the spelling to keep }
function buildNameMap(fixes) {
  return fixes.reduce(
    (all, [kept, others]) => others.reduce((map, other) => ({ ...map, [normalizeKey(other)]: kept }), all),
    {}
  );
}

// What to do on one sheet. products: [{ key, label, row }] (readModel).
// Returns { renames: [{ row, from, to }], clashes: [label] }.
function planRenames(products, nameMap) {
  const present = new Set(products.map((product) => product.key));
  const wanted = products.filter((product) => nameMap[product.key] !== undefined && product.label !== nameMap[product.key]);
  // Free: the kept spelling is not on the sheet (or is this row's own, written in other capitals or
  // spaces), and no earlier row is about to get it.
  const isFree = (product) =>
    (!present.has(normalizeKey(nameMap[product.key])) || normalizeKey(nameMap[product.key]) === product.key) &&
    wanted.find((other) => nameMap[other.key] === nameMap[product.key]) === product;
  return {
    renames: wanted
      .filter(isFree)
      .map((product) => ({ row: product.row, from: product.label, to: nameMap[product.key] })),
    clashes: wanted.filter((product) => !isFree(product)).map((product) => product.label),
  };
}

function numberedName(label, number) {
  return `${String(label).trim()} (${number})`;
}

// The label itself when no other row has it, otherwise the label with the first free number.
function freeName(label, takenKeys) {
  const taken = new Set(takenKeys);
  if (!taken.has(normalizeKey(label))) {
    return label;
  }
  let number = NAME_FIRST_NUMBER;
  while (taken.has(normalizeKey(numberedName(label, number)))) {
    number += 1;
  }
  return numberedName(label, number);
}

// The names the rows of a sheet should have so that none stands twice.
// previous / previousNames: the rows of the sheet before as they were read, and
// the names they were given. A row that continues one of them takes its name;
// the others keep theirs when it is free and get a number when it is not.
function planNumbering(previous, previousNames, current) {
  const pairs = pairProducts(previous, current);
  const inherited = current.map((product, at) => {
    if (pairs[at] === -1) {
      return null;
    }
    const name = previousNames[pairs[at]];
    return normalizeKey(name) === product.key ? product.label : name; // same name: the spelling stays
  });
  const taken = inherited.filter((name) => name !== null).map(normalizeKey);
  return current.reduce(
    (plan, product, at) => {
      const name = inherited[at] !== null ? inherited[at] : freeName(product.label, plan.taken);
      return { names: [...plan.names, name], taken: [...plan.taken, normalizeKey(name)] };
    },
    { names: [], taken }
  ).names;
}

// ---------- spreadsheet access ----------

// The rule of column A: the dropdown of the product list (product-list.gs) once the
// list exists, the "no name twice" rule before.
function buildNameRule(spreadsheet) {
  const list = productListRange(spreadsheet);
  const builder = SpreadsheetApp.newDataValidation();
  if (list === null) {
    return builder.requireFormulaSatisfied(NAME_RULE_FORMULA).setAllowInvalid(false).setHelpText(NAME_RULE_MESSAGE).build();
  }
  return builder.requireValueInRange(list, true).setAllowInvalid(!PRODUCT_LIST_STRICT).setHelpText(PRODUCT_LIST_HELP).build();
}

// What a rule checks, to tell whether the wanted one is already there. The height of
// the list counts: rows added at the bottom of the list must reach the dropdown.
function nameRuleSignature(rule) {
  const [source] = rule.getCriteriaValues();
  const isRange = source !== null && typeof source === 'object' && typeof source.getSheet === 'function';
  return isRange
    ? `${source.getSheet().getName()}!${source.getRow()}:${source.getNumRows()}|${rule.getAllowInvalid()}`
    : String(source);
}

// Puts the rule on column A. Returns true when it was missing.
function installNameRule(sheet) {
  const rows = sheet.getMaxRows() - NAME_FIRST_ROW + 1;
  if (rows <= 0) {
    return false;
  }
  const wanted = buildNameRule(sheet.getParent());
  const current = sheet.getRange(NAME_FIRST_ROW, SYNC_NAME_COL).getDataValidation();
  if (current !== null && nameRuleSignature(current) === nameRuleSignature(wanted)) {
    return false;
  }
  if (!DRY_RUN) {
    sheet.getRange(NAME_FIRST_ROW, SYNC_NAME_COL, rows, 1).setDataValidation(wanted);
  }
  return true;
}

// fromPosition (optional): only the half-months from this one on (config.gs sheetSlot).
function runNumberRepeated(spreadsheet, fromPosition = 0) {
  const counts = {};
  findChains(spreadsheet).forEach(({ label, sheets }) => {
    let previous = [];
    let previousNames = [];
    sheets
      .filter((sheet) => sheetSlot(sheet.getName()).position >= fromPosition)
      .forEach((sheet) => {
        const products = readModel(sheet).products;
        const names = planNumbering(previous, previousNames, products);
        products.forEach((product, at) => {
          if (normalizeKey(names[at]) === product.key) {
            return;
          }
          if (!DRY_RUN) {
            sheet.getRange(product.row, SYNC_NAME_COL).setValue(names[at]);
          }
          const id = `${label}: ${product.label} → ${names[at]}`;
          counts[id] = (counts[id] || 0) + 1;
        });
        previous = products;
        previousNames = names;
      });
  });
  const total = Object.values(counts).reduce((sum, count) => sum + count, 0);
  if (total === 0) {
    return [];
  }
  if (!DRY_RUN) {
    saveListsAnew(spreadsheet);
  }
  return [
    `ижил нэртэй мөрийг дугаарлах: ${total} нүд ${DRY_RUN ? 'солигдоно' : 'солигдлоо'}`,
    ...Object.entries(counts).map(([id, count]) => `  ${id} (${count} sheet)`),
  ];
}

// fixes (optional): another table than NAME_FIXES.
function runUnifyNames(spreadsheet, fixes = NAME_FIXES) {
  const counts = {};
  const clashes = [];
  findChains(spreadsheet).forEach(({ label, sheets }) => {
    const nameMap = buildNameMap(fixes[label] || []);
    sheets.forEach((sheet) => {
      const plan = planRenames(readModel(sheet).products, nameMap);
      plan.renames.forEach((rename) => {
        if (!DRY_RUN) {
          sheet.getRange(rename.row, SYNC_NAME_COL).setValue(rename.to);
        }
        const id = `${label}: ${rename.from} → ${rename.to}`;
        counts[id] = (counts[id] || 0) + 1;
      });
      plan.clashes.forEach((name) => clashes.push(`${sheet.getName()}: "${name}" хэвээр (хадгалах нэртэй мөр энэ sheet дээр бас байна)`));
    });
  });
  const total = Object.values(counts).reduce((sum, count) => sum + count, 0);
  if (!DRY_RUN && total > 0) {
    saveListsAnew(spreadsheet);
  }
  return [
    `нэр нэгтгэх: ${total} нүд ${DRY_RUN ? 'солигдоно' : 'солигдлоо'}`,
    ...Object.entries(counts).map(([id, count]) => `  ${id} (${count} sheet)`),
    ...clashes,
  ];
}

// The lists the syncs compare with are taken from the sheets as they are now, so
// the new names are nobody's "added" or "removed" product and no price "changed".
function saveListsAnew(spreadsheet) {
  const sheets = findChains(spreadsheet).flatMap((chain) => chain.sheets);
  const lists = sheets.reduce(
    (all, sheet) => ({ ...all, [sheet.getName()]: readModel(sheet).products.map((product) => product.key) }),
    loadSnapshots(spreadsheet)
  );
  const prices = sheets.reduce(
    (all, sheet) => ({ ...all, [sheet.getName()]: readPrices(sheet).values }),
    loadSnapshots(spreadsheet, PRICE_STATE_SHEET)
  );
  saveSnapshots(spreadsheet, lists);
  saveSnapshots(spreadsheet, prices, PRICE_STATE_SHEET);
}
