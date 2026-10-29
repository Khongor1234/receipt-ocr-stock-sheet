/**
 * Keeps the tabs of the year sheets in a fixed order.
 *
 * Google Sheets cannot lock the position of a tab: anyone who may edit can drag
 * it, and rename it. So both are restored instead: setupYear arranges the tabs,
 * a dragged or renamed tab is put back at once (instant-sync.gs), and the
 * background job (year-sync.gs) does it again every hour.
 *
 * A year sheet is known by its tab name, so a renamed one would drop out of the
 * order (and out of the year). The name of every year sheet is therefore
 * remembered by sheet id (document property TAB_NAMES_KEY), and a year sheet that
 * no longer carries a year name gets its remembered name back.
 *
 * TAB_ORDER_MODE
 *   'pairs'   one period after the other, both shops next to each other:
 *             Соронзон 1/1-15, My home 1/1-15, Соронзон 1/16-31, My home 1/16-31, ...
 *   'blocks'  the whole year of one shop, then the whole year of the other:
 *             Соронзон January ... December, then My home January ... December
 * TAB_ORDER_FIRST_CHAIN is the shop that comes first ('соронзон' or 'my home').
 *
 * Sheets that are not year sheets (Sheet200, Sheet100, ...) keep their side:
 * those standing before most of the year sheets stay in front, the others go last.
 * Needs config.gs.
 */

const TAB_ORDER_MODE = 'pairs';
const TAB_ORDER_FIRST_CHAIN = 'соронзон';
const TAB_NAMES_KEY = 'tabNames'; // JSON: { sheet id: tab name of that year sheet }

function arrangeTabs() {
  const lines = runArrangeTabs(currentBook());
  Logger.log('%s\n%s', modeLabel(), lines.join('\n'));
}

// ---------- pure helpers (no Spreadsheet calls) ----------

function tabSortKey(name) {
  const slot = sheetSlot(name);
  const chainRank = slot.chainKey === TAB_ORDER_FIRST_CHAIN ? 0 : 1;
  return TAB_ORDER_MODE === 'blocks'
    ? chainRank * CHAIN_PERIOD_COUNT + slot.position
    : slot.position * CHAINS.length + chainRank;
}

function desiredTabOrder(names) {
  const isYearSheet = (name) => sheetSlot(name) !== null;
  const yearNames = names.filter(isYearSheet);
  // A sheet counts as "in front" while most year sheets come after it, so one
  // year tab dragged to the very front does not push Sheet200 to the back.
  const isInFront = (name) =>
    names.slice(0, names.indexOf(name)).filter(isYearSheet).length < yearNames.length / 2;
  const others = names.filter((name) => !isYearSheet(name));
  const year = [...yearNames].sort((a, b) => tabSortKey(a) - tabSortKey(b));
  return [...others.filter(isInFront), ...year, ...others.filter((name) => !isInFront(name))];
}

// The tabs that have to be moved, as { name, toIndex } in the order to apply them.
function planTabMoves(names) {
  const wanted = desiredTabOrder(names);
  const moves = [];
  let current = names;
  wanted.forEach((name, index) => {
    if (current[index] === name) {
      return;
    }
    moves.push({ name, toIndex: index });
    const without = current.filter((other) => other !== name);
    current = [...without.slice(0, index), name, ...without.slice(index)];
  });
  return moves;
}

// tabs: [{ id, name }] as they are now. saved: { id: year name } from the last run.
// Returns the renames to make and the names to remember from now on.
function planTabRenames(tabs, saved) {
  const taken = new Set(tabs.map((tab) => tab.name));
  return tabs.reduce(
    (plan, tab) => {
      if (sheetSlot(tab.name) !== null) {
        return { ...plan, names: { ...plan.names, [tab.id]: tab.name } };
      }
      const wanted = saved[tab.id];
      if (wanted === undefined || taken.has(wanted)) {
        return plan; // never was a year sheet, or its old name is in use
      }
      taken.add(wanted);
      return {
        renames: [...plan.renames, { id: tab.id, from: tab.name, to: wanted }],
        names: { ...plan.names, [tab.id]: wanted },
      };
    },
    { renames: [], names: {} }
  );
}

// ---------- spreadsheet access ----------

function readSavedTabNames() {
  const saved = bookProperties().getProperty(TAB_NAMES_KEY);
  try {
    const parsed = saved ? JSON.parse(saved) : {};
    return parsed !== null && typeof parsed === 'object' ? parsed : {};
  } catch (error) {
    return {}; // unreadable note: start remembering again
  }
}

// Gives a renamed year sheet its name back. Returns the log lines.
function restoreTabNames(spreadsheet) {
  const sheets = spreadsheet.getSheets();
  const saved = readSavedTabNames();
  const plan = planTabRenames(
    sheets.map((sheet) => ({ id: String(sheet.getSheetId()), name: sheet.getName() })),
    saved
  );
  if (!DRY_RUN) {
    plan.renames.forEach((rename) => {
      sheets.find((sheet) => String(sheet.getSheetId()) === rename.id).setName(rename.to);
    });
    if (JSON.stringify(plan.names) !== JSON.stringify(saved)) {
      bookProperties().setProperty(TAB_NAMES_KEY, JSON.stringify(plan.names));
    }
  }
  return plan.renames.map(
    (rename) => `"${rename.from}" табын нэрийг "${rename.to}" болгож ${DRY_RUN ? 'сэргээнэ' : 'сэргээлээ'}`
  );
}

function runArrangeTabs(spreadsheet) {
  const renamed = restoreTabNames(spreadsheet);
  const names = spreadsheet.getSheets().map((sheet) => sheet.getName());
  const moves = planTabMoves(names);
  if (moves.length === 0) {
    return [...renamed, 'табын дараалал зөв байна'];
  }
  if (!DRY_RUN) {
    moves.forEach((move) => {
      spreadsheet.setActiveSheet(spreadsheet.getSheetByName(move.name));
      spreadsheet.moveActiveSheet(move.toIndex + 1);
    });
  }
  return [
    ...renamed,
    `${moves.length} табыг байранд нь ${DRY_RUN ? 'зөөнө' : 'зөөлөө'}: ${moves.map((move) => move.name).join(', ')}`,
  ];
}
