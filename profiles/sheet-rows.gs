/**
 * Reads the product rows of a sheet into plain objects. Used by legacy-months.gs (the old
 * half-month sheets, 15 day columns) and by site-feed.gs (the month sheets, 30 day columns).
 *
 * A row with a name and no number anywhere is a category; the products below it belong to it.
 * layout: the column layout of the sheet (config.gs buildColumnLayout). offset: how many
 * columns every column after A stands to the right of where it should (one old sheet has an
 * extra empty column B).
 */

const sumNumbers = (values) => values.reduce((sum, value) => sum + (typeof value === 'number' ? value : 0), 0);

const orBlank = (value) => (value === undefined || value === null ? '' : value);

function parseSheetRows(values, layout, offset = 0) {
  const at = (row, column) => orBlank(row[column - 1 + (column >= 2 ? offset : 0)]);
  const slice = (row, first, last) => Array.from({ length: last - first + 1 }, (_, step) => at(row, first + step));
  const rows = [];
  let category = { key: null, label: '' };
  values.forEach((row) => {
    const label = String(row[0]).trim();
    if (label === '') {
      return;
    }
    const entry = {
      key: normalizeKey(label),
      label,
      categoryKey: category.key,
      category: category.label,
      opening: at(row, 2),
      cost: at(row, 3),
      price: at(row, 4),
      shelfPrice: layout.shelfPrice ? at(row, layout.shelfPrice) : '', // the shops' layout only
      income: slice(row, layout.incomeFirst, layout.transferIn - 1),
      transferIn: at(row, layout.transferIn),
      expense: slice(row, layout.expenseFirst, layout.transferOut - 1),
      transferOut: at(row, layout.transferOut),
      gift: layout.gift ? at(row, layout.gift) : '', // the shops' gift column only
      count: at(row, layout.count),
      balance: at(row, layout.balance),
    };
    const hasNumber = [entry.opening, entry.cost, entry.price, entry.count, entry.transferIn, entry.transferOut, entry.gift, ...entry.income, ...entry.expense]
      .some((value) => typeof value === 'number');
    if (hasNumber) {
      rows.push(entry);
    } else {
      category = { key: entry.key, label };
    }
  });
  return rows;
}
