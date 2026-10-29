/**
 * Charts on the yearly summary sheet (Sheet200).
 *
 * Sheet200 holds the months across (C5:N5) and one line per figure (rows 6-12).
 * Charts read columns more reliably than rows, so the figures are first laid out
 * in columns on a hidden helper sheet "_graph", and the charts are drawn from
 * there. The helper cells are FORMULAS on Sheet200, so every chart follows the
 * numbers by itself: nothing has to be run again when a new month is entered.
 * One chart is kept: the monthly sales and profit of each shop, the amounts in tugrik
 * on the vertical axis. (The year-so-far, monthly-profit and profit-share charts
 * of an earlier version are removed.) The chart is placed under the table of
 * Sheet200. Also fills the empty yearly total column "Нийт" (O) and the missing
 * label of the profit-share row, adds one line per shop that names the seller of
 * each month, and colours the table: one colour per shop.
 *
 * installCharts is run by setupYear; run it by hand to put back a deleted chart.
 * A chart that already exists (same title) is left alone, so it may be moved or
 * resized freely. DRY_RUN (config.gs) makes it report only.
 */

const CHART_SOURCE_SHEET = 'Sheet200';
const CHART_DATA_SHEET = '_graph';
const CHART_MONTH_COUNT = 12;
const CHART_MONTHS_RANGE = 'C5:N5';
const CHART_FIRST_ANCHOR_ROW = 15; // first row under the table of Sheet200
const CHART_ROWS_PER_CHART = 19;
const CHART_ANCHOR_COLUMN = 1;
const CHART_WIDTH = 760;
const CHART_HEIGHT = 340;
const CHART_SHARE_AXIS = '0%';
const CHART_SHARE_LABEL = 'Ашгийн хувь (ашиг / борлуулалт)';
const CHART_TOTAL_COLUMN = 15; // O, "Нийт"
const CHART_TOTAL_LETTER = 'O';
const CHART_TOTAL_ROWS = [6, 7, 8, 9, 10, 11];
const CHART_SHARE_ROW = 12;
const CHART_LABEL_COLUMN = 2; // B

// One block of columns on the helper sheet per chart.
//   series: one column per series, each read from one row of Sheet200, month by month
const CHART_SPECS = [
  {
    title: 'Сарын борлуулалт ба ашиг',
    type: 'COLUMN',
    // Each shop has its own colour; its profit is the lighter shade of it.
    series: [
      { name: 'My home борлуулалт', sourceRow: 6, color: '#4285f4' },
      { name: 'My home ашиг', sourceRow: 7, color: '#a4c2f4' },
      { name: 'Соронзон борлуулалт', sourceRow: 8, color: '#db4437' },
      { name: 'Соронзон ашиг', sourceRow: 9, color: '#ea9999' },
    ],
  },
];
// Charts of earlier versions, removed when found. The old monthly-sales chart is
// among them: its axis showed 1.5E+8 instead of tugrik, so it is drawn again.
const CHART_RETIRED_TITLES = [
  'Жилийн эхнээс өнөөг хүртэл: нийт борлуулалт ба ашиг (₮)',
  'Сарын борлуулалт (₮)',
  'Сарын борлуулалт',
  'Сарын борлуулалтын ашиг (₮)',
  'Ашгийн хувь (ашиг / борлуулалт)',
];

// Colours of the table on Sheet200 (columns A..O): one colour per shop.
const CHART_TABLE_FIRST_ROW = 4;
const CHART_TABLE_FIRST_COLUMN = 1;
const CHART_TABLE_COLORS = [
  '#ffffff', // 4  (empty)
  '#d9d9d9', // 5  months
  '#cfe2f3', // 6  My home sales
  '#cfe2f3', // 7  My home profit
  '#f4cccc', // 8  Соронзон sales
  '#f4cccc', // 9  Соронзон profit
  '#d9ead3', // 10 both, sales
  '#d9ead3', // 11 both, profit
  '#fff2cc', // 12 profit share
  '#cfe2f3', // 13 My home seller
  '#f4cccc', // 14 Соронзон seller
];
// Who sold: one line per shop under the table, the seller's name (B1 of the two
// half-month sheets) month by month. Empty while no name is typed there.
const CHART_SELLER_CELL = '$B$1';
const CHART_SELLER_ROWS = [
  { chainKey: 'my home', row: 13, label: 'My home худалдагч' },
  { chainKey: 'соронзон', row: 14, label: 'Соронзон худалдагч' },
];

function installCharts() {
  const lines = runInstallCharts(currentBook());
  Logger.log('%s\n%s', modeLabel(), lines.join('\n'));
}

// ---------- pure helpers (no Spreadsheet calls) ----------

function sourceCell(reference) {
  return `'${CHART_SOURCE_SHEET}'!${reference}`;
}

// The cells of one block as rows of text; a text that starts with "=" is a formula.
// IFERROR keeps a month that still shows an error on Sheet200 from breaking a chart.
function buildBlockCells(spec) {
  const monthly = spec.series.map(
    (series) =>
      `=ARRAYFORMULA(IFERROR(TRANSPOSE(${sourceCell(`C${series.sourceRow}:N${series.sourceRow}`)}),0))`
  );
  return [
    ['Сар', ...spec.series.map((series) => series.name)],
    [`=TRANSPOSE(${sourceCell(CHART_MONTHS_RANGE)})`, ...monthly],
  ];
}

// Where each chart's block sits on the helper sheet: side by side, one empty column
// between them.
function layoutChartBlocks(specs) {
  return specs.reduce(
    (layout, spec) => {
      const cells = buildBlockCells(spec);
      const block = {
        spec,
        cells,
        firstColumn: layout.nextColumn,
        width: cells[0].length,
      };
      return { nextColumn: layout.nextColumn + block.width + 1, blocks: [...layout.blocks, block] };
    },
    { nextColumn: 1, blocks: [] }
  ).blocks;
}

function buildTotalFormula(row) {
  return row === CHART_SHARE_ROW
    ? `=IFERROR(${CHART_TOTAL_LETTER}11/${CHART_TOTAL_LETTER}10,0)`
    : `=SUM(C${row}:N${row})`;
}

// The names typed on the two sheets of a month, each once, joined by a comma.
function buildSellerFormula(chain, monthIndex) {
  const cells = [0, 1]
    .map((half) => `'${chainSheetName(chain, monthIndex * 2 + half)}'!${CHART_SELLER_CELL}`)
    .join(';');
  return `=IFERROR(TEXTJOIN(", ",TRUE,UNIQUE({${cells}})),"")`;
}

function buildSellerFormulas(chainKey) {
  const chain = CHAINS.find((candidate) => candidate.key === chainKey);
  return Array.from({ length: CHART_MONTH_COUNT }, (_, monthIndex) => buildSellerFormula(chain, monthIndex));
}

function buildTableColors() {
  return CHART_TABLE_COLORS.map((color) => Array(CHART_TOTAL_COLUMN).fill(color));
}

function onlyFormulas(rows) {
  return rows.map((row) => row.map((text) => (String(text).startsWith('=') ? text : '')));
}

// ---------- spreadsheet access ----------

function runInstallCharts(spreadsheet) {
  const source = spreadsheet.getSheetByName(CHART_SOURCE_SHEET);
  if (!source) {
    return [`${CHART_SOURCE_SHEET}: олдсонгүй, график нэмээгүй`];
  }
  const blocks = layoutChartBlocks(CHART_SPECS);
  return [
    ...fillYearTotals(source),
    ...CHART_SELLER_ROWS.map((line) => installSellerRow(source, line)),
    colorSummaryTable(source),
    removeRetiredCharts(source),
    prepareChartData(spreadsheet, blocks),
    ...blocks.map((block, index) => ensureChart(spreadsheet, source, block, index)),
  ].filter((line) => line !== null);
}

// The "Нийт" column and the label of the profit-share row, only where still empty.
function fillYearTotals(source) {
  const rows = [...CHART_TOTAL_ROWS, CHART_SHARE_ROW];
  const empty = rows.filter((row) => {
    const cell = source.getRange(row, CHART_TOTAL_COLUMN);
    return cell.getFormula() === '' && String(cell.getValue()) === '';
  });
  const label = source.getRange(CHART_SHARE_ROW, CHART_LABEL_COLUMN);
  const needsLabel = String(label.getValue()).trim() === '';
  if (!DRY_RUN) {
    empty.forEach((row) => {
      source
        .getRange(row, CHART_TOTAL_COLUMN)
        .setFormula(buildTotalFormula(row))
        .setNumberFormat(row === CHART_SHARE_ROW ? CHART_SHARE_AXIS : MONEY_FORMAT);
    });
    if (needsLabel) {
      label.setValue(CHART_SHARE_LABEL);
      source.getRange(CHART_SHARE_ROW, 3, 1, CHART_MONTH_COUNT).setNumberFormat(CHART_SHARE_AXIS);
    }
  }
  return [
    empty.length > 0
      ? `${CHART_SOURCE_SHEET}: "Нийт" баганын ${empty.length} нүдийг ${DRY_RUN ? 'бөглөнө' : 'бөглөлөө'}`
      : null,
    needsLabel ? `${CHART_SOURCE_SHEET}: ашгийн хувийн мөрд нэр ${DRY_RUN ? 'өгнө' : 'өглөө'}` : null,
  ];
}

function installSellerRow(source, line) {
  const label = source.getRange(line.row, CHART_LABEL_COLUMN);
  const labelNow = String(label.getValue()).trim();
  if (labelNow !== '' && labelNow !== line.label) {
    return `${CHART_SOURCE_SHEET}: ${line.row}-р мөр завгүй ("${labelNow}"), "${line.label}" мөрийг нэмээгүй`;
  }
  const wanted = buildSellerFormulas(line.chainKey);
  const range = source.getRange(line.row, 3, 1, CHART_MONTH_COUNT);
  if (labelNow === line.label && JSON.stringify(range.getFormulas()[0]) === JSON.stringify(wanted)) {
    return null;
  }
  if (!DRY_RUN) {
    label.setValue(line.label);
    range.setFormulas([wanted]);
  }
  return `${CHART_SOURCE_SHEET}: "${line.label}" мөрийг ${DRY_RUN ? 'нэмнэ' : 'нэмлээ'}`;
}

function colorSummaryTable(source) {
  const wanted = buildTableColors();
  const range = source.getRange(
    CHART_TABLE_FIRST_ROW,
    CHART_TABLE_FIRST_COLUMN,
    wanted.length,
    CHART_TOTAL_COLUMN
  );
  if (JSON.stringify(range.getBackgrounds()) === JSON.stringify(wanted)) {
    return null;
  }
  if (!DRY_RUN) {
    range.setBackgrounds(wanted);
  }
  return `${CHART_SOURCE_SHEET}: хүснэгтийг дэлгүүрээр нь ${DRY_RUN ? 'өнгөөр ялгана' : 'өнгөөр ялгалаа'}`;
}

function removeRetiredCharts(source) {
  const retired = source
    .getCharts()
    .filter((chart) => CHART_RETIRED_TITLES.includes(chart.getOptions().get('title')));
  if (retired.length === 0) {
    return null;
  }
  if (!DRY_RUN) {
    retired.forEach((chart) => source.removeChart(chart));
  }
  return `${CHART_SOURCE_SHEET}: хуучин ${retired.length} графикийг ${DRY_RUN ? 'авна' : 'авлаа'}`;
}

function prepareChartData(spreadsheet, blocks) {
  const existing = spreadsheet.getSheetByName(CHART_DATA_SHEET);
  const isReady =
    existing !== null &&
    blocks.every((block) => {
      const range = existing.getRange(1, block.firstColumn, block.cells.length, block.width);
      return JSON.stringify(range.getFormulas()) === JSON.stringify(onlyFormulas(block.cells));
    });
  if (isReady) {
    return null;
  }
  if (!DRY_RUN) {
    const data = existing || spreadsheet.insertSheet(CHART_DATA_SHEET);
    data.clear();
    blocks.forEach((block) => {
      // setValues stores a text that starts with "=" as a formula.
      data.getRange(1, block.firstColumn, block.cells.length, block.width).setValues(block.cells);
      // A chart takes the number format of its axis from the cells it reads.
      data
        .getRange(2, block.firstColumn + 1, CHART_MONTH_COUNT, block.width - 1)
        .setNumberFormat(MONEY_FORMAT);
    });
    data.hideSheet();
  }
  return `${CHART_DATA_SHEET}: графикийн өгөгдлийг ${DRY_RUN ? 'бэлтгэнэ' : 'бэлтгэлээ'}`;
}

function hasChart(sheet, title) {
  return sheet.getCharts().some((chart) => chart.getOptions().get('title') === title);
}

function ensureChart(spreadsheet, source, block, index) {
  const title = block.spec.title;
  if (hasChart(source, title)) {
    return null;
  }
  if (!DRY_RUN) {
    const data = spreadsheet.getSheetByName(CHART_DATA_SHEET);
    const chart = source
      .newChart()
      .setChartType(Charts.ChartType[block.spec.type])
      .addRange(data.getRange(1, block.firstColumn, CHART_MONTH_COUNT + 1, block.width))
      .setNumHeaders(1)
      .setOption('title', title)
      .setOption('width', CHART_WIDTH)
      .setOption('height', CHART_HEIGHT)
      .setOption('legend', { position: 'bottom' })
      .setOption('vAxis', { format: MONEY_FORMAT })
      .setOption('colors', block.spec.series.map((series) => series.color))
      .setPosition(CHART_FIRST_ANCHOR_ROW + index * CHART_ROWS_PER_CHART, CHART_ANCHOR_COLUMN, 0, 0)
      .build();
    source.insertChart(chart);
  }
  return `${CHART_SOURCE_SHEET}: "${title}" графикийг ${DRY_RUN ? 'нэмнэ' : 'нэмлээ'}`;
}
