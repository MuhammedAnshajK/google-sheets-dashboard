/**
 * Two-way Sales Dashboard for Google Sheets
 *
 * Builds a sales dashboard where the data can be changed in either place:
 *   - Edit the Data sheet, and the Dashboard updates through formulas.
 *   - Type into the white grid on the Dashboard, and the new value is written back
 *     to the Data sheet (by the onEdit function below), so both always agree.
 *
 * How to install: open a blank Google Sheet, go to Extensions > Apps Script,
 * paste this whole file, save, choose "setup" in the function list and click Run.
 * Run setup again at any time to reset everything to the sample data.
 */

const DATA = 'Data';
const DASH = 'Dashboard';
const MONTHS = ['Jan 2025', 'Feb 2025', 'Mar 2025', 'Apr 2025', 'May 2025', 'Jun 2025',
                'Jul 2025', 'Aug 2025', 'Sep 2025', 'Oct 2025', 'Nov 2025', 'Dec 2025'];
const REGIONS = ['North', 'South', 'West'];
const METRICS = ['Revenue', 'Cost', 'Target'];           // editable columns C, D, E on Data
const SAMPLE = {
  North: { Revenue: [388, 435, 429, 427, 477, 445, 445, 481, 489, 450, 406, 431],
           Cost:    [246, 273, 273, 269, 309, 275, 289, 290, 322, 287, 244, 260],
           Target:  [391, 407, 428, 450, 467, 476, 476, 467, 454, 439, 430, 429] },
  South: { Revenue: [275, 278, 313, 333, 345, 350, 379, 364, 325, 313, 331, 335],
           Cost:    [169, 175, 204, 213, 221, 216, 250, 234, 199, 189, 207, 209],
           Target:  [289, 301, 316, 332, 345, 352, 351, 345, 335, 324, 317, 317] },
  West:  { Revenue: [250, 223, 272, 289, 259, 297, 264, 299, 253, 245, 268, 257],
           Cost:    [163, 137, 171, 180, 165, 183, 164, 193, 156, 148, 164, 161],
           Target:  [233, 242, 255, 268, 278, 284, 283, 278, 270, 262, 256, 255] },
};

// Dashboard layout
const REGION_CELL = 'C4';       // region filter
const METRIC_CELL = 'C10';      // which metric the grid shows and edits
const GRID_HEAD = 12;           // header row of the grid (region names)
const GRID_FIRST = 13;          // first month row
const GRID_LAST = GRID_FIRST + MONTHS.length - 1;
const GRID_COL1 = 3;            // column C = first region
const GRID_COLN = GRID_COL1 + REGIONS.length - 1;
const CHART_COL = 14;           // column N: chart data block

const NUM = '#,##0';
const BLUE = '#2a78d6', ORANGE = '#eb6834', NAVY = '#1f4e79', LIGHT = '#dce6f1', GREY = '#f2f2f2';

// ---------------------------------------------------------------- setup

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = getOrCreate_(ss, DATA);
  const dash = getOrCreate_(ss, DASH);
  [data, dash].forEach(function (sh) {
    sh.getCharts().forEach(function (c) { sh.removeChart(c); });
    sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) { p.remove(); });
    sh.clear();
    sh.clearConditionalFormatRules();
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
  });
  buildData_(data);
  buildDashboard_(dash);
  ss.setActiveSheet(dash);
  ss.moveActiveSheet(1);
  const blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0) ss.deleteSheet(blank);
  SpreadsheetApp.flush();
  SpreadsheetApp.getActive().toast('Dashboard is ready. Try typing in the white cells.', 'Setup complete', 6);
}

function getOrCreate_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function buildData_(sh) {
  const n = MONTHS.length * REGIONS.length;
  sh.getRange('A1:G1').setValues([['Month', 'Region', 'Revenue', 'Cost', 'Target', 'Profit', 'Margin %']])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground(NAVY);
  const rows = [];
  REGIONS.forEach(function (region) {
    MONTHS.forEach(function (month, i) {
      rows.push([month, region, SAMPLE[region].Revenue[i], SAMPLE[region].Cost[i], SAMPLE[region].Target[i]]);
    });
  });
  sh.getRange(2, 1, n, 1).setNumberFormat('@');      // keep "Jan 2025" as text, not a date
  sh.getRange(2, 1, n, 5).setValues(rows);
  sh.getRange(2, 3, n, 3).setNumberFormat(NUM).setFontColor('#0000ff');   // inputs in blue
  sh.getRange(2, 6, n, 1).setFormulaR1C1('=RC[-3]-RC[-2]').setNumberFormat(NUM);
  sh.getRange(2, 7, n, 1).setFormulaR1C1('=IF(RC[-4]=0,"",RC[-1]/RC[-4])').setNumberFormat('0.0%');
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 100);
  sh.setColumnWidths(2, 6, 90);
  sh.getRange('I1').setValue('Blue numbers can be edited here or on the Dashboard grid. Profit and margin are formulas. Amounts in $ thousands.')
    .setFontStyle('italic').setFontColor('#595959');
}

function buildDashboard_(sh) {
  const dataRows = MONTHS.length * REGIONS.length + 1;     // last data row
  const D = function (col) { return 'Data!$' + col + '$2:$' + col + '$' + dataRows; };
  const regionCrit = 'IF($C$4="All regions","*",$C$4)';

  sh.setHiddenGridlines(true);
  sh.getRange('B1').setValue('Sales Dashboard 2025').setFontSize(18).setFontWeight('bold').setFontColor(NAVY);
  sh.getRange('B2').setValue('Two-way: change numbers here or on the Data sheet, and both stay in sync. Amounts in $ thousands.')
    .setFontStyle('italic').setFontColor('#595959');

  // filter
  sh.getRange('B4').setValue('Region').setFontWeight('bold');
  sh.getRange(REGION_CELL).setValue('All regions').setBackground('#fff2cc')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['All regions'].concat(REGIONS), true).build());

  // KPI tiles
  sh.getRange('B6:F6').setValues([['Revenue', 'Target', 'Vs target', 'Profit', 'Margin']])
    .setFontColor('#595959').setBackground(LIGHT);
  sh.getRange('B7:F7').setFormulas([[
    '=SUMIFS(' + D('C') + ',' + D('B') + ',' + regionCrit + ')',
    '=SUMIFS(' + D('E') + ',' + D('B') + ',' + regionCrit + ')',
    '=B7-C7',
    '=SUMIFS(' + D('F') + ',' + D('B') + ',' + regionCrit + ')',
    '=IF(B7=0,"",E7/B7)',
  ]]).setFontSize(16).setFontWeight('bold').setBackground(LIGHT);
  sh.getRange('B7:E7').setNumberFormat(NUM);
  sh.getRange('D7').setNumberFormat('+#,##0;-#,##0');
  sh.getRange('F7').setNumberFormat('0.0%');
  sh.getRange('D8').setFormula('=IF(C7=0,"",D7/C7)').setNumberFormat('+0.0%" vs target";-0.0%" vs target"')
    .setFontColor('#595959');
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor('#c00000').setRanges([sh.getRange('D7:D8')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0).setFontColor('#006300').setRanges([sh.getRange('D7:D8')]).build(),
  ]);

  // editable grid
  sh.getRange('B10').setValue('Show and edit').setFontWeight('bold');
  sh.getRange(METRIC_CELL).setValue('Revenue').setBackground('#fff2cc')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(METRICS, true).build());
  sh.getRange('D10').setValue('← pick a metric, then type in the white cells below. Changes are saved to the Data sheet.')
    .setFontStyle('italic').setFontColor('#595959');
  sh.getRange(GRID_HEAD, 2, 1, REGIONS.length + 2).setValues([['Month'].concat(REGIONS, ['Total'])])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground(NAVY);
  sh.getRange(GRID_FIRST, 2, MONTHS.length, 1).setNumberFormat('@')
    .setValues(MONTHS.map(function (m) { return [m]; }));
  for (let r = GRID_FIRST; r <= GRID_LAST; r++) {
    for (let c = GRID_COL1; c <= GRID_COLN; c++) sh.getRange(r, c).setFormula(gridFormula_(r, c));
    sh.getRange(r, GRID_COLN + 1).setFormula('=SUM(' + a1_(r, GRID_COL1) + ':' + a1_(r, GRID_COLN) + ')');
  }
  const totalRow = GRID_LAST + 1;
  sh.getRange(totalRow, 2).setValue('Total').setFontWeight('bold');
  for (let c = GRID_COL1; c <= GRID_COLN + 1; c++) {
    sh.getRange(totalRow, c).setFormula('=SUM(' + a1_(GRID_FIRST, c) + ':' + a1_(GRID_LAST, c) + ')').setFontWeight('bold');
  }
  const grid = sh.getRange(GRID_FIRST, GRID_COL1, MONTHS.length, REGIONS.length);
  grid.setBackground('#ffffff').setFontColor('#0000ff')
    .setBorder(true, true, true, true, true, true, '#bfbfbf', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(GRID_FIRST, GRID_COLN + 1, MONTHS.length + 1, 1).setBackground(GREY);
  sh.getRange(totalRow, GRID_COL1, 1, REGIONS.length).setBackground(GREY);
  sh.getRange(GRID_FIRST, GRID_COL1, MONTHS.length + 1, REGIONS.length + 1).setNumberFormat(NUM);

  // chart data (follows the region filter)
  const cc = CHART_COL;
  sh.getRange(GRID_HEAD - 1, cc).setValue('Chart data (formulas)').setFontStyle('italic').setFontColor('#595959');
  sh.getRange(GRID_HEAD, cc, 1, 4).setValues([['Month', 'Revenue', 'Target', 'Profit']]).setFontWeight('bold');
  for (let r = GRID_FIRST; r <= GRID_LAST; r++) {
    const m = '$' + colLetter_(cc) + r;
    sh.getRange(r, cc).setFormula('=' + a1_(r, 2));
    sh.getRange(r, cc + 1, 1, 3).setFormulas([[
      '=SUMIFS(' + D('C') + ',' + D('A') + ',' + m + ',' + D('B') + ',' + regionCrit + ')',
      '=SUMIFS(' + D('E') + ',' + D('A') + ',' + m + ',' + D('B') + ',' + regionCrit + ')',
      '=SUMIFS(' + D('F') + ',' + D('A') + ',' + m + ',' + D('B') + ',' + regionCrit + ')',
    ]]);
  }
  const regHead = GRID_LAST + 3;
  sh.getRange(regHead, cc, 1, 3).setValues([['Region', 'Revenue', 'Target']]).setFontWeight('bold');
  REGIONS.forEach(function (reg, i) {
    const r = regHead + 1 + i;
    sh.getRange(r, cc).setValue(reg);
    sh.getRange(r, cc + 1, 1, 2).setFormulas([[
      '=SUMIFS(' + D('C') + ',' + D('B') + ',$' + colLetter_(cc) + r + ')',
      '=SUMIFS(' + D('E') + ',' + D('B') + ',$' + colLetter_(cc) + r + ')',
    ]]);
  });
  sh.getRange(GRID_FIRST, cc + 1, regHead + REGIONS.length - GRID_FIRST + 1, 3).setNumberFormat(NUM);

  // charts
  const c1 = sh.newChart().setChartType(Charts.ChartType.LINE)
    .addRange(sh.getRange(GRID_HEAD, cc, MONTHS.length + 1, 3))
    .setNumHeaders(1)
    .setPosition(3, 8, 0, 0)
    .setOption('title', 'Revenue vs target by month')
    .setOption('colors', [BLUE, ORANGE])
    .setOption('legend', { position: 'bottom' })
    .setOption('width', 560).setOption('height', 260)
    .build();
  sh.insertChart(c1);
  const c2 = sh.newChart().setChartType(Charts.ChartType.BAR)
    .addRange(sh.getRange(regHead, cc, REGIONS.length + 1, 3))
    .setNumHeaders(1)
    .setPosition(17, 8, 0, 0)
    .setOption('title', 'Revenue vs target by region (full year)')
    .setOption('colors', [BLUE, ORANGE])
    .setOption('legend', { position: 'bottom' })
    .setOption('width', 560).setOption('height', 240)
    .build();
  sh.insertChart(c2);

  // warn (not block) if someone types over a formula outside the editable grid
  [sh.getRange('B6:F8'), sh.getRange(GRID_FIRST, GRID_COLN + 1, MONTHS.length + 1, 1),
   sh.getRange(totalRow, GRID_COL1, 1, REGIONS.length), sh.getRange(GRID_HEAD, cc, regHead + REGIONS.length - GRID_HEAD + 1, 4)]
    .forEach(function (rg) { rg.protect().setDescription('Formula: edit the white cells instead').setWarningOnly(true); });

  sh.setColumnWidth(1, 20);
  sh.setColumnWidths(2, 5, 110);
  sh.setColumnWidth(7, 20);
}

// ---------------------------------------------------------------- two-way sync

/**
 * Runs automatically whenever someone edits the spreadsheet.
 * If the edit is in the Dashboard grid, it saves the new value to the matching
 * Data row (month + region + selected metric) and puts the grid formula back,
 * so the Dashboard keeps reading from the Data sheet.
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() !== DASH) return;

  const r0 = e.range.getRow(), c0 = e.range.getColumn();
  const r1 = e.range.getLastRow(), c1 = e.range.getLastColumn();
  if (r1 < GRID_FIRST || r0 > GRID_LAST || c1 < GRID_COL1 || c0 > GRID_COLN) return;

  const metric = sh.getRange(METRIC_CELL).getValue();
  const metricIdx = METRICS.indexOf(metric);
  const data = e.source.getSheetByName(DATA);
  if (metricIdx < 0 || !data) return;
  const dataCol = 3 + metricIdx;
  const keys = data.getRange(2, 1, data.getLastRow() - 1, 2).getDisplayValues();   // [month, region]

  for (let r = Math.max(r0, GRID_FIRST); r <= Math.min(r1, GRID_LAST); r++) {
    for (let c = Math.max(c0, GRID_COL1); c <= Math.min(c1, GRID_COLN); c++) {
      const cell = sh.getRange(r, c);
      const value = cell.getValue();
      const month = sh.getRange(r, 2).getDisplayValue();
      const region = sh.getRange(GRID_HEAD, c).getDisplayValue();
      const idx = keys.findIndex(function (k) { return k[0] === month && k[1] === region; });
      if (idx >= 0 && value !== '' && !isNaN(Number(value))) {
        data.getRange(idx + 2, dataCol).setValue(Number(value));
      }
      cell.setFormula(gridFormula_(r, c));    // always restore the link to the Data sheet
    }
  }
}

// ---------------------------------------------------------------- helpers

function gridFormula_(r, c) {
  const last = MONTHS.length * REGIONS.length + 1;
  return '=SUMIFS(INDEX(Data!$C$2:$E$' + last + ',0,MATCH($C$10,Data!$C$1:$E$1,0)),' +
         'Data!$A$2:$A$' + last + ',$B' + r + ',' +
         'Data!$B$2:$B$' + last + ',' + colLetter_(c) + '$' + GRID_HEAD + ')';
}

function colLetter_(c) {
  let s = '';
  while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
  return s;
}

function a1_(r, c) {
  return colLetter_(c) + r;
}
