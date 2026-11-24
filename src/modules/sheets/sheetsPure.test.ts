import {test} from 'vitest';
import assert from 'node:assert/strict';
import {
  colIndexToLetter,
  evaluateWorkbook,
  formatCellValue,
  letterToColIndex,
  shiftFormulaReferences,
} from './formulaEngine';
import {
  computeAutoFitColWidth,
  computeStickyOffsets,
  computeVisibleRowTopOffsets,
  createBlankWorkbook,
  createEmptySheet,
  evaluateConditionalFormatting,
  findSheetMatches,
  sheetToCsvString,
  sortSheetByColumn,
  validateCellInput,
  workbookToPdfBytes,
  workbookToXlsxBytes,
  workbookToXlsxBytesAsync,
  xlsxBytesToWorkbook,
  xlsxBytesToWorkbookAsync,
} from './sheetsPure';
import {cellKey} from './sheetsTypes';

test('column coordinate conversion and formula reference shifting work accurately', () => {
  assert.equal(colIndexToLetter(0), 'A');
  assert.equal(colIndexToLetter(25), 'Z');
  assert.equal(colIndexToLetter(26), 'AA');
  assert.equal(letterToColIndex('AA'), 26);

  assert.equal(shiftFormulaReferences('=B2-C2', 1, 0), '=B3-C3');
  assert.equal(shiftFormulaReferences('=$B$1+B2', 2, 1), '=$B$1+C4');
  assert.equal(shiftFormulaReferences('=SUM(A1:A5)&" A1"', 1, 0), '=SUM(A2:A6)&" A1"');
});

test('formula engine evaluates arithmetic, aggregate, lookup, logic, financial, and cross-sheet formulas', () => {
  const wb = createBlankWorkbook('Budget');
  const s1 = wb.sheets[0];
  s1.cells[cellKey(0, 0)] = {raw: 'Item'};
  s1.cells[cellKey(0, 1)] = {raw: 'Planned'};
  s1.cells[cellKey(0, 2)] = {raw: 'Actual'};
  s1.cells[cellKey(1, 0)] = {raw: 'Operations'};
  s1.cells[cellKey(1, 1)] = {raw: '4200'};
  s1.cells[cellKey(1, 2)] = {raw: '3980'};
  s1.cells[cellKey(2, 0)] = {raw: 'Hardware'};
  s1.cells[cellKey(2, 1)] = {raw: '1800'};
  s1.cells[cellKey(2, 2)] = {raw: '1745'};
  s1.cells[cellKey(3, 1)] = {raw: '=SUM(B2:B3)'};
  s1.cells[cellKey(3, 2)] = {raw: '=AVERAGE(C2:C3)'};
  s1.cells[cellKey(4, 0)] = {raw: '=IF(B2>C2,"Under","Over")'};
  s1.cells[cellKey(4, 1)] = {raw: '=COUNTIF(B2:B3,">2000")'};
  s1.cells[cellKey(4, 2)] = {raw: '=VLOOKUP("Hardware",A2:C3,3,FALSE)'};
  s1.cells[cellKey(5, 0)] = {raw: '=XLOOKUP("Operations",A2:A3,B2:B3)'};
  s1.cells[cellKey(5, 1)] = {raw: '=ROUND(PMT(0.05/12,12,1200),2)'};
  s1.cells[cellKey(5, 2)] = {raw: '=IFERROR(10/0,"Safe")'};
  s1.cells[cellKey(6, 0)] = {raw: '=10/0'};
  // Circular reference
  s1.cells[cellKey(6, 1)] = {raw: '=C7+1'};
  s1.cells[cellKey(6, 2)] = {raw: '=B7+1'};

  const s2 = createEmptySheet('sheet_2', 'Summary');
  s2.cells[cellKey(0, 0)] = {raw: '=Budget!B4*2'};
  wb.sheets.push(s2);

  const evals = evaluateWorkbook(wb);
  const e1 = evals[s1.id];
  const e2 = evals[s2.id];

  assert.equal(e1[cellKey(3, 1)].value, 6000);
  assert.equal(e1[cellKey(3, 2)].value, 2862.5);
  assert.equal(e1[cellKey(4, 0)].value, 'Under');
  assert.equal(e1[cellKey(4, 1)].value, 1);
  assert.equal(e1[cellKey(4, 2)].value, 1745);
  assert.equal(e1[cellKey(5, 0)].value, 4200);
  assert.equal(e1[cellKey(5, 1)].value, -102.73);
  assert.equal(e1[cellKey(5, 2)].value, 'Safe');
  assert.equal(e1[cellKey(6, 0)].error, '#DIV/0!');
  assert.equal(e1[cellKey(6, 1)].error, '#CYCLE!');
  assert.equal(e2[cellKey(0, 0)].value, 12000);
});

test('number formatting, conditional formatting, sorting, and data validation work cleanly', () => {
  assert.equal(formatCellValue(4200, 'currency', 2), '$4,200.00');
  assert.equal(formatCellValue(0.245, 'percent', 1), '24.5%');
  assert.equal(formatCellValue(12500, 'scientific', 2), '1.25E+4');

  const wb = createBlankWorkbook('Data');
  const s = wb.sheets[0];
  s.cells[cellKey(0, 0)] = {raw: 'Name'};
  s.cells[cellKey(0, 1)] = {raw: 'Score'};
  s.cells[cellKey(1, 0)] = {raw: 'Beta'};
  s.cells[cellKey(1, 1)] = {raw: '40'};
  s.cells[cellKey(2, 0)] = {raw: 'Alpha'};
  s.cells[cellKey(2, 1)] = {raw: '95'};

  const evals = evaluateWorkbook(wb);
  const sorted = sortSheetByColumn(s, evals[s.id], 1, 'desc', true);
  assert.equal(sorted.cells[cellKey(1, 0)].raw, 'Alpha');
  assert.equal(sorted.cells[cellKey(2, 0)].raw, 'Beta');

  const cond = evaluateConditionalFormatting(2, 1, evals[s.id][cellKey(2, 1)], [
    {id: 'r1', range: 'B2:B10', operator: 'gt', value1: '80', bg: '#DCFCE7', color: '#166534'},
  ]);
  assert.equal(cond?.bg, '#DCFCE7');

  assert.equal(
    validateCellInput('150', {id: 'v1', range: 'B2:B10', kind: 'number-range', min: 0, max: 100})
      .valid,
    false,
  );
  assert.equal(
    validateCellInput('Open', {
      id: 'v2',
      range: 'C2:C10',
      kind: 'list',
      listItems: ['Open', 'Closed'],
    }).valid,
    true,
  );
});

test('XLSX binary round-trip, CSV serialization, and PDF export produce valid outputs', async () => {
  const wb = createBlankWorkbook('Q4');
  wb.sheets[0].cells[cellKey(0, 0)] = {raw: 'Metric'};
  wb.sheets[0].cells[cellKey(0, 1)] = {raw: 'Amount'};
  wb.sheets[0].cells[cellKey(1, 0)] = {raw: 'Revenue'};
  wb.sheets[0].cells[cellKey(1, 1)] = {raw: '12500', style: {format: 'currency'}};
  wb.sheets[0].cells[cellKey(2, 0)] = {raw: 'Doubled'};
  wb.sheets[0].cells[cellKey(2, 1)] = {raw: '=B2*2'};
  wb.sheets[0].merges.push({startRow: 4, startCol: 0, endRow: 4, endCol: 1});

  const xlsxBytes = workbookToXlsxBytes(wb);
  // ZIP magic bytes PK\x03\x04
  assert.equal(xlsxBytes[0], 0x50);
  assert.equal(xlsxBytes[1], 0x4b);

  const restored = xlsxBytesToWorkbook(xlsxBytes, '.xlsx');
  assert.equal(restored.sheets[0].name, 'Q4');
  assert.equal(restored.sheets[0].cells[cellKey(2, 1)].raw, '=B2*2');
  assert.equal(restored.sheets[0].merges.length, 1);

  const evals = evaluateWorkbook(restored);
  assert.equal(evals[restored.sheets[0].id][cellKey(2, 1)].value, 25000);

  const csv = sheetToCsvString(restored.sheets[0], evals[restored.sheets[0].id]);
  assert.match(csv, /Revenue/);
  assert.match(csv, /25000/);

  const pdfBytes = await workbookToPdfBytes(restored, 'Q4 Report');
  const header = new TextDecoder().decode(pdfBytes.slice(0, 5));
  assert.equal(header, '%PDF-');
});

test('async XLSX round-trip preserves OOXML freeze panes and column widths across multiple sheets', async () => {
  const wb = createBlankWorkbook('Ledger');
  const s1 = wb.sheets[0];
  s1.cells[cellKey(0, 0)] = {raw: 'Account'};
  s1.cells[cellKey(0, 1)] = {raw: 'Q1'};
  s1.cells[cellKey(1, 0)] = {raw: 'Cloud Hosting'};
  s1.cells[cellKey(1, 1)] = {raw: '4800'};
  s1.freezeRows = 1;
  s1.freezeCols = 1;
  s1.colWidths = {0: 190, 1: 120};

  const s2 = createEmptySheet('sheet_2', 'Summary');
  s2.cells[cellKey(0, 0)] = {raw: 'Quarter'};
  s2.cells[cellKey(1, 0)] = {raw: 'Total'};
  s2.cells[cellKey(2, 0)] = {raw: '=Ledger!B2'};
  s2.freezeRows = 2;
  s2.freezeCols = 0;

  const s3 = createEmptySheet('sheet_3', 'Notes');
  s3.cells[cellKey(0, 0)] = {raw: 'Unfrozen sheet'};
  s3.freezeRows = 0;
  s3.freezeCols = 0;

  wb.sheets.push(s2, s3);

  const bytes = await workbookToXlsxBytesAsync(wb, 'xlsx');
  const restored = await xlsxBytesToWorkbookAsync(bytes, '.xlsx');

  assert.equal(restored.sheets.length, 3);
  assert.equal(restored.sheets[0].freezeRows, 1);
  assert.equal(restored.sheets[0].freezeCols, 1);
  assert.equal(restored.sheets[1].freezeRows, 2);
  assert.equal(restored.sheets[1].freezeCols, 0);
  assert.equal(restored.sheets[2].freezeRows, 0);
  assert.equal(restored.sheets[2].freezeCols, 0);
  assert.ok(
    Math.abs((restored.sheets[0].colWidths[0] ?? 0) - 190) <= 12,
    `expected col 0 width near 190, got ${restored.sheets[0].colWidths[0]}`,
  );
});

test('findSheetMatches matches raw and evaluated values, honors case sensitivity and visibleRows', () => {
  const wb = createBlankWorkbook('Sales');
  const s = wb.sheets[0];
  s.cells[cellKey(0, 0)] = {raw: 'Region'};
  s.cells[cellKey(0, 1)] = {raw: 'Total'};
  s.cells[cellKey(1, 0)] = {raw: 'North America'};
  s.cells[cellKey(1, 1)] = {raw: '=1000+250'};
  s.cells[cellKey(2, 0)] = {raw: 'north europe'};
  s.cells[cellKey(2, 1)] = {raw: '500'};

  const evals = evaluateWorkbook(wb)[s.id];

  // Case-insensitive matches both North rows
  const ci = findSheetMatches(s, evals, 'north', false);
  assert.deepEqual(ci, [
    {row: 1, col: 0},
    {row: 2, col: 0},
  ]);

  // Case-sensitive matches only exact case
  const cs = findSheetMatches(s, evals, 'North', true);
  assert.deepEqual(cs, [{row: 1, col: 0}]);

  // Matches evaluated formula result (1250) as well as raw formula text
  const byDisplay = findSheetMatches(s, evals, '1250', false);
  assert.deepEqual(byDisplay, [{row: 1, col: 1}]);

  // Respects visibleRows when Auto-Filter hides row 2
  const filtered = findSheetMatches(s, evals, 'north', false, [0, 1]);
  assert.deepEqual(filtered, [{row: 1, col: 0}]);
});

test('computeAutoFitColWidth and sticky offset helpers calculate zoom-scaled and filtered positions without mutating sheet', () => {
  const wb = createBlankWorkbook('Metrics');
  const s = wb.sheets[0];
  s.cells[cellKey(0, 0)] = {raw: 'Short'};
  s.cells[cellKey(1, 0)] = {raw: 'Very Long Descriptive Category Name'};
  const beforeWidths = {...s.colWidths};

  const evals = evaluateWorkbook(wb)[s.id];
  const autoW = computeAutoFitColWidth(s, evals, 0);
  assert.ok(autoW > 200 && autoW <= 360, `expected wide autofit width, got ${autoW}`);
  assert.deepEqual(s.colWidths, beforeWidths, 'computeAutoFitColWidth must never mutate sheet');

  // Sticky column offsets at 1.25x zoom
  const colOffsets = computeStickyOffsets(3, (c) => Math.round((c === 0 ? 100 : 80) * 1.25), 55);
  assert.deepEqual(colOffsets, [55, 180, 280]);

  // Visible row top offsets skipping hidden row 1
  const rowOffsets = computeVisibleRowTopOffsets([0, 2, 3], () => 30, 35);
  assert.equal(rowOffsets[0], 35);
  assert.equal(rowOffsets[1], undefined);
  assert.equal(rowOffsets[2], 65);
  assert.equal(rowOffsets[3], 95);
});

