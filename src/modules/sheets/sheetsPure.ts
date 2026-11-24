import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import {PDFDocument, StandardFonts, rgb} from 'pdf-lib';
import {
  colIndexToLetter,
  evaluateWorkbook,
  formatCellRef,
  isCoordInRange,
  shiftFormulaReferences,
} from './formulaEngine';
import {
  cellKey,
  DEFAULT_COL_COUNT,
  DEFAULT_COL_WIDTH,
  DEFAULT_ROW_COUNT,
  type CellCoord,
  type CellData,
  type ConditionalRule,
  type DataValidationRule,
  type EvaluatedCell,
  type EvaluatedScalar,
  type MergeRange,
  type PersistedSheetMeta,
  type SheetData,
  type WorkbookData,
} from './sheetsTypes';

export function createEmptySheet(id: string, name: string): SheetData {
  return {
    id,
    name,
    rowCount: DEFAULT_ROW_COUNT,
    colCount: DEFAULT_COL_COUNT,
    cells: {},
    colWidths: {},
    rowHeights: {},
    merges: [],
    freezeRows: 0,
    freezeCols: 0,
    autoFilterEnabled: false,
    columnFilters: {},
    conditionalRules: [],
    validations: [],
    charts: [],
  };
}

export function createBlankWorkbook(sheetName = 'Sheet1'): WorkbookData {
  const sheet = createEmptySheet('sheet_1', sheetName);
  return {
    activeSheetId: sheet.id,
    sheets: [sheet],
  };
}

export function extractPersistedMeta(wb: WorkbookData): PersistedSheetMeta {
  const sheetsMeta: PersistedSheetMeta['sheetsMeta'] = {};
  for (const s of wb.sheets) {
    const styles: Record<string, NonNullable<CellData['style']>> = {};
    for (const [k, c] of Object.entries(s.cells)) {
      if (c.style && Object.keys(c.style).length > 0) {
        styles[k] = c.style;
      }
    }
    sheetsMeta[s.name] = {
      styles,
      colWidths: s.colWidths,
      rowHeights: s.rowHeights,
      merges: s.merges,
      freezeRows: s.freezeRows,
      freezeCols: s.freezeCols,
      autoFilterEnabled: s.autoFilterEnabled,
      conditionalRules: s.conditionalRules,
      validations: s.validations,
      charts: s.charts,
    };
  }
  return {
    version: 1,
    activeSheetId: wb.activeSheetId,
    sheetsMeta,
  };
}

export function applyPersistedMeta(wb: WorkbookData, meta: PersistedSheetMeta | null): WorkbookData {
  if (!meta || meta.version !== 1) return wb;
  const sheets = wb.sheets.map((s) => {
    const sm = meta.sheetsMeta[s.name];
    if (!sm) return s;
    const nextCells: Record<string, CellData> = {...s.cells};
    for (const [k, st] of Object.entries(sm.styles ?? {})) {
      if (nextCells[k]) {
        nextCells[k] = {...nextCells[k], style: {...nextCells[k].style, ...st}};
      } else {
        nextCells[k] = {raw: '', style: st};
      }
    }
    return {
      ...s,
      cells: nextCells,
      colWidths: {...s.colWidths, ...(sm.colWidths ?? {})},
      rowHeights: {...s.rowHeights, ...(sm.rowHeights ?? {})},
      merges: sm.merges?.length ? sm.merges : s.merges,
      freezeRows: sm.freezeRows ?? s.freezeRows,
      freezeCols: sm.freezeCols ?? s.freezeCols,
      autoFilterEnabled: Boolean(sm.autoFilterEnabled),
      conditionalRules: sm.conditionalRules ?? [],
      validations: sm.validations ?? [],
      charts: sm.charts ?? [],
    };
  });
  const hasActive = sheets.some((s) => s.id === meta.activeSheetId);
  return {
    activeSheetId: hasActive ? meta.activeSheetId : sheets[0]?.id ?? 'sheet_1',
    sheets,
  };
}

function isZipHeader(bytes: Uint8Array): boolean {
  return bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 0x03 && bytes[3] === 0x04;
}

function isOleXlsHeader(bytes: Uint8Array): boolean {
  return bytes.length >= 4 && bytes[0] === 0xd0 && bytes[1] === 0xcf && bytes[2] === 0x11 && bytes[3] === 0xe0;
}

/**
 * Parses .xlsx, .xls, or .csv bytes into our multi-sheet WorkbookData model,
 * preserving formulas (=...), values, merges, and column widths.
 */
export function xlsxBytesToWorkbook(bytes: Uint8Array, extension = '.xlsx'): WorkbookData {
  if (bytes.length === 0) return createBlankWorkbook();

  const isBinary = isZipHeader(bytes) || isOleXlsHeader(bytes);
  let xwb: XLSX.WorkBook;

  if (isBinary) {
    xwb = XLSX.read(bytes, {type: 'array', cellFormula: true, cellNF: true, cellStyles: true});
  } else {
    const text = new TextDecoder('utf-8', {fatal: false}).decode(bytes);
    if (!text.trim()) return createBlankWorkbook();
    // Starter files or CSV text
    xwb = XLSX.read(text, {type: 'string', cellFormula: true});
  }

  const sheetNames = xwb.SheetNames.length > 0 ? xwb.SheetNames : ['Sheet1'];
  const sheets: SheetData[] = sheetNames.map((name, idx) => {
    const ws = xwb.Sheets[name];
    const sheet = createEmptySheet(`sheet_${idx + 1}`, name);
    if (!ws) return sheet;

    const ref = ws['!ref'];
    let maxR = DEFAULT_ROW_COUNT - 1;
    let maxC = DEFAULT_COL_COUNT - 1;

    if (ref) {
      const range = XLSX.utils.decode_range(ref);
      maxR = Math.max(maxR, range.e.r + 10);
      maxC = Math.max(maxC, range.e.c + 4);
      for (let r = range.s.r; r <= range.e.r; r++) {
        for (let c = range.s.c; c <= range.e.c; c++) {
          const addr = XLSX.utils.encode_cell({r, c});
          const xcell = ws[addr] as XLSX.CellObject | undefined;
          if (!xcell) continue;
          let raw = '';
          if (xcell.f) {
            raw = xcell.f.startsWith('=') ? xcell.f : `=${xcell.f}`;
          } else if (xcell.v !== undefined && xcell.v !== null) {
            raw = String(xcell.v);
          }
          if (raw !== '') {
            const style: CellData['style'] = {};
            if (xcell.z && typeof xcell.z === 'string') {
              if (xcell.z.includes('$')) style.format = 'currency';
              else if (xcell.z.includes('%')) style.format = 'percent';
              else if (xcell.z.includes('yy') || xcell.z.includes('mm')) style.format = 'date';
            }
            sheet.cells[cellKey(r, c)] = {
              raw,
              style: Object.keys(style).length > 0 ? style : undefined,
            };
          }
        }
      }
    }

    sheet.rowCount = Math.min(200, maxR + 1);
    sheet.colCount = Math.min(52, maxC + 1);

    if (Array.isArray(ws['!merges'])) {
      sheet.merges = ws['!merges'].map((m) => ({
        startRow: m.s.r,
        startCol: m.s.c,
        endRow: m.e.r,
        endCol: m.e.c,
      }));
    }

    if (Array.isArray(ws['!cols'])) {
      ws['!cols'].forEach((colInfo, cIdx) => {
        if (colInfo?.wpx) {
          sheet.colWidths[cIdx] = Math.max(56, Math.min(360, Math.round(colInfo.wpx)));
        } else if (colInfo?.wch) {
          sheet.colWidths[cIdx] = Math.max(56, Math.min(360, Math.round(colInfo.wch * 8.5)));
        }
      });
    }

    // Upgrade header styling cleanly when loading starter CSV/XLSX text
    if (!isBinary && extension === '.xlsx' && idx === 0 && Object.keys(sheet.cells).length > 0) {
      for (let c = 0; c < sheet.colCount; c++) {
        const headerCell = sheet.cells[cellKey(0, c)];
        if (headerCell && headerCell.raw) {
          headerCell.style = {...headerCell.style, bold: true, bg: '#EFF6FF'};
        }
      }
    }

    return sheet;
  });

  return {
    activeSheetId: sheets[0].id,
    sheets,
  };
}

/**
 * Serializes WorkbookData into real .xlsx (or .xls) bytes via SheetJS CE,
 * writing formulas, cached values, merges, and column widths.
 */
export function workbookToXlsxBytes(
  workbook: WorkbookData,
  bookType: 'xlsx' | 'xls' = 'xlsx',
): Uint8Array {
  const evalMaps = evaluateWorkbook(workbook);
  const xwb = XLSX.utils.book_new();

  for (const sheet of workbook.sheets) {
    const ws: XLSX.WorkSheet = {};
    const evalMap = evalMaps[sheet.id] ?? {};
    let maxR = 0;
    let maxC = 0;
    let hasAny = false;

    for (const [k, cell] of Object.entries(sheet.cells)) {
      if (!cell.raw) continue;
      const [r, c] = k.split(',').map(Number);
      hasAny = true;
      if (r > maxR) maxR = r;
      if (c > maxC) maxC = c;
      const addr = XLSX.utils.encode_cell({r, c});
      const evaluated = evalMap[k];

      const xcell: XLSX.CellObject = {t: 's', v: ''};
      if (cell.raw.startsWith('=')) {
        xcell.f = cell.raw.slice(1);
        if (evaluated && typeof evaluated.value === 'number') {
          xcell.t = 'n';
          xcell.v = evaluated.value;
        } else if (evaluated && typeof evaluated.value === 'boolean') {
          xcell.t = 'b';
          xcell.v = evaluated.value;
        } else {
          xcell.t = 's';
          xcell.v = evaluated?.display ?? '';
        }
      } else {
        const val = evaluated?.value;
        if (typeof val === 'number') {
          xcell.t = 'n';
          xcell.v = val;
        } else if (typeof val === 'boolean') {
          xcell.t = 'b';
          xcell.v = val;
        } else {
          xcell.t = 's';
          xcell.v = cell.raw;
        }
      }

      if (cell.style?.format === 'currency') xcell.z = '$#,##0.00';
      else if (cell.style?.format === 'percent') xcell.z = '0.00%';
      else if (cell.style?.format === 'date') xcell.z = 'yyyy-mm-dd';

      ws[addr] = xcell;
    }

    ws['!ref'] = hasAny
      ? XLSX.utils.encode_range({s: {r: 0, c: 0}, e: {r: maxR, c: maxC}})
      : 'A1:A1';

    if (sheet.merges.length > 0) {
      ws['!merges'] = sheet.merges.map((m) => ({
        s: {r: m.startRow, c: m.startCol},
        e: {r: m.endRow, c: m.endCol},
      }));
    }

    const cols: XLSX.ColInfo[] = [];
    for (let c = 0; c <= maxC; c++) {
      cols.push({wpx: sheet.colWidths[c] ?? DEFAULT_COL_WIDTH});
    }
    ws['!cols'] = cols;

    XLSX.utils.book_append_sheet(xwb, ws, sheet.name.slice(0, 31) || 'Sheet1');
  }

  const out = XLSX.write(xwb, {type: 'array', bookType}) as ArrayBuffer | Uint8Array;
  return out instanceof Uint8Array ? out : new Uint8Array(out);
}

/**
 * Resolves ordered worksheet XML paths inside an OOXML .xlsx archive
 * by reading xl/workbook.xml and xl/_rels/workbook.xml.rels.
 */
async function resolveXlsxSheetXmlPaths(zip: JSZip, fallbackCount: number): Promise<string[]> {
  const wbXml = await zip.file('xl/workbook.xml')?.async('string');
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('string');
  if (wbXml && relsXml) {
    const relMap = new Map<string, string>();
    const relRegex = /<Relationship\b[^>]*>/g;
    let rm: RegExpExecArray | null;
    while ((rm = relRegex.exec(relsXml)) !== null) {
      const tag = rm[0];
      const idMatch = /\bId="([^"]+)"/.exec(tag);
      const targetMatch = /\bTarget="([^"]+)"/.exec(tag);
      if (idMatch && targetMatch) {
        let target = targetMatch[1].replace(/^\/+/, '');
        if (!target.startsWith('xl/')) target = `xl/${target}`;
        relMap.set(idMatch[1], target);
      }
    }

    const paths: string[] = [];
    const sheetRegex = /<sheet\b[^>]*>/g;
    let sm: RegExpExecArray | null;
    while ((sm = sheetRegex.exec(wbXml)) !== null) {
      const tag = sm[0];
      const ridMatch = /\br:id="([^"]+)"/.exec(tag);
      if (ridMatch) {
        const resolved = relMap.get(ridMatch[1]);
        if (resolved) paths.push(resolved);
      }
    }
    if (paths.length > 0) return paths;
  }

  return Array.from({length: fallbackCount}, (_, i) => `xl/worksheets/sheet${i + 1}.xml`);
}

function parseFrozenPaneFromSheetXml(xml: string): {freezeRows: number; freezeCols: number} {
  const paneMatch = /<pane\b[^>]*>/i.exec(xml);
  if (!paneMatch) return {freezeRows: 0, freezeCols: 0};
  const tag = paneMatch[0];
  const stateMatch = /\bstate="([^"]+)"/i.exec(tag);
  const state = stateMatch?.[1]?.toLowerCase() ?? '';
  if (state !== 'frozen' && state !== 'frozensplit') {
    return {freezeRows: 0, freezeCols: 0};
  }
  const xMatch = /\bxSplit="(\d+(?:\.\d+)?)"/i.exec(tag);
  const yMatch = /\bySplit="(\d+(?:\.\d+)?)"/i.exec(tag);
  return {
    freezeCols: xMatch ? Math.max(0, Math.floor(Number(xMatch[1]))) : 0,
    freezeRows: yMatch ? Math.max(0, Math.floor(Number(yMatch[1]))) : 0,
  };
}

function injectFrozenPaneIntoSheetXml(
  xml: string,
  freezeRows: number,
  freezeCols: number,
): string {
  const fRows = Math.max(0, Math.floor(freezeRows));
  const fCols = Math.max(0, Math.floor(freezeCols));
  if (fRows === 0 && fCols === 0) return xml;

  const attrs: string[] = [];
  if (fCols > 0) attrs.push(`xSplit="${fCols}"`);
  if (fRows > 0) attrs.push(`ySplit="${fRows}"`);
  attrs.push(`topLeftCell="${formatCellRef(fRows, fCols)}"`);
  const activePane =
    fRows > 0 && fCols > 0 ? 'bottomRight' : fRows > 0 ? 'bottomLeft' : 'topRight';
  attrs.push(`activePane="${activePane}"`);
  attrs.push('state="frozen"');
  const paneXml = `<pane ${attrs.join(' ')}/>`;

  // Strip any existing <pane .../> inside sheetViews first
  let cleaned = xml.replace(/<pane\b[^>]*\/?>(?:<\/pane>)?/gi, '');

  if (/<sheetView\b[^>]*\/>/i.test(cleaned)) {
    return cleaned.replace(
      /<sheetView\b([^>]*)\/>/i,
      `<sheetView$1>${paneXml}</sheetView>`,
    );
  }
  if (/<sheetView\b[^>]*>/i.test(cleaned)) {
    return cleaned.replace(
      /<sheetView\b([^>]*)>/i,
      `<sheetView$1>${paneXml}`,
    );
  }
  const sheetViewsBlock = `<sheetViews><sheetView workbookViewId="0">${paneXml}</sheetView></sheetViews>`;
  if (/<sheetFormatPr\b/i.test(cleaned)) {
    return cleaned.replace(/<sheetFormatPr\b/i, `${sheetViewsBlock}<sheetFormatPr`);
  }
  if (/<cols\b/i.test(cleaned)) {
    return cleaned.replace(/<cols\b/i, `${sheetViewsBlock}<cols`);
  }
  if (/<sheetData\b/i.test(cleaned)) {
    return cleaned.replace(/<sheetData\b/i, `${sheetViewsBlock}<sheetData`);
  }
  return cleaned;
}

/**
 * Async OOXML-aware .xlsx/.xls/.csv parser that also extracts frozen panes
 * (<sheetViews><sheetView><pane state="frozen" xSplit ySplit/></sheetView></sheetViews>)
 * which SheetJS CE omits on read.
 */
export async function xlsxBytesToWorkbookAsync(
  bytes: Uint8Array,
  extension = '.xlsx',
): Promise<WorkbookData> {
  const wb = xlsxBytesToWorkbook(bytes, extension);
  if (!isZipHeader(bytes)) return wb;
  try {
    const zip = await JSZip.loadAsync(bytes);
    const paths = await resolveXlsxSheetXmlPaths(zip, wb.sheets.length);
    for (let i = 0; i < wb.sheets.length; i++) {
      const xmlPath = paths[i];
      if (!xmlPath) continue;
      const xml = await zip.file(xmlPath)?.async('string');
      if (!xml) continue;
      const {freezeRows, freezeCols} = parseFrozenPaneFromSheetXml(xml);
      if (freezeRows > 0 || freezeCols > 0) {
        wb.sheets[i] = {
          ...wb.sheets[i],
          freezeRows: Math.min(wb.sheets[i].rowCount - 1, freezeRows),
          freezeCols: Math.min(wb.sheets[i].colCount - 1, freezeCols),
        };
      }
    }
  } catch {
    // Non-fatal: fall back to SheetJS CE result if archive lacks sheetViews
  }
  return wb;
}

/**
 * Async OOXML-aware .xlsx/.xls serializer that injects standard Excel freeze panes
 * (<pane state="frozen"/>) into worksheet XML files so Excel/Calc/Sheets honor them.
 */
export async function workbookToXlsxBytesAsync(
  workbook: WorkbookData,
  bookType: 'xlsx' | 'xls' = 'xlsx',
): Promise<Uint8Array> {
  const baseBytes = workbookToXlsxBytes(workbook, bookType);
  if (bookType !== 'xlsx') return baseBytes;
  const hasAnyFreeze = workbook.sheets.some((s) => s.freezeRows > 0 || s.freezeCols > 0);
  if (!hasAnyFreeze) return baseBytes;

  try {
    const zip = await JSZip.loadAsync(baseBytes);
    const paths = await resolveXlsxSheetXmlPaths(zip, workbook.sheets.length);
    for (let i = 0; i < workbook.sheets.length; i++) {
      const sheet = workbook.sheets[i];
      if (sheet.freezeRows <= 0 && sheet.freezeCols <= 0) continue;
      const xmlPath = paths[i];
      if (!xmlPath) continue;
      const xml = await zip.file(xmlPath)?.async('string');
      if (!xml) continue;
      const patched = injectFrozenPaneIntoSheetXml(xml, sheet.freezeRows, sheet.freezeCols);
      zip.file(xmlPath, patched);
    }
    return await zip.generateAsync({type: 'uint8array', compression: 'DEFLATE'});
  } catch {
    return baseBytes;
  }
}

/**
 * Pure search helper matching both raw formulas and evaluated display values,
 * respecting currently visible rows when Auto-Filter is active.
 */
export function findSheetMatches(
  sheet: SheetData,
  evalMap: Record<string, EvaluatedCell>,
  query: string,
  caseSensitive: boolean,
  visibleRows?: number[],
): CellCoord[] {
  const trimmed = query;
  if (!trimmed) return [];
  const needle = caseSensitive ? trimmed : trimmed.toLowerCase();
  const rowsToScan =
    visibleRows ?? Array.from({length: sheet.rowCount}, (_, r) => r);
  const matches: CellCoord[] = [];
  for (const r of rowsToScan) {
    for (let c = 0; c < sheet.colCount; c++) {
      const k = cellKey(r, c);
      const raw = sheet.cells[k]?.raw ?? '';
      const disp = evalMap[k]?.display ?? '';
      if (!raw && !disp) continue;
      const hay = caseSensitive ? `${raw}\n${disp}` : `${raw}\n${disp}`.toLowerCase();
      if (hay.includes(needle)) {
        matches.push({row: r, col: c});
      }
    }
  }
  return matches;
}

/**
 * Computes an ideal unscaled column width (px) to fit the widest visible cell in colIdx.
 * Pure — never mutates the sheet.
 */
export function computeAutoFitColWidth(
  sheet: SheetData,
  evalMap: Record<string, EvaluatedCell>,
  colIdx: number,
  visibleRows?: number[],
): number {
  const rowsToScan =
    visibleRows ?? Array.from({length: sheet.rowCount}, (_, r) => r);
  let maxChars = 4;
  let maxFontPx = 13;
  for (const r of rowsToScan) {
    const k = cellKey(r, colIdx);
    const cell = sheet.cells[k];
    const display = evalMap[k]?.display ?? cell?.raw ?? '';
    if (!display) continue;
    // Use longest line when cell contains manual line breaks
    const longestLine = display
      .split(/\r?\n/)
      .reduce((acc, line) => Math.max(acc, line.length), 0);
    if (longestLine > maxChars) {
      maxChars = longestLine;
      maxFontPx = cell?.style?.fontSize ?? 13;
    }
  }
  const approxCharWidth = Math.max(6.5, maxFontPx * 0.62);
  const estimated = Math.round(maxChars * approxCharWidth + 24);
  return Math.max(64, Math.min(360, estimated));
}

/**
 * Computes cumulative left sticky offsets for columns (used for Freeze Columns).
 */
export function computeStickyOffsets(
  count: number,
  getSize: (index: number) => number,
  baseOffset: number,
): number[] {
  const offsets: number[] = [];
  let acc = baseOffset;
  for (let i = 0; i < count; i++) {
    offsets.push(acc);
    acc += getSize(i);
  }
  return offsets;
}

/**
 * Computes cumulative top sticky offsets keyed by row index across visible rows,
 * so frozen rows stay pixel-aligned even when Auto-Filter hides rows or zoom changes.
 */
export function computeVisibleRowTopOffsets(
  visibleRows: number[],
  getRowHeight: (rowIdx: number) => number,
  headerHeight: number,
): Record<number, number> {
  const offsets: Record<number, number> = {};
  let acc = headerHeight;
  for (const r of visibleRows) {
    offsets[r] = acc;
    acc += getRowHeight(r);
  }
  return offsets;
}

export function sheetToCsvString(
  sheet: SheetData,
  evalMap: Record<string, EvaluatedCell>,
): string {
  let maxR = -1;
  let maxC = -1;
  for (const [k, cell] of Object.entries(sheet.cells)) {
    if (!cell.raw) continue;
    const [r, c] = k.split(',').map(Number);
    if (r > maxR) maxR = r;
    if (c > maxC) maxC = c;
  }
  if (maxR < 0 || maxC < 0) return '';

  const escapeCsv = (val: string) => {
    if (/[",\n\r]/.test(val)) {
      return `"${val.replace(/"/g, '""')}"`;
    }
    return val;
  };

  const lines: string[] = [];
  for (let r = 0; r <= maxR; r++) {
    const row: string[] = [];
    for (let c = 0; c <= maxC; c++) {
      const k = cellKey(r, c);
      const display = evalMap[k]?.display ?? sheet.cells[k]?.raw ?? '';
      row.push(escapeCsv(display));
    }
    lines.push(row.join(','));
  }
  return lines.join('\n');
}

export function sortSheetByColumn(
  sheet: SheetData,
  evalMap: Record<string, EvaluatedCell>,
  colIndex: number,
  direction: 'asc' | 'desc',
  hasHeaderRow = true,
): SheetData {
  let maxR = -1;
  let maxC = -1;
  for (const [k, cell] of Object.entries(sheet.cells)) {
    if (!cell.raw) continue;
    const [r, c] = k.split(',').map(Number);
    if (r > maxR) maxR = r;
    if (c > maxC) maxC = c;
  }
  const startRow = hasHeaderRow ? 1 : 0;
  if (maxR <= startRow) return sheet;

  const rows: Array<{
    origRow: number;
    sortVal: EvaluatedScalar;
    cells: Record<number, CellData>;
  }> = [];

  for (let r = startRow; r <= maxR; r++) {
    const rowCells: Record<number, CellData> = {};
    let hasContent = false;
    for (let c = 0; c <= maxC; c++) {
      const cell = sheet.cells[cellKey(r, c)];
      if (cell) {
        rowCells[c] = cell;
        if (cell.raw !== '') hasContent = true;
      }
    }
    if (!hasContent) continue;
    rows.push({
      origRow: r,
      sortVal: evalMap[cellKey(r, colIndex)]?.value ?? null,
      cells: rowCells,
    });
  }

  rows.sort((a, b) => {
    const av = a.sortVal;
    const bv = b.sortVal;
    if (av === null && bv === null) return 0;
    if (av === null) return 1;
    if (bv === null) return -1;
    let cmp = 0;
    if (typeof av === 'number' && typeof bv === 'number') {
      cmp = av - bv;
    } else {
      cmp = String(av).localeCompare(String(bv), 'en-US', {numeric: true, sensitivity: 'base'});
    }
    return direction === 'asc' ? cmp : -cmp;
  });

  const nextCells: Record<string, CellData> = {};
  // Keep header rows and untouched rows above startRow or below maxR
  for (const [k, cell] of Object.entries(sheet.cells)) {
    const [r] = k.split(',').map(Number);
    if (r < startRow || r > maxR) {
      nextCells[k] = cell;
    }
  }

  rows.forEach((entry, idx) => {
    const targetRow = startRow + idx;
    const dRow = targetRow - entry.origRow;
    for (const [cStr, cell] of Object.entries(entry.cells)) {
      const shiftedRaw =
        cell.raw.startsWith('=') && dRow !== 0
          ? shiftFormulaReferences(cell.raw, dRow, 0)
          : cell.raw;
      nextCells[cellKey(targetRow, Number(cStr))] = {...cell, raw: shiftedRaw};
    }
  });

  return {...sheet, cells: nextCells};
}

export function evaluateConditionalFormatting(
  row: number,
  col: number,
  evaluated: EvaluatedCell | undefined,
  rules: ConditionalRule[],
): {bg?: string; color?: string} | null {
  if (!evaluated || evaluated.value === null || evaluated.value === '') return null;
  for (const rule of rules) {
    if (!isCoordInRange(row, col, rule.range)) continue;
    const val = evaluated.value;
    const numVal = typeof val === 'number' ? val : Number(val);
    const n1 = Number(rule.value1);
    const n2 = Number(rule.value2 ?? '');
    let matched = false;
    switch (rule.operator) {
      case 'gt':
        matched = Number.isFinite(numVal) && Number.isFinite(n1) && numVal > n1;
        break;
      case 'lt':
        matched = Number.isFinite(numVal) && Number.isFinite(n1) && numVal < n1;
        break;
      case 'between':
        matched =
          Number.isFinite(numVal) &&
          Number.isFinite(n1) &&
          Number.isFinite(n2) &&
          numVal >= Math.min(n1, n2) &&
          numVal <= Math.max(n1, n2);
        break;
      case 'eq':
        matched = String(val).toLowerCase() === rule.value1.trim().toLowerCase();
        break;
      case 'contains':
        matched = String(val).toLowerCase().includes(rule.value1.trim().toLowerCase());
        break;
    }
    if (matched) return {bg: rule.bg, color: rule.color};
  }
  return null;
}

export function getValidationForCell(
  row: number,
  col: number,
  validations: DataValidationRule[],
): DataValidationRule | null {
  for (let i = validations.length - 1; i >= 0; i--) {
    if (isCoordInRange(row, col, validations[i].range)) {
      return validations[i];
    }
  }
  return null;
}

export function validateCellInput(
  raw: string,
  rule: DataValidationRule | null,
): {valid: boolean; message?: string} {
  if (!rule || rule.kind === 'none' || raw.trim() === '' || raw.startsWith('=')) {
    return {valid: true};
  }
  if (rule.kind === 'number-range') {
    const n = Number(raw);
    if (!Number.isFinite(n)) {
      return {valid: false, message: 'Value must be a valid number.'};
    }
    if (rule.min !== undefined && n < rule.min) {
      return {valid: false, message: `Value must be at least ${rule.min}.`};
    }
    if (rule.max !== undefined && n > rule.max) {
      return {valid: false, message: `Value must be at most ${rule.max}.`};
    }
  } else if (rule.kind === 'list') {
    const allowed = rule.listItems ?? [];
    if (
      allowed.length > 0 &&
      !allowed.some((item) => item.toLowerCase() === raw.trim().toLowerCase())
    ) {
      return {valid: false, message: `Value must be one of: ${allowed.join(', ')}`};
    }
  } else if (rule.kind === 'text-length') {
    const len = raw.length;
    if (rule.min !== undefined && len < rule.min) {
      return {valid: false, message: `Text must be at least ${rule.min} characters.`};
    }
    if (rule.max !== undefined && len > rule.max) {
      return {valid: false, message: `Text must be at most ${rule.max} characters.`};
    }
  }
  return {valid: true};
}

export function findMergeAt(row: number, col: number, merges: MergeRange[]): MergeRange | null {
  for (const m of merges) {
    if (row >= m.startRow && row <= m.endRow && col >= m.startCol && col <= m.endCol) {
      return m;
    }
  }
  return null;
}

/**
 * Exports all populated sheets into a clean tabular PDF document via pdf-lib.
 */
export async function workbookToPdfBytes(
  workbook: WorkbookData,
  title: string,
): Promise<Uint8Array> {
  const evalMaps = evaluateWorkbook(workbook);
  const pdf = await PDFDocument.create();
  pdf.setTitle(title);
  pdf.setCreator('eSZ OFFICE');

  const regular = await pdf.embedFont(StandardFonts.Helvetica);
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold);

  const pageW = 841.89; // A4 Landscape
  const pageH = 595.28;
  const marginX = 36;
  const marginY = 44;
  const sanitize = (s: string) => s.replace(/[^\x20-\x7E]/g, ' ');

  for (const sheet of workbook.sheets) {
    const evalMap = evalMaps[sheet.id] ?? {};
    let maxR = -1;
    let maxC = -1;
    for (const [k, cell] of Object.entries(sheet.cells)) {
      if (!cell.raw) continue;
      const [r, c] = k.split(',').map(Number);
      if (r > maxR) maxR = r;
      if (c > maxC) maxC = c;
    }

    let page = pdf.addPage([pageW, pageH]);
    let y = pageH - marginY;

    const drawHeader = () => {
      page.drawText(sanitize(`${title} — ${sheet.name}`), {
        x: marginX,
        y,
        size: 13,
        font: bold,
        color: rgb(0.11, 0.23, 0.54),
      });
      y -= 22;
    };
    drawHeader();

    if (maxR < 0 || maxC < 0) {
      page.drawText('(Empty sheet)', {
        x: marginX,
        y,
        size: 10,
        font: regular,
        color: rgb(0.45, 0.5, 0.58),
      });
      continue;
    }

    const colCount = Math.min(12, maxC + 1);
    const colW = (pageW - marginX * 2) / colCount;
    const rowH = 19;

    // Column letter band
    page.drawRectangle({
      x: marginX,
      y: y - rowH + 4,
      width: pageW - marginX * 2,
      height: rowH,
      color: rgb(0.94, 0.96, 0.98),
      borderColor: rgb(0.8, 0.84, 0.9),
      borderWidth: 0.6,
    });
    for (let c = 0; c < colCount; c++) {
      page.drawText(colIndexToLetter(c), {
        x: marginX + c * colW + 4,
        y: y - 9,
        size: 8.5,
        font: bold,
        color: rgb(0.35, 0.4, 0.48),
      });
    }
    y -= rowH;

    for (let r = 0; r <= maxR; r++) {
      if (y < marginY + rowH + 16) {
        page = pdf.addPage([pageW, pageH]);
        y = pageH - marginY;
        drawHeader();
      }
      for (let c = 0; c < colCount; c++) {
        const k = cellKey(r, c);
        const cell = sheet.cells[k];
        const evaluated = evalMap[k];
        const display = sanitize(evaluated?.display ?? '').slice(0, Math.max(6, Math.floor(colW / 5.4)));
        page.drawRectangle({
          x: marginX + c * colW,
          y: y - rowH + 4,
          width: colW,
          height: rowH,
          color: r === 0 ? rgb(0.96, 0.97, 0.99) : rgb(1, 1, 1),
          borderColor: rgb(0.85, 0.88, 0.92),
          borderWidth: 0.5,
        });
        if (display) {
          page.drawText(display, {
            x: marginX + c * colW + 4,
            y: y - 9,
            size: 8.5,
            font: r === 0 || cell?.style?.bold ? bold : regular,
            color: rgb(0.08, 0.12, 0.2),
          });
        }
      }
      y -= rowH;
    }
  }

  const pages = pdf.getPages();
  pages.forEach((p, idx) => {
    p.drawText(`Page ${idx + 1} of ${pages.length} · Generated offline by eSZ OFFICE`, {
      x: marginX,
      y: 18,
      size: 8,
      font: regular,
      color: rgb(0.45, 0.5, 0.58),
    });
  });

  return pdf.save();
}
