export type SheetsRibbonTab = 'home' | 'insert' | 'data' | 'view';

export type CellNumberFormat =
  | 'general'
  | 'number'
  | 'currency'
  | 'percent'
  | 'date'
  | 'scientific'
  | 'text';

export interface CellStyle {
  fontFamily?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
  bg?: string;
  align?: 'left' | 'center' | 'right';
  wrap?: boolean;
  border?: boolean;
  format?: CellNumberFormat;
  decimals?: number;
}

export interface CellData {
  /** Raw user input (plain value or formula starting with '=') */
  raw: string;
  style?: CellStyle;
}

export interface MergeRange {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

export type ConditionalOperator = 'gt' | 'lt' | 'eq' | 'between' | 'contains';

export interface ConditionalRule {
  id: string;
  range: string; // e.g., "B2:D10"
  operator: ConditionalOperator;
  value1: string;
  value2?: string;
  bg: string;
  color: string;
}

export type ValidationKind = 'none' | 'number-range' | 'list' | 'text-length';

export interface DataValidationRule {
  id: string;
  range: string; // e.g., "A2:A20"
  kind: ValidationKind;
  min?: number;
  max?: number;
  listItems?: string[];
  prompt?: string;
}

export type ChartKind = 'bar' | 'line' | 'pie' | 'scatter';

export interface ChartSpec {
  id: string;
  title: string;
  kind: ChartKind;
  range: string; // e.g., "A1:C4"
}

export interface SheetData {
  id: string;
  name: string;
  rowCount: number;
  colCount: number;
  /** Sparse map keyed by "r,c" (0-indexed) */
  cells: Record<string, CellData>;
  colWidths: Record<number, number>;
  rowHeights: Record<number, number>;
  merges: MergeRange[];
  freezeRows: number;
  freezeCols: number;
  autoFilterEnabled: boolean;
  /** Column index -> selected allowed values (empty/undefined means show all) */
  columnFilters: Record<number, string[]>;
  conditionalRules: ConditionalRule[];
  validations: DataValidationRule[];
  charts: ChartSpec[];
}

export interface WorkbookData {
  activeSheetId: string;
  sheets: SheetData[];
}

export interface CellCoord {
  row: number;
  col: number;
}

export interface SelectionRange {
  start: CellCoord;
  end: CellCoord;
}

export type EvaluatedScalar = number | string | boolean | null;

export interface EvaluatedCell {
  value: EvaluatedScalar;
  error?: string;
  display: string;
}

export interface PersistedSheetMeta {
  version: 1;
  activeSheetId: string;
  sheetsMeta: Record<
    string,
    {
      styles: Record<string, CellStyle>;
      colWidths: Record<number, number>;
      rowHeights: Record<number, number>;
      merges: MergeRange[];
      freezeRows: number;
      freezeCols: number;
      autoFilterEnabled: boolean;
      conditionalRules: ConditionalRule[];
      validations: DataValidationRule[];
      charts: ChartSpec[];
    }
  >;
}

export const DEFAULT_ROW_COUNT = 40;
export const DEFAULT_COL_COUNT = 16;
export const DEFAULT_COL_WIDTH = 108;
export const DEFAULT_ROW_HEIGHT = 28;

export const CHART_COLORS = [
  '#1D4ED8',
  '#059669',
  '#D97706',
  '#DC2626',
  '#7C3AED',
  '#0891B2',
  '#DB2777',
  '#4F46E5',
];

export const CONDITIONAL_PRESETS: Array<{label: string; bg: string; color: string}> = [
  {label: 'Green Positive', bg: '#DCFCE7', color: '#166534'},
  {label: 'Red Alert', bg: '#FEE2E2', color: '#991B1B'},
  {label: 'Amber Warning', bg: '#FEF3C7', color: '#92400E'},
  {label: 'Blue Accent', bg: '#DBEAFE', color: '#1E40AF'},
];

export function cellKey(row: number, col: number): string {
  return `${row},${col}`;
}

export function normalizeSelection(sel: SelectionRange): {
  minRow: number;
  maxRow: number;
  minCol: number;
  maxCol: number;
} {
  return {
    minRow: Math.min(sel.start.row, sel.end.row),
    maxRow: Math.max(sel.start.row, sel.end.row),
    minCol: Math.min(sel.start.col, sel.end.col),
    maxCol: Math.max(sel.start.col, sel.end.col),
  };
}
