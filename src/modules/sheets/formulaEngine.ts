import {
  cellKey,
  type CellNumberFormat,
  type EvaluatedCell,
  type EvaluatedScalar,
  type SheetData,
  type WorkbookData,
} from './sheetsTypes';

export function colIndexToLetter(col: number): string {
  let n = col + 1;
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function letterToColIndex(letters: string): number {
  const clean = letters.toUpperCase();
  let acc = 0;
  for (let i = 0; i < clean.length; i++) {
    acc = acc * 26 + (clean.charCodeAt(i) - 64);
  }
  return acc - 1;
}

export function formatCellRef(row: number, col: number): string {
  return `${colIndexToLetter(col)}${row + 1}`;
}

export function parseCellRef(ref: string): {
  row: number;
  col: number;
  absRow: boolean;
  absCol: boolean;
} | null {
  const m = ref.trim().match(/^(\$?)([A-Za-z]+)(\$?)(\d+)$/);
  if (!m) return null;
  const col = letterToColIndex(m[2]);
  const row = Number.parseInt(m[4], 10) - 1;
  if (row < 0 || col < 0) return null;
  return {row, col, absCol: m[1] === '$', absRow: m[3] === '$'};
}

export function parseRangeRef(range: string): {
  sheetName?: string;
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
} | null {
  let target = range.trim();
  let sheetName: string | undefined;
  const bangIdx = target.indexOf('!');
  if (bangIdx !== -1) {
    sheetName = target.slice(0, bangIdx).replace(/^'|'$/g, '');
    target = target.slice(bangIdx + 1);
  }
  const parts = target.split(':');
  const start = parseCellRef(parts[0]);
  if (!start) return null;
  const end = parts[1] ? parseCellRef(parts[1]) : start;
  if (!end) return null;
  return {
    sheetName,
    startRow: Math.min(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endRow: Math.max(start.row, end.row),
    endCol: Math.max(start.col, end.col),
  };
}

export function isCoordInRange(row: number, col: number, rangeStr: string): boolean {
  const parsed = parseRangeRef(rangeStr);
  if (!parsed) return false;
  return (
    row >= parsed.startRow &&
    row <= parsed.endRow &&
    col >= parsed.startCol &&
    col <= parsed.endCol
  );
}

/**
 * Adjusts relative A1 references inside a formula when filling/copying across cells.
 * Absolute references ($A$1) stay locked.
 */
export function shiftFormulaReferences(formula: string, dRow: number, dCol: number): string {
  if (!formula.startsWith('=')) return formula;
  return formula.replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)/g, (match, absC, colL, absR, rowStr, offset, src) => {
    // Avoid modifying text inside double quotes
    const prefix = src.slice(0, offset);
    const quoteCount = (prefix.match(/"/g) || []).length;
    if (quoteCount % 2 === 1) return match;
    // Avoid function names immediately followed by '('
    if (src[offset + match.length] === '(') return match;

    const col = letterToColIndex(colL);
    const row = Number.parseInt(rowStr, 10) - 1;
    const nextCol = absC ? col : Math.max(0, col + dCol);
    const nextRow = absR ? row : Math.max(0, row + dRow);
    return `${absC}${colIndexToLetter(nextCol)}${absR}${nextRow + 1}`;
  });
}

class FormulaError extends Error {
  code: string;
  constructor(code: string) {
    super(code);
    this.code = code;
  }
}

type TokenType = 'num' | 'str' | 'ident' | 'op' | 'lparen' | 'rparen' | 'comma' | 'colon' | 'eof';

interface Token {
  type: TokenType;
  value: string;
}

function tokenizeFormula(input: string): Token[] {
  const tokens: Token[] = [];
  let i = 0;
  const len = input.length;

  while (i < len) {
    const ch = input[i];
    if (/\s/.test(ch)) {
      i++;
      continue;
    }

    // String literal
    if (ch === '"') {
      let str = '';
      i++;
      while (i < len) {
        if (input[i] === '"' && input[i + 1] === '"') {
          str += '"';
          i += 2;
        } else if (input[i] === '"') {
          i++;
          break;
        } else {
          str += input[i++];
        }
      }
      tokens.push({type: 'str', value: str});
      continue;
    }

    // Quoted sheet name e.g. 'Q4 Budget'!A1
    if (ch === "'") {
      let sheetRef = "'";
      i++;
      while (i < len && input[i] !== "'") {
        sheetRef += input[i++];
      }
      if (i < len) sheetRef += input[i++]; // closing quote
      if (input[i] === '!') {
        sheetRef += input[i++];
        while (i < len && /[\w$]/.test(input[i])) {
          sheetRef += input[i++];
        }
      }
      tokens.push({type: 'ident', value: sheetRef});
      continue;
    }

    // Numbers
    if (/\d/.test(ch) || (ch === '.' && /\d/.test(input[i + 1] || ''))) {
      let num = '';
      while (i < len && /[\d.]/.test(input[i])) {
        num += input[i++];
      }
      if ((input[i] === 'e' || input[i] === 'E') && /[+-]?\d/.test(input.slice(i + 1))) {
        num += input[i++];
        if (input[i] === '+' || input[i] === '-') num += input[i++];
        while (i < len && /\d/.test(input[i])) num += input[i++];
      }
      if (input[i] === '%') {
        i++;
        tokens.push({type: 'num', value: String(Number(num) / 100)});
      } else {
        tokens.push({type: 'num', value: num});
      }
      continue;
    }

    // Two-char comparison operators
    const two = input.slice(i, i + 2);
    if (two === '<>' || two === '<=' || two === '>=') {
      tokens.push({type: 'op', value: two});
      i += 2;
      continue;
    }

    if ('+-*/^=&<>'.includes(ch)) {
      tokens.push({type: 'op', value: ch});
      i++;
      continue;
    }
    if (ch === '(') {
      tokens.push({type: 'lparen', value: '('});
      i++;
      continue;
    }
    if (ch === ')') {
      tokens.push({type: 'rparen', value: ')'});
      i++;
      continue;
    }
    if (ch === ',') {
      tokens.push({type: 'comma', value: ','});
      i++;
      continue;
    }
    if (ch === ':') {
      tokens.push({type: 'colon', value: ':'});
      i++;
      continue;
    }

    // Identifiers, cell refs, cross-sheet refs (Sheet2!A1)
    if (/[A-Za-z_$]/.test(ch)) {
      let ident = '';
      while (i < len && /[A-Za-z0-9_$!.?]/.test(input[i])) {
        ident += input[i++];
      }
      tokens.push({type: 'ident', value: ident});
      continue;
    }

    throw new FormulaError('#NAME?');
  }

  tokens.push({type: 'eof', value: ''});
  return tokens;
}

type EvalArg = EvaluatedScalar | EvaluatedScalar[][];

function isMatrix(v: EvalArg): v is EvaluatedScalar[][] {
  return Array.isArray(v);
}

function toScalar(v: EvalArg): EvaluatedScalar {
  if (isMatrix(v)) {
    return v[0]?.[0] ?? null;
  }
  return v;
}

function toNumber(v: EvalArg): number {
  const s = toScalar(v);
  if (s === null || s === '') return 0;
  if (typeof s === 'number') return s;
  if (typeof s === 'boolean') return s ? 1 : 0;
  const parsed = Number(String(s).replace(/[$,%\s]/g, ''));
  if (!Number.isFinite(parsed)) throw new FormulaError('#VALUE!');
  return parsed;
}

function toStringVal(v: EvalArg): string {
  const s = toScalar(v);
  if (s === null) return '';
  if (typeof s === 'boolean') return s ? 'TRUE' : 'FALSE';
  return String(s);
}

function toBool(v: EvalArg): boolean {
  const s = toScalar(v);
  if (typeof s === 'boolean') return s;
  if (typeof s === 'number') return s !== 0;
  if (typeof s === 'string') {
    const u = s.trim().toUpperCase();
    if (u === 'TRUE') return true;
    if (u === 'FALSE' || u === '') return false;
  }
  return Boolean(s);
}

function flattenArgs(args: EvalArg[]): EvaluatedScalar[] {
  const out: EvaluatedScalar[] = [];
  for (const a of args) {
    if (isMatrix(a)) {
      for (const row of a) {
        for (const cell of row) out.push(cell);
      }
    } else {
      out.push(a);
    }
  }
  return out;
}

function matchesCriterion(val: EvaluatedScalar, rawCrit: EvaluatedScalar): boolean {
  if (rawCrit === null) return val === null || val === '';
  const crit = String(rawCrit).trim();
  const opMatch = crit.match(/^(<>|<=|>=|<|>|=)(.*)$/);
  if (opMatch) {
    const op = opMatch[1];
    const rhsRaw = opMatch[2].trim();
    const rhsNum = Number(rhsRaw);
    const lhsNum = typeof val === 'number' ? val : Number(val);
    if (Number.isFinite(rhsNum) && rhsRaw !== '' && val !== null && val !== '' && Number.isFinite(lhsNum)) {
      if (op === '>') return lhsNum > rhsNum;
      if (op === '>=') return lhsNum >= rhsNum;
      if (op === '<') return lhsNum < rhsNum;
      if (op === '<=') return lhsNum <= rhsNum;
      if (op === '=') return lhsNum === rhsNum;
      if (op === '<>') return lhsNum !== rhsNum;
    }
    const lhsStr = val === null ? '' : String(val).toLowerCase();
    const rhsStr = rhsRaw.toLowerCase();
    if (op === '=') return lhsStr === rhsStr;
    if (op === '<>') return lhsStr !== rhsStr;
    return false;
  }
  if (typeof val === 'number' && typeof rawCrit === 'number') {
    return val === rawCrit;
  }
  return String(val ?? '').toLowerCase() === crit.toLowerCase();
}

function parseRawScalar(raw: string): EvaluatedScalar {
  const trimmed = raw.trim();
  if (trimmed === '') return null;
  if (trimmed.toUpperCase() === 'TRUE') return true;
  if (trimmed.toUpperCase() === 'FALSE') return false;
  if (/^-?\d+(\.\d+)?([eE][+-]?\d+)?$/.test(trimmed)) {
    const n = Number(trimmed);
    if (Number.isFinite(n)) return n;
  }
  return raw;
}

export function formatCellValue(
  value: EvaluatedScalar,
  format: CellNumberFormat = 'general',
  decimals = 2,
): string {
  if (value === null || value === undefined) return '';
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value === 'string') return value;
  if (!Number.isFinite(value)) return '#NUM!';

  const dec = Math.max(0, Math.min(8, decimals));

  switch (format) {
    case 'currency':
      return value.toLocaleString('en-US', {
        style: 'currency',
        currency: 'USD',
        minimumFractionDigits: dec,
        maximumFractionDigits: dec,
      });
    case 'percent':
      return `${(value * 100).toLocaleString('en-US', {
        minimumFractionDigits: dec,
        maximumFractionDigits: dec,
      })}%`;
    case 'number':
      return value.toLocaleString('en-US', {
        minimumFractionDigits: dec,
        maximumFractionDigits: dec,
      });
    case 'scientific':
      return value.toExponential(dec).toUpperCase();
    case 'date': {
      // Excel serial date (days since 1899-12-30) or JS timestamp
      const ms = value > 1e10 ? value : Date.UTC(1899, 11, 30) + value * 86_400_000;
      const d = new Date(ms);
      if (Number.isNaN(d.getTime())) return String(value);
      return d.toISOString().slice(0, 10);
    }
    case 'text':
      return String(value);
    case 'general':
    default: {
      if (Number.isInteger(value)) return String(value);
      return String(Number(value.toFixed(6)));
    }
  }
}

/**
 * Evaluates all formulas across every sheet in a workbook with cycle detection and memoization.
 * Returns a map: sheetId -> cellKey("r,c") -> EvaluatedCell.
 */
export function evaluateWorkbook(
  workbook: WorkbookData,
): Record<string, Record<string, EvaluatedCell>> {
  const sheetById = new Map<string, SheetData>();
  const sheetByLowerName = new Map<string, SheetData>();
  for (const s of workbook.sheets) {
    sheetById.set(s.id, s);
    sheetByLowerName.set(s.name.toLowerCase(), s);
  }

  const cache = new Map<string, {value: EvaluatedScalar; error?: string}>();
  const visiting = new Set<string>();

  const evalCell = (sheetId: string, row: number, col: number): {value: EvaluatedScalar; error?: string} => {
    const k = `${sheetId}:${row},${col}`;
    const hit = cache.get(k);
    if (hit) return hit;

    if (visiting.has(k)) {
      return {value: '#CYCLE!', error: '#CYCLE!'};
    }

    const sheet = sheetById.get(sheetId);
    const cell = sheet?.cells[cellKey(row, col)];
    if (!cell || cell.raw === '') {
      const empty = {value: null};
      cache.set(k, empty);
      return empty;
    }

    if (!cell.raw.startsWith('=')) {
      const scalar = {value: parseRawScalar(cell.raw)};
      cache.set(k, scalar);
      return scalar;
    }

    visiting.add(k);
    let result: {value: EvaluatedScalar; error?: string};
    try {
      const val = evaluateFormulaExpression(cell.raw.slice(1), sheetId);
      const scalar = toScalar(val);
      if (typeof scalar === 'string' && scalar.startsWith('#')) {
        result = {value: scalar, error: scalar};
      } else {
        result = {value: scalar};
      }
    } catch (err) {
      const code = err instanceof FormulaError ? err.code : '#ERROR!';
      result = {value: code, error: code};
    } finally {
      visiting.delete(k);
    }

    cache.set(k, result);
    return result;
  };

  const resolveSheet = (name: string | undefined, currentSheetId: string): SheetData => {
    if (!name) {
      const s = sheetById.get(currentSheetId);
      if (!s) throw new FormulaError('#REF!');
      return s;
    }
    const found = sheetByLowerName.get(name.toLowerCase());
    if (!found) throw new FormulaError('#REF!');
    return found;
  };

  const readSingleRef = (tokenVal: string, currentSheetId: string): EvaluatedScalar => {
    let sheetName: string | undefined;
    let refPart = tokenVal;
    const bang = tokenVal.indexOf('!');
    if (bang !== -1) {
      sheetName = tokenVal.slice(0, bang).replace(/^'|'$/g, '');
      refPart = tokenVal.slice(bang + 1);
    }
    const parsed = parseCellRef(refPart);
    if (!parsed) throw new FormulaError('#NAME?');
    const targetSheet = resolveSheet(sheetName, currentSheetId);
    const res = evalCell(targetSheet.id, parsed.row, parsed.col);
    if (res.error) throw new FormulaError(res.error);
    return res.value;
  };

  const readRangeMatrix = (
    leftToken: string,
    rightToken: string,
    currentSheetId: string,
  ): EvaluatedScalar[][] => {
    let sheetName: string | undefined;
    let startPart = leftToken;
    const bang = leftToken.indexOf('!');
    if (bang !== -1) {
      sheetName = leftToken.slice(0, bang).replace(/^'|'$/g, '');
      startPart = leftToken.slice(bang + 1);
    }
    const endPart = rightToken.includes('!') ? rightToken.split('!')[1] : rightToken;
    const s = parseCellRef(startPart);
    const e = parseCellRef(endPart);
    if (!s || !e) throw new FormulaError('#REF!');
    const targetSheet = resolveSheet(sheetName, currentSheetId);

    const r1 = Math.min(s.row, e.row);
    const r2 = Math.max(s.row, e.row);
    const c1 = Math.min(s.col, e.col);
    const c2 = Math.max(s.col, e.col);
    const matrix: EvaluatedScalar[][] = [];
    for (let r = r1; r <= r2; r++) {
      const rowVals: EvaluatedScalar[] = [];
      for (let c = c1; c <= c2; c++) {
        const res = evalCell(targetSheet.id, r, c);
        if (res.error) throw new FormulaError(res.error);
        rowVals.push(res.value);
      }
      matrix.push(rowVals);
    }
    return matrix;
  };

  const callFunction = (nameRaw: string, rawArgNodes: Array<() => EvalArg>): EvalArg => {
    const name = nameRaw.toUpperCase();

    // Short-circuiting functions first
    if (name === 'IF') {
      if (rawArgNodes.length < 2) throw new FormulaError('#VALUE!');
      const cond = toBool(rawArgNodes[0]());
      if (cond) return rawArgNodes[1]();
      return rawArgNodes[2] ? rawArgNodes[2]() : false;
    }

    if (name === 'IFERROR') {
      if (rawArgNodes.length < 2) throw new FormulaError('#VALUE!');
      try {
        const v = rawArgNodes[0]();
        const s = toScalar(v);
        if (typeof s === 'string' && s.startsWith('#')) return rawArgNodes[1]();
        return v;
      } catch {
        return rawArgNodes[1]();
      }
    }

    if (name === 'IFS') {
      for (let i = 0; i + 1 < rawArgNodes.length; i += 2) {
        if (toBool(rawArgNodes[i]())) return rawArgNodes[i + 1]();
      }
      throw new FormulaError('#N/A');
    }

    const args = rawArgNodes.map((fn) => fn());

    switch (name) {
      case 'SUM': {
        return flattenArgs(args).reduce<number>(
          (acc, v) => (typeof v === 'number' ? acc + v : acc),
          0,
        );
      }
      case 'AVERAGE': {
        const nums = flattenArgs(args).filter((v): v is number => typeof v === 'number');
        if (nums.length === 0) throw new FormulaError('#DIV/0!');
        return nums.reduce((a, b) => a + b, 0) / nums.length;
      }
      case 'MIN': {
        const nums = flattenArgs(args).filter((v): v is number => typeof v === 'number');
        return nums.length === 0 ? 0 : Math.min(...nums);
      }
      case 'MAX': {
        const nums = flattenArgs(args).filter((v): v is number => typeof v === 'number');
        return nums.length === 0 ? 0 : Math.max(...nums);
      }
      case 'MEDIAN': {
        const nums = flattenArgs(args)
          .filter((v): v is number => typeof v === 'number')
          .sort((a, b) => a - b);
        if (nums.length === 0) throw new FormulaError('#NUM!');
        const mid = Math.floor(nums.length / 2);
        return nums.length % 2 === 1 ? nums[mid] : (nums[mid - 1] + nums[mid]) / 2;
      }
      case 'PRODUCT': {
        const nums = flattenArgs(args).filter((v): v is number => typeof v === 'number');
        return nums.length === 0 ? 0 : nums.reduce((a, b) => a * b, 1);
      }
      case 'COUNT': {
        return flattenArgs(args).filter((v) => typeof v === 'number').length;
      }
      case 'COUNTA': {
        return flattenArgs(args).filter((v) => v !== null && v !== '').length;
      }
      case 'COUNTBLANK': {
        return flattenArgs(args).filter((v) => v === null || v === '').length;
      }
      case 'COUNTIF': {
        if (args.length < 2) throw new FormulaError('#VALUE!');
        const vals = flattenArgs([args[0]]);
        const crit = toScalar(args[1]);
        return vals.filter((v) => matchesCriterion(v, crit)).length;
      }
      case 'SUMIF': {
        if (args.length < 2) throw new FormulaError('#VALUE!');
        const rangeVals = flattenArgs([args[0]]);
        const crit = toScalar(args[1]);
        const sumVals = args[2] ? flattenArgs([args[2]]) : rangeVals;
        let total = 0;
        for (let i = 0; i < rangeVals.length; i++) {
          if (matchesCriterion(rangeVals[i], crit)) {
            const n = sumVals[i];
            if (typeof n === 'number') total += n;
          }
        }
        return total;
      }
      case 'AVERAGEIF': {
        if (args.length < 2) throw new FormulaError('#VALUE!');
        const rangeVals = flattenArgs([args[0]]);
        const crit = toScalar(args[1]);
        const avgVals = args[2] ? flattenArgs([args[2]]) : rangeVals;
        let sum = 0;
        let count = 0;
        for (let i = 0; i < rangeVals.length; i++) {
          if (matchesCriterion(rangeVals[i], crit) && typeof avgVals[i] === 'number') {
            sum += avgVals[i] as number;
            count++;
          }
        }
        if (count === 0) throw new FormulaError('#DIV/0!');
        return sum / count;
      }
      case 'AND': {
        const vals = flattenArgs(args).filter((v) => v !== null && v !== '');
        if (vals.length === 0) throw new FormulaError('#VALUE!');
        return vals.every((v) => toBool(v));
      }
      case 'OR': {
        const vals = flattenArgs(args).filter((v) => v !== null && v !== '');
        if (vals.length === 0) throw new FormulaError('#VALUE!');
        return vals.some((v) => toBool(v));
      }
      case 'NOT':
        return !toBool(args[0]);
      case 'ISBLANK': {
        const s = toScalar(args[0]);
        return s === null || s === '';
      }
      case 'ISNUMBER':
        return typeof toScalar(args[0]) === 'number';
      case 'ABS':
        return Math.abs(toNumber(args[0]));
      case 'ROUND': {
        const n = toNumber(args[0]);
        const d = args[1] !== undefined ? toNumber(args[1]) : 0;
        const factor = Math.pow(10, d);
        return Math.round(n * factor) / factor;
      }
      case 'ROUNDUP': {
        const n = toNumber(args[0]);
        const d = args[1] !== undefined ? toNumber(args[1]) : 0;
        const factor = Math.pow(10, d);
        return (n >= 0 ? Math.ceil(n * factor) : Math.floor(n * factor)) / factor;
      }
      case 'ROUNDDOWN':
      case 'INT': {
        const n = toNumber(args[0]);
        const d = name === 'INT' ? 0 : args[1] !== undefined ? toNumber(args[1]) : 0;
        const factor = Math.pow(10, d);
        return Math.trunc(n * factor) / factor;
      }
      case 'FLOOR': {
        const n = toNumber(args[0]);
        const sig = args[1] !== undefined ? toNumber(args[1]) : 1;
        if (sig === 0) throw new FormulaError('#DIV/0!');
        return Math.floor(n / sig) * sig;
      }
      case 'CEILING': {
        const n = toNumber(args[0]);
        const sig = args[1] !== undefined ? toNumber(args[1]) : 1;
        if (sig === 0) throw new FormulaError('#DIV/0!');
        return Math.ceil(n / sig) * sig;
      }
      case 'POWER':
        return Math.pow(toNumber(args[0]), toNumber(args[1]));
      case 'SQRT': {
        const n = toNumber(args[0]);
        if (n < 0) throw new FormulaError('#NUM!');
        return Math.sqrt(n);
      }
      case 'MOD': {
        const n = toNumber(args[0]);
        const d = toNumber(args[1]);
        if (d === 0) throw new FormulaError('#DIV/0!');
        return ((n % d) + d) % d;
      }
      case 'VLOOKUP': {
        if (args.length < 3 || !isMatrix(args[1])) throw new FormulaError('#VALUE!');
        const needle = toScalar(args[0]);
        const table = args[1];
        const colIdx = Math.trunc(toNumber(args[2])) - 1;
        const approx = args[3] !== undefined ? toBool(args[3]) : true;
        if (colIdx < 0 || colIdx >= (table[0]?.length ?? 0)) throw new FormulaError('#REF!');
        if (!approx) {
          for (const row of table) {
            if (String(row[0] ?? '').toLowerCase() === String(needle ?? '').toLowerCase()) {
              return row[colIdx] ?? null;
            }
          }
          throw new FormulaError('#N/A');
        }
        let bestRow: EvaluatedScalar[] | null = null;
        for (const row of table) {
          if (typeof row[0] === 'number' && typeof needle === 'number') {
            if (row[0] <= needle) bestRow = row;
          } else if (String(row[0] ?? '').toLowerCase() === String(needle ?? '').toLowerCase()) {
            bestRow = row;
            break;
          }
        }
        if (!bestRow) throw new FormulaError('#N/A');
        return bestRow[colIdx] ?? null;
      }
      case 'HLOOKUP': {
        if (args.length < 3 || !isMatrix(args[1])) throw new FormulaError('#VALUE!');
        const needle = toScalar(args[0]);
        const table = args[1];
        const rowIdx = Math.trunc(toNumber(args[2])) - 1;
        if (rowIdx < 0 || rowIdx >= table.length) throw new FormulaError('#REF!');
        const header = table[0] ?? [];
        for (let c = 0; c < header.length; c++) {
          if (String(header[c] ?? '').toLowerCase() === String(needle ?? '').toLowerCase()) {
            return table[rowIdx][c] ?? null;
          }
        }
        throw new FormulaError('#N/A');
      }
      case 'XLOOKUP': {
        if (args.length < 3) throw new FormulaError('#VALUE!');
        const needle = toScalar(args[0]);
        const lookupArr = flattenArgs([args[1]]);
        const returnArr = flattenArgs([args[2]]);
        const notFound = args[3] !== undefined ? toScalar(args[3]) : undefined;
        const idx = lookupArr.findIndex(
          (v) => String(v ?? '').toLowerCase() === String(needle ?? '').toLowerCase(),
        );
        if (idx === -1) {
          if (notFound !== undefined) return notFound;
          throw new FormulaError('#N/A');
        }
        return returnArr[idx] ?? null;
      }
      case 'INDEX': {
        if (args.length < 2 || !isMatrix(args[0])) throw new FormulaError('#VALUE!');
        const mat = args[0];
        const r = Math.trunc(toNumber(args[1])) - 1;
        const c = args[2] !== undefined ? Math.trunc(toNumber(args[2])) - 1 : 0;
        if (r < 0 || r >= mat.length || c < 0 || c >= (mat[r]?.length ?? 0)) {
          throw new FormulaError('#REF!');
        }
        return mat[r][c] ?? null;
      }
      case 'MATCH': {
        if (args.length < 2) throw new FormulaError('#VALUE!');
        const needle = toScalar(args[0]);
        const arr = flattenArgs([args[1]]);
        const idx = arr.findIndex(
          (v) => String(v ?? '').toLowerCase() === String(needle ?? '').toLowerCase(),
        );
        if (idx === -1) throw new FormulaError('#N/A');
        return idx + 1;
      }
      case 'CONCAT':
      case 'CONCATENATE':
        return flattenArgs(args)
          .map((v) => (v === null ? '' : String(v)))
          .join('');
      case 'LEFT': {
        const s = toStringVal(args[0]);
        const n = args[1] !== undefined ? Math.max(0, Math.trunc(toNumber(args[1]))) : 1;
        return s.slice(0, n);
      }
      case 'RIGHT': {
        const s = toStringVal(args[0]);
        const n = args[1] !== undefined ? Math.max(0, Math.trunc(toNumber(args[1]))) : 1;
        return n === 0 ? '' : s.slice(-n);
      }
      case 'MID': {
        const s = toStringVal(args[0]);
        const start = Math.max(1, Math.trunc(toNumber(args[1]))) - 1;
        const count = Math.max(0, Math.trunc(toNumber(args[2])));
        return s.slice(start, start + count);
      }
      case 'LEN':
        return toStringVal(args[0]).length;
      case 'UPPER':
        return toStringVal(args[0]).toUpperCase();
      case 'LOWER':
        return toStringVal(args[0]).toLowerCase();
      case 'PROPER':
        return toStringVal(args[0]).replace(/\b\w/g, (ch) => ch.toUpperCase());
      case 'TRIM':
        return toStringVal(args[0]).trim().replace(/\s+/g, ' ');
      case 'SUBSTITUTE': {
        const text = toStringVal(args[0]);
        const oldStr = toStringVal(args[1]);
        const newStr = toStringVal(args[2]);
        return oldStr ? text.split(oldStr).join(newStr) : text;
      }
      case 'TEXT': {
        const val = toScalar(args[0]);
        const fmt = toStringVal(args[1]).toLowerCase();
        if (typeof val === 'number') {
          if (fmt.includes('$')) return formatCellValue(val, 'currency', 2);
          if (fmt.includes('%')) return formatCellValue(val, 'percent', 1);
          if (fmt.includes('0.00')) return formatCellValue(val, 'number', 2);
        }
        return String(val ?? '');
      }
      case 'TODAY': {
        const now = new Date();
        return new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()))
          .toISOString()
          .slice(0, 10);
      }
      case 'NOW':
        return new Date().toISOString().slice(0, 16).replace('T', ' ');
      case 'DATE': {
        const y = Math.trunc(toNumber(args[0]));
        const m = Math.trunc(toNumber(args[1])) - 1;
        const d = Math.trunc(toNumber(args[2]));
        return new Date(Date.UTC(y, m, d)).toISOString().slice(0, 10);
      }
      case 'YEAR': {
        const d = new Date(toStringVal(args[0]));
        if (Number.isNaN(d.getTime())) throw new FormulaError('#VALUE!');
        return d.getUTCFullYear();
      }
      case 'MONTH': {
        const d = new Date(toStringVal(args[0]));
        if (Number.isNaN(d.getTime())) throw new FormulaError('#VALUE!');
        return d.getUTCMonth() + 1;
      }
      case 'DAY': {
        const d = new Date(toStringVal(args[0]));
        if (Number.isNaN(d.getTime())) throw new FormulaError('#VALUE!');
        return d.getUTCDate();
      }
      case 'DAYS': {
        const end = new Date(toStringVal(args[0])).getTime();
        const start = new Date(toStringVal(args[1])).getTime();
        if (Number.isNaN(end) || Number.isNaN(start)) throw new FormulaError('#VALUE!');
        return Math.round((end - start) / 86_400_000);
      }
      case 'PMT': {
        const rate = toNumber(args[0]);
        const nper = toNumber(args[1]);
        const pv = toNumber(args[2]);
        const fv = args[3] !== undefined ? toNumber(args[3]) : 0;
        if (nper === 0) throw new FormulaError('#NUM!');
        if (rate === 0) return -(pv + fv) / nper;
        const factor = Math.pow(1 + rate, nper);
        return -(pv * factor + fv) * (rate / (factor - 1));
      }
      case 'FV': {
        const rate = toNumber(args[0]);
        const nper = toNumber(args[1]);
        const pmt = toNumber(args[2]);
        const pv = args[3] !== undefined ? toNumber(args[3]) : 0;
        if (rate === 0) return -(pv + pmt * nper);
        const factor = Math.pow(1 + rate, nper);
        return -(pv * factor + pmt * ((factor - 1) / rate));
      }
      case 'PV': {
        const rate = toNumber(args[0]);
        const nper = toNumber(args[1]);
        const pmt = toNumber(args[2]);
        const fv = args[3] !== undefined ? toNumber(args[3]) : 0;
        if (rate === 0) return -(fv + pmt * nper);
        const factor = Math.pow(1 + rate, nper);
        return -(fv + pmt * ((factor - 1) / rate)) / factor;
      }
      case 'NPV': {
        if (args.length < 2) throw new FormulaError('#VALUE!');
        const rate = toNumber(args[0]);
        const flows = flattenArgs(args.slice(1)).filter((v): v is number => typeof v === 'number');
        return flows.reduce((acc, cf, i) => acc + cf / Math.pow(1 + rate, i + 1), 0);
      }
      default:
        throw new FormulaError('#NAME?');
    }
  };

  const evaluateFormulaExpression = (expr: string, currentSheetId: string): EvalArg => {
    const tokens = tokenizeFormula(expr);
    let pos = 0;

    const peek = () => tokens[pos];
    const consume = () => tokens[pos++];

    const parseComparison = (): EvalArg => {
      let left = parseConcat();
      while (peek().type === 'op' && ['=', '<>', '<', '<=', '>', '>='].includes(peek().value)) {
        const op = consume().value;
        const right = parseConcat();
        const l = toScalar(left);
        const r = toScalar(right);
        if (op === '=') {
          left =
            typeof l === 'string' && typeof r === 'string'
              ? l.toLowerCase() === r.toLowerCase()
              : l === r;
        } else if (op === '<>') {
          left =
            typeof l === 'string' && typeof r === 'string'
              ? l.toLowerCase() !== r.toLowerCase()
              : l !== r;
        } else {
          const ln = toNumber(l);
          const rn = toNumber(r);
          if (op === '<') left = ln < rn;
          else if (op === '<=') left = ln <= rn;
          else if (op === '>') left = ln > rn;
          else if (op === '>=') left = ln >= rn;
        }
      }
      return left;
    };

    const parseConcat = (): EvalArg => {
      let left = parseAdditive();
      while (peek().type === 'op' && peek().value === '&') {
        consume();
        const right = parseAdditive();
        left = toStringVal(left) + toStringVal(right);
      }
      return left;
    };

    const parseAdditive = (): EvalArg => {
      let left = parseMultiplicative();
      while (peek().type === 'op' && (peek().value === '+' || peek().value === '-')) {
        const op = consume().value;
        const right = parseMultiplicative();
        const ln = toNumber(left);
        const rn = toNumber(right);
        left = op === '+' ? ln + rn : ln - rn;
      }
      return left;
    };

    const parseMultiplicative = (): EvalArg => {
      let left = parsePower();
      while (peek().type === 'op' && (peek().value === '*' || peek().value === '/')) {
        const op = consume().value;
        const right = parsePower();
        const ln = toNumber(left);
        const rn = toNumber(right);
        if (op === '/') {
          if (rn === 0) throw new FormulaError('#DIV/0!');
          left = ln / rn;
        } else {
          left = ln * rn;
        }
      }
      return left;
    };

    const parsePower = (): EvalArg => {
      let left = parseUnary();
      if (peek().type === 'op' && peek().value === '^') {
        consume();
        const right = parsePower();
        left = Math.pow(toNumber(left), toNumber(right));
      }
      return left;
    };

    const parseUnary = (): EvalArg => {
      if (peek().type === 'op' && peek().value === '+') {
        consume();
        return toNumber(parseUnary());
      }
      if (peek().type === 'op' && peek().value === '-') {
        consume();
        return -toNumber(parseUnary());
      }
      return parsePrimary();
    };

    const parsePrimary = (): EvalArg => {
      const tok = peek();
      if (tok.type === 'num') {
        consume();
        return Number(tok.value);
      }
      if (tok.type === 'str') {
        consume();
        return tok.value;
      }
      if (tok.type === 'lparen') {
        consume();
        const inner = parseComparison();
        if (peek().type === 'rparen') consume();
        return inner;
      }
      if (tok.type === 'ident') {
        consume();
        const upper = tok.value.toUpperCase();
        if (upper === 'TRUE' && peek().type !== 'lparen') return true;
        if (upper === 'FALSE' && peek().type !== 'lparen') return false;

        // Function call
        if (peek().type === 'lparen') {
          consume(); // '('
          // Capture lazy argument token spans so IF / IFERROR can short-circuit
          const argSlices: Token[][] = [];
          if (peek().type !== 'rparen') {
            let depth = 0;
            let current: Token[] = [];
            while (peek().type !== 'eof') {
              if (peek().type === 'rparen' && depth === 0) {
                consume();
                break;
              }
              if (peek().type === 'comma' && depth === 0) {
                consume();
                argSlices.push(current);
                current = [];
                continue;
              }
              if (peek().type === 'lparen') depth++;
              else if (peek().type === 'rparen') depth--;
              current.push(consume());
            }
            argSlices.push(current);
          } else {
            consume(); // ')'
          }

          const lazyEvaluators = argSlices.map((slice) => () => {
            if (slice.length === 0) return null;
            const subTokens = [...slice, {type: 'eof' as const, value: ''}];
            return evaluateSubTokens(subTokens);
          });
          return callFunction(tok.value, lazyEvaluators);
        }

        // Range reference A1:B5 or Sheet2!A1:B5
        if (peek().type === 'colon') {
          consume();
          const rightTok = consume();
          if (rightTok.type !== 'ident') throw new FormulaError('#REF!');
          return readRangeMatrix(tok.value, rightTok.value, currentSheetId);
        }

        // Single cell reference
        return readSingleRef(tok.value, currentSheetId);
      }

      throw new FormulaError('#VALUE!');
    };

    const evaluateSubTokens = (subTokens: Token[]): EvalArg => {
      const savedPos = pos;
      const savedTokens = tokens.slice();
      tokens.length = 0;
      tokens.push(...subTokens);
      pos = 0;
      try {
        return parseComparison();
      } finally {
        tokens.length = 0;
        tokens.push(...savedTokens);
        pos = savedPos;
      }
    };

    return parseComparison();
  };

  const out: Record<string, Record<string, EvaluatedCell>> = {};
  for (const sheet of workbook.sheets) {
    const map: Record<string, EvaluatedCell> = {};
    for (const [k, cell] of Object.entries(sheet.cells)) {
      const [rStr, cStr] = k.split(',');
      const r = Number(rStr);
      const c = Number(cStr);
      const res = evalCell(sheet.id, r, c);
      const display = res.error
        ? res.error
        : formatCellValue(res.value, cell.style?.format ?? 'general', cell.style?.decimals ?? 2);
      map[k] = {
        value: res.value,
        error: res.error,
        display,
      };
    }
    out[sheet.id] = map;
  }

  return out;
}
