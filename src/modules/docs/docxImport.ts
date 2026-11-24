import JSZip from 'jszip';
import {canonicalizeWordFontFamily, resolveRenderFontStack} from './docFonts';
import {
  computeBlockSignatureFromHtml,
  encodeUtf8Base64,
  injectHtmlAttrs,
} from './docxOoxmlBridge';
import type {
  DocBorderStyle,
  DocMarginPreset,
  DocPageBorderConfig,
  DocPageBorderSide,
  DocPageLayout,
  DocPageSize,
  DocTabStop,
} from './docsTypes';

export interface DocxImportWarning {
  code: string;
  feature: string;
  message: string;
  severity: 'info' | 'warning';
  count: number;
}

export interface DocxImportResult {
  html: string;
  layoutPatch: Partial<DocPageLayout>;
  warnings: DocxImportWarning[];
  hasLockedElements: boolean;
  hasLossyElements: boolean;
  stats: {
    paragraphs: number;
    tables: number;
    images: number;
    editableShapes: number;
    lockedElements: number;
    alternateContentResolved: number;
  };
}

// ---------------------------------------------------------------------------
// Universal zero-dependency XML DOM parser (identical in Browser & Node)
// ---------------------------------------------------------------------------

export interface XmlNode {
  type: 'element' | 'text';
  tag: string;
  localName: string;
  attrs: Record<string, string>;
  children: XmlNode[];
  text: string;
  rawXml?: string;
}

function decodeXmlEntities(raw: string): string {
  return raw
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex: string) =>
      String.fromCodePoint(parseInt(hex, 16)),
    )
    .replace(/&#(\d+);/g, (_, dec: string) => String.fromCodePoint(parseInt(dec, 10)))
    .replace(/&amp;/g, '&');
}

function escapeHtmlAttr(val: string): string {
  return val
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function escapeHtmlText(val: string): string {
  return val.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

export function parseXml(xml: string): XmlNode | null {
  if (!xml) return null;
  // Strip UTF-8 BOM, XML declarations, DOCTYPE, processing instructions, and comments
  const cleaned = xml
    .replace(/^\uFEFF/, '')
    .replace(/<\?[\s\S]*?\?>/g, '')
    .replace(/<!--[\s\S]*?-->/g, '')
    .replace(/<!DOCTYPE[\s\S]*?>/gi, '');

  const root: XmlNode = {
    type: 'element',
    tag: '#root',
    localName: '#root',
    attrs: {},
    children: [],
    text: '',
  };
  const stack: XmlNode[] = [root];
  const startOffsets: number[] = [0];

  const tokenRe = /<!\[CDATA\[([\s\S]*?)\]\]>|<(\/?)([A-Za-z_][\w:.-]*)([^>]*?)(\/?)>|([^<]+)/g;
  const attrRe = /([A-Za-z_][\w:.-]*)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

  let m: RegExpExecArray | null;
  while ((m = tokenRe.exec(cleaned)) !== null) {
    const parent = stack[stack.length - 1];
    if (m[1] !== undefined) {
      // CDATA
      parent.children.push({
        type: 'text',
        tag: '#text',
        localName: '#text',
        attrs: {},
        children: [],
        text: m[1],
      });
      continue;
    }
    if (m[6] !== undefined) {
      // Text node
      parent.children.push({
        type: 'text',
        tag: '#text',
        localName: '#text',
        attrs: {},
        children: [],
        text: decodeXmlEntities(m[6]),
      });
      continue;
    }

    const isClose = m[2] === '/';
    const tag = m[3];
    const attrStr = m[4] ?? '';
    const isSelfClose = m[5] === '/';

    if (isClose) {
      for (let i = stack.length - 1; i > 0; i--) {
        if (stack[i].tag === tag) {
          stack[i].rawXml = cleaned.slice(startOffsets[i], tokenRe.lastIndex);
          stack.length = i;
          startOffsets.length = i;
          break;
        }
      }
      continue;
    }

    const attrs: Record<string, string> = {};
    attrRe.lastIndex = 0;
    let am: RegExpExecArray | null;
    while ((am = attrRe.exec(attrStr)) !== null) {
      attrs[am[1]] = decodeXmlEntities(am[2] ?? am[3] ?? '');
    }

    const colonIdx = tag.indexOf(':');
    const localName = colonIdx >= 0 ? tag.slice(colonIdx + 1) : tag;
    const node: XmlNode = {
      type: 'element',
      tag,
      localName,
      attrs,
      children: [],
      text: '',
      ...(isSelfClose ? {rawXml: cleaned.slice(m.index, tokenRe.lastIndex)} : {}),
    };
    parent.children.push(node);
    if (!isSelfClose) {
      stack.push(node);
      startOffsets.push(m.index);
    }
  }

  return root.children.find((c) => c.type === 'element') ?? null;
}

function childEls(node: XmlNode | null | undefined, tagOrLocal?: string): XmlNode[] {
  if (!node) return [];
  const els = node.children.filter((c) => c.type === 'element');
  if (!tagOrLocal) return els;
  return els.filter((c) => c.tag === tagOrLocal || c.localName === tagOrLocal);
}

function firstChild(node: XmlNode | null | undefined, tagOrLocal: string): XmlNode | undefined {
  if (!node) return undefined;
  return node.children.find(
    (c) => c.type === 'element' && (c.tag === tagOrLocal || c.localName === tagOrLocal),
  );
}

function findDescendants(node: XmlNode | null | undefined, tagOrLocal: string): XmlNode[] {
  const out: XmlNode[] = [];
  if (!node) return out;
  const walk = (n: XmlNode) => {
    for (const c of n.children) {
      if (c.type !== 'element') continue;
      if (c.tag === tagOrLocal || c.localName === tagOrLocal) out.push(c);
      walk(c);
    }
  };
  walk(node);
  return out;
}

function attr(node: XmlNode | null | undefined, name: string): string | undefined {
  if (!node) return undefined;
  if (node.attrs[name] !== undefined) return node.attrs[name];
  // Check prefixed forms (e.g. w:val, r:embed, mc:Requires)
  const suffix = `:${name}`;
  for (const [k, v] of Object.entries(node.attrs)) {
    if (k === name || k.endsWith(suffix)) return v;
  }
  return undefined;
}

function textContent(node: XmlNode | null | undefined): string {
  if (!node) return '';
  if (node.type === 'text') return node.text;
  let s = '';
  for (const c of node.children) s += textContent(c);
  return s;
}

// ---------------------------------------------------------------------------
// Units & Color Helpers
// ---------------------------------------------------------------------------

const TWIPS_PER_PT = 20;
const TWIPS_PER_PX = 15; // 1440 twips = 1 in = 96 CSS px -> 15 twips/px
const EMU_PER_PX = 9525;
const PT_TO_PX = 96 / 72;

function num(val: string | undefined, fallback = 0): number {
  if (val == null || val === '') return fallback;
  const n = Number(val);
  return Number.isFinite(n) ? n : fallback;
}

function twipsToPt(twips: number): number {
  return Number((twips / TWIPS_PER_PT).toFixed(2));
}

function twipsToPx(twips: number): number {
  return Math.round(twips / TWIPS_PER_PX);
}

function normalizeColor(
  val: string | undefined,
  themeColors: Record<string, string>,
  themeAttr?: string,
  autoDefault?: string,
): string | undefined {
  if (val) {
    const t = val.trim();
    const low = t.toLowerCase();
    if (low === 'none' || low === 'nil') return undefined;
    if (low === 'auto') return autoDefault;
    if (/^[0-9a-fA-F]{6}$/.test(t)) return `#${t.toUpperCase()}`;
    if (/^#[0-9a-fA-F]{6}$/.test(t)) return t.toUpperCase();
  }
  if (themeAttr) {
    const mapped = themeColors[themeAttr.toLowerCase()];
    if (mapped) return mapped;
  }
  return undefined;
}

const WORD_HIGHLIGHT_MAP: Record<string, string> = {
  yellow: '#FEF08A',
  green: '#BBF7D0',
  cyan: '#A5F3FC',
  magenta: '#F5D0FE',
  blue: '#BFDBFE',
  red: '#FECACA',
  darkBlue: '#1E3A8A',
  darkCyan: '#155E75',
  darkGreen: '#166534',
  darkMagenta: '#701A75',
  darkRed: '#991B1B',
  darkYellow: '#CA8A04',
  darkGray: '#64748B',
  lightGray: '#E2E8F0',
  black: '#000000',
  white: '#FFFFFF',
};

function mapBorderCssStyle(val: string): string {
  switch (val.toLowerCase()) {
    case 'double':
    case 'triple':
    case 'thinthicksmallgap':
    case 'thickthinsmallgap':
      return 'double';
    case 'dashed':
    case 'dashsmallgap':
    case 'dotdash':
      return 'dashed';
    case 'dotted':
      return 'dotted';
    case 'ridge':
    case 'outset':
      return 'ridge';
    case 'groove':
    case 'inset':
      return 'groove';
    default:
      return 'solid';
  }
}

interface ParsedBorderSide {
  val: string;
  szPt: number;
  color: string;
  spacePt: number;
}

interface ParsedBorderSet {
  top?: ParsedBorderSide | 'none';
  right?: ParsedBorderSide | 'none';
  bottom?: ParsedBorderSide | 'none';
  left?: ParsedBorderSide | 'none';
  insideH?: ParsedBorderSide | 'none';
  insideV?: ParsedBorderSide | 'none';
}

function parseBorderSideEl(
  el: XmlNode | undefined,
  themeColors: Record<string, string>,
): ParsedBorderSide | 'none' | undefined {
  if (!el) return undefined;
  const val = (attr(el, 'val') ?? '').toLowerCase();
  if (!val || val === 'none' || val === 'nil') return 'none';
  const szEighths = num(attr(el, 'sz'), 4);
  const szPt = Math.max(0.5, Number((szEighths / 8).toFixed(2)));
  const color =
    normalizeColor(attr(el, 'color'), themeColors, attr(el, 'themeColor'), '#000000') ??
    '#000000';
  const spacePt = Math.max(0, num(attr(el, 'space'), 0));
  return {val, szPt, color, spacePt};
}

function parseBorderSet(
  el: XmlNode | undefined,
  themeColors: Record<string, string>,
): ParsedBorderSet {
  if (!el) return {};
  return {
    top: parseBorderSideEl(firstChild(el, 'w:top'), themeColors),
    right: parseBorderSideEl(firstChild(el, 'w:right') ?? firstChild(el, 'w:end'), themeColors),
    bottom: parseBorderSideEl(firstChild(el, 'w:bottom'), themeColors),
    left: parseBorderSideEl(firstChild(el, 'w:left') ?? firstChild(el, 'w:start'), themeColors),
    insideH: parseBorderSideEl(firstChild(el, 'w:insideH'), themeColors),
    insideV: parseBorderSideEl(firstChild(el, 'w:insideV'), themeColors),
  };
}

function borderSideToCss(side: ParsedBorderSide | 'none' | undefined): string | undefined {
  if (!side) return undefined;
  if (side === 'none') return 'none';
  return `${side.szPt}pt ${mapBorderCssStyle(side.val)} ${side.color}`;
}

// ---------------------------------------------------------------------------
// Theme, Style Inheritance, and Numbering Parsing
// ---------------------------------------------------------------------------

interface ThemeContext {
  majorFont: string | null;
  minorFont: string | null;
  colors: Record<string, string>;
}

function parseThemeXml(themeXml: string | null): ThemeContext {
  const ctx: ThemeContext = {
    majorFont: null,
    minorFont: null,
    colors: {
      dk1: '#000000',
      tx1: '#000000',
      lt1: '#FFFFFF',
      bg1: '#FFFFFF',
      dk2: '#44546A',
      lt2: '#E7E6E6',
      accent1: '#4472C4',
      accent2: '#ED7D31',
      accent3: '#A5A5A5',
      accent4: '#FFC000',
      accent5: '#5B9BD5',
      accent6: '#70AD47',
      hlink: '#0563C1',
      folhlink: '#954F72',
    },
  };
  if (!themeXml) return ctx;
  const root = parseXml(themeXml);
  if (!root) return ctx;

  const majorFont = findDescendants(root, 'a:majorFont')[0];
  const minorFont = findDescendants(root, 'a:minorFont')[0];
  const majorLatin = firstChild(majorFont, 'a:latin');
  const minorLatin = firstChild(minorFont, 'a:latin');
  ctx.majorFont = attr(majorLatin, 'typeface') ?? null;
  ctx.minorFont = attr(minorLatin, 'typeface') ?? null;

  const clrScheme = findDescendants(root, 'a:clrScheme')[0];
  if (clrScheme) {
    for (const child of childEls(clrScheme)) {
      const srgb = firstChild(child, 'a:srgbClr');
      const sysClr = firstChild(child, 'a:sysClr');
      const hex = attr(srgb, 'val') ?? attr(sysClr, 'lastClr');
      if (hex && /^[0-9a-fA-F]{6}$/.test(hex)) {
        const key = child.localName.toLowerCase();
        ctx.colors[key] = `#${hex.toUpperCase()}`;
        if (key === 'dk1') ctx.colors.tx1 = `#${hex.toUpperCase()}`;
        if (key === 'lt1') ctx.colors.bg1 = `#${hex.toUpperCase()}`;
      }
    }
  }

  return ctx;
}

interface ResolvedRunProps {
  fontFamily?: string;
  fontSizePt?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
  highlight?: string;
  vertAlign?: 'subscript' | 'superscript' | 'baseline';
  allCaps?: boolean;
  smallCaps?: boolean;
  letterSpacingPt?: number;
}

interface ResolvedParaProps {
  align?: 'left' | 'center' | 'right' | 'justify';
  indentLeftPx?: number;
  indentRightPx?: number;
  firstLineIndentPx?: number;
  spaceBeforePt?: number;
  spaceAfterPt?: number;
  lineSpacing?: string;
  keepNext?: boolean;
  keepLines?: boolean;
  pageBreakBefore?: boolean;
  outlineLvl?: number;
  numId?: string;
  ilvl?: number;
  tabStops?: DocTabStop[];
}

interface ResolvedTableProps {
  borders?: ParsedBorderSet;
  cellMarginsPx?: {top: number; right: number; bottom: number; left: number};
}

interface StyleDefinition {
  id: string;
  type: 'paragraph' | 'character' | 'table' | 'numbering';
  name: string;
  basedOn?: string;
  isDefault?: boolean;
  pPr: ResolvedParaProps;
  rPr: ResolvedRunProps;
  tblPr?: ResolvedTableProps;
}

function parseOnOff(el: XmlNode | undefined): boolean | undefined {
  if (!el) return undefined;
  const v = attr(el, 'val');
  if (v === undefined) return true;
  const low = v.toLowerCase();
  return low !== '0' && low !== 'false' && low !== 'off' && low !== 'none';
}

function parseRunPropsEl(rPr: XmlNode | undefined, theme: ThemeContext): ResolvedRunProps {
  const out: ResolvedRunProps = {};
  if (!rPr) return out;

  const rFonts = firstChild(rPr, 'w:rFonts');
  if (rFonts) {
    const explicit = attr(rFonts, 'ascii') || attr(rFonts, 'hAnsi');
    if (explicit) {
      out.fontFamily = canonicalizeWordFontFamily(explicit) ?? explicit;
    } else {
      const themeFont = (attr(rFonts, 'asciiTheme') || attr(rFonts, 'hAnsiTheme') || '').toLowerCase();
      if (themeFont.startsWith('major') && theme.majorFont) {
        out.fontFamily = canonicalizeWordFontFamily(theme.majorFont) ?? theme.majorFont;
      } else if (themeFont.startsWith('minor') && theme.minorFont) {
        out.fontFamily = canonicalizeWordFontFamily(theme.minorFont) ?? theme.minorFont;
      }
    }
  }

  const sz = firstChild(rPr, 'w:sz');
  if (sz) {
    const halfPts = num(attr(sz, 'val'), 0);
    if (halfPts > 0) out.fontSizePt = halfPts / 2;
  }

  const b = parseOnOff(firstChild(rPr, 'w:b'));
  if (b !== undefined) out.bold = b;

  const i = parseOnOff(firstChild(rPr, 'w:i'));
  if (i !== undefined) out.italic = i;

  const u = firstChild(rPr, 'w:u');
  if (u) {
    const val = (attr(u, 'val') ?? 'single').toLowerCase();
    out.underline = val !== 'none' && val !== '0';
  }

  const strike = parseOnOff(firstChild(rPr, 'w:strike') ?? firstChild(rPr, 'w:dstrike'));
  if (strike !== undefined) out.strike = strike;

  const colorEl = firstChild(rPr, 'w:color');
  if (colorEl) {
    const c = normalizeColor(
      attr(colorEl, 'val'),
      theme.colors,
      attr(colorEl, 'themeColor'),
      '#000000',
    );
    if (c) out.color = c;
  }
  if (!out.color) {
    // Word 2010+ textFill solid color
    const w14Fill = findDescendants(rPr, 'w14:srgbClr')[0];
    if (w14Fill) {
      const c = normalizeColor(attr(w14Fill, 'val'), theme.colors);
      if (c) out.color = c;
    }
  }

  const hlEl = firstChild(rPr, 'w:highlight');
  if (hlEl) {
    const hlVal = attr(hlEl, 'val') ?? '';
    if (hlVal && hlVal !== 'none') {
      out.highlight = WORD_HIGHLIGHT_MAP[hlVal] ?? hlVal;
    }
  } else {
    const shdEl = firstChild(rPr, 'w:shd');
    if (shdEl) {
      const fill = normalizeColor(attr(shdEl, 'fill'), theme.colors, attr(shdEl, 'themeFill'));
      if (fill && fill !== '#FFFFFF') out.highlight = fill;
    }
  }

  const vertAlign = firstChild(rPr, 'w:vertAlign');
  if (vertAlign) {
    const v = (attr(vertAlign, 'val') ?? '').toLowerCase();
    if (v === 'subscript') out.vertAlign = 'subscript';
    else if (v === 'superscript') out.vertAlign = 'superscript';
    else if (v === 'baseline') out.vertAlign = 'baseline';
  }

  const caps = parseOnOff(firstChild(rPr, 'w:caps'));
  if (caps !== undefined) out.allCaps = caps;

  const smallCaps = parseOnOff(firstChild(rPr, 'w:smallCaps'));
  if (smallCaps !== undefined) out.smallCaps = smallCaps;

  const spacing = firstChild(rPr, 'w:spacing');
  if (spacing) {
    const valTwips = num(attr(spacing, 'val'), 0);
    if (valTwips !== 0) out.letterSpacingPt = twipsToPt(valTwips);
  }

  return out;
}

function parseParaPropsEl(pPr: XmlNode | undefined): ResolvedParaProps {
  const out: ResolvedParaProps = {};
  if (!pPr) return out;

  const jcEl = firstChild(pPr, 'w:jc');
  if (jcEl) {
    const v = (attr(jcEl, 'val') ?? '').toLowerCase();
    if (v === 'center') out.align = 'center';
    else if (v === 'right' || v === 'end') out.align = 'right';
    else if (v === 'both' || v === 'distribute') out.align = 'justify';
    else if (v === 'left' || v === 'start') out.align = 'left';
  }

  const indEl = firstChild(pPr, 'w:ind');
  if (indEl) {
    const leftRaw = attr(indEl, 'left') ?? attr(indEl, 'start');
    const rightRaw = attr(indEl, 'right') ?? attr(indEl, 'end');
    const firstRaw = attr(indEl, 'firstLine');
    const hangingRaw = attr(indEl, 'hanging');
    if (leftRaw !== undefined) out.indentLeftPx = Math.max(0, twipsToPx(num(leftRaw, 0)));
    if (rightRaw !== undefined) out.indentRightPx = Math.max(0, twipsToPx(num(rightRaw, 0)));
    if (firstRaw !== undefined) out.firstLineIndentPx = twipsToPx(num(firstRaw, 0));
    else if (hangingRaw !== undefined) out.firstLineIndentPx = -Math.abs(twipsToPx(num(hangingRaw, 0)));
  }

  const spEl = firstChild(pPr, 'w:spacing');
  if (spEl) {
    const beforeRaw = attr(spEl, 'before');
    const afterRaw = attr(spEl, 'after');
    const lineRaw = attr(spEl, 'line');
    const lineRule = (attr(spEl, 'lineRule') ?? 'auto').toLowerCase();
    if (beforeRaw !== undefined) out.spaceBeforePt = Math.max(0, twipsToPt(num(beforeRaw, 0)));
    if (afterRaw !== undefined) out.spaceAfterPt = Math.max(0, twipsToPt(num(afterRaw, 0)));
    if (lineRaw !== undefined) {
      const lineVal = num(lineRaw, 240);
      if (lineVal > 0) {
        if (lineRule === 'exact' || lineRule === 'atleast') {
          out.lineSpacing = `${twipsToPt(lineVal)}pt`;
        } else {
          const factor = Number((lineVal / 240).toFixed(2));
          if (factor >= 0.6 && factor <= 4) {
            out.lineSpacing = String(factor);
          }
        }
      }
    }
  }

  const keepNext = parseOnOff(firstChild(pPr, 'w:keepNext'));
  if (keepNext !== undefined) out.keepNext = keepNext;

  const keepLines = parseOnOff(firstChild(pPr, 'w:keepLines'));
  if (keepLines !== undefined) out.keepLines = keepLines;

  const pageBreakBefore = parseOnOff(firstChild(pPr, 'w:pageBreakBefore'));
  if (pageBreakBefore !== undefined) out.pageBreakBefore = pageBreakBefore;

  const outlineLvl = firstChild(pPr, 'w:outlineLvl');
  if (outlineLvl) {
    const lvl = num(attr(outlineLvl, 'val'), -1);
    if (lvl >= 0 && lvl <= 8) out.outlineLvl = lvl;
  }

  const numPr = firstChild(pPr, 'w:numPr');
  if (numPr) {
    const numIdEl = firstChild(numPr, 'w:numId');
    const ilvlEl = firstChild(numPr, 'w:ilvl');
    const nid = attr(numIdEl, 'val');
    if (nid && nid !== '0') {
      out.numId = nid;
      out.ilvl = Math.max(0, num(attr(ilvlEl, 'val'), 0));
    }
  }

  const tabsEl = firstChild(pPr, 'w:tabs');
  if (tabsEl) {
    const parsedStops: DocTabStop[] = [];
    for (const tabEl of childEls(tabsEl, 'w:tab')) {
      const rawVal = (attr(tabEl, 'val') ?? 'left').toLowerCase();
      const rawPos = attr(tabEl, 'pos');
      if (rawPos === undefined) continue;
      const posTwips = Math.round(num(rawPos, NaN));
      if (!Number.isFinite(posTwips)) continue;
      const val: DocTabStop['val'] =
        rawVal === 'clear'
          ? 'clear'
          : rawVal === 'center'
            ? 'center'
            : rawVal === 'right' || rawVal === 'end'
              ? 'right'
              : rawVal === 'decimal' || rawVal === 'num'
                ? 'decimal'
                : rawVal === 'bar'
                  ? 'bar'
                  : 'left';
      const rawLeader = (attr(tabEl, 'leader') ?? 'none').toLowerCase();
      const leader: DocTabStop['leader'] =
        rawLeader === 'dot'
          ? 'dot'
          : rawLeader === 'hyphen'
            ? 'hyphen'
            : rawLeader === 'underscore'
              ? 'underscore'
              : rawLeader === 'heavy'
                ? 'heavy'
                : rawLeader === 'middledot'
                  ? 'middleDot'
                  : 'none';
      parsedStops.push({
        posTwips,
        posPx: Math.max(0, twipsToPx(posTwips)),
        val,
        leader,
      });
    }
    if (parsedStops.length > 0) {
      out.tabStops = parsedStops;
    }
  }

  return out;
}

function parseCellMarginsEl(
  el: XmlNode | undefined,
): {top: number; right: number; bottom: number; left: number} | undefined {
  if (!el) return undefined;
  const topEl = firstChild(el, 'w:top');
  const rightEl = firstChild(el, 'w:right') ?? firstChild(el, 'w:end');
  const bottomEl = firstChild(el, 'w:bottom');
  const leftEl = firstChild(el, 'w:left') ?? firstChild(el, 'w:start');
  if (!topEl && !rightEl && !bottomEl && !leftEl) return undefined;
  return {
    top: topEl ? Math.max(1, twipsToPx(num(attr(topEl, 'w'), 60))) : 4,
    right: rightEl ? Math.max(2, twipsToPx(num(attr(rightEl, 'w'), 108))) : 7,
    bottom: bottomEl ? Math.max(1, twipsToPx(num(attr(bottomEl, 'w'), 60))) : 4,
    left: leftEl ? Math.max(2, twipsToPx(num(attr(leftEl, 'w'), 108))) : 7,
  };
}

interface StyleSheet {
  docDefaultsRPr: ResolvedRunProps;
  docDefaultsPPr: ResolvedParaProps;
  normalStyleId: string | null;
  styles: Map<string, StyleDefinition>;
}

function parseStylesXml(stylesXml: string | null, theme: ThemeContext): StyleSheet {
  const sheet: StyleSheet = {
    docDefaultsRPr: {},
    docDefaultsPPr: {},
    normalStyleId: null,
    styles: new Map(),
  };
  if (!stylesXml) return sheet;
  const root = parseXml(stylesXml);
  if (!root) return sheet;

  const docDefaults = firstChild(root, 'w:docDefaults');
  if (docDefaults) {
    const rPrDef = firstChild(firstChild(docDefaults, 'w:rPrDefault'), 'w:rPr');
    const pPrDef = firstChild(firstChild(docDefaults, 'w:pPrDefault'), 'w:pPr');
    sheet.docDefaultsRPr = parseRunPropsEl(rPrDef, theme);
    sheet.docDefaultsPPr = parseParaPropsEl(pPrDef);
  }

  for (const styleEl of childEls(root, 'w:style')) {
    const id = attr(styleEl, 'styleId');
    if (!id) continue;
    const typeRaw = (attr(styleEl, 'type') ?? 'paragraph') as StyleDefinition['type'];
    const nameEl = firstChild(styleEl, 'w:name');
    const basedOnEl = firstChild(styleEl, 'w:basedOn');
    const isDefault = attr(styleEl, 'default') === '1' || attr(styleEl, 'default') === 'true';
    const pPr = parseParaPropsEl(firstChild(styleEl, 'w:pPr'));
    const rPr = parseRunPropsEl(firstChild(styleEl, 'w:rPr'), theme);
    let tblPr: ResolvedTableProps | undefined;
    const tblPrEl = firstChild(styleEl, 'w:tblPr');
    if (tblPrEl) {
      tblPr = {
        borders: parseBorderSet(firstChild(tblPrEl, 'w:tblBorders'), theme.colors),
        cellMarginsPx: parseCellMarginsEl(firstChild(tblPrEl, 'w:tblCellMar')),
      };
    }
    const def: StyleDefinition = {
      id,
      type: typeRaw,
      name: attr(nameEl, 'val') ?? id,
      basedOn: attr(basedOnEl, 'val'),
      isDefault,
      pPr,
      rPr,
      tblPr,
    };
    sheet.styles.set(id, def);
    if (typeRaw === 'paragraph' && (id === 'Normal' || (isDefault && !sheet.normalStyleId))) {
      sheet.normalStyleId = id;
    }
  }

  return sheet;
}

function mergeTabStops(
  base: DocTabStop[] | undefined,
  override: DocTabStop[] | undefined,
): DocTabStop[] | undefined {
  if (!base?.length && !override?.length) return undefined;
  const byPos = new Map<number, DocTabStop>();
  for (const s of base ?? []) {
    if (s.val !== 'clear') byPos.set(s.posTwips, s);
  }
  for (const s of override ?? []) {
    if (s.val === 'clear') {
      byPos.delete(s.posTwips);
    } else {
      byPos.set(s.posTwips, s);
    }
  }
  if (byPos.size === 0) return undefined;
  return Array.from(byPos.values()).sort((a, b) => a.posTwips - b.posTwips);
}

function mergeDefined<T extends object>(base: T, override: Partial<T>): T {
  const out = {...base};
  for (const [k, v] of Object.entries(override)) {
    if (v !== undefined) {
      (out as Record<string, unknown>)[k] = v;
    }
  }
  return out;
}

function mergeParaProps(base: ResolvedParaProps, override: ResolvedParaProps): ResolvedParaProps {
  const merged = mergeDefined(base, override);
  const tabStops = mergeTabStops(base.tabStops, override.tabStops);
  if (tabStops) merged.tabStops = tabStops;
  else delete merged.tabStops;
  return merged;
}

function resolveStyleChain(
  sheet: StyleSheet,
  styleId: string | undefined,
): {pPr: ResolvedParaProps; rPr: ResolvedRunProps; tblPr?: ResolvedTableProps; name?: string} {
  const chain: StyleDefinition[] = [];
  const visited = new Set<string>();
  let currId = styleId ?? sheet.normalStyleId ?? undefined;
  while (currId && !visited.has(currId)) {
    visited.add(currId);
    const def = sheet.styles.get(currId);
    if (!def) break;
    chain.unshift(def);
    currId = def.basedOn;
  }

  let pPr: ResolvedParaProps = {...sheet.docDefaultsPPr};
  let rPr: ResolvedRunProps = {...sheet.docDefaultsRPr};
  let tblPr: ResolvedTableProps | undefined;
  for (const def of chain) {
    pPr = mergeParaProps(pPr, def.pPr);
    rPr = mergeDefined(rPr, def.rPr);
    if (def.tblPr) {
      tblPr = {
        borders: {...(tblPr?.borders ?? {}), ...(def.tblPr.borders ?? {})},
        cellMarginsPx: def.tblPr.cellMarginsPx ?? tblPr?.cellMarginsPx,
      };
    }
  }
  const leaf = styleId ? sheet.styles.get(styleId) : undefined;
  return {pPr, rPr, tblPr, name: leaf?.name};
}

interface NumberingLevelSpec {
  ilvl: number;
  numFmt: string;
  isOrdered: boolean;
  indentLeftPx?: number;
}

type NumberingMap = Map<string, Map<number, NumberingLevelSpec>>;

function parseNumberingXml(numberingXml: string | null): NumberingMap {
  const result: NumberingMap = new Map();
  if (!numberingXml) return result;
  const root = parseXml(numberingXml);
  if (!root) return result;

  const abstractMap = new Map<string, Map<number, NumberingLevelSpec>>();
  for (const abs of childEls(root, 'w:abstractNum')) {
    const absId = attr(abs, 'abstractNumId');
    if (!absId) continue;
    const levels = new Map<number, NumberingLevelSpec>();
    for (const lvl of childEls(abs, 'w:lvl')) {
      const ilvl = num(attr(lvl, 'ilvl'), 0);
      const numFmt = (attr(firstChild(lvl, 'w:numFmt'), 'val') ?? 'bullet').toLowerCase();
      const isOrdered = numFmt !== 'bullet' && numFmt !== 'none';
      const indEl = firstChild(firstChild(lvl, 'w:pPr'), 'w:ind');
      const leftRaw = attr(indEl, 'left') ?? attr(indEl, 'start');
      levels.set(ilvl, {
        ilvl,
        numFmt,
        isOrdered,
        indentLeftPx: leftRaw !== undefined ? twipsToPx(num(leftRaw, 0)) : undefined,
      });
    }
    abstractMap.set(absId, levels);
  }

  for (const numEl of childEls(root, 'w:num')) {
    const numId = attr(numEl, 'numId');
    const absId = attr(firstChild(numEl, 'w:abstractNumId'), 'val');
    if (!numId || !absId) continue;
    const absLevels = abstractMap.get(absId);
    if (!absLevels) continue;
    const copy = new Map(absLevels);
    for (const ov of childEls(numEl, 'w:lvlOverride')) {
      const lvl = firstChild(ov, 'w:lvl');
      if (lvl) {
        const ilvl = num(attr(lvl, 'ilvl'), num(attr(ov, 'ilvl'), 0));
        const numFmt = (attr(firstChild(lvl, 'w:numFmt'), 'val') ?? 'bullet').toLowerCase();
        copy.set(ilvl, {ilvl, numFmt, isOrdered: numFmt !== 'bullet' && numFmt !== 'none'});
      }
    }
    result.set(numId, copy);
  }

  return result;
}

// ---------------------------------------------------------------------------
// Relationships & Media
// ---------------------------------------------------------------------------

interface RelEntry {
  id: string;
  target: string;
  type: string;
  isExternal: boolean;
}

function parseRelsXml(relsXml: string | null): Map<string, RelEntry> {
  const map = new Map<string, RelEntry>();
  if (!relsXml) return map;
  const root = parseXml(relsXml);
  if (!root) return map;
  for (const rel of childEls(root, 'Relationship')) {
    const id = attr(rel, 'Id');
    const target = attr(rel, 'Target');
    const type = attr(rel, 'Type') ?? '';
    const isExternal = (attr(rel, 'TargetMode') ?? '').toLowerCase() === 'external';
    if (!id || !target) continue;
    const normalized = isExternal
      ? target
      : target.startsWith('/')
        ? target.slice(1)
        : target.startsWith('word/')
          ? target
          : `word/${target.replace(/^\.\//, '')}`;
    map.set(id, {id, target: normalized, type, isExternal});
  }
  return map;
}

function guessMediaMime(filePath: string): string {
  const lower = filePath.toLowerCase();
  if (lower.endsWith('.png')) return 'image/png';
  if (lower.endsWith('.jpg') || lower.endsWith('.jpeg')) return 'image/jpeg';
  if (lower.endsWith('.gif')) return 'image/gif';
  if (lower.endsWith('.webp')) return 'image/webp';
  if (lower.endsWith('.svg')) return 'image/svg+xml';
  if (lower.endsWith('.bmp')) return 'image/bmp';
  return 'application/octet-stream';
}

// ---------------------------------------------------------------------------
// Section Properties, Page Borders, Headers & Footers
// ---------------------------------------------------------------------------

function closestPageSize(widthPx: number, heightPx: number): DocPageSize {
  const minDim = Math.min(widthPx, heightPx);
  const maxDim = Math.max(widthPx, heightPx);
  // A4: 794 x 1123, Letter: 816 x 1056, Legal: 816 x 1344
  if (maxDim > 1230) return 'legal';
  if (Math.abs(minDim - 816) < Math.abs(minDim - 794) && maxDim < 1090) return 'letter';
  return 'a4';
}

function closestMarginPreset(marginsPx: {
  top: number;
  right: number;
  bottom: number;
  left: number;
}): DocMarginPreset {
  const avgH = (marginsPx.left + marginsPx.right) / 2;
  if (avgH <= 46) return 'narrow';
  if (avgH <= 64) return 'moderate';
  if (avgH >= 96) return 'wide';
  return 'normal';
}

function parseHeaderFooterPart(xml: string | null): {
  text: string;
  hasPageNumber: boolean;
  align: 'left' | 'center' | 'right';
} {
  if (!xml) return {text: '', hasPageNumber: false, align: 'center'};
  const root = parseXml(xml);
  if (!root) return {text: '', hasPageNumber: false, align: 'center'};

  let hasPageNumber = false;
  let align: 'left' | 'center' | 'right' = 'center';
  const lines: string[] = [];

  for (const p of findDescendants(root, 'w:p')) {
    const jc = (attr(firstChild(firstChild(p, 'w:pPr'), 'w:jc'), 'val') ?? '').toLowerCase();
    if (jc === 'right' || jc === 'end') align = 'right';
    else if (jc === 'left' || jc === 'start') align = 'left';
    else if (jc === 'center') align = 'center';

    // Detect PAGE / NUMPAGES fields
    for (const fld of findDescendants(p, 'w:fldSimple')) {
      const instr = (attr(fld, 'instr') ?? '').toUpperCase();
      if (/\bPAGE\b|\bNUMPAGES\b/.test(instr)) hasPageNumber = true;
    }
    for (const instr of findDescendants(p, 'w:instrText')) {
      if (/\bPAGE\b|\bNUMPAGES\b/.test(textContent(instr).toUpperCase())) {
        hasPageNumber = true;
      }
    }

    // Collect non-field text (skip Page X of Y boilerplate if dynamic fields are present)
    let inField = false;
    let fieldIsPage = false;
    let lineText = '';
    for (const child of childEls(p)) {
      if (child.tag === 'w:fldSimple') {
        const instr = (attr(child, 'instr') ?? '').toUpperCase();
        if (!/\bPAGE\b|\bNUMPAGES\b/.test(instr)) {
          lineText += textContent(child);
        }
        continue;
      }
      if (child.tag !== 'w:r') continue;
      const fldChar = firstChild(child, 'w:fldChar');
      if (fldChar) {
        const t = (attr(fldChar, 'fldCharType') ?? '').toLowerCase();
        if (t === 'begin') {
          inField = true;
          fieldIsPage = false;
        } else if (t === 'end') {
          inField = false;
          fieldIsPage = false;
        }
        continue;
      }
      const instrEl = firstChild(child, 'w:instrText');
      if (instrEl) {
        if (/\bPAGE\b|\bNUMPAGES\b/.test(textContent(instrEl).toUpperCase())) {
          fieldIsPage = true;
        }
        continue;
      }
      if (inField && fieldIsPage) continue;
      for (const rChild of childEls(child)) {
        if (rChild.tag === 'w:t') lineText += textContent(rChild);
        else if (rChild.tag === 'w:tab') lineText += ' ';
      }
    }

    const cleaned = lineText
      .replace(/\bPage\s+of\b/gi, '')
      .replace(/\s+/g, ' ')
      .trim();
    if (cleaned) lines.push(cleaned);
  }

  return {
    text: lines.join(' · ').trim(),
    hasPageNumber,
    align,
  };
}

// ---------------------------------------------------------------------------
// DrawingML / Shapes / Images / Unsupported Elements
// ---------------------------------------------------------------------------

interface SectionGeometry {
  pageWidthPx: number;
  pageHeightPx: number;
  marginLeftPx: number;
  marginRightPx: number;
  marginTopPx: number;
  marginBottomPx: number;
}

interface ImportContext {
  theme: ThemeContext;
  styles: StyleSheet;
  numbering: NumberingMap;
  rels: Map<string, RelEntry>;
  mediaDataUrl: Map<string, string>;
  defaultFontFamily: string | null;
  defaultFontSizePt: number | null;
  defaultColor: string | null;
  defaultTabStopPx: number;
  sectionGeometry: SectionGeometry;
  warningsMap: Map<string, DocxImportWarning>;
  hasLockedElements: boolean;
  hasLossyElements: boolean;
  stats: DocxImportResult['stats'];
}

function addWarning(
  ctx: ImportContext,
  code: string,
  feature: string,
  message: string,
  severity: 'info' | 'warning' = 'warning',
  markLossy = true,
) {
  if (markLossy) ctx.hasLossyElements = true;
  const existing = ctx.warningsMap.get(code);
  if (existing) {
    existing.count += 1;
  } else {
    ctx.warningsMap.set(code, {code, feature, message, severity, count: 1});
  }
}

function extractDrawingFillColor(node: XmlNode | undefined, theme: ThemeContext): string | null {
  if (!node) return null;
  if (firstChild(node, 'a:noFill')) return 'transparent';
  const solidFill = firstChild(node, 'a:solidFill');
  if (!solidFill) return null;
  const srgb = firstChild(solidFill, 'a:srgbClr');
  if (srgb) {
    return normalizeColor(attr(srgb, 'val'), theme.colors) ?? null;
  }
  const scheme = firstChild(solidFill, 'a:schemeClr');
  if (scheme) {
    const val = (attr(scheme, 'val') ?? '').toLowerCase();
    return theme.colors[val] ?? '#000000';
  }
  return null;
}

function renderLockedElementHtml(
  label: string,
  detail: string,
  imageDataUrl?: string,
  widthPx?: number,
  heightPx?: number,
): string {
  const attrs: string[] = [
    'data-doc-locked="true"',
    `data-locked-label="${escapeHtmlAttr(label)}"`,
    `data-locked-detail="${escapeHtmlAttr(detail)}"`,
  ];
  if (imageDataUrl) attrs.push(`data-locked-image="${escapeHtmlAttr(imageDataUrl)}"`);
  if (widthPx) attrs.push(`data-width="${Math.round(widthPx)}"`);
  if (heightPx) attrs.push(`data-height="${Math.round(heightPx)}"`);

  if (imageDataUrl) {
    const styleParts = ['display:block', 'margin:8px auto', 'max-width:100%'];
    if (widthPx) styleParts.push(`width:${Math.round(widthPx)}px`);
    return `<div ${attrs.join(' ')}><img src="${escapeHtmlAttr(imageDataUrl)}" alt="${escapeHtmlAttr(label)}" style="${styleParts.join(';')}" /></div>`;
  }
  return `<div ${attrs.join(' ')}>${escapeHtmlText(label)}</div>`;
}

interface ImageRenderOptions {
  wrapType?: 'topAndBottom' | 'none' | 'square' | 'tight';
  floatLeftPx?: number | null;
  floatTopPx?: number | null;
  relativeFromH?: string | null;
  relativeFromV?: string | null;
  marginTopPx?: number | null;
  marginBottomPx?: number | null;
  ooxmlDrawingXml?: string | null;
}

function renderImageHtml(
  dataUrl: string,
  widthPx: number,
  heightPx: number,
  align: 'left' | 'center' | 'right',
  floating: boolean,
  alt = 'Embedded image',
  opts?: ImageRenderOptions,
): string {
  const safeW = Math.max(16, Math.round(widthPx));
  const safeH = Math.max(16, Math.round(heightPx));
  const styleParts = [
    `width:${safeW}px`,
    `height:${safeH}px`,
    'max-width:100%',
    'object-fit:contain',
  ];

  const isSideFloat =
    floating && (opts?.wrapType === 'square' || opts?.wrapType === 'tight');
  if (isSideFloat && align === 'right') {
    styleParts.push('float:right', 'margin-left:12px');
  } else if (isSideFloat && align === 'left') {
    styleParts.push('float:left', 'margin-right:12px');
  } else if (align === 'center') {
    styleParts.push('display:block', 'margin-left:auto', 'margin-right:auto');
  } else if (align === 'right') {
    styleParts.push('display:block', 'margin-left:auto', 'margin-right:0');
  } else if (floating && opts?.floatLeftPx && opts.floatLeftPx > 0) {
    styleParts.push('display:block', `margin-left:${Math.round(opts.floatLeftPx)}px`);
  } else {
    styleParts.push('display:inline-block');
  }

  if (opts?.marginTopPx !== null && opts?.marginTopPx !== undefined) {
    styleParts.push(`margin-top:${Math.round(opts.marginTopPx)}px`);
  }
  if (opts?.marginBottomPx !== null && opts?.marginBottomPx !== undefined) {
    styleParts.push(`margin-bottom:${Math.round(opts.marginBottomPx)}px`);
  }

  const extraAttrs: string[] = [
    `data-width="${safeW}"`,
    `data-height="${safeH}"`,
    `data-align="${align}"`,
  ];
  if (floating) extraAttrs.push('data-floating="true"');
  if (opts?.wrapType) extraAttrs.push(`data-wrap="${escapeHtmlAttr(opts.wrapType)}"`);
  if (opts?.floatLeftPx !== null && opts?.floatLeftPx !== undefined) {
    extraAttrs.push(`data-float-left="${Math.round(opts.floatLeftPx)}"`);
  }
  if (opts?.floatTopPx !== null && opts?.floatTopPx !== undefined) {
    extraAttrs.push(`data-float-top="${Math.round(opts.floatTopPx)}"`);
  }
  if (opts?.relativeFromH) extraAttrs.push(`data-rel-h="${escapeHtmlAttr(opts.relativeFromH)}"`);
  if (opts?.relativeFromV) extraAttrs.push(`data-rel-v="${escapeHtmlAttr(opts.relativeFromV)}"`);
  if (opts?.marginTopPx !== null && opts?.marginTopPx !== undefined) {
    extraAttrs.push(`data-margin-top="${Math.round(opts.marginTopPx)}"`);
  }
  if (opts?.marginBottomPx !== null && opts?.marginBottomPx !== undefined) {
    extraAttrs.push(`data-margin-bottom="${Math.round(opts.marginBottomPx)}"`);
  }
  if (opts?.ooxmlDrawingXml) {
    extraAttrs.push(`data-ooxml-drawing="${encodeUtf8Base64(opts.ooxmlDrawingXml)}"`);
  }

  return `<img src="${escapeHtmlAttr(dataUrl)}" alt="${escapeHtmlAttr(alt)}" ${extraAttrs.join(' ')} style="${styleParts.join(';')}" />`;
}

function renderRunHtml(
  text: string,
  rPr: ResolvedRunProps,
  inheritedDefaults: {fontFamily: string | null; fontSizePt: number | null; color: string | null},
  forceExplicitHeading = false,
): string {
  if (!text) return '';
  let content = escapeHtmlText(text);

  // Innermost text-style span (font, size, color, caps, letterSpacing)
  const styles: string[] = [];
  const attrs: string[] = [];

  const canonicalFont = rPr.fontFamily
    ? canonicalizeWordFontFamily(rPr.fontFamily) ?? rPr.fontFamily
    : null;
  if (canonicalFont && canonicalFont !== inheritedDefaults.fontFamily) {
    attrs.push(`data-doc-font="${escapeHtmlAttr(canonicalFont)}"`);
    styles.push(`font-family: ${resolveRenderFontStack(canonicalFont)}`);
  }

  const effectiveSizePt = rPr.fontSizePt ?? (forceExplicitHeading ? inheritedDefaults.fontSizePt : undefined);
  if (effectiveSizePt && (forceExplicitHeading || effectiveSizePt !== inheritedDefaults.fontSizePt)) {
    styles.push(`font-size: ${effectiveSizePt}pt`);
  }

  const effectiveColor = rPr.color ?? (forceExplicitHeading ? inheritedDefaults.color ?? '#000000' : undefined);
  if (effectiveColor && (forceExplicitHeading || effectiveColor !== inheritedDefaults.color)) {
    styles.push(`color: ${effectiveColor}`);
  }

  if (rPr.allCaps) {
    styles.push('text-transform: uppercase');
  }
  if (rPr.smallCaps) {
    styles.push('font-variant: small-caps');
  }
  if (rPr.letterSpacingPt) {
    styles.push(`letter-spacing: ${rPr.letterSpacingPt}pt`);
  }

  if (styles.length > 0 || attrs.length > 0) {
    const attrPrefix = attrs.length > 0 ? ` ${attrs.join(' ')}` : '';
    const styleAttr = styles.length > 0 ? ` style="${escapeHtmlAttr(styles.join('; '))}"` : '';
    content = `<span${attrPrefix}${styleAttr}>${content}</span>`;
  }

  if (rPr.highlight) {
    content = `<mark data-color="${escapeHtmlAttr(rPr.highlight)}" style="background-color: ${escapeHtmlAttr(rPr.highlight)}">${content}</mark>`;
  }
  if (rPr.bold) content = `<strong>${content}</strong>`;
  if (rPr.italic) content = `<em>${content}</em>`;
  if (rPr.underline) content = `<u>${content}</u>`;
  if (rPr.strike) content = `<s>${content}</s>`;
  if (rPr.vertAlign === 'subscript') content = `<sub>${content}</sub>`;
  else if (rPr.vertAlign === 'superscript') content = `<sup>${content}</sup>`;

  return content;
}

function renderTabHtml(stop: DocTabStop | undefined, defaultTabStopPx: number): string {
  const align = stop && stop.val !== 'clear' ? stop.val : 'left';
  const leader = stop?.leader ?? 'none';
  const widthPx = Math.max(12, Math.round(defaultTabStopPx || 48));
  const posAttr = stop ? ` data-tab-pos="${stop.posTwips}"` : '';
  return `<span data-doc-tab="true" data-tab-align="${align}" data-tab-leader="${leader}" data-tab-width="${widthPx}"${posAttr} class="esz-doc-tab" style="width: ${widthPx}px">\t</span>`;
}

function parseDrawingElement(
  drawingEl: XmlNode,
  paraAlign: 'left' | 'center' | 'right' | 'justify' | undefined,
  ctx: ImportContext,
): {inlineHtml: string; beforeBlockHtmls: string[]; blockHtmls: string[]} {
  const inlineHtmlParts: string[] = [];
  const beforeBlockHtmls: string[] = [];
  const blockHtmls: string[] = [];

  // Check for SmartArt or Chart inside graphicData
  if (findDescendants(drawingEl, 'dgm:relIds').length > 0) {
    ctx.hasLockedElements = true;
    ctx.stats.lockedElements += 1;
    addWarning(
      ctx,
      'smartart',
      'SmartArt Graphic',
      'SmartArt diagrams are not editable in Edit mode and are shown as locked placeholders.',
    );
    blockHtmls.push(
      renderLockedElementHtml(
        'SmartArt Graphic (Locked)',
        'SmartArt diagrams cannot be edited directly. Switch to View mode or Save as Copy to preserve the original.',
      ),
    );
    return {inlineHtml: '', beforeBlockHtmls, blockHtmls};
  }

  if (findDescendants(drawingEl, 'c:chart').length > 0) {
    ctx.hasLockedElements = true;
    ctx.stats.lockedElements += 1;
    addWarning(
      ctx,
      'chart',
      'Embedded Chart',
      'Word charts cannot be edited directly and are shown as locked placeholders.',
    );
    blockHtmls.push(
      renderLockedElementHtml(
        'Embedded Word Chart (Locked)',
        'Charts cannot be edited directly in Docs Edit mode.',
      ),
    );
    return {inlineHtml: '', beforeBlockHtmls, blockHtmls};
  }

  const isAnchor = Boolean(firstChild(drawingEl, 'wp:anchor'));
  const container = firstChild(drawingEl, 'wp:anchor') ?? firstChild(drawingEl, 'wp:inline') ?? drawingEl;

  // Extent dimensions
  const extentEl = firstChild(container, 'wp:extent') ?? findDescendants(container, 'a:ext')[0];
  const widthPx = Math.max(24, Math.round(num(attr(extentEl, 'cx'), 240 * EMU_PER_PX) / EMU_PER_PX));
  const heightPx = Math.max(20, Math.round(num(attr(extentEl, 'cy'), 120 * EMU_PER_PX) / EMU_PER_PX));

  // Horizontal alignment from wp:positionH/wp:align or enclosing paragraph alignment
  let align: 'left' | 'center' | 'right' =
    paraAlign === 'center' ? 'center' : paraAlign === 'right' ? 'right' : 'left';
  const posH = firstChild(container, 'wp:positionH');
  let posHAlignTxt = '';
  if (posH) {
    const alignEl = firstChild(posH, 'wp:align');
    posHAlignTxt = textContent(alignEl).trim().toLowerCase();
    if (posHAlignTxt === 'center' || posHAlignTxt === 'left' || posHAlignTxt === 'right') {
      align = posHAlignTxt;
    }
  }

  // Case 1: DrawingML Shape / TextBox / Rounded Rectangle (<wps:wsp> or <wpg:wgp>)
  const wsp = findDescendants(container, 'wps:wsp')[0];
  const wgp = findDescendants(container, 'wpg:wgp')[0];
  if (wsp || wgp) {
    const has3dOrWarp =
      findDescendants(container, 'a:prstTxWarp').length > 0 ||
      findDescendants(container, 'a:scene3d').length > 0 ||
      findDescendants(container, 'a:sp3d').length > 0;
    const txbxContents = findDescendants(container, 'w:txbxContent');

    // Check optional blip image inside shape
    const blipEl = findDescendants(container, 'a:blip')[0];
    const blipEmbedId = blipEl ? attr(blipEl, 'embed') : undefined;
    const blipRel = blipEmbedId ? ctx.rels.get(blipEmbedId) : undefined;
    const blipDataUrl = blipRel ? ctx.mediaDataUrl.get(blipRel.target) : undefined;

    // Complex 3D WordArt or grouped non-text drawing -> locked visual fallback
    if (has3dOrWarp && txbxContents.length === 0) {
      ctx.hasLockedElements = true;
      ctx.stats.lockedElements += 1;
      addWarning(
        ctx,
        'wordart-3d',
        '3D WordArt / Advanced Drawing',
        '3D WordArt warp/bevel effects cannot be edited directly and are shown as locked elements.',
      );
      blockHtmls.push(
        renderLockedElementHtml(
          'WordArt / Advanced Drawing (Locked)',
          'Advanced 3D WordArt effects are preserved in View mode.',
          blipDataUrl,
          widthPx,
          heightPx,
        ),
      );
      return {inlineHtml: '', beforeBlockHtmls, blockHtmls};
    }

    if (wgp && txbxContents.length === 0) {
      ctx.hasLockedElements = true;
      ctx.stats.lockedElements += 1;
      addWarning(
        ctx,
        'grouped-shape',
        'Grouped Drawing Shape',
        'Grouped drawings without editable text are shown as locked elements.',
      );
      blockHtmls.push(
        renderLockedElementHtml(
          'Grouped Drawing (Locked)',
          'Grouped drawing objects cannot be edited directly.',
          blipDataUrl,
          widthPx,
          heightPx,
        ),
      );
      return {inlineHtml: '', beforeBlockHtmls, blockHtmls};
    }

    // Simple DrawingML shape / text box / rounded rectangle -> editable styled docTextBox!
    if (has3dOrWarp) {
      addWarning(
        ctx,
        'wordart-effects',
        'WordArt Text Effects',
        'WordArt text was imported as an editable styled box; advanced 3D warp/shadow effects are simplified in Edit mode.',
        'info',
      );
    } else {
      addWarning(
        ctx,
        'drawingml-textbox',
        'DrawingML Text Box / Shape',
        'Word text boxes and rounded shapes were converted to editable styled boxes. Choose "Save as Copy" to keep the original DrawingML layout untouched.',
        'info',
      );
    }

    const spPr = firstChild(wsp ?? container, 'wps:spPr') ?? findDescendants(container, 'wps:spPr')[0];
    const prstGeomEl = firstChild(spPr, 'a:prstGeom');
    const prstGeom = attr(prstGeomEl, 'prst') ?? 'rect';

    let borderRadiusPx = 0;
    if (/round/i.test(prstGeom)) {
      const adjGd = findDescendants(prstGeomEl, 'a:gd').find((g) => attr(g, 'name') === 'adj');
      const fmla = attr(adjGd, 'fmla') ?? '';
      const adjMatch = fmla.match(/val\s+(\d+)/i);
      const adjVal = adjMatch ? num(adjMatch[1], 16667) : 16667;
      const minDim = Math.min(widthPx, heightPx);
      borderRadiusPx = Math.max(
        4,
        Math.min(Math.round(minDim / 2), Math.round((minDim * adjVal) / 100000)),
      );
    } else if (/ellipse|oval/i.test(prstGeom)) {
      borderRadiusPx = 9999;
    }

    const fillColor = extractDrawingFillColor(spPr, ctx.theme) ?? 'transparent';

    const lnEl = firstChild(spPr, 'a:ln');
    let borderColor: string | null = null;
    let borderWidthPx = 0;
    let borderStyle: 'solid' | 'dashed' | 'dotted' | 'none' = 'none';
    if (lnEl) {
      const strokeColor = extractDrawingFillColor(lnEl, ctx.theme);
      if (strokeColor && strokeColor !== 'transparent') {
        borderColor = strokeColor;
        borderWidthPx = Math.max(1, Number((num(attr(lnEl, 'w'), EMU_PER_PX) / EMU_PER_PX).toFixed(2)));
        const dashVal = attr(firstChild(lnEl, 'a:prstDash'), 'val') ?? 'solid';
        borderStyle = /dot/i.test(dashVal)
          ? 'dotted'
          : /dash/i.test(dashVal)
            ? 'dashed'
            : 'solid';
      }
    } else if (/roundRect/i.test(prstGeom)) {
      const styleEl = firstChild(wsp ?? container, 'wps:style');
      const lnRef = firstChild(styleEl, 'a:lnRef');
      if (lnRef && attr(lnRef, 'idx') && attr(lnRef, 'idx') !== '0') {
        borderColor = '#000000';
        borderWidthPx = 1;
        borderStyle = 'solid';
      }
    }

    const bodyPr = firstChild(wsp ?? container, 'wps:bodyPr');
    const padL = Math.round(num(attr(bodyPr, 'lIns'), 91440) / EMU_PER_PX);
    const padR = Math.round(num(attr(bodyPr, 'rIns'), 91440) / EMU_PER_PX);
    const padT = Math.round(num(attr(bodyPr, 'tIns'), 45720) / EMU_PER_PX);
    const padB = Math.round(num(attr(bodyPr, 'bIns'), 45720) / EMU_PER_PX);
    const anchorV = attr(bodyPr, 'anchor');
    const vAlign = anchorV === 'ctr' ? 'middle' : anchorV === 'b' ? 'bottom' : 'top';

    const innerBlocks: string[] = [];
    for (const tb of txbxContents) {
      for (const child of childEls(tb)) {
        if (child.tag === 'w:p') {
          innerBlocks.push(...convertParagraphNode(child, ctx));
        } else if (child.tag === 'w:tbl') {
          innerBlocks.push(convertTableNode(child, ctx));
        }
      }
    }

    if (innerBlocks.length === 0) {
      if (!borderColor && fillColor === 'transparent' && !blipDataUrl) {
        return {inlineHtml: '', beforeBlockHtmls, blockHtmls: []};
      }
      innerBlocks.push('<p></p>');
    }

    const boxAlign =
      posH || paraAlign
        ? align
        : /roundRect/i.test(prstGeom) || has3dOrWarp
          ? 'center'
          : 'left';

    const boxAttrs: string[] = [
      'data-doc-textbox="true"',
      `data-box-width="${widthPx}"`,
      `data-box-min-height="${heightPx}"`,
      `data-box-align="${boxAlign}"`,
      `data-box-radius="${borderRadiusPx}"`,
      `data-box-fill="${escapeHtmlAttr(fillColor)}"`,
      `data-box-border-style="${borderStyle}"`,
      `data-box-border-width="${borderWidthPx}"`,
      `data-box-valign="${vAlign}"`,
      `data-box-padding="${padT}px ${padR}px ${padB}px ${padL}px"`,
    ];
    if (borderColor) boxAttrs.push(`data-box-border-color="${escapeHtmlAttr(borderColor)}"`);
    if (has3dOrWarp) boxAttrs.push('data-box-wordart="true"');

    ctx.stats.editableShapes += 1;
    blockHtmls.push(`<div ${boxAttrs.join(' ')}>${innerBlocks.join('')}</div>`);
    return {inlineHtml: '', beforeBlockHtmls, blockHtmls};
  }

  // Case 2: Standard DrawingML Picture (<pic:pic> / <a:blip>)
  const blipEl = findDescendants(container, 'a:blip')[0];
  if (blipEl) {
    const embedId = attr(blipEl, 'embed') ?? attr(blipEl, 'link');
    const rel = embedId ? ctx.rels.get(embedId) : undefined;
    const dataUrl = rel ? ctx.mediaDataUrl.get(rel.target) : undefined;
    if (dataUrl) {
      const docPr = firstChild(container, 'wp:docPr');
      const alt = attr(docPr, 'descr') || attr(docPr, 'name') || 'Embedded image';
      ctx.stats.images += 1;

      if (isAnchor) {
        const wrapType: ImageRenderOptions['wrapType'] = firstChild(container, 'wp:wrapTopAndBottom')
          ? 'topAndBottom'
          : firstChild(container, 'wp:wrapNone')
            ? 'none'
            : firstChild(container, 'wp:wrapSquare')
              ? 'square'
              : firstChild(container, 'wp:wrapTight') || firstChild(container, 'wp:wrapThrough')
                ? 'tight'
                : 'topAndBottom';

        const geom = ctx.sectionGeometry;
        const relativeFromH = (attr(posH, 'relativeFrom') ?? 'column').toLowerCase();
        const posHOffsetEl = firstChild(posH, 'wp:posOffset');
        const hasPosHOffset = Boolean(posHOffsetEl);
        const posHOffsetPx = hasPosHOffset
          ? Math.round(num(textContent(posHOffsetEl), 0) / EMU_PER_PX)
          : 0;

        let pageLeftPx: number;
        if (posHAlignTxt === 'center') {
          pageLeftPx = (geom.pageWidthPx - widthPx) / 2;
        } else if (posHAlignTxt === 'right') {
          pageLeftPx =
            relativeFromH === 'page'
              ? geom.pageWidthPx - widthPx
              : geom.pageWidthPx - geom.marginRightPx - widthPx;
        } else if (relativeFromH === 'page') {
          pageLeftPx = posHOffsetPx;
        } else {
          pageLeftPx = geom.marginLeftPx + posHOffsetPx;
        }

        const centerOffsetRatio =
          Math.abs(pageLeftPx + widthPx / 2 - geom.pageWidthPx / 2) / Math.max(1, geom.pageWidthPx);
        const isCenteredAnchor =
          posHAlignTxt === 'center' || (hasPosHOffset && centerOffsetRatio <= 0.05);

        const posV = firstChild(container, 'wp:positionV');
        const relativeFromV = (attr(posV, 'relativeFrom') ?? 'paragraph').toLowerCase();
        const posVOffsetEl = firstChild(posV, 'wp:posOffset');
        const posVOffsetPx = posVOffsetEl
          ? Math.round(num(textContent(posVOffsetEl), 0) / EMU_PER_PX)
          : null;

        const distTPx = Math.max(0, Math.round(num(attr(container, 'distT'), 0) / EMU_PER_PX));
        const distBPx = Math.max(0, Math.round(num(attr(container, 'distB'), 0) / EMU_PER_PX));

        let placeBefore = false;
        let marginTopPx = distTPx;
        let marginBottomPx = distBPx;
        if (
          (relativeFromV === 'paragraph' || relativeFromV === 'line') &&
          posVOffsetPx !== null
        ) {
          if (posVOffsetPx < -heightPx * 0.5) {
            placeBefore = true;
            marginTopPx = distTPx;
            marginBottomPx = Math.max(distBPx, 4);
          } else if (posVOffsetPx > 0) {
            marginTopPx = Math.max(distTPx, Math.min(240, posVOffsetPx));
          }
        }

        if (isCenteredAnchor && (wrapType === 'topAndBottom' || wrapType === 'none')) {
          const imgHtml = renderImageHtml(dataUrl, widthPx, heightPx, 'center', false, alt, {
            wrapType,
            floatTopPx: posVOffsetPx,
            relativeFromH,
            relativeFromV,
            marginTopPx,
            marginBottomPx,
            ooxmlDrawingXml: drawingEl.rawXml,
          });
          if (placeBefore) beforeBlockHtmls.push(imgHtml);
          else blockHtmls.push(imgHtml);
        } else {
          const contentLeftPx = Math.max(0, Math.round(pageLeftPx - geom.marginLeftPx));
          const floatAlign: 'left' | 'center' | 'right' =
            isCenteredAnchor
              ? 'center'
              : posHAlignTxt === 'right' || pageLeftPx + widthPx / 2 > geom.pageWidthPx * 0.6
                ? 'right'
                : 'left';
          const imgHtml = renderImageHtml(dataUrl, widthPx, heightPx, floatAlign, true, alt, {
            wrapType,
            floatLeftPx: contentLeftPx,
            floatTopPx: posVOffsetPx,
            relativeFromH,
            relativeFromV,
            marginTopPx,
            marginBottomPx,
            ooxmlDrawingXml: drawingEl.rawXml,
          });
          if (placeBefore) beforeBlockHtmls.push(imgHtml);
          else blockHtmls.push(imgHtml);
        }
      } else {
        // Emit as a block image so TipTap's block DocImage node preserves OOXML source metadata cleanly
        const imgHtml = renderImageHtml(dataUrl, widthPx, heightPx, align, false, alt, {
          ooxmlDrawingXml: drawingEl.rawXml,
        });
        blockHtmls.push(imgHtml);
      }
    }
  }

  return {inlineHtml: inlineHtmlParts.join(''), beforeBlockHtmls, blockHtmls};
}

// ---------------------------------------------------------------------------
// Paragraph & Run Conversion
// ---------------------------------------------------------------------------

function extractParagraphListInfo(
  pEl: XmlNode,
  resolvedPPr: ResolvedParaProps,
  ctx: ImportContext,
): {isList: boolean; isOrdered: boolean; level: number; numId: string} | null {
  if (!resolvedPPr.numId) return null;
  const numLevels = ctx.numbering.get(resolvedPPr.numId);
  const ilvl = resolvedPPr.ilvl ?? 0;
  const spec = numLevels?.get(ilvl) ?? numLevels?.get(0);
  if (!spec) {
    // Fallback if numbering.xml didn't define this numId
    const pStyle = attr(firstChild(firstChild(pEl, 'w:pPr'), 'w:pStyle'), 'val') ?? '';
    return {
      isList: true,
      isOrdered: /number/i.test(pStyle),
      level: ilvl,
      numId: resolvedPPr.numId,
    };
  }
  return {
    isList: true,
    isOrdered: spec.isOrdered,
    level: ilvl,
    numId: resolvedPPr.numId,
  };
}

function resolveHeadingLevel(styleId: string | undefined, styleName: string | undefined, pPr: ResolvedParaProps): 1 | 2 | 3 | 4 | null {
  const probe = `${styleId ?? ''} ${styleName ?? ''}`;
  const m = probe.match(/\bheading\s*([1-4])\b/i);
  if (m) return Number(m[1]) as 1 | 2 | 3 | 4;
  if (pPr.outlineLvl !== undefined && pPr.outlineLvl >= 0 && pPr.outlineLvl <= 3) {
    return (pPr.outlineLvl + 1) as 1 | 2 | 3 | 4;
  }
  return null;
}

type RunEvent =
  | {kind: 'inline'; html: string}
  | {kind: 'block-before'; html: string}
  | {kind: 'block-after'; html: string}
  | {kind: 'page-break'};

function collectRunPieces(
  runEl: XmlNode,
  baseRPr: ResolvedRunProps,
  paraAlign: ResolvedParaProps['align'],
  ctx: ImportContext,
  tabStops: DocTabStop[] | undefined,
  tabState: {index: number},
  isHeading = false,
): RunEvent[] {
  const events: RunEvent[] = [];

  const rPrEl = firstChild(runEl, 'w:rPr');
  const rStyleId = attr(firstChild(rPrEl, 'w:rStyle'), 'val');
  const charStyleRPr = rStyleId ? resolveStyleChain(ctx.styles, rStyleId).rPr : {};
  const directRPr = parseRunPropsEl(rPrEl, ctx.theme);
  const mergedRPr = mergeDefined(mergeDefined(baseRPr, charStyleRPr), directRPr);

  const defaults = {
    fontFamily: ctx.defaultFontFamily,
    fontSizePt: ctx.defaultFontSizePt,
    color: ctx.defaultColor,
  };

  const processChild = (child: XmlNode) => {
    if (child.tag === 'w:t') {
      const html = renderRunHtml(textContent(child), mergedRPr, defaults, isHeading);
      if (html) events.push({kind: 'inline', html});
    } else if (child.tag === 'w:tab') {
      const activeStops = tabStops?.filter((s) => s.val !== 'clear');
      const stop = activeStops?.[tabState.index];
      tabState.index += 1;
      events.push({kind: 'inline', html: renderTabHtml(stop, ctx.defaultTabStopPx)});
    } else if (child.tag === 'w:br') {
      const brType = (attr(child, 'type') ?? '').toLowerCase();
      if (brType === 'page') {
        events.push({kind: 'page-break'});
      } else {
        events.push({kind: 'inline', html: '<br />'});
      }
    } else if (child.tag === 'w:cr') {
      events.push({kind: 'inline', html: '<br />'});
    } else if (child.tag === 'w:drawing') {
      const res = parseDrawingElement(child, paraAlign, ctx);
      for (const b of res.beforeBlockHtmls) events.push({kind: 'block-before', html: b});
      if (res.inlineHtml) events.push({kind: 'inline', html: res.inlineHtml});
      for (const b of res.blockHtmls) events.push({kind: 'block-after', html: b});
    } else if (child.tag === 'mc:AlternateContent') {
      // ALWAYS use mc:Choice and ignore mc:Fallback so shapes/WordArt never duplicate
      ctx.stats.alternateContentResolved += 1;
      const choices = childEls(child, 'mc:Choice');
      const chosen =
        choices.find(
          (c) =>
            findDescendants(c, 'wps:wsp').length > 0 ||
            findDescendants(c, 'wpg:wgp').length > 0 ||
            findDescendants(c, 'w:drawing').length > 0,
        ) ?? choices[0];
      if (chosen) {
        for (const c of childEls(chosen)) processChild(c);
      }
    } else if (child.tag === 'w:pict') {
      // Legacy VML drawing outside AlternateContent (mc:Fallback was already skipped above)
      const imgData = findDescendants(child, 'v:imagedata')[0];
      const rId = imgData ? attr(imgData, 'id') : undefined;
      const rel = rId ? ctx.rels.get(rId) : undefined;
      const dataUrl = rel ? ctx.mediaDataUrl.get(rel.target) : undefined;
      if (dataUrl) {
        ctx.stats.images += 1;
        events.push({
          kind: 'inline',
          html: renderImageHtml(dataUrl, 240, 120, 'left', false, 'VML image'),
        });
      } else {
        ctx.hasLockedElements = true;
        ctx.stats.lockedElements += 1;
        addWarning(
          ctx,
          'legacy-vml',
          'Legacy VML Shape',
          'Legacy VML drawings cannot be edited directly and are shown as locked placeholders.',
        );
        events.push({
          kind: 'block-after',
          html: renderLockedElementHtml(
            'Legacy VML Drawing (Locked)',
            'Legacy VML shapes are preserved in View mode.',
          ),
        });
      }
    } else if (child.tag === 'w:object') {
      ctx.hasLockedElements = true;
      ctx.stats.lockedElements += 1;
      addWarning(
        ctx,
        'ole-object',
        'Embedded OLE Object',
        'Embedded OLE objects cannot be edited in the browser and are shown as locked placeholders.',
      );
      events.push({
        kind: 'block-after',
        html: renderLockedElementHtml(
          'Embedded OLE Object (Locked)',
          'OLE embedded objects require desktop Word to edit.',
        ),
      });
    }
  };

  for (const child of childEls(runEl)) {
    if (child.tag === 'w:rPr') continue;
    processChild(child);
  }

  return events;
}

function buildParagraphAttributes(
  pPr: ResolvedParaProps,
  styleId: string | undefined,
  _isHeading: boolean,
): string {
  const attrs: string[] = [];
  const styles: string[] = [];

  if (styleId) attrs.push(`data-style-id="${escapeHtmlAttr(styleId)}"`);
  if (pPr.align && pPr.align !== 'left') styles.push(`text-align: ${pPr.align}`);
  if (pPr.indentLeftPx && pPr.indentLeftPx > 0) {
    attrs.push(`data-indent-left="${pPr.indentLeftPx}"`);
    styles.push(`margin-left: ${pPr.indentLeftPx}px`);
  }
  if (pPr.indentRightPx && pPr.indentRightPx > 0) {
    attrs.push(`data-indent-right="${pPr.indentRightPx}"`);
    styles.push(`margin-right: ${pPr.indentRightPx}px`);
  }
  if (pPr.firstLineIndentPx && pPr.firstLineIndentPx !== 0) {
    attrs.push(`data-first-line-indent="${pPr.firstLineIndentPx}"`);
    styles.push(`text-indent: ${pPr.firstLineIndentPx}px`);
  }
  if (pPr.spaceBeforePt !== undefined) {
    attrs.push(`data-space-before="${pPr.spaceBeforePt}"`);
    styles.push(`margin-top: ${pPr.spaceBeforePt}pt`);
  }
  if (pPr.spaceAfterPt !== undefined) {
    attrs.push(`data-space-after="${pPr.spaceAfterPt}"`);
    styles.push(`margin-bottom: ${pPr.spaceAfterPt}pt`);
  }
  if (pPr.lineSpacing) {
    attrs.push(`data-line-spacing="${escapeHtmlAttr(pPr.lineSpacing)}"`);
    styles.push(`line-height: ${pPr.lineSpacing}`);
  }
  if (pPr.keepNext) attrs.push('data-keep-next="true"');
  if (pPr.keepLines) attrs.push('data-keep-lines="true"');
  if (pPr.tabStops && pPr.tabStops.length > 0) {
    attrs.push(`data-tab-stops="${escapeHtmlAttr(JSON.stringify(pPr.tabStops))}"`);
  }

  if (styles.length > 0) attrs.push(`style="${escapeHtmlAttr(styles.join('; '))}"`);
  return attrs.length > 0 ? ` ${attrs.join(' ')}` : '';
}

function convertParagraphNode(
  pEl: XmlNode,
  ctx: ImportContext,
  tableRunOverride?: ResolvedRunProps,
  skipListWrapper = false,
): string[] {
  const outBlocks: string[] = [];
  const pPrEl = firstChild(pEl, 'w:pPr');
  const styleId = attr(firstChild(pPrEl, 'w:pStyle'), 'val');
  const inherited = resolveStyleChain(ctx.styles, styleId);
  const directPPr = parseParaPropsEl(pPrEl);

  // Per ECMA-376 §17.3.1.29, w:pPr/w:rPr formats ONLY the paragraph mark glyph (¶)
  // and is NEVER inherited by <w:r> text runs inside the paragraph.
  const mergedPPr = mergeParaProps(inherited.pPr, directPPr);
  let mergedRPr: ResolvedRunProps = {...inherited.rPr};
  if (tableRunOverride) {
    mergedRPr = mergeDefined(tableRunOverride, mergedRPr);
  }

  const headingLevel = skipListWrapper
    ? null
    : resolveHeadingLevel(styleId, inherited.name, mergedPPr);
  const tag = headingLevel ? `h${headingLevel}` : 'p';
  const attrStr = buildParagraphAttributes(mergedPPr, styleId, Boolean(headingLevel));

  if (mergedPPr.pageBreakBefore) {
    outBlocks.push('<div data-page-break="true">Page Break</div>');
  }

  let inlineBuf = '';
  let beforeBlocks: string[] = [];
  let afterBlocks: string[] = [];
  let emittedAnyBlockOrBreak = false;
  const tabState = {index: 0};

  const flushSegment = () => {
    const hasInline = inlineBuf.trim().length > 0;
    if (beforeBlocks.length > 0) {
      outBlocks.push(...beforeBlocks);
      beforeBlocks = [];
      emittedAnyBlockOrBreak = true;
    }
    if (hasInline) {
      outBlocks.push(`<${tag}${attrStr}>${inlineBuf}</${tag}>`);
      inlineBuf = '';
      emittedAnyBlockOrBreak = true;
    }
    if (afterBlocks.length > 0) {
      outBlocks.push(...afterBlocks);
      afterBlocks = [];
      emittedAnyBlockOrBreak = true;
    }
  };

  const applyEvents = (events: RunEvent[]) => {
    for (const ev of events) {
      if (ev.kind === 'inline') {
        inlineBuf += ev.html;
      } else if (ev.kind === 'block-before') {
        beforeBlocks.push(ev.html);
      } else if (ev.kind === 'block-after') {
        afterBlocks.push(ev.html);
      } else if (ev.kind === 'page-break') {
        flushSegment();
        outBlocks.push('<div data-page-break="true">Page Break</div>');
        emittedAnyBlockOrBreak = true;
        tabState.index = 0;
      }
    }
  };

  const processParaChild = (child: XmlNode) => {
    if (child.tag === 'w:pPr') return;
    if (child.tag === 'w:r') {
      const events = collectRunPieces(
        child,
        mergedRPr,
        mergedPPr.align,
        ctx,
        mergedPPr.tabStops,
        tabState,
        Boolean(headingLevel),
      );
      applyEvents(events);
    } else if (child.tag === 'w:hyperlink') {
      const rId = attr(child, 'id');
      const rel = rId ? ctx.rels.get(rId) : undefined;
      const href = rel?.target ?? (attr(child, 'anchor') ? `#${attr(child, 'anchor')}` : '');
      let linkInner = '';
      for (const lr of childEls(child, 'w:r')) {
        const events = collectRunPieces(
          lr,
          mergedRPr,
          mergedPPr.align,
          ctx,
          mergedPPr.tabStops,
          tabState,
          Boolean(headingLevel),
        );
        for (const ev of events) {
          if (ev.kind === 'inline') linkInner += ev.html;
          else if (ev.kind === 'block-before') beforeBlocks.push(ev.html);
          else if (ev.kind === 'block-after') afterBlocks.push(ev.html);
        }
      }
      if (linkInner) {
        inlineBuf += href
          ? `<a href="${escapeHtmlAttr(href)}" target="_blank" rel="noopener noreferrer">${linkInner}</a>`
          : linkInner;
      }
    } else if (child.tag === 'mc:AlternateContent') {
      ctx.stats.alternateContentResolved += 1;
      const choices = childEls(child, 'mc:Choice');
      const chosen =
        choices.find(
          (c) =>
            findDescendants(c, 'wps:wsp').length > 0 ||
            findDescendants(c, 'w:drawing').length > 0 ||
            findDescendants(c, 'w:r').length > 0,
        ) ?? choices[0];
      if (chosen) {
        for (const c of childEls(chosen)) processParaChild(c);
      }
    } else if (child.tag === 'm:oMath' || child.tag === 'm:oMathPara') {
      const formula = textContent(child).replace(/\s+/g, ' ').trim();
      if (formula) {
        addWarning(
          ctx,
          'omml-math',
          'Word Equation (OMML)',
          'Math equations were imported as plain formula blocks; complex stacked equation layouts are simplified in Edit mode.',
          'info',
        );
        afterBlocks.push(`<div data-doc-equation="true">${escapeHtmlText(formula)}</div>`);
      }
    } else if (child.tag === 'w:sdt') {
      // Structured Document Tag (content control): unwrap w:sdtContent
      const sdtContent = firstChild(child, 'w:sdtContent');
      if (sdtContent) {
        for (const c of childEls(sdtContent)) processParaChild(c);
      }
    } else if (child.tag === 'w:ins') {
      for (const c of childEls(child)) processParaChild(c);
    }
  };

  for (const child of childEls(pEl)) {
    processParaChild(child);
  }

  ctx.stats.paragraphs += 1;

  // Flush remaining content; preserve intentional empty spacer paragraphs only if nothing else was emitted
  flushSegment();
  if (!emittedAnyBlockOrBreak) {
    outBlocks.push(`<${tag}${attrStr}></${tag}>`);
  }

  return outBlocks;
}

// ---------------------------------------------------------------------------
// Table Conversion (tblGrid, tcW, gridSpan, vMerge, borders, shading, vAlign)
// ---------------------------------------------------------------------------

function convertTableNode(tblEl: XmlNode, ctx: ImportContext): string {
  ctx.stats.tables += 1;

  const tblPr = firstChild(tblEl, 'w:tblPr');
  const tblStyleId = attr(firstChild(tblPr, 'w:tblStyle'), 'val');
  const inheritedTbl = tblStyleId ? resolveStyleChain(ctx.styles, tblStyleId) : null;

  const tblBorders: ParsedBorderSet = {
    ...(inheritedTbl?.tblPr?.borders ?? {}),
    ...parseBorderSet(firstChild(tblPr, 'w:tblBorders'), ctx.theme.colors),
  };
  const defaultCellMargins =
    parseCellMarginsEl(firstChild(tblPr, 'w:tblCellMar')) ??
    inheritedTbl?.tblPr?.cellMarginsPx ??
    {top: 4, right: 8, bottom: 4, left: 8};

  // Table alignment & width
  const tblJcRaw = (attr(firstChild(tblPr, 'w:jc'), 'val') ?? '').toLowerCase();
  const tableAlign =
    tblJcRaw === 'center' ? 'center' : tblJcRaw === 'right' || tblJcRaw === 'end' ? 'right' : 'left';
  const tblWEl = firstChild(tblPr, 'w:tblW');
  const tblWType = (attr(tblWEl, 'type') ?? '').toLowerCase();
  const tblWVal = num(attr(tblWEl, 'w'), 0);

  // Column widths from w:tblGrid > w:gridCol
  const tblGrid = firstChild(tblEl, 'w:tblGrid');
  const gridColWidthsPx = childEls(tblGrid, 'w:gridCol').map((gc) =>
    Math.max(24, twipsToPx(num(attr(gc, 'w'), 1440))),
  );
  const gridTotalPx = gridColWidthsPx.reduce((a, b) => a + b, 0);
  const tableWidthPx =
    tblWType === 'dxa' && tblWVal > 0
      ? twipsToPx(tblWVal)
      : gridTotalPx > 0
        ? gridTotalPx
        : undefined;

  // First pass: collect rows & resolve vertical merges (vMerge -> rowspan)
  interface CellGridEntry {
    tcEl: XmlNode;
    isHeader: boolean;
    colStart: number;
    colSpan: number;
    rowSpan: number;
    skipped: boolean;
  }

  const trEls = childEls(tblEl, 'w:tr');
  const rowCount = trEls.length;
  const grid: CellGridEntry[][] = [];
  const activeVmergeByCol = new Map<number, CellGridEntry>();

  for (let rIdx = 0; rIdx < rowCount; rIdx++) {
    const trEl = trEls[rIdx];
    const trPr = firstChild(trEl, 'w:trPr');
    const isHeaderRow = Boolean(firstChild(trPr, 'w:tblHeader'));
    const rowEntries: CellGridEntry[] = [];
    let colCursor = 0;

    for (const tcEl of childEls(trEl, 'w:tc')) {
      const tcPr = firstChild(tcEl, 'w:tcPr');
      const colSpan = Math.max(1, num(attr(firstChild(tcPr, 'w:gridSpan'), 'val'), 1));
      const vMergeEl = firstChild(tcPr, 'w:vMerge');

      if (vMergeEl) {
        const vVal = (attr(vMergeEl, 'val') ?? '').toLowerCase();
        if (vVal === 'restart') {
          const entry: CellGridEntry = {
            tcEl,
            isHeader: isHeaderRow,
            colStart: colCursor,
            colSpan,
            rowSpan: 1,
            skipped: false,
          };
          rowEntries.push(entry);
          for (let c = colCursor; c < colCursor + colSpan; c++) {
            activeVmergeByCol.set(c, entry);
          }
        } else {
          // Continuation of vertical merge: increment rowspan on top anchor cell & skip emitting this cell
          const anchor = activeVmergeByCol.get(colCursor);
          if (anchor) anchor.rowSpan += 1;
          rowEntries.push({
            tcEl,
            isHeader: isHeaderRow,
            colStart: colCursor,
            colSpan,
            rowSpan: 0,
            skipped: true,
          });
        }
      } else {
        for (let c = colCursor; c < colCursor + colSpan; c++) {
          activeVmergeByCol.delete(c);
        }
        rowEntries.push({
          tcEl,
          isHeader: isHeaderRow,
          colStart: colCursor,
          colSpan,
          rowSpan: 1,
          skipped: false,
        });
      }

      colCursor += colSpan;
    }
    grid.push(rowEntries);
  }

  const totalCols = Math.max(
    gridColWidthsPx.length,
    ...grid.map((row) =>
      row.reduce((acc, cell) => Math.max(acc, cell.colStart + cell.colSpan), 0),
    ),
  );

  // Second pass: serialize rows & cells with exact borders, shading, vAlign, colwidth
  const rowsHtml: string[] = [];
  for (let rIdx = 0; rIdx < rowCount; rIdx++) {
    const cellsHtml: string[] = [];
    for (const cell of grid[rIdx]) {
      if (cell.skipped) continue;
      const tcPr = firstChild(cell.tcEl, 'w:tcPr');
      const tcBorders = parseBorderSet(firstChild(tcPr, 'w:tcBorders'), ctx.theme.colors);
      const cellMar = parseCellMarginsEl(firstChild(tcPr, 'w:tcMar')) ?? defaultCellMargins;

      // Resolve borders per side (cell override -> table outer/inside border)
      const isTopEdge = rIdx === 0;
      const isBottomEdge = rIdx + cell.rowSpan >= rowCount;
      const isLeftEdge = cell.colStart === 0;
      const isRightEdge = cell.colStart + cell.colSpan >= totalCols;

      const topCss = borderSideToCss(tcBorders.top ?? (isTopEdge ? tblBorders.top : tblBorders.insideH));
      const bottomCss = borderSideToCss(
        tcBorders.bottom ?? (isBottomEdge ? tblBorders.bottom : tblBorders.insideH),
      );
      const leftCss = borderSideToCss(
        tcBorders.left ?? (isLeftEdge ? tblBorders.left : tblBorders.insideV),
      );
      const rightCss = borderSideToCss(
        tcBorders.right ?? (isRightEdge ? tblBorders.right : tblBorders.insideV),
      );

      // Cell shading
      const shdEl = firstChild(tcPr, 'w:shd');
      const bgColor = shdEl
        ? normalizeColor(attr(shdEl, 'fill'), ctx.theme.colors, attr(shdEl, 'themeFill'))
        : undefined;

      // Vertical alignment
      const vAlignRaw = (attr(firstChild(tcPr, 'w:vAlign'), 'val') ?? '').toLowerCase();
      const vAlign =
        vAlignRaw === 'center' ? 'middle' : vAlignRaw === 'bottom' ? 'bottom' : 'top';

      // Exact column widths for TipTap colwidth attribute
      const colWidthsForCell: number[] = [];
      for (let c = cell.colStart; c < cell.colStart + cell.colSpan; c++) {
        if (gridColWidthsPx[c]) colWidthsForCell.push(gridColWidthsPx[c]);
      }
      if (colWidthsForCell.length === 0) {
        const tcWEl = firstChild(tcPr, 'w:tcW');
        const tcWType = (attr(tcWEl, 'type') ?? '').toLowerCase();
        const tcWVal = num(attr(tcWEl, 'w'), 0);
        if (tcWType === 'dxa' && tcWVal > 0) {
          colWidthsForCell.push(Math.max(24, twipsToPx(tcWVal)));
        }
      }

      const cellAttrs: string[] = [];
      const cellStyles: string[] = [];
      if (cell.colSpan > 1) cellAttrs.push(`colspan="${cell.colSpan}"`);
      if (cell.rowSpan > 1) cellAttrs.push(`rowspan="${cell.rowSpan}"`);
      if (colWidthsForCell.length > 0) {
        cellAttrs.push(`colwidth="${colWidthsForCell.join(',')}"`);
        const sumW = colWidthsForCell.reduce((a, b) => a + b, 0);
        cellStyles.push(`width: ${sumW}px`);
      }
      if (bgColor) {
        cellAttrs.push(`data-bg-color="${escapeHtmlAttr(bgColor)}"`);
        cellStyles.push(`background-color: ${bgColor}`);
      }
      if (vAlign) {
        cellAttrs.push(`data-valign="${vAlign}"`);
        cellStyles.push(`vertical-align: ${vAlign}`);
      }
      if (topCss) {
        cellAttrs.push(`data-border-top="${escapeHtmlAttr(topCss)}"`);
        cellStyles.push(`border-top: ${topCss}`);
      }
      if (rightCss) {
        cellAttrs.push(`data-border-right="${escapeHtmlAttr(rightCss)}"`);
        cellStyles.push(`border-right: ${rightCss}`);
      }
      if (bottomCss) {
        cellAttrs.push(`data-border-bottom="${escapeHtmlAttr(bottomCss)}"`);
        cellStyles.push(`border-bottom: ${bottomCss}`);
      }
      if (leftCss) {
        cellAttrs.push(`data-border-left="${escapeHtmlAttr(leftCss)}"`);
        cellStyles.push(`border-left: ${leftCss}`);
      }
      const padCss = `${cellMar.top}px ${cellMar.right}px ${cellMar.bottom}px ${cellMar.left}px`;
      cellAttrs.push(`data-cell-padding="${escapeHtmlAttr(padCss)}"`);
      cellStyles.push(`padding: ${padCss}`);

      if (cellStyles.length > 0) {
        cellAttrs.push(`style="${escapeHtmlAttr(cellStyles.join('; '))}"`);
      }

      // Cell inner paragraphs (and nested tables flattened safely)
      const innerParts: string[] = [];
      for (const tcChild of childEls(cell.tcEl)) {
        if (tcChild.tag === 'w:p') {
          innerParts.push(...convertParagraphNode(tcChild, ctx, inheritedTbl?.rPr, true));
        } else if (tcChild.tag === 'w:tbl') {
          innerParts.push(convertTableNode(tcChild, ctx));
        }
      }
      if (innerParts.length === 0) innerParts.push('<p></p>');

      const cellTag = cell.isHeader ? 'th' : 'td';
      cellsHtml.push(
        `<${cellTag}${cellAttrs.length ? ` ${cellAttrs.join(' ')}` : ''}>${innerParts.join('')}</${cellTag}>`,
      );
    }
    rowsHtml.push(`<tr>${cellsHtml.join('')}</tr>`);
  }

  const colgroupHtml =
    gridColWidthsPx.length > 0
      ? `<colgroup>${gridColWidthsPx.map((w) => `<col style="width: ${w}px" />`).join('')}</colgroup>`
      : '';

  const tblAttrs: string[] = [`data-table-align="${tableAlign}"`];
  const tblStyles: string[] = ['border-collapse: collapse', 'table-layout: fixed'];
  if (tableWidthPx && tableWidthPx > 0) {
    tblAttrs.push(`data-table-width="${tableWidthPx}"`);
    tblStyles.push(`width: ${tableWidthPx}px`, 'max-width: 100%');
  } else {
    tblStyles.push('width: 100%');
  }
  if (tableAlign === 'center') tblStyles.push('margin-left: auto', 'margin-right: auto');
  else if (tableAlign === 'right') tblStyles.push('margin-left: auto', 'margin-right: 0');
  tblAttrs.push(`style="${escapeHtmlAttr(tblStyles.join('; '))}"`);

  return `<table ${tblAttrs.join(' ')}>${colgroupHtml}<tbody>${rowsHtml.join('')}</tbody></table>`;
}

// ---------------------------------------------------------------------------
// Main Entrypoint: importDocxToHtml
// ---------------------------------------------------------------------------

/**
 * High-fidelity OOXML (.docx) importer built on JSZip + XML DOM parsing.
 * Replaces mammoth.convertToHtml for Edit Mode by preserving fonts, sizes,
 * colors, alignment, indents, spacing, style inheritance, numbering lists,
 * fixed-width tables (tblGrid, borders, shading, merges, vAlign), inline/anchored
 * images, DrawingML text boxes & rounded rectangles (mc:Choice only), page borders,
 * margins, and headers/footers.
 */
export async function importDocxToHtml(buffer: ArrayBuffer): Promise<DocxImportResult> {
  const zip = await JSZip.loadAsync(buffer);

  const docXmlStr = await zip.file('word/document.xml')?.async('string');
  if (!docXmlStr) {
    throw new Error('Invalid .docx archive: missing word/document.xml.');
  }

  const [stylesXmlStr, themeXmlStr, numberingXmlStr, settingsXmlStr, relsXmlStr] =
    await Promise.all([
      zip.file('word/styles.xml')?.async('string') ?? Promise.resolve(null),
      zip.file('word/theme/theme1.xml')?.async('string') ?? Promise.resolve(null),
      zip.file('word/numbering.xml')?.async('string') ?? Promise.resolve(null),
      zip.file('word/settings.xml')?.async('string') ?? Promise.resolve(null),
      zip.file('word/_rels/document.xml.rels')?.async('string') ?? Promise.resolve(null),
    ]);

  const theme = parseThemeXml(themeXmlStr);
  const styles = parseStylesXml(stylesXmlStr, theme);
  const numbering = parseNumberingXml(numberingXmlStr);
  const rels = parseRelsXml(relsXmlStr);

  // Preload media files into base64 data URLs
  const mediaDataUrl = new Map<string, string>();
  const mediaPaths = Object.keys(zip.files).filter(
    (p) => p.startsWith('word/media/') && !zip.files[p].dir,
  );
  for (const mediaPath of mediaPaths) {
    const b64 = await zip.file(mediaPath)?.async('base64');
    if (b64) {
      mediaDataUrl.set(mediaPath, `data:${guessMediaMime(mediaPath)};base64,${b64}`);
    }
  }

  const normalDefaults = resolveStyleChain(styles, styles.normalStyleId ?? undefined);
  const defaultFontFamily =
    canonicalizeWordFontFamily(normalDefaults.rPr.fontFamily) ?? normalDefaults.rPr.fontFamily ?? null;
  const defaultFontSizePt = normalDefaults.rPr.fontSizePt ?? null;
  const defaultColor = normalDefaults.rPr.color ?? '#000000';

  const settingsRoot = settingsXmlStr ? parseXml(settingsXmlStr) : null;
  const defaultTabStopTwips = num(
    attr(firstChild(settingsRoot, 'w:defaultTabStop'), 'val'),
    720,
  );
  const defaultTabStopPx = Math.max(24, twipsToPx(defaultTabStopTwips > 0 ? defaultTabStopTwips : 720));

  const docRoot = parseXml(docXmlStr);
  const bodyEl = firstChild(docRoot, 'w:body');
  if (!bodyEl) {
    throw new Error('Invalid word/document.xml: missing w:body element.');
  }

  // Extract section geometry BEFORE body conversion so anchored images know exact page/margin bounds
  const sectPrs = findDescendants(bodyEl, 'w:sectPr');
  const mainSectPr = firstChild(bodyEl, 'w:sectPr') ?? sectPrs[sectPrs.length - 1];
  const mainPgSz = firstChild(mainSectPr, 'w:pgSz');
  const mainPgMar = firstChild(mainSectPr, 'w:pgMar');
  const sectionGeometry: SectionGeometry = {
    pageWidthPx: twipsToPx(num(attr(mainPgSz, 'w'), 11906)),
    pageHeightPx: twipsToPx(num(attr(mainPgSz, 'h'), 16838)),
    marginTopPx: Math.max(16, twipsToPx(num(attr(mainPgMar, 'top'), 1440))),
    marginRightPx: Math.max(16, twipsToPx(num(attr(mainPgMar, 'right'), 1440))),
    marginBottomPx: Math.max(16, twipsToPx(num(attr(mainPgMar, 'bottom'), 1440))),
    marginLeftPx: Math.max(16, twipsToPx(num(attr(mainPgMar, 'left'), 1440))),
  };

  const ctx: ImportContext = {
    theme,
    styles,
    numbering,
    rels,
    mediaDataUrl,
    defaultFontFamily,
    defaultFontSizePt,
    defaultColor,
    defaultTabStopPx,
    sectionGeometry,
    warningsMap: new Map(),
    hasLockedElements: false,
    hasLossyElements: false,
    stats: {
      paragraphs: 0,
      tables: 0,
      images: 0,
      editableShapes: 0,
      lockedElements: 0,
      alternateContentResolved: 0,
    },
  };

  // Detect VBA macros / SmartArt diagrams / Charts at archive level
  if (zip.file('word/vbaProject.bin')) {
    addWarning(
      ctx,
      'vba-macros',
      'VBA Macros',
      'VBA macros (vbaProject.bin) are stripped for security and cannot be edited or executed in the browser.',
    );
  }
  if (Object.keys(zip.files).some((p) => p.startsWith('word/diagrams/'))) {
    addWarning(
      ctx,
      'smartart-part',
      'SmartArt Diagrams',
      'This document includes SmartArt XML parts which cannot be edited directly.',
    );
  }
  if (Object.keys(zip.files).some((p) => p.startsWith('word/charts/'))) {
    addWarning(
      ctx,
      'charts-part',
      'Embedded Charts',
      'This document includes embedded Word charts which cannot be edited directly.',
    );
  }

  // Convert body children with proper list grouping (<ul> / <ol>) and attach original OOXML metadata
  const htmlBlocks: string[] = [];
  let openListStack: Array<{isOrdered: boolean; level: number; numId: string}> = [];
  let ooxmlGroupSeq = 0;

  const attachOoxmlMetadata = (blocks: string[], rawXml: string | undefined): string[] => {
    if (!rawXml || blocks.length === 0) return blocks;
    const groupId = `b${ooxmlGroupSeq++}`;
    const encodedSource = encodeUtf8Base64(rawXml);
    const total = String(blocks.length);
    return blocks.map((blockHtml, idx) => {
      const sig = computeBlockSignatureFromHtml(blockHtml);
      const attrs: Record<string, string> = {
        'data-ooxml-group': groupId,
        'data-ooxml-gi': String(idx),
        'data-ooxml-gc': total,
        'data-ooxml-sig': sig,
      };
      if (idx === 0) {
        attrs['data-ooxml-source'] = encodedSource;
      }
      return injectHtmlAttrs(blockHtml, attrs);
    });
  };

  const closeListsToLevel = (targetLevel: number) => {
    while (openListStack.length > targetLevel) {
      const top = openListStack.pop()!;
      htmlBlocks.push(top.isOrdered ? '</ol>' : '</ul>');
    }
  };

  for (const child of childEls(bodyEl)) {
    if (child.tag === 'w:sectPr') continue;

    if (child.tag === 'w:tbl') {
      closeListsToLevel(0);
      htmlBlocks.push(...attachOoxmlMetadata([convertTableNode(child, ctx)], child.rawXml));
      continue;
    }

    if (child.tag === 'w:p') {
      const pPrEl = firstChild(child, 'w:pPr');
      const styleId = attr(firstChild(pPrEl, 'w:pStyle'), 'val');
      const inherited = resolveStyleChain(styles, styleId);
      const mergedPPr = mergeParaProps(inherited.pPr, parseParaPropsEl(pPrEl));
      const listInfo = extractParagraphListInfo(child, mergedPPr, ctx);

      if (listInfo) {
        const desiredDepth = Math.min(3, listInfo.level) + 1;
        // If list type changed at current depth, close and reopen
        if (
          openListStack.length >= desiredDepth &&
          openListStack[desiredDepth - 1].isOrdered !== listInfo.isOrdered
        ) {
          closeListsToLevel(desiredDepth - 1);
        } else if (openListStack.length > desiredDepth) {
          closeListsToLevel(desiredDepth);
        }
        while (openListStack.length < desiredDepth) {
          const nextLvl = openListStack.length;
          htmlBlocks.push(listInfo.isOrdered ? '<ol>' : '<ul>');
          openListStack.push({
            isOrdered: listInfo.isOrdered,
            level: nextLvl,
            numId: listInfo.numId,
          });
        }
        const itemBlocks = attachOoxmlMetadata(
          convertParagraphNode(child, ctx, undefined, true),
          child.rawXml,
        );
        htmlBlocks.push(`<li>${itemBlocks.join('')}</li>`);
      } else {
        closeListsToLevel(0);
        htmlBlocks.push(...attachOoxmlMetadata(convertParagraphNode(child, ctx), child.rawXml));
      }
      continue;
    }

    if (child.tag === 'w:sdt') {
      closeListsToLevel(0);
      const sdtContent = firstChild(child, 'w:sdtContent');
      if (sdtContent) {
        const sdtBlocks: string[] = [];
        for (const sc of childEls(sdtContent)) {
          if (sc.tag === 'w:p') sdtBlocks.push(...convertParagraphNode(sc, ctx));
          else if (sc.tag === 'w:tbl') sdtBlocks.push(convertTableNode(sc, ctx));
        }
        htmlBlocks.push(...attachOoxmlMetadata(sdtBlocks, child.rawXml));
      }
    }
  }
  closeListsToLevel(0);

  // Build section layout patch (page size, margins, page borders, columns, headers & footers)
  const layoutPatch: Partial<DocPageLayout> = {
    hasImportedStyles: true,
    docDefaultColor: defaultColor,
    defaultTabStopPx,
  };

  if (defaultFontFamily) layoutPatch.docDefaultFont = defaultFontFamily;
  if (defaultFontSizePt) layoutPatch.docDefaultFontSizePt = defaultFontSizePt;

  if (mainSectPr) {
    // Page size & orientation
    const pgSz = firstChild(mainSectPr, 'w:pgSz');
    if (pgSz) {
      const wPx = twipsToPx(num(attr(pgSz, 'w'), 11906));
      const hPx = twipsToPx(num(attr(pgSz, 'h'), 16838));
      const orientRaw = (attr(pgSz, 'orient') ?? '').toLowerCase();
      const orientation = orientRaw === 'landscape' || wPx > hPx ? 'landscape' : 'portrait';
      layoutPatch.orientation = orientation;
      layoutPatch.size = closestPageSize(wPx, hPx);
      layoutPatch.customPageWidthPx = wPx;
      layoutPatch.customPageHeightPx = hPx;
    }

    // Margins
    const pgMar = firstChild(mainSectPr, 'w:pgMar');
    if (pgMar) {
      const marginsPx = {
        top: Math.max(16, twipsToPx(num(attr(pgMar, 'top'), 1440))),
        right: Math.max(16, twipsToPx(num(attr(pgMar, 'right'), 1440))),
        bottom: Math.max(16, twipsToPx(num(attr(pgMar, 'bottom'), 1440))),
        left: Math.max(16, twipsToPx(num(attr(pgMar, 'left'), 1440))),
      };
      layoutPatch.margins = closestMarginPreset(marginsPx);
      layoutPatch.customMarginsPx = marginsPx;
    }

    // Columns
    const colsEl = firstChild(mainSectPr, 'w:cols');
    if (colsEl) {
      const colCount = num(attr(colsEl, 'num'), 1);
      if (colCount === 2 || colCount === 3) layoutPatch.columns = colCount;
      else layoutPatch.columns = 1;
    }

    // Page Borders (w:pgBorders)
    const pgBorders = firstChild(mainSectPr, 'w:pgBorders');
    if (pgBorders) {
      const parsed = parseBorderSet(pgBorders, theme.colors);
      const toSide = (s: ParsedBorderSide | 'none' | undefined): DocPageBorderSide | undefined =>
        s && s !== 'none' ? s : undefined;
      const top = toSide(parsed.top);
      const right = toSide(parsed.right);
      const bottom = toSide(parsed.bottom);
      const left = toSide(parsed.left);
      if (top || right || bottom || left) {
        const offsetFromRaw = (attr(pgBorders, 'offsetFrom') ?? 'text').toLowerCase();
        const displayRaw = attr(pgBorders, 'display') ?? 'allPages';
        const config: DocPageBorderConfig = {
          offsetFrom: offsetFromRaw === 'page' ? 'page' : 'text',
          display:
            displayRaw === 'firstPage' || displayRaw === 'notFirstPage' ? displayRaw : 'allPages',
          top,
          right,
          bottom,
          left,
        };
        layoutPatch.pageBorderConfig = config;
        const primaryVal = (top ?? right ?? bottom ?? left)!.val.toLowerCase();
        const borderStyle: DocBorderStyle = primaryVal === 'double' ? 'double' : 'box';
        layoutPatch.pageBorder = borderStyle;
      }
    }

    // Different first page
    const titlePg = firstChild(mainSectPr, 'w:titlePg');
    if (titlePg && parseOnOff(titlePg) !== false) {
      layoutPatch.differentFirstPage = true;
    }

    // Headers & Footers
    let foundFooterPageNum = false;
    let foundHeaderPageNum = false;
    let footerNumAlign: 'left' | 'center' | 'right' = 'center';

    for (const hRef of childEls(mainSectPr, 'w:headerReference')) {
      const hType = (attr(hRef, 'type') ?? 'default').toLowerCase();
      const rId = attr(hRef, 'id');
      const rel = rId ? rels.get(rId) : undefined;
      if (!rel) continue;
      const hfXml = (await zip.file(rel.target)?.async('string')) ?? null;
      const parsedHf = parseHeaderFooterPart(hfXml);
      if (hType === 'first') {
        layoutPatch.firstPageHeaderText = parsedHf.text;
      } else if (hType === 'default') {
        layoutPatch.headerText = parsedHf.text;
        if (parsedHf.hasPageNumber) foundHeaderPageNum = true;
      }
    }

    for (const fRef of childEls(mainSectPr, 'w:footerReference')) {
      const fType = (attr(fRef, 'type') ?? 'default').toLowerCase();
      const rId = attr(fRef, 'id');
      const rel = rId ? rels.get(rId) : undefined;
      if (!rel) continue;
      const hfXml = (await zip.file(rel.target)?.async('string')) ?? null;
      const parsedHf = parseHeaderFooterPart(hfXml);
      if (fType === 'first') {
        layoutPatch.firstPageFooterText = parsedHf.text;
      } else if (fType === 'default') {
        layoutPatch.footerText = parsedHf.text;
        if (parsedHf.hasPageNumber) {
          foundFooterPageNum = true;
          footerNumAlign = parsedHf.align;
        }
      }
    }

    if (foundFooterPageNum) {
      layoutPatch.showPageNumbers = true;
      layoutPatch.pageNumberPosition =
        footerNumAlign === 'right' ? 'bottom-right' : 'bottom-center';
    } else if (foundHeaderPageNum) {
      layoutPatch.showPageNumbers = true;
      layoutPatch.pageNumberPosition = 'top-right';
    } else {
      layoutPatch.showPageNumbers = false;
    }
  }

  // Check settings.xml for even/odd headers (informational note)
  if (settingsXmlStr && /<w:evenAndOddHeaders\b/.test(settingsXmlStr)) {
    addWarning(
      ctx,
      'even-odd-headers',
      'Different Odd & Even Headers',
      'Separate odd/even page headers are simplified to a unified header in Edit mode.',
      'info',
    );
  }

  const html = htmlBlocks.join('').trim() || '<p></p>';

  return {
    html,
    layoutPatch,
    warnings: Array.from(ctx.warningsMap.values()),
    hasLockedElements: ctx.hasLockedElements,
    hasLossyElements: ctx.hasLossyElements,
    stats: ctx.stats,
  };
}
