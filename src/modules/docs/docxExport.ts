import {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  Header,
  HeadingLevel,
  ImageRun,
  LeaderType,
  LevelFormat,
  Packer,
  PageBorderDisplay,
  PageBorderOffsetFrom,
  PageBreak as DocxPageBreak,
  PageNumber,
  PageOrientation,
  Paragraph,
  ShadingType,
  Tab,
  TabStopType,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  UnderlineType,
  VerticalAlign,
  WidthType,
  type IBorderOptions,
  type IParagraphOptions,
  type IRunOptions,
} from 'docx';
import {canonicalizeWordFontFamily} from './docFonts';
import {patchOriginalDocxPackage} from './docxLosslessSave';
import {
  DEFAULT_DOC_LAYOUT,
  DOC_MARGIN_PX,
  DOC_PAGE_DIMENSIONS_PX,
  DOC_THEMES,
  type DocNoteItem,
  type DocPageBorderSide,
  type DocPageLayout,
  type DocTabStop,
} from './docsTypes';

// ---------------------------------------------------------------------------
// Universal lightweight HTML tree parser (identical in Browser & Node/Vitest)
// ---------------------------------------------------------------------------

interface HtmlNode {
  type: 'element' | 'text';
  tag: string;
  attrs: Record<string, string>;
  children: HtmlNode[];
  text: string;
}

const VOID_HTML_TAGS = new Set([
  'br',
  'img',
  'hr',
  'col',
  'input',
  'meta',
  'link',
  'source',
  'wbr',
]);

function decodeEntities(raw: string): string {
  return raw
    .replace(/&nbsp;/g, '\u00A0')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;|&apos;/g, "'")
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex: string) =>
      String.fromCodePoint(parseInt(hex, 16)),
    )
    .replace(/&#(\d+);/g, (_, dec: string) => String.fromCodePoint(parseInt(dec, 10)))
    .replace(/&amp;/g, '&');
}

function parseHtmlTree(html: string): HtmlNode {
  const root: HtmlNode = {type: 'element', tag: '#root', attrs: {}, children: [], text: ''};
  const stack: HtmlNode[] = [root];

  const tokenRe = /<!--[\s\S]*?-->|<(\/?)([A-Za-z][A-Za-z0-9-]*)([^>]*?)(\/?)>|([^<]+)/g;
  const attrRe = /([A-Za-z0-9_:-]+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'=<>`]+)))?/g;

  let m: RegExpExecArray | null;
  while ((m = tokenRe.exec(html)) !== null) {
    const parent = stack[stack.length - 1];
    if (m[0].startsWith('<!--')) continue;

    if (m[5] !== undefined) {
      parent.children.push({
        type: 'text',
        tag: '#text',
        attrs: {},
        children: [],
        text: decodeEntities(m[5]),
      });
      continue;
    }

    const isClose = m[1] === '/';
    const tag = m[2].toLowerCase();
    const attrStr = m[3] ?? '';
    const isSelfClose = m[4] === '/' || VOID_HTML_TAGS.has(tag);

    if (isClose) {
      for (let i = stack.length - 1; i > 0; i--) {
        if (stack[i].tag === tag) {
          stack.length = i;
          break;
        }
      }
      continue;
    }

    const attrs: Record<string, string> = {};
    attrRe.lastIndex = 0;
    let am: RegExpExecArray | null;
    while ((am = attrRe.exec(attrStr)) !== null) {
      const k = am[1].toLowerCase();
      const v = am[2] ?? am[3] ?? am[4] ?? '';
      attrs[k] = decodeEntities(v);
    }

    const el: HtmlNode = {type: 'element', tag, attrs, children: [], text: ''};
    parent.children.push(el);
    if (!isSelfClose) {
      stack.push(el);
    }
  }

  return root;
}

function parseInlineCss(styleAttr: string | undefined): Record<string, string> {
  const map: Record<string, string> = {};
  if (!styleAttr) return map;
  for (const part of styleAttr.split(';')) {
    const idx = part.indexOf(':');
    if (idx <= 0) continue;
    const key = part.slice(0, idx).trim().toLowerCase();
    const val = part.slice(idx + 1).trim();
    if (key && val) map[key] = val;
  }
  return map;
}

function pxToTwips(px: number): number {
  return Math.round(px * 15);
}

function ptToTwips(pt: number): number {
  return Math.round(pt * 20);
}

function parseLengthPx(raw: string | undefined): number | undefined {
  if (!raw) return undefined;
  const m = raw.trim().match(/^(-?\d+(?:\.\d+)?)\s*(px|pt|em|rem)?$/i);
  if (!m) return undefined;
  const val = Number(m[1]);
  if (!Number.isFinite(val)) return undefined;
  const unit = (m[2] ?? 'px').toLowerCase();
  if (unit === 'pt') return val * (96 / 72);
  if (unit === 'em' || unit === 'rem') return val * 16;
  return val;
}

function parseLengthPt(raw: string | undefined): number | undefined {
  if (!raw) return undefined;
  const m = raw.trim().match(/^(-?\d+(?:\.\d+)?)\s*(pt|px)?$/i);
  if (!m) return undefined;
  const val = Number(m[1]);
  if (!Number.isFinite(val)) return undefined;
  const unit = (m[2] ?? 'pt').toLowerCase();
  return unit === 'px' ? val * 0.75 : val;
}

function normalizeHex6(raw: string | undefined): string | undefined {
  if (!raw) return undefined;
  const trimmed = raw.trim();
  const hexMatch = trimmed.match(/^#?([0-9a-fA-F]{6})$/);
  if (hexMatch) return hexMatch[1].toUpperCase();
  const shortHex = trimmed.match(/^#([0-9a-fA-F]{3})$/);
  if (shortHex) {
    const [r, g, b] = shortHex[1].split('');
    return `${r}${r}${g}${g}${b}${b}`.toUpperCase();
  }
  const rgbMatch = trimmed.match(/^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i);
  if (rgbMatch) {
    const hex = [rgbMatch[1], rgbMatch[2], rgbMatch[3]]
      .map((n) => Math.max(0, Math.min(255, Number(n))).toString(16).padStart(2, '0'))
      .join('');
    return hex.toUpperCase();
  }
  return undefined;
}

function mapDocxBorderStyle(cssStyle: string): (typeof BorderStyle)[keyof typeof BorderStyle] {
  switch (cssStyle.toLowerCase()) {
    case 'none':
    case 'nil':
      return BorderStyle.NONE;
    case 'double':
    case 'triple':
      return BorderStyle.DOUBLE;
    case 'dashed':
      return BorderStyle.DASHED;
    case 'dotted':
      return BorderStyle.DOTTED;
    default:
      return BorderStyle.SINGLE;
  }
}

function parseCssBorder(raw: string | undefined): IBorderOptions | undefined {
  if (!raw) return undefined;
  const trimmed = raw.trim().toLowerCase();
  if (trimmed === 'none' || trimmed === '0') {
    return {style: BorderStyle.NONE, size: 0, color: 'auto'};
  }
  const parts = raw.trim().split(/\s+/);
  let widthPt = 0.5;
  let style: (typeof BorderStyle)[keyof typeof BorderStyle] = BorderStyle.SINGLE;
  let color = 'CBD5E1';

  for (const p of parts) {
    const pt = parseLengthPt(p);
    if (pt !== undefined && /\d/.test(p)) {
      widthPt = Math.max(0.25, pt);
      continue;
    }
    if (/^(solid|double|dashed|dotted|none|ridge|groove)$/i.test(p)) {
      style = mapDocxBorderStyle(p);
      continue;
    }
    const hex = normalizeHex6(p);
    if (hex) color = hex;
  }

  return {
    style,
    size: Math.max(2, Math.round(widthPt * 8)),
    color,
  };
}

function decodeBase64DataUrl(
  dataUrl: string,
): {bytes: Uint8Array; type: 'png' | 'jpg' | 'gif' | 'bmp'} | null {
  const m = dataUrl.match(/^data:image\/([a-zA-Z0-9+.-]+);base64,([\s\S]+)$/);
  if (!m) return null;
  const mimeSub = m[1].toLowerCase();
  const imgType: 'png' | 'jpg' | 'gif' | 'bmp' =
    mimeSub === 'jpeg' || mimeSub === 'jpg'
      ? 'jpg'
      : mimeSub === 'gif'
        ? 'gif'
        : mimeSub === 'bmp'
          ? 'bmp'
          : 'png';
  try {
    const cleanB64 = m[2].replace(/\s+/g, '');
    if (typeof Buffer !== 'undefined') {
      return {bytes: new Uint8Array(Buffer.from(cleanB64, 'base64')), type: imgType};
    }
    const bin = atob(cleanB64);
    const arr = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
    return {bytes: arr, type: imgType};
  } catch {
    return null;
  }
}

// ---------------------------------------------------------------------------
// Inline Run Serialization
// ---------------------------------------------------------------------------

interface RunState {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  sub?: boolean;
  superScript?: boolean;
  color?: string;
  highlightHex?: string;
  fontFamily?: string;
  fontSizeHalfPts?: number;
  allCaps?: boolean;
  smallCaps?: boolean;
  characterSpacingTwips?: number;
}

function collectRunsFromNode(
  node: HtmlNode,
  inherited: RunState,
  out: Array<TextRun | ImageRun>,
): void {
  if (node.type === 'text') {
    if (!node.text) return;
    const opts: IRunOptions = {
      text: node.text,
      bold: inherited.bold || undefined,
      italics: inherited.italic || undefined,
      underline: inherited.underline ? {type: UnderlineType.SINGLE} : undefined,
      strike: inherited.strike || undefined,
      subScript: inherited.sub || undefined,
      superScript: inherited.superScript || undefined,
      color: inherited.color,
      font: inherited.fontFamily,
      size: inherited.fontSizeHalfPts,
      allCaps: inherited.allCaps || undefined,
      smallCaps: inherited.smallCaps || undefined,
      characterSpacing: inherited.characterSpacingTwips,
      shading: inherited.highlightHex
        ? {type: ShadingType.CLEAR, color: 'auto', fill: inherited.highlightHex}
        : undefined,
    };
    out.push(new TextRun(opts));
    return;
  }

  const tag = node.tag;
  if (tag === 'br') {
    out.push(new TextRun({break: 1}));
    return;
  }

  if (tag === 'span' && node.attrs['data-doc-tab'] === 'true') {
    out.push(
      new TextRun({
        children: [new Tab()],
        bold: inherited.bold || undefined,
        italics: inherited.italic || undefined,
        font: inherited.fontFamily,
        size: inherited.fontSizeHalfPts,
        color: inherited.color,
      }),
    );
    return;
  }

  if (tag === 'img') {
    const src = node.attrs.src ?? '';
    const decoded = decodeBase64DataUrl(src);
    if (decoded) {
      const css = parseInlineCss(node.attrs.style);
      const w = Math.max(
        16,
        Math.round(
          Number(node.attrs['data-width']) ||
            parseLengthPx(css.width) ||
            Number(node.attrs.width) ||
            240,
        ),
      );
      const h = Math.max(
        16,
        Math.round(
          Number(node.attrs['data-height']) ||
            parseLengthPx(css.height) ||
            Number(node.attrs.height) ||
            120,
        ),
      );
      out.push(
        new ImageRun({
          data: decoded.bytes,
          type: decoded.type,
          transformation: {width: w, height: h},
        }),
      );
    }
    return;
  }

  const next: RunState = {...inherited};
  if (tag === 'strong' || tag === 'b') next.bold = true;
  else if (tag === 'em' || tag === 'i') next.italic = true;
  else if (tag === 'u') next.underline = true;
  else if (tag === 's' || tag === 'del') next.strike = true;
  else if (tag === 'sub') next.sub = true;
  else if (tag === 'sup') next.superScript = true;
  else if (tag === 'code') next.fontFamily = 'Courier New';

  const css = parseInlineCss(node.attrs.style);
  if (node.attrs['data-doc-font'] || css['font-family']) {
    const canonical = canonicalizeWordFontFamily(
      node.attrs['data-doc-font'] || css['font-family'],
    );
    if (canonical) next.fontFamily = canonical;
  }
  if (css['font-size']) {
    const pt = parseLengthPt(css['font-size']);
    if (pt && pt > 0) next.fontSizeHalfPts = Math.max(8, Math.round(pt * 2));
  }
  if (css.color) {
    const hex = normalizeHex6(css.color);
    if (hex) next.color = hex;
  }
  if (tag === 'mark' || css['background-color']) {
    const hex = normalizeHex6(node.attrs['data-color'] || css['background-color']);
    if (hex) next.highlightHex = hex;
  }
  if (css['text-transform']?.toLowerCase() === 'uppercase') next.allCaps = true;
  if (css['font-variant']?.toLowerCase() === 'small-caps') next.smallCaps = true;
  if (css['letter-spacing']) {
    const lsPt = parseLengthPt(css['letter-spacing']);
    if (lsPt && lsPt !== 0) next.characterSpacingTwips = ptToTwips(lsPt);
  }

  for (const child of node.children) {
    collectRunsFromNode(child, next, out);
  }
}

// ---------------------------------------------------------------------------
// Block Serialization (Paragraphs, Headings, Lists, Tables, Images, TextBoxes)
// ---------------------------------------------------------------------------

const ALIGN_MAP: Record<string, (typeof AlignmentType)[keyof typeof AlignmentType]> = {
  left: AlignmentType.LEFT,
  center: AlignmentType.CENTER,
  right: AlignmentType.RIGHT,
  justify: AlignmentType.JUSTIFIED,
};

const HEADING_MAP: Record<string, (typeof HeadingLevel)[keyof typeof HeadingLevel]> = {
  h1: HeadingLevel.HEADING_1,
  h2: HeadingLevel.HEADING_2,
  h3: HeadingLevel.HEADING_3,
  h4: HeadingLevel.HEADING_4,
};

function buildParagraphFromElement(
  el: HtmlNode,
  options?: {
    bulletLevel?: number;
    orderedLevel?: number;
    extraIndentLeftTwips?: number;
    defaultBold?: boolean;
  },
): Paragraph {
  const css = parseInlineCss(el.attrs.style);
  const runs: Array<TextRun | ImageRun> = [];
  collectRunsFromNode(el, {bold: options?.defaultBold}, runs);

  const alignKey = (css['text-align'] ?? '').toLowerCase();
  const alignment = ALIGN_MAP[alignKey];

  const indentLeftPx =
    Number(el.attrs['data-indent-left']) ||
    (el.attrs['data-indent'] ? Number(el.attrs['data-indent']) * 32 : 0) ||
    parseLengthPx(css['margin-left']) ||
    0;
  const indentRightPx =
    Number(el.attrs['data-indent-right']) || parseLengthPx(css['margin-right']) || 0;
  const firstLinePx =
    Number(el.attrs['data-first-line-indent']) || parseLengthPx(css['text-indent']) || 0;

  const totalLeftTwips = Math.max(
    0,
    pxToTwips(indentLeftPx) + (options?.extraIndentLeftTwips ?? 0),
  );
  const indent: IParagraphOptions['indent'] =
    totalLeftTwips > 0 || indentRightPx > 0 || firstLinePx !== 0
      ? {
          ...(totalLeftTwips > 0 ? {left: totalLeftTwips} : {}),
          ...(indentRightPx > 0 ? {right: pxToTwips(indentRightPx)} : {}),
          ...(firstLinePx > 0
            ? {firstLine: pxToTwips(firstLinePx)}
            : firstLinePx < 0
              ? {hanging: pxToTwips(Math.abs(firstLinePx))}
              : {}),
        }
      : undefined;

  const beforePt =
    el.attrs['data-space-before'] !== undefined
      ? Number(el.attrs['data-space-before'])
      : parseLengthPt(css['margin-top']);
  const afterPt =
    el.attrs['data-space-after'] !== undefined
      ? Number(el.attrs['data-space-after'])
      : parseLengthPt(css['margin-bottom']);
  const lineSpacingRaw = el.attrs['data-line-spacing'] ?? css['line-height'];
  let lineTwips: number | undefined;
  if (lineSpacingRaw) {
    const pt = lineSpacingRaw.endsWith('pt') ? parseLengthPt(lineSpacingRaw) : undefined;
    if (pt && pt > 0) {
      lineTwips = ptToTwips(pt);
    } else {
      const mult = Number(lineSpacingRaw);
      if (Number.isFinite(mult) && mult >= 0.6 && mult <= 4) {
        lineTwips = Math.round(mult * 240);
      }
    }
  }

  const spacing: IParagraphOptions['spacing'] =
    beforePt !== undefined || afterPt !== undefined || lineTwips !== undefined
      ? {
          ...(beforePt !== undefined && Number.isFinite(beforePt)
            ? {before: Math.max(0, ptToTwips(beforePt))}
            : {}),
          ...(afterPt !== undefined && Number.isFinite(afterPt)
            ? {after: Math.max(0, ptToTwips(afterPt))}
            : {}),
          ...(lineTwips !== undefined ? {line: lineTwips} : {}),
        }
      : undefined;

  const isHeading = el.tag in HEADING_MAP;
  const keepNext = el.attrs['data-keep-next'] === 'true' || isHeading ? true : undefined;
  const keepLines = el.attrs['data-keep-lines'] === 'true' ? true : undefined;

  let tabStops: IParagraphOptions['tabStops'];
  if (el.attrs['data-tab-stops']) {
    try {
      const parsed = JSON.parse(el.attrs['data-tab-stops']) as DocTabStop[];
      if (Array.isArray(parsed) && parsed.length > 0) {
        const mapped = parsed
          .filter((s) => s && typeof s.posTwips === 'number' && Number.isFinite(s.posTwips))
          .map((s) => ({
            position: Math.round(s.posTwips),
            type:
              s.val === 'right'
                ? TabStopType.RIGHT
                : s.val === 'center'
                  ? TabStopType.CENTER
                  : s.val === 'decimal'
                    ? TabStopType.DECIMAL
                    : s.val === 'bar'
                      ? TabStopType.BAR
                      : s.val === 'clear'
                        ? TabStopType.CLEAR
                        : TabStopType.LEFT,
            leader:
              s.leader === 'dot'
                ? LeaderType.DOT
                : s.leader === 'hyphen'
                  ? LeaderType.HYPHEN
                  : s.leader === 'underscore'
                    ? LeaderType.UNDERSCORE
                    : s.leader === 'middleDot'
                      ? LeaderType.MIDDLE_DOT
                      : LeaderType.NONE,
          }));
        if (mapped.length > 0) tabStops = mapped;
      }
    } catch {
      // Ignore malformed tab stops JSON
    }
  }

  return new Paragraph({
    heading: HEADING_MAP[el.tag],
    alignment,
    indent,
    spacing,
    tabStops,
    keepNext,
    keepLines,
    widowControl: true,
    bullet: options?.bulletLevel !== undefined ? {level: options.bulletLevel} : undefined,
    numbering:
      options?.orderedLevel !== undefined
        ? {reference: 'esz-numbered-list', level: options.orderedLevel}
        : undefined,
    children: runs.length > 0 ? runs : [new TextRun('')],
  });
}

function buildTableFromElement(tblEl: HtmlNode): Table {
  // Collect column widths from <colgroup><col> or first-row colwidth attributes
  const colWidthsPx: number[] = [];
  const colgroup = tblEl.children.find((c) => c.type === 'element' && c.tag === 'colgroup');
  if (colgroup) {
    for (const col of colgroup.children) {
      if (col.type === 'element' && col.tag === 'col') {
        const w = parseLengthPx(parseInlineCss(col.attrs.style).width);
        if (w && w > 0) colWidthsPx.push(w);
      }
    }
  }

  const trEls: HtmlNode[] = [];
  const collectTrs = (n: HtmlNode) => {
    for (const c of n.children) {
      if (c.type !== 'element') continue;
      if (c.tag === 'tr') trEls.push(c);
      else if (c.tag === 'thead' || c.tag === 'tbody' || c.tag === 'tfoot') collectTrs(c);
    }
  };
  collectTrs(tblEl);

  if (colWidthsPx.length === 0 && trEls.length > 0) {
    for (const cell of trEls[0].children) {
      if (cell.type !== 'element' || (cell.tag !== 'td' && cell.tag !== 'th')) continue;
      const cw = cell.attrs.colwidth;
      if (cw) {
        for (const piece of cw.split(',')) {
          const n = Number(piece.trim());
          if (Number.isFinite(n) && n > 0) colWidthsPx.push(n);
        }
      }
    }
  }

  const docxRows: TableRow[] = trEls.map((trEl, rIdx) => {
    const cells = trEl.children.filter(
      (c) => c.type === 'element' && (c.tag === 'td' || c.tag === 'th'),
    );
    const isHeaderRow = cells.length > 0 && cells.every((c) => c.tag === 'th');

    const docxCells = cells.map((cellEl) => {
      const css = parseInlineCss(cellEl.attrs.style);
      const colSpan = Math.max(1, Number(cellEl.attrs.colspan) || 1);
      const rowSpan = Math.max(1, Number(cellEl.attrs.rowspan) || 1);

      let cellWidthPx = 0;
      if (cellEl.attrs.colwidth) {
        cellWidthPx = cellEl.attrs.colwidth
          .split(',')
          .map((s) => Number(s.trim()))
          .filter((n) => Number.isFinite(n) && n > 0)
          .reduce((a, b) => a + b, 0);
      }
      if (!cellWidthPx) {
        cellWidthPx = parseLengthPx(css.width) ?? 0;
      }

      const bgHex = normalizeHex6(cellEl.attrs['data-bg-color'] || css['background-color']);
      const vAlignRaw = (cellEl.attrs['data-valign'] || css['vertical-align'] || '').toLowerCase();
      const verticalAlign =
        vAlignRaw === 'middle' || vAlignRaw === 'center'
          ? VerticalAlign.CENTER
          : vAlignRaw === 'bottom'
            ? VerticalAlign.BOTTOM
            : VerticalAlign.TOP;

      const borderTop = parseCssBorder(cellEl.attrs['data-border-top'] || css['border-top']);
      const borderRight = parseCssBorder(cellEl.attrs['data-border-right'] || css['border-right']);
      const borderBottom = parseCssBorder(
        cellEl.attrs['data-border-bottom'] || css['border-bottom'],
      );
      const borderLeft = parseCssBorder(cellEl.attrs['data-border-left'] || css['border-left']);
      const hasCustomBorders = Boolean(borderTop || borderRight || borderBottom || borderLeft);

      const cellChildren: Array<Paragraph | Table> = [];
      for (const child of cellEl.children) {
        if (child.type === 'text') {
          if (child.text.trim()) {
            cellChildren.push(
              new Paragraph({
                children: [new TextRun({text: child.text, bold: cellEl.tag === 'th'})],
              }),
            );
          }
        } else {
          appendBlockNode(child, cellChildren, 0);
        }
      }
      if (cellChildren.length === 0) cellChildren.push(new Paragraph(''));

      return new TableCell({
        columnSpan: colSpan > 1 ? colSpan : undefined,
        rowSpan: rowSpan > 1 ? rowSpan : undefined,
        width:
          cellWidthPx > 0 ? {size: pxToTwips(cellWidthPx), type: WidthType.DXA} : undefined,
        verticalAlign,
        shading: bgHex
          ? {type: ShadingType.CLEAR, color: 'auto', fill: bgHex}
          : cellEl.tag === 'th'
            ? {type: ShadingType.CLEAR, color: 'auto', fill: 'F1F5F9'}
            : undefined,
        borders: hasCustomBorders
          ? {
              top: borderTop,
              right: borderRight,
              bottom: borderBottom,
              left: borderLeft,
            }
          : undefined,
        children: cellChildren,
      });
    });

    return new TableRow({
      tableHeader: isHeaderRow || rIdx === 0 ? isHeaderRow : undefined,
      cantSplit: true,
      children: docxCells,
    });
  });

  const tblAlignRaw = (tblEl.attrs['data-table-align'] || '').toLowerCase();
  const tableWidthPx =
    Number(tblEl.attrs['data-table-width']) ||
    (colWidthsPx.length > 0 ? colWidthsPx.reduce((a, b) => a + b, 0) : 0);

  return new Table({
    alignment: ALIGN_MAP[tblAlignRaw],
    layout: colWidthsPx.length > 0 ? TableLayoutType.FIXED : TableLayoutType.AUTOFIT,
    columnWidths: colWidthsPx.length > 0 ? colWidthsPx.map((w) => pxToTwips(w)) : undefined,
    width:
      tableWidthPx > 0
        ? {size: pxToTwips(tableWidthPx), type: WidthType.DXA}
        : {size: 100, type: WidthType.PERCENTAGE},
    rows: docxRows.length > 0 ? docxRows : [new TableRow({children: [new TableCell({children: [new Paragraph('')]})]})],
  });
}

function buildTextBoxTableFromElement(boxEl: HtmlNode): Table {
  const widthPx = Math.max(80, Number(boxEl.attrs['data-box-width']) || 440);
  const alignRaw = (boxEl.attrs['data-box-align'] || 'center').toLowerCase();
  const fillHex = normalizeHex6(boxEl.attrs['data-box-fill']);
  const borderColorHex = normalizeHex6(boxEl.attrs['data-box-border-color']);
  const borderStyleRaw = (boxEl.attrs['data-box-border-style'] || 'solid').toLowerCase();
  const borderWidthPx = Number(boxEl.attrs['data-box-border-width']) || (borderColorHex ? 1 : 0);
  const vAlignRaw = (boxEl.attrs['data-box-valign'] || 'top').toLowerCase();

  const borderOpt: IBorderOptions =
    borderColorHex && borderStyleRaw !== 'none' && borderWidthPx > 0
      ? {
          style: mapDocxBorderStyle(borderStyleRaw),
          size: Math.max(4, Math.round(borderWidthPx * 0.75 * 8)),
          color: borderColorHex,
        }
      : {style: BorderStyle.NONE, size: 0, color: 'auto'};

  const innerChildren: Array<Paragraph | Table> = [];
  for (const child of boxEl.children) {
    if (child.type === 'element') {
      appendBlockNode(child, innerChildren, 0);
    } else if (child.text.trim()) {
      innerChildren.push(new Paragraph({children: [new TextRun(child.text)]}));
    }
  }
  if (innerChildren.length === 0) innerChildren.push(new Paragraph(''));

  return new Table({
    alignment: ALIGN_MAP[alignRaw] ?? AlignmentType.CENTER,
    layout: TableLayoutType.FIXED,
    columnWidths: [pxToTwips(widthPx)],
    width: {size: pxToTwips(widthPx), type: WidthType.DXA},
    rows: [
      new TableRow({
        cantSplit: true,
        children: [
          new TableCell({
            width: {size: pxToTwips(widthPx), type: WidthType.DXA},
            verticalAlign:
              vAlignRaw === 'middle'
                ? VerticalAlign.CENTER
                : vAlignRaw === 'bottom'
                  ? VerticalAlign.BOTTOM
                  : VerticalAlign.TOP,
            shading: fillHex ? {type: ShadingType.CLEAR, color: 'auto', fill: fillHex} : undefined,
            borders: {
              top: borderOpt,
              right: borderOpt,
              bottom: borderOpt,
              left: borderOpt,
            },
            margins: {top: 100, right: 140, bottom: 100, left: 140},
            children: innerChildren,
          }),
        ],
      }),
    ],
  });
}

function appendBlockNode(
  node: HtmlNode,
  out: Array<Paragraph | Table>,
  listDepth: number,
  listType?: 'ul' | 'ol',
): void {
  if (node.type === 'text') {
    if (node.text.trim()) {
      out.push(new Paragraph({children: [new TextRun(node.text)]}));
    }
    return;
  }

  const tag = node.tag;

  if (node.attrs['data-page-break'] === 'true') {
    out.push(new Paragraph({children: [new DocxPageBreak()]}));
    return;
  }

  if (node.attrs['data-doc-textbox'] === 'true') {
    out.push(buildTextBoxTableFromElement(node));
    return;
  }

  if (node.attrs['data-doc-locked'] === 'true') {
    const imgSrc = node.attrs['data-locked-image'];
    if (imgSrc) {
      const decoded = decodeBase64DataUrl(imgSrc);
      if (decoded) {
        const w = Math.max(24, Number(node.attrs['data-width']) || 240);
        const h = Math.max(24, Number(node.attrs['data-height']) || 120);
        out.push(
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new ImageRun({
                data: decoded.bytes,
                type: decoded.type,
                transformation: {width: w, height: h},
              }),
            ],
          }),
        );
        return;
      }
    }
    const label = node.attrs['data-locked-label'] || 'Locked Word Element';
    out.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({text: `[${label}]`, italics: true, color: '64748B'})],
      }),
    );
    return;
  }

  if (tag === 'img') {
    const alignRaw = (node.attrs['data-align'] || 'left').toLowerCase();
    const runs: Array<TextRun | ImageRun> = [];
    collectRunsFromNode(node, {}, runs);
    if (runs.length > 0) {
      out.push(
        new Paragraph({
          alignment: ALIGN_MAP[alignRaw],
          children: runs,
        }),
      );
    }
    return;
  }

  if (tag === 'table') {
    out.push(buildTableFromElement(node));
    return;
  }

  if (tag === 'ul' || tag === 'ol') {
    for (const child of node.children) {
      if (child.type === 'element' && child.tag === 'li') {
        appendBlockNode(child, out, listDepth, tag);
      }
    }
    return;
  }

  if (tag === 'li') {
    const clampedLevel = Math.min(3, listDepth);
    let emittedPrimary = false;
    for (const child of node.children) {
      if (child.type === 'element' && (child.tag === 'ul' || child.tag === 'ol')) {
        appendBlockNode(child, out, listDepth + 1, child.tag);
      } else if (child.type === 'element' && child.tag === 'p') {
        out.push(
          buildParagraphFromElement(child, {
            bulletLevel: listType === 'ul' ? clampedLevel : undefined,
            orderedLevel: listType === 'ol' ? clampedLevel : undefined,
          }),
        );
        emittedPrimary = true;
      } else if (child.type === 'element' && child.tag === 'table') {
        out.push(buildTableFromElement(child));
      }
    }
    if (!emittedPrimary) {
      out.push(
        buildParagraphFromElement(node, {
          bulletLevel: listType === 'ul' ? clampedLevel : undefined,
          orderedLevel: listType === 'ol' ? clampedLevel : undefined,
        }),
      );
    }
    return;
  }

  if (tag === 'blockquote') {
    const childParas = node.children.filter((c) => c.type === 'element' && c.tag === 'p');
    if (childParas.length > 0) {
      for (const p of childParas) {
        out.push(buildParagraphFromElement(p, {extraIndentLeftTwips: 720}));
      }
    } else {
      out.push(buildParagraphFromElement(node, {extraIndentLeftTwips: 720}));
    }
    return;
  }

  if (tag === 'p' || tag === 'h1' || tag === 'h2' || tag === 'h3' || tag === 'h4') {
    out.push(buildParagraphFromElement(node));
    return;
  }

  if (tag === 'div' || tag === 'section' || tag === 'article' || tag === 'body') {
    // TipTap wraps tables inside <div class="tableWrapper">
    const hasBlockChildren = node.children.some(
      (c) =>
        c.type === 'element' &&
        ['p', 'h1', 'h2', 'h3', 'h4', 'table', 'ul', 'ol', 'div', 'blockquote', 'img'].includes(
          c.tag,
        ),
    );
    if (hasBlockChildren) {
      for (const child of node.children) {
        appendBlockNode(child, out, listDepth, listType);
      }
    } else if (node.children.length > 0) {
      out.push(buildParagraphFromElement(node));
    }
  }
}

function buildPageNumberRuns(): TextRun[] {
  return [
    new TextRun({
      children: ['Page ', PageNumber.CURRENT, ' of ', PageNumber.TOTAL_PAGES],
      size: 18,
      color: '64748B',
    }),
  ];
}

function buildDocxPageBorderSide(side: DocPageBorderSide | undefined): IBorderOptions | undefined {
  if (!side) return undefined;
  return {
    style: mapDocxBorderStyle(side.val),
    size: Math.max(4, Math.round(side.szPt * 8)),
    color: normalizeHex6(side.color) ?? '000000',
    space: Math.max(1, Math.min(31, Math.round(side.spacePt))),
  };
}

/**
 * High-fidelity HTML -> OOXML (.docx) exporter.
 *
 * When `originalDocxBuffer` is provided, performs a lossless package-preserving save
 * (rewriting only `word/document.xml` and newly added media/relationships, keeping
 * untouched paragraphs/tables/anchors/styles/headers/footers/settings verbatim).
 * Otherwise builds a clean OOXML package from scratch with canonical Word fonts,
 * sizes, run styles, tab stops, paragraph spacing/indents/alignment, fixed-width tables,
 * images, styled boxes, exact page size/margins, page borders, headers, and footers.
 */
export async function exportHtmlToDocxBytes(
  html: string,
  layout: DocPageLayout = DEFAULT_DOC_LAYOUT,
  notes: DocNoteItem[] = [],
  originalDocxBuffer?: ArrayBuffer | Uint8Array | null,
): Promise<Uint8Array> {
  if (originalDocxBuffer && originalDocxBuffer.byteLength > 0) {
    const lossless = await patchOriginalDocxPackage(html, layout, notes, originalDocxBuffer);
    if (lossless) {
      return lossless.bytes;
    }
  }

  const root = parseHtmlTree(html);
  const children: Array<Paragraph | Table> = [];

  for (const child of root.children) {
    appendBlockNode(child, children, 0);
  }

  if (notes.length > 0) {
    children.push(
      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        keepNext: true,
        widowControl: true,
        children: [new TextRun('Footnotes & Endnotes')],
      }),
    );
    for (const n of notes) {
      children.push(
        new Paragraph({
          widowControl: true,
          children: [
            new TextRun({text: `[${n.number}] (${n.kind}) `, bold: true}),
            new TextRun({text: n.text}),
          ],
        }),
      );
    }
  }

  // Exact margins in twips (prefer imported customMarginsPx when present)
  const presetMargins = DOC_MARGIN_PX[layout.margins] ?? DOC_MARGIN_PX.normal;
  const marginTopTwips = layout.customMarginsPx
    ? pxToTwips(layout.customMarginsPx.top)
    : layout.margins === 'narrow'
      ? 720
      : 1440;
  const marginBottomTwips = layout.customMarginsPx
    ? pxToTwips(layout.customMarginsPx.bottom)
    : layout.margins === 'narrow'
      ? 720
      : 1440;
  const marginLeftTwips = layout.customMarginsPx
    ? pxToTwips(layout.customMarginsPx.left)
    : layout.margins === 'narrow'
      ? 720
      : layout.margins === 'moderate'
        ? 1080
        : layout.margins === 'wide'
          ? 2160
          : pxToTwips(presetMargins.left);
  const marginRightTwips = layout.customMarginsPx
    ? pxToTwips(layout.customMarginsPx.right)
    : marginLeftTwips;

  // Exact page size in twips
  const presetDims = DOC_PAGE_DIMENSIONS_PX[layout.size] ?? DOC_PAGE_DIMENSIONS_PX.a4;
  const rawW = layout.customPageWidthPx
    ? pxToTwips(layout.customPageWidthPx)
    : layout.size === 'a4'
      ? 11906
      : layout.size === 'letter' || layout.size === 'legal'
        ? 12240
        : pxToTwips(presetDims.width);
  const rawH = layout.customPageHeightPx
    ? pxToTwips(layout.customPageHeightPx)
    : layout.size === 'a4'
      ? 16838
      : layout.size === 'legal'
        ? 20160
        : 15840;

  const pageWidthTwips =
    layout.customPageWidthPx != null
      ? rawW
      : layout.orientation === 'landscape'
        ? rawH
        : rawW;
  const pageHeightTwips =
    layout.customPageHeightPx != null
      ? rawH
      : layout.orientation === 'landscape'
        ? rawW
        : rawH;

  // Page borders (w:pgBorders)
  let pageBorders:
    | {
        pageBorderTop?: IBorderOptions;
        pageBorderRight?: IBorderOptions;
        pageBorderBottom?: IBorderOptions;
        pageBorderLeft?: IBorderOptions;
        pageBorders?: {
          display?: (typeof PageBorderDisplay)[keyof typeof PageBorderDisplay];
          offsetFrom?: (typeof PageBorderOffsetFrom)[keyof typeof PageBorderOffsetFrom];
        };
      }
    | undefined;

  if (layout.pageBorderConfig) {
    const cfg = layout.pageBorderConfig;
    pageBorders = {
      pageBorderTop: buildDocxPageBorderSide(cfg.top),
      pageBorderRight: buildDocxPageBorderSide(cfg.right),
      pageBorderBottom: buildDocxPageBorderSide(cfg.bottom),
      pageBorderLeft: buildDocxPageBorderSide(cfg.left),
      pageBorders: {
        display:
          cfg.display === 'firstPage'
            ? PageBorderDisplay.FIRST_PAGE
            : cfg.display === 'notFirstPage'
              ? PageBorderDisplay.NOT_FIRST_PAGE
              : PageBorderDisplay.ALL_PAGES,
        offsetFrom:
          cfg.offsetFrom === 'page' ? PageBorderOffsetFrom.PAGE : PageBorderOffsetFrom.TEXT,
      },
    };
  } else if (layout.pageBorder && layout.pageBorder !== 'none') {
    const bStyle =
      layout.pageBorder === 'double' ? BorderStyle.DOUBLE : BorderStyle.SINGLE;
    const bOpt: IBorderOptions = {style: bStyle, size: 12, color: '000000', space: 24};
    pageBorders = {
      pageBorderTop: bOpt,
      pageBorderRight: bOpt,
      pageBorderBottom: bOpt,
      pageBorderLeft: bOpt,
      pageBorders: {
        display: PageBorderDisplay.ALL_PAGES,
        offsetFrom: PageBorderOffsetFrom.PAGE,
      },
    };
  }

  // Headers & Footers
  const showTopRightPageNum =
    layout.showPageNumbers && layout.pageNumberPosition === 'top-right';
  const showFooterPageNum =
    layout.showPageNumbers && layout.pageNumberPosition !== 'top-right';

  const hasDefaultHeader = Boolean(layout.headerText.trim()) || showTopRightPageNum;
  const hasDefaultFooter = Boolean(layout.footerText.trim()) || showFooterPageNum;

  const defaultHeaderChildren: TextRun[] = [];
  if (layout.headerText.trim()) {
    defaultHeaderChildren.push(
      new TextRun({
        text: showTopRightPageNum ? `${layout.headerText.trim()}    ` : layout.headerText.trim(),
        size: 18,
        color: '64748B',
      }),
    );
  }
  if (showTopRightPageNum) defaultHeaderChildren.push(...buildPageNumberRuns());

  const defaultFooterChildren: TextRun[] = [];
  if (layout.footerText.trim()) {
    defaultFooterChildren.push(
      new TextRun({
        text: showFooterPageNum ? `${layout.footerText.trim()}    ` : layout.footerText.trim(),
        size: 18,
        color: '64748B',
      }),
    );
  }
  if (showFooterPageNum) defaultFooterChildren.push(...buildPageNumberRuns());

  const headersObj =
    hasDefaultHeader || (layout.differentFirstPage && layout.firstPageHeaderText?.trim())
      ? {
          ...(hasDefaultHeader
            ? {
                default: new Header({
                  children: [
                    new Paragraph({
                      alignment:
                        showTopRightPageNum && !layout.headerText.trim()
                          ? AlignmentType.RIGHT
                          : AlignmentType.LEFT,
                      children: defaultHeaderChildren,
                    }),
                  ],
                }),
              }
            : {}),
          ...(layout.differentFirstPage
            ? {
                first: new Header({
                  children: [
                    new Paragraph({
                      children: layout.firstPageHeaderText?.trim()
                        ? [
                            new TextRun({
                              text: layout.firstPageHeaderText.trim(),
                              size: 18,
                              color: '64748B',
                            }),
                          ]
                        : [new TextRun('')],
                    }),
                  ],
                }),
              }
            : {}),
        }
      : undefined;

  const footersObj =
    hasDefaultFooter || (layout.differentFirstPage && layout.firstPageFooterText?.trim())
      ? {
          ...(hasDefaultFooter
            ? {
                default: new Footer({
                  children: [
                    new Paragraph({
                      alignment:
                        layout.pageNumberPosition === 'bottom-right'
                          ? AlignmentType.RIGHT
                          : AlignmentType.CENTER,
                      children: defaultFooterChildren,
                    }),
                  ],
                }),
              }
            : {}),
          ...(layout.differentFirstPage
            ? {
                first: new Footer({
                  children: [
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      children: layout.firstPageFooterText?.trim()
                        ? [
                            new TextRun({
                              text: layout.firstPageFooterText.trim(),
                              size: 18,
                              color: '64748B',
                            }),
                          ]
                        : [new TextRun('')],
                    }),
                  ],
                }),
              }
            : {}),
        }
      : undefined;

  // Canonical Word default font (NEVER a substituted SIL OFL or app UI font!)
  const defaultDocFont =
    canonicalizeWordFontFamily(layout.docDefaultFont) ||
    DOC_THEMES[layout.theme].canonicalFontFamily;
  const defaultDocSizeHalfPts = Math.max(
    12,
    Math.round((layout.docDefaultFontSizePt ?? 11) * 2),
  );
  const defaultDocColor = normalizeHex6(layout.docDefaultColor);

  const doc = new Document({
    creator: 'eSZ OFFICE',
    numbering: {
      config: [
        {
          reference: 'esz-numbered-list',
          levels: [
            {
              level: 0,
              format: LevelFormat.DECIMAL,
              text: '%1.',
              alignment: AlignmentType.LEFT,
              style: {paragraph: {indent: {left: 720, hanging: 360}}},
            },
            {
              level: 1,
              format: LevelFormat.LOWER_LETTER,
              text: '%2.',
              alignment: AlignmentType.LEFT,
              style: {paragraph: {indent: {left: 1440, hanging: 360}}},
            },
            {
              level: 2,
              format: LevelFormat.LOWER_ROMAN,
              text: '%3.',
              alignment: AlignmentType.LEFT,
              style: {paragraph: {indent: {left: 2160, hanging: 360}}},
            },
          ],
        },
      ],
    },
    styles: {
      default: {
        document: {
          run: {
            font: defaultDocFont,
            size: defaultDocSizeHalfPts,
            ...(defaultDocColor ? {color: defaultDocColor} : {}),
          },
        },
      },
    },
    sections: [
      {
        properties: {
          titlePage: Boolean(layout.differentFirstPage),
          column:
            layout.columns > 1
              ? {
                  count: layout.columns,
                  space: 708,
                }
              : undefined,
          page: {
            margin: {
              top: marginTopTwips,
              bottom: marginBottomTwips,
              left: marginLeftTwips,
              right: marginRightTwips,
              header: 540,
              footer: 540,
            },
            size: {
              width: pageWidthTwips,
              height: pageHeightTwips,
              orientation:
                layout.orientation === 'landscape'
                  ? PageOrientation.LANDSCAPE
                  : PageOrientation.PORTRAIT,
            },
            ...(pageBorders ? {borders: pageBorders} : {}),
          },
        },
        headers: headersObj,
        footers: footersObj,
        children: children.length > 0 ? children : [new Paragraph('')],
      },
    ],
  });

  const buffer = await Packer.toArrayBuffer(doc);
  return new Uint8Array(buffer);
}
