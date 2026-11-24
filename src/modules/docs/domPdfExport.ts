import {
  PDFDocument,
  type PDFFont,
  type PDFImage,
  type PDFPage,
  degrees,
  rgb,
} from 'pdf-lib';
import {
  type HtmlNode,
  normalizeHex6,
  parseHtmlTree,
  parseInlineCss,
  parseLengthPt,
} from './docxOoxmlBridge';
import {getPageMetrics} from './docsPagination';
import {
  DEFAULT_DOC_LAYOUT,
  DOC_PAGE_DIMENSIONS_PX,
  type DocNoteItem,
  type DocPageBorderSide,
  type DocPageLayout,
  type DocTabStop,
} from './docsTypes';
import {PdfFontManager, resolveBundledPdfFontFamily} from './pdfFonts';

export interface DomPdfExportOptions {
  title: string;
  layout?: DocPageLayout;
  notes?: DocNoteItem[];
}

interface RgbaColor {
  r: number;
  g: number;
  b: number;
  a: number;
}

interface MappedPage {
  pageIndex: number;
  pdfPage: PDFPage;
  pageWpt: number;
  pageHpt: number;
  rect: DOMRect;
  sx: number;
  sy: number;
}

interface TextLineFragment {
  text: string;
  left: number;
  top: number;
  width: number;
  height: number;
}

const STANDARD_SIZES_PT: Array<[number, number]> = [
  [595.28, 841.89], // A4 portrait
  [841.89, 595.28], // A4 landscape
  [612, 792], // Letter portrait
  [792, 612], // Letter landscape
  [612, 1008], // Legal portrait
  [1008, 612], // Legal landscape
];

function normalizeStandardPageSizePt(widthPt: number, heightPt: number): [number, number] {
  for (const [sw, sh] of STANDARD_SIZES_PT) {
    if (Math.abs(widthPt - sw) <= 3 && Math.abs(heightPt - sh) <= 3) {
      return [sw, sh];
    }
  }
  return [Math.max(72, widthPt), Math.max(72, heightPt)];
}

function getLayoutPageSizePt(layout: DocPageLayout): [number, number] {
  if (
    layout.customPageWidthPx &&
    layout.customPageHeightPx &&
    layout.customPageWidthPx > 100 &&
    layout.customPageHeightPx > 100
  ) {
    const rawW = layout.customPageWidthPx * 0.75;
    const rawH = layout.customPageHeightPx * 0.75;
    const isPortrait = layout.orientation !== 'landscape';
    const w = isPortrait ? Math.min(rawW, rawH) : Math.max(rawW, rawH);
    const h = isPortrait ? Math.max(rawW, rawH) : Math.min(rawW, rawH);
    return normalizeStandardPageSizePt(w, h);
  }
  const base = DOC_PAGE_DIMENSIONS_PX[layout.size] ?? DOC_PAGE_DIMENSIONS_PX.letter;
  const wPt = (layout.orientation === 'landscape' ? base.height : base.width) * 0.75;
  const hPt = (layout.orientation === 'landscape' ? base.width : base.height) * 0.75;
  return normalizeStandardPageSizePt(wPt, hPt);
}

export function parseCssColor(raw: string | null | undefined): RgbaColor | null {
  if (!raw) return null;
  const val = raw.trim().toLowerCase();
  if (!val || val === 'transparent' || val === 'none' || val === 'initial' || val === 'inherit') {
    return null;
  }
  if (val === 'black') return {r: 0, g: 0, b: 0, a: 1};
  if (val === 'white') return {r: 1, g: 1, b: 1, a: 1};

  const rgbaMatch = val.match(
    /^rgba?\(\s*([0-9.]+)\s*[,\s]\s*([0-9.]+)\s*[,\s]\s*([0-9.]+)(?:\s*[,/]\s*([0-9.]+%?))?\s*\)$/,
  );
  if (rgbaMatch) {
    const r = Math.min(255, Math.max(0, Number(rgbaMatch[1]))) / 255;
    const g = Math.min(255, Math.max(0, Number(rgbaMatch[2]))) / 255;
    const b = Math.min(255, Math.max(0, Number(rgbaMatch[3]))) / 255;
    let a = 1;
    if (rgbaMatch[4] !== undefined) {
      a = rgbaMatch[4].endsWith('%')
        ? parseFloat(rgbaMatch[4]) / 100
        : parseFloat(rgbaMatch[4]);
    }
    if (!Number.isFinite(a) || a <= 0.005) return null;
    return {r, g, b, a: Math.min(1, Math.max(0, a))};
  }

  const hex = normalizeHex6(val);
  if (hex) {
    const r = Number.parseInt(hex.slice(0, 2), 16) / 255;
    const g = Number.parseInt(hex.slice(2, 4), 16) / 255;
    const b = Number.parseInt(hex.slice(4, 6), 16) / 255;
    return {r, g, b, a: 1};
  }

  return null;
}

function isNearWhite(color: RgbaColor): boolean {
  return color.r > 0.985 && color.g > 0.985 && color.b > 0.985;
}

function isBoldFontWeight(weight: string | null | undefined): boolean {
  if (!weight) return false;
  const w = weight.trim().toLowerCase();
  if (w === 'bold' || w === 'bolder') return true;
  const num = Number.parseInt(w, 10);
  return Number.isFinite(num) && num >= 600;
}

function isItalicFontStyle(style: string | null | undefined): boolean {
  if (!style) return false;
  const s = style.trim().toLowerCase();
  return s === 'italic' || s === 'oblique';
}

function applyTextTransform(text: string, transform: string | null | undefined): string {
  if (!transform || transform === 'none') return text;
  const t = transform.toLowerCase();
  if (t === 'uppercase') return text.toUpperCase();
  if (t === 'lowercase') return text.toLowerCase();
  if (t === 'capitalize') {
    return text.replace(/\b([a-z])/g, (m) => m.toUpperCase());
  }
  return text;
}

function drawStyledLine(
  page: PDFPage,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  thicknessPt: number,
  style: string,
  color: RgbaColor,
  opacityMultiplier = 1,
): void {
  const thickness = Math.max(0.35, thicknessPt);
  const opacity = Math.min(1, Math.max(0, color.a * opacityMultiplier));
  if (opacity <= 0.005) return;
  const pdfColor = rgb(color.r, color.g, color.b);
  const lowerStyle = (style || 'solid').toLowerCase();

  if (lowerStyle === 'none' || lowerStyle === 'hidden') return;

  if (lowerStyle === 'double' && thickness >= 1.2) {
    const strokeW = Math.max(0.45, thickness / 3);
    const offset = thickness / 3;
    const isHorizontal = Math.abs(y2 - y1) < Math.abs(x2 - x1);
    if (isHorizontal) {
      page.drawLine({
        start: {x: x1, y: y1 + offset},
        end: {x: x2, y: y2 + offset},
        thickness: strokeW,
        color: pdfColor,
        opacity,
      });
      page.drawLine({
        start: {x: x1, y: y1 - offset},
        end: {x: x2, y: y2 - offset},
        thickness: strokeW,
        color: pdfColor,
        opacity,
      });
    } else {
      page.drawLine({
        start: {x: x1 - offset, y: y1},
        end: {x: x2 - offset, y: y2},
        thickness: strokeW,
        color: pdfColor,
        opacity,
      });
      page.drawLine({
        start: {x: x1 + offset, y: y1},
        end: {x: x2 + offset, y: y2},
        thickness: strokeW,
        color: pdfColor,
        opacity,
      });
    }
    return;
  }

  let dashArray: number[] | undefined;
  if (lowerStyle === 'dotted') {
    dashArray = [Math.max(0.8, thickness * 1.1), Math.max(1.8, thickness * 2.0)];
  } else if (lowerStyle === 'dashed') {
    dashArray = [Math.max(3.0, thickness * 3.0), Math.max(2.0, thickness * 2.0)];
  }

  page.drawLine({
    start: {x: x1, y: y1},
    end: {x: x2, y: y2},
    thickness,
    color: pdfColor,
    opacity,
    dashArray,
  });
}

function decodeBase64ToBytes(b64: string): Uint8Array {
  const clean = b64.replace(/\s+/g, '');
  if (typeof Buffer !== 'undefined') {
    const buf = Buffer.from(clean, 'base64');
    return new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength);
  }
  const binary = atob(clean);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

function detectImageFormat(bytes: Uint8Array): 'png' | 'jpg' | null {
  if (
    bytes.length >= 8 &&
    bytes[0] === 0x89 &&
    bytes[1] === 0x50 &&
    bytes[2] === 0x4e &&
    bytes[3] === 0x47
  ) {
    return 'png';
  }
  if (bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) {
    return 'jpg';
  }
  return null;
}

class PdfImageCache {
  private readonly pdf: PDFDocument;
  private readonly cache = new Map<string, PDFImage | null>();

  constructor(pdf: PDFDocument) {
    this.pdf = pdf;
  }

  async embedFromBytes(bytes: Uint8Array, cacheKey?: string): Promise<PDFImage | null> {
    if (cacheKey && this.cache.has(cacheKey)) {
      return this.cache.get(cacheKey) ?? null;
    }
    const fmt = detectImageFormat(bytes);
    let img: PDFImage | null = null;
    try {
      if (fmt === 'png') {
        img = await this.pdf.embedPng(bytes);
      } else if (fmt === 'jpg') {
        img = await this.pdf.embedJpg(bytes);
      }
    } catch {
      img = null;
    }
    if (cacheKey) this.cache.set(cacheKey, img);
    return img;
  }

  async embedFromDataUrl(src: string): Promise<PDFImage | null> {
    if (this.cache.has(src)) return this.cache.get(src) ?? null;
    const match = src.match(/^data:image\/([a-zA-Z0-9.+-]+)(?:;charset=[^;,]+)?;base64,([\s\S]+)$/i);
    if (!match) {
      this.cache.set(src, null);
      return null;
    }
    const bytes = decodeBase64ToBytes(match[2]);
    return this.embedFromBytes(bytes, src);
  }

  async embedHtmlImage(imgEl: HTMLImageElement, targetWidthPx: number, targetHeightPx: number): Promise<PDFImage | null> {
    const src = imgEl.currentSrc || imgEl.getAttribute('src') || '';
    if (!src) return null;
    if (this.cache.has(src)) return this.cache.get(src) ?? null;

    // 1. Direct lossless PNG/JPEG embedding from data URL
    if (src.startsWith('data:image/')) {
      const direct = await this.embedFromDataUrl(src);
      if (direct) return direct;
    }

    // 2. Direct lossless PNG/JPEG fetch from blob: or same-origin URL
    if (typeof window !== 'undefined' && typeof window.fetch === 'function' && (src.startsWith('blob:') || src.startsWith('http') || src.startsWith('/'))) {
      try {
        const res = await window.fetch(src);
        if (res.ok) {
          const bytes = new Uint8Array(await res.arrayBuffer());
          const direct = await this.embedFromBytes(bytes, src);
          if (direct) return direct;
        }
      } catch {
        // Fall through to canvas rasterization
      }
    }

    // 3. Rasterize WebP/GIF/SVG/etc. onto an offscreen high-DPI canvas -> PNG
    if (typeof document !== 'undefined') {
      try {
        const w = Math.max(
          1,
          Math.min(2400, Math.round(imgEl.naturalWidth || Math.max(2, targetWidthPx * 2))),
        );
        const h = Math.max(
          1,
          Math.min(2400, Math.round(imgEl.naturalHeight || Math.max(2, targetHeightPx * 2))),
        );
        const canvas = document.createElement('canvas');
        canvas.width = w;
        canvas.height = h;
        const ctx = canvas.getContext('2d');
        if (ctx) {
          ctx.drawImage(imgEl, 0, 0, w, h);
          const pngUrl = canvas.toDataURL('image/png');
          const match = pngUrl.match(/^data:image\/png;base64,([\s\S]+)$/);
          if (match) {
            const embedded = await this.embedFromBytes(decodeBase64ToBytes(match[1]), src);
            return embedded;
          }
        }
      } catch {
        // Ignore tainted canvas
      }
    }

    this.cache.set(src, null);
    return null;
  }

  async embedCanvas(canvas: HTMLCanvasElement): Promise<PDFImage | null> {
    try {
      const dataUrl = canvas.toDataURL('image/png');
      return await this.embedFromDataUrl(dataUrl);
    } catch {
      return null;
    }
  }

  async embedSvg(svgEl: SVGSVGElement, widthPx: number, heightPx: number): Promise<PDFImage | null> {
    if (typeof window === 'undefined' || typeof document === 'undefined') return null;
    try {
      const clone = svgEl.cloneNode(true) as SVGSVGElement;
      if (!clone.getAttribute('xmlns')) {
        clone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
      }
      const w = Math.max(1, Math.min(2000, Math.round(widthPx * 2)));
      const h = Math.max(1, Math.min(2000, Math.round(heightPx * 2)));
      clone.setAttribute('width', String(w));
      clone.setAttribute('height', String(h));
      const xml = new XMLSerializer().serializeToString(clone);
      const svgUrl = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(xml)}`;

      const img = await new Promise<HTMLImageElement | null>((resolve) => {
        const image = new Image();
        image.onload = () => resolve(image);
        image.onerror = () => resolve(null);
        image.src = svgUrl;
      });
      if (!img) return null;

      const canvas = document.createElement('canvas');
      canvas.width = w;
      canvas.height = h;
      const ctx = canvas.getContext('2d');
      if (!ctx) return null;
      ctx.drawImage(img, 0, 0, w, h);
      return await this.embedCanvas(canvas);
    } catch {
      return null;
    }
  }
}

// ---------------------------------------------------------------------------
// DOM Walking Helpers (Browser)
// ---------------------------------------------------------------------------

function shouldSkipElement(el: Element): boolean {
  const tag = el.tagName.toLowerCase();
  if (
    tag === 'style' ||
    tag === 'script' ||
    tag === 'noscript' ||
    tag === 'template' ||
    tag === 'button'
  ) {
    return true;
  }
  if (el.getAttribute('data-page-break') === 'true') return true;
  if (el.getAttribute('data-testid') === 'docs-page-watermark') return true;
  if (
    el.classList.contains('esz-page-break-spacer') ||
    el.classList.contains('esz-page-spacer-inline') ||
    el.classList.contains('ProseMirror-trailingBreak') ||
    el.classList.contains('ProseMirror-separator') ||
    el.classList.contains('ProseMirror-gapcursor') ||
    el.classList.contains('column-resize-handle')
  ) {
    return true;
  }
  const style = window.getComputedStyle(el);
  if (style.display === 'none' || style.visibility === 'hidden') return true;
  if (parseFloat(style.opacity || '1') <= 0.005) return true;
  return false;
}

function getEffectiveTextDecorations(el: Element): {
  underline: boolean;
  strike: boolean;
  decorationColor: RgbaColor | null;
} {
  let underline = false;
  let strike = false;
  let decorationColor: RgbaColor | null = null;
  let cur: Element | null = el;

  while (cur) {
    const cs = window.getComputedStyle(cur);
    const line = (cs.textDecorationLine || cs.textDecoration || '').toLowerCase();
    if (line.includes('underline')) {
      underline = true;
      if (!decorationColor) {
        decorationColor = parseCssColor(cs.textDecorationColor) ?? parseCssColor(cs.color);
      }
    }
    if (line.includes('line-through')) {
      strike = true;
      if (!decorationColor) {
        decorationColor = parseCssColor(cs.textDecorationColor) ?? parseCssColor(cs.color);
      }
    }
    const display = cs.display;
    if (
      display === 'block' ||
      display === 'list-item' ||
      display === 'table-cell' ||
      display === 'flex' ||
      display === 'grid'
    ) {
      break;
    }
    cur = cur.parentElement;
  }

  return {underline, strike, decorationColor};
}

function getEnclosingBlockTextAlign(el: Element): string {
  let cur: Element | null = el;
  while (cur) {
    const cs = window.getComputedStyle(cur);
    const display = cs.display;
    if (
      display === 'block' ||
      display === 'list-item' ||
      display === 'table-cell' ||
      display === 'flex' ||
      display === 'grid'
    ) {
      return (cs.textAlign || 'left').toLowerCase();
    }
    cur = cur.parentElement;
  }
  return 'left';
}

function collapseTextWhitespace(raw: string, whiteSpace: string): string {
  const ws = (whiteSpace || 'normal').toLowerCase();
  if (ws === 'pre' || ws === 'pre-wrap' || ws === 'break-spaces') {
    return raw.replace(/\r/g, '');
  }
  return raw.replace(/[\r\n\t ]+/g, ' ');
}

/**
 * Extracts exact visual line fragments (and justified word positions) from a DOM Text node
 * using Range.getClientRects().
 */
function extractTextLineFragments(
  textNode: Text,
  parentStyle: CSSStyleDeclaration,
  isJustified: boolean,
  scaleX: number,
  fontSizePx: number,
): TextLineFragment[] {
  const raw = textNode.nodeValue ?? '';
  if (!raw) return [];

  const whiteSpace = (parentStyle.whiteSpace || 'normal').toLowerCase();
  const preserveWs = whiteSpace === 'pre' || whiteSpace === 'pre-wrap' || whiteSpace === 'break-spaces';
  if (!preserveWs && !raw.trim()) return [];

  const doc = textNode.ownerDocument;
  if (!doc) return [];

  const range = doc.createRange();
  range.selectNodeContents(textNode);
  const wholeRects = Array.from(range.getClientRects()).filter(
    (r) => r.width > 0.15 && r.height > 0.15,
  );
  if (wholeRects.length === 0) return [];

  const textTransform = parentStyle.textTransform;

  // Fast path: single visual rect, not justified, no internal newlines
  if (wholeRects.length === 1 && !isJustified && !raw.includes('\n')) {
    const normalized = applyTextTransform(collapseTextWhitespace(raw, whiteSpace), textTransform);
    const trimmed = preserveWs ? normalized : normalized.trim();
    if (!trimmed) return [];
    const r = wholeRects[0];
    // Measure trimmed slice when leading/trailing whitespace existed in the text node
    if (!preserveWs && (raw.length !== raw.trim().length)) {
      const leadMatch = raw.match(/^\s*/);
      const trailMatch = raw.match(/\s*$/);
      const startIdx = leadMatch ? leadMatch[0].length : 0;
      const endIdx = raw.length - (trailMatch ? trailMatch[0].length : 0);
      if (endIdx > startIdx) {
        range.setStart(textNode, startIdx);
        range.setEnd(textNode, endIdx);
        const tightRects = Array.from(range.getClientRects()).filter(
          (tr) => tr.width > 0.15 && tr.height > 0.15,
        );
        if (tightRects.length === 1) {
          return [
            {
              text: trimmed,
              left: tightRects[0].left,
              top: tightRects[0].top,
              width: tightRects[0].width,
              height: tightRects[0].height,
            },
          ];
        }
      }
    }
    return [
      {
        text: trimmed,
        left: r.left,
        top: r.top,
        width: r.width,
        height: r.height,
      },
    ];
  }

  // Multi-line OR justified OR preformatted text: measure word-by-word via Range.getClientRects()
  interface WordPiece {
    text: string;
    left: number;
    top: number;
    right: number;
    width: number;
    height: number;
  }

  const pieces: WordPiece[] = [];
  const tokenRegex = /\S+/g;
  let match: RegExpExecArray | null;

  while ((match = tokenRegex.exec(raw)) !== null) {
    const start = match.index;
    const end = start + match[0].length;
    range.setStart(textNode, start);
    range.setEnd(textNode, end);
    const wRects = Array.from(range.getClientRects()).filter(
      (r) => r.width > 0.1 && r.height > 0.1,
    );
    if (wRects.length === 0) continue;

    if (wRects.length === 1) {
      const r = wRects[0];
      pieces.push({
        text: applyTextTransform(match[0], textTransform),
        left: r.left,
        top: r.top,
        right: r.right,
        width: r.width,
        height: r.height,
      });
    } else {
      // Single long word wrapped mid-word: split per visual line character-by-character
      let segStart = start;
      let segTop: number | null = null;
      let segRect: DOMRect | null = null;

      for (let i = start; i < end; i++) {
        range.setStart(textNode, i);
        range.setEnd(textNode, i + 1);
        const cRect = range.getClientRects()[0];
        if (!cRect || cRect.height <= 0.1) continue;
        if (segTop === null) {
          segTop = cRect.top;
          segRect = cRect;
          segStart = i;
        } else if (Math.abs(cRect.top - segTop) > Math.max(3, cRect.height * 0.45)) {
          range.setStart(textNode, segStart);
          range.setEnd(textNode, i);
          const sliceRect = range.getClientRects()[0] ?? segRect!;
          pieces.push({
            text: applyTextTransform(raw.slice(segStart, i), textTransform),
            left: sliceRect.left,
            top: sliceRect.top,
            right: sliceRect.right,
            width: sliceRect.width,
            height: sliceRect.height,
          });
          segStart = i;
          segTop = cRect.top;
          segRect = cRect;
        }
      }
      if (segStart < end) {
        range.setStart(textNode, segStart);
        range.setEnd(textNode, end);
        const sliceRect = range.getClientRects()[0] ?? segRect;
        if (sliceRect) {
          pieces.push({
            text: applyTextTransform(raw.slice(segStart, end), textTransform),
            left: sliceRect.left,
            top: sliceRect.top,
            right: sliceRect.right,
            width: sliceRect.width,
            height: sliceRect.height,
          });
        }
      }
    }
  }

  if (pieces.length === 0) return [];

  // Group pieces into visual lines
  const lines: WordPiece[][] = [];
  for (const p of pieces) {
    const lastLine = lines[lines.length - 1];
    if (
      lastLine &&
      Math.abs(p.top - lastLine[0].top) <= Math.max(3, lastLine[0].height * 0.45) &&
      p.left >= lastLine[lastLine.length - 1].left - 2
    ) {
      lastLine.push(p);
    } else {
      lines.push([p]);
    }
  }

  const fragments: TextLineFragment[] = [];
  const naturalSpacePx = Math.max(2, fontSizePx * 0.28) / Math.max(0.1, scaleX > 0 ? 1 / (scaleX / 0.75) : 1);
  const largeGapThresholdPx = Math.max(8, fontSizePx * 1.15);

  for (const line of lines) {
    if (line.length === 1) {
      const p = line[0];
      fragments.push({
        text: p.text,
        left: p.left,
        top: p.top,
        width: p.width,
        height: p.height,
      });
      continue;
    }

    // Detect if this line is stretched by CSS justification OR contains wide tab gaps
    let hasStretchedJustification = false;
    if (isJustified) {
      for (let i = 1; i < line.length; i++) {
        const gap = line[i].left - line[i - 1].right;
        if (gap > naturalSpacePx * 1.45) {
          hasStretchedJustification = true;
          break;
        }
      }
    }

    if (hasStretchedJustification) {
      // Emit each word at its exact browser-justified X coordinate
      for (const p of line) {
        fragments.push({
          text: p.text,
          left: p.left,
          top: p.top,
          width: p.width,
          height: p.height,
        });
      }
      continue;
    }

    // Merge contiguous words on the same line, splitting only on wide tab-like horizontal jumps
    let runStart = line[0];
    let runText = line[0].text;
    let runRight = line[0].right;
    let runTop = line[0].top;
    let runHeight = line[0].height;

    for (let i = 1; i < line.length; i++) {
      const p = line[i];
      const gap = p.left - runRight;
      if (gap > largeGapThresholdPx) {
        fragments.push({
          text: runText,
          left: runStart.left,
          top: runTop,
          width: Math.max(0.5, runRight - runStart.left),
          height: runHeight,
        });
        runStart = p;
        runText = p.text;
        runRight = p.right;
        runTop = p.top;
        runHeight = p.height;
      } else {
        runText += ` ${p.text}`;
        runRight = Math.max(runRight, p.right);
        runTop = Math.min(runTop, p.top);
        runHeight = Math.max(runHeight, p.height);
      }
    }
    fragments.push({
      text: runText,
      left: runStart.left,
      top: runTop,
      width: Math.max(0.5, runRight - runStart.left),
      height: runHeight,
    });
  }

  return fragments;
}

function getListMarkerText(li: HTMLElement, listStyleType: string): string | null {
  const type = listStyleType.toLowerCase();
  if (type === 'none') return null;
  if (type === 'disc') return '\u2022';
  if (type === 'circle') return '\u25E6';
  if (type === 'square') return '\u25AA';

  const parent = li.parentElement;
  if (!parent) return '\u2022';
  const startAttr = Number.parseInt(parent.getAttribute('start') || '1', 10);
  const start = Number.isFinite(startAttr) ? startAttr : 1;
  let index = 0;
  for (const child of Array.from(parent.children)) {
    if (child === li) break;
    if (child.tagName.toLowerCase() === 'li') index += 1;
  }
  const num = start + index;

  if (type === 'lower-alpha' || type === 'lower-latin') {
    return `${String.fromCharCode(97 + ((num - 1) % 26))}.`;
  }
  if (type === 'upper-alpha' || type === 'upper-latin') {
    return `${String.fromCharCode(65 + ((num - 1) % 26))}.`;
  }
  if (type === 'lower-roman') return `${toRoman(num).toLowerCase()}.`;
  if (type === 'upper-roman') return `${toRoman(num)}.`;
  return `${num}.`;
}

function toRoman(num: number): string {
  const pairs: Array<[number, string]> = [
    [1000, 'M'],
    [900, 'CM'],
    [500, 'D'],
    [400, 'CD'],
    [100, 'C'],
    [90, 'XC'],
    [50, 'L'],
    [40, 'XL'],
    [10, 'X'],
    [9, 'IX'],
    [5, 'V'],
    [4, 'IV'],
    [1, 'I'],
  ];
  let n = Math.max(1, Math.floor(num));
  let out = '';
  for (const [val, sym] of pairs) {
    while (n >= val) {
      out += sym;
      n -= val;
    }
  }
  return out;
}

function drawElementBoxBackgroundAndBorders(
  el: HTMLElement,
  mapped: MappedPage,
  rect: DOMRect,
  skipWhiteBg = true,
  skipBorders = false,
): void {
  const cs = window.getComputedStyle(el);
  const opacity = Math.min(1, Math.max(0, parseFloat(cs.opacity || '1')));
  if (opacity <= 0.005) return;

  const xPt = (rect.left - mapped.rect.left) * mapped.sx;
  const yTopPt = (rect.top - mapped.rect.top) * mapped.sy;
  const wPt = rect.width * mapped.sx;
  const hPt = rect.height * mapped.sy;
  if (wPt <= 0.1 || hPt <= 0.1) return;

  const yBottomPt = mapped.pageHpt - (yTopPt + hPt);
  const yTopPdf = mapped.pageHpt - yTopPt;

  // 1. Background fill
  const bg = parseCssColor(cs.backgroundColor);
  if (bg && (!skipWhiteBg || !isNearWhite(bg))) {
    mapped.pdfPage.drawRectangle({
      x: xPt,
      y: yBottomPt,
      width: wPt,
      height: hPt,
      color: rgb(bg.r, bg.g, bg.b),
      opacity: bg.a * opacity,
    });
  }

  if (skipBorders) return;

  // 2. Per-side borders (solid, dotted, dashed, double)
  const btW = parseFloat(cs.borderTopWidth || '0') * 0.75;
  const brW = parseFloat(cs.borderRightWidth || '0') * 0.75;
  const bbW = parseFloat(cs.borderBottomWidth || '0') * 0.75;
  const blW = parseFloat(cs.borderLeftWidth || '0') * 0.75;

  if (btW > 0.05 && cs.borderTopStyle !== 'none' && cs.borderTopStyle !== 'hidden') {
    const col = parseCssColor(cs.borderTopColor);
    if (col) {
      drawStyledLine(
        mapped.pdfPage,
        xPt,
        yTopPdf - btW / 2,
        xPt + wPt,
        yTopPdf - btW / 2,
        btW,
        cs.borderTopStyle,
        col,
        opacity,
      );
    }
  }
  if (bbW > 0.05 && cs.borderBottomStyle !== 'none' && cs.borderBottomStyle !== 'hidden') {
    const col = parseCssColor(cs.borderBottomColor);
    if (col) {
      drawStyledLine(
        mapped.pdfPage,
        xPt,
        yBottomPt + bbW / 2,
        xPt + wPt,
        yBottomPt + bbW / 2,
        bbW,
        cs.borderBottomStyle,
        col,
        opacity,
      );
    }
  }
  if (blW > 0.05 && cs.borderLeftStyle !== 'none' && cs.borderLeftStyle !== 'hidden') {
    const col = parseCssColor(cs.borderLeftColor);
    if (col) {
      drawStyledLine(
        mapped.pdfPage,
        xPt + blW / 2,
        yBottomPt,
        xPt + blW / 2,
        yTopPdf,
        blW,
        cs.borderLeftStyle,
        col,
        opacity,
      );
    }
  }
  if (brW > 0.05 && cs.borderRightStyle !== 'none' && cs.borderRightStyle !== 'hidden') {
    const col = parseCssColor(cs.borderRightColor);
    if (col) {
      drawStyledLine(
        mapped.pdfPage,
        xPt + wPt - brW / 2,
        yBottomPt,
        xPt + wPt - brW / 2,
        yTopPdf,
        brW,
        cs.borderRightStyle,
        col,
        opacity,
      );
    }
  }
}

function drawTabLeaderInDom(el: HTMLElement, mapped: MappedPage, rect: DOMRect): void {
  const leader = (el.getAttribute('data-tab-leader') || 'none').toLowerCase();
  if (!leader || leader === 'none') return;
  const cs = window.getComputedStyle(el);
  const col = parseCssColor(cs.color) ?? {r: 0, g: 0, b: 0, a: 1};
  const opacity = Math.min(1, Math.max(0, parseFloat(cs.opacity || '1')));
  const xPt = (rect.left - mapped.rect.left) * mapped.sx;
  const wPt = rect.width * mapped.sx;
  const yTopPt = (rect.top - mapped.rect.top) * mapped.sy;
  const hPt = rect.height * mapped.sy;
  if (wPt <= 2) return;

  const isMiddle = leader === 'middledot' || leader === 'hyphen';
  const lineY = mapped.pageHpt - (yTopPt + hPt * (isMiddle ? 0.55 : 0.82));
  const style =
    leader === 'dot' || leader === 'middledot'
      ? 'dotted'
      : leader === 'hyphen'
        ? 'dashed'
        : 'solid';
  const thickness = leader === 'heavy' ? 1.5 : 0.85;
  drawStyledLine(
    mapped.pdfPage,
    xPt + 1,
    lineY,
    xPt + wPt - 1,
    lineY,
    thickness,
    style,
    col,
    opacity,
  );
}

function resolveMappedPageForRect(rect: DOMRect, pages: MappedPage[]): MappedPage | null {
  if (pages.length === 0) return null;
  if (pages.length === 1) return pages[0];
  const centerY = (rect.top + rect.bottom) / 2;
  for (const p of pages) {
    if (centerY >= p.rect.top - 2 && centerY <= p.rect.bottom + 2) {
      return p;
    }
  }
  // Nearest page fallback
  let best = pages[0];
  let bestDist = Infinity;
  for (const p of pages) {
    const dist =
      centerY < p.rect.top
        ? p.rect.top - centerY
        : centerY > p.rect.bottom
          ? centerY - p.rect.bottom
          : 0;
    if (dist < bestDist) {
      bestDist = dist;
      best = p;
    }
  }
  return best;
}

async function drawWatermarkOnPage(
  pdfPage: PDFPage,
  pageWpt: number,
  pageHpt: number,
  watermarkText: string,
  fontManager: PdfFontManager,
): Promise<void> {
  const trimmed = watermarkText.trim();
  if (!trimmed) return;
  const font = await fontManager.getFont('Liberation Sans', true, false);
  const safe = fontManager.sanitizeTextForFont(trimmed.toUpperCase(), font);
  if (!safe) return;

  const fontSize = 48;
  const textW = font.widthOfTextAtSize(safe, fontSize);
  const angleRad = Math.PI / 4;
  // Center the rotated watermark string geometrically on the page
  const cx = pageWpt / 2;
  const cy = pageHpt / 2;
  const x = cx - (textW / 2) * Math.cos(angleRad);
  const y = cy - (textW / 2) * Math.sin(angleRad) - fontSize * 0.25;

  pdfPage.drawText(safe, {
    x,
    y,
    size: fontSize,
    font,
    color: rgb(0.58, 0.64, 0.72),
    opacity: 0.16,
    rotate: degrees(45),
  });
}

function drawPageBorderOverlay(
  pdfPage: PDFPage,
  pageWpt: number,
  pageHpt: number,
  layout: DocPageLayout,
  pageIndex: number,
): void {
  const cfg = layout.pageBorderConfig;
  if (cfg) {
    if (cfg.display === 'firstPage' && pageIndex !== 0) return;
    if (cfg.display === 'notFirstPage' && pageIndex === 0) return;

    const pm = getPageMetrics(layout);
    const drawSide = (
      side: DocPageBorderSide | undefined,
      orientation: 'top' | 'right' | 'bottom' | 'left',
    ) => {
      if (!side || side.val === 'none' || side.val === 'nil') return;
      const color = parseCssColor(side.color) ?? {r: 0, g: 0, b: 0, a: 1};
      const thickness = Math.max(0.5, side.szPt);
      const cssStyle =
        side.val === 'double'
          ? 'double'
          : side.val === 'dotted'
            ? 'dotted'
            : /dash/i.test(side.val)
              ? 'dashed'
              : 'solid';
      const insetFromMargin = cfg.offsetFrom === 'text';
      const spacePt = side.spacePt;

      const leftX = insetFromMargin
        ? Math.max(4, pm.marginLeft * 0.75 - spacePt)
        : Math.max(4, spacePt);
      const rightX = insetFromMargin
        ? Math.min(pageWpt - 4, pageWpt - pm.marginRight * 0.75 + spacePt)
        : Math.min(pageWpt - 4, pageWpt - spacePt);
      const topY = insetFromMargin
        ? Math.min(pageHpt - 4, pageHpt - pm.marginTop * 0.75 + spacePt)
        : Math.min(pageHpt - 4, pageHpt - spacePt);
      const bottomY = insetFromMargin
        ? Math.max(4, pm.marginBottom * 0.75 - spacePt)
        : Math.max(4, spacePt);

      if (orientation === 'top') {
        drawStyledLine(pdfPage, leftX, topY, rightX, topY, thickness, cssStyle, color);
      } else if (orientation === 'bottom') {
        drawStyledLine(pdfPage, leftX, bottomY, rightX, bottomY, thickness, cssStyle, color);
      } else if (orientation === 'left') {
        drawStyledLine(pdfPage, leftX, bottomY, leftX, topY, thickness, cssStyle, color);
      } else {
        drawStyledLine(pdfPage, rightX, bottomY, rightX, topY, thickness, cssStyle, color);
      }
    };

    drawSide(cfg.top, 'top');
    drawSide(cfg.right, 'right');
    drawSide(cfg.bottom, 'bottom');
    drawSide(cfg.left, 'left');
    return;
  }

  if (layout.pageBorder === 'none') return;
  const inset = 18;
  const col: RgbaColor = {r: 0.12, g: 0.16, b: 0.23, a: 1};
  const style = layout.pageBorder === 'double' ? 'double' : 'solid';
  const thickness = layout.pageBorder === 'double' ? 2.5 : 1.2;
  drawStyledLine(pdfPage, inset, pageHpt - inset, pageWpt - inset, pageHpt - inset, thickness, style, col);
  drawStyledLine(pdfPage, inset, inset, pageWpt - inset, inset, thickness, style, col);
  drawStyledLine(pdfPage, inset, inset, inset, pageHpt - inset, thickness, style, col);
  drawStyledLine(pdfPage, pageWpt - inset, inset, pageWpt - inset, pageHpt - inset, thickness, style, col);
}

/**
 * Walks a DOM subtree and renders backgrounds, borders, tab leaders, images, SVGs,
 * list markers, and Range.getClientRects()-positioned text runs onto the mapped PDF page(s).
 */
async function renderDomSubtreeToPdfPages(
  root: HTMLElement,
  pages: MappedPage[],
  fontManager: PdfFontManager,
  imageCache: PdfImageCache,
  options: {
    skipRootBox?: boolean;
    skipTableOuterBoxWhenMultiPage?: boolean;
  } = {},
): Promise<void> {
  // Collect nodes in document order so we can paint cleanly in 3 layers:
  // 1) Backgrounds, borders & tab leaders
  // 2) Images, canvases & inline SVGs
  // 3) Text runs, list markers & editable footnote inputs
  const elements: HTMLElement[] = [];
  const svgs: SVGSVGElement[] = [];
  const textNodes: Text[] = [];
  const inputNodes: HTMLInputElement[] = [];

  const traverse = (node: Node) => {
    if (node.nodeType === Node.ELEMENT_NODE) {
      const el = node as Element;
      if (shouldSkipElement(el)) return;

      if (el instanceof SVGSVGElement) {
        svgs.push(el);
        return;
      }

      if (el instanceof HTMLInputElement) {
        inputNodes.push(el);
        return;
      }

      if (el instanceof HTMLElement) {
        elements.push(el);
        // Do not descend into tab node's dummy "\t" text child
        if (el.getAttribute('data-doc-tab') === 'true') {
          return;
        }
      }
      for (const child of Array.from(el.childNodes)) {
        traverse(child);
      }
    } else if (node.nodeType === Node.TEXT_NODE) {
      textNodes.push(node as Text);
    }
  };

  traverse(root);

  // Layer 1: Element backgrounds, borders & tab leaders
  for (const el of elements) {
    if (options.skipRootBox && el === root) continue;
    if (el.getAttribute('data-doc-tab') === 'true') {
      const rect = el.getBoundingClientRect();
      const mapped = resolveMappedPageForRect(rect, pages);
      if (mapped) drawTabLeaderInDom(el, mapped, rect);
      continue;
    }

    const tag = el.tagName.toLowerCase();
    if (tag === 'img' || tag === 'canvas') continue;

    const cs = window.getComputedStyle(el);
    const isInline = cs.display === 'inline';

    if (isInline) {
      // Inline text highlight (<mark>, <span style="background-color: ...">) uses Range.getClientRects()
      const bg = parseCssColor(cs.backgroundColor);
      if (bg && !isNearWhite(bg)) {
        const range = el.ownerDocument.createRange();
        range.selectNodeContents(el);
        const rects = Array.from(range.getClientRects()).filter(
          (r) => r.width > 0.2 && r.height > 0.2,
        );
        const opacity = Math.min(1, Math.max(0, parseFloat(cs.opacity || '1')));
        for (const r of rects) {
          const mapped = resolveMappedPageForRect(r, pages);
          if (!mapped) continue;
          const xPt = (r.left - mapped.rect.left) * mapped.sx;
          const yTopPt = (r.top - mapped.rect.top) * mapped.sy;
          const wPt = r.width * mapped.sx;
          const hPt = r.height * mapped.sy;
          mapped.pdfPage.drawRectangle({
            x: xPt,
            y: mapped.pageHpt - (yTopPt + hPt),
            width: wPt,
            height: hPt,
            color: rgb(bg.r, bg.g, bg.b),
            opacity: bg.a * opacity,
          });
        }
      }
      continue;
    }

    const rect = el.getBoundingClientRect();
    if (rect.width <= 0.2 || rect.height <= 0.2) continue;

    // Skip drawing a single outer border for a multi-page spanning <table> or wrapper div;
    // its individual <td>/<th> cells carry their own borders on each page.
    if (
      options.skipTableOuterBoxWhenMultiPage &&
      (tag === 'table' || tag === 'tbody' || tag === 'thead') &&
      pages.length > 1
    ) {
      const firstPage = resolveMappedPageForRect(
        new DOMRect(rect.left, rect.top, rect.width, 1),
        pages,
      );
      const lastPage = resolveMappedPageForRect(
        new DOMRect(rect.left, rect.bottom - 1, rect.width, 1),
        pages,
      );
      if (firstPage && lastPage && firstPage.pageIndex !== lastPage.pageIndex) {
        continue;
      }
    }

    const mapped = resolveMappedPageForRect(rect, pages);
    if (!mapped) continue;
    drawElementBoxBackgroundAndBorders(el, mapped, rect, true, false);
  }

  // Layer 2: Images, Canvases, and Inline SVGs
  for (const el of elements) {
    const tag = el.tagName.toLowerCase();
    if (tag === 'img') {
      const imgEl = el as HTMLImageElement;
      const rect = imgEl.getBoundingClientRect();
      if (rect.width <= 0.5 || rect.height <= 0.5) continue;
      const mapped = resolveMappedPageForRect(rect, pages);
      if (!mapped) continue;

      const pdfImg = await imageCache.embedHtmlImage(imgEl, rect.width, rect.height);
      if (!pdfImg) continue;

      const cs = window.getComputedStyle(imgEl);
      const opacity = Math.min(1, Math.max(0, parseFloat(cs.opacity || '1')));
      const xPt = (rect.left - mapped.rect.left) * mapped.sx;
      const yTopPt = (rect.top - mapped.rect.top) * mapped.sy;
      const wPt = rect.width * mapped.sx;
      const hPt = rect.height * mapped.sy;

      mapped.pdfPage.drawImage(pdfImg, {
        x: xPt,
        y: mapped.pageHpt - (yTopPt + hPt),
        width: wPt,
        height: hPt,
        opacity,
      });

      // Also draw image border if styled
      drawElementBoxBackgroundAndBorders(imgEl, mapped, rect, true, false);
    } else if (tag === 'canvas') {
      const canvasEl = el as HTMLCanvasElement;
      const rect = canvasEl.getBoundingClientRect();
      if (rect.width <= 0.5 || rect.height <= 0.5) continue;
      const mapped = resolveMappedPageForRect(rect, pages);
      if (!mapped) continue;

      const pdfImg = await imageCache.embedCanvas(canvasEl);
      if (!pdfImg) continue;

      const xPt = (rect.left - mapped.rect.left) * mapped.sx;
      const yTopPt = (rect.top - mapped.rect.top) * mapped.sy;
      const wPt = rect.width * mapped.sx;
      const hPt = rect.height * mapped.sy;
      mapped.pdfPage.drawImage(pdfImg, {
        x: xPt,
        y: mapped.pageHpt - (yTopPt + hPt),
        width: wPt,
        height: hPt,
      });
    }
  }

  for (const svgEl of svgs) {
    const rect = svgEl.getBoundingClientRect();
    if (rect.width <= 1 || rect.height <= 1) continue;
    const mapped = resolveMappedPageForRect(rect, pages);
    if (!mapped) continue;
    const pdfImg = await imageCache.embedSvg(svgEl, rect.width, rect.height);
    if (!pdfImg) continue;
    const xPt = (rect.left - mapped.rect.left) * mapped.sx;
    const yTopPt = (rect.top - mapped.rect.top) * mapped.sy;
    const wPt = rect.width * mapped.sx;
    const hPt = rect.height * mapped.sy;
    mapped.pdfPage.drawImage(pdfImg, {
      x: xPt,
      y: mapped.pageHpt - (yTopPt + hPt),
      width: wPt,
      height: hPt,
    });
  }

  // Layer 3A: List markers (<li> bullets & ordered numbers rendered via CSS ::marker)
  for (const el of elements) {
    if (el.tagName.toLowerCase() !== 'li') continue;
    const cs = window.getComputedStyle(el);
    const marker = getListMarkerText(el, cs.listStyleType || 'none');
    if (!marker) continue;

    // Anchor marker to the first text rect inside the <li> (or the <li> top-left)
    const liRect = el.getBoundingClientRect();
    if (liRect.width <= 0.5 || liRect.height <= 0.5) continue;
    const mapped = resolveMappedPageForRect(liRect, pages);
    if (!mapped) continue;

    const fontSizePx = parseFloat(cs.fontSize || '14.67') || 14.67;
    const fontSizePt = Math.max(4, fontSizePx * 0.75);
    const bold = isBoldFontWeight(cs.fontWeight);
    const italic = isItalicFontStyle(cs.fontStyle);
    const font = await fontManager.getFont(cs.fontFamily, bold, italic);
    const safeMarker = fontManager.sanitizeTextForFont(marker, font);
    if (!safeMarker) continue;

    const color = parseCssColor(cs.color) ?? {r: 0, g: 0, b: 0, a: 1};
    const opacity = Math.min(1, Math.max(0, parseFloat(cs.opacity || '1'))) * color.a;
    const markerW = font.widthOfTextAtSize(safeMarker, fontSizePt);
    const xPt = (liRect.left - mapped.rect.left) * mapped.sx - markerW - 5;
    const yTopPt = (liRect.top - mapped.rect.top) * mapped.sy;
    const ascentFrac = fontManager.getAscentFraction(font, fontSizePt);
    const lineH = Math.min(liRect.height * mapped.sy, fontSizePt * 1.25);
    const baselineY = mapped.pageHpt - (yTopPt + lineH * ascentFrac);

    mapped.pdfPage.drawText(safeMarker, {
      x: Math.max(4, xPt),
      y: baselineY,
      size: fontSizePt,
      font,
      color: rgb(color.r, color.g, color.b),
      opacity,
    });
  }

  // Layer 3B: All DOM Text Nodes via Range.getClientRects()
  for (const textNode of textNodes) {
    const parent = textNode.parentElement;
    if (!parent) continue;
    if (shouldSkipElement(parent)) continue;

    const cs = window.getComputedStyle(parent);
    const fontSizePx = parseFloat(cs.fontSize || '14.67') || 14.67;
    const fontSizePt = Math.max(3, fontSizePx * 0.75);
    const bold = isBoldFontWeight(cs.fontWeight);
    const italic = isItalicFontStyle(cs.fontStyle);
    const blockAlign = getEnclosingBlockTextAlign(parent);
    const isJustified = blockAlign === 'justify';

    const referenceScale = pages[0]?.sx ?? 0.75;
    const fragments = extractTextLineFragments(
      textNode,
      cs,
      isJustified,
      referenceScale,
      fontSizePx,
    );
    if (fragments.length === 0) continue;

    const font = await fontManager.getFont(cs.fontFamily, bold, italic);
    const ascentFrac = fontManager.getAscentFraction(font, fontSizePt);
    const color = parseCssColor(cs.color) ?? {r: 0, g: 0, b: 0, a: 1};
    const opacity = Math.min(1, Math.max(0, parseFloat(cs.opacity || '1'))) * color.a;
    const {underline, strike, decorationColor} = getEffectiveTextDecorations(parent);
    const decoCol = decorationColor ?? color;

    for (const frag of fragments) {
      const safeText = fontManager.sanitizeTextForFont(frag.text, font);
      if (!safeText) continue;

      const fragDomRect = new DOMRect(frag.left, frag.top, frag.width, frag.height);
      const mapped = resolveMappedPageForRect(fragDomRect, pages);
      if (!mapped) continue;

      const xPt = (frag.left - mapped.rect.left) * mapped.sx;
      const yTopPt = (frag.top - mapped.rect.top) * mapped.sy;
      const wPt = frag.width * mapped.sx;
      const hPt = frag.height * mapped.sy;
      const baselineY = mapped.pageHpt - (yTopPt + hPt * ascentFrac);

      mapped.pdfPage.drawText(safeText, {
        x: xPt,
        y: baselineY,
        size: fontSizePt,
        font,
        color: rgb(color.r, color.g, color.b),
        opacity,
      });

      const lineThickness = Math.max(0.45, fontSizePt * 0.065);
      if (underline) {
        const ulY = baselineY - fontSizePt * 0.12;
        drawStyledLine(
          mapped.pdfPage,
          xPt,
          ulY,
          xPt + wPt,
          ulY,
          lineThickness,
          'solid',
          decoCol,
          opacity,
        );
      }
      if (strike) {
        const stY = baselineY + fontSizePt * 0.28;
        drawStyledLine(
          mapped.pdfPage,
          xPt,
          stY,
          xPt + wPt,
          stY,
          lineThickness,
          'solid',
          decoCol,
          opacity,
        );
      }
    }
  }

  // Layer 3C: Editable footnote/endnote <input> elements in Edit mode
  for (const input of inputNodes) {
    const val = input.value.trim();
    if (!val) continue;
    const rect = input.getBoundingClientRect();
    if (rect.width <= 1 || rect.height <= 1) continue;
    const mapped = resolveMappedPageForRect(rect, pages);
    if (!mapped) continue;

    const cs = window.getComputedStyle(input);
    const fontSizePt = Math.max(4, (parseFloat(cs.fontSize || '11') || 11) * 0.75);
    const font = await fontManager.getFont(
      cs.fontFamily,
      isBoldFontWeight(cs.fontWeight),
      isItalicFontStyle(cs.fontStyle),
    );
    const safeVal = fontManager.sanitizeTextForFont(val, font);
    if (!safeVal) continue;
    const color = parseCssColor(cs.color) ?? {r: 0.2, g: 0.25, b: 0.33, a: 1};
    const padLeftPt = (parseFloat(cs.paddingLeft || '0') || 0) * 0.75;
    const xPt = (rect.left - mapped.rect.left) * mapped.sx + padLeftPt;
    const yTopPt = (rect.top - mapped.rect.top) * mapped.sy;
    const hPt = rect.height * mapped.sy;
    const ascentFrac = fontManager.getAscentFraction(font, fontSizePt);
    const baselineY = mapped.pageHpt - (yTopPt + (hPt - fontSizePt) / 2 + fontSizePt * ascentFrac);

    mapped.pdfPage.drawText(safeVal, {
      x: xPt,
      y: baselineY,
      size: fontSizePt,
      font,
      color: rgb(color.r, color.g, color.b),
      opacity: color.a,
    });
  }
}

// ---------------------------------------------------------------------------
// Public Browser DOM Exporters
// ---------------------------------------------------------------------------

/**
 * Exports View-mode `docx-preview` rendered pages (`section.esz-docx`) directly to PDF
 * using `Range.getClientRects()` and embedded TrueType fonts (never StandardFonts).
 */
export async function exportDocxPreviewDomToPdf(
  previewRoot: HTMLElement,
  options: DomPdfExportOptions,
): Promise<Uint8Array> {
  const sections = Array.from(previewRoot.querySelectorAll<HTMLElement>('section.esz-docx'));
  if (sections.length === 0) {
    throw new Error('No rendered Word preview pages found to export.');
  }

  if (typeof document !== 'undefined' && document.fonts?.ready) {
    await document.fonts.ready;
  }

  const layout = options.layout ?? DEFAULT_DOC_LAYOUT;
  const pdf = await PDFDocument.create();
  pdf.setTitle(options.title);
  pdf.setCreator('eSZ OFFICE');

  const fontManager = new PdfFontManager(pdf, layout.docDefaultFont ?? 'Times New Roman');
  const imageCache = new PdfImageCache(pdf);

  for (let i = 0; i < sections.length; i++) {
    const sec = sections[i];
    const secRect = sec.getBoundingClientRect();
    if (secRect.width <= 10 || secRect.height <= 10) {
      throw new Error(`Preview page ${i + 1} has invalid rendered dimensions.`);
    }

    const rawWpt = (sec.offsetWidth || secRect.width) * 0.75;
    const rawHpt = (sec.offsetHeight || secRect.height) * 0.75;
    const [pageWpt, pageHpt] = normalizeStandardPageSizePt(rawWpt, rawHpt);

    const pdfPage = pdf.addPage([pageWpt, pageHpt]);
    const sx = pageWpt / secRect.width;
    const sy = pageHpt / secRect.height;

    const mapped: MappedPage = {
      pageIndex: i,
      pdfPage,
      pageWpt,
      pageHpt,
      rect: secRect,
      sx,
      sy,
    };

    // Draw section's own page borders (if any real document border exists on the section)
    const secCs = window.getComputedStyle(sec);
    const hasDocBorder =
      (parseFloat(secCs.borderTopWidth || '0') > 1.2 && secCs.borderTopStyle !== 'none') ||
      (parseFloat(secCs.borderBottomWidth || '0') > 1.2 && secCs.borderBottomStyle !== 'none') ||
      (parseFloat(secCs.borderLeftWidth || '0') > 1.2 && secCs.borderLeftStyle !== 'none') ||
      (parseFloat(secCs.borderRightWidth || '0') > 1.2 && secCs.borderRightStyle !== 'none');

    if (hasDocBorder) {
      drawElementBoxBackgroundAndBorders(sec, mapped, secRect, true, false);
    }

    if (layout.watermarkEnabled && layout.watermarkText.trim()) {
      await drawWatermarkOnPage(pdfPage, pageWpt, pageHpt, layout.watermarkText, fontManager);
    }

    await renderDomSubtreeToPdfPages(sec, [mapped], fontManager, imageCache, {
      skipRootBox: true,
    });
  }

  return pdf.save();
}

/**
 * Exports Edit-mode paginated TipTap sheet (`[data-testid="docs-print-sheet"]`) directly to PDF
 * using `Range.getClientRects()` and embedded TrueType fonts (never StandardFonts).
 * Page breaks match the on-screen physical page cards 1:1.
 */
export async function exportPaginatedEditorDomToPdf(
  sheetEl: HTMLElement,
  options: DomPdfExportOptions,
): Promise<Uint8Array> {
  const pageSections = Array.from(
    sheetEl.querySelectorAll<HTMLElement>('section[data-testid="docs-page"]'),
  );
  if (pageSections.length === 0) {
    throw new Error('No paginated document pages found to export.');
  }

  if (typeof document !== 'undefined' && document.fonts?.ready) {
    await document.fonts.ready;
  }

  const layout = options.layout ?? DEFAULT_DOC_LAYOUT;
  const pdf = await PDFDocument.create();
  pdf.setTitle(options.title);
  pdf.setCreator('eSZ OFFICE');

  const fontManager = new PdfFontManager(pdf, layout.docDefaultFont ?? 'Times New Roman');
  const imageCache = new PdfImageCache(pdf);

  const [pageWpt, pageHpt] = getLayoutPageSizePt(layout);
  const mappedPages: MappedPage[] = [];

  // 1. Create physical PDF pages & paint Layer 1 (page borders, watermarks, headers, footers)
  for (let i = 0; i < pageSections.length; i++) {
    const sec = pageSections[i];
    const secRect = sec.getBoundingClientRect();
    if (secRect.width <= 10 || secRect.height <= 10) {
      throw new Error(`Document page ${i + 1} has invalid rendered dimensions.`);
    }

    const pdfPage = pdf.addPage([pageWpt, pageHpt]);
    const mapped: MappedPage = {
      pageIndex: i,
      pdfPage,
      pageWpt,
      pageHpt,
      rect: secRect,
      sx: pageWpt / secRect.width,
      sy: pageHpt / secRect.height,
    };
    mappedPages.push(mapped);

    // Draw page border (explicit OOXML overlay frame OR preset border)
    const borderFrame = sec.querySelector<HTMLElement>('[data-testid="docs-page-border-frame"]');
    if (borderFrame) {
      drawElementBoxBackgroundAndBorders(
        borderFrame,
        mapped,
        borderFrame.getBoundingClientRect(),
        true,
        false,
      );
    } else if (layout.pageBorder !== 'none' || layout.pageBorderConfig) {
      drawPageBorderOverlay(pdfPage, pageWpt, pageHpt, layout, i);
    }

    // Draw watermark behind content
    if (layout.watermarkEnabled && layout.watermarkText.trim()) {
      await drawWatermarkOnPage(pdfPage, pageWpt, pageHpt, layout.watermarkText, fontManager);
    }

    // Draw header & footer DOM elements
    const headerEl = sec.querySelector<HTMLElement>('[data-testid="docs-page-header"]');
    if (headerEl) {
      await renderDomSubtreeToPdfPages(headerEl, [mapped], fontManager, imageCache, {
        skipRootBox: true,
      });
    }
    const footerEl = sec.querySelector<HTMLElement>('[data-testid="docs-page-footer"]');
    if (footerEl) {
      await renderDomSubtreeToPdfPages(footerEl, [mapped], fontManager, imageCache, {
        skipRootBox: true,
      });
    }
  }

  // 2. Paint Layer 2 (Paginated ProseMirror rich content assigned 1:1 to exact screen pages)
  const proseMirror =
    sheetEl.querySelector<HTMLElement>('.ProseMirror') ??
    sheetEl.querySelector<HTMLElement>('.esz-doc-editor');
  if (proseMirror) {
    await renderDomSubtreeToPdfPages(proseMirror, mappedPages, fontManager, imageCache, {
      skipRootBox: true,
      skipTableOuterBoxWhenMultiPage: true,
    });
  }

  // 3. Paint Layer 3 (Per-page Footnotes & final-page Endnotes)
  const footnoteContainers = Array.from(
    sheetEl.querySelectorAll<HTMLElement>('[data-testid="docs-page-footnotes"]'),
  );
  for (const fnEl of footnoteContainers) {
    const pageNum = Number.parseInt(fnEl.getAttribute('data-page-number') || '1', 10);
    const targetPage = mappedPages[Math.max(0, Math.min(mappedPages.length - 1, pageNum - 1))];
    if (!targetPage) continue;
    await renderDomSubtreeToPdfPages(fnEl, [targetPage], fontManager, imageCache, {
      skipRootBox: false,
    });
  }

  return pdf.save();
}

// ---------------------------------------------------------------------------
// Headless Embedded-Font HTML-to-PDF Engine (Node / Vitest / Web-View Fallback)
// Never uses StandardFonts — always embeds bundled TrueType fonts via fontkit!
// ---------------------------------------------------------------------------

interface HeadlessRunStyle {
  fontFamily: string;
  fontSizePt: number;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  sub: boolean;
  sup: boolean;
  color: RgbaColor;
  bgColor: RgbaColor | null;
}

interface HeadlessSegment {
  kind: 'text' | 'tab' | 'br';
  text: string;
  style: HeadlessRunStyle;
  tabLeader?: string;
  tabAlign?: string;
  tabPosPt?: number;
  tabWidthPt?: number;
}

function parseCssBorderSide(raw: string | undefined): {
  widthPt: number;
  style: string;
  color: RgbaColor;
} | null {
  if (!raw) return null;
  const trimmed = raw.trim();
  if (!trimmed || /^none$/i.test(trimmed) || /^0(px|pt)?$/i.test(trimmed)) return null;
  const parts = trimmed.split(/\s+/);
  let widthPt = 0.75;
  let style = 'solid';
  let color: RgbaColor = {r: 0, g: 0, b: 0, a: 1};
  for (const part of parts) {
    const lower = part.toLowerCase();
    if (/^(solid|dashed|dotted|double|groove|ridge|none|hidden)$/.test(lower)) {
      style = lower;
    } else if (/^\d+(?:\.\d+)?(pt|px)$/.test(lower)) {
      widthPt = parseLengthPt(lower) ?? 0.75;
    } else {
      const parsedCol = parseCssColor(part);
      if (parsedCol) color = parsedCol;
    }
  }
  if (style === 'none' || style === 'hidden' || widthPt <= 0.05) return null;
  return {widthPt, style, color};
}

function collectHeadlessSegments(
  nodes: HtmlNode[],
  inherited: HeadlessRunStyle,
  out: HeadlessSegment[],
  footnoteIds: string[],
): void {
  for (const node of nodes) {
    if (node.type === 'text') {
      const text = node.text.replace(/\r/g, '');
      if (text) {
        out.push({kind: 'text', text, style: inherited});
      }
      continue;
    }

    const tag = node.tag;
    if (tag === 'br') {
      out.push({kind: 'br', text: '', style: inherited});
      continue;
    }

    if (node.attrs['data-note-kind'] === 'footnote' && node.attrs['data-note-ref']) {
      const id = node.attrs['data-note-ref'];
      if (!footnoteIds.includes(id)) footnoteIds.push(id);
    }

    if (tag === 'span' && node.attrs['data-doc-tab'] === 'true') {
      const posPx = node.attrs['data-tab-pos'] ? parseFloat(node.attrs['data-tab-pos']) : undefined;
      const widthPx = node.attrs['data-tab-width']
        ? parseFloat(node.attrs['data-tab-width'])
        : undefined;
      out.push({
        kind: 'tab',
        text: '\t',
        style: inherited,
        tabLeader: node.attrs['data-tab-leader'] || 'none',
        tabAlign: node.attrs['data-tab-align'] || 'left',
        tabPosPt: posPx && Number.isFinite(posPx) ? posPx * 0.75 : undefined,
        tabWidthPt: widthPx && Number.isFinite(widthPx) ? widthPx * 0.75 : 36,
      });
      continue;
    }

    const css = parseInlineCss(node.attrs.style);
    const next: HeadlessRunStyle = {...inherited};

    if (tag === 'strong' || tag === 'b') next.bold = true;
    if (tag === 'em' || tag === 'i') next.italic = true;
    if (tag === 'u') next.underline = true;
    if (tag === 's' || tag === 'del' || tag === 'strike') next.strike = true;
    if (tag === 'sub') {
      next.sub = true;
      next.fontSizePt = Math.max(6, inherited.fontSizePt * 0.72);
    }
    if (tag === 'sup') {
      next.sup = true;
      next.fontSizePt = Math.max(6, inherited.fontSizePt * 0.72);
    }
    if (tag === 'code') {
      next.fontFamily = 'Courier New';
    }

    if (css['font-weight']) {
      next.bold = isBoldFontWeight(css['font-weight']);
    }
    if (css['font-style']) {
      next.italic = isItalicFontStyle(css['font-style']);
    }
    if (css['text-decoration'] || css['text-decoration-line']) {
      const dec = (css['text-decoration-line'] || css['text-decoration'] || '').toLowerCase();
      if (dec.includes('underline')) next.underline = true;
      if (dec.includes('line-through')) next.strike = true;
      if (dec === 'none') {
        next.underline = false;
        next.strike = false;
      }
    }
    if (node.attrs['data-doc-font']) {
      next.fontFamily = node.attrs['data-doc-font'];
    } else if (css['font-family']) {
      next.fontFamily = css['font-family'];
    }
    if (css['font-size']) {
      const sz = parseLengthPt(css['font-size']);
      if (sz && sz >= 4) next.fontSizePt = sz;
    }
    if (css.color) {
      const col = parseCssColor(css.color);
      if (col) next.color = col;
    }
    if (css['background-color'] || css.background || tag === 'mark') {
      const bg =
        parseCssColor(css['background-color'] || css.background) ??
        (tag === 'mark' ? {r: 1, g: 0.96, b: 0.5, a: 1} : null);
      if (bg && !isNearWhite(bg)) next.bgColor = bg;
    }

    collectHeadlessSegments(node.children, next, out, footnoteIds);
  }
}

interface PositionedPiece {
  text: string;
  font: PDFFont;
  size: number;
  width: number;
  style: HeadlessRunStyle;
  isSpace: boolean;
  tabLeader?: string;
  tabStartX?: number;
  tabEndX?: number;
  explicitX?: number;
}

/**
 * Full-featured headless HTML-to-PDF exporter powered by pdf-lib + @pdf-lib/fontkit
 * and our bundled TrueType fonts (Liberation Serif/Sans/Mono, Carlito, Caladea).
 * Never uses any of the 14 PDF standard fonts.
 */
export async function exportHtmlToPdfHeadless(
  html: string,
  title: string,
  layout: DocPageLayout = DEFAULT_DOC_LAYOUT,
  notes: DocNoteItem[] = [],
): Promise<Uint8Array> {
  const pdf = await PDFDocument.create();
  pdf.setTitle(title);
  pdf.setCreator('eSZ OFFICE');

  const defaultFontFamily = layout.docDefaultFont ?? 'Times New Roman';
  const defaultFontSizePt = layout.docDefaultFontSizePt ?? 11;
  const defaultColor = parseCssColor(layout.docDefaultColor) ?? {r: 0, g: 0, b: 0, a: 1};
  const isImported = Boolean(layout.hasImportedStyles || layout.docDefaultFont);

  const fontManager = new PdfFontManager(pdf, defaultFontFamily);
  const imageCache = new PdfImageCache(pdf);

  const [pageW, pageH] = getLayoutPageSizePt(layout);
  const pm = getPageMetrics(layout);
  const marginTop = Math.max(18, pm.marginTop * 0.75);
  const marginBottom = Math.max(18, pm.marginBottom * 0.75);
  const marginLeft = Math.max(18, pm.marginLeft * 0.75);
  const marginRight = Math.max(18, pm.marginRight * 0.75);
  const contentWidth = Math.max(120, pageW - marginLeft - marginRight);

  const notesById = new Map<string, DocNoteItem>();
  for (const n of notes) notesById.set(n.id, n);
  const pageFootnoteMap = new Map<number, DocNoteItem[]>();
  const referencedFootnoteIds = new Set<string>();

  let pageIndex = 0;
  let page = pdf.addPage([pageW, pageH]);
  let y = pageH - marginTop;

  const getFootnoteReservePt = (pIdx: number, extraIds?: string[]): number => {
    const existing = pageFootnoteMap.get(pIdx) ?? [];
    const merged = new Set(existing.map((n) => n.id));
    if (extraIds) {
      for (const id of extraIds) {
        const note = notesById.get(id);
        if (note && note.kind === 'footnote') merged.add(id);
      }
    }
    return merged.size === 0 ? 0 : 18 + merged.size * 14;
  };

  const registerFootnotesOnCurrentPage = (ids: string[]) => {
    if (ids.length === 0) return;
    const list = pageFootnoteMap.get(pageIndex) ?? [];
    for (const id of ids) {
      const note = notesById.get(id);
      if (note && note.kind === 'footnote' && !list.some((x) => x.id === id)) {
        list.push(note);
        referencedFootnoteIds.add(id);
      }
    }
    if (list.length > 0) pageFootnoteMap.set(pageIndex, list);
  };

  const getMinY = (extraIds?: string[]) => marginBottom + 6 + getFootnoteReservePt(pageIndex, extraIds);

  const addFreshPage = () => {
    page = pdf.addPage([pageW, pageH]);
    pageIndex += 1;
    y = pageH - marginTop;
  };

  const drawRunDecorations = (
    targetPage: PDFPage,
    x: number,
    baselineY: number,
    width: number,
    size: number,
    style: HeadlessRunStyle,
  ) => {
    const thick = Math.max(0.45, size * 0.065);
    if (style.underline) {
      drawStyledLine(
        targetPage,
        x,
        baselineY - size * 0.12,
        x + width,
        baselineY - size * 0.12,
        thick,
        'solid',
        style.color,
      );
    }
    if (style.strike) {
      drawStyledLine(
        targetPage,
        x,
        baselineY + size * 0.28,
        x + width,
        baselineY + size * 0.28,
        thick,
        'solid',
        style.color,
      );
    }
  };

  const renderStyledParagraph = async (
    el: HtmlNode,
    prefixText?: string,
    extraIndentLeftPt = 0,
    maxWidthOverride?: number,
    baseXOverride?: number,
    targetPageOverride?: PDFPage,
    yStartOverride?: number,
    dryRunMeasureOnly = false,
  ): Promise<number> => {
    const css = parseInlineCss(el.attrs.style);
    const tag = el.tag;
    const isHeading = tag === 'h1' || tag === 'h2' || tag === 'h3' || tag === 'h4';

    const baseSizePt =
      parseLengthPt(css['font-size']) ??
      (isHeading && !isImported
        ? tag === 'h1'
          ? 18
          : tag === 'h2'
            ? 15
            : tag === 'h3'
              ? 13
              : 12
        : defaultFontSizePt);

    const baseColor =
      parseCssColor(css.color) ??
      (isHeading && !isImported ? {r: 0.11, g: 0.23, b: 0.54, a: 1} : defaultColor);

    const baseStyle: HeadlessRunStyle = {
      fontFamily: el.attrs['data-doc-font'] || css['font-family'] || defaultFontFamily,
      fontSizePt: baseSizePt,
      bold: css['font-weight'] ? isBoldFontWeight(css['font-weight']) : isHeading && !isImported,
      italic: css['font-style'] ? isItalicFontStyle(css['font-style']) : tag === 'blockquote',
      underline: (css['text-decoration'] || '').toLowerCase().includes('underline'),
      strike: (css['text-decoration'] || '').toLowerCase().includes('line-through'),
      sub: false,
      sup: false,
      color: baseColor,
      bgColor: null,
    };

    const align = (css['text-align'] || 'left').toLowerCase();
    const spaceBeforePt = parseLengthPt(css['margin-top']) ?? (isHeading && !isImported ? 8 : 0);
    const spaceAfterPt =
      parseLengthPt(css['margin-bottom']) ?? (isImported ? 0 : isHeading ? 5 : 6);
    const indentLeftPt = (parseLengthPt(css['margin-left']) ?? 0) + extraIndentLeftPt;
    const indentRightPt = parseLengthPt(css['margin-right']) ?? 0;
    const firstLineIndentPt = parseLengthPt(css['text-indent']) ?? 0;
    const lineHeightFactor = (() => {
      const raw = css['line-height'];
      if (!raw) return 1.25;
      const num = parseFloat(raw);
      if (!Number.isFinite(num) || num <= 0) return 1.25;
      if (raw.endsWith('pt') || raw.endsWith('px')) {
        const lhPt = parseLengthPt(raw) ?? baseSizePt * 1.25;
        return Math.max(0.9, lhPt / Math.max(6, baseSizePt));
      }
      return Math.max(0.9, Math.min(3.5, num));
    })();

    let tabStops: DocTabStop[] = [];
    if (el.attrs['data-tab-stops']) {
      try {
        const parsed = JSON.parse(el.attrs['data-tab-stops']);
        if (Array.isArray(parsed)) tabStops = parsed;
      } catch {
        // Ignore malformed stops
      }
    }
    const defaultTabStepPt = (layout.defaultTabStopPx ?? 48) * 0.75;

    const footnoteIds: string[] = [];
    const segments: HeadlessSegment[] = [];
    if (prefixText) {
      segments.push({kind: 'text', text: prefixText, style: baseStyle});
    }
    collectHeadlessSegments(el.children, baseStyle, segments, footnoteIds);

    const availW = Math.max(
      40,
      (maxWidthOverride ?? contentWidth) - indentLeftPt - indentRightPt,
    );
    const baseX = (baseXOverride ?? marginLeft) + indentLeftPt;

    // Build visual lines (respecting explicit tabs, <br>, and word wrapping)
    interface BuiltLine {
      pieces: PositionedPiece[];
      maxFontSizePt: number;
      lineHeightPt: number;
      hasExplicitBreak: boolean;
      hasTabs: boolean;
    }

    const lines: BuiltLine[] = [];
    let curPieces: PositionedPiece[] = [];
    let curX = firstLineIndentPt;
    let curMaxSize = baseSizePt;
    let lineHasTabs = false;

    const flushLine = (hasExplicitBreak: boolean) => {
      // Trim trailing space
      while (curPieces.length > 0 && curPieces[curPieces.length - 1].isSpace) {
        curPieces.pop();
      }
      lines.push({
        pieces: curPieces,
        maxFontSizePt: curMaxSize,
        lineHeightPt: Math.max(curMaxSize * 1.15, curMaxSize * lineHeightFactor),
        hasExplicitBreak,
        hasTabs: lineHasTabs,
      });
      curPieces = [];
      curX = 0;
      curMaxSize = baseSizePt;
      lineHasTabs = false;
    };

    // Check if paragraph contains any inline tabs so we can compute exact tab-stop positions
    for (let sIdx = 0; sIdx < segments.length; sIdx++) {
      const seg = segments[sIdx];
      if (seg.kind === 'br') {
        flushLine(true);
        continue;
      }

      if (seg.kind === 'tab') {
        lineHasTabs = true;
        // Find next explicit tab stop > curX
        let targetX: number | null = null;
        let stopAlign = seg.tabAlign || 'left';
        let stopLeader = seg.tabLeader || 'none';

        if (seg.tabPosPt !== undefined && seg.tabPosPt > curX + 1) {
          targetX = seg.tabPosPt;
        } else {
          const sortedStops = [...tabStops].sort((a, b) => a.posPx - b.posPx);
          const nextStop = sortedStops.find((s) => s.posPx * 0.75 > curX + 1);
          if (nextStop) {
            targetX = nextStop.posPx * 0.75;
            stopAlign = nextStop.val;
            stopLeader = nextStop.leader;
          } else {
            targetX = Math.floor(curX / defaultTabStepPt + 1) * defaultTabStepPt;
          }
        }

        // Measure following run width until next tab/br to support right/center tab stops!
        let nextRunWidthPt = 0;
        if (stopAlign === 'right' || stopAlign === 'center') {
          for (let j = sIdx + 1; j < segments.length; j++) {
            const nxt = segments[j];
            if (nxt.kind === 'tab' || nxt.kind === 'br') break;
            const nxtFont = await fontManager.getFont(
              nxt.style.fontFamily,
              nxt.style.bold,
              nxt.style.italic,
            );
            const clean = fontManager.sanitizeTextForFont(nxt.text, nxtFont);
            if (clean) {
              nextRunWidthPt += nxtFont.widthOfTextAtSize(clean, nxt.style.fontSizePt);
            }
          }
        }

        const clampedStopX = Math.min(availW, targetX);
        let landingX = clampedStopX;
        if (stopAlign === 'right') {
          landingX = Math.max(curX + 4, clampedStopX - nextRunWidthPt);
        } else if (stopAlign === 'center') {
          landingX = Math.max(curX + 4, clampedStopX - nextRunWidthPt / 2);
        }

        const font = await fontManager.getFont(
          seg.style.fontFamily,
          seg.style.bold,
          seg.style.italic,
        );
        curPieces.push({
          text: '',
          font,
          size: seg.style.fontSizePt,
          width: Math.max(4, landingX - curX),
          style: seg.style,
          isSpace: false,
          tabLeader: stopLeader,
          tabStartX: curX,
          tabEndX: landingX,
          explicitX: curX,
        });
        curX = landingX;
        continue;
      }

      const font = await fontManager.getFont(
        seg.style.fontFamily,
        seg.style.bold,
        seg.style.italic,
      );
      const clean = fontManager.sanitizeTextForFont(seg.text, font);
      if (!clean) continue;

      const tokens = clean.match(/\S+|\s+/g) ?? [];
      for (const token of tokens) {
        const isSpace = /^\s+$/.test(token);
        const normalizedTok = isSpace ? ' ' : token;
        if (isSpace && curPieces.length === 0) continue;

        const w = font.widthOfTextAtSize(normalizedTok, seg.style.fontSizePt);
        if (!isSpace && curPieces.length > 0 && curX + w > availW + 0.5) {
          flushLine(false);
        }

        curPieces.push({
          text: normalizedTok,
          font,
          size: seg.style.fontSizePt,
          width: w,
          style: seg.style,
          isSpace,
          explicitX: lineHasTabs ? curX : undefined,
        });
        curX += w;
        curMaxSize = Math.max(curMaxSize, seg.style.fontSizePt);
      }
    }

    if (curPieces.length > 0 || lines.length === 0) {
      flushLine(true);
    }

    const totalLinesH = lines.reduce((sum, l) => sum + l.lineHeightPt, 0);
    const totalH = spaceBeforePt + totalLinesH + spaceAfterPt;
    if (dryRunMeasureOnly) return totalH;

    // Render onto PDF page(s)
    const isTableCell = targetPageOverride !== undefined;
    let curY = yStartOverride ?? y;
    let activePage = targetPageOverride ?? page;

    if (!isTableCell) {
      if (curY < pageH - marginTop) {
        curY -= spaceBeforePt;
      }
      // Keep-with-next for headings / orphan control
      const needed = isHeading ? Math.min(totalLinesH + 24, pageH - marginTop - marginBottom) : lines[0].lineHeightPt;
      if (curY - needed < getMinY(footnoteIds) && curY < pageH - marginTop) {
        addFreshPage();
        activePage = page;
        curY = y;
      }
    }

    const paraStartY = curY;
    const paraBg = parseCssColor(css['background-color']);
    if (paraBg && !isNearWhite(paraBg)) {
      activePage.drawRectangle({
        x: baseX - 2,
        y: curY - totalLinesH - 2,
        width: availW + 4,
        height: totalLinesH + 4,
        color: rgb(paraBg.r, paraBg.g, paraBg.b),
        opacity: paraBg.a,
      });
    }

    for (let li = 0; li < lines.length; li++) {
      const line = lines[li];
      if (!isTableCell && curY - line.lineHeightPt < getMinY(footnoteIds) && curY < pageH - marginTop) {
        addFreshPage();
        activePage = page;
        curY = y;
      }

      const lineTotalW = line.hasTabs
        ? line.pieces.reduce((maxX, p) => Math.max(maxX, (p.explicitX ?? 0) + p.width), 0)
        : line.pieces.reduce((sum, p) => sum + p.width, 0);

      let lineStartX = baseX + (li === 0 ? firstLineIndentPt : 0);
      if (!line.hasTabs) {
        if (align === 'center') {
          lineStartX = baseX + Math.max(0, (availW - lineTotalW) / 2);
        } else if (align === 'right') {
          lineStartX = baseX + Math.max(0, availW - lineTotalW);
        }
      }

      // Justified extra space per inter-word gap (only on wrapped non-final lines without tabs)
      const isJustifyLine =
        align === 'justify' &&
        !line.hasExplicitBreak &&
        !line.hasTabs &&
        li < lines.length - 1;
      const spaceCount = isJustifyLine
        ? line.pieces.filter((p) => p.isSpace).length
        : 0;
      const extraSpacePerGap =
        isJustifyLine && spaceCount > 0 && availW > lineTotalW
          ? (availW - lineTotalW) / spaceCount
          : 0;

      const baselineY = curY - line.maxFontSizePt * 0.85;
      let drawX = lineStartX;

      for (const piece of line.pieces) {
        if (piece.tabStartX !== undefined && piece.tabEndX !== undefined) {
          if (piece.tabLeader && piece.tabLeader !== 'none') {
            const isMid = piece.tabLeader === 'middleDot' || piece.tabLeader === 'hyphen';
            const ly = baselineY + (isMid ? piece.size * 0.25 : 0);
            const lStyle =
              piece.tabLeader === 'dot' || piece.tabLeader === 'middleDot'
                ? 'dotted'
                : piece.tabLeader === 'hyphen'
                  ? 'dashed'
                  : 'solid';
            drawStyledLine(
              activePage,
              baseX + piece.tabStartX + 2,
              ly,
              baseX + piece.tabEndX - 2,
              ly,
              piece.tabLeader === 'heavy' ? 1.5 : 0.8,
              lStyle,
              piece.style.color,
            );
          }
          drawX = baseX + piece.tabEndX;
          continue;
        }

        if (line.hasTabs && piece.explicitX !== undefined) {
          drawX = baseX + piece.explicitX;
        }

        if (piece.style.bgColor) {
          activePage.drawRectangle({
            x: drawX,
            y: baselineY - piece.size * 0.2,
            width: piece.width,
            height: piece.size * 1.15,
            color: rgb(piece.style.bgColor.r, piece.style.bgColor.g, piece.style.bgColor.b),
            opacity: piece.style.bgColor.a,
          });
        }

        if (piece.text) {
          const yOffset = piece.style.sup
            ? piece.size * 0.42
            : piece.style.sub
              ? -piece.size * 0.22
              : 0;
          activePage.drawText(piece.text, {
            x: drawX,
            y: baselineY + yOffset,
            size: piece.size,
            font: piece.font,
            color: rgb(piece.style.color.r, piece.style.color.g, piece.style.color.b),
            opacity: piece.style.color.a,
          });
          drawRunDecorations(
            activePage,
            drawX,
            baselineY + yOffset,
            piece.width,
            piece.size,
            piece.style,
          );
        }

        drawX += piece.width + (piece.isSpace ? extraSpacePerGap : 0);
      }

      curY -= line.lineHeightPt;
    }

    // Paragraph borders (bottom / top / left blockquote)
    const bBottom = parseCssBorderSide(css['border-bottom']);
    if (bBottom) {
      drawStyledLine(
        activePage,
        baseX,
        curY + 1,
        baseX + availW,
        curY + 1,
        bBottom.widthPt,
        bBottom.style,
        bBottom.color,
      );
    }
    const bTop = parseCssBorderSide(css['border-top']);
    if (bTop) {
      drawStyledLine(
        activePage,
        baseX,
        paraStartY + 1,
        baseX + availW,
        paraStartY + 1,
        bTop.widthPt,
        bTop.style,
        bTop.color,
      );
    }
    if (tag === 'blockquote') {
      drawStyledLine(
        activePage,
        baseX - 8,
        curY + 2,
        baseX - 8,
        paraStartY,
        2,
        'solid',
        {r: 0.35, g: 0.45, b: 0.65, a: 1},
      );
    }

    if (!isTableCell) {
      registerFootnotesOnCurrentPage(footnoteIds);
      y = curY - spaceAfterPt;
    }

    return paraStartY - curY;
  };

  const renderImageElement = async (el: HtmlNode) => {
    const src = el.attrs.src || '';
    if (!src.startsWith('data:image/')) return;
    const pdfImg = await imageCache.embedFromDataUrl(src);
    if (!pdfImg) return;

    const css = parseInlineCss(el.attrs.style);
    const rawW =
      parseLengthPt(css.width) ??
      (el.attrs.width ? parseFloat(el.attrs.width) * 0.75 : undefined) ??
      Math.min(contentWidth, pdfImg.width * 0.75);
    const rawH =
      parseLengthPt(css.height) ??
      (el.attrs.height ? parseFloat(el.attrs.height) * 0.75 : undefined) ??
      rawW * (pdfImg.height / Math.max(1, pdfImg.width));

    const scale = rawW > contentWidth ? contentWidth / rawW : 1;
    const drawW = Math.max(8, rawW * scale);
    const drawH = Math.max(8, rawH * scale);
    const mt = parseLengthPt(css['margin-top']) ?? 4;
    const mb = parseLengthPt(css['margin-bottom']) ?? 4;

    if (y - (mt + drawH) < getMinY() && y < pageH - marginTop) {
      addFreshPage();
    }

    let x = marginLeft;
    const alignAttr = (el.attrs['data-align'] || '').toLowerCase();
    const isCentered =
      alignAttr === 'center' ||
      (css['margin-left'] === 'auto' && css['margin-right'] === 'auto');
    if (isCentered) {
      x = marginLeft + Math.max(0, (contentWidth - drawW) / 2);
    } else if (alignAttr === 'right') {
      x = pageW - marginRight - drawW;
    } else if (css['margin-left']) {
      const ml = parseLengthPt(css['margin-left']);
      if (ml !== undefined && ml > 0) x = marginLeft + ml;
    }

    y -= mt;
    page.drawImage(pdfImg, {
      x,
      y: y - drawH,
      width: drawW,
      height: drawH,
    });
    y -= drawH + mb;
  };

  const renderTableElement = async (tbl: HtmlNode) => {
    const rows: HtmlNode[] = [];
    const collectRows = (node: HtmlNode) => {
      for (const ch of node.children) {
        if (ch.type !== 'element') continue;
        if (ch.tag === 'tr') rows.push(ch);
        else if (ch.tag === 'thead' || ch.tag === 'tbody' || ch.tag === 'tfoot') {
          collectRows(ch);
        }
      }
    };
    collectRows(tbl);
    if (rows.length === 0) return;

    const maxCols = Math.max(
      1,
      ...rows.map(
        (r) =>
          r.children.filter((c) => c.type === 'element' && (c.tag === 'td' || c.tag === 'th'))
            .length,
      ),
    );

    // Derive column widths from first row's colwidth/width CSS or equal split
    const colWidths: number[] = Array(maxCols).fill(contentWidth / maxCols);
    const firstRowCells = rows[0].children.filter(
      (c) => c.type === 'element' && (c.tag === 'td' || c.tag === 'th'),
    );
    const parsedWidths = firstRowCells.map((cell) => {
      const cwAttr = cell.attrs.colwidth;
      if (cwAttr) {
        const px = parseFloat(cwAttr.split(',')[0]);
        if (Number.isFinite(px) && px > 10) return px * 0.75;
      }
      const css = parseInlineCss(cell.attrs.style);
      const wPt = parseLengthPt(css.width);
      if (wPt && wPt > 10) return wPt;
      return null;
    });
    if (parsedWidths.some((w) => w !== null)) {
      const knownSum = parsedWidths.reduce<number>((s, w) => s + (w ?? 0), 0);
      const unknownCount = parsedWidths.filter((w) => w === null).length;
      const fallbackW =
        unknownCount > 0 ? Math.max(30, (contentWidth - knownSum) / unknownCount) : 0;
      const raw = parsedWidths.map((w) => w ?? fallbackW);
      const rawSum = raw.reduce((s, v) => s + v, 0);
      const factor = rawSum > contentWidth ? contentWidth / rawSum : 1;
      for (let c = 0; c < Math.min(maxCols, raw.length); c++) {
        colWidths[c] = Math.max(24, raw[c] * factor);
      }
    }

    const cellPadX = 5;
    const cellPadY = 4;

    for (const row of rows) {
      const cells = row.children.filter(
        (c) => c.type === 'element' && (c.tag === 'td' || c.tag === 'th'),
      );
      if (cells.length === 0) continue;

      // Measure row height across all cells
      let rowHeight = 18;
      let colCursor = 0;
      const cellColSpans: Array<{cell: HtmlNode; xOffset: number; width: number}> = [];
      for (const cell of cells) {
        const colspan = Math.max(1, Number.parseInt(cell.attrs.colspan || '1', 10) || 1);
        let cellW = 0;
        for (let k = 0; k < colspan && colCursor + k < colWidths.length; k++) {
          cellW += colWidths[colCursor + k];
        }
        if (cellW <= 0) cellW = contentWidth / maxCols;
        const xOffset = colWidths.slice(0, colCursor).reduce((s, v) => s + v, 0);
        cellColSpans.push({cell, xOffset, width: cellW});
        colCursor += colspan;

        let cellContentH = cellPadY * 2;
        const blockKids = cell.children.filter(
          (ch) =>
            ch.type === 'element' &&
            ['p', 'h1', 'h2', 'h3', 'h4', 'div'].includes(ch.tag),
        );
        if (blockKids.length > 0) {
          for (const bk of blockKids) {
            cellContentH += await renderStyledParagraph(
              bk,
              undefined,
              0,
              Math.max(20, cellW - cellPadX * 2),
              marginLeft + xOffset + cellPadX,
              page,
              y,
              true,
            );
          }
        } else {
          cellContentH += await renderStyledParagraph(
            cell,
            undefined,
            0,
            Math.max(20, cellW - cellPadX * 2),
            marginLeft + xOffset + cellPadX,
            page,
            y,
            true,
          );
        }
        rowHeight = Math.max(rowHeight, cellContentH);
      }

      if (y - rowHeight < getMinY() && y < pageH - marginTop) {
        addFreshPage();
      }

      // Draw each cell's shading, borders, and content
      for (const {cell, xOffset, width: cellW} of cellColSpans) {
        const cx = marginLeft + xOffset;
        const cyBottom = y - rowHeight;
        const cellCss = parseInlineCss(cell.attrs.style);
        const bg =
          parseCssColor(cellCss['background-color']) ??
          (cell.tag === 'th' && !isImported ? {r: 0.945, g: 0.96, b: 0.976, a: 1} : null);

        if (bg && !isNearWhite(bg)) {
          page.drawRectangle({
            x: cx,
            y: cyBottom,
            width: cellW,
            height: rowHeight,
            color: rgb(bg.r, bg.g, bg.b),
            opacity: bg.a,
          });
        }

        const defaultBorder = {
          widthPt: 0.6,
          style: 'solid',
          color: {r: 0.65, g: 0.7, b: 0.76, a: 1},
        };
        const bTop = parseCssBorderSide(cellCss['border-top'] || cellCss.border) ?? defaultBorder;
        const bBottom =
          parseCssBorderSide(cellCss['border-bottom'] || cellCss.border) ?? defaultBorder;
        const bLeft = parseCssBorderSide(cellCss['border-left'] || cellCss.border) ?? defaultBorder;
        const bRight =
          parseCssBorderSide(cellCss['border-right'] || cellCss.border) ?? defaultBorder;

        drawStyledLine(page, cx, y, cx + cellW, y, bTop.widthPt, bTop.style, bTop.color);
        drawStyledLine(
          page,
          cx,
          cyBottom,
          cx + cellW,
          cyBottom,
          bBottom.widthPt,
          bBottom.style,
          bBottom.color,
        );
        drawStyledLine(page, cx, cyBottom, cx, y, bLeft.widthPt, bLeft.style, bLeft.color);
        drawStyledLine(
          page,
          cx + cellW,
          cyBottom,
          cx + cellW,
          y,
          bRight.widthPt,
          bRight.style,
          bRight.color,
        );

        let innerY = y - cellPadY;
        const blockKids = cell.children.filter(
          (ch) =>
            ch.type === 'element' &&
            ['p', 'h1', 'h2', 'h3', 'h4', 'div'].includes(ch.tag),
        );
        if (blockKids.length > 0) {
          for (const bk of blockKids) {
            const used = await renderStyledParagraph(
              bk,
              undefined,
              0,
              Math.max(20, cellW - cellPadX * 2),
              cx + cellPadX,
              page,
              innerY,
              false,
            );
            innerY -= used;
          }
        } else {
          await renderStyledParagraph(
            cell,
            undefined,
            0,
            Math.max(20, cellW - cellPadX * 2),
            cx + cellPadX,
            page,
            innerY,
            false,
          );
        }
      }

      y -= rowHeight;
    }
    y -= 6;
  };

  const walkBlockNodes = async (nodes: HtmlNode[], listDepth = 0): Promise<void> => {
    for (const node of nodes) {
      if (node.type === 'text') {
        if (node.text.trim()) {
          const fakeP: HtmlNode = {
            type: 'element',
            tag: 'p',
            attrs: {},
            children: [node],
            text: '',
          };
          await renderStyledParagraph(fakeP);
        }
        continue;
      }

      const tag = node.tag;
      if (node.attrs['data-page-break'] === 'true') {
        addFreshPage();
        continue;
      }
      if (tag === 'hr') {
        if (y - 12 < getMinY()) addFreshPage();
        drawStyledLine(page, marginLeft, y - 4, pageW - marginRight, y - 4, 0.75, 'solid', {
          r: 0.75,
          g: 0.8,
          b: 0.86,
          a: 1,
        });
        y -= 10;
        continue;
      }
      if (tag === 'img') {
        await renderImageElement(node);
        continue;
      }
      if (tag === 'table') {
        await renderTableElement(node);
        continue;
      }
      if (tag === 'ul' || tag === 'ol') {
        let itemIndex = Number.parseInt(node.attrs.start || '1', 10) || 1;
        for (const child of node.children) {
          if (child.type !== 'element' || child.tag !== 'li') continue;
          const prefix = tag === 'ol' ? `${itemIndex}. ` : '\u2022  ';
          itemIndex += 1;
          const innerBlocks = child.children.filter(
            (c) =>
              c.type === 'element' && ['p', 'h1', 'h2', 'h3', 'h4', 'ul', 'ol'].includes(c.tag),
          );
          if (innerBlocks.length > 0) {
            let appliedPrefix = false;
            for (const bk of innerBlocks) {
              if (bk.tag === 'ul' || bk.tag === 'ol') {
                await walkBlockNodes([bk], listDepth + 1);
              } else {
                await renderStyledParagraph(
                  bk,
                  appliedPrefix ? undefined : prefix,
                  (listDepth + 1) * 16,
                );
                appliedPrefix = true;
              }
            }
          } else {
            await renderStyledParagraph(child, prefix, (listDepth + 1) * 16);
          }
        }
        continue;
      }
      if (tag === 'blockquote') {
        const innerBlocks = node.children.filter(
          (c) => c.type === 'element' && c.tag === 'p',
        );
        if (innerBlocks.length > 0) {
          for (const bk of innerBlocks) {
            await renderStyledParagraph(bk, undefined, 16);
          }
        } else {
          await renderStyledParagraph(node, undefined, 16);
        }
        continue;
      }
      if (tag === 'p' || tag === 'h1' || tag === 'h2' || tag === 'h3' || tag === 'h4') {
        // Check if paragraph contains inline <img> children
        const imgs = node.children.filter(
          (c) => c.type === 'element' && c.tag === 'img',
        );
        const nonImgs = node.children.filter(
          (c) => !(c.type === 'element' && c.tag === 'img'),
        );
        const hasText = nonImgs.some(
          (c) => (c.type === 'text' && c.text.trim().length > 0) || c.type === 'element',
        );
        if (hasText || imgs.length === 0) {
          await renderStyledParagraph({...node, children: nonImgs});
        }
        for (const img of imgs) {
          await renderImageElement(img);
        }
        continue;
      }
      if (tag === 'div' || tag === 'section' || tag === 'article' || tag === 'body') {
        await walkBlockNodes(node.children, listDepth);
        continue;
      }
    }
  };

  const parsedTree = parseHtmlTree(html);
  await walkBlockNodes(parsedTree.children);

  // Render endnotes (and any unreferenced footnotes) at the end of the document
  const endnotes = notes.filter(
    (n) => n.kind === 'endnote' || !referencedFootnoteIds.has(n.id),
  );
  if (endnotes.length > 0) {
    const endnoteBold = await fontManager.getFont(defaultFontFamily, true, false);
    const endnoteReg = await fontManager.getFont(defaultFontFamily, false, false);
    if (y - 42 < getMinY()) addFreshPage();
    y -= 10;
    page.drawText('Endnotes', {
      x: marginLeft,
      y: y - 12,
      size: 12,
      font: endnoteBold,
      color: rgb(0.12, 0.16, 0.23),
    });
    y -= 18;
    for (const n of endnotes) {
      if (y - 14 < getMinY()) addFreshPage();
      const line = fontManager.sanitizeTextForFont(`[${n.number}] ${n.text}`, endnoteReg);
      page.drawText(line, {
        x: marginLeft,
        y: y - 9.5,
        size: 9.5,
        font: endnoteReg,
        color: rgb(0.25, 0.3, 0.38),
      });
      y -= 13;
    }
  }

  // Stamp page borders, watermarks, per-page footnotes, headers, footers & page numbers
  const chromeFont = await fontManager.getFont(defaultFontFamily, false, false);
  const allPages = pdf.getPages();
  for (let idx = 0; idx < allPages.length; idx++) {
    const p = allPages[idx];

    if (layout.pageBorderConfig || layout.pageBorder !== 'none') {
      drawPageBorderOverlay(p, pageW, pageH, layout, idx);
    }

    if (layout.watermarkEnabled && layout.watermarkText.trim()) {
      await drawWatermarkOnPage(p, pageW, pageH, layout.watermarkText, fontManager);
    }

    const pageFns = pageFootnoteMap.get(idx) ?? [];
    if (pageFns.length > 0) {
      let fnY = marginBottom + 8 + (pageFns.length - 1) * 13;
      drawStyledLine(
        p,
        marginLeft,
        fnY + 12,
        Math.min(pageW - marginRight, marginLeft + 160),
        fnY + 12,
        0.6,
        'solid',
        {r: 0.75, g: 0.8, b: 0.86, a: 1},
      );
      for (const fn of pageFns) {
        const fnLine = fontManager
          .sanitizeTextForFont(`[${fn.number}] ${fn.text}`, chromeFont)
          .slice(0, 120);
        p.drawText(fnLine, {
          x: marginLeft,
          y: fnY,
          size: 8.5,
          font: chromeFont,
          color: rgb(0.28, 0.34, 0.42),
        });
        fnY -= 13;
      }
    }

    const isSuppressedFirst = Boolean(layout.differentFirstPage && idx === 0);
    const headerStr = isSuppressedFirst
      ? (layout.firstPageHeaderText ?? '').trim()
      : layout.headerText.trim();
    const footerStr = isSuppressedFirst
      ? (layout.firstPageFooterText ?? '').trim()
      : layout.footerText.trim();

    const headerY = pageH - Math.max(18, Math.round(marginTop * 0.48));
    const footerY = Math.max(16, Math.round(marginBottom * 0.42));

    if (headerStr) {
      const safeHeader = fontManager.sanitizeTextForFont(headerStr, chromeFont);
      p.drawText(safeHeader, {
        x: marginLeft,
        y: headerY,
        size: 9,
        font: chromeFont,
        color: rgb(0.4, 0.45, 0.52),
      });
    }
    if (footerStr) {
      const safeFooter = fontManager.sanitizeTextForFont(footerStr, chromeFont);
      p.drawText(safeFooter, {
        x: marginLeft,
        y: footerY,
        size: 9,
        font: chromeFont,
        color: rgb(0.4, 0.45, 0.52),
      });
    }
    if (layout.showPageNumbers && !isSuppressedFirst) {
      const label = `Page ${idx + 1} of ${allPages.length}`;
      const lw = chromeFont.widthOfTextAtSize(label, 9);
      const px =
        layout.pageNumberPosition === 'bottom-right' || layout.pageNumberPosition === 'top-right'
          ? pageW - marginRight - lw
          : (pageW - lw) / 2;
      const py = layout.pageNumberPosition === 'top-right' ? headerY : footerY;
      p.drawText(label, {
        x: px,
        y: py,
        size: 9,
        font: chromeFont,
        color: rgb(0.35, 0.4, 0.48),
      });
    }
  }

  // Ensure at least one bundled font is referenced even for an empty document
  if (allPages.length > 0 && html.trim().length === 0) {
    await fontManager.getFont(defaultFontFamily, false, false);
  }

  return pdf.save();
}

export {resolveBundledPdfFontFamily};
