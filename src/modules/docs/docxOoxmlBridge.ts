import {canonicalizeWordFontFamily} from './docFonts';

// ---------------------------------------------------------------------------
// Browser & Node safe UTF-8 <-> Base64 helpers (for OOXML node attributes)
// ---------------------------------------------------------------------------

export function encodeUtf8Base64(str: string): string {
  if (typeof Buffer !== 'undefined') {
    return Buffer.from(str, 'utf8').toString('base64');
  }
  const bytes = new TextEncoder().encode(str);
  let binary = '';
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    const slice = bytes.subarray(i, i + chunkSize);
    binary += String.fromCharCode(...slice);
  }
  return btoa(binary);
}

export function decodeUtf8Base64(b64: string): string | null {
  try {
    if (typeof Buffer !== 'undefined') {
      return Buffer.from(b64, 'base64').toString('utf8');
    }
    const binary = atob(b64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
      bytes[i] = binary.charCodeAt(i);
    }
    return new TextDecoder('utf-8').decode(bytes);
  } catch {
    return null;
  }
}

// ---------------------------------------------------------------------------
// XML & HTML escaping utilities
// ---------------------------------------------------------------------------

export function escapeXmlText(val: string): string {
  return val
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

export function escapeXmlAttr(val: string): string {
  return val
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Injects HTML attributes onto the opening tag of a serialized HTML element string.
 * Handles both normal tags (`<p ...>`) and self-closing tags (`<img ... />`).
 */
export function injectHtmlAttrs(blockHtml: string, attrs: Record<string, string>): string {
  const entries = Object.entries(attrs).filter(([, v]) => v !== undefined && v !== '');
  if (entries.length === 0) return blockHtml;

  const trimmedStart = blockHtml.search(/<[A-Za-z]/);
  if (trimmedStart < 0) return blockHtml;

  // Scan forward to find the closing '>' of the opening tag, respecting quoted attribute values
  let inQuote: '"' | "'" | null = null;
  let endIdx = -1;
  for (let i = trimmedStart + 1; i < blockHtml.length; i++) {
    const ch = blockHtml[i];
    if (inQuote) {
      if (ch === inQuote) inQuote = null;
      continue;
    }
    if (ch === '"' || ch === "'") {
      inQuote = ch;
    } else if (ch === '>') {
      endIdx = i;
      break;
    }
  }
  if (endIdx < 0) return blockHtml;

  const isSelfClosing = blockHtml[endIdx - 1] === '/';
  const insertAt = isSelfClosing ? endIdx - 1 : endIdx;
  const serialized = entries.map(([k, v]) => `${k}="${escapeXmlAttr(v)}"`).join(' ');
  const prefix = blockHtml.slice(0, insertAt).replace(/\s+$/, '');
  const suffix = blockHtml.slice(insertAt);
  return `${prefix} ${serialized}${isSelfClosing ? ' ' : ''}${suffix}`;
}

// ---------------------------------------------------------------------------
// Universal lightweight HTML tree parser (identical in Browser & Node/Vitest)
// ---------------------------------------------------------------------------

export interface HtmlNode {
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

export function decodeHtmlEntityString(raw: string): string {
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

export function parseHtmlTree(html: string): HtmlNode {
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
        text: decodeHtmlEntityString(m[5]),
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
      attrs[k] = decodeHtmlEntityString(v);
    }

    const el: HtmlNode = {type: 'element', tag, attrs, children: [], text: ''};
    parent.children.push(el);
    if (!isSelfClose) {
      stack.push(el);
    }
  }

  return root;
}

export function parseInlineCss(styleAttr: string | undefined): Record<string, string> {
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

export function parseLengthPx(raw: string | undefined): number | undefined {
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

export function parseLengthPt(raw: string | undefined): number | undefined {
  if (!raw) return undefined;
  const m = raw.trim().match(/^(-?\d+(?:\.\d+)?)\s*(pt|px)?$/i);
  if (!m) return undefined;
  const val = Number(m[1]);
  if (!Number.isFinite(val)) return undefined;
  const unit = (m[2] ?? 'pt').toLowerCase();
  return unit === 'px' ? val * 0.75 : val;
}

export function normalizeHex6(raw: string | undefined): string | undefined {
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

// ---------------------------------------------------------------------------
// Canonical TipTap-Normalization-Invariant Block Signature
// ---------------------------------------------------------------------------

interface InlineTokenStyle {
  b?: 1;
  i?: 1;
  u?: 1;
  s?: 1;
  sub?: 1;
  sup?: 1;
  c?: string;
  bg?: string;
  f?: string;
  sz?: number;
  caps?: 1;
  scaps?: 1;
  ls?: number;
  href?: string;
}

type CanonicalInlineToken =
  | {k: 't'; v: string; st: string}
  | {k: 'tab'; pos: string; al: string; ld: string}
  | {k: 'br'}
  | {k: 'img'; src: string; w: number; h: number};

function fingerprintSrc(src: string | undefined): string {
  if (!src) return '';
  if (src.length <= 64) return src;
  return `${src.length}:${src.slice(0, 32)}:${src.slice(-32)}`;
}

function serializeInlineStyle(st: InlineTokenStyle): string {
  return [
    st.b ? 'b' : '',
    st.i ? 'i' : '',
    st.u ? 'u' : '',
    st.s ? 's' : '',
    st.sub ? 'sub' : '',
    st.sup ? 'sup' : '',
    st.c ? `c:${st.c}` : '',
    st.bg ? `bg:${st.bg}` : '',
    st.f ? `f:${st.f}` : '',
    st.sz ? `sz:${st.sz}` : '',
    st.caps ? 'caps' : '',
    st.scaps ? 'scaps' : '',
    st.ls ? `ls:${st.ls}` : '',
    st.href ? `href:${st.href}` : '',
  ]
    .filter(Boolean)
    .join('|');
}

function collectCanonicalInlineTokens(
  node: HtmlNode,
  inherited: InlineTokenStyle,
  out: CanonicalInlineToken[],
): void {
  if (node.type === 'text') {
    const normalized = node.text.replace(/[ \t\r\n]+/g, ' ');
    if (!normalized) return;
    const st = serializeInlineStyle(inherited);
    const prev = out[out.length - 1];
    if (prev && prev.k === 't' && prev.st === st) {
      prev.v += normalized;
    } else {
      out.push({k: 't', v: normalized, st});
    }
    return;
  }

  const tag = node.tag;
  if (tag === 'br') {
    out.push({k: 'br'});
    return;
  }

  if (tag === 'span' && node.attrs['data-doc-tab'] === 'true') {
    out.push({
      k: 'tab',
      pos: node.attrs['data-tab-pos'] ?? '',
      al: (node.attrs['data-tab-align'] ?? 'left').toLowerCase(),
      ld: (node.attrs['data-tab-leader'] ?? 'none').toLowerCase(),
    });
    return;
  }

  if (tag === 'img') {
    const css = parseInlineCss(node.attrs.style);
    const w = Math.round(
      Number(node.attrs['data-width']) || parseLengthPx(css.width) || Number(node.attrs.width) || 0,
    );
    const h = Math.round(
      Number(node.attrs['data-height']) || parseLengthPx(css.height) || Number(node.attrs.height) || 0,
    );
    out.push({k: 'img', src: fingerprintSrc(node.attrs.src), w, h});
    return;
  }

  const next: InlineTokenStyle = {...inherited};
  if (tag === 'strong' || tag === 'b') next.b = 1;
  else if (tag === 'em' || tag === 'i') next.i = 1;
  else if (tag === 'u') next.u = 1;
  else if (tag === 's' || tag === 'del') next.s = 1;
  else if (tag === 'sub') next.sub = 1;
  else if (tag === 'sup') next.sup = 1;
  else if (tag === 'code') next.f = 'Courier New';
  else if (tag === 'a' && node.attrs.href) next.href = node.attrs.href;

  const css = parseInlineCss(node.attrs.style);
  const rawFont = node.attrs['data-doc-font'] || css['font-family'];
  if (rawFont) {
    const canonical = canonicalizeWordFontFamily(rawFont);
    if (canonical) next.f = canonical;
  }
  if (css['font-size']) {
    const pt = parseLengthPt(css['font-size']);
    if (pt && pt > 0) next.sz = Math.round(pt * 2);
  }
  if (css.color) {
    const hex = normalizeHex6(css.color);
    if (hex) next.c = hex;
  }
  if (tag === 'mark' || css['background-color']) {
    const hex = normalizeHex6(node.attrs['data-color'] || css['background-color']);
    if (hex) next.bg = hex;
  }
  if (css['text-transform']?.toLowerCase() === 'uppercase') next.caps = 1;
  if (css['font-variant']?.toLowerCase() === 'small-caps') next.scaps = 1;
  if (css['letter-spacing']) {
    const lsPt = parseLengthPt(css['letter-spacing']);
    if (lsPt && lsPt !== 0) next.ls = Math.round(lsPt * 20);
  }

  for (const child of node.children) {
    collectCanonicalInlineTokens(child, next, out);
  }
}

function normalizeLineSpacingSig(raw: string | undefined): string {
  if (!raw) return '';
  const trimmed = raw.trim().toLowerCase();
  if (trimmed.endsWith('pt')) {
    const pt = parseLengthPt(trimmed);
    return pt !== undefined ? `${Math.round(pt * 10) / 10}pt` : trimmed;
  }
  const n = Number(trimmed);
  return Number.isFinite(n) ? String(Math.round(n * 100) / 100) : trimmed;
}

function normalizeNumSig(raw: string | undefined): string {
  if (raw === undefined || raw === '') return '';
  const n = Number(raw);
  return Number.isFinite(n) ? String(Math.round(n * 10) / 10) : '';
}

function fnv1aHashHex(input: string): string {
  let h1 = 0x811c9dc5;
  let h2 = 0x9e3779b9;
  for (let i = 0; i < input.length; i++) {
    const c = input.charCodeAt(i);
    h1 ^= c;
    h1 = Math.imul(h1, 0x01000193);
    h2 ^= c + (i & 0xff);
    h2 = Math.imul(h2, 0x85ebca6b);
  }
  return `${(h1 >>> 0).toString(16).padStart(8, '0')}${(h2 >>> 0).toString(16).padStart(8, '0')}`;
}

function buildCanonicalBlockObject(node: HtmlNode): unknown {
  if (node.type === 'text') {
    return {k: 'txt', v: node.text.replace(/[ \t\r\n]+/g, ' ')};
  }

  const tag = node.tag;

  if (node.attrs['data-page-break'] === 'true') {
    return {k: 'pb'};
  }

  if (node.attrs['data-doc-locked'] === 'true') {
    return {
      k: 'locked',
      lbl: node.attrs['data-locked-label'] ?? '',
      img: fingerprintSrc(node.attrs['data-locked-image']),
      w: normalizeNumSig(node.attrs['data-width']),
      h: normalizeNumSig(node.attrs['data-height']),
    };
  }

  if (node.attrs['data-doc-equation'] === 'true') {
    const tokens: CanonicalInlineToken[] = [];
    collectCanonicalInlineTokens(node, {}, tokens);
    return {k: 'eq', t: tokens};
  }

  if (node.attrs['data-doc-textbox'] === 'true') {
    const inner = node.children
      .filter((c) => c.type === 'element')
      .map((c) => buildCanonicalBlockObject(c));
    return {
      k: 'box',
      w: normalizeNumSig(node.attrs['data-box-width']),
      al: (node.attrs['data-box-align'] ?? '').toLowerCase(),
      fill: normalizeHex6(node.attrs['data-box-fill']) ?? node.attrs['data-box-fill'] ?? '',
      bc: normalizeHex6(node.attrs['data-box-border-color']) ?? '',
      bs: (node.attrs['data-box-border-style'] ?? '').toLowerCase(),
      bw: normalizeNumSig(node.attrs['data-box-border-width']),
      va: (node.attrs['data-box-valign'] ?? '').toLowerCase(),
      inner,
    };
  }

  if (tag === 'img') {
    const css = parseInlineCss(node.attrs.style);
    const w = Math.round(
      Number(node.attrs['data-width']) || parseLengthPx(css.width) || Number(node.attrs.width) || 0,
    );
    const h = Math.round(
      Number(node.attrs['data-height']) || parseLengthPx(css.height) || Number(node.attrs.height) || 0,
    );
    return {
      k: 'img',
      src: fingerprintSrc(node.attrs.src),
      w,
      h,
      al: (node.attrs['data-align'] ?? 'left').toLowerCase(),
      fl: node.attrs['data-floating'] === 'true' ? 1 : 0,
      wr: (node.attrs['data-wrap'] ?? '').toLowerCase(),
      flL: normalizeNumSig(node.attrs['data-float-left']),
      flT: normalizeNumSig(node.attrs['data-float-top']),
      mt: normalizeNumSig(node.attrs['data-margin-top']),
      mb: normalizeNumSig(node.attrs['data-margin-bottom']),
    };
  }

  if (tag === 'table') {
    const trEls: HtmlNode[] = [];
    const collectTrs = (n: HtmlNode) => {
      for (const c of n.children) {
        if (c.type !== 'element') continue;
        if (c.tag === 'tr') trEls.push(c);
        else if (c.tag === 'thead' || c.tag === 'tbody' || c.tag === 'tfoot') collectTrs(c);
      }
    };
    collectTrs(node);

    const rows = trEls.map((tr) =>
      tr.children
        .filter((c) => c.type === 'element' && (c.tag === 'td' || c.tag === 'th'))
        .map((cell) => {
          const cCss = parseInlineCss(cell.attrs.style);
          const inner = cell.children
            .filter((c) => c.type === 'element' || c.text.trim().length > 0)
            .map((c) => buildCanonicalBlockObject(c));
          return {
            tag: cell.tag,
            cs: Math.max(1, Number(cell.attrs.colspan) || 1),
            rs: Math.max(1, Number(cell.attrs.rowspan) || 1),
            bg: normalizeHex6(cell.attrs['data-bg-color'] || cCss['background-color']) ?? '',
            va: (cell.attrs['data-valign'] || cCss['vertical-align'] || 'top').toLowerCase(),
            inner,
          };
        }),
    );

    return {
      k: 'tbl',
      al: (node.attrs['data-table-align'] ?? 'left').toLowerCase(),
      w: normalizeNumSig(node.attrs['data-table-width']),
      rows,
    };
  }

  // Paragraph or heading (p, h1-h4, etc.)
  const css = parseInlineCss(node.attrs.style);
  const align = (css['text-align'] ?? 'left').toLowerCase();
  const indentLeft = Math.round(
    Number(node.attrs['data-indent-left']) ||
      (node.attrs['data-indent'] ? Number(node.attrs['data-indent']) * 32 : 0) ||
      parseLengthPx(css['margin-left']) ||
      0,
  );
  const indentRight = Math.round(
    Number(node.attrs['data-indent-right']) || parseLengthPx(css['margin-right']) || 0,
  );
  const firstLine = Math.round(
    Number(node.attrs['data-first-line-indent']) || parseLengthPx(css['text-indent']) || 0,
  );

  const tokens: CanonicalInlineToken[] = [];
  collectCanonicalInlineTokens(node, {}, tokens);

  return {
    k: tag,
    sid: node.attrs['data-style-id'] ?? '',
    al: align,
    il: indentLeft,
    ir: indentRight,
    fl: firstLine,
    sb: normalizeNumSig(node.attrs['data-space-before']),
    sa: normalizeNumSig(node.attrs['data-space-after']),
    ls: normalizeLineSpacingSig(node.attrs['data-line-spacing'] ?? css['line-height']),
    kn: node.attrs['data-keep-next'] === 'true' ? 1 : 0,
    kl: node.attrs['data-keep-lines'] === 'true' ? 1 : 0,
    ts: node.attrs['data-tab-stops'] ?? '',
    r: tokens,
  };
}

/**
 * Computes a deterministic, TipTap-DOM-normalization-invariant signature for a block HtmlNode.
 * Identical before TipTap import and after TipTap `editor.getHTML()` when the user made no edits,
 * and guaranteed to change when text, runs, alignment, spacing, tabs, images, or table cells change.
 */
export function computeBlockSignature(node: HtmlNode): string {
  const canonicalJson = JSON.stringify(buildCanonicalBlockObject(node));
  return fnv1aHashHex(canonicalJson);
}

export function computeBlockSignatureFromHtml(blockHtml: string): string {
  const tree = parseHtmlTree(blockHtml);
  const firstEl = tree.children.find((c) => c.type === 'element');
  if (!firstEl) return fnv1aHashHex(blockHtml);
  return computeBlockSignature(firstEl);
}
