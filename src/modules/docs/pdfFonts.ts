import fontkit from '@pdf-lib/fontkit';
import type {PDFDocument, PDFFont} from 'pdf-lib';
import {BUNDLED_FONT_FACES, isAppShellUiFont} from './docFonts';

export type BundledPdfFontFamilyId =
  | 'liberation-serif'
  | 'liberation-sans'
  | 'liberation-mono'
  | 'carlito'
  | 'caladea';

export interface ResolvedPdfFontSpec {
  familyId: BundledPdfFontFamilyId;
  weight: 400 | 700;
  style: 'normal' | 'italic';
  fileName: string;
  url: string;
  cacheKey: string;
}

const FONT_BYTE_CACHE = new Map<string, Uint8Array>();

function splitFontTokens(raw: string): string[] {
  const out: string[] = [];
  let current = '';
  let quote: '"' | "'" | null = null;
  for (let i = 0; i < raw.length; i++) {
    const ch = raw[i];
    if (quote) {
      if (ch === quote) quote = null;
      current += ch;
    } else if (ch === '"' || ch === "'") {
      quote = ch;
      current += ch;
    } else if (ch === ',') {
      const cleaned = current.trim().replace(/^["']+|["']+$/g, '').trim();
      if (cleaned) out.push(cleaned);
      current = '';
    } else {
      current += ch;
    }
  }
  const tail = current.trim().replace(/^["']+|["']+$/g, '').trim();
  if (tail) out.push(tail);
  return out;
}

/**
 * Maps any Word/CSS font-family declaration to one of the 5 bundled TrueType font families.
 * Never returns an app-shell UI font or any of the 14 PDF standard fonts.
 */
export function resolveBundledPdfFontFamily(
  fontFamily: string | null | undefined,
  fallback: BundledPdfFontFamilyId = 'liberation-serif',
): BundledPdfFontFamilyId {
  if (!fontFamily || !fontFamily.trim()) return fallback;
  const tokens = splitFontTokens(fontFamily);

  for (const token of tokens) {
    if (isAppShellUiFont(token)) continue;
    const lower = token.toLowerCase();

    if (
      lower === 'liberation serif' ||
      lower === 'times new roman' ||
      lower === 'times' ||
      lower === 'georgia' ||
      lower === 'garamond' ||
      lower === 'palatino' ||
      lower === 'palatino linotype' ||
      lower === 'book antiqua' ||
      lower === 'serif'
    ) {
      return 'liberation-serif';
    }
    if (
      lower === 'carlito' ||
      lower === 'calibri' ||
      lower === 'calibri light'
    ) {
      return 'carlito';
    }
    if (
      lower === 'caladea' ||
      lower === 'cambria' ||
      lower === 'cambria math'
    ) {
      return 'caladea';
    }
    if (
      lower === 'liberation mono' ||
      lower === 'courier new' ||
      lower === 'courier' ||
      lower === 'consolas' ||
      lower === 'menlo' ||
      lower === 'monaco' ||
      lower === 'monospace'
    ) {
      return 'liberation-mono';
    }
    if (
      lower === 'liberation sans' ||
      lower === 'arial' ||
      lower === 'helvetica' ||
      lower === 'helvetica neue' ||
      lower === 'verdana' ||
      lower === 'tahoma' ||
      lower === 'segoe ui' ||
      lower === 'trebuchet ms' ||
      lower === 'sans-serif'
    ) {
      return 'liberation-sans';
    }
  }

  // Secondary heuristic based on substrings if an uncommon named font was used
  for (const token of tokens) {
    if (isAppShellUiFont(token)) continue;
    const lower = token.toLowerCase();
    if (/mono|courier|consolas|fixed|code/i.test(lower)) return 'liberation-mono';
    if (/cambria|caladea/i.test(lower)) return 'caladea';
    if (/calibri|carlito/i.test(lower)) return 'carlito';
    if (/serif|times|roman|georgia|garamond|palatino|book/i.test(lower)) return 'liberation-serif';
    if (/sans|arial|helvetica|verdana|tahoma|gothic|segoe/i.test(lower)) return 'liberation-sans';
  }

  return fallback;
}

export function resolveBundledPdfFontSpec(
  fontFamily: string | null | undefined,
  bold: boolean,
  italic: boolean,
  fallback: BundledPdfFontFamilyId = 'liberation-serif',
): ResolvedPdfFontSpec {
  const familyId = resolveBundledPdfFontFamily(fontFamily, fallback);
  const weight: 400 | 700 = bold ? 700 : 400;
  const style: 'normal' | 'italic' = italic ? 'italic' : 'normal';
  const fileName = `${familyId}-${weight}-${style}.ttf`;
  return {
    familyId,
    weight,
    style,
    fileName,
    url: `/fonts/${fileName}`,
    cacheKey: `${familyId}:${weight}:${style}`,
  };
}

/**
 * Loads raw TrueType (.ttf) font bytes for a bundled font face, supporting both
 * browser runtime (fetch from /fonts/*.ttf) and Node/Vitest runtime (read from public/fonts/*.ttf).
 */
export async function loadBundledTtfFontBytes(spec: ResolvedPdfFontSpec): Promise<Uint8Array> {
  const cached = FONT_BYTE_CACHE.get(spec.cacheKey);
  if (cached) return cached;

  const isBrowserLike =
    typeof window !== 'undefined' &&
    typeof window.fetch === 'function' &&
    typeof window.location?.origin === 'string' &&
    window.location.origin.startsWith('http');

  if (isBrowserLike) {
    const absoluteUrl = new URL(spec.url, window.location.origin).href;
    const res = await window.fetch(absoluteUrl);
    if (!res.ok) {
      throw new Error(`Failed to load bundled PDF font ${spec.fileName} (HTTP ${res.status})`);
    }
    const bytes = new Uint8Array(await res.arrayBuffer());
    if (bytes.byteLength < 1024) {
      throw new Error(`Bundled PDF font ${spec.fileName} appears corrupt or truncated`);
    }
    FONT_BYTE_CACHE.set(spec.cacheKey, bytes);
    return bytes;
  }

  // Node / Vitest fallback: read directly from public/fonts/
  const fsMod = 'node:fs';
  const pathMod = 'node:path';
  const fs = (await import(/* @vite-ignore */ fsMod)) as typeof import('node:fs');
  const path = (await import(/* @vite-ignore */ pathMod)) as typeof import('node:path');
  const candidates = [
    path.resolve(process.cwd(), 'public', 'fonts', spec.fileName),
    path.resolve(process.cwd(), 'dist', 'fonts', spec.fileName),
  ];
  for (const candidate of candidates) {
    if (fs.existsSync(candidate)) {
      const buf = fs.readFileSync(candidate);
      const bytes = new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength);
      FONT_BYTE_CACHE.set(spec.cacheKey, bytes);
      return bytes;
    }
  }

  throw new Error(`Bundled PDF font file not found: ${spec.fileName}`);
}

export class PdfFontManager {
  private readonly pdf: PDFDocument;
  private readonly defaultFamily: BundledPdfFontFamilyId;
  private readonly embedded = new Map<string, PDFFont>();
  private readonly pending = new Map<string, Promise<PDFFont>>();

  constructor(pdf: PDFDocument, defaultFontFamily?: string | null) {
    this.pdf = pdf;
    this.pdf.registerFontkit(fontkit);
    this.defaultFamily = resolveBundledPdfFontFamily(defaultFontFamily, 'liberation-serif');
  }

  async getFont(
    fontFamily: string | null | undefined,
    bold = false,
    italic = false,
  ): Promise<PDFFont> {
    const spec = resolveBundledPdfFontSpec(fontFamily, bold, italic, this.defaultFamily);
    const hit = this.embedded.get(spec.cacheKey);
    if (hit) return hit;

    const inflight = this.pending.get(spec.cacheKey);
    if (inflight) return inflight;

    const loadPromise = (async () => {
      const bytes = await loadBundledTtfFontBytes(spec);
      const font = await this.pdf.embedFont(bytes, {subset: true});
      this.embedded.set(spec.cacheKey, font);
      this.pending.delete(spec.cacheKey);
      return font;
    })();

    this.pending.set(spec.cacheKey, loadPromise);
    return loadPromise;
  }

  /**
   * Computes the exact baseline ratio (ascender / (ascender + |descender|))
   * from the embedded TrueType font's metrics.
   */
  getAscentFraction(font: PDFFont, fontSizePt: number): number {
    try {
      const fullH = font.heightAtSize(fontSizePt, {descender: true});
      const ascH = font.heightAtSize(fontSizePt, {descender: false});
      if (fullH > 0 && ascH > 0 && ascH <= fullH) {
        return ascH / fullH;
      }
    } catch {
      // Fall through to typographic default
    }
    return 0.8;
  }

  /**
   * Cleans text for PDF embedding while preserving Unicode punctuation, accented characters,
   * Greek/Cyrillic, bullets, smart quotes, dashes, and math symbols supported by the TrueType font.
   */
  sanitizeTextForFont(raw: string, font: PDFFont): string {
    const cleaned = raw
      .replace(/\u00A0/g, ' ')
      .replace(/[\u200B-\u200D\uFEFF\r]/g, '')
      .replace(/\t/g, ' ');
    if (!cleaned) return '';

    // Fast path: attempt encoding the entire string directly
    try {
      font.encodeText(cleaned);
      return cleaned;
    } catch {
      // Fallback character-by-character so one rare emoji/symbol never fails the export
      let out = '';
      for (const ch of Array.from(cleaned)) {
        try {
          font.encodeText(ch);
          out += ch;
        } catch {
          if (ch === '\u2013' || ch === '\u2014') out += '-';
          else if (ch === '\u2018' || ch === '\u2019') out += "'";
          else if (ch === '\u201C' || ch === '\u201D') out += '"';
          else if (ch === '\u2022') out += '*';
          else if (ch === '\u2026') out += '...';
          else out += ' ';
        }
      }
      return out;
    }
  }
}

/**
 * Generates self-hosted @font-face CSS declarations for the browser print fallback iframe
 * so printed PDFs use the exact bundled metric-compatible fonts instead of OS fallbacks.
 */
export function buildEmbeddedFontFaceCss(origin = ''): string {
  const base = origin.replace(/\/$/, '');
  return BUNDLED_FONT_FACES.map(
    (face) => `@font-face {
  font-family: "${face.family}";
  src: url("${base}${face.url}") format("woff2");
  font-weight: ${face.weight};
  font-style: ${face.style};
  font-display: block;
}`,
  ).join('\n');
}
