import JSZip from 'jszip';
import {
  applyRenderFontSubstitutionToPreviewDom,
  canonicalizeWordFontFamily,
  resolveRenderFontStack,
} from './docFonts';

export interface DocxBorderSide {
  val: string;
  szPt: number;
  color: string;
  spacePt: number;
}

export interface DocxPageBordersSpec {
  offsetFrom: 'page' | 'text';
  display: 'allPages' | 'firstPage' | 'notFirstPage';
  top?: DocxBorderSide;
  right?: DocxBorderSide;
  bottom?: DocxBorderSide;
  left?: DocxBorderSide;
  marginsPt: {
    top: number;
    right: number;
    bottom: number;
    left: number;
  };
}

export interface ExtractedShapeRun {
  text: string;
  fontFamily?: string;
  fontSizePt?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
  highlight?: string;
  allCaps?: boolean;
  smallCaps?: boolean;
  letterSpacingPt?: number;
}

export interface ExtractedShapeParagraph {
  align?: 'left' | 'center' | 'right' | 'justify';
  spaceBeforePt?: number;
  spaceAfterPt?: number;
  lineHeight?: number;
  runs: ExtractedShapeRun[];
}

export interface ExtractedDrawingShape {
  id: string;
  widthPx: number;
  heightPx: number;
  isAnchor: boolean;
  wrapType: 'inline' | 'topAndBottom' | 'square' | 'none';
  alignH: 'left' | 'center' | 'right' | null;
  offsetXPx: number | null;
  offsetYPx: number | null;
  relativeFromH: string;
  relativeFromV: string;
  prstGeom: string;
  borderRadiusPx: number;
  fillColor: string;
  borderColor: string | null;
  borderWidthPx: number;
  borderStyle: 'solid' | 'dashed' | 'dotted' | 'none';
  insetsPx: {top: number; right: number; bottom: number; left: number};
  vAlign: 'top' | 'middle' | 'bottom';
  isWordArt: boolean;
  imageDataUrl?: string;
  paragraphs: ExtractedShapeParagraph[];
}

export interface ExtractedFloatingImage {
  id: string;
  dataUrl: string;
  widthPx: number;
  heightPx: number;
  leftPx: number;
  /** Offset relative to page top when relativeFromV is page/margin, or delta from anchor paragraph when paragraph/line */
  topPx: number;
  relativeFromV: 'page' | 'margin' | 'paragraph';
  behindDoc: boolean;
  alt: string;
}

export interface DocxDocDefaults {
  fontFamily: string | null;
  fontSizePt: number | null;
  color: string | null;
}

export interface PreprocessedDocxPreview {
  /** In-memory sanitized .docx bytes for docx-preview (original stored blob is never mutated) */
  buffer: ArrayBuffer;
  /** Extracted DrawingML mc:Choice shapes/textboxes/WordArt keyed by placeholder token */
  shapes: Record<string, ExtractedDrawingShape>;
  /** Off-centre anchored images rendered with position:absolute on the owning page element */
  floatingImages: Record<string, ExtractedFloatingImage>;
  /** Page borders extracted per w:sectPr in document order */
  sectionBorders: Array<DocxPageBordersSpec | null>;
  /** Default document typography from word/styles.xml docDefaults / Normal */
  docDefaults: DocxDocDefaults;
  /** Stats for diagnostics & unit testing */
  stats: {
    alternateContentCount: number;
    choiceUsedCount: number;
    fallbackDroppedCount: number;
    shapesExtractedCount: number;
    anchoredImagesCenteredCount: number;
    anchoredImagesAbsoluteCount: number;
    hasPageBorders: boolean;
  };
}

const EMU_PER_PX = 9525;
const TWIPS_PER_PT = 20;
const PT_TO_PX = 96 / 72;

export function isZipArchiveHeader(bytes: Uint8Array): boolean {
  return (
    bytes.length >= 4 &&
    bytes[0] === 0x50 &&
    bytes[1] === 0x4b &&
    bytes[2] === 0x03 &&
    bytes[3] === 0x04
  );
}

/**
 * Checks whether a Blob/ArrayBuffer is a genuine OOXML .docx archive containing word/document.xml.
 */
export async function isRealDocxArchive(buffer: ArrayBuffer): Promise<boolean> {
  const bytes = new Uint8Array(buffer);
  if (!isZipArchiveHeader(bytes)) return false;
  try {
    const zip = await JSZip.loadAsync(buffer);
    return Boolean(zip.file('word/document.xml'));
  } catch {
    return false;
  }
}

function decodeXmlEntities(raw: string): string {
  return raw
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex: string) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, dec: string) => String.fromCodePoint(parseInt(dec, 10)))
    .replace(/&amp;/g, '&');
}

function getAttr(tag: string, name: string): string | undefined {
  // Match either prefixed (w:val="...") or unprefixed (val="...")
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const re = new RegExp(`(?:^|[\\s<])(?:[A-Za-z0-9_-]+:)?${escaped}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`);
  const m = tag.match(re);
  if (!m) return undefined;
  return decodeXmlEntities(m[1] ?? m[2] ?? '');
}

function getExactAttr(tag: string, fullName: string): string | undefined {
  const escaped = fullName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const re = new RegExp(`(?:^|[\\s<])${escaped}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`);
  const m = tag.match(re);
  if (!m) return undefined;
  return decodeXmlEntities(m[1] ?? m[2] ?? '');
}

function parseNum(val: string | undefined, fallback: number): number {
  if (val == null || val === '') return fallback;
  const n = Number(val);
  return Number.isFinite(n) ? n : fallback;
}

function normalizeHexColor(val: string | undefined, autoDefault = '#000000'): string | undefined {
  if (!val) return undefined;
  const trimmed = val.trim();
  if (!trimmed || trimmed.toLowerCase() === 'none') return undefined;
  if (trimmed.toLowerCase() === 'auto') return autoDefault;
  if (/^[0-9a-fA-F]{6}$/.test(trimmed)) return `#${trimmed}`;
  if (/^#[0-9a-fA-F]{6}$/.test(trimmed)) return trimmed;
  return trimmed;
}

const HIGHLIGHT_MAP: Record<string, string> = {
  yellow: '#fef08a',
  green: '#bbf7d0',
  cyan: '#a5f3fc',
  magenta: '#f5d0fe',
  blue: '#bfdbfe',
  red: '#fecaca',
  darkBlue: '#1e3a8a',
  darkCyan: '#155e75',
  darkGreen: '#166534',
  darkMagenta: '#701a75',
  darkRed: '#991b1b',
  darkYellow: '#a16207',
  darkGray: '#475569',
  lightGray: '#e2e8f0',
  black: '#000000',
  white: '#ffffff',
};

/**
 * Finds all top-level balanced occurrences of an XML element by tag name (e.g., `mc:AlternateContent`).
 */
export function findBalancedElements(
  xml: string,
  tagName: string,
): Array<{start: number; end: number; openTag: string; inner: string; full: string}> {
  const results: Array<{start: number; end: number; openTag: string; inner: string; full: string}> =
    [];
  const escaped = tagName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const tokenRe = new RegExp(`<(\\/?)${escaped}(?=[\\s/>])([^>]*)>`, 'g');

  let depth = 0;
  let blockStart = -1;
  let openTag = '';
  let innerStart = -1;

  let match: RegExpExecArray | null;
  while ((match = tokenRe.exec(xml)) !== null) {
    const isClose = match[1] === '/';
    const rest = match[2] ?? '';
    const isSelfClosing = !isClose && /\/\s*$/.test(rest);
    const fullMatch = match[0];

    if (isSelfClosing) {
      if (depth === 0) {
        results.push({
          start: match.index,
          end: match.index + fullMatch.length,
          openTag: fullMatch,
          inner: '',
          full: fullMatch,
        });
      }
      continue;
    }

    if (!isClose) {
      if (depth === 0) {
        blockStart = match.index;
        openTag = fullMatch;
        innerStart = match.index + fullMatch.length;
      }
      depth += 1;
    } else if (depth > 0) {
      depth -= 1;
      if (depth === 0 && blockStart >= 0) {
        const end = match.index + fullMatch.length;
        results.push({
          start: blockStart,
          end,
          openTag,
          inner: xml.slice(innerStart, match.index),
          full: xml.slice(blockStart, end),
        });
        blockStart = -1;
      }
    }
  }

  return results;
}

function parseThemeFonts(themeXml: string | null): {major: string | null; minor: string | null} {
  if (!themeXml) return {major: null, minor: null};
  const majorEl = findBalancedElements(themeXml, 'a:majorFont')[0];
  const minorEl = findBalancedElements(themeXml, 'a:minorFont')[0];
  const majorTag = majorEl?.inner.match(/<a:latin\b[^>]*>/)?.[0];
  const minorTag = minorEl?.inner.match(/<a:latin\b[^>]*>/)?.[0];
  return {
    major: majorTag ? getAttr(majorTag, 'typeface') ?? null : null,
    minor: minorTag ? getAttr(minorTag, 'typeface') ?? null : null,
  };
}

function resolveRFontsTag(
  rFontsTag: string,
  themeFonts: {major: string | null; minor: string | null},
): string | null {
  const explicit = getAttr(rFontsTag, 'ascii') || getAttr(rFontsTag, 'hAnsi');
  if (explicit) return canonicalizeWordFontFamily(explicit) ?? explicit;

  const themeAttr = (
    getAttr(rFontsTag, 'asciiTheme') ||
    getAttr(rFontsTag, 'hAnsiTheme') ||
    ''
  ).toLowerCase();
  if (themeAttr.startsWith('major') && themeFonts.major) {
    return canonicalizeWordFontFamily(themeFonts.major) ?? themeFonts.major;
  }
  if (themeAttr.startsWith('minor') && themeFonts.minor) {
    return canonicalizeWordFontFamily(themeFonts.minor) ?? themeFonts.minor;
  }
  return null;
}

export function parseDocDefaultsFromStyles(
  stylesXml: string | null,
  themeXml: string | null = null,
): DocxDocDefaults {
  const result: DocxDocDefaults = {fontFamily: null, fontSizePt: null, color: null};
  if (!stylesXml) return result;

  const themeFonts = parseThemeFonts(themeXml);

  // 1. w:docDefaults > w:rPrDefault > w:rPr
  const docDefaults = findBalancedElements(stylesXml, 'w:docDefaults')[0];
  if (docDefaults) {
    const rPr = findBalancedElements(docDefaults.inner, 'w:rPr')[0];
    if (rPr) {
      const rFontsTag = rPr.inner.match(/<w:rFonts\b[^>]*>/);
      if (rFontsTag) {
        const f = resolveRFontsTag(rFontsTag[0], themeFonts);
        if (f) result.fontFamily = f;
      }
      const szTag = rPr.inner.match(/<w:sz\b[^>]*>/);
      if (szTag) {
        const halfPts = parseNum(getAttr(szTag[0], 'val'), 0);
        if (halfPts > 0) result.fontSizePt = halfPts / 2;
      }
      const colorTag = rPr.inner.match(/<w:color\b[^>]*>/);
      if (colorTag) {
        const c = normalizeHexColor(getAttr(colorTag[0], 'val'));
        if (c) result.color = c;
      }
    }
  }

  // 2. Normal style overrides (w:styleId="Normal" or w:default="1" paragraph style)
  const styles = findBalancedElements(stylesXml, 'w:style');
  const normalStyle = styles.find(
    (s) =>
      getAttr(s.openTag, 'type') === 'paragraph' &&
      (getAttr(s.openTag, 'styleId') === 'Normal' || getAttr(s.openTag, 'default') === '1'),
  );
  if (normalStyle) {
    const rPr = findBalancedElements(normalStyle.inner, 'w:rPr')[0];
    if (rPr) {
      const rFontsTag = rPr.inner.match(/<w:rFonts\b[^>]*>/);
      if (rFontsTag) {
        const f = resolveRFontsTag(rFontsTag[0], themeFonts);
        if (f) result.fontFamily = f;
      }
      const szTag = rPr.inner.match(/<w:sz\b[^>]*>/);
      if (szTag) {
        const halfPts = parseNum(getAttr(szTag[0], 'val'), 0);
        if (halfPts > 0) result.fontSizePt = halfPts / 2;
      }
      const colorTag = rPr.inner.match(/<w:color\b[^>]*>/);
      if (colorTag) {
        const c = normalizeHexColor(getAttr(colorTag[0], 'val'));
        if (c) result.color = c;
      }
    }
  }

  return result;
}

/**
 * Extracts canonical default document typography (font family, size, color) from
 * word/styles.xml (+ word/theme/theme1.xml) without mutating the source ArrayBuffer.
 */
export async function extractDocxDefaultTypography(
  buffer: ArrayBuffer,
): Promise<DocxDocDefaults> {
  try {
    const zip = await JSZip.loadAsync(buffer);
    const stylesXml = (await zip.file('word/styles.xml')?.async('string')) ?? null;
    const themeXml = (await zip.file('word/theme/theme1.xml')?.async('string')) ?? null;
    return parseDocDefaultsFromStyles(stylesXml, themeXml);
  } catch {
    return {fontFamily: null, fontSizePt: null, color: null};
  }
}

function parseBorderSideTag(tag: string | undefined): DocxBorderSide | undefined {
  if (!tag) return undefined;
  const val = (getAttr(tag, 'val') ?? 'none').toLowerCase();
  if (!val || val === 'none' || val === 'nil') return undefined;
  const szEighths = parseNum(getAttr(tag, 'sz'), 8);
  const szPt = Math.max(0.5, szEighths / 8);
  const color = normalizeHexColor(getAttr(tag, 'color'), '#000000') ?? '#000000';
  const spacePt = Math.max(0, parseNum(getAttr(tag, 'space'), 24));
  return {val, szPt, color, spacePt};
}

/**
 * Extracts w:pgBorders + w:pgMar from each w:sectPr in word/document.xml.
 */
export function extractSectionPageBorders(documentXml: string): Array<DocxPageBordersSpec | null> {
  const sectPrs = findBalancedElements(documentXml, 'w:sectPr');
  return sectPrs.map((sect) => {
    const pgBorders = findBalancedElements(sect.inner, 'w:pgBorders')[0];
    if (!pgBorders) return null;

    const offsetFromRaw = (getAttr(pgBorders.openTag, 'offsetFrom') ?? 'text').toLowerCase();
    const offsetFrom: 'page' | 'text' = offsetFromRaw === 'page' ? 'page' : 'text';
    const displayRaw = getAttr(pgBorders.openTag, 'display') ?? 'allPages';
    const display: 'allPages' | 'firstPage' | 'notFirstPage' =
      displayRaw === 'firstPage' || displayRaw === 'notFirstPage' ? displayRaw : 'allPages';

    const top = parseBorderSideTag(pgBorders.inner.match(/<w:top\b[^>]*>/)?.[0]);
    const right = parseBorderSideTag(pgBorders.inner.match(/<w:right\b[^>]*>/)?.[0]);
    const bottom = parseBorderSideTag(pgBorders.inner.match(/<w:bottom\b[^>]*>/)?.[0]);
    const left = parseBorderSideTag(pgBorders.inner.match(/<w:left\b[^>]*>/)?.[0]);

    if (!top && !right && !bottom && !left) return null;

    const pgMarTag = sect.inner.match(/<w:pgMar\b[^>]*>/)?.[0];
    const marginsPt = {
      top: parseNum(pgMarTag ? getAttr(pgMarTag, 'top') : undefined, 1440) / TWIPS_PER_PT,
      right: parseNum(pgMarTag ? getAttr(pgMarTag, 'right') : undefined, 1440) / TWIPS_PER_PT,
      bottom: parseNum(pgMarTag ? getAttr(pgMarTag, 'bottom') : undefined, 1440) / TWIPS_PER_PT,
      left: parseNum(pgMarTag ? getAttr(pgMarTag, 'left') : undefined, 1440) / TWIPS_PER_PT,
    };

    return {offsetFrom, display, top, right, bottom, left, marginsPt};
  });
}

function parseTextBoxParagraphs(
  txbxContentXml: string,
  docDefaults: DocxDocDefaults,
): ExtractedShapeParagraph[] {
  const paragraphs: ExtractedShapeParagraph[] = [];
  const pEls = findBalancedElements(txbxContentXml, 'w:p');

  for (const p of pEls) {
    const pPr = findBalancedElements(p.inner, 'w:pPr')[0];
    let align: ExtractedShapeParagraph['align'];
    let spaceBeforePt: number | undefined;
    let spaceAfterPt: number | undefined;
    let lineHeight: number | undefined;

    if (pPr) {
      const jcTag = pPr.inner.match(/<w:jc\b[^>]*>/)?.[0];
      const jc = jcTag ? getAttr(jcTag, 'val')?.toLowerCase() : undefined;
      if (jc === 'center') align = 'center';
      else if (jc === 'right' || jc === 'end') align = 'right';
      else if (jc === 'both' || jc === 'distribute') align = 'justify';
      else if (jc === 'left' || jc === 'start') align = 'left';

      const spacingTag = pPr.inner.match(/<w:spacing\b[^>]*>/)?.[0];
      if (spacingTag) {
        const before = getAttr(spacingTag, 'before');
        const after = getAttr(spacingTag, 'after');
        const line = getAttr(spacingTag, 'line');
        const lineRule = getAttr(spacingTag, 'lineRule') ?? 'auto';
        if (before != null) spaceBeforePt = parseNum(before, 0) / TWIPS_PER_PT;
        if (after != null) spaceAfterPt = parseNum(after, 0) / TWIPS_PER_PT;
        if (line != null && lineRule === 'auto') {
          const factor = parseNum(line, 240) / 240;
          if (factor > 0.5 && factor < 4) lineHeight = factor;
        }
      }
    }

    const runs: ExtractedShapeRun[] = [];
    const rEls = findBalancedElements(p.inner, 'w:r');
    for (const r of rEls) {
      const rPr = findBalancedElements(r.inner, 'w:rPr')[0];
      const run: ExtractedShapeRun = {text: ''};

      if (rPr) {
        const rFontsTag = rPr.inner.match(/<w:rFonts\b[^>]*>/)?.[0];
        if (rFontsTag) {
          run.fontFamily =
            getAttr(rFontsTag, 'ascii') || getAttr(rFontsTag, 'hAnsi') || undefined;
        }
        const szTag = rPr.inner.match(/<w:sz\b[^>]*>/)?.[0];
        if (szTag) {
          const halfPts = parseNum(getAttr(szTag, 'val'), 0);
          if (halfPts > 0) run.fontSizePt = halfPts / 2;
        }
        const bTag = rPr.inner.match(/<w:b\b[^>]*>/)?.[0];
        if (bTag) {
          const v = getAttr(bTag, 'val');
          run.bold = v !== '0' && v !== 'false';
        }
        const iTag = rPr.inner.match(/<w:i\b[^>]*>/)?.[0];
        if (iTag) {
          const v = getAttr(iTag, 'val');
          run.italic = v !== '0' && v !== 'false';
        }
        const uTag = rPr.inner.match(/<w:u\b[^>]*>/)?.[0];
        if (uTag) {
          const v = getAttr(uTag, 'val');
          run.underline = v !== 'none' && v !== '0';
        }
        const strikeTag = rPr.inner.match(/<w:strike\b[^>]*>/)?.[0];
        if (strikeTag) {
          const v = getAttr(strikeTag, 'val');
          run.strike = v !== '0' && v !== 'false';
        }
        const colorTag = rPr.inner.match(/<w:color\b[^>]*>/)?.[0];
        if (colorTag) {
          run.color = normalizeHexColor(getAttr(colorTag, 'val'));
        }
        // Check Word 2010+ WordArt solid text fill (<w14:textFill><w14:solidFill><w14:srgbClr val="..."/>)
        if (!run.color) {
          const w14Color = rPr.inner.match(/<w14:srgbClr\b[^>]*val="([0-9a-fA-F]{6})"/i)?.[1];
          if (w14Color) run.color = `#${w14Color}`;
        }
        const hlTag = rPr.inner.match(/<w:highlight\b[^>]*>/)?.[0];
        if (hlTag) {
          const hlName = getAttr(hlTag, 'val');
          if (hlName && HIGHLIGHT_MAP[hlName]) run.highlight = HIGHLIGHT_MAP[hlName];
        }
        const capsTag = rPr.inner.match(/<w:caps\b[^>]*>/)?.[0];
        if (capsTag && getAttr(capsTag, 'val') !== '0') run.allCaps = true;
        const smallCapsTag = rPr.inner.match(/<w:smallCaps\b[^>]*>/)?.[0];
        if (smallCapsTag && getAttr(smallCapsTag, 'val') !== '0') run.smallCaps = true;
        const spacingTag = rPr.inner.match(/<w:spacing\b[^>]*>/)?.[0];
        if (spacingTag) {
          const valTwips = parseNum(getAttr(spacingTag, 'val'), 0);
          if (valTwips !== 0) run.letterSpacingPt = valTwips / TWIPS_PER_PT;
        }
      }

      if (!run.fontFamily && docDefaults.fontFamily) run.fontFamily = docDefaults.fontFamily;
      if (!run.fontSizePt && docDefaults.fontSizePt) run.fontSizePt = docDefaults.fontSizePt;

      // Collect <w:t>, <w:tab/>, <w:br/> in order
      const pieceRe = /<w:t\b[^>]*>([\s\S]*?)<\/w:t>|<w:tab\b[^>]*\/>|<w:br\b[^>]*\/>/g;
      let textAcc = '';
      let pm: RegExpExecArray | null;
      while ((pm = pieceRe.exec(r.inner)) !== null) {
        if (pm[0].startsWith('<w:tab')) {
          textAcc += '\t';
        } else if (pm[0].startsWith('<w:br')) {
          textAcc += '\n';
        } else {
          textAcc += decodeXmlEntities(pm[1] ?? '');
        }
      }
      if (textAcc.length > 0) {
        run.text = textAcc;
        runs.push(run);
      }
    }

    paragraphs.push({
      align,
      spaceBeforePt,
      spaceAfterPt,
      lineHeight,
      runs,
    });
  }

  return paragraphs;
}

function extractColorFromDrawingFill(xmlSegment: string): string | null {
  if (/<a:noFill\b/.test(xmlSegment)) return 'transparent';
  const solidFill = findBalancedElements(xmlSegment, 'a:solidFill')[0];
  if (!solidFill) return null;
  const srgb = solidFill.inner.match(/<a:srgbClr\b[^>]*>/)?.[0];
  if (srgb) {
    return normalizeHexColor(getAttr(srgb, 'val')) ?? null;
  }
  const scheme = solidFill.inner.match(/<a:schemeClr\b[^>]*>/)?.[0];
  if (scheme) {
    const val = (getAttr(scheme, 'val') ?? '').toLowerCase();
    const schemeMap: Record<string, string> = {
      tx1: '#000000',
      dk1: '#000000',
      bg1: '#ffffff',
      lt1: '#ffffff',
      accent1: '#1d4ed8',
      accent2: '#b91c1c',
      accent3: '#15803d',
      accent4: '#6d28d9',
      accent5: '#0369a1',
      accent6: '#c2410c',
    };
    return schemeMap[val] ?? '#000000';
  }
  return null;
}

function parseDrawingShapeFromChoice(
  choiceXml: string,
  id: string,
  docDefaults: DocxDocDefaults,
  relTargets: Map<string, string>,
  imageDataByPath: Map<string, string>,
): ExtractedDrawingShape | null {
  const hasWsp = /<wps:wsp\b/.test(choiceXml) || /<wpg:wgp\b/.test(choiceXml);
  if (!hasWsp) return null;

  const isAnchor = /<wp:anchor\b/.test(choiceXml);

  // Dimensions from wp:extent or a:ext
  const wpExtent = choiceXml.match(/<wp:extent\b[^>]*>/)?.[0];
  const aExt = choiceXml.match(/<a:ext\b[^>]*>/)?.[0];
  const cxEmu = parseNum(
    (wpExtent ? getAttr(wpExtent, 'cx') : undefined) ?? (aExt ? getAttr(aExt, 'cx') : undefined),
    3600000,
  );
  const cyEmu = parseNum(
    (wpExtent ? getAttr(wpExtent, 'cy') : undefined) ?? (aExt ? getAttr(aExt, 'cy') : undefined),
    720000,
  );
  const widthPx = Math.max(24, Math.round(cxEmu / EMU_PER_PX));
  const heightPx = Math.max(20, Math.round(cyEmu / EMU_PER_PX));

  // Wrap type
  let wrapType: ExtractedDrawingShape['wrapType'] = isAnchor ? 'topAndBottom' : 'inline';
  if (/<wp:wrapNone\b/.test(choiceXml)) wrapType = 'none';
  else if (/<wp:wrapTopAndBottom\b/.test(choiceXml)) wrapType = 'topAndBottom';
  else if (/<wp:wrapSquare\b|<wp:wrapTight\b|<wp:wrapThrough\b/.test(choiceXml))
    wrapType = 'square';

  // Horizontal & vertical placement
  const posH = findBalancedElements(choiceXml, 'wp:positionH')[0];
  const posV = findBalancedElements(choiceXml, 'wp:positionV')[0];
  const relativeFromH = (posH ? getAttr(posH.openTag, 'relativeFrom') : undefined) ?? 'column';
  const relativeFromV = (posV ? getAttr(posV.openTag, 'relativeFrom') : undefined) ?? 'paragraph';

  let alignH: ExtractedDrawingShape['alignH'] = null;
  let offsetXPx: number | null = null;
  let offsetYPx: number | null = null;

  if (posH) {
    const alignMatch = posH.inner.match(/<wp:align>([^<]+)<\/wp:align>/);
    const alignVal = alignMatch?.[1]?.trim().toLowerCase();
    if (alignVal === 'center' || alignVal === 'left' || alignVal === 'right') {
      alignH = alignVal;
    }
    const offMatch = posH.inner.match(/<wp:posOffset>([^<]+)<\/wp:posOffset>/);
    if (offMatch) {
      offsetXPx = Math.round(parseNum(offMatch[1], 0) / EMU_PER_PX);
    }
  }
  if (posV) {
    const offMatch = posV.inner.match(/<wp:posOffset>([^<]+)<\/wp:posOffset>/);
    if (offMatch) {
      offsetYPx = Math.round(parseNum(offMatch[1], 0) / EMU_PER_PX);
    }
  }

  // Geometry, fill, stroke from wps:spPr
  const spPr = findBalancedElements(choiceXml, 'wps:spPr')[0];
  const spPrXml = spPr ? spPr.inner : choiceXml;
  const prstGeomTag = spPrXml.match(/<a:prstGeom\b[^>]*>/)?.[0];
  const prstGeom = (prstGeomTag ? getAttr(prstGeomTag, 'prst') : undefined) ?? 'rect';

  let borderRadiusPx = 0;
  if (/round/i.test(prstGeom)) {
    const adjMatch = spPrXml.match(/<a:gd\b[^>]*name="adj"[^>]*fmla="val\s+(\d+)"/i);
    const adjVal = adjMatch ? parseNum(adjMatch[1], 16667) : 16667;
    const minDim = Math.min(widthPx, heightPx);
    borderRadiusPx = Math.max(4, Math.min(Math.round(minDim / 2), Math.round((minDim * adjVal) / 100000)));
  } else if (/ellipse|oval/i.test(prstGeom)) {
    borderRadiusPx = 9999;
  }

  // Fill color (remove a:ln before inspecting shape fill so stroke fill isn't confused with shape fill)
  const lnEl = findBalancedElements(spPrXml, 'a:ln')[0];
  const spPrWithoutLn = lnEl
    ? spPrXml.slice(0, lnEl.start) + spPrXml.slice(lnEl.end)
    : spPrXml;
  const fillColor = extractColorFromDrawingFill(spPrWithoutLn) ?? 'transparent';

  // Stroke / border from a:ln
  let borderColor: string | null = null;
  let borderWidthPx = 0;
  let borderStyle: ExtractedDrawingShape['borderStyle'] = 'none';
  if (lnEl) {
    const strokeFill = extractColorFromDrawingFill(lnEl.inner);
    if (strokeFill && strokeFill !== 'transparent') {
      borderColor = strokeFill;
      const wEmu = parseNum(getAttr(lnEl.openTag, 'w'), 9525);
      borderWidthPx = Math.max(1, Number((wEmu / EMU_PER_PX).toFixed(2)));
      const dashTag = lnEl.inner.match(/<a:prstDash\b[^>]*>/)?.[0];
      const dashVal = (dashTag ? getAttr(dashTag, 'val') : 'solid') ?? 'solid';
      if (/dot/i.test(dashVal)) borderStyle = 'dotted';
      else if (/dash/i.test(dashVal)) borderStyle = 'dashed';
      else borderStyle = 'solid';
    }
  } else if (/roundRect/i.test(prstGeom)) {
    // Word rounded rectangle shapes default to a crisp border if styled via wps:style lnRef
    const styleEl = findBalancedElements(choiceXml, 'wps:style')[0];
    if (styleEl && /<a:lnRef\b[^>]*idx="(?!0")\d+"/.test(styleEl.inner)) {
      borderColor = '#000000';
      borderWidthPx = 1;
      borderStyle = 'solid';
    }
  }

  // Text body insets & vertical alignment
  const bodyPrTag = choiceXml.match(/<wps:bodyPr\b[^>]*>/)?.[0];
  const insetsPx = {
    left: Math.round(parseNum(bodyPrTag ? getAttr(bodyPrTag, 'lIns') : undefined, 91440) / EMU_PER_PX),
    right: Math.round(parseNum(bodyPrTag ? getAttr(bodyPrTag, 'rIns') : undefined, 91440) / EMU_PER_PX),
    top: Math.round(parseNum(bodyPrTag ? getAttr(bodyPrTag, 'tIns') : undefined, 45720) / EMU_PER_PX),
    bottom: Math.round(parseNum(bodyPrTag ? getAttr(bodyPrTag, 'bIns') : undefined, 45720) / EMU_PER_PX),
  };
  const anchorAttr = bodyPrTag ? getAttr(bodyPrTag, 'anchor') : undefined;
  const vAlign: ExtractedDrawingShape['vAlign'] =
    anchorAttr === 'ctr' ? 'middle' : anchorAttr === 'b' ? 'bottom' : 'top';

  const isWordArt = /<a:prstTxWarp\b|<w14:textFill\b|<w14:textOutline\b|<w14:shadow\b/.test(
    choiceXml,
  );

  // Extract optional blip image inside shape fill
  let imageDataUrl: string | undefined;
  const blipTag = spPrWithoutLn.match(/<a:blip\b[^>]*>/)?.[0];
  if (blipTag) {
    const rId = getExactAttr(blipTag, 'r:embed') || getAttr(blipTag, 'embed');
    if (rId && relTargets.has(rId)) {
      const targetPath = relTargets.get(rId)!;
      imageDataUrl = imageDataByPath.get(targetPath);
    }
  }

  // Extract all w:txbxContent blocks inside this shape (or grouped shapes)
  const txbxBlocks = findBalancedElements(choiceXml, 'w:txbxContent');
  const paragraphs: ExtractedShapeParagraph[] = [];
  for (const tb of txbxBlocks) {
    paragraphs.push(...parseTextBoxParagraphs(tb.inner, docDefaults));
  }

  // Skip empty invisible decorative boxes with no text, no border, no fill, and no image
  const hasVisibleText = paragraphs.some((p) => p.runs.some((r) => r.text.trim().length > 0));
  if (!hasVisibleText && !imageDataUrl && !borderColor && fillColor === 'transparent') {
    return null;
  }

  return {
    id,
    widthPx,
    heightPx,
    isAnchor,
    wrapType,
    alignH,
    offsetXPx,
    offsetYPx,
    relativeFromH,
    relativeFromV,
    prstGeom,
    borderRadiusPx,
    fillColor,
    borderColor,
    borderWidthPx,
    borderStyle,
    insetsPx,
    vAlign,
    isWordArt,
    imageDataUrl,
    paragraphs,
  };
}

/**
 * Resolves all <mc:AlternateContent> blocks in an OOXML XML part:
 * - Always selects <mc:Choice> and drops <mc:Fallback> (never rendering duplicates or blurry VML fallback)
 * - Extracts DrawingML shapes/text boxes/WordArt (<wps:wsp> / <wpg:wgp>) into structured placeholders
 *   for crisp post-render hydration inside docx-preview.
 */
export function resolveAlternateContentXml(
  xml: string,
  docDefaults: DocxDocDefaults,
  relTargets: Map<string, string>,
  imageDataByPath: Map<string, string>,
  shapesOut: Record<string, ExtractedDrawingShape>,
  counters: {
    alternateContentCount: number;
    choiceUsedCount: number;
    fallbackDroppedCount: number;
    shapeSeq: number;
  },
): string {
  const blocks = findBalancedElements(xml, 'mc:AlternateContent');
  if (blocks.length === 0) return xml;

  let output = '';
  let cursor = 0;

  for (const blk of blocks) {
    output += xml.slice(cursor, blk.start);
    counters.alternateContentCount += 1;

    const choices = findBalancedElements(blk.inner, 'mc:Choice');
    const fallbacks = findBalancedElements(blk.inner, 'mc:Fallback');
    if (fallbacks.length > 0) {
      counters.fallbackDroppedCount += fallbacks.length;
    }

    // Prefer a Choice that contains DrawingML shape or picture, otherwise first Choice
    const chosen =
      choices.find((c) => /<wps:wsp\b|<wpg:wgp\b|<pic:pic\b|<w:drawing\b/.test(c.inner)) ??
      choices[0];

    if (chosen) {
      counters.choiceUsedCount += 1;
      // Recursively resolve any nested mc:AlternateContent inside Choice
      const resolvedChoiceInner = resolveAlternateContentXml(
        chosen.inner,
        docDefaults,
        relTargets,
        imageDataByPath,
        shapesOut,
        counters,
      );

      if (/<wps:wsp\b|<wpg:wgp\b/.test(resolvedChoiceInner)) {
        const shapeId = `ESZ_SHAPE_${counters.shapeSeq++}`;
        const shape = parseDrawingShapeFromChoice(
          resolvedChoiceInner,
          shapeId,
          docDefaults,
          relTargets,
          imageDataByPath,
        );
        if (shape) {
          shapesOut[shapeId] = shape;
          // Check whether mc:AlternateContent is directly inside a <w:r> or outside a <w:r>
          const beforeSnippet = xml.slice(Math.max(0, blk.start - 200), blk.start);
          const lastROpen = beforeSnippet.lastIndexOf('<w:r');
          const lastRClose = beforeSnippet.lastIndexOf('</w:r>');
          const insideRun = lastROpen > lastRClose;
          output += insideRun
            ? `<w:t>__${shapeId}__</w:t>`
            : `<w:r><w:t>__${shapeId}__</w:t></w:r>`;
        }
      } else {
        // DrawingML picture (<pic:pic>) or other Choice content: unwrap cleanly, discarding Fallback
        output += resolvedChoiceInner;
      }
    }
    // Always ignore <mc:Fallback> so WordArt/text boxes never duplicate or degrade to blurry images

    cursor = blk.end;
  }

  output += xml.slice(cursor);
  return output;
}

/**
 * Replaces Word PAGE and NUMPAGES fields in header/footer XML with dynamic tokens
 * so each rendered page in docx-preview shows its true 1-based page number and total count.
 */
export function normalizePageNumberFieldsXml(xml: string): string {
  // 1. Simple fields: <w:fldSimple w:instr=" PAGE ">...</w:fldSimple>
  let out = xml.replace(
    /<w:fldSimple\b([^>]*)>([\s\S]*?)<\/w:fldSimple>/g,
    (full, attrs: string, inner: string) => {
      const instr = (getAttr(`<x ${attrs}>`, 'instr') ?? '').trim().toUpperCase();
      const firstRPr = inner.match(/<w:rPr\b[\s\S]*?<\/w:rPr>/)?.[0] ?? '';
      if (/^PAGE(\s|$)/.test(instr)) {
        return `<w:r>${firstRPr}<w:t>__ESZ_PAGE_NUM__</w:t></w:r>`;
      }
      if (/^NUMPAGES(\s|$)/.test(instr)) {
        return `<w:r>${firstRPr}<w:t>__ESZ_NUM_PAGES__</w:t></w:r>`;
      }
      return full;
    },
  );

  // 2. Complex fields: fldChar begin -> instrText (PAGE|NUMPAGES) -> separate -> display run -> end
  out = out.replace(
    /<w:r\b[^>]*>(?:(?!<\/w:r>)[\s\S])*?<w:fldChar\b[^>]*fldCharType="begin"[^>]*\/>[\s\S]*?<w:fldChar\b[^>]*fldCharType="end"[^>]*\/>(?:(?!<\/w:r>)[\s\S])*?<\/w:r>/g,
    (block) => {
      const instrTexts = Array.from(block.matchAll(/<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>/g))
        .map((m) => decodeXmlEntities(m[1] ?? ''))
        .join(' ')
        .trim()
        .toUpperCase();
      const firstRPr = block.match(/<w:rPr\b[\s\S]*?<\/w:rPr>/)?.[0] ?? '';
      if (/^PAGE(\s|$)/.test(instrTexts)) {
        return `<w:r>${firstRPr}<w:t>__ESZ_PAGE_NUM__</w:t></w:r>`;
      }
      if (/^NUMPAGES(\s|$)/.test(instrTexts)) {
        return `<w:r>${firstRPr}<w:t>__ESZ_NUM_PAGES__</w:t></w:r>`;
      }
      return block;
    },
  );

  return out;
}

function guessMimeType(filePath: string): string {
  const lower = filePath.toLowerCase();
  if (lower.endsWith('.png')) return 'image/png';
  if (lower.endsWith('.jpg') || lower.endsWith('.jpeg')) return 'image/jpeg';
  if (lower.endsWith('.gif')) return 'image/gif';
  if (lower.endsWith('.webp')) return 'image/webp';
  if (lower.endsWith('.svg')) return 'image/svg+xml';
  if (lower.endsWith('.bmp')) return 'image/bmp';
  return 'application/octet-stream';
}

function parseRelsMap(relsXml: string | null): Map<string, string> {
  const map = new Map<string, string>();
  if (!relsXml) return map;
  const relRe = /<Relationship\b[^>]*>/g;
  let m: RegExpExecArray | null;
  while ((m = relRe.exec(relsXml)) !== null) {
    const id = getAttr(m[0], 'Id');
    const target = getAttr(m[0], 'Target');
    if (id && target) {
      const normalized = target.startsWith('/')
        ? target.slice(1)
        : target.startsWith('word/')
          ? target
          : `word/${target.replace(/^\.\//, '')}`;
      map.set(id, normalized);
    }
  }
  return map;
}

interface SectionPageGeometry {
  endOffset: number;
  pageWidthPx: number;
  pageHeightPx: number;
  marginLeftPx: number;
  marginRightPx: number;
  marginTopPx: number;
  marginBottomPx: number;
}

function extractSectionGeometries(documentXml: string): SectionPageGeometry[] {
  const sectPrs = findBalancedElements(documentXml, 'w:sectPr');
  if (sectPrs.length === 0) {
    return [
      {
        endOffset: Number.POSITIVE_INFINITY,
        pageWidthPx: (11906 * 635) / EMU_PER_PX,
        pageHeightPx: (16838 * 635) / EMU_PER_PX,
        marginLeftPx: (1440 * 635) / EMU_PER_PX,
        marginRightPx: (1440 * 635) / EMU_PER_PX,
        marginTopPx: (1440 * 635) / EMU_PER_PX,
        marginBottomPx: (1440 * 635) / EMU_PER_PX,
      },
    ];
  }
  return sectPrs.map((sect) => {
    const pgSzTag = sect.inner.match(/<w:pgSz\b[^>]*>/)?.[0];
    const pgMarTag = sect.inner.match(/<w:pgMar\b[^>]*>/)?.[0];
    const wTwips = parseNum(pgSzTag ? getAttr(pgSzTag, 'w') : undefined, 11906);
    const hTwips = parseNum(pgSzTag ? getAttr(pgSzTag, 'h') : undefined, 16838);
    const leftTwips = parseNum(pgMarTag ? getAttr(pgMarTag, 'left') : undefined, 1440);
    const rightTwips = parseNum(pgMarTag ? getAttr(pgMarTag, 'right') : undefined, 1440);
    const topTwips = parseNum(pgMarTag ? getAttr(pgMarTag, 'top') : undefined, 1440);
    const bottomTwips = parseNum(pgMarTag ? getAttr(pgMarTag, 'bottom') : undefined, 1440);
    const twipsToPx = 635 / EMU_PER_PX;
    return {
      endOffset: sect.end,
      pageWidthPx: wTwips * twipsToPx,
      pageHeightPx: hTwips * twipsToPx,
      marginLeftPx: leftTwips * twipsToPx,
      marginRightPx: rightTwips * twipsToPx,
      marginTopPx: topTwips * twipsToPx,
      marginBottomPx: bottomTwips * twipsToPx,
    };
  });
}

function getGeometryForOffset(
  geometries: SectionPageGeometry[],
  offset: number,
): SectionPageGeometry {
  for (const g of geometries) {
    if (offset <= g.endOffset) return g;
  }
  return geometries[geometries.length - 1];
}

function buildInlineFromAnchor(anchorXml: string): string | null {
  const extentTag = anchorXml.match(/<wp:extent\b[^>]*\/>/)?.[0];
  const graphicEl = findBalancedElements(anchorXml, 'a:graphic')[0];
  if (!extentTag || !graphicEl) return null;

  const effectExtentTag =
    anchorXml.match(/<wp:effectExtent\b[^>]*\/>/)?.[0] ??
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>';
  const docPrTag =
    anchorXml.match(/<wp:docPr\b[^>]*\/>/)?.[0] ?? '<wp:docPr id="1" name="Picture"/>';
  const cNvEl = findBalancedElements(anchorXml, 'wp:cNvGraphicFramePr')[0];
  const cNvXml = cNvEl
    ? cNvEl.full
    : '<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>';

  return `<wp:inline distT="0" distB="0" distL="0" distR="0">${extentTag}${effectExtentTag}${docPrTag}${cNvXml}${graphicEl.full}</wp:inline>`;
}

function setParagraphCenteredPPr(pInner: string, distTTwips: number, distBTwips: number): string {
  const existingPPr = findBalancedElements(pInner, 'w:pPr')[0];
  if (!existingPPr) {
    const spacingAttr =
      distTTwips > 0 || distBTwips > 0
        ? `<w:spacing${distTTwips > 0 ? ` w:before="${distTTwips}"` : ''}${distBTwips > 0 ? ` w:after="${distBTwips}"` : ''}/>`
        : '';
    return `<w:pPr>${spacingAttr}<w:ind w:left="0" w:right="0" w:firstLine="0"/><w:jc w:val="center"/></w:pPr>${pInner}`;
  }

  let pPrInner = existingPPr.inner;
  // Force horizontal alignment to center
  if (/<w:jc\b[^>]*\/>/.test(pPrInner)) {
    pPrInner = pPrInner.replace(/<w:jc\b[^>]*\/>/, '<w:jc w:val="center"/>');
  } else {
    pPrInner += '<w:jc w:val="center"/>';
  }
  // Explicitly zero left/right/firstLine indents so style-inherited indents cannot skew centring
  pPrInner = pPrInner.replace(/<w:ind\b[^>]*\/>/g, '');
  pPrInner += '<w:ind w:left="0" w:right="0" w:firstLine="0"/>';
  if ((distTTwips > 0 || distBTwips > 0) && !/<w:spacing\b/.test(pPrInner)) {
    pPrInner += `<w:spacing${distTTwips > 0 ? ` w:before="${distTTwips}"` : ''}${distBTwips > 0 ? ` w:after="${distBTwips}"` : ''}/>`;
  }

  return (
    pInner.slice(0, existingPPr.start) +
    `${existingPPr.openTag}${pPrInner}</w:pPr>` +
    pInner.slice(existingPPr.end)
  );
}

/**
 * Step A.2: Normalizes anchored pictures (<wp:anchor> containing <pic:pic>) using
 * wrapTopAndBottom or wrapNone with relativeFrom="page" | "column" | "margin":
 * - When horizontal centre is within 5% of the page centre (or wp:align="center"),
 *   converts the anchor to a <wp:inline> inside a centred paragraph.
 * - Otherwise computes exact page-relative offsets and extracts the image into
 *   `floatingImagesOut` so hydrateDocxPreviewDom renders it with `position: absolute`
 *   directly on the owning `<section class="esz-docx">` page element.
 */
export function normalizeAnchoredImagesForPreviewXml(
  documentXml: string,
  relTargets: Map<string, string>,
  imageDataByPath: Map<string, string>,
  floatingImagesOut: Record<string, ExtractedFloatingImage>,
  counters: {
    anchoredImagesCenteredCount: number;
    anchoredImagesAbsoluteCount: number;
    floatImgSeq: number;
  },
): string {
  if (!/<wp:anchor\b/.test(documentXml)) return documentXml;

  const geometries = extractSectionGeometries(documentXml);
  const paragraphs = findBalancedElements(documentXml, 'w:p');
  if (paragraphs.length === 0) return documentXml;

  let out = '';
  let cursor = 0;

  for (const p of paragraphs) {
    out += documentXml.slice(cursor, p.start);
    cursor = p.end;

    if (!/<wp:anchor\b/.test(p.inner)) {
      out += p.full;
      continue;
    }

    const geom = getGeometryForOffset(geometries, p.start);
    const contentWidthPx = Math.max(
      60,
      geom.pageWidthPx - geom.marginLeftPx - geom.marginRightPx,
    );

    const drawings = findBalancedElements(p.inner, 'w:drawing');
    if (drawings.length === 0) {
      out += p.full;
      continue;
    }

    // Check whether this paragraph has visible text outside drawings
    let pWithoutDrawings = '';
    let dCur = 0;
    for (const d of drawings) {
      pWithoutDrawings += p.inner.slice(dCur, d.start);
      dCur = d.end;
    }
    pWithoutDrawings += p.inner.slice(dCur);
    const visibleText = Array.from(pWithoutDrawings.matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g))
      .map((m) => decodeXmlEntities(m[1] ?? ''))
      .join('')
      .trim();
    const hasVisibleText = visibleText.length > 0;

    let newPInner = '';
    let innerCursor = 0;
    let shouldCenterCurrentP = false;
    let centerDistTTwips = 0;
    let centerDistBTwips = 0;
    const insertedBeforePs: string[] = [];
    const insertedAfterPs: string[] = [];

    for (const d of drawings) {
      newPInner += p.inner.slice(innerCursor, d.start);
      innerCursor = d.end;

      const anchor = findBalancedElements(d.inner, 'wp:anchor')[0];
      if (
        !anchor ||
        !/<pic:pic\b/.test(anchor.inner) ||
        /<wps:wsp\b|<wpg:wgp\b/.test(anchor.inner)
      ) {
        newPInner += d.full;
        continue;
      }

      const isTopBottomOrNone =
        /<wp:wrapTopAndBottom\b/.test(anchor.inner) || /<wp:wrapNone\b/.test(anchor.inner);
      if (!isTopBottomOrNone) {
        newPInner += d.full;
        continue;
      }

      const posH = findBalancedElements(anchor.inner, 'wp:positionH')[0];
      const relativeFromH = (posH ? getAttr(posH.openTag, 'relativeFrom') : undefined)?.toLowerCase() ?? '';
      if (
        relativeFromH !== 'page' &&
        relativeFromH !== 'column' &&
        relativeFromH !== 'margin'
      ) {
        newPInner += d.full;
        continue;
      }

      const extentTag = anchor.inner.match(/<wp:extent\b[^>]*>/)?.[0];
      const cxEmu = parseNum(extentTag ? getAttr(extentTag, 'cx') : undefined, 0);
      const cyEmu = parseNum(extentTag ? getAttr(extentTag, 'cy') : undefined, 0);
      if (cxEmu <= 0 || cyEmu <= 0) {
        newPInner += d.full;
        continue;
      }
      const widthPx = cxEmu / EMU_PER_PX;
      const heightPx = cyEmu / EMU_PER_PX;

      // Compute horizontal left edge relative to page left edge
      const alignHVal = posH?.inner
        .match(/<wp:align>([^<]+)<\/wp:align>/)?.[1]
        ?.trim()
        .toLowerCase();
      const posOffsetHMatch = posH?.inner.match(/<wp:posOffset>([^<]+)<\/wp:posOffset>/);
      const offsetHPx = posOffsetHMatch ? parseNum(posOffsetHMatch[1], 0) / EMU_PER_PX : null;

      let pageLeftPx: number | null = null;
      if (alignHVal === 'center') {
        pageLeftPx =
          relativeFromH === 'page'
            ? (geom.pageWidthPx - widthPx) / 2
            : geom.marginLeftPx + (contentWidthPx - widthPx) / 2;
      } else if (alignHVal === 'left') {
        pageLeftPx = relativeFromH === 'page' ? 0 : geom.marginLeftPx;
      } else if (alignHVal === 'right') {
        pageLeftPx =
          relativeFromH === 'page'
            ? geom.pageWidthPx - widthPx
            : geom.marginLeftPx + contentWidthPx - widthPx;
      } else if (offsetHPx != null) {
        pageLeftPx = relativeFromH === 'page' ? offsetHPx : geom.marginLeftPx + offsetHPx;
      }

      if (pageLeftPx == null) {
        newPInner += d.full;
        continue;
      }

      const imgCenterX = pageLeftPx + widthPx / 2;
      const pageCenterX = geom.pageWidthPx / 2;
      const centerOffsetRatio = Math.abs(imgCenterX - pageCenterX) / geom.pageWidthPx;
      const isCentered = alignHVal === 'center' || centerOffsetRatio <= 0.05;

      // Parse vertical positioning for placement decision or absolute coordinates
      const posV = findBalancedElements(anchor.inner, 'wp:positionV')[0];
      const relativeFromVRaw =
        (posV ? getAttr(posV.openTag, 'relativeFrom') : undefined)?.toLowerCase() ?? 'paragraph';
      const alignVVal = posV?.inner
        .match(/<wp:align>([^<]+)<\/wp:align>/)?.[1]
        ?.trim()
        .toLowerCase();
      const posOffsetVMatch = posV?.inner.match(/<wp:posOffset>([^<]+)<\/wp:posOffset>/);
      const offsetVPx = posOffsetVMatch ? parseNum(posOffsetVMatch[1], 0) / EMU_PER_PX : 0;

      if (isCentered) {
        const inlineXml = buildInlineFromAnchor(anchor.inner);
        if (!inlineXml) {
          newPInner += d.full;
          continue;
        }
        const distTEmu = parseNum(getAttr(anchor.openTag, 'distT'), 0);
        const distBEmu = parseNum(getAttr(anchor.openTag, 'distB'), 0);
        const distTTwips = Math.min(720, Math.round(distTEmu / 635));
        const distBTwips = Math.min(720, Math.round(distBEmu / 635));

        counters.anchoredImagesCenteredCount += 1;

        if (hasVisibleText) {
          // Place the converted inline image in its own centred paragraph directly above or below the text
          const spacingXml =
            distTTwips > 0 || distBTwips > 0
              ? `<w:spacing${distTTwips > 0 ? ` w:before="${distTTwips}"` : ''}${distBTwips > 0 ? ` w:after="${distBTwips}"` : ''}/>`
              : '';
          const standaloneP = `<w:p><w:pPr>${spacingXml}<w:ind w:left="0" w:right="0" w:firstLine="0"/><w:jc w:val="center"/></w:pPr><w:r><w:drawing>${inlineXml}</w:drawing></w:r></w:p>`;
          if (
            (relativeFromVRaw === 'paragraph' || relativeFromVRaw === 'line') &&
            offsetVPx < -heightPx * 0.5
          ) {
            insertedBeforePs.push(standaloneP);
          } else {
            insertedAfterPs.push(standaloneP);
          }
          // Omit drawing from the text paragraph itself
        } else {
          shouldCenterCurrentP = true;
          centerDistTTwips = Math.max(centerDistTTwips, distTTwips);
          centerDistBTwips = Math.max(centerDistBTwips, distBTwips);
          newPInner += `<w:drawing>${inlineXml}</w:drawing>`;
        }
      } else {
        // Off-centre (>5% from page centre): extract and position:absolute relative to the owning page element
        const blipTag = anchor.inner.match(/<a:blip\b[^>]*>/)?.[0];
        const rId = blipTag ? getExactAttr(blipTag, 'r:embed') || getAttr(blipTag, 'embed') : undefined;
        const targetPath = rId ? relTargets.get(rId) : undefined;
        const dataUrl = targetPath ? imageDataByPath.get(targetPath) : undefined;
        if (!dataUrl) {
          newPInner += d.full;
          continue;
        }

        const docPrTag = anchor.inner.match(/<wp:docPr\b[^>]*>/)?.[0];
        const alt = (docPrTag ? getAttr(docPrTag, 'descr') || getAttr(docPrTag, 'name') : '') ?? '';
        const behindDoc = getAttr(anchor.openTag, 'behindDoc') === '1';

        let relativeFromV: ExtractedFloatingImage['relativeFromV'] = 'paragraph';
        let topPx = offsetVPx;
        if (relativeFromVRaw === 'page') {
          relativeFromV = 'page';
          if (alignVVal === 'top') topPx = 0;
          else if (alignVVal === 'center') topPx = (geom.pageHeightPx - heightPx) / 2;
          else if (alignVVal === 'bottom') topPx = geom.pageHeightPx - heightPx;
        } else if (relativeFromVRaw === 'margin' || relativeFromVRaw === 'topmargin') {
          relativeFromV = 'margin';
          const contentHeightPx = Math.max(
            60,
            geom.pageHeightPx - geom.marginTopPx - geom.marginBottomPx,
          );
          if (alignVVal === 'top') topPx = geom.marginTopPx;
          else if (alignVVal === 'center') {
            topPx = geom.marginTopPx + (contentHeightPx - heightPx) / 2;
          } else if (alignVVal === 'bottom') {
            topPx = geom.marginTopPx + contentHeightPx - heightPx;
          } else {
            topPx = geom.marginTopPx + offsetVPx;
          }
        } else {
          relativeFromV = 'paragraph';
          topPx = offsetVPx;
        }

        const floatId = `ESZ_FLOAT_IMG_${counters.floatImgSeq++}`;
        floatingImagesOut[floatId] = {
          id: floatId,
          dataUrl,
          widthPx: Math.round(widthPx * 100) / 100,
          heightPx: Math.round(heightPx * 100) / 100,
          leftPx: Math.round(pageLeftPx * 100) / 100,
          topPx: Math.round(topPx * 100) / 100,
          relativeFromV,
          behindDoc,
          alt,
        };
        counters.anchoredImagesAbsoluteCount += 1;

        // Leave an invisible inline marker token inside the paragraph so hydrateDocxPreviewDom
        // mounts the absolute-positioned image on the exact rendered <section.esz-docx> page.
        newPInner += `<w:t>__${floatId}__</w:t>`;
      }
    }

    newPInner += p.inner.slice(innerCursor);

    if (shouldCenterCurrentP) {
      newPInner = setParagraphCenteredPPr(newPInner, centerDistTTwips, centerDistBTwips);
    }

    out += insertedBeforePs.join('') + `${p.openTag}${newPInner}</w:p>` + insertedAfterPs.join('');
  }

  out += documentXml.slice(cursor);
  return out;
}

/**
 * Prepares an in-memory copy of a .docx ArrayBuffer for faithful View-mode rendering in docx-preview:
 * - Never writes or mutates the caller's original stored Blob/ArrayBuffer
 * - Resolves mc:AlternateContent using mc:Choice and ignoring mc:Fallback
 * - Normalizes anchored images (centred inline when within 5% of page centre; page-absolute otherwise)
 * - Extracts w:pgBorders, docDefaults, and DrawingML shapes/text boxes/WordArt
 * - Normalizes PAGE / NUMPAGES header/footer fields for exact per-page stamping
 */
export async function preprocessDocxForPreview(
  originalBuffer: ArrayBuffer,
): Promise<PreprocessedDocxPreview> {
  const zip = await JSZip.loadAsync(originalBuffer);
  const stylesXml = (await zip.file('word/styles.xml')?.async('string')) ?? null;
  const themeXml = (await zip.file('word/theme/theme1.xml')?.async('string')) ?? null;
  const docDefaults = parseDocDefaultsFromStyles(stylesXml, themeXml);

  const relsXml = (await zip.file('word/_rels/document.xml.rels')?.async('string')) ?? null;
  const relTargets = parseRelsMap(relsXml);

  // Preload media files referenced by shapes & floating images if any exist
  const imageDataByPath = new Map<string, string>();
  const mediaPaths = Object.keys(zip.files).filter(
    (p) => p.startsWith('word/media/') && !zip.files[p].dir,
  );
  for (const mediaPath of mediaPaths) {
    const b64 = await zip.file(mediaPath)?.async('base64');
    if (b64) {
      imageDataByPath.set(mediaPath, `data:${guessMimeType(mediaPath)};base64,${b64}`);
    }
  }

  const shapes: Record<string, ExtractedDrawingShape> = {};
  const floatingImages: Record<string, ExtractedFloatingImage> = {};
  const counters = {
    alternateContentCount: 0,
    choiceUsedCount: 0,
    fallbackDroppedCount: 0,
    shapeSeq: 0,
    anchoredImagesCenteredCount: 0,
    anchoredImagesAbsoluteCount: 0,
    floatImgSeq: 0,
  };

  let sectionBorders: Array<DocxPageBordersSpec | null> = [];
  const docXmlFile = zip.file('word/document.xml');
  if (docXmlFile) {
    const rawDocXml = await docXmlFile.async('string');
    sectionBorders = extractSectionPageBorders(rawDocXml);
    let processedDocXml = resolveAlternateContentXml(
      rawDocXml,
      docDefaults,
      relTargets,
      imageDataByPath,
      shapes,
      counters,
    );
    processedDocXml = normalizeAnchoredImagesForPreviewXml(
      processedDocXml,
      relTargets,
      imageDataByPath,
      floatingImages,
      counters,
    );
    zip.file('word/document.xml', processedDocXml);
  }

  // Process all headers and footers for mc:AlternateContent + dynamic PAGE / NUMPAGES fields
  const hfPaths = Object.keys(zip.files).filter((p) =>
    /^word\/(header|footer)\d*\.xml$/i.test(p),
  );
  for (const hfPath of hfPaths) {
    const rawHf = await zip.file(hfPath)?.async('string');
    if (!rawHf) continue;
    const hfRelsPath = hfPath.replace(/^word\//, 'word/_rels/') + '.rels';
    const hfRelsXml = (await zip.file(hfRelsPath)?.async('string')) ?? null;
    const hfRels = hfRelsXml ? parseRelsMap(hfRelsXml) : relTargets;

    let processedHf = resolveAlternateContentXml(
      rawHf,
      docDefaults,
      hfRels,
      imageDataByPath,
      shapes,
      counters,
    );
    processedHf = normalizePageNumberFieldsXml(processedHf);
    zip.file(hfPath, processedHf);
  }

  const outArrayBuffer = await zip.generateAsync({type: 'arraybuffer'});

  return {
    buffer: outArrayBuffer,
    shapes,
    floatingImages,
    sectionBorders,
    docDefaults,
    stats: {
      alternateContentCount: counters.alternateContentCount,
      choiceUsedCount: counters.choiceUsedCount,
      fallbackDroppedCount: counters.fallbackDroppedCount,
      shapesExtractedCount: Object.keys(shapes).length,
      anchoredImagesCenteredCount: counters.anchoredImagesCenteredCount,
      anchoredImagesAbsoluteCount: counters.anchoredImagesAbsoluteCount,
      hasPageBorders: sectionBorders.some((b) => b !== null),
    },
  };
}

function mapWordBorderStyleToCss(val: string): string {
  switch (val.toLowerCase()) {
    case 'double':
    case 'triple':
    case 'thinThickSmallGap':
    case 'thickThinSmallGap':
      return 'double';
    case 'dashed':
    case 'dashSmallGap':
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

function createShapeDomElement(shape: ExtractedDrawingShape, doc: Document): HTMLElement {
  const wrapper = doc.createElement('div');
  wrapper.setAttribute('data-testid', 'docx-shape-box');
  wrapper.setAttribute('data-shape-geom', shape.prstGeom);
  wrapper.className = 'esz-docx-shape-wrapper';
  wrapper.style.display = 'flex';
  wrapper.style.width = '100%';
  wrapper.style.boxSizing = 'border-box';

  if (shape.alignH === 'center' || shape.isWordArt || /roundRect/i.test(shape.prstGeom)) {
    wrapper.style.justifyContent =
      shape.alignH === 'left'
        ? 'flex-start'
        : shape.alignH === 'right'
          ? 'flex-end'
          : 'center';
  } else if (shape.alignH === 'right') {
    wrapper.style.justifyContent = 'flex-end';
  } else {
    wrapper.style.justifyContent = 'flex-start';
  }

  wrapper.style.marginTop = '6px';
  wrapper.style.marginBottom = '6px';

  const box = doc.createElement('div');
  box.className = 'esz-docx-shape-box';
  box.style.boxSizing = 'border-box';
  box.style.width = `${shape.widthPx}px`;
  box.style.maxWidth = '100%';
  box.style.minHeight = `${shape.heightPx}px`;
  box.style.display = 'flex';
  box.style.flexDirection = 'column';
  box.style.justifyContent =
    shape.vAlign === 'middle' ? 'center' : shape.vAlign === 'bottom' ? 'flex-end' : 'flex-start';
  box.style.padding = `${shape.insetsPx.top}px ${shape.insetsPx.right}px ${shape.insetsPx.bottom}px ${shape.insetsPx.left}px`;
  box.style.backgroundColor = shape.fillColor;

  if (shape.borderRadiusPx > 0) {
    box.style.borderRadius = `${shape.borderRadiusPx}px`;
  }
  if (shape.borderColor && shape.borderWidthPx > 0 && shape.borderStyle !== 'none') {
    box.style.border = `${shape.borderWidthPx}px ${shape.borderStyle} ${shape.borderColor}`;
  }

  if (shape.imageDataUrl) {
    const img = doc.createElement('img');
    img.src = shape.imageDataUrl;
    img.alt = '';
    img.style.maxWidth = '100%';
    img.style.height = 'auto';
    img.style.display = 'block';
    img.style.margin = '0 auto';
    box.appendChild(img);
  }

  for (const para of shape.paragraphs) {
    const pEl = doc.createElement('p');
    pEl.style.margin = '0';
    if (para.spaceBeforePt != null) pEl.style.marginTop = `${para.spaceBeforePt}pt`;
    if (para.spaceAfterPt != null) pEl.style.marginBottom = `${para.spaceAfterPt}pt`;
    if (para.lineHeight != null) pEl.style.lineHeight = String(para.lineHeight);
    if (para.align) pEl.style.textAlign = para.align;
    else if (shape.isWordArt) pEl.style.textAlign = 'center';

    if (para.runs.length === 0) {
      pEl.innerHTML = '&nbsp;';
    } else {
      for (const run of para.runs) {
        const span = doc.createElement('span');
        if (run.fontFamily) span.style.fontFamily = resolveRenderFontStack(run.fontFamily);
        if (run.fontSizePt) span.style.fontSize = `${run.fontSizePt}pt`;
        if (run.bold) span.style.fontWeight = '700';
        if (run.italic) span.style.fontStyle = 'italic';
        if (run.underline && run.strike) span.style.textDecoration = 'underline line-through';
        else if (run.underline) span.style.textDecoration = 'underline';
        else if (run.strike) span.style.textDecoration = 'line-through';
        if (run.color) span.style.color = run.color;
        if (run.highlight) span.style.backgroundColor = run.highlight;
        if (run.allCaps) span.style.textTransform = 'uppercase';
        if (run.smallCaps) span.style.fontVariant = 'small-caps';
        if (run.letterSpacingPt) span.style.letterSpacing = `${run.letterSpacingPt}pt`;
        span.textContent = run.text;
        pEl.appendChild(span);
      }
    }
    box.appendChild(pEl);
  }

  wrapper.appendChild(box);
  return wrapper;
}

/**
 * Post-processes the DOM generated by docx-preview:
 * 1. Hydrates extracted mc:Choice DrawingML shape/textbox/WordArt markers into crisp HTML boxes
 * 2. Stamps dynamic PAGE and NUMPAGES fields on each rendered section.esz-docx page
 * 3. Renders exact w:pgBorders frames inside each section.esz-docx page
 */
export function hydrateDocxPreviewDom(
  container: HTMLElement,
  preview: PreprocessedDocxPreview,
  styleContainer?: HTMLElement | null,
): number {
  const doc = container.ownerDocument;
  const pages = Array.from(container.querySelectorAll<HTMLElement>('section.esz-docx'));
  const totalPages = pages.length;

  // 0. Apply render-only SIL OFL metric-compatible font substitution to generated styles & inline spans
  applyRenderFontSubstitutionToPreviewDom(container, styleContainer);

  // 1. Hydrate extracted DrawingML mc:Choice shapes
  const shapeTokens = Object.keys(preview.shapes);
  if (shapeTokens.length > 0) {
    const walker = doc.createTreeWalker(container, NodeFilter.SHOW_TEXT);
    const matches: Array<{node: Text; shapeId: string}> = [];
    let current: Node | null;
    while ((current = walker.nextNode()) !== null) {
      const textNode = current as Text;
      const val = textNode.nodeValue ?? '';
      const m = val.match(/__(ESZ_SHAPE_\d+)__/);
      if (m && preview.shapes[m[1]]) {
        matches.push({node: textNode, shapeId: m[1]});
      }
    }
    for (const {node, shapeId} of matches) {
      const shape = preview.shapes[shapeId];
      const shapeEl = createShapeDomElement(shape, doc);
      const val = node.nodeValue ?? '';
      const token = `__${shapeId}__`;
      const parts = val.split(token);
      const parent = node.parentNode;
      if (!parent) continue;
      if (parts[0]) parent.insertBefore(doc.createTextNode(parts[0]), node);
      parent.insertBefore(shapeEl, node);
      if (parts.slice(1).join(token)) {
        parent.insertBefore(doc.createTextNode(parts.slice(1).join(token)), node);
      }
      parent.removeChild(node);
    }
  }

  // 1b. Hydrate off-centre anchored floating images with position:absolute on their owning page element
  const floatTokens = Object.keys(preview.floatingImages ?? {});
  if (floatTokens.length > 0) {
    const walker = doc.createTreeWalker(container, NodeFilter.SHOW_TEXT);
    const matches: Array<{node: Text; floatId: string}> = [];
    let current: Node | null;
    while ((current = walker.nextNode()) !== null) {
      const textNode = current as Text;
      const val = textNode.nodeValue ?? '';
      const re = /__(ESZ_FLOAT_IMG_\d+)__/g;
      let m: RegExpExecArray | null;
      while ((m = re.exec(val)) !== null) {
        if (preview.floatingImages[m[1]]) {
          matches.push({node: textNode, floatId: m[1]});
        }
      }
    }

    for (const {node, floatId} of matches) {
      const spec = preview.floatingImages[floatId];
      const anchorEl = node.parentElement;
      const pageEl = anchorEl?.closest<HTMLElement>('section.esz-docx') ?? pages[0];
      if (!pageEl) continue;

      if (!pageEl.style.position) {
        pageEl.style.position = 'relative';
      }

      let topPx = spec.topPx;
      if (spec.relativeFromV === 'paragraph' && anchorEl) {
        // Compute unscaled vertical offset of anchor paragraph relative to its section page
        const liveScale =
          pageEl.offsetWidth > 0 ? pageEl.getBoundingClientRect().width / pageEl.offsetWidth : 1;
        const safeScale = liveScale > 0.05 ? liveScale : 1;
        const pageRect = pageEl.getBoundingClientRect();
        const anchorRect = anchorEl.getBoundingClientRect();
        const anchorTopInPage = (anchorRect.top - pageRect.top) / safeScale;
        topPx = Math.max(0, anchorTopInPage + spec.topPx);
      }

      const img = doc.createElement('img');
      img.src = spec.dataUrl;
      img.alt = spec.alt;
      img.setAttribute('data-testid', 'docx-floating-image');
      img.setAttribute('data-float-id', floatId);
      img.className = 'esz-docx-floating-image';
      img.style.position = 'absolute';
      img.style.left = `${spec.leftPx}px`;
      img.style.top = `${Math.round(topPx * 100) / 100}px`;
      img.style.width = `${spec.widthPx}px`;
      img.style.height = `${spec.heightPx}px`;
      img.style.maxWidth = 'none';
      img.style.pointerEvents = 'none';
      img.style.zIndex = spec.behindDoc ? '0' : '3';
      pageEl.appendChild(img);

      // Strip placeholder token from text node
      if (node.nodeValue) {
        node.nodeValue = node.nodeValue.replace(`__${floatId}__`, '');
      }
    }
  }

  // 2. Replace PAGE / NUMPAGES tokens and apply exact w:pgBorders per rendered page
  const defaultBorderSpec =
    preview.sectionBorders.find((b) => b !== null) ?? null;

  pages.forEach((pageEl, pageIdx) => {
    const pageNum = pageIdx + 1;
    pageEl.setAttribute('data-testid', 'docx-preview-page');
    pageEl.setAttribute('data-page-number', String(pageNum));

    // Stamp PAGE and NUMPAGES tokens inside this page
    const walker = doc.createTreeWalker(pageEl, NodeFilter.SHOW_TEXT);
    let current: Node | null;
    while ((current = walker.nextNode()) !== null) {
      const t = current as Text;
      if (!t.nodeValue) continue;
      if (t.nodeValue.includes('__ESZ_PAGE_NUM__') || t.nodeValue.includes('__ESZ_NUM_PAGES__')) {
        t.nodeValue = t.nodeValue
          .replace(/__ESZ_PAGE_NUM__/g, String(pageNum))
          .replace(/__ESZ_NUM_PAGES__/g, String(Math.max(1, totalPages)));
      }
    }

    // Apply w:pgBorders overlay frame if section defines page borders
    const borderSpec = preview.sectionBorders[pageIdx] ?? defaultBorderSpec;
    if (borderSpec) {
      const shouldApply =
        borderSpec.display === 'allPages' ||
        (borderSpec.display === 'firstPage' && pageIdx === 0) ||
        (borderSpec.display === 'notFirstPage' && pageIdx > 0);

      if (shouldApply) {
        const frame = doc.createElement('div');
        frame.setAttribute('data-testid', 'docx-page-border');
        frame.className = 'esz-docx-page-border';
        frame.style.position = 'absolute';
        frame.style.pointerEvents = 'none';
        frame.style.boxSizing = 'border-box';
        frame.style.zIndex = '2';

        const computeInsetPx = (marginPt: number, side?: DocxBorderSide): number => {
          const spacePt = side?.spacePt ?? 24;
          if (borderSpec.offsetFrom === 'page') {
            return Math.max(2, Math.round(spacePt * PT_TO_PX));
          }
          // offsetFrom === 'text': offset outwards from text margin by spacePt, clamped inside page edge
          const insetPt = Math.max(6, marginPt - spacePt);
          return Math.round(insetPt * PT_TO_PX);
        };

        frame.style.top = `${computeInsetPx(borderSpec.marginsPt.top, borderSpec.top)}px`;
        frame.style.right = `${computeInsetPx(borderSpec.marginsPt.right, borderSpec.right)}px`;
        frame.style.bottom = `${computeInsetPx(borderSpec.marginsPt.bottom, borderSpec.bottom)}px`;
        frame.style.left = `${computeInsetPx(borderSpec.marginsPt.left, borderSpec.left)}px`;

        const formatBorder = (side?: DocxBorderSide) =>
          side ? `${side.szPt}pt ${mapWordBorderStyleToCss(side.val)} ${side.color}` : 'none';

        frame.style.borderTop = formatBorder(borderSpec.top);
        frame.style.borderRight = formatBorder(borderSpec.right);
        frame.style.borderBottom = formatBorder(borderSpec.bottom);
        frame.style.borderLeft = formatBorder(borderSpec.left);

        pageEl.appendChild(frame);
      }
    }
  });

  return totalPages;
}

/**
 * Pure helper to determine which 1-based page is currently visible in a scroll container.
 * Uses the actual rendered page bounds (already scaled or unscaled consistently with scrollTop)
 * and picks the page covering the upper-third viewport reading line (or maximum visible overlap).
 */
export function computeVisiblePageNumber(
  scrollTop: number,
  viewportHeight: number,
  pageBounds: Array<{top: number; height: number}>,
): number {
  if (pageBounds.length <= 1) return 1;

  const viewTop = scrollTop;
  const viewBottom = scrollTop + Math.max(1, viewportHeight);
  const anchorY = viewTop + Math.min(viewportHeight * 0.35, 220);

  // First check if a page directly straddles the reading anchor line
  for (let i = 0; i < pageBounds.length; i++) {
    const p = pageBounds[i];
    if (anchorY >= p.top && anchorY <= p.top + p.height) {
      return i + 1;
    }
  }

  // Otherwise pick the page with the largest vertical intersection with the viewport
  let bestIdx = 0;
  let bestOverlap = -1;
  for (let i = 0; i < pageBounds.length; i++) {
    const p = pageBounds[i];
    const overlap = Math.min(viewBottom, p.top + p.height) - Math.max(viewTop, p.top);
    if (overlap > bestOverlap) {
      bestOverlap = overlap;
      bestIdx = i;
    }
  }
  return bestIdx + 1;
}
