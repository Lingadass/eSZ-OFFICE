import JSZip from 'jszip';
import {canonicalizeWordFontFamily} from './docFonts';
import {parseXml, type XmlNode} from './docxImport';
import {
  computeBlockSignature,
  decodeUtf8Base64,
  escapeXmlAttr,
  escapeXmlText,
  normalizeHex6,
  parseHtmlTree,
  parseInlineCss,
  parseLengthPt,
  parseLengthPx,
  type HtmlNode,
} from './docxOoxmlBridge';
import {
  DOC_MARGIN_PX,
  DOC_PAGE_DIMENSIONS_PX,
  type DocNoteItem,
  type DocPageBorderSide,
  type DocPageLayout,
  type DocTabStop,
} from './docsTypes';

const EMU_PER_PX = 9525;

export interface LosslessSaveStats {
  preservedBlocks: number;
  generatedBlocks: number;
  addedMediaCount: number;
  sectPrPreservedVerbatim: boolean;
}

export interface LosslessSaveResult {
  bytes: Uint8Array;
  stats: LosslessSaveStats;
}

function pxToTwips(px: number): number {
  return Math.round(px * 15);
}

function ptToTwips(pt: number): number {
  return Math.round(pt * 20);
}

function decodeBase64DataUrl(
  dataUrl: string,
): {bytes: Uint8Array; ext: 'png' | 'jpeg' | 'gif' | 'bmp'; contentType: string} | null {
  const m = dataUrl.match(/^data:image\/([a-zA-Z0-9+.-]+);base64,([\s\S]+)$/);
  if (!m) return null;
  const mimeSub = m[1].toLowerCase();
  const ext: 'png' | 'jpeg' | 'gif' | 'bmp' =
    mimeSub === 'jpeg' || mimeSub === 'jpg'
      ? 'jpeg'
      : mimeSub === 'gif'
        ? 'gif'
        : mimeSub === 'bmp'
          ? 'bmp'
          : 'png';
  const contentType = `image/${ext}`;
  try {
    const cleanB64 = m[2].replace(/\s+/g, '');
    if (typeof Buffer !== 'undefined') {
      return {bytes: new Uint8Array(Buffer.from(cleanB64, 'base64')), ext, contentType};
    }
    const bin = atob(cleanB64);
    const arr = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
    return {bytes: arr, ext, contentType};
  } catch {
    return null;
  }
}

function bytesFingerprint(bytes: Uint8Array): string {
  let h1 = 0x811c9dc5;
  let h2 = 0x9e3779b9;
  for (let i = 0; i < bytes.length; i++) {
    h1 ^= bytes[i];
    h1 = Math.imul(h1, 0x01000193);
    h2 ^= bytes[i] + (i & 0xff);
    h2 = Math.imul(h2, 0x85ebca6b);
  }
  return `${bytes.length}:${(h1 >>> 0).toString(16)}:${(h2 >>> 0).toString(16)}`;
}

// ---------------------------------------------------------------------------
// Package Relationship & Media Manager (Mutates rels/content-types ONLY on additions)
// ---------------------------------------------------------------------------

interface PackagePatchContext {
  zip: JSZip;
  nextRelNum: number;
  nextDocPrId: number;
  mediaCounter: number;
  existingMediaByHash: Map<string, string>; // bytesFingerprint -> existing rId
  existingHyperlinkByUrl: Map<string, string>; // url -> existing rId
  newRelationshipsXml: string[];
  newContentTypeExts: Map<string, string>; // ext -> mime
  existingContentTypeExts: Set<string>;
  needsDrawingNamespaces: boolean;
  needsRelNamespace: boolean;
  addedMediaCount: number;
}

async function createPackagePatchContext(
  zip: JSZip,
  origDocXml: string,
): Promise<PackagePatchContext> {
  let maxRelNum = 0;
  const existingMediaByHash = new Map<string, string>();
  const existingHyperlinkByUrl = new Map<string, string>();

  const relsStr = (await zip.file('word/_rels/document.xml.rels')?.async('string')) ?? '';
  const relTagRe = /<Relationship\b([^>]*?)\/?\s*>/gi;
  let rm: RegExpExecArray | null;
  while ((rm = relTagRe.exec(relsStr)) !== null) {
    const attrs = rm[1];
    const id = attrs.match(/\bId="([^"]+)"/i)?.[1] ?? '';
    const target = attrs.match(/\bTarget="([^"]+)"/i)?.[1] ?? '';
    const type = attrs.match(/\bType="([^"]+)"/i)?.[1] ?? '';
    const targetMode = attrs.match(/\bTargetMode="([^"]+)"/i)?.[1] ?? '';

    const numMatch = id.match(/^rId(\d+)$/i);
    if (numMatch) {
      maxRelNum = Math.max(maxRelNum, Number(numMatch[1]));
    }

    if (id && target && type.endsWith('/image')) {
      const normalizedPath = target.startsWith('/')
        ? target.slice(1)
        : target.startsWith('word/')
          ? target
          : `word/${target.replace(/^\.\//, '')}`;
      const mediaFile = zip.file(normalizedPath);
      if (mediaFile) {
        const mediaBytes = await mediaFile.async('uint8array');
        existingMediaByHash.set(bytesFingerprint(mediaBytes), id);
      }
    } else if (id && target && targetMode.toLowerCase() === 'external' && type.endsWith('/hyperlink')) {
      existingHyperlinkByUrl.set(target, id);
    }
  }

  // Find max wp:docPr id so new drawings never collide
  let maxDocPrId = 100;
  const docPrRe = /<wp:docPr\b[^>]*\bid="(\d+)"/gi;
  let dm: RegExpExecArray | null;
  while ((dm = docPrRe.exec(origDocXml)) !== null) {
    maxDocPrId = Math.max(maxDocPrId, Number(dm[1]));
  }

  // Track existing Default extensions in [Content_Types].xml
  const existingContentTypeExts = new Set<string>();
  const ctStr = (await zip.file('[Content_Types].xml')?.async('string')) ?? '';
  const ctDefRe = /<Default\b[^>]*\bExtension="([^"]+)"/gi;
  let cm: RegExpExecArray | null;
  while ((cm = ctDefRe.exec(ctStr)) !== null) {
    existingContentTypeExts.add(cm[1].toLowerCase());
  }

  return {
    zip,
    nextRelNum: maxRelNum + 1,
    nextDocPrId: maxDocPrId + 1,
    mediaCounter: 1,
    existingMediaByHash,
    existingHyperlinkByUrl,
    newRelationshipsXml: [],
    newContentTypeExts: new Map(),
    existingContentTypeExts,
    needsDrawingNamespaces: false,
    needsRelNamespace: false,
    addedMediaCount: 0,
  };
}

function ensureImageRelationship(
  ctx: PackagePatchContext,
  dataUrl: string,
): {rId: string; ext: string} | null {
  const decoded = decodeBase64DataUrl(dataUrl);
  if (!decoded) return null;

  const fp = bytesFingerprint(decoded.bytes);
  const existingRId = ctx.existingMediaByHash.get(fp);
  if (existingRId) {
    return {rId: existingRId, ext: decoded.ext};
  }

  // Allocate unique filename inside word/media/
  let fileName = `esz_image_${ctx.mediaCounter++}.${decoded.ext}`;
  while (ctx.zip.file(`word/media/${fileName}`)) {
    fileName = `esz_image_${ctx.mediaCounter++}.${decoded.ext}`;
  }
  ctx.zip.file(`word/media/${fileName}`, decoded.bytes);

  const rId = `rId${ctx.nextRelNum++}`;
  ctx.existingMediaByHash.set(fp, rId);
  ctx.addedMediaCount += 1;

  ctx.newRelationshipsXml.push(
    `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${fileName}"/>`,
  );

  if (!ctx.existingContentTypeExts.has(decoded.ext)) {
    ctx.newContentTypeExts.set(decoded.ext, decoded.contentType);
    ctx.existingContentTypeExts.add(decoded.ext);
  }

  return {rId, ext: decoded.ext};
}

function ensureHyperlinkRelationship(ctx: PackagePatchContext, href: string): string {
  const existing = ctx.existingHyperlinkByUrl.get(href);
  if (existing) return existing;

  const rId = `rId${ctx.nextRelNum++}`;
  ctx.existingHyperlinkByUrl.set(href, rId);
  ctx.needsRelNamespace = true;
  ctx.newRelationshipsXml.push(
    `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${escapeXmlAttr(href)}" TargetMode="External"/>`,
  );
  return rId;
}

// ---------------------------------------------------------------------------
// Direct OOXML Serializers (for Edited or Newly Inserted Blocks ONLY)
// ---------------------------------------------------------------------------

interface RunFormat {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  sub?: boolean;
  superScript?: boolean;
  colorHex?: string;
  highlightHex?: string;
  fontFamily?: string;
  fontSizeHalfPts?: number;
  allCaps?: boolean;
  smallCaps?: boolean;
  characterSpacingTwips?: number;
}

function buildRPrXml(fmt: RunFormat): string {
  const parts: string[] = [];
  if (fmt.fontFamily) {
    const f = escapeXmlAttr(fmt.fontFamily);
    parts.push(`<w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:cs="${f}"/>`);
  }
  if (fmt.bold) parts.push('<w:b/><w:bCs/>');
  if (fmt.italic) parts.push('<w:i/><w:iCs/>');
  if (fmt.allCaps) parts.push('<w:caps/>');
  if (fmt.smallCaps) parts.push('<w:smallCaps/>');
  if (fmt.strike) parts.push('<w:strike/>');
  if (fmt.colorHex) parts.push(`<w:color w:val="${fmt.colorHex}"/>`);
  if (fmt.characterSpacingTwips) {
    parts.push(`<w:spacing w:val="${fmt.characterSpacingTwips}"/>`);
  }
  if (fmt.fontSizeHalfPts) {
    parts.push(`<w:sz w:val="${fmt.fontSizeHalfPts}"/><w:szCs w:val="${fmt.fontSizeHalfPts}"/>`);
  }
  if (fmt.underline) parts.push('<w:u w:val="single"/>');
  if (fmt.highlightHex) {
    parts.push(`<w:shd w:val="clear" w:color="auto" w:fill="${fmt.highlightHex}"/>`);
  }
  if (fmt.superScript) parts.push('<w:vertAlign w:val="superscript"/>');
  else if (fmt.sub) parts.push('<w:vertAlign w:val="subscript"/>');

  return parts.length > 0 ? `<w:rPr>${parts.join('')}</w:rPr>` : '';
}

function buildDrawingXmlForImage(
  imgEl: HtmlNode,
  ctx: PackagePatchContext,
): string {
  // If this image carried its original DrawingML XML and wasn't modified, keep exact wp:anchor/wp:inline!
  const origDrawingB64 = imgEl.attrs['data-ooxml-drawing'];
  const origSig = imgEl.attrs['data-ooxml-sig'];
  if (origDrawingB64 && (!origSig || origSig === computeBlockSignature(imgEl))) {
    const decoded = decodeUtf8Base64(origDrawingB64);
    if (decoded && decoded.includes('<w:drawing')) {
      return decoded;
    }
  }

  const src = imgEl.attrs.src ?? '';
  const rel = ensureImageRelationship(ctx, src);
  if (!rel) return '';

  ctx.needsDrawingNamespaces = true;
  ctx.needsRelNamespace = true;

  const css = parseInlineCss(imgEl.attrs.style);
  const wPx = Math.max(
    16,
    Math.round(
      Number(imgEl.attrs['data-width']) ||
        parseLengthPx(css.width) ||
        Number(imgEl.attrs.width) ||
        240,
    ),
  );
  const hPx = Math.max(
    16,
    Math.round(
      Number(imgEl.attrs['data-height']) ||
        parseLengthPx(css.height) ||
        Number(imgEl.attrs.height) ||
        120,
    ),
  );
  const cx = wPx * EMU_PER_PX;
  const cy = hPx * EMU_PER_PX;
  const docPrId = ctx.nextDocPrId++;
  const alt = escapeXmlAttr(imgEl.attrs.alt || `Image ${docPrId}`);
  const floating = imgEl.attrs['data-floating'] === 'true';
  const wrapRaw = (imgEl.attrs['data-wrap'] || '').toLowerCase();

  const picXml =
    `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
    `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:nvPicPr><pic:cNvPr id="${docPrId}" name="${alt}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${rel.rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic>`;

  if (floating) {
    const relH = escapeXmlAttr(imgEl.attrs['data-rel-h'] || 'column');
    const relV = escapeXmlAttr(imgEl.attrs['data-rel-v'] || 'paragraph');
    const align = (imgEl.attrs['data-align'] || 'left').toLowerCase();
    const floatLeftPx = imgEl.attrs['data-float-left'] !== undefined ? Number(imgEl.attrs['data-float-left']) : null;
    const floatTopPx = imgEl.attrs['data-float-top'] !== undefined ? Number(imgEl.attrs['data-float-top']) : 0;
    const distT = Math.max(0, Number(imgEl.attrs['data-margin-top']) || 0) * EMU_PER_PX;
    const distB = Math.max(0, Number(imgEl.attrs['data-margin-bottom']) || 0) * EMU_PER_PX;

    const posHInner =
      floatLeftPx !== null && Number.isFinite(floatLeftPx) && align !== 'center' && align !== 'right'
        ? `<wp:posOffset>${Math.round(floatLeftPx * EMU_PER_PX)}</wp:posOffset>`
        : `<wp:align>${align === 'center' || align === 'right' ? align : 'left'}</wp:align>`;
    const posVInner = `<wp:posOffset>${Math.round((Number.isFinite(floatTopPx) ? floatTopPx : 0) * EMU_PER_PX)}</wp:posOffset>`;
    const wrapXml =
      wrapRaw === 'none'
        ? '<wp:wrapNone/>'
        : wrapRaw === 'square'
          ? '<wp:wrapSquare wrapText="bothSides"/>'
          : wrapRaw === 'tight'
            ? '<wp:wrapTight wrapText="bothSides"/>'
            : '<wp:wrapTopAndBottom/>';

    return (
      `<w:drawing>` +
      `<wp:anchor distT="${distT}" distB="${distB}" distL="114300" distR="114300" simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
      `<wp:simplePos x="0" y="0"/>` +
      `<wp:positionH relativeFrom="${relH}">${posHInner}</wp:positionH>` +
      `<wp:positionV relativeFrom="${relV}">${posVInner}</wp:positionV>` +
      `<wp:extent cx="${cx}" cy="${cy}"/>` +
      `<wp:effectExtent l="0" t="0" r="0" b="0"/>` +
      `${wrapXml}` +
      `<wp:docPr id="${docPrId}" name="${alt}"/>` +
      `<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
      `${picXml}` +
      `</wp:anchor></w:drawing>`
    );
  }

  return (
    `<w:drawing>` +
    `<wp:inline distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="${cx}" cy="${cy}"/>` +
    `<wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `<wp:docPr id="${docPrId}" name="${alt}"/>` +
    `<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
    `${picXml}` +
    `</wp:inline></w:drawing>`
  );
}

function serializeInlineRunsToOoxml(
  node: HtmlNode,
  inherited: RunFormat,
  ctx: PackagePatchContext,
  out: string[],
): void {
  if (node.type === 'text') {
    if (!node.text) return;
    const rPr = buildRPrXml(inherited);
    const needsSpacePreserve = /^\s|\s$|\s{2,}/.test(node.text);
    const spaceAttr = needsSpacePreserve ? ' xml:space="preserve"' : '';
    out.push(`<w:r>${rPr}<w:t${spaceAttr}>${escapeXmlText(node.text)}</w:t></w:r>`);
    return;
  }

  const tag = node.tag;

  if (tag === 'br') {
    out.push('<w:r><w:br/></w:r>');
    return;
  }

  if (tag === 'span' && node.attrs['data-doc-tab'] === 'true') {
    const rPr = buildRPrXml(inherited);
    out.push(`<w:r>${rPr}<w:tab/></w:r>`);
    return;
  }

  if (tag === 'img') {
    const drawingXml = buildDrawingXmlForImage(node, ctx);
    if (drawingXml) {
      out.push(`<w:r>${drawingXml}</w:r>`);
    }
    return;
  }

  const next: RunFormat = {...inherited};
  if (tag === 'strong' || tag === 'b') next.bold = true;
  else if (tag === 'em' || tag === 'i') next.italic = true;
  else if (tag === 'u') next.underline = true;
  else if (tag === 's' || tag === 'del') next.strike = true;
  else if (tag === 'sub') next.sub = true;
  else if (tag === 'sup') next.superScript = true;
  else if (tag === 'code') next.fontFamily = 'Courier New';

  const css = parseInlineCss(node.attrs.style);
  const rawFont = node.attrs['data-doc-font'] || css['font-family'];
  if (rawFont) {
    const canonical = canonicalizeWordFontFamily(rawFont);
    if (canonical) next.fontFamily = canonical;
  }
  if (css['font-size']) {
    const pt = parseLengthPt(css['font-size']);
    if (pt && pt > 0) next.fontSizeHalfPts = Math.max(8, Math.round(pt * 2));
  }
  if (css.color) {
    const hex = normalizeHex6(css.color);
    if (hex) next.colorHex = hex;
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

  if (tag === 'a' && node.attrs.href && !node.attrs.href.startsWith('#')) {
    const rId = ensureHyperlinkRelationship(ctx, node.attrs.href);
    const linkRuns: string[] = [];
    for (const child of node.children) {
      serializeInlineRunsToOoxml(child, next, ctx, linkRuns);
    }
    if (linkRuns.length > 0) {
      out.push(`<w:hyperlink r:id="${rId}">${linkRuns.join('')}</w:hyperlink>`);
    }
    return;
  }

  for (const child of node.children) {
    serializeInlineRunsToOoxml(child, next, ctx, out);
  }
}

function extractNumPrXmlFromSource(ooxmlSourceB64: string | undefined): string | null {
  if (!ooxmlSourceB64) return null;
  const raw = decodeUtf8Base64(ooxmlSourceB64);
  if (!raw) return null;
  const m = raw.match(/<w:numPr\b[\s\S]*?<\/w:numPr>/);
  return m ? m[0] : null;
}

const HEADING_STYLE_FALLBACK: Record<string, string> = {
  h1: 'Heading1',
  h2: 'Heading2',
  h3: 'Heading3',
  h4: 'Heading4',
};

function mapAlignmentToOoxml(alignKey: string): string | null {
  switch (alignKey.toLowerCase()) {
    case 'center':
      return 'center';
    case 'right':
      return 'right';
    case 'justify':
      return 'both';
    case 'left':
      return 'left';
    default:
      return null;
  }
}

function buildOoxmlParagraph(
  el: HtmlNode,
  ctx: PackagePatchContext,
  options?: {
    extraIndentLeftTwips?: number;
    preservedNumPrXml?: string | null;
    defaultBold?: boolean;
  },
): string {
  const css = parseInlineCss(el.attrs.style);
  const pPrParts: string[] = [];

  // 1. w:pStyle
  const styleId = el.attrs['data-style-id'] || HEADING_STYLE_FALLBACK[el.tag] || null;
  if (styleId) {
    pPrParts.push(`<w:pStyle w:val="${escapeXmlAttr(styleId)}"/>`);
  }

  // 2. w:keepNext / w:keepLines
  if (el.attrs['data-keep-next'] === 'true' || el.tag in HEADING_STYLE_FALLBACK) {
    pPrParts.push('<w:keepNext/>');
  }
  if (el.attrs['data-keep-lines'] === 'true') {
    pPrParts.push('<w:keepLines/>');
  }

  // 3. w:numPr (for edited list items preserving original numbering.xml binding)
  const numPrXml =
    options?.preservedNumPrXml ?? extractNumPrXmlFromSource(el.attrs['data-ooxml-source']);
  if (numPrXml) {
    pPrParts.push(numPrXml);
  }

  // 4. w:tabs (explicit tab stops)
  const rawTabStops = el.attrs['data-tab-stops'];
  if (rawTabStops) {
    try {
      const stops = JSON.parse(rawTabStops) as DocTabStop[];
      if (Array.isArray(stops) && stops.length > 0) {
        const tabsXml = stops
          .filter((s) => s && typeof s.posTwips === 'number' && Number.isFinite(s.posTwips))
          .map((s) => {
            const valAttr = escapeXmlAttr(s.val || 'left');
            const leaderAttr =
              s.leader && s.leader !== 'none' ? ` w:leader="${escapeXmlAttr(s.leader)}"` : '';
            return `<w:tab w:val="${valAttr}"${leaderAttr} w:pos="${Math.round(s.posTwips)}"/>`;
          })
          .join('');
        if (tabsXml) pPrParts.push(`<w:tabs>${tabsXml}</w:tabs>`);
      }
    } catch {
      // Ignore malformed tab stops JSON
    }
  }

  // 5. w:spacing
  const beforePt =
    el.attrs['data-space-before'] !== undefined
      ? Number(el.attrs['data-space-before'])
      : parseLengthPt(css['margin-top']);
  const afterPt =
    el.attrs['data-space-after'] !== undefined
      ? Number(el.attrs['data-space-after'])
      : parseLengthPt(css['margin-bottom']);
  const lineSpacingRaw = el.attrs['data-line-spacing'] ?? css['line-height'];

  const spacingAttrs: string[] = [];
  if (beforePt !== undefined && Number.isFinite(beforePt)) {
    spacingAttrs.push(`w:before="${Math.max(0, ptToTwips(beforePt))}"`);
  }
  if (afterPt !== undefined && Number.isFinite(afterPt)) {
    spacingAttrs.push(`w:after="${Math.max(0, ptToTwips(afterPt))}"`);
  }
  if (lineSpacingRaw) {
    const trimmed = lineSpacingRaw.trim().toLowerCase();
    if (trimmed.endsWith('pt')) {
      const pt = parseLengthPt(trimmed);
      if (pt && pt > 0) {
        spacingAttrs.push(`w:line="${ptToTwips(pt)}"`, 'w:lineRule="exact"');
      }
    } else {
      const mult = Number(trimmed);
      if (Number.isFinite(mult) && mult >= 0.6 && mult <= 4) {
        spacingAttrs.push(`w:line="${Math.round(mult * 240)}"`, 'w:lineRule="auto"');
      }
    }
  }
  if (spacingAttrs.length > 0) {
    pPrParts.push(`<w:spacing ${spacingAttrs.join(' ')}/>`);
  }

  // 6. w:ind
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
  const indAttrs: string[] = [];
  if (totalLeftTwips > 0) indAttrs.push(`w:left="${totalLeftTwips}"`);
  if (indentRightPx > 0) indAttrs.push(`w:right="${pxToTwips(indentRightPx)}"`);
  if (firstLinePx > 0) indAttrs.push(`w:firstLine="${pxToTwips(firstLinePx)}"`);
  else if (firstLinePx < 0) indAttrs.push(`w:hanging="${pxToTwips(Math.abs(firstLinePx))}"`);
  if (indAttrs.length > 0) {
    pPrParts.push(`<w:ind ${indAttrs.join(' ')}/>`);
  }

  // 7. w:jc
  const jcVal = mapAlignmentToOoxml(css['text-align'] ?? '');
  if (jcVal && jcVal !== 'left') {
    pPrParts.push(`<w:jc w:val="${jcVal}"/>`);
  } else if (css['text-align']?.toLowerCase() === 'left') {
    pPrParts.push('<w:jc w:val="left"/>');
  }

  const pPrXml = pPrParts.length > 0 ? `<w:pPr>${pPrParts.join('')}</w:pPr>` : '';

  const runXmlParts: string[] = [];
  serializeInlineRunsToOoxml(el, {bold: options?.defaultBold}, ctx, runXmlParts);

  return `<w:p>${pPrXml}${runXmlParts.join('')}</w:p>`;
}

function buildCssBorderOoxml(sideName: string, raw: string | undefined): string {
  if (!raw) return '';
  const trimmed = raw.trim().toLowerCase();
  if (trimmed === 'none' || trimmed === '0') {
    return `<w:${sideName} w:val="none" w:sz="0" w:space="0" w:color="auto"/>`;
  }
  const parts = raw.trim().split(/\s+/);
  let widthPt = 0.5;
  let val = 'single';
  let color = '000000';
  for (const p of parts) {
    const pt = parseLengthPt(p);
    if (pt !== undefined && /\d/.test(p)) {
      widthPt = Math.max(0.25, pt);
      continue;
    }
    const low = p.toLowerCase();
    if (low === 'double' || low === 'triple') {
      val = 'double';
      continue;
    }
    if (low === 'dashed') {
      val = 'dashed';
      continue;
    }
    if (low === 'dotted') {
      val = 'dotted';
      continue;
    }
    if (low === 'solid') {
      val = 'single';
      continue;
    }
    const hex = normalizeHex6(p);
    if (hex) color = hex;
  }
  const sz = Math.max(2, Math.round(widthPt * 8));
  return `<w:${sideName} w:val="${val}" w:sz="${sz}" w:space="0" w:color="${color}"/>`;
}

function buildOoxmlTable(tblEl: HtmlNode, ctx: PackagePatchContext): string {
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
      if (cell.attrs.colwidth) {
        for (const piece of cell.attrs.colwidth.split(',')) {
          const n = Number(piece.trim());
          if (Number.isFinite(n) && n > 0) colWidthsPx.push(n);
        }
      }
    }
  }

  const tblAlignRaw = (tblEl.attrs['data-table-align'] || '').toLowerCase();
  const tableWidthPx =
    Number(tblEl.attrs['data-table-width']) ||
    (colWidthsPx.length > 0 ? colWidthsPx.reduce((a, b) => a + b, 0) : 0);

  const tblPrParts: string[] = [];
  if (tableWidthPx > 0) {
    tblPrParts.push(`<w:tblW w:w="${pxToTwips(tableWidthPx)}" w:type="dxa"/>`);
  } else {
    tblPrParts.push('<w:tblW w:w="5000" w:type="pct"/>');
  }
  const tblJc = mapAlignmentToOoxml(tblAlignRaw);
  if (tblJc && tblJc !== 'both') {
    tblPrParts.push(`<w:jc w:val="${tblJc}"/>`);
  }
  if (colWidthsPx.length > 0) {
    tblPrParts.push('<w:tblLayout w:type="fixed"/>');
  }

  const tblGridXml =
    colWidthsPx.length > 0
      ? `<w:tblGrid>${colWidthsPx.map((w) => `<w:gridCol w:w="${pxToTwips(w)}"/>`).join('')}</w:tblGrid>`
      : '';

  // Track vertical merges across rows so rowspan > 1 emits continuation <w:vMerge/> cells
  const activeRowSpans = new Map<number, {remaining: number; colSpan: number; widthTwips: number}>();
  const rowsXml: string[] = [];

  for (const trEl of trEls) {
    const cells = trEl.children.filter(
      (c) => c.type === 'element' && (c.tag === 'td' || c.tag === 'th'),
    );
    const isHeaderRow = cells.length > 0 && cells.every((c) => c.tag === 'th');
    const trPrXml = `<w:trPr><w:cantSplit/>${isHeaderRow ? '<w:tblHeader/>' : ''}</w:trPr>`;

    const rowCellsXml: string[] = [];
    let colIdx = 0;

    const flushVerticalContinuations = () => {
      while (activeRowSpans.has(colIdx)) {
        const span = activeRowSpans.get(colIdx)!;
        const contPr: string[] = [];
        if (span.widthTwips > 0) contPr.push(`<w:tcW w:w="${span.widthTwips}" w:type="dxa"/>`);
        if (span.colSpan > 1) contPr.push(`<w:gridSpan w:val="${span.colSpan}"/>`);
        contPr.push('<w:vMerge/>');
        rowCellsXml.push(`<w:tc><w:tcPr>${contPr.join('')}</w:tcPr><w:p/></w:tc>`);
        span.remaining -= 1;
        if (span.remaining <= 0) activeRowSpans.delete(colIdx);
        colIdx += span.colSpan;
      }
    };

    for (const cellEl of cells) {
      flushVerticalContinuations();
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
      if (!cellWidthPx) cellWidthPx = parseLengthPx(css.width) ?? 0;
      const cellWidthTwips = cellWidthPx > 0 ? pxToTwips(cellWidthPx) : 0;

      if (rowSpan > 1) {
        activeRowSpans.set(colIdx, {
          remaining: rowSpan - 1,
          colSpan,
          widthTwips: cellWidthTwips,
        });
      }

      const tcPrParts: string[] = [];
      if (cellWidthTwips > 0) tcPrParts.push(`<w:tcW w:w="${cellWidthTwips}" w:type="dxa"/>`);
      if (colSpan > 1) tcPrParts.push(`<w:gridSpan w:val="${colSpan}"/>`);
      if (rowSpan > 1) tcPrParts.push('<w:vMerge w:val="restart"/>');

      const bTop = buildCssBorderOoxml('top', cellEl.attrs['data-border-top'] || css['border-top']);
      const bLeft = buildCssBorderOoxml('left', cellEl.attrs['data-border-left'] || css['border-left']);
      const bBottom = buildCssBorderOoxml(
        'bottom',
        cellEl.attrs['data-border-bottom'] || css['border-bottom'],
      );
      const bRight = buildCssBorderOoxml(
        'right',
        cellEl.attrs['data-border-right'] || css['border-right'],
      );
      if (bTop || bLeft || bBottom || bRight) {
        tcPrParts.push(`<w:tcBorders>${bTop}${bLeft}${bBottom}${bRight}</w:tcBorders>`);
      }

      const bgHex = normalizeHex6(cellEl.attrs['data-bg-color'] || css['background-color']);
      if (bgHex) {
        tcPrParts.push(`<w:shd w:val="clear" w:color="auto" w:fill="${bgHex}"/>`);
      }

      const vAlignRaw = (cellEl.attrs['data-valign'] || css['vertical-align'] || '').toLowerCase();
      if (vAlignRaw === 'middle' || vAlignRaw === 'center') {
        tcPrParts.push('<w:vAlign w:val="center"/>');
      } else if (vAlignRaw === 'bottom') {
        tcPrParts.push('<w:vAlign w:val="bottom"/>');
      } else if (vAlignRaw === 'top') {
        tcPrParts.push('<w:vAlign w:val="top"/>');
      }

      const innerXml: string[] = [];
      for (const child of cellEl.children) {
        if (child.type === 'text') {
          if (child.text.trim()) {
            const rPr = cellEl.tag === 'th' ? '<w:rPr><w:b/></w:rPr>' : '';
            innerXml.push(`<w:p><w:r>${rPr}<w:t xml:space="preserve">${escapeXmlText(child.text)}</w:t></w:r></w:p>`);
          }
        } else if (child.tag === 'table') {
          innerXml.push(buildOoxmlTable(child, ctx));
        } else {
          innerXml.push(
            buildOoxmlParagraph(child, ctx, {defaultBold: cellEl.tag === 'th'}),
          );
        }
      }
      if (innerXml.length === 0) innerXml.push('<w:p/>');

      const tcPrXml = tcPrParts.length > 0 ? `<w:tcPr>${tcPrParts.join('')}</w:tcPr>` : '';
      rowCellsXml.push(`<w:tc>${tcPrXml}${innerXml.join('')}</w:tc>`);
      colIdx += colSpan;
    }
    flushVerticalContinuations();

    rowsXml.push(`<w:tr>${trPrXml}${rowCellsXml.join('')}</w:tr>`);
  }

  return `<w:tbl><w:tblPr>${tblPrParts.join('')}</w:tblPr>${tblGridXml}${rowsXml.join('')}</w:tbl>`;
}

function buildOoxmlTextBoxTable(boxEl: HtmlNode, ctx: PackagePatchContext): string {
  const widthPx = Math.max(80, Number(boxEl.attrs['data-box-width']) || 440);
  const widthTwips = pxToTwips(widthPx);
  const align = mapAlignmentToOoxml(boxEl.attrs['data-box-align'] || 'center') ?? 'center';
  const fillHex = normalizeHex6(boxEl.attrs['data-box-fill']);
  const borderColorHex = normalizeHex6(boxEl.attrs['data-box-border-color']);
  const borderStyleRaw = (boxEl.attrs['data-box-border-style'] || 'solid').toLowerCase();
  const borderWidthPx = Number(boxEl.attrs['data-box-border-width']) || (borderColorHex ? 1 : 0);
  const vAlignRaw = (boxEl.attrs['data-box-valign'] || 'top').toLowerCase();

  const bVal =
    !borderColorHex || borderStyleRaw === 'none' || borderWidthPx <= 0
      ? 'none'
      : borderStyleRaw === 'dashed'
        ? 'dashed'
        : borderStyleRaw === 'dotted'
          ? 'dotted'
          : borderStyleRaw === 'double'
            ? 'double'
            : 'single';
  const bSz = bVal === 'none' ? 0 : Math.max(4, Math.round(borderWidthPx * 0.75 * 8));
  const bCol = bVal === 'none' ? 'auto' : borderColorHex ?? '000000';
  const bordersXml = `<w:tcBorders>` +
    `<w:top w:val="${bVal}" w:sz="${bSz}" w:space="0" w:color="${bCol}"/>` +
    `<w:left w:val="${bVal}" w:sz="${bSz}" w:space="0" w:color="${bCol}"/>` +
    `<w:bottom w:val="${bVal}" w:sz="${bSz}" w:space="0" w:color="${bCol}"/>` +
    `<w:right w:val="${bVal}" w:sz="${bSz}" w:space="0" w:color="${bCol}"/>` +
    `</w:tcBorders>`;

  const tcPrParts = [
    `<w:tcW w:w="${widthTwips}" w:type="dxa"/>`,
    bordersXml,
    ...(fillHex ? [`<w:shd w:val="clear" w:color="auto" w:fill="${fillHex}"/>`] : []),
    `<w:tcMar><w:top w:w="100" w:type="dxa"/><w:left w:w="140" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/><w:right w:w="140" w:type="dxa"/></w:tcMar>`,
    ...(vAlignRaw === 'middle'
      ? ['<w:vAlign w:val="center"/>']
      : vAlignRaw === 'bottom'
        ? ['<w:vAlign w:val="bottom"/>']
        : []),
  ];

  const innerXml: string[] = [];
  for (const child of boxEl.children) {
    if (child.type === 'element') {
      if (child.tag === 'table') innerXml.push(buildOoxmlTable(child, ctx));
      else innerXml.push(buildOoxmlParagraph(child, ctx));
    } else if (child.text.trim()) {
      innerXml.push(`<w:p><w:r><w:t xml:space="preserve">${escapeXmlText(child.text)}</w:t></w:r></w:p>`);
    }
  }
  if (innerXml.length === 0) innerXml.push('<w:p/>');

  return (
    `<w:tbl><w:tblPr><w:tblW w:w="${widthTwips}" w:type="dxa"/><w:jc w:val="${align}"/><w:tblLayout w:type="fixed"/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="${widthTwips}"/></w:tblGrid>` +
    `<w:tr><w:trPr><w:cantSplit/></w:trPr><w:tc><w:tcPr>${tcPrParts.join('')}</w:tcPr>${innerXml.join('')}</w:tc></w:tr></w:tbl>`
  );
}

function buildOoxmlImageParagraph(imgEl: HtmlNode, ctx: PackagePatchContext): string {
  const drawingXml = buildDrawingXmlForImage(imgEl, ctx);
  if (!drawingXml) return '';
  const align = (imgEl.attrs['data-align'] || 'left').toLowerCase();
  const floating = imgEl.attrs['data-floating'] === 'true';
  const pPrParts: string[] = [];
  if (!floating && (align === 'center' || align === 'right')) {
    pPrParts.push(`<w:jc w:val="${align}"/>`);
  }
  const marginTopPx = Number(imgEl.attrs['data-margin-top']);
  const marginBottomPx = Number(imgEl.attrs['data-margin-bottom']);
  if (!floating && (Number.isFinite(marginTopPx) || Number.isFinite(marginBottomPx))) {
    const spAttrs: string[] = [];
    if (Number.isFinite(marginTopPx) && marginTopPx > 0) {
      spAttrs.push(`w:before="${pxToTwips(marginTopPx)}"`);
    }
    if (Number.isFinite(marginBottomPx) && marginBottomPx > 0) {
      spAttrs.push(`w:after="${pxToTwips(marginBottomPx)}"`);
    }
    if (spAttrs.length > 0) pPrParts.push(`<w:spacing ${spAttrs.join(' ')}/>`);
  }
  const pPrXml = pPrParts.length > 0 ? `<w:pPr>${pPrParts.join('')}</w:pPr>` : '';
  return `<w:p>${pPrXml}<w:r>${drawingXml}</w:r></w:p>`;
}

// ---------------------------------------------------------------------------
// Flattened Block Stream & Group-Aware Lossless Reconstruction
// ---------------------------------------------------------------------------

interface FlattenedBlock {
  node: HtmlNode;
  listContext?: {
    listType: 'ul' | 'ol';
    depth: number;
    siblingNumPrXml: string | null;
  };
}

function findListSiblingNumPr(listEl: HtmlNode): string | null {
  for (const li of listEl.children) {
    if (li.type !== 'element' || li.tag !== 'li') continue;
    for (const c of li.children) {
      if (c.type === 'element' && c.tag === 'p') {
        const found = extractNumPrXmlFromSource(c.attrs['data-ooxml-source']);
        if (found) return found;
      }
    }
  }
  return null;
}

function flattenHtmlBlocks(
  nodes: HtmlNode[],
  out: FlattenedBlock[],
  listDepth = 0,
  listType?: 'ul' | 'ol',
  siblingNumPrXml: string | null = null,
): void {
  for (const node of nodes) {
    if (node.type === 'text') {
      if (node.text.trim()) {
        out.push({node});
      }
      continue;
    }

    const tag = node.tag;

    if (
      node.attrs['data-page-break'] === 'true' ||
      node.attrs['data-doc-textbox'] === 'true' ||
      node.attrs['data-doc-locked'] === 'true' ||
      node.attrs['data-doc-equation'] === 'true'
    ) {
      out.push({node});
      continue;
    }

    if (tag === 'ul' || tag === 'ol') {
      const numPr = findListSiblingNumPr(node);
      for (const child of node.children) {
        if (child.type === 'element' && child.tag === 'li') {
          flattenHtmlBlocks([child], out, listDepth, tag, numPr);
        }
      }
      continue;
    }

    if (tag === 'li') {
      let emittedPrimary = false;
      for (const child of node.children) {
        if (child.type === 'element' && (child.tag === 'ul' || child.tag === 'ol')) {
          flattenHtmlBlocks([child], out, listDepth + 1, child.tag, findListSiblingNumPr(child));
        } else if (child.type === 'element' && (child.tag === 'p' || child.tag === 'img' || child.tag === 'div')) {
          out.push({
            node: child,
            listContext: listType
              ? {listType, depth: listDepth, siblingNumPrXml}
              : undefined,
          });
          emittedPrimary = true;
        } else if (child.type === 'element' && child.tag === 'table') {
          out.push({node: child});
          emittedPrimary = true;
        }
      }
      if (!emittedPrimary) {
        out.push({
          node,
          listContext: listType ? {listType, depth: listDepth, siblingNumPrXml} : undefined,
        });
      }
      continue;
    }

    if (tag === 'div' || tag === 'section' || tag === 'article' || tag === 'body') {
      const hasBlockChildren = node.children.some(
        (c) =>
          c.type === 'element' &&
          ['p', 'h1', 'h2', 'h3', 'h4', 'table', 'ul', 'ol', 'div', 'blockquote', 'img'].includes(
            c.tag,
          ),
      );
      if (hasBlockChildren) {
        flattenHtmlBlocks(node.children, out, listDepth, listType, siblingNumPrXml);
      } else if (node.children.length > 0) {
        out.push({node});
      }
      continue;
    }

    out.push({node});
  }
}

function serializeEditedBlock(item: FlattenedBlock, ctx: PackagePatchContext): string {
  const node = item.node;
  if (node.type === 'text') {
    return `<w:p><w:r><w:t xml:space="preserve">${escapeXmlText(node.text)}</w:t></w:r></w:p>`;
  }

  const tag = node.tag;

  if (node.attrs['data-page-break'] === 'true') {
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
  }

  if (node.attrs['data-doc-textbox'] === 'true') {
    return buildOoxmlTextBoxTable(node, ctx);
  }

  if (node.attrs['data-doc-locked'] === 'true') {
    const imgSrc = node.attrs['data-locked-image'];
    if (imgSrc) {
      const syntheticImg: HtmlNode = {
        type: 'element',
        tag: 'img',
        attrs: {
          src: imgSrc,
          'data-width': node.attrs['data-width'] || '240',
          'data-height': node.attrs['data-height'] || '120',
          'data-align': 'center',
          alt: node.attrs['data-locked-label'] || 'Locked Element',
        },
        children: [],
        text: '',
      };
      const imgPara = buildOoxmlImageParagraph(syntheticImg, ctx);
      if (imgPara) return imgPara;
    }
    const label = node.attrs['data-locked-label'] || 'Locked Word Element';
    return `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:i/><w:color w:val="64748B"/></w:rPr><w:t>[${escapeXmlText(label)}]</w:t></w:r></w:p>`;
  }

  if (node.attrs['data-doc-equation'] === 'true') {
    return buildOoxmlParagraph(node, ctx);
  }

  if (tag === 'img') {
    return buildOoxmlImageParagraph(node, ctx);
  }

  if (tag === 'table') {
    return buildOoxmlTable(node, ctx);
  }

  if (tag === 'blockquote') {
    const childParas = node.children.filter((c) => c.type === 'element' && c.tag === 'p');
    if (childParas.length > 0) {
      return childParas.map((p) => buildOoxmlParagraph(p, ctx, {extraIndentLeftTwips: 720})).join('');
    }
    return buildOoxmlParagraph(node, ctx, {extraIndentLeftTwips: 720});
  }

  return buildOoxmlParagraph(node, ctx, {
    preservedNumPrXml: item.listContext
      ? extractNumPrXmlFromSource(node.attrs['data-ooxml-source']) ??
        item.listContext.siblingNumPrXml
      : undefined,
  });
}

// ---------------------------------------------------------------------------
// Section Properties (<w:sectPr>) Preservation & Targeted Patching
// ---------------------------------------------------------------------------

function xmlChildAttr(node: XmlNode | undefined, childTag: string, attrName: string): string | undefined {
  if (!node) return undefined;
  const child = node.children.find(
    (c) => c.type === 'element' && (c.tag === childTag || c.localName === childTag),
  );
  if (!child) return undefined;
  if (child.attrs[attrName] !== undefined) return child.attrs[attrName];
  const suffix = `:${attrName}`;
  for (const [k, v] of Object.entries(child.attrs)) {
    if (k === attrName || k.endsWith(suffix)) return v;
  }
  return undefined;
}

function buildOoxmlPageBorderSide(sideName: string, side: DocPageBorderSide | undefined): string {
  if (!side) return '';
  const val = side.val === 'double' || side.val === 'triple'
    ? 'double'
    : side.val === 'dashed'
      ? 'dashed'
      : side.val === 'dotted'
        ? 'dotted'
        : side.val === 'none' || side.val === 'nil'
          ? 'none'
          : 'single';
  const sz = Math.max(4, Math.round(side.szPt * 8));
  const space = Math.max(1, Math.min(31, Math.round(side.spacePt)));
  const color = normalizeHex6(side.color) ?? '000000';
  return `<w:${sideName} w:val="${val}" w:sz="${sz}" w:space="${space}" w:color="${color}"/>`;
}

function resolveSectionPropertiesXml(
  origSectPrNode: XmlNode | undefined,
  layout: DocPageLayout,
): {xml: string; preservedVerbatim: boolean} {
  if (!origSectPrNode?.rawXml) {
    // Fallback if original document somehow lacked top-level w:sectPr
    const w = layout.customPageWidthPx ? pxToTwips(layout.customPageWidthPx) : 11906;
    const h = layout.customPageHeightPx ? pxToTwips(layout.customPageHeightPx) : 16838;
    const m = layout.customMarginsPx
      ? {
          top: pxToTwips(layout.customMarginsPx.top),
          right: pxToTwips(layout.customMarginsPx.right),
          bottom: pxToTwips(layout.customMarginsPx.bottom),
          left: pxToTwips(layout.customMarginsPx.left),
        }
      : {top: 1440, right: 1440, bottom: 1440, left: 1440};
    return {
      xml: `<w:sectPr><w:pgSz w:w="${w}" w:h="${h}"/><w:pgMar w:top="${m.top}" w:right="${m.right}" w:bottom="${m.bottom}" w:left="${m.left}" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>`,
      preservedVerbatim: false,
    };
  }

  const origXml = origSectPrNode.rawXml;

  // Compare original w:pgSz and w:pgMar against layout (with 20-twip tolerance for px rounding)
  const origW = Number(xmlChildAttr(origSectPrNode, 'w:pgSz', 'w') ?? 11906);
  const origH = Number(xmlChildAttr(origSectPrNode, 'w:pgSz', 'h') ?? 16838);
  const origOrient =
    (xmlChildAttr(origSectPrNode, 'w:pgSz', 'orient') ?? '').toLowerCase() === 'landscape' ||
    origW > origH
      ? 'landscape'
      : 'portrait';

  const presetDims = DOC_PAGE_DIMENSIONS_PX[layout.size] ?? DOC_PAGE_DIMENSIONS_PX.a4;
  const rawDesiredW = layout.customPageWidthPx
    ? pxToTwips(layout.customPageWidthPx)
    : layout.size === 'a4'
      ? 11906
      : layout.size === 'letter' || layout.size === 'legal'
        ? 12240
        : pxToTwips(presetDims.width);
  const rawDesiredH = layout.customPageHeightPx
    ? pxToTwips(layout.customPageHeightPx)
    : layout.size === 'a4'
      ? 16838
      : layout.size === 'legal'
        ? 20160
        : 15840;
  const desiredW =
    layout.customPageWidthPx != null
      ? rawDesiredW
      : layout.orientation === 'landscape'
        ? rawDesiredH
        : rawDesiredW;
  const desiredH =
    layout.customPageHeightPx != null
      ? rawDesiredH
      : layout.orientation === 'landscape'
        ? rawDesiredW
        : rawDesiredH;

  const origTop = Number(xmlChildAttr(origSectPrNode, 'w:pgMar', 'top') ?? 1440);
  const origRight = Number(xmlChildAttr(origSectPrNode, 'w:pgMar', 'right') ?? 1440);
  const origBottom = Number(xmlChildAttr(origSectPrNode, 'w:pgMar', 'bottom') ?? 1440);
  const origLeft = Number(xmlChildAttr(origSectPrNode, 'w:pgMar', 'left') ?? 1440);
  const origHeader = xmlChildAttr(origSectPrNode, 'w:pgMar', 'header') ?? '708';
  const origFooter = xmlChildAttr(origSectPrNode, 'w:pgMar', 'footer') ?? '708';
  const origGutter = xmlChildAttr(origSectPrNode, 'w:pgMar', 'gutter') ?? '0';

  const presetMargins = DOC_MARGIN_PX[layout.margins] ?? DOC_MARGIN_PX.normal;
  const desiredTop = layout.customMarginsPx
    ? pxToTwips(layout.customMarginsPx.top)
    : layout.margins === 'narrow'
      ? 720
      : 1440;
  const desiredBottom = layout.customMarginsPx
    ? pxToTwips(layout.customMarginsPx.bottom)
    : layout.margins === 'narrow'
      ? 720
      : 1440;
  const desiredLeft = layout.customMarginsPx
    ? pxToTwips(layout.customMarginsPx.left)
    : layout.margins === 'narrow'
      ? 720
      : layout.margins === 'moderate'
        ? 1080
        : layout.margins === 'wide'
          ? 2160
          : pxToTwips(presetMargins.left);
  const desiredRight = layout.customMarginsPx
    ? pxToTwips(layout.customMarginsPx.right)
    : desiredLeft;

  const origCols = Math.max(1, Number(xmlChildAttr(origSectPrNode, 'w:cols', 'num') ?? 1));
  const sizeChanged =
    origOrient !== layout.orientation ||
    Math.abs(origW - desiredW) > 25 ||
    Math.abs(origH - desiredH) > 25;
  const marginsChanged =
    Math.abs(origTop - desiredTop) > 25 ||
    Math.abs(origRight - desiredRight) > 25 ||
    Math.abs(origBottom - desiredBottom) > 25 ||
    Math.abs(origLeft - desiredLeft) > 25;
  const colsChanged = origCols !== layout.columns;

  const hasOrigPgBorders = Boolean(
    origSectPrNode.children.find((c) => c.type === 'element' && c.localName === 'pgBorders'),
  );
  const wantsCustomBorder = Boolean(layout.pageBorderConfig);
  const wantsPresetBorder = !wantsCustomBorder && layout.pageBorder !== 'none';
  const borderChanged =
    (!hasOrigPgBorders && (wantsCustomBorder || wantsPresetBorder)) ||
    (hasOrigPgBorders && !wantsCustomBorder && layout.pageBorder === 'none');

  if (!sizeChanged && !marginsChanged && !colsChanged && !borderChanged) {
    return {xml: origXml, preservedVerbatim: true};
  }

  let patched = origXml;

  if (sizeChanged) {
    const orientAttr = layout.orientation === 'landscape' ? ' w:orient="landscape"' : '';
    const newPgSz = `<w:pgSz w:w="${desiredW}" w:h="${desiredH}"${orientAttr}/>`;
    if (/<w:pgSz\b[^>]*\/>/.test(patched)) {
      patched = patched.replace(/<w:pgSz\b[^>]*\/>/, newPgSz);
    } else {
      patched = patched.replace(/<\/w:sectPr>$/, `${newPgSz}</w:sectPr>`);
    }
  }

  if (marginsChanged) {
    const newPgMar = `<w:pgMar w:top="${desiredTop}" w:right="${desiredRight}" w:bottom="${desiredBottom}" w:left="${desiredLeft}" w:header="${origHeader}" w:footer="${origFooter}" w:gutter="${origGutter}"/>`;
    if (/<w:pgMar\b[^>]*\/>/.test(patched)) {
      patched = patched.replace(/<w:pgMar\b[^>]*\/>/, newPgMar);
    } else {
      patched = patched.replace(/<\/w:sectPr>$/, `${newPgMar}</w:sectPr>`);
    }
  }

  if (colsChanged) {
    const newCols =
      layout.columns > 1
        ? `<w:cols w:num="${layout.columns}" w:space="708"/>`
        : '<w:cols w:space="708"/>';
    if (/<w:cols\b[^>]*\/>/.test(patched)) {
      patched = patched.replace(/<w:cols\b[^>]*\/>/, newCols);
    } else {
      patched = patched.replace(/<\/w:sectPr>$/, `${newCols}</w:sectPr>`);
    }
  }

  if (borderChanged) {
    patched = patched.replace(/<w:pgBorders\b[\s\S]*?<\/w:pgBorders>/g, '');
    if (layout.pageBorderConfig) {
      const cfg = layout.pageBorderConfig;
      const sides = [
        buildOoxmlPageBorderSide('top', cfg.top),
        buildOoxmlPageBorderSide('left', cfg.left),
        buildOoxmlPageBorderSide('bottom', cfg.bottom),
        buildOoxmlPageBorderSide('right', cfg.right),
      ].join('');
      const newBorders = `<w:pgBorders w:offsetFrom="${cfg.offsetFrom}" w:display="${cfg.display}">${sides}</w:pgBorders>`;
      patched = patched.replace(/<\/w:sectPr>$/, `${newBorders}</w:sectPr>`);
    } else if (layout.pageBorder !== 'none') {
      const bVal = layout.pageBorder === 'double' ? 'double' : 'single';
      const newBorders =
        `<w:pgBorders w:offsetFrom="page" w:display="allPages">` +
        `<w:top w:val="${bVal}" w:sz="12" w:space="24" w:color="000000"/>` +
        `<w:left w:val="${bVal}" w:sz="12" w:space="24" w:color="000000"/>` +
        `<w:bottom w:val="${bVal}" w:sz="12" w:space="24" w:color="000000"/>` +
        `<w:right w:val="${bVal}" w:sz="12" w:space="24" w:color="000000"/>` +
        `</w:pgBorders>`;
      patched = patched.replace(/<\/w:sectPr>$/, `${newBorders}</w:sectPr>`);
    }
  }

  return {xml: patched, preservedVerbatim: false};
}

function ensureDocumentNamespaces(prefixXml: string, ctx: PackagePatchContext): string {
  if (!ctx.needsDrawingNamespaces && !ctx.needsRelNamespace) return prefixXml;
  const docTagMatch = prefixXml.match(/<w:document\b[^>]*>/);
  if (!docTagMatch) return prefixXml;

  let docTag = docTagMatch[0];
  const required: Array<[string, string]> = [];
  if (ctx.needsRelNamespace || ctx.needsDrawingNamespaces) {
    required.push(['xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships']);
  }
  if (ctx.needsDrawingNamespaces) {
    required.push(
      ['xmlns:wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'],
      ['xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main'],
      ['xmlns:pic', 'http://schemas.openxmlformats.org/drawingml/2006/picture'],
    );
  }

  const toAdd: string[] = [];
  for (const [attrName, uri] of required) {
    if (!new RegExp(`\\b${attrName}\\s*=`).test(docTag)) {
      toAdd.push(`${attrName}="${uri}"`);
    }
  }
  if (toAdd.length === 0) return prefixXml;

  const updatedDocTag = docTag.replace(/>$/, ` ${toAdd.join(' ')}>`);
  return prefixXml.replace(docTag, updatedDocTag);
}

// ---------------------------------------------------------------------------
// Main Entrypoint: Lossless Package-Preserving .docx Save
// ---------------------------------------------------------------------------

/**
 * Lossless `.docx` save that starts from the original `.docx` ZIP package and rewrites
 * ONLY `word/document.xml` (plus new `word/media/*`, `word/_rels/document.xml.rels`, and
 * `[Content_Types].xml` entries when new images or hyperlinks were inserted).
 *
 * - Keeps original `styles.xml`, `numbering.xml`, `theme`, `fontTable`, `settings`,
 *   `headers`, `footers`, `docProps`, and existing relationships/media untouched.
 * - Writes every unedited imported paragraph, table, page break, anchored image, and
 *   content control back byte-for-byte unchanged from its stored `data-ooxml-source`.
 * - Generates schema-valid OOXML only for edited or newly inserted blocks.
 */
export async function patchOriginalDocxPackage(
  html: string,
  layout: DocPageLayout,
  notes: DocNoteItem[],
  originalDocxBuffer: ArrayBuffer | Uint8Array,
): Promise<LosslessSaveResult | null> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(originalDocxBuffer);
  } catch {
    return null;
  }

  const origDocXml = await zip.file('word/document.xml')?.async('string');
  if (!origDocXml) return null;

  const bodyOpenMatch = origDocXml.match(/<w:body\b[^>]*>/);
  const bodyCloseIdx = origDocXml.lastIndexOf('</w:body>');
  if (!bodyOpenMatch || bodyOpenMatch.index === undefined || bodyCloseIdx < 0) {
    return null;
  }

  const bodyContentStart = bodyOpenMatch.index + bodyOpenMatch[0].length;
  const origPrefix = origDocXml.slice(0, bodyContentStart);
  const origSuffix = origDocXml.slice(bodyCloseIdx);

  // Extract direct-child document-level <w:sectPr> accurately via XML tree
  const parsedDoc = parseXml(origDocXml);
  const parsedBody = parsedDoc?.children.find(
    (c) => c.type === 'element' && (c.tag === 'w:body' || c.localName === 'body'),
  );
  const directSectPrs = (parsedBody?.children ?? []).filter(
    (c) => c.type === 'element' && (c.tag === 'w:sectPr' || c.localName === 'sectPr'),
  );
  const origSectPrNode = directSectPrs[directSectPrs.length - 1];

  const patchCtx = await createPackagePatchContext(zip, origDocXml);
  const htmlRoot = parseHtmlTree(html);
  const flattened: FlattenedBlock[] = [];
  flattenHtmlBlocks(htmlRoot.children, flattened);

  const bodyXmlParts: string[] = [];
  const stats: LosslessSaveStats = {
    preservedBlocks: 0,
    generatedBlocks: 0,
    addedMediaCount: 0,
    sectPrPreservedVerbatim: true,
  };

  let i = 0;
  while (i < flattened.length) {
    const item = flattened[i];
    const el = item.node;

    if (el.type === 'element') {
      const groupId = el.attrs['data-ooxml-group'];
      const gi = el.attrs['data-ooxml-gi'] !== undefined ? Number(el.attrs['data-ooxml-gi']) : -1;
      const gc = el.attrs['data-ooxml-gc'] !== undefined ? Number(el.attrs['data-ooxml-gc']) : 0;
      const srcB64 = el.attrs['data-ooxml-source'];

      if (groupId && srcB64 && gi === 0 && gc >= 1 && i + gc <= flattened.length) {
        let groupIntact = true;
        for (let k = 0; k < gc; k++) {
          const member = flattened[i + k].node;
          if (
            member.type !== 'element' ||
            member.attrs['data-ooxml-group'] !== groupId ||
            Number(member.attrs['data-ooxml-gi']) !== k ||
            Number(member.attrs['data-ooxml-gc']) !== gc
          ) {
            groupIntact = false;
            break;
          }
          const storedSig = member.attrs['data-ooxml-sig'];
          if (!storedSig || computeBlockSignature(member) !== storedSig) {
            groupIntact = false;
            break;
          }
        }

        if (groupIntact) {
          const decodedRawXml = decodeUtf8Base64(srcB64);
          if (decodedRawXml) {
            bodyXmlParts.push(decodedRawXml);
            stats.preservedBlocks += 1;
            i += gc;
            continue;
          }
        }
      }
    }

    const generated = serializeEditedBlock(item, patchCtx);
    if (generated) {
      bodyXmlParts.push(generated);
      stats.generatedBlocks += 1;
    }
    i += 1;
  }

  // Append newly authored notes if any exist outside the original package's footnotes/endnotes
  if (notes.length > 0 && !zip.file('word/footnotes.xml') && !zip.file('word/endnotes.xml')) {
    bodyXmlParts.push(
      '<w:p><w:pPr><w:pStyle w:val="Heading3"/><w:keepNext/></w:pPr><w:r><w:t>Footnotes &amp; Endnotes</w:t></w:r></w:p>',
    );
    for (const n of notes) {
      bodyXmlParts.push(
        `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">[${n.number}] (${escapeXmlText(n.kind)}) </w:t></w:r><w:r><w:t xml:space="preserve">${escapeXmlText(n.text)}</w:t></w:r></w:p>`,
      );
    }
  }

  if (bodyXmlParts.length === 0) {
    bodyXmlParts.push('<w:p/>');
  }

  const sectPrResolved = resolveSectionPropertiesXml(origSectPrNode, layout);
  stats.sectPrPreservedVerbatim = sectPrResolved.preservedVerbatim;
  stats.addedMediaCount = patchCtx.addedMediaCount;

  const finalPrefix = ensureDocumentNamespaces(origPrefix, patchCtx);
  const nextDocXml = `${finalPrefix}${bodyXmlParts.join('')}${sectPrResolved.xml}${origSuffix}`;
  zip.file('word/document.xml', nextDocXml);

  // Only rewrite word/_rels/document.xml.rels if new media or hyperlinks were added
  if (patchCtx.newRelationshipsXml.length > 0) {
    const existingRels =
      (await zip.file('word/_rels/document.xml.rels')?.async('string')) ??
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
    const updatedRels = existingRels.replace(
      /<\/Relationships>/i,
      `${patchCtx.newRelationshipsXml.join('')}</Relationships>`,
    );
    zip.file('word/_rels/document.xml.rels', updatedRels);
  }

  // Only rewrite [Content_Types].xml if a brand-new image extension was added
  if (patchCtx.newContentTypeExts.size > 0) {
    const existingCt = await zip.file('[Content_Types].xml')?.async('string');
    if (existingCt) {
      const additions = Array.from(patchCtx.newContentTypeExts.entries())
        .map(([ext, mime]) => `<Default Extension="${ext}" ContentType="${mime}"/>`)
        .join('');
      const updatedCt = existingCt.replace(/(<Types\b[^>]*>)/i, `$1${additions}`);
      zip.file('[Content_Types].xml', updatedCt);
    }
  }

  const bytes = await zip.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: {level: 6},
  });

  return {bytes, stats};
}
