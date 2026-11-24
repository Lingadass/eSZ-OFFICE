import {
  PDFDocument,
  StandardFonts,
  degrees,
  rgb,
  type PDFFont,
  type PDFPage,
} from 'pdf-lib';
import type {
  PdfAnnotationItem,
  PdfHeaderFooterConfig,
  PdfMetadataInfo,
  PdfPageNumberConfig,
  PdfPageTextClearance,
  PdfWatermarkConfig,
} from './pdfTypes';
import {computeOverlayLayoutForPage} from './pdfOverlayLayout';

export function parsePageRange(spec: string, totalPages: number): number[] {
  if (totalPages <= 0) return [];
  const trimmed = spec.trim();
  if (!trimmed) {
    return Array.from({length: totalPages}, (_, i) => i);
  }
  const indices: number[] = [];
  const seen = new Set<number>();
  const parts = trimmed.split(',');

  for (const rawPart of parts) {
    const part = rawPart.trim().toLowerCase();
    if (!part) continue;
    const dashIdx = part.indexOf('-');
    if (dashIdx !== -1) {
      const startRaw = part.slice(0, dashIdx).trim();
      const endRaw = part.slice(dashIdx + 1).trim();
      const start = startRaw === 'z' ? totalPages : Number.parseInt(startRaw, 10);
      const end = endRaw === 'z' ? totalPages : Number.parseInt(endRaw, 10);
      if (Number.isFinite(start) && Number.isFinite(end)) {
        const step = start <= end ? 1 : -1;
        for (let p = start; step > 0 ? p <= end : p >= end; p += step) {
          if (p >= 1 && p <= totalPages) {
            const idx = p - 1;
            if (!seen.has(idx)) {
              seen.add(idx);
              indices.push(idx);
            }
          }
        }
      }
    } else {
      const p = part === 'z' ? totalPages : Number.parseInt(part, 10);
      if (Number.isFinite(p) && p >= 1 && p <= totalPages) {
        const idx = p - 1;
        if (!seen.has(idx)) {
          seen.add(idx);
          indices.push(idx);
        }
      }
    }
  }
  return indices;
}

export function splitIntoChunkSpecs(totalPages: number, pagesPerFile: number): number[][] {
  if (totalPages <= 0) return [];
  const chunk = Math.max(1, Math.floor(pagesPerFile));
  const groups: number[][] = [];
  for (let i = 0; i < totalPages; i += chunk) {
    const group: number[] = [];
    for (let j = i; j < Math.min(totalPages, i + chunk); j++) {
      group.push(j);
    }
    groups.push(group);
  }
  return groups;
}

export function hexToRgbColor(hex: string) {
  const clean = hex.replace('#', '').trim();
  if (clean.length === 3) {
    const r = Number.parseInt(clean[0] + clean[0], 16) / 255;
    const g = Number.parseInt(clean[1] + clean[1], 16) / 255;
    const b = Number.parseInt(clean[2] + clean[2], 16) / 255;
    if (!Number.isNaN(r) && !Number.isNaN(g) && !Number.isNaN(b)) {
      return rgb(r, g, b);
    }
  }
  if (clean.length === 6) {
    const r = Number.parseInt(clean.slice(0, 2), 16) / 255;
    const g = Number.parseInt(clean.slice(2, 4), 16) / 255;
    const b = Number.parseInt(clean.slice(4, 6), 16) / 255;
    if (!Number.isNaN(r) && !Number.isNaN(g) && !Number.isNaN(b)) {
      return rgb(r, g, b);
    }
  }
  return rgb(0.11, 0.31, 0.85);
}

function wrapLine(text: string, font: PDFFont, fontSize: number, maxWidth: number): string[] {
  const words = text.split(/\s+/);
  const lines: string[] = [];
  let current = '';

  for (const word of words) {
    const candidate = current ? `${current} ${word}` : word;
    if (font.widthOfTextAtSize(candidate, fontSize) <= maxWidth) {
      current = candidate;
    } else {
      if (current) lines.push(current);
      current = word;
    }
  }
  if (current) lines.push(current);
  return lines.length > 0 ? lines : [''];
}

export async function createBlankPdfBytes(
  title = 'Untitled PDF',
  bodyText = '',
): Promise<Uint8Array> {
  if (bodyText.trim()) {
    return textToPdfBytes(bodyText, title);
  }
  const doc = await PDFDocument.create();
  doc.setTitle(title);
  doc.setCreator('eSZ OFFICE');
  doc.addPage([595.28, 841.89]); // A4 portrait
  return doc.save();
}

export async function textToPdfBytes(text: string, title = 'Document'): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  doc.setTitle(title);
  doc.setCreator('eSZ OFFICE');
  const font = await doc.embedFont(StandardFonts.Helvetica);
  const boldFont = await doc.embedFont(StandardFonts.HelveticaBold);

  const pageWidth = 595.28;
  const pageHeight = 841.89;
  const margin = 52;
  const maxWidth = pageWidth - margin * 2;
  const fontSize = 11;
  const lineHeight = 16.5;

  let page: PDFPage = doc.addPage([pageWidth, pageHeight]);
  let y = pageHeight - margin;

  if (title.trim()) {
    const safeTitle = title.replace(/[^\x20-\x7E]/g, ' ');
    page.drawText(safeTitle, {
      x: margin,
      y,
      size: 16,
      font: boldFont,
      color: rgb(0.06, 0.09, 0.16),
    });
    y -= 28;
  }

  const paragraphs = text.replace(/\r\n/g, '\n').split('\n');
  for (const para of paragraphs) {
    const sanitized = para.replace(/[^\x20-\x7E]/g, ' ');
    const lines = sanitized.trim() ? wrapLine(sanitized, font, fontSize, maxWidth) : [''];
    for (const line of lines) {
      if (y < margin + lineHeight) {
        page = doc.addPage([pageWidth, pageHeight]);
        y = pageHeight - margin;
      }
      if (line) {
        page.drawText(line, {
          x: margin,
          y,
          size: fontSize,
          font,
          color: rgb(0.12, 0.16, 0.23),
        });
      }
      y -= lineHeight;
    }
  }

  return doc.save();
}

export async function mergePdfBytes(inputs: Uint8Array[]): Promise<Uint8Array> {
  const outDoc = await PDFDocument.create();
  outDoc.setCreator('eSZ OFFICE');
  for (const bytes of inputs) {
    const src = await PDFDocument.load(bytes, {ignoreEncryption: true});
    const copied = await outDoc.copyPages(src, src.getPageIndices());
    for (const page of copied) {
      outDoc.addPage(page);
    }
  }
  return outDoc.save();
}

export async function extractPdfPagesBytes(
  pdfBytes: Uint8Array,
  pageIndices: number[],
): Promise<Uint8Array> {
  const src = await PDFDocument.load(pdfBytes, {ignoreEncryption: true});
  const out = await PDFDocument.create();
  out.setCreator('eSZ OFFICE');
  const valid = pageIndices.filter((i) => i >= 0 && i < src.getPageCount());
  const copied = await out.copyPages(src, valid);
  for (const p of copied) out.addPage(p);
  return out.save();
}

export async function reorderAndTransformPdfBytes(
  pdfBytes: Uint8Array,
  pageOrder: Array<{
    sourceIndex: number | 'blank-a4' | 'blank-letter';
    extraRotationDeg?: number;
  }>,
): Promise<Uint8Array> {
  const src = await PDFDocument.load(pdfBytes, {ignoreEncryption: true});
  const out = await PDFDocument.create();
  out.setTitle(src.getTitle() ?? '');
  out.setAuthor(src.getAuthor() ?? '');
  out.setSubject(src.getSubject() ?? '');
  out.setKeywords(src.getKeywords() ? [src.getKeywords()!] : []);
  out.setCreator(src.getCreator() || 'eSZ OFFICE');

  for (const spec of pageOrder) {
    if (spec.sourceIndex === 'blank-a4') {
      out.addPage([595.28, 841.89]);
      continue;
    }
    if (spec.sourceIndex === 'blank-letter') {
      out.addPage([612, 792]);
      continue;
    }
    if (spec.sourceIndex >= 0 && spec.sourceIndex < src.getPageCount()) {
      const [copied] = await out.copyPages(src, [spec.sourceIndex]);
      if (spec.extraRotationDeg) {
        const current = copied.getRotation().angle;
        copied.setRotation(degrees((current + spec.extraRotationDeg + 360) % 360));
      }
      out.addPage(copied);
    }
  }
  return out.save();
}

export async function flattenPdfBytes(pdfBytes: Uint8Array): Promise<Uint8Array> {
  const doc = await PDFDocument.load(pdfBytes, {ignoreEncryption: true});
  try {
    const form = doc.getForm();
    form.flatten();
  } catch {
    // Document may not contain an AcroForm
  }
  return doc.save();
}

export async function readPdfMetadata(pdfBytes: Uint8Array): Promise<PdfMetadataInfo> {
  const doc = await PDFDocument.load(pdfBytes, {ignoreEncryption: true});
  return {
    title: doc.getTitle() ?? '',
    author: doc.getAuthor() ?? '',
    subject: doc.getSubject() ?? '',
    keywords: doc.getKeywords() ?? '',
    creator: doc.getCreator() ?? 'eSZ OFFICE',
  };
}

export async function updatePdfMetadata(
  pdfBytes: Uint8Array,
  meta: PdfMetadataInfo,
): Promise<Uint8Array> {
  const doc = await PDFDocument.load(pdfBytes, {ignoreEncryption: true});
  doc.setTitle(meta.title);
  doc.setAuthor(meta.author);
  doc.setSubject(meta.subject);
  doc.setKeywords(
    meta.keywords
      .split(',')
      .map((k) => k.trim())
      .filter(Boolean),
  );
  doc.setCreator(meta.creator || 'eSZ OFFICE');
  return doc.save();
}

function dataUrlToUint8Array(dataUrl: string): {bytes: Uint8Array; isPng: boolean} {
  const [header, base64] = dataUrl.split(',');
  const binary = atob(base64 || '');
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return {bytes, isPng: header.includes('image/png')};
}

export async function bakeAnnotationsAndOverlays(
  pdfBytes: Uint8Array,
  annotations: PdfAnnotationItem[],
  extras?: {
    watermark?: PdfWatermarkConfig;
    headerFooter?: PdfHeaderFooterConfig;
    pageNumbers?: PdfPageNumberConfig;
    pageClearances?: Array<PdfPageTextClearance | null>;
  },
): Promise<Uint8Array> {
  const sourceDoc = await PDFDocument.load(pdfBytes, {ignoreEncryption: true});
  const sourcePages = sourceDoc.getPages();
  const totalPages = sourcePages.length;

  // Pre-compute safe-margin layouts to determine if any page needs vector reframing
  const preliminaryLayouts = sourcePages.map((srcPage, idx) => {
    const {width, height} = srcPage.getSize();
    return computeOverlayLayoutForPage({
      pageIndex: idx,
      totalPages,
      pageWidth: width,
      pageHeight: height,
      watermark: extras?.watermark,
      headerFooter: extras?.headerFooter,
      pageNumbers: extras?.pageNumbers,
      clearance: extras?.pageClearances?.[idx] ?? null,
    });
  });

  const anyReframe = preliminaryLayouts.some((l) => l.reframe.needed);

  let doc: PDFDocument;
  if (anyReframe) {
    doc = await PDFDocument.create();
    const title = sourceDoc.getTitle();
    const author = sourceDoc.getAuthor();
    const subject = sourceDoc.getSubject();
    if (title) doc.setTitle(title);
    if (author) doc.setAuthor(author);
    if (subject) doc.setSubject(subject);
    doc.setCreator('eSZ OFFICE');

    for (let idx = 0; idx < totalPages; idx++) {
      const srcPage = sourcePages[idx];
      const layout = preliminaryLayouts[idx];
      if (layout.reframe.needed) {
        const {width, height} = srcPage.getSize();
        const newPage = doc.addPage([width, height]);
        newPage.drawRectangle({
          x: 0,
          y: 0,
          width,
          height,
          color: rgb(1, 1, 1),
        });
        try {
          const embedded = await doc.embedPage(srcPage);
          newPage.drawPage(embedded, {
            x: layout.reframe.offsetX,
            y: layout.reframe.offsetY,
            xScale: layout.reframe.scale,
            yScale: layout.reframe.scale,
          });
        } catch {
          // Fallback: copy page un-embedded if source page stream is empty
          const [copied] = await doc.copyPages(sourceDoc, [idx]);
          doc.removePage(doc.getPageCount() - 1);
          doc.addPage(copied);
        }
      } else {
        const [copied] = await doc.copyPages(sourceDoc, [idx]);
        doc.addPage(copied);
      }
    }
  } else {
    doc = sourceDoc;
  }

  const font = await doc.embedFont(StandardFonts.Helvetica);
  const boldFont = await doc.embedFont(StandardFonts.HelveticaBold);
  const pages = doc.getPages();

  const measureWidth = (text: string, fontSize: number, bold = false) =>
    (bold ? boldFont : font).widthOfTextAtSize(text, fontSize);

  for (let pageIdx = 0; pageIdx < totalPages; pageIdx++) {
    const page = pages[pageIdx];
    const {width, height} = page.getSize();

    // 1. Page-level annotations
    const pageAnns = annotations.filter((a) => a.pageIndex === pageIdx);
    for (const ann of pageAnns) {
      const px = ann.x * width;
      const pw = Math.max(4, ann.width * width);
      const ph = Math.max(4, ann.height * height);
      const py = height - (ann.y * height + ph);
      const color = hexToRgbColor(ann.color);

      if (ann.tool === 'highlight') {
        page.drawRectangle({
          x: px,
          y: py,
          width: pw,
          height: ph,
          color,
          opacity: ann.opacity ?? 0.35,
        });
      } else if (ann.tool === 'underline') {
        page.drawLine({
          start: {x: px, y: py + 2},
          end: {x: px + pw, y: py + 2},
          thickness: Math.max(1.5, ann.strokeWidth || 2),
          color,
        });
      } else if (ann.tool === 'strikethrough') {
        page.drawLine({
          start: {x: px, y: py + ph / 2},
          end: {x: px + pw, y: py + ph / 2},
          thickness: Math.max(1.5, ann.strokeWidth || 2),
          color,
        });
      } else if (ann.tool === 'rect') {
        page.drawRectangle({
          x: px,
          y: py,
          width: pw,
          height: ph,
          borderColor: color,
          borderWidth: Math.max(1.5, ann.strokeWidth || 2),
        });
      } else if (ann.tool === 'ellipse') {
        page.drawEllipse({
          x: px + pw / 2,
          y: py + ph / 2,
          xScale: pw / 2,
          yScale: ph / 2,
          borderColor: color,
          borderWidth: Math.max(1.5, ann.strokeWidth || 2),
        });
      } else if (ann.tool === 'arrow') {
        const startX = px;
        const startY = height - ann.y * height;
        const endX = (ann.x + ann.width) * width;
        const endY = height - (ann.y + ann.height) * height;
        page.drawLine({
          start: {x: startX, y: startY},
          end: {x: endX, y: endY},
          thickness: Math.max(2, ann.strokeWidth || 2),
          color,
        });
      } else if (ann.tool === 'freehand' && ann.points && ann.points.length > 1) {
        for (let i = 1; i < ann.points.length; i++) {
          const prev = ann.points[i - 1];
          const curr = ann.points[i];
          page.drawLine({
            start: {x: prev.x * width, y: height - prev.y * height},
            end: {x: curr.x * width, y: height - curr.y * height},
            thickness: Math.max(1.5, ann.strokeWidth || 2),
            color,
            opacity: ann.opacity ?? 1,
          });
        }
      } else if (ann.tool === 'text' && ann.text) {
        const safe = ann.text.replace(/[^\x20-\x7E]/g, ' ');
        page.drawText(safe, {
          x: px,
          y: height - ann.y * height - 14,
          size: 12,
          font,
          color,
        });
      } else if (ann.tool === 'sticky') {
        const noteText = (ann.text || 'Note').replace(/[^\x20-\x7E]/g, ' ');
        const boxW = Math.min(180, Math.max(110, pw));
        const boxH = Math.min(90, Math.max(46, ph));
        const boxY = Math.max(12, height - ann.y * height - boxH);
        page.drawRectangle({
          x: px,
          y: boxY,
          width: boxW,
          height: boxH,
          color: rgb(1, 0.96, 0.7),
          borderColor: rgb(0.85, 0.65, 0.08),
          borderWidth: 1,
        });
        page.drawText(noteText.slice(0, 80), {
          x: px + 6,
          y: boxY + boxH - 16,
          size: 9,
          font,
          color: rgb(0.15, 0.15, 0.15),
          maxWidth: boxW - 12,
        });
      } else if (ann.tool === 'stamp' && ann.text) {
        const label = ann.text.toUpperCase().replace(/[^\x20-\x7E]/g, ' ');
        const boxW = Math.max(110, pw);
        const boxH = Math.max(32, ph);
        const textW = boldFont.widthOfTextAtSize(label, 13);
        page.drawRectangle({
          x: px,
          y: py,
          width: boxW,
          height: boxH,
          borderColor: color,
          borderWidth: 2,
          opacity: 0.85,
        });
        page.drawText(label, {
          x: px + Math.max(8, (boxW - textW) / 2),
          y: py + Math.max(10, (boxH - 10) / 2),
          size: 13,
          font: boldFont,
          color,
        });
      } else if ((ann.tool === 'image' || ann.tool === 'signature') && ann.imageDataUrl) {
        try {
          const {bytes, isPng} = dataUrlToUint8Array(ann.imageDataUrl);
          const embedded = isPng ? await doc.embedPng(bytes) : await doc.embedJpg(bytes);
          page.drawImage(embedded, {
            x: px,
            y: py,
            width: pw,
            height: ph,
          });
        } catch {
          // Skip invalid image data URL
        }
      } else if (ann.tool === 'redact') {
        page.drawRectangle({
          x: px,
          y: py,
          width: pw,
          height: ph,
          color: rgb(0, 0, 0),
          opacity: 1,
        });
      }
    }

    // 2-4. Compute exact safe-margin overlay coordinates with true embedded font metrics
    const overlayLayout = computeOverlayLayoutForPage({
      pageIndex: pageIdx,
      totalPages,
      pageWidth: width,
      pageHeight: height,
      watermark: extras?.watermark,
      headerFooter: extras?.headerFooter,
      pageNumbers: extras?.pageNumbers,
      clearance: extras?.pageClearances?.[pageIdx] ?? null,
      measureWidth,
    });

    // 2. Watermark
    if (overlayLayout.watermarkItem) {
      const wmItem = overlayLayout.watermarkItem;
      page.drawText(wmItem.text, {
        x: wmItem.pdfX,
        y: wmItem.pdfY,
        size: wmItem.fontSize,
        font: boldFont,
        color: hexToRgbColor(wmItem.color),
        opacity: wmItem.opacity,
        rotate: degrees(wmItem.rotationDeg),
      });
    }

    // 3. Header & Footer (safe margins, true left/center/right alignment, non-overlapping slots)
    for (const item of [...overlayLayout.headerItems, ...overlayLayout.footerItems]) {
      page.drawText(item.text, {
        x: item.x,
        y: item.y,
        size: item.fontSize,
        font,
        color: rgb(0.32, 0.37, 0.45),
      });
    }

    // 4. Page numbers (all 6 safe-margin presets, collision-free with header/footer & page text)
    if (overlayLayout.pageNumberItem) {
      const pnItem = overlayLayout.pageNumberItem;
      page.drawText(pnItem.text, {
        x: pnItem.x,
        y: pnItem.y,
        size: pnItem.fontSize,
        font,
        color: rgb(0.25, 0.31, 0.39),
      });
    }
  }

  return doc.save();
}
