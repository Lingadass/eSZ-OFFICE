import JSZip from 'jszip';
import * as pdfjsLib from 'pdfjs-dist';
import {PDFDocument} from 'pdf-lib';
import {ensurePdfJsWorker} from './pdfEngine';
import {textToPdfBytes} from './pdfPure';

export async function imagesToPdfBytes(imageFiles: File[]): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  doc.setCreator('eSZ OFFICE');

  for (const file of imageFiles) {
    const buffer = await file.arrayBuffer();
    const isPng = file.type === 'image/png' || file.name.toLowerCase().endsWith('.png');
    const isJpg =
      file.type === 'image/jpeg' ||
      file.name.toLowerCase().endsWith('.jpg') ||
      file.name.toLowerCase().endsWith('.jpeg');

    let embedded;
    if (isPng) {
      embedded = await doc.embedPng(buffer);
    } else if (isJpg) {
      embedded = await doc.embedJpg(buffer);
    } else {
      // Convert WebP/BMP/GIF to PNG via canvas
      const bitmap = await createImageBitmap(file);
      const canvas = document.createElement('canvas');
      canvas.width = bitmap.width;
      canvas.height = bitmap.height;
      const ctx = canvas.getContext('2d');
      ctx?.drawImage(bitmap, 0, 0);
      bitmap.close();
      const pngBlob = await new Promise<Blob>((resolve, reject) =>
        canvas.toBlob((b) => (b ? resolve(b) : reject(new Error('Image conversion failed'))), 'image/png'),
      );
      embedded = await doc.embedPng(await pngBlob.arrayBuffer());
    }

    const page = doc.addPage([embedded.width, embedded.height]);
    page.drawImage(embedded, {
      x: 0,
      y: 0,
      width: embedded.width,
      height: embedded.height,
    });
  }

  return doc.save();
}

export async function pdfToImagesZip(
  pdfBytes: Uint8Array,
  format: 'png' | 'jpeg' = 'png',
  scale = 2,
  onProgress?: (page: number, total: number) => void,
): Promise<{zipBlob: Blob; pageCount: number}> {
  ensurePdfJsWorker();
  const loadingTask = pdfjsLib.getDocument({data: new Uint8Array(pdfBytes)});
  const doc = await loadingTask.promise;
  const zip = new JSZip();

  try {
    const total = doc.numPages;
    for (let p = 1; p <= total; p++) {
      onProgress?.(p, total);
      const page = await doc.getPage(p);
      const viewport = page.getViewport({scale});
      const canvas = document.createElement('canvas');
      canvas.width = Math.max(1, Math.floor(viewport.width));
      canvas.height = Math.max(1, Math.floor(viewport.height));
      const ctx = canvas.getContext('2d');
      if (!ctx) continue;

      if (format === 'jpeg') {
        ctx.fillStyle = '#FFFFFF';
        ctx.fillRect(0, 0, canvas.width, canvas.height);
      }

      await page.render({canvasContext: ctx, viewport, canvas}).promise;
      const mime = format === 'jpeg' ? 'image/jpeg' : 'image/png';
      const blob = await new Promise<Blob | null>((res) =>
        canvas.toBlob((b) => res(b), mime, 0.92),
      );
      if (blob) {
        const ext = format === 'jpeg' ? 'jpg' : 'png';
        const pad = String(p).padStart(3, '0');
        zip.file(`page-${pad}.${ext}`, await blob.arrayBuffer());
      }
    }

    const zipBlob = await zip.generateAsync({type: 'blob'});
    return {zipBlob, pageCount: total};
  } finally {
    await loadingTask.destroy();
  }
}

/**
 * Best-effort client-side conversion for .txt, .md, .docx, and .pptx to PDF.
 */
export async function officeFileToPdfBestEffort(file: File): Promise<Uint8Array> {
  const lower = file.name.toLowerCase();

  if (lower.endsWith('.txt') || lower.endsWith('.md') || lower.endsWith('.csv')) {
    const text = await file.text();
    return textToPdfBytes(text, file.name);
  }

  if (lower.endsWith('.docx')) {
    const zip = await JSZip.loadAsync(file);
    const xmlStr = await zip.file('word/document.xml')?.async('string');
    if (!xmlStr) {
      throw new Error('Could not read word/document.xml from .docx archive.');
    }
    const xmlDoc = new DOMParser().parseFromString(xmlStr, 'application/xml');
    const paragraphs = Array.from(xmlDoc.getElementsByTagName('w:p'));
    const lines = paragraphs
      .map((p) =>
        Array.from(p.getElementsByTagName('w:t'))
          .map((t) => t.textContent ?? '')
          .join(''),
      )
      .filter((line) => line.trim().length > 0);
    return textToPdfBytes(
      lines.join('\n\n') || '(Empty DOCX document)',
      file.name.replace(/\.docx$/i, ''),
    );
  }

  if (lower.endsWith('.pptx')) {
    const zip = await JSZip.loadAsync(file);
    const slidePaths = Object.keys(zip.files)
      .filter((p) => /^ppt\/slides\/slide\d+\.xml$/i.test(p))
      .sort((a, b) => {
        const na = Number.parseInt(a.match(/slide(\d+)\.xml/i)?.[1] ?? '0', 10);
        const nb = Number.parseInt(b.match(/slide(\d+)\.xml/i)?.[1] ?? '0', 10);
        return na - nb;
      });
    const sections: string[] = [];
    for (let i = 0; i < slidePaths.length; i++) {
      const xmlStr = await zip.file(slidePaths[i])?.async('string');
      if (!xmlStr) continue;
      const xmlDoc = new DOMParser().parseFromString(xmlStr, 'application/xml');
      const texts = Array.from(xmlDoc.getElementsByTagName('a:t'))
        .map((t) => t.textContent ?? '')
        .filter((s) => s.trim().length > 0);
      sections.push(`--- Slide ${i + 1} ---\n${texts.join('\n')}`);
    }
    return textToPdfBytes(
      sections.join('\n\n') || '(Empty PPTX presentation)',
      file.name.replace(/\.pptx$/i, ''),
    );
  }

  throw new Error('Supported conversion inputs: PNG, JPG, WebP, TXT, MD, CSV, DOCX, PPTX.');
}
