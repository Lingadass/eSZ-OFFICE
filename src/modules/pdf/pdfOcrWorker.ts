import * as pdfjsLib from 'pdfjs-dist';
import {PDFDocument, StandardFonts, rgb, TextRenderingMode} from 'pdf-lib';
import {ensurePdfJsWorker} from './pdfEngine';

export interface OcrProgressUpdate {
  page: number;
  totalPages: number;
  status: string;
  percent: number;
}

/**
 * Lazy-loads Tesseract.js Web Worker to convert scanned PDF pages into a searchable PDF
 * with an invisible positioned text layer over high-resolution page images.
 */
export async function runOcrToSearchablePdf(
  pdfBytes: Uint8Array,
  pageIndices: number[],
  lang = 'eng',
  onProgress?: (info: OcrProgressUpdate) => void,
): Promise<{pdfBytes: Uint8Array; extractedText: string}> {
  ensurePdfJsWorker();
  const {createWorker} = await import('tesseract.js');

  const loadingTask = pdfjsLib.getDocument({data: new Uint8Array(pdfBytes)});
  const pdfJsDoc = await loadingTask.promise;
  const outDoc = await PDFDocument.create();
  outDoc.setCreator('eSZ OFFICE OCR');
  const font = await outDoc.embedFont(StandardFonts.Helvetica);

  const targets =
    pageIndices.length > 0
      ? pageIndices.filter((i) => i >= 0 && i < pdfJsDoc.numPages)
      : Array.from({length: pdfJsDoc.numPages}, (_, i) => i);

  let currentPageIdx = 0;
  const worker = await createWorker(lang, undefined, {
    logger: (m) => {
      if (typeof m.progress === 'number') {
        onProgress?.({
          page: currentPageIdx + 1,
          totalPages: targets.length,
          status: m.status || 'Recognizing text...',
          percent: Math.round(m.progress * 100),
        });
      }
    },
  });

  const allText: string[] = [];

  try {
    for (let i = 0; i < targets.length; i++) {
      currentPageIdx = i;
      const pageNum = targets[i] + 1;
      onProgress?.({
        page: i + 1,
        totalPages: targets.length,
        status: `Rendering page ${pageNum}...`,
        percent: 5,
      });

      const pdfPage = await pdfJsDoc.getPage(pageNum);
      const origViewport = pdfPage.getViewport({scale: 1});
      const ocrScale = 2;
      const viewport = pdfPage.getViewport({scale: ocrScale});

      const canvas = document.createElement('canvas');
      canvas.width = Math.max(1, Math.floor(viewport.width));
      canvas.height = Math.max(1, Math.floor(viewport.height));
      const ctx = canvas.getContext('2d');
      if (!ctx) continue;

      ctx.fillStyle = '#FFFFFF';
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      await pdfPage.render({canvasContext: ctx, viewport, canvas}).promise;

      const result = await worker.recognize(canvas, {}, {blocks: true, text: true});
      if (result.data.text) {
        allText.push(result.data.text);
      }

      const jpegBlob = await new Promise<Blob | null>((res) =>
        canvas.toBlob((b) => res(b), 'image/jpeg', 0.9),
      );
      if (!jpegBlob) continue;
      const jpgEmbedded = await outDoc.embedJpg(await jpegBlob.arrayBuffer());

      const page = outDoc.addPage([origViewport.width, origViewport.height]);
      page.drawImage(jpgEmbedded, {
        x: 0,
        y: 0,
        width: origViewport.width,
        height: origViewport.height,
      });

      // Embed invisible selectable text layer from Tesseract word bounding boxes
      const blocks = result.data.blocks ?? [];
      for (const block of blocks) {
        for (const para of block.paragraphs ?? []) {
          for (const line of para.lines ?? []) {
            for (const word of line.words ?? []) {
              const clean = (word.text || '').replace(/[^\x20-\x7E]/g, ' ').trim();
              if (!clean) continue;
              const {x0, y0, y1} = word.bbox;
              const pdfX = x0 / ocrScale;
              const boxH = Math.max(6, (y1 - y0) / ocrScale);
              const pdfY = origViewport.height - y1 / ocrScale;
              page.drawText(clean, {
                x: pdfX,
                y: pdfY,
                size: Math.min(32, Math.max(6, boxH * 0.85)),
                font,
                color: rgb(0, 0, 0),
                opacity: 0,
              });
              void TextRenderingMode;
            }
          }
        }
      }
    }

    return {
      pdfBytes: await outDoc.save(),
      extractedText: allText.join('\n\n'),
    };
  } finally {
    await worker.terminate();
    await loadingTask.destroy();
  }
}
