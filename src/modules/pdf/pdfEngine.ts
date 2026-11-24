import * as pdfjsLib from 'pdfjs-dist';
import {
  PDFCheckBox,
  PDFDocument,
  PDFDropdown,
  PDFRadioGroup,
  PDFTextField,
} from 'pdf-lib';
import {createBlankPdfBytes} from './pdfPure';
import type {PdfAnnotationItem, PdfFormFieldInfo} from './pdfTypes';

let workerConfigured = false;
export function ensurePdfJsWorker(): void {
  if (!workerConfigured && typeof window !== 'undefined') {
    pdfjsLib.GlobalWorkerOptions.workerSrc = '/workers/pdf.worker.min.mjs';
    workerConfigured = true;
  }
}

export async function ensureValidPdfBytes(rawBlob: Blob, title: string): Promise<Uint8Array> {
  const buf = new Uint8Array(await rawBlob.arrayBuffer());
  // Check for %PDF- header within first 1024 bytes
  const headerSlice = new TextDecoder('latin1').decode(buf.slice(0, 1024));
  if (headerSlice.includes('%PDF-')) {
    return buf;
  }
  // Convert Phase 0 text starter PDF cleanly into a real multi-line PDF binary
  const text = new TextDecoder('utf-8', {fatal: false}).decode(buf);
  return createBlankPdfBytes(title, text);
}

export async function getPdfFormFields(pdfBytes: Uint8Array): Promise<PdfFormFieldInfo[]> {
  const doc = await PDFDocument.load(pdfBytes, {ignoreEncryption: true});
  try {
    const form = doc.getForm();
    return form.getFields().map((field) => {
      const name = field.getName();
      if (field instanceof PDFTextField) {
        return {name, type: 'text', value: field.getText() ?? ''};
      }
      if (field instanceof PDFCheckBox) {
        return {name, type: 'checkbox', value: field.isChecked()};
      }
      if (field instanceof PDFDropdown) {
        return {
          name,
          type: 'dropdown',
          value: field.getSelected()[0] ?? '',
          options: field.getOptions(),
        };
      }
      if (field instanceof PDFRadioGroup) {
        return {
          name,
          type: 'radio',
          value: field.getSelected() ?? '',
          options: field.getOptions(),
        };
      }
      return {name, type: 'unknown', value: ''};
    });
  } catch {
    return [];
  }
}

export async function applyPdfFormFields(
  pdfBytes: Uint8Array,
  fields: PdfFormFieldInfo[],
  flattenAfter = false,
): Promise<Uint8Array> {
  const doc = await PDFDocument.load(pdfBytes, {ignoreEncryption: true});
  const form = doc.getForm();

  for (const item of fields) {
    try {
      const field = form.getField(item.name);
      if (field instanceof PDFTextField) {
        field.setText(String(item.value ?? ''));
      } else if (field instanceof PDFCheckBox) {
        if (item.value) field.check();
        else field.uncheck();
      } else if (field instanceof PDFDropdown) {
        if (String(item.value)) field.select(String(item.value));
      } else if (field instanceof PDFRadioGroup) {
        if (String(item.value)) field.select(String(item.value));
      }
    } catch {
      // Skip non-editable field
    }
  }

  if (flattenAfter) {
    form.flatten();
  }
  return doc.save();
}

async function getQpdfInstance() {
  const {createQpdf} = await import('@mssio/qpdf-wasm');
  return createQpdf({
    wasmUrl: '/wasm/qpdf.wasm',
    workers: 1,
  });
}

export async function encryptPdfWithQpdf(
  pdfBytes: Uint8Array,
  options: {
    userPassword: string;
    ownerPassword: string;
    allowPrint: boolean;
    allowModify: boolean;
    allowExtract: boolean;
    allowAnnotate: boolean;
  },
): Promise<Uint8Array> {
  const qpdf = await getQpdfInstance();
  try {
    const res = await qpdf.encrypt(pdfBytes.slice(), {
      userPassword: options.userPassword,
      ownerPassword: options.ownerPassword || options.userPassword,
      bits: 256,
      allow: {
        print: options.allowPrint,
        modify: options.allowModify,
        extract: options.allowExtract,
        annotate: options.allowAnnotate,
      },
    });
    return new Uint8Array(res.output);
  } finally {
    qpdf.terminate();
  }
}

export async function decryptPdfWithQpdf(
  pdfBytes: Uint8Array,
  password: string,
): Promise<Uint8Array> {
  const qpdf = await getQpdfInstance();
  try {
    const res = await qpdf.decrypt(pdfBytes.slice(), {password});
    return new Uint8Array(res.output);
  } finally {
    qpdf.terminate();
  }
}

export async function compressPdfWithQpdf(pdfBytes: Uint8Array): Promise<Uint8Array> {
  const qpdf = await getQpdfInstance();
  try {
    const res = await qpdf.compress(pdfBytes.slice(), {level: 9});
    return new Uint8Array(res.output);
  } finally {
    qpdf.terminate();
  }
}

/**
 * True pixel-level redaction: for pages containing redact boxes, renders the page via pdf.js,
 * paints solid opaque black over redacted regions, and replaces those pages in pdf-lib so
 * underlying selectable text is permanently destroyed.
 */
export async function applyTrueRedactions(
  pdfBytes: Uint8Array,
  redactions: PdfAnnotationItem[],
  password?: string,
): Promise<Uint8Array> {
  if (redactions.length === 0) return pdfBytes;
  ensurePdfJsWorker();

  const loadingTask = pdfjsLib.getDocument({
    data: new Uint8Array(pdfBytes),
    password,
  });
  const pdfJsDoc = await loadingTask.promise;
  const outDoc = await PDFDocument.load(pdfBytes, {ignoreEncryption: true});
  const pagesToRedact = new Set(redactions.map((r) => r.pageIndex));

  try {
    for (const pageIndex of Array.from(pagesToRedact).sort((a, b) => a - b)) {
      const pdfPage = await pdfJsDoc.getPage(pageIndex + 1);
      const viewport = pdfPage.getViewport({scale: 2});
      const canvas = document.createElement('canvas');
      canvas.width = Math.max(1, Math.floor(viewport.width));
      canvas.height = Math.max(1, Math.floor(viewport.height));
      const ctx = canvas.getContext('2d');
      if (!ctx) continue;

      await pdfPage.render({canvasContext: ctx, viewport, canvas}).promise;

      // Paint solid black over redacted regions
      ctx.fillStyle = '#000000';
      for (const box of redactions.filter((r) => r.pageIndex === pageIndex)) {
        ctx.fillRect(
          box.x * canvas.width,
          box.y * canvas.height,
          Math.max(4, box.width * canvas.width),
          Math.max(4, box.height * canvas.height),
        );
      }

      const pngBlob = await new Promise<Blob | null>((res) =>
        canvas.toBlob((b) => res(b), 'image/png'),
      );
      if (!pngBlob) continue;
      const pngBytes = new Uint8Array(await pngBlob.arrayBuffer());
      const embedded = await outDoc.embedPng(pngBytes);
      const origPage = outDoc.getPage(pageIndex);
      const {width, height} = origPage.getSize();

      // Replace the original page with a clean flattened image-only page
      outDoc.removePage(pageIndex);
      const newPage = outDoc.insertPage(pageIndex, [width, height]);
      newPage.drawImage(embedded, {x: 0, y: 0, width, height});
    }
    return await outDoc.save();
  } finally {
    await loadingTask.destroy();
  }
}
