import {
  htmlToDocxBytes,
  htmlToMarkdown,
  htmlToPlainText,
  htmlToRtf,
  markdownToHtml,
  plainTextToHtml,
  rtfToHtml,
} from './docsPure';
import {resolveRenderFontStack} from './docFonts';
import {importDocxToHtml, type DocxImportWarning} from './docxImport';
import {DOC_THEMES, type DocNoteItem, type DocPageLayout} from './docsTypes';

export interface LoadedDocumentResult {
  html: string;
  warnings: string[];
  detailedWarnings: DocxImportWarning[];
  layoutPatch?: Partial<DocPageLayout>;
  hasLockedElements: boolean;
  hasLossyElements: boolean;
}

function isZipBinary(bytes: Uint8Array): boolean {
  return bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 0x03 && bytes[3] === 0x04;
}

/**
 * Loads a stored Blob (.docx, .html, .md, .rtf, .txt) into TipTap-compatible HTML.
 * Uses our custom OOXML parser (JSZip + XML DOM) for real .docx archives, preserving
 * fonts, sizes, alignment, tables, borders, images, lists, shapes, and headers/footers.
 */
export async function loadDocumentBlobAsHtml(
  blob: Blob,
  extension: string,
): Promise<LoadedDocumentResult> {
  const ext = extension.toLowerCase();
  const buffer = await blob.arrayBuffer();
  const bytes = new Uint8Array(buffer);

  if (ext === '.docx' && isZipBinary(bytes)) {
    const result = await importDocxToHtml(buffer);
    return {
      html: result.html,
      warnings: result.warnings.map((w) => `${w.feature}: ${w.message}`),
      detailedWarnings: result.warnings,
      layoutPatch: result.layoutPatch,
      hasLockedElements: result.hasLockedElements,
      hasLossyElements: result.hasLossyElements,
    };
  }

  const text = new TextDecoder('utf-8', {fatal: false}).decode(bytes);
  if (!text.trim()) {
    return {
      html: '<p></p>',
      warnings: [],
      detailedWarnings: [],
      hasLockedElements: false,
      hasLossyElements: false,
    };
  }

  if (ext === '.html' || ext === '.htm' || text.trim().startsWith('<')) {
    const bodyMatch = text.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
    return {
      html: (bodyMatch ? bodyMatch[1] : text).trim() || '<p></p>',
      warnings: [],
      detailedWarnings: [],
      hasLockedElements: false,
      hasLossyElements: false,
    };
  }

  if (ext === '.md') {
    return {
      html: markdownToHtml(text),
      warnings: [],
      detailedWarnings: [],
      hasLockedElements: false,
      hasLossyElements: false,
    };
  }

  if (ext === '.rtf' || text.trim().startsWith('{\\rtf')) {
    return {
      html: rtfToHtml(text),
      warnings: [],
      detailedWarnings: [],
      hasLockedElements: false,
      hasLossyElements: false,
    };
  }

  // Starter .docx seeded as plain text or .txt file
  return {
    html: plainTextToHtml(text),
    warnings: [],
    detailedWarnings: [],
    hasLockedElements: false,
    hasLossyElements: false,
  };
}

/**
 * Serializes editor HTML into a Blob matching the requested extension (.docx, .html, .md, .rtf, .txt).
 */
export async function serializeDocumentToBlob(
  html: string,
  extension: string,
  title: string,
  layout: DocPageLayout,
  notes: DocNoteItem[],
  originalDocxBuffer?: ArrayBuffer | Uint8Array | null,
): Promise<{blob: Blob; mimeType: string; ext: string}> {
  const ext = extension.toLowerCase();

  if (ext === '.docx') {
    const bytes = await htmlToDocxBytes(html, layout, notes, originalDocxBuffer);
    const mimeType =
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
    return {blob: new Blob([new Uint8Array(bytes)], {type: mimeType}), mimeType, ext: '.docx'};
  }

  if (ext === '.html' || ext === '.htm') {
    const htmlFontStack = resolveRenderFontStack(
      layout.docDefaultFont || DOC_THEMES[layout.theme].canonicalFontFamily,
      'Calibri',
    );
    const fullHtml = `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<title>${title.replace(/</g, '&lt;')}</title>
<style>body{font-family:${htmlFontStack};max-width:794px;margin:2rem auto;padding:0 1.5rem;line-height:1.5;color:${layout.docDefaultColor || '#0f172a'}}table{border-collapse:collapse;width:100%}td,th{border:1px solid #cbd5e1;padding:6px 10px}</style>
</head>
<body>
${html}
</body>
</html>`;
    return {blob: new Blob([fullHtml], {type: 'text/html'}), mimeType: 'text/html', ext: '.html'};
  }

  if (ext === '.md') {
    const md = htmlToMarkdown(html);
    return {blob: new Blob([md], {type: 'text/markdown'}), mimeType: 'text/markdown', ext: '.md'};
  }

  if (ext === '.rtf') {
    const rtf = htmlToRtf(html);
    return {blob: new Blob([rtf], {type: 'application/rtf'}), mimeType: 'application/rtf', ext: '.rtf'};
  }

  const txt = htmlToPlainText(html);
  return {blob: new Blob([txt], {type: 'text/plain'}), mimeType: 'text/plain', ext: '.txt'};
}
