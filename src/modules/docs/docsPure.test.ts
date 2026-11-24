import assert from 'node:assert/strict';
import {test} from 'vitest';
import {
  buildTocHtml,
  computeDocumentStats,
  extractHeadingsFromHtml,
  htmlToDocxBytes,
  htmlToMarkdown,
  htmlToPdfBytes,
  htmlToRtf,
  markdownToHtml,
  rtfToHtml,
} from './docsPure';
import {DEFAULT_DOC_LAYOUT} from './docsTypes';

test('computeDocumentStats counts words, characters, paragraphs, and reading time', () => {
  const html = '<h1>Quarterly Review</h1><p>Offline-first documents stay private on device.</p><p>Zero cloud uploads.</p>';
  const stats = computeDocumentStats(html);
  assert.equal(stats.words, 11);
  assert.equal(stats.paragraphs, 3);
  assert.equal(stats.estimatedPages, 1);
  assert.equal(stats.readingMinutes, 1);
});

test('extractHeadingsFromHtml and buildTocHtml extract H1-H3 hierarchy cleanly', () => {
  const html = '<h1>Overview</h1><p>Intro</p><h2>Architecture</h2><h3>Storage</h3>';
  const headings = extractHeadingsFromHtml(html);
  assert.equal(headings.length, 3);
  assert.deepEqual(
    headings.map((h) => [h.level, h.text]),
    [
      [1, 'Overview'],
      [2, 'Architecture'],
      [3, 'Storage'],
    ],
  );
  const toc = buildTocHtml(headings);
  assert.ok(toc.includes('data-doc-toc="true"'));
  assert.ok(toc.includes('Architecture'));
});

test('Markdown and RTF conversions round-trip structured headings and paragraphs', () => {
  const md = '# Project Plan\n\nThis is **bold** and *italic*.\n\n- Alpha\n- Beta';
  const html = markdownToHtml(md);
  assert.ok(html.includes('<h1>Project Plan</h1>'));
  assert.ok(html.includes('<strong>bold</strong>'));
  assert.ok(html.includes('<ul>'));

  const backMd = htmlToMarkdown(html);
  assert.ok(backMd.includes('# Project Plan'));
  assert.ok(backMd.includes('**bold**'));

  const rtf = htmlToRtf(html);
  assert.ok(rtf.startsWith('{\\rtf1'));
  const fromRtf = rtfToHtml(rtf);
  assert.ok(fromRtf.includes('Project Plan'));
});

test('htmlToDocxBytes and htmlToPdfBytes emit valid DOCX zip and PDF binaries', async () => {
  const sampleHtml =
    '<h1>eSZ OFFICE Spec</h1><p>Built for <strong>offline</strong> reliability.</p><table><tr><th>Module</th><th>Status</th></tr><tr><td>Docs</td><td>Ready</td></tr></table>';

  const docxBytes = await htmlToDocxBytes(sampleHtml, {
    ...DEFAULT_DOC_LAYOUT,
    headerText: 'eSZ OFFICE Draft',
    footerText: 'Confidential',
  });
  // Valid OOXML ZIP magic header PK\x03\x04
  assert.equal(docxBytes[0], 0x50);
  assert.equal(docxBytes[1], 0x4b);
  assert.ok(docxBytes.byteLength > 1000);

  const pdfBytes = await htmlToPdfBytes(sampleHtml, 'eSZ OFFICE Spec', {
    ...DEFAULT_DOC_LAYOUT,
    watermarkEnabled: true,
    watermarkText: 'DRAFT',
  });
  const pdfHeader = new TextDecoder('latin1').decode(pdfBytes.slice(0, 8));
  assert.ok(pdfHeader.startsWith('%PDF-'));
  assert.ok(pdfBytes.byteLength > 800);
});
