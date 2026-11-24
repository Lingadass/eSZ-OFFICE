import {test} from 'vitest';
import assert from 'node:assert/strict';
import {PDFDocument} from 'pdf-lib';
import {
  bakeAnnotationsAndOverlays,
  createBlankPdfBytes,
  extractPdfPagesBytes,
  mergePdfBytes,
  parsePageRange,
  reorderAndTransformPdfBytes,
  splitIntoChunkSpecs,
  textToPdfBytes,
  updatePdfMetadata,
  readPdfMetadata,
} from './pdfPure';

test('parsePageRange handles single pages, ranges, reverse ranges, and z token', () => {
  assert.deepEqual(parsePageRange('1-3, 5, z', 7), [0, 1, 2, 4, 6]);
  assert.deepEqual(parsePageRange('', 3), [0, 1, 2]);
  assert.deepEqual(parsePageRange('99', 4), []);
});

test('splitIntoChunkSpecs partitions pages evenly', () => {
  assert.deepEqual(splitIntoChunkSpecs(5, 2), [[0, 1], [2, 3], [4]]);
});

test('createBlankPdfBytes, textToPdfBytes, merge, extract, reorder, and metadata work end-to-end', async () => {
  const pdf1 = await createBlankPdfBytes('Doc One');
  const pdf2 = await textToPdfBytes('Hello eSZ OFFICE PDF module.\nSecond line.', 'Doc Two');

  const merged = await mergePdfBytes([pdf1, pdf2]);
  const mergedDoc = await PDFDocument.load(merged);
  assert.equal(mergedDoc.getPageCount(), 2);

  const reordered = await reorderAndTransformPdfBytes(merged, [
    {sourceIndex: 1, extraRotationDeg: 90},
    {sourceIndex: 'blank-a4'},
    {sourceIndex: 0},
  ]);
  const reorderedDoc = await PDFDocument.load(reordered);
  assert.equal(reorderedDoc.getPageCount(), 3);
  assert.equal(reorderedDoc.getPage(0).getRotation().angle, 90);

  const extracted = await extractPdfPagesBytes(reordered, [0, 2]);
  const extractedDoc = await PDFDocument.load(extracted);
  assert.equal(extractedDoc.getPageCount(), 2);

  const withMeta = await updatePdfMetadata(extracted, {
    title: 'Updated Spec',
    author: 'Tester',
    subject: 'Unit Test',
    keywords: 'pdf, offline',
    creator: 'eSZ OFFICE',
  });
  const meta = await readPdfMetadata(withMeta);
  assert.equal(meta.title, 'Updated Spec');
  assert.equal(meta.author, 'Tester');

  const annotated = await bakeAnnotationsAndOverlays(
    withMeta,
    [
      {
        id: 'a1',
        pageIndex: 0,
        tool: 'highlight',
        x: 0.1,
        y: 0.1,
        width: 0.3,
        height: 0.05,
        color: '#FACC15',
        opacity: 0.35,
        strokeWidth: 2,
      },
      {
        id: 'a2',
        pageIndex: 0,
        tool: 'stamp',
        x: 0.2,
        y: 0.3,
        width: 0.25,
        height: 0.08,
        color: '#DC2626',
        opacity: 0.9,
        strokeWidth: 2,
        text: 'APPROVED',
      },
    ],
    {
      watermark: {
        enabled: true,
        text: 'CONFIDENTIAL',
        opacity: 0.2,
        fontSize: 40,
        rotationDeg: 45,
        color: '#64748B',
      },
      pageNumbers: {
        enabled: true,
        position: 'bottom-center',
        format: 'page-of-total',
        startNumber: 1,
      },
    },
  );
  const annotatedDoc = await PDFDocument.load(annotated);
  assert.equal(annotatedDoc.getPageCount(), 2);
});
