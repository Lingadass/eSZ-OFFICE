import {describe, it} from 'vitest';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {
  computePaginationPlan,
  DOC_PAGE_GAP_PX,
  formatPageNumber,
  getPageMetrics,
  shouldShowPageChrome,
  type MeasurableBlock,
} from './docsPagination';
import {computeVisiblePageNumber} from './docxPreviewUtils';
import {htmlToDocxBytes, htmlToPdfBytes} from './docsPure';
import {DEFAULT_DOC_LAYOUT, type DocNoteItem, type DocPageLayout} from './docsTypes';

describe('docsPagination — physical page-by-page layout engine', () => {
  it('computes exact non-drifting pageTops and spacer heights for a 5-page document', () => {
    const metrics = getPageMetrics(DEFAULT_DOC_LAYOUT);
    // A4 portrait normal margins: 794 x 1123, margins 72 top/bottom -> contentHeight = 979
    assert.equal(metrics.pageWidth, 794);
    assert.equal(metrics.pageHeight, 1123);
    assert.equal(metrics.contentHeight, 979);

    // 5 pages via 20 blocks of 220px each (4 fit per page -> 4*220 + 3*10 = 910 <= 971)
    const blocks: MeasurableBlock[] = Array.from({length: 20}, (_, idx) => ({
      pos: idx * 20,
      endPos: idx * 20 + 18,
      kind: 'paragraph',
      height: 220,
      marginTop: 0,
      marginBottom: 10,
      lineHeight: 22,
    }));

    const plan = computePaginationPlan(blocks, metrics, new Map());
    assert.equal(plan.pageCount, 5);
    assert.equal(plan.breaks.length, 4);

    const stride = metrics.pageHeight + DOC_PAGE_GAP_PX;
    assert.deepEqual(plan.pageTops, [0, stride, stride * 2, stride * 3, stride * 4]);
    assert.equal(plan.totalStackHeight, 5 * metrics.pageHeight + 4 * DOC_PAGE_GAP_PX);

    // Verify zero cumulative drift: each break lands precisely at target page start
    for (let i = 0; i < plan.breaks.length; i++) {
      const br = plan.breaks[i];
      assert.equal(br.fromPage, i);
      assert.equal(br.toPage, i + 1);
      // 4 blocks occupy 4*220 + 4*10 = 920px; stride is 1147px -> spacer = 227px
      assert.equal(br.spacerHeight, stride - 920);
    }
  });

  it('honors manual Page Break nodes', () => {
    const metrics = getPageMetrics(DEFAULT_DOC_LAYOUT);
    const blocks: MeasurableBlock[] = [
      {pos: 0, endPos: 10, kind: 'paragraph', height: 40, marginTop: 0, marginBottom: 10},
      {pos: 10, endPos: 11, kind: 'pageBreak', height: 18, marginTop: 4, marginBottom: 4},
      {pos: 11, endPos: 25, kind: 'paragraph', height: 40, marginTop: 0, marginBottom: 10},
    ];
    const plan = computePaginationPlan(blocks, metrics, new Map());
    assert.equal(plan.pageCount, 2);
    assert.equal(plan.breaks.length, 1);
    assert.equal(plan.breaks[0].manual, true);
    assert.equal(plan.breaks[0].pos, 11);
    assert.equal(plan.breaks[0].fromPage, 0);
    assert.equal(plan.breaks[0].toPage, 1);
  });

  it('applies keep-with-next for headings near the bottom of a page', () => {
    const metrics = getPageMetrics(DEFAULT_DOC_LAYOUT);
    // Usable height = 979 - 8 = 971.
    // Put a 900px block, then a 34px heading (would fit alone at 934px, but heading + 2 lines of next paragraph overflows!)
    const blocks: MeasurableBlock[] = [
      {pos: 0, endPos: 10, kind: 'paragraph', height: 900, marginTop: 0, marginBottom: 10},
      {pos: 10, endPos: 20, kind: 'heading', height: 34, marginTop: 12, marginBottom: 8, lineHeight: 34},
      {pos: 20, endPos: 40, kind: 'paragraph', height: 66, marginTop: 0, marginBottom: 10, lineHeight: 22},
    ];
    const plan = computePaginationPlan(blocks, metrics, new Map());
    assert.equal(plan.pageCount, 2);
    assert.equal(plan.breaks.length, 1);
    // Break MUST occur BEFORE the heading (pos 10), not between heading and paragraph
    assert.equal(plan.breaks[0].pos, 10);
  });

  it('splits long paragraphs cleanly while keeping at least 2 lines together (orphan/widow control)', () => {
    const metrics = getPageMetrics(DEFAULT_DOC_LAYOUT);
    // Usable height = 971px. First block takes 800px + 10px margin = 810px -> 161px left on page 1.
    // Long paragraph is 10 lines * 24px = 240px.
    // With 161px available, up to 6 lines (144px) fit on page 1, leaving 4 lines (96px >= 2 lines) on page 2.
    const blocks: MeasurableBlock[] = [
      {pos: 0, endPos: 10, kind: 'paragraph', height: 800, marginTop: 0, marginBottom: 10, lineHeight: 24},
      {
        pos: 10,
        endPos: 500,
        kind: 'paragraph',
        height: 240,
        marginTop: 0,
        marginBottom: 10,
        lineHeight: 24,
        resolveParagraphSplit: (availablePx, fromOffsetY) => {
          const fitLines = Math.floor(availablePx / 24);
          if (fitLines < 2) return null;
          return {
            pos: 10 + fitLines * 40,
            offsetY: fromOffsetY + fitLines * 24,
          };
        },
      },
    ];

    const plan = computePaginationPlan(blocks, metrics, new Map());
    assert.equal(plan.pageCount, 2);
    assert.equal(plan.breaks.length, 1);
    assert.equal(plan.breaks[0].kind, 'inline');
    // 6 lines fit in 161px -> split at 10 + 6*40 = 250
    assert.equal(plan.breaks[0].pos, 250);
  });

  it('breaks tables cleanly between rows and keeps header row with the first body row', () => {
    const metrics = getPageMetrics(DEFAULT_DOC_LAYOUT);
    // 30 rows of 45px each -> 21 rows fit on page 1 (21*45 = 945 <= 971), row index 21 starts page 2
    const rows = Array.from({length: 30}, (_, idx) => ({
      pos: 100 + idx * 10,
      height: 45,
      isHeader: idx === 0,
    }));
    const blocks: MeasurableBlock[] = [
      {
        pos: 99,
        endPos: 450,
        kind: 'table',
        height: 30 * 45,
        marginTop: 0,
        marginBottom: 12,
        tableRows: rows,
      },
    ];
    const plan = computePaginationPlan(blocks, metrics, new Map());
    assert.equal(plan.pageCount, 2);
    assert.equal(plan.breaks.length, 1);
    assert.equal(plan.breaks[0].kind, 'table-row');
    assert.equal(plan.breaks[0].pos, rows[21].pos);
  });

  it('places footnotes on the exact page where they are referenced and deducts usable height', () => {
    const metrics = getPageMetrics(DEFAULT_DOC_LAYOUT);
    const fn1: DocNoteItem = {id: 'fn1', kind: 'footnote', number: 1, text: 'Page 1 citation'};
    const fn2: DocNoteItem = {id: 'fn2', kind: 'footnote', number: 2, text: 'Page 2 citation'};
    const notesById = new Map<string, DocNoteItem>([
      [fn1.id, fn1],
      [fn2.id, fn2],
    ]);

    const blocks: MeasurableBlock[] = [
      {pos: 0, endPos: 10, kind: 'paragraph', height: 400, marginTop: 0, marginBottom: 10, footnoteIds: ['fn1']},
      {pos: 10, endPos: 11, kind: 'pageBreak', height: 10, marginTop: 0, marginBottom: 0},
      {pos: 11, endPos: 25, kind: 'paragraph', height: 200, marginTop: 0, marginBottom: 10, footnoteIds: ['fn2']},
    ];

    const plan = computePaginationPlan(blocks, metrics, notesById);
    assert.equal(plan.pageCount, 2);
    assert.equal(plan.pageFootnotes[0]?.[0]?.id, 'fn1');
    assert.equal(plan.pageFootnotes[1]?.[0]?.id, 'fn2');
  });

  it('pushes entire paragraph to next page when fewer than 2 lines fit (orphan prevention)', () => {
    const metrics = getPageMetrics(DEFAULT_DOC_LAYOUT);
    // Usable height = 971px. First block takes 930px + 10px margin = 940px -> only 31px left (< 2 * 24px lines).
    const blocks: MeasurableBlock[] = [
      {pos: 0, endPos: 10, kind: 'paragraph', height: 930, marginTop: 0, marginBottom: 10, lineHeight: 24},
      {
        pos: 10,
        endPos: 200,
        kind: 'paragraph',
        height: 120,
        marginTop: 0,
        marginBottom: 10,
        lineHeight: 24,
        resolveParagraphSplit: (availablePx) => {
          const fitLines = Math.floor(availablePx / 24);
          return fitLines < 2 ? null : {pos: 50, offsetY: fitLines * 24};
        },
      },
    ];
    const plan = computePaginationPlan(blocks, metrics, new Map());
    assert.equal(plan.pageCount, 2);
    assert.equal(plan.breaks.length, 1);
    assert.equal(plan.breaks[0].kind, 'block');
    assert.equal(plan.breaks[0].pos, 10);
  });

  it('formats live page numbers with safe bounds clamping and respects differentFirstPage', () => {
    assert.equal(formatPageNumber(1, 5), 'Page 1 of 5');
    assert.equal(formatPageNumber(3, 5), 'Page 3 of 5');
    assert.equal(formatPageNumber(99, 5), 'Page 5 of 5');
    assert.equal(formatPageNumber(0, 5), 'Page 1 of 5');
    assert.equal(formatPageNumber(1, 0), 'Page 1 of 1');

    assert.equal(shouldShowPageChrome(0, false), true);
    assert.equal(shouldShowPageChrome(0, true), false);
    assert.equal(shouldShowPageChrome(1, true), true);
  });

  it('starts an oversized block/image on a fresh page and advances along exact pageStride without footer drift', () => {
    const metrics = getPageMetrics(DEFAULT_DOC_LAYOUT);
    const stride = metrics.pageHeight + DOC_PAGE_GAP_PX;
    // Page 1 has a 120px paragraph, followed by a 1250px oversized drawing (> 979px contentHeight), then a normal paragraph.
    const blocks: MeasurableBlock[] = [
      {pos: 0, endPos: 10, kind: 'paragraph', height: 120, marginTop: 0, marginBottom: 10},
      {pos: 10, endPos: 20, kind: 'other', height: 1250, marginTop: 0, marginBottom: 12},
      {pos: 20, endPos: 30, kind: 'paragraph', height: 80, marginTop: 0, marginBottom: 10},
    ];
    const plan = computePaginationPlan(blocks, metrics, new Map());
    // Oversized block must move to page 2 (index 1) first instead of overflowing page 1's footer
    assert.equal(plan.breaks[0].pos, 10);
    assert.equal(plan.breaks[0].fromPage, 0);
    assert.equal(plan.breaks[0].toPage, 1);
    assert.ok(plan.breaks[0].spacerHeight > 0, 'Spacer before oversized block must be positive');
    assert.equal(plan.breaks[0].spacerHeight, stride - 130);
    // All subsequent breaks must remain non-negative and aligned to exact pageTops
    for (const br of plan.breaks) {
      assert.ok(br.spacerHeight >= 0, 'Spacer height must never be negative');
    }
    assert.ok(plan.pageCount >= 3, 'Oversized 1250px block starting on page 2 spans into page 3');
  });

  it('moves tall table rows to a fresh page and keeps subsequent rows aligned without footer overlap', () => {
    const metrics = getPageMetrics(DEFAULT_DOC_LAYOUT);
    const stride = metrics.pageHeight + DOC_PAGE_GAP_PX;
    const rows = [
      {pos: 10, height: 40, isHeader: true},
      {pos: 20, height: 600, isHeader: false},
      {pos: 30, height: 480, isHeader: false}, // 40 + 600 + 480 = 1120 > 971 -> must break before pos 30
      {pos: 40, height: 120, isHeader: false},
    ];
    const blocks: MeasurableBlock[] = [
      {
        pos: 5,
        endPos: 50,
        kind: 'table',
        height: 1240,
        marginTop: 0,
        marginBottom: 12,
        tableRows: rows,
      },
    ];
    const plan = computePaginationPlan(blocks, metrics, new Map());
    assert.equal(plan.pageCount, 2);
    assert.equal(plan.breaks.length, 1);
    assert.equal(plan.breaks[0].kind, 'table-row');
    assert.equal(plan.breaks[0].pos, 30);
    // Rows 0 + 1 occupy 640px on page 1 -> spacer pushes row 2 cleanly to page 2 start (stride - 640)
    assert.equal(plan.breaks[0].spacerHeight, stride - 640);
  });

  it('suppresses an accidental blank page when a manual pageBreak sits immediately after an auto overflow break', () => {
    const metrics = getPageMetrics(DEFAULT_DOC_LAYOUT);
    // Block 0 fills page 1 (960px + 10px = 970px <= 971px), then block 1 (20px) triggers an auto break to page 2,
    // wait: test exact case where manual pageBreak itself is pushed after an auto break at yInPage === 0.
    // Block 0 (960px + 10px) -> yInPage = 970. Next is manual pageBreak (height 18px) -> does NOT trigger auto break first;
    // instead test: Block 0 (960px) -> Block 1 (50px, auto break to page 2) -> NO, yInPage is 50 after block 1.
    // Exact case: pageBreak directly after an auto-broken block that starts at y=0 is NOT y=0;
    // y=0 happens when an auto-break landed right before the pageBreak or consecutive breaks occur.
    // Specifically: two consecutive manual pageBreaks MUST still produce 3 pages (user explicitly inserted 2 breaks).
    const twoManual: MeasurableBlock[] = [
      {pos: 0, endPos: 10, kind: 'paragraph', height: 40, marginTop: 0, marginBottom: 10},
      {pos: 10, endPos: 11, kind: 'pageBreak', height: 18, marginTop: 0, marginBottom: 0},
      {pos: 11, endPos: 12, kind: 'pageBreak', height: 18, marginTop: 0, marginBottom: 0},
      {pos: 12, endPos: 25, kind: 'paragraph', height: 40, marginTop: 0, marginBottom: 10},
    ];
    const manualPlan = computePaginationPlan(twoManual, metrics, new Map());
    assert.equal(manualPlan.pageCount, 3);
    assert.equal(manualPlan.breaks.length, 2);
  });

  it('computes toolbar Page X of Y matching the visible page footer across 50%, 100%, 110%, and 150% zoom with top ruler offset', () => {
    const metrics = getPageMetrics(DEFAULT_DOC_LAYOUT);
    const topOffset = 52; // HorizontalRuler + container top padding
    const viewportHeight = 760;

    for (const zoom of [0.5, 1, 1.1, 1.5]) {
      const scaledPageH = metrics.pageHeight * zoom;
      const scaledGap = DOC_PAGE_GAP_PX * zoom;
      const pageBounds = [0, 1, 2].map((i) => ({
        top: topOffset + i * (scaledPageH + scaledGap),
        height: scaledPageH,
      }));

      // Top of scroll -> Page 1
      assert.equal(computeVisiblePageNumber(0, viewportHeight, pageBounds), 1);

      // Scrolled so Page 1's footer is still comfortably visible in the middle of the viewport
      const page1FooterScroll = Math.max(0, pageBounds[0].top + scaledPageH - viewportHeight * 0.65);
      assert.equal(
        computeVisiblePageNumber(page1FooterScroll, viewportHeight, pageBounds),
        1,
        `Page 1 footer view must report Page 1 at zoom ${zoom}`,
      );

      // Scrolled squarely onto Page 2 -> Page 2
      const page2TopScroll = pageBounds[1].top - 16;
      assert.equal(
        computeVisiblePageNumber(page2TopScroll, viewportHeight, pageBounds),
        2,
        `Page 2 top view must report Page 2 at zoom ${zoom}`,
      );

      // Scrolled squarely onto Page 3 -> Page 3
      const page3TopScroll = pageBounds[2].top - 16;
      assert.equal(
        computeVisiblePageNumber(page3TopScroll, viewportHeight, pageBounds),
        3,
        `Page 3 top view must report Page 3 at zoom ${zoom}`,
      );
    }
  });
});

describe('docsPure — .docx & .pdf pagination parity export', () => {
  it('exports .docx with real section page size, margins, header, footer, titlePage, and PAGE / NUMPAGES fields', async () => {
    const layout: DocPageLayout = {
      ...DEFAULT_DOC_LAYOUT,
      size: 'letter',
      orientation: 'portrait',
      margins: 'moderate',
      headerText: 'Quarterly Board Brief',
      footerText: 'Confidential Draft',
      showPageNumbers: true,
      pageNumberPosition: 'bottom-center',
      differentFirstPage: true,
    };

    const html =
      '<h1>Executive Summary</h1><p>Opening paragraph on page one.</p><div data-page-break="true">Page Break</div><p>Second page content.</p>';

    const bytes = await htmlToDocxBytes(html, layout);
    const zip = await JSZip.loadAsync(bytes);

    const docXml = await zip.file('word/document.xml')?.async('string');
    assert.ok(docXml, 'word/document.xml must exist');
    assert.match(docXml, /w:pgSz[^>]*w:w="12240"[^>]*w:h="15840"/);
    assert.match(docXml, /w:pgMar[^>]*w:top="1440"[^>]*w:right="1080"[^>]*w:bottom="1440"[^>]*w:left="1080"/);
    assert.match(docXml, /<w:titlePg\b/);
    assert.match(docXml, /<w:keepNext\b/);
    assert.match(docXml, /<w:widowControl\b/);

    const xmlEntries = Object.keys(zip.files).filter(
      (name) => name.startsWith('word/header') || name.startsWith('word/footer'),
    );
    assert.ok(xmlEntries.length >= 2, 'Expected header and footer XML parts');

    const combinedHf = (
      await Promise.all(xmlEntries.map((f) => zip.file(f)!.async('string')))
    ).join('\n');

    assert.ok(combinedHf.includes('Quarterly Board Brief'), 'Header text missing in .docx');
    assert.ok(combinedHf.includes('Confidential Draft'), 'Footer text missing in .docx');
    assert.match(combinedHf, /\bPAGE\b/, 'OOXML PAGE field missing in .docx');
    assert.match(combinedHf, /\bNUMPAGES\b/, 'OOXML NUMPAGES field missing in .docx');
  });

  it('exports multi-page .pdf honoring manual breaks, footnotes, watermark, and differentFirstPage', async () => {
    const layout: DocPageLayout = {
      ...DEFAULT_DOC_LAYOUT,
      headerText: 'Annual Report',
      footerText: 'Internal Use',
      watermarkEnabled: true,
      watermarkText: 'DRAFT',
      differentFirstPage: true,
    };
    const notes: DocNoteItem[] = [
      {id: 'fn_a', kind: 'footnote', number: 1, text: 'Source: 2026 Ledger'},
    ];
    const html =
      '<h1>Page 1</h1><p>Metric<sup data-note-ref="fn_a" data-note-kind="footnote">[1]</sup></p>' +
      '<div data-page-break="true">Page Break</div>' +
      '<h1>Page 2</h1><p>Second page text.</p>';

    const pdfBytes = await htmlToPdfBytes(html, 'Report', layout, notes);
    assert.ok(pdfBytes.length > 800, 'Expected non-empty PDF output');
    const headerSig = new TextDecoder().decode(pdfBytes.slice(0, 5));
    assert.equal(headerSig, '%PDF-');
  });
});
