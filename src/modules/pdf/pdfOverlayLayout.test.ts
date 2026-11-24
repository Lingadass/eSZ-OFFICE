import {test} from 'vitest';
import assert from 'node:assert/strict';
import {PDFDocument} from 'pdf-lib';
import {
  computeOverlayLayoutForPage,
  computePageTextClearance,
  formatPdfPageNumberLabel,
} from './pdfOverlayLayout';
import {bakeAnnotationsAndOverlays, textToPdfBytes} from './pdfPure';
import type {PdfPageNumberPosition} from './pdfTypes';

test('computePageTextClearance measures top and bottom clearance and ignores whitespace', () => {
  const clearance = computePageTextClearance(
    [
      {str: '   ', transform: [12, 0, 0, 12, 10, 835], width: 20, height: 12},
      {str: 'Top Heading', transform: [14, 0, 0, 14, 54, 780], width: 120, height: 14},
      {str: 'Bottom Line', transform: [11, 0, 0, 11, 54, 60], width: 110, height: 11},
    ],
    595,
    842,
  );

  assert.equal(clearance.hasText, true);
  // 842 - (780 + 14 * 0.95) = 48.7
  assert.ok(clearance.topClearancePt > 45 && clearance.topClearancePt < 52);
  // 60 - 11 * 0.25 = 57.25
  assert.ok(clearance.bottomClearancePt > 54 && clearance.bottomClearancePt < 60);
});

test('all 6 page number presets place labels strictly inside safe page bounds', () => {
  const positions: PdfPageNumberPosition[] = [
    'top-left',
    'top-center',
    'top-right',
    'bottom-left',
    'bottom-center',
    'bottom-right',
  ];

  for (const pos of positions) {
    const layout = computeOverlayLayoutForPage({
      pageIndex: 1,
      totalPages: 5,
      pageWidth: 595,
      pageHeight: 842,
      pageNumbers: {
        enabled: true,
        position: pos,
        format: 'page-of-total',
        startNumber: 1,
      },
      clearance: {
        pageWidth: 595,
        pageHeight: 842,
        topClearancePt: 72,
        bottomClearancePt: 72,
        leftClearancePt: 54,
        rightClearancePt: 54,
        hasText: true,
      },
    });

    assert.equal(layout.clearanceStatus, 'safe');
    assert.ok(layout.pageNumberItem);
    assert.equal(layout.pageNumberItem.text, 'Page 2 of 5');
    assert.ok(layout.pageNumberItem.x >= layout.safeMarginX - 0.1);
    assert.ok(
      layout.pageNumberItem.x + layout.pageNumberItem.measuredWidth <=
        595 - layout.safeMarginX + 0.5,
    );
    if (pos.startsWith('top-')) {
      assert.ok(layout.pageNumberItem.y > 842 - 72);
    } else {
      assert.ok(layout.pageNumberItem.y + layout.pageNumberItem.fontSize < 72);
    }
  }
});

test('header/footer slots truncate cleanly so left, center, and right never overlap', () => {
  const layout = computeOverlayLayoutForPage({
    pageIndex: 0,
    totalPages: 3,
    pageWidth: 595,
    pageHeight: 842,
    headerFooter: {
      enabled: true,
      headerLeft: 'Very Long Company Confidentiality Notice For Left Slot',
      headerCenter: 'Very Long Document Title That Would Overlap Without Truncation',
      headerRight: 'Very Long Classification Label For Right Slot',
      footerLeft: '',
      footerCenter: '',
      footerRight: '',
    },
  });

  assert.equal(layout.headerItems.length, 3);
  const [left, center, right] = layout.headerItems;
  assert.ok(left.x + left.measuredWidth < center.x, 'left header must not overlap center');
  assert.ok(center.x + center.measuredWidth < right.x, 'center header must not overlap right');
  assert.ok(right.x + right.measuredWidth <= 595 - layout.safeMarginX + 1);
});

test('page number stacks cleanly when colliding with occupied footer slot', () => {
  const layout = computeOverlayLayoutForPage({
    pageIndex: 0,
    totalPages: 4,
    pageWidth: 595,
    pageHeight: 842,
    headerFooter: {
      enabled: true,
      headerLeft: '',
      headerCenter: '',
      headerRight: '',
      footerLeft: '',
      footerCenter: 'Quarterly Report',
      footerRight: '',
    },
    pageNumbers: {
      enabled: true,
      position: 'bottom-center',
      format: 'slash',
      startNumber: 1,
    },
  });

  assert.equal(layout.footerItems.length, 1);
  assert.ok(layout.pageNumberItem);
  assert.equal(layout.pageNumberItem.text, '1 / 4');
  assert.ok(
    Math.abs(layout.pageNumberItem.y - layout.footerItems[0].y) >= 10,
    'colliding page number and footer center must be vertically separated',
  );
});

test('edge-to-edge existing text triggers safe-margin reframing instead of overlapping', () => {
  const tightClearance = {
    pageWidth: 595,
    pageHeight: 842,
    topClearancePt: 6,
    bottomClearancePt: 5,
    leftClearancePt: 10,
    rightClearancePt: 10,
    hasText: true,
  };

  const withSafe = computeOverlayLayoutForPage({
    pageIndex: 0,
    totalPages: 2,
    pageWidth: 595,
    pageHeight: 842,
    pageNumbers: {
      enabled: true,
      position: 'bottom-center',
      format: 'page-of-total',
      startNumber: 1,
      ensureSafeMargins: true,
    },
    clearance: tightClearance,
  });

  assert.equal(withSafe.clearanceStatus, 'reframed');
  assert.equal(withSafe.reframe.needed, true);
  assert.ok(withSafe.reframe.scale < 1);
  assert.ok(withSafe.reframe.offsetY >= withSafe.bottomBandPt - 0.1);

  const withoutSafe = computeOverlayLayoutForPage({
    pageIndex: 0,
    totalPages: 2,
    pageWidth: 595,
    pageHeight: 842,
    pageNumbers: {
      enabled: true,
      position: 'bottom-center',
      format: 'page-of-total',
      startNumber: 1,
      ensureSafeMargins: false,
    },
    clearance: tightClearance,
  });

  assert.equal(withoutSafe.clearanceStatus, 'overlap-risk');
  assert.equal(withoutSafe.reframe.needed, false);
});

test('watermark position presets clamp size and support safe top/bottom margins and skipFirstPage', () => {
  const skipped = computeOverlayLayoutForPage({
    pageIndex: 0,
    totalPages: 3,
    pageWidth: 595,
    pageHeight: 842,
    watermark: {
      enabled: true,
      text: 'CONFIDENTIAL',
      opacity: 0.2,
      fontSize: 48,
      rotationDeg: 45,
      color: '#64748B',
      position: 'center-diagonal',
      skipFirstPage: true,
    },
  });
  assert.equal(skipped.watermarkItem, null);

  const topWm = computeOverlayLayoutForPage({
    pageIndex: 1,
    totalPages: 3,
    pageWidth: 595,
    pageHeight: 842,
    watermark: {
      enabled: true,
      text: 'INTERNAL DISTRIBUTION ONLY',
      opacity: 0.2,
      fontSize: 48,
      rotationDeg: 45,
      color: '#64748B',
      position: 'top-margin',
    },
  });
  assert.ok(topWm.watermarkItem);
  assert.equal(topWm.watermarkItem.rotationDeg, 0);
  assert.ok(topWm.watermarkItem.centerY > 800, 'top-margin watermark must sit in top safe band');
  assert.ok(topWm.watermarkItem.fontSize <= 18);
});

test('bakeAnnotationsAndOverlays succeeds with vector safe-margin reframing on tight pages', async () => {
  const samplePdf = await textToPdfBytes(
    'Line 1\nLine 2\nLine 3\nLine 4\nLine 5',
    'Tight Margin Sample',
  );
  const baked = await bakeAnnotationsAndOverlays(samplePdf, [], {
    headerFooter: {
      enabled: true,
      headerLeft: 'Spec v3',
      headerCenter: 'Safe Margin Test',
      headerRight: '2026',
      footerLeft: 'eSZ OFFICE',
      footerCenter: '',
      footerRight: 'Confidential',
      ensureSafeMargins: true,
    },
    pageNumbers: {
      enabled: true,
      position: 'bottom-center',
      format: 'page-of-total',
      startNumber: 1,
      ensureSafeMargins: true,
    },
    watermark: {
      enabled: true,
      text: 'DRAFT',
      opacity: 0.15,
      fontSize: 42,
      rotationDeg: 45,
      color: '#64748B',
      position: 'center-diagonal',
    },
    pageClearances: [
      {
        pageWidth: 595,
        pageHeight: 842,
        topClearancePt: 5,
        bottomClearancePt: 5,
        leftClearancePt: 12,
        rightClearancePt: 12,
        hasText: true,
      },
    ],
  });

  const doc = await PDFDocument.load(baked);
  assert.equal(doc.getPageCount(), 1);
  assert.equal(formatPdfPageNumberLabel(0, 3, {enabled: true, position: 'bottom-center', format: 'page-of-total', startNumber: 1}), 'Page 1 of 3');
});
