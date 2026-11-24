import type {
  PdfHeaderFooterConfig,
  PdfPageNumberConfig,
  PdfPageNumberPosition,
  PdfPageTextClearance,
  PdfWatermarkConfig,
  PdfWatermarkPosition,
} from './pdfTypes';

export interface RawPdfTextItem {
  str?: string;
  transform?: number[];
  width?: number;
  height?: number;
}

export interface TextDrawPlacement {
  text: string;
  /** Baseline X in bottom-up PDF points */
  x: number;
  /** Baseline Y in bottom-up PDF points */
  y: number;
  /** Top-down Y (from top edge to top of text box) for HTML/CSS preview */
  topY: number;
  fontSize: number;
  align: 'left' | 'center' | 'right';
  maxWidth: number;
  measuredWidth: number;
}

export interface WatermarkDrawPlacement {
  text: string;
  /** Center X in bottom-up PDF points */
  centerX: number;
  /** Center Y in bottom-up PDF points */
  centerY: number;
  /** Baseline X for pdf-lib drawText after rotation compensation */
  pdfX: number;
  /** Baseline Y for pdf-lib drawText after rotation compensation */
  pdfY: number;
  /** Top-down center Y for CSS preview */
  previewTopY: number;
  fontSize: number;
  rotationDeg: number;
  opacity: number;
  color: string;
  position: PdfWatermarkPosition;
}

export interface PageReframeBox {
  needed: boolean;
  scale: number;
  offsetX: number;
  offsetY: number;
  topBandPt: number;
  bottomBandPt: number;
}

export type MarginClearanceStatus = 'safe' | 'tight-adjusted' | 'reframed' | 'overlap-risk';

export interface ComputedPageOverlayLayout {
  pageWidth: number;
  pageHeight: number;
  safeMarginX: number;
  topBandPt: number;
  bottomBandPt: number;
  reframe: PageReframeBox;
  clearanceStatus: MarginClearanceStatus;
  clearanceMessage: string;
  headerItems: TextDrawPlacement[];
  footerItems: TextDrawPlacement[];
  pageNumberItem: TextDrawPlacement | null;
  watermarkItem: WatermarkDrawPlacement | null;
}

export type TextWidthMeasurer = (text: string, fontSize: number, bold?: boolean) => number;

const EDGE_PADDING_PT = 10;
const CONTENT_GAP_PT = 8;
const DEFAULT_HORIZONTAL_MARGIN_PT = 36;
const DEFAULT_VERTICAL_MARGIN_PT = 24;

export function sanitizePdfAscii(text: string): string {
  return text.replace(/[^\x20-\x7E]/g, ' ').trim();
}

/**
 * Fallback character-width estimator for environments without pdf-lib font metrics
 * (e.g. fast live preview and unit tests). Closely matches Helvetica / Helvetica-Bold.
 */
export function estimateHelveticaWidth(text: string, fontSize: number, bold = false): number {
  const factor = bold ? 0.59 : 0.54;
  return text.length * fontSize * factor;
}

/**
 * Truncates text with an ellipsis so it never exceeds maxWidth and collides with adjacent slots.
 */
export function fitTextToMaxWidth(
  rawText: string,
  fontSize: number,
  maxWidth: number,
  measureWidth: TextWidthMeasurer = estimateHelveticaWidth,
  bold = false,
): {text: string; width: number} {
  const clean = sanitizePdfAscii(rawText);
  if (!clean) return {text: '', width: 0};
  const fullW = measureWidth(clean, fontSize, bold);
  if (fullW <= maxWidth || maxWidth <= 12) {
    return {text: clean, width: fullW};
  }
  const ellipsis = '...';
  let lo = 1;
  let hi = clean.length;
  let best = ellipsis;
  while (lo <= hi) {
    const mid = Math.floor((lo + hi) / 2);
    const candidate = `${clean.slice(0, mid).trimEnd()}${ellipsis}`;
    if (measureWidth(candidate, fontSize, bold) <= maxWidth) {
      best = candidate;
      lo = mid + 1;
    } else {
      hi = mid - 1;
    }
  }
  return {text: best, width: measureWidth(best, fontSize, bold)};
}

/**
 * Inspects pdf.js getTextContent() items to determine the closest existing text to
 * all four edges of a PDF page (in PDF points).
 */
export function computePageTextClearance(
  items: RawPdfTextItem[],
  pageWidth: number,
  pageHeight: number,
): PdfPageTextClearance {
  let minBottom = pageHeight;
  let maxTop = 0;
  let minLeft = pageWidth;
  let maxRight = 0;
  let hasText = false;

  for (const item of items) {
    const str = (item.str ?? '').trim();
    if (!str) continue;
    const t = item.transform;
    if (!Array.isArray(t) || t.length < 6) continue;

    const tx = Number(t[4]);
    const ty = Number(t[5]);
    if (!Number.isFinite(tx) || !Number.isFinite(ty)) continue;

    const scaleY = Math.hypot(Number(t[2]) || 0, Number(t[3]) || 0);
    const scaleX = Math.hypot(Number(t[0]) || 0, Number(t[1]) || 0);
    const fontHeight = Math.max(4, Math.abs(item.height || 0) || scaleY || scaleX || 10);
    const itemWidth = Math.max(0, Math.abs(item.width || 0) || str.length * fontHeight * 0.5);

    const itemBottom = Math.max(0, ty - fontHeight * 0.25);
    const itemTop = Math.min(pageHeight, ty + fontHeight * 0.95);
    const itemLeft = Math.max(0, tx);
    const itemRight = Math.min(pageWidth, tx + itemWidth);

    hasText = true;
    if (itemBottom < minBottom) minBottom = itemBottom;
    if (itemTop > maxTop) maxTop = itemTop;
    if (itemLeft < minLeft) minLeft = itemLeft;
    if (itemRight > maxRight) maxRight = itemRight;
  }

  if (!hasText) {
    return {
      pageWidth,
      pageHeight,
      topClearancePt: pageHeight,
      bottomClearancePt: pageHeight,
      leftClearancePt: pageWidth,
      rightClearancePt: pageWidth,
      hasText: false,
    };
  }

  return {
    pageWidth,
    pageHeight,
    topClearancePt: Math.max(0, pageHeight - maxTop),
    bottomClearancePt: Math.max(0, minBottom),
    leftClearancePt: Math.max(0, minLeft),
    rightClearancePt: Math.max(0, pageWidth - maxRight),
    hasText: true,
  };
}

export function formatPdfPageNumberLabel(
  pageIndex: number,
  totalPages: number,
  config: PdfPageNumberConfig,
): string {
  const start = Number.isFinite(config.startNumber) ? Math.max(1, Math.floor(config.startNumber)) : 1;
  const num = pageIndex + start;
  const endNum = totalPages + start - 1;
  if (config.format === 'page-of-total') {
    return `Page ${num} of ${endNum}`;
  }
  if (config.format === 'slash') {
    return `${num} / ${endNum}`;
  }
  return `${num}`;
}

export function defaultRotationForWatermarkPosition(pos: PdfWatermarkPosition): number {
  return pos === 'center-diagonal' ? 45 : 0;
}

function slotHorizontalPosition(pos: PdfPageNumberPosition): 'left' | 'center' | 'right' {
  if (pos.endsWith('left')) return 'left';
  if (pos.endsWith('right')) return 'right';
  return 'center';
}

function slotVerticalBand(pos: PdfPageNumberPosition): 'top' | 'bottom' {
  return pos.startsWith('top-') ? 'top' : 'bottom';
}

function buildRowPlacements(
  slots: {left: string; center: string; right: string},
  baselineY: number,
  pageWidth: number,
  pageHeight: number,
  safeMarginX: number,
  fontSize: number,
  measureWidth: TextWidthMeasurer,
): TextDrawPlacement[] {
  const contentWidth = Math.max(60, pageWidth - safeMarginX * 2);
  const gap = 12;
  const hasL = Boolean(sanitizePdfAscii(slots.left));
  const hasC = Boolean(sanitizePdfAscii(slots.center));
  const hasR = Boolean(sanitizePdfAscii(slots.right));
  const activeCount = (hasL ? 1 : 0) + (hasC ? 1 : 0) + (hasR ? 1 : 0);
  if (activeCount === 0) return [];

  const maxSlotW =
    activeCount === 1
      ? contentWidth
      : activeCount === 2
        ? Math.max(40, (contentWidth - gap) / 2)
        : Math.max(36, (contentWidth - gap * 2) / 3);

  const out: TextDrawPlacement[] = [];
  const topY = Math.max(2, pageHeight - baselineY - fontSize);

  if (hasL) {
    const fitted = fitTextToMaxWidth(slots.left, fontSize, maxSlotW, measureWidth, false);
    if (fitted.text) {
      out.push({
        text: fitted.text,
        x: safeMarginX,
        y: baselineY,
        topY,
        fontSize,
        align: 'left',
        maxWidth: maxSlotW,
        measuredWidth: fitted.width,
      });
    }
  }
  if (hasC) {
    const fitted = fitTextToMaxWidth(slots.center, fontSize, maxSlotW, measureWidth, false);
    if (fitted.text) {
      out.push({
        text: fitted.text,
        x: Math.max(safeMarginX, (pageWidth - fitted.width) / 2),
        y: baselineY,
        topY,
        fontSize,
        align: 'center',
        maxWidth: maxSlotW,
        measuredWidth: fitted.width,
      });
    }
  }
  if (hasR) {
    const fitted = fitTextToMaxWidth(slots.right, fontSize, maxSlotW, measureWidth, false);
    if (fitted.text) {
      out.push({
        text: fitted.text,
        x: Math.max(safeMarginX, pageWidth - safeMarginX - fitted.width),
        y: baselineY,
        topY,
        fontSize,
        align: 'right',
        maxWidth: maxSlotW,
        measuredWidth: fitted.width,
      });
    }
  }
  return out;
}

/**
 * Computes exact safe-margin coordinates for Header/Footer, Page Numbers, and Watermark
 * on a single page. Guarantees zero overlap with existing page text when clearances are supplied.
 */
export function computeOverlayLayoutForPage(params: {
  pageIndex: number;
  totalPages: number;
  pageWidth: number;
  pageHeight: number;
  watermark?: PdfWatermarkConfig;
  headerFooter?: PdfHeaderFooterConfig;
  pageNumbers?: PdfPageNumberConfig;
  clearance?: PdfPageTextClearance | null;
  measureWidth?: TextWidthMeasurer;
}): ComputedPageOverlayLayout {
  const {
    pageIndex,
    totalPages,
    pageWidth,
    pageHeight,
    watermark,
    headerFooter,
    pageNumbers,
    clearance,
    measureWidth = estimateHelveticaWidth,
  } = params;

  const safeMarginX = Math.min(DEFAULT_HORIZONTAL_MARGIN_PT, Math.max(18, pageWidth * 0.06));
  const hfActive = Boolean(headerFooter?.enabled && !(headerFooter.skipFirstPage && pageIndex === 0));
  const pnActive = Boolean(pageNumbers?.enabled && !(pageNumbers.skipFirstPage && pageIndex === 0));
  const wmActive = Boolean(
    watermark?.enabled &&
      sanitizePdfAscii(watermark.text) &&
      !(watermark.skipFirstPage && pageIndex === 0),
  );

  const hfFontSize = Math.min(14, Math.max(7, headerFooter?.fontSize ?? 9));
  const pnFontSize = Math.min(14, Math.max(7, pageNumbers?.fontSize ?? 9));

  const hasHeaderContent =
    hfActive &&
    Boolean(
      sanitizePdfAscii(headerFooter?.headerLeft || '') ||
        sanitizePdfAscii(headerFooter?.headerCenter || '') ||
        sanitizePdfAscii(headerFooter?.headerRight || ''),
    );
  const hasFooterContent =
    hfActive &&
    Boolean(
      sanitizePdfAscii(headerFooter?.footerLeft || '') ||
        sanitizePdfAscii(headerFooter?.footerCenter || '') ||
        sanitizePdfAscii(headerFooter?.footerRight || ''),
    );

  const pnBand = pnActive && pageNumbers ? slotVerticalBand(pageNumbers.position) : null;
  const pnAlign = pnActive && pageNumbers ? slotHorizontalPosition(pageNumbers.position) : null;

  const wmPosition: PdfWatermarkPosition = watermark?.position ?? 'center-diagonal';
  const wmInTopBand = wmActive && wmPosition === 'top-margin';
  const wmInBottomBand = wmActive && wmPosition === 'bottom-margin';

  // Detect if Page Number collides with an already-occupied Header/Footer slot in the same band
  const pnCollidesWithHeaderSlot =
    pnBand === 'top' &&
    hasHeaderContent &&
    ((pnAlign === 'left' && Boolean(sanitizePdfAscii(headerFooter?.headerLeft || ''))) ||
      (pnAlign === 'center' && Boolean(sanitizePdfAscii(headerFooter?.headerCenter || ''))) ||
      (pnAlign === 'right' && Boolean(sanitizePdfAscii(headerFooter?.headerRight || ''))));

  const pnCollidesWithFooterSlot =
    pnBand === 'bottom' &&
    hasFooterContent &&
    ((pnAlign === 'left' && Boolean(sanitizePdfAscii(headerFooter?.footerLeft || ''))) ||
      (pnAlign === 'center' && Boolean(sanitizePdfAscii(headerFooter?.footerCenter || ''))) ||
      (pnAlign === 'right' && Boolean(sanitizePdfAscii(headerFooter?.footerRight || ''))));

  const topLineCount =
    (hasHeaderContent ? 1 : 0) +
    (pnBand === 'top' && (!hasHeaderContent || pnCollidesWithHeaderSlot) ? 1 : 0) +
    (wmInTopBand ? 1 : 0);

  const bottomLineCount =
    (hasFooterContent ? 1 : 0) +
    (pnBand === 'bottom' && (!hasFooterContent || pnCollidesWithFooterSlot) ? 1 : 0) +
    (wmInBottomBand ? 1 : 0);

  const desiredTopMargin = Math.max(
    DEFAULT_VERTICAL_MARGIN_PT,
    headerFooter?.marginPt ?? 0,
    pnBand === 'top' ? (pageNumbers?.marginPt ?? 0) : 0,
  );
  const desiredBottomMargin = Math.max(
    DEFAULT_VERTICAL_MARGIN_PT,
    headerFooter?.marginPt ?? 0,
    pnBand === 'bottom' ? (pageNumbers?.marginPt ?? 0) : 0,
  );

  const topFontMax = Math.max(
    hasHeaderContent ? hfFontSize : 0,
    pnBand === 'top' ? pnFontSize : 0,
    wmInTopBand ? Math.min(18, Math.max(10, watermark?.fontSize ?? 13)) : 0,
  );
  const bottomFontMax = Math.max(
    hasFooterContent ? hfFontSize : 0,
    pnBand === 'bottom' ? pnFontSize : 0,
    wmInBottomBand ? Math.min(18, Math.max(10, watermark?.fontSize ?? 13)) : 0,
  );

  const lineStepTop = topFontMax + 4;
  const lineStepBottom = bottomFontMax + 4;

  const minTopBandNeeded =
    topLineCount > 0 ? EDGE_PADDING_PT + topLineCount * lineStepTop + CONTENT_GAP_PT : 0;
  const minBottomBandNeeded =
    bottomLineCount > 0 ? EDGE_PADDING_PT + bottomLineCount * lineStepBottom + CONTENT_GAP_PT : 0;

  const targetTopBand =
    topLineCount > 0 ? Math.max(minTopBandNeeded, desiredTopMargin + (topLineCount - 1) * lineStepTop + 6) : 0;
  const targetBottomBand =
    bottomLineCount > 0
      ? Math.max(minBottomBandNeeded, desiredBottomMargin + (bottomLineCount - 1) * lineStepBottom + 6)
      : 0;

  const ensureSafe =
    (headerFooter?.enabled ? headerFooter.ensureSafeMargins !== false : true) &&
    (pageNumbers?.enabled ? pageNumbers.ensureSafeMargins !== false : true);

  const topClear = clearance?.hasText ? clearance.topClearancePt : pageHeight;
  const bottomClear = clearance?.hasText ? clearance.bottomClearancePt : pageHeight;

  let topBandPt = targetTopBand;
  let bottomBandPt = targetBottomBand;
  let adjusted = false;
  let needsTopReframe = false;
  let needsBottomReframe = false;

  if (topLineCount > 0 && topClear < targetTopBand) {
    if (topClear >= minTopBandNeeded) {
      topBandPt = topClear;
      adjusted = true;
    } else {
      needsTopReframe = true;
      topBandPt = Math.max(minTopBandNeeded, targetTopBand);
    }
  }

  if (bottomLineCount > 0 && bottomClear < targetBottomBand) {
    if (bottomClear >= minBottomBandNeeded) {
      bottomBandPt = bottomClear;
      adjusted = true;
    } else {
      needsBottomReframe = true;
      bottomBandPt = Math.max(minBottomBandNeeded, targetBottomBand);
    }
  }

  const wantsReframe = (needsTopReframe || needsBottomReframe) && ensureSafe;
  let reframe: PageReframeBox = {
    needed: false,
    scale: 1,
    offsetX: 0,
    offsetY: 0,
    topBandPt,
    bottomBandPt,
  };

  if (wantsReframe) {
    const reserveTop = needsTopReframe ? topBandPt : 0;
    const reserveBottom = needsBottomReframe ? bottomBandPt : 0;
    const availH = Math.max(pageHeight * 0.65, pageHeight - reserveTop - reserveBottom);
    const scale = Math.min(1, availH / pageHeight);
    const scaledW = pageWidth * scale;
    const scaledH = pageHeight * scale;
    const offsetX = (pageWidth - scaledW) / 2;
    // Place reframed page strictly between reserveBottom and (pageHeight - reserveTop)
    const remainingSlackY = Math.max(0, pageHeight - reserveTop - reserveBottom - scaledH);
    const offsetY = reserveBottom + remainingSlackY / 2;
    reframe = {
      needed: true,
      scale,
      offsetX,
      offsetY,
      topBandPt,
      bottomBandPt,
    };
  }

  let clearanceStatus: MarginClearanceStatus = 'safe';
  let clearanceMessage = 'Safe margins verified — overlays sit clear of page text.';
  if (needsTopReframe || needsBottomReframe) {
    if (ensureSafe) {
      clearanceStatus = 'reframed';
      clearanceMessage =
        'Page text sits near the edge — safe margin inset is active so overlays never overlap content.';
    } else {
      clearanceStatus = 'overlap-risk';
      clearanceMessage =
        'Existing page text reaches into this margin band. Enable Safe Margin Inset or pick another preset.';
    }
  } else if (adjusted) {
    clearanceStatus = 'tight-adjusted';
    clearanceMessage = 'Tight margin detected — overlay automatically tucked into the clear outer strip.';
  }

  // Compute top band baselines (ordered from page top downward inside safe band)
  let topCursorFromTop =
    topBandPt > 0
      ? Math.max(EDGE_PADDING_PT, Math.min(topBandPt - topLineCount * lineStepTop - 4, (topBandPt - topLineCount * lineStepTop) / 2))
      : EDGE_PADDING_PT;
  topCursorFromTop = Math.max(6, topCursorFromTop);

  let headerItems: TextDrawPlacement[] = [];
  if (hasHeaderContent && headerFooter) {
    const baselineY = Math.max(6, pageHeight - topCursorFromTop - hfFontSize);
    headerItems = buildRowPlacements(
      {
        left: headerFooter.headerLeft,
        center: headerFooter.headerCenter,
        right: headerFooter.headerRight,
      },
      baselineY,
      pageWidth,
      pageHeight,
      safeMarginX,
      hfFontSize,
      measureWidth,
    );
    if (pnCollidesWithHeaderSlot || wmInTopBand) {
      topCursorFromTop += hfFontSize + 4;
    }
  }

  // Compute bottom band baselines (ordered from bottom edge upward inside safe band)
  let bottomCursorFromBottom =
    bottomBandPt > 0
      ? Math.max(EDGE_PADDING_PT, Math.min(bottomBandPt - bottomLineCount * lineStepBottom - 4, (bottomBandPt - bottomLineCount * lineStepBottom) / 2))
      : EDGE_PADDING_PT;
  bottomCursorFromBottom = Math.max(6, bottomCursorFromBottom);

  let footerItems: TextDrawPlacement[] = [];
  if (hasFooterContent && headerFooter) {
    const baselineY = bottomCursorFromBottom;
    footerItems = buildRowPlacements(
      {
        left: headerFooter.footerLeft,
        center: headerFooter.footerCenter,
        right: headerFooter.footerRight,
      },
      baselineY,
      pageWidth,
      pageHeight,
      safeMarginX,
      hfFontSize,
      measureWidth,
    );
    if (pnCollidesWithFooterSlot || wmInBottomBand) {
      bottomCursorFromBottom += hfFontSize + 4;
    }
  }

  // Compute Page Number placement
  let pageNumberItem: TextDrawPlacement | null = null;
  if (pnActive && pageNumbers && pnBand && pnAlign) {
    const label = formatPdfPageNumberLabel(pageIndex, totalPages, pageNumbers);
    const contentWidth = Math.max(60, pageWidth - safeMarginX * 2);
    const maxW =
      (pnBand === 'top' && hasHeaderContent && !pnCollidesWithHeaderSlot) ||
      (pnBand === 'bottom' && hasFooterContent && !pnCollidesWithFooterSlot)
        ? Math.max(60, contentWidth / 3)
        : contentWidth;
    const fitted = fitTextToMaxWidth(label, pnFontSize, maxW, measureWidth, false);
    const baselineY =
      pnBand === 'top'
        ? Math.max(6, pageHeight - topCursorFromTop - pnFontSize)
        : bottomCursorFromBottom;

    let x = safeMarginX;
    if (pnAlign === 'center') {
      x = Math.max(safeMarginX, (pageWidth - fitted.width) / 2);
    } else if (pnAlign === 'right') {
      x = Math.max(safeMarginX, pageWidth - safeMarginX - fitted.width);
    }

    pageNumberItem = {
      text: fitted.text,
      x,
      y: baselineY,
      topY: Math.max(2, pageHeight - baselineY - pnFontSize),
      fontSize: pnFontSize,
      align: pnAlign,
      maxWidth: maxW,
      measuredWidth: fitted.width,
    };

    if (pnBand === 'top' && wmInTopBand) {
      topCursorFromTop += pnFontSize + 4;
    }
    if (pnBand === 'bottom' && wmInBottomBand) {
      bottomCursorFromBottom += pnFontSize + 4;
    }
  }

  // Compute Watermark placement
  let watermarkItem: WatermarkDrawPlacement | null = null;
  if (wmActive && watermark) {
    const cleanText = sanitizePdfAscii(watermark.text);
    const opacity = Math.min(0.55, Math.max(0.05, watermark.opacity ?? 0.16));
    const color = watermark.color || '#64748B';

    let fontSize = watermark.fontSize || 42;
    let rotationDeg = watermark.rotationDeg ?? defaultRotationForWatermarkPosition(wmPosition);
    let centerX = pageWidth / 2;
    let centerY = pageHeight / 2;

    if (wmPosition === 'top-margin') {
      fontSize = Math.min(18, Math.max(10, Math.min(fontSize, 14)));
      rotationDeg = 0;
      const baselineY = Math.max(6, pageHeight - topCursorFromTop - fontSize);
      centerY = baselineY + fontSize / 2;
    } else if (wmPosition === 'bottom-margin') {
      fontSize = Math.min(18, Math.max(10, Math.min(fontSize, 14)));
      rotationDeg = 0;
      centerY = bottomCursorFromBottom + fontSize / 2;
    } else if (wmPosition === 'bottom-right') {
      fontSize = Math.min(24, Math.max(10, Math.min(fontSize, 18)));
      rotationDeg = 0;
      const w = measureWidth(cleanText, fontSize, true);
      centerX = Math.max(safeMarginX + w / 2, pageWidth - safeMarginX - w / 2);
      centerY = Math.max(EDGE_PADDING_PT + fontSize / 2, (bottomBandPt > 0 ? bottomBandPt + 8 : 22) + fontSize / 2);
    } else if (wmPosition === 'center-horizontal') {
      rotationDeg = 0;
    }

    // Clamp font size so watermark never overflows page width
    const maxSpan =
      wmPosition === 'center-diagonal'
        ? Math.hypot(pageWidth - safeMarginX * 2, pageHeight - safeMarginX * 2) * 0.62
        : (pageWidth - safeMarginX * 2) * 0.92;
    const rawW = measureWidth(cleanText, fontSize, true);
    if (rawW > maxSpan && rawW > 0) {
      fontSize = Math.max(9, Math.floor((fontSize * maxSpan) / rawW));
    }

    const finalW = measureWidth(cleanText, fontSize, true);
    const finalH = fontSize * 0.72;
    const rad = (rotationDeg * Math.PI) / 180;
    // Exact rotation-compensated baseline origin so text visual center lands precisely on (centerX, centerY)
    const pdfX = centerX - (finalW / 2) * Math.cos(rad) + (finalH / 2) * Math.sin(rad);
    const pdfY = centerY - (finalW / 2) * Math.sin(rad) - (finalH / 2) * Math.cos(rad);

    watermarkItem = {
      text: cleanText,
      centerX,
      centerY,
      pdfX,
      pdfY,
      previewTopY: pageHeight - centerY,
      fontSize,
      rotationDeg,
      opacity,
      color,
      position: wmPosition,
    };
  }

  return {
    pageWidth,
    pageHeight,
    safeMarginX,
    topBandPt,
    bottomBandPt,
    reframe,
    clearanceStatus,
    clearanceMessage,
    headerItems,
    footerItems,
    pageNumberItem,
    watermarkItem,
  };
}
