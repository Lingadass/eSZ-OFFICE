import {
  DOC_MARGIN_PX,
  DOC_PAGE_DIMENSIONS_PX,
  type DocNoteItem,
  type DocPageLayout,
} from './docsTypes';

export const DOC_PAGE_GAP_PX = 24;
export const DOC_FOOTNOTE_HEADER_PX = 18;
export const DOC_FOOTNOTE_ROW_PX = 24;
export const DOC_FOOTNOTE_PADDING_PX = 12;
export const DOC_BOTTOM_SAFETY_PX = 8;
export const MIN_LINES_TOGETHER = 2;

export interface PageMetrics {
  pageWidth: number;
  pageHeight: number;
  marginTop: number;
  marginBottom: number;
  marginLeft: number;
  marginRight: number;
  pageGap: number;
  contentWidth: number;
  contentHeight: number;
  columns: 1 | 2 | 3;
}

export interface MeasurableTableRow {
  pos: number;
  height: number;
  isHeader?: boolean;
  footnoteIds?: string[];
}

export interface MeasurableListItem {
  pos: number;
  height: number;
  marginTop: number;
  marginBottom: number;
  footnoteIds?: string[];
}

export interface MeasurableParagraphSplit {
  /** ProseMirror position right before the first character of the continuation line */
  pos: number;
  /** Offset in unscaled px from top of the paragraph to the bottom of the last line on the current page */
  offsetY: number;
}

export interface MeasurableBlock {
  /** ProseMirror start position of the top-level block */
  pos: number;
  /** ProseMirror end position of the top-level block */
  endPos: number;
  kind: 'paragraph' | 'heading' | 'table' | 'pageBreak' | 'list' | 'other';
  /** Natural unscaled border-box height (excluding any existing pagination spacers) */
  height: number;
  marginTop: number;
  marginBottom: number;
  /** Estimated visual line height in px (for paragraph splitting) */
  lineHeight?: number;
  /** Footnote IDs referenced inside this block */
  footnoteIds?: string[];
  /** Table rows for row-level breaking when kind === 'table' */
  tableRows?: MeasurableTableRow[];
  /** Exact column count of a table so row spacer widgets match colspan without altering table-layout: fixed */
  tableColCount?: number;
  /** List items for item-level breaking when kind === 'list' */
  listItems?: MeasurableListItem[];
  /**
   * Optional callback to resolve an exact inline ProseMirror split position
   * inside a long paragraph given how many pixels fit on the current page
   * (already respecting minimum orphan/widow line count).
   */
  resolveParagraphSplit?: (
    availablePx: number,
    fromOffsetY: number,
  ) => MeasurableParagraphSplit | null;
}

export interface PaginationBreak {
  /** ProseMirror position where the spacer widget should be inserted */
  pos: number;
  /** Spacer widget type */
  kind: 'block' | 'inline' | 'table-row' | 'list-item';
  /** Exact unscaled height in px of the spacer widget */
  spacerHeight: number;
  /** 0-based page index that this break transitions FROM */
  fromPage: number;
  /** 0-based page index that this break transitions TO */
  toPage: number;
  /** Whether triggered by an explicit manual PageBreak node */
  manual?: boolean;
  /** Exact table column count when kind === 'table-row' */
  tableColCount?: number;
}

export interface PaginationPlan {
  pageCount: number;
  breaks: PaginationBreak[];
  /** Footnotes keyed by 0-based page index where they are referenced */
  pageFootnotes: Record<number, DocNoteItem[]>;
  /** Endnotes displayed at the end of the document */
  endnotes: DocNoteItem[];
  /** Unscaled Y offset of each page card top in the page stack */
  pageTops: number[];
  /** Total unscaled height of the full physical page stack */
  totalStackHeight: number;
  /** Compact structural signature for flicker-free diffing */
  signature: string;
}

export function getPageMetrics(layout: DocPageLayout): PageMetrics {
  const dims = DOC_PAGE_DIMENSIONS_PX[layout.size] ?? DOC_PAGE_DIMENSIONS_PX.a4;
  const presetWidth = layout.orientation === 'landscape' ? dims.height : dims.width;
  const presetHeight = layout.orientation === 'landscape' ? dims.width : dims.height;
  const pageWidth =
    layout.customPageWidthPx && layout.customPageWidthPx >= 240
      ? Math.round(layout.customPageWidthPx)
      : presetWidth;
  const pageHeight =
    layout.customPageHeightPx && layout.customPageHeightPx >= 240
      ? Math.round(layout.customPageHeightPx)
      : presetHeight;
  const presetMargins = DOC_MARGIN_PX[layout.margins] ?? DOC_MARGIN_PX.normal;
  const margins = layout.customMarginsPx
    ? {
        top: Math.max(12, Math.round(layout.customMarginsPx.top)),
        right: Math.max(12, Math.round(layout.customMarginsPx.right)),
        bottom: Math.max(12, Math.round(layout.customMarginsPx.bottom)),
        left: Math.max(12, Math.round(layout.customMarginsPx.left)),
      }
    : presetMargins;

  return {
    pageWidth,
    pageHeight,
    marginTop: margins.top,
    marginBottom: margins.bottom,
    marginLeft: margins.left,
    marginRight: margins.right,
    pageGap: DOC_PAGE_GAP_PX,
    contentWidth: Math.max(120, pageWidth - margins.left - margins.right),
    contentHeight: Math.max(160, pageHeight - margins.top - margins.bottom),
    columns: layout.columns,
  };
}

export function formatPageNumber(pageNumber: number, totalPages: number): string {
  const safeTotal = Math.max(1, Math.floor(totalPages));
  const safeCurrent = Math.min(safeTotal, Math.max(1, Math.floor(pageNumber)));
  return `Page ${safeCurrent} of ${safeTotal}`;
}

export function shouldShowPageChrome(
  pageIndex: number,
  differentFirstPage?: boolean,
): boolean {
  if (differentFirstPage && pageIndex === 0) return false;
  return true;
}

export function estimateFootnotesBoxHeight(footnotes: DocNoteItem[]): number {
  if (footnotes.length === 0) return 0;
  return DOC_FOOTNOTE_HEADER_PX + DOC_FOOTNOTE_PADDING_PX + footnotes.length * DOC_FOOTNOTE_ROW_PX;
}

/**
 * Pure deterministic page-break calculator.
 * Measures sequential blocks against exact physical page boundaries and returns:
 * - Widget spacer positions & pixel-exact heights (zero cumulative drift across 100+ pages)
 * - Manual Page Break support
 * - Heading keep-with-next (prevents orphaned headings)
 * - Long paragraph splitting with 2-line orphan/widow protection
 * - Table row-level breaking (keeps header row with first body row)
 * - Per-page footnote placement & height reservation
 */
export function computePaginationPlan(
  blocks: MeasurableBlock[],
  metrics: PageMetrics,
  notesById: Map<string, DocNoteItem>,
): PaginationPlan {
  const {pageHeight, pageGap, contentHeight, columns} = metrics;
  const pageStride = pageHeight + pageGap;

  // Multi-column pages pack `columns` columns side-by-side inside each page's contentHeight.
  // Because each break spacer uses `column-span: all`, we approximate un-split capacity per page
  // by multiplying contentHeight by columns (with a small 6% balancing safety margin when columns > 1).
  const columnFactor = columns > 1 ? columns * 0.94 : 1;

  const breaks: PaginationBreak[] = [];
  const pageFootnoteIds: Record<number, string[]> = {};
  const pageFootnotes: Record<number, DocNoteItem[]> = {};

  const getPageContentStart = (pageIdx: number) => pageIdx * pageStride;

  const getFootnotesForIds = (ids: string[]): DocNoteItem[] => {
    const out: DocNoteItem[] = [];
    for (const id of ids) {
      const note = notesById.get(id);
      if (note && note.kind === 'footnote') {
        out.push(note);
      }
    }
    return out;
  };

  const getPageFootnoteHeight = (pageIdx: number, additionalIds?: string[]): number => {
    const existing = pageFootnoteIds[pageIdx] ?? [];
    if (!additionalIds || additionalIds.length === 0) {
      return estimateFootnotesBoxHeight(getFootnotesForIds(existing));
    }
    const merged = [...existing];
    for (const id of additionalIds) {
      if (!merged.includes(id)) merged.push(id);
    }
    return estimateFootnotesBoxHeight(getFootnotesForIds(merged));
  };

  const registerPageFootnotes = (pageIdx: number, ids?: string[]) => {
    if (!ids || ids.length === 0) return;
    if (!pageFootnoteIds[pageIdx]) pageFootnoteIds[pageIdx] = [];
    for (const id of ids) {
      const note = notesById.get(id);
      if (note && note.kind === 'footnote' && !pageFootnoteIds[pageIdx].includes(id)) {
        pageFootnoteIds[pageIdx].push(id);
      }
    }
  };

  const getUsableHeight = (pageIdx: number, extraFootnoteIds?: string[]): number => {
    const fnH = getPageFootnoteHeight(pageIdx, extraFootnoteIds);
    return Math.max(80, (contentHeight - fnH - DOC_BOTTOM_SAFETY_PX) * columnFactor);
  };

  let currentPage = 0;
  // Unscaled Y position of cursor inside the current page's content area [0 .. usableHeight]
  let yInPage = 0;
  let prevMarginBottom = 0;

  const pushBlockBreak = (
    pos: number,
    manual = false,
    extraBottomOffset = 0,
  ) => {
    const absCursor = getPageContentStart(currentPage) + yInPage + extraBottomOffset;
    const nextPage = currentPage + 1;
    const targetStart = getPageContentStart(nextPage);
    // For multi-column pages, visual vertical height used is yInPage / columnFactor
    const effectiveYInPage = columns > 1 ? Math.min(contentHeight, yInPage / columnFactor) : yInPage;
    const effectiveAbs = getPageContentStart(currentPage) + effectiveYInPage + extraBottomOffset;
    const spacerHeight = Math.max(
      metrics.marginBottom + pageGap + metrics.marginTop,
      Math.round(targetStart - effectiveAbs),
    );
    breaks.push({
      pos,
      kind: 'block',
      spacerHeight: columns > 1 ? Math.round(spacerHeight) : Math.max(0, Math.round(targetStart - absCursor)),
      fromPage: currentPage,
      toPage: nextPage,
      manual,
    });
    currentPage = nextPage;
    yInPage = 0;
    prevMarginBottom = 0;
  };

  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i];

    // 1. Manual Page Break node
    if (block.kind === 'pageBreak') {
      // If an automatic overflow break JUST landed us at yInPage === 0 on a fresh page,
      // absorb an immediate redundant manual page break so we don't produce a blank empty gap.
      const prevBreak = breaks.length > 0 ? breaks[breaks.length - 1] : null;
      if (yInPage === 0 && prevBreak && !prevBreak.manual) {
        continue;
      }
      const collapsedTop = yInPage === 0 ? 0 : Math.max(prevMarginBottom, block.marginTop);
      const breakEndY = yInPage + collapsedTop + block.height;
      // Place spacer immediately after the PageBreak node so the next block lands at top of next page
      const absEnd =
        getPageContentStart(currentPage) +
        (columns > 1 ? Math.min(contentHeight, breakEndY / columnFactor) : breakEndY) +
        block.marginBottom;
      const nextPage = currentPage + 1;
      const targetStart = getPageContentStart(nextPage);
      const spacerHeight = Math.max(0, Math.round(targetStart - absEnd));
      breaks.push({
        pos: block.endPos,
        kind: 'block',
        spacerHeight,
        fromPage: currentPage,
        toPage: nextPage,
        manual: true,
      });
      currentPage = nextPage;
      yInPage = 0;
      prevMarginBottom = 0;
      continue;
    }

    const collapsedMarginTop = yInPage === 0 ? 0 : Math.max(prevMarginBottom, block.marginTop);
    const blockTopInPage = yInPage + collapsedMarginTop;
    const usable = getUsableHeight(currentPage, block.footnoteIds);

    // 2. Heading Keep-with-Next rule: avoid orphaning a heading at the bottom of a page
    if (block.kind === 'heading' && yInPage > 0) {
      const nextBlock = blocks[i + 1];
      let nextLookahead = 0;
      if (nextBlock) {
        if (nextBlock.kind === 'table' && nextBlock.tableRows && nextBlock.tableRows.length > 0) {
          const firstGroupCount =
            nextBlock.tableRows.length >= 2 && nextBlock.tableRows[0].isHeader ? 2 : 1;
          nextLookahead =
            nextBlock.tableRows
              .slice(0, firstGroupCount)
              .reduce((sum, r) => sum + r.height, 0) +
            Math.max(block.marginBottom, nextBlock.marginTop);
        } else {
          nextLookahead =
            Math.min(
              nextBlock.height,
              Math.max(36, (nextBlock.lineHeight ?? 22) * MIN_LINES_TOGETHER),
            ) + Math.max(block.marginBottom, nextBlock.marginTop);
        }
      }
      if (blockTopInPage + block.height + nextLookahead > usable) {
        pushBlockBreak(block.pos, false, prevMarginBottom);
      }
    }

    // Recompute after possible heading break
    const curCollapsedTop = yInPage === 0 ? 0 : Math.max(prevMarginBottom, block.marginTop);
    const curTopInPage = yInPage + curCollapsedTop;
    const curUsable = getUsableHeight(currentPage, block.footnoteIds);

    // 3. Block fits completely on current page
    if (curTopInPage + block.height <= curUsable) {
      yInPage = curTopInPage + block.height;
      prevMarginBottom = block.marginBottom;
      registerPageFootnotes(currentPage, block.footnoteIds);
      continue;
    }

    // 4. Table row-level breaking (when single-column and rows are provided)
    if (block.kind === 'table' && columns === 1 && block.tableRows && block.tableRows.length > 0) {
      const rows = block.tableRows;
      // Check if at least the first row (or header + first data row) fits on the current page
      const firstGroupCount = rows.length >= 2 && rows[0].isHeader ? 2 : 1;
      const firstGroupHeight = rows
        .slice(0, firstGroupCount)
        .reduce((sum, r) => sum + r.height, 0);

      if (yInPage > 0 && curTopInPage + firstGroupHeight > curUsable) {
        // Move entire table start to next page so we don't orphan a table header
        pushBlockBreak(block.pos, false, prevMarginBottom);
      }

      let tableY = yInPage === 0 ? 0 : yInPage + Math.max(prevMarginBottom, block.marginTop);
      for (let rIdx = 0; rIdx < rows.length; rIdx++) {
        const row = rows[rIdx];
        const rowUsable = getUsableHeight(currentPage, row.footnoteIds);
        if (rIdx > 0 && tableY > 0 && tableY + row.height > rowUsable) {
          const absRowTop = getPageContentStart(currentPage) + tableY;
          const nextPage = currentPage + 1;
          const targetStart = getPageContentStart(nextPage);
          breaks.push({
            pos: row.pos,
            kind: 'table-row',
            spacerHeight: Math.max(0, Math.round(targetStart - absRowTop)),
            fromPage: currentPage,
            toPage: nextPage,
            tableColCount: block.tableColCount,
          });
          currentPage = nextPage;
          tableY = 0;
        }
        const activeRowUsable = getUsableHeight(currentPage, row.footnoteIds);
        if (row.height > activeRowUsable && tableY === 0) {
          // Ultra-tall single table row exceeding one full page: advance cleanly along physical pageStride
          const absRowEnd = getPageContentStart(currentPage) + row.height;
          const endPage = Math.floor(absRowEnd / pageStride);
          const offsetInEndPage = absRowEnd - endPage * pageStride;
          currentPage = endPage;
          tableY = offsetInEndPage;
        } else {
          tableY += row.height;
        }
        registerPageFootnotes(currentPage, row.footnoteIds);
      }
      yInPage = tableY;
      prevMarginBottom = block.marginBottom;
      continue;
    }

    // 4b. List item-level breaking (for long bulleted/numbered lists)
    if (block.kind === 'list' && columns === 1 && block.listItems && block.listItems.length > 0) {
      const items = block.listItems;
      if (yInPage > 0 && curTopInPage + items[0].height > curUsable) {
        pushBlockBreak(block.pos, false, prevMarginBottom);
      }

      let listY = yInPage === 0 ? 0 : yInPage + Math.max(prevMarginBottom, block.marginTop);
      let prevItemMarginBottom = 0;

      for (let itemIdx = 0; itemIdx < items.length; itemIdx++) {
        const item = items[itemIdx];
        const itemCollapsedTop =
          itemIdx === 0 ? 0 : Math.max(prevItemMarginBottom, item.marginTop);
        const itemUsable = getUsableHeight(currentPage, item.footnoteIds);
        if (itemIdx > 0 && listY + itemCollapsedTop + item.height > itemUsable && listY > 0) {
          const absItemBreak = getPageContentStart(currentPage) + listY + prevItemMarginBottom;
          const nextPage = currentPage + 1;
          const targetStart = getPageContentStart(nextPage);
          breaks.push({
            pos: item.pos,
            kind: 'list-item',
            spacerHeight: Math.max(0, Math.round(targetStart - absItemBreak)),
            fromPage: currentPage,
            toPage: nextPage,
          });
          currentPage = nextPage;
          listY = 0;
          prevItemMarginBottom = 0;
        }
        const topGap = listY === 0 ? 0 : Math.max(prevItemMarginBottom, item.marginTop);
        listY += topGap + item.height;
        prevItemMarginBottom = item.marginBottom;
        registerPageFootnotes(currentPage, item.footnoteIds);
      }

      yInPage = listY;
      prevMarginBottom = Math.max(block.marginBottom, prevItemMarginBottom);
      continue;
    }

    // 5. Long paragraph splitting (keep at least 2 lines together — widow/orphan control)
    if (block.kind === 'paragraph' && columns === 1 && block.resolveParagraphSplit) {
      const lineH = Math.max(14, block.lineHeight ?? 22);
      const totalLines = Math.max(1, Math.round(block.height / lineH));
      const minChunkHeight = lineH * MIN_LINES_TOGETHER;
      const remainingOnCur = curUsable - curTopInPage;

      // If fewer than 2 lines fit on the current page (or total paragraph <= 3 lines), move whole paragraph
      if (yInPage > 0 && (remainingOnCur < minChunkHeight || totalLines < MIN_LINES_TOGETHER * 2)) {
        pushBlockBreak(block.pos, false, prevMarginBottom);
      }

      let paraTopInPage = yInPage === 0 ? 0 : yInPage + Math.max(prevMarginBottom, block.marginTop);
      let consumedPx = 0;
      let safety = 0;

      while (safety < 80) {
        safety += 1;
        const pageUsable = getUsableHeight(currentPage, block.footnoteIds);
        const remainingParaPx = Math.max(0, block.height - consumedPx);
        const spaceLeft = pageUsable - paraTopInPage;

        if (remainingParaPx <= spaceLeft + 1) {
          yInPage = paraTopInPage + remainingParaPx;
          prevMarginBottom = block.marginBottom;
          registerPageFootnotes(currentPage, block.footnoteIds);
          break;
        }

        // Keep at least MIN_LINES_TOGETHER lines on the next page (widow control)
        const maxTakeForWidow = Math.max(0, remainingParaPx - minChunkHeight);
        const allowedTake = Math.min(spaceLeft, maxTakeForWidow);

        if (allowedTake < minChunkHeight) {
          // Cannot satisfy 2-line orphan/widow rule on this page
          if (consumedPx === 0 && yInPage > 0) {
            pushBlockBreak(block.pos, false, prevMarginBottom);
            paraTopInPage = 0;
            continue;
          }
          // Already at top of page and still cannot split cleanly — allow forced fit on current page
          const split = block.resolveParagraphSplit(spaceLeft, consumedPx);
          if (!split || split.offsetY <= consumedPx + 1) {
            yInPage = paraTopInPage + remainingParaPx;
            prevMarginBottom = block.marginBottom;
            registerPageFootnotes(currentPage, block.footnoteIds);
            break;
          }
          const sliceHeight = split.offsetY - consumedPx;
          const absSplitY = getPageContentStart(currentPage) + paraTopInPage + sliceHeight;
          const nextPage = currentPage + 1;
          const targetStart = getPageContentStart(nextPage);
          breaks.push({
            pos: split.pos,
            kind: 'inline',
            spacerHeight: Math.max(0, Math.round(targetStart - absSplitY)),
            fromPage: currentPage,
            toPage: nextPage,
          });
          registerPageFootnotes(currentPage, block.footnoteIds);
          currentPage = nextPage;
          paraTopInPage = 0;
          consumedPx = split.offsetY;
          continue;
        }

        const split = block.resolveParagraphSplit(allowedTake, consumedPx);
        if (!split || split.offsetY <= consumedPx + 1) {
          // Could not locate an inline split point -> move whole block if not at top of page
          if (consumedPx === 0 && yInPage > 0) {
            pushBlockBreak(block.pos, false, prevMarginBottom);
            paraTopInPage = 0;
            continue;
          }
          yInPage = paraTopInPage + remainingParaPx;
          prevMarginBottom = block.marginBottom;
          registerPageFootnotes(currentPage, block.footnoteIds);
          break;
        }

        const sliceHeight = split.offsetY - consumedPx;
        const absSplitY = getPageContentStart(currentPage) + paraTopInPage + sliceHeight;
        const nextPage = currentPage + 1;
        const targetStart = getPageContentStart(nextPage);
        breaks.push({
          pos: split.pos,
          kind: 'inline',
          spacerHeight: Math.max(0, Math.round(targetStart - absSplitY)),
          fromPage: currentPage,
          toPage: nextPage,
        });
        registerPageFootnotes(currentPage, block.footnoteIds);
        currentPage = nextPage;
        paraTopInPage = 0;
        consumedPx = split.offsetY;
      }
      continue;
    }

    // 6. Standard overflowing block -> move cleanly to the next page (if not already at top)
    if (yInPage > 0) {
      pushBlockBreak(block.pos, false, prevMarginBottom);
    }

    // Place block at top of new page; if it still exceeds a full single page, advance using physical pageStride
    const freshUsable = getUsableHeight(currentPage, block.footnoteIds);
    registerPageFootnotes(currentPage, block.footnoteIds);
    if (block.height <= freshUsable) {
      yInPage = block.height;
      prevMarginBottom = block.marginBottom;
    } else {
      const absBlockEnd = getPageContentStart(currentPage) + block.height;
      const endPage = Math.floor(absBlockEnd / pageStride);
      const offsetInEndPage = absBlockEnd - endPage * pageStride;
      currentPage = endPage;
      yInPage = offsetInEndPage;
      prevMarginBottom = block.marginBottom;
    }
  }

  // Finalize per-page footnote maps
  for (const [pageIdxStr, ids] of Object.entries(pageFootnoteIds)) {
    const idx = Number(pageIdxStr);
    const items = getFootnotesForIds(ids);
    if (items.length > 0) {
      pageFootnotes[idx] = items;
    }
  }

  const endnotes = Array.from(notesById.values())
    .filter((n) => n.kind === 'endnote')
    .sort((a, b) => a.number - b.number);

  // If endnotes exist and wouldn't fit on the final page, add a clean page for endnotes
  let pageCount = Math.max(1, currentPage + 1);
  if (endnotes.length > 0) {
    const endnotesHeight = estimateFootnotesBoxHeight(endnotes) + 12;
    const finalUsable = getUsableHeight(pageCount - 1);
    const effectiveFinalY = columns > 1 ? yInPage / columnFactor : yInPage;
    if (effectiveFinalY > 0 && effectiveFinalY + endnotesHeight > finalUsable) {
      pageCount += 1;
    }
  }

  const pageTops = Array.from({length: pageCount}, (_, idx) => idx * pageStride);
  const totalStackHeight = pageCount * pageHeight + Math.max(0, pageCount - 1) * pageGap;

  const signature = `${pageCount}|${pageWidthKey(metrics)}|${breaks
    .map((b) => `${b.pos}:${b.kind}:${b.spacerHeight}`)
    .join(',')}|fn:${Object.entries(pageFootnotes)
    .map(([p, list]) => `${p}=${list.map((n) => n.id).join('.')}`)
    .join(';')}|en:${endnotes.length}`;

  return {
    pageCount,
    breaks,
    pageFootnotes,
    endnotes,
    pageTops,
    totalStackHeight,
    signature,
  };
}

function pageWidthKey(m: PageMetrics): string {
  return `${m.pageWidth}x${m.pageHeight}:${m.marginTop}-${m.marginBottom}-${m.columns}`;
}
