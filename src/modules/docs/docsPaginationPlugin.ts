import {Extension} from '@tiptap/core';
import {Plugin, PluginKey} from '@tiptap/pm/state';
import {Decoration, DecorationSet, type EditorView} from '@tiptap/pm/view';
import type {Node as PMNode} from '@tiptap/pm/model';
import {
  computePaginationPlan,
  getPageMetrics,
  type MeasurableBlock,
  type MeasurableListItem,
  type MeasurableParagraphSplit,
  type MeasurableTableRow,
  type PaginationBreak,
  type PaginationPlan,
} from './docsPagination';
import type {DocNoteItem, DocPageLayout, DocViewMode} from './docsTypes';

export const PaginationPluginKey = new PluginKey<PaginationPluginState>('eszDocsPagination');

interface PaginationPluginState {
  decorations: DecorationSet;
  signature: string;
}

export interface PaginationRuntimeConfig {
  enabled: boolean;
  layout: DocPageLayout;
  notes: DocNoteItem[];
  zoom: number;
  viewMode: DocViewMode;
  onPlanChange: (plan: PaginationPlan) => void;
}

function collectFootnoteIds(node: PMNode): string[] {
  const ids: string[] = [];
  node.descendants((child) => {
    if (child.type.name === 'noteRef' && child.attrs?.kind === 'footnote' && child.attrs?.noteId) {
      const id = String(child.attrs.noteId);
      if (!ids.includes(id)) ids.push(id);
    }
  });
  return ids;
}

function parsePx(value: string | null | undefined, fallback = 0): number {
  if (!value) return fallback;
  const n = Number.parseFloat(value);
  return Number.isFinite(n) ? n : fallback;
}

function getLiveTransformScale(container: HTMLElement): number {
  const ow = container.offsetWidth;
  if (ow > 0) {
    const rw = container.getBoundingClientRect().width;
    if (rw > 0) return rw / ow;
  }
  return 1;
}

function getNaturalBlockHeight(dom: HTMLElement, liveScale: number): number {
  const safeScale = Math.max(0.25, liveScale);
  const raw = dom.getBoundingClientRect().height / safeScale;

  // Subtract exact unscaled inline heights of any existing pagination spacer widgets nested inside this block
  const spacers = dom.querySelectorAll<HTMLElement>(
    '.esz-page-spacer, .esz-page-spacer-inline, .esz-page-spacer-row, .esz-page-spacer-li',
  );
  let spacerSum = 0;
  spacers.forEach((s) => {
    if (s.classList.contains('esz-page-spacer-row')) {
      const td = s.querySelector<HTMLElement>('td');
      spacerSum += parsePx(td?.style.height, s.offsetHeight);
    } else {
      spacerSum += parsePx(s.style.height, s.offsetHeight);
    }
  });
  let measured = Math.max(0, raw - spacerSum);

  // Include declared image height fallback when an image has not finished decoding yet
  if (dom.tagName === 'IMG') {
    const img = dom as HTMLImageElement;
    if (!img.complete || measured < 2) {
      const declared = parsePx(img.getAttribute('data-height') || img.getAttribute('height'), 0);
      if (declared > measured) measured = declared;
    }
  } else {
    const nestedImgs = dom.querySelectorAll<HTMLImageElement>('img');
    nestedImgs.forEach((img) => {
      if (!img.complete) {
        const declared = parsePx(img.getAttribute('data-height') || img.getAttribute('height'), 0);
        if (declared > 0 && img.offsetHeight < declared) {
          measured += declared - img.offsetHeight;
        }
      }
    });
  }

  return measured;
}

/**
 * Computes effective bottom margin of a block including any CSS margin that collapses
 * out of its last child element (e.g. <p> inside <blockquote> or <li> inside <ul>).
 */
function getEffectiveMarginBottom(dom: HTMLElement, style: CSSStyleDeclaration): number {
  let mb = parsePx(style.marginBottom, 0);
  const padBottom = parsePx(style.paddingBottom, 0);
  const borderBottom = parsePx(style.borderBottomWidth, 0);
  if (padBottom === 0 && borderBottom === 0 && dom.lastElementChild instanceof HTMLElement) {
    const lastChild = dom.lastElementChild;
    if (!lastChild.classList.contains('esz-page-spacer') && !lastChild.classList.contains('esz-page-spacer-li')) {
      const childStyle = window.getComputedStyle(lastChild);
      mb = Math.max(mb, parsePx(childStyle.marginBottom, 0));
    }
  }
  return mb;
}

/**
 * Resolves a clean ProseMirror character offset inside a paragraph where a page break
 * should occur so at most `availablePx` from `fromOffsetY` stays on the current page.
 * Uses binary search over `view.coordsAtPos` (O(log N) calls) and snaps to word boundaries.
 */
function createParagraphSplitResolver(
  view: EditorView,
  paraDom: HTMLElement,
  paraNode: PMNode,
  paraStartPos: number,
  liveScale: number,
): (availablePx: number, fromOffsetY: number) => MeasurableParagraphSplit | null {
  const safeScale = Math.max(0.25, liveScale);
  return (availablePx: number, fromOffsetY: number): MeasurableParagraphSplit | null => {
    const contentStart = paraStartPos + 1;
    const contentEnd = paraStartPos + paraNode.nodeSize - 1;
    if (contentEnd - contentStart < 4) return null;

    const paraRect = paraDom.getBoundingClientRect();
    // Collect any existing inline spacers inside this paragraph so we map visual Y -> natural Y
    const existingSpacers: Array<{top: number; height: number}> = [];
    paraDom.querySelectorAll<HTMLElement>('.esz-page-spacer-inline').forEach((sp) => {
      const r = sp.getBoundingClientRect();
      existingSpacers.push({
        top: (r.top - paraRect.top) / safeScale,
        height: parsePx(sp.style.height, sp.offsetHeight),
      });
    });

    const naturalYOfPos = (pos: number): {top: number; bottom: number} | null => {
      try {
        const c = view.coordsAtPos(pos);
        const relTop = (c.top - paraRect.top) / safeScale;
        const relBottom = (c.bottom - paraRect.top) / safeScale;
        let spacerAbove = 0;
        for (const sp of existingSpacers) {
          if (relTop >= sp.top - 1) {
            spacerAbove += sp.height;
          }
        }
        return {
          top: Math.max(0, relTop - spacerAbove),
          bottom: Math.max(0, relBottom - spacerAbove),
        };
      } catch {
        return null;
      }
    };

    const targetMaxBottom = fromOffsetY + availablePx;

    // Binary search for the largest pos whose character bottom <= targetMaxBottom
    let lo = contentStart;
    let hi = contentEnd;
    let bestPos = contentStart;

    while (lo <= hi) {
      const mid = Math.floor((lo + hi) / 2);
      const metrics = naturalYOfPos(mid);
      if (!metrics) {
        hi = mid - 1;
        continue;
      }
      if (metrics.bottom <= targetMaxBottom + 0.5) {
        bestPos = mid;
        lo = mid + 1;
      } else {
        hi = mid - 1;
      }
    }

    if (bestPos <= contentStart || bestPos >= contentEnd) {
      return null;
    }

    // Advance bestPos to the first position on the NEXT visual line
    const bestCoords = naturalYOfPos(bestPos);
    if (!bestCoords) return null;

    let nextLinePos = bestPos;
    const scanLimit = Math.min(contentEnd, bestPos + 64);
    for (let p = bestPos; p < scanLimit; p++) {
      const c = naturalYOfPos(p);
      if (c && c.top >= bestCoords.bottom - 2) {
        nextLinePos = p;
        break;
      }
    }

    // If we didn't find a strictly new line, snap cleanly to the nearest preceding space
    if (nextLinePos === bestPos) {
      const textBefore = view.state.doc.textBetween(contentStart, bestPos, ' ', ' ');
      const lastSpace = textBefore.lastIndexOf(' ');
      if (lastSpace > 4) {
        nextLinePos = contentStart + lastSpace + 1;
      }
    }

    if (nextLinePos <= contentStart || nextLinePos >= contentEnd) {
      return null;
    }

    const prevCoords = naturalYOfPos(Math.max(contentStart, nextLinePos - 1));
    const nextCoords = naturalYOfPos(nextLinePos);
    const offsetY =
      prevCoords && nextCoords && nextCoords.top >= prevCoords.bottom - 1
        ? (prevCoords.bottom + nextCoords.top) / 2
        : prevCoords
          ? prevCoords.bottom
          : targetMaxBottom;

    if (offsetY <= fromOffsetY + 4) return null;
    return {pos: nextLinePos, offsetY};
  };
}

function buildDecorations(doc: PMNode, breaks: PaginationBreak[]): DecorationSet {
  if (breaks.length === 0) return DecorationSet.empty;
  const decos: Decoration[] = [];

  for (const br of breaks) {
    if (br.pos < 0 || br.pos > doc.content.size) continue;

    if (br.kind === 'table-row') {
      const widget = Decoration.widget(
        br.pos,
        () => {
          const tr = document.createElement('tr');
          tr.className = 'esz-page-spacer-row';
          tr.setAttribute('aria-hidden', 'true');
          tr.setAttribute('contenteditable', 'false');
          tr.setAttribute('data-page-break-spacer', String(br.toPage + 1));
          const td = document.createElement('td');
          td.colSpan = Math.max(1, br.tableColCount ?? 1);
          td.style.height = `${br.spacerHeight}px`;
          td.style.padding = '0';
          td.style.border = 'none';
          td.style.background = 'transparent';
          tr.appendChild(td);
          return tr;
        },
        {
          side: -1,
          key: `row-br-${br.fromPage}-${br.toPage}-${br.spacerHeight}`,
          ignoreSelection: true,
        },
      );
      decos.push(widget);
    } else if (br.kind === 'list-item') {
      const widget = Decoration.widget(
        br.pos,
        () => {
          const li = document.createElement('li');
          li.className = 'esz-page-spacer-li';
          li.setAttribute('aria-hidden', 'true');
          li.setAttribute('contenteditable', 'false');
          li.setAttribute('data-page-break-spacer', String(br.toPage + 1));
          li.style.height = `${br.spacerHeight}px`;
          return li;
        },
        {
          side: -1,
          key: `li-br-${br.fromPage}-${br.toPage}-${br.spacerHeight}`,
          ignoreSelection: true,
        },
      );
      decos.push(widget);
    } else if (br.kind === 'inline') {
      const widget = Decoration.widget(
        br.pos,
        () => {
          const span = document.createElement('span');
          span.className = 'esz-page-spacer-inline';
          span.setAttribute('aria-hidden', 'true');
          span.setAttribute('contenteditable', 'false');
          span.setAttribute('data-page-break-spacer', String(br.toPage + 1));
          span.style.height = `${br.spacerHeight}px`;
          return span;
        },
        {
          side: -1,
          key: `inline-br-${br.fromPage}-${br.toPage}-${br.spacerHeight}`,
          ignoreSelection: true,
        },
      );
      decos.push(widget);
    } else {
      const widget = Decoration.widget(
        br.pos,
        () => {
          const div = document.createElement('div');
          div.className = 'esz-page-spacer';
          div.setAttribute('aria-hidden', 'true');
          div.setAttribute('contenteditable', 'false');
          div.setAttribute('data-page-break-spacer', String(br.toPage + 1));
          div.style.height = `${br.spacerHeight}px`;
          return div;
        },
        {
          side: br.manual ? 1 : -1,
          key: `block-br-${br.fromPage}-${br.toPage}-${br.spacerHeight}`,
          ignoreSelection: true,
        },
      );
      decos.push(widget);
    }
  }

  return DecorationSet.create(doc, decos);
}

export function measureAndPaginateEditor(
  view: EditorView,
  config: PaginationRuntimeConfig,
): PaginationPlan {
  const metrics = getPageMetrics(config.layout);
  const notesById = new Map<string, DocNoteItem>();
  for (const n of config.notes) {
    notesById.set(n.id, n);
  }

  if (!config.enabled || config.viewMode !== 'print') {
    const emptyPlan = computePaginationPlan([], metrics, notesById);
    const current = PaginationPluginKey.getState(view.state);
    if (current && current.signature !== 'disabled') {
      const tr = view.state.tr
        .setMeta(PaginationPluginKey, {
          decorations: DecorationSet.empty,
          signature: 'disabled',
        })
        .setMeta('addToHistory', false);
      view.dispatch(tr);
    }
    config.onPlanChange(emptyPlan);
    return emptyPlan;
  }

  const doc = view.state.doc;
  const blocks: MeasurableBlock[] = [];
  // Derive live transform scale strictly from DOM so pagination measurements are invariant at 50%/100%/110%/150% zoom
  const liveScale = getLiveTransformScale(view.dom);

  doc.forEach((node, offset) => {
    const pos = offset;
    const endPos = offset + node.nodeSize;
    const domNode = view.nodeDOM(pos);
    if (!(domNode instanceof HTMLElement)) return;

    const style = window.getComputedStyle(domNode);
    let marginTop = parsePx(style.marginTop, 0);
    // If CSS `.esz-page-spacer + *` is currently zeroing margin-top, preserve declared spaceBefore for reflow checks
    if (marginTop === 0 && domNode.previousElementSibling?.classList.contains('esz-page-spacer')) {
      const spaceBeforePt = parsePx(domNode.getAttribute('data-space-before'), 0);
      if (spaceBeforePt > 0) {
        marginTop = spaceBeforePt * (96 / 72);
      }
    }
    const marginBottom = getEffectiveMarginBottom(domNode, style);
    const fontSize = parsePx(style.fontSize, 15);
    const lineHeight = parsePx(style.lineHeight, fontSize * 1.5);
    const height = getNaturalBlockHeight(domNode, liveScale);
    const footnoteIds = collectFootnoteIds(node);

    const typeName = node.type.name;
    let kind: MeasurableBlock['kind'] = 'other';
    if (typeName === 'pageBreak') kind = 'pageBreak';
    else if (typeName === 'heading') kind = 'heading';
    else if (typeName === 'paragraph') kind = 'paragraph';
    else if (typeName === 'table') kind = 'table';
    else if (typeName === 'bulletList' || typeName === 'orderedList') kind = 'list';

    let tableRows: MeasurableTableRow[] | undefined;
    let tableColCount: number | undefined;
    if (kind === 'table') {
      tableRows = [];
      let maxCols = 1;
      let rowPos = pos + 1;
      node.forEach((rowNode) => {
        if (rowNode.type.name === 'tableRow') {
          let rowCols = 0;
          rowNode.forEach((cellNode) => {
            rowCols += Math.max(1, Number(cellNode.attrs?.colspan ?? 1));
          });
          if (rowCols > maxCols) maxCols = rowCols;
          const rowDom = view.nodeDOM(rowPos);
          const rowHeight =
            rowDom instanceof HTMLElement
              ? getNaturalBlockHeight(rowDom, liveScale)
              : 32;
          const firstCell = rowNode.firstChild;
          const isHeader = firstCell?.type.name === 'tableHeader';
          tableRows!.push({
            pos: rowPos,
            height: Math.max(20, rowHeight),
            isHeader,
            footnoteIds: collectFootnoteIds(rowNode),
          });
        }
        rowPos += rowNode.nodeSize;
      });
      tableColCount = maxCols;
    }

    let listItems: MeasurableListItem[] | undefined;
    if (kind === 'list') {
      listItems = [];
      let itemPos = pos + 1;
      node.forEach((itemNode) => {
        if (itemNode.type.name === 'listItem') {
          const itemDom = view.nodeDOM(itemPos);
          if (itemDom instanceof HTMLElement) {
            const itemStyle = window.getComputedStyle(itemDom);
            listItems!.push({
              pos: itemPos,
              height: Math.max(18, getNaturalBlockHeight(itemDom, liveScale)),
              marginTop: parsePx(itemStyle.marginTop, 0),
              marginBottom: getEffectiveMarginBottom(itemDom, itemStyle),
              footnoteIds: collectFootnoteIds(itemNode),
            });
          }
        }
        itemPos += itemNode.nodeSize;
      });
    }

    blocks.push({
      pos,
      endPos,
      kind,
      height: Math.max(kind === 'pageBreak' ? 0 : 12, height),
      marginTop,
      marginBottom,
      lineHeight,
      footnoteIds: footnoteIds.length > 0 ? footnoteIds : undefined,
      tableRows,
      tableColCount,
      listItems: listItems && listItems.length > 0 ? listItems : undefined,
      resolveParagraphSplit:
        kind === 'paragraph' && height > lineHeight * 3 && !domNode.querySelector('img')
          ? createParagraphSplitResolver(view, domNode, node, pos, liveScale)
          : undefined,
    });
  });

  const plan = computePaginationPlan(blocks, metrics, notesById);
  const current = PaginationPluginKey.getState(view.state);

  if (!current || current.signature !== plan.signature) {
    const decorations = buildDecorations(doc, plan.breaks);
    const tr = view.state.tr
      .setMeta(PaginationPluginKey, {
        decorations,
        signature: plan.signature,
      })
      .setMeta('addToHistory', false);
    view.dispatch(tr);
  }

  config.onPlanChange(plan);
  return plan;
}

export interface PaginationExtensionOptions {
  getConfig: () => PaginationRuntimeConfig;
}

export const PaginationExtension = Extension.create<PaginationExtensionOptions>({
  name: 'eszPagination',

  addOptions() {
    return {
      getConfig: () => {
        throw new Error('PaginationExtension requires getConfig option');
      },
    };
  },

  addProseMirrorPlugins() {
    const getConfig = this.options.getConfig;
    let rafId: number | null = null;
    let timerId: number | null = null;

    const schedulePass = (view: EditorView, immediate = false) => {
      if (timerId !== null) {
        window.clearTimeout(timerId);
        timerId = null;
      }
      if (rafId !== null) {
        window.cancelAnimationFrame(rafId);
        rafId = null;
      }
      const run = () => {
        rafId = window.requestAnimationFrame(() => {
          rafId = null;
          if (view.isDestroyed) return;
          try {
            measureAndPaginateEditor(view, getConfig());
          } catch {
            // Avoid throwing if DOM is mid-unmount
          }
        });
      };

      if (immediate) {
        run();
      } else {
        timerId = window.setTimeout(run, 50);
      }
    };

    return [
      new Plugin<PaginationPluginState>({
        key: PaginationPluginKey,
        state: {
          init() {
            return {
              decorations: DecorationSet.empty,
              signature: '',
            };
          },
          apply(tr, prev) {
            const meta = tr.getMeta(PaginationPluginKey) as
              | PaginationPluginState
              | undefined;
            if (meta) {
              return meta;
            }
            if (tr.docChanged) {
              return {
                decorations: prev.decorations.map(tr.mapping, tr.doc),
                signature: prev.signature,
              };
            }
            return prev;
          },
        },
        props: {
          decorations(state) {
            const decos = PaginationPluginKey.getState(state)?.decorations;
            if (!decos || decos === DecorationSet.empty) return null;
            return decos;
          },
        },
        view(editorView) {
          schedulePass(editorView, true);

          // Re-paginate whenever any embedded image finishes loading/decoding
          const onMediaLoad = (e: Event) => {
            if (e.target instanceof HTMLImageElement) {
              schedulePass(editorView, true);
            }
          };
          editorView.dom.addEventListener('load', onMediaLoad, true);

          // Re-paginate whenever web fonts finish loading
          const onFontsLoaded = () => {
            if (!editorView.isDestroyed) schedulePass(editorView, true);
          };
          const fontSet = typeof document !== 'undefined' ? document.fonts : undefined;
          fontSet?.addEventListener?.('loadingdone', onFontsLoaded);
          void fontSet?.ready?.then(onFontsLoaded);

          // Observe non-spacer DOM height changes (e.g. image decode, table reflow)
          let resizeObserver: ResizeObserver | null = null;
          if (typeof ResizeObserver !== 'undefined') {
            resizeObserver = new ResizeObserver(() => {
              if (!editorView.isDestroyed) {
                schedulePass(editorView, false);
              }
            });
            resizeObserver.observe(editorView.dom);
          }

          return {
            update(view, prevState) {
              if (!view.state.doc.eq(prevState.doc)) {
                schedulePass(view, false);
              }
            },
            destroy() {
              editorView.dom.removeEventListener('load', onMediaLoad, true);
              fontSet?.removeEventListener?.('loadingdone', onFontsLoaded);
              resizeObserver?.disconnect();
              if (timerId !== null) window.clearTimeout(timerId);
              if (rafId !== null) window.cancelAnimationFrame(rafId);
            },
          };
        },
      }),
    ];
  },
});
