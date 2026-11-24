import {Extension, Mark, Node, mergeAttributes} from '@tiptap/core';
import {FontFamily} from '@tiptap/extension-text-style';
import Image from '@tiptap/extension-image';
import {Plugin, PluginKey} from '@tiptap/pm/state';
import {Table, TableCell, TableHeader} from '@tiptap/extension-table';
import {canonicalizeWordFontFamily, resolveRenderFontStack} from './docFonts';
import type {DocTabStop} from './docsTypes';

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    indent: {
      increaseIndent: () => ReturnType;
      decreaseIndent: () => ReturnType;
    };
    pageBreak: {
      insertPageBreak: () => ReturnType;
    };
    docTab: {
      insertDocTab: () => ReturnType;
    };
    docShape: {
      insertDocShape: (attrs: {
        shapeType: 'rectangle' | 'rounded' | 'circle' | 'callout';
        label: string;
        fill: string;
        border: string;
      }) => ReturnType;
    };
    docEquation: {
      insertDocEquation: (formula: string) => ReturnType;
    };
    docTextBox: {
      insertDocTextBox: (content?: string) => ReturnType;
    };
    docCaption: {
      insertDocCaption: (label: string) => ReturnType;
    };
    bookmark: {
      setBookmark: (name: string) => ReturnType;
    };
  }
}

/**
 * TipTap FontFamily extension that keeps the canonical Word font name (e.g. "Times New Roman",
 * "Calibri") in the document model and HTML `data-doc-font`, while rendering with our bundled
 * offline metric-compatible SIL OFL font stack.
 */
export const DocFontFamily = FontFamily.extend({
  addGlobalAttributes() {
    return [
      {
        types: this.options.types,
        attributes: {
          fontFamily: {
            default: null,
            parseHTML: (element) => {
              const explicit = element.getAttribute('data-doc-font');
              if (explicit) return canonicalizeWordFontFamily(explicit);
              const raw = element.style.fontFamily;
              return raw ? canonicalizeWordFontFamily(raw) : null;
            },
            renderHTML: (attributes) => {
              const canonical = canonicalizeWordFontFamily(attributes.fontFamily as string | null);
              if (!canonical) return {};
              return {
                'data-doc-font': canonical,
                style: `font-family: ${resolveRenderFontStack(canonical)}`,
              };
            },
          },
        },
      },
    ];
  },
});

/**
 * Preserves Word run typography marks (all-caps, small-caps, character spacing) on TextStyle spans.
 */
export const DocRunStyleExtension = Extension.create({
  name: 'docRunStyle',

  addGlobalAttributes() {
    return [
      {
        types: ['textStyle'],
        attributes: {
          textTransform: {
            default: null,
            parseHTML: (el) => el.style.textTransform || null,
            renderHTML: (attrs) =>
              attrs.textTransform ? {style: `text-transform: ${attrs.textTransform}`} : {},
          },
          fontVariant: {
            default: null,
            parseHTML: (el) => el.style.fontVariant || null,
            renderHTML: (attrs) =>
              attrs.fontVariant ? {style: `font-variant: ${attrs.fontVariant}`} : {},
          },
          letterSpacing: {
            default: null,
            parseHTML: (el) => el.style.letterSpacing || null,
            renderHTML: (attrs) =>
              attrs.letterSpacing ? {style: `letter-spacing: ${attrs.letterSpacing}`} : {},
          },
        },
      },
    ];
  },
});

/**
 * Preserves OOXML paragraph attributes: styleId, spacing before/after, line spacing,
 * exact left/right/first-line indents, keep-with-next, and keep-lines.
 */
export const DocParagraphStyleExtension = Extension.create({
  name: 'docParagraphStyle',

  addGlobalAttributes() {
    return [
      {
        types: ['paragraph', 'heading'],
        attributes: {
          styleId: {
            default: null,
            parseHTML: (el) => el.getAttribute('data-style-id') || null,
            renderHTML: (attrs) => (attrs.styleId ? {'data-style-id': attrs.styleId} : {}),
          },
          spaceBeforePt: {
            default: null,
            parseHTML: (el) => {
              const raw = el.getAttribute('data-space-before');
              if (raw === null) return null;
              const n = Number(raw);
              return Number.isFinite(n) ? n : null;
            },
            renderHTML: (attrs) =>
              attrs.spaceBeforePt !== null && attrs.spaceBeforePt !== undefined
                ? {
                    'data-space-before': String(attrs.spaceBeforePt),
                    style: `margin-top: ${attrs.spaceBeforePt}pt`,
                  }
                : {},
          },
          spaceAfterPt: {
            default: null,
            parseHTML: (el) => {
              const raw = el.getAttribute('data-space-after');
              if (raw === null) return null;
              const n = Number(raw);
              return Number.isFinite(n) ? n : null;
            },
            renderHTML: (attrs) =>
              attrs.spaceAfterPt !== null && attrs.spaceAfterPt !== undefined
                ? {
                    'data-space-after': String(attrs.spaceAfterPt),
                    style: `margin-bottom: ${attrs.spaceAfterPt}pt`,
                  }
                : {},
          },
          lineSpacing: {
            default: null,
            parseHTML: (el) => el.getAttribute('data-line-spacing') || el.style.lineHeight || null,
            renderHTML: (attrs) =>
              attrs.lineSpacing
                ? {
                    'data-line-spacing': String(attrs.lineSpacing),
                    style: `line-height: ${attrs.lineSpacing}`,
                  }
                : {},
          },
          indentLeftPx: {
            default: null,
            parseHTML: (el) => {
              const raw = el.getAttribute('data-indent-left');
              if (raw === null) return null;
              const n = Number(raw);
              return Number.isFinite(n) && n > 0 ? n : null;
            },
            renderHTML: (attrs) =>
              attrs.indentLeftPx
                ? {
                    'data-indent-left': String(attrs.indentLeftPx),
                    style: `margin-left: ${attrs.indentLeftPx}px`,
                  }
                : {},
          },
          indentRightPx: {
            default: null,
            parseHTML: (el) => {
              const raw = el.getAttribute('data-indent-right');
              if (raw === null) return null;
              const n = Number(raw);
              return Number.isFinite(n) && n > 0 ? n : null;
            },
            renderHTML: (attrs) =>
              attrs.indentRightPx
                ? {
                    'data-indent-right': String(attrs.indentRightPx),
                    style: `margin-right: ${attrs.indentRightPx}px`,
                  }
                : {},
          },
          firstLineIndentPx: {
            default: null,
            parseHTML: (el) => {
              const raw = el.getAttribute('data-first-line-indent');
              if (raw === null) return null;
              const n = Number(raw);
              return Number.isFinite(n) && n !== 0 ? n : null;
            },
            renderHTML: (attrs) =>
              attrs.firstLineIndentPx
                ? {
                    'data-first-line-indent': String(attrs.firstLineIndentPx),
                    style: `text-indent: ${attrs.firstLineIndentPx}px`,
                  }
                : {},
          },
          keepNext: {
            default: false,
            parseHTML: (el) => el.getAttribute('data-keep-next') === 'true',
            renderHTML: (attrs) =>
              attrs.keepNext
                ? {'data-keep-next': 'true', style: 'break-after: avoid; page-break-after: avoid'}
                : {},
          },
          keepLines: {
            default: false,
            parseHTML: (el) => el.getAttribute('data-keep-lines') === 'true',
            renderHTML: (attrs) =>
              attrs.keepLines
                ? {'data-keep-lines': 'true', style: 'break-inside: avoid'}
                : {},
          },
          tabStops: {
            default: null,
            parseHTML: (el) => {
              const raw = el.getAttribute('data-tab-stops');
              if (!raw) return null;
              try {
                const parsed = JSON.parse(raw);
                return Array.isArray(parsed) && parsed.length > 0 ? parsed : null;
              } catch {
                return null;
              }
            },
            renderHTML: (attrs) =>
              Array.isArray(attrs.tabStops) && attrs.tabStops.length > 0
                ? {'data-tab-stops': JSON.stringify(attrs.tabStops)}
                : {},
          },
        },
      },
    ];
  },
});

/**
 * Preserves each imported OOXML body block's original XML (`data-ooxml-source`), group membership
 * (`data-ooxml-group`, `data-ooxml-gi`, `data-ooxml-gc`), and canonical import signature
 * (`data-ooxml-sig`) on TipTap block nodes so lossless .docx save can write untouched blocks back
 * byte-for-byte unchanged and regenerate XML only for edited or newly inserted blocks.
 */
export const DocOoxmlPreserveExtension = Extension.create({
  name: 'docOoxmlPreserve',

  addGlobalAttributes() {
    return [
      {
        types: [
          'paragraph',
          'heading',
          'pageBreak',
          'image',
          'table',
          'docTextBox',
          'docLockedElement',
          'docEquation',
        ],
        attributes: {
          ooxmlSource: {
            default: null,
            parseHTML: (el) => el.getAttribute('data-ooxml-source') || null,
            renderHTML: (attrs) =>
              attrs.ooxmlSource ? {'data-ooxml-source': String(attrs.ooxmlSource)} : {},
          },
          ooxmlGroup: {
            default: null,
            parseHTML: (el) => el.getAttribute('data-ooxml-group') || null,
            renderHTML: (attrs) =>
              attrs.ooxmlGroup ? {'data-ooxml-group': String(attrs.ooxmlGroup)} : {},
          },
          ooxmlGroupIndex: {
            default: null,
            parseHTML: (el) => el.getAttribute('data-ooxml-gi') || null,
            renderHTML: (attrs) =>
              attrs.ooxmlGroupIndex !== null && attrs.ooxmlGroupIndex !== undefined
                ? {'data-ooxml-gi': String(attrs.ooxmlGroupIndex)}
                : {},
          },
          ooxmlGroupCount: {
            default: null,
            parseHTML: (el) => el.getAttribute('data-ooxml-gc') || null,
            renderHTML: (attrs) =>
              attrs.ooxmlGroupCount !== null && attrs.ooxmlGroupCount !== undefined
                ? {'data-ooxml-gc': String(attrs.ooxmlGroupCount)}
                : {},
          },
          ooxmlSig: {
            default: null,
            parseHTML: (el) => el.getAttribute('data-ooxml-sig') || null,
            renderHTML: (attrs) =>
              attrs.ooxmlSig ? {'data-ooxml-sig': String(attrs.ooxmlSig)} : {},
          },
        },
      },
    ];
  },
});

/**
 * Extended Table / TableCell / TableHeader preserving fixed OOXML widths, alignment,
 * cell borders, shading, padding, vertical alignment, and merges.
 */
export const DocTable = Table.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      tableAlign: {
        default: null,
        parseHTML: (el) => el.getAttribute('data-table-align') || null,
        renderHTML: (attrs) => {
          if (!attrs.tableAlign) return {};
          const marginStyle =
            attrs.tableAlign === 'center'
              ? 'margin-left: auto; margin-right: auto;'
              : attrs.tableAlign === 'right'
                ? 'margin-left: auto; margin-right: 0;'
                : '';
          return {
            'data-table-align': attrs.tableAlign,
            ...(marginStyle ? {style: marginStyle} : {}),
          };
        },
      },
      tableWidthPx: {
        default: null,
        parseHTML: (el) => {
          const raw = el.getAttribute('data-table-width');
          const n = raw ? Number(raw) : NaN;
          return Number.isFinite(n) && n > 0 ? n : null;
        },
        renderHTML: (attrs) =>
          attrs.tableWidthPx
            ? {
                'data-table-width': String(attrs.tableWidthPx),
                style: `width: ${attrs.tableWidthPx}px; max-width: 100%;`,
              }
            : {},
      },
    };
  },
});

function buildCellExtraAttributes() {
  return {
    backgroundColor: {
      default: null,
      parseHTML: (el: HTMLElement) =>
        el.getAttribute('data-bg-color') || el.style.backgroundColor || null,
      renderHTML: (attrs: Record<string, unknown>) =>
        attrs.backgroundColor
          ? {
              'data-bg-color': String(attrs.backgroundColor),
              style: `background-color: ${attrs.backgroundColor}`,
            }
          : {},
    },
    verticalAlign: {
      default: null,
      parseHTML: (el: HTMLElement) =>
        el.getAttribute('data-valign') || el.style.verticalAlign || null,
      renderHTML: (attrs: Record<string, unknown>) =>
        attrs.verticalAlign
          ? {
              'data-valign': String(attrs.verticalAlign),
              style: `vertical-align: ${attrs.verticalAlign}`,
            }
          : {},
    },
    borderTop: {
      default: null,
      parseHTML: (el: HTMLElement) => el.getAttribute('data-border-top') || null,
      renderHTML: (attrs: Record<string, unknown>) =>
        attrs.borderTop
          ? {'data-border-top': String(attrs.borderTop), style: `border-top: ${attrs.borderTop}`}
          : {},
    },
    borderRight: {
      default: null,
      parseHTML: (el: HTMLElement) => el.getAttribute('data-border-right') || null,
      renderHTML: (attrs: Record<string, unknown>) =>
        attrs.borderRight
          ? {
              'data-border-right': String(attrs.borderRight),
              style: `border-right: ${attrs.borderRight}`,
            }
          : {},
    },
    borderBottom: {
      default: null,
      parseHTML: (el: HTMLElement) => el.getAttribute('data-border-bottom') || null,
      renderHTML: (attrs: Record<string, unknown>) =>
        attrs.borderBottom
          ? {
              'data-border-bottom': String(attrs.borderBottom),
              style: `border-bottom: ${attrs.borderBottom}`,
            }
          : {},
    },
    borderLeft: {
      default: null,
      parseHTML: (el: HTMLElement) => el.getAttribute('data-border-left') || null,
      renderHTML: (attrs: Record<string, unknown>) =>
        attrs.borderLeft
          ? {
              'data-border-left': String(attrs.borderLeft),
              style: `border-left: ${attrs.borderLeft}`,
            }
          : {},
    },
    cellPadding: {
      default: null,
      parseHTML: (el: HTMLElement) => el.getAttribute('data-cell-padding') || null,
      renderHTML: (attrs: Record<string, unknown>) =>
        attrs.cellPadding
          ? {
              'data-cell-padding': String(attrs.cellPadding),
              style: `padding: ${attrs.cellPadding}`,
            }
          : {},
    },
  };
}

export const DocTableCell = TableCell.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      ...buildCellExtraAttributes(),
    };
  },
});

export const DocTableHeader = TableHeader.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      ...buildCellExtraAttributes(),
    };
  },
});

/**
 * Extended Image node preserving imported OOXML width, height, horizontal alignment, and floating state.
 */
export const DocImage = Image.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      widthPx: {
        default: null,
        parseHTML: (el) => {
          const raw = el.getAttribute('data-width') || el.getAttribute('width');
          const n = raw ? Number.parseFloat(raw) : NaN;
          return Number.isFinite(n) && n > 0 ? Math.round(n) : null;
        },
        renderHTML: (attrs) =>
          attrs.widthPx
            ? {'data-width': String(attrs.widthPx), style: `width: ${attrs.widthPx}px; max-width: 100%;`}
            : {},
      },
      heightPx: {
        default: null,
        parseHTML: (el) => {
          const raw = el.getAttribute('data-height') || el.getAttribute('height');
          const n = raw ? Number.parseFloat(raw) : NaN;
          return Number.isFinite(n) && n > 0 ? Math.round(n) : null;
        },
        renderHTML: (attrs) =>
          attrs.heightPx
            ? {'data-height': String(attrs.heightPx), style: `height: ${attrs.heightPx}px; object-fit: contain;`}
            : {},
      },
      align: {
        default: null,
        parseHTML: (el) => el.getAttribute('data-align') || null,
        renderHTML: (attrs) => {
          if (attrs.align === 'center') {
            return {
              'data-align': 'center',
              style: 'display: block; margin-left: auto; margin-right: auto;',
            };
          }
          if (attrs.align === 'right') {
            return {
              'data-align': 'right',
              style: 'display: block; margin-left: auto; margin-right: 0;',
            };
          }
          return attrs.align ? {'data-align': String(attrs.align)} : {};
        },
      },
      floating: {
        default: false,
        parseHTML: (el) => el.getAttribute('data-floating') === 'true',
        renderHTML: (attrs) => (attrs.floating ? {'data-floating': 'true'} : {}),
      },
      wrapType: {
        default: null,
        parseHTML: (el) => el.getAttribute('data-wrap') || null,
        renderHTML: (attrs) => {
          if (!attrs.wrapType) return {};
          const w = String(attrs.wrapType);
          const extraStyles: string[] = [];
          if (w === 'square' || w === 'tight') {
            if (attrs.align === 'right') {
              extraStyles.push('float: right', 'margin-left: 12px', 'margin-bottom: 8px');
            } else if (attrs.align === 'left') {
              extraStyles.push('float: left', 'margin-right: 12px', 'margin-bottom: 8px');
            }
          }
          return {
            'data-wrap': w,
            ...(extraStyles.length > 0 ? {style: extraStyles.join('; ')} : {}),
          };
        },
      },
      floatLeftPx: {
        default: null,
        parseHTML: (el) => {
          const raw = el.getAttribute('data-float-left');
          if (raw === null) return null;
          const n = Number(raw);
          return Number.isFinite(n) ? Math.round(n) : null;
        },
        renderHTML: (attrs) => {
          if (attrs.floatLeftPx === null || attrs.floatLeftPx === undefined) return {};
          const leftVal = Number(attrs.floatLeftPx);
          const res: Record<string, string> = {'data-float-left': String(leftVal)};
          if (attrs.align !== 'center' && attrs.wrapType !== 'square' && attrs.wrapType !== 'tight' && leftVal > 0) {
            res.style = `display: block; margin-left: ${leftVal}px`;
          }
          return res;
        },
      },
      floatTopPx: {
        default: null,
        parseHTML: (el) => {
          const raw = el.getAttribute('data-float-top');
          if (raw === null) return null;
          const n = Number(raw);
          return Number.isFinite(n) ? Math.round(n) : null;
        },
        renderHTML: (attrs) =>
          attrs.floatTopPx !== null && attrs.floatTopPx !== undefined
            ? {'data-float-top': String(attrs.floatTopPx)}
            : {},
      },
      relativeFromH: {
        default: null,
        parseHTML: (el) => el.getAttribute('data-rel-h') || null,
        renderHTML: (attrs) => (attrs.relativeFromH ? {'data-rel-h': String(attrs.relativeFromH)} : {}),
      },
      relativeFromV: {
        default: null,
        parseHTML: (el) => el.getAttribute('data-rel-v') || null,
        renderHTML: (attrs) => (attrs.relativeFromV ? {'data-rel-v': String(attrs.relativeFromV)} : {}),
      },
      marginTopPx: {
        default: null,
        parseHTML: (el) => {
          const raw = el.getAttribute('data-margin-top');
          if (raw === null) return null;
          const n = Number(raw);
          return Number.isFinite(n) ? Math.round(n) : null;
        },
        renderHTML: (attrs) =>
          attrs.marginTopPx !== null && attrs.marginTopPx !== undefined
            ? {'data-margin-top': String(attrs.marginTopPx), style: `margin-top: ${attrs.marginTopPx}px`}
            : {},
      },
      marginBottomPx: {
        default: null,
        parseHTML: (el) => {
          const raw = el.getAttribute('data-margin-bottom');
          if (raw === null) return null;
          const n = Number(raw);
          return Number.isFinite(n) ? Math.round(n) : null;
        },
        renderHTML: (attrs) =>
          attrs.marginBottomPx !== null && attrs.marginBottomPx !== undefined
            ? {
                'data-margin-bottom': String(attrs.marginBottomPx),
                style: `margin-bottom: ${attrs.marginBottomPx}px`,
              }
            : {},
      },
      ooxmlDrawing: {
        default: null,
        parseHTML: (el) => el.getAttribute('data-ooxml-drawing') || null,
        renderHTML: (attrs) =>
          attrs.ooxmlDrawing ? {'data-ooxml-drawing': String(attrs.ooxmlDrawing)} : {},
      },
    };
  },
});

const docTabLayoutKey = new PluginKey('docTabLayout');

function measureInlineWidthUntilNextTab(tabEl: HTMLElement, liveScale: number): number {
  const parent = tabEl.closest('p, h1, h2, h3, h4, li');
  if (!parent) return 0;

  // Walk sibling/descendant text and inline nodes after tabEl up to the next tab or <br>
  const doc = tabEl.ownerDocument;
  const range = doc.createRange();
  let endNode: globalThis.Node | null = null;

  const walker = doc.createTreeWalker(parent, NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT);
  let passedTab = false;
  let cur: globalThis.Node | null;
  while ((cur = walker.nextNode()) !== null) {
    if (cur === tabEl) {
      passedTab = true;
      continue;
    }
    if (!passedTab) continue;
    if (cur.nodeType === globalThis.Node.ELEMENT_NODE) {
      const el = cur as HTMLElement;
      if (el.hasAttribute('data-doc-tab') || el.tagName === 'BR' || el.classList.contains('esz-page-spacer-inline')) {
        break;
      }
    } else if (cur.nodeType === globalThis.Node.TEXT_NODE && (cur.nodeValue ?? '').length > 0) {
      if (!cur.parentElement?.closest('[data-doc-tab="true"]')) {
        endNode = cur;
      }
    }
  }

  if (!endNode || !tabEl.parentNode) return 0;
  try {
    range.setStartAfter(tabEl);
    range.setEnd(endNode, (endNode.nodeValue ?? '').length);
    const rects = range.getClientRects();
    let totalW = 0;
    for (let i = 0; i < rects.length; i++) {
      totalW += rects[i].width;
    }
    return totalW / liveScale;
  } catch {
    return 0;
  }
}

function layoutTabsInEditorDom(rootDom: HTMLElement): void {
  const tabEls = Array.from(rootDom.querySelectorAll<HTMLElement>('span[data-doc-tab="true"]'));
  if (tabEls.length === 0) return;

  const rootRect = rootDom.getBoundingClientRect();
  const liveScale =
    rootDom.offsetWidth > 0 && rootRect.width > 0 ? rootRect.width / rootDom.offsetWidth : 1;
  const safeScale = liveScale > 0.05 ? liveScale : 1;

  const defaultStopAttr = rootDom.closest<HTMLElement>('[data-default-tab-stop]')?.getAttribute('data-default-tab-stop');
  const defaultTabStopPx = Math.max(24, Number(defaultStopAttr) || 48);

  const paras = new Set<HTMLElement>();
  for (const t of tabEls) {
    const p = t.closest<HTMLElement>('p, h1, h2, h3, h4, li');
    if (p) paras.add(p);
  }

  for (const para of paras) {
    const paraTabs = Array.from(para.querySelectorAll<HTMLElement>('span[data-doc-tab="true"]'));
    if (paraTabs.length === 0) continue;

    let stops: DocTabStop[] = [];
    const rawStops = para.getAttribute('data-tab-stops');
    if (rawStops) {
      try {
        const parsed = JSON.parse(rawStops);
        if (Array.isArray(parsed)) {
          stops = parsed
            .filter((s) => s && typeof s.posPx === 'number' && s.val !== 'clear')
            .sort((a, b) => a.posPx - b.posPx);
        }
      } catch {
        stops = [];
      }
    }

    const paraRect = para.getBoundingClientRect();
    const paraStyle = getComputedStyle(para);
    const padLeft = Number.parseFloat(paraStyle.paddingLeft) || 0;
    const paraContentLeft = paraRect.left + padLeft * safeScale;
    const paraContentWidth = Math.max(120, paraRect.width / safeScale - padLeft);

    // Measure left-to-right sequentially so each tab's width advances subsequent tabs accurately
    for (const tabEl of paraTabs) {
      const tabRect = tabEl.getBoundingClientRect();
      const currentX = Math.max(0, (tabRect.left - paraContentLeft) / safeScale);

      // Find first explicit stop strictly to the right of currentX
      const explicitStop = stops.find((s) => s.posPx > currentX + 2);
      let targetStopPx: number;
      let align: string = tabEl.getAttribute('data-tab-align') || 'left';
      let leader: string = tabEl.getAttribute('data-tab-leader') || 'none';

      if (explicitStop) {
        targetStopPx = Math.min(paraContentWidth, explicitStop.posPx);
        align = explicitStop.val;
        leader = explicitStop.leader;
      } else {
        targetStopPx = Math.min(
          paraContentWidth,
          Math.ceil((currentX + 2) / defaultTabStopPx) * defaultTabStopPx,
        );
      }

      let widthPx = targetStopPx - currentX;
      if (align === 'right' || align === 'center' || align === 'decimal') {
        const afterW = measureInlineWidthUntilNextTab(tabEl, safeScale);
        if (align === 'right' || align === 'decimal') {
          widthPx = targetStopPx - currentX - afterW;
        } else {
          widthPx = targetStopPx - currentX - afterW / 2;
        }
      }

      const clampedW = Math.max(12, Math.min(paraContentWidth, Math.round(widthPx * 10) / 10));
      tabEl.style.width = `${clampedW}px`;
      tabEl.setAttribute('data-tab-align', align);
      tabEl.setAttribute('data-tab-leader', leader);
      tabEl.setAttribute('data-tab-computed-width', String(clampedW));
    }
  }
}

/**
 * Real OOXML inline tab node (<w:tab/>) with live tab-stop alignment (left/center/right/decimal)
 * and tab leader rendering (dot, hyphen, underscore) synced to paragraph w:tabs stops.
 */
export const DocTabNode = Node.create({
  name: 'docTab',
  group: 'inline',
  inline: true,
  atom: true,
  selectable: true,

  addAttributes() {
    return {
      posTwips: {
        default: null,
        parseHTML: (el) => {
          const raw = el.getAttribute('data-tab-pos');
          if (raw === null) return null;
          const n = Number(raw);
          return Number.isFinite(n) ? n : null;
        },
      },
      align: {
        default: 'left',
        parseHTML: (el) => el.getAttribute('data-tab-align') || 'left',
      },
      leader: {
        default: 'none',
        parseHTML: (el) => el.getAttribute('data-tab-leader') || 'none',
      },
      widthPx: {
        default: 48,
        parseHTML: (el) => {
          const raw = el.getAttribute('data-tab-width');
          const n = raw ? Number.parseFloat(raw) : NaN;
          return Number.isFinite(n) && n > 0 ? Math.round(n) : 48;
        },
      },
    };
  },

  parseHTML() {
    return [{tag: 'span[data-doc-tab="true"]'}];
  },

  renderHTML({node, HTMLAttributes}) {
    const widthPx = Math.max(12, Number(node.attrs.widthPx) || 48);
    const align = String(node.attrs.align || 'left');
    const leader = String(node.attrs.leader || 'none');
    const attrs: Record<string, string> = {
      'data-doc-tab': 'true',
      'data-tab-align': align,
      'data-tab-leader': leader,
      'data-tab-width': String(widthPx),
      class: 'esz-doc-tab',
      style: `width: ${widthPx}px`,
    };
    if (node.attrs.posTwips !== null && node.attrs.posTwips !== undefined) {
      attrs['data-tab-pos'] = String(node.attrs.posTwips);
    }
    return ['span', mergeAttributes(HTMLAttributes, attrs), '\t'];
  },

  addCommands() {
    return {
      insertDocTab:
        () =>
        ({commands}) =>
          commands.insertContent({type: this.name}),
    };
  },

  addKeyboardShortcuts() {
    return {
      Tab: () => {
        if (this.editor.isActive('listItem') || this.editor.isActive('table')) {
          return false;
        }
        return this.editor.commands.insertDocTab();
      },
    };
  },

  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: docTabLayoutKey,
        view(editorView) {
          let rafId: number | null = null;
          const schedule = () => {
            if (rafId !== null) window.cancelAnimationFrame(rafId);
            rafId = window.requestAnimationFrame(() => {
              rafId = null;
              if (!editorView.isDestroyed) {
                layoutTabsInEditorDom(editorView.dom);
              }
            });
          };
          schedule();
          if (typeof document !== 'undefined' && document.fonts?.ready) {
            void document.fonts.ready.then(schedule);
          }
          return {
            update() {
              schedule();
            },
            destroy() {
              if (rafId !== null) window.cancelAnimationFrame(rafId);
            },
          };
        },
      }),
    ];
  },
});

export const IndentExtension = Extension.create({
  name: 'indent',

  addGlobalAttributes() {
    return [
      {
        types: ['paragraph', 'heading'],
        attributes: {
          indent: {
            default: 0,
            parseHTML: (element) => {
              const raw = element.getAttribute('data-indent');
              const parsed = raw ? Number.parseInt(raw, 10) : 0;
              return Number.isFinite(parsed) ? Math.min(6, Math.max(0, parsed)) : 0;
            },
            renderHTML: (attributes) => {
              const lvl = Number(attributes.indent) || 0;
              if (lvl <= 0) return {};
              return {
                'data-indent': String(lvl),
                style: `margin-left: ${lvl * 32}px`,
              };
            },
          },
        },
      },
    ];
  },

  addCommands() {
    return {
      increaseIndent:
        () =>
        ({tr, state, dispatch}) => {
          const {from, to} = state.selection;
          let changed = false;
          state.doc.nodesBetween(from, to, (node, pos) => {
            if (node.type.name === 'paragraph' || node.type.name === 'heading') {
              const curr = Number(node.attrs.indent) || 0;
              const next = Math.min(6, curr + 1);
              if (next !== curr) {
                tr.setNodeMarkup(pos, undefined, {...node.attrs, indent: next});
                changed = true;
              }
            }
          });
          if (changed && dispatch) dispatch(tr);
          return changed;
        },
      decreaseIndent:
        () =>
        ({tr, state, dispatch}) => {
          const {from, to} = state.selection;
          let changed = false;
          state.doc.nodesBetween(from, to, (node, pos) => {
            if (node.type.name === 'paragraph' || node.type.name === 'heading') {
              const curr = Number(node.attrs.indent) || 0;
              const next = Math.max(0, curr - 1);
              if (next !== curr) {
                tr.setNodeMarkup(pos, undefined, {...node.attrs, indent: next});
                changed = true;
              }
            }
          });
          if (changed && dispatch) dispatch(tr);
          return changed;
        },
    };
  },
});

export const PageBreakNode = Node.create({
  name: 'pageBreak',
  group: 'block',
  atom: true,
  selectable: true,

  parseHTML() {
    return [{tag: 'div[data-page-break="true"]'}];
  },

  renderHTML({HTMLAttributes}) {
    return [
      'div',
      mergeAttributes(HTMLAttributes, {'data-page-break': 'true', contenteditable: 'false'}),
      'Page Break',
    ];
  },

  addCommands() {
    return {
      insertPageBreak:
        () =>
        ({commands}) =>
          commands.insertContent([{type: this.name}, {type: 'paragraph'}]),
    };
  },
});

/**
 * Editable styled text box / rounded rectangle block node.
 * Preserves Word DrawingML dimensions, border radius, border, fill, vertical alignment,
 * padding, and horizontal placement while allowing full rich-text editing inside.
 */
export const DocTextBoxNode = Node.create({
  name: 'docTextBox',
  group: 'block',
  content: 'block+',
  defining: true,

  addAttributes() {
    return {
      widthPx: {
        default: null,
        parseHTML: (el) => {
          const n = Number(el.getAttribute('data-box-width'));
          return Number.isFinite(n) && n > 0 ? n : null;
        },
      },
      minHeightPx: {
        default: null,
        parseHTML: (el) => {
          const n = Number(el.getAttribute('data-box-min-height'));
          return Number.isFinite(n) && n > 0 ? n : null;
        },
      },
      align: {
        default: null,
        parseHTML: (el) => el.getAttribute('data-box-align') || null,
      },
      borderRadiusPx: {
        default: null,
        parseHTML: (el) => {
          const raw = el.getAttribute('data-box-radius');
          if (raw === null) return null;
          const n = Number(raw);
          return Number.isFinite(n) && n >= 0 ? n : null;
        },
      },
      fillColor: {
        default: null,
        parseHTML: (el) => el.getAttribute('data-box-fill') || null,
      },
      borderColor: {
        default: null,
        parseHTML: (el) => el.getAttribute('data-box-border-color') || null,
      },
      borderWidthPx: {
        default: null,
        parseHTML: (el) => {
          const raw = el.getAttribute('data-box-border-width');
          if (raw === null) return null;
          const n = Number(raw);
          return Number.isFinite(n) && n >= 0 ? n : null;
        },
      },
      borderStyle: {
        default: null,
        parseHTML: (el) => el.getAttribute('data-box-border-style') || null,
      },
      vAlign: {
        default: null,
        parseHTML: (el) => el.getAttribute('data-box-valign') || null,
      },
      paddingCss: {
        default: null,
        parseHTML: (el) => el.getAttribute('data-box-padding') || null,
      },
      isWordArt: {
        default: false,
        parseHTML: (el) => el.getAttribute('data-box-wordart') === 'true',
      },
    };
  },

  parseHTML() {
    return [{tag: 'div[data-doc-textbox="true"]'}];
  },

  renderHTML({node, HTMLAttributes}) {
    const a = node.attrs;
    const isCustom =
      a.widthPx !== null ||
      a.borderRadiusPx !== null ||
      a.borderColor !== null ||
      a.fillColor !== null;

    const styles: string[] = [];
    const dataAttrs: Record<string, string> = {'data-doc-textbox': 'true'};

    if (isCustom) {
      dataAttrs['data-box-custom'] = 'true';
      styles.push('box-sizing: border-box', 'max-width: 100%', 'display: flex', 'flex-direction: column');
      if (a.widthPx) {
        dataAttrs['data-box-width'] = String(a.widthPx);
        styles.push(`width: ${a.widthPx}px`);
      }
      if (a.minHeightPx) {
        dataAttrs['data-box-min-height'] = String(a.minHeightPx);
        styles.push(`min-height: ${a.minHeightPx}px`);
      }
      if (a.align) {
        dataAttrs['data-box-align'] = String(a.align);
        if (a.align === 'center') styles.push('margin-left: auto', 'margin-right: auto');
        else if (a.align === 'right') styles.push('margin-left: auto', 'margin-right: 0');
        else styles.push('margin-left: 0', 'margin-right: auto');
      }
      if (a.borderRadiusPx !== null) {
        dataAttrs['data-box-radius'] = String(a.borderRadiusPx);
        styles.push(`border-radius: ${a.borderRadiusPx}px`);
      }
      if (a.fillColor) {
        dataAttrs['data-box-fill'] = String(a.fillColor);
        styles.push(`background-color: ${a.fillColor}`);
      }
      if (a.borderStyle) dataAttrs['data-box-border-style'] = String(a.borderStyle);
      if (a.borderWidthPx !== null) dataAttrs['data-box-border-width'] = String(a.borderWidthPx);
      if (a.borderColor) dataAttrs['data-box-border-color'] = String(a.borderColor);

      if (a.borderColor && a.borderStyle !== 'none' && (a.borderWidthPx ?? 1) > 0) {
        styles.push(`border: ${a.borderWidthPx ?? 1}px ${a.borderStyle ?? 'solid'} ${a.borderColor}`);
      } else {
        styles.push('border: none');
      }

      if (a.vAlign) {
        dataAttrs['data-box-valign'] = String(a.vAlign);
        styles.push(
          `justify-content: ${a.vAlign === 'middle' ? 'center' : a.vAlign === 'bottom' ? 'flex-end' : 'flex-start'}`,
        );
      }
      if (a.paddingCss) {
        dataAttrs['data-box-padding'] = String(a.paddingCss);
        styles.push(`padding: ${a.paddingCss}`);
      }
      if (a.isWordArt) dataAttrs['data-box-wordart'] = 'true';
    }

    return [
      'div',
      mergeAttributes(HTMLAttributes, dataAttrs, styles.length > 0 ? {style: styles.join('; ')} : {}),
      0,
    ];
  },

  addCommands() {
    return {
      insertDocTextBox:
        (content = 'Type callout or sidebar text here...') =>
        ({commands}) =>
          commands.insertContent({
            type: this.name,
            content: [
              {
                type: 'paragraph',
                content: [{type: 'text', text: content}],
              },
            ],
          }),
    };
  },
});

/**
 * Non-editable locked visual block for unsupported advanced Word elements
 * (SmartArt, Charts, 3D WordArt, OLE objects, legacy VML drawings).
 */
export const DocLockedElementNode = Node.create({
  name: 'docLockedElement',
  group: 'block',
  atom: true,
  selectable: true,

  addAttributes() {
    return {
      label: {default: 'Advanced Word Element (Locked)'},
      detail: {default: 'Shown as a locked visual element in Edit mode.'},
      imageSrc: {default: null},
      widthPx: {default: null},
      heightPx: {default: null},
    };
  },

  parseHTML() {
    return [
      {
        tag: 'div[data-doc-locked="true"]',
        getAttrs: (el) => {
          const elem = el as HTMLElement;
          const img = elem.querySelector('img');
          const w = Number(elem.getAttribute('data-width'));
          const h = Number(elem.getAttribute('data-height'));
          return {
            label: elem.getAttribute('data-locked-label') || elem.textContent || 'Locked Word Element',
            detail: elem.getAttribute('data-locked-detail') || '',
            imageSrc: elem.getAttribute('data-locked-image') || img?.getAttribute('src') || null,
            widthPx: Number.isFinite(w) && w > 0 ? w : null,
            heightPx: Number.isFinite(h) && h > 0 ? h : null,
          };
        },
      },
    ];
  },

  renderHTML({node, HTMLAttributes}) {
    const label = String(node.attrs.label || 'Locked Word Element');
    const detail = String(node.attrs.detail || '');
    const imageSrc = node.attrs.imageSrc as string | null;
    const widthPx = node.attrs.widthPx as number | null;
    const heightPx = node.attrs.heightPx as number | null;

    const attrs: Record<string, string> = {
      'data-doc-locked': 'true',
      'data-locked-label': label,
      'data-locked-detail': detail,
      contenteditable: 'false',
    };
    if (imageSrc) attrs['data-locked-image'] = imageSrc;
    if (widthPx) attrs['data-width'] = String(widthPx);
    if (heightPx) attrs['data-height'] = String(heightPx);

    if (imageSrc) {
      return [
        'div',
        mergeAttributes(HTMLAttributes, attrs),
        [
          'img',
          {
            src: imageSrc,
            alt: label,
            style: `display:block;margin:0 auto;max-width:100%;${widthPx ? `width:${widthPx}px;` : ''}`,
          },
        ],
        ['div', {class: 'esz-locked-badge'}, `${label} · Locked`],
      ];
    }

    return [
      'div',
      mergeAttributes(HTMLAttributes, attrs),
      ['div', {class: 'esz-locked-title'}, label],
      ...(detail ? [['div', {class: 'esz-locked-detail'}, detail]] : []),
    ];
  },
});

export const DocEquationNode = Node.create({
  name: 'docEquation',
  group: 'block',
  atom: true,
  selectable: true,

  addAttributes() {
    return {
      formula: {
        default: 'E = mc²',
      },
    };
  },

  parseHTML() {
    return [
      {
        tag: 'div[data-doc-equation="true"]',
        getAttrs: (el) => ({
          formula: (el as HTMLElement).textContent || 'E = mc²',
        }),
      },
    ];
  },

  renderHTML({node, HTMLAttributes}) {
    return [
      'div',
      mergeAttributes(HTMLAttributes, {'data-doc-equation': 'true'}),
      String(node.attrs.formula || 'E = mc²'),
    ];
  },

  addCommands() {
    return {
      insertDocEquation:
        (formula: string) =>
        ({commands}) =>
          commands.insertContent([
            {type: this.name, attrs: {formula}},
            {type: 'paragraph'},
          ]),
    };
  },
});

export const DocShapeNode = Node.create({
  name: 'docShape',
  group: 'block',
  atom: true,
  selectable: true,

  addAttributes() {
    return {
      shapeType: {default: 'rounded'},
      label: {default: 'Shape'},
      fill: {default: '#DBEAFE'},
      border: {default: '#1D4ED8'},
    };
  },

  parseHTML() {
    return [
      {
        tag: 'div[data-doc-shape="true"]',
        getAttrs: (el) => {
          const elem = el as HTMLElement;
          return {
            shapeType: elem.getAttribute('data-shape-type') || 'rounded',
            label: elem.textContent || 'Shape',
            fill: elem.getAttribute('data-fill') || '#DBEAFE',
            border: elem.getAttribute('data-border') || '#1D4ED8',
          };
        },
      },
    ];
  },

  renderHTML({node, HTMLAttributes}) {
    const shape = String(node.attrs.shapeType || 'rounded');
    const fill = String(node.attrs.fill || '#DBEAFE');
    const border = String(node.attrs.border || '#1D4ED8');
    const radius =
      shape === 'circle'
        ? '9999px'
        : shape === 'rounded'
          ? '12px'
          : shape === 'callout'
            ? '12px 12px 12px 2px'
            : '2px';
    const width = shape === 'circle' ? '140px' : '260px';
    const height = shape === 'circle' ? '140px' : '74px';

    return [
      'div',
      mergeAttributes(HTMLAttributes, {
        'data-doc-shape': 'true',
        'data-shape-type': shape,
        'data-fill': fill,
        'data-border': border,
        style: `width:${width};height:${height};border-radius:${radius};background:${fill};border:2px solid ${border};color:#0f172a;padding:8px 14px;text-align:center;`,
      }),
      String(node.attrs.label || 'Shape'),
    ];
  },

  addCommands() {
    return {
      insertDocShape:
        (attrs) =>
        ({commands}) =>
          commands.insertContent([
            {type: this.name, attrs},
            {type: 'paragraph'},
          ]),
    };
  },
});

export const DocCaptionNode = Node.create({
  name: 'docCaption',
  group: 'block',
  content: 'inline*',

  parseHTML() {
    return [{tag: 'p[data-doc-caption="true"]'}];
  },

  renderHTML({HTMLAttributes}) {
    return ['p', mergeAttributes(HTMLAttributes, {'data-doc-caption': 'true'}), 0];
  },

  addCommands() {
    return {
      insertDocCaption:
        (label: string) =>
        ({commands}) =>
          commands.insertContent({
            type: this.name,
            content: [{type: 'text', text: label}],
          }),
    };
  },
});

export const DocTocNode = Node.create({
  name: 'docToc',
  group: 'block',
  content: 'block+',
  defining: true,

  parseHTML() {
    return [{tag: 'div[data-doc-toc="true"]'}];
  },

  renderHTML({HTMLAttributes}) {
    return ['div', mergeAttributes(HTMLAttributes, {'data-doc-toc': 'true'}), 0];
  },
});

export const BookmarkMark = Mark.create({
  name: 'bookmark',

  addAttributes() {
    return {
      name: {
        default: 'bookmark',
        parseHTML: (el) => el.getAttribute('data-bookmark') || 'bookmark',
        renderHTML: (attrs) => ({
          'data-bookmark': attrs.name,
          id: `bm-${attrs.name}`,
          title: `Bookmark: ${attrs.name}`,
        }),
      },
    };
  },

  parseHTML() {
    return [{tag: 'span[data-bookmark]'}];
  },

  renderHTML({HTMLAttributes}) {
    return ['span', mergeAttributes(HTMLAttributes), 0];
  },

  addCommands() {
    return {
      setBookmark:
        (name: string) =>
        ({commands}) =>
          commands.setMark(this.name, {name}),
    };
  },
});

export const NoteRefNode = Node.create({
  name: 'noteRef',
  group: 'inline',
  inline: true,
  atom: true,

  addAttributes() {
    return {
      noteId: {default: ''},
      number: {default: 1},
      kind: {default: 'footnote'},
    };
  },

  parseHTML() {
    return [
      {
        tag: 'sup[data-note-ref]',
        getAttrs: (el) => {
          const elem = el as HTMLElement;
          return {
            noteId: elem.getAttribute('data-note-id') || '',
            number: Number(elem.getAttribute('data-note-num') || '1'),
            kind: elem.getAttribute('data-note-kind') || 'footnote',
          };
        },
      },
    ];
  },

  renderHTML({node, HTMLAttributes}) {
    const num = String(node.attrs.number ?? 1);
    const kind = String(node.attrs.kind || 'footnote');
    return [
      'sup',
      mergeAttributes(HTMLAttributes, {
        'data-note-ref': 'true',
        'data-note-id': String(node.attrs.noteId || ''),
        'data-note-num': num,
        'data-note-kind': kind,
        title: `${kind === 'footnote' ? 'Footnote' : 'Endnote'} ${num}`,
      }),
      `[${num}]`,
    ];
  },
});
