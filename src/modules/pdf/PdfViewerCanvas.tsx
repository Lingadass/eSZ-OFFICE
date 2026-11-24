import React, {useCallback, useEffect, useMemo, useRef, useState} from 'react';
import type {PDFDocumentProxy, RenderTask} from 'pdfjs-dist';
import {StickyNote, Trash2} from 'lucide-react';
import type {
  PdfAnnotationItem,
  PdfAnnotationTool,
  PdfHeaderFooterConfig,
  PdfLayoutMode,
  PdfPageNumberConfig,
  PdfPoint,
  PdfWatermarkConfig,
  PdfZoomMode,
} from './pdfTypes';
import {computeOverlayLayoutForPage} from './pdfOverlayLayout';

interface PdfViewerCanvasProps {
  pdfDoc: PDFDocumentProxy;
  numPages: number;
  currentPage: number;
  scrollSignal: number;
  onPageChange: (page: number) => void;
  layoutMode: PdfLayoutMode;
  zoomMode: PdfZoomMode;
  zoom: number;
  viewRotation: number;
  darkReading: boolean;
  activeTool: PdfAnnotationTool;
  toolColor: string;
  toolOpacity: number;
  toolStrokeWidth: number;
  pendingStampText: string;
  pendingImageDataUrl: string | null;
  onConsumePendingImage: () => void;
  annotations: PdfAnnotationItem[];
  selectedAnnotationId: string | null;
  onSelectAnnotation: (id: string | null) => void;
  onAddAnnotation: (ann: PdfAnnotationItem) => void;
  onDeleteAnnotation: (id: string) => void;
  watermark: PdfWatermarkConfig;
  headerFooter: PdfHeaderFooterConfig;
  pageNumbers: PdfPageNumberConfig;
  searchHighlightPages: Set<number>;
  activeSearchPage: number | null;
  readOnly?: boolean;
}

interface DraftBox {
  pageIndex: number;
  startX: number;
  startY: number;
  currX: number;
  currY: number;
  points: PdfPoint[];
}

interface TextPromptState {
  pageIndex: number;
  x: number;
  y: number;
  tool: 'text' | 'sticky';
  value: string;
}

function makeAnnId(): string {
  return `ann_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
}

interface PageSurfaceProps {
  pdfDoc: PDFDocumentProxy;
  pageNumber: number;
  totalPages: number;
  containerSize: {width: number; height: number};
  layoutMode: PdfLayoutMode;
  zoomMode: PdfZoomMode;
  zoom: number;
  viewRotation: number;
  darkReading: boolean;
  activeTool: PdfAnnotationTool;
  toolColor: string;
  toolOpacity: number;
  toolStrokeWidth: number;
  pendingStampText: string;
  pendingImageDataUrl: string | null;
  onConsumePendingImage: () => void;
  annotations: PdfAnnotationItem[];
  selectedAnnotationId: string | null;
  onSelectAnnotation: (id: string | null) => void;
  onAddAnnotation: (ann: PdfAnnotationItem) => void;
  onDeleteAnnotation: (id: string) => void;
  onOpenTextPrompt: (prompt: TextPromptState) => void;
  watermark: PdfWatermarkConfig;
  headerFooter: PdfHeaderFooterConfig;
  pageNumbers: PdfPageNumberConfig;
  isSearchMatch: boolean;
  isActiveSearchMatch: boolean;
  readOnly?: boolean;
  registerPageElement: (pageNum: number, el: HTMLDivElement | null) => void;
}

const PdfPageSurface: React.FC<PageSurfaceProps> = ({
  pdfDoc,
  pageNumber,
  totalPages,
  containerSize,
  layoutMode,
  zoomMode,
  zoom,
  viewRotation,
  darkReading,
  activeTool,
  toolColor,
  toolOpacity,
  toolStrokeWidth,
  pendingStampText,
  pendingImageDataUrl,
  onConsumePendingImage,
  annotations,
  selectedAnnotationId,
  onSelectAnnotation,
  onAddAnnotation,
  onDeleteAnnotation,
  onOpenTextPrompt,
  watermark,
  headerFooter,
  pageNumbers,
  isSearchMatch,
  isActiveSearchMatch,
  readOnly = false,
  registerPageElement,
}) => {
  const wrapperRef = useRef<HTMLDivElement | null>(null);
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const overlayRef = useRef<HTMLDivElement | null>(null);
  const renderTaskRef = useRef<RenderTask | null>(null);

  const [dims, setDims] = useState<{width: number; height: number}>({width: 595, height: 842});
  const [draft, setDraft] = useState<DraftBox | null>(null);

  const setRefs = useCallback(
    (el: HTMLDivElement | null) => {
      wrapperRef.current = el;
      registerPageElement(pageNumber, el);
    },
    [pageNumber, registerPageElement],
  );

  useEffect(() => {
    let cancelled = false;

    const renderPage = async () => {
      try {
        const page = await pdfDoc.getPage(pageNumber);
        if (cancelled) return;

        const baseRotation = (page.rotate + viewRotation) % 360;
        const unscaled = page.getViewport({scale: 1, rotation: baseRotation});

        const availWidth = Math.max(
          260,
          layoutMode === 'two-page'
            ? (containerSize.width - 64) / 2
            : containerSize.width - 48,
        );
        const availHeight = Math.max(320, containerSize.height - 48);

        let effectiveScale = zoom;
        if (zoomMode === 'fit-width') {
          effectiveScale = Math.min(3, Math.max(0.4, availWidth / unscaled.width));
        } else if (zoomMode === 'fit-page') {
          effectiveScale = Math.min(
            3,
            Math.max(0.4, Math.min(availWidth / unscaled.width, availHeight / unscaled.height)),
          );
        }

        const viewport = page.getViewport({scale: effectiveScale, rotation: baseRotation});
        setDims({width: Math.round(viewport.width), height: Math.round(viewport.height)});

        const canvas = canvasRef.current;
        if (!canvas || cancelled) return;
        const ctx = canvas.getContext('2d');
        if (!ctx) return;

        if (renderTaskRef.current) {
          try {
            renderTaskRef.current.cancel();
          } catch {
            // ignore cancellation errors
          }
          renderTaskRef.current = null;
        }

        const dpr = Math.min(window.devicePixelRatio || 1, 2);
        canvas.width = Math.max(1, Math.floor(viewport.width * dpr));
        canvas.height = Math.max(1, Math.floor(viewport.height * dpr));
        canvas.style.width = `${Math.round(viewport.width)}px`;
        canvas.style.height = `${Math.round(viewport.height)}px`;

        ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
        ctx.fillStyle = '#FFFFFF';
        ctx.fillRect(0, 0, viewport.width, viewport.height);

        const task = page.render({
          canvasContext: ctx,
          viewport,
          canvas,
        });
        renderTaskRef.current = task;
        await task.promise;
        renderTaskRef.current = null;
      } catch (err) {
        if (err instanceof Error && err.name === 'RenderingCancelledException') {
          return;
        }
      }
    };

    void renderPage();

    return () => {
      cancelled = true;
      if (renderTaskRef.current) {
        try {
          renderTaskRef.current.cancel();
        } catch {
          // ignore
        }
      }
    };
  }, [pdfDoc, pageNumber, containerSize.width, containerSize.height, layoutMode, zoomMode, zoom, viewRotation]);

  const getRelativeCoords = (clientX: number, clientY: number): PdfPoint => {
    const rect = overlayRef.current?.getBoundingClientRect();
    if (!rect || rect.width === 0 || rect.height === 0) return {x: 0.5, y: 0.5};
    const x = Math.min(1, Math.max(0, (clientX - rect.left) / rect.width));
    const y = Math.min(1, Math.max(0, (clientY - rect.top) / rect.height));
    return {x, y};
  };

  const handlePointerDown = (e: React.PointerEvent<HTMLDivElement>) => {
    if (readOnly) return;
    if (activeTool === 'select') {
      if (e.target === overlayRef.current) {
        onSelectAnnotation(null);
      }
      return;
    }

    e.preventDefault();
    const pt = getRelativeCoords(e.clientX, e.clientY);
    const pageIndex = pageNumber - 1;

    if (activeTool === 'text' || activeTool === 'sticky') {
      onOpenTextPrompt({
        pageIndex,
        x: pt.x,
        y: pt.y,
        tool: activeTool,
        value: '',
      });
      return;
    }

    if (activeTool === 'stamp') {
      onAddAnnotation({
        id: makeAnnId(),
        pageIndex,
        tool: 'stamp',
        x: Math.min(0.72, pt.x),
        y: Math.min(0.88, pt.y),
        width: 0.24,
        height: 0.06,
        color: toolColor,
        opacity: 0.9,
        strokeWidth: 2,
        text: pendingStampText || 'APPROVED',
      });
      return;
    }

    if ((activeTool === 'image' || activeTool === 'signature') && pendingImageDataUrl) {
      onAddAnnotation({
        id: makeAnnId(),
        pageIndex,
        tool: activeTool,
        x: Math.min(0.65, pt.x),
        y: Math.min(0.82, pt.y),
        width: activeTool === 'signature' ? 0.28 : 0.3,
        height: activeTool === 'signature' ? 0.1 : 0.18,
        color: toolColor,
        opacity: 1,
        strokeWidth: 1,
        imageDataUrl: pendingImageDataUrl,
      });
      onConsumePendingImage();
      return;
    }

    e.currentTarget.setPointerCapture(e.pointerId);
    setDraft({
      pageIndex,
      startX: pt.x,
      startY: pt.y,
      currX: pt.x,
      currY: pt.y,
      points: [pt],
    });
  };

  const handlePointerMove = (e: React.PointerEvent<HTMLDivElement>) => {
    if (readOnly || !draft) return;
    const pt = getRelativeCoords(e.clientX, e.clientY);
    setDraft((prev) =>
      prev
        ? {
            ...prev,
            currX: pt.x,
            currY: pt.y,
            points: activeTool === 'freehand' ? [...prev.points, pt] : prev.points,
          }
        : null,
    );
  };

  const handlePointerUp = () => {
    if (readOnly || !draft || activeTool === 'select') return;
    const x = Math.min(draft.startX, draft.currX);
    const y = Math.min(draft.startY, draft.currY);
    const width = Math.abs(draft.currX - draft.startX);
    const height = Math.abs(draft.currY - draft.startY);

    if (activeTool === 'freehand') {
      if (draft.points.length > 1) {
        const xs = draft.points.map((p) => p.x);
        const ys = draft.points.map((p) => p.y);
        onAddAnnotation({
          id: makeAnnId(),
          pageIndex: draft.pageIndex,
          tool: 'freehand',
          x: Math.min(...xs),
          y: Math.min(...ys),
          width: Math.max(0.02, Math.max(...xs) - Math.min(...xs)),
          height: Math.max(0.02, Math.max(...ys) - Math.min(...ys)),
          color: toolColor,
          opacity: toolOpacity,
          strokeWidth: toolStrokeWidth,
          points: draft.points,
        });
      }
      setDraft(null);
      return;
    }

    if (activeTool === 'arrow') {
      if (Math.hypot(draft.currX - draft.startX, draft.currY - draft.startY) > 0.015) {
        onAddAnnotation({
          id: makeAnnId(),
          pageIndex: draft.pageIndex,
          tool: 'arrow',
          x: draft.startX,
          y: draft.startY,
          width: draft.currX - draft.startX,
          height: draft.currY - draft.startY,
          color: toolColor,
          opacity: toolOpacity,
          strokeWidth: toolStrokeWidth,
        });
      }
      setDraft(null);
      return;
    }

    if (
      activeTool === 'highlight' ||
      activeTool === 'underline' ||
      activeTool === 'strikethrough' ||
      activeTool === 'rect' ||
      activeTool === 'ellipse' ||
      activeTool === 'redact'
    ) {
      if (width > 0.01 || height > 0.01) {
        onAddAnnotation({
          id: makeAnnId(),
          pageIndex: draft.pageIndex,
          tool: activeTool,
          x,
          y,
          width: Math.max(0.02, width),
          height: Math.max(0.018, height),
          color: activeTool === 'redact' ? '#000000' : toolColor,
          opacity: activeTool === 'highlight' ? Math.min(0.45, toolOpacity) : activeTool === 'redact' ? 1 : toolOpacity,
          strokeWidth: toolStrokeWidth,
        });
      }
    }

    setDraft(null);
  };

  const pageAnns = annotations.filter((a) => a.pageIndex === pageNumber - 1);

  const previewScale = dims.width / 595;
  const liveOverlayLayout = useMemo(() => {
    if (!watermark.enabled && !headerFooter.enabled && !pageNumbers.enabled) {
      return null;
    }
    const baseHeightPt = dims.width > 0 ? (dims.height / dims.width) * 595 : 842;
    return computeOverlayLayoutForPage({
      pageIndex: pageNumber - 1,
      totalPages,
      pageWidth: 595,
      pageHeight: baseHeightPt,
      watermark,
      headerFooter,
      pageNumbers,
    });
  }, [watermark, headerFooter, pageNumbers, pageNumber, totalPages, dims.width, dims.height]);

  return (
    <div
      ref={setRefs}
      data-page-number={pageNumber}
      data-testid={`pdf-page-wrapper-${pageNumber}`}
      className="mx-auto flex shrink-0 flex-col items-center gap-2"
    >
      <div
        className={`relative select-none rounded-md bg-white shadow-md ring-1 transition-shadow ${
          isActiveSearchMatch
            ? 'ring-2 ring-amber-500'
            : isSearchMatch
              ? 'ring-2 ring-blue-500/70'
              : 'ring-slate-300 dark:ring-slate-700'
        }`}
        style={{width: `${dims.width}px`, height: `${dims.height}px`}}
      >
        {/* Rendered PDF.js page canvas */}
        <canvas
          ref={canvasRef}
          className="block rounded-md"
          style={darkReading ? {filter: 'invert(0.9) hue-rotate(180deg)'} : undefined}
        />

        {/* Live Watermark, Header/Footer, and Page Number Overlay Preview */}
        {liveOverlayLayout?.watermarkItem && (
          <div
            aria-hidden="true"
            className="pointer-events-none absolute inset-0 overflow-hidden"
          >
            <span
              className="absolute whitespace-nowrap font-bold uppercase tracking-wider select-none"
              style={{
                left: `${liveOverlayLayout.watermarkItem.centerX * previewScale}px`,
                top: `${liveOverlayLayout.watermarkItem.previewTopY * previewScale}px`,
                color: liveOverlayLayout.watermarkItem.color,
                opacity: liveOverlayLayout.watermarkItem.opacity,
                fontSize: `${Math.max(10, Math.round(liveOverlayLayout.watermarkItem.fontSize * previewScale))}px`,
                transform: `translate(-50%, -50%) rotate(${-liveOverlayLayout.watermarkItem.rotationDeg}deg)`,
                lineHeight: 1,
              }}
            >
              {liveOverlayLayout.watermarkItem.text}
            </span>
          </div>
        )}

        {liveOverlayLayout &&
          [...liveOverlayLayout.headerItems, ...liveOverlayLayout.footerItems].map((item, idx) => (
            <div
              key={`draft_hf_${idx}`}
              aria-hidden="true"
              className="pointer-events-none absolute whitespace-nowrap text-slate-600 select-none"
              style={{
                left: `${item.x * previewScale}px`,
                top: `${item.topY * previewScale}px`,
                fontSize: `${Math.max(8, item.fontSize * previewScale)}px`,
                lineHeight: 1.1,
              }}
            >
              {item.text}
            </div>
          ))}

        {liveOverlayLayout?.pageNumberItem && (
          <div
            aria-hidden="true"
            className="pointer-events-none absolute whitespace-nowrap font-medium text-slate-700 select-none"
            style={{
              left: `${liveOverlayLayout.pageNumberItem.x * previewScale}px`,
              top: `${liveOverlayLayout.pageNumberItem.topY * previewScale}px`,
              fontSize: `${Math.max(8, liveOverlayLayout.pageNumberItem.fontSize * previewScale)}px`,
              lineHeight: 1.1,
            }}
          >
            {liveOverlayLayout.pageNumberItem.text}
          </div>
        )}

      {/* Interactive Annotation Surface */}
      <div
        ref={overlayRef}
        onPointerDown={handlePointerDown}
        onPointerMove={handlePointerMove}
        onPointerUp={handlePointerUp}
        onPointerCancel={() => setDraft(null)}
        className={`absolute inset-0 ${
          readOnly || activeTool === 'select'
            ? 'cursor-default touch-auto'
            : 'cursor-crosshair touch-none'
        }`}
      >
        <svg
          className="pointer-events-none absolute inset-0 h-full w-full"
          viewBox={`0 0 ${dims.width} ${dims.height}`}
        >
          {pageAnns.map((ann) => {
            const px = ann.x * dims.width;
            const py = ann.y * dims.height;
            const pw = ann.width * dims.width;
            const ph = ann.height * dims.height;

            if (ann.tool === 'highlight') {
              return (
                <rect
                  key={ann.id}
                  x={px}
                  y={py}
                  width={pw}
                  height={ph}
                  fill={ann.color}
                  fillOpacity={ann.opacity ?? 0.35}
                  rx={2}
                />
              );
            }
            if (ann.tool === 'underline') {
              return (
                <line
                  key={ann.id}
                  x1={px}
                  y1={py + ph}
                  x2={px + pw}
                  y2={py + ph}
                  stroke={ann.color}
                  strokeWidth={Math.max(2, ann.strokeWidth)}
                />
              );
            }
            if (ann.tool === 'strikethrough') {
              return (
                <line
                  key={ann.id}
                  x1={px}
                  y1={py + ph / 2}
                  x2={px + pw}
                  y2={py + ph / 2}
                  stroke={ann.color}
                  strokeWidth={Math.max(2, ann.strokeWidth)}
                />
              );
            }
            if (ann.tool === 'rect') {
              return (
                <rect
                  key={ann.id}
                  x={px}
                  y={py}
                  width={pw}
                  height={ph}
                  fill="none"
                  stroke={ann.color}
                  strokeWidth={Math.max(2, ann.strokeWidth)}
                />
              );
            }
            if (ann.tool === 'ellipse') {
              return (
                <ellipse
                  key={ann.id}
                  cx={px + pw / 2}
                  cy={py + ph / 2}
                  rx={Math.max(2, pw / 2)}
                  ry={Math.max(2, ph / 2)}
                  fill="none"
                  stroke={ann.color}
                  strokeWidth={Math.max(2, ann.strokeWidth)}
                />
              );
            }
            if (ann.tool === 'arrow') {
              const x1 = ann.x * dims.width;
              const y1 = ann.y * dims.height;
              const x2 = (ann.x + ann.width) * dims.width;
              const y2 = (ann.y + ann.height) * dims.height;
              const angle = Math.atan2(y2 - y1, x2 - x1);
              const head = 10;
              return (
                <g key={ann.id}>
                  <line
                    x1={x1}
                    y1={y1}
                    x2={x2}
                    y2={y2}
                    stroke={ann.color}
                    strokeWidth={Math.max(2, ann.strokeWidth)}
                  />
                  <polygon
                    points={`${x2},${y2} ${x2 - head * Math.cos(angle - Math.PI / 6)},${y2 - head * Math.sin(angle - Math.PI / 6)} ${x2 - head * Math.cos(angle + Math.PI / 6)},${y2 - head * Math.sin(angle + Math.PI / 6)}`}
                    fill={ann.color}
                  />
                </g>
              );
            }
            if (ann.tool === 'freehand' && ann.points && ann.points.length > 1) {
              const d = ann.points
                .map((p, idx) => `${idx === 0 ? 'M' : 'L'} ${p.x * dims.width} ${p.y * dims.height}`)
                .join(' ');
              return (
                <path
                  key={ann.id}
                  d={d}
                  fill="none"
                  stroke={ann.color}
                  strokeWidth={Math.max(2, ann.strokeWidth)}
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  strokeOpacity={ann.opacity}
                />
              );
            }
            if (ann.tool === 'redact') {
              return (
                <rect
                  key={ann.id}
                  x={px}
                  y={py}
                  width={pw}
                  height={ph}
                  fill="#000000"
                  fillOpacity={0.92}
                  stroke="#EF4444"
                  strokeWidth={1}
                  strokeDasharray="4 2"
                />
              );
            }
            return null;
          })}

          {/* Active Draft Shape Preview */}
          {draft && draft.pageIndex === pageNumber - 1 && (
            <>
              {activeTool === 'freehand' && draft.points.length > 1 && (
                <path
                  d={draft.points
                    .map((p, idx) => `${idx === 0 ? 'M' : 'L'} ${p.x * dims.width} ${p.y * dims.height}`)
                    .join(' ')}
                  fill="none"
                  stroke={toolColor}
                  strokeWidth={Math.max(2, toolStrokeWidth)}
                  strokeLinecap="round"
                />
              )}
              {activeTool === 'arrow' && (
                <line
                  x1={draft.startX * dims.width}
                  y1={draft.startY * dims.height}
                  x2={draft.currX * dims.width}
                  y2={draft.currY * dims.height}
                  stroke={toolColor}
                  strokeWidth={Math.max(2, toolStrokeWidth)}
                />
              )}
              {(activeTool === 'highlight' ||
                activeTool === 'underline' ||
                activeTool === 'strikethrough' ||
                activeTool === 'rect' ||
                activeTool === 'ellipse' ||
                activeTool === 'redact') && (
                <rect
                  x={Math.min(draft.startX, draft.currX) * dims.width}
                  y={Math.min(draft.startY, draft.currY) * dims.height}
                  width={Math.abs(draft.currX - draft.startX) * dims.width}
                  height={Math.abs(draft.currY - draft.startY) * dims.height}
                  fill={activeTool === 'redact' ? '#000000' : toolColor}
                  fillOpacity={activeTool === 'redact' ? 0.75 : 0.25}
                  stroke={activeTool === 'redact' ? '#EF4444' : toolColor}
                  strokeWidth={1.5}
                  strokeDasharray="4 2"
                />
              )}
            </>
          )}
        </svg>

        {/* HTML Annotation Items (Text, Sticky, Stamp, Image, Signature, and Selection handles) */}
        {pageAnns.map((ann) => {
          const isSelected = ann.id === selectedAnnotationId;
          const left = `${ann.x * 100}%`;
          const top = `${ann.y * 100}%`;
          const width = `${Math.max(4, ann.width * 100)}%`;
          const height = `${Math.max(3, ann.height * 100)}%`;

          return (
            <div
              key={ann.id}
              onClick={(e) => {
                if (readOnly) return;
                e.stopPropagation();
                onSelectAnnotation(ann.id);
              }}
              style={{
                left: ann.tool === 'arrow' && ann.width < 0 ? `${(ann.x + ann.width) * 100}%` : left,
                top: ann.tool === 'arrow' && ann.height < 0 ? `${(ann.y + ann.height) * 100}%` : top,
                width: ann.tool === 'arrow' ? `${Math.max(4, Math.abs(ann.width) * 100)}%` : width,
                height: ann.tool === 'arrow' ? `${Math.max(3, Math.abs(ann.height) * 100)}%` : height,
              }}
              className={`absolute ${
                !readOnly && activeTool === 'select'
                  ? 'pointer-events-auto cursor-pointer'
                  : 'pointer-events-none'
              } ${isSelected && !readOnly ? 'ring-2 ring-blue-600 ring-offset-1' : ''}`}
            >
              {ann.tool === 'text' && (
                <div
                  className="px-1 text-xs font-medium leading-snug break-words"
                  style={{color: ann.color}}
                >
                  {ann.text}
                </div>
              )}

              {ann.tool === 'sticky' && (
                <div className="flex min-w-[140px] max-w-[200px] items-start gap-1.5 rounded-lg border border-amber-400 bg-amber-100/95 p-2 text-[11px] text-slate-900 shadow-sm">
                  <StickyNote className="mt-0.5 h-3.5 w-3.5 shrink-0 text-amber-700" />
                  <span className="break-words">{ann.text}</span>
                </div>
              )}

              {ann.tool === 'stamp' && (
                <div
                  className="flex h-full w-full items-center justify-center rounded border-2 px-2 text-xs font-bold uppercase tracking-wider"
                  style={{borderColor: ann.color, color: ann.color}}
                >
                  {ann.text}
                </div>
              )}

              {(ann.tool === 'image' || ann.tool === 'signature') && ann.imageDataUrl && (
                <img
                  src={ann.imageDataUrl}
                  alt={ann.tool === 'signature' ? 'PDF signature' : 'Embedded image'}
                  className="h-full w-full object-contain"
                />
              )}

              {isSelected && !readOnly && (
                <button
                  type="button"
                  onClick={(e) => {
                    e.stopPropagation();
                    onDeleteAnnotation(ann.id);
                  }}
                  aria-label="Delete selected annotation"
                  className="absolute -right-3 -top-3 flex h-7 w-7 items-center justify-center rounded-full bg-rose-600 text-white shadow hover:bg-rose-700"
                >
                  <Trash2 className="h-3.5 w-3.5" />
                </button>
              )}
            </div>
          );
        })}
      </div>
      </div>

      {/* Per-page number badge placed cleanly outside the PDF canvas sheet */}
      <div
        data-testid="pdf-page-badge"
        data-page-badge={pageNumber}
        className="inline-flex items-center rounded-md border border-slate-300/80 bg-white/95 px-2.5 py-0.5 text-[11px] font-semibold tabular-nums text-slate-600 shadow-2xs dark:border-slate-700 dark:bg-slate-900/90 dark:text-slate-300"
      >
        Page {pageNumber} of {totalPages}
      </div>
    </div>
  );
};

export const PdfViewerCanvas: React.FC<PdfViewerCanvasProps> = ({
  pdfDoc,
  numPages,
  currentPage,
  scrollSignal,
  onPageChange,
  layoutMode,
  zoomMode,
  zoom,
  viewRotation,
  darkReading,
  activeTool,
  toolColor,
  toolOpacity,
  toolStrokeWidth,
  pendingStampText,
  pendingImageDataUrl,
  onConsumePendingImage,
  annotations,
  selectedAnnotationId,
  onSelectAnnotation,
  onAddAnnotation,
  onDeleteAnnotation,
  watermark,
  headerFooter,
  pageNumbers,
  searchHighlightPages,
  activeSearchPage,
  readOnly = false,
}) => {
  const scrollContainerRef = useRef<HTMLDivElement | null>(null);
  const pageElementsRef = useRef<Map<number, HTMLDivElement>>(new Map());
  const [containerSize, setContainerSize] = useState({width: 820, height: 680});
  const [textPrompt, setTextPrompt] = useState<TextPromptState | null>(null);
  const suppressObserverRef = useRef(false);
  const currentPageRef = useRef(currentPage);
  currentPageRef.current = currentPage;
  const lastHandledScrollSignalRef = useRef(scrollSignal);

  useEffect(() => {
    if (readOnly) {
      setTextPrompt(null);
    }
  }, [readOnly]);

  const registerPageElement = useCallback((pageNum: number, el: HTMLDivElement | null) => {
    if (el) {
      pageElementsRef.current.set(pageNum, el);
    } else {
      pageElementsRef.current.delete(pageNum);
    }
  }, []);

  useEffect(() => {
    const el = scrollContainerRef.current;
    if (!el) return;
    const updateSize = () => {
      setContainerSize({width: el.clientWidth, height: el.clientHeight});
    };
    updateSize();
    const ro = new ResizeObserver(updateSize);
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  // Scroll to currentPage ONLY when explicitly requested via scrollSignal (thumbnail click, Next/Prev, search jump, page input)
  useEffect(() => {
    if (scrollSignal === lastHandledScrollSignalRef.current) return;
    lastHandledScrollSignalRef.current = scrollSignal;
    if (layoutMode !== 'continuous') return;
    const el = pageElementsRef.current.get(currentPage);
    if (!el) return;
    suppressObserverRef.current = true;
    el.scrollIntoView({behavior: 'smooth', block: 'start'});
    const timer = window.setTimeout(() => {
      suppressObserverRef.current = false;
    }, 520);
    return () => window.clearTimeout(timer);
  }, [scrollSignal, layoutMode, currentPage]);

  // Accurately track the currently visible page during natural continuous scrolling at any zoom level
  useEffect(() => {
    if (layoutMode !== 'continuous') return;
    const root = scrollContainerRef.current;
    if (!root) return;

    let rafId: number | null = null;

    const detectVisiblePage = () => {
      rafId = null;
      if (suppressObserverRef.current) return;
      const rootRect = root.getBoundingClientRect();
      if (rootRect.height <= 0) return;

      // Upper-middle focus line feels natural as a user scrolls down through pages
      const focusY = rootRect.top + rootRect.height * 0.35;
      let bestPage: number | null = null;
      let bestDistance = Number.POSITIVE_INFINITY;

      pageElementsRef.current.forEach((node, pageNum) => {
        const rect = node.getBoundingClientRect();
        if (rect.height <= 0) return;
        if (rect.top <= focusY && rect.bottom >= focusY) {
          bestPage = pageNum;
          bestDistance = 0;
          return;
        }
        const dist = focusY < rect.top ? rect.top - focusY : focusY - rect.bottom;
        if (dist < bestDistance) {
          bestDistance = dist;
          bestPage = pageNum;
        }
      });

      if (bestPage !== null && bestPage !== currentPageRef.current) {
        onPageChange(bestPage);
      }
    };

    const onScroll = () => {
      if (rafId !== null) return;
      rafId = window.requestAnimationFrame(detectVisiblePage);
    };

    root.addEventListener('scroll', onScroll, {passive: true});
    return () => {
      root.removeEventListener('scroll', onScroll);
      if (rafId !== null) window.cancelAnimationFrame(rafId);
    };
  }, [layoutMode, numPages, onPageChange]);

  const visiblePages = useMemo(() => {
    if (layoutMode === 'continuous') {
      return Array.from({length: numPages}, (_, i) => i + 1);
    }
    if (layoutMode === 'two-page') {
      const left = currentPage % 2 === 1 ? currentPage : Math.max(1, currentPage - 1);
      return left + 1 <= numPages ? [left, left + 1] : [left];
    }
    return [Math.min(numPages, Math.max(1, currentPage))];
  }, [layoutMode, numPages, currentPage]);

  const handleConfirmTextPrompt = (e: React.FormEvent) => {
    e.preventDefault();
    if (!textPrompt || !textPrompt.value.trim()) {
      setTextPrompt(null);
      return;
    }
    onAddAnnotation({
      id: makeAnnId(),
      pageIndex: textPrompt.pageIndex,
      tool: textPrompt.tool,
      x: Math.min(0.8, textPrompt.x),
      y: Math.min(0.88, textPrompt.y),
      width: textPrompt.tool === 'sticky' ? 0.24 : 0.28,
      height: textPrompt.tool === 'sticky' ? 0.1 : 0.05,
      color: textPrompt.tool === 'sticky' ? '#B45309' : toolColor,
      opacity: 1,
      strokeWidth: 1,
      text: textPrompt.value.trim(),
    });
    setTextPrompt(null);
  };

  return (
    <div className="relative flex-1 overflow-hidden bg-slate-200/75 dark:bg-slate-950">
      <div
        ref={scrollContainerRef}
        data-testid="pdf-scroll-container"
        className="h-full w-full overflow-auto px-4 py-6 pb-14 sm:px-6"
      >
        <div
          className={
            layoutMode === 'two-page'
              ? 'mx-auto flex flex-col items-center justify-center gap-6 lg:flex-row lg:items-start'
              : 'mx-auto flex flex-col items-center gap-7'
          }
        >
          {visiblePages.map((pageNumber) => (
            <PdfPageSurface
              key={pageNumber}
              pdfDoc={pdfDoc}
              pageNumber={pageNumber}
              totalPages={numPages}
              containerSize={containerSize}
              layoutMode={layoutMode}
              zoomMode={zoomMode}
              zoom={zoom}
              viewRotation={viewRotation}
              darkReading={darkReading}
              activeTool={activeTool}
              toolColor={toolColor}
              toolOpacity={toolOpacity}
              toolStrokeWidth={toolStrokeWidth}
              pendingStampText={pendingStampText}
              pendingImageDataUrl={pendingImageDataUrl}
              onConsumePendingImage={onConsumePendingImage}
              annotations={annotations}
              selectedAnnotationId={selectedAnnotationId}
              onSelectAnnotation={onSelectAnnotation}
              onAddAnnotation={onAddAnnotation}
              onDeleteAnnotation={onDeleteAnnotation}
              onOpenTextPrompt={setTextPrompt}
              watermark={watermark}
              headerFooter={headerFooter}
              pageNumbers={pageNumbers}
              isSearchMatch={searchHighlightPages.has(pageNumber)}
              isActiveSearchMatch={activeSearchPage === pageNumber}
              readOnly={readOnly}
              registerPageElement={registerPageElement}
            />
          ))}
        </div>
      </div>

      {/* Scroll-following floating Page X of Y indicator */}
      {numPages > 0 && (
        <div
          data-testid="pdf-page-indicator"
          aria-live="polite"
          className="pointer-events-none absolute bottom-3 right-4 z-10 rounded-lg border border-slate-300/90 bg-slate-900/85 px-3 py-1.5 text-xs font-semibold tabular-nums text-white shadow-md backdrop-blur-xs dark:border-slate-700 dark:bg-slate-800/90"
        >
          Page {currentPage} of {numPages}
        </div>
      )}

      {/* Inline Text / Sticky Note Entry Modal */}
      {textPrompt && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/40 p-4">
          <form
            onSubmit={handleConfirmTextPrompt}
            className="w-full max-w-sm rounded-2xl border border-slate-200 bg-white p-5 shadow-xl dark:border-slate-800 dark:bg-slate-900"
          >
            <h3 className="text-sm font-bold text-slate-900 dark:text-slate-100">
              {textPrompt.tool === 'sticky' ? 'Add Sticky Note' : 'Add Text Annotation'}
            </h3>
            <p className="mt-1 text-xs text-slate-500 dark:text-slate-400">
              Page {textPrompt.pageIndex + 1} (stored on-device)
            </p>
            <textarea
              autoFocus
              rows={3}
              value={textPrompt.value}
              onChange={(e) => setTextPrompt({...textPrompt, value: e.target.value})}
              placeholder={
                textPrompt.tool === 'sticky'
                  ? 'Write your review note...'
                  : 'Enter text to place on page...'
              }
              className="mt-3 w-full rounded-xl border border-slate-200 bg-slate-50 p-3 text-sm text-slate-900 focus:border-blue-600 focus:outline-none dark:border-slate-700 dark:bg-slate-950 dark:text-slate-100"
            />
            <div className="mt-4 flex justify-end gap-2">
              <button
                type="button"
                onClick={() => setTextPrompt(null)}
                className="min-h-[40px] rounded-xl border border-slate-200 px-3.5 text-xs font-semibold text-slate-700 hover:bg-slate-100 dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800"
              >
                Cancel
              </button>
              <button
                type="submit"
                className="min-h-[40px] rounded-xl bg-blue-700 px-4 text-xs font-semibold text-white hover:bg-blue-800 dark:bg-blue-600"
              >
                Place on Page
              </button>
            </div>
          </form>
        </div>
      )}
    </div>
  );
};
