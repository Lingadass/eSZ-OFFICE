import React, {useEffect, useMemo, useRef, useState} from 'react';
import type {PDFDocumentProxy} from 'pdfjs-dist';
import {
  AlertTriangle,
  CheckCircle2,
  ChevronLeft,
  ChevronRight,
  FileText,
  Hash,
  ShieldCheck,
  Sparkles,
  X,
} from 'lucide-react';
import type {
  PdfHeaderFooterConfig,
  PdfPageNumberConfig,
  PdfPageNumberFormat,
  PdfPageNumberPosition,
  PdfPageTextClearance,
  PdfWatermarkConfig,
  PdfWatermarkPosition,
} from './pdfTypes';
import {
  computeOverlayLayoutForPage,
  computePageTextClearance,
  defaultRotationForWatermarkPosition,
  type RawPdfTextItem,
} from './pdfOverlayLayout';
import {bakeAnnotationsAndOverlays} from './pdfPure';

export type PdfOverlayTab = 'page-numbers' | 'header-footer' | 'watermark';

interface PdfOverlaysModalProps {
  open: boolean;
  initialTab: PdfOverlayTab;
  onClose: () => void;
  fileName: string;
  pdfBytes: Uint8Array;
  pdfDoc: PDFDocumentProxy | null;
  numPages: number;
  initialPage: number;
  activePassword?: string;
  initialWatermark: PdfWatermarkConfig;
  initialHeaderFooter: PdfHeaderFooterConfig;
  initialPageNumbers: PdfPageNumberConfig;
  onApplyBytes: (nextBytes: Uint8Array, summary: string) => void;
  onToast: (message: string, tone?: 'info' | 'success' | 'warning' | 'error') => void;
}

const PAGE_NUMBER_PRESETS: Array<{id: PdfPageNumberPosition; label: string; short: string}> = [
  {id: 'top-left', label: 'Top Left', short: 'TL'},
  {id: 'top-center', label: 'Top Center', short: 'TC'},
  {id: 'top-right', label: 'Top Right', short: 'TR'},
  {id: 'bottom-left', label: 'Bottom Left', short: 'BL'},
  {id: 'bottom-center', label: 'Bottom Center', short: 'BC'},
  {id: 'bottom-right', label: 'Bottom Right', short: 'BR'},
];

const WATERMARK_POSITIONS: Array<{
  id: PdfWatermarkPosition;
  label: string;
  hint: string;
}> = [
  {id: 'center-diagonal', label: 'Center Diagonal', hint: '45° translucent center'},
  {id: 'center-horizontal', label: 'Center Horizontal', hint: '0° translucent center'},
  {id: 'top-margin', label: 'Top Safe Margin', hint: 'Inside top margin band'},
  {id: 'bottom-margin', label: 'Bottom Safe Margin', hint: 'Inside bottom margin band'},
  {id: 'bottom-right', label: 'Bottom Right Stamp', hint: 'Corner stamp inside margin'},
];

const WATERMARK_CHIPS = ['CONFIDENTIAL', 'DRAFT', 'INTERNAL USE', 'REVIEW COPY', 'FINAL'];

async function extractPageClearance(
  pdfDoc: PDFDocumentProxy,
  pageNumber: number,
): Promise<PdfPageTextClearance | null> {
  try {
    const page = await pdfDoc.getPage(pageNumber);
    const viewport = page.getViewport({scale: 1});
    const textContent = await page.getTextContent();
    const rawItems = (textContent.items ?? []) as RawPdfTextItem[];
    return computePageTextClearance(rawItems, viewport.width, viewport.height);
  } catch {
    return null;
  }
}

export const PdfOverlaysModal: React.FC<PdfOverlaysModalProps> = ({
  open,
  initialTab,
  onClose,
  fileName,
  pdfBytes,
  pdfDoc,
  numPages,
  initialPage,
  activePassword,
  initialWatermark,
  initialHeaderFooter,
  initialPageNumbers,
  onApplyBytes,
  onToast,
}) => {
  const [activeTab, setActiveTab] = useState<PdfOverlayTab>(initialTab);
  const [previewPage, setPreviewPage] = useState<number>(Math.max(1, initialPage));
  const [showGuides, setShowGuides] = useState(true);
  const [applying, setApplying] = useState(false);

  // Local draft configs (never mutate document until Apply is clicked)
  const [pn, setPn] = useState<PdfPageNumberConfig>(initialPageNumbers);
  const [hf, setHf] = useState<PdfHeaderFooterConfig>(initialHeaderFooter);
  const [wm, setWm] = useState<PdfWatermarkConfig>(initialWatermark);

  // Rendered preview state
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const [pageSizePt, setPageSizePt] = useState<{width: number; height: number}>({
    width: 595,
    height: 842,
  });
  const [previewScale, setPreviewScale] = useState<number>(0.55);
  const [pageClearance, setPageClearance] = useState<PdfPageTextClearance | null>(null);
  const [renderingPreview, setRenderingPreview] = useState(false);

  // Reset local draft when opening modal
  useEffect(() => {
    if (!open) return;
    setActiveTab(initialTab);
    setPreviewPage(Math.min(Math.max(1, numPages), Math.max(1, initialPage)));

    // Enable the tool corresponding to the tab the user explicitly opened so live preview is immediate
    setPn({
      ...initialPageNumbers,
      enabled: initialTab === 'page-numbers' ? true : initialPageNumbers.enabled,
      fontSize: initialPageNumbers.fontSize ?? 9,
      marginPt: initialPageNumbers.marginPt ?? 24,
      ensureSafeMargins: initialPageNumbers.ensureSafeMargins ?? true,
      skipFirstPage: initialPageNumbers.skipFirstPage ?? false,
    });
    setHf({
      ...initialHeaderFooter,
      enabled: initialTab === 'header-footer' ? true : initialHeaderFooter.enabled,
      headerCenter:
        initialTab === 'header-footer' &&
        !initialHeaderFooter.headerLeft &&
        !initialHeaderFooter.headerCenter &&
        !initialHeaderFooter.headerRight &&
        !initialHeaderFooter.footerLeft &&
        !initialHeaderFooter.footerCenter &&
        !initialHeaderFooter.footerRight
          ? fileName.replace(/\.pdf$/i, '')
          : initialHeaderFooter.headerCenter,
      fontSize: initialHeaderFooter.fontSize ?? 9,
      marginPt: initialHeaderFooter.marginPt ?? 24,
      ensureSafeMargins: initialHeaderFooter.ensureSafeMargins ?? true,
      skipFirstPage: initialHeaderFooter.skipFirstPage ?? false,
    });
    setWm({
      ...initialWatermark,
      enabled: initialTab === 'watermark' ? true : initialWatermark.enabled,
      text: initialWatermark.text || 'CONFIDENTIAL',
      opacity: initialWatermark.opacity ?? 0.16,
      fontSize: initialWatermark.fontSize ?? 42,
      position: initialWatermark.position ?? 'center-diagonal',
      rotationDeg:
        initialWatermark.rotationDeg ??
        defaultRotationForWatermarkPosition(initialWatermark.position ?? 'center-diagonal'),
      color: initialWatermark.color || '#64748B',
      skipFirstPage: initialWatermark.skipFirstPage ?? false,
    });
  }, [
    open,
    initialTab,
    initialPage,
    numPages,
    fileName,
    initialPageNumbers,
    initialHeaderFooter,
    initialWatermark,
  ]);

  // Render the selected PDF page into the live preview canvas & inspect its text clearance
  useEffect(() => {
    if (!open || !pdfDoc) return;
    let cancelled = false;
    let renderTask: {cancel: () => void; promise: Promise<void>} | null = null;

    const render = async () => {
      setRenderingPreview(true);
      try {
        const page = await pdfDoc.getPage(previewPage);
        if (cancelled) return;
        const unscaled = page.getViewport({scale: 1});
        const clearance = await extractPageClearance(pdfDoc, previewPage);
        if (cancelled) return;
        setPageSizePt({width: unscaled.width, height: unscaled.height});
        setPageClearance(clearance);

        const maxW = 360;
        const maxH = 440;
        const fitScale = Math.min(maxW / unscaled.width, maxH / unscaled.height);
        setPreviewScale(fitScale);

        const viewport = page.getViewport({scale: fitScale});
        const dpr = Math.min(window.devicePixelRatio || 1, 2);
        const canvas = canvasRef.current;
        if (!canvas || cancelled) return;

        canvas.width = Math.max(1, Math.floor(viewport.width * dpr));
        canvas.height = Math.max(1, Math.floor(viewport.height * dpr));
        canvas.style.width = `${Math.floor(viewport.width)}px`;
        canvas.style.height = `${Math.floor(viewport.height)}px`;

        const ctx = canvas.getContext('2d');
        if (!ctx) return;
        ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
        ctx.fillStyle = '#FFFFFF';
        ctx.fillRect(0, 0, viewport.width, viewport.height);

        renderTask = page.render({canvasContext: ctx, viewport, canvas});
        await renderTask.promise;
      } catch {
        // Ignore cancelled renders
      } finally {
        if (!cancelled) setRenderingPreview(false);
      }
    };

    void render();
    return () => {
      cancelled = true;
      if (renderTask) {
        try {
          renderTask.cancel();
        } catch {
          // ignore
        }
      }
    };
  }, [open, pdfDoc, previewPage]);

  // Live safe-margin layout calculation for the selected preview page
  const liveLayout = useMemo(
    () =>
      computeOverlayLayoutForPage({
        pageIndex: previewPage - 1,
        totalPages: Math.max(1, numPages),
        pageWidth: pageSizePt.width,
        pageHeight: pageSizePt.height,
        watermark: wm,
        headerFooter: hf,
        pageNumbers: pn,
        clearance: pageClearance,
      }),
    [previewPage, numPages, pageSizePt, wm, hf, pn, pageClearance],
  );

  const anyEnabled =
    pn.enabled ||
    (hf.enabled &&
      Boolean(
        hf.headerLeft.trim() ||
          hf.headerCenter.trim() ||
          hf.headerRight.trim() ||
          hf.footerLeft.trim() ||
          hf.footerCenter.trim() ||
          hf.footerRight.trim(),
      )) ||
    (wm.enabled && Boolean(wm.text.trim()));

  const handleApply = async () => {
    if (!anyEnabled) {
      onToast('Enable at least one overlay before applying.', 'warning');
      return;
    }
    if (activePassword) {
      onToast('Remove password protection first before embedding overlays into this PDF.', 'warning');
      return;
    }
    setApplying(true);
    try {
      // Inspect text clearances across all pages so every page gets verified safe-margin placement
      const pageClearances: Array<PdfPageTextClearance | null> = [];
      if (pdfDoc) {
        const count = Math.min(numPages, pdfDoc.numPages);
        for (let p = 1; p <= count; p++) {
          pageClearances.push(await extractPageClearance(pdfDoc, p));
        }
      }

      const baked = await bakeAnnotationsAndOverlays(pdfBytes, [], {
        watermark: wm,
        headerFooter: hf,
        pageNumbers: pn,
        pageClearances,
      });

      const appliedLabels: string[] = [];
      if (pn.enabled) appliedLabels.push('Page Numbers');
      if (hf.enabled) appliedLabels.push('Header/Footer');
      if (wm.enabled && wm.text.trim()) appliedLabels.push('Watermark');

      onApplyBytes(baked, `Applied ${appliedLabels.join(', ')} inside safe margins`);
      onClose();
    } catch (err) {
      onToast(err instanceof Error ? err.message : 'Failed to apply overlays to PDF.', 'error');
    } finally {
      setApplying(false);
    }
  };

  if (!open) return null;

  const previewWidthPx = Math.max(1, Math.floor(pageSizePt.width * previewScale));
  const previewHeightPx = Math.max(1, Math.floor(pageSizePt.height * previewScale));

  const applyHeaderFooterPreset = (preset: 'header-only' | 'footer-only' | 'both' | 'clear') => {
    const baseTitle = fileName.replace(/\.pdf$/i, '');
    const today = new Date().toLocaleDateString('en-US', {
      year: 'numeric',
      month: 'short',
      day: 'numeric',
    });
    if (preset === 'clear') {
      setHf((prev) => ({
        ...prev,
        headerLeft: '',
        headerCenter: '',
        headerRight: '',
        footerLeft: '',
        footerCenter: '',
        footerRight: '',
      }));
      return;
    }
    if (preset === 'header-only') {
      setHf((prev) => ({
        ...prev,
        enabled: true,
        headerLeft: baseTitle,
        headerCenter: '',
        headerRight: today,
        footerLeft: '',
        footerCenter: '',
        footerRight: '',
      }));
      return;
    }
    if (preset === 'footer-only') {
      setHf((prev) => ({
        ...prev,
        enabled: true,
        headerLeft: '',
        headerCenter: '',
        headerRight: '',
        footerLeft: baseTitle,
        footerCenter: '',
        footerRight: today,
      }));
      return;
    }
    setHf((prev) => ({
      ...prev,
      enabled: true,
      headerLeft: baseTitle,
      headerCenter: '',
      headerRight: today,
      footerLeft: 'Confidential',
      footerCenter: '',
      footerRight: 'eSZ OFFICE',
    }));
  };

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/55 p-3 sm:p-5"
      role="dialog"
      aria-modal="true"
      aria-label="PDF Page Numbers, Header/Footer and Watermark"
      data-testid="pdf-overlays-modal"
    >
      <div className="flex max-h-[92vh] w-full max-w-5xl flex-col overflow-hidden rounded-2xl border border-slate-200 bg-white shadow-2xl dark:border-slate-800 dark:bg-slate-900">
        {/* Header */}
        <div className="flex items-center justify-between border-b border-slate-200 px-5 py-3.5 dark:border-slate-800">
          <div>
            <h2 className="text-sm font-bold text-slate-900 dark:text-slate-100">
              Page Numbers, Header &amp; Footer, and Watermark
            </h2>
            <p className="text-xs text-slate-500 dark:text-slate-400">
              Preview placement on any page and verify safe margins before applying to{' '}
              <span className="font-medium text-slate-700 dark:text-slate-300">{fileName}</span>.
            </p>
          </div>
          <button
            type="button"
            onClick={onClose}
            aria-label="Close overlays dialog"
            className="flex h-8 w-8 items-center justify-center rounded-lg text-slate-500 hover:bg-slate-100 dark:hover:bg-slate-800"
          >
            <X className="h-4 w-4" />
          </button>
        </div>

        {/* Body: Left Controls + Right Live Preview */}
        <div className="grid min-h-0 flex-1 grid-cols-1 overflow-y-auto lg:grid-cols-12">
          {/* Left Controls */}
          <div className="flex flex-col border-b border-slate-200 p-4 lg:col-span-6 lg:border-b-0 lg:border-r dark:border-slate-800">
            {/* Tabs */}
            <div className="mb-4 grid grid-cols-3 gap-1 rounded-xl bg-slate-100 p-1 dark:bg-slate-950">
              {(
                [
                  {id: 'page-numbers', label: 'Page Numbers', Icon: Hash, active: pn.enabled},
                  {id: 'header-footer', label: 'Header / Footer', Icon: FileText, active: hf.enabled},
                  {id: 'watermark', label: 'Watermark', Icon: Sparkles, active: wm.enabled},
                ] as const
              ).map(({id, label, Icon, active}) => (
                <button
                  key={id}
                  type="button"
                  onClick={() => setActiveTab(id)}
                  data-testid={`pdf-overlay-tab-${id}`}
                  className={`flex items-center justify-center gap-1.5 rounded-lg px-2.5 py-2 text-xs font-semibold transition-colors ${
                    activeTab === id
                      ? 'bg-white text-blue-700 shadow-xs dark:bg-slate-800 dark:text-blue-400'
                      : 'text-slate-600 hover:text-slate-900 dark:text-slate-400 dark:hover:text-slate-200'
                  }`}
                >
                  <Icon className="h-3.5 w-3.5 shrink-0" />
                  <span className="truncate">{label}</span>
                  {active && (
                    <span
                      className="h-1.5 w-1.5 shrink-0 rounded-full bg-emerald-500"
                      title="Enabled"
                    />
                  )}
                </button>
              ))}
            </div>

            {/* Tab 1: Page Numbers */}
            {activeTab === 'page-numbers' && (
              <div className="space-y-4">
                <label className="flex cursor-pointer items-center justify-between rounded-xl border border-slate-200 bg-slate-50 px-3.5 py-2.5 dark:border-slate-800 dark:bg-slate-950">
                  <div>
                    <div className="text-xs font-bold text-slate-800 dark:text-slate-100">
                      Include Page Numbers
                    </div>
                    <div className="text-[11px] text-slate-500 dark:text-slate-400">
                      Placed inside safe top or bottom margin bands
                    </div>
                  </div>
                  <input
                    type="checkbox"
                    checked={pn.enabled}
                    onChange={(e) => setPn((prev) => ({...prev, enabled: e.target.checked}))}
                    className="h-4 w-4 accent-blue-600"
                  />
                </label>

                <div>
                  <span className="block text-[11px] font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">
                    Position Preset
                  </span>
                  <div className="mt-1.5 grid grid-cols-3 gap-1.5">
                    {PAGE_NUMBER_PRESETS.map((preset) => (
                      <button
                        key={preset.id}
                        type="button"
                        data-testid={`pn-pos-${preset.id}`}
                        onClick={() =>
                          setPn((prev) => ({...prev, enabled: true, position: preset.id}))
                        }
                        className={`rounded-lg border px-2.5 py-2 text-xs font-semibold transition-colors ${
                          pn.position === preset.id
                            ? 'border-blue-600 bg-blue-50 text-blue-700 dark:border-blue-500 dark:bg-blue-950/50 dark:text-blue-300'
                            : 'border-slate-200 bg-white text-slate-700 hover:bg-slate-50 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-200'
                        }`}
                      >
                        {preset.label}
                      </button>
                    ))}
                  </div>
                </div>

                <div className="grid grid-cols-2 gap-3">
                  <label className="space-y-1 text-xs">
                    <span className="font-semibold text-slate-700 dark:text-slate-300">Format</span>
                    <select
                      aria-label="Page number format"
                      value={pn.format}
                      onChange={(e) =>
                        setPn((prev) => ({
                          ...prev,
                          format: e.target.value as PdfPageNumberFormat,
                        }))
                      }
                      className="h-9 w-full rounded-lg border border-slate-200 bg-white px-2.5 text-xs dark:border-slate-700 dark:bg-slate-950"
                    >
                      <option value="page-of-total">Page 1 of {numPages || 1}</option>
                      <option value="slash">1 / {numPages || 1}</option>
                      <option value="number">1, 2, 3</option>
                    </select>
                  </label>
                  <label className="space-y-1 text-xs">
                    <span className="font-semibold text-slate-700 dark:text-slate-300">
                      Start Number
                    </span>
                    <input
                      type="number"
                      min={1}
                      max={9999}
                      value={pn.startNumber}
                      onChange={(e) =>
                        setPn((prev) => ({
                          ...prev,
                          startNumber: Math.max(1, Number.parseInt(e.target.value, 10) || 1),
                        }))
                      }
                      className="h-9 w-full rounded-lg border border-slate-200 bg-white px-2.5 text-xs tabular-nums dark:border-slate-700 dark:bg-slate-950"
                    />
                  </label>
                </div>

                <div className="grid grid-cols-2 gap-3">
                  <label className="space-y-1 text-xs">
                    <span className="flex justify-between font-semibold text-slate-700 dark:text-slate-300">
                      <span>Font Size</span>
                      <span className="tabular-nums text-slate-500">{pn.fontSize ?? 9} pt</span>
                    </span>
                    <input
                      type="range"
                      min={8}
                      max={14}
                      step={1}
                      value={pn.fontSize ?? 9}
                      onChange={(e) => setPn((prev) => ({...prev, fontSize: Number(e.target.value)}))}
                      className="w-full accent-blue-600"
                    />
                  </label>
                  <label className="space-y-1 text-xs">
                    <span className="flex justify-between font-semibold text-slate-700 dark:text-slate-300">
                      <span>Safe Margin Inset</span>
                      <span className="tabular-nums text-slate-500">{pn.marginPt ?? 24} pt</span>
                    </span>
                    <input
                      type="range"
                      min={18}
                      max={48}
                      step={2}
                      value={pn.marginPt ?? 24}
                      onChange={(e) => setPn((prev) => ({...prev, marginPt: Number(e.target.value)}))}
                      className="w-full accent-blue-600"
                    />
                  </label>
                </div>

                <div className="space-y-2 pt-1">
                  <label className="flex cursor-pointer items-center gap-2 text-xs text-slate-700 dark:text-slate-300">
                    <input
                      type="checkbox"
                      checked={Boolean(pn.skipFirstPage)}
                      onChange={(e) => setPn((prev) => ({...prev, skipFirstPage: e.target.checked}))}
                      className="h-4 w-4 accent-blue-600"
                    />
                    <span>Skip cover / first page</span>
                  </label>
                  <label className="flex cursor-pointer items-center gap-2 text-xs text-slate-700 dark:text-slate-300">
                    <input
                      type="checkbox"
                      checked={pn.ensureSafeMargins !== false}
                      onChange={(e) =>
                        setPn((prev) => ({...prev, ensureSafeMargins: e.target.checked}))
                      }
                      className="h-4 w-4 accent-blue-600"
                    />
                    <span>Auto-inset page content if existing text reaches margin edge</span>
                  </label>
                </div>
              </div>
            )}

            {/* Tab 2: Header & Footer */}
            {activeTab === 'header-footer' && (
              <div className="space-y-3.5">
                <label className="flex cursor-pointer items-center justify-between rounded-xl border border-slate-200 bg-slate-50 px-3.5 py-2.5 dark:border-slate-800 dark:bg-slate-950">
                  <div>
                    <div className="text-xs font-bold text-slate-800 dark:text-slate-100">
                      Include Header &amp; Footer
                    </div>
                    <div className="text-[11px] text-slate-500 dark:text-slate-400">
                      Auto-measured Left / Center / Right slots never overlap each other
                    </div>
                  </div>
                  <input
                    type="checkbox"
                    checked={hf.enabled}
                    onChange={(e) => setHf((prev) => ({...prev, enabled: e.target.checked}))}
                    className="h-4 w-4 accent-blue-600"
                  />
                </label>

                <div>
                  <span className="block text-[11px] font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">
                    Quick Position Presets
                  </span>
                  <div className="mt-1.5 flex flex-wrap gap-1.5">
                    <button
                      type="button"
                      onClick={() => applyHeaderFooterPreset('header-only')}
                      className="rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs font-semibold text-slate-700 hover:bg-slate-50 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-200"
                    >
                      Header Only
                    </button>
                    <button
                      type="button"
                      onClick={() => applyHeaderFooterPreset('footer-only')}
                      className="rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs font-semibold text-slate-700 hover:bg-slate-50 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-200"
                    >
                      Footer Only
                    </button>
                    <button
                      type="button"
                      onClick={() => applyHeaderFooterPreset('both')}
                      className="rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs font-semibold text-slate-700 hover:bg-slate-50 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-200"
                    >
                      Header &amp; Footer
                    </button>
                    <button
                      type="button"
                      onClick={() => applyHeaderFooterPreset('clear')}
                      className="rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs font-medium text-slate-500 hover:bg-slate-50 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-400"
                    >
                      Clear
                    </button>
                  </div>
                </div>

                <div className="space-y-2">
                  <span className="block text-[11px] font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">
                    Top Header (Safe Top Band)
                  </span>
                  <div className="grid grid-cols-3 gap-1.5">
                    <input
                      type="text"
                      value={hf.headerLeft}
                      onChange={(e) =>
                        setHf((prev) => ({...prev, enabled: true, headerLeft: e.target.value}))
                      }
                      placeholder="Left"
                      aria-label="Header left text"
                      className="h-9 rounded-lg border border-slate-200 bg-white px-2.5 text-xs dark:border-slate-700 dark:bg-slate-950"
                    />
                    <input
                      type="text"
                      value={hf.headerCenter}
                      onChange={(e) =>
                        setHf((prev) => ({...prev, enabled: true, headerCenter: e.target.value}))
                      }
                      placeholder="Center"
                      aria-label="Header center text"
                      className="h-9 rounded-lg border border-slate-200 bg-white px-2.5 text-xs dark:border-slate-700 dark:bg-slate-950"
                    />
                    <input
                      type="text"
                      value={hf.headerRight}
                      onChange={(e) =>
                        setHf((prev) => ({...prev, enabled: true, headerRight: e.target.value}))
                      }
                      placeholder="Right"
                      aria-label="Header right text"
                      className="h-9 rounded-lg border border-slate-200 bg-white px-2.5 text-xs dark:border-slate-700 dark:bg-slate-950"
                    />
                  </div>
                </div>

                <div className="space-y-2">
                  <span className="block text-[11px] font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">
                    Bottom Footer (Safe Bottom Band)
                  </span>
                  <div className="grid grid-cols-3 gap-1.5">
                    <input
                      type="text"
                      value={hf.footerLeft}
                      onChange={(e) =>
                        setHf((prev) => ({...prev, enabled: true, footerLeft: e.target.value}))
                      }
                      placeholder="Left"
                      aria-label="Footer left text"
                      className="h-9 rounded-lg border border-slate-200 bg-white px-2.5 text-xs dark:border-slate-700 dark:bg-slate-950"
                    />
                    <input
                      type="text"
                      value={hf.footerCenter}
                      onChange={(e) =>
                        setHf((prev) => ({...prev, enabled: true, footerCenter: e.target.value}))
                      }
                      placeholder="Center"
                      aria-label="Footer center text"
                      className="h-9 rounded-lg border border-slate-200 bg-white px-2.5 text-xs dark:border-slate-700 dark:bg-slate-950"
                    />
                    <input
                      type="text"
                      value={hf.footerRight}
                      onChange={(e) =>
                        setHf((prev) => ({...prev, enabled: true, footerRight: e.target.value}))
                      }
                      placeholder="Right"
                      aria-label="Footer right text"
                      className="h-9 rounded-lg border border-slate-200 bg-white px-2.5 text-xs dark:border-slate-700 dark:bg-slate-950"
                    />
                  </div>
                </div>

                <div className="grid grid-cols-2 gap-3">
                  <label className="space-y-1 text-xs">
                    <span className="flex justify-between font-semibold text-slate-700 dark:text-slate-300">
                      <span>Font Size</span>
                      <span className="tabular-nums text-slate-500">{hf.fontSize ?? 9} pt</span>
                    </span>
                    <input
                      type="range"
                      min={7}
                      max={13}
                      step={1}
                      value={hf.fontSize ?? 9}
                      onChange={(e) => setHf((prev) => ({...prev, fontSize: Number(e.target.value)}))}
                      className="w-full accent-blue-600"
                    />
                  </label>
                  <label className="space-y-1 text-xs">
                    <span className="flex justify-between font-semibold text-slate-700 dark:text-slate-300">
                      <span>Safe Margin Inset</span>
                      <span className="tabular-nums text-slate-500">{hf.marginPt ?? 24} pt</span>
                    </span>
                    <input
                      type="range"
                      min={18}
                      max={48}
                      step={2}
                      value={hf.marginPt ?? 24}
                      onChange={(e) => setHf((prev) => ({...prev, marginPt: Number(e.target.value)}))}
                      className="w-full accent-blue-600"
                    />
                  </label>
                </div>

                <div className="space-y-2">
                  <label className="flex cursor-pointer items-center gap-2 text-xs text-slate-700 dark:text-slate-300">
                    <input
                      type="checkbox"
                      checked={Boolean(hf.skipFirstPage)}
                      onChange={(e) => setHf((prev) => ({...prev, skipFirstPage: e.target.checked}))}
                      className="h-4 w-4 accent-blue-600"
                    />
                    <span>Skip cover / first page</span>
                  </label>
                  <label className="flex cursor-pointer items-center gap-2 text-xs text-slate-700 dark:text-slate-300">
                    <input
                      type="checkbox"
                      checked={hf.ensureSafeMargins !== false}
                      onChange={(e) =>
                        setHf((prev) => ({...prev, ensureSafeMargins: e.target.checked}))
                      }
                      className="h-4 w-4 accent-blue-600"
                    />
                    <span>Auto-inset page content if existing text reaches margin edge</span>
                  </label>
                </div>
              </div>
            )}

            {/* Tab 3: Watermark */}
            {activeTab === 'watermark' && (
              <div className="space-y-3.5">
                <label className="flex cursor-pointer items-center justify-between rounded-xl border border-slate-200 bg-slate-50 px-3.5 py-2.5 dark:border-slate-800 dark:bg-slate-950">
                  <div>
                    <div className="text-xs font-bold text-slate-800 dark:text-slate-100">
                      Include Watermark
                    </div>
                    <div className="text-[11px] text-slate-500 dark:text-slate-400">
                      Translucent center or safe margin-band stamp
                    </div>
                  </div>
                  <input
                    type="checkbox"
                    checked={wm.enabled}
                    onChange={(e) => setWm((prev) => ({...prev, enabled: e.target.checked}))}
                    className="h-4 w-4 accent-blue-600"
                  />
                </label>

                <div className="space-y-1.5">
                  <label className="block text-xs font-semibold text-slate-700 dark:text-slate-300">
                    Watermark Text
                  </label>
                  <input
                    type="text"
                    value={wm.text}
                    onChange={(e) =>
                      setWm((prev) => ({...prev, enabled: true, text: e.target.value}))
                    }
                    placeholder="CONFIDENTIAL"
                    className="h-9 w-full rounded-lg border border-slate-200 bg-white px-3 text-xs font-medium dark:border-slate-700 dark:bg-slate-950"
                  />
                  <div className="flex flex-wrap gap-1 pt-0.5">
                    {WATERMARK_CHIPS.map((chip) => (
                      <button
                        key={chip}
                        type="button"
                        onClick={() => setWm((prev) => ({...prev, enabled: true, text: chip}))}
                        className="rounded-md border border-slate-200 bg-slate-50 px-2 py-0.5 text-[11px] font-medium text-slate-600 hover:bg-slate-100 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-300"
                      >
                        {chip}
                      </button>
                    ))}
                  </div>
                </div>

                <div>
                  <span className="block text-[11px] font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">
                    Position Preset
                  </span>
                  <div className="mt-1.5 grid grid-cols-2 gap-1.5">
                    {WATERMARK_POSITIONS.map((preset) => (
                      <button
                        key={preset.id}
                        type="button"
                        data-testid={`wm-pos-${preset.id}`}
                        onClick={() =>
                          setWm((prev) => ({
                            ...prev,
                            enabled: true,
                            position: preset.id,
                            rotationDeg: defaultRotationForWatermarkPosition(preset.id),
                          }))
                        }
                        className={`rounded-lg border px-2.5 py-2 text-left transition-colors ${
                          (wm.position ?? 'center-diagonal') === preset.id
                            ? 'border-blue-600 bg-blue-50 text-blue-700 dark:border-blue-500 dark:bg-blue-950/50 dark:text-blue-300'
                            : 'border-slate-200 bg-white text-slate-700 hover:bg-slate-50 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-200'
                        }`}
                      >
                        <div className="text-xs font-semibold">{preset.label}</div>
                        <div className="text-[10px] opacity-75">{preset.hint}</div>
                      </button>
                    ))}
                  </div>
                </div>

                <div className="grid grid-cols-3 gap-2.5">
                  <label className="space-y-1 text-xs">
                    <span className="flex justify-between font-semibold text-slate-700 dark:text-slate-300">
                      <span>Opacity</span>
                      <span className="tabular-nums text-slate-500">
                        {Math.round(wm.opacity * 100)}%
                      </span>
                    </span>
                    <input
                      type="range"
                      min={0.06}
                      max={0.45}
                      step={0.02}
                      value={wm.opacity}
                      onChange={(e) => setWm((prev) => ({...prev, opacity: Number(e.target.value)}))}
                      className="w-full accent-blue-600"
                    />
                  </label>
                  <label className="space-y-1 text-xs">
                    <span className="flex justify-between font-semibold text-slate-700 dark:text-slate-300">
                      <span>Size</span>
                      <span className="tabular-nums text-slate-500">{wm.fontSize} pt</span>
                    </span>
                    <input
                      type="range"
                      min={14}
                      max={68}
                      step={2}
                      value={wm.fontSize}
                      onChange={(e) => setWm((prev) => ({...prev, fontSize: Number(e.target.value)}))}
                      className="w-full accent-blue-600"
                    />
                  </label>
                  <label className="space-y-1 text-xs">
                    <span className="block font-semibold text-slate-700 dark:text-slate-300">
                      Color
                    </span>
                    <input
                      type="color"
                      value={wm.color}
                      onChange={(e) => setWm((prev) => ({...prev, color: e.target.value}))}
                      className="h-8 w-full cursor-pointer rounded-lg border border-slate-200 bg-transparent p-0.5 dark:border-slate-700"
                    />
                  </label>
                </div>

                <label className="flex cursor-pointer items-center gap-2 text-xs text-slate-700 dark:text-slate-300">
                  <input
                    type="checkbox"
                    checked={Boolean(wm.skipFirstPage)}
                    onChange={(e) => setWm((prev) => ({...prev, skipFirstPage: e.target.checked}))}
                    className="h-4 w-4 accent-blue-600"
                  />
                  <span>Skip cover / first page</span>
                </label>
              </div>
            )}
          </div>

          {/* Right Pane: Live Preview on Selected Page */}
          <div className="flex flex-col bg-slate-100/80 p-4 lg:col-span-6 dark:bg-slate-950/70">
            {/* Preview Top Controls */}
            <div className="mb-3 flex flex-wrap items-center justify-between gap-2">
              <div className="flex items-center gap-1.5">
                <button
                  type="button"
                  disabled={previewPage <= 1}
                  onClick={() => setPreviewPage((p) => Math.max(1, p - 1))}
                  aria-label="Preview previous page"
                  className="flex h-7 w-7 items-center justify-center rounded-lg border border-slate-200 bg-white text-slate-700 hover:bg-slate-50 disabled:opacity-40 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-200"
                >
                  <ChevronLeft className="h-4 w-4" />
                </button>
                <span
                  data-testid="pdf-overlay-preview-page"
                  className="text-xs font-semibold tabular-nums text-slate-700 dark:text-slate-200"
                >
                  Preview: Page {previewPage} of {numPages || 1}
                </span>
                <button
                  type="button"
                  disabled={previewPage >= numPages}
                  onClick={() => setPreviewPage((p) => Math.min(numPages, p + 1))}
                  aria-label="Preview next page"
                  className="flex h-7 w-7 items-center justify-center rounded-lg border border-slate-200 bg-white text-slate-700 hover:bg-slate-50 disabled:opacity-40 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-200"
                >
                  <ChevronRight className="h-4 w-4" />
                </button>
              </div>

              <label className="flex cursor-pointer items-center gap-1.5 text-[11px] font-medium text-slate-600 dark:text-slate-400">
                <input
                  type="checkbox"
                  checked={showGuides}
                  onChange={(e) => setShowGuides(e.target.checked)}
                  className="h-3.5 w-3.5 accent-blue-600"
                />
                <span>Safe margin guides</span>
              </label>
            </div>

            {/* Safe Margin Clearance Badge */}
            <div
              data-testid="pdf-overlay-clearance-status"
              data-status={liveLayout.clearanceStatus}
              className={`mb-3 flex items-start gap-2 rounded-xl border px-3 py-2 text-xs ${
                liveLayout.clearanceStatus === 'overlap-risk'
                  ? 'border-amber-300 bg-amber-50 text-amber-900 dark:border-amber-800 dark:bg-amber-950/50 dark:text-amber-200'
                  : liveLayout.clearanceStatus === 'reframed'
                    ? 'border-blue-200 bg-blue-50 text-blue-900 dark:border-blue-800 dark:bg-blue-950/50 dark:text-blue-200'
                    : 'border-emerald-200 bg-emerald-50 text-emerald-900 dark:border-emerald-800 dark:bg-emerald-950/50 dark:text-emerald-200'
              }`}
            >
              {liveLayout.clearanceStatus === 'overlap-risk' ? (
                <AlertTriangle className="mt-0.5 h-4 w-4 shrink-0 text-amber-600 dark:text-amber-400" />
              ) : liveLayout.clearanceStatus === 'reframed' ? (
                <ShieldCheck className="mt-0.5 h-4 w-4 shrink-0 text-blue-600 dark:text-blue-400" />
              ) : (
                <CheckCircle2 className="mt-0.5 h-4 w-4 shrink-0 text-emerald-600 dark:text-emerald-400" />
              )}
              <div className="min-w-0">
                <div className="font-semibold">{liveLayout.clearanceMessage}</div>
                {pageClearance?.hasText && (
                  <div className="mt-0.5 text-[11px] opacity-80 tabular-nums">
                    Measured text clearance — Top: {Math.round(pageClearance.topClearancePt)} pt ·
                    Bottom: {Math.round(pageClearance.bottomClearancePt)} pt
                  </div>
                )}
              </div>
            </div>

            {/* Live Page Canvas + Exact Scaled Overlay Layer */}
            <div className="flex flex-1 items-center justify-center overflow-auto py-1">
              <div
                data-testid="pdf-overlay-live-preview"
                className="relative overflow-hidden rounded-md bg-white shadow-lg ring-1 ring-slate-300 dark:ring-slate-700"
                style={{width: `${previewWidthPx}px`, height: `${previewHeightPx}px`}}
              >
                {/* Underlying PDF page canvas (mirrors vector safe-margin reframe when active) */}
                <div
                  className="absolute inset-0 origin-top-left"
                  style={
                    liveLayout.reframe.needed
                      ? {
                          transform: `translate(${liveLayout.reframe.offsetX * previewScale}px, ${
                            (pageSizePt.height -
                              liveLayout.reframe.offsetY -
                              pageSizePt.height * liveLayout.reframe.scale) *
                            previewScale
                          }px) scale(${liveLayout.reframe.scale})`,
                        }
                      : undefined
                  }
                >
                  <canvas ref={canvasRef} className="block" />
                </div>

                {/* Safe Margin Boundary Guides */}
                {showGuides && (
                  <div aria-hidden="true" className="pointer-events-none absolute inset-0">
                    {liveLayout.topBandPt > 0 && (
                      <div
                        className="absolute inset-x-0 top-0 border-b border-dashed border-emerald-500/70 bg-emerald-500/5"
                        style={{height: `${liveLayout.topBandPt * previewScale}px`}}
                      />
                    )}
                    {liveLayout.bottomBandPt > 0 && (
                      <div
                        className="absolute inset-x-0 bottom-0 border-t border-dashed border-emerald-500/70 bg-emerald-500/5"
                        style={{height: `${liveLayout.bottomBandPt * previewScale}px`}}
                      />
                    )}
                    <div
                      className="absolute inset-y-0 left-0 border-r border-dashed border-slate-400/35"
                      style={{width: `${liveLayout.safeMarginX * previewScale}px`}}
                    />
                    <div
                      className="absolute inset-y-0 right-0 border-l border-dashed border-slate-400/35"
                      style={{width: `${liveLayout.safeMarginX * previewScale}px`}}
                    />
                  </div>
                )}

                {/* Watermark Preview */}
                {liveLayout.watermarkItem && (
                  <div
                    aria-hidden="true"
                    data-testid="pdf-preview-watermark"
                    className="pointer-events-none absolute whitespace-nowrap font-bold uppercase tracking-wider select-none"
                    style={{
                      left: `${liveLayout.watermarkItem.centerX * previewScale}px`,
                      top: `${liveLayout.watermarkItem.previewTopY * previewScale}px`,
                      fontSize: `${Math.max(7, liveLayout.watermarkItem.fontSize * previewScale)}px`,
                      color: liveLayout.watermarkItem.color,
                      opacity: liveLayout.watermarkItem.opacity,
                      transform: `translate(-50%, -50%) rotate(${-liveLayout.watermarkItem.rotationDeg}deg)`,
                      lineHeight: 1,
                    }}
                  >
                    {liveLayout.watermarkItem.text}
                  </div>
                )}

                {/* Header & Footer Preview Items */}
                {[...liveLayout.headerItems, ...liveLayout.footerItems].map((item, idx) => (
                  <div
                    key={`hf_${idx}`}
                    aria-hidden="true"
                    data-testid="pdf-preview-header-footer"
                    className="pointer-events-none absolute whitespace-nowrap text-slate-600 select-none"
                    style={{
                      left: `${item.x * previewScale}px`,
                      top: `${item.topY * previewScale}px`,
                      fontSize: `${Math.max(6, item.fontSize * previewScale)}px`,
                      lineHeight: 1.1,
                    }}
                  >
                    {item.text}
                  </div>
                ))}

                {/* Page Number Preview Item */}
                {liveLayout.pageNumberItem && (
                  <div
                    aria-hidden="true"
                    data-testid="pdf-preview-page-number"
                    className="pointer-events-none absolute whitespace-nowrap font-medium text-slate-700 select-none"
                    style={{
                      left: `${liveLayout.pageNumberItem.x * previewScale}px`,
                      top: `${liveLayout.pageNumberItem.topY * previewScale}px`,
                      fontSize: `${Math.max(6, liveLayout.pageNumberItem.fontSize * previewScale)}px`,
                      lineHeight: 1.1,
                    }}
                  >
                    {liveLayout.pageNumberItem.text}
                  </div>
                )}

                {renderingPreview && (
                  <div className="absolute inset-0 flex items-center justify-center bg-white/50 text-xs font-medium text-slate-500">
                    Rendering page...
                  </div>
                )}
              </div>
            </div>
          </div>
        </div>

        {/* Footer Actions */}
        <div className="flex flex-wrap items-center justify-between gap-2 border-t border-slate-200 bg-slate-50 px-5 py-3 dark:border-slate-800 dark:bg-slate-950">
          <div className="flex flex-wrap items-center gap-2 text-[11px] text-slate-500 dark:text-slate-400">
            <span>Active on apply:</span>
            {pn.enabled && (
              <span className="rounded bg-blue-100 px-1.5 py-0.5 font-semibold text-blue-800 dark:bg-blue-950 dark:text-blue-300">
                Page Numbers
              </span>
            )}
            {hf.enabled && (
              <span className="rounded bg-blue-100 px-1.5 py-0.5 font-semibold text-blue-800 dark:bg-blue-950 dark:text-blue-300">
                Header/Footer
              </span>
            )}
            {wm.enabled && wm.text.trim() && (
              <span className="rounded bg-blue-100 px-1.5 py-0.5 font-semibold text-blue-800 dark:bg-blue-950 dark:text-blue-300">
                Watermark
              </span>
            )}
            {!anyEnabled && <span className="italic">None selected</span>}
          </div>
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={onClose}
              className="min-h-[36px] rounded-xl border border-slate-200 bg-white px-4 text-xs font-semibold text-slate-700 hover:bg-slate-100 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-200 dark:hover:bg-slate-800"
            >
              Cancel
            </button>
            <button
              type="button"
              disabled={!anyEnabled || applying}
              onClick={() => void handleApply()}
              data-testid="pdf-overlay-apply-btn"
              className="min-h-[36px] rounded-xl bg-blue-700 px-4 text-xs font-semibold text-white hover:bg-blue-800 disabled:opacity-45 dark:bg-blue-600 dark:hover:bg-blue-500"
            >
              {applying ? 'Applying safely...' : 'Apply to PDF'}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
