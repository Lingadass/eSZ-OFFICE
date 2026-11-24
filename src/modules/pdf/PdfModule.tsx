import React, {useCallback, useEffect, useMemo, useRef, useState} from 'react';
import * as pdfjsLib from 'pdfjs-dist';
import type {PDFDocumentProxy} from 'pdfjs-dist';
import {
  ArrowUpRight,
  BookOpen,
  Check,
  ChevronLeft,
  ChevronRight,
  Circle,
  Columns2,
  Download,
  EyeOff,
  FileImage,
  FileText,
  Hash,
  Highlighter,
  Layers,
  Lock,
  Maximize2,
  Minimize2,
  Moon,
  MousePointer2,
  PanelLeft,
  Pencil,
  PenTool,
  Printer,
  Redo2,
  RotateCw,
  Save,
  ScanText,
  ScrollText,
  Search,
  Square,
  Stamp,
  Star,
  StickyNote,
  Strikethrough,
  Sun,
  Type,
  Underline,
  Undo2,
  Wand2,
  X,
  ZoomIn,
  ZoomOut,
} from 'lucide-react';
import type {OfficeFileRecord, TabMode, WorkspaceTab} from '../../core/types';
import {loadFileBlob, updateFileLastPage} from '../../core/storage/fileRepository';
import {formatBytes} from '../../core/storage/fileRules';
import {useWorkspaceStore} from '../../core/workspace/useWorkspaceStore';
import {ViewModeMobileBar, ViewModeToolbar} from '../../shell/ViewModeToolbar';
import {ensurePdfJsWorker, ensureValidPdfBytes} from './pdfEngine';
import {bakeAnnotationsAndOverlays} from './pdfPure';
import type {
  PdfAnnotationItem,
  PdfAnnotationTool,
  PdfHeaderFooterConfig,
  PdfLayoutMode,
  PdfPageNumberConfig,
  PdfWatermarkConfig,
  PdfZoomMode,
} from './pdfTypes';
import {PdfViewerCanvas} from './PdfViewerCanvas';
import {PdfOrganizeModal} from './PdfOrganizeModal';
import {PdfFormsAndSignModal} from './PdfFormsAndSignModal';
import {PdfSecurityAndConvertModal} from './PdfSecurityAndConvertModal';
import {PdfOverlaysModal, type PdfOverlayTab} from './PdfOverlaysModal';

export interface OfficeModuleProps {
  tab: WorkspaceTab;
  mode: TabMode;
  fileRecord?: OfficeFileRecord;
  canUndo: boolean;
  canRedo: boolean;
  onChangeContent: (next: string) => void;
  onUndo: () => void;
  onRedo: () => void;
  onSaveNow: () => void;
  onExportDevice: () => void;
  onToggleFavorite: () => void;
  onDismissRecovery: () => void;
  onEnterEdit: () => void;
  onDoneEditing: () => void;
  onRequestExitEdit: () => void;
}

interface OutlineNode {
  title: string;
  pageNumber: number | null;
}

interface PdfSnapshot {
  bytes: Uint8Array;
  annotations: PdfAnnotationItem[];
}

const SidebarThumbnail: React.FC<{
  pdfDoc: PDFDocumentProxy;
  pageNumber: number;
  active: boolean;
  onClick: () => void;
}> = ({pdfDoc, pageNumber, active, onClick}) => {
  const canvasRef = useRef<HTMLCanvasElement | null>(null);

  useEffect(() => {
    let cancelled = false;
    const renderThumb = async () => {
      try {
        const page = await pdfDoc.getPage(pageNumber);
        if (cancelled) return;
        const unscaled = page.getViewport({scale: 1});
        const scale = Math.min(120 / unscaled.width, 152 / unscaled.height);
        const viewport = page.getViewport({scale});
        const canvas = canvasRef.current;
        if (!canvas || cancelled) return;
        canvas.width = Math.max(1, Math.floor(viewport.width));
        canvas.height = Math.max(1, Math.floor(viewport.height));
        const ctx = canvas.getContext('2d');
        if (!ctx) return;
        ctx.fillStyle = '#FFFFFF';
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        await page.render({canvasContext: ctx, viewport, canvas}).promise;
      } catch {
        // ignore thumbnail render abort
      }
    };
    void renderThumb();
    return () => {
      cancelled = true;
    };
  }, [pdfDoc, pageNumber]);

  return (
    <button
      type="button"
      onClick={onClick}
      className={`flex w-full flex-col items-center rounded-xl border p-2 transition-colors ${
        active
          ? 'border-blue-600 bg-blue-50 dark:border-blue-500 dark:bg-blue-950/40'
          : 'border-slate-200 bg-white hover:bg-slate-100 dark:border-slate-800 dark:bg-slate-900 dark:hover:bg-slate-800'
      }`}
    >
      <canvas ref={canvasRef} className="max-h-36 max-w-[120px] rounded border border-slate-200 bg-white" />
      <span className="mt-1.5 text-[11px] font-semibold tabular-nums text-slate-600 dark:text-slate-300">
        Page {pageNumber}
      </span>
    </button>
  );
};

const TOOL_ITEMS: Array<{
  id: PdfAnnotationTool;
  label: string;
  Icon: React.FC<{className?: string}>;
}> = [
  {id: 'select', label: 'Select', Icon: MousePointer2},
  {id: 'highlight', label: 'Highlight', Icon: Highlighter},
  {id: 'underline', label: 'Underline', Icon: Underline},
  {id: 'strikethrough', label: 'Strike', Icon: Strikethrough},
  {id: 'sticky', label: 'Note', Icon: StickyNote},
  {id: 'freehand', label: 'Draw', Icon: Pencil},
  {id: 'rect', label: 'Box', Icon: Square},
  {id: 'ellipse', label: 'Oval', Icon: Circle},
  {id: 'arrow', label: 'Arrow', Icon: ArrowUpRight},
  {id: 'text', label: 'Text', Icon: Type},
  {id: 'stamp', label: 'Stamp', Icon: Stamp},
  {id: 'image', label: 'Image', Icon: FileImage},
  {id: 'redact', label: 'Redact', Icon: EyeOff},
];

const STAMP_PRESETS = ['APPROVED', 'CONFIDENTIAL', 'DRAFT', 'REVIEWED', 'FINAL'];

const PdfModule: React.FC<OfficeModuleProps> = ({
  tab,
  mode,
  fileRecord,
  onExportDevice,
  onToggleFavorite,
  onEnterEdit,
  onDoneEditing,
}) => {
  const isEditing = mode === 'edit';
  const {
    files,
    saveTabBlob,
    markTabDirty,
    registerBinarySaveHandler,
    createFileFromBytes,
    pushToast,
  } = useWorkspaceStore();

  const rootRef = useRef<HTMLDivElement | null>(null);
  const imageInputRef = useRef<HTMLInputElement | null>(null);

  // Core PDF binary & pdf.js document state
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState<string | null>(null);
  const [pdfBytes, setPdfBytes] = useState<Uint8Array | null>(null);
  const [pdfDoc, setPdfDoc] = useState<PDFDocumentProxy | null>(null);
  const [numPages, setNumPages] = useState(0);
  const [passwordNeeded, setPasswordNeeded] = useState(false);
  const [passwordInput, setPasswordInput] = useState('');
  const [activePassword, setActivePassword] = useState<string | undefined>(undefined);

  // Undo / Redo snapshot stacks
  const [past, setPast] = useState<PdfSnapshot[]>([]);
  const [future, setFuture] = useState<PdfSnapshot[]>([]);

  // Viewer navigation & display settings
  const [currentPage, setCurrentPage] = useState<number>(fileRecord?.lastPage ?? 1);
  const [pageInput, setPageInput] = useState<string>(String(fileRecord?.lastPage ?? 1));
  const [scrollSignal, setScrollSignal] = useState(0);
  const [layoutMode, setLayoutMode] = useState<PdfLayoutMode>('continuous');
  const [zoomMode, setZoomMode] = useState<PdfZoomMode>('fit-width');
  const [zoom, setZoom] = useState(1.1);
  const [viewRotation, setViewRotation] = useState(0);
  const [darkReading, setDarkReading] = useState(false);
  const [isFullscreen, setIsFullscreen] = useState(false);

  // Sidebar (Thumbnails / Outline)
  const [sidebarOpen, setSidebarOpen] = useState(() =>
    typeof window !== 'undefined' ? window.innerWidth >= 1024 : true,
  );
  const [sidebarTab, setSidebarTab] = useState<'thumbnails' | 'outline'>('thumbnails');
  const [outline, setOutline] = useState<OutlineNode[]>([]);

  // Search state
  const [searchOpen, setSearchOpen] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [searchMatches, setSearchMatches] = useState<number[]>([]);
  const [searchIdx, setSearchIdx] = useState(0);
  const [searching, setSearching] = useState(false);

  // Annotation & overlay state
  const [activeTool, setActiveTool] = useState<PdfAnnotationTool>('select');
  const [toolColor, setToolColor] = useState('#FACC15');
  const [toolOpacity, setToolOpacity] = useState(0.9);
  const [toolStrokeWidth, setToolStrokeWidth] = useState(2);
  const [pendingStampText, setPendingStampText] = useState('APPROVED');
  const [pendingImageDataUrl, setPendingImageDataUrl] = useState<string | null>(null);
  const [annotations, setAnnotations] = useState<PdfAnnotationItem[]>([]);
  const [selectedAnnotationId, setSelectedAnnotationId] = useState<string | null>(null);

  // Overlays studio modal (Page Numbers / Header & Footer / Watermark with live preview)
  const [overlaysModal, setOverlaysModal] = useState<{open: boolean; tab: PdfOverlayTab}>({
    open: false,
    tab: 'page-numbers',
  });
  const [watermark, setWatermark] = useState<PdfWatermarkConfig>({
    enabled: false,
    text: 'CONFIDENTIAL',
    opacity: 0.16,
    fontSize: 42,
    rotationDeg: 45,
    color: '#64748B',
    position: 'center-diagonal',
  });
  const [headerFooter, setHeaderFooter] = useState<PdfHeaderFooterConfig>({
    enabled: false,
    headerLeft: '',
    headerCenter: '',
    headerRight: '',
    footerLeft: '',
    footerCenter: '',
    footerRight: '',
    fontSize: 9,
    marginPt: 24,
    ensureSafeMargins: true,
  });
  const [pageNumbers, setPageNumbers] = useState<PdfPageNumberConfig>({
    enabled: false,
    position: 'bottom-center',
    format: 'page-of-total',
    startNumber: 1,
    fontSize: 9,
    marginPt: 24,
    ensureSafeMargins: true,
  });

  // Modals
  const [organizeOpen, setOrganizeOpen] = useState(false);
  const [formsModal, setFormsModal] = useState<{open: boolean; tab: 'forms' | 'signature'}>({
    open: false,
    tab: 'signature',
  });
  const [securityModal, setSecurityModal] = useState<{
    open: boolean;
    tab: 'security' | 'convert' | 'ocr';
  }>({open: false, tab: 'security'});

  // Refs for latest state inside autosave handler
  const latestStateRef = useRef({
    pdfBytes,
    annotations,
    watermark,
    headerFooter,
    pageNumbers,
    currentPage,
  });
  latestStateRef.current = {
    pdfBytes,
    annotations,
    watermark,
    headerFooter,
    pageNumbers,
    currentPage,
  };

  // Load binary blob for this tab
  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    setLoadError(null);
    setPasswordNeeded(false);
    setAnnotations([]);
    setPast([]);
    setFuture([]);
    setSearchMatches([]);
    setActivePassword(undefined);

    const loadInitial = async () => {
      try {
        const rawBlob = await loadFileBlob(tab.fileId);
        if (cancelled) return;
        const blobToUse = rawBlob ?? new Blob([''], {type: 'application/pdf'});
        const validBytes = await ensureValidPdfBytes(blobToUse, tab.name);
        if (cancelled) return;

        // Keep upgraded PDF bytes purely in memory on open so view mode never writes or changes updatedAt.
        setPdfBytes(validBytes);
      } catch (err) {
        if (!cancelled) {
          setLoadError(err instanceof Error ? err.message : 'Unable to open PDF file.');
          setLoading(false);
        }
      }
    };

    void loadInitial();
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tab.fileId, tab.discardNonce]);

  // Reset edit-only drawers/tools whenever switching into read-only view mode
  useEffect(() => {
    if (!isEditing) {
      setActiveTool('select');
      setSelectedAnnotationId(null);
      setOverlaysModal((m) => ({...m, open: false}));
      setOrganizeOpen(false);
      setFormsModal((m) => ({...m, open: false}));
      setSecurityModal((m) => ({...m, open: false}));
    }
  }, [isEditing]);

  // Parse pdfBytes with pdf.js whenever bytes or activePassword changes
  useEffect(() => {
    if (!pdfBytes) return;
    let cancelled = false;
    ensurePdfJsWorker();

    const task = pdfjsLib.getDocument({
      data: new Uint8Array(pdfBytes),
      password: activePassword,
    });

    const openWithPdfJs = async () => {
      setLoading(true);
      setLoadError(null);
      try {
        const doc = await task.promise;
        if (cancelled) {
          await task.destroy();
          return;
        }
        setPdfDoc(doc);
        setNumPages(doc.numPages);
        setPasswordNeeded(false);

        const initialPage = Math.min(
          doc.numPages,
          Math.max(1, latestStateRef.current.currentPage || fileRecord?.lastPage || 1),
        );
        setCurrentPage(initialPage);
        setPageInput(String(initialPage));

        // Extract outline bookmarks if present
        try {
          const rawOutline = await doc.getOutline();
          if (rawOutline && rawOutline.length > 0 && !cancelled) {
            const parsed: OutlineNode[] = [];
            for (const item of rawOutline.slice(0, 80)) {
              let pageNum: number | null = null;
              try {
                let dest = item.dest;
                if (typeof dest === 'string') {
                  dest = await doc.getDestination(dest);
                }
                if (Array.isArray(dest) && dest[0]) {
                  const idx = await doc.getPageIndex(dest[0]);
                  pageNum = idx + 1;
                }
              } catch {
                // ignore unresolvable destination
              }
              parsed.push({title: item.title || 'Section', pageNumber: pageNum});
            }
            setOutline(parsed);
          } else {
            setOutline([]);
          }
        } catch {
          setOutline([]);
        }
      } catch (err: unknown) {
        if (cancelled) return;
        const errName = (err as {name?: string})?.name;
        if (errName === 'PasswordException') {
          setPasswordNeeded(true);
          setLoadError(null);
        } else {
          setLoadError(err instanceof Error ? err.message : 'Failed to render PDF document.');
        }
      } finally {
        if (!cancelled) setLoading(false);
      }
    };

    void openWithPdfJs();

    return () => {
      cancelled = true;
      void task.destroy();
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [pdfBytes, activePassword]);

  // Sync page input & remember last viewed page in Dexie only when user navigates away from stored page (never touches updatedAt)
  useEffect(() => {
    setPageInput(String(currentPage));
    if (
      numPages > 0 &&
      currentPage >= 1 &&
      currentPage <= numPages &&
      currentPage !== (fileRecord?.lastPage ?? 1)
    ) {
      void updateFileLastPage(tab.fileId, currentPage);
    }
  }, [currentPage, numPages, tab.fileId, fileRecord?.lastPage]);

  const hasActiveOverlays = watermark.enabled || headerFooter.enabled || pageNumbers.enabled;

  // Persist current PDF bytes (baking any pending non-redaction annotations/overlays)
  const commitSaveToVault = useCallback(
    async (silent = false) => {
      if (mode !== 'edit') return;
      const state = latestStateRef.current;
      if (!state.pdfBytes) return;

      const overlaysActive =
        state.watermark.enabled || state.headerFooter.enabled || state.pageNumbers.enabled;
      const bakeableAnns = state.annotations.filter((a) => a.tool !== 'redact');
      const remainingRedacts = state.annotations.filter((a) => a.tool === 'redact');

      let outputBytes = state.pdfBytes;
      if (bakeableAnns.length > 0 || overlaysActive) {
        try {
          if (activePassword) {
            // Encrypted PDFs cannot be modified directly by pdf-lib without decrypting first
            if (!silent) {
              pushToast('Remove password protection first to embed new annotations.', 'warning');
            }
          } else {
            outputBytes = await bakeAnnotationsAndOverlays(state.pdfBytes, bakeableAnns, {
              watermark: state.watermark,
              headerFooter: state.headerFooter,
              pageNumbers: state.pageNumbers,
            });
            setPdfBytes(outputBytes);
            setAnnotations(remainingRedacts);
            if (overlaysActive) {
              setWatermark((w) => ({...w, enabled: false}));
              setHeaderFooter((h) => ({...h, enabled: false}));
              setPageNumbers((p) => ({...p, enabled: false}));
            }
          }
        } catch (err) {
          pushToast(
            err instanceof Error ? err.message : 'Could not embed annotations into PDF.',
            'error',
          );
          return;
        }
      }

      const blob = new Blob([new Uint8Array(outputBytes)], {type: 'application/pdf'});
      const ok = await saveTabBlob(
        tab.fileId,
        blob,
        `${tab.name} (${numPages || 1} pages)`,
        state.currentPage,
      );
      if (ok && !silent) {
        pushToast('Saved PDF changes locally', 'success');
      }
    },
    [mode, tab.fileId, tab.name, numPages, activePassword, saveTabBlob, pushToast],
  );

  // Register binary save handler so WorkspaceView's 5s autosave and Export use it safely
  useEffect(() => {
    registerBinarySaveHandler(tab.fileId, () => commitSaveToVault(true));
    return () => {
      registerBinarySaveHandler(tab.fileId, null);
    };
  }, [tab.fileId, registerBinarySaveHandler, commitSaveToVault]);

  // Snapshot helpers for Undo/Redo
  const pushSnapshot = useCallback(() => {
    if (!pdfBytes || mode !== 'edit') return;
    setPast((prev) => [...prev.slice(-24), {bytes: pdfBytes, annotations}]);
    setFuture([]);
    markTabDirty(tab.fileId, true);
  }, [pdfBytes, annotations, markTabDirty, tab.fileId, mode]);

  const handleApplyNewBytes = useCallback(
    (nextBytes: Uint8Array, summary: string, nextPassword?: string) => {
      if (mode !== 'edit') return;
      pushSnapshot();
      if (nextPassword !== undefined || summary.toLowerCase().includes('remove')) {
        setActivePassword(nextPassword);
      }
      setPdfBytes(nextBytes);
      pushToast(summary, 'success');
    },
    [pushSnapshot, pushToast, mode],
  );

  const handleAddAnnotation = useCallback(
    (ann: PdfAnnotationItem) => {
      if (mode !== 'edit') return;
      pushSnapshot();
      setAnnotations((prev) => [...prev, ann]);
      if (ann.tool === 'stamp' || ann.tool === 'signature' || ann.tool === 'image') {
        setActiveTool('select');
      }
    },
    [pushSnapshot, mode],
  );

  const handleDeleteAnnotation = useCallback(
    (id: string) => {
      if (mode !== 'edit') return;
      pushSnapshot();
      setAnnotations((prev) => prev.filter((a) => a.id !== id));
      setSelectedAnnotationId(null);
    },
    [pushSnapshot, mode],
  );

  const handleUndo = useCallback(() => {
    if (mode !== 'edit' || past.length === 0 || !pdfBytes) return;
    const prev = past[past.length - 1];
    setFuture((f) => [{bytes: pdfBytes, annotations}, ...f.slice(0, 24)]);
    setPast((p) => p.slice(0, -1));
    setPdfBytes(prev.bytes);
    setAnnotations(prev.annotations);
    markTabDirty(tab.fileId, true);
  }, [mode, past, pdfBytes, annotations, markTabDirty, tab.fileId]);

  const handleRedo = useCallback(() => {
    if (mode !== 'edit' || future.length === 0 || !pdfBytes) return;
    const next = future[0];
    setPast((p) => [...p.slice(-24), {bytes: pdfBytes, annotations}]);
    setFuture((f) => f.slice(1));
    setPdfBytes(next.bytes);
    setAnnotations(next.annotations);
    markTabDirty(tab.fileId, true);
  }, [mode, future, pdfBytes, annotations, markTabDirty, tab.fileId]);

  // Keyboard shortcuts (Ctrl/Cmd+S, Ctrl/Cmd+Z, Ctrl/Cmd+F)
  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      const target = e.target as HTMLElement | null;
      const inInput =
        target &&
        (target.tagName === 'INPUT' ||
          target.tagName === 'TEXTAREA' ||
          target.isContentEditable);

      if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 's') {
        if (!isEditing) return;
        e.preventDefault();
        void commitSaveToVault(false);
        return;
      }
      if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'f') {
        e.preventDefault();
        setSearchOpen(true);
        return;
      }
      if (isEditing && !inInput && (e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'z') {
        e.preventDefault();
        if (e.shiftKey) handleRedo();
        else handleUndo();
      }
    };
    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, [isEditing, commitSaveToVault, handleUndo, handleRedo]);

  // Fullscreen listener
  useEffect(() => {
    const onFs = () => setIsFullscreen(Boolean(document.fullscreenElement));
    document.addEventListener('fullscreenchange', onFs);
    return () => document.removeEventListener('fullscreenchange', onFs);
  }, []);

  // Auto-collapse sidebar on narrow screens so mobile canvas stays unobstructed
  useEffect(() => {
    if (typeof window === 'undefined' || typeof window.matchMedia !== 'function') return;
    const mq = window.matchMedia('(min-width: 1024px)');
    const onChange = (e: MediaQueryListEvent) => {
      if (!e.matches) setSidebarOpen(false);
    };
    mq.addEventListener('change', onChange);
    return () => mq.removeEventListener('change', onChange);
  }, []);

  const toggleFullscreen = async () => {
    if (!document.fullscreenElement && rootRef.current) {
      await rootRef.current.requestFullscreen().catch(() => undefined);
    } else if (document.fullscreenElement) {
      await document.exitFullscreen().catch(() => undefined);
    }
  };

  const handlePrint = async () => {
    if (!pdfBytes) return;
    if (isEditing && tab.dirty) {
      await commitSaveToVault(true);
    }
    const bytesToPrint = latestStateRef.current.pdfBytes ?? pdfBytes;
    const url = URL.createObjectURL(
      new Blob([new Uint8Array(bytesToPrint)], {type: 'application/pdf'}),
    );
    const iframe = document.createElement('iframe');
    iframe.style.position = 'fixed';
    iframe.style.right = '0';
    iframe.style.bottom = '0';
    iframe.style.width = '0';
    iframe.style.height = '0';
    iframe.style.border = '0';
    iframe.src = url;
    document.body.appendChild(iframe);
    iframe.onload = () => {
      try {
        iframe.contentWindow?.focus();
        iframe.contentWindow?.print();
      } catch {
        pushToast('Use Export to download and print on this browser.', 'info');
      }
      window.setTimeout(() => {
        iframe.remove();
        URL.revokeObjectURL(url);
      }, 60_000);
    };
  };

  // Search across all PDF pages via pdf.js textContent
  const runSearch = async (e?: React.FormEvent) => {
    e?.preventDefault();
    if (!pdfDoc || !searchQuery.trim()) {
      setSearchMatches([]);
      return;
    }
    setSearching(true);
    try {
      const needle = searchQuery.trim().toLowerCase();
      const matchedPages: number[] = [];
      for (let p = 1; p <= pdfDoc.numPages; p++) {
        const page = await pdfDoc.getPage(p);
        const content = await page.getTextContent();
        const text = content.items
          .map((item) => ('str' in item ? item.str : ''))
          .join(' ')
          .toLowerCase();
        if (text.includes(needle)) {
          matchedPages.push(p);
        }
      }
      setSearchMatches(matchedPages);
      setSearchIdx(0);
      if (matchedPages.length > 0) {
        jumpToPage(matchedPages[0]);
      } else {
        pushToast(`No matches for "${searchQuery.trim()}"`, 'info');
      }
    } finally {
      setSearching(false);
    }
  };

  const jumpToPage = (pageNum: number) => {
    const clamped = Math.min(Math.max(1, numPages), Math.max(1, pageNum));
    setCurrentPage(clamped);
    setScrollSignal((s) => s + 1);
  };

  const cycleSearchMatch = (delta: number) => {
    if (searchMatches.length === 0) return;
    const nextIdx = (searchIdx + delta + searchMatches.length) % searchMatches.length;
    setSearchIdx(nextIdx);
    jumpToPage(searchMatches[nextIdx]);
  };

  const handleToolSelect = (tool: PdfAnnotationTool) => {
    if (tool === 'image') {
      imageInputRef.current?.click();
      return;
    }
    setActiveTool(tool);
    if (tool === 'highlight') {
      setToolColor('#FACC15');
      setToolOpacity(0.38);
    } else if (tool === 'redact') {
      setToolColor('#000000');
      setToolOpacity(1);
    } else if (tool === 'stamp') {
      setToolColor('#DC2626');
      setToolOpacity(0.9);
    } else if (tool === 'underline' || tool === 'strikethrough' || tool === 'freehand' || tool === 'rect' || tool === 'ellipse' || tool === 'arrow' || tool === 'text') {
      if (toolColor === '#FACC15' || toolColor === '#000000') {
        setToolColor('#1D4ED8');
      }
      setToolOpacity(1);
    }
  };

  const handleImagePicked = (file: File | undefined) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      if (typeof reader.result === 'string') {
        setPendingImageDataUrl(reader.result);
        setActiveTool('image');
        pushToast('Click anywhere on a PDF page to place the image.', 'info');
      }
    };
    reader.readAsDataURL(file);
  };

  const handleReadySignature = (dataUrl: string, placeImmediately: boolean) => {
    if (placeImmediately) {
      handleAddAnnotation({
        id: `sig_${Date.now()}`,
        pageIndex: Math.max(0, currentPage - 1),
        tool: 'signature',
        x: 0.58,
        y: 0.78,
        width: 0.28,
        height: 0.1,
        color: '#0F172A',
        opacity: 1,
        strokeWidth: 1,
        imageDataUrl: dataUrl,
      });
      pushToast(`Placed signature on page ${currentPage}`, 'success');
    } else {
      setPendingImageDataUrl(dataUrl);
      setActiveTool('signature');
      pushToast('Click on the PDF page where you want to place your signature.', 'info');
    }
  };

  const vaultPdfFiles = useMemo(
    () => files.filter((f) => f.moduleType === 'pdf' && f.id !== tab.fileId),
    [files, tab.fileId],
  );

  const pendingRedactions = useMemo(
    () => annotations.filter((a) => a.tool === 'redact'),
    [annotations],
  );

  const searchHighlightSet = useMemo(() => new Set(searchMatches), [searchMatches]);
  const activeSearchPage = searchMatches.length > 0 ? searchMatches[searchIdx] ?? null : null;

  const zoomDisplayLabel =
    zoomMode === 'fit-width'
      ? 'Fit'
      : zoomMode === 'fit-page'
        ? 'Page'
        : `${Math.round(zoom * 100)}%`;

  const pageNavControl = (
    <div className="flex shrink-0 items-center gap-0.5 rounded-lg border border-slate-200 bg-slate-50 px-0.5 py-0.5 dark:border-slate-700 dark:bg-slate-950">
      <button
        type="button"
        disabled={currentPage <= 1}
        onClick={() => jumpToPage(currentPage - 1)}
        aria-label="Previous page"
        className="flex h-7 w-7 items-center justify-center rounded-md text-slate-700 hover:bg-slate-200/70 disabled:opacity-40 dark:text-slate-200 dark:hover:bg-slate-800"
      >
        <ChevronLeft className="h-4 w-4" />
      </button>
      <form
        onSubmit={(e) => {
          e.preventDefault();
          const parsed = Number.parseInt(pageInput, 10);
          if (Number.isFinite(parsed)) jumpToPage(parsed);
        }}
        className="flex items-center gap-1 px-1 text-xs tabular-nums"
      >
        <span className="hidden text-slate-500 sm:inline dark:text-slate-400">Page</span>
        <input
          type="text"
          inputMode="numeric"
          aria-label="Current page number"
          value={pageInput}
          onChange={(e) => setPageInput(e.target.value)}
          className="h-6 w-8 rounded border border-slate-200 bg-white text-center text-xs font-semibold text-slate-800 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-100"
        />
        <span className="text-slate-500 dark:text-slate-400">of {numPages || 1}</span>
      </form>
      <button
        type="button"
        disabled={currentPage >= numPages}
        onClick={() => jumpToPage(currentPage + 1)}
        aria-label="Next page"
        className="flex h-7 w-7 items-center justify-center rounded-md text-slate-700 hover:bg-slate-200/70 disabled:opacity-40 dark:text-slate-200 dark:hover:bg-slate-800"
      >
        <ChevronRight className="h-4 w-4" />
      </button>
    </div>
  );

  return (
    <div
      ref={rootRef}
      className="flex h-full min-h-0 flex-col bg-slate-50 dark:bg-slate-950"
    >
      {/* Hidden input for Insert Image annotation */}
      <input
        ref={imageInputRef}
        type="file"
        accept="image/png,image/jpeg"
        className="hidden"
        onChange={(e) => {
          handleImagePicked(e.target.files?.[0]);
          e.target.value = '';
        }}
      />

      {!isEditing ? (
        <ViewModeToolbar
          fileName={tab.name}
          moduleType="pdf"
          zoomLabel={zoomDisplayLabel}
          onZoomOut={() => {
            setZoomMode('custom');
            setZoom((z) => Math.max(0.5, Number((z - 0.15).toFixed(2))));
          }}
          onZoomIn={() => {
            setZoomMode('custom');
            setZoom((z) => Math.min(3, Number((z + 0.15).toFixed(2))));
          }}
          onZoomReset={() =>
            setZoomMode((m) => (m === 'fit-width' ? 'fit-page' : 'fit-width'))
          }
          isFavorite={Boolean(fileRecord?.isFavorite)}
          onToggleFavorite={onToggleFavorite}
          onSearch={() => setSearchOpen((s) => !s)}
          searchActive={searchOpen}
          onPrint={() => void handlePrint()}
          onExport={onExportDevice}
          exportLabel="Download PDF"
          onEdit={onEnterEdit}
          leadingSlot={
            <button
              type="button"
              onClick={() => setSidebarOpen((o) => !o)}
              aria-label="Toggle thumbnails and outline sidebar"
              className={`flex h-8 w-8 shrink-0 items-center justify-center rounded-lg border text-slate-700 transition-colors dark:text-slate-200 ${
                sidebarOpen
                  ? 'border-blue-600 bg-blue-50 text-blue-700 dark:border-blue-500 dark:bg-blue-950/50 dark:text-blue-300'
                  : 'border-slate-200 hover:bg-slate-100 dark:border-slate-700 dark:hover:bg-slate-800'
              }`}
            >
              <PanelLeft className="h-4 w-4" />
            </button>
          }
          navigationSlot={pageNavControl}
          extraActionsSlot={
            <>
              <button
                type="button"
                onClick={() => setViewRotation((r) => (r + 90) % 360)}
                aria-label="Rotate view 90 degrees clockwise"
                title="Rotate view"
                className="hidden h-8 w-8 shrink-0 items-center justify-center rounded-lg border border-slate-200 text-slate-700 hover:bg-slate-100 lg:flex dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800"
              >
                <RotateCw className="h-3.5 w-3.5" />
              </button>
              <button
                type="button"
                onClick={() => setDarkReading((d) => !d)}
                aria-label="Toggle PDF dark reading mode"
                title="Dark reading mode"
                className={`hidden h-8 w-8 shrink-0 items-center justify-center rounded-lg border transition-colors lg:flex ${
                  darkReading
                    ? 'border-blue-600 bg-blue-50 text-blue-700 dark:border-blue-500 dark:bg-blue-950/50 dark:text-blue-300'
                    : 'border-slate-200 text-slate-700 hover:bg-slate-100 dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800'
                }`}
              >
                {darkReading ? <Sun className="h-3.5 w-3.5" /> : <Moon className="h-3.5 w-3.5" />}
              </button>
              <button
                type="button"
                onClick={() => void toggleFullscreen()}
                aria-label={isFullscreen ? 'Exit fullscreen' : 'Enter fullscreen'}
                title="Fullscreen"
                className="hidden h-8 w-8 shrink-0 items-center justify-center rounded-lg border border-slate-200 text-slate-700 hover:bg-slate-100 xl:flex dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800"
              >
                {isFullscreen ? (
                  <Minimize2 className="h-3.5 w-3.5" />
                ) : (
                  <Maximize2 className="h-3.5 w-3.5" />
                )}
              </button>
            </>
          }
        />
      ) : (
        <>
          {/* Row 1: Compact Single-Row PDF Viewer & Action Toolbar */}
          <div className="flex shrink-0 items-center justify-between gap-1.5 overflow-x-auto border-b border-slate-200 bg-white px-2.5 py-1.5 dark:border-slate-800 dark:bg-slate-900">
            {/* Left: Sidebar toggle + Page Navigation + Zoom */}
            <div className="flex shrink-0 items-center gap-1">
              <button
                type="button"
                onClick={() => setSidebarOpen((o) => !o)}
                aria-label="Toggle thumbnails and outline sidebar"
                className={`flex h-8 w-8 shrink-0 items-center justify-center rounded-lg border text-slate-700 transition-colors dark:text-slate-200 ${
                  sidebarOpen
                    ? 'border-blue-600 bg-blue-50 text-blue-700 dark:border-blue-500 dark:bg-blue-950/50 dark:text-blue-300'
                    : 'border-slate-200 hover:bg-slate-100 dark:border-slate-700 dark:hover:bg-slate-800'
                }`}
              >
                <PanelLeft className="h-4 w-4" />
              </button>

          <button
            type="button"
            onClick={onToggleFavorite}
            aria-label={fileRecord?.isFavorite ? 'Remove from favourites' : 'Add to favourites'}
            className="flex h-8 w-8 shrink-0 items-center justify-center rounded-lg border border-slate-200 text-slate-600 hover:bg-slate-100 dark:border-slate-700 dark:text-slate-300 dark:hover:bg-slate-800"
          >
            <Star
              className={`h-4 w-4 ${
                fileRecord?.isFavorite ? 'fill-amber-400 text-amber-500' : ''
              }`}
            />
          </button>

          {/* Page Nav */}
          <div className="flex shrink-0 items-center gap-0.5 rounded-lg border border-slate-200 bg-slate-50 px-0.5 py-0.5 dark:border-slate-700 dark:bg-slate-950">
            <button
              type="button"
              disabled={currentPage <= 1}
              onClick={() => jumpToPage(currentPage - 1)}
              aria-label="Previous page"
              className="flex h-7 w-7 items-center justify-center rounded-md text-slate-700 hover:bg-slate-200/70 disabled:opacity-40 dark:text-slate-200 dark:hover:bg-slate-800"
            >
              <ChevronLeft className="h-4 w-4" />
            </button>
            <form
              onSubmit={(e) => {
                e.preventDefault();
                const parsed = Number.parseInt(pageInput, 10);
                if (Number.isFinite(parsed)) jumpToPage(parsed);
              }}
              className="flex items-center gap-1 px-1 text-xs tabular-nums"
            >
              <span className="hidden text-slate-500 sm:inline dark:text-slate-400">Page</span>
              <input
                type="text"
                inputMode="numeric"
                aria-label="Current page number"
                value={pageInput}
                onChange={(e) => setPageInput(e.target.value)}
                className="h-6 w-8 rounded border border-slate-200 bg-white text-center text-xs font-semibold text-slate-800 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-100"
              />
              <span className="text-slate-500 dark:text-slate-400">of {numPages || 1}</span>
            </form>
            <button
              type="button"
              disabled={currentPage >= numPages}
              onClick={() => jumpToPage(currentPage + 1)}
              aria-label="Next page"
              className="flex h-7 w-7 items-center justify-center rounded-md text-slate-700 hover:bg-slate-200/70 disabled:opacity-40 dark:text-slate-200 dark:hover:bg-slate-800"
            >
              <ChevronRight className="h-4 w-4" />
            </button>
          </div>

          {/* Zoom & Fit Controls */}
          <div className="flex shrink-0 items-center gap-0.5 rounded-lg border border-slate-200 bg-slate-50 px-0.5 py-0.5 dark:border-slate-700 dark:bg-slate-950">
            <button
              type="button"
              onClick={() => {
                setZoomMode('custom');
                setZoom((z) => Math.max(0.5, Number((z - 0.15).toFixed(2))));
              }}
              aria-label="Zoom out"
              className="flex h-7 w-7 items-center justify-center rounded-md text-slate-700 hover:bg-slate-200/70 dark:text-slate-200 dark:hover:bg-slate-800"
            >
              <ZoomOut className="h-3.5 w-3.5" />
            </button>
            <button
              type="button"
              onClick={() => setZoomMode((m) => (m === 'fit-width' ? 'fit-page' : 'fit-width'))}
              className="min-h-[28px] whitespace-nowrap rounded-md px-1.5 text-xs font-semibold tabular-nums text-slate-700 hover:bg-slate-200/70 dark:text-slate-200 dark:hover:bg-slate-800"
              title="Toggle Fit Width / Fit Page"
            >
              {zoomMode === 'fit-width'
                ? 'Fit'
                : zoomMode === 'fit-page'
                  ? 'Page'
                  : `${Math.round(zoom * 100)}%`}
            </button>
            <button
              type="button"
              onClick={() => {
                setZoomMode('custom');
                setZoom((z) => Math.min(3, Number((z + 0.15).toFixed(2))));
              }}
              aria-label="Zoom in"
              className="flex h-7 w-7 items-center justify-center rounded-md text-slate-700 hover:bg-slate-200/70 dark:text-slate-200 dark:hover:bg-slate-800"
            >
              <ZoomIn className="h-3.5 w-3.5" />
            </button>
          </div>

          {/* Layout Mode */}
          <div className="hidden shrink-0 items-center gap-0.5 rounded-lg border border-slate-200 bg-slate-50 p-0.5 xl:flex dark:border-slate-700 dark:bg-slate-950">
            {(
              [
                ['single', 'Single page view', FileText],
                ['continuous', 'Continuous scroll view', ScrollText],
                ['two-page', 'Two-page spread view', Columns2],
              ] as const
            ).map(([mode, label, Icon]) => (
              <button
                key={mode}
                type="button"
                onClick={() => setLayoutMode(mode)}
                aria-label={label}
                title={label}
                className={`flex h-7 w-7 items-center justify-center rounded-md ${
                  layoutMode === mode
                    ? 'bg-white text-blue-700 shadow-xs dark:bg-slate-800 dark:text-blue-400'
                    : 'text-slate-600 hover:text-slate-900 dark:text-slate-400 dark:hover:text-slate-100'
                }`}
              >
                <Icon className="h-3.5 w-3.5" />
              </button>
            ))}
          </div>

          {/* Rotate View, Dark Reading, Search, Fullscreen, Print */}
          <button
            type="button"
            onClick={() => setViewRotation((r) => (r + 90) % 360)}
            aria-label="Rotate view 90 degrees clockwise"
            title="Rotate view"
            className="flex h-8 w-8 shrink-0 items-center justify-center rounded-lg border border-slate-200 text-slate-700 hover:bg-slate-100 dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800"
          >
            <RotateCw className="h-3.5 w-3.5" />
          </button>

          <button
            type="button"
            onClick={() => setDarkReading((d) => !d)}
            aria-label="Toggle PDF dark reading mode"
            title="Dark reading mode"
            className={`flex h-8 w-8 shrink-0 items-center justify-center rounded-lg border transition-colors ${
              darkReading
                ? 'border-blue-600 bg-blue-50 text-blue-700 dark:border-blue-500 dark:bg-blue-950/50 dark:text-blue-300'
                : 'border-slate-200 text-slate-700 hover:bg-slate-100 dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800'
            }`}
          >
            {darkReading ? <Sun className="h-3.5 w-3.5" /> : <Moon className="h-3.5 w-3.5" />}
          </button>

          <button
            type="button"
            onClick={() => setSearchOpen((s) => !s)}
            aria-label="Search inside PDF"
            title="Search text (Ctrl+F)"
            className={`flex h-8 w-8 shrink-0 items-center justify-center rounded-lg border transition-colors ${
              searchOpen
                ? 'border-blue-600 bg-blue-50 text-blue-700 dark:border-blue-500 dark:bg-blue-950/50 dark:text-blue-300'
                : 'border-slate-200 text-slate-700 hover:bg-slate-100 dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800'
            }`}
          >
            <Search className="h-3.5 w-3.5" />
          </button>

          <button
            type="button"
            onClick={() => void toggleFullscreen()}
            aria-label={isFullscreen ? 'Exit fullscreen' : 'Enter fullscreen'}
            title="Fullscreen"
            className="hidden h-8 w-8 shrink-0 items-center justify-center rounded-lg border border-slate-200 text-slate-700 hover:bg-slate-100 2xl:flex dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800"
          >
            {isFullscreen ? <Minimize2 className="h-3.5 w-3.5" /> : <Maximize2 className="h-3.5 w-3.5" />}
          </button>

          <button
            type="button"
            onClick={() => void handlePrint()}
            aria-label="Print PDF"
            title="Print PDF"
            className="hidden h-8 w-8 shrink-0 items-center justify-center rounded-lg border border-slate-200 text-slate-700 hover:bg-slate-100 xl:flex dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800"
          >
            <Printer className="h-3.5 w-3.5" />
          </button>
        </div>

        {/* Right: Feature Modals, Undo/Redo, Save, Export */}
        <div className="flex shrink-0 items-center gap-1">
          <button
            type="button"
            onClick={() => setOrganizeOpen(true)}
            className="inline-flex min-h-[32px] shrink-0 items-center gap-1 whitespace-nowrap rounded-lg border border-slate-200 bg-white px-2 text-xs font-semibold text-slate-800 hover:bg-slate-100 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-100 dark:hover:bg-slate-800"
          >
            <Layers className="h-3.5 w-3.5 text-blue-700 dark:text-blue-400" />
            Organize
          </button>

          <button
            type="button"
            onClick={() => setFormsModal({open: true, tab: 'signature'})}
            className="inline-flex min-h-[32px] shrink-0 items-center gap-1 whitespace-nowrap rounded-lg border border-slate-200 bg-white px-2 text-xs font-semibold text-slate-800 hover:bg-slate-100 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-100 dark:hover:bg-slate-800"
          >
            <PenTool className="h-3.5 w-3.5 text-blue-700 dark:text-blue-400" />
            Sign
          </button>

          <button
            type="button"
            onClick={async () => {
              if (!activePassword && (annotations.some((a) => a.tool !== 'redact') || hasActiveOverlays)) {
                await commitSaveToVault(true);
              }
              setSecurityModal({open: true, tab: 'security'});
            }}
            className="inline-flex min-h-[32px] shrink-0 items-center gap-1 whitespace-nowrap rounded-lg border border-slate-200 bg-white px-2 text-xs font-semibold text-slate-800 hover:bg-slate-100 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-100 dark:hover:bg-slate-800"
          >
            <Lock className="h-3.5 w-3.5 text-blue-700 dark:text-blue-400" />
            Protect
          </button>

          <button
            type="button"
            onClick={() => setSecurityModal({open: true, tab: 'convert'})}
            className="inline-flex min-h-[32px] shrink-0 items-center gap-1 whitespace-nowrap rounded-lg border border-slate-200 bg-white px-2 text-xs font-semibold text-slate-800 hover:bg-slate-100 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-100 dark:hover:bg-slate-800"
          >
            <FileImage className="h-3.5 w-3.5 text-blue-700 dark:text-blue-400" />
            Convert
          </button>

          <button
            type="button"
            onClick={() => setSecurityModal({open: true, tab: 'ocr'})}
            className="inline-flex min-h-[32px] shrink-0 items-center gap-1 whitespace-nowrap rounded-lg border border-slate-200 bg-white px-2 text-xs font-semibold text-slate-800 hover:bg-slate-100 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-100 dark:hover:bg-slate-800"
          >
            <ScanText className="h-3.5 w-3.5 text-blue-700 dark:text-blue-400" />
            OCR
          </button>

          <div className="flex shrink-0 items-center gap-0.5">
            <button
              type="button"
              disabled={past.length === 0}
              onClick={handleUndo}
              aria-label="Undo PDF change"
              className="flex h-8 w-7 items-center justify-center rounded-lg border border-slate-200 text-slate-700 hover:bg-slate-100 disabled:opacity-40 dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800"
            >
              <Undo2 className="h-3.5 w-3.5" />
            </button>
            <button
              type="button"
              disabled={future.length === 0}
              onClick={handleRedo}
              aria-label="Redo PDF change"
              className="flex h-8 w-7 items-center justify-center rounded-lg border border-slate-200 text-slate-700 hover:bg-slate-100 disabled:opacity-40 dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800"
            >
              <Redo2 className="h-3.5 w-3.5" />
            </button>
          </div>

          <button
            type="button"
            onClick={() => void commitSaveToVault(false)}
            className="inline-flex min-h-[32px] shrink-0 items-center gap-1 whitespace-nowrap rounded-lg border border-slate-200 bg-white px-2.5 text-xs font-semibold text-slate-800 hover:bg-slate-100 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-100 dark:hover:bg-slate-800"
          >
            <Save className="h-3.5 w-3.5 text-blue-700 dark:text-blue-400" />
            Save
          </button>

          <button
            type="button"
            onClick={onExportDevice}
            aria-label="Export PDF to device"
            title="Export PDF"
            className="flex h-8 w-8 shrink-0 items-center justify-center rounded-lg border border-slate-200 bg-white text-slate-800 hover:bg-slate-100 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-100 dark:hover:bg-slate-800"
          >
            <Download className="h-3.5 w-3.5" />
          </button>

          <button
            type="button"
            onClick={onDoneEditing}
            aria-label="Done editing"
            title="Save and return to View mode (Esc)"
            className="inline-flex min-h-[32px] shrink-0 items-center gap-1 whitespace-nowrap rounded-lg bg-blue-700 px-3 text-xs font-semibold text-white hover:bg-blue-800 dark:bg-blue-600 dark:hover:bg-blue-500"
          >
            <Check className="h-3.5 w-3.5" />
            Done
          </button>
        </div>
      </div>

      {/* Row 2: Compact Single-Row Annotation & Overlay Strip */}
      <div className="flex shrink-0 items-center justify-between gap-2 overflow-x-auto border-b border-slate-200 bg-slate-50 px-2.5 py-1.5 dark:border-slate-800 dark:bg-slate-950">
        <div className="flex shrink-0 items-center gap-0.5">
          {TOOL_ITEMS.map(({id, label, Icon}) => (
            <button
              key={id}
              type="button"
              onClick={() => handleToolSelect(id)}
              className={`inline-flex min-h-[32px] shrink-0 items-center gap-1 whitespace-nowrap rounded-lg px-2 text-xs font-medium transition-colors ${
                activeTool === id
                  ? 'bg-slate-900 text-white dark:bg-blue-600 dark:text-white'
                  : 'text-slate-700 hover:bg-slate-200/70 dark:text-slate-300 dark:hover:bg-slate-800'
              }`}
            >
              <Icon className="h-3.5 w-3.5" />
              <span>{label}</span>
            </button>
          ))}
        </div>

        {/* Contextual Tool Options + Watermark/Header/Footer/Page Numbers */}
        <div className="flex shrink-0 items-center gap-1.5">
          {activeTool === 'stamp' && (
            <select
              aria-label="Stamp text preset"
              value={pendingStampText}
              onChange={(e) => setPendingStampText(e.target.value)}
              className="h-8 rounded-lg border border-slate-200 bg-white px-2 text-xs font-semibold text-slate-800 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-100"
            >
              {STAMP_PRESETS.map((s) => (
                <option key={s} value={s}>
                  {s}
                </option>
              ))}
            </select>
          )}

          {activeTool !== 'select' && activeTool !== 'redact' && (
            <div className="flex items-center gap-1">
              <label className="sr-only" htmlFor="pdf-tool-color">
                Annotation color
              </label>
              <input
                id="pdf-tool-color"
                type="color"
                value={toolColor}
                onChange={(e) => setToolColor(e.target.value)}
                className="h-7 w-7 cursor-pointer rounded border border-slate-300 bg-transparent p-0.5 dark:border-slate-700"
              />
              <select
                aria-label="Stroke width"
                value={toolStrokeWidth}
                onChange={(e) => setToolStrokeWidth(Number(e.target.value))}
                className="h-8 rounded-lg border border-slate-200 bg-white px-1.5 text-xs tabular-nums dark:border-slate-700 dark:bg-slate-900"
              >
                <option value={2}>2px</option>
                <option value={4}>4px</option>
                <option value={6}>6px</option>
              </select>
            </div>
          )}

          {pendingRedactions.length > 0 && (
            <button
              type="button"
              onClick={() => setSecurityModal({open: true, tab: 'security'})}
              className="inline-flex min-h-[32px] shrink-0 items-center gap-1.5 whitespace-nowrap rounded-lg bg-rose-700 px-2.5 text-xs font-semibold text-white hover:bg-rose-800"
            >
              Commit {pendingRedactions.length} Redaction{pendingRedactions.length > 1 ? 's' : ''}
            </button>
          )}

          <div className="flex shrink-0 items-center gap-1">
            <button
              type="button"
              data-testid="pdf-btn-page-numbers"
              onClick={() => setOverlaysModal({open: true, tab: 'page-numbers'})}
              className="inline-flex min-h-[32px] shrink-0 items-center gap-1 whitespace-nowrap rounded-lg border border-slate-200 bg-white px-2 text-xs font-semibold text-slate-700 hover:bg-slate-100 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-200 dark:hover:bg-slate-800"
              title="Add Page Numbers with safe-margin live preview"
            >
              <Hash className="h-3.5 w-3.5 text-blue-700 dark:text-blue-400" />
              Page #
            </button>
            <button
              type="button"
              data-testid="pdf-btn-header-footer"
              onClick={() => setOverlaysModal({open: true, tab: 'header-footer'})}
              className="inline-flex min-h-[32px] shrink-0 items-center gap-1 whitespace-nowrap rounded-lg border border-slate-200 bg-white px-2 text-xs font-semibold text-slate-700 hover:bg-slate-100 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-200 dark:hover:bg-slate-800"
              title="Add Header & Footer with safe-margin live preview"
            >
              <FileText className="h-3.5 w-3.5 text-blue-700 dark:text-blue-400" />
              Header/Footer
            </button>
            <button
              type="button"
              data-testid="pdf-btn-watermark"
              onClick={() => setOverlaysModal({open: true, tab: 'watermark'})}
              className="inline-flex min-h-[32px] shrink-0 items-center gap-1 whitespace-nowrap rounded-lg border border-slate-200 bg-white px-2 text-xs font-semibold text-slate-700 hover:bg-slate-100 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-200 dark:hover:bg-slate-800"
              title="Add Watermark with position presets and live preview"
            >
              <Wand2 className="h-3.5 w-3.5 text-blue-700 dark:text-blue-400" />
              Watermark
            </button>
          </div>

          {pdfBytes && (
            <span className="hidden shrink-0 whitespace-nowrap text-[11px] font-medium tabular-nums text-slate-500 2xl:inline dark:text-slate-400">
              {formatBytes(pdfBytes.byteLength)}
              {annotations.length > 0 ? ` · ${annotations.length} draft` : ''}
            </span>
          )}
        </div>
      </div>
        </>
      )}

      {/* Inline Search Bar */}
      {searchOpen && (
        <form
          onSubmit={(e) => void runSearch(e)}
          className="flex flex-wrap items-center gap-2 border-b border-slate-200 bg-white px-4 py-2 dark:border-slate-800 dark:bg-slate-900"
        >
          <Search className="h-4 w-4 text-slate-400" />
          <input
            type="text"
            autoFocus
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            placeholder="Find text across all pages..."
            className="h-9 min-w-[200px] flex-1 rounded-lg border border-slate-200 bg-slate-50 px-3 text-xs dark:border-slate-700 dark:bg-slate-950"
          />
          <button
            type="submit"
            disabled={searching}
            className="min-h-[36px] rounded-lg bg-blue-700 px-3.5 text-xs font-semibold text-white hover:bg-blue-800"
          >
            {searching ? 'Searching...' : 'Find'}
          </button>
          {searchMatches.length > 0 && (
            <div className="flex items-center gap-1.5 text-xs tabular-nums text-slate-600 dark:text-slate-300">
              <span>
                {searchIdx + 1} / {searchMatches.length} pages
              </span>
              <button
                type="button"
                onClick={() => cycleSearchMatch(-1)}
                aria-label="Previous search match"
                className="flex h-8 w-8 items-center justify-center rounded border border-slate-200 hover:bg-slate-100 dark:border-slate-700 dark:hover:bg-slate-800"
              >
                <ChevronLeft className="h-3.5 w-3.5" />
              </button>
              <button
                type="button"
                onClick={() => cycleSearchMatch(1)}
                aria-label="Next search match"
                className="flex h-8 w-8 items-center justify-center rounded border border-slate-200 hover:bg-slate-100 dark:border-slate-700 dark:hover:bg-slate-800"
              >
                <ChevronRight className="h-3.5 w-3.5" />
              </button>
            </div>
          )}
          <button
            type="button"
            onClick={() => {
              setSearchOpen(false);
              setSearchMatches([]);
            }}
            aria-label="Close search bar"
            className="flex h-8 w-8 items-center justify-center rounded-lg text-slate-500 hover:bg-slate-100 dark:hover:bg-slate-800"
          >
            <X className="h-4 w-4" />
          </button>
        </form>
      )}



      {/* Main Split Workspace: Left Sidebar + Viewer Canvas */}
      <div className="relative flex min-h-0 flex-1 overflow-hidden">
        {/* Left Sidebar: Thumbnails & Outline (overlay on <lg viewports so mobile canvas stays full width) */}
        {sidebarOpen && pdfDoc && (
          <aside className="absolute inset-y-0 left-0 z-20 flex w-48 shrink-0 flex-col border-r border-slate-200 bg-white shadow-xl lg:static lg:shadow-none dark:border-slate-800 dark:bg-slate-900">
            <div className="flex items-center gap-1 border-b border-slate-200 p-1.5 dark:border-slate-800">
              <button
                type="button"
                onClick={() => setSidebarTab('thumbnails')}
                className={`flex-1 rounded-lg py-1.5 text-xs font-semibold ${
                  sidebarTab === 'thumbnails'
                    ? 'bg-slate-900 text-white dark:bg-slate-100 dark:text-slate-900'
                    : 'text-slate-600 hover:bg-slate-100 dark:text-slate-400 dark:hover:bg-slate-800'
                }`}
              >
                Pages
              </button>
              <button
                type="button"
                onClick={() => setSidebarTab('outline')}
                className={`flex-1 rounded-lg py-1.5 text-xs font-semibold ${
                  sidebarTab === 'outline'
                    ? 'bg-slate-900 text-white dark:bg-slate-100 dark:text-slate-900'
                    : 'text-slate-600 hover:bg-slate-100 dark:text-slate-400 dark:hover:bg-slate-800'
                }`}
              >
                Outline
              </button>
              <button
                type="button"
                onClick={() => setSidebarOpen(false)}
                aria-label="Close pages sidebar"
                className="flex h-7 w-7 items-center justify-center rounded-md text-slate-500 hover:bg-slate-100 lg:hidden dark:hover:bg-slate-800"
              >
                <X className="h-3.5 w-3.5" />
              </button>
            </div>

            <div className="flex-1 overflow-y-auto p-2.5">
              {sidebarTab === 'thumbnails' ? (
                <div className="space-y-2.5">
                  {Array.from({length: numPages}, (_, i) => i + 1).map((p) => (
                    <SidebarThumbnail
                      key={p}
                      pdfDoc={pdfDoc}
                      pageNumber={p}
                      active={currentPage === p}
                      onClick={() => jumpToPage(p)}
                    />
                  ))}
                </div>
              ) : outline.length === 0 ? (
                <div className="p-3 text-center text-xs text-slate-500 dark:text-slate-400">
                  <BookOpen className="mx-auto mb-2 h-5 w-5 text-slate-400" />
                  This PDF has no embedded bookmark outline.
                </div>
              ) : (
                <ul className="space-y-1">
                  {outline.map((item, idx) => (
                    <li key={`${item.title}_${idx}`}>
                      <button
                        type="button"
                        disabled={!item.pageNumber}
                        onClick={() => item.pageNumber && jumpToPage(item.pageNumber)}
                        className="flex w-full items-center justify-between gap-2 rounded-lg px-2.5 py-2 text-left text-xs text-slate-700 hover:bg-slate-100 disabled:opacity-50 dark:text-slate-300 dark:hover:bg-slate-800"
                      >
                        <span className="truncate">{item.title}</span>
                        {item.pageNumber && (
                          <span className="shrink-0 text-[11px] tabular-nums text-slate-400">
                            {item.pageNumber}
                          </span>
                        )}
                      </button>
                    </li>
                  ))}
                </ul>
              )}
            </div>
          </aside>
        )}

        {/* Center Canvas / States */}
        {passwordNeeded ? (
          <div className="flex flex-1 items-center justify-center p-6">
            <form
              onSubmit={(e) => {
                e.preventDefault();
                setActivePassword(passwordInput);
              }}
              className="w-full max-w-sm rounded-2xl border border-slate-200 bg-white p-6 shadow-sm dark:border-slate-800 dark:bg-slate-900"
            >
              <Lock className="h-6 w-6 text-blue-700 dark:text-blue-400" />
              <h2 className="mt-3 text-base font-bold text-slate-900 dark:text-slate-100">
                Password Protected PDF
              </h2>
              <p className="mt-1 text-xs text-slate-500 dark:text-slate-400">
                Enter the password to unlock <strong>{tab.name}</strong> locally on this device.
              </p>
              <input
                type="password"
                autoFocus
                value={passwordInput}
                onChange={(e) => setPasswordInput(e.target.value)}
                placeholder="PDF password..."
                className="mt-4 h-11 w-full rounded-xl border border-slate-200 bg-slate-50 px-3.5 text-sm dark:border-slate-700 dark:bg-slate-950"
              />
              <button
                type="submit"
                className="mt-3 min-h-[42px] w-full rounded-xl bg-blue-700 text-xs font-semibold text-white hover:bg-blue-800"
              >
                Unlock Document
              </button>
            </form>
          </div>
        ) : loadError ? (
          <div className="flex flex-1 items-center justify-center p-6">
            <div className="max-w-md rounded-2xl border border-rose-200 bg-white p-6 text-center dark:border-rose-900 dark:bg-slate-900">
              <p className="text-sm font-bold text-rose-700 dark:text-rose-400">
                Could not open PDF
              </p>
              <p className="mt-1.5 text-xs text-slate-600 dark:text-slate-300">{loadError}</p>
            </div>
          </div>
        ) : loading || !pdfDoc ? (
          <div className="flex flex-1 flex-col items-center justify-center gap-3 p-8">
            <div className="h-10 w-64 animate-pulse rounded-xl bg-slate-200 dark:bg-slate-800" />
            <div className="h-[460px] w-full max-w-lg animate-pulse rounded-xl bg-slate-200/80 dark:bg-slate-800/60" />
          </div>
        ) : (
          <PdfViewerCanvas
            pdfDoc={pdfDoc}
            numPages={numPages}
            currentPage={currentPage}
            scrollSignal={scrollSignal}
            onPageChange={setCurrentPage}
            layoutMode={layoutMode}
            zoomMode={zoomMode}
            zoom={zoom}
            viewRotation={viewRotation}
            darkReading={darkReading}
            readOnly={!isEditing}
            activeTool={isEditing ? activeTool : 'select'}
            toolColor={toolColor}
            toolOpacity={toolOpacity}
            toolStrokeWidth={toolStrokeWidth}
            pendingStampText={pendingStampText}
            pendingImageDataUrl={pendingImageDataUrl}
            onConsumePendingImage={() => setPendingImageDataUrl(null)}
            annotations={annotations}
            selectedAnnotationId={isEditing ? selectedAnnotationId : null}
            onSelectAnnotation={isEditing ? setSelectedAnnotationId : () => undefined}
            onAddAnnotation={handleAddAnnotation}
            onDeleteAnnotation={handleDeleteAnnotation}
            watermark={watermark}
            headerFooter={headerFooter}
            pageNumbers={pageNumbers}
            searchHighlightPages={searchHighlightSet}
            activeSearchPage={activeSearchPage}
          />
        )}
      </div>

      {!isEditing && (
        <ViewModeMobileBar
          moduleType="pdf"
          onEdit={onEnterEdit}
          onSearch={() => setSearchOpen((s) => !s)}
          onExport={onExportDevice}
        />
      )}

      {/* Organize Modal */}
      {isEditing && pdfBytes && (
        <PdfOrganizeModal
          open={organizeOpen}
          onClose={() => setOrganizeOpen(false)}
          fileName={tab.name}
          pdfBytes={pdfBytes}
          pdfDoc={pdfDoc}
          vaultPdfFiles={vaultPdfFiles}
          onApplyBytes={(next, summary) => handleApplyNewBytes(next, summary)}
          onCreateNewPdf={async (name, bytes) => {
            await createFileFromBytes(name, 'pdf', bytes, fileRecord?.folderId ?? null, true);
          }}
          onToast={pushToast}
        />
      )}

      {/* Forms & Signature Modal */}
      {isEditing && pdfBytes && (
        <PdfFormsAndSignModal
          open={formsModal.open}
          initialTab={formsModal.tab}
          onClose={() => setFormsModal((m) => ({...m, open: false}))}
          pdfBytes={pdfBytes}
          onApplyBytes={(next, summary) => handleApplyNewBytes(next, summary)}
          onReadySignature={handleReadySignature}
          onToast={pushToast}
        />
      )}

      {/* Security, Convert & OCR Modal */}
      {isEditing && pdfBytes && (
        <PdfSecurityAndConvertModal
          open={securityModal.open}
          initialTab={securityModal.tab}
          onClose={() => setSecurityModal((m) => ({...m, open: false}))}
          fileName={tab.name}
          numPages={numPages}
          pdfBytes={pdfBytes}
          currentPassword={activePassword}
          pendingRedactions={pendingRedactions}
          onApplyBytes={(next, summary, nextPw) => handleApplyNewBytes(next, summary, nextPw)}
          onClearRedactions={() =>
            setAnnotations((prev) => prev.filter((a) => a.tool !== 'redact'))
          }
          onCreateNewPdf={async (name, bytes) => {
            await createFileFromBytes(name, 'pdf', bytes, fileRecord?.folderId ?? null, true);
          }}
          onToast={pushToast}
        />
      )}

      {/* Page Numbers, Header/Footer & Watermark Studio with Live Preview */}
      {isEditing && pdfBytes && (
        <PdfOverlaysModal
          open={overlaysModal.open}
          initialTab={overlaysModal.tab}
          onClose={() => setOverlaysModal((m) => ({...m, open: false}))}
          fileName={tab.name}
          pdfBytes={pdfBytes}
          pdfDoc={pdfDoc}
          numPages={numPages}
          initialPage={currentPage}
          activePassword={activePassword}
          initialWatermark={watermark}
          initialHeaderFooter={headerFooter}
          initialPageNumbers={pageNumbers}
          onApplyBytes={(next, summary) => handleApplyNewBytes(next, summary)}
          onToast={pushToast}
        />
      )}
    </div>
  );
};

export default PdfModule;
