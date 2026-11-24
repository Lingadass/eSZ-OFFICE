import React, {useCallback, useEffect, useMemo, useRef, useState} from 'react';
import {EditorContent, useEditor} from '@tiptap/react';
import StarterKit from '@tiptap/starter-kit';
import {
  Color,
  FontSize,
  LineHeight,
  TextStyle,
} from '@tiptap/extension-text-style';
import Highlight from '@tiptap/extension-highlight';
import Subscript from '@tiptap/extension-subscript';
import Superscript from '@tiptap/extension-superscript';
import TextAlign from '@tiptap/extension-text-align';
import {TableRow} from '@tiptap/extension-table';
import Placeholder from '@tiptap/extension-placeholder';
import {AlertTriangle, ChevronLeft, ChevronRight, FileText, Globe, Info, Trash2} from 'lucide-react';
import type {DocVersionRecord} from '../../core/types';
import {
  deleteDocVersion,
  getDocMeta,
  listDocVersions,
  loadFileBlob,
  saveDocMeta,
  saveDocVersion,
} from '../../core/storage/fileRepository';
import {saveBlobToDevice} from '../../core/file-io/fileSystemAccess';
import {useWorkspaceStore} from '../../core/workspace/useWorkspaceStore';
import type {OfficeModuleProps} from '../pdf/PdfModule';
import {
  BookmarkMark,
  DocCaptionNode,
  DocEquationNode,
  DocFontFamily,
  DocImage,
  DocLockedElementNode,
  DocOoxmlPreserveExtension,
  DocParagraphStyleExtension,
  DocRunStyleExtension,
  DocShapeNode,
  DocTable,
  DocTableCell,
  DocTableHeader,
  DocTabNode,
  DocTextBoxNode,
  DocTocNode,
  IndentExtension,
  NoteRefNode,
  PageBreakNode,
} from './docsExtensions';
import {resolveRenderFontStack} from './docFonts';
import {loadDocumentBlobAsHtml, serializeDocumentToBlob} from './docsFileIo';
import type {DocxImportWarning} from './docxImport';
import {DocxPreviewView, type DocxPreviewMetrics} from './DocxPreviewView';
import {
  computeVisiblePageNumber,
  extractDocxDefaultTypography,
  isRealDocxArchive,
  type DocxDocDefaults,
} from './docxPreviewUtils';
import {
  computePaginationPlan,
  DOC_BOTTOM_SAFETY_PX,
  DOC_PAGE_GAP_PX,
  formatPageNumber,
  getPageMetrics,
  shouldShowPageChrome,
  type PaginationPlan,
} from './docsPagination';
import {
  measureAndPaginateEditor,
  PaginationExtension,
  type PaginationRuntimeConfig,
} from './docsPaginationPlugin';
import {
  exportDocxPreviewDomToPdf,
  exportPaginatedEditorDomToPdf,
} from './domPdfExport';
import {buildEmbeddedFontFaceCss} from './pdfFonts';
import {
  buildTocHtml,
  computeDocumentStats,
  escapeHtml,
  extractHeadingsFromHtml,
  htmlToPdfBytes,
  htmlToPlainText,
} from './docsPure';
import {
  DEFAULT_DOC_LAYOUT,
  DOC_PAGE_DIMENSIONS_PX,
  DOC_THEMES,
  type DocComment,
  type DocNoteItem,
  type DocPageBorderConfig,
  type DocPageBorderSide,
  type DocPageLayout,
  type DocPersistedMeta,
  type DocsRibbonTab,
  type DocViewMode,
  type FormatPainterMarks,
} from './docsTypes';
import {DocsRibbon} from './DocsRibbon';
import {DocsCommentsSidebar, DocsDialogs, type DocsDialogKind} from './DocsDialogs';
import {ViewModeMobileBar, ViewModeToolbar} from '../../shell/ViewModeToolbar';

function borderSideToCss(side?: DocPageBorderSide): string | undefined {
  if (!side || side.val === 'none' || side.val === 'nil') return undefined;
  const widthPx = Math.max(0.75, side.szPt * (96 / 72));
  const v = side.val.toLowerCase();
  const style =
    v === 'double' || v === 'triple'
      ? 'double'
      : v === 'dotted'
        ? 'dotted'
        : v.includes('dash')
          ? 'dashed'
          : 'solid';
  const finalWidth = style === 'double' ? Math.max(3, widthPx * 2.2) : widthPx;
  return `${finalWidth.toFixed(2)}px ${style} ${side.color || '#000000'}`;
}

function buildPageBorderOverlayStyle(
  config: DocPageBorderConfig | null | undefined,
  pageIdx: number,
  metrics: {marginTop: number; marginRight: number; marginBottom: number; marginLeft: number},
): React.CSSProperties | null {
  if (!config) return null;
  if (config.display === 'firstPage' && pageIdx !== 0) return null;
  if (config.display === 'notFirstPage' && pageIdx === 0) return null;

  const borderTop = borderSideToCss(config.top);
  const borderRight = borderSideToCss(config.right);
  const borderBottom = borderSideToCss(config.bottom);
  const borderLeft = borderSideToCss(config.left);
  if (!borderTop && !borderRight && !borderBottom && !borderLeft) return null;

  const ptToPx = (pt: number) => Math.round(pt * (96 / 72));
  const topSpace = ptToPx(config.top?.spacePt ?? 24);
  const rightSpace = ptToPx(config.right?.spacePt ?? 24);
  const bottomSpace = ptToPx(config.bottom?.spacePt ?? 24);
  const leftSpace = ptToPx(config.left?.spacePt ?? 24);

  const isFromText = config.offsetFrom === 'text';
  const top = isFromText ? Math.max(6, metrics.marginTop - topSpace) : Math.max(6, topSpace);
  const right = isFromText ? Math.max(6, metrics.marginRight - rightSpace) : Math.max(6, rightSpace);
  const bottom = isFromText
    ? Math.max(6, metrics.marginBottom - bottomSpace)
    : Math.max(6, bottomSpace);
  const left = isFromText ? Math.max(6, metrics.marginLeft - leftSpace) : Math.max(6, leftSpace);

  return {
    top: `${top}px`,
    right: `${right}px`,
    bottom: `${bottom}px`,
    left: `${left}px`,
    borderTop,
    borderRight,
    borderBottom,
    borderLeft,
  };
}

const HorizontalRuler: React.FC<{
  pageWidthPx: number;
  marginLeftPx: number;
  marginRightPx: number;
  zoom: number;
}> = ({pageWidthPx, marginLeftPx, marginRightPx, zoom}) => {
  const scaledW = Math.round(pageWidthPx * zoom);
  const scaledLeft = Math.round(marginLeftPx * zoom);
  const scaledRight = Math.round(marginRightPx * zoom);
  const totalInches = Math.ceil(pageWidthPx / 96);
  const ticks: number[] = [];
  for (let i = 0; i <= totalInches * 4; i++) {
    ticks.push(i);
  }

  return (
    <div
      aria-label="Document horizontal ruler"
      className="relative mx-auto mb-2 h-6 shrink-0 select-none overflow-hidden rounded border border-slate-300 bg-white text-[9px] tabular-nums text-slate-500 shadow-xs dark:border-slate-700 dark:bg-slate-900 dark:text-slate-400"
      style={{width: `${scaledW}px`}}
    >
      {/* Left & right margin shading */}
      <div
        className="pointer-events-none absolute inset-y-0 left-0 bg-slate-200/80 dark:bg-slate-800/90"
        style={{width: `${scaledLeft}px`}}
      />
      <div
        className="pointer-events-none absolute inset-y-0 right-0 bg-slate-200/80 dark:bg-slate-800/90"
        style={{width: `${scaledRight}px`}}
      />
      {/* Inch ticks */}
      {ticks.map((quarter) => {
        const px = Math.round((quarter * 24) * zoom);
        if (px > scaledW) return null;
        const isWhole = quarter % 4 === 0;
        const isHalf = quarter % 2 === 0 && !isWhole;
        return (
          <div
            key={quarter}
            className="pointer-events-none absolute bottom-0 flex flex-col items-center"
            style={{left: `${px}px`}}
          >
            {isWhole && quarter > 0 && (
              <span className="-translate-x-1/2 leading-none">{quarter / 4}</span>
            )}
            <div
              className={`w-px bg-slate-400 dark:bg-slate-500 ${
                isWhole ? 'h-2.5' : isHalf ? 'h-2' : 'h-1'
              }`}
            />
          </div>
        );
      })}
    </div>
  );
};

const DocsModule: React.FC<OfficeModuleProps> = ({
  tab,
  fileRecord,
  mode,
  onEnterEdit,
  onDoneEditing,
  onToggleFavorite,
}) => {
  const {
    markTabDirty,
    registerBinarySaveHandler,
    saveTabBlob,
    createFileFromBytes,
    discardTabChanges,
    pushToast,
  } = useWorkspaceStore();

  const isEditing = mode === 'edit';
  const modeRef = useRef(mode);
  modeRef.current = mode;

  const [activeTab, setActiveTab] = useState<DocsRibbonTab>('home');
  const [layout, setLayout] = useState<DocPageLayout>(DEFAULT_DOC_LAYOUT);
  const [comments, setComments] = useState<DocComment[]>([]);
  const [notes, setNotes] = useState<DocNoteItem[]>([]);
  const [versions, setVersions] = useState<DocVersionRecord[]>([]);
  const [viewMode, setViewMode] = useState<DocViewMode>('print');
  const [zoom, setZoom] = useState<number>(() =>
    typeof window !== 'undefined' && window.innerWidth < 768 ? 0.45 : 1,
  );
  const [showRuler, setShowRuler] = useState(true);
  const [showFormattingMarks, setShowFormattingMarks] = useState(false);
  const [spellCheck, setSpellCheck] = useState(true);
  const [openDialog, setOpenDialog] = useState<DocsDialogKind | null>(null);
  const [commentsOpen, setCommentsOpen] = useState(false);
  const [commentComposerOpen, setCommentComposerOpen] = useState(false);
  const [draftCommentQuote, setDraftCommentQuote] = useState('');
  const [formatPainterMarks, setFormatPainterMarks] = useState<FormatPainterMarks | null>(null);
  const [loadingDoc, setLoadingDoc] = useState(true);
  const [rawDocxBuffer, setRawDocxBuffer] = useState<ArrayBuffer | null>(null);
  const [docxPreviewFailed, setDocxPreviewFailed] = useState(false);
  const [docxPreviewMetrics, setDocxPreviewMetrics] = useState<DocxPreviewMetrics | null>(null);
  const [importWarnings, setImportWarnings] = useState<DocxImportWarning[]>([]);
  const [hasLockedElements, setHasLockedElements] = useState(false);
  const [hasLossyElements, setHasLossyElements] = useState(false);
  const [allowOverwriteOriginal, setAllowOverwriteOriginal] = useState(false);
  const [pendingLossyAction, setPendingLossyAction] = useState<'save' | 'done' | null>(null);
  const [htmlSnapshot, setHtmlSnapshot] = useState('<p></p>');
  const [paginationPlan, setPaginationPlan] = useState<PaginationPlan>(() =>
    computePaginationPlan([], getPageMetrics(DEFAULT_DOC_LAYOUT), new Map()),
  );
  const [currentPage, setCurrentPage] = useState(1);

  const imageInputRef = useRef<HTMLInputElement>(null);
  const sheetRef = useRef<HTMLDivElement>(null);
  const previewRootRef = useRef<HTMLDivElement>(null);
  const scrollContainerRef = useRef<HTMLDivElement>(null);
  const isHydratingRef = useRef(true);
  const layoutRef = useRef(layout);
  const commentsRef = useRef(comments);
  const notesRef = useRef(notes);
  const zoomRef = useRef(zoom);
  zoomRef.current = zoom;
  const hasLossyElementsRef = useRef(hasLossyElements);
  hasLossyElementsRef.current = hasLossyElements;
  const allowOverwriteRef = useRef(allowOverwriteOriginal);
  allowOverwriteRef.current = allowOverwriteOriginal;
  const rawDocxBufferRef = useRef<ArrayBuffer | null>(rawDocxBuffer);
  rawDocxBufferRef.current = rawDocxBuffer;

  const useDocxPreview =
    !isEditing && viewMode === 'print' && rawDocxBuffer !== null && !docxPreviewFailed;

  const handlePlanChange = useCallback((nextPlan: PaginationPlan) => {
    setPaginationPlan((prev) => (prev.signature === nextPlan.signature ? prev : nextPlan));
  }, []);

  const paginationConfigRef = useRef<PaginationRuntimeConfig>({
    enabled: false,
    layout: DEFAULT_DOC_LAYOUT,
    notes: [],
    zoom: 1,
    viewMode: 'print',
    onPlanChange: handlePlanChange,
  });
  paginationConfigRef.current = {
    enabled: !loadingDoc && viewMode === 'print' && !useDocxPreview,
    layout,
    notes,
    zoom,
    viewMode,
    onPlanChange: handlePlanChange,
  };

  useEffect(() => {
    layoutRef.current = layout;
  }, [layout]);
  useEffect(() => {
    commentsRef.current = comments;
  }, [comments]);
  useEffect(() => {
    notesRef.current = notes;
  }, [notes]);

  const editor = useEditor({
    extensions: [
      StarterKit.configure({
        link: {
          openOnClick: false,
          autolink: true,
        },
        underline: {},
      }),
      TextStyle,
      Color,
      DocFontFamily,
      DocRunStyleExtension,
      DocParagraphStyleExtension,
      DocOoxmlPreserveExtension,
      FontSize,
      LineHeight,
      Highlight.configure({multicolor: true}),
      Subscript,
      Superscript,
      TextAlign.configure({types: ['heading', 'paragraph']}),
      DocImage.configure({allowBase64: true}),
      DocTable.configure({resizable: true}),
      TableRow,
      DocTableHeader,
      DocTableCell,
      Placeholder.configure({
        placeholder: 'Start writing your document...',
      }),
      IndentExtension,
      PageBreakNode,
      DocTabNode,
      DocTextBoxNode,
      DocLockedElementNode,
      DocEquationNode,
      DocShapeNode,
      DocCaptionNode,
      DocTocNode,
      BookmarkMark,
      NoteRefNode,
      PaginationExtension.configure({
        getConfig: () => paginationConfigRef.current,
      }),
    ],
    content: '<p></p>',
    editable: mode === 'edit',
    editorProps: {
      attributes: {
        'aria-label': 'Document editor canvas',
        spellcheck: 'true',
      },
    },
    onUpdate: ({editor: ed}) => {
      const html = ed.getHTML();
      setHtmlSnapshot(html);
      if (!isHydratingRef.current && modeRef.current === 'edit') {
        markTabDirty(tab.fileId, true);
      }
    },
    onSelectionUpdate: ({editor: ed}) => {
      const sheetEl = sheetRef.current;
      if (!sheetEl || paginationConfigRef.current.viewMode !== 'print') return;
      try {
        const coords = ed.view.coordsAtPos(ed.state.selection.head);
        const sheetRect = sheetEl.getBoundingClientRect();
        const liveScale =
          sheetEl.offsetWidth > 0 && sheetRect.width > 0
            ? sheetRect.width / sheetEl.offsetWidth
            : Math.max(0.25, zoomRef.current);
        const unscaledY = (coords.top - sheetRect.top) / liveScale;
        const metrics = getPageMetrics(layoutRef.current);
        const stride = metrics.pageHeight + DOC_PAGE_GAP_PX;
        const pageIdx = Math.max(0, Math.floor(unscaledY / stride));
        setCurrentPage(pageIdx + 1);
      } catch {
        // Ignore transient selection coord failures during mount
      }
    },
  });

  // Enforce real TipTap read-only mode and clear edit-only UI state on mode change
  useEffect(() => {
    if (editor && editor.isEditable !== isEditing) {
      editor.setEditable(isEditing, false);
    }
    if (!isEditing) {
      setFormatPainterMarks(null);
      setCommentComposerOpen(false);
      setOpenDialog((prev) =>
        prev === 'find' || prev === 'export' || prev === 'word-count' ? prev : null,
      );
    }
  }, [editor, isEditing]);

  // Keep spellcheck attribute synchronized once EditorContent mounts
  useEffect(() => {
    if (!editor || editor.isDestroyed || loadingDoc) return;
    try {
      const dom = editor.view?.dom;
      if (dom) {
        dom.setAttribute('spellcheck', spellCheck ? 'true' : 'false');
      }
    } catch {
      // TipTap v3 throws if view is accessed before EditorContent finishes mounting
    }
  }, [editor, spellCheck, loadingDoc]);

  // Hydrate document blob, metadata, and versions on fileId change
  useEffect(() => {
    if (!editor) return;
    let cancelled = false;
    isHydratingRef.current = true;
    setLoadingDoc(true);

    (async () => {
      try {
        const [blob, meta, existingVersions] = await Promise.all([
          loadFileBlob(tab.fileId),
          getDocMeta<DocPersistedMeta>(tab.fileId),
          listDocVersions(tab.fileId),
        ]);
        if (cancelled) return;

        let loadedHtml = '<p></p>';
        let detectedDocxBuffer: ArrayBuffer | null = null;
        let detectedDocDefaults: DocxDocDefaults | null = null;
        let importedLayoutPatch: Partial<DocPageLayout> = {};
        let loadedDetailedWarnings: DocxImportWarning[] = [];
        let loadedHasLocked = false;
        let loadedHasLossy = false;

        if (blob) {
          if (tab.extension.toLowerCase() === '.docx') {
            const rawBuf = await blob.arrayBuffer();
            if (await isRealDocxArchive(rawBuf)) {
              detectedDocxBuffer = rawBuf;
              detectedDocDefaults = await extractDocxDefaultTypography(rawBuf);
            }
          }
          const parsed = await loadDocumentBlobAsHtml(blob, tab.extension);
          loadedHtml = parsed.html;
          importedLayoutPatch = parsed.layoutPatch ?? {};
          loadedDetailedWarnings = parsed.detailedWarnings;
          loadedHasLocked = parsed.hasLockedElements;
          loadedHasLossy = parsed.hasLossyElements;
          if (parsed.warnings.length > 0 && !detectedDocxBuffer) {
            pushToast(
              `Imported ${tab.name} with ${parsed.warnings.length} formatting note(s).`,
              'info',
            );
          }
        }

        if (cancelled) return;
        setRawDocxBuffer(detectedDocxBuffer);
        rawDocxBufferRef.current = detectedDocxBuffer;
        setDocxPreviewFailed(false);
        setDocxPreviewMetrics(null);
        setImportWarnings(loadedDetailedWarnings);
        setHasLockedElements(loadedHasLocked);
        setHasLossyElements(loadedHasLossy);
        hasLossyElementsRef.current = loadedHasLossy;
        setAllowOverwriteOriginal(false);
        allowOverwriteRef.current = false;
        setPendingLossyAction(null);
        setCurrentPage(1);
        editor.commands.setContent(loadedHtml, {emitUpdate: false});
        editor.setEditable(modeRef.current === 'edit', false);
        setHtmlSnapshot(editor.getHTML());

        const docxDefaultsPatch: Partial<DocPageLayout> = detectedDocDefaults
          ? {
              ...(detectedDocDefaults.fontFamily
                ? {docDefaultFont: detectedDocDefaults.fontFamily}
                : {}),
              ...(detectedDocDefaults.fontSizePt
                ? {docDefaultFontSizePt: detectedDocDefaults.fontSizePt}
                : {}),
              ...(detectedDocDefaults.color
                ? {docDefaultColor: detectedDocDefaults.color}
                : {}),
            }
          : {};

        const mergedImportPatch: Partial<DocPageLayout> = {
          ...docxDefaultsPatch,
          ...importedLayoutPatch,
        };

        if (meta) {
          setLayout({...DEFAULT_DOC_LAYOUT, ...mergedImportPatch, ...meta.layout});
          setComments(Array.isArray(meta.comments) ? meta.comments : []);
          setNotes(Array.isArray(meta.notes) ? meta.notes : []);
        } else {
          setLayout({...DEFAULT_DOC_LAYOUT, ...mergedImportPatch});
          setComments([]);
          setNotes([]);
        }

        // Never write to IndexedDB when opening a file in View mode
        if (existingVersions.length === 0 && modeRef.current === 'edit') {
          const statsNow = computeDocumentStats(loadedHtml);
          const initialSnap = await saveDocVersion({
            fileId: tab.fileId,
            label: 'Initial version',
            html: loadedHtml,
            wordCount: statsNow.words,
          });
          if (!cancelled) setVersions([initialSnap]);
        } else {
          setVersions(existingVersions);
        }
      } catch (err) {
        if (!cancelled) {
          pushToast(
            err instanceof Error ? err.message : 'Unable to open document.',
            'error',
          );
        }
      } finally {
        if (!cancelled) {
          isHydratingRef.current = false;
          setLoadingDoc(false);
        }
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [editor, tab.fileId, tab.extension, tab.name, tab.discardNonce, pushToast]);

  // Binary save handler wired to workspace autosave & manual save
  // Protects original .docx bytes when the source file contains lossy/unsupported Word features
  const persistCurrentDocument = useCallback(async () => {
    if (!editor || modeRef.current !== 'edit') return;
    if (hasLossyElementsRef.current && !allowOverwriteRef.current) {
      return;
    }
    const html = editor.getHTML();
    const {blob} = await serializeDocumentToBlob(
      html,
      tab.extension,
      tab.name,
      layoutRef.current,
      notesRef.current,
      rawDocxBufferRef.current,
    );
    const preview = htmlToPlainText(html).slice(0, 200);
    await saveDocMeta<DocPersistedMeta>(tab.fileId, {
      layout: layoutRef.current,
      comments: commentsRef.current,
      notes: notesRef.current,
    });
    await saveTabBlob(tab.fileId, blob, preview);
    if (tab.extension.toLowerCase() === '.docx') {
      const savedBuf = await blob.arrayBuffer();
      if (await isRealDocxArchive(savedBuf)) {
        setRawDocxBuffer(savedBuf);
        rawDocxBufferRef.current = savedBuf;
        setDocxPreviewFailed(false);
      }
    }
  }, [editor, tab.fileId, tab.extension, tab.name, saveTabBlob]);

  const handleRequestSave = useCallback(() => {
    if (hasLossyElementsRef.current && !allowOverwriteRef.current) {
      setPendingLossyAction('save');
      setOpenDialog('save-lossy');
      return;
    }
    void persistCurrentDocument().then(() => pushToast('Document saved', 'success'));
  }, [persistCurrentDocument, pushToast]);

  const handleRequestDoneEditing = useCallback(() => {
    if (tab.dirty && hasLossyElementsRef.current && !allowOverwriteRef.current) {
      setPendingLossyAction('done');
      setOpenDialog('save-lossy');
      return;
    }
    onDoneEditing();
  }, [tab.dirty, onDoneEditing]);

  const handleSaveAsCopy = useCallback(async () => {
    if (!editor) return;
    const html = editor.getHTML();
    const baseName = tab.name.replace(/\.[^.]+$/, '');
    const ext = tab.extension.toLowerCase() || '.docx';
    const copyName = `${baseName} (Edited)${ext}`;
    const {blob} = await serializeDocumentToBlob(
      html,
      ext,
      baseName,
      layoutRef.current,
      notesRef.current,
      rawDocxBufferRef.current,
    );
    const bytes = new Uint8Array(await blob.arrayBuffer());
    await createFileFromBytes(copyName, 'docs', bytes, fileRecord?.folderId ?? null, false);
    pushToast(`Saved copy "${copyName}" to vault (original untouched)`, 'success');
    if (pendingLossyAction === 'done') {
      setPendingLossyAction(null);
      await discardTabChanges(tab.fileId);
    } else {
      setPendingLossyAction(null);
    }
  }, [
    editor,
    tab.name,
    tab.extension,
    tab.fileId,
    fileRecord?.folderId,
    createFileFromBytes,
    discardTabChanges,
    pendingLossyAction,
    pushToast,
  ]);

  const handleOverwriteOriginal = useCallback(async () => {
    allowOverwriteRef.current = true;
    setAllowOverwriteOriginal(true);
    await persistCurrentDocument();
    pushToast('Original document updated', 'success');
    const action = pendingLossyAction;
    setPendingLossyAction(null);
    if (action === 'done') {
      onDoneEditing();
    }
  }, [persistCurrentDocument, pendingLossyAction, onDoneEditing, pushToast]);

  useEffect(() => {
    registerBinarySaveHandler(tab.fileId, persistCurrentDocument);
    return () => registerBinarySaveHandler(tab.fileId, null);
  }, [tab.fileId, persistCurrentDocument, registerBinarySaveHandler]);

  // Recompute real page-by-page pagination on layout, notes, mode, or load changes
  // Intentionally NOT dependent on `zoom`: measurement is strictly at scale 1, and zoom only applies CSS transform.
  useEffect(() => {
    if (!editor || loadingDoc || useDocxPreview) return;
    let cancelled = false;
    const runPass = () => {
      if (cancelled || editor.isDestroyed) return;
      try {
        measureAndPaginateEditor(editor.view, paginationConfigRef.current);
      } catch {
        // Ignore if view is mid-transition
      }
    };
    const rafId = window.requestAnimationFrame(runPass);
    // Re-measure once webfonts finish loading to prevent late line-wrap shifts
    const fontsReady = typeof document !== 'undefined' ? document.fonts?.ready : undefined;
    if (fontsReady) {
      void fontsReady.then(() => {
        if (!cancelled) window.requestAnimationFrame(runPass);
      });
    }
    // Re-measure once any pending images in the editor finish decoding
    try {
      const imgs = editor.view?.dom?.querySelectorAll<HTMLImageElement>('img') ?? [];
      imgs.forEach((img) => {
        if (!img.complete) {
          img.addEventListener('load', () => {
            if (!cancelled) window.requestAnimationFrame(runPass);
          }, {once: true});
          if (typeof img.decode === 'function') {
            void img.decode().then(() => {
              if (!cancelled) window.requestAnimationFrame(runPass);
            }).catch(() => {});
          }
        }
      });
    } catch {
      // Ignore if view is mid-mount
    }
    return () => {
      cancelled = true;
      window.cancelAnimationFrame(rafId);
    };
  }, [editor, layout, notes, viewMode, loadingDoc, isEditing, useDocxPreview]);

  // Keep currentPage clamped within [1 .. activeTotalPages]
  const activeTotalPages = useDocxPreview
    ? Math.max(1, docxPreviewMetrics?.pageCount ?? 1)
    : paginationPlan.pageCount;

  useEffect(() => {
    setCurrentPage((prev) => Math.min(activeTotalPages, Math.max(1, prev)));
  }, [activeTotalPages]);

  // Keyboard shortcuts (Ctrl+S save in edit mode, Ctrl+F/Ctrl+H find)
  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      const key = e.key.toLowerCase();
      if ((e.ctrlKey || e.metaKey) && key === 's') {
        if (modeRef.current !== 'edit') return;
        e.preventDefault();
        handleRequestSave();
      } else if ((e.ctrlKey || e.metaKey) && (key === 'h' || key === 'f')) {
        e.preventDefault();
        setOpenDialog('find');
      }
    };
    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, [handleRequestSave]);

  const handleUpdateLayout = useCallback(
    (patch: Partial<DocPageLayout>) => {
      if (modeRef.current !== 'edit') return;
      setLayout((prev) => ({...prev, ...patch}));
      markTabDirty(tab.fileId, true);
    },
    [markTabDirty, tab.fileId],
  );

  // Format Painter copy & apply
  const handleToggleFormatPainter = () => {
    if (!editor) return;
    if (formatPainterMarks) {
      setFormatPainterMarks(null);
      return;
    }
    const ts = editor.getAttributes('textStyle') as {
      color?: string;
      fontFamily?: string;
      fontSize?: string;
    };
    const hl = editor.getAttributes('highlight') as {color?: string};
    setFormatPainterMarks({
      bold: editor.isActive('bold'),
      italic: editor.isActive('italic'),
      underline: editor.isActive('underline'),
      strike: editor.isActive('strike'),
      color: ts?.color,
      fontFamily: ts?.fontFamily,
      fontSize: ts?.fontSize,
      highlight: editor.isActive('highlight') ? hl?.color || '#FEF08A' : undefined,
    });
    pushToast('Format Painter active — select target text to apply formatting.', 'info');
  };

  const applyFormatPainterIfActive = () => {
    if (!editor || !formatPainterMarks || modeRef.current !== 'edit') return;
    if (editor.state.selection.empty) return;
    let chain = editor.chain().focus();
    chain = formatPainterMarks.bold ? chain.setBold() : chain.unsetBold();
    chain = formatPainterMarks.italic ? chain.setItalic() : chain.unsetItalic();
    chain = formatPainterMarks.underline ? chain.setUnderline() : chain.unsetUnderline();
    chain = formatPainterMarks.strike ? chain.setStrike() : chain.unsetStrike();
    if (formatPainterMarks.color) chain = chain.setColor(formatPainterMarks.color);
    else chain = chain.unsetColor();
    if (formatPainterMarks.fontFamily) chain = chain.setFontFamily(formatPainterMarks.fontFamily);
    if (formatPainterMarks.fontSize) chain = chain.setFontSize(formatPainterMarks.fontSize);
    if (formatPainterMarks.highlight) {
      chain = chain.setHighlight({color: formatPainterMarks.highlight});
    } else {
      chain = chain.unsetHighlight();
    }
    chain.run();
    setFormatPainterMarks(null);
  };

  const handlePickImageFile = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file || !editor) return;
    const reader = new FileReader();
    reader.onload = () => {
      if (typeof reader.result === 'string') {
        editor.chain().focus().setImage({src: reader.result, alt: file.name}).run();
        pushToast(`Inserted ${file.name}`, 'success');
      }
    };
    reader.readAsDataURL(file);
  };

  const handleInsertCoverPage = () => {
    if (!editor) return;
    const cleanTitle = tab.name.replace(/\.[^.]+$/, '');
    const dateStr = new Date().toLocaleDateString('en-US', {
      year: 'numeric',
      month: 'long',
      day: 'numeric',
    });
    editor
      .chain()
      .focus()
      .insertContentAt(
        0,
        `<h1>${cleanTitle}</h1><p><strong>Subtitle or Document Summary</strong></p><p>Prepared locally with eSZ OFFICE — ${dateStr}</p><div data-page-break="true">Page Break</div>`,
      )
      .run();
    pushToast('Inserted Cover Page', 'success');
  };

  const handleInsertToc = () => {
    if (!editor) return;
    const headings = extractHeadingsFromHtml(editor.getHTML());
    const tocHtml = buildTocHtml(headings);
    editor.chain().focus().insertContent(tocHtml).run();
    pushToast(
      headings.length > 0
        ? `Inserted Table of Contents (${headings.length} heading${headings.length === 1 ? '' : 's'})`
        : 'Inserted Table of Contents placeholder',
      'success',
    );
  };

  const handleInsertNote = (kind: 'footnote' | 'endnote') => {
    if (!editor || modeRef.current !== 'edit') return;
    const nextNum = notes.length + 1;
    const noteId = `note_${Date.now()}_${nextNum}`;
    const newNote: DocNoteItem = {
      id: noteId,
      kind,
      number: nextNum,
      text: `${kind === 'footnote' ? 'Footnote' : 'Endnote'} ${nextNum} reference details.`,
    };
    setNotes((prev) => [...prev, newNote]);
    editor
      .chain()
      .focus()
      .insertContent({
        type: 'noteRef',
        attrs: {noteId, number: nextNum, kind},
      })
      .run();
    markTabDirty(tab.fileId, true);
    pushToast(`Inserted ${kind} [${nextNum}]`, 'success');
  };

  const handleInsertCaption = () => {
    if (!editor) return;
    const count = (editor.getHTML().match(/data-doc-caption="true"/g) || []).length + 1;
    editor.chain().focus().insertDocCaption(`Figure ${count}: Enter caption description here`).run();
  };

  const handleStartComment = () => {
    if (editor) {
      const {from, to} = editor.state.selection;
      const selected = editor.state.doc.textBetween(from, to, ' ').trim();
      setDraftCommentQuote(selected.slice(0, 140));
    } else {
      setDraftCommentQuote('');
    }
    setCommentsOpen(true);
    setCommentComposerOpen(true);
  };

  const handleSubmitComment = (text: string) => {
    if (modeRef.current !== 'edit') return;
    const item: DocComment = {
      id: `cmt_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
      author: 'Local Author',
      quote: draftCommentQuote,
      text,
      createdAt: Date.now(),
      resolved: false,
    };
    setComments((prev) => [item, ...prev]);
    setCommentComposerOpen(false);
    setDraftCommentQuote('');
    markTabDirty(tab.fileId, true);
    pushToast('Comment added', 'success');
  };

  const handleExportFormat = async (
    format: '.docx' | '.pdf' | '.html' | '.md' | '.rtf' | '.txt',
    target: 'device' | 'vault',
  ) => {
    if (!editor) return;
    const html = editor.getHTML();
    const baseName = tab.name.replace(/\.[^.]+$/, '');
    const outName = `${baseName}${format}`;

    if (format === '.pdf') {
      try {
        let pdfBytes: Uint8Array;
        if (useDocxPreview && previewRootRef.current) {
          pdfBytes = await exportDocxPreviewDomToPdf(previewRootRef.current, {
            title: baseName,
            layout,
            notes,
          });
        } else if (viewMode === 'print' && sheetRef.current) {
          pdfBytes = await exportPaginatedEditorDomToPdf(sheetRef.current, {
            title: baseName,
            layout,
            notes,
          });
        } else {
          pdfBytes = await htmlToPdfBytes(html, baseName, layout, notes);
        }

        if (target === 'vault') {
          await createFileFromBytes(outName, 'pdf', pdfBytes, fileRecord?.folderId ?? null, true);
        } else {
          const blob = new Blob([new Uint8Array(pdfBytes)], {type: 'application/pdf'});
          await saveBlobToDevice(blob, outName, 'pdf');
          pushToast(`Exported ${outName}`, 'success');
        }
      } catch (err) {
        const reason = err instanceof Error ? err.message : 'PDF export error';
        pushToast(
          `Direct PDF export failed (${reason}). Falling back to browser Print to PDF.`,
          'error',
        );
        handlePrint();
      }
      return;
    }

    // In View mode with an unmodified original .docx, preserve exact original bytes on download/copy
    if (format === '.docx' && !isEditing && !tab.dirty && rawDocxBuffer) {
      const originalBytes = new Uint8Array(rawDocxBuffer);
      if (target === 'vault') {
        await createFileFromBytes(outName, 'docs', originalBytes, fileRecord?.folderId ?? null, false);
      } else {
        const mimeType =
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
        await saveBlobToDevice(new Blob([originalBytes], {type: mimeType}), outName, 'docs');
        pushToast(`Saved ${outName}`, 'success');
      }
      return;
    }

    const {blob} = await serializeDocumentToBlob(
      html,
      format,
      baseName,
      layout,
      notes,
      format === '.docx' ? rawDocxBufferRef.current : null,
    );
    if (target === 'vault') {
      const bytes = new Uint8Array(await blob.arrayBuffer());
      await createFileFromBytes(outName, 'docs', bytes, fileRecord?.folderId ?? null, false);
    } else {
      await saveBlobToDevice(blob, outName, 'docs');
      pushToast(`Saved ${outName}`, 'success');
    }
  };

  const handleQuickExportPdf = async () => {
    await handleExportFormat('.pdf', 'device');
  };

  const handlePrint = () => {
    const embeddedFontCss = buildEmbeddedFontFaceCss(
      typeof window !== 'undefined' ? window.location.origin : '',
    );

    const triggerIframePrint = (iframe: HTMLIFrameElement, htmlDoc: string) => {
      const win = iframe.contentWindow;
      if (!win) {
        iframe.remove();
        return;
      }
      win.document.open();
      win.document.write(htmlDoc);
      win.document.close();
      win.focus();

      let printed = false;
      const runPrint = () => {
        if (printed) return;
        printed = true;
        try {
          win.print();
        } finally {
          setTimeout(() => iframe.remove(), 1200);
        }
      };

      const fontsReady = win.document.fonts?.ready;
      if (fontsReady) {
        void Promise.race([
          fontsReady,
          new Promise((resolve) => setTimeout(resolve, 600)),
        ]).then(() => setTimeout(runPrint, 60));
      } else {
        setTimeout(runPrint, 180);
      }
    };

    if (useDocxPreview && previewRootRef.current) {
      const previewRoot = previewRootRef.current;
      const previewParent = previewRoot.closest('[data-testid="docx-preview-root"]');
      const styleTags = Array.from(previewParent?.querySelectorAll('style') ?? [])
        .map((s) => s.outerHTML)
        .join('\n');
      const firstSec = previewRoot.querySelector<HTMLElement>('section.esz-docx');
      const secWpt = firstSec?.offsetWidth ? `${(firstSec.offsetWidth * 0.75).toFixed(2)}pt` : null;
      const secHpt = firstSec?.offsetHeight ? `${(firstSec.offsetHeight * 0.75).toFixed(2)}pt` : null;
      const pageSizeRule = secWpt && secHpt ? `size: ${secWpt} ${secHpt};` : '';
      const iframe = document.createElement('iframe');
      iframe.style.position = 'fixed';
      iframe.style.right = '0';
      iframe.style.bottom = '0';
      iframe.style.width = '0';
      iframe.style.height = '0';
      iframe.style.border = '0';
      document.body.appendChild(iframe);
      const printHtml = `<!doctype html>
<html><head><title>${escapeHtml(tab.name)}</title>
<style>${embeddedFontCss}</style>
${styleTags}
<style>
@page { ${pageSizeRule} margin: 0; }
body { margin: 0; background: #fff; }
.esz-docx-wrapper { background: transparent !important; padding: 0 !important; display: block !important; }
section.esz-docx { margin: 0 auto !important; box-shadow: none !important; border: none !important; page-break-after: always; break-after: page; position: relative; overflow: hidden; }
section.esz-docx:last-child { page-break-after: auto; break-after: auto; }
</style></head><body>${previewRoot.innerHTML}</body></html>`;
      triggerIframePrint(iframe, printHtml);
      return;
    }

    if (!editor) return;
    const iframe = document.createElement('iframe');
    iframe.style.position = 'fixed';
    iframe.style.right = '0';
    iframe.style.bottom = '0';
    iframe.style.width = '0';
    iframe.style.height = '0';
    iframe.style.border = '0';
    document.body.appendChild(iframe);
    const theme = DOC_THEMES[layout.theme];
    const printIsImported = Boolean(layout.hasImportedStyles || layout.docDefaultFont);
    const pm = getPageMetrics(layout);
    const pageCssSize =
      layout.customPageWidthPx && layout.customPageHeightPx
        ? `${(pm.pageWidth * 0.75).toFixed(2)}pt ${(pm.pageHeight * 0.75).toFixed(2)}pt`
        : layout.size === 'a4'
          ? `A4 ${layout.orientation}`
          : layout.size === 'legal'
            ? `legal ${layout.orientation}`
            : `letter ${layout.orientation}`;
    const mTopPt = Math.round(pm.marginTop * 0.75);
    const mRightPt = Math.round(pm.marginRight * 0.75);
    const mBottomPt = Math.round(pm.marginBottom * 0.75);
    const mLeftPt = Math.round(pm.marginLeft * 0.75);

    const cssEsc = (s: string) => s.replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/\n/g, ' ');
    const headerCss = cssEsc(layout.headerText.trim());
    const footerCss = cssEsc(layout.footerText.trim());
    const firstHeaderCss = cssEsc((layout.firstPageHeaderText ?? '').trim());
    const firstFooterCss = cssEsc((layout.firstPageFooterText ?? '').trim());
    const pageCounterCss = '"Page " counter(page) " of " counter(pages)';

    const topRightContent =
      layout.showPageNumbers && layout.pageNumberPosition === 'top-right'
        ? pageCounterCss
        : '""';
    const bottomCenterContent =
      layout.showPageNumbers && layout.pageNumberPosition === 'bottom-center'
        ? pageCounterCss
        : '""';
    const bottomRightContent =
      layout.showPageNumbers && layout.pageNumberPosition === 'bottom-right'
        ? pageCounterCss
        : '""';

    const firstPageRule = layout.differentFirstPage
      ? `
@page :first {
  @top-left { content: "${firstHeaderCss}"; font-size: 9pt; color: #64748b; }
  @top-right { content: ""; }
  @bottom-left { content: "${firstFooterCss}"; font-size: 9pt; color: #64748b; }
  @bottom-center { content: ""; }
  @bottom-right { content: ""; }
}`
      : '';

    const docHtml = `<!doctype html>
<html><head><title>${escapeHtml(tab.name)}</title>
<style>
${embeddedFontCss}
@page {
  size: ${pageCssSize};
  margin: ${mTopPt}pt ${mRightPt}pt ${mBottomPt}pt ${mLeftPt}pt;
  @top-left { content: "${headerCss}"; font-size: 9pt; color: #64748b; vertical-align: bottom; padding-bottom: 8pt; }
  @top-right { content: ${topRightContent}; font-size: 9pt; color: #64748b; vertical-align: bottom; padding-bottom: 8pt; }
  @bottom-left { content: "${footerCss}"; font-size: 9pt; color: #64748b; vertical-align: top; padding-top: 8pt; }
  @bottom-center { content: ${bottomCenterContent}; font-size: 9pt; color: #64748b; vertical-align: top; padding-top: 8pt; }
  @bottom-right { content: ${bottomRightContent}; font-size: 9pt; color: #64748b; vertical-align: top; padding-top: 8pt; }
}
${firstPageRule}
* { box-sizing: border-box; }
body {
  font-family: ${layout.docDefaultFont ? resolveRenderFontStack(layout.docDefaultFont) : theme.fontFamily};
  color: ${layout.docDefaultColor ?? theme.textColor};
  font-size: ${layout.docDefaultFontSizePt ? `${layout.docDefaultFontSizePt}pt` : '11pt'};
  line-height: 1.5;
  margin: 0;
  column-count: ${layout.columns};
  column-gap: ${layout.columns > 1 ? '21pt' : '0'};
  column-rule: ${layout.columns > 1 ? '0.75pt solid #e2e8f0' : 'none'};
}
h1, h2, h3, h4 {
  color: ${printIsImported ? (layout.docDefaultColor ?? '#000000') : theme.headingColor};
  ${printIsImported ? 'font-size: inherit; font-weight: inherit; font-style: inherit; line-height: inherit; margin-top: 0; margin-bottom: 0;' : ''}
  break-after: avoid;
  page-break-after: avoid;
}
p {
  margin-top: 0;
  margin-bottom: ${printIsImported ? '0' : '0.65em'};
  orphans: 2;
  widows: 2;
}
table { border-collapse: collapse; width: 100%; margin: 10pt 0; break-inside: auto; }
thead { display: table-header-group; }
tr { break-inside: avoid; page-break-inside: avoid; }
th, td { border: 1px solid #cbd5e1; padding: 5pt 8pt; vertical-align: top; }
th { background: #f1f5f9; font-weight: 600; text-align: left; }
blockquote { border-left: 2.5pt solid ${theme.accentColor}; padding-left: 10pt; font-style: italic; color: #475569; }
[data-page-break="true"] { break-after: page; page-break-after: always; visibility: hidden; height: 0; margin: 0; padding: 0; border: 0; }
.print-watermark {
  position: fixed;
  inset: 0;
  display: flex;
  align-items: center;
  justify-content: center;
  pointer-events: none;
  z-index: 0;
}
.print-watermark span {
  transform: rotate(-45deg);
  font-size: 48pt;
  font-weight: 800;
  letter-spacing: 0.14em;
  text-transform: uppercase;
  color: rgba(148, 163, 184, 0.18);
}
.print-body { position: relative; z-index: 1; }
.print-notes { margin-top: 16pt; border-top: 0.75pt solid #cbd5e1; padding-top: 8pt; font-size: 9pt; color: #475569; break-inside: avoid; }
</style></head>
<body>
${
  layout.watermarkEnabled && layout.watermarkText.trim()
    ? `<div class="print-watermark"><span>${escapeHtml(layout.watermarkText.trim())}</span></div>`
    : ''
}
<div class="print-body">
${editor.getHTML()}
${
  notes.length > 0
    ? `<div class="print-notes"><strong>Footnotes &amp; Endnotes</strong>${notes
        .map((n) => `<p style="margin:3pt 0">[${n.number}] (${n.kind}) ${escapeHtml(n.text)}</p>`)
        .join('')}</div>`
    : ''
}
</div>
</body></html>`;
    triggerIframePrint(iframe, docHtml);
  };

  const handleSaveVersion = async (label: string) => {
    if (!editor || modeRef.current !== 'edit') return;
    const html = editor.getHTML();
    const s = computeDocumentStats(html);
    await saveDocVersion({
      fileId: tab.fileId,
      label,
      html,
      wordCount: s.words,
    });
    const refreshed = await listDocVersions(tab.fileId);
    setVersions(refreshed);
    pushToast(`Saved snapshot "${label}"`, 'success');
  };

  const handleRestoreVersion = (version: DocVersionRecord) => {
    if (!editor || modeRef.current !== 'edit') return;
    editor.commands.setContent(version.html, {emitUpdate: true});
    setHtmlSnapshot(version.html);
    markTabDirty(tab.fileId, true);
    pushToast(`Restored "${version.label}"`, 'success');
  };

  const handleDeleteVersion = async (versionId: string) => {
    if (modeRef.current !== 'edit') return;
    await deleteDocVersion(versionId);
    setVersions((prev) => prev.filter((v) => v.id !== versionId));
  };

  // Geometry & styling derived from layout + real measured pagination
  const stats = useMemo(() => computeDocumentStats(htmlSnapshot), [htmlSnapshot]);
  const pageMetrics = useMemo(() => getPageMetrics(layout), [layout]);
  const pageWidthPx = pageMetrics.pageWidth;
  const pageHeightPx = pageMetrics.pageHeight;
  const margins = {
    top: pageMetrics.marginTop,
    right: pageMetrics.marginRight,
    bottom: pageMetrics.marginBottom,
    left: pageMetrics.marginLeft,
  };
  const theme = DOC_THEMES[layout.theme];
  const isImportedDocx = Boolean(layout.hasImportedStyles || layout.docDefaultFont);
  const effectiveDocFontFamily = layout.docDefaultFont
    ? resolveRenderFontStack(layout.docDefaultFont)
    : theme.fontFamily;
  const effectiveDocFontSize = layout.docDefaultFontSizePt
    ? `${layout.docDefaultFontSizePt}pt`
    : '11pt';
  const effectiveDocTextColor = layout.docDefaultColor ?? theme.textColor;
  const effectiveDocHeadingColor = isImportedDocx
    ? (layout.docDefaultColor ?? '#000000')
    : theme.headingColor;

  const totalPages =
    viewMode === 'print'
      ? useDocxPreview
        ? Math.max(1, docxPreviewMetrics?.pageCount ?? 1)
        : paginationPlan.pageCount
      : stats.estimatedPages;
  const safeCurrentPage = Math.min(totalPages, Math.max(1, currentPage));
  const scaledPaperWidth = Math.round(pageWidthPx * zoom);
  const scaledStackHeight = Math.round(paginationPlan.totalStackHeight * zoom);

  const handleCanvasScroll = useCallback(() => {
    if (viewMode !== 'print') return;
    const container = scrollContainerRef.current;
    if (!container) return;

    const pageElements = useDocxPreview
      ? Array.from(
          previewRootRef.current?.querySelectorAll<HTMLElement>('section.esz-docx') ?? [],
        )
      : Array.from(
          sheetRef.current?.querySelectorAll<HTMLElement>('section[data-testid="docs-page"]') ?? [],
        );

    if (pageElements.length > 0) {
      const containerRect = container.getBoundingClientRect();
      const pageBounds = pageElements.map((sec) => {
        const r = sec.getBoundingClientRect();
        return {
          top: r.top - containerRect.top + container.scrollTop,
          height: Math.max(1, r.height),
        };
      });
      const visible = computeVisiblePageNumber(
        container.scrollTop,
        container.clientHeight,
        pageBounds,
      );
      setCurrentPage((prev) => (prev === visible ? prev : visible));
      return;
    }

    const strideScaled = (pageHeightPx + DOC_PAGE_GAP_PX) * zoom;
    if (strideScaled <= 0) return;
    const approx = Math.floor((container.scrollTop + container.clientHeight * 0.5) / strideScaled) + 1;
    const clamped = Math.min(paginationPlan.pageCount, Math.max(1, approx));
    setCurrentPage((prev) => (prev === clamped ? prev : clamped));
  }, [viewMode, useDocxPreview, pageHeightPx, zoom, paginationPlan.pageCount]);

  // Keep toolbar "Page X of Y" synchronized with visible page after zoom or page-count changes
  useEffect(() => {
    if (viewMode !== 'print' || loadingDoc) return;
    const rafId = window.requestAnimationFrame(() => {
      handleCanvasScroll();
    });
    return () => window.cancelAnimationFrame(rafId);
  }, [zoom, activeTotalPages, viewMode, loadingDoc, useDocxPreview, handleCanvasScroll]);

  const jumpToPage = useCallback(
    (targetPage: number) => {
      const clamped = Math.min(totalPages, Math.max(1, targetPage));
      setCurrentPage(clamped);
      const container = scrollContainerRef.current;
      if (!container || viewMode !== 'print') return;

      const pageElements = useDocxPreview
        ? Array.from(
            previewRootRef.current?.querySelectorAll<HTMLElement>('section.esz-docx') ?? [],
          )
        : Array.from(
            sheetRef.current?.querySelectorAll<HTMLElement>('section[data-testid="docs-page"]') ??
              [],
          );
      const targetSec = pageElements[clamped - 1];
      if (targetSec) {
        const containerRect = container.getBoundingClientRect();
        const targetTop =
          targetSec.getBoundingClientRect().top - containerRect.top + container.scrollTop - 16;
        container.scrollTo({
          top: Math.max(0, Math.round(targetTop)),
          behavior: 'smooth',
        });
        return;
      }

      const pageTop = paginationPlan.pageTops[clamped - 1] ?? 0;
      container.scrollTo({
        top: Math.max(0, Math.round(pageTop * zoom)),
        behavior: 'smooth',
      });
    },
    [totalPages, useDocxPreview, paginationPlan.pageTops, viewMode, zoom],
  );

  const borderClass = layout.pageBorderConfig
    ? ''
    : layout.pageBorder === 'box'
      ? 'ring-2 ring-slate-800 ring-inset'
      : layout.pageBorder === 'double'
        ? 'outline-4 outline-double outline-slate-700 -outline-offset-8'
        : layout.pageBorder === 'shadow'
          ? 'ring-4 ring-blue-700/30 ring-inset'
          : '';

  return (
    <div className="flex h-full min-h-0 flex-col bg-slate-100 dark:bg-slate-950">
      {/* Hidden local image input */}
      <input
        ref={imageInputRef}
        type="file"
        accept="image/png,image/jpeg,image/webp,image/gif,image/svg+xml"
        onChange={handlePickImageFile}
        className="hidden"
      />

      {/* Top Bar: Slim View Mode Toolbar OR Full 6-tab Editing Ribbon */}
      {isEditing ? (
        <DocsRibbon
          editor={editor}
          activeTab={activeTab}
          onChangeTab={setActiveTab}
          layout={layout}
          onUpdateLayout={handleUpdateLayout}
          viewMode={viewMode}
          onChangeViewMode={setViewMode}
          zoom={zoom}
          onChangeZoom={setZoom}
          showRuler={showRuler}
          onToggleRuler={() => setShowRuler((v) => !v)}
          showFormattingMarks={showFormattingMarks}
          onToggleFormattingMarks={() => setShowFormattingMarks((v) => !v)}
          spellCheck={spellCheck}
          onToggleSpellCheck={() => setSpellCheck((v) => !v)}
          formatPainterActive={Boolean(formatPainterMarks)}
          onToggleFormatPainter={handleToggleFormatPainter}
          isFavorite={Boolean(fileRecord?.isFavorite)}
          onToggleFavorite={onToggleFavorite}
          commentCount={comments.length}
          commentsOpen={commentsOpen}
          onToggleComments={() => setCommentsOpen((v) => !v)}
          onOpenDialog={(kind) => setOpenDialog(kind)}
          onPickImage={() => imageInputRef.current?.click()}
          onInsertCoverPage={handleInsertCoverPage}
          onInsertToc={handleInsertToc}
          onInsertNote={handleInsertNote}
          onInsertCaption={handleInsertCaption}
          onAddComment={handleStartComment}
          onSaveNow={handleRequestSave}
          onExportPdf={() => void handleQuickExportPdf()}
          onPrint={handlePrint}
          onDoneEditing={handleRequestDoneEditing}
          currentPage={safeCurrentPage}
          totalPages={totalPages}
          onJumpToPage={jumpToPage}
        />
      ) : (
        <ViewModeToolbar
          fileName={tab.name}
          moduleType="docs"
          zoomLabel={`${Math.round(zoom * 100)}%`}
          onZoomOut={() => setZoom((z) => Math.max(0.45, Number((z - 0.1).toFixed(2))))}
          onZoomIn={() => setZoom((z) => Math.min(1.8, Number((z + 0.1).toFixed(2))))}
          onZoomReset={() => setZoom(1)}
          isFavorite={Boolean(fileRecord?.isFavorite)}
          onToggleFavorite={onToggleFavorite}
          onSearch={() => setOpenDialog('find')}
          searchActive={openDialog === 'find'}
          onPrint={handlePrint}
          onExport={() => setOpenDialog('export')}
          exportLabel="Export or download document"
          onEdit={onEnterEdit}
          navigationSlot={
            <div className="flex shrink-0 items-center gap-1.5">
              {viewMode === 'print' && (
                <div className="flex items-center rounded-lg border border-slate-200 bg-slate-50 dark:border-slate-700 dark:bg-slate-950">
                  <button
                    type="button"
                    onClick={() => jumpToPage(safeCurrentPage - 1)}
                    disabled={safeCurrentPage <= 1}
                    aria-label="Previous page"
                    title="Previous page"
                    className="rounded-l-lg p-1 text-slate-600 hover:bg-slate-200/70 disabled:opacity-35 dark:text-slate-300 dark:hover:bg-slate-800"
                  >
                    <ChevronLeft className="h-3.5 w-3.5" />
                  </button>
                  <span
                    data-testid="docs-toolbar-page-indicator"
                    className="whitespace-nowrap px-1.5 text-xs font-medium tabular-nums text-slate-700 dark:text-slate-200"
                  >
                    {formatPageNumber(safeCurrentPage, totalPages)}
                  </span>
                  <button
                    type="button"
                    onClick={() => jumpToPage(safeCurrentPage + 1)}
                    disabled={safeCurrentPage >= totalPages}
                    aria-label="Next page"
                    title="Next page"
                    className="rounded-r-lg p-1 text-slate-600 hover:bg-slate-200/70 disabled:opacity-35 dark:text-slate-300 dark:hover:bg-slate-800"
                  >
                    <ChevronRight className="h-3.5 w-3.5" />
                  </button>
                </div>
              )}
              {viewMode === 'web' && (
                <span className="whitespace-nowrap rounded-lg border border-slate-200 bg-slate-50 px-2 py-1 text-xs font-medium tabular-nums text-slate-700 dark:border-slate-700 dark:bg-slate-950 dark:text-slate-200">
                  Web Layout
                </span>
              )}
              <div className="hidden items-center rounded-lg border border-slate-200 bg-slate-50 p-0.5 sm:flex dark:border-slate-700 dark:bg-slate-950">
                <button
                  type="button"
                  onClick={() => setViewMode('print')}
                  aria-label="Print Layout"
                  title="Print Layout"
                  className={`rounded-md p-1 ${
                    viewMode === 'print'
                      ? 'bg-blue-100 text-blue-800 dark:bg-blue-950 dark:text-blue-300'
                      : 'text-slate-500 hover:text-slate-800 dark:text-slate-400'
                  }`}
                >
                  <FileText className="h-3.5 w-3.5" />
                </button>
                <button
                  type="button"
                  onClick={() => setViewMode('web')}
                  aria-label="Web Layout"
                  title="Web Layout"
                  className={`rounded-md p-1 ${
                    viewMode === 'web'
                      ? 'bg-blue-100 text-blue-800 dark:bg-blue-950 dark:text-blue-300'
                      : 'text-slate-500 hover:text-slate-800 dark:text-slate-400'
                  }`}
                >
                  <Globe className="h-3.5 w-3.5" />
                </button>
              </div>
            </div>
          }
        />
      )}

      {/* Format Painter active banner */}
      {isEditing && formatPainterMarks && (
        <div className="flex items-center justify-between border-b border-blue-200 bg-blue-50 px-4 py-1.5 text-xs text-blue-900 dark:border-blue-900 dark:bg-blue-950/60 dark:text-blue-200">
          <span>
            Format Painter is armed — highlight target text inside the document to paste formatting.
          </span>
          <button
            type="button"
            onClick={() => setFormatPainterMarks(null)}
            className="font-semibold underline"
          >
            Cancel
          </button>
        </div>
      )}

      {/* Edit-mode Word compatibility / locked-element banner */}
      {isEditing && (hasLockedElements || importWarnings.length > 0) && (
        <div
          data-testid="docs-import-warnings-banner"
          className={`flex flex-wrap items-center justify-between gap-2 border-b px-4 py-1.5 text-xs ${
            hasLockedElements
              ? 'border-amber-200 bg-amber-50 text-amber-950 dark:border-amber-900/70 dark:bg-amber-950/40 dark:text-amber-100'
              : 'border-slate-200 bg-slate-50 text-slate-700 dark:border-slate-800 dark:bg-slate-900/80 dark:text-slate-300'
          }`}
        >
          <span className="inline-flex items-center gap-1.5 font-medium">
            {hasLockedElements ? (
              <AlertTriangle className="h-3.5 w-3.5 shrink-0 text-amber-600 dark:text-amber-400" />
            ) : (
              <Info className="h-3.5 w-3.5 shrink-0 text-blue-600 dark:text-blue-400" />
            )}
            {hasLockedElements
              ? 'Some advanced elements are shown as images and cannot be edited.'
              : `Imported with ${importWarnings.length} Word compatibility note${importWarnings.length === 1 ? '' : 's'}.`}
          </span>
          <button
            type="button"
            onClick={() => setOpenDialog('warnings')}
            className="font-semibold underline hover:opacity-80"
          >
            View details
          </button>
        </div>
      )}

      {/* Main Workspace Canvas + Optional Comments Drawer */}
      <div className="relative flex min-h-0 flex-1 overflow-hidden">
        <div
          ref={scrollContainerRef}
          onScroll={handleCanvasScroll}
          className="min-w-0 flex-1 overflow-auto px-3 py-4 sm:px-6 sm:py-5"
        >
          {loadingDoc ? (
            <div className="mx-auto max-w-3xl space-y-4 rounded-xl bg-white p-10 shadow-sm dark:bg-slate-900">
              <div className="h-7 w-1/2 animate-pulse rounded bg-slate-200 dark:bg-slate-800" />
              <div className="h-4 w-full animate-pulse rounded bg-slate-100 dark:bg-slate-800" />
              <div className="h-4 w-5/6 animate-pulse rounded bg-slate-100 dark:bg-slate-800" />
              <div className="h-4 w-4/6 animate-pulse rounded bg-slate-100 dark:bg-slate-800" />
            </div>
          ) : useDocxPreview && rawDocxBuffer ? (
            <DocxPreviewView
              buffer={rawDocxBuffer}
              zoom={zoom}
              previewRootRef={previewRootRef}
              onMetricsChange={(metrics) => {
                setDocxPreviewMetrics(metrics);
              }}
              onError={(err) => {
                setDocxPreviewFailed(true);
                pushToast(
                  `Fell back to standard viewer (${err.message}).`,
                  'info',
                );
              }}
            />
          ) : viewMode === 'print' ? (
            <div className="flex min-w-full w-max flex-col items-center pb-8">
              {isEditing && showRuler && (
                <HorizontalRuler
                  pageWidthPx={pageWidthPx}
                  marginLeftPx={margins.left}
                  marginRightPx={margins.right}
                  zoom={zoom}
                />
              )}

              {/* Outer scaled page-stack wrapper */}
              <div
                style={{
                  width: `${scaledPaperWidth}px`,
                  height: `${scaledStackHeight}px`,
                }}
                className="relative shrink-0"
              >
                {/* Unscaled multi-page physical stack transformed by zoom */}
                <div
                  ref={sheetRef}
                  data-testid="docs-print-sheet"
                  onMouseUp={applyFormatPainterIfActive}
                  className="relative origin-top-left"
                  style={
                    {
                      width: `${pageWidthPx}px`,
                      height: `${paginationPlan.totalStackHeight}px`,
                      transform: `scale(${zoom})`,
                      fontFamily: effectiveDocFontFamily,
                      fontSize: effectiveDocFontSize,
                      color: effectiveDocTextColor,
                      '--doc-heading-color': effectiveDocHeadingColor,
                      '--doc-accent-color': theme.accentColor,
                    } as React.CSSProperties
                  }
                >
                  {/* Layer 1: Individual physical page cards with own margins, shadow, gap, header, footer & clipped watermark */}
                  {Array.from({length: paginationPlan.pageCount}, (_, pageIdx) => {
                    const pageNum = pageIdx + 1;
                    const pageTop = paginationPlan.pageTops[pageIdx] ?? pageIdx * (pageHeightPx + DOC_PAGE_GAP_PX);
                    const showChrome = shouldShowPageChrome(pageIdx, layout.differentFirstPage);
                    const pageHeaderText = showChrome
                      ? layout.headerText
                      : (layout.firstPageHeaderText ?? '');
                    const pageFooterText = showChrome
                      ? layout.footerText
                      : (layout.firstPageFooterText ?? '');
                    const pageNumberText = formatPageNumber(pageNum, paginationPlan.pageCount);
                    const showTopRightNum =
                      showChrome &&
                      layout.showPageNumbers &&
                      layout.pageNumberPosition === 'top-right';
                    const showBottomCenterNum =
                      showChrome &&
                      layout.showPageNumbers &&
                      layout.pageNumberPosition === 'bottom-center';
                    const showBottomRightNum =
                      showChrome &&
                      layout.showPageNumbers &&
                      layout.pageNumberPosition === 'bottom-right';

                    const pageBorderOverlayStyle = buildPageBorderOverlayStyle(
                      layout.pageBorderConfig,
                      pageIdx,
                      pageMetrics,
                    );

                    return (
                      <section
                        key={pageIdx}
                        data-testid="docs-page"
                        data-page-number={pageNum}
                        aria-label={`Page ${pageNum} of ${paginationPlan.pageCount}`}
                        onMouseDown={(e) => {
                          if (!isEditing) return;
                          const target = e.target as HTMLElement;
                          if (target.closest('input, button, a')) return;
                          editor?.commands.focus();
                        }}
                        className={`absolute left-0 w-full overflow-hidden border border-slate-300 bg-white shadow-md dark:border-slate-700 ${borderClass}`}
                        style={{
                          top: `${pageTop}px`,
                          height: `${pageHeightPx}px`,
                        }}
                      >
                        {/* Exact OOXML Page Border Frame (w:pgBorders) */}
                        {pageBorderOverlayStyle && (
                          <div
                            aria-hidden="true"
                            data-testid="docs-page-border-frame"
                            className="pointer-events-none absolute z-10 box-border"
                            style={pageBorderOverlayStyle}
                          />
                        )}

                        {/* Per-page centered diagonal Watermark (clipped strictly to this page) */}
                        {layout.watermarkEnabled && layout.watermarkText.trim() && (
                          <div
                            aria-hidden="true"
                            data-testid="docs-page-watermark"
                            className="pointer-events-none absolute inset-0 z-0 flex items-center justify-center overflow-hidden"
                          >
                            <span className="-rotate-45 select-none whitespace-nowrap text-6xl font-extrabold tracking-widest text-slate-400/20 uppercase">
                              {layout.watermarkText}
                            </span>
                          </div>
                        )}

                        {/* Per-page Header inside top margin */}
                        <div
                          data-testid="docs-page-header"
                          className="pointer-events-none absolute inset-x-0 top-0 z-10 flex items-end justify-between pb-3 text-xs text-slate-500"
                          style={{
                            height: `${margins.top}px`,
                            paddingLeft: `${margins.left}px`,
                            paddingRight: `${margins.right}px`,
                          }}
                        >
                          <span className="min-w-0 truncate">{pageHeaderText}</span>
                          {showTopRightNum && (
                            <span
                              data-testid="docs-page-number"
                              className="ml-4 shrink-0 tabular-nums"
                            >
                              {pageNumberText}
                            </span>
                          )}
                        </div>

                        {/* Per-page Footer inside bottom margin */}
                        <div
                          data-testid="docs-page-footer"
                          className="pointer-events-none absolute inset-x-0 bottom-0 z-10 grid grid-cols-3 items-start pt-3 text-xs text-slate-500"
                          style={{
                            height: `${margins.bottom}px`,
                            paddingLeft: `${margins.left}px`,
                            paddingRight: `${margins.right}px`,
                          }}
                        >
                          <span className="truncate text-left">{pageFooterText}</span>
                          <span className="text-center tabular-nums">
                            {showBottomCenterNum && (
                              <span data-testid="docs-page-number">{pageNumberText}</span>
                            )}
                          </span>
                          <span className="text-right tabular-nums">
                            {showBottomRightNum && (
                              <span data-testid="docs-page-number">{pageNumberText}</span>
                            )}
                          </span>
                        </div>
                      </section>
                    );
                  })}

                  {/* Layer 2: Paginated Rich ProseMirror Flow (transparent over physical page cards) */}
                  <div
                    data-default-tab-stop={layout.defaultTabStopPx ?? 48}
                    className={`esz-doc-editor esz-paginated absolute inset-x-0 z-10 ${
                      isImportedDocx ? 'esz-imported-docx' : ''
                    } ${showFormattingMarks ? 'esz-show-formatting-marks' : ''} ${
                      layout.columns > 1 ? 'esz-multi-col' : ''
                    } ${!isEditing ? 'esz-view-readonly caret-transparent' : ''}`}
                    style={
                      {
                        top: `${margins.top}px`,
                        paddingRight: `${margins.right}px`,
                        paddingLeft: `${margins.left}px`,
                        columnCount: layout.columns > 1 ? layout.columns : undefined,
                        columnGap: layout.columns > 1 ? '28px' : undefined,
                        columnRule: layout.columns > 1 ? '1px solid #e2e8f0' : undefined,
                        '--doc-max-block-height': `${Math.max(160, pageMetrics.contentHeight - DOC_BOTTOM_SAFETY_PX)}px`,
                      } as React.CSSProperties
                    }
                  >
                    <EditorContent editor={editor} />
                  </div>

                  {/* Layer 3: Interactive per-page Footnotes & final-page Endnotes anchored above each page's bottom margin */}
                  <div className="pointer-events-none absolute inset-0 z-20">
                    {Array.from({length: paginationPlan.pageCount}, (_, pageIdx) => {
                      const pageFns = paginationPlan.pageFootnotes[pageIdx] ?? [];
                      const isLastPage = pageIdx === paginationPlan.pageCount - 1;
                      const pageEndnotes = isLastPage ? paginationPlan.endnotes : [];
                      const pageNotes = [...pageFns, ...pageEndnotes];
                      if (pageNotes.length === 0) return null;

                      const pageTop =
                        paginationPlan.pageTops[pageIdx] ??
                        pageIdx * (pageHeightPx + DOC_PAGE_GAP_PX);
                      const bottomAnchorFromStackTop =
                        pageTop + pageHeightPx - margins.bottom - 4;

                      return (
                        <div
                          key={`notes-${pageIdx}`}
                          data-testid="docs-page-footnotes"
                          data-page-number={pageIdx + 1}
                          className="pointer-events-auto absolute border-t border-slate-300 bg-white/95 pt-1.5 text-[11px] text-slate-600"
                          style={{
                            left: `${margins.left}px`,
                            right: `${margins.right}px`,
                            top: `${bottomAnchorFromStackTop}px`,
                            transform: 'translateY(-100%)',
                          }}
                        >
                          <div className="space-y-1">
                            {pageNotes.map((n) => (
                              <div key={n.id} className="flex items-center gap-1.5">
                                <span className="shrink-0 font-semibold tabular-nums text-blue-700">
                                  [{n.number}]
                                </span>
                                <span className="shrink-0 text-[9px] font-medium uppercase text-slate-400">
                                  {n.kind === 'endnote' ? 'end' : 'fn'}
                                </span>
                                {isEditing ? (
                                  <>
                                    <input
                                      type="text"
                                      aria-label={`${n.kind} ${n.number} text`}
                                      value={n.text}
                                      onChange={(e) => {
                                        const val = e.target.value;
                                        setNotes((prev) =>
                                          prev.map((item) =>
                                            item.id === n.id ? {...item, text: val} : item,
                                          ),
                                        );
                                        markTabDirty(tab.fileId, true);
                                      }}
                                      className="min-w-0 flex-1 rounded border border-transparent bg-slate-50 px-1.5 py-0.5 text-[11px] text-slate-800 focus:border-blue-500 focus:bg-white focus:outline-none"
                                    />
                                    <button
                                      type="button"
                                      aria-label={`Delete ${n.kind} ${n.number}`}
                                      onClick={() => {
                                        setNotes((prev) => prev.filter((item) => item.id !== n.id));
                                        markTabDirty(tab.fileId, true);
                                      }}
                                      className="rounded p-0.5 text-slate-400 hover:text-rose-600"
                                    >
                                      <Trash2 className="h-3 w-3" />
                                    </button>
                                  </>
                                ) : (
                                  <span className="min-w-0 flex-1 truncate text-[11px] text-slate-700">
                                    {n.text}
                                  </span>
                                )}
                              </div>
                            ))}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              </div>
            </div>
          ) : (
            /* Web Layout View */
            <div
              data-testid="docs-web-sheet"
              onMouseUp={applyFormatPainterIfActive}
              className="mx-auto min-h-[75vh] max-w-5xl rounded-xl border border-slate-200 bg-white px-6 py-8 shadow-sm sm:px-10 dark:border-slate-800 dark:bg-slate-900"
              style={
                {
                  fontFamily: effectiveDocFontFamily,
                  fontSize: effectiveDocFontSize,
                  color: effectiveDocTextColor,
                  '--doc-heading-color': effectiveDocHeadingColor,
                  '--doc-accent-color': theme.accentColor,
                } as React.CSSProperties
              }
            >
              <div
                data-default-tab-stop={layout.defaultTabStopPx ?? 48}
                className={`esz-doc-editor dark:text-slate-100 ${
                  isImportedDocx ? 'esz-imported-docx' : ''
                } ${showFormattingMarks ? 'esz-show-formatting-marks' : ''} ${
                  !isEditing ? 'esz-view-readonly caret-transparent' : ''
                }`}
              >
                <EditorContent editor={editor} />
              </div>
            </div>
          )}
        </div>

        {/* Comments Panel (overlay on small screens, docked sidebar on sm+) */}
        {isEditing && commentsOpen && (
          <div className="absolute inset-y-0 right-0 z-30 flex w-full sm:static sm:w-auto">
            <DocsCommentsSidebar
              comments={comments}
              draftQuote={draftCommentQuote}
              composerOpen={commentComposerOpen}
              onClosePanel={() => setCommentsOpen(false)}
              onCancelComposer={() => setCommentComposerOpen(false)}
              onSubmitComment={handleSubmitComment}
              onToggleResolve={(id) => {
                if (modeRef.current !== 'edit') return;
                setComments((prev) =>
                  prev.map((c) => (c.id === id ? {...c, resolved: !c.resolved} : c)),
                );
                markTabDirty(tab.fileId, true);
              }}
              onDeleteComment={(id) => {
                if (modeRef.current !== 'edit') return;
                setComments((prev) => prev.filter((c) => c.id !== id));
                markTabDirty(tab.fileId, true);
              }}
            />
          </div>
        )}
      </div>

      {/* Bottom Document Status Bar */}
      <footer className="flex flex-wrap items-center justify-between gap-2 border-t border-slate-200 bg-white px-3 py-1.5 text-xs text-slate-600 dark:border-slate-800 dark:bg-slate-900 dark:text-slate-400">
        <div className="flex items-center gap-3 tabular-nums">
          <span data-testid="docs-status-page-indicator">
            {formatPageNumber(safeCurrentPage, totalPages)}
          </span>
          <button
            type="button"
            onClick={() => setOpenDialog('word-count')}
            className="font-medium text-slate-700 hover:text-blue-700 hover:underline dark:text-slate-200 dark:hover:text-blue-400"
          >
            {stats.words.toLocaleString('en-US')} words
          </button>
          <span className="hidden sm:inline">
            {DOC_PAGE_DIMENSIONS_PX[layout.size].label.split(' ')[0]} ·{' '}
            {layout.orientation === 'landscape' ? 'Landscape' : 'Portrait'}
          </span>
          <span className="hidden md:inline">
            Spell Check: {spellCheck ? 'On' : 'Off'}
          </span>
        </div>

        <div className="flex items-center gap-2">
          <div className="flex items-center rounded-md border border-slate-200 p-0.5 dark:border-slate-700">
            <button
              type="button"
              onClick={() => setViewMode('print')}
              aria-label="Print Layout"
              title="Print Layout"
              className={`rounded px-1.5 py-0.5 ${
                viewMode === 'print'
                  ? 'bg-blue-100 text-blue-800 dark:bg-blue-950 dark:text-blue-300'
                  : 'text-slate-500 hover:text-slate-800'
              }`}
            >
              <FileText className="h-3.5 w-3.5" />
            </button>
            <button
              type="button"
              onClick={() => setViewMode('web')}
              aria-label="Web Layout"
              title="Web Layout"
              className={`rounded px-1.5 py-0.5 ${
                viewMode === 'web'
                  ? 'bg-blue-100 text-blue-800 dark:bg-blue-950 dark:text-blue-300'
                  : 'text-slate-500 hover:text-slate-800'
              }`}
            >
              <Globe className="h-3.5 w-3.5" />
            </button>
          </div>

          <input
            type="range"
            aria-label="Document zoom level"
            min={45}
            max={180}
            step={5}
            value={Math.round(zoom * 100)}
            onChange={(e) => setZoom(Number(e.target.value) / 100)}
            className="hidden h-1.5 w-24 cursor-pointer accent-blue-600 sm:inline-block"
          />
          <span className="w-10 text-right font-semibold tabular-nums text-slate-700 dark:text-slate-200">
            {Math.round(zoom * 100)}%
          </span>
        </div>
      </footer>

      {/* Mobile View-Mode Bottom Action Bar */}
      {!isEditing && (
        <ViewModeMobileBar
          moduleType="docs"
          onEdit={onEnterEdit}
          onSearch={() => setOpenDialog('find')}
          onExport={() => setOpenDialog('export')}
        />
      )}

      {/* All Modal Dialogs */}
      <DocsDialogs
        openDialog={openDialog}
        onClose={() => {
          setOpenDialog(null);
          setPendingLossyAction(null);
        }}
        editor={editor}
        fileName={tab.name}
        layout={layout}
        onUpdateLayout={handleUpdateLayout}
        stats={stats}
        actualPageCount={totalPages}
        versions={versions}
        onSaveVersion={handleSaveVersion}
        onRestoreVersion={handleRestoreVersion}
        onDeleteVersion={handleDeleteVersion}
        onExportFormat={handleExportFormat}
        importWarnings={importWarnings}
        hasLockedElements={hasLockedElements}
        onSaveAsCopy={handleSaveAsCopy}
        onOverwriteOriginal={handleOverwriteOriginal}
        readOnly={!isEditing}
      />
    </div>
  );
};

export default DocsModule;
