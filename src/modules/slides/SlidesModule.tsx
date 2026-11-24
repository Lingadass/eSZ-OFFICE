import React, {useCallback, useEffect, useMemo, useRef, useState} from 'react';
import {
  AlertTriangle,
  ChevronLeft,
  ChevronRight,
  MessageSquareText,
  PanelLeft,
  PanelRight,
  Play,
  ZoomIn,
  ZoomOut,
} from 'lucide-react';
import {saveBlobToDevice} from '../../core/file-io/fileSystemAccess';
import {getDocMeta, loadFileBlob, saveDocMeta} from '../../core/storage/fileRepository';
import {useWorkspaceStore} from '../../core/workspace/useWorkspaceStore';
import {ViewModeMobileBar, ViewModeToolbar} from '../../shell/ViewModeToolbar';
import type {OfficeModuleProps} from '../pdf/PdfModule';
import {SlidesCanvas, type SlidesCanvasHandle} from './SlidesCanvas';
import {SlidesDialogs} from './SlidesDialogs';
import {SlidesPresenter} from './SlidesPresenter';
import {SlidesRibbon} from './SlidesRibbon';
import {SlideFilmstrip, SlideInspector} from './SlidesSidebar';
import {
  alignElementOnSlide,
  applyThemeToDeck,
  buildDeckPreviewText,
  buildPersistedSlidesMeta,
  createBlankDeck,
  createDeckFromTemplate,
  createNewElement,
  createSlideFromLayout,
  createSlideUid,
  deckToPdfBytes,
  deckToPptxBytes,
  duplicateSlideItem,
  getSlideTheme,
  moveElementLayer,
  parseOutlineToSlides,
  pptxBytesToDeck,
  reorderSlides,
} from './slidesPure';
import {analyzeSlideTextFit, getSlideDimensions} from './slidesTextFit';
import type {
  PersistedSlidesMeta,
  PresentationDeck,
  ShapeKind,
  SlideAspectRatio,
  SlideChartKind,
  SlideDialogKind,
  SlideElement,
  SlideLayoutKind,
  SlideMasterSettings,
  SlideRibbonTab,
  SlideTransitionKind,
} from './slidesTypes';

const MAX_UNDO = 40;

const SlidesModule: React.FC<OfficeModuleProps> = ({
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
    pushToast,
  } = useWorkspaceStore();

  const isEditing = mode === 'edit';
  const modeRef = useRef(mode);
  modeRef.current = mode;

  const canvasRef = useRef<SlidesCanvasHandle | null>(null);
  const isHydratingRef = useRef(true);

  const [loadingDeck, setLoadingDeck] = useState(true);
  const [deck, setDeck] = useState<PresentationDeck>(() => createBlankDeck(tab.name));
  const [past, setPast] = useState<PresentationDeck[]>([]);
  const [future, setFuture] = useState<PresentationDeck[]>([]);

  const [ribbonTab, setRibbonTab] = useState<SlideRibbonTab>('home');
  const [selectedElementId, setSelectedElementId] = useState<string | null>(null);
  const [zoom, setZoom] = useState(100);
  const [showFilmstrip, setShowFilmstrip] = useState(() =>
    typeof window === 'undefined' ? true : window.innerWidth >= 768,
  );
  const [showInspector, setShowInspector] = useState(() =>
    typeof window === 'undefined' ? true : window.innerWidth >= 1180,
  );
  const [showNotes, setShowNotes] = useState(true);

  const [openDialog, setOpenDialog] = useState<SlideDialogKind>(null);
  const [defaultChartKind, setDefaultChartKind] = useState<SlideChartKind>('bar');
  const [presenting, setPresenting] = useState<{open: boolean; startIndex: number}>({
    open: false,
    startIndex: 0,
  });

  const deckRef = useRef(deck);
  deckRef.current = deck;
  const pastRef = useRef(past);
  pastRef.current = past;
  const futureRef = useRef(future);
  futureRef.current = future;

  const activeSlideIndex = useMemo(() => {
    const idx = deck.slides.findIndex((s) => s.id === deck.activeSlideId);
    return idx >= 0 ? idx : 0;
  }, [deck.slides, deck.activeSlideId]);

  const activeSlide = deck.slides[activeSlideIndex] ?? deck.slides[0];
  const activeTheme = useMemo(() => getSlideTheme(deck.themeId), [deck.themeId]);
  const selectedElement = useMemo(
    () => activeSlide.elements.find((e) => e.id === selectedElementId) ?? null,
    [activeSlide.elements, selectedElementId],
  );
  const slideTextFit = useMemo(() => analyzeSlideTextFit(activeSlide), [activeSlide]);

  // Clear selection and edit-only modals when switching back to View mode
  useEffect(() => {
    if (!isEditing) {
      setSelectedElementId(null);
      setOpenDialog((prev) => (prev === 'export' ? prev : null));
    }
  }, [isEditing]);

  // Hydrate deck from persisted Dexie metadata or PPTX binary blob
  useEffect(() => {
    let cancelled = false;
    isHydratingRef.current = true;
    setLoadingDeck(true);

    (async () => {
      try {
        const [meta, blob] = await Promise.all([
          getDocMeta<PersistedSlidesMeta>(tab.fileId),
          loadFileBlob(tab.fileId),
        ]);
        if (cancelled) return;

        let loadedDeck: PresentationDeck;
        if (meta?.version === 1 && meta.deck?.slides?.length > 0) {
          loadedDeck = meta.deck;
        } else if (blob && blob.size > 0) {
          const bytes = new Uint8Array(await blob.arrayBuffer());
          loadedDeck = await pptxBytesToDeck(bytes, tab.name);
        } else {
          loadedDeck = createBlankDeck(tab.name);
        }

        if (cancelled) return;
        setDeck(loadedDeck);
        setPast([]);
        setFuture([]);
        setSelectedElementId(null);
      } catch (err) {
        if (!cancelled) {
          pushToast(
            err instanceof Error ? err.message : 'Unable to open presentation.',
            'error',
          );
          setDeck(createBlankDeck(tab.name));
        }
      } finally {
        if (!cancelled) {
          isHydratingRef.current = false;
          setLoadingDeck(false);
        }
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [tab.fileId, tab.name, tab.discardNonce, pushToast]);

  // Commit deck update with undo history (flat state updates, never inside an updater)
  const commitDeck = useCallback(
    (updater: (prev: PresentationDeck) => PresentationDeck, recordUndo = true) => {
      if (modeRef.current !== 'edit') return;
      const prev = deckRef.current;
      const next = updater(prev);
      if (next === prev) return;
      deckRef.current = next;
      setDeck(next);
      if (recordUndo && !isHydratingRef.current) {
        const nextPast = [...pastRef.current.slice(-(MAX_UNDO - 1)), prev];
        pastRef.current = nextPast;
        futureRef.current = [];
        setPast(nextPast);
        setFuture([]);
        markTabDirty(tab.fileId, true);
      }
    },
    [markTabDirty, tab.fileId],
  );

  const handleUndo = useCallback(() => {
    if (modeRef.current !== 'edit') return;
    const currentPast = pastRef.current;
    if (currentPast.length === 0) return;
    const previous = currentPast[currentPast.length - 1];
    const nextPast = currentPast.slice(0, -1);
    const nextFuture = [deckRef.current, ...futureRef.current.slice(0, MAX_UNDO - 1)];
    pastRef.current = nextPast;
    futureRef.current = nextFuture;
    deckRef.current = previous;
    setPast(nextPast);
    setFuture(nextFuture);
    setDeck(previous);
    markTabDirty(tab.fileId, true);
  }, [markTabDirty, tab.fileId]);

  const handleRedo = useCallback(() => {
    if (modeRef.current !== 'edit') return;
    const currentFuture = futureRef.current;
    if (currentFuture.length === 0) return;
    const next = currentFuture[0];
    const nextFuture = currentFuture.slice(1);
    const nextPast = [...pastRef.current.slice(-(MAX_UNDO - 1)), deckRef.current];
    pastRef.current = nextPast;
    futureRef.current = nextFuture;
    deckRef.current = next;
    setPast(nextPast);
    setFuture(nextFuture);
    setDeck(next);
    markTabDirty(tab.fileId, true);
  }, [markTabDirty, tab.fileId]);

  // Binary PPTX + Dexie metadata persistence
  const persistDeck = useCallback(
    async (silent = false) => {
      if (modeRef.current !== 'edit') return;
      const current = deckRef.current;
      try {
        await saveDocMeta(tab.fileId, buildPersistedSlidesMeta(current));
        const bytes = await deckToPptxBytes(current);
        const blob = new Blob([new Uint8Array(bytes)], {
          type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        });
        const ok = await saveTabBlob(tab.fileId, blob, buildDeckPreviewText(current));
        if (ok && !silent) {
          pushToast('Presentation saved', 'success');
        }
      } catch (err) {
        if (!silent) {
          pushToast(
            err instanceof Error ? err.message : 'Failed to save presentation.',
            'error',
          );
        }
      }
    },
    [pushToast, saveTabBlob, tab.fileId],
  );

  useEffect(() => {
    registerBinarySaveHandler(tab.fileId, () => persistDeck(true));
    return () => registerBinarySaveHandler(tab.fileId, null);
  }, [persistDeck, registerBinarySaveHandler, tab.fileId]);

  // Element & Slide Actions
  const handleSelectSlide = useCallback((slideId: string) => {
    setSelectedElementId(null);
    setDeck((prev) => ({...prev, activeSlideId: slideId}));
  }, []);

  const handleAddSlide = useCallback(
    (layout: SlideLayoutKind) => {
      commitDeck((prev) => {
        const theme = getSlideTheme(prev.themeId);
        const created = createSlideFromLayout(layout, theme);
        const idx = prev.slides.findIndex((s) => s.id === prev.activeSlideId);
        const nextSlides = [...prev.slides];
        nextSlides.splice(idx >= 0 ? idx + 1 : nextSlides.length, 0, created);
        return {...prev, slides: nextSlides, activeSlideId: created.id};
      });
      setSelectedElementId(null);
    },
    [commitDeck],
  );

  const handleDuplicateSlide = useCallback(
    (slideId?: string) => {
      commitDeck((prev) => {
        const targetId = slideId ?? prev.activeSlideId;
        const idx = prev.slides.findIndex((s) => s.id === targetId);
        if (idx < 0) return prev;
        const copy = duplicateSlideItem(prev.slides[idx]);
        const nextSlides = [...prev.slides];
        nextSlides.splice(idx + 1, 0, copy);
        return {...prev, slides: nextSlides, activeSlideId: copy.id};
      });
      setSelectedElementId(null);
    },
    [commitDeck],
  );

  const handleDeleteSlide = useCallback(
    (slideId?: string) => {
      commitDeck((prev) => {
        if (prev.slides.length <= 1) return prev;
        const targetId = slideId ?? prev.activeSlideId;
        const idx = prev.slides.findIndex((s) => s.id === targetId);
        if (idx < 0) return prev;
        const nextSlides = prev.slides.filter((s) => s.id !== targetId);
        const nextActive =
          prev.activeSlideId === targetId
            ? nextSlides[Math.max(0, idx - 1)].id
            : prev.activeSlideId;
        return {...prev, slides: nextSlides, activeSlideId: nextActive};
      });
      setSelectedElementId(null);
    },
    [commitDeck],
  );

  const handleReorderSlides = useCallback(
    (fromIndex: number, toIndex: number) => {
      commitDeck((prev) => ({
        ...prev,
        slides: reorderSlides(prev.slides, fromIndex, toIndex),
      }));
    },
    [commitDeck],
  );

  const handleApplyLayout = useCallback(
    (layout: SlideLayoutKind) => {
      commitDeck((prev) => {
        const theme = getSlideTheme(prev.themeId);
        return {
          ...prev,
          slides: prev.slides.map((s) => {
            if (s.id !== prev.activeSlideId) return s;
            const existingTitle = s.elements.find((e) => e.type === 'text')?.text;
            const fresh = createSlideFromLayout(layout, theme, {
              title: existingTitle,
              notes: s.notes,
            });
            return {...fresh, id: s.id, transition: s.transition};
          }),
        };
      });
      setSelectedElementId(null);
    },
    [commitDeck],
  );

  const handleUpdateElement = useCallback(
    (elementId: string, patch: Partial<SlideElement>) => {
      commitDeck((prev) => ({
        ...prev,
        slides: prev.slides.map((s) =>
          s.id !== prev.activeSlideId
            ? s
            : {
                ...s,
                elements: s.elements.map((el) =>
                  el.id === elementId ? ({...el, ...patch} as SlideElement) : el,
                ),
              },
        ),
      }));
    },
    [commitDeck],
  );

  const handleUpdateSelectedElement = useCallback(
    (patch: Partial<SlideElement>) => {
      if (!selectedElementId) return;
      handleUpdateElement(selectedElementId, patch);
    },
    [handleUpdateElement, selectedElementId],
  );

  const handleDeleteElement = useCallback(
    (elementId?: string) => {
      const target = elementId ?? selectedElementId;
      if (!target) return;
      commitDeck((prev) => ({
        ...prev,
        slides: prev.slides.map((s) =>
          s.id !== prev.activeSlideId
            ? s
            : {...s, elements: s.elements.filter((el) => el.id !== target)},
        ),
      }));
      if (target === selectedElementId) setSelectedElementId(null);
    },
    [commitDeck, selectedElementId],
  );

  const handleDuplicateSelectedElement = useCallback(() => {
    if (!selectedElement) return;
    const dims = getSlideDimensions(deckRef.current.aspectRatio);
    const clone: SlideElement = {
      ...selectedElement,
      id: createSlideUid(selectedElement.type.slice(0, 3)),
      x: Math.min(dims.width - 140, selectedElement.x + 24),
      y: Math.min(dims.height - 100, selectedElement.y + 24),
    };
    commitDeck((prev) => ({
      ...prev,
      slides: prev.slides.map((s) =>
        s.id !== prev.activeSlideId ? s : {...s, elements: [...s.elements, clone]},
      ),
    }));
    setSelectedElementId(clone.id);
  }, [commitDeck, selectedElement]);

  const appendElementToActiveSlide = useCallback(
    (el: SlideElement) => {
      commitDeck((prev) => ({
        ...prev,
        slides: prev.slides.map((s) =>
          s.id !== prev.activeSlideId ? s : {...s, elements: [...s.elements, el]},
        ),
      }));
      setSelectedElementId(el.id);
    },
    [commitDeck],
  );

  const handleAddText = useCallback(() => {
    appendElementToActiveSlide(createNewElement('text', activeTheme));
  }, [activeTheme, appendElementToActiveSlide]);

  const handleAddShape = useCallback(
    (shapeKind: ShapeKind) => {
      appendElementToActiveSlide(createNewElement('shape', activeTheme, {shapeKind}));
    },
    [activeTheme, appendElementToActiveSlide],
  );

  const handleInsertImageFile = useCallback(
    (file: File) => {
      const reader = new FileReader();
      reader.onload = () => {
        if (typeof reader.result !== 'string') return;
        const el = createNewElement('image', activeTheme, {
          imageSrc: reader.result,
          imageAlt: file.name,
        });
        appendElementToActiveSlide(el);
        pushToast(`Inserted ${file.name}`, 'success');
      };
      reader.readAsDataURL(file);
    },
    [activeTheme, appendElementToActiveSlide, pushToast],
  );

  const handleLayerAction = useCallback(
    (dir: 'front' | 'back') => {
      if (!selectedElementId) return;
      commitDeck((prev) => ({
        ...prev,
        slides: prev.slides.map((s) =>
          s.id !== prev.activeSlideId
            ? s
            : {...s, elements: moveElementLayer(s.elements, selectedElementId, dir)},
        ),
      }));
    },
    [commitDeck, selectedElementId],
  );

  const handleAlignElement = useCallback(
    (alignment: 'left' | 'center' | 'right' | 'top' | 'middle' | 'bottom') => {
      if (!selectedElementId) return;
      commitDeck((prev) => ({
        ...prev,
        slides: prev.slides.map((s) =>
          s.id !== prev.activeSlideId
            ? s
            : {
                ...s,
                elements: s.elements.map((el) =>
                  el.id === selectedElementId
                    ? alignElementOnSlide(el, alignment, prev.aspectRatio)
                    : el,
                ),
              },
        ),
      }));
    },
    [commitDeck, selectedElementId],
  );

  const handleSelectTheme = useCallback(
    (nextThemeId: string) => {
      commitDeck((prev) => applyThemeToDeck(prev, nextThemeId));
      pushToast(`Applied ${getSlideTheme(nextThemeId).name} theme`, 'info');
    },
    [commitDeck, pushToast],
  );

  const handleChangeAspectRatio = useCallback(
    (nextRatio: SlideAspectRatio) => {
      commitDeck((prev) => {
        if (prev.aspectRatio === nextRatio) return prev;
        const oldDims = getSlideDimensions(prev.aspectRatio);
        const newDims = getSlideDimensions(nextRatio);
        const yScale = newDims.height / oldDims.height;
        return {
          ...prev,
          aspectRatio: nextRatio,
          slides: prev.slides.map((s) => ({
            ...s,
            elements: s.elements.map((el) => {
              const scaledY = Math.round(el.y * yScale);
              const scaledH = Math.max(24, Math.round(el.height * yScale));
              const clampedY = Math.max(0, Math.min(newDims.height - 24, scaledY));
              const clampedH = Math.min(scaledH, newDims.height - clampedY);
              return {...el, y: clampedY, height: clampedH};
            }),
          })),
        };
      });
      pushToast(`Slide ratio set to ${nextRatio}`, 'info');
    },
    [commitDeck, pushToast],
  );

  const handleChangeBackground = useCallback(
    (hex: string) => {
      commitDeck((prev) => ({
        ...prev,
        slides: prev.slides.map((s) =>
          s.id !== prev.activeSlideId ? s : {...s, background: hex},
        ),
      }));
    },
    [commitDeck],
  );

  const handleChangeTransition = useCallback(
    (transition: SlideTransitionKind) => {
      commitDeck((prev) => ({
        ...prev,
        slides: prev.slides.map((s) =>
          s.id !== prev.activeSlideId ? s : {...s, transition},
        ),
      }));
    },
    [commitDeck],
  );

  const handleApplyTransitionAll = useCallback(
    (transition: SlideTransitionKind) => {
      commitDeck((prev) => ({
        ...prev,
        slides: prev.slides.map((s) => ({...s, transition})),
      }));
      pushToast(`Applied "${transition}" transition to all slides`, 'info');
    },
    [commitDeck, pushToast],
  );

  const handleUpdateMaster = useCallback(
    (patch: Partial<SlideMasterSettings>) => {
      commitDeck((prev) => ({...prev, master: {...prev.master, ...patch}}));
    },
    [commitDeck],
  );

  const handleUpdateNotes = useCallback(
    (notes: string) => {
      commitDeck((prev) => ({
        ...prev,
        slides: prev.slides.map((s) =>
          s.id !== prev.activeSlideId ? s : {...s, notes},
        ),
      }));
    },
    [commitDeck],
  );

  const handleApplyTemplate = useCallback(
    (templateId: string) => {
      commitDeck((prev) => createDeckFromTemplate(templateId, prev.title));
      setSelectedElementId(null);
      pushToast('Template applied to presentation', 'success');
    },
    [commitDeck, pushToast],
  );

  const handleGenerateFromOutline = useCallback(
    (outlineText: string, replaceAll: boolean) => {
      commitDeck((prev) => {
        const generated = parseOutlineToSlides(outlineText, prev.themeId);
        const nextSlides = replaceAll ? generated : [...prev.slides, ...generated];
        return {
          ...prev,
          slides: nextSlides,
          activeSlideId: generated[0]?.id ?? prev.activeSlideId,
        };
      });
      setSelectedElementId(null);
      pushToast('Slides generated from outline', 'success');
    },
    [commitDeck, pushToast],
  );

  // Exports (PPTX, PDF, PNG)
  const handleExportFormat = useCallback(
    async (format: 'pptx' | 'pdf' | 'png', destination: 'device' | 'vault') => {
      const current = deckRef.current;
      const baseName = tab.name.replace(/\.[^.]+$/, '') || 'Presentation';

      try {
        if (format === 'pptx') {
          const bytes = await deckToPptxBytes(current);
          const outName = `${baseName}.pptx`;
          if (destination === 'vault') {
            await createFileFromBytes(outName, 'slides', bytes, fileRecord?.folderId ?? null, false);
          } else {
            const blob = new Blob([new Uint8Array(bytes)], {
              type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
            });
            await saveBlobToDevice(blob, outName, 'slides');
            pushToast(`Exported ${outName}`, 'success');
          }
          return;
        }

        if (format === 'pdf') {
          const pdfBytes = await deckToPdfBytes(current);
          const outName = `${baseName}.pdf`;
          if (destination === 'vault') {
            await createFileFromBytes(outName, 'pdf', pdfBytes, fileRecord?.folderId ?? null, true);
          } else {
            const blob = new Blob([new Uint8Array(pdfBytes)], {type: 'application/pdf'});
            await saveBlobToDevice(blob, outName, 'pdf');
            pushToast(`Exported ${outName}`, 'success');
          }
          return;
        }

        // PNG snapshot of active slide via Konva stage
        const dataUrl = canvasRef.current?.exportPngDataUrl();
        if (!dataUrl) {
          pushToast('Slide canvas is not ready for PNG export.', 'warning');
          return;
        }
        const res = await fetch(dataUrl);
        const pngBlob = await res.blob();
        const outName = `${baseName}-slide-${activeSlideIndex + 1}.png`;
        const url = URL.createObjectURL(pngBlob);
        const link = document.createElement('a');
        link.href = url;
        link.download = outName;
        document.body.appendChild(link);
        link.click();
        link.remove();
        setTimeout(() => URL.revokeObjectURL(url), 1000);
        pushToast(`Exported ${outName}`, 'success');
      } catch (err) {
        pushToast(
          err instanceof Error ? err.message : `Failed to export ${format.toUpperCase()}.`,
          'error',
        );
      }
    },
    [activeSlideIndex, createFileFromBytes, fileRecord?.folderId, pushToast, tab.name],
  );

  const goRelativeSlide = useCallback(
    (delta: number) => {
      const slides = deckRef.current.slides;
      if (slides.length <= 1) return;
      const currIdx = slides.findIndex((s) => s.id === deckRef.current.activeSlideId);
      const nextIdx = Math.max(0, Math.min(slides.length - 1, (currIdx >= 0 ? currIdx : 0) + delta));
      handleSelectSlide(slides[nextIdx].id);
    },
    [handleSelectSlide],
  );

  // Global keyboard shortcuts (F5 Present; slide nav in View mode; editing keys in Edit mode)
  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      if (presenting.open || openDialog) return;
      const target = e.target as HTMLElement | null;
      const isInput =
        target?.tagName === 'INPUT' ||
        target?.tagName === 'TEXTAREA' ||
        target?.tagName === 'SELECT' ||
        target?.isContentEditable;
      const editingNow = modeRef.current === 'edit';

      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
        if (!editingNow) return;
        e.preventDefault();
        void persistDeck(false);
        return;
      }
      if (e.key === 'F5') {
        e.preventDefault();
        setPresenting({open: true, startIndex: 0});
        return;
      }
      if (isInput) return;

      if (!editingNow) {
        if (e.key === 'ArrowRight' || e.key === 'PageDown') {
          e.preventDefault();
          goRelativeSlide(1);
        } else if (e.key === 'ArrowLeft' || e.key === 'PageUp') {
          e.preventDefault();
          goRelativeSlide(-1);
        }
        return;
      }

      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z' && !e.shiftKey) {
        e.preventDefault();
        handleUndo();
      } else if (
        (e.ctrlKey || e.metaKey) &&
        (e.key.toLowerCase() === 'y' || (e.shiftKey && e.key.toLowerCase() === 'z'))
      ) {
        e.preventDefault();
        handleRedo();
      } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'd' && selectedElementId) {
        e.preventDefault();
        handleDuplicateSelectedElement();
      } else if ((e.key === 'Delete' || e.key === 'Backspace') && selectedElementId) {
        e.preventDefault();
        handleDeleteElement(selectedElementId);
      } else if (
        selectedElement &&
        ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)
      ) {
        e.preventDefault();
        const step = e.shiftKey ? 10 : 2;
        const dx = e.key === 'ArrowLeft' ? -step : e.key === 'ArrowRight' ? step : 0;
        const dy = e.key === 'ArrowUp' ? -step : e.key === 'ArrowDown' ? step : 0;
        handleUpdateElement(selectedElement.id, {
          x: Math.max(0, selectedElement.x + dx),
          y: Math.max(0, selectedElement.y + dy),
        });
      }
    };

    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, [
    goRelativeSlide,
    handleDeleteElement,
    handleDuplicateSelectedElement,
    handleRedo,
    handleUndo,
    handleUpdateElement,
    openDialog,
    persistDeck,
    presenting.open,
    selectedElement,
    selectedElementId,
  ]);

  if (loadingDeck) {
    return (
      <div className="flex h-full flex-col gap-3 p-6">
        <div className="h-16 animate-pulse rounded-xl bg-slate-200 dark:bg-slate-800" />
        <div className="flex flex-1 gap-4">
          <div className="hidden w-52 animate-pulse rounded-xl bg-slate-200 md:block dark:bg-slate-800" />
          <div className="flex-1 animate-pulse rounded-xl bg-slate-100 dark:bg-slate-900" />
        </div>
      </div>
    );
  }

  return (
    <div className="flex h-full min-h-0 flex-col bg-slate-100 dark:bg-slate-950">
      {/* Top Bar: Slim View Mode Toolbar OR Full 5-tab Ribbon */}
      {isEditing ? (
        <SlidesRibbon
          activeTab={ribbonTab}
          onChangeTab={setRibbonTab}
          dirty={tab.dirty}
          isFavorite={Boolean(fileRecord?.isFavorite)}
          canUndo={past.length > 0}
          canRedo={future.length > 0}
          themeId={deck.themeId}
          aspectRatio={deck.aspectRatio}
          slideTransition={activeSlide.transition}
          slideBackground={activeSlide.background}
          selectedElement={selectedElement}
          onSave={() => void persistDeck(false)}
          onUndo={handleUndo}
          onRedo={handleRedo}
          onToggleFavorite={onToggleFavorite}
          onAddSlide={handleAddSlide}
          onDuplicateSlide={() => handleDuplicateSlide()}
          onDeleteSlide={() => handleDeleteSlide()}
          onApplyLayout={handleApplyLayout}
          onAddText={handleAddText}
          onAddShape={handleAddShape}
          onInsertImageFile={handleInsertImageFile}
          onOpenDialog={(dlg, chartK) => {
            if (chartK) setDefaultChartKind(chartK);
            setOpenDialog(dlg);
          }}
          onUpdateSelectedElement={handleUpdateSelectedElement}
          onDeleteSelectedElement={() => handleDeleteElement()}
          onDuplicateSelectedElement={handleDuplicateSelectedElement}
          onLayerAction={handleLayerAction}
          onAlignElement={handleAlignElement}
          onSelectTheme={handleSelectTheme}
          onChangeAspectRatio={handleChangeAspectRatio}
          onChangeBackground={handleChangeBackground}
          onChangeTransition={handleChangeTransition}
          onApplyTransitionAll={handleApplyTransitionAll}
          onStartPresent={(fromStart) =>
            setPresenting({open: true, startIndex: fromStart ? 0 : activeSlideIndex})
          }
          onExportPdf={() => void handleExportFormat('pdf', 'vault')}
          onExportPng={() => void handleExportFormat('png', 'device')}
          onPrint={() => window.print()}
          onDoneEditing={onDoneEditing}
        />
      ) : (
        <ViewModeToolbar
          fileName={tab.name}
          moduleType="slides"
          zoomLabel={`${zoom}%`}
          onZoomOut={() => setZoom((z) => Math.max(50, z - 10))}
          onZoomIn={() => setZoom((z) => Math.min(180, z + 10))}
          onZoomReset={() => setZoom(100)}
          isFavorite={Boolean(fileRecord?.isFavorite)}
          onToggleFavorite={onToggleFavorite}
          onPrint={() => window.print()}
          onExport={() => setOpenDialog('export')}
          exportLabel="Export or download presentation"
          onEdit={onEnterEdit}
          leadingSlot={
            <button
              type="button"
              onClick={() => setShowFilmstrip((v) => !v)}
              aria-label="Toggle slide thumbnails"
              aria-pressed={showFilmstrip}
              title="Toggle slide thumbnails"
              className={`inline-flex h-8 w-8 shrink-0 items-center justify-center rounded-lg border transition-colors ${
                showFilmstrip
                  ? 'border-orange-500 bg-orange-50 text-orange-700 dark:border-orange-500 dark:bg-orange-950/60 dark:text-orange-300'
                  : 'border-slate-200 text-slate-700 hover:bg-slate-100 dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800'
              }`}
            >
              <PanelLeft className="h-3.5 w-3.5" />
            </button>
          }
          navigationSlot={
            <div className="flex shrink-0 items-center gap-0.5 rounded-lg border border-slate-200 bg-slate-50 p-0.5 dark:border-slate-700 dark:bg-slate-950">
              <button
                type="button"
                disabled={activeSlideIndex <= 0}
                onClick={() => goRelativeSlide(-1)}
                aria-label="Previous slide"
                className="flex h-7 w-7 items-center justify-center rounded-md text-slate-700 hover:bg-slate-200/70 disabled:opacity-35 dark:text-slate-200 dark:hover:bg-slate-800"
              >
                <ChevronLeft className="h-3.5 w-3.5" />
              </button>
              <span className="whitespace-nowrap px-1.5 text-xs font-semibold tabular-nums text-slate-700 dark:text-slate-200">
                {activeSlideIndex + 1} / {deck.slides.length}
              </span>
              <button
                type="button"
                disabled={activeSlideIndex >= deck.slides.length - 1}
                onClick={() => goRelativeSlide(1)}
                aria-label="Next slide"
                className="flex h-7 w-7 items-center justify-center rounded-md text-slate-700 hover:bg-slate-200/70 disabled:opacity-35 dark:text-slate-200 dark:hover:bg-slate-800"
              >
                <ChevronRight className="h-3.5 w-3.5" />
              </button>
            </div>
          }
          extraActionsSlot={
            <button
              type="button"
              data-testid="slides-present-btn"
              onClick={() => setPresenting({open: true, startIndex: activeSlideIndex})}
              aria-label="Present slide show (F5)"
              title="Present slide show (F5)"
              className="hidden h-8 shrink-0 items-center gap-1.5 whitespace-nowrap rounded-lg bg-slate-900 px-3 text-xs font-semibold text-white hover:bg-slate-800 sm:inline-flex dark:bg-slate-100 dark:text-slate-900 dark:hover:bg-white"
            >
              <Play className="h-3.5 w-3.5 fill-current" />
              Present
            </button>
          }
        />
      )}

      {/* Main Workspace */}
      <div className="relative flex min-h-0 flex-1 overflow-hidden">
        {/* Left Slide Filmstrip (overlay drawer on mobile, static sidebar on md+) */}
        {showFilmstrip && (
          <div className="absolute inset-y-0 left-0 z-20 flex shadow-xl md:static md:z-auto md:shadow-none">
            <SlideFilmstrip
              slides={deck.slides}
              activeSlideId={activeSlide.id}
              themeId={deck.themeId}
              aspectRatio={deck.aspectRatio}
              master={deck.master}
              readOnly={!isEditing}
              onSelectSlide={(id) => {
                handleSelectSlide(id);
                if (typeof window !== 'undefined' && window.innerWidth < 768) {
                  setShowFilmstrip(false);
                }
              }}
              onAddSlide={handleAddSlide}
              onDuplicateSlide={handleDuplicateSlide}
              onDeleteSlide={handleDeleteSlide}
              onReorderSlides={handleReorderSlides}
              onCloseMobile={() => setShowFilmstrip(false)}
            />
          </div>
        )}

        {/* Center Konva Canvas + Speaker Notes */}
        <div className="flex min-h-0 min-w-0 flex-1 flex-col overflow-hidden">
          <SlidesCanvas
            ref={canvasRef}
            slide={activeSlide}
            slideIndex={activeSlideIndex}
            slideCount={deck.slides.length}
            themeId={deck.themeId}
            aspectRatio={deck.aspectRatio}
            master={deck.master}
            selectedElementId={isEditing ? selectedElementId : null}
            zoom={zoom}
            readOnly={!isEditing}
            onSelectElement={setSelectedElementId}
            onUpdateElement={handleUpdateElement}
          />

          {showNotes && (
            <div
              data-testid="slides-notes-panel"
              className="shrink-0 border-t border-slate-200 bg-white px-4 py-2 dark:border-slate-800 dark:bg-slate-900"
            >
              <div className="mb-1 flex items-center justify-between">
                <label
                  htmlFor={`speaker-notes-${tab.fileId}`}
                  className="text-[11px] font-semibold text-slate-500 dark:text-slate-400"
                >
                  Speaker Notes — Slide {activeSlideIndex + 1}
                </label>
                {isEditing && (
                  <span className="hidden text-[11px] text-slate-400 sm:inline">
                    Double-click text on canvas to edit inline
                  </span>
                )}
              </div>
              {isEditing ? (
                <textarea
                  id={`speaker-notes-${tab.fileId}`}
                  rows={2}
                  value={activeSlide.notes}
                  onChange={(e) => handleUpdateNotes(e.target.value)}
                  placeholder="Click to add speaker notes for Presenter Mode..."
                  className="w-full resize-none rounded-md border border-slate-200 bg-slate-50 px-2.5 py-1.5 text-xs text-slate-800 outline-none focus:border-orange-500 dark:border-slate-800 dark:bg-slate-950 dark:text-slate-100"
                />
              ) : (
                <div
                  id={`speaker-notes-${tab.fileId}`}
                  className="max-h-20 min-h-[38px] overflow-y-auto rounded-md border border-slate-200 bg-slate-50 px-2.5 py-1.5 text-xs whitespace-pre-wrap text-slate-700 dark:border-slate-800 dark:bg-slate-950 dark:text-slate-200"
                >
                  {activeSlide.notes.trim() || 'No speaker notes for this slide.'}
                </div>
              )}
            </div>
          )}
        </div>

        {/* Right Contextual Inspector (Edit Mode only) */}
        {isEditing && showInspector && (
          <div className="absolute inset-y-0 right-0 z-20 flex shadow-xl xl:static xl:z-auto xl:shadow-none">
            <SlideInspector
              slide={activeSlide}
              selectedElement={selectedElement}
              aspectRatio={deck.aspectRatio}
              master={deck.master}
              onSelectElement={setSelectedElementId}
              onUpdateElement={handleUpdateElement}
              onDeleteElement={handleDeleteElement}
              onLayerAction={handleLayerAction}
              onAlignElement={handleAlignElement}
              onChangeBackground={handleChangeBackground}
              onChangeTransition={handleChangeTransition}
              onUpdateMaster={handleUpdateMaster}
              onClose={() => setShowInspector(false)}
            />
          </div>
        )}
      </div>

      {/* Bottom Status Bar */}
      <footer className="flex shrink-0 items-center justify-between gap-2 border-t border-slate-200 bg-white px-3 py-1 text-xs text-slate-600 dark:border-slate-800 dark:bg-slate-900 dark:text-slate-300">
        <div className="flex min-w-0 items-center gap-2">
          <button
            type="button"
            onClick={() => setShowFilmstrip((v) => !v)}
            aria-label="Toggle slides filmstrip"
            aria-pressed={showFilmstrip}
            className={`inline-flex h-6 items-center gap-1 rounded px-2 text-[11px] font-medium ${
              showFilmstrip
                ? 'bg-orange-100 text-orange-800 dark:bg-orange-950/60 dark:text-orange-300'
                : 'hover:bg-slate-100 dark:hover:bg-slate-800'
            }`}
          >
            <PanelLeft className="h-3.5 w-3.5" />
            <span className="hidden sm:inline">Slides</span>
          </button>
          <span className="font-mono text-[11px] tabular-nums">
            Slide {activeSlideIndex + 1} of {deck.slides.length}
          </span>
          <span className="hidden text-slate-300 sm:inline dark:text-slate-700">·</span>
          <span
            data-testid="slides-aspect-ratio"
            className="hidden font-mono text-[11px] text-slate-500 sm:inline"
          >
            {deck.aspectRatio}
          </span>
          <span className="hidden text-slate-300 md:inline dark:text-slate-700">·</span>
          <span className="hidden truncate text-[11px] text-slate-500 md:inline">
            Theme: {activeTheme.name}
          </span>
          {slideTextFit.overflowIds.length > 0 && (
            <button
              type="button"
              data-testid="slides-overflow-warning"
              onClick={() => {
                if (isEditing) {
                  setSelectedElementId(slideTextFit.overflowIds[0]);
                  setShowInspector(true);
                }
              }}
              title={
                isEditing
                  ? 'Click to select overflowing text box and expand it'
                  : 'Text exceeds box bounds even after auto-shrink; switch to Edit to resize'
              }
              className="inline-flex h-5 shrink-0 items-center gap-1 rounded border border-amber-300 bg-amber-50 px-1.5 text-[11px] font-medium text-amber-800 dark:border-amber-800 dark:bg-amber-950/60 dark:text-amber-200"
            >
              <AlertTriangle className="h-3 w-3 shrink-0 text-amber-600 dark:text-amber-400" />
              <span>
                Text overflow ({slideTextFit.overflowIds.length})
              </span>
            </button>
          )}
        </div>

        <div className="flex items-center gap-1.5">
          <button
            type="button"
            onClick={() => setShowNotes((v) => !v)}
            aria-label="Toggle speaker notes panel"
            aria-pressed={showNotes}
            className={`inline-flex h-6 items-center gap-1 rounded px-2 text-[11px] font-medium ${
              showNotes
                ? 'bg-orange-100 text-orange-800 dark:bg-orange-950/60 dark:text-orange-300'
                : 'hover:bg-slate-100 dark:hover:bg-slate-800'
            }`}
          >
            <MessageSquareText className="h-3.5 w-3.5" />
            <span className="hidden sm:inline">Notes</span>
          </button>

          {isEditing && (
            <button
              type="button"
              onClick={() => setShowInspector((v) => !v)}
              aria-label="Toggle inspector sidebar"
              aria-pressed={showInspector}
              className={`inline-flex h-6 items-center gap-1 rounded px-2 text-[11px] font-medium ${
                showInspector
                  ? 'bg-orange-100 text-orange-800 dark:bg-orange-950/60 dark:text-orange-300'
                  : 'hover:bg-slate-100 dark:hover:bg-slate-800'
              }`}
            >
              <PanelRight className="h-3.5 w-3.5" />
              <span className="hidden sm:inline">Inspector</span>
            </button>
          )}

          <div className="mx-1 h-3.5 w-px bg-slate-200 dark:bg-slate-700" />

          <button
            type="button"
            onClick={() => setZoom((z) => Math.max(50, z - 10))}
            aria-label="Zoom out slide"
            className="inline-flex h-6 w-6 items-center justify-center rounded hover:bg-slate-100 dark:hover:bg-slate-800"
          >
            <ZoomOut className="h-3.5 w-3.5" />
          </button>
          <span className="w-9 text-center font-mono text-[11px] tabular-nums">{zoom}%</span>
          <button
            type="button"
            onClick={() => setZoom((z) => Math.min(180, z + 10))}
            aria-label="Zoom in slide"
            className="inline-flex h-6 w-6 items-center justify-center rounded hover:bg-slate-100 dark:hover:bg-slate-800"
          >
            <ZoomIn className="h-3.5 w-3.5" />
          </button>
        </div>
      </footer>

      {/* Mobile View-Mode Bottom Action Bar */}
      {!isEditing && (
        <ViewModeMobileBar
          moduleType="slides"
          onEdit={onEnterEdit}
          onExport={() => setOpenDialog('export')}
          secondarySlot={
            <button
              type="button"
              onClick={() => setPresenting({open: true, startIndex: activeSlideIndex})}
              aria-label="Present slide show"
              className="inline-flex min-h-[40px] shrink-0 items-center gap-1.5 whitespace-nowrap rounded-xl bg-slate-900 px-3.5 text-xs font-semibold text-white hover:bg-slate-800 dark:bg-slate-100 dark:text-slate-900"
            >
              <Play className="h-3.5 w-3.5 fill-current" />
              <span>Present</span>
            </button>
          }
        />
      )}

      {/* Modal Dialogs */}
      <SlidesDialogs
        openDialog={openDialog}
        defaultChartKind={defaultChartKind}
        deckTitle={deck.title}
        master={deck.master}
        onClose={() => setOpenDialog(null)}
        onApplyTemplate={handleApplyTemplate}
        onGenerateFromOutline={handleGenerateFromOutline}
        onCreateChart={(spec) => {
          const el = createNewElement('chart', activeTheme, {
            chartKind: spec.kind,
            chartTitle: spec.title,
            chartLabels: spec.labels,
            chartValues: spec.values,
          });
          appendElementToActiveSlide(el);
          pushToast(`Inserted ${spec.kind} chart`, 'success');
        }}
        onCreateTable={(rows, cols) => {
          const el = createNewElement('table', activeTheme, {
            tableRows: rows,
            tableCols: cols,
          });
          appendElementToActiveSlide(el);
          pushToast(`Inserted ${rows}×${cols} table`, 'success');
        }}
        onUpdateMaster={handleUpdateMaster}
        onExportFormat={(fmt, dest) => void handleExportFormat(fmt, dest)}
      />

      {/* Full-Screen Presenter Mode */}
      {presenting.open && (
        <SlidesPresenter
          deck={deck}
          initialIndex={presenting.startIndex}
          onExit={() => setPresenting({open: false, startIndex: activeSlideIndex})}
        />
      )}
    </div>
  );
};

export default SlidesModule;
