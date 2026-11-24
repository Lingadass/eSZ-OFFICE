import React, {useCallback, useEffect, useMemo, useRef, useState} from 'react';
import {
  BarChart3,
  Check,
  ChevronDown,
  ChevronUp,
  Copy,
  Filter,
  Pencil,
  Plus,
  Search,
  Sigma,
  Trash2,
  X,
} from 'lucide-react';
import {
  getDocMeta,
  loadFileBlob,
  saveDocMeta,
} from '../../core/storage/fileRepository';
import {saveBlobToDevice} from '../../core/file-io/fileSystemAccess';
import {useWorkspaceStore} from '../../core/workspace/useWorkspaceStore';
import type {OfficeModuleProps} from '../pdf/PdfModule';
import {
  colIndexToLetter,
  evaluateWorkbook,
  formatCellRef,
  parseCellRef,
  shiftFormulaReferences,
} from './formulaEngine';
import {
  applyPersistedMeta,
  computeAutoFitColWidth,
  computeStickyOffsets,
  computeVisibleRowTopOffsets,
  createBlankWorkbook,
  createEmptySheet,
  evaluateConditionalFormatting,
  extractPersistedMeta,
  findMergeAt,
  findSheetMatches,
  getValidationForCell,
  sheetToCsvString,
  sortSheetByColumn,
  validateCellInput,
  workbookToPdfBytes,
  workbookToXlsxBytesAsync,
  xlsxBytesToWorkbookAsync,
} from './sheetsPure';
import {
  cellKey,
  DEFAULT_COL_WIDTH,
  DEFAULT_ROW_HEIGHT,
  normalizeSelection,
  type CellCoord,
  type CellStyle,
  type ChartKind,
  type ConditionalRule,
  type DataValidationRule,
  type PersistedSheetMeta,
  type SelectionRange,
  type SheetData,
  type SheetsRibbonTab,
  type WorkbookData,
} from './sheetsTypes';
import {SheetsRibbon} from './SheetsRibbon';
import {SheetsChartsPanel} from './SheetsCharts';
import {SheetsDialogs, type SheetsDialogKind} from './SheetsDialogs';
import {ViewModeMobileBar, ViewModeToolbar} from '../../shell/ViewModeToolbar';

const ROW_HEADER_WIDTH = 44;
const COL_HEADER_HEIGHT = 28;

const SheetsModule: React.FC<OfficeModuleProps> = ({
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

  const [workbook, setWorkbook] = useState<WorkbookData>(() => createBlankWorkbook('Sheet1'));
  const [past, setPast] = useState<WorkbookData[]>([]);
  const [future, setFuture] = useState<WorkbookData[]>([]);
  const [loadingSheet, setLoadingSheet] = useState(true);

  const [ribbonTab, setRibbonTab] = useState<SheetsRibbonTab>('home');
  const [selection, setSelection] = useState<SelectionRange>({
    start: {row: 0, col: 0},
    end: {row: 0, col: 0},
  });
  const [editingCoord, setEditingCoord] = useState<CellCoord | null>(null);
  const [draftValue, setDraftValue] = useState('');
  const [nameBoxText, setNameBoxText] = useState('A1');

  const [showGridlines, setShowGridlines] = useState(true);
  const [showFormulaBar, setShowFormulaBar] = useState(true);
  const [zoom, setZoom] = useState(1);
  const [chartsPanelOpen, setChartsPanelOpen] = useState(false);

  const [openDialog, setOpenDialog] = useState<SheetsDialogKind | null>(null);
  const [defaultChartKind, setDefaultChartKind] = useState<ChartKind>('bar');

  // Auto-Filter open popover column index
  const [openFilterCol, setOpenFilterCol] = useState<number | null>(null);

  // Sheet rename inline state
  const [renamingSheetId, setRenamingSheetId] = useState<string | null>(null);
  const [renameDraft, setRenameDraft] = useState('');
  // Ephemeral column widths used exclusively in View mode (never dirties or writes file)
  const [viewColWidths, setViewColWidths] = useState<Record<string, Record<number, number>>>({});
  // Non-blocking View-mode floating Find bar state
  const [viewFindOpen, setViewFindOpen] = useState(false);
  const [viewFindQuery, setViewFindQuery] = useState('');
  const [viewFindCaseSensitive, setViewFindCaseSensitive] = useState(false);
  const [viewFindIndex, setViewFindIndex] = useState(0);
  const viewFindInputRef = useRef<HTMLInputElement>(null);

  // Drag selection & Fill handle refs
  const isSelectingRef = useRef(false);
  const isFillingRef = useRef(false);
  const [fillTarget, setFillTarget] = useState<CellCoord | null>(null);
  const findCursorRef = useRef(0);
  const gridContainerRef = useRef<HTMLDivElement>(null);
  const inlineInputRef = useRef<HTMLInputElement>(null);
  const workbookRef = useRef(workbook);
  const isHydratingRef = useRef(true);

  useEffect(() => {
    workbookRef.current = workbook;
  }, [workbook]);

  // Clear edit-only interactive states whenever switching back to View mode
  useEffect(() => {
    if (!isEditing) {
      setEditingCoord(null);
      setRenamingSheetId(null);
      setFillTarget(null);
      isFillingRef.current = false;
      setOpenDialog((prev) => (prev === 'export' ? prev : null));
    } else {
      setViewFindOpen(false);
    }
  }, [isEditing]);

  // Active sheet + evaluated values across entire workbook
  const activeSheet: SheetData = useMemo(() => {
    return (
      workbook.sheets.find((s) => s.id === workbook.activeSheetId) ??
      workbook.sheets[0]
    );
  }, [workbook]);

  const evaluatedWorkbook = useMemo(() => evaluateWorkbook(workbook), [workbook]);
  const activeEvalMap = useMemo(
    () => evaluatedWorkbook[activeSheet.id] ?? {},
    [evaluatedWorkbook, activeSheet.id],
  );

  const normSel = useMemo(() => normalizeSelection(selection), [selection]);
  const activeCoord = selection.start;
  const activeCellKey = cellKey(activeCoord.row, activeCoord.col);
  const activeCell = activeSheet.cells[activeCellKey];
  const activeStyle: CellStyle = activeCell?.style ?? {};

  const selectionRangeLabel = useMemo(() => {
    const s = formatCellRef(normSel.minRow, normSel.minCol);
    const e = formatCellRef(normSel.maxRow, normSel.maxCol);
    return s === e ? s : `${s}:${e}`;
  }, [normSel]);

  // Sync Name Box & formula bar draft when selection changes (outside inline edit)
  useEffect(() => {
    setNameBoxText(selectionRangeLabel);
    if (!editingCoord) {
      setDraftValue(activeSheet.cells[activeCellKey]?.raw ?? '');
    }
  }, [selectionRangeLabel, activeCellKey, activeSheet, editingCoord]);

  // Hydrate workbook from binary blob + persisted Dexie metadata
  useEffect(() => {
    let cancelled = false;
    isHydratingRef.current = true;
    setLoadingSheet(true);

    (async () => {
      try {
        const [blob, meta] = await Promise.all([
          loadFileBlob(tab.fileId),
          getDocMeta<PersistedSheetMeta>(tab.fileId),
        ]);
        if (cancelled) return;

        let loadedWb = createBlankWorkbook('Sheet1');
        if (blob && blob.size > 0) {
          const bytes = new Uint8Array(await blob.arrayBuffer());
          loadedWb = await xlsxBytesToWorkbookAsync(bytes, tab.extension);
        }
        const hydrated = applyPersistedMeta(loadedWb, meta);
        if (cancelled) return;

        setWorkbook(hydrated);
        setPast([]);
        setFuture([]);
        setViewColWidths({});
        setViewFindOpen(false);
        setViewFindQuery('');
        setViewFindIndex(0);
        setSelection({start: {row: 0, col: 0}, end: {row: 0, col: 0}});
        setEditingCoord(null);
        const isDesktop = typeof window === 'undefined' ? true : window.innerWidth >= 768;
        setChartsPanelOpen(
          isDesktop &&
            (hydrated.sheets.find((s) => s.id === hydrated.activeSheetId)?.charts.length ?? 0) > 0,
        );
      } catch (err) {
        if (!cancelled) {
          pushToast(
            err instanceof Error ? err.message : 'Unable to open spreadsheet.',
            'error',
          );
        }
      } finally {
        if (!cancelled) {
          isHydratingRef.current = false;
          setLoadingSheet(false);
        }
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [tab.fileId, tab.extension, tab.discardNonce, pushToast]);

  // Helper to mutate workbook with undo history
  const commitWorkbook = useCallback(
    (updater: (prev: WorkbookData) => WorkbookData, recordUndo = true) => {
      if (modeRef.current !== 'edit') return;
      setWorkbook((prev) => {
        const next = updater(prev);
        if (next === prev) return prev;
        if (recordUndo) {
          setPast((p) => [...p.slice(-49), prev]);
          setFuture([]);
        }
        return next;
      });
      if (!isHydratingRef.current) {
        markTabDirty(tab.fileId, true);
      }
    },
    [markTabDirty, tab.fileId],
  );

  const updateActiveSheet = useCallback(
    (updater: (sheet: SheetData) => SheetData, recordUndo = true) => {
      if (modeRef.current !== 'edit') return;
      commitWorkbook(
        (prev) => ({
          ...prev,
          sheets: prev.sheets.map((s) =>
            s.id === prev.activeSheetId ? updater(s) : s,
          ),
        }),
        recordUndo,
      );
    },
    [commitWorkbook],
  );

  const handleUndo = useCallback(() => {
    if (modeRef.current !== 'edit') return;
    setPast((p) => {
      if (p.length === 0) return p;
      const previous = p[p.length - 1];
      setFuture((f) => [workbookRef.current, ...f.slice(0, 49)]);
      setWorkbook(previous);
      markTabDirty(tab.fileId, true);
      return p.slice(0, -1);
    });
  }, [markTabDirty, tab.fileId]);

  const handleRedo = useCallback(() => {
    if (modeRef.current !== 'edit') return;
    setFuture((f) => {
      if (f.length === 0) return f;
      const next = f[0];
      setPast((p) => [...p.slice(-49), workbookRef.current]);
      setWorkbook(next);
      markTabDirty(tab.fileId, true);
      return f.slice(1);
    });
  }, [markTabDirty, tab.fileId]);

  // Persist binary .xlsx/.xls/.csv blob + persisted visual metadata
  const persistWorkbook = useCallback(async () => {
    if (modeRef.current !== 'edit') return;
    const wb = workbookRef.current;
    const ext = tab.extension.toLowerCase();
    let blob: Blob;
    if (ext === '.csv') {
      const sheet = wb.sheets.find((s) => s.id === wb.activeSheetId) ?? wb.sheets[0];
      const evals = evaluateWorkbook(wb)[sheet.id] ?? {};
      const csv = sheetToCsvString(sheet, evals);
      blob = new Blob([csv], {type: 'text/csv'});
    } else {
      const bytes = await workbookToXlsxBytesAsync(wb, ext === '.xls' ? 'xls' : 'xlsx');
      blob = new Blob([new Uint8Array(bytes)], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
    }

    const active = wb.sheets.find((s) => s.id === wb.activeSheetId) ?? wb.sheets[0];
    const evals = evaluateWorkbook(wb)[active.id] ?? {};
    const preview = sheetToCsvString(active, evals).slice(0, 200);

    await saveDocMeta<PersistedSheetMeta>(tab.fileId, extractPersistedMeta(wb));
    await saveTabBlob(tab.fileId, blob, preview);
  }, [tab.fileId, tab.extension, saveTabBlob]);

  useEffect(() => {
    registerBinarySaveHandler(tab.fileId, persistWorkbook);
    return () => registerBinarySaveHandler(tab.fileId, null);
  }, [tab.fileId, persistWorkbook, registerBinarySaveHandler]);

  // Commit cell raw text with Data Validation check
  const commitCellValue = useCallback(
    (row: number, col: number, raw: string) => {
      if (modeRef.current !== 'edit') return false;
      const rule = getValidationForCell(row, col, activeSheet.validations);
      const check = validateCellInput(raw, rule);
      if (!check.valid) {
        pushToast(check.message || 'Input rejected by Data Validation rule.', 'warning');
        return false;
      }

      const k = cellKey(row, col);
      const existing = activeSheet.cells[k];
      if ((existing?.raw ?? '') === raw) return true;

      updateActiveSheet((s) => {
        const nextCells = {...s.cells};
        if (raw === '' && !existing?.style) {
          delete nextCells[k];
        } else {
          nextCells[k] = {...existing, raw};
        }
        return {...s, cells: nextCells};
      });
      return true;
    },
    [activeSheet.cells, activeSheet.validations, updateActiveSheet, pushToast],
  );

  const finishEditing = useCallback(
    (move?: 'down' | 'up' | 'right' | 'left') => {
      if (editingCoord) {
        commitCellValue(editingCoord.row, editingCoord.col, draftValue);
        setEditingCoord(null);
      }
      if (move) {
        setSelection((prev) => {
          let r = prev.start.row;
          let c = prev.start.col;
          if (move === 'down') r = Math.min(activeSheet.rowCount - 1, r + 1);
          if (move === 'up') r = Math.max(0, r - 1);
          if (move === 'right') c = Math.min(activeSheet.colCount - 1, c + 1);
          if (move === 'left') c = Math.max(0, c - 1);
          return {start: {row: r, col: c}, end: {row: r, col: c}};
        });
      }
    },
    [editingCoord, draftValue, commitCellValue, activeSheet.rowCount, activeSheet.colCount],
  );

  const startEditingAt = useCallback(
    (coord: CellCoord, initialText?: string) => {
      if (modeRef.current !== 'edit') return;
      const k = cellKey(coord.row, coord.col);
      setEditingCoord(coord);
      setDraftValue(initialText !== undefined ? initialText : activeSheet.cells[k]?.raw ?? '');
      requestAnimationFrame(() => inlineInputRef.current?.focus());
    },
    [activeSheet.cells],
  );

  // Apply style patch across entire selection
  const handleApplyStyle = useCallback(
    (patch: Partial<CellStyle>) => {
      updateActiveSheet((s) => {
        const nextCells = {...s.cells};
        for (let r = normSel.minRow; r <= normSel.maxRow; r++) {
          for (let c = normSel.minCol; c <= normSel.maxCol; c++) {
            const k = cellKey(r, c);
            const curr = nextCells[k] ?? {raw: ''};
            nextCells[k] = {
              ...curr,
              style: {...curr.style, ...patch},
            };
          }
        }
        return {...s, cells: nextCells};
      });
    },
    [normSel, updateActiveSheet],
  );

  const handleStepDecimals = (delta: number) => {
    const currDec = activeStyle.decimals ?? 2;
    const nextDec = Math.max(0, Math.min(6, currDec + delta));
    handleApplyStyle({
      format:
        activeStyle.format && activeStyle.format !== 'general'
          ? activeStyle.format
          : 'number',
      decimals: nextDec,
    });
  };

  const activeMerge = useMemo(
    () => findMergeAt(activeCoord.row, activeCoord.col, activeSheet.merges),
    [activeCoord, activeSheet.merges],
  );

  const handleToggleMerge = () => {
    if (activeMerge) {
      updateActiveSheet((s) => ({
        ...s,
        merges: s.merges.filter(
          (m) =>
            !(
              m.startRow === activeMerge.startRow &&
              m.startCol === activeMerge.startCol
            ),
        ),
      }));
      pushToast('Unmerged cells', 'info');
      return;
    }
    if (normSel.minRow === normSel.maxRow && normSel.minCol === normSel.maxCol) {
      pushToast('Select 2 or more cells to merge.', 'info');
      return;
    }
    updateActiveSheet((s) => {
      // Remove any overlapping existing merges first
      const filtered = s.merges.filter(
        (m) =>
          m.endRow < normSel.minRow ||
          m.startRow > normSel.maxRow ||
          m.endCol < normSel.minCol ||
          m.startCol > normSel.maxCol,
      );
      const anchorKey = cellKey(normSel.minRow, normSel.minCol);
      const nextCells = {...s.cells};
      const anchor = nextCells[anchorKey] ?? {raw: ''};
      nextCells[anchorKey] = {
        ...anchor,
        style: {...anchor.style, align: 'center'},
      };
      return {
        ...s,
        cells: nextCells,
        merges: [
          ...filtered,
          {
            startRow: normSel.minRow,
            startCol: normSel.minCol,
            endRow: normSel.maxRow,
            endCol: normSel.maxCol,
          },
        ],
      };
    });
    pushToast('Merged & centered selection', 'success');
  };

  const handleClearSelection = useCallback(() => {
    updateActiveSheet((s) => {
      const nextCells = {...s.cells};
      for (let r = normSel.minRow; r <= normSel.maxRow; r++) {
        for (let c = normSel.minCol; c <= normSel.maxCol; c++) {
          delete nextCells[cellKey(r, c)];
        }
      }
      return {...s, cells: nextCells};
    });
  }, [normSel, updateActiveSheet]);

  const handleInsertRow = (where: 'above' | 'below') => {
    const insertIdx = where === 'above' ? activeCoord.row : activeCoord.row + 1;
    updateActiveSheet((s) => {
      const nextCells: typeof s.cells = {};
      for (const [k, cell] of Object.entries(s.cells)) {
        const [r, c] = k.split(',').map(Number);
        const nr = r >= insertIdx ? r + 1 : r;
        nextCells[cellKey(nr, c)] = cell;
      }
      return {...s, rowCount: s.rowCount + 1, cells: nextCells};
    });
  };

  const handleDeleteRow = () => {
    if (activeSheet.rowCount <= 1) return;
    const delIdx = activeCoord.row;
    updateActiveSheet((s) => {
      const nextCells: typeof s.cells = {};
      for (const [k, cell] of Object.entries(s.cells)) {
        const [r, c] = k.split(',').map(Number);
        if (r === delIdx) continue;
        nextCells[cellKey(r > delIdx ? r - 1 : r, c)] = cell;
      }
      return {...s, rowCount: Math.max(10, s.rowCount - 1), cells: nextCells};
    });
  };

  const handleInsertCol = (where: 'left' | 'right') => {
    const insertIdx = where === 'left' ? activeCoord.col : activeCoord.col + 1;
    updateActiveSheet((s) => {
      const nextCells: typeof s.cells = {};
      for (const [k, cell] of Object.entries(s.cells)) {
        const [r, c] = k.split(',').map(Number);
        const nc = c >= insertIdx ? c + 1 : c;
        nextCells[cellKey(r, nc)] = cell;
      }
      return {...s, colCount: s.colCount + 1, cells: nextCells};
    });
  };

  const handleDeleteCol = () => {
    if (activeSheet.colCount <= 1) return;
    const delIdx = activeCoord.col;
    updateActiveSheet((s) => {
      const nextCells: typeof s.cells = {};
      for (const [k, cell] of Object.entries(s.cells)) {
        const [r, c] = k.split(',').map(Number);
        if (c === delIdx) continue;
        nextCells[cellKey(r, c > delIdx ? c - 1 : c)] = cell;
      }
      return {...s, colCount: Math.max(6, s.colCount - 1), cells: nextCells};
    });
  };

  const handleInsertFunction = (fnName: string) => {
    // Auto-detect numeric range above active cell if single cell is selected
    let rangeArg = selectionRangeLabel;
    if (normSel.minRow === normSel.maxRow && normSel.minCol === normSel.maxCol) {
      let top = activeCoord.row - 1;
      while (top >= 0) {
        const v = activeEvalMap[cellKey(top, activeCoord.col)]?.value;
        if (typeof v === 'number') top--;
        else break;
      }
      const firstNumRow = top + 1;
      if (firstNumRow < activeCoord.row) {
        rangeArg = `${formatCellRef(firstNumRow, activeCoord.col)}:${formatCellRef(activeCoord.row - 1, activeCoord.col)}`;
      } else {
        rangeArg = 'A1:A5';
      }
    }

    const templateMap: Record<string, string> = {
      SUM: `=SUM(${rangeArg})`,
      AVERAGE: `=AVERAGE(${rangeArg})`,
      MIN: `=MIN(${rangeArg})`,
      MAX: `=MAX(${rangeArg})`,
      COUNT: `=COUNT(${rangeArg})`,
      COUNTIF: `=COUNTIF(${rangeArg},">0")`,
      IF: `=IF(${formatCellRef(activeCoord.row, Math.max(0, activeCoord.col - 1))}>0,"Yes","No")`,
      VLOOKUP: `=VLOOKUP(A2,A2:C10,2,FALSE)`,
      XLOOKUP: `=XLOOKUP(A2,A2:A10,B2:B10,"Not found")`,
      PMT: `=PMT(0.05/12,12,1200)`,
    };

    const formula = templateMap[fnName] ?? `=${fnName}(${rangeArg})`;
    if (fnName === 'SUM' && rangeArg !== 'A1:A5') {
      commitCellValue(activeCoord.row, activeCoord.col, formula);
    } else {
      startEditingAt(activeCoord, formula);
    }
  };

  const handleSort = (dir: 'asc' | 'desc') => {
    updateActiveSheet((s) => sortSheetByColumn(s, activeEvalMap, activeCoord.col, dir, true));
    pushToast(`Sorted column ${colIndexToLetter(activeCoord.col)} (${dir === 'asc' ? 'A→Z' : 'Z→A'})`, 'success');
  };

  const handleSetFreeze = (mode: 'top-row' | 'first-col' | 'selection' | 'none') => {
    updateActiveSheet((s) => {
      if (mode === 'top-row') return {...s, freezeRows: 1, freezeCols: 0};
      if (mode === 'first-col') return {...s, freezeRows: 0, freezeCols: 1};
      if (mode === 'selection') return {...s, freezeRows: activeCoord.row, freezeCols: activeCoord.col};
      return {...s, freezeRows: 0, freezeCols: 0};
    });
    pushToast(mode === 'none' ? 'Unfroze panes' : 'Freeze panes updated', 'info');
  };

  // Fill handle drag completion
  useEffect(() => {
    const onWindowMouseUp = () => {
      isSelectingRef.current = false;
      if (modeRef.current === 'edit' && isFillingRef.current && fillTarget) {
        isFillingRef.current = false;
        const target = fillTarget;
        setFillTarget(null);
        const srcMinR = normSel.minRow;
        const srcMaxR = normSel.maxRow;
        const srcMinC = normSel.minCol;
        const srcMaxC = normSel.maxCol;
        const srcH = srcMaxR - srcMinR + 1;
        const srcW = srcMaxC - srcMinC + 1;

        const destMinR = Math.min(srcMinR, target.row);
        const destMaxR = Math.max(srcMaxR, target.row);
        const destMinC = Math.min(srcMinC, target.col);
        const destMaxC = Math.max(srcMaxC, target.col);

        if (destMaxR > srcMaxR || destMaxC > srcMaxC) {
          updateActiveSheet((s) => {
            const nextCells = {...s.cells};
            for (let r = destMinR; r <= destMaxR; r++) {
              for (let c = destMinC; c <= destMaxC; c++) {
                if (r >= srcMinR && r <= srcMaxR && c >= srcMinC && c <= srcMaxC) continue;
                const srcR = srcMinR + ((r - srcMinR) % srcH);
                const srcC = srcMinC + ((c - srcMinC) % srcW);
                const sourceCell = s.cells[cellKey(srcR, srcC)];
                if (!sourceCell) continue;
                const dR = r - srcR;
                const dC = c - srcC;
                const nextRaw = sourceCell.raw.startsWith('=')
                  ? shiftFormulaReferences(sourceCell.raw, dR, dC)
                  : sourceCell.raw;
                nextCells[cellKey(r, c)] = {
                  raw: nextRaw,
                  style: sourceCell.style ? {...sourceCell.style} : undefined,
                };
              }
            }
            return {...s, cells: nextCells};
          });
          setSelection({
            start: {row: destMinR, col: destMinC},
            end: {row: destMaxR, col: destMaxC},
          });
        }
      } else {
        isFillingRef.current = false;
        setFillTarget(null);
      }
    };

    window.addEventListener('mouseup', onWindowMouseUp);
    return () => window.removeEventListener('mouseup', onWindowMouseUp);
  }, [fillTarget, normSel, updateActiveSheet]);

  // Global keyboard shortcuts inside Sheets workspace
  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      const targetEl = e.target as HTMLElement | null;
      const inInput =
        targetEl?.tagName === 'INPUT' ||
        targetEl?.tagName === 'TEXTAREA' ||
        targetEl?.tagName === 'SELECT';
      const editingNow = modeRef.current === 'edit';
      const key = e.key.toLowerCase();

      if ((e.ctrlKey || e.metaKey) && key === 's') {
        if (!editingNow) return;
        e.preventDefault();
        void persistWorkbook().then(() => pushToast('Workbook saved', 'success'));
        return;
      }
      if ((e.ctrlKey || e.metaKey) && (key === 'h' || key === 'f')) {
        e.preventDefault();
        if (!editingNow) {
          setViewFindOpen(true);
          requestAnimationFrame(() => viewFindInputRef.current?.focus());
        } else {
          setOpenDialog('find');
        }
        return;
      }

      if (e.key === 'Escape' && !editingNow && viewFindOpen) {
        e.preventDefault();
        setViewFindOpen(false);
        return;
      }

      if (inInput || openDialog) return;

      if ((e.ctrlKey || e.metaKey) && key === 'c') {
        const lines: string[] = [];
        for (let r = normSel.minRow; r <= normSel.maxRow; r++) {
          const vals: string[] = [];
          for (let c = normSel.minCol; c <= normSel.maxCol; c++) {
            vals.push(activeSheet.cells[cellKey(r, c)]?.raw ?? '');
          }
          lines.push(vals.join('\t'));
        }
        void navigator.clipboard?.writeText(lines.join('\n')).catch(() => undefined);
        return;
      }

      if (e.key === 'ArrowDown' || e.key === 'ArrowUp' || e.key === 'ArrowLeft' || e.key === 'ArrowRight') {
        e.preventDefault();
        const dr = e.key === 'ArrowDown' ? 1 : e.key === 'ArrowUp' ? -1 : 0;
        const dc = e.key === 'ArrowRight' ? 1 : e.key === 'ArrowLeft' ? -1 : 0;
        setSelection((prev) => {
          const base = e.shiftKey ? prev.end : prev.start;
          const next = {
            row: Math.max(0, Math.min(activeSheet.rowCount - 1, base.row + dr)),
            col: Math.max(0, Math.min(activeSheet.colCount - 1, base.col + dc)),
          };
          return e.shiftKey ? {start: prev.start, end: next} : {start: next, end: next};
        });
        return;
      }
      if (e.key === 'Tab') {
        e.preventDefault();
        const dc = e.shiftKey ? -1 : 1;
        setSelection((prev) => {
          const nextCol = Math.max(0, Math.min(activeSheet.colCount - 1, prev.start.col + dc));
          return {start: {row: prev.start.row, col: nextCol}, end: {row: prev.start.row, col: nextCol}};
        });
        return;
      }

      // View mode is strictly read-only: block cell editing, paste, delete, undo/redo, and formatting
      if (!editingNow) return;

      if ((e.ctrlKey || e.metaKey) && key === 'z') {
        e.preventDefault();
        if (e.shiftKey) handleRedo();
        else handleUndo();
        return;
      }
      if ((e.ctrlKey || e.metaKey) && key === 'y') {
        e.preventDefault();
        handleRedo();
        return;
      }
      if ((e.ctrlKey || e.metaKey) && key === 'b') {
        e.preventDefault();
        handleApplyStyle({bold: !activeStyle.bold});
        return;
      }
      if ((e.ctrlKey || e.metaKey) && key === 'i') {
        e.preventDefault();
        handleApplyStyle({italic: !activeStyle.italic});
        return;
      }
      if ((e.ctrlKey || e.metaKey) && key === 'u') {
        e.preventDefault();
        handleApplyStyle({underline: !activeStyle.underline});
        return;
      }
      if ((e.ctrlKey || e.metaKey) && key === 'v') {
        void navigator.clipboard
          ?.readText()
          .then((clip) => {
            if (!clip || modeRef.current !== 'edit') return;
            const rows = clip.replace(/\r\n/g, '\n').split('\n');
            updateActiveSheet((s) => {
              const nextCells = {...s.cells};
              rows.forEach((line, rOff) => {
                line.split('\t').forEach((val, cOff) => {
                  const tr = activeCoord.row + rOff;
                  const tc = activeCoord.col + cOff;
                  if (tr < s.rowCount && tc < s.colCount) {
                    const k = cellKey(tr, tc);
                    nextCells[k] = {...nextCells[k], raw: val};
                  }
                });
              });
              return {...s, cells: nextCells};
            });
          })
          .catch(() => undefined);
        return;
      }
      if (e.key === 'Enter' || e.key === 'F2') {
        e.preventDefault();
        startEditingAt(activeCoord);
        return;
      }
      if (e.key === 'Delete' || e.key === 'Backspace') {
        e.preventDefault();
        handleClearSelection();
        return;
      }
      if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
        e.preventDefault();
        startEditingAt(activeCoord, e.key);
      }
    };

    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, [
    openDialog,
    viewFindOpen,
    persistWorkbook,
    pushToast,
    handleUndo,
    handleRedo,
    handleApplyStyle,
    activeStyle,
    normSel,
    activeSheet,
    activeCoord,
    updateActiveSheet,
    startEditingAt,
    handleClearSelection,
  ]);

  const getColWidth = useCallback(
    (colIdx: number) => {
      const viewOverride = !isEditing ? viewColWidths[activeSheet.id]?.[colIdx] : undefined;
      return viewOverride ?? activeSheet.colWidths[colIdx] ?? DEFAULT_COL_WIDTH;
    },
    [activeSheet.colWidths, activeSheet.id, isEditing, viewColWidths],
  );

  // Zoom-scaled geometry (native CSS dimensions keep position: sticky freeze panes accurate at every zoom level)
  const scaledRowHeaderW = Math.round(ROW_HEADER_WIDTH * zoom);
  const scaledColHeaderH = Math.round(COL_HEADER_HEIGHT * zoom);
  const getScaledColWidth = useCallback(
    (colIdx: number) => Math.max(40, Math.round(getColWidth(colIdx) * zoom)),
    [getColWidth, zoom],
  );
  const getScaledRowHeight = useCallback(
    (rowIdx: number) =>
      Math.max(18, Math.round((activeSheet.rowHeights[rowIdx] ?? DEFAULT_ROW_HEIGHT) * zoom)),
    [activeSheet.rowHeights, zoom],
  );

  // Column resize drag (allowed in View mode for reading only, without mutating workbook)
  const startColResize = (colIdx: number, startX: number) => {
    const sheetId = activeSheet.id;
    const startW = getColWidth(colIdx);
    const onMove = (ev: MouseEvent) => {
      const nextW = Math.max(54, Math.min(360, Math.round(startW + (ev.clientX - startX) / zoom)));
      if (modeRef.current === 'edit') {
        updateActiveSheet(
          (s) => ({...s, colWidths: {...s.colWidths, [colIdx]: nextW}}),
          false,
        );
      } else {
        setViewColWidths((prev) => ({
          ...prev,
          [sheetId]: {...(prev[sheetId] ?? {}), [colIdx]: nextW},
        }));
      }
    };
    const onUp = () => {
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
    };
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
  };

  // Row resize drag (edit mode only)
  const startRowResize = (rowIdx: number, startY: number) => {
    if (modeRef.current !== 'edit') return;
    const startH = activeSheet.rowHeights[rowIdx] ?? DEFAULT_ROW_HEIGHT;
    const onMove = (ev: MouseEvent) => {
      const nextH = Math.max(22, Math.min(140, Math.round(startH + (ev.clientY - startY) / zoom)));
      updateActiveSheet(
        (s) => ({...s, rowHeights: {...s.rowHeights, [rowIdx]: nextH}}),
        false,
      );
    };
    const onUp = () => {
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
    };
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
  };

  // Auto-Filter visible row mask
  const visibleRows = useMemo(() => {
    const activeFilters = Object.entries(activeSheet.columnFilters).filter(
      ([, allowed]) => allowed && allowed.length > 0,
    );
    const list: number[] = [];
    for (let r = 0; r < activeSheet.rowCount; r++) {
      if (r === 0 || !activeSheet.autoFilterEnabled || activeFilters.length === 0) {
        list.push(r);
        continue;
      }
      const matchesAll = activeFilters.every(([colStr, allowed]) => {
        const display = activeEvalMap[cellKey(r, Number(colStr))]?.display ?? '';
        return allowed.includes(display);
      });
      if (matchesAll) list.push(r);
    }
    return list;
  }, [activeSheet.rowCount, activeSheet.autoFilterEnabled, activeSheet.columnFilters, activeEvalMap]);

  // Double-click column resize handle to auto-fit column width (viewing-only in View mode)
  const handleAutoFitCol = useCallback(
    (colIdx: number) => {
      const idealW = computeAutoFitColWidth(activeSheet, activeEvalMap, colIdx, visibleRows);
      if (modeRef.current === 'edit') {
        updateActiveSheet(
          (s) => ({...s, colWidths: {...s.colWidths, [colIdx]: idealW}}),
          false,
        );
      } else {
        const sheetId = activeSheet.id;
        setViewColWidths((prev) => ({
          ...prev,
          [sheetId]: {...(prev[sheetId] ?? {}), [colIdx]: idealW},
        }));
      }
    },
    [activeEvalMap, activeSheet, updateActiveSheet, visibleRows],
  );

  // Unique values for the currently open Auto-Filter column dropdown
  const filterColumnOptions = useMemo(() => {
    if (openFilterCol === null) return [];
    const set = new Set<string>();
    for (let r = 1; r < activeSheet.rowCount; r++) {
      const disp = activeEvalMap[cellKey(r, openFilterCol)]?.display ?? '';
      if (disp !== '') set.add(disp);
    }
    return Array.from(set).sort((a, b) => a.localeCompare(b, 'en-US', {numeric: true}));
  }, [openFilterCol, activeSheet.rowCount, activeEvalMap]);

  // Cumulative zoom-scaled sticky offsets for Freeze Panes
  const colLeftOffsets = useMemo(
    () => computeStickyOffsets(activeSheet.colCount, getScaledColWidth, scaledRowHeaderW),
    [activeSheet.colCount, getScaledColWidth, scaledRowHeaderW],
  );

  const rowTopOffsets = useMemo(
    () => computeVisibleRowTopOffsets(visibleRows, getScaledRowHeight, scaledColHeaderH),
    [visibleRows, getScaledRowHeight, scaledColHeaderH],
  );

  const totalTableWidth = useMemo(() => {
    let w = scaledRowHeaderW;
    for (let c = 0; c < activeSheet.colCount; c++) {
      w += getScaledColWidth(c);
    }
    return w;
  }, [activeSheet.colCount, getScaledColWidth, scaledRowHeaderW]);

  // Scroll the grid container so a target cell is cleanly visible beyond any frozen panes
  const scrollCellIntoView = useCallback(
    (coord: CellCoord) => {
      const container = gridContainerRef.current;
      if (!container) return;

      const cellLeft = colLeftOffsets[coord.col] ?? scaledRowHeaderW;
      const cellW = getScaledColWidth(coord.col);
      const cellTop = rowTopOffsets[coord.row] ?? scaledColHeaderH;
      const cellH = getScaledRowHeight(coord.row);

      let frozenW = scaledRowHeaderW;
      for (let c = 0; c < activeSheet.freezeCols && c < activeSheet.colCount; c++) {
        frozenW += getScaledColWidth(c);
      }
      let frozenH = scaledColHeaderH;
      for (const r of visibleRows) {
        if (r < activeSheet.freezeRows) frozenH += getScaledRowHeight(r);
      }

      if (coord.col >= activeSheet.freezeCols) {
        const visLeft = container.scrollLeft + frozenW;
        const visRight = container.scrollLeft + container.clientWidth;
        if (cellLeft < visLeft) {
          container.scrollLeft = Math.max(0, cellLeft - frozenW - 16);
        } else if (cellLeft + cellW > visRight) {
          container.scrollLeft = Math.max(0, cellLeft + cellW - container.clientWidth + 24);
        }
      }

      if (coord.row >= activeSheet.freezeRows) {
        const visTop = container.scrollTop + frozenH;
        const visBottom = container.scrollTop + container.clientHeight;
        if (cellTop < visTop) {
          container.scrollTop = Math.max(0, cellTop - frozenH - 16);
        } else if (cellTop + cellH > visBottom) {
          container.scrollTop = Math.max(0, cellTop + cellH - container.clientHeight + 24);
        }
      }
    },
    [
      activeSheet.colCount,
      activeSheet.freezeCols,
      activeSheet.freezeRows,
      colLeftOffsets,
      getScaledColWidth,
      getScaledRowHeight,
      rowTopOffsets,
      scaledColHeaderH,
      scaledRowHeaderW,
      visibleRows,
    ],
  );

  // Unified sheet tab switcher (strictly local view state — never dirties or writes file)
  const handleSelectSheet = useCallback(
    (sheetId: string) => {
      setWorkbook((prev) => (prev.activeSheetId === sheetId ? prev : {...prev, activeSheetId: sheetId}));
      setSelection({start: {row: 0, col: 0}, end: {row: 0, col: 0}});
      setEditingCoord(null);
      setOpenFilterCol(null);
      findCursorRef.current = 0;
      setViewFindIndex(0);
      if (gridContainerRef.current) {
        gridContainerRef.current.scrollTo({top: 0, left: 0});
      }
      const nextSheet = workbookRef.current.sheets.find((s) => s.id === sheetId);
      const isDesktop = typeof window === 'undefined' ? true : window.innerWidth >= 768;
      if (nextSheet) {
        setChartsPanelOpen(isDesktop && nextSheet.charts.length > 0);
      }
    },
    [],
  );

  // Live View-mode search matches & highlighting
  const viewFindMatches = useMemo(() => {
    if (!viewFindOpen || !viewFindQuery) return [];
    return findSheetMatches(
      activeSheet,
      activeEvalMap,
      viewFindQuery,
      viewFindCaseSensitive,
      visibleRows,
    );
  }, [activeEvalMap, activeSheet, viewFindCaseSensitive, viewFindOpen, viewFindQuery, visibleRows]);

  const viewFindMatchKeySet = useMemo(() => {
    const set = new Set<string>();
    for (const m of viewFindMatches) set.add(cellKey(m.row, m.col));
    return set;
  }, [viewFindMatches]);

  const activeViewMatchCoord =
    viewFindMatches.length > 0
      ? viewFindMatches[viewFindIndex % viewFindMatches.length]
      : null;

  const stepViewFind = useCallback(
    (delta: number) => {
      if (viewFindMatches.length === 0) return;
      const nextIdx =
        ((viewFindIndex + delta) % viewFindMatches.length + viewFindMatches.length) %
        viewFindMatches.length;
      setViewFindIndex(nextIdx);
      const target = viewFindMatches[nextIdx];
      setSelection({start: target, end: target});
      scrollCellIntoView(target);
    },
    [scrollCellIntoView, viewFindIndex, viewFindMatches],
  );

  // When View-mode query changes, jump immediately to the first match
  useEffect(() => {
    if (!viewFindOpen || !viewFindQuery) {
      setViewFindIndex(0);
      return;
    }
    setViewFindIndex(0);
    if (viewFindMatches.length > 0) {
      const first = viewFindMatches[0];
      setSelection({start: first, end: first});
      scrollCellIntoView(first);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [viewFindQuery, viewFindCaseSensitive, viewFindOpen, activeSheet.id]);

  // Live selection summary stats (SUM, AVG, COUNT)
  const selectionStats = useMemo(() => {
    const nums: number[] = [];
    let countAll = 0;
    for (let r = normSel.minRow; r <= normSel.maxRow; r++) {
      for (let c = normSel.minCol; c <= normSel.maxCol; c++) {
        const ev = activeEvalMap[cellKey(r, c)];
        if (ev && ev.value !== null && ev.value !== '') {
          countAll++;
          if (typeof ev.value === 'number' && Number.isFinite(ev.value)) {
            nums.push(ev.value);
          }
        }
      }
    }
    const sum = nums.reduce((a, b) => a + b, 0);
    const avg = nums.length > 0 ? sum / nums.length : 0;
    return {countAll, numCount: nums.length, sum, avg};
  }, [normSel, activeEvalMap]);

  // Sheet Tab actions
  const handleAddSheet = () => {
    const idx = workbook.sheets.length + 1;
    let name = `Sheet${idx}`;
    while (workbook.sheets.some((s) => s.name.toLowerCase() === name.toLowerCase())) {
      name = `Sheet${Math.floor(Math.random() * 900 + 100)}`;
    }
    const fresh = createEmptySheet(`sheet_${Date.now()}`, name);
    commitWorkbook((prev) => ({
      activeSheetId: fresh.id,
      sheets: [...prev.sheets, fresh],
    }));
    setSelection({start: {row: 0, col: 0}, end: {row: 0, col: 0}});
  };

  const handleDuplicateSheet = (sheetId: string) => {
    const src = workbook.sheets.find((s) => s.id === sheetId);
    if (!src) return;
    const copy: SheetData = {
      ...JSON.parse(JSON.stringify(src)),
      id: `sheet_${Date.now()}`,
      name: `${src.name} Copy`.slice(0, 31),
    };
    commitWorkbook((prev) => ({
      activeSheetId: copy.id,
      sheets: [...prev.sheets, copy],
    }));
  };

  const handleDeleteSheet = (sheetId: string) => {
    if (workbook.sheets.length <= 1) {
      pushToast('A workbook must keep at least one sheet.', 'info');
      return;
    }
    commitWorkbook((prev) => {
      const remaining = prev.sheets.filter((s) => s.id !== sheetId);
      return {
        activeSheetId:
          prev.activeSheetId === sheetId ? remaining[0].id : prev.activeSheetId,
        sheets: remaining,
      };
    });
  };

  const commitSheetRename = () => {
    if (!renamingSheetId) return;
    const clean = renameDraft.trim().slice(0, 31);
    if (clean) {
      commitWorkbook((prev) => ({
        ...prev,
        sheets: prev.sheets.map((s) =>
          s.id === renamingSheetId ? {...s, name: clean} : s,
        ),
      }));
    }
    setRenamingSheetId(null);
  };

  // Find & Replace callbacks (used by Edit-mode dialog)
  const handleFindStep = (query: string, caseSensitive: boolean, delta: 1 | -1): string => {
    const matches = findSheetMatches(activeSheet, activeEvalMap, query, caseSensitive, visibleRows);
    if (matches.length === 0) return 'No matching cells found.';
    const idx =
      ((findCursorRef.current % matches.length) + matches.length) % matches.length;
    const target = matches[idx];
    findCursorRef.current = ((idx + delta) % matches.length + matches.length) % matches.length;
    setSelection({start: target, end: target});
    scrollCellIntoView(target);
    return `Selected ${formatCellRef(target.row, target.col)} (${idx + 1} of ${matches.length})`;
  };

  const handleFindNext = (query: string, caseSensitive: boolean): string =>
    handleFindStep(query, caseSensitive, 1);

  const handleFindPrev = (query: string, caseSensitive: boolean): string => {
    const matches = findSheetMatches(activeSheet, activeEvalMap, query, caseSensitive, visibleRows);
    if (matches.length === 0) return 'No matching cells found.';
    // Step back from the last visited index
    findCursorRef.current =
      ((findCursorRef.current - 2) % matches.length + matches.length) % matches.length;
    return handleFindStep(query, caseSensitive, 1);
  };

  const handleReplaceAll = (
    query: string,
    replacement: string,
    caseSensitive: boolean,
  ): string => {
    if (modeRef.current !== 'edit') return 'Switch to Edit mode to replace cells.';
    let count = 0;
    const flags = caseSensitive ? 'g' : 'gi';
    const escaped = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const regex = new RegExp(escaped, flags);
    updateActiveSheet((s) => {
      const nextCells = {...s.cells};
      for (const [k, cell] of Object.entries(s.cells)) {
        if (regex.test(cell.raw)) {
          regex.lastIndex = 0;
          nextCells[k] = {...cell, raw: cell.raw.replace(regex, replacement)};
          count++;
        }
      }
      return {...s, cells: nextCells};
    });
    return count > 0
      ? `Replaced matches in ${count} cell${count === 1 ? '' : 's'}.`
      : 'No matching cells found.';
  };

  // Export & Print handlers
  const handleExportFormat = async (
    format: '.xlsx' | '.xls' | '.csv' | '.pdf',
    target: 'device' | 'vault',
  ) => {
    const baseName = tab.name.replace(/\.[^.]+$/, '');
    const outName = `${baseName}${format}`;

    if (format === '.pdf') {
      const pdfBytes = await workbookToPdfBytes(workbookRef.current, baseName);
      if (target === 'vault') {
        await createFileFromBytes(outName, 'pdf', pdfBytes, fileRecord?.folderId ?? null, true);
      } else {
        const blob = new Blob([new Uint8Array(pdfBytes)], {type: 'application/pdf'});
        await saveBlobToDevice(blob, outName, 'pdf');
        pushToast(`Exported ${outName}`, 'success');
      }
      return;
    }

    let bytes: Uint8Array;
    let mime: string;
    if (format === '.csv') {
      const csv = sheetToCsvString(activeSheet, activeEvalMap);
      bytes = new TextEncoder().encode(csv);
      mime = 'text/csv';
    } else {
      bytes = await workbookToXlsxBytesAsync(
        workbookRef.current,
        format === '.xls' ? 'xls' : 'xlsx',
      );
      mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    }

    if (target === 'vault') {
      await createFileFromBytes(outName, 'sheets', bytes, fileRecord?.folderId ?? null, false);
    } else {
      await saveBlobToDevice(new Blob([new Uint8Array(bytes)], {type: mime}), outName, 'sheets');
      pushToast(`Saved ${outName}`, 'success');
    }
  };

  const handlePrint = () => {
    const iframe = document.createElement('iframe');
    iframe.style.position = 'fixed';
    iframe.style.right = '0';
    iframe.style.bottom = '0';
    iframe.style.width = '0';
    iframe.style.height = '0';
    iframe.style.border = '0';
    document.body.appendChild(iframe);

    let maxR = 0;
    let maxC = 0;
    for (const [k, cell] of Object.entries(activeSheet.cells)) {
      if (!cell.raw) continue;
      const [r, c] = k.split(',').map(Number);
      if (r > maxR) maxR = r;
      if (c > maxC) maxC = c;
    }

    const rowsHtml: string[] = [];
    for (let r = 0; r <= maxR; r++) {
      const cellsHtml: string[] = [];
      for (let c = 0; c <= maxC; c++) {
        const disp = activeEvalMap[cellKey(r, c)]?.display ?? '';
        const tag = r === 0 ? 'th' : 'td';
        cellsHtml.push(`<${tag}>${disp.replace(/</g, '&lt;')}</${tag}>`);
      }
      rowsHtml.push(`<tr>${cellsHtml.join('')}</tr>`);
    }

    const win = iframe.contentWindow;
    if (win) {
      win.document.open();
      win.document.write(`<!doctype html><html><head><title>${tab.name}</title>
<style>@page{size:landscape;margin:12mm}body{font-family:Calibri,system-ui,sans-serif;font-size:12px}h2{margin:0 0 10px;color:#065f46}table{border-collapse:collapse;width:100%}th,td{border:1px solid #cbd5e1;padding:5px 8px;text-align:left}th{background:#f1f5f9;font-weight:600}</style>
</head><body><h2>${tab.name} — ${activeSheet.name}</h2><table>${rowsHtml.join('')}</table></body></html>`);
      win.document.close();
      win.focus();
      setTimeout(() => {
        win.print();
        setTimeout(() => iframe.remove(), 1200);
      }, 150);
    }
  };

  // Active cell validation rule (for inline dropdown support)
  const activeValidation = useMemo(
    () => getValidationForCell(activeCoord.row, activeCoord.col, activeSheet.validations),
    [activeCoord, activeSheet.validations],
  );

  return (
    <div className="flex h-full min-h-0 flex-col bg-slate-100 dark:bg-slate-950">
      {/* Top Bar: Slim View Mode Toolbar OR Full 4-Tab Ribbon */}
      {isEditing ? (
        <SheetsRibbon
          activeTab={ribbonTab}
          onChangeTab={setRibbonTab}
          activeStyle={activeStyle}
          canUndo={past.length > 0}
          canRedo={future.length > 0}
          isFavorite={Boolean(fileRecord?.isFavorite)}
          isMerged={Boolean(activeMerge)}
          autoFilterEnabled={activeSheet.autoFilterEnabled}
          freezeRows={activeSheet.freezeRows}
          freezeCols={activeSheet.freezeCols}
          showGridlines={showGridlines}
          showFormulaBar={showFormulaBar}
          zoom={zoom}
          chartCount={activeSheet.charts.length}
          chartsPanelOpen={chartsPanelOpen}
          onUndo={handleUndo}
          onRedo={handleRedo}
          onSaveNow={() =>
            void persistWorkbook().then(() => pushToast('Workbook saved', 'success'))
          }
          onToggleFavorite={onToggleFavorite}
          onExportPdf={() => void handleExportFormat('.pdf', 'device')}
          onPrint={handlePrint}
          onOpenDialog={(kind, chartKind) => {
            if (chartKind) setDefaultChartKind(chartKind);
            setOpenDialog(kind);
          }}
          onApplyStyle={handleApplyStyle}
          onStepDecimals={handleStepDecimals}
          onToggleMerge={handleToggleMerge}
          onClearSelection={handleClearSelection}
          onInsertRow={handleInsertRow}
          onDeleteRow={handleDeleteRow}
          onInsertCol={handleInsertCol}
          onDeleteCol={handleDeleteCol}
          onInsertFunction={handleInsertFunction}
          onSort={handleSort}
          onToggleAutoFilter={() => {
            updateActiveSheet((s) => ({
              ...s,
              autoFilterEnabled: !s.autoFilterEnabled,
              columnFilters: s.autoFilterEnabled ? {} : s.columnFilters,
            }));
            setOpenFilterCol(null);
          }}
          onSetFreeze={handleSetFreeze}
          onToggleGridlines={() => setShowGridlines((v) => !v)}
          onToggleFormulaBar={() => setShowFormulaBar((v) => !v)}
          onChangeZoom={setZoom}
          onToggleChartsPanel={() => setChartsPanelOpen((v) => !v)}
          onDoneEditing={onDoneEditing}
        />
      ) : (
        <ViewModeToolbar
          fileName={tab.name}
          moduleType="sheets"
          zoomLabel={`${Math.round(zoom * 100)}%`}
          onZoomOut={() => setZoom((z) => Math.max(0.75, Number((z - 0.1).toFixed(2))))}
          onZoomIn={() => setZoom((z) => Math.min(1.5, Number((z + 0.1).toFixed(2))))}
          onZoomReset={() => setZoom(1)}
          isFavorite={Boolean(fileRecord?.isFavorite)}
          onToggleFavorite={onToggleFavorite}
          onSearch={() => {
            setViewFindOpen((open) => {
              const next = !open;
              if (next) requestAnimationFrame(() => viewFindInputRef.current?.focus());
              return next;
            });
          }}
          searchActive={viewFindOpen}
          onPrint={handlePrint}
          onExport={() => setOpenDialog('export')}
          exportLabel="Export or download workbook"
          onEdit={onEnterEdit}
          navigationSlot={
            workbook.sheets.length > 1 ? (
              <select
                aria-label="Select worksheet"
                value={activeSheet.id}
                onChange={(e) => handleSelectSheet(e.target.value)}
                className="h-8 max-w-[140px] rounded-lg border border-slate-200 bg-slate-50 px-2 text-xs font-semibold text-slate-700 dark:border-slate-700 dark:bg-slate-950 dark:text-slate-200"
              >
                {workbook.sheets.map((s) => (
                  <option key={s.id} value={s.id}>
                    {s.name}
                  </option>
                ))}
              </select>
            ) : (
              <span className="hidden whitespace-nowrap rounded-lg border border-slate-200 bg-slate-50 px-2 py-1 text-xs font-medium text-slate-700 sm:inline-block dark:border-slate-700 dark:bg-slate-950 dark:text-slate-200">
                {activeSheet.name}
              </span>
            )
          }
          extraActionsSlot={
            activeSheet.charts.length > 0 ? (
              <button
                type="button"
                onClick={() => setChartsPanelOpen((v) => !v)}
                aria-label="Toggle charts panel"
                aria-pressed={chartsPanelOpen}
                className={`inline-flex h-8 shrink-0 items-center gap-1 whitespace-nowrap rounded-lg border px-2 text-xs font-semibold transition-colors ${
                  chartsPanelOpen
                    ? 'border-emerald-600 bg-emerald-50 text-emerald-800 dark:border-emerald-500 dark:bg-emerald-950/60 dark:text-emerald-300'
                    : 'border-slate-200 text-slate-700 hover:bg-slate-100 dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800'
                }`}
              >
                <BarChart3 className="h-3.5 w-3.5" />
                <span className="sm:hidden">{activeSheet.charts.length}</span>
                <span className="hidden sm:inline">Charts ({activeSheet.charts.length})</span>
              </button>
            ) : undefined
          }
        />
      )}

      {/* Name Box + Formula Bar (Edit Mode only) */}
      {isEditing && showFormulaBar && (
        <div className="flex items-center gap-1.5 border-b border-slate-200 bg-white px-2.5 py-1.5 dark:border-slate-800 dark:bg-slate-900">
          <form
            onSubmit={(e) => {
              e.preventDefault();
              const parsed = parseCellRef(nameBoxText.split(':')[0]);
              if (parsed && parsed.row < activeSheet.rowCount && parsed.col < activeSheet.colCount) {
                setSelection({
                  start: {row: parsed.row, col: parsed.col},
                  end: {row: parsed.row, col: parsed.col},
                });
              } else {
                setNameBoxText(selectionRangeLabel);
              }
            }}
          >
            <input
              type="text"
              aria-label="Active cell name box"
              value={nameBoxText}
              onChange={(e) => setNameBoxText(e.target.value.toUpperCase())}
              className="h-7 w-20 rounded border border-slate-200 bg-slate-50 px-2 text-center font-mono text-xs font-semibold text-slate-800 focus:border-emerald-600 focus:bg-white focus:outline-none dark:border-slate-700 dark:bg-slate-950 dark:text-slate-100"
            />
          </form>

          <button
            type="button"
            title="Insert SUM formula"
            aria-label="Insert SUM formula"
            onClick={() => handleInsertFunction('SUM')}
            className="inline-flex h-7 w-7 shrink-0 items-center justify-center rounded border border-slate-200 bg-slate-50 text-emerald-700 hover:bg-emerald-50 dark:border-slate-700 dark:bg-slate-950 dark:text-emerald-400"
          >
            <Sigma className="h-3.5 w-3.5" />
          </button>

          <input
            type="text"
            aria-label="Formula bar input"
            value={draftValue}
            onFocus={() => {
              if (!editingCoord) setEditingCoord(activeCoord);
            }}
            onChange={(e) => {
              setDraftValue(e.target.value);
              if (!editingCoord) setEditingCoord(activeCoord);
            }}
            onKeyDown={(e) => {
              if (e.key === 'Enter') {
                e.preventDefault();
                finishEditing('down');
                (e.target as HTMLInputElement).blur();
              } else if (e.key === 'Escape') {
                e.preventDefault();
                setEditingCoord(null);
                setDraftValue(activeCell?.raw ?? '');
                (e.target as HTMLInputElement).blur();
              }
            }}
            onBlur={() => {
              if (editingCoord) finishEditing();
            }}
            placeholder="Enter value or formula (e.g. =SUM(B2:B4))"
            className="h-7 min-w-0 flex-1 rounded border border-slate-200 bg-slate-50 px-2.5 font-mono text-xs text-slate-900 focus:border-emerald-600 focus:bg-white focus:outline-none dark:border-slate-700 dark:bg-slate-950 dark:text-slate-100"
          />

          {editingCoord && (
            <div className="flex shrink-0 items-center gap-0.5">
              <button
                type="button"
                aria-label="Commit formula"
                onMouseDown={(e) => {
                  e.preventDefault();
                  finishEditing();
                }}
                className="inline-flex h-7 w-7 items-center justify-center rounded bg-emerald-600 text-white hover:bg-emerald-700"
              >
                <Check className="h-3.5 w-3.5" />
              </button>
              <button
                type="button"
                aria-label="Cancel formula"
                onMouseDown={(e) => {
                  e.preventDefault();
                  setEditingCoord(null);
                  setDraftValue(activeCell?.raw ?? '');
                }}
                className="inline-flex h-7 w-7 items-center justify-center rounded border border-slate-200 text-slate-500 hover:bg-slate-100 dark:border-slate-700"
              >
                <X className="h-3.5 w-3.5" />
              </button>
            </div>
          )}
        </div>
      )}

      {/* Main Grid Area + Optional Embedded Charts Panel */}
      <div className="relative flex min-h-0 flex-1 overflow-hidden">
        {/* Non-blocking Floating Find Bar in View Mode */}
        {!isEditing && viewFindOpen && (
          <div
            role="search"
            aria-label="Find in sheet"
            data-testid="sheets-view-find-bar"
            className="absolute top-2 right-4 z-50 flex flex-wrap items-center gap-1.5 rounded-xl border border-slate-200 bg-white/95 px-2.5 py-1.5 shadow-lg backdrop-blur-sm dark:border-slate-700 dark:bg-slate-900/95"
          >
            <Search className="h-3.5 w-3.5 shrink-0 text-slate-400" />
            <input
              ref={viewFindInputRef}
              type="text"
              aria-label="Find in sheet"
              value={viewFindQuery}
              onChange={(e) => setViewFindQuery(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') {
                  e.preventDefault();
                  stepViewFind(e.shiftKey ? -1 : 1);
                } else if (e.key === 'Escape') {
                  e.preventDefault();
                  setViewFindOpen(false);
                }
              }}
              placeholder="Find in sheet…"
              className="h-7 w-40 rounded-md border border-slate-200 bg-slate-50 px-2 text-xs text-slate-900 focus:border-emerald-600 focus:bg-white focus:outline-none sm:w-48 dark:border-slate-700 dark:bg-slate-950 dark:text-slate-100"
            />
            <span
              data-testid="sheets-find-counter"
              className="min-w-[60px] text-center font-mono text-[11px] tabular-nums text-slate-500 dark:text-slate-400"
            >
              {!viewFindQuery
                ? ''
                : viewFindMatches.length === 0
                  ? '0 / 0'
                  : `${(viewFindIndex % viewFindMatches.length) + 1} / ${viewFindMatches.length}`}
            </span>
            <button
              type="button"
              disabled={viewFindMatches.length === 0}
              onClick={() => stepViewFind(-1)}
              aria-label="Previous match"
              title="Previous match (Shift+Enter)"
              className="inline-flex h-7 w-7 items-center justify-center rounded-md border border-slate-200 text-slate-600 hover:bg-slate-100 disabled:opacity-35 dark:border-slate-700 dark:text-slate-300 dark:hover:bg-slate-800"
            >
              <ChevronUp className="h-3.5 w-3.5" />
            </button>
            <button
              type="button"
              disabled={viewFindMatches.length === 0}
              onClick={() => stepViewFind(1)}
              aria-label="Next match"
              title="Next match (Enter)"
              className="inline-flex h-7 w-7 items-center justify-center rounded-md border border-slate-200 text-slate-600 hover:bg-slate-100 disabled:opacity-35 dark:border-slate-700 dark:text-slate-300 dark:hover:bg-slate-800"
            >
              <ChevronDown className="h-3.5 w-3.5" />
            </button>
            <button
              type="button"
              onClick={() => setViewFindCaseSensitive((v) => !v)}
              aria-label="Match case"
              aria-pressed={viewFindCaseSensitive}
              title="Match case"
              className={`inline-flex h-7 items-center justify-center rounded-md border px-1.5 font-mono text-[11px] font-semibold transition-colors ${
                viewFindCaseSensitive
                  ? 'border-emerald-600 bg-emerald-50 text-emerald-800 dark:border-emerald-500 dark:bg-emerald-950/60 dark:text-emerald-300'
                  : 'border-slate-200 text-slate-500 hover:bg-slate-100 dark:border-slate-700 dark:text-slate-400 dark:hover:bg-slate-800'
              }`}
            >
              Aa
            </button>
            <button
              type="button"
              onClick={() => setViewFindOpen(false)}
              aria-label="Close find bar"
              title="Close (Esc)"
              className="inline-flex h-7 w-7 items-center justify-center rounded-md text-slate-400 hover:bg-slate-100 hover:text-slate-700 dark:hover:bg-slate-800 dark:hover:text-slate-200"
            >
              <X className="h-3.5 w-3.5" />
            </button>
          </div>
        )}

        <div
          ref={gridContainerRef}
          data-testid="sheets-scroll-container"
          className="relative min-w-0 flex-1 overflow-auto bg-white dark:bg-slate-950"
          onClick={() => {
            if (openFilterCol !== null) setOpenFilterCol(null);
          }}
        >
          {loadingSheet ? (
            <div className="space-y-2 p-6">
              <div className="h-8 w-full animate-pulse rounded bg-slate-200 dark:bg-slate-800" />
              <div className="h-8 w-full animate-pulse rounded bg-slate-100 dark:bg-slate-900" />
              <div className="h-8 w-5/6 animate-pulse rounded bg-slate-100 dark:bg-slate-900" />
            </div>
          ) : (
            <table
              role="grid"
              aria-label={`${activeSheet.name} spreadsheet grid`}
              data-testid="sheets-grid"
              className="table-fixed border-separate border-spacing-0 select-none"
              style={{
                width: `${totalTableWidth}px`,
                fontSize: `${Math.max(9, Math.round(12 * zoom))}px`,
              }}
            >
              <thead>
                <tr style={{height: `${scaledColHeaderH}px`}}>
                  {/* Top-left corner select-all cell */}
                  <th
                    onClick={() =>
                      setSelection({
                        start: {row: 0, col: 0},
                        end: {row: activeSheet.rowCount - 1, col: activeSheet.colCount - 1},
                      })
                    }
                    title="Select all cells"
                    className="sticky top-0 left-0 z-40 border-r border-b border-slate-300 bg-slate-100 text-center text-[10px] font-semibold text-slate-500 dark:border-slate-700 dark:bg-slate-900"
                    style={{
                      width: `${scaledRowHeaderW}px`,
                      minWidth: `${scaledRowHeaderW}px`,
                      maxWidth: `${scaledRowHeaderW}px`,
                    }}
                  />

                  {Array.from({length: activeSheet.colCount}, (_, c) => {
                    const colW = getScaledColWidth(c);
                    const isColSelected = c >= normSel.minCol && c <= normSel.maxCol;
                    const isFrozenCol = c < activeSheet.freezeCols;
                    const isFiltered = (activeSheet.columnFilters[c]?.length ?? 0) > 0;
                    return (
                      <th
                        key={c}
                        data-col-header={c}
                        onClick={() =>
                          setSelection({
                            start: {row: 0, col: c},
                            end: {row: activeSheet.rowCount - 1, col: c},
                          })
                        }
                        className={`sticky top-0 border-r border-b border-slate-300 px-1 text-center font-semibold dark:border-slate-700 ${
                          isFrozenCol ? 'z-30' : 'z-20'
                        } ${
                          isColSelected
                            ? 'bg-emerald-100 text-emerald-900 dark:bg-emerald-950 dark:text-emerald-300'
                            : 'bg-slate-100 text-slate-600 dark:bg-slate-900 dark:text-slate-300'
                        } ${
                          c === activeSheet.freezeCols - 1
                            ? 'shadow-[2px_0_0_0_rgba(16,185,129,0.55)]'
                            : ''
                        }`}
                        style={{
                          width: `${colW}px`,
                          minWidth: `${colW}px`,
                          maxWidth: `${colW}px`,
                          left: isFrozenCol ? `${colLeftOffsets[c]}px` : undefined,
                        }}
                      >
                        <div className="relative flex h-full items-center justify-center gap-1">
                          <span>{colIndexToLetter(c)}</span>
                          {isEditing && activeSheet.autoFilterEnabled && (
                            <button
                              type="button"
                              aria-label={`Filter column ${colIndexToLetter(c)}`}
                              onClick={(e) => {
                                e.stopPropagation();
                                setOpenFilterCol(openFilterCol === c ? null : c);
                              }}
                              className={`rounded p-0.5 ${
                                isFiltered
                                  ? 'bg-emerald-600 text-white'
                                  : 'text-slate-400 hover:bg-slate-200 hover:text-slate-700 dark:hover:bg-slate-800'
                              }`}
                            >
                              <Filter className="h-3 w-3" />
                            </button>
                          )}

                          {/* Column resize drag handle (viewing-only in View mode; double-click to auto-fit) */}
                          <div
                            data-col-resize={c}
                            title={
                              isEditing
                                ? 'Drag to resize column (double-click to auto-fit)'
                                : 'Drag to resize column for reading (double-click to auto-fit)'
                            }
                            onMouseDown={(e) => {
                              e.stopPropagation();
                              startColResize(c, e.clientX);
                            }}
                            onDoubleClick={(e) => {
                              e.stopPropagation();
                              handleAutoFitCol(c);
                            }}
                            className="absolute top-0 -right-1 bottom-0 w-2 cursor-col-resize hover:bg-emerald-500/40"
                          />
                        </div>

                        {/* Auto-Filter dropdown popover */}
                        {openFilterCol === c && (
                          <div
                            onClick={(e) => e.stopPropagation()}
                            className="absolute top-full left-0 z-50 mt-0.5 w-48 rounded-lg border border-slate-200 bg-white p-2 text-left font-normal shadow-xl dark:border-slate-700 dark:bg-slate-900"
                          >
                            <div className="mb-1.5 flex items-center justify-between text-[11px] font-semibold text-slate-700 dark:text-slate-200">
                              <span>Filter {colIndexToLetter(c)}</span>
                              <button
                                type="button"
                                onClick={() => {
                                  updateActiveSheet((s) => {
                                    const next = {...s.columnFilters};
                                    delete next[c];
                                    return {...s, columnFilters: next};
                                  }, false);
                                  setOpenFilterCol(null);
                                }}
                                className="text-emerald-700 hover:underline dark:text-emerald-400"
                              >
                                Clear
                              </button>
                            </div>
                            <div className="max-h-40 space-y-1 overflow-y-auto">
                              {filterColumnOptions.length === 0 ? (
                                <p className="text-[11px] text-slate-400">No values</p>
                              ) : (
                                filterColumnOptions.map((val) => {
                                  const activeList = activeSheet.columnFilters[c] ?? [];
                                  const checked =
                                    activeList.length === 0 || activeList.includes(val);
                                  return (
                                    <label
                                      key={val}
                                      className="flex cursor-pointer items-center gap-1.5 rounded px-1 py-0.5 text-[11px] text-slate-700 hover:bg-slate-100 dark:text-slate-200 dark:hover:bg-slate-800"
                                    >
                                      <input
                                        type="checkbox"
                                        checked={checked}
                                        onChange={(e) => {
                                          const base =
                                            activeList.length === 0
                                              ? filterColumnOptions
                                              : activeList;
                                          const next = e.target.checked
                                            ? Array.from(new Set([...base, val]))
                                            : base.filter((x) => x !== val);
                                          updateActiveSheet(
                                            (s) => ({
                                              ...s,
                                              columnFilters: {
                                                ...s.columnFilters,
                                                [c]:
                                                  next.length === filterColumnOptions.length
                                                    ? []
                                                    : next,
                                              },
                                            }),
                                            false,
                                          );
                                        }}
                                        className="h-3 w-3 accent-emerald-600"
                                      />
                                      <span className="truncate">{val}</span>
                                    </label>
                                  );
                                })
                              )}
                            </div>
                          </div>
                        )}
                      </th>
                    );
                  })}
                </tr>
              </thead>

              <tbody>
                {visibleRows.map((r) => {
                  const rowH = getScaledRowHeight(r);
                  const isRowSelected = r >= normSel.minRow && r <= normSel.maxRow;
                  const isFrozenRow = r < activeSheet.freezeRows;

                  return (
                    <tr key={r} style={{height: `${rowH}px`}}>
                      {/* Row Number Gutter */}
                      <th
                        data-row-header={r}
                        onClick={() =>
                          setSelection({
                            start: {row: r, col: 0},
                            end: {row: r, col: activeSheet.colCount - 1},
                          })
                        }
                        className={`sticky left-0 border-r border-b border-slate-300 text-center font-mono font-medium tabular-nums dark:border-slate-700 ${
                          isFrozenRow ? 'z-30' : 'z-20'
                        } ${
                          isRowSelected
                            ? 'bg-emerald-100 text-emerald-900 dark:bg-emerald-950 dark:text-emerald-300'
                            : 'bg-slate-100 text-slate-500 dark:bg-slate-900 dark:text-slate-400'
                        } ${
                          r === activeSheet.freezeRows - 1
                            ? 'shadow-[0_2px_0_0_rgba(16,185,129,0.55)]'
                            : ''
                        }`}
                        style={{
                          width: `${scaledRowHeaderW}px`,
                          minWidth: `${scaledRowHeaderW}px`,
                          maxWidth: `${scaledRowHeaderW}px`,
                          fontSize: `${Math.max(9, Math.round(11 * zoom))}px`,
                          top: isFrozenRow ? `${rowTopOffsets[r]}px` : undefined,
                        }}
                      >
                        <div className="relative flex h-full items-center justify-center">
                          <span>{r + 1}</span>
                          {isEditing && (
                            <div
                              onMouseDown={(e) => {
                                e.stopPropagation();
                                startRowResize(r, e.clientY);
                              }}
                              className="absolute right-0 -bottom-1 left-0 h-2 cursor-row-resize hover:bg-emerald-500/40"
                            />
                          )}
                        </div>
                      </th>

                      {/* Row Data Cells */}
                      {Array.from({length: activeSheet.colCount}, (_, c) => {
                        const merge = findMergeAt(r, c, activeSheet.merges);
                        if (merge && (merge.startRow !== r || merge.startCol !== c)) {
                          return null;
                        }

                        const k = cellKey(r, c);
                        const cell = activeSheet.cells[k];
                        const evaluated = activeEvalMap[k];
                        const cond = evaluateConditionalFormatting(
                          r,
                          c,
                          evaluated,
                          activeSheet.conditionalRules,
                        );
                        const isActive = activeCoord.row === r && activeCoord.col === c;
                        const isInSel =
                          r >= normSel.minRow &&
                          r <= normSel.maxRow &&
                          c >= normSel.minCol &&
                          c <= normSel.maxCol;
                        const isCellEditing =
                          isEditing && editingCoord?.row === r && editingCoord?.col === c;
                        const isSelectionCorner =
                          r === normSel.maxRow && c === normSel.maxCol;
                        const isInFillPreview =
                          isEditing &&
                          Boolean(fillTarget) &&
                          r >= Math.min(normSel.minRow, fillTarget!.row) &&
                          r <= Math.max(normSel.maxRow, fillTarget!.row) &&
                          c >= Math.min(normSel.minCol, fillTarget!.col) &&
                          c <= Math.max(normSel.maxCol, fillTarget!.col) &&
                          !isInSel;

                        const isViewFindMatch = !isEditing && viewFindMatchKeySet.has(k);
                        const isActiveViewMatch =
                          !isEditing &&
                          activeViewMatchCoord?.row === r &&
                          activeViewMatchCoord?.col === c;

                        const isFrozenCol = c < activeSheet.freezeCols;
                        const isFrozenCell = isFrozenRow || isFrozenCol;
                        const st = cell?.style;
                        const isNumericVal = typeof evaluated?.value === 'number';
                        const textAlign = st?.align ?? (isNumericVal ? 'right' : 'left');
                        const bgColor = cond?.bg ?? st?.bg;
                        const fgColor =
                          evaluated?.error
                            ? '#DC2626'
                            : cond?.color ?? st?.color;

                        const colW = getScaledColWidth(c);
                        const scaledFontPx = Math.max(9, Math.round((st?.fontSize ?? 13) * zoom));

                        // Strict z-index layering so normal cells never overlap frozen rows/cols
                        const zClass =
                          isFrozenRow && isFrozenCol
                            ? isActive
                              ? 'z-[21]'
                              : 'z-20'
                            : isFrozenCell
                              ? isActive
                                ? 'z-[11]'
                                : 'z-10'
                              : isActive
                                ? 'z-[1]'
                                : 'z-0';

                        // Combine freeze divider shadows cleanly at frozen row/col intersection
                        const isFreezeBottom = r === activeSheet.freezeRows - 1;
                        const isFreezeRight = c === activeSheet.freezeCols - 1;
                        const freezeShadow =
                          isFreezeBottom && isFreezeRight
                            ? '0 2px 0 0 rgba(16,185,129,0.5), 2px 0 0 0 rgba(16,185,129,0.5)'
                            : isFreezeBottom
                              ? '0 2px 0 0 rgba(16,185,129,0.5)'
                              : isFreezeRight
                                ? '2px 0 0 0 rgba(16,185,129,0.5)'
                                : undefined;

                        // Keep opaque backgrounds on frozen cells so scrolled cells behind never bleed through
                        const stateBgClass = isActiveViewMatch
                          ? 'bg-amber-200 ring-2 ring-amber-500 ring-inset dark:bg-amber-800/80'
                          : isViewFindMatch
                            ? 'bg-amber-100 dark:bg-amber-950'
                            : isActive
                              ? 'bg-white ring-2 ring-emerald-600 ring-inset dark:bg-slate-950'
                              : isInSel
                                ? isFrozenCell
                                  ? 'bg-emerald-50 dark:bg-emerald-950'
                                  : 'bg-emerald-50/70 dark:bg-emerald-950/40'
                                : isInFillPreview
                                  ? 'bg-emerald-100/40 ring-1 ring-emerald-500 ring-dashed ring-inset'
                                  : 'bg-white dark:bg-slate-950';

                        return (
                          <td
                            key={c}
                            role="gridcell"
                            aria-label={`Cell ${formatCellRef(r, c)}`}
                            data-cell={formatCellRef(r, c)}
                            data-frozen={isFrozenCell ? 'true' : undefined}
                            data-match={isViewFindMatch ? 'true' : undefined}
                            rowSpan={merge ? merge.endRow - merge.startRow + 1 : undefined}
                            colSpan={merge ? merge.endCol - merge.startCol + 1 : undefined}
                            onMouseDown={(e) => {
                              if (e.button !== 0) return;
                              if (editingCoord && (editingCoord.row !== r || editingCoord.col !== c)) {
                                finishEditing();
                              }
                              if (e.shiftKey) {
                                setSelection((prev) => ({start: prev.start, end: {row: r, col: c}}));
                              } else {
                                isSelectingRef.current = true;
                                setSelection({start: {row: r, col: c}, end: {row: r, col: c}});
                              }
                            }}
                            onMouseEnter={() => {
                              if (isSelectingRef.current) {
                                setSelection((prev) => ({start: prev.start, end: {row: r, col: c}}));
                              } else if (isEditing && isFillingRef.current) {
                                setFillTarget({row: r, col: c});
                              }
                            }}
                            onDoubleClick={() => {
                              if (isEditing) startEditingAt({row: r, col: c});
                            }}
                            className={`relative px-1.5 py-0.5 align-middle ${
                              showGridlines
                                ? 'border-r border-b border-slate-200 dark:border-slate-800'
                                : ''
                            } ${
                              st?.border ? 'ring-1 ring-slate-600 ring-inset dark:ring-slate-400' : ''
                            } ${isFrozenCell ? 'sticky' : ''} ${zClass} ${stateBgClass}`}
                            style={{
                              width: `${colW}px`,
                              maxWidth: merge ? undefined : `${colW}px`,
                              top: isFrozenRow ? `${rowTopOffsets[r]}px` : undefined,
                              left: isFrozenCol ? `${colLeftOffsets[c]}px` : undefined,
                              boxShadow: freezeShadow,
                              backgroundColor:
                                isActiveViewMatch || isViewFindMatch
                                  ? undefined
                                  : bgColor || undefined,
                              color: fgColor || undefined,
                              fontSize: `${scaledFontPx}px`,
                              fontWeight: st?.bold ? 700 : undefined,
                              fontStyle: st?.italic ? 'italic' : undefined,
                              textDecoration:
                                [st?.underline ? 'underline' : '', st?.strike ? 'line-through' : '']
                                  .filter(Boolean)
                                  .join(' ') || undefined,
                              textAlign,
                            }}
                          >
                            {isCellEditing ? (
                              activeValidation?.kind === 'list' &&
                              activeValidation.listItems?.length ? (
                                <select
                                  aria-label={`Select value for ${formatCellRef(r, c)}`}
                                  autoFocus
                                  value={draftValue}
                                  onChange={(e) => {
                                    setDraftValue(e.target.value);
                                    commitCellValue(r, c, e.target.value);
                                    setEditingCoord(null);
                                  }}
                                  onBlur={() => setEditingCoord(null)}
                                  className="h-full w-full bg-white text-xs text-slate-900 focus:outline-none dark:bg-slate-900 dark:text-slate-100"
                                >
                                  <option value="">-- Choose --</option>
                                  {activeValidation.listItems.map((opt) => (
                                    <option key={opt} value={opt}>
                                      {opt}
                                    </option>
                                  ))}
                                </select>
                              ) : (
                                <input
                                  ref={inlineInputRef}
                                  type="text"
                                  aria-label={`Edit cell ${formatCellRef(r, c)}`}
                                  value={draftValue}
                                  onChange={(e) => setDraftValue(e.target.value)}
                                  onKeyDown={(e) => {
                                    e.stopPropagation();
                                    if (e.key === 'Enter') {
                                      e.preventDefault();
                                      finishEditing(e.shiftKey ? 'up' : 'down');
                                    } else if (e.key === 'Tab') {
                                      e.preventDefault();
                                      finishEditing(e.shiftKey ? 'left' : 'right');
                                    } else if (e.key === 'Escape') {
                                      e.preventDefault();
                                      setEditingCoord(null);
                                      setDraftValue(cell?.raw ?? '');
                                    }
                                  }}
                                  onBlur={() => finishEditing()}
                                  className="h-full w-full bg-white font-mono text-xs text-slate-900 focus:outline-none dark:bg-slate-900 dark:text-slate-100"
                                />
                              )
                            ) : (
                              <div
                                className={`tabular-nums ${
                                  st?.wrap
                                    ? 'whitespace-pre-wrap break-words'
                                    : 'overflow-hidden text-ellipsis whitespace-nowrap'
                                }`}
                              >
                                {evaluated?.display ?? ''}
                              </div>
                            )}

                            {/* Fill Handle on bottom-right corner of selection (Edit Mode only) */}
                            {isEditing && isSelectionCorner && !isCellEditing && (
                              <div
                                title="Drag to fill values or formulas"
                                onMouseDown={(e) => {
                                  e.stopPropagation();
                                  isFillingRef.current = true;
                                  setFillTarget({row: r, col: c});
                                }}
                                className="absolute -right-[4px] -bottom-[4px] z-20 h-2.5 w-2.5 cursor-crosshair border border-white bg-emerald-600"
                              />
                            )}
                          </td>
                        );
                      })}
                    </tr>
                  );
                })}
              </tbody>
            </table>
          )}
        </div>

        {/* Embedded Charts Panel */}
        {chartsPanelOpen && activeSheet.charts.length > 0 && (
          <div className="absolute inset-y-0 right-0 z-30 flex w-full sm:static sm:w-auto">
            <SheetsChartsPanel
              sheet={activeSheet}
              evalMap={activeEvalMap}
              readOnly={!isEditing}
              onChangeChartKind={(chartId, kind) =>
                updateActiveSheet((s) => ({
                  ...s,
                  charts: s.charts.map((ch) => (ch.id === chartId ? {...ch, kind} : ch)),
                }))
              }
              onDeleteChart={(chartId) =>
                updateActiveSheet((s) => ({
                  ...s,
                  charts: s.charts.filter((ch) => ch.id !== chartId),
                }))
              }
              onClose={() => setChartsPanelOpen(false)}
            />
          </div>
        )}
      </div>

      {/* Bottom Sheet Tabs + Live Calculation Status Bar */}
      <footer className="flex flex-wrap items-center justify-between gap-2 border-t border-slate-200 bg-white px-2.5 py-1 text-xs text-slate-600 dark:border-slate-800 dark:bg-slate-900 dark:text-slate-400">
        {/* Multi-Sheet Tab Bar */}
        <div className="flex min-w-0 flex-1 items-center gap-1 overflow-x-auto">
          {isEditing && (
            <button
              type="button"
              onClick={handleAddSheet}
              aria-label="Add new sheet"
              title="Add new sheet"
              className="inline-flex h-7 w-7 shrink-0 items-center justify-center rounded-md border border-slate-200 bg-slate-50 text-slate-700 hover:bg-emerald-50 hover:text-emerald-700 dark:border-slate-700 dark:bg-slate-950 dark:text-slate-200"
            >
              <Plus className="h-3.5 w-3.5" />
            </button>
          )}

          <div role="tablist" aria-label="Workbook sheets" className="flex items-center gap-1">
            {workbook.sheets.map((s) => {
              const active = s.id === activeSheet.id;
              const isRenaming = isEditing && renamingSheetId === s.id;
              return (
                <div
                  key={s.id}
                  className={`flex shrink-0 items-center gap-1 rounded-md border px-2 py-1 ${
                    active
                      ? 'border-emerald-500 bg-emerald-50 font-semibold text-emerald-900 dark:border-emerald-700 dark:bg-emerald-950/60 dark:text-emerald-200'
                      : 'border-slate-200 bg-slate-50 text-slate-600 hover:bg-slate-100 dark:border-slate-800 dark:bg-slate-950 dark:text-slate-400'
                  }`}
                >
                  {isRenaming ? (
                    <form
                      onSubmit={(e) => {
                        e.preventDefault();
                        commitSheetRename();
                      }}
                    >
                      <input
                        type="text"
                        autoFocus
                        aria-label="Rename sheet"
                        value={renameDraft}
                        onChange={(e) => setRenameDraft(e.target.value)}
                        onBlur={commitSheetRename}
                        className="w-24 rounded border border-emerald-500 bg-white px-1 text-xs text-slate-900 focus:outline-none dark:bg-slate-900 dark:text-slate-100"
                      />
                    </form>
                  ) : (
                    <button
                      type="button"
                      role="tab"
                      aria-selected={active}
                      onClick={() => handleSelectSheet(s.id)}
                      onDoubleClick={() => {
                        if (!isEditing) return;
                        setRenamingSheetId(s.id);
                        setRenameDraft(s.name);
                      }}
                      className="whitespace-nowrap"
                    >
                      {s.name}
                    </button>
                  )}

                  {isEditing && active && !isRenaming && (
                    <div className="flex items-center gap-0.5 pl-1">
                      <button
                        type="button"
                        aria-label={`Rename ${s.name}`}
                        title="Rename sheet"
                        onClick={() => {
                          setRenamingSheetId(s.id);
                          setRenameDraft(s.name);
                        }}
                        className="rounded p-0.5 text-slate-400 hover:text-slate-700 dark:hover:text-slate-200"
                      >
                        <Pencil className="h-2.5 w-2.5" />
                      </button>
                      <button
                        type="button"
                        aria-label={`Duplicate ${s.name}`}
                        title="Duplicate sheet"
                        onClick={() => handleDuplicateSheet(s.id)}
                        className="rounded p-0.5 text-slate-400 hover:text-slate-700 dark:hover:text-slate-200"
                      >
                        <Copy className="h-2.5 w-2.5" />
                      </button>
                      {workbook.sheets.length > 1 && (
                        <button
                          type="button"
                          aria-label={`Delete ${s.name}`}
                          title="Delete sheet"
                          onClick={() => handleDeleteSheet(s.id)}
                          className="rounded p-0.5 text-slate-400 hover:text-rose-600"
                        >
                          <Trash2 className="h-2.5 w-2.5" />
                        </button>
                      )}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        </div>

        {/* Live Selection Calculations */}
        <div className="flex shrink-0 items-center gap-3 tabular-nums text-[11px]">
          {selectionStats.countAll > 0 && (
            <>
              <span>Count: {selectionStats.countAll}</span>
              {selectionStats.numCount > 0 && (
                <>
                  <span className="hidden sm:inline">
                    Avg: {selectionStats.avg.toLocaleString('en-US', {maximumFractionDigits: 2})}
                  </span>
                  <span className="font-semibold text-slate-800 dark:text-slate-200">
                    Sum: {selectionStats.sum.toLocaleString('en-US', {maximumFractionDigits: 2})}
                  </span>
                </>
              )}
            </>
          )}
          {(activeSheet.freezeRows > 0 || activeSheet.freezeCols > 0) && (
            <span className="hidden font-medium text-emerald-700 md:inline dark:text-emerald-400">
              Frozen ({activeSheet.freezeRows}R × {activeSheet.freezeCols}C)
            </span>
          )}
          <span className="font-semibold text-slate-700 dark:text-slate-200">
            {Math.round(zoom * 100)}%
          </span>
        </div>
      </footer>

      {/* Mobile View-Mode Bottom Action Bar */}
      {!isEditing && (
        <ViewModeMobileBar
          moduleType="sheets"
          onEdit={onEnterEdit}
          onSearch={() => {
            setViewFindOpen((open) => {
              const next = !open;
              if (next) requestAnimationFrame(() => viewFindInputRef.current?.focus());
              return next;
            });
          }}
          onExport={() => setOpenDialog('export')}
        />
      )}

      {/* Spreadsheets Modal Suite */}
      <SheetsDialogs
        openDialog={openDialog}
        defaultChartKind={defaultChartKind}
        selectionRangeLabel={selectionRangeLabel}
        fileName={tab.name}
        conditionalRules={activeSheet.conditionalRules}
        validations={activeSheet.validations}
        readOnly={!isEditing}
        onClose={() => setOpenDialog(null)}
        onFindNext={handleFindNext}
        onFindPrev={handleFindPrev}
        onReplaceAll={handleReplaceAll}
        onAddConditionalRule={(rule: Omit<ConditionalRule, 'id'>) => {
          updateActiveSheet((s) => ({
            ...s,
            conditionalRules: [
              ...s.conditionalRules,
              {...rule, id: `cf_${Date.now()}_${Math.random().toString(36).slice(2, 5)}`},
            ],
          }));
          pushToast('Conditional formatting rule added', 'success');
        }}
        onDeleteConditionalRule={(id) =>
          updateActiveSheet((s) => ({
            ...s,
            conditionalRules: s.conditionalRules.filter((r) => r.id !== id),
          }))
        }
        onSaveValidation={(rule: Omit<DataValidationRule, 'id'>) => {
          updateActiveSheet((s) => ({
            ...s,
            validations: [
              ...s.validations,
              {...rule, id: `dv_${Date.now()}_${Math.random().toString(36).slice(2, 5)}`},
            ],
          }));
          pushToast('Data validation rule applied', 'success');
        }}
        onDeleteValidation={(id) =>
          updateActiveSheet((s) => ({
            ...s,
            validations: s.validations.filter((v) => v.id !== id),
          }))
        }
        onCreateChart={(spec) => {
          updateActiveSheet((s) => ({
            ...s,
            charts: [
              ...s.charts,
              {...spec, id: `ch_${Date.now()}_${Math.random().toString(36).slice(2, 5)}`},
            ],
          }));
          setChartsPanelOpen(true);
          pushToast(`Created ${spec.kind} chart "${spec.title}"`, 'success');
        }}
        onExportFormat={handleExportFormat}
      />
    </div>
  );
};

export default SheetsModule;
