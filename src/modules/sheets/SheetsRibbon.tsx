import React from 'react';
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  ArrowDownZA,
  ArrowUpAZ,
  BarChart3,
  Bold,
  Check,
  Columns3,
  Download,
  Eraser,
  FileDown,
  Filter,
  Grid3X3,
  Italic,
  Merge,
  PaintBucket,
  Palette,
  Pin,
  Plus,
  Printer,
  Redo2,
  Rows3,
  Save,
  Search,
  ShieldCheck,
  Sigma,
  Star,
  Strikethrough,
  Trash2,
  Type,
  Underline,
  Undo2,
  WrapText,
  ZoomIn,
  ZoomOut,
} from 'lucide-react';
import type {
  CellNumberFormat,
  CellStyle,
  ChartKind,
  SheetsRibbonTab,
} from './sheetsTypes';
import type {SheetsDialogKind} from './SheetsDialogs';

interface SheetsRibbonProps {
  activeTab: SheetsRibbonTab;
  onChangeTab: (tab: SheetsRibbonTab) => void;
  activeStyle: CellStyle;
  canUndo: boolean;
  canRedo: boolean;
  isFavorite: boolean;
  isMerged: boolean;
  autoFilterEnabled: boolean;
  freezeRows: number;
  freezeCols: number;
  showGridlines: boolean;
  showFormulaBar: boolean;
  zoom: number;
  chartCount: number;
  chartsPanelOpen: boolean;
  onUndo: () => void;
  onRedo: () => void;
  onSaveNow: () => void;
  onToggleFavorite: () => void;
  onExportPdf: () => void;
  onPrint: () => void;
  onOpenDialog: (kind: SheetsDialogKind, defaultChartKind?: ChartKind) => void;
  onApplyStyle: (patch: Partial<CellStyle>) => void;
  onStepDecimals: (delta: number) => void;
  onToggleMerge: () => void;
  onClearSelection: () => void;
  onInsertRow: (where: 'above' | 'below') => void;
  onDeleteRow: () => void;
  onInsertCol: (where: 'left' | 'right') => void;
  onDeleteCol: () => void;
  onInsertFunction: (fnName: string) => void;
  onSort: (dir: 'asc' | 'desc') => void;
  onToggleAutoFilter: () => void;
  onSetFreeze: (mode: 'top-row' | 'first-col' | 'selection' | 'none') => void;
  onToggleGridlines: () => void;
  onToggleFormulaBar: () => void;
  onChangeZoom: (z: number) => void;
  onToggleChartsPanel: () => void;
  onDoneEditing: () => void;
}

const RIBBON_TABS: Array<{id: SheetsRibbonTab; label: string}> = [
  {id: 'home', label: 'Home'},
  {id: 'insert', label: 'Insert'},
  {id: 'data', label: 'Data'},
  {id: 'view', label: 'View'},
];

const QUICK_FUNCTIONS = [
  'SUM',
  'AVERAGE',
  'MIN',
  'MAX',
  'COUNT',
  'COUNTIF',
  'IF',
  'VLOOKUP',
  'XLOOKUP',
  'PMT',
];

const RibbonGroup: React.FC<{label: string; children: React.ReactNode}> = ({label, children}) => (
  <div className="flex shrink-0 flex-col justify-between border-r border-slate-200 pr-2 last:border-r-0 dark:border-slate-800">
    <div className="flex items-center gap-0.5">{children}</div>
    <span className="mt-0.5 text-center text-[10px] font-medium text-slate-400 dark:text-slate-500">
      {label}
    </span>
  </div>
);

const ToolBtn: React.FC<{
  active?: boolean;
  disabled?: boolean;
  onClick: () => void;
  title: string;
  label?: string;
  children: React.ReactNode;
}> = ({active, disabled, onClick, title, label, children}) => (
  <button
    type="button"
    disabled={disabled}
    onClick={onClick}
    title={title}
    aria-label={title}
    className={`inline-flex min-h-[30px] shrink-0 items-center justify-center gap-1 whitespace-nowrap rounded-md px-1.5 text-xs font-medium transition-colors disabled:opacity-40 ${
      active
        ? 'bg-emerald-100 text-emerald-900 dark:bg-emerald-950/70 dark:text-emerald-300'
        : 'text-slate-700 hover:bg-slate-100 dark:text-slate-200 dark:hover:bg-slate-800'
    }`}
  >
    {children}
    {label && <span>{label}</span>}
  </button>
);

export const SheetsRibbon: React.FC<SheetsRibbonProps> = ({
  activeTab,
  onChangeTab,
  activeStyle,
  canUndo,
  canRedo,
  isFavorite,
  isMerged,
  autoFilterEnabled,
  freezeRows,
  freezeCols,
  showGridlines,
  showFormulaBar,
  zoom,
  chartCount,
  chartsPanelOpen,
  onUndo,
  onRedo,
  onSaveNow,
  onToggleFavorite,
  onExportPdf,
  onPrint,
  onOpenDialog,
  onApplyStyle,
  onStepDecimals,
  onToggleMerge,
  onClearSelection,
  onInsertRow,
  onDeleteRow,
  onInsertCol,
  onDeleteCol,
  onInsertFunction,
  onSort,
  onToggleAutoFilter,
  onSetFreeze,
  onToggleGridlines,
  onToggleFormulaBar,
  onChangeZoom,
  onToggleChartsPanel,
  onDoneEditing,
}) => {
  return (
    <div className="border-b border-slate-200 bg-white dark:border-slate-800 dark:bg-slate-900">
      {/* Quick Access + 4 Ribbon Tabs + File Actions */}
      <div className="flex items-center justify-between gap-2 overflow-x-auto border-b border-slate-200 bg-slate-50 px-2.5 pt-1.5 dark:border-slate-800 dark:bg-slate-950">
        <div className="flex shrink-0 items-center gap-1">
          <div className="mr-2 flex shrink-0 items-center gap-0.5 border-r border-slate-200 pr-2 pb-1 dark:border-slate-800">
            <ToolBtn onClick={onSaveNow} title="Save workbook (Ctrl+S)">
              <Save className="h-3.5 w-3.5 text-emerald-700 dark:text-emerald-400" />
            </ToolBtn>
            <ToolBtn disabled={!canUndo} onClick={onUndo} title="Undo (Ctrl+Z)">
              <Undo2 className="h-3.5 w-3.5" />
            </ToolBtn>
            <ToolBtn disabled={!canRedo} onClick={onRedo} title="Redo (Ctrl+Y)">
              <Redo2 className="h-3.5 w-3.5" />
            </ToolBtn>
            <ToolBtn
              onClick={onToggleFavorite}
              title={isFavorite ? 'Remove favourite' : 'Add favourite'}
            >
              <Star className={`h-3.5 w-3.5 ${isFavorite ? 'fill-amber-400 text-amber-500' : ''}`} />
            </ToolBtn>
          </div>

          <div role="tablist" aria-label="Spreadsheet ribbon tabs" className="flex shrink-0 items-end gap-0.5">
            {RIBBON_TABS.map((t) => {
              const active = activeTab === t.id;
              return (
                <button
                  key={t.id}
                  type="button"
                  role="tab"
                  aria-selected={active}
                  onClick={() => onChangeTab(t.id)}
                  className={`whitespace-nowrap rounded-t-lg border-x border-t px-3 py-1.5 text-xs font-semibold transition-colors ${
                    active
                      ? '-mb-px border-slate-200 bg-white text-emerald-700 dark:border-slate-800 dark:bg-slate-900 dark:text-emerald-400'
                      : 'border-transparent text-slate-600 hover:bg-slate-200/60 hover:text-slate-900 dark:text-slate-400 dark:hover:bg-slate-800 dark:hover:text-slate-100'
                  }`}
                >
                  {t.label}
                </button>
              );
            })}
          </div>
        </div>

        <div className="flex shrink-0 items-center gap-1 pb-1">
          {chartCount > 0 && (
            <ToolBtn
              active={chartsPanelOpen}
              onClick={onToggleChartsPanel}
              title="Toggle Charts Panel"
              label={`Charts (${chartCount})`}
            >
              <BarChart3 className="h-3.5 w-3.5 text-emerald-700 dark:text-emerald-400" />
            </ToolBtn>
          )}
          <ToolBtn onClick={onExportPdf} title="Export workbook as PDF" label="PDF">
            <FileDown className="h-3.5 w-3.5 text-rose-600 dark:text-rose-400" />
          </ToolBtn>
          <ToolBtn onClick={onPrint} title="Print sheet">
            <Printer className="h-3.5 w-3.5" />
          </ToolBtn>
          <button
            type="button"
            onClick={() => onOpenDialog('export')}
            className="inline-flex min-h-[30px] shrink-0 items-center gap-1 whitespace-nowrap rounded-lg border border-slate-300 bg-white px-2.5 text-xs font-semibold text-slate-800 hover:bg-slate-100 dark:border-slate-700 dark:bg-slate-800 dark:text-slate-100 dark:hover:bg-slate-700"
          >
            <Download className="h-3.5 w-3.5" />
            Save As
          </button>
          <button
            type="button"
            onClick={onDoneEditing}
            aria-label="Done editing"
            title="Save and return to View mode (Esc)"
            className="inline-flex min-h-[30px] shrink-0 items-center gap-1 whitespace-nowrap rounded-lg bg-emerald-700 px-3 text-xs font-semibold text-white hover:bg-emerald-800 dark:bg-emerald-600 dark:hover:bg-emerald-500"
          >
            <Check className="h-3.5 w-3.5" />
            Done
          </button>
        </div>
      </div>

      {/* Ribbon Command Surface */}
      <div className="flex items-stretch gap-2 overflow-x-auto px-2.5 py-1.5">
        {activeTab === 'home' && (
          <>
            <RibbonGroup label="Font">
              <select
                aria-label="Cell font size"
                value={activeStyle.fontSize ?? 13}
                onChange={(e) => onApplyStyle({fontSize: Number(e.target.value)})}
                className="h-7 w-14 rounded border border-slate-200 bg-slate-50 px-1 text-xs tabular-nums text-slate-800 dark:border-slate-700 dark:bg-slate-950 dark:text-slate-100"
              >
                {[10, 11, 12, 13, 14, 16, 18, 20].map((sz) => (
                  <option key={sz} value={sz}>
                    {sz}px
                  </option>
                ))}
              </select>
              <ToolBtn
                active={Boolean(activeStyle.bold)}
                onClick={() => onApplyStyle({bold: !activeStyle.bold})}
                title="Bold (Ctrl+B)"
              >
                <Bold className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                active={Boolean(activeStyle.italic)}
                onClick={() => onApplyStyle({italic: !activeStyle.italic})}
                title="Italic (Ctrl+I)"
              >
                <Italic className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                active={Boolean(activeStyle.underline)}
                onClick={() => onApplyStyle({underline: !activeStyle.underline})}
                title="Underline (Ctrl+U)"
              >
                <Underline className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                active={Boolean(activeStyle.strike)}
                onClick={() => onApplyStyle({strike: !activeStyle.strike})}
                title="Strikethrough"
              >
                <Strikethrough className="h-3.5 w-3.5" />
              </ToolBtn>
              <label
                title="Text color"
                className="inline-flex h-7 cursor-pointer items-center gap-0.5 rounded border border-slate-200 px-1 dark:border-slate-700"
              >
                <Type className="h-3.5 w-3.5 text-slate-600 dark:text-slate-300" />
                <input
                  type="color"
                  aria-label="Text color"
                  value={activeStyle.color || '#0f172a'}
                  onChange={(e) => onApplyStyle({color: e.target.value})}
                  className="h-4 w-4 cursor-pointer border-0 bg-transparent p-0"
                />
              </label>
              <label
                title="Fill color"
                className="inline-flex h-7 cursor-pointer items-center gap-0.5 rounded border border-slate-200 px-1 dark:border-slate-700"
              >
                <PaintBucket className="h-3.5 w-3.5 text-slate-600 dark:text-slate-300" />
                <input
                  type="color"
                  aria-label="Cell fill color"
                  value={activeStyle.bg || '#ffffff'}
                  onChange={(e) => onApplyStyle({bg: e.target.value})}
                  className="h-4 w-4 cursor-pointer border-0 bg-transparent p-0"
                />
              </label>
              <ToolBtn
                active={Boolean(activeStyle.border)}
                onClick={() => onApplyStyle({border: !activeStyle.border})}
                title="Toggle cell border"
              >
                <Grid3X3 className="h-3.5 w-3.5" />
              </ToolBtn>
            </RibbonGroup>

            <RibbonGroup label="Alignment">
              <ToolBtn
                active={activeStyle.align === 'left'}
                onClick={() => onApplyStyle({align: 'left'})}
                title="Align Left"
              >
                <AlignLeft className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                active={activeStyle.align === 'center'}
                onClick={() => onApplyStyle({align: 'center'})}
                title="Align Center"
              >
                <AlignCenter className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                active={activeStyle.align === 'right'}
                onClick={() => onApplyStyle({align: 'right'})}
                title="Align Right"
              >
                <AlignRight className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                active={Boolean(activeStyle.wrap)}
                onClick={() => onApplyStyle({wrap: !activeStyle.wrap})}
                title="Wrap Text"
              >
                <WrapText className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                active={isMerged}
                onClick={onToggleMerge}
                title={isMerged ? 'Unmerge cells' : 'Merge & Center selected cells'}
                label="Merge"
              >
                <Merge className="h-3.5 w-3.5" />
              </ToolBtn>
            </RibbonGroup>

            <RibbonGroup label="Number">
              <select
                aria-label="Number format"
                value={activeStyle.format ?? 'general'}
                onChange={(e) => onApplyStyle({format: e.target.value as CellNumberFormat})}
                className="h-7 rounded border border-slate-200 bg-slate-50 px-1.5 text-xs text-slate-800 dark:border-slate-700 dark:bg-slate-950 dark:text-slate-100"
              >
                <option value="general">General</option>
                <option value="number">Number (1,234.00)</option>
                <option value="currency">Currency ($)</option>
                <option value="percent">Percent (%)</option>
                <option value="date">Date (YYYY-MM-DD)</option>
                <option value="scientific">Scientific (E+)</option>
                <option value="text">Plain Text</option>
              </select>
              <ToolBtn
                active={activeStyle.format === 'currency'}
                onClick={() => onApplyStyle({format: 'currency'})}
                title="Format as Currency ($)"
                label="$"
              >
                <span />
              </ToolBtn>
              <ToolBtn
                active={activeStyle.format === 'percent'}
                onClick={() => onApplyStyle({format: 'percent'})}
                title="Format as Percent (%)"
                label="%"
              >
                <span />
              </ToolBtn>
              <ToolBtn onClick={() => onStepDecimals(1)} title="Increase decimals" label=".0+">
                <span />
              </ToolBtn>
              <ToolBtn onClick={() => onStepDecimals(-1)} title="Decrease decimals" label=".0-">
                <span />
              </ToolBtn>
            </RibbonGroup>

            <RibbonGroup label="Editing">
              <ToolBtn onClick={() => onInsertFunction('SUM')} title="AutoSum selected cells" label="AutoSum">
                <Sigma className="h-3.5 w-3.5 text-emerald-700 dark:text-emerald-400" />
              </ToolBtn>
              <ToolBtn onClick={onClearSelection} title="Clear selected cells" label="Clear">
                <Eraser className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn onClick={() => onOpenDialog('find')} title="Find & Replace (Ctrl+H)" label="Find">
                <Search className="h-3.5 w-3.5" />
              </ToolBtn>
            </RibbonGroup>
          </>
        )}

        {activeTab === 'insert' && (
          <>
            <RibbonGroup label="Rows & Columns">
              <ToolBtn onClick={() => onInsertRow('below')} title="Insert Row Below" label="+Row">
                <Rows3 className="h-3.5 w-3.5 text-emerald-700 dark:text-emerald-400" />
              </ToolBtn>
              <ToolBtn onClick={onDeleteRow} title="Delete Active Row" label="Del Row">
                <Trash2 className="h-3.5 w-3.5 text-rose-600 dark:text-rose-400" />
              </ToolBtn>
              <ToolBtn onClick={() => onInsertCol('right')} title="Insert Column Right" label="+Col">
                <Columns3 className="h-3.5 w-3.5 text-emerald-700 dark:text-emerald-400" />
              </ToolBtn>
              <ToolBtn onClick={onDeleteCol} title="Delete Active Column" label="Del Col">
                <Trash2 className="h-3.5 w-3.5 text-rose-600 dark:text-rose-400" />
              </ToolBtn>
            </RibbonGroup>

            <RibbonGroup label="Charts">
              {(
                [
                  ['bar', 'Bar Chart'],
                  ['line', 'Line Chart'],
                  ['pie', 'Pie Chart'],
                  ['scatter', 'Scatter Plot'],
                ] as const
              ).map(([kind, label]) => (
                <ToolBtn
                  key={kind}
                  onClick={() => onOpenDialog('chart', kind)}
                  title={`Insert ${label} from selected range`}
                  label={label}
                >
                  <BarChart3 className="h-3.5 w-3.5 text-emerald-700 dark:text-emerald-400" />
                </ToolBtn>
              ))}
            </RibbonGroup>

            <RibbonGroup label="Quick Functions">
              <select
                aria-label="Insert formula function"
                defaultValue=""
                onChange={(e) => {
                  if (e.target.value) {
                    onInsertFunction(e.target.value);
                    e.target.value = '';
                  }
                }}
                className="h-7 rounded border border-slate-200 bg-slate-50 px-2 text-xs font-medium text-slate-800 dark:border-slate-700 dark:bg-slate-950 dark:text-slate-100"
              >
                <option value="" disabled>
                  fx Insert Function...
                </option>
                {QUICK_FUNCTIONS.map((fn) => (
                  <option key={fn} value={fn}>
                    ={fn}(...)
                  </option>
                ))}
              </select>
              <ToolBtn onClick={() => onInsertRow('above')} title="Insert Row Above" label="Row Above">
                <Plus className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn onClick={() => onInsertCol('left')} title="Insert Column Left" label="Col Left">
                <Plus className="h-3.5 w-3.5" />
              </ToolBtn>
            </RibbonGroup>
          </>
        )}

        {activeTab === 'data' && (
          <>
            <RibbonGroup label="Sort & Filter">
              <ToolBtn onClick={() => onSort('asc')} title="Sort column ascending (A to Z)" label="Sort A→Z">
                <ArrowUpAZ className="h-3.5 w-3.5 text-emerald-700 dark:text-emerald-400" />
              </ToolBtn>
              <ToolBtn onClick={() => onSort('desc')} title="Sort column descending (Z to A)" label="Sort Z→A">
                <ArrowDownZA className="h-3.5 w-3.5 text-emerald-700 dark:text-emerald-400" />
              </ToolBtn>
              <ToolBtn
                active={autoFilterEnabled}
                onClick={onToggleAutoFilter}
                title="Toggle column Auto-Filter dropdowns"
                label="Auto-Filter"
              >
                <Filter className="h-3.5 w-3.5" />
              </ToolBtn>
            </RibbonGroup>

            <RibbonGroup label="Rules & Validation">
              <ToolBtn
                onClick={() => onOpenDialog('conditional')}
                title="Manage Conditional Formatting rules"
                label="Conditional Formatting"
              >
                <Palette className="h-3.5 w-3.5 text-amber-600 dark:text-amber-400" />
              </ToolBtn>
              <ToolBtn
                onClick={() => onOpenDialog('validation')}
                title="Set Data Validation rules on selection"
                label="Data Validation"
              >
                <ShieldCheck className="h-3.5 w-3.5 text-blue-600 dark:text-blue-400" />
              </ToolBtn>
            </RibbonGroup>
          </>
        )}

        {activeTab === 'view' && (
          <>
            <RibbonGroup label="Freeze Panes">
              <ToolBtn
                active={freezeRows === 1 && freezeCols === 0}
                onClick={() => onSetFreeze('top-row')}
                title="Freeze Top Row"
                label="Freeze Top Row"
              >
                <Pin className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                active={freezeCols === 1 && freezeRows === 0}
                onClick={() => onSetFreeze('first-col')}
                title="Freeze First Column"
                label="Freeze First Col"
              >
                <Pin className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                onClick={() => onSetFreeze('selection')}
                title="Freeze Rows & Columns above/left of active cell"
                label="Freeze at Cell"
              >
                <Pin className="h-3.5 w-3.5" />
              </ToolBtn>
              {(freezeRows > 0 || freezeCols > 0) && (
                <ToolBtn onClick={() => onSetFreeze('none')} title="Unfreeze Panes" label="Unfreeze">
                  <span />
                </ToolBtn>
              )}
            </RibbonGroup>

            <RibbonGroup label="Show">
              <ToolBtn
                active={showGridlines}
                onClick={onToggleGridlines}
                title="Toggle sheet gridlines"
                label="Gridlines"
              >
                <Grid3X3 className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                active={showFormulaBar}
                onClick={onToggleFormulaBar}
                title="Toggle Formula Bar"
                label="Formula Bar"
              >
                <Sigma className="h-3.5 w-3.5" />
              </ToolBtn>
            </RibbonGroup>

            <RibbonGroup label="Zoom">
              <ToolBtn
                onClick={() => onChangeZoom(Math.max(0.75, Number((zoom - 0.1).toFixed(2))))}
                title="Zoom Out"
              >
                <ZoomOut className="h-3.5 w-3.5" />
              </ToolBtn>
              <span className="min-w-[44px] text-center text-xs font-semibold tabular-nums text-slate-700 dark:text-slate-200">
                {Math.round(zoom * 100)}%
              </span>
              <ToolBtn
                onClick={() => onChangeZoom(Math.min(1.5, Number((zoom + 0.1).toFixed(2))))}
                title="Zoom In"
              >
                <ZoomIn className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn onClick={() => onChangeZoom(1)} title="Reset to 100%" label="100%">
                <span />
              </ToolBtn>
            </RibbonGroup>
          </>
        )}
      </div>
    </div>
  );
};
