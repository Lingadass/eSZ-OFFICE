import React, {useEffect, useState} from 'react';
import {Download, FileDown, FolderPlus, Search, Trash2, X} from 'lucide-react';
import {
  CONDITIONAL_PRESETS,
  type ChartKind,
  type ConditionalOperator,
  type ConditionalRule,
  type DataValidationRule,
  type ValidationKind,
} from './sheetsTypes';

export type SheetsDialogKind = 'find' | 'conditional' | 'validation' | 'chart' | 'export';

interface SheetsDialogsProps {
  openDialog: SheetsDialogKind | null;
  defaultChartKind: ChartKind;
  selectionRangeLabel: string;
  fileName: string;
  conditionalRules: ConditionalRule[];
  validations: DataValidationRule[];
  onClose: () => void;
  onFindNext: (query: string, caseSensitive: boolean) => string;
  onFindPrev?: (query: string, caseSensitive: boolean) => string;
  onReplaceAll: (query: string, replacement: string, caseSensitive: boolean) => string;
  onAddConditionalRule: (rule: Omit<ConditionalRule, 'id'>) => void;
  onDeleteConditionalRule: (id: string) => void;
  onSaveValidation: (rule: Omit<DataValidationRule, 'id'>) => void;
  onDeleteValidation: (id: string) => void;
  onCreateChart: (spec: {title: string; kind: ChartKind; range: string}) => void;
  onExportFormat: (
    format: '.xlsx' | '.xls' | '.csv' | '.pdf',
    target: 'device' | 'vault',
  ) => Promise<void>;
  readOnly?: boolean;
}

const EXPORT_FORMATS: Array<{
  ext: '.xlsx' | '.xls' | '.csv' | '.pdf';
  title: string;
  desc: string;
}> = [
  {ext: '.xlsx', title: 'Excel Workbook (.xlsx)', desc: 'Modern OpenXML workbook with multi-sheet formulas & merges'},
  {ext: '.xls', title: 'Legacy Excel 97-2003 (.xls)', desc: 'Classic binary spreadsheet format for older software'},
  {ext: '.csv', title: 'Comma-Separated Values (.csv)', desc: 'Plain text table of the active sheet for universal import'},
  {ext: '.pdf', title: 'PDF Document (.pdf)', desc: 'Landscape printable table across all populated sheets'},
];

export const SheetsDialogs: React.FC<SheetsDialogsProps> = ({
  openDialog,
  defaultChartKind,
  selectionRangeLabel,
  fileName,
  conditionalRules,
  validations,
  onClose,
  onFindNext,
  onFindPrev,
  onReplaceAll,
  onAddConditionalRule,
  onDeleteConditionalRule,
  onSaveValidation,
  onDeleteValidation,
  onCreateChart,
  onExportFormat,
  readOnly = false,
}) => {
  // Find & Replace
  const [findQuery, setFindQuery] = useState('');
  const [replaceText, setReplaceText] = useState('');
  const [caseSensitive, setCaseSensitive] = useState(false);
  const [findStatus, setFindStatus] = useState('');

  // Conditional Formatting
  const [condRange, setCondRange] = useState(selectionRangeLabel);
  const [condOp, setCondOp] = useState<ConditionalOperator>('gt');
  const [condVal1, setCondVal1] = useState('1000');
  const [condVal2, setCondVal2] = useState('5000');
  const [condPresetIdx, setCondPresetIdx] = useState(0);

  // Data Validation
  const [valRange, setValRange] = useState(selectionRangeLabel);
  const [valKind, setValKind] = useState<ValidationKind>('number-range');
  const [valMin, setValMin] = useState('0');
  const [valMax, setValMax] = useState('10000');
  const [valList, setValList] = useState('Planned, Approved, Paid');

  // Insert Chart
  const [chartTitle, setChartTitle] = useState('Budget Overview');
  const [chartKind, setChartKind] = useState<ChartKind>(defaultChartKind);
  const [chartRange, setChartRange] = useState(selectionRangeLabel);

  const [busy, setBusy] = useState(false);

  useEffect(() => {
    if (!openDialog) return;
    setCondRange(selectionRangeLabel);
    setValRange(selectionRangeLabel);
    setChartRange(selectionRangeLabel.includes(':') ? selectionRangeLabel : 'A1:C4');
    setChartKind(defaultChartKind);
    setFindStatus('');
  }, [openDialog, selectionRangeLabel, defaultChartKind]);

  useEffect(() => {
    if (!openDialog) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [openDialog, onClose]);

  if (!openDialog) return null;

  const titles: Record<SheetsDialogKind, string> = {
    find: readOnly ? 'Find in Sheet' : 'Find & Replace in Sheet',
    conditional: 'Conditional Formatting Rules',
    validation: 'Data Validation',
    chart: 'Insert Embedded Chart',
    export: 'Save As & Export Workbook',
  };

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/50 p-4 backdrop-blur-[1px]"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div
        role="dialog"
        aria-modal="true"
        aria-labelledby="sheets-dialog-title"
        className="flex max-h-[88vh] w-full max-w-lg flex-col overflow-hidden rounded-xl border border-slate-200 bg-white shadow-2xl dark:border-slate-700 dark:bg-slate-900"
      >
        <div className="flex items-center justify-between border-b border-slate-200 px-4 py-3 dark:border-slate-800">
          <h2 id="sheets-dialog-title" className="text-sm font-semibold text-slate-900 dark:text-slate-100">
            {titles[openDialog]}
          </h2>
          <button
            type="button"
            onClick={onClose}
            aria-label="Close dialog"
            className="rounded-md p-1 text-slate-500 hover:bg-slate-100 dark:hover:bg-slate-800"
          >
            <X className="h-4 w-4" />
          </button>
        </div>

        <div className="overflow-y-auto p-4 text-sm">
          {openDialog === 'find' && (
            <div className="space-y-3">
              <div>
                <label className="mb-1 block text-xs font-medium text-slate-600 dark:text-slate-300">
                  Find value or formula text
                </label>
                <div className="relative">
                  <Search className="pointer-events-none absolute top-2.5 left-2.5 h-4 w-4 text-slate-400" />
                  <input
                    type="text"
                    autoFocus
                    value={findQuery}
                    onChange={(e) => setFindQuery(e.target.value)}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter' && findQuery) {
                        e.preventDefault();
                        setFindStatus(
                          e.shiftKey && onFindPrev
                            ? onFindPrev(findQuery, caseSensitive)
                            : onFindNext(findQuery, caseSensitive),
                        );
                      }
                    }}
                    placeholder="Search cells..."
                    className="w-full rounded-lg border border-slate-300 bg-white py-2 pr-3 pl-8 text-sm dark:border-slate-700 dark:bg-slate-950"
                  />
                </div>
              </div>
              {!readOnly && (
                <div>
                  <label className="mb-1 block text-xs font-medium text-slate-600 dark:text-slate-300">
                    Replace with
                  </label>
                  <input
                    type="text"
                    value={replaceText}
                    onChange={(e) => setReplaceText(e.target.value)}
                    placeholder="Replacement..."
                    className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm dark:border-slate-700 dark:bg-slate-950"
                  />
                </div>
              )}
              <label className="flex items-center gap-2 text-xs text-slate-700 dark:text-slate-300">
                <input
                  type="checkbox"
                  checked={caseSensitive}
                  onChange={(e) => setCaseSensitive(e.target.checked)}
                  className="h-3.5 w-3.5 accent-emerald-600"
                />
                Match case
              </label>
              {findStatus && (
                <p className="rounded-md bg-slate-100 px-3 py-1.5 text-xs text-slate-700 dark:bg-slate-800 dark:text-slate-200">
                  {findStatus}
                </p>
              )}
              <div className="flex justify-end gap-2 pt-2">
                {onFindPrev && (
                  <button
                    type="button"
                    disabled={!findQuery}
                    onClick={() => setFindStatus(onFindPrev(findQuery, caseSensitive))}
                    className="rounded-lg border border-slate-300 px-3 py-1.5 text-xs font-semibold text-slate-700 hover:bg-slate-100 disabled:opacity-40 dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800"
                  >
                    Previous
                  </button>
                )}
                <button
                  type="button"
                  disabled={!findQuery}
                  onClick={() => setFindStatus(onFindNext(findQuery, caseSensitive))}
                  className={`rounded-lg px-3.5 py-1.5 text-xs font-semibold disabled:opacity-40 ${
                    readOnly
                      ? 'bg-emerald-700 text-white hover:bg-emerald-800 dark:bg-emerald-600'
                      : 'border border-slate-300 text-slate-700 hover:bg-slate-100 dark:border-slate-700 dark:text-slate-200'
                  }`}
                >
                  Find Next
                </button>
                {!readOnly && (
                  <button
                    type="button"
                    disabled={!findQuery}
                    onClick={() => setFindStatus(onReplaceAll(findQuery, replaceText, caseSensitive))}
                    className="rounded-lg bg-emerald-700 px-3.5 py-1.5 text-xs font-semibold text-white hover:bg-emerald-800 disabled:opacity-40 dark:bg-emerald-600"
                  >
                    Replace All
                  </button>
                )}
              </div>
            </div>
          )}

          {openDialog === 'conditional' && (
            <div className="space-y-4">
              <form
                className="space-y-3 rounded-lg border border-slate-200 bg-slate-50 p-3 dark:border-slate-800 dark:bg-slate-950"
                onSubmit={(e) => {
                  e.preventDefault();
                  const preset = CONDITIONAL_PRESETS[condPresetIdx] ?? CONDITIONAL_PRESETS[0];
                  onAddConditionalRule({
                    range: condRange.trim().toUpperCase() || 'A1:Z40',
                    operator: condOp,
                    value1: condVal1,
                    value2: condOp === 'between' ? condVal2 : undefined,
                    bg: preset.bg,
                    color: preset.color,
                  });
                }}
              >
                <div className="grid grid-cols-2 gap-2.5">
                  <div>
                    <label className="mb-1 block text-xs font-medium text-slate-600 dark:text-slate-300">
                      Cell range
                    </label>
                    <input
                      type="text"
                      required
                      value={condRange}
                      onChange={(e) => setCondRange(e.target.value)}
                      placeholder="B2:D10"
                      className="w-full rounded-md border border-slate-300 bg-white px-2.5 py-1.5 font-mono text-xs dark:border-slate-700 dark:bg-slate-900"
                    />
                  </div>
                  <div>
                    <label className="mb-1 block text-xs font-medium text-slate-600 dark:text-slate-300">
                      Condition
                    </label>
                    <select
                      aria-label="Conditional operator"
                      value={condOp}
                      onChange={(e) => setCondOp(e.target.value as ConditionalOperator)}
                      className="w-full rounded-md border border-slate-300 bg-white px-2 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-900"
                    >
                      <option value="gt">Greater than (&gt;)</option>
                      <option value="lt">Less than (&lt;)</option>
                      <option value="between">Between</option>
                      <option value="eq">Equal to (=)</option>
                      <option value="contains">Text contains</option>
                    </select>
                  </div>
                </div>

                <div className="grid grid-cols-2 gap-2.5">
                  <div>
                    <label className="mb-1 block text-xs font-medium text-slate-600 dark:text-slate-300">
                      Value
                    </label>
                    <input
                      type="text"
                      required
                      value={condVal1}
                      onChange={(e) => setCondVal1(e.target.value)}
                      className="w-full rounded-md border border-slate-300 bg-white px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-900"
                    />
                  </div>
                  {condOp === 'between' ? (
                    <div>
                      <label className="mb-1 block text-xs font-medium text-slate-600 dark:text-slate-300">
                        And value
                      </label>
                      <input
                        type="text"
                        required
                        value={condVal2}
                        onChange={(e) => setCondVal2(e.target.value)}
                        className="w-full rounded-md border border-slate-300 bg-white px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-900"
                      />
                    </div>
                  ) : (
                    <div>
                      <label className="mb-1 block text-xs font-medium text-slate-600 dark:text-slate-300">
                        Highlight style
                      </label>
                      <select
                        aria-label="Highlight style"
                        value={condPresetIdx}
                        onChange={(e) => setCondPresetIdx(Number(e.target.value))}
                        className="w-full rounded-md border border-slate-300 bg-white px-2 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-900"
                      >
                        {CONDITIONAL_PRESETS.map((p, idx) => (
                          <option key={p.label} value={idx}>
                            {p.label}
                          </option>
                        ))}
                      </select>
                    </div>
                  )}
                </div>

                <div className="flex justify-end">
                  <button
                    type="submit"
                    className="rounded-lg bg-emerald-700 px-3.5 py-1.5 text-xs font-semibold text-white hover:bg-emerald-800 dark:bg-emerald-600"
                  >
                    Add Rule
                  </button>
                </div>
              </form>

              <div className="space-y-1.5">
                <div className="text-xs font-semibold text-slate-700 dark:text-slate-300">
                  Active Rules ({conditionalRules.length})
                </div>
                {conditionalRules.length === 0 ? (
                  <p className="text-xs text-slate-500">No conditional formatting rules on this sheet.</p>
                ) : (
                  conditionalRules.map((r) => (
                    <div
                      key={r.id}
                      className="flex items-center justify-between rounded-lg border border-slate-200 px-3 py-2 text-xs dark:border-slate-800"
                    >
                      <div className="flex items-center gap-2">
                        <span
                          className="inline-block rounded px-1.5 py-0.5 font-mono text-[11px] font-semibold"
                          style={{backgroundColor: r.bg, color: r.color}}
                        >
                          {r.range}
                        </span>
                        <span className="text-slate-700 dark:text-slate-300">
                          {r.operator} {r.value1}
                          {r.value2 ? ` – ${r.value2}` : ''}
                        </span>
                      </div>
                      <button
                        type="button"
                        onClick={() => onDeleteConditionalRule(r.id)}
                        aria-label={`Delete conditional rule ${r.range}`}
                        className="rounded p-1 text-slate-400 hover:text-rose-600"
                      >
                        <Trash2 className="h-3.5 w-3.5" />
                      </button>
                    </div>
                  ))
                )}
              </div>
            </div>
          )}

          {openDialog === 'validation' && (
            <div className="space-y-4">
              <form
                className="space-y-3 rounded-lg border border-slate-200 bg-slate-50 p-3 dark:border-slate-800 dark:bg-slate-950"
                onSubmit={(e) => {
                  e.preventDefault();
                  onSaveValidation({
                    range: valRange.trim().toUpperCase() || selectionRangeLabel,
                    kind: valKind,
                    min: valKind !== 'list' ? Number(valMin) : undefined,
                    max: valKind !== 'list' ? Number(valMax) : undefined,
                    listItems:
                      valKind === 'list'
                        ? valList
                            .split(',')
                            .map((s) => s.trim())
                            .filter(Boolean)
                        : undefined,
                  });
                }}
              >
                <div className="grid grid-cols-2 gap-2.5">
                  <div>
                    <label className="mb-1 block text-xs font-medium text-slate-600 dark:text-slate-300">
                      Cell range
                    </label>
                    <input
                      type="text"
                      required
                      value={valRange}
                      onChange={(e) => setValRange(e.target.value)}
                      className="w-full rounded-md border border-slate-300 bg-white px-2.5 py-1.5 font-mono text-xs dark:border-slate-700 dark:bg-slate-900"
                    />
                  </div>
                  <div>
                    <label className="mb-1 block text-xs font-medium text-slate-600 dark:text-slate-300">
                      Validation type
                    </label>
                    <select
                      aria-label="Validation type"
                      value={valKind}
                      onChange={(e) => setValKind(e.target.value as ValidationKind)}
                      className="w-full rounded-md border border-slate-300 bg-white px-2 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-900"
                    >
                      <option value="number-range">Number range</option>
                      <option value="list">Dropdown list</option>
                      <option value="text-length">Text length</option>
                    </select>
                  </div>
                </div>

                {valKind === 'list' ? (
                  <div>
                    <label className="mb-1 block text-xs font-medium text-slate-600 dark:text-slate-300">
                      Comma-separated allowed options
                    </label>
                    <input
                      type="text"
                      required
                      value={valList}
                      onChange={(e) => setValList(e.target.value)}
                      placeholder="Open, In Progress, Done"
                      className="w-full rounded-md border border-slate-300 bg-white px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-900"
                    />
                  </div>
                ) : (
                  <div className="grid grid-cols-2 gap-2.5">
                    <div>
                      <label className="mb-1 block text-xs font-medium text-slate-600 dark:text-slate-300">
                        Minimum
                      </label>
                      <input
                        type="number"
                        value={valMin}
                        onChange={(e) => setValMin(e.target.value)}
                        className="w-full rounded-md border border-slate-300 bg-white px-2.5 py-1.5 text-xs tabular-nums dark:border-slate-700 dark:bg-slate-900"
                      />
                    </div>
                    <div>
                      <label className="mb-1 block text-xs font-medium text-slate-600 dark:text-slate-300">
                        Maximum
                      </label>
                      <input
                        type="number"
                        value={valMax}
                        onChange={(e) => setValMax(e.target.value)}
                        className="w-full rounded-md border border-slate-300 bg-white px-2.5 py-1.5 text-xs tabular-nums dark:border-slate-700 dark:bg-slate-900"
                      />
                    </div>
                  </div>
                )}

                <div className="flex justify-end">
                  <button
                    type="submit"
                    className="rounded-lg bg-emerald-700 px-3.5 py-1.5 text-xs font-semibold text-white hover:bg-emerald-800 dark:bg-emerald-600"
                  >
                    Apply Validation
                  </button>
                </div>
              </form>

              <div className="space-y-1.5">
                <div className="text-xs font-semibold text-slate-700 dark:text-slate-300">
                  Active Rules ({validations.length})
                </div>
                {validations.map((v) => (
                  <div
                    key={v.id}
                    className="flex items-center justify-between rounded-lg border border-slate-200 px-3 py-2 text-xs dark:border-slate-800"
                  >
                    <span>
                      <strong className="font-mono">{v.range}</strong> — {v.kind}
                      {v.listItems ? ` (${v.listItems.join(', ')})` : ` (${v.min}–${v.max})`}
                    </span>
                    <button
                      type="button"
                      onClick={() => onDeleteValidation(v.id)}
                      aria-label={`Delete validation ${v.range}`}
                      className="rounded p-1 text-slate-400 hover:text-rose-600"
                    >
                      <Trash2 className="h-3.5 w-3.5" />
                    </button>
                  </div>
                ))}
              </div>
            </div>
          )}

          {openDialog === 'chart' && (
            <form
              className="space-y-3"
              onSubmit={(e) => {
                e.preventDefault();
                onCreateChart({
                  title: chartTitle.trim() || 'Sheet Chart',
                  kind: chartKind,
                  range: chartRange.trim().toUpperCase() || 'A1:C4',
                });
                onClose();
              }}
            >
              <div>
                <label className="mb-1 block text-xs font-medium text-slate-600 dark:text-slate-300">
                  Chart title
                </label>
                <input
                  type="text"
                  required
                  value={chartTitle}
                  onChange={(e) => setChartTitle(e.target.value)}
                  className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm dark:border-slate-700 dark:bg-slate-950"
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="mb-1 block text-xs font-medium text-slate-600 dark:text-slate-300">
                    Chart type
                  </label>
                  <select
                    aria-label="Chart type"
                    value={chartKind}
                    onChange={(e) => setChartKind(e.target.value as ChartKind)}
                    className="w-full rounded-lg border border-slate-300 bg-white px-2.5 py-2 text-sm dark:border-slate-700 dark:bg-slate-950"
                  >
                    <option value="bar">Bar / Column</option>
                    <option value="line">Line</option>
                    <option value="pie">Pie</option>
                    <option value="scatter">Scatter Plot</option>
                  </select>
                </div>
                <div>
                  <label className="mb-1 block text-xs font-medium text-slate-600 dark:text-slate-300">
                    Data range
                  </label>
                  <input
                    type="text"
                    required
                    value={chartRange}
                    onChange={(e) => setChartRange(e.target.value)}
                    placeholder="A1:C4"
                    className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 font-mono text-sm dark:border-slate-700 dark:bg-slate-950"
                  />
                </div>
              </div>

              <p className="text-xs text-slate-500 dark:text-slate-400">
                First column supplies category labels; subsequent columns plot numeric series automatically.
              </p>

              <div className="flex justify-end gap-2 pt-2">
                <button
                  type="button"
                  onClick={onClose}
                  className="rounded-lg border border-slate-300 px-3 py-1.5 text-xs font-semibold text-slate-700 dark:border-slate-700 dark:text-slate-200"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="rounded-lg bg-emerald-700 px-4 py-1.5 text-xs font-semibold text-white hover:bg-emerald-800 dark:bg-emerald-600"
                >
                  Create Chart
                </button>
              </div>
            </form>
          )}

          {openDialog === 'export' && (
            <div className="space-y-2.5">
              <p className="text-xs text-slate-600 dark:text-slate-400">
                Export <span className="font-semibold text-slate-800 dark:text-slate-200">{fileName}</span> completely offline.
              </p>
              <div className="space-y-2">
                {EXPORT_FORMATS.map((fmt) => (
                  <div
                    key={fmt.ext}
                    className="flex flex-col gap-2 rounded-lg border border-slate-200 p-3 sm:flex-row sm:items-center sm:justify-between dark:border-slate-800"
                  >
                    <div className="min-w-0">
                      <div className="text-xs font-semibold text-slate-900 dark:text-slate-100">
                        {fmt.title}
                      </div>
                      <div className="text-[11px] text-slate-500 dark:text-slate-400">
                        {fmt.desc}
                      </div>
                    </div>
                    <div className="flex shrink-0 items-center gap-1.5">
                      <button
                        type="button"
                        disabled={busy}
                        onClick={async () => {
                          setBusy(true);
                          try {
                            await onExportFormat(fmt.ext, 'vault');
                            onClose();
                          } finally {
                            setBusy(false);
                          }
                        }}
                        className="inline-flex items-center gap-1 whitespace-nowrap rounded-md border border-slate-300 px-2.5 py-1.5 text-xs font-semibold text-slate-700 hover:bg-slate-100 disabled:opacity-40 dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800"
                      >
                        <FolderPlus className="h-3.5 w-3.5" />
                        Vault
                      </button>
                      <button
                        type="button"
                        disabled={busy}
                        onClick={async () => {
                          setBusy(true);
                          try {
                            await onExportFormat(fmt.ext, 'device');
                            onClose();
                          } finally {
                            setBusy(false);
                          }
                        }}
                        className="inline-flex items-center gap-1 whitespace-nowrap rounded-md bg-emerald-700 px-2.5 py-1.5 text-xs font-semibold text-white hover:bg-emerald-800 disabled:opacity-40 dark:bg-emerald-600"
                      >
                        {fmt.ext === '.pdf' ? (
                          <FileDown className="h-3.5 w-3.5" />
                        ) : (
                          <Download className="h-3.5 w-3.5" />
                        )}
                        Save
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};
