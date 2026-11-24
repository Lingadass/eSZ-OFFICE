import React, {useEffect, useState} from 'react';
import {Download, FileText, Image as ImageIcon, LayoutTemplate, Wand2, X} from 'lucide-react';
import {DECK_TEMPLATES, parseOutlineToSlides} from './slidesPure';
import type {
  SlideChartKind,
  SlideDialogKind,
  SlideMasterSettings,
} from './slidesTypes';

interface SlidesDialogsProps {
  openDialog: SlideDialogKind;
  defaultChartKind: SlideChartKind;
  deckTitle: string;
  master: SlideMasterSettings;
  onClose: () => void;
  onApplyTemplate: (templateId: string) => void;
  onGenerateFromOutline: (outlineText: string, replaceAll: boolean) => void;
  onCreateChart: (spec: {
    kind: SlideChartKind;
    title: string;
    labels: string[];
    values: number[];
  }) => void;
  onCreateTable: (rows: number, cols: number) => void;
  onUpdateMaster: (patch: Partial<SlideMasterSettings>) => void;
  onExportFormat: (format: 'pptx' | 'pdf' | 'png', destination: 'device' | 'vault') => void;
}

const SAMPLE_OUTLINE = `# Q4 Offline Suite Launch
## Engineering & Operations Summary
> Open with our zero-cloud privacy promise.

# Key Milestones Achieved
- Full PDF viewer, redaction, signing, and OCR
- Word 2007-style Documents with real DOCX export
- Multi-sheet Spreadsheets with 40+ Excel formulas
- Interactive Presentations with Konva & PPTX export
> Walk through each completed pillar briefly.

# Next Steps
- Run Play Store Trusted Web Activity packaging
- Finalize Digital Asset Links verification
- Publish offline-first release to beta testers`;

export const SlidesDialogs: React.FC<SlidesDialogsProps> = ({
  openDialog,
  defaultChartKind,
  deckTitle,
  master,
  onClose,
  onApplyTemplate,
  onGenerateFromOutline,
  onCreateChart,
  onCreateTable,
  onUpdateMaster,
  onExportFormat,
}) => {
  const [outlineText, setOutlineText] = useState(SAMPLE_OUTLINE);
  const [replaceAll, setReplaceAll] = useState(false);

  const [chartKind, setChartKind] = useState<SlideChartKind>(defaultChartKind);
  const [chartTitle, setChartTitle] = useState('Quarterly Performance');
  const [chartRowsText, setChartRowsText] = useState(
    'Q1, 42\nQ2, 68\nQ3, 84\nQ4, 112',
  );

  const [tableRows, setTableRows] = useState(4);
  const [tableCols, setTableCols] = useState(3);

  const [footerDraft, setFooterDraft] = useState(master.footerText);
  const [showNumbersDraft, setShowNumbersDraft] = useState(master.showSlideNumbers);
  const [showAccentDraft, setShowAccentDraft] = useState(master.showAccentBar);

  useEffect(() => {
    setChartKind(defaultChartKind);
  }, [defaultChartKind, openDialog]);

  useEffect(() => {
    setFooterDraft(master.footerText);
    setShowNumbersDraft(master.showSlideNumbers);
    setShowAccentDraft(master.showAccentBar);
  }, [master, openDialog]);

  if (!openDialog) return null;

  const previewSlideCount = parseOutlineToSlides(outlineText).length;

  const titleMap: Record<Exclude<SlideDialogKind, null>, string> = {
    templates: 'Presentation Templates',
    outline: 'Outline to Slides Generator',
    chart: 'Insert Data Chart',
    table: 'Insert Table',
    master: 'Slide Master Settings',
    export: `Export "${deckTitle}"`,
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-label={titleMap[openDialog]}
      className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/60 p-4"
      onClick={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div className="flex max-h-[88vh] w-full max-w-xl flex-col overflow-hidden rounded-2xl border border-slate-200 bg-white shadow-2xl dark:border-slate-800 dark:bg-slate-900">
        <div className="flex items-center justify-between border-b border-slate-200 px-5 py-3.5 dark:border-slate-800">
          <h2 className="text-sm font-bold text-slate-900 dark:text-slate-100">
            {titleMap[openDialog]}
          </h2>
          <button
            type="button"
            onClick={onClose}
            aria-label="Close dialog"
            className="inline-flex h-8 w-8 items-center justify-center rounded-lg text-slate-500 hover:bg-slate-100 dark:hover:bg-slate-800"
          >
            <X className="h-4 w-4" />
          </button>
        </div>

        <div className="overflow-y-auto p-5">
          {openDialog === 'templates' && (
            <div className="space-y-3">
              <p className="text-xs text-slate-600 dark:text-slate-300">
                Choose a curated presentation template to replace the active deck with ready-made layouts and themes.
              </p>
              <div className="grid gap-3 sm:grid-cols-2">
                {DECK_TEMPLATES.map((tpl) => (
                  <button
                    key={tpl.id}
                    type="button"
                    onClick={() => {
                      onApplyTemplate(tpl.id);
                      onClose();
                    }}
                    className="flex flex-col items-start rounded-xl border border-slate-200 bg-slate-50 p-3.5 text-left transition-colors hover:border-orange-500 hover:bg-orange-50/40 dark:border-slate-800 dark:bg-slate-950 dark:hover:border-orange-500"
                  >
                    <div className="flex w-full items-center justify-between gap-2">
                      <span className="text-sm font-bold text-slate-900 dark:text-slate-100">
                        {tpl.name}
                      </span>
                      <LayoutTemplate className="h-4 w-4 shrink-0 text-orange-600 dark:text-orange-400" />
                    </div>
                    <span className="mt-0.5 text-[11px] font-medium text-slate-500">
                      {tpl.category}
                    </span>
                    <p className="mt-2 text-xs leading-relaxed text-slate-600 dark:text-slate-300">
                      {tpl.description}
                    </p>
                  </button>
                ))}
              </div>
            </div>
          )}

          {openDialog === 'outline' && (
            <form
              className="space-y-3"
              onSubmit={(e) => {
                e.preventDefault();
                if (!outlineText.trim()) return;
                onGenerateFromOutline(outlineText, replaceAll);
                onClose();
              }}
            >
              <p className="text-xs text-slate-600 dark:text-slate-300">
                Paste Markdown or bulleted notes. Use <code className="font-mono"># Title</code> for each slide,{' '}
                <code className="font-mono">- Bullet</code> for points, and{' '}
                <code className="font-mono">&gt; Note</code> for speaker notes.
              </p>
              <textarea
                aria-label="Presentation outline text"
                required
                rows={10}
                value={outlineText}
                onChange={(e) => setOutlineText(e.target.value)}
                className="w-full rounded-lg border border-slate-300 bg-white p-3 font-mono text-xs leading-relaxed dark:border-slate-700 dark:bg-slate-950"
              />
              <div className="flex flex-wrap items-center justify-between gap-2">
                <label className="flex cursor-pointer items-center gap-2 text-xs text-slate-700 dark:text-slate-300">
                  <input
                    type="checkbox"
                    checked={replaceAll}
                    onChange={(e) => setReplaceAll(e.target.checked)}
                  />
                  Replace existing slides instead of appending
                </label>
                <span className="text-xs font-medium tabular-nums text-slate-500">
                  Will generate {previewSlideCount} slide{previewSlideCount === 1 ? '' : 's'}
                </span>
              </div>
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
                  className="inline-flex items-center gap-1.5 rounded-lg bg-orange-600 px-4 py-1.5 text-xs font-semibold text-white hover:bg-orange-700"
                >
                  <Wand2 className="h-3.5 w-3.5" />
                  Generate Slides
                </button>
              </div>
            </form>
          )}

          {openDialog === 'chart' && (
            <form
              className="space-y-3"
              onSubmit={(e) => {
                e.preventDefault();
                const parsed = chartRowsText
                  .split('\n')
                  .map((line) => line.split(',').map((s) => s.trim()))
                  .filter((parts) => parts.length >= 2 && parts[0].length > 0);
                const labels = parsed.map((p) => p[0]);
                const values = parsed.map((p) => Number(p[1]) || 0);
                onCreateChart({
                  kind: chartKind,
                  title: chartTitle.trim() || 'Slide Chart',
                  labels: labels.length > 0 ? labels : ['Q1', 'Q2', 'Q3'],
                  values: values.length > 0 ? values : [25, 50, 75],
                });
                onClose();
              }}
            >
              <div className="grid grid-cols-2 gap-3">
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
                <div>
                  <label className="mb-1 block text-xs font-medium text-slate-600 dark:text-slate-300">
                    Chart type
                  </label>
                  <select
                    aria-label="Slide chart type"
                    value={chartKind}
                    onChange={(e) => setChartKind(e.target.value as SlideChartKind)}
                    className="w-full rounded-lg border border-slate-300 bg-white px-2.5 py-2 text-sm dark:border-slate-700 dark:bg-slate-950"
                  >
                    <option value="bar">Bar Chart</option>
                    <option value="line">Line Chart</option>
                    <option value="pie">Pie Chart</option>
                  </select>
                </div>
              </div>
              <div>
                <label className="mb-1 block text-xs font-medium text-slate-600 dark:text-slate-300">
                  Data rows (Label, Value per line)
                </label>
                <textarea
                  aria-label="Chart data rows"
                  rows={5}
                  value={chartRowsText}
                  onChange={(e) => setChartRowsText(e.target.value)}
                  className="w-full rounded-lg border border-slate-300 bg-white p-2.5 font-mono text-xs dark:border-slate-700 dark:bg-slate-950"
                />
              </div>
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
                  className="rounded-lg bg-orange-600 px-4 py-1.5 text-xs font-semibold text-white hover:bg-orange-700"
                >
                  Insert Chart
                </button>
              </div>
            </form>
          )}

          {openDialog === 'table' && (
            <form
              className="space-y-4"
              onSubmit={(e) => {
                e.preventDefault();
                onCreateTable(tableRows, tableCols);
                onClose();
              }}
            >
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="mb-1 block text-xs font-medium text-slate-600 dark:text-slate-300">
                    Rows (2–8)
                  </label>
                  <input
                    type="number"
                    min={2}
                    max={8}
                    value={tableRows}
                    onChange={(e) => setTableRows(Number(e.target.value) || 3)}
                    className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm tabular-nums dark:border-slate-700 dark:bg-slate-950"
                  />
                </div>
                <div>
                  <label className="mb-1 block text-xs font-medium text-slate-600 dark:text-slate-300">
                    Columns (2–6)
                  </label>
                  <input
                    type="number"
                    min={2}
                    max={6}
                    value={tableCols}
                    onChange={(e) => setTableCols(Number(e.target.value) || 3)}
                    className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm tabular-nums dark:border-slate-700 dark:bg-slate-950"
                  />
                </div>
              </div>
              <div className="flex justify-end gap-2">
                <button
                  type="button"
                  onClick={onClose}
                  className="rounded-lg border border-slate-300 px-3 py-1.5 text-xs font-semibold text-slate-700 dark:border-slate-700 dark:text-slate-200"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="rounded-lg bg-orange-600 px-4 py-1.5 text-xs font-semibold text-white hover:bg-orange-700"
                >
                  Insert Table
                </button>
              </div>
            </form>
          )}

          {openDialog === 'master' && (
            <form
              className="space-y-3"
              onSubmit={(e) => {
                e.preventDefault();
                onUpdateMaster({
                  footerText: footerDraft,
                  showSlideNumbers: showNumbersDraft,
                  showAccentBar: showAccentDraft,
                });
                onClose();
              }}
            >
              <p className="text-xs text-slate-600 dark:text-slate-300">
                Slide Master settings apply consistently across every slide in this presentation and carry over to PPTX and PDF exports.
              </p>
              <div>
                <label className="mb-1 block text-xs font-medium text-slate-600 dark:text-slate-300">
                  Footer text
                </label>
                <input
                  type="text"
                  value={footerDraft}
                  onChange={(e) => setFooterDraft(e.target.value)}
                  className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm dark:border-slate-700 dark:bg-slate-950"
                />
              </div>
              <label className="flex cursor-pointer items-center gap-2 text-xs text-slate-700 dark:text-slate-200">
                <input
                  type="checkbox"
                  checked={showNumbersDraft}
                  onChange={(e) => setShowNumbersDraft(e.target.checked)}
                />
                Display slide numbers on bottom right
              </label>
              <label className="flex cursor-pointer items-center gap-2 text-xs text-slate-700 dark:text-slate-200">
                <input
                  type="checkbox"
                  checked={showAccentDraft}
                  onChange={(e) => setShowAccentDraft(e.target.checked)}
                />
                Display top theme accent bar
              </label>
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
                  className="rounded-lg bg-orange-600 px-4 py-1.5 text-xs font-semibold text-white hover:bg-orange-700"
                >
                  Apply to All Slides
                </button>
              </div>
            </form>
          )}

          {openDialog === 'export' && (
            <div className="space-y-3">
              <p className="text-xs text-slate-600 dark:text-slate-300">
                All conversions run locally on your device. Choose an output format and destination.
              </p>
              {(
                [
                  {
                    fmt: 'pptx' as const,
                    title: 'PowerPoint Presentation (.pptx)',
                    desc: 'OpenXML presentation with editable text, shapes, tables, charts, and speaker notes.',
                    icon: Download,
                  },
                  {
                    fmt: 'pdf' as const,
                    title: 'Widescreen PDF Handout (.pdf)',
                    desc: '16:9 vector PDF ready for sharing or printing in the PDF module.',
                    icon: FileText,
                  },
                  {
                    fmt: 'png' as const,
                    title: 'Active Slide Image (.png)',
                    desc: 'High-resolution 2x PNG snapshot of the currently selected slide.',
                    icon: ImageIcon,
                  },
                ] as const
              ).map(({fmt, title, desc, icon: Icon}) => (
                <div
                  key={fmt}
                  className="flex flex-col gap-2.5 rounded-xl border border-slate-200 p-3.5 sm:flex-row sm:items-center sm:justify-between dark:border-slate-800"
                >
                  <div className="min-w-0">
                    <div className="flex items-center gap-2 text-xs font-bold text-slate-900 dark:text-slate-100">
                      <Icon className="h-4 w-4 text-orange-600 dark:text-orange-400" />
                      {title}
                    </div>
                    <p className="mt-0.5 text-[11px] text-slate-500 dark:text-slate-400">{desc}</p>
                  </div>
                  <div className="flex shrink-0 gap-1.5">
                    <button
                      type="button"
                      onClick={() => {
                        onExportFormat(fmt, 'device');
                        onClose();
                      }}
                      className="rounded-lg bg-orange-600 px-3 py-1.5 text-xs font-semibold text-white hover:bg-orange-700"
                    >
                      Download
                    </button>
                    {fmt !== 'png' && (
                      <button
                        type="button"
                        onClick={() => {
                          onExportFormat(fmt, 'vault');
                          onClose();
                        }}
                        className="rounded-lg border border-slate-300 px-3 py-1.5 text-xs font-semibold text-slate-700 hover:bg-slate-50 dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800"
                      >
                        Save to Vault
                      </button>
                    )}
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
};
