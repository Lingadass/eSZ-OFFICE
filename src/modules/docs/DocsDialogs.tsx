import React, {useEffect, useMemo, useState} from 'react';
import type {Editor} from '@tiptap/react';
import {
  AlertTriangle,
  Check,
  CheckCircle2,
  Clock,
  Copy,
  Download,
  FileDown,
  FolderPlus,
  Info,
  MessageSquare,
  RotateCcw,
  Search,
  ShieldAlert,
  Trash2,
  X,
} from 'lucide-react';
import type {DocVersionRecord} from '../../core/types';
import type {DocxImportWarning} from './docxImport';
import type {DocComment, DocPageLayout, DocStats} from './docsTypes';

export type DocsDialogKind =
  | 'find'
  | 'table'
  | 'link'
  | 'bookmark'
  | 'symbol'
  | 'equation'
  | 'shape'
  | 'header-footer'
  | 'word-count'
  | 'versions'
  | 'export'
  | 'warnings'
  | 'save-lossy';

interface DocsDialogsProps {
  openDialog: DocsDialogKind | null;
  onClose: () => void;
  editor: Editor | null;
  fileName: string;
  readOnly?: boolean;
  layout: DocPageLayout;
  onUpdateLayout: (patch: Partial<DocPageLayout>) => void;
  stats: DocStats;
  actualPageCount?: number;
  versions: DocVersionRecord[];
  onSaveVersion: (label: string) => Promise<void>;
  onRestoreVersion: (version: DocVersionRecord) => void;
  onDeleteVersion: (versionId: string) => Promise<void>;
  onExportFormat: (
    format: '.docx' | '.pdf' | '.html' | '.md' | '.rtf' | '.txt',
    target: 'device' | 'vault',
  ) => Promise<void>;
  importWarnings?: DocxImportWarning[];
  hasLockedElements?: boolean;
  onSaveAsCopy?: () => Promise<void>;
  onOverwriteOriginal?: () => Promise<void>;
}

const SYMBOLS = [
  '©', '®', '™', '§', '¶', '°', '±', '×',
  '÷', '≤', '≥', '≠', '≈', '∞', '∑', '√',
  'π', 'Ω', 'Δ', 'α', 'β', 'μ', '€', '£',
  '¥', '¢', '→', '←', '↑', '↓', '↔', '✓',
  '✗', '•', '—', '–', '…', '«', '»', '‰',
];

const EQUATION_TEMPLATES: Array<{label: string; formula: string}> = [
  {label: 'Quadratic Formula', formula: 'x = (-b ± √(b² - 4ac)) / 2a'},
  {label: 'Pythagorean Theorem', formula: 'a² + b² = c²'},
  {label: 'Mass-Energy', formula: 'E = mc²'},
  {label: 'Summation', formula: '∑(i=1 → n) xᵢ = x₁ + x₂ + … + xₙ'},
  {label: 'Definite Integral', formula: '∫ₐᵇ f(x) dx = F(b) - F(a)'},
  {label: 'Standard Deviation', formula: 'σ = √( (1/N) ∑(xᵢ - μ)² )'},
  {label: 'Euler Identity', formula: 'e^(iπ) + 1 = 0'},
];

const EXPORT_FORMATS: Array<{
  ext: '.docx' | '.pdf' | '.html' | '.md' | '.rtf' | '.txt';
  title: string;
  desc: string;
}> = [
  {ext: '.docx', title: 'Word Document (.docx)', desc: 'OpenXML format compatible with Microsoft Word & LibreOffice'},
  {ext: '.pdf', title: 'PDF Document (.pdf)', desc: 'Fixed-layout printable document with margins, headers & watermark'},
  {ext: '.html', title: 'Web Page (.html)', desc: 'Standalone semantic HTML document with clean styling'},
  {ext: '.md', title: 'Markdown (.md)', desc: 'Portable plain-text Markdown headings, lists & emphasis'},
  {ext: '.rtf', title: 'Rich Text Format (.rtf)', desc: 'Legacy cross-platform rich text format'},
  {ext: '.txt', title: 'Plain Text (.txt)', desc: 'Unformatted UTF-8 text without markup'},
];

function findMatchesInEditor(
  editor: Editor,
  query: string,
  caseSensitive: boolean,
): Array<{from: number; to: number}> {
  if (!query) return [];
  const needle = caseSensitive ? query : query.toLowerCase();
  const matches: Array<{from: number; to: number}> = [];
  editor.state.doc.descendants((node, pos) => {
    if (!node.isText || !node.text) return;
    const hay = caseSensitive ? node.text : node.text.toLowerCase();
    let idx = hay.indexOf(needle);
    while (idx !== -1) {
      matches.push({from: pos + idx, to: pos + idx + query.length});
      idx = hay.indexOf(needle, idx + Math.max(1, query.length));
    }
  });
  return matches;
}

export const DocsDialogs: React.FC<DocsDialogsProps> = ({
  openDialog,
  onClose,
  editor,
  fileName,
  readOnly = false,
  layout,
  onUpdateLayout,
  stats,
  actualPageCount,
  versions,
  onSaveVersion,
  onRestoreVersion,
  onDeleteVersion,
  onExportFormat,
  importWarnings = [],
  hasLockedElements = false,
  onSaveAsCopy,
  onOverwriteOriginal,
}) => {
  // Find & Replace state
  const [findQuery, setFindQuery] = useState('');
  const [replaceText, setReplaceText] = useState('');
  const [caseSensitive, setCaseSensitive] = useState(false);
  const [matchCursor, setMatchCursor] = useState(0);
  const [findMessage, setFindMessage] = useState('');

  // Table state
  const [tableRows, setTableRows] = useState(3);
  const [tableCols, setTableCols] = useState(3);
  const [tableHeader, setTableHeader] = useState(true);

  // Link state
  const [linkUrl, setLinkUrl] = useState('https://');
  const [linkLabel, setLinkLabel] = useState('');

  // Bookmark state
  const [bookmarkName, setBookmarkName] = useState('');

  // Equation state
  const [equationFormula, setEquationFormula] = useState('x = (-b ± √(b² - 4ac)) / 2a');

  // Shape state
  const [shapeType, setShapeType] = useState<'rectangle' | 'rounded' | 'circle' | 'callout'>('rounded');
  const [shapeLabel, setShapeLabel] = useState('Key Callout');
  const [shapeFill, setShapeFill] = useState('#DBEAFE');
  const [shapeBorder, setShapeBorder] = useState('#1D4ED8');

  // Header & Footer draft
  const [headerDraft, setHeaderDraft] = useState(layout.headerText);
  const [footerDraft, setFooterDraft] = useState(layout.footerText);
  const [showNumDraft, setShowNumDraft] = useState(layout.showPageNumbers);
  const [numPosDraft, setNumPosDraft] = useState(layout.pageNumberPosition);
  const [diffFirstDraft, setDiffFirstDraft] = useState(Boolean(layout.differentFirstPage));
  const [firstHeaderDraft, setFirstHeaderDraft] = useState(layout.firstPageHeaderText ?? '');
  const [firstFooterDraft, setFirstFooterDraft] = useState(layout.firstPageFooterText ?? '');

  // Version snapshot label
  const [versionLabel, setVersionLabel] = useState('');
  const [busy, setBusy] = useState(false);

  // Sync dialog drafts when opening
  useEffect(() => {
    if (!openDialog) return;
    if (openDialog === 'header-footer') {
      setHeaderDraft(layout.headerText);
      setFooterDraft(layout.footerText);
      setShowNumDraft(layout.showPageNumbers);
      setNumPosDraft(layout.pageNumberPosition);
      setDiffFirstDraft(Boolean(layout.differentFirstPage));
      setFirstHeaderDraft(layout.firstPageHeaderText ?? '');
      setFirstFooterDraft(layout.firstPageFooterText ?? '');
    } else if (openDialog === 'link' && editor) {
      const existingHref = (editor.getAttributes('link')?.href as string | undefined) || 'https://';
      const {from, to} = editor.state.selection;
      const selected = editor.state.doc.textBetween(from, to, ' ');
      setLinkUrl(existingHref);
      setLinkLabel(selected);
    } else if (openDialog === 'find') {
      setFindMessage('');
      setMatchCursor(0);
    } else if (openDialog === 'versions') {
      setVersionLabel('');
    }
  }, [openDialog, editor, layout]);

  useEffect(() => {
    if (!openDialog) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [openDialog, onClose]);

  const currentMatches = useMemo(() => {
    if (openDialog !== 'find' || !editor || !findQuery) return [];
    return findMatchesInEditor(editor, findQuery, caseSensitive);
  }, [openDialog, editor, findQuery, caseSensitive]);

  if (!openDialog) return null;

  const dialogTitleMap: Record<DocsDialogKind, string> = {
    find: readOnly ? 'Find in Document' : 'Find & Replace',
    table: 'Insert Table',
    link: 'Insert Hyperlink',
    bookmark: 'Insert Bookmark',
    symbol: 'Insert Special Symbol',
    equation: 'Insert Equation',
    shape: 'Insert Shape',
    'header-footer': 'Header, Footer & Page Numbers',
    'word-count': 'Word Count & Document Statistics',
    versions: 'Version History',
    export: 'Save As & Export Document',
    warnings: 'Word Compatibility & Import Details',
    'save-lossy': 'Save Edited Word Document',
  };

  const handleFindNext = () => {
    if (!editor || !findQuery) return;
    const matches = findMatchesInEditor(editor, findQuery, caseSensitive);
    if (matches.length === 0) {
      setFindMessage('No matches found.');
      return;
    }
    const idx = matchCursor % matches.length;
    const target = matches[idx];
    editor.chain().focus().setTextSelection(target).scrollIntoView().run();
    setMatchCursor((idx + 1) % matches.length);
    setFindMessage(`Match ${idx + 1} of ${matches.length}`);
  };

  const handleReplaceOne = () => {
    if (!editor || !findQuery) return;
    const matches = findMatchesInEditor(editor, findQuery, caseSensitive);
    if (matches.length === 0) {
      setFindMessage('No matches left to replace.');
      return;
    }
    const {from, to} = editor.state.selection;
    const active = matches.find((m) => m.from === from && m.to === to) ?? matches[0];
    editor
      .chain()
      .focus()
      .insertContentAt({from: active.from, to: active.to}, replaceText)
      .run();
    const remaining = findMatchesInEditor(editor, findQuery, caseSensitive).length;
    setFindMessage(`Replaced 1 occurrence (${remaining} remaining).`);
  };

  const handleReplaceAll = () => {
    if (!editor || !findQuery) return;
    const matches = findMatchesInEditor(editor, findQuery, caseSensitive);
    if (matches.length === 0) {
      setFindMessage('No matches found.');
      return;
    }
    const tr = editor.state.tr;
    for (let i = matches.length - 1; i >= 0; i--) {
      tr.insertText(replaceText, matches[i].from, matches[i].to);
    }
    editor.view.dispatch(tr);
    setFindMessage(`Replaced all ${matches.length} occurrence${matches.length === 1 ? '' : 's'}.`);
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
        aria-labelledby="docs-dialog-title"
        className="flex max-h-[88vh] w-full max-w-lg flex-col overflow-hidden rounded-xl border border-slate-200 bg-white shadow-2xl dark:border-slate-700 dark:bg-slate-900"
      >
        <div className="flex items-center justify-between border-b border-slate-200 px-4 py-3 dark:border-slate-800">
          <h2 id="docs-dialog-title" className="text-sm font-semibold text-slate-900 dark:text-slate-100">
            {dialogTitleMap[openDialog]}
          </h2>
          <button
            type="button"
            onClick={onClose}
            aria-label="Close dialog"
            className="rounded-md p-1 text-slate-500 hover:bg-slate-100 hover:text-slate-900 dark:text-slate-400 dark:hover:bg-slate-800 dark:hover:text-slate-100"
          >
            <X className="h-4 w-4" />
          </button>
        </div>

        <div className="overflow-y-auto p-4 text-sm">
          {openDialog === 'find' && (
            <div className="space-y-3">
              <div>
                <label className="mb-1 block text-xs font-medium text-slate-600 dark:text-slate-300">
                  Find text
                </label>
                <div className="relative">
                  <Search className="pointer-events-none absolute top-2.5 left-2.5 h-4 w-4 text-slate-400" />
                  <input
                    type="text"
                    value={findQuery}
                    onChange={(e) => {
                      setFindQuery(e.target.value);
                      setMatchCursor(0);
                    }}
                    placeholder="Enter text to search..."
                    autoFocus
                    className="w-full rounded-lg border border-slate-300 bg-white py-2 pr-3 pl-8 text-sm text-slate-900 focus:border-blue-600 focus:outline-none dark:border-slate-700 dark:bg-slate-950 dark:text-slate-100"
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
                    placeholder="Replacement text..."
                    className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm text-slate-900 focus:border-blue-600 focus:outline-none dark:border-slate-700 dark:bg-slate-950 dark:text-slate-100"
                  />
                </div>
              )}

              <div className="flex items-center justify-between">
                <label className="flex items-center gap-2 text-xs text-slate-700 dark:text-slate-300">
                  <input
                    type="checkbox"
                    checked={caseSensitive}
                    onChange={(e) => setCaseSensitive(e.target.checked)}
                    className="h-3.5 w-3.5 accent-blue-600"
                  />
                  Match case
                </label>
                <span className="text-xs font-medium tabular-nums text-slate-500 dark:text-slate-400">
                  {findQuery ? `${currentMatches.length} found` : 'Type to search'}
                </span>
              </div>

              {findMessage && (
                <p className="rounded-md bg-slate-100 px-3 py-1.5 text-xs text-slate-700 dark:bg-slate-800 dark:text-slate-200">
                  {findMessage}
                </p>
              )}

              <div className="flex flex-wrap justify-end gap-2 pt-2">
                <button
                  type="button"
                  onClick={handleFindNext}
                  disabled={!findQuery}
                  className="rounded-lg border border-slate-300 px-3 py-1.5 text-xs font-semibold text-slate-700 hover:bg-slate-100 disabled:opacity-40 dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800"
                >
                  Find Next
                </button>
                {!readOnly && (
                  <>
                    <button
                      type="button"
                      onClick={handleReplaceOne}
                      disabled={!findQuery}
                      className="rounded-lg border border-slate-300 px-3 py-1.5 text-xs font-semibold text-slate-700 hover:bg-slate-100 disabled:opacity-40 dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800"
                    >
                      Replace
                    </button>
                    <button
                      type="button"
                      onClick={handleReplaceAll}
                      disabled={!findQuery}
                      className="rounded-lg bg-blue-700 px-3.5 py-1.5 text-xs font-semibold text-white hover:bg-blue-800 disabled:opacity-40 dark:bg-blue-600"
                    >
                      Replace All
                    </button>
                  </>
                )}
              </div>
            </div>
          )}

          {openDialog === 'table' && (
            <form
              className="space-y-4"
              onSubmit={(e) => {
                e.preventDefault();
                editor
                  ?.chain()
                  .focus()
                  .insertTable({
                    rows: Math.max(1, Math.min(20, tableRows)),
                    cols: Math.max(1, Math.min(10, tableCols)),
                    withHeaderRow: tableHeader,
                  })
                  .run();
                onClose();
              }}
            >
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="mb-1 block text-xs font-medium text-slate-600 dark:text-slate-300">
                    Rows (1–20)
                  </label>
                  <input
                    type="number"
                    min={1}
                    max={20}
                    value={tableRows}
                    onChange={(e) => setTableRows(Number(e.target.value))}
                    className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm tabular-nums dark:border-slate-700 dark:bg-slate-950"
                  />
                </div>
                <div>
                  <label className="mb-1 block text-xs font-medium text-slate-600 dark:text-slate-300">
                    Columns (1–10)
                  </label>
                  <input
                    type="number"
                    min={1}
                    max={10}
                    value={tableCols}
                    onChange={(e) => setTableCols(Number(e.target.value))}
                    className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm tabular-nums dark:border-slate-700 dark:bg-slate-950"
                  />
                </div>
              </div>

              <label className="flex items-center gap-2 text-xs text-slate-700 dark:text-slate-300">
                <input
                  type="checkbox"
                  checked={tableHeader}
                  onChange={(e) => setTableHeader(e.target.checked)}
                  className="h-3.5 w-3.5 accent-blue-600"
                />
                Include styled header row
              </label>

              <div className="flex justify-end gap-2 pt-2">
                <button
                  type="button"
                  onClick={onClose}
                  className="rounded-lg border border-slate-300 px-3 py-1.5 text-xs font-semibold text-slate-700 hover:bg-slate-100 dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="rounded-lg bg-blue-700 px-4 py-1.5 text-xs font-semibold text-white hover:bg-blue-800 dark:bg-blue-600"
                >
                  Insert Table
                </button>
              </div>
            </form>
          )}

          {openDialog === 'link' && (
            <form
              className="space-y-3"
              onSubmit={(e) => {
                e.preventDefault();
                if (!editor) return;
                const cleanUrl = linkUrl.trim();
                if (!cleanUrl) return;
                const {empty} = editor.state.selection;
                if (empty && linkLabel.trim()) {
                  editor
                    .chain()
                    .focus()
                    .insertContent(`<a href="${cleanUrl.replace(/"/g, '&quot;')}">${linkLabel.trim()}</a>`)
                    .run();
                } else {
                  editor
                    .chain()
                    .focus()
                    .extendMarkRange('link')
                    .setLink({href: cleanUrl})
                    .run();
                }
                onClose();
              }}
            >
              <div>
                <label className="mb-1 block text-xs font-medium text-slate-600 dark:text-slate-300">
                  Display text (when no text is selected)
                </label>
                <input
                  type="text"
                  value={linkLabel}
                  onChange={(e) => setLinkLabel(e.target.value)}
                  placeholder="Reference link"
                  className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm dark:border-slate-700 dark:bg-slate-950"
                />
              </div>
              <div>
                <label className="mb-1 block text-xs font-medium text-slate-600 dark:text-slate-300">
                  Target URL
                </label>
                <input
                  type="url"
                  required
                  value={linkUrl}
                  onChange={(e) => setLinkUrl(e.target.value)}
                  placeholder="https://example.org"
                  className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm dark:border-slate-700 dark:bg-slate-950"
                />
              </div>
              <div className="flex items-center justify-between pt-2">
                {editor?.isActive('link') ? (
                  <button
                    type="button"
                    onClick={() => {
                      editor.chain().focus().unsetLink().run();
                      onClose();
                    }}
                    className="rounded-lg border border-rose-200 px-3 py-1.5 text-xs font-semibold text-rose-700 hover:bg-rose-50 dark:border-rose-900 dark:text-rose-400"
                  >
                    Remove Link
                  </button>
                ) : (
                  <span />
                )}
                <div className="flex gap-2">
                  <button
                    type="button"
                    onClick={onClose}
                    className="rounded-lg border border-slate-300 px-3 py-1.5 text-xs font-semibold text-slate-700 dark:border-slate-700 dark:text-slate-200"
                  >
                    Cancel
                  </button>
                  <button
                    type="submit"
                    className="rounded-lg bg-blue-700 px-4 py-1.5 text-xs font-semibold text-white hover:bg-blue-800 dark:bg-blue-600"
                  >
                    Apply Link
                  </button>
                </div>
              </div>
            </form>
          )}

          {openDialog === 'bookmark' && (
            <form
              className="space-y-3"
              onSubmit={(e) => {
                e.preventDefault();
                if (!editor || !bookmarkName.trim()) return;
                const clean = bookmarkName.trim().replace(/\s+/g, '-');
                if (editor.state.selection.empty) {
                  editor
                    .chain()
                    .focus()
                    .insertContent(`<span data-bookmark="${clean}" id="bm-${clean}">${bookmarkName.trim()}</span> `)
                    .run();
                } else {
                  editor.chain().focus().setBookmark(clean).run();
                }
                setBookmarkName('');
                onClose();
              }}
            >
              <p className="text-xs text-slate-600 dark:text-slate-400">
                Bookmarks mark a named anchor on the selected text (or at the cursor) so you can reference it across the document.
              </p>
              <div>
                <label className="mb-1 block text-xs font-medium text-slate-600 dark:text-slate-300">
                  Bookmark name
                </label>
                <input
                  type="text"
                  required
                  autoFocus
                  value={bookmarkName}
                  onChange={(e) => setBookmarkName(e.target.value)}
                  placeholder="executive-summary"
                  className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm dark:border-slate-700 dark:bg-slate-950"
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
                  className="rounded-lg bg-blue-700 px-4 py-1.5 text-xs font-semibold text-white hover:bg-blue-800 dark:bg-blue-600"
                >
                  Insert Bookmark
                </button>
              </div>
            </form>
          )}

          {openDialog === 'symbol' && (
            <div className="space-y-3">
              <p className="text-xs text-slate-600 dark:text-slate-400">
                Click any symbol to insert it at the current cursor position.
              </p>
              <div className="grid grid-cols-8 gap-1.5">
                {SYMBOLS.map((sym) => (
                  <button
                    key={sym}
                    type="button"
                    onClick={() => {
                      editor?.chain().focus().insertContent(sym).run();
                      onClose();
                    }}
                    className="flex h-10 items-center justify-center rounded-lg border border-slate-200 bg-slate-50 text-base font-medium text-slate-800 hover:border-blue-500 hover:bg-blue-50 dark:border-slate-700 dark:bg-slate-950 dark:text-slate-100 dark:hover:bg-blue-950/60"
                  >
                    {sym}
                  </button>
                ))}
              </div>
            </div>
          )}

          {openDialog === 'equation' && (
            <form
              className="space-y-3"
              onSubmit={(e) => {
                e.preventDefault();
                if (!equationFormula.trim()) return;
                editor?.chain().focus().insertDocEquation(equationFormula.trim()).run();
                onClose();
              }}
            >
              <div>
                <label className="mb-1 block text-xs font-medium text-slate-600 dark:text-slate-300">
                  Equation expression
                </label>
                <input
                  type="text"
                  required
                  value={equationFormula}
                  onChange={(e) => setEquationFormula(e.target.value)}
                  className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 font-serif text-base dark:border-slate-700 dark:bg-slate-950"
                />
              </div>

              <div>
                <p className="mb-1.5 text-xs font-medium text-slate-600 dark:text-slate-300">
                  Quick templates
                </p>
                <div className="grid grid-cols-1 gap-1.5 sm:grid-cols-2">
                  {EQUATION_TEMPLATES.map((tpl) => (
                    <button
                      key={tpl.label}
                      type="button"
                      onClick={() => setEquationFormula(tpl.formula)}
                      className="rounded-lg border border-slate-200 px-2.5 py-2 text-left hover:border-blue-500 hover:bg-blue-50/50 dark:border-slate-700 dark:hover:bg-slate-800"
                    >
                      <div className="text-[11px] font-semibold text-slate-700 dark:text-slate-200">
                        {tpl.label}
                      </div>
                      <div className="truncate font-serif text-xs text-slate-500 dark:text-slate-400">
                        {tpl.formula}
                      </div>
                    </button>
                  ))}
                </div>
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
                  className="rounded-lg bg-blue-700 px-4 py-1.5 text-xs font-semibold text-white hover:bg-blue-800 dark:bg-blue-600"
                >
                  Insert Equation
                </button>
              </div>
            </form>
          )}

          {openDialog === 'shape' && (
            <form
              className="space-y-3"
              onSubmit={(e) => {
                e.preventDefault();
                editor
                  ?.chain()
                  .focus()
                  .insertDocShape({
                    shapeType,
                    label: shapeLabel.trim() || 'Shape',
                    fill: shapeFill,
                    border: shapeBorder,
                  })
                  .run();
                onClose();
              }}
            >
              <div className="grid grid-cols-2 gap-2 sm:grid-cols-4">
                {(
                  [
                    ['rectangle', 'Rectangle'],
                    ['rounded', 'Rounded'],
                    ['circle', 'Circle'],
                    ['callout', 'Callout'],
                  ] as const
                ).map(([type, label]) => (
                  <button
                    key={type}
                    type="button"
                    onClick={() => setShapeType(type)}
                    className={`rounded-lg border px-2.5 py-2 text-xs font-semibold ${
                      shapeType === type
                        ? 'border-blue-600 bg-blue-50 text-blue-800 dark:bg-blue-950/60 dark:text-blue-300'
                        : 'border-slate-200 text-slate-700 hover:bg-slate-50 dark:border-slate-700 dark:text-slate-300'
                    }`}
                  >
                    {label}
                  </button>
                ))}
              </div>

              <div>
                <label className="mb-1 block text-xs font-medium text-slate-600 dark:text-slate-300">
                  Shape text
                </label>
                <input
                  type="text"
                  value={shapeLabel}
                  onChange={(e) => setShapeLabel(e.target.value)}
                  className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm dark:border-slate-700 dark:bg-slate-950"
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <label className="flex items-center justify-between rounded-lg border border-slate-200 px-3 py-2 text-xs dark:border-slate-700">
                  <span>Fill color</span>
                  <input
                    type="color"
                    value={shapeFill}
                    onChange={(e) => setShapeFill(e.target.value)}
                    className="h-5 w-7 cursor-pointer border-0 bg-transparent"
                  />
                </label>
                <label className="flex items-center justify-between rounded-lg border border-slate-200 px-3 py-2 text-xs dark:border-slate-700">
                  <span>Border color</span>
                  <input
                    type="color"
                    value={shapeBorder}
                    onChange={(e) => setShapeBorder(e.target.value)}
                    className="h-5 w-7 cursor-pointer border-0 bg-transparent"
                  />
                </label>
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
                  className="rounded-lg bg-blue-700 px-4 py-1.5 text-xs font-semibold text-white hover:bg-blue-800 dark:bg-blue-600"
                >
                  Insert Shape
                </button>
              </div>
            </form>
          )}

          {openDialog === 'header-footer' && (
            <form
              className="space-y-3"
              onSubmit={(e) => {
                e.preventDefault();
                onUpdateLayout({
                  headerText: headerDraft,
                  footerText: footerDraft,
                  showPageNumbers: showNumDraft,
                  pageNumberPosition: numPosDraft,
                  differentFirstPage: diffFirstDraft,
                  firstPageHeaderText: firstHeaderDraft,
                  firstPageFooterText: firstFooterDraft,
                });
                onClose();
              }}
            >
              <div>
                <label className="mb-1 block text-xs font-medium text-slate-600 dark:text-slate-300">
                  Header text
                </label>
                <input
                  type="text"
                  value={headerDraft}
                  onChange={(e) => setHeaderDraft(e.target.value)}
                  placeholder="e.g., Q4 Executive Report — Internal"
                  className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm dark:border-slate-700 dark:bg-slate-950"
                />
              </div>
              <div>
                <label className="mb-1 block text-xs font-medium text-slate-600 dark:text-slate-300">
                  Footer text
                </label>
                <input
                  type="text"
                  value={footerDraft}
                  onChange={(e) => setFooterDraft(e.target.value)}
                  placeholder="e.g., Prepared with eSZ OFFICE"
                  className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm dark:border-slate-700 dark:bg-slate-950"
                />
              </div>

              <div className="flex flex-wrap items-center justify-between gap-2 rounded-lg border border-slate-200 px-3 py-2.5 dark:border-slate-700">
                <label className="flex items-center gap-2 text-xs font-medium text-slate-700 dark:text-slate-200">
                  <input
                    type="checkbox"
                    checked={showNumDraft}
                    onChange={(e) => setShowNumDraft(e.target.checked)}
                    className="h-3.5 w-3.5 accent-blue-600"
                  />
                  Show page numbers
                </label>
                <select
                  aria-label="Page number position"
                  value={numPosDraft}
                  onChange={(e) =>
                    setNumPosDraft(e.target.value as DocPageLayout['pageNumberPosition'])
                  }
                  disabled={!showNumDraft}
                  className="rounded border border-slate-300 bg-white px-2 py-1 text-xs disabled:opacity-40 dark:border-slate-700 dark:bg-slate-950"
                >
                  <option value="bottom-center">Bottom Center</option>
                  <option value="bottom-right">Bottom Right</option>
                  <option value="top-right">Top Right</option>
                </select>
              </div>

              <div className="space-y-2.5 rounded-lg border border-slate-200 px-3 py-2.5 dark:border-slate-700">
                <label className="flex items-center gap-2 text-xs font-medium text-slate-700 dark:text-slate-200">
                  <input
                    type="checkbox"
                    checked={diffFirstDraft}
                    onChange={(e) => setDiffFirstDraft(e.target.checked)}
                    className="h-3.5 w-3.5 accent-blue-600"
                  />
                  Different first page (hide page number &amp; use separate Page 1 header/footer)
                </label>
                {diffFirstDraft && (
                  <div className="grid grid-cols-1 gap-2 pt-1 sm:grid-cols-2">
                    <div>
                      <label className="mb-1 block text-[11px] font-medium text-slate-500 dark:text-slate-400">
                        Page 1 header (optional)
                      </label>
                      <input
                        type="text"
                        value={firstHeaderDraft}
                        onChange={(e) => setFirstHeaderDraft(e.target.value)}
                        placeholder="Leave blank for none"
                        className="w-full rounded-md border border-slate-300 bg-white px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-950"
                      />
                    </div>
                    <div>
                      <label className="mb-1 block text-[11px] font-medium text-slate-500 dark:text-slate-400">
                        Page 1 footer (optional)
                      </label>
                      <input
                        type="text"
                        value={firstFooterDraft}
                        onChange={(e) => setFirstFooterDraft(e.target.value)}
                        placeholder="Leave blank for none"
                        className="w-full rounded-md border border-slate-300 bg-white px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-950"
                      />
                    </div>
                  </div>
                )}
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
                  className="rounded-lg bg-blue-700 px-4 py-1.5 text-xs font-semibold text-white hover:bg-blue-800 dark:bg-blue-600"
                >
                  Save Header &amp; Footer
                </button>
              </div>
            </form>
          )}

          {openDialog === 'word-count' && (
            <div className="space-y-3">
              <div className="divide-y divide-slate-200 rounded-lg border border-slate-200 dark:divide-slate-800 dark:border-slate-800">
                {(
                  [
                    ['Words', stats.words],
                    ['Characters (with spaces)', stats.characters],
                    ['Characters (no spaces)', stats.charactersNoSpaces],
                    ['Paragraphs', stats.paragraphs],
                    ['Pages', actualPageCount ?? stats.estimatedPages],
                    ['Reading time (min)', stats.readingMinutes],
                  ] as const
                ).map(([label, val]) => (
                  <div key={label} className="flex items-center justify-between px-3.5 py-2.5 text-xs">
                    <span className="text-slate-600 dark:text-slate-400">{label}</span>
                    <span className="font-semibold tabular-nums text-slate-900 dark:text-slate-100">
                      {val.toLocaleString('en-US')}
                    </span>
                  </div>
                ))}
              </div>
              <div className="flex justify-end">
                <button
                  type="button"
                  onClick={onClose}
                  className="rounded-lg bg-blue-700 px-4 py-1.5 text-xs font-semibold text-white hover:bg-blue-800 dark:bg-blue-600"
                >
                  Done
                </button>
              </div>
            </div>
          )}

          {openDialog === 'versions' && (
            <div className="space-y-4">
              <form
                className="flex gap-2"
                onSubmit={async (e) => {
                  e.preventDefault();
                  setBusy(true);
                  try {
                    await onSaveVersion(
                      versionLabel.trim() || `Snapshot ${new Date().toLocaleTimeString('en-US')}`,
                    );
                    setVersionLabel('');
                  } finally {
                    setBusy(false);
                  }
                }}
              >
                <input
                  type="text"
                  value={versionLabel}
                  onChange={(e) => setVersionLabel(e.target.value)}
                  placeholder="Optional snapshot label (e.g., Draft 2)"
                  className="min-w-0 flex-1 rounded-lg border border-slate-300 bg-white px-3 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-950"
                />
                <button
                  type="submit"
                  disabled={busy}
                  className="shrink-0 whitespace-nowrap rounded-lg bg-blue-700 px-3 py-1.5 text-xs font-semibold text-white hover:bg-blue-800 disabled:opacity-50 dark:bg-blue-600"
                >
                  Save Snapshot
                </button>
              </form>

              {versions.length === 0 ? (
                <div className="rounded-lg border border-dashed border-slate-300 p-6 text-center text-xs text-slate-500 dark:border-slate-700 dark:text-slate-400">
                  No snapshots saved yet. Create a snapshot above to capture this revision locally.
                </div>
              ) : (
                <div className="max-h-72 space-y-2 overflow-y-auto pr-1">
                  {versions.map((v) => (
                    <div
                      key={v.id}
                      className="flex items-center justify-between gap-2 rounded-lg border border-slate-200 bg-slate-50 p-2.5 dark:border-slate-800 dark:bg-slate-950"
                    >
                      <div className="min-w-0">
                        <div className="truncate text-xs font-semibold text-slate-800 dark:text-slate-100">
                          {v.label}
                        </div>
                        <div className="mt-0.5 flex items-center gap-2 text-[11px] tabular-nums text-slate-500 dark:text-slate-400">
                          <span className="inline-flex items-center gap-1">
                            <Clock className="h-3 w-3" />
                            {new Date(v.savedAt).toLocaleString('en-US', {
                              month: 'short',
                              day: 'numeric',
                              hour: '2-digit',
                              minute: '2-digit',
                            })}
                          </span>
                          <span>{v.wordCount} words</span>
                        </div>
                      </div>
                      <div className="flex shrink-0 items-center gap-1">
                        <button
                          type="button"
                          onClick={() => {
                            onRestoreVersion(v);
                            onClose();
                          }}
                          className="inline-flex items-center gap-1 rounded-md border border-slate-300 bg-white px-2.5 py-1 text-xs font-semibold text-slate-700 hover:bg-blue-50 hover:text-blue-700 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-200"
                        >
                          <RotateCcw className="h-3 w-3" />
                          Restore
                        </button>
                        <button
                          type="button"
                          aria-label={`Delete snapshot ${v.label}`}
                          onClick={() => void onDeleteVersion(v.id)}
                          className="rounded-md p-1.5 text-slate-400 hover:bg-rose-50 hover:text-rose-600 dark:hover:bg-rose-950/50"
                        >
                          <Trash2 className="h-3.5 w-3.5" />
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}

          {openDialog === 'export' && (
            <div className="space-y-2.5">
              <p className="text-xs text-slate-600 dark:text-slate-400">
                Export <span className="font-semibold text-slate-800 dark:text-slate-200">{fileName}</span> completely offline. Choose to download to your device or save a copy in your local vault.
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
                        className="inline-flex items-center gap-1 whitespace-nowrap rounded-md bg-blue-700 px-2.5 py-1.5 text-xs font-semibold text-white hover:bg-blue-800 disabled:opacity-40 dark:bg-blue-600"
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

          {openDialog === 'warnings' && (
            <div className="space-y-3" data-testid="docs-warnings-dialog">
              <p className="text-xs text-slate-600 dark:text-slate-300">
                {hasLockedElements
                  ? 'Some advanced Word elements in this file cannot be edited directly and are preserved as locked visual blocks in Edit mode. View mode always renders the untouched original .docx.'
                  : 'Compatibility notes detected while importing this Word document for editing. View mode always preserves the untouched original .docx.'}
              </p>

              {importWarnings.length === 0 ? (
                <div className="rounded-lg border border-slate-200 bg-slate-50 p-4 text-center text-xs text-slate-600 dark:border-slate-800 dark:bg-slate-950 dark:text-slate-400">
                  No compatibility issues were detected for this document.
                </div>
              ) : (
                <div className="max-h-72 space-y-2 overflow-y-auto pr-1">
                  {importWarnings.map((w) => (
                    <div
                      key={w.code}
                      className={`rounded-lg border p-3 text-xs ${
                        w.severity === 'warning'
                          ? 'border-amber-200 bg-amber-50/70 dark:border-amber-900/60 dark:bg-amber-950/30'
                          : 'border-slate-200 bg-slate-50 dark:border-slate-800 dark:bg-slate-950'
                      }`}
                    >
                      <div className="flex items-center justify-between gap-2">
                        <span className="inline-flex items-center gap-1.5 font-semibold text-slate-900 dark:text-slate-100">
                          {w.severity === 'warning' ? (
                            <AlertTriangle className="h-3.5 w-3.5 shrink-0 text-amber-600 dark:text-amber-400" />
                          ) : (
                            <Info className="h-3.5 w-3.5 shrink-0 text-blue-600 dark:text-blue-400" />
                          )}
                          {w.feature}
                        </span>
                        <span className="shrink-0 rounded bg-white/80 px-1.5 py-0.5 text-[10px] font-semibold tabular-nums text-slate-600 ring-1 ring-slate-200 dark:bg-slate-900 dark:text-slate-300 dark:ring-slate-700">
                          {w.count}×
                        </span>
                      </div>
                      <p className="mt-1 text-[11px] leading-relaxed text-slate-600 dark:text-slate-300">
                        {w.message}
                      </p>
                    </div>
                  ))}
                </div>
              )}

              <div className="flex justify-end pt-1">
                <button
                  type="button"
                  onClick={onClose}
                  className="rounded-lg bg-blue-700 px-4 py-1.5 text-xs font-semibold text-white hover:bg-blue-800 dark:bg-blue-600"
                >
                  Close
                </button>
              </div>
            </div>
          )}

          {openDialog === 'save-lossy' && (
            <div className="space-y-3.5" data-testid="docs-save-lossy-dialog">
              <div className="flex items-start gap-2.5 rounded-lg border border-amber-200 bg-amber-50/80 p-3 text-xs text-amber-950 dark:border-amber-900/60 dark:bg-amber-950/30 dark:text-amber-100">
                <ShieldAlert className="mt-0.5 h-4 w-4 shrink-0 text-amber-600 dark:text-amber-400" />
                <div className="space-y-1">
                  <p className="font-semibold">
                    Original file contains advanced Word features
                  </p>
                  <p className="text-[11px] leading-relaxed text-amber-900/90 dark:text-amber-200/90">
                    Saving as a new copy keeps <span className="font-semibold">{fileName}</span> completely untouched while writing your edits to a separate .docx file in your vault.
                  </p>
                </div>
              </div>

              {importWarnings.length > 0 && (
                <div className="max-h-36 space-y-1 overflow-y-auto rounded-lg border border-slate-200 bg-slate-50 p-2.5 text-[11px] text-slate-700 dark:border-slate-800 dark:bg-slate-950 dark:text-slate-300">
                  {importWarnings.map((w) => (
                    <div key={w.code} className="flex items-center justify-between gap-2">
                      <span className="truncate font-medium">{w.feature}</span>
                      <span className="shrink-0 tabular-nums text-slate-500">{w.count}×</span>
                    </div>
                  ))}
                </div>
              )}

              <div className="flex flex-wrap items-center justify-end gap-2 pt-1">
                <button
                  type="button"
                  disabled={busy}
                  onClick={onClose}
                  className="rounded-lg border border-slate-300 px-3 py-1.5 text-xs font-semibold text-slate-700 hover:bg-slate-100 disabled:opacity-40 dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800"
                >
                  Cancel
                </button>
                {onOverwriteOriginal && (
                  <button
                    type="button"
                    disabled={busy}
                    onClick={async () => {
                      setBusy(true);
                      try {
                        await onOverwriteOriginal();
                        onClose();
                      } finally {
                        setBusy(false);
                      }
                    }}
                    className="rounded-lg border border-amber-300 bg-amber-50 px-3 py-1.5 text-xs font-semibold text-amber-900 hover:bg-amber-100 disabled:opacity-40 dark:border-amber-800 dark:bg-amber-950/60 dark:text-amber-200"
                  >
                    Overwrite Original
                  </button>
                )}
                {onSaveAsCopy && (
                  <button
                    type="button"
                    disabled={busy}
                    onClick={async () => {
                      setBusy(true);
                      try {
                        await onSaveAsCopy();
                        onClose();
                      } finally {
                        setBusy(false);
                      }
                    }}
                    className="inline-flex items-center gap-1.5 rounded-lg bg-blue-700 px-3.5 py-1.5 text-xs font-semibold text-white hover:bg-blue-800 disabled:opacity-40 dark:bg-blue-600"
                  >
                    <Copy className="h-3.5 w-3.5" />
                    Save as Copy
                  </button>
                )}
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

interface DocsCommentsSidebarProps {
  comments: DocComment[];
  draftQuote: string;
  composerOpen: boolean;
  onClosePanel: () => void;
  onCancelComposer: () => void;
  onSubmitComment: (text: string) => void;
  onToggleResolve: (id: string) => void;
  onDeleteComment: (id: string) => void;
}

export const DocsCommentsSidebar: React.FC<DocsCommentsSidebarProps> = ({
  comments,
  draftQuote,
  composerOpen,
  onClosePanel,
  onCancelComposer,
  onSubmitComment,
  onToggleResolve,
  onDeleteComment,
}) => {
  const [commentText, setCommentText] = useState('');

  return (
    <aside
      aria-label="Document comments"
      className="flex w-full shrink-0 flex-col border-l border-slate-200 bg-white sm:w-80 dark:border-slate-800 dark:bg-slate-900"
    >
      <div className="flex items-center justify-between border-b border-slate-200 px-3.5 py-2.5 dark:border-slate-800">
        <div className="flex items-center gap-1.5 text-xs font-semibold text-slate-900 dark:text-slate-100">
          <MessageSquare className="h-3.5 w-3.5 text-blue-600 dark:text-blue-400" />
          Comments ({comments.length})
        </div>
        <button
          type="button"
          onClick={onClosePanel}
          aria-label="Close comments panel"
          className="rounded p-1 text-slate-500 hover:bg-slate-100 dark:hover:bg-slate-800"
        >
          <X className="h-4 w-4" />
        </button>
      </div>

      <div className="flex-1 space-y-3 overflow-y-auto p-3">
        {composerOpen && (
          <form
            className="space-y-2 rounded-lg border border-blue-200 bg-blue-50/50 p-2.5 dark:border-blue-900 dark:bg-blue-950/30"
            onSubmit={(e) => {
              e.preventDefault();
              if (!commentText.trim()) return;
              onSubmitComment(commentText.trim());
              setCommentText('');
            }}
          >
            {draftQuote && (
              <div className="truncate border-l-2 border-blue-500 pl-2 text-[11px] italic text-slate-600 dark:text-slate-300">
                &ldquo;{draftQuote}&rdquo;
              </div>
            )}
            <textarea
              rows={3}
              autoFocus
              required
              value={commentText}
              onChange={(e) => setCommentText(e.target.value)}
              placeholder="Write a review comment..."
              className="w-full resize-none rounded-md border border-slate-300 bg-white p-2 text-xs text-slate-900 focus:border-blue-600 focus:outline-none dark:border-slate-700 dark:bg-slate-950 dark:text-slate-100"
            />
            <div className="flex justify-end gap-1.5">
              <button
                type="button"
                onClick={() => {
                  setCommentText('');
                  onCancelComposer();
                }}
                className="rounded-md border border-slate-300 px-2.5 py-1 text-xs font-medium text-slate-600 dark:border-slate-700 dark:text-slate-300"
              >
                Cancel
              </button>
              <button
                type="submit"
                className="rounded-md bg-blue-700 px-2.5 py-1 text-xs font-semibold text-white hover:bg-blue-800 dark:bg-blue-600"
              >
                Post Comment
              </button>
            </div>
          </form>
        )}

        {comments.length === 0 && !composerOpen ? (
          <div className="rounded-lg border border-dashed border-slate-200 p-5 text-center text-xs text-slate-500 dark:border-slate-800 dark:text-slate-400">
            No comments yet. Select text and click <strong>New Comment</strong> in the Review tab.
          </div>
        ) : (
          comments.map((c) => (
            <div
              key={c.id}
              className={`rounded-lg border p-2.5 text-xs ${
                c.resolved
                  ? 'border-slate-200 bg-slate-50/70 opacity-70 dark:border-slate-800 dark:bg-slate-950/60'
                  : 'border-slate-200 bg-white dark:border-slate-800 dark:bg-slate-950'
              }`}
            >
              <div className="flex items-center justify-between gap-2">
                <span className="font-semibold text-slate-800 dark:text-slate-200">
                  {c.author}
                </span>
                <span className="text-[10px] tabular-nums text-slate-400">
                  {new Date(c.createdAt).toLocaleTimeString('en-US', {
                    hour: '2-digit',
                    minute: '2-digit',
                  })}
                </span>
              </div>
              {c.quote && (
                <div className="mt-1 truncate border-l-2 border-amber-400 pl-2 text-[11px] italic text-slate-500 dark:text-slate-400">
                  &ldquo;{c.quote}&rdquo;
                </div>
              )}
              <p className="mt-1.5 whitespace-pre-wrap text-slate-700 dark:text-slate-300">
                {c.text}
              </p>
              <div className="mt-2 flex items-center justify-between border-t border-slate-100 pt-1.5 dark:border-slate-800">
                <button
                  type="button"
                  onClick={() => onToggleResolve(c.id)}
                  className="inline-flex items-center gap-1 text-[11px] font-medium text-blue-700 hover:underline dark:text-blue-400"
                >
                  {c.resolved ? (
                    <>
                      <CheckCircle2 className="h-3 w-3" />
                      Reopen
                    </>
                  ) : (
                    <>
                      <Check className="h-3 w-3" />
                      Resolve
                    </>
                  )}
                </button>
                <button
                  type="button"
                  aria-label={`Delete comment ${c.id}`}
                  onClick={() => onDeleteComment(c.id)}
                  className="rounded p-1 text-slate-400 hover:text-rose-600"
                >
                  <Trash2 className="h-3 w-3" />
                </button>
              </div>
            </div>
          ))
        )}
      </div>
    </aside>
  );
};
