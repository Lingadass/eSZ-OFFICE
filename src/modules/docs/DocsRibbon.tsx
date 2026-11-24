import React from 'react';
import type {Editor} from '@tiptap/react';
import {
  AlignCenter,
  AlignJustify,
  AlignLeft,
  AlignRight,
  Bold,
  BookMarked,
  Bookmark,
  BookOpen,
  Check,
  ChevronLeft,
  ChevronRight,
  Columns3,
  Download,
  Eraser,
  FileDown,
  FileText,
  Globe,
  Hash,
  Highlighter,
  History,
  ImagePlus,
  IndentDecrease,
  IndentIncrease,
  Italic,
  Link2,
  List,
  ListOrdered,
  ListTree,
  MessageSquarePlus,
  Omega,
  Paintbrush,
  PanelTop,
  Pilcrow,
  Printer,
  Redo2,
  Ruler,
  Save,
  Scissors,
  Search,
  Shapes,
  Sigma,
  SpellCheck2,
  SquareDashedBottom,
  Star,
  Strikethrough,
  Subscript,
  Superscript,
  Table2,
  Type,
  Underline,
  Undo2,
  ZoomIn,
  ZoomOut,
} from 'lucide-react';
import {canonicalizeWordFontFamily, DOCUMENT_FONT_PRESETS} from './docFonts';
import {
  DOC_THEMES,
  type DocBorderStyle,
  type DocColumns,
  type DocMarginPreset,
  type DocPageLayout,
  type DocPageSize,
  type DocsRibbonTab,
  type DocThemeId,
  type DocViewMode,
} from './docsTypes';

interface DocsRibbonProps {
  editor: Editor | null;
  activeTab: DocsRibbonTab;
  onChangeTab: (tab: DocsRibbonTab) => void;
  layout: DocPageLayout;
  onUpdateLayout: (patch: Partial<DocPageLayout>) => void;
  viewMode: DocViewMode;
  onChangeViewMode: (mode: DocViewMode) => void;
  zoom: number;
  onChangeZoom: (z: number) => void;
  showRuler: boolean;
  onToggleRuler: () => void;
  showFormattingMarks: boolean;
  onToggleFormattingMarks: () => void;
  spellCheck: boolean;
  onToggleSpellCheck: () => void;
  formatPainterActive: boolean;
  onToggleFormatPainter: () => void;
  isFavorite: boolean;
  onToggleFavorite: () => void;
  commentCount: number;
  commentsOpen: boolean;
  onToggleComments: () => void;
  onOpenDialog: (
    kind:
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
      | 'export',
  ) => void;
  onPickImage: () => void;
  onInsertCoverPage: () => void;
  onInsertToc: () => void;
  onInsertNote: (kind: 'footnote' | 'endnote') => void;
  onInsertCaption: () => void;
  onAddComment: () => void;
  onSaveNow: () => void;
  onExportPdf: () => void;
  onPrint: () => void;
  onDoneEditing: () => void;
  currentPage?: number;
  totalPages?: number;
  onJumpToPage?: (page: number) => void;
}

const RIBBON_TABS: Array<{id: DocsRibbonTab; label: string}> = [
  {id: 'home', label: 'Home'},
  {id: 'insert', label: 'Insert'},
  {id: 'page-layout', label: 'Page Layout'},
  {id: 'references', label: 'References'},
  {id: 'review', label: 'Review'},
  {id: 'view', label: 'View'},
];

const FONT_FAMILIES = [...DOCUMENT_FONT_PRESETS];

const FONT_SIZES = ['10pt', '11pt', '12pt', '14pt', '16pt', '18pt', '22pt', '26pt', '32pt'];

const LINE_SPACINGS = [
  {label: '1.0', value: '1.0'},
  {label: '1.15', value: '1.15'},
  {label: '1.5', value: '1.5'},
  {label: '2.0', value: '2.0'},
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
        ? 'bg-blue-100 text-blue-800 dark:bg-blue-950/70 dark:text-blue-300'
        : 'text-slate-700 hover:bg-slate-100 dark:text-slate-200 dark:hover:bg-slate-800'
    }`}
  >
    {children}
    {label && <span>{label}</span>}
  </button>
);

export const DocsRibbon: React.FC<DocsRibbonProps> = ({
  editor,
  activeTab,
  onChangeTab,
  layout,
  onUpdateLayout,
  viewMode,
  onChangeViewMode,
  zoom,
  onChangeZoom,
  showRuler,
  onToggleRuler,
  showFormattingMarks,
  onToggleFormattingMarks,
  spellCheck,
  onToggleSpellCheck,
  formatPainterActive,
  onToggleFormatPainter,
  isFavorite,
  onToggleFavorite,
  commentCount,
  commentsOpen,
  onToggleComments,
  onOpenDialog,
  onPickImage,
  onInsertCoverPage,
  onInsertToc,
  onInsertNote,
  onInsertCaption,
  onAddComment,
  onSaveNow,
  onExportPdf,
  onPrint,
  onDoneEditing,
  currentPage = 1,
  totalPages = 1,
  onJumpToPage,
}) => {
  const defaultDocFont =
    canonicalizeWordFontFamily(layout.docDefaultFont) ||
    DOC_THEMES[layout.theme].canonicalFontFamily;
  const activeFont =
    canonicalizeWordFontFamily(
      editor?.getAttributes('textStyle')?.fontFamily as string | undefined,
    ) || defaultDocFont;
  const fontOptions = FONT_FAMILIES.includes(activeFont)
    ? FONT_FAMILIES
    : [activeFont, ...FONT_FAMILIES];
  const activeSize =
    (editor?.getAttributes('textStyle')?.fontSize as string | undefined) ||
    (layout.docDefaultFontSizePt ? `${layout.docDefaultFontSizePt}pt` : '11pt');
  const activeColor =
    (editor?.getAttributes('textStyle')?.color as string | undefined) ||
    layout.docDefaultColor ||
    '#0f172a';
  const activeLineHeight =
    (editor?.getAttributes('textStyle')?.lineHeight as string | undefined) || '1.5';

  return (
    <div className="border-b border-slate-200 bg-white dark:border-slate-800 dark:bg-slate-900">
      {/* Word 2007-style Quick Access + Ribbon Tab Strip */}
      <div className="flex items-center justify-between gap-2 overflow-x-auto border-b border-slate-200 bg-slate-50 px-2.5 pt-1.5 dark:border-slate-800 dark:bg-slate-950">
        <div className="flex shrink-0 items-center gap-1">
          {/* Quick Access Buttons */}
          <div className="mr-2 flex shrink-0 items-center gap-0.5 border-r border-slate-200 pr-2 pb-1 dark:border-slate-800">
            <ToolBtn onClick={onSaveNow} title="Save document (Ctrl+S)">
              <Save className="h-3.5 w-3.5 text-blue-700 dark:text-blue-400" />
            </ToolBtn>
            <ToolBtn
              disabled={!editor?.can().undo()}
              onClick={() => editor?.chain().focus().undo().run()}
              title="Undo (Ctrl+Z)"
            >
              <Undo2 className="h-3.5 w-3.5" />
            </ToolBtn>
            <ToolBtn
              disabled={!editor?.can().redo()}
              onClick={() => editor?.chain().focus().redo().run()}
              title="Redo (Ctrl+Y)"
            >
              <Redo2 className="h-3.5 w-3.5" />
            </ToolBtn>
            <ToolBtn onClick={onToggleFavorite} title={isFavorite ? 'Remove favourite' : 'Add favourite'}>
              <Star className={`h-3.5 w-3.5 ${isFavorite ? 'fill-amber-400 text-amber-500' : ''}`} />
            </ToolBtn>
          </div>

          {/* 6 Classic Ribbon Tabs */}
          <div role="tablist" aria-label="Document ribbon tabs" className="flex shrink-0 items-end gap-0.5">
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
                      ? '-mb-px border-slate-200 bg-white text-blue-700 dark:border-slate-800 dark:bg-slate-900 dark:text-blue-400'
                      : 'border-transparent text-slate-600 hover:bg-slate-200/60 hover:text-slate-900 dark:text-slate-400 dark:hover:bg-slate-800 dark:hover:text-slate-100'
                  }`}
                >
                  {t.label}
                </button>
              );
            })}
          </div>
        </div>

        {/* Right File Actions */}
        <div className="flex shrink-0 items-center gap-1 pb-1">
          {viewMode === 'print' && onJumpToPage && (
            <div className="mr-1 hidden items-center rounded-lg border border-slate-200 bg-white sm:flex dark:border-slate-700 dark:bg-slate-900">
              <button
                type="button"
                onClick={() => onJumpToPage(currentPage - 1)}
                disabled={currentPage <= 1}
                aria-label="Previous page"
                title="Previous page"
                className="rounded-l-lg p-1 text-slate-600 hover:bg-slate-100 disabled:opacity-35 dark:text-slate-300 dark:hover:bg-slate-800"
              >
                <ChevronLeft className="h-3.5 w-3.5" />
              </button>
              <span
                data-testid="docs-toolbar-page-indicator"
                className="whitespace-nowrap px-1.5 text-xs font-medium tabular-nums text-slate-700 dark:text-slate-200"
              >
                Page {currentPage} of {totalPages}
              </span>
              <button
                type="button"
                onClick={() => onJumpToPage(currentPage + 1)}
                disabled={currentPage >= totalPages}
                aria-label="Next page"
                title="Next page"
                className="rounded-r-lg p-1 text-slate-600 hover:bg-slate-100 disabled:opacity-35 dark:text-slate-300 dark:hover:bg-slate-800"
              >
                <ChevronRight className="h-3.5 w-3.5" />
              </button>
            </div>
          )}
          <ToolBtn onClick={() => onOpenDialog('versions')} title="Version History" label="History">
            <History className="h-3.5 w-3.5 text-slate-600 dark:text-slate-300" />
          </ToolBtn>
          <ToolBtn onClick={onExportPdf} title="Export as PDF" label="PDF">
            <FileDown className="h-3.5 w-3.5 text-rose-600 dark:text-rose-400" />
          </ToolBtn>
          <ToolBtn onClick={onPrint} title="Print document">
            <Printer className="h-3.5 w-3.5" />
          </ToolBtn>
          <button
            type="button"
            onClick={() => onOpenDialog('export')}
            className="inline-flex min-h-[30px] shrink-0 items-center gap-1 whitespace-nowrap rounded-lg border border-slate-200 bg-white px-2.5 text-xs font-semibold text-slate-800 hover:bg-slate-100 dark:border-slate-700 dark:bg-slate-800 dark:text-slate-100 dark:hover:bg-slate-700"
          >
            <Download className="h-3.5 w-3.5" />
            Save As
          </button>
          <button
            type="button"
            onClick={onDoneEditing}
            aria-label="Done editing"
            title="Save and return to View mode (Esc)"
            className="inline-flex min-h-[30px] shrink-0 items-center gap-1 whitespace-nowrap rounded-lg bg-blue-700 px-3 text-xs font-semibold text-white hover:bg-blue-800 dark:bg-blue-600 dark:hover:bg-blue-500"
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
            {/* Clipboard & Format Painter */}
            <RibbonGroup label="Format">
              <ToolBtn
                active={formatPainterActive}
                onClick={onToggleFormatPainter}
                title="Format Painter (copy selection formatting, click target text to apply)"
                label="Painter"
              >
                <Paintbrush className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                onClick={() => editor?.chain().focus().clearNodes().unsetAllMarks().run()}
                title="Clear all formatting"
                label="Clear"
              >
                <Eraser className="h-3.5 w-3.5" />
              </ToolBtn>
            </RibbonGroup>

            {/* Font */}
            <RibbonGroup label="Font">
              <select
                aria-label="Font family"
                value={activeFont}
                onChange={(e) => editor?.chain().focus().setFontFamily(e.target.value).run()}
                className="h-7 w-[104px] rounded border border-slate-200 bg-slate-50 px-1.5 text-xs text-slate-800 dark:border-slate-700 dark:bg-slate-950 dark:text-slate-100"
              >
                {fontOptions.map((f) => (
                  <option key={f} value={f}>
                    {f}
                  </option>
                ))}
              </select>
              <select
                aria-label="Font size"
                value={activeSize}
                onChange={(e) => editor?.chain().focus().setFontSize(e.target.value).run()}
                className="h-7 w-14 rounded border border-slate-200 bg-slate-50 px-1 text-xs tabular-nums text-slate-800 dark:border-slate-700 dark:bg-slate-950 dark:text-slate-100"
              >
                {FONT_SIZES.map((s) => (
                  <option key={s} value={s}>
                    {s}
                  </option>
                ))}
              </select>
              <ToolBtn
                active={editor?.isActive('bold')}
                onClick={() => editor?.chain().focus().toggleBold().run()}
                title="Bold (Ctrl+B)"
              >
                <Bold className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                active={editor?.isActive('italic')}
                onClick={() => editor?.chain().focus().toggleItalic().run()}
                title="Italic (Ctrl+I)"
              >
                <Italic className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                active={editor?.isActive('underline')}
                onClick={() => editor?.chain().focus().toggleUnderline().run()}
                title="Underline (Ctrl+U)"
              >
                <Underline className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                active={editor?.isActive('strike')}
                onClick={() => editor?.chain().focus().toggleStrike().run()}
                title="Strikethrough"
              >
                <Strikethrough className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                active={editor?.isActive('subscript')}
                onClick={() => editor?.chain().focus().toggleSubscript().run()}
                title="Subscript"
              >
                <Subscript className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                active={editor?.isActive('superscript')}
                onClick={() => editor?.chain().focus().toggleSuperscript().run()}
                title="Superscript"
              >
                <Superscript className="h-3.5 w-3.5" />
              </ToolBtn>

              {/* Text Color */}
              <label
                title="Font color"
                className="inline-flex h-7 cursor-pointer items-center gap-0.5 rounded border border-slate-200 px-1 dark:border-slate-700"
              >
                <Type className="h-3.5 w-3.5 text-slate-600 dark:text-slate-300" />
                <input
                  type="color"
                  aria-label="Font color"
                  value={activeColor}
                  onChange={(e) => editor?.chain().focus().setColor(e.target.value).run()}
                  className="h-4 w-4 cursor-pointer border-0 bg-transparent p-0"
                />
              </label>

              {/* Highlight */}
              <ToolBtn
                active={editor?.isActive('highlight')}
                onClick={() =>
                  editor?.chain().focus().toggleHighlight({color: '#FEF08A'}).run()
                }
                title="Text highlight color"
              >
                <Highlighter className="h-3.5 w-3.5 text-amber-600 dark:text-amber-400" />
              </ToolBtn>
            </RibbonGroup>

            {/* Paragraph */}
            <RibbonGroup label="Paragraph">
              <ToolBtn
                active={editor?.isActive({textAlign: 'left'})}
                onClick={() => editor?.chain().focus().setTextAlign('left').run()}
                title="Align Left"
              >
                <AlignLeft className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                active={editor?.isActive({textAlign: 'center'})}
                onClick={() => editor?.chain().focus().setTextAlign('center').run()}
                title="Center"
              >
                <AlignCenter className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                active={editor?.isActive({textAlign: 'right'})}
                onClick={() => editor?.chain().focus().setTextAlign('right').run()}
                title="Align Right"
              >
                <AlignRight className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                active={editor?.isActive({textAlign: 'justify'})}
                onClick={() => editor?.chain().focus().setTextAlign('justify').run()}
                title="Justify"
              >
                <AlignJustify className="h-3.5 w-3.5" />
              </ToolBtn>
              <select
                aria-label="Line spacing"
                title="Line spacing"
                value={activeLineHeight}
                onChange={(e) => editor?.chain().focus().setLineHeight(e.target.value).run()}
                className="h-7 rounded border border-slate-200 bg-slate-50 px-1 text-xs tabular-nums dark:border-slate-700 dark:bg-slate-950"
              >
                {LINE_SPACINGS.map((ls) => (
                  <option key={ls.value} value={ls.value}>
                    {ls.label}x
                  </option>
                ))}
              </select>
              <ToolBtn
                active={editor?.isActive('bulletList')}
                onClick={() => editor?.chain().focus().toggleBulletList().run()}
                title="Bullets"
              >
                <List className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                active={editor?.isActive('orderedList')}
                onClick={() => editor?.chain().focus().toggleOrderedList().run()}
                title="Numbering"
              >
                <ListOrdered className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                onClick={() => {
                  if (editor?.isActive('listItem')) {
                    editor.chain().focus().sinkListItem('listItem').run();
                  } else {
                    editor?.chain().focus().toggleOrderedList().run();
                  }
                }}
                title="Multilevel list (nest list item)"
              >
                <ListTree className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                onClick={() => {
                  if (editor?.isActive('listItem')) {
                    editor.chain().focus().liftListItem('listItem').run();
                  } else {
                    editor?.chain().focus().decreaseIndent().run();
                  }
                }}
                title="Decrease Indent"
              >
                <IndentDecrease className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                onClick={() => {
                  if (editor?.isActive('listItem')) {
                    editor.chain().focus().sinkListItem('listItem').run();
                  } else {
                    editor?.chain().focus().increaseIndent().run();
                  }
                }}
                title="Increase Indent"
              >
                <IndentIncrease className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                active={showFormattingMarks}
                onClick={onToggleFormattingMarks}
                title="Show formatting marks (page breaks & tabs)"
              >
                <Pilcrow className="h-3.5 w-3.5" />
              </ToolBtn>
            </RibbonGroup>

            {/* Styles Gallery */}
            <RibbonGroup label="Styles">
              {(
                [
                  ['Normal', 'Normal', () => editor?.chain().focus().setParagraph().run(), editor?.isActive('paragraph')],
                  ['H1', 'Heading 1', () => editor?.chain().focus().toggleHeading({level: 1}).run(), editor?.isActive('heading', {level: 1})],
                  ['H2', 'Heading 2', () => editor?.chain().focus().toggleHeading({level: 2}).run(), editor?.isActive('heading', {level: 2})],
                  ['H3', 'Heading 3', () => editor?.chain().focus().toggleHeading({level: 3}).run(), editor?.isActive('heading', {level: 3})],
                  ['Quote', 'Quote', () => editor?.chain().focus().toggleBlockquote().run(), editor?.isActive('blockquote')],
                ] as const
              ).map(([shortLabel, ariaTitle, handler, active]) => (
                <button
                  key={ariaTitle}
                  type="button"
                  title={ariaTitle}
                  aria-label={ariaTitle}
                  onClick={handler}
                  className={`h-7 shrink-0 whitespace-nowrap rounded border px-1.5 text-xs transition-colors ${
                    active
                      ? 'border-blue-600 bg-blue-50 font-bold text-blue-800 dark:border-blue-500 dark:bg-blue-950/60 dark:text-blue-300'
                      : 'border-slate-200 bg-slate-50 text-slate-700 hover:bg-slate-100 dark:border-slate-700 dark:bg-slate-950 dark:text-slate-200'
                  }`}
                >
                  {shortLabel}
                </button>
              ))}
            </RibbonGroup>

            {/* Editing */}
            <RibbonGroup label="Editing">
              <ToolBtn
                onClick={() => onOpenDialog('find')}
                title="Find & Replace (Ctrl+H)"
                label="Find"
              >
                <Search className="h-3.5 w-3.5" />
              </ToolBtn>
            </RibbonGroup>
          </>
        )}

        {activeTab === 'insert' && (
          <>
            <RibbonGroup label="Pages">
              <ToolBtn onClick={onInsertCoverPage} title="Insert formatted Cover Page" label="Cover Page">
                <BookOpen className="h-3.5 w-3.5 text-blue-700 dark:text-blue-400" />
              </ToolBtn>
              <ToolBtn
                onClick={() => editor?.chain().focus().insertPageBreak().run()}
                title="Insert Page Break"
                label="Page Break"
              >
                <Scissors className="h-3.5 w-3.5" />
              </ToolBtn>
            </RibbonGroup>

            <RibbonGroup label="Tables">
              <ToolBtn onClick={() => onOpenDialog('table')} title="Insert Table" label="Table">
                <Table2 className="h-3.5 w-3.5 text-blue-700 dark:text-blue-400" />
              </ToolBtn>
              {editor?.isActive('table') && (
                <>
                  <ToolBtn
                    onClick={() => editor.chain().focus().addRowAfter().run()}
                    title="Add Row Below"
                    label="+Row"
                  >
                    <span />
                  </ToolBtn>
                  <ToolBtn
                    onClick={() => editor.chain().focus().addColumnAfter().run()}
                    title="Add Column Right"
                    label="+Col"
                  >
                    <span />
                  </ToolBtn>
                  <ToolBtn
                    onClick={() => editor.chain().focus().deleteTable().run()}
                    title="Delete Table"
                    label="Del Table"
                  >
                    <span />
                  </ToolBtn>
                </>
              )}
            </RibbonGroup>

            <RibbonGroup label="Illustrations">
              <ToolBtn onClick={onPickImage} title="Insert Image from device" label="Image">
                <ImagePlus className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn onClick={() => onOpenDialog('shape')} title="Insert Shape" label="Shapes">
                <Shapes className="h-3.5 w-3.5" />
              </ToolBtn>
            </RibbonGroup>

            <RibbonGroup label="Links">
              <ToolBtn onClick={() => onOpenDialog('link')} title="Insert Hyperlink" label="Hyperlink">
                <Link2 className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn onClick={() => onOpenDialog('bookmark')} title="Insert Bookmark anchor" label="Bookmark">
                <Bookmark className="h-3.5 w-3.5" />
              </ToolBtn>
            </RibbonGroup>

            <RibbonGroup label="Header & Footer">
              <ToolBtn
                onClick={() => onOpenDialog('header-footer')}
                title="Edit Header, Footer & Page Numbers"
                label="Header & Footer"
              >
                <PanelTop className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                active={layout.showPageNumbers}
                onClick={() => onUpdateLayout({showPageNumbers: !layout.showPageNumbers})}
                title="Toggle Page Numbers"
                label="Page Numbers"
              >
                <Hash className="h-3.5 w-3.5" />
              </ToolBtn>
            </RibbonGroup>

            <RibbonGroup label="Text & Symbols">
              <ToolBtn
                onClick={() => editor?.chain().focus().insertDocTextBox().run()}
                title="Insert Callout Text Box"
                label="Text Box"
              >
                <SquareDashedBottom className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn onClick={() => onOpenDialog('symbol')} title="Insert Special Symbol" label="Symbol">
                <Omega className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn onClick={() => onOpenDialog('equation')} title="Insert Equation block" label="Equation">
                <Sigma className="h-3.5 w-3.5" />
              </ToolBtn>
            </RibbonGroup>
          </>
        )}

        {activeTab === 'page-layout' && (
          <>
            <RibbonGroup label="Themes">
              <select
                aria-label="Document theme"
                value={layout.theme}
                onChange={(e) => onUpdateLayout({theme: e.target.value as DocThemeId})}
                className="h-7 rounded border border-slate-200 bg-slate-50 px-2 text-xs font-medium dark:border-slate-700 dark:bg-slate-950"
              >
                {(Object.keys(DOC_THEMES) as DocThemeId[]).map((id) => (
                  <option key={id} value={id}>
                    {DOC_THEMES[id].name}
                  </option>
                ))}
              </select>
            </RibbonGroup>

            <RibbonGroup label="Page Setup">
              <select
                aria-label="Page margins"
                value={layout.margins}
                onChange={(e) => onUpdateLayout({margins: e.target.value as DocMarginPreset})}
                className="h-7 rounded border border-slate-200 bg-slate-50 px-2 text-xs dark:border-slate-700 dark:bg-slate-950"
              >
                <option value="normal">Margins: Normal (1&quot;)</option>
                <option value="narrow">Margins: Narrow (0.5&quot;)</option>
                <option value="moderate">Margins: Moderate (0.75&quot;)</option>
                <option value="wide">Margins: Wide (1.5&quot;)</option>
              </select>

              <select
                aria-label="Page orientation"
                value={layout.orientation}
                onChange={(e) =>
                  onUpdateLayout({orientation: e.target.value as 'portrait' | 'landscape'})
                }
                className="h-7 rounded border border-slate-200 bg-slate-50 px-2 text-xs dark:border-slate-700 dark:bg-slate-950"
              >
                <option value="portrait">Portrait</option>
                <option value="landscape">Landscape</option>
              </select>

              <select
                aria-label="Paper size"
                value={layout.size}
                onChange={(e) => onUpdateLayout({size: e.target.value as DocPageSize})}
                className="h-7 rounded border border-slate-200 bg-slate-50 px-2 text-xs dark:border-slate-700 dark:bg-slate-950"
              >
                <option value="a4">A4</option>
                <option value="letter">US Letter</option>
                <option value="legal">Legal</option>
              </select>

              <div className="flex items-center gap-1">
                <Columns3 className="h-3.5 w-3.5 text-slate-500" />
                <select
                  aria-label="Page columns"
                  value={layout.columns}
                  onChange={(e) => onUpdateLayout({columns: Number(e.target.value) as DocColumns})}
                  className="h-7 rounded border border-slate-200 bg-slate-50 px-1.5 text-xs dark:border-slate-700 dark:bg-slate-950"
                >
                  <option value={1}>1 Column</option>
                  <option value={2}>2 Columns</option>
                  <option value={3}>3 Columns</option>
                </select>
              </div>
            </RibbonGroup>

            <RibbonGroup label="Page Background">
              <label className="flex items-center gap-1.5 text-xs font-medium text-slate-700 dark:text-slate-300">
                <input
                  type="checkbox"
                  checked={layout.watermarkEnabled}
                  onChange={(e) => onUpdateLayout({watermarkEnabled: e.target.checked})}
                  className="h-3.5 w-3.5 accent-blue-600"
                />
                Watermark
              </label>
              {layout.watermarkEnabled && (
                <input
                  type="text"
                  aria-label="Watermark text"
                  value={layout.watermarkText}
                  onChange={(e) => onUpdateLayout({watermarkText: e.target.value})}
                  className="h-7 w-28 rounded border border-slate-200 bg-slate-50 px-2 text-xs dark:border-slate-700 dark:bg-slate-950"
                />
              )}
              <select
                aria-label="Page border style"
                value={layout.pageBorder}
                onChange={(e) => onUpdateLayout({pageBorder: e.target.value as DocBorderStyle})}
                className="h-7 rounded border border-slate-200 bg-slate-50 px-2 text-xs dark:border-slate-700 dark:bg-slate-950"
              >
                <option value="none">Border: None</option>
                <option value="box">Border: Box</option>
                <option value="double">Border: Double</option>
                <option value="shadow">Border: Accent Frame</option>
              </select>
            </RibbonGroup>
          </>
        )}

        {activeTab === 'references' && (
          <>
            <RibbonGroup label="Table of Contents">
              <ToolBtn
                onClick={onInsertToc}
                title="Generate and insert Table of Contents from Headings 1-3"
                label="Insert Table of Contents"
              >
                <BookMarked className="h-3.5 w-3.5 text-blue-700 dark:text-blue-400" />
              </ToolBtn>
            </RibbonGroup>

            <RibbonGroup label="Footnotes & Endnotes">
              <ToolBtn
                onClick={() => onInsertNote('footnote')}
                title="Insert numbered Footnote at cursor"
                label="Insert Footnote"
              >
                <Pilcrow className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                onClick={() => onInsertNote('endnote')}
                title="Insert numbered Endnote at cursor"
                label="Insert Endnote"
              >
                <FileText className="h-3.5 w-3.5" />
              </ToolBtn>
            </RibbonGroup>

            <RibbonGroup label="Captions">
              <ToolBtn
                onClick={onInsertCaption}
                title="Insert Figure / Table Caption"
                label="Insert Caption"
              >
                <Type className="h-3.5 w-3.5" />
              </ToolBtn>
            </RibbonGroup>
          </>
        )}

        {activeTab === 'review' && (
          <>
            <RibbonGroup label="Proofing">
              <ToolBtn
                active={spellCheck}
                onClick={onToggleSpellCheck}
                title="Toggle browser spell check"
                label={spellCheck ? 'Spell Check: On' : 'Spell Check: Off'}
              >
                <SpellCheck2 className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                onClick={() => onOpenDialog('word-count')}
                title="View detailed Word Count statistics"
                label="Word Count"
              >
                <Hash className="h-3.5 w-3.5" />
              </ToolBtn>
            </RibbonGroup>

            <RibbonGroup label="Comments">
              <ToolBtn onClick={onAddComment} title="New Comment on current selection" label="New Comment">
                <MessageSquarePlus className="h-3.5 w-3.5 text-blue-700 dark:text-blue-400" />
              </ToolBtn>
              <ToolBtn
                active={commentsOpen}
                onClick={onToggleComments}
                title="Show or hide Comments panel"
                label={`Comments (${commentCount})`}
              >
                <span />
              </ToolBtn>
            </RibbonGroup>

            <RibbonGroup label="Tracking">
              <span className="px-2 text-[11px] text-slate-500 dark:text-slate-400">
                Track Changes planned for a later phase; use Version History to compare snapshots.
              </span>
            </RibbonGroup>
          </>
        )}

        {activeTab === 'view' && (
          <>
            <RibbonGroup label="Document Views">
              <ToolBtn
                active={viewMode === 'print'}
                onClick={() => onChangeViewMode('print')}
                title="Print Layout with real page boundaries"
                label="Print Layout"
              >
                <FileText className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                active={viewMode === 'web'}
                onClick={() => onChangeViewMode('web')}
                title="Fluid Web Layout"
                label="Web Layout"
              >
                <Globe className="h-3.5 w-3.5" />
              </ToolBtn>
            </RibbonGroup>

            <RibbonGroup label="Show">
              <ToolBtn
                active={showRuler}
                onClick={onToggleRuler}
                title="Toggle Horizontal Ruler"
                label="Ruler"
              >
                <Ruler className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                active={showFormattingMarks}
                onClick={onToggleFormattingMarks}
                title="Show formatting marks (page breaks & tabs)"
                label="Formatting Marks"
              >
                <Pilcrow className="h-3.5 w-3.5" />
              </ToolBtn>
            </RibbonGroup>

            <RibbonGroup label="Zoom">
              <ToolBtn
                onClick={() => onChangeZoom(Math.max(0.5, Number((zoom - 0.1).toFixed(2))))}
                title="Zoom Out"
              >
                <ZoomOut className="h-3.5 w-3.5" />
              </ToolBtn>
              <span className="min-w-[44px] text-center text-xs font-semibold tabular-nums text-slate-700 dark:text-slate-200">
                {Math.round(zoom * 100)}%
              </span>
              <ToolBtn
                onClick={() => onChangeZoom(Math.min(1.8, Number((zoom + 0.1).toFixed(2))))}
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
