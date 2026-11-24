import React, {useRef} from 'react';
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  ArrowRight,
  BarChart3,
  Bold,
  Check,
  Circle,
  Copy,
  Download,
  FileText,
  Image as ImageIcon,
  Italic,
  LayoutTemplate,
  Layers,
  List,
  ListOrdered,
  Minus,
  MonitorPlay,
  Palette,
  Play,
  Plus,
  Printer,
  Redo2,
  Save,
  Sparkles,
  Square,
  Star,
  Table2,
  Trash2,
  Triangle,
  Type,
  Underline,
  Undo2,
  Wand2,
} from 'lucide-react';
import {SLIDE_THEMES} from './slidesPure';
import {
  SLIDE_FONT_FAMILIES,
  SLIDE_FONT_SIZES,
  type ShapeKind,
  type SlideAspectRatio,
  type SlideChartKind,
  type SlideDialogKind,
  type SlideElement,
  type SlideLayoutKind,
  type SlideRibbonTab,
  type SlideTransitionKind,
} from './slidesTypes';

interface SlidesRibbonProps {
  activeTab: SlideRibbonTab;
  onChangeTab: (tab: SlideRibbonTab) => void;
  dirty: boolean;
  isFavorite: boolean;
  canUndo: boolean;
  canRedo: boolean;
  themeId: string;
  aspectRatio?: SlideAspectRatio;
  slideTransition: SlideTransitionKind;
  slideBackground: string;
  selectedElement: SlideElement | null;
  onSave: () => void;
  onUndo: () => void;
  onRedo: () => void;
  onToggleFavorite: () => void;
  onAddSlide: (layout: SlideLayoutKind) => void;
  onDuplicateSlide: () => void;
  onDeleteSlide: () => void;
  onApplyLayout: (layout: SlideLayoutKind) => void;
  onAddText: () => void;
  onAddShape: (kind: ShapeKind) => void;
  onInsertImageFile: (file: File) => void;
  onOpenDialog: (dialog: SlideDialogKind, defaultChartKind?: SlideChartKind) => void;
  onUpdateSelectedElement: (patch: Partial<SlideElement>) => void;
  onDeleteSelectedElement: () => void;
  onDuplicateSelectedElement: () => void;
  onLayerAction: (dir: 'front' | 'back') => void;
  onAlignElement: (align: 'left' | 'center' | 'right' | 'top' | 'middle' | 'bottom') => void;
  onSelectTheme: (themeId: string) => void;
  onChangeAspectRatio?: (ratio: SlideAspectRatio) => void;
  onChangeBackground: (hex: string) => void;
  onChangeTransition: (transition: SlideTransitionKind) => void;
  onApplyTransitionAll: (transition: SlideTransitionKind) => void;
  onStartPresent: (fromStart: boolean) => void;
  onExportPdf: () => void;
  onExportPng: () => void;
  onPrint: () => void;
  onDoneEditing: () => void;
}

const ToolBtn: React.FC<{
  onClick: () => void;
  active?: boolean;
  disabled?: boolean;
  title: string;
  label?: string;
  children: React.ReactNode;
}> = ({onClick, active, disabled, title, label, children}) => (
  <button
    type="button"
    onClick={onClick}
    disabled={disabled}
    title={title}
    aria-label={title}
    aria-pressed={active}
    className={`inline-flex h-7 shrink-0 items-center justify-center gap-1 whitespace-nowrap rounded px-2 text-xs font-medium transition-colors disabled:opacity-40 ${
      active
        ? 'bg-orange-100 text-orange-800 dark:bg-orange-950/70 dark:text-orange-300'
        : 'text-slate-700 hover:bg-slate-100 dark:text-slate-200 dark:hover:bg-slate-800'
    }`}
  >
    {children}
    {label && <span>{label}</span>}
  </button>
);

const RibbonGroup: React.FC<{label: string; children: React.ReactNode}> = ({label, children}) => (
  <div className="flex shrink-0 flex-col justify-between border-r border-slate-200 pr-2.5 last:border-r-0 dark:border-slate-800">
    <div className="flex items-center gap-1">{children}</div>
    <span className="mt-0.5 text-center text-[10px] font-medium text-slate-400 dark:text-slate-500">
      {label}
    </span>
  </div>
);

const TABS: Array<{id: SlideRibbonTab; label: string}> = [
  {id: 'home', label: 'Home'},
  {id: 'insert', label: 'Insert'},
  {id: 'design', label: 'Design'},
  {id: 'transitions', label: 'Transitions'},
  {id: 'slideshow', label: 'Slide Show'},
];

const LAYOUTS: Array<{id: SlideLayoutKind; label: string}> = [
  {id: 'title', label: 'Title'},
  {id: 'title-content', label: 'Title & Content'},
  {id: 'two-column', label: 'Two Column'},
  {id: 'section', label: 'Section'},
  {id: 'comparison', label: 'Table Matrix'},
  {id: 'image-caption', label: 'Chart & Notes'},
  {id: 'blank', label: 'Blank'},
];

const TRANSITIONS: Array<{id: SlideTransitionKind; label: string}> = [
  {id: 'none', label: 'None'},
  {id: 'fade', label: 'Fade'},
  {id: 'slide', label: 'Slide'},
  {id: 'zoom', label: 'Zoom'},
];

export const SlidesRibbon: React.FC<SlidesRibbonProps> = ({
  activeTab,
  onChangeTab,
  dirty,
  isFavorite,
  canUndo,
  canRedo,
  themeId,
  aspectRatio = '16:9',
  slideTransition,
  slideBackground,
  selectedElement,
  onSave,
  onUndo,
  onRedo,
  onToggleFavorite,
  onAddSlide,
  onDuplicateSlide,
  onDeleteSlide,
  onApplyLayout,
  onAddText,
  onAddShape,
  onInsertImageFile,
  onOpenDialog,
  onUpdateSelectedElement,
  onDeleteSelectedElement,
  onDuplicateSelectedElement,
  onLayerAction,
  onAlignElement,
  onSelectTheme,
  onChangeAspectRatio,
  onChangeBackground,
  onChangeTransition,
  onApplyTransitionAll,
  onStartPresent,
  onExportPdf,
  onExportPng,
  onPrint,
  onDoneEditing,
}) => {
  const imageInputRef = useRef<HTMLInputElement | null>(null);
  const textEl = selectedElement?.type === 'text' ? selectedElement : null;
  const fontOptions =
    textEl?.fontFamily && !SLIDE_FONT_FAMILIES.includes(textEl.fontFamily)
      ? [textEl.fontFamily, ...SLIDE_FONT_FAMILIES]
      : SLIDE_FONT_FAMILIES;

  return (
    <div className="shrink-0 border-b border-slate-200 bg-white dark:border-slate-800 dark:bg-slate-900">
      <input
        ref={imageInputRef}
        type="file"
        accept="image/png,image/jpeg,image/webp,image/gif"
        className="hidden"
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) onInsertImageFile(f);
          e.target.value = '';
        }}
      />

      {/* Row 1: Quick Access + Ribbon Tabs + Primary Actions */}
      <div className="flex items-center justify-between gap-2 overflow-x-auto border-b border-slate-100 px-3 py-1 dark:border-slate-800">
        <div className="flex shrink-0 items-center gap-1">
          <ToolBtn onClick={onSave} title="Save presentation (Ctrl+S)" active={dirty}>
            <Save className="h-3.5 w-3.5 text-orange-700 dark:text-orange-400" />
          </ToolBtn>
          <ToolBtn onClick={onUndo} disabled={!canUndo} title="Undo (Ctrl+Z)">
            <Undo2 className="h-3.5 w-3.5" />
          </ToolBtn>
          <ToolBtn onClick={onRedo} disabled={!canRedo} title="Redo (Ctrl+Y)">
            <Redo2 className="h-3.5 w-3.5" />
          </ToolBtn>
          <ToolBtn
            onClick={onToggleFavorite}
            title={isFavorite ? 'Remove from Favourites' : 'Add to Favourites'}
          >
            <Star
              className={`h-3.5 w-3.5 ${
                isFavorite ? 'fill-amber-400 text-amber-500' : 'text-slate-500 dark:text-slate-400'
              }`}
            />
          </ToolBtn>

          <div className="mx-1.5 h-4 w-px bg-slate-200 dark:bg-slate-700" />

          <div role="tablist" aria-label="Presentation ribbon tabs" className="flex items-center gap-0.5">
            {TABS.map((t) => {
              const active = activeTab === t.id;
              return (
                <button
                  key={t.id}
                  type="button"
                  role="tab"
                  aria-selected={active}
                  onClick={() => onChangeTab(t.id)}
                  className={`whitespace-nowrap rounded-t-md px-3 py-1 text-xs font-semibold transition-colors ${
                    active
                      ? 'border-b-2 border-orange-600 bg-orange-50/70 text-orange-800 dark:border-orange-400 dark:bg-orange-950/40 dark:text-orange-300'
                      : 'text-slate-600 hover:bg-slate-100 hover:text-slate-900 dark:text-slate-300 dark:hover:bg-slate-800'
                  }`}
                >
                  {t.label}
                </button>
              );
            })}
          </div>
        </div>

        <div className="flex shrink-0 items-center gap-1.5">
          <button
            type="button"
            onClick={() => onStartPresent(false)}
            aria-label="Present slide show (F5)"
            className="inline-flex h-7 shrink-0 items-center gap-1.5 whitespace-nowrap rounded-md bg-orange-600 px-3 text-xs font-semibold text-white hover:bg-orange-700 dark:bg-orange-600 dark:hover:bg-orange-500"
          >
            <Play className="h-3.5 w-3.5 fill-current" />
            Present
          </button>
          <ToolBtn onClick={onExportPdf} title="Export deck to PDF" label="PDF">
            <FileText className="h-3.5 w-3.5 text-rose-600 dark:text-rose-400" />
          </ToolBtn>
          <ToolBtn onClick={onExportPng} title="Export active slide as PNG" label="PNG">
            <ImageIcon className="h-3.5 w-3.5 text-sky-600 dark:text-sky-400" />
          </ToolBtn>
          <ToolBtn onClick={onPrint} title="Print presentation">
            <Printer className="h-3.5 w-3.5" />
          </ToolBtn>
          <button
            type="button"
            onClick={() => onOpenDialog('export')}
            className="inline-flex h-7 shrink-0 items-center gap-1 whitespace-nowrap rounded-md border border-slate-300 bg-white px-2.5 text-xs font-semibold text-slate-800 hover:bg-slate-50 dark:border-slate-700 dark:bg-slate-800 dark:text-slate-100 dark:hover:bg-slate-700"
          >
            <Download className="h-3.5 w-3.5" />
            Save As
          </button>
          <button
            type="button"
            onClick={onDoneEditing}
            aria-label="Done editing"
            title="Save and return to View mode (Esc)"
            className="inline-flex h-7 shrink-0 items-center gap-1 whitespace-nowrap rounded-md bg-slate-900 px-3 text-xs font-semibold text-white hover:bg-slate-800 dark:bg-slate-100 dark:text-slate-900 dark:hover:bg-white"
          >
            <Check className="h-3.5 w-3.5" />
            Done
          </button>
        </div>
      </div>

      {/* Row 2: Active Tab Commands */}
      <div className="flex items-stretch gap-3 overflow-x-auto px-3 py-1.5">
        {activeTab === 'home' && (
          <>
            <RibbonGroup label="Slides">
              <ToolBtn
                onClick={() => onAddSlide('title-content')}
                title="Add new slide"
                label="New Slide"
              >
                <Plus className="h-3.5 w-3.5 text-orange-600 dark:text-orange-400" />
              </ToolBtn>
              <select
                aria-label="Change slide layout"
                defaultValue=""
                onChange={(e) => {
                  if (e.target.value) {
                    onApplyLayout(e.target.value as SlideLayoutKind);
                    e.target.value = '';
                  }
                }}
                className="h-7 rounded border border-slate-200 bg-slate-50 px-1.5 text-xs text-slate-800 dark:border-slate-700 dark:bg-slate-950 dark:text-slate-100"
              >
                <option value="" disabled>
                  Layout...
                </option>
                {LAYOUTS.map((l) => (
                  <option key={l.id} value={l.id}>
                    {l.label}
                  </option>
                ))}
              </select>
              <ToolBtn onClick={onDuplicateSlide} title="Duplicate active slide">
                <Copy className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn onClick={onDeleteSlide} title="Delete active slide">
                <Trash2 className="h-3.5 w-3.5 text-rose-600 dark:text-rose-400" />
              </ToolBtn>
            </RibbonGroup>

            <RibbonGroup label="Font">
              <select
                aria-label="Text font family"
                disabled={!textEl}
                value={textEl?.fontFamily ?? 'Plus Jakarta Sans'}
                onChange={(e) => onUpdateSelectedElement({fontFamily: e.target.value})}
                className="h-7 w-28 rounded border border-slate-200 bg-slate-50 px-1.5 text-xs text-slate-800 disabled:opacity-45 dark:border-slate-700 dark:bg-slate-950 dark:text-slate-100"
              >
                {fontOptions.map((f) => (
                  <option key={f} value={f}>
                    {f}
                  </option>
                ))}
              </select>
              <select
                aria-label="Text font size"
                disabled={!textEl}
                value={textEl?.fontSize ?? 20}
                onChange={(e) => onUpdateSelectedElement({fontSize: Number(e.target.value)})}
                className="h-7 w-14 rounded border border-slate-200 bg-slate-50 px-1 text-xs tabular-nums text-slate-800 disabled:opacity-45 dark:border-slate-700 dark:bg-slate-950 dark:text-slate-100"
              >
                {SLIDE_FONT_SIZES.map((sz) => (
                  <option key={sz} value={sz}>
                    {sz}
                  </option>
                ))}
              </select>
              <ToolBtn
                disabled={!textEl}
                active={Boolean(textEl?.bold)}
                onClick={() => onUpdateSelectedElement({bold: !textEl?.bold})}
                title="Bold"
              >
                <Bold className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                disabled={!textEl}
                active={Boolean(textEl?.italic)}
                onClick={() => onUpdateSelectedElement({italic: !textEl?.italic})}
                title="Italic"
              >
                <Italic className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                disabled={!textEl}
                active={Boolean(textEl?.underline)}
                onClick={() => onUpdateSelectedElement({underline: !textEl?.underline})}
                title="Underline"
              >
                <Underline className="h-3.5 w-3.5" />
              </ToolBtn>
              <label
                title="Text color"
                className={`inline-flex h-7 items-center gap-0.5 rounded border border-slate-200 px-1 dark:border-slate-700 ${
                  textEl ? 'cursor-pointer' : 'pointer-events-none opacity-45'
                }`}
              >
                <Type className="h-3 w-3 text-slate-600 dark:text-slate-300" />
                <input
                  type="color"
                  aria-label="Text color"
                  disabled={!textEl}
                  value={textEl?.color ?? '#0F172A'}
                  onChange={(e) => onUpdateSelectedElement({color: e.target.value})}
                  className="h-4 w-4 cursor-pointer border-0 bg-transparent p-0"
                />
              </label>
            </RibbonGroup>

            <RibbonGroup label="Paragraph">
              <ToolBtn
                disabled={!textEl}
                active={textEl?.align === 'left'}
                onClick={() => onUpdateSelectedElement({align: 'left'})}
                title="Align Left"
              >
                <AlignLeft className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                disabled={!textEl}
                active={textEl?.align === 'center'}
                onClick={() => onUpdateSelectedElement({align: 'center'})}
                title="Align Center"
              >
                <AlignCenter className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                disabled={!textEl}
                active={textEl?.align === 'right'}
                onClick={() => onUpdateSelectedElement({align: 'right'})}
                title="Align Right"
              >
                <AlignRight className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                disabled={!textEl}
                active={textEl?.listStyle === 'bullet'}
                onClick={() =>
                  onUpdateSelectedElement({
                    listStyle: textEl?.listStyle === 'bullet' ? 'none' : 'bullet',
                  })
                }
                title="Bullet List"
              >
                <List className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                disabled={!textEl}
                active={textEl?.listStyle === 'numbered'}
                onClick={() =>
                  onUpdateSelectedElement({
                    listStyle: textEl?.listStyle === 'numbered' ? 'none' : 'numbered',
                  })
                }
                title="Numbered List"
              >
                <ListOrdered className="h-3.5 w-3.5" />
              </ToolBtn>
            </RibbonGroup>

            <RibbonGroup label="Arrange">
              <ToolBtn onClick={onAddText} title="Insert Text Box" label="Text">
                <Type className="h-3.5 w-3.5 text-orange-600 dark:text-orange-400" />
              </ToolBtn>
              <ToolBtn
                disabled={!selectedElement}
                onClick={() => onLayerAction('front')}
                title="Bring element to front"
                label="Front"
              >
                <Layers className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                disabled={!selectedElement}
                onClick={() => onAlignElement('center')}
                title="Center horizontally on slide"
                label="Center"
              >
                <AlignCenter className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                disabled={!selectedElement}
                onClick={onDuplicateSelectedElement}
                title="Duplicate selected object (Ctrl+D)"
              >
                <Copy className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                disabled={!selectedElement}
                onClick={onDeleteSelectedElement}
                title="Delete selected object (Delete)"
              >
                <Trash2 className="h-3.5 w-3.5 text-rose-600 dark:text-rose-400" />
              </ToolBtn>
            </RibbonGroup>
          </>
        )}

        {activeTab === 'insert' && (
          <>
            <RibbonGroup label="Text & Media">
              <ToolBtn onClick={onAddText} title="Insert Text Box" label="Text Box">
                <Type className="h-3.5 w-3.5 text-orange-600 dark:text-orange-400" />
              </ToolBtn>
              <ToolBtn
                onClick={() => imageInputRef.current?.click()}
                title="Insert image from device"
                label="Image"
              >
                <ImageIcon className="h-3.5 w-3.5 text-sky-600 dark:text-sky-400" />
              </ToolBtn>
              <ToolBtn
                onClick={() => onOpenDialog('table')}
                title="Insert editable table"
                label="Table"
              >
                <Table2 className="h-3.5 w-3.5 text-emerald-600 dark:text-emerald-400" />
              </ToolBtn>
            </RibbonGroup>

            <RibbonGroup label="Shapes">
              <ToolBtn
                onClick={() => onAddShape('rounded-rect')}
                title="Insert rounded rectangle"
                label="Rect"
              >
                <Square className="h-3.5 w-3.5 text-blue-600 dark:text-blue-400" />
              </ToolBtn>
              <ToolBtn onClick={() => onAddShape('ellipse')} title="Insert circle or ellipse" label="Circle">
                <Circle className="h-3.5 w-3.5 text-indigo-600 dark:text-indigo-400" />
              </ToolBtn>
              <ToolBtn onClick={() => onAddShape('triangle')} title="Insert triangle" label="Triangle">
                <Triangle className="h-3.5 w-3.5 text-amber-600 dark:text-amber-400" />
              </ToolBtn>
              <ToolBtn onClick={() => onAddShape('arrow-right')} title="Insert right arrow" label="Arrow">
                <ArrowRight className="h-3.5 w-3.5 text-emerald-600 dark:text-emerald-400" />
              </ToolBtn>
              <ToolBtn onClick={() => onAddShape('line')} title="Insert divider line" label="Line">
                <Minus className="h-3.5 w-3.5" />
              </ToolBtn>
            </RibbonGroup>

            <RibbonGroup label="Charts">
              <ToolBtn
                onClick={() => onOpenDialog('chart', 'bar')}
                title="Insert Bar Chart"
                label="Bar Chart"
              >
                <BarChart3 className="h-3.5 w-3.5 text-orange-600 dark:text-orange-400" />
              </ToolBtn>
              <ToolBtn
                onClick={() => onOpenDialog('chart', 'line')}
                title="Insert Line Chart"
                label="Line Chart"
              >
                <BarChart3 className="h-3.5 w-3.5 text-sky-600 dark:text-sky-400" />
              </ToolBtn>
              <ToolBtn
                onClick={() => onOpenDialog('chart', 'pie')}
                title="Insert Pie Chart"
                label="Pie Chart"
              >
                <BarChart3 className="h-3.5 w-3.5 text-emerald-600 dark:text-emerald-400" />
              </ToolBtn>
            </RibbonGroup>

            <RibbonGroup label="Quick Builders">
              <ToolBtn
                onClick={() => onOpenDialog('outline')}
                title="Generate slides from structured outline text"
                label="Outline to Slides"
              >
                <Wand2 className="h-3.5 w-3.5 text-orange-600 dark:text-orange-400" />
              </ToolBtn>
              <ToolBtn
                onClick={() => onOpenDialog('templates')}
                title="Browse deck templates"
                label="Templates"
              >
                <LayoutTemplate className="h-3.5 w-3.5 text-blue-600 dark:text-blue-400" />
              </ToolBtn>
            </RibbonGroup>
          </>
        )}

        {activeTab === 'design' && (
          <>
            <RibbonGroup label="Themes">
              {SLIDE_THEMES.map((t) => {
                const active = themeId === t.id;
                return (
                  <button
                    key={t.id}
                    type="button"
                    onClick={() => onSelectTheme(t.id)}
                    className={`inline-flex h-7 shrink-0 items-center gap-1.5 whitespace-nowrap rounded border px-2 text-xs font-medium transition-colors ${
                      active
                        ? 'border-orange-500 bg-orange-50 text-orange-900 dark:bg-orange-950/60 dark:text-orange-200'
                        : 'border-slate-200 bg-white text-slate-700 hover:bg-slate-50 dark:border-slate-700 dark:bg-slate-950 dark:text-slate-200'
                    }`}
                  >
                    <span
                      className="h-3 w-3 rounded-sm border border-black/15"
                      style={{background: `linear-gradient(135deg, ${t.background} 50%, ${t.primary} 50%)`}}
                    />
                    {t.name}
                  </button>
                );
              })}
            </RibbonGroup>

            {onChangeAspectRatio && (
              <RibbonGroup label="Slide Size">
                <ToolBtn
                  active={aspectRatio === '16:9'}
                  onClick={() => onChangeAspectRatio('16:9')}
                  title="Widescreen 16:9 slide ratio"
                  label="16:9"
                >
                  <MonitorPlay className="h-3.5 w-3.5 text-orange-600 dark:text-orange-400" />
                </ToolBtn>
                <ToolBtn
                  active={aspectRatio === '4:3'}
                  onClick={() => onChangeAspectRatio('4:3')}
                  title="Standard 4:3 slide ratio"
                  label="4:3"
                >
                  <Square className="h-3.5 w-3.5 text-slate-600 dark:text-slate-300" />
                </ToolBtn>
              </RibbonGroup>
            )}

            <RibbonGroup label="Customize">
              <label
                title="Slide background color"
                className="inline-flex h-7 cursor-pointer items-center gap-1.5 rounded border border-slate-200 px-2 text-xs font-medium text-slate-700 dark:border-slate-700 dark:text-slate-200"
              >
                <Palette className="h-3.5 w-3.5 text-orange-600 dark:text-orange-400" />
                <span>Slide Fill</span>
                <input
                  type="color"
                  aria-label="Slide background fill"
                  value={slideBackground}
                  onChange={(e) => onChangeBackground(e.target.value)}
                  className="h-4 w-4 cursor-pointer border-0 bg-transparent p-0"
                />
              </label>
              <ToolBtn
                onClick={() => onOpenDialog('master')}
                title="Configure Slide Master footer, slide numbers, and accent bar"
                label="Slide Master"
              >
                <LayoutTemplate className="h-3.5 w-3.5" />
              </ToolBtn>
              <ToolBtn
                onClick={() => onOpenDialog('templates')}
                title="Start from a curated presentation template"
                label="Deck Templates"
              >
                <Sparkles className="h-3.5 w-3.5 text-orange-600 dark:text-orange-400" />
              </ToolBtn>
            </RibbonGroup>
          </>
        )}

        {activeTab === 'transitions' && (
          <>
            <RibbonGroup label="Slide Transition">
              {TRANSITIONS.map((tr) => (
                <ToolBtn
                  key={tr.id}
                  active={slideTransition === tr.id}
                  onClick={() => onChangeTransition(tr.id)}
                  title={`Set ${tr.label} transition for active slide`}
                  label={tr.label}
                >
                  <Sparkles className="h-3.5 w-3.5 text-orange-600 dark:text-orange-400" />
                </ToolBtn>
              ))}
            </RibbonGroup>

            <RibbonGroup label="Apply">
              <ToolBtn
                onClick={() => onApplyTransitionAll(slideTransition)}
                title="Apply current transition to all slides"
                label="Apply to All Slides"
              >
                <Layers className="h-3.5 w-3.5" />
              </ToolBtn>
            </RibbonGroup>
          </>
        )}

        {activeTab === 'slideshow' && (
          <>
            <RibbonGroup label="Start Slide Show">
              <ToolBtn
                onClick={() => onStartPresent(true)}
                title="Start presentation from Slide 1"
                label="From Beginning"
              >
                <MonitorPlay className="h-3.5 w-3.5 text-orange-600 dark:text-orange-400" />
              </ToolBtn>
              <ToolBtn
                onClick={() => onStartPresent(false)}
                title="Start presentation from current slide"
                label="From Current Slide"
              >
                <Play className="h-3.5 w-3.5 text-emerald-600 dark:text-emerald-400" />
              </ToolBtn>
            </RibbonGroup>

            <RibbonGroup label="Export & Handouts">
              <ToolBtn onClick={onExportPdf} title="Export widescreen PDF" label="Export PDF">
                <FileText className="h-3.5 w-3.5 text-rose-600 dark:text-rose-400" />
              </ToolBtn>
              <ToolBtn onClick={onExportPng} title="Save active slide as PNG" label="Export Slide PNG">
                <ImageIcon className="h-3.5 w-3.5 text-sky-600 dark:text-sky-400" />
              </ToolBtn>
              <ToolBtn onClick={() => onOpenDialog('export')} title="Save As PPTX or PDF" label="Export PPTX">
                <Download className="h-3.5 w-3.5" />
              </ToolBtn>
            </RibbonGroup>
          </>
        )}
      </div>
    </div>
  );
};
