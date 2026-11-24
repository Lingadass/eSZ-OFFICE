import React, {useState} from 'react';
import {
  AlertTriangle,
  ArrowDown,
  ArrowUp,
  Copy,
  Layers,
  Plus,
  Trash2,
  X,
} from 'lucide-react';
import {SlideSvgRenderer} from './SlidesCanvas';
import {fitTextToSlideBox, getSlideDimensions} from './slidesTextFit';
import type {
  SlideAspectRatio,
  SlideChartKind,
  SlideElement,
  SlideItem,
  SlideLayoutKind,
  SlideMasterSettings,
  SlideTransitionKind,
} from './slidesTypes';

interface SlideFilmstripProps {
  slides: SlideItem[];
  activeSlideId: string;
  themeId: string;
  aspectRatio?: SlideAspectRatio;
  master: SlideMasterSettings;
  onSelectSlide: (id: string) => void;
  onAddSlide: (layout: SlideLayoutKind) => void;
  onDuplicateSlide: (id: string) => void;
  onDeleteSlide: (id: string) => void;
  onReorderSlides: (fromIndex: number, toIndex: number) => void;
  onCloseMobile?: () => void;
  readOnly?: boolean;
}

export const SlideFilmstrip: React.FC<SlideFilmstripProps> = ({
  slides,
  activeSlideId,
  themeId,
  aspectRatio = '16:9',
  master,
  onSelectSlide,
  onAddSlide,
  onDuplicateSlide,
  onDeleteSlide,
  onReorderSlides,
  onCloseMobile,
  readOnly = false,
}) => {
  const [dragIdx, setDragIdx] = useState<number | null>(null);
  const dims = getSlideDimensions(aspectRatio);

  return (
    <aside
      aria-label="Slide thumbnails"
      data-testid="slides-filmstrip"
      className="flex h-full w-56 shrink-0 flex-col border-r border-slate-200 bg-slate-50 dark:border-slate-800 dark:bg-slate-900"
    >
      <div className="flex items-center justify-between border-b border-slate-200 px-3 py-2 dark:border-slate-800">
        <span className="text-xs font-semibold text-slate-700 dark:text-slate-200">
          Slides ({slides.length})
        </span>
        <div className="flex items-center gap-1">
          {!readOnly && (
            <button
              type="button"
              onClick={() => onAddSlide('title-content')}
              aria-label="Add new slide to filmstrip"
              className="inline-flex h-7 items-center gap-1 rounded-md bg-orange-600 px-2 text-xs font-semibold text-white hover:bg-orange-700"
            >
              <Plus className="h-3.5 w-3.5" />
              Slide
            </button>
          )}
          {onCloseMobile && (
            <button
              type="button"
              onClick={onCloseMobile}
              aria-label="Close slides panel"
              className="inline-flex h-7 w-7 items-center justify-center rounded-md text-slate-500 hover:bg-slate-200 dark:hover:bg-slate-800 md:hidden"
            >
              <X className="h-4 w-4" />
            </button>
          )}
        </div>
      </div>

      <div className="flex-1 space-y-2.5 overflow-y-auto p-2.5">
        {slides.map((slide, idx) => {
          const active = slide.id === activeSlideId;
          return (
            <div
              key={slide.id}
              draggable={!readOnly}
              onDragStart={() => {
                if (!readOnly) setDragIdx(idx);
              }}
              onDragOver={(e) => {
                if (!readOnly) e.preventDefault();
              }}
              onDrop={() => {
                if (!readOnly && dragIdx !== null && dragIdx !== idx) {
                  onReorderSlides(dragIdx, idx);
                }
                setDragIdx(null);
              }}
              onDragEnd={() => setDragIdx(null)}
              className={`group rounded-lg border p-1.5 transition-colors ${
                active
                  ? 'border-orange-500 bg-orange-50/70 dark:border-orange-500 dark:bg-orange-950/30'
                  : 'border-slate-200 bg-white hover:border-slate-300 dark:border-slate-800 dark:bg-slate-950 dark:hover:border-slate-700'
              }`}
            >
              <div className="mb-1 flex items-center justify-between text-[11px]">
                <span className="font-mono font-semibold tabular-nums text-slate-500 dark:text-slate-400">
                  {idx + 1}
                </span>
                {!readOnly && (
                  <div className="flex items-center gap-0.5 opacity-80 group-hover:opacity-100">
                    <button
                      type="button"
                      disabled={idx === 0}
                      onClick={() => onReorderSlides(idx, idx - 1)}
                      aria-label={`Move slide ${idx + 1} up`}
                      className="rounded p-0.5 text-slate-500 hover:bg-slate-200 disabled:opacity-30 dark:hover:bg-slate-800"
                    >
                      <ArrowUp className="h-3 w-3" />
                    </button>
                    <button
                      type="button"
                      disabled={idx === slides.length - 1}
                      onClick={() => onReorderSlides(idx, idx + 1)}
                      aria-label={`Move slide ${idx + 1} down`}
                      className="rounded p-0.5 text-slate-500 hover:bg-slate-200 disabled:opacity-30 dark:hover:bg-slate-800"
                    >
                      <ArrowDown className="h-3 w-3" />
                    </button>
                    <button
                      type="button"
                      onClick={() => onDuplicateSlide(slide.id)}
                      aria-label={`Duplicate slide ${idx + 1}`}
                      className="rounded p-0.5 text-slate-500 hover:bg-slate-200 dark:hover:bg-slate-800"
                    >
                      <Copy className="h-3 w-3" />
                    </button>
                    <button
                      type="button"
                      disabled={slides.length <= 1}
                      onClick={() => onDeleteSlide(slide.id)}
                      aria-label={`Delete slide ${idx + 1}`}
                      className="rounded p-0.5 text-slate-500 hover:bg-rose-100 hover:text-rose-600 disabled:opacity-30 dark:hover:bg-rose-950/60"
                    >
                      <Trash2 className="h-3 w-3" />
                    </button>
                  </div>
                )}
              </div>

              <button
                type="button"
                onClick={() => onSelectSlide(slide.id)}
                aria-label={`Select slide ${idx + 1}`}
                aria-current={active ? 'page' : undefined}
                className={`block ${dims.cssAspect} w-full overflow-hidden rounded border border-slate-200 bg-white dark:border-slate-700`}
              >
                <SlideSvgRenderer
                  slide={slide}
                  slideIndex={idx}
                  slideCount={slides.length}
                  themeId={themeId}
                  aspectRatio={aspectRatio}
                  master={master}
                />
              </button>
            </div>
          );
        })}
      </div>
    </aside>
  );
};

interface SlideInspectorProps {
  slide: SlideItem;
  selectedElement: SlideElement | null;
  aspectRatio?: SlideAspectRatio;
  master: SlideMasterSettings;
  onSelectElement: (id: string | null) => void;
  onUpdateElement: (id: string, patch: Partial<SlideElement>) => void;
  onDeleteElement: (id: string) => void;
  onLayerAction: (dir: 'front' | 'back') => void;
  onAlignElement: (align: 'left' | 'center' | 'right' | 'top' | 'middle' | 'bottom') => void;
  onChangeBackground: (hex: string) => void;
  onChangeTransition: (tr: SlideTransitionKind) => void;
  onUpdateMaster: (patch: Partial<SlideMasterSettings>) => void;
  onClose: () => void;
}

export const SlideInspector: React.FC<SlideInspectorProps> = ({
  slide,
  selectedElement,
  aspectRatio = '16:9',
  master,
  onSelectElement,
  onUpdateElement,
  onDeleteElement,
  onLayerAction,
  onAlignElement,
  onChangeBackground,
  onChangeTransition,
  onUpdateMaster,
  onClose,
}) => {
  const dims = getSlideDimensions(aspectRatio);
  const selectedTextFit =
    selectedElement?.type === 'text'
      ? fitTextToSlideBox({
          text: selectedElement.text,
          listStyle: selectedElement.listStyle,
          width: selectedElement.width,
          height: selectedElement.height,
          fontSize: selectedElement.fontSize,
          lineHeight: selectedElement.lineHeight ?? 1.35,
          padding: selectedElement.padding ?? 4,
          bold: selectedElement.bold,
          italic: selectedElement.italic,
          fontFamily: selectedElement.fontFamily,
          allowAutoShrink: selectedElement.autoShrink !== false,
        })
      : null;
  return (
    <aside
      aria-label="Slide and element inspector"
      className="flex h-full w-72 shrink-0 flex-col border-l border-slate-200 bg-white dark:border-slate-800 dark:bg-slate-900"
    >
      <div className="flex items-center justify-between border-b border-slate-200 px-3 py-2 dark:border-slate-800">
        <span className="text-xs font-semibold text-slate-800 dark:text-slate-100">
          {selectedElement ? `Format ${selectedElement.type}` : 'Slide Properties'}
        </span>
        <button
          type="button"
          onClick={onClose}
          aria-label="Close inspector"
          className="inline-flex h-7 w-7 items-center justify-center rounded-md text-slate-500 hover:bg-slate-100 dark:hover:bg-slate-800"
        >
          <X className="h-4 w-4" />
        </button>
      </div>

      <div className="flex-1 space-y-4 overflow-y-auto p-3 text-xs">
        {/* Accessible Element Layers list for quick keyboard/touch selection */}
        <div>
          <div className="mb-1.5 flex items-center justify-between text-[11px] font-semibold text-slate-500 dark:text-slate-400">
            <span className="inline-flex items-center gap-1">
              <Layers className="h-3.5 w-3.5" />
              Objects on Slide ({slide.elements.length})
            </span>
          </div>
          <div className="max-h-32 space-y-1 overflow-y-auto rounded-lg border border-slate-200 bg-slate-50 p-1.5 dark:border-slate-800 dark:bg-slate-950">
            {slide.elements.length === 0 ? (
              <p className="px-1.5 py-1 text-[11px] text-slate-400">No objects on this slide.</p>
            ) : (
              slide.elements.map((el, idx) => {
                const isSel = selectedElement?.id === el.id;
                const label =
                  el.type === 'text'
                    ? el.text.split('\n')[0].slice(0, 24) || 'Text Box'
                    : el.type === 'chart'
                      ? `Chart: ${el.title}`
                      : el.type === 'table'
                        ? `Table (${el.rows.length}×${el.rows[0]?.length ?? 0})`
                        : el.type === 'shape'
                          ? `Shape: ${el.shapeKind}`
                          : `Image: ${el.alt}`;
                return (
                  <button
                    key={el.id}
                    type="button"
                    onClick={() => onSelectElement(el.id)}
                    className={`flex w-full items-center justify-between rounded px-2 py-1 text-left text-[11px] font-medium transition-colors ${
                      isSel
                        ? 'bg-orange-100 text-orange-900 dark:bg-orange-950/70 dark:text-orange-200'
                        : 'text-slate-700 hover:bg-slate-200/70 dark:text-slate-300 dark:hover:bg-slate-800'
                    }`}
                  >
                    <span className="truncate">{label}</span>
                    <span className="ml-1 shrink-0 font-mono text-[10px] text-slate-400">
                      #{idx + 1}
                    </span>
                  </button>
                );
              })
            )}
          </div>
        </div>

        {!selectedElement ? (
          <div className="space-y-3 border-t border-slate-200 pt-3 dark:border-slate-800">
            <div>
              <label className="mb-1 block font-medium text-slate-600 dark:text-slate-300">
                Slide Background
              </label>
              <div className="flex items-center gap-2">
                <input
                  type="color"
                  aria-label="Inspector slide background"
                  value={slide.background}
                  onChange={(e) => onChangeBackground(e.target.value)}
                  className="h-8 w-10 cursor-pointer rounded border border-slate-300 bg-transparent dark:border-slate-700"
                />
                <span className="font-mono text-slate-600 dark:text-slate-300">
                  {slide.background.toUpperCase()}
                </span>
              </div>
            </div>

            <div>
              <label className="mb-1 block font-medium text-slate-600 dark:text-slate-300">
                Slide Transition
              </label>
              <select
                aria-label="Inspector slide transition"
                value={slide.transition}
                onChange={(e) => onChangeTransition(e.target.value as SlideTransitionKind)}
                className="w-full rounded-md border border-slate-300 bg-white px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-950"
              >
                <option value="none">None</option>
                <option value="fade">Fade</option>
                <option value="slide">Slide</option>
                <option value="zoom">Zoom</option>
              </select>
            </div>

            <div className="space-y-2 border-t border-slate-200 pt-3 dark:border-slate-800">
              <div className="font-semibold text-slate-700 dark:text-slate-200">Slide Master</div>
              <div>
                <label className="mb-1 block text-slate-600 dark:text-slate-300">Footer Text</label>
                <input
                  type="text"
                  value={master.footerText}
                  onChange={(e) => onUpdateMaster({footerText: e.target.value})}
                  placeholder="eSZ OFFICE"
                  className="w-full rounded-md border border-slate-300 bg-white px-2.5 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-950"
                />
              </div>
              <label className="flex cursor-pointer items-center gap-2 text-slate-700 dark:text-slate-300">
                <input
                  type="checkbox"
                  checked={master.showSlideNumbers}
                  onChange={(e) => onUpdateMaster({showSlideNumbers: e.target.checked})}
                />
                Show slide numbers
              </label>
              <label className="flex cursor-pointer items-center gap-2 text-slate-700 dark:text-slate-300">
                <input
                  type="checkbox"
                  checked={master.showAccentBar}
                  onChange={(e) => onUpdateMaster({showAccentBar: e.target.checked})}
                />
                Show top theme accent bar
              </label>
            </div>
          </div>
        ) : (
          <div className="space-y-3 border-t border-slate-200 pt-3 dark:border-slate-800">
            {/* Text Editor Controls */}
            {selectedElement.type === 'text' && (
              <div className="space-y-2">
                <label className="block font-medium text-slate-600 dark:text-slate-300">
                  Text Content
                </label>
                <textarea
                  aria-label="Selected text content"
                  rows={4}
                  value={selectedElement.text}
                  onChange={(e) =>
                    onUpdateElement(selectedElement.id, {text: e.target.value})
                  }
                  className="w-full rounded-md border border-slate-300 bg-white p-2 text-xs dark:border-slate-700 dark:bg-slate-950"
                />
                <div className="grid grid-cols-2 gap-2">
                  <div>
                    <label className="mb-0.5 block text-[11px] text-slate-500">Text Fill</label>
                    <input
                      type="color"
                      aria-label="Inspector text color"
                      value={selectedElement.color}
                      onChange={(e) =>
                        onUpdateElement(selectedElement.id, {color: e.target.value})
                      }
                      className="h-7 w-full cursor-pointer rounded border border-slate-300 dark:border-slate-700"
                    />
                  </div>
                  <div>
                    <label className="mb-0.5 block text-[11px] text-slate-500">Box Fill</label>
                    <input
                      type="color"
                      aria-label="Inspector text box background"
                      value={
                        selectedElement.fill && selectedElement.fill !== 'transparent'
                          ? selectedElement.fill
                          : '#FFFFFF'
                      }
                      onChange={(e) =>
                        onUpdateElement(selectedElement.id, {fill: e.target.value})
                      }
                      className="h-7 w-full cursor-pointer rounded border border-slate-300 dark:border-slate-700"
                    />
                  </div>
                </div>
                <label className="flex cursor-pointer items-center gap-2 pt-1 text-[11px] text-slate-700 dark:text-slate-300">
                  <input
                    type="checkbox"
                    checked={selectedElement.autoShrink !== false}
                    onChange={(e) =>
                      onUpdateElement(selectedElement.id, {autoShrink: e.target.checked})
                    }
                  />
                  Auto-shrink text to fit box
                </label>
                {selectedTextFit &&
                  (selectedTextFit.wasAutoShrunk ||
                    (selectedElement.originalFontSize &&
                      selectedElement.originalFontSize > selectedElement.fontSize)) &&
                  !selectedTextFit.stillOverflows && (
                    <div className="rounded-md border border-slate-200 bg-slate-50 px-2 py-1.5 text-[11px] text-slate-600 dark:border-slate-700 dark:bg-slate-950 dark:text-slate-300">
                      Text auto-fitted to{' '}
                      <span className="font-mono font-semibold">
                        {selectedTextFit.fittedFontSize}px
                      </span>{' '}
                      to fit cleanly inside its box.
                    </div>
                  )}
                {selectedTextFit?.stillOverflows && (
                  <div className="space-y-1.5 rounded-md border border-amber-300 bg-amber-50 p-2 text-[11px] text-amber-900 dark:border-amber-800 dark:bg-amber-950/50 dark:text-amber-200">
                    <div className="flex items-start gap-1.5">
                      <AlertTriangle className="mt-0.5 h-3.5 w-3.5 shrink-0 text-amber-600 dark:text-amber-400" />
                      <span>Text exceeds box height even at minimum font size.</span>
                    </div>
                    <button
                      type="button"
                      onClick={() =>
                        onUpdateElement(selectedElement.id, {
                          height: Math.min(
                            Math.max(selectedElement.height, selectedTextFit.recommendedHeight),
                            Math.max(selectedElement.height, dims.height - selectedElement.y - 12),
                          ),
                          hasOverflow: false,
                        })
                      }
                      className="w-full rounded bg-amber-600 px-2 py-1 text-[11px] font-semibold text-white hover:bg-amber-700"
                    >
                      Expand box to fit
                    </button>
                  </div>
                )}
              </div>
            )}

            {/* Shape Controls */}
            {selectedElement.type === 'shape' && (
              <div className="space-y-2">
                <div className="grid grid-cols-2 gap-2">
                  <div>
                    <label className="mb-0.5 block text-[11px] text-slate-500">Shape Fill</label>
                    <input
                      type="color"
                      aria-label="Shape fill color"
                      value={selectedElement.fill}
                      onChange={(e) =>
                        onUpdateElement(selectedElement.id, {fill: e.target.value})
                      }
                      className="h-7 w-full cursor-pointer rounded border border-slate-300 dark:border-slate-700"
                    />
                  </div>
                  <div>
                    <label className="mb-0.5 block text-[11px] text-slate-500">Border</label>
                    <input
                      type="color"
                      aria-label="Shape stroke color"
                      value={
                        selectedElement.stroke && selectedElement.stroke !== 'transparent'
                          ? selectedElement.stroke
                          : '#0F172A'
                      }
                      onChange={(e) =>
                        onUpdateElement(selectedElement.id, {
                          stroke: e.target.value,
                          strokeWidth: Math.max(2, selectedElement.strokeWidth || 2),
                        })
                      }
                      className="h-7 w-full cursor-pointer rounded border border-slate-300 dark:border-slate-700"
                    />
                  </div>
                </div>
                {selectedElement.shapeKind !== 'line' && (
                  <div>
                    <label className="mb-0.5 block text-[11px] text-slate-500">Shape Label</label>
                    <input
                      type="text"
                      value={selectedElement.label ?? ''}
                      onChange={(e) =>
                        onUpdateElement(selectedElement.id, {label: e.target.value})
                      }
                      placeholder="Optional centered label"
                      className="w-full rounded-md border border-slate-300 bg-white px-2 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-950"
                    />
                  </div>
                )}
              </div>
            )}

            {/* Table Editor Controls */}
            {selectedElement.type === 'table' && (
              <div className="space-y-2">
                <div className="flex items-center justify-between">
                  <span className="font-medium text-slate-700 dark:text-slate-200">
                    Table Cells
                  </span>
                  <div className="flex gap-1">
                    <button
                      type="button"
                      onClick={() => {
                        const cols = selectedElement.rows[0]?.length ?? 3;
                        onUpdateElement(selectedElement.id, {
                          rows: [
                            ...selectedElement.rows,
                            Array.from({length: cols}, (_, c) => `Cell ${c + 1}`),
                          ],
                        });
                      }}
                      className="rounded border border-slate-300 px-1.5 py-0.5 text-[11px] hover:bg-slate-100 dark:border-slate-700 dark:hover:bg-slate-800"
                    >
                      +Row
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        onUpdateElement(selectedElement.id, {
                          rows: selectedElement.rows.map((r, rIdx) => [
                            ...r,
                            rIdx === 0 ? `Col ${r.length + 1}` : '',
                          ]),
                        });
                      }}
                      className="rounded border border-slate-300 px-1.5 py-0.5 text-[11px] hover:bg-slate-100 dark:border-slate-700 dark:hover:bg-slate-800"
                    >
                      +Col
                    </button>
                  </div>
                </div>
                <div className="max-h-44 space-y-1.5 overflow-y-auto">
                  {selectedElement.rows.map((row, rIdx) => (
                    <div key={rIdx} className="grid gap-1" style={{gridTemplateColumns: `repeat(${row.length}, minmax(0, 1fr))`}}>
                      {row.map((cell, cIdx) => (
                        <input
                          key={cIdx}
                          type="text"
                          aria-label={`Table row ${rIdx + 1} column ${cIdx + 1}`}
                          value={cell}
                          onChange={(e) => {
                            const nextRows = selectedElement.rows.map((r, ri) =>
                              ri === rIdx
                                ? r.map((v, ci) => (ci === cIdx ? e.target.value : v))
                                : r,
                            );
                            onUpdateElement(selectedElement.id, {rows: nextRows});
                          }}
                          className="rounded border border-slate-300 bg-white px-1.5 py-1 text-[11px] dark:border-slate-700 dark:bg-slate-950"
                        />
                      ))}
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* Chart Editor Controls */}
            {selectedElement.type === 'chart' && (
              <div className="space-y-2">
                <div>
                  <label className="mb-0.5 block text-[11px] text-slate-500">Chart Title</label>
                  <input
                    type="text"
                    value={selectedElement.title}
                    onChange={(e) =>
                      onUpdateElement(selectedElement.id, {title: e.target.value})
                    }
                    className="w-full rounded-md border border-slate-300 bg-white px-2 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-950"
                  />
                </div>
                <div>
                  <label className="mb-0.5 block text-[11px] text-slate-500">Chart Type</label>
                  <select
                    aria-label="Selected chart type"
                    value={selectedElement.chartKind}
                    onChange={(e) =>
                      onUpdateElement(selectedElement.id, {
                        chartKind: e.target.value as SlideChartKind,
                      })
                    }
                    className="w-full rounded-md border border-slate-300 bg-white px-2 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-950"
                  >
                    <option value="bar">Bar Chart</option>
                    <option value="line">Line Chart</option>
                    <option value="pie">Pie Chart</option>
                  </select>
                </div>
                <div className="space-y-1">
                  <div className="text-[11px] font-medium text-slate-500">Data Points</div>
                  {selectedElement.labels.map((lbl, idx) => (
                    <div key={idx} className="grid grid-cols-2 gap-1.5">
                      <input
                        type="text"
                        aria-label={`Chart label ${idx + 1}`}
                        value={lbl}
                        onChange={(e) => {
                          const nextLabels = [...selectedElement.labels];
                          nextLabels[idx] = e.target.value;
                          onUpdateElement(selectedElement.id, {labels: nextLabels});
                        }}
                        className="rounded border border-slate-300 bg-white px-1.5 py-1 text-[11px] dark:border-slate-700 dark:bg-slate-950"
                      />
                      <input
                        type="number"
                        aria-label={`Chart value ${idx + 1}`}
                        value={selectedElement.values[idx] ?? 0}
                        onChange={(e) => {
                          const nextVals = [...selectedElement.values];
                          nextVals[idx] = Number(e.target.value) || 0;
                          onUpdateElement(selectedElement.id, {values: nextVals});
                        }}
                        className="rounded border border-slate-300 bg-white px-1.5 py-1 text-[11px] tabular-nums dark:border-slate-700 dark:bg-slate-950"
                      />
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* Geometry & Arrange */}
            <div className="space-y-2 border-t border-slate-200 pt-3 dark:border-slate-800">
              <div className="font-semibold text-slate-700 dark:text-slate-200">
                Position & Size
              </div>
              <div className="grid grid-cols-2 gap-1.5">
                {(
                  [
                    ['X', 'x', selectedElement.x],
                    ['Y', 'y', selectedElement.y],
                    ['W', 'width', selectedElement.width],
                    ['H', 'height', selectedElement.height],
                  ] as const
                ).map(([lbl, key, val]) => (
                  <label key={key} className="flex items-center gap-1.5 rounded border border-slate-200 px-2 py-1 dark:border-slate-800">
                    <span className="text-[10px] font-semibold text-slate-400">{lbl}</span>
                    <input
                      type="number"
                      aria-label={`Element ${lbl}`}
                      value={Math.round(val)}
                      onChange={(e) =>
                        onUpdateElement(selectedElement.id, {
                          [key]: Math.max(key === 'width' || key === 'height' ? 20 : 0, Number(e.target.value) || 0),
                        })
                      }
                      className="w-full bg-transparent text-right font-mono text-xs tabular-nums outline-none"
                    />
                  </label>
                ))}
              </div>

              <div className="grid grid-cols-3 gap-1">
                {(
                  [
                    ['left', 'Left'],
                    ['center', 'Center'],
                    ['right', 'Right'],
                    ['top', 'Top'],
                    ['middle', 'Middle'],
                    ['bottom', 'Bottom'],
                  ] as const
                ).map(([al, label]) => (
                  <button
                    key={al}
                    type="button"
                    onClick={() => onAlignElement(al)}
                    className="rounded border border-slate-200 py-1 text-[11px] font-medium text-slate-700 hover:bg-slate-100 dark:border-slate-800 dark:text-slate-300 dark:hover:bg-slate-800"
                  >
                    {label}
                  </button>
                ))}
              </div>

              <div className="flex gap-1.5">
                <button
                  type="button"
                  onClick={() => onLayerAction('front')}
                  className="flex-1 rounded border border-slate-200 py-1 text-[11px] font-medium text-slate-700 hover:bg-slate-100 dark:border-slate-800 dark:text-slate-300 dark:hover:bg-slate-800"
                >
                  Bring Front
                </button>
                <button
                  type="button"
                  onClick={() => onLayerAction('back')}
                  className="flex-1 rounded border border-slate-200 py-1 text-[11px] font-medium text-slate-700 hover:bg-slate-100 dark:border-slate-800 dark:text-slate-300 dark:hover:bg-slate-800"
                >
                  Send Back
                </button>
              </div>

              <button
                type="button"
                onClick={() => onDeleteElement(selectedElement.id)}
                className="flex w-full items-center justify-center gap-1.5 rounded-md border border-rose-200 bg-rose-50 py-1.5 text-xs font-semibold text-rose-700 hover:bg-rose-100 dark:border-rose-900/60 dark:bg-rose-950/40 dark:text-rose-300"
              >
                <Trash2 className="h-3.5 w-3.5" />
                Delete Object
              </button>
            </div>
          </div>
        )}
      </div>
    </aside>
  );
};
