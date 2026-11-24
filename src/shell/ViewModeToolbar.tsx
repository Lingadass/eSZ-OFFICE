import React from 'react';
import {
  Download,
  Pencil,
  Printer,
  Search,
  Star,
  ZoomIn,
  ZoomOut,
} from 'lucide-react';
import type {OfficeModuleType} from '../core/types';

export interface ViewModeToolbarProps {
  fileName: string;
  moduleType: OfficeModuleType;
  zoomLabel: string;
  onZoomOut: () => void;
  onZoomIn: () => void;
  onZoomReset?: () => void;
  isFavorite: boolean;
  onToggleFavorite: () => void;
  onSearch?: () => void;
  searchActive?: boolean;
  onPrint?: () => void;
  onExport?: () => void;
  exportLabel?: string;
  onEdit: () => void;
  leadingSlot?: React.ReactNode;
  navigationSlot?: React.ReactNode;
  extraActionsSlot?: React.ReactNode;
}

const EDIT_CTA_CLASS: Record<OfficeModuleType, string> = {
  pdf: 'bg-blue-700 hover:bg-blue-800 dark:bg-blue-600 dark:hover:bg-blue-500',
  docs: 'bg-blue-700 hover:bg-blue-800 dark:bg-blue-600 dark:hover:bg-blue-500',
  sheets: 'bg-emerald-700 hover:bg-emerald-800 dark:bg-emerald-600 dark:hover:bg-emerald-500',
  slides: 'bg-orange-600 hover:bg-orange-700 dark:bg-orange-600 dark:hover:bg-orange-500',
};

/**
 * Slim, distraction-free View Mode toolbar shared across all four office modules.
 * Displays file name, zoom, module navigation slot, search, print, export/download,
 * favourite star, and a prominent "Edit" CTA.
 */
export const ViewModeToolbar: React.FC<ViewModeToolbarProps> = ({
  fileName,
  moduleType,
  zoomLabel,
  onZoomOut,
  onZoomIn,
  onZoomReset,
  isFavorite,
  onToggleFavorite,
  onSearch,
  searchActive = false,
  onPrint,
  onExport,
  exportLabel = 'Download / Export',
  onEdit,
  leadingSlot,
  navigationSlot,
  extraActionsSlot,
}) => {
  return (
    <div
      data-testid="view-mode-toolbar"
      className="flex shrink-0 items-center justify-between gap-2 overflow-x-auto border-b border-slate-200 bg-white px-3 py-1.5 dark:border-slate-800 dark:bg-slate-900"
    >
      {/* Left: optional sidebar toggle, favourite star, and file name */}
      <div className="flex min-w-0 items-center gap-1.5">
        {leadingSlot}
        <button
          type="button"
          onClick={onToggleFavorite}
          aria-label={isFavorite ? 'Remove from favourites' : 'Add to favourites'}
          title={isFavorite ? 'Remove from favourites' : 'Add to favourites'}
          className="inline-flex h-8 w-8 shrink-0 items-center justify-center rounded-lg border border-slate-200 text-slate-600 transition-colors hover:bg-slate-100 dark:border-slate-700 dark:text-slate-300 dark:hover:bg-slate-800"
        >
          <Star
            className={`h-3.5 w-3.5 ${
              isFavorite ? 'fill-amber-400 text-amber-500' : ''
            }`}
          />
        </button>
        <span
          title={fileName}
          className="max-w-[160px] truncate text-xs font-semibold text-slate-900 sm:max-w-[240px] lg:max-w-[300px] dark:text-slate-100"
        >
          {fileName}
        </span>
      </div>

      {/* Center: module navigation (pages / slides / sheets) + zoom controls */}
      <div className="flex shrink-0 items-center gap-1.5">
        {navigationSlot}

        <div className="flex shrink-0 items-center gap-0.5 rounded-lg border border-slate-200 bg-slate-50 p-0.5 dark:border-slate-700 dark:bg-slate-950">
          <button
            type="button"
            onClick={onZoomOut}
            aria-label="Zoom out"
            title="Zoom out"
            className="flex h-7 w-7 items-center justify-center rounded-md text-slate-700 hover:bg-slate-200/70 dark:text-slate-200 dark:hover:bg-slate-800"
          >
            <ZoomOut className="h-3.5 w-3.5" />
          </button>
          <button
            type="button"
            onClick={onZoomReset}
            disabled={!onZoomReset}
            title={onZoomReset ? 'Reset / toggle zoom' : undefined}
            className="min-h-[28px] whitespace-nowrap rounded-md px-1.5 text-xs font-semibold tabular-nums text-slate-700 hover:bg-slate-200/70 disabled:hover:bg-transparent dark:text-slate-200 dark:hover:bg-slate-800"
          >
            {zoomLabel}
          </button>
          <button
            type="button"
            onClick={onZoomIn}
            aria-label="Zoom in"
            title="Zoom in"
            className="flex h-7 w-7 items-center justify-center rounded-md text-slate-700 hover:bg-slate-200/70 dark:text-slate-200 dark:hover:bg-slate-800"
          >
            <ZoomIn className="h-3.5 w-3.5" />
          </button>
        </div>
      </div>

      {/* Right: Search, module extras, Print, Export, and desktop Edit CTA */}
      <div className="flex shrink-0 items-center gap-1.5">
        {extraActionsSlot}

        {onSearch && (
          <button
            type="button"
            onClick={onSearch}
            aria-label="Search in document"
            title="Search (Ctrl+F)"
            className={`inline-flex h-8 w-8 shrink-0 items-center justify-center rounded-lg border transition-colors ${
              searchActive
                ? 'border-blue-600 bg-blue-50 text-blue-700 dark:border-blue-500 dark:bg-blue-950/50 dark:text-blue-300'
                : 'border-slate-200 text-slate-700 hover:bg-slate-100 dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800'
            }`}
          >
            <Search className="h-3.5 w-3.5" />
          </button>
        )}

        {onPrint && (
          <button
            type="button"
            onClick={onPrint}
            aria-label="Print document"
            title="Print"
            className="hidden h-8 w-8 shrink-0 items-center justify-center rounded-lg border border-slate-200 text-slate-700 hover:bg-slate-100 sm:inline-flex dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800"
          >
            <Printer className="h-3.5 w-3.5" />
          </button>
        )}

        {onExport && (
          <button
            type="button"
            onClick={onExport}
            aria-label={exportLabel}
            title={exportLabel}
            className="inline-flex h-8 w-8 shrink-0 items-center justify-center rounded-lg border border-slate-200 text-slate-700 hover:bg-slate-100 dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800"
          >
            <Download className="h-3.5 w-3.5" />
          </button>
        )}

        <button
          type="button"
          onClick={onEdit}
          aria-label="Edit"
          title="Switch to Edit mode (E or Ctrl+E)"
          className={`hidden min-h-[34px] shrink-0 items-center gap-1.5 whitespace-nowrap rounded-lg px-3.5 text-xs font-semibold text-white shadow-xs transition-colors md:inline-flex ${EDIT_CTA_CLASS[moduleType]}`}
        >
          <Pencil className="h-3.5 w-3.5" />
          <span>Edit</span>
        </button>
      </div>
    </div>
  );
};

interface ViewModeMobileBarProps {
  moduleType: OfficeModuleType;
  onEdit: () => void;
  onSearch?: () => void;
  onExport?: () => void;
  secondarySlot?: React.ReactNode;
}

/**
 * Mobile bottom action bar shown in View Mode on screens < md (768px),
 * providing thumb-friendly access to the prominent Edit CTA without colliding
 * with the global mobile navigation bar.
 */
export const ViewModeMobileBar: React.FC<ViewModeMobileBarProps> = ({
  moduleType,
  onEdit,
  onSearch,
  onExport,
  secondarySlot,
}) => {
  return (
    <div
      data-testid="view-mode-mobile-bar"
      className="flex shrink-0 items-center justify-between gap-2 border-t border-slate-200 bg-white px-3 py-2 md:hidden dark:border-slate-800 dark:bg-slate-900"
    >
      <div className="flex min-w-0 items-center gap-1.5">
        {secondarySlot}
        {onSearch && (
          <button
            type="button"
            onClick={onSearch}
            aria-label="Search document"
            className="inline-flex h-10 w-10 shrink-0 items-center justify-center rounded-xl border border-slate-200 text-slate-700 hover:bg-slate-100 dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800"
          >
            <Search className="h-4 w-4" />
          </button>
        )}
        {onExport && (
          <button
            type="button"
            onClick={onExport}
            aria-label="Export document"
            className="inline-flex h-10 w-10 shrink-0 items-center justify-center rounded-xl border border-slate-200 text-slate-700 hover:bg-slate-100 dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800"
          >
            <Download className="h-4 w-4" />
          </button>
        )}
      </div>

      <button
        type="button"
        onClick={onEdit}
        aria-label="Edit document"
        className={`inline-flex min-h-[42px] shrink-0 items-center justify-center gap-2 whitespace-nowrap rounded-xl px-5 text-xs font-semibold text-white shadow-xs transition-colors ${EDIT_CTA_CLASS[moduleType]}`}
      >
        <Pencil className="h-4 w-4" />
        <span>Edit</span>
      </button>
    </div>
  );
};
