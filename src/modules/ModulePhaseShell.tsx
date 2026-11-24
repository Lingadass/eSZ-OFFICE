import React from 'react';
import {
  Download,
  FileSpreadsheet,
  FileText,
  Presentation,
  Redo2,
  Save,
  Star,
  Undo2,
} from 'lucide-react';
import {MODULE_METADATA} from '../core/storage/fileRules';
import type {OfficeFileRecord, WorkspaceTab} from '../core/types';

interface ModulePhaseShellProps {
  tab: WorkspaceTab;
  fileRecord?: OfficeFileRecord;
  canUndo: boolean;
  canRedo: boolean;
  onChangeContent: (next: string) => void;
  onUndo: () => void;
  onRedo: () => void;
  onSaveNow: () => void;
  onExportDevice: () => void;
  onToggleFavorite: () => void;
  onDismissRecovery: () => void;
  plannedHighlights: string[];
  honestNotice: string;
}

function getModuleIcon(type: WorkspaceTab['moduleType']) {
  switch (type) {
    case 'pdf':
      return <FileText className="h-5 w-5 text-red-600 dark:text-red-400" />;
    case 'docs':
      return <FileText className="h-5 w-5 text-blue-600 dark:text-blue-400" />;
    case 'sheets':
      return <FileSpreadsheet className="h-5 w-5 text-emerald-700 dark:text-emerald-400" />;
    case 'slides':
      return <Presentation className="h-5 w-5 text-amber-600 dark:text-amber-400" />;
  }
}

export const ModulePhaseShell: React.FC<ModulePhaseShellProps> = ({
  tab,
  fileRecord,
  canUndo,
  canRedo,
  onChangeContent,
  onUndo,
  onRedo,
  onSaveNow,
  onExportDevice,
  onToggleFavorite,
  onDismissRecovery,
  plannedHighlights,
  honestNotice,
}) => {
  const meta = MODULE_METADATA[tab.moduleType];

  return (
    <div className="flex h-full flex-col bg-white dark:bg-slate-900">
      {/* Crash recovery banner */}
      {tab.recoveredFromCrash && (
        <div
          role="status"
          className="flex flex-wrap items-center justify-between gap-3 border-b border-amber-200 bg-amber-50 px-4 py-2.5 text-sm text-amber-900 dark:border-amber-900/50 dark:bg-amber-950/40 dark:text-amber-200"
        >
          <span>
            Recovered unsaved edits from your previous session. Autosave will persist your changes.
          </span>
          <button
            type="button"
            onClick={onDismissRecovery}
            className="min-h-[36px] rounded-lg border border-amber-300 bg-white px-3 py-1 text-xs font-medium text-amber-900 hover:bg-amber-100 dark:border-amber-800 dark:bg-amber-900/50 dark:text-amber-100"
          >
            Dismiss
          </button>
        </div>
      )}

      {/* Workspace toolbar */}
      <div className="flex flex-wrap items-center justify-between gap-3 border-b border-slate-200 px-4 py-2.5 dark:border-slate-800">
        <div className="flex min-w-0 items-center gap-3">
          <div className="flex h-9 w-9 shrink-0 items-center justify-center rounded-lg bg-slate-100 dark:bg-slate-800">
            {getModuleIcon(tab.moduleType)}
          </div>
          <div className="min-w-0">
            <h2 className="truncate text-sm font-semibold text-slate-900 dark:text-slate-100">
              {tab.name}
            </h2>
            <p className="text-xs text-slate-500 tabular-nums dark:text-slate-400">
              {meta.label} · {tab.dirty ? 'Unsaved changes' : 'Saved on device'} ·{' '}
              {new Date(tab.lastSavedAt).toLocaleTimeString('en-US', {
                hour: '2-digit',
                minute: '2-digit',
                second: '2-digit',
              })}
            </p>
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-1.5">
          <button
            type="button"
            onClick={onUndo}
            disabled={!canUndo}
            aria-label="Undo"
            title="Undo (Ctrl+Z)"
            className="inline-flex min-h-[40px] min-w-[40px] items-center justify-center rounded-lg border border-slate-200 bg-white px-2.5 text-slate-700 transition-colors hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-40 dark:border-slate-700 dark:bg-slate-800 dark:text-slate-200 dark:hover:bg-slate-700"
          >
            <Undo2 className="h-4 w-4" />
          </button>
          <button
            type="button"
            onClick={onRedo}
            disabled={!canRedo}
            aria-label="Redo"
            title="Redo (Ctrl+Shift+Z)"
            className="inline-flex min-h-[40px] min-w-[40px] items-center justify-center rounded-lg border border-slate-200 bg-white px-2.5 text-slate-700 transition-colors hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-40 dark:border-slate-700 dark:bg-slate-800 dark:text-slate-200 dark:hover:bg-slate-700"
          >
            <Redo2 className="h-4 w-4" />
          </button>
          <button
            type="button"
            onClick={onToggleFavorite}
            aria-label={fileRecord?.isFavorite ? 'Remove from favourites' : 'Add to favourites'}
            className="inline-flex min-h-[40px] min-w-[40px] items-center justify-center rounded-lg border border-slate-200 bg-white px-2.5 text-slate-700 transition-colors hover:bg-slate-50 dark:border-slate-700 dark:bg-slate-800 dark:text-slate-200 dark:hover:bg-slate-700"
          >
            <Star
              className={`h-4 w-4 ${
                fileRecord?.isFavorite
                  ? 'fill-amber-400 text-amber-500'
                  : 'text-slate-500 dark:text-slate-400'
              }`}
            />
          </button>
          <button
            type="button"
            onClick={onSaveNow}
            className="inline-flex min-h-[40px] items-center gap-2 whitespace-nowrap rounded-lg bg-blue-700 px-3.5 py-2 text-xs font-semibold text-white transition-colors hover:bg-blue-800 dark:bg-blue-600 dark:hover:bg-blue-500"
          >
            <Save className="h-4 w-4" />
            Save
          </button>
          <button
            type="button"
            onClick={onExportDevice}
            className="inline-flex min-h-[40px] items-center gap-2 whitespace-nowrap rounded-lg border border-slate-200 bg-white px-3.5 py-2 text-xs font-medium text-slate-700 transition-colors hover:bg-slate-50 dark:border-slate-700 dark:bg-slate-800 dark:text-slate-200 dark:hover:bg-slate-700"
          >
            <Download className="h-4 w-4" />
            Export
          </button>
        </div>
      </div>

      {/* Main module body: live editable draft buffer + Phase roadmap notice */}
      <div className="grid flex-1 grid-cols-1 gap-6 overflow-y-auto p-4 lg:grid-cols-3 lg:p-6">
        <div className="flex flex-col lg:col-span-2">
          <label
            htmlFor={`editor-${tab.fileId}`}
            className="mb-2 text-xs font-medium text-slate-600 dark:text-slate-400"
          >
            Local Document Buffer (Autosaves every 5 seconds to {fileRecord?.storageBackend === 'opfs' ? 'OPFS' : 'IndexedDB'})
          </label>
          <textarea
            id={`editor-${tab.fileId}`}
            value={tab.contentText}
            onChange={(e) => onChangeContent(e.target.value)}
            rows={16}
            spellCheck
            aria-label={`Document content for ${tab.name}`}
            className="min-h-[360px] w-full flex-1 resize-y rounded-xl border border-slate-200 bg-slate-50 p-4 font-mono text-sm leading-relaxed text-slate-900 focus:border-blue-600 focus:bg-white focus:outline-none dark:border-slate-800 dark:bg-slate-950 dark:text-slate-100 dark:focus:border-blue-500 dark:focus:bg-slate-900"
          />
        </div>

        <aside className="flex flex-col gap-5 rounded-xl border border-slate-200 bg-slate-50 p-5 dark:border-slate-800 dark:bg-slate-950">
          <div>
            <h3 className="text-sm font-semibold text-slate-900 dark:text-slate-100">
              {meta.label} Engine Status
            </h3>
            <p className="mt-1.5 text-sm leading-relaxed text-slate-600 dark:text-slate-300">
              {meta.phaseNote}
            </p>
          </div>

          <div className="border-t border-slate-200 pt-4 dark:border-slate-800">
            <h4 className="text-xs font-semibold tracking-wide text-slate-700 dark:text-slate-300">
              Included Capabilities
            </h4>
            <ul className="mt-2.5 space-y-2 text-xs text-slate-600 dark:text-slate-400">
              {plannedHighlights.map((item) => (
                <li key={item} className="flex items-start gap-2">
                  <span aria-hidden="true" className="mt-0.5 text-blue-600 dark:text-blue-400">
                    •
                  </span>
                  <span>{item}</span>
                </li>
              ))}
            </ul>
          </div>

          <div className="border-t border-slate-200 pt-4 dark:border-slate-800">
            <h4 className="text-xs font-semibold text-slate-700 dark:text-slate-300">
              Format Compatibility Notice
            </h4>
            <p className="mt-1.5 text-xs leading-relaxed text-slate-500 dark:text-slate-400">
              {honestNotice}
            </p>
          </div>
        </aside>
      </div>
    </div>
  );
};
