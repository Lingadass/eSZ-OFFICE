import React, {Suspense, useCallback, useEffect, useState} from 'react';
import {useNavigate} from 'react-router-dom';
import {FilePlus2, FolderOpen, Plus, X} from 'lucide-react';
import {pickLocalFiles, saveBlobToDevice} from '../core/file-io/fileSystemAccess';
import {loadFileBlob} from '../core/storage/fileRepository';
import {MODULE_METADATA} from '../core/storage/fileRules';
import type {OfficeModuleType} from '../core/types';
import {useWorkspaceStore} from '../core/workspace/useWorkspaceStore';
import {
  resolveModeShortcut,
  shouldConfirmExitEdit,
} from '../core/workspace/viewModeRules';
import {ModuleIcon} from './FileThumb';

const PdfModule = React.lazy(() => import('../modules/pdf/PdfModule'));
const DocsModule = React.lazy(() => import('../modules/docs/DocsModule'));
const SheetsModule = React.lazy(() => import('../modules/sheets/SheetsModule'));
const SlidesModule = React.lazy(() => import('../modules/slides/SlidesModule'));

export const WorkspaceView: React.FC = () => {
  const navigate = useNavigate();
  const {
    tabs,
    activeTabId,
    files,
    recoveryQueue,
    searchOpen,
    setActiveTab,
    setTabMode,
    discardTabChanges,
    closeTab,
    openFileInTab,
    createAndOpenFile,
    importFiles,
    updateTabContent,
    undoTab,
    redoTab,
    canUndoTab,
    canRedoTab,
    saveTabNow,
    autosaveDirtyTabs,
    toggleFavoriteById,
    dismissCrashRecoveryBanner,
    discardRecoverySnapshot,
    pushToast,
  } = useWorkspaceStore();

  const [exitConfirmFileId, setExitConfirmFileId] = useState<string | null>(null);
  const [savingExit, setSavingExit] = useState(false);

  // 5-second autosave loop (store skips view-mode tabs automatically)
  useEffect(() => {
    const timer = window.setInterval(() => {
      void autosaveDirtyTabs();
    }, 5000);
    return () => window.clearInterval(timer);
  }, [autosaveDirtyTabs]);

  const activeTab = tabs.find((t) => t.fileId === activeTabId) ?? null;
  const activeRecord = activeTab ? files.find((f) => f.id === activeTab.fileId) : undefined;
  const confirmTab = exitConfirmFileId
    ? tabs.find((t) => t.fileId === exitConfirmFileId) ?? null
    : null;

  const handleEnterEdit = useCallback(
    (fileId: string) => {
      setTabMode(fileId, 'edit');
    },
    [setTabMode],
  );

  const handleDoneEditing = useCallback(
    async (fileId: string) => {
      const target = useWorkspaceStore.getState().tabs.find((t) => t.fileId === fileId);
      if (!target) return;
      if (target.dirty && target.mode === 'edit') {
        setSavingExit(true);
        try {
          await saveTabNow(fileId);
        } finally {
          setSavingExit(false);
        }
      }
      setTabMode(fileId, 'view');
      setExitConfirmFileId(null);
    },
    [saveTabNow, setTabMode],
  );

  const handleRequestExitEdit = useCallback(
    (fileId: string) => {
      const target = useWorkspaceStore.getState().tabs.find((t) => t.fileId === fileId);
      if (!target) return;
      if (shouldConfirmExitEdit(target.mode, target.dirty)) {
        setExitConfirmFileId(fileId);
      } else {
        setTabMode(fileId, 'view');
      }
    },
    [setTabMode],
  );

  const handleDiscardExitEdit = useCallback(
    async (fileId: string) => {
      setSavingExit(true);
      try {
        await discardTabChanges(fileId);
        setExitConfirmFileId(null);
        pushToast('Discarded unsaved changes', 'info');
      } finally {
        setSavingExit(false);
      }
    },
    [discardTabChanges, pushToast],
  );

  // Global keyboard shortcuts for View/Edit mode: E, Ctrl+E, Esc
  useEffect(() => {
    if (!activeTab) return;
    const onKeyDown = (e: KeyboardEvent) => {
      const target = e.target as HTMLElement | null;
      const isProseMirrorTarget = Boolean(target?.closest?.('.ProseMirror'));
      if (e.defaultPrevented && !(e.key === 'Escape' && isProseMirrorTarget)) return;

      const isEditableTarget = Boolean(
        target &&
          (target.tagName === 'INPUT' ||
            target.tagName === 'TEXTAREA' ||
            target.tagName === 'SELECT' ||
            target.isContentEditable),
      );
      const hasBlockingOverlay =
        searchOpen ||
        Boolean(exitConfirmFileId) ||
        Boolean(document.querySelector('[role="dialog"], [aria-modal="true"]'));

      const action = resolveModeShortcut({
        key: e.key,
        ctrlOrMeta: e.ctrlKey || e.metaKey,
        alt: e.altKey,
        shift: e.shiftKey,
        mode: activeTab.mode,
        isEditableTarget,
        hasBlockingOverlay,
      });

      if (!action) return;
      e.preventDefault();
      if (action === 'enter-edit') {
        handleEnterEdit(activeTab.fileId);
      } else if (action === 'request-exit-edit') {
        handleRequestExitEdit(activeTab.fileId);
      }
    };

    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, [activeTab, searchOpen, exitConfirmFileId, handleEnterEdit, handleRequestExitEdit]);

  const handleExportDevice = async () => {
    if (!activeTab) return;
    if (activeTab.mode === 'edit' && activeTab.dirty) {
      await saveTabNow(activeTab.fileId);
    }
    const blob =
      (await loadFileBlob(activeTab.fileId)) ??
      new Blob([activeTab.contentText], {
        type: MODULE_METADATA[activeTab.moduleType].mimeType,
      });
    const result = await saveBlobToDevice(blob, activeTab.name, activeTab.moduleType);
    if (result === 'fs-access') {
      pushToast(`Saved ${activeTab.name} to device`, 'success');
    } else if (result === 'download') {
      pushToast(`Downloaded ${activeTab.name}`, 'success');
    }
  };

  const handleOpenPicker = async () => {
    const picked = await pickLocalFiles();
    if (picked.length > 0) {
      await importFiles(picked, null);
    }
  };

  const renderActiveModule = () => {
    if (!activeTab) return null;
    const commonProps = {
      tab: activeTab,
      mode: activeTab.mode,
      fileRecord: activeRecord,
      canUndo: canUndoTab(activeTab.fileId),
      canRedo: canRedoTab(activeTab.fileId),
      onChangeContent: (next: string) => updateTabContent(activeTab.fileId, next),
      onUndo: () => undoTab(activeTab.fileId),
      onRedo: () => redoTab(activeTab.fileId),
      onSaveNow: () => void saveTabNow(activeTab.fileId),
      onExportDevice: () => void handleExportDevice(),
      onToggleFavorite: () => void toggleFavoriteById(activeTab.fileId),
      onDismissRecovery: () => dismissCrashRecoveryBanner(activeTab.fileId),
      onEnterEdit: () => handleEnterEdit(activeTab.fileId),
      onDoneEditing: () => void handleDoneEditing(activeTab.fileId),
      onRequestExitEdit: () => handleRequestExitEdit(activeTab.fileId),
    };

    switch (activeTab.moduleType) {
      case 'pdf':
        return <PdfModule {...commonProps} />;
      case 'docs':
        return <DocsModule {...commonProps} />;
      case 'sheets':
        return <SheetsModule {...commonProps} />;
      case 'slides':
        return <SlidesModule {...commonProps} />;
    }
  };

  return (
    <div className="flex min-h-0 flex-1 flex-col overflow-hidden">
      {/* Unopened crash recovery queue */}
      {recoveryQueue.length > 0 && (
        <div
          role="region"
          aria-label="Crash recovery"
          className="border-b border-amber-200 bg-amber-50 px-4 py-3 text-sm text-amber-900 dark:border-amber-900/60 dark:bg-amber-950/40 dark:text-amber-200"
        >
          <div className="flex flex-wrap items-center justify-between gap-3">
            <span>
              Found {recoveryQueue.length} unsaved session snapshot
              {recoveryQueue.length > 1 ? 's' : ''} from an unexpected close.
            </span>
            <div className="flex flex-wrap gap-2">
              {recoveryQueue.map((snap) => (
                <div key={snap.fileId} className="flex items-center gap-1.5">
                  <button
                    type="button"
                    onClick={() => void openFileInTab(snap.fileId)}
                    className="min-h-[34px] rounded-lg bg-amber-700 px-3 text-xs font-semibold text-white hover:bg-amber-800"
                  >
                    Recover {snap.title}
                  </button>
                  <button
                    type="button"
                    onClick={() => void discardRecoverySnapshot(snap.fileId)}
                    className="min-h-[34px] rounded-lg border border-amber-300 px-2.5 text-xs font-medium hover:bg-amber-100 dark:border-amber-800"
                  >
                    Discard
                  </button>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* Multi-document tab strip */}
      <div className="flex items-center justify-between gap-2 border-b border-slate-200 bg-slate-100/80 px-3 py-1.5 dark:border-slate-800 dark:bg-slate-950">
        <div
          role="tablist"
          aria-label="Open documents"
          className="flex min-w-0 flex-1 items-center gap-1 overflow-x-auto"
        >
          {tabs.map((tab) => {
            const active = tab.fileId === activeTab?.fileId;
            return (
              <div
                key={tab.fileId}
                className={`group flex max-w-[220px] shrink-0 items-center gap-1 rounded-xl border px-2.5 py-1.5 text-xs transition-colors ${
                  active
                    ? 'border-slate-300 bg-white text-slate-900 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-100'
                    : 'border-transparent text-slate-600 hover:bg-slate-200/70 dark:text-slate-400 dark:hover:bg-slate-800'
                }`}
              >
                <button
                  type="button"
                  role="tab"
                  aria-selected={active}
                  onClick={() => setActiveTab(tab.fileId)}
                  className="flex min-w-0 flex-1 items-center gap-2 text-left"
                >
                  <ModuleIcon type={tab.moduleType} size="sm" />
                  <span className="truncate font-medium">
                    {tab.name}
                    {tab.dirty ? ' *' : ''}
                  </span>
                </button>
                <button
                  type="button"
                  onClick={() => void closeTab(tab.fileId)}
                  aria-label={`Close ${tab.name}`}
                  className="inline-flex h-6 w-6 shrink-0 items-center justify-center rounded-md text-slate-400 hover:bg-slate-200 hover:text-slate-700 dark:hover:bg-slate-700 dark:hover:text-slate-200"
                >
                  <X className="h-3.5 w-3.5" />
                </button>
              </div>
            );
          })}
        </div>

        <div className="flex shrink-0 items-center gap-1.5">
          <button
            type="button"
            onClick={() => void createAndOpenFile('docs')}
            aria-label="New blank document tab"
            className="inline-flex min-h-[36px] items-center gap-1.5 whitespace-nowrap rounded-lg border border-slate-200 bg-white px-2.5 text-xs font-medium text-slate-700 hover:bg-slate-50 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-200 dark:hover:bg-slate-800"
          >
            <Plus className="h-3.5 w-3.5" />
            <span className="hidden sm:inline">New Doc</span>
          </button>
          <button
            type="button"
            onClick={() => void handleOpenPicker()}
            aria-label="Open file into workspace"
            className="inline-flex min-h-[36px] items-center gap-1.5 whitespace-nowrap rounded-lg border border-slate-200 bg-white px-2.5 text-xs font-medium text-slate-700 hover:bg-slate-50 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-200 dark:hover:bg-slate-800"
          >
            <FolderOpen className="h-3.5 w-3.5" />
            <span className="hidden sm:inline">Open</span>
          </button>
        </div>
      </div>

      {/* Active Document Module or Empty State */}
      {activeTab ? (
        <div
          data-testid="active-workspace-module"
          data-tab-mode={activeTab.mode}
          data-tab-module={activeTab.moduleType}
          className="min-h-0 flex-1 overflow-hidden"
        >
          <Suspense
            fallback={
              <div className="space-y-4 p-6">
                <div className="h-10 w-1/3 animate-pulse rounded-xl bg-slate-200 dark:bg-slate-800" />
                <div className="h-72 animate-pulse rounded-2xl bg-slate-100 dark:bg-slate-900" />
              </div>
            }
          >
            {renderActiveModule()}
          </Suspense>
        </div>
      ) : (
        <div className="flex flex-1 flex-col items-center justify-center p-8 text-center">
          <div className="max-w-md rounded-2xl border border-slate-200 bg-white p-8 dark:border-slate-800 dark:bg-slate-900">
            <h1 className="text-lg font-bold text-slate-900 dark:text-slate-100">
              No open documents
            </h1>
            <p className="mt-2 text-sm text-slate-600 dark:text-slate-300">
              Open multiple PDFs, Documents, Spreadsheets, or Presentations side-by-side in tabs.
            </p>
            <div className="mt-6 grid grid-cols-2 gap-2.5">
              {(['pdf', 'docs', 'sheets', 'slides'] as OfficeModuleType[]).map((m) => (
                <button
                  key={m}
                  type="button"
                  onClick={() => void createAndOpenFile(m)}
                  className="inline-flex min-h-[44px] items-center justify-center gap-2 rounded-xl border border-slate-200 bg-slate-50 px-3 text-xs font-semibold text-slate-800 hover:bg-slate-100 dark:border-slate-700 dark:bg-slate-800 dark:text-slate-100 dark:hover:bg-slate-700"
                >
                  <FilePlus2 className="h-4 w-4" />
                  New {MODULE_METADATA[m].shortLabel}
                </button>
              ))}
            </div>
            <button
              type="button"
              onClick={() => navigate('/files')}
              className="mt-4 w-full min-h-[44px] rounded-xl bg-blue-700 text-xs font-semibold text-white hover:bg-blue-800 dark:bg-blue-600"
            >
              Browse Local Files
            </button>
          </div>
        </div>
      )}

      {/* Unsaved Changes Exit-Edit Confirmation Modal */}
      {confirmTab && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/50 p-4 backdrop-blur-[1px]"
          onMouseDown={(e) => {
            if (e.target === e.currentTarget && !savingExit) {
              setExitConfirmFileId(null);
            }
          }}
        >
          <div
            role="dialog"
            aria-modal="true"
            aria-labelledby="exit-edit-confirm-title"
            className="w-full max-w-md rounded-2xl border border-slate-200 bg-white p-5 shadow-2xl dark:border-slate-700 dark:bg-slate-900"
          >
            <h2
              id="exit-edit-confirm-title"
              className="text-base font-bold text-slate-900 dark:text-slate-100"
            >
              Unsaved changes in {confirmTab.name}
            </h2>
            <p className="mt-1.5 text-xs leading-relaxed text-slate-600 dark:text-slate-300">
              Save your changes before returning to read-only View mode, or discard them to revert to the last saved version.
            </p>
            <div className="mt-5 flex flex-wrap justify-end gap-2">
              <button
                type="button"
                disabled={savingExit}
                onClick={() => setExitConfirmFileId(null)}
                className="min-h-[38px] rounded-lg border border-slate-200 px-3 text-xs font-semibold text-slate-700 hover:bg-slate-100 disabled:opacity-50 dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800"
              >
                Keep Editing
              </button>
              <button
                type="button"
                disabled={savingExit}
                onClick={() => void handleDiscardExitEdit(confirmTab.fileId)}
                className="min-h-[38px] rounded-lg border border-rose-200 bg-rose-50 px-3 text-xs font-semibold text-rose-700 hover:bg-rose-100 disabled:opacity-50 dark:border-rose-900 dark:bg-rose-950/40 dark:text-rose-300"
              >
                Discard Changes
              </button>
              <button
                type="button"
                disabled={savingExit}
                onClick={() => void handleDoneEditing(confirmTab.fileId)}
                className="min-h-[38px] rounded-lg bg-blue-700 px-4 text-xs font-semibold text-white hover:bg-blue-800 disabled:opacity-50 dark:bg-blue-600"
              >
                {savingExit ? 'Saving...' : 'Save & Exit Edit'}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

