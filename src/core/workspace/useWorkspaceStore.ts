import {create} from 'zustand';
import {
  clearRecoverySnapshot,
  createFolder,
  createNewFile,
  duplicateFile,
  emptyTrash,
  ensureStarterWorkspace,
  getDirtyRecoverySnapshots,
  getFileRecord,
  getSetting,
  getStorageStats,
  importExternalFile,
  listActiveFiles,
  listActiveFolders,
  listTrashedFiles,
  listTrashedFolders,
  loadFileContentText,
  markFileOpened,
  moveFileToFolder,
  permanentlyDeleteFile,
  renameFile,
  renameFolder,
  restoreFile,
  restoreFolder,
  saveFileBlob,
  saveFileContentText,
  setSetting,
  toggleFavoriteFile,
  trashFile,
  trashFolder,
  writeRecoverySnapshot,
} from '../storage/fileRepository';
import {StorageQuotaError} from '../storage/opfs';
import {
  canRedo,
  canUndo,
  createHistory,
  type HistoryState,
  pushHistory,
  redoHistory,
  undoHistory,
} from './history';
import type {
  DefaultOpenMode,
  DefaultSaveMode,
  OfficeFileRecord,
  OfficeFolderRecord,
  OfficeModuleType,
  RecoverySnapshot,
  StorageStats,
  TabMode,
  WorkspaceTab,
} from '../types';
import {
  canMutateTab,
  resolveDirtyState,
  resolveInitialTabMode,
  shouldAutosaveTab,
} from './viewModeRules';

export interface ToastItem {
  id: string;
  message: string;
  tone: 'info' | 'success' | 'warning' | 'error';
}

interface WorkspaceState {
  loading: boolean;
  files: OfficeFileRecord[];
  trashedFiles: OfficeFileRecord[];
  folders: OfficeFolderRecord[];
  trashedFolders: OfficeFolderRecord[];
  storageStats: StorageStats | null;
  tabs: WorkspaceTab[];
  activeTabId: string | null;
  histories: Record<string, HistoryState<string>>;
  recoveryQueue: RecoverySnapshot[];
  searchOpen: boolean;
  onboardingCompleted: boolean;
  defaultSaveMode: DefaultSaveMode;
  defaultOpenMode: DefaultOpenMode;
  toasts: ToastItem[];

  bootstrap: (skipOnboarding?: boolean) => Promise<void>;
  refreshLibrary: () => Promise<void>;
  setSearchOpen: (open: boolean) => void;
  completeOnboarding: () => Promise<void>;
  replayOnboarding: () => void;
  setDefaultSaveMode: (mode: DefaultSaveMode) => Promise<void>;
  setDefaultOpenMode: (mode: DefaultOpenMode) => Promise<void>;
  pushToast: (message: string, tone?: ToastItem['tone']) => void;
  dismissToast: (id: string) => void;

  openFileInTab: (fileId: string, explicitMode?: TabMode) => Promise<WorkspaceTab | null>;
  createAndOpenFile: (
    moduleType: OfficeModuleType,
    name?: string,
    folderId?: string | null,
  ) => Promise<OfficeFileRecord | null>;
  importFiles: (files: File[], folderId?: string | null) => Promise<OfficeFileRecord[]>;
  setActiveTab: (fileId: string) => void;
  setTabMode: (fileId: string, mode: TabMode) => void;
  discardTabChanges: (fileId: string) => Promise<void>;
  closeTab: (fileId: string) => Promise<void>;
  updateTabContent: (fileId: string, nextText: string) => void;
  markTabDirty: (fileId: string, dirty?: boolean) => void;
  registerBinarySaveHandler: (fileId: string, handler: (() => Promise<void>) | null) => void;
  saveTabBlob: (
    fileId: string,
    blob: Blob,
    previewText?: string,
    lastPage?: number,
  ) => Promise<boolean>;
  createFileFromBytes: (
    name: string,
    moduleType: OfficeModuleType,
    bytes: Uint8Array,
    folderId?: string | null,
    openTab?: boolean,
  ) => Promise<OfficeFileRecord | null>;
  undoTab: (fileId: string) => void;
  redoTab: (fileId: string) => void;
  canUndoTab: (fileId: string) => boolean;
  canRedoTab: (fileId: string) => boolean;
  saveTabNow: (fileId: string) => Promise<boolean>;
  autosaveDirtyTabs: () => Promise<void>;
  dismissCrashRecoveryBanner: (fileId: string) => void;
  discardRecoverySnapshot: (fileId: string) => Promise<void>;

  renameFileById: (fileId: string, name: string) => Promise<void>;
  duplicateFileById: (fileId: string) => Promise<void>;
  toggleFavoriteById: (fileId: string) => Promise<void>;
  moveFileById: (fileId: string, folderId: string | null) => Promise<void>;
  trashFileById: (fileId: string) => Promise<void>;
  restoreFileById: (fileId: string) => Promise<void>;
  deleteForeverById: (fileId: string) => Promise<void>;
  emptyAllTrash: () => Promise<void>;

  createNewFolder: (name: string, parentId: string | null) => Promise<void>;
  renameFolderById: (folderId: string, name: string) => Promise<void>;
  trashFolderById: (folderId: string) => Promise<void>;
  restoreFolderById: (folderId: string) => Promise<void>;
}

export const useWorkspaceStore = create<WorkspaceState>((set, get) => ({
  loading: true,
  files: [],
  trashedFiles: [],
  folders: [],
  trashedFolders: [],
  storageStats: null,
  tabs: [],
  activeTabId: null,
  histories: {},
  recoveryQueue: [],
  searchOpen: false,
  onboardingCompleted: true,
  defaultSaveMode: 'device-vault',
  defaultOpenMode: 'view',
  toasts: [],
  // Non-serialized registry on store closure for active binary modules (e.g., PDF)
  binarySaveHandlers: {} as Record<string, () => Promise<void>>,

  pushToast: (message, tone = 'info') => {
    const id = `${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
    set((s) => ({toasts: [...s.toasts.slice(-3), {id, message, tone}]}));
    setTimeout(() => {
      set((s) => ({toasts: s.toasts.filter((t) => t.id !== id)}));
    }, 4200);
  },

  dismissToast: (id) => {
    set((s) => ({toasts: s.toasts.filter((t) => t.id !== id)}));
  },

  bootstrap: async (skipOnboarding = false) => {
    set({loading: true});
    try {
      await ensureStarterWorkspace();
      const [
        files,
        trashedFiles,
        folders,
        trashedFolders,
        storageStats,
        onboardingDone,
        saveMode,
        openMode,
        dirtySnapshots,
      ] = await Promise.all([
        listActiveFiles(),
        listTrashedFiles(),
        listActiveFolders(),
        listTrashedFolders(),
        getStorageStats(),
        getSetting<boolean>('onboardingCompleted', false),
        getSetting<string>('defaultSaveMode', 'device-vault'),
        getSetting<string>('defaultOpenMode', 'view'),
        getDirtyRecoverySnapshots(),
      ]);

      set({
        loading: false,
        files,
        trashedFiles,
        folders,
        trashedFolders,
        storageStats,
        onboardingCompleted: skipOnboarding ? true : Boolean(onboardingDone),
        defaultSaveMode: saveMode === 'ask-save-as' ? 'ask-save-as' : 'device-vault',
        defaultOpenMode: openMode === 'edit' ? 'edit' : 'view',
        recoveryQueue: dirtySnapshots,
      });
    } catch (err) {
      set({loading: false});
      get().pushToast(
        err instanceof Error ? err.message : 'Unable to initialize local storage.',
        'error',
      );
    }
  },

  refreshLibrary: async () => {
    const [files, trashedFiles, folders, trashedFolders, storageStats] = await Promise.all([
      listActiveFiles(),
      listTrashedFiles(),
      listActiveFolders(),
      listTrashedFolders(),
      getStorageStats(),
    ]);
    set({files, trashedFiles, folders, trashedFolders, storageStats});
  },

  setSearchOpen: (open) => set({searchOpen: open}),

  completeOnboarding: async () => {
    set({onboardingCompleted: true});
    await setSetting('onboardingCompleted', true);
  },

  replayOnboarding: () => set({onboardingCompleted: false}),

  setDefaultSaveMode: async (mode) => {
    set({defaultSaveMode: mode});
    await setSetting('defaultSaveMode', mode);
  },

  setDefaultOpenMode: async (mode) => {
    set({defaultOpenMode: mode});
    await setSetting('defaultOpenMode', mode);
  },

  openFileInTab: async (fileId, explicitMode) => {
    const existing = get().tabs.find((t) => t.fileId === fileId);
    if (existing) {
      if (explicitMode && existing.mode !== explicitMode) {
        set((s) => ({
          activeTabId: fileId,
          tabs: s.tabs.map((t) =>
            t.fileId === fileId ? {...t, mode: explicitMode} : t,
          ),
        }));
      } else {
        set({activeTabId: fileId});
      }
      await markFileOpened(fileId);
      await get().refreshLibrary();
      return get().tabs.find((t) => t.fileId === fileId) ?? existing;
    }

    const record = await getFileRecord(fileId);
    if (!record) {
      get().pushToast('File could not be found in local vault.', 'error');
      return null;
    }

    const recovery = get().recoveryQueue.find((r) => r.fileId === fileId);
    const contentText = recovery ? recovery.contentText : await loadFileContentText(fileId);
    const mode = resolveInitialTabMode({
      hasRecovery: Boolean(recovery),
      defaultOpenMode: get().defaultOpenMode,
      explicitMode,
    });
    const tab: WorkspaceTab = {
      fileId: record.id,
      name: record.name,
      moduleType: record.moduleType,
      extension: record.extension,
      mode,
      discardNonce: 0,
      dirty: resolveDirtyState(mode, Boolean(recovery)),
      lastSavedAt: record.updatedAt,
      autosaveStatus: 'saved',
      contentText,
      recoveredFromCrash: Boolean(recovery),
    };

    await markFileOpened(fileId);
    set((s) => ({
      tabs: [...s.tabs, tab],
      activeTabId: record.id,
      histories: {
        ...s.histories,
        [record.id]: createHistory(contentText),
      },
      recoveryQueue: s.recoveryQueue.filter((r) => r.fileId !== fileId),
    }));
    await get().refreshLibrary();
    return tab;
  },

  createAndOpenFile: async (moduleType, name, folderId = null) => {
    try {
      const rec = await createNewFile({moduleType, name, folderId});
      await get().refreshLibrary();
      await get().openFileInTab(rec.id, 'edit');
      get().pushToast(`Created ${rec.name}`, 'success');
      return rec;
    } catch (err) {
      get().pushToast(
        err instanceof StorageQuotaError
          ? err.message
          : 'Failed to create file in device storage.',
        'error',
      );
      return null;
    }
  },

  importFiles: async (incoming, folderId = null) => {
    const imported: OfficeFileRecord[] = [];
    for (const file of incoming) {
      try {
        const rec = await importExternalFile(file, folderId);
        imported.push(rec);
      } catch (err) {
        get().pushToast(
          err instanceof StorageQuotaError
            ? err.message
            : `Failed to import ${file.name}.`,
          'error',
        );
      }
    }
    if (imported.length > 0) {
      await get().refreshLibrary();
      await get().openFileInTab(imported[0].id);
      get().pushToast(
        imported.length === 1
          ? `Opened ${imported[0].name}`
          : `Imported ${imported.length} files`,
        'success',
      );
    }
    return imported;
  },

  setActiveTab: (fileId) => set({activeTabId: fileId}),

  setTabMode: (fileId, mode) => {
    set((s) => ({
      tabs: s.tabs.map((t) =>
        t.fileId === fileId
          ? {...t, mode, dirty: mode === 'view' ? false : t.dirty}
          : t,
      ),
    }));
  },

  discardTabChanges: async (fileId) => {
    await clearRecoverySnapshot(fileId);
    const contentText = await loadFileContentText(fileId);
    set((s) => ({
      tabs: s.tabs.map((t) =>
        t.fileId === fileId
          ? {
              ...t,
              mode: 'view',
              dirty: false,
              autosaveStatus: 'saved',
              contentText,
              discardNonce: t.discardNonce + 1,
            }
          : t,
      ),
      histories: {
        ...s.histories,
        [fileId]: createHistory(contentText),
      },
    }));
  },

  closeTab: async (fileId) => {
    const tab = get().tabs.find((t) => t.fileId === fileId);
    if (tab && shouldAutosaveTab(tab)) {
      await get().saveTabNow(fileId);
    }
    set((s) => {
      const remaining = s.tabs.filter((t) => t.fileId !== fileId);
      let nextActive = s.activeTabId;
      if (s.activeTabId === fileId) {
        nextActive = remaining.length > 0 ? remaining[remaining.length - 1].fileId : null;
      }
      const nextHistories = {...s.histories};
      delete nextHistories[fileId];
      return {tabs: remaining, activeTabId: nextActive, histories: nextHistories};
    });
  },

  updateTabContent: (fileId, nextText) => {
    const currentTab = get().tabs.find((t) => t.fileId === fileId);
    if (!currentTab || !canMutateTab(currentTab.mode)) return;
    set((s) => {
      const hist = s.histories[fileId] ?? createHistory(nextText);
      return {
        tabs: s.tabs.map((t) =>
          t.fileId === fileId ? {...t, contentText: nextText, dirty: true} : t,
        ),
        histories: {
          ...s.histories,
          [fileId]: pushHistory(hist, nextText),
        },
      };
    });
    const tab = get().tabs.find((t) => t.fileId === fileId);
    if (tab && canMutateTab(tab.mode)) {
      void writeRecoverySnapshot({
        fileId,
        title: tab.name,
        contentText: nextText,
        savedAt: Date.now(),
        dirty: true,
      });
    }
  },

  markTabDirty: (fileId, dirty = true) => {
    set((s) => ({
      tabs: s.tabs.map((t) =>
        t.fileId === fileId ? {...t, dirty: resolveDirtyState(t.mode, dirty)} : t,
      ),
    }));
  },

  registerBinarySaveHandler: (fileId, handler) => {
    set((s) => {
      const current = (s as unknown as {binarySaveHandlers: Record<string, () => Promise<void>>})
        .binarySaveHandlers;
      const next = {...current};
      if (handler) {
        next[fileId] = handler;
      } else {
        delete next[fileId];
      }
      return {binarySaveHandlers: next} as Partial<WorkspaceState>;
    });
  },

  saveTabBlob: async (fileId, blob, previewText, lastPage) => {
    const targetTab = get().tabs.find((t) => t.fileId === fileId);
    if (targetTab && !canMutateTab(targetTab.mode)) {
      return false;
    }
    set((s) => ({
      tabs: s.tabs.map((t) =>
        t.fileId === fileId ? {...t, autosaveStatus: 'saving'} : t,
      ),
    }));
    try {
      const updated = await saveFileBlob(fileId, blob, previewText, lastPage);
      set((s) => ({
        tabs: s.tabs.map((t) =>
          t.fileId === fileId
            ? {
                ...t,
                dirty: false,
                autosaveStatus: 'saved',
                lastSavedAt: updated?.updatedAt ?? Date.now(),
              }
            : t,
        ),
      }));
      await get().refreshLibrary();
      return true;
    } catch (err) {
      set((s) => ({
        tabs: s.tabs.map((t) =>
          t.fileId === fileId ? {...t, autosaveStatus: 'error'} : t,
        ),
      }));
      get().pushToast(
        err instanceof StorageQuotaError
          ? err.message
          : 'Failed to save binary document to local storage.',
        'error',
      );
      return false;
    }
  },

  createFileFromBytes: async (
    name,
    moduleType,
    bytes,
    folderId = null,
    openTab = true,
  ) => {
    try {
      const mimeMap: Record<OfficeModuleType, string> = {
        pdf: 'application/pdf',
        docs: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        sheets: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        slides: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      };
      const file = new File([new Uint8Array(bytes)], name, {type: mimeMap[moduleType]});
      const rec = await importExternalFile(file, folderId);
      await get().refreshLibrary();
      if (openTab) {
        await get().openFileInTab(rec.id);
      }
      get().pushToast(`Saved ${rec.name}`, 'success');
      return rec;
    } catch (err) {
      get().pushToast(
        err instanceof StorageQuotaError
          ? err.message
          : `Failed to save ${name} to local vault.`,
        'error',
      );
      return null;
    }
  },

  undoTab: (fileId) => {
    const tab = get().tabs.find((t) => t.fileId === fileId);
    if (!tab || !canMutateTab(tab.mode)) return;
    const hist = get().histories[fileId];
    if (!hist || !canUndo(hist)) return;
    const prev = undoHistory(hist);
    set((s) => ({
      histories: {...s.histories, [fileId]: prev},
      tabs: s.tabs.map((t) =>
        t.fileId === fileId ? {...t, contentText: prev.present, dirty: true} : t,
      ),
    }));
  },

  redoTab: (fileId) => {
    const tab = get().tabs.find((t) => t.fileId === fileId);
    if (!tab || !canMutateTab(tab.mode)) return;
    const hist = get().histories[fileId];
    if (!hist || !canRedo(hist)) return;
    const next = redoHistory(hist);
    set((s) => ({
      histories: {...s.histories, [fileId]: next},
      tabs: s.tabs.map((t) =>
        t.fileId === fileId ? {...t, contentText: next.present, dirty: true} : t,
      ),
    }));
  },

  canUndoTab: (fileId) => {
    const hist = get().histories[fileId];
    return hist ? canUndo(hist) : false;
  },

  canRedoTab: (fileId) => {
    const hist = get().histories[fileId];
    return hist ? canRedo(hist) : false;
  },

  saveTabNow: async (fileId) => {
    const tab = get().tabs.find((t) => t.fileId === fileId);
    if (!tab || !canMutateTab(tab.mode)) return false;
    const handlers = (
      get() as unknown as {binarySaveHandlers?: Record<string, () => Promise<void>>}
    ).binarySaveHandlers;
    const binaryHandler = handlers?.[fileId];
    if (binaryHandler) {
      await binaryHandler();
      return true;
    }
    // Never overwrite binary PDF, DOCX, XLSX, or PPTX files with preview text
    if (
      tab.moduleType === 'pdf' ||
      tab.moduleType === 'docs' ||
      tab.moduleType === 'sheets' ||
      tab.moduleType === 'slides'
    ) {
      return true;
    }
    set((s) => ({
      tabs: s.tabs.map((t) =>
        t.fileId === fileId ? {...t, autosaveStatus: 'saving'} : t,
      ),
    }));
    try {
      const updated = await saveFileContentText(fileId, tab.contentText);
      set((s) => ({
        tabs: s.tabs.map((t) =>
          t.fileId === fileId
            ? {
                ...t,
                dirty: false,
                autosaveStatus: 'saved',
                lastSavedAt: updated?.updatedAt ?? Date.now(),
              }
            : t,
        ),
      }));
      await get().refreshLibrary();
      return true;
    } catch (err) {
      set((s) => ({
        tabs: s.tabs.map((t) =>
          t.fileId === fileId ? {...t, autosaveStatus: 'error'} : t,
        ),
      }));
      get().pushToast(
        err instanceof StorageQuotaError
          ? err.message
          : 'Autosave failed due to local storage error.',
        'error',
      );
      return false;
    }
  },

  autosaveDirtyTabs: async () => {
    const dirtyTabs = get().tabs.filter((t) => shouldAutosaveTab(t));
    for (const tab of dirtyTabs) {
      await get().saveTabNow(tab.fileId);
    }
  },

  dismissCrashRecoveryBanner: (fileId) => {
    set((s) => ({
      tabs: s.tabs.map((t) =>
        t.fileId === fileId ? {...t, recoveredFromCrash: false} : t,
      ),
    }));
  },

  discardRecoverySnapshot: async (fileId) => {
    await clearRecoverySnapshot(fileId);
    set((s) => ({
      recoveryQueue: s.recoveryQueue.filter((r) => r.fileId !== fileId),
    }));
  },

  renameFileById: async (fileId, name) => {
    const updated = await renameFile(fileId, name);
    if (updated) {
      set((s) => ({
        tabs: s.tabs.map((t) =>
          t.fileId === fileId
            ? {...t, name: updated.name, extension: updated.extension}
            : t,
        ),
      }));
      await get().refreshLibrary();
      get().pushToast(`Renamed to ${updated.name}`, 'success');
    }
  },

  duplicateFileById: async (fileId) => {
    try {
      const copy = await duplicateFile(fileId);
      if (copy) {
        await get().refreshLibrary();
        get().pushToast(`Duplicated as ${copy.name}`, 'success');
      }
    } catch (err) {
      get().pushToast(
        err instanceof StorageQuotaError ? err.message : 'Failed to duplicate file.',
        'error',
      );
    }
  },

  toggleFavoriteById: async (fileId) => {
    const next = await toggleFavoriteFile(fileId);
    await get().refreshLibrary();
    get().pushToast(next ? 'Added to Favourites' : 'Removed from Favourites', 'info');
  },

  moveFileById: async (fileId, folderId) => {
    await moveFileToFolder(fileId, folderId);
    await get().refreshLibrary();
    get().pushToast('Moved file', 'success');
  },

  trashFileById: async (fileId) => {
    await trashFile(fileId);
    set((s) => {
      const remaining = s.tabs.filter((t) => t.fileId !== fileId);
      const nextActive =
        s.activeTabId === fileId
          ? remaining.length > 0
            ? remaining[remaining.length - 1].fileId
            : null
          : s.activeTabId;
      return {tabs: remaining, activeTabId: nextActive};
    });
    await get().refreshLibrary();
    get().pushToast('Moved to Trash', 'info');
  },

  restoreFileById: async (fileId) => {
    await restoreFile(fileId);
    await get().refreshLibrary();
    get().pushToast('File restored', 'success');
  },

  deleteForeverById: async (fileId) => {
    await permanentlyDeleteFile(fileId);
    await get().refreshLibrary();
    get().pushToast('File permanently deleted', 'info');
  },

  emptyAllTrash: async () => {
    await emptyTrash();
    await get().refreshLibrary();
    get().pushToast('Trash emptied', 'info');
  },

  createNewFolder: async (name, parentId) => {
    const created = await createFolder(name, parentId);
    await get().refreshLibrary();
    get().pushToast(`Created folder "${created.name}"`, 'success');
  },

  renameFolderById: async (folderId, name) => {
    await renameFolder(folderId, name);
    await get().refreshLibrary();
    get().pushToast('Folder renamed', 'success');
  },

  trashFolderById: async (folderId) => {
    await trashFolder(folderId);
    await get().refreshLibrary();
    get().pushToast('Folder moved to Trash', 'info');
  },

  restoreFolderById: async (folderId) => {
    await restoreFolder(folderId);
    await get().refreshLibrary();
    get().pushToast('Folder restored', 'success');
  },
}));
