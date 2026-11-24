import {beforeEach, describe, expect, it, vi} from 'vitest';
import type {OfficeFileRecord} from '../types';

const mockFileRepo = vi.hoisted(() => ({
  clearRecoverySnapshot: vi.fn(async () => undefined),
  createFolder: vi.fn(),
  createNewFile: vi.fn(),
  duplicateFile: vi.fn(),
  emptyTrash: vi.fn(),
  ensureStarterWorkspace: vi.fn(async () => undefined),
  getDirtyRecoverySnapshots: vi.fn(async () => []),
  getFileRecord: vi.fn(),
  getSetting: vi.fn(),
  getStorageStats: vi.fn(async () => ({
    totalBytes: 0,
    totalFiles: 0,
    byType: {
      pdf: {count: 0, bytes: 0},
      docs: {count: 0, bytes: 0},
      sheets: {count: 0, bytes: 0},
      slides: {count: 0, bytes: 0},
    },
  })),
  importExternalFile: vi.fn(),
  listActiveFiles: vi.fn<() => Promise<OfficeFileRecord[]>>(async () => []),
  listActiveFolders: vi.fn(async () => []),
  listTrashedFiles: vi.fn(async () => []),
  listTrashedFolders: vi.fn(async () => []),
  loadFileContentText: vi.fn(async () => 'Original document body'),
  markFileOpened: vi.fn(async () => undefined),
  moveFileToFolder: vi.fn(),
  permanentlyDeleteFile: vi.fn(),
  renameFile: vi.fn(),
  renameFolder: vi.fn(),
  restoreFile: vi.fn(),
  restoreFolder: vi.fn(),
  saveFileBlob: vi.fn(async () => undefined),
  saveFileContentText: vi.fn(async () => undefined),
  setSetting: vi.fn(async () => undefined),
  toggleFavoriteFile: vi.fn(),
  trashFile: vi.fn(),
  trashFolder: vi.fn(),
  writeRecoverySnapshot: vi.fn(async () => undefined),
}));

vi.mock('../storage/fileRepository', () => mockFileRepo);

import {useWorkspaceStore} from './useWorkspaceStore';

function makeFileRecord(id: string, moduleType: OfficeFileRecord['moduleType']): OfficeFileRecord {
  const ts = 1_700_000_000_000;
  return {
    id,
    name: `Sample.${moduleType === 'docs' ? 'docx' : moduleType === 'sheets' ? 'xlsx' : moduleType === 'slides' ? 'pptx' : 'pdf'}`,
    moduleType,
    mimeType: 'application/octet-stream',
    extension:
      moduleType === 'docs'
        ? '.docx'
        : moduleType === 'sheets'
          ? '.xlsx'
          : moduleType === 'slides'
            ? '.pptx'
            : '.pdf',
    folderId: null,
    size: 512,
    createdAt: ts,
    updatedAt: ts,
    lastOpenedAt: ts,
    isFavorite: false,
    isTrashed: false,
    trashedAt: null,
    storageBackend: 'idb',
    opfsPath: null,
    previewText: 'Original document body',
  };
}

describe('useWorkspaceStore — View mode never writes to storage', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    useWorkspaceStore.setState({
      loading: false,
      files: [],
      trashedFiles: [],
      folders: [],
      trashedFolders: [],
      storageStats: null,
      tabs: [],
      activeTabId: null,
      histories: {},
      recoveryQueue: [],
      defaultOpenMode: 'view',
      toasts: [],
    });
  });

  it('opens existing files in View mode by default and blocks all mutations, dirty flags, saves, and autosaves', async () => {
    const rec = makeFileRecord('file_doc_1', 'docs');
    mockFileRepo.getFileRecord.mockResolvedValueOnce(rec);
    mockFileRepo.listActiveFiles.mockResolvedValue([rec]);

    const tab = await useWorkspaceStore.getState().openFileInTab(rec.id);
    expect(tab?.mode).toBe('view');
    expect(tab?.dirty).toBe(false);

    // Attempt every write/mutation path while in View mode
    useWorkspaceStore.getState().updateTabContent(rec.id, 'Unauthorized edit in View mode');
    useWorkspaceStore.getState().markTabDirty(rec.id, true);
    const savedNow = await useWorkspaceStore.getState().saveTabNow(rec.id);
    const savedBlob = await useWorkspaceStore
      .getState()
      .saveTabBlob(rec.id, new Blob(['mutated']), 'mutated');
    await useWorkspaceStore.getState().autosaveDirtyTabs();

    const currentTab = useWorkspaceStore.getState().tabs.find((t) => t.fileId === rec.id);
    expect(currentTab?.mode).toBe('view');
    expect(currentTab?.dirty).toBe(false);
    expect(currentTab?.contentText).toBe('Original document body');
    expect(savedNow).toBe(false);
    expect(savedBlob).toBe(false);

    // Confirm zero storage write calls occurred
    expect(mockFileRepo.saveFileContentText).not.toHaveBeenCalled();
    expect(mockFileRepo.saveFileBlob).not.toHaveBeenCalled();
    expect(mockFileRepo.writeRecoverySnapshot).not.toHaveBeenCalled();
  });

  it('allows editing and saving only after switching to Edit mode, then re-enforces read-only on Done', async () => {
    const rec = makeFileRecord('file_sheet_1', 'sheets');
    mockFileRepo.getFileRecord.mockResolvedValueOnce(rec);
    mockFileRepo.listActiveFiles.mockResolvedValue([rec]);

    await useWorkspaceStore.getState().openFileInTab(rec.id);

    // Switch to Edit mode
    useWorkspaceStore.getState().setTabMode(rec.id, 'edit');
    useWorkspaceStore.getState().markTabDirty(rec.id, true);
    expect(useWorkspaceStore.getState().tabs[0]?.dirty).toBe(true);

    const blobOk = await useWorkspaceStore
      .getState()
      .saveTabBlob(rec.id, new Blob(['xlsx-bytes']), 'A1,B1');
    expect(blobOk).toBe(true);
    expect(mockFileRepo.saveFileBlob).toHaveBeenCalledTimes(1);

    // Return to View mode -> writes blocked again
    useWorkspaceStore.getState().setTabMode(rec.id, 'view');
    const blockedAfterDone = await useWorkspaceStore
      .getState()
      .saveTabBlob(rec.id, new Blob(['should-not-write']), 'nope');
    expect(blockedAfterDone).toBe(false);
    expect(mockFileRepo.saveFileBlob).toHaveBeenCalledTimes(1);
  });
});
