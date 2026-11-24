export type OfficeModuleType = 'pdf' | 'docs' | 'sheets' | 'slides';

export type SortField = 'updatedAt' | 'name' | 'size' | 'moduleType';
export type SortDirection = 'asc' | 'desc';
export type ViewMode = 'grid' | 'list';
export type ThemePreference = 'light' | 'dark' | 'system';
export type DefaultSaveMode = 'device-vault' | 'ask-save-as';
export type TabMode = 'view' | 'edit';
export type DefaultOpenMode = 'view' | 'edit';

export interface OfficeFileRecord {
  id: string;
  name: string;
  moduleType: OfficeModuleType;
  mimeType: string;
  extension: string;
  folderId: string | null;
  size: number;
  createdAt: number;
  updatedAt: number;
  lastOpenedAt: number | null;
  isFavorite: boolean;
  isTrashed: boolean;
  trashedAt: number | null;
  storageBackend: 'opfs' | 'idb';
  opfsPath: string | null;
  previewText: string;
  lastPage?: number;
}

export interface OfficeFileBlobRecord {
  fileId: string;
  blob: Blob;
  updatedAt: number;
}

export interface OfficeFolderRecord {
  id: string;
  name: string;
  parentId: string | null;
  createdAt: number;
  updatedAt: number;
  isTrashed: boolean;
  trashedAt: number | null;
}

export interface RecoverySnapshot {
  fileId: string;
  title: string;
  contentText: string;
  savedAt: number;
  dirty: boolean;
}

export interface SettingEntry {
  key: string;
  value: string | boolean | number;
}

export interface StorageStats {
  usageBytes: number;
  quotaBytes: number;
  percentUsed: number;
  persisted: boolean;
  opfsSupported: boolean;
  fileCount: number;
  trashedCount: number;
}

export interface WorkspaceTab {
  fileId: string;
  name: string;
  moduleType: OfficeModuleType;
  extension: string;
  mode: TabMode;
  discardNonce: number;
  dirty: boolean;
  lastSavedAt: number;
  autosaveStatus: 'saved' | 'saving' | 'error';
  contentText: string;
  recoveredFromCrash: boolean;
}

export interface DocVersionRecord {
  id: string;
  fileId: string;
  label: string;
  html: string;
  wordCount: number;
  savedAt: number;
}

export interface DocMetaRecord {
  fileId: string;
  metaJson: string;
  updatedAt: number;
}
