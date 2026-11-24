import {db} from './db';
import {
  buildDuplicateName,
  ensureExtension,
  extractExtension,
  inferModuleFromFilename,
  MODULE_METADATA,
} from './fileRules';
import {
  deleteFileBlob,
  getStorageEstimate,
  isOpfsSupported,
  isStoragePersisted,
  readFileBlob,
  writeFileBlob,
} from './opfs';
import type {
  OfficeFileRecord,
  OfficeFolderRecord,
  OfficeModuleType,
  RecoverySnapshot,
  StorageStats,
} from '../types';

function createId(prefix: string): string {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    return `${prefix}_${crypto.randomUUID()}`;
  }
  return `${prefix}_${Date.now()}_${Math.random().toString(36).slice(2, 9)}`;
}

export async function ensureStarterWorkspace(): Promise<void> {
  const existingFlag = await db.settings.get('starterSeeded');
  if (existingFlag?.value === true) return;

  const fileCount = await db.files.count();
  if (fileCount > 0) {
    await db.settings.put({key: 'starterSeeded', value: true});
    return;
  }

  const now = Date.now();
  const folderId = createId('fld');
  const starterFolder: OfficeFolderRecord = {
    id: folderId,
    name: 'Getting Started',
    parentId: null,
    createdAt: now - 3600_000,
    updatedAt: now - 3600_000,
    isTrashed: false,
    trashedAt: null,
  };
  await db.folders.put(starterFolder);

  const samples: Array<{
    name: string;
    moduleType: OfficeModuleType;
    folderId: string | null;
    isFavorite: boolean;
    content: string;
    offsetMs: number;
  }> = [
    {
      name: 'Welcome to eSZ OFFICE.docx',
      moduleType: 'docs',
      folderId: null,
      isFavorite: true,
      offsetMs: 60_000,
      content:
        'eSZ OFFICE runs completely on your device.\n\n- Zero uploads, zero accounts, zero tracking.\n- Works offline via Service Worker, IndexedDB, and Origin Private File System.\n- Use Settings -> Export Backup (.zip) anytime to archive your local vault.',
    },
    {
      name: 'Q4 Budget Tracker.xlsx',
      moduleType: 'sheets',
      folderId: folderId,
      isFavorite: false,
      offsetMs: 180_000,
      content: 'Category,Planned,Actual,Variance\nOperations,4200,3980,220\nSoftware,0,0,0\nHardware,1800,1745,55',
    },
    {
      name: 'Product Roadmap Deck.pptx',
      moduleType: 'slides',
      folderId: folderId,
      isFavorite: false,
      offsetMs: 320_000,
      content: 'Slide 1: eSZ OFFICE Roadmap\nSlide 2: Offline-First Architecture\nSlide 3: PDF, Docs, Sheets & Slides',
    },
    {
      name: 'Offline Privacy Specification.pdf',
      moduleType: 'pdf',
      folderId: null,
      isFavorite: true,
      offsetMs: 500_000,
      content: 'eSZ OFFICE Local Security & Privacy Reference\nAll document processing stays strictly on-device.',
    },
  ];

  for (const s of samples) {
    const id = createId('file');
    const blob = new Blob([s.content], {type: MODULE_METADATA[s.moduleType].mimeType});
    const {storageBackend, opfsPath} = await writeFileBlob(id, blob);
    const ts = now - s.offsetMs;
    await db.files.put({
      id,
      name: s.name,
      moduleType: s.moduleType,
      mimeType: MODULE_METADATA[s.moduleType].mimeType,
      extension: extractExtension(s.name, s.moduleType),
      folderId: s.folderId,
      size: blob.size,
      createdAt: ts,
      updatedAt: ts,
      lastOpenedAt: ts,
      isFavorite: s.isFavorite,
      isTrashed: false,
      trashedAt: null,
      storageBackend,
      opfsPath,
      previewText: s.content.slice(0, 200),
      lastPage: s.moduleType === 'pdf' ? 1 : undefined,
    });
  }

  await db.settings.put({key: 'starterSeeded', value: true});
}

export async function listActiveFiles(): Promise<OfficeFileRecord[]> {
  const all = await db.files.toArray();
  return all.filter((f) => !f.isTrashed);
}

export async function listTrashedFiles(): Promise<OfficeFileRecord[]> {
  const all = await db.files.toArray();
  return all.filter((f) => f.isTrashed);
}

export async function listActiveFolders(): Promise<OfficeFolderRecord[]> {
  const all = await db.folders.toArray();
  return all.filter((f) => !f.isTrashed);
}

export async function listTrashedFolders(): Promise<OfficeFolderRecord[]> {
  const all = await db.folders.toArray();
  return all.filter((f) => f.isTrashed);
}

export async function createNewFile(params: {
  moduleType: OfficeModuleType;
  name?: string;
  folderId?: string | null;
  initialText?: string;
}): Promise<OfficeFileRecord> {
  const meta = MODULE_METADATA[params.moduleType];
  const baseName = ensureExtension(
    params.name || `Untitled ${meta.shortLabel}${meta.defaultExtension}`,
    params.moduleType,
  );
  const siblings = await db.files.where('folderId').equals(params.folderId ?? '').toArray().catch(() => []);
  const activeNames = (await listActiveFiles())
    .filter((f) => f.folderId === (params.folderId ?? null))
    .map((f) => f.name);
  void siblings;
  const uniqueName = activeNames.some((n) => n.toLowerCase() === baseName.toLowerCase())
    ? buildDuplicateName(baseName, activeNames)
    : baseName;

  const initialContent =
    params.initialText ??
    `${uniqueName}\nCreated in eSZ OFFICE (${meta.label}) on ${new Date().toLocaleDateString('en-US')}.\n\n`;

  let blob: Blob;
  if (params.moduleType === 'pdf') {
    const {createBlankPdfBytes} = await import('../../modules/pdf/pdfPure');
    const pdfBytes = await createBlankPdfBytes(uniqueName, params.initialText ?? '');
    blob = new Blob([new Uint8Array(pdfBytes)], {type: meta.mimeType});
  } else {
    blob = new Blob([initialContent], {type: meta.mimeType});
  }

  const id = createId('file');
  const {storageBackend, opfsPath} = await writeFileBlob(id, blob);
  const now = Date.now();

  const record: OfficeFileRecord = {
    id,
    name: uniqueName,
    moduleType: params.moduleType,
    mimeType: meta.mimeType,
    extension: extractExtension(uniqueName, params.moduleType),
    folderId: params.folderId ?? null,
    size: blob.size,
    createdAt: now,
    updatedAt: now,
    lastOpenedAt: now,
    isFavorite: false,
    isTrashed: false,
    trashedAt: null,
    storageBackend,
    opfsPath,
    previewText:
      params.moduleType === 'pdf'
        ? params.initialText?.slice(0, 200) || `PDF Document (${uniqueName})`
        : initialContent.slice(0, 200),
    lastPage: params.moduleType === 'pdf' ? 1 : undefined,
  };
  await db.files.put(record);
  return record;
}

export async function importExternalFile(
  file: File,
  folderId: string | null = null,
): Promise<OfficeFileRecord> {
  const moduleType = inferModuleFromFilename(file.name, file.type);
  const meta = MODULE_METADATA[moduleType];
  const id = createId('file');
  const {storageBackend, opfsPath} = await writeFileBlob(id, file);
  const now = Date.now();

  let previewText = `${file.name} (${file.size} bytes)`;
  if (
    file.type.startsWith('text/') ||
    file.name.toLowerCase().endsWith('.csv') ||
    file.name.toLowerCase().endsWith('.md') ||
    file.name.toLowerCase().endsWith('.txt')
  ) {
    try {
      previewText = (await file.text()).slice(0, 280);
    } catch {
      // keep default
    }
  }

  const record: OfficeFileRecord = {
    id,
    name: file.name,
    moduleType,
    mimeType: file.type || meta.mimeType,
    extension: extractExtension(file.name, moduleType),
    folderId,
    size: file.size,
    createdAt: now,
    updatedAt: now,
    lastOpenedAt: now,
    isFavorite: false,
    isTrashed: false,
    trashedAt: null,
    storageBackend,
    opfsPath,
    previewText,
    lastPage: moduleType === 'pdf' ? 1 : undefined,
  };
  await db.files.put(record);
  return record;
}

export async function getFileRecord(fileId: string): Promise<OfficeFileRecord | undefined> {
  return db.files.get(fileId);
}

export async function loadFileContentText(fileId: string): Promise<string> {
  const rec = await db.files.get(fileId);
  if (!rec) return '';
  const blob = await readFileBlob(fileId, rec.storageBackend);
  if (!blob) return rec.previewText || '';
  if (
    rec.mimeType.startsWith('text/') ||
    ['.txt', '.csv', '.md', '.html', '.rtf'].includes(rec.extension) ||
    blob.size < 256_000
  ) {
    try {
      const text = await blob.text();
      // Guard against binary PDF/OOXML header garbage in preview editor
      if (!text.startsWith('%PDF-') && !text.startsWith('PK\x03\x04')) {
        return text;
      }
    } catch {
      // fallback
    }
  }
  return rec.previewText || `Binary ${rec.extension.toUpperCase()} document (${rec.size} bytes)`;
}

export async function loadFileBlob(fileId: string): Promise<Blob | null> {
  const rec = await db.files.get(fileId);
  if (!rec) return null;
  return readFileBlob(fileId, rec.storageBackend);
}

export async function saveFileContentText(fileId: string, contentText: string): Promise<OfficeFileRecord | null> {
  const rec = await db.files.get(fileId);
  if (!rec) return null;
  const blob = new Blob([contentText], {type: rec.mimeType});
  const {storageBackend, opfsPath} = await writeFileBlob(fileId, blob);
  const now = Date.now();
  const updated: OfficeFileRecord = {
    ...rec,
    size: blob.size,
    updatedAt: now,
    storageBackend,
    opfsPath,
    previewText: contentText.slice(0, 200),
  };
  await db.files.put(updated);
  await db.recoverySnapshots.delete(fileId);
  return updated;
}

export async function saveFileBlob(
  fileId: string,
  blob: Blob,
  previewText?: string,
  lastPage?: number,
): Promise<OfficeFileRecord | null> {
  const rec = await db.files.get(fileId);
  if (!rec) return null;
  const {storageBackend, opfsPath} = await writeFileBlob(fileId, blob);
  const now = Date.now();
  const updated: OfficeFileRecord = {
    ...rec,
    size: blob.size,
    updatedAt: now,
    storageBackend,
    opfsPath,
    previewText: previewText ?? rec.previewText,
    lastPage: lastPage ?? rec.lastPage,
  };
  await db.files.put(updated);
  await db.recoverySnapshots.delete(fileId);
  return updated;
}

export async function updateFileLastPage(fileId: string, lastPage: number): Promise<void> {
  await db.files.update(fileId, {lastPage});
}

export async function markFileOpened(fileId: string): Promise<void> {
  await db.files.update(fileId, {lastOpenedAt: Date.now()});
}

export async function renameFile(fileId: string, newName: string): Promise<OfficeFileRecord | null> {
  const rec = await db.files.get(fileId);
  if (!rec) return null;
  const cleaned = ensureExtension(newName, rec.moduleType);
  const updated: OfficeFileRecord = {
    ...rec,
    name: cleaned,
    extension: extractExtension(cleaned, rec.moduleType),
    updatedAt: Date.now(),
  };
  await db.files.put(updated);
  return updated;
}

export async function duplicateFile(fileId: string): Promise<OfficeFileRecord | null> {
  const rec = await db.files.get(fileId);
  if (!rec) return null;
  const siblings = (await listActiveFiles()).filter((f) => f.folderId === rec.folderId);
  const copyName = buildDuplicateName(
    rec.name,
    siblings.map((f) => f.name),
  );
  const blob = (await readFileBlob(fileId, rec.storageBackend)) ?? new Blob([rec.previewText], {type: rec.mimeType});
  const newId = createId('file');
  const {storageBackend, opfsPath} = await writeFileBlob(newId, blob);
  const now = Date.now();
  const copy: OfficeFileRecord = {
    ...rec,
    id: newId,
    name: copyName,
    createdAt: now,
    updatedAt: now,
    lastOpenedAt: null,
    isFavorite: false,
    isTrashed: false,
    trashedAt: null,
    storageBackend,
    opfsPath,
  };
  await db.files.put(copy);
  return copy;
}

export async function toggleFavoriteFile(fileId: string): Promise<boolean> {
  const rec = await db.files.get(fileId);
  if (!rec) return false;
  const next = !rec.isFavorite;
  await db.files.update(fileId, {isFavorite: next});
  return next;
}

export async function moveFileToFolder(fileId: string, folderId: string | null): Promise<void> {
  await db.files.update(fileId, {folderId, updatedAt: Date.now()});
}

export async function trashFile(fileId: string): Promise<void> {
  await db.files.update(fileId, {isTrashed: true, trashedAt: Date.now()});
}

export async function restoreFile(fileId: string): Promise<void> {
  await db.files.update(fileId, {isTrashed: false, trashedAt: null, updatedAt: Date.now()});
}

export async function permanentlyDeleteFile(fileId: string): Promise<void> {
  await deleteFileBlob(fileId);
  await db.recoverySnapshots.delete(fileId).catch(() => undefined);
  await db.files.delete(fileId);
}

export async function emptyTrash(): Promise<void> {
  const trashedFiles = await listTrashedFiles();
  for (const f of trashedFiles) {
    await permanentlyDeleteFile(f.id);
  }
  const trashedFolders = await listTrashedFolders();
  for (const fld of trashedFolders) {
    await db.folders.delete(fld.id);
  }
}

export async function createFolder(name: string, parentId: string | null = null): Promise<OfficeFolderRecord> {
  const trimmed = name.trim() || 'New Folder';
  const now = Date.now();
  const folder: OfficeFolderRecord = {
    id: createId('fld'),
    name: trimmed,
    parentId,
    createdAt: now,
    updatedAt: now,
    isTrashed: false,
    trashedAt: null,
  };
  await db.folders.put(folder);
  return folder;
}

export async function renameFolder(folderId: string, name: string): Promise<void> {
  const trimmed = name.trim();
  if (!trimmed) return;
  await db.folders.update(folderId, {name: trimmed, updatedAt: Date.now()});
}

export async function trashFolder(folderId: string): Promise<void> {
  const now = Date.now();
  await db.folders.update(folderId, {isTrashed: true, trashedAt: now});
  const childFiles = (await listActiveFiles()).filter((f) => f.folderId === folderId);
  for (const f of childFiles) {
    await db.files.update(f.id, {isTrashed: true, trashedAt: now});
  }
}

export async function restoreFolder(folderId: string): Promise<void> {
  await db.folders.update(folderId, {isTrashed: false, trashedAt: null, updatedAt: Date.now()});
  const childFiles = (await listTrashedFiles()).filter((f) => f.folderId === folderId);
  for (const f of childFiles) {
    await db.files.update(f.id, {isTrashed: false, trashedAt: null});
  }
}

export async function writeRecoverySnapshot(snapshot: RecoverySnapshot): Promise<void> {
  await db.recoverySnapshots.put(snapshot);
}

export async function getDirtyRecoverySnapshots(): Promise<RecoverySnapshot[]> {
  const all = await db.recoverySnapshots.toArray();
  return all.filter((s) => s.dirty);
}

export async function clearRecoverySnapshot(fileId: string): Promise<void> {
  await db.recoverySnapshots.delete(fileId);
}

export async function getSetting<T extends string | boolean | number>(
  key: string,
  fallback: T,
): Promise<T> {
  const entry = await db.settings.get(key);
  return entry ? (entry.value as T) : fallback;
}

export async function setSetting(key: string, value: string | boolean | number): Promise<void> {
  await db.settings.put({key, value});
}

export async function getStorageStats(): Promise<StorageStats> {
  const [estimate, persisted, allFiles] = await Promise.all([
    getStorageEstimate(),
    isStoragePersisted(),
    db.files.toArray(),
  ]);
  const trashedCount = allFiles.filter((f) => f.isTrashed).length;
  return {
    usageBytes: estimate.usageBytes,
    quotaBytes: estimate.quotaBytes,
    percentUsed: estimate.percentUsed,
    persisted,
    opfsSupported: isOpfsSupported(),
    fileCount: allFiles.length - trashedCount,
    trashedCount,
  };
}

export async function listDocVersions(fileId: string) {
  const rows = await db.docVersions.where('fileId').equals(fileId).toArray();
  return rows.sort((a, b) => b.savedAt - a.savedAt);
}

export async function saveDocVersion(params: {
  fileId: string;
  label: string;
  html: string;
  wordCount: number;
}) {
  const rec = {
    id: createId('ver'),
    fileId: params.fileId,
    label: params.label.trim() || `Snapshot ${new Date().toLocaleTimeString('en-US')}`,
    html: params.html,
    wordCount: params.wordCount,
    savedAt: Date.now(),
  };
  await db.docVersions.put(rec);
  // Keep latest 30 snapshots per document to stay storage-friendly
  const all = await listDocVersions(params.fileId);
  if (all.length > 30) {
    const stale = all.slice(30);
    await db.docVersions.bulkDelete(stale.map((s) => s.id));
  }
  return rec;
}

export async function deleteDocVersion(versionId: string): Promise<void> {
  await db.docVersions.delete(versionId);
}

export async function getDocMeta<T>(fileId: string): Promise<T | null> {
  const row = await db.docMeta.get(fileId);
  if (!row) return null;
  try {
    return JSON.parse(row.metaJson) as T;
  } catch {
    return null;
  }
}

export async function saveDocMeta<T>(fileId: string, meta: T): Promise<void> {
  await db.docMeta.put({
    fileId,
    metaJson: JSON.stringify(meta),
    updatedAt: Date.now(),
  });
}
