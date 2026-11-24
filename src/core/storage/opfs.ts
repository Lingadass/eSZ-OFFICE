import {db} from './db';
import {isQuotaExceededError} from './fileRules';

const OPFS_DIR_NAME = 'esz-office-vault';

export function isOpfsSupported(): boolean {
  return (
    typeof navigator !== 'undefined' &&
    typeof navigator.storage !== 'undefined' &&
    typeof navigator.storage.getDirectory === 'function'
  );
}

async function getVaultDirectory(): Promise<FileSystemDirectoryHandle | null> {
  if (!isOpfsSupported()) return null;
  try {
    const root = await navigator.storage.getDirectory();
    return await root.getDirectoryHandle(OPFS_DIR_NAME, {create: true});
  } catch {
    return null;
  }
}

export class StorageQuotaError extends Error {
  constructor(message = 'Device storage quota exceeded. Free space or export a backup before deleting old files.') {
    super(message);
    this.name = 'StorageQuotaError';
  }
}

export async function writeFileBlob(
  fileId: string,
  blob: Blob,
): Promise<{storageBackend: 'opfs' | 'idb'; opfsPath: string | null}> {
  try {
    const dir = await getVaultDirectory();
    if (dir) {
      const handle = await dir.getFileHandle(fileId, {create: true});
      if (typeof handle.createWritable === 'function') {
        const writable = await handle.createWritable();
        await writable.write(blob);
        await writable.close();
        // Keep IndexedDB clean if previously stored there
        await db.fileBlobs.delete(fileId).catch(() => undefined);
        return {storageBackend: 'opfs', opfsPath: `${OPFS_DIR_NAME}/${fileId}`};
      }
    }
    await db.fileBlobs.put({fileId, blob, updatedAt: Date.now()});
    return {storageBackend: 'idb', opfsPath: null};
  } catch (err) {
    if (isQuotaExceededError(err)) {
      throw new StorageQuotaError();
    }
    // Fallback to IndexedDB if OPFS throws a non-quota error
    try {
      await db.fileBlobs.put({fileId, blob, updatedAt: Date.now()});
      return {storageBackend: 'idb', opfsPath: null};
    } catch (idbErr) {
      if (isQuotaExceededError(idbErr)) {
        throw new StorageQuotaError();
      }
      throw idbErr;
    }
  }
}

export async function readFileBlob(
  fileId: string,
  storageBackend: 'opfs' | 'idb',
): Promise<Blob | null> {
  if (storageBackend === 'opfs') {
    try {
      const dir = await getVaultDirectory();
      if (dir) {
        const handle = await dir.getFileHandle(fileId, {create: false});
        return await handle.getFile();
      }
    } catch {
      // Fallback to IDB check below
    }
  }
  const rec = await db.fileBlobs.get(fileId);
  return rec ? rec.blob : null;
}

export async function deleteFileBlob(fileId: string): Promise<void> {
  try {
    const dir = await getVaultDirectory();
    if (dir) {
      await dir.removeEntry(fileId).catch(() => undefined);
    }
  } catch {
    // Ignore OPFS removal failure
  }
  await db.fileBlobs.delete(fileId).catch(() => undefined);
}

export async function requestPersistentStorage(): Promise<boolean> {
  if (typeof navigator === 'undefined' || !navigator.storage?.persist) {
    return false;
  }
  try {
    return await navigator.storage.persist();
  } catch {
    return false;
  }
}

export async function isStoragePersisted(): Promise<boolean> {
  if (typeof navigator === 'undefined' || !navigator.storage?.persisted) {
    return false;
  }
  try {
    return await navigator.storage.persisted();
  } catch {
    return false;
  }
}

export async function getStorageEstimate(): Promise<{
  usageBytes: number;
  quotaBytes: number;
  percentUsed: number;
}> {
  if (typeof navigator !== 'undefined' && navigator.storage?.estimate) {
    try {
      const est = await navigator.storage.estimate();
      const usageBytes = est.usage ?? 0;
      const quotaBytes = est.quota ?? 0;
      const percentUsed =
        quotaBytes > 0 ? Math.min(100, Math.round((usageBytes / quotaBytes) * 1000) / 10) : 0;
      return {usageBytes, quotaBytes, percentUsed};
    } catch {
      // Fallback
    }
  }
  const blobs = await db.fileBlobs.toArray();
  const usageBytes = blobs.reduce((sum, b) => sum + b.blob.size, 0);
  return {usageBytes, quotaBytes: 0, percentUsed: 0};
}
