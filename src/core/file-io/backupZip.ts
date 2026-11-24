import JSZip from 'jszip';
import {db} from '../storage/db';
import {readFileBlob, writeFileBlob} from '../storage/opfs';
import type {OfficeFileRecord, OfficeFolderRecord} from '../types';

interface BackupManifestV1 {
  format: 'esz-office-backup';
  version: 1;
  exportedAt: number;
  folders: OfficeFolderRecord[];
  files: OfficeFileRecord[];
}

export async function exportWorkspaceBackupZip(): Promise<{blob: Blob; filename: string; fileCount: number}> {
  const zip = new JSZip();
  const [folders, files] = await Promise.all([db.folders.toArray(), db.files.toArray()]);

  const manifest: BackupManifestV1 = {
    format: 'esz-office-backup',
    version: 1,
    exportedAt: Date.now(),
    folders,
    files,
  };
  zip.file('manifest.json', JSON.stringify(manifest, null, 2));

  const blobsFolder = zip.folder('blobs');
  for (const file of files) {
    const blob = await readFileBlob(file.id, file.storageBackend);
    if (blob && blobsFolder) {
      const buf = await blob.arrayBuffer();
      blobsFolder.file(file.id, buf);
    }
  }

  const zipBlob = await zip.generateAsync({
    type: 'blob',
    compression: 'DEFLATE',
    compressionOptions: {level: 6},
  });
  const dateStamp = new Date().toISOString().slice(0, 10);
  return {
    blob: zipBlob,
    filename: `esz-office-backup-${dateStamp}.zip`,
    fileCount: files.length,
  };
}

export async function importWorkspaceBackupZip(zipFile: File): Promise<{
  importedFiles: number;
  importedFolders: number;
}> {
  const zip = await JSZip.loadAsync(zipFile);
  const manifestEntry = zip.file('manifest.json');
  if (!manifestEntry) {
    throw new Error('Invalid backup archive: missing manifest.json');
  }

  const parsed = JSON.parse(await manifestEntry.async('string')) as BackupManifestV1;
  if (parsed.format !== 'esz-office-backup' || parsed.version !== 1) {
    throw new Error('Unsupported backup archive version.');
  }

  let importedFolders = 0;
  for (const folder of parsed.folders ?? []) {
    await db.folders.put(folder);
    importedFolders += 1;
  }

  let importedFiles = 0;
  for (const file of parsed.files ?? []) {
    const blobEntry = zip.file(`blobs/${file.id}`);
    let storageBackend: 'opfs' | 'idb' = 'idb';
    let opfsPath: string | null = null;
    if (blobEntry) {
      const data = await blobEntry.async('arraybuffer');
      const blob = new Blob([data], {type: file.mimeType});
      const res = await writeFileBlob(file.id, blob);
      storageBackend = res.storageBackend;
      opfsPath = res.opfsPath;
    }
    await db.files.put({
      ...file,
      storageBackend,
      opfsPath,
    });
    importedFiles += 1;
  }

  return {importedFiles, importedFolders};
}
