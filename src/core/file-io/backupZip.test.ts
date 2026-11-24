import {test} from 'vitest';
import assert from 'node:assert/strict';
import JSZip from 'jszip';

test('JSZip backup archive structure serializes and restores manifest and file blobs', async () => {
  const zip = new JSZip();
  const manifest = {
    format: 'esz-office-backup',
    version: 1,
    exportedAt: 1700000000000,
    folders: [
      {
        id: 'fld_1',
        name: 'Contracts',
        parentId: null,
        createdAt: 1700000000000,
        updatedAt: 1700000000000,
        isTrashed: false,
        trashedAt: null,
      },
    ],
    files: [
      {
        id: 'file_1',
        name: 'Agreement.docx',
        moduleType: 'docs',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        extension: '.docx',
        folderId: 'fld_1',
        size: 18,
        createdAt: 1700000000000,
        updatedAt: 1700000000000,
        lastOpenedAt: null,
        isFavorite: true,
        isTrashed: false,
        trashedAt: null,
        storageBackend: 'idb',
        opfsPath: null,
        previewText: 'Sample agreement',
      },
    ],
  };

  zip.file('manifest.json', JSON.stringify(manifest));
  zip.folder('blobs')?.file('file_1', 'Sample agreement!!');

  const archiveBytes = await zip.generateAsync({type: 'uint8array'});
  const reloaded = await JSZip.loadAsync(archiveBytes);
  const parsedManifest = JSON.parse(await reloaded.file('manifest.json')!.async('string'));
  const restoredText = await reloaded.file('blobs/file_1')!.async('string');

  assert.equal(parsedManifest.format, 'esz-office-backup');
  assert.equal(parsedManifest.files[0].name, 'Agreement.docx');
  assert.equal(restoredText, 'Sample agreement!!');
});
