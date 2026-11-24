import {test} from 'vitest';
import assert from 'node:assert/strict';
import {
  buildDuplicateName,
  buildFolderBreadcrumbs,
  ensureExtension,
  filterFilesByQuery,
  formatBytes,
  inferModuleFromFilename,
  isQuotaExceededError,
  sortFiles,
} from './fileRules';
import type {OfficeFileRecord, OfficeFolderRecord} from '../types';

test('inferModuleFromFilename maps extensions and mime types accurately', () => {
  assert.equal(inferModuleFromFilename('report.pdf'), 'pdf');
  assert.equal(inferModuleFromFilename('notes.docx'), 'docs');
  assert.equal(inferModuleFromFilename('budget.xlsx'), 'sheets');
  assert.equal(inferModuleFromFilename('data.csv', 'text/csv'), 'sheets');
  assert.equal(inferModuleFromFilename('deck.pptx'), 'slides');
  assert.equal(inferModuleFromFilename('readme.txt'), 'docs');
});

test('ensureExtension appends default module extension when missing', () => {
  assert.equal(ensureExtension('Quarterly Plan', 'docs'), 'Quarterly Plan.docx');
  assert.equal(ensureExtension('Quarterly Plan.pdf', 'pdf'), 'Quarterly Plan.pdf');
});

test('buildDuplicateName generates non-colliding copy filenames', () => {
  const existing = ['Budget.xlsx', 'Budget (Copy).xlsx', 'Budget (Copy 2).xlsx'];
  assert.equal(buildDuplicateName('Budget.xlsx', existing), 'Budget (Copy 3).xlsx');
  assert.equal(buildDuplicateName('New.pdf', existing), 'New (Copy).pdf');
});

test('buildFolderBreadcrumbs walks parent chain and prevents cycles', () => {
  const folders: OfficeFolderRecord[] = [
    {id: 'f1', name: 'Work', parentId: null, createdAt: 1, updatedAt: 1, isTrashed: false, trashedAt: null},
    {id: 'f2', name: '2026', parentId: 'f1', createdAt: 2, updatedAt: 2, isTrashed: false, trashedAt: null},
    {id: 'f3', name: 'Q4', parentId: 'f2', createdAt: 3, updatedAt: 3, isTrashed: false, trashedAt: null},
  ];
  const crumbs = buildFolderBreadcrumbs('f3', folders);
  assert.deepEqual(
    crumbs.map((c) => c.name),
    ['Work', '2026', 'Q4'],
  );
  assert.deepEqual(buildFolderBreadcrumbs(null, folders), []);
});

test('sortFiles and filterFilesByQuery work deterministically', () => {
  const base: OfficeFileRecord = {
    id: '1',
    name: 'Alpha.pdf',
    moduleType: 'pdf',
    mimeType: 'application/pdf',
    extension: '.pdf',
    folderId: null,
    size: 1024,
    createdAt: 1,
    updatedAt: 100,
    lastOpenedAt: null,
    isFavorite: false,
    isTrashed: false,
    trashedAt: null,
    storageBackend: 'idb',
    opfsPath: null,
    previewText: 'Contract terms',
  };
  const files: OfficeFileRecord[] = [
    base,
    {...base, id: '2', name: 'Zeta.docx', moduleType: 'docs', size: 4096, updatedAt: 200, previewText: 'Meeting notes'},
  ];

  assert.deepEqual(
    sortFiles(files, 'size', 'desc').map((f) => f.id),
    ['2', '1'],
  );
  assert.deepEqual(
    filterFilesByQuery(files, 'contract').map((f) => f.id),
    ['1'],
  );
});

test('formatBytes and isQuotaExceededError handle edge cases', () => {
  assert.equal(formatBytes(0), '0 B');
  assert.equal(formatBytes(1536), '1.5 KB');
  assert.equal(isQuotaExceededError({name: 'QuotaExceededError'}), true);
  assert.equal(isQuotaExceededError(new Error('Normal failure')), false);
});
