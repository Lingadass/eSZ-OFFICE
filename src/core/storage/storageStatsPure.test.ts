import {test} from 'vitest';
import assert from 'node:assert/strict';
import {computeStorageBreakdown} from './storageStatsPure';

test('computeStorageBreakdown returns zeroed stats for an empty IndexedDB file list', () => {
  const result = computeStorageBreakdown([]);
  assert.equal(result.totalBytes, 0);
  assert.equal(result.totalFiles, 0);
  assert.equal(result.trashedBytes, 0);
  assert.equal(result.trashedFiles, 0);
  assert.deepEqual(
    result.byType.map((t) => [t.type, t.bytes, t.count, t.percentOfUsed]),
    [
      ['pdf', 0, 0, 0],
      ['docs', 0, 0, 0],
      ['sheets', 0, 0, 0],
      ['slides', 0, 0, 0],
    ],
  );
});

test('computeStorageBreakdown calculates per-type bytes, counts, and percentages accurately', () => {
  const result = computeStorageBreakdown([
    {moduleType: 'pdf', size: 500, isTrashed: false},
    {moduleType: 'docs', size: 250, isTrashed: false},
    {moduleType: 'sheets', size: 150, isTrashed: false},
    {moduleType: 'slides', size: 100, isTrashed: false},
  ]);

  assert.equal(result.totalBytes, 1000);
  assert.equal(result.totalFiles, 4);
  assert.equal(result.byType[0].percentOfUsed, 50);
  assert.equal(result.byType[1].percentOfUsed, 25);
  assert.equal(result.byType[2].percentOfUsed, 15);
  assert.equal(result.byType[3].percentOfUsed, 10);
});

test('computeStorageBreakdown toggles trashed files and sanitizes invalid sizes', () => {
  const records = [
    {moduleType: 'pdf' as const, size: 600, isTrashed: false},
    {moduleType: 'pdf' as const, size: 400, isTrashed: true},
    {moduleType: 'docs' as const, size: -20, isTrashed: false},
    {moduleType: 'sheets' as const, size: Number.NaN, isTrashed: false},
  ];

  const withTrash = computeStorageBreakdown(records, true);
  assert.equal(withTrash.totalBytes, 1000);
  assert.equal(withTrash.totalFiles, 4);
  assert.equal(withTrash.trashedBytes, 400);
  assert.equal(withTrash.trashedFiles, 1);
  assert.equal(withTrash.byType.find((t) => t.type === 'pdf')?.bytes, 1000);

  const activeOnly = computeStorageBreakdown(records, false);
  assert.equal(activeOnly.totalBytes, 600);
  assert.equal(activeOnly.totalFiles, 3);
  assert.equal(activeOnly.trashedBytes, 400);
  assert.equal(activeOnly.trashedFiles, 1);
  assert.equal(activeOnly.byType.find((t) => t.type === 'pdf')?.bytes, 600);
  assert.equal(activeOnly.byType.find((t) => t.type === 'pdf')?.percentOfUsed, 100);
});
