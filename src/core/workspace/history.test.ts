import {test} from 'vitest';
import assert from 'node:assert/strict';
import {
  canRedo,
  canUndo,
  createHistory,
  pushHistory,
  redoHistory,
  undoHistory,
} from './history';

test('undo/redo history tracks edits and clears redo branch on new edit', () => {
  let h = createHistory('v1', 5);
  assert.equal(canUndo(h), false);
  assert.equal(canRedo(h), false);

  h = pushHistory(h, 'v2');
  h = pushHistory(h, 'v3');
  assert.equal(h.present, 'v3');
  assert.equal(canUndo(h), true);

  h = undoHistory(h);
  assert.equal(h.present, 'v2');
  assert.equal(canRedo(h), true);

  h = redoHistory(h);
  assert.equal(h.present, 'v3');

  h = undoHistory(h);
  h = pushHistory(h, 'v2-branch');
  assert.equal(canRedo(h), false);
  assert.equal(h.present, 'v2-branch');
});

test('history enforces maxDepth cap', () => {
  let h = createHistory(0, 2);
  h = pushHistory(h, 1);
  h = pushHistory(h, 2);
  h = pushHistory(h, 3);
  assert.equal(h.past.length, 2);
  assert.deepEqual(h.past, [1, 2]);
});
