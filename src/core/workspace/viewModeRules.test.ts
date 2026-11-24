import {test} from 'vitest';
import assert from 'node:assert/strict';
import {
  canMutateTab,
  resolveDirtyState,
  resolveInitialTabMode,
  resolveModeShortcut,
  shouldAutosaveTab,
  shouldConfirmExitEdit,
} from './viewModeRules';

test('resolveInitialTabMode opens existing files in view mode by default', () => {
  assert.equal(resolveInitialTabMode({defaultOpenMode: 'view'}), 'view');
  assert.equal(resolveInitialTabMode({defaultOpenMode: 'edit'}), 'edit');
});

test('resolveInitialTabMode opens newly created blank files and recovery in edit mode', () => {
  assert.equal(resolveInitialTabMode({defaultOpenMode: 'view', isNewFile: true}), 'edit');
  assert.equal(resolveInitialTabMode({defaultOpenMode: 'view', hasRecovery: true}), 'edit');
  assert.equal(resolveInitialTabMode({defaultOpenMode: 'view', explicitMode: 'edit'}), 'edit');
  assert.equal(resolveInitialTabMode({defaultOpenMode: 'edit', explicitMode: 'view'}), 'view');
});

test('view mode never allows mutation, dirty state, or autosave', () => {
  assert.equal(canMutateTab('view'), false);
  assert.equal(canMutateTab('edit'), true);

  assert.equal(resolveDirtyState('view', true), false);
  assert.equal(resolveDirtyState('view', false), false);
  assert.equal(resolveDirtyState('edit', true), true);
  assert.equal(resolveDirtyState('edit', false), false);

  assert.equal(shouldAutosaveTab({mode: 'view', dirty: true}), false);
  assert.equal(shouldAutosaveTab({mode: 'view', dirty: false}), false);
  assert.equal(shouldAutosaveTab({mode: 'edit', dirty: false}), false);
  assert.equal(shouldAutosaveTab({mode: 'edit', dirty: true}), true);
});

test('shouldConfirmExitEdit only confirms when leaving dirty edit mode', () => {
  assert.equal(shouldConfirmExitEdit('view', false), false);
  assert.equal(shouldConfirmExitEdit('view', true), false);
  assert.equal(shouldConfirmExitEdit('edit', false), false);
  assert.equal(shouldConfirmExitEdit('edit', true), true);
});

test('resolveModeShortcut handles E, Ctrl+E, and Esc safely', () => {
  // Plain E in view mode enters edit when not in an input
  assert.equal(
    resolveModeShortcut({
      key: 'e',
      ctrlOrMeta: false,
      alt: false,
      shift: false,
      mode: 'view',
      isEditableTarget: false,
      hasBlockingOverlay: false,
    }),
    'enter-edit',
  );

  // Plain E inside an input (e.g. search box) must not switch modes
  assert.equal(
    resolveModeShortcut({
      key: 'e',
      ctrlOrMeta: false,
      alt: false,
      shift: false,
      mode: 'view',
      isEditableTarget: true,
      hasBlockingOverlay: false,
    }),
    null,
  );

  // Plain E in edit mode must not exit edit (user is typing text)
  assert.equal(
    resolveModeShortcut({
      key: 'e',
      ctrlOrMeta: false,
      alt: false,
      shift: false,
      mode: 'edit',
      isEditableTarget: false,
      hasBlockingOverlay: false,
    }),
    null,
  );

  // Ctrl+E toggles both directions
  assert.equal(
    resolveModeShortcut({
      key: 'e',
      ctrlOrMeta: true,
      alt: false,
      shift: false,
      mode: 'view',
      isEditableTarget: true,
      hasBlockingOverlay: false,
    }),
    'enter-edit',
  );
  assert.equal(
    resolveModeShortcut({
      key: 'E',
      ctrlOrMeta: true,
      alt: false,
      shift: false,
      mode: 'edit',
      isEditableTarget: true,
      hasBlockingOverlay: false,
    }),
    'request-exit-edit',
  );

  // Esc in edit mode requests exit; Esc in view mode or with modal open does nothing
  assert.equal(
    resolveModeShortcut({
      key: 'Escape',
      ctrlOrMeta: false,
      alt: false,
      shift: false,
      mode: 'edit',
      isEditableTarget: false,
      hasBlockingOverlay: false,
    }),
    'request-exit-edit',
  );
  assert.equal(
    resolveModeShortcut({
      key: 'Escape',
      ctrlOrMeta: false,
      alt: false,
      shift: false,
      mode: 'view',
      isEditableTarget: false,
      hasBlockingOverlay: false,
    }),
    null,
  );
  assert.equal(
    resolveModeShortcut({
      key: 'Escape',
      ctrlOrMeta: false,
      alt: false,
      shift: false,
      mode: 'edit',
      isEditableTarget: false,
      hasBlockingOverlay: true,
    }),
    null,
  );
});
