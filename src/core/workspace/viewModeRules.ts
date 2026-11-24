import type {DefaultOpenMode, TabMode} from '../types';

export interface InitialTabModeInput {
  isNewFile?: boolean;
  hasRecovery?: boolean;
  defaultOpenMode: DefaultOpenMode;
  explicitMode?: TabMode;
}

/**
 * Determines whether a tab should open in view or edit mode.
 * - Explicit mode override always wins (e.g. newly created blank file passes 'edit').
 * - Blank newly created files and crash-recovered sessions open in 'edit'.
 * - Existing files opened from Home, Files, Recent, or "Open with" respect defaultOpenMode ('view' by default).
 */
export function resolveInitialTabMode(input: InitialTabModeInput): TabMode {
  if (input.explicitMode) return input.explicitMode;
  if (input.isNewFile || input.hasRecovery) return 'edit';
  return input.defaultOpenMode === 'edit' ? 'edit' : 'view';
}

/**
 * View mode is strictly read-only: tabs in view mode may never mutate content,
 * become dirty, trigger autosave, or write back to storage.
 */
export function canMutateTab(mode: TabMode): boolean {
  return mode === 'edit';
}

export function resolveDirtyState(mode: TabMode, requestedDirty: boolean): boolean {
  if (mode !== 'edit') return false;
  return requestedDirty;
}

export function shouldAutosaveTab(tab: {mode: TabMode; dirty: boolean}): boolean {
  return tab.mode === 'edit' && tab.dirty;
}

export function shouldConfirmExitEdit(mode: TabMode, dirty: boolean): boolean {
  return mode === 'edit' && dirty;
}

export type ModeShortcutAction = 'enter-edit' | 'request-exit-edit' | null;

export interface ModeShortcutInput {
  key: string;
  ctrlOrMeta: boolean;
  alt: boolean;
  shift: boolean;
  mode: TabMode;
  isEditableTarget: boolean;
  hasBlockingOverlay: boolean;
}

/**
 * Keyboard mode shortcuts:
 * - Ctrl+E / Cmd+E toggles between view and edit mode (unless a blocking modal is open).
 * - Plain `E` enters edit mode from view mode when focus is not inside a text control or modal.
 * - `Escape` requests exiting edit mode back to view mode when no modal/overlay is consuming it.
 */
export function resolveModeShortcut(input: ModeShortcutInput): ModeShortcutAction {
  if (input.hasBlockingOverlay) return null;

  const lower = input.key.toLowerCase();

  if (input.ctrlOrMeta && !input.alt && lower === 'e') {
    return input.mode === 'view' ? 'enter-edit' : 'request-exit-edit';
  }

  if (
    !input.ctrlOrMeta &&
    !input.alt &&
    !input.shift &&
    lower === 'e' &&
    input.mode === 'view' &&
    !input.isEditableTarget
  ) {
    return 'enter-edit';
  }

  if (!input.ctrlOrMeta && !input.alt && input.key === 'Escape' && input.mode === 'edit') {
    return 'request-exit-edit';
  }

  return null;
}
