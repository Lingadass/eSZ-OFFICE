import type {
  OfficeFileRecord,
  OfficeFolderRecord,
  OfficeModuleType,
  SortDirection,
  SortField,
} from '../types';

export interface ModuleMetadata {
  type: OfficeModuleType;
  label: string;
  shortLabel: string;
  defaultExtension: string;
  mimeType: string;
  phaseNote: string;
}

export const MODULE_METADATA: Record<OfficeModuleType, ModuleMetadata> = {
  pdf: {
    type: 'pdf',
    label: 'PDF Document',
    shortLabel: 'PDF',
    defaultExtension: '.pdf',
    mimeType: 'application/pdf',
    phaseNote: 'Phase 1 delivers full PDF viewer, organizer, annotations, signing, security, and OCR worker.',
  },
  docs: {
    type: 'docs',
    label: 'Document',
    shortLabel: 'Doc',
    defaultExtension: '.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    phaseNote: 'Phase 2 delivers Word-style ribbon, TipTap editor, .docx import/export, and PDF export.',
  },
  sheets: {
    type: 'sheets',
    label: 'Spreadsheet',
    shortLabel: 'Sheet',
    defaultExtension: '.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    phaseNote: 'Phase 3 delivers multi-sheet grid, formula engine, charts, and .xlsx/.csv import/export.',
  },
  slides: {
    type: 'slides',
    label: 'Presentation',
    shortLabel: 'Slides',
    defaultExtension: '.pptx',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    phaseNote: 'Phase 4 delivers Konva slide editor, presenter mode, themes, and .pptx/.pdf export.',
  },
};

export function inferModuleFromFilename(filename: string, mimeType = ''): OfficeModuleType {
  const lower = filename.toLowerCase().trim();
  if (lower.endsWith('.pdf') || mimeType === 'application/pdf') {
    return 'pdf';
  }
  if (
    lower.endsWith('.xlsx') ||
    lower.endsWith('.xls') ||
    lower.endsWith('.csv') ||
    mimeType.includes('spreadsheet') ||
    mimeType === 'text/csv'
  ) {
    return 'sheets';
  }
  if (
    lower.endsWith('.pptx') ||
    lower.endsWith('.ppt') ||
    mimeType.includes('presentation')
  ) {
    return 'slides';
  }
  return 'docs';
}

export function extractExtension(filename: string, fallbackModule: OfficeModuleType): string {
  const trimmed = filename.trim();
  const dotIdx = trimmed.lastIndexOf('.');
  if (dotIdx > 0 && dotIdx < trimmed.length - 1) {
    return trimmed.slice(dotIdx).toLowerCase();
  }
  return MODULE_METADATA[fallbackModule].defaultExtension;
}

export function ensureExtension(filename: string, moduleType: OfficeModuleType): string {
  const trimmed = filename.trim() || `Untitled${MODULE_METADATA[moduleType].defaultExtension}`;
  if (/\.[a-z0-9]{1,8}$/i.test(trimmed)) {
    return trimmed;
  }
  return `${trimmed}${MODULE_METADATA[moduleType].defaultExtension}`;
}

export function buildDuplicateName(originalName: string, existingNames: Iterable<string>): string {
  const used = new Set(Array.from(existingNames, (n) => n.toLowerCase()));
  const dotIdx = originalName.lastIndexOf('.');
  const hasExt = dotIdx > 0;
  const stem = hasExt ? originalName.slice(0, dotIdx) : originalName;
  const ext = hasExt ? originalName.slice(dotIdx) : '';

  const baseCopy = `${stem} (Copy)${ext}`;
  if (!used.has(baseCopy.toLowerCase())) {
    return baseCopy;
  }

  let counter = 2;
  while (counter < 1000) {
    const candidate = `${stem} (Copy ${counter})${ext}`;
    if (!used.has(candidate.toLowerCase())) {
      return candidate;
    }
    counter += 1;
  }
  return `${stem} (${Date.now()})${ext}`;
}

export function sortFiles(
  files: OfficeFileRecord[],
  field: SortField,
  direction: SortDirection,
): OfficeFileRecord[] {
  const dir = direction === 'asc' ? 1 : -1;
  return [...files].sort((a, b) => {
    switch (field) {
      case 'name':
        return a.name.localeCompare(b.name, 'en', {numeric: true, sensitivity: 'base'}) * dir;
      case 'size':
        return (a.size - b.size) * dir;
      case 'moduleType':
        return a.moduleType.localeCompare(b.moduleType) * dir || a.name.localeCompare(b.name) * dir;
      case 'updatedAt':
      default:
        return (a.updatedAt - b.updatedAt) * dir;
    }
  });
}

export function sortFolders(
  folders: OfficeFolderRecord[],
  direction: SortDirection,
): OfficeFolderRecord[] {
  const dir = direction === 'asc' ? 1 : -1;
  return [...folders].sort(
    (a, b) => a.name.localeCompare(b.name, 'en', {numeric: true, sensitivity: 'base'}) * dir,
  );
}

export function buildFolderBreadcrumbs(
  currentFolderId: string | null,
  folders: OfficeFolderRecord[],
): OfficeFolderRecord[] {
  if (!currentFolderId) return [];
  const byId = new Map(folders.map((f) => [f.id, f]));
  const chain: OfficeFolderRecord[] = [];
  const visited = new Set<string>();
  let cursor: string | null = currentFolderId;

  while (cursor && !visited.has(cursor)) {
    visited.add(cursor);
    const folder = byId.get(cursor);
    if (!folder) break;
    chain.unshift(folder);
    cursor = folder.parentId;
  }
  return chain;
}

export function filterFilesByQuery(
  files: OfficeFileRecord[],
  query: string,
): OfficeFileRecord[] {
  const q = query.trim().toLowerCase();
  if (!q) return files;
  return files.filter(
    (f) =>
      f.name.toLowerCase().includes(q) ||
      f.moduleType.toLowerCase().includes(q) ||
      f.previewText.toLowerCase().includes(q),
  );
}

export function formatBytes(bytes: number): string {
  if (!Number.isFinite(bytes) || bytes <= 0) return '0 B';
  const units = ['B', 'KB', 'MB', 'GB', 'TB'];
  const exp = Math.min(units.length - 1, Math.floor(Math.log(bytes) / Math.log(1024)));
  const val = bytes / Math.pow(1024, exp);
  return `${val < 10 && exp > 0 ? val.toFixed(1) : Math.round(val)} ${units[exp]}`;
}

export function isQuotaExceededError(err: unknown): boolean {
  if (!err || typeof err !== 'object') return false;
  const maybeErr = err as {name?: string; code?: number; message?: string};
  return (
    maybeErr.name === 'QuotaExceededError' ||
    maybeErr.name === 'NS_ERROR_DOM_QUOTA_REACHED' ||
    maybeErr.code === 22 ||
    maybeErr.code === 1014 ||
    (typeof maybeErr.message === 'string' && maybeErr.message.toLowerCase().includes('quota'))
  );
}
