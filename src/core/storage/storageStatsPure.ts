import type {OfficeFileRecord, OfficeModuleType} from '../types';

export interface FileTypeStorageStat {
  type: OfficeModuleType;
  label: string;
  bytes: number;
  count: number;
  percentOfUsed: number;
}

export interface StorageBreakdown {
  totalBytes: number;
  totalFiles: number;
  trashedBytes: number;
  trashedFiles: number;
  byType: FileTypeStorageStat[];
}

export const FILE_TYPE_ORDER: Array<{type: OfficeModuleType; label: string}> = [
  {type: 'pdf', label: 'PDF'},
  {type: 'docs', label: 'Docs'},
  {type: 'sheets', label: 'Sheets'},
  {type: 'slides', label: 'Slides'},
];

/**
 * Aggregates IndexedDB file records into exact byte usage, file counts,
 * and normalized percentage shares for PDF, Docs, Sheets, and Slides.
 */
export function computeStorageBreakdown(
  files: Array<Pick<OfficeFileRecord, 'moduleType' | 'size' | 'isTrashed'>>,
  includeTrashed = true,
): StorageBreakdown {
  const buckets: Record<OfficeModuleType, {bytes: number; count: number}> = {
    pdf: {bytes: 0, count: 0},
    docs: {bytes: 0, count: 0},
    sheets: {bytes: 0, count: 0},
    slides: {bytes: 0, count: 0},
  };

  let trashedBytes = 0;
  let trashedFiles = 0;
  let totalBytes = 0;
  let totalFiles = 0;

  for (const file of files) {
    const safeBytes =
      typeof file.size === 'number' && Number.isFinite(file.size) && file.size > 0
        ? file.size
        : 0;

    if (file.isTrashed) {
      trashedBytes += safeBytes;
      trashedFiles += 1;
      if (!includeTrashed) continue;
    }

    const bucket = buckets[file.moduleType];
    if (!bucket) continue;
    bucket.bytes += safeBytes;
    bucket.count += 1;
    totalBytes += safeBytes;
    totalFiles += 1;
  }

  const byType: FileTypeStorageStat[] = FILE_TYPE_ORDER.map(({type, label}) => {
    const {bytes, count} = buckets[type];
    const percentOfUsed =
      totalBytes > 0 ? Math.round((bytes / totalBytes) * 1000) / 10 : 0;
    return {
      type,
      label,
      bytes,
      count,
      percentOfUsed,
    };
  });

  return {
    totalBytes,
    totalFiles,
    trashedBytes,
    trashedFiles,
    byType,
  };
}
