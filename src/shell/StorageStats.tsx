import React, {useCallback, useEffect, useState} from 'react';
import {Link} from 'react-router-dom';
import {HardDrive, RefreshCw} from 'lucide-react';
import {db} from '../core/storage/db';
import {formatBytes} from '../core/storage/fileRules';
import {
  computeStorageBreakdown,
  type StorageBreakdown,
} from '../core/storage/storageStatsPure';
import type {OfficeModuleType} from '../core/types';
import {useWorkspaceStore} from '../core/workspace/useWorkspaceStore';

interface StorageStatsProps {
  /** Compact layout for Workspace Home; full layout for Settings */
  variant?: 'full' | 'compact';
  /** Show a subtle link to /settings when rendered outside Settings */
  showManageLink?: boolean;
  /** Optional wrapper style override when embedded inside an existing section */
  embedded?: boolean;
}

const TYPE_STYLES: Record<
  OfficeModuleType,
  {
    barClass: string;
    swatchClass: string;
    textClass: string;
  }
> = {
  pdf: {
    barClass: 'bg-red-600 dark:bg-red-500',
    swatchClass: 'bg-red-600 dark:bg-red-500',
    textClass: 'text-red-700 dark:text-red-400',
  },
  docs: {
    barClass: 'bg-blue-600 dark:bg-blue-500',
    swatchClass: 'bg-blue-600 dark:bg-blue-500',
    textClass: 'text-blue-700 dark:text-blue-400',
  },
  sheets: {
    barClass: 'bg-emerald-600 dark:bg-emerald-500',
    swatchClass: 'bg-emerald-600 dark:bg-emerald-500',
    textClass: 'text-emerald-700 dark:text-emerald-400',
  },
  slides: {
    barClass: 'bg-amber-500 dark:bg-amber-400',
    swatchClass: 'bg-amber-500 dark:bg-amber-400',
    textClass: 'text-amber-700 dark:text-amber-400',
  },
};

export const StorageStats: React.FC<StorageStatsProps> = ({
  variant = 'full',
  showManageLink = false,
  embedded = false,
}) => {
  // Subscribe to library changes so IndexedDB stats stay fresh after creates/edits/deletes
  const fileListSignal = useWorkspaceStore((s) => s.files.length + s.trashedFiles.length);
  const storageStats = useWorkspaceStore((s) => s.storageStats);

  const [includeTrashed, setIncludeTrashed] = useState(true);
  const [breakdown, setBreakdown] = useState<StorageBreakdown | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const loadFromIndexedDb = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const records = await db.files.toArray();
      setBreakdown(computeStorageBreakdown(records, includeTrashed));
    } catch (err) {
      setError(
        err instanceof Error ? err.message : 'Unable to read storage usage from IndexedDB.',
      );
    } finally {
      setLoading(false);
    }
  }, [includeTrashed]);

  useEffect(() => {
    void loadFromIndexedDb();
  }, [loadFromIndexedDb, fileListSignal, storageStats?.usageBytes]);

  const wrapperClass = embedded
    ? 'border-t border-slate-200 pt-5 dark:border-slate-800'
    : 'rounded-2xl border border-slate-200 bg-white p-5 dark:border-slate-800 dark:bg-slate-900';

  const stackedAriaLabel = breakdown
    ? `IndexedDB storage distribution across file types: ${breakdown.byType
        .map((t) => `${t.label} ${formatBytes(t.bytes)} (${t.percentOfUsed}%)`)
        .join(', ')}`
    : 'IndexedDB storage distribution across file types';

  return (
    <section aria-labelledby="storage-stats-heading" className={wrapperClass}>
      <div className="flex flex-wrap items-start justify-between gap-3">
        <div className="min-w-0">
          <div className="flex items-center gap-2">
            <HardDrive className="h-4 w-4 shrink-0 text-slate-500 dark:text-slate-400" />
            <h3
              id="storage-stats-heading"
              className="text-sm font-bold text-slate-900 dark:text-slate-100"
            >
              Storage Stats by File Type
            </h3>
          </div>
          <p className="mt-1 text-xs text-slate-500 tabular-nums dark:text-slate-400">
            {breakdown ? (
              <>
                <span>{formatBytes(breakdown.totalBytes)} across </span>
                <span>
                  {breakdown.totalFiles} {breakdown.totalFiles === 1 ? 'file' : 'files'}
                </span>
                {breakdown.trashedFiles > 0 && (
                  <>
                    <span aria-hidden="true"> · </span>
                    <span>
                      Trash: {formatBytes(breakdown.trashedBytes)} ({breakdown.trashedFiles})
                    </span>
                  </>
                )}
              </>
            ) : (
              'Reading IndexedDB vault...'
            )}
          </p>
        </div>

        <div className="flex shrink-0 items-center gap-2">
          {variant === 'full' && (
            <div
              role="group"
              aria-label="Filter storage scope"
              className="inline-flex items-center gap-0.5 rounded-lg bg-slate-100 p-0.5 dark:bg-slate-800"
            >
              <button
                type="button"
                onClick={() => setIncludeTrashed(true)}
                aria-pressed={includeTrashed}
                className={`whitespace-nowrap rounded-md px-2.5 py-1 text-xs font-medium transition-colors ${
                  includeTrashed
                    ? 'bg-white text-slate-900 shadow-xs dark:bg-slate-900 dark:text-slate-100'
                    : 'text-slate-600 hover:text-slate-900 dark:text-slate-400 dark:hover:text-slate-100'
                }`}
              >
                All Files
              </button>
              <button
                type="button"
                onClick={() => setIncludeTrashed(false)}
                aria-pressed={!includeTrashed}
                className={`whitespace-nowrap rounded-md px-2.5 py-1 text-xs font-medium transition-colors ${
                  !includeTrashed
                    ? 'bg-white text-slate-900 shadow-xs dark:bg-slate-900 dark:text-slate-100'
                    : 'text-slate-600 hover:text-slate-900 dark:text-slate-400 dark:hover:text-slate-100'
                }`}
              >
                Active Only
              </button>
            </div>
          )}

          <button
            type="button"
            onClick={() => void loadFromIndexedDb()}
            disabled={loading}
            aria-label="Refresh IndexedDB storage stats"
            className="inline-flex h-8 w-8 items-center justify-center rounded-lg border border-slate-200 text-slate-600 transition-colors hover:bg-slate-50 hover:text-slate-900 disabled:opacity-50 dark:border-slate-700 dark:text-slate-300 dark:hover:bg-slate-800"
          >
            <RefreshCw className={`h-3.5 w-3.5 ${loading ? 'animate-spin' : ''}`} />
          </button>

          {showManageLink && (
            <Link
              to="/settings"
              className="whitespace-nowrap text-xs font-semibold text-blue-700 hover:underline dark:text-blue-400"
            >
              Manage
            </Link>
          )}
        </div>
      </div>

      {error ? (
        <div className="mt-4 flex items-center justify-between gap-3 rounded-xl border border-red-200 bg-red-50 px-3.5 py-2.5 text-xs text-red-800 dark:border-red-900/60 dark:bg-red-950/40 dark:text-red-200">
          <span>{error}</span>
          <button
            type="button"
            onClick={() => void loadFromIndexedDb()}
            className="shrink-0 font-semibold underline"
          >
            Retry
          </button>
        </div>
      ) : loading && !breakdown ? (
        <div className="mt-4 space-y-3" aria-busy="true">
          <div className="h-3 w-full animate-pulse rounded-full bg-slate-200 dark:bg-slate-800" />
          <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
            {[0, 1, 2, 3].map((i) => (
              <div
                key={i}
                className="h-11 animate-pulse rounded-lg bg-slate-100 dark:bg-slate-800/70"
              />
            ))}
          </div>
        </div>
      ) : breakdown ? (
        <div className="mt-4 space-y-4">
          {/* Combined Multi-Segment Progress Bar */}
          <div>
            <div
              role="img"
              aria-label={stackedAriaLabel}
              data-testid="storage-stats-stacked-bar"
              className="flex h-3 w-full overflow-hidden rounded-full bg-slate-200 dark:bg-slate-800"
            >
              {breakdown.totalBytes > 0 &&
                breakdown.byType.map((item) => {
                  if (item.bytes <= 0) return null;
                  const style = TYPE_STYLES[item.type];
                  return (
                    <div
                      key={item.type}
                      title={`${item.label}: ${formatBytes(item.bytes)} (${item.percentOfUsed}%)`}
                      style={{width: `${item.percentOfUsed}%`}}
                      className={`h-full transition-all duration-200 first:rounded-l-full last:rounded-r-full ${style.barClass}`}
                    />
                  );
                })}
            </div>

            {breakdown.totalBytes === 0 && (
              <p className="mt-2 text-xs text-slate-500 dark:text-slate-400">
                No files stored in IndexedDB yet. Create or import a PDF, Document, Spreadsheet, or Presentation to see space usage.
              </p>
            )}
          </div>

          {/* Per-File-Type Breakdown with Individual Progress Bars */}
          <div
            className={
              variant === 'compact'
                ? 'grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-4'
                : 'grid grid-cols-1 gap-x-6 gap-y-3.5 sm:grid-cols-2'
            }
          >
            {breakdown.byType.map((item) => {
              const style = TYPE_STYLES[item.type];
              return (
                <div key={item.type} data-testid={`storage-stat-${item.type}`}>
                  <div className="flex items-baseline justify-between gap-2 text-xs">
                    <div className="flex min-w-0 items-center gap-2">
                      <span
                        aria-hidden="true"
                        className={`h-2.5 w-2.5 shrink-0 rounded-xs ${style.swatchClass}`}
                      />
                      <span className="font-semibold text-slate-900 dark:text-slate-100">
                        {item.label}
                      </span>
                      <span className="text-slate-400 dark:text-slate-500" aria-hidden="true">
                        ·
                      </span>
                      <span className="truncate text-slate-500 tabular-nums dark:text-slate-400">
                        {item.count} {item.count === 1 ? 'file' : 'files'}
                      </span>
                    </div>
                    <div className="shrink-0 font-mono text-xs tabular-nums text-slate-700 dark:text-slate-300">
                      <span className="font-semibold">{formatBytes(item.bytes)}</span>
                      <span className="ml-1.5 text-slate-500 dark:text-slate-400">
                        ({item.percentOfUsed}%)
                      </span>
                    </div>
                  </div>

                  <div
                    role="progressbar"
                    aria-label={`${item.label} storage usage`}
                    aria-valuemin={0}
                    aria-valuemax={100}
                    aria-valuenow={Math.round(item.percentOfUsed)}
                    aria-valuetext={`${formatBytes(item.bytes)} (${item.percentOfUsed}% of used space)`}
                    className="mt-1.5 h-2 w-full overflow-hidden rounded-full bg-slate-100 dark:bg-slate-800"
                  >
                    <div
                      className={`h-full rounded-full transition-all duration-200 ${style.barClass}`}
                      style={{
                        width: item.bytes > 0 ? `${Math.max(2, item.percentOfUsed)}%` : '0%',
                      }}
                    />
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      ) : null}
    </section>
  );
};
