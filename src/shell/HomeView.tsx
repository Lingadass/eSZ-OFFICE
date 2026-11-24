import React, {useMemo, useState} from 'react';
import {Link, useNavigate} from 'react-router-dom';
import {
  CheckCircle2,
  FilePlus2,
  FileSpreadsheet,
  FileText,
  HardDrive,
  Presentation,
  Search,
  Star,
} from 'lucide-react';
import {filterFilesByQuery, formatBytes, sortFiles} from '../core/storage/fileRules';
import type {OfficeModuleType} from '../core/types';
import {useWorkspaceStore} from '../core/workspace/useWorkspaceStore';
import {ModuleIcon} from './FileThumb';
import {StorageStats} from './StorageStats';

interface HomeViewProps {
  offlineReady: boolean;
}

const CREATE_OPTIONS: Array<{
  type: OfficeModuleType;
  label: string;
  subtitle: string;
  icon: React.FC<{className?: string}>;
  accentClass: string;
}> = [
  {
    type: 'pdf',
    label: 'New PDF',
    subtitle: 'View, organize & sign',
    icon: FileText,
    accentClass: 'bg-red-50 text-red-700 dark:bg-red-950/60 dark:text-red-300',
  },
  {
    type: 'docs',
    label: 'New Document',
    subtitle: 'Word-style rich text',
    icon: FileText,
    accentClass: 'bg-blue-50 text-blue-700 dark:bg-blue-950/60 dark:text-blue-300',
  },
  {
    type: 'sheets',
    label: 'New Spreadsheet',
    subtitle: 'Tables, formulas & CSV',
    icon: FileSpreadsheet,
    accentClass: 'bg-emerald-50 text-emerald-700 dark:bg-emerald-950/60 dark:text-emerald-300',
  },
  {
    type: 'slides',
    label: 'New Presentation',
    subtitle: 'Decks & speaker notes',
    icon: Presentation,
    accentClass: 'bg-amber-50 text-amber-700 dark:bg-amber-950/60 dark:text-amber-300',
  },
];

export const HomeView: React.FC<HomeViewProps> = ({offlineReady}) => {
  const navigate = useNavigate();
  const {
    loading,
    files,
    storageStats,
    createAndOpenFile,
    openFileInTab,
    toggleFavoriteById,
  } = useWorkspaceStore();
  const [query, setQuery] = useState('');

  const filtered = useMemo(() => filterFilesByQuery(files, query), [files, query]);
  const recentFiles = useMemo(
    () =>
      sortFiles(filtered, 'updatedAt', 'desc').slice(0, 6),
    [filtered],
  );
  const favoriteFiles = useMemo(
    () => sortFiles(filtered.filter((f) => f.isFavorite), 'updatedAt', 'desc').slice(0, 6),
    [filtered],
  );

  const handleCreate = async (type: OfficeModuleType) => {
    const created = await createAndOpenFile(type);
    if (created) {
      navigate('/workspace');
    }
  };

  const handleOpenDoc = async (fileId: string) => {
    const tab = await openFileInTab(fileId);
    if (tab) navigate('/workspace');
  };

  return (
    <div className="mx-auto w-full max-w-6xl space-y-8 p-4 sm:p-6 lg:p-8">
      {/* Hero / Quick Create & Status */}
      <section aria-labelledby="home-create-heading">
        <div>
          <h1
            id="home-create-heading"
            className="text-2xl font-bold tracking-tight text-slate-900 dark:text-slate-100 sm:text-3xl"
          >
            Workspace Home
          </h1>
          <div className="mt-2 flex flex-wrap items-center gap-x-3 gap-y-1 text-xs font-medium text-slate-600 dark:text-slate-400">
            <span className="inline-flex items-center gap-1.5 text-emerald-700 dark:text-emerald-400">
              <CheckCircle2 className="h-3.5 w-3.5" />
              {offlineReady ? 'Available offline' : 'Offline cache initializing'}
            </span>
            <span aria-hidden="true">·</span>
            <span>All processing stays on-device</span>
            {storageStats && (
              <>
                <span aria-hidden="true">·</span>
                <span className="tabular-nums">
                  {formatBytes(storageStats.usageBytes)} used
                </span>
              </>
            )}
          </div>
        </div>

        <div className="mt-5 grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-4">
          {CREATE_OPTIONS.map((opt) => {
            const Icon = opt.icon;
            return (
              <button
                key={opt.type}
                type="button"
                onClick={() => void handleCreate(opt.type)}
                className="group flex min-h-[84px] items-center gap-3.5 rounded-2xl border border-slate-200 bg-white p-4 text-left transition-colors hover:border-blue-300 hover:bg-slate-50 dark:border-slate-800 dark:bg-slate-900 dark:hover:border-blue-700 dark:hover:bg-slate-800/70"
              >
                <div
                  className={`flex h-11 w-11 shrink-0 items-center justify-center rounded-xl ${opt.accentClass}`}
                >
                  <Icon className="h-5 w-5" />
                </div>
                <div className="min-w-0 flex-1">
                  <div className="flex items-center gap-1.5 text-sm font-semibold text-slate-900 dark:text-slate-100">
                    <span className="truncate">{opt.label}</span>
                    <FilePlus2 className="h-3.5 w-3.5 shrink-0 opacity-0 transition-opacity group-hover:opacity-100" />
                  </div>
                  <p className="mt-0.5 truncate text-xs text-slate-500 dark:text-slate-400">
                    {opt.subtitle}
                  </p>
                </div>
              </button>
            );
          })}
        </div>
      </section>

      {/* Search + Storage summary row */}
      <section className="grid grid-cols-1 gap-4 lg:grid-cols-3">
        <div className="flex items-center gap-3 rounded-2xl border border-slate-200 bg-white px-4 py-2.5 dark:border-slate-800 dark:bg-slate-900 lg:col-span-2">
          <Search className="h-4 w-4 shrink-0 text-slate-400" />
          <input
            type="search"
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="Filter recent and favourite files..."
            aria-label="Filter home documents"
            className="min-h-[36px] w-full bg-transparent text-sm text-slate-900 placeholder:text-slate-400 focus:outline-none dark:text-slate-100"
          />
        </div>

        <Link
          to="/settings"
          className="flex items-center justify-between gap-4 rounded-2xl border border-slate-200 bg-white px-4 py-3 transition-colors hover:bg-slate-50 dark:border-slate-800 dark:bg-slate-900 dark:hover:bg-slate-800/70"
        >
          <div className="flex min-w-0 items-center gap-3">
            <HardDrive className="h-5 w-5 shrink-0 text-slate-500 dark:text-slate-400" />
            <div className="min-w-0">
              <div className="text-xs font-semibold text-slate-900 dark:text-slate-100">
                Device Vault Storage
              </div>
              <div className="truncate text-xs text-slate-500 tabular-nums dark:text-slate-400">
                {storageStats
                  ? `${formatBytes(storageStats.usageBytes)}${
                      storageStats.quotaBytes > 0
                        ? ` / ${formatBytes(storageStats.quotaBytes)}`
                        : ''
                    }`
                  : 'Calculating...'}
              </div>
            </div>
          </div>
          <span className="shrink-0 text-xs font-semibold text-blue-700 dark:text-blue-400">
            Manage
          </span>
        </Link>
      </section>

      {/* IndexedDB Storage Breakdown by File Type */}
      <StorageStats variant="compact" showManageLink />

      {/* Favourites */}
      <section aria-labelledby="home-favorites-heading">
        <div className="flex items-center justify-between">
          <h2
            id="home-favorites-heading"
            className="text-base font-bold text-slate-900 dark:text-slate-100"
          >
            Favourites
          </h2>
          <Link
            to="/files?view=favorites"
            className="text-xs font-semibold text-blue-700 hover:underline dark:text-blue-400"
          >
            View all
          </Link>
        </div>

        {loading ? (
          <div className="mt-3 grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-3">
            {[0, 1, 2].map((idx) => (
              <div
                key={idx}
                className="h-20 animate-pulse rounded-2xl border border-slate-200 bg-slate-100 dark:border-slate-800 dark:bg-slate-900"
              />
            ))}
          </div>
        ) : favoriteFiles.length === 0 ? (
          <div className="mt-3 rounded-2xl border border-slate-200 bg-white p-6 text-sm text-slate-500 dark:border-slate-800 dark:bg-slate-900 dark:text-slate-400">
            Star any document in Recent or All Files to pin it here for quick offline access.
          </div>
        ) : (
          <div className="mt-3 grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-3">
            {favoriteFiles.map((file) => (
              <div
                key={file.id}
                className="flex items-center justify-between gap-3 rounded-2xl border border-slate-200 bg-white p-3.5 transition-colors hover:border-slate-300 dark:border-slate-800 dark:bg-slate-900 dark:hover:border-slate-700"
              >
                <button
                  type="button"
                  onClick={() => void handleOpenDoc(file.id)}
                  className="flex min-w-0 flex-1 items-center gap-3 text-left"
                >
                  <ModuleIcon type={file.moduleType} />
                  <div className="min-w-0">
                    <div className="truncate text-sm font-semibold text-slate-900 dark:text-slate-100">
                      {file.name}
                    </div>
                    <div className="mt-0.5 text-xs text-slate-500 tabular-nums dark:text-slate-400">
                      {file.moduleType.toUpperCase()} · {formatBytes(file.size)}
                    </div>
                  </div>
                </button>
                <button
                  type="button"
                  onClick={() => void toggleFavoriteById(file.id)}
                  aria-label={`Remove ${file.name} from favourites`}
                  className="inline-flex h-10 w-10 shrink-0 items-center justify-center rounded-xl text-amber-500 hover:bg-slate-100 dark:hover:bg-slate-800"
                >
                  <Star className="h-4 w-4 fill-amber-400" />
                </button>
              </div>
            ))}
          </div>
        )}
      </section>

      {/* Recent Files */}
      <section aria-labelledby="home-recent-heading">
        <div className="flex items-center justify-between">
          <h2
            id="home-recent-heading"
            className="text-base font-bold text-slate-900 dark:text-slate-100"
          >
            Recent Documents
          </h2>
          <Link
            to="/files"
            className="text-xs font-semibold text-blue-700 hover:underline dark:text-blue-400"
          >
            Open File Manager
          </Link>
        </div>

        {loading ? (
          <div className="mt-3 space-y-2">
            {[0, 1, 2, 3].map((idx) => (
              <div
                key={idx}
                className="h-14 animate-pulse rounded-xl border border-slate-200 bg-slate-100 dark:border-slate-800 dark:bg-slate-900"
              />
            ))}
          </div>
        ) : recentFiles.length === 0 ? (
          <div className="mt-3 rounded-2xl border border-slate-200 bg-white p-8 text-center dark:border-slate-800 dark:bg-slate-900">
            <p className="text-sm font-medium text-slate-700 dark:text-slate-300">
              {query ? 'No documents match your filter.' : 'No local documents yet.'}
            </p>
            <p className="mt-1 text-xs text-slate-500 dark:text-slate-400">
              Create a new PDF, Document, Spreadsheet, or Presentation above, or open a file from your device.
            </p>
          </div>
        ) : (
          <div className="mt-3 overflow-hidden rounded-2xl border border-slate-200 bg-white divide-y divide-slate-100 dark:border-slate-800 dark:bg-slate-900 dark:divide-slate-800">
            {recentFiles.map((file) => (
              <div
                key={file.id}
                className="flex items-center justify-between gap-3 px-4 py-3 transition-colors hover:bg-slate-50 dark:hover:bg-slate-800/60"
              >
                <button
                  type="button"
                  onClick={() => void handleOpenDoc(file.id)}
                  className="flex min-w-0 flex-1 items-center gap-3.5 text-left"
                >
                  <ModuleIcon type={file.moduleType} size="sm" />
                  <div className="min-w-0 flex-1">
                    <div className="truncate text-sm font-medium text-slate-900 dark:text-slate-100">
                      {file.name}
                    </div>
                    <div className="mt-0.5 text-xs text-slate-500 tabular-nums dark:text-slate-400">
                      {file.moduleType.toUpperCase()} · {formatBytes(file.size)} · Updated{' '}
                      {new Date(file.updatedAt).toLocaleDateString('en-US', {
                        month: 'short',
                        day: 'numeric',
                      })}
                    </div>
                  </div>
                </button>
                <button
                  type="button"
                  onClick={() => void toggleFavoriteById(file.id)}
                  aria-label={
                    file.isFavorite
                      ? `Remove ${file.name} from favourites`
                      : `Add ${file.name} to favourites`
                  }
                  className="inline-flex h-10 w-10 shrink-0 items-center justify-center rounded-xl text-slate-400 hover:bg-slate-100 hover:text-amber-500 dark:hover:bg-slate-800"
                >
                  <Star
                    className={`h-4 w-4 ${file.isFavorite ? 'fill-amber-400 text-amber-500' : ''}`}
                  />
                </button>
              </div>
            ))}
          </div>
        )}
      </section>
    </div>
  );
};
