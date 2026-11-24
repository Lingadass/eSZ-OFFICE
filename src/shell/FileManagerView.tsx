import React, {useMemo, useState} from 'react';
import {useNavigate, useSearchParams} from 'react-router-dom';
import {
  ArrowDownUp,
  ChevronRight,
  Copy,
  Edit3,
  Folder,
  FolderInput,
  FolderPlus,
  LayoutGrid,
  List,
  RotateCcw,
  Search,
  Star,
  Trash2,
  Upload,
} from 'lucide-react';
import {pickLocalFiles} from '../core/file-io/fileSystemAccess';
import {
  buildFolderBreadcrumbs,
  filterFilesByQuery,
  formatBytes,
  sortFiles,
  sortFolders,
} from '../core/storage/fileRules';
import type {
  OfficeFileRecord,
  OfficeFolderRecord,
  SortDirection,
  SortField,
  ViewMode,
} from '../core/types';
import {useWorkspaceStore} from '../core/workspace/useWorkspaceStore';
import {DocumentThumbnail, ModuleIcon} from './FileThumb';

export const FileManagerView: React.FC = () => {
  const navigate = useNavigate();
  const [searchParams, setSearchParams] = useSearchParams();
  const currentView = searchParams.get('view') || 'all'; // 'all' | 'favorites' | 'trash'
  const currentFolderId = searchParams.get('folder') || null;

  const {
    loading,
    files,
    trashedFiles,
    folders,
    trashedFolders,
    openFileInTab,
    importFiles,
    createNewFolder,
    renameFileById,
    renameFolderById,
    duplicateFileById,
    toggleFavoriteById,
    moveFileById,
    trashFileById,
    trashFolderById,
    restoreFileById,
    restoreFolderById,
    deleteForeverById,
    emptyAllTrash,
  } = useWorkspaceStore();

  const [query, setQuery] = useState('');
  const [sortField, setSortField] = useState<SortField>('updatedAt');
  const [sortDir, setSortDir] = useState<SortDirection>('desc');
  const [viewMode, setViewMode] = useState<ViewMode>('grid');

  // Dialog states
  const [newFolderOpen, setNewFolderOpen] = useState(false);
  const [newFolderName, setNewFolderName] = useState('');
  const [renameTarget, setRenameTarget] = useState<{
    kind: 'file' | 'folder';
    id: string;
    name: string;
  } | null>(null);
  const [moveTarget, setMoveTarget] = useState<OfficeFileRecord | null>(null);
  const [confirmEmptyTrash, setConfirmEmptyTrash] = useState(false);

  const breadcrumbs = useMemo(
    () => (currentView === 'all' ? buildFolderBreadcrumbs(currentFolderId, folders) : []),
    [currentView, currentFolderId, folders],
  );

  const visibleFolders = useMemo(() => {
    if (currentView === 'trash') {
      return sortFolders(trashedFolders, 'asc');
    }
    if (currentView === 'favorites') return [];
    const base = folders.filter((f) => f.parentId === currentFolderId);
    const q = query.trim().toLowerCase();
    const filtered = q ? base.filter((f) => f.name.toLowerCase().includes(q)) : base;
    return sortFolders(filtered, sortDir === 'asc' ? 'asc' : 'asc');
  }, [currentView, folders, trashedFolders, currentFolderId, query, sortDir]);

  const visibleFiles = useMemo(() => {
    let base: OfficeFileRecord[] = [];
    if (currentView === 'trash') {
      base = trashedFiles;
    } else if (currentView === 'favorites') {
      base = files.filter((f) => f.isFavorite);
    } else {
      base = query.trim()
        ? files
        : files.filter((f) => f.folderId === currentFolderId);
    }
    const filtered = filterFilesByQuery(base, query);
    return sortFiles(filtered, sortField, sortDir);
  }, [currentView, files, trashedFiles, currentFolderId, query, sortField, sortDir]);

  const handleSetView = (view: 'all' | 'favorites' | 'trash', folderId: string | null = null) => {
    const next = new URLSearchParams();
    if (view !== 'all') next.set('view', view);
    if (folderId) next.set('folder', folderId);
    setSearchParams(next);
  };

  const handleOpenFile = async (fileId: string) => {
    const tab = await openFileInTab(fileId);
    if (tab) navigate('/workspace');
  };

  const handleImport = async () => {
    const picked = await pickLocalFiles();
    if (picked.length > 0) {
      await importFiles(picked, currentView === 'all' ? currentFolderId : null);
    }
  };

  const handleCreateFolderSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newFolderName.trim()) return;
    await createNewFolder(newFolderName.trim(), currentFolderId);
    setNewFolderName('');
    setNewFolderOpen(false);
  };

  const handleRenameSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!renameTarget || !renameTarget.name.trim()) return;
    if (renameTarget.kind === 'file') {
      await renameFileById(renameTarget.id, renameTarget.name.trim());
    } else {
      await renameFolderById(renameTarget.id, renameTarget.name.trim());
    }
    setRenameTarget(null);
  };

  return (
    <div className="mx-auto w-full max-w-6xl space-y-6 p-4 sm:p-6 lg:p-8">
      {/* Header & Primary Controls */}
      <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
        <div>
          <h1 className="text-2xl font-bold tracking-tight text-slate-900 dark:text-slate-100">
            {currentView === 'trash'
              ? 'Trash'
              : currentView === 'favorites'
                ? 'Favourites'
                : 'Local File Manager'}
          </h1>
          <p className="mt-1 text-xs text-slate-500 dark:text-slate-400">
            {currentView === 'trash'
              ? 'Restore deleted items or permanently clear local storage.'
              : 'Organize documents in local folders stored privately on this device.'}
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2">
          {currentView === 'trash' ? (
            (trashedFiles.length > 0 || trashedFolders.length > 0) && (
              <button
                type="button"
                onClick={() => setConfirmEmptyTrash(true)}
                className="inline-flex min-h-[40px] items-center gap-2 whitespace-nowrap rounded-xl border border-red-200 bg-red-50 px-3.5 py-2 text-xs font-semibold text-red-700 hover:bg-red-100 dark:border-red-900 dark:bg-red-950/60 dark:text-red-300"
              >
                <Trash2 className="h-4 w-4" />
                Empty Trash
              </button>
            )
          ) : (
            <>
              <button
                type="button"
                onClick={() => setNewFolderOpen(true)}
                className="inline-flex min-h-[40px] items-center gap-2 whitespace-nowrap rounded-xl border border-slate-200 bg-white px-3.5 py-2 text-xs font-semibold text-slate-700 hover:bg-slate-50 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-200 dark:hover:bg-slate-800"
              >
                <FolderPlus className="h-4 w-4" />
                New Folder
              </button>
              <button
                type="button"
                onClick={() => void handleImport()}
                className="inline-flex min-h-[40px] items-center gap-2 whitespace-nowrap rounded-xl bg-blue-700 px-4 py-2 text-xs font-semibold text-white hover:bg-blue-800 dark:bg-blue-600 dark:hover:bg-blue-500"
              >
                <Upload className="h-4 w-4" />
                Import File
              </button>
            </>
          )}
        </div>
      </div>

      {/* Filter tabs + Sort + View toggle */}
      <div className="flex flex-col gap-3 rounded-2xl border border-slate-200 bg-white p-3 dark:border-slate-800 dark:bg-slate-900 lg:flex-row lg:items-center lg:justify-between">
        <div className="flex flex-wrap items-center gap-2">
          <div className="inline-flex rounded-xl bg-slate-100 p-1 dark:bg-slate-800">
            <button
              type="button"
              onClick={() => handleSetView('all', null)}
              className={`min-h-[36px] whitespace-nowrap rounded-lg px-3 text-xs font-semibold transition-colors ${
                currentView === 'all'
                  ? 'bg-white text-slate-900 shadow-xs dark:bg-slate-700 dark:text-slate-100'
                  : 'text-slate-600 hover:text-slate-900 dark:text-slate-400 dark:hover:text-slate-100'
              }`}
            >
              All Files ({files.length})
            </button>
            <button
              type="button"
              onClick={() => handleSetView('favorites')}
              className={`min-h-[36px] whitespace-nowrap rounded-lg px-3 text-xs font-semibold transition-colors ${
                currentView === 'favorites'
                  ? 'bg-white text-slate-900 shadow-xs dark:bg-slate-700 dark:text-slate-100'
                  : 'text-slate-600 hover:text-slate-900 dark:text-slate-400 dark:hover:text-slate-100'
              }`}
            >
              Favourites ({files.filter((f) => f.isFavorite).length})
            </button>
            <button
              type="button"
              onClick={() => handleSetView('trash')}
              className={`min-h-[36px] whitespace-nowrap rounded-lg px-3 text-xs font-semibold transition-colors ${
                currentView === 'trash'
                  ? 'bg-white text-slate-900 shadow-xs dark:bg-slate-700 dark:text-slate-100'
                  : 'text-slate-600 hover:text-slate-900 dark:text-slate-400 dark:hover:text-slate-100'
              }`}
            >
              Trash ({trashedFiles.length})
            </button>
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-2">
          <div className="flex min-w-[200px] flex-1 items-center gap-2 rounded-xl border border-slate-200 bg-slate-50 px-3 dark:border-slate-700 dark:bg-slate-950 sm:flex-initial">
            <Search className="h-3.5 w-3.5 text-slate-400" />
            <input
              type="search"
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              placeholder="Search in view..."
              aria-label="Search files in current view"
              className="min-h-[36px] w-full bg-transparent text-xs text-slate-900 placeholder:text-slate-400 focus:outline-none dark:text-slate-100"
            />
          </div>

          <label className="sr-only" htmlFor="sort-field-select">
            Sort files by
          </label>
          <select
            id="sort-field-select"
            value={sortField}
            onChange={(e) => setSortField(e.target.value as SortField)}
            className="min-h-[38px] rounded-xl border border-slate-200 bg-white px-3 text-xs font-medium text-slate-700 dark:border-slate-700 dark:bg-slate-800 dark:text-slate-200"
          >
            <option value="updatedAt">Sort: Modified</option>
            <option value="name">Sort: Name</option>
            <option value="size">Sort: Size</option>
            <option value="moduleType">Sort: Type</option>
          </select>

          <button
            type="button"
            onClick={() => setSortDir((d) => (d === 'asc' ? 'desc' : 'asc'))}
            aria-label={`Sort direction: ${sortDir === 'asc' ? 'ascending' : 'descending'}`}
            className="inline-flex min-h-[38px] min-w-[38px] items-center justify-center rounded-xl border border-slate-200 bg-white px-2.5 text-xs font-medium text-slate-700 hover:bg-slate-50 dark:border-slate-700 dark:bg-slate-800 dark:text-slate-200"
          >
            <ArrowDownUp className="h-3.5 w-3.5" />
          </button>

          <div className="inline-flex rounded-xl border border-slate-200 p-0.5 dark:border-slate-700">
            <button
              type="button"
              onClick={() => setViewMode('grid')}
              aria-label="Grid view"
              aria-pressed={viewMode === 'grid'}
              className={`inline-flex h-8 w-8 items-center justify-center rounded-lg ${
                viewMode === 'grid'
                  ? 'bg-slate-100 text-slate-900 dark:bg-slate-800 dark:text-slate-100'
                  : 'text-slate-500'
              }`}
            >
              <LayoutGrid className="h-4 w-4" />
            </button>
            <button
              type="button"
              onClick={() => setViewMode('list')}
              aria-label="List view"
              aria-pressed={viewMode === 'list'}
              className={`inline-flex h-8 w-8 items-center justify-center rounded-lg ${
                viewMode === 'list'
                  ? 'bg-slate-100 text-slate-900 dark:bg-slate-800 dark:text-slate-100'
                  : 'text-slate-500'
              }`}
            >
              <List className="h-4 w-4" />
            </button>
          </div>
        </div>
      </div>

      {/* Folder Breadcrumbs */}
      {currentView === 'all' && (
        <nav aria-label="Folder breadcrumbs" className="flex flex-wrap items-center gap-1.5 text-xs">
          <button
            type="button"
            onClick={() => handleSetView('all', null)}
            className={`rounded-lg px-2.5 py-1.5 font-medium ${
              currentFolderId === null
                ? 'bg-slate-200/70 text-slate-900 dark:bg-slate-800 dark:text-slate-100'
                : 'text-slate-600 hover:bg-slate-100 dark:text-slate-400 dark:hover:bg-slate-800'
            }`}
          >
            Root Vault
          </button>
          {breadcrumbs.map((crumb) => (
            <React.Fragment key={crumb.id}>
              <ChevronRight className="h-3.5 w-3.5 text-slate-400" />
              <button
                type="button"
                onClick={() => handleSetView('all', crumb.id)}
                className={`rounded-lg px-2.5 py-1.5 font-medium ${
                  currentFolderId === crumb.id
                    ? 'bg-slate-200/70 text-slate-900 dark:bg-slate-800 dark:text-slate-100'
                    : 'text-slate-600 hover:bg-slate-100 dark:text-slate-400 dark:hover:bg-slate-800'
                }`}
              >
                {crumb.name}
              </button>
            </React.Fragment>
          ))}
        </nav>
      )}

      {/* Folders Section */}
      {visibleFolders.length > 0 && (
        <section aria-label="Folders">
          <h2 className="mb-2.5 text-xs font-semibold text-slate-500 dark:text-slate-400">
            Folders
          </h2>
          <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-3">
            {visibleFolders.map((folder: OfficeFolderRecord) => (
              <div
                key={folder.id}
                className="flex items-center justify-between gap-2 rounded-2xl border border-slate-200 bg-white p-3 dark:border-slate-800 dark:bg-slate-900"
              >
                <button
                  type="button"
                  onClick={() => {
                    if (currentView !== 'trash') handleSetView('all', folder.id);
                  }}
                  className="flex min-w-0 flex-1 items-center gap-3 text-left"
                >
                  <div className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl bg-slate-100 text-slate-700 dark:bg-slate-800 dark:text-slate-300">
                    <Folder className="h-4 w-4" />
                  </div>
                  <span className="truncate text-sm font-semibold text-slate-900 dark:text-slate-100">
                    {folder.name}
                  </span>
                </button>
                <div className="flex items-center gap-1">
                  {currentView === 'trash' ? (
                    <button
                      type="button"
                      onClick={() => void restoreFolderById(folder.id)}
                      aria-label={`Restore folder ${folder.name}`}
                      className="inline-flex h-9 w-9 items-center justify-center rounded-lg text-slate-600 hover:bg-slate-100 dark:text-slate-300 dark:hover:bg-slate-800"
                    >
                      <RotateCcw className="h-4 w-4" />
                    </button>
                  ) : (
                    <>
                      <button
                        type="button"
                        onClick={() =>
                          setRenameTarget({kind: 'folder', id: folder.id, name: folder.name})
                        }
                        aria-label={`Rename folder ${folder.name}`}
                        className="inline-flex h-9 w-9 items-center justify-center rounded-lg text-slate-500 hover:bg-slate-100 dark:text-slate-400 dark:hover:bg-slate-800"
                      >
                        <Edit3 className="h-3.5 w-3.5" />
                      </button>
                      <button
                        type="button"
                        onClick={() => void trashFolderById(folder.id)}
                        aria-label={`Move folder ${folder.name} to Trash`}
                        className="inline-flex h-9 w-9 items-center justify-center rounded-lg text-slate-500 hover:bg-red-50 hover:text-red-600 dark:text-slate-400 dark:hover:bg-red-950/60"
                      >
                        <Trash2 className="h-3.5 w-3.5" />
                      </button>
                    </>
                  )}
                </div>
              </div>
            ))}
          </div>
        </section>
      )}

      {/* Files Grid / List */}
      {loading ? (
        <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-3">
          {[0, 1, 2].map((i) => (
            <div
              key={i}
              className="h-44 animate-pulse rounded-2xl border border-slate-200 bg-slate-100 dark:border-slate-800 dark:bg-slate-900"
            />
          ))}
        </div>
      ) : visibleFiles.length === 0 && visibleFolders.length === 0 ? (
        <div className="rounded-2xl border border-slate-200 bg-white p-10 text-center dark:border-slate-800 dark:bg-slate-900">
          <p className="text-sm font-semibold text-slate-800 dark:text-slate-200">
            {currentView === 'trash'
              ? 'Trash is empty'
              : currentView === 'favorites'
                ? 'No favourite documents'
                : 'This folder has no files yet'}
          </p>
          <p className="mt-1 text-xs text-slate-500 dark:text-slate-400">
            {currentView === 'all'
              ? 'Import a file from your device or create a new document from Home.'
              : 'Items you star or delete will appear here.'}
          </p>
        </div>
      ) : viewMode === 'grid' ? (
        <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-3">
          {visibleFiles.map((file) => (
            <article
              key={file.id}
              className="flex flex-col justify-between gap-3 rounded-2xl border border-slate-200 bg-white p-4 transition-colors hover:border-slate-300 dark:border-slate-800 dark:bg-slate-900 dark:hover:border-slate-700"
            >
              <button
                type="button"
                onClick={() => {
                  if (currentView !== 'trash') void handleOpenFile(file.id);
                }}
                className="text-left"
              >
                <DocumentThumbnail type={file.moduleType} previewText={file.previewText} />
                <h3 className="mt-3 truncate text-sm font-semibold text-slate-900 dark:text-slate-100">
                  {file.name}
                </h3>
                <p className="mt-0.5 text-xs text-slate-500 tabular-nums dark:text-slate-400">
                  {formatBytes(file.size)} ·{' '}
                  {new Date(file.updatedAt).toLocaleDateString('en-US', {
                    month: 'short',
                    day: 'numeric',
                    year: 'numeric',
                  })}
                </p>
              </button>

              <div className="flex flex-wrap items-center justify-between gap-1 border-t border-slate-100 pt-2.5 dark:border-slate-800">
                {currentView === 'trash' ? (
                  <div className="flex w-full items-center justify-between gap-2">
                    <button
                      type="button"
                      onClick={() => void restoreFileById(file.id)}
                      className="inline-flex min-h-[38px] items-center gap-1.5 rounded-lg border border-slate-200 px-3 text-xs font-semibold text-slate-700 hover:bg-slate-50 dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800"
                    >
                      <RotateCcw className="h-3.5 w-3.5" />
                      Restore
                    </button>
                    <button
                      type="button"
                      onClick={() => void deleteForeverById(file.id)}
                      className="inline-flex min-h-[38px] items-center gap-1.5 rounded-lg px-3 text-xs font-medium text-red-600 hover:bg-red-50 dark:text-red-400 dark:hover:bg-red-950/60"
                    >
                      Delete Forever
                    </button>
                  </div>
                ) : (
                  <>
                    <button
                      type="button"
                      onClick={() => void toggleFavoriteById(file.id)}
                      aria-label={
                        file.isFavorite
                          ? `Remove ${file.name} from favourites`
                          : `Add ${file.name} to favourites`
                      }
                      className="inline-flex h-9 w-9 items-center justify-center rounded-lg text-slate-500 hover:bg-slate-100 dark:text-slate-400 dark:hover:bg-slate-800"
                    >
                      <Star
                        className={`h-4 w-4 ${
                          file.isFavorite ? 'fill-amber-400 text-amber-500' : ''
                        }`}
                      />
                    </button>
                    <div className="flex items-center gap-1">
                      <button
                        type="button"
                        onClick={() =>
                          setRenameTarget({kind: 'file', id: file.id, name: file.name})
                        }
                        aria-label={`Rename ${file.name}`}
                        title="Rename"
                        className="inline-flex h-9 w-9 items-center justify-center rounded-lg text-slate-500 hover:bg-slate-100 dark:text-slate-400 dark:hover:bg-slate-800"
                      >
                        <Edit3 className="h-4 w-4" />
                      </button>
                      <button
                        type="button"
                        onClick={() => void duplicateFileById(file.id)}
                        aria-label={`Duplicate ${file.name}`}
                        title="Duplicate"
                        className="inline-flex h-9 w-9 items-center justify-center rounded-lg text-slate-500 hover:bg-slate-100 dark:text-slate-400 dark:hover:bg-slate-800"
                      >
                        <Copy className="h-4 w-4" />
                      </button>
                      <button
                        type="button"
                        onClick={() => setMoveTarget(file)}
                        aria-label={`Move ${file.name} to folder`}
                        title="Move to folder"
                        className="inline-flex h-9 w-9 items-center justify-center rounded-lg text-slate-500 hover:bg-slate-100 dark:text-slate-400 dark:hover:bg-slate-800"
                      >
                        <FolderInput className="h-4 w-4" />
                      </button>
                      <button
                        type="button"
                        onClick={() => void trashFileById(file.id)}
                        aria-label={`Delete ${file.name}`}
                        title="Move to Trash"
                        className="inline-flex h-9 w-9 items-center justify-center rounded-lg text-slate-500 hover:bg-red-50 hover:text-red-600 dark:text-slate-400 dark:hover:bg-red-950/60"
                      >
                        <Trash2 className="h-4 w-4" />
                      </button>
                    </div>
                  </>
                )}
              </div>
            </article>
          ))}
        </div>
      ) : (
        <div className="overflow-hidden rounded-2xl border border-slate-200 bg-white divide-y divide-slate-100 dark:border-slate-800 dark:bg-slate-900 dark:divide-slate-800">
          {visibleFiles.map((file) => (
            <div
              key={file.id}
              className="flex flex-wrap items-center justify-between gap-3 px-4 py-3 hover:bg-slate-50 dark:hover:bg-slate-800/60"
            >
              <button
                type="button"
                onClick={() => {
                  if (currentView !== 'trash') void handleOpenFile(file.id);
                }}
                className="flex min-w-0 flex-1 items-center gap-3 text-left"
              >
                <ModuleIcon type={file.moduleType} size="sm" />
                <div className="min-w-0">
                  <div className="truncate text-sm font-medium text-slate-900 dark:text-slate-100">
                    {file.name}
                  </div>
                  <div className="text-xs text-slate-500 tabular-nums dark:text-slate-400">
                    {file.moduleType.toUpperCase()} · {formatBytes(file.size)} ·{' '}
                    {new Date(file.updatedAt).toLocaleDateString('en-US')}
                  </div>
                </div>
              </button>

              <div className="flex items-center gap-1">
                {currentView === 'trash' ? (
                  <>
                    <button
                      type="button"
                      onClick={() => void restoreFileById(file.id)}
                      aria-label={`Restore ${file.name}`}
                      className="inline-flex min-h-[38px] items-center gap-1.5 rounded-lg border border-slate-200 px-3 text-xs font-semibold text-slate-700 hover:bg-slate-100 dark:border-slate-700 dark:text-slate-200"
                    >
                      <RotateCcw className="h-3.5 w-3.5" />
                      Restore
                    </button>
                    <button
                      type="button"
                      onClick={() => void deleteForeverById(file.id)}
                      aria-label={`Permanently delete ${file.name}`}
                      className="inline-flex min-h-[38px] items-center rounded-lg px-3 text-xs font-medium text-red-600 hover:bg-red-50 dark:text-red-400"
                    >
                      Delete Forever
                    </button>
                  </>
                ) : (
                  <>
                    <button
                      type="button"
                      onClick={() => void toggleFavoriteById(file.id)}
                      aria-label={
                        file.isFavorite
                          ? `Remove ${file.name} from favourites`
                          : `Add ${file.name} to favourites`
                      }
                      className="inline-flex h-9 w-9 items-center justify-center rounded-lg text-slate-500 hover:bg-slate-100 dark:hover:bg-slate-800"
                    >
                      <Star
                        className={`h-4 w-4 ${
                          file.isFavorite ? 'fill-amber-400 text-amber-500' : ''
                        }`}
                      />
                    </button>
                    <button
                      type="button"
                      onClick={() =>
                        setRenameTarget({kind: 'file', id: file.id, name: file.name})
                      }
                      aria-label={`Rename ${file.name}`}
                      className="inline-flex h-9 w-9 items-center justify-center rounded-lg text-slate-500 hover:bg-slate-100 dark:hover:bg-slate-800"
                    >
                      <Edit3 className="h-4 w-4" />
                    </button>
                    <button
                      type="button"
                      onClick={() => void duplicateFileById(file.id)}
                      aria-label={`Duplicate ${file.name}`}
                      className="inline-flex h-9 w-9 items-center justify-center rounded-lg text-slate-500 hover:bg-slate-100 dark:hover:bg-slate-800"
                    >
                      <Copy className="h-4 w-4" />
                    </button>
                    <button
                      type="button"
                      onClick={() => setMoveTarget(file)}
                      aria-label={`Move ${file.name} to folder`}
                      className="inline-flex h-9 w-9 items-center justify-center rounded-lg text-slate-500 hover:bg-slate-100 dark:hover:bg-slate-800"
                    >
                      <FolderInput className="h-4 w-4" />
                    </button>
                    <button
                      type="button"
                      onClick={() => void trashFileById(file.id)}
                      aria-label={`Delete ${file.name}`}
                      className="inline-flex h-9 w-9 items-center justify-center rounded-lg text-slate-500 hover:bg-red-50 hover:text-red-600 dark:hover:bg-red-950/60"
                    >
                      <Trash2 className="h-4 w-4" />
                    </button>
                  </>
                )}
              </div>
            </div>
          ))}
        </div>
      )}

      {/* New Folder Dialog */}
      {newFolderOpen && (
        <div
          role="dialog"
          aria-modal="true"
          aria-labelledby="new-folder-title"
          className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/50 p-4 backdrop-blur-xs"
        >
          <form
            onSubmit={(e) => void handleCreateFolderSubmit(e)}
            className="w-full max-w-sm rounded-2xl border border-slate-200 bg-white p-6 dark:border-slate-800 dark:bg-slate-900"
          >
            <h2 id="new-folder-title" className="text-base font-bold text-slate-900 dark:text-slate-100">
              Create Folder
            </h2>
            <input
              type="text"
              autoFocus
              required
              value={newFolderName}
              onChange={(e) => setNewFolderName(e.target.value)}
              placeholder="Folder name"
              aria-label="Folder name"
              className="mt-4 w-full min-h-[44px] rounded-xl border border-slate-200 bg-slate-50 px-3.5 text-sm text-slate-900 focus:border-blue-600 focus:bg-white focus:outline-none dark:border-slate-700 dark:bg-slate-950 dark:text-slate-100"
            />
            <div className="mt-5 flex justify-end gap-2">
              <button
                type="button"
                onClick={() => {
                  setNewFolderOpen(false);
                  setNewFolderName('');
                }}
                className="min-h-[40px] rounded-xl border border-slate-200 px-4 text-xs font-semibold text-slate-700 hover:bg-slate-50 dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800"
              >
                Cancel
              </button>
              <button
                type="submit"
                className="min-h-[40px] rounded-xl bg-blue-700 px-4 text-xs font-semibold text-white hover:bg-blue-800 dark:bg-blue-600"
              >
                Create
              </button>
            </div>
          </form>
        </div>
      )}

      {/* Rename Dialog */}
      {renameTarget && (
        <div
          role="dialog"
          aria-modal="true"
          aria-labelledby="rename-dialog-title"
          className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/50 p-4 backdrop-blur-xs"
        >
          <form
            onSubmit={(e) => void handleRenameSubmit(e)}
            className="w-full max-w-sm rounded-2xl border border-slate-200 bg-white p-6 dark:border-slate-800 dark:bg-slate-900"
          >
            <h2 id="rename-dialog-title" className="text-base font-bold text-slate-900 dark:text-slate-100">
              Rename {renameTarget.kind === 'file' ? 'Document' : 'Folder'}
            </h2>
            <input
              type="text"
              autoFocus
              required
              value={renameTarget.name}
              onChange={(e) =>
                setRenameTarget({...renameTarget, name: e.target.value})
              }
              aria-label="New name"
              className="mt-4 w-full min-h-[44px] rounded-xl border border-slate-200 bg-slate-50 px-3.5 text-sm text-slate-900 focus:border-blue-600 focus:bg-white focus:outline-none dark:border-slate-700 dark:bg-slate-950 dark:text-slate-100"
            />
            <div className="mt-5 flex justify-end gap-2">
              <button
                type="button"
                onClick={() => setRenameTarget(null)}
                className="min-h-[40px] rounded-xl border border-slate-200 px-4 text-xs font-semibold text-slate-700 hover:bg-slate-50 dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800"
              >
                Cancel
              </button>
              <button
                type="submit"
                className="min-h-[40px] rounded-xl bg-blue-700 px-4 text-xs font-semibold text-white hover:bg-blue-800 dark:bg-blue-600"
              >
                Save Name
              </button>
            </div>
          </form>
        </div>
      )}

      {/* Move to Folder Dialog */}
      {moveTarget && (
        <div
          role="dialog"
          aria-modal="true"
          aria-labelledby="move-dialog-title"
          className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/50 p-4 backdrop-blur-xs"
        >
          <div className="w-full max-w-sm rounded-2xl border border-slate-200 bg-white p-6 dark:border-slate-800 dark:bg-slate-900">
            <h2 id="move-dialog-title" className="text-base font-bold text-slate-900 dark:text-slate-100">
              Move &ldquo;{moveTarget.name}&rdquo;
            </h2>
            <div className="mt-4 max-h-60 space-y-1 overflow-y-auto">
              <button
                type="button"
                onClick={async () => {
                  await moveFileById(moveTarget.id, null);
                  setMoveTarget(null);
                }}
                className="flex w-full min-h-[42px] items-center gap-2.5 rounded-xl px-3 text-left text-sm font-medium text-slate-800 hover:bg-slate-100 dark:text-slate-200 dark:hover:bg-slate-800"
              >
                <Folder className="h-4 w-4 text-slate-500" />
                Root Vault
              </button>
              {folders.map((fld) => (
                <button
                  key={fld.id}
                  type="button"
                  onClick={async () => {
                    await moveFileById(moveTarget.id, fld.id);
                    setMoveTarget(null);
                  }}
                  className="flex w-full min-h-[42px] items-center gap-2.5 rounded-xl px-3 text-left text-sm font-medium text-slate-800 hover:bg-slate-100 dark:text-slate-200 dark:hover:bg-slate-800"
                >
                  <Folder className="h-4 w-4 text-slate-500" />
                  {fld.name}
                </button>
              ))}
            </div>
            <div className="mt-5 flex justify-end">
              <button
                type="button"
                onClick={() => setMoveTarget(null)}
                className="min-h-[40px] rounded-xl border border-slate-200 px-4 text-xs font-semibold text-slate-700 hover:bg-slate-50 dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800"
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Empty Trash Confirm Dialog */}
      {confirmEmptyTrash && (
        <div
          role="dialog"
          aria-modal="true"
          aria-labelledby="empty-trash-title"
          className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/50 p-4 backdrop-blur-xs"
        >
          <div className="w-full max-w-sm rounded-2xl border border-slate-200 bg-white p-6 dark:border-slate-800 dark:bg-slate-900">
            <h2 id="empty-trash-title" className="text-base font-bold text-slate-900 dark:text-slate-100">
              Empty Trash permanently?
            </h2>
            <p className="mt-2 text-sm text-slate-600 dark:text-slate-300">
              All trashed files and folders will be removed from this device and cannot be restored.
            </p>
            <div className="mt-5 flex justify-end gap-2">
              <button
                type="button"
                onClick={() => setConfirmEmptyTrash(false)}
                className="min-h-[40px] rounded-xl border border-slate-200 px-4 text-xs font-semibold text-slate-700 hover:bg-slate-50 dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={async () => {
                  await emptyAllTrash();
                  setConfirmEmptyTrash(false);
                }}
                className="min-h-[40px] rounded-xl bg-red-600 px-4 text-xs font-semibold text-white hover:bg-red-700"
              >
                Empty Trash
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
