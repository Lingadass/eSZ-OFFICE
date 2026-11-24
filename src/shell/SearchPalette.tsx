import React, {useEffect, useMemo, useRef, useState} from 'react';
import {useNavigate} from 'react-router-dom';
import {Folder, Search, X} from 'lucide-react';
import {filterFilesByQuery, formatBytes} from '../core/storage/fileRules';
import {useWorkspaceStore} from '../core/workspace/useWorkspaceStore';
import {ModuleIcon} from './FileThumb';

export const SearchPalette: React.FC = () => {
  const navigate = useNavigate();
  const inputRef = useRef<HTMLInputElement>(null);
  const {searchOpen, setSearchOpen, files, folders, openFileInTab} = useWorkspaceStore();
  const [query, setQuery] = useState('');

  useEffect(() => {
    if (searchOpen) {
      setQuery('');
      setTimeout(() => inputRef.current?.focus(), 20);
    }
  }, [searchOpen]);

  const matchedFiles = useMemo(
    () => filterFilesByQuery(files, query).slice(0, 8),
    [files, query],
  );
  const matchedFolders = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return folders.slice(0, 3);
    return folders.filter((f) => f.name.toLowerCase().includes(q)).slice(0, 4);
  }, [folders, query]);

  if (!searchOpen) return null;

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-label="Search files and folders"
      className="fixed inset-0 z-50 flex items-start justify-center bg-slate-950/50 p-4 pt-16 backdrop-blur-xs"
      onClick={(e) => {
        if (e.target === e.currentTarget) setSearchOpen(false);
      }}
    >
      <div className="w-full max-w-xl overflow-hidden rounded-2xl border border-slate-200 bg-white dark:border-slate-800 dark:bg-slate-900">
        <div className="flex items-center gap-3 border-b border-slate-200 px-4 py-3 dark:border-slate-800">
          <Search className="h-4 w-4 text-slate-400" />
          <input
            ref={inputRef}
            type="search"
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Escape') setSearchOpen(false);
            }}
            placeholder="Search documents, spreadsheets, PDFs, or folders..."
            aria-label="Search local files"
            className="w-full bg-transparent text-sm text-slate-900 placeholder:text-slate-400 focus:outline-none dark:text-slate-100"
          />
          <button
            type="button"
            onClick={() => setSearchOpen(false)}
            aria-label="Close search"
            className="inline-flex h-8 w-8 items-center justify-center rounded-lg text-slate-500 hover:bg-slate-100 dark:hover:bg-slate-800"
          >
            <X className="h-4 w-4" />
          </button>
        </div>

        <div className="max-h-96 overflow-y-auto p-2">
          {matchedFiles.length === 0 && matchedFolders.length === 0 ? (
            <div className="px-4 py-10 text-center text-sm text-slate-500 dark:text-slate-400">
              No matching files or folders found on this device.
            </div>
          ) : (
            <>
              {matchedFolders.length > 0 && (
                <div className="mb-2">
                  <div className="px-3 py-1.5 text-xs font-semibold text-slate-500 dark:text-slate-400">
                    Folders
                  </div>
                  {matchedFolders.map((folder) => (
                    <button
                      key={folder.id}
                      type="button"
                      onClick={() => {
                        setSearchOpen(false);
                        navigate(`/files?folder=${encodeURIComponent(folder.id)}`);
                      }}
                      className="flex w-full min-h-[44px] items-center gap-3 rounded-xl px-3 py-2 text-left text-sm text-slate-800 hover:bg-slate-100 dark:text-slate-200 dark:hover:bg-slate-800"
                    >
                      <Folder className="h-4 w-4 text-slate-500" />
                      <span className="truncate font-medium">{folder.name}</span>
                    </button>
                  ))}
                </div>
              )}

              {matchedFiles.length > 0 && (
                <div>
                  <div className="px-3 py-1.5 text-xs font-semibold text-slate-500 dark:text-slate-400">
                    Documents
                  </div>
                  {matchedFiles.map((file) => (
                    <button
                      key={file.id}
                      type="button"
                      onClick={async () => {
                        setSearchOpen(false);
                        await openFileInTab(file.id);
                        navigate('/workspace');
                      }}
                      className="flex w-full min-h-[48px] items-center justify-between gap-3 rounded-xl px-3 py-2 text-left hover:bg-slate-100 dark:hover:bg-slate-800"
                    >
                      <div className="flex min-w-0 items-center gap-3">
                        <ModuleIcon type={file.moduleType} size="sm" />
                        <div className="min-w-0">
                          <div className="truncate text-sm font-medium text-slate-900 dark:text-slate-100">
                            {file.name}
                          </div>
                          <div className="text-xs text-slate-500 tabular-nums dark:text-slate-400">
                            {file.moduleType.toUpperCase()} · {formatBytes(file.size)}
                          </div>
                        </div>
                      </div>
                    </button>
                  ))}
                </div>
              )}
            </>
          )}
        </div>
      </div>
    </div>
  );
};
