import React, {useRef, useState} from 'react';
import {Link} from 'react-router-dom';
import {
  Download,
  HardDrive,
  Monitor,
  Moon,
  RotateCcw,
  ShieldCheck,
  Sun,
  Upload,
} from 'lucide-react';
import {exportWorkspaceBackupZip, importWorkspaceBackupZip} from '../core/file-io/backupZip';
import {saveBlobToDevice} from '../core/file-io/fileSystemAccess';
import {formatBytes} from '../core/storage/fileRules';
import {requestPersistentStorage} from '../core/storage/opfs';
import {useThemeStore} from '../core/theme/useThemeStore';
import type {ThemePreference} from '../core/types';
import {useWorkspaceStore} from '../core/workspace/useWorkspaceStore';
import {StorageStats} from './StorageStats';

export const SettingsView: React.FC = () => {
  const {preference, setTheme} = useThemeStore();
  const {
    storageStats,
    defaultSaveMode,
    setDefaultSaveMode,
    defaultOpenMode,
    setDefaultOpenMode,
    refreshLibrary,
    replayOnboarding,
    pushToast,
  } = useWorkspaceStore();

  const [busyBackup, setBusyBackup] = useState(false);
  const [persisting, setPersisting] = useState(false);
  const restoreInputRef = useRef<HTMLInputElement>(null);

  const handleRequestPersist = async () => {
    setPersisting(true);
    const granted = await requestPersistentStorage();
    await refreshLibrary();
    setPersisting(false);
    pushToast(
      granted
        ? 'Persistent device storage granted by browser.'
        : 'Browser kept best-effort storage. Export regular .zip backups to protect files.',
      granted ? 'success' : 'info',
    );
  };

  const handleExportBackup = async () => {
    setBusyBackup(true);
    try {
      const {blob, filename, fileCount} = await exportWorkspaceBackupZip();
      await saveBlobToDevice(blob, filename, 'docs');
      pushToast(`Exported ${fileCount} files to ${filename}`, 'success');
    } catch {
      pushToast('Failed to generate backup archive.', 'error');
    } finally {
      setBusyBackup(false);
    }
  };

  const handleImportBackup = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file) return;
    setBusyBackup(true);
    try {
      const res = await importWorkspaceBackupZip(file);
      await refreshLibrary();
      pushToast(
        `Restored ${res.importedFiles} files and ${res.importedFolders} folders.`,
        'success',
      );
    } catch (err) {
      pushToast(
        err instanceof Error ? err.message : 'Could not import backup archive.',
        'error',
      );
    } finally {
      setBusyBackup(false);
    }
  };

  const themeOptions: Array<{value: ThemePreference; label: string; icon: React.FC<{className?: string}>}> = [
    {value: 'light', label: 'Light', icon: Sun},
    {value: 'dark', label: 'Dark', icon: Moon},
    {value: 'system', label: 'System', icon: Monitor},
  ];

  return (
    <div className="mx-auto w-full max-w-4xl space-y-6 p-4 sm:p-6 lg:p-8">
      <div>
        <h1 className="text-2xl font-bold tracking-tight text-slate-900 dark:text-slate-100">
          Settings
        </h1>
        <p className="mt-1 text-sm text-slate-600 dark:text-slate-400">
          Configure appearance, local device storage, backups, and legal information.
        </p>
      </div>

      {/* Appearance */}
      <section className="rounded-2xl border border-slate-200 bg-white p-5 dark:border-slate-800 dark:bg-slate-900">
        <h2 className="text-base font-bold text-slate-900 dark:text-slate-100">Theme</h2>
        <p className="mt-1 text-xs text-slate-500 dark:text-slate-400">
          Choose light, dark, or match your device system setting.
        </p>
        <div className="mt-4 grid grid-cols-3 gap-2.5 sm:max-w-md">
          {themeOptions.map((opt) => {
            const Icon = opt.icon;
            const active = preference === opt.value;
            return (
              <button
                key={opt.value}
                type="button"
                onClick={() => void setTheme(opt.value)}
                aria-pressed={active}
                className={`flex min-h-[44px] items-center justify-center gap-2 rounded-xl border px-3 text-xs font-semibold transition-colors ${
                  active
                    ? 'border-blue-600 bg-blue-50 text-blue-800 dark:border-blue-500 dark:bg-blue-950/60 dark:text-blue-200'
                    : 'border-slate-200 bg-white text-slate-700 hover:bg-slate-50 dark:border-slate-700 dark:bg-slate-800 dark:text-slate-200 dark:hover:bg-slate-700'
                }`}
              >
                <Icon className="h-4 w-4" />
                {opt.label}
              </button>
            );
          })}
        </div>
      </section>

      {/* Opening Existing Files Mode */}
      <section className="rounded-2xl border border-slate-200 bg-white p-5 dark:border-slate-800 dark:bg-slate-900">
        <h2 className="text-base font-bold text-slate-900 dark:text-slate-100">
          Open Existing Files In
        </h2>
        <p className="mt-1 text-xs text-slate-500 dark:text-slate-400">
          Choose whether existing documents open in clean read-only View mode or directly in Edit mode. Newly created blank files always open in Edit mode.
        </p>
        <div className="mt-4 grid gap-2.5 sm:grid-cols-2">
          <label
            className={`flex cursor-pointer items-start gap-3 rounded-xl border p-3.5 transition-colors ${
              defaultOpenMode === 'view'
                ? 'border-blue-600 bg-blue-50/60 dark:border-blue-500 dark:bg-blue-950/40'
                : 'border-slate-200 dark:border-slate-800'
            }`}
          >
            <input
              type="radio"
              name="default-open-mode"
              aria-label="Open existing files in View (default)"
              checked={defaultOpenMode === 'view'}
              onChange={() => void setDefaultOpenMode('view')}
              className="mt-1 h-4 w-4 accent-blue-700"
            />
            <div>
              <div className="text-sm font-semibold text-slate-900 dark:text-slate-100">
                View (default)
              </div>
              <div className="text-xs text-slate-500 dark:text-slate-400">
                Distraction-free read-only viewer. Press Edit or E whenever you want to make changes.
              </div>
            </div>
          </label>
          <label
            className={`flex cursor-pointer items-start gap-3 rounded-xl border p-3.5 transition-colors ${
              defaultOpenMode === 'edit'
                ? 'border-blue-600 bg-blue-50/60 dark:border-blue-500 dark:bg-blue-950/40'
                : 'border-slate-200 dark:border-slate-800'
            }`}
          >
            <input
              type="radio"
              name="default-open-mode"
              aria-label="Open existing files in Edit"
              checked={defaultOpenMode === 'edit'}
              onChange={() => void setDefaultOpenMode('edit')}
              className="mt-1 h-4 w-4 accent-blue-700"
            />
            <div>
              <div className="text-sm font-semibold text-slate-900 dark:text-slate-100">
                Edit
              </div>
              <div className="text-xs text-slate-500 dark:text-slate-400">
                Open directly with the full editing ribbon and tools active.
              </div>
            </div>
          </label>
        </div>
      </section>

      {/* Default Save Location */}
      <section className="rounded-2xl border border-slate-200 bg-white p-5 dark:border-slate-800 dark:bg-slate-900">
        <h2 className="text-base font-bold text-slate-900 dark:text-slate-100">
          Default Save Behaviour
        </h2>
        <p className="mt-1 text-xs text-slate-500 dark:text-slate-400">
          Control how documents are saved when editing or exporting.
        </p>
        <div className="mt-4 space-y-2.5">
          <label className="flex cursor-pointer items-start gap-3 rounded-xl border border-slate-200 p-3.5 dark:border-slate-800">
            <input
              type="radio"
              name="save-mode"
              checked={defaultSaveMode === 'device-vault'}
              onChange={() => void setDefaultSaveMode('device-vault')}
              className="mt-1 h-4 w-4 accent-blue-700"
            />
            <div>
              <div className="text-sm font-semibold text-slate-900 dark:text-slate-100">
                On-Device Vault (OPFS + IndexedDB)
              </div>
              <div className="text-xs text-slate-500 dark:text-slate-400">
                Autosaves every 5 seconds locally with instant offline reopening.
              </div>
            </div>
          </label>
          <label className="flex cursor-pointer items-start gap-3 rounded-xl border border-slate-200 p-3.5 dark:border-slate-800">
            <input
              type="radio"
              name="save-mode"
              checked={defaultSaveMode === 'ask-save-as'}
              onChange={() => void setDefaultSaveMode('ask-save-as')}
              className="mt-1 h-4 w-4 accent-blue-700"
            />
            <div>
              <div className="text-sm font-semibold text-slate-900 dark:text-slate-100">
                Vault + Prompt File System Save As
              </div>
              <div className="text-xs text-slate-500 dark:text-slate-400">
                Keeps local autosave active and opens the native save dialog on export.
              </div>
            </div>
          </label>
        </div>
      </section>

      {/* Storage & Backup */}
      <section className="rounded-2xl border border-slate-200 bg-white p-5 dark:border-slate-800 dark:bg-slate-900">
        <div className="flex flex-wrap items-start justify-between gap-4">
          <div>
            <h2 className="text-base font-bold text-slate-900 dark:text-slate-100">
              Storage Usage &amp; Local Backups
            </h2>
            <p className="mt-1 text-xs text-slate-500 dark:text-slate-400">
              All files remain strictly on your device. Persistent storage protects data from automatic browser eviction.
            </p>
          </div>
          <span className="inline-flex items-center gap-1.5 text-xs font-semibold text-slate-700 dark:text-slate-300">
            <ShieldCheck className="h-4 w-4 text-blue-700 dark:text-blue-400" />
            {storageStats?.persisted ? 'Persistent Storage Active' : 'Best-Effort Storage'}
          </span>
        </div>

        <div className="mt-4 rounded-xl border border-slate-200 bg-slate-50 p-4 dark:border-slate-800 dark:bg-slate-950">
          <div className="flex items-center justify-between text-xs font-medium text-slate-700 tabular-nums dark:text-slate-300">
            <span className="inline-flex items-center gap-2">
              <HardDrive className="h-4 w-4 text-slate-500" />
              {storageStats ? formatBytes(storageStats.usageBytes) : '0 B'} used
              {storageStats && storageStats.quotaBytes > 0
                ? ` of ${formatBytes(storageStats.quotaBytes)}`
                : ''}
            </span>
            <span>{storageStats?.percentUsed ?? 0}%</span>
          </div>
          <div className="mt-2 h-2 w-full overflow-hidden rounded-full bg-slate-200 dark:bg-slate-800">
            <div
              className="h-full rounded-full bg-blue-700 transition-all dark:bg-blue-500"
              style={{width: `${Math.max(2, storageStats?.percentUsed ?? 2)}%`}}
            />
          </div>
          <div className="mt-2.5 flex flex-wrap gap-x-4 gap-y-1 text-xs text-slate-500 tabular-nums dark:text-slate-400">
            <span>Files: {storageStats?.fileCount ?? 0}</span>
            <span>·</span>
            <span>Trash: {storageStats?.trashedCount ?? 0}</span>
            <span>·</span>
            <span>
              Blob engine: {storageStats?.opfsSupported ? 'OPFS + IndexedDB' : 'IndexedDB'}
            </span>
          </div>
        </div>

        <div className="mt-5">
          <StorageStats embedded />
        </div>

        <div className="mt-4 flex flex-wrap gap-2.5">
          {!storageStats?.persisted && (
            <button
              type="button"
              disabled={persisting}
              onClick={() => void handleRequestPersist()}
              className="inline-flex min-h-[42px] items-center gap-2 rounded-xl border border-slate-200 bg-white px-4 text-xs font-semibold text-slate-800 hover:bg-slate-50 disabled:opacity-50 dark:border-slate-700 dark:bg-slate-800 dark:text-slate-100 dark:hover:bg-slate-700"
            >
              <ShieldCheck className="h-4 w-4 text-blue-700 dark:text-blue-400" />
              Request Persistent Storage
            </button>
          )}
          <button
            type="button"
            disabled={busyBackup}
            onClick={() => void handleExportBackup()}
            className="inline-flex min-h-[42px] items-center gap-2 rounded-xl bg-blue-700 px-4 text-xs font-semibold text-white hover:bg-blue-800 disabled:opacity-50 dark:bg-blue-600"
          >
            <Download className="h-4 w-4" />
            Export Backup (.zip)
          </button>
          <button
            type="button"
            disabled={busyBackup}
            onClick={() => restoreInputRef.current?.click()}
            className="inline-flex min-h-[42px] items-center gap-2 rounded-xl border border-slate-200 bg-white px-4 text-xs font-semibold text-slate-800 hover:bg-slate-50 disabled:opacity-50 dark:border-slate-700 dark:bg-slate-800 dark:text-slate-100 dark:hover:bg-slate-700"
          >
            <Upload className="h-4 w-4" />
            Import Backup (.zip)
          </button>
          <input
            ref={restoreInputRef}
            type="file"
            accept=".zip,application/zip"
            onChange={(e) => void handleImportBackup(e)}
            className="hidden"
          />
        </div>
      </section>

      {/* About, Tour, Privacy & Licences */}
      <section className="rounded-2xl border border-slate-200 bg-white p-5 dark:border-slate-800 dark:bg-slate-900">
        <h2 className="text-base font-bold text-slate-900 dark:text-slate-100">
          About eSZ OFFICE &amp; Legal
        </h2>
        <p className="mt-1 text-xs leading-relaxed text-slate-600 dark:text-slate-400">
          eSZ OFFICE is a free, offline-first office suite with zero ads, zero telemetry, and zero server uploads.
        </p>

        <div className="mt-4 flex flex-wrap gap-3">
          <button
            type="button"
            onClick={replayOnboarding}
            className="inline-flex min-h-[40px] items-center gap-2 rounded-xl border border-slate-200 px-3.5 text-xs font-semibold text-slate-700 hover:bg-slate-50 dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800"
          >
            <RotateCcw className="h-3.5 w-3.5" />
            Replay Onboarding Tour
          </button>
          <Link
            to="/privacy-policy"
            className="inline-flex min-h-[40px] items-center rounded-xl border border-slate-200 px-3.5 text-xs font-semibold text-slate-700 hover:bg-slate-50 dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800"
          >
            Privacy Policy
          </Link>
          <Link
            to="/terms"
            className="inline-flex min-h-[40px] items-center rounded-xl border border-slate-200 px-3.5 text-xs font-semibold text-slate-700 hover:bg-slate-50 dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800"
          >
            Terms of Use
          </Link>
          <Link
            to="/licenses"
            className="inline-flex min-h-[40px] items-center rounded-xl border border-slate-200 px-3.5 text-xs font-semibold text-slate-700 hover:bg-slate-50 dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800"
          >
            Open-Source Licences
          </Link>
        </div>
      </section>
    </div>
  );
};
