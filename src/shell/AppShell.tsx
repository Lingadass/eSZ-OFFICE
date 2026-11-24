import React, {useEffect} from 'react';
import {Link, NavLink, useLocation, useNavigate, useSearchParams} from 'react-router-dom';
import {
  FileStack,
  FolderOpen,
  FolderTree,
  Home,
  RefreshCw,
  Search,
  Settings,
  Star,
  Trash2,
  WifiOff,
  X,
} from 'lucide-react';
import {pickLocalFiles, registerLaunchQueueConsumer} from '../core/file-io/fileSystemAccess';
import {useOnlineStatus} from '../core/pwa/useOnlineStatus';
import {useSWUpdate} from '../core/pwa/useSWUpdate';
import {useThemeStore} from '../core/theme/useThemeStore';
import type {OfficeModuleType} from '../core/types';
import {useWorkspaceStore} from '../core/workspace/useWorkspaceStore';
import {OnboardingModal} from './OnboardingModal';
import {PWAInstallButton} from './PWAInstallButton';
import {SearchPalette} from './SearchPalette';

interface AppShellProps {
  children: React.ReactNode;
  offlineReady: boolean;
}

export const AppShell: React.FC<AppShellProps> = ({children, offlineReady}) => {
  const navigate = useNavigate();
  const location = useLocation();
  const [searchParams, setSearchParams] = useSearchParams();

  const isOnline = useOnlineStatus();
  const {needRefresh, applyUpdate, dismissUpdate} = useSWUpdate();
  const {initTheme} = useThemeStore();

  const {
    tabs,
    activeTabId,
    files,
    trashedFiles,
    onboardingCompleted,
    toasts,
    bootstrap,
    completeOnboarding,
    setSearchOpen,
    createAndOpenFile,
    importFiles,
    saveTabNow,
    closeTab,
    undoTab,
    redoTab,
    dismissToast,
  } = useWorkspaceStore();

  // Initialize theme + Dexie/OPFS vault on mount
  useEffect(() => {
    const autostart = searchParams.get('autostart') === '1';
    void initTheme();
    void bootstrap(autostart);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Handle PWA manifest shortcut URLs: /?action=new&type=pdf|docs|sheets|slides
  useEffect(() => {
    const action = searchParams.get('action');
    const type = searchParams.get('type') as OfficeModuleType | null;
    if (
      action === 'new' &&
      type &&
      ['pdf', 'docs', 'sheets', 'slides'].includes(type)
    ) {
      const nextParams = new URLSearchParams(searchParams);
      nextParams.delete('action');
      nextParams.delete('type');
      setSearchParams(nextParams, {replace: true});
      void createAndOpenFile(type).then((created) => {
        if (created) navigate('/workspace');
      });
    }
  }, [searchParams, setSearchParams, createAndOpenFile, navigate]);

  // Handle PWA File Handlers launchQueue
  useEffect(() => {
    return registerLaunchQueueConsumer((launchedFiles) => {
      if (launchedFiles.length > 0) {
        void importFiles(launchedFiles, null).then(() => {
          navigate('/workspace');
        });
      }
    });
  }, [importFiles, navigate]);

  // Global keyboard shortcuts
  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      const mod = e.ctrlKey || e.metaKey;
      if (!mod) return;

      const targetTag = (e.target as HTMLElement | null)?.tagName?.toLowerCase();
      const isTextInput = targetTag === 'textarea' || targetTag === 'input';

      if (e.key.toLowerCase() === 'k') {
        e.preventDefault();
        setSearchOpen(true);
      } else if (e.key.toLowerCase() === 'o') {
        e.preventDefault();
        void pickLocalFiles().then((picked) => {
          if (picked.length > 0) {
            void importFiles(picked, null).then(() => navigate('/workspace'));
          }
        });
      } else if (e.key.toLowerCase() === 's' && activeTabId) {
        const activeTab = tabs.find((t) => t.fileId === activeTabId);
        if (activeTab?.mode === 'edit') {
          e.preventDefault();
          void saveTabNow(activeTabId);
        }
      } else if (e.key.toLowerCase() === 'w' && activeTabId && location.pathname === '/workspace') {
        e.preventDefault();
        void closeTab(activeTabId);
      } else if (!isTextInput && e.key.toLowerCase() === 'z' && activeTabId) {
        const activeTab = tabs.find((t) => t.fileId === activeTabId);
        if (activeTab?.mode === 'edit') {
          e.preventDefault();
          if (e.shiftKey) {
            redoTab(activeTabId);
          } else {
            undoTab(activeTabId);
          }
        }
      }
    };

    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, [
    tabs,
    activeTabId,
    location.pathname,
    setSearchOpen,
    importFiles,
    navigate,
    saveTabNow,
    closeTab,
    undoTab,
    redoTab,
  ]);

  const handleOpenExternal = async () => {
    const picked = await pickLocalFiles();
    if (picked.length > 0) {
      await importFiles(picked, null);
      navigate('/workspace');
    }
  };

  const navLinkClass = (active: boolean) =>
    `flex min-h-[42px] items-center justify-between gap-3 rounded-xl px-3 text-sm font-medium transition-colors ${
      active
        ? 'bg-blue-50 text-blue-800 dark:bg-blue-950/60 dark:text-blue-200'
        : 'text-slate-700 hover:bg-slate-100 dark:text-slate-300 dark:hover:bg-slate-800/70'
    }`;

  const isFilesAll = location.pathname === '/files' && !searchParams.get('view');
  const isFilesFav =
    location.pathname === '/files' && searchParams.get('view') === 'favorites';
  const isFilesTrash =
    location.pathname === '/files' && searchParams.get('view') === 'trash';

  const isWorkspaceRoute = location.pathname === '/workspace';

  return (
    <div
      className={`flex bg-slate-50 text-slate-900 dark:bg-slate-950 dark:text-slate-100 ${
        isWorkspaceRoute ? 'h-[100dvh] overflow-hidden' : 'min-h-screen'
      }`}
    >
      {/* Desktop Sidebar */}
      <aside className="hidden w-64 shrink-0 flex-col border-r border-slate-200 bg-white dark:border-slate-800 dark:bg-slate-900 lg:flex">
        <div className="flex h-16 items-center border-b border-slate-200 px-5 dark:border-slate-800">
          <Link
            to="/"
            className="inline-flex items-center rounded-lg transition-opacity hover:opacity-90"
            aria-label="eSZ OFFICE TOOLS Home"
          >
            <img
              src="/icons/icon-512.png"
              alt="eSZ OFFICE TOOLS"
              className="h-7 w-auto"
            />
          </Link>
        </div>

        <nav aria-label="Main navigation" className="flex-1 space-y-1 p-3">
          <NavLink to="/" end className={({isActive}) => navLinkClass(isActive)}>
            <span className="flex items-center gap-3">
              <Home className="h-4 w-4" />
              Home
            </span>
          </NavLink>

          <NavLink to="/workspace" className={({isActive}) => navLinkClass(isActive)}>
            <span className="flex items-center gap-3">
              <FileStack className="h-4 w-4" />
              Workspace
            </span>
            {tabs.length > 0 && (
              <span className="font-mono text-xs tabular-nums text-slate-500 dark:text-slate-400">
                {tabs.length}
              </span>
            )}
          </NavLink>

          <Link to="/files" className={navLinkClass(isFilesAll)}>
            <span className="flex items-center gap-3">
              <FolderTree className="h-4 w-4" />
              All Files
            </span>
            <span className="font-mono text-xs tabular-nums text-slate-500 dark:text-slate-400">
              {files.length}
            </span>
          </Link>

          <Link to="/files?view=favorites" className={navLinkClass(isFilesFav)}>
            <span className="flex items-center gap-3">
              <Star className="h-4 w-4" />
              Favourites
            </span>
            <span className="font-mono text-xs tabular-nums text-slate-500 dark:text-slate-400">
              {files.filter((f) => f.isFavorite).length}
            </span>
          </Link>

          <Link to="/files?view=trash" className={navLinkClass(isFilesTrash)}>
            <span className="flex items-center gap-3">
              <Trash2 className="h-4 w-4" />
              Trash
            </span>
            <span className="font-mono text-xs tabular-nums text-slate-500 dark:text-slate-400">
              {trashedFiles.length}
            </span>
          </Link>

          <NavLink to="/settings" className={({isActive}) => navLinkClass(isActive)}>
            <span className="flex items-center gap-3">
              <Settings className="h-4 w-4" />
              Settings
            </span>
          </NavLink>
        </nav>

        <div className="border-t border-slate-200 p-4 text-xs text-slate-500 dark:border-slate-800 dark:text-slate-400">
          <div className="font-medium text-slate-700 dark:text-slate-300">
            {offlineReady ? 'Offline ready' : 'Preparing offline cache'}
          </div>
          <div className="mt-0.5">Zero uploads · Device vault</div>
        </div>
      </aside>

      {/* Main Content Area */}
      <div
        className={`flex min-w-0 flex-1 flex-col ${
          isWorkspaceRoute ? 'min-h-0 overflow-hidden pb-14 lg:pb-0' : 'pb-16 lg:pb-0'
        }`}
      >
        {/* Top Bar (Desktop & Mobile) */}
        <header className="sticky top-0 z-30 flex h-12 items-center justify-between gap-4 border-b border-slate-200 bg-white/95 px-4 backdrop-blur-xs dark:border-slate-800 dark:bg-slate-900/95 lg:h-16 lg:px-6">
          <div className="flex min-w-0 items-center gap-3">
            <Link
              to="/"
              className="inline-flex items-center rounded-lg lg:hidden"
              aria-label="eSZ OFFICE TOOLS Home"
            >
              <img
                src="/icons/icon-512.png"
                alt="eSZ OFFICE TOOLS"
                className="h-6 w-auto"
              />
            </Link>
            <button
              type="button"
              onClick={() => setSearchOpen(true)}
              aria-label="Search files (Ctrl+K)"
              className="hidden min-h-[38px] w-72 items-center justify-between gap-2 rounded-xl border border-slate-200 bg-slate-50 px-3 text-xs text-slate-500 transition-colors hover:border-slate-300 hover:bg-white dark:border-slate-700 dark:bg-slate-950 dark:text-slate-400 dark:hover:border-slate-600 lg:inline-flex"
            >
              <span className="flex items-center gap-2">
                <Search className="h-3.5 w-3.5" />
                Search documents...
              </span>
              <kbd className="font-mono text-[11px] text-slate-400">Ctrl+K</kbd>
            </button>
          </div>

          <div className="flex shrink-0 items-center gap-2">
            <button
              type="button"
              onClick={() => setSearchOpen(true)}
              aria-label="Search documents"
              className="inline-flex h-10 w-10 items-center justify-center rounded-xl text-slate-600 hover:bg-slate-100 dark:text-slate-300 dark:hover:bg-slate-800 lg:hidden"
            >
              <Search className="h-4 w-4" />
            </button>
            <PWAInstallButton compact />
            <button
              type="button"
              onClick={() => void handleOpenExternal()}
              className="inline-flex min-h-[38px] items-center gap-2 whitespace-nowrap shrink-0 rounded-xl bg-blue-700 px-3.5 text-xs font-semibold text-white transition-colors hover:bg-blue-800 dark:bg-blue-600 dark:hover:bg-blue-500"
            >
              <FolderOpen className="h-4 w-4" />
              <span>Open File</span>
            </button>
          </div>
        </header>

        {/* SW Update Available Prompt */}
        {needRefresh && (
          <div
            role="status"
            className="flex flex-wrap items-center justify-between gap-3 border-b border-blue-200 bg-blue-50 px-4 py-2.5 text-xs text-blue-900 dark:border-blue-900 dark:bg-blue-950/60 dark:text-blue-100"
          >
            <span className="font-medium">Update available for eSZ OFFICE. Reload to apply the latest offline version.</span>
            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={() => void applyUpdate()}
                className="inline-flex min-h-[34px] items-center gap-1.5 rounded-lg bg-blue-700 px-3 font-semibold text-white hover:bg-blue-800"
              >
                <RefreshCw className="h-3.5 w-3.5" />
                Reload
              </button>
              <button
                type="button"
                onClick={dismissUpdate}
                aria-label="Dismiss update prompt"
                className="inline-flex h-8 w-8 items-center justify-center rounded-lg hover:bg-blue-100 dark:hover:bg-blue-900"
              >
                <X className="h-3.5 w-3.5" />
              </button>
            </div>
          </div>
        )}

        <main
          className={
            isWorkspaceRoute ? 'flex min-h-0 flex-1 flex-col overflow-hidden' : 'flex-1'
          }
        >
          {children}
        </main>
      </div>

      {/* Mobile Bottom Thumb Navigation (Height: 56px, <= 15% sticky cap with 48px top bar) */}
      <nav
        aria-label="Mobile navigation"
        className="fixed bottom-0 left-0 right-0 z-40 grid h-14 grid-cols-4 border-t border-slate-200 bg-white/95 backdrop-blur-xs dark:border-slate-800 dark:bg-slate-900/95 lg:hidden"
      >
        <NavLink
          to="/"
          end
          className={({isActive}) =>
            `flex min-h-[44px] flex-col items-center justify-center gap-0.5 text-[11px] font-medium ${
              isActive
                ? 'text-blue-700 dark:text-blue-400'
                : 'text-slate-600 dark:text-slate-400'
            }`
          }
        >
          <Home className="h-5 w-5" />
          <span>Home</span>
        </NavLink>
        <NavLink
          to="/files"
          className={({isActive}) =>
            `flex min-h-[44px] flex-col items-center justify-center gap-0.5 text-[11px] font-medium ${
              isActive
                ? 'text-blue-700 dark:text-blue-400'
                : 'text-slate-600 dark:text-slate-400'
            }`
          }
        >
          <FolderTree className="h-5 w-5" />
          <span>Files</span>
        </NavLink>
        <NavLink
          to="/workspace"
          className={({isActive}) =>
            `flex min-h-[44px] flex-col items-center justify-center gap-0.5 text-[11px] font-medium ${
              isActive
                ? 'text-blue-700 dark:text-blue-400'
                : 'text-slate-600 dark:text-slate-400'
            }`
          }
        >
          <FileStack className="h-5 w-5" />
          <span>Tabs ({tabs.length})</span>
        </NavLink>
        <NavLink
          to="/settings"
          className={({isActive}) =>
            `flex min-h-[44px] flex-col items-center justify-center gap-0.5 text-[11px] font-medium ${
              isActive
                ? 'text-blue-700 dark:text-blue-400'
                : 'text-slate-600 dark:text-slate-400'
            }`
          }
        >
          <Settings className="h-5 w-5" />
          <span>Settings</span>
        </NavLink>
      </nav>

      {/* Non-intrusive Offline Indicator */}
      {!isOnline && (
        <div
          role="status"
          className="fixed bottom-16 left-4 z-40 flex items-center gap-2 rounded-xl border border-amber-300 bg-amber-50 px-3 py-2 text-xs font-medium text-amber-900 shadow-sm dark:border-amber-800 dark:bg-amber-950 dark:text-amber-100 lg:bottom-4"
        >
          <WifiOff className="h-3.5 w-3.5" />
          <span>Offline Mode · Local device vault active</span>
        </div>
      )}

      {/* Toast Stack */}
      <div
        aria-live="polite"
        className="fixed bottom-16 right-4 z-50 flex max-w-sm flex-col gap-2 lg:bottom-4"
      >
        {toasts.map((toast) => (
          <div
            key={toast.id}
            className="flex items-center justify-between gap-3 rounded-xl border border-slate-200 bg-white px-4 py-3 text-xs font-medium text-slate-900 shadow-lg dark:border-slate-700 dark:bg-slate-800 dark:text-slate-100"
          >
            <span>{toast.message}</span>
            <button
              type="button"
              onClick={() => dismissToast(toast.id)}
              aria-label="Dismiss notification"
              className="text-slate-400 hover:text-slate-700 dark:hover:text-slate-200"
            >
              <X className="h-3.5 w-3.5" />
            </button>
          </div>
        ))}
      </div>

      {/* Onboarding Modal (3 screens) */}
      <OnboardingModal open={!onboardingCompleted} onComplete={() => void completeOnboarding()} />

      {/* Ctrl+K Search Palette */}
      <SearchPalette />
    </div>
  );
};
