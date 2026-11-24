import {create} from 'zustand';
import {getSetting, setSetting} from '../storage/fileRepository';
import type {ThemePreference} from '../types';

interface ThemeState {
  preference: ThemePreference;
  resolved: 'light' | 'dark';
  initialized: boolean;
  initTheme: () => Promise<void>;
  setTheme: (pref: ThemePreference) => Promise<void>;
}

function resolveSystemTheme(): 'light' | 'dark' {
  if (typeof window === 'undefined' || !window.matchMedia) return 'light';
  return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
}

function applyThemeToDom(resolved: 'light' | 'dark'): void {
  if (typeof document === 'undefined') return;
  const root = document.documentElement;
  if (resolved === 'dark') {
    root.classList.add('dark');
  } else {
    root.classList.remove('dark');
  }
  const metaTheme = document.querySelector('meta[name="theme-color"]');
  if (metaTheme) {
    metaTheme.setAttribute('content', resolved === 'dark' ? '#0B1120' : '#1D4ED8');
  }
}

export const useThemeStore = create<ThemeState>((set, get) => ({
  preference: 'system',
  resolved: 'light',
  initialized: false,

  initTheme: async () => {
    const saved = await getSetting<string>('theme', 'system');
    const preference: ThemePreference =
      saved === 'light' || saved === 'dark' || saved === 'system' ? saved : 'system';
    const resolved = preference === 'system' ? resolveSystemTheme() : preference;
    applyThemeToDom(resolved);
    set({preference, resolved, initialized: true});

    if (typeof window !== 'undefined' && window.matchMedia) {
      const mql = window.matchMedia('(prefers-color-scheme: dark)');
      const onChange = () => {
        if (get().preference === 'system') {
          const next = resolveSystemTheme();
          applyThemeToDom(next);
          set({resolved: next});
        }
      };
      mql.addEventListener('change', onChange);
    }
  },

  setTheme: async (preference: ThemePreference) => {
    const resolved = preference === 'system' ? resolveSystemTheme() : preference;
    applyThemeToDom(resolved);
    set({preference, resolved});
    await setSetting('theme', preference);
  },
}));
