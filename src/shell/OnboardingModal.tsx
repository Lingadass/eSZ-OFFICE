import React, {useState} from 'react';
import {ArrowLeft, ArrowRight, Check, FolderLock, HardDriveDownload, Layers, X} from 'lucide-react';

interface OnboardingModalProps {
  open: boolean;
  onComplete: () => void;
}

const SLIDES = [
  {
    title: 'Private, Zero-Upload Office Suite',
    description:
      'Every PDF, Document, Spreadsheet, and Presentation is processed directly on your device. No accounts, no tracking, and no cloud uploads.',
    icon: FolderLock,
    bullets: [
      'Works completely offline once loaded',
      'Files stay inside your browser IndexedDB and OPFS vault',
      'Free forever with zero ads',
    ],
  },
  {
    title: 'Four Integrated Workspaces',
    description:
      'Create, organize, and switch seamlessly between multiple open tabs across PDF, Docs, Sheets, and Slides in one unified interface.',
    icon: Layers,
    bullets: [
      'Autosaves every 5 seconds with crash recovery',
      'Full local folder hierarchy, favourites, and recoverable Trash',
      'Keyboard shortcuts and touch-ready mobile navigation',
    ],
  },
  {
    title: 'Installable & Portable Backups',
    description:
      'Install eSZ OFFICE to your home screen or desktop for standalone launch, and export your entire workspace as a standard .zip archive anytime.',
    icon: HardDriveDownload,
    bullets: [
      'Uses native File System Access API where supported',
      'Download & upload fallback on all browsers and WebViews',
      'One-click .zip backup and restore in Settings',
    ],
  },
];

export const OnboardingModal: React.FC<OnboardingModalProps> = ({open, onComplete}) => {
  const [step, setStep] = useState(0);

  if (!open) return null;

  const current = SLIDES[step];
  const Icon = current.icon;
  const isLast = step === SLIDES.length - 1;

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="onboarding-heading"
      className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/60 p-4 backdrop-blur-xs"
    >
      <div className="w-full max-w-lg rounded-2xl border border-slate-200 bg-white p-6 dark:border-slate-800 dark:bg-slate-900 sm:p-8">
        <div className="flex items-center justify-between">
          <span className="text-xs font-semibold text-slate-500 tabular-nums dark:text-slate-400">
            Step {step + 1} of {SLIDES.length}
          </span>
          <button
            type="button"
            onClick={onComplete}
            aria-label="Skip onboarding tour"
            className="inline-flex min-h-[36px] items-center gap-1 rounded-lg px-2.5 text-xs font-medium text-slate-600 hover:bg-slate-100 dark:text-slate-300 dark:hover:bg-slate-800"
          >
            <span>Skip</span>
            <X className="h-3.5 w-3.5" />
          </button>
        </div>

        <div className="mt-6 flex h-12 w-12 items-center justify-center rounded-xl bg-blue-50 text-blue-700 dark:bg-blue-950 dark:text-blue-300">
          <Icon className="h-6 w-6" />
        </div>

        <h2
          id="onboarding-heading"
          className="mt-5 text-xl font-bold tracking-tight text-slate-900 dark:text-slate-100"
        >
          {current.title}
        </h2>
        <p className="mt-2.5 text-sm leading-relaxed text-slate-600 dark:text-slate-300">
          {current.description}
        </p>

        <ul className="mt-5 space-y-2.5 border-t border-slate-200 pt-4 text-sm text-slate-700 dark:border-slate-800 dark:text-slate-200">
          {current.bullets.map((b) => (
            <li key={b} className="flex items-start gap-2.5">
              <Check className="mt-0.5 h-4 w-4 shrink-0 text-blue-700 dark:text-blue-400" />
              <span>{b}</span>
            </li>
          ))}
        </ul>

        <div className="mt-8 flex items-center justify-between gap-3">
          <div className="flex items-center gap-1.5" aria-hidden="true">
            {SLIDES.map((_, idx) => (
              <span
                key={idx}
                className={`h-1.5 rounded-full transition-all ${
                  idx === step
                    ? 'w-6 bg-blue-700 dark:bg-blue-500'
                    : 'w-2 bg-slate-200 dark:bg-slate-700'
                }`}
              />
            ))}
          </div>

          <div className="flex items-center gap-2">
            {step > 0 && (
              <button
                type="button"
                onClick={() => setStep((s) => s - 1)}
                className="inline-flex min-h-[44px] items-center gap-1.5 rounded-xl border border-slate-200 px-4 py-2 text-sm font-medium text-slate-700 hover:bg-slate-50 dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800"
              >
                <ArrowLeft className="h-4 w-4" />
                Back
              </button>
            )}
            {isLast ? (
              <button
                type="button"
                onClick={onComplete}
                className="inline-flex min-h-[44px] items-center gap-1.5 rounded-xl bg-blue-700 px-5 py-2 text-sm font-semibold text-white hover:bg-blue-800 dark:bg-blue-600 dark:hover:bg-blue-500"
              >
                Start Using eSZ OFFICE
              </button>
            ) : (
              <button
                type="button"
                onClick={() => setStep((s) => s + 1)}
                className="inline-flex min-h-[44px] items-center gap-1.5 rounded-xl bg-blue-700 px-5 py-2 text-sm font-semibold text-white hover:bg-blue-800 dark:bg-blue-600 dark:hover:bg-blue-500"
              >
                Next
                <ArrowRight className="h-4 w-4" />
              </button>
            )}
          </div>
        </div>
      </div>
    </div>
  );
};
