import React from 'react';
import {Link} from 'react-router-dom';
import {ArrowLeft} from 'lucide-react';

const LegalFrame: React.FC<{title: string; updated: string; children: React.ReactNode}> = ({
  title,
  updated,
  children,
}) => (
  <div className="mx-auto w-full max-w-3xl p-4 sm:p-6 lg:p-8">
    <Link
      to="/settings"
      className="inline-flex min-h-[40px] items-center gap-1.5 text-xs font-semibold text-blue-700 hover:underline dark:text-blue-400"
    >
      <ArrowLeft className="h-3.5 w-3.5" />
      Back to Settings
    </Link>
    <article className="mt-4 rounded-2xl border border-slate-200 bg-white p-6 dark:border-slate-800 dark:bg-slate-900 sm:p-8">
      <h1 className="text-2xl font-bold tracking-tight text-slate-900 dark:text-slate-100">
        {title}
      </h1>
      <p className="mt-1 text-xs text-slate-500 dark:text-slate-400">
        Effective {updated}
      </p>
      <div className="mt-6 space-y-5 text-sm leading-relaxed text-slate-700 dark:text-slate-300">
        {children}
      </div>
    </article>
  </div>
);

export const PrivacyPolicyView: React.FC = () => (
  <LegalFrame title="Privacy Policy" updated="November 2026">
    <section>
      <h2 className="text-base font-bold text-slate-900 dark:text-slate-100">
        1. Zero Data Collection
      </h2>
      <p className="mt-1.5">
        eSZ OFFICE does not collect, transmit, store, sell, or share any personal information, usage telemetry, crash logs, advertising identifiers, or document contents. No account or sign-in is required.
      </p>
    </section>
    <section>
      <h2 className="text-base font-bold text-slate-900 dark:text-slate-100">
        2. On-Device Processing Only
      </h2>
      <p className="mt-1.5">
        All PDF rendering, document editing, spreadsheet calculations, presentation tools, file conversions, and backups execute strictly on your device inside your browser or Android Trusted Web Activity container. Your files never leave your device unless you explicitly export or share them using your operating system&apos;s share/save tools.
      </p>
    </section>
    <section>
      <h2 className="text-base font-bold text-slate-900 dark:text-slate-100">
        3. Local Storage
      </h2>
      <p className="mt-1.5">
        Documents, folders, settings, and recovery snapshots are stored exclusively in your browser&apos;s local Origin Private File System (OPFS) and IndexedDB storage. Clearing site data or uninstalling the application removes this local vault unless you exported a `.zip` backup.
      </p>
    </section>
    <section>
      <h2 className="text-base font-bold text-slate-900 dark:text-slate-100">
        4. Third-Party Services &amp; Ads
      </h2>
      <p className="mt-1.5">
        eSZ OFFICE contains no advertisements, no analytics SDKs, and no third-party tracking scripts.
      </p>
    </section>
  </LegalFrame>
);

export const TermsView: React.FC = () => (
  <LegalFrame title="Terms of Use" updated="November 2026">
    <section>
      <h2 className="text-base font-bold text-slate-900 dark:text-slate-100">
        1. Free Offline Productivity Software
      </h2>
      <p className="mt-1.5">
        eSZ OFFICE is provided free of charge for personal, educational, and commercial productivity without ads or subscriptions.
      </p>
    </section>
    <section>
      <h2 className="text-base font-bold text-slate-900 dark:text-slate-100">
        2. Your Files &amp; Backups
      </h2>
      <p className="mt-1.5">
        You retain 100% ownership and control over every file you create or open in eSZ OFFICE. Because eSZ OFFICE has no cloud servers, you are responsible for keeping regular backups (for example via Settings → Export Backup) before clearing browser storage or resetting your device.
      </p>
    </section>
    <section>
      <h2 className="text-base font-bold text-slate-900 dark:text-slate-100">
        3. Warranty Disclaimer
      </h2>
      <p className="mt-1.5">
        The application is provided &ldquo;as is&rdquo; without warranties of any kind. Complex proprietary desktop formats (such as macros or SmartArt) may render differently in browser-based modules.
      </p>
    </section>
  </LegalFrame>
);

const LICENSE_ROWS = [
  {pkg: 'react / react-dom', role: 'UI component runtime', license: 'MIT'},
  {pkg: 'react-router-dom', role: 'Client-side routing', license: 'MIT'},
  {pkg: 'zustand', role: 'Workspace & theme state', license: 'MIT'},
  {pkg: 'dexie', role: 'IndexedDB structured vault', license: 'Apache-2.0'},
  {pkg: 'jszip', role: 'Local .zip backup export/import', license: 'MIT (dual MIT / GPL-3.0)'},
  {pkg: 'lucide-react', role: 'Vector interface icons', license: 'ISC'},
  {pkg: 'motion', role: 'UI motion primitives', license: 'MIT'},
  {pkg: 'vite-plugin-pwa / workbox-window', role: 'Offline service worker & manifest', license: 'MIT'},
  {pkg: 'tailwindcss', role: 'Styling system', license: 'MIT'},
];

export const LicensesView: React.FC = () => (
  <LegalFrame title="Open-Source Licences" updated="November 2026">
    <p>
      eSZ OFFICE is built exclusively with permissive open-source packages (MIT, Apache-2.0, BSD, ISC). The complete table including planned Phase 1–4 modules is maintained in{' '}
      <code className="rounded bg-slate-100 px-1.5 py-0.5 text-xs dark:bg-slate-800">
        /docs/LICENSES.md
      </code>
      .
    </p>
    <div className="overflow-x-auto rounded-xl border border-slate-200 dark:border-slate-800">
      <table className="w-full text-left text-xs">
        <thead className="bg-slate-50 text-slate-600 dark:bg-slate-950 dark:text-slate-400">
          <tr>
            <th className="px-4 py-3 font-semibold">Package</th>
            <th className="px-4 py-3 font-semibold">Purpose</th>
            <th className="px-4 py-3 font-semibold">Licence</th>
          </tr>
        </thead>
        <tbody className="divide-y divide-slate-100 dark:divide-slate-800">
          {LICENSE_ROWS.map((row) => (
            <tr key={row.pkg}>
              <td className="px-4 py-2.5 font-mono font-medium text-slate-900 dark:text-slate-100">
                {row.pkg}
              </td>
              <td className="px-4 py-2.5 text-slate-600 dark:text-slate-300">{row.role}</td>
              <td className="px-4 py-2.5 font-mono text-slate-700 dark:text-slate-200">
                {row.license}
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  </LegalFrame>
);
