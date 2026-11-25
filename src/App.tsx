import {Analytics} from '@vercel/analytics/react';
import {BrowserRouter, Navigate, Route, Routes} from 'react-router-dom';
import {useSWUpdate} from './core/pwa/useSWUpdate';
import {AppShell} from './shell/AppShell';
import {ErrorBoundary} from './shell/ErrorBoundary';
import {FileManagerView} from './shell/FileManagerView';
import {HomeView} from './shell/HomeView';
import {LicensesView, PrivacyPolicyView, TermsView} from './shell/LegalViews';
import {SettingsView} from './shell/SettingsView';
import {WorkspaceView} from './shell/WorkspaceView';

export default function App() {
  const {offlineReady} = useSWUpdate();

  return (
    <ErrorBoundary>
      <BrowserRouter>
        <AppShell offlineReady={offlineReady}>
          <Routes>
            <Route path="/" element={<HomeView offlineReady={offlineReady} />} />
            <Route path="/files" element={<FileManagerView />} />
            <Route path="/workspace" element={<WorkspaceView />} />
            <Route path="/settings" element={<SettingsView />} />
            <Route path="/privacy-policy" element={<PrivacyPolicyView />} />
            <Route path="/terms" element={<TermsView />} />
            <Route path="/licenses" element={<LicensesView />} />
            <Route path="*" element={<Navigate to="/" replace />} />
          </Routes>
        </AppShell>
        <Analytics />
      </BrowserRouter>
    </ErrorBoundary>
  );
}
