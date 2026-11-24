import {useCallback, useEffect, useRef, useState} from 'react';
import {registerSW} from 'virtual:pwa-register';

export function useSWUpdate() {
  const [needRefresh, setNeedRefresh] = useState(false);
  const [offlineReady, setOfflineReady] = useState(false);
  const updateFnRef = useRef<((reloadPage?: boolean) => Promise<void>) | null>(null);

  useEffect(() => {
    if (typeof window === 'undefined' || !('serviceWorker' in navigator)) {
      return;
    }
    try {
      const updateSW = registerSW({
        immediate: true,
        onNeedRefresh() {
          setNeedRefresh(true);
        },
        onOfflineReady() {
          setOfflineReady(true);
        },
        onRegisteredSW(_swUrl, registration) {
          if (registration?.active) {
            setOfflineReady(true);
          }
        },
      });
      updateFnRef.current = updateSW;

      navigator.serviceWorker.ready
        .then(() => setOfflineReady(true))
        .catch(() => undefined);
    } catch {
      // Ignore SW registration failure in restricted environments
    }
  }, []);

  const applyUpdate = useCallback(async () => {
    if (updateFnRef.current) {
      await updateFnRef.current(true);
    } else {
      window.location.reload();
    }
  }, []);

  const dismissUpdate = useCallback(() => {
    setNeedRefresh(false);
  }, []);

  return {
    needRefresh,
    offlineReady,
    applyUpdate,
    dismissUpdate,
  };
}
