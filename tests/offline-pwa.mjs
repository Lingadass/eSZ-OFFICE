// Offline & PWA verification script run via `playwright-cli run-code` against production preview
export async function verifyOfflinePWA(page) {
  await page.waitForSelector('text=Workspace Home');

  // Wait for service worker to register and activate
  const swReady = await page.evaluate(async () => {
    if (!('serviceWorker' in navigator)) return false;
    const reg = await navigator.serviceWorker.ready;
    return Boolean(reg.active);
  });
  if (!swReady) {
    throw new Error('Service worker failed to activate.');
  }

  // Verify manifest.webmanifest + icons + assetlinks + offline.html
  const checks = await page.evaluate(async () => {
    const manifestRes = await fetch('/manifest.webmanifest');
    const manifest = await manifestRes.json();
    const assetLinksRes = await fetch('/.well-known/assetlinks.json');
    const offlineRes = await fetch('/offline.html');
    const logoRes = await fetch('/icons/logo.svg');
    const icon192Res = await fetch('/icons/icon-192.png');
    const icon512Res = await fetch('/icons/icon-512.png');
    const maskableRes = await fetch('/icons/icon-maskable-512.png');
    return {
      manifestOk: manifestRes.ok,
      name: manifest.name,
      shortName: manifest.short_name,
      display: manifest.display,
      startUrl: manifest.start_url,
      shortcuts: manifest.shortcuts?.length ?? 0,
      fileHandlers: manifest.file_handlers?.length ?? 0,
      shareTarget: Boolean(manifest.share_target),
      assetLinksOk: assetLinksRes.ok,
      offlineOk: offlineRes.ok,
      logoOk: logoRes.ok,
      icon192Ok: icon192Res.ok,
      icon512Ok: icon512Res.ok,
      maskableOk: maskableRes.ok,
    };
  });

  for (const [k, v] of Object.entries(checks)) {
    if (v === false) throw new Error(`PWA check failed for ${k}`);
  }

  // Simulate full network offline mode and reload
  await page.context().setOffline(true);
  await page.reload({waitUntil: 'domcontentloaded'});
  await page.waitForSelector('text=Workspace Home');
  await page.waitForSelector('text=Offline Mode');
  await page.context().setOffline(false);

  return checks;
}
