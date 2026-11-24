/**
 * End-to-end Playwright verification checklist for eSZ OFFICE:
 *  1. All 4 modules (PDF, Docs, Sheets, Slides) open existing files in View mode by default
 *     with zero IndexedDB updatedAt mutations.
 *  2. Edit -> Done round trip works cleanly across all 4 modules.
 *  3. Offline reload via Service Worker + IndexedDB succeeds when network is disabled.
 *
 * Designed to be executed against the local app URL using `playwright-cli run-code`.
 */
export async function verifyAllModulesAndOffline(page) {
  await page.setViewportSize({width: 1280, height: 800});
  let isOffline = false;
  const errors = [];
  const offlineNetworkNotices = [];
  const failedRequests = [];

  page.on('console', (msg) => {
    if (msg.type() !== 'error') return;
    const text = msg.text();
    const locUrl = msg.location()?.url || '';
    const isExternalFont = /fonts\.googleapis\.com|fonts\.gstatic\.com/i.test(locUrl);
    if (
      /Failed to load resource|net::ERR_FAILED|ERR_INTERNET_DISCONNECTED/i.test(text) &&
      (isOffline || isExternalFont)
    ) {
      offlineNetworkNotices.push({text, url: locUrl, isOffline});
      return;
    }
    errors.push(`${text}${locUrl ? ` (${locUrl})` : ''}`);
  });
  page.on('pageerror', (err) => errors.push(err.message));
  page.on('requestfailed', (req) => {
    failedRequests.push({
      url: req.url(),
      errorText: req.failure()?.errorText || 'unknown',
      isOffline,
    });
  });

  // Dismiss onboarding modal if present on a fresh browser profile
  const onboarding = page.locator('[aria-labelledby="onboarding-heading"]');
  if (
    await onboarding
      .waitFor({state: 'visible', timeout: 4000})
      .then(() => true)
      .catch(() => false)
  ) {
    await onboarding
      .getByRole('button', {name: /skip|close|get started|start/i})
      .first()
      .click();
    await onboarding.waitFor({state: 'hidden', timeout: 3000});
  }

  const readFileUpdatedAt = async (fileName) =>
    page.evaluate(
      (targetName) =>
        new Promise((resolve, reject) => {
          const req = indexedDB.open('eSZOfficeDB');
          req.onerror = () => reject(req.error);
          req.onsuccess = () => {
            const db = req.result;
            const tx = db.transaction('files', 'readonly');
            const all = tx.objectStore('files').getAll();
            all.onsuccess = () => {
              const rec = all.result.find((f) => f.name === targetName);
              resolve(rec ? rec.updatedAt : null);
            };
            all.onerror = () => reject(all.error);
          };
        }),
      fileName,
    );

  const starterFiles = [
    {
      module: 'pdf',
      name: 'Offline Privacy Specification.pdf',
      editReadySelector: 'button[aria-label="Done editing"]',
    },
    {
      module: 'docs',
      name: 'Welcome to eSZ OFFICE.docx',
      editReadySelector: 'button[aria-label="Done editing"]',
    },
    {
      module: 'sheets',
      name: 'Q4 Budget Tracker.xlsx',
      editReadySelector: 'input[aria-label="Formula bar input"]',
    },
    {
      module: 'slides',
      name: 'Product Roadmap Deck.pptx',
      editReadySelector: 'button[aria-label="Done editing"]',
    },
  ];

  const moduleResults = [];

  for (const item of starterFiles) {
    // Navigate to Home and open the starter document from Recent Documents
    await page.getByRole('link', {name: /^home$/i}).first().click();
    await page.waitForURL((url) => url.pathname === '/', {timeout: 10000});
    const recentSection = page.locator('section', {hasText: 'Recent Documents'}).first();
    const fileBtn = recentSection.getByRole('button', {name: new RegExp(item.name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'))}).first();
    await fileBtn.waitFor({state: 'visible', timeout: 12000});

    const beforeUpdatedAt = await readFileUpdatedAt(item.name);

    await fileBtn.click();
    await page.waitForURL((url) => url.pathname.startsWith('/workspace'), {timeout: 10000});

    const viewToolbar = page.locator('[data-testid="view-mode-toolbar"]');
    const editBtn = viewToolbar.locator('button[aria-label="Edit"]');
    try {
      await editBtn.waitFor({state: 'visible', timeout: 15000});
    } catch (err) {
      const snippet = await page.evaluate(() => document.body.innerText.slice(0, 500));
      throw new Error(
        `[${item.module}] timed out waiting for ViewModeToolbar Edit button at ${page.url()} | errors=${JSON.stringify(errors)} | text=${JSON.stringify(snippet)}`,
      );
    }
    await page.waitForTimeout(300);

    // 1. Verify View mode default
    const openedInView = await editBtn.isVisible();
    if (!openedInView) {
      throw new Error(`${item.module} (${item.name}) did not open in View mode by default`);
    }

    const afterViewUpdatedAt = await readFileUpdatedAt(item.name);
    if (beforeUpdatedAt !== afterViewUpdatedAt) {
      throw new Error(`${item.module} (${item.name}) mutated updatedAt in View mode`);
    }

    // 2. Verify Edit -> Done round trip
    await editBtn.click();
    await page.locator(item.editReadySelector).first().waitFor({state: 'visible', timeout: 10000});

    const doneBtn = page.locator('button[aria-label="Done editing"]').first();
    await doneBtn.click();
    await editBtn.waitFor({state: 'visible', timeout: 10000});

    const returnedToView = await editBtn.isVisible();
    moduleResults.push({
      module: item.module,
      openedInView,
      zeroViewWrites: beforeUpdatedAt === afterViewUpdatedAt,
      editDoneRoundTrip: returnedToView,
    });
  }

  // 3. Offline reload check
  await page.getByRole('link', {name: /^home$/i}).first().click();
  await page.getByText('Workspace Home').waitFor({state: 'visible', timeout: 8000});

  const swReady = await page.evaluate(async () => {
    if (!('serviceWorker' in navigator)) return false;
    const reg = await navigator.serviceWorker.ready;
    return Boolean(reg.active);
  });

  if (swReady) {
    const controlled = await page.evaluate(() => Boolean(navigator.serviceWorker.controller));
    if (!controlled) {
      await page.reload({waitUntil: 'domcontentloaded'});
      await page.getByText('Workspace Home').waitFor({state: 'visible', timeout: 10000});
    }
  }

  isOffline = true;
  await page.context().setOffline(true);
  await page.reload({waitUntil: 'domcontentloaded'});
  await page.getByText('Workspace Home').waitFor({state: 'visible', timeout: 10000});
  await page.getByText('Welcome to eSZ OFFICE.docx').first().waitFor({state: 'visible', timeout: 8000});
  await page.waitForTimeout(300);
  await page.context().setOffline(false);
  isOffline = false;

  if (errors.length > 0) {
    throw new Error(`Console/page errors detected: ${errors.join('; ')}`);
  }

  return {
    moduleResults,
    swReady,
    offlineReloadOk: true,
    errorCount: errors.length,
    offlineNetworkNoticesCount: offlineNetworkNotices.length,
    failedRequests,
  };
}
