const { test, expect } = require('@playwright/test');

function observeRuntimeErrors(page) {
  const pageErrors = [];
  page.on('pageerror', (error) => pageErrors.push(String(error && error.message || error)));
  return async function expectCleanRuntime() {
    expect(pageErrors).toEqual([]);
    const recorded = await page.evaluate(() => window.__runtime);
    expect(recorded.pageErrors).toEqual([]);
    expect(recorded.unhandledRejections).toEqual([]);
  };
}

async function enqueue(page, item) {
  await page.evaluate((item) => window.__gasMock.enqueue('getSharedPinAudioData', item), item);
}

test('shared page starts without editor, vendor or audio fetch and uses only the share token', async ({ page }) => {
  const expectCleanRuntime = observeRuntimeErrors(page);
  await page.goto('/shared');
  const initial = await page.evaluate(() => ({
    calls: window.__gasMock.calls.length,
    editorCount: document.querySelectorAll('#audio-editor-overlay, [data-hemisphere-audio-editor]').length,
    vendorGlobal: typeof window.Mediabunny,
    vendorRequests: performance.getEntriesByType('resource').filter((entry) => entry.name.includes('/audio-vendor.js')).length
  }));
  expect(initial).toEqual({ calls: 0, editorCount: 0, vendorGlobal: 'undefined', vendorRequests: 0 });

  await enqueue(page, { audioSeed: 7 });
  await page.locator('[data-open-pin="pin-a"]').click();
  await expect(page.locator('#pin-audio-runtime')).toBeVisible();
  await expect(page.locator('#pin-audio-runtime')).toHaveAttribute('controlsList', 'nodownload');
  expect(await page.locator('#pin-audio-runtime').evaluate((audio) => audio.controls)).toBe(true);
  const call = await page.evaluate(() => window.__gasMock.calls[0]);
  expect(call).toEqual({
    method: 'getSharedPinAudioData',
    payload: { shareToken: 'share-token-browser-test', pinId: 'pin-a' }
  });
  expect(call.payload).not.toHaveProperty('editToken');
  await expectCleanRuntime();
});

test('shared out-of-projection failure stays safe and a later shared pin can play', async ({ page }) => {
  const expectCleanRuntime = observeRuntimeErrors(page);
  await page.goto('/shared');
  await enqueue(page, { failure: 'Audio is unavailable.' });
  await page.locator('[data-open-pin="pin-b"]').click();
  await expect(page.locator('#pin-audio-player-status')).toHaveText('音声を再生できませんでした');
  await expect(page.locator('#pin-audio-player-retry')).toBeVisible();

  await enqueue(page, { audioSeed: 8 });
  await page.locator('[data-open-pin="pin-c"]').click();
  await expect(page.locator('#pin-audio-runtime')).toBeVisible();
  await page.locator('#pin-audio-runtime').evaluate((audio) => audio.play());
  expect(await page.evaluate(() => window.__media.plays)).toBe(1);

  const opener = page.locator('[data-open-pin="pin-c"]');
  await page.locator('#detail-close').click();
  await expect(page.locator('#pin-detail')).toBeHidden();
  await expect(opener).toBeFocused();
  expect(await page.evaluate(() => window.__media.pauses)).toBeGreaterThan(0);
  await expectCleanRuntime();
});

test('shared pagehide revokes cached audio without ever loading vendor code', async ({ page }) => {
  const expectCleanRuntime = observeRuntimeErrors(page);
  await page.goto('/shared');
  await enqueue(page, { audioSeed: 9 });
  await page.locator('[data-open-pin="pin-d"]').click();
  await expect(page.locator('#pin-audio-runtime')).toBeVisible();
  expect(await page.evaluate(() => window.__media.created.length)).toBe(1);

  await page.evaluate(() => window.dispatchEvent(new Event('pagehide')));
  await expect.poll(() => page.evaluate(() => window.__media.revoked.length)).toBe(1);
  const separation = await page.evaluate(() => ({
    vendorGlobal: typeof window.Mediabunny,
    editorCount: document.querySelectorAll('#audio-editor-overlay, [data-hemisphere-audio-editor]').length,
    vendorRequests: performance.getEntriesByType('resource').filter((entry) => entry.name.includes('/audio-vendor.js')).length
  }));
  expect(separation).toEqual({ vendorGlobal: 'undefined', editorCount: 0, vendorRequests: 0 });
  await expectCleanRuntime();
});
