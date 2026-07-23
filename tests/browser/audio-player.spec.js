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

async function enqueue(page, method, item) {
  await page.evaluate(({ method, item }) => window.__gasMock.enqueue(method, item), { method, item });
}

test('native player covers loading, playback contract and detail close', async ({ page }) => {
  const expectCleanRuntime = observeRuntimeErrors(page);
  await page.goto('/player');
  expect(await page.evaluate(() => window.__gasMock.calls.length)).toBe(0);

  await enqueue(page, 'getPinAudioData', { defer: 'pin-a-ready', audioSeed: 1 });
  const opener = page.locator('[data-open-pin="pin-a"]');
  await opener.click();

  await expect(page.locator('#pin-detail')).toBeVisible();
  await expect(page.locator('#pin-audio-player')).toBeVisible();
  await expect(page.locator('#pin-audio-player')).toHaveAttribute('aria-busy', 'true');
  await expect(page.locator('#pin-audio-player-status')).toHaveText('音声を準備中…');
  await expect(page.locator('#pin-audio-runtime')).toBeHidden();

  await page.evaluate(() => window.__gasMock.resolve('pin-a-ready'));
  await expect(page.locator('#pin-audio-runtime')).toBeVisible();
  await expect(page.locator('#pin-audio-player')).toHaveAttribute('aria-busy', 'false');
  await expect(page.locator('#pin-audio-player-status')).toHaveText('');

  const request = await page.evaluate(() => window.__gasMock.calls[0]);
  expect(request).toEqual({
    method: 'getPinAudioData',
    payload: { editToken: 'edit-token-browser-test', pinId: 'pin-a' }
  });

  await page.locator('#pin-audio-runtime').evaluate((audio) => audio.play());
  expect(await page.evaluate(() => window.__media.plays)).toBe(1);

  const audioContract = await page.evaluate(() => {
    const audio = document.getElementById('pin-audio-runtime');
    return {
      hidden: audio.hidden,
      controls: audio.controls,
      controlsAttribute: audio.hasAttribute('controls'),
      controlsList: Array.from(audio.controlsList),
      ranges: document.querySelectorAll('#pin-audio-player input[type="range"]').length,
      buttons: Array.from(document.querySelectorAll('#pin-audio-player button')).map((button) => button.textContent.trim())
    };
  });
  expect(audioContract).toEqual({
    hidden: false,
    controls: true,
    controlsAttribute: true,
    controlsList: ['nodownload'],
    ranges: 0,
    buttons: ['再試行']
  });

  await page.locator('#detail-close').click();
  await expect(page.locator('#pin-detail')).toBeHidden();
  await expect(opener).toBeFocused();
  await expect(page.locator('#pin-audio-player')).toBeHidden();
  await expectCleanRuntime();
});

test('failed fetch exposes one retry action and recovers', async ({ page }) => {
  const expectCleanRuntime = observeRuntimeErrors(page);
  await page.goto('/player');
  await enqueue(page, 'getPinAudioData', { failure: 'temporary failure' });
  await page.locator('[data-open-pin="pin-b"]').click();

  await expect(page.locator('#pin-audio-player-status')).toHaveText('音声を再生できませんでした');
  await expect(page.locator('#pin-audio-runtime')).toBeHidden();
  await expect(page.locator('#pin-audio-player-retry')).toBeVisible();
  await expect(page.locator('#pin-audio-player-retry')).toBeEnabled();

  await enqueue(page, 'getPinAudioData', { audioSeed: 2 });
  await page.locator('#pin-audio-player-retry').click();
  await expect(page.locator('#pin-audio-runtime')).toBeVisible();
  expect(await page.evaluate(() => window.__gasMock.calls.length)).toBe(2);
  await expectCleanRuntime();
});

test('stale responses are ignored and four ready pins evict the LRU object URL', async ({ page }) => {
  const expectCleanRuntime = observeRuntimeErrors(page);
  await page.goto('/player');

  await enqueue(page, 'getPinAudioData', { defer: 'stale-a', audioSeed: 1 });
  await page.locator('[data-open-pin="pin-a"]').click();
  await expect(page.locator('#pin-audio-player')).toHaveAttribute('aria-busy', 'true');

  await enqueue(page, 'getPinAudioData', { audioSeed: 2 });
  await page.locator('[data-open-pin="pin-b"]').click();
  await expect(page.locator('#pin-audio-runtime')).toBeVisible();
  const sourceForB = await page.locator('#pin-audio-runtime').getAttribute('src');
  await page.evaluate(() => window.__gasMock.resolve('stale-a'));
  await expect.poll(() => page.evaluate(() => window.__currentPin)).toBe('pin-b');
  expect(await page.locator('#pin-audio-runtime').getAttribute('src')).toBe(sourceForB);
  expect(await page.evaluate(() => window.__media.created.length)).toBe(1);

  for (const [pinId, seed] of [['pin-c', 3], ['pin-d', 4], ['pin-e', 5]]) {
    await enqueue(page, 'getPinAudioData', { audioSeed: seed });
    await page.locator(`[data-open-pin="${pinId}"]`).click();
    await expect(page.locator('#pin-audio-runtime')).toBeVisible();
  }
  expect(await page.evaluate(() => window.__media.created.length)).toBe(4);
  expect(await page.evaluate(() => window.__media.revoked.length)).toBe(1);

  await enqueue(page, 'getPinAudioData', { audioSeed: 6 });
  await page.locator('[data-open-pin="pin-b"]').click();
  await expect(page.locator('#pin-audio-runtime')).toBeVisible();
  const pinBCalls = await page.evaluate(() => window.__gasMock.calls.filter((call) => call.payload.pinId === 'pin-b').length);
  expect(pinBCalls).toBe(2);
  expect(await page.evaluate(() => window.__media.revoked.length)).toBe(2);

  await page.evaluate(() => window.dispatchEvent(new Event('pagehide')));
  await expect.poll(() => page.evaluate(() => window.__media.revoked.length)).toBe(5);
  await expectCleanRuntime();
});

test('pin detail native player has no horizontal overflow and close keeps a 44 pixel target', async ({ page }) => {
  const expectCleanRuntime = observeRuntimeErrors(page);
  await page.goto('/player');
  await enqueue(page, 'getPinAudioData', { audioSeed: 1 });
  await page.locator('[data-open-pin="pin-a"]').click();
  await expect(page.locator('#pin-audio-runtime')).toBeVisible();

  for (const width of [375, 768, 1280]) {
    await page.setViewportSize({ width, height: 720 });
    expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth + 1)).toBe(true);
    const detailFits = await page.locator('#pin-detail').evaluate((element) => element.scrollWidth <= element.clientWidth + 1);
    expect(detailFits).toBe(true);
    const audioFits = await page.locator('#pin-audio-runtime').evaluate((audio) =>
      audio.scrollWidth <= audio.clientWidth + 1
      && audio.getBoundingClientRect().right
        <= audio.parentElement.getBoundingClientRect().right + 1);
    expect(audioFits).toBe(true);
    const closeBox = await page.locator('#detail-close').boundingBox();
    expect(closeBox.width).toBeGreaterThanOrEqual(44);
    expect(closeBox.height).toBeGreaterThanOrEqual(44);
  }
  await expectCleanRuntime();
});
