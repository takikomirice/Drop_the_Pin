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

function productionPin(overrides) {
  return Object.assign({
    id: 'production-audio-pin',
    title: '本番テンプレートの音声ピン',
    description: '統合テスト',
    eventAt: '2026-07-22T10:00',
    timestamp: '2026-07-22T09:00:00.000Z',
    updatedAt: '2026-07-22T10:00:00.000Z',
    lat: 35.0,
    lng: 135.0,
    color: '#2196f3',
    icon: 'default',
    status: '未対応',
    tags: ['音声'],
    links: [],
    fileId: '',
    imageUrl: '',
    hasAudio: true
  }, overrides || {});
}

function photoResponse() {
  const bytes = Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
    'base64'
  );
  return {
    ok: true,
    mimeType: 'image/png',
    byteLength: bytes.length,
    base64: bytes.toString('base64')
  };
}

test('processed index template uses real detail, hasAudio and source chooser wiring', async ({ page }) => {
  const expectCleanRuntime = observeRuntimeErrors(page);
  await page.goto('/production-edit');
  await expect.poll(() => page.evaluate(() => typeof window.__productionEdit)).toBe('object');
  expect(await page.evaluate(() => window.__productionInitializationSuppressed)).toBe(true);
  await expect(page.locator('#audio-editor-overlay')).toHaveCount(1);
  await expect(page.locator('#pin-audio-player')).toHaveCount(1);

  const pin = productionPin();
  await enqueue(page, 'getPinAudioData', { audioSeed: 11 });
  await page.evaluate((pin) => {
    window.__productionEdit.state.pins = [pin];
    window.__productionEdit.openPinDetail(pin);
  }, pin);

  await expect(page.locator('#pin-detail-overlay')).toHaveClass(/open/);
  await expect(page.locator('#pin-detail-title')).toHaveText(pin.title);
  await expect(page.locator('#pin-detail-audio-add')).toBeHidden();
  await expect(page.locator('#pin-detail-audio-replace')).toBeVisible();
  await expect(page.locator('#pin-detail-audio-delete')).toBeVisible();
  await expect(page.locator('#pin-audio-runtime')).toBeVisible();
  await expect(page.locator('#pin-audio-runtime')).toHaveAttribute('controlsList', 'nodownload');
  expect(await page.locator('#pin-audio-runtime').evaluate((audio) => audio.controls)).toBe(true);
  const audioCall = await page.evaluate(() => window.__gasMock.calls[0]);
  expect(audioCall).toEqual({
    method: 'getPinAudioData',
    payload: { __editToken: 'edit-token-browser-test', pinId: pin.id }
  });

  await page.locator('#pin-detail-audio-replace').click();
  await expect(page.locator('#pin-detail-overlay')).not.toHaveClass(/open/);
  await expect(page.locator('#pin-audio-source-overlay')).toHaveClass(/open/);
  await expect(page.locator('#pin-audio-source-title')).toHaveText('音声を差し替え');
  await expect(page.locator('#pin-audio-source-target')).toContainText(pin.title);
  await expect(page.locator('#pin-audio-source-local')).toBeVisible();
  await expect(page.locator('#pin-audio-source-drive')).toBeVisible();

  await page.locator('#pin-audio-source-cancel').click();
  await expect(page.locator('#pin-audio-source-overlay')).not.toHaveClass(/open/);
  await expect(page.locator('#pin-detail-overlay')).toHaveClass(/open/);
  expect(await page.evaluate(() => window.__gasMock.calls.length)).toBe(1);

  const noAudioPin = productionPin({
    id: 'production-no-audio', title: '音声なしピン', hasAudio: false,
    updatedAt: '2026-07-22T11:00:00.000Z'
  });
  await page.evaluate((pin) => {
    window.__productionEdit.closePinDetail({ restoreFocus: false });
    window.__productionEdit.state.pins.push(pin);
    window.__productionEdit.openPinDetail(pin);
  }, noAudioPin);
  await expect(page.locator('#pin-detail-audio-add')).toBeVisible();
  await expect(page.locator('#pin-detail-audio-replace')).toBeHidden();
  await expect(page.locator('#pin-audio-player')).toBeHidden();
  await page.locator('#pin-detail-audio-add').click();
  await expect(page.locator('#pin-audio-source-title')).toHaveText('音声を追加');
  await expectCleanRuntime();
});

test('processed edit template lazy-loads photo bytes and hides unavailable mobile actions', async ({ page }) => {
  const expectCleanRuntime = observeRuntimeErrors(page);
  await page.goto('/production-edit');
  await expect.poll(() => page.evaluate(() => typeof window.__productionEdit)).toBe('object');

  const pin = productionPin({
    id: 'production-photo-pin',
    title: '本番テンプレートの写真ピン',
    hasPhoto: true,
    hasAudio: false,
    fileId: 'private-drive-photo-id',
    imageUrl: 'https://drive.google.com/thumbnail?id=private-drive-photo-id'
  });
  await enqueue(page, 'getPinPhotoData', { response: photoResponse() });
  await page.evaluate((photoPin) => {
    window.__productionEdit.state.pins = [photoPin];
    window.__productionEdit.openPinDetail(photoPin);
  }, pin);

  await expect(page.locator('#pin-detail-image-trigger')).toBeVisible();
  const detailSource = await page.locator('#pin-detail-image').getAttribute('src');
  expect(detailSource).toMatch(/^blob:/);
  expect(detailSource).not.toContain('drive.google.com');
  const photoCalls = await page.evaluate(() =>
    window.__gasMock.calls.filter((call) => call.method === 'getPinPhotoData'));
  expect(photoCalls).toEqual([{
    method: 'getPinPhotoData',
    payload: { __editToken: 'edit-token-browser-test', pinId: pin.id }
  }]);

  await page.locator('#pin-detail-image-trigger').click();
  await expect(page.locator('#photo-viewer-overlay')).toHaveClass(/open/);
  await expect(page.locator('#photo-viewer-image')).toHaveAttribute('src', detailSource);

  await page.setViewportSize({ width: 390, height: 844 });
  await page.evaluate(() => {
    document.body.classList.add('narrow-view');
    window.__productionEdit.state.narrowView = true;
  });
  await expect(page.locator('#share-open-btn')).toBeHidden();
  await expect(page.locator('#data-toggle')).toBeHidden();

  await page.setViewportSize({ width: 1280, height: 720 });
  await page.evaluate(() => {
    document.body.classList.remove('narrow-view');
    window.__productionEdit.state.narrowView = false;
  });
  await expect(page.locator('#share-open-btn')).toBeVisible();
  await expect(page.locator('#data-toggle')).toBeVisible();
  await expectCleanRuntime();
});

test('edit preview keeps native audio playback while mutation actions stay hidden', async ({ page }) => {
  const expectCleanRuntime = observeRuntimeErrors(page);
  await page.goto('/production-edit');
  await expect.poll(() => page.evaluate(() => typeof window.__productionEdit)).toBe('object');
  const pin = productionPin({ id: 'preview-audio-pin' });
  await enqueue(page, 'getPinAudioData', { audioSeed: 13 });

  await page.evaluate((previewPin) => {
    window.__productionEdit.state.previewMode = true;
    window.__productionEdit.state.pins = [previewPin];
    window.__productionEdit.openPinDetail(previewPin);
  }, pin);

  await expect(page.locator('#pin-audio-runtime')).toBeVisible();
  await expect(page.locator('#pin-audio-runtime')).toHaveAttribute('controlsList', 'nodownload');
  await expect(page.locator('#pin-detail-audio-actions')).toBeHidden();
  const calls = await page.evaluate(() =>
    window.__gasMock.calls.filter((call) => call.method === 'getPinAudioData'));
  expect(calls).toEqual([{
    method: 'getPinAudioData',
    payload: { __editToken: 'edit-token-browser-test', pinId: pin.id }
  }]);
  await expectCleanRuntime();
});

test('real edit initialization warms vendor in background, tolerates failure and retries on editor open', async ({ page }) => {
  const expectCleanRuntime = observeRuntimeErrors(page);
  await page.goto('/production-edit?mock=editing');
  await expect.poll(() => page.evaluate(() => typeof window.__productionEdit)).toBe('object');

  await enqueue(page, 'getAppSettings', {
    response: { ok: true, rootFolderId: 'root-folder', rootFolderUrl: '#', renameFileWithTitle: false }
  });
  await enqueue(page, 'getMapData', { response: [] });
  await enqueue(page, 'getRouteGroups', { response: [] });
  await enqueue(page, 'getTracks', { response: { ok: true, tracks: [], warnings: [] } });
  await enqueue(page, 'getAudioVendorBundle', { failure: 'configured background failure' });

  await page.evaluate(() => window.__productionEdit.initializeApp());
  await expect.poll(() => page.evaluate(() => window.__gasMock.calls.filter((call) => call.method === 'getAudioVendorBundle').length)).toBe(1);
  expect(await page.evaluate(() => window.__productionEdit.state.initializing)).toBe(false);
  expect(await page.evaluate(() => typeof window.Mediabunny)).toBe('undefined');

  const mapStillWorks = await page.evaluate(() => {
    const map = window.__productionEdit.map;
    map.setView([35, 135], 10);
    return map.getZoom();
  });
  expect(mapStillWorks).toBe(5);

  const vendorSource = await page.evaluate(() => fetch('/audio-vendor.js').then((response) => response.text()));
  await enqueue(page, 'getAudioVendorBundle', {
    response: { version: '1.50.8', source: vendorSource }
  });
  await page.evaluate(() => window.HemisphereAudioEditor.init());
  await expect(page.locator('#audio-editor-overlay')).toHaveClass(/open/);
  await expect.poll(() => page.evaluate(() => !!(window.Mediabunny && window.MediabunnyMp3Encoder)), {
    timeout: 30_000
  }).toBe(true);
  expect(await page.evaluate(() => window.__gasMock.calls.filter((call) => call.method === 'getAudioVendorBundle').length)).toBe(2);
  await expectCleanRuntime();
});

test('processed shared template consumes the projected DTO and never includes editor or vendor', async ({ page }) => {
  const expectCleanRuntime = observeRuntimeErrors(page);
  await page.goto('/production-shared');
  await expect.poll(() => page.evaluate(() => typeof window.__productionShared)).toBe('object');
  expect(await page.evaluate(() => window.__productionInitializationSuppressed)).toBe(true);

  const allowedPin = productionPin({ id: 'shared-allowed', title: '共有対象の音声ピン' });
  const forbiddenId = 'shared-not-projected';
  await enqueue(page, 'getSharedViewData', {
    response: {
      ok: true,
      pins: [allowedPin],
      routes: [],
      allowedRouteIds: [],
      allowedTags: ['音声'],
      allowedColors: ['#2196f3'],
      shareLink: { label: '音声共有', tagMode: 'or' }
    }
  });
  await page.evaluate(() => window.__productionShared.initializeSharedView());

  await expect(page.locator(`#shared-list [data-pin-id="${allowedPin.id}"]`)).toHaveCount(1);
  await expect(page.locator(`#shared-list [data-pin-id="${forbiddenId}"]`)).toHaveCount(0);
  expect(await page.evaluate(() => window.__gasMock.calls.filter((call) => call.method === 'getSharedPinAudioData').length)).toBe(0);
  await expect(page.locator('#audio-editor-overlay, [data-hemisphere-audio-editor]')).toHaveCount(0);
  expect(await page.evaluate(() => typeof window.Mediabunny)).toBe('undefined');
  expect(await page.evaluate(() => performance.getEntriesByType('resource').filter((entry) => entry.name.includes('/audio-vendor.js')).length)).toBe(0);

  await enqueue(page, 'getSharedPinAudioData', { audioSeed: 12 });
  await page.locator(`#shared-list [data-pin-id="${allowedPin.id}"]`).click();
  await expect(page.locator('#shared-detail-overlay')).toHaveClass(/open/);
  await expect(page.locator('#shared-detail-title')).toHaveText(allowedPin.title);
  await expect(page.locator('#pin-audio-runtime')).toBeVisible();
  await expect(page.locator('#pin-audio-runtime')).toHaveAttribute('controlsList', 'nodownload');
  expect(await page.locator('#pin-audio-runtime').evaluate((audio) => audio.controls)).toBe(true);
  const audioCall = await page.evaluate(() => window.__gasMock.calls.find((call) => call.method === 'getSharedPinAudioData'));
  expect(audioCall).toEqual({
    method: 'getSharedPinAudioData',
    payload: { shareToken: 'share-token-browser-test', pinId: allowedPin.id }
  });
  expect(audioCall.payload).not.toHaveProperty('editToken');
  expect(audioCall.payload).not.toHaveProperty('__editToken');

  await page.locator('#shared-detail-close').click();
  await expect(page.locator('#shared-detail-overlay')).not.toHaveClass(/open/);
  await expect(page.locator('#pin-audio-player')).toBeHidden();
  await expectCleanRuntime();
});
