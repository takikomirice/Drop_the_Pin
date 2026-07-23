const { test, expect } = require('@playwright/test');
const { wavFilePayload } = require('./audio-fixtures');

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

async function enqueue(page, method, response) {
  await page.evaluate(({ method, response }) => window.__gasMock.enqueue(method, { response }), { method, response });
}

function savedPin(overrides) {
  return Object.assign({
    id: 'new-pin', title: '録音', description: '', eventAt: '', lat: 34.987, lng: 135.765,
    color: '#176a8d', icon: 'pin', status: 'active', tags: [], links: [],
    updatedAt: '2026-07-22T01:00:00.000Z', hasAudio: true
  }, overrides || {});
}

async function rejectedCode(page, expression) {
  return page.evaluate(expression).then(
    () => '',
    (error) => String(error && error.message || error)
  );
}

test('local audio creates a pin only after an explicit map or unplaced choice', async ({ page }) => {
  const expectCleanRuntime = observeRuntimeErrors(page);
  await page.goto('/workflow');
  await page.locator('#workflow-local-file').setInputFiles(wavFilePayload({
    name: '大阪城の録音.wav', channels: 1, durationSeconds: 1
  }));
  await page.evaluate(() => window.__workflow.beginLocal('create-pin'));

  await expect(page.locator('#workflow-preview')).toBeVisible();
  await expect(page.locator('#workflow-title')).toHaveValue('大阪城の録音');
  const missingLocation = await page.evaluate(() => window.__workflow.save().then(
    () => '', (error) => String(error.code || error.message)
  ));
  expect(missingLocation).toBe('AUDIO_LOCATION_REQUIRED');
  expect(await page.evaluate(() => window.__gasMock.calls.length)).toBe(0);

  await page.locator('#choose-map').click();
  await enqueue(page, 'saveImportAudioItem', { ok: true, pin: savedPin({ id: 'local-created', title: '大阪城の録音' }) });
  await page.evaluate(() => window.__workflow.save());

  const call = await page.evaluate(() => window.__gasMock.calls[0]);
  expect(call.method).toBe('saveImportAudioItem');
  expect(call.payload.editToken).toBe('edit-token-browser-test');
  expect(call.payload.sourceKind).toBe('local');
  expect(call.payload.sourceDriveFileId).toBe('');
  expect(call.payload.sourceFileName).toBe('大阪城の録音.wav');
  expect(call.payload.audioMimeType).toBe('audio/mpeg');
  expect(call.payload.audioBase64).toMatch(/^[A-Za-z0-9+/]+={0,2}$/);
  expect(call.payload.pin.lat).toBe(34.987);
  expect(call.payload.pin.lng).toBe(135.765);
  expect(await page.evaluate(() => window.__runtime.geocoderCalls)).toBe(0);
  await expectCleanRuntime();
});

test('Drive audio materializes as a File and creates an unplaced pin without geocoding', async ({ page }) => {
  const expectCleanRuntime = observeRuntimeErrors(page);
  await page.goto('/workflow');
  await page.evaluate(() => window.__workflow.beginDrive('create-pin'));
  await page.locator('#choose-unplaced').click();
  await enqueue(page, 'saveImportAudioItem', {
    ok: true,
    pin: savedPin({ id: 'drive-created', title: 'Driveの録音', lat: null, lng: null })
  });
  await page.evaluate(() => window.__workflow.save());

  const state = await page.evaluate(() => ({
    call: window.__gasMock.calls[0],
    editorSource: window.__workflowState.editorSources[0],
    geocoderCalls: window.__runtime.geocoderCalls
  }));
  expect(state.call.payload.sourceKind).toBe('drive');
  expect(state.call.payload.sourceDriveFileId).toBe('driveaudio123');
  expect(state.call.payload.pin.lat).toBeNull();
  expect(state.call.payload.pin.lng).toBeNull();
  expect(state.editorSource).toEqual({ sourceKind: 'drive', sourceDriveFileId: 'driveaudio123' });
  expect(state.geocoderCalls).toBe(0);
  await expectCleanRuntime();
});

test('local attach preserves every existing coordinate and omits location from the payload', async ({ page }) => {
  const expectCleanRuntime = observeRuntimeErrors(page);
  await page.goto('/workflow');
  await page.locator('#workflow-local-file').setInputFiles(wavFilePayload({
    name: '別地点名.wav', channels: 1, durationSeconds: 1
  }));
  await page.evaluate(() => window.__workflow.beginLocal('attach-existing-pin'));

  await expect(page.locator('#location-choices')).toBeHidden();
  const locationCode = await page.evaluate(() => {
    try { window.__workflow.chooseLocation('map'); return ''; }
    catch (error) { return String(error.code || error.message); }
  });
  expect(locationCode).toBe('AUDIO_LOCATION_NOT_APPLICABLE');

  await enqueue(page, 'saveImportAudioItem', {
    ok: true,
    pin: savedPin({
      id: 'existing', title: '既存ピン', description: '説明', lat: 35.1, lng: 135.2,
      tags: ['音声'], updatedAt: '2026-07-22T02:00:00.000Z'
    })
  });
  await page.evaluate(() => window.__workflow.save());
  const result = await page.evaluate(() => ({
    payload: window.__gasMock.calls[0].payload,
    pin: window.__workflowState.pins.existing,
    options: window.__workflowState.previewOptions
  }));
  expect(result.options.readOnlyFields).toBe(true);
  expect(result.payload.operationMode).toBe('attach-existing-pin');
  expect(result.payload.targetPinId).toBe('existing');
  expect(result.payload.expectedUpdatedAt).toBe('2026-07-22T00:00:00.000Z');
  expect(result.payload).not.toHaveProperty('pin');
  expect(result.payload).not.toHaveProperty('lat');
  expect(result.payload).not.toHaveProperty('lng');
  expect(result.pin.lat).toBe(35.1);
  expect(result.pin.lng).toBe(135.2);
  expect(await page.evaluate(() => window.__runtime.geocoderCalls)).toBe(0);
  await expectCleanRuntime();
});

test('Drive replace failure preserves the old player and retry keeps the same payload', async ({ page }) => {
  const expectCleanRuntime = observeRuntimeErrors(page);
  await page.goto('/workflow');
  await page.evaluate(() => window.__workflow.beginDrive('replace-existing-audio'));
  await expect(page.locator('#existing-player')).toBeVisible();

  await enqueue(page, 'saveImportAudioItem', {
    ok: false, errorCode: 'AUDIO_IMPORT_SAVE_FAILED', error: '保存できませんでした', retryable: true
  });
  const failure = await page.evaluate(() => window.__workflow.save().then(
    () => '', (error) => String(error.code || error.message)
  ));
  expect(failure).toBe('AUDIO_IMPORT_SAVE_FAILED');
  await expect(page.locator('#existing-player')).toBeVisible();
  expect(await page.evaluate(() => window.__workflowState.pins.existing.lat)).toBe(35.1);

  await enqueue(page, 'saveImportAudioItem', {
    ok: true,
    pin: savedPin({ id: 'existing', title: '既存ピン', lat: 35.1, lng: 135.2, updatedAt: '2026-07-22T03:00:00.000Z' })
  });
  await page.evaluate(() => window.__workflow.retry());
  const calls = await page.evaluate(() => window.__gasMock.calls);
  expect(calls).toHaveLength(2);
  expect(calls[1].payload).toEqual(calls[0].payload);
  expect(calls[1].payload.operationMode).toBe('replace-existing-audio');
  expect(calls[1].payload.sourceKind).toBe('drive');
  expect(calls[1].payload).not.toHaveProperty('pin');
  await expect(page.locator('#existing-player')).toBeVisible();
  await expectCleanRuntime();
});

test('existing audio removal hides the old player only after server success', async ({ page }) => {
  const expectCleanRuntime = observeRuntimeErrors(page);
  await page.goto('/workflow');
  await enqueue(page, 'removePinAudio', {
    ok: true,
    pin: savedPin({ id: 'existing', title: '既存ピン', lat: 35.1, lng: 135.2, hasAudio: false })
  });
  await page.evaluate(() => window.__workflow.removeExisting());
  await expect(page.locator('#existing-player')).toBeHidden();
  const call = await page.evaluate(() => window.__gasMock.calls[0]);
  expect(call).toEqual({
    method: 'removePinAudio',
    payload: {
      editToken: 'edit-token-browser-test',
      pinId: 'existing',
      expectedUpdatedAt: '2026-07-22T00:00:00.000Z'
    }
  });
  await expectCleanRuntime();
});

test('source chooser and preview fit all target widths and cancel restores focus', async ({ page }) => {
  const expectCleanRuntime = observeRuntimeErrors(page);
  await page.goto('/workflow');
  const trigger = page.locator('#workflow-trigger');
  await trigger.click();
  await expect(page.locator('#source-chooser')).toBeVisible();

  for (const width of [375, 768, 1280]) {
    await page.setViewportSize({ width, height: 720 });
    expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth + 1)).toBe(true);
    expect(await page.locator('#source-chooser').evaluate((element) => element.scrollWidth <= element.clientWidth + 1)).toBe(true);
    for (const selector of ['[data-source="local"]', '[data-source="drive"]']) {
      const box = await page.locator(selector).boundingBox();
      expect(box.width).toBeGreaterThanOrEqual(44);
      expect(box.height).toBeGreaterThanOrEqual(44);
    }
  }

  await page.locator('#workflow-local-file').setInputFiles(wavFilePayload({
    name: 'responsive.wav', channels: 1, durationSeconds: 1
  }));
  await page.evaluate(() => window.__workflow.beginLocal('create-pin'));
  expect(await page.locator('#workflow-preview').evaluate((element) => element.scrollWidth <= element.clientWidth + 1)).toBe(true);
  await page.locator('#workflow-cancel').click();
  await expect(trigger).toBeFocused();
  await expectCleanRuntime();
});
