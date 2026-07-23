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

test('production editor fragment opens and restores focus when closed', async ({ page }) => {
  const expectCleanRuntime = observeRuntimeErrors(page);
  await page.goto('/editor?mock=editing');
  await page.locator('#editor-trigger').click();

  await expect(page.locator('#audio-editor-overlay')).toHaveClass(/open/);
  await expect.poll(() => page.evaluate(() => typeof window.HemisphereAudioEditor)).toBe('object');
  await expect(page.locator('[data-hae-overlay-close]')).toBeFocused();
  await expect(page.locator('[data-hae-result-audio]')).toHaveAttribute('controlsList', 'nodownload');

  await page.locator('[data-hae-overlay-close]').click();
  await expect(page.locator('#audio-editor-overlay')).not.toHaveClass(/open/);
  await expect(page.locator('#editor-trigger')).toBeFocused();
  await expectCleanRuntime();
});

test('editor loads deterministic 48 kHz stereo PCM without changing channel count', async ({ page }) => {
  const expectCleanRuntime = observeRuntimeErrors(page);
  await page.goto('/editor');
  await page.locator('#editor-trigger').click();
  const fixture = wavFilePayload({
    name: 'stereo-fixture.wav', channels: 2, durationSeconds: 1.2
  });
  expect(fixture.buffer.readUInt16LE(22)).toBe(2);
  expect(fixture.buffer.readUInt32LE(24)).toBe(48000);
  await page.locator('#hae-audio-file').setInputFiles(fixture);

  await expect.poll(() => page.evaluate(() => window.HemisphereAudioEditor.getState().mode)).toBe('editing');
  await expect(page.locator('[data-hae-file-name]')).toHaveText('stereo-fixture.wav');
  await expect(page.locator('[data-hae-file-meta]')).toContainText('WAV');
  await expectCleanRuntime();
});

test('editor performs a real mono WAV to MP3 encode and browser decode', async ({ page }) => {
  test.setTimeout(120_000);
  const expectCleanRuntime = observeRuntimeErrors(page);
  await page.goto('/editor');
  await page.locator('#editor-trigger').click();
  await page.locator('#hae-audio-file').setInputFiles(wavFilePayload({
    name: 'forty-six-seconds.wav', channels: 1, durationSeconds: 46
  }));

  await expect.poll(() => page.evaluate(() => window.HemisphereAudioEditor.getState().mode), {
    timeout: 30_000
  }).toBe('editing');
  const encodeOutcome = await page.evaluate(async () => {
    window.HemisphereAudioEditor.setSelection(0, 46);
    const result = await window.HemisphereAudioEditor.encodeSelection();
    return {
      sizeBytes: result ? result.sizeBytes : 0,
      error: String(window.__lastEncodeError || '')
    };
  });
  expect(encodeOutcome.error).toBe('');
  expect(encodeOutcome.sizeBytes).toBeGreaterThanOrEqual(1024 * 1024);

  const result = await page.evaluate(async () => {
    const value = window.HemisphereAudioEditor.getResult();
    const bytes = await value.blob.arrayBuffer();
    const context = new AudioContext();
    const decoded = await context.decodeAudioData(bytes.slice(0));
    const output = {
      mimeType: value.mimeType,
      blobType: value.blob.type,
      sizeBytes: value.sizeBytes,
      sampleRate: value.sampleRate,
      bitrate: value.bitrate,
      channels: value.numberOfChannels,
      decodedChannels: decoded.numberOfChannels,
      decodedDuration: decoded.duration
    };
    await context.close();
    return output;
  });

  expect(result.mimeType).toBe('audio/mpeg');
  expect(result.blobType).toBe('audio/mpeg');
  expect(result.sizeBytes).toBeGreaterThanOrEqual(1024 * 1024);
  expect(result.sizeBytes).toBeLessThanOrEqual(4 * 1024 * 1024);
  expect(result.sampleRate).toBe(48000);
  expect(result.bitrate).toBe(192000);
  expect(result.channels).toBe(1);
  expect(result.decodedChannels).toBe(1);
  expect(result.decodedDuration).toBeGreaterThan(45);
  expect(result.decodedDuration).toBeLessThan(47);

  const saved = await page.evaluate(() => window.__saveRealEditorResult());
  expect(saved.method).toBe('saveImportAudioItem');
  expect(saved.mimeType).toBe('audio/mpeg');
  expect(saved.decodedBytes).toBe(result.sizeBytes);
  expect(saved.blobBytes).toBe(result.sizeBytes);
  expect(saved.decodedBytes).toBeGreaterThanOrEqual(1024 * 1024);
  expect(saved.decodedBytes).toBeLessThanOrEqual(4 * 1024 * 1024);
  expect(saved.base64Length).toBeGreaterThan(saved.decodedBytes);
  expect(saved.sourceKind).toBe('local');
  expect(saved.sourceFileName).toBe('real-source.wav');
  expect(saved.pin.lat).toBe(35);
  expect(saved.pin.lng).toBe(135);
  await expectCleanRuntime();
});

test('editor saves a real half-second MP3 without a one MiB minimum', async ({ page }) => {
  test.setTimeout(120_000);
  const expectCleanRuntime = observeRuntimeErrors(page);
  await page.goto('/editor');
  await page.locator('#editor-trigger').click();
  await page.locator('#hae-audio-file').setInputFiles(wavFilePayload({
    name: 'half-second-selection.wav', channels: 1, durationSeconds: 1
  }));

  await expect.poll(() => page.evaluate(() => window.HemisphereAudioEditor.getState().mode), {
    timeout: 30_000
  }).toBe('editing');
  const result = await page.evaluate(async () => {
    window.HemisphereAudioEditor.setSelection(0, 0.5);
    await window.HemisphereAudioEditor.encodeSelection();
    const value = window.HemisphereAudioEditor.getResult();
    const context = new AudioContext();
    const decoded = await context.decodeAudioData(await value.blob.arrayBuffer());
    const output = {
      sizeBytes: value.sizeBytes,
      decodedDuration: decoded.duration
    };
    await context.close();
    return output;
  });

  expect(result.sizeBytes).toBeGreaterThanOrEqual(1024);
  expect(result.sizeBytes).toBeLessThan(1024 * 1024);
  expect(result.decodedDuration).toBeGreaterThan(0.4);
  expect(result.decodedDuration).toBeLessThan(0.7);
  const saved = await page.evaluate(() => window.__saveRealEditorResult());
  expect(saved.decodedBytes).toBe(result.sizeBytes);
  expect(saved.decodedBytes).toBeLessThan(1024 * 1024);
  await expectCleanRuntime();
});

test('three-channel PCM is rejected in the UI without an unhandled error', async ({ page }) => {
  const expectCleanRuntime = observeRuntimeErrors(page);
  await page.goto('/editor');
  await page.locator('#editor-trigger').click();
  await page.locator('#hae-audio-file').setInputFiles(wavFilePayload({
    name: 'unsupported-3ch.wav', channels: 3, durationSeconds: 1
  }));

  await expect.poll(() => page.evaluate(() => window.HemisphereAudioEditor.getState().mode)).toBe('editing');
  await page.locator('[data-hae-confirm]').click();
  await expect(page.locator('[data-hae-encode-message]')).toContainText(/モノラル|ステレオ/);
  await expectCleanRuntime();
});

test('editor remains usable after background vendor failure and retries on reopen', async ({ page }) => {
  const expectCleanRuntime = observeRuntimeErrors(page);
  await page.goto('/editor?mock=editing&vendorFailures=1');
  await page.locator('#editor-trigger').click();
  await expect.poll(() => page.evaluate(() => window.__vendor.failures)).toBe(1);

  await expect(page.locator('[data-hae-editor]')).toBeVisible();

  await page.locator('[data-hae-overlay-close]').click();
  await page.locator('#map-action').click();
  await expect(page.locator('#map-action-count')).toHaveText('1');
  await page.locator('#editor-trigger').click();
  await expect.poll(() => page.evaluate(() => window.__vendor.loaded), { timeout: 30_000 }).toBe(true);
  await expectCleanRuntime();
});

test('editor geometry stays inside 375, 768 and 1280 pixel viewports', async ({ page }) => {
  const expectCleanRuntime = observeRuntimeErrors(page);
  await page.goto('/editor?mock=editing');
  await page.locator('#editor-trigger').click();

  for (const width of [375, 768, 1280]) {
    await page.setViewportSize({ width, height: width === 375 ? 720 : 820 });
    await expect.poll(() => page.evaluate(() => {
      const sheet = document.querySelector('.audio-editor-sheet');
      return sheet ? sheet.scrollWidth <= sheet.clientWidth + 1 : false;
    })).toBe(true);
    const closeBox = await page.locator('[data-hae-overlay-close]').boundingBox();
    const confirmBox = await page.locator('[data-hae-confirm]').boundingBox();
    expect(closeBox.width).toBeGreaterThanOrEqual(44);
    expect(closeBox.height).toBeGreaterThanOrEqual(44);
    expect(confirmBox.width).toBeGreaterThanOrEqual(44);
    expect(confirmBox.height).toBeGreaterThanOrEqual(44);
    expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth + 1)).toBe(true);
  }
  await expectCleanRuntime();
});
