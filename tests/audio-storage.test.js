const assert = require('node:assert/strict');
const test = require('node:test');
const {
  makeHarness, audioPayload, validMp3Bytes, RECEIPT_HEADERS
} = require('./audio-storage-harness');

function plain(value) { return JSON.parse(JSON.stringify(value)); }
function receiptValue(harness, name) {
  return harness.receipt()[RECEIPT_HEADERS.indexOf(name)];
}

test('saveImportAudioItem exposes the Task 4 storage API', () => {
  const harness = makeHarness();
  assert.equal(typeof harness.api.saveImportAudioItem, 'function');
  assert.equal(typeof harness.api.readPinAudioBlobByPinId_, 'function');
});

test('MP3 validation accepts canonical 1 KiB-4 MiB payloads and rejects malformed boundaries before Drive writes', () => {
  const validSizes = [1024, 4 * 1024 * 1024];
  validSizes.forEach((size, index) => {
    const harness = makeHarness();
    const request = audioPayload({
      jobId: `audio-valid-${index}`, itemId: 'item',
      idempotencyKey: `audio-valid-${index}:item`,
      audioBase64: Buffer.from(validMp3Bytes(size)).toString('base64')
    });
    assert.equal(harness.api.saveImportAudioItem(request).ok, true);
  });
  const mpegHarness = makeHarness();
  const mpegBytes = Buffer.alloc(1024, 0);
  mpegBytes[0] = 0xff; mpegBytes[1] = 0xfb;
  assert.equal(mpegHarness.api.saveImportAudioItem(audioPayload({
    jobId: 'audio-mpeg', itemId: 'item', idempotencyKey: 'audio-mpeg:item',
    audioBase64: mpegBytes.toString('base64')
  })).ok, true);

  const invalidPayloads = [
    { audioBase64: '' },
    { audioBase64: '***not-base64***' },
    { audioMimeType: 'audio/wav' },
    { audioBase64: Buffer.from(validMp3Bytes(1023)).toString('base64') },
    { audioBase64: Buffer.from(validMp3Bytes(4 * 1024 * 1024 + 1)).toString('base64') },
    { audioBase64: Buffer.alloc(1024).toString('base64') }
  ];
  invalidPayloads.forEach((changes, index) => {
    const harness = makeHarness();
    const result = plain(harness.api.saveImportAudioItem(audioPayload({
      jobId: `audio-invalid-${index}`, itemId: 'item',
      idempotencyKey: `audio-invalid-${index}:item`, ...changes
    })));
    assert.equal(result.ok, false);
    assert.equal(result.errorCode, 'INVALID_AUDIO_PAYLOAD');
    assert.equal(harness.audit.drive.creates, 0);
    assert.equal(harness.audit.locks.attempts, 0);
  });
});

test('audio validation and provider failures never expose Base64, edit token, or source Drive ID', () => {
  const harness = makeHarness();
  const secretToken = 'edit-token-SECRET-987';
  const secretDriveId = 'source_audio_SECRET_12345';
  const invalidSecretBase64 = 'not-canonical-base64-SECRET';
  const secretBase64 = Buffer.from(validMp3Bytes()).toString('base64');
  const validationResult = plain(harness.api.saveImportAudioItem(audioPayload({
    __editToken: secretToken,
    sourceKind: 'drive', sourceDriveFileId: secretDriveId,
    audioBase64: invalidSecretBase64
  })));
  const providerResult = plain(harness.api.saveImportAudioItem(audioPayload({
    __editToken: secretToken,
    sourceKind: 'drive', sourceDriveFileId: secretDriveId,
    audioBase64: secretBase64
  })));
  assert.equal(validationResult.errorCode, 'INVALID_AUDIO_PAYLOAD');
  assert.equal(providerResult.errorCode, 'IMPORT_AUDIO_SOURCE_CHECK_FAILED');
  const observable = JSON.stringify({
    validationResult, providerResult, errors: harness.audit.errors,
    sheets: Array.from(harness.sheets.values()).map((s) => s.rows)
  });
  assert.equal(observable.includes(secretToken), false);
  assert.equal(observable.includes(secretDriveId), false);
  assert.equal(observable.includes(invalidSecretBase64), false);
  assert.equal(observable.includes(secretBase64), false);
});

test('create-pin saves one private managed MP3 and returns only safe pin projection', () => {
  const harness = makeHarness();
  const result = plain(harness.api.saveImportAudioItem(audioPayload()));
  assert.equal(result.ok, true);
  assert.equal(result.pin.hasAudio, true);
  assert.equal(Object.hasOwn(result.pin, 'audioId'), false);
  assert.equal(result.pin.title, 'voice');
  assert.equal(harness.liveManagedAudioFiles().length, 1);
  assert.equal(harness.audit.drive.shares, 0);
  assert.equal(receiptValue(harness, 'mediaKind'), 'audio');
  assert.equal(receiptValue(harness, 'operationMode'), 'create-pin');
  assert.match(receiptValue(harness, 'fileId'), /^managed_audio_/);
  assert.equal(receiptValue(harness, 'imageUrl'), '');
  assert.equal(harness.audit.locks.nestedAttempts, 0);
  assert.equal(harness.audit.events.find((event) => event.type === 'drive-create').lockHeld, false);
});

test('readPinAudioBlobByPinId_ resolves the server-only managed audio ID without returning it', () => {
  const harness = makeHarness();
  const result = plain(harness.api.saveImportAudioItem(audioPayload()));
  const blob = harness.api.readPinAudioBlobByPinId_(result.pin.id);
  assert.equal(blob.getContentType(), 'audio/mpeg');
  assert.deepEqual(blob.getBytes().slice(0, 3), [0x49, 0x44, 0x33]);
  assert.equal(JSON.stringify(result).includes(receiptValue(harness, 'fileId')), false);
});
