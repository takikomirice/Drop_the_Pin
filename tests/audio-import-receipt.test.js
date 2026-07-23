const assert = require('node:assert/strict');
const test = require('node:test');
const {
  makeHarness, audioPayload, validMp3Bytes, RECEIPT_HEADERS
} = require('./audio-storage-harness');

function plain(value) { return JSON.parse(JSON.stringify(value)); }
function receiptValue(harness, name) {
  return harness.receipt()[RECEIPT_HEADERS.indexOf(name)];
}

test('audio receipt replay returns the same pin without creating another MP3', () => {
  const harness = makeHarness();
  const first = plain(harness.api.saveImportAudioItem(audioPayload()));
  const replay = plain(harness.api.saveImportAudioItem(audioPayload()));
  assert.equal(first.ok, true);
  assert.equal(replay.ok, true);
  assert.equal(replay.deduplicated, true);
  assert.equal(replay.pin.id, first.pin.id);
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.liveManagedAudioFiles().length, 1);
  assert.equal(receiptValue(harness, 'state'), 'completed');
});
test('audio receipt rejects a changed payload hash without a Drive write', () => {
  const harness = makeHarness();
  assert.equal(harness.api.saveImportAudioItem(audioPayload()).ok, true);
  const creates = harness.audit.drive.creates;
  const result = plain(harness.api.saveImportAudioItem(audioPayload({
    audioBase64: Buffer.from(validMp3Bytes(1024 * 1024 + 1)).toString('base64')
  })));
  assert.equal(result.ok, false);
  assert.equal(result.errorCode, 'IDEMPOTENCY_PAYLOAD_CONFLICT');
  assert.equal(harness.audit.drive.creates, creates);
});

test('response loss after sheet linkage replays from the linked row without another pin or MP3', () => {
  const harness = makeHarness({ failReceiptState: 'linked' });
  const first = plain(harness.api.saveImportAudioItem(audioPayload()));
  assert.equal(first.ok, false);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
  assert.equal(harness.liveManagedAudioFiles().length, 1);
  const replay = plain(harness.api.saveImportAudioItem(audioPayload()));
  assert.equal(replay.ok, true);
  assert.equal(replay.pin.hasAudio, true);
  assert.equal(replay.deduplicated, true);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(receiptValue(harness, 'state'), 'completed');
});

test('pre-link map failure compensates only the newly created unreferenced MP3', () => {
  const harness = makeHarness({ failMapAppendOnce: true });
  const result = plain(harness.api.saveImportAudioItem(audioPayload()));
  assert.equal(result.ok, false);
  assert.equal(harness.liveManagedAudioFiles().length, 0);
  assert.equal(harness.audit.drive.trashes, 1);
  assert.equal(harness.sheets.get('map_info').rows.length, 1);
});

test('Drive source archives to original/audio only after linkage and cleanup retry converges', () => {
  const sourceId = 'source_audio_1234567';
  const harness = makeHarness({ failMoveOnceId: sourceId });
  harness.addFile(sourceId, 'voice.wav', 'audio/wav', [1, 2, 3], [harness.rootId]);
  const request = audioPayload({ sourceKind: 'drive', sourceDriveFileId: sourceId });
  const first = plain(harness.api.saveImportAudioItem(request));
  assert.equal(first.ok, true);
  assert.equal(first.cleanupRequired, true);
  assert.equal(receiptValue(harness, 'state'), 'cleanup_pending');
  assert.equal(harness.mapRow(first.pin.id)[15], receiptValue(harness, 'fileId'));
  assert.equal(harness.files.get(sourceId).parentIds[0], harness.rootId);

  const replay = plain(harness.api.saveImportAudioItem(request));
  assert.equal(replay.ok, true);
  assert.equal(replay.cleanupRequired, false);
  assert.equal(replay.deduplicated, true);
  assert.equal(receiptValue(harness, 'state'), 'completed');
  assert.deepEqual(harness.files.get(sourceId).parentIds, [harness.originalAudioFolder.id]);
  assert.equal(harness.audit.drive.creates, 1);
});
