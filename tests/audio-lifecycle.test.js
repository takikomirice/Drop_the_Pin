const assert = require('node:assert/strict');
const test = require('node:test');
const {
  makeHarness, audioPayload, pinRow, MAP_HEADERS, RECEIPT_HEADERS
} = require('./audio-storage-harness');

function plain(value) { return JSON.parse(JSON.stringify(value)); }
function sheetsWithPin(row) {
  return {
    map_info: [MAP_HEADERS, row],
    import_receipts: [RECEIPT_HEADERS],
    config: [
      ['設定項目', '値', '説明'],
      ['IMAGE_DRIVE_URL', 'https://drive.google.com/drive/folders/root_media_1234567890', ''],
      ['EDIT_KEY', 'key', '']
    ]
  };
}
function existingRequest(mode, overrides = {}) {
  return audioPayload({
    jobId: `${mode}-job`, itemId: 'item', idempotencyKey: `${mode}-job:item`,
    operationMode: mode, targetPinId: 'pin-existing-0001',
    expectedUpdatedAt: '2026/07/22 11:59:00',
    pin: { title: 'ignored', lat: -10, lng: -20 }, ...overrides
  });
}

function photoOwnershipReceipt(pinId, photoId, targetFolderId) {
  const row = Array(RECEIPT_HEADERS.length).fill('');
  row[RECEIPT_HEADERS.indexOf('idempotencyKey')] = `photo-owner-${pinId}`;
  row[RECEIPT_HEADERS.indexOf('payloadHash')] = `photo-hash-${pinId}`;
  row[RECEIPT_HEADERS.indexOf('state')] = 'completed';
  row[RECEIPT_HEADERS.indexOf('pinId')] = pinId;
  row[RECEIPT_HEADERS.indexOf('targetFolderId')] = targetFolderId;
  row[RECEIPT_HEADERS.indexOf('fileId')] = photoId;
  row[RECEIPT_HEADERS.indexOf('mediaKind')] = 'photo';
  return row;
}

function audioCleanupRows(harness, operationMode, pinId) {
  return harness.sheets.get('import_receipts').rows.slice(1).filter((row) =>
    row[RECEIPT_HEADERS.indexOf('operationMode')] === operationMode
      && row[RECEIPT_HEADERS.indexOf('targetPinId')] === pinId);
}

test('client import cannot preempt the internal audio cleanup receipt namespace', () => {
  const targetPinId = 'pin-cleanup-namespace';
  const targetAudioId = 'managed_cleanup_namespace_1';
  const harness = makeHarness({
    sheets: sheetsWithPin(pinRow({ id: targetPinId, audioId: targetAudioId }))
  });
  harness.addFile(
    targetAudioId, 'target.mp3', 'audio/mpeg', [0x49, 0x44, 0x33],
    [harness.audioFolder.id]
  );
  const legacyCleanupPayload = JSON.stringify({
    kind: 'audio-cleanup', operationMode: 'delete-pin-audio',
    pinId: targetPinId, ownerValue: ''
  });
  const clientRequest = audioPayload({
    jobId: 'audio-cleanup-key',
    itemId: legacyCleanupPayload,
    idempotencyKey: `audio-cleanup-key:${legacyCleanupPayload}`
  });
  const clientReceipt = plain(harness.api.saveImportAudioItem(clientRequest));
  assert.equal(clientReceipt.ok, true);

  const result = plain(harness.api.deletePin({
    __editToken: 'valid-token', id: targetPinId
  }));
  assert.equal(result.ok, true);
  assert.equal(harness.mapRow(targetPinId), undefined);
  assert.equal(harness.fileExists(targetAudioId), false);
  const cleanupRows = audioCleanupRows(harness, 'delete-pin-audio', targetPinId);
  assert.equal(cleanupRows.length, 1);
  const identityMetadata = JSON.stringify([
    cleanupRows[0][RECEIPT_HEADERS.indexOf('idempotencyKey')],
    cleanupRows[0][RECEIPT_HEADERS.indexOf('jobId')],
    cleanupRows[0][RECEIPT_HEADERS.indexOf('itemId')],
    cleanupRows[0][RECEIPT_HEADERS.indexOf('payloadHash')]
  ]);
  [targetAudioId, clientRequest.__editToken, clientRequest.audioBase64].forEach((secret) => {
    assert.equal(identityMetadata.includes(secret), false);
    assert.equal(JSON.stringify({ result, errors: harness.audit.errors }).includes(secret), false);
  });
});

test('attach-existing-pin preserves all metadata and rejects CAS or existing audio before Drive writes', () => {
  const original = pinRow();
  const harness = makeHarness({ sheets: sheetsWithPin(original) });
  const result = plain(harness.api.saveImportAudioItem(existingRequest('attach-existing-pin')));
  assert.equal(result.ok, true);
  assert.equal(result.pin.lat, 35.1);
  assert.equal(result.pin.lng, 139.2);
  assert.equal(result.pin.title, '既存ピン');
  assert.equal(result.pin.description, '維持する説明');
  assert.equal(result.pin.hasAudio, true);
  assert.equal(Object.hasOwn(result.pin, 'audioId'), false);

  const stale = makeHarness({ sheets: sheetsWithPin(pinRow()) });
  const staleResult = plain(stale.api.saveImportAudioItem(existingRequest('attach-existing-pin', {
    expectedUpdatedAt: 'stale timestamp'
  })));
  assert.equal(staleResult.ok, false);
  assert.equal(staleResult.errorCode, 'PIN_AUDIO_CONFLICT');
  assert.equal(stale.audit.drive.creates, 0);
  assert.equal(stale.audit.locks.attempts, 0);

  const occupied = makeHarness({ sheets: sheetsWithPin(pinRow({ audioId: 'old_audio_123456789' })) });
  const occupiedResult = plain(occupied.api.saveImportAudioItem(existingRequest('attach-existing-pin')));
  assert.equal(occupiedResult.ok, false);
  assert.equal(occupiedResult.errorCode, 'PIN_AUDIO_ALREADY_ATTACHED');
  assert.equal(occupied.audit.drive.creates, 0);
});

test('replace requires existing audio and keeps the old MP3 until the new ID is durably linked', () => {
  const oldAudioId = 'old_managed_audio_1234';
  const harness = makeHarness({
    sheets: sheetsWithPin(pinRow({ audioId: oldAudioId })),
    failMapAudioWriteOnce: true
  });
  harness.addFile(oldAudioId, 'old.mp3', 'audio/mpeg', [0x49, 0x44, 0x33], [harness.audioFolder.id]);
  const failure = plain(harness.api.saveImportAudioItem(existingRequest('replace-existing-audio')));
  assert.equal(failure.ok, false);
  assert.equal(failure.pin && failure.pin.hasAudio, true);
  assert.equal(harness.fileExists(oldAudioId), true);
  const trashEvent = harness.audit.events.find((event) => event.type === 'drive-trash' && event.id === oldAudioId);
  assert.equal(trashEvent, undefined);

  const missing = makeHarness({ sheets: sheetsWithPin(pinRow()) });
  const missingResult = plain(missing.api.saveImportAudioItem(existingRequest('replace-existing-audio')));
  assert.equal(missingResult.ok, false);
  assert.equal(missingResult.errorCode, 'PIN_AUDIO_MISSING');
  assert.equal(missing.audit.drive.creates, 0);
});

test('stale existing-pin audio requests report CAS conflict before attachment-state errors', () => {
  const staleAttach = makeHarness({
    sheets: sheetsWithPin(pinRow({ audioId: 'old_audio_123456789' }))
  });
  const staleAttachResult = plain(staleAttach.api.saveImportAudioItem(existingRequest(
    'attach-existing-pin',
    { expectedUpdatedAt: 'stale timestamp' }
  )));
  assert.equal(staleAttachResult.ok, false);
  assert.equal(staleAttachResult.errorCode, 'PIN_AUDIO_CONFLICT');
  assert.equal(staleAttach.audit.drive.creates, 0);

  const staleReplace = makeHarness({ sheets: sheetsWithPin(pinRow()) });
  const staleReplaceResult = plain(staleReplace.api.saveImportAudioItem(existingRequest(
    'replace-existing-audio',
    { expectedUpdatedAt: 'stale timestamp' }
  )));
  assert.equal(staleReplaceResult.ok, false);
  assert.equal(staleReplaceResult.errorCode, 'PIN_AUDIO_CONFLICT');
  assert.equal(staleReplace.audit.drive.creates, 0);
});

test('replace cleanup failure returns linked success and replay removes only the old managed MP3', () => {
  const oldAudioId = 'old_managed_audio_1234';
  const harness = makeHarness({
    sheets: sheetsWithPin(pinRow({ audioId: oldAudioId })),
    failTrashOnceId: oldAudioId
  });
  harness.addFile(oldAudioId, 'old.mp3', 'audio/mpeg', [0x49, 0x44, 0x33], [harness.audioFolder.id]);
  const request = existingRequest('replace-existing-audio');
  const first = plain(harness.api.saveImportAudioItem(request));
  assert.equal(first.ok, true);
  assert.equal(first.cleanupRequired, true);
  assert.equal(harness.fileExists(oldAudioId), true);
  assert.equal(harness.receipt()[RECEIPT_HEADERS.indexOf('state')], 'cleanup_pending');
  const linkedEventIndex = harness.audit.events.findIndex((event) =>
    event.type === 'sheet-write' && event.sheet === 'map_info' && event.column === 16);
  const trashEventIndex = harness.audit.events.findIndex((event) => event.type === 'drive-trash' && event.id === oldAudioId);
  assert.equal(linkedEventIndex >= 0 && trashEventIndex > linkedEventIndex, true);

  const replay = plain(harness.api.saveImportAudioItem(request));
  assert.equal(replay.ok, true);
  assert.equal(replay.cleanupRequired, false);
  assert.equal(harness.fileExists(oldAudioId), false);
  assert.equal(harness.audit.drive.creates, 1);
});

test('existing-pin replay repairs an audio link whose updatedAt write lost its response', () => {
  const harness = makeHarness({
    sheets: sheetsWithPin(pinRow()),
    failMapAudioTimestampOnce: true
  });
  const request = existingRequest('attach-existing-pin');
  const first = plain(harness.api.saveImportAudioItem(request));
  assert.equal(first.ok, false);
  assert.notEqual(harness.mapRow('pin-existing-0001')[15], '');
  assert.equal(harness.mapRow('pin-existing-0001')[13], '2026/07/22 11:59:00');

  const replay = plain(harness.api.saveImportAudioItem(request));
  assert.equal(replay.ok, true);
  assert.equal(replay.deduplicated, true);
  assert.notEqual(replay.pin.updatedAt, '2026/07/22 11:59:00');
  assert.equal(harness.audit.drive.creates, 1);
});

test('removePinAudio clears the sheet first, preserves metadata, and deletes only managed audio', () => {
  const managedId = 'managed_remove_audio_123';
  const harness = makeHarness({ sheets: sheetsWithPin(pinRow({ audioId: managedId })) });
  harness.addFile(managedId, 'remove.mp3', 'audio/mpeg', [0x49, 0x44, 0x33], [harness.audioFolder.id]);
  const result = plain(harness.api.removePinAudio({
    __editToken: 'valid-token', pinId: 'pin-existing-0001',
    expectedUpdatedAt: '2026/07/22 11:59:00'
  }));
  assert.equal(result.ok, true);
  assert.equal(result.pin.hasAudio, false);
  assert.equal(result.pin.title, '既存ピン');
  assert.equal(Object.hasOwn(result.pin, 'audioId'), false);
  assert.equal(harness.mapRow('pin-existing-0001')[15], '');
  assert.equal(harness.fileExists(managedId), false);
  const clearIndex = harness.audit.events.findIndex((event) =>
    event.type === 'sheet-write' && event.sheet === 'map_info' && event.column === 16);
  const trashIndex = harness.audit.events.findIndex((event) => event.type === 'drive-trash' && event.id === managedId);
  assert.equal(clearIndex >= 0 && trashIndex > clearIndex, true);

  const originalId = 'original_audio_file_123';
  const original = makeHarness({ sheets: sheetsWithPin(pinRow({ audioId: originalId })) });
  original.addFile(originalId, 'source.mp3', 'audio/mpeg', [0x49, 0x44, 0x33], [original.originalAudioFolder.id]);
  const originalResult = plain(original.api.removePinAudio({
    __editToken: 'valid-token', pinId: 'pin-existing-0001',
    expectedUpdatedAt: '2026/07/22 11:59:00'
  }));
  assert.equal(originalResult.ok, true);
  assert.equal(original.fileExists(originalId), true);
});

test('removePinAudio rejects stale CAS before Drive mutation', () => {
  const audioId = 'managed_remove_audio_123';
  const harness = makeHarness({ sheets: sheetsWithPin(pinRow({ audioId })) });
  harness.addFile(audioId, 'remove.mp3', 'audio/mpeg', [0x49, 0x44, 0x33], [harness.audioFolder.id]);
  const result = plain(harness.api.removePinAudio({
    __editToken: 'valid-token', pinId: 'pin-existing-0001', expectedUpdatedAt: 'stale'
  }));
  assert.equal(result.ok, false);
  assert.equal(result.errorCode, 'PIN_AUDIO_CONFLICT');
  assert.equal(harness.fileExists(audioId), true);
  assert.equal(harness.audit.drive.trashes, 0);
  assert.equal(harness.audit.locks.attempts, 0);
});

test('removePinAudio reconciles a cleared audio ID after timestamp response loss and replays', () => {
  const audioId = 'managed_remove_response_12';
  const harness = makeHarness({
    sheets: sheetsWithPin(pinRow({ audioId })),
    failMapAudioTimestampOnce: true
  });
  harness.addFile(audioId, 'remove.mp3', 'audio/mpeg', [0x49, 0x44, 0x33], [harness.audioFolder.id]);
  const request = {
    __editToken: 'valid-token', pinId: 'pin-existing-0001',
    expectedUpdatedAt: '2026/07/22 11:59:00'
  };

  const first = plain(harness.api.removePinAudio(request));
  assert.equal(first.ok, true);
  assert.equal(first.cleanupRequired, false);
  assert.equal(first.pin.hasAudio, false);
  assert.notEqual(first.pin.updatedAt, request.expectedUpdatedAt);
  assert.equal(harness.fileExists(audioId), false);

  const replay = plain(harness.api.removePinAudio(request));
  assert.equal(replay.ok, true);
  assert.equal(replay.cleanupRequired, false);
  assert.equal(replay.pin.hasAudio, false);
  assert.equal(harness.audit.drive.trashes, 1);
  const cleanupRows = harness.sheets.get('import_receipts').rows.slice(1).filter((row) =>
    row[RECEIPT_HEADERS.indexOf('operationMode')] === 'remove-pin-audio');
  assert.equal(cleanupRows.length, 1);
  assert.equal(cleanupRows[0][RECEIPT_HEADERS.indexOf('state')], 'completed');
});

test('removePinAudio journals trash-once cleanup and converges on the same request', () => {
  const audioId = 'managed_remove_trash_once';
  const harness = makeHarness({
    sheets: sheetsWithPin(pinRow({ audioId })), failTrashOnceId: audioId
  });
  harness.addFile(audioId, 'remove.mp3', 'audio/mpeg', [0x49, 0x44, 0x33], [harness.audioFolder.id]);
  const request = {
    __editToken: 'valid-token', pinId: 'pin-existing-0001',
    expectedUpdatedAt: '2026/07/22 11:59:00'
  };
  const first = plain(harness.api.removePinAudio(request));
  assert.equal(first.ok, true);
  assert.equal(first.cleanupRequired, true);
  assert.equal(harness.fileExists(audioId), true);
  assert.equal(audioCleanupRows(harness, 'remove-pin-audio', 'pin-existing-0001')[0][4], 'cleanup_pending');

  const replay = plain(harness.api.removePinAudio(request));
  assert.equal(replay.ok, true);
  assert.equal(replay.cleanupRequired, false);
  assert.equal(harness.fileExists(audioId), false);
  assert.equal(audioCleanupRows(harness, 'remove-pin-audio', 'pin-existing-0001')[0][4], 'completed');
});

test('removePinAudio old replay drains A without unlinking reattached B and a new CAS removes B', () => {
  const audioA = 'managed_remove_pending_AA';
  const audioB = 'managed_remove_reattached_B';
  const harness = makeHarness({
    sheets: sheetsWithPin(pinRow({ audioId: audioA })), failTrashOnceId: audioA
  });
  harness.addFile(audioA, 'a.mp3', 'audio/mpeg', [0x49, 0x44, 0x33], [harness.audioFolder.id]);
  harness.addFile(audioB, 'b.mp3', 'audio/mpeg', [0x49, 0x44, 0x33], [harness.audioFolder.id]);
  const oldRequest = {
    __editToken: 'valid-token', pinId: 'pin-existing-0001',
    expectedUpdatedAt: '2026/07/22 11:59:00'
  };
  const first = plain(harness.api.removePinAudio(oldRequest));
  assert.equal(first.ok, true);
  assert.equal(first.cleanupRequired, true);
  harness.mapRow('pin-existing-0001')[15] = audioB;
  harness.mapRow('pin-existing-0001')[13] = '2026/07/22 12:30:00';

  const oldReplay = plain(harness.api.removePinAudio(oldRequest));
  assert.equal(oldReplay.ok, true);
  assert.equal(oldReplay.cleanupRequired, false);
  assert.equal(oldReplay.pin.hasAudio, true);
  assert.equal(Object.hasOwn(oldReplay.pin, 'audioId'), false);
  assert.equal(harness.mapRow('pin-existing-0001')[15], audioB);
  assert.equal(harness.fileExists(audioA), false);
  assert.equal(harness.fileExists(audioB), true);

  const newRequest = {
    __editToken: 'valid-token', pinId: 'pin-existing-0001',
    expectedUpdatedAt: '2026/07/22 12:30:00'
  };
  const freshRemove = plain(harness.api.removePinAudio(newRequest));
  assert.equal(freshRemove.ok, true);
  assert.equal(freshRemove.pin.hasAudio, false);
  assert.equal(harness.fileExists(audioB), false);
  const journals = audioCleanupRows(harness, 'remove-pin-audio', 'pin-existing-0001');
  assert.equal(journals.length, 2);
  assert.equal(journals.every((row) => row[RECEIPT_HEADERS.indexOf('state')] === 'completed'), true);
});

test('removePinAudio old replay preserves A when another pin references it and preserves B', () => {
  const audioA = 'managed_remove_shared_AAA';
  const audioB = 'managed_remove_shared_BBB';
  const sheets = sheetsWithPin(pinRow({ id: 'pin-remove-shared-target', audioId: audioA }));
  sheets.map_info.push(pinRow({ id: 'pin-remove-shared-keeper', audioId: '' }));
  const harness = makeHarness({ sheets, failTrashOnceId: audioA });
  harness.addFile(audioA, 'a.mp3', 'audio/mpeg', [0x49, 0x44, 0x33], [harness.audioFolder.id]);
  harness.addFile(audioB, 'b.mp3', 'audio/mpeg', [0x49, 0x44, 0x33], [harness.audioFolder.id]);
  const oldRequest = {
    __editToken: 'valid-token', pinId: 'pin-remove-shared-target',
    expectedUpdatedAt: '2026/07/22 11:59:00'
  };
  const first = plain(harness.api.removePinAudio(oldRequest));
  assert.equal(first.ok, true);
  assert.equal(first.cleanupRequired, true);
  harness.mapRow('pin-remove-shared-target')[15] = audioB;
  harness.mapRow('pin-remove-shared-target')[13] = '2026/07/22 12:31:00';
  harness.mapRow('pin-remove-shared-keeper')[15] = audioA;

  const replay = plain(harness.api.removePinAudio(oldRequest));
  assert.equal(replay.ok, true);
  assert.equal(replay.cleanupRequired, false);
  assert.equal(replay.pin.hasAudio, true);
  assert.equal(harness.mapRow('pin-remove-shared-target')[15], audioB);
  assert.equal(harness.fileExists(audioA), true);
  assert.equal(harness.fileExists(audioB), true);
  assert.equal(
    audioCleanupRows(harness, 'remove-pin-audio', 'pin-remove-shared-target')[0][4],
    'completed'
  );
});

test('deletePin removes the sheet row before managed audio cleanup', () => {
  const audioId = 'managed_delete_audio_123';
  const harness = makeHarness({ sheets: sheetsWithPin(pinRow({ audioId })) });
  harness.addFile(audioId, 'delete.mp3', 'audio/mpeg', [0x49, 0x44, 0x33], [harness.audioFolder.id]);
  const result = plain(harness.api.deletePin({ __editToken: 'valid-token', id: 'pin-existing-0001' }));
  assert.equal(result.ok, true);
  assert.equal(harness.sheets.get('map_info').rows.length, 1);
  assert.equal(harness.fileExists(audioId), false);
  const deleteIndex = harness.audit.events.findIndex((event) => event.type === 'sheet-delete');
  const trashIndex = harness.audit.events.findIndex((event) => event.type === 'drive-trash' && event.id === audioId);
  assert.equal(deleteIndex >= 0 && trashIndex > deleteIndex, true);
});

test('bulkDeletePins removes rows before cleaning each managed audio attachment', () => {
  const firstAudioId = 'managed_bulk_audio_1234';
  const secondAudioId = 'managed_bulk_audio_5678';
  const sheets = sheetsWithPin(pinRow({ id: 'pin-audio-a', audioId: firstAudioId }));
  sheets.map_info.push(pinRow({ id: 'pin-audio-b', audioId: secondAudioId }));
  const harness = makeHarness({ sheets });
  harness.addFile(firstAudioId, 'first.mp3', 'audio/mpeg', [0x49, 0x44, 0x33], [harness.audioFolder.id]);
  harness.addFile(secondAudioId, 'second.mp3', 'audio/mpeg', [0x49, 0x44, 0x33], [harness.audioFolder.id]);
  const result = plain(harness.api.bulkDeletePins({
    __editToken: 'valid-token', ids: ['pin-audio-a', 'pin-audio-b']
  }));
  assert.equal(result.ok, true);
  assert.equal(result.deletedCount, 2);
  assert.equal(harness.fileExists(firstAudioId), false);
  assert.equal(harness.fileExists(secondAudioId), false);
  const deleteIndex = harness.audit.events.findIndex((event) => event.type === 'sheet-delete');
  const trashIndexes = harness.audit.events
    .map((event, index) => ({ event, index }))
    .filter(({ event }) => event.type === 'drive-trash')
    .map(({ index }) => index);
  assert.equal(trashIndexes.length, 2);
  assert.equal(trashIndexes.every((index) => index > deleteIndex), true);
});

test('deletePin uses the locked current audio after photo trash when only audio changed', () => {
  const photoId = 'managed_photo_concurrent_1';
  const oldAudioId = 'managed_audio_concurrent_old';
  const currentAudioId = 'managed_audio_concurrent_new';
  const sheets = sheetsWithPin(pinRow({ fileId: photoId, audioId: oldAudioId }));
  sheets.import_receipts.push(photoOwnershipReceipt(
    'pin-existing-0001', photoId, 'managed_photos_12345'
  ));
  let harness;
  harness = makeHarness({
    sheets,
    afterTrash(file) {
      if (file.id === photoId) harness.mapRow('pin-existing-0001')[15] = currentAudioId;
    }
  });
  harness.addFile(photoId, 'photo.jpg', 'image/jpeg', [0xff, 0xd8, 0xff], [harness.photosFolder.id]);
  harness.addFile(oldAudioId, 'old.mp3', 'audio/mpeg', [0x49, 0x44, 0x33], [harness.audioFolder.id]);
  harness.addFile(currentAudioId, 'new.mp3', 'audio/mpeg', [0x49, 0x44, 0x33], [harness.audioFolder.id]);

  const result = plain(harness.api.deletePin({
    __editToken: 'valid-token', id: 'pin-existing-0001'
  }));
  assert.equal(result.ok, true);
  assert.equal(harness.mapRow('pin-existing-0001'), undefined);
  assert.equal(harness.fileExists(photoId), false);
  assert.equal(harness.fileExists(currentAudioId), false);
  assert.equal(harness.fileExists(oldAudioId), true);
});

test('bulkDeletePins uses each locked current audio after photo trash when only audio changed', () => {
  const photoId = 'managed_photo_concurrent_2';
  const oldAudioId = 'managed_audio_bulk_old_12';
  const currentAudioId = 'managed_audio_bulk_new_12';
  const sheets = sheetsWithPin(pinRow({
    id: 'pin-bulk-concurrent', fileId: photoId, audioId: oldAudioId
  }));
  sheets.import_receipts.push(photoOwnershipReceipt(
    'pin-bulk-concurrent', photoId, 'managed_photos_12345'
  ));
  let harness;
  harness = makeHarness({
    sheets,
    afterTrash(file) {
      if (file.id === photoId) harness.mapRow('pin-bulk-concurrent')[15] = currentAudioId;
    }
  });
  harness.addFile(photoId, 'photo.jpg', 'image/jpeg', [0xff, 0xd8, 0xff], [harness.photosFolder.id]);
  harness.addFile(oldAudioId, 'old.mp3', 'audio/mpeg', [0x49, 0x44, 0x33], [harness.audioFolder.id]);
  harness.addFile(currentAudioId, 'new.mp3', 'audio/mpeg', [0x49, 0x44, 0x33], [harness.audioFolder.id]);

  const result = plain(harness.api.bulkDeletePins({
    __editToken: 'valid-token', ids: ['pin-bulk-concurrent']
  }));
  assert.equal(result.ok, true);
  assert.equal(result.deletedCount, 1);
  assert.equal(harness.mapRow('pin-bulk-concurrent'), undefined);
  assert.equal(harness.fileExists(photoId), false);
  assert.equal(harness.fileExists(currentAudioId), false);
  assert.equal(harness.fileExists(oldAudioId), true);
});

test('deletePin creates a new cleanup generation after an A delete failure and B replacement', () => {
  const audioA = 'managed_delete_generation_A';
  const audioB = 'managed_delete_generation_B';
  const sheets = sheetsWithPin(pinRow({ id: 'pin-generation-single', audioId: audioA }));
  sheets.map_info.push(pinRow({ id: 'pin-generation-keeper', audioId: audioA }));
  const harness = makeHarness({ sheets, failMapDeleteOnce: true });
  harness.addFile(audioA, 'a.mp3', 'audio/mpeg', [0x49, 0x44, 0x33], [harness.audioFolder.id]);
  harness.addFile(audioB, 'b.mp3', 'audio/mpeg', [0x49, 0x44, 0x33], [harness.audioFolder.id]);
  const request = { __editToken: 'valid-token', id: 'pin-generation-single' };

  assert.throws(() => harness.api.deletePin(request), /private map delete failure/);
  assert.notEqual(harness.mapRow('pin-generation-single'), undefined);
  harness.mapRow('pin-generation-single')[15] = audioB;

  let replay;
  assert.doesNotThrow(() => { replay = plain(harness.api.deletePin(request)); });
  assert.equal(replay.ok, true);
  assert.equal(harness.mapRow('pin-generation-single'), undefined);
  assert.equal(harness.fileExists(audioA), true);
  assert.equal(harness.fileExists(audioB), false);
  const journals = audioCleanupRows(harness, 'delete-pin-audio', 'pin-generation-single');
  assert.equal(journals.length, 2);
  assert.equal(journals.every((row) => row[RECEIPT_HEADERS.indexOf('state')] === 'completed'), true);
});

test('bulkDeletePins creates a new cleanup generation after an A delete failure and B replacement', () => {
  const audioA = 'managed_bulk_generation_AA';
  const audioB = 'managed_bulk_generation_BB';
  const sheets = sheetsWithPin(pinRow({ id: 'pin-generation-bulk', audioId: audioA }));
  sheets.map_info.push(pinRow({ id: 'pin-generation-bulk-keeper', audioId: audioA }));
  const harness = makeHarness({ sheets, failMapDeleteRowsOnce: true });
  harness.addFile(audioA, 'a.mp3', 'audio/mpeg', [0x49, 0x44, 0x33], [harness.audioFolder.id]);
  harness.addFile(audioB, 'b.mp3', 'audio/mpeg', [0x49, 0x44, 0x33], [harness.audioFolder.id]);
  const request = { __editToken: 'valid-token', ids: ['pin-generation-bulk'] };

  const first = plain(harness.api.bulkDeletePins(request));
  assert.equal(first.ok, true);
  assert.equal(first.deletedCount, 0);
  assert.notEqual(harness.mapRow('pin-generation-bulk'), undefined);
  harness.mapRow('pin-generation-bulk')[15] = audioB;

  let replay;
  assert.doesNotThrow(() => { replay = plain(harness.api.bulkDeletePins(request)); });
  assert.equal(replay.ok, true);
  assert.equal(replay.deletedCount, 1);
  assert.equal(harness.mapRow('pin-generation-bulk'), undefined);
  assert.equal(harness.fileExists(audioA), true);
  assert.equal(harness.fileExists(audioB), false);
  const journals = audioCleanupRows(harness, 'delete-pin-audio', 'pin-generation-bulk');
  assert.equal(journals.length, 2);
  assert.equal(journals.every((row) => row[RECEIPT_HEADERS.indexOf('state')] === 'completed'), true);
});

test('deletePin does not adopt a completed old cleanup generation for a new photo-only row', () => {
  const pinId = 'pin-completed-generation-reuse';
  const audioA = 'managed_completed_generation_A';
  const harness = makeHarness({
    sheets: sheetsWithPin(pinRow({ id: pinId, audioId: audioA }))
  });
  harness.addFile(audioA, 'a.mp3', 'audio/mpeg', [0x49, 0x44, 0x33], [harness.audioFolder.id]);
  const request = { __editToken: 'valid-token', id: pinId };
  const first = plain(harness.api.deletePin(request));
  assert.equal(first.ok, true);
  assert.equal(audioCleanupRows(harness, 'delete-pin-audio', pinId).length, 1);
  assert.equal(harness.audit.locks.nestedAttempts, 0);
  assert.equal(
    harness.audit.events.find((event) => event.type === 'drive-trash' && event.id === audioA).lockHeld,
    false
  );

  harness.sheets.get('map_info').rows.push(pinRow({ id: pinId, audioId: '' }));
  const newDelete = plain(harness.api.deletePin(request));
  assert.equal(newDelete.ok, true);
  assert.equal(Object.hasOwn(newDelete, 'cleanupRequired'), false);
  assert.equal(harness.mapRow(pinId), undefined);
  assert.equal(audioCleanupRows(harness, 'delete-pin-audio', pinId).length, 1);
  assert.equal(harness.audit.locks.nestedAttempts, 0);
});

test('deletePin journals audio attached after photo trash and replays a transient trash failure', () => {
  const pinId = 'pin-none-audio-single-trash';
  const photoId = 'managed_photo_none_single_trash';
  const audioB = 'managed_audio_none_single_trash';
  const sheets = sheetsWithPin(pinRow({ id: pinId, fileId: photoId, audioId: '' }));
  sheets.import_receipts.push(photoOwnershipReceipt(pinId, photoId, 'managed_photos_12345'));
  let harness;
  harness = makeHarness({
    sheets,
    failTrashOnceId: audioB,
    afterTrash(file) {
      if (file.id === photoId) harness.mapRow(pinId)[15] = audioB;
    }
  });
  harness.addFile(photoId, 'photo.jpg', 'image/jpeg', [0xff, 0xd8, 0xff], [harness.photosFolder.id]);
  harness.addFile(audioB, 'b.mp3', 'audio/mpeg', [0x49, 0x44, 0x33], [harness.audioFolder.id]);
  const request = { __editToken: 'valid-token', id: pinId };

  const first = plain(harness.api.deletePin(request));
  assert.equal(first.ok, true);
  assert.equal(first.cleanupRequired, true);
  assert.equal(harness.mapRow(pinId), undefined);
  assert.equal(harness.fileExists(audioB), true);
  assert.equal(audioCleanupRows(harness, 'delete-pin-audio', pinId).length, 1);

  const replay = plain(harness.api.deletePin(request));
  assert.equal(replay.ok, true);
  assert.equal(replay.cleanupRequired, false);
  assert.equal(harness.fileExists(audioB), false);
});

test('bulkDeletePins journals audio attached after photo trash and replays a transient trash failure', () => {
  const pinId = 'pin-none-audio-bulk-trash';
  const photoId = 'managed_photo_none_bulk_trash';
  const audioB = 'managed_audio_none_bulk_trash';
  const sheets = sheetsWithPin(pinRow({ id: pinId, fileId: photoId, audioId: '' }));
  sheets.import_receipts.push(photoOwnershipReceipt(pinId, photoId, 'managed_photos_12345'));
  let harness;
  harness = makeHarness({
    sheets,
    failTrashOnceId: audioB,
    afterTrash(file) {
      if (file.id === photoId) harness.mapRow(pinId)[15] = audioB;
    }
  });
  harness.addFile(photoId, 'photo.jpg', 'image/jpeg', [0xff, 0xd8, 0xff], [harness.photosFolder.id]);
  harness.addFile(audioB, 'b.mp3', 'audio/mpeg', [0x49, 0x44, 0x33], [harness.audioFolder.id]);
  const request = { __editToken: 'valid-token', ids: [pinId] };

  const first = plain(harness.api.bulkDeletePins(request));
  assert.equal(first.ok, true);
  assert.equal(first.deletedCount, 1);
  assert.equal(first.cleanupRequired, true);
  assert.equal(harness.mapRow(pinId), undefined);
  assert.equal(harness.fileExists(audioB), true);
  assert.equal(audioCleanupRows(harness, 'delete-pin-audio', pinId).length, 1);

  const replay = plain(harness.api.bulkDeletePins(request));
  assert.equal(replay.ok, true);
  assert.equal(replay.deletedCount, 0);
  assert.equal(replay.cleanupRequired, false);
  assert.equal(harness.fileExists(audioB), false);
});

test('deletePin journals audio attached after photo trash before a transient structure failure', () => {
  const pinId = 'pin-none-audio-single-structure';
  const photoId = 'managed_photo_none_single_structure';
  const audioB = 'managed_audio_none_single_structure';
  const sheets = sheetsWithPin(pinRow({ id: pinId, fileId: photoId, audioId: '' }));
  sheets.import_receipts.push(photoOwnershipReceipt(pinId, photoId, 'managed_photos_12345'));
  let harness;
  harness = makeHarness({
    sheets,
    failStructureOnce: true,
    afterTrash(file) {
      if (file.id === photoId) harness.mapRow(pinId)[15] = audioB;
    }
  });
  harness.addFile(photoId, 'photo.jpg', 'image/jpeg', [0xff, 0xd8, 0xff], [harness.photosFolder.id]);
  harness.addFile(audioB, 'b.mp3', 'audio/mpeg', [0x49, 0x44, 0x33], [harness.audioFolder.id]);
  const request = { __editToken: 'valid-token', id: pinId };

  const first = plain(harness.api.deletePin(request));
  assert.equal(first.ok, true);
  assert.equal(first.cleanupRequired, true);
  assert.equal(harness.mapRow(pinId), undefined);
  assert.equal(harness.fileExists(audioB), true);
  assert.equal(audioCleanupRows(harness, 'delete-pin-audio', pinId).length, 1);

  const replay = plain(harness.api.deletePin(request));
  assert.equal(replay.ok, true);
  assert.equal(replay.cleanupRequired, false);
  assert.equal(harness.fileExists(audioB), false);
});

test('bulkDeletePins journals audio attached after photo trash before a transient structure failure', () => {
  const pinId = 'pin-none-audio-bulk-structure';
  const photoId = 'managed_photo_none_bulk_structure';
  const audioB = 'managed_audio_none_bulk_structure';
  const sheets = sheetsWithPin(pinRow({ id: pinId, fileId: photoId, audioId: '' }));
  sheets.import_receipts.push(photoOwnershipReceipt(pinId, photoId, 'managed_photos_12345'));
  let harness;
  harness = makeHarness({
    sheets,
    failStructureOnce: true,
    afterTrash(file) {
      if (file.id === photoId) harness.mapRow(pinId)[15] = audioB;
    }
  });
  harness.addFile(photoId, 'photo.jpg', 'image/jpeg', [0xff, 0xd8, 0xff], [harness.photosFolder.id]);
  harness.addFile(audioB, 'b.mp3', 'audio/mpeg', [0x49, 0x44, 0x33], [harness.audioFolder.id]);
  const request = { __editToken: 'valid-token', ids: [pinId] };

  const first = plain(harness.api.bulkDeletePins(request));
  assert.equal(first.ok, true);
  assert.equal(first.deletedCount, 1);
  assert.equal(first.cleanupRequired, true);
  assert.equal(harness.mapRow(pinId), undefined);
  assert.equal(harness.fileExists(audioB), true);
  assert.equal(audioCleanupRows(harness, 'delete-pin-audio', pinId).length, 1);

  const replay = plain(harness.api.bulkDeletePins(request));
  assert.equal(replay.ok, true);
  assert.equal(replay.deletedCount, 0);
  assert.equal(replay.cleanupRequired, false);
  assert.equal(harness.fileExists(audioB), false);
});

test('deletePin journals trash-once cleanup and replay converges after the row is gone', () => {
  const audioId = 'managed_delete_trash_once';
  const harness = makeHarness({
    sheets: sheetsWithPin(pinRow({ audioId })), failTrashOnceId: audioId
  });
  harness.addFile(audioId, 'delete.mp3', 'audio/mpeg', [0x49, 0x44, 0x33], [harness.audioFolder.id]);
  const request = { __editToken: 'valid-token', id: 'pin-existing-0001' };
  const first = plain(harness.api.deletePin(request));
  assert.equal(first.ok, true);
  assert.equal(first.cleanupRequired, true);
  assert.equal(harness.mapRow('pin-existing-0001'), undefined);
  assert.equal(harness.fileExists(audioId), true);
  const journal = audioCleanupRows(harness, 'delete-pin-audio', 'pin-existing-0001');
  assert.equal(journal.length, 1);
  assert.equal(journal[0][4], 'cleanup_pending');

  const replay = plain(harness.api.deletePin(request));
  assert.equal(replay.ok, true);
  assert.equal(replay.cleanupRequired, false);
  assert.equal(harness.fileExists(audioId), false);
  assert.equal(audioCleanupRows(harness, 'delete-pin-audio', 'pin-existing-0001')[0][4], 'completed');
});

test('bulkDeletePins journals trash-once cleanup and replay converges after the row is gone', () => {
  const audioId = 'managed_bulk_trash_once_1';
  const sheets = sheetsWithPin(pinRow({ id: 'pin-bulk-trash', audioId }));
  const harness = makeHarness({ sheets, failTrashOnceId: audioId });
  harness.addFile(audioId, 'delete.mp3', 'audio/mpeg', [0x49, 0x44, 0x33], [harness.audioFolder.id]);
  const request = { __editToken: 'valid-token', ids: ['pin-bulk-trash'] };
  const first = plain(harness.api.bulkDeletePins(request));
  assert.equal(first.ok, true);
  assert.equal(first.deletedCount, 1);
  assert.equal(first.cleanupRequired, true);
  assert.equal(harness.fileExists(audioId), true);
  const journal = audioCleanupRows(harness, 'delete-pin-audio', 'pin-bulk-trash');
  assert.equal(journal.length, 1);
  assert.equal(journal[0][4], 'cleanup_pending');

  const replay = plain(harness.api.bulkDeletePins(request));
  assert.equal(replay.ok, true);
  assert.equal(replay.cleanupRequired, false);
  assert.equal(harness.fileExists(audioId), false);
  assert.equal(audioCleanupRows(harness, 'delete-pin-audio', 'pin-bulk-trash')[0][4], 'completed');
});

test('deletePin structure failure leaves the audio-linked row intact and retry succeeds', () => {
  const audioId = 'managed_delete_structure_1';
  const harness = makeHarness({
    sheets: sheetsWithPin(pinRow({ audioId })), failStructureOnce: true
  });
  harness.addFile(audioId, 'delete.mp3', 'audio/mpeg', [0x49, 0x44, 0x33], [harness.audioFolder.id]);
  const request = { __editToken: 'valid-token', id: 'pin-existing-0001' };
  const first = plain(harness.api.deletePin(request));
  assert.equal(first.ok, false);
  assert.notEqual(harness.mapRow('pin-existing-0001'), undefined);
  assert.equal(harness.fileExists(audioId), true);

  const replay = plain(harness.api.deletePin(request));
  assert.equal(replay.ok, true);
  assert.equal(harness.mapRow('pin-existing-0001'), undefined);
  assert.equal(harness.fileExists(audioId), false);
});

test('bulkDeletePins structure failure leaves every audio-linked row intact and retry succeeds', () => {
  const audioId = 'managed_bulk_structure_12';
  const sheets = sheetsWithPin(pinRow({ id: 'pin-bulk-structure', audioId }));
  const harness = makeHarness({ sheets, failStructureOnce: true });
  harness.addFile(audioId, 'delete.mp3', 'audio/mpeg', [0x49, 0x44, 0x33], [harness.audioFolder.id]);
  const request = { __editToken: 'valid-token', ids: ['pin-bulk-structure'] };
  const first = plain(harness.api.bulkDeletePins(request));
  assert.equal(first.ok, false);
  assert.notEqual(harness.mapRow('pin-bulk-structure'), undefined);
  assert.equal(harness.fileExists(audioId), true);

  const replay = plain(harness.api.bulkDeletePins(request));
  assert.equal(replay.ok, true);
  assert.equal(replay.deletedCount, 1);
  assert.equal(harness.mapRow('pin-bulk-structure'), undefined);
  assert.equal(harness.fileExists(audioId), false);
});

test('deletePin reaches journaled audio cleanup after post-map relation failure', () => {
  const audioId = 'managed_delete_relation_12';
  const sheets = sheetsWithPin(pinRow({ audioId }));
  sheets.route_pins = [
    ['routeId', 'pinId', 'pinOrder', 'createdAt', 'updatedAt'],
    ['route-1', 'pin-existing-0001', 1, '', '']
  ];
  const harness = makeHarness({ sheets, failRelationWriteOnce: true });
  harness.addFile(audioId, 'delete.mp3', 'audio/mpeg', [0x49, 0x44, 0x33], [harness.audioFolder.id]);
  const request = { __editToken: 'valid-token', id: 'pin-existing-0001' };
  assert.throws(() => harness.api.deletePin(request), /private relation write failure/);
  assert.equal(harness.mapRow('pin-existing-0001'), undefined);
  assert.equal(harness.fileExists(audioId), false);
  assert.equal(audioCleanupRows(harness, 'delete-pin-audio', 'pin-existing-0001')[0][4], 'completed');

  const replay = plain(harness.api.deletePin(request));
  assert.equal(replay.ok, true);
  assert.equal(harness.sheets.get('route_pins').getLastRow(), 1);
});

test('bulkDeletePins reaches every journaled audio cleanup after post-map relation failure', () => {
  const audioId = 'managed_bulk_relation_123';
  const sheets = sheetsWithPin(pinRow({ id: 'pin-bulk-relation', audioId }));
  sheets.route_pins = [
    ['routeId', 'pinId', 'pinOrder', 'createdAt', 'updatedAt'],
    ['route-1', 'pin-bulk-relation', 1, '', '']
  ];
  const harness = makeHarness({ sheets, failRelationWriteOnce: true });
  harness.addFile(audioId, 'delete.mp3', 'audio/mpeg', [0x49, 0x44, 0x33], [harness.audioFolder.id]);
  const request = { __editToken: 'valid-token', ids: ['pin-bulk-relation'] };
  assert.throws(() => harness.api.bulkDeletePins(request), /private relation write failure/);
  assert.equal(harness.mapRow('pin-bulk-relation'), undefined);
  assert.equal(harness.fileExists(audioId), false);
  assert.equal(audioCleanupRows(harness, 'delete-pin-audio', 'pin-bulk-relation')[0][4], 'completed');

  const replay = plain(harness.api.bulkDeletePins(request));
  assert.equal(replay.ok, true);
  assert.equal(harness.sheets.get('route_pins').getLastRow(), 1);
});
