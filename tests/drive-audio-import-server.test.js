const assert = require('node:assert/strict');
const test = require('node:test');
const { createHarness } = require('./drive-photo-import-server-harness');

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

function seed(harness) {
  const result = harness.api.ensureMediaDriveStructure(harness.tokenPayload({}));
  assert.deepEqual(plain(result), { ok: true });
  const original = harness.folderId('original');
  const ids = {
    photos: harness.folderId('photos'),
    audio: harness.folderId('audio'),
    original,
    originalPhotos: harness.folderId('photos', original),
    originalAudio: harness.folderId('audio', original)
  };
  harness.audit.writes.length = 0;
  return ids;
}

test('media inbox classifies only supported direct-root photos and audio with exact descriptors', () => {
  const managedAudioId = 'managed_audio_AAAA';
  const mapHeaders = ['ID', 'ファイルID', '音声ID'];
  const harness = createHarness({ sheets: {
    map_info: [mapHeaders, ['pin-audio', '', managedAudioId]]
  } });
  const ids = seed(harness);
  harness.addFile('photo_JPGAAAAAAAA', 'photo.jpg', 'image/jpeg', [1], [harness.rootId]);
  harness.addFile('photo_HEIFAAAAAAA', 'photo.heif', 'image/heif', [2], [harness.rootId]);
  harness.addFile('audio_M4AAAAAAAAAA', 'voice.m4a', 'audio/mp4', [3], [harness.rootId]);
  harness.addFile('audio_MP3AAAAAAAAA', 'voice.mp3', 'audio/mpeg', [4], [harness.rootId]);
  harness.addFile('audio_WAVAAAAAAAAA', 'voice.wav', 'audio/wav', [5], [harness.rootId]);
  harness.addFile(managedAudioId, 'managed.mp3', 'audio/mpeg', [6], [harness.rootId]);
  harness.addFile('nested_audio_AAAAA', 'nested.mp3', 'audio/mpeg', [7], [ids.audio]);
  harness.addFile('archive_audio_AAAA', 'archive.wav', 'audio/wav', [8], [ids.originalAudio]);
  harness.addFile('trashed_audio_AAAA', 'trash.mp3', 'audio/mpeg', [9], [harness.rootId], { trashed: true });
  harness.addFile('shortcut_audio_AAA', 'shortcut.mp3', 'application/vnd.google-apps.shortcut', [10], [harness.rootId]);
  harness.addFile('unsupported_AAAAA', 'notes.txt', 'text/plain', [11], [harness.rootId]);

  const photo = harness.api.listDriveMediaInbox(harness.tokenPayload({ mediaKind: 'photo' }));
  const audio = harness.api.listDriveMediaInbox(harness.tokenPayload({ mediaKind: 'audio' }));

  assert.deepEqual(plain(photo), {
    ok: true,
    items: [
      { id: 'photo_HEIFAAAAAAA', name: 'photo.heif', mimeType: 'image/heif', sizeBytes: 1,
        modifiedAt: '2026-07-12T01:02:03.000Z', kind: 'heic' },
      { id: 'photo_JPGAAAAAAAA', name: 'photo.jpg', mimeType: 'image/jpeg', sizeBytes: 1,
        modifiedAt: '2026-07-12T01:02:03.000Z', kind: 'jpeg' }
    ]
  });
  assert.deepEqual(plain(audio), {
    ok: true,
    items: [
      { id: 'audio_M4AAAAAAAAAA', name: 'voice.m4a', mimeType: 'audio/mp4', sizeBytes: 1,
        modifiedAt: '2026-07-12T01:02:03.000Z', kind: 'm4a' },
      { id: 'audio_MP3AAAAAAAAA', name: 'voice.mp3', mimeType: 'audio/mpeg', sizeBytes: 1,
        modifiedAt: '2026-07-12T01:02:03.000Z', kind: 'mp3' },
      { id: 'audio_WAVAAAAAAAAA', name: 'voice.wav', mimeType: 'audio/wav', sizeBytes: 1,
        modifiedAt: '2026-07-12T01:02:03.000Z', kind: 'wav' }
    ]
  });
  assert.deepEqual(harness.audit.writes, []);
});

test('audio read accepts M4A, MP3, and WAV and returns exact actual bytes', () => {
  for (const [name, mimeType, kind] of [
    ['voice.m4a', 'audio/mp4', 'm4a'],
    ['voice.mp3', 'audio/mpeg', 'mp3'],
    ['voice.wav', 'audio/wav', 'wav']
  ]) {
    const harness = createHarness();
    seed(harness);
    const fileId = `audio_${kind}_AAAAAAAAAA`;
    harness.addFile(fileId, name, mimeType, [1, 2, 3], [harness.rootId], { metadataSize: 1 });

    const result = harness.api.readDriveAudioImportFile(harness.tokenPayload({ fileId }));

    assert.equal(result.ok, true);
    assert.deepEqual(plain(result.file), {
      id: fileId,
      name,
      mimeType,
      sizeBytes: 3,
      modifiedAt: '2026-07-12T01:02:03.000Z',
      kind,
      base64: 'AQID'
    });
    assert.equal(harness.audit.blobReads, 1);
    assert.equal(harness.audit.byteReads, 1);
    assert.deepEqual(harness.audit.writes, []);
  }
});

test('audio read rejects MIME-extension mismatches, empty data, and enforces the actual-byte 15MiB boundary', () => {
  const mismatch = createHarness();
  seed(mismatch);
  mismatch.addFile('mismatch_audio_AAA', 'voice.mp3', 'audio/wav', [1], [mismatch.rootId]);
  assert.equal(mismatch.api.readDriveAudioImportFile(
    mismatch.tokenPayload({ fileId: 'mismatch_audio_AAA' })
  ).errorCode, 'DRIVE_AUDIO_FILE_TYPE_UNSUPPORTED');

  const empty = createHarness();
  seed(empty);
  empty.addFile('empty_audio_AAAAA', 'empty.wav', 'audio/wav', [], [empty.rootId], { metadataSize: 123 });
  assert.equal(empty.api.readDriveAudioImportFile(
    empty.tokenPayload({ fileId: 'empty_audio_AAAAA' })
  ).errorCode, 'DRIVE_AUDIO_FILE_EMPTY');

  const exact = createHarness();
  seed(exact);
  exact.addFile('exact_audio_AAAAA', 'exact.mp3', 'audio/mpeg', new Uint8Array(15 * 1024 * 1024),
    [exact.rootId], { metadataSize: 1 });
  assert.equal(exact.api.readDriveAudioImportFile(
    exact.tokenPayload({ fileId: 'exact_audio_AAAAA' })
  ).file.sizeBytes, 15 * 1024 * 1024);

  const over = createHarness();
  seed(over);
  over.addFile('large_audio_AAAAA', 'large.mp3', 'audio/mpeg', new Uint8Array(15 * 1024 * 1024 + 1),
    [over.rootId], { metadataSize: 1 });
  const overResult = over.api.readDriveAudioImportFile(over.tokenPayload({ fileId: 'large_audio_AAAAA' }));
  assert.equal(overResult.errorCode, 'DRIVE_AUDIO_FILE_TOO_LARGE');
  assert.equal(over.audit.blobReads, 1);
  assert.equal(over.audit.byteReads, 1);
});

test('audio read is root-direct-only and revalidates trash, shortcut, and replacement state', () => {
  const harness = createHarness();
  const ids = seed(harness);
  harness.addFile('nested_audio_AAAAA', 'nested.mp3', 'audio/mpeg', [1], [ids.audio]);
  harness.addFile('archive_audio_AAAA', 'archive.mp3', 'audio/mpeg', [1], [ids.originalAudio]);
  harness.addFile('trashed_audio_AAAA', 'trash.mp3', 'audio/mpeg', [1], [harness.rootId], { trashed: true });
  harness.addFile('shortcut_audio_AAA', 'shortcut.mp3', 'application/vnd.google-apps.shortcut', [1], [harness.rootId]);
  harness.addFile('replace_audio_AAAA', 'replace.mp3', 'audio/mpeg', [1], [harness.rootId]);

  const listing = harness.api.listDriveMediaInbox(harness.tokenPayload({ mediaKind: 'audio' }));
  assert.equal(listing.items.some((item) => item.id === 'replace_audio_AAAA'), true);
  harness.replaceFile('replace_audio_AAAA', 'replace.wav', 'audio/wav', [1], [ids.audio]);

  assert.equal(harness.api.readDriveAudioImportFile(
    harness.tokenPayload({ fileId: 'nested_audio_AAAAA' })
  ).errorCode, 'DRIVE_AUDIO_FILE_OUTSIDE_INBOX');
  assert.equal(harness.api.readDriveAudioImportFile(
    harness.tokenPayload({ fileId: 'archive_audio_AAAA' })
  ).errorCode, 'DRIVE_AUDIO_FILE_OUTSIDE_INBOX');
  assert.equal(harness.api.readDriveAudioImportFile(
    harness.tokenPayload({ fileId: 'trashed_audio_AAAA' })
  ).errorCode, 'DRIVE_AUDIO_FILE_NOT_FOUND');
  assert.equal(harness.api.readDriveAudioImportFile(
    harness.tokenPayload({ fileId: 'shortcut_audio_AAA' })
  ).errorCode, 'DRIVE_AUDIO_FILE_TYPE_UNSUPPORTED');
  assert.equal(harness.api.readDriveAudioImportFile(
    harness.tokenPayload({ fileId: 'replace_audio_AAAA' })
  ).errorCode, 'DRIVE_AUDIO_FILE_OUTSIDE_INBOX');
});

test('audio read rejects completed root sources and managed root audio IDs just like listing', () => {
  const receiptHeaders = [
    'idempotencyKey', 'jobId', 'itemId', 'payloadHash', 'state', 'leaseOwner',
    'leaseUntil', 'pinId', 'targetFolderId', 'tempFileName', 'fileId', 'imageUrl',
    'folderUrl', 'createdAt', 'updatedAt', 'lastErrorCode', 'sourceDriveFileId',
    'mediaKind', 'operationMode', 'targetPinId', 'cleanupFileId'
  ];
  const completed = Array(receiptHeaders.length).fill('');
  completed[receiptHeaders.indexOf('state')] = 'completed';
  completed[receiptHeaders.indexOf('pinId')] = 'pin-completed-audio';
  completed[receiptHeaders.indexOf('sourceDriveFileId')] = 'completed_audio_AA';
  completed[receiptHeaders.indexOf('mediaKind')] = 'audio';
  const harness = createHarness({ sheets: {
    map_info: [['ID', 'ファイルID', '音声ID'], ['pin-managed-audio', '', 'managed_audio_AAAA']],
    import_receipts: [receiptHeaders, completed]
  } });
  seed(harness);
  harness.addFile('completed_audio_AA', 'completed.mp3', 'audio/mpeg', [1], [harness.rootId]);
  harness.addFile('managed_audio_AAAA', 'managed.wav', 'audio/wav', [2], [harness.rootId]);

  assert.equal(harness.api.readDriveAudioImportFile(
    harness.tokenPayload({ fileId: 'completed_audio_AA' })
  ).errorCode, 'DRIVE_AUDIO_FILE_OUTSIDE_INBOX');
  assert.equal(harness.api.readDriveAudioImportFile(
    harness.tokenPayload({ fileId: 'managed_audio_AAAA' })
  ).errorCode, 'DRIVE_AUDIO_FILE_OUTSIDE_INBOX');
  assert.equal(harness.audit.blobReads, 0);
});

test('media APIs reject invalid kinds and unauthenticated requests before Drive writes', () => {
  const harness = createHarness();
  assert.equal(harness.api.listDriveMediaInbox({ __editToken: 'bad', mediaKind: 'audio' }).errorCode,
    'DRIVE_MEDIA_ACCESS_DENIED');
  assert.equal(harness.api.readDriveAudioImportFile({ __editToken: 'bad', fileId: 'audio_AAAAAAAAAAA' }).errorCode,
    'DRIVE_MEDIA_ACCESS_DENIED');
  assert.equal(harness.api.listDriveMediaInbox(harness.tokenPayload({ mediaKind: 'video' })).errorCode,
    'DRIVE_MEDIA_KIND_INVALID');
  assert.deepEqual(harness.audit.writes, []);
});
