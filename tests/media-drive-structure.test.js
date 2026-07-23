const assert = require('node:assert/strict');
const test = require('node:test');
const { createHarness } = require('./drive-photo-import-server-harness');

const GUIDE_NAME = 'ここに直接ファイルを入れてください.txt';

function plainStructure(value) {
  return JSON.parse(JSON.stringify(value));
}

function seedExpectedStructure(harness) {
  const result = harness.api.ensureMediaDriveStructure(harness.tokenPayload({}));
  assert.deepEqual(JSON.parse(JSON.stringify(result)), { ok: true });
  return {
    photos: harness.folderId('photos'),
    audio: harness.folderId('audio'),
    original: harness.folderId('original'),
    originalPhotos: harness.folderId('photos', harness.folderId('original')),
    originalAudio: harness.folderId('audio', harness.folderId('original'))
  };
}

test('ensure creates the exact media hierarchy and an empty root guide without exposing IDs', () => {
  const harness = createHarness();

  const result = harness.api.ensureMediaDriveStructure(harness.tokenPayload({}));

  assert.deepEqual(JSON.parse(JSON.stringify(result)), { ok: true });
  assert.deepEqual(harness.directFolderNames(harness.rootId), ['audio', 'original', 'photos']);
  assert.deepEqual(harness.directFolderNames(harness.folderId('original')), ['audio', 'photos']);
  assert.deepEqual(harness.directFileNames(harness.rootId), [GUIDE_NAME]);
  assert.equal(harness.fileBytes(GUIDE_NAME).length, 0);
  assert.equal(JSON.stringify(result).includes(harness.rootId), false);
});

test('ensure is idempotent and never reads or overwrites an existing guide', () => {
  const harness = createHarness();
  seedExpectedStructure(harness);
  const guideId = Array.from(harness.files.values()).find((file) => file.__name === GUIDE_NAME).getId();
  harness.replaceFile(guideId, GUIDE_NAME, 'text/plain', [7, 8, 9], [harness.rootId]);
  harness.audit.writes.length = 0;
  const blobReadsBefore = harness.audit.blobReads;

  const result = harness.api.ensureMediaDriveStructure(harness.tokenPayload({}));

  assert.deepEqual(JSON.parse(JSON.stringify(result)), { ok: true });
  assert.deepEqual(harness.audit.writes, []);
  assert.equal(harness.audit.blobReads, blobReadsBefore);
  assert.deepEqual(harness.fileBytes(GUIDE_NAME), [7, 8, 9]);
});

test('concurrent missing checks serialize into one hierarchy and return the same IDs', () => {
  const harness = createHarness();
  const completed = [];
  harness.onNextFolderCreate(() => {
    harness.runConcurrent(() => {
      completed.push(harness.api.ensureMediaDriveStructure_());
    });
  });

  const first = harness.api.ensureMediaDriveStructure_();

  assert.equal(completed.length, 1);
  assert.deepEqual(plainStructure(completed[0]), plainStructure(first));
  assert.deepEqual(harness.directFolderNames(harness.rootId), ['audio', 'original', 'photos']);
  assert.deepEqual(harness.directFolderNames(first.original), ['audio', 'photos']);
  assert.deepEqual(harness.directFileNames(harness.rootId), [GUIDE_NAME]);
  assert.equal(harness.audit.writes.filter((entry) => entry.method === 'createFolder').length, 5);
  assert.equal(harness.audit.writes.filter((entry) => entry.method === 'createFile').length, 1);
  assert.equal(harness.audit.locks.queuedRuns, 1);
  assert.equal(harness.audit.locks.attempts, 2);
  assert.equal(harness.audit.locks.releases, 2);
  assert.equal(harness.audit.locks.maxDepth, 1);
});

test('late conflicts under existing original are preflighted before every write', () => {
  const cases = [
    (harness, original) => {
      harness.addFolder('late_photo_AAAAAA', 'photos', [original.getId()]);
      harness.addFolder('late_photo_BBBBBB', 'photos', [original.getId()]);
      return 'DRIVE_MEDIA_STRUCTURE_AMBIGUOUS';
    },
    (harness, original) => {
      harness.addFile('late_audio_file_A', 'audio', 'application/octet-stream', [1], [original.getId()]);
      return 'DRIVE_MEDIA_STRUCTURE_NAME_CONFLICT';
    }
  ];

  cases.forEach((seedConflict) => {
    const harness = createHarness();
    const original = harness.addFolder('existing_original_A', 'original', [harness.rootId]);
    const expectedCode = seedConflict(harness, original);

    const result = harness.api.ensureMediaDriveStructure(harness.tokenPayload({}));

    assert.equal(result.errorCode, expectedCode);
    assert.deepEqual(harness.audit.writes, []);
    assert.deepEqual(harness.directFolderNames(harness.rootId), ['original']);
  });
});

test('ensure rejects duplicate folders, same-name files, duplicate guides, and wrong created parents with stable codes', () => {
  const duplicate = createHarness();
  duplicate.addFolder('photos_AAAAAAAAAA', 'photos', [duplicate.rootId]);
  duplicate.addFolder('photos_BBBBBBBBBB', 'photos', [duplicate.rootId]);
  assert.equal(duplicate.api.ensureMediaDriveStructure(duplicate.tokenPayload({})).errorCode,
    'DRIVE_MEDIA_STRUCTURE_AMBIGUOUS');

  const fileConflict = createHarness();
  fileConflict.addFile('photos_file_AAAAAA', 'photos', 'application/octet-stream', [1], [fileConflict.rootId]);
  assert.equal(fileConflict.api.ensureMediaDriveStructure(fileConflict.tokenPayload({})).errorCode,
    'DRIVE_MEDIA_STRUCTURE_NAME_CONFLICT');

  const duplicateGuide = createHarness();
  duplicateGuide.addFile('guide_file_AAAAAAA', GUIDE_NAME, 'text/plain', [], [duplicateGuide.rootId]);
  duplicateGuide.addFile('guide_file_BBBBBBB', GUIDE_NAME, 'text/plain', [], [duplicateGuide.rootId]);
  assert.equal(duplicateGuide.api.ensureMediaDriveStructure(duplicateGuide.tokenPayload({})).errorCode,
    'DRIVE_MEDIA_STRUCTURE_AMBIGUOUS');

  const wrongParent = createHarness();
  const outside = wrongParent.addFolder('outside_AAAAAAAAA', 'Outside');
  wrongParent.folders.get(wrongParent.rootId).createFolder = function(name) {
    wrongParent.audit.writes.push({ method: 'createFolder', parentId: wrongParent.rootId, name: String(name) });
    return wrongParent.addFolder('wrong_parent_AAAA', String(name), [outside.getId()]);
  };
  assert.equal(wrongParent.api.ensureMediaDriveStructure(wrongParent.tokenPayload({})).errorCode,
    'DRIVE_MEDIA_STRUCTURE_PARENT_INVALID');
});

test('ensure rejects unauthenticated and missing-root calls before Drive writes', () => {
  const denied = createHarness();
  assert.equal(denied.api.ensureMediaDriveStructure({ __editToken: 'bad' }).errorCode,
    'DRIVE_MEDIA_ACCESS_DENIED');
  assert.deepEqual(denied.audit.writes, []);
  assert.deepEqual(denied.audit.folderReads, []);

  const missing = createHarness({ rootMissing: true });
  assert.equal(missing.api.ensureMediaDriveStructure(missing.tokenPayload({})).errorCode,
    'DRIVE_MEDIA_ROOT_MISSING');
  assert.deepEqual(missing.audit.writes, []);
  assert.deepEqual(missing.audit.folderReads, []);
});

test('ensure migrates only uniquely owned legacy managed JPEGs from the root into photos', () => {
  const receiptHeaders = [
    'idempotencyKey', 'jobId', 'itemId', 'payloadHash', 'state', 'leaseOwner',
    'leaseUntil', 'pinId', 'targetFolderId', 'tempFileName', 'fileId', 'imageUrl',
    'folderUrl', 'createdAt', 'updatedAt', 'lastErrorCode', 'sourceDriveFileId',
    'mediaKind', 'operationMode', 'targetPinId', 'cleanupFileId'
  ];
  const mapHeaders = [
    'タイムスタンプ', 'タイトル', '説明', '緯度', '経度', 'ピンの色',
    'ファイルID', '画像URL', 'ID', '参考URL一覧', '状態', 'タグ',
    'イベント時刻', '更新時刻', 'アイコン', '音声ID'
  ];
  const mapRow = (pinId, fileId, audioId = '') => {
    const row = Array(mapHeaders.length).fill('');
    row[mapHeaders.indexOf('ID')] = pinId;
    row[mapHeaders.indexOf('ファイルID')] = fileId;
    row[mapHeaders.indexOf('音声ID')] = audioId;
    return row;
  };
  const receipt = (state, pinId, fileId, mediaKind) => {
    const row = Array(receiptHeaders.length).fill('');
    row[receiptHeaders.indexOf('state')] = state;
    row[receiptHeaders.indexOf('pinId')] = pinId;
    row[receiptHeaders.indexOf('fileId')] = fileId;
    row[receiptHeaders.indexOf('mediaKind')] = mediaKind;
    return row;
  };
  const harness = createHarness({ sheets: {
    map_info: [mapHeaders,
      mapRow('pin-map', 'managed_map_AAAAA'),
      mapRow('pin-duplicate-a', 'managed_duplicate'),
      mapRow('pin-duplicate-b', 'managed_duplicate'),
      mapRow('pin-audio', '', 'managed_audio_AAAA')],
    import_receipts: [receiptHeaders,
      receipt('completed', 'pin-receipt', 'managed_receipt_A', 'photo'),
      receipt('failed', 'pin-failed', 'managed_failed_AAA', 'photo'),
      receipt('completed', 'pin-audio-receipt', 'managed_audio_AAAA', 'audio')]
  } });
  [
    'managed_map_AAAAA', 'managed_receipt_A', 'managed_duplicate',
    'managed_failed_AAA', 'managed_audio_AAAA', 'unknown_candidate_A'
  ].forEach((id) => harness.addFile(id, `${id}.jpg`, 'image/jpeg', [1, 2, 3], [harness.rootId]));

  const ids = seedExpectedStructure(harness);

  assert.deepEqual(harness.directFileNames(ids.photos), [
    'managed_map_AAAAA.jpg', 'managed_receipt_A.jpg'
  ]);
  assert.deepEqual(harness.directFileNames(harness.rootId).sort(), [
    GUIDE_NAME,
    'managed_audio_AAAA.jpg',
    'managed_duplicate.jpg',
    'managed_failed_AAA.jpg',
    'unknown_candidate_A.jpg'
  ].sort());
});
