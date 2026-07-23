const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const { createHarness } = require('./drive-photo-import-server-harness');

const codeJs = fs.readFileSync(path.join(__dirname, '..', 'Code.js'), 'utf8');

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

function seedStructure(harness) {
  const result = harness.api.ensureMediaDriveStructure(harness.tokenPayload({}));
  assert.deepEqual(plain(result), { ok: true });
  harness.audit.writes.length = 0;
  return {
    photos: harness.folderId('photos'),
    audio: harness.folderId('audio'),
    original: harness.folderId('original'),
    originalPhotos: harness.folderId('photos', harness.folderId('original')),
    originalAudio: harness.folderId('audio', harness.folderId('original'))
  };
}

test('legacy photo listing delegates to a root-only inbox without exposing folder IDs', () => {
  const harness = createHarness();
  const ids = seedStructure(harness);
  harness.addFolder('nested_AAAAAAAAAA', 'Nested', [harness.rootId]);
  harness.addFile('root_jpeg_AAAAAAA', 'A-photo.jpg', 'image/jpeg', [1], [harness.rootId]);
  harness.addFile('root_png_AAAAAAAA', 'b-photo.png', 'image/png', [2], [harness.rootId]);
  harness.addFile('root_webp_AAAAAAA', 'c-photo.webp', 'image/webp', [3], [harness.rootId]);
  harness.addFile('root_heic_AAAAAAA', 'd-photo.heic', 'image/heic', [4], [harness.rootId]);
  harness.addFile('root_heif_AAAAAAA', 'e-photo.heif', 'image/heif', [5], [harness.rootId]);
  harness.addFile('nested_photo_AAAA', 'nested.jpg', 'image/jpeg', [6], ['nested_AAAAAAAAAA']);
  harness.addFile('managed_photo_AAA', 'managed.jpg', 'image/jpeg', [7], [ids.photos]);
  harness.addFile('archived_photo_AA', 'archived.jpg', 'image/jpeg', [8], [ids.originalPhotos]);

  const result = harness.api.listDrivePhotoImportFolder(harness.tokenPayload({ folderId: '' }));

  assert.equal(result.ok, true);
  assert.deepEqual(plain(result.folder), { id: '', name: '取込Inbox', isRoot: true });
  assert.equal(result.parent, null);
  assert.deepEqual(plain(result.folders), []);
  assert.deepEqual(plain(result.photos).map((item) => [item.name, item.kind]), [
    ['A-photo.jpg', 'jpeg'], ['b-photo.png', 'png'], ['c-photo.webp', 'webp'],
    ['d-photo.heic', 'heic'], ['e-photo.heif', 'heic']
  ]);
  assert.deepEqual(plain(result.counts), { folders: 0, photos: 5 });
  assert.equal(JSON.stringify(result).includes(harness.rootId), false);
  assert.deepEqual(harness.audit.writes, []);
});

test('legacy photo listing rejects every non-root folder ID with one stable boundary code', () => {
  const harness = createHarness();
  seedStructure(harness);
  const nested = harness.addFolder('nested_AAAAAAAAAA', 'Nested', [harness.rootId]);

  const result = harness.api.listDrivePhotoImportFolder(
    harness.tokenPayload({ folderId: nested.getId() })
  );

  assert.equal(result.ok, false);
  assert.equal(result.errorCode, 'DRIVE_IMPORT_FOLDER_OUTSIDE_ROOT');
  assert.doesNotMatch(JSON.stringify(result), /nested_A|Nested/);
  assert.deepEqual(harness.audit.writes, []);
});

test('photo inbox excludes managed IDs and completed source IDs while leaving unknown root candidates', () => {
  const receiptHeaders = [
    'idempotencyKey', 'jobId', 'itemId', 'payloadHash', 'state', 'leaseOwner',
    'leaseUntil', 'pinId', 'targetFolderId', 'tempFileName', 'fileId', 'imageUrl',
    'folderUrl', 'createdAt', 'updatedAt', 'lastErrorCode', 'sourceDriveFileId',
    'mediaKind', 'operationMode', 'targetPinId', 'cleanupFileId'
  ];
  const mapHeaders = ['ID', 'ファイルID', '音声ID'];
  const receipt = Array(receiptHeaders.length).fill('');
  receipt[receiptHeaders.indexOf('state')] = 'completed';
  receipt[receiptHeaders.indexOf('pinId')] = 'pin-source';
  receipt[receiptHeaders.indexOf('sourceDriveFileId')] = 'source_completed_A';
  receipt[receiptHeaders.indexOf('mediaKind')] = 'photo';
  const harness = createHarness({ sheets: {
    map_info: [mapHeaders, ['pin-managed', 'managed_photo_AAA', '']],
    import_receipts: [receiptHeaders, receipt]
  } });
  seedStructure(harness);
  harness.addFile('managed_photo_AAA', 'managed.jpg', 'image/jpeg', [1], [harness.rootId]);
  harness.addFile('source_completed_A', 'completed.jpg', 'image/jpeg', [2], [harness.rootId]);
  harness.addFile('unknown_photo_AAA', 'unknown.jpg', 'image/jpeg', [3], [harness.rootId]);

  const result = harness.api.listDriveMediaInbox(harness.tokenPayload({ mediaKind: 'photo' }));

  assert.deepEqual(plain(result.items).map((item) => item.id), ['unknown_photo_AAA']);
  assert.deepEqual(Object.keys(plain(result.items[0])).sort(),
    ['id', 'kind', 'mimeType', 'modifiedAt', 'name', 'sizeBytes']);
});

test('photo list preserves strict type, safe-name, trash, shortcut, size, and 500-entry behavior', () => {
  const harness = createHarness();
  seedStructure(harness);
  harness.addFile('generic_jpg_AAAAAA', 'generic.jpg', 'application/octet-stream', [1], [harness.rootId]);
  harness.addFile('mismatch_AAAAAAAAA', 'wrong.png', 'image/jpeg', [1], [harness.rootId]);
  harness.addFile('unsafe_AAAAAAAAAAA', '../unsafe.jpg', 'image/jpeg', [1], [harness.rootId]);
  harness.addFile('zero_AAAAAAAAAAAAA', 'zero.jpg', 'image/jpeg', [], [harness.rootId]);
  harness.addFile('shortcut_AAAAAAAAA', 'shortcut.jpg', 'application/vnd.google-apps.shortcut', [1], [harness.rootId]);
  harness.addFile('trash_AAAAAAAAAAAA', 'trash.jpg', 'image/jpeg', [1], [harness.rootId], { trashed: true });
  const result = harness.api.listDriveMediaInbox(harness.tokenPayload({ mediaKind: 'photo' }));
  assert.deepEqual(plain(result.items).map((item) => item.id), ['generic_jpg_AAAAAA']);

  const accepted = createHarness();
  seedStructure(accepted);
  for (let index = 0; index < 500; index += 1) {
    accepted.addFile(`unsupported_${String(index).padStart(4, '0')}`, `item-${index}.txt`,
      'text/plain', [1], [accepted.rootId]);
  }
  assert.equal(accepted.api.listDriveMediaInbox(
    accepted.tokenPayload({ mediaKind: 'photo' })
  ).ok, true);

  const rejected = createHarness();
  seedStructure(rejected);
  for (let index = 0; index < 501; index += 1) {
    rejected.addFile(`unsupported_${String(index).padStart(4, '0')}`, `item-${index}.txt`,
      'text/plain', [1], [rejected.rootId]);
  }
  assert.equal(rejected.api.listDriveMediaInbox(
    rejected.tokenPayload({ mediaKind: 'photo' })
  ).errorCode, 'DRIVE_MEDIA_INBOX_TOO_LARGE');
});

test('legacy photo read is root-direct-only and preserves exact actual-byte materialization', () => {
  const harness = createHarness();
  const ids = seedStructure(harness);
  harness.addFile('root_photo_AAAAAAA', 'root.webp', 'image/webp', [1, 2, 3], [harness.rootId], {
    metadataSize: 1
  });
  harness.addFile('nested_photo_AAAA', 'nested.jpg', 'image/jpeg', [1], [ids.photos]);
  harness.addFile('trashed_photo_AAA', 'trash.jpg', 'image/jpeg', [1], [harness.rootId], { trashed: true });
  harness.addFile('unsupported_AAAAA', 'notes.txt', 'text/plain', [1], [harness.rootId]);

  const result = harness.api.readDrivePhotoImportFile(
    harness.tokenPayload({ fileId: 'root_photo_AAAAAAA' })
  );

  assert.equal(result.ok, true);
  assert.equal(result.file.base64, 'AQID');
  assert.equal(result.file.sizeBytes, 3);
  assert.equal(result.file.kind, 'webp');
  assert.equal(harness.api.readDrivePhotoImportFile(
    harness.tokenPayload({ fileId: 'nested_photo_AAAA' })
  ).errorCode, 'DRIVE_IMPORT_FILE_OUTSIDE_ROOT');
  assert.equal(harness.api.readDrivePhotoImportFile(
    harness.tokenPayload({ fileId: 'trashed_photo_AAA' })
  ).errorCode, 'DRIVE_IMPORT_FILE_NOT_FOUND');
  assert.equal(harness.api.readDrivePhotoImportFile(
    harness.tokenPayload({ fileId: 'unsupported_AAAAA' })
  ).errorCode, 'DRIVE_IMPORT_FILE_TYPE_UNSUPPORTED');
  assert.deepEqual(harness.audit.writes, []);
});

test('photo read rejects completed root sources and proven managed root files just like listing', () => {
  const receiptHeaders = [
    'idempotencyKey', 'jobId', 'itemId', 'payloadHash', 'state', 'leaseOwner',
    'leaseUntil', 'pinId', 'targetFolderId', 'tempFileName', 'fileId', 'imageUrl',
    'folderUrl', 'createdAt', 'updatedAt', 'lastErrorCode', 'sourceDriveFileId',
    'mediaKind', 'operationMode', 'targetPinId', 'cleanupFileId'
  ];
  const completed = Array(receiptHeaders.length).fill('');
  completed[receiptHeaders.indexOf('state')] = 'completed';
  completed[receiptHeaders.indexOf('pinId')] = 'pin-completed-photo';
  completed[receiptHeaders.indexOf('sourceDriveFileId')] = 'completed_photo_A';
  completed[receiptHeaders.indexOf('mediaKind')] = 'photo';
  const harness = createHarness({ sheets: {
    map_info: [['ID', 'ファイルID', '音声ID'], ['pin-managed-photo', 'managed_photo_AAA', '']],
    import_receipts: [receiptHeaders, completed]
  } });
  seedStructure(harness);
  harness.addFile('completed_photo_A', 'completed.jpg', 'image/jpeg', [1], [harness.rootId]);
  harness.addFile('managed_photo_AAA', 'managed.jpg', 'image/jpeg', [2], [harness.rootId]);

  assert.equal(harness.api.readDrivePhotoImportFile(
    harness.tokenPayload({ fileId: 'completed_photo_A' })
  ).errorCode, 'DRIVE_IMPORT_FILE_OUTSIDE_ROOT');
  assert.equal(harness.api.readDrivePhotoImportFile(
    harness.tokenPayload({ fileId: 'managed_photo_AAA' })
  ).errorCode, 'DRIVE_IMPORT_FILE_OUTSIDE_ROOT');
  assert.equal(harness.audit.blobReads, 0);
});

test('photo read enforces the actual-byte 15MiB limit and revalidates replacement state', () => {
  const exact = createHarness();
  seedStructure(exact);
  exact.addFile('exact_photo_AAAAA', 'exact.jpg', 'image/jpeg', new Uint8Array(15 * 1024 * 1024),
    [exact.rootId], { metadataSize: 1 });
  assert.equal(exact.api.readDrivePhotoImportFile(
    exact.tokenPayload({ fileId: 'exact_photo_AAAAA' })
  ).file.sizeBytes, 15 * 1024 * 1024);

  const over = createHarness();
  seedStructure(over);
  over.addFile('large_photo_AAAAA', 'large.jpg', 'image/jpeg', new Uint8Array(15 * 1024 * 1024 + 1),
    [over.rootId], { metadataSize: 1 });
  assert.equal(over.api.readDrivePhotoImportFile(
    over.tokenPayload({ fileId: 'large_photo_AAAAA' })
  ).errorCode, 'DRIVE_IMPORT_FILE_TOO_LARGE');

  const replaced = createHarness();
  const ids = seedStructure(replaced);
  replaced.addFile('replace_photo_AAA', 'replace.jpg', 'image/jpeg', [1], [replaced.rootId]);
  assert.equal(replaced.api.listDriveMediaInbox(
    replaced.tokenPayload({ mediaKind: 'photo' })
  ).items.length, 1);
  replaced.replaceFile('replace_photo_AAA', 'replace.jpg', 'image/jpeg', [1], [ids.photos]);
  assert.equal(replaced.api.readDrivePhotoImportFile(
    replaced.tokenPayload({ fileId: 'replace_photo_AAA' })
  ).errorCode, 'DRIVE_IMPORT_FILE_OUTSIDE_ROOT');
});

test('compatibility wrappers delegate to media APIs instead of retaining descendant traversal', () => {
  const listBody = codeJs.slice(
    codeJs.indexOf('function listDrivePhotoImportFolder'),
    codeJs.indexOf('function readDrivePhotoImportFile')
  );
  const readBody = codeJs.slice(
    codeJs.indexOf('function readDrivePhotoImportFile'),
    codeJs.indexOf('\nfunction ', codeJs.indexOf('function readDrivePhotoImportFile') + 10)
  );
  assert.match(listBody, /listDriveMediaInbox/);
  assert.match(readBody, /readDriveMediaImportFile_/);
  assert.doesNotMatch(listBody, /getFolders\(/);
});
