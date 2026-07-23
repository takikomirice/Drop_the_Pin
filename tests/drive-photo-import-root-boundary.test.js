const assert = require('node:assert/strict');
const test = require('node:test');
const { createHarness } = require('./drive-photo-import-server-harness');

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

test('root containment helper still rejects outside, cycles, trash, and parent failures', () => {
  const harness = createHarness();
  const folderA = harness.addFolder('folder_AAAAAAAAAA', 'Folder A', [harness.rootId]);
  const nested = harness.addFolder('folder_BBBBBBBBBB', 'Nested', [folderA.getId()]);
  const outside = harness.addFolder('outside_AAAAAAAAA', 'Outside');
  const cycleA = harness.addFolder('cycle_AAAAAAAAAAA', 'Cycle A', ['cycle_BBBBBBBBBBB']);
  harness.addFolder('cycle_BBBBBBBBBBB', 'Cycle B', [cycleA.getId()]);
  const broken = harness.addFolder('broken_AAAAAAAAAA', 'Broken', [harness.rootId], { parentError: true });
  const insideFile = harness.addFile('file_inside_AAAAAA', 'inside.jpg', 'image/jpeg', [1], [nested.getId()]);
  const outsideFile = harness.addFile('file_outside_AAAAA', 'outside.jpg', 'image/jpeg', [1], [outside.getId()]);

  assert.equal(harness.api.isDriveFolderWithinRoot_(harness.folders.get(harness.rootId), harness.rootId), true);
  assert.equal(harness.api.isDriveFolderWithinRoot_(nested, harness.rootId), true);
  assert.equal(harness.api.isDriveFolderWithinRoot_(outside, harness.rootId), false);
  assert.equal(harness.api.isDriveFolderWithinRoot_(cycleA, harness.rootId), false);
  assert.equal(harness.api.isDriveFolderWithinRoot_(broken, harness.rootId), false);
  assert.equal(harness.api.isDriveFileWithinRoot_(insideFile, harness.rootId), true);
  assert.equal(harness.api.isDriveFileWithinRoot_(outsideFile, harness.rootId), false);
});

test('root-only photo wrapper validates IDs before Drive access and rejects every non-root value', () => {
  const harness = createHarness();
  harness.api.ensureMediaDriveStructure(harness.tokenPayload({}));
  harness.audit.writes.length = 0;
  const child = harness.addFolder('child_AAAAAAAAAAA', 'Child', [harness.rootId]);
  const invalidValues = [
    'short', 'with space AAAAA', 'https://drive.google.com/x', '../folder_AAAAA',
    'folder/AAAAAAAAAA', 'folder\\AAAAAAAAAA', '__proto__', 'constructor', 'prototype',
    'x'.repeat(201), 'line\nbreakAAAAA'
  ];
  invalidValues.forEach((folderId) => {
    const before = harness.audit.folderReads.length;
    const result = harness.api.listDrivePhotoImportFolder(harness.tokenPayload({ folderId }));
    assert.equal(result.errorCode, 'DRIVE_IMPORT_FOLDER_ID_INVALID');
    assert.equal(harness.audit.folderReads.length, before);
  });
  const childResult = harness.api.listDrivePhotoImportFolder(
    harness.tokenPayload({ folderId: child.getId() })
  );
  assert.equal(childResult.errorCode, 'DRIVE_IMPORT_FOLDER_OUTSIDE_ROOT');
  assert.doesNotMatch(JSON.stringify(childResult), /child_A|Child/);
  assert.deepEqual(harness.audit.writes, []);
});

test('media APIs reject inherited or invalid properties before Drive access', () => {
  const harness = createHarness();
  const inheritedKind = Object.create({ mediaKind: 'audio' });
  inheritedKind.__editToken = 'valid-token';
  assert.equal(harness.api.listDriveMediaInbox(inheritedKind).errorCode, 'DRIVE_MEDIA_KIND_INVALID');

  const inheritedFile = Object.create({ fileId: 'audio_AAAAAAAAAAA' });
  inheritedFile.__editToken = 'valid-token';
  assert.equal(harness.api.readDriveAudioImportFile(inheritedFile).errorCode,
    'DRIVE_AUDIO_FILE_ID_INVALID');

  const readsBefore = harness.audit.fileReads.length;
  assert.equal(harness.api.readDriveAudioImportFile(
    harness.tokenPayload({ fileId: '../audio_AAAAAAA' })
  ).errorCode, 'DRIVE_AUDIO_FILE_ID_INVALID');
  assert.equal(harness.audit.fileReads.length, readsBefore);
  assert.deepEqual(harness.audit.writes, []);
});

test('all media entry points reject invalid token and missing root without Drive access or writes', () => {
  const denied = createHarness();
  const deniedResults = [
    denied.api.ensureMediaDriveStructure({ __editToken: 'bad' }),
    denied.api.listDriveMediaInbox({ __editToken: 'bad', mediaKind: 'photo' }),
    denied.api.readDriveAudioImportFile({ __editToken: 'bad', fileId: 'audio_AAAAAAAAAAA' }),
    denied.api.listDrivePhotoImportFolder({ __editToken: 'bad', folderId: '' }),
    denied.api.readDrivePhotoImportFile({ __editToken: 'bad', fileId: 'photo_AAAAAAAAAAA' })
  ];
  assert.deepEqual(plain(deniedResults).map((item) => item.errorCode), [
    'DRIVE_MEDIA_ACCESS_DENIED', 'DRIVE_MEDIA_ACCESS_DENIED', 'DRIVE_MEDIA_ACCESS_DENIED',
    'DRIVE_IMPORT_ACCESS_DENIED', 'DRIVE_IMPORT_ACCESS_DENIED'
  ]);
  assert.deepEqual(denied.audit.folderReads, []);
  assert.deepEqual(denied.audit.fileReads, []);
  assert.deepEqual(denied.audit.writes, []);

  const missing = createHarness({ rootMissing: true });
  assert.equal(missing.api.ensureMediaDriveStructure(missing.tokenPayload({})).errorCode,
    'DRIVE_MEDIA_ROOT_MISSING');
  assert.equal(missing.api.listDriveMediaInbox(
    missing.tokenPayload({ mediaKind: 'photo' })
  ).errorCode, 'DRIVE_MEDIA_ROOT_MISSING');
  assert.equal(missing.api.readDriveAudioImportFile(
    missing.tokenPayload({ fileId: 'audio_AAAAAAAAAAA' })
  ).errorCode, 'DRIVE_MEDIA_ROOT_MISSING');
  assert.deepEqual(missing.audit.folderReads, []);
  assert.deepEqual(missing.audit.fileReads, []);
  assert.deepEqual(missing.audit.writes, []);
});
