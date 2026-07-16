const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const { createHarness } = require('./drive-photo-import-server-harness');

const codeJs = fs.readFileSync(path.resolve(__dirname, '..', 'Code.js'), 'utf8');

function functionBody(name) {
  const start = codeJs.indexOf(`function ${name}(`);
  assert.notEqual(start, -1);
  const open = codeJs.indexOf('{', start);
  let depth = 0;
  for (let index = open; index < codeJs.length; index += 1) {
    if (codeJs[index] === '{') depth += 1;
    if (codeJs[index] === '}') depth -= 1;
    if (depth === 0) return codeJs.slice(open + 1, index);
  }
  assert.fail(`Could not read ${name}`);
}

function seedTree(harness) {
  harness.addFolder('folder_AAAAAAAAAA', 'folder-a', [harness.rootId]);
  harness.addFolder('empty_AAAAAAAAAAA', 'Empty-folder', [harness.rootId]);
  harness.addFolder('trashed_folder_AAA', 'Trashed-folder', [harness.rootId], { trashed: true });
  harness.addFolder('nested_AAAAAAAAAA', 'Nested', ['folder_AAAAAAAAAA']);
  harness.addFolder('outside_AAAAAAAAA', 'Outside');
  harness.addFile('photo_AAAAAAAAAAA', 'z-photo.jpg', 'image/jpeg', [1, 2, 3], ['folder_AAAAAAAAAA']);
  harness.addFile('photo_BBBBBBBBBBB', 'A-photo.heic', 'image/heic', [4, 5], ['folder_AAAAAAAAAA']);
  harness.addFile('unsupported_AAAAA', 'notes.txt', 'text/plain', [6], ['folder_AAAAAAAAAA']);
  harness.addFile('root_photo_AAAAAA', 'root.webp', 'image/webp', [7, 8], [harness.rootId]);
  harness.addFile('shortcut_AAAAAAAAA', 'outside shortcut', 'application/vnd.google-apps.shortcut', [9], [harness.rootId]);
  harness.addFile('trashed_file_AAAAA', 'trashed.jpg', 'image/jpeg', [9], [harness.rootId], { trashed: true });
  harness.addFile('outside_photo_AAAA', 'private.jpg', 'image/jpeg', [10], ['outside_AAAAAAAAA']);
}

test('folder listing returns direct safe children, parent, stable sort, and unsupported count', () => {
  const harness = createHarness();
  seedTree(harness);
  assert.equal(typeof harness.api.listDrivePhotoImportFolder, 'function');

  const rootResult = harness.api.listDrivePhotoImportFolder(harness.tokenPayload({ folderId: '' }));
  assert.equal(rootResult.ok, true);
  assert.deepEqual(JSON.parse(JSON.stringify(rootResult.folders)).map((item) => item.name),
    ['Empty-folder', 'folder-a']);
  assert.deepEqual(JSON.parse(JSON.stringify(rootResult.photos)).map((item) => item.name), ['root.webp']);
  assert.equal(rootResult.parent, null);
  assert.equal(rootResult.folder.isRoot, true);
  assert.deepEqual(JSON.parse(JSON.stringify(rootResult.counts)), { folders: 2, photos: 1 });
  assert.equal(rootResult.ignoredUnsupportedFileCount, 1);
  assert.deepEqual(Object.keys(rootResult.photos[0]).sort(),
    ['id', 'imported', 'kind', 'mimeType', 'modifiedAt', 'name', 'sizeBytes']);
  assert.equal(rootResult.photos[0].imported, false);

  const childResult = harness.api.listDrivePhotoImportFolder(
    harness.tokenPayload({ folderId: 'folder_AAAAAAAAAA' })
  );
  assert.equal(childResult.ok, true);
  assert.deepEqual(JSON.parse(JSON.stringify(childResult.folders)).map((item) => item.name), ['Nested']);
  assert.deepEqual(JSON.parse(JSON.stringify(childResult.photos)).map((item) => item.name),
    ['A-photo.heic', 'z-photo.jpg']);
  assert.deepEqual(JSON.parse(JSON.stringify(childResult.parent)), { id: harness.rootId, name: 'Root' });
  assert.equal(childResult.ignoredUnsupportedFileCount, 1);
  assert.doesNotMatch(JSON.stringify(childResult), /notes\.txt|outside|owner|permission|description|https?:/i);
  assert.deepEqual(harness.audit.writes, []);
});

test('folder listing excludes every surviving managed file except completed receipt sources', () => {
  const receiptHeaders = [
    'idempotencyKey', 'jobId', 'itemId', 'payloadHash', 'state', 'leaseOwner',
    'leaseUntil', 'pinId', 'targetFolderId', 'tempFileName', 'fileId', 'imageUrl',
    'folderUrl', 'createdAt', 'updatedAt', 'lastErrorCode', 'sourceDriveFileId'
  ];
  const mapHeaders = [
    'タイムスタンプ', 'タイトル', '説明', '緯度', '経度', 'ピンの色',
    'ファイルID', '画像URL', 'ID', '参考URL一覧', '状態', 'タグ',
    'イベント時刻', '更新時刻', 'アイコン'
  ];
  const receipt = (state, pinId, sourceDriveFileId) => {
    const row = Array(receiptHeaders.length).fill('');
    row[receiptHeaders.indexOf('state')] = state;
    row[receiptHeaders.indexOf('pinId')] = pinId;
    row[receiptHeaders.indexOf('sourceDriveFileId')] = sourceDriveFileId;
    return row;
  };
  const mapRow = (pinId, fileId = '') => {
    const row = Array(mapHeaders.length).fill('');
    row[mapHeaders.indexOf('ID')] = pinId;
    row[mapHeaders.indexOf('ファイルID')] = fileId;
    return row;
  };
  const ids = {
    directSource: 'photo_DIRECTAAAAAAA',
    legacySource: 'photo_LEGACYAAAAAA',
    legacyCopy: 'managed_COPYAAAAA',
    localUpload: 'managed_LOCALAAAA',
    singleAdd: 'managed_SINGLEAAA',
    receiptless: 'managed_NORECEIPT',
    deletedPin: 'photo_DELETEDAAAAA',
    unmanaged: 'photo_NEWAAAAAAAAA'
  };
  const harness = createHarness({ sheets: {
    import_receipts: [receiptHeaders,
      receipt('completed', 'pin-direct', ids.directSource),
      receipt('completed', 'pin-legacy', ids.legacySource),
      receipt('completed', 'pin-deleted', ids.deletedPin)],
    map_info: [mapHeaders,
      mapRow('pin-direct', ids.directSource),
      mapRow('pin-legacy', ids.legacyCopy),
      mapRow('pin-local', ids.localUpload),
      mapRow('pin-single', ids.singleAdd),
      mapRow('pin-receiptless', ids.receiptless)]
  } });
  Object.values(ids).forEach((id, index) => {
    harness.addFile(id, `${String(index).padStart(2, '0')}.jpg`, 'image/jpeg', [index], [harness.rootId]);
  });

  const result = harness.api.listDrivePhotoImportFolder(harness.tokenPayload({ folderId: '' }));
  assert.equal(result.ok, true);
  const listed = Object.fromEntries(
    JSON.parse(JSON.stringify(result.photos)).map((photo) => [photo.id, photo.imported])
  );
  assert.deepEqual(listed, {
    [ids.directSource]: true,
    [ids.legacySource]: true,
    [ids.deletedPin]: false,
    [ids.unmanaged]: false
  });
  for (const managedId of [ids.legacyCopy, ids.localUpload, ids.singleAdd, ids.receiptless]) {
    assert.equal(Object.hasOwn(listed, managedId), false, managedId);
  }
  assert.equal(harness.audit.sheetReads.filter((read) => read.name === 'import_receipts').length, 1);
  assert.equal(harness.audit.sheetReads.filter((read) => read.name === 'map_info').length, 1);
  assert.deepEqual(harness.audit.writes, []);
});

test('folder listing reports a safe list failure instead of treating a sheet read error as success', () => {
  const harness = createHarness({
    failSheetRead: true,
    sheets: { import_receipts: [['sourceDriveFileId']], map_info: [['ID']] }
  });
  harness.addFile('photo_AAAAAAAAAAA', 'photo.jpg', 'image/jpeg', [1], [harness.rootId]);
  const result = harness.api.listDrivePhotoImportFolder(harness.tokenPayload({ folderId: '' }));
  assert.equal(result.ok, false);
  assert.equal(result.errorCode, 'DRIVE_IMPORT_FOLDER_READ_FAILED');
  assert.doesNotMatch(JSON.stringify(result), /private|spreadsheet|detail/);
  assert.deepEqual(harness.audit.writes, []);
});

test('folder listing accepts 500 entries and atomically rejects 501', () => {
  const accepted = createHarness();
  for (let index = 0; index < 500; index += 1) {
    accepted.addFile(`unsupported_${String(index).padStart(4, '0')}`, `item-${index}.txt`, 'text/plain', [1], [accepted.rootId]);
  }
  const acceptedResult = accepted.api.listDrivePhotoImportFolder(accepted.tokenPayload({ folderId: '' }));
  assert.equal(acceptedResult.ok, true);
  assert.equal(acceptedResult.ignoredUnsupportedFileCount, 500);

  const rejected = createHarness();
  for (let index = 0; index < 501; index += 1) {
    rejected.addFile(`unsupported_${String(index).padStart(4, '0')}`, `item-${index}.txt`, 'text/plain', [1], [rejected.rootId]);
  }
  const rejectedResult = rejected.api.listDrivePhotoImportFolder(rejected.tokenPayload({ folderId: '' }));
  assert.equal(rejectedResult.ok, false);
  assert.equal(rejectedResult.errorCode, 'DRIVE_IMPORT_FOLDER_TOO_LARGE');
  assert.equal(Object.prototype.hasOwnProperty.call(rejectedResult, 'folders'), false);
  assert.doesNotMatch(JSON.stringify(rejectedResult), /item-|unsupported_/);
});

test('folder listing counts trashed entries toward the atomic traversal limit', () => {
  const accepted = createHarness();
  for (let index = 0; index < 500; index += 1) {
    accepted.addFile(`trashed_${String(index).padStart(12, '0')}`, `item-${index}.jpg`,
      'image/jpeg', [1], [accepted.rootId], { trashed: true });
  }
  const acceptedResult = accepted.api.listDrivePhotoImportFolder(accepted.tokenPayload({ folderId: '' }));
  assert.equal(acceptedResult.ok, true);
  assert.equal(acceptedResult.photos.length, 0);

  const rejected = createHarness();
  for (let index = 0; index < 501; index += 1) {
    rejected.addFile(`trashed_${String(index).padStart(12, '0')}`, `item-${index}.jpg`,
      'image/jpeg', [1], [rejected.rootId], { trashed: true });
  }
  const rejectedResult = rejected.api.listDrivePhotoImportFolder(rejected.tokenPayload({ folderId: '' }));
  assert.equal(rejectedResult.ok, false);
  assert.equal(rejectedResult.errorCode, 'DRIVE_IMPORT_FOLDER_TOO_LARGE');
  assert.equal(Object.hasOwn(rejectedResult, 'photos'), false);
});

test('folder listing uses case-insensitive name then ID order and the shared photo type contract', () => {
  const harness = createHarness();
  harness.addFolder('folder_BBBBBBBBBB', 'Same', [harness.rootId]);
  harness.addFolder('folder_AAAAAAAAAA', 'same', [harness.rootId]);
  harness.addFile('photo_BBBBBBBBBBB', 'Same.jpg', 'image/jpeg', [1], [harness.rootId]);
  harness.addFile('photo_AAAAAAAAAAA', 'same.JPG', 'image/jpg', [2], [harness.rootId]);
  harness.addFile('photo_PNGAAAAAAAA', 'photo.png', 'image/png', [3], [harness.rootId]);
  harness.addFile('photo_WEBPAAAAAAA', 'photo.webp', 'application/octet-stream', [4], [harness.rootId]);
  harness.addFile('photo_JPGAAAAAAAA', 'photo.jpg', 'application/octet-stream', [4], [harness.rootId]);
  harness.addFile('photo_HEICAAAAAAA', 'photo.heic', 'image/heic', [5], [harness.rootId]);
  harness.addFile('photo_HEIFAAAAAAA', 'photo.heif', 'binary/octet-stream', [6], [harness.rootId]);
  harness.addFile('photo_MISMATCHAAA', 'wrong.png', 'image/jpeg', [7], [harness.rootId]);

  const result = harness.api.listDrivePhotoImportFolder(harness.tokenPayload({ folderId: '' }));
  assert.equal(result.ok, true);
  assert.deepEqual(JSON.parse(JSON.stringify(result.folders)).map((item) => item.id),
    ['folder_AAAAAAAAAA', 'folder_BBBBBBBBBB']);
  assert.deepEqual(JSON.parse(JSON.stringify(result.photos)).map((item) => [item.id, item.kind]), [
    ['photo_HEICAAAAAAA', 'heic'],
    ['photo_HEIFAAAAAAA', 'heic'],
    ['photo_JPGAAAAAAAA', 'jpeg'],
    ['photo_PNGAAAAAAAA', 'png'],
    ['photo_WEBPAAAAAAA', 'webp'],
    ['photo_AAAAAAAAAAA', 'jpeg'],
    ['photo_BBBBBBBBBBB', 'jpeg']
  ]);
  assert.equal(result.photos.find((item) => item.id === 'photo_JPGAAAAAAAA').mimeType, 'image/jpeg');
  assert.equal(result.ignoredUnsupportedFileCount, 1);
});

test('photo read validates containment/type/trash and returns exact base64 with one blob and byte read', () => {
  const harness = createHarness();
  seedTree(harness);
  assert.equal(typeof harness.api.readDrivePhotoImportFile, 'function');

  const result = harness.api.readDrivePhotoImportFile(
    harness.tokenPayload({ fileId: 'photo_AAAAAAAAAAA' })
  );
  assert.equal(result.ok, true);
  assert.equal(result.file.base64, 'AQID');
  assert.equal(result.file.sizeBytes, 3);
  assert.equal(result.file.kind, 'jpeg');
  assert.equal(result.file.modifiedAt, '2026-07-12T01:02:03.000Z');
  assert.equal(harness.audit.blobReads, 1);
  assert.equal(harness.audit.byteReads, 1);
  assert.deepEqual(harness.audit.writes, []);

  const rootDirect = harness.api.readDrivePhotoImportFile(
    harness.tokenPayload({ fileId: 'root_photo_AAAAAA' })
  );
  assert.equal(rootDirect.ok, true);
  assert.equal(rootDirect.file.kind, 'webp');

  const outside = harness.api.readDrivePhotoImportFile(
    harness.tokenPayload({ fileId: 'outside_photo_AAAA' })
  );
  assert.equal(outside.errorCode, 'DRIVE_IMPORT_FILE_OUTSIDE_ROOT');
  const unsupported = harness.api.readDrivePhotoImportFile(
    harness.tokenPayload({ fileId: 'unsupported_AAAAA' })
  );
  assert.equal(unsupported.errorCode, 'DRIVE_IMPORT_FILE_TYPE_UNSUPPORTED');
  const shortcut = harness.api.readDrivePhotoImportFile(
    harness.tokenPayload({ fileId: 'shortcut_AAAAAAAAA' })
  );
  assert.equal(shortcut.errorCode, 'DRIVE_IMPORT_FILE_TYPE_UNSUPPORTED');
  const trashed = harness.api.readDrivePhotoImportFile(
    harness.tokenPayload({ fileId: 'trashed_file_AAAAA' })
  );
  assert.equal(trashed.errorCode, 'DRIVE_IMPORT_FILE_NOT_FOUND');
  assert.doesNotMatch(JSON.stringify([outside, unsupported, shortcut, trashed]),
    /private|outside_photo|unsupported_A|shortcut_A|trashed_file/);
});

test('photo read enforces the actual-byte 15MB boundary and sanitizes read failures', () => {
  const exact = createHarness();
  exact.addFile('exact_file_AAAAAAA', 'exact.jpg', 'image/jpeg', new Uint8Array(15 * 1024 * 1024), [exact.rootId]);
  const exactResult = exact.api.readDrivePhotoImportFile(exact.tokenPayload({ fileId: 'exact_file_AAAAAAA' }));
  assert.equal(exactResult.ok, true);
  assert.equal(exactResult.file.sizeBytes, 15 * 1024 * 1024);

  const over = createHarness();
  over.addFile('large_file_AAAAAAA', 'large.jpg', 'image/jpeg',
    new Uint8Array(15 * 1024 * 1024 + 1), [over.rootId], { metadataSize: 1 });
  const overResult = over.api.readDrivePhotoImportFile(over.tokenPayload({ fileId: 'large_file_AAAAAAA' }));
  assert.equal(overResult.errorCode, 'DRIVE_IMPORT_FILE_TOO_LARGE');
  assert.equal(over.audit.blobReads, 1);
  assert.equal(over.audit.byteReads, 1);

  const failed = createHarness();
  failed.addFile('failed_file_AAAAAA', 'failed.jpg', 'image/jpeg', [1], [failed.rootId], { blobError: true });
  const failedResult = failed.api.readDrivePhotoImportFile(failed.tokenPayload({ fileId: 'failed_file_AAAAAA' }));
  assert.equal(failedResult.errorCode, 'DRIVE_IMPORT_FILE_READ_FAILED');
  assert.doesNotMatch(JSON.stringify(failedResult), /private|failed_file|failed\.jpg/);

  const mismatch = createHarness();
  mismatch.addFile('mismatch_file_AAAA', 'mismatch.jpg', 'image/jpeg', [1, 2], [mismatch.rootId], {
    metadataSize: 1
  });
  const mismatchResult = mismatch.api.readDrivePhotoImportFile(
    mismatch.tokenPayload({ fileId: 'mismatch_file_AAAA' })
  );
  assert.equal(mismatchResult.ok, true);
  assert.equal(mismatchResult.file.sizeBytes, 2);
  assert.equal(mismatchResult.file.base64, 'AQI=');
  assert.equal(mismatch.audit.blobReads, 1);
  assert.equal(mismatch.audit.byteReads, 1);
});

test('folder listing hides the exact root original folder and refuses direct browsing of its photos', () => {
  const harness = createHarness();
  harness.addFolder('original_AAAAAAAAA', 'original', [harness.rootId]);
  harness.addFolder('Original_AAAAAAAAA', 'Original', [harness.rootId]);
  harness.addFile('archived_AAAAAAAAA', 'archived.jpg', 'image/jpeg', [1, 2, 3], ['original_AAAAAAAAA']);

  const rootResult = harness.api.listDrivePhotoImportFolder(harness.tokenPayload({ folderId: '' }));
  assert.equal(rootResult.ok, true);
  assert.deepEqual(JSON.parse(JSON.stringify(rootResult.folders)).map((folder) => folder.name), ['Original']);
  assert.equal(rootResult.photos.some((photo) => photo.name === 'archived.jpg'), false);

  const archivedResult = harness.api.listDrivePhotoImportFolder(
    harness.tokenPayload({ folderId: 'original_AAAAAAAAA' })
  );
  assert.equal(archivedResult.ok, false);
  assert.equal(archivedResult.errorCode, 'DRIVE_IMPORT_FOLDER_NOT_FOUND');
  assert.doesNotMatch(JSON.stringify(archivedResult), /archived\.jpg|archived_A/);
});

test('server never emits an unsafe basename or invalid modifiedAt descriptor', () => {
  const harness = createHarness();
  harness.addFile('unsafe_name_AAAAAAA', '../photo.jpg', 'image/jpeg', [1], [harness.rootId]);
  harness.addFile('unsafe_c1_AAAAAAAAA', 'bad\u0085name.jpg', 'image/jpeg', [1], [harness.rootId]);
  harness.addFile('unsafe_bidi_AAAAAAA', 'bad\u202Ename.jpg', 'image/jpeg', [1], [harness.rootId]);
  harness.addFile('invalid_date_AAAAAA', 'photo.jpg', 'image/jpeg', [1], [harness.rootId], {
    modifiedAt: 'not-a-date'
  });
  const listing = harness.api.listDrivePhotoImportFolder(harness.tokenPayload({ folderId: '' }));
  assert.equal(listing.ok, true);
  assert.deepEqual(JSON.parse(JSON.stringify(listing.photos)), []);
  assert.equal(listing.ignoredUnsupportedFileCount, 4);
  assert.equal(harness.api.readDrivePhotoImportFile(
    harness.tokenPayload({ fileId: 'unsafe_name_AAAAAAA' })
  ).errorCode, 'DRIVE_IMPORT_FILE_READ_FAILED');
  assert.equal(harness.api.readDrivePhotoImportFile(
    harness.tokenPayload({ fileId: 'invalid_date_AAAAAA' })
  ).errorCode, 'DRIVE_IMPORT_FILE_READ_FAILED');
  assert.equal(harness.api.readDrivePhotoImportFile(
    harness.tokenPayload({ fileId: 'unsafe_c1_AAAAAAAAA' })
  ).errorCode, 'DRIVE_IMPORT_FILE_READ_FAILED');
  assert.equal(harness.api.readDrivePhotoImportFile(
    harness.tokenPayload({ fileId: 'unsafe_bidi_AAAAAAA' })
  ).errorCode, 'DRIVE_IMPORT_FILE_READ_FAILED');
});

test('folder listing accepts displayable path-like names and rejects dangerous folder controls safely', () => {
  const safe = createHarness();
  safe.addFolder('safe_name_AAAAAAAA', '<b>A/B\\C</b>', [safe.rootId]);
  const safeResult = safe.api.listDrivePhotoImportFolder(safe.tokenPayload({ folderId: '' }));
  assert.equal(safeResult.ok, true);
  assert.equal(safeResult.folders[0].name, '<b>A/B\\C</b>');

  ['bad\nname', 'bad\u0085name', 'bad\u202Ename'].forEach((name, index) => {
    const hostile = createHarness();
    hostile.addFolder(`hostile_${String(index).padStart(12, '0')}`, name, [hostile.rootId]);
    const result = hostile.api.listDrivePhotoImportFolder(hostile.tokenPayload({ folderId: '' }));
    assert.equal(result.ok, false);
    assert.equal(result.errorCode, 'DRIVE_IMPORT_FOLDER_NOT_FOUND');
    assert.doesNotMatch(JSON.stringify(result), /bad|name/);
  });
});

test('folder failures ignore unknown or throwing provider codes and keep a safe folder fallback', () => {
  ['unknown', 'throwing'].forEach((mode) => {
    const harness = createHarness();
    harness.folders.get(harness.rootId).getFolders = function() {
      const error = new Error('private provider detail ' + harness.rootId);
      if (mode === 'unknown') error.code = 'EIO';
      else Object.defineProperty(error, 'code', {
        get() { throw new Error('private code getter ' + harness.rootId); }
      });
      throw error;
    };
    const result = harness.api.listDrivePhotoImportFolder(harness.tokenPayload({ folderId: '' }));
    assert.equal(result.ok, false);
    assert.equal(result.errorCode, 'DRIVE_IMPORT_FOLDER_NOT_FOUND');
    assert.doesNotMatch(JSON.stringify(result), /private|EIO|root_/);
  });
});

test('non-root listing fails safely when its validated parent cannot be projected', () => {
  const harness = createHarness();
  const child = harness.addFolder('child_AAAAAAAAAAA', 'Child', [harness.rootId]);
  const originalGetParents = child.getParents;
  let parentReads = 0;
  child.getParents = function() {
    parentReads += 1;
    if (parentReads > 1) throw new Error('private parent changed after containment');
    return originalGetParents();
  };
  const result = harness.api.listDrivePhotoImportFolder(
    harness.tokenPayload({ folderId: child.getId() })
  );
  assert.equal(result.ok, false);
  assert.equal(result.errorCode, 'DRIVE_IMPORT_FOLDER_NOT_FOUND');
  assert.doesNotMatch(JSON.stringify(result), /private|parent changed|child_A/);
});

test('Phase 9-1 server APIs contain no Drive or Sheet write path', () => {
  const forbidden = [
    'createFile', 'createFolder', 'makeCopy', 'moveTo', 'setName', 'setSharing',
    'setTrashed', 'appendRow', 'setValues', 'saveImportPhotoItem', 'saveMapData'
  ];
  ['listDrivePhotoImportFolder', 'readDrivePhotoImportFile'].forEach((name) => {
    const body = functionBody(name);
    forbidden.forEach((operation) => {
      assert.equal(body.includes(operation + '('), false, `${name} must not call ${operation}`);
    });
  });
});
