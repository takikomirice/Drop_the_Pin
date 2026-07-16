const assert = require('node:assert/strict');
const test = require('node:test');
const { createHarness } = require('./drive-photo-import-server-harness');

test('root containment accepts root and descendants but rejects outside, cycles, and parent failures', () => {
  const harness = createHarness();
  const folderA = harness.addFolder('folder_AAAAAAAAAA', 'Folder A', [harness.rootId]);
  const nested = harness.addFolder('folder_BBBBBBBBBB', 'Nested', ['folder_AAAAAAAAAA']);
  const outside = harness.addFolder('outside_AAAAAAAAA', 'Outside');
  const cycleA = harness.addFolder('cycle_AAAAAAAAAAA', 'Cycle A', ['cycle_BBBBBBBBBBB']);
  harness.addFolder('cycle_BBBBBBBBBBB', 'Cycle B', ['cycle_AAAAAAAAAAA']);
  const broken = harness.addFolder('broken_AAAAAAAAAA', 'Broken', [harness.rootId], { parentError: true });
  const insideFile = harness.addFile('file_inside_AAAAAA', 'inside.jpg', 'image/jpeg', [1], [nested.getId()]);
  const outsideFile = harness.addFile('file_outside_AAAAA', 'outside.jpg', 'image/jpeg', [1], [outside.getId()]);

  assert.equal(typeof harness.api.isDriveFolderWithinRoot_, 'function');
  assert.equal(harness.api.isDriveFolderWithinRoot_(harness.folders.get(harness.rootId), harness.rootId), true);
  assert.equal(harness.api.isDriveFolderWithinRoot_(folderA, harness.rootId), true);
  assert.equal(harness.api.isDriveFolderWithinRoot_(nested, harness.rootId), true);
  assert.equal(harness.api.isDriveFolderWithinRoot_(outside, harness.rootId), false);
  assert.equal(harness.api.isDriveFolderWithinRoot_(cycleA, harness.rootId), false);
  assert.equal(harness.api.isDriveFolderWithinRoot_(broken, harness.rootId), false);
  assert.equal(harness.api.isDriveFileWithinRoot_(insideFile, harness.rootId), true);
  assert.equal(harness.api.isDriveFileWithinRoot_(outsideFile, harness.rootId), false);
});

test('folder API validates own strict IDs before Drive access and sanitizes outside or missing folders', () => {
  const harness = createHarness();
  harness.addFolder('outside_AAAAAAAAA', 'Outside');
  assert.equal(typeof harness.api.listDrivePhotoImportFolder, 'function');
  const invalidValues = [
    'short', 'with space AAAAA', 'https://drive.google.com/x', '../folder_AAAAA',
    'folder/AAAAAAAAAA', 'folder\\AAAAAAAAAA', '__proto__', 'constructor', 'prototype',
    'x'.repeat(201), 'line\nbreakAAAAA'
  ];
  invalidValues.forEach((folderId) => {
    const before = harness.audit.folderReads.length;
    const result = harness.api.listDrivePhotoImportFolder(
      harness.tokenPayload({ folderId })
    );
    assert.equal(result.ok, false);
    assert.equal(result.errorCode, 'DRIVE_IMPORT_FOLDER_ID_INVALID');
    assert.equal(harness.audit.folderReads.length, before);
    assert.equal(JSON.stringify(result).includes(folderId), false);
  });

  const inherited = Object.create({ folderId: 'folder_AAAAAAAAAA' });
  inherited.__editToken = 'valid-token';
  const inheritedResult = harness.api.listDrivePhotoImportFolder(inherited);
  assert.equal(inheritedResult.errorCode, 'DRIVE_IMPORT_FOLDER_ID_INVALID');

  const inheritedFile = Object.create({ fileId: 'file_inside_AAAAAA' });
  inheritedFile.__editToken = 'valid-token';
  const beforeInheritedFile = harness.audit.fileReads.length;
  assert.equal(harness.api.readDrivePhotoImportFile(inheritedFile).errorCode,
    'DRIVE_IMPORT_FILE_ID_INVALID');
  assert.equal(harness.audit.fileReads.length, beforeInheritedFile);

  const outsideResult = harness.api.listDrivePhotoImportFolder(
    harness.tokenPayload({ folderId: 'outside_AAAAAAAAA' })
  );
  assert.equal(outsideResult.errorCode, 'DRIVE_IMPORT_FOLDER_OUTSIDE_ROOT');
  assert.doesNotMatch(JSON.stringify(outsideResult), /outside_AAAAAAAAA|Outside/);

  const missingResult = harness.api.listDrivePhotoImportFolder(
    harness.tokenPayload({ folderId: 'missing_AAAAAAAAA' })
  );
  assert.equal(missingResult.errorCode, 'DRIVE_IMPORT_FOLDER_NOT_FOUND');
  assert.doesNotMatch(JSON.stringify(missingResult), /missing_AAAAAAAAA|private/);
});

test('folder and file APIs reject invalid token and missing root without Drive access', () => {
  const denied = createHarness();
  assert.equal(denied.api.listDrivePhotoImportFolder({ folderId: '', __editToken: 'bad' }).errorCode,
    'DRIVE_IMPORT_ACCESS_DENIED');
  assert.equal(denied.api.readDrivePhotoImportFile({ fileId: 'file_inside_AAAAAA', __editToken: 'bad' }).errorCode,
    'DRIVE_IMPORT_ACCESS_DENIED');
  assert.deepEqual(denied.audit.folderReads, []);
  assert.deepEqual(denied.audit.fileReads, []);
  assert.deepEqual(denied.audit.sheetReads, []);

  const missing = createHarness({ rootMissing: true });
  assert.equal(missing.api.listDrivePhotoImportFolder(missing.tokenPayload({ folderId: '' })).errorCode,
    'DRIVE_IMPORT_ROOT_MISSING');
  assert.equal(missing.api.readDrivePhotoImportFile(
    missing.tokenPayload({ fileId: 'file_inside_AAAAAA' })
  ).errorCode, 'DRIVE_IMPORT_ROOT_MISSING');
  assert.deepEqual(missing.audit.folderReads, []);
  assert.deepEqual(missing.audit.fileReads, []);
  assert.deepEqual(missing.audit.sheetReads, []);
});

test('configured root IDs accept exactly 10 and 200 characters but reject 201', () => {
  ['r'.repeat(10), 'r'.repeat(200)].forEach((rootId) => {
    const harness = createHarness({ rootId });
    const result = harness.api.listDrivePhotoImportFolder(harness.tokenPayload({ folderId: '' }));
    assert.equal(result.ok, true);
    assert.equal(result.folder.id, rootId);
  });
  const overlong = createHarness({ rootId: 'r'.repeat(201) });
  assert.equal(overlong.api.listDrivePhotoImportFolder(
    overlong.tokenPayload({ folderId: '' })
  ).errorCode, 'DRIVE_IMPORT_ROOT_MISSING');
  assert.deepEqual(overlong.audit.folderReads, []);
});

test('containment rejects a trashed configured root and trashed intermediate ancestor', () => {
  const rootTrashed = createHarness();
  rootTrashed.addFolder(rootTrashed.rootId, 'Root', [], { trashed: true });
  const child = rootTrashed.addFolder('child_AAAAAAAAAAA', 'Child', [rootTrashed.rootId]);
  assert.equal(rootTrashed.api.isDriveFolderWithinRoot_(child, rootTrashed.rootId), false);
  assert.equal(rootTrashed.api.listDrivePhotoImportFolder(
    rootTrashed.tokenPayload({ folderId: 'child_AAAAAAAAAAA' })
  ).errorCode, 'DRIVE_IMPORT_FOLDER_OUTSIDE_ROOT');

  const middleTrashed = createHarness();
  middleTrashed.addFolder('middle_AAAAAAAAAA', 'Middle', [middleTrashed.rootId], { trashed: true });
  const nested = middleTrashed.addFolder('nested_AAAAAAAAAA', 'Nested', ['middle_AAAAAAAAAA']);
  assert.equal(middleTrashed.api.isDriveFolderWithinRoot_(nested, middleTrashed.rootId), false);
});

test('containment enforces exact depth and node bounds without parent-order exceptions', () => {
  const depth = createHarness();
  let parentId = depth.rootId;
  for (let index = 0; index < 64; index += 1) {
    const id = `depth_${String(index).padStart(12, '0')}`;
    depth.addFolder(id, `Depth ${index}`, [parentId]);
    parentId = id;
  }
  const depth64 = depth.folders.get(parentId);
  const depth65 = depth.addFolder('depth_over_AAAAAAA', 'Depth over', [parentId]);
  assert.equal(depth.api.isDriveFolderWithinRoot_(depth64, depth.rootId), true);
  assert.equal(depth.api.isDriveFolderWithinRoot_(depth65, depth.rootId), false);

  const nodes = createHarness();
  const outsideParents = Array.from({ length: 256 }, (_, index) => {
    const id = `outside_${String(index).padStart(12, '0')}`;
    nodes.addFolder(id, `Outside ${index}`);
    return id;
  });
  const atLimit = nodes.addFolder('nodes_limit_AAAAAA', 'At limit', outsideParents.slice(0, 255).concat(nodes.rootId));
  const overLimit = nodes.addFolder('nodes_over_AAAAAAA', 'Over limit', outsideParents.concat(nodes.rootId));
  assert.equal(nodes.api.isDriveFolderWithinRoot_(atLimit, nodes.rootId), true);
  assert.equal(nodes.api.isDriveFolderWithinRoot_(overLimit, nodes.rootId), false);

  ['valid-first', 'broken-first'].forEach((order) => {
    const harness = createHarness();
    const valid = harness.addFolder('valid_parent_AAAAA', 'Valid', [harness.rootId]);
    const broken = harness.addFolder('broken_parent_AAAA', 'Broken', [], { parentError: true });
    const parentIds = order === 'valid-first'
      ? [valid.getId(), broken.getId()] : [broken.getId(), valid.getId()];
    const subject = harness.addFolder(`subject_${order.replace('-', '_')}AAAA`, 'Subject', parentIds);
    assert.equal(harness.api.isDriveFolderWithinRoot_(subject, harness.rootId), false, order);
  });

  ['valid-first', 'error-first'].forEach((order) => {
    const harness = createHarness();
    const valid = harness.addFolder('valid_trash_AAAAAA', 'Valid', [harness.rootId]);
    const unreadable = harness.addFolder('unreadable_trash_A', 'Unreadable', [], { trashError: true });
    const parentIds = order === 'valid-first'
      ? [valid.getId(), unreadable.getId()] : [unreadable.getId(), valid.getId()];
    const subject = harness.addFolder(`trash_subject_${order.replace('-', '_')}AA`, 'Subject', parentIds);
    assert.equal(harness.api.isDriveFolderWithinRoot_(subject, harness.rootId), false, order);
  });

  const multiple = createHarness();
  const valid = multiple.addFolder('valid_branch_AAAAA', 'Valid', [multiple.rootId]);
  const outside = multiple.addFolder('outside_branch_AAA', 'Outside');
  const cycleA = multiple.addFolder('cycle_branch_AAAAA', 'Cycle A', ['cycle_branch_BBBBB']);
  multiple.addFolder('cycle_branch_BBBBB', 'Cycle B', [cycleA.getId()]);
  const trashed = multiple.addFolder('trashed_branch_AAA', 'Trashed', [], { trashed: true });
  const insideOutside = multiple.addFile('multiple_inside_AAA', 'inside.jpg', 'image/jpeg', [1],
    [outside.getId(), valid.getId()]);
  const insideCycle = multiple.addFile('multiple_cycle_AAAA', 'cycle.jpg', 'image/jpeg', [1],
    [cycleA.getId(), valid.getId()]);
  const insideTrashed = multiple.addFile('multiple_trash_AAAA', 'trash.jpg', 'image/jpeg', [1],
    [trashed.getId(), valid.getId()]);
  assert.equal(multiple.api.isDriveFileWithinRoot_(insideOutside, multiple.rootId), true);
  assert.equal(multiple.api.isDriveFileWithinRoot_(insideCycle, multiple.rootId), true);
  assert.equal(multiple.api.isDriveFileWithinRoot_(insideTrashed, multiple.rootId), true);
});
