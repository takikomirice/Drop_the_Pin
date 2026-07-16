const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const {
  loadDriveClientModules, plain, descriptor, fileResponse
} = require('./drive-photo-import-client-harness');

const indexHtml = fs.readFileSync(path.resolve(__dirname, '..', 'index.html'), 'utf8');

function functionSource(name) {
  const start = indexHtml.indexOf(`function ${name}(`);
  assert.notEqual(start, -1, `Expected function ${name}`);
  const bodyStart = indexHtml.indexOf('{', start);
  let depth = 0;
  for (let index = bodyStart; index < indexHtml.length; index += 1) {
    if (indexHtml[index] === '{') depth += 1;
    if (indexHtml[index] === '}') {
      depth -= 1;
      if (depth === 0) return indexHtml.slice(start, index + 1);
    }
  }
  throw new Error(`Unclosed function ${name}`);
}

test('photo type core preserves local JPEG PNG WebP HEIC HEIF and generic MIME behavior', () => {
  const { fileTypeCore } = loadDriveClientModules();
  const accepted = [
    ['a.jpg', 'image/jpeg', 'jpeg', 'image/jpeg'],
    ['a.jpeg', 'image/jpg', 'jpeg', 'image/jpeg'],
    ['a.png', 'image/png', 'png', 'image/png'],
    ['a.webp', '', 'webp', 'image/webp'],
    ['a.heic', 'application/octet-stream', 'heic', 'image/heic'],
    ['a.heif', 'binary/octet-stream', 'heic', 'image/heif']
  ];
  accepted.forEach(([name, type, kind, normalizedMimeType]) => {
    assert.deepEqual(plain(fileTypeCore.classify({ name, type, size: 1 })), {
      supported: true,
      kind,
      normalizedMimeType,
      extension: name.split('.').pop() === 'jpeg' ? 'jpg' : name.split('.').pop()
    });
  });
  [
    ['photo.jpg', 'image/svg+xml'],
    ['photo.svg', 'image/jpeg'],
    ['photo.jpg', 'video/mp4'],
    ['photo.gif', 'image/gif'],
    ['photo.bin', 'application/octet-stream'],
    ['photo.png', 'image/jpeg']
  ].forEach(([name, type]) => assert.equal(fileTypeCore.classify({ name, type, size: 1 }).supported, false));
});

test('descriptor validation uses own fields, fixed projection, safe basename, and consistent type contract', () => {
  const { sourceCore } = loadDriveClientModules();
  const source = descriptor({ arbitrary: 'secret' });
  const normalized = sourceCore.validateFileDescriptor(source);
  assert.deepEqual(plain(normalized), descriptor());
  assert.equal(Object.isFrozen(normalized), true);
  assert.deepEqual(source, descriptor({ arbitrary: 'secret' }));

  const inherited = Object.create(descriptor());
  assert.throws(() => sourceCore.validateFileDescriptor(inherited),
    (error) => error.code === 'DRIVE_IMPORT_SELECTION_INVALID');
  const invalid = [
    descriptor({ id: '__proto__' }),
    descriptor({ id: 'constructor' }),
    descriptor({ name: '../photo.jpg' }),
    descriptor({ name: 'folder\\photo.jpg' }),
    descriptor({ name: 'bad\nname.jpg' }),
    descriptor({ name: 'x'.repeat(256) }),
    descriptor({ mimeType: ' image/jpeg' }),
    descriptor({ mimeType: 'image/png' }),
    descriptor({ kind: 'png' }),
    descriptor({ sizeBytes: 0 }),
    descriptor({ sizeBytes: 1.5 }),
    descriptor({ sizeBytes: 15 * 1024 * 1024 + 1 }),
    descriptor({ modifiedAt: 'not-a-date' }),
    descriptor({ modifiedAt: 'July 12, 2026 01:02:03 UTC' }),
    descriptor({ imported: 'false' }),
    (() => {
      const value = descriptor();
      delete value.imported;
      return value;
    })()
  ];
  invalid.forEach((value) => assert.throws(
    () => sourceCore.validateFileDescriptor(value),
    (error) => error.code === 'DRIVE_IMPORT_SELECTION_INVALID'
  ));

  let sizeReads = 0;
  const changingSize = descriptor();
  Object.defineProperty(changingSize, 'sizeBytes', {
    enumerable: true,
    get() {
      sizeReads += 1;
      return sizeReads <= 4 ? 3 : -1;
    }
  });
  assert.equal(sourceCore.validateFileDescriptor(changingSize).sizeBytes, 3);
  assert.equal(sizeReads, 1);

  const throwingId = descriptor();
  Object.defineProperty(throwingId, 'id', {
    enumerable: true,
    get() { throw new Error('private descriptor getter'); }
  });
  assert.throws(() => sourceCore.validateFileDescriptor(throwingId), (error) => (
    error.code === 'DRIVE_IMPORT_SELECTION_INVALID' && !/private/.test(error.message)
  ));
});

test('listing descriptors retain only a boolean imported marker while file responses keep the base descriptor contract', () => {
  const { sourceCore } = loadDriveClientModules();
  const imported = descriptor({ imported: true, owner: 'private-owner' });
  const response = {
    ok: true,
    folder: { id: 'root_ABCDEFGHIJKLMNO', name: 'Root', isRoot: true },
    parent: null,
    folders: [],
    photos: [imported],
    ignoredUnsupportedFileCount: 0,
    counts: { folders: 0, photos: 1 }
  };
  const normalized = sourceCore.validateFolderResponse(response);
  assert.equal(normalized.photos[0].imported, true);
  assert.equal(Object.hasOwn(normalized.photos[0], 'owner'), false);
  const validatedFile = plain(sourceCore.validateFileResponse(fileResponse(imported), imported));
  assert.equal(validatedFile.file.id, imported.id);
  assert.equal(validatedFile.file.base64, 'AQID');
  assert.equal(Object.hasOwn(validatedFile.file, 'imported'), false);
  assert.equal(Object.hasOwn(validatedFile.file, 'owner'), false);
});

test('selection accepts one and twenty, preserves order, and enforces duplicate/count/100MB atomically', () => {
  const { sourceCore } = loadDriveClientModules();
  const one = sourceCore.validateSelection([descriptor()]);
  assert.equal(one.length, 1);
  assert.equal(sourceCore.validateSelection([
    descriptor({ sizeBytes: 15 * 1024 * 1024 })
  ])[0].sizeBytes, 15 * 1024 * 1024);

  const twentyInput = Array.from({ length: 20 }, (_, index) => descriptor({
    id: `photo_${String(index).padStart(12, '0')}`,
    name: `${index}.jpg`,
    sizeBytes: 5 * 1024 * 1024
  }));
  const twenty = sourceCore.validateSelection(twentyInput);
  assert.equal(twenty.length, 20);
  assert.deepEqual(plain(twenty).map((item) => item.id), twentyInput.map((item) => item.id));
  assert.equal(Object.isFrozen(twenty), true);

  assert.throws(() => sourceCore.validateSelection([]),
    (error) => error.code === 'DRIVE_IMPORT_SELECTION_EMPTY');
  assert.throws(() => sourceCore.validateSelection(new Array(1)),
    (error) => error.code === 'DRIVE_IMPORT_SELECTION_INVALID');
  assert.throws(() => sourceCore.validateSelection(twentyInput.concat(descriptor({ id: 'photo_extra_AAAAAA' }))),
    (error) => error.code === 'DRIVE_IMPORT_SELECTION_LIMIT_EXCEEDED');
  assert.throws(() => sourceCore.validateSelection([descriptor(), descriptor()]),
    (error) => error.code === 'DRIVE_IMPORT_SELECTION_DUPLICATE');

  const exact = Array.from({ length: 10 }, (_, index) => descriptor({
    id: `exact_${String(index).padStart(12, '0')}`,
    name: `${index}.jpg`,
    sizeBytes: 10 * 1024 * 1024
  }));
  assert.equal(sourceCore.validateSelection(exact).length, 10);
  const over = exact.map((item, index) => index === 9 ? { ...item, sizeBytes: item.sizeBytes + 1 } : item);
  assert.throws(() => sourceCore.validateSelection(over),
    (error) => error.code === 'DRIVE_IMPORT_SELECTION_TOO_LARGE');
});

test('selection rejects formally imported photos even when a hostile caller bypasses the disabled checkbox', () => {
  const { sourceCore } = loadDriveClientModules();
  assert.throws(
    () => sourceCore.validateSelection([descriptor({ imported: true })]),
    (error) => error.code === 'DRIVE_IMPORT_SELECTION_INVALID'
  );
});

test('folder response is strictly whitelisted and normalizes photo descriptors without mutation', () => {
  const { sourceCore } = loadDriveClientModules();
  const response = {
    ok: true,
    folder: { id: 'folder_AAAAAAAAAA', name: 'Folder', isRoot: false, secret: 'x' },
    parent: { id: 'root_ABCDEFGHIJKLMNO', name: 'Root', url: 'secret' },
    folders: [{ id: 'child_AAAAAAAAAAA', name: 'Child', owner: 'secret' }],
    photos: [descriptor({ arbitrary: 'secret' })],
    ignoredUnsupportedFileCount: 2,
    counts: { folders: 1, photos: 1, arbitrary: 9 },
    arbitrary: 'secret'
  };
  const normalized = sourceCore.validateFolderResponse(response);
  assert.deepEqual(plain(normalized), {
    ok: true,
    folder: { id: 'folder_AAAAAAAAAA', name: 'Folder', isRoot: false },
    parent: { id: 'root_ABCDEFGHIJKLMNO', name: 'Root' },
    folders: [{ id: 'child_AAAAAAAAAAA', name: 'Child' }],
    photos: [descriptor()],
    ignoredUnsupportedFileCount: 2,
    counts: { folders: 1, photos: 1 }
  });
  assert.equal(response.arbitrary, 'secret');

  const sparsePhotos = { ...response, photos: new Array(1) };
  assert.throws(() => sourceCore.validateFolderResponse(sparsePhotos),
    (error) => error.code === 'DRIVE_IMPORT_RESPONSE_INVALID');

  const throwingFolder = { ...response };
  Object.defineProperty(throwingFolder, 'folder', {
    enumerable: true,
    get() { throw new Error('private folder getter'); }
  });
  assert.throws(() => sourceCore.validateFolderResponse(throwingFolder), (error) => (
    error.code === 'DRIVE_IMPORT_RESPONSE_INVALID' && !/private/.test(error.message)
  ));
});

test('folder response rejects duplicate identities and inconsistent root parent topology', () => {
  const { sourceCore } = loadDriveClientModules();
  const base = {
    ok: true,
    folder: { id: 'folder_AAAAAAAAAA', name: 'Folder', isRoot: false },
    parent: { id: 'root_ABCDEFGHIJKLMNO', name: 'Root' },
    folders: [{ id: 'child_AAAAAAAAAAA', name: 'Child' }],
    photos: [descriptor()],
    ignoredUnsupportedFileCount: 0,
    counts: { folders: 1, photos: 1 }
  };
  const invalid = [
    { ...base, folders: [base.folders[0], { ...base.folders[0] }], counts: { folders: 2, photos: 1 } },
    { ...base, photos: [base.photos[0], { ...base.photos[0] }], counts: { folders: 1, photos: 2 } },
    { ...base, folders: [{ id: base.photos[0].id, name: 'Collision' }] },
    { ...base, parent: { id: base.folder.id, name: 'Self' } },
    { ...base, folders: [{ id: base.folder.id, name: 'Current again' }] },
    { ...base, folder: { ...base.folder, isRoot: true } },
    { ...base, parent: null }
  ];
  invalid.forEach((response) => assert.throws(
    () => sourceCore.validateFolderResponse(response),
    (error) => error.code === 'DRIVE_IMPORT_RESPONSE_INVALID'
  ));

  const root = sourceCore.validateFolderResponse({
    ...base,
    folder: { id: 'root_ABCDEFGHIJKLMNO', name: 'Root', isRoot: true },
    parent: null
  });
  assert.equal(root.folder.isRoot, true);
  assert.equal(root.parent, null);
});

test('folder response bounds raw arrays before reading entries and permits only safe display names', () => {
  const { sourceCore } = loadDriveClientModules();
  let entryReads = 0;
  const oversized = Array.from({ length: 501 }, (_, index) => {
    const entry = { name: `Folder ${index}` };
    Object.defineProperty(entry, 'id', {
      enumerable: true,
      get() { entryReads += 1; return `folder_${String(index).padStart(12, '0')}`; }
    });
    return entry;
  });
  assert.throws(() => sourceCore.validateFolderResponse({
    ok: true,
    folder: { id: 'root_ABCDEFGHIJKLMNO', name: 'Root', isRoot: true },
    parent: null,
    folders: oversized,
    photos: [],
    ignoredUnsupportedFileCount: 0,
    counts: { folders: 501, photos: 0 }
  }), (error) => error.code === 'DRIVE_IMPORT_RESPONSE_INVALID');
  assert.equal(entryReads, 0);

  const valid = sourceCore.validateFolderResponse({
    ok: true,
    folder: { id: 'root_ABCDEFGHIJKLMNO', name: '<b>A/B\\C</b>', isRoot: true },
    parent: null,
    folders: [],
    photos: [],
    ignoredUnsupportedFileCount: 0,
    counts: { folders: 0, photos: 0 }
  });
  assert.equal(valid.folder.name, '<b>A/B\\C</b>');

  ['bad\nname', 'bad\u0085name', 'bad\u202Ename'].forEach((name) => {
    assert.throws(() => sourceCore.validateFolderResponse({
      ok: true,
      folder: { id: 'root_ABCDEFGHIJKLMNO', name, isRoot: true },
      parent: null,
      folders: [],
      photos: [],
      ignoredUnsupportedFileCount: 0,
      counts: { folders: 0, photos: 0 }
    }), (error) => error.code === 'DRIVE_IMPORT_RESPONSE_INVALID');
  });
});

test('file response uses only the requested id and keeps decoded size out of metadata validation', () => {
  const { sourceCore } = loadDriveClientModules();
  const expected = descriptor();
  const bytes = Buffer.alloc(4096, 0x5a);
  const responseFile = descriptor({
    name: 'current.jpeg',
    mimeType: 'IMAGE/JPG',
    sizeBytes: 1,
    modifiedAt: '2026-07-15T10:20:30.000Z',
    kind: 'ignored-server-kind'
  });
  const valid = sourceCore.validateFileResponse(fileResponse(responseFile, bytes.toString('base64')), expected);
  assert.deepEqual(plain(valid), {
    ok: true,
    file: {
      id: expected.id,
      name: 'current.jpeg',
      mimeType: 'image/jpeg',
      modifiedAt: '2026-07-15T10:20:30.000Z',
      kind: 'jpeg',
      base64: bytes.toString('base64')
    }
  });
  assert.equal(Object.isFrozen(valid), true);
  assert.equal(Object.hasOwn(valid.file, 'sizeBytes'), false);

  assert.throws(() => sourceCore.validateFileResponse(
    fileResponse(descriptor({ id: 'other_AAAAAAAAAAAA' })), expected
  ), (error) => error.code === 'DRIVE_IMPORT_RESPONSE_INVALID');
  [
    descriptor({ name: '../photo.jpg' }),
    descriptor({ name: 'photo.gif', mimeType: 'image/gif', kind: 'gif' }),
    descriptor({ name: 'photo.png', mimeType: 'image/jpeg' })
  ].forEach((actual) => assert.throws(
    () => sourceCore.validateFileResponse(fileResponse(actual), expected),
    (error) => error.code === 'DRIVE_IMPORT_RESPONSE_INVALID'
  ));
  ['', 'data:image/jpeg;base64,AQID', 'AQ I D', '***=', 'AQI', 'AQID\n'].forEach((base64) => {
    assert.throws(() => sourceCore.validateFileResponse(fileResponse(expected, base64), expected),
      (error) => error.code === 'DRIVE_IMPORT_RESPONSE_INVALID');
  });
  let base64Reads = 0;
  const changingResponse = fileResponse(expected);
  Object.defineProperty(changingResponse.file, 'base64', {
    enumerable: true,
    get() {
      base64Reads += 1;
      return base64Reads <= 2 ? 'AQID' : 'data:image/jpeg;base64,AQID';
    }
  });
  assert.equal(sourceCore.validateFileResponse(changingResponse, expected).file.base64, 'AQID');
  assert.equal(base64Reads, 1);
});

test('file response accepts standard Base64 alphabet and every valid padding form', () => {
  const { sourceCore } = loadDriveClientModules();
  const expected = descriptor({ sizeBytes: 3 });
  const fixtures = [
    { bytes: Buffer.from([0xfb, 0xff, 0xff]), base64: '+///' },
    { bytes: Buffer.from([0xfb, 0xff]), base64: '+/8=' },
    { bytes: Buffer.from([0xff]), base64: '/w==' }
  ];
  fixtures.forEach(({ bytes, base64 }) => {
    assert.equal(bytes.toString('base64'), base64);
    const validated = sourceCore.validateFileResponse(fileResponse(descriptor({
      sizeBytes: 999,
      modifiedAt: '2026-07-16T01:02:03.000Z'
    }), base64), expected);
    assert.equal(Object.hasOwn(validated.file, 'sizeBytes'), false);
    assert.equal(validated.file.base64, base64);
  });
});

test('file response accepts hundreds of KB through multi-megabyte Base64 without stack overflow', () => {
  const { sourceCore } = loadDriveClientModules();
  const expected = descriptor({ sizeBytes: 3 });
  [384 * 1024, 2 * 1024 * 1024, 4 * 1024 * 1024].forEach((sizeBytes) => {
    const bytes = Buffer.allocUnsafe(sizeBytes);
    for (let index = 0; index < bytes.length; index += 1) {
      bytes[index] = (index * 31 + 17) & 0xff;
    }
    const base64 = bytes.toString('base64');

    const validated = sourceCore.validateFileResponse(fileResponse(expected, base64), expected);

    assert.equal(Object.hasOwn(validated.file, 'sizeBytes'), false);
    assert.equal(validated.file.base64, base64);
  });
});

test('file response reports only safe Base64 diagnostic stages', () => {
  const { sourceCore } = loadDriveClientModules();
  const expected = descriptor();
  const missing = fileResponse(expected);
  delete missing.file.base64;
  assert.throws(() => sourceCore.validateFileResponse(missing, expected), (error) => (
    error.code === 'DRIVE_IMPORT_RESPONSE_INVALID'
      && error.diagnosticStage === 'base64_missing'
      && !/photo_|AQID/.test(error.message)
  ));

  const unreadable = fileResponse(expected);
  Object.defineProperty(unreadable.file, 'base64', {
    enumerable: true,
    get() { throw new Error('private Base64 getter AQID'); }
  });
  assert.throws(() => sourceCore.validateFileResponse(unreadable, expected), (error) => (
    error.code === 'DRIVE_IMPORT_RESPONSE_INVALID'
      && error.diagnosticStage === 'base64_read_failed'
      && !/private|AQID/.test(error.message)
  ));

  for (const base64 of [null, 123, '', 'data:image/jpeg;base64,AQID', 'AQ I', 'AQ\nI=', 'A=ID', 'AQ*D']) {
    const expectedStage = typeof base64 === 'string' && base64
      ? 'base64_format_invalid' : 'base64_not_string';
    assert.throws(() => sourceCore.validateFileResponse(fileResponse(expected, base64), expected),
      (error) => error.code === 'DRIVE_IMPORT_RESPONSE_INVALID'
        && error.diagnosticStage === expectedStage);
  }
});

test('file response snapshots metadata without projecting Base64 and reads Base64 once', () => {
  const source = functionSource('validateFileResponse');
  const snapshotCalls = source.match(/snapshotOwnFields\([\s\S]*?\);/g) || [];
  assert.equal(snapshotCalls.some((call) => /base64/.test(call)), false);
  assert.match(source, /base64\s*=\s*responseSnapshot\.file\.base64/);
  assert.doesNotMatch(source, /console\.|JSON\.stringify|normalizeDisplayName\(base64/);
});
