const assert = require('node:assert/strict');
const test = require('node:test');
const {
  loadDriveClientModules, descriptor, fileResponse, imageFixtures
} = require('./drive-photo-import-client-harness');

class BrowserFile extends Blob {
  constructor(parts, name, options = {}) {
    super(parts, options);
    this.name = name;
    this.lastModified = options.lastModified;
  }
}

function environment(overrides = {}) {
  return {
    atob(value) { return Buffer.from(value, 'base64').toString('binary'); },
    Uint8Array,
    Blob,
    File: BrowserFile,
    ...overrides
  };
}

test('materializeFile returns a real File with exact name type size and lastModified', () => {
  const { sourceCore } = loadDriveClientModules();
  const expected = descriptor({ sizeBytes: imageFixtures.jpeg.length });
  const validated = sourceCore.validateFileResponse(
    fileResponse(expected, imageFixtures.jpeg.toString('base64')), expected
  );
  const file = sourceCore.materializeFile(validated, environment());
  assert.equal(file instanceof BrowserFile, true);
  assert.equal(file.name, expected.name);
  assert.equal(file.type, expected.mimeType);
  assert.equal(file.size, imageFixtures.jpeg.length);
  assert.equal(file.lastModified, Date.parse(expected.modifiedAt));
  assert.equal(Object.prototype.hasOwnProperty.call(file, 'id'), false);
  assert.equal(Object.prototype.hasOwnProperty.call(file, '__editToken'), false);
  assert.equal(Object.prototype.hasOwnProperty.call(file, 'base64'), false);
});

test('materializeFile uses read metadata and actual padded bytes instead of listing metadata', () => {
  const { sourceCore } = loadDriveClientModules();
  const listed = descriptor({
    name: 'listed.heic',
    mimeType: 'image/heic',
    sizeBytes: 3,
    modifiedAt: '2026-07-12T01:02:03.000Z',
    kind: 'heic'
  });
  const bytes = Buffer.alloc(4096, 0x7f);
  const read = descriptor({
    id: listed.id,
    name: 'current.jpeg',
    mimeType: 'IMAGE/JPG',
    sizeBytes: 1,
    modifiedAt: '2026-07-15T10:20:30.000Z',
    kind: 'ignored'
  });
  const validated = sourceCore.validateFileResponse(fileResponse(read, bytes.toString('base64')), listed);
  const file = sourceCore.materializeFile(validated, environment());
  assert.equal(file.name, 'current.jpeg');
  assert.equal(file.type, 'image/jpeg');
  assert.equal(file.size, bytes.length);
  assert.equal(file.lastModified, Date.parse(read.modifiedAt));
});

test('materializeFile treats invalid modifiedAt as an auxiliary value and falls back to now', () => {
  const { sourceCore } = loadDriveClientModules();
  const response = fileResponse(descriptor({ modifiedAt: 'not-a-date' }));
  const validated = sourceCore.validateFileResponse(response, descriptor());
  const before = Date.now();
  const file = sourceCore.materializeFile(validated, environment());
  const after = Date.now();
  assert.equal(file.lastModified >= before && file.lastModified <= after, true);
});

test('materializeFile safely rejects unavailable or lookalike File APIs and decode anomalies', () => {
  const { sourceCore } = loadDriveClientModules();
  const expected = descriptor();
  const validated = sourceCore.validateFileResponse(fileResponse(expected), expected);
  assert.throws(() => sourceCore.materializeFile(validated, environment({ File: null })),
    (error) => error.code === 'DRIVE_IMPORT_FILE_API_UNAVAILABLE');
  function LookalikeFile(parts, name, options) {
    return new Blob(parts, options);
  }
  assert.throws(() => sourceCore.materializeFile(validated, environment({ File: LookalikeFile })),
    (error) => error.code === 'DRIVE_IMPORT_RESPONSE_INVALID'
      && error.diagnosticStage === 'file_result_invalid');
  class NonBlobFile {
    constructor(parts, name, options) {
      this.name = name;
      this.type = options.type;
      this.size = 3;
      this.lastModified = options.lastModified;
    }
  }
  assert.throws(() => sourceCore.materializeFile(validated, environment({ File: NonBlobFile })),
    (error) => error.code === 'DRIVE_IMPORT_RESPONSE_INVALID'
      && error.diagnosticStage === 'file_result_invalid');
  assert.throws(() => sourceCore.materializeFile(validated, environment({
    atob() { throw new Error('private base64'); }
  })), (error) => error.code === 'DRIVE_IMPORT_RESPONSE_INVALID'
    && error.diagnosticStage === 'base64_decode_failed'
    && !/private|AQID/.test(error.message));
  assert.throws(() => sourceCore.materializeFile(validated, environment({
    atob() { return null; }
  })), (error) => error.code === 'DRIVE_IMPORT_RESPONSE_INVALID'
    && error.diagnosticStage === 'base64_decode_type_invalid');
  assert.throws(() => sourceCore.materializeFile(validated, environment({
    atob() { return '\x01\x02' + String.fromCharCode(300); }
  })), (error) => error.code === 'DRIVE_IMPORT_RESPONSE_INVALID'
    && error.diagnosticStage === 'base64_decode_type_invalid');
});

test('materializeFile uses decoded binary length as the only size source', () => {
  const { sourceCore } = loadDriveClientModules();
  const expected = descriptor({ sizeBytes: 999 });
  const validated = sourceCore.validateFileResponse(fileResponse(expected), expected);
  const file = sourceCore.materializeFile(validated, environment({
    atob() { return '\x01\x02'; }
  }));

  assert.equal(file.size, 2);
  assert.equal(Object.hasOwn(validated.file, 'sizeBytes'), false);

  assert.throws(() => sourceCore.materializeFile(validated, environment({
    atob() { return ''; }
  })), (error) => error.code === 'DRIVE_IMPORT_RESPONSE_INVALID'
    && error.diagnosticStage === 'decoded_size_empty');
  assert.throws(() => sourceCore.materializeFile(validated, environment({
    atob() { return 'x'.repeat(sourceCore.MAX_FILE_BYTES + 1); }
  })), (error) => error.code === 'DRIVE_IMPORT_RESPONSE_INVALID'
    && error.diagnosticStage === 'decoded_size_too_large');
});

test('materializeFile accepts cross-realm Blob identity and browser-normalized optional fields', () => {
  const { sourceCore } = loadDriveClientModules();
  const expected = descriptor();
  const validated = sourceCore.validateFileResponse(fileResponse(expected), expected);
  class CrossRealmFile extends Blob {
    constructor(parts, name) {
      super(parts);
      this.name = name;
      this.lastModified = 0;
    }
  }
  class OtherRealmBlob {}

  const file = sourceCore.materializeFile(validated, environment({
    File: CrossRealmFile,
    Blob: OtherRealmBlob
  }));

  assert.equal(file instanceof CrossRealmFile, true);
  assert.equal(file instanceof OtherRealmBlob, false);
  assert.equal(file.name, expected.name);
  assert.equal(file.type, '');
  assert.equal(file.size, 3);
  assert.equal(file.lastModified, 0);
});

test('materializeFile preserves real JPEG PNG and WebP binary signatures', async () => {
  const { sourceCore } = loadDriveClientModules();
  const samples = [
    ['jpeg', descriptor({ name: 'pixel.jpg', mimeType: 'image/jpeg', kind: 'jpeg' })],
    ['png', descriptor({ name: 'pixel.png', mimeType: 'image/png', kind: 'png' })],
    ['webp', descriptor({ name: 'pixel.webp', mimeType: 'image/webp', kind: 'webp' })]
  ];

  for (const [kind, expected] of samples) {
    const fixture = imageFixtures[kind];
    const validated = sourceCore.validateFileResponse(
      fileResponse(expected, fixture.toString('base64')), expected
    );
    const file = sourceCore.materializeFile(validated, environment());
    const bytes = Buffer.from(await file.arrayBuffer());
    assert.deepEqual(bytes, fixture);
  }

  assert.deepEqual(imageFixtures.jpeg.subarray(0, 3), Buffer.from([0xff, 0xd8, 0xff]));
  assert.deepEqual(imageFixtures.png.subarray(0, 8), Buffer.from('89504e470d0a1a0a', 'hex'));
  assert.equal(imageFixtures.webp.subarray(0, 4).toString('ascii'), 'RIFF');
  assert.equal(imageFixtures.webp.subarray(8, 12).toString('ascii'), 'WEBP');
});
