const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');

const imageFixtures = Object.freeze({
  jpeg: Buffer.from(
    '/9j/4AAQSkZJRgABAQEASABIAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/wAALCAABAAEBAREA/8QAFAABAAAAAAAAAAAAAAAAAAAACf/EABQQAQAAAAAAAAAAAAAAAAAAAAD/2gAIAQEAAD8AVIP/2Q==',
    'base64'
  ),
  png: Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
    'base64'
  ),
  webp: Buffer.from('UklGRiIAAABXRUJQVlA4IBYAAAAwAQCdASoBAAEADsD+JaQAA3AAAAAA', 'base64'),
  heic: Buffer.from('00000018667479706865696300000000686569636d696631', 'hex'),
  heif: Buffer.from('00000018667479706865696600000000686569666d696631', 'hex')
});

function loadDriveClientModules(extra = {}) {
  const start = indexHtml.indexOf('const PhotoImportFileTypeCore = (function()');
  const end = indexHtml.indexOf('    const MultiPhotoImportBuilder =', start);
  assert.notEqual(start, -1, 'Expected PhotoImportFileTypeCore');
  assert.notEqual(end, -1, 'Expected Drive source modules before MultiPhotoImportBuilder');
  const context = {
    console,
    Number,
    Object,
    Array,
    String,
    Error,
    RegExp,
    Set,
    Promise,
    Date,
    Math,
    ImportJobCore: { MAX_ITEMS: 20 },
    ...extra
  };
  vm.createContext(context);
  vm.runInContext(
    indexHtml.slice(start, end) + '\n'
      + 'globalThis.__fileTypeCore = PhotoImportFileTypeCore;\n'
      + 'globalThis.__sourceCore = DrivePhotoImportSourceCore;\n'
      + 'globalThis.__loader = typeof DrivePhotoImportLoader === "undefined" ? null : DrivePhotoImportLoader;\n'
      + 'globalThis.__ui = typeof DrivePhotoImportUI === "undefined" ? null : DrivePhotoImportUI;',
    context
  );
  return {
    fileTypeCore: context.__fileTypeCore,
    sourceCore: context.__sourceCore,
    loader: context.__loader,
    ui: context.__ui
  };
}

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

function descriptor(overrides = {}) {
  return {
    id: 'photo_AAAAAAAAAAA',
    name: 'photo.jpg',
    mimeType: 'image/jpeg',
    sizeBytes: 3,
    modifiedAt: '2026-07-12T01:02:03.000Z',
    kind: 'jpeg',
    imported: false,
    ...overrides
  };
}

function fileResponse(value = descriptor(), base64 = 'AQID') {
  const { imported: _imported, ...file } = value;
  return { ok: true, file: { ...file, base64 } };
}

module.exports = { loadDriveClientModules, plain, descriptor, fileResponse, imageFixtures };
