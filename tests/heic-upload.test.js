const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const codeJs = fs.readFileSync(path.join(root, 'Code.js'), 'utf8');

function extractFunction(source, name) {
  const functionStart = source.indexOf(`function ${name}(`);
  assert.notEqual(functionStart, -1, `Expected function ${name} to exist`);
  const start = source.slice(Math.max(0, functionStart - 6), functionStart) === 'async '
    ? functionStart - 6
    : functionStart;
  const open = source.indexOf('{', start);
  let depth = 0;
  for (let index = open; index < source.length; index += 1) {
    if (source[index] === '{') depth += 1;
    if (source[index] === '}') depth -= 1;
    if (depth === 0) return source.slice(start, index + 1);
  }
  assert.fail(`Could not parse function ${name}`);
}

function createCoreApi(extra = {}) {
  const names = [
    'isHeicFile',
    'normalizeJpegFileName',
    'rationalToNumber',
    'gpsValueToDecimal',
    'tagValue',
    'normalizeGpsResult',
    'isUnsupportedMetadataError',
    'parseExifDateTimeOriginal',
    'readHeicMetadata',
    'convertHeicToJpeg',
    'prepareUploadPhoto',
    'locationMessageForMetadataStatus'
  ];
  const context = {
    Blob,
    File: globalThis.File,
    Promise,
    Number,
    String,
    Array,
    Object,
    Math,
    ...extra
  };
  vm.runInNewContext(`${names.map((name) => extractFunction(indexHtml, name)).join('\n')}
globalThis.__heicApi = { ${names.join(', ')} };`, context);
  return context.__heicApi;
}

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

test('detects HEIC and HEIF by MIME type or filename extension', () => {
  const { isHeicFile } = createCoreApi();

  assert.equal(isHeicFile({ name: 'photo.bin', type: 'image/heic' }), true);
  assert.equal(isHeicFile({ name: 'photo.HEIF', type: '' }), true);
  assert.equal(isHeicFile({ name: 'photo.heic', type: 'application/octet-stream' }), true);
  assert.equal(isHeicFile({ name: 'photo.jpeg', type: 'image/jpeg' }), false);
  assert.equal(isHeicFile({ name: 'photo.png', type: 'image/png' }), false);
  assert.equal(isHeicFile({ name: 'photo.webp', type: 'image/webp' }), false);
});

test('normalizes converted filenames to a JPEG suffix', () => {
  const { normalizeJpegFileName } = createCoreApi();

  assert.equal(normalizeJpegFileName('IMG_0001.HEIC'), 'IMG_0001.jpg');
  assert.equal(normalizeJpegFileName('trip.heif'), 'trip.jpg');
  assert.equal(normalizeJpegFileName('untitled'), 'untitled.jpg');
  assert.equal(normalizeJpegFileName('.heic'), 'image.jpg');
});

test('Drive filename sync preserves the converted JPEG extension', () => {
  const source = extractFunction(codeJs, 'buildFileNameForSave');
  const context = {};
  vm.runInNewContext(`${source}; globalThis.__build = buildFileNameForSave;`, context);

  assert.equal(context.__build('夏休み.heic', 'IMG_0001.jpg', true), '夏休み.jpg');
  assert.equal(context.__build('夏休み', 'IMG_0001.jpg', true), '夏休み.jpg');
  assert.equal(context.__build('旅行 1.2', 'IMG_0001.jpg', true), '旅行 1.2.jpg');
  assert.equal(context.__build('.hidden', 'IMG_0001.jpg', true), '.hidden.jpg');
  assert.equal(context.__build('夏休み', 'IMG_0001.jpg', false), 'IMG_0001.jpg');
});

test('normalizes numeric, rational, and DMS GPS values with correct hemispheres', () => {
  const { normalizeGpsResult } = createCoreApi();

  assert.deepEqual(plain(normalizeGpsResult({ gps: { Latitude: 35.25, Longitude: 139.5 } })), {
    status: 'success', lat: 35.25, lng: 139.5
  });
  assert.deepEqual(plain(normalizeGpsResult({
    GPSLatitude: { value: [{ numerator: 35, denominator: 1 }, { numerator: 30, denominator: 1 }, { numerator: 0, denominator: 1 }] },
    GPSLatitudeRef: { value: ['S'] },
    GPSLongitude: { value: [[139, 1], [45, 1], [0, 1]] },
    GPSLongitudeRef: { value: ['W'] }
  })), { status: 'success', lat: -35.5, lng: -139.75 });
});

test('distinguishes missing GPS from invalid and out-of-range coordinates', () => {
  const { normalizeGpsResult } = createCoreApi();

  assert.deepEqual(plain(normalizeGpsResult({})), { status: 'no-gps', lat: null, lng: null });
  assert.deepEqual(plain(normalizeGpsResult({ gps: { Latitude: 35 } })), { status: 'invalid-gps', lat: null, lng: null });
  assert.deepEqual(plain(normalizeGpsResult({ gps: { Latitude: 91, Longitude: 139 } })), { status: 'invalid-gps', lat: null, lng: null });
  assert.deepEqual(plain(normalizeGpsResult({ gps: { Latitude: 35, Longitude: 181 } })), { status: 'invalid-gps', lat: null, lng: null });
  assert.deepEqual(plain(normalizeGpsResult({ GPSLatitude: [35, 60, 0], GPSLongitude: [139, 0, 0] })), { status: 'invalid-gps', lat: null, lng: null });
  assert.deepEqual(plain(normalizeGpsResult({ GPSLatitude: [35, -1, 0], GPSLongitude: [139, 0, 0] })), { status: 'invalid-gps', lat: null, lng: null });
  assert.deepEqual(plain(normalizeGpsResult({ GPSLatitude: [35, 0, 60], GPSLongitude: [139, 0, 0] })), { status: 'invalid-gps', lat: null, lng: null });
});

test('reads HEIC metadata from the original file and keeps GPS/date outcomes distinct', async () => {
  const { readHeicMetadata } = createCoreApi();
  const original = { name: 'IMG_0001.heic' };
  let received = null;
  const reader = {
    async load(file, options) {
      received = file;
      assert.equal(options.expanded, true);
      return {
        gps: { Latitude: 35.1, Longitude: 139.2 },
        exif: { DateTimeOriginal: { description: '2026:07:10 12:34:56' } }
      };
    }
  };

  assert.deepEqual(plain(await readHeicMetadata(original, reader)), {
    status: 'success', lat: 35.1, lng: 139.2, capturedAt: '2026-07-10T12:34'
  });
  assert.equal(received, original);
  assert.deepEqual(plain(await readHeicMetadata(original, null)), {
    status: 'unsupported', lat: null, lng: null, capturedAt: ''
  });
  assert.deepEqual(plain(await readHeicMetadata(original, { load: async () => ({}) })), {
    status: 'no-gps', lat: null, lng: null, capturedAt: ''
  });
  assert.deepEqual(plain(await readHeicMetadata(original, { load: async () => { throw new Error('Invalid image format'); } })), {
    status: 'unsupported', lat: null, lng: null, capturedAt: ''
  });
  assert.deepEqual(plain(await readHeicMetadata(original, { load: async () => { throw new Error('bad metadata'); } })), {
    status: 'read-failed', lat: null, lng: null, capturedAt: ''
  });
});

test('manual EXIF date read cannot restore stale metadata after removal or replacement', async () => {
  let resolveRead;
  const pendingRead = new Promise((resolve) => { resolveRead = resolve; });
  const original = { name: 'old.jpg' };
  const eventAt = { value: '' };
  const context = {
    state: { upload: { capturedAt: '', originalFile: original, selectionToken: 10 } },
    readExifDateTimeOriginal: () => pendingRead,
    document: { getElementById: () => eventAt },
    alert() {}
  };
  const source = extractFunction(indexHtml, 'fillUploadEventAtFromExif');
  vm.runInNewContext(`${source}; globalThis.__fill = fillUploadEventAtFromExif;`, context);
  const reading = context.__fill(false);
  context.state.upload.originalFile = { name: 'new.jpg' };
  context.state.upload.selectionToken += 1;
  resolveRead('2026-07-10T12:34');
  await reading;

  assert.equal(eventAt.value, '');
});

test('converts HEIC once after metadata parsing and reuses the JPEG result', async () => {
  class TestFile extends Blob {
    constructor(parts, name, options) {
      super(parts, options);
      this.name = name;
      this.lastModified = options.lastModified;
    }
  }
  const calls = [];
  const original = new TestFile(['heic'], 'IMG_0001.HEIC', { type: 'image/heic', lastModified: 123 });
  const jpegBlob = new Blob(['jpeg'], { type: 'image/jpeg' });
  const { prepareUploadPhoto } = createCoreApi();

  const result = await prepareUploadPhoto(original, {
    metadataReader: {
      load: async (file) => {
        calls.push(['metadata', file]);
        return { gps: { Latitude: 35, Longitude: 139 } };
      }
    },
    converter: async (options) => {
      calls.push(['convert', options.blob, options.type, options.quality]);
      return jpegBlob;
    },
    FileCtor: TestFile
  });

  assert.equal(calls.length, 2);
  assert.equal(calls[0][0], 'metadata');
  assert.equal(calls[0][1], original);
  assert.equal(calls[1][0], 'convert');
  assert.equal(calls[1][1], original);
  assert.equal(result.originalFile, original);
  assert.notEqual(result.uploadFile, original);
  assert.equal(result.uploadFile.type, 'image/jpeg');
  assert.equal(result.uploadFile.name, 'IMG_0001.jpg');
  assert.equal(result.metadataStatus, 'success');
  assert.equal(result.lat, 35);
  assert.equal(result.lng, 139);
});

test('uses only the first converter frame and reports conversion failure without an upload file', async () => {
  class TestFile extends Blob {
    constructor(parts, name, options) {
      super(parts, options);
      this.name = name;
      this.lastModified = options.lastModified;
    }
  }
  const original = new TestFile(['heic'], 'multi.heif', { type: 'image/heif', lastModified: 1 });
  const first = new Blob(['first'], { type: 'image/jpeg' });
  const second = new Blob(['second'], { type: 'image/jpeg' });
  const { prepareUploadPhoto } = createCoreApi();
  const metadataReader = { load: async () => ({}) };

  const converted = await prepareUploadPhoto(original, {
    metadataReader,
    converter: async () => [first, second],
    FileCtor: TestFile
  });
  assert.equal(await converted.uploadFile.text(), 'first');

  const failed = await prepareUploadPhoto(original, {
    metadataReader,
    converter: async () => { throw new Error('decode failed'); },
    FileCtor: TestFile
  });
  assert.equal(failed.originalFile, original);
  assert.equal(failed.uploadFile, null);
  assert.equal(failed.metadataStatus, 'conversion-failed');
  assert.match(failed.conversionError, /decode failed/);
});

test('keeps JPEG, PNG, and WebP on the existing path without conversion', async () => {
  const { prepareUploadPhoto } = createCoreApi();
  for (const file of [
    { name: 'a.jpg', type: 'image/jpeg' },
    { name: 'b.png', type: 'image/png' },
    { name: 'c.webp', type: 'image/webp' }
  ]) {
    let conversions = 0;
    const result = await prepareUploadPhoto(file, {
      converter: async () => { conversions += 1; },
      readStandardGps: async () => ({ lat: 35, lng: 139 }),
      readStandardDate: async () => '2026-07-10T12:34'
    });
    assert.equal(conversions, 0);
    assert.equal(result.originalFile, file);
    assert.equal(result.uploadFile, file);
    assert.equal(result.capturedAt, '2026-07-10T12:34');
    assert.equal(result.metadataStatus, 'success');
  }
});

test('preserves an invalid GPS classification on the existing image path', async () => {
  const { prepareUploadPhoto } = createCoreApi();
  const file = { name: 'bad-gps.jpg', type: 'image/jpeg' };
  const result = await prepareUploadPhoto(file, {
    readStandardGps: async () => ({ status: 'invalid-gps', lat: null, lng: null }),
    readStandardDate: async () => '2026-07-10T12:34'
  });

  assert.equal(result.uploadFile, file);
  assert.equal(result.metadataStatus, 'invalid-gps');
  assert.equal(result.lat, null);
  assert.equal(result.lng, null);
  assert.equal(result.capturedAt, '2026-07-10T12:34');
});

test('renders distinct location messages for no GPS and metadata failures', () => {
  const { locationMessageForMetadataStatus } = createCoreApi();

  const noGps = locationMessageForMetadataStatus('no-gps');
  const unsupported = locationMessageForMetadataStatus('unsupported');
  const invalid = locationMessageForMetadataStatus('invalid-gps');
  const failed = locationMessageForMetadataStatus('read-failed');
  assert.match(noGps, /GPS 情報がありません/);
  assert.match(unsupported, /解析に対応していません/);
  assert.match(invalid, /不正/);
  assert.match(failed, /読み取れませんでした/);
  assert.equal(new Set([noGps, unsupported, invalid, failed]).size, 4);
});

test('upload UI pins library versions and exposes HEIC conversion status', () => {
  assert.match(indexHtml, /heic-to@1\.5\.2\/dist\/iife\/heic-to\.js/);
  assert.match(indexHtml, /exifreader@4\.41\.0\/dist\/exif-reader\.js/);
  assert.match(indexHtml, /accept="[^"]*image\/heic[^"]*image\/heif[^"]*\.heic[^"]*\.heif/);
  assert.match(indexHtml, /id="upload-photo-status"/);
  assert.match(indexHtml, /HEIC写真を変換しています/);
  assert.match(indexHtml, /位置情報を取得しました/);
});

test('conversion state disables both save actions and conversion failure blocks draft creation', () => {
  const buttons = {
    'upload-submit': { disabled: false },
    'upload-save-unplaced': { disabled: false }
  };
  const context = {
    state: { upload: { converting: true, conversionError: '' } },
    document: { getElementById: (id) => buttons[id] }
  };
  const refreshSource = extractFunction(indexHtml, 'refreshUploadSubmitState');
  vm.runInNewContext(`${refreshSource}; globalThis.__refresh = refreshUploadSubmitState;`, context);
  context.__refresh();
  assert.equal(buttons['upload-submit'].disabled, true);
  assert.equal(buttons['upload-save-unplaced'].disabled, true);

  context.state.upload.converting = false;
  context.state.upload.conversionError = 'decode failed';
  context.__refresh();
  assert.equal(buttons['upload-submit'].disabled, true);
  assert.equal(buttons['upload-save-unplaced'].disabled, true);

  const draftContext = {
    state: { upload: { converting: false, conversionError: 'decode failed', originalFile: {}, uploadFile: null } },
    document: { getElementById: () => ({ value: 'title' }) }
  };
  const draftSource = extractFunction(indexHtml, 'getUploadDraft');
  vm.runInNewContext(`${draftSource}; globalThis.__draft = getUploadDraft;`, draftContext);
  assert.throws(() => draftContext.__draft(), /変換に失敗/);
});

test('save path uses the prepared JPEG and never invokes HEIC conversion', () => {
  const getDraftSource = extractFunction(indexHtml, 'getUploadDraft');
  const saveSource = extractFunction(indexHtml, 'saveNewPin');
  const submitSource = extractFunction(indexHtml, 'handleUploadSubmit');
  const selectionStart = indexHtml.indexOf("document.getElementById('file-input').addEventListener('change'");
  const selectionEnd = indexHtml.indexOf("document.getElementById('upload-title').addEventListener", selectionStart);
  const selectionSource = indexHtml.slice(selectionStart, selectionEnd);

  assert.match(getDraftSource, /file:\s*state\.upload\.uploadFile/);
  assert.match(getDraftSource, /filename:\s*state\.upload\.uploadFile/);
  assert.match(getDraftSource, /state\.upload\.conversionError/);
  assert.match(saveSource, /resizeWithOrientation\(draft\.file,\s*1920\)/);
  assert.match(saveSource, /filename:\s*draft\.filename/);
  assert.doesNotMatch(saveSource, /HeicTo|convertHeic|prepareUploadPhoto/);
  assert.match(submitSource, /openLocationChoice\(locationMessageForMetadataStatus\(state\.upload\.metadataStatus\)/);
  assert.match(selectionSource, /URL\.createObjectURL\(prepared\.uploadFile\)/);
  assert.match(selectionSource, /selectionToken !== state\.upload\.selectionToken/);
});

test('save sends the prepared JPEG filename and GPS payload without reconversion', async () => {
  const preparedJpeg = { name: 'IMG_0001.jpg', type: 'image/jpeg' };
  const calls = [];
  const context = {
    Promise,
    Date,
    state: { pins: [] },
    document: { getElementById: () => ({ disabled: false }) },
    resizeWithOrientation: async (file, maxSize) => {
      assert.equal(file, preparedJpeg);
      assert.equal(maxSize, 1920);
      return 'data:image/jpeg;base64,converted';
    },
    withEditToken: (payload) => payload,
    withGAS: async (method, payload) => {
      calls.push({ method, payload });
      return { ok: true, id: 'pin-1', imageUrl: 'image', fileId: 'file-1', folderUrl: 'folder' };
    },
    currentFolderIdOrRoot: () => 'target-folder',
    clonePin: (pin) => pin,
    cachePinFolderUrl_() {},
    clearUploadPhotoState() {},
    closeOverlay() {},
    renderPins() {},
    renderSidePanel() {},
    updateUnplacedBadge() {},
    renderColorFilterUI() {},
    renderIconFilterUI() {},
    renderTagFilterUI() {}
  };
  const source = extractFunction(indexHtml, 'saveNewPin');
  vm.runInNewContext(`${source}; globalThis.__save = saveNewPin;`, context);
  await context.__save({
    file: preparedJpeg,
    filename: 'IMG_0001.jpg',
    title: 'GPS HEIC', description: '', eventAt: '2026-07-10T12:34',
    color: '#e53935', icon: 'photo', links: [], status: '未対応', tags: []
  }, { lat: 35.25, lng: 139.5 });

  assert.equal(calls.length, 1);
  assert.equal(calls[0].method, 'saveMapData');
  assert.equal(calls[0].payload.filename, 'IMG_0001.jpg');
  assert.equal(calls[0].payload.lat, 35.25);
  assert.equal(calls[0].payload.lng, 139.5);
  assert.match(calls[0].payload.base64, /^data:image\/jpeg/);
});

test('clearing a photo revokes the preview URL and resets every image-derived field', () => {
  const source = extractFunction(indexHtml, 'clearUploadPhotoState');
  const revoked = [];
  const elements = {
    'file-input': { value: 'selected' },
    'upload-preview': { src: 'blob:preview', style: { display: 'block' } },
    'upload-event-at': { value: '2026-07-10T12:34' },
    'file-drop': { textContent: '', classList: { remove() {} } },
    'upload-photo-status': { textContent: 'old', className: 'old', style: {} },
    'upload-submit': { disabled: true },
    'upload-save-unplaced': { disabled: true }
  };
  const context = {
    state: {
      upload: {
        originalFile: { name: 'source.heic' }, uploadFile: { name: 'source.jpg' },
        previewUrl: 'blob:preview', lat: 35, lng: 139, capturedAt: '2026-07-10T12:34',
        metadataStatus: 'success', converting: true, conversionError: 'old', selectionToken: 4
      }
    },
    URL: { revokeObjectURL: (value) => revoked.push(value) },
    document: { getElementById: (id) => elements[id] },
    refreshUploadPhotoStatus() {},
    refreshUploadSubmitState() {},
    refreshRenameNotes() {}
  };
  vm.runInNewContext(`${source}; globalThis.clearUploadPhotoState();`, context);

  assert.deepEqual(revoked, ['blob:preview']);
  assert.equal(context.state.upload.originalFile, null);
  assert.equal(context.state.upload.uploadFile, null);
  assert.equal(context.state.upload.previewUrl, '');
  assert.equal(context.state.upload.lat, null);
  assert.equal(context.state.upload.lng, null);
  assert.equal(context.state.upload.capturedAt, '');
  assert.equal(context.state.upload.metadataStatus, 'idle');
  assert.equal(context.state.upload.converting, false);
  assert.equal(context.state.upload.conversionError, '');
  assert.equal(context.state.upload.selectionToken, 5);
  assert.equal(elements['upload-preview'].src, '');
  assert.equal(elements['upload-preview'].style.display, 'none');
});

test('backdrop dismissal releases upload preview resources', () => {
  const source = extractFunction(indexHtml, 'closeOverlayFromBackdrop');
  assert.match(source, /id === 'upload-overlay'[\s\S]*clearUploadPhotoState\(\)/);
});
