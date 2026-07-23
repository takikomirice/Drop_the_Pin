const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.resolve(__dirname, '..', 'index.html'), 'utf8');

class BrowserFile extends Blob {
  constructor(parts, name, options = {}) {
    super(parts, options);
    this.name = name;
    this.lastModified = options.lastModified || 0;
  }
}

function loadProcessor() {
  const start = indexHtml.indexOf('const ImportAudioItemProcessor =');
  const end = indexHtml.indexOf('const AudioPinImportWorkflow =', start);
  assert.notEqual(start, -1, 'ImportAudioItemProcessor must exist');
  assert.notEqual(end, -1, 'AudioPinImportWorkflow boundary must exist');
  const context = { console, Blob, Uint8Array, ArrayBuffer, Promise };
  vm.runInNewContext(`${indexHtml.slice(start, end)}\nglobalThis.__api = ImportAudioItemProcessor;`, context);
  return { api: context.__api, source: indexHtml.slice(start, end) };
}

function mp3Result(overrides = {}) {
  const blob = new Blob([new Uint8Array([0x49, 0x44, 0x33, 4, 0, 0, 0, 0, 0, 0])], { type: 'audio/mpeg' });
  return {
    blob, fileName: 'voice-trimmed.mp3', mimeType: 'audio/mpeg', sizeBytes: blob.size,
    sourceDurationSeconds: 45, durationSeconds: 30,
    selectionStart: 0, selectionEnd: 30,
    sampleRate: 48000, bitrate: 192000, numberOfChannels: 2,
    ...overrides
  };
}

function localFile(size, name = 'voice.m4a', type = 'audio/mp4') {
  return { name, type, size, lastModified: 123 };
}

function deferred() {
  let resolve;
  let reject;
  const promise = new Promise((resolvePromise, rejectPromise) => {
    resolve = resolvePromise;
    reject = rejectPromise;
  });
  return { promise, resolve, reject };
}

test('local 200MB maximum is enforced before editor/vendor handoff', async () => {
  const { api } = loadProcessor();
  let editorCalls = 0;
  const processor = api.create({
    openAudioEditor: async () => { editorCalls += 1; return mp3Result(); },
    confirmLargeLocal: async () => true
  });
  await assert.rejects(
    processor.processLocalFile(localFile(200 * 1024 * 1024 + 1)),
    (error) => error && error.code === 'AUDIO_FILE_TOO_LARGE'
  );
  assert.equal(editorCalls, 0);
});

test('local files over 50MB continue only after explicit confirmation', async () => {
  const { api } = loadProcessor();
  const calls = [];
  const rejected = api.create({
    confirmLargeLocal: async (file) => { calls.push(['confirm', file.name]); return false; },
    openAudioEditor: async () => { calls.push(['editor']); return mp3Result(); }
  });
  await assert.rejects(
    rejected.processLocalFile(localFile(50 * 1024 * 1024 + 1)),
    (error) => error && error.code === 'AUDIO_LARGE_FILE_CANCELLED'
  );
  assert.deepEqual(calls, [['confirm', 'voice.m4a']]);

  const accepted = api.create({
    confirmLargeLocal: async () => true,
    openAudioEditor: async (file) => { calls.push(['editor', file.name]); return mp3Result(); }
  });
  const result = await accepted.processLocalFile(localFile(50 * 1024 * 1024 + 1));
  assert.equal(result.mimeType, 'audio/mpeg');
  assert.deepEqual(calls.at(-1), ['editor', 'voice.m4a']);
});

test('local browser MIME aliases and empty MIME use the editor-supported extension contract', async () => {
  const { api } = loadProcessor();
  const accepted = [
    ['voice.m4a', 'audio/mp4'], ['voice.m4a', 'audio/x-m4a'], ['voice.m4a', 'video/mp4'],
    ['voice.mp3', 'audio/mpeg'], ['voice.mp3', 'audio/mp3'],
    ['voice.wav', 'audio/wav'], ['voice.wav', 'audio/x-wav'], ['voice.wav', 'audio/wave'],
    ['voice.wav', 'audio/vnd.wave'],
    ['voice.m4a', ''], ['voice.mp3', ''], ['voice.wav', '']
  ];
  const seen = [];
  const processor = api.create({
    openAudioEditor: async (file) => { seen.push([file.name, file.type]); return mp3Result(); }
  });

  for (const [name, type] of accepted) {
    const output = await processor.processLocalFile(localFile(1024, name, type));
    assert.equal(output.mimeType, 'audio/mpeg', `${name} ${type || '(empty)'}`);
    processor.release();
  }
  assert.deepEqual(seen, accepted);

  await assert.rejects(
    processor.processLocalFile(localFile(1024, 'voice.m4a', 'audio/mpeg')),
    (error) => error && error.code === 'AUDIO_FILE_TYPE_UNSUPPORTED'
  );
});

test('Drive response validates requested id, basename, MIME, canonical Base64, and decoded size before File handoff', async () => {
  const { api } = loadProcessor();
  const seen = [];
  const bytes = Buffer.from([1, 2, 3, 4]);
  const processor = api.create({
    environment: {
      atob: (value) => Buffer.from(value, 'base64').toString('binary'),
      Uint8Array, Blob, File: BrowserFile
    },
    openAudioEditor: async (file) => { seen.push(file); return mp3Result(); }
  });
  const response = {
    ok: true,
    file: {
      id: 'audio_AAAAAAAAAAA', name: 'field.wav', mimeType: 'audio/wav',
      sizeBytes: bytes.length, modifiedAt: '2026-07-22T10:20:30.000Z',
      kind: 'wav', base64: bytes.toString('base64')
    }
  };
  const result = await processor.processDriveResponse(response, {
    requestFileId: 'audio_AAAAAAAAAAA'
  });
  assert.equal(result.mimeType, 'audio/mpeg');
  assert.equal(seen.length, 1);
  assert.equal(seen[0] instanceof BrowserFile, true);
  assert.equal(seen[0].name, 'field.wav');
  assert.equal(seen[0].type, 'audio/wav');
  assert.equal(seen[0].size, bytes.length);
  assert.equal(Object.hasOwn(seen[0], 'id'), false);

  const invalids = [
    { file: { ...response.file, id: 'audio_BBBBBBBBBBB' } },
    { file: { ...response.file, name: '../field.wav' } },
    { file: { ...response.file, mimeType: 'audio/mpeg' } },
    { file: { ...response.file, base64: 'A===' } },
    { file: { ...response.file, sizeBytes: bytes.length + 1 } }
  ];
  const strictDriveAliases = [
    { file: { ...response.file, name: 'field.m4a', mimeType: 'audio/x-m4a', kind: 'm4a' } },
    { file: { ...response.file, name: 'field.mp3', mimeType: 'audio/mp3', kind: 'mp3' } },
    { file: { ...response.file, mimeType: 'audio/x-wav' } }
  ];
  invalids.push(...strictDriveAliases);
  for (const invalid of invalids) {
    await assert.rejects(
      processor.processDriveResponse({ ok: true, ...invalid }, { requestFileId: 'audio_AAAAAAAAAAA' }),
      (error) => error && error.code === 'DRIVE_AUDIO_RESPONSE_INVALID'
    );
  }
  assert.equal(seen.length, 1);
});

test('editor public result is independently checked for duration, selection, channels, and MP3 output', async () => {
  const { api, source } = loadProcessor();
  assert.doesNotMatch(source, /EXIF|heic|readExifGps|readAudioMetadata/);
  const invalidResults = [
    mp3Result({ sourceDurationSeconds: 601 }),
    mp3Result({ selectionEnd: 0.4, durationSeconds: 0.4 }),
    mp3Result({ selectionEnd: 121, durationSeconds: 121 }),
    mp3Result({ numberOfChannels: 3 }),
    mp3Result({ sampleRate: 44100 }),
    mp3Result({ bitrate: 128000 }),
    mp3Result({ mimeType: 'audio/wav' })
  ];
  for (const editorResult of invalidResults) {
    const processor = api.create({ openAudioEditor: async () => editorResult });
    await assert.rejects(
      processor.processLocalFile(localFile(1024)),
      (error) => error && error.code === 'AUDIO_EDITOR_RESULT_INVALID'
    );
  }
});

test('cancel and release delegate resource ownership without double release', async () => {
  const { api } = loadProcessor();
  let releases = 0;
  const processor = api.create({
    openAudioEditor: async () => mp3Result(),
    releaseEditorResult: () => { releases += 1; }
  });
  await processor.processLocalFile(localFile(1024));
  assert.equal(processor.release(), true);
  assert.equal(processor.release(), false);
  assert.equal(releases, 1);
});

test('cancel during asynchronous MP3 validation rejects the stale editor result before adoption', async () => {
  const { api } = loadProcessor();
  const headerGate = deferred();
  const headerStarted = deferred();
  class DelayedHeaderBlob extends Blob {
    slice() {
      return {
        arrayBuffer() {
          headerStarted.resolve();
          return headerGate.promise;
        }
      };
    }
  }
  const blob = new DelayedHeaderBlob(
    [new Uint8Array([0x49, 0x44, 0x33, 4, 0, 0, 0, 0, 0, 0])],
    { type: 'audio/mpeg' }
  );
  const processor = api.create({
    openAudioEditor: async () => mp3Result({ blob, sizeBytes: blob.size })
  });

  const processing = processor.processLocalFile(localFile(1024));
  await headerStarted.promise;
  processor.cancel();
  headerGate.resolve(Uint8Array.from([0x49, 0x44, 0x33, 4]).buffer);

  await assert.rejects(processing, (error) => error && error.code === 'AUDIO_IMPORT_CANCELLED');
  assert.equal(processor.getResult(), null);
});

test('cancel during long-duration confirmation rejects the stale result after the awaited choice', async () => {
  const { api } = loadProcessor();
  const confirmationGate = deferred();
  const confirmationStarted = deferred();
  const processor = api.create({
    openAudioEditor: async () => mp3Result({ sourceDurationSeconds: 301 }),
    confirmLongDuration() {
      confirmationStarted.resolve();
      return confirmationGate.promise;
    }
  });

  const processing = processor.processLocalFile(localFile(1024));
  await confirmationStarted.promise;
  processor.cancel();
  confirmationGate.resolve(true);

  await assert.rejects(processing, (error) => error && error.code === 'AUDIO_IMPORT_CANCELLED');
  assert.equal(processor.getResult(), null);
});
