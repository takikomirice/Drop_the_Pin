const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.resolve(__dirname, '..', 'index.html'), 'utf8');

function loadShell() {
  const start = indexHtml.indexOf('const MediaImportShell =');
  const end = indexHtml.indexOf('const ImportAudioItemProcessor =', start);
  assert.notEqual(start, -1, 'MediaImportShell must exist');
  assert.notEqual(end, -1, 'ImportAudioItemProcessor boundary must exist');
  const context = { console };
  vm.runInNewContext(`${indexHtml.slice(start, end)}\nglobalThis.__api = MediaImportShell;`, context);
  return { api: context.__api, source: indexHtml.slice(start, end) };
}

test('shared shell owns generic source, mode, identity, retry, cancel, and return-target state only', () => {
  const { api, source } = loadShell();
  const ids = ['job-shared', 'item-shared'];
  const shell = api.create({ createId: () => ids.shift() });
  const operation = shell.begin({
    mediaKind: 'audio',
    sourceKind: 'local',
    operationMode: 'create-pin',
    selectionLimit: 1,
    returnOverlayId: 'add-menu-overlay'
  });

  assert.deepEqual(JSON.parse(JSON.stringify(operation)), {
    mediaKind: 'audio', sourceKind: 'local', operationMode: 'create-pin',
    selectionLimit: 1, returnOverlayId: 'add-menu-overlay',
    targetPinId: '', expectedUpdatedAt: '',
    jobId: 'job-shared', itemId: 'item-shared',
    idempotencyKey: 'job-shared:item-shared', sourceDriveFileId: '', status: 'active'
  });
  assert.doesNotMatch(source, /EXIF|heic|decodeAudioData|Mp3OutputFormat|50\s*\*\s*1024|200\s*\*\s*1024/);
  assert.doesNotMatch(source, /mediaKind\s*===\s*['"](?:audio|photo)['"]/);
});

test('cancel is idempotent and retry preserves identity while invoking generic hooks', async () => {
  const { api } = loadShell();
  const calls = [];
  const shell = api.create({ createId: (() => { let id = 0; return () => `shared-${++id}`; })() });
  const first = shell.begin({
    mediaKind: 'photo', sourceKind: 'drive', operationMode: 'create-pin',
    selectionLimit: 20, returnOverlayId: 'add-menu-overlay',
    sourceDriveFileId: 'drive_AAAAAAAAAAA',
    onCancel: (operation, reason) => calls.push(['cancel', operation.idempotencyKey, reason]),
    onRetry: (operation) => calls.push(['retry', operation.idempotencyKey])
  });

  assert.equal(shell.cancel('user'), true);
  assert.equal(shell.cancel('again'), false);
  const retried = await shell.retry();
  assert.equal(retried.idempotencyKey, first.idempotencyKey);
  assert.equal(retried.status, 'active');
  assert.deepEqual(calls, [
    ['cancel', first.idempotencyKey, 'user'],
    ['retry', first.idempotencyKey]
  ]);
  assert.equal(shell.getReturnTarget(), 'add-menu-overlay');
});

test('shell rejects unsupported source, mode, and selection limits before creating an operation', () => {
  const { api } = loadShell();
  const shell = api.create({ createId: () => 'unused' });
  assert.throws(() => shell.begin({ mediaKind: 'video', sourceKind: 'local', operationMode: 'create-pin', selectionLimit: 1 }), /media/i);
  assert.throws(() => shell.begin({ mediaKind: 'audio', sourceKind: 'cloud', operationMode: 'create-pin', selectionLimit: 1 }), /source/i);
  assert.throws(() => shell.begin({ mediaKind: 'audio', sourceKind: 'local', operationMode: 'delete', selectionLimit: 1 }), /mode/i);
  assert.throws(() => shell.begin({ mediaKind: 'audio', sourceKind: 'local', operationMode: 'create-pin', selectionLimit: 21 }), /selection/i);
});

test('default shell supports the documented MediaImportShell.begin entry point', () => {
  const { api } = loadShell();
  const operation = api.begin({
    mediaKind: 'audio', sourceKind: 'local', operationMode: 'create-pin',
    selectionLimit: 1, returnOverlayId: 'add-menu-overlay'
  });
  assert.match(operation.jobId, /^media-/);
  assert.equal(operation.idempotencyKey, `${operation.jobId}:${operation.itemId}`);
  api.cancel('test-cleanup');
});
