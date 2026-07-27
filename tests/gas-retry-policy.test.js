const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

const READ_METHODS = [
  'getAppSettings',
  'getMapData',
  'getPinAudioData',
  'getPinPhotoData',
  'getPinDriveMeta',
  'getRootFolderContents',
  'getRouteCache',
  'getRouteGroups',
  'getTracks',
  'listDriveMediaInbox',
  'listDrivePhotoImportFolder',
  'listInputPresets',
  'listShareLinks',
  'navigateToFolder',
  'readDrivePhotoImportFile',
  'readDriveAudioImportFile'
];

const MUTATION_METHODS = [
  'bulkDeletePins',
  'bulkUpdatePinMetadata',
  'bulkUpdatePinStatus',
  'createShareLink',
  'deleteInputPreset',
  'deletePin',
  'deleteRouteGroup',
  'deleteShareLink',
  'deleteTrack',
  'duplicatePin',
  'ensureMediaDriveStructure',
  'movePin',
  'putRouteCache',
  'removePinAudio',
  'saveImportPhotoItem',
  'saveImportAudioItem',
  'saveImportPinItem',
  'saveInputPreset',
  'saveMapData',
  'saveRouteGroup',
  'saveTrackBundle',
  'setRoutePins',
  'setShareLinkEnabled',
  'unplacePin',
  'updateAppSettings',
  'updateInputPresetOrder',
  'updatePinDetails',
  'updateTrackDisplaySettings',
  'updateRoutesOrder',
  'updateTracksOrder'
];

function retryPolicySource() {
  const start = indexHtml.indexOf('    const GAS_RETRY_MAX');
  const end = indexHtml.indexOf('    function withEditToken', start);
  assert.notEqual(start, -1, 'Expected GAS retry constants');
  assert.notEqual(end, -1, 'Expected withEditToken after GAS helpers');
  return indexHtml.slice(start, end);
}

function createHarness(outcomes) {
  const queue = Array.isArray(outcomes) ? outcomes.slice() : [];
  const calls = [];
  const delays = [];
  let successHandler = null;
  let failureHandler = null;
  let runner;

  runner = new Proxy({
    withSuccessHandler(handler) {
      successHandler = handler;
      return runner;
    },
    withFailureHandler(handler) {
      failureHandler = handler;
      return runner;
    }
  }, {
    get(target, property) {
      if (Object.prototype.hasOwnProperty.call(target, property)) return target[property];
      return function(...args) {
        calls.push({ method: String(property), args });
        const outcome = queue.length ? queue.shift() : { ok: true, value: { ok: true } };
        if (outcome.ok) {
          successHandler(outcome.value);
        } else {
          failureHandler(outcome.error);
        }
      };
    }
  });

  const deterministicMath = Object.create(Math);
  deterministicMath.random = () => 0;
  const context = {
    google: { script: { run: runner } },
    Math: deterministicMath,
    setTimeout(handler, delay) {
      delays.push(delay);
      handler();
    }
  };
  vm.runInNewContext(`${retryPolicySource()}
globalThis.__api = {
  callGAS,
  withGAS,
  withGASNoArg,
  isRetryableGASReadMethod: typeof isRetryableGASReadMethod === 'function'
    ? isRetryableGASReadMethod : null,
  readMethods: typeof GAS_RETRY_READ_METHODS === 'undefined'
    ? [] : Object.keys(GAS_RETRY_READ_METHODS).sort()
};`, context);

  return { api: context.__api, calls, delays };
}

function failure(message) {
  return { ok: false, error: new Error(message) };
}

function success(value) {
  return { ok: true, value };
}

test('read allowlist exactly covers side-effect-free index call sites', () => {
  const { api } = createHarness();
  assert.equal(typeof api.isRetryableGASReadMethod, 'function');
  assert.deepEqual(Array.from(api.readMethods), READ_METHODS.slice().sort());
  for (const method of READ_METHODS) assert.equal(api.isRetryableGASReadMethod(method), true, method);
  for (const method of MUTATION_METHODS) assert.equal(api.isRetryableGASReadMethod(method), false, method);
  assert.equal(api.isRetryableGASReadMethod('futureUnknownMethod'), false);

  const calledMethods = new Set(Array.from(indexHtml.matchAll(
    /\b(?:withGAS(?:NoArg)?|callGAS|gasCall|config\.callGAS)\('([^']+)'/g
  ), (match) => match[1]));
  assert.deepEqual(
    Array.from(calledMethods).sort(),
    READ_METHODS.concat(MUTATION_METHODS).sort(),
    'Every index call site must be classified as read or mutation'
  );
});

test('read API retries transient failures with the existing backoff and identical payload', async () => {
  const payload = Object.freeze({ __editToken: 'edit-token', folderId: 'folder-1' });
  const expected = { ok: true, items: [] };
  const harness = createHarness([
    failure('temporary-1'),
    failure('temporary-2'),
    success(expected)
  ]);

  const result = await harness.api.withGAS('listDrivePhotoImportFolder', payload);

  assert.deepEqual(result, expected);
  assert.deepEqual(harness.delays, [1000, 2000]);
  assert.equal(harness.calls.length, 3);
  for (const call of harness.calls) {
    assert.equal(call.method, 'listDrivePhotoImportFolder');
    assert.equal(call.args.length, 1);
    assert.equal(call.args[0], payload);
    assert.equal(call.args[0].__editToken, 'edit-token');
  }
});

test('read API keeps the existing maximum of four retries', async () => {
  const errors = [0, 1, 2, 3, 4].map((index) => failure(`temporary-${index}`));
  const harness = createHarness(errors);

  await assert.rejects(harness.api.withGASNoArg('getMapData'), /temporary-4/);

  assert.equal(harness.calls.length, 5);
  assert.deepEqual(harness.calls.map((call) => call.args.length), [0, 0, 0, 0, 0]);
  assert.deepEqual(harness.delays, [1000, 2000, 4000, 8000]);
});

test('every known mutation API executes only once after transport failure', async () => {
  for (const method of MUTATION_METHODS) {
    const payload = Object.freeze({ __editToken: 'same-token', method });
    const harness = createHarness([failure(`${method}-failed`)]);

    await assert.rejects(harness.api.withGAS(method, payload), new RegExp(`${method}-failed`));

    assert.equal(harness.calls.length, 1, method);
    assert.equal(harness.calls[0].args[0], payload, method);
    assert.equal(harness.calls[0].args[0].__editToken, 'same-token', method);
    assert.deepEqual(harness.delays, [], method);
  }
});

test('unknown API executes only once after transport failure', async () => {
  const payload = Object.freeze({ __editToken: 'same-token', value: 1 });
  const harness = createHarness([failure('unknown-failed')]);

  await assert.rejects(harness.api.withGAS('futureUnknownMethod', payload), /unknown-failed/);

  assert.equal(harness.calls.length, 1);
  assert.equal(harness.calls[0].args[0], payload);
  assert.deepEqual(harness.delays, []);
});

test('successful read and mutation APIs both execute once without changing arguments', async () => {
  const readHarness = createHarness([success({ ok: true, presets: [] })]);
  const readPayload = Object.freeze({ __editToken: 'read-token' });
  await readHarness.api.withGAS('listInputPresets', readPayload);
  assert.equal(readHarness.calls.length, 1);
  assert.equal(readHarness.calls[0].args[0], readPayload);
  assert.deepEqual(readHarness.delays, []);

  const mutationHarness = createHarness([success({ ok: true })]);
  const mutationPayload = Object.freeze({ __editToken: 'mutation-token', title: 'unchanged' });
  await mutationHarness.api.withGAS('saveMapData', mutationPayload);
  assert.equal(mutationHarness.calls.length, 1);
  assert.equal(mutationHarness.calls[0].args[0], mutationPayload);
  assert.deepEqual(mutationHarness.delays, []);
});
