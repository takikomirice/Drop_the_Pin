const assert = require('node:assert/strict');
const test = require('node:test');
const {
  makeHarness, pinRow, validMp3Bytes, MAP_HEADERS, RECEIPT_HEADERS
} = require('./audio-storage-harness');

const SHARE_HEADERS = [
  'createdAt', 'label', 'token', 'tags', 'tagMode', 'enabled', 'revokedAt',
  'colors', 'routeIds', 'routeTargetsJson'
];
const ROUTE_HEADERS = [
  'routeId', 'name', 'color', 'routeMode', 'closed', 'startPinId', 'endPinId',
  'createdAt', 'updatedAt', 'orderIndex', 'visible', 'showNumbers', 'showLine', 'lineStyle'
];
const ROUTE_PIN_HEADERS = ['routeId', 'pinId', 'pinOrder', 'createdAt', 'updatedAt'];

function shareRow(token, overrides = {}) {
  const values = {
    label: token, tags: '', tagMode: 'or', enabled: true, revokedAt: '',
    colors: '', routeIds: '', routeTargetsJson: '', ...overrides
  };
  return [
    '2026-07-22T00:00:00.000Z', values.label, token, values.tags, values.tagMode,
    values.enabled, values.revokedAt, values.colors, values.routeIds,
    values.routeTargetsJson
  ];
}

function createProjectionHarness() {
  const sheets = {
    map_info: [
      MAP_HEADERS,
      pinRow({ id: 'pin-route', title: 'Route', tags: 'alpha', color: '#4caf50', audioId: 'audio_route_123456' }),
      pinRow({ id: 'pin-orphan', title: 'Orphan', tags: 'alpha', color: '#4caf50', audioId: 'audio_orphan_12345' }),
      pinRow({ id: 'pin-beta', title: 'Beta', tags: 'beta', color: '#4caf50', audioId: 'audio_beta_1234567' }),
      pinRow({ id: 'pin-no-audio', title: 'Silent', tags: 'alpha', color: '#4caf50', audioId: '' })
    ],
    import_receipts: [RECEIPT_HEADERS],
    share_links: [
      SHARE_HEADERS,
      shareRow('share-alpha', { tags: 'alpha', colors: '#4caf50', routeIds: 'route-a' }),
      shareRow('share-beta', { tags: 'beta', colors: '#4caf50' }),
      shareRow('share-disabled', { tags: 'alpha', enabled: false }),
      shareRow('share-revoked', { tags: 'alpha', enabled: true, revokedAt: '2026-07-22T01:00:00.000Z' })
    ],
    routes: [
      ROUTE_HEADERS,
      ['route-a', 'Route A', '#1e88e5', 'straight', false, '', '', '', '', 0, true, true, true, 'solid']
    ],
    route_pins: [ROUTE_PIN_HEADERS, ['route-a', 'pin-route', 0, '', '']]
  };
  const harness = makeHarness({ sheets });
  for (const [id, name] of [
    ['audio_route_123456', 'route.mp3'],
    ['audio_orphan_12345', 'orphan.mp3'],
    ['audio_beta_1234567', 'beta.mp3']
  ]) {
    harness.addFile(id, name, 'audio/mpeg', validMp3Bytes(), [harness.audioFolder.id]);
  }
  return harness;
}

function errorFingerprint(callback) {
  try {
    callback();
    assert.fail('expected shared audio access to fail');
  } catch (error) {
    if (error && error.code === 'ERR_ASSERTION') throw error;
    return {
      name: String(error && error.name || ''),
      code: String(error && error.code || ''),
      message: String(error && error.message || ''),
      retryable: error && error.retryable === true
    };
  }
}

test('shared audio succeeds only for pins in the exact shared view projection', () => {
  const harness = createProjectionHarness();
  assert.equal(typeof harness.api.resolveSharedProjection_, 'function');
  assert.equal(typeof harness.api.getSharedPinAudioData, 'function');

  const shared = harness.api.getSharedViewData('share-alpha');
  assert.equal(shared.ok, true);
  assert.deepEqual(
    JSON.parse(JSON.stringify(shared.pins.map((pin) => pin.id))),
    ['pin-route', 'pin-orphan', 'pin-no-audio'],
    'route selection must keep the existing shared pin projection semantics'
  );

  for (const pinId of ['pin-route', 'pin-orphan']) {
    const result = harness.api.getSharedPinAudioData({ shareToken: 'share-alpha', pinId });
    assert.deepEqual(Object.keys(result).sort(), ['base64', 'byteLength', 'mimeType', 'ok']);
    assert.equal(result.ok, true);
    assert.equal(result.mimeType, 'audio/mpeg');
    assert.equal(result.byteLength, 1024 * 1024);
    assert.equal(Buffer.from(result.base64, 'base64').length, result.byteLength);
    assert.equal(JSON.stringify(result).includes('audio_'), false);
    for (const field of ['audioId', 'fileId', 'name', 'folder', 'url', 'sourceDriveFileId']) {
      assert.equal(Object.hasOwn(result, field), false);
    }
  }

  assert.equal(harness.api.getSharedPinAudioData({
    shareToken: 'share-beta', pinId: 'pin-beta'
  }).ok, true);
});

test('all denied shared audio cases use one safe indistinguishable error contract', () => {
  const harness = createProjectionHarness();
  assert.equal(typeof harness.api.getSharedPinAudioData, 'function');
  const attempts = [
    () => harness.api.getSharedPinAudioData({ shareToken: 'missing', pinId: 'pin-route' }),
    () => harness.api.getSharedPinAudioData({ shareToken: 'share-disabled', pinId: 'pin-route' }),
    () => harness.api.getSharedPinAudioData({ shareToken: 'share-revoked', pinId: 'pin-route' }),
    () => harness.api.getSharedPinAudioData({ shareToken: 'share-alpha', pinId: 'pin-beta' }),
    () => harness.api.getSharedPinAudioData({ shareToken: 'share-alpha', pinId: 'pin-no-audio' }),
    () => harness.api.getSharedPinAudioData({ shareToken: 'share-alpha', pinId: 'deleted-pin' }),
    () => harness.api.getSharedPinAudioData({ shareToken: 'share-alpha', pinId: '=formula' }),
    () => harness.api.getSharedPinAudioData({ shareToken: 'share-alpha', audioId: 'audio_route_123456' })
  ];
  const fingerprints = attempts.map(errorFingerprint);
  fingerprints.forEach((fingerprint) => assert.deepEqual(fingerprint, fingerprints[0]));
  assert.equal(fingerprints[0].retryable, false);
  assert.equal(JSON.stringify(fingerprints[0]).includes('share-alpha'), false);
  assert.equal(JSON.stringify(fingerprints[0]).includes('audio_route_123456'), false);
});

test('malformed or unmanaged shared audio fails safely without leaking storage details', () => {
  function malformed(overrides, file) {
    const harness = makeHarness({
      sheets: {
        map_info: [MAP_HEADERS, pinRow({
          id: 'pin-bad-audio', tags: 'alpha', color: '#4caf50', audioId: file.id,
          ...overrides
        })],
        import_receipts: [RECEIPT_HEADERS],
        share_links: [SHARE_HEADERS, shareRow('share-bad', { tags: 'alpha' })]
      }
    });
    harness.addFile(file.id, file.name, file.mimeType, file.bytes, file.parents(harness));
    return harness;
  }

  const cases = [
    malformed({}, {
      id: 'audio_short_123456', name: 'short.mp3', mimeType: 'audio/mpeg',
      bytes: [0x49, 0x44, 0x33], parents: (harness) => [harness.audioFolder.id]
    }),
    malformed({}, {
      id: 'audio_wav_12345678', name: 'wrong.wav', mimeType: 'audio/wav',
      bytes: validMp3Bytes(), parents: (harness) => [harness.audioFolder.id]
    }),
    malformed({}, {
      id: 'audio_unmanaged_123', name: 'outside.mp3', mimeType: 'audio/mpeg',
      bytes: validMp3Bytes(), parents: (harness) => [harness.rootId]
    }),
    malformed({}, {
      id: 'audio_too_large_123', name: 'large.mp3', mimeType: 'audio/mpeg',
      bytes: validMp3Bytes(4 * 1024 * 1024 + 1), parents: (harness) => [harness.audioFolder.id]
    })
  ];
  const fingerprints = cases.map((harness) => errorFingerprint(() =>
    harness.api.getSharedPinAudioData({ shareToken: 'share-bad', pinId: 'pin-bad-audio' })
  ));
  fingerprints.forEach((fingerprint) => assert.deepEqual(fingerprint, fingerprints[0]));
  assert.equal(JSON.stringify(fingerprints[0]).includes('audio_'), false);
});

test('shared audio reads do not create folders, mutate sheets, or change Drive sharing', () => {
  const harness = createProjectionHarness();
  const before = {
    sheetWrites: harness.audit.sheetWrites,
    lockAttempts: harness.audit.locks.attempts,
    drive: { ...harness.audit.drive },
    events: harness.audit.events.length
  };

  assert.equal(harness.api.getSharedPinAudioData({
    shareToken: 'share-alpha', pinId: 'pin-route'
  }).ok, true);

  assert.equal(harness.audit.sheetWrites, before.sheetWrites);
  assert.equal(harness.audit.locks.attempts, before.lockAttempts);
  assert.deepEqual(harness.audit.drive, before.drive);
  assert.equal(harness.audit.events.length, before.events);
});
