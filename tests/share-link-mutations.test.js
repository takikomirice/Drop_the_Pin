const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const codeJs = fs.readFileSync(path.join(root, 'Code.js'), 'utf8');

const EDIT_TOKEN = 'share-mutation-edit-token';
const LEGACY_HEADERS = [
  'createdAt', 'label', 'token', 'tags', 'tagMode', 'enabled', 'revokedAt', 'colors', 'routeIds'
];

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

function shareRow(headers, values = {}) {
  return headers.map((header) => Object.prototype.hasOwnProperty.call(values, header) ? values[header] : '');
}

function loadApi({ headers = LEGACY_HEADERS, dataRows = [], authorized = true, lockAcquired = true, enforceLock = false } = {}) {
  const audit = {
    authChecks: 0,
    spreadsheetAccesses: 0,
    reads: [],
    writes: [],
    deletes: [],
    lockAttempts: 0,
    lockReleases: 0,
    lockHeld: false,
    events: []
  };
  const rows = [headers.slice()].concat(dataRows.map((row) => row.slice()));

  function assertLockHeld(operation) {
    if (enforceLock) assert.equal(audit.lockHeld, true, `${operation} must run while the mutation lock is held`);
  }

  const sheet = {
    rows,
    getDataRange() {
      assertLockHeld('getDataRange');
      audit.events.push('read');
      audit.reads.push({ lockHeld: audit.lockHeld });
      return { getValues: () => this.rows.map((row) => row.slice()) };
    },
    getRange(row, column) {
      return {
        setValue: (value) => {
          assertLockHeld('setValue');
          audit.events.push('write');
          audit.writes.push({ row, column, value, lockHeld: audit.lockHeld });
          this.rows[row - 1][column - 1] = value;
          return this;
        }
      };
    },
    deleteRow(row) {
      assertLockHeld('deleteRow');
      audit.events.push('delete');
      audit.deletes.push({ row, lockHeld: audit.lockHeld });
      this.rows.splice(row - 1, 1);
    }
  };
  const spreadsheet = {
    getSheetByName(name) {
      assertLockHeld('getSheetByName');
      assert.equal(name, 'share_links');
      return sheet;
    }
  };
  const lock = {
    tryLock() {
      audit.events.push('lock');
      audit.lockAttempts += 1;
      if (lockAcquired) audit.lockHeld = true;
      return lockAcquired;
    },
    releaseLock() {
      assert.equal(audit.lockHeld, true);
      audit.events.push('unlock');
      audit.lockReleases += 1;
      audit.lockHeld = false;
    }
  };
  const context = {
    CacheService: {
      getScriptCache: () => ({
        get(key) {
          audit.events.push('auth');
          audit.authChecks += 1;
          return authorized && key === `EDIT_TOKEN_${EDIT_TOKEN}` ? '1' : null;
        }
      })
    },
    LockService: { getScriptLock: () => lock },
    SpreadsheetApp: {
      getActiveSpreadsheet() {
        audit.events.push('spreadsheet');
        audit.spreadsheetAccesses += 1;
        assert.ok(audit.authChecks > 0, 'edit token must be checked before Spreadsheet access');
        assertLockHeld('getActiveSpreadsheet');
        return spreadsheet;
      }
    }
  };
  vm.runInNewContext(`${codeJs}
globalThis.__shareMutationApi = { setShareLinkEnabled, deleteShareLink, revokeShareLink };`, context);
  return { api: context.__shareMutationApi, audit, sheet };
}

function authorized(payload = {}) {
  return Object.assign({}, payload, { __editToken: EDIT_TOKEN });
}

test('legacy headers enable and disable the requested link and maintain revokedAt', () => {
  const initialRevokedAt = '2026-01-01T00:00:00.000Z';
  const { api, audit, sheet } = loadApi({
    dataRows: [shareRow(LEGACY_HEADERS, {
      label: 'Legacy', token: 'legacy-token', enabled: false, revokedAt: initialRevokedAt,
      routeIds: 'route-a'
    })]
  });

  assert.deepEqual(plain(api.setShareLinkEnabled(authorized({ token: 'legacy-token', enabled: true }))), { ok: true });
  assert.equal(sheet.rows[1][LEGACY_HEADERS.indexOf('enabled')], true);
  assert.equal(sheet.rows[1][LEGACY_HEADERS.indexOf('revokedAt')], '');

  assert.deepEqual(plain(api.revokeShareLink(authorized({ token: 'legacy-token' }))), { ok: true });
  assert.equal(sheet.rows[1][LEGACY_HEADERS.indexOf('enabled')], false);
  assert.ok(Number.isFinite(Date.parse(sheet.rows[1][LEGACY_HEADERS.indexOf('revokedAt')])));
  assert.deepEqual(audit.writes.map(({ row, column }) => [row, column]), [[2, 6], [2, 7], [2, 6], [2, 7]]);
});

test('reordered headers and inserted extension columns update only enabled and revokedAt', () => {
  const headers = [
    'createdAt', 'extensionBeforeEnabled', 'enabled', 'label', 'token', 'routeTargetsJson',
    'tags', 'extensionBeforeRevokedAt', 'revokedAt', 'tagMode', 'colors', 'routeIds'
  ];
  const original = shareRow(headers, {
    createdAt: '2026-01-01T00:00:00.000Z', extensionBeforeEnabled: 'keep-a', enabled: true,
    label: 'Reordered', token: 'target-token', routeTargetsJson: '{"v":1,"targets":[]}',
    tags: 'alpha', extensionBeforeRevokedAt: 'keep-b', revokedAt: '', tagMode: 'and',
    colors: '#1e88e5', routeIds: '__share_no_routes__'
  });
  const { api, audit, sheet } = loadApi({ headers, dataRows: [original] });

  const result = api.setShareLinkEnabled(authorized({ token: 'target-token', enabled: false }));

  assert.deepEqual(plain(result), { ok: true });
  const changedIndexes = sheet.rows[1].map((value, index) => value === original[index] ? -1 : index).filter((index) => index >= 0);
  assert.deepEqual(changedIndexes, [headers.indexOf('enabled'), headers.indexOf('revokedAt')]);
  assert.deepEqual(audit.writes.map(({ row, column }) => [row, column]), [
    [2, headers.indexOf('enabled') + 1],
    [2, headers.indexOf('revokedAt') + 1]
  ]);
  assert.equal(sheet.rows[1][headers.indexOf('routeTargetsJson')], '{"v":1,"targets":[]}');
  assert.equal(sheet.rows[1][headers.indexOf('extensionBeforeEnabled')], 'keep-a');
  assert.equal(sheet.rows[1][headers.indexOf('extensionBeforeRevokedAt')], 'keep-b');
});

test('deleteShareLink finds the requested row from a reordered token header', () => {
  const headers = ['extension', 'label', 'enabled', 'routeTargetsJson', 'token', 'revokedAt'];
  const first = shareRow(headers, { extension: 'keep-1', label: 'First', enabled: true, token: 'first' });
  const target = shareRow(headers, { extension: 'delete-me', label: 'Target', enabled: true, token: 'target', routeTargetsJson: 'target-json' });
  const last = shareRow(headers, { extension: 'keep-2', label: 'Last', enabled: false, token: 'last', routeTargetsJson: 'last-json' });
  const { api, audit, sheet } = loadApi({ headers, dataRows: [first, target, last] });

  assert.deepEqual(plain(api.deleteShareLink(authorized({ token: 'target' }))), { ok: true });
  assert.deepEqual(audit.deletes, [{ row: 3, lockHeld: true }]);
  assert.deepEqual(sheet.rows, [headers, first, last]);
});

test('missing request tokens and unknown tokens never write or delete', () => {
  for (const method of ['setShareLinkEnabled', 'deleteShareLink']) {
    for (const token of ['', 'not-present']) {
      const { api, audit } = loadApi({
        dataRows: [shareRow(LEGACY_HEADERS, { token: 'existing', enabled: true })]
      });
      const result = api[method](authorized({ token, enabled: false }));
      assert.equal(result.ok, false, `${method}(${JSON.stringify(token)}) must fail closed`);
      assert.deepEqual(audit.writes, []);
      assert.deepEqual(audit.deletes, []);
    }
  }
});

test('setShareLinkEnabled fails closed for every missing or duplicate required header', () => {
  for (const requiredHeader of ['token', 'enabled', 'revokedAt']) {
    for (const mode of ['missing', 'duplicate']) {
      const headers = mode === 'missing'
        ? LEGACY_HEADERS.filter((header) => header !== requiredHeader)
        : LEGACY_HEADERS.concat([requiredHeader]);
      const values = { token: 'target', enabled: true, revokedAt: '', label: 'Keep' };
      const row = shareRow(headers, values);
      if (mode === 'duplicate') row[row.length - 1] = values[requiredHeader];
      const { api, audit, sheet } = loadApi({ headers, dataRows: [row] });
      const before = sheet.rows.map((item) => item.slice());

      const result = api.setShareLinkEnabled(authorized({ token: 'target', enabled: false }));

      assert.equal(result.ok, false, `${mode} ${requiredHeader} must fail closed`);
      assert.deepEqual(sheet.rows, before);
      assert.deepEqual(audit.writes, []);
    }
  }
});

test('deleteShareLink fails closed when the token header is missing or duplicated', () => {
  for (const headers of [
    LEGACY_HEADERS.filter((header) => header !== 'token'),
    LEGACY_HEADERS.concat(['token'])
  ]) {
    const row = shareRow(headers, { token: 'target', enabled: true });
    if (headers.filter((header) => header === 'token').length === 2) row[row.length - 1] = 'target';
    const { api, audit, sheet } = loadApi({ headers, dataRows: [row] });
    const before = sheet.rows.map((item) => item.slice());

    const result = api.deleteShareLink(authorized({ token: 'target' }));

    assert.equal(result.ok, false);
    assert.deepEqual(sheet.rows, before);
    assert.deepEqual(audit.deletes, []);
  }
});

test('duplicate token rows prevent both update and deletion', () => {
  for (const method of ['setShareLinkEnabled', 'deleteShareLink']) {
    const { api, audit, sheet } = loadApi({
      dataRows: [
        shareRow(LEGACY_HEADERS, { label: 'First', token: 'duplicate', enabled: true }),
        shareRow(LEGACY_HEADERS, { label: 'Second', token: 'duplicate', enabled: false })
      ]
    });
    const before = sheet.rows.map((item) => item.slice());

    const result = api[method](authorized({ token: 'duplicate', enabled: false }));

    assert.equal(result.ok, false);
    assert.deepEqual(sheet.rows, before);
    assert.deepEqual(audit.writes, []);
    assert.deepEqual(audit.deletes, []);
  }
});

test('read, target lookup, update and delete run only while the mutation lock is held', () => {
  for (const method of ['setShareLinkEnabled', 'deleteShareLink']) {
    const { api, audit } = loadApi({
      dataRows: [shareRow(LEGACY_HEADERS, { token: 'target', enabled: true })],
      enforceLock: true
    });

    const result = api[method](authorized({ token: 'target', enabled: false }));

    assert.equal(result.ok, true);
    assert.equal(audit.lockAttempts, 1);
    assert.equal(audit.lockReleases, 1);
    assert.equal(audit.reads.every((entry) => entry.lockHeld), true);
    assert.equal(audit.writes.every((entry) => entry.lockHeld), true);
    assert.equal(audit.deletes.every((entry) => entry.lockHeld), true);
    assert.equal(audit.events.indexOf('lock') < audit.events.indexOf('spreadsheet'), true);
    assert.equal(audit.events.at(-1), 'unlock');
  }
});

test('failed authorization and lock acquisition never access Spreadsheet or mutate rows', () => {
  const unauthorized = loadApi({
    dataRows: [shareRow(LEGACY_HEADERS, { token: 'target', enabled: true })],
    authorized: false,
    enforceLock: true
  });
  assert.throws(
    () => unauthorized.api.setShareLinkEnabled(authorized({ token: 'target', enabled: false })),
    /編集権限/
  );
  assert.equal(unauthorized.audit.spreadsheetAccesses, 0);
  assert.equal(unauthorized.audit.lockAttempts, 0);

  const busy = loadApi({
    dataRows: [shareRow(LEGACY_HEADERS, { token: 'target', enabled: true })],
    lockAcquired: false,
    enforceLock: true
  });
  const result = busy.api.deleteShareLink(authorized({ token: 'target' }));
  assert.equal(result.ok, false);
  assert.equal(busy.audit.spreadsheetAccesses, 0);
  assert.deepEqual(busy.audit.deletes, []);
  assert.equal(busy.audit.lockReleases, 0);
});

test('revokeShareLink continues to delegate to setShareLinkEnabled', () => {
  assert.match(
    codeJs,
    /function revokeShareLink\(data\)[\s\S]*?return setShareLinkEnabled\(\{ token: token, enabled: false, __editToken:/
  );
});
