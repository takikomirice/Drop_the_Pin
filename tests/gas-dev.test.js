const test = require('node:test');
const assert = require('node:assert/strict');
const { validateTarget, compareFiles, assertUnchanged, preparePush } = require('../scripts/gas-dev');

test('rejects a clasp project pointing at another script or source directory', () => {
  assert.throws(() => validateTarget({ scriptId: 'other', rootDir: '.' }, { scriptId: 'target' }), /scriptId/);
  assert.throws(() => validateTarget({ scriptId: 'target', rootDir: 'elsewhere' }, { scriptId: 'target' }), /rootDir/);
  assert.doesNotThrow(() => validateTarget({ scriptId: 'target', rootDir: '.' }, { scriptId: 'target' }));
});

test('comparison ignores line endings but detects changed, added and removed files', () => {
  assert.deepEqual(compareFiles({ 'Code.js': 'a\r\n', 'old.html': 'x' }, { 'Code.js': 'a\n', 'new.html': 'x' }), ['new.html', 'old.html']);
  assert.deepEqual(compareFiles({ 'Code.js': 'a' }, { 'Code.js': 'b' }), ['Code.js']);
});

test('a changed remote blocks push even if only one file changed', () => {
  assert.throws(() => assertUnchanged({ 'Code.js': 'old' }, { 'Code.js': 'new' }), /Remote changed/);
  assert.doesNotThrow(() => assertUnchanged({ 'Code.js': 'same' }, { 'Code.js': 'same' }));
});

test('push keeps the remote manifest and excludes local tooling', () => {
  const remote = { 'Code.js': 'old', 'index.html': 'old', 'shared.html': 'old', 'appsscript.json': '{"webapp":{"access":"MYSELF"}}' };
  const local = { 'Code.js': 'new', 'index.html': 'new', 'shared.html': 'new', 'appsscript.json': '{"webapp":{"access":"ANYONE"}}', 'secret.json': 'private' };
  assert.deepEqual(preparePush(local, remote), { ...remote, 'Code.js': 'new', 'index.html': 'new', 'shared.html': 'new' });
  assert.throws(() => preparePush(local, { ...remote, 'extra.gs': 'keep me' }), /Unexpected remote files/);
  assert.throws(() => preparePush({ ...local, 'index.html': '' }, remote), /Missing local source/);
});

test('validation preserves a failing exit code and never executes the next check', () => {
  const { runChecks } = require('../scripts/validate');
  assert.equal(runChecks([['-e', 'process.exit(7)'], ['-e', 'process.exit(9)']]), 7);
  assert.equal(runChecks([['-e', 'process.exit(0)']]), 0);
});
