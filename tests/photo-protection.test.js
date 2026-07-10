const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const sharedHtml = fs.readFileSync(path.join(root, 'shared.html'), 'utf8');

function assertIncludes(source, needle, message) {
  assert.ok(source.includes(needle), message || `Expected source to include ${needle}`);
}

function sourceFunction(source, name) {
  const start = source.indexOf(`function ${name}`);
  assert.notEqual(start, -1, `Expected function ${name} to exist`);
  const open = source.indexOf('{', start);
  let depth = 0;
  for (let index = open; index < source.length; index += 1) {
    if (source[index] === '{') depth += 1;
    if (source[index] === '}') depth -= 1;
    if (depth === 0) return source.slice(start, index + 1);
  }
  assert.fail(`Could not parse ${name}`);
}

function loadCopyGuard(source, selection, photos) {
  const context = {
    window: { getSelection: () => selection },
    document: { querySelectorAll: () => photos }
  };
  vm.runInNewContext([
    sourceFunction(source, 'isProtectedPhotoTarget'),
    sourceFunction(source, 'selectionIncludesProtectedPhoto'),
    sourceFunction(source, 'shouldBlockProtectedPhotoCopy'),
    'globalThis.shouldBlockProtectedPhotoCopy = shouldBlockProtectedPhotoCopy;'
  ].join('\n'), context);
  return context.shouldBlockProtectedPhotoCopy;
}

function loadPhotoGuards(source, selection, photos) {
  const listeners = {};
  const context = {
    window: { getSelection: () => selection },
    document: {
      querySelectorAll: () => photos,
      addEventListener: (name, listener) => { listeners[name] = listener; }
    }
  };
  vm.runInNewContext([
    sourceFunction(source, 'isProtectedPhotoTarget'),
    sourceFunction(source, 'selectionIncludesProtectedPhoto'),
    sourceFunction(source, 'shouldBlockProtectedPhotoCopy'),
    sourceFunction(source, 'installProtectedPhotoGuards'),
    'installProtectedPhotoGuards();'
  ].join('\n'), context);
  return listeners;
}

test('normal detail image is marked as a non-draggable protected photo', () => {
  assertIncludes(indexHtml, 'id="pin-detail-image" class="protected-photo" draggable="false"');
  assertIncludes(indexHtml, '.protected-photo');
});

test('shared detail creates protected non-draggable photos dynamically', () => {
  assertIncludes(sharedHtml, 'class="detail-image protected-photo"');
  assertIncludes(sharedHtml, 'draggable="false"');
  assertIncludes(sharedHtml, '.protected-photo');
});

test('both views delegate context, drag, and copy guards only to protected photos', () => {
  [indexHtml, sharedHtml].forEach((source) => {
    const photo = { id: 'photo' };
    const selection = { rangeCount: 0, isCollapsed: true };
    const guards = loadPhotoGuards(source, selection, [photo]);
    const target = { closest: () => photo };

    ['contextmenu', 'dragstart'].forEach((eventName) => {
      let prevented = false;
      guards[eventName]({ target, preventDefault: () => { prevented = true; } });
      assert.equal(prevented, true, `${eventName} should be prevented on a protected photo`);
    });
  });
});

test('copy guard blocks a protected-photo target or a selection containing one', () => {
  const photo = { id: 'photo' };
  const selected = { rangeCount: 1, isCollapsed: false, getRangeAt: () => ({ intersectsNode: (node) => node === photo }) };
  const guard = loadCopyGuard(indexHtml, selected, [photo]);
  const guards = loadPhotoGuards(indexHtml, selected, [photo]);

  assert.equal(guard({ target: { closest: () => null } }), true);
  assert.equal(guard({ target: { closest: () => photo } }), true);
  let prevented = false;
  guards.copy({ target: { closest: () => null }, preventDefault: () => { prevented = true; } });
  assert.equal(prevented, true);
});

test('copy guard allows a description-only selection and handles missing photos', () => {
  const descriptionSelection = { rangeCount: 1, isCollapsed: false, getRangeAt: () => ({ intersectsNode: () => false }) };
  let guard = loadCopyGuard(sharedHtml, descriptionSelection, [{ id: 'photo' }]);
  assert.equal(guard({ target: { closest: () => null } }), false);
  const guards = loadPhotoGuards(sharedHtml, descriptionSelection, [{ id: 'photo' }]);
  let prevented = false;
  guards.copy({ target: { closest: () => null }, preventDefault: () => { prevented = true; } });
  assert.equal(prevented, false);

  guard = loadCopyGuard(sharedHtml, null, []);
  assert.equal(guard({ target: null }), false);
});
