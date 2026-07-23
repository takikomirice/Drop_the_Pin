const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const root = path.resolve(__dirname, '..');
const { replaceVendorRegion } = require(path.join(root, 'scripts', 'sync-audio-vendor.js'));
const VENDOR_START = 'AUDIO_VENDOR_BUNDLE_START';
const VENDOR_END = 'AUDIO_VENDOR_BUNDLE_END';

function vendorRegion(source, newline = '\n') {
  return `${VENDOR_START}${newline}<script>${newline}${source}${newline}</script>${newline}${VENDOR_END}`;
}

function rawIndexWithVendor(source, options = {}) {
  const newline = options.newline || '\n';
  const prefix = options.prefix || '';
  return `${prefix}${vendorRegion(source, newline)}${newline}<!DOCTYPE html>${newline}<html></html>${newline}`;
}

function replaceRegion(indexSource, desiredRegion) {
  assert.equal(typeof replaceVendorRegion, 'function', 'generator must export its pure region replacer');
  return replaceVendorRegion(indexSource, desiredRegion);
}

function countOccurrences(text, needle) {
  return text.split(needle).length - 1;
}

test('vendor generator replaces a valid stale prefix region', () => {
  const original = rawIndexWithVendor('globalThis.staleVendor=true;', { prefix: ' \t\n' });
  const desired = vendorRegion('globalThis.desiredVendor=true;');
  const updated = replaceRegion(original, desired);

  assert.equal(updated, ` \t\n${desired}\n<!DOCTYPE html>\n<html></html>\n`);
  assert.equal(updated.includes('globalThis.staleVendor=true;'), false);
  assert.equal(countOccurrences(updated, 'globalThis.desiredVendor=true;'), 1);
});

test('vendor generator replacement is byte-idempotent for the same desired region', () => {
  const original = rawIndexWithVendor('globalThis.staleVendor=true;', { newline: '\r\n' });
  const desired = vendorRegion('globalThis.desiredVendor=true;');
  const once = replaceRegion(original, desired);
  const twice = replaceRegion(once, desired);

  assert.equal(twice, once);
});

test('vendor generator rejects a marker region after non-whitespace content', () => {
  const source = rawIndexWithVendor('globalThis.staleVendor=true;', { prefix: 'unexpected prefix\n' });
  assert.throws(
    () => replaceRegion(source, vendorRegion('globalThis.desiredVendor=true;')),
    /prefix|whitespace|start/i
  );
});

test('vendor generator rejects duplicate and reversed markers', () => {
  const valid = rawIndexWithVendor('globalThis.staleVendor=true;');
  const duplicateStart = valid.replace(VENDOR_START, `${VENDOR_START}\n${VENDOR_START}`);
  const duplicateEnd = valid.replace(VENDOR_END, `${VENDOR_END}\n${VENDOR_END}`);
  const reversed = `${VENDOR_END}\n<script>\nglobalThis.staleVendor=true;\n</script>\n${VENDOR_START}\n<!DOCTYPE html>`;
  const desired = vendorRegion('globalThis.desiredVendor=true;');

  assert.throws(() => replaceRegion(duplicateStart, desired), /marker|pair|once/i);
  assert.throws(() => replaceRegion(duplicateEnd, desired), /marker|pair|once/i);
  assert.throws(() => replaceRegion(reversed, desired), /marker|order/i);
});

test('vendor generator rejects a non-plain script wrapper', () => {
  const source = rawIndexWithVendor('globalThis.staleVendor=true;')
    .replace('<script>', '<script type="text/javascript">');
  assert.throws(
    () => replaceRegion(source, vendorRegion('globalThis.desiredVendor=true;')),
    /plain script|script wrapper/i
  );
});

test('vendor generator rejects an empty current script region', () => {
  assert.throws(
    () => replaceRegion(
      rawIndexWithVendor('   '),
      vendorRegion('globalThis.desiredVendor=true;')
    ),
    /empty|nonempty|source/i
  );
});

test('vendor generator requires an adjacent index doctype after the end sentinel', () => {
  const region = vendorRegion('globalThis.staleVendor=true;');
  const desired = vendorRegion('globalThis.desiredVendor=true;');

  assert.throws(() => replaceRegion(`${region}\n<html></html>`, desired), /doctype|document|adjacent/i);
  assert.throws(() => replaceRegion(`${region}\n\n<!DOCTYPE html>`, desired), /doctype|document|adjacent/i);
});

test('vendor generator makes stale valid source observably different for check mode', () => {
  const original = rawIndexWithVendor('globalThis.staleVendor=true;');
  const updated = replaceRegion(original, vendorRegion('globalThis.desiredVendor=true;'));
  assert.notEqual(updated, original);
});

test('audio vendor package and generated bundle follow the delivery contract', () => {
  const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
  assert.deepEqual(pkg.dependencies, {
    '@mediabunny/mp3-encoder': '1.50.8',
    mediabunny: '1.50.8'
  });
  assert.equal(pkg.devDependencies.esbuild, '0.25.6');
  assert.equal(pkg.devDependencies['@playwright/test'], '1.61.1');

  const vendorHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
  const start = 'AUDIO_VENDOR_BUNDLE_START\n<script>\n';
  const end = '\n</script>\nAUDIO_VENDOR_BUNDLE_END\n';
  const startIndex = vendorHtml.indexOf(start);
  const endIndex = vendorHtml.indexOf(end, startIndex + start.length);
  assert.equal(vendorHtml.slice(0, startIndex).trim(), '');
  assert.equal(startIndex >= 0, true);
  assert.equal(endIndex > startIndex, true);
  assert.equal(countOccurrences(vendorHtml, 'AUDIO_VENDOR_BUNDLE_START'), 1);
  assert.equal(countOccurrences(vendorHtml, 'AUDIO_VENDOR_BUNDLE_END'), 1);

  const bundle = vendorHtml.slice(startIndex + start.length, endIndex);
  assert.ok(bundle.length > 300000);
  assert.equal(countOccurrences(bundle, '</script>'), 0);
  assert.equal(countOccurrences(bundle, 'globalThis.Mediabunny='), 1);
  assert.equal(countOccurrences(bundle, 'globalThis.MediabunnyMp3Encoder='), 1);
  const strippedIndex = vendorHtml.slice(endIndex + end.length);
  assert.match(strippedIndex, /^<!DOCTYPE html>/);
  assert.equal(strippedIndex.includes('AUDIO_VENDOR_BUNDLE_START'), false);
  assert.equal(strippedIndex.includes('globalThis.Mediabunny='), false);
  assert.equal(strippedIndex.includes('globalThis.MediabunnyMp3Encoder='), false);

  const sharedSource = fs.readFileSync(path.join(root, 'shared.html'), 'utf8');
  assert.equal(sharedSource.includes('AUDIO_VENDOR_BUNDLE_START'), false);
  assert.equal(sharedSource.includes('globalThis.Mediabunny='), false);
  assert.equal(sharedSource.includes('globalThis.MediabunnyMp3Encoder='), false);

  const syncScript = fs.readFileSync(path.join(root, 'scripts', 'sync-audio-vendor.js'), 'utf8');
  assert.match(syncScript, /path\.join\(root, ['"]index\.html['"]\)/);

  assert.equal(
    fs.readFileSync(path.join(root, 'vendor', 'mediabunny-LICENSE.txt'), 'utf8'),
    fs.readFileSync(path.join(root, 'node_modules', 'mediabunny', 'LICENSE'), 'utf8')
  );
  assert.equal(
    fs.readFileSync(path.join(root, 'vendor', 'mediabunny-mp3-encoder-LICENSE.txt'), 'utf8'),
    fs.readFileSync(path.join(root, 'node_modules', '@mediabunny', 'mp3-encoder', 'LICENSE'), 'utf8')
  );

  const claspIgnore = fs.readFileSync(path.join(root, '.claspignore'), 'utf8');
  for (const ignoredPath of [
    'node_modules/**',
    'scripts/**',
    'vendor/**',
    'package.json',
    'package-lock.json',
    'playwright.config.js',
    'playwright-report/**',
    'blob-report/**',
    'test-results/**'
  ]) {
    assert.equal(claspIgnore.includes(`${ignoredPath}\n`), true);
  }
});
