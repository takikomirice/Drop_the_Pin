const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const root = path.resolve(__dirname, '..');
const codeJs = fs.readFileSync(path.join(root, 'Code.js'), 'utf8');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const readme = fs.readFileSync(path.join(root, 'README.md'), 'utf8');

function assertIncludes(source, needle, message) {
  assert.ok(source.includes(needle), message || `Expected source to include ${needle}`);
}

test('share link list exposes QR action and modal UI', () => {
  assertIncludes(indexHtml, 'data-action="qr"');
  assertIncludes(indexHtml, 'id="share-qr-overlay"');
  assertIncludes(indexHtml, 'id="share-qr-title"');
  assertIncludes(indexHtml, 'id="share-qr-code"');
  assertIncludes(indexHtml, 'id="share-qr-url"');
  assertIncludes(indexHtml, 'id="share-qr-copy"');
  assertIncludes(indexHtml, 'id="share-qr-close"');
});

test('QR modal warns that only read-only shared view URLs are encoded', () => {
  assertIncludes(indexHtml, 'このQRは閲覧専用の共有ビューURLです。編集URLや editKey は含みません。');
  assertIncludes(indexHtml, '共有前に、写真・説明・地点に個人情報や詳細すぎる位置情報が含まれていないか確認してください。');
  assertIncludes(indexHtml, 'QRコードを生成できませんでした。URLをコピーして共有してください。');
});

test('QR generation is browser-local and does not use external QR APIs', () => {
  assertIncludes(indexHtml, 'function createQrMatrix');
  assertIncludes(indexHtml, 'function renderQrSvg');
  assertIncludes(indexHtml, 'function renderShareQr');
  ['api.qrserver.com', 'chart.googleapis.com', 'quickchart.io'].forEach((needle) => {
    assert.equal(indexHtml.includes(needle), false, `Unexpected external QR API: ${needle}`);
  });
});

test('QR target URL excludes edit credentials and requires shared view URL', () => {
  assertIncludes(indexHtml, 'function getSafeSharedQrUrl');
  assertIncludes(indexHtml, "params.get('view') !== 'shared'");
  assertIncludes(indexHtml, "params.has('editKey')");
  assertIncludes(indexHtml, "params.has('mode')");
  assertIncludes(indexHtml, "params.has('__editToken')");
  assertIncludes(indexHtml, "params.has('EDIT_KEY')");
  assert.equal(indexHtml.includes('mode=edit'), false);
  assert.equal(indexHtml.includes('__editToken='), false);
});

test('shared link URLs come from buildSharedViewUrl and README documents shared QR', () => {
  assertIncludes(codeJs, 'url: buildSharedViewUrl_(token)');
  assertIncludes(codeJs, 'item.url = buildSharedViewUrl_(item.token)');
  assertIncludes(readme, '共有リンク一覧から共有ビューURLをコピー');
  assertIncludes(readme, '共有リンクごとにQRコードを表示');
  assertIncludes(readme, 'QRは閲覧専用共有ビューURL');
  assertIncludes(readme, 'QRには編集URL、editKey、編集トークンを含めません');
  assertIncludes(readme, '外部QR生成APIを使わず');
  assertIncludes(readme, '写真・説明・地点・タグ・リンク');
});
