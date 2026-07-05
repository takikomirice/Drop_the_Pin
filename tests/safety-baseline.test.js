const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const root = path.resolve(__dirname, '..');
const codeJs = fs.readFileSync(path.join(root, 'Code.js'), 'utf8');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const sharedHtml = fs.readFileSync(path.join(root, 'shared.html'), 'utf8');
const readme = fs.readFileSync(path.join(root, 'README.md'), 'utf8');
const appsscript = fs.readFileSync(path.join(root, 'appsscript.json'), 'utf8');

function assertIncludes(source, needle, message) {
  assert.ok(source.includes(needle), message || `Expected source to include ${needle}`);
}

test('README documents map providers and location privacy baseline', () => {
  assertIncludes(readme, 'Leaflet.js + OpenStreetMap');
  assertIncludes(readme, 'Google Maps API は使っていません');
  assertIncludes(readme, 'EXIF GPS');
  assertIncludes(readme, '自宅・通学路・個人が特定される場所');
  assertIncludes(readme, '共有リンクを作成する前');
  assertIncludes(readme, 'OSRM public demo server');
  assertIncludes(readme, '本番・大規模利用向けではありません');
  assertIncludes(readme, '住所 / 地名検索');
  assertIncludes(readme, 'Nominatim');
});

test('index help includes concise location privacy guidance', () => {
  assertIncludes(indexHtml, '<h3>位置情報</h3>');
  assertIncludes(indexHtml, 'GPS付き写真は撮影場所をもとに自動でピンを置きます。');
  assertIncludes(indexHtml, '自宅・通学路・個人が特定される場所の写真は投稿しないでください。');
  assertIncludes(indexHtml, '必要に応じてピンを移動するか、未配置に戻してください。');
});

test('README documents edit URL operation and shared-view safety', () => {
  assertIncludes(readme, '通常URLでは閲覧専用');
  assertIncludes(readme, 'mode=edit&editKey=');
  assertIncludes(readme, 'スプレッドシートメニューの「編集URLを表示」');
  assertIncludes(readme, 'EDIT_KEY は config シート');
  assertIncludes(readme, 'WEB_APP_URL は config シート');
  assertIncludes(readme, 'EDIT_KEY は個人認証ではなく共有鍵');
  assertIncludes(readme, '編集URLを知っている人は共同編集できます');
  assertIncludes(readme, '6時間で切れるのは EDIT_KEY ではなく一時編集トークン');
  assertIncludes(readme, '開きっぱなしで編集できなくなった場合は、同じ編集URLを再読み込み');
  assertIncludes(readme, '編集キーを再生成すると、古い編集URLは無効');
  assertIncludes(readme, '共有ビューは閲覧専用');
  assertIncludes(readme, '写真・説明・タグ・リンク・地点');
});

test('index help briefly explains edit URL usage', () => {
  assertIncludes(indexHtml, '<h3>編集について</h3>');
  assertIncludes(indexHtml, '通常URLでは閲覧専用です。');
  assertIncludes(indexHtml, '先生がスプレッドシートメニューから取得した編集URLで開いてください。');
});

test('README marks legacy Apps Script mutation test helpers as non-recommended after edit tokens', () => {
  assertIncludes(readme, 'testSaveMapData()');
  assertIncludes(readme, 'testUpdatePin()');
  assertIncludes(readme, 'testRouteCRUD()');
  assertIncludes(readme, 'Phase 1以降');
  assertIncludes(readme, '変更系関数は編集トークンで保護');
  assertIncludes(readme, '直接実行用の動作確認としては非推奨');
  assertIncludes(readme, '編集トークン免除');
});

test('normal application code does not include Google Maps API hooks or keys', () => {
  const normalCode = [codeJs, indexHtml, sharedHtml, appsscript].join('\n');
  ['google.maps', 'maps.googleapis.com', 'AIza'].forEach((needle) => {
    assert.equal(normalCode.includes(needle), false, `Unexpected Google Maps API marker: ${needle}`);
  });
});

test('application loads Leaflet, OpenStreetMap tiles, and exif-js', () => {
  assertIncludes(indexHtml, 'https://unpkg.com/leaflet@1.9.4/dist/leaflet.css');
  assertIncludes(indexHtml, 'https://unpkg.com/leaflet@1.9.4/dist/leaflet.js');
  assertIncludes(sharedHtml, 'https://unpkg.com/leaflet@1.9.4/dist/leaflet.css');
  assertIncludes(sharedHtml, 'https://unpkg.com/leaflet@1.9.4/dist/leaflet.js');
  assertIncludes(indexHtml, 'https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png');
  assertIncludes(sharedHtml, '{s}.tile.openstreetmap.org/{z}/{x}/{y}.png');
  assertIncludes(indexHtml, 'https://cdn.jsdelivr.net/npm/exif-js@2.3.0');
  assertIncludes(indexHtml, 'https://router.project-osrm.org/route/v1/driving/');
});
