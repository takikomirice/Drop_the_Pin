const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');

function countId(id) {
  return (indexHtml.match(new RegExp(`id=["']${id}["']`, 'g')) || []).length;
}

test('Drive photo import is a root-scoped workbench with a persistent archive notice', () => {
  assert.match(indexHtml, /class="drive-photo-import-workbench"/);
  assert.match(indexHtml, /class="[^"\n]*drive-photo-import-folder-pane[^"\n]*"/);
  assert.match(indexHtml, /class="[^"\n]*drive-photo-import-photo-pane[^"\n]*"/);
  assert.match(indexHtml, /id="drive-photo-import-root-name"/);
  assert.match(indexHtml, /id="drive-photo-import-folder-name"/);
  assert.match(indexHtml, /取込元/);
  assert.doesNotMatch(indexHtml.match(/<div id="drive-photo-import-overlay"[\s\S]*?<div id="settings-overlay"/)[0], /管理用コピーの保存先/);
  assert.match(indexHtml, /id="drive-photo-import-storage-note"/);
  assert.match(indexHtml, /登録成功後、表示用の管理JPEGを作成し、元写真はoriginalフォルダへ移動します。元写真自体は削除・変換しません。/);
  assert.doesNotMatch(indexHtml, /drive-photo-import-consent/);
  assert.match(indexHtml, /1〜20枚[^<]*1枚15MBまで[^<]*合計100MBまで/);

  [
    'drive-photo-import-folder-name', 'drive-photo-import-folder-list',
    'drive-photo-import-photo-list', 'drive-photo-import-photo-count',
    'drive-photo-import-hide-imported', 'drive-photo-import-selection-count',
    'drive-photo-import-selection-size', 'drive-photo-import-status',
    'drive-photo-import-error', 'drive-photo-import-confirm'
  ].forEach((id) => assert.equal(countId(id), 1, id));
  assert.match(indexHtml, /id="drive-photo-import-photo-count"[^>]*aria-live="polite"/);
  assert.match(indexHtml, /id="drive-photo-import-hide-imported"[^>]*type="checkbox"/);
  assert.match(indexHtml, /取込済みの写真を表示しない/);
  assert.match(indexHtml, /drive-photo-import-badge/);
});

test('Drive dialog keeps its footer fixed while only the two lists scroll at desktop and mobile sizes', () => {
  assert.match(indexHtml, /\.drive-photo-import-workbench\s*\{[^}]*min-height:\s*0[^}]*overflow:\s*hidden/s);
  assert.match(indexHtml, /\.drive-photo-import-footer-actions\s*\{[^}]*flex-wrap:\s*nowrap/s);
  assert.match(indexHtml, /@media \(max-width:\s*900px\)[\s\S]*?\.drive-photo-import-footer-actions\s*\{[^}]*flex-wrap:\s*wrap/);
  assert.match(indexHtml, /@media \(max-width:\s*900px\) and \(max-height:\s*600px\)[\s\S]*?\.drive-photo-import-workbench\s*\{[^}]*minmax\(0,\s*1fr\)/);
});

test('Drive workbench exposes user-facing loading, empty, limit, and preserved-navigation states', () => {
  assert.match(indexHtml, /drive-photo-import-folder-loading/);
  assert.match(indexHtml, /drive-photo-import-photo-empty/);
  assert.match(indexHtml, /drive-photo-import-limit-status/);
  assert.match(indexHtml, /'以前の一覧と'\s*\+\s*sourceState\.selectedIds\.length[\s\S]*?'件の選択を維持しています。'/);
  assert.match(indexHtml, /rootFolderName/);
  assert.match(indexHtml, /sourceState\.rootFolderId\s*=\s*validated\.folder\.id/);
  assert.doesNotMatch(indexHtml, /getTargetFolderLabel/);
  assert.match(indexHtml, /resolvedTargetFolderId[\s\S]*?driveSourceState\.rootFolderId/);
  assert.doesNotMatch(indexHtml, /if\s*\(\s*!driveSourceImport\s*&&/);
  assert.doesNotMatch(
    indexHtml.match(/<div id="drive-photo-import-overlay"[\s\S]*?<div id="settings-overlay"/)[0],
    /マイドライブ|Drive ID|owner|permission|完全path/
  );
});
