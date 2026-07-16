const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');

test('common Import Preview provides Pencil 09 photo workbench and variant separation', () => {
  assert.match(indexHtml, /class="import-preview-layout"/);
  assert.match(indexHtml, /id="import-preview-photo-pane"/);
  assert.match(indexHtml, /data-import-variant/);
  assert.match(indexHtml, /\.import-preview-sheet\[data-import-variant="photo"\][\s\S]*grid-template-columns:\s*270px\s+minmax\(0,\s*1fr\)\s+430px/);
  assert.match(indexHtml, /\.import-preview-sheet\[data-import-variant="(?:csv|geojson)"\][^}]*#import-preview-photo-pane[^}]*display:\s*none/);
  assert.match(indexHtml, /sourceType === 'multi-photo'[\s\S]*'photo'/);
  assert.match(indexHtml, /sourceType === 'csv'[\s\S]*'csv'/);
  assert.match(indexHtml, /sourceType === 'geojson'[\s\S]*'geojson'/);
});

test('Import Preview has five status filters and required common actions', () => {
  const filters = ['all', 'needs-review', 'processing', 'succeeded', 'failed'];
  filters.forEach((filter) => {
    assert.match(indexHtml, new RegExp(`data-import-filter=["']${filter}["']`), filter);
  });
  ['全件', '要確認', '登録中', '登録済み', '失敗'].forEach((label) => {
    assert.match(indexHtml, new RegExp(`>${label}<`), label);
  });
  assert.match(indexHtml, /選択項目へ適用/);
  assert.match(indexHtml, /すべてへ適用/);
  assert.match(indexHtml, /id="import-preview-primary"[^>]*>[\s\S]*?<span class="action-label">取込<\/span><\/button>/);
  assert.match(indexHtml, />取込を破棄</);
  assert.match(indexHtml, />再試行</);
  assert.match(indexHtml, />残りを再開</);
});

test('photo editor and explicit route-time matching use product language', () => {
  [
    'タイトル', '説明', '撮影日時', 'タグ', '色', 'アイコン', '状態', 'URL', '緯度', '経度',
    '保存済みルートから写真の位置を推定', '保存済みルート',
    '写真のタイムゾーン（+09:00形式）', 'カメラ時計補正（秒）',
    '最大補間間隔', '端点許容時間', '照合する',
    '選択した写真へ位置を適用', '照合結果をクリア'
  ].forEach((label) => assert.equal(indexHtml.includes(label), true, label));
  assert.match(indexHtml, /保存結果を確認できませんでした。安全に再確認します。/);

  const preview = indexHtml.match(/<div id="import-preview-overlay"[\s\S]*?<div id="track-import-preview-overlay"/)[0];
  assert.doesNotMatch(preview, /response loss|idempotency|ジョブ状態|current revision/);
});

test('Pencil 10 preparation and result states remain available without changing ImportJob statuses', () => {
  ['写真準備中', 'HEIC変換中', 'キャンセル中', '一部失敗', '再試行', '残りを再開', '完了']
    .forEach((label) => assert.match(indexHtml, new RegExp(label), label));
  assert.match(indexHtml, /queued:\s*'要確認'/);
  assert.match(indexHtml, /processing:\s*'登録中'/);
  assert.match(indexHtml, /succeeded:\s*'登録済み'/);
  assert.match(indexHtml, /failed:\s*'失敗'/);
});
