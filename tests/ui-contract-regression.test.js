const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const sharedHtml = fs.readFileSync(path.join(root, 'shared.html'), 'utf8');
const readme = fs.readFileSync(path.join(root, 'README.md'), 'utf8');

function bodyMarkupBeforeScript(source) {
  const bodyStart = source.indexOf('<body');
  const scriptStart = source.indexOf('<script>', bodyStart);
  return source.slice(bodyStart, scriptStart);
}

test('integrated shells retain Pencil foundations and 44px operation targets', () => {
  assert.match(indexHtml, /\.icon-btn,\s*\.text-btn,\s*\.ghost-btn,\s*\.danger-btn,\s*\.btn-primary[\s\S]{0,180}min-height:\s*44px/);
  assert.match(indexHtml, /\.icon-btn[\s\S]{0,120}min-width:\s*44px[\s\S]{0,120}height:\s*44px/);
  assert.match(indexHtml, /\.color-filter-swatch\s*\{[^}]*width:\s*44px[^}]*height:\s*44px[^}]*background:\s*transparent/);
  assert.match(indexHtml, /\.color-filter-pin-svg\s*\{[^}]*width:\s*18px[^}]*height:\s*25px/);

  assert.match(sharedHtml, /font-family:\s*"Noto Sans JP"/);
  assert.match(sharedHtml, /#shared-help-open-btn,[\s\S]{0,220}width:\s*44px[\s\S]{0,120}height:\s*44px/);
  assert.match(sharedHtml, /\.shared-control-btn,[\s\S]{0,180}min-height:\s*44px/);
  assert.match(sharedHtml, /\.route-visibility[\s\S]{0,180}min-width:\s*44px[\s\S]{0,80}min-height:\s*44px/);
  assert.match(sharedHtml, /\.route-fit[\s\S]{0,180}min-width:\s*44px[\s\S]{0,80}min-height:\s*44px/);
  assert.match(sharedHtml, /\.shared-color-swatch[\s\S]{0,180}width:\s*44px[\s\S]{0,180}background-clip:\s*content-box/);
});

test('shared interaction states use the current light and dark design tokens', () => {
  assert.doesNotMatch(sharedHtml, /185,\s*95,\s*45|71,\s*49,\s*25|#4f6b44/i);
  assert.match(sharedHtml, /\.filter-toggle-btn\.active,[\s\S]{0,180}border-color:\s*var\(--accent\)[\s\S]{0,180}background:\s*var\(--accent-soft\)/);
});

test('desktop docks and narrow bottom sheets reclaim space without overlap', () => {
  assert.match(indexHtml, /#map-search-bar[\s\S]{0,260}max-height:\s*calc\(100dvh/);
  assert.match(sharedHtml, /body\.shared-panel-hidden #shared-map\s*\{\s*right:\s*0/);
  assert.match(sharedHtml, /#shared-list-panel\s*\{\s*display:\s*contents/);
  assert.match(sharedHtml, /#shared-side-panel\s*\{[\s\S]{0,260}width:\s*min\(var\(--app-dock-width\),\s*100%\)/);
  assert.match(sharedHtml, /@media \(max-width:\s*640px\)[\s\S]*?#shared-side-panel\s*\{[^}]*border-radius:\s*20px 20px 0 0/);
  assert.match(sharedHtml, /id="shared-mobile-sheet-handle"/);
  assert.match(sharedHtml, /#shared-map-search-bar[\s\S]{0,480}max-height:\s*calc\(100%\s*-\s*40px\)/);
  assert.match(sharedHtml, /@media \(prefers-reduced-motion: reduce\)[\s\S]{0,240}\*, \*::before, \*::after/);
});

test('shared view exposes keyboard and in-panel loading empty error states', () => {
  assert.match(sharedHtml, /item\s*=\s*document\.createElement\('button'\)/);
  assert.match(sharedHtml, /item\.type\s*=\s*'button'/);
  assert.match(sharedHtml, /item\.setAttribute\('aria-label'/);
  assert.doesNotMatch(sharedHtml, /item\.setAttribute\('role',\s*'button'\)/);
  assert.doesNotMatch(sharedHtml, /item\.addEventListener\('keydown'/);
  assert.match(sharedHtml, /共有リンクを読み込んでいます/);
  assert.match(sharedHtml, /公開されているルートはありません/);
  assert.match(sharedHtml, /ピンを表示できません/);
  assert.match(sharedHtml, /id="shared-mobile-pins-tab"[\s\S]{0,180}aria-controls="shared-pin-section"/);
  assert.match(sharedHtml, /id="shared-mobile-routes-tab"[\s\S]{0,180}aria-controls="shared-route-section"/);
});

test('viewing markup stays free of edit affordances design notes and internal retry terms', () => {
  const sharedProductMarkup = bodyMarkupBeforeScript(sharedHtml).replace(/\sdraggable="false"/g, '');
  assert.doesNotMatch(sharedProductMarkup, /draggable|drag-handle|編集トークン|response loss|idempotency|Decision|desktopのみ|sharedで3種類/);
  assert.doesNotMatch(sharedHtml, /Sortable|edit-token|share-open-btn/);
  assert.doesNotMatch(indexHtml, /maximum-scale/);
});

test('README describes current typed route sharing without stale unsupported claims', () => {
  assert.match(readme, /ピンルート、GPXルート、GeoJSONルート/);
  assert.doesNotMatch(readme, /保存済みGPXトラックはshared viewへ公開せず/);
  assert.doesNotMatch(readme, /shared viewのトラック表示は未対応/);
  assert.doesNotMatch(readme, /track shared viewは後続Phase/);
});
