const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');

function countId(id) {
  const escaped = id.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return (indexHtml.match(new RegExp(`\\bid=["']${escaped}["']`, 'g')) || []).length;
}

test('app shell keeps every required DOM id exactly once and has no duplicate ids', () => {
  [
    'app-shell',
    'topbar',
    'theme-toggle',
    'panel-toggle',
    'help-open-btn',
    'share-open-btn',
    'settings-toggle',
    'mode-badge',
    'edit-toggle',
    'map',
    'map-search-bar',
    'map-search-input',
    'map-search-toggle',
    'side-panel',
    'side-routes',
    'route-list',
    'side-tracks',
    'track-list',
    'side-unplaced',
    'side-placed',
    'pin-add-btn'
  ].forEach((id) => assert.equal(countId(id), 1, `${id} must exist exactly once`));
  assert.equal(countId('pin-panel-toggle'), 0, 'pin-panel-toggle must not exist');

  const ids = Array.from(indexHtml.matchAll(/\bid=(["'])([^"']+)\1/g), (match) => match[2]);
  const duplicates = ids.filter((id, index) => ids.indexOf(id) !== index);
  assert.deepEqual(duplicates, []);

  const shell = indexHtml.match(/<main id="app-shell">([\s\S]*?)<\/main>/);
  assert.ok(shell, 'app-shell main must wrap the map workspace');
  ['id="map"', 'id="map-search-bar"', 'id="side-panel"'].forEach((needle) => {
    assert.ok(shell[1].includes(needle), `app-shell must contain ${needle}`);
  });
});

test('Pencil Foundations are exposed as light and dark CSS tokens', () => {
  [
    '--font-family: "Noto Sans JP"',
    '--font-size-caption: 13px',
    '--font-size-body: 16px',
    '--font-size-heading: 24px',
    '--font-size-display: 36px',
    '--sp-1: 4px',
    '--sp-2: 8px',
    '--sp-3: 12px',
    '--sp-4: 16px',
    '--sp-6: 24px',
    '--sp-8: 32px',
    '--radius-sm: 8px',
    '--radius-md: 12px',
    '--radius-lg: 18px',
    '--border-width: 1px',
    '--focus: #2E90FA',
    '--route-pin: #287A4B',
    '--route-gpx: #3567C8',
    '--route-geo: #8A4CB3'
  ].forEach((token) => assert.ok(indexHtml.includes(token), `missing token ${token}`));

  [
    '--color-surface: #F7F8F5',
    '--color-surface-muted: #EDF1EC',
    '--color-surface-raised: #FFFFFF',
    '--color-text: #172019',
    '--color-text-muted: #536158',
    '--accent: #22643D',
    '--accent-soft: #DCECDF',
    '--accent-strong: #174A2D',
    '--border: #C9D2CB',
    '--danger: #B42318',
    '--warning: #8A4B08',
    '--info: #175CD3'
  ].forEach((token) => assert.ok(indexHtml.includes(token), `missing light token ${token}`));

  [
    '--color-surface: #111713',
    '--color-surface-muted: #252E27',
    '--color-surface-raised: #1A211C',
    '--color-text: #F2F6F3',
    '--color-text-muted: #B8C3BB',
    '--accent: #64C987',
    '--accent-soft: #1F3A29',
    '--accent-strong: #89DDA4',
    '--border: #3C4940',
    '--danger: #FF8B82',
    '--warning: #F3B35D',
    '--info: #84ADFF'
  ].forEach((token) => assert.ok(indexHtml.includes(token), `missing dark token ${token}`));
});

test('the 56px shell preserves body mode contracts and accessible motion rules', () => {
  assert.ok(indexHtml.includes('--app-header-height: 56px'));
  assert.match(indexHtml, /#topbar\s*\{[\s\S]*?height:\s*var\(--app-header-height\)/);
  assert.match(indexHtml, /#app-shell\s*\{[\s\S]*?inset:\s*var\(--app-header-height\) 0 0/);
  assert.match(indexHtml, /#topbar \.icon-btn[\s\S]*?min-width:\s*44px[\s\S]*?min-height:\s*44px/);
  assert.match(indexHtml, /--app-dock-width:\s*clamp\(340px,\s*26\.35vw,\s*360px\)/);
  assert.match(indexHtml, /body:not\(\.panel-hidden\) #map\s*\{\s*right:\s*var\(--app-dock-width\)/);

  [
    'body.edit-mode',
    'body.view-only',
    'body.share-mode',
    'body.panel-visible',
    'body.panel-hidden',
    'body:not(.has-edit-token)',
    'body.edit-url-denied'
  ].forEach((mode) => assert.ok(indexHtml.includes(mode), `missing body mode contract ${mode}`));

  assert.match(indexHtml, /:where\(a, button, input, textarea, select, \[tabindex\]\):focus-visible\s*\{/);
  assert.ok(indexHtml.includes('outline: var(--focus-ring-width) solid var(--focus)'));
  assert.ok(indexHtml.includes('outline-offset: var(--focus-ring-offset)'));
  assert.match(indexHtml, /\.route-visibility:focus-visible\s*\{\s*outline:\s*var\(--focus-ring-width\) solid var\(--focus\)/);
  assert.ok(indexHtml.includes('@media (prefers-reduced-motion: reduce)'));
  assert.ok(indexHtml.includes('animation-duration: 0.01ms !important'));
  assert.ok(indexHtml.includes('transition-duration: 0.01ms !important'));
});

test('viewport allows zoom and design annotations are absent from product copy', () => {
  const viewport = indexHtml.match(/<meta name="viewport" content="([^"]+)">/);
  assert.ok(viewport);
  assert.equal(viewport[1], 'width=device-width, initial-scale=1.0, viewport-fit=cover');

  [
    'Decision',
    'desktopのみ',
    'sharedで3種類のルートを表示',
    '失敗分だけ再開'
  ].forEach((annotation) => assert.equal(indexHtml.includes(annotation), false, `${annotation} must not be product copy`));
});
