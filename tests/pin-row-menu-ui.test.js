const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const indexHtml = fs.readFileSync(path.resolve(__dirname, '..', 'index.html'), 'utf8');
const css = indexHtml.match(/<style>([\s\S]*?)<\/style>/)[1];

function functionSource(name) {
  const start = indexHtml.indexOf(`function ${name}(`);
  assert.notEqual(start, -1, `Expected function ${name}`);
  const open = indexHtml.indexOf('{', start);
  let depth = 0;
  for (let index = open; index < indexHtml.length; index += 1) {
    if (indexHtml[index] === '{') depth += 1;
    if (indexHtml[index] === '}') depth -= 1;
    if (depth === 0) return indexHtml.slice(start, index + 1);
  }
  assert.fail(`Could not parse ${name}`);
}

test('pin list builds sibling main and vertical-menu buttons without nested buttons', () => {
  const build = functionSource('buildListItem');
  assert.match(build, /const row = document\.createElement\('div'\)/);
  assert.match(build, /row\.className = 'pin-list-row'/);
  assert.match(build, /dragHandle\.className = 'pin-drag-handle'/);
  assert.match(build, /button\.className = 'list-item'/);
  assert.match(build, /menuButton\.className = 'pin-row-menu-btn icon-btn'/);
  assert.match(build, /row\.appendChild\(dragHandle\)/);
  assert.match(build, /row\.appendChild\(button\)/);
  assert.match(build, /row\.appendChild\(menuButton\)/);
  assert.match(build, /return row/);
  assert.ok(build.indexOf('row.appendChild(dragHandle)') < build.indexOf('row.appendChild(button)'));
  assert.ok(build.indexOf('row.appendChild(button)') < build.indexOf('row.appendChild(menuButton)'));
  assert.doesNotMatch(build, /button\.appendChild\(menuButton\)/);
  assert.doesNotMatch(build, /button\.appendChild\(dragHandle\)/);
});

test('row menu button is accessible isolated from selection and excluded from drag', () => {
  const build = functionSource('buildListItem');
  assert.match(build, /menuButton\.setAttribute\('aria-label', 'ピンのメニューを開く'\)/);
  assert.match(build, /menuButton\.title = 'ピンのメニューを開く'/);
  assert.match(build, /menuButton\.setAttribute\('aria-haspopup', 'menu'\)/);
  assert.match(build, /menuButton\.setAttribute\('aria-controls', 'pin-menu-overlay'\)/);
  assert.match(build, /menuButton\.draggable = false/);
  assert.match(build, /menuButton\.addEventListener\('pointerdown',[\s\S]*stopPropagation/);
  assert.match(build, /menuButton\.addEventListener\('click',[\s\S]*preventDefault[\s\S]*stopPropagation[\s\S]*options\.onMenu\(\)/);
  assert.match(build, /if \(canEdit\(\) && options\.onMenu\)/);
  assert.match(css, /\.pin-row-menu-btn\s*\{[^}]*width:\s*44px[^}]*height:\s*44px/);
  assert.match(css, /\.pin-row-menu-btn\s*\{[^}]*border:\s*0[^}]*background:\s*transparent/);
  assert.ok(
    css.indexOf('#side-panel .pin-list-row .pin-row-menu-btn.icon-btn')
      > css.indexOf('.icon-btn, .text-btn, .ghost-btn, .danger-btn, .btn-primary'),
    'the borderless row-menu override must follow the common bordered button rule'
  );
});

test('row owns the frame hover and selection while drag handle is an isolated 44px control', () => {
  const build = functionSource('buildListItem');
  assert.match(build, /dragHandle\.setAttribute\('aria-label', 'ピンを並べ替え'\)/);
  assert.match(build, /dragHandle\.title = 'ピンを並べ替え'/);
  assert.doesNotMatch(build, /dragHandle\.draggable = true/);
  assert.doesNotMatch(build, /dragHandle\.addEventListener\('dragstart'/);
  assert.doesNotMatch(build, /dragHandle\.addEventListener\('dragend'/);
  assert.doesNotMatch(build, /button\.draggable = true/);

  assert.match(css, /\.pin-list-row\s*\{[^}]*min-height:\s*62px[^}]*border:[^}]*background:\s*var\(--color-surface-raised\)[^}]*box-shadow:\s*var\(--shadow-sm\)/);
  assert.match(css, /\.pin-drag-handle,\s*\.pin-row-menu-btn\s*\{[^}]*width:\s*44px[^}]*height:\s*44px[^}]*border:\s*0[^}]*background:\s*transparent/);
  assert.match(css, /#side-panel \.pin-list-row:hover\s*\{[^}]*background:\s*var\(--color-surface-muted\)/);
  assert.match(css, /#side-panel \.pin-list-row\.is-selected\s*\{[^}]*border-left-width:\s*4px[^}]*background:\s*var\(--accent-soft\)/);
  assert.match(css, /#side-panel \.pin-list-row \.list-item\s*\{[^}]*border:\s*0[^}]*background:\s*transparent[^}]*box-shadow:\s*none/);
  assert.match(functionSource('syncListItemSelection'), /closest\('\.pin-list-row'\)[\s\S]*classList\.toggle\('is-selected'/);
});

test('row menu click contextmenu and long press all reuse the existing callback', () => {
  const build = functionSource('buildListItem');
  assert.match(build, /button\.addEventListener\('contextmenu',[\s\S]*options\.onMenu\(\)/);
  assert.match(build, /attachLongPress\(button,[\s\S]*options\.onMenu\(\)/);
  assert.doesNotMatch(build, /openPinMenu\(/);
  assert.match(functionSource('openPinMenu'), /openOverlay\('pin-menu-overlay'\)/);
});

test('existing overlay keeps menu in viewport and restores explicit-button focus on Escape', () => {
  assert.match(css, /#pin-menu-overlay \.sheet-body\s*\{[^}]*width:\s*min\(360px,[^}]*max-height:\s*calc\(100dvh - 24px\)[^}]*overflow-y:\s*auto/);
  assert.match(functionSource('openOverlay'), /opener:\s*document\.activeElement/);
  assert.match(functionSource('closeOverlay'), /restoreSurfaceFocus\(record && record\.opener\)/);
  assert.match(functionSource('dismissOverlayById'), /id === 'pin-menu-overlay'\) return closePinMenu\(\)/);
});
