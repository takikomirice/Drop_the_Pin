const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const indexHtml = fs.readFileSync(path.resolve(__dirname, '..', 'index.html'), 'utf8');
const css = indexHtml.match(/<style>([\s\S]*?)<\/style>/)[1];
const mainScriptStart = indexHtml.lastIndexOf('<script>', indexHtml.indexOf('const appStartupStartedAt'));
const bodyMarkup = indexHtml.slice(indexHtml.indexOf('<body'), mainScriptStart)
  .replace(/<script(?:\s[^>]*)?>[\s\S]*?<\/script>/gi, '');

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function cssRule(selector) {
  const match = css.match(new RegExp(`(?:^|\\})\\s*${escapeRegExp(selector)}\\s*\\{([^}]*)\\}`, 'm'));
  assert.ok(match, `Expected CSS rule ${selector}`);
  return match[1];
}

function blockStartingAt(marker) {
  const start = css.indexOf(marker);
  assert.notEqual(start, -1, `Expected CSS block ${marker}`);
  const open = css.indexOf('{', start);
  let depth = 0;
  for (let index = open; index < css.length; index += 1) {
    if (css[index] === '{') depth += 1;
    if (css[index] === '}') depth -= 1;
    if (depth === 0) return css.slice(open + 1, index);
  }
  assert.fail(`Could not parse CSS block ${marker}`);
}

function elementWithId(id) {
  const match = bodyMarkup.match(new RegExp(`<([a-z]+)\\b[^>]*\\bid="${id}"[^>]*>[\\s\\S]*?<\\/\\1>`));
  assert.ok(match, `Expected element #${id}`);
  return match[0];
}

function visibleText(markup) {
  return markup.replace(/<svg\b[\s\S]*?<\/svg>/g, '').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
}

test('light and dark themes expose reusable foreground tokens for accent and danger surfaces', () => {
  const light = blockStartingAt(':root, [data-theme="light"]');
  const dark = blockStartingAt('[data-theme="dark"]');
  for (const block of [light, dark]) {
    assert.match(block, /--on-accent:\s*#[0-9A-F]{6}/i);
    assert.match(block, /--on-danger:\s*#[0-9A-F]{6}/i);
  }

  assert.match(cssRule('.btn-primary'), /color:\s*var\(--on-accent\)/);
  assert.match(cssRule('.danger-btn'), /color:\s*var\(--on-danger\)/);
  assert.match(cssRule('#pin-add-btn'), /background:\s*transparent/);
  assert.match(cssRule('#placement-confirm'), /color:\s*var\(--on-accent\)/);
  assert.match(cssRule('#dnd-delete-zone'), /color:\s*var\(--on-danger\)/);
  assert.match(cssRule('.dialog-validation-summary'), /color:\s*var\(--danger\)/);
  assert.doesNotMatch(cssRule('[data-theme="dark"] .icon-btn, [data-theme="dark"] .text-btn, [data-theme="dark"] .ghost-btn'), /rgba\(/);
});

test('all common button variants share a legible disabled state without changing their labels', () => {
  const disabled = cssRule(':is(.icon-btn, .text-btn, .ghost-btn, .danger-btn, .btn-primary):disabled');
  assert.match(disabled, /opacity:\s*0\.5[025]/);
  assert.match(disabled, /cursor:\s*not-allowed/);

  [
    'upload-submit', 'settings-save', 'edit-save', 'delete-confirm',
    'drive-photo-import-confirm', 'import-preview-primary', 'import-preview-cancel',
    'track-import-preview-save', 'app-confirmation-confirm', 'app-confirmation-cancel'
  ].forEach((id) => assert.notEqual(visibleText(elementWithId(id)), '', `#${id} must retain its label`));

  const bulkDelete = elementWithId('bulk-delete-btn');
  assert.equal(visibleText(bulkDelete), '', '#bulk-delete-btn is intentionally icon-only');
  assert.match(bulkDelete, /\btitle="削除"/);
  assert.match(bulkDelete, /\baria-label="削除"/);
});

test('button and navigation labels stay on one line while dialog titles allow at most two lines', () => {
  assert.match(cssRule('.icon-btn, .text-btn, .ghost-btn, .danger-btn, .btn-primary'), /white-space:\s*nowrap/);
  assert.match(cssRule('.dock-pin-tab, #mobile-sheet-tabs button, .import-preview-count, .topbar-menu-item'), /white-space:\s*nowrap/);

  const title = cssRule('.sheet-title');
  assert.match(title, /line-height:\s*1\.(?:3[45]|4)/);
  assert.match(title, /max-block-size:\s*2\.(?:6|7|8)em/);
  assert.match(title, /overflow:\s*hidden/);
  assert.match(title, /overflow-wrap:\s*anywhere/);
});

test('long labels and user data cannot force action rows or dialog headers outside the viewport', () => {
  assert.match(cssRule('.action-label'), /overflow:\s*hidden/);
  assert.match(cssRule('.action-label'), /text-overflow:\s*ellipsis/);
  assert.match(cssRule('#file-drop .action-label'), /text-overflow:\s*ellipsis/);
  assert.match(cssRule('#file-drop .action-label'), /white-space:\s*nowrap/);
  assert.match(cssRule('.workbench-header > div, .single-pin-header > div, .import-preview-heading-row > div'), /min-width:\s*0/);
  assert.match(cssRule('.share-link-head > div:first-child'), /min-width:\s*0/);
  assert.match(cssRule('.share-link-title'), /overflow-wrap:\s*anywhere/);
  assert.match(cssRule('.drive-photo-import-footer-actions'), /flex-wrap:\s*nowrap/);
  assert.match(blockStartingAt('@media (max-width: 900px)'),
    /\.drive-photo-import-footer-actions\s*\{[^}]*flex-wrap:\s*wrap/);
  assert.match(cssRule('.track-actions'), /flex-wrap:\s*wrap/);
});

test('sheet handles are reserved for mobile bottom sheets', () => {
  assert.doesNotMatch(cssRule('.sheet-handle'), /display:\s*none/);
  assert.match(css, /@media\s*\(min-width:\s*900px\)\s*\{\s*\.sheet-handle\s*\{\s*display:\s*none;?\s*\}\s*\}/);
  assert.match(blockStartingAt('@media not all and (min-width: 900px)'), /#import-preview-overlay \.sheet-handle\s*\{\s*display:\s*none;?\s*\}/);
});

test('reduced motion keeps modal interaction immediate', () => {
  const reducedMotion = blockStartingAt('@media (prefers-reduced-motion: reduce)');
  assert.match(reducedMotion, /animation-duration:\s*0\.01ms\s*!important/);
  assert.match(reducedMotion, /animation-iteration-count:\s*1\s*!important/);
  assert.match(reducedMotion, /transition-duration:\s*0\.01ms\s*!important/);
});
