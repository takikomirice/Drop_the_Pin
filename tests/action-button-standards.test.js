const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const indexHtml = fs.readFileSync(path.resolve(__dirname, '..', 'index.html'), 'utf8');
const css = indexHtml.match(/<style>([\s\S]*?)<\/style>/)[1];
const mainScriptStart = indexHtml.lastIndexOf('<script>', indexHtml.indexOf('const appStartupStartedAt'));
const bodyMarkup = indexHtml.slice(indexHtml.indexOf('<body'), mainScriptStart)
  .replace(/<script(?:\s[^>]*)?>[\s\S]*?<\/script>/gi, '');

function elementWithId(id) {
  const match = bodyMarkup.match(new RegExp(`<([a-z]+)\\b[^>]*\\bid="${id}"[^>]*>[\\s\\S]*?<\\/\\1>`));
  assert.ok(match, `Expected element #${id}`);
  return match[0];
}

function visibleText(markup) {
  return markup.replace(/<svg\b[\s\S]*?<\/svg>/g, '').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
}

function cssBlock(selector) {
  const start = css.indexOf(selector);
  assert.notEqual(start, -1, `Expected CSS selector ${selector}`);
  const open = css.indexOf('{', start);
  let depth = 0;
  for (let index = open; index < css.length; index += 1) {
    if (css[index] === '{') depth += 1;
    if (css[index] === '}') depth -= 1;
    if (depth === 0) return css.slice(open + 1, index);
  }
  assert.fail(`Could not parse CSS selector ${selector}`);
}

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
  assert.fail(`Could not parse function ${name}`);
}

test('operation labels use the approved add and import terminology without character icons', () => {
  const buttonMarkup = Array.from(bodyMarkup.matchAll(/<button\b[^>]*>[\s\S]*?<\/button>/g), (match) => match[0]).join('\n');
  assert.doesNotMatch(buttonMarkup, /＋|📋|📷|←|×|▼|▲|⌖|↝|▧|◇/u);

  ['input-preset-add', 'pin-detail-route-add-btn'].forEach((id) => {
    assert.match(visibleText(elementWithId(id)), /追加/);
  });
  assert.match(elementWithId('pin-add-btn'), /aria-label="ピンを追加"/);
  ['csv-import-button', 'geojson-import-button', 'gpx-track-import-button', 'geojson-track-import-button'].forEach((id) => {
    assert.equal(visibleText(elementWithId(id)), '取込');
  });
  assert.equal(visibleText(elementWithId('import-preview-primary')), '取込');
  assert.equal(visibleText(elementWithId('track-import-preview-save')), '取込');
});

test('major desktop actions use an aria-hidden inline SVG followed by a visible label', () => {
  [
    'upload-submit', 'settings-save', 'input-preset-add', 'input-preset-save',
    'edit-save', 'bulk-metadata-apply', 'bulk-route-add-apply',
    'delete-confirm', 'share-create-btn', 'csv-import-button', 'geojson-import-button',
    'gpx-track-import-button', 'geojson-track-import-button', 'import-preview-primary',
    'track-import-preview-save'
  ].forEach((id) => {
    const markup = elementWithId(id);
    assert.match(markup, /<svg\b[^>]*class="[^"]*action-icon[^"]*"[^>]*aria-hidden="true"/,
      `#${id} must expose the common action icon`);
    assert.match(markup, /<span class="action-label">[^<]+<\/span>/,
      `#${id} must keep a visible action label`);
  });
});

test('mobile topbar keeps core actions as labelled 44px icon controls while desktop keeps text', () => {
  ['share-open-btn', 'data-toggle', 'more-menu-toggle'].forEach((id) => {
    const markup = elementWithId(id);
    assert.match(markup, /class="[^"]*mobile-icon-action[^"]*"/);
    assert.match(markup, /aria-label="[^"]+"/);
    assert.match(markup, /title="[^"]+"/);
    assert.match(markup, /<svg\b[^>]*aria-hidden="true"/);
    assert.match(markup, /<span class="topbar-action-label">[^<]+<\/span>/);
  });

  assert.match(cssBlock('.mobile-icon-action .topbar-action-label'), /display:\s*inline/);
  const mobileControl = cssBlock('body.narrow-view .mobile-icon-action');
  assert.match(mobileControl, /width:\s*44px/);
  assert.match(mobileControl, /min-width:\s*44px/);
  assert.match(mobileControl, /height:\s*44px/);
  assert.match(cssBlock('body.narrow-view .mobile-icon-action .topbar-action-label'), /display:\s*none/);
});

test('icon-only controls have names, titles, SVG icons, and the shared icon-button class', () => {
  [
    'panel-toggle', 'map-search-toggle', 'pin-add-btn', 'route-add-btn',
    'route-dock-toggle', 'settings-cancel', 'data-close', 'input-presets-close',
    'bulk-cancel-btn', 'bulk-metadata-btn', 'bulk-delete-btn'
  ].forEach((id) => {
    const markup = elementWithId(id);
    assert.match(markup, /class="[^"]*icon-btn[^"]*"/, `#${id} must use icon-btn`);
    assert.match(markup, /aria-label="[^"]+"/, `#${id} must have aria-label`);
    assert.match(markup, /title="[^"]+"/, `#${id} must have title`);
    assert.match(markup, /<svg\b[^>]*aria-hidden="true"/, `#${id} must use a hidden inline SVG`);
    assert.equal(visibleText(markup), '', `#${id} must be visually icon-only`);
  });
});

test('static button SVGs are decorative and common buttons follow the 44px standard', () => {
  for (const match of bodyMarkup.matchAll(/<button\b[^>]*>[\s\S]*?<\/button>/g)) {
    for (const svg of match[0].matchAll(/<svg\b[^>]*>/g)) {
      assert.match(svg[0], /aria-hidden="true"/, `Button SVG must be decorative: ${match[0].slice(0, 100)}`);
    }
  }

  const common = cssBlock('.icon-btn, .text-btn, .ghost-btn, .danger-btn, .btn-primary');
  assert.match(common, /min-height:\s*44px/);
  assert.match(common, /border-radius:\s*10px/);
  assert.match(common, /display:\s*inline-flex/);
  assert.match(common, /align-items:\s*center/);
  assert.match(common, /justify-content:\s*center/);
  assert.match(common, /gap:\s*(?:6|7|8)px/);
  assert.match(common, /white-space:\s*nowrap/);
  assert.match(cssBlock('.action-icon'), /width:\s*(?:18|19|20)px/);
  assert.match(cssBlock('.action-icon'), /height:\s*(?:18|19|20)px/);

  const iconButton = cssBlock('\n    .icon-btn {');
  assert.match(iconButton, /width:\s*44px/);
  assert.match(iconButton, /height:\s*44px/);
  assert.match(iconButton, /padding:\s*0/);
});

test('edit, duplicate, share, delete, reorder, and sort reuse semantic SVG icons', () => {
  [
    ['pin-menu-edit', 'action-icon-edit'],
    ['pin-menu-duplicate-unplaced', 'action-icon-duplicate'],
    ['pin-menu-duplicate-same', 'action-icon-duplicate'],
    ['pin-menu-duplicate-point', 'action-icon-duplicate'],
    ['pin-detail-share', 'action-icon-share'],
    ['pin-menu-delete', 'action-icon-delete'],
    ['share-qr-copy', 'action-icon-copy']
  ].forEach(([id, iconClass]) => {
    const markup = elementWithId(id);
    assert.match(markup, new RegExp(`<svg\\b[^>]*class="[^"]*${iconClass}[^"]*"[^>]*aria-hidden="true"`));
    assert.match(markup, /<span class="action-label">[^<]+<\/span>/);
  });

  const presetRenderer = functionSource('renderInputPresetList');
  assert.doesNotMatch(presetRenderer, /dragHandle\.textContent\s*=\s*'⠿'/);
  assert.match(presetRenderer, /action-icon-sort/);
  const presetButtonFactory = functionSource('createInputPresetActionButton');
  [
    'action-icon-edit', 'action-icon-duplicate', 'action-icon-enable', 'action-icon-disable',
    'action-icon-delete', 'action-icon-up', 'action-icon-down'
  ]
    .forEach((iconClass) => assert.match(presetButtonFactory, new RegExp(iconClass)));
  assert.match(presetButtonFactory, /button\.title = label/);
  assert.match(presetButtonFactory, /button\.setAttribute\('aria-label', label\)/);

  const shareRenderer = functionSource('renderShareLinks');
  ['action-icon-copy', 'action-icon-qr', 'action-icon-delete']
    .forEach((iconClass) => assert.match(shareRenderer, new RegExp(iconClass)));

  const routeRenderer = functionSource('buildRouteItem');
  const trackRenderer = functionSource('buildUnifiedTrackRouteItem');
  assert.match(routeRenderer, /createActionIconElement\(document, 'delete'\)/);
  assert.match(trackRenderer, /createActionIconElement\(document, 'settings'\)/);
  assert.match(trackRenderer, /createActionIconElement\(document, 'delete'\)/);
});

test('important and dangerous mobile actions retain text and established event ids remain wired', () => {
  [
    'upload-submit', 'settings-save', 'edit-save', 'delete-confirm',
    'drive-photo-import-confirm', 'import-preview-primary', 'import-preview-cancel',
    'import-preview-discard', 'track-import-preview-save', 'track-import-preview-discard',
    'app-confirmation-confirm', 'app-confirmation-cancel'
  ].forEach((id) => assert.notEqual(visibleText(elementWithId(id)), '', `#${id} must retain its text label`));

  [
    'share-open-btn', 'data-toggle', 'more-menu-toggle', 'panel-toggle', 'pin-add-btn',
    'upload-submit', 'settings-save', 'input-preset-add', 'edit-save', 'share-create-btn'
  ].forEach((id) => assert.equal((bodyMarkup.match(new RegExp(`\\bid="${id}"`, 'g')) || []).length, 1));

  ['share-open-btn', 'data-toggle', 'more-menu-toggle', 'panel-toggle', 'pin-add-btn'].forEach((id) => {
    assert.match(indexHtml, new RegExp(`getElementById\\('${id}'\\)\\.addEventListener\\('click'`), `#${id} click wiring changed`);
  });
});

test('dynamic label updates preserve SVG markup and remain compatible with isolated UI harnesses', () => {
  const helper = functionSource('setActionButtonLabel');
  assert.match(helper, /typeof button\.querySelector === 'function'/);
  assert.match(helper, /querySelector\('\.action-label'\)/);
  assert.match(helper, /labelNode\.textContent/);
  assert.match(helper, /else button\.textContent/);

  [
    'renderCsvInterchangeBusy', 'renderGeoJsonInterchangeBusy', 'renderTrackImportBusy',
    'updateInputPresetEditorControls', 'saveAppSettings',
    'clearUploadPhotoState', 'handleUploadPhotoSelected', 'renderShareFilterUi'
  ].forEach((name) => assert.match(functionSource(name), /setActionButtonLabel/));
});
