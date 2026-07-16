const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.resolve(__dirname, '..', 'index.html'), 'utf8');

function countId(id) {
  const escaped = id.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return (indexHtml.match(new RegExp(`\\bid=["']${escaped}["']`, 'g')) || []).length;
}

function functionBody(name) {
  const start = indexHtml.indexOf(`function ${name}`);
  assert.notEqual(start, -1, `Expected function ${name}`);
  const open = indexHtml.indexOf('{', start);
  let depth = 0;
  for (let index = open; index < indexHtml.length; index += 1) {
    if (indexHtml[index] === '{') depth += 1;
    if (indexHtml[index] === '}') depth -= 1;
    if (depth === 0) return indexHtml.slice(open + 1, index);
  }
  assert.fail(`Could not parse ${name}`);
}

test('right dock uses a slimmer bounded width with stable pin containers', () => {
  assert.match(indexHtml, /--app-dock-width:\s*clamp\(340px,\s*26\.35vw,\s*360px\)/);
  ['side-panel', 'dock-pin-region', 'dock-route-region', 'side-placed', 'side-unplaced']
    .forEach((id) => assert.equal(countId(id), 1, `${id} must exist exactly once`));

  const dock = indexHtml.match(/<aside id="side-panel"[^>]*>([\s\S]*?)<\/aside>/);
  assert.ok(dock);
  assert.ok(dock[1].indexOf('id="dock-pin-region"') < dock[1].indexOf('id="dock-route-region"'));
  assert.match(indexHtml, /#side-panel\s*\{[\s\S]*?width:\s*min\(var\(--app-dock-width\),\s*100%\)/);
  assert.match(indexHtml, /#side-panel\s*\{[\s\S]*?grid-template-rows:\s*minmax\(0,\s*1fr\)/);
});

test('placed and unplaced tabs preserve IDs and data-pin-id card behavior', () => {
  ['pin-tab-placed', 'pin-tab-unplaced', 'side-placed-count', 'side-unplaced-count']
    .forEach((id) => assert.equal(countId(id), 1, `${id} must exist exactly once`));
  assert.match(indexHtml, /id="pin-tab-placed"[^>]*role="tab"[^>]*aria-controls="side-placed"/);
  assert.match(indexHtml, /id="pin-tab-unplaced"[^>]*role="tab"[^>]*aria-controls="side-unplaced"/);

  const build = functionBody('buildListItem');
  assert.match(build, /button\.className\s*=\s*'list-item'/);
  assert.match(build, /button\.dataset\.pinId\s*=\s*pin\.id/);

  const render = functionBody('renderSidePanel');
  assert.match(render, /getElementById\('side-unplaced'\)/);
  assert.match(render, /getElementById\('side-placed'\)/);
  assert.match(render, /buildListItem/);

  const tabs = functionBody('renderPinDockTabs');
  assert.match(tabs, /aria-selected/);
  assert.match(tabs, /placedContainer\.hidden/);
  assert.match(tabs, /unplacedContainer\.hidden/);
});

test('tab switch keeps the existing placed and unplaced drag paths reachable', () => {
  const dnd = functionBody('setupDndDropTargets');
  assert.match(dnd, /getElementById\('pin-tab-unplaced'\)/);
  assert.match(dnd, /getElementById\('pin-tab-placed'\)/);
  assert.match(dnd, /withGAS\('unplacePin',\s*withEditToken/);
  assert.match(dnd, /withGAS\('movePin',\s*withEditToken/);
  assert.match(dnd, /setPinDockTab\('unplaced'\)/);
  assert.match(dnd, /setPinDockTab\('placed'\)/);
});

test('one compact bookmark tab exposes its state and keeps a 44px height', () => {
  assert.equal(countId('panel-toggle'), 1);
  assert.match(indexHtml, /id="panel-toggle"[^>]*aria-controls="side-panel"[^>]*aria-expanded="true"/);
  assert.match(indexHtml, /#panel-toggle\s*\{[\s\S]*?min-width:\s*44px[\s\S]*?min-height:\s*44px/);
  assert.match(indexHtml, /body\.panel-visible #panel-toggle\s*\{[\s\S]*?right:\s*var\(--app-dock-width\)/);
  assert.match(indexHtml, /body\.panel-hidden #panel-toggle\s*\{[\s\S]*?right:\s*0/);
  assert.match(indexHtml, /@media not all and \(min-width:\s*900px\)[\s\S]*?body\.panel-visible #panel-toggle\s*\{\s*right:\s*min\(var\(--app-dock-width\),\s*calc\(100vw - 44px\)\)/);

  const render = functionBody('renderPanelToggle');
  assert.match(render, /setAttribute\('aria-expanded'/);
  assert.match(render, /一覧を閉じる/);
  assert.match(render, /一覧を開く/);
  assert.doesNotMatch(render, /textContent/);
  assert.match(indexHtml, /class="panel-toggle-icon panel-toggle-icon-close"/);
  assert.match(indexHtml, /class="panel-toggle-icon panel-toggle-icon-open"/);
  assert.match(indexHtml, /#panel-toggle\[aria-expanded="false"\] \.panel-toggle-icon-open\s*\{\s*display:\s*block/);
});

test('panel resize invalidates Leaflet after transition and under reduced motion', () => {
  const body = functionBody('invalidateMapAfterPanelTransition');
  assert.match(body, /prefers-reduced-motion:\s*reduce/);
  assert.match(body, /transitionend/);
  assert.match(body, /map\.invalidateSize\(\)/);
  assert.match(body, /requestAnimationFrame/);
  assert.match(body, /setTimeout/);
});

test('panel resize waits for right transition once and reduced motion uses the next frame', () => {
  function createHarness(reduceMotion) {
    const listeners = {};
    const timers = [];
    const frames = [];
    let invalidations = 0;
    const mapElement = {
      addEventListener(type, listener) { listeners[type] = listener; },
      removeEventListener(type, listener) {
        if (listeners[type] === listener) delete listeners[type];
      }
    };
    const run = vm.runInNewContext(`(function() {${functionBody('invalidateMapAfterPanelTransition')}})`, {
      document: { getElementById: () => mapElement },
      map: { invalidateSize() { invalidations += 1; } },
      window: { matchMedia: () => ({ matches: reduceMotion }) },
      requestAnimationFrame(callback) { frames.push(callback); },
      setTimeout(callback) { timers.push(callback); }
    });
    return { run, listeners, timers, frames, mapElement, invalidations: () => invalidations };
  }

  const animated = createHarness(false);
  animated.run();
  assert.equal(animated.invalidations(), 0);
  animated.listeners.transitionend({ target: {}, propertyName: 'right' });
  assert.equal(animated.invalidations(), 0);
  animated.listeners.transitionend({ target: animated.mapElement, propertyName: 'right' });
  assert.equal(animated.invalidations(), 1);
  animated.timers[0]();
  assert.equal(animated.invalidations(), 1);

  const reduced = createHarness(true);
  reduced.run();
  assert.equal(reduced.listeners.transitionend, undefined);
  assert.equal(reduced.invalidations(), 0);
  reduced.frames[0]();
  assert.equal(reduced.invalidations(), 1);
});

test('multi-selection toolbar keeps only count clear edit and danger delete icon actions', () => {
  const toolbar = indexHtml.match(/<div id="bulk-action-bar"[^>]*>([\s\S]*?)<\/div>/);
  assert.ok(toolbar);
  ['bulk-count', 'bulk-cancel-btn', 'bulk-metadata-btn', 'bulk-delete-btn']
    .forEach((id) => assert.ok(toolbar[1].includes(`id="${id}"`), `${id} must be in toolbar`));
  ['bulk-route-add-btn', 'bulk-unplace-btn', 'ルート追加', '未配置へ']
    .forEach((value) => assert.equal(toolbar[1].includes(value), false, `${value} must be removed`));
  assert.equal(toolbar[1].includes('bulk-status-select'), false, 'status editor belongs in the bulk edit dialog');
  assert.match(indexHtml, /#bulk-action-bar\s*\{[\s\S]*?display:\s*none[\s\S]*?flex-wrap:\s*nowrap/);
  ['bulk-cancel-btn', 'bulk-metadata-btn', 'bulk-delete-btn'].forEach((id) => {
    const button = toolbar[1].match(new RegExp(`<button id="${id}"[\\s\\S]*?<\\/button>`));
    assert.ok(button, id);
    assert.match(button[0], /title="[^"]+"/);
    assert.match(button[0], /aria-label="[^"]+"/);
    assert.match(button[0], /<svg\b[^>]*aria-hidden="true"/);
    assert.equal(button[0].replace(/<svg\b[\s\S]*?<\/svg>/, '').replace(/<[^>]+>/g, '').trim(), '');
  });
  assert.match(toolbar[1], /id="bulk-delete-btn" class="[^"]*danger-btn[^"]*"/);
});

test('edit actions stay hidden and guarded when edit permission is unavailable', () => {
  assert.match(indexHtml, /body:not\(\.has-edit-token\) #bulk-action-bar/);
  assert.match(indexHtml, /body:not\(\.edit-mode\) #bulk-action-bar/);
  assert.match(functionBody('requestBulkDelete'), /if \(!canEdit\(\)\) return/);
  assert.match(functionBody('applyBulkStatus'), /if \(!canEdit\(\)\) return/);
});

test('desktop detail uses the dock while the existing overlay contract remains', () => {
  assert.equal(countId('pin-detail-overlay'), 1);
  assert.match(indexHtml, /#pin-detail-overlay\.open/);
  assert.match(indexHtml, /@media \(min-width:\s*900px\)[\s\S]*?body:not\(\.panel-hidden\) #pin-detail-overlay\.open/);
  assert.match(indexHtml, /@media not all and \(min-width:\s*900px\)[\s\S]*?#pin-detail-overlay/);
  assert.match(functionBody('openPinDetail'), /openOverlay\('pin-detail-overlay'\)/);
  assert.match(functionBody('closePinDetail'), /closeOverlay\('pin-detail-overlay'/);
});
