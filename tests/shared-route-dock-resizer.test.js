const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const sharedHtml = fs.readFileSync(path.resolve(__dirname, '..', 'shared.html'), 'utf8');
const css = sharedHtml.match(/<style>([\s\S]*?)<\/style>/)[1];
const mainScriptStart = sharedHtml.lastIndexOf('<script>', sharedHtml.indexOf('const SHARED_DEFAULT_COLOR'));
const body = sharedHtml.slice(sharedHtml.indexOf('<body'), mainScriptStart)
  .replace(/<script(?:\s[^>]*)?>[\s\S]*?<\/script>/gi, '');

function functionSource(name) {
  const start = sharedHtml.indexOf(`function ${name}(`);
  assert.notEqual(start, -1, `Expected function ${name}`);
  const open = sharedHtml.indexOf('{', start);
  let depth = 0;
  for (let index = open; index < sharedHtml.length; index += 1) {
    if (sharedHtml[index] === '{') depth += 1;
    if (sharedHtml[index] === '}') depth -= 1;
    if (depth === 0) return sharedHtml.slice(start, index + 1);
  }
  assert.fail(`Could not parse ${name}`);
}

function ruleBody(selector, predicate = () => true) {
  const normalizedCss = css.replace(/\r\n?/g, '\n');
  const normalizedSelector = selector.replace(/\r\n?/g, '\n');
  const escaped = normalizedSelector.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const matches = Array.from(normalizedCss.matchAll(new RegExp(`${escaped}\\s*\\{([^}]*)\\}`, 'g')), (match) => match[1]);
  const match = matches.find(predicate);
  assert.ok(match, `Expected CSS rule ${selector}`);
  return match;
}

function classList(initial = []) {
  const values = new Set(initial);
  return {
    add(...names) { names.forEach((name) => values.add(name)); },
    remove(...names) { names.forEach((name) => values.delete(name)); },
    toggle(name, force) {
      if (typeof force === 'boolean') force ? values.add(name) : values.delete(name);
      else if (values.has(name)) values.delete(name); else values.add(name);
      return values.has(name);
    },
    contains(name) { return values.has(name); }
  };
}

function createStyle() {
  const values = new Map();
  return {
    setProperty(name, value) { values.set(name, value); },
    removeProperty(name) { values.delete(name); },
    getPropertyValue(name) { return values.get(name) || ''; },
    values
  };
}

function createHarness(options = {}) {
  const listeners = new Map();
  const captured = [];
  const released = [];
  const panel = {
    classList: classList(['shared-route-dock-expanded']),
    style: createStyle(),
    getBoundingClientRect() { return { height: 536 }; }
  };
  const resizer = {
    classList: classList(), attributes: {},
    getBoundingClientRect() { return { height: 12 }; },
    setAttribute(name, value) { this.attributes[name] = String(value); },
    addEventListener(name, listener) { listeners.set(name, listener); },
    removeEventListener(name, listener) {
      if (listeners.get(name) === listener) listeners.delete(name);
    },
    setPointerCapture(pointerId) {
      if (options.captureFails) throw new Error('capture unavailable');
      captured.push(pointerId);
    },
    releasePointerCapture(pointerId) { released.push(pointerId); }
  };
  const documentBody = { classList: classList() };
  const elements = {
    'shared-side-panel': panel,
    'shared-route-resizer': resizer
  };
  const stored = new Map([
    ['dropThePin.routeDockSplitPx.v1', '287']
  ]);
  const storage = {
    getItem(key) { return stored.has(key) ? stored.get(key) : null; },
    setItem(key, value) { stored.set(key, String(value)); }
  };
  const context = {
    SHARED_ROUTE_DOCK_SPLIT_STORAGE_KEY: 'dropThePin.sharedRouteDockSplitPx.v1',
    state: {
      narrowView: false,
      panelVisible: true,
      routeDockExpanded: true,
      routeDockPinHeightPx: 320,
      routeDockResize: null
    },
    sharedDesktopViewMedia: { matches: true },
    document: { body: documentBody, getElementById(id) { return elements[id] || null; } },
    window: {
      localStorage: storage,
      getComputedStyle() {
        return {
          getPropertyValue(name) {
            return ({
              '--dock-pin-region-min-height': '200px',
              '--dock-pin-region-preferred-height': '64%',
              '--dock-pin-region-max-height': '450px',
              '--dock-route-region-min-height': '160px',
              '--dock-route-resizer-size': '12px'
            })[name] || '';
          }
        };
      },
      addEventListener(name, listener) { listeners.set(`window:${name}`, listener); },
      removeEventListener(name, listener) {
        if (listeners.get(`window:${name}`) === listener) listeners.delete(`window:${name}`);
      }
    },
    Math,
    Number
  };
  [
    'getSharedRouteDockCssNumber', 'getSharedRouteDockSplitBounds', 'clampSharedRouteDockPinHeight',
    'readSharedRouteDockStoredPinHeight', 'persistSharedRouteDockPinHeight', 'isSharedDesktopRouteDock',
    'getSharedRouteDockSplitMetrics', 'updateSharedRouteDockResizerAria', 'setSharedRouteDockPinHeight',
    'syncSharedRouteDockSplit', 'handleSharedRouteDockResizeMove', 'finishSharedRouteDockResize',
    'cancelSharedRouteDockResize', 'beginSharedRouteDockResize', 'handleSharedRouteDockResizerKeydown'
  ].forEach((name) => {
    context[name] = vm.runInNewContext(`(${functionSource(name)})`, context);
  });
  return { context, panel, resizer, listeners, captured, released, storage, stored };
}

function pointerEvent(overrides = {}) {
  return Object.assign({
    pointerId: 11, pointerType: 'mouse', button: 0, isPrimary: true, clientY: 200,
    preventDefault() {}, stopPropagation() {}
  }, overrides);
}

test('desktop shared route expansion owns one accessible separator and mobile hides it', () => {
  assert.match(
    sharedHtml,
    /id="shared-route-resizer"[^>]*role="separator"[^>]*aria-orientation="horizontal"[^>]*tabindex="0"[^>]*aria-label="ピン領域とルート領域の高さを調整"[^>]*aria-valuemin="0"[^>]*aria-valuemax="0"[^>]*aria-valuenow="0"/
  );
  assert.match(css, /#shared-route-resizer\s*\{[^}]*display:\s*none/);
  assert.match(css, /@media \(min-width:\s*900px\)[\s\S]*#shared-side-panel\.shared-route-dock-expanded #shared-route-resizer\s*\{[^}]*display:\s*block/);
  const pinIndex = body.indexOf('id="shared-pin-section"');
  const resizerIndex = body.indexOf('id="shared-route-resizer"');
  const routeIndex = body.indexOf('id="shared-route-section"');
  assert.ok(pinIndex < resizerIndex && resizerIndex < routeIndex, 'separator stays on the pin/route boundary');
  const mobile = css.slice(css.indexOf('@media (max-width: 640px)'));
  assert.match(mobile, /#shared-route-resizer[^}]*display:\s*none\s*!important/);
});

test('subdesktop expansion keeps two usable rows and drag disables grid track animation', () => {
  const baseExpanded = ruleBody('#shared-side-panel.shared-route-dock-expanded');
  assert.doesNotMatch(baseExpanded, /--dock-route-resizer-size/);
  assert.match(baseExpanded, /--dock-pin-region-preferred-height/);
  assert.match(baseExpanded, /--dock-route-region-min-height/);
  assert.match(
    css,
    /@media \(min-width:\s*900px\)[\s\S]*#shared-side-panel\.shared-route-dock-expanded\s*\{[^}]*--dock-route-resizer-size[^}]*--shared-dock-route-region-effective-min-height/
  );
  const dragging = ruleBody('body.shared-route-dock-resizing #shared-side-panel');
  assert.doesNotMatch(dragging, /grid-template-rows/);
});

test('shared split bounds preserve minimum pin and route heights on normal and short panels', () => {
  const bounds = vm.runInNewContext(`(${functionSource('getSharedRouteDockSplitBounds')})`);
  assert.deepEqual({ ...bounds(536, 12, 200, 160) }, { min: 200, max: 364 });
  const low = bounds(120, 12, 200, 160);
  assert.ok(low.min > 0);
  assert.ok(low.max >= low.min);
  assert.ok(108 - low.max > 0, 'route region stays positive');
});

test('shared storage uses its own key, rejects invalid values, and tolerates storage errors', () => {
  assert.match(sharedHtml, /dropThePin\.sharedRouteDockSplitPx\.v1/);
  const h = createHarness();
  assert.equal(h.context.readSharedRouteDockStoredPinHeight(h.storage), null);
  h.stored.set('dropThePin.sharedRouteDockSplitPx.v1', '312.5');
  assert.equal(h.context.readSharedRouteDockStoredPinHeight(h.storage), 312.5);
  for (const value of ['', '-1', 'NaN', 'Infinity', '0']) {
    h.stored.set('dropThePin.sharedRouteDockSplitPx.v1', value);
    assert.equal(h.context.readSharedRouteDockStoredPinHeight(h.storage), null, value);
  }
  assert.equal(h.context.readSharedRouteDockStoredPinHeight({ getItem() { throw new Error('blocked'); } }), null);
  assert.doesNotThrow(() => h.context.persistSharedRouteDockPinHeight(300, {
    setItem() { throw new Error('blocked'); }
  }));
  h.context.persistSharedRouteDockPinHeight(300, h.storage);
  assert.equal(h.stored.get('dropThePin.sharedRouteDockSplitPx.v1'), '300');
  assert.equal(h.stored.get('dropThePin.routeDockSplitPx.v1'), '287', 'editing split remains untouched');
});

test('shared pointer drag captures, clamps, persists, releases, and removes transient listeners', () => {
  const h = createHarness();
  assert.equal(h.context.beginSharedRouteDockResize(pointerEvent()), true);
  assert.deepEqual(h.captured, [11]);
  assert.ok(h.listeners.has('pointermove'));
  assert.ok(h.listeners.has('pointerup'));
  assert.ok(h.listeners.has('pointercancel'));

  h.context.handleSharedRouteDockResizeMove(pointerEvent({ clientY: 260 }));
  assert.equal(h.context.state.routeDockPinHeightPx, 364);
  assert.equal(h.resizer.attributes['aria-valuenow'], '364');

  h.context.finishSharedRouteDockResize(pointerEvent({ clientY: 260 }), false);
  assert.deepEqual(h.released, [11]);
  assert.equal(h.context.state.routeDockResize, null);
  assert.equal(h.listeners.has('pointermove'), false);
  assert.equal(h.listeners.has('pointerup'), false);
  assert.equal(h.listeners.has('pointercancel'), false);
  assert.equal(h.stored.get('dropThePin.sharedRouteDockSplitPx.v1'), '364');
});

test('pointercancel restores the shared drag start and a viewport shrink reclamps saved height', () => {
  const h = createHarness();
  h.context.beginSharedRouteDockResize(pointerEvent());
  h.context.handleSharedRouteDockResizeMove(pointerEvent({ clientY: 120 }));
  assert.equal(h.context.state.routeDockPinHeightPx, 240);
  h.context.finishSharedRouteDockResize(pointerEvent({ clientY: 120 }), true);
  assert.equal(h.context.state.routeDockPinHeightPx, 320);
  assert.equal(h.stored.has('dropThePin.sharedRouteDockSplitPx.v1'), false);

  h.panel.getBoundingClientRect = () => ({ height: 430 });
  h.context.state.routeDockPinHeightPx = 400;
  h.context.syncSharedRouteDockSplit({ persist: true });
  assert.equal(h.context.state.routeDockPinHeightPx, 258);
  assert.equal(h.stored.get('dropThePin.sharedRouteDockSplitPx.v1'), '258');
});

test('capture failure leaves no shared drag state or pointer listeners', () => {
  const h = createHarness({ captureFails: true });
  assert.equal(h.context.beginSharedRouteDockResize(pointerEvent()), false);
  assert.equal(h.context.state.routeDockResize, null);
  assert.equal(h.listeners.has('pointermove'), false);
  assert.equal(h.listeners.has('pointerup'), false);
  assert.equal(h.listeners.has('pointercancel'), false);
  assert.equal(h.context.document.body.classList.contains('shared-route-dock-resizing'), false);
});

test('shared keyboard directions, Shift, Home, End, and aria values stay synchronized', () => {
  const h = createHarness();
  function press(key, shiftKey = false) {
    let prevented = 0;
    h.context.handleSharedRouteDockResizerKeydown({
      key, shiftKey,
      preventDefault() { prevented += 1; }, stopPropagation() {}
    });
    assert.equal(prevented, 1, key);
  }
  press('ArrowUp');
  assert.equal(h.context.state.routeDockPinHeightPx, 304);
  press('ArrowDown', true);
  assert.equal(h.context.state.routeDockPinHeightPx, 336);
  press('Home');
  assert.equal(h.context.state.routeDockPinHeightPx, 200);
  press('End');
  assert.equal(h.context.state.routeDockPinHeightPx, 364);
  assert.deepEqual(
    [h.resizer.attributes['aria-valuemin'], h.resizer.attributes['aria-valuemax'], h.resizer.attributes['aria-valuenow']],
    ['200', '364', '364']
  );
});

test('mobile shared removes desktop split and preserves the pin and route tab model', () => {
  const h = createHarness();
  h.context.state.narrowView = true;
  h.context.sharedDesktopViewMedia.matches = false;
  h.panel.style.setProperty('--shared-dock-pin-region-height', '320px');
  h.context.syncSharedRouteDockSplit();
  assert.equal(h.panel.style.getPropertyValue('--shared-dock-pin-region-height'), '');
  assert.equal(h.context.beginSharedRouteDockResize(pointerEvent()), false);
  assert.equal(h.context.handleSharedRouteDockResizerKeydown({
    key: 'ArrowUp', preventDefault() {}, stopPropagation() {}
  }), false);

  const tabs = functionSource('setSharedMobileSheetTab');
  assert.match(tabs, /shared-mobile-routes/);
  assert.match(tabs, /state\.mobileSheetTab === 'routes'/);
  assert.match(css, /body\.shared-mobile-routes #shared-pin-section\s*\{\s*display:\s*none/);
  assert.match(css, /body:not\(\.shared-mobile-routes\) #shared-route-section\s*\{\s*display:\s*none/);
  assert.match(css, /--mobile-route-map-min-height/);
});

test('closing shared panel and viewport changes cancel an active drag and reclamp the split', () => {
  assert.match(functionSource('renderSharedPanelUi'), /if \(!state\.panelVisible\)[\s\S]*cancelSharedRouteDockResize\(\)/);
  assert.match(sharedHtml, /addSharedMediaListener\(sharedNarrowViewMedia,[\s\S]*cancelSharedRouteDockResize\(\)/);
  assert.match(sharedHtml, /addSharedMediaListener\(sharedDesktopViewMedia,[\s\S]*cancelSharedRouteDockResize\(\)/);
  assert.match(sharedHtml, /window\.addEventListener\('resize',[\s\S]*cancelSharedRouteDockResize\(\)[\s\S]*syncSharedRouteDockSplit/);
});

test('shared pin, route, and nested route-pin lists own vertical scrolling without horizontal overflow', () => {
  const listRule = ruleBody('#shared-pin-list,\n    #shared-route-list');
  assert.match(listRule, /min-height:\s*0/);
  assert.match(listRule, /overflow-x:\s*hidden/);
  assert.match(listRule, /overflow-y:\s*auto/);
  assert.match(listRule, /overscroll-behavior:\s*contain/);
  assert.match(listRule, /touch-action:\s*pan-y/);

  const cardRule = ruleBody('.unified-route-card');
  assert.match(cardRule, /flex:\s*0 0 auto/);
  const routePinList = ruleBody('.route-pin-list');
  assert.match(routePinList, /max-height:\s*min\(42dvh, 352px\)/);
  assert.match(routePinList, /overflow-x:\s*hidden/);
  assert.match(routePinList, /overflow-y:\s*auto/);
  assert.match(routePinList, /touch-action:\s*pan-y/);

  assert.match(sharedHtml, /getElementById\('shared-pin-list'\)\.addEventListener\('wheel'/);
  assert.match(sharedHtml, /getElementById\('shared-route-list'\)\.addEventListener\('wheel'/);
});

test('shared dock stays read-only and adds no edit or import entry points', () => {
  assert.doesNotMatch(body, /id="shared-route-add/);
  assert.doesNotMatch(body, /id="shared-route-import/);
  assert.doesNotMatch(body, /id="shared-(?:route-name-edit|route-undo|route-detail-settings)/);
  assert.doesNotMatch(sharedHtml, /dropThePin\.routeDockSplitPx\.v1/);
});
