const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

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
    classList: classList(['route-dock-expanded']), style: createStyle(),
    getBoundingClientRect() { return { height: 536 }; }
  };
  const pinRegion = { getBoundingClientRect() { return { height: 320 }; } };
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
  const body = { classList: classList() };
  const elements = {
    'side-panel': panel,
    'dock-pin-region': pinRegion,
    'dock-route-resizer': resizer
  };
  const stored = new Map();
  const storage = {
    getItem(key) { return stored.has(key) ? stored.get(key) : null; },
    setItem(key, value) { stored.set(key, String(value)); }
  };
  const context = {
    ROUTE_DOCK_SPLIT_STORAGE_KEY: 'dropThePin.routeDockSplitPx.v1',
    state: {
      narrowView: false,
      routeDockExpanded: true,
      routeDockPinHeightPx: 320,
      routeDockResize: null
    },
    desktopViewMedia: { matches: true },
    document: { body, getElementById(id) { return elements[id] || null; } },
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
    Math, Number
  };
  [
    'getRouteDockCssNumber', 'getRouteDockSplitBounds', 'clampRouteDockPinHeight',
    'readRouteDockStoredPinHeight', 'persistRouteDockPinHeight', 'isDesktopRouteDock',
    'getRouteDockSplitMetrics', 'updateRouteDockResizerAria', 'setRouteDockPinHeight',
    'syncRouteDockSplit', 'handleRouteDockResizeMove', 'finishRouteDockResize',
    'cancelRouteDockResize', 'beginRouteDockResize', 'handleRouteDockResizerKeydown'
  ].forEach((name) => {
    context[name] = vm.runInNewContext(`(${functionSource(name)})`, context);
  });
  return { context, panel, resizer, listeners, captured, released, storage, stored };
}

function pointerEvent(overrides = {}) {
  return Object.assign({
    pointerId: 7, pointerType: 'mouse', button: 0, isPrimary: true, clientY: 200,
    preventDefault() {}, stopPropagation() {}
  }, overrides);
}

test('expanded desktop dock owns one accessible horizontal separator and mobile hides it', () => {
  assert.match(
    indexHtml,
    /id="dock-route-resizer"[^>]*role="separator"[^>]*aria-orientation="horizontal"[^>]*tabindex="0"[^>]*aria-label="ピン領域とルート領域の高さを調整"/
  );
  assert.match(css, /#dock-route-resizer\s*\{[^}]*display:\s*none/);
  assert.match(css, /@media \(min-width:\s*900px\)[\s\S]*#side-panel\.route-dock-expanded #dock-route-resizer\s*\{[^}]*display:\s*block/);
  const mobile = css.slice(css.indexOf('@media (max-width: 640px)'));
  assert.match(mobile, /#dock-route-resizer[^}]*display:\s*none\s*!important/);
});

test('split bounds preserve both configured minimums and degrade to positive tracks on very low panels', () => {
  const bounds = vm.runInNewContext(`(${functionSource('getRouteDockSplitBounds')})`);
  assert.deepEqual({ ...bounds(536, 12, 200, 160) }, { min: 200, max: 364 });
  const low = bounds(120, 12, 200, 160);
  assert.ok(low.min > 0);
  assert.ok(low.max >= low.min);
  assert.ok(108 - low.max > 0, 'route region stays positive');
});

test('storage accepts only finite positive values and storage failures are harmless', () => {
  const read = vm.runInNewContext(`(${functionSource('readRouteDockStoredPinHeight')})`, {
    ROUTE_DOCK_SPLIT_STORAGE_KEY: 'dropThePin.routeDockSplitPx.v1', Number
  });
  assert.equal(read({ getItem() { return '312.5'; } }), 312.5);
  for (const value of [null, '', '-1', 'NaN', 'Infinity', '0']) {
    assert.equal(read({ getItem() { return value; } }), null, String(value));
  }
  assert.equal(read({ getItem() { throw new Error('blocked'); } }), null);

  const persist = vm.runInNewContext(`(${functionSource('persistRouteDockPinHeight')})`, {
    ROUTE_DOCK_SPLIT_STORAGE_KEY: 'dropThePin.routeDockSplitPx.v1', window: {
      localStorage: { setItem() { throw new Error('blocked'); } }
    }, Number
  });
  assert.doesNotThrow(() => persist(300));
});

test('pointer drag captures, clamps, persists, releases, and removes transient listeners', () => {
  const h = createHarness();
  assert.equal(h.context.beginRouteDockResize(pointerEvent()), true);
  assert.deepEqual(h.captured, [7]);
  assert.ok(h.listeners.has('pointermove'));
  assert.ok(h.listeners.has('pointerup'));
  assert.ok(h.listeners.has('pointercancel'));

  h.context.handleRouteDockResizeMove(pointerEvent({ clientY: 260 }));
  assert.equal(h.context.state.routeDockPinHeightPx, 364, 'route minimum clamps downward drag');
  assert.equal(h.resizer.attributes['aria-valuenow'], '364');

  h.context.finishRouteDockResize(pointerEvent({ clientY: 260 }), false);
  assert.deepEqual(h.released, [7]);
  assert.equal(h.context.state.routeDockResize, null);
  assert.equal(h.listeners.has('pointermove'), false);
  assert.equal(h.listeners.has('pointerup'), false);
  assert.equal(h.listeners.has('pointercancel'), false);
  assert.equal(h.stored.get('dropThePin.routeDockSplitPx.v1'), '364');
});

test('pointercancel restores the drag start and window shrink reclamps a saved height', () => {
  const h = createHarness();
  h.context.beginRouteDockResize(pointerEvent());
  h.context.handleRouteDockResizeMove(pointerEvent({ clientY: 120 }));
  assert.equal(h.context.state.routeDockPinHeightPx, 240);
  h.context.finishRouteDockResize(pointerEvent({ clientY: 120 }), true);
  assert.equal(h.context.state.routeDockPinHeightPx, 320);
  assert.equal(h.stored.has('dropThePin.routeDockSplitPx.v1'), false);

  h.panel.getBoundingClientRect = () => ({ height: 430 });
  h.context.state.routeDockPinHeightPx = 400;
  h.context.syncRouteDockSplit({ persist: true });
  assert.equal(h.context.state.routeDockPinHeightPx, 258);
  assert.equal(h.stored.get('dropThePin.routeDockSplitPx.v1'), '258');
});

test('capture failure aborts the resize without leaving drag state or listeners', () => {
  const h = createHarness({ captureFails: true });
  assert.equal(h.context.beginRouteDockResize(pointerEvent()), false);
  assert.equal(h.context.state.routeDockResize, null);
  assert.equal(h.listeners.has('pointermove'), false);
  assert.equal(h.listeners.has('pointerup'), false);
  assert.equal(h.listeners.has('pointercancel'), false);
  assert.equal(h.context.document.body.classList.contains('route-dock-resizing'), false);
});

test('keyboard directions, Shift, Home, End, and aria values remain synchronized', () => {
  const h = createHarness();
  function press(key, shiftKey = false) {
    let prevented = 0;
    h.context.handleRouteDockResizerKeydown({
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

test('mobile mode removes the applied split and refuses pointer or keyboard resizing', () => {
  const h = createHarness();
  h.context.state.narrowView = true;
  h.context.desktopViewMedia.matches = false;
  h.panel.style.setProperty('--dock-pin-region-height', '320px');
  h.context.syncRouteDockSplit();
  assert.equal(h.panel.style.getPropertyValue('--dock-pin-region-height'), '');
  assert.equal(h.context.beginRouteDockResize(pointerEvent()), false);
  let prevented = 0;
  assert.equal(h.context.handleRouteDockResizerKeydown({
    key: 'ArrowUp', preventDefault() { prevented += 1; }, stopPropagation() {}
  }), false);
  assert.equal(prevented, 0);
});

test('dock closing and access-mode transitions cancel an active resize', () => {
  assert.match(functionSource('renderPanelToggle'), /if \(!isPanelVisible\)[\s\S]*cancelRouteDockResize\(\)/);
  assert.match(functionSource('renderAccessMode'), /if \(!canEdit\(\)\)[\s\S]*cancelRouteDockResize\(\)/);
});
