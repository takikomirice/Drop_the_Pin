const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const sharedHtml = fs.readFileSync(path.join(root, 'shared.html'), 'utf8');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const sharedCss = sharedHtml.match(/<style>([\s\S]*?)<\/style>/)[1];

function functionSource(source, name) {
  const start = source.indexOf(`function ${name}`);
  assert.notEqual(start, -1, `Expected function ${name}`);
  const open = source.indexOf('{', start);
  let depth = 0;
  for (let index = open; index < source.length; index += 1) {
    if (source[index] === '{') depth += 1;
    if (source[index] === '}') depth -= 1;
    if (depth === 0) return source.slice(start, index + 1);
  }
  assert.fail(`Could not parse function ${name}`);
}

function listenerSource(id) {
  const marker = `document.getElementById('${id}').addEventListener('click', `;
  const markerStart = sharedHtml.indexOf(marker);
  assert.notEqual(markerStart, -1, `Expected click listener for #${id}`);
  const functionStart = sharedHtml.indexOf('function', markerStart + marker.length);
  const open = sharedHtml.indexOf('{', functionStart);
  let depth = 0;
  for (let index = open; index < sharedHtml.length; index += 1) {
    if (sharedHtml[index] === '{') depth += 1;
    if (sharedHtml[index] === '}') depth -= 1;
    if (depth === 0) return sharedHtml.slice(functionStart, index + 1);
  }
  assert.fail(`Could not parse click listener for #${id}`);
}

function element(id, { display = '', scrollTop = 0 } = {}) {
  const attributes = Object.create(null);
  return {
    id,
    hidden: false,
    className: '',
    style: { display, visibility: '', opacity: '' },
    value: '',
    textContent: '',
    scrollTop,
    setAttribute(name, value) { attributes[name] = String(value); },
    getAttribute(name) { return Object.hasOwn(attributes, name) ? attributes[name] : null; }
  };
}

function createHarness(overrides = {}) {
  const elements = {
    'shared-map-search-bar': element('shared-map-search-bar', { scrollTop: overrides.scrollTop || 0 }),
    'shared-search-input': element('shared-search-input'),
    'shared-search-expanded': element('shared-search-expanded', { display: 'none' }),
    'shared-search-toggle': element('shared-search-toggle'),
    'shared-search-controls': element('shared-search-controls'),
    'shared-sort': element('shared-sort'),
    'shared-filter-tag-btn': element('shared-filter-tag-btn'),
    'shared-filter-tag-detail': element('shared-filter-tag-detail', { display: 'none' }),
    'shared-filter-color-btn': element('shared-filter-color-btn'),
    'shared-filter-color-detail': element('shared-filter-color-detail', { display: 'none' }),
    'shared-clear-filters': element('shared-clear-filters'),
    'shared-reset-routes': element('shared-reset-routes', { display: overrides.resetDisplay || '' }),
    'shared-tag-mode': element('shared-tag-mode')
  };
  const calls = { tags: 0, colors: 0, pins: 0, map: 0 };
  const state = {
    listQuery: overrides.listQuery || 'query',
    activeTags: (overrides.activeTags || []).slice(),
    activeColors: (overrides.activeColors || []).slice(),
    tagMode: overrides.tagMode || 'and',
    initialTagMode: overrides.initialTagMode || 'or',
    sort: 'newest',
    searchExpanded: overrides.searchExpanded !== false,
    tagFilterExpanded: overrides.tagFilterExpanded === true,
    colorFilterExpanded: overrides.colorFilterExpanded === true
  };
  elements['shared-search-input'].value = state.listQuery;
  const context = {
    state,
    document: { getElementById(id) { return elements[id] || null; } },
    renderSharedTagChips() { calls.tags += 1; },
    renderSharedColorFilter() { calls.colors += 1; },
    renderSharedPins() { calls.pins += 1; },
    renderSharedMap() { calls.map += 1; }
  };
  vm.runInNewContext([
    functionSource(sharedHtml, 'renderSharedSearchUi'),
    functionSource(sharedHtml, 'clearSharedFilters'),
    'this.renderSearch = renderSharedSearchUi;',
    'this.clearFilters = clearSharedFilters;'
  ].join('\n'), context);
  return {
    state,
    elements,
    calls,
    render: context.renderSearch,
    clear: context.clearFilters,
    click(id) {
      const listener = vm.runInNewContext(`(${listenerSource(id)})`, context);
      listener.call(elements[id]);
    }
  };
}

function isActive(elementValue) {
  return elementValue.className.split(/\s+/).includes('active');
}

function computedButtonStyle(elementValue) {
  return {
    display: elementValue.style.display || 'inline-flex',
    visibility: elementValue.style.visibility || 'visible',
    opacity: elementValue.style.opacity || '1'
  };
}

function buttonRect(searchBar, topAtZero, height = 44) {
  return { top: topAtZero - searchBar.scrollTop, bottom: topAtZero - searchBar.scrollTop + height };
}

function isVerticallyClipped(rect, viewport) {
  return rect.top < viewport.top || rect.bottom > viewport.bottom;
}

test('clearing a selected tag preserves every control and exposes it above the scroll clip', () => {
  const harness = createHarness({
    activeTags: ['food'],
    tagFilterExpanded: true,
    scrollTop: 180,
    resetDisplay: 'none'
  });
  const tagButton = harness.elements['shared-filter-tag-btn'];
  const searchBar = harness.elements['shared-map-search-bar'];
  const viewport = { top: 12, bottom: 344 };

  harness.render();
  assert.equal(isVerticallyClipped(buttonRect(searchBar, 112), viewport), true, 'fixture must reproduce scroll clipping');
  harness.clear();

  assert.strictEqual(harness.elements['shared-filter-tag-btn'], tagButton, 'the control node must not be rebuilt');
  assert.equal(tagButton.hidden, false);
  assert.deepEqual(computedButtonStyle(tagButton), { display: 'inline-flex', visibility: 'visible', opacity: '1' });
  assert.equal(isActive(tagButton), false);
  assert.equal(tagButton.getAttribute('aria-pressed'), 'false');
  assert.equal(tagButton.getAttribute('aria-expanded'), 'true');
  assert.equal(harness.elements['shared-search-expanded'].style.display, '');
  assert.equal(harness.elements['shared-filter-tag-detail'].style.display, '');
  assert.equal(searchBar.scrollTop, 0);
  assert.equal(isVerticallyClipped(buttonRect(searchBar, 112), viewport), false);
  assert.equal(harness.elements['shared-filter-color-btn'].hidden, false);
  assert.equal(harness.elements['shared-clear-filters'].hidden, false);
  assert.equal(harness.elements['shared-reset-routes'].style.display, 'none', 'route reset visibility remains route-owned');
  assert.equal(harness.state.activeTags.length, 0);
  assert.equal(harness.state.listQuery, '');
  assert.equal(harness.state.tagMode, 'or');
  assert.deepEqual(harness.calls, { tags: 1, colors: 1, pins: 1, map: 1 });
});

test('clearing a selected color removes its filter indicator without closing its detail', () => {
  const harness = createHarness({ activeColors: ['#43a047'], colorFilterExpanded: true, scrollTop: 96 });
  const colorButton = harness.elements['shared-filter-color-btn'];

  harness.clear();

  assert.strictEqual(harness.elements['shared-filter-color-btn'], colorButton);
  assert.equal(colorButton.hidden, false);
  assert.equal(isActive(colorButton), false);
  assert.equal(colorButton.getAttribute('aria-pressed'), 'false');
  assert.equal(colorButton.getAttribute('aria-expanded'), 'true');
  assert.equal(harness.elements['shared-filter-color-detail'].style.display, '');
  assert.equal(harness.state.activeColors.length, 0);
  assert.equal(harness.elements['shared-map-search-bar'].scrollTop, 0);
});

test('clear follows index by preserving an open tag detail independently of filter selection', () => {
  const indexClear = functionSource(indexHtml, 'clearMapSearchFilters');
  assert.doesNotMatch(indexClear, /map-filter-(?:tag|color)-detail/);
  const harness = createHarness({ tagFilterExpanded: true, activeTags: ['food'] });

  harness.clear();

  assert.equal(harness.state.tagFilterExpanded, true);
  assert.equal(harness.state.colorFilterExpanded, false);
  assert.equal(harness.elements['shared-filter-tag-detail'].style.display, '');
  assert.equal(harness.elements['shared-filter-tag-btn'].getAttribute('aria-expanded'), 'true');
  assert.equal(harness.elements['shared-filter-tag-btn'].getAttribute('aria-pressed'), 'false');
});

test('clear follows index by preserving an open color detail independently of filter selection', () => {
  const harness = createHarness({ colorFilterExpanded: true, activeColors: ['#43a047'] });

  harness.clear();

  assert.equal(harness.state.colorFilterExpanded, true);
  assert.equal(harness.state.tagFilterExpanded, false);
  assert.equal(harness.elements['shared-filter-color-detail'].style.display, '');
  assert.equal(harness.elements['shared-filter-color-btn'].getAttribute('aria-expanded'), 'true');
  assert.equal(harness.elements['shared-filter-color-btn'].getAttribute('aria-pressed'), 'false');
});

test('tag and color detail buttons follow index mutual exclusion and clear leaves all operations intact', () => {
  const harness = createHarness({ activeTags: ['food'], activeColors: ['#43a047'] });

  harness.click('shared-filter-tag-btn');
  assert.equal(harness.state.tagFilterExpanded, true);
  assert.equal(harness.state.colorFilterExpanded, false);
  harness.click('shared-filter-color-btn');
  assert.equal(harness.state.tagFilterExpanded, false);
  assert.equal(harness.state.colorFilterExpanded, true);
  harness.clear();

  for (const id of [
    'shared-filter-tag-btn',
    'shared-filter-color-btn',
    'shared-clear-filters',
    'shared-reset-routes'
  ]) {
    assert.ok(harness.elements[id], `#${id} remains in the controls`);
    assert.equal(harness.elements[id].hidden, false);
  }
  assert.equal(harness.elements['shared-search-expanded'].style.display, '');
  assert.equal(harness.elements['shared-filter-color-detail'].style.display, '');
  assert.equal(harness.elements['shared-filter-tag-detail'].style.display, 'none');
});

test('320px 390px and desktop search controls wrap before the tag button can clip horizontally', () => {
  assert.match(sharedCss, /#shared-map-search-bar\s*\{[^}]*overflow-x:\s*hidden[^}]*overflow-y:\s*auto/s);
  assert.match(sharedCss, /#shared-search-controls\s*\{[^}]*display:\s*flex[^}]*flex-wrap:\s*wrap/s);
  assert.match(sharedCss, /\.shared-text-btn\s*\{[^}]*padding:\s*8px 12px[^}]*white-space:\s*nowrap/s);
  const fixtures = [
    { viewport: 320, panel: 320 - 32, padding: 14 },
    { viewport: 390, panel: 390 - 32, padding: 14 },
    { viewport: 1440, panel: 500, padding: 19 }
  ];
  const itemWidths = [150, 98, 98, 98, 122];
  const tagIndex = 2;

  fixtures.forEach(({ viewport, panel, padding }) => {
    const contentWidth = panel - padding;
    let rowWidth = 0;
    let tagRight = 0;
    itemWidths.forEach((width, index) => {
      const proposed = rowWidth === 0 ? width : rowWidth + 6 + width;
      if (proposed > contentWidth) rowWidth = width;
      else rowWidth = proposed;
      if (index === tagIndex) tagRight = rowWidth;
    });
    assert.ok(tagRight <= contentWidth, `${viewport}px tag button must fit its wrapped row`);
    assert.ok(98 <= contentWidth, `${viewport}px individual filter controls must fit the panel`);
  });
});
