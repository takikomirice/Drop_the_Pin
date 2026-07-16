const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const sharedHtml = fs.readFileSync(path.resolve(__dirname, '..', 'shared.html'), 'utf8');
const css = sharedHtml.match(/<style>([\s\S]*?)<\/style>/)[1];
const bodyMarkup = sharedHtml.slice(sharedHtml.indexOf('<body'), sharedHtml.indexOf('<script>', sharedHtml.indexOf('<body')));

function ruleBodies(selector) {
  const escaped = selector.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return Array.from(css.matchAll(new RegExp(`${escaped}\\s*\\{([^}]*)\\}`, 'g')), (match) => match[1]);
}

function ruleBody(selector, predicate = () => true) {
  const body = ruleBodies(selector).find(predicate);
  assert.ok(body, `Expected CSS rule ${selector}`);
  return body;
}

function declaration(body, property) {
  const escaped = property.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const match = body.match(new RegExp(`(?:^|;)\\s*${escaped}:\\s*([^;]+)`));
  assert.ok(match, `Expected ${property} in ${body}`);
  return match[1].trim();
}

function pixels(value) {
  const match = String(value).match(/^(-?\d+(?:\.\d+)?)px$/);
  assert.ok(match, `Expected pixel value, received ${value}`);
  return Number(match[1]);
}

function maxHeightFromContract(expression, viewportHeight, appShellHeight, headerHeight) {
  const subtraction = expression.match(/-\s*(\d+)px\s*\)$/);
  assert.ok(subtraction, `Expected bounded calc expression, received ${expression}`);
  const inset = Number(subtraction[1]);
  if (/100%/.test(expression)) return appShellHeight - inset;
  if (/100dvh/.test(expression)) return viewportHeight - headerHeight - inset;
  assert.fail(`Unknown max-height basis: ${expression}`);
}

const root = ruleBody(':root', (body) => /--app-header-height/.test(body));
const appShell = ruleBody('#shared-app-shell');
const searchPanel = ruleBody('#shared-map-search-bar', (body) => /max-height/.test(body));
const mobileSearchPanel = ruleBody('body.shared-narrow-view #shared-map-search-bar');
const searchInput = ruleBody('#shared-search-input', (body) => /min-width/.test(body));
const iconButton = ruleBody('.shared-icon-btn');
const searchPanelBase = ruleBody('#shared-map-search-bar', (body) => /z-index/.test(body));
const listPanelBase = ruleBody('#shared-side-panel', (body) => /z-index/.test(body));

const headerHeight = pixels(declaration(root, '--app-header-height'));
const searchTop = pixels(declaration(searchPanel, 'top'));
const maxHeightExpression = declaration(searchPanel, 'max-height');

test('search panel bottom stays inside the app shell for desktop mobile and safe-area fixtures', () => {
  const fixtures = [
    { viewport: '320x568', height: 568, safeTop: 0, expected: { appTop: 56, appHeight: 512, maxHeight: 472, bottom: 544 } },
    { viewport: '375x667', height: 667, safeTop: 20, expected: { appTop: 76, appHeight: 591, maxHeight: 551, bottom: 643 } },
    { viewport: '390x844', height: 844, safeTop: 47, expected: { appTop: 103, appHeight: 741, maxHeight: 701, bottom: 820 } },
    { viewport: '430x932', height: 932, safeTop: 59, expected: { appTop: 115, appHeight: 817, maxHeight: 777, bottom: 908 } },
    { viewport: '1440x900', height: 900, safeTop: 0, expected: { appTop: 56, appHeight: 844, maxHeight: 804, bottom: 876 } }
  ];

  for (const fixture of fixtures) {
    const appTop = headerHeight + fixture.safeTop;
    const appHeight = fixture.height - appTop;
    const maxHeight = maxHeightFromContract(maxHeightExpression, fixture.height, appHeight, headerHeight);
    const bottom = appTop + searchTop + maxHeight;
    assert.deepEqual(
      { appTop, appHeight, maxHeight, bottom },
      fixture.expected,
      `${fixture.viewport} safe-area-top ${fixture.safeTop}px geometry`
    );
    assert.ok(bottom <= fixture.height, `${fixture.viewport} search bottom ${bottom}px exceeds app shell bottom ${fixture.height}px`);
  }
});

test('search max-height is parent-relative and remains the primary contained scroller', () => {
  assert.equal(declaration(appShell, 'top'), 'calc(var(--app-header-height) + var(--safe-area-top))');
  assert.equal(declaration(appShell, 'bottom'), '0');
  assert.equal(maxHeightExpression, 'calc(100% - 40px)');
  assert.equal(declaration(searchPanel, 'overflow-y'), 'auto');
  assert.equal(declaration(searchPanel, 'overflow-x'), 'hidden');
  assert.equal(declaration(searchPanel, 'overscroll-behavior'), 'contain');
});

test('mobile search keeps bounded width flexible input and a 44px toggle above the sheet', () => {
  const toggle = bodyMarkup.match(/<button\b[^>]*\bid="shared-search-toggle"[^>]*>/);
  assert.ok(toggle, 'Expected #shared-search-toggle');
  assert.match(toggle[0], /class="[^"]*shared-icon-btn[^"]*"/);
  assert.equal(declaration(mobileSearchPanel, 'width'), 'calc(100% - 32px)');
  assert.equal(declaration(mobileSearchPanel, 'max-height'), 'calc(100% - 252px)');
  assert.equal(declaration(searchInput, 'min-width'), '0');
  assert.equal(declaration(iconButton, 'width'), '44px');
  assert.equal(declaration(iconButton, 'height'), '44px');
  assert.match(css, /\*, \*::before, \*::after\s*\{\s*box-sizing:\s*border-box/);
  assert.ok(
    Number(declaration(searchPanelBase, 'z-index')) > Number(declaration(listPanelBase, 'z-index')),
    'expanded search must remain operable when it overlaps the mobile sheet'
  );
});

test('topbar status note and side panel positioning contracts remain unchanged', () => {
  const topbar = ruleBody('#shared-topbar', (body) => /var\(--safe-area-top\)/.test(body));
  const status = ruleBody('#shared-status-note');
  const sidePanel = ruleBody('#shared-side-panel');
  const mobileSidePanel = ruleBody('#shared-side-panel', (body) => /height:\s*220px/.test(body));
  assert.equal(declaration(topbar, 'height'), 'calc(var(--app-header-height) + var(--safe-area-top))');
  assert.equal(declaration(status, 'top'), 'calc(var(--app-header-height) + var(--safe-area-top) + var(--sp-2))');
  assert.equal(declaration(sidePanel, 'position'), 'absolute');
  assert.equal(declaration(sidePanel, 'top'), '0');
  assert.equal(declaration(sidePanel, 'right'), '0');
  assert.equal(declaration(sidePanel, 'bottom'), '0');
  assert.equal(declaration(mobileSidePanel, 'top'), 'auto');
  assert.equal(declaration(mobileSidePanel, 'right'), '0');
  assert.equal(declaration(mobileSidePanel, 'bottom'), '0');
  assert.equal(declaration(mobileSidePanel, 'left'), '0');
  assert.equal(declaration(mobileSidePanel, 'height'), '220px');
});
