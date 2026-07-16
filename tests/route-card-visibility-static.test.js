const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');

function assertIncludes(source, needle, message) {
  assert.ok(source.includes(needle), message || `Expected source to include ${needle}`);
}

function sourceFunctionBody(source, name) {
  const index = source.indexOf(`function ${name}`);
  assert.notEqual(index, -1, `Expected function ${name} to exist`);
  const openIndex = source.indexOf('{', index);
  assert.notEqual(openIndex, -1, `Expected function ${name} to have a body`);
  let depth = 0;
  for (let i = openIndex; i < source.length; i += 1) {
    const char = source[i];
    if (char === '{') depth += 1;
    if (char === '}') depth -= 1;
    if (depth === 0) return source.slice(openIndex + 1, i);
  }
  assert.fail(`Could not parse function ${name}`);
}

test('route card visibility label is an accessible optimistic toggle', () => {
  const body = sourceFunctionBody(indexHtml, 'buildRouteItem');

  assertIncludes(indexHtml, 'routeVisibilityOverrides: {},');
  assertIncludes(indexHtml, 'function getRouteDisplayVisible(group)');
  assertIncludes(indexHtml, 'Object.prototype.hasOwnProperty.call(state.routeVisibilityOverrides, routeId)');
  assertIncludes(body, 'const isVisible = getRouteDisplayVisible(group);');
  assertIncludes(body, "visibility.type = 'button';");
  assertIncludes(body, "visibility.setAttribute('aria-pressed', isVisible ? 'true' : 'false');");
  assertIncludes(body, "visibility.setAttribute('aria-label',");
  assertIncludes(body, 'function toggleRouteVisibilityFromCard(event)');
  assertIncludes(body, 'event.stopPropagation();');
  assertIncludes(body, 'if (state.shareMode) return;');
  assertIncludes(body, 'if (canEditRouteControls()) {');
  assertIncludes(body, 'updateRouteGroupOptimistic(routeId, { visible: nextVisible });');
  assertIncludes(body, 'state.routeVisibilityOverrides[routeId] = nextVisible;');
  assertIncludes(body, 'renderPins();');
  assertIncludes(body, 'renderSidePanel();');
  assertIncludes(body, "visibility.addEventListener('click', toggleRouteVisibilityFromCard);");
  assert.equal(body.includes("visibility.addEventListener('keydown'"), false);
  assert.equal(body.includes('summary.appendChild(visibility)'), false);
});

test('route display visibility drives lines, numbers, and route card state without hiding pins', () => {
  const layerBody = sourceFunctionBody(indexHtml, 'renderRouteGroupLayers');
  const numberBody = sourceFunctionBody(indexHtml, 'getRouteNumberDisplayForPin');
  const pinRouteBody = sourceFunctionBody(indexHtml, 'isPinInAnyVisibleRoute');

  assertIncludes(layerBody, 'if (!getRouteDisplayVisible(group) || group.showLine === false)');
  assertIncludes(numberBody, 'if (!group || !getRouteDisplayVisible(group) || group.showNumbers === false) continue;');
  assert.equal(pinRouteBody.includes('getRouteDisplayVisible(group)'), false);
});

test('entering edit mode clears temporary route visibility overrides', () => {
  const body = sourceFunctionBody(indexHtml, 'setEditMode');

  assertIncludes(body, 'if (next) state.routeVisibilityOverrides = {};');
});

test('route detail edit controls no longer include the route visibility switch', () => {
  const body = sourceFunctionBody(indexHtml, 'renderRouteEditOverlay');

  assert.equal(body.includes("buildRouteEditToggleControl('表示/非表示'"), false);
  assertIncludes(body, "buildRouteEditToggleControl('番号を表示'");
  assertIncludes(body, "buildRouteEditToggleControl('線を表示'");
  assertIncludes(body, "buildRouteEditToggleControl('道路に沿わせる'");
  assertIncludes(body, "buildRouteEditToggleControl('循環ルート'");
});

test('share route clear action is labeled and behaves as explicit no-route selection', () => {
  assertIncludes(indexHtml, "const SHARE_ROUTE_NONE_SENTINEL = '__share_no_routes__';");
  assertIncludes(indexHtml, '<button id="share-route-clear-btn" class="ghost-btn" type="button">選択解除</button>');
  assert.equal(indexHtml.includes('<button id="share-route-clear-btn" class="ghost-btn" type="button">全ルートに戻す</button>'), false);
  assertIncludes(indexHtml, 'state.shareManager.routeTargets = [];');
  assertIncludes(indexHtml, "if (isShareRouteSelectionNone(routeIds)) return '表示するルート: なし';");

  const filterBody = sourceFunctionBody(indexHtml, 'getActivePinFilterState');
  assert.equal(filterBody.includes('routeIds: state.shareManager.routeIds'), false);
});
