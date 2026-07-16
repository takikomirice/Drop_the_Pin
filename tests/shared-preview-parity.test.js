const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const sharedHtml = fs.readFileSync(path.join(root, 'shared.html'), 'utf8');
const indexCss = indexHtml.match(/<style>([\s\S]*?)<\/style>/)[1];
const sharedCss = sharedHtml.match(/<style>([\s\S]*?)<\/style>/)[1];
const sharedBody = sharedHtml.slice(
  sharedHtml.indexOf('<body'),
  sharedHtml.indexOf('<script>', sharedHtml.indexOf('<body'))
);

function ruleBodies(source, selector) {
  const escaped = selector.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return Array.from(source.matchAll(new RegExp(`${escaped}\\s*\\{([^}]*)\\}`, 'g')), (match) => match[1]);
}

function ruleBody(source, selector, predicate = () => true) {
  const body = ruleBodies(source, selector).find(predicate);
  assert.ok(body, `Expected CSS rule ${selector}`);
  return body;
}

function declaration(body, property) {
  const escaped = property.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const match = body.match(new RegExp(`(?:^|;)\\s*${escaped}:\\s*([^;]+)`));
  assert.ok(match, `Expected ${property} in ${body}`);
  return match[1].trim();
}

function customProperty(source, name) {
  const match = source.match(new RegExp(`${name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}:\\s*([^;]+)`));
  assert.ok(match, `Expected ${name}`);
  return match[1].trim();
}

function markupBetween(source, startNeedle, endNeedle) {
  const start = source.indexOf(startNeedle);
  assert.notEqual(start, -1, `Expected ${startNeedle}`);
  const end = source.indexOf(endNeedle, start);
  assert.notEqual(end, -1, `Expected ${endNeedle}`);
  return source.slice(start, end);
}

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

test('shared preview shell uses the index header and dock geometry tokens', () => {
  assert.equal(customProperty(sharedCss, '--app-header-height'), customProperty(indexCss, '--app-header-height'));
  assert.equal(customProperty(sharedCss, '--app-dock-width'), customProperty(indexCss, '--app-dock-width'));

  const indexTopbar = ruleBody(indexCss, '#topbar', (body) => /position:\s*fixed/.test(body));
  const sharedTopbar = ruleBody(sharedCss, '#shared-topbar', (body) => /position:\s*fixed/.test(body));
  ['position', 'z-index', 'display', 'align-items', 'justify-content', 'background', 'backdrop-filter'].forEach((property) => {
    assert.equal(declaration(sharedTopbar, property), declaration(indexTopbar, property), property);
  });
  assert.equal(declaration(sharedTopbar, 'height'), 'calc(var(--app-header-height) + var(--safe-area-top))');

  const sharedShell = ruleBody(sharedCss, '#shared-app-shell');
  assert.equal(declaration(sharedShell, 'top'), 'calc(var(--app-header-height) + var(--safe-area-top))');
  assert.equal(declaration(sharedShell, 'bottom'), '0');
  assert.equal(declaration(sharedShell, 'left'), 'var(--safe-area-left)');
  assert.equal(declaration(sharedShell, 'right'), 'var(--safe-area-right)');

  const sharedDock = ruleBody(sharedCss, '#shared-side-panel', (body) => /grid-template-rows/.test(body));
  assert.equal(declaration(sharedDock, 'width'), 'min(var(--app-dock-width), 100%)');
  assert.equal(declaration(sharedDock, 'z-index'), '700');
  assert.equal(declaration(sharedDock, 'border-left'), 'var(--border-width) solid var(--border)');
  assert.equal(declaration(sharedDock, 'transform'), 'translateX(0)');
});

test('shared desktop dock follows the index pin then route region order', () => {
  const sharedDock = markupBetween(sharedBody, '<div id="shared-list-panel">', '</aside>');
  const indexDock = markupBetween(indexHtml, '<aside id="side-panel"', '</aside>');

  const sharedHeaderIndex = sharedDock.indexOf('id="shared-mobile-sheet-header"');
  const sharedPinIndex = sharedDock.indexOf('id="shared-pin-section"');
  const sharedRouteIndex = sharedDock.indexOf('id="shared-route-section"');
  assert.ok(sharedHeaderIndex < sharedPinIndex, 'mobile sheet header must remain first');
  assert.ok(sharedPinIndex < sharedRouteIndex, 'pin region must precede route region');

  const indexPinIndex = indexDock.indexOf('id="dock-pin-region"');
  const indexRouteIndex = indexDock.indexOf('id="dock-route-region"');
  assert.ok(indexPinIndex < indexRouteIndex, 'index pin region must precede route region');

  assert.equal(declaration(ruleBody(sharedCss, '#shared-list-panel'), 'display'), 'contents');
  const expanded = ruleBody(sharedCss, '#shared-side-panel.shared-route-dock-expanded');
  const tracks = declaration(expanded, 'grid-template-rows');
  assert.match(tracks, /--dock-pin-region-(?:min-height|preferred-height|max-height)/);
  assert.match(tracks, /--dock-route-region-min-height/);
  assert.ok(
    tracks.indexOf('--dock-pin-region') < tracks.indexOf('--dock-route-region'),
    'desktop grid row 1 must be pin and row 2 must be route'
  );
});

test('shared route dock expansion follows the index preview state contract', () => {
  assert.match(indexHtml, /routeDockExpanded:\s*false/);
  assert.doesNotMatch(sharedBody, /<aside id="shared-side-panel"[^>]*\bshared-route-dock-expanded\b/);
  assert.match(sharedHtml, /routeDockExpanded:\s*false/);

  const render = functionSource(sharedHtml, 'renderSharedRouteDockState');
  assert.match(render, /classList\.toggle\('shared-route-dock-expanded', expanded\)/);
  assert.match(render, /list\.hidden\s*=\s*!expanded/);
  assert.match(render, /aria-expanded/);
  assert.match(functionSource(sharedHtml, 'setSharedRouteDockExpanded'), /renderSharedRouteDockState\(\)/);
  assert.match(functionSource(sharedHtml, 'buildSharedPinRouteCard'), /setSharedRouteDockExpanded\(true\)/);
  assert.match(functionSource(sharedHtml, 'buildSharedTrackRouteCard'), /setSharedRouteDockExpanded\(true\)/);
  assert.match(functionSource(sharedHtml, 'setSharedMobileSheetTab'), /setSharedRouteDockExpanded\(true\)/);
});

test('shared bookmark is a topbar-external map and dock boundary control', () => {
  const topbarMarkup = markupBetween(sharedBody, '<div id="shared-topbar"', '<div id="shared-status-note"');
  const appShellMarkup = markupBetween(sharedBody, '<main id="shared-app-shell"', '</main>');
  assert.doesNotMatch(topbarMarkup, /id="shared-panel-toggle"/);
  assert.match(appShellMarkup, /id="shared-panel-toggle"/);

  const toggle = ruleBody(sharedCss, '#shared-panel-toggle', (body) => /position:\s*absolute/.test(body));
  assert.equal(declaration(toggle, 'top'), '18px');
  assert.equal(declaration(toggle, 'right'), 'var(--app-dock-width)');
  assert.equal(declaration(toggle, 'z-index'), '1120');
  assert.equal(declaration(toggle, 'width'), '44px');
  assert.equal(declaration(toggle, 'height'), '44px');
  assert.equal(declaration(toggle, 'border-radius'), '8px 0 0 8px');
  assert.equal(declaration(toggle, 'background'), 'var(--color-surface-raised)');
  assert.equal(declaration(toggle, 'box-shadow'), 'var(--shadow-sm)');
  assert.match(sharedCss, /body\.shared-panel-visible #shared-panel-toggle\s*\{\s*right:\s*var\(--app-dock-width\)/);
  assert.match(sharedCss, /body\.shared-panel-hidden #shared-panel-toggle\s*\{\s*right:\s*0/);
  assert.match(sharedBody, /class="shared-panel-toggle-icon shared-panel-toggle-icon-close"[\s\S]*class="shared-panel-toggle-icon shared-panel-toggle-icon-open"/);
});

test('shared topbar has one brand lockup and defers the shared map name', () => {
  const topbarMarkup = markupBetween(sharedBody, '<div id="shared-topbar"', '<div id="shared-status-note"');
  assert.equal((topbarMarkup.match(/>Drop the Pin!</g) || []).length, 1);
  assert.match(topbarMarkup, /class="shared-topbar-brand-icon"/);
  assert.match(topbarMarkup, /id="shared-display-name"[^>]*hidden/);
  assert.match(topbarMarkup, /id="shared-title"><\/span>/);
  assert.doesNotMatch(topbarMarkup, /shared-topbar-kicker/);
  assert.match(sharedHtml, /function syncSharedDisplayName\(/);
  assert.match(sharedHtml, /state\.status === 'ready'/);
  assert.match(sharedHtml, /label !== 'Drop the Pin!'/);
});

test('shared desktop map reclaims the dock and narrow view uses the index bottom sheet geometry', () => {
  assert.match(
    sharedCss,
    /@media \(min-width:\s*900px\)[\s\S]*body:not\(\.shared-panel-hidden\) #shared-map,[\s\S]*?\{\s*right:\s*var\(--app-dock-width\)/
  );
  assert.match(sharedCss, /body\.shared-panel-hidden #shared-map\s*\{\s*right:\s*0/);
  assert.match(sharedCss, /@media \(max-width:\s*640px\)[\s\S]*#shared-panel-toggle\s*\{\s*display:\s*none/);
  assert.match(
    sharedCss,
    /@media \(max-width:\s*640px\)[\s\S]*#shared-side-panel\s*\{[^}]*height:\s*220px[^}]*border-radius:\s*20px 20px 0 0/
  );
  assert.match(sharedCss, /#shared-mobile-sheet-tabs\s*\{[^}]*grid-template-columns:\s*repeat\(2, minmax\(0, 1fr\)\)/);
  assert.match(sharedCss, /#shared-mobile-sheet-handle\s*\{[^}]*width:\s*42px[^}]*height:\s*4px/);
  assert.match(sharedCss, /#shared-pin-list,[\s\S]*#shared-route-list\s*\{[^}]*overflow-y:\s*auto/);
  assert.doesNotMatch(sharedCss, /@media \(max-width:\s*899px\)/);
});

test('shared shell has no superseded floating-panel compatibility geometry', () => {
  assert.doesNotMatch(sharedCss, /--shared-header-height/);
  assert.doesNotMatch(sharedCss, /--shared-dock-width/);
  assert.doesNotMatch(sharedCss, /#shared-topbar\s*\{[^}]*height:\s*52px/);
  assert.doesNotMatch(sharedCss, /#shared-list-panel\s*\{[^}]*top:\s*64px/);
  assert.doesNotMatch(sharedCss, /#shared-list-panel\s*\{[^}]*width:\s*280px/);
  assert.doesNotMatch(sharedCss, /body\.shared-panel-hidden #shared-list-panel\s*\{\s*display:\s*none\s*!important/);
  const reducedMotion = sharedCss.slice(sharedCss.lastIndexOf('@media (prefers-reduced-motion: reduce)'));
  assert.match(reducedMotion, /#shared-map[^}]*transition:\s*none/);
  assert.match(reducedMotion, /#shared-side-panel[^}]*transition:\s*none/);
  assert.match(reducedMotion, /#shared-panel-toggle[^}]*transition:\s*none/);
});

test('shared preview search copies the index search hierarchy and accessible labels', () => {
  const searchMarkup = markupBetween(sharedBody, '<div id="shared-map-search-bar">', '<aside id="shared-side-panel"');
  assert.match(
    searchMarkup,
    /<span class="shared-map-search-leading-icon" aria-hidden="true">[\s\S]*?<div class="shared-map-search-field">[\s\S]*?<input id="shared-search-input"/
  );
  assert.match(searchMarkup, /placeholder="ピン・タグ・場所を検索" aria-label="ピン・タグ・場所を検索"/);
  assert.match(
    searchMarkup,
    /id="shared-search-toggle"[^>]*aria-label="フィルタ・ソート"[^>]*aria-controls="shared-search-expanded"[^>]*aria-expanded="false"[^>]*title="フィルタ・ソート"/
  );
  assert.match(searchMarkup, /id="shared-geocode-results"[\s\S]*id="shared-search-expanded"/);
  assert.doesNotMatch(searchMarkup, /shared-inline-help/);
});

test('shared preview search uses the translated index desktop geometry', () => {
  const search = ruleBody(sharedCss, '#shared-map-search-bar', (body) => /position:\s*absolute/.test(body));
  assert.equal(declaration(search, 'top'), '16px');
  assert.equal(declaration(search, 'left'), 'var(--sp-6)');
  assert.equal(declaration(search, 'width'), 'min(500px, calc(100% - 48px))');
  assert.equal(declaration(search, 'background'), 'var(--sheet-bg)');
  assert.equal(declaration(search, 'border-radius'), 'var(--radius-md)');
  assert.equal(declaration(search, 'box-shadow'), 'var(--shadow-md)');
  assert.equal(declaration(search, 'max-height'), 'calc(100% - 40px)');

  const desktopSearch = ruleBody(sharedCss, '#shared-map-search-bar', (body) => /padding:\s*4px 7px 4px 12px/.test(body));
  assert.equal(declaration(desktopSearch, 'padding'), '4px 7px 4px 12px');
  assert.equal(declaration(desktopSearch, 'border-radius'), '12px');
  const desktopRow = ruleBody(sharedCss, '#shared-map-search-bar-row1', (body) => /min-height:\s*34px/.test(body));
  assert.equal(declaration(desktopRow, 'min-height'), '34px');
  assert.equal(declaration(desktopRow, 'gap'), '8px');

  const field = ruleBody(sharedCss, '.shared-map-search-field');
  assert.equal(declaration(field, 'height'), '32px');
  assert.equal(declaration(field, 'padding'), '0 8px');
  assert.equal(declaration(field, 'border-radius'), '8px');
  assert.match(declaration(field, 'background'), /var\(--color-surface-raised\).*88%/);
  const input = ruleBody(sharedCss, '#shared-search-input', (body) => /height:\s*30px/.test(body));
  assert.equal(declaration(input, 'height'), '30px');
  assert.equal(declaration(input, 'min-height'), '30px');
  assert.equal(declaration(input, 'padding'), '0');
  assert.equal(declaration(input, 'border'), '0');
  assert.equal(declaration(input, 'line-height'), '30px');
});

test('shared preview search reserves dock width at compact desktop sizes', () => {
  assert.match(
    sharedCss,
    /@media \(min-width:\s*641px\) and \(max-width:\s*1279px\)[\s\S]*?--app-effective-dock-width:\s*clamp\(320px, 38vw, var\(--app-dock-width\)\)[\s\S]*?--app-compact-search-left:\s*var\(--sp-6\)[\s\S]*?--app-compact-control-gap:\s*var\(--sp-3\)/
  );
  assert.match(
    sharedCss,
    /body\.shared-panel-visible #shared-map-search-bar\s*\{[\s\S]*?width:\s*min\([\s\S]*?500px,[\s\S]*?100%[\s\S]*?- var\(--app-effective-dock-width\)[\s\S]*?- 36px[\s\S]*?- var\(--app-compact-search-left\)[\s\S]*?- var\(--app-compact-control-gap\)/
  );
  assert.match(
    sharedCss,
    /body\.shared-panel-hidden #shared-map-search-bar\s*\{[\s\S]*?width:\s*min\([\s\S]*?500px,[\s\S]*?100%[\s\S]*?- var\(--app-compact-search-left\)[\s\S]*?- var\(--app-compact-search-left\)/
  );
});

test('shared narrow search uses the index mobile dimensions above the bottom sheet', () => {
  const narrow = ruleBody(sharedCss, 'body.shared-narrow-view #shared-map-search-bar');
  assert.equal(declaration(narrow, 'top'), '12px');
  assert.equal(declaration(narrow, 'left'), '16px');
  assert.equal(declaration(narrow, 'width'), 'calc(100% - 32px)');
  assert.equal(declaration(narrow, 'max-height'), 'calc(100% - 252px)');
  assert.equal(declaration(narrow, 'padding'), '5px 7px');
  assert.equal(declaration(narrow, 'border-radius'), '12px');
  assert.equal(declaration(narrow, 'box-shadow'), 'var(--shadow-md)');
  const narrowInput = ruleBody(sharedCss, 'body.shared-narrow-view #shared-search-input');
  assert.equal(declaration(narrowInput, 'height'), '28px');
  assert.equal(declaration(narrowInput, 'min-height'), '28px');
  assert.equal(declaration(narrowInput, 'font-size'), '13px');
  assert.equal(declaration(narrowInput, 'line-height'), '28px');
});

test('shared readonly pin cards keep their hierarchy while index editable rows own separate chrome', () => {
  const renderer = functionSource(sharedHtml, 'renderSharedPins');
  const thumbRenderer = functionSource(sharedHtml, 'sharedPinListThumbMarkup');
  assert.match(renderer, /createRouteNumberBadge\(routeNumberDisplay\.number, 'list-route-number-badge', routeNumberDisplay\.color\)/);
  assert.match(renderer, /class="list-route-number-spacer"/);
  assert.match(thumbRenderer, /class="list-thumb"[\s\S]*'list-thumb-icon'/);
  assert.match(renderer, /class="list-meta"[\s\S]*class="list-title"[\s\S]*class="list-subtitle"/);
  assert.match(renderer, /class="view-card-chevron"[\s\S]*<svg/);
  assert.match(renderer, /aria-label="[^"]*の詳細を表示"/);
  assert.doesNotMatch(renderer, /shared-list-copy|shared-list-pin-icon/);
  assert.doesNotMatch(renderer, /<button[\s\S]*<button/);
  assert.doesNotMatch(renderer, /pin-drag-handle|pin-row-menu-btn/);

  const dockCard = ruleBody(sharedCss, '#shared-side-panel .list-item');
  assert.equal(declaration(dockCard, 'min-height'), '62px');
  assert.equal(declaration(dockCard, 'margin-bottom'), '5px');
  assert.equal(declaration(dockCard, 'padding'), '6px 9px');
  assert.equal(declaration(dockCard, 'border-radius'), '10px');
  const indexRow = ruleBody(indexCss, '#side-panel .pin-list-row');
  assert.equal(declaration(indexRow, 'min-height'), '62px');
  assert.equal(declaration(indexRow, 'margin-bottom'), '5px');
  assert.equal(declaration(indexRow, 'border-radius'), '10px');
  assert.equal(declaration(indexRow, 'background'), 'var(--color-surface-raised)');
  const indexMain = ruleBody(indexCss, '#side-panel .pin-list-row .list-item');
  assert.equal(declaration(indexMain, 'border'), '0');
  assert.equal(declaration(indexMain, 'background'), 'transparent');
  const mobileCard = ruleBody(sharedCss, 'body.shared-narrow-view #shared-side-panel .list-item');
  assert.equal(declaration(mobileCard, 'min-height'), '66px');
  assert.equal(declaration(mobileCard, 'margin-bottom'), '5px');
  assert.equal(declaration(mobileCard, 'padding'), '7px 10px');
});

test('shared PIN GPX and GeoJSON routes use one index unified card hierarchy', () => {
  const pinBuilder = functionSource(sharedHtml, 'buildSharedPinRouteCard');
  const trackBuilder = functionSource(sharedHtml, 'buildSharedTrackRouteCard');
  const renderer = functionSource(sharedHtml, 'renderSharedRouteList');
  assert.match(pinBuilder, /unified-route-card route-item/);
  assert.match(trackBuilder, /unified-route-card imported-route-item/);
  assert.doesNotMatch(trackBuilder, /unified-route-card imported-route-item track-item/);
  [pinBuilder, trackBuilder].forEach((builder) => {
    assert.match(builder, /className = 'route-card-header'/);
    assert.match(builder, /className = 'route-summary'/);
    assert.match(builder, /className = 'route-fit'/);
    assert.match(builder, /className = 'route-visibility'/);
    assert.match(builder, /setAttribute\('aria-pressed'/);
    assert.match(builder, /header\.appendChild\(summary\);[\s\S]*header\.appendChild\(fitButton\);[\s\S]*header\.appendChild\(visibility\);/);
    assert.doesNotMatch(builder, /summary\.appendChild\((?:fitButton|visibility)\)/);
  });
  assert.match(pinBuilder, /createSharedRouteTypeBadge\('pin'\)/);
  assert.match(trackBuilder, /createSharedRouteTypeBadge\(route\.type === 'gpx-route' \? 'gpx' : 'geojson'\)/);
  assert.match(renderer, /buildSharedPinRouteCard\(group\)/);
  assert.match(renderer, /buildSharedTrackRouteCard\(route\)/);
});

test('shared unified route cards copy index preview dimensions without parent opacity', () => {
  const card = ruleBody(sharedCss, '.unified-route-card');
  const indexCard = ruleBody(indexCss, '.unified-route-card');
  ['position', 'min-height', 'border', 'border-radius', 'background', 'color', 'overflow', 'box-shadow'].forEach((property) => {
    assert.equal(declaration(card, property), declaration(indexCard, property), property);
  });
  const summary = ruleBody(sharedCss, '.route-summary');
  const indexSummary = ruleBody(indexCss, '.route-summary', (body) => /width:\s*100%/.test(body));
  ['width', 'min-width', 'display', 'align-items', 'gap', 'padding', 'border', 'background', 'text-align', 'cursor'].forEach((property) => {
    assert.equal(declaration(summary, property), declaration(indexSummary, property), property);
  });
  const visibility = ruleBody(sharedCss, '.route-visibility');
  assert.equal(declaration(visibility, 'min-width'), '44px');
  assert.equal(declaration(visibility, 'min-height'), '44px');
  const fit = ruleBody(sharedCss, '.route-fit');
  assert.equal(declaration(fit, 'min-width'), '44px');
  assert.equal(declaration(fit, 'min-height'), '44px');
  assert.doesNotMatch(ruleBody(sharedCss, '.unified-route-card.is-hidden'), /(?:^|;)\s*opacity\s*:/);
  assert.doesNotMatch(sharedCss, /\.shared-route-card(?:\s|\{|,)/);
});

test('shared route pin rows copy index order typography and retain readonly detail access', () => {
  const builder = functionSource(sharedHtml, 'buildSharedRoutePinList');
  assert.match(builder, /className = 'route-pin-row'/);
  assert.match(builder, /className = 'route-pin-order'/);
  assert.doesNotMatch(builder, /order\.style\.(?:backgroundColor|color)/);
  const sharedOrder = ruleBody(sharedCss, '.route-pin-order');
  const indexOrder = ruleBody(indexCss, '.route-pin-order');
  ['color', 'font-size', 'font-weight', 'text-align', 'user-select'].forEach((property) => {
    assert.equal(declaration(sharedOrder, property), declaration(indexOrder, property), property);
  });
  assert.match(builder, /className = 'shared-route-pin-action'/);
  assert.match(builder, /openSharedDetail\(pin\)/);
  assert.match(builder, /highlightLinkedMarker\(pin\.id\)/);
  assert.doesNotMatch(builder, /(?:edit|delete|remove|drag|reorder)/i);
});

test('shared mobile tabs copy index icon label structure and support keyboard navigation', () => {
  const tabs = markupBetween(sharedBody, '<div id="shared-mobile-sheet-tabs"', '</div>');
  assert.match(tabs, /id="shared-mobile-pins-tab"[^>]*role="tab"[^>]*aria-selected="true"[^>]*aria-controls="shared-pin-section"[^>]*tabindex="0"[\s\S]*<svg[^>]*aria-hidden="true"[\s\S]*<span>ピン<\/span>/);
  assert.match(tabs, /id="shared-mobile-routes-tab"[^>]*role="tab"[^>]*aria-selected="false"[^>]*aria-controls="shared-route-section"[^>]*tabindex="-1"[\s\S]*<svg[^>]*aria-hidden="true"[\s\S]*<span>ルート<\/span>/);
  const tabRule = ruleBody(sharedCss, '#shared-mobile-sheet-tabs button');
  const indexTabRule = ruleBody(indexCss, '#mobile-sheet-tabs button');
  ['min-width', 'min-height', 'padding', 'border', 'border-radius', 'background', 'color', 'font-size', 'font-weight'].forEach((property) => {
    assert.equal(declaration(tabRule, property), declaration(indexTabRule, property), property);
  });
  const keyboard = functionSource(sharedHtml, 'handleSharedMobileSheetTabKeydown');
  ['ArrowLeft', 'ArrowRight', 'Home', 'End'].forEach((key) => assert.match(keyboard, new RegExp(key)));
  assert.match(sharedHtml, /getElementById\('shared-mobile-sheet-tabs'\)\.addEventListener\('keydown', handleSharedMobileSheetTabKeydown\)/);
  assert.doesNotMatch(sharedCss, /\.shared-mobile-sheet-tab(?:\s|\{|,|\[)/);
});

test('shared readonly detail uses the index sheet hierarchy and desktop dock geometry', () => {
  const detailMarkup = sharedBody.slice(sharedBody.indexOf('<div id="shared-detail-overlay"'));
  assert.match(detailMarkup, /<div id="shared-detail-overlay" class="sheet-overlay"[^>]*>[\s\S]*?<div class="sheet-body">/);
  assert.match(detailMarkup, /class="sheet-handle" aria-hidden="true"/);
  assert.match(detailMarkup, /id="shared-detail-image" class="photo-fit-cover protected-photo" draggable="false"/);
  assert.match(detailMarkup, /id="shared-detail-title"/);
  assert.match(detailMarkup, /id="shared-detail-status"/);
  assert.match(detailMarkup, /id="shared-detail-tags"/);
  assert.match(detailMarkup, /id="shared-detail-description" class="sheet-subtitle"/);
  assert.match(detailMarkup, /id="shared-detail-routes"/);
  assert.match(detailMarkup, /id="shared-detail-links"/);
  assert.match(detailMarkup, /id="shared-detail-time" class="muted"/);
  assert.match(detailMarkup, /id="shared-detail-close" class="ghost-btn"[^>]*data-shared-initial-focus/);

  const indexOverlay = ruleBody(indexCss, '.sheet-overlay', (body) => /position:\s*fixed/.test(body));
  const sharedOverlay = ruleBody(sharedCss, '.sheet-overlay', (body) => /position:\s*fixed/.test(body));
  ['position', 'z-index', 'width', 'height', 'display', 'align-items', 'justify-content', 'overflow', 'overscroll-behavior', 'background'].forEach((property) => {
    assert.equal(declaration(sharedOverlay, property), declaration(indexOverlay, property), property);
  });
  const indexSheet = ruleBody(indexCss, '.sheet-body', (body) => /width:\s*min\(100%, 520px\)/.test(body));
  const sharedSheet = ruleBody(sharedCss, '.sheet-body', (body) => /width:\s*min\(100%, 520px\)/.test(body));
  ['width', 'max-height', 'min-height', 'overflow-y', 'overscroll-behavior', 'scroll-padding-block', 'background', 'border', 'border-radius', 'box-shadow', 'padding'].forEach((property) => {
    assert.equal(declaration(sharedSheet, property), declaration(indexSheet, property), property);
  });

  const docked = ruleBody(sharedCss, 'body:not(.shared-panel-hidden) #shared-detail-overlay.open');
  assert.equal(declaration(docked, 'width'), 'var(--app-dock-width)');
  assert.equal(declaration(docked, 'height'), 'calc((100dvh - var(--app-header-height) - var(--safe-area-top)) * 0.68)');
  assert.equal(declaration(docked, 'padding'), '0');
  assert.equal(declaration(docked, 'background'), 'var(--color-surface-raised)');
  const dockedSheet = ruleBody(sharedCss, 'body:not(.shared-panel-hidden) #shared-detail-overlay .sheet-body');
  assert.equal(declaration(dockedSheet, 'width'), '100%');
  assert.equal(declaration(dockedSheet, 'max-width'), 'none');
  assert.equal(declaration(dockedSheet, 'max-height'), 'none');
  assert.equal(declaration(dockedSheet, 'border-radius'), '0');
  assert.equal(declaration(dockedSheet, 'box-shadow'), 'none');
});

test('shared readonly detail uses the index narrow bottom sheet and populates every viewing field', () => {
  const narrowOverlay = ruleBody(sharedCss, 'body.shared-narrow-view #shared-detail-overlay.open');
  assert.equal(declaration(narrowOverlay, 'display'), 'flex');
  assert.equal(declaration(narrowOverlay, 'align-items'), 'flex-end');
  assert.equal(declaration(narrowOverlay, 'justify-content'), 'stretch');
  assert.equal(declaration(narrowOverlay, 'padding'), '0');
  assert.equal(declaration(narrowOverlay, 'background'), 'transparent');
  const narrowSheet = ruleBody(sharedCss, 'body.shared-narrow-view #shared-detail-overlay .sheet-body');
  assert.equal(declaration(narrowSheet, 'width'), '100%');
  assert.equal(declaration(narrowSheet, 'height'), 'min(75dvh, calc(var(--dialog-viewport-height, 100dvh) - env(safe-area-inset-top)), 634px)');
  assert.equal(declaration(narrowSheet, 'max-height'), 'min(75dvh, calc(var(--dialog-viewport-height, 100dvh) - env(safe-area-inset-top)), 634px)');
  assert.equal(declaration(narrowSheet, 'border-radius'), '20px 20px 0 0');
  assert.equal(declaration(narrowSheet, 'border-bottom'), '0');

  const detail = functionSource(sharedHtml, 'openSharedDetail');
  [
    'shared-detail-title', 'shared-detail-image', 'shared-detail-status',
    'shared-detail-tags', 'shared-detail-description', 'shared-detail-routes',
    'shared-detail-links', 'shared-detail-time'
  ].forEach((id) => assert.match(detail, new RegExp(`getElementById\\('${id}'\\)`)));
  assert.doesNotMatch(detail, /sharedPinListIconMarkup|shared-detail-icon/);
  assert.match(detail, /getSharedPinRouteLabels\(pin\.id\)/);
  assert.match(detail, /renderTagChips\(pin\.tags\)/);
  assert.match(detail, /openSharedSurface\('shared-detail-overlay'\)/);
  assert.doesNotMatch(detail, /(?:編集|削除|保存|共有作成)/);
});

test('shared detail switches between docked non-modal and modal accessibility states', () => {
  const docked = functionSource(sharedHtml, 'isSharedDockedPinDetailOverlay');
  assert.match(docked, /overlay\.id !== 'shared-detail-overlay'/);
  assert.match(docked, /matchMedia\('\(min-width: 900px\)'\)/);
  assert.match(docked, /classList\.contains\('shared-panel-visible'\)/);

  const sync = functionSource(sharedHtml, 'syncSharedSurfaceInteractionState');
  assert.match(sync, /!isSharedDockedPinDetailOverlay/);
  assert.match(sync, /active \? 'true' : 'false'/);
  const trap = functionSource(sharedHtml, 'trapSharedSurfaceFocus');
  assert.match(trap, /getTopSharedModalSurfaceRecord\(\)/);
  const backdrop = functionSource(sharedHtml, 'setupSharedOverlayBackdropDismissal');
  assert.match(backdrop, /isSharedDockedPinDetailOverlay\(overlay\)/);
});

test('shared help uses the index dialog shell and superseded shared dialog CSS is absent', () => {
  const helpMarkup = markupBetween(sharedBody, '<div id="shared-help-overlay"', '<div id="shared-detail-overlay"');
  assert.match(helpMarkup, /<div class="sheet-body">/);
  assert.match(helpMarkup, /class="sheet-handle" aria-hidden="true"/);
  assert.match(helpMarkup, /id="shared-help-title" class="sheet-title"/);
  assert.match(helpMarkup, /class="help-content"/);
  assert.match(helpMarkup, /id="shared-help-close" class="ghost-btn"[^>]*data-shared-initial-focus/);
  assert.match(sharedCss, /\.help-content\s*\{[^}]*display:\s*flex[^}]*gap:\s*12px/);

  [
    'detail-card', 'shared-sheet-handle', 'shared-dialog-header', 'shared-dialog-heading',
    'shared-dialog-body', 'shared-dialog-close', 'shared-help-content', 'detail-image',
    'detail-body', 'detail-section-title', 'detail-links', 'detail-link'
  ].forEach((legacyClass) => {
    assert.doesNotMatch(sharedHtml, new RegExp(`\\.${legacyClass}(?:\\s|\\{|,|\\[)`), legacyClass);
    assert.doesNotMatch(sharedBody, new RegExp(`class="[^"]*\\b${legacyClass}\\b`), legacyClass);
  });
});

test('shared final preview shell contains no superseded UI selectors tokens or duplicate breakpoint blocks', () => {
  [
    'shared-topbar-title-wrap', 'shared-topbar-kicker', 'shared-brand',
    'shared-panel-icon-show', 'shared-panel-icon-hide', 'shared-list-copy', 'shared-list-pin-icon',
    'shared-route-card', 'shared-route-card-header', 'shared-route-fit', 'shared-route-color-dot',
    'shared-route-name', 'shared-route-display-mode', 'shared-route-meta', 'shared-route-pin-item',
    'shared-route-pin-num', 'shared-route-pin-title', 'shared-route-type-badge',
    'detail-card', 'shared-dialog-sheet', 'shared-dialog-header', 'shared-dialog-body'
  ].forEach((legacyName) => assert.doesNotMatch(sharedHtml, new RegExp(`\\b${legacyName}\\b`), legacyName));

  [
    '--shared-header-height', '--shared-dock-width', '--route-type-badge-bg',
    '--route-type-badge-text', '--surface-strong', '--surface-muted', '--note-bg'
  ].forEach((legacyToken) => assert.doesNotMatch(sharedCss, new RegExp(legacyToken), legacyToken));
  assert.doesNotMatch(sharedCss, /\.shared-control-btn\.is-off/);
  assert.doesNotMatch(sharedCss, /@media \(max-width:\s*899px\)/);
  assert.equal((sharedCss.match(/@media \(min-width: 900px\)\s*\{/g) || []).length, 1);
  assert.equal((sharedCss.match(/@media \(max-width: 640px\)\s*\{/g) || []).length, 1);
  assert.equal((sharedCss.match(/@media \(prefers-reduced-motion: reduce\)\s*\{/g) || []).length, 1);

  const tokenBlocks = [
    sharedCss.match(/:root\s*\{([^}]*)\}/)[1],
    sharedCss.match(/:root, \[data-theme="light"\]\s*\{([^}]*)\}/)[1],
    sharedCss.match(/\[data-theme="dark"\]\s*\{([^}]*)\}/)[1]
  ];
  tokenBlocks.forEach((block) => {
    const names = Array.from(block.matchAll(/(--[a-z0-9-]+)\s*:/gi), (match) => match[1]);
    assert.equal(new Set(names).size, names.length, `duplicate token in ${block.slice(0, 30)}`);
  });
});
