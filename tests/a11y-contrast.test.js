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

function constantArray(name) {
  const match = indexHtml.match(new RegExp(`const ${name} = \\[([\\s\\S]*?)\\n\\s*\\];`));
  assert.ok(match, `Expected ${name}`);
  return vm.runInNewContext(`[${match[1]}]`);
}

function optionalRuleBody(selector) {
  const escaped = selector.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const match = css.match(new RegExp(`${escaped}\\s*\\{([\\s\\S]*?)\\}`));
  return match ? match[1] : '';
}

function ruleBody(selector) {
  const body = optionalRuleBody(selector);
  const match = body !== '';
  assert.ok(match, `Expected CSS rule ${selector}`);
  return body;
}

function token(body, name) {
  const match = body.match(new RegExp(`${name}:\\s*(#[0-9a-f]{6})`, 'i'));
  assert.ok(match, `Expected ${name}`);
  return match[1].toUpperCase();
}

function rgb(hex) {
  const value = hex.replace('#', '');
  return [0, 2, 4].map((offset) => parseInt(value.slice(offset, offset + 2), 16) / 255);
}

function luminance(hex) {
  return rgb(hex)
    .map((channel) => channel <= 0.04045 ? channel / 12.92 : ((channel + 0.055) / 1.055) ** 2.4)
    .reduce((sum, channel, index) => sum + channel * [0.2126, 0.7152, 0.0722][index], 0);
}

function contrast(first, second) {
  const [lighter, darker] = [luminance(first), luminance(second)].sort((a, b) => b - a);
  return (lighter + 0.05) / (darker + 0.05);
}

function assertContrast(first, second, minimum, label) {
  const ratio = contrast(first, second);
  assert.ok(ratio >= minimum, `${label}: ${ratio.toFixed(2)}:1 is below ${minimum}:1 (${first} / ${second})`);
}

function declaredOpacity(selector) {
  const match = optionalRuleBody(selector).match(/(?:^|;)\s*opacity:\s*(0(?:\.\d+)?|1(?:\.0+)?)\s*;?/);
  return match ? Number(match[1]) : 1;
}

function composite(foreground, background, alpha) {
  const foregroundRgb = rgb(foreground);
  const backgroundRgb = rgb(background);
  const channels = foregroundRgb.map((channel, index) => Math.round((channel * alpha + backgroundRgb[index] * (1 - alpha)) * 255));
  return `#${channels.map((channel) => channel.toString(16).padStart(2, '0')).join('')}`.toUpperCase();
}

function assertEffectiveContrast(foreground, background, underlay, opacity, minimum, label) {
  assertContrast(
    composite(foreground, underlay, opacity),
    composite(background, underlay, opacity),
    minimum,
    label
  );
}

test('all route colors receive a computed badge foreground of at least 4.5:1', () => {
  const context = {
    escHtml: (value) => String(value),
    safeColor: (value) => value
  };
  vm.runInNewContext([
    functionSource('hexToRgb'),
    functionSource('relativeLuminance'),
    functionSource('contrastRatio'),
    functionSource('getReadableTextColor'),
    functionSource('createRouteNumberBadge')
  ].join('\n'), context);

  for (const color of constantArray('PIN_COLORS')) {
    const html = context.createRouteNumberBadge(1, 'list-route-number-badge', color.hex);
    const match = html.match(/--route-badge-foreground:\s*(#[0-9a-f]{6})/i);
    assert.ok(match, `Expected computed foreground for ${color.hex}`);
    assertContrast(match[1], color.hex, 4.5, `route badge ${color.label}`);
  }
  assert.match(ruleBody('.route-number-badge'), /color:\s*var\(--route-badge-foreground/);
});

test('route type badge theme tokens meet 4.5:1 in light and dark themes', () => {
  const themes = [ruleBody(':root, [data-theme="light"]'), ruleBody('[data-theme="dark"]')];
  for (const [themeIndex, body] of themes.entries()) {
    for (const kind of ['pin', 'gpx', 'geojson']) {
      const foreground = token(body, `--route-type-${kind}-foreground`);
      const background = token(body, `--route-type-${kind}-background`);
      assertContrast(foreground, background, 4.5, `${themeIndex ? 'dark' : 'light'} ${kind} badge`);
    }
  }
  for (const kind of ['pin', 'gpx', 'geojson']) {
    const body = ruleBody(`.route-type-badge.type-${kind}`);
    assert.match(body, new RegExp(`color:\\s*var\\(--route-type-${kind}-foreground\\)`));
    assert.match(body, new RegExp(`background:\\s*var\\(--route-type-${kind}-background\\)`));
  }
});

test('hidden route type badges retain 4.5:1 after ancestor opacity is composited', () => {
  const cardSelectors = ['.unified-route-card.is-hidden', '.route-item.is-hidden'];
  const themes = [
    ['light', ruleBody(':root, [data-theme="light"]')],
    ['dark', ruleBody('[data-theme="dark"]')]
  ];

  for (const [themeName, body] of themes) {
    const underlay = token(body, '--color-surface');
    for (const cardSelector of cardSelectors) {
      const opacity = declaredOpacity(cardSelector);
      for (const kind of ['pin', 'gpx', 'geojson']) {
        assertEffectiveContrast(
          token(body, `--route-type-${kind}-foreground`),
          token(body, `--route-type-${kind}-background`),
          underlay,
          opacity,
          4.5,
          `${themeName} hidden ${kind} badge in ${cardSelector}`
        );
      }
    }
  }
});

test('hidden route cards do not attenuate badges, focus rings, or controls through ancestor opacity', () => {
  for (const selector of ['.track-item.is-hidden', '.unified-route-card.is-hidden', '.route-item.is-hidden']) {
    assert.doesNotMatch(ruleBody(selector), /(?:^|;)\s*opacity\s*:/, selector);
  }
});

test('selected swatch uses light and dark indicators distinguishable from every route color', () => {
  const root = ruleBody(':root');
  const lightRing = token(root, '--selected-swatch-ring-light');
  const darkRing = token(root, '--selected-swatch-ring-dark');
  for (const color of constantArray('PIN_COLORS')) {
    const strongest = Math.max(contrast(lightRing, color.hex), contrast(darkRing, color.hex));
    assert.ok(strongest >= 3, `${color.label} swatch selection indicator is only ${strongest.toFixed(2)}:1`);
  }
  const selected = ruleBody('.color-swatch.selected::after');
  assert.match(selected, /content:\s*['"]✓['"]/);
  assert.match(selected, /--selected-swatch-ring-light/);
  assert.match(selected, /--selected-swatch-ring-dark/);
});

test('form borders meet 3:1 and state styles use dedicated tokens in both themes', () => {
  for (const [name, body] of [
    ['light', ruleBody(':root, [data-theme="light"]')],
    ['dark', ruleBody('[data-theme="dark"]')]
  ]) {
    const background = token(body, '--form-background');
    for (const borderToken of ['--form-border', '--form-border-hover']) {
      assertContrast(token(body, borderToken), background, 3, `${name} ${borderToken}`);
    }
    assertContrast(token(body, '--form-border-disabled'), token(body, '--form-background-disabled'), 3, `${name} disabled form border`);
  }
  const controls = ruleBody('.form-input, .form-textarea, .form-select');
  assert.match(controls, /border:\s*1px solid var\(--form-border\)/);
  assert.match(controls, /background:\s*var\(--form-background\)/);
  assert.match(css, /:where\(\.form-input, \.form-textarea, \.form-select\):hover:not\(:disabled\)[\s\S]*?var\(--form-border-hover\)/);
  assert.match(css, /:where\(\.form-input, \.form-textarea, \.form-select\):disabled[\s\S]*?var\(--form-border-disabled\)/);
  assert.match(css, /:where\(\.form-input, \.form-textarea, \.form-select\):focus-visible[\s\S]*?var\(--accent\)/);
  assert.match(css, /\[aria-invalid="true"\][\s\S]*?var\(--danger\)/);
});

test('disabled preset form borders retain 3:1 after ancestor opacity is composited', () => {
  const opacity = declaredOpacity('.input-preset-field-disabled');
  for (const [name, body] of [
    ['light', ruleBody(':root, [data-theme="light"]')],
    ['dark', ruleBody('[data-theme="dark"]')]
  ]) {
    const underlay = token(body, '--color-surface-raised');
    assertEffectiveContrast(
      token(body, '--form-border-disabled'),
      token(body, '--form-background-disabled'),
      underlay,
      opacity,
      3,
      `${name} disabled preset form border`
    );
  }
});

test('disabled preset fields use native disabled controls and control-level state tokens', () => {
  assert.doesNotMatch(optionalRuleBody('.input-preset-field-disabled'), /(?:^|;)\s*opacity\s*:/);
  const disabledControls = ruleBody('.input-preset-field-disabled :where(.form-input, .form-textarea, .form-select):disabled');
  assert.match(disabledControls, /background:\s*var\(--form-background-disabled\)/);
  assert.match(disabledControls, /border-color:\s*var\(--form-border-disabled\)/);
  assert.match(disabledControls, /color:\s*var\(--text-sub\)/);
  assert.match(disabledControls, /cursor:\s*not-allowed/);

  assert.match(indexHtml, /id="input-preset-tags" class="form-input"/);
  assert.match(indexHtml, /id="input-preset-status" class="form-select"/);
  const updateControls = functionSource('updateInputPresetEditorControls');
  assert.match(updateControls, /controlId: 'input-preset-tags'/);
  assert.match(updateControls, /controlId: 'input-preset-status'/);
  assert.match(updateControls, /getElementById\(dependency\.controlId\)\.disabled = saving \|\| !visible/);
});

test('primary, danger, and warning button theme tokens retain 4.5:1 text contrast', () => {
  for (const [name, body] of [
    ['light', ruleBody(':root, [data-theme="light"]')],
    ['dark', ruleBody('[data-theme="dark"]')]
  ]) {
    assertContrast(token(body, '--on-accent'), token(body, '--accent'), 4.5, `${name} primary button`);
    assertContrast(token(body, '--on-danger'), token(body, '--danger'), 4.5, `${name} danger button`);
    assertContrast(token(body, '--on-warning'), token(body, '--warning'), 4.5, `${name} warning button`);
  }
});
