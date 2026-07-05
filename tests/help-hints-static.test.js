const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const sharedHtml = fs.readFileSync(path.join(root, 'shared.html'), 'utf8');

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

test('index help panel is available from the topbar and explains core concepts', () => {
  assertIncludes(indexHtml, '<button id="help-open-btn"');
  assertIncludes(indexHtml, '<div id="help-overlay" class="sheet-overlay"');
  assertIncludes(indexHtml, '<button id="help-close" class="ghost-btn" type="button">閉じる</button>');
  assertIncludes(indexHtml, 'dtp-help-seen-v112');

  ['基本', '編集モード', '未配置ピン', 'ルート', '共有リンク'].forEach((heading) => {
    assertIncludes(indexHtml, `<h3>${heading}</h3>`);
  });
  assertIncludes(indexHtml, '地図上のピンを押すと詳細を確認できます。');
  assertIncludes(indexHtml, '閲覧モードでは編集操作はできません。');
  assertIncludes(indexHtml, 'ピンを挿さずに説明用データとして残す用途もあります。');
  assertIncludes(indexHtml, '道路表示と直線表示を切り替えできます。');
  assertIncludes(indexHtml, 'ピンそのものをルート所属ピンだけに限定する設定ではありません。');

  const openBody = sourceFunctionBody(indexHtml, 'openHelpPanel');
  assertIncludes(openBody, "openOverlay('help-overlay');");
  assertIncludes(openBody, 'rememberHelpSeen();');
  assertIncludes(indexHtml, "document.getElementById('help-open-btn').addEventListener('click', openHelpPanel);");
  assertIncludes(indexHtml, "document.getElementById('help-close').addEventListener('click', closeHelpPanel);");
});

test('index share route wording describes visible routes without changing filtering behavior', () => {
  assertIncludes(indexHtml, '<div class="form-label">表示するルート</div>');
  assertIncludes(indexHtml, '選んだルート線だけを共有ビューに表示します。ピンはタグ・色の条件で絞り込まれます。');
  assertIncludes(indexHtml, '表示するルートはルート線の表示対象です。ピンの絞り込みはタグ・色で行います。');
  assertIncludes(indexHtml, "if (isShareRouteSelectionNone(routeIds)) return '表示するルート: なし';");
  assertIncludes(indexHtml, "if (!ids.length) return '表示するルート: すべて';");
  assertIncludes(indexHtml, "return '表示するルート: ' + names.join('、');");
  assert.equal(indexHtml.includes('<div class="form-label">対象ルート</div>'), false);
});

test('shared help panel is read-only focused and available from the topbar', () => {
  assertIncludes(sharedHtml, '<button id="shared-help-open-btn"');
  assertIncludes(sharedHtml, '<div id="shared-help-overlay" class="sheet-overlay"');
  assertIncludes(sharedHtml, '<button id="shared-help-close" class="shared-control-btn" type="button">閉じる</button>');
  assertIncludes(sharedHtml, 'dtp-shared-help-seen-v112');

  ['このページは閲覧専用', 'ピン', '検索・絞り込み', 'ルート', '注意'].forEach((heading) => {
    assertIncludes(sharedHtml, `<h3>${heading}</h3>`);
  });
  assertIncludes(sharedHtml, 'ピンの追加・編集・削除はできません。');
  assertIncludes(sharedHtml, '地図上のピンや一覧のピンを押すと詳細を表示できます。');
  assertIncludes(sharedHtml, 'ルートカードの表示ボタンで「道路 / 直線 / 非表示」を切り替えできます。');
  assertIncludes(sharedHtml, '説明用の未所属ピンが表示されることがあります。');

  const openBody = sourceFunctionBody(sharedHtml, 'openSharedHelpPanel');
  assertIncludes(openBody, "openSharedHelpOverlay();");
  assertIncludes(openBody, 'rememberSharedHelpSeen();');
  assertIncludes(sharedHtml, "document.getElementById('shared-help-open-btn').addEventListener('click', openSharedHelpPanel);");
  assertIncludes(sharedHtml, "document.getElementById('shared-help-close').addEventListener('click', closeSharedHelpPanel);");
});

test('help panels and lightweight hints support non-intrusive dismissal', () => {
  assertIncludes(indexHtml, "'help-overlay'");
  assertIncludes(indexHtml, "if (event.key === 'Escape' && document.getElementById('help-overlay').classList.contains('open'))");
  assertIncludes(indexHtml, 'ルートの表示/非表示は閲覧中でも切り替えできます。');

  assertIncludes(sharedHtml, "if (event.key === 'Escape' && document.getElementById('shared-help-overlay').classList.contains('open'))");
  assertIncludes(sharedHtml, 'ピンは検索・タグ・色で絞り込めます。');
  assertIncludes(sharedHtml, 'ルートのボタンで 道路 / 直線 / 非表示 を切り替えできます。');
});
