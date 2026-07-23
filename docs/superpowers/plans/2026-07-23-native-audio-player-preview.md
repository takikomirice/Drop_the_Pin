# Native Audio Player Preview Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 編集モード、編集URL内プレビュー、共有URLで保存済み音声をブラウザ標準プレーヤーにより再生でき、対応ブラウザのダウンロード項目だけを非表示にする。

**Architecture:** `index.html`と`shared.html`の共通`AudioPinPlayer`ブロックをネイティブ`audio[controls]`向けに置き換え、既存コントローラーは取得、検証、LRU、停止、Object URL解放だけを担当する。編集ページの表示条件は変更権限`canEdit()`から編集トークン所持へ分離し、共有ページは既存の共有投影認可を維持する。

**Tech Stack:** Google Apps Script HTML、vanilla JavaScript、Node.js `node:test`、Playwright

## Global Constraints

- `index.html`と`shared.html`の`data-dtp-audio-player-boundary`間はバイト単位で同一にする。
- `getPinAudioData`、`getSharedPinAudioData`、レスポンス形状、Drive配置、Spreadsheet列を変更しない。
- 最大3件、合計12MBのLRUと、1KiB以上4MiB以下のMP3検証を維持する。
- 通常URLから編集用音声APIを呼ばず、共有ページへ編集トークンを持ち込まない。
- 音声編集結果、保存済み音声の両方へ`controlsList="nodownload"`だけを指定する。
- ネイティブプレーヤーは幅100%以内に収め、375px、768px、1280pxで横方向に溢れさせない。
- ブラウザ標準UIの外観、追加メニュー、`nodownload`対応差を許容する。

---

### Task 1: 共通プレーヤーをネイティブコントロールへ置き換える

**Files:**
- Modify: `tests/audio-player-core.test.js`
- Modify: `tests/audio-player-integration.test.js`
- Modify: `tests/audio-editor-integration.test.js`
- Modify: `index.html`
- Modify: `shared.html`

**Interfaces:**
- Consumes: `AudioPinPlayer.create({ audioElement, fetchAudio, renderState })`
- Produces: `AudioPinPlayer.createRenderer({ container, statusElement, audioElement, retryButton })`
- Produces: controller methods `open(pinId)`, `retry()`, `close()`, `invalidate(pinId)`, `destroy()`

- [ ] **Step 1: ネイティブUIの失敗テストを書く**

`tests/audio-player-core.test.js`の断片テストを、次の契約へ変更する。

```js
test('player fragment exposes the native accessible playback UI without a download action', () => {
  const html = playerHtml();
  assert.equal((html.match(/<audio\s+id="pin-audio-runtime"/g) || []).length, 1);
  assert.match(html, /<audio\s+id="pin-audio-runtime"[^>]*\bcontrols\b/);
  assert.match(html, /<audio\s+id="pin-audio-runtime"[^>]*\bcontrolsList="nodownload"/);
  assert.match(html, /<audio\s+id="pin-audio-runtime"[^>]*\bpreload="metadata"/);
  assert.doesNotMatch(html, /id="pin-audio-player-toggle"/);
  assert.match(html, /aria-live="polite"/);
  assert.match(html, /音声を準備中…/);
  assert.match(html, /音声を再生できませんでした/);
  assert.match(html, /再試行/);
});
```

rendererテストの要素を`audioElement`へ変更し、次の状態を固定する。

```js
const elements = {
  container: createElement(),
  statusElement: createElement(),
  audioElement: createElement(),
  retryButton: createElement()
};
const render = api.createRenderer(elements);

render({ status: 'loading' });
assert.equal(elements.container.hidden, false);
assert.equal(elements.statusElement.textContent, '音声を準備中…');
assert.equal(elements.audioElement.hidden, true);

render({ status: 'ready' });
assert.equal(elements.statusElement.textContent, '');
assert.equal(elements.audioElement.hidden, false);
assert.equal(elements.retryButton.hidden, true);

render({ status: 'error' });
assert.equal(elements.statusElement.textContent, '音声を再生できませんでした');
assert.equal(elements.audioElement.hidden, true);
assert.equal(elements.retryButton.hidden, false);

render({ status: 'hidden' });
assert.equal(elements.container.hidden, true);
```

`tests/audio-editor-integration.test.js`の結果プレーヤーテストへ次を追加する。

```js
assert.match(html, /data-hae-result-audio[^>]*controlsList="nodownload"/);
assert.doesNotMatch(html, /controlsList="(?:[^"]*\s)?(?:nofullscreen|noplaybackrate|noremoteplayback)/);
```

- [ ] **Step 2: 集中テストを実行してREDを確認する**

Run:

```powershell
node --test tests/audio-player-core.test.js tests/audio-player-integration.test.js tests/audio-editor-integration.test.js
```

Expected: `controls`、`controlsList="nodownload"`、rendererの`audioElement`契約が未実装でFAIL。

- [ ] **Step 3: 共通マークアップとrendererを最小実装する**

`index.html`と`shared.html`の共通ブロックを同じ内容にし、独自トグルを削除する。

```html
<section id="pin-audio-player" class="pin-audio-player" aria-label="音声" aria-busy="false" hidden>
  <span id="pin-audio-player-status" class="pin-audio-player-status" role="status" aria-live="polite"></span>
  <audio id="pin-audio-runtime" controls controlsList="nodownload" preload="metadata"
    aria-label="音声を再生" hidden></audio>
  <button id="pin-audio-player-retry" class="ghost-btn" type="button" hidden>再試行</button>
</section>
```

共通CSSへ次を設定する。

```css
.pin-audio-player audio {
  display: block;
  width: 100%;
  min-width: 0;
  max-width: 100%;
}
.pin-audio-player-status:empty {
  display: none;
}
```

`createRenderer()`を次の状態投影へ変更する。

```js
function createRenderer(elements) {
  const source = elements || {};
  const container = source.container;
  const statusElement = source.statusElement;
  const audioElement = source.audioElement;
  const retryButton = source.retryButton;
  if (!container || !statusElement || !audioElement || !retryButton) {
    throw new Error('Audio player elements are unavailable.');
  }

  return function render(view) {
    const status = view && typeof view.status === 'string' ? view.status : 'hidden';
    const hidden = status === 'hidden';
    const loading = status === 'loading';
    const ready = status === 'ready';
    const failed = status === 'error';

    container.hidden = hidden;
    container.setAttribute('aria-busy', loading ? 'true' : 'false');
    statusElement.textContent = loading
      ? '音声を準備中…'
      : (failed ? '音声を再生できませんでした' : '');
    audioElement.hidden = !ready;
    retryButton.hidden = !failed;
    retryButton.disabled = !failed;
    retryButton.textContent = '再試行';
  };
}
```

- [ ] **Step 4: コントローラーから独自再生状態を除く**

`AudioPinPlayer.create()`から`playbackGeneration`、`toggle()`、`stop()`、`handleEnded()`、`ended` listenerを削除する。`pauseAndReset()`はクローズ、ピン切替、失効、破棄時の停止だけを担当する。

```js
function pauseAndReset() {
  try { audioElement.pause(); } catch (_error) {}
  try { audioElement.currentTime = 0; } catch (_error) {}
}

return Object.freeze({
  open: open,
  retry: retry,
  close: close,
  invalidate: invalidate,
  destroy: destroy
});
```

編集・共有factoryは`audioElement`をrendererへ渡し、retryだけを配線する。

```js
const audioElement = document.getElementById('pin-audio-runtime');
const retryButton = document.getElementById('pin-audio-player-retry');
const controller = AudioPinPlayer.create({
  audioElement: audioElement,
  fetchAudio: fetchAudio,
  renderState: AudioPinPlayer.createRenderer({
    container: document.getElementById('pin-audio-player'),
    statusElement: document.getElementById('pin-audio-player-status'),
    audioElement: audioElement,
    retryButton: retryButton
  })
});
retryButton.addEventListener('click', function() { controller.retry(); });
```

音声編集結果の要素も同じダウンロード抑止にする。

```html
<audio class="hae-result-audio" data-hae-result-audio controls
  controlsList="nodownload" preload="metadata" aria-label="生成したMP3を確認再生"></audio>
```

- [ ] **Step 5: コアテストをGREENへ更新する**

独自`toggle()`を検証していたテストを削除し、ネイティブ再生中でもcontrollerのクローズ処理が`pause()`、`currentTime = 0`、`src`解除を行うことを検証する。

```js
test('player stops native playback on pin switch close and pagehide', async () => {
  const loaded = loadPlayer();
  const audio = createAudioElement();
  const controller = loaded.api.create({
    audioElement: audio,
    fetchAudio() { return Promise.resolve(response()); },
    renderState() {}
  });

  await controller.open('a');
  await audio.play();
  audio.currentTime = 8;
  await controller.open('b');
  assert.equal(audio.currentTime, 0);

  audio.currentTime = 4;
  controller.close();
  assert.equal(audio.currentTime, 0);

  await controller.open('c');
  loaded.emitPagehide();
  assert.equal(audio.src, '');
  assert.equal(await controller.open('after-destroy'), false);
});
```

- [ ] **Step 6: 集中テストを実行してGREENを確認する**

Run:

```powershell
node --test tests/audio-player-core.test.js tests/audio-player-integration.test.js tests/audio-editor-integration.test.js
```

Expected: 全件PASS、FAIL 0。

- [ ] **Step 7: 実装単位をコミットする**

```powershell
git add index.html shared.html tests/audio-player-core.test.js tests/audio-player-integration.test.js tests/audio-editor-integration.test.js
git commit -m "feat: use native controls for pin audio"
```

---

### Task 2: 編集URL内プレビューで音声閲覧を許可する

**Files:**
- Modify: `tests/audio-player-integration.test.js`
- Modify: `tests/browser/production-template-integration.spec.js`
- Modify: `index.html`

**Interfaces:**
- Consumes: template定数`hasEditToken`
- Consumes: `openPinDetail(pin)`
- Produces: `hasEditToken && pin.hasAudio`による編集用音声閲覧判定

- [ ] **Step 1: プレビュー回帰の失敗テストを書く**

`tests/audio-player-integration.test.js`で`openPinDetail()`の条件を固定する。

```js
assert.match(
  openSource,
  /hasEditToken\s*&&\s*pin\.hasAudio[\s\S]*?pinAudioPlayer\.open\s*\(\s*pin\.id\s*\)/
);
assert.doesNotMatch(openSource, /canEdit\(\)\s*&&\s*pin\.hasAudio/);
```

`tests/browser/production-template-integration.spec.js`へプレビュー動作を追加する。

```js
test('edit preview keeps native audio playback while mutation actions stay hidden', async ({ page }) => {
  await page.goto('/production-edit');
  await expect.poll(() => page.evaluate(() => typeof window.__productionEdit)).toBe('object');
  const pin = productionPin({ id: 'preview-audio-pin' });
  await enqueue(page, 'getPinAudioData', { audioSeed: 13 });

  await page.evaluate((pin) => {
    window.__productionEdit.state.previewMode = true;
    window.__productionEdit.state.pins = [pin];
    window.__productionEdit.openPinDetail(pin);
  }, pin);

  await expect(page.locator('#pin-audio-runtime')).toBeVisible();
  await expect(page.locator('#pin-audio-runtime')).toHaveAttribute('controlsList', 'nodownload');
  await expect(page.locator('#pin-detail-audio-actions')).toBeHidden();
  const calls = await page.evaluate(() =>
    window.__gasMock.calls.filter((call) => call.method === 'getPinAudioData'));
  expect(calls).toEqual([{
    method: 'getPinAudioData',
    payload: { __editToken: 'edit-token-browser-test', pinId: pin.id }
  }]);
});
```

- [ ] **Step 2: REDを確認する**

Run:

```powershell
node --test tests/audio-player-integration.test.js
npx playwright test tests/browser/production-template-integration.spec.js --grep "edit preview"
```

Expected: `openPinDetail()`が`canEdit()`を使っているためFAIL。

- [ ] **Step 3: 表示条件を編集権限から閲覧認証へ分離する**

`index.html`の`openPinDetail()`を次へ変更する。

```js
openOverlay('pin-detail-overlay');
if (hasEditToken && pin.hasAudio) {
  pinAudioPlayer.open(pin.id);
} else {
  pinAudioPlayer.close();
}
```

`renderPinDetailAudioActions()`は従来どおり`canEdit()`を使い、プレビューで追加、差し替え、削除を表示しない。

- [ ] **Step 4: NodeとPlaywrightのGREENを確認する**

Run:

```powershell
node --test tests/audio-player-integration.test.js tests/preview-mobile-ui.test.js
npx playwright test tests/browser/production-template-integration.spec.js --grep "edit preview"
```

Expected: 全件PASS、プレビューの取得は1回、変更操作は非表示。

- [ ] **Step 5: 実装単位をコミットする**

```powershell
git add index.html tests/audio-player-integration.test.js tests/browser/production-template-integration.spec.js
git commit -m "fix: show pin audio in edit preview"
```

---

### Task 3: 編集・共有ブラウザ回帰とREADMEを更新する

**Files:**
- Modify: `tests/browser/audio-player.spec.js`
- Modify: `tests/browser/shared-audio.spec.js`
- Modify: `tests/browser/audio-editor.spec.js`
- Modify: `tests/browser/production-template-integration.spec.js`
- Modify: `tests/shared-audio-ui.test.js`
- Modify: `README.md`

**Interfaces:**
- Consumes: `#pin-audio-runtime`
- Consumes: `controlsList="nodownload"`
- Produces: 編集、プレビュー、共有、音声編集結果のエンドツーエンド回帰

- [ ] **Step 1: ブラウザ契約をネイティブUIへ更新する**

`tests/browser/audio-player.spec.js`では独自ボタン操作をネイティブ要素へ置き換える。

```js
await expect(page.locator('#pin-audio-player-status')).toHaveText('音声を準備中…');
await expect(page.locator('#pin-audio-runtime')).toBeHidden();
await page.evaluate(() => window.__gasMock.resolve('pin-a-ready'));
await expect(page.locator('#pin-audio-runtime')).toBeVisible();
await expect(page.locator('#pin-audio-runtime')).toHaveAttribute('controlsList', 'nodownload');
expect(await page.locator('#pin-audio-runtime').evaluate((audio) => audio.controls)).toBe(true);
await page.locator('#pin-audio-runtime').evaluate((audio) => audio.play());
expect(await page.evaluate(() => window.__media.plays)).toBe(1);
```

失敗時はretryだけを検証する。

```js
await expect(page.locator('#pin-audio-player-status')).toHaveText('音声を再生できませんでした');
await expect(page.locator('#pin-audio-runtime')).toBeHidden();
await expect(page.locator('#pin-audio-player-retry')).toBeVisible();
```

幅テストはプレーヤーがコンテナ内に収まることを検証する。

```js
const fits = await page.locator('#pin-audio-runtime').evaluate((audio) =>
  audio.scrollWidth <= audio.clientWidth + 1
  && audio.getBoundingClientRect().right
    <= audio.parentElement.getBoundingClientRect().right + 1);
expect(fits).toBe(true);
```

- [ ] **Step 2: 共有と音声編集結果の契約を更新する**

`tests/browser/shared-audio.spec.js`と本番テンプレート統合テストで、`#pin-audio-runtime`が可視、`controls`がtrue、`controlsList`が`nodownload`、共有トークンだけが送信されることを確認する。

`tests/browser/audio-editor.spec.js`へ次を追加する。

```js
await expect(page.locator('[data-hae-result-audio]')).toHaveAttribute('controlsList', 'nodownload');
```

`tests/shared-audio-ui.test.js`では独自toggle配線の期待を削除し、retry配線とネイティブ要素を確認する。

```js
assert.match(sharedHtml, /pin-audio-runtime[^>]*controls[^>]*controlsList="nodownload"/);
assert.equal(sharedHtml.includes('pin-audio-player-toggle'), false);
assert.match(sharedHtml, /pin-audio-player-retry[\s\S]*?addEventListener\s*\(\s*['"]click['"]/);
```

- [ ] **Step 3: ブラウザ集中テストを実行する**

Run:

```powershell
npx playwright test tests/browser/audio-player.spec.js tests/browser/shared-audio.spec.js tests/browser/audio-editor.spec.js tests/browser/production-template-integration.spec.js
```

Expected: 全件PASS、pageerrorとunhandled rejectionが0。

- [ ] **Step 4: READMEを現在のUI仕様へ更新する**

主な機能を次の趣旨へ変更する。

```markdown
- 音声はピン詳細を開いたときだけ取得し、編集・プレビュー・共有URLでブラウザ標準プレーヤーにより再生
```

安全設計を次の趣旨へ変更する。

```markdown
- 音声付きピンではブラウザ標準プレーヤーを表示します。対応ブラウザではダウンロード項目だけを非表示にし、再生、一時停止、シーク、音量などの表示や外観はブラウザにより異なります。
- ダウンロードUIは非表示にしますが、ブラウザへ配信した音声の保存や取得を技術的に完全禁止するものではありません。
```

- [ ] **Step 5: 全Node検証を実行する**

Run:

```powershell
npm run check
```

Expected: `node --check Code.js`成功、全NodeテストPASS、FAIL 0。

- [ ] **Step 6: 全Playwright検証を実行する**

Run:

```powershell
npm run test:browser
```

Expected: 全PlaywrightテストPASS、FAIL 0。

- [ ] **Step 7: 差分と互換性を自己レビューする**

Run:

```powershell
git diff --check
git diff --stat
git status --short
```

確認事項:

- `Code.js`、サーバーAPI、永続化形式に差分がない。
- `index.html`と`shared.html`の共通ブロックがテストで同一。
- `controlsList`には`nodownload`以外を指定していない。
- 通常URL、共有認可、音声差し替え・削除後のinvalidateを弱めていない。
- 今回と無関係なファイルをコミットしない。

- [ ] **Step 8: ドキュメントとブラウザ回帰をコミットする**

```powershell
git add README.md tests/browser/audio-player.spec.js tests/browser/shared-audio.spec.js tests/browser/audio-editor.spec.js tests/browser/production-template-integration.spec.js tests/shared-audio-ui.test.js
git commit -m "test: cover native audio playback surfaces"
```

- [ ] **Step 9: 現在のブランチへプッシュする**

```powershell
git push origin v2.1.0
```

Expected: `v2.1.0 -> v2.1.0`、force pushなし。
