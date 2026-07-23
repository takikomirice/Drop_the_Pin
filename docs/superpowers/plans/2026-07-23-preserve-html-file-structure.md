# Preserve Two-File HTML Structure Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** v2.1.0の音声機能と遅延読込を維持したまま、Apps ScriptプロジェクトのHTMLを`index.html`と`shared.html`の2ファイルだけへ戻す。

**Architecture:** playerは同一のinert marker blockとして`index.html`と`shared.html`へ埋め込み、editorと生成済みvendorは`index.html`だけへ埋め込む。`Code.js`はraw `index.html`からvendor領域を検証・除去して初期テンプレートを生成し、認証済み`getAudioVendorBundle`だけが同じ領域のJavaScript sourceを返す。

**Tech Stack:** Google Apps Script HTML Service、Vanilla HTML/CSS/JavaScript、Node.js `node:test`、esbuild、Playwright Chromium

---

## File map

- Modify: `Code.js` — index raw sourceのvendor抽出・除去、stripped template生成、認証済みvendor取得。
- Modify: `index.html` — vendor、共通player、edit-token限定editorをinline化。
- Modify: `shared.html` —共通playerだけをinline化。
- Modify: `scripts/sync-audio-vendor.js` — `audio-vendor.html`ではなく`index.html`内vendor marker regionを決定的に更新。license同期は維持。
- Delete: `audio-editor.html` — 内容を`index.html`へ移動。
- Delete: `audio-player.html` — 内容を既存2 HTMLへ移動。
- Delete: `audio-vendor.html` — 生成先を`index.html`内へ移動。
- Create: `tests/html-file-structure.test.js` — HTML実在ファイルとinline block契約。
- Modify: `tests/audio-editor-integration.test.js` — inline editorとindex内vendorの認証・抽出・初期除外。
- Modify: `tests/audio-player-integration.test.js` — 2ページのplayer block同一性。
- Modify: `tests/audio-vendor.test.js` — index内生成物、license、sync/check契約。
- Modify: `tests/audio-editor-review-regressions.test.js` — editor include抽出をinline marker抽出へ変更。
- Modify: `tests/escape-overlay-dispatcher.test.js` — editor include抽出をinline marker抽出へ変更。
- Modify: `tests/shared-audio-ui.test.js` — shared inline player契約へ変更。
- Modify: `tests/browser/harness-server.js` — 追加HTML読込を廃止し、index/sharedのinline marker blockを利用。
- Modify only if failing assertions identify the old include contract: other `tests/*.test.js` files returned by `rg "audio-(editor|player|vendor)\\.html|includeHtml_" tests`.

### Marker contracts

共通playerは、表示やアクセシビリティへ影響しない次の境界で囲む。

```html
<template data-dtp-audio-player-boundary="start"></template>
<!-- existing audio-player.html contents, byte-identical in both pages -->
<template data-dtp-audio-player-boundary="end"></template>
```

editorはindexの既存edit-token条件内へ置く。

```html
<? if (editToken) { ?>
  <template data-dtp-audio-editor-boundary="start"></template>
  <!-- existing audio-editor.html contents -->
  <template data-dtp-audio-editor-boundary="end"></template>
<? } ?>
```

vendorはraw indexの先頭で、既存のplain sentinelを1組だけ使う。

```html
AUDIO_VENDOR_BUNDLE_START
<script>
/* generated bundle; literal closing script sequences are escaped */
</script>
AUDIO_VENDOR_BUNDLE_END
<!DOCTYPE html>
```

## Task 1: Commit the approved design and reviewed implementation plan

**Files:**
- Create: `docs/superpowers/specs/2026-07-23-preserve-html-file-structure-design.md`
- Create: `docs/superpowers/plans/2026-07-23-preserve-html-file-structure.md`

- [ ] **Step 1: Verify branch and remote have not moved**

Run:

```powershell
git fetch origin v2.1.0
git status --short --branch
git rev-parse HEAD
git rev-parse origin/v2.1.0
```

Expected: clean tracked worktree; local and remote hashes both equal the previously inspected base or are reconciled before continuing.

- [ ] **Step 2: Force-add the ignored reviewed documents**

`docs/superpowers/` is intentionally ignored, so add only these two exact files.

```powershell
git add -f docs/superpowers/specs/2026-07-23-preserve-html-file-structure-design.md docs/superpowers/plans/2026-07-23-preserve-html-file-structure.md
git diff --cached --check
```

Expected: only the two Markdown documents are staged and the check exits 0.

- [ ] **Step 3: Commit and push the documents**

```powershell
git commit -m "docs: plan two-file audio html structure"
git push origin v2.1.0
```

Expected: non-force push succeeds. Do not merge, tag, release, or deploy.

## Task 2: Establish RED tests for the two-file and inline-block contracts

**Files:**
- Create: `tests/html-file-structure.test.js`
- Modify: `tests/audio-player-integration.test.js`
- Modify: `tests/audio-editor-integration.test.js`
- Modify: `tests/audio-vendor.test.js`
- Modify: `tests/shared-audio-ui.test.js`

- [ ] **Step 1: Add a strict root HTML inventory test**

Use the actual directory contents, not only `git ls-files`:

```js
const htmlNames = fs.readdirSync(root, { withFileTypes: true })
  .filter((entry) => entry.isFile() && entry.name.endsWith('.html'))
  .map((entry) => entry.name)
  .sort();
assert.deepEqual(htmlNames, ['index.html', 'shared.html']);
```

- [ ] **Step 2: Add reusable exact marker extraction in the test**

The helper must reject missing, duplicate, or reversed boundaries, then compare the complete player regions from both files:

```js
function extractSingleRegion(source, start, end) {
  assert.equal((source.match(new RegExp(start, 'g')) || []).length, 1);
  assert.equal((source.match(new RegExp(end, 'g')) || []).length, 1);
  const startIndex = source.indexOf(start);
  const endIndex = source.indexOf(end, startIndex + start.length);
  assert.ok(startIndex >= 0 && endIndex > startIndex);
  return source.slice(startIndex + start.length, endIndex).trim();
}
```

Assert:

- player region exists once in both pages and is equal;
- each page contains exactly one `id="pin-audio-runtime"`;
- editor region exists only in index, inside `<? if (editToken) { ?>`;
- vendor sentinels exist once only in index;
- shared has neither editor nor vendor markers.

- [ ] **Step 3: Replace fragment-file assertions in focused tests**

Read editor/player/vendor regions from `index.html` and `shared.html`. Remove expectations for `includeHtml_` and the three added files. Preserve all behavioral assertions for limits, cancellation, MP3 output, playback state, `nodownload`, authorization, and no-download copy.

- [ ] **Step 4: Run the focused tests and verify RED**

Run:

```powershell
node --test tests/html-file-structure.test.js tests/audio-player-integration.test.js tests/audio-editor-integration.test.js tests/audio-vendor.test.js tests/shared-audio-ui.test.js
```

Expected: FAIL because five root HTML files exist and the inline markers are absent. Failures must be contract failures, not syntax or fixture errors.

- [ ] **Step 5: Verify the RED state against clasp's actual upload filter**

Use the worktree-local example project file explicitly so clasp does not discover the parent checkout's `.clasp.json`:

```powershell
$status = clasp -P .clasp.json.example -I .claspignore --json status | ConvertFrom-Json
$htmlToPush = @($status.filesToPush | Where-Object { $_ -like '*.html' } | Sort-Object)
if (($htmlToPush -join ',') -ne 'index.html,shared.html') { throw "Unexpected Apps Script HTML upload set: $($htmlToPush -join ', ')" }
```

Expected: throw in RED because the three added HTML files are still in `filesToPush`. This is the authoritative `.claspignore`-filtered check; do not use the parent checkout configuration discovered by a bare `clasp status`.

## Task 3: Establish RED tests for server-side vendor stripping and authentication

**Files:**
- Modify: `tests/audio-editor-integration.test.js`
- Modify: `tests/edit-key-config.test.js`

- [ ] **Step 1: Extend the HTML Service harness**

Track calls to:

- `createHtmlOutputFromFile('index')` for raw source;
- `createTemplate(strippedIndexSource)` for the index;
- `createTemplateFromFile('shared')` for shared.

The template double must capture assigned `execUrl`, `token`, and `editToken`, and `evaluate()` must expose the source used.

- [ ] **Step 2: Test index initial-output stripping**

Call `doGet` for index with a valid edit-key request and assert:

- the raw source is read once;
- the evaluated template source contains neither sentinel, generated source, nor `globalThis.Mediabunny`;
- the stripped template source, after optional leading whitespace removal, begins with `<!DOCTYPE html>`;
- `execUrl`, `token`, and nonempty `editToken` are assigned exactly as before;
- malformed/duplicate/out-of-order vendor markers fail closed rather than returning vendor-bearing HTML.

- [ ] **Step 3: Test authentication-before-index-read**

For missing/invalid `__editToken`, call `getAudioVendorBundle` and assert no `index` file read occurred. For a valid token, assert exactly one index read and a response shaped as `{ version: '1.50.8', source }` with no marker or `<script>` wrapper.

- [ ] **Step 4: Run focused tests and verify RED**

Run:

```powershell
node --test tests/audio-editor-integration.test.js tests/edit-key-config.test.js
```

Expected: FAIL because `doGet` still evaluates `index.html` directly and vendor retrieval still reads `audio-vendor.html`.

## Task 4: Generate vendor into index and inline editor/player

**Files:**
- Modify: `scripts/sync-audio-vendor.js`
- Modify: `index.html`
- Modify: `shared.html`
- Delete: `audio-editor.html`
- Delete: `audio-player.html`
- Delete: `audio-vendor.html`

- [ ] **Step 1: Add index region replacement to the vendor generator**

Define the same start/end sentinel constants as `Code.js`. Read `index.html`, require exactly one correctly ordered region, and replace from the start sentinel through the end sentinel with `buildVendorHtml().trimEnd()`. Add the resulting full index content to `desiredArtifacts()` instead of `audio-vendor.html`. Keep the two license artifacts unchanged.

The check path must report `index.html` when its embedded bundle is stale. Running sync twice must be byte-idempotent.

- [ ] **Step 2: Add a placeholder vendor region to index and run sync**

Insert one marker-wrapped `<script>` before `<!DOCTYPE html>`, then run:

```powershell
npm run vendor:sync
npm run vendor:check
```

Expected: both exit 0; the first updates only the vendor region and licenses if required; the second reports no mismatch.

- [ ] **Step 3: Inline player and editor without behavior edits**

Mechanically replace each player include with the marker-wrapped exact `audio-player.html` content. Mechanically replace the editor include inside the existing edit-token condition with the marker-wrapped exact `audio-editor.html` content. Do not rename DOM IDs, functions, globals, events, CSS selectors, copy, limits, playback behavior, or audio-editor public API.

- [ ] **Step 4: Remove the three added HTML files**

Delete only:

```text
audio-editor.html
audio-player.html
audio-vendor.html
```

Verify `Get-ChildItem -File -Filter *.html` reports only `index.html` and `shared.html`.

- [ ] **Step 5: Run the structure/vendor focused tests**

Run the Task 2 command again.

Expected: the root inventory and inline-region tests advance to GREEN; server-specific tests may remain RED until Task 5.

## Task 5: Strip vendor from index output and serve it only after authentication

**Files:**
- Modify: `Code.js`
- Modify: `tests/audio-editor-integration.test.js`
- Modify: `tests/edit-key-config.test.js`

- [ ] **Step 1: Replace the include layer with vendor region helpers**

Remove `HTML_INCLUDE_ALLOWLIST_` and `includeHtml_`, because no fragment files remain. Introduce one validator returning indexes and source:

```js
function locateAudioVendorRegion_(html) {
  const content = String(html || '');
  // require one start, one end, correct order, exact plain <script> wrapper,
  // nonempty escaped source, and one copy of both required public APIs
  return { startIndex, endIndex, source };
}

function stripAudioVendorRegion_(html) {
  const region = locateAudioVendorRegion_(html);
  return String(html).slice(0, region.startIndex)
    + String(html).slice(region.endIndex);
}
```

`endIndex` must include the end sentinel and its adjacent line ending so stripped output begins with `<!DOCTYPE html>` and contains no marker residue.

- [ ] **Step 2: Create stripped index templates**

For index only:

```js
const rawIndex = HtmlService.createHtmlOutputFromFile('index').getContent();
const template = HtmlService.createTemplate(stripAudioVendorRegion_(rawIndex));
```

Continue assigning `execUrl`, `token`, and `editToken` on this template exactly as before. Shared continues using `createTemplateFromFile('shared')`.

- [ ] **Step 3: Read index only after edit-token authentication**

Keep `assertEditToken_(payload)` as the first statement of `getAudioVendorBundle`. Then read raw `index.html`, call `locateAudioVendorRegion_`, and return only `{ version, source }`.

- [ ] **Step 4: Run focused server tests and syntax checks**

Run:

```powershell
node --check Code.js
node --test tests/audio-editor-integration.test.js tests/edit-key-config.test.js tests/html-file-structure.test.js tests/audio-vendor.test.js
```

Expected: all pass with no warnings.

## Task 6: Adapt remaining unit and browser harness extraction

**Files:**
- Modify: `tests/audio-editor-review-regressions.test.js`
- Modify: `tests/escape-overlay-dispatcher.test.js`
- Modify: `tests/browser/harness-server.js`
- Modify: any remaining old-fragment assertions identified by search

- [ ] **Step 1: Search for obsolete contracts**

Run:

```powershell
rg -n "audio-(editor|player|vendor)\\.html|includeHtml_|audio(Editor|Player|Vendor)Fragment|audioVendorFile" tests Code.js index.html shared.html scripts
```

Expected after the edits: no production references; test references only when explicitly asserting absence.

- [ ] **Step 2: Extract production marker regions in the browser harness**

Remove reads of the three deleted HTML files. Extract editor/player/vendor directly from `indexSource`; verify the shared player extraction equals the index player extraction. For `productionEditPage()`, remove the vendor region before serving and evaluate the edit-token conditional by removing only its Apps Script wrapper while retaining its inline content. For `productionSharedPage()`, no fragment replacement is needed.

- [ ] **Step 3: Preserve the deferred vendor endpoint**

Keep `/audio-vendor.js` in the local harness, backed by the source extracted from index. Initial page responses must not contain the vendor source. Existing browser assertions that no vendor resource is requested on initial/shared/100-pin flows must remain unchanged.

- [ ] **Step 4: Run all Node tests**

Run:

```powershell
npm run check
```

Expected: all Node tests pass; no missing-file or unprocessed template directive failures.

- [ ] **Step 5: Commit the implementation if the Node suite is green**

Before commit, inspect `git diff --check`, `git status --short`, and the exact diff. Commit only the planned source/test deletions and modifications:

```powershell
git add Code.js index.html shared.html scripts/sync-audio-vendor.js tests audio-editor.html audio-player.html audio-vendor.html
git commit -m "refactor: preserve two-file audio html structure"
```

Expected: one non-destructive commit; no push yet if browser verification remains.

## Task 7: Full browser, delivery-size, and regression verification

**Files:**
- No expected production changes; fix only demonstrated regressions using a new failing test first.

- [ ] **Step 1: Run complete Chromium coverage**

Run:

```powershell
npm run test:browser
```

Expected: all Chromium tests pass, including real MP3 generation, edit player, shared player, initial deferred vendor, and 100-pin metadata-only startup.

- [ ] **Step 2: Verify generated and deployment file contracts**

Run:

```powershell
npm run vendor:check
Get-ChildItem -File -Filter *.html | Select-Object -ExpandProperty Name
git ls-files "*.html"
$status = clasp -P .clasp.json.example -I .claspignore --json status | ConvertFrom-Json
$htmlToPush = @($status.filesToPush | Where-Object { $_ -like '*.html' } | Sort-Object)
if (($htmlToPush -join ',') -ne 'index.html,shared.html') { throw "Unexpected Apps Script HTML upload set: $($htmlToPush -join ', ')" }
```

Expected: vendor is current; filesystem, Git, and `.claspignore`-filtered clasp upload HTML lists contain only `index.html` and `shared.html`.

- [ ] **Step 3: Measure raw and stripped initial HTML**

Use a read-only Node helper or existing test harness to report:

- raw `index.html` bytes and gzip bytes;
- `index.html` after vendor-region removal, raw and gzip bytes;
- `shared.html` raw and gzip bytes;
- v2.0.0 equivalents when available through `git show v2.0.0:<file>`.

Acceptance: raw index may grow because it stores generated vendor/editor, but the normal server-render input/output path excludes the approximately 388KB vendor. Shared remains vendor/editor-free. Report measurements as evidence, not as a deployed GAS timing claim.

- [ ] **Step 4: Inspect security and privacy invariants**

Confirm tests cover:

- invalid vendor requests perform no index file read;
- initial/shared responses contain no vendor source or audio Base64;
- shared payloads contain no audio/Drive IDs;
- playback elements retain `controlsList="nodownload"`/`nodownload` behavior;
- no new token, ID, Base64, or filename logging.

- [ ] **Step 5: Request independent code review**

Review the complete diff against the approved spec. Resolve only concrete correctness, regression, security, or scope issues; rerun affected tests after each fix.

## Task 8: Synchronize Git safely

**Files:**
- No code changes expected.

- [ ] **Step 1: Fetch and compare before pushing**

Run:

```powershell
git fetch origin v2.1.0
git status --short --branch
git rev-parse HEAD
git rev-parse origin/v2.1.0
git log --oneline --decorate --graph -8
```

Expected: worktree clean; remote is an ancestor of local. If the user moved the remote, stop and reconcile without force-push or history rewriting.

- [ ] **Step 2: Push the verified commits**

```powershell
git push origin v2.1.0
```

Expected: normal fast-forward push succeeds.

- [ ] **Step 3: Verify remote identity**

Run:

```powershell
git ls-remote --heads origin v2.1.0
git rev-parse HEAD
```

Expected: local HEAD and remote branch hashes match. Do not run `clasp push`, create a deployment, merge to master, tag, release, or open a PR without separate authorization.
