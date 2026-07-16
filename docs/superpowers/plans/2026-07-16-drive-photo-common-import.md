# Drive Photo Common Import Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Drive写真を安全なブラウザ`File`へ変換した直後から端末写真と同じ複数写真取込経路へ合流させ、キャンセル時に単品ピン画面を開かない。

**Architecture:** サーバーは実読込bytesを上限判定とレスポンスサイズの根拠にし、クライアントは要求IDと実データだけを検証して`File`化する。写真取込元選択・共通準備への遷移を単品ピンフォームから分離し、Drive固有値は既存runtime/payloadの`sourceDriveFileId`だけに限定する。

**Tech Stack:** Google Apps Script JavaScript、ブラウザFile/Blob API、Node.js built-in test runner、VMテストハーネス

## Global Constraints

- 新しいバージョン管理、設定項目、Sheet列、Drive専用Preview、Drive専用保存API、Advanced Drive APIは追加しない。
- `readDrivePhotoImportFile()`の編集権限、ルート配下、trash、shortcut、形式、15MB上限の検証を維持する。
- Driveと端末は`startPhotoImportFromFiles()`以降のBuilder、Preview、processor、save、onSavedを共有する。
- 単品ピンフォームの入力値をDrive開始・キャンセル・失敗で変更または初期化しない。

---

### Task 1: DriveレスポンスとFile化の契約を固定する

**Files:**
- Modify: `tests/drive-photo-import-source-core.test.js`
- Modify: `tests/drive-photo-import-loader.test.js`
- Modify: `tests/drive-photo-import-integration.test.js`
- Verify: `Code.js`
- Verify: `index.html`

**Interfaces:**
- Consumes: `DrivePhotoImportSourceCore.validateFileResponse(response, expectedDescriptor)`、`materializeFile(response, environment)`、`DrivePhotoImportLoader.create(options).start(descriptors)`
- Produces: 一覧descriptorに依存せず、要求IDと復号実データから生成された`File[]`

- [ ] **Step 1: 対応5形式と拒否条件の失敗テストを書く**

```js
for (const sample of [
  ['photo.jpg', 'image/jpeg'], ['photo.png', 'image/png'],
  ['photo.webp', 'image/webp'], ['photo.heic', 'image/heic'],
  ['photo.heif', 'image/heif']
]) {
  const validated = sourceCore.validateFileResponse(fileResponse(
    descriptor({ name: sample[0], mimeType: sample[1] })
  ), descriptor());
  const file = sourceCore.materializeFile(validated, environment());
  assert.equal(file.type, sample[1]);
}
```

ID不一致、不正Base64、非対応形式、復号後0byte、15MB超も同じ公開APIで拒否されることを追加する。

- [ ] **Step 2: REDを確認する**

Run: `node --test tests/drive-photo-import-source-core.test.js tests/drive-photo-import-loader.test.js tests/drive-photo-import-integration.test.js`

Expected: 対応5形式の一括File化または実データ境界の新規assertionが現行契約との差分でFAILする。

- [ ] **Step 3: 必要最小限のレスポンス検証・File生成を実装する**

`validateFileResponse()`では`ok`、要求ID、安全なbasename、`PhotoImportFileTypeCore.classify()`、正規Base64、復号サイズだけを検証する。`materializeFile()`は復号結果の1〜15MBを再確認し、有効な`modifiedAt`だけを`lastModified`へ使う。`Code.js`はblob bytes lengthを`sizeBytes`と上限判定に使い、`getSize()`との一致を要求しない。

- [ ] **Step 4: GREENを確認する**

Run: `node --test tests/drive-photo-import-source-core.test.js tests/drive-photo-import-loader.test.js tests/drive-photo-import-integration.test.js`

Expected: 全テストPASS。

### Task 2: Drive pickerの戻り先と単品フォームを分離する

**Files:**
- Modify: `tests/drive-photo-import-ui.test.js`
- Modify: `tests/drive-photo-import-workflow.test.js`
- Modify: `tests/drive-photo-import-integration.test.js`
- Modify: `index.html`

**Interfaces:**
- Consumes: `DrivePhotoImportUI.create({ closePicker, closeConsent, returnToPhotoSource, onFilesReady })`
- Produces: picker cancel/handoff failure/preparation cancelが`add-menu-overlay`へ戻る遷移と、単品フォーム非変更保証

- [ ] **Step 1: キャンセル・失敗・成功遷移の失敗テストを書く**

```js
const picker = createPicker({
  returnToPhotoSource: (message) => calls.push(['source', message])
});
await picker.controller.open();
assert.equal(picker.controller.cancel(), true);
assert.equal(calls.some(([name]) => name === 'source'), true);
```

production wiringについて`returnToPhotoSource`が`add-menu-overlay`だけを開き、`resetUploadState()`、`loadUploadInputPresets()`、`openOverlay('upload-overlay')`を呼ばないことをassertする。loader rejectでは`state.open === true`、picker close/戻り呼出しなし、成功時だけ`multi-photo-preparation-overlay`が開くことも追加する。

- [ ] **Step 2: REDを確認する**

Run: `node --test tests/drive-photo-import-ui.test.js tests/drive-photo-import-workflow.test.js tests/drive-photo-import-integration.test.js`

Expected: 現在の`returnToUpload`、`resetUploadState()`、`upload-overlay`復帰を検出してFAILする。

- [ ] **Step 3: 写真取込専用の初期値と復帰処理を実装する**

`snapshotPhotoImportDefaults()`は空タグ、既定色、既定アイコン、既定状態、アプリのroot folderを返し、単品フォームを読まない。`prepareAddMenuPhotoImport()`から`resetUploadState()`を削除する。Drive UI callbackを`returnToPhotoSource`へ改名し、`closePicker()`はpickerを閉じるだけ、`returnToPhotoSource()`は準備overlayを閉じて`add-menu-overlay`を開き、安全なメッセージを`multi-photo-conflict-message`へ表示する。`returnFromMultiPhotoPreparation()`と`startPhotoImportFromFiles()`の失敗も同じ復帰処理を使う。

- [ ] **Step 4: GREENを確認する**

Run: `node --test tests/drive-photo-import-ui.test.js tests/drive-photo-import-workflow.test.js tests/drive-photo-import-integration.test.js`

Expected: 全テストPASS。

### Task 3: 共通経路と全回帰を検証する

**Files:**
- Review: `index.html`
- Review: `Code.js`
- Review: `tests/drive-photo-import-*.test.js`

**Interfaces:**
- Consumes: Drive生成`File[]`、`sourceDriveFileIds[]`、既存`startPhotoImportFromFiles()`
- Produces: 端末/Drive共通のready job、共通Preview、`saveImportPhotoItem()` payload、`onSaved`反映

- [ ] **Step 1: 禁止構造と共通経路を静的・統合テストで確認する**

```js
assert.equal(indexHtml.includes('DrivePhotoImportPreview'), false);
assert.equal(indexHtml.includes('saveDrivePhotoImportItem'), false);
assert.match(indexHtml, /onFilesReady:[\s\S]*startPhotoImportFromFiles/);
```

- [ ] **Step 2: 指定検証をすべて実行する**

Run:

```text
node --check Code.js
node --test tests/drive-photo-import-source-core.test.js
node --test tests/drive-photo-import-loader.test.js
node --test tests/drive-photo-import-ui.test.js
node --test tests/drive-photo-import-workflow.test.js
node --test tests/drive-photo-import-integration.test.js
node --test tests/*.test.js
```

Expected: 各コマンドexit 0、fail 0。

- [ ] **Step 3: 差分を自己レビューする**

`git diff --check`、`git diff --stat`、`git diff -- Code.js index.html tests`で無関係な整形、公開payload変更、Sheet列追加、Advanced Drive API、Drive専用Preview/save API、単品フォーム変更がないことを確認する。

- [ ] **Step 4: 今回分だけをコミットしてpushする**

```text
git add index.html Code.js tests/drive-photo-import-source-core.test.js tests/drive-photo-import-loader.test.js tests/drive-photo-import-ui.test.js tests/drive-photo-import-workflow.test.js tests/drive-photo-import-integration.test.js docs/superpowers
git commit -m "fix: unify Drive photo import workflow"
git push origin v1.4.0
```
