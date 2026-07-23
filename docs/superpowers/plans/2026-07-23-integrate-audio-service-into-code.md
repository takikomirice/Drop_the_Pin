# AudioService.js Integration Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** `AudioService.js` の既存コードを挙動を変えずに `Code.js` へ統合し、ルートのApps Script JavaScriptを1ファイルにする。

**Architecture:** 現在のテスト評価順序 `Code.js` → `AudioService.js` を保つため、サービス本体を `Code.js` の末尾へそのまま移す。テストハーネスと構文検査は単一の `Code.js` だけを読み込み、ファイル構造テストで分離ファイルの再導入を防ぐ。

**Tech Stack:** Google Apps Script JavaScript、Node.js built-in test runner、PowerShell、Git

## Global Constraints

- 公開関数、レスポンス形状、シート列、Driveフォルダ構造、エラーコードを変更しない。
- `AudioService.js` の関数・定数は内容と相対順序を保って移動する。
- 統合と無関係なリファクタリング、命名変更、依存関係更新を行わない。
- 利用方法と設定は変わらないためREADMEは変更しない。

---

### Task 1: 単一JavaScriptファイル契約をテストで固定する

**Files:**
- Modify: `tests/html-file-structure.test.js`

**Interfaces:**
- Consumes: リポジトリルート直下の `.js` ファイル一覧。
- Produces: ルートのApps Script JavaScriptが `Code.js` だけであるという回帰契約。

- [ ] **Step 1: 失敗する構造テストを追加する**

`tests/html-file-structure.test.js` のHTMLファイル構造テストの直後へ追加する。

```js
test('the Apps Script project has exactly one deployable root JavaScript file', () => {
  const claspIgnoredPaths = new Set(
    fs.readFileSync(path.join(root, '.claspignore'), 'utf8')
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean)
  );
  const jsFiles = fs.readdirSync(root, { withFileTypes: true })
    .filter((entry) => entry.isFile() && path.extname(entry.name).toLowerCase() === '.js')
    .filter((entry) => !claspIgnoredPaths.has(entry.name))
    .map((entry) => entry.name)
    .sort();
  assert.deepEqual(jsFiles, ['Code.js']);
});
```

- [ ] **Step 2: REDを確認する**

Run: `node --test tests/html-file-structure.test.js`

Expected: 新規テストが `['AudioService.js', 'Code.js']` と `['Code.js']` の不一致で失敗する。

---

### Task 2: AudioServiceを移動して単一ファイル利用へ切り替える

**Files:**
- Modify: `Code.js`
- Delete: `AudioService.js`
- Modify: `tests/audio-storage-harness.js`
- Modify: `tests/drive-photo-import-server-harness.js`
- Modify: `tests/import-item-save.test.js`
- Modify: `package.json`

**Interfaces:**
- Consumes: `AudioService.js` の現在の全定数・全関数と、`Code.js` から参照される既存ヘルパー。
- Produces: `Code.js` だけで `ensureMediaDriveStructure`、`listDriveMediaInbox`、`readDriveAudioImportFile`、音声保存・再生・削除用内部関数を定義するサーバーコード。

- [ ] **Step 1: サービス本体を評価順序どおり移動する**

`Code.js` の最後の `module.exports` ブロックの後へ空行を1行置き、`AudioService.js` の先頭 `var MEDIA_DRIVE_NAMES = Object.freeze({` から末尾 `readPinAudioBlobByPinId_` の閉じ括弧までを順序・内容を変えずに追加する。その後 `AudioService.js` を削除する。境界は次の形にする。

```js
    invalidateRouteCacheForRoute: invalidateRouteCacheForRoute,
    testRouteCRUD: testRouteCRUD,
    setupSheet: setupSheet
  };
}

var MEDIA_DRIVE_NAMES = Object.freeze({
  photos: 'photos',
  audio: 'audio',
  original: 'original',
  guide: 'ここに直接ファイルを入れてください.txt'
});
```

- [ ] **Step 2: テストハーネスを `Code.js` 単体評価へ変更する**

3ファイルから次の読込定義を削除する。

```js
const audioServiceJs = fs.readFileSync(path.join(root, 'AudioService.js'), 'utf8');

const audioServicePath = path.join(root, 'AudioService.js');
const audioServiceJs = fs.existsSync(audioServicePath)
  ? fs.readFileSync(audioServicePath, 'utf8') : '';

const audioServiceJs = fs.readFileSync(path.join(__dirname, '..', 'AudioService.js'), 'utf8');
```

評価文字列は、`tests/audio-storage-harness.js` で次の1行へ置換する。

```js
vm.runInContext(codeJs + '\n' + [
```

`tests/drive-photo-import-server-harness.js` では連結式を次の1行へ置換する。

```js
codeJs + '\n' + [
```

`tests/import-item-save.test.js` ではテンプレート文字列の先頭を次の1行へ置換する。各ファイルの後続API公開行は変更しない。

```js
`${codeJs}\nthis.__api = {\n`
```

- [ ] **Step 3: 構文検査を単一ファイル前提へ変更する**

`package.json` の `check` を次の値にする。

```json
"check": "node --check Code.js && node --test tests/*.test.js"
```

- [ ] **Step 4: GREENと音声・Drive回帰を確認する**

Run: `node --test tests/html-file-structure.test.js tests/audio-storage.test.js tests/audio-lifecycle.test.js tests/audio-player-integration.test.js tests/shared-audio-access.test.js tests/drive-photo-import-server.test.js tests/import-item-save.test.js`

Expected: 全テストがpassし、失敗が0件になる。

---

### Task 3: 全体検証、自己レビュー、コミット、プッシュ

**Files:**
- Review: `Code.js`, `package.json`, `tests/html-file-structure.test.js`, `tests/audio-storage-harness.js`, `tests/drive-photo-import-server-harness.js`, `tests/import-item-save.test.js`

**Interfaces:**
- Consumes: Task 1とTask 2の統合済みコードとテスト。
- Produces: 現在ブランチへコミット・プッシュされた、検証済みの単一JavaScript構成。

- [ ] **Step 1: 構文検査と全Nodeテストを実行する**

Run: `npm run check`

Expected: `node --check Code.js` と全Nodeテストがexit code 0で完了する。

- [ ] **Step 2: 生成済みvendorの同期を検査する**

Run: `npm run vendor:check`

Expected: 差分なしでexit code 0になる。

- [ ] **Step 3: ブラウザ回帰テストを実行する**

Run: `npm run test:browser`

Expected: Playwrightの全テストがpassする。

- [ ] **Step 4: 差分を自己レビューする**

Run: `git diff --check`

Expected: 出力なし、exit code 0。

Run: `git diff --stat`

Expected: `Code.js` への798行追加、`AudioService.js` の同量削除、4つのテスト関連ファイルと `package.json` の小さな変更だけが表示される。

Run: `rg -n "AudioService\\.js|audioServiceJs|audioServicePath" package.json tests Code.js`

Expected: 該当なし、exit code 1。

- [ ] **Step 5: 実装差分だけをコミットする**

```powershell
git add Code.js AudioService.js package.json tests/html-file-structure.test.js tests/audio-storage-harness.js tests/drive-photo-import-server-harness.js tests/import-item-save.test.js docs/superpowers/plans/2026-07-23-integrate-audio-service-into-code.md
git commit -m "refactor: integrate AudioService into Code"
```

- [ ] **Step 6: 現在ブランチを通常プッシュする**

Run: `git push origin v2.1.0`

Expected: 強制プッシュなしで `origin/v2.1.0` が新しいコミットへ進む。
