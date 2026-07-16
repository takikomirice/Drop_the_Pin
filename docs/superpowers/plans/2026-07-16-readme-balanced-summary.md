# README Balanced Summary Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** READMEを利用者・導入者向けのバランス要約へ整理し、最終内容をv2.0.0としてGitHub Releaseへ公開する。

**Architecture:** READMEを単一の入口として維持し、現在の機能、導入、運用、安全、主要仕様を本文に残す。フェーズ別説明、仕様、手動確認、変更履歴は詳細列挙をやめて機能カテゴリ単位の要約へ置き換え、v2.0.0の個別変更はGitHub Releaseを正本にする。

**Tech Stack:** Markdown、Node.js組み込みテストランナー、Git、GitHub CLI

## Global Constraints

- 厳密な行数上限は設けず、重複と内部実装の説明を優先して削る。
- フェーズ別説明、主要仕様、手動確認、変更履歴は削除せず要約して残す。
- 既存データとの互換性、初期設定の再実行、対応形式、主要な件数・容量上限を残す。
- 未実施のApps Script・実ブラウザ確認を「確認済み」と記載しない。
- v2.0.0の詳細はGitHub Releaseを正本とし、コミット列挙は行わない。
- Release公開が確認できるまでv1.4.0ブランチを削除しない。

---

### Task 1: READMEをバランス要約へ置き換える

**Files:**
- Modify: `README.md`
- Modify: `tests/csv-interchange-ui.test.js`
- Modify: `tests/geojson-import-workflow.test.js`
- Modify: `tests/geojson-track-import-workflow.test.js`
- Modify: `tests/gpx-track-import-ui.test.js`
- Modify: `tests/import-item-save.test.js`
- Modify: `tests/multi-photo-import-builder.test.js`
- Reference: `docs/superpowers/specs/2026-07-16-readme-balanced-summary-design.md`
- Reference: `RELEASE_NOTES.md`
- Test: `tests/*.test.js`

**Interfaces:**
- Consumes: 現在のREADMEに記載された機能、設定、制限、確認状況
- Produces: 利用者と導入者が単独で参照できる要約版README

- [ ] **Step 1: 現在の見出しと保持対象を確認する**

Run:

```powershell
rg -n '^#{1,4} ' README.md
```

Expected: 製品概要、機能、安全、セットアップ、データ構造、各Phase、テスト、バージョン、注意点の見出しが確認できる。

- [ ] **Step 2: README本文を次の構成へ再編集する**

`README.md`を以下の順序にする。

```markdown
# Drop the Pin

Google Apps Script と Leaflet.js で動作する、写真・ピン・ルートを共有できる地図投稿アプリの概要。

## 主な機能
- 写真付き／写真なしピン、GPS・EXIF、未配置ピン
- 端末／Driveからの複数写真取込、HEIC／HEIF変換
- CSV／GeoJSON Point入出力、GPX／GeoJSONトラック取込
- ピンルートとトラックの管理、共有ビュー
- 検索、フィルタ、プリセット、一括操作、レスポンシブUI

## 対応形式と主な上限
写真形式、20件取込、2MBデータ交換、5MB GPX、15MB Drive写真、20,000トラックpoint、200 segment、100 trackを表で示す。

## 安全設計と外部サービス
編集URL、共有リンク、位置情報、Drive共有、OpenStreetMap、Nominatim、OSRMの注意を要約する。

## 技術構成
Apps Script、HTML、Leaflet、Drive、Spreadsheet、Node.jsテストを列挙する。

## セットアップとデプロイ
clasp導入、Apps Script／Spreadsheet／Drive準備、setupSheet、clasp push、Webアプリ再デプロイ、編集URL確認を短い手順で示す。

## データ構造
map_info、share_links、routes、route_pins、route_cache、tracks、track_segments、input_presets、import_receiptsの用途を表で示す。既存行の互換性と更新後の初期設定再実行を明記する。

## フェーズ別概要
基盤、共有・安全、複数写真・データ交換、トラック、写真時刻照合・Drive、UI統合の6区分を各1〜3文で要約する。

## テストと手動確認
node --testコマンドを示し、実機確認が未実施であることを明記する。初期設定、写真、CSV／GeoJSON、GPX／GeoJSON track、共有、端末互換の6カテゴリを代表項目へ集約する。

## バージョン
v1.0.0〜v2.0.0を各1行で要約し、GitHub Releasesへのリンクを付ける。

## 運用上の注意
timezone、Drive共有、READMEがclasp対象外であることを残す。

## ライセンス
MIT Licenseへのリンク。
```

各節では公開API名、内部クラス名、receiptの状態遷移、詳細エラーコード、個別のテストケース列挙を削除する。これらを固定していたREADME契約テストは、対応形式、主要上限、安全上の注意、実機確認状況を検証する要約後の契約へ更新する。

- [ ] **Step 3: READMEの要約状態を確認する**

Run:

```powershell
(Get-Content README.md).Count
rg -n '^## ' README.md
rg -n '未実施|setupSheet|20,000|15MB|v2.0.0|GitHub Releases' README.md
```

Expected: 必須セクションと運用情報が存在し、元の728行から明確に短縮されている。厳密な最大行数は判定条件にしない。

### Task 2: README変更を検証してコミットする

**Files:**
- Modify: `README.md`
- Modify: `tests/csv-interchange-ui.test.js`
- Modify: `tests/geojson-import-workflow.test.js`
- Modify: `tests/geojson-track-import-workflow.test.js`
- Modify: `tests/gpx-track-import-ui.test.js`
- Modify: `tests/import-item-save.test.js`
- Modify: `tests/multi-photo-import-builder.test.js`
- Test: `tests/*.test.js`

**Interfaces:**
- Consumes: Task 1の要約版README
- Produces: 検証済みREADMEコミット

- [ ] **Step 1: Markdown差分の空白エラーを確認する**

Run:

```powershell
git diff --check
git diff --stat
git diff -- README.md
```

Expected: `git diff --check`は出力なし・終了コード0。差分はREADMEの要約とREADME契約テストだけで、製品コードを含まない。

- [ ] **Step 2: 全ローカルテストを実行する**

Run:

```powershell
node --test tests/*.test.js
```

Expected: 1,354 tests、0 failures。

- [ ] **Step 3: READMEをコミットする**

Run:

```powershell
git add README.md tests/csv-interchange-ui.test.js tests/geojson-import-workflow.test.js tests/geojson-track-import-workflow.test.js tests/gpx-track-import-ui.test.js tests/import-item-save.test.js tests/multi-photo-import-builder.test.js docs/superpowers/plans/2026-07-16-readme-balanced-summary.md
git commit -m "docs: summarize README for v2.0.0"
```

Expected: READMEだけを含む新しいコミットが作成される。

### Task 3: masterとv2.0.0タグを公開する

**Files:**
- Read: `.git/` repository state

**Interfaces:**
- Consumes: 設計、計画、README要約を含むmaster最終コミット
- Produces: GitHub上のmasterと、そのコミットを指す注釈付きv2.0.0タグ

- [ ] **Step 1: 公開前状態を確認する**

Run:

```powershell
git status --short --branch
git log -3 --oneline
gh auth status
```

Expected: 作業ツリーがクリーン、masterがorigin/masterより必要なコミット数だけ進み、GitHub認証が有効。

- [ ] **Step 2: masterをpushする**

Run:

```powershell
git push origin master
```

Expected: origin/masterがローカルmasterと一致する。

- [ ] **Step 3: 公開前タグを最終コミットへ更新する**

Run:

```powershell
git tag -fa v2.0.0 -m "v2.0.0"
git push --force origin refs/tags/v2.0.0
```

Expected: `refs/tags/v2.0.0^{}`がmasterのHEADと同じcommitを指す。Release未公開のタグ更新に限定し、masterはforce-pushしない。

### Task 4: GitHub Releaseを公開する

**Files:**
- Create temporarily: `.release-v2.0.0.md`
- Delete after publish: `.release-v2.0.0.md`

**Interfaces:**
- Consumes: v2.0.0タグと承認済み日本語リリースノート
- Produces: タイトル`v2.0.0`の公開GitHub Release

- [ ] **Step 1: 承認済みノートを一時ファイルへ用意する**

本文は「主な新機能」「改善」「不具合修正」「注意点」「実機確認済み項目」の順とする。破壊的変更がないこと、初期設定の再実行、実機未確認、ローカル自動テスト1,354件成功を明記する。

- [ ] **Step 2: Releaseを公開する**

Run:

```powershell
gh release create v2.0.0 --title v2.0.0 --notes-file .release-v2.0.0.md --verify-tag
```

Expected: draftでもprereleaseでもないv2.0.0 Release URLが返る。

- [ ] **Step 3: 公開状態を確認して一時ファイルを削除する**

Run:

```powershell
gh release view v2.0.0 --json name,tagName,isDraft,isPrerelease,url
```

Expected: `name`と`tagName`が`v2.0.0`、`isDraft`と`isPrerelease`がfalse。

### Task 5: v1.4.0ブランチを削除して最終確認する

**Files:**
- Read: `.git/` repository state

**Interfaces:**
- Consumes: 公開確認済みv2.0.0 Release
- Produces: v1.4.0ローカル・リモートブランチが存在しない完了状態

- [ ] **Step 1: リモートブランチを削除する**

Run:

```powershell
git push origin --delete v1.4.0
```

Expected: GitHubの`refs/heads/v1.4.0`が削除される。

- [ ] **Step 2: ローカルブランチを削除する**

Run:

```powershell
git branch -D v1.4.0
```

Expected: squash mergeのため未マージ扱いのローカルv1.4.0を、ユーザーの明示指示に基づいて削除する。

- [ ] **Step 3: 最終状態を検証する**

Run:

```powershell
git fetch origin --prune
git status --short --branch
git branch --all
git ls-remote origin refs/heads/master refs/heads/v1.4.0 'refs/tags/v2.0.0*'
gh release view v2.0.0 --json name,tagName,isDraft,isPrerelease,url
```

Expected: 作業ツリーはクリーン、masterとorigin/masterは一致、v1.4.0はローカル・リモートとも不在、v2.0.0タグはmaster HEADを指し、Releaseは公開済み。
