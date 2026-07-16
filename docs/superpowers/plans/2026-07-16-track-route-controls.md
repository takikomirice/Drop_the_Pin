# Track Route Controls Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox notation so progress can be audited.

**Goal:** GPX/GeoJSON の線幅を常に 4 に固定し、色をスウォッチで選べる取込・編集UIと、ピンルート／取込ルートの統一されたアイコン操作を提供する。

**Architecture:** 既存の取込プレビューオーバーレイを新規取込と保存済みトラック編集で共用する。クライアントは取込と編集の状態を分離し、保存済みトラックの表示設定だけを更新する専用 Apps Script API を呼ぶ。サーバーは既存形状・順序・共有情報を保持し、通常表示・共有表示・再編集のすべてで線幅を 4 に正規化する。

**Tech Stack:** Google Apps Script (`Code.js`)、vanilla JavaScript/CSS (`index.html`)、Node.js built-in test runner (`node:test`)

## Global Constraints

- 作業ブランチは `v1.4.0` のまま使用し、既存の未関連差分は変更しない。
- GPX/GeoJSON の線幅入力は表示しない。新規保存・編集保存・通常表示・共有表示はいずれも `lineWidth: 4` とする。
- 既存シートに保存済みの 1〜10 の線幅値は読み込み不能にせず、読み出し時に 4 として扱う。
- 保存済みトラック編集で変更できるのは名前・説明・色・線種・表示状態だけ。形状、日時、順序、共有リンク、作成日時、由来情報は保持する。
- ピンルートの操作は 44px のアイコンのみを「追加・時系列・Undo・編集・削除」の順で並べる。
- GPX/GeoJSON カードは既存の表示切替と地図フィットを残し、右端に 44px の編集・削除アイコンを置く。共有／読み取り専用画面には編集・削除を出さない。
- すべてのボタンに `aria-label` と `title` を持たせ、PCとモバイルで同じDOMと動作を使う。

---

## Task 1: Fix imported track line width and replace color select with swatches

**Files:**

- Modify: `index.html`
- Modify: `Code.js`
- Test: `tests/gpx-track-import-ui.test.js`
- Test: `tests/geojson-track-import-ui.test.js`
- Test: `tests/gpx-track-import-workflow.test.js`
- Test: `tests/geojson-track-import-workflow.test.js`
- Test: `tests/track-storage.test.js`

- [ ] Add failing UI tests asserting that the preview contains a palette container, does not contain `track-import-preview-line-width`, does not construct color `<option>` elements, and uses the existing `renderColorPaletteButtons` helper.
- [ ] Run `node --test tests/gpx-track-import-ui.test.js tests/geojson-track-import-ui.test.js` and confirm the new assertions fail for the current `<select>` and range/number input.
- [ ] Add failing workflow/storage assertions that an attempted non-4 line width is normalized to 4 on the client payload, server save result, normal DTO, and legacy-row DTO.
- [ ] Run `node --test tests/gpx-track-import-workflow.test.js tests/geojson-track-import-workflow.test.js tests/track-storage.test.js` and confirm the new fixed-width assertions fail.
- [ ] Replace the color `<select>` in `index.html` with a palette container that stores the selected color in `container.value`; remove all line-width markup and form lookups.
- [ ] Refactor the import preview renderer to call a shared helper similar to:

  ```js
  function renderTrackSettingsColorPalette(selectedColor, disabled) {
    renderColorPaletteButtons(colorContainer, selectedColor, handleColorChange, {
      disabled: disabled
    });
  }
  ```

- [ ] Remove line-width validation from `validateTrackImportPreviewForm()` and force `lineWidth: 4` when building a draft/payload.
- [ ] In `Code.js`, keep accepting legacy stored widths where required for parsing, but force normalized saved metadata and returned metadata to `lineWidth: 4`.
- [ ] Re-run the five focused tests and confirm they pass.
- [ ] Commit this logical unit with `refactor: fix imported track display controls`.

## Task 2: Add a server API for saved track display settings

**Files:**

- Modify: `Code.js`
- Modify: `tests/track-storage.test.js`
- Modify: `tests/edit-token-guard.test.js`
- Modify: `tests/gas-retry-policy.test.js`

- [ ] Add failing tests for `updateTrackDisplaySettings(data)` covering success, missing/invalid edit token, unknown track, invalid name/color/line style/visibility, preservation of geometry/order/revision/source/share fields/formulas/extension columns, fixed line width 4, and retry classification.
- [ ] Run `node --test tests/track-storage.test.js tests/edit-token-guard.test.js tests/gas-retry-policy.test.js` and confirm failures identify the missing API.
- [ ] Implement a narrow normalizer for `{ trackId, name, description, color, visible, lineStyle }`, reusing existing validation/constants and setting `lineWidth: 4` internally.
- [ ] Implement `updateTrackDisplaySettings(data)` using the edit-token guard and `LockService.getScriptLock()`. Find the current row by normalized track ID and update only metadata cells while copying formulas and extension columns unchanged.
- [ ] Return a whitelisted settings DTO, for example:

  ```js
  {
    trackId,
    name,
    description,
    color,
    visible,
    lineStyle,
    lineWidth: 4,
    updatedAt
  }
  ```

- [ ] Map storage/lock/not-found/validation failures through the existing structured result conventions; do not return raw server exceptions.
- [ ] Add the mutation name to retry/edit-token policy allowlists where applicable.
- [ ] Re-run the three focused tests and confirm they pass.
- [ ] Commit this logical unit with `feat: update saved track display settings`.

## Task 3: Reuse the import overlay for editing saved tracks

**Files:**

- Modify: `index.html`
- Add: `tests/track-settings-ui.test.js`
- Modify: `tests/route-pending-mutation.test.js`
- Modify: `tests/track-deletion-ui.test.js`

- [ ] Add failing static/runtime tests for opening an existing track, pre-filling name/description/color/line style/visibility, showing fixed width only implicitly, saving through `updateTrackDisplaySettings`, merging returned metadata without losing geometry, cancel/Escape behavior, retryable errors, and disabled controls during a pending save.
- [ ] Run `node --test tests/track-settings-ui.test.js tests/route-pending-mutation.test.js tests/track-deletion-ui.test.js` and confirm the edit-flow assertions fail.
- [ ] Add a dedicated `state.trackEdit` record with `trackId`, snapshot, and pending/error state; do not overload the import controller's file lifecycle state.
- [ ] Extract small overlay helpers for title, summary visibility, field values, palette rendering, validation, and button labels so import and edit use the same controls.
- [ ] Implement `openTrackDisplaySettingsEditor(trackId)`, edit save/cancel/retry handlers, and mode-aware save/discard/close/Escape dispatch.
- [ ] Add the edit mutation to `hasPendingMutationWork()` and related busy guards so sorting, deletion, closing, and duplicate submissions cannot race it.
- [ ] On success, merge only the server-returned settings into the existing normalized track and rerender layers/cards. On failure, preserve the user's draft and expose the existing retry/cancel UI.
- [ ] Re-run the three focused tests and confirm they pass.
- [ ] Commit this logical unit with `feat: edit imported route settings`.

## Task 4: Unify pin-route and imported-route card actions and heights

**Files:**

- Modify: `index.html`
- Modify: `tests/unified-route-ui.test.js`
- Modify: `tests/route-card-visibility-static.test.js`
- Modify: `tests/action-button-standards.test.js`
- Modify: `tests/track-deletion-ui.test.js`

- [ ] Add failing tests that pin routes render five icon-only 44px controls in the required order, that delete is no longer inside the details/edit overlay, and that imported routes render edit/delete controls only in editable mode.
- [ ] Add a failing card-structure assertion that imported cards do not carry the legacy `.track-item` class/padding that makes their collapsed height differ from pin routes.
- [ ] Run `node --test tests/unified-route-ui.test.js tests/route-card-visibility-static.test.js tests/action-button-standards.test.js tests/track-deletion-ui.test.js` and confirm the new assertions fail.
- [ ] Change the pin action row to use existing action-icon creation helpers for Add, Chronological, Undo, Edit, and Delete. Remove visible text labels while preserving accessible labels/tooltips and disabled states.
- [ ] Move pin deletion out of the details overlay into its own fifth action, reusing the existing confirmation and deletion mutation.
- [ ] In editable imported cards, retain current visibility and fit controls, then place Edit/Delete icon buttons in a right-aligned action group. Hide both in shared/read-only rendering.
- [ ] Remove the legacy `.track-item` class from unified imported route cards and adjust only the minimal CSS needed to make collapsed height match the pin-route baseline.
- [ ] Re-run the focused tests and inspect both desktop/mobile CSS assertions.
- [ ] Commit this logical unit with `refactor: unify route card actions`.

## Task 5: Documentation, regression verification, and delivery

**Files:**

- Modify: `README.md` only where import/edit controls or line-width behavior are documented
- Review: `Code.js`
- Review: `index.html`
- Review: `tests/*.test.js`

- [ ] Update README statements that advertise editable line width or omit saved-track settings editing. Do not add unrelated documentation changes.
- [ ] Run syntax and focused suites:

  ```powershell
  node --check Code.js
  node --test tests/gpx-track-import-ui.test.js tests/geojson-track-import-ui.test.js tests/gpx-track-import-workflow.test.js tests/geojson-track-import-workflow.test.js
  node --test tests/track-storage.test.js tests/edit-token-guard.test.js tests/gas-retry-policy.test.js
  node --test tests/track-settings-ui.test.js tests/unified-route-ui.test.js tests/route-card-visibility-static.test.js tests/route-pending-mutation.test.js tests/track-deletion-ui.test.js
  ```

- [ ] Run broad regressions:

  ```powershell
  node --test tests/*track*.test.js
  node --test tests/*route*.test.js
  node --test tests/*shared*.test.js
  node --test tests/*.test.js
  git diff --check
  ```

- [ ] Review the final diff for accidental public-format changes, geometry mutation, share-link mutation, authorization weakening, stale text labels, missing mobile behavior, or unrelated formatting.
- [ ] If self-review finds an issue, fix it and rerun the affected focused suite plus the full suite.
- [ ] Commit any final README/test-only cleanup with a specific message if it is not naturally part of the preceding commits.
- [ ] Verify `git status --short --branch`, push `v1.4.0` to `origin`, and report the final commit SHA, push destination, test counts/results, README status, compatibility impact, and any residual risk.
