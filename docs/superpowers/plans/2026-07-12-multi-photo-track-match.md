# Phase 8-2 Multi-photo Track Match Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Connect `PhotoTrackMatchCore` to only the multi-photo Preview and apply explicitly selected safe location candidates through the existing import-save flow.

**Architecture:** Add a DOM-free `PhotoTrackMatchPreviewCore` for option parsing, track filtering, ownership, stale detection, matching, selection, and application. Add a lat/lng-only immutable API to `ImportJobCore`, expose a fixed track-match panel through a dedicated `ImportPreviewUI` configuration object, and inject `state.tracks` only from the production multi-photo workflow.

**Tech Stack:** Google Apps Script HTML client JavaScript, DOM APIs, Node.js `node:test`, VM-based module harnesses.

## Global Constraints

- Only `sourceType === "multi-photo"` exposes matching.
- Never mutate coordinates during Preview open or match execution; only explicit apply changes `lat` / `lng`.
- Never overwrite original GPS, partial GPS, or manually edited coordinates.
- Do not add GAS APIs, Sheets, routes, GPX parsing, or persistence fields.
- Keep the existing `ImportFlowController -> ImportPhotoItemProcessor -> saveImportPhotoItem()` path.
- Keep at most 20 photos and use null-prototype maps for item IDs.
- Render external names with DOM APIs and `textContent`; never expose coordinates, points, XML, File, EXIF, or error bodies.

---

### Task 1: UTC offset and immutable location patch contracts

**Files:**
- Modify: `index.html` (`ImportJobCore`, new `PhotoTrackMatchPreviewCore`)
- Test: `tests/photo-track-match-options-ui.test.js`
- Test: `tests/photo-track-match-apply.test.js`

**Interfaces:**
- Produces: `PhotoTrackMatchPreviewCore.parseUtcOffsetText(value): number|null`
- Produces: `PhotoTrackMatchPreviewCore.formatUtcOffsetMinutes(value): string`
- Produces: `ImportJobCore.applyLocationPatches(job, patches): ImportJob`

- [ ] **Step 1: Write failing parser tests** for strict signed `HH:MM`, ±14:00 boundaries, minute bounds, numeric/whitespace rejection, and formatter round trips.
- [ ] **Step 2: Run `node --test tests/photo-track-match-options-ui.test.js`** and verify failure because `PhotoTrackMatchPreviewCore` does not exist.
- [ ] **Step 3: Implement strict helpers** using `/^([+-])(\d{2}):(\d{2})$/`, rejecting absolute values above 840 minutes and non-integer formatter input.
- [ ] **Step 4: Write failing patch tests** covering identity/order/runtime preservation, input non-mutation, duplicate/unknown/partial/string/out-of-range rejection.
- [ ] **Step 5: Run `node --test tests/photo-track-match-apply.test.js`** and verify failure because `applyLocationPatches` does not exist.
- [ ] **Step 6: Implement minimal immutable patching** by validating all patches before cloning every item with cloned arrays/runtime and changing only `lat` / `lng` for selected IDs.
- [ ] **Step 7: Run both tests** and verify all parser and patch cases pass.

### Task 2: Matching state, ownership, projection, and stale rules

**Files:**
- Modify: `index.html` (`PhotoTrackMatchPreviewCore`)
- Test: `tests/photo-track-match-apply.test.js`

**Interfaces:**
- Produces: `PhotoTrackMatchPreviewCore.create(options)` with `initialize`, `getViewModel`, `setOption`, `run`, `setSelected`, `apply`, `clearResult`, `onDraftChange`, `cleanup`.
- Consumes: `PhotoTrackMatchCore.matchPhotos`, `TrackGeometryCore.normalizeTrack`, `ImportJobCore.applyLocationPatches`.

- [ ] **Step 1: Write failing state tests** for safe initial state, local UTC default, original snapshots, candidate filtering/current revision, and no retained track/points/File data.
- [ ] **Step 2: Write failing ownership tests** for original GPS protection, matcher-owned replacement, manual override ownership loss, cleared-coordinate eligibility, partial GPS rejection, and input non-mutation.
- [ ] **Step 3: Write failing stale tests** for track/revision/options/capturedAt/item ID/order/count/effective-coordinate changes.
- [ ] **Step 4: Run `node --test tests/photo-track-match-apply.test.js`** and verify expected missing behavior failures.
- [ ] **Step 5: Implement controller state transitions** with null-prototype maps, safe projections, result input snapshots, matched-only selection, fixed error-code mapping, and synchronous running guard.
- [ ] **Step 6: Implement explicit apply** with idle/open/current-revision/current-input checks and minimal applied match summaries.
- [ ] **Step 7: Run the test and refactor only after green.**

### Task 3: Dedicated Preview panel and safe rendering

**Files:**
- Modify: `index.html` (CSS, fixed panel DOM, `ImportPreviewUI`)
- Test: `tests/photo-track-match-preview.test.js`

**Interfaces:**
- Consumes: the exact controller returned by `PhotoTrackMatchPreviewCore.create` through `ImportPreviewUI.open({ trackMatch })`.
- Produces: fixed IDs `multi-photo-track-match-panel`, `multi-photo-track-select`, `multi-photo-track-utc-offset`, `multi-photo-track-clock-correction`, `multi-photo-track-max-gap`, `multi-photo-track-endpoint-tolerance`, `multi-photo-track-run`, `multi-photo-track-status`, `multi-photo-track-error`, `multi-photo-track-counts`, `multi-photo-track-warnings`, `multi-photo-track-results`, `multi-photo-track-apply`, `multi-photo-track-clear`.

- [ ] **Step 1: Write failing DOM tests** for multi-photo-only visibility, other-format invisibility, fixed labels/defaults, accessible live/error regions, and disabled explanation.
- [ ] **Step 2: Write failing render tests** for safe track options, no point/coordinate/XML exposure, Japanese counts/warnings/status labels, matched-only default checkboxes, and item order.
- [ ] **Step 3: Run `node --test tests/photo-track-match-preview.test.js`** and verify missing panel/controller wiring failures.
- [ ] **Step 4: Add fixed semantic HTML and minimal responsive CSS** without changing the overall Preview layout.
- [ ] **Step 5: Extend `ImportPreviewUI`** with a dedicated `trackMatch` slot, DOM-only rendering, input/click/change handlers, normal draft notification after apply, focus notification, and cleanup.
- [ ] **Step 6: Run the Preview test** and verify all cases pass without `innerHTML`.

### Task 4: Production multi-photo workflow and save integration

**Files:**
- Modify: `index.html` (`state.multiPhotoImport`, `MultiPhotoImportWorkflow`, production configuration, `ImportFlowController` option forwarding)
- Test: `tests/photo-track-match-import-integration.test.js`
- Update: `tests/multi-photo-import-workflow.test.js`
- Update: `tests/import-preview-ui.test.js`

**Interfaces:**
- Consumes: `state.tracks`, `PhotoTrackMatchPreviewCore`, existing flow and processor.
- Produces: no new server interface; resulting job remains the only input to `ImportPhotoItemProcessor`.

- [ ] **Step 1: Write failing integration tests** for initialization/cleanup, track state injection, draft propagation, flow exclusion while matching, and non-multi Preview invisibility.
- [ ] **Step 2: Write failing save tests** for A existing GPS/B matched/C missing time and assert the existing payload has A original coordinates, B applied coordinates, C null coordinates, with no track/revision/match fields.
- [ ] **Step 3: Run targeted integration tests** and verify failures precede implementation.
- [ ] **Step 4: Add `trackMatch` to `state.multiPhotoImport`** and reset it at selection start, Preview discard/close, completion cleanup, and preparation failure.
- [ ] **Step 5: Forward only the dedicated controller** through `MultiPhotoImportWorkflow -> ImportFlowController.open -> ImportPreviewUI.open`; inject current tracks in production.
- [ ] **Step 6: Preserve existing busy/beforeunload boundaries** and disable match operations unless the job is idle and Preview is open.
- [ ] **Step 7: Run targeted tests** and verify save APIs/routes/tracks remain unchanged.

### Task 5: Documentation, regression, and delivery

**Files:**
- Modify: `README.md`

- [ ] **Step 1: Update README Phase 8-2 behavior and exclusions** including explicit apply, timezone/correction/gap/endpoint, ownership protections, existing save path, non-persistence, unsupported single/Drive inputs, and unexecuted manual checklist.
- [ ] **Step 2: Run `node --check Code.js`.** Expected: exit 0.
- [ ] **Step 3: Run `node --test tests/*.test.js`.** Expected: all tests pass with zero failures.
- [ ] **Step 4: Run `git diff --check`.** Expected: no whitespace errors.
- [ ] **Step 5: Self-review the diff** for unrelated changes, route/API/Sheet changes, secret/error leakage, mutation, coordinate ownership, cleanup, and regressions.
- [ ] **Step 6: Commit only task files** with `feat: match multi photos to track`.
- [ ] **Step 7: Push the current `v1.4.0` branch to `origin/v1.4.0`** without merging, opening a PR, deploying Apps Script, or rewriting history.
