# Multi-photo Import Builder Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build reusable multi-photo draft preparation and preview integration without connecting it to registration or the normal UI.

**Architecture:** Add a dependency-injected preparation session around `ImportJobCore`, reuse existing image metadata and HEIC helpers in its browser defaults, and extend the existing preview with one opt-in hidden-primary-action flag. Fixed result slots preserve selection order while two worker loops bound concurrency.

**Tech Stack:** Google Apps Script HTML, browser File/Blob/Object URL APIs, existing ExifReader/heic-to/exif-js helpers, Node.js built-in test runner and VM harnesses.

## Global Constraints

- Accept only JPEG/JPG, PNG, WebP, HEIC, and HEIF evidence; always reject SVG.
- Require 1 through `ImportJobCore.MAX_ITEMS` files and reject an oversized selection before work.
- Default and maximum concurrency is 2; accepted values are 1 and 2.
- Keep original File, upload File, and Object URL only under `runtime`.
- Do not call GAS, Base64 conversion, `saveMapData()`, `ImportQueueRunner.start()`, or registration.
- Do not expose any new normal-screen launch control.

---

### Task 1: Builder validation and ordered preparation

**Files:**
- Modify: `index.html`
- Create: `tests/multi-photo-import-builder.test.js`

**Interfaces:**
- Produces: `MultiPhotoImportBuilder.create(options)` returning `{ start(files, defaults), cancel(), isRunning(), release() }`.
- Produces coded errors for invalid configuration, selection, defaults, running state, and cancellation.

- [ ] Add failing tests for accepted/rejected MIME and suffix evidence, empty and 21-item rejection, default snapshots, IDs, titles, runtime-only resources, and standard-photo metadata outcomes.
- [ ] Run `node --test tests/multi-photo-import-builder.test.js` and verify failure because the builder is absent.
- [ ] Implement validation, safe IDs/errors/titles, fixed result slots, ImportItem construction, and `ImportJobCore.createJob()` integration.
- [ ] Rerun the focused tests until the initial behavior passes.

### Task 2: HEIC preparation, concurrency, progress, and cancellation

**Files:**
- Modify: `index.html`
- Modify: `tests/multi-photo-import-builder.test.js`

**Interfaces:**
- Browser default preparation returns `{ originalFile, uploadFile, metadataStatus, conversionStatus, lat, lng, capturedAt, error }`.
- Progress emits `{ total, pending, processing, ready, failed, cancelled, eventType, index, filename }` without File data.

- [ ] Add failing delayed-Promise tests for concurrency 1/2, slot refill, reverse completion order, one claim per index, partial failure, HEIC metadata source/conversion-once/JPEG URL routing, observer exceptions, and every cancellation race.
- [ ] Run the focused test and verify the new assertions fail for missing behavior.
- [ ] Implement bounded workers, existing-helper-backed browser preparation, progress isolation, generation-local cancellation, one-time URL release, and coded cancellation rejection.
- [ ] Rerun the focused tests and refactor only while green.

### Task 3: Preview integration

**Files:**
- Modify: `index.html`
- Modify: `tests/import-preview-ui.test.js`
- Modify: `tests/multi-photo-import-builder.test.js`

**Interfaces:**
- `ImportPreviewUI.open({ hidePrimaryAction })` hides the primary action only when true.
- `openMultiPhotoImportPreview(job, options)` opens with `複数写真を確認`, `写真 N件`, no primary action, and forwarded draft/close callbacks.

- [ ] Add failing UI tests for hidden multi-photo primary action, unchanged normal preview action, safe item/error rendering, editing/deletion, normal-close retention, and discard cleanup.
- [ ] Run the two focused test files and confirm expected failures.
- [ ] Add the opt-in UI state and public multi-photo opener without adding any visible launch element.
- [ ] Rerun the focused tests.

### Task 4: Full regression and delivery

**Files:**
- Review: `index.html`, `tests/multi-photo-import-builder.test.js`, `tests/import-preview-ui.test.js`, design and plan documents.

- [x] Run `node --check Code.js`.
- [x] Run `node --test tests/*.test.js` and confirm all existing HEIC/import/preset/flow/queue/status behavior passes.
- [x] Run `git diff --check` and inspect the diff for UI exposure, leaked runtime data, cancellation races, and unrelated edits.
- [ ] Commit only task files as `feat: prepare multi-photo import drafts` and push `v1.4.0` to `origin`.
