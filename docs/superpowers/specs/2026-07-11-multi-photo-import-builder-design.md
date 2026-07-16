# Multi-photo Import Builder Design

## Goal

Prepare one to twenty selected JPEG, PNG, WebP, HEIC, or HEIF photos as an editable `ImportJobCore` draft without exposing a production upload entry point or invoking GAS, Base64 conversion, the queue runner, or registration.

## Architecture

`MultiPhotoImportBuilder` is a DOM- and GAS-independent session factory placed beside the existing import core. A session snapshots files and defaults before work begins, validates the whole selection limit, and runs at most one or two workers. Photo preparation is dependency-injected for deterministic tests; the browser default reuses the existing EXIF, HEIC metadata, JPEG conversion, and filename normalization helpers.

The builder owns Object URLs only while a run is being assembled. A successful run transfers those URLs to the returned ImportJob runtime and records its matching revoker in a WeakMap. Multi-photo preview edits propagate that ownership to each immutable next job, so delete and discard use the matching URL API and deduplicate repeated URL strings. Cancellation or fatal assembly failure revokes all run-owned URLs once, immediately drops unstarted File references, waits for in-flight preparation to settle, and rejects with a coded cancellation error. Completed sessions are reusable and retain a session-wide ID set.

## Validation and item results

The selection must contain 1 through `ImportJobCore.MAX_ITEMS` entries. Extension and MIME evidence are both inspected case-insensitively. The five allowed families are JPEG/JPG, PNG, WebP, HEIC, and HEIF; SVG, GIF, BMP, TIFF, PDF, video, and unknown files become failed draft items rather than aborting other files. An empty MIME is acceptable when the extension is allowed.

Defaults are copied and validated before workers start: tags are at most five non-empty strings, color is both a safe hex and one of `PIN_COLORS`, icon is a `PIN_ICONS` id, and status is empty or a `PIN_STATUSES` value. Each item receives its own tags array.

Successful standard images retain the same object for `originalFile` and `uploadFile`. GPS and capture date failures produce metadata statuses without failing the upload draft. HEIC/HEIF metadata is read from the original, conversion is invoked exactly once, the first converted frame follows the existing converter behavior, and the normalized JPEG becomes `uploadFile`. Conversion and metadata statuses remain separate.

Unsupported, invalid-reference, conversion, preparation, and Object URL failures remain visible as failed ImportItems with sanitized fixed messages. File contents, parser details, stacks, Object URLs, and File objects never enter progress events or persistable items.

## Concurrency, progress, and cancellation

Workers claim each selection index once and write into a fixed result slot, so completion order cannot change selection order. Every per-index operation, including malformed getters from injected values, is caught inside the worker loop and converted to a sanitized failed item. Progress snapshots report total, pending, processing, ready, failed, and cancelled at start, item start, item completion, cancellation request, and final settlement. Observer exceptions are ignored.

`cancel()` is idempotent. It prevents new claims, releases already prepared results immediately, discards late in-flight results, and lets the original asynchronous library calls settle naturally. A generation-local run record prevents late work from entering a later `start()` call.

## Preview integration

`ImportPreviewUI.open()` gains a backward-compatible `hidePrimaryAction` option. `openMultiPhotoImportPreview(job, options)` supplies the multi-photo title and count label, always hides the primary action, and forwards draft/close observers so callers can retain the current immutable job after editing or deletion. Presets, editing, failed-item deletion, normal close retention, and discard cleanup continue through the existing preview implementation.

No button, file input, URL parameter, or other normal-screen entry point is added.

## Verification

A dedicated Node test file covers format evidence, defaults, normal and HEIC preparation, ordering, concurrency limits, partial failures, progress isolation, cancellation races, resource cleanup, ID uniqueness, persistable output, and preview opening. Existing import, HEIC, preset, flow, queue, and HTML suites remain the regression boundary.
