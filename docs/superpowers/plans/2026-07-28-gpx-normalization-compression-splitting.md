# GPX Normalization, Compression, and Splitting Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** GPXの完全重複を安全に除去し、4時間以上の中断で行程を分け、各行程を最大5秒間隔まで必要最小限に圧縮して、必要なら`名前(1/N)`形式の複数トラックとして再試行可能に保存する。

**Architecture:** `GpxTrackInterchangeCore`へsource解析と保存用batch生成を分離した純粋処理を追加し、既存の単一トラックAPIは維持する。`TrackFileImportUI`は単一draftを1件batchとして扱う形へ拡張し、複数GPX draftを共通Previewで編集して既存`saveTrackBundle`へ直列・冪等に送る。サーバーのトラックpayload、Sheetsスキーマ、1トラック20,000 point制約は変更しない。

**Tech Stack:** JavaScript、ブラウザDOMParser、Leaflet既存トラックモデル、Google Apps Script、Node.js `node:test`／`vm`

## Global Constraints

- GPXファイル上限は5MB、source point安全上限は100,000、1回の生成トラック上限は20件とする。
- 保存される各トラックは最大200 segment、20,000 point、保存済み全体は最大100トラックとする。
- 完全重複は同一元segment内の時刻・緯度・経度・標高が正規化後に一致する時刻付きpointだけとする。
- 4時間以上の時刻差で行程を分け、日付変更だけでは分けない。
- 圧縮は20,000 point超過時だけ実行し、1～5秒のうち必要最小の間隔を選ぶ。
- 最大5秒でも超過する場合はpointを追加削除せず複数トラックへ分ける。
- 元GPX、server payload、tracks／track_segmentsスキーマ、GeoJSON取込、共有形式は変更しない。
- 複数保存は直列実行し、同じtrackId／revisionId／payloadで安全に再試行できるようにする。

---

### Task 1: GPX source normalization and batch draft core

**Files:**
- Modify: `index.html:14730-15240` (`GpxTrackInterchangeCore`)
- Modify: `tests/gpx-track-interchange-core.test.js`

**Interfaces:**
- Consumes: `buildDraftBatch(text, { sourceName, parseXml, generateId? })`
- Produces:
  - `{ drafts, baseName, summary, stats, warnings }`
  - `stats = { sourcePointCount, pointCount, duplicatePointCount, interruptionCount, compressedPointCount, generatedTrackCount, compressionIntervals }`
  - `updateDraftBatch(batch, patch)`
  - `toSavePayloads(batch)`
- Preserves: single-output `parse`, `buildDraft`, `updateDraft`, `toSavePayload`

- [ ] **Step 1: Read the test-quality rules before changing tests**

Run:

```powershell
Get-Content -Raw 'C:\Users\takik\.codex\plugins\cache\openai-curated-remote\superpowers\6.2.0\skills\test-driven-development\writing-good-tests.md'
```

Expected: testが本番動作を直接検証し、実装詳細やmockだけを検証しないルールを確認できる。

- [ ] **Step 2: Add failing duplicate-removal tests**

Add helpers and tests to `tests/gpx-track-interchange-core.test.js`:

```js
function trkpt(lat, lon, time, elevation = 10) {
  return `<trkpt lat="${lat}" lon="${lon}"><ele>${elevation}</ele><time>${time}</time></trkpt>`;
}

test('GPX batch removes exact timed duplicates before applying the saved point limit', () => {
  const point = trkpt(35, 139, '2026-01-01T00:00:00Z');
  const batch = loadCore().buildDraftBatch(
    `<gpx version="1.1"><trk><trkseg>${point.repeat(20001)}</trkseg></trk></gpx>`,
    options()
  );
  assert.equal(batch.stats.sourcePointCount, 20001);
  assert.equal(batch.stats.duplicatePointCount, 20000);
  assert.equal(batch.drafts.length, 1);
  assert.equal(batch.drafts[0].summary.pointCount, 1);
});

test('GPX duplicate removal is exact timed and scoped to its source segment', () => {
  const same = trkpt(35, 139, '2026-01-01T00:00:00Z');
  const noTime = '<trkpt lat="35" lon="139"><ele>10</ele></trkpt>';
  const differentElevation = trkpt(35, 139, '2026-01-01T00:00:00Z', 11);
  const xml = `<gpx version="1.1"><trk>`
    + `<trkseg>${same}${noTime}${same}${noTime}${differentElevation}</trkseg>`
    + `<trkseg>${same}</trkseg></trk></gpx>`;
  const batch = loadCore().buildDraftBatch(xml, options());
  assert.equal(batch.stats.duplicatePointCount, 1);
  assert.deepEqual(plain(batch.drafts[0].segments.map((segment) => segment.points.length)), [4, 1]);
});
```

Name the production change that makes these tests pass: `buildDraftBatch()` must validate every source point, then deduplicate exact timed points before `MAX_POINTS` is enforced.

- [ ] **Step 3: Run duplicate tests and verify RED**

Run:

```powershell
node --test --test-name-pattern "batch removes exact|duplicate removal is exact" tests/gpx-track-interchange-core.test.js
```

Expected: FAIL because `buildDraftBatch` does not exist.

- [ ] **Step 4: Implement validated source collection and exact deduplication**

In `GpxTrackInterchangeCore`:

```js
const MAX_SOURCE_POINTS = 100000;
const MAX_GENERATED_TRACKS = 20;
const INTERRUPTION_MS = 4 * 60 * 60 * 1000;
const COMPRESSION_INTERVAL_SECONDS = [1, 2, 3, 4, 5];

function duplicateKey(point) {
  if (!point.time) return '';
  return [
    point.time, point.lat, point.lng,
    point.elevation === null ? 'null' : point.elevation
  ].join('|');
}

function normalizeSourcePoints(pointElements, sourceStats) {
  if (sourceStats.sourcePointCount + pointElements.length > MAX_SOURCE_POINTS) {
    fail('GPX_SOURCE_POINT_LIMIT_EXCEEDED', 'GPXのsource point数が上限を超えています。');
  }
  const seen = new Set();
  const points = pointElements.map(pointFromElement).filter(function(point) {
    sourceStats.sourcePointCount += 1;
    const key = duplicateKey(point);
    if (!key || !seen.has(key)) {
      if (key) seen.add(key);
      return true;
    }
    sourceStats.duplicatePointCount += 1;
    return false;
  });
  return points;
}
```

Refactor `collectSegments()` so source parsing does not apply `trackGeometry.MAX_POINTS` before deduplication. Keep the existing 200 source segment limit and safe metadata traversal.

- [ ] **Step 5: Run duplicate tests and verify GREEN**

Run:

```powershell
node --test --test-name-pattern "batch removes exact|duplicate removal is exact" tests/gpx-track-interchange-core.test.js
```

Expected: PASS.

- [ ] **Step 6: Add failing interruption, compression, and overflow-splitting tests**

Add:

```js
test('GPX batch splits a four-hour interruption but not a midnight crossing', () => {
  const continuous = [
    trkpt(35, 139, '2026-01-01T23:59:59Z'),
    trkpt(35.001, 139.001, '2026-01-02T00:00:01Z')
  ].join('');
  assert.equal(loadCore().buildDraftBatch(
    `<gpx version="1.1"><trk><trkseg>${continuous}</trkseg></trk></gpx>`, options()
  ).drafts.length, 1);

  const interrupted = [
    trkpt(35, 139, '2026-01-01T18:00:00Z'),
    trkpt(35.001, 139.001, '2026-01-01T21:59:59Z'),
    trkpt(35.002, 139.002, '2026-01-02T02:00:00Z')
  ].join('');
  const batch = loadCore().buildDraftBatch(
    `<gpx version="1.1"><trk><trkseg>${interrupted}</trkseg></trk></gpx>`,
    options()
  );
  assert.equal(batch.stats.interruptionCount, 1);
  assert.deepEqual(plain(batch.drafts.map((draft) => draft.name)), ['walk(1/2)', 'walk(2/2)']);
  assert.deepEqual(plain(batch.drafts.map((draft) => draft.summary.pointCount)), [2, 1]);
});

test('GPX batch chooses the smallest whole-second interval that fits 20000 points', () => {
  const start = Date.parse('2026-01-01T00:00:00Z');
  const points = Array.from({ length: 20001 }, (_, index) =>
    trkpt(35 + index / 1000000, 139, new Date(start + index * 1000).toISOString())
  ).join('');
  const batch = loadCore().buildDraftBatch(
    `<gpx version="1.1"><trk><trkseg>${points}</trkseg></trk></gpx>`,
    options()
  );
  assert.deepEqual(plain(batch.stats.compressionIntervals), [2]);
  assert.equal(batch.drafts.length, 1);
  assert.equal(batch.drafts[0].summary.pointCount, 10001);
  assert.equal(batch.drafts[0].segments[0].points[0].time, '2026-01-01T00:00:00.000Z');
  assert.equal(batch.drafts[0].segments[0].points.at(-1).time,
    new Date(start + 20000 * 1000).toISOString());
});

test('GPX batch partitions a stage that still exceeds 20000 points without losing sampled points', () => {
  const start = Date.parse('2026-01-01T00:00:00Z');
  const points = Array.from({ length: 60001 }, (_, index) =>
    `<trkpt lat="35" lon="139"><time>${
      new Date(start + index * 2000).toISOString()
    }</time></trkpt>`
  ).join('');
  const batch = loadCore().buildDraftBatch(
    `<gpx version="1.1"><trk><trkseg>${points}</trkseg></trk></gpx>`,
    options()
  );
  assert.equal(batch.drafts.length, 2);
  assert.deepEqual(plain(batch.drafts.map((draft) => draft.name)), ['walk(1/2)', 'walk(2/2)']);
  assert.deepEqual(plain(batch.drafts.map((draft) => draft.summary.pointCount)), [20000, 1]);
  assert.equal(batch.stats.compressedPointCount, 40000);
  const savedTimes = batch.drafts.flatMap((draft) =>
    draft.segments.flatMap((segment) => segment.points.map((point) => point.time)));
  assert.equal(savedTimes.length, 20001);
  assert.equal(new Set(savedTimes).size, 20001);
});
```

The last case uses two-second source points over more than 27 hours. Five-second sampling still leaves 20,001 points, so it proves that the sampled points are partitioned in order without another deletion pass.

- [ ] **Step 7: Run new transformation tests and verify RED**

Run:

```powershell
node --test --test-name-pattern "four-hour interruption|smallest whole-second|partitions a stage" tests/gpx-track-interchange-core.test.js
```

Expected: FAIL because interruption splitting, interval selection, partitioning, and suffix naming are missing.

- [ ] **Step 8: Implement stage splitting, compression, partitioning, and batch draft APIs**

Add these pure helpers:

```js
function segmentPointCount(segments) {
  return segments.reduce(function(total, segment) {
    return total + segment.points.length;
  }, 0);
}

function splitSegmentsAtInterruptions(segments, stats) {
  const stages = [];
  let stage = [];
  function flushStage() {
    if (stage.length) stages.push(stage);
    stage = [];
  }
  segments.forEach(function(segment) {
    let points = [];
    let previousTimedMs = null;
    segment.points.forEach(function(point) {
      const currentMs = point.time ? Date.parse(point.time) : NaN;
      if (points.length && previousTimedMs !== null && Number.isFinite(currentMs)
          && currentMs - previousTimedMs >= INTERRUPTION_MS) {
        stage.push({ index: stage.length, points: points });
        flushStage();
        points = [];
        stats.interruptionCount += 1;
      }
      points.push(point);
      previousTimedMs = Number.isFinite(currentMs) ? currentMs : null;
    });
    if (points.length) stage.push({ index: stage.length, points: points });
  });
  flushStage();
  return stages;
}

function sampleSegmentBySeconds(segment, intervalSeconds) {
  if (segment.points.length <= 2) {
    return { index: segment.index, points: segment.points.slice() };
  }
  const result = [segment.points[0]];
  let lastTimedMs = segment.points[0].time ? Date.parse(segment.points[0].time) : null;
  segment.points.slice(1, -1).forEach(function(point) {
    const currentMs = point.time ? Date.parse(point.time) : NaN;
    if (!Number.isFinite(currentMs) || lastTimedMs === null
        || currentMs - lastTimedMs >= intervalSeconds * 1000) {
      result.push(point);
      lastTimedMs = Number.isFinite(currentMs) ? currentMs : null;
    }
  });
  const last = segment.points[segment.points.length - 1];
  if (result[result.length - 1] !== last) result.push(last);
  return { index: segment.index, points: result };
}

function partitionStage(stageSegments) {
  const parts = [];
  let current = [];
  let currentCount = 0;
  function flush() {
    if (current.length) parts.push(current);
    current = [];
    currentCount = 0;
  }
  stageSegments.forEach(function(segment) {
    let offset = 0;
    while (offset < segment.points.length) {
      if (currentCount === trackGeometry.MAX_POINTS) flush();
      const take = Math.min(
        trackGeometry.MAX_POINTS - currentCount,
        segment.points.length - offset
      );
      current.push({
        index: current.length,
        points: segment.points.slice(offset, offset + take)
      });
      currentCount += take;
      offset += take;
    }
  });
  flush();
  return parts;
}

function partName(baseName, index, total) {
  if (total === 1) return baseName;
  const suffix = '(' + (index + 1) + '/' + total + ')';
  return baseName.slice(0, Math.max(0, 100 - suffix.length)) + suffix;
}
```

Implementation rules:

- Treat each original segment independently for dedup and interval sampling.
- Create a new stage only for adjacent timed points in the same source segment whose difference is `>= INTERRUPTION_MS`.
- Do not split merely because UTC/local dates differ.
- Evaluate 1, 2, 3, 4, then 5 seconds and choose the first candidate at or below 20,000 points.
- `compressStage(stageSegments, stats)` returns the original stage when it is already at or below 20,000 points. Otherwise it evaluates cloned candidates from the original stage for each interval, stores only the first fitting candidate, and records the chosen interval and `originalCount - retainedCount`.
- After the five-second candidate, retain every timed point in that candidate. If excess points are untimed, preserve every segment endpoint and choose the remaining untimed indexes deterministically with `Math.floor(slot * candidateCount / requiredCount)` until the stage reaches 20,000. If the timed and endpoint set alone exceeds 20,000, do not delete another point.
- If a stage still exceeds 20,000, partition at existing segment boundaries first, then split an oversized segment into ordered slices.
- Generate IDs only after every source point and every transformation has succeeded.
- Aggregate summary from the produced drafts; never reuse pre-compression distance or bounds.
- `buildDraftBatch()` calls source parsing, deduplication, interruption splitting, compression, and partitioning in that order; rejects more than 20 final parts; then generates two IDs per part and normalizes each draft with `trackGeometry.normalizeTrack()`.
- `updateDraftBatch()` validates a common patch through the existing `updateDraft()`, applies the common description/color/visibility/style to every draft, and regenerates all names through `partName()`.
- `toSavePayloads()` maps the batch drafts through the existing `toSavePayload()` whitelist and returns a deep-independent array.

- [ ] **Step 9: Run the complete GPX core test file**

Run:

```powershell
node --test tests/gpx-track-interchange-core.test.js
```

Expected: PASS, including existing GPX 1.0／1.1, XML safety, metadata, and single-draft tests.

- [ ] **Step 10: Commit the core**

Run:

```powershell
git add index.html tests/gpx-track-interchange-core.test.js
git commit -m "feat: normalize and split large GPX tracks"
```

Expected: one commit containing only core behavior and core tests.

### Task 2: Batch preview and common editing

**Files:**
- Modify: `index.html:3520-3555` (track import preview markup)
- Modify: `index.html:15245-15720` (adapter and preview workflow)
- Modify: `index.html:15970-16020` (`state.trackImport`)
- Modify: `tests/gpx-track-import-workflow.test.js`

**Interfaces:**
- Consumes: Task 1 `buildDraftBatch`, `updateDraftBatch`, `toSavePayloads`
- Produces: `importState.batch`, preview list `#track-import-preview-parts`, common form editing across all drafts
- Preserves: GeoJSON as a one-draft batch and the existing single-track preview

- [ ] **Step 1: Add failing preview tests**

Extend the test document and workflow state to include `batch`, `submittedPayloads`, and `savedTrackIds`, then add:

```js
test('GPX preview renders generated part names and aggregate transformation counts', async () => {
  const setup = createWorkflow();
  const xml = '<gpx version="1.1"><trk><trkseg>'
    + trkpt(35, 139, '2026-01-01T18:00:00Z')
    + trkpt(35.1, 139.1, '2026-01-02T02:00:00Z')
    + '</trkseg></trk></gpx>';
  const batch = await setup.gpx.importFile(gpxFile({ text: async () => xml }));
  assert.equal(batch.drafts.length, 2);
  assert.match(setup.documentApi.getElementById('track-import-preview-parts').textContent,
    /walk\\(1\\/2\\).*walk\\(2\\/2\\)/s);
  assert.match(setup.documentApi.getElementById('track-import-preview-stats').textContent,
    /記録中断 1件.*生成ルート 2件/s);
});

test('editing a GPX batch applies common metadata and regenerates every suffix', async () => {
  const setup = createWorkflow();
  await setup.gpx.importFile(gpxFile({ text: async () => twoStageGpx() }));
  setup.documentApi.getElementById('track-import-preview-name').value = '縦走';
  setup.documentApi.getElementById('track-import-preview-description').value = '2泊3日';
  setup.documentApi.getElementById('track-import-preview-color').value = '#2196f3';
  const batch = setup.gpx.syncDraftFromForm();
  assert.deepEqual(plain(batch.drafts.map((draft) => draft.name)), ['縦走(1/2)', '縦走(2/2)']);
  assert.ok(batch.drafts.every((draft) =>
    draft.description === '2泊3日' && draft.color === '#2196f3'));
});
```

- [ ] **Step 2: Run preview tests and verify RED**

Run:

```powershell
node --test --test-name-pattern "preview renders generated|editing a GPX batch" tests/gpx-track-import-workflow.test.js
```

Expected: FAIL because the workflow stores one `draft` and has no parts preview.

- [ ] **Step 3: Implement adapter batch normalization**

Extend `createTrackFileImportAdapter()`:

```js
buildBatch: function(contents, options) {
  if (typeof adapterConfig.core.buildDraftBatch === 'function') {
    return adapterConfig.core.buildDraftBatch(contents, options);
  }
  const draft = adapterConfig.core.buildDraft(contents, options);
  return {
    drafts: [draft], baseName: draft.name, summary: draft.summary,
    stats: null, warnings: draft.warnings || []
  };
},
updateBatch: function(batch, patch) {
  return typeof adapterConfig.core.updateDraftBatch === 'function'
    ? adapterConfig.core.updateDraftBatch(batch, patch)
    : { ...batch, drafts: [adapterConfig.core.updateDraft(batch.drafts[0], patch)] };
},
toSavePayloads: function(batch) {
  return typeof adapterConfig.core.toSavePayloads === 'function'
    ? adapterConfig.core.toSavePayloads(batch)
    : [adapterConfig.core.toSavePayload(batch.drafts[0])];
}
```

Keep the existing adapter methods as compatibility aliases.

- [ ] **Step 4: Implement batch state and preview list**

Add `#track-import-preview-parts` after the legacy summary using the existing `section-card` class. In `TrackFileImportUI`:

- set `importState.batch` from `adapter.buildBatch()`;
- keep `importState.draft = batch.drafts[0]` as a compatibility alias;
- use aggregate summary for the top summary;
- render each generated part as text containing name, time range, segment count, and point count;
- hide the part list for one draft;
- pass batch stats to the GPX stats renderer;
- have `syncDraftFromForm()` call `adapter.updateBatch()` and refresh both `batch` and `draft`.

When the saved-track edit overlay reuses this DOM, clear and hide `track-import-preview-parts`.

- [ ] **Step 5: Run focused preview tests**

Run:

```powershell
node --test --test-name-pattern "preview renders generated|editing a GPX batch|preview" tests/gpx-track-import-workflow.test.js
```

Expected: PASS.

- [ ] **Step 6: Run all GPX UI workflow tests**

Run:

```powershell
node --test tests/gpx-track-import-workflow.test.js
```

Expected: PASS; GeoJSON continues to behave as a one-draft batch.

- [ ] **Step 7: Commit preview support**

Run:

```powershell
git add index.html tests/gpx-track-import-workflow.test.js
git commit -m "feat: preview split GPX imports"
```

Expected: one commit containing preview/state changes and their tests.

### Task 3: Sequential idempotent multi-track saving

**Files:**
- Modify: `index.html:15720-15910` (`TrackFileImportUI` submit/save/retry)
- Modify: `index.html:16000-16020` (`state.trackImport`)
- Modify: `index.html:16318-16340` (production workflow options)
- Modify: `tests/gpx-track-import-workflow.test.js`

**Interfaces:**
- Consumes: Task 2 `importState.batch` and `adapter.toSavePayloads(batch)`
- Produces: stable `submittedPayloads`, `savedTrackIds`, sequential `saveTrackBundle` calls, retry progress
- Preserves: single GeoJSON/GPX save response validation and `onSaved(track)`

- [ ] **Step 1: Add failing sequential-save and retry tests**

Add:

```js
test('split GPX payloads save sequentially with stable identities', async () => {
  const gates = [deferred(), deferred()];
  let callIndex = 0;
  const setup = createWorkflow({
    callGAS: (method, payload) => {
      setup.calls.push({ method, payload: plain(payload) });
      return gates[callIndex++].promise;
    }
  });
  await setup.gpx.importFile(gpxFile({ text: async () => twoStageGpx() }));
  const saving = setup.gpx.save();
  await Promise.resolve();
  assert.equal(setup.calls.length, 1);
  gates[0].resolve({ ok: true, track: { ...setup.calls[0].payload, orderIndex: 0 } });
  await Promise.resolve();
  await Promise.resolve();
  assert.equal(setup.calls.length, 2);
  gates[1].resolve({ ok: true, track: { ...setup.calls[1].payload, orderIndex: 1 } });
  assert.equal(await saving, true);
  assert.equal(setup.savedTracks.length, 2);
  assert.notEqual(setup.calls[0].payload.trackId, setup.calls[1].payload.trackId);
});

test('split GPX retry reuses every payload and does not duplicate applied tracks', async () => {
  let attempt = 0;
  const setup = createWorkflow({
    callGAS: async (_method, payload) => {
      setup.calls.push(plain(payload));
      attempt += 1;
      if (attempt === 2) return { ok: false, errorCode: 'TRACK_STORAGE_BUSY', retryable: true };
      return { ok: true, deduplicated: attempt > 2, track: { ...plain(payload), orderIndex: attempt - 1 } };
    }
  });
  await setup.gpx.importFile(gpxFile({ text: async () => twoStageGpx() }));
  assert.equal(await setup.gpx.save(), false);
  const firstAttemptPayloads = plain(setup.state.submittedPayloads);
  assert.equal(setup.savedTracks.length, 1);
  assert.equal(await setup.gpx.retry(), true);
  assert.deepEqual(plain(setup.state.submittedPayloads), firstAttemptPayloads);
  assert.equal(setup.savedTracks.length, 2);
});
```

- [ ] **Step 2: Run save tests and verify RED**

Run:

```powershell
node --test --test-name-pattern "payloads save sequentially|retry reuses every payload" tests/gpx-track-import-workflow.test.js
```

Expected: FAIL because only one submitted payload is supported.

- [ ] **Step 3: Implement stable sequential submission**

Refactor save state:

```js
submittedPayloads: null,
savedTrackIds: Object.create(null),
savedCount: 0
```

On first `save(kind)`:

```js
importState.submittedPayloads = clonePayload(adapter.toSavePayloads(importState.batch));
```

In `submit(kind)`, reduce the frozen payloads into a Promise chain. For every item:

1. Call `saveTrackBundle` with the same payload and edit token.
2. Reuse the existing strict response projection and submitted-payload comparison.
3. Call `onSaved` only when `savedTrackIds[trackId]` is not set.
4. Mark the ID and update `savedCount`.
5. Stop at the first failure and leave the payload array intact.

On retry, start from the same payload array. Already-saved revisions may be sent again and must be accepted through `deduplicated: true`, but must not call `onSaved` twice.

Render `保存中 X/N` while saving and `X/N件を保存しました。再試行してください。` on partial failure.

- [ ] **Step 4: Add and implement the client track-capacity preflight**

Pass:

```js
getTrackCount: function() { return state.tracks.length; }
```

to `TrackFileImportUI.create()`. Before freezing payloads, reject when:

```js
getTrackCount() + importState.batch.drafts.length > 100
```

with `TRACK_LIMIT_EXCEEDED`. The server remains authoritative and still checks each payload under lock.

- [ ] **Step 5: Run focused save tests**

Run:

```powershell
node --test --test-name-pattern "payloads save sequentially|retry reuses every payload|TRACK_LIMIT_EXCEEDED" tests/gpx-track-import-workflow.test.js
```

Expected: PASS.

- [ ] **Step 6: Run core, workflow, and server storage regression tests**

Run:

```powershell
node --test tests/gpx-track-interchange-core.test.js tests/gpx-track-import-workflow.test.js tests/track-storage.test.js
```

Expected: PASS. `Code.js` continues rejecting any individual payload above 20,000 points.

- [ ] **Step 7: Commit sequential saving**

Run:

```powershell
git add index.html tests/gpx-track-import-workflow.test.js
git commit -m "feat: save split GPX tracks sequentially"
```

Expected: one commit containing save/retry/capacity behavior and tests.

### Task 4: Documentation, real-file regression, and final verification

**Files:**
- Modify: `README.md:53`
- Modify: `tests/gpx-track-import-ui.test.js`
- Modify: `docs/superpowers/plans/2026-07-28-gpx-normalization-compression-splitting.md` (checkbox completion)

**Interfaces:**
- Documents: 5MB、100,000 source point、完全重複除去、最大5秒圧縮、4時間中断分割、各20,000 point／最大20生成トラック
- Verifies: `yamap_2025-11-23_09_07.gpx` becomes one 17,724-point draft without modifying the file

- [ ] **Step 1: Add a failing README contract test**

Update `tests/gpx-track-import-ui.test.js`:

```js
test('README documents adaptive GPX normalization limits', () => {
  assert.match(readme,
    /GPXトラック[^\n]*5MB[^\n]*100,000 source point[^\n]*重複除去[^\n]*最大5秒[^\n]*4時間[^\n]*20,000 point/);
});
```

- [ ] **Step 2: Run the documentation test and verify RED**

Run:

```powershell
node --test --test-name-pattern "adaptive GPX normalization" tests/gpx-track-import-ui.test.js
```

Expected: FAIL because README still says only 20,000 points.

- [ ] **Step 3: Update README**

Replace the GPX capability row with concise wording that includes:

```markdown
GPX 1.0 / 1.1、trk／rte、5MB・100,000 source pointまで。完全重複を除去し、必要時は最大5秒間隔に圧縮、4時間以上の中断で分割。保存は1ルート20,000 point、1回最大20ルート。
```

Do not change unrelated setup or Drive-account documentation.

- [ ] **Step 4: Run the documentation test and verify GREEN**

Run:

```powershell
node --test tests/gpx-track-import-ui.test.js
```

Expected: PASS.

- [ ] **Step 5: Run the actual YAMAP file through the production core**

Run a read-only Node/PowerShell harness against:

```text
C:\Users\takik\Downloads\yamap_2025-11-23_09_07.gpx
```

Expected:

- sourcePointCount: 35,582
- duplicatePointCount: 17,858
- drafts: 1
- saved pointCount: 17,724
- compression interval: none
- original file hash and byte length unchanged

- [ ] **Step 6: Run required full verification**

Run:

```powershell
node --check Code.js
node --test tests/*.test.js
git diff --check
git status --short
git diff --stat HEAD~3
git diff HEAD~3
```

Expected: all tests pass; no XML contents, personal coordinates, unrelated Drive/photo changes, server schema changes, or generated artifacts are committed.

- [ ] **Step 7: Review safety and compatibility**

Confirm:

- every saved payload is at or below 20,000 point and 200 segment;
- invalid duplicate candidates still fail validation;
- midnight crossings without a 4-hour gap remain one route;
- split saves are sequential and retry with stable IDs;
- partial success never duplicates client state;
- GeoJSON remains a one-track import;
- shared.html and storage schemas are unchanged;
- source GPX is never written.

- [ ] **Step 8: Commit documentation and push**

Run:

```powershell
git add README.md tests/gpx-track-import-ui.test.js docs/superpowers/plans/2026-07-28-gpx-normalization-compression-splitting.md
git commit -m "docs: describe adaptive GPX imports"
git push origin fix/inherited-drive-sharing
```

Expected: the current branch is pushed without force and contains the design, plan, implementation, tests, and README update.
