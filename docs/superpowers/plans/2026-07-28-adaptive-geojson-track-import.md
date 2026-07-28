# Adaptive GeoJSON Track Import Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** GeoJSONルートを最大100,000 source pointまで検証し、完全重複除去、4時間中断分割、最大5秒の時刻圧縮、時刻なしpointの形状圧縮、20,000 point単位の追加分割を経て安全に複数保存できるようにする。

**Architecture:** GPX内にあるformat非依存のbatch変換を`TrackBatchTransformCore`へ抽出し、GPXの既存結果を維持したままGeoJSONからも利用する。GeoJSON固有の時刻なし圧縮は`TrackShapeSimplificationCore`へ分離し、Visvalingam–Whyatt方式でsegment端・時刻付きpoint・標高極値・急カーブを優先して残す。既存`TrackFileImportUI`のbatch Preview、共通編集、直列保存、固定payload再試行をそのまま利用し、サーバーpayloadとSheets形式は変更しない。

**Tech Stack:** JavaScript、GeoJSON FeatureCollection、Leaflet既存トラックモデル、Google Apps Script、Node.js `node:test`／`vm`

## Global Constraints

- GeoJSONファイル上限は5MB、source point安全上限は100,000 pointとする。
- Feature上限20件、source segment上限200件を維持する。
- 各保存トラックは200 segment・20,000 point、1回の生成は20トラック、保存済み全体は100トラックまでとする。
- 完全重複は同一元segment内で時刻・緯度・経度・標高がすべて一致する時刻付きpointだけとする。
- 4時間以上の連続時刻差で行程を分け、日付変更だけでは分けない。
- 時刻圧縮は20,000 point超過時だけ実行し、1～5秒の必要最小間隔を選ぶ。
- 形状圧縮ではsegment端、時刻付きpoint、行程の標高最高点・最低点を削除しない。
- 元GeoJSON、server payload、tracks／track_segments、共有形式を変更しない。
- 既存GPXの時刻なし均等抽出を形状圧縮へ変更しない。
- 新規・変更テストは本番Coreの出力を検証し、実装文字列やmockの存在だけを検証しない。

---

### Task 1: Format非依存batch変換Coreの抽出

**Files:**
- Create: `tests/track-batch-transform-core.test.js`
- Modify: `index.html:12890-15580`
- Test: `tests/gpx-track-interchange-core.test.js`

**Interfaces:**
- Consumes: `segments: Array<{index:number, points:Array<{lat,lng,elevation,time}>}>`
- Produces: `TrackBatchTransformCore.transform(segments, options)`
- `options = { maxPoints, interruptionMs, compressionIntervals, reduceOverflow? }`
- Returns: `{ parts, interruptionCount, timeCompressedPointCount, overflowCompressedPointCount, compressedPointCount, compressionIntervals }`
- Produces: `TrackBatchTransformCore.partName(baseName, index, total)`
- Produces: `TrackBatchTransformCore.aggregateSummaries(drafts)`
- Preserves: `GpxTrackInterchangeCore.buildDraftBatch()`の既存出力とstats

- [ ] **Step 1: 共通変換Coreが存在しないことを示す失敗テストを書く**

`tests/track-batch-transform-core.test.js`を追加し、`TrackGeometryCore`から`GeoJsonTrackInterchangeCore`直前までを`vm`で評価する。次の実データテストを追加する。

```js
test('shared transform chooses the smallest time interval and preserves segment endpoints', () => {
  const core = loadCore();
  const start = Date.parse('2026-01-01T00:00:00Z');
  const points = Array.from({ length: 7 }, (_, index) => ({
    lat: 35 + index / 1000,
    lng: 139,
    elevation: null,
    time: new Date(start + index * 1000).toISOString()
  }));
  const result = core.transform([{ index: 0, points }], {
    maxPoints: 4,
    interruptionMs: 4 * 60 * 60 * 1000,
    compressionIntervals: [1, 2, 3, 4, 5]
  });
  assert.deepEqual(plain(result.parts[0][0].points.map((point) => point.time)), [
    '2026-01-01T00:00:00.000Z',
    '2026-01-01T00:00:02.000Z',
    '2026-01-01T00:00:04.000Z',
    '2026-01-01T00:00:06.000Z'
  ]);
  assert.deepEqual(plain(result.compressionIntervals), [2]);
  assert.equal(result.timeCompressedPointCount, 3);
});
```

追加で次を別テストにする。

- 4時間未満と日付変更では1行程、4時間以上では2行程になる。
- 5秒圧縮後も上限超過ならpoint順を変えず複数partへ分ける。
- `reduceOverflow`が返したpoint数だけ`overflowCompressedPointCount`へ計上する。
- `partName('x'.repeat(100), 0, 2)`が100文字以内で`(1/2)`を残す。
- `aggregateSummaries`が距離、point、segment、標高、時刻、boundsを手計算値どおり集計する。

このテストを通す本番変更は、GPX内のローカル変換関数ではなく、独立した`TrackBatchTransformCore`の追加である。

- [ ] **Step 2: 新しいCoreテストを実行してREDを確認する**

Run:

```powershell
node --test tests/track-batch-transform-core.test.js
```

Expected: `TrackBatchTransformCore`が未定義のためFAILする。

- [ ] **Step 3: `TrackBatchTransformCore`を最小実装する**

`TrackGeometryCore`の直後へ次の境界を追加する。

```js
const TrackBatchTransformCore = (function() {
  function cloneSegments(segments) {
    return segments.map(function(segment, index) {
      return { index: index, points: segment.points.slice() };
    });
  }

  function countPoints(segments) {
    return segments.reduce(function(total, segment) {
      return total + segment.points.length;
    }, 0);
  }

  function transform(segments, options) {
    // 4時間中断分割 → 1～5秒の最小間隔圧縮
    // → reduceOverflow callback → maxPoints単位のpart分割を行う。
    // 入力point objectは変更せず、segment indexだけ出力単位で振り直す。
  }

  return Object.freeze({
    transform: transform,
    partName: partName,
    aggregateSummaries: aggregateSummaries,
    cloneSegments: cloneSegments,
    countPoints: countPoints
  });
})();
```

`transform`は次を固定する。

- `maxPoints`は正の整数、`interruptionMs`は非負有限数、`compressionIntervals`は昇順の正整数だけを受ける。
- segment内の4時間境界でstageを分ける。
- 各stageが上限超過時だけ時刻圧縮を試す。
- 最小候補で上限以下になれば、それ以降の候補を試さない。
- 5秒でも超過する場合だけ`reduceOverflow(clonedSegments, maxPoints)`を呼ぶ。
- callbackが不正なsegmentを返した場合は元stageを使い、後段のpart分割へ進む。
- `compressedPointCount = timeCompressedPointCount + overflowCompressedPointCount`とする。

- [ ] **Step 4: Coreテストを実行してGREENを確認する**

Run:

```powershell
node --test tests/track-batch-transform-core.test.js
```

Expected: 全テストPASS。

- [ ] **Step 5: GPXを共通Coreへ移行する前に既存回帰を実行する**

Run:

```powershell
node --test tests/gpx-track-interchange-core.test.js tests/gpx-track-import-workflow.test.js
```

Expected: PASS。移行前の基準結果として保持する。

- [ ] **Step 6: GPX batchを共通Coreへ接続する**

`GpxTrackInterchangeCore.buildDraftBatch`内の変換を次の形へ置き換える。

```js
const transformed = TrackBatchTransformCore.transform(
  source.collected.segments,
  {
    maxPoints: trackGeometry.MAX_POINTS,
    interruptionMs: INTERRUPTION_MS,
    compressionIntervals: COMPRESSION_INTERVAL_SECONDS,
    reduceOverflow: function(segments) {
      return reduceUntimedPoints(segments);
    }
  }
);
const parts = transformed.parts;
transformStats.interruptionCount = transformed.interruptionCount;
transformStats.compressedPointCount = transformed.compressedPointCount;
transformStats.compressionIntervals = transformed.compressionIntervals.slice();
```

GPX固有のXML解析、完全重複、warning、source上限、均等抽出、draft生成は`GpxTrackInterchangeCore`に残す。移行後に不要になった`splitSegmentsAtInterruptions`、`sampleSegmentBySeconds`、`compressStage`、`partitionStage`、`partName`、`aggregateSummaries`のローカル重複だけを削除する。

- [ ] **Step 7: 共通CoreとGPX回帰を実行してGREENを確認する**

Run:

```powershell
node --test tests/track-batch-transform-core.test.js tests/gpx-track-interchange-core.test.js tests/gpx-track-import-workflow.test.js
```

Expected: 全テストPASS。既存GPXのdraft数、保存point、suffix、statsが変わらない。

- [ ] **Step 8: Task 1をコミットする**

```powershell
git add index.html tests/track-batch-transform-core.test.js
git commit -m "refactor: share track import batch transforms"
```

---

### Task 2: 時刻なしGeoJSONの形状簡略化Core

**Files:**
- Create: `tests/track-shape-simplification-core.test.js`
- Modify: `index.html`（`TrackBatchTransformCore`直後）

**Interfaces:**
- Consumes: 正規化済みの`segments`と正整数`maxPoints`
- Produces: `TrackShapeSimplificationCore.reduce(segments, maxPoints)`
- Returns: 新しいsegment／points配列。point object自体は共有してよいが、入力配列とindexは変更しない。
- Protected: segment先頭・末尾、全時刻付きpoint、全行程の標高最高点・最低点

- [ ] **Step 1: 形状簡略化の失敗テストを書く**

`tests/track-shape-simplification-core.test.js`で本番Coreを評価し、次を追加する。

```js
test('shape reduction removes collinear density before sharp corners', () => {
  const result = loadCore().reduce([{
    index: 0,
    points: [
      point(0, 0), point(0, 0.25), point(0, 0.5),
      point(1, 0.5), point(1, 0.75), point(1, 1)
    ]
  }], 4);
  assert.deepEqual(plain(result[0].points.map((value) => [value.lat, value.lng])), [
    [0, 0], [0, 0.5], [1, 0.5], [1, 1]
  ]);
});
```

別テストで次を検証する。

- segment端を必ず残す。
- 直線上でも時刻付きpointを残す。
- 標高最高点と最低点を残す。
- 複数segmentの元順とindexを維持する。
- protected pointが上限を超える場合は追加削除せず全pointを返す。
- 同面積の入力で複数回実行して同じpoint列になる。
- 経度179.9度から-179.9度をまたぐ折れ線で、日付変更線上の小移動を世界一周相当として扱わない。
- 入力segment／points配列を変更せず、削除可能なpointがあれば正確に指定件数を返す。
- `maxPoints`が0、負数、小数、文字列の場合は安全に入力cloneを返す。

このテストを通す本番変更は、単純なindex均等抽出ではなく、隣接形状を再評価する`TrackShapeSimplificationCore.reduce`である。

- [ ] **Step 2: 形状Coreテストを実行してREDを確認する**

Run:

```powershell
node --test tests/track-shape-simplification-core.test.js
```

Expected: `TrackShapeSimplificationCore`が未定義のためFAILする。

- [ ] **Step 3: Visvalingam–Whyatt簡略化を実装する**

次のデータ構造を使う。

```js
const node = {
  segmentIndex: segmentIndex,
  pointIndex: pointIndex,
  point: point,
  previous: previousNode,
  next: nextNode,
  protected: protectedPoint,
  removed: false,
  version: 0
};
```

実装規則:

- 全nodeをsegmentごとの双方向linked listにする。
- segment端、`point.time !== ''`、全segmentを通した最初の標高最小・最大nodeを`protected`にする。
- 削除候補の面積は中央pointを原点にした局所equirectangular投影の三角形面積とする。
- 経度差は`((delta + 540) % 360) - 180`で`[-180, 180)`へ正規化する。
- min-heapの比較順は`area`、`segmentIndex`、`pointIndex`とする。
- node削除時に左右を接続し、左右候補の`version`を増やして再計算結果をheapへ追加する。
- heapから取り出した古いversionは無視する。
- point総数が`maxPoints`になるか候補がなくなるまで削除する。
- 最後に元segment順・point順で未削除nodeを新しい配列へ投影する。

公開境界:

```js
const TrackShapeSimplificationCore = (function() {
  function reduce(segments, maxPoints) {
    // 上記規則で決定的に簡略化する。
  }
  return Object.freeze({ reduce: reduce });
})();
```

- [ ] **Step 4: 形状Coreテストを実行してGREENを確認する**

Run:

```powershell
node --test tests/track-shape-simplification-core.test.js
```

Expected: 全テストPASS。

- [ ] **Step 5: Task 2をコミットする**

```powershell
git add index.html tests/track-shape-simplification-core.test.js
git commit -m "feat: preserve GeoJSON shape during point reduction"
```

---

### Task 3: GeoJSON source batch・重複除去・圧縮・分割

**Files:**
- Modify: `index.html:14370-14730`
- Modify: `tests/geojson-track-interchange-core.test.js`
- Modify: `tests/geojson-track-import-ui.test.js`

**Interfaces:**
- Produces: `GeoJsonTrackInterchangeCore.buildDraftBatch(input, options)`
- Returns:

```js
{
  drafts,
  baseName,
  summary,
  stats: {
    sourcePointCount,
    pointCount,
    duplicatePointCount,
    interruptionCount,
    timeCompressedPointCount,
    shapeCompressedPointCount,
    compressedPointCount,
    generatedTrackCount,
    compressionIntervals
  },
  warnings: []
}
```

- Produces: `GeoJsonTrackInterchangeCore.updateDraftBatch(batch, patch)`
- Produces: `GeoJsonTrackInterchangeCore.toSavePayloads(batch)`
- Preserves: `parse`、`buildDraft`、`updateDraft`、`toSavePayload`

- [ ] **Step 1: GeoJSON batchの全要求を表す失敗テストを書く**

`tests/geojson-track-interchange-core.test.js`へ次を追加する。

```js
test('GeoJSON batch removes exact timed duplicates before the saved point limit', () => {
  const coordinate = [139, 35, 10];
  const time = '2026-01-01T00:00:00Z';
  const batch = loadCore().geo.buildDraftBatch(collection([
    feature(line(Array.from({ length: 20001 }, () => coordinate.slice())), {
      coordTimes: Array.from({ length: 20001 }, () => time)
    })
  ]), { sourceName: 'walk.geojson' });
  assert.equal(batch.stats.sourcePointCount, 20001);
  assert.equal(batch.stats.duplicatePointCount, 20000);
  assert.equal(batch.drafts.length, 1);
  assert.equal(batch.drafts[0].summary.pointCount, 1);
});
```

別々のテストとして次を追加する。

- 時刻なし、座標差、標高差、時刻差、別segmentの同一pointを重複除去しない。
- 日付をまたぐ2秒差は1 draft、4時間以上の差は2 draftと`interruptionCount: 1`。
- 20,001個の1秒pointは2秒間隔を選び10,001 pointになり、segment端と時刻を保持する。
- 20,001個の時刻なし折れ線は形状圧縮で20,000 pointになり、`shapeCompressedPointCount: 1`。
- 20,001個の10秒間隔pointは5秒圧縮で減らないため`[20000, 1]`へ追加分割する。
- 時刻付きpointと時刻なしpointが混在しても、形状圧縮で時刻付きpointを削除しない。
- 標高最高・最低、急カーブ、日付変更線のpointをbatch結果でも維持する。
- 20,000以下では入力point列を一切変更しない。
- source 100,000 pointを受け付け、100,001で`GEOJSON_SOURCE_POINT_LIMIT_EXCEEDED`。
- 20個の4時間中断は20 drafts、21個は`GEOJSON_GENERATED_TRACK_LIMIT_EXCEEDED`。
- 200 source segmentを受け付け、201で`TRACK_SEGMENT_LIMIT_EXCEEDED`。
- エラー時に`generateId`呼出し回数が0。
- 100文字base nameの複数出力がsuffixを含め100文字以内。
- `updateDraftBatch`が全draftの共通metadataとsuffixを更新する。
- `toSavePayloads`がdraft順に独立payloadを返す。

`tests/geojson-track-import-ui.test.js`へ4時間中断を持つ小さなGeoJSONを追加し、実UI Core経由で次を検証する。

```js
const batch = await setup.controller.importFile(file({
  text: async () => interruptedGeoJson()
}));
assert.equal(batch.drafts.length, 2);
assert.deepEqual(plain(batch.drafts.map((draft) => draft.name)), [
  'Morning walk(1/2)', 'Morning walk(2/2)'
]);
```

同じintegration testで共通編集後のsuffix、2 payloadの直列保存、2件目失敗後の固定ID・revision・payload再試行、`onSaved`のtrackId単位upsertを確認する。

- [ ] **Step 2: GeoJSON batchテストを実行してREDを確認する**

Run:

```powershell
node --test --test-name-pattern "GeoJSON batch|GeoJSON generated routes" tests/geojson-track-interchange-core.test.js tests/geojson-track-import-ui.test.js
```

Expected: `buildDraftBatch`が未定義のためFAILする。

- [ ] **Step 3: source収集を保存用上限から分離する**

`addSegment`と`appendFeatureSegments`へcollection optionsを渡し、次を実装する。

```js
const MAX_SOURCE_POINTS = 100000;
const MAX_GENERATED_TRACKS = 20;
const INTERRUPTION_MS = 4 * 60 * 60 * 1000;
const COMPRESSION_INTERVAL_SECONDS = [1, 2, 3, 4, 5];

function duplicatePointKey(point) {
  if (!point.time) return '';
  return [
    point.time,
    point.lat,
    point.lng,
    point.elevation === null ? 'null' : point.elevation
  ].join('|');
}
```

source収集規則:

- `sourcePointCount + coordinates.length`をmap前に検査する。
- batchでは100,000、既存`parse`では20,000を上限にする。
- 全coordinateと対応timeを正規化してから、segmentローカル`Set`で完全重複を除去する。
- `sourcePointCount`は除去前、`pointCount`は除去後とする。
- source segmentは200件で拒否し、圧縮前にsegmentを捨てない。

- [ ] **Step 4: GeoJSON batch変換とdraft APIを実装する**

`buildDraftBatch`は`TrackBatchTransformCore.transform`へ次を渡す。

```js
const transformed = TrackBatchTransformCore.transform(source.segments, {
  maxPoints: trackGeometry.MAX_POINTS,
  interruptionMs: INTERRUPTION_MS,
  compressionIntervals: COMPRESSION_INTERVAL_SECONDS,
  reduceOverflow: function(segments, maxPoints) {
    return TrackShapeSimplificationCore.reduce(segments, maxPoints);
  }
});
```

各partを既存`normalizeDraft`で独立draftへ変換し、IDは全変換成功後に文書順で生成する。`TrackBatchTransformCore.partName`と`TrackBatchTransformCore.aggregateSummaries`を使い、20件超過時は`GEOJSON_GENERATED_TRACK_LIMIT_EXCEEDED`で拒否する。

stats mapping:

```js
const stats = {
  sourcePointCount: source.sourcePointCount,
  pointCount: drafts.reduce(function(total, draft) {
    return total + draft.summary.pointCount;
  }, 0),
  duplicatePointCount: source.duplicatePointCount,
  interruptionCount: transformed.interruptionCount,
  timeCompressedPointCount: transformed.timeCompressedPointCount,
  shapeCompressedPointCount: transformed.overflowCompressedPointCount,
  compressedPointCount: transformed.compressedPointCount,
  generatedTrackCount: drafts.length,
  compressionIntervals: transformed.compressionIntervals.slice()
};
```

`updateDraftBatch`は共通metadataを全draftへ適用し、`partName`でsuffixを再生成する。`toSavePayloads`はdraft順に既存`toSavePayload`を適用する。`warnings`は空配列を維持し、座標やpropertiesを含めない。

安全メッセージ用codeは次で固定する。

- `GEOJSON_SOURCE_POINT_LIMIT_EXCEEDED`
- `GEOJSON_GENERATED_TRACK_LIMIT_EXCEEDED`

- [ ] **Step 5: GeoJSON Core・UI integrationを実行してGREENを確認する**

Run:

```powershell
node --test tests/track-batch-transform-core.test.js tests/track-shape-simplification-core.test.js tests/geojson-track-interchange-core.test.js tests/geojson-track-import-ui.test.js tests/gpx-track-interchange-core.test.js
```

Expected: 全テストPASS。

- [ ] **Step 6: Task 3をコミットする**

```powershell
git add index.html tests/geojson-track-interchange-core.test.js tests/geojson-track-import-ui.test.js
git commit -m "feat: adaptively normalize GeoJSON tracks"
```

---

### Task 4: GeoJSON adapter・Preview表示

**Files:**
- Modify: `index.html:15635-15910`
- Modify: `tests/geojson-track-import-ui.test.js`

**Interfaces:**
- GeoJSON file size: `5 * 1024 * 1024`
- Adapter consumes: `buildDraftBatch`、`updateDraftBatch`、`toSavePayloads`
- Preview consumes: Task 3の`batch.stats`
- Save behavior: 既存`TrackFileImportUI`による文書順の`saveTrackBundle`

- [ ] **Step 1: 5MB file境界の失敗テストへ更新する**

`tests/geojson-track-import-ui.test.js`の既存file sizeテストを次へ変更する。

```js
assert.ok(await exact.controller.importFile(file({ size: 5 * 1024 * 1024 })));
input.files = [file({
  size: 5 * 1024 * 1024 + 1,
  text: async () => { reads += 1; return geoJson(); }
})];
assert.equal(
  documentApi.getElementById('geojson-track-operation-error').textContent,
  'GeoJSONファイルは5MB以内にしてください。'
);
```

Run:

```powershell
node --test --test-name-pattern "file selection resets" tests/geojson-track-import-ui.test.js
```

Expected: 現在は2MB上限のためFAILする。

- [ ] **Step 2: adapterと安全メッセージを5MBへ変更してGREENを確認する**

`GeoJsonTrackImportAdapter.maxFileBytes`を`5 * 1024 * 1024`へ変更し、`GEOJSON_FILE_TOO_LARGE`メッセージを5MBへ更新する。次のcodeも固定メッセージへ追加する。

```js
if (code === 'GEOJSON_SOURCE_POINT_LIMIT_EXCEEDED') {
  return 'GeoJSONの元point数は100,000件以内にしてください。';
}
if (code === 'GEOJSON_GENERATED_TRACK_LIMIT_EXCEEDED') {
  return 'GeoJSONから生成されるルート数が20件を超えています。';
}
```

Run:

```powershell
node --test --test-name-pattern "file selection resets" tests/geojson-track-import-ui.test.js
```

Expected: PASS。

- [ ] **Step 3: GeoJSON複数Previewの失敗テストを書く**

4時間中断を持つ小さなGeoJSONをfixture helperで生成し、次を検証する。

```js
const batch = await setup.controller.importFile(file({
  text: async () => interruptedGeoJson()
}));
assert.equal(batch.drafts.length, 2);
assert.deepEqual(plain(batch.drafts.map((draft) => draft.name)), [
  'Morning walk(1/2)', 'Morning walk(2/2)'
]);
assert.match(
  setup.documentApi.getElementById('track-import-preview-stats').textContent,
  /元point 3/
);
assert.match(
  setup.documentApi.getElementById('track-import-preview-stats').textContent,
  /生成ルート 2件/
);
assert.match(
  setup.documentApi.getElementById('track-import-preview-parts').textContent,
  /1\..*2\./s
);
```

Run:

```powershell
node --test --test-name-pattern "GeoJSON multiple preview" tests/geojson-track-import-ui.test.js
```

Expected: GeoJSON `renderStats`が空文字のためFAILする。

- [ ] **Step 4: GeoJSON処理統計を安全に表示する**

`GeoJsonTrackImportAdapter.renderStats`を次の項目だけで構築する。

```text
元point N / 保存point N / 完全重複除去 N / 記録中断 N件 /
時刻圧縮除去 N / 形状圧縮除去 N / 生成ルート N件 /
圧縮間隔 2秒・5秒
```

`compressionIntervals`は重複を除き、空なら圧縮間隔文言を出さない。入力JSON、座標、properties、ファイルパスを連結しない。

Run:

```powershell
node --test --test-name-pattern "GeoJSON multiple preview" tests/geojson-track-import-ui.test.js
```

Expected: PASS。

- [ ] **Step 5: Task 4をコミットする**

```powershell
git add index.html tests/geojson-track-import-ui.test.js
git commit -m "feat: surface adaptive GeoJSON import results"
```

---

### Task 5: README・全体回帰・最終レビュー

**Files:**
- Modify: `README.md:45-55`
- Modify: `README.md:198-205`
- Modify: `README.md:238-246`
- Modify: `tests/geojson-track-import-workflow.test.js`

**Interfaces:**
- Documentation reflects: 5MB、100,000 source point、重複除去、最大5秒、形状圧縮、4時間中断、20,000 point／20生成トラック
- Server contract remains: 1 payload最大20,000 point

- [ ] **Step 1: README契約テストを新仕様へ更新してREDを確認する**

`tests/geojson-track-import-workflow.test.js`のREADMEテストを次の契約へ変更する。

```js
assert.match(
  readme,
  /GeoJSONトラック[^\n]*LineString／MultiLineString[^\n]*5MB[^\n]*100,000 source point/
);
assert.match(
  readme,
  /GeoJSON[^\n]*完全重複[^\n]*最大5秒[^\n]*形状[^\n]*4時間[^\n]*20,000 point/
);
```

Run:

```powershell
node --test --test-name-pattern "README summarizes GeoJSON" tests/geojson-track-import-workflow.test.js
```

Expected: READMEが旧2MB・20,000 point記述のためFAILする。

- [ ] **Step 2: READMEを更新してGREENを確認する**

制約表とトラック説明を設計書どおり更新する。手動確認チェックには、実GeoJSONで次を確認する項目を残す。

- 時刻あり最大5秒圧縮
- 時刻なし形状圧縮
- 4時間中断
- 複数suffix
- segment、時刻、標高、表示順

Run:

```powershell
node --test --test-name-pattern "README summarizes GeoJSON" tests/geojson-track-import-workflow.test.js
```

Expected: PASS。

- [ ] **Step 3: 関連テストをまとめて実行する**

Run:

```powershell
node --test tests/track-batch-transform-core.test.js tests/track-shape-simplification-core.test.js tests/geojson-track-interchange-core.test.js tests/geojson-track-import-ui.test.js tests/geojson-track-import-workflow.test.js tests/gpx-track-interchange-core.test.js tests/gpx-track-import-workflow.test.js tests/track-file-import-common.test.js tests/track-storage.test.js tests/photo-track-match-interpolation.test.js
```

Expected: 全テストPASS、warning・unhandled rejectionなし。

- [ ] **Step 4: 本番コードと全Nodeテストを検証する**

Run:

```powershell
pnpm run check
```

Expected:

- `node --check Code.js` exit 0
- 全`tests/*.test.js` PASS
- failure、cancelled、skipped、todoが0

- [ ] **Step 5: 差分と互換性を自己レビューする**

Run:

```powershell
git diff --check
git status --short
git diff -- index.html README.md tests
```

確認項目:

- GeoJSON以外のPoint取込、GPX、保存済みトラック編集、共有形式へ意図しない変更がない。
- `buildDraft`の既存20,000 point拒否テストが残っている。
- `saveTrackBundle`へ送る各payloadは20,000 point以下。
- 形状圧縮後のsummaryが保存pointから再計算される。
- batch statsやエラーへ座標・properties・stackが入らない。
- ユーザーの既存未コミット変更を含めていない。

- [ ] **Step 6: READMEと実装をコミットする**

```powershell
git add README.md tests/geojson-track-import-workflow.test.js
git commit -m "docs: describe adaptive GeoJSON imports"
```

- [ ] **Step 7: 完了前のfresh verificationを実行する**

Run:

```powershell
pnpm run check
git status --short --branch
git log -5 --oneline
```

Expected: 全テストPASS、worktree clean、今回の実装コミットが現在ブランチ先頭に並ぶ。

- [ ] **Step 8: 現在のブランチをpushする**

```powershell
git push origin fix/inherited-drive-sharing
git ls-remote --heads origin fix/inherited-drive-sharing
git rev-parse HEAD
```

Expected: remote SHAとlocal HEADが一致する。force pushは使わない。
