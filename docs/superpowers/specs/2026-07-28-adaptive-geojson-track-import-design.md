# GeoJSONルート重複除去・形状圧縮・分割 設計

## 目的

GPSログやGPX変換結果として作られた大容量GeoJSONを、20,000 pointを超えても安全に取り込めるようにする。時刻付きデータは既存GPXと同じ最大5秒間隔の圧縮を使い、時刻のないデータは単純な偶数間引きではなく、線形状の重要点を優先して残す。

保存済みトラックの1件200 segment・20,000 point上限、全体100トラック上限、既存のtrack payloadとSheets形式は変更しない。

## 対象データ

ルート取込で引き続き受け付けるのは、`FeatureCollection`内の`LineString`／`MultiLineString`だけとする。各pointから保存する情報は次のとおり。

- 経度・緯度
- 任意の標高
- `properties.coordTimes`または`properties.times`にある任意の時刻

ルート名、説明、色、表示状態、線種の既存metadata優先順位を維持する。任意のFeature properties、Feature ID、4番目以降の座標値など、現在保存していない情報を新たに保存対象にはしない。

## 基本処理

GeoJSONの保存用batch生成は次の順序で行う。

1. JSON、FeatureCollection、Feature、geometry、座標、標高、時刻、metadataをすべて検証・正規化する。
2. 同一元segment内の証明可能な完全重複pointを除去する。
3. 連続する時刻付きpoint間に4時間以上の空白があれば行程を分ける。
4. 各行程が20,000 pointを超える場合だけ、時刻付きpointを1～5秒の必要最小間隔で圧縮する。
5. まだ20,000 pointを超える場合、時刻のないpointを形状考慮方式で必要数だけ除去する。
6. それでも20,000 pointを超える場合、複数の保存トラックへ分割する。
7. 複数出力は`元の名前(1/N)`形式で命名し、既存の共通プレビューと直列保存を使う。

元GeoJSON文字列は変更しない。IDは解析・変換がすべて成功してから生成する。

## 入力上限

- GeoJSONファイル上限を2MBから5MBへ変更する。
- source point安全上限を100,000 pointとする。
- Feature上限は既存どおり20件とする。
- source segment上限は既存どおり200件とする。
- 生成できる保存トラックは1回20件までとする。
- 各保存トラックは200 segment・20,000 pointまでとする。
- 保存済みトラック全体は既存どおり100件までとする。

source point上限は重複除去・圧縮より前に数える。いずれかの入力上限を超えた場合は部分取込せず、安全な固定メッセージで拒否する。

## 完全重複の除去

同じ元segment内で、正規化後の次の値がすべて一致する時刻付きpointだけを完全重複とする。

- 時刻
- 緯度
- 経度
- 標高（未設定を含む）

最初のpointだけを残す。時刻のない同一座標は、停止、折り返し、自己交差などの意図を判別できないため、重複として除去しない。別segmentの同一pointもsegment境界を維持するため除去しない。

不正なpointを重複として読み飛ばしてはならず、すべてのsource pointを検証した後に重複判定する。

## 時刻による分割と圧縮

既存GPXと同じ規則を使う。

- 日付変更だけでは分割しない。
- 同一segment内の連続する時刻付きpointの差が4時間以上なら新しい行程とする。
- 時刻のないpointを挟む箇所は自動分割しない。
- 元segment境界は常に維持する。
- 20,000 point以下の行程は圧縮しない。
- 超過時は1、2、3、4、5秒を順に試し、20,000以下になる最小間隔を選ぶ。
- segment先頭・末尾は必ず残し、間隔判定はsegmentごとにリセットする。

時刻圧縮後に残った時刻付きpointは、形状圧縮では削除しない。これにより、採用した最大5秒間隔を超える追加削除を行わず、写真時刻照合に使う時系列を維持する。

## 時刻なしpointの形状圧縮

時刻圧縮後も20,000 pointを超える場合だけ、削除可能な時刻なしpointへVisvalingam–Whyatt方式の形状簡略化を適用する。

必ず残すpointは次のとおり。

- 各segmentの先頭と末尾
- 時刻圧縮後に残った全時刻付きpoint
- 各行程の標高最高点と最低点

残りの時刻なしpointについて、前後pointと作る三角形の実効面積が小さい順に除去する。面積は緯度・経度を局所的なメートル座標へ近似投影して計算し、経度差は日付変更線を考慮して正規化する。point除去後は隣接pointの実効面積を再計算する。

同じ実効面積の場合は元segment順・point順で決定し、同じ入力から常に同じ結果を得る。全行程で20,000 pointになるか、削除可能なpointがなくなるまで処理する。

この方式により、直線上の密集点を先に減らし、急カーブ、折り返し、segment端、標高範囲を優先して残す。元point順とsegment順は変更しない。

## 追加分割

時刻圧縮と形状圧縮の後も20,000 pointを超える場合、既存GPXと同じ規則で複数トラックへ分割する。

- segment境界を優先する。
- segment内で切る場合も各保存トラックを20,000 point以下にする。
- 分割でpointを捨てたり重複保存したりしない。
- 各保存トラックのsegment indexを0から振り直す。
- 生成数が20件を超える場合は保存前に拒否する。

別トラック間をLeafletで自動接続しない既存仕様を維持する。

## batchと共通処理

GeoJSON用に`buildDraftBatch`、`updateDraftBatch`、`toSavePayloads`を追加し、既存`TrackFileImportUI`の複数トラックPreview・共通編集・直列保存・冪等再試行を利用する。

GPXとGeoJSONで共通する次の純粋処理は、format非依存の内部Coreへ抽出する。

- segment／pointの複製と集計
- 4時間中断による行程分割
- 1～5秒間隔の時刻圧縮
- 20,000 point単位の追加分割
- suffix付き命名
- 複数summaryの集計

GPXの既存動作は維持し、時刻なしGPXの均等抽出を今回の形状圧縮へ無断変更しない。形状圧縮はGeoJSONのbatch生成からだけ選択する。

既存のGeoJSON単一APIである`parse`、`buildDraft`、`updateDraft`、`toSavePayload`は互換性のため維持する。通常の取込UIは新しいbatch APIを使う。

## Preview

GeoJSONの取込プレビューへ次を表示する。

- source point数
- 完全重複除去数
- 4時間以上の中断数
- 時刻圧縮除去数と採用間隔
- 形状圧縮除去数
- 保存point数
- 生成ルート数

複数出力時は既存GPXと同じ一覧で、保存名、開始・終了時刻、segment数、point数を示す。処理を行っていない項目を警告扱いにはしない。

生のGeoJSON、座標一覧、任意properties、ファイルパス、例外stackをPreviewやエラーへ表示しない。

## 写真時刻照合

写真時刻照合は保存後の各segmentに残る時刻付きpointを使う。最大5秒間隔までの時刻圧縮と、その間の線形補間という既存方針を維持する。

- 形状圧縮では時刻付きpointを削除しない。
- 4時間中断、元segment境界、追加分割をまたいで補間しない。
- 時刻のないGeoJSONは従来どおり撮影時刻だけでは照合できない。

## エラー処理

- 5MB、100,000 source point、20 Feature、200 source segment、20生成トラック、100保存トラックの上限を区別する。
- malformed JSON、Feature、geometry、座標、標高、時刻、metadataの既存エラー分類を維持する。
- 圧縮・分割後に空segmentや空トラックを生成しない。
- 解析・変換失敗時はIDを消費せず、1件も保存しない。
- 複数保存の途中失敗は、既存の固定payloadとrevision冪等性で再試行する。
- 内部例外や入力内容を利用者向けエラーへ露出しない。

## テスト設計

### GeoJSON Core

- 20,000 point以下ではpoint、順序、segment、時刻、標高を変更しない。
- 100,000 source point境界を検証し、上限超過を部分採用しない。
- 完全重複を保存上限判定前に除去する。
- 時刻なし、座標差、標高差、時刻差、別segmentのpointを重複除去しない。
- 4時間以上の中断だけを分割し、日付変更と4時間未満を分割しない。
- 1～5秒の最小圧縮間隔を選び、segment端を保持する。
- 時刻と座標が常に同じpointとして移動・除去される。
- 形状圧縮が直線上の点を優先して減らし、急カーブ、折り返し、segment端を残す。
- 標高最高点・最低点を残し、summaryを保存対象pointから再計算する。
- 日付変更線付近でも形状判定が大きく歪まない。
- 削除不能な20,000超過を複数トラックへ順序どおり分割する。
- suffix、100文字名、20生成トラックの境界を確認する。
- 既存`parse`／`buildDraft`の20,000 point契約を維持する。

### UI workflow

- GeoJSON adapterがbatch APIを使う。
- 単一出力は従来どおり1件のPreviewとpayloadになる。
- 複数出力は一覧、共通編集、suffix再生成、集計を表示する。
- GeoJSONの処理統計だけを安全に表示する。
- 複数payloadを直列保存し、途中失敗後も同じID・payloadで再試行する。
- 5MB境界をFile.text前に検証する。

### 回帰

- GeoJSONのFeatureCollection、LineString／MultiLineString、metadata、`coordTimes`／`times`の既存検証を維持する。
- GPXの重複除去、圧縮、分割、Preview、保存結果を変更しない。
- サーバーは20,000 pointを超える単一payloadを引き続き拒否する。
- 保存済みトラック編集、Leaflet描画、共有payload形式を変更しない。

## ドキュメント

READMEのGeoJSONトラック制約を次へ更新する。

- 5MB、最大100,000 source point
- 完全重複の自動除去
- 時刻付きpointの必要最小限・最大5秒間隔圧縮
- 時刻なしpointの形状考慮圧縮
- 4時間中断と20,000 point上限による複数トラック化
- 各保存トラック200 segment・20,000 point、1回最大20トラック

## 非対象

- Polygon、MultiPolygon、Point、GeometryCollectionのルート取込
- 任意Feature propertiesや4番目以降の座標値の保存
- 元GeoJSONファイルの書換え
- 時刻なしpointの重複推測
- 10秒以上の時刻圧縮
- ルートグループ用の新しい永続化形式
- server payload、tracks／track_segments、共有形式の変更
