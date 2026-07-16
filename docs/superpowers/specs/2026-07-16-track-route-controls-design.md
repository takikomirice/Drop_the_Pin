# GPX／GeoJSON表示設定編集とルート操作UI統一 設計

## 目的

GPX／GeoJSONルートの取込時と保存後の表示設定を、既存のピンルートUIに揃えて扱いやすくする。同時に、閉じたルートカードの高さと操作ボタンの見た目をPC／モバイルで統一する。

## スコープ

- ルート取込Previewの色選択を、既存ピン編集と同じ丸型カラースウォッチへ変更する。
- GPX／GeoJSONの線幅を常に4へ固定し、取込・編集画面から線幅入力を削除する。
- 保存済みGPX／GeoJSONへ表示設定の編集機能を追加する。
- 閉じたGPX／GeoJSONカードの高さを、閉じたピンルートカードに揃える。
- ピンルートとGPX／GeoJSONの操作ボタンを、PC／モバイルともアイコン中心の構成へ統一する。

## UI設計

### 色選択

`track-import-preview-color`は`select`ではなく、既存の`renderColorPaletteButtons()`と`.color-palette`／`.color-swatch`を再利用する。各色は44px四方のbuttonとし、色名を`aria-label`と`title`へ設定し、選択状態を`aria-pressed`とチェック表示で示す。

同じpaletteを新規取込と保存済みルート編集の両方で使用する。フォーム値はpaletteコンテナの正規化済み`value`で管理し、`PIN_COLORS`にない値は保存しない。

### 線幅

`track-import-preview-line-width`をDOM、編集対象一覧、フォーム同期、フォーム検証から削除する。GPX／GeoJSONの解析元に`lineWidth`があっても新規保存payloadは4へ正規化する。保存済みmetadataに1〜10の旧値が残っている場合は破損扱いにせず読み取るが、通常画面・Shared・再編集では4として扱い、編集保存時に4へ収束させる。

### ピンルートの操作列

ピンルートを開いたときの操作列は、PC／モバイルとも次の5個を同じ順で横一列に置く。

1. 追加
2. 時系列
3. Undo
4. 編集
5. 削除

全ボタンを44px四方以上のアイコンbuttonにし、可視テキストは置かない。SVGは装飾扱いにし、`aria-label`と`title`で操作名を提供する。320px幅でも5個が折り返さず収まるgapとpaddingにする。

従来の「詳細」ボタンは「編集」へ変更し、開くダイアログの内容は現在の詳細設定を維持する。削除は操作列の独立ボタンから既存の確認フローを起動する。編集ダイアログ内の重複する削除入口は除去する。

### GPX／GeoJSONの操作列

既存のヘッダー内「表示／非表示」と、展開時の「地図で全体表示」は維持する。展開領域では「地図で全体表示」を左側に残し、右端に44px以上のアイコンbuttonで次を並べる。

1. 編集
2. 削除

編集・削除は編集モードだけに表示する。閲覧モードと`shared.html`には追加しない。削除は既存の`deleteTrackFromUi()`を再利用する。

### カード高さ

統合ルートカードのGPX／GeoJSON要素から、旧トラック一覧用の`.track-item` classによるpadding重複を取り除く。閉じたカードはピンルートと同じ`.unified-route-card`、`.route-card-header`、`.route-summary`の寸法だけで構成し、名前やbadgeの内容にかかわらず同じ最小高さとpaddingにする。

## 保存済みトラック編集

### クライアント状態

既存の`track-import-preview-overlay`を新規取込と編集で共用する。編集用stateは取込ownerとは分離し、最低限次を保持する。

- 対象`trackId`
- 保存中フラグ
- 開始時snapshot

編集開始時に現在のtrackから、名前、説明、色、線種、表示状態をフォームへ入れる。概要欄は現在保存されているsourceType、sourceName、segment数、point数、距離、時刻、標高を読み取り専用で表示する。タイトルとprimary buttonは編集用の「ルートを編集」「保存」に切り替える。

保存成功時はサーバー応答を固定フィールドで検証し、既存trackのgeometry、revision、source情報、並び順を維持したままmetadataをstateへmergeする。対象Leaflet layerとサイドパネルを即時再描画し、取込成功時と同じ完了表示と「閉じる」操作へ遷移する。キャンセル時はサーバーを呼ばずsnapshotを変更しない。

### サーバーAPI

編集トークン必須の`updateTrackDisplaySettings(data)`を追加する。

入力:

- `trackId`
- `name`
- `description`
- `color`
- `visible`
- `lineStyle`

`lineWidth`は入力として受け付けず、保存値を4に固定する。ScriptLock内で`tracks`の対象metadataがちょうど1行であることと既存metadata全体を検証し、次だけを更新する。

- name
- description
- color
- updatedAt
- visible
- lineStyle
- lineWidth = 4

active revision、payload hash、summary、bounds、sourceType、sourceName、createdAt、orderIndex、`track_segments`、journal、retired hash、`share_links`は変更しない。対象行と拡張列の既存値／数式を保った1回の行書込みで更新する。

成功応答は`ok: true`と、クライアントmerge用に`trackId / name / description / color / visible / lineStyle / lineWidth / updatedAt`だけを含むsettingsを返す。

### 排他とエラー

- 編集モードかつ有効な編集トークンがある場合だけ開始する。
- 取込、トラック並び順保存、別トラック編集、削除、写真時刻照合で対象トラックを使用中の場合は開始または保存を拒否する。
- 保存中はフォームと対象の編集・削除操作を無効化し、二重送信を防止する。
- 編集保存中はbeforeunloadと既存のmutation pending判定へ含める。
- 対象なし、storage busy、metadata破損、metadata書込み失敗を固定error codeへ分離する。
- 生のサーバー例外、trackId、Sheet値を画面へ表示しない。
- サーバー成功後のクライアント反映失敗は再読み込みを案内する。

## テスト設計

### 色・線幅

- 取込Previewに色`select`と線幅inputが存在しない。
- paletteが全`PIN_COLORS`、44px target、色名、選択状態を持つ。
- GPX／GeoJSONの新規保存payloadは入力元にかかわらず`lineWidth: 4`になる。
- 旧metadataの1〜10は読み込めるが公開／通常DTOは4になる。

### カードと操作

- 閉じたピンルートとGPX／GeoJSONが同じpadding／最小高さ契約を使う。
- ピンルート操作が指定順の5アイコンで、可視テキストなし、各44px以上になる。
- ピンルート編集は既存詳細設定を開き、削除は独立確認フローを使う。
- GPX／GeoJSONは表示切替と全体表示を維持し、編集モードだけ編集／削除アイコンを持つ。
- 閲覧モードと`shared.html`には編集／削除がない。

### metadata更新

- 編集トークンを最初に検証する。
- 正常更新でgeometry、revision、source、order、createdAt、extension列、segmentsを変更しない。
- name、description、color、visible、lineStyle、lineWidth=4、updatedAtだけを更新する。
- 対象なし、重複metadata、busy、書込み失敗を区別する。
- キャンセルと二重押しでは更新APIを余分に呼ばない。
- 成功時にstate、Leaflet、カードへ即時反映し、失敗時は旧stateを維持する。

## 非対象

- 軌跡座標や標高、時刻の編集
- GPX／GeoJSON原本の再取込、保存、削除
- pin-routeの保存形式変更
- 新しいSheetや列の追加
- shared.htmlからの編集
- 既存トラック全行を事前移行する一括migration
