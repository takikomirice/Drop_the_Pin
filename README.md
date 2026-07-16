# Drop the Pin

Google Apps Script と Leaflet.js で動作する、写真・ピン・ルートを地図上で管理・共有するアプリです。写真の位置情報を利用した自動配置、GPSのない写真の手動配置、複数形式のデータ交換、閲覧専用の共有ビューに対応しています。

## 主な機能

### ピンと写真

- 写真付き／写真なしピンの追加、編集、複製、移動、未配置化、削除
- EXIF GPSによる自動配置と、EXIF撮影日時によるイベント時刻入力
- GPSのない写真を未配置で保存し、あとから地図上へ配置
- JPEG、PNG、WebP、GIF、HEIC / HEIFに対応
- HEIC / HEIFはGPSと撮影日時を読み取ったあと、ブラウザ内でJPEG へ変換してから Google Drive へ保存
- 端末またはGoogle Driveから最大20枚の写真をまとめて取り込み
- 写真のない既存ピンへの写真追加と、編集・共有画面の全体表示ビューア
- タイトル、説明、タグ、色、アイコン、状態、イベント時刻、参考URLの管理
- 複数ピンの選択、一括状態変更、一括タグ・アイコン変更、一括削除

### データ交換とルート

- CSV／GeoJSON Pointによる全ピンのエクスポートと写真なしピンのインポート
- GPX 1.0 / 1.1およびGeoJSONのLineString／MultiLineStringによるトラック取込
- 写真の撮影時刻とトラック時刻を照合し、GPSのない写真の配置候補を表示
- ピンルートの作成、ピン追加、並べ替え、表示切替、直線／道路沿い表示、循環ルート
- GPX／GeoJSONトラックの並べ替え、表示設定編集、地図フィット、削除
- ピンルート、GPXルート、GeoJSONルートを統合した一覧と共有表示

### 閲覧と操作

- タイトル・説明・タグ検索、色・アイコン・状態・配置状態・並び順による絞り込み
- ダーク／ライトテーマと、PC・タブレット・スマートフォン向けレスポンシブUI
- 入力プリセットの追加、編集、複製、有効切替、並べ替え、明示適用
- 通常URLの閲覧専用画面、編集URLの編集画面、条件付き共有リンク
- 共有ビューでの検索、フィルタ、ルート表示切替、キャッシュ済み道路ルート表示
- 共有リンク一覧から共有ビューURLをコピーし、共有リンクごとにQRコードを表示

## 対応形式と主な上限

| 対象 | 対応形式・上限 |
| --- | --- |
| 単体写真 | JPEG、PNG、WebP、GIF、HEIC / HEIF |
| 複数写真 | JPEG、PNG、WebP、HEIC / HEIF、1回1〜20枚、合計100MBまで |
| Drive写真 | 1ファイル15MBまで、取込対象は設定したルートフォルダ配下 |
| CSVピン | UTF-8、2MB・20データ行まで |
| GeoJSONピン | FeatureCollection、Pointまたはgeometry: null、2MB・20 Featureまで |
| GeoJSONトラック | LineString／MultiLineString、2MB、20 Feature、200 segment、20,000 pointまで |
| GPXトラック | GPX 1.0 / 1.1、trk／rte、5MB、200 segment、20,000 pointまで |
| 保存済みトラック | 100件まで、線幅は4で統一 |
| 入力プリセット | 100件まで |
| ピンのタグ | 1ピン5件まで |

上限を超えたデータは切り捨てや自動分割を行わず、取込前にエラーとして扱います。CSV／GeoJSON Pointのエクスポートには写真、Drive ID、編集・共有トークン、ルート、取込管理情報を含めません。

## 位置情報と安全設計

### 位置情報・写真

Drop the PinはEXIF GPSを使って写真の撮影場所へピンを配置します。写真側の位置情報設定や書き出し・共有方法によってGPSや撮影日時が削除されている場合は、地図上で位置を選ぶか未配置で保存してください。

位置情報にはプライバシー上の注意が必要です。自宅・通学路・個人が特定される場所の写真は投稿しないでください。共有リンクを作成する前に、写真・説明・タグ・リンク・地点に個人情報や詳細すぎる位置情報が含まれていないか確認してください。

アプリが新規作成する表示用画像だけをリンク共有へ設定します。元のDrive写真の共有設定は変更しません。組織でリンク共有が禁止されている場合は、管理者と公開可能な保存先を確認してください。

### 編集URLと共有ビュー

- 通常URLでは閲覧専用です。編集には`mode=edit&editKey=`を含む編集URLを使います。
- 編集URLはconfig シートの `EDIT_URL` から開きます。必要に応じてメニューの「編集URLを更新・開く」で再生成できます。
- EDIT_KEY は config シートで管理します。EDIT_KEY は個人認証ではなく共有鍵であり、編集URLを知っている人は共同編集できます。
- WEB_APP_URL は config シートで管理するデプロイ済みWebアプリURLです。EDIT_URL を知っている人は共同編集できるため、共有範囲を限定してください。
- 6時間で切れるのは EDIT_KEY ではなく一時編集トークンです。開きっぱなしで編集できなくなった場合は、同じ編集URLを再読み込みしてください。
- 編集キーを再生成すると EDIT_URL も更新され、古い編集URLは無効になります。
- 共有ビューは閲覧専用で、ピンやルートの追加・編集・削除はできません。
- QRは閲覧専用共有ビューURLです。QRには編集URL、editKey、編集トークンを含めません。外部QR生成APIを使わずブラウザ内で生成します。
- 共有前に、写真・説明・地点・タグ・リンクの公開範囲を確認してください。

### 地図・検索・経路サービス

地図はLeaflet.js + OpenStreetMapを使用し、Google Maps API は使っていません。住所 / 地名検索にはNominatim、道路沿いルートにはOSRM public demo serverを使う場合があります。OSRM public demo serverは本番・大規模利用向けではありません。大量検索や高頻度利用を避け、必要に応じて自己ホストや代替サービスを検討してください。

## 技術構成

- Backend: Google Apps Script
- Frontend: `index.html`
- Shared view: `shared.html`
- Map: Leaflet.js + OpenStreetMap
- EXIF: `exif-js`
- HEIC / HEIF conversion: `heic-to@1.5.2`
- HEIC / HEIF metadata: `ExifReader@4.41.0`
- Drag and drop: SortableJS
- Storage: Google Drive、Google Spreadsheet
- Development: Node.js組み込みテストランナー、`clasp`

## セットアップとデプロイ

### 必要なもの

- Googleアカウント
- Google Apps Scriptプロジェクト
- 投稿データ用Googleスプレッドシート
- 画像保存用Google Driveフォルダ
- Node.js
- `clasp`

### 手順

1. `clasp`をインストールしてログインします。

   ```bash
   npm install -g @google/clasp
   clasp login
   ```

2. `.clasp.json.example`を`.clasp.json`へコピーし、Apps Scriptの`scriptId`を設定します。

   ```json
   {
     "scriptId": "YOUR_SCRIPT_ID_HERE",
     "rootDir": "."
   }
   ```

3. スプレッドシートに紐づいたApps Scriptから、`設定 -> 初期設定`を実行します。`setupSheet()`により必要なシートと不足ヘッダーが作成・補修されます。

4. `config`シートの`IMAGE_DRIVE_URL`へ画像保存用DriveフォルダURLを設定します。必要に応じて`RENAME_FILE_WITH_TITLE`を`true`にします。

5. テスト後にApps Scriptへ反映します。

   ```bash
   node --test tests/*.test.js
   clasp status
   clasp push
   ```

6. Apps Scriptの「デプロイを管理」からWebアプリの新しいバージョンを作成します。実行者とアクセス範囲を運用方針に合わせ、デプロイ後に通常URLと`EDIT_URL`を確認します。

v2.0.0へ更新する場合も、既存データを維持したまま新しいシートと列を補うため、デプロイ前後に初期設定を再実行してください。

## データ構造

| シート | 用途 |
| --- | --- |
| `map_info` | ピン、写真、位置、タグ、状態、イベント時刻、アイコン |
| `config` | Drive保存先、WebアプリURL、編集URL、動作設定 |
| `share_links` | 共有リンク、絞り込み条件、公開対象ルート |
| `routes` | ピンルート本体、表示設定、並び順 |
| `route_pins` | ルートに含めるピンと順序 |
| `route_cache` | 道路沿いルートの座標キャッシュ |
| `tracks` | GPX／GeoJSONトラックの概要、表示設定、並び順 |
| `track_segments` | トラックのsegment座標 |
| `input_presets` | タグ、色、アイコン、状態の入力プリセット |
| `import_receipts` | 写真・CSV・GeoJSON取込の重複防止と再試行管理 |

`map_info`のK〜O列が空の既存行も読み込めます。更新時の初期設定は既存行を削除せず、不足シートや末尾列を追加・補修します。破壊的なデータ移行はありませんが、作業前にスプレッドシートとDriveのバックアップを推奨します。

## フェーズ別概要

### Phase 1〜2: 基盤、安全、共有

写真付きピン、未配置ピン、検索・フィルタ、編集URL、編集トークン、ルート管理、共有リンクとQRを整備しました。通常URLと共有ビューは閲覧専用とし、変更操作を編集URLへ限定しています。

### Phase 3: 複数写真インポート

端末写真を最大20枚まで準備・確認・編集して登録できるようにし、キャンセル、再開、失敗項目の再試行、応答喪失時の重複防止を追加しました。

### Phase 4〜5: CSV／GeoJSON Point交換

写真を含まないピンをCSVまたは標準GeoJSONで入出力できるようにしました。既存IDと一致する`sourceId`があっても既存ピンを上書きせず、新規ピンとして取り込みます。

### Phase 6〜7: GeoJSON／GPXトラック

LineString、MultiLineString、GPXのtrk／rteを独立したトラックとして保存・表示できるようにしました。segment順を維持し、トラック間やsegment間を自動接続しません。

### Phase 8〜9: 写真時刻照合とDrive写真

写真時刻照合をPhase 8で追加し、撮影時刻とトラック時刻から配置候補を表示します。候補に表示するトラック時刻は現在のブラウザのローカル時刻です。Phase 9では設定済みDriveフォルダからの写真選択、既存ピンへの写真追加、取込元写真の保護を追加しました。

### Phase 9.5-F: UIと共有の統合

PC・タブレット・スマートフォン向けUIを刷新し、ピンルートと取込ルートの操作を統合しました。入力プリセット、データ画面、写真ビューア、アクセシビリティ改善、ピンルート・GPXルート・GeoJSONルートの共有表示を追加しました。

## 診断、テスト、手動確認

### ローカルテスト

```bash
node --test tests/*.test.js
```

### 起動診断

編集URLへ`debugStartup=1`を追加すると、編集画面だけに起動stageと`pin-add-ready`までの処理時間を表示します。診断のためにGASの呼び出しは追加しません。この診断は原因調査用であり、読み込みの高速化そのものは実装していません。

### 手動確認の要約

ローカル自動テストは実施していますが、Apps Script・実スプレッドシート・実ブラウザでの確認は未実施です。デプロイ前後に次を確認してください。

- [ ] 初期設定を再実行し、既存データを維持したまま必要なシートと列が補修される
- [ ] 端末とDriveから1枚／20枚の写真を取り込み、GPSあり／なし、キャンセル、再開、失敗再試行を確認する
- [ ] GPS付きiPhone HEIC、GPSなしiPhone HEIC、GPS付きJPEGを取り込み、JPEG変換、未配置保存と自動配置を確認する
- [ ] Chrome または EdgeとSafariで、写真を外して再選択した際に以前のプレビューや位置情報が残らないことを確認する
- [ ] CSV／GeoJSON Pointの1件／20件取込、上限超過、不正項目削除、エクスポートからの再取込を確認する
- [ ] YAMAPやGarminなどから出力した実GPXとGeoJSONトラックを取り込み、segment、時刻、標高、表示順を確認する
- [ ] 写真時刻照合でUTC offsetとカメラ時計補正を確認し、GPSのない写真へ正しい候補を適用できる
- [ ] ピンルート、GPXルート、GeoJSONルートを通常画面と共有画面で表示し、共有条件と表示順を確認する
- [ ] スマートフォン、タブレット、PCで、ダイアログ、検索、一覧、取込Preview、保存、破棄、再試行を操作できる
- [ ] Classroomなど実際の配布経路から編集URLを開き、閲覧モードと編集モードの権限制御を確認する

### Apps Scriptのテスト補助関数

`Code.js`には`testSaveMapData()`、`testUpdatePin()`、`testRouteCRUD()`があります。ただしPhase 1以降、変更系関数は編集トークンで保護されています。これらはApps Scriptエディタからの直接実行用の動作確認としては非推奨です。編集トークン免除やテスト専用の迂回経路は用意していないため、編集URLからの画面操作とローカルテストを利用してください。

## バージョン

- `v1.0.0`: 写真付きピン、未配置ピン、検索・フィルタ、基本的な地図編集
- `v1.1.0`: ピンアイコン、イベント時刻、Drive保存先、共有リンクとルート表示
- `v1.2.0`: 編集URL、一時編集トークン、共有QR、安全案内
- `v1.3.0`: 編集URL運用、HEIC / HEIF、シート書込みとUIの安定化
- `v2.0.0`: 複数写真・データ交換・GPX／GeoJSONトラック、Drive写真、UIと共有機能の大幅更新

各バージョンの詳細は[GitHub Releases](https://github.com/takikomirice/Drop_the_Pin/releases)を参照してください。

## 運用上の注意

- `appsscript.json`の`timeZone`は現在`America/New_York`です。運用地域に合わせて変更してください。
- Drive画像の共有可否はGoogle Workspaceの管理ポリシーに従います。
- NominatimとOSRMの公開サービスへ高頻度アクセスしないでください。
- `.claspignore`に`*.md`が含まれるため、READMEはApps Scriptへpushされません。

## ライセンス

[MIT License](LICENSE)
