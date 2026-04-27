# Drop the Pin

Google Apps Script と Leaflet.js で作る、写真付きのマップ投稿アプリです。スマホで写真を投稿し、GPS がある写真はそのまま地図へ、GPS がない写真はあとから地図上で配置できます。

## できること

- 写真をアップロードして Google Drive に保存
- 投稿内容を Google スプレッドシートに保存
- Exif GPS を読んで自動でピン配置
- GPS なし写真を未配置として保存し、あとから地図上で配置
- 写真なしでタイトルだけのピンを保存
- モバイル向けボトムシート UI と PC 向けサイドパネル UI
- ダーク / ライトテーマ切り替え
- 閲覧モード / 編集モード切り替え
- 既存ピンの編集 / ドラッグ移動 / 未配置へ戻す / 削除
- ピンごとの参考 URL 登録
- 画面から Drive フォルダを開く導線
- タイトル・説明・タグによる検索
- 配置状態・運用状態・並び替えによるフィルタリング
- ピンへのタグ付け（`#tree #bird` 形式、最大5件）
- 運用状態管理（未対応 / 対応中 / 完了 / 保留）— 編集モード限定表示
- 編集モードでの複数選択と一括状態変更
- 編集モードでの複数選択ピン一括削除
- 住所 / 地名検索による地図移動
- ルートグループの作成、ピン追加、並び替え、表示切り替え
- ルート線の表示、線種設定、経路キャッシュ
- タグ / 色で絞り込んだ共有リンクの作成と閲覧専用共有ビュー
- 共有ビューでの条件クリア、直線ルート線、ルート内番号付きピン表示

## 技術構成

- Backend: Google Apps Script
- Frontend: `index.html` 1ファイル構成
- Shared view: `shared.html`
- Map: Leaflet.js + OpenStreetMap
- Exif: `exif-js`
- Storage:
  - Google Drive: 画像保存
  - Google Spreadsheet: 投稿データ保存
- Local dev / deploy: `clasp`

## ファイル構成

```text
.
├── Code.js
├── index.html
├── appsscript.json
├── .clasp.json.example
├── .claspignore
├── .gitignore
├── shared.html
└── tests/
```

## データ構造

保存先シート名は `map_info` です。ヘッダーは `setupSheet()` で作成できます。

| 列 | 項目 | 備考 |
| --- | --- | --- |
| A | タイムスタンプ | |
| B | タイトル | |
| C | 説明 | |
| D | 緯度 | |
| E | 経度 | |
| F | ピンの色 | |
| G | ファイルID | |
| H | 画像URL | |
| I | ID | |
| J | 参考URL一覧 | `\|` 区切り |
| K | 状態 | `未対応 / 対応中 / 完了 / 保留`、または空欄 |
| L | タグ | `\|` 区切り（内部表現）、UI 上は `#tag` 形式 |

`lat` / `lng` が空の行は未配置データとして扱われます。K / L 列が空の既存行も問題なく読み込まれます。

ルートと共有リンクは次のシートに保存されます。いずれも `setupSheet()` で作成できます。

| シート | 用途 |
| --- | --- |
| `share_links` | 共有ビュー用トークン、ラベル、タグ条件、色条件、有効 / 無効状態 |
| `routes` | ルートグループ本体、色、表示設定、線種、始点 / 終点、並び順 |
| `route_pins` | ルートに含めるピン ID とルート内の並び順 |
| `route_cache` | 通常画面の経路描画用の座標キャッシュ |

`config` シートには少なくとも次の設定値を置きます。

| 設定項目 | 意味 |
| --- | --- |
| `IMAGE_DRIVE_URL` | 写真を保存するルート Drive フォルダの URL |
| `RENAME_FILE_WITH_TITLE` | `true` のとき、写真付きピンのタイトル変更時に Drive 上の写真名も変更 |

## 事前準備

必要なもの:

- Google アカウント
- Google Drive の保存先フォルダ
- Google スプレッドシート
- Node.js
- `clasp`

`clasp` のインストール:

```bash
npm install -g @google/clasp
clasp login
```

## セットアップ

### 1. Apps Script プロジェクトを作成

Apps Script プロジェクトを用意し、スクリプト ID を控えます。

`.clasp.json.example` をコピーして `.clasp.json` を作成し、`scriptId` を設定します。

```json
{
  "scriptId": "YOUR_SCRIPT_ID_HERE",
  "rootDir": "."
}
```

### 2. スプレッドシートと Drive フォルダを用意

- 投稿データ保存用のスプレッドシートを作成
- 画像保存用の Drive フォルダを作成

### 3. `map_info` シートを初期化

このリポジトリには `setupSheet()` が入っています。スプレッドシートに紐づいた状態でスクリプトを開き、シートを開いて `設定 -> 初期設定` を実行すると、`map_info` / `config` / `share_links` / `routes` / `route_pins` / `route_cache` シートとヘッダーを作成できます。

初期化後、`config` シートの `IMAGE_DRIVE_URL` に画像保存用 Drive フォルダの URL を設定してください。必要に応じて、`RENAME_FILE_WITH_TITLE` を `true` に変更できます。

### 4. push

```bash
clasp push
```

## デプロイ

Apps Script エディタでウェブアプリとしてデプロイします。

- 実行者: 自分
- アクセスできるユーザー: 必要に応じて設定

デプロイ後、発行された Web アプリ URL にアクセスすると地図画面が開きます。

## 開発時の見方

- `Code.js`
  - `doGet()`: フロント画面を返す
  - `saveMapData()`: 写真あり / なしのピンを保存（`status` / `tags` 含む）
  - `getMapData()`: 一覧取得（`status` / `tags` 含む）
  - `updatePinDetails()`: タイトル / 説明 / 色 / URL / 状態 / タグ更新
  - `bulkUpdatePinStatus({ ids, status })`: 複数 ID の状態を一括更新
  - `bulkDeletePins({ ids })`: 複数 ID のピンを一括削除し、関連する写真 / ルート参照 / キャッシュを更新
  - `movePin()`: ピン位置更新
  - `unplacePin()`: 配置済みピンを未配置へ戻す
  - `deletePin()`: ピン削除 + 写真の削除
  - `createShareLink()` / `listShareLinks()` / `getSharedViewData()`: 共有リンク作成、管理、共有ビュー用データ取得（共有許可ピンとルートグループを返す）
  - `getRouteGroups()` / `saveRouteGroup()` / `setRoutePins()` / `deleteRouteGroup()` / `updateRoutesOrder()`: ルート管理
  - `getRouteCache()` / `putRouteCache()`: 経路描画用キャッシュの読み書き
  - `getAppSettings()` / `updateAppSettings()`: 全体設定取得 / 更新
  - `setupSheet()`: シート初期化（`map_info` は A〜L 列。既存 J 列までのシートに K/L 列を補完する）
  - `PinData.normalizeTags(values)`: タグ配列を正規化（先頭 `#` 除去・重複排除・5件上限）
  - `PinData.serializeTags(values)` / `PinData.deserializeTags(value)`: `|` 区切りへの変換
  - `PinData.normalizeStatus(value)`: 状態値のバリデーション
- `index.html`
  - 地図描画
  - アップロード UI（タグ・状態入力付き）
  - Exif 読み取り
  - 未配置一覧 UI
  - 既存ピンの右クリック / 長押しメニュー
  - ピンのドラッグ移動
  - テーマ切り替え
  - サイドパネルの検索 / フィルタ / 並び替え UI
  - 編集モード限定の複数選択と一括状態変更バー
  - 編集モード限定の一括削除バー
  - 住所 / 地名検索
  - ルート管理 UI
  - 共有リンク作成 / 管理 UI
- `shared.html`
  - 共有リンクから開く閲覧専用マップ
  - タイトル / 説明 / タグ検索、タグ / 色フィルタ、並び替え
  - 検索語、タグ、色条件のクリア
  - 住所 / 地名検索
  - 共有対象ルートの直線表示、線種反映、距離ポップアップ
  - ルート内順序にもとづく番号付きピン表示
  - 道路ルートは閲覧時に外部 API を呼ばず、現時点では直線表示に統一

`index.html` には `google.script.run` の簡易モックも入っているので、Apps Script 外でもレイアウト確認がしやすくなっています。

## テスト用関数

`Code.js` には次のテスト補助関数があります。

- `testSaveMapData()`
- `testUpdatePin()`
- `testRouteCRUD()`

Apps Script エディタから直接実行して、Drive 保存やスプレッドシート更新の確認に使えます。

ローカルでは、少なくとも `Code.js` の構文確認を実行できます。

```bash
node --check Code.js
```

## バージョン

- `v0.1.0`: `master` の初期リリース基準
- `v0.5.0`: 一括削除、住所検索、共有ビュー、ルート管理までを含むリリース基準
- `v1.0.0`: 共有ビューに条件クリア、ルートグループ配信、直線ルート線、ルート内番号付きピンを追加
- `v1.0.1`: `README.md` を v1.0 系の共有ビュー機能に合わせて更新

## 注意点

- `appsscript.json` の `timeZone` は現在 `America/New_York` です。運用地域に合わせて変更してください。
- `Code.js` の ID 定数はプレースホルダーのままでは動きません。
- 画像は `ANYONE_WITH_LINK` で共有設定されます。公開範囲は運用方針に合わせて確認してください。
- `.claspignore` に `*.md` が入っているため、`README.md` は Apps Script へは push されません。

## 今後の導入検討

次の機能は、検索 / フィルタ / 並び替え・タグ / 状態管理・一括操作・ルート管理・共有ビューのルート表示が導入済みとなった現時点では、後続フェーズの検討候補として残っています。

- 現在地 / 座標入力
  - 現在地への移動、緯度経度直接指定などの地図移動補助
- pinの一括入力機能
- 共有ビューのルート一覧、ルート別表示切り替え、ルート未所属ピンの表示制御
- キャッシュ済み道路ルートの共有ビュー表示

## ライセンス

MIT License です。商用利用、改変、再配布が可能です。詳細は `LICENSE` を参照してください。

## 参考資料
