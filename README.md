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

## 技術構成

- Backend: Google Apps Script
- Frontend: `index.html` 1ファイル構成
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

このリポジトリには `setupSheet()` が入っています。スプレッドシートに紐づいた状態でスクリプトを開き、シートを開いて `設定 -> 初期設定` を実行すると、`map_info` / `config` / `share_links` シートとヘッダーを作成できます。

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
  - `movePin()`: ピン位置更新
  - `unplacePin()`: 配置済みピンを未配置へ戻す
  - `deletePin()`: ピン削除 + 写真の削除
  - `getAppSettings()` / `updateAppSettings()`: 全体設定取得 / 更新
  - `setupSheet()`: シート初期化（A〜L 列。既存 J 列までのシートに K/L 列を補完する）
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

`index.html` には `google.script.run` の簡易モックも入っているので、Apps Script 外でもレイアウト確認がしやすくなっています。

## テスト用関数

`Code.js` には次のテスト補助関数があります。

- `testSaveMapData()`
- `testUpdatePin()`

Apps Script エディタから直接実行して、Drive 保存やスプレッドシート更新の確認に使えます。

## 注意点

- `appsscript.json` の `timeZone` は現在 `America/New_York` です。運用地域に合わせて変更してください。
- `Code.js` の ID 定数はプレースホルダーのままでは動きません。
- 画像は `ANYONE_WITH_LINK` で共有設定されます。公開範囲は運用方針に合わせて確認してください。
- `.claspignore` に `*.md` が入っているため、`README.md` は Apps Script へは push されません。

## 今後の導入検討

次の機能は、検索 / フィルタ / 並び替え・タグ / 状態管理・一括操作が導入済みとなった現時点では、後続フェーズの検討候補として残っています。

- 住所検索 / 現在地 / 座標入力
  - 地名検索、現在地への移動、緯度経度直接指定などの地図移動補助
- pinの一括入力 / 一括削除機能

## ライセンス

MIT License です。商用利用、改変、再配布が可能です。詳細は `LICENSE` を参照してください。

## 参考資料
