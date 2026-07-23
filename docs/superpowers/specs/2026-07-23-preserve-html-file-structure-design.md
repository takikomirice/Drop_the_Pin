# v2.1.0 音声機能のHTMLファイル構成復元設計

## 状態

2026-07-23、利用者承認済み。v2.1.0の音声機能は維持しつつ、HTMLファイルをv2.0.0と同じ`index.html`、`shared.html`の2ファイルだけへ戻す。

## 背景

v2.0.0のHTMLは`index.html`と`shared.html`だけだった。音声機能の実装で`audio-editor.html`、`audio-player.html`、`audio-vendor.html`を追加したが、既存のHTMLファイル構成を変えないという要件に反する。

## 目標

- リポジトリ直下およびApps Script配信対象のHTMLを`index.html`と`shared.html`だけにする。Git追跡状態だけで判定せず、実在ファイルと`.claspignore`適用後の対象も検査する。
- 音声編集、新規音声ピン、既存ピンへの追加・差し替え・削除、編集／共有再生を維持する。
- 約388KBの音声vendorを初期HTMLへ出力せず、編集画面の主要表示後にだけ遅延取得する。
- `getAudioVendorBundle`の編集トークン認証、共有投影認可、音声ID非公開を維持する。
- v2.1.0のDrive／Spreadsheet構造と音声データを変更しない。

## 検討した方式

### A. 既存HTMLへ統合し、vendor領域だけサーバーで除外する（採用）

`audio-editor.html`を`index.html`へ、`audio-player.html`を`index.html`と`shared.html`へ統合する。vendorは固定マーカー付きで`index.html`へ置き、通常の`doGet`では評価前にその領域を除外する。認証済み`getAudioVendorBundle`だけが同じ領域を抽出して返す。

既存2ファイル構成、遅延取得、認証境界を同時に維持できる。Hemisphereで利用している方式と同じ境界を再利用できる。

### B. vendorを`Code.js`の巨大文字列へ移す（不採用）

HTML数は増えないが、生成物と業務ロジックが混在し、エスケープ、差分、Apps Scriptエディター上の保守性が悪化する。

### C. 追加HTMLを残す（不採用）

現在の共通化は維持できるが、利用者が指定したファイル構成を満たさない。

## ファイル構成

完了後のHTMLは次の2ファイルだけとする。

```text
index.html
shared.html
```

- `index.html`
  - 固定マーカーで囲んだ生成済みaudio vendor領域
  - 編集／通常画面用の最小音声プレーヤー
  - 編集トークンがある場合だけ評価される音声エディター
  - 現在の音声取込・既存ピン管理コード
- `shared.html`
  - `index.html`と同一内容の最小音声プレーヤーブロック
  - 共有音声取得コード
  - 音声エディターとvendorは含めない

`audio-editor.html`、`audio-player.html`、`audio-vendor.html`は削除する。

## vendorの生成と配信

`scripts/sync-audio-vendor.js`は`audio-vendor.html`を書かず、`index.html`内の次の領域だけを決定的に更新する。既存の`vendor/mediabunny-LICENSE.txt`と`vendor/mediabunny-mp3-encoder-LICENSE.txt`の同期は変更しない。

```text
AUDIO_VENDOR_BUNDLE_START
<script>
生成済みvendor source
</script>
AUDIO_VENDOR_BUNDLE_END
```

`Code.js`は`index.html`のraw sourceからマーカーを一度だけ検出する。通常のindex表示ではマーカー領域を除去した文字列からHTMLテンプレートを作り、vendor source、マーカー、vendor公開APIを初期レスポンスへ含めない。文字列から作るテンプレートにも、現在と同じ`execUrl`、`token`、`editToken`を設定する。共有表示は従来どおり`shared.html`を直接評価する。

`getAudioVendorBundle`は編集トークンを先に検証し、その後に`index.html`を読み、同じマーカー領域からvendor sourceだけを返す。マーカー重複、順序違反、script境界違反、危険な`</script>`、公開API欠落は失敗として扱う。

## エディターとプレーヤーの統合

音声エディターは現在の内容をそのまま`index.html`の編集トークン条件内へ移す。通常URLと共有画面には生成しない。

音声プレーヤーは固定の開始／終了マーカーを付けて`index.html`と`shared.html`へ置く。Nodeテストで両ブロックが同一であること、各ページに非表示`audio`要素が1件だけあることを固定する。プレーヤーのAPI、LRU、再生／停止、`nodownload`、共有認可は変更しない。

## エラー処理と安全境界

- index raw sourceの読込またはvendor領域検証に失敗した場合、vendor付きHTMLを返さずindex表示を失敗させる。
- vendor取得失敗は既存どおり通常の地図操作を止めず、エディターを開いた際に再試行する。
- 認証前にindex raw sourceからvendorを抽出しない。
- shared初期表示、100音声ピン、地図移動ではMP3本体もvendorも取得しない。
- Drive ID、音声ID、Base64、編集／共有トークンをエラーやログへ追加しない。

## テスト方針

実装前に、現在の3追加HTMLが存在するため失敗する次の契約を追加する。

1. リポジトリ直下の実在HTMLとApps Script配信対象HTMLが`index.html`と`shared.html`だけである。
2. editor、player、vendorの各領域が既存HTMLへ埋め込まれている。
3. index初期出力からvendor領域が除外され、認証済みvendor APIは同じ領域を返す。
4. vendor同期を2回実行しても`index.html`の差分が増えず、check modeは古い領域を検出する。
5. indexとsharedのplayerブロックが一致し、sharedにeditor／vendorがない。
6. ブラウザーハーネスが分割HTMLではなくinline marker blockを抽出して、本番テンプレートと同じ構造を評価する。

GREEN後に次を実行する。

- vendor同期／check、GAS JavaScript構文検査
- 音声vendor、エディター、プレーヤー、テンプレート、認証の集中Nodeテスト
- 全Nodeテスト
- 全Chromiumテスト
- v2.0.0比の初期HTML raw／gzip再計測
- 独立コードレビュー

ローカル合格はApps Script実環境の配信確認とは分けて報告し、`clasp push`やデプロイは別許可なしに行わない。

## Git方針

既にpush済みの履歴は書き換えない。設計書を先に独立コミットし、実装と検証を後続コミットとして`v2.1.0`へpushする。masterへのマージ、PR、タグ、リリース、デプロイは行わない。
