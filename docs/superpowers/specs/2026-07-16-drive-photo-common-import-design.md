# Drive写真共通取込設計

## 目的

Drive固有処理を、Drive APIレスポンスから安全なブラウザ`File`を生成する境界に限定する。生成後は端末写真と同じ`startPhotoImportFromFiles()`、`MultiPhotoImportBuilder`、複数写真Preview、`ImportPhotoItemProcessor`、`saveImportPhotoItem()`を使用する。

## アーキテクチャ

- `readDrivePhotoImportFile()`は編集権限、取込ルート配下、trash、shortcut、対応形式、実読込バイトの15MB上限を検証し、実際に読み取ったbytesから`sizeBytes`とBase64を返す。`file.getSize()`と実バイト長の一致は要求しない。
- `DrivePhotoImportSourceCore.validateFileResponse()`は`ok === true`、要求IDとの一致、安全なファイル名、JPEG/PNG/WebP/HEIC/HEIF、正規Base64、復号後1〜15MBだけを信頼境界で検証する。一覧descriptorの`name`、`mimeType`、`sizeBytes`、`modifiedAt`、`kind`との一致は要求しない。
- `DrivePhotoImportLoader`は検証済み実データからブラウザ`File`を作り、`modifiedAt`は有効な場合だけ`lastModified`補助値に使う。その`File`と対応する`sourceDriveFileId`を既存共通取込入口へ渡す。
- 写真取込の初期値と戻り先は単品ピンフォームから独立させる。Drive picker、同意画面、共通写真準備から戻る場合は`add-menu-overlay`へ戻し、`upload-overlay`を開かず、単品フォームの値を読み書きしない。

## 画面遷移

- Drive pickerキャンセル: 同意画面とpickerを閉じ、追加方法選択へ戻る。
- 同意画面の戻る: 選択を維持してpickerへ戻る。
- Driveファイル読込失敗: pickerを開いたまま安全なエラーを表示する。
- File生成成功: pickerを閉じ、共通写真準備へ進む。
- 共通写真準備キャンセルまたは準備開始失敗: 追加方法選択へ戻る。
- 単品ピン登録画面は、追加方法選択で「ピンを追加」を選んだときだけ開く。

## データ保持

Drive固有値は各取込itemのruntimeにある`sourceDriveFileId`だけとし、既存processorが保存payloadへ投影する。Base64、一覧descriptor、Drive専用Preview状態、Drive専用保存APIは追加・保持しない。

## エラー処理

レスポンス検証またはBase64復号に失敗した場合はFileを共通経路へ渡さない。loader失敗中はpicker sessionと選択を維持し、再試行またはキャンセルを可能にする。共通経路へのhandoff自体が失敗した場合はDrive状態を解放し、追加方法選択に安全なメッセージを表示する。

## テスト

JPG/PNG/WebP/HEIC/HEIFのFile化、一覧/読込メタデータ差異、ID/Base64/形式/サイズ拒否、Builder ready化、共通Preview/processor/save経路、Driveキャンセル時の単品フォーム不変、読込失敗時picker維持、成功時だけ共通準備へ進むこと、ローカル写真経路の回帰をNodeテストで検証する。

## 制約

新しいバージョン管理、設定項目、Sheet列、Drive専用Preview、Drive専用保存API、Advanced Drive APIは追加しない。公開保存payloadと既存ローカル写真取込の形状は維持する。
