# Phase 8-2 複数写真トラック時刻照合接続 設計

## 目的

複数写真Preview（`sourceType === "multi-photo"`）に限り、保存済みトラックと写真撮影時刻を `PhotoTrackMatchCore` で照合し、利用者が選択して「適用」した候補だけをImportJobの `lat` / `lng` へ反映する。照合実行だけでは座標を変更せず、既存GPSと照合後の手動編集を上書きしない。

## 採用アーキテクチャ

1. `PhotoTrackMatchPreviewCore` を純粋な状態・所有権レイヤーとして追加する。UTC offset parser/formatter、時刻付きトラック候補作成、写真projection、照合、選択、stale判定、明示適用、cleanupを担当し、DOM・GAS・File・Blob・XMLへ依存しない。
2. `ImportJobCore.applyLocationPatches()` を緯度経度限定のimmutable patch APIとして追加する。job/item identity、順序、runtime、preview URL、編集済みmetadata、upload stateを維持し、重複・unknown ID・partial/文字列/範囲外座標を拒否する。
3. `ImportPreviewUI.open()` に `trackMatch` 専用の構造化拡張点を追加する。固定DOMをUI自身がDOM APIと `textContent` で描画し、任意HTMLは受け取らない。他形式ではパネルを隠し、状態も作らない。
4. `MultiPhotoImportWorkflow` が複数写真Previewを開く直前に照合Coreを初期化し、Previewのdraft変更・適用jobを通常 `onDraftChange` / controller経路へ戻す。cleanup時は照合stateを全解除する。
5. 本番構成は `state.multiPhotoImport.trackMatch` と `state.tracks` を注入する。トラックを再解析・保存せず、現在revisionの正規化済みトラックだけを参照する。

## 状態と所有権

照合stateは設定値、Core結果、選択item ID、original coordinate snapshot、照合機能が現在所有する最小match概要、結果入力snapshot、stale/running/errorだけを保持する。mapはすべて `Object.create(null)` とし、File、Blob、preview URL、track/point配列、XML、token、server errorを保持しない。

Preview開始時に全itemの元座標をsnapshotする。projectionでは、照合所有中で座標が最後の適用値と一致する場合だけ `null/null` として再照合可能にする。元GPS、partial GPS、または手動座標は現在値をCoreへ渡す。適用後の座標を利用者が変更したら所有権を解除し、再適用対象外にする。`null/null` へ戻した場合は再照合可能にする。

結果作成時にtrack ID/revision、正規化options、item順・ID・capturedAt・effective coordinatesの安全なsnapshotを保持する。track/revision/options、item構成、capturedAt、effective original coordinatesの変更はstaleにする。適用時にも同じ条件を再検証し、古い結果を拒否する。

## UIとデータフロー

Previewの固定パネルはトラックselect、UTC offset、時計補正、最大gap、endpoint tolerance、照合、status/error、counts、warnings、item別結果checkbox、適用、クリアを持つ。候補trackは `TrackGeometryCore.normalizeTrack()` 可能で時刻付きpointが1件以上ある現在revisionだけとし、optionはDOM APIで生成する。

`照合` → safe projection → `PhotoTrackMatchCore.matchPhotos()` → 結果表示、の時点ではjobを変更しない。matched結果だけを既定選択する。`選択したN件に位置を適用` → 所有権/stale/job statusを再検証 → `ImportJobCore.applyLocationPatches()` → 通常draft change通知 → 既存保存フロー、の順とする。保存payloadへtrack ID/revision/match詳細は追加しない。

## エラー・排他・アクセシビリティ

Core error codeは固定の安全な日本語へ変換し、例外本文、座標一覧、point/XML/File/EXIF/stackを表示しない。照合中は二重実行を禁止し、ImportJobがidleでない場合は照合・適用を無効化する。既存の複数写真workflow所有中判定を維持するためCSV/GeoJSON/track/単体追加との排他とbeforeunload保護は継続する。

label関連付け、disabled理由、`aria-live="polite"`、`role="alert"`、写真名を含むcheckbox aria-label、適用後status focusでキーボード操作と通知を確保する。

## テスト方針

UTC offset境界、候補track filtering、安全表示、match前後の非mutation、選択適用、既存GPS/partial/手動編集保護、再照合、revision/options/capturedAt/item変更stale、immutable patch、保存payload非拡張、他形式非表示、cleanupをNode実行型テストで検証する。既存647テスト、`node --check Code.js`、`git diff --check`も再実行する。

## 対象外

単体写真、Driveフォルダ、自動適用、IANA/DST、EXIF OffsetTimeOriginal、track変更・保存、match詳細永続化、UI全面再設計、Apps Scriptデプロイ・実機確認は行わない。
