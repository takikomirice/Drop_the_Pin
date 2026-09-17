# Codex 開発メモ

- 開始時に `docs/DEVELOPMENT.md` と `gas-project.json` を確認する。
- Windows PowerShellでは `.\npm.cmd run <command>` を使用できる。
- GASの対象は `gas-project.json` のテストプロジェクト。`gas:doctor` で認証と対象を確認する。
- リモート確認は `gas:snapshot` を使い、作業ツリーへ直接 `clasp pull` しない。
- GAS反映は、差分レビュー後に `gas:push -- --reviewed <snapshot>` を使う。現行リモートのマニフェストと固定バージョンを維持する。
- `appsscript.json` の匿名公開設定をテストGASへ直接pushしない。現在のテストWebアプリは自分のみ。
- `.codex-remote/`、OAuth情報、編集キー、編集URLをGitへ追加しない。
- アプリ変更は `validate` で検証し、ブラウザの `/dev` で実GASの動作も確認する。ローカルハーネスは地図・GASをスタブ化した部品確認用。
