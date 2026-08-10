# tarasisterga4 (taraga4) 作業ルール

## clasp 運用ルール（厳守）

このマシンの `~/.clasprc.json` には複数のGoogleアカウントの認証情報が保存されており、
**`-u` 未指定（defaultプロファイル）は `otajigyokyo@gmail.com`（本プロジェクトとは無関係の別アカウント）** になっている。

本プロジェクトのGASプロジェクト「tarasister-analytics」（scriptId:
`12WiaV2dpGP9bTN_Rp_YSnxdvCNq0yFf9EPfmq8k0xjDjZ5NZKb3rhesU`）は
`tokyoflowerco.ltd@gmail.com` アカウントで認証する必要がある。

- **このプロジェクトで clasp コマンドを実行する際は、例外なく `-u tokyoflower` を付けること。**
- `-u` を付けない clasp コマンドの実行は禁止（誤って別アカウント・別GASプロジェクトを操作するリスクがあるため）。
- 誤実行防止のため、リポジトリ直下に `clasp.ps1` ラッパーを用意している。
  `.\clasp.ps1 <サブコマンド>` の形で呼び出せば、内部で自動的に `-u tokyoflower` が付与される
  （例: `.\clasp.ps1 push`, `.\clasp.ps1 status`）。
- Bashツールから直接 `clasp` を呼ぶ場合も、必ず `clasp -u tokyoflower ...` の形で実行すること。

## GASプロジェクトについて

- `コード.js` / `appsscript.json` はリポジトリ直下にあり、**スタンドアロン型**のGASプロジェクト
  （スプレッドシート紐付けではない）。
- `.clasp.json` は `.gitignore` 対象のため、新しい環境でクローンした直後は存在しない。
  新環境では以下の内容で作成すること（scriptIdは変わらない想定）:
  ```json
  {
    "scriptId": "12WiaV2dpGP9bTN_Rp_YSnxdvCNq0yFf9EPfmq8k0xjDjZ5NZKb3rhesU",
    "rootDir": "."
  }
  ```
- push前は必ず本番のGASコードとローカルの `コード.js` に差分がないか確認してから実行すること
  （本番側にリポジトリへ反映されていない手動修正が入っている可能性があるため）。
