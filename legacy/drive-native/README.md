# 退役: Drive ネイティブ版（2026-02 〜 2026-08-23）

このフォルダは、**配信から外した**旧本番です。歴史資料として残しています。

- 形: GitHub Pages が配る静的な HTML/JS から、ブラウザが Google Drive REST API を直接叩く
- 入口: `reflection-journal.giga-school.com` の共通 URL 1 本（先生も児童も同じ）
- 保存先: 児童のポートフォリオは児童の Drive、先生のチャンネルは先生の Drive、相互に閲覧共有

## なぜ外したか

配布の形を「先生ごとにデプロイする」ものへ変えたためです。
共通 URL を 1 本持つ形は、運営者側で OAuth クライアントと制限付きスコープ（`drive.readonly`）の
審査を通し続ける必要があり、学校ごとの Workspace ポリシーで止まることも多くありました。
いまは、スプレッドシートにコンテナバインドした Google Apps Script を先生がコピーして公開します。
運営者のアカウントも中央の設定も要りません。

## いまここに残っている理由

`kit/` を、他アプリを「Drive ネイティブ・分散ポートフォリオ」方式へ移すための部品として
公開しているためです（`PORTING_FROM_GAS.md` と `kit/README.md`）。
`scripts/test-legacy-drive-*.mjs` と `tests/legacy-drive-ui.spec.mjs` が、ここを回帰テストしています。

## 触るときの注意

- **配信されていません。** `docs/` の下ではないので GitHub Pages は配りません。
- ここを直しても、誰の端末にも届きません。いまの本番はリポジトリ直下の `.gs` です。
- `offline.html` にはアプリへ戻るリンクがありません。当時からの積み残しで、
  配信から外したので直していません（いま配信している `docs/offline.html` には入っています）。
