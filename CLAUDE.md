# CLAUDE.md — このリポジトリを触るときに先に読むこと

@.agents/rules/gigaschool-standards.md

⚠️ **上の 1 行を消さないこと。** 艦隊共通のルール（Zero-CDN・Zero-PII・正本同期）は
正本 `GIGAyama.github.io/standards/agents/rules/` に 1 本だけ置いてある。
Claude Code はこの取りこみを通して読む。以下はこのリポジトリ固有の話。

## 何が本番か（2026-08-23 に入れ替わった）

**本番は、スプレッドシートにコンテナバインドした Google Apps Script です。**

| 場所 | 何か | 配信されるか |
| --- | --- | --- |
| リポジトリ直下の `.gs` / `index.html` / `app.html` ほか | アプリ本体 | **GitHub からは配信されない**（下記） |
| `docs/` | 導入案内のページ | GitHub Pages（`docs/CNAME`） |
| `legacy/drive-native/` | 退役した旧本番。配信していない | されない |

`src/app.jsx` / `tools/extra.css` / `tailwind.config.js` / `index.html` が原本で、
`app.html` / `css.html` / `vendor.html` / `qr.html` は **生成物**です。手で編集しないこと。
`npm run build` で作り直し、`scripts/check-generated.mjs` が食い違いを止めます。

## 反映経路（マージ＝反映ではない）

- `docs/` を変えた → main にマージで GitHub Pages に出ます。
  **配信物を 1 バイトでも変えたら `node tools/build-sw.mjs` で SW の版を上げること**
  （定数名は `sw-build.config.json` の `versionConst`。この repo は `APP_VERSION`）。
- `.gs` / `*.html` / `appsscript.json` を変えた → **自動では誰にも届きません。**
  先生方はスプレッドシートのコピーを持っていて、その中のコードで動いています。
  届けるには配布用テンプレートを更新し、貼り直しをお願いする必要があります
  （`docs/copy-distribution.md`）。「main にマージしたので直りました」と書かないこと。

## 触る前に

```bash
npm ci
npm run build && npm run check      # ビルド → 検査 の順（逆だと生成物を見ずに緑になる）
node scripts/selftest-check.mjs     # 検査をわざと壊して落ちることを見る
npm test
node /home/user/GIGAyama.github.io/standards/check-drift.mjs --standards /home/user/GIGAyama.github.io/standards
```

配布テンプレートのコピーリンクは **4 か所**にあります（案内ページ・README 2 か所・紹介記事）。
テンプレートを作り直して ID が変わったら 4 か所すべてを差し替えること。
**G8 が突き合わせていて、1 か所でも古い ID が残ると落ちます。**
古い ID を踏んだ先生は、前のテンプレートをコピーしてしまいます。

## この repo 固有の落とし穴

- **`.gs` の公開関数（末尾 `_` 無し）は `google.script.run` から誰でも呼べます。**
  新設したら必ず `assertOwner_()` か `guardMember_()` を通すこと。G1 が数えています。
- **先生が誰かは「登録」ではなく「デプロイした人」で決まります**（`resolveOwner_`）。
  根拠は `Session.getEffectiveUser()` です。`getActiveUser()`（＝開いた本人）に
  戻さないこと。戻すと、先生より先に URL を開いた児童が恒久的に先生になります。
  この前提は `appsscript.json` の `executeAs: USER_DEPLOYING` が支えています。
  そこが `USER_ACCESSING` だと `getEffectiveUser` は開いた本人を返し、全員が
  先生を名乗れるようになります。**両方まとめて G12 が見ています。**
- **設定作業をスプレッドシート側に戻さないこと。** メニューは「準備の状態を見る」
  「シートを点検する」「シートを直せる範囲で直す」の 3 つだけで、書き込む入口はありません。
  足すと「コピーしてデプロイするだけ」という案内と実物がずれます。G13 が見ています。
- **コピーは `fileId` で見分けます。** 控えた時点のファイル ID と、いま開いている
  ファイルの ID が違えば、それはコピーです（`resolveOwner_`）。そのとき `ownerEmail` は
  自動で貼り替わりますが、**名簿と提出は消しません。** 消すのは先生が
  ［準備の状態］で押したときだけです（`opClearInheritedData`）。
  「コピーだから前の学級のもの」と決めつけると、去年の記録を持ち越したい先生の分まで消えます。
- **列は見出しの名前で引きます**（`headerMap_` / `colOf_`）。列番号の直書きに戻さないこと。
  先生が 1 列挿しただけで「返却が別の列に入る」が、画面に何も出ないまま起きます。G6 が見ています。
- **点検（`inspectSheets_`）は 1 セルも書き換えません。** 見出しだけ正しくすると、
  ずれた列に正しいラベルが付いて事故が見えなくなります。G7 が見ています。
- **修整（`repairSheets_`）は右端に足すだけ。** 既にある列を動かさず、名前も変えず、何も消しません。
- **差し込みは `getRawContent()` で。`createHtmlOutputFromFile(...).getContent()` を使わないこと。**
  後者は中身を **HTML として読み直し、組み立て直して** 返します。app.html の中身は
  `<script>` 1 個ぶんの JavaScript で、その中には HTML の断片を組み立てる文字列があります。
  読み直されるとその断片が本物のタグとして扱われ、**バッククォートの対応が崩れた**
  JavaScript が返ってきます。これで「タブに題は出るが画面が出ない」状態になりました
  （2026-08-24）。G11 が見ています。
- **ブラウザのエラー行は、文書の行とはかぎりません。**
  上の事故で出た `userCodeAppPanel…:1314` は、貼り合わせた 1 枚の 1314 行目ではなく、
  **`app.html` の 1315 行目**でした（`<script>` の中身を 1 行目として数えるため、1 引く）。
  文書の行として読んだせいで、当てずっぽうの修正を 2 回本番へ出しました。
  まず `app.html` 単体の行として読み、合わなければ貼り合わせた側を疑う順にすること。
- **手元で再現しない事故は、手元に無い工程を疑うこと。**
  上の事故のとき app.html 単体は `node --check` を通り、`scripts/assemble-gas-page.mjs` で
  貼り合わせた 1 枚も本物のブラウザで動きました。手元に無かったのは
  「GAS が中身を読み直す」工程そのもので、そこが原因でした。
  `tests/gas-page.spec.mjs` が、組み立てた 1 枚を本物のブラウザに読ませています。
- **`<?` を、GAS が差し込むファイル（app/css/vendor/qr.html）に入れないこと。**（G10）
  こちらは上の事故の**原因ではありませんでした**（消しても直らなかった）。
  それでも `<?` はスクリプトレットの開き記号なので、念のため止めています。
- **GAS の列挙子は、無い名前を書いても JavaScript が黙って undefined を返します。**
  `HtmlService.XFrameOptionsMode` は `ALLOWALL` と `DEFAULT` の 2 つだけ。
  `SAMEORIGIN` と書いたら doGet が「引数は null にできません: mode」で落ち、
  **画面が 1 つも開かなくなりました**（2026-08-23）。G9 が名前の表と突き合わせています。
- **テストの偽物を本物より寛容にしないこと。** 上の事故のとき、偽の HtmlService は
  `evaluate: () => ({})` で、`setXFrameOptionsMode` を呼びもしませんでした。
  テストは 23 件すべて緑のまま、開かない画面を通しました。
  `scripts/test-gas.mjs` の `makeHtmlService` は、GAS と同じように undefined で落ちます。
- `Gemini.gs` / `scripts/gas-deploy.mjs` / `.github/workflows/deploy.yml` /
  `tools/build-sw.mjs` / `tools/check-secrets.mjs` は**正本のコピー**です
  （`standards-map.json`）。直すときは `GIGAyama.github.io/standards/` を直してから配ること。

## 測っていないこと

- 本番疎通・実機での動作は、この環境からは確かめられません（外部通信が塞がれています）。
- ブラウザ実測は案内ページ（`docs/`）だけです。GAS 本体の画面は実測していません。
