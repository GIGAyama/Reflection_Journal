# CLAUDE.md — このリポジトリを触るときに先に読むこと

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

`npm run check` の G8 は「配布テンプレートの ID が未設定」の**注意**です（落ちません）。
テンプレートを作ったら `docs/copy-distribution.md` §4 の 3 か所を差し替えて消してください。

## この repo 固有の落とし穴

- **`.gs` の公開関数（末尾 `_` 無し）は `google.script.run` から誰でも呼べます。**
  新設したら必ず `assertOwner_()` か `guardMember_()` を通すこと。G1 が数えています。
- **列は見出しの名前で引きます**（`headerMap_` / `colOf_`）。列番号の直書きに戻さないこと。
  先生が 1 列挿しただけで「返却が別の列に入る」が、画面に何も出ないまま起きます。G6 が見ています。
- **点検（`inspectSheets_`）は 1 セルも書き換えません。** 見出しだけ正しくすると、
  ずれた列に正しいラベルが付いて事故が見えなくなります。G7 が見ています。
- **修整（`repairSheets_`）は右端に足すだけ。** 既にある列を動かさず、名前も変えず、何も消しません。
- **`<?` を、GAS が差し込むファイル（app/css/vendor/qr.html）に入れないこと。**
  doGet は index.html をテンプレートとして評価し、`<?!= include('app') ?>` で
  4 つを差し込みます。**ブラウザが受け取るのは貼り合わせた 1 枚**で、`<?` は
  スクリプトレットの開き記号です。QR の SVG に XML 宣言を書いていたせいで、
  貼り合わせた側だけが壊れ「タブに題は出るが画面が出ない」状態になりました（2026-08-24）。
  G10 が見ています。
- **ファイル単位の検査では、貼り合わせた 1 枚が動くことは分かりません。**
  上の事故のとき app.html 単体は `node --check` も通り、テストも 85 件すべて緑でした。
  `tests/gas-page.spec.mjs` が、組み立てた 1 枚を本物のブラウザに読ませています。
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
