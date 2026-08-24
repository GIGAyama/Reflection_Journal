# 🚚 ロールアウト記録 — Reflection_Journal

GIGA Standard v5 の改修モード（`/rollout`）の作業記録。
**他リポジトリにも効く知見は、その1本の問題で終わらせずにここへ書く。**

- 型: **C+型**（GitHub Pages シェル + GAS 2デプロイ）
- 監査時のコミット: `b781dbd`
- 監査結果と実測値: [AUDIT.md](AUDIT.md)

## 2026-08-23: コンテナバインド配布へ作り直した

配布の形を「先生ごとにデプロイ」へ変えた。共通 URL は無くなり、
`reflection-journal.giga-school.com` は導入案内のページになった。

| 変えたところ | 前 | 後 |
| --- | --- | --- |
| アプリの形 | GitHub Pages + Drive REST API（Drive ネイティブ版） | スプレッドシートにコンテナバインドした GAS |
| 配り方 | 共通 URL 1 本 | スプレッドシートのコピー → 先生ごとにデプロイ |
| 本人確認 | GIS の ID トークン検証（運営者の OAuth クライアント） | `Session.getActiveUser()`（同じ組織内の全員でデプロイ） |
| 記録の置き場所 | 児童と先生それぞれの Drive に分散 | コピーしたスプレッドシート 1 つ |
| 権限 | `drive.file` + `drive.readonly`（制限付きスコープ・審査が要る） | `spreadsheets.currentonly` ほか 3 つ |
| `docs/` | アプリ本体（PWA） | 導入案内のページ |
| 品質ゲート | Drive ネイティブ版を見ていた（`.gs` は対象外） | GAS 本体と案内ページを見る（G1〜G8 を新設） |

### 他リポジトリにも効く知見

**「配布の形を変える」は、検査の向き先を変える作業でもある。**
`scripts/check-project.mjs` は `gsFiles = []`（「ルート直下の旧GAS資産は品質判定に含めない」）と
書いてあり、コメントは正しかった。ただしその前提が入れ替わった瞬間、
**本番になったコードを 1 行も見ないゲートが緑を返し続ける状態**になる。
配信物の場所を動かしたら、`check-project.mjs` の対象配列と `selftest-check.mjs` の
「こわしかた」を同じ PR で付け替えること。

**GAS 側の検査を足したら、その場で 1 件見つかった。** `getDb()` が末尾 `_` 無しで、
`google.script.run` から誰でも呼べる状態だった（G1）。
正本ゲートの 38 項目に GAS 関連は 1 つも無いので、ここは自前で持つしかない。

**GAS の `<?` は、差し込まれる側のファイルにも効く。**
QR の SVG に `<?xml version="1.0" encoding="UTF-8"?>` を書いていた。
doGet は index.html をテンプレートとして評価し、`<?!= include('app') ?>` で
app/css/vendor/qr を差し込むので、ブラウザが受け取るのは貼り合わせた 1 枚である。
その 1 枚の中で `<?` はスクリプトレットの開き記号になる。
**app.html にあった唯一の `<?` と唯一の `?>` がこれで、貼り合わせた側だけが壊れた。**
症状は「タブに題は出るが画面が出ない」＋
`Uncaught SyntaxError: Unexpected identifier 'ふりかえりジャーナル_$'`。
テンプレート文字列の開きのバッククォートごと食われると、次のバッククォートまでが
文字列になり、その先が code 位置に出てくる、という壊れ方である。
`include()` を使う GAS アプリすべてに同じ形がある。G10 で名指しして止めた。

**ファイル単位の検査では、貼り合わせた 1 枚が動くことは一度も見ていない。**
このとき app.html 単体は `node --check` も通り、8 つのスクリプトブロックも全部妥当で、
テスト 85 件・ゲート 27 項目・CI 3 ジョブがすべて緑だった。
`scripts/assemble-gas-page.mjs` で doGet と同じ貼り合わせを再現し、
`tests/gas-page.spec.mjs` が本物のブラウザに読ませる形にした。
**組み立てたものを一度もブラウザに読ませていないなら、画面が出ることは誰も見ていない。**

**貼り合わせを `String.replace(文字列, 文字列)` で書かないこと。**
置換文字列の中の `$&` や `$'` が特殊記号として解釈される。
React の minify 済みコードには実際に `$'` が入っており、
差し込んだはずの中身が別のものに化けた。置換は関数で渡す。

**GAS の列挙子は、無い名前を書いても静かに undefined になる。**
`HtmlService.XFrameOptionsMode.SAMEORIGIN` と書いた（実在するのは `ALLOWALL` と `DEFAULT` だけ）。
`setXFrameOptionsMode(undefined)` は「引数は null にできません: mode」で落ちるので、
**doGet がまるごと死に、画面が 1 つも開かない**。ゲートも 26 項目すべて緑、
テストも 23 件すべて緑、CI も 3 ジョブすべて緑のまま本番に出た。
名前の表と突き合わせる検査（G9）を足した。ほかの repo でも
`ContentService.MimeType` / `Utilities.DigestAlgorithm` / `ui.ButtonSet` は同じ形で壊れる。

**テストの偽物を、本物より寛容にしないこと。** これがいちばんの教訓。
当時の偽 HtmlService は `createTemplateFromFile: () => ({ evaluate: () => ({}) })` で、
`setXFrameOptionsMode` を**呼びもしなかった**。だから壊れた行はテストに一度も触れられず、
23 件の緑は「doGet が動く」ことを 1 ミリも保証していなかった。
偽物は、本物が落ちるところで落ちる形にする（undefined を黙って受けない）。

**`onOpen` のメニュー関数も公開エンドポイントである。** `SpreadsheetApp.getUi()` を
関数の**先頭**で取ることが、そのまま「画面が無い文脈では 1 セルも書かない」保証になる。
`setupAsTeacher` は先生を決める関数なので、ここが逆順だと最初に開いた児童が先生になる。

## やったこと

| フェーズ | 内容 | 状態 |
| --- | --- | --- |
| Phase 0 | 実測（静的 + 実ブラウザ6画面 + PWA挙動） | ✅ `AUDIT.md` |
| P0 | CI ワークフロー（`pull_request` と `push` の両方） | ✅ |
| P0.5 | ブラウザ内 Babel / CDN 依存の解消 | ✅ 3,153.9KB → 271.0KB |
| P1 | ふりがな・拡大禁止・色・タップ・動き・a11y | ✅ 199件/101件 → 0件/0件 |
| P2 | 画像・maskable | ✅ アイコン差し替えに伴い作り直し |
| P4 | 品質ゲート（わざと壊して確認） | ✅ 9項目 |
| P1-f | CSP | ⏸ GAS 固有の事情で見送り（下記） |

## 他リポジトリにも効く知見

### 1. `context.setOffline()` は Service Worker の fetch に効かない 🆕

v5 §7-5 は「圏外で起動するか」を `context.setOffline(true)` で測ると書いているが、
**Chromium ではこれが Service Worker 内の `fetch()` に効かない。**
実測で `fromServiceWorker: true` かつ `status: 200` で本物の応答が返ってきた。
つまり「オフラインでも起動した」という結果は、**まったく当てにならない。**

`offline.html` が出るはずの条件（本体キャッシュを消して圏外）でも `index.html` が出てしまい、
最初は「offline.html が出ない不具合」だと誤診した。

**確かめ方: HTTP サーバーのプロセスを実際に落とす。**
落としてから測り直したところ、①圏外でシェルが起動し、②本体キャッシュを消すと
`offline.html` が出ることを、どちらも応答の `title` で確認できた。

### 2. maskable の「セーフゾーン外の中身」は、下地を数えると 130倍ずれる

§3-7 は「下地と中身を色で区別する」と書いているが、実際にどれだけずれるかの例として。
四隅の色との差だけで判定したところ **25.6%**（＝深刻に見える）。
下地のオレンジ系を除外して数え直すと **0.000%** だった。既存アイコンは元から問題なかった。

**下地がグラデーションだと、単純な「四隅の色との差」は下地そのものを中身として数える。**
色域（この例ではオレンジ系）で除外するのが確実。

### 3. パレット PNG は「いちばん軽い版」を選んではいけない 🆕

§2-6 の「いちばん軽くなった版を選ぶ」を素直に実装すると、常に最少色数（64色）が選ばれる。
その結果、**アイコンの葉の緑が茶色に潰れて別の絵になった。**

正しくは「**予算（512は60KB、faviconは30KB）に収まる中で、いちばん色数の多い版**」を選ぶ。
実測: 256色で 44.3KB。予算内なので画質を落とす理由がない。

### 4. `rt` の色決め打ちは、この構成では 1.08 まで悪化する

v5 §4 の実例は 1.28 / 1.47 だが、**グラデーションの帯の上では 1.08** まで落ちた。
Tailwind の `bg-gradient-to-r from-blue-500 to-indigo-500` のような面に
`text-gray-500` のふりがなが乗ると、ほぼ見えない。

横断で数えるには:

```bash
grep -rn '<rt[^>]*text-\(gray\|slate\|zinc\|neutral\)-' $(git ls-files '*.html' '*.jsx')
grep -rn '^\s*rt\s*{' $(git ls-files '*.html' '*.css')
```

### 5. アイコンコンポーネントの既定 className は「上書き」される 🆕

```jsx
const Icon = ({ path, className = "w-5 h-5" }) => <svg className={className} .../>;
```

この形は、呼び出し側が `className="text-indigo-500"` のように**大きさ以外だけを渡すと
既定値ごと消える**。SVG は親いっぱいに広がり、時計アイコンが 200px 超になっていた。
React の既定引数は「未指定のときだけ」効くので、気づきにくい。

対処は「大きさの指定が無いときだけ足す」:

```jsx
const hasSize = /(^|\s)(w-|h-|size-)/.test(className);
className={`${hasSize ? '' : 'w-5 h-5 '}${className}`.trim()}
```

**同じ形は他リポジトリにもありうる。** 探し方:

```bash
grep -rn 'className = "w-\|className="w-5 h-5" }' $(git ls-files '*.jsx' '*.html')
```

### 6. 品質ゲートの「ふりがな検査」は、受け皿を1つ外しても落ちない

§P4 の「わざと壊して通ることを確認する」を実行して見つかった4件目の不具合。
`rt { color: inherit }` の**存在だけ**を見る実装だと、
`[class*="text-white"] rt` のような受け皿を1つ削っても通ってしまう。

§4 は「1か所ずつ潰さない。まとめて継がせるのが正しい」と書いている。
検査も同じで、**必要な受け皿がそろっているか**を見る必要がある。

### 7. GAS には CSP を入れられない（`script-src 'self'` は全部止める）

GAS は `include()` で本体を `<script>` として差し込む構成であり、
`google.script.run` の橋渡しもインラインで入る。
`script-src 'self'` を入れると**アプリ全体が動かなくなる**。
`'unsafe-inline'` を足して回避するのは、CSP を入れた意味をほとんど無くす。

**C型 / C+型では、CSP は「入れない」を明示して理由を残すのが正しい。**
静的型（A型・B型）とは扱いを分けること。

### 8. C+型でも §6 の手順はそのまま使える

`vendor.html` / `css.html` / `app.html` を作り、`Main.gs` に `include()` を足して
`index.html` から貼り合わせるだけ。実測 3,153.9KB → 271.0KB。

ただし1点だけ注意: **JSX の文字列リテラルに `</script>` が入っていると、
生成した `app.html` がそこで切れる。** 本アプリは印刷帳票を
`w.document.write('…<script>…<\/script>…')` で組み立てており、これに該当した。
ビルド側で `</script` → `<\/script` に割ってから包む（文字列リテラルの中では等価）。

## 横断で回すべきもの

```bash
# ブラウザ内 Babel（v5 Part V の「残り10本」）
grep -rln "babel/standalone" $(git ls-files '*.html')

# ふりがなの色
grep -rn '<rt[^>]*text-\(gray\|slate\)-' $(git ls-files '*.html' '*.jsx')

# アイコンの既定 className 上書き
grep -rn 'className = "w-5 h-5"' $(git ls-files '*.html' '*.jsx')

# 拡大禁止（.gs 側も見ること）
grep -rn "user-scalable=no\|maximum-scale" $(git ls-files '*.html' '*.gs')
```

## 未了・人間の判断待ち

- **本番でのテストデプロイ**（作業環境から `script.google.com` へ到達できない）
- **マージ判断**（ビルド導入はアーキテクチャ変更のため）
- Tab 順・読み上げ順の計測（F3）
- GAS 本体の fluid type（D4）
- 帳票（別ウィンドウ）側の `@page`（D13）
