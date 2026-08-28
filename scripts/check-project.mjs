/**
 * 品質ゲート（GIGA Standard v5 Part I）+ このリポジトリ固有の GAS 検査。
 *
 * ── 何が本番なのかが変わった ────────────────────────────────────
 * 以前は docs/ の Drive ネイティブ版が本番で、ルート直下の .gs は履歴資料だった。
 * いまは逆で、**本番はスプレッドシートにコンテナバインドした .gs とその画面**、
 * docs/ は導入案内のページだけである。だから検査の対象もそちらへ移した。
 *
 * 正本ゲートの 38 項目に GAS 関連は 1 つも無い（standards/docs/app-hardening.md §5）。
 * .gs は CI に一切守られていない領域なので、ここで自前の G* を持つ。
 *
 * ⚠️ 検査を足したら、同じ PR で scripts/selftest-check.mjs の「こわしかた」も足すこと。
 *    反応しない検査は、緑を返すだけの飾りになる。
 */
import { readFileSync, statSync, existsSync, readdirSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const read = (p) => readFileSync(join(ROOT, p), 'utf8');
const size = (p) => statSync(join(ROOT, p)).size;
const has = (p) => existsSync(join(ROOT, p));
const problems = [];
const notes = [];
const fail = (id, msg) => problems.push(`${id}: ${msg}`);
const ok = [];
const pass = (id, msg) => ok.push(`${id}: ${msg}`);

/** JS/CSS のコメントを落とす。注意書きに反応する誤検知を止めるため */
const stripComments = (s) =>
  s.replace(/\/\*[\s\S]*?\*\//g, ' ').replace(/(^|[^:])\/\/[^\n]*/g, '$1 ');

// GitHub Pages が配る導入案内のページ
const siteHtml = ['docs/index.html', 'docs/offline.html', 'docs/privacy.html', 'docs/terms.html'];
// GAS が配る画面（外枠は原本、app/css/vendor/qr は生成物）
const gasHtml = ['index.html'];
// index.html が差し込んでいるファイルを、index.html 自身から読む。
// ⚠️ 決め打ちの一覧にしない。2026-08-28 に fonts.html を足したとき、
//    一覧が ['app','css','vendor','qr'] のままだったので G10 は
//    「4 ファイルに <? なし」と緑を出しつづけた。**足したファイルだけが
//    検査されない**という、いちばん気づけない形になる。
const gasGenerated = (() => {
  const src = existsSync(join(ROOT, 'index.html'))
    ? readFileSync(join(ROOT, 'index.html'), 'utf8')
    : '';
  const found = [...src.matchAll(/include_?\(\s*'([^']+)'\s*\)/g)].map((m) => m[1] + '.html');
  // index.html から読めなかったときのために、もとの 4 つを下限として残す
  return [...new Set([...found, 'app.html', 'css.html', 'vendor.html', 'qr.html'])].sort();
})();
// 児童の端末が最初に読み込む JS（GAS 側）。qr.html は先生の画面だけなので含めない
const memberJs = ['vendor.html', 'app.html'];
// サーバー側（.gs）。ここが CI に守られていない領域
const gsFiles = readdirSync(ROOT).filter((f) => f.endsWith('.gs')).sort();

// ── B6: CDN から取る実行コードが 0 バイト ──
{
  const bad = [];
  for (const f of [...siteHtml, ...gasHtml]) {
    const html = read(f);
    for (const m of html.matchAll(/<script[^>]+src=["'](https?:\/\/[^"']+)["']/gi)) bad.push(`${f} → ${m[1]}`);
    // ブラウザ内で CSS を生成する Tailwind CDN も実行コード扱い
    const noComments = html.replace(/<!--[\s\S]*?-->/g, '');
    if (/cdn\.tailwindcss\.com/.test(noComments)) bad.push(`${f} → cdn.tailwindcss.com`);
    if (/@babel\/standalone/.test(noComments)) bad.push(`${f} → @babel/standalone`);
  }
  bad.length ? fail('B6', `CDN から実行コードを読んでいる: ${bad.join(' / ')}`)
             : pass('B6', 'CDN から取る実行コード 0 件');
}

// ── B8: QRコードはローカル生成し、児童用URLを第三者へ送らない ──
{
  const source = read('src/app.jsx') + read('index.html') + read('app.html');
  const shell = read('index.html');
  if (/api\.qrserver\.com|chart\.googleapis\.com[^"']*cht=qr/i.test(source)) {
    fail('B8', 'QR生成のため児童用URLを外部サービスへ送っている');
  } else if (!has('qr.html')) {
    fail('B8', 'ローカルQR生成の生成物 qr.html が無い（npm run build を実行）');
  } else if (!/bootMode\s*===\s*'owner'[\s\S]{0,120}include_\('qr'\)/.test(shell)) {
    fail('B8', 'QR生成コードが先生の画面だけの読み込みになっていない（児童の初回JSに混ざる）');
  } else {
    pass('B8', 'QRはブラウザ内で生成し、生成コードは先生の画面だけに配信');
  }
}

// ── D14 / D1: 拡大禁止と viewport-fit（.gs 側の addMetaTag も見る） ──
{
  const targets = [...siteHtml, ...gasHtml, 'Main.gs'];
  const zoomBlocked = targets.filter((f) => /user-scalable\s*=\s*no|maximum-scale\s*=\s*1/.test(read(f)));
  zoomBlocked.length ? fail('D14', `拡大を禁止している: ${zoomBlocked.join(', ')}`)
                     : pass('D14', '拡大を禁止していない');

  const viewportDecls = [];
  for (const f of targets) {
    const s = read(f);
    for (const m of s.matchAll(/(?:name=["']viewport["']\s+content\s*=\s*|addMetaTag\('viewport',\s*)["']([^"']+)["']/g)) {
      viewportDecls.push({ file: f, value: m[1] });
    }
  }
  const missing = viewportDecls.filter((v) => !/viewport-fit\s*=\s*cover/.test(v.value));
  missing.length ? fail('D1', `viewport-fit=cover が無い: ${missing.map((m) => m.file).join(', ')}`)
                 : pass('D1', `viewport 宣言 ${viewportDecls.length} 件すべてに viewport-fit=cover`);
}

// ── D2: 100vh の単独使用（@supports のフォールバックは正しい書き方なので見逃す） ──
{
  const bad = [];
  for (const f of [...siteHtml, 'docs/style.css', 'tools/extra.css', 'css.html']) {
    if (!has(f)) continue;
    const s = stripComments(read(f));
    for (const m of s.matchAll(/100vh/g)) {
      const before = s.slice(Math.max(0, m.index - 400), m.index);
      if (/@supports\s+not\s*\(height:\s*100dvh\)/.test(before)) continue;   // 前方も見る
      if (/100dvh/.test(s)) continue;                                        // dvh を併記している
      bad.push(f);
      break;
    }
  }
  bad.length ? fail('D2', `100vh を単独で使っている: ${[...new Set(bad)].join(', ')}`)
             : pass('D2', '100vh の単独使用なし');
}

// ── F4: rt（ふりがな）の色の決め打ち ──
{
  const css = read('tools/extra.css');
  const jsx = read('src/app.jsx');
  const flat = css.replace(/\s+/g, ' ');
  const hardCodedInJsx = /<rt[^>]*className="[^"]*text-(gray|slate|zinc|neutral)-\d/.test(jsx);

  // 色のついた面で親を継がせているか。1か所ずつ潰すと必ず漏れるので、
  // 「まとめて継がせているか」を、必要な受け皿がそろっているかで見る。
  const inheritRules = [...flat.matchAll(/([^{}]+)\{\s*color:\s*inherit[^}]*\}/g)]
    .map((m) => m[1]).filter((sel) => /\brt\b/.test(sel)).join(',');
  const REQUIRED = ['button rt', 'a rt', '[class*="bg-"] rt', '[class*="text-white"] rt'];
  const missing = REQUIRED.filter((sel) => !inheritRules.includes(sel));

  if (hardCodedInJsx) fail('F4', 'rt の色をマークアップ側で決め打ちしている（色のついた面の上で読めなくなる）');
  else if (!inheritRules) fail('F4', 'rt の色に、色のついた面で親を継がせる指定（color: inherit）が無い');
  else if (missing.length) fail('F4', `rt を継がせる受け皿が足りない: ${missing.join(' / ')}（1か所ずつ潰すと必ず漏れる）`);
  else pass('F4', 'rt の色は決め打ちしていない（色のついた面ではまとめて継がせている）');
}

// ── D10 / D11: 動きの配慮とハイコントラスト（案内ページと GAS 本体の両方） ──
{
  for (const [label, files] of [['案内ページ', ['docs/style.css']], ['GAS 本体', ['tools/extra.css']]]) {
    const css = files.map(read).join('\n');
    if (!/prefers-reduced-motion/.test(css)) fail('D10', `${label} に prefers-reduced-motion が無い`);
    if (/animation-duration:\s*0\s*(!important)?\s*;/.test(css)) {
      fail('D10', `${label} の animation-duration が 0（.01ms にしないと fill-mode: forwards が壊れて中身が消える）`);
    }
    if (!/forced-colors/.test(css)) fail('D11', `${label} に forced-colors が無い`);
  }
  if (!problems.some((p) => p.startsWith('D10'))) pass('D10', 'prefers-reduced-motion あり（案内ページ・GAS 本体とも）');
  if (!problems.some((p) => p.startsWith('D11'))) pass('D11', 'forced-colors あり（案内ページ・GAS 本体とも）');
}

// ── E5 / E6 / E7: sw.js ──
{
  const swRaw = read('docs/sw.js');
  const sw = stripComments(swRaw);
  // 「消す式」ではなく「startsWith で絞る式があるか」を見る
  if (/caches\.keys\(\)/.test(sw) && !/startsWith\s*\(/.test(sw)) {
    fail('E5', 'sw.js が caches.keys() を接頭辞で絞らずに扱っている（同一オリジンの他アプリを巻き添えにする）');
  } else pass('E5', 'sw.js は自アプリ接頭辞のキャッシュだけを扱っている');

  /localStorage/.test(sw) ? fail('E6', 'sw.js が localStorage に触れている') : pass('E6', 'sw.js は localStorage に触れていない');

  const installBlock = sw.slice(sw.indexOf("addEventListener('install'"), sw.indexOf("addEventListener('activate'"));
  /skipWaiting\s*\(/.test(installBlock)
    ? fail('E7', 'install の中で skipWaiting している（読んでいる最中に画面が入れ替わる）')
    : pass('E7', 'install では skipWaiting していない');
}

// ── E1: manifest の id / scope / start_url ──
{
  const m = JSON.parse(read('docs/manifest.webmanifest'));
  // 独自ドメイン reflection-journal.giga-school.com の直下で配信されるので "/"。
  // 旧構成（gigayama.github.io/Reflection_Journal/）のリポジトリ名の絶対パスに
  // 戻すと、scope がページの URL を含まなくなって manifest ごと無視され、
  // PWA としてインストールできなくなる。実際にその状態で残っていた。
  const want = '/';
  const bad = ['id', 'start_url', 'scope'].filter((k) => m[k] !== want);
  bad.length ? fail('E1', `manifest の ${bad.join(', ')} が配信場所（${want}）と合っていない`)
             : pass('E1', 'manifest の id / scope / start_url が配信場所と合っている');
}

// ── E2: apple-touch-icon に透明を含まない（PNG のヘッダで判定） ──
{
  const html = read('docs/index.html');
  const m = html.match(/rel="apple-touch-icon"[^>]*href="\.\/([^"]+)"/);
  if (!m) fail('E2', 'apple-touch-icon の指定が無い');
  else {
    const buf = readFileSync(join(ROOT, 'docs', m[1]));
    const colourType = buf[25];                       // IHDR の colour type
    const hasAlphaChannel = colourType === 4 || colourType === 6;
    const hasTRNS = buf.includes(Buffer.from('tRNS'));
    (hasAlphaChannel || hasTRNS)
      ? fail('E2', `${m[1]} が透明を持っている（iOS は透明を黒で埋めるため四隅が黒くなる）`)
      : pass('E2', `${m[1]} は透明を持っていない`);
  }
}

// ── E8: offline.html は JavaScript に頼らず、必ず戻るリンクを持つ ──
{
  const f = 'docs/offline.html';
  const html = read(f).replace(/<!--[\s\S]*?-->/g, '');   // 「使わない理由」のコメント本文は除く
  const bad = [];
  if (/<script/i.test(html)) bad.push('<script> がある');
  if (/\son[a-z]+\s*=/i.test(html)) bad.push('onclick= 等がある');
  if (!/href="\.\/"/.test(html)) bad.push('<a href="./"> の戻るリンクが無い');
  bad.length
    ? fail('E8', `offline.html: ${bad.join(' / ')}（本体が読めていない状況で JS に頼ると、その JS も読めない）`)
    : pass('E8', 'offline.html は JS を使わず、戻るリンクがある');
}

// ── C5: localStorage.clear() ──
{
  const files = [...siteHtml, 'src/app.jsx'].filter(has);
  const bad = files.filter((f) => /localStorage\.clear\s*\(/.test(stripComments(read(f))));
  bad.length ? fail('C5', `localStorage.clear() を使っている: ${bad.join(', ')}（同一オリジンの他アプリの保存も消える）`)
             : pass('C5', 'localStorage.clear() を使っていない');
}

// ── F5: 児童の初回 JS 300KB 以下（vendor + app。qr は先生の画面だけなので含めない） ──
{
  const kb = memberJs.reduce((s, f) => s + size(f), 0) / 1024;
  kb <= 300 ? pass('F5', `児童の初回 JS ${kb.toFixed(1)} KB（300KB 以下）`)
            : fail('F5', `児童の初回 JS が ${kb.toFixed(1)} KB（300KB を超えている）`);
}

// ── F6: 1ファイル 5,000行 / 400KB（生成物は除く） ──
{
  const bad = [];
  for (const f of [...siteHtml, ...gasHtml, 'src/app.jsx', 'docs/style.css', 'tools/extra.css', ...gsFiles]) {
    if (!has(f)) continue;
    const s = read(f);
    if (s.split('\n').length > 5000 || size(f) > 400 * 1024) bad.push(f);
  }
  bad.length ? fail('F6', `5,000行 / 400KB を超えるファイル: ${bad.join(', ')}`)
             : pass('F6', '5,000行 / 400KB を超えるファイルなし（生成物を除く）');
}

// ── B4: postMessage の宛先 ──
{
  const bad = [];
  for (const f of [...siteHtml, ...gasHtml, 'src/app.jsx']) {
    const s = stripComments(read(f));
    for (const m of s.matchAll(/postMessage\([^)]*,\s*['"]\*['"]\s*\)/g)) bad.push(`${f}: ${m[0].slice(0, 50)}`);
    // 変数経由でも、設定が空のときに '*' へ落ちる書き方は同じ危うさがある
    for (const m of s.matchAll(/(\w+)\s*=\s*[^;\n]*\|\|\s*['"]\*['"]/g)) {
      if (new RegExp(`postMessage\\([^)]*,\\s*${m[1]}\\s*\\)`).test(s)) bad.push(`${f}: ${m[0].slice(0, 60)}`);
    }
  }
  bad.length ? notes.push(`B4: postMessage の宛先に '*' を直接書いている箇所がある（${bad.join(' / ')}）`)
             : pass('B4', "postMessage の宛先に '*' の直書きなし");
}

// ── D7: 画像 ──
{
  const bad = [];
  for (const f of readdirSync(join(ROOT, 'docs')).filter((f) => f.endsWith('.png'))) {
    const kb = size(join('docs', f)) / 1024;
    const limit = /512/.test(f) ? 60 : /192|180/.test(f) ? 30 : 150;
    if (kb > limit) bad.push(`${f} ${kb.toFixed(1)}KB > ${limit}KB`);
  }
  bad.length ? fail('D7', `画像が上限を超えている: ${bad.join(', ')}`) : pass('D7', '画像はすべて上限内');
}

// ════════════════════════════════════════════════════════════════
// G*: GAS（.gs）— 正本ゲートの 38 項目に 1 つも無い領域
// ════════════════════════════════════════════════════════════════

/** .gs を「トップレベル関数の名前 → 本文」に割る */
function gasFunctions() {
  const out = [];
  for (const f of gsFiles) {
    const src = read(f);
    const re = /^function\s+([A-Za-z_$][\w$]*)\s*\(([^)]*)\)/gm;
    const hits = [...src.matchAll(re)];
    hits.forEach((m, i) => {
      const start = m.index;
      const end = i + 1 < hits.length ? hits[i + 1].index : src.length;
      out.push({ file: f, name: m[1], args: m[2], body: src.slice(start, end) });
    });
  }
  return out;
}

// ── G1: 公開エンドポイント（末尾 `_` 無し）はすべて認可を通る ──
// google.script.run は末尾 `_` の無い関数を誰でも直接呼べる。1 つの抜けで境界が破れる。
{
  // doGet / onOpen は入口。メニュー用の関数は先に getUi() を取ることで
  // 「画面が無い文脈（ウェブアプリ）では 1 行も進まない」ことを保証している。
  // 差し込み（include_）は末尾 `_` なので、そもそも公開されていない。
  const ENTRY = new Set(['doGet', 'onOpen']);
  const GUARDS = ['assertOwner_(', 'guardMember_(', 'requireEmail_(', 'SpreadsheetApp.getUi()'];
  const bad = [];
  for (const fn of gasFunctions()) {
    if (fn.name.endsWith('_') || ENTRY.has(fn.name)) continue;
    if (!GUARDS.some((g) => fn.body.includes(g))) bad.push(`${fn.file}:${fn.name}`);
  }
  bad.length
    ? fail('G1', `認可の無い公開関数がある（google.script.run から誰でも呼べる）: ${bad.join(', ')}`)
    : pass('G1', `公開関数はすべて認可を通っている（${gasFunctions().filter((f) => !f.name.endsWith('_')).length} 本）`);
}

// ── G2: 児童の書き込みはロックの中（40台が一斉に叩く前提） ──
{
  const WRITE_APIS = ['mbSaveJournal', 'mbAddPastComment', 'mbRequestJoin'];
  const fns = gasFunctions();
  const bad = WRITE_APIS.filter((name) => {
    const fn = fns.find((f) => f.name === name);
    return !fn || !fn.body.includes('withScriptLock_(');
  });
  bad.length
    ? fail('G2', `児童の書き込みが LockService で囲まれていない: ${bad.join(', ')}（同じ児童の行が2つ入り、誰も気づかない）`)
    : pass('G2', `児童の書き込みはすべてロックの中（${WRITE_APIS.length} 本）`);
}

// ── G3: 児童 API は画面から「誰か」を受け取らない ──
// 引数で受け取ると、児童がコンソールから他人のアドレスを渡すだけで成りすませる。
{
  const bad = gasFunctions()
    .filter((f) => f.file === 'MemberApi.gs' && !f.name.endsWith('_'))
    .filter((f) => /\bemail\b/i.test(f.args))
    .map((f) => `${f.name}(${f.args})`);
  bad.length
    ? fail('G3', `児童 API が画面からメールアドレスを受け取っている: ${bad.join(', ')}`)
    : pass('G3', '児童 API は画面から「誰か」を受け取っていない');
}

// ── G4: 児童へ返す一覧は必ずサニタイズを通る（ほかの児童のアドレスを出さない） ──
{
  const fn = gasFunctions().find((f) => f.name === 'mbSync');
  (fn && fn.body.includes('sanitizeJournals_('))
    ? pass('G4', '児童へ返す一覧はサニタイズを通している')
    : fail('G4', 'mbSync が sanitizeJournals_ を通していない（ほかの児童のメールアドレスが出る）');
}

// ── G5: appsscript.json のスコープ ──
{
  if (!has('appsscript.json')) {
    // clasp push は GAS 側のマニフェストを丸ごと上書きする。無いまま送ると入口が消える。
    fail('G5', 'appsscript.json が無い（clasp push でウェブアプリの入口が消える）');
  } else {
    const m = JSON.parse(read('appsscript.json'));
    const scopes = m.oauthScopes || [];
    const bad = [];
    // フル Drive は「子どものドライブ全部を読めます」と保護者に説明せざるを得なくなる
    if (scopes.some((s) => /auth\/drive$|auth\/drive\.readonly$/.test(s))) bad.push('ドライブ全体のスコープを要求している');
    // コンテナバインドなら、束ねられた 1 ファイルだけで足りる
    if (scopes.some((s) => /auth\/spreadsheets$/.test(s))) bad.push('spreadsheets（全スプレッドシート）を要求している。currentonly で足りる');
    if (!scopes.includes('https://www.googleapis.com/auth/spreadsheets.currentonly')) bad.push('spreadsheets.currentonly が無い');
    if (!m.webapp || !m.webapp.executeAs || !m.webapp.access) bad.push('webapp の executeAs / access が無い');
    if (m.webapp && m.webapp.access === 'ANYONE_ANONYMOUS') bad.push('access が ANYONE_ANONYMOUS（誰が書いたかを確かめられない）');
    bad.length ? fail('G5', `appsscript.json: ${bad.join(' / ')}`)
               : pass('G5', `oauthScopes は ${scopes.length} 本で、ドライブ全体を要求していない`);
  }
}

// ── G6: シートの列は見出しの名前で引く ──
// 列番号の直書きに戻すと、先生が誤字を直すついでに 1 列挿しただけで
// 「返却が別の列に入る」「自分の記録が出ない」が、画面に何も出ないまま起きる。
{
  const db = stripComments(read('Db.gs'));
  const bad = [];
  if (!/function headerMap_\(/.test(db)) bad.push('headerMap_ が無い');
  if (!/function colOf_\(/.test(db)) bad.push('colOf_ が無い');
  // getRosterRows_ が固定の列番号（r[0], r[1]…）に戻っていないか
  const roster = gasFunctions().find((f) => f.name === 'getRosterRows_');
  if (roster && /\br\[\d+\]/.test(stripComments(roster.body))) bad.push('getRosterRows_ が列番号の直書きに戻っている');
  bad.length ? fail('G6', `列の引き方: ${bad.join(' / ')}`)
             : pass('G6', 'シートの列は見出しの名前で引いている');
}

// ── G7: 点検（読むだけ）と修整（人が押したときだけ）が両方ある ──
{
  const names = new Set(gasFunctions().map((f) => f.name));
  const bad = ['inspectSheets_', 'repairSheets_', 'ensureSheets_'].filter((n) => !names.has(n));
  const inspect = gasFunctions().find((f) => f.name === 'inspectSheets_');
  if (bad.length) fail('G7', `シートの点検・修整が足りない: ${bad.join(', ')}`);
  // 点検は読むだけ。ここで書き換えると、ずれている列に正しいラベルが付いて事故が見えなくなる
  else if (inspect && /\.setValue\(|\.setValues\(|insertSheet\(|deleteRow/.test(stripComments(inspect.body))) {
    fail('G7', 'inspectSheets_ がシートを書き換えている（点検は読むだけにする）');
  } else pass('G7', '点検は読むだけ、修整は人が押したときだけ');
}

// ── G10: GAS が組み立てるファイルに `<?` を入れない ──
//
// doGet は index.html を「テンプレート」として評価し、`<?!= include('app') ?>` などで
// app / css / vendor / qr を差し込む。ブラウザが受け取るのは貼り合わせた 1 枚である。
// `<?` は GAS のスクリプトレットの開き記号なので、差し込まれる側に置いてはいけない。
//
// 2026-08-24、QR の SVG に XML 宣言（<?xml version="1.0" encoding="UTF-8"?>）が入っていた。
// これが app.html にある唯一の `<?` と唯一の `?>` で、貼り合わせた 1 枚だけが壊れ、
// 「タブに題は出るが画面が出ない」状態になった。ファイル単位の構文検査は全部通っていた。
{
  const bad = [];
  for (const f of gasGenerated) {
    if (!has(f)) continue;
    const src = read(f);
    const hits = [...src.matchAll(/<\?/g)];
    if (hits.length) {
      const line = src.slice(0, hits[0].index).split('\n').length;
      bad.push(`${f}:${line}（${hits.length} か所）`);
    }
  }
  bad.length
    ? fail('G10', `GAS が差し込むファイルに「<?」がある（スクリプトレットと解釈され、貼り合わせた1枚が壊れる）: ${bad.join(' / ')}`)
    : pass('G10', `GAS が差し込む ${gasGenerated.length} ファイルに「<?」なし`);
}

// ── G11: 差し込みは、中身を読み直さずそのまま渡す ──
//
// include は 2 通り書ける。見た目はほぼ同じで、片方だけが本番を壊す。
//
//   ❌ HtmlService.createHtmlOutputFromFile(name).getContent()
//        中身を **HTML として読み直し、組み立て直して** 返す。
//   ✅ HtmlService.createTemplateFromFile(name).getRawContent()
//        ファイルの中身をそのまま返す（解釈を挟まない）。
//
// app.html の中身は <script> 1 個ぶんの JavaScript で、その中には HTML の断片を
// 組み立てる文字列がある。読み直されるとその断片が本物のタグとして扱われ、
// バッククォートの対応が崩れた JavaScript が返ってくる。
//
// 2026-08-24、これで画面が出なくなった。ブラウザは app.html の 1315 行目
// （QR の保存名 `ふりかえりジャーナル_${...}_QR.svg`）を構文エラーとして指した。
// ファイル単位の検査も、手元で貼り合わせた 1 枚も、すべて緑のままだった。
// 手元には「読み直し」が無いので、手元では再現しない種類の事故である。
{
  const src = has('Main.gs') ? stripComments(read('Main.gs')) : '';
  const bad = [];
  if (/createHtmlOutputFromFile\s*\(/.test(src)) {
    bad.push('createHtmlOutputFromFile(...).getContent() を使っている（中身が読み直される）');
  }
  if (!/getRawContent\s*\(\)/.test(src)) {
    bad.push('getRawContent() で差し込んでいない');
  }
  bad.length
    ? fail('G11', `差し込みが中身を読み直している: ${bad.join(' / ')}`)
    : pass('G11', '差し込みは getRawContent（中身を読み直さない）');
}

// ── G9: GAS の列挙子に、実在しない名前を書いていない ──
//
// 2026-08-23、doGet が HtmlService.XFrameOptionsMode.SAMEORIGIN を渡していた。
// この名前は無い（ALLOWALL と DEFAULT の 2 つだけ）ので undefined になり、GAS は
// 「引数は null にできません: mode」で落ちた。**画面が 1 つも開かない**状態である。
// JavaScript は存在しないプロパティを undefined として黙って通すので、
// 目でも動かしても気づけない。名前の表と突き合わせるしかない。
{
  const GAS_ENUMS = {
    'HtmlService.XFrameOptionsMode': ['ALLOWALL', 'DEFAULT'],
    'HtmlService.SandboxMode': ['IFRAME'],
    'ContentService.MimeType': ['ATOM', 'CSV', 'ICAL', 'JAVASCRIPT', 'JSON', 'RSS', 'TEXT', 'VCARD', 'XML'],
    'Utilities.DigestAlgorithm': ['MD2', 'MD5', 'SHA_1', 'SHA_256', 'SHA_384', 'SHA_512'],
    'Utilities.Charset': ['US_ASCII', 'UTF_8'],
    // メニューのダイアログ。ui は SpreadsheetApp.getUi() の戻り値なので接頭辞が違う
    'ui.ButtonSet': ['OK', 'OK_CANCEL', 'YES_NO', 'YES_NO_CANCEL'],
    'ui.Button': ['CLOSE', 'OK', 'CANCEL', 'YES', 'NO']
  };
  const bad = [];
  let checked = 0;
  for (const f of gsFiles) {
    const src = stripComments(read(f));
    for (const [path, members] of Object.entries(GAS_ENUMS)) {
      const re = new RegExp(`\\b${path.replace('.', '\\.')}\\.([A-Za-z0-9_]+)`, 'g');
      for (const m of src.matchAll(re)) {
        checked++;
        if (!members.includes(m[1])) {
          bad.push(`${f}: ${path}.${m[1]}（実在するのは ${members.join(' / ')}）`);
        }
      }
    }
  }
  bad.length
    ? fail('G9', `実在しない GAS の列挙子を使っている（undefined が渡り、その画面が開かなくなる）: ${bad.join(' / ')}`)
    : pass('G9', `GAS の列挙子 ${checked} か所はすべて実在する名前`);
}

// ── G12: 先生を「開いた本人」から導いていないこと ──
//
// resolveOwner_ が Session.getActiveUser を根拠にすると、先生より先に URL を開いた
// 児童が恒久的に先生になる。根拠は必ず getEffectiveUser（＝デプロイした人）である。
// 併せて、マニフェストの executeAs が USER_DEPLOYING であることも見る。
// そこが USER_ACCESSING だと getEffectiveUser は開いた本人を返し、前提ごと崩れる。
{
  const main = read('Main.gs');
  const body = main.slice(main.indexOf('function resolveOwner_'), main.indexOf('function classNameOf_'));
  const problems = [];
  if (!body.includes('effectiveEmail_()')) {
    problems.push('resolveOwner_ が getEffectiveUser を根拠にしていない');
  }
  if (/=\s*activeEmail_\(\)/.test(body)) {
    problems.push('resolveOwner_ が activeEmail_（開いた本人）を根拠にしている');
  }
  let manifest = {};
  try { manifest = JSON.parse(read('appsscript.json')); } catch (e) { problems.push('appsscript.json が読めない'); }
  const executeAs = manifest.webapp && manifest.webapp.executeAs;
  if (executeAs !== 'USER_DEPLOYING') {
    problems.push(`executeAs が USER_DEPLOYING でない（${executeAs}）。getEffectiveUser が開いた本人を返すようになる`);
  }
  problems.length
    ? fail('G12', `先生の導き方が危うい: ${problems.join(' / ')}`)
    : pass('G12', '先生はデプロイした本人から導いている（executeAs は USER_DEPLOYING）');
}

// ── G13: 設定作業をスプレッドシート側に残していないこと ──
//
// 「コピーしてデプロイするだけ」が売りなので、メニューに書き込み系の入口を戻さない。
// 点検（読むだけ）と修整（先生が押したときだけ）以外を足すと、案内と実物がずれる。
{
  const main = read('Main.gs');
  const menu = main.slice(main.indexOf('function onOpen'), main.indexOf('function showSetupStatus'));
  const items = [...menu.matchAll(/addItem\(\s*'([^']+)'\s*,\s*'([^']+)'/g)].map((m) => m[2]);
  const allowed = new Set(['showSetupStatus', 'showSheetCheck', 'showSheetRepair']);
  const extra = items.filter((fn) => !allowed.has(fn));
  extra.length
    ? fail('G13', `スプレッドシートのメニューに、想定外の入口がある: ${extra.join(' / ')}（設定作業はアプリ側に置くこと）`)
    : pass('G13', `メニューは ${items.length} 項目で、設定作業は 1 つも無い`);
}

// ── G8: 配布テンプレートのコピーリンクが、どこも同じ 1 つを指している ──
//
// コピーリンクは案内ページ・README（2 か所）・紹介記事の 4 か所にある。
// テンプレートを作り直すと ID が変わるが、**変えるのは人の手**なので、
// 1 か所だけ古い ID が残る形で必ずずれる。古い ID を踏んだ先生は、
// 前のテンプレート（＝別人の _meta が入ったままのファイル）をコピーしてしまう。
// 埋まっていない（PUT-TEMPLATE-FILE-ID-HERE のまま）ときは注意にとどめる。
{
  const SITES = [
    'docs/index.html',
    'README.md',
    'docs/note/reflection-journal-note-article.md'
  ];
  const RE = /docs\.google\.com\/spreadsheets\/d\/([A-Za-z0-9_-]+)\/copy/g;
  const found = [];
  for (const f of SITES) {
    if (!has(f)) continue;
    for (const m of read(f).matchAll(RE)) found.push({ file: f, id: m[1] });
  }
  const placeholder = found.filter((x) => x.id === 'PUT-TEMPLATE-FILE-ID-HERE');
  const real = found.filter((x) => x.id !== 'PUT-TEMPLATE-FILE-ID-HERE');
  const ids = [...new Set(real.map((x) => x.id))];

  if (!found.length) {
    fail('G8', '配布テンプレートのコピーリンクが 1 つも無い（案内ページ・README・紹介記事のどこにも）');
  } else if (placeholder.length && real.length) {
    fail('G8', `コピーリンクの一部だけ ID が入っている（未設定: ${placeholder.map((x) => x.file).join(' / ')}）`);
  } else if (placeholder.length) {
    notes.push('G8: 案内ページの「コピーして始める」がテンプレート未設定のままです（docs/copy-distribution.md の手順で ID を差し替えてください）');
  } else if (ids.length > 1) {
    fail('G8', `コピーリンクが別々のテンプレートを指している（${real.map((x) => `${x.file}=${x.id.slice(0, 12)}…`).join(' / ')}）`);
  } else {
    pass('G8', `コピーリンク ${found.length} か所は、すべて同じテンプレートを指している`);
  }
}

// ── 出力 ──
console.log('── 通った項目 ──');
for (const o of ok) console.log('  ✅ ' + o);
if (notes.length) {
  console.log('── 注意（落とさない） ──');
  for (const n of notes) console.log('  ⚠️  ' + n);
}
if (problems.length) {
  console.log('── 直すべき項目 ──');
  for (const p of problems) console.log('  ❌ ' + p);
  console.log(`\n${problems.length} 件の問題があります。`);
  process.exit(1);
}
console.log(`\n問題は見つかりませんでした（${ok.length} 項目を確認）。`);
