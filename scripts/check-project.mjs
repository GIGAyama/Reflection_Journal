/**
 * 品質ゲート — GIGA Standard v5 Part I のうち、静的に確かめられるものを機械で見る。
 *
 * 「0件でした」だけでは、検査が動いているのか何も見ていないのか区別できない。
 * わざと壊して落ちることを確かめてから使うこと（scripts/selftest-check.mjs）。
 *
 * 既知の誤検知に対処してある:
 *   - localStorage は「注意書きのコメント」に反応しやすい → 判定前にコメントを落とす
 *   - 100vh は @supports not (height: 100dvh) の中に書くのが正しい → 前方も見る
 *   - キャッシュ削除は「消す式」ではなく「startsWith で絞る式があるか」を見る
 */
import { readFileSync, existsSync, statSync, readdirSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const read = (p) => readFileSync(join(ROOT, p), 'utf8');
const size = (p) => statSync(join(ROOT, p)).size;
const problems = [];
const notes = [];
const fail = (id, msg) => problems.push(`${id}: ${msg}`);
const ok = [];
const pass = (id, msg) => ok.push(`${id}: ${msg}`);

/** JS/CSS のコメントを落とす。注意書きに反応する誤検知を止めるため */
const stripComments = (s) =>
  s.replace(/\/\*[\s\S]*?\*\//g, ' ').replace(/(^|[^:])\/\/[^\n]*/g, '$1 ');

const htmlFiles = ['docs/index.html', 'docs/diag.html', 'docs/offline.html'];
// 初期表示で読み込むJS。キットも同じ予算に入れる（切り出したぶんを見落とさないため）。
const kitFiles = ['namespace.js', 'invite.js', 'drive-client.js', 'records.js', 'session.js', 'index.js']
  .map((name) => `docs/kit/${name}`);
const appJsFiles = ['docs/drive-app.js', 'docs/drive-api.js', 'docs/drive-core.js', ...kitFiles];
// GitHub Pages + Drive API が本番実装。ルート直下の旧GAS資産は品質判定に含めない。
const gsFiles = [];

// ── B6: CDN から取る実行コードが 0 バイト ──
{
  const bad = [];
  for (const f of htmlFiles) {
    const html = read(f);
    for (const m of html.matchAll(/<script[^>]+src=["'](https?:\/\/[^"']+)["']/gi)) bad.push(`${f} → ${m[1]}`);
    // ブラウザ内で CSS を生成する Tailwind CDN も実行コード扱い
    if (/cdn\.tailwindcss\.com/.test(html.replace(/<!--[\s\S]*?-->/g, ''))) bad.push(`${f} → cdn.tailwindcss.com`);
    if (/@babel\/standalone/.test(html.replace(/<!--[\s\S]*?-->/g, ''))) bad.push(`${f} → @babel/standalone`);
  }
  bad.length ? fail('B6', `CDN から実行コードを読んでいる: ${bad.join(' / ')}`)
             : pass('B6', 'CDN から取る実行コード 0 件');
}

// ── B8: QRコードはローカル生成し、児童用URLを第三者へ送らない ──
{
  const source = read('docs/drive-app.js');
  const index = read('docs/index.html');
  if (/api\.qrserver\.com|chart\.googleapis\.com.*cht=qr/i.test(source + index)) {
    fail('B8', 'QR生成のため児童用URLを外部サービスへ送っている');
  } else if (!existsSync(join(ROOT, 'docs/qrcode.js')) || /<script[^>]+qrcode\.js/.test(index) || !/ensureQr\s*\(\)/.test(source)) {
    fail('B8', 'ローカルQR生成コードがクラス設定画面からの遅延読込になっていない');
  } else {
    pass('B8', 'QRはブラウザ内で生成し、生成コードは教師ポータルだけに配信');
  }
}

// ── D14 / D1: 拡大禁止と viewport-fit（GAS は .gs 側の addMetaTag も見る） ──
{
  const targets = htmlFiles;
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
  for (const f of [...htmlFiles, 'docs/drive.css']) {
    if (!existsSync(join(ROOT, f))) continue;
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
  const css = read('docs/drive.css');
  const jsx = read('docs/drive-app.js');
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

// ── D10 / D11: 動きの配慮とハイコントラスト ──
{
  const css = [read('docs/drive.css'), read('docs/index.html')].join('\n');
  /prefers-reduced-motion/.test(css) ? pass('D10', 'prefers-reduced-motion あり')
                                     : fail('D10', 'prefers-reduced-motion が無い');
  if (/animation-duration:\s*0\s*(!important)?\s*;/.test(css)) {
    fail('D10', 'animation-duration: 0 になっている（.01ms にしないと fill-mode: forwards が壊れて中身が消える）');
  }
  /forced-colors/.test(css) ? pass('D11', 'forced-colors あり') : fail('D11', 'forced-colors が無い');
}

// ── E5 / E6 / sw.js ──
{
  const swPath = 'docs/sw.js';
  const swRaw = read(swPath);
  const sw = stripComments(swRaw);
  // 「消す式」ではなく「startsWith で絞る式があるか」を見る
  if (/caches\.keys\(\)/.test(sw) && !/startsWith\s*\(/.test(sw)) {
    fail('E5', 'sw.js が caches.keys() を接頭辞で絞らずに扱っている（同一オリジンの他アプリを巻き添えにする）');
  } else pass('E5', 'sw.js は自アプリ接頭辞のキャッシュだけを扱っている');

  /localStorage/.test(sw) ? fail('E6', 'sw.js が localStorage に触れている') : pass('E6', 'sw.js は localStorage に触れていない');

  const installBlock = sw.slice(sw.indexOf("addEventListener('install'"), sw.indexOf("addEventListener('activate'"));
  /skipWaiting\s*\(/.test(installBlock)
    ? fail('E7', 'install の中で skipWaiting している（書いている最中に画面が入れ替わる）')
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

// ── C5: localStorage.clear() ──
{
  const files = [...htmlFiles, ...appJsFiles].filter((f) => existsSync(join(ROOT, f)));
  const bad = files.filter((f) => /localStorage\.clear\s*\(/.test(stripComments(read(f))));
  bad.length ? fail('C5', `localStorage.clear() を使っている: ${bad.join(', ')}`)
             : pass('C5', 'localStorage.clear() を使っていない');
}

// ── F5: 初回 JS 300KB 以下 ──
{
  const js = appJsFiles.reduce((s, f) => s + size(f), 0);
  const kb = js / 1024;
  kb <= 300 ? pass('F5', `初回 JS ${kb.toFixed(1)} KB（300KB 以下）`)
            : fail('F5', `初回 JS が ${kb.toFixed(1)} KB（300KB を超えている）`);
}

// ── F6: 1ファイル 5,000行 / 400KB ──
{
  const bad = [];
  for (const f of [...htmlFiles, 'docs/qrcode.js', ...appJsFiles]) {
    if (!existsSync(join(ROOT, f))) continue;
    const s = read(f);
    if (s.split('\n').length > 5000 || size(f) > 400 * 1024) bad.push(f);
  }
  bad.length ? fail('F6', `5,000行 / 400KB を超えるファイル: ${bad.join(', ')}`)
             : pass('F6', '5,000行 / 400KB を超えるファイルなし');
}

// ── B4: postMessage の宛先 ──
{
  const bad = [];
  for (const f of htmlFiles) {
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

// ── 画像 ──
{
  const bad = [];
  for (const f of readdirSync(join(ROOT, 'docs')).filter((f) => f.endsWith('.png'))) {
    const kb = size(join('docs', f)) / 1024;
    const limit = /512/.test(f) ? 60 : /192|180/.test(f) ? 30 : 150;
    if (kb > limit) bad.push(`${f} ${kb.toFixed(1)}KB > ${limit}KB`);
  }
  bad.length ? fail('D7', `画像が上限を超えている: ${bad.join(', ')}`) : pass('D7', '画像はすべて上限内');
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
