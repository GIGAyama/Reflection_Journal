/**
 * ビルド — 原本（src/ と tools/）から、GAS に置く生成物を作る。
 *
 * なぜビルドするのか:
 *   以前は React・ReactDOM・Babel・Tailwind をブラウザへ CDN から読み込み、
 *   1,400 行の JSX を「開くたびに」ブラウザの中で翻訳していた。
 *   学校のネットワークは cdn.jsdelivr.net / unpkg.com / cdn.tailwindcss.com を
 *   塞いでいることがあり、1本でも届かないと画面が白いまま何も出ない。
 *   児童からは「壊れている」としか見えず、原因はアプリの外にあるので先生も追えない。
 *   @babel/standalone だけで 2,997.6 KB あった。
 *
 * 生成物（手で編集しない）:
 *   vendor.html … react / react-dom / canvas-confetti（npm の実バイト）
 *   qr.html     … QRコード生成（先生の画面だけで読み込む）
 *   css.html    … Tailwind が生成した CSS + tools/extra.css
 *   app.html    … src/app.jsx をコンパイルした JS
 *
 * 原本（ここを直す）:
 *   src/app.jsx / tools/extra.css / tailwind.config.js / index.html
 *
 * ⚠️ 生成物をコミットしている。原本を直してビルドを走らせずに push すると、
 *    GAS には古い画面が出たままになる。CI の check-generated.mjs がそれを止める。
 */
import { readFileSync, writeFileSync, existsSync, mkdtempSync, rmSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { tmpdir } from 'node:os';
import { transformSync } from '@babel/core';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const kb = (s) => (Buffer.byteLength(s, 'utf8') / 1024).toFixed(1) + ' KB';

/** 生成した JS/CSS を GAS の .html に包む。
 *  文字列リテラルの中の </script> がそのまま出ると、そこで script が終わってしまうので必ず割る。 */
function wrapScript(js) {
  return '<script>\n' + js.replace(/<\/script/gi, '<\\/script') + '\n</script>\n';
}
function wrapStyle(css) {
  return '<style>\n' + css.replace(/<\/style/gi, '<\\/style') + '\n</style>\n';
}

// ── ① vendor.html：実行コードは自分側に持つ（CDN から取る実行コードを 0 バイトにする） ──
const VENDOR = [
  ['react',           'node_modules/react/umd/react.production.min.js'],
  ['react-dom',       'node_modules/react-dom/umd/react-dom.production.min.js'],
  // npm の canvas-confetti に confetti.browser.min.js は無い（jsDelivr が自動生成しているだけ）
  ['canvas-confetti', 'node_modules/canvas-confetti/dist/confetti.browser.js']
];
let vendor = '<!-- 生成物。手で編集しない（npm run build で作り直す） -->\n';
for (const [name, rel] of VENDOR) {
  const p = join(ROOT, rel);
  if (!existsSync(p)) throw new Error(`${name} が見つかりません: ${rel}（npm ci を実行してください）`);
  vendor += `<!-- ${name} -->\n` + wrapScript(readFileSync(p, 'utf8'));
}
writeFileSync(join(ROOT, 'vendor.html'), vendor);

// ── ①-b qr.html：児童用URLを外部QRサービスへ送らず、ブラウザ内で生成する ──
// QRは先生の画面だけで使う。index.html 側で owner のときだけ include し、
// 児童端末の初回JS（F5: 300KB以下）には一切含めない。
const qrSource = readFileSync(join(ROOT, 'node_modules/qrcode-generator/qrcode.js'), 'utf8');
const { code: qrCode } = transformSync(qrSource, {
  comments: false,
  compact: true,
  minified: true,
  babelrc: false,
  configFile: false
});
writeFileSync(join(ROOT, 'qr.html'),
  '<!-- 生成物。手で編集しない（npm run build で作り直す・先生の画面専用） -->\n' + wrapScript(qrCode));

// ── ② css.html：使うクラスだけを先に作る（ブラウザ内で CSS を生成しない） ──
const tmp = mkdtempSync(join(tmpdir(), 'rj-build-'));
const inCss = join(tmp, 'in.css');
const outCss = join(tmp, 'out.css');
writeFileSync(inCss, '@tailwind base;\n@tailwind components;\n@tailwind utilities;\n');
const twBin = join(ROOT, 'node_modules/.bin/tailwindcss');
execFileSync(twBin, ['-c', join(ROOT, 'tailwind.config.js'), '-i', inCss, '-o', outCss, '--minify'],
  { stdio: ['ignore', 'ignore', 'inherit'] });
const css = readFileSync(outCss, 'utf8') + '\n' + readFileSync(join(ROOT, 'tools/extra.css'), 'utf8');
rmSync(tmp, { recursive: true, force: true });
writeFileSync(join(ROOT, 'css.html'),
  '<!-- 生成物。手で編集しない（原本は tools/extra.css と tailwind.config.js） -->\n' + wrapStyle(css));

// ── ③ app.html：JSX の翻訳はビルド時に1回だけ ──
const jsx = readFileSync(join(ROOT, 'src/app.jsx'), 'utf8');
const { code } = transformSync(jsx, {
  filename: 'app.jsx',
  presets: [['@babel/preset-react', { runtime: 'classic' }]],
  comments: false,          // 配る側では落とす。「なぜ」の説明は src/app.jsx に残る
  compact: false,
  babelrc: false,
  configFile: false
});
writeFileSync(join(ROOT, 'app.html'),
  '<!-- 生成物。手で編集しない（原本は src/app.jsx） -->\n' + wrapScript(code));

console.log('ビルド完了');
console.log('  vendor.html', kb(vendor));
console.log('  qr.html    ', kb(qrCode), '（先生の画面のみ）');
console.log('  css.html   ', kb(css));
console.log('  app.html   ', kb(code));
// F5（初回JS 300KB以下）に数えるのは JS だけ。CSS は別に出す。
console.log('  児童の初回JS', kb(vendor + code), '（vendor + app。CSS は含めない）');
