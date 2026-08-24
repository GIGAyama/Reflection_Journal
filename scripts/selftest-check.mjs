/**
 * 品質ゲートの自己点検 — わざと壊して、ちゃんと落ちることを確かめる。
 *
 * 「0件でした」だけでは、検査が動いているのか何も見ていないのか区別できない。
 * 実際、この確認をしたことで検査そのものの不具合が見つかることがある。
 *
 * 一時ディレクトリに複製してから壊す。リポジトリには一切手を入れない。
 */
import { cpSync, readFileSync, writeFileSync, mkdtempSync, rmSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { tmpdir } from 'node:os';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');

/** 壊し方と、そのとき落ちてほしい検査 ID */
const CASES = [
  ['B6',  'docs/index.html',      (s) => s.replace('</head>', '<script src="https://cdn.tailwindcss.com"></script></head>')],
  ['B8',  'index.html',           (s) => s.replace("<? if (bootMode === 'owner') { ?><?!= include_('qr'); ?><? } ?>", "<?!= include_('qr'); ?>")],
  ['D14', 'docs/index.html',      (s) => s.replace('initial-scale=1.0, viewport-fit=cover', 'initial-scale=1.0, maximum-scale=1, user-scalable=no')],
  ['D1',  'docs/index.html',      (s) => s.replace(', viewport-fit=cover', '')],
  ['D2',  'tools/extra.css',      (s) => s.replace('@supports (height: 100dvh) {', '@supports (height: 1px) {')
                                     .replace('.h-screen { height: 100dvh !important; }', '.h-screen { height: 100vh !important; }')],
  ['F4',  'tools/extra.css',      (s) => s.replace('rt { color: #5f6368; font-weight: 500; }', 'rt { color: #666; }')
                                     .replace('[class*="text-white"] rt,\n', '')],
  ['E6',  'docs/sw.js',           (s) => s.replace('self.clients.claim();', "localStorage.removeItem('x'); self.clients.claim();")],
  ['E5',  'docs/sw.js',           (s) => s.replace('.filter((k) => k.startsWith(CACHE_PREFIX) && k !== CACHE_NAME)', '.filter((k) => k !== CACHE_NAME)')],
  ['E7',  'docs/sw.js',           (s) => s.replace('    // ここでは skipWaiting しない。', '    self.skipWaiting();\n    // ここでは skipWaiting しない。')],
  ['E8',  'docs/offline.html',    (s) => s.replace('<a class="btn btn--primary" href="./">案内ページをひらく</a>', '<button onclick="location.reload()">もういちど ためす</button>')],
  ['D10', 'tools/extra.css',      (s) => s.replace('animation-duration: .01ms !important;', 'animation-duration: 0 !important;')],
  ['D11', 'docs/style.css',       (s) => s.replace('@media (forced-colors: active) {', '@media (min-width: 1px) {')],
  ['E1',  'docs/manifest.webmanifest', (s) => s.replace('"id": "/"', '"id": "/Reflection_Journal/"')],
  ['C5',  'src/app.jsx',          (s) => s.replace("localStorage.removeItem(storeKey('draft'));", 'localStorage.clear();')],

  // ── GAS（正本ゲートに 1 つも無い領域。ここが働かないと .gs は誰にも見られない） ──
  // 認可を 1 つ外す。google.script.run は末尾 `_` の無い関数を誰でも呼べる
  ['G1',  'OwnerApi.gs',          (s) => s.replace('function opResetData() {\n  try {\n    const ctx = assertOwner_();',
                                                   'function opResetData() {\n  try {\n    const ctx = { ss: getDb_() };')],
  // 児童の書き込みからロックを外す
  ['G2',  'MemberApi.gs',         (s) => s.replace('    withScriptLock_(function () {                         // 保持区間は最小に',
                                                   '    (function () {')],
  // 画面から「誰の記録か」を受け取る形に戻す
  ['G3',  'MemberApi.gs',         (s) => s.replace('function mbSync() {', 'function mbSync(email) {')],
  // ほかの児童のメールアドレスを素通しにする
  ['G4',  'MemberApi.gs',         (s) => s.replace('sanitizeJournals_(getJournalsForEmail_(g.ss, g.email))',
                                                   'getJournalsForEmail_(g.ss, g.email)')],
  // スコープを広げる
  ['G5',  'appsscript.json',      (s) => s.replace('"https://www.googleapis.com/auth/spreadsheets.currentonly"',
                                                   '"https://www.googleapis.com/auth/drive"')],
  // 列を番号で引く形に戻す
  ['G6',  'Db.gs',                (s) => s.replace('function headerMap_(sheet) {', 'function headerMapDisabled_(sheet) {')],
  // GAS のスクリプトレットの開き記号を、差し込まれる側に入れる
  ['G10', 'app.html',             (s) => s.replace('const svg = new XMLSerializer().serializeToString(svgEl);',
                                                   'const svg = \'<?xml version="1.0"?>\' + new XMLSerializer().serializeToString(svgEl);')],
  // 差し込みを「中身を読み直す」形に戻す（app.html の中の JavaScript が壊れて届く）
  ['G11', 'Main.gs',              (s) => s.replace('HtmlService.createTemplateFromFile(filename).getRawContent()',
                                                   'HtmlService.createHtmlOutputFromFile(filename).getContent()')],
  // 先生を「開いた本人」から導く（先に開いた児童が先生になる）
  ['G12', 'Main.gs',            (s) => s.replace('const effective = effectiveEmail_();',
                                                 'const effective = activeEmail_();')],
  // 実行ユーザーが「アクセスしているユーザー」に戻る（全員が先生を名乗れる）
  ['G12', 'appsscript.json',    (s) => s.replace('"USER_DEPLOYING"', '"USER_ACCESSING"')],
  // スプレッドシート側に設定作業が戻る
  ['G13', 'Main.gs',            (s) => s.replace(".addItem('準備の状態を見る', 'showSetupStatus')",
                                                 ".addItem('はじめの設定', 'setupAsTeacher')")],
  // コピーリンクの 1 か所だけ古い ID が残る（前のテンプレートを配ってしまう）
  ['G8',  'README.md',            (s) => s.replace('15yVHzXFrPQudGZu9nLt5zXJRtAEwJrdgW-3pub-KJWU',
                                                   '1OLDoldOLDoldOLDoldOLDoldOLDoldOLDoldOLDold')],
  // 実在しない列挙子（undefined が渡り、その画面が開かなくなる）
  ['G9',  'Main.gs',              (s) => s.replace('HtmlService.XFrameOptionsMode.DEFAULT',
                                                   'HtmlService.XFrameOptionsMode.SAMEORIGIN')],
  // 点検が書き換えを始める
  ['G7',  'Db.gs',                (s) => s.replace("      found.push({\n        sheet: spec.name, kind: '見出しが無い',",
                                                   "      sheet.getRange(1, 1).setValue('こわした');\n      found.push({\n        sheet: spec.name, kind: '見出しが無い',")]
];

let bad = 0;
for (const [id, file, breakIt] of CASES) {
  const dir = mkdtempSync(join(tmpdir(), 'rj-selftest-'));
  try {
    cpSync(ROOT, dir, {
      recursive: true,
      filter: (src) => !/node_modules|[\\/]\.git[\\/]?$|[\\/]\.git[\\/]/.test(src)
    });
    const p = join(dir, file);
    const before = readFileSync(p, 'utf8');
    const after = breakIt(before);
    if (after === before) { console.error(`⚠️  ${id}: 壊し方が当たっていない（${file} が変わらなかった）`); bad++; continue; }
    writeFileSync(p, after);

    let failed = false, out = '';
    try {
      out = execFileSync(process.execPath, [join(dir, 'scripts/check-project.mjs')], { encoding: 'utf8' });
    } catch (e) {
      failed = true;
      out = (e.stdout || '') + (e.stderr || '');
    }
    if (failed && out.includes(`❌ ${id}`)) {
      console.log(`✅ ${id}: 壊したら ${id} で落ちた`);
    } else {
      console.error(`❌ ${id}: 壊したのに ${id} で落ちなかった（検査が働いていない）`);
      bad++;
    }
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

if (bad) {
  console.error(`\n${bad} 件の検査が働いていません。`);
  process.exit(1);
}
console.log(`\n${CASES.length} 件すべて、壊したときに落ちることを確認しました。`);
