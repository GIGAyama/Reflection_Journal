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
  ['B6',  'index.html',      (s) => s.replace('</head>', '<script src="https://cdn.tailwindcss.com"></script></head>')],
  ['B8',  'src/app.jsx',     (s) => s.replace('const makeQrModel', "const externalQr = 'https://api.qrserver.com';\n    const makeQrModel")],
  ['D14', 'index.html',      (s) => s.replace('initial-scale=1.0, viewport-fit=cover', 'initial-scale=1.0, maximum-scale=1.0, user-scalable=no')],
  ['D1',  'Main.gs',         (s) => s.replace(", viewport-fit=cover'", "'")],
  ['F4',  'tools/extra.css', (s) => s.replace(/rt \{ color: #5f6368; font-weight: 500; \}/, 'rt { color: #666; }')
                                     .replace(/\[class\*="text-white"\] rt,/, '')],
  ['E6',  'docs/sw.js',      (s) => s.replace("self.clients.claim();", "localStorage.removeItem('x'); self.clients.claim();")],
  ['E5',  'docs/sw.js',      (s) => s.replace(".filter((k) => k.startsWith(CACHE_PREFIX) && k !== CACHE_NAME)", ".filter((k) => k !== CACHE_NAME)")],
  ['D10', 'tools/extra.css', (s) => s.replace('animation-duration: .01ms !important;', 'animation-duration: 0 !important;')],
  ['E1',  'docs/manifest.webmanifest', (s) => s.replace('"id": "/Reflection_Journal/"', '"id": "/"')],
  ['C5',  'src/app.jsx',     (s) => s.replace('localStorage.removeItem(storeKey(\'draft\'));', 'localStorage.clear();')]
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
