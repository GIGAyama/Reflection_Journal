/**
 * 生成物が原本と食い違っていないかを見る。
 *
 *   app.html    ← src/app.jsx
 *   css.html    ← tools/extra.css / tailwind.config.js
 *   vendor.html ← node_modules の react / react-dom / canvas-confetti
 *   qr.html     ← node_modules の qrcode-generator
 *
 * 原本を直したのにビルドを走らせずに push すると、GAS には古い画面が出たままになる。
 * その食い違いは、動かしてみるまで誰にも見えない。CI で必ず止める。
 *
 * 使い方: npm run build のあとに実行する（build → このチェック、の順で意味を持つ）。
 */
import { execSync } from 'node:child_process';

const GENERATED = ['app.html', 'css.html', 'vendor.html', 'qr.html'];

let diff = '';
try {
  diff = execSync(`git diff --name-only -- ${GENERATED.join(' ')}`, { encoding: 'utf8' }).trim();
} catch (e) {
  console.log('git が使えないため、生成物の照合はできませんでした（未計測）。');
  process.exit(0);
}

if (diff) {
  console.error('❌ 生成物が原本と食い違っています:\n  ' + diff.split('\n').join('\n  '));
console.error('\n依存関係または tools/build.mjs を直したら、');
  console.error('必ず `npm run build` を走らせてから push してください。');
  process.exit(1);
}
console.log('✅ 生成物は原本と一致しています');
