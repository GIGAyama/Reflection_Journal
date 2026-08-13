/**
 * GitHub Pages配信物の生成。
 * QRコードは外部サービスへ招待URLを送らず、npmに固定した実装をローカル配信する。
 */
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const sourcePath = join(ROOT, 'node_modules/qrcode-generator/qrcode.js');
if (!existsSync(sourcePath)) throw new Error('qrcode-generator が見つかりません。npm ci を実行してください。');

const code = readFileSync(sourcePath, 'utf8');

writeFileSync(join(ROOT, 'docs/qrcode.js'), `/* 生成物。手で編集しない（npm run build で作り直す） */\n${code}\n`);
console.log(`ビルド完了: docs/qrcode.js ${(Buffer.byteLength(code) / 1024).toFixed(1)} KB`);
