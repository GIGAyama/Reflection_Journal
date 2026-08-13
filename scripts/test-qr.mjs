import assert from 'node:assert/strict';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
const qrcode = require('qrcode-generator');
const memberUrl = 'https://example.github.io/Reflection_Journal/?t=ABCDEFGH';

const make = () => {
  const qr = qrcode(0, 'M');
  qr.addData(memberUrl, 'Byte');
  qr.make();
  return qr;
};

const first = make();
const second = make();
const count = first.getModuleCount();
let darkModules = 0;
let signature = '';

for (let row = 0; row < count; row += 1) {
  for (let col = 0; col < count; col += 1) {
    if (first.isDark(row, col)) darkModules += 1;
    signature += first.isDark(row, col) ? '1' : '0';
  }
}

assert.ok(count >= 21 && count <= 177, `QRの一辺が規格外です: ${count}`);
assert.ok(darkModules > 0 && darkModules < count * count, 'QRの明暗モジュールを生成できていません');

let secondSignature = '';
for (let row = 0; row < second.getModuleCount(); row += 1) {
  for (let col = 0; col < second.getModuleCount(); col += 1) {
    secondSignature += second.isDark(row, col) ? '1' : '0';
  }
}
assert.equal(second.getModuleCount(), count, '同じURLでQRサイズが変化しました');
assert.equal(secondSignature, signature, '同じURLから同じQRを再現できません');

console.log(`✅ QR生成テスト成功: ${count}×${count}、暗色 ${darkModules} モジュール`);
