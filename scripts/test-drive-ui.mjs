import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

const index = readFileSync(new URL('../docs/index.html', import.meta.url), 'utf8');
const app = readFileSync(new URL('../docs/drive-app.js', import.meta.url), 'utf8');
const config = readFileSync(new URL('../docs/config.js', import.meta.url), 'utf8');

test('GitHub Pages共通URLがDrive版を直接起動する', () => {
  assert.match(index, /drive-app\.js/);
  assert.match(index, /drive\.css/);
  assert.doesNotMatch(index, /<iframe|script\.google\.com|execUrl/);
  assert.doesNotMatch(config, /execUrlA|execUrlB|backendMode/);
});

test('認証情報をブラウザストレージへ保存しない', () => {
  assert.match(app, /initTokenClient/);
  assert.match(app, /drive\.file/);
  assert.doesNotMatch(app, /localStorage\.setItem\([^\n]*(token|credential|auth)/i);
  assert.doesNotMatch(app, /sessionStorage/);
});

test('完全移行した主要な教師・児童フローを含む', () => {
  for (const marker of [
    '参加申請・名簿',
    'テーマを保存して児童へ配信',
    '保存して返却',
    'AIで下書き',
    'クラスの心の波',
    'CSVをダウンロード',
    '過去の自分と対話',
    'Driveに保存して先生へ共有'
  ]) assert.ok(app.includes(marker), `UIに「${marker}」がありません`);
});

