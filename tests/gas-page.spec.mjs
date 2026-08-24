/**
 * GAS が組み立てて配る 1 枚の HTML を、本物のブラウザに読ませる。
 *
 * ★ なぜ要るか ★
 * 2026-08-24、貼り合わせた側だけが壊れて「タブに題は出るが画面が出ない」状態になった。
 * app.html 単体は構文として妥当で、`node --check` も通り、テストも 85 件すべて緑だった。
 * **ファイル単位の検査では、貼り合わせた 1 枚が動くことを一度も見ていなかった。**
 */
import { test, expect } from '@playwright/test';
import { mkdirSync, writeFileSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { assembleGasPage } from '../scripts/assemble-gas-page.mjs';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');

/** 組み立てた 1 枚を dist/ へ置き、ブラウザに読ませて、出たエラーを全部返す */
async function open(page, mode, boot) {
  const dir = join(ROOT, 'dist/gas-page', mode);
  mkdirSync(dir, { recursive: true });
  writeFileSync(join(dir, 'index.html'), assembleGasPage(mode, boot));

  const errors = [];
  page.on('pageerror', (e) => errors.push(String(e.message)));
  // 外部の書体やアイコンが届かないのは想定内。画面が出るかどうかだけを見る。
  await page.goto(`/dist/gas-page/${mode}/`, { waitUntil: 'load' });
  await page.waitForTimeout(500);
  return errors;
}

test('先生の画面: 貼り合わせた1枚が構文エラーなく動き、中身が描かれる', async ({ page }) => {
  const errors = await open(page, 'owner');
  expect(errors.filter((e) => /SyntaxError/i.test(e)), errors.join('\n')).toEqual([]);
  const html = await page.evaluate(() => document.getElementById('root')?.innerHTML ?? '');
  expect(html.length).toBeGreaterThan(100);
});

test('児童の画面: 貼り合わせた1枚が構文エラーなく動き、中身が描かれる', async ({ page }) => {
  const errors = await open(page, 'member');
  expect(errors.filter((e) => /SyntaxError/i.test(e)), errors.join('\n')).toEqual([]);
  const html = await page.evaluate(() => document.getElementById('root')?.innerHTML ?? '');
  expect(html.length).toBeGreaterThan(100);
});

test('ログインが確かめられないときは、その理由が画面に出る', async ({ page }) => {
  const errors = await open(page, 'member', { signedIn: false });
  expect(errors.filter((e) => /SyntaxError/i.test(e)), errors.join('\n')).toEqual([]);
  await expect(page.getByText('学校のアカウントでログインしてから')).toBeVisible();
});

test('先生が「はじめの設定」を押す前は、準備中とだけ出る', async ({ page }) => {
  const errors = await open(page, 'member', { setupDone: false });
  expect(errors.filter((e) => /SyntaxError/i.test(e)), errors.join('\n')).toEqual([]);
  await expect(page.getByText('先生の じゅんびが おわるまで まってね。')).toBeVisible();
});
