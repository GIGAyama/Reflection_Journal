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
  /* ⚠️ テストごとに別のディレクトリへ書く。置き場を mode だけで決めてはいけない。
   *
   *    2026-08-29 まで `dist/gas-page/<mode>/` に書いていた。ところが
   *    mode='member' を使うテストが 3 つあり、それぞれ別の boot を渡している
   *    （既定 / signedIn:false / bootError）。playwright.config は
   *    fullyParallel なので、3 つが**同じ index.html を互いに上書き**し、
   *    書いた本人とは違う中身を読みにいくことがあった。
   *
   *    そのため「ログインが確かめられないときは、その理由が画面に出る」が
   *    ときどきだけ落ちる。手元で --workers=4 にして 10 回走らせると 4 回落ちた。
   *    CI でも 2026-08-27 以降、成功と失敗が無関係なコミットで交互に出ていた。
   *
   *    ⚠️ これは「たまに落ちる検査」ではなく、**検査どうしがぶつかっていた**。
   *       落ちた回は、見たかった画面をそもそも見ていない。 */
  const key = `${mode}-${test.info().testId}`.replace(/[^A-Za-z0-9_-]/g, '');
  const dir = join(ROOT, 'dist/gas-page', key);
  mkdirSync(dir, { recursive: true });
  writeFileSync(join(dir, 'index.html'), assembleGasPage(mode, boot));

  const errors = [];
  page.on('pageerror', (e) => errors.push(String(e.message)));
  // 外部の書体やアイコンが届かないのは想定内。画面が出るかどうかだけを見る。
  await page.goto(`/dist/gas-page/${key}/`, { waitUntil: 'load' });
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

test('実行ユーザーの設定ミスは、画面に理由が出る', async ({ page }) => {
  // 「はじめの設定」を押す前の「準備中」画面は無くなった（押すものが無いため）。
  // 代わりに、画面が出ない唯一の理由がこれになる。
  const errors = await open(page, 'member', {
    bootError: 'デプロイの「実行するユーザー」が「自分」になっていません。'
  });
  expect(errors.filter((e) => /SyntaxError/i.test(e)), errors.join('\n')).toEqual([]);
  await expect(page.getByText('実行するユーザー')).toBeVisible();
});
