/**
 * 配信中の導入案内ページ（docs/）のブラウザ回帰テスト。
 *
 * このページは「アプリの入口」ではなく「入れかたの案内」である。
 * 児童がここへ来てしまったときに、行き先を間違えないことを見る。
 */
import { test, expect } from '@playwright/test';

test('案内ページは、外部へ1本も取りに行かずに表示できる', async ({ page }) => {
  const external = [];
  page.on('request', (req) => {
    const url = new URL(req.url());
    if (url.hostname !== '127.0.0.1' && url.hostname !== 'localhost') external.push(req.url());
  });
  await page.goto('/docs/');
  await expect(page.getByRole('heading', { name: 'ふりかえりジャーナル', level: 1 })).toBeVisible();
  expect(external, `外部へ取りに行っています: ${external.join(', ')}`).toEqual([]);
});

test('児童がここへ来ても、行き先が最初の画面で分かる', async ({ page }) => {
  await page.goto('/docs/');
  const notice = page.getByText('このページからは、ジャーナルを開けません');
  await expect(notice).toBeVisible();
  // 「先生から配られた URL を開いてください」が画面の上のほうにある
  const box = await notice.boundingBox();
  expect(box.y).toBeLessThan(900);
  await expect(page.getByText('先生から配られた URL を開いてください')).toBeVisible();
});

test('入れかたの手順が、デプロイの設定まで書いてある', async ({ page }) => {
  await page.goto('/docs/');
  await expect(page.getByRole('heading', { name: '入れかた' })).toBeVisible();
  await expect(page.getByText('実行するユーザーは「自分」、アクセスできるユーザーは「同じ組織内の全員」')).toBeVisible();
  // 匿名アクセスを選ばせない注意が同じページにある
  await expect(page.getByText('アクセスできるユーザーに「全員（匿名ユーザーを含む）」を選ばないでください。')).toBeVisible();
});

test('横スクロールが出ない（学習者用端末の幅）', async ({ page }) => {
  await page.setViewportSize({ width: 320, height: 568 });
  await page.goto('/docs/');
  const overflow = await page.evaluate(() =>
    document.documentElement.scrollWidth - document.documentElement.clientWidth);
  expect(overflow).toBeLessThanOrEqual(1);
});

test('オフライン用のページから、案内ページへ戻れる', async ({ page }) => {
  await page.goto('/docs/offline.html');
  await page.getByRole('link', { name: '案内ページをひらく' }).click();
  await expect(page.getByRole('heading', { name: 'ふりかえりジャーナル', level: 1 })).toBeVisible();
});
