import { test, expect } from '@playwright/test';

const now = new Date().toISOString();
const classRecord = {
  schemaVersion: 2,
  kind: 'reflection-journal-class',
  classId: 'class-1',
  classCode: 'ABC23456',
  className: '5年1組',
  teacher: { email: 'teacher@example.ed.jp', name: '山田先生' },
  settings: { approvalRequired: true, todayTheme: { date: '', text: '' }, weeklyThemes: { 1: '今日の学び' }, geminiApiKey: '', geminiModel: 'gemini-3.1-flash-lite' },
  members: [{ email: 'student@example.ed.jp', name: '鈴木花子', role: 'student', status: 'active', portfolioFileId: 'portfolio-file', channelFileId: 'channel-file', joinedAt: now }],
  createdAt: now,
  updatedAt: now
};
const portfolio = {
  schemaVersion: 2,
  kind: 'reflection-journal-portfolio',
  class: { id: 'class-1', code: 'ABC23456', name: '5年1組', teacherEmail: 'teacher@example.ed.jp', teacherName: '山田先生', approvalRequired: true, acceptingMembers: true },
  student: { email: 'student@example.ed.jp', name: '鈴木花子' },
  journals: [{ id: 'journal-1', theme: '今日の学び', content: '友だちの考えを聞いて、別の方法でも答えを見つけられました。', emotion: '💡', imageFileId: null, pastComment: '', createdAt: now, updatedAt: now }],
  createdAt: now,
  updatedAt: now
};
const channel = {
  schemaVersion: 2,
  kind: 'reflection-journal-channel',
  class: { id: 'class-1', code: 'ABC23456', name: '5年1組' },
  teacher: classRecord.teacher,
  student: portfolio.student,
  status: 'active',
  themes: { todayTheme: { date: '', text: '' }, weeklyThemes: { 1: '今日の学び' } },
  feedback: { 'journal-1': { comment: '友だちの考えから方法を広げられましたね。', stamp: '⭐', highlights: [], returned: true, updatedAt: now } },
  createdAt: now,
  updatedAt: now
};

async function mockGoogle(page, role, { denyShared = false } = {}) {
  const email = role === 'teacher' ? 'teacher@example.ed.jp' : 'student@example.ed.jp';
  await page.addInitScript(({ email, denyShared }) => {
    const baseScopes = 'openid email profile https://www.googleapis.com/auth/drive.file';
    const sharedScope = 'https://www.googleapis.com/auth/drive.readonly';
    window.__requestedScopes = [];
    window.google = { accounts: { oauth2: {
      initTokenClient: ({ callback, scope }) => ({ requestAccessToken: () => {
        window.__requestedScopes.push(scope);
        const shared = scope.includes('drive.readonly');
        if (shared && denyShared) return callback({ error: 'access_denied', error_description: '閲覧許可がキャンセルされました。' });
        callback({ access_token: shared ? 'shared-token' : 'base-token', scope: shared ? `${baseScopes} ${sharedScope}` : baseScopes });
      } }),
      hasGrantedAllScopes: (response, ...scopes) => scopes.every((scope) => String(response.scope || '').split(' ').includes(scope))
    } } };
    window.APP_CONFIG = { googleClientId: 'e2e-client', publicEntryUrl: 'http://127.0.0.1:4173/' };
    window.__e2eEmail = email;
  }, { email, denyShared });
  await page.route('https://openidconnect.googleapis.com/v1/userinfo', (route) => route.fulfill({ json: { email, name: role === 'teacher' ? '山田先生' : '鈴木花子' } }));
  await page.route('https://www.googleapis.com/drive/v3/**', async (route) => {
    const url = new URL(route.request().url());
    const path = url.pathname;
    const query = decodeURIComponent(url.searchParams.get('q') || '');
    if (query.includes('sharedWithMe') && route.request().headers().authorization !== 'Bearer shared-token') return route.fulfill({ json: { files: [] } });
    if (path.endsWith('/files') && query.includes("value='class'")) return route.fulfill({ json: { files: role === 'teacher' ? [{ id: 'class-file', modifiedTime: now }] : [] } });
    if (path.endsWith('/files') && query.includes("value='portfolio'")) return route.fulfill({ json: { files: [{ id: 'portfolio-file', modifiedTime: now }] } });
    if (path.endsWith('/files') && query.includes("value='channel'")) return route.fulfill({ json: { files: [{ id: 'channel-file', modifiedTime: now }] } });
    if (path.endsWith('/files/class-file')) return route.fulfill({ json: classRecord });
    if (path.endsWith('/files/portfolio-file')) return route.fulfill({ json: portfolio });
    if (path.endsWith('/files/channel-file')) return route.fulfill({ json: channel });
    return route.fulfill({ json: {} });
  });
  await page.route('https://www.googleapis.com/upload/drive/v3/**', (route) => route.fulfill({ json: { id: 'updated-file', modifiedTime: now } }));
}

async function loginAs(page, role) {
  await mockGoogle(page, role);
  await page.goto('/');
  await page.getByRole('button', { name: 'Googleアカウントで続ける' }).click();
  await page.getByRole('button', { name: role === 'teacher' ? /先生として使う/ : /児童として使う/ }).click();
  await page.getByRole('button', { name: '共有された記録の同期を許可する' }).click();
}

test('児童のモバイル画面はノートを先頭にし、横にはみ出さない', async ({ page }) => {
  await page.setViewportSize({ width: 390, height: 844 });
  await loginAs(page, 'student');
  await page.getByRole('button', { name: /5年1組/ }).click();
  await expect(page.locator('.notebook')).toBeVisible();
  const order = await page.evaluate(() => ({
    notebook: document.querySelector('.notebook').getBoundingClientRect().top,
    tools: document.querySelector('.writing-tools').getBoundingClientRect().top,
    calendar: document.querySelector('.student-calendar').getBoundingClientRect().top,
    overflow: document.documentElement.scrollWidth - document.documentElement.clientWidth
  }));
  expect(order.notebook).toBeLessThan(order.tools);
  expect(order.tools).toBeLessThan(order.calendar);
  expect(order.overflow).toBeLessThanOrEqual(1);
  await expect(page.locator('.writing-tools > .auxiliary-details')).not.toHaveAttribute('open', '');
  await page.getByRole('button', { name: /広く書く/ }).click();
  await expect(page.locator('html')).toHaveClass(/writing-focus/);
  await expect(page.locator('.topbar')).toBeHidden();
  await page.getByRole('button', { name: /もとに戻す/ }).click();
  await page.locator('.writing-tools > .auxiliary-details > summary').click();
  await page.getByRole('button', { name: 'おまかせ書き出し' }).click();
  await expect(page.locator('#content')).not.toHaveValue('');
  await expect(page.locator('#writing-count')).not.toHaveText('0文字');
  await expect(page.getByRole('button', { name: /新しいおへんじ/ })).toBeVisible();
  await page.getByRole('button', { name: /新しいおへんじ/ }).click();
  await expect(page.locator('dialog')).toContainText('友だちの考えから方法を広げられましたね');
});

test('児童のデスクトップ画面はカレンダー・ノート・支援を3列表示する', async ({ page }) => {
  await page.setViewportSize({ width: 1440, height: 1000 });
  await loginAs(page, 'student');
  await page.getByRole('button', { name: /5年1組/ }).click();
  await expect(page.locator('.notebook')).toBeVisible();
  const columns = await page.evaluate(() => ({
    calendar: document.querySelector('.student-calendar').getBoundingClientRect().left,
    notebook: document.querySelector('.notebook').getBoundingClientRect().left,
    tools: document.querySelector('.writing-tools').getBoundingClientRect().left
  }));
  expect(columns.calendar).toBeLessThan(columns.notebook);
  expect(columns.notebook).toBeLessThan(columns.tools);
  await expect(page.locator('.writing-tools > .auxiliary-details')).toHaveAttribute('open', '');
});

for (const device of [
  { name: '小型スマートフォン', viewport: { width: 320, height: 568 }, minWritingWidth: 270, minWritingHeight: 300, compact: true },
  { name: 'スマートフォン横向き', viewport: { width: 844, height: 390 }, minWritingWidth: 790, minWritingHeight: 220, compact: true },
  { name: 'タブレット縦向き', viewport: { width: 768, height: 1024 }, minWritingWidth: 440, minWritingHeight: 360, compact: false },
  { name: 'Chromebook', viewport: { width: 1366, height: 768 }, minWritingWidth: 760, minWritingHeight: 360, compact: false }
]) {
  test(`${device.name}で原稿欄を十分に広く表示する`, async ({ page }) => {
    await page.setViewportSize(device.viewport);
    await loginAs(page, 'student');
    await page.getByRole('button', { name: /5年1組/ }).click();
    const layout = await page.evaluate(() => {
      const notebook = document.querySelector('.notebook').getBoundingClientRect();
      const writing = document.querySelector('.lined-paper').getBoundingClientRect();
      const tools = document.querySelector('.writing-tools').getBoundingClientRect();
      return {
        notebookTop: notebook.top,
        notebookWidth: notebook.width,
        writingWidth: writing.width,
        writingHeight: writing.height,
        toolsTop: tools.top,
        toolsLeft: tools.left,
        overflow: document.documentElement.scrollWidth - document.documentElement.clientWidth
      };
    });
    expect(layout.writingWidth).toBeGreaterThanOrEqual(device.minWritingWidth);
    expect(layout.writingHeight).toBeGreaterThanOrEqual(device.minWritingHeight);
    expect(layout.overflow).toBeLessThanOrEqual(1);
    if (device.compact) {
      expect(layout.notebookTop).toBeLessThan(layout.toolsTop);
      await expect(page.locator('.writing-tools > .auxiliary-details')).not.toHaveAttribute('open', '');
    } else {
      expect(layout.toolsLeft).toBeGreaterThan(0);
      await expect(page.locator('.writing-tools > .auxiliary-details')).toHaveAttribute('open', '');
    }
  });
}

test('教師は提出率・絞り込み・クイック返却・範囲コメントを操作できる', async ({ page }) => {
  await page.setViewportSize({ width: 1280, height: 900 });
  await loginAs(page, 'teacher');
  await page.getByRole('button', { name: /5年1組/ }).click();
  await expect(page.locator('.submission-donut')).toContainText('100%');
  await page.locator('[data-filter="returned"]').click();
  await expect(page.locator('.teacher-journal-card')).toHaveCount(1);
  await page.getByRole('button', { name: /最新に更新/ }).click();
  await expect(page.locator('.submission-donut')).toContainText('100%');
  await page.locator('.journal-open').click();
  await page.locator('#journal-source').evaluate((node) => {
    const range = document.createRange();
    range.setStart(node.firstChild, 0);
    range.setEnd(node.firstChild, 8);
    const selection = getSelection();
    selection.removeAllRanges();
    selection.addRange(range);
  });
  await page.getByRole('button', { name: /選んだ部分にコメント/ }).click();
  await expect(page.locator('.highlight-row')).toHaveCount(1);
});

test('別アカウントの共有記録は説明後の追加許可で同期する', async ({ page }) => {
  await loginAs(page, 'teacher');
  await expect(page.getByRole('button', { name: /5年1組/ })).toBeVisible();
  const scopes = await page.evaluate(() => window.__requestedScopes);
  expect(scopes).toEqual([
    'openid email profile https://www.googleapis.com/auth/drive.file',
    'https://www.googleapis.com/auth/drive.readonly'
  ]);
});

test('共有記録の閲覧を許可しなかった場合は理由と再試行を表示する', async ({ page }) => {
  await mockGoogle(page, 'student', { denyShared: true });
  await page.goto('/');
  await page.getByRole('button', { name: 'Googleアカウントで続ける' }).click();
  await page.getByRole('button', { name: /児童として使う/ }).click();
  await page.getByRole('button', { name: '共有された記録の同期を許可する' }).click();
  await expect(page.getByRole('alert')).toContainText('閲覧許可がキャンセルされました');
  await expect(page.getByRole('button', { name: '共有された記録の同期を許可する' })).toBeVisible();
});
