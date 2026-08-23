import { defineConfig } from '@playwright/test';

export default defineConfig({
  testDir: './tests',
  timeout: 30_000,
  fullyParallel: true,
  use: {
    baseURL: 'http://127.0.0.1:4173',
    browserName: 'chromium',
    locale: 'ja-JP',
    timezoneId: 'Asia/Tokyo',
    reducedMotion: 'reduce'
  },
  webServer: {
    // 配信中の案内ページ（docs/）と、退役した版（legacy/drive-native/）の
    // 両方を同じサーバーから見るため、リポジトリの根から配る。
    command: 'python3 -m http.server 4173',
    url: 'http://127.0.0.1:4173',
    reuseExistingServer: true
  },
  reporter: [['list']]
});
