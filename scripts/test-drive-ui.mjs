import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

const index = readFileSync(new URL('../docs/index.html', import.meta.url), 'utf8');
const app = readFileSync(new URL('../docs/drive-app.js', import.meta.url), 'utf8');
const css = readFileSync(new URL('../docs/drive.css', import.meta.url), 'utf8');
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
  assert.match(app, /drive\.readonly/);
  assert.match(app, /requireSharedRead/);
  assert.match(app, /共有された記録の同期を許可する/);
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
    'できた！ 先生にとどける'
  ]) assert.ok(app.includes(marker), `UIに「${marker}」がありません`);
});

test('PR #10の児童向け学習体験をDrive版で復元している', () => {
  for (const marker of [
    'JOURNAL_TEMPLATES',
    'RANDOM_STARTERS',
    'HINT_GROUPS',
    'studentCalendar',
    'unreadFeedbackJournals',
    'openStudentJournal',
    'celebrateSubmission',
    '<ruby>'
  ]) assert.ok(app.includes(marker), `児童UIに「${marker}」がありません`);
  assert.match(css, /\.student-workspace\s*\{/);
  assert.match(css, /\.lined-paper/);
  assert.match(css, /\.feedback-alert/);
  assert.match(css, /@keyframes confetti-fall/);
});

test('PR #10の教師支援・招待・PWA操作をDrive版で復元している', () => {
  for (const marker of [
    'submission-donut',
    'quickFeedback',
    'batchReturn',
    'batchAiDrafts',
    'addSelectionHighlight',
    'class-switch',
    'addBulkMembers',
    'downloadQrCode',
    'initPwaControls',
    'SKIP_WAITING'
  ]) assert.ok(app.includes(marker) || index.includes(marker), `復元UIに「${marker}」がありません`);
  assert.match(index, /id="pwa-install"/);
  assert.match(index, /id="pwa-update"/);
  assert.match(css, /\.submission-donut/);
  assert.match(css, /\.teacher-highlight/);
});

test('別アカウント共有と教師画面の更新を明示的に扱う', () => {
  for (const marker of ['SHARED_READ_SCOPE', 'handleSharedToken', 'refreshTeacherClass', 'scheduleTeacherRefresh', '最新に更新', '最終同期']) {
    assert.ok(app.includes(marker), `複数アカウント同期に「${marker}」がありません`);
  }
  assert.match(css, /\.permission-card/);
  assert.match(css, /\.sync-status/);
});

test('トップ画面は中央配置され、内部用語やクリックイベントを表示しない', () => {
  assert.match(app, /class="role-home"/);
  assert.match(css, /\.role-home\s*\{[^}]*margin-inline:\s*auto/s);
  assert.match(css, /\.center-screen\s*\{[^}]*justify-content:\s*center/s);
  assert.doesNotMatch(app, /Driveネイティブ/);
  assert.doesNotMatch(app, /addEventListener\('click',\s*render(?:Home|TeacherHome|StudentHome)\)/);
  assert.match(app, /typeof message === 'string'/);
});

test('ログイン画面とヘッダーはPWAと同じアプリアイコンを使う', () => {
  assert.match(app, /class="app-logo" src="\.\/icon-192\.png"/);
  assert.match(app, /class="brand-icon" src="\.\/icon-192\.png"/);
  assert.doesNotMatch(app, /📔/);
});
