import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

const index = readFileSync(new URL('../docs/index.html', import.meta.url), 'utf8');
const app = readFileSync(new URL('../docs/drive-app.js', import.meta.url), 'utf8');
const css = readFileSync(new URL('../docs/drive.css', import.meta.url), 'utf8');
const config = readFileSync(new URL('../docs/config.js', import.meta.url), 'utf8');
const readme = readFileSync(new URL('../README.md', import.meta.url), 'utf8');
const manual = readFileSync(new URL('../MANUAL.md', import.meta.url), 'utf8');
const architecture = readFileSync(new URL('../docs/DRIVE_NATIVE_ARCHITECTURE.md', import.meta.url), 'utf8');
const oauthGuide = readFileSync(new URL('../docs/OAUTH_SHARED_RECORDS_SETUP.md', import.meta.url), 'utf8');
// 共通部分（分散ポートフォリオのキット）。本体と同じく初期表示で読み込まれる。
const kitSources = Object.fromEntries(
  ['namespace.js', 'invite.js', 'drive-client.js', 'records.js', 'session.js', 'index.js']
    .map((name) => [name, readFileSync(new URL(`../docs/kit/${name}`, import.meta.url), 'utf8')])
);
const kitSession = kitSources['session.js'];
const kitIndex = kitSources['index.js'];
const core = readFileSync(new URL('../docs/drive-core.js', import.meta.url), 'utf8');
const api = readFileSync(new URL('../docs/drive-api.js', import.meta.url), 'utf8');
const sw = readFileSync(new URL('../docs/sw.js', import.meta.url), 'utf8');
const portingGuide = readFileSync(new URL('../docs/PORTING_FROM_GAS.md', import.meta.url), 'utf8');
const kitReadme = readFileSync(new URL('../docs/kit/README.md', import.meta.url), 'utf8');

test('GitHub Pages共通URLがDrive版を直接起動する', () => {
  assert.match(index, /drive-app\.js/);
  assert.match(index, /drive\.css/);
  assert.doesNotMatch(index, /<iframe|script\.google\.com|execUrl/);
  assert.doesNotMatch(config, /execUrlA|execUrlB|backendMode/);
});

test('本番配信と共有データの安全境界を実装する', () => {
  assert.match(index, /Content-Security-Policy/);
  assert.match(index, /default-src 'self'/);
  assert.doesNotMatch(index, /<script>[^<]/s);
  assert.match(config, /allowedOrigins/);
  assert.match(config, /allowedWorkspaceDomains/);
  assert.match(app, /validatePortfolioForClass/);
  assert.match(app, /validateChannelForStudent/);
  assert.match(app, /encodeSignedInvite/);
  assert.match(app, /expectedVersion/);
  assert.match(app, /Driveへ完全バックアップ/);
  assert.match(app, /geminiDataConsent/);
});

test('認証情報と児童下書きは共有オリジンの永続領域へ残さない', () => {
  assert.match(app, /initTokenClient/);
  assert.match(kitIndex, /drive\.file/);
  assert.match(kitIndex, /drive\.readonly/);
  assert.match(app, /requireSharedRead/);
  assert.match(app, /INITIAL_SCOPES/);
  assert.match(app, /tokenExpiresAt/);
  assert.match(config, /persistSessionToken:\s*false/);
  assert.match(config, /persistLocalDrafts:\s*false/);

  // 保存の可否はキットの SessionPolicy が握る。保存を選んだときだけ sessionStorage を使う。
  assert.match(app, /persist:\s*config\.persistSessionToken === true/);
  assert.match(kitSession, /if \(!this\.persist[^)]*\) return false;/);
  assert.match(kitSession, /this\.storage\?\.setItem\(this\.storageKey/);
  assert.match(kitSession, /this\.storage\?\.removeItem\(this\.storageKey/);
  assert.match(kitSession, /globalThis\.sessionStorage/);
  // 起動時に読むJSのどこにも、localStorage への認証情報の書き込みを置かない。
  for (const source of [app, kitSession, kitIndex]) {
    assert.doesNotMatch(source, /localStorage\.setItem\([^\n]*(token|credential|auth|session)/i);
  }
});

test('完全移行した主要な教師・児童フローを含む', () => {
  for (const marker of [
    '参加申請・名簿',
    '保存して児童へ配信',
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

test('端末の戻る操作をアプリ内階層として管理する', () => {
  for (const marker of ['APP_HISTORY_ID', 'pushAppRoute', 'appBack', 'renderHistoryRoute', "window.addEventListener('popstate'", 'student-writing-focus', 'student-journal', 'teacher-feedback']) {
    assert.ok(app.includes(marker), `アプリ内履歴に「${marker}」がありません`);
  }
  assert.match(app, /history\.pushState/);
  assert.match(app, /history\.back\(\)/);
});

test('児童向け固定UIはふりがなを使い、ノートの赤い縦線を表示しない', () => {
  for (const marker of ['STUDENT_READINGS', 'studentText', "studentText('参加しているクラス')", "studentText('自分のことばで書こう')", "studentText('文章についたおへんじ')"]) {
    assert.ok(app.includes(marker), `ふりがな対応に「${marker}」がありません`);
  }
  assert.match(app, /placeholder="できたこと、かんがえたこと、つぎにやってみたいことをかこう"/);
  assert.doesNotMatch(css, /\.notebook::before/);
});

test('ふりがなは漢字にだけ付き、送り仮名には付かない', () => {
  // 文字列の一致だけでは、割り方が正しいかまでは分からない。
  // ふりがなの部分だけを実際に動かして、語彙表の全件を確かめる。
  const source = app.slice(app.indexOf('const escapeHtml'), app.indexOf('function appHistoryState'));
  const { STUDENT_READINGS, splitReading, studentText } = new Function(
    `${source}; return { STUDENT_READINGS, splitReading, studentText };`
  )();

  assert.ok(STUDENT_READINGS.length > 50, '語彙表が読めていません');
  const kana = /[\u3041-\u309f\u30a1-\u30fa\u30fc]/u;
  for (const [word, reading, parts] of STUDENT_READINGS) {
    assert.ok(splitReading(word, reading), `「${word}」を漢字と送り仮名へ割れていません`);
    assert.equal(parts.map((part) => part.text).join(''), word, `「${word}」の綴りが変わっています`);
    assert.equal(parts.map((part) => part.reading || part.text).join(''), reading, `「${word}」の読みが元と変わっています`);
    for (const part of parts) {
      if (!part.reading) continue;
      assert.doesNotMatch(part.text, kana, `「${word}」の送り仮名にふりがなが付いています`);
    }
  }

  // 代表例。送り仮名は ruby の外に出す
  assert.equal(studentText('使い方'), '<ruby>使<rt>つか</rt></ruby>い<ruby>方<rt>かた</rt></ruby>');
  assert.equal(studentText('書こう'), '<ruby>書<rt>か</rt></ruby>こう');
  assert.equal(studentText('参加している'), '<ruby>参加<rt>さんか</rt></ruby>している');
  assert.equal(studentText('もう一度'), 'もう<ruby>一度<rt>いちど</rt></ruby>');
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

test('制限付きスコープは、共有記録を読む直前にだけ求める', () => {
  // 初回は drive.file まで。粒度別同意を有効にし、Driveを外されたまま進ませない。
  assert.match(app, /const INITIAL_SCOPES = BASE_SCOPES/);
  assert.match(app, /enable_granular_consent: true/);
  assert.equal((app.match(/enable_granular_consent: true/g) || []).length, 2, '初回と追加要求の両方で粒度別同意を有効にしていません');
  assert.match(app, /if \(!state\.grantedScopes\.has\(DRIVE_FILE_SCOPE\)\)/);

  // 一覧・作成の画面からは要求しない。要求は共有記録を読む2か所だけ。
  const gates = app.match(/requireSharedRead\(/g) || [];
  assert.equal(gates.length, 3, `requireSharedRead の呼び出しが ${gates.length} 箇所あります（定義1 + 要求2 のはず）`);
  for (const name of ['openTeacherClass', 'openPortfolio']) {
    const body = app.slice(app.indexOf(`async function ${name}(`));
    assert.match(body.slice(0, 500), /if \(!state\.grantedScopes\.has\(SHARED_READ_SCOPE\)\) return requireSharedRead\(/,
      `${name} の冒頭で共有読み取りを求めていません`);
  }
  for (const name of ['renderTeacherHome', 'renderStudentHome', 'renderJoin', 'renderHome']) {
    const start = app.indexOf(`function ${name}(`);
    const body = app.slice(start, start + 3000);
    assert.doesNotMatch(body, /requireSharedRead\(/, `${name} が共有読み取りを先回りして求めています`);
  }

  // 検索式も必要な範囲へ絞る（先生が開くのは自分が作ったクラスだけ）
  assert.match(api, /listClasses\(\)\s*\{[^}]*owner: 'me'/s);
  assert.doesNotMatch(api, /listByType\(\{ type: TYPE\.class \}\)/);
});

test('教師ダッシュボードは提出確認を主役に構造化している', () => {
  for (const marker of ['teacher-dashboard-grid', 'teacher-submissions', 'teacher-sidebar', 'today-status-panel', 'teacher-settings-grid', 'vitals-grid']) {
    assert.ok(app.includes(marker) || css.includes(marker), `教師画面に「${marker}」がありません`);
  }
  assert.match(css, /\.teacher-journal-list/);
  assert.match(css, /\.teacher-overview/);
});

test('トップ画面は中央配置され、内部用語やクリックイベントを表示しない', () => {
  assert.match(app, /class="role-home"/);
  assert.match(css, /\.role-home\s*\{[^}]*margin-inline:\s*auto/s);
  assert.match(css, /\.center-screen\s*\{[^}]*justify-content:\s*center/s);
  assert.doesNotMatch(app, /Driveネイティブ/);
  assert.doesNotMatch(app, /addEventListener\('click',\s*render(?:Home|TeacherHome|StudentHome)\)/);
  assert.match(app, /typeof message === 'string'/);
});

test('ログイン後の最初の画面が役割選択からクラスまで最短でつながる', () => {
  // 役割選択カードは絵記号を持ち、児童側は低学年でも読めるようふりがなを付ける
  assert.match(app, /class="item-card role-card"/);
  assert.match(app, /class="role-icon" aria-hidden="true"/);
  assert.match(app, /studentText\('児童として使う'\)/);
  assert.match(css, /\.role-icon\s*\{/);

  // 先生ホームは、作成フォームより先に担当クラスを見せる
  const teacherHome = app.slice(app.indexOf('async function renderTeacherHome'), app.indexOf('async function createClass'));
  assert.ok(teacherHome.includes('teacher-classes'), '先生ホームにクラス一覧の区画がありません');
  assert.ok(teacherHome.includes('create-class-panel'), '先生ホームにクラス作成の区画がありません');
  assert.ok(teacherHome.indexOf('teacher-classes') < teacherHome.indexOf('create-class-panel'), 'クラス作成フォームがクラス一覧より前にあります');
  assert.match(teacherHome, /classes\.length === 1 \? classes\[0\] : null/);

  // 児童のクラスカードは、その日の提出状況とこれまでの件数を示す
  const studentHome = app.slice(app.indexOf('async function renderStudentHome'), app.indexOf('async function renderJoin'));
  assert.match(studentHome, /class-today/);
  assert.match(studentHome, /studentText\('今日のふりかえりを書こう'\)/);
  assert.match(studentHome, /studentText\('今日のふりかえりを提出しました'\)/);
  assert.match(css, /\.class-today\.done\s*\{/);

  // 参加クラスが1つだけなら一覧を飛ばして開く。開く前に旗を下ろし、失敗時の往復を作らない
  assert.match(studentHome, /portfolios\.length === 1/);
  assert.ok(studentHome.includes('state.restoreSingleStudentClass = false;'), '自動オープンの旗を下ろしていません');
  assert.ok(studentHome.includes('return openPortfolio(only)'), '1クラスのときの自動オープンがありません');
  assert.ok(
    studentHome.indexOf('state.restoreSingleStudentClass = false;') < studentHome.indexOf('return openPortfolio(only)'),
    '自動オープンの旗を下ろす前にポートフォリオを開いています'
  );

  // 戻る操作で一覧へ帰ったときは、開き直さない
  const studentRoute = app.slice(app.indexOf("if (route === 'student-home')"), app.indexOf("if (route === 'student-join')"));
  assert.match(studentRoute, /state\.restoreSingleStudentClass = false;/);
});

test('ログイン画面とヘッダーはPWAと同じアプリアイコンを使う', () => {
  assert.match(app, /class="app-logo" src="\.\/icon-192\.png"/);
  assert.match(app, /class="brand-icon" src="\.\/icon-192\.png"/);
  assert.doesNotMatch(app, /📔/);
});

test('READMEとMANUALが現在のDriveネイティブ実装を説明している', () => {
  const userDocs = `${readme}\n${manual}`;
  const technicalDocs = `${architecture}\n${oauthGuide}`;

  assert.match(readme, /現在版はGASを使いません/);
  assert.match(readme, /本番入口 `docs\/index\.html`/);
  assert.match(readme, /ブラウザがGoogle Drive REST APIへ直接通信/);
  assert.match(readme, /ルート直下の `\*\.gs`.*移行前のGAS版/s);
  assert.doesNotMatch(manual, /GAS.*デプロイ.*必要です/);

  assert.match(app, /const INITIAL_SCOPES = BASE_SCOPES/);
  assert.match(app, /requireSharedRead/);
  for (const document of [readme, manual, architecture, oauthGuide]) {
    assert.match(document, /初回|最初/);
    assert.match(document, /drive\.readonly|Drive閲覧権限|Driveの閲覧権限/);
  }
  assert.match(technicalDocs, /drive\.readonly[^\n]*段階的|段階的[^\n]*drive\.readonly/);

  assert.match(manual, /先生へもう一度届ける/);
  assert.doesNotMatch(manual, /先生へ再共有/);
  assert.match(userDocs, /児童.*参加済みクラス一覧/s);
  assert.doesNotMatch(userDocs, /前回の画面へ戻/);
  assert.match(manual, /端末の戻るジェスチャー／ボタン、ブラウザの戻るボタン/);
  assert.match(userDocs, /利用者が入力した文章.*自動.*ふりがな/s);
});

test('本体は共通部分をキットから読み、キットは本体に依存しない', () => {
  // 本体側がキットを実際に使っていること（コピーを持ち直していないこと）
  assert.match(app, /from '\.\/kit\/(index|records|session)\.js'/);
  assert.match(core, /from '\.\/kit\/(namespace|invite|records)\.js'/);
  assert.match(api, /from '\.\/kit\/drive-client\.js'/);

  // キットは他アプリへそのまま持ち出せる状態を保つ（アプリ固有の名前・文言を持たない）
  for (const [name, source] of Object.entries(kitSources)) {
    assert.doesNotMatch(source, /\.\.\/|from '\.\/(drive-core|drive-api|drive-app)/, `${name} が本体を読んでいます`);
    assert.doesNotMatch(source, /ふりかえりジャーナル|reflection-journal|rjType|rjClassId/, `${name} にアプリ固有の名前が残っています`);
  }
});

test('オフラインでもキットを含むシェル一式が配信される', () => {
  for (const asset of ['namespace.js', 'invite.js', 'drive-client.js', 'records.js', 'session.js', 'index.js']) {
    assert.ok(sw.includes(`./kit/${asset}`), `Service Workerのキャッシュ対象に kit/${asset} がありません`);
  }
  assert.match(sw, /const CACHE_PREFIX = 'rj-shell-'/);
});

test('横展開の手順書が、移し替えの前提と落とし穴を説明している', () => {
  assert.match(portingGuide, /appId[^\n]*propertyPrefix|propertyPrefix[^\n]*appId/);
  assert.match(portingGuide, /drive\.file[\s\S]*drive\.readonly/);
  // 共有が禁止された環境では動かないこと、drive.file だけでは共有同期できないことを必ず書く
  assert.match(portingGuide, /共有.*禁止|禁止.*共有/);
  assert.match(portingGuide, /drive\.file[^\n]*だけで/);
  assert.match(portingGuide, /Registry\.gs/);
  assert.match(portingGuide, /version/);
  assert.match(kitReadme, /createDriveNativeApp/);
  assert.match(readme, /PORTING_FROM_GAS\.md/);
});
