import {
  analyzeClass,
  appendJournal,
  computeClassId,
  createChannel,
  createClassRecord,
  createPortfolio,
  currentTheme,
  decodeInvite,
  encodeSignedInvite,
  ensureInviteSecurity,
  exportCsv,
  isEmail,
  mergePortfoliosIntoMembers,
  normalizeClassCode,
  normalizeEmail,
  randomClassCode,
  rotateInviteSecurity,
  setFeedback,
  signedInviteUrl,
  studentKey,
  syncChannel,
  updatePastComment,
  validateChannelForStudent,
  validatePortfolioForClass
} from './drive-core.js';
import { DriveApiError, DriveClient } from './drive-api.js';
import { BASE_SCOPES, SHARED_READ_SCOPE } from './kit/index.js';
import { RecordCache } from './kit/records.js';
import { ScopeGrant, SessionPolicy, tokenExpiryFrom } from './kit/session.js';

const app = document.getElementById('app');
const toastElement = document.getElementById('toast');
const config = window.APP_CONFIG || {};
const INITIAL_SCOPES = BASE_SCOPES;
const LAST_ROLE_KEY = 'rj_last_role';
const state = {
  user: null,
  drive: null,
  accessToken: '',
  tokenExpiresAt: 0,
  forceAccountSelection: false,
  invite: null,
  tokenClient: null,
  sharedTokenClient: null,
  grantedScopes: new ScopeGrant(),
  pendingSharedAction: null,
  pendingSharedRole: '',
  teacher: null,
  portfolio: null,
  portfolioFile: null,
  channel: null,
  channelFile: null,
  studentMonth: new Date(),
  studentPortfolios: [],
  teacherFilter: 'all',
  teacherSearch: '',
  teacherClasses: [],
  restoreLastTeacherClass: false,
  restoreSingleStudentClass: false,
  feedbackDraftKey: '',
  feedbackDraftHighlights: [],
  teacherTab: 'journals',
  teacherRefreshTimer: null,
  jsonCache: new RecordCache(),
  studentDrafts: new Map(),
  rejectedPortfolios: []
};

// 公開元・ドメイン・トークン保存の方針はキットへ寄せる（他アプリでも同じ規則を使う）。
const sessionPolicy = new SessionPolicy({
  clientId: config.googleClientId,
  storageKey: 'rj_oauth_session_v1',
  allowedOrigins: Array.isArray(config.allowedOrigins) ? config.allowedOrigins : [],
  allowedDomains: Array.isArray(config.allowedWorkspaceDomains) ? config.allowedWorkspaceDomains : [],
  persist: config.persistSessionToken === true
});
const originSecurityError = sessionPolicy.originAllowed(location.origin, location.hostname)
  ? ''
  : 'この公開元はアプリの許可リストにありません。正しい学校用URLから開いてください。';

function allowedWorkspaceDomain(email) {
  return sessionPolicy.domainAllowed(email);
}

const escapeHtml = (value) => String(value ?? '').replace(/[&<>'"]/g, (char) => ({
  '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;'
})[char]);
const todayKey = (date = new Date()) => [date.getFullYear(), String(date.getMonth() + 1).padStart(2, '0'), String(date.getDate()).padStart(2, '0')].join('-');
const ruby = (word, reading) => `<ruby>${escapeHtml(word)}<rt>${escapeHtml(reading)}</rt></ruby>`;
const KANJI_RUN = /[\u4e00-\u9faf\u3005]+|[^\u4e00-\u9faf\u3005]+/g;
const IS_KANJI = /[\u4e00-\u9faf\u3005]/;

// 「使い方（つかいかた）」を 使(つか)い方(かた) に割る。ふりがなは漢字にだけ付け、
// 送り仮名には付けない。送り仮名は読みの中にもそのまま現れるので、その位置を手がかりに
// 漢字の読みを切り出せる。割れない語（読みと綴りが対応しない当て字など）は null を返し、
// 語まるごとのふりがなへ落とす。
function splitReading(word, reading) {
  const runs = word.match(KANJI_RUN) || [];
  const parts = [];
  let rest = reading;
  for (let index = 0; index < runs.length; index += 1) {
    const run = runs[index];
    if (!IS_KANJI.test(run)) {
      if (!rest.startsWith(run)) return null;
      rest = rest.slice(run.length);
      parts.push({ text: run });
      continue;
    }
    const following = runs[index + 1];
    if (!following) {
      if (!rest) return null;
      parts.push({ text: run, reading: rest });
      rest = '';
      continue;
    }
    const at = rest.indexOf(following, 1);   // 漢字には最低1文字の読みを残す
    if (at < 1) return null;
    parts.push({ text: run, reading: rest.slice(0, at) });
    rest = rest.slice(at);
  }
  return rest ? null : parts;
}

const rubyParts = (parts) => parts.map((part) => (part.reading ? ruby(part.text, part.reading) : escapeHtml(part.text))).join('');
const APP_HISTORY_ID = 'reflection-journal';
const STUDENT_READINGS = [
  ['参加している', 'さんかしている'], ['参加済み', 'さんかずみ'], ['参加申請', 'さんかしんせい'], ['参加状況', 'さんかじょうきょう'],
  ['受け付けて', 'うけつけて'], ['読み込んで', 'よみこんで'], ['書き出し', 'かきだし'], ['書き方', 'かきかた'],
  ['使い方', 'つかいかた'], ['使う', 'つかう'], ['新しい', 'あたらしい'], ['表示できません', 'ひょうじできません'], ['表示する', 'ひょうじする'], ['確認できる', 'かくにんできる'], ['探しています', 'さがしています'], ['開いて', 'ひらいて'],
  ['もう一度', 'もういちど'], ['過去', 'かこ'], ['自分', 'じぶん'], ['今日', 'きょう'], ['明日', 'あした'],
  ['先生', 'せんせい'], ['児童', 'じどう'], ['招待', 'しょうたい'], ['専用', 'せんよう'], ['名前', 'なまえ'],
  ['一覧', 'いちらん'], ['戻る', 'もどる'], ['戻す', 'もどす'], ['変える', 'かえる'], ['件', 'けん'], ['届けます', 'とどけます'], ['届ける', 'とどける'], ['届いて', 'とどいて'], ['届きます', 'とどきます'],
  ['参加', 'さんか'], ['承認', 'しょうにん'], ['申請', 'しんせい'], ['提出しました', 'ていしゅつしました'], ['提出', 'ていしゅつ'], ['送る', 'おくる'], ['送りました', 'おくりました'],
  ['記録', 'きろく'], ['書いた日', 'かいたひ'], ['書ける', 'かける'], ['書こう', 'かこう'], ['書いて', 'かいて'], ['書く', 'かく'], ['書き', 'かき'],
  ['考えた', 'かんがえた'], ['気持ち', 'きもち'], ['気持', 'きも'], ['気づき', 'きづき'], ['学び', 'まなび'],
  ['広く', 'ひろく'], ['作品', 'さくひん'], ['写真', 'しゃしん'], ['画像', 'がぞう'], ['任意', 'にんい'],
  ['見やすい', 'みやすい'], ['大きさ', 'おおきさ'], ['共有', 'きょうゆう'], ['学校側', 'がっこうがわ'], ['拒否', 'きょひ'],
  ['完了', 'かんりょう'], ['設定', 'せってい'], ['確認', 'かくにん'], ['文章', 'ぶんしょう'], ['注目', 'ちゅうもく'], ['今', 'いま'],
  ['次', 'つぎ'], ['前', 'まえ'], ['見る', 'みる'], ['待っています', 'まっています'], ['閉じる', 'とじる'],
  ['対話', 'たいわ'], ['保存', 'ほぞん'], ['文字', 'もじ'], ['文', 'ぶん'], ['入ります', 'はいります'],
  ['心', 'こころ'], ['残った', 'のこった'], ['初めて', 'はじめて'], ['瞬間', 'しゅんかん'], ['新しく', 'あたらしく'],
  ['調べたい', 'しらべたい'], ['知りたい', 'しりたい'], ['友だち', 'ともだち'], ['思った', 'おもった'], ['言いたい', 'いいたい'], ['言う', 'いう'], ['協力', 'きょうりょく'],
  ['年間', 'ねんかん'], ['年', 'ねん'], ['組', 'くみ'], ['月', 'がつ'], ['日', 'にち'], ['火', 'か'], ['水', 'すい'], ['木', 'もく'], ['金', 'きん'], ['土', 'ど']
].sort((a, b) => b[0].length - a[0].length)
  .map(([word, reading]) => [word, reading, splitReading(word, reading) || [{ text: word, reading }]]);

function studentText(value) {
  const source = String(value ?? '');
  let result = '';
  let plain = '';
  const flush = () => { result += escapeHtml(plain); plain = ''; };
  for (let index = 0; index < source.length;) {
    const match = STUDENT_READINGS.find(([word]) => source.startsWith(word, index));
    if (!match) { plain += source[index]; index += 1; continue; }
    flush();
    result += rubyParts(match[2]);
    index += match[0].length;
  }
  flush();
  return result;
}

function appHistoryState() {
  return history.state?.app === APP_HISTORY_ID ? history.state : null;
}

function replaceAppRoute(route, data = {}, url = location.href) {
  const current = appHistoryState();
  history.replaceState({ app: APP_HISTORY_ID, route, data, depth: current?.depth || 0 }, '', url);
}

function resetAppRoute(route = 'home', data = {}, url = location.href) {
  history.replaceState({ app: APP_HISTORY_ID, route, data, depth: 0 }, '', url);
}

function pushAppRoute(route, data = {}) {
  const current = appHistoryState();
  if (current?.route === route && JSON.stringify(current.data || {}) === JSON.stringify(data)) return;
  history.pushState({ app: APP_HISTORY_ID, route, data, depth: (current?.depth || 0) + 1 }, '', location.href);
}

function appBack(fallback) {
  const current = appHistoryState();
  if (current && current.depth > 0) history.back();
  else fallback?.();
}

function closeStudentDialog() {
  const dialog = document.querySelector('dialog.journal-dialog');
  if (!dialog) return;
  dialog.dataset.historyClose = 'true';
  if (dialog.open) dialog.close();
  else dialog.remove();
}

async function renderHistoryRoute(entry) {
  closeStudentDialog();
  setWritingFocus(false);
  const route = entry?.route || 'home';
  const data = entry?.data || {};
  if (!state.user || !state.drive) return renderLogin();
  if (route === 'home') {
    forgetRole();
    state.restoreLastTeacherClass = false;
    state.restoreSingleStudentClass = false;
    return renderHome();
  }
  if (route === 'teacher-home') {
    rememberRole('teacher');
    state.restoreLastTeacherClass = false;
    return requireSharedRead('teacher', () => renderTeacherHome());
  }
  if (route === 'teacher-class') {
    rememberRole('teacher');
    const current = state.teacher?.record?.classId === data.classId ? { file: state.teacher.file, record: state.teacher.record } : null;
    const item = current || state.teacherClasses.find(({ record }) => record.classId === data.classId);
    return item ? openTeacherClass(item, data.tab || 'journals') : renderTeacherHome();
  }
  if (route === 'teacher-feedback') {
    const portfolioItem = state.teacher?.portfolios?.find(({ record }) => normalizeEmail(record.student.email) === normalizeEmail(data.email));
    const journal = portfolioItem?.record?.journals?.find((item) => item.id === data.journalId);
    return portfolioItem && journal ? renderFeedbackEditor(portfolioItem, journal) : renderTeacherClass('journals');
  }
  if (route === 'student-home') {
    rememberRole('student');
    state.restoreSingleStudentClass = false;
    state.invite = null;
    if (/^#join=/.test(location.hash)) history.replaceState(entry, '', location.pathname + location.search);
    return requireSharedRead('student', () => renderStudentHome());
  }
  if (route === 'student-join') return state.invite ? renderJoin() : renderStudentHome();
  if (['student-portfolio', 'student-writing-focus', 'student-journal'].includes(route)) {
    const item = state.portfolio?.class?.id === data.classId
      ? { file: state.portfolioFile, record: state.portfolio }
      : state.studentPortfolios?.find(({ file, record }) => file.id === data.fileId || record.class.id === data.classId);
    if (!item) return renderStudentHome();
    if (state.portfolio?.class?.id !== item.record.class.id) await openPortfolio(item);
    else renderPortfolio();
    if (route === 'student-writing-focus') setWritingFocus(true);
    if (route === 'student-journal') openStudentJournal(data.journalId);
    return;
  }
  return renderHome();
}

const JOURNAL_TEMPLATES = [
  ['YWT', '【Y: やったこと】\n\n【W: わかったこと】\n\n【T: つぎにやること】\n'],
  ['KPT', '【K: Keep よかったこと・つづけること】\n\n【P: Problem こまったこと・やめること】\n\n【T: Try つぎにためすこと】\n'],
  ['5W1H', '【いつ】\n\n【どこで】\n\n【だれが】\n\n【なにを】\n\n【なぜ】\n\n【どのように】\n'],
  ['3つの気づき', '【わかったこと】\n\n【おどろいたこと】\n\n【もっとしりたいこと】\n']
];
const RANDOM_STARTERS = ['今日いちばん心にのこったのは、', '今日、はじめてできたことは、', '今日うれしかったのは、', '明日がんばりたいことは、'];
const HINT_GROUPS = [
  ['気持ち', ['今日、いちばんうれしかったことは、', 'くやしかったこと・かなしかったことは、', 'ドキドキ・ワクワクした瞬間は、']],
  ['学び', ['新しくわかったこと・できたことは、', '「なぜだろう？」と思ったことは、', 'もっと調べたい・知りたいことは、']],
  ['つながり', ['友だちに「ありがとう」と言いたいことは、', 'だれかの「すごいな」と思ったところは、', 'みんなで協力してできたことは、']],
  ['次へ', ['明日がんばりたいことは、', 'もう一度やるなら、どうする？', '今日の自分にひとこと言うなら、']]
];

function toast(message) {
  toastElement.textContent = message;
  toastElement.classList.add('visible');
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => toastElement.classList.remove('visible'), 3200);
}

function studentToast(message) {
  toastElement.innerHTML = studentText(message);
  toastElement.classList.add('visible');
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => toastElement.classList.remove('visible'), 3200);
}

function setBusy(message = 'データを読み込んでいます…') {
  app.innerHTML = `<section class="center-screen"><div class="loader" aria-hidden="true"></div><p>${escapeHtml(message)}</p></section>`;
}

function setStudentBusy(message) {
  app.innerHTML = `<section class="center-screen student-ui"><div class="loader" aria-hidden="true"></div><p>${studentText(message)}</p></section>`;
}

function friendlyError(error) {
  console.error(error);
  if (error instanceof DriveApiError) return error.message;
  return error?.message || '操作を完了できませんでした。もう一度お試しください。';
}

function rememberGrantedScopes(response, required = []) {
  state.grantedScopes.remember(response, required, window.google?.accounts?.oauth2);
}

function errorNotice(message) {
  const text = typeof message === 'string' ? message.trim() : '';
  return text ? `<div class="error" role="alert">${escapeHtml(text)}</div>` : '';
}

function studentErrorNotice(message) {
  const text = typeof message === 'string' ? message.trim() : '';
  return text ? `<div class="error" role="alert">${studentText(text)}</div>` : '';
}

async function withError(action, fallback) {
  try { return await action(); }
  catch (error) {
    if (error instanceof DriveApiError && error.status === 401) {
      clearSession();
      return renderLogin(error.message);
    }
    fallback(friendlyError(error));
    return null;
  }
}

function shell(content) {
  return `<header class="topbar">
    <div class="brand"><img class="brand-icon" src="./icon-192.png" alt="" aria-hidden="true"><span>ふりかえりジャーナル</span></div>
    <button class="account" data-change-account type="button" title="Googleアカウントを変更"><strong>${escapeHtml(state.user?.name || '')}</strong><span>${escapeHtml(state.user?.email || '')}</span><small>アカウントを変更</small></button>
  </header><div class="page">${content}</div>`;
}

function renderLogin(error = '') {
  const blockingError = originSecurityError || error;
  app.innerHTML = `<section class="center-screen"><div class="login-card">
    <img class="app-logo" src="./icon-192.png" alt="" aria-hidden="true">
    <h1>毎日のふりかえりを<br>学びの成長へ</h1>
    <p>自分の言葉で学びを残し、先生からのおへんじを受け取れます。</p>
    ${errorNotice(blockingError)}
    <button id="login" class="primary wide" type="button" ${originSecurityError ? 'disabled' : ''}>Googleアカウントで続ける</button>
    <p class="muted small">${config.persistSessionToken === true ? '同じタブでは有効期限までログイン状態を保ちます。タブを閉じると認証情報は消去されます。' : '認証情報を端末の保存領域に残しません。ページ更新後は再ログインが必要です。'}</p>
  </div></section>`;
  document.getElementById('login').addEventListener('click', requestAccess);
}

function loadGoogleIdentity() {
  return new Promise((resolve, reject) => {
    if (window.google?.accounts?.oauth2) return resolve();
    const script = document.createElement('script');
    script.src = 'https://accounts.google.com/gsi/client';
    script.async = true;
    script.onload = resolve;
    script.onerror = () => reject(new Error('Googleログインを読み込めませんでした。学校のネットワーク設定を確認してください。'));
    document.head.appendChild(script);
  });
}

async function requestAccess() {
  if (originSecurityError) return renderLogin();
  if (!config.googleClientId) return renderLogin('ログインの準備が完了していません。アプリ管理者に連絡してください。');
  setBusy('Googleログインをひらいています…');
  try {
    await loadGoogleIdentity();
    state.tokenClient ||= google.accounts.oauth2.initTokenClient({
      client_id: config.googleClientId,
      scope: INITIAL_SCOPES,
      include_granted_scopes: true,
      callback: handleToken
    });
    state.tokenClient.requestAccessToken({ prompt: state.forceAccountSelection ? 'select_account' : '' });
  } catch (error) { renderLogin(error.message); }
}

async function handleToken(response) {
  if (!response?.access_token) return renderLogin(response?.error_description || 'Googleログインがキャンセルされました。');
  rememberGrantedScopes(response, INITIAL_SCOPES.split(' '));
  state.accessToken = response.access_token;
  state.tokenExpiresAt = tokenExpiryFrom(response);
  state.forceAccountSelection = false;
  setBusy('アカウントを確認しています…');
  try {
    const userResponse = await fetch('https://openidconnect.googleapis.com/v1/userinfo', { headers: { Authorization: `Bearer ${response.access_token}` } });
    if (!userResponse.ok) throw new Error('Googleアカウント情報を確認できませんでした。');
    const profile = await userResponse.json();
    if (!allowedWorkspaceDomain(profile.email)) {
      clearSession();
      return renderLogin('このGoogleアカウントのドメインは、アプリ管理者から許可されていません。');
    }
    state.user = { email: normalizeEmail(profile.email), name: profile.name || profile.email };
    state.drive = new DriveClient(response.access_token);
    saveSession();
    await resolveEntryRoute();
  } catch (error) { renderLogin(error.message); }
}

async function resolveEntryRoute() {
  const encoded = location.hash.match(/^#join=([^&]+)/)?.[1];
  if (encoded) {
    try {
      state.invite = await decodeInvite(encoded);
      rememberRole('student');
      if (appHistoryState()?.route === 'home') pushAppRoute('student-home');
      else if (appHistoryState()?.route !== 'student-home' && appHistoryState()?.route !== 'student-join') replaceAppRoute('student-home');
      pushAppRoute('student-join');
      return requireSharedRead('student', () => renderJoin());
    }
    catch (error) { return renderHome(error.message); }
  }
  const role = preferredRole();
  if (role === 'teacher') {
    state.restoreLastTeacherClass = true;
    if (appHistoryState()?.route === 'home') pushAppRoute('teacher-home');
    else if (appHistoryState()?.route !== 'teacher-home') replaceAppRoute('teacher-home');
    return requireSharedRead('teacher', () => renderTeacherHome());
  }
  if (role === 'student') {
    state.restoreSingleStudentClass = true;
    if (appHistoryState()?.route === 'home') pushAppRoute('student-home');
    else if (appHistoryState()?.route !== 'student-home') replaceAppRoute('student-home');
    return requireSharedRead('student', () => renderStudentHome());
  }
  replaceAppRoute('home');
  return renderHome();
}

function saveSession() {
  sessionPolicy.save({
    accessToken: state.accessToken,
    expiresAt: state.tokenExpiresAt,
    scopes: state.grantedScopes.list(),
    user: state.user
  });
}

function restoreSession() {
  if (originSecurityError) return false;
  const saved = sessionPolicy.restore();
  if (!saved) return false;
  state.accessToken = saved.accessToken;
  state.tokenExpiresAt = saved.expiresAt;
  state.user = saved.user;
  state.grantedScopes = new ScopeGrant(saved.scopes);
  state.drive = new DriveClient(saved.accessToken);
  return true;
}

function clearSession() {
  sessionPolicy.clear();
  state.accessToken = '';
  state.tokenExpiresAt = 0;
  state.user = null;
  state.drive = null;
  state.grantedScopes = new ScopeGrant();
}

function preferredRole() {
  try { return localStorage.getItem(LAST_ROLE_KEY) || ''; } catch (error) { return ''; }
}

function rememberRole(role) {
  try { localStorage.setItem(LAST_ROLE_KEY, role); } catch (error) {}
}

function forgetRole() {
  try { localStorage.removeItem(LAST_ROLE_KEY); } catch (error) {}
}

function requireSharedRead(role, action) {
  if (state.grantedScopes.has(SHARED_READ_SCOPE)) return action();
  state.pendingSharedRole = role;
  state.pendingSharedAction = action;
  renderSharedReadPermission(role);
}

function renderSharedReadPermission(role, error = '') {
  const isTeacher = role === 'teacher';
  app.innerHTML = shell(`<section class="permission-page"><div class="permission-card panel">
    <div class="permission-icon" aria-hidden="true">🔄</div><span class="eyebrow">SHARED RECORDS</span>
    <h1>${isTeacher ? '児童の提出を受け取る準備' : '先生のおへんじを受け取る準備'}</h1>
    <p>${isTeacher ? '別のアカウントから共有されたふりかえり' : '先生のアカウントから共有されたテーマやおへんじ'}を自動で見つけるため、Google Driveの閲覧許可が必要です。</p>
    ${errorNotice(error)}
    <div class="permission-points"><p><strong>許可の範囲</strong><br><span>Googleの許可画面ではDrive全体の閲覧権限です。アプリの現在の実装は、共有済みで専用の印が付いた記録だけを検索・表示します。</span></p><p><strong>書き換えは専用ファイルだけ</strong><br><span>書込みには、アプリが作成したファイルだけを扱う権限を使います。通常の記録を運営者サーバーへ保存しません。</span></p></div>
    <button id="grant-shared-read" class="primary wide" type="button">共有された記録の同期を許可する</button>
    <button id="permission-back" class="quiet wide" type="button">使い方の選択へ戻る</button>
  </div></section>`);
  document.getElementById('grant-shared-read').addEventListener('click', requestSharedRead);
  document.getElementById('permission-back').addEventListener('click', () => {
    state.pendingSharedAction = null;
    state.pendingSharedRole = '';
    appBack(() => { forgetRole(); replaceAppRoute('home'); renderHome(); });
  });
}

async function requestSharedRead() {
  setBusy('Google Driveの共有記録を同期する準備をしています…');
  try {
    await loadGoogleIdentity();
    state.sharedTokenClient ||= google.accounts.oauth2.initTokenClient({
      client_id: config.googleClientId,
      scope: SHARED_READ_SCOPE,
      include_granted_scopes: true,
      enable_granular_consent: true,
      callback: handleSharedToken
    });
    state.sharedTokenClient.requestAccessToken({ prompt: '' });
  } catch (error) { renderSharedReadPermission(state.pendingSharedRole, friendlyError(error)); }
}

async function handleSharedToken(response) {
  if (!response?.access_token) return renderSharedReadPermission(state.pendingSharedRole, response?.error_description || '共有記録の同期が許可されませんでした。');
  rememberGrantedScopes(response, [SHARED_READ_SCOPE]);
  if (!state.grantedScopes.has(SHARED_READ_SCOPE)) return renderSharedReadPermission(state.pendingSharedRole, 'Google Driveの閲覧許可が選択されていません。共有記録を受け取るには、この許可が必要です。');
  state.accessToken = response.access_token;
  state.tokenExpiresAt = tokenExpiryFrom(response);
  state.drive = new DriveClient(response.access_token);
  saveSession();
  const action = state.pendingSharedAction;
  state.pendingSharedAction = null;
  state.pendingSharedRole = '';
  if (action) await action();
  else renderHome();
}

function renderHome(error = '') {
  // ログイン直後に最初に見る画面。児童側のカードは、低学年でも読めるようふりがなを付ける。
  app.innerHTML = shell(`<section class="role-home"><div class="role-home-intro"><h1>どの使い方をしますか？</h1><p>あなたに合った入口を選んでください。あとから［使い方を変える］で切り替えられます。</p></div>
    ${errorNotice(error)}
    <div class="grid role-choice-grid">
      <button class="item-card role-card" id="teacher-home" type="button" aria-label="先生として使う。クラス作成、招待、返却、分析、名簿とテーマを管理します"><span class="role-icon" aria-hidden="true">🧑‍🏫</span><h2>先生として使う</h2><p class="muted">クラス作成、招待、返却、分析、名簿とテーマを管理します。</p></button>
      <button class="item-card role-card student-ui" id="student-home" type="button" aria-label="児童として使う。参加したクラスで書き、先生からのおへんじを受け取ります"><span class="role-icon" aria-hidden="true">🎒</span><h2>${studentText('児童として使う')}</h2><p class="muted">${studentText('クラスで書いて、先生からのおへんじが届きます。')}</p></button>
    </div></section>`);
  document.getElementById('teacher-home').addEventListener('click', () => { rememberRole('teacher'); state.restoreLastTeacherClass = true; pushAppRoute('teacher-home'); requireSharedRead('teacher', () => renderTeacherHome()); });
  document.getElementById('student-home').addEventListener('click', () => { rememberRole('student'); state.restoreSingleStudentClass = true; pushAppRoute('student-home'); requireSharedRead('student', () => renderStudentHome()); });
}

async function renderTeacherHome(error = '') {
  setBusy('先生のクラスを探しています…');
  const files = await withError(() => state.drive.listClasses(), (message) => renderHome(message));
  if (!files) return;
  const classes = (await loadJsonItems(files)).filter(({ file, record }) =>
    record?.kind === 'reflection-journal-class'
    && normalizeEmail(record.teacher?.email) === state.user.email
    && normalizeEmail(file.owners?.[0]?.emailAddress) === state.user.email);
  state.teacherClasses = classes;
  if (state.restoreLastTeacherClass && classes.length) {
    state.restoreLastTeacherClass = false;
    let last = '';
    try { last = localStorage.getItem('rj_last_teacher_class') || ''; } catch (storageError) {}
    // 最後に開いたクラスが分からなくても、1クラスしか無ければ迷う余地はない。そのまま開く。
    const match = classes.find(({ file }) => file.id === last) || (classes.length === 1 ? classes[0] : null);
    if (match) { pushAppRoute('teacher-class', { classId: match.record.classId, tab: 'journals' }); return openTeacherClass(match); }
  }
  state.restoreLastTeacherClass = false;
  // 毎日の用事は「担当クラスを開く」ほう。作成フォームは一覧の後ろへ置く。
  app.innerHTML = shell(`<div class="page-heading"><div><span class="badge">先生</span><h1>クラス</h1></div><button id="back" class="quiet" type="button">使い方を変える</button></div>
    ${errorNotice(error)}
    ${classes.length ? `<section class="teacher-classes"><div class="compact-section-heading"><h2>作成済みのクラス</h2><span class="muted small">${classes.length}クラス</span></div><div class="grid">${classes.map(({ file, record }) => {
      const active = (record.members || []).filter((member) => member.status === 'active').length;
      return `<button class="item-card class-item" data-file="${escapeHtml(file.id)}" type="button" aria-label="${escapeHtml(record.className)}（クラスコード ${escapeHtml(record.classCode)}）を開く。参加中 ${active}人"><span class="badge">${escapeHtml(record.classCode)}</span><h3>${escapeHtml(record.className)}</h3><p>参加中 ${active}人</p><p class="muted small">更新: ${escapeHtml(formatDate(file.modifiedTime))}</p></button>`;
    }).join('')}</div></section>` : ''}
    <section class="panel create-class-panel"><h2>${classes.length ? 'クラスをもう1つ作る' : '最初のクラスを作る'}</h2>
      <p class="muted small">${classes.length ? '学年やクラスごとに、いくつでも作れます。' : 'クラス名を入れると、招待用のクラスコードとQRコードが作られます。'}</p>
      <form id="create-class">
      <label><span>クラス名</span><input id="class-name" maxlength="80" required placeholder="例：5年1組"></label>
      <button class="primary" type="submit">クラスを作成</button>
    </form></section>`);
  document.getElementById('back').addEventListener('click', () => appBack(() => { forgetRole(); replaceAppRoute('home'); renderHome(); }));
  document.getElementById('create-class').addEventListener('submit', createClass);
  document.querySelectorAll('.class-item').forEach((button) => button.addEventListener('click', () => {
    const item = classes.find(({ file }) => file.id === button.dataset.file);
    pushAppRoute('teacher-class', { classId: item.record.classId, tab: 'journals' });
    openTeacherClass(item);
  }));
}

async function createClass(event) {
  event.preventDefault();
  const className = document.getElementById('class-name').value.trim();
  if (!className) return;
  setBusy('クラスを作成しています…');
  await withError(async () => {
    const classCode = randomClassCode();
    const classId = await computeClassId(state.user.email, classCode);
    const secured = await ensureInviteSecurity(createClassRecord({ classId, classCode, className, teacher: state.user }));
    const record = secured.record;
    const file = await state.drive.createClass(record);
    rememberJson(file, record);
    pushAppRoute('teacher-class', { classId: record.classId, tab: 'journals' });
    await openTeacherClass({ file, record });
  }, (message) => renderTeacherHome(message));
}

function rememberJson(file, record) {
  state.jsonCache.remember(file, record);
}

async function loadJsonItem(file) {
  const cached = state.jsonCache.read(file);
  if (cached !== null) return { file, record: cached };
  const record = await state.drive.getJson(file.id);
  rememberJson(file, record);
  return { file, record };
}

async function loadJsonItems(files) {
  return (await Promise.all(files.map(async (file) => {
    try { return await loadJsonItem(file); } catch (error) { return null; }
  }))).filter(Boolean);
}

async function updateDocument(file, record) {
  const updated = await state.drive.updateJson(file.id, record, { expectedVersion: file.version || '' });
  Object.assign(file, updated || {});
  rememberJson(file, record);
  return updated;
}

async function persistTeacherChannel(email, channel) {
  const normalized = normalizeEmail(email);
  const file = state.teacher.channelFiles.get(normalized);
  if (!file) throw new Error('おへんじ用ファイルが見つかりません。最新の状態に更新してください。');
  await updateDocument(file, channel);
  state.teacher.channels.set(normalized, channel);
}

async function openTeacherClass(item, tab = 'journals', error = '', options = {}) {
  if (!options.silent) setBusy('クラスのデータを読み込んでいます…');
  const previousCount = state.teacher?.record?.classId === item.record.classId ? analyzeClass(activePortfolioItems(), state.teacher.channels).all.length : 0;
  await withError(async () => {
    const security = await ensureInviteSecurity(item.record);
    let record = security.record;
    const inviteValidationRecord = security.changed ? item.record : record;
    if (security.changed) await updateDocument(item.file, record);
    const original = JSON.stringify(record);
    const portfolioFiles = await state.drive.listSharedPortfolios(record.classId);
    const candidates = await loadJsonItems(portfolioFiles);
    const portfolios = [];
    const rejectedPortfolios = [];
    for (const candidate of candidates) {
      const validation = await validatePortfolioForClass(candidate.file, candidate.record, inviteValidationRecord);
      if (validation.ok) portfolios.push(candidate);
      else rejectedPortfolios.push({ file: candidate.file, reason: validation.reason });
    }
    record = mergePortfoliosIntoMembers(record, portfolios);
    const ownChannelFiles = await state.drive.listOwnChannels(record.classId);
    const channelFilesById = new Map(ownChannelFiles.map((file) => [file.id, file]));
    const channels = new Map();
    const channelFiles = new Map();
    for (const member of record.members || []) {
      if (member.status !== 'active') continue;
      let channel;
      let channelFile = channelFilesById.get(member.channelFileId);
      if (channelFile) {
        try {
          const loaded = await loadJsonItem(channelFile);
          if (loaded.record?.kind === 'reflection-journal-channel'
            && loaded.record.class?.id === record.classId
            && normalizeEmail(loaded.record.student?.email) === normalizeEmail(member.email)
            && normalizeEmail(loaded.record.teacher?.email) === state.user.email) channel = loaded.record;
        } catch (loadError) {}
      }
      if (!channel && member.portfolioFileId) {
        channel = createChannel({ classRecord: record, member });
        channelFile = await state.drive.createChannel(channel, await studentKey(member.email));
        await state.drive.shareWithUser(channelFile.id, member.email, 'reader');
        member.channelFileId = channelFile.id;
        rememberJson(channelFile, channel);
      }
      if (channel && channelFile) {
        const email = normalizeEmail(member.email);
        channels.set(email, channel);
        channelFiles.set(email, channelFile);
      }
    }
    if (JSON.stringify(record) !== original) await updateDocument(item.file, record);
    state.rejectedPortfolios = rejectedPortfolios;
    state.teacher = { file: item.file, record, portfolios, channels, channelFiles, rejectedPortfolios, syncedAt: new Date() };
    try { localStorage.setItem('rj_last_teacher_class', item.file.id); } catch (storageError) {}
    renderTeacherClass(tab, error);
    const nextCount = analyzeClass(activePortfolioItems(), channels).all.length;
    if (options.silent && nextCount > previousCount) toast(`新しい提出を${nextCount - previousCount}件受け取りました。`);
  }, (message) => renderTeacherHome(message));
}

function teacherTabs(active) {
  const pending = (state.teacher.record.members || []).filter((member) => member.status === 'pending').length;
  return `<nav class="tabs teacher-tabs" aria-label="クラスメニュー">${[
    ['journals', 'ジャーナル管理'], ['vitals', '心のバイタル'], ['settings', `クラス設定${pending ? ` (${pending})` : ''}`]
  ].map(([id, label]) => `<button class="tab ${active === id ? 'active' : ''}" data-tab="${id}" type="button">${label}</button>`).join('')}</nav>`;
}

function bindTeacherTabs() {
  document.querySelectorAll('[data-tab]').forEach((button) => button.addEventListener('click', () => {
    replaceAppRoute('teacher-class', { classId: state.teacher.record.classId, tab: button.dataset.tab });
    renderTeacherClass(button.dataset.tab);
  }));
  document.getElementById('classes')?.addEventListener('click', () => appBack(() => { replaceAppRoute('teacher-home'); renderTeacherHome(); }));
  document.getElementById('refresh-class')?.addEventListener('click', () => refreshTeacherClass(false));
  document.getElementById('class-switch')?.addEventListener('change', (event) => {
    if (event.target.value === 'new') return appBack(() => { replaceAppRoute('teacher-home'); renderTeacherHome(); });
    const next = state.teacherClasses.find(({ file }) => file.id === event.target.value);
    if (next) {
      replaceAppRoute('teacher-class', { classId: next.record.classId, tab: 'journals' });
      openTeacherClass(next);
    }
  });
}

function renderTeacherClass(tab = 'journals', error = '') {
  const ctx = state.teacher;
  state.teacherTab = tab;
  app.innerHTML = shell(`<div class="page-heading teacher-class-heading"><div class="teacher-class-title"><span class="badge">${escapeHtml(ctx.record.classCode)}</span><h1>${escapeHtml(ctx.record.className)}</h1><span class="sync-status">最終同期 ${escapeHtml(formatTime(ctx.syncedAt))}</span></div><div class="class-switcher"><label for="class-switch">表示するクラス</label><select id="class-switch">${state.teacherClasses.map(({ file, record }) => `<option value="${escapeHtml(file.id)}" ${file.id === ctx.file.id ? 'selected' : ''}>${escapeHtml(record.className)}</option>`).join('')}<option value="new">＋ クラス一覧・新規作成</option></select><button id="refresh-class" class="secondary" type="button">↻ 最新に更新</button><button id="classes" class="quiet" type="button">一覧</button></div></div>
    ${errorNotice(error)}${teacherTabs(tab)}<div id="teacher-content"></div>`);
  bindTeacherTabs();
  if (tab === 'journals') renderJournalManagement();
  else if (tab === 'vitals') renderVitals();
  else renderClassSettings();
  scheduleTeacherRefresh();
}

async function refreshTeacherClass(silent = true) {
  const ctx = state.teacher;
  if (!ctx?.file?.id) return;
  if (!silent) setBusy('最新の提出を確認しています…');
  await withError(async () => {
    const [file, record] = await Promise.all([state.drive.getMetadata(ctx.file.id), state.drive.getJson(ctx.file.id)]);
    rememberJson(file, record);
    await openTeacherClass({ file, record }, state.teacherTab, '', { silent });
  }, (message) => renderTeacherClass(state.teacherTab, message));
}

function scheduleTeacherRefresh() {
  clearTimeout(state.teacherRefreshTimer);
  if (state.teacherTab !== 'journals') return;
  state.teacherRefreshTimer = setTimeout(async () => {
    if (!document.getElementById('teacher-content')) return;
    const activeTag = document.activeElement?.tagName;
    if (document.visibilityState !== 'visible' || ['INPUT', 'TEXTAREA', 'SELECT'].includes(activeTag)) return scheduleTeacherRefresh();
    await refreshTeacherClass(true);
  }, 30_000);
}

function activePortfolioItems() {
  const active = new Set((state.teacher.record.members || []).filter((member) => member.status === 'active').map((member) => normalizeEmail(member.email)));
  return state.teacher.portfolios.filter(({ record }) => active.has(normalizeEmail(record.student.email)));
}

function renderJournalManagement() {
  const ctx = state.teacher;
  const stats = analyzeClass(activePortfolioItems(), ctx.channels);
  const content = document.getElementById('teacher-content');
  const activeMembers = (ctx.record.members || []).filter((member) => member.status === 'active');
  const today = todayKey();
  const submittedEmails = new Set(stats.all.filter((journal) => todayKey(new Date(journal.createdAt)) === today).map((journal) => normalizeEmail(journal.student.email)));
  const submittedMembers = activeMembers.filter((member) => submittedEmails.has(normalizeEmail(member.email)));
  const missingMembers = activeMembers.filter((member) => !submittedEmails.has(normalizeEmail(member.email)));
  const rate = activeMembers.length ? Math.round(submittedMembers.length / activeMembers.length * 100) : 0;
  const unreturned = stats.all.length - stats.returned;
  const filtered = filterTeacherJournals(stats.all);
  content.innerHTML = `<section class="teacher-overview" aria-label="クラス概要">
    <div class="metric submission-donut-card"><div class="submission-donut" style="--rate:${rate * 3.6}deg"><span><strong>${rate}%</strong><small>${submittedMembers.length}/${activeMembers.length}人</small></span></div><span>今日の提出率</span></div>
    <div class="metric attention-metric"><strong>${missingMembers.length}</strong><span>今日の未提出</span></div>
    <div class="metric attention-metric"><strong>${unreturned}</strong><span>未返却の記録</span></div>
    <div class="metric"><strong>${stats.all.length}</strong><span>すべての記録</span></div>
  </section>
  <div class="teacher-dashboard-grid">
    <main class="teacher-submissions" aria-labelledby="submission-heading"><div class="teacher-list-heading"><div><span class="eyebrow">STUDENT JOURNALS</span><h2 id="submission-heading">提出一覧</h2></div><span id="journal-count" class="badge">${filtered.length}件</span></div>
      <div class="teacher-list-controls"><input id="journal-search" class="compact-input" value="${escapeHtml(state.teacherSearch)}" placeholder="氏名・本文・テーマを検索" aria-label="提出を検索"><div class="filter-row" role="group" aria-label="返却状況で絞り込む">${[['all',`すべて ${stats.all.length}`],['unreturned',`未返却 ${unreturned}`],['returned',`返却済み ${stats.returned}`]].map(([value, label]) => `<button class="filter-chip ${state.teacherFilter === value ? 'active' : ''}" data-filter="${value}" type="button">${label}</button>`).join('')}</div></div>
      <details class="batch-bar"><summary>まとめて支援・返却</summary><p class="muted small">未返却の記録が対象です。AIの内容は返却前に確認してください。</p><div class="button-row compact"><button id="batch-ai-simple" class="secondary" type="button">AI下書き</button><button id="batch-ai-detail" class="secondary" type="button">AIで詳しく</button><button id="batch-return" class="primary" type="button">一括返却</button></div></details>
      <div id="journal-list" class="journal-list teacher-journal-list">${teacherJournalCards(filtered)}</div>
    </main>
    <aside class="teacher-sidebar" aria-label="授業設定と今日の状況">
      <section class="panel teacher-theme-panel"><div class="compact-section-heading"><div><span class="eyebrow">TODAY'S THEME</span><h2>テーマ設定</h2></div></div><form id="theme-form">
        <label><span>適用日</span><input id="theme-date" type="date" value="${escapeHtml(ctx.record.settings.todayTheme?.date || today)}"></label><label><span>今日のテーマ</span><input id="today-theme" maxlength="200" value="${escapeHtml(ctx.record.settings.todayTheme?.text || '')}" placeholder="今日の学びをふり返ろう"></label>
        <details><summary>曜日ごとのテーマ</summary><div class="form-grid">${['月','火','水','木','金'].map((day, index) => `<label><span>${day}曜日</span><input class="weekly-theme" data-day="${index + 1}" maxlength="200" value="${escapeHtml(ctx.record.settings.weeklyThemes?.[index + 1] || '')}"></label>`).join('')}</div></details>
        <button class="primary wide" type="submit">保存して児童へ配信</button>
      </form></section>
      <section class="panel today-status-panel"><div class="compact-section-heading"><div><span class="eyebrow">TODAY'S STATUS</span><h2>今日の状況</h2></div><span class="muted small">${activeMembers.length}人</span></div>
        <h3>未提出 ${missingMembers.length}人</h3><div class="student-status-list missing">${missingMembers.length ? missingMembers.map((member) => `<span>${escapeHtml(member.name)}</span>`).join('') : '<p class="muted small">全員提出しています。</p>'}</div>
        <details><summary>提出済み ${submittedMembers.length}人</summary><div class="student-status-list submitted">${submittedMembers.map((member) => `<span>${escapeHtml(member.name)}</span>`).join('')}</div></details>
      </section>
    </aside>
  </div>`;
  document.getElementById('theme-form').addEventListener('submit', saveThemes);
  document.getElementById('journal-search').addEventListener('input', (event) => {
    state.teacherSearch = event.target.value;
    const next = filterTeacherJournals(stats.all);
    document.getElementById('journal-list').innerHTML = teacherJournalCards(next);
    document.getElementById('journal-count').textContent = `${next.length}件`;
    bindJournalCards(stats.all);
  });
  document.querySelectorAll('[data-filter]').forEach((button) => button.addEventListener('click', () => {
    state.teacherFilter = button.dataset.filter;
    renderJournalManagement();
  }));
  document.getElementById('batch-return').addEventListener('click', () => batchReturn(stats.all));
  document.getElementById('batch-ai-simple').addEventListener('click', () => batchAiDrafts(stats.all, false));
  document.getElementById('batch-ai-detail').addEventListener('click', () => batchAiDrafts(stats.all, true));
  bindJournalCards(stats.all);
}

function filterTeacherJournals(journals) {
  const query = state.teacherSearch.trim().toLowerCase();
  return [...journals].filter((journal) => {
    const returned = Boolean(state.teacher.channels.get(normalizeEmail(journal.student.email))?.feedback?.[journal.id]?.returned);
    const matchesStatus = state.teacherFilter === 'all' || (state.teacherFilter === 'returned' ? returned : !returned);
    const matchesQuery = !query || `${journal.student.name} ${journal.theme} ${journal.content}`.toLowerCase().includes(query);
    return matchesStatus && matchesQuery;
  }).sort((a, b) => {
    const aReturned = Boolean(state.teacher.channels.get(normalizeEmail(a.student.email))?.feedback?.[a.id]?.returned);
    const bReturned = Boolean(state.teacher.channels.get(normalizeEmail(b.student.email))?.feedback?.[b.id]?.returned);
    return Number(aReturned) - Number(bReturned) || new Date(b.createdAt) - new Date(a.createdAt);
  });
}

function teacherJournalCards(journals) {
  return journals.length ? journals.map((journal) => {
    const feedback = state.teacher.channels.get(normalizeEmail(journal.student.email))?.feedback?.[journal.id];
    return `<article class="journal-card teacher-journal-card ${feedback?.returned ? 'is-returned' : 'needs-reply'}"><button class="journal-main journal-open" data-email="${escapeHtml(journal.student.email)}" data-journal="${escapeHtml(journal.id)}" type="button">
      <div class="journal-meta"><span><strong>${escapeHtml(journal.student.name)}</strong> · ${escapeHtml(formatDate(journal.createdAt))}</span><span>${escapeHtml(journal.emotion || '')}</span></div>
      <h3>${escapeHtml(journal.theme || 'テーマなし')}</h3><div class="journal-body clamp">${escapeHtml(journal.content)}</div>
      <span class="badge ${feedback?.returned ? 'success-badge' : ''}">${feedback?.returned ? '返却済み' : '未返却'}</span></button>
      <div class="quick-feedback" aria-label="クイック返却">${feedback?.returned ? `<button class="quiet" data-quick-undo data-email="${escapeHtml(journal.student.email)}" data-journal="${escapeHtml(journal.id)}" type="button">↩ 返却を取り消す</button>` : `<span class="muted small">すぐ返す:</span>${['😊','👏','💪','⭐'].map((stamp) => `<button class="quick-stamp" data-quick-stamp="${stamp}" data-email="${escapeHtml(journal.student.email)}" data-journal="${escapeHtml(journal.id)}" type="button" aria-label="${stamp}で返却">${stamp}</button>`).join('')}</div>`}
    </article>`;
  }).join('') : '<div class="empty">提出はまだありません。</div>';
}

function bindJournalCards(all) {
  document.querySelectorAll('.journal-open').forEach((button) => button.addEventListener('click', () => {
    const journal = all.find((item) => item.id === button.dataset.journal && normalizeEmail(item.student.email) === normalizeEmail(button.dataset.email));
    const portfolio = state.teacher.portfolios.find(({ record }) => normalizeEmail(record.student.email) === normalizeEmail(button.dataset.email));
    pushAppRoute('teacher-feedback', { classId: state.teacher.record.classId, email: button.dataset.email, journalId: button.dataset.journal });
    renderFeedbackEditor(portfolio, journal);
  }));
  document.querySelectorAll('[data-quick-stamp]').forEach((button) => button.addEventListener('click', () => quickFeedback(button.dataset.email, button.dataset.journal, button.dataset.quickStamp, true)));
  document.querySelectorAll('[data-quick-undo]').forEach((button) => button.addEventListener('click', () => quickFeedback(button.dataset.email, button.dataset.journal, '', false)));
}

async function quickFeedback(email, journalId, stamp, returned) {
  const normalized = normalizeEmail(email);
  const member = state.teacher.record.members.find((item) => normalizeEmail(item.email) === normalized);
  let channel = state.teacher.channels.get(normalized);
  if (!channel || !member?.channelFileId) return toast('先に児童の参加を承認してください。');
  const existing = channel.feedback?.[journalId] || {};
  channel = setFeedback(channel, journalId, { ...existing, stamp: stamp || existing.stamp, returned });
  await withError(async () => {
    await persistTeacherChannel(normalized, channel);
    toast(returned ? `${stamp} を返しました。` : '返却を取り消しました。');
    renderJournalManagement();
  }, (message) => renderTeacherClass('journals', message));
}

function unreturnedJournals(journals) {
  return journals.filter((journal) => !state.teacher.channels.get(normalizeEmail(journal.student.email))?.feedback?.[journal.id]?.returned);
}

async function batchReturn(journals) {
  const targets = unreturnedJournals(journals);
  if (!targets.length) return toast('未返却の記録はありません。');
  if (!window.confirm(`${targets.length}件をまとめて返却します。よろしいですか？`)) return;
  setBusy(`${targets.length}件のおへんじを返しています…`);
  await withError(async () => {
    const changed = new Map();
    for (const journal of targets) {
      const email = normalizeEmail(journal.student.email);
      let channel = changed.get(email) || state.teacher.channels.get(email);
      if (!channel) continue;
      const existing = channel.feedback?.[journal.id] || {};
      channel = setFeedback(channel, journal.id, { ...existing, stamp: existing.stamp || '👏', returned: true });
      changed.set(email, channel);
    }
    for (const [email, channel] of changed) {
      await persistTeacherChannel(email, channel);
    }
    toast(`${targets.length}件を返却しました。`);
    renderTeacherClass('journals');
  }, (message) => renderTeacherClass('journals', message));
}

async function geminiText(prompt) {
  const apiKey = state.teacher.record.settings.geminiApiKey;
  if (!apiKey) throw new Error('クラス設定でGemini APIキーを保存してください。');
  if (!state.teacher.record.settings.geminiDataConsent) throw new Error('児童の文章をGemini APIへ送信することについて、学校の承認と同意をクラス設定で確認してください。');
  const model = state.teacher.record.settings.geminiModel || 'gemini-3.1-flash-lite';
  const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'x-goog-api-key': apiKey },
    body: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] })
  });
  const result = await response.json();
  if (!response.ok) throw new Error(result.error?.message || 'AI支援を利用できませんでした。');
  return result.candidates?.[0]?.content?.parts?.map((part) => part.text || '').join('')?.trim() || '';
}

function detailedAiResult(text, journal) {
  const jsonText = text.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/, '');
  const parsed = JSON.parse(jsonText);
  const highlights = (Array.isArray(parsed.highlights) ? parsed.highlights : []).map((item, index) => {
    const quote = String(item.quote || '').trim();
    const start = journal.content.indexOf(quote);
    return start < 0 || !quote ? null : { id: crypto.randomUUID?.() || `ai-${Date.now()}-${index}`, start, end: start + quote.length, text: quote, comment: String(item.comment || '').trim(), stamp: String(item.stamp || '').slice(0, 8) };
  }).filter(Boolean);
  return { comment: String(parsed.comment || '').trim(), stamp: String(parsed.stamp || '🌱').slice(0, 8), highlights };
}

async function batchAiDrafts(journals, detailed) {
  const targets = unreturnedJournals(journals);
  if (!targets.length) return toast('未返却の記録はありません。');
  if (!state.teacher.record.settings.geminiApiKey) return renderTeacherClass('settings', 'AI支援を使うにはGemini APIキーを設定してください。');
  if (!state.teacher.record.settings.geminiDataConsent) return renderTeacherClass('settings', 'AI送信に関する学校の承認と同意を確認してください。');
  if (!window.confirm(`${targets.length}件のAI下書きを作ります。API利用料が発生する場合があります。続けますか？`)) return;
  setBusy(`AIが${targets.length}件の下書きを作っています…`);
  await withError(async () => {
    const changed = new Map();
    let completed = 0;
    for (const journal of targets) {
      const prompt = detailed
        ? `あなたは小学校の担任です。児童のふりかえりを読み、温かい総合コメントと、本文中の具体的なよさ2箇所以内を選んでください。必ず次のJSONだけを返してください。{"comment":"100字以内","stamp":"絵文字1個","highlights":[{"quote":"本文から完全一致する短い引用","comment":"具体的な称賛","stamp":"絵文字1個"}]}\nテーマ: ${journal.theme}\n本文: ${journal.content}`
        : `あなたは小学校の担任です。次の児童のふりかえりへ、具体的なよさを認め、次の学びにつながる温かいコメントを100字以内で日本語で書いてください。コメント本文だけを返してください。\nテーマ: ${journal.theme}\n本文: ${journal.content}`;
      const generated = await geminiText(prompt);
      const draft = detailed ? detailedAiResult(generated, journal) : { comment: generated, stamp: '🌱', highlights: [] };
      const email = normalizeEmail(journal.student.email);
      let channel = changed.get(email) || state.teacher.channels.get(email);
      if (!channel) continue;
      channel = setFeedback(channel, journal.id, { ...draft, returned: false });
      changed.set(email, channel);
      completed += 1;
    }
    for (const [email, channel] of changed) {
      await persistTeacherChannel(email, channel);
    }
    toast(`${completed}件の下書きを保存しました。確認後に返却してください。`);
    renderTeacherClass('journals');
  }, (message) => renderTeacherClass('journals', message));
}

async function saveThemes(event) {
  event.preventDefault();
  const ctx = state.teacher;
  ctx.record.settings.todayTheme = { date: document.getElementById('theme-date').value, text: document.getElementById('today-theme').value.trim() };
  document.querySelectorAll('.weekly-theme').forEach((input) => { ctx.record.settings.weeklyThemes[input.dataset.day] = input.value.trim(); });
  ctx.record.updatedAt = new Date().toISOString();
  setBusy('テーマを児童へ配信しています…');
  await withError(async () => {
    await updateDocument(ctx.file, ctx.record);
    for (const member of ctx.record.members.filter((item) => item.status === 'active' && item.channelFileId)) {
      const email = normalizeEmail(member.email);
      const channel = syncChannel(ctx.channels.get(email), ctx.record, member);
      await persistTeacherChannel(email, channel);
    }
    toast('テーマを保存しました。');
    renderTeacherClass('journals');
  }, (message) => renderTeacherClass('journals', message));
}

function renderFeedbackEditor(portfolioItem, journal, error = '') {
  const ctx = state.teacher;
  const email = normalizeEmail(portfolioItem.record.student.email);
  const channel = ctx.channels.get(email);
  const existing = channel?.feedback?.[journal.id] || {};
  const draftKey = `${email}:${journal.id}`;
  if (state.feedbackDraftKey !== draftKey) {
    state.feedbackDraftKey = draftKey;
    state.feedbackDraftHighlights = (existing.highlights || []).map((item) => ({ ...item }));
  }
  app.innerHTML = shell(`<div class="page-heading"><div><span class="badge">おへんじ</span><h1>${escapeHtml(portfolioItem.record.student.name)}</h1></div><button id="journal-back" class="quiet" type="button">提出一覧へ</button></div>
    ${errorNotice(error)}
    <div class="detail-grid"><article class="panel"><div class="journal-meta"><span>${escapeHtml(formatDate(journal.createdAt))}</span><span>${escapeHtml(journal.emotion || '')}</span></div><h2>${escapeHtml(journal.theme || 'テーマなし')}</h2><p class="field-note">ほめたい文章を選び、「選んだ部分にコメント」を押すと、範囲ごとのおへんじを付けられます。</p><div id="journal-source" class="journal-body selectable-journal">${escapeHtml(journal.content)}</div><button id="add-highlight" class="secondary wide highlight-add" type="button">🖍️ 選んだ部分にコメント</button>${journal.imageFileId ? `<div class="image-slot" data-file="${escapeHtml(journal.imageFileId)}"><p>画像を読み込んでいます…</p></div>` : ''}</article>
    <section class="panel"><h2>先生のおへんじ</h2><form id="feedback-form">
      <label><span>コメント</span><textarea id="feedback-comment" maxlength="4000" placeholder="よかったところや、次につながるひとこと">${escapeHtml(existing.comment || '')}</textarea></label>
      <label><span>スタンプ</span><select id="feedback-stamp"><option value="">なし</option>${['😊','👏','💪','⭐','🌱','✨'].map((stamp) => `<option ${existing.stamp === stamp ? 'selected' : ''}>${stamp}</option>`).join('')}</select></label>
      <div class="highlight-editor"><h3>文章ごとのおへんじ</h3><div id="highlight-list">${feedbackHighlightRows()}</div></div>
      <div class="button-row"><button class="primary" type="submit">保存して返却</button><button id="ai-draft" class="secondary" type="button">AIで下書き</button></div>
    </form></section></div>`);
  document.getElementById('journal-back').addEventListener('click', () => appBack(() => {
    replaceAppRoute('teacher-class', { classId: state.teacher.record.classId, tab: 'journals' });
    renderTeacherClass('journals');
  }));
  document.getElementById('feedback-form').addEventListener('submit', (event) => saveFeedback(event, portfolioItem, journal));
  document.getElementById('ai-draft').addEventListener('click', () => generateAiDraft(journal, portfolioItem));
  document.getElementById('add-highlight').addEventListener('click', () => addSelectionHighlight(portfolioItem, journal));
  bindHighlightRows(portfolioItem, journal);
  loadImages();
}

function feedbackHighlightRows() {
  return state.feedbackDraftHighlights.length ? state.feedbackDraftHighlights.map((item, index) => `<div class="highlight-row"><div class="highlight-quote">「${escapeHtml(item.text)}」</div><label><span>この部分へのひとこと</span><textarea data-highlight-comment="${index}" maxlength="1000">${escapeHtml(item.comment || '')}</textarea></label><div class="highlight-actions"><select data-highlight-stamp="${index}" aria-label="範囲コメントのスタンプ"><option value="">なし</option>${['😊','👏','💪','⭐','🌱','✨'].map((stamp) => `<option ${item.stamp === stamp ? 'selected' : ''}>${stamp}</option>`).join('')}</select><button class="danger" data-remove-highlight="${index}" type="button">削除</button></div></div>`).join('') : '<p class="muted small">まだ文章ごとのおへんじはありません。</p>';
}

function bindHighlightRows(portfolioItem, journal) {
  document.querySelectorAll('[data-highlight-comment]').forEach((input) => input.addEventListener('input', () => { state.feedbackDraftHighlights[Number(input.dataset.highlightComment)].comment = input.value; }));
  document.querySelectorAll('[data-highlight-stamp]').forEach((input) => input.addEventListener('change', () => { state.feedbackDraftHighlights[Number(input.dataset.highlightStamp)].stamp = input.value; }));
  document.querySelectorAll('[data-remove-highlight]').forEach((button) => button.addEventListener('click', () => {
    state.feedbackDraftHighlights.splice(Number(button.dataset.removeHighlight), 1);
    renderFeedbackEditor(portfolioItem, journal);
  }));
}

function addSelectionHighlight(portfolioItem, journal) {
  const source = document.getElementById('journal-source');
  const selection = window.getSelection();
  if (!selection?.rangeCount || selection.isCollapsed) return toast('本文から、ほめたい文章を選んでください。');
  const range = selection.getRangeAt(0);
  if (!source.contains(range.commonAncestorContainer)) return toast('左側の本文から文章を選んでください。');
  const before = range.cloneRange();
  before.selectNodeContents(source);
  before.setEnd(range.startContainer, range.startOffset);
  const start = before.toString().length;
  const text = range.toString().trim();
  if (!text) return toast('本文から文章を選んでください。');
  const adjustedStart = journal.content.indexOf(text, start);
  const safeStart = adjustedStart >= 0 ? adjustedStart : start;
  state.feedbackDraftHighlights.push({ id: crypto.randomUUID(), start: safeStart, end: safeStart + text.length, text, comment: '', stamp: '⭐' });
  selection.removeAllRanges();
  renderFeedbackEditor(portfolioItem, journal);
}

async function saveFeedback(event, portfolioItem, journal) {
  event.preventDefault();
  const email = normalizeEmail(portfolioItem.record.student.email);
  const member = state.teacher.record.members.find((item) => normalizeEmail(item.email) === email);
  let channel = state.teacher.channels.get(email);
  if (!channel || !member?.channelFileId) return renderFeedbackEditor(portfolioItem, journal, 'おへんじを届ける準備が完了していません。クラス設定で児童を承認してください。');
  channel = setFeedback(channel, journal.id, { comment: document.getElementById('feedback-comment').value, stamp: document.getElementById('feedback-stamp').value, highlights: state.feedbackDraftHighlights, returned: true });
  setBusy('おへんじを保存しています…');
  await withError(async () => {
    await persistTeacherChannel(email, channel);
    state.feedbackDraftKey = '';
    state.feedbackDraftHighlights = [];
    toast('児童へ返却しました。');
    appBack(() => {
      replaceAppRoute('teacher-class', { classId: state.teacher.record.classId, tab: 'journals' });
      renderTeacherClass('journals');
    });
  }, (message) => renderFeedbackEditor(portfolioItem, journal, message));
}

async function generateAiDraft(journal, portfolioItem) {
  const apiKey = state.teacher.record.settings.geminiApiKey;
  if (!apiKey) return renderFeedbackEditor(portfolioItem, journal, 'クラス設定でGemini APIキーを保存してください。');
  if (!state.teacher.record.settings.geminiDataConsent) return renderFeedbackEditor(portfolioItem, journal, '先にクラス設定で、AI送信に関する学校の承認と同意を確認してください。');
  const button = document.getElementById('ai-draft');
  button.disabled = true;
  button.textContent = '下書きを作成中…';
  try {
    document.getElementById('feedback-comment').value = await geminiText(`あなたは小学校の担任です。次の児童のふりかえりへ、具体的なよさを認め、次の学びにつながる温かいコメントを100字以内で日本語で書いてください。コメント本文だけを返してください。\n\nテーマ: ${journal.theme}\n本文: ${journal.content}`);
    toast('AIの下書きを作成しました。内容を確認してください。');
  } catch (error) { renderFeedbackEditor(portfolioItem, journal, friendlyError(error)); }
  finally { if (button.isConnected) { button.disabled = false; button.textContent = 'AIで下書き'; } }
}

function renderVitals() {
  const stats = analyzeClass(activePortfolioItems(), state.teacher.channels);
  const days = Array.from({ length: 14 }, (_, index) => { const date = new Date(); date.setDate(date.getDate() - (13 - index)); return date; });
  const content = document.getElementById('teacher-content');
  content.innerHTML = `<div class="vitals-grid"><section class="panel vitals-alerts"><span class="eyebrow">CARE SIGNALS</span><h2>気にかけたい児童</h2>${stats.alerts.length ? stats.alerts.map((alert) => `<div class="notice"><strong>${escapeHtml(alert.name)}</strong> — ${escapeHtml(alert.reason)}</div>`).join('') : '<div class="empty">現在の記録から強い変化は検出されていません。対面での様子も必ず合わせて見てください。</div>'}</section>
    <section class="panel vitals-heatmap"><span class="eyebrow">14 DAYS</span><h2>クラスの心の波</h2><div class="heatmap"><div class="heatmap-row heatmap-head"><span>児童</span>${days.map((date) => `<span>${date.getMonth() + 1}/${date.getDate()}</span>`).join('')}</div>${activePortfolioItems().map(({ record }) => `<div class="heatmap-row"><strong>${escapeHtml(record.student.name)}</strong>${days.map((date) => { const journal = (record.journals || []).find((item) => todayKey(new Date(item.createdAt)) === todayKey(date)); return `<span title="${escapeHtml(journal?.content || '未提出')}">${escapeHtml(journal?.emotion || '·')}</span>`; }).join('')}</div>`).join('')}</div></section></div>`;
}

function publicEntryUrl() {
  return config.publicEntryUrl || new URL('./', location.href).href;
}

async function renderClassSettings() {
  const ctx = state.teacher;
  const settings = ctx.record.settings;
  const token = await encodeSignedInvite({
    classCode: ctx.record.classCode,
    className: ctx.record.className,
    teacherEmail: ctx.record.teacher.email,
    teacherName: ctx.record.teacher.name,
    approvalRequired: settings.approvalRequired,
    acceptingMembers: settings.acceptingMembers !== false
  }, settings.inviteSecurity);
  const link = signedInviteUrl(publicEntryUrl(), token);
  const content = document.getElementById('teacher-content');
  if (state.teacher !== ctx || !content) return;
  const rejectedNotice = ctx.rejectedPortfolios?.length
    ? `<div class="error"><strong>安全性チェックで${ctx.rejectedPortfolios.length}件を除外しました。</strong><p class="small">所有者と児童アカウントの不一致、または失効した招待が原因です。Drive上の原ファイルは変更していません。</p></div>` : '';
  content.innerHTML = `<div class="teacher-settings-grid"><section class="panel invite-layout wide-setting"><div><span class="eyebrow">INVITE STUDENTS</span><h2>児童を招待する</h2><p class="muted">教室ではQRコード、遠隔では専用URLを配ると簡単です。</p><div class="class-code">${escapeHtml(ctx.record.classCode)}</div><div class="button-row"><button id="copy-code" class="secondary" type="button">クラスコードをコピー</button><button id="student-preview" class="quiet" type="button">児童画面を確認</button></div><div class="url-box">${escapeHtml(link)}</div><div class="button-row" style="margin-top:12px"><button id="copy-url" class="primary" type="button">専用URLをコピー</button><button id="copy-text" class="secondary" type="button">招待文をコピー</button><button id="rotate-invite" class="danger" type="button">招待URLを失効・再発行</button></div><p class="muted small">有効期限: ${escapeHtml(formatDate(settings.inviteSecurity?.expiresAt))}。再発行すると旧URLでの新規参加は拒否されます。</p></div><div class="qr-actions"><div id="qr" class="qr-box" aria-label="児童招待用QRコード"></div><button id="download-qr" class="secondary wide" type="button">QRを画像で保存</button></div></section>
  <section class="panel"><h2>参加設定</h2><form id="admission-form"><label class="check-label"><input id="accepting-members" type="checkbox" ${settings.acceptingMembers !== false ? 'checked' : ''}><span>新しい児童の参加を受け付ける</span></label><label class="check-label"><input id="approval" type="checkbox" ${settings.approvalRequired ? 'checked' : ''}><span>参加には先生の承認が必要</span></label><p class="muted small">受付停止は新規参加だけを止め、参加済みの児童の記録は保持します。</p><button class="primary" type="submit">設定を保存</button></form></section>
  <section class="panel wide-setting"><div class="section-heading"><div><span class="eyebrow">ROSTER</span><h2>参加申請・名簿</h2></div><span class="badge">${(ctx.record.members || []).filter((member) => member.status === 'active').length}人 参加中</span></div>${rejectedNotice}<div id="member-list">${memberRows()}</div><form id="invite-member" class="inline-form"><input id="member-name" required placeholder="児童名"><input id="member-email" type="email" required placeholder="児童のGoogleメール"><button class="secondary" type="submit">1人追加</button></form><details><summary>名簿をまとめて追加</summary><form id="bulk-members"><label><span>1行に「名前, メールアドレス」</span><textarea id="bulk-member-text" placeholder="山田 花子, hanako@example.ed.jp\n鈴木 太郎, taro@example.ed.jp"></textarea></label><button class="secondary" type="submit">まとめて名簿へ追加</button></form></details></section>
  <section class="panel"><h2>Gemini AI支援（任意）</h2><form id="gemini-form"><label><span>APIキー</span><input id="gemini-key" type="password" value="${escapeHtml(settings.geminiApiKey || '')}" autocomplete="off"></label><label><span>モデル</span><input id="gemini-model" value="${escapeHtml(settings.geminiModel || 'gemini-3.1-flash-lite')}"></label><label class="check-label"><input id="gemini-consent" type="checkbox" ${settings.geminiDataConsent ? 'checked' : ''}><span>児童の本文がGoogle Generative Language APIへ送信されることを理解し、学校の承認を得ている</span></label><p class="muted small">APIキーは先生の非共有Driveファイルに平文保存されます。Google Cloud側でキーを学校用オリジンとAPIに制限し、AIコメントは必ず先生が確認してください。</p><button class="primary" type="submit">AI設定を保存</button></form></section>
  <section class="panel"><h2>データ出力・保全</h2><p class="muted small">完全バックアップは、クラス設定・児童の記録・おへんじ・共有画像の複製を先生のDriveに作成します。</p><div class="button-row"><button id="archive" class="primary" type="button">Driveへ完全バックアップ</button><button id="csv" class="secondary" type="button">CSVをダウンロード</button><button id="print" class="secondary" type="button">印刷・PDF</button></div></section></div>`;
  document.getElementById('copy-url').addEventListener('click', () => copy(link, '専用URLをコピーしました。'));
  document.getElementById('copy-code').addEventListener('click', () => copy(ctx.record.classCode, 'クラスコードをコピーしました。'));
  document.getElementById('copy-text').addEventListener('click', () => copy(`「${ctx.record.className}」のふりかえりジャーナルに参加してください。\nクラスコード: ${ctx.record.classCode}\n${link}`, '招待文をコピーしました。'));
  document.getElementById('student-preview').addEventListener('click', () => window.open(link, '_blank', 'noopener'));
  document.getElementById('download-qr').addEventListener('click', () => downloadQrCode(ctx.record.className));
  document.getElementById('rotate-invite').addEventListener('click', rotateInviteLink);
  document.getElementById('admission-form').addEventListener('submit', saveAdmissionSettings);
  document.getElementById('invite-member').addEventListener('submit', addInvitedMember);
  document.getElementById('bulk-members').addEventListener('submit', addBulkMembers);
  document.getElementById('gemini-form').addEventListener('submit', saveGeminiSettings);
  document.getElementById('csv').addEventListener('click', downloadCsv);
  document.getElementById('print').addEventListener('click', printClass);
  document.getElementById('archive').addEventListener('click', archiveClassToDrive);
  document.querySelectorAll('[data-member-action]').forEach((button) => button.addEventListener('click', () => updateMemberStatus(button.dataset.email, button.dataset.memberAction)));
  ensureQr().then(() => renderQr(link)).catch(() => {
    const target = document.getElementById('qr');
    if (target) target.textContent = 'QRコードを読み込めませんでした。専用URLをコピーしてください。';
  });
}

function memberRows() {
  const members = state.teacher.record.members || [];
  return members.length ? `<div class="member-table">${members.map((member) => `<div class="member-row"><div><strong>${escapeHtml(member.name)}</strong><span>${escapeHtml(member.email)}</span></div><span class="badge">${statusLabel(member.status)}</span><div class="button-row compact">${member.status !== 'active' ? `<button class="primary" data-member-action="active" data-email="${escapeHtml(member.email)}" type="button">承認</button>` : ''}${member.status !== 'rejected' ? `<button class="danger" data-member-action="rejected" data-email="${escapeHtml(member.email)}" type="button">除外</button>` : ''}</div></div>`).join('')}</div>` : '<div class="empty">まだ児童はいません。</div>';
}

function statusLabel(status) {
  return ({ active: '参加中', pending: '承認待ち', invited: '招待済み', rejected: '除外' })[status] || status;
}

async function saveAdmissionSettings(event) {
  event.preventDefault();
  state.teacher.record.settings.approvalRequired = document.getElementById('approval').checked;
  state.teacher.record.settings.acceptingMembers = document.getElementById('accepting-members').checked;
  await saveTeacherRecord('参加設定を保存しました。');
}

async function addInvitedMember(event) {
  event.preventDefault();
  const email = normalizeEmail(document.getElementById('member-email').value);
  const name = document.getElementById('member-name').value.trim();
  if (!isEmail(email)) return toast('メールアドレスを確認してください。');
  const existing = state.teacher.record.members.find((member) => normalizeEmail(member.email) === email);
  if (existing) { existing.name = name; if (existing.status === 'rejected') existing.status = 'invited'; }
  else state.teacher.record.members.push({ email, name, role: 'student', status: 'invited', portfolioFileId: '', channelFileId: '', joinedAt: '' });
  await saveTeacherRecord('名簿へ追加しました。');
}

async function addBulkMembers(event) {
  event.preventDefault();
  const rows = document.getElementById('bulk-member-text').value.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
  const invalid = [];
  let added = 0;
  for (const line of rows) {
    const separator = line.includes('\t') ? '\t' : ',';
    const parts = line.split(separator).map((part) => part.trim());
    const email = normalizeEmail(parts.pop());
    const name = parts.join(' ').trim();
    if (!name || !isEmail(email)) { invalid.push(line); continue; }
    const existing = state.teacher.record.members.find((member) => normalizeEmail(member.email) === email);
    if (existing) { existing.name = name; if (existing.status === 'rejected') existing.status = 'invited'; }
    else state.teacher.record.members.push({ email, name, role: 'student', status: 'invited', portfolioFileId: '', channelFileId: '', joinedAt: '' });
    added += 1;
  }
  if (!added) return toast('「名前, メールアドレス」の形で入力してください。');
  await saveTeacherRecord(invalid.length ? `${added}人を追加しました。${invalid.length}行は形式を確認してください。` : `${added}人を名簿へ追加しました。`);
}

async function updateMemberStatus(email, status) {
  const ctx = state.teacher;
  const member = ctx.record.members.find((item) => normalizeEmail(item.email) === normalizeEmail(email));
  if (!member) return;
  member.status = status;
  setBusy(status === 'active' ? '児童を承認しています…' : '参加状態を更新しています…');
  await withError(async () => {
    let channel = ctx.channels.get(normalizeEmail(email));
    if (!channel && member.portfolioFileId) {
      channel = createChannel({ classRecord: ctx.record, member, status });
      const created = await state.drive.createChannel(channel, await studentKey(member.email));
      await state.drive.shareWithUser(created.id, member.email, 'reader');
      member.channelFileId = created.id;
      ctx.channelFiles.set(normalizeEmail(email), created);
      rememberJson(created, channel);
    } else if (channel && member.channelFileId) {
      channel = syncChannel(channel, ctx.record, member);
      await persistTeacherChannel(email, channel);
    }
    if (channel) ctx.channels.set(normalizeEmail(email), channel);
    await updateDocument(ctx.file, ctx.record);
    toast(status === 'active' ? '児童を承認しました。' : '参加状態を更新しました。');
    renderTeacherClass('settings');
  }, (message) => renderTeacherClass('settings', message));
}

async function saveGeminiSettings(event) {
  event.preventDefault();
  const apiKey = document.getElementById('gemini-key').value.trim();
  const consent = document.getElementById('gemini-consent').checked;
  if (apiKey && !consent) return renderTeacherClass('settings', '児童の文章の外部送信について、学校の承認と同意を確認してください。');
  state.teacher.record.settings.geminiApiKey = apiKey;
  state.teacher.record.settings.geminiModel = document.getElementById('gemini-model').value.trim() || 'gemini-3.1-flash-lite';
  state.teacher.record.settings.geminiDataConsent = Boolean(apiKey && consent);
  await saveTeacherRecord('AI設定を保存しました。');
}

async function saveTeacherRecord(message) {
  const ctx = state.teacher;
  ctx.record.updatedAt = new Date().toISOString();
  setBusy('設定を保存しています…');
  await withError(async () => { await updateDocument(ctx.file, ctx.record); toast(message); renderTeacherClass('settings'); }, (error) => renderTeacherClass('settings', error));
}

async function rotateInviteLink() {
  if (!window.confirm('現在の招待URLを失効させ、新しいURLを発行しますか？参加済みの児童には影響しません。')) return;
  const ctx = state.teacher;
  setBusy('招待URLを再発行しています…');
  await withError(async () => {
    ctx.record = await rotateInviteSecurity(ctx.record);
    await updateDocument(ctx.file, ctx.record);
    toast('旧しい招待URLを失効させ、新しいURLを発行しました。');
    renderTeacherClass('settings');
  }, (message) => renderTeacherClass('settings', message));
}

async function archiveClassToDrive() {
  if (!window.confirm('現在表示しているクラスの完全バックアップを、先生のGoogle Driveに作成しますか？')) return;
  const ctx = state.teacher;
  const archivedAt = new Date().toISOString();
  setBusy('クラスの完全バックアップを作成しています…');
  await withError(async () => {
    const folder = await state.drive.createFolder(
      `ふりかえりジャーナル_バックアップ_${ctx.record.className}_${todayKey()}`,
      { rjType: 'archive-folder', rjClassId: ctx.record.classId, rjArchivedAt: archivedAt }
    );
    const properties = { rjType: 'archive-item', rjClassId: ctx.record.classId, rjArchivedAt: archivedAt };
    await state.drive.createJson('class.json', ctx.record, properties, [folder.id]);
    for (const { record } of ctx.portfolios) {
      const key = await studentKey(record.student.email);
      await state.drive.createJson(`portfolio_${key.slice(0, 12)}.json`, record, properties, [folder.id]);
      for (const journal of record.journals || []) {
        if (!journal.imageFileId) continue;
        try {
          const blob = await state.drive.getBlob(journal.imageFileId);
          await state.drive.createFile({
            name: `image_${journal.id}${blob.type === 'image/png' ? '.png' : '.jpg'}`,
            mimeType: blob.type || 'application/octet-stream',
            parents: [folder.id],
            appProperties: { ...properties, rjJournalId: journal.id }
          }, blob);
        } catch (imageError) {
          // The manifest keeps the source ID so an administrator can inspect a policy-blocked image later.
        }
      }
    }
    for (const [email, channel] of ctx.channels) {
      const key = await studentKey(email);
      await state.drive.createJson(`channel_${key.slice(0, 12)}.json`, channel, properties, [folder.id]);
    }
    await state.drive.createJson('manifest.json', {
      schemaVersion: 1,
      kind: 'reflection-journal-archive',
      classId: ctx.record.classId,
      className: ctx.record.className,
      archivedAt,
      archivedBy: state.user.email,
      portfolioCount: ctx.portfolios.length,
      channelCount: ctx.channels.size,
      sourceFileIds: {
        class: ctx.file.id,
        portfolios: ctx.portfolios.map(({ file }) => file.id),
        channels: [...ctx.channelFiles.values()].map((file) => file.id)
      }
    }, properties, [folder.id]);
    toast('先生のDriveに完全バックアップを作成しました。');
    renderTeacherClass('settings');
  }, (message) => renderTeacherClass('settings', message));
}

function downloadCsv() {
  const blob = new Blob([exportCsv(activePortfolioItems(), state.teacher.channels)], { type: 'text/csv;charset=utf-8' });
  downloadBlob(blob, `ふりかえりジャーナル_${state.teacher.record.className}_${todayKey()}.csv`);
}

function printClass() {
  const win = window.open('', '_blank', 'noopener');
  if (!win) return toast('印刷画面を開けませんでした。');
  const body = activePortfolioItems().map(({ record }) => `<section><h2>${escapeHtml(record.student.name)}</h2>${(record.journals || []).map((journal) => { const feedback = state.teacher.channels.get(normalizeEmail(record.student.email))?.feedback?.[journal.id] || {}; return `<article><h3>${escapeHtml(formatDate(journal.createdAt))} ${escapeHtml(journal.theme)}</h3><p>${escapeHtml(journal.content)}</p>${feedback.returned ? `<p><strong>先生:</strong> ${escapeHtml(feedback.stamp)} ${escapeHtml(feedback.comment)}</p>` : ''}</article>`; }).join('')}</section>`).join('');
  win.document.write(`<!doctype html><html lang="ja"><meta charset="utf-8"><title>${escapeHtml(state.teacher.record.className)}</title><style>body{font-family:sans-serif;line-height:1.8}section{break-after:page}article{border-bottom:1px solid #ccc;padding:12px}</style><body><h1>${escapeHtml(state.teacher.record.className)}</h1>${body}</body></html>`);
  win.document.close();
  win.print();
}

async function renderStudentHome(error = '') {
  setStudentBusy('参加済みのクラスを探しています…');
  const files = await withError(() => state.drive.listOwnPortfolios(), (message) => renderHome(message));
  if (!files) return;
  const portfolios = (await loadJsonItems(files)).filter(({ file, record }) =>
    record?.kind === 'reflection-journal-portfolio'
    && normalizeEmail(record.student?.email) === state.user.email
    && normalizeEmail(file.owners?.[0]?.emailAddress) === state.user.email);
  state.studentPortfolios = portfolios;
  // 参加クラスが1つだけなら、一覧を出さずにそのクラスを開く。低学年ほど「押す所を探す」ことが負担になる。
  // 一覧へは［クラス一覧へ］と端末の戻る操作で戻れる（renderHistoryRoute がこの旗を下ろす）。
  if (state.restoreSingleStudentClass) {
    state.restoreSingleStudentClass = false;
    if (portfolios.length === 1) {
      const only = portfolios[0];
      pushAppRoute('student-portfolio', { fileId: only.file.id, classId: only.record.class.id });
      return openPortfolio(only);
    }
  }
  const today = todayKey();
  app.innerHTML = shell(`<div class="student-ui"><div class="page-heading"><div><span class="badge">${studentText('児童')}</span><h1>${studentText('参加しているクラス')}</h1></div><button id="back" class="quiet" type="button">${studentText('使い方を変える')}</button></div>
    ${studentErrorNotice(error)}${portfolios.length ? `<div class="grid">${portfolios.map(({ file, record }) => {
      const journals = record.journals || [];
      const wroteToday = journals.some((journal) => todayKey(new Date(journal.createdAt)) === today);
      return `<button class="item-card student-portfolio" data-file="${escapeHtml(file.id)}" type="button" aria-label="${escapeHtml(record.class?.name || 'クラス')}をひらく。ふりかえり${journals.length}件。${wroteToday ? '今日のぶんは提出ずみ' : '今日のぶんはまだ'}"><span class="badge">${escapeHtml(record.class?.code || '')}</span><h3 class="user-content">${studentText(record.class?.name || 'クラス')}</h3><p class="class-today ${wroteToday ? 'done' : 'todo'}">${wroteToday ? `✅ ${studentText('今日のふりかえりを提出しました')}` : `✏️ ${studentText('今日のふりかえりを書こう')}`}</p><p class="muted small">${journals.length}${studentText('件')}のふりかえり</p></button>`;
    }).join('')}</div>` : `<div class="empty">${studentText('参加済みのクラスはありません。先生のQRコードまたは専用URLから参加してください。')}</div>`}</div>`);
  document.getElementById('back').addEventListener('click', () => appBack(() => { forgetRole(); replaceAppRoute('home'); renderHome(); }));
  document.querySelectorAll('.student-portfolio').forEach((button) => button.addEventListener('click', () => {
    const item = portfolios.find(({ file }) => file.id === button.dataset.file);
    pushAppRoute('student-portfolio', { fileId: item.file.id, classId: item.record.class.id });
    openPortfolio(item);
  }));
}

async function renderJoin(error = '') {
  const invite = state.invite;
  if (!allowedWorkspaceDomain(invite.teacherEmail)) return renderHome('招待元のGoogle Workspaceドメインは、このアプリで許可されていません。');
  if (!invite.acceptingMembers) return renderHome('この招待では新しい参加を受け付けていません。先生から新しいURLを受け取ってください。');
  setStudentBusy('参加状況を確認しています…');
  const files = await withError(() => state.drive.listOwnPortfolios(invite.classId), (message) => renderHome(message));
  if (!files) return;
  if (files.length) {
    const record = await withError(() => state.drive.getJson(files[0].id), (message) => renderHome(message));
    if (record) {
      replaceAppRoute('student-portfolio', { fileId: files[0].id, classId: record.class.id }, location.pathname + location.search);
      return openPortfolio({ file: files[0], record });
    }
  }
  if (!invite.signed) return renderHome('この招待URLは旧形式です。安全に参加するため、先生から新しいURLまたはQRコードを受け取ってください。');
  app.innerHTML = shell(`<div class="student-ui"><div class="page-heading"><div><span class="badge">${studentText('招待')}</span><h1 class="user-content">${studentText(invite.className)}</h1></div><button id="cancel" class="quiet" type="button">${studentText('戻る')}</button></div>
    ${studentErrorNotice(error)}<section class="panel"><p>${studentText('先生')}: <strong class="user-content">${escapeHtml(invite.teacherName || invite.teacherEmail)}</strong></p><p>クラスコード: <strong>${escapeHtml(invite.classCode)}</strong></p><form id="join-class"><label><span>${studentText('先生に表示する名前')}</span><input id="student-name" maxlength="80" required value="${escapeHtml(state.user.name || '')}"></label><button class="primary wide" type="submit">${studentText('このクラスに参加する')}</button></form><p class="muted small">${studentText('参加すると、あなたのふりかえりを先生が確認できるようになります。')}</p></section></div>`);
  document.getElementById('cancel').addEventListener('click', () => appBack(() => {
    state.invite = null;
    replaceAppRoute('student-home', {}, location.pathname + location.search);
    renderStudentHome();
  }));
  document.getElementById('join-class').addEventListener('submit', joinClass);
}

async function joinClass(event) {
  event.preventDefault();
  const name = document.getElementById('student-name').value.trim();
  if (!name) return;
  setStudentBusy('クラスに参加しています…');
  await withError(async () => {
    if (!state.invite?.signed || new Date(state.invite.expiresAt).getTime() <= Date.now()) throw new Error('招待URLが失効しています。先生から新しいURLを受け取ってください。');
    const record = createPortfolio({ invite: state.invite, student: { ...state.user, name } });
    const file = await state.drive.createPortfolio(record);
    rememberJson(file, record);
    try { await state.drive.shareWithUser(file.id, state.invite.teacherEmail, 'reader'); }
    catch (error) {
      replaceAppRoute('student-portfolio', { fileId: file.id, classId: record.class.id }, location.pathname + location.search);
      return openPortfolio({ file, record }, 'クラスへの参加は完了しましたが、先生へ記録を届ける設定が学校側で拒否されました。「先生へもう一度届ける」を押してください。');
    }
    replaceAppRoute('student-portfolio', { fileId: file.id, classId: record.class.id }, location.pathname + location.search);
    await openPortfolio({ file, record });
  }, (message) => renderJoin(message));
}

async function findStudentChannel(classId) {
  const files = await state.drive.listSharedChannels(classId);
  for (const file of files) {
    try {
      const { record } = await loadJsonItem(file);
      if (validateChannelForStudent(file, record, state.user.email, classId)) return { file, record };
    } catch (error) {}
  }
  return null;
}

async function openPortfolio(item, error = '') {
  setStudentBusy('ふりかえりを開いています…');
  await withError(async () => {
    if (item.record?.kind !== 'reflection-journal-portfolio'
      || normalizeEmail(item.record.student?.email) !== state.user.email
      || normalizeEmail(item.file.owners?.[0]?.emailAddress) !== state.user.email) throw new Error('このポートフォリオは、ログイン中の児童のDrive記録として確認できません。');
    const channelItem = await findStudentChannel(item.record.class.id);
    state.portfolioFile = item.file;
    state.portfolio = item.record;
    state.channelFile = channelItem?.file || null;
    state.channel = channelItem?.record || null;
    renderPortfolio(error);
  }, (message) => renderStudentHome(message));
}

function draftKey() { return `rj_draft_v2_${state.portfolio.class.id}_${state.user.email}`; }
function readKey() { return `rj_read_v2_${state.portfolio.class.id}_${state.user.email}`; }

function readJournalIds() {
  try { return new Set(JSON.parse(localStorage.getItem(readKey()) || '[]')); }
  catch (error) { return new Set(); }
}

function unreadFeedbackJournals() {
  const read = readJournalIds();
  return (state.portfolio.journals || []).filter((journal) => state.channel?.feedback?.[journal.id]?.returned && !read.has(journal.id));
}

function markJournalRead(journalId) {
  const read = readJournalIds();
  read.add(journalId);
  try { localStorage.setItem(readKey(), JSON.stringify([...read].slice(-500))); } catch (error) {}
}

function renderPortfolio(error = '') {
  const portfolio = state.portfolio;
  const channel = state.channel;
  const pending = !channel && portfolio.class.approvalRequired;
  const rejected = channel?.status === 'rejected';
  const theme = currentTheme(channel?.themes || {}, new Date());
  let draft = state.studentDrafts.get(draftKey()) || {};
  if (config.persistLocalDrafts === true) {
    try { draft = JSON.parse(localStorage.getItem(draftKey()) || '{}'); } catch (loadError) {}
  }
  const unread = unreadFeedbackJournals();
  app.innerHTML = shell(`<div class="student-ui"><div class="page-heading student-page-heading"><div><span class="badge">${escapeHtml(portfolio.class.code)}</span><h1 class="user-content">${studentText(portfolio.class.name)}</h1></div><button id="student-classes" class="quiet" type="button">${studentText('クラス一覧へ')}</button></div>
    ${studentErrorNotice(error)}${typeof error === 'string' && error ? `<button id="reshare" class="secondary" type="button">${studentText('先生へもう一度届ける')}</button>` : ''}
    ${unread.length ? `<button id="unread-feedback" class="feedback-alert" type="button" aria-label="先生から新しいおへんじが ${unread.length}件 届いています"><span>💌</span><strong>${studentText('先生から新しいおへんじが')} ${unread.length}${studentText('件 届いています')}</strong><span>${studentText('見る')} →</span></button>` : ''}
    ${rejected ? `<div class="error">${studentText('このクラスへの参加は承認されませんでした。先生へ確認してください。')}</div>` : pending ? `<div class="notice">${studentText('参加申請を先生へ送りました。先生が承認すると、ふりかえりを書けるようになります。')}</div>` : studentWorkspace(theme, draft)}
    <section class="student-history"><div class="section-heading"><div><span class="eyebrow">MY JOURNAL</span><h2>これまでのふりかえり</h2></div><span class="muted">${(portfolio.journals || []).length}${studentText('件')}</span></div><div class="journal-list">${studentJournalCards()}</div></section>
    ${pastSelfPanel()}</div>`);
  document.getElementById('student-classes').addEventListener('click', () => {
    setWritingFocus(false);
    appBack(() => { replaceAppRoute('student-home'); renderStudentHome(); });
  });
  document.getElementById('reshare')?.addEventListener('click', async () => {
    await withError(async () => { await state.drive.shareWithUser(state.portfolioFile.id, portfolio.class.teacherEmail, 'reader'); studentToast('先生へ共有しました。'); renderPortfolio(); }, (message) => renderPortfolio(message));
  });
  if (!pending && !rejected) {
    const form = document.getElementById('journal-form');
    form.addEventListener('submit', saveJournal);
    ['theme','content'].forEach((id) => document.getElementById(id).addEventListener('input', saveDraft));
    document.querySelectorAll('[data-insert]').forEach((button) => button.addEventListener('click', () => insertText(button.dataset.insert)));
    document.querySelectorAll('[data-template]').forEach((button) => button.addEventListener('click', () => applyTemplate(button.dataset.template)));
    document.getElementById('random-starter')?.addEventListener('click', insertRandomStarter);
    document.querySelectorAll('[name="emotion"]').forEach((input) => input.addEventListener('change', saveDraft));
    document.getElementById('writing-focus')?.addEventListener('click', () => {
      const active = document.documentElement.classList.contains('writing-focus');
      if (active) appBack(() => setWritingFocus(false));
      else { pushAppRoute('student-writing-focus', { fileId: state.portfolioFile.id, classId: portfolio.class.id }); setWritingFocus(true); }
    });
    updateWritingCount();
  }
  document.querySelectorAll('[data-student-journal]').forEach((button) => button.addEventListener('click', () => {
    pushAppRoute('student-journal', { fileId: state.portfolioFile.id, classId: portfolio.class.id, journalId: button.dataset.studentJournal });
    openStudentJournal(button.dataset.studentJournal);
  }));
  document.getElementById('calendar-prev')?.addEventListener('click', () => changeStudentMonth(-1));
  document.getElementById('calendar-next')?.addEventListener('click', () => changeStudentMonth(1));
  document.getElementById('unread-feedback')?.addEventListener('click', () => {
    if (!unread[0]?.id) return;
    pushAppRoute('student-journal', { fileId: state.portfolioFile.id, classId: portfolio.class.id, journalId: unread[0].id });
    openStudentJournal(unread[0].id);
  });
  document.getElementById('past-comment-form')?.addEventListener('submit', savePastComment);
}

function studentShortDate(date = new Date()) {
  const weekdays = [['日', 'にち'], ['月', 'げつ'], ['火', 'か'], ['水', 'すい'], ['木', 'もく'], ['金', 'きん'], ['土', 'ど']];
  const weekday = weekdays[date.getDay()];
  return `${date.getMonth() + 1}${ruby('月', 'がつ')}${date.getDate()}${ruby('日', 'にち')}（${ruby(weekday[0], weekday[1])}）`;
}

function studentWorkspace(theme, draft) {
  const auxiliaryOpen = !window.matchMedia?.('(max-width: 680px), (max-height: 520px) and (orientation: landscape)').matches;
  const focusActive = document.documentElement.classList.contains('writing-focus');
  return `<section class="student-workspace">
    <aside class="student-calendar panel"><details class="auxiliary-details" ${auxiliaryOpen ? 'open' : ''}><summary class="auxiliary-summary">📅 ${studentText('記録カレンダー')}</summary><div class="auxiliary-content"><div class="calendar-heading"><button id="calendar-prev" class="icon-button" type="button" aria-label="前の月">‹</button><strong>${state.studentMonth.getFullYear()}${ruby('年', 'ねん')} ${state.studentMonth.getMonth() + 1}${ruby('月', 'がつ')}</strong><button id="calendar-next" class="icon-button" type="button" aria-label="次の月">›</button></div>${studentCalendar()}</div></details></aside>
    <section class="notebook panel"><div class="notebook-top"><div><span class="eyebrow">TODAY'S JOURNAL</span><h2>${studentText('今日のふりかえり')}</h2></div><div class="notebook-controls"><span class="notebook-date">${studentShortDate()}</span><button id="writing-focus" class="focus-button" type="button" aria-label="${focusActive ? 'もとに戻す' : '広く書く'}" aria-pressed="${focusActive}">${focusActive ? `↙ ${studentText('もとに戻す')}` : `↗ ${studentText('広く書く')}`}</button></div></div>
      <div class="theme-banner"><span>${studentText('今日のテーマ')}</span><strong class="user-content">${studentText(theme)}</strong></div><form id="journal-form">
      <label class="visually-hidden"><span>テーマ</span><input id="theme" maxlength="200" value="${escapeHtml(draft.theme || theme)}"></label>
      <label><span>${studentText('自分のことばで書こう')}</span><textarea id="content" class="lined-paper" maxlength="20000" required placeholder="できたこと、かんがえたこと、つぎにやってみたいことをかこう">${escapeHtml(draft.content || '')}</textarea></label>
      <div class="notebook-options"><fieldset class="emotion-fieldset"><legend><strong>${studentText('今の気持ち')}</strong> <span class="muted small">（えらばなくてもOK）</span></legend><div class="emotion-row labeled-emotions" role="radiogroup" aria-label="いまの気持ち">${[['😊','うれしい'],['😠','くやしい'],['💡','なるほど'],['🤔','もやもや']].map(([emotion, label]) => `<label><input type="radio" name="emotion" value="${emotion}" ${draft.emotion === emotion ? 'checked' : ''}><span><b>${emotion}</b><small>${label}</small></span></label>`).join('')}</div></fieldset>
      <details class="attachment"><summary>📎 ${studentText('作品や写真をつける')}</summary><label><span>${studentText('画像・作品（任意）')}</span><input id="image" type="file" accept="image/*"></label><p class="field-note">${studentText('見やすい大きさにして先生へ届けます。')}</p></details></div>
      <div class="notebook-submit"><span id="writing-count" class="writing-count" aria-live="polite">0${studentText('文字')}</span><button class="primary submit-journal" type="submit" aria-label="できた！ 先生にとどける">できた！ ${studentText('先生')}にとどける</button></div>
    </form></section>
    <aside class="writing-tools panel"><details class="auxiliary-details" ${auxiliaryOpen ? 'open' : ''}><summary class="auxiliary-summary">✏️ ${studentText('書き方サポート')}</summary><div class="auxiliary-content"><p class="muted small">${studentText('ボタンをおすと、文の書き出しが入ります。')}</p>
      <button id="random-starter" class="secondary wide" type="button" aria-label="おまかせ書き出し">🎲 おまかせ${studentText('書き出し')}</button>
      <details open><summary>かたをえらぶ</summary><div class="template-grid">${JOURNAL_TEMPLATES.map(([label, value]) => `<button class="template-button" data-template="${escapeHtml(value)}" type="button">${studentText(label)}</button>`).join('')}</div></details>
      ${HINT_GROUPS.map(([label, hints], index) => `<details ${index === 0 ? 'open' : ''}><summary>${studentText(label)}のヒント</summary><div class="hint-list">${hints.map((hint) => `<button class="hint-button" data-insert="${escapeHtml(hint)}" type="button">${studentText(hint)}</button>`).join('')}</div></details>`).join('')}
    </div></details></aside>
  </section>`;
}

function studentCalendar() {
  const year = state.studentMonth.getFullYear();
  const month = state.studentMonth.getMonth();
  const firstDay = new Date(year, month, 1).getDay();
  const lastDate = new Date(year, month + 1, 0).getDate();
  const journals = state.portfolio.journals || [];
  const cells = Array.from({ length: firstDay }, () => '<span class="calendar-cell empty-day"></span>');
  for (let day = 1; day <= lastDate; day += 1) {
    const key = todayKey(new Date(year, month, day));
    const journal = [...journals].reverse().find((item) => todayKey(new Date(item.createdAt)) === key);
    const returned = journal && state.channel?.feedback?.[journal.id]?.returned;
    cells.push(journal ? `<button class="calendar-cell has-journal ${returned ? 'has-feedback' : ''}" data-student-journal="${escapeHtml(journal.id)}" type="button"><span>${day}</span><small>${returned ? '💬' : '●'}</small></button>` : `<span class="calendar-cell"><span>${day}</span></span>`);
  }
  const weekdays = [['日', 'にち'], ['月', 'げつ'], ['火', 'か'], ['水', 'すい'], ['木', 'もく'], ['金', 'きん'], ['土', 'ど']];
  return `<div class="calendar-week">${weekdays.map(([day, reading]) => `<strong>${ruby(day, reading)}</strong>`).join('')}</div><div class="calendar-grid">${cells.join('')}</div><p class="calendar-key"><span>● ${studentText('書いた日')}</span><span>💬 おへんじ</span></p>`;
}

function changeStudentMonth(offset) {
  state.studentMonth = new Date(state.studentMonth.getFullYear(), state.studentMonth.getMonth() + offset, 1);
  renderPortfolio();
}

function studentJournalCards() {
  const journals = state.portfolio.journals || [];
  return journals.length ? [...journals].reverse().map((journal) => {
    const feedback = state.channel?.feedback?.[journal.id];
    const unread = feedback?.returned && !readJournalIds().has(journal.id);
    return `<button class="journal-card student-journal-card ${unread ? 'unread' : ''}" data-student-journal="${escapeHtml(journal.id)}" type="button"><div class="journal-meta"><span>${escapeHtml(formatDate(journal.createdAt))}</span><span>${escapeHtml(journal.emotion || '')}</span></div><h3 class="user-content">${studentText(journal.theme || '')}</h3><div class="journal-body clamp user-content">${escapeHtml(journal.content)}</div><div class="journal-footer">${journal.imageFileId ? `<span>📎 ${studentText('作品')}つき</span>` : '<span></span>'}${feedback?.returned ? `<span class="badge success-badge">${unread ? 'NEW ' : ''}${escapeHtml(feedback.stamp || '💬')} おへんじ</span>` : ''}</div></button>`;
  }).join('') : '<div class="empty">まだふりかえりはありません。</div>';
}

function highlightedJournalHtml(journal, feedback) {
  const highlights = (feedback?.highlights || []).filter((item) => Number.isInteger(item.start) && Number.isInteger(item.end) && item.start >= 0 && item.end > item.start && item.end <= journal.content.length)
    .sort((a, b) => a.start - b.start || a.end - b.end);
  let cursor = 0;
  const parts = [];
  for (const item of highlights) {
    if (item.start < cursor) continue;
    parts.push(escapeHtml(journal.content.slice(cursor, item.start)));
    parts.push(`<mark class="teacher-highlight" title="${escapeHtml(item.comment || '先生が注目したところ')}">${escapeHtml(journal.content.slice(item.start, item.end))}<span>${escapeHtml(item.stamp || '⭐')}</span></mark>`);
    cursor = item.end;
  }
  parts.push(escapeHtml(journal.content.slice(cursor)));
  return parts.join('');
}

function studentHighlightNotes(feedback) {
  const items = (feedback?.highlights || []).filter((item) => item.comment || item.stamp);
  return items.length ? `<div class="highlight-notes"><h3>${studentText('文章についたおへんじ')}</h3>${items.map((item) => `<div><strong>${escapeHtml(item.stamp || '⭐')} 「${escapeHtml(item.text || '')}」</strong><p class="user-content">${escapeHtml(item.comment || '')}</p></div>`).join('')}</div>` : '';
}

function openStudentJournal(journalId) {
  if (!journalId) return;
  const journal = (state.portfolio.journals || []).find((item) => item.id === journalId);
  if (!journal) return;
  const feedback = state.channel?.feedback?.[journal.id];
  if (feedback?.returned) markJournalRead(journal.id);
  closeStudentDialog();
  const dialog = document.createElement('dialog');
  dialog.className = 'journal-dialog';
  dialog.innerHTML = `<article class="student-ui"><form method="dialog"><button class="dialog-close" aria-label="閉じる">×</button></form><div class="journal-meta"><span>${escapeHtml(formatDate(journal.createdAt))}</span><span>${escapeHtml(journal.emotion || '')}</span></div><h2 class="user-content">${studentText(journal.theme || 'ふりかえり')}</h2><div class="journal-body lined-reading user-content">${highlightedJournalHtml(journal, feedback)}</div>${journal.imageFileId ? `<div class="image-slot" data-file="${escapeHtml(journal.imageFileId)}"><p>${studentText('作品を読み込んでいます…')}</p></div>` : ''}${feedback?.returned ? `${studentHighlightNotes(feedback)}<div class="feedback-box large"><strong>${escapeHtml(feedback.stamp || '💬')} ${studentText('先生からのおへんじ')}</strong><p class="user-content">${escapeHtml(feedback.comment)}</p></div>` : `<div class="muted dialog-note">${studentText('先生からのおへんじを待っています。')}</div>`}${journal.pastComment ? `<div class="notice"><strong>${studentText('今の自分から')}:</strong> <span class="user-content">${escapeHtml(journal.pastComment)}</span></div>` : ''}</article>`;
  document.body.appendChild(dialog);
  dialog.addEventListener('close', () => {
    const historyClose = dialog.dataset.historyClose === 'true';
    dialog.remove();
    if (!historyClose) appBack(() => renderPortfolio());
  });
  dialog.addEventListener('click', (event) => { if (event.target === dialog) dialog.close(); });
  dialog.showModal();
  loadImages();
}

function pastSelfPanel() {
  const cutoff = Date.now() - 30 * 86400000;
  const old = (state.portfolio.journals || []).filter((journal) => new Date(journal.createdAt).getTime() < cutoff);
  if (!old.length) return '';
  const journal = old[Math.floor(Math.random() * old.length)];
  return `<section class="panel past-self-panel"><h2>${studentText('過去の自分と対話')}</h2><blockquote class="user-content">${escapeHtml(journal.content)}</blockquote><form id="past-comment-form" data-journal="${escapeHtml(journal.id)}"><label><span>${studentText('今の自分からひとこと')}</span><textarea id="past-comment" maxlength="4000">${escapeHtml(journal.pastComment || '')}</textarea></label><button class="secondary" type="submit">メッセージを${studentText('保存')}</button></form></section>`;
}

function saveDraft() {
  const draft = { theme: document.getElementById('theme').value, content: document.getElementById('content').value, emotion: document.querySelector('[name="emotion"]:checked')?.value || '' };
  state.studentDrafts.set(draftKey(), draft);
  if (config.persistLocalDrafts === true) {
    try { localStorage.setItem(draftKey(), JSON.stringify(draft)); } catch (error) {}
  }
  updateWritingCount();
}

function updateWritingCount() {
  const count = document.getElementById('content')?.value.length || 0;
  const output = document.getElementById('writing-count');
  if (output) output.innerHTML = `${count.toLocaleString('ja-JP')}${studentText('文字')}`;
}

function setWritingFocus(active) {
  document.documentElement.classList.toggle('writing-focus', Boolean(active));
  const button = document.getElementById('writing-focus');
  if (!button) return;
  button.setAttribute('aria-pressed', String(Boolean(active)));
  button.setAttribute('aria-label', active ? 'もとに戻す' : '広く書く');
  button.innerHTML = active ? `↙ ${studentText('もとに戻す')}` : `↗ ${studentText('広く書く')}`;
  if (active) document.getElementById('content')?.focus({ preventScroll: true });
}

function insertText(text) {
  const area = document.getElementById('content');
  const start = area.selectionStart;
  area.value = area.value.slice(0, start) + text + area.value.slice(area.selectionEnd);
  area.focus();
  area.selectionStart = area.selectionEnd = start + text.length;
  saveDraft();
}

function applyTemplate(template) {
  const area = document.getElementById('content');
  if (area.value.trim() && !window.confirm('今書いている内容を、選んだ「かた」に入れ替えますか？')) return;
  area.value = template;
  area.focus();
  area.selectionStart = area.selectionEnd = template.length;
  saveDraft();
}

function insertRandomStarter() {
  insertText(RANDOM_STARTERS[Math.floor(Math.random() * RANDOM_STARTERS.length)]);
}

function celebrateSubmission() {
  if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) return;
  const layer = document.createElement('div');
  layer.className = 'confetti-layer';
  layer.setAttribute('aria-hidden', 'true');
  layer.innerHTML = Array.from({ length: 42 }, (_, index) => `<i style="--x:${(index * 47) % 100}%;--delay:${(index % 9) * .035}s;--spin:${(index % 2 ? 1 : -1) * (240 + index * 11)}deg;--color:${['#f97316','#fbbf24','#38bdf8','#34d399'][index % 4]}"></i>`).join('');
  document.body.appendChild(layer);
  setTimeout(() => layer.remove(), 1800);
}

async function saveJournal(event) {
  event.preventDefault();
  const theme = document.getElementById('theme').value;
  const content = document.getElementById('content').value.trim();
  const emotion = document.querySelector('[name="emotion"]:checked')?.value || '';
  const inputFile = document.getElementById('image').files[0] || null;
  if (!content) return;
  const journalId = crypto.randomUUID();
  setStudentBusy('ふりかえりを保存しています…');
  await withError(async () => {
    let imageFileId = null;
    let imageName = null;
    if (inputFile) {
      const image = await resizeImage(inputFile);
      const created = await state.drive.createJournalImage({ classId: state.portfolio.class.id, journalId, studentName: state.portfolio.student.name, file: image });
      await state.drive.shareWithUser(created.id, state.portfolio.class.teacherEmail, 'reader');
      imageFileId = created.id;
      imageName = inputFile.name;
    }
    const updated = appendJournal(state.portfolio, { id: journalId, theme, content, emotion, imageFileId, imageName });
    await updateDocument(state.portfolioFile, updated);
    state.portfolio = updated;
    state.studentDrafts.delete(draftKey());
    try { localStorage.removeItem(draftKey()); } catch (error) {}
    studentToast('ふりかえりを提出しました。');
    renderPortfolio();
    celebrateSubmission();
  }, (message) => renderPortfolio(message));
}

async function savePastComment(event) {
  event.preventDefault();
  const updated = updatePastComment(state.portfolio, event.currentTarget.dataset.journal, document.getElementById('past-comment').value);
  setStudentBusy('メッセージを保存しています…');
  await withError(async () => { await updateDocument(state.portfolioFile, updated); state.portfolio = updated; studentToast('メッセージを保存しました。'); renderPortfolio(); }, (message) => renderPortfolio(message));
}

async function resizeImage(file) {
  if (!file.type.startsWith('image/')) throw new Error('画像ファイルを選んでください。');
  const bitmap = await createImageBitmap(file);
  const scale = Math.min(1, 1600 / Math.max(bitmap.width, bitmap.height));
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.round(bitmap.width * scale));
  canvas.height = Math.max(1, Math.round(bitmap.height * scale));
  canvas.getContext('2d').drawImage(bitmap, 0, 0, canvas.width, canvas.height);
  bitmap.close();
  const blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/jpeg', .82));
  if (!blob) throw new Error('画像を変換できませんでした。');
  return new File([blob], file.name.replace(/\.[^.]+$/, '') + '.jpg', { type: 'image/jpeg' });
}

async function loadImages() {
  document.querySelectorAll('.image-slot').forEach(async (slot) => {
    try { const blob = await state.drive.getBlob(slot.dataset.file); const url = URL.createObjectURL(blob); slot.innerHTML = '<img class="journal-image" alt="児童が添付した成果物">'; slot.querySelector('img').src = url; }
    catch (error) { slot.innerHTML = `<p class="error small">${slot.closest('.student-ui') ? studentText('画像を表示できませんでした。') : '画像を表示できませんでした。'}</p>`; }
  });
}

function renderQr(value) {
  const target = document.getElementById('qr');
  if (!target || typeof window.qrcode !== 'function') { if (target) target.textContent = 'QRコードを生成できませんでした。'; return; }
  const qr = window.qrcode(0, 'M');
  qr.addData(value);
  qr.make();
  target.innerHTML = qr.createSvgTag({ cellSize: 5, margin: 2, scalable: true });
}

function downloadQrCode(className) {
  const svg = document.querySelector('#qr svg');
  if (!svg) return toast('QRコードの準備ができていません。');
  const source = new XMLSerializer().serializeToString(svg);
  downloadBlob(new Blob([source], { type: 'image/svg+xml;charset=utf-8' }), `${className}_招待QR.svg`);
}

function ensureQr() {
  if (typeof window.qrcode === 'function') return Promise.resolve();
  if (ensureQr.promise) return ensureQr.promise;
  ensureQr.promise = new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = './qrcode.js';
    script.onload = resolve;
    script.onerror = reject;
    document.head.appendChild(script);
  });
  return ensureQr.promise;
}

async function copy(value, message) {
  try { await navigator.clipboard.writeText(value); toast(message); }
  catch (error) { window.prompt('この内容をコピーしてください', value); }
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename.replace(/[\\/:*?"<>|]/g, '_');
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

function formatDate(value) {
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? '' : new Intl.DateTimeFormat('ja-JP', { dateStyle: 'medium', timeStyle: 'short' }).format(date);
}

function formatTime(value) {
  const date = value instanceof Date ? value : new Date(value);
  return Number.isNaN(date.getTime()) ? '' : new Intl.DateTimeFormat('ja-JP', { hour: '2-digit', minute: '2-digit', second: '2-digit' }).format(date);
}

function initPwaControls() {
  const installButton = document.getElementById('pwa-install');
  const updateBox = document.getElementById('pwa-update');
  const refreshButton = document.getElementById('pwa-refresh');
  const laterButton = document.getElementById('pwa-later');
  if (!installButton || !updateBox) return;

  const showInstall = () => { installButton.hidden = !window.__deferredInstallPrompt; };
  showInstall();
  window.addEventListener('pwa-installable', showInstall);
  window.addEventListener('pwa-installed', () => { installButton.hidden = true; toast('アプリをインストールしました。'); });
  installButton.addEventListener('click', async () => {
    const promptEvent = window.__deferredInstallPrompt;
    if (!promptEvent) return;
    promptEvent.prompt();
    await promptEvent.userChoice;
    window.__deferredInstallPrompt = null;
    installButton.hidden = true;
  });

  let waitingWorker = null;
  let refreshing = false;
  const showUpdate = (worker) => {
    waitingWorker = worker;
    updateBox.hidden = false;
  };
  const watchRegistration = (registration) => {
    if (registration.waiting && navigator.serviceWorker.controller) showUpdate(registration.waiting);
    registration.addEventListener('updatefound', () => {
      const worker = registration.installing;
      worker?.addEventListener('statechange', () => {
        if (worker.state === 'installed' && navigator.serviceWorker.controller) showUpdate(worker);
      });
    });
  };
  const registrationReady = () => { if (window.__pwaRegistration) watchRegistration(window.__pwaRegistration); };
  window.addEventListener('pwa-registration-ready', registrationReady);
  registrationReady();
  refreshButton?.addEventListener('click', () => waitingWorker?.postMessage({ type: 'SKIP_WAITING' }));
  laterButton?.addEventListener('click', () => { updateBox.hidden = true; });
  navigator.serviceWorker?.addEventListener('controllerchange', () => {
    if (refreshing) return;
    refreshing = true;
    location.reload();
  });
}

app.addEventListener('click', (event) => {
  if (!event.target.closest('[data-change-account]')) return;
  clearSession();
  forgetRole();
  resetAppRoute('home', {}, location.pathname + location.search);
  state.tokenClient = null;
  state.sharedTokenClient = null;
  state.forceAccountSelection = true;
  renderLogin('別のGoogleアカウントで続けてください。');
});

initPwaControls();
if (!appHistoryState()) resetAppRoute('home');
window.addEventListener('popstate', (event) => {
  const entry = event.state?.app === APP_HISTORY_ID ? event.state : null;
  if (!entry) return;
  renderHistoryRoute(entry);
});
document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape' && document.documentElement.classList.contains('writing-focus')) appBack(() => setWritingFocus(false));
});

if (config.persistLocalDrafts !== true) {
  try {
    for (let index = localStorage.length - 1; index >= 0; index -= 1) {
      const key = localStorage.key(index);
      if (key?.startsWith('rj_draft_v2_')) localStorage.removeItem(key);
    }
  } catch (error) {}
}

if (restoreSession()) {
  setBusy('前回の画面をひらいています…');
  resolveEntryRoute();
} else {
  renderLogin();
}
