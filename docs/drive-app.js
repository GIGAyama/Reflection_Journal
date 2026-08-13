import {
  appendJournal,
  computeClassId,
  createClassRecord,
  createPortfolio,
  decodeInvite,
  inviteUrl,
  isEmail,
  normalizeClassCode,
  randomClassCode
} from './drive-core.js';
import { DriveApiError, DriveClient } from './drive-api.js';

const app = document.getElementById('app');
const toastElement = document.getElementById('toast');
const config = window.APP_CONFIG || {};
const state = { user: null, drive: null, invite: null, tokenClient: null, portfolio: null, portfolioFile: null };

const escapeHtml = (value) => String(value ?? '').replace(/[&<>'"]/g, (char) => ({
  '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;'
})[char]);

function toast(message) {
  toastElement.textContent = message;
  toastElement.classList.add('visible');
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => toastElement.classList.remove('visible'), 3000);
}

function setBusy(message = 'Google Driveと通信しています…') {
  app.innerHTML = `<section class="center-screen"><div class="loader" aria-hidden="true"></div><p>${escapeHtml(message)}</p></section>`;
}

function friendlyError(error) {
  console.error(error);
  if (error instanceof DriveApiError) return `${error.message}${error.detail ? `<br><span class="small">詳細: ${escapeHtml(error.detail)}</span>` : ''}`;
  return escapeHtml(error?.message || '予期しないエラーが発生しました。');
}

async function withError(action, fallback) {
  try { return await action(); } catch (error) {
    if (error instanceof DriveApiError && error.status === 401) return renderLogin(error.message);
    fallback(friendlyError(error));
    return null;
  }
}

function shell(content) {
  return `
    <header class="topbar">
      <div class="brand"><span aria-hidden="true">📔</span><span>ふりかえりジャーナル</span></div>
      <div class="account"><strong>${escapeHtml(state.user?.name || '')}</strong><span>${escapeHtml(state.user?.email || '')}</span></div>
    </header>
    <div class="page">${content}</div>`;
}

function renderLogin(error = '') {
  app.innerHTML = `
    <section class="center-screen">
      <div class="login-card">
        <div class="app-logo" aria-hidden="true">📔</div>
        <h1>Google Driveで<br>ふりかえりをつなぐ</h1>
        <p>成果物は児童自身のGoogle Driveへ保存され、先生へ直接共有されます。学校ごとのGASデプロイは不要です。</p>
        ${error ? `<div class="error">${escapeHtml(error)}</div>` : ''}
        <button id="login" class="primary wide" type="button">Googleアカウントで続ける</button>
        <p class="muted small">このアプリが作成・選択したファイルだけにアクセスします。アクセストークンは端末へ保存しません。</p>
      </div>
    </section>`;
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
  if (!config.googleClientId) return renderLogin('OAuthクライアントIDが設定されていません。');
  setBusy('Googleログインをひらいています…');
  try {
    await loadGoogleIdentity();
    if (!state.tokenClient) {
      state.tokenClient = google.accounts.oauth2.initTokenClient({
        client_id: config.googleClientId,
        scope: 'openid email profile https://www.googleapis.com/auth/drive.file',
        include_granted_scopes: true,
        callback: handleToken
      });
    }
    state.tokenClient.requestAccessToken({ prompt: '' });
  } catch (error) { renderLogin(error.message); }
}

async function handleToken(response) {
  if (!response?.access_token) return renderLogin(response?.error_description || 'Googleログインがキャンセルされました。');
  setBusy('アカウントを確認しています…');
  try {
    const userResponse = await fetch('https://openidconnect.googleapis.com/v1/userinfo', {
      headers: { Authorization: `Bearer ${response.access_token}` }
    });
    if (!userResponse.ok) throw new Error('Googleアカウント情報を確認できませんでした。');
    const profile = await userResponse.json();
    state.user = { email: profile.email, name: profile.name || profile.email };
    state.drive = new DriveClient(response.access_token);
    await resolveEntryRoute();
  } catch (error) { renderLogin(error.message); }
}

async function resolveEntryRoute() {
  const encoded = location.hash.match(/^#join=([^&]+)/)?.[1];
  if (encoded) {
    try { state.invite = await decodeInvite(encoded); return renderJoin(); }
    catch (error) { return renderHome(error.message); }
  }
  return renderHome();
}

function renderHome(error = '') {
  app.innerHTML = shell(`
    <div class="page-heading"><div><span class="badge">Drive版</span><h1>どの使い方をしますか？</h1></div></div>
    ${error ? `<div class="error">${escapeHtml(error)}</div>` : ''}
    <div class="grid">
      <button class="item-card" id="teacher-home" type="button">
        <h2>先生として使う</h2><p class="muted">クラスを作成し、児童を招待して、共有された成果物を閲覧します。</p>
      </button>
      <button class="item-card" id="student-home" type="button">
        <h2>児童として使う</h2><p class="muted">参加済みのクラスを開くか、先生のメールアドレスとクラスコードで参加します。</p>
      </button>
    </div>`);
  document.getElementById('teacher-home').addEventListener('click', renderTeacherHome);
  document.getElementById('student-home').addEventListener('click', renderStudentHome);
}

async function renderTeacherHome(error = '') {
  setBusy('先生のクラスを探しています…');
  const files = await withError(() => state.drive.listClasses(), (message) => renderHome(message));
  if (!files) return;
  const classes = (await Promise.all(files.map(async (file) => {
    try { return { file, record: await state.drive.getJson(file.id) }; } catch (error) { return null; }
  }))).filter(Boolean);
  app.innerHTML = shell(`
    <div class="page-heading"><div><span class="badge">先生</span><h1>クラス</h1></div><button id="back" class="quiet" type="button">使い方を変える</button></div>
    ${error ? `<div class="error">${error}</div>` : ''}
    <section class="panel">
      <h2>新しいクラスを作る</h2>
      <form id="create-class">
        <label><span>クラス名</span><input id="class-name" maxlength="80" required placeholder="例：5年1組"></label>
        <button class="primary" type="submit">クラスを作成</button>
      </form>
    </section>
    <section>
      <h2>作成済みのクラス</h2>
      ${classes.length ? `<div class="grid">${classes.map(({ file, record }) => `
        <button class="item-card class-item" data-file="${escapeHtml(file.id)}" type="button">
          <span class="badge">${escapeHtml(record.classCode)}</span>
          <h3>${escapeHtml(record.className)}</h3>
          <p class="muted small">更新: ${escapeHtml(formatDate(file.modifiedTime))}</p>
        </button>`).join('')}</div>` : '<div class="empty">まだクラスはありません。</div>'}
    </section>`);
  document.getElementById('back').addEventListener('click', renderHome);
  document.getElementById('create-class').addEventListener('submit', createClass);
  document.querySelectorAll('.class-item').forEach((button) => button.addEventListener('click', () => {
    const item = classes.find(({ file }) => file.id === button.dataset.file);
    renderTeacherClass(item);
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
    const record = createClassRecord({ classId, classCode, className, teacher: state.user });
    const file = await state.drive.createClass(record);
    await renderTeacherClass({ file, record });
  }, (message) => renderTeacherHome(message));
}

function publicEntryUrl() {
  if (config.publicEntryUrl) return config.publicEntryUrl;
  return new URL('./', location.href).href;
}

async function renderTeacherClass(item, error = '') {
  const { file, record } = item;
  setBusy('児童の成果物を探しています…');
  const portfolioFiles = await withError(() => state.drive.listSharedPortfolios(record.classId), (message) => renderTeacherHome(message));
  if (!portfolioFiles) return;
  const portfolios = (await Promise.all(portfolioFiles.map(async (portfolioFile) => {
    try { return { file: portfolioFile, record: await state.drive.getJson(portfolioFile.id) }; }
    catch (loadError) { return null; }
  }))).filter((entry) => entry?.record?.class?.id === record.classId);
  const link = inviteUrl(publicEntryUrl(), {
    classCode: record.classCode,
    className: record.className,
    teacherEmail: record.teacher.email,
    teacherName: record.teacher.name
  });
  app.innerHTML = shell(`
    <div class="page-heading"><div><span class="badge">${escapeHtml(record.classCode)}</span><h1>${escapeHtml(record.className)}</h1></div><button id="classes" class="quiet" type="button">クラス一覧へ</button></div>
    ${error ? `<div class="error">${error}</div>` : ''}
    <section class="panel invite-layout">
      <div>
        <h2>児童を招待する</h2>
        <p class="muted">QRコードまたは専用URLを配ってください。児童は自分のDriveに保存し、この先生へ直接共有します。</p>
        <div class="class-code">${escapeHtml(record.classCode)}</div>
        <div class="url-box" id="invite-url">${escapeHtml(link)}</div>
        <div class="button-row" style="margin-top:12px">
          <button id="copy-url" class="primary" type="button">専用URLをコピー</button>
          <button id="copy-text" class="secondary" type="button">案内文をコピー</button>
        </div>
      </div>
      <div id="qr" class="qr-box" aria-label="児童招待用QRコード"></div>
    </section>
    <section>
      <h2>共有された成果物 <span class="muted small">${portfolios.length}人</span></h2>
      ${portfolios.length ? `<div class="grid">${portfolios.map(({ file: portfolioFile, record: portfolio }) => `
        <button class="item-card portfolio-item" data-file="${escapeHtml(portfolioFile.id)}" type="button">
          <h3>${escapeHtml(portfolio.student?.name || '名前未設定')}</h3>
          <p>${portfolio.journals?.length || 0}件のふりかえり</p>
          <p class="muted small">更新: ${escapeHtml(formatDate(portfolio.updatedAt || portfolioFile.modifiedTime))}</p>
        </button>`).join('')}</div>` : '<div class="empty">まだ共有された成果物はありません。児童が最初の参加を完了すると、ここに表示されます。</div>'}
    </section>`);
  document.getElementById('classes').addEventListener('click', renderTeacherHome);
  document.getElementById('copy-url').addEventListener('click', () => copy(link, '専用URLをコピーしました。'));
  document.getElementById('copy-text').addEventListener('click', () => copy(`「${record.className}」のふりかえりジャーナルに参加してください。\nクラスコード: ${record.classCode}\n${link}`, '案内文をコピーしました。'));
  renderQr(link);
  document.querySelectorAll('.portfolio-item').forEach((button) => button.addEventListener('click', () => {
    const portfolio = portfolios.find(({ file: portfolioFile }) => portfolioFile.id === button.dataset.file);
    renderTeacherPortfolio(item, portfolio);
  }));
}

async function renderTeacherPortfolio(classItem, portfolioItem) {
  const portfolio = portfolioItem.record;
  app.innerHTML = shell(`
    <div class="page-heading"><div><span class="badge">児童の成果物</span><h1>${escapeHtml(portfolio.student?.name || '名前未設定')}</h1></div><button id="class-back" class="quiet" type="button">クラスへ戻る</button></div>
    <section class="journal-list">
      ${(portfolio.journals || []).length ? [...portfolio.journals].reverse().map((journal) => `
        <article class="journal-card">
          <div class="journal-meta"><span>${escapeHtml(formatDate(journal.createdAt))}</span><span aria-label="気持ち">${escapeHtml(journal.emotion || '')}</span></div>
          ${journal.theme ? `<h2>${escapeHtml(journal.theme)}</h2>` : ''}
          <div class="journal-body">${escapeHtml(journal.content)}</div>
          ${journal.imageFileId ? `<div class="image-slot" data-file="${escapeHtml(journal.imageFileId)}"><p class="muted small">画像を読み込んでいます…</p></div>` : ''}
        </article>`).join('') : '<div class="empty">ふりかえりはまだありません。</div>'}
    </section>`);
  document.getElementById('class-back').addEventListener('click', () => renderTeacherClass(classItem));
  document.querySelectorAll('.image-slot').forEach(async (slot) => {
    try {
      const blob = await state.drive.getBlob(slot.dataset.file);
      const url = URL.createObjectURL(blob);
      slot.innerHTML = `<img class="journal-image" alt="児童が添付した成果物">`;
      slot.querySelector('img').src = url;
    } catch (error) { slot.innerHTML = '<p class="error small">画像を表示できませんでした。</p>'; }
  });
}

async function renderStudentHome(error = '') {
  setBusy('参加済みのクラスを探しています…');
  const files = await withError(() => state.drive.listOwnPortfolios(), (message) => renderHome(message));
  if (!files) return;
  const portfolios = (await Promise.all(files.map(async (file) => {
    try { return { file, record: await state.drive.getJson(file.id) }; } catch (loadError) { return null; }
  }))).filter(Boolean);
  app.innerHTML = shell(`
    <div class="page-heading"><div><span class="badge">児童</span><h1>参加しているクラス</h1></div><button id="back" class="quiet" type="button">使い方を変える</button></div>
    ${error ? `<div class="error">${error}</div>` : ''}
    ${portfolios.length ? `<div class="grid">${portfolios.map(({ file, record }) => `
      <button class="item-card student-portfolio" data-file="${escapeHtml(file.id)}" type="button">
        <span class="badge">${escapeHtml(record.class?.code || '')}</span><h3>${escapeHtml(record.class?.name || 'クラス')}</h3>
        <p>${record.journals?.length || 0}件のふりかえり</p>
      </button>`).join('')}</div>` : '<div class="empty">参加済みのクラスはありません。先生のQRコードまたは専用URLから参加できます。</div>'}
    <section class="panel" style="margin-top:20px">
      <h2>コードで参加する</h2>
      <p class="muted">専用URLを使えない場合は、先生のメールアドレス・クラス名・コードを入力してください。</p>
      <form id="manual-join">
        <label><span>先生のGoogleメールアドレス</span><input id="teacher-email" type="email" required></label>
        <label><span>クラス名</span><input id="manual-class-name" maxlength="80" required></label>
        <label><span>クラスコード</span><input id="manual-code" maxlength="10" autocapitalize="characters" required></label>
        <button class="primary" type="submit">クラスを確認</button>
      </form>
    </section>`);
  document.getElementById('back').addEventListener('click', renderHome);
  document.querySelectorAll('.student-portfolio').forEach((button) => button.addEventListener('click', () => {
    const item = portfolios.find(({ file }) => file.id === button.dataset.file);
    openPortfolio(item);
  }));
  document.getElementById('manual-join').addEventListener('submit', async (event) => {
    event.preventDefault();
    const teacherEmail = document.getElementById('teacher-email').value.trim().toLowerCase();
    const className = document.getElementById('manual-class-name').value.trim();
    const classCode = normalizeClassCode(document.getElementById('manual-code').value);
    if (!isEmail(teacherEmail) || classCode.length < 6) return renderStudentHome('メールアドレスまたはクラスコードを確認してください。');
    state.invite = { teacherEmail, teacherName: '', className, classCode, classId: await computeClassId(teacherEmail, classCode) };
    renderJoin();
  });
}

async function renderJoin(error = '') {
  const invite = state.invite;
  setBusy('参加状況を確認しています…');
  const files = await withError(() => state.drive.listOwnPortfolios(invite.classId), (message) => renderHome(message));
  if (!files) return;
  if (files.length) {
    const record = await withError(() => state.drive.getJson(files[0].id), (message) => renderHome(message));
    if (record) return openPortfolio({ file: files[0], record });
  }
  app.innerHTML = shell(`
    <div class="page-heading"><div><span class="badge">招待</span><h1>${escapeHtml(invite.className)}</h1></div><button id="cancel" class="quiet" type="button">戻る</button></div>
    ${error ? `<div class="error">${error}</div>` : ''}
    <section class="panel">
      <p>先生: <strong>${escapeHtml(invite.teacherName || invite.teacherEmail)}</strong></p>
      <p>クラスコード: <strong>${escapeHtml(invite.classCode)}</strong></p>
      <form id="join-class">
        <label><span>先生に表示する名前</span><input id="student-name" maxlength="80" required value="${escapeHtml(state.user.name || '')}"></label>
        <button class="primary wide" type="submit">このクラスに参加する</button>
      </form>
      <p class="muted small">参加すると、このクラス用のポートフォリオがあなたのGoogle Driveに作られ、先生へ閲覧共有されます。</p>
    </section>`);
  document.getElementById('cancel').addEventListener('click', () => { history.replaceState(null, '', location.pathname + location.search); state.invite = null; renderStudentHome(); });
  document.getElementById('join-class').addEventListener('submit', joinClass);
}

async function joinClass(event) {
  event.preventDefault();
  const name = document.getElementById('student-name').value.trim();
  if (!name) return;
  setBusy('クラスに参加しています…');
  await withError(async () => {
    const record = createPortfolio({ invite: state.invite, student: { ...state.user, name } });
    const file = await state.drive.createPortfolio(record);
    try {
      await state.drive.shareWithUser(file.id, state.invite.teacherEmail, 'reader');
    } catch (error) {
      return openPortfolio({ file, record }, 'ポートフォリオは作成されましたが、先生への共有が学校の設定で拒否されました。「先生へ再共有」を押すか、管理者へ共有設定を確認してください。');
    }
    history.replaceState(null, '', location.pathname + location.search);
    await openPortfolio({ file, record });
  }, (message) => renderJoin(message));
}

async function openPortfolio(item, error = '') {
  state.portfolioFile = item.file;
  state.portfolio = item.record;
  const portfolio = item.record;
  app.innerHTML = shell(`
    <div class="page-heading"><div><span class="badge">${escapeHtml(portfolio.class.code)}</span><h1>${escapeHtml(portfolio.class.name)}</h1></div><button id="student-classes" class="quiet" type="button">クラス一覧へ</button></div>
    ${error ? `<div class="error">${escapeHtml(error)}</div><button id="reshare" class="secondary" type="button">先生へ再共有</button>` : ''}
    <section class="panel">
      <h2>今日のふりかえり</h2>
      <form id="journal-form">
        <label><span>テーマ（任意）</span><input id="theme" maxlength="200" placeholder="例：今日の算数で分かったこと"></label>
        <label><span>ふりかえり</span><textarea id="content" maxlength="20000" required placeholder="できたこと、考えたこと、次にやってみたいことを書こう"></textarea></label>
        <span><strong>いまの気持ち（任意）</strong></span>
        <div class="emotion-row" role="radiogroup" aria-label="いまの気持ち">
          ${['😊','🙂','😐','😕','💪'].map((emotion) => `<label><input type="radio" name="emotion" value="${emotion}"><span>${emotion}</span></label>`).join('')}
        </div>
        <label><span>画像・作品（任意）</span><input id="image" type="file" accept="image/*"></label>
        <p class="field-note">画像は端末内で縮小してから自分のDriveへ保存し、先生へ直接共有します。</p>
        <button class="primary wide" type="submit">Driveに保存して先生へ共有</button>
      </form>
    </section>
    <section>
      <h2>これまでのふりかえり</h2>
      <div class="journal-list">${(portfolio.journals || []).length ? [...portfolio.journals].reverse().map((journal) => `
        <article class="journal-card"><div class="journal-meta"><span>${escapeHtml(formatDate(journal.createdAt))}</span><span>${escapeHtml(journal.emotion || '')}</span></div>${journal.theme ? `<h3>${escapeHtml(journal.theme)}</h3>` : ''}<div class="journal-body">${escapeHtml(journal.content)}</div>${journal.imageFileId ? '<p class="muted small">📎 画像を添付済み</p>' : ''}</article>`).join('') : '<div class="empty">まだふりかえりはありません。</div>'}</div>
    </section>`);
  document.getElementById('student-classes').addEventListener('click', renderStudentHome);
  document.getElementById('journal-form').addEventListener('submit', saveJournal);
  document.getElementById('reshare')?.addEventListener('click', async () => {
    await withError(async () => { await state.drive.shareWithUser(item.file.id, portfolio.class.teacherEmail, 'reader'); toast('先生へ共有しました。'); openPortfolio(item); }, (message) => openPortfolio(item, message));
  });
}

async function saveJournal(event) {
  event.preventDefault();
  const theme = document.getElementById('theme').value;
  const content = document.getElementById('content').value.trim();
  const emotion = document.querySelector('[name="emotion"]:checked')?.value || '';
  const inputFile = document.getElementById('image').files[0] || null;
  if (!content) return;
  const journalId = crypto.randomUUID();
  setBusy('ふりかえりを保存しています…');
  await withError(async () => {
    let imageFileId = null;
    let imageName = null;
    if (inputFile) {
      const image = await resizeImage(inputFile);
      const created = await state.drive.createJournalImage({
        classId: state.portfolio.class.id,
        journalId,
        studentName: state.portfolio.student.name,
        file: image
      });
      await state.drive.shareWithUser(created.id, state.portfolio.class.teacherEmail, 'reader');
      imageFileId = created.id;
      imageName = inputFile.name;
    }
    const updated = appendJournal(state.portfolio, { id: journalId, theme, content, emotion, imageFileId, imageName });
    await state.drive.updateJson(state.portfolioFile.id, updated);
    state.portfolio = updated;
    toast('Driveに保存しました。');
    openPortfolio({ file: state.portfolioFile, record: updated });
  }, (message) => openPortfolio({ file: state.portfolioFile, record: state.portfolio }, message));
}

async function resizeImage(file) {
  if (!file.type.startsWith('image/')) throw new Error('画像ファイルを選んでください。');
  const bitmap = await createImageBitmap(file);
  const maxSide = 1600;
  const scale = Math.min(1, maxSide / Math.max(bitmap.width, bitmap.height));
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.round(bitmap.width * scale));
  canvas.height = Math.max(1, Math.round(bitmap.height * scale));
  canvas.getContext('2d').drawImage(bitmap, 0, 0, canvas.width, canvas.height);
  bitmap.close();
  const blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/jpeg', .82));
  if (!blob) throw new Error('画像を変換できませんでした。');
  return new File([blob], file.name.replace(/\.[^.]+$/, '') + '.jpg', { type: 'image/jpeg' });
}

function renderQr(value) {
  const target = document.getElementById('qr');
  if (!target || typeof window.qrcode !== 'function') {
    if (target) target.textContent = 'QRコードを生成できませんでした。';
    return;
  }
  const qr = window.qrcode(0, 'M');
  qr.addData(value);
  qr.make();
  target.innerHTML = qr.createSvgTag({ cellSize: 5, margin: 2, scalable: true });
}

async function copy(value, message) {
  try { await navigator.clipboard.writeText(value); toast(message); }
  catch (error) { toast('コピーできませんでした。URLを選択してコピーしてください。'); }
}

function formatDate(value) {
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? '' : new Intl.DateTimeFormat('ja-JP', { dateStyle: 'medium', timeStyle: 'short' }).format(date);
}

renderLogin();
