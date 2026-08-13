import {
  analyzeClass,
  appendJournal,
  computeClassId,
  createChannel,
  createClassRecord,
  createPortfolio,
  currentTheme,
  decodeInvite,
  exportCsv,
  inviteUrl,
  isEmail,
  mergePortfoliosIntoMembers,
  normalizeClassCode,
  normalizeEmail,
  randomClassCode,
  setFeedback,
  studentKey,
  syncChannel,
  updatePastComment
} from './drive-core.js';
import { DriveApiError, DriveClient } from './drive-api.js';

const app = document.getElementById('app');
const toastElement = document.getElementById('toast');
const config = window.APP_CONFIG || {};
const state = {
  user: null,
  drive: null,
  invite: null,
  tokenClient: null,
  teacher: null,
  portfolio: null,
  portfolioFile: null,
  channel: null,
  channelFile: null,
  studentMonth: new Date(),
  teacherFilter: 'all',
  teacherSearch: '',
  teacherClasses: [],
  restoreLastTeacherClass: false,
  feedbackDraftKey: '',
  feedbackDraftHighlights: []
};

const escapeHtml = (value) => String(value ?? '').replace(/[&<>'"]/g, (char) => ({
  '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;'
})[char]);
const todayKey = (date = new Date()) => [date.getFullYear(), String(date.getMonth() + 1).padStart(2, '0'), String(date.getDate()).padStart(2, '0')].join('-');
const ruby = (word, reading) => `<ruby>${escapeHtml(word)}<rt>${escapeHtml(reading)}</rt></ruby>`;

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

function setBusy(message = 'データを読み込んでいます…') {
  app.innerHTML = `<section class="center-screen"><div class="loader" aria-hidden="true"></div><p>${escapeHtml(message)}</p></section>`;
}

function friendlyError(error) {
  console.error(error);
  if (error instanceof DriveApiError) return error.message;
  return error?.message || '操作を完了できませんでした。もう一度お試しください。';
}

function errorNotice(message) {
  const text = typeof message === 'string' ? message.trim() : '';
  return text ? `<div class="error" role="alert">${escapeHtml(text)}</div>` : '';
}

async function withError(action, fallback) {
  try { return await action(); }
  catch (error) {
    if (error instanceof DriveApiError && error.status === 401) return renderLogin(error.message);
    fallback(friendlyError(error));
    return null;
  }
}

function shell(content) {
  return `<header class="topbar">
    <div class="brand"><img class="brand-icon" src="./icon-192.png" alt="" aria-hidden="true"><span>ふりかえりジャーナル</span></div>
    <div class="account"><strong>${escapeHtml(state.user?.name || '')}</strong><span>${escapeHtml(state.user?.email || '')}</span></div>
  </header><div class="page">${content}</div>`;
}

function renderLogin(error = '') {
  app.innerHTML = `<section class="center-screen"><div class="login-card">
    <img class="app-logo" src="./icon-192.png" alt="" aria-hidden="true">
    <h1>毎日のふりかえりを<br>学びの成長へ</h1>
    <p>自分の言葉で学びを残し、先生からのおへんじを受け取れます。</p>
    ${errorNotice(error)}
    <button id="login" class="primary wide" type="button">Googleアカウントで続ける</button>
    <p class="muted small">必要な記録だけを安全に保存します。ログイン情報は端末へ保存しません。</p>
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
  if (!config.googleClientId) return renderLogin('ログインの準備が完了していません。アプリ管理者に連絡してください。');
  setBusy('Googleログインをひらいています…');
  try {
    await loadGoogleIdentity();
    state.tokenClient ||= google.accounts.oauth2.initTokenClient({
      client_id: config.googleClientId,
      scope: 'openid email profile https://www.googleapis.com/auth/drive.file',
      include_granted_scopes: true,
      callback: handleToken
    });
    state.tokenClient.requestAccessToken({ prompt: '' });
  } catch (error) { renderLogin(error.message); }
}

async function handleToken(response) {
  if (!response?.access_token) return renderLogin(response?.error_description || 'Googleログインがキャンセルされました。');
  setBusy('アカウントを確認しています…');
  try {
    const userResponse = await fetch('https://openidconnect.googleapis.com/v1/userinfo', { headers: { Authorization: `Bearer ${response.access_token}` } });
    if (!userResponse.ok) throw new Error('Googleアカウント情報を確認できませんでした。');
    const profile = await userResponse.json();
    state.user = { email: normalizeEmail(profile.email), name: profile.name || profile.email };
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
  app.innerHTML = shell(`<section class="role-home"><div class="role-home-intro"><h1>どの使い方をしますか？</h1><p>あなたに合った入口を選んでください。</p></div>
    ${errorNotice(error)}
    <div class="grid role-choice-grid">
      <button class="item-card" id="teacher-home" type="button"><h2>先生として使う</h2><p class="muted">クラス作成、招待、返却、分析、名簿とテーマを管理します。</p></button>
      <button class="item-card" id="student-home" type="button"><h2>児童として使う</h2><p class="muted">参加したクラスで書き、おへんじを受け取ります。</p></button>
    </div></section>`);
  document.getElementById('teacher-home').addEventListener('click', () => { state.restoreLastTeacherClass = true; renderTeacherHome(); });
  document.getElementById('student-home').addEventListener('click', () => renderStudentHome());
}

async function renderTeacherHome(error = '') {
  setBusy('先生のクラスを探しています…');
  const files = await withError(() => state.drive.listClasses(), (message) => renderHome(message));
  if (!files) return;
  const classes = (await Promise.all(files.map(async (file) => {
    try { return { file, record: await state.drive.getJson(file.id) }; } catch (loadError) { return null; }
  }))).filter(Boolean);
  state.teacherClasses = classes;
  if (state.restoreLastTeacherClass && classes.length) {
    state.restoreLastTeacherClass = false;
    let last = '';
    try { last = localStorage.getItem('rj_last_teacher_class') || ''; } catch (storageError) {}
    const match = classes.find(({ file }) => file.id === last);
    if (match) return openTeacherClass(match);
  }
  state.restoreLastTeacherClass = false;
  app.innerHTML = shell(`<div class="page-heading"><div><span class="badge">先生</span><h1>クラス</h1></div><button id="back" class="quiet" type="button">使い方を変える</button></div>
    ${errorNotice(error)}
    <section class="panel"><h2>新しいクラスを作る</h2><form id="create-class">
      <label><span>クラス名</span><input id="class-name" maxlength="80" required placeholder="例：5年1組"></label>
      <button class="primary" type="submit">クラスを作成</button>
    </form></section>
    <section><h2>作成済みのクラス</h2>${classes.length ? `<div class="grid">${classes.map(({ file, record }) => `
      <button class="item-card class-item" data-file="${escapeHtml(file.id)}" type="button"><span class="badge">${escapeHtml(record.classCode)}</span><h3>${escapeHtml(record.className)}</h3><p>${(record.members || []).filter((member) => member.status === 'active').length}人</p><p class="muted small">更新: ${escapeHtml(formatDate(file.modifiedTime))}</p></button>`).join('')}</div>` : '<div class="empty">まだクラスはありません。</div>'}</section>`);
  document.getElementById('back').addEventListener('click', () => renderHome());
  document.getElementById('create-class').addEventListener('submit', createClass);
  document.querySelectorAll('.class-item').forEach((button) => button.addEventListener('click', () => openTeacherClass(classes.find(({ file }) => file.id === button.dataset.file))));
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
    await openTeacherClass({ file, record });
  }, (message) => renderTeacherHome(message));
}

async function loadJsonItems(files) {
  return (await Promise.all(files.map(async (file) => {
    try { return { file, record: await state.drive.getJson(file.id) }; } catch (error) { return null; }
  }))).filter(Boolean);
}

async function openTeacherClass(item, tab = 'journals', error = '') {
  setBusy('クラスのデータを読み込んでいます…');
  await withError(async () => {
    const portfolioFiles = await state.drive.listSharedPortfolios(item.record.classId);
    const portfolios = (await loadJsonItems(portfolioFiles)).filter((entry) => entry.record?.class?.id === item.record.classId);
    let record = mergePortfoliosIntoMembers(item.record, portfolios);
    const original = JSON.stringify(item.record);
    const channels = new Map();
    for (const member of record.members || []) {
      if (member.status !== 'active') continue;
      let channel;
      if (member.channelFileId) {
        try { channel = await state.drive.getJson(member.channelFileId); } catch (loadError) { member.channelFileId = ''; }
      }
      if (!channel && member.portfolioFileId) {
        channel = createChannel({ classRecord: record, member });
        const created = await state.drive.createChannel(channel, await studentKey(member.email));
        await state.drive.shareWithUser(created.id, member.email, 'reader');
        member.channelFileId = created.id;
      }
      if (channel) channels.set(normalizeEmail(member.email), channel);
    }
    if (JSON.stringify(record) !== original) await state.drive.updateJson(item.file.id, record);
    state.teacher = { file: item.file, record, portfolios, channels };
    try { localStorage.setItem('rj_last_teacher_class', item.file.id); } catch (storageError) {}
    renderTeacherClass(tab, error);
  }, (message) => renderTeacherHome(message));
}

function teacherTabs(active) {
  const pending = (state.teacher.record.members || []).filter((member) => member.status === 'pending').length;
  return `<nav class="tabs" aria-label="クラスメニュー">${[
    ['journals', 'ジャーナル管理'], ['vitals', '心のバイタル'], ['settings', `クラス設定${pending ? ` (${pending})` : ''}`]
  ].map(([id, label]) => `<button class="tab ${active === id ? 'active' : ''}" data-tab="${id}" type="button">${label}</button>`).join('')}</nav>`;
}

function bindTeacherTabs() {
  document.querySelectorAll('[data-tab]').forEach((button) => button.addEventListener('click', () => renderTeacherClass(button.dataset.tab)));
  document.getElementById('classes')?.addEventListener('click', () => renderTeacherHome());
  document.getElementById('class-switch')?.addEventListener('change', (event) => {
    if (event.target.value === 'new') return renderTeacherHome();
    const next = state.teacherClasses.find(({ file }) => file.id === event.target.value);
    if (next) openTeacherClass(next);
  });
}

function renderTeacherClass(tab = 'journals', error = '') {
  const ctx = state.teacher;
  app.innerHTML = shell(`<div class="page-heading teacher-class-heading"><div><span class="badge">${escapeHtml(ctx.record.classCode)}</span><h1>${escapeHtml(ctx.record.className)}</h1></div><div class="class-switcher"><label for="class-switch">表示するクラス</label><select id="class-switch">${state.teacherClasses.map(({ file, record }) => `<option value="${escapeHtml(file.id)}" ${file.id === ctx.file.id ? 'selected' : ''}>${escapeHtml(record.className)}</option>`).join('')}<option value="new">＋ クラス一覧・新規作成</option></select><button id="classes" class="quiet" type="button">一覧</button></div></div>
    ${errorNotice(error)}${teacherTabs(tab)}<div id="teacher-content"></div>`);
  bindTeacherTabs();
  if (tab === 'journals') renderJournalManagement();
  else if (tab === 'vitals') renderVitals();
  else renderClassSettings();
}

function activePortfolioItems() {
  const active = new Set((state.teacher.record.members || []).filter((member) => member.status === 'active').map((member) => normalizeEmail(member.email)));
  return state.teacher.portfolios.filter(({ record }) => active.has(normalizeEmail(record.student.email)));
}

function renderJournalManagement() {
  const ctx = state.teacher;
  const stats = analyzeClass(activePortfolioItems(), ctx.channels);
  const content = document.getElementById('teacher-content');
  const rate = stats.totalStudents ? Math.round(stats.submittedToday / stats.totalStudents * 100) : 0;
  content.innerHTML = `<section class="metrics teacher-overview">
    <div class="metric submission-donut-card"><div class="submission-donut" style="--rate:${rate * 3.6}deg"><span><strong>${rate}%</strong><small>${stats.submittedToday}/${stats.totalStudents}人</small></span></div><span>今日の提出率</span></div>
    <div class="metric"><strong>${stats.all.length}</strong><span>すべての記録</span></div>
    <div class="metric"><strong>${stats.returned}</strong><span>返却済み</span></div>
  </section>
  <section class="panel"><h2>テーマ設定</h2><form id="theme-form">
    <div class="form-grid"><label><span>適用日</span><input id="theme-date" type="date" value="${escapeHtml(ctx.record.settings.todayTheme?.date || todayKey())}"></label><label><span>今日のテーマ</span><input id="today-theme" maxlength="200" value="${escapeHtml(ctx.record.settings.todayTheme?.text || '')}" placeholder="今日の学びをふり返ろう"></label></div>
    <details><summary>曜日ごとのテーマ</summary><div class="form-grid">${['月','火','水','木','金'].map((day, index) => `<label><span>${day}曜日</span><input class="weekly-theme" data-day="${index + 1}" maxlength="200" value="${escapeHtml(ctx.record.settings.weeklyThemes?.[index + 1] || '')}"></label>`).join('')}</div></details>
    <button class="primary" type="submit">テーマを保存して児童へ配信</button>
  </form></section>
  <section><div class="page-heading"><h2>提出一覧</h2><input id="journal-search" class="compact-input" value="${escapeHtml(state.teacherSearch)}" placeholder="氏名・本文・テーマを検索"></div>
    <div class="filter-row" role="group" aria-label="返却状況で絞り込む">${[['all','すべて'],['unreturned','未返却'],['returned','返却済み']].map(([value, label]) => `<button class="filter-chip ${state.teacherFilter === value ? 'active' : ''}" data-filter="${value}" type="button">${label}</button>`).join('')}</div>
    <div class="batch-bar"><div><strong>まとめて支援</strong><span class="muted small">未返却の記録が対象です。AIの内容は返却前に確認してください。</span></div><div class="button-row compact"><button id="batch-ai-simple" class="secondary" type="button">AI下書き</button><button id="batch-ai-detail" class="secondary" type="button">AIで詳しく</button><button id="batch-return" class="primary" type="button">一括返却</button></div></div>
    <div id="journal-list" class="journal-list">${teacherJournalCards(filterTeacherJournals(stats.all))}</div>
  </section>`;
  document.getElementById('theme-form').addEventListener('submit', saveThemes);
  document.getElementById('journal-search').addEventListener('input', (event) => {
    state.teacherSearch = event.target.value;
    document.getElementById('journal-list').innerHTML = teacherJournalCards(filterTeacherJournals(stats.all));
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
  return journals.filter((journal) => {
    const returned = Boolean(state.teacher.channels.get(normalizeEmail(journal.student.email))?.feedback?.[journal.id]?.returned);
    const matchesStatus = state.teacherFilter === 'all' || (state.teacherFilter === 'returned' ? returned : !returned);
    const matchesQuery = !query || `${journal.student.name} ${journal.theme} ${journal.content}`.toLowerCase().includes(query);
    return matchesStatus && matchesQuery;
  });
}

function teacherJournalCards(journals) {
  return journals.length ? journals.map((journal) => {
    const feedback = state.teacher.channels.get(normalizeEmail(journal.student.email))?.feedback?.[journal.id];
    return `<article class="journal-card teacher-journal-card"><button class="journal-main journal-open" data-email="${escapeHtml(journal.student.email)}" data-journal="${escapeHtml(journal.id)}" type="button">
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
    await state.drive.updateJson(member.channelFileId, channel);
    state.teacher.channels.set(normalized, channel);
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
      const member = state.teacher.record.members.find((item) => normalizeEmail(item.email) === email);
      if (member?.channelFileId) await state.drive.updateJson(member.channelFileId, channel);
      state.teacher.channels.set(email, channel);
    }
    toast(`${targets.length}件を返却しました。`);
    renderTeacherClass('journals');
  }, (message) => renderTeacherClass('journals', message));
}

async function geminiText(prompt) {
  const apiKey = state.teacher.record.settings.geminiApiKey;
  if (!apiKey) throw new Error('クラス設定でGemini APIキーを保存してください。');
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
      const member = state.teacher.record.members.find((item) => normalizeEmail(item.email) === email);
      if (member?.channelFileId) await state.drive.updateJson(member.channelFileId, channel);
      state.teacher.channels.set(email, channel);
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
    await state.drive.updateJson(ctx.file.id, ctx.record);
    for (const member of ctx.record.members.filter((item) => item.status === 'active' && item.channelFileId)) {
      const email = normalizeEmail(member.email);
      const channel = syncChannel(ctx.channels.get(email), ctx.record, member);
      await state.drive.updateJson(member.channelFileId, channel);
      ctx.channels.set(email, channel);
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
  document.getElementById('journal-back').addEventListener('click', () => renderTeacherClass('journals'));
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
    await state.drive.updateJson(member.channelFileId, channel);
    state.teacher.channels.set(email, channel);
    state.feedbackDraftKey = '';
    state.feedbackDraftHighlights = [];
    toast('児童へ返却しました。');
    renderTeacherClass('journals');
  }, (message) => renderFeedbackEditor(portfolioItem, journal, message));
}

async function generateAiDraft(journal, portfolioItem) {
  const apiKey = state.teacher.record.settings.geminiApiKey;
  if (!apiKey) return renderFeedbackEditor(portfolioItem, journal, 'クラス設定でGemini APIキーを保存してください。');
  const button = document.getElementById('ai-draft');
  button.disabled = true;
  button.textContent = '下書きを作成中…';
  try {
    const model = state.teacher.record.settings.geminiModel || 'gemini-3.1-flash-lite';
    const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-goog-api-key': apiKey },
      body: JSON.stringify({ contents: [{ parts: [{ text: `あなたは小学校の担任です。次の児童のふりかえりへ、具体的なよさを認め、次の学びにつながる温かいコメントを100字以内で日本語で書いてください。コメント本文だけを返してください。\n\nテーマ: ${journal.theme}\n本文: ${journal.content}` }] }] })
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error?.message || 'Gemini APIでエラーが発生しました。');
    document.getElementById('feedback-comment').value = result.candidates?.[0]?.content?.parts?.map((part) => part.text || '').join('')?.trim() || '';
    toast('AIの下書きを作成しました。内容を確認してください。');
  } catch (error) { renderFeedbackEditor(portfolioItem, journal, friendlyError(error)); }
  finally { if (button.isConnected) { button.disabled = false; button.textContent = 'AIで下書き'; } }
}

function renderVitals() {
  const stats = analyzeClass(activePortfolioItems(), state.teacher.channels);
  const days = Array.from({ length: 14 }, (_, index) => { const date = new Date(); date.setDate(date.getDate() - (13 - index)); return date; });
  const content = document.getElementById('teacher-content');
  content.innerHTML = `<section class="panel"><h2>気にかけたい児童</h2>${stats.alerts.length ? stats.alerts.map((alert) => `<div class="notice"><strong>${escapeHtml(alert.name)}</strong> — ${escapeHtml(alert.reason)}</div>`).join('') : '<div class="empty">現在の記録から強い変化は検出されていません。対面での様子も必ず合わせて見てください。</div>'}</section>
    <section class="panel"><h2>クラスの心の波（14日間）</h2><div class="heatmap"><div class="heatmap-row heatmap-head"><span>児童</span>${days.map((date) => `<span>${date.getMonth() + 1}/${date.getDate()}</span>`).join('')}</div>${activePortfolioItems().map(({ record }) => `<div class="heatmap-row"><strong>${escapeHtml(record.student.name)}</strong>${days.map((date) => { const journal = (record.journals || []).find((item) => todayKey(new Date(item.createdAt)) === todayKey(date)); return `<span title="${escapeHtml(journal?.content || '未提出')}">${escapeHtml(journal?.emotion || '·')}</span>`; }).join('')}</div>`).join('')}</div></section>`;
}

function publicEntryUrl() {
  return config.publicEntryUrl || new URL('./', location.href).href;
}

function renderClassSettings() {
  const ctx = state.teacher;
  const settings = ctx.record.settings;
  const link = inviteUrl(publicEntryUrl(), {
    classCode: ctx.record.classCode,
    className: ctx.record.className,
    teacherEmail: ctx.record.teacher.email,
    teacherName: ctx.record.teacher.name,
    approvalRequired: settings.approvalRequired,
    acceptingMembers: true
  });
  const content = document.getElementById('teacher-content');
  content.innerHTML = `<section class="panel invite-layout"><div><span class="eyebrow">INVITE STUDENTS</span><h2>児童を招待する</h2><p class="muted">教室ではQRコード、遠隔では専用URLを配ると簡単です。</p><div class="class-code">${escapeHtml(ctx.record.classCode)}</div><div class="button-row"><button id="copy-code" class="secondary" type="button">クラスコードをコピー</button><button id="student-preview" class="quiet" type="button">児童画面を確認</button></div><div class="url-box">${escapeHtml(link)}</div><div class="button-row" style="margin-top:12px"><button id="copy-url" class="primary" type="button">専用URLをコピー</button><button id="copy-text" class="secondary" type="button">招待文をコピー</button></div></div><div class="qr-actions"><div id="qr" class="qr-box" aria-label="児童招待用QRコード"></div><button id="download-qr" class="secondary wide" type="button">QRを画像で保存</button></div></section>
  <section class="panel"><h2>参加設定</h2><form id="admission-form"><label class="check-label"><input id="approval" type="checkbox" ${settings.approvalRequired ? 'checked' : ''}><span>参加には先生の承認が必要</span></label><p class="muted small">承認なしにすると、児童の参加後に先生がクラス画面を開いた時点で自動承認されます。</p><button class="primary" type="submit">設定を保存</button></form></section>
  <section class="panel"><div class="section-heading"><div><span class="eyebrow">ROSTER</span><h2>参加申請・名簿</h2></div><span class="badge">${(ctx.record.members || []).filter((member) => member.status === 'active').length}人 参加中</span></div><div id="member-list">${memberRows()}</div><form id="invite-member" class="inline-form"><input id="member-name" required placeholder="児童名"><input id="member-email" type="email" required placeholder="児童のGoogleメール"><button class="secondary" type="submit">1人追加</button></form><details><summary>名簿をまとめて追加</summary><form id="bulk-members"><label><span>1行に「名前, メールアドレス」</span><textarea id="bulk-member-text" placeholder="山田 花子, hanako@example.ed.jp\n鈴木 太郎, taro@example.ed.jp"></textarea></label><button class="secondary" type="submit">まとめて名簿へ追加</button></form></details></section>
  <section class="panel"><h2>Gemini AI支援（任意）</h2><form id="gemini-form"><label><span>APIキー</span><input id="gemini-key" type="password" value="${escapeHtml(settings.geminiApiKey || '')}" autocomplete="off"></label><label><span>モデル</span><input id="gemini-model" value="${escapeHtml(settings.geminiModel || 'gemini-3.1-flash-lite')}"></label><p class="muted small">キーは先生所有の非共有クラス設定ファイルに保存され、児童へは共有されません。AIコメントは必ず先生が確認してから返却します。</p><button class="primary" type="submit">AI設定を保存</button></form></section>
  <section class="panel"><h2>データ出力</h2><div class="button-row"><button id="csv" class="secondary" type="button">CSVをダウンロード</button><button id="print" class="secondary" type="button">印刷・PDF</button></div></section>`;
  document.getElementById('copy-url').addEventListener('click', () => copy(link, '専用URLをコピーしました。'));
  document.getElementById('copy-code').addEventListener('click', () => copy(ctx.record.classCode, 'クラスコードをコピーしました。'));
  document.getElementById('copy-text').addEventListener('click', () => copy(`「${ctx.record.className}」のふりかえりジャーナルに参加してください。\nクラスコード: ${ctx.record.classCode}\n${link}`, '招待文をコピーしました。'));
  document.getElementById('student-preview').addEventListener('click', () => window.open(link, '_blank', 'noopener'));
  document.getElementById('download-qr').addEventListener('click', () => downloadQrCode(ctx.record.className));
  document.getElementById('admission-form').addEventListener('submit', saveAdmissionSettings);
  document.getElementById('invite-member').addEventListener('submit', addInvitedMember);
  document.getElementById('bulk-members').addEventListener('submit', addBulkMembers);
  document.getElementById('gemini-form').addEventListener('submit', saveGeminiSettings);
  document.getElementById('csv').addEventListener('click', downloadCsv);
  document.getElementById('print').addEventListener('click', printClass);
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
    } else if (channel && member.channelFileId) {
      channel = syncChannel(channel, ctx.record, member);
      await state.drive.updateJson(member.channelFileId, channel);
    }
    if (channel) ctx.channels.set(normalizeEmail(email), channel);
    await state.drive.updateJson(ctx.file.id, ctx.record);
    toast(status === 'active' ? '児童を承認しました。' : '参加状態を更新しました。');
    renderTeacherClass('settings');
  }, (message) => renderTeacherClass('settings', message));
}

async function saveGeminiSettings(event) {
  event.preventDefault();
  state.teacher.record.settings.geminiApiKey = document.getElementById('gemini-key').value.trim();
  state.teacher.record.settings.geminiModel = document.getElementById('gemini-model').value.trim() || 'gemini-3.1-flash-lite';
  await saveTeacherRecord('AI設定を保存しました。');
}

async function saveTeacherRecord(message) {
  const ctx = state.teacher;
  ctx.record.updatedAt = new Date().toISOString();
  setBusy('設定を保存しています…');
  await withError(async () => { await state.drive.updateJson(ctx.file.id, ctx.record); toast(message); renderTeacherClass('settings'); }, (error) => renderTeacherClass('settings', error));
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
  setBusy('参加済みのクラスを探しています…');
  const files = await withError(() => state.drive.listOwnPortfolios(), (message) => renderHome(message));
  if (!files) return;
  const portfolios = await loadJsonItems(files);
  app.innerHTML = shell(`<div class="page-heading"><div><span class="badge">児童</span><h1>参加しているクラス</h1></div><button id="back" class="quiet" type="button">使い方を変える</button></div>
    ${errorNotice(error)}${portfolios.length ? `<div class="grid">${portfolios.map(({ file, record }) => `<button class="item-card student-portfolio" data-file="${escapeHtml(file.id)}" type="button"><span class="badge">${escapeHtml(record.class?.code || '')}</span><h3>${escapeHtml(record.class?.name || 'クラス')}</h3><p>${record.journals?.length || 0}件のふりかえり</p></button>`).join('')}</div>` : '<div class="empty">参加済みのクラスはありません。先生のQRコードまたは専用URLから参加してください。</div>'}`);
  document.getElementById('back').addEventListener('click', () => renderHome());
  document.querySelectorAll('.student-portfolio').forEach((button) => button.addEventListener('click', () => openPortfolio(portfolios.find(({ file }) => file.id === button.dataset.file))));
}

async function renderJoin(error = '') {
  const invite = state.invite;
  if (!invite.acceptingMembers) return renderHome('この招待では新しい参加を受け付けていません。先生から新しいURLを受け取ってください。');
  setBusy('参加状況を確認しています…');
  const files = await withError(() => state.drive.listOwnPortfolios(invite.classId), (message) => renderHome(message));
  if (!files) return;
  if (files.length) {
    const record = await withError(() => state.drive.getJson(files[0].id), (message) => renderHome(message));
    if (record) return openPortfolio({ file: files[0], record });
  }
  app.innerHTML = shell(`<div class="page-heading"><div><span class="badge">招待</span><h1>${escapeHtml(invite.className)}</h1></div><button id="cancel" class="quiet" type="button">戻る</button></div>
    ${errorNotice(error)}<section class="panel"><p>先生: <strong>${escapeHtml(invite.teacherName || invite.teacherEmail)}</strong></p><p>クラスコード: <strong>${escapeHtml(invite.classCode)}</strong></p><form id="join-class"><label><span>先生に表示する名前</span><input id="student-name" maxlength="80" required value="${escapeHtml(state.user.name || '')}"></label><button class="primary wide" type="submit">このクラスに参加する</button></form><p class="muted small">参加すると、あなたのふりかえりを先生が確認できるようになります。</p></section>`);
  document.getElementById('cancel').addEventListener('click', () => { history.replaceState(null, '', location.pathname); state.invite = null; renderStudentHome(); });
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
    try { await state.drive.shareWithUser(file.id, state.invite.teacherEmail, 'reader'); }
    catch (error) { return openPortfolio({ file, record }, 'クラスへの参加は完了しましたが、先生へ記録を届ける設定が学校側で拒否されました。「先生へもう一度届ける」を押してください。'); }
    history.replaceState(null, '', location.pathname);
    await openPortfolio({ file, record });
  }, (message) => renderJoin(message));
}

async function findStudentChannel(classId) {
  const files = await state.drive.listSharedChannels(classId);
  for (const file of files) {
    try {
      const record = await state.drive.getJson(file.id);
      if (record.class?.id === classId && normalizeEmail(record.student?.email) === state.user.email) return { file, record };
    } catch (error) {}
  }
  return null;
}

async function openPortfolio(item, error = '') {
  setBusy('ふりかえりを開いています…');
  await withError(async () => {
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
  let draft = {};
  try { draft = JSON.parse(localStorage.getItem(draftKey()) || '{}'); } catch (loadError) {}
  const unread = unreadFeedbackJournals();
  app.innerHTML = shell(`<div class="page-heading"><div><span class="badge">${escapeHtml(portfolio.class.code)}</span><h1>${escapeHtml(portfolio.class.name)}</h1></div><button id="student-classes" class="quiet" type="button">クラス一覧へ</button></div>
    ${errorNotice(error)}${typeof error === 'string' && error ? '<button id="reshare" class="secondary" type="button">先生へもう一度届ける</button>' : ''}
    ${unread.length ? `<button id="unread-feedback" class="feedback-alert" type="button"><span>💌</span><strong>先生から新しいおへんじが ${unread.length}件 届いています</strong><span>見る →</span></button>` : ''}
    ${rejected ? '<div class="error">このクラスへの参加は承認されませんでした。先生へ確認してください。</div>' : pending ? '<div class="notice">参加申請を先生へ送りました。先生が承認すると、ふりかえりを書けるようになります。</div>' : studentWorkspace(theme, draft)}
    <section class="student-history"><div class="section-heading"><div><span class="eyebrow">MY JOURNAL</span><h2>これまでのふりかえり</h2></div><span class="muted">${(portfolio.journals || []).length}件</span></div><div class="journal-list">${studentJournalCards()}</div></section>
    ${pastSelfPanel()}`);
  document.getElementById('student-classes').addEventListener('click', () => renderStudentHome());
  document.getElementById('reshare')?.addEventListener('click', async () => {
    await withError(async () => { await state.drive.shareWithUser(state.portfolioFile.id, portfolio.class.teacherEmail, 'reader'); toast('先生へ共有しました。'); renderPortfolio(); }, (message) => renderPortfolio(message));
  });
  if (!pending && !rejected) {
    const form = document.getElementById('journal-form');
    form.addEventListener('submit', saveJournal);
    ['theme','content'].forEach((id) => document.getElementById(id).addEventListener('input', saveDraft));
    document.querySelectorAll('[data-insert]').forEach((button) => button.addEventListener('click', () => insertText(button.dataset.insert)));
    document.querySelectorAll('[data-template]').forEach((button) => button.addEventListener('click', () => applyTemplate(button.dataset.template)));
    document.getElementById('random-starter')?.addEventListener('click', insertRandomStarter);
    document.querySelectorAll('[name="emotion"]').forEach((input) => input.addEventListener('change', saveDraft));
  }
  document.querySelectorAll('[data-student-journal]').forEach((button) => button.addEventListener('click', () => openStudentJournal(button.dataset.studentJournal)));
  document.getElementById('calendar-prev')?.addEventListener('click', () => changeStudentMonth(-1));
  document.getElementById('calendar-next')?.addEventListener('click', () => changeStudentMonth(1));
  document.getElementById('unread-feedback')?.addEventListener('click', () => openStudentJournal(unread[0]?.id));
  document.getElementById('past-comment-form')?.addEventListener('submit', savePastComment);
}

function studentWorkspace(theme, draft) {
  return `<section class="student-workspace">
    <aside class="student-calendar panel"><div class="calendar-heading"><button id="calendar-prev" class="icon-button" type="button" aria-label="前の月">‹</button><strong>${state.studentMonth.getFullYear()}年 ${state.studentMonth.getMonth() + 1}月</strong><button id="calendar-next" class="icon-button" type="button" aria-label="次の月">›</button></div>${studentCalendar()}</aside>
    <section class="notebook panel"><div class="notebook-top"><div><span class="eyebrow">TODAY'S JOURNAL</span><h2>${ruby('今日', 'きょう')}のふりかえり</h2></div><span class="notebook-date">${new Intl.DateTimeFormat('ja-JP', { month: 'long', day: 'numeric', weekday: 'short' }).format(new Date())}</span></div>
      <div class="theme-banner"><span>${ruby('今日', 'きょう')}のテーマ</span><strong>${escapeHtml(theme)}</strong></div><form id="journal-form">
      <label class="visually-hidden"><span>テーマ</span><input id="theme" maxlength="200" value="${escapeHtml(draft.theme || theme)}"></label>
      <label><span>${ruby('自分', 'じぶん')}のことばで${ruby('書', 'か')}こう</span><textarea id="content" class="lined-paper" maxlength="20000" required placeholder="できたこと、考えたこと、次にやってみたいことを書こう">${escapeHtml(draft.content || '')}</textarea></label>
      <span><strong>${ruby('今', 'いま')}の${ruby('気持', 'きも')}ち</strong> <span class="muted small">（えらばなくてもOK）</span></span><div class="emotion-row labeled-emotions" role="radiogroup" aria-label="いまの気持ち">${[['😊','うれしい'],['😠','くやしい'],['💡','なるほど'],['🤔','もやもや']].map(([emotion, label]) => `<label><input type="radio" name="emotion" value="${emotion}" ${draft.emotion === emotion ? 'checked' : ''}><span><b>${emotion}</b><small>${label}</small></span></label>`).join('')}</div>
      <details class="attachment"><summary>📎 ${ruby('作品', 'さくひん')}や${ruby('写真', 'しゃしん')}をつける</summary><label><span>画像・作品（任意）</span><input id="image" type="file" accept="image/*"></label><p class="field-note">見やすい大きさにして先生へ届けます。</p></details>
      <button class="primary wide submit-journal" type="submit">できた！ 先生にとどける</button>
    </form></section>
    <aside class="writing-tools panel"><h3>✏️ ${ruby('書', 'か')}き${ruby('方', 'かた')}サポート</h3><p class="muted small">ボタンをおすと、文の書き出しが入ります。</p>
      <button id="random-starter" class="secondary wide" type="button">🎲 おまかせ書き出し</button>
      <details open><summary>かたをえらぶ</summary><div class="template-grid">${JOURNAL_TEMPLATES.map(([label, value]) => `<button class="template-button" data-template="${escapeHtml(value)}" type="button">${escapeHtml(label)}</button>`).join('')}</div></details>
      ${HINT_GROUPS.map(([label, hints], index) => `<details ${index === 0 ? 'open' : ''}><summary>${escapeHtml(label)}のヒント</summary><div class="hint-list">${hints.map((hint) => `<button class="hint-button" data-insert="${escapeHtml(hint)}" type="button">${escapeHtml(hint)}</button>`).join('')}</div></details>`).join('')}
    </aside>
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
  return `<div class="calendar-week">${['日','月','火','水','木','金','土'].map((day) => `<strong>${day}</strong>`).join('')}</div><div class="calendar-grid">${cells.join('')}</div><p class="calendar-key"><span>● 書いた日</span><span>💬 おへんじ</span></p>`;
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
    return `<button class="journal-card student-journal-card ${unread ? 'unread' : ''}" data-student-journal="${escapeHtml(journal.id)}" type="button"><div class="journal-meta"><span>${escapeHtml(formatDate(journal.createdAt))}</span><span>${escapeHtml(journal.emotion || '')}</span></div><h3>${escapeHtml(journal.theme || '')}</h3><div class="journal-body clamp">${escapeHtml(journal.content)}</div><div class="journal-footer">${journal.imageFileId ? '<span>📎 作品つき</span>' : '<span></span>'}${feedback?.returned ? `<span class="badge success-badge">${unread ? 'NEW ' : ''}${escapeHtml(feedback.stamp || '💬')} おへんじ</span>` : ''}</div></button>`;
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
  return items.length ? `<div class="highlight-notes"><h3>文章についたおへんじ</h3>${items.map((item) => `<div><strong>${escapeHtml(item.stamp || '⭐')} 「${escapeHtml(item.text || '')}」</strong><p>${escapeHtml(item.comment || '')}</p></div>`).join('')}</div>` : '';
}

function openStudentJournal(journalId) {
  if (!journalId) return;
  const journal = (state.portfolio.journals || []).find((item) => item.id === journalId);
  if (!journal) return;
  const feedback = state.channel?.feedback?.[journal.id];
  if (feedback?.returned) markJournalRead(journal.id);
  const dialog = document.createElement('dialog');
  dialog.className = 'journal-dialog';
  dialog.innerHTML = `<article><form method="dialog"><button class="dialog-close" aria-label="閉じる">×</button></form><div class="journal-meta"><span>${escapeHtml(formatDate(journal.createdAt))}</span><span>${escapeHtml(journal.emotion || '')}</span></div><h2>${escapeHtml(journal.theme || 'ふりかえり')}</h2><div class="journal-body lined-reading">${highlightedJournalHtml(journal, feedback)}</div>${journal.imageFileId ? `<div class="image-slot" data-file="${escapeHtml(journal.imageFileId)}"><p>作品を読み込んでいます…</p></div>` : ''}${feedback?.returned ? `${studentHighlightNotes(feedback)}<div class="feedback-box large"><strong>${escapeHtml(feedback.stamp || '💬')} 先生からのおへんじ</strong><p>${escapeHtml(feedback.comment)}</p></div>` : '<div class="muted dialog-note">先生からのおへんじを待っています。</div>'}${journal.pastComment ? `<div class="notice"><strong>今の自分から:</strong> ${escapeHtml(journal.pastComment)}</div>` : ''}</article>`;
  document.body.appendChild(dialog);
  dialog.addEventListener('close', () => { dialog.remove(); renderPortfolio(); });
  dialog.addEventListener('click', (event) => { if (event.target === dialog) dialog.close(); });
  dialog.showModal();
  loadImages();
}

function pastSelfPanel() {
  const cutoff = Date.now() - 30 * 86400000;
  const old = (state.portfolio.journals || []).filter((journal) => new Date(journal.createdAt).getTime() < cutoff);
  if (!old.length) return '';
  const journal = old[Math.floor(Math.random() * old.length)];
  return `<section class="panel"><h2>過去の自分と対話</h2><blockquote>${escapeHtml(journal.content)}</blockquote><form id="past-comment-form" data-journal="${escapeHtml(journal.id)}"><label><span>今の自分からひとこと</span><textarea id="past-comment" maxlength="4000">${escapeHtml(journal.pastComment || '')}</textarea></label><button class="secondary" type="submit">メッセージを保存</button></form></section>`;
}

function saveDraft() {
  try { localStorage.setItem(draftKey(), JSON.stringify({ theme: document.getElementById('theme').value, content: document.getElementById('content').value, emotion: document.querySelector('[name="emotion"]:checked')?.value || '' })); } catch (error) {}
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
  setBusy('ふりかえりを保存しています…');
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
    await state.drive.updateJson(state.portfolioFile.id, updated);
    state.portfolio = updated;
    try { localStorage.removeItem(draftKey()); } catch (error) {}
    toast('ふりかえりを提出しました。');
    renderPortfolio();
    celebrateSubmission();
  }, (message) => renderPortfolio(message));
}

async function savePastComment(event) {
  event.preventDefault();
  const updated = updatePastComment(state.portfolio, event.currentTarget.dataset.journal, document.getElementById('past-comment').value);
  setBusy('メッセージを保存しています…');
  await withError(async () => { await state.drive.updateJson(state.portfolioFile.id, updated); state.portfolio = updated; toast('メッセージを保存しました。'); renderPortfolio(); }, (message) => renderPortfolio(message));
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
    catch (error) { slot.innerHTML = '<p class="error small">画像を表示できませんでした。</p>'; }
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

initPwaControls();
renderLogin();
