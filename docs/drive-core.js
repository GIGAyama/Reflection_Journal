const CLASS_CODE_ALPHABET = '23456789ABCDEFGHJKLMNPQRSTUVWXYZ';
export const SCHEMA_VERSION = 2;

export function normalizeClassCode(value) {
  return String(value || '').toUpperCase().replace(/[^23456789ABCDEFGHJKLMNPQRSTUVWXYZ]/g, '').slice(0, 10);
}

export function normalizeEmail(value) {
  return String(value || '').trim().toLowerCase();
}

export function isEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normalizeEmail(value));
}

export function randomClassCode(length = 8, cryptoObject = globalThis.crypto) {
  const bytes = new Uint8Array(length);
  cryptoObject.getRandomValues(bytes);
  return Array.from(bytes, (byte) => CLASS_CODE_ALPHABET[byte % CLASS_CODE_ALPHABET.length]).join('');
}

export async function sha256(value, cryptoObject = globalThis.crypto) {
  const digest = await cryptoObject.subtle.digest('SHA-256', new TextEncoder().encode(String(value)));
  return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, '0')).join('');
}

export function computeClassId(teacherEmail, classCode, cryptoObject = globalThis.crypto) {
  return sha256(`reflection-journal:v2:${normalizeEmail(teacherEmail)}:${normalizeClassCode(classCode)}`, cryptoObject);
}

export function studentKey(email) {
  return sha256(`reflection-journal:student:v2:${normalizeEmail(email)}`);
}

function bytesToBase64Url(bytes) {
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
}

function base64UrlToBytes(value) {
  const input = String(value);
  const padded = input.replace(/-/g, '+').replace(/_/g, '/') + '==='.slice((input.length + 3) % 4);
  return Uint8Array.from(atob(padded), (char) => char.charCodeAt(0));
}

export function encodeInvite(invite) {
  const compact = {
    v: SCHEMA_VERSION,
    c: normalizeClassCode(invite.classCode),
    n: String(invite.className || '').trim().slice(0, 80),
    e: normalizeEmail(invite.teacherEmail),
    t: String(invite.teacherName || '').trim().slice(0, 80),
    a: invite.approvalRequired !== false,
    o: invite.acceptingMembers !== false
  };
  return bytesToBase64Url(new TextEncoder().encode(JSON.stringify(compact)));
}

export async function decodeInvite(encoded) {
  let raw;
  try { raw = JSON.parse(new TextDecoder().decode(base64UrlToBytes(encoded))); }
  catch (error) { throw new Error('招待情報を読み取れませんでした。先生から新しいURLを受け取ってください。'); }
  const invite = {
    version: Number(raw.v),
    classCode: normalizeClassCode(raw.c),
    className: String(raw.n || '').trim().slice(0, 80),
    teacherEmail: normalizeEmail(raw.e),
    teacherName: String(raw.t || '').trim().slice(0, 80),
    approvalRequired: raw.a !== false,
    acceptingMembers: raw.o !== false
  };
  if (invite.version !== SCHEMA_VERSION || invite.classCode.length < 6 || !invite.className || !isEmail(invite.teacherEmail)) {
    throw new Error('招待情報が正しくありません。先生から新しいURLを受け取ってください。');
  }
  invite.classId = await computeClassId(invite.teacherEmail, invite.classCode);
  return invite;
}

export function defaultSettings() {
  return {
    approvalRequired: true,
    acceptingMembers: true,
    todayTheme: { date: '', text: '' },
    weeklyThemes: { 1: '', 2: '', 3: '', 4: '', 5: '' },
    geminiApiKey: '',
    geminiModel: 'gemini-3.1-flash-lite'
  };
}

export function createClassRecord({ classId, classCode, className, teacher, now = new Date().toISOString() }) {
  return {
    schemaVersion: SCHEMA_VERSION,
    kind: 'reflection-journal-class',
    classId,
    classCode: normalizeClassCode(classCode),
    className: String(className || '').trim(),
    teacher: { email: normalizeEmail(teacher.email), name: String(teacher.name || '').trim() },
    settings: defaultSettings(),
    members: [],
    createdAt: now,
    updatedAt: now
  };
}

export function createPortfolio({ invite, student, now = new Date().toISOString() }) {
  return {
    schemaVersion: SCHEMA_VERSION,
    kind: 'reflection-journal-portfolio',
    class: {
      id: invite.classId,
      code: normalizeClassCode(invite.classCode),
      name: String(invite.className || '').trim(),
      teacherEmail: normalizeEmail(invite.teacherEmail),
      teacherName: String(invite.teacherName || '').trim(),
      approvalRequired: invite.approvalRequired !== false,
      acceptingMembers: invite.acceptingMembers !== false
    },
    student: { email: normalizeEmail(student.email), name: String(student.name || '').trim() },
    journals: [],
    createdAt: now,
    updatedAt: now
  };
}

export function createChannel({ classRecord, member, status = 'active', now = new Date().toISOString() }) {
  return {
    schemaVersion: SCHEMA_VERSION,
    kind: 'reflection-journal-channel',
    class: { id: classRecord.classId, code: classRecord.classCode, name: classRecord.className },
    teacher: { ...classRecord.teacher },
    student: { email: normalizeEmail(member.email), name: String(member.name || '').trim() },
    status,
    themes: {
      todayTheme: { ...(classRecord.settings?.todayTheme || { date: '', text: '' }) },
      weeklyThemes: { ...(classRecord.settings?.weeklyThemes || {}) }
    },
    feedback: {},
    createdAt: now,
    updatedAt: now
  };
}

export function syncChannel(channel, classRecord, member, now = new Date().toISOString()) {
  return {
    ...channel,
    class: { id: classRecord.classId, code: classRecord.classCode, name: classRecord.className },
    teacher: { ...classRecord.teacher },
    student: { email: normalizeEmail(member.email), name: String(member.name || '').trim() },
    status: member.status,
    themes: {
      todayTheme: { ...(classRecord.settings?.todayTheme || { date: '', text: '' }) },
      weeklyThemes: { ...(classRecord.settings?.weeklyThemes || {}) }
    },
    updatedAt: now
  };
}

export function setFeedback(channel, journalId, feedback, now = new Date().toISOString()) {
  const id = String(journalId || '');
  if (!id) throw new Error('ふりかえりIDがありません。');
  return {
    ...channel,
    feedback: {
      ...(channel.feedback || {}),
      [id]: {
        comment: String(feedback.comment || '').trim().slice(0, 4000),
        stamp: String(feedback.stamp || '').slice(0, 8),
        highlights: Array.isArray(feedback.highlights) ? feedback.highlights.slice(0, 20) : [],
        returned: feedback.returned !== false,
        updatedAt: now
      }
    },
    updatedAt: now
  };
}

export function appendJournal(portfolio, journal, now = new Date().toISOString()) {
  if (!portfolio || portfolio.kind !== 'reflection-journal-portfolio') throw new Error('ポートフォリオ形式が正しくありません。');
  const item = {
    id: String(journal.id || ''),
    theme: String(journal.theme || '').trim().slice(0, 200),
    content: String(journal.content || '').trim().slice(0, 20000),
    emotion: String(journal.emotion || '').slice(0, 8),
    imageFileId: journal.imageFileId ? String(journal.imageFileId) : null,
    imageName: journal.imageName ? String(journal.imageName).slice(0, 200) : null,
    pastComment: '',
    createdAt: journal.createdAt || now,
    updatedAt: now
  };
  if (!item.id || !item.content) throw new Error('本文を入力してください。');
  return { ...portfolio, journals: [...(portfolio.journals || []), item], updatedAt: now };
}

export function updatePastComment(portfolio, journalId, comment, now = new Date().toISOString()) {
  return {
    ...portfolio,
    journals: (portfolio.journals || []).map((journal) => journal.id === journalId
      ? { ...journal, pastComment: String(comment || '').trim().slice(0, 4000), updatedAt: now }
      : journal),
    updatedAt: now
  };
}

export function currentTheme(themes, date = new Date()) {
  const localDate = [date.getFullYear(), String(date.getMonth() + 1).padStart(2, '0'), String(date.getDate()).padStart(2, '0')].join('-');
  if (themes?.todayTheme?.date === localDate && themes.todayTheme.text) return themes.todayTheme.text;
  return themes?.weeklyThemes?.[date.getDay()] || '今日の学びをふり返ろう';
}

export function mergePortfoliosIntoMembers(classRecord, portfolioItems, now = new Date().toISOString()) {
  const members = (classRecord.members || []).map((member) => ({ ...member }));
  for (const item of portfolioItems) {
    const email = normalizeEmail(item.record?.student?.email);
    if (!email) continue;
    const existing = members.find((member) => normalizeEmail(member.email) === email);
    if (existing) {
      existing.portfolioFileId = item.file.id;
      existing.name = existing.name || item.record.student.name;
      existing.joinedAt = existing.joinedAt || item.record.createdAt || now;
      if (existing.status === 'invited') existing.status = 'active';
    } else {
      members.push({
        email,
        name: item.record.student.name || email,
        role: 'student',
        status: classRecord.settings?.approvalRequired === false ? 'active' : 'pending',
        portfolioFileId: item.file.id,
        channelFileId: '',
        joinedAt: item.record.createdAt || now
      });
    }
  }
  return { ...classRecord, members, updatedAt: now };
}

export function analyzeClass(portfolioItems, channels = new Map(), today = new Date()) {
  const all = portfolioItems.flatMap(({ record }) => (record.journals || []).map((journal) => ({ ...journal, student: record.student })))
    .sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
  const dateKey = (value) => {
    const date = new Date(value);
    return [date.getFullYear(), String(date.getMonth() + 1).padStart(2, '0'), String(date.getDate()).padStart(2, '0')].join('-');
  };
  const todayKey = dateKey(today);
  const submittedToday = new Set(all.filter((journal) => dateKey(journal.createdAt) === todayKey).map((journal) => journal.student.email)).size;
  const alerts = [];
  for (const { record } of portfolioItems) {
    const journals = [...(record.journals || [])].sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
    const recent = journals.slice(0, 3);
    if (recent.filter((journal) => ['😕', '😠', '🤔'].includes(journal.emotion)).length >= 2) {
      alerts.push({ email: record.student.email, name: record.student.name, reason: '気になる気持ちが続いています' });
    }
    if (journals.length >= 3) {
      const average = journals.slice(1).reduce((sum, journal) => sum + journal.content.length, 0) / (journals.length - 1);
      if (average > 0 && journals[0].content.length < average * .3) alerts.push({ email: record.student.email, name: record.student.name, reason: '記述量が急に減っています' });
    }
  }
  const returned = all.filter((journal) => channels.get(normalizeEmail(journal.student.email))?.feedback?.[journal.id]?.returned).length;
  return { all, submittedToday, totalStudents: portfolioItems.length, returned, alerts };
}

export function exportCsv(portfolioItems, channels = new Map()) {
  const rows = [['日時', '児童名', 'メールアドレス', 'テーマ', '本文', '気持ち', '教師コメント', 'スタンプ', '返却状況']];
  for (const { record } of portfolioItems) {
    const channel = channels.get(normalizeEmail(record.student.email));
    for (const journal of record.journals || []) {
      const feedback = channel?.feedback?.[journal.id] || {};
      rows.push([journal.createdAt, record.student.name, record.student.email, journal.theme, journal.content, journal.emotion, feedback.comment, feedback.stamp, feedback.returned ? '返却済み' : '未返却']);
    }
  }
  const csv = rows.map((row) => row.map((value) => `"${String(value || '').replace(/"/g, '""')}"`).join(',')).join('\r\n');
  return '\ufeff' + csv;
}

export function driveQueryValue(value) {
  return String(value || '').replace(/\\/g, '\\\\').replace(/'/g, "\\'");
}

export function classAppProperties(classId, type, extra = {}) {
  return { rjSchema: String(SCHEMA_VERSION), rjType: type, rjClassId: String(classId), ...extra };
}

export function inviteUrl(baseUrl, invite) {
  const url = new URL(baseUrl);
  url.search = '';
  url.hash = `join=${encodeInvite(invite)}`;
  return url.href;
}
