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

function canonicalPublicKey(jwk) {
  return { kty: String(jwk?.kty || ''), crv: String(jwk?.crv || ''), x: String(jwk?.x || ''), y: String(jwk?.y || '') };
}

function invitePayload(invite, security = null) {
  const compact = {
    v: SCHEMA_VERSION,
    c: normalizeClassCode(invite.classCode),
    n: String(invite.className || '').trim().slice(0, 80),
    e: normalizeEmail(invite.teacherEmail),
    t: String(invite.teacherName || '').trim().slice(0, 80),
    a: invite.approvalRequired !== false,
    o: invite.acceptingMembers !== false
  };
  if (security) {
    compact.g = Math.max(1, Number.parseInt(security.generation, 10) || 1);
    compact.x = String(security.expiresAt || '');
    compact.i = String(security.keyId || '');
    compact.k = canonicalPublicKey(security.publicKeyJwk);
  }
  return compact;
}

export function encodeInvite(invite) {
  return bytesToBase64Url(new TextEncoder().encode(JSON.stringify(invitePayload(invite))));
}

export async function createInviteSecurity({ generation = 1, validityDays = 35, now = new Date(), cryptoObject = globalThis.crypto } = {}) {
  const pair = await cryptoObject.subtle.generateKey(
    { name: 'ECDSA', namedCurve: 'P-256' }, true, ['sign', 'verify']
  );
  const publicKeyJwk = canonicalPublicKey(await cryptoObject.subtle.exportKey('jwk', pair.publicKey));
  const privateKeyJwk = await cryptoObject.subtle.exportKey('jwk', pair.privateKey);
  const keyId = await sha256(`reflection-journal:invite-key:v1:${publicKeyJwk.crv}:${publicKeyJwk.x}:${publicKeyJwk.y}`, cryptoObject);
  const expiresAt = new Date(now.getTime() + Math.max(1, Number(validityDays) || 35) * 86400000).toISOString();
  return { algorithm: 'ECDSA-P256-SHA256', generation, keyId, publicKeyJwk, privateKeyJwk, expiresAt, createdAt: now.toISOString() };
}

export async function ensureInviteSecurity(classRecord, { now = new Date(), cryptoObject = globalThis.crypto } = {}) {
  const current = classRecord?.settings?.inviteSecurity;
  if (current?.keyId && current?.publicKeyJwk?.x && current?.privateKeyJwk?.d && new Date(current.expiresAt).getTime() > now.getTime()) {
    return { record: classRecord, changed: false };
  }
  const generation = Math.max(1, Number.parseInt(current?.generation, 10) || 0) + (current ? 1 : 0);
  const inviteSecurity = await createInviteSecurity({ generation, validityDays: classRecord?.settings?.inviteValidityDays || 35, now, cryptoObject });
  return {
    changed: true,
    record: {
      ...classRecord,
      settings: { ...defaultSettings(), ...(classRecord.settings || {}), inviteSecurity },
      updatedAt: now.toISOString()
    }
  };
}

export async function rotateInviteSecurity(classRecord, options = {}) {
  const currentGeneration = Number.parseInt(classRecord?.settings?.inviteSecurity?.generation, 10) || 0;
  const inviteSecurity = await createInviteSecurity({
    generation: currentGeneration + 1,
    validityDays: classRecord?.settings?.inviteValidityDays || 35,
    ...options
  });
  return {
    ...classRecord,
    settings: { ...defaultSettings(), ...(classRecord.settings || {}), inviteSecurity },
    updatedAt: (options.now || new Date()).toISOString()
  };
}

export async function encodeSignedInvite(invite, security, cryptoObject = globalThis.crypto) {
  if (!security?.privateKeyJwk?.d || !security?.publicKeyJwk?.x || !security?.keyId) throw new Error('招待署名の準備が完了していません。');
  const payload = bytesToBase64Url(new TextEncoder().encode(JSON.stringify(invitePayload(invite, security))));
  const privateKey = await cryptoObject.subtle.importKey(
    'jwk', security.privateKeyJwk, { name: 'ECDSA', namedCurve: 'P-256' }, false, ['sign']
  );
  const signature = await cryptoObject.subtle.sign(
    { name: 'ECDSA', hash: 'SHA-256' }, privateKey, new TextEncoder().encode(payload)
  );
  return `rj2.${payload}.${bytesToBase64Url(new Uint8Array(signature))}`;
}

export async function decodeInvite(encoded, { allowExpired = false, now = Date.now() } = {}) {
  let raw;
  let signed = false;
  try {
    const parts = String(encoded || '').split('.');
    if (parts.length === 3 && parts[0] === 'rj2') {
      const payload = parts[1];
      raw = JSON.parse(new TextDecoder().decode(base64UrlToBytes(payload)));
      const publicKey = await globalThis.crypto.subtle.importKey(
        'jwk', canonicalPublicKey(raw.k), { name: 'ECDSA', namedCurve: 'P-256' }, false, ['verify']
      );
      signed = await globalThis.crypto.subtle.verify(
        { name: 'ECDSA', hash: 'SHA-256' }, publicKey, base64UrlToBytes(parts[2]), new TextEncoder().encode(payload)
      );
      if (!signed) throw new Error('invalid signature');
    } else {
      raw = JSON.parse(new TextDecoder().decode(base64UrlToBytes(encoded)));
    }
  }
  catch (error) { throw new Error('招待情報を読み取れませんでした。先生から新しいURLを受け取ってください。'); }
  const invite = {
    version: Number(raw.v),
    classCode: normalizeClassCode(raw.c),
    className: String(raw.n || '').trim().slice(0, 80),
    teacherEmail: normalizeEmail(raw.e),
    teacherName: String(raw.t || '').trim().slice(0, 80),
    approvalRequired: raw.a !== false,
    acceptingMembers: raw.o !== false,
    signed,
    generation: Math.max(0, Number.parseInt(raw.g, 10) || 0),
    expiresAt: String(raw.x || ''),
    keyId: String(raw.i || ''),
    publicKeyJwk: raw.k ? canonicalPublicKey(raw.k) : null,
    token: String(encoded || '')
  };
  if (invite.version !== SCHEMA_VERSION || invite.classCode.length < 6 || !invite.className || !isEmail(invite.teacherEmail)) {
    throw new Error('招待情報が正しくありません。先生から新しいURLを受け取ってください。');
  }
  if (signed && (!invite.keyId || !invite.expiresAt || (!allowExpired && new Date(invite.expiresAt).getTime() <= now))) {
    throw new Error('この招待URLの有効期限が切れています。先生から新しいURLを受け取ってください。');
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
    inviteValidityDays: 35,
    inviteSecurity: null,
    geminiApiKey: '',
    geminiModel: 'gemini-3.1-flash-lite',
    geminiDataConsent: false
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
      acceptingMembers: invite.acceptingMembers !== false,
      inviteToken: String(invite.token || ''),
      inviteKeyId: String(invite.keyId || ''),
      inviteGeneration: Math.max(0, Number.parseInt(invite.generation, 10) || 0),
      inviteExpiresAt: String(invite.expiresAt || '')
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
  const highlights = (Array.isArray(feedback.highlights) ? feedback.highlights : []).slice(0, 20).map((item, index) => ({
    id: String(item?.id || `highlight-${index}`).slice(0, 100),
    start: Math.max(0, Number.parseInt(item?.start, 10) || 0),
    end: Math.max(0, Number.parseInt(item?.end, 10) || 0),
    text: String(item?.text || '').slice(0, 1000),
    comment: String(item?.comment || '').trim().slice(0, 1000),
    stamp: String(item?.stamp || '').slice(0, 8)
  })).filter((item) => item.end > item.start && item.text);
  return {
    ...channel,
    feedback: {
      ...(channel.feedback || {}),
      [id]: {
        comment: String(feedback.comment || '').trim().slice(0, 4000),
        stamp: String(feedback.stamp || '').slice(0, 8),
        highlights,
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
  let changed = false;
  for (const item of portfolioItems) {
    const email = normalizeEmail(item.record?.student?.email);
    if (!email) continue;
    const existing = members.find((member) => normalizeEmail(member.email) === email);
    if (existing) {
      const nextFileId = item.file.id;
      const nextName = existing.name || item.record.student.name;
      const nextJoinedAt = existing.joinedAt || item.record.createdAt || now;
      const nextStatus = existing.status === 'invited' ? 'active' : existing.status;
      if (existing.portfolioFileId !== nextFileId || existing.name !== nextName || existing.joinedAt !== nextJoinedAt || existing.status !== nextStatus) changed = true;
      existing.portfolioFileId = nextFileId;
      existing.name = nextName;
      existing.joinedAt = nextJoinedAt;
      existing.status = nextStatus;
    } else {
      changed = true;
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
  return changed ? { ...classRecord, members, updatedAt: now } : classRecord;
}

function fileOwnerEmail(file) {
  return normalizeEmail(file?.owners?.[0]?.emailAddress || '');
}

export async function validatePortfolioForClass(file, record, classRecord) {
  const studentEmail = normalizeEmail(record?.student?.email);
  const ownerEmail = fileOwnerEmail(file);
  if (record?.kind !== 'reflection-journal-portfolio') return { ok: false, reason: 'ポートフォリオ形式が正しくありません。' };
  if (record?.class?.id !== classRecord?.classId || normalizeEmail(record?.class?.teacherEmail) !== normalizeEmail(classRecord?.teacher?.email)) {
    return { ok: false, reason: 'クラス情報が一致しません。' };
  }
  if (!studentEmail || !ownerEmail || studentEmail !== ownerEmail) return { ok: false, reason: '児童アカウントとDrive所有者が一致しません。' };
  const existing = (classRecord.members || []).find((member) => normalizeEmail(member.email) === studentEmail);
  if (existing?.portfolioFileId === file.id) return { ok: true, legacy: !record.class.inviteToken };
  if (!record.class.inviteToken) return { ok: false, reason: '署名された招待情報がありません。' };
  let invite;
  try { invite = await decodeInvite(record.class.inviteToken, { allowExpired: true }); }
  catch (error) { return { ok: false, reason: error.message }; }
  const security = classRecord?.settings?.inviteSecurity;
  if (!invite.signed || invite.classId !== classRecord.classId || invite.keyId !== security?.keyId || invite.generation !== security?.generation) {
    return { ok: false, reason: '招待URLが失効しているか、正しいクラスから発行されていません。' };
  }
  const createdAt = new Date(file?.createdTime || record?.createdAt || '').getTime();
  const expiresAt = new Date(invite.expiresAt).getTime();
  if (!Number.isFinite(createdAt) || createdAt > expiresAt) return { ok: false, reason: '招待URLの有効期限後に作成された記録です。' };
  return { ok: true, invite };
}

export function validateChannelForStudent(file, record, studentEmail, classId) {
  if (record?.kind !== 'reflection-journal-channel' || record?.class?.id !== classId) return false;
  if (normalizeEmail(record?.student?.email) !== normalizeEmail(studentEmail)) return false;
  return fileOwnerEmail(file) === normalizeEmail(record?.teacher?.email);
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

export function signedInviteUrl(baseUrl, token) {
  const url = new URL(baseUrl);
  url.search = '';
  url.hash = `join=${String(token || '')}`;
  return url.href;
}
