// ============================================================
// ふりかえりジャーナル — 記録の形とクラス運用の規則
// ------------------------------------------------------------
// 分散ポートフォリオの共通部分（ID・招待の署名・共有記録の検証・名簿統合）は
// docs/kit/ に切り出してある。このファイルはふりかえりジャーナル固有の
// 「何を書き、どう返すか」だけを持つ。
// 他アプリへ横展開するときは kit/ をそのまま持っていき、このファイルに当たる
// 部分だけを書き直す（docs/PORTING_FROM_GAS.md）。
// ============================================================

import {
  createInviteKey,
  createdWithinInvite,
  decodeInvite as decodeKitInvite,
  encodeInvite as encodeKitInvite,
  encodeSignedInvite as signKitInvite,
  inviteKeyUsable,
  inviteUrl as kitInviteUrl,
  matchesIssuedKey
} from './kit/invite.js';
import { defineAppNamespace, driveQueryValue, isEmail, normalizeEmail, sha256 } from './kit/namespace.js';
import { mergeSharedIntoRoster, ownerEmailOf, validateSharedRecord } from './kit/records.js';

export const SCHEMA_VERSION = 2;

/** このアプリの印。IDと検索条件はすべてここから作る。 */
export const RJ = defineAppNamespace({
  appId: 'reflection-journal',
  propertyPrefix: 'rj',
  schemaVersion: SCHEMA_VERSION,
  terms: { tenant: 'Class', member: 'Student' }
});

export const KIND = {
  class: 'reflection-journal-class',
  portfolio: 'reflection-journal-portfolio',
  channel: 'reflection-journal-channel'
};

// 招待の失敗理由を、児童にも読める日本語へ直す。キットは文言を持たない。
const INVITE_MESSAGES = {
  unreadable: '招待情報を読み取れませんでした。先生から新しいURLを受け取ってください。',
  invalid: '招待情報が正しくありません。先生から新しいURLを受け取ってください。',
  expired: 'この招待URLの有効期限が切れています。先生から新しいURLを受け取ってください。',
  unsigned: '招待署名の準備が完了していません。'
};

function inviteError(error) {
  return new Error(INVITE_MESSAGES[error?.code] || INVITE_MESSAGES.unreadable);
}

export { driveQueryValue, isEmail, normalizeEmail, sha256 };

export function normalizeClassCode(value) {
  return RJ.normalizeCode(value);
}

export function randomClassCode(length = 8, cryptoObject = globalThis.crypto) {
  return RJ.randomCode(length, cryptoObject);
}

export function computeClassId(teacherEmail, classCode, cryptoObject = globalThis.crypto) {
  return RJ.tenantId(teacherEmail, classCode, cryptoObject);
}

export function studentKey(email) {
  return RJ.memberKey(email);
}

export function classAppProperties(classId, type, extra = {}) {
  return RJ.appProperties(classId, type, extra);
}

// ── 招待（クラス用語 ⇄ キットの汎用用語の変換だけを持つ） ──

function toKitInvite(invite) {
  return {
    code: invite.classCode,
    name: invite.className,
    hostEmail: invite.teacherEmail,
    hostName: invite.teacherName,
    approvalRequired: invite.approvalRequired,
    acceptingMembers: invite.acceptingMembers
  };
}

function fromKitInvite(invite) {
  return {
    version: invite.version,
    classCode: invite.code,
    className: invite.name,
    teacherEmail: invite.hostEmail,
    teacherName: invite.hostName,
    approvalRequired: invite.approvalRequired,
    acceptingMembers: invite.acceptingMembers,
    signed: invite.signed,
    generation: invite.generation,
    expiresAt: invite.expiresAt,
    keyId: invite.keyId,
    publicKeyJwk: invite.publicKeyJwk,
    token: invite.token,
    classId: invite.tenantId
  };
}

export function encodeInvite(invite) {
  return encodeKitInvite(RJ, toKitInvite(invite));
}

export function createInviteSecurity(options = {}) {
  return createInviteKey(RJ, options);
}

export async function encodeSignedInvite(invite, security, cryptoObject = globalThis.crypto) {
  try { return await signKitInvite(RJ, toKitInvite(invite), security, cryptoObject); }
  catch (error) { throw error?.code ? inviteError(error) : error; }
}

export async function decodeInvite(encoded, options = {}) {
  try { return fromKitInvite(await decodeKitInvite(RJ, encoded, options)); }
  catch (error) { throw error?.code ? inviteError(error) : error; }
}

export async function ensureInviteSecurity(classRecord, { now = new Date(), cryptoObject = globalThis.crypto } = {}) {
  const current = classRecord?.settings?.inviteSecurity;
  if (inviteKeyUsable(current, now)) return { record: classRecord, changed: false };
  const generation = Math.max(1, Number.parseInt(current?.generation, 10) || 0) + (current ? 1 : 0);
  const inviteSecurity = await createInviteKey(RJ, {
    generation, validityDays: classRecord?.settings?.inviteValidityDays || 35, now, cryptoObject
  });
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
  const inviteSecurity = await createInviteKey(RJ, {
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

export function inviteUrl(baseUrl, invite) {
  return kitInviteUrl(baseUrl, encodeInvite(invite));
}

export function signedInviteUrl(baseUrl, token) {
  return kitInviteUrl(baseUrl, token);
}

// ── 記録の形 ──

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
    kind: KIND.class,
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
    kind: KIND.portfolio,
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
    kind: KIND.channel,
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
  if (!portfolio || portfolio.kind !== KIND.portfolio) throw new Error('ポートフォリオ形式が正しくありません。');
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

// ── 共有記録の取り込み ──

export function mergePortfoliosIntoMembers(classRecord, portfolioItems, now = new Date().toISOString()) {
  const { members, changed } = mergeSharedIntoRoster(classRecord.members, portfolioItems, {
    subjectOf: (record) => record?.student,
    fileIdField: 'portfolioFileId',
    status: classRecord.settings?.approvalRequired === false ? 'active' : 'pending',
    defaults: { role: 'student', channelFileId: '' },
    now
  });
  return changed ? { ...classRecord, members, updatedAt: now } : classRecord;
}

/**
 * 共有されたポートフォリオを名簿へ入れてよいかを決める。
 * キットの基本検証（種別・クラス・Drive所有者）に、このアプリ固有の
 * 「署名付き招待が現世代の鍵から出ていること」を足している。
 */
export async function validatePortfolioForClass(file, record, classRecord) {
  const base = validateSharedRecord(file, record, {
    kind: KIND.portfolio,
    tenantId: classRecord?.classId,
    tenantIdOf: (value) => value?.class?.id,
    ownerMustBe: (value) => value?.student?.email
  });
  if (!base.ok) {
    if (base.reason === 'kind') return { ok: false, reason: 'ポートフォリオ形式が正しくありません。' };
    if (base.reason === 'owner') return { ok: false, reason: '児童アカウントとDrive所有者が一致しません。' };
    return { ok: false, reason: 'クラス情報が一致しません。' };
  }
  if (normalizeEmail(record?.class?.teacherEmail) !== normalizeEmail(classRecord?.teacher?.email)) {
    return { ok: false, reason: 'クラス情報が一致しません。' };
  }
  const studentEmail = normalizeEmail(record.student.email);
  const existing = (classRecord.members || []).find((member) => normalizeEmail(member.email) === studentEmail);
  if (existing?.portfolioFileId === file.id) return { ok: true, legacy: !record.class.inviteToken };
  if (!record.class.inviteToken) return { ok: false, reason: '署名された招待情報がありません。' };
  let invite;
  try { invite = await decodeInvite(record.class.inviteToken, { allowExpired: true }); }
  catch (error) { return { ok: false, reason: error.message }; }
  const issued = matchesIssuedKey(
    { ...invite, tenantId: invite.classId },
    classRecord?.settings?.inviteSecurity,
    { tenantId: classRecord.classId }
  );
  if (!issued.ok) return { ok: false, reason: '招待URLが失効しているか、正しいクラスから発行されていません。' };
  if (!createdWithinInvite(file?.createdTime || record?.createdAt, invite)) {
    return { ok: false, reason: '招待URLの有効期限後に作成された記録です。' };
  }
  return { ok: true, invite };
}

export function validateChannelForStudent(file, record, studentEmail, classId) {
  return validateSharedRecord(file, record, {
    kind: KIND.channel,
    tenantId: classId,
    tenantIdOf: (value) => value?.class?.id,
    ownerMustBe: (value) => value?.teacher?.email,
    subjectEmail: studentEmail,
    subjectEmailOf: (value) => value?.student?.email
  }).ok;
}

export { ownerEmailOf };

// ── 学級全体の見立て ──

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
