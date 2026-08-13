const CLASS_CODE_ALPHABET = '23456789ABCDEFGHJKLMNPQRSTUVWXYZ';
const SCHEMA_VERSION = 1;

export function normalizeClassCode(value) {
  return String(value || '')
    .toUpperCase()
    .replace(/[^23456789ABCDEFGHJKLMNPQRSTUVWXYZ]/g, '')
    .slice(0, 10);
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

export async function computeClassId(teacherEmail, classCode, cryptoObject = globalThis.crypto) {
  const source = `reflection-journal:v1:${normalizeEmail(teacherEmail)}:${normalizeClassCode(classCode)}`;
  const digest = await cryptoObject.subtle.digest('SHA-256', new TextEncoder().encode(source));
  return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, '0')).join('');
}

function bytesToBase64Url(bytes) {
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
}

function base64UrlToBytes(value) {
  const padded = String(value).replace(/-/g, '+').replace(/_/g, '/') + '==='.slice((String(value).length + 3) % 4);
  const binary = atob(padded);
  return Uint8Array.from(binary, (char) => char.charCodeAt(0));
}

export function encodeInvite(invite) {
  const compact = {
    v: SCHEMA_VERSION,
    c: normalizeClassCode(invite.classCode),
    n: String(invite.className || '').trim().slice(0, 80),
    e: normalizeEmail(invite.teacherEmail),
    t: String(invite.teacherName || '').trim().slice(0, 80)
  };
  return bytesToBase64Url(new TextEncoder().encode(JSON.stringify(compact)));
}

export async function decodeInvite(encoded) {
  let raw;
  try {
    raw = JSON.parse(new TextDecoder().decode(base64UrlToBytes(encoded)));
  } catch (error) {
    throw new Error('招待情報を読み取れませんでした。先生から新しいURLを受け取ってください。');
  }
  const invite = {
    version: Number(raw.v),
    classCode: normalizeClassCode(raw.c),
    className: String(raw.n || '').trim().slice(0, 80),
    teacherEmail: normalizeEmail(raw.e),
    teacherName: String(raw.t || '').trim().slice(0, 80)
  };
  if (invite.version !== SCHEMA_VERSION || invite.classCode.length < 6 || !invite.className || !isEmail(invite.teacherEmail)) {
    throw new Error('招待情報が正しくありません。先生から新しいURLを受け取ってください。');
  }
  invite.classId = await computeClassId(invite.teacherEmail, invite.classCode);
  return invite;
}

export function createClassRecord({ classId, classCode, className, teacher, now = new Date().toISOString() }) {
  return {
    schemaVersion: SCHEMA_VERSION,
    kind: 'reflection-journal-class',
    classId,
    classCode: normalizeClassCode(classCode),
    className: String(className || '').trim(),
    teacher: { email: normalizeEmail(teacher.email), name: String(teacher.name || '').trim() },
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
      teacherName: String(invite.teacherName || '').trim()
    },
    student: { email: normalizeEmail(student.email), name: String(student.name || '').trim() },
    journals: [],
    createdAt: now,
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
    createdAt: journal.createdAt || now,
    updatedAt: now
  };
  if (!item.id || !item.content) throw new Error('本文を入力してください。');
  return { ...portfolio, journals: [...(portfolio.journals || []), item], updatedAt: now };
}

export function driveQueryValue(value) {
  return String(value || '').replace(/\\/g, '\\\\').replace(/'/g, "\\'");
}

export function classAppProperties(classId, type) {
  return { rjSchema: String(SCHEMA_VERSION), rjType: type, rjClassId: String(classId) };
}

export function inviteUrl(baseUrl, invite) {
  const url = new URL(baseUrl);
  url.search = '';
  url.hash = `join=${encodeInvite(invite)}`;
  return url.href;
}

