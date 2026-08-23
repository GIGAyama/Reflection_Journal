// ============================================================
// Driveネイティブ・キット — 共有記録の検証・キャッシュ・名簿統合
// ------------------------------------------------------------
// 共有されたJSONは「相手のDriveにある、相手が書き換えられるファイル」である。
// 中身の自己申告（メールアドレスやテナントID）だけを信じると、共有さえできれば
// 他人の名簿へ入り込める。取り込む前に必ず
//   ① 種別 ② テナント ③ Driveの所有者 ④ 宛先
// の4点をDriveのメタデータ側と突き合わせる。
// ============================================================

import { normalizeEmail } from './namespace.js';

/** Driveメタデータ上の所有者。JSONの自己申告ではなく、こちらを真とする。 */
export function ownerEmailOf(file) {
  return normalizeEmail(file?.owners?.[0]?.emailAddress || '');
}

/**
 * 共有記録の基本検証。
 * @returns {{ok: boolean, reason?: 'kind'|'tenant'|'owner'|'subject'}}
 */
export function validateSharedRecord(file, record, {
  kind = '',
  tenantId = '',
  tenantIdOf = (value) => value?.tenant?.id,
  // このファイルを所有していなければならない人のメール（記録から引く）
  ownerMustBe = null,
  // 記録の宛先（誰のための記録か）と、それに一致すべきメール
  subjectEmail = '',
  subjectEmailOf = null
} = {}) {
  if (kind && record?.kind !== kind) return { ok: false, reason: 'kind' };
  if (tenantId && tenantIdOf(record) !== tenantId) return { ok: false, reason: 'tenant' };
  if (ownerMustBe) {
    const expected = normalizeEmail(ownerMustBe(record));
    const actual = ownerEmailOf(file);
    if (!expected || !actual || expected !== actual) return { ok: false, reason: 'owner' };
  }
  if (subjectEmail && subjectEmailOf) {
    if (normalizeEmail(subjectEmailOf(record)) !== normalizeEmail(subjectEmail)) return { ok: false, reason: 'subject' };
  }
  return { ok: true };
}

/**
 * 読み込んだJSONの再利用キャッシュ。
 * 一覧の定期更新のたびに全員分の本文を取り直すと、児童40人×30秒でDriveの
 * 実行回数を使い切る。ファイルID・modifiedTime・version が同じなら中身も同じ。
 */
export class RecordCache {
  constructor() {
    this.entries = new Map();
  }

  static stamp(file) {
    return `${String(file?.modifiedTime || '')}|${String(file?.version || '')}`;
  }

  remember(file, record) {
    if (!file?.id) return record;
    this.entries.set(file.id, { stamp: RecordCache.stamp(file), record });
    return record;
  }

  /** 版が一致すれば記録を返し、変わっていれば null（＝取り直しが必要）。 */
  read(file) {
    const cached = this.entries.get(file?.id);
    return cached && cached.stamp === RecordCache.stamp(file) ? cached.record : null;
  }

  forget(fileId) {
    this.entries.delete(fileId);
  }

  clear() {
    this.entries.clear();
  }
}

/**
 * 共有された記録を名簿へ取り込む。呼ぶ前に validateSharedRecord を通すこと。
 * 既存の行は上書きせず、ファイルIDと参加状態だけを追いつかせる。
 *
 * @returns {{members: object[], changed: boolean}}
 */
export function mergeSharedIntoRoster(members, items, {
  subjectOf = (record) => record?.member || {},
  fileIdField = 'fileId',
  createdAtOf = (record) => record?.createdAt,
  status = 'pending',
  defaults = {},
  now = new Date().toISOString()
} = {}) {
  const next = (members || []).map((member) => ({ ...member }));
  let changed = false;
  for (const item of items) {
    const subject = subjectOf(item.record) || {};
    const email = normalizeEmail(subject.email);
    if (!email) continue;
    const existing = next.find((member) => normalizeEmail(member.email) === email);
    if (existing) {
      const nextFileId = item.file.id;
      const nextName = existing.name || subject.name;
      const nextJoinedAt = existing.joinedAt || createdAtOf(item.record) || now;
      // 招待だけして未参加だった行は、記録が届いた時点で参加済みに進める。
      const nextStatus = existing.status === 'invited' ? 'active' : existing.status;
      if (existing[fileIdField] !== nextFileId || existing.name !== nextName
        || existing.joinedAt !== nextJoinedAt || existing.status !== nextStatus) changed = true;
      existing[fileIdField] = nextFileId;
      existing.name = nextName;
      existing.joinedAt = nextJoinedAt;
      existing.status = nextStatus;
    } else {
      changed = true;
      next.push({
        email,
        name: subject.name || email,
        ...defaults,
        status,
        [fileIdField]: item.file.id,
        joinedAt: createdAtOf(item.record) || now
      });
    }
  }
  return { members: next, changed };
}
