/**
 * Tenant.gs — テナント（クラス）解決とアクセス制御
 *
 * 認可の鉄則: すべての API はサーバー側で
 *   ① トークン検証（B）/ Session（A） → ② テナント解決 → ③ 名簿照合（状態=active）
 *   → ④ 役割チェック → ⑤ 行所有者チェック
 * の順にガードする。フロントの出し分けは防御とみなさない。
 *
 * スプレッドシート ID は児童側（URL・API レスポンス・HTML）に一切露出しない。
 * 児童向けレスポンスでは email も露出せず、匿名 ID（uid = sha256(email) 先頭12桁）に置換する。
 */

/** テナント解決: レジストリ → openById。開けない場合は復旧手順つきのエラー */
function openTenantSs_(code) {
  const rec = requireTenantRecord_(code);
  try {
    return SpreadsheetApp.openById(rec.spreadsheetId);
  } catch (e) {
    // 先生が共有を外した / シートを削除した場合にここへ来る
    throw new Error('TENANT_UNAVAILABLE: データベースにアクセスできません。先生に確認してください。' +
      '（先生へ: スプレッドシートが削除されていないか、共有設定で ' +
      getSetting_(PROP_KEYS.APP_ACCOUNT, false) + ' が「編集者」のままかを確認し、' +
      '外れている場合は共有し直してください）');
  }
}

/** 児童側に email を出さないための匿名 ID */
function uidOf_(email) {
  return 'u' + sha256Hex_(String(email).toLowerCase()).slice(0, 12);
}

// ────────────────────────────────────────────────────────────────
// 名簿照合（児童名簿シート: 役割 / 氏名 / メールアドレス / 状態）
//   役割: '担任' | '児童'   状態: 'active' | 'pending'
// ────────────────────────────────────────────────────────────────

function getMembers_(ss) {
  return getRosterRows_(ss);   // Db.gs 側で実装
}

function getMemberRow_(ss, email) {
  const target = String(email).toLowerCase();
  const members = getMembers_(ss);
  for (let i = 0; i < members.length; i++) {
    if (String(members[i].email).toLowerCase() === target) return members[i];
  }
  return null;
}

function assertActiveMember_(ss, email) {
  const m = getMemberRow_(ss, email);
  if (!m) {
    throw new Error('NOT_MEMBER: 名簿に登録されていません。先生に確認してください');
  }
  if (m.status !== 'active') {
    throw new Error('MEMBER_PENDING: 参加申請は先生の承認待ちです。承認されたらもう一度開いてください');
  }
  return m;
}

/**
 * 児童 API 共通ガード（多段ガードの ①〜③）。
 * 戻り値 { user, rec, ss, code, member } を各 API が使う。
 * ロック外で行うこと（トークン検証・シート読み取りはロック不要）。
 */
function guardMember_(idToken, tenantCode) {
  const user = verifyIdToken_(idToken);                // ① トークン検証
  const code = normalizeTenantCode_(tenantCode);
  const rec = requireTenantRecord_(code);              // ② テナント解決
  const ss = openTenantSs_(code);
  const member = assertActiveMember_(ss, user.email);  // ③ 名簿照合
  return { user: user, rec: rec, ss: ss, code: code, member: member };
}

/**
 * ⑤ 行所有者チェック。update/delete 系の API は必ずこれを通す。
 * 担任は他人の行も操作できてよい場合のみ allowOwnerRole を true にする。
 */
function assertRowOwner_(rowEmail, email, member, allowOwnerRole) {
  if (allowOwnerRole && member && member.role === '担任') return;
  if (!rowEmail || String(rowEmail).toLowerCase() !== String(email).toLowerCase()) {
    throw new Error('FORBIDDEN: 自分のデータ以外は変更できません');
  }
}

// ────────────────────────────────────────────────────────────────
// 児童向けレスポンスのサニタイズ（email → uid 置換）
// レスポンス組み立てはここに集約する。API ごとに手で消すと必ず漏れる。
// ────────────────────────────────────────────────────────────────

function sanitizeJournals_(journals) {
  return journals.map(function (j) {
    const out = {};
    Object.keys(j).forEach(function (k) { out[k] = j[k]; });
    out.email = uidOf_(j.email);   // 生 email を必ず潰す
    return out;
  });
}

// ────────────────────────────────────────────────────────────────
// 入力バリデーション（payload はホワイトリストしたキーのみ書き込む）
// ────────────────────────────────────────────────────────────────

function vStr_(v, max, label) {
  const s = (v === null || v === undefined) ? '' : String(v);
  if (s.length > max) throw new Error('BAD_INPUT: ' + label + 'が長すぎます');
  return s;
}

function vNum_(v, min, max, label) {
  const n = Number(v);
  if (!isFinite(n) || n < min || n > max) throw new Error('BAD_INPUT: ' + label + 'の値が正しくありません');
  return n;
}

/**
 * ジャーナル ID の検証。UUID 形式でなければ空を返し、呼び出し側で
 * NOT_FOUND 扱いにする（他人の行に当たらないことを保証する）。
 */
function vJournalId_(v) {
  const s = String(v || '');
  return /^[A-Za-z0-9\-]{8,50}$/.test(s) ? s : '';
}
