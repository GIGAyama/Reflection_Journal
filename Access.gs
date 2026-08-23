/**
 * Access.gs — 誰が何をしてよいかの判断
 *
 * 認可の鉄則: すべての API はサーバー側で
 *   ① 本人確認（Session.getActiveUser）→ ② 名簿照合（状態 = active）
 *   → ③ 役割チェック → ④ 行の持ち主チェック
 * の順にガードする。画面側の出し分けは防御とみなさない。
 *
 * `google.script.run` は末尾 `_` の無いトップレベル関数を誰でも直接呼べる。
 * 児童の端末のコンソールから opSaveFeedback('...') を打てば、その関数に認可が
 * 無いかぎり通ってしまう。**新しく公開関数を作ったら、必ずここのガードを通すこと。**
 *
 * 児童向けのレスポンスでは、ほかの児童のメールアドレスを一切出さない。
 * 匿名 ID（uid = sha256(email) の先頭 12 桁）に置き換える。
 */

/** 児童側にメールアドレスを出さないための匿名 ID */
function uidOf_(email) {
  return 'u' + sha256Hex_(String(email).toLowerCase()).slice(0, 12);
}

// ────────────────────────────────────────────────────────────────
// 名簿照合（児童名簿シート: 役割 / 氏名 / メールアドレス / 状態）
//   役割: '担任' | '児童'   状態: 'active' | 'pending'
// ────────────────────────────────────────────────────────────────

function getMembers_(ss) {
  return getRosterRows_(ss);   // Db.gs 側で実装（列は見出しの名前で引く）
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
 * 児童 API 共通ガード（①②）。戻り値 { email, ss, member } を各 API が使う。
 * ロックの外で行うこと（本人確認とシート読み取りにロックは要らない）。
 */
function guardMember_() {
  const email = requireEmail_();               // ① 本人確認
  const ss = getDb_();
  const member = assertActiveMember_(ss, email); // ② 名簿照合
  return { email: email, ss: ss, member: member };
}

/**
 * 先生 API 共通ガード（①③）。
 *
 * 先生かどうかは _meta の ownerEmail か、名簿の「担任」で決まる。どちらも
 * スプレッドシートを開ける人しか書けない場所なので、ウェブアプリから先生になる道は無い。
 */
function assertOwner_() {
  const email = requireEmail_();               // ① 本人確認
  const ss = getDb_();
  if (!ownerEmailOf_(ss)) {
    throw new Error('SETUP_REQUIRED: このクラスはまだ設定されていません。' +
      '（先生へ: スプレッドシートを開き、メニュー「' + CONFIG.APP_NAME + '」＞「はじめの設定」を 1 回押してください）');
  }
  if (!isOwnerEmail_(ss, email)) {             // ③ 役割チェック
    throw new Error('FORBIDDEN: この操作ができるのは、このクラスの先生だけです');
  }
  return { email: email, ss: ss };
}

/**
 * ④ 行の持ち主チェック。update / delete 系の API は必ずこれを通す。
 * 先生が他人の行も操作してよい場合だけ allowOwnerRole を true にする。
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
    out.email = uidOf_(j.email);   // 生のメールアドレスを必ず潰す
    return out;
  });
}

// ────────────────────────────────────────────────────────────────
// 入力の確かめ（payload は決めたキーだけ書き込む）
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
 * ジャーナル ID の確かめ。形が違えば空を返し、呼び出し側で NOT_FOUND 扱いにする
 * （他人の行に当たらないことを保証する）。
 */
function vJournalId_(v) {
  const s = String(v || '');
  return /^[A-Za-z0-9\-]{8,50}$/.test(s) ? s : '';
}
