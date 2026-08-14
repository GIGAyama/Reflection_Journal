// ============================================================
// Driveネイティブ・キット — 名前空間（アプリ固有の「印」を1か所で決める）
// ------------------------------------------------------------
// 中央DBを持たない分散ポートフォリオでは、テナント（学級・教室・講座）の
// 識別子とDriveの appProperties だけが「どのアプリの、どのテナントのファイルか」
// を示す。ここを各アプリで手書きすると、衝突・取り違え・移行不能が起きる。
//
// アプリごとに defineAppNamespace() を1回だけ呼び、以降のIDと検索条件は
// すべてこの戻り値から作る。
// ============================================================

const DEFAULT_CODE_ALPHABET = '23456789ABCDEFGHJKLMNPQRSTUVWXYZ';

export function normalizeEmail(value) {
  return String(value || '').trim().toLowerCase();
}

export function isEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normalizeEmail(value));
}

export async function sha256(value, cryptoObject = globalThis.crypto) {
  const digest = await cryptoObject.subtle.digest('SHA-256', new TextEncoder().encode(String(value)));
  return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, '0')).join('');
}

export function bytesToBase64Url(bytes) {
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
}

export function base64UrlToBytes(value) {
  const input = String(value);
  const padded = input.replace(/-/g, '+').replace(/_/g, '/') + '==='.slice((input.length + 3) % 4);
  return Uint8Array.from(atob(padded), (char) => char.charCodeAt(0));
}

/** Drive検索式へ値を埋めるときのエスケープ（引用符とバックスラッシュ） */
export function driveQueryValue(value) {
  return String(value || '').replace(/\\/g, '\\\\').replace(/'/g, "\\'");
}

/**
 * アプリ固有の名前空間を作る。
 *
 * @param {object} options
 * @param {string} options.appId          ID生成の接頭辞。他アプリと必ず異なる値にする（例 'daily-journal'）。
 * @param {string} options.propertyPrefix Driveの appProperties キーの接頭辞（例 'rj'）。短く、他アプリと衝突しない値。
 * @param {number} options.schemaVersion  保存形式の版。上げると過去のIDと招待は無効になる。
 * @param {object} [options.terms]        画面用語。{ tenant: 'Class', member: 'Student' } のように大文字始まりで渡す。
 * @param {string} [options.memberLabel]  メンバー鍵の用途文字列（既定は terms.member の小文字）。
 */
export function defineAppNamespace({
  appId,
  propertyPrefix,
  schemaVersion = 1,
  terms = {},
  memberLabel = '',
  codeAlphabet = DEFAULT_CODE_ALPHABET,
  codeLength = 8,
  codeMaxLength = 10,
  minCodeLength = 6,
  tokenPrefix = ''
} = {}) {
  if (!appId || !propertyPrefix) throw new Error('defineAppNamespace: appId と propertyPrefix は必須です。');
  const tenantTerm = String(terms.tenant || 'Tenant');
  const memberTerm = String(terms.member || 'Member');
  const memberPurpose = memberLabel || memberTerm.toLowerCase();
  const codePattern = new RegExp(`[^${codeAlphabet.replace(/[\\\]^-]/g, '\\$&')}]`, 'g');
  const properties = Object.freeze({
    schema: `${propertyPrefix}Schema`,
    type: `${propertyPrefix}Type`,
    tenantId: `${propertyPrefix}${tenantTerm}Id`,
    member: `${propertyPrefix}${memberTerm}`
  });

  return Object.freeze({
    appId,
    propertyPrefix,
    schemaVersion,
    properties,
    codeLength,
    minCodeLength,
    // 署名付き招待の先頭に付く版。復号側が形式を取り違えないようにする。
    tokenPrefix: tokenPrefix || `${propertyPrefix}${schemaVersion}`,

    normalizeCode(value) {
      return String(value || '').toUpperCase().replace(codePattern, '').slice(0, codeMaxLength);
    },

    randomCode(length = codeLength, cryptoObject = globalThis.crypto) {
      const bytes = new Uint8Array(length);
      cryptoObject.getRandomValues(bytes);
      return Array.from(bytes, (byte) => codeAlphabet[byte % codeAlphabet.length]).join('');
    },

    /** テナントID。主催者メールと参加コードから決定的に作るので、中央の採番簿が要らない。 */
    tenantId(hostEmail, code, cryptoObject = globalThis.crypto) {
      return sha256(`${appId}:v${schemaVersion}:${normalizeEmail(hostEmail)}:${this.normalizeCode(code)}`, cryptoObject);
    },

    /** メンバー識別子。Driveの appProperties に生のメールを残さないためのハッシュ。 */
    memberKey(email, cryptoObject = globalThis.crypto) {
      return sha256(`${appId}:${memberPurpose}:v${schemaVersion}:${normalizeEmail(email)}`, cryptoObject);
    },

    /** 招待署名鍵のID。公開鍵そのものから決まるので、鍵の取り違えを検出できる。 */
    inviteKeyId(publicKeyJwk, cryptoObject = globalThis.crypto) {
      return sha256(`${appId}:invite-key:v1:${publicKeyJwk?.crv}:${publicKeyJwk?.x}:${publicKeyJwk?.y}`, cryptoObject);
    },

    /** Driveファイルへ付ける印。これが検索の唯一の手がかりになる。 */
    appProperties(tenantId, type, extra = {}) {
      return {
        [properties.schema]: String(schemaVersion),
        [properties.type]: String(type),
        [properties.tenantId]: String(tenantId),
        ...extra
      };
    },

    /**
     * ファイル検索式。
     * @param {object} options
     * @param {string} options.type       ファイルの種別（<prefix>Type の値）。
     * @param {'me'|'shared'|'any'} [options.owner] 'me'=自分が所有 / 'shared'=共有された / 'any'=両方。
     * @param {string} [options.tenantId] 指定するとそのテナントだけに絞る。
     * @param {object} [options.properties] 追加の appProperties 一致条件。
     */
    query({ type, owner = 'any', tenantId = '', properties: extra = {} } = {}) {
      const has = (key, value) => `appProperties has { key='${driveQueryValue(key)}' and value='${driveQueryValue(value)}' }`;
      const parts = ['trashed = false'];
      if (owner === 'me') parts.push("'me' in owners");
      if (owner === 'shared') parts.push('sharedWithMe');
      parts.push(has(properties.type, type));
      if (tenantId) parts.push(has(properties.tenantId, tenantId));
      for (const [key, value] of Object.entries(extra)) parts.push(has(key, value));
      return parts.join(' and ');
    }
  });
}
