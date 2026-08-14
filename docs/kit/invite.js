// ============================================================
// Driveネイティブ・キット — 署名付き招待
// ------------------------------------------------------------
// 中央サーバーが無いので、「このテナントに参加してよい」という事実は
// 招待URLそのものが証明する。主催者だけが持つ秘密鍵で署名し、
// 参加者が作った記録を主催者が取り込むときに署名を検証する。
//
// 署名が無い招待は「誰でも作れるただの文字列」であり、名簿に他人が
// 紛れ込む余地を残す。新規テナントは必ず署名付きで発行する。
//
// 招待はURLのフラグメント（#）に載せる。フラグメントはサーバーへ送信されない
// ため、GitHub Pagesや中継のアクセスログに参加情報が残らない。
// ============================================================

import { base64UrlToBytes, bytesToBase64Url, isEmail, normalizeEmail } from './namespace.js';

const ALGORITHM = { name: 'ECDSA', namedCurve: 'P-256' };
const SIGN_PARAMS = { name: 'ECDSA', hash: 'SHA-256' };

/**
 * 招待の失敗理由。画面文言はアプリ側が code から決める（キットは日本語文言を持たない）。
 * unreadable = 読めない・署名不一致 / invalid = 中身が条件を満たさない / expired = 期限切れ
 */
export class InviteError extends Error {
  constructor(code, message = code) {
    super(message);
    this.name = 'InviteError';
    this.code = code;
  }
}

function canonicalPublicKey(jwk) {
  return { kty: String(jwk?.kty || ''), crv: String(jwk?.crv || ''), x: String(jwk?.x || ''), y: String(jwk?.y || '') };
}

function compactPayload(namespace, invite, security = null) {
  const compact = {
    v: namespace.schemaVersion,
    c: namespace.normalizeCode(invite.code),
    n: String(invite.name || '').trim().slice(0, 80),
    e: normalizeEmail(invite.hostEmail),
    t: String(invite.hostName || '').trim().slice(0, 80),
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

/**
 * 招待署名鍵を作る。秘密鍵は主催者だけが持つ非共有ファイルへ保存すること。
 * generation は世代番号。鍵を入れ替えると古い招待URLは検証に落ちる。
 */
export async function createInviteKey(namespace, {
  generation = 1, validityDays = 35, now = new Date(), cryptoObject = globalThis.crypto
} = {}) {
  const pair = await cryptoObject.subtle.generateKey(ALGORITHM, true, ['sign', 'verify']);
  const publicKeyJwk = canonicalPublicKey(await cryptoObject.subtle.exportKey('jwk', pair.publicKey));
  const privateKeyJwk = await cryptoObject.subtle.exportKey('jwk', pair.privateKey);
  const keyId = await namespace.inviteKeyId(publicKeyJwk, cryptoObject);
  const expiresAt = new Date(now.getTime() + Math.max(1, Number(validityDays) || 35) * 86400000).toISOString();
  return {
    algorithm: 'ECDSA-P256-SHA256',
    generation,
    keyId,
    publicKeyJwk,
    privateKeyJwk,
    expiresAt,
    createdAt: now.toISOString()
  };
}

/** 鍵として使える形がそろっていて、まだ期限内か。 */
export function inviteKeyUsable(security, now = new Date()) {
  return Boolean(
    security?.keyId && security?.publicKeyJwk?.x && security?.privateKeyJwk?.d
    && new Date(security.expiresAt).getTime() > now.getTime()
  );
}

/** 署名なし招待。互換読取り用に残すが、新規発行には使わない。 */
export function encodeInvite(namespace, invite) {
  return bytesToBase64Url(new TextEncoder().encode(JSON.stringify(compactPayload(namespace, invite))));
}

export async function encodeSignedInvite(namespace, invite, security, cryptoObject = globalThis.crypto) {
  if (!security?.privateKeyJwk?.d || !security?.publicKeyJwk?.x || !security?.keyId) {
    throw new InviteError('unsigned', 'invite signing key is not ready');
  }
  const payload = bytesToBase64Url(new TextEncoder().encode(JSON.stringify(compactPayload(namespace, invite, security))));
  const privateKey = await cryptoObject.subtle.importKey('jwk', security.privateKeyJwk, ALGORITHM, false, ['sign']);
  const signature = await cryptoObject.subtle.sign(SIGN_PARAMS, privateKey, new TextEncoder().encode(payload));
  return `${namespace.tokenPrefix}.${payload}.${bytesToBase64Url(new Uint8Array(signature))}`;
}

/**
 * 招待を復号し、署名・期限・必須項目を確かめる。
 * 署名の検証に使う公開鍵は招待自身に入っているため、これだけでは「本物の主催者か」は
 * 決まらない。取り込み側で keyId と generation を主催者のファイルと突き合わせること
 * （records.js の validateSharedRecord とアプリ側の照合を必ず通す）。
 */
export async function decodeInvite(namespace, encoded, {
  allowExpired = false, now = Date.now(), cryptoObject = globalThis.crypto
} = {}) {
  let raw;
  let signed = false;
  try {
    const parts = String(encoded || '').split('.');
    if (parts.length === 3 && parts[0] === namespace.tokenPrefix) {
      raw = JSON.parse(new TextDecoder().decode(base64UrlToBytes(parts[1])));
      const publicKey = await cryptoObject.subtle.importKey('jwk', canonicalPublicKey(raw.k), ALGORITHM, false, ['verify']);
      signed = await cryptoObject.subtle.verify(
        SIGN_PARAMS, publicKey, base64UrlToBytes(parts[2]), new TextEncoder().encode(parts[1])
      );
      if (!signed) throw new Error('invalid signature');
    } else {
      raw = JSON.parse(new TextDecoder().decode(base64UrlToBytes(encoded)));
    }
  } catch (error) {
    throw new InviteError('unreadable', 'invite token could not be read');
  }
  const invite = {
    version: Number(raw.v),
    code: namespace.normalizeCode(raw.c),
    name: String(raw.n || '').trim().slice(0, 80),
    hostEmail: normalizeEmail(raw.e),
    hostName: String(raw.t || '').trim().slice(0, 80),
    approvalRequired: raw.a !== false,
    acceptingMembers: raw.o !== false,
    signed,
    generation: Math.max(0, Number.parseInt(raw.g, 10) || 0),
    expiresAt: String(raw.x || ''),
    keyId: String(raw.i || ''),
    publicKeyJwk: raw.k ? canonicalPublicKey(raw.k) : null,
    token: String(encoded || '')
  };
  if (invite.version !== namespace.schemaVersion || invite.code.length < namespace.minCodeLength
    || !invite.name || !isEmail(invite.hostEmail)) {
    throw new InviteError('invalid', 'invite content is not valid');
  }
  if (signed && (!invite.keyId || !invite.expiresAt || (!allowExpired && new Date(invite.expiresAt).getTime() <= now))) {
    throw new InviteError('expired', 'invite token has expired');
  }
  invite.tenantId = await namespace.tenantId(invite.hostEmail, invite.code, cryptoObject);
  return invite;
}

/**
 * 参加者が持ち込んだ招待が、主催者の現在の鍵から出たものかを確かめる。
 * 取り込み時に必ず通す。落ちた場合は名簿へ入れない。
 */
export function matchesIssuedKey(invite, security, { tenantId = '' } = {}) {
  if (!invite?.signed) return { ok: false, reason: 'unsigned' };
  if (tenantId && invite.tenantId !== tenantId) return { ok: false, reason: 'tenant_mismatch' };
  if (!security?.keyId || invite.keyId !== security.keyId) return { ok: false, reason: 'key_mismatch' };
  if (invite.generation !== security.generation) return { ok: false, reason: 'generation_mismatch' };
  return { ok: true };
}

/** 記録の作成時刻が招待の有効期限内かを見る（期限切れURLの使い回しを弾く）。 */
export function createdWithinInvite(createdTime, invite) {
  const createdAt = new Date(createdTime || '').getTime();
  const expiresAt = new Date(invite?.expiresAt || '').getTime();
  return Number.isFinite(createdAt) && Number.isFinite(expiresAt) && createdAt <= expiresAt;
}

/** 招待URL。クエリではなくフラグメントへ載せる。 */
export function inviteUrl(baseUrl, token) {
  const url = new URL(baseUrl);
  url.search = '';
  url.hash = `join=${String(token || '')}`;
  return url.href;
}

/** URL・QRから読み取った文字列から招待トークンを取り出す。 */
export function inviteTokenFromUrl(value) {
  const hash = String(value || '').split('#')[1] || '';
  return hash.match(/(?:^|&)join=([^&]+)/)?.[1] || '';
}
