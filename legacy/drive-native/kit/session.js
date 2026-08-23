// ============================================================
// Driveネイティブ・キット — ログイン方針（画面を持たない部分）
// ------------------------------------------------------------
// GitHub Pagesの共有オリジンでは、同じドメインに数十本のアプリが同居する。
// 保存領域は隣のアプリからも読めるため、既定ではアクセストークンをどこにも
// 保存しない。保存を選ぶ場合も sessionStorage までにとどめる。
//
// 権限は段階的に求める。最初は自分のファイルを作る drive.file だけ、
// 共有記録の同期を始める直前に drive.readonly を説明してから求める。
// ============================================================

import { normalizeEmail } from './namespace.js';

const LOOPBACK_HOSTS = ['localhost', '127.0.0.1', '[::1]'];

/** トークン応答で実際に許可されたスコープ。要求と一致するとは限らない。 */
export function grantedScopesFrom(response) {
  return new Set(String(response?.scope || '').split(/\s+/).filter(Boolean));
}

/** 許可済みスコープの集合。granular consent で外された権限を取りこぼさないための入れ物。 */
export class ScopeGrant {
  constructor(scopes = []) {
    this.scopes = new Set(scopes);
  }

  /**
   * @param {object} response GISのトークン応答。
   * @param {string[]} [required] 応答に載らないことがある要求済みスコープ。
   * @param {object} [oauth2] google.accounts.oauth2（hasGrantedAllScopes の確認に使う）。
   */
  remember(response, required = [], oauth2 = null) {
    const granted = grantedScopesFrom(response);
    for (const scope of granted) this.scopes.add(scope);
    for (const scope of required) {
      if (granted.has(scope) || oauth2?.hasGrantedAllScopes?.(response, scope)) this.scopes.add(scope);
    }
    return this;
  }

  has(scope) {
    return this.scopes.has(scope);
  }

  list() {
    return [...this.scopes];
  }
}

export class SessionPolicy {
  /**
   * @param {object} options
   * @param {string} options.clientId          OAuthクライアントID。変わった保存分は捨てる。
   * @param {string} options.storageKey        保存キー。他アプリと衝突しない名前にする。
   * @param {string[]} [options.allowedOrigins] 実行を許可する公開元。空なら制限しない。
   * @param {string[]} [options.allowedDomains] 利用を許可するWorkspaceドメイン。空なら制限しない。
   * @param {boolean} [options.persist]        true のときだけ sessionStorage へ残す。
   * @param {object} [options.storage]         既定は sessionStorage。
   */
  constructor({
    clientId = '', storageKey = '', allowedOrigins = [], allowedDomains = [],
    persist = false, storage = null, loopbackHosts = LOOPBACK_HOSTS
  } = {}) {
    this.clientId = clientId;
    this.storageKey = storageKey;
    this.allowedOrigins = new Set((allowedOrigins || []).filter(Boolean));
    this.allowedDomains = (allowedDomains || []).map((value) => String(value).trim().toLowerCase()).filter(Boolean);
    this.persist = persist === true;
    this.loopbackHosts = loopbackHosts;
    this._storage = storage;
  }

  get storage() {
    if (this._storage !== null) return this._storage;
    try { return globalThis.sessionStorage; } catch (error) { return null; }
  }

  /** 設定した公開元以外へ資産をコピーされたときに動かさないための確認。 */
  originAllowed(origin, hostname = '') {
    if (!this.allowedOrigins.size) return true;
    return this.allowedOrigins.has(origin) || this.loopbackHosts.includes(hostname);
  }

  domainAllowed(email) {
    if (!this.allowedDomains.length) return true;
    return this.allowedDomains.includes(normalizeEmail(email).split('@')[1] || '');
  }

  save(session) {
    if (!this.persist || !this.storageKey) return false;
    if (!session?.accessToken || !session?.user?.email || !session?.expiresAt) return false;
    try {
      this.storage?.setItem(this.storageKey, JSON.stringify({
        clientId: this.clientId,
        accessToken: session.accessToken,
        expiresAt: session.expiresAt,
        scopes: session.scopes || [],
        user: session.user
      }));
      return true;
    } catch (error) { return false; }
  }

  /** 復元できないときは必ず消してから null を返す（古いトークンを残さない）。 */
  restore({ now = Date.now(), marginMs = 60_000 } = {}) {
    if (!this.persist || !this.storageKey) {
      this.clear();
      return null;
    }
    try {
      const saved = JSON.parse(this.storage?.getItem(this.storageKey) || 'null');
      const usable = saved && saved.clientId === this.clientId && saved.accessToken
        && saved.user?.email && Number(saved.expiresAt) >= now + marginMs
        && this.domainAllowed(saved.user.email);
      if (!usable) {
        this.clear();
        return null;
      }
      return {
        accessToken: saved.accessToken,
        expiresAt: Number(saved.expiresAt),
        scopes: Array.isArray(saved.scopes) ? saved.scopes : [],
        user: { email: normalizeEmail(saved.user.email), name: saved.user.name || saved.user.email }
      };
    } catch (error) {
      this.clear();
      return null;
    }
  }

  clear() {
    if (!this.storageKey) return;
    try { this.storage?.removeItem(this.storageKey); } catch (error) {}
  }
}

/** Google Identity Services を必要になった時点で読み込む（初期表示を遅らせないため）。 */
export function loadGoogleIdentity({
  documentRef = globalThis.document,
  windowRef = globalThis,
  src = 'https://accounts.google.com/gsi/client',
  errorMessage = 'Googleログインを読み込めませんでした。ネットワーク設定を確認してください。'
} = {}) {
  return new Promise((resolve, reject) => {
    if (windowRef.google?.accounts?.oauth2) return resolve(windowRef.google);
    const script = documentRef.createElement('script');
    script.src = src;
    script.async = true;
    script.onload = () => resolve(windowRef.google);
    script.onerror = () => reject(new Error(errorMessage));
    documentRef.head.appendChild(script);
  });
}

/** アクセストークンから本人のメールと表示名を得る。フロントの自己申告は使わない。 */
export async function fetchUserInfo(accessToken, fetchImpl = globalThis.fetch) {
  const response = await Reflect.apply(fetchImpl, globalThis, [
    'https://openidconnect.googleapis.com/v1/userinfo',
    { headers: { Authorization: `Bearer ${accessToken}` } }
  ]);
  if (!response.ok) throw new Error('Googleアカウント情報を確認できませんでした。');
  const profile = await response.json();
  return { email: normalizeEmail(profile.email), name: profile.name || profile.email };
}

/** 有効期限をミリ秒の絶対時刻へ直す（応答の expires_in は秒で、短いことがある）。 */
export function tokenExpiryFrom(response, now = Date.now()) {
  return now + Math.max(60, Number(response?.expires_in) || 3600) * 1000;
}
