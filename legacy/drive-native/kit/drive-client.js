// ============================================================
// Driveネイティブ・キット — Drive REST クライアント
// ------------------------------------------------------------
// 運営者のサーバーを一切通さず、ログイン中の本人の権限でDrive APIを呼ぶ。
// アクセストークンはこのオブジェクトの中だけに置き、保存領域へ書かない。
//
// 書込みは drive.file（アプリが作ったファイルだけ）で足りる。
// 他人から共有された記録の「検索」だけが drive.readonly を必要とする。
// ============================================================

import { defineAppNamespace } from './namespace.js';

const API = 'https://www.googleapis.com/drive/v3';
const UPLOAD = 'https://www.googleapis.com/upload/drive/v3';
const FILE_FIELDS = 'id,name,mimeType,createdTime,modifiedTime,version,size,appProperties,owners(displayName,emailAddress)';

/** 既定の利用者向け文言。学校向けの言い回しなので、アプリ側で差し替えてよい。 */
export const DEFAULT_MESSAGES = {
  unauthorized: 'Googleへの接続期限が切れました。もう一度ログインしてください。',
  forbidden: '学校のGoogle Workspace設定により、この操作が許可されていません。管理者にアプリの許可設定を確認してください。',
  conflict: '別の端末でデータが更新されました。最新の内容を読み込み直してから、もう一度保存してください。',
  unknown: 'Google Driveとの通信に失敗しました。時間をおいてもう一度お試しください。'
};

export class DriveApiError extends Error {
  constructor(message, status, detail, code = 'unknown') {
    super(message);
    this.name = 'DriveApiError';
    this.status = status;
    this.detail = detail;
    this.code = code;
  }
}

export class DriveConflictError extends DriveApiError {
  constructor(message = DEFAULT_MESSAGES.conflict) {
    super(message, 409, 'version_conflict', 'conflict');
    this.name = 'DriveConflictError';
  }
}

function errorCode(status) {
  if (status === 401) return 'unauthorized';
  if (status === 403) return 'forbidden';
  if (status === 409 || status === 412) return 'conflict';
  return 'unknown';
}

export class KitDriveClient {
  /**
   * @param {object} options
   * @param {string} options.accessToken   GISトークンモデルで取得したアクセストークン。
   * @param {object} options.namespace     defineAppNamespace() の戻り値。
   * @param {Function} [options.fetchImpl] テスト用の差し替え。
   * @param {object} [options.messages]    利用者向け文言の差し替え。
   */
  constructor({ accessToken, namespace, fetchImpl, messages = {} } = {}) {
    this.accessToken = accessToken;
    this.namespace = namespace || defineAppNamespace({ appId: 'drive-native', propertyPrefix: 'dn' });
    this.messages = { ...DEFAULT_MESSAGES, ...messages };
    const fetcher = fetchImpl ?? globalThis.fetch;
    // ブラウザ標準のfetchはWindow/globalをレシーバーに要求する実装がある。
    // 包んでおくと、テストでの差し替えも同じ経路で効く。
    this.fetch = (...args) => Reflect.apply(fetcher, globalThis, args);
  }

  async request(url, options = {}) {
    const headers = new Headers(options.headers || {});
    headers.set('Authorization', `Bearer ${this.accessToken}`);
    const response = await this.fetch(url, { ...options, headers });
    if (!response.ok) {
      let detail = '';
      try { detail = (await response.json()).error?.message || ''; } catch (error) { detail = await response.text().catch(() => ''); }
      const code = errorCode(response.status);
      throw new DriveApiError(this.messages[code], response.status, detail, code);
    }
    if (response.status === 204) return null;
    const type = response.headers.get('content-type') || '';
    return type.includes('application/json') ? response.json() : response.blob();
  }

  /** 検索式に一致するファイルを、ページングを最後までたどって返す。 */
  async list(query) {
    const files = [];
    let pageToken = '';
    do {
      const params = new URLSearchParams({
        q: query,
        spaces: 'drive',
        pageSize: '100',
        orderBy: 'modifiedTime desc',
        fields: `nextPageToken,files(${FILE_FIELDS})`
      });
      if (pageToken) params.set('pageToken', pageToken);
      const result = await this.request(`${API}/files?${params}`);
      files.push(...(result.files || []));
      pageToken = result.nextPageToken || '';
    } while (pageToken);
    return files;
  }

  /** 種別・所有関係・テナントで絞って一覧する（検索式の組み立ては名前空間に任せる）。 */
  listByType(options) {
    return this.list(this.namespace.query(options));
  }

  async createJson(name, data, appProperties, parents = []) {
    const metadata = { name, mimeType: 'application/json', appProperties };
    if (parents.length) metadata.parents = parents;
    return this.createFile(metadata, new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' }));
  }

  async createFile(metadata, content) {
    const boundary = `dn_${crypto.randomUUID().replace(/-/g, '')}`;
    const body = new Blob([
      `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n`,
      JSON.stringify(metadata),
      `\r\n--${boundary}\r\nContent-Type: ${content.type || 'application/octet-stream'}\r\n\r\n`,
      content,
      `\r\n--${boundary}--`
    ], { type: `multipart/related; boundary=${boundary}` });
    return this.request(`${UPLOAD}/files?uploadType=multipart&fields=${FILE_FIELDS}`, { method: 'POST', body });
  }

  /**
   * 楽観的競合検知つきの更新。画面が読んだ版と食い違ったら書かずに落とす。
   * Drive APIには条件付き書込みが無いため、これはベストエフォートである。
   */
  async updateJson(fileId, data, { expectedVersion = '' } = {}) {
    if (expectedVersion) {
      const current = await this.getMetadata(fileId);
      if (String(current.version || '') !== String(expectedVersion)) throw new DriveConflictError(this.messages.conflict);
    }
    return this.request(`${UPLOAD}/files/${encodeURIComponent(fileId)}?uploadType=media&fields=id,modifiedTime,version,size`, {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json; charset=UTF-8' },
      body: JSON.stringify(data, null, 2)
    });
  }

  getMetadata(fileId) {
    return this.request(`${API}/files/${encodeURIComponent(fileId)}?fields=${FILE_FIELDS}`);
  }

  async getJson(fileId) {
    const result = await this.request(`${API}/files/${encodeURIComponent(fileId)}?alt=media`);
    if (result instanceof Blob) return JSON.parse(await result.text());
    return result;
  }

  getBlob(fileId) {
    return this.request(`${API}/files/${encodeURIComponent(fileId)}?alt=media`);
  }

  /** 相手へ直接共有する。通知メールは送らない（児童のメールボックスを埋めないため）。 */
  shareWithUser(fileId, email, role = 'reader') {
    const params = new URLSearchParams({ sendNotificationEmail: 'false', fields: 'id,type,role,emailAddress' });
    return this.request(`${API}/files/${encodeURIComponent(fileId)}/permissions?${params}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ type: 'user', role, emailAddress: email })
    });
  }

  createFolder(name, appProperties = {}) {
    return this.request(`${API}/files?fields=id,name,mimeType,createdTime,modifiedTime,version,appProperties`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name, mimeType: 'application/vnd.google-apps.folder', appProperties })
    });
  }
}
