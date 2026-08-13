import { classAppProperties, driveQueryValue } from './drive-core.js';

const API = 'https://www.googleapis.com/drive/v3';
const UPLOAD = 'https://www.googleapis.com/upload/drive/v3';

export class DriveApiError extends Error {
  constructor(message, status, detail) {
    super(message);
    this.name = 'DriveApiError';
    this.status = status;
    this.detail = detail;
  }
}

export class DriveClient {
  constructor(accessToken, fetchImpl) {
    this.accessToken = accessToken;
    const fetcher = fetchImpl ?? globalThis.fetch;
    // Browser-native fetch requires the Window/global receiver in some engines.
    // Wrapping it also keeps dependency injection available for automated tests.
    this.fetch = (...args) => Reflect.apply(fetcher, globalThis, args);
  }

  async request(url, options = {}) {
    const headers = new Headers(options.headers || {});
    headers.set('Authorization', `Bearer ${this.accessToken}`);
    const response = await this.fetch(url, { ...options, headers });
    if (!response.ok) {
      let detail = '';
      try { detail = (await response.json()).error?.message || ''; } catch (error) { detail = await response.text().catch(() => ''); }
      const friendly = response.status === 401
        ? 'Googleへの接続期限が切れました。もう一度ログインしてください。'
        : response.status === 403
          ? '学校のGoogle Workspace設定により、この操作が許可されていません。管理者にアプリの許可設定を確認してください。'
          : 'Google Driveとの通信に失敗しました。時間をおいてもう一度お試しください。';
      throw new DriveApiError(friendly, response.status, detail);
    }
    if (response.status === 204) return null;
    const type = response.headers.get('content-type') || '';
    return type.includes('application/json') ? response.json() : response.blob();
  }

  async list(query) {
    const files = [];
    let pageToken = '';
    do {
      const params = new URLSearchParams({
        q: query,
        spaces: 'drive',
        pageSize: '100',
        orderBy: 'modifiedTime desc',
        fields: 'nextPageToken,files(id,name,mimeType,createdTime,modifiedTime,appProperties,owners(displayName,emailAddress))'
      });
      if (pageToken) params.set('pageToken', pageToken);
      const result = await this.request(`${API}/files?${params}`);
      files.push(...(result.files || []));
      pageToken = result.nextPageToken || '';
    } while (pageToken);
    return files;
  }

  async createJson(name, data, appProperties) {
    return this.createFile({ name, mimeType: 'application/json', appProperties }, new Blob([
      JSON.stringify(data, null, 2)
    ], { type: 'application/json' }));
  }

  async createFile(metadata, content) {
    const boundary = `rj_${crypto.randomUUID().replace(/-/g, '')}`;
    const body = new Blob([
      `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n`,
      JSON.stringify(metadata),
      `\r\n--${boundary}\r\nContent-Type: ${content.type || 'application/octet-stream'}\r\n\r\n`,
      content,
      `\r\n--${boundary}--`
    ], { type: `multipart/related; boundary=${boundary}` });
    return this.request(`${UPLOAD}/files?uploadType=multipart&fields=id,name,mimeType,appProperties`, { method: 'POST', body });
  }

  async updateJson(fileId, data) {
    return this.request(`${UPLOAD}/files/${encodeURIComponent(fileId)}?uploadType=media&fields=id,modifiedTime`, {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json; charset=UTF-8' },
      body: JSON.stringify(data, null, 2)
    });
  }

  async getJson(fileId) {
    const result = await this.request(`${API}/files/${encodeURIComponent(fileId)}?alt=media`);
    if (result instanceof Blob) return JSON.parse(await result.text());
    return result;
  }

  async getBlob(fileId) {
    return this.request(`${API}/files/${encodeURIComponent(fileId)}?alt=media`);
  }

  async shareWithUser(fileId, email, role = 'reader') {
    const params = new URLSearchParams({ sendNotificationEmail: 'false', fields: 'id,type,role,emailAddress' });
    return this.request(`${API}/files/${encodeURIComponent(fileId)}/permissions?${params}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ type: 'user', role, emailAddress: email })
    });
  }

  listClasses() {
    return this.list("trashed = false and appProperties has { key='rjType' and value='class' }");
  }

  listOwnPortfolios(classId = '') {
    const classPart = classId ? ` and appProperties has { key='rjClassId' and value='${driveQueryValue(classId)}' }` : '';
    return this.list(`trashed = false and 'me' in owners and appProperties has { key='rjType' and value='portfolio' }${classPart}`);
  }

  listSharedPortfolios(classId) {
    return this.list(`trashed = false and sharedWithMe and appProperties has { key='rjType' and value='portfolio' } and appProperties has { key='rjClassId' and value='${driveQueryValue(classId)}' }`);
  }

  listSharedChannels(classId = '') {
    const classPart = classId ? ` and appProperties has { key='rjClassId' and value='${driveQueryValue(classId)}' }` : '';
    return this.list(`trashed = false and sharedWithMe and appProperties has { key='rjType' and value='channel' }${classPart}`);
  }

  createClass(record) {
    return this.createJson(`ふりかえりジャーナル_クラス_${record.className}.rj-class.json`, record, classAppProperties(record.classId, 'class'));
  }

  createPortfolio(record) {
    return this.createJson(`ふりかえりジャーナル_${record.class.name}_${record.student.name}.rj.json`, record, classAppProperties(record.class.id, 'portfolio'));
  }

  createChannel(record, studentHash) {
    return this.createJson(
      `ふりかえりジャーナル_おへんじ_${record.class.name}_${record.student.name}.rj-channel.json`,
      record,
      classAppProperties(record.class.id, 'channel', { rjStudent: studentHash })
    );
  }

  createJournalImage({ classId, journalId, studentName, file }) {
    const extension = String(file.name || '').match(/\.[a-z0-9]+$/i)?.[0] || '.jpg';
    return this.createFile({
      name: `ふりかえり_${studentName}_${journalId}${extension}`,
      mimeType: file.type || 'image/jpeg',
      appProperties: { ...classAppProperties(classId, 'journal-image'), rjJournalId: journalId }
    }, file);
  }
}
