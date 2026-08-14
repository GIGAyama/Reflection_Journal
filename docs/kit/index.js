// ============================================================
// Driveネイティブ・キット — 入口
// ------------------------------------------------------------
// 「運営者のサーバーやDBへ利用者データを集めず、GitHub Pages上のアプリが
// ログイン中の本人の権限でDrive APIを呼ぶ」構成を、アプリ間で使い回すための一式。
//
// アプリ側がやることは3つだけ。
//   1. createDriveNativeApp() へ自分のアプリの名前と用語を渡す
//   2. 記録の形（JSONの中身）と画面を書く
//   3. 取り込み時の照合規則を書く（キットの検証を通したうえで、アプリ固有の条件を足す）
//
// 使い方の詳細と、GAS版からの移し替え手順は docs/PORTING_FROM_GAS.md を見る。
// ============================================================

import { defineAppNamespace } from './namespace.js';
import { KitDriveClient } from './drive-client.js';
import { SessionPolicy, ScopeGrant } from './session.js';
import {
  createInviteKey, createdWithinInvite, decodeInvite, encodeInvite, encodeSignedInvite,
  inviteKeyUsable, inviteTokenFromUrl, inviteUrl, matchesIssuedKey
} from './invite.js';

export * from './namespace.js';
export * from './invite.js';
export * from './records.js';
export * from './session.js';
export { KitDriveClient, DriveApiError, DriveConflictError, DEFAULT_MESSAGES } from './drive-client.js';

export const DRIVE_FILE_SCOPE = 'https://www.googleapis.com/auth/drive.file';
export const DRIVE_READONLY_SCOPE = 'https://www.googleapis.com/auth/drive.readonly';
/** 最初に求める権限。自分のファイルの作成・更新・共有はこれだけで足りる。 */
export const BASE_SCOPES = `openid email profile ${DRIVE_FILE_SCOPE}`;
/**
 * 共有された記録の同期にだけ必要な制限付きスコープ。
 * Driveの共有だけでは相手側アプリのファイル認可にならないため、drive.file だけでは
 * 他人から共有された記録を検索できない。求める直前に用途を画面で説明すること。
 */
export const SHARED_READ_SCOPE = DRIVE_READONLY_SCOPE;

/**
 * アプリ1本ぶんの設定をまとめる。
 *
 * @param {object} config
 * @param {string} config.appId           他アプリと必ず異なるID（例 'daily-journal'）。
 * @param {string} config.propertyPrefix  appProperties の接頭辞（例 'rj'）。
 * @param {number} config.schemaVersion   保存形式の版。
 * @param {object} [config.terms]         { tenant: 'Class', member: 'Student' }。
 * @param {string} [config.clientId]      共通OAuthクライアントID。
 * @param {string[]} [config.allowedOrigins] 実行を許可する公開元。
 * @param {string[]} [config.allowedDomains] 利用を許可するWorkspaceドメイン。
 * @param {boolean} [config.persistSessionToken] 既定 false（トークンを保存しない）。
 * @param {object} [config.messages]      Driveエラーの利用者向け文言。
 */
export function createDriveNativeApp({
  appId,
  propertyPrefix,
  schemaVersion = 1,
  terms = {},
  codeLength = 8,
  clientId = '',
  allowedOrigins = [],
  allowedDomains = [],
  persistSessionToken = false,
  storageKey = `${propertyPrefix}_oauth_session_v1`,
  messages = {},
  entryUrl = ''
} = {}) {
  const namespace = defineAppNamespace({ appId, propertyPrefix, schemaVersion, terms, codeLength });
  const session = new SessionPolicy({
    clientId, storageKey, allowedOrigins, allowedDomains, persist: persistSessionToken
  });

  return {
    namespace,
    session,
    clientId,
    entryUrl,
    scopes: { base: BASE_SCOPES, sharedRead: SHARED_READ_SCOPE },
    newScopeGrant: (scopes = []) => new ScopeGrant(scopes),

    /** アクセストークン1本ぶんのDriveクライアント。トークンを差し替えたら作り直す。 */
    client(accessToken, fetchImpl) {
      return new KitDriveClient({ accessToken, namespace, fetchImpl, messages });
    },

    /** 招待まわりは名前空間を束ねた形で使う（tokenPrefixの取り違えを防ぐため）。 */
    invites: {
      createKey: (options) => createInviteKey(namespace, options),
      usable: inviteKeyUsable,
      encode: (invite) => encodeInvite(namespace, invite),
      sign: (invite, security, cryptoObject) => encodeSignedInvite(namespace, invite, security, cryptoObject),
      decode: (token, options) => decodeInvite(namespace, token, options),
      matchesIssuedKey,
      createdWithinInvite,
      url: (baseUrl, token) => inviteUrl(baseUrl || entryUrl, token),
      tokenFromUrl: inviteTokenFromUrl
    }
  };
}
