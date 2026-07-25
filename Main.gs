/**
 * ============================================================
 * ふりかえりジャーナル — Main.gs
 * 設定・doGet ルーティング・共通ユーティリティ・診断 API
 * ============================================================
 *
 * ── アーキテクチャ概要（マルチテナント 3 層構成） ──────────────────
 * 同一プロジェクトから 2 本の Web アプリデプロイを発行して運用する。
 *
 *   デプロイ A（先生ポータル）: 実行 = ウェブアプリケーションにアクセスしているユーザー
 *     → クラス（テナント）作成が先生本人の権限で走り、DB シートは最初から先生所有になる。
 *       同じ実行内で addEditor(アプリアカウント) を呼び、共有まで自動化する。
 *   デプロイ B（児童アプリ）: 実行 = 自分（アプリアカウント）/ アクセス = 全員（匿名可）
 *     → すべての読み書きがアプリアカウント権限で走る。児童はシートへの権限を一切持たない。
 *       Session.getActiveUser() が使えないため、本人確認は ID トークン検証（Auth.gs）で行う。
 *
 * 入口は GitHub Pages のシェル（共通 URL 1 つ）。シェルが Google サインイン（GIS）を担当し、
 * ID トークンを iframe 内の本アプリへ postMessage で渡す。
 *
 * ⚠️ appsscript.json の webapp は A 用の値。B は必ずデプロイ画面（UI）で
 *    「自分として実行 / 全員」に設定すること（clasp で作ると A の値で作られてしまう）。
 */

const CONFIG = {
  APP_NAME: 'ふりかえりジャーナル',
  SCHEMA_VERSION: 2,
  LOCK_TIMEOUT_MS: 10000,
  REGISTRY_CACHE_SEC: 600,            // ScriptProperties の日次上限対策
  TOKEN_CACHE_SEC: 300,               // UrlFetch の日次上限対策
  CODE_ALPHABET: 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789', // 紛らわしい I,1,O,0 を除外
  CODE_LENGTH: 8
};

/**
 * 配布元設定（ScriptProperties）。レジストリ(tn_/own_)とこの sp_* 以外を
 * ScriptProperties に置いてはならない（1値9KB/全体500KB制限のため）。
 */
const PROP_KEYS = {
  APP_ACCOUNT: 'sp_appAccountEmail',  // 必須: アプリ運営用アカウント
  CLIENT_ID:   'sp_googleClientId',   // 必須: GIS 用クライアント ID（aud 検証に使用）
  SHELL_URL:   'sp_shellUrl',         // 必須: GitHub Pages シェルの URL
  TEMPLATE:    'sp_dbTemplateId'      // 任意: DB テンプレートのスプレッドシート ID
};

/**
 * 承認トリガー（GAS エディタから手動実行する）。
 *
 * デプロイ B は「自分として実行」のため、実行時の権限は「アプリアカウントが事前に
 * このスクリプトへ与えた承認」で決まる。oauthScopes を変更した場合や初回承認を
 * 飛ばした場合、児童側で「UrlFetchApp.fetch を呼び出す権限がありません」になる。
 *
 * 対処: アプリアカウントで GAS エディタを開き、この関数を選択して「実行」→ すべて許可。
 * 再デプロイは不要（既存デプロイに即反映される）。
 */
function authorizeApp() {
  const results = [];
  results.push('実行者: ' + Session.getEffectiveUser().getEmail());          // userinfo.email
  const res = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=check',
    { muteHttpExceptions: true });                                            // script.external_request
  results.push('UrlFetch(トークン検証): OK (HTTP ' + res.getResponseCode() + ' は正常です)');
  results.push('spreadsheets スコープ: ' + (ScriptApp.getOAuthToken() ? '承認済み' : '不明'));
  const summary = results.join(' / ');
  Logger.log(summary);
  return summary;
}

// ────────────────────────────────────────────────────────────────
// 設定アクセス
// ────────────────────────────────────────────────────────────────

function getSetting_(key, required) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  if (!v && required) {
    throw new Error('SERVER_ERROR: アプリの初期設定（' + key + '）がまだ行われていません。配布元（運営者）に連絡してください。');
  }
  return v || '';
}

function getShellUrl_() {
  let url = getSetting_(PROP_KEYS.SHELL_URL, false);
  if (url && url.slice(-1) !== '/') url += '/';
  return url;
}

function getShellOrigin_() {
  const m = getShellUrl_().match(/^(https?:\/\/[^\/]+)/);
  return m ? m[1] : '';
}

// ────────────────────────────────────────────────────────────────
// 共通ユーティリティ
// ────────────────────────────────────────────────────────────────

function sha256Hex_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(function (b) { return ((b + 256) % 256).toString(16).padStart(2, '0'); })
    .join('');
}

/** 書き込みの直列化。保持区間は最小に（appendRow 数回程度）保つこと */
function withScriptLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(CONFIG.LOCK_TIMEOUT_MS)) {
    throw new Error('LOCK_BUSY: 処理が混み合っています。数秒待ってもう一度お試しください');
  }
  try { return fn(); } finally { try { lock.releaseLock(); } catch (e) {} }
}

/** API レスポンスは必ずこの 2 つで返す。エラーは "CODE: 日本語メッセージ" 形式で throw する */
function jsonOk_(obj) {
  const out = obj || {};
  out.success = true;
  return JSON.stringify(out);
}

function jsonErr_(e) {
  const msg = (e && e.message) ? e.message : String(e);
  const m = msg.match(/^([A-Z_]+):\s*(.*)$/);
  return JSON.stringify({
    success: false,
    code: m ? m[1] : '',
    error: m ? (m[2] || msg) : msg
  });
}

// ────────────────────────────────────────────────────────────────
// ルーティング
// ────────────────────────────────────────────────────────────────

/**
 * デプロイの判別は exec URL の照合ではなく URL パラメータで行う:
 *   - 先生ポータル(A) はシェルが必ず ?portal=1 を付けて開く
 *   - 児童アプリ(B) はシェルが素の exec URL で開き、クラスコードは
 *     postMessage で渡す（パラメータ付き URL の iframe 読み込みを
 *     拒否する環境が存在するため）。旧方式の ?t= も受ける。
 */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.diag === '1') return doGetDiag_();

  const mode = p.portal === '1' ? 'owner' : 'member';

  const t = HtmlService.createTemplateFromFile('index');
  t.bootMode = mode;
  t.bootTenantCode = String(p.t || '').replace(/[^A-Z2-9]/gi, '').toUpperCase().slice(0, 16);
  t.bootShellUrl = getShellUrl_();
  t.bootShellOrigin = getShellOrigin_();
  return t.evaluate()
    .setTitle(CONFIG.APP_NAME)
    // GitHub Pages シェルが iframe 埋め込みするため必須。
    // GAS は frame-ancestors の限定ができないため ALLOWALL 一択。その代償として
    // ID トークン検証（Auth.gs）とシェル側 origin 検証を必須の防御線とする。
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

/**
 * 接続診断エンドポイント（?diag=1）。docs/diag.html が credentials:'omit' で fetch する。
 * 秘密情報（メールアドレス・ID・トークン）は一切返さない。
 *
 * この JSON が Cookie なしの fetch で読めた時点で「アクセスできるユーザー: 全員
 * （匿名アクセス可）」が確定する。ログイン必須設定だと Google が accounts.google.com へ
 * リダイレクトするため fetch 自体が失敗する。
 */
function doGetDiag_() {
  let effective = '', active = '';
  try { effective = String(Session.getEffectiveUser().getEmail() || '').toLowerCase(); } catch (e) {}
  try { active = String(Session.getActiveUser().getEmail() || '').toLowerCase(); } catch (e) {}
  const appAccount = String(getSetting_(PROP_KEYS.APP_ACCOUNT, false) || '').toLowerCase();
  return ContentService.createTextOutput(JSON.stringify({
    ok: true,
    app: CONFIG.APP_NAME,
    schemaVersion: CONFIG.SCHEMA_VERSION,
    // 実効ユーザーがアプリアカウントなら B（自分として実行）、それ以外なら A。
    // 注: アプリアカウント本人のブラウザから A を開いた場合も 'B' と出る
    deployKind: !effective ? 'unknown' : (appAccount && effective === appAccount) ? 'B' : 'A',
    anonymousAccess: !active,
    config: {
      appAccount: !!appAccount,
      clientId: !!getSetting_(PROP_KEYS.CLIENT_ID, false),
      shellUrl: getShellUrl_()   // 児童用 URL に含まれる公開情報のため露出可
    }
  })).setMimeType(ContentService.MimeType.JSON);
}
