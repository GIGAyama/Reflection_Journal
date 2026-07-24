/**
 * ============================================================
 * ふりかえりジャーナル — テナント解決モジュール (マルチテナント対応)
 * ============================================================
 * スタンドアロン Webアプリを
 *   「次のユーザーとして実行: ウェブアプリケーションにアクセスしているユーザー」
 * でデプロイすると、UserProperties が Google アカウント単位で自動的に分離される。
 * このモジュールは「共通URLにログインするだけで各自の DB に接続される」ための
 * テナント（ユーザーごとのスプレッドシート）解決を担当する。
 *
 * プロパティの使い分け:
 *   - UserProperties   … ユーザー個別の設定（DB ID、APIキー、週間テーマ等）
 *   - ScriptProperties … 配布元が設定する共有設定（DBテンプレートID）と旧環境互換のみ
 */

// UserProperties: ユーザー個別の DB スプレッドシート ID
const UP_KEY_SPREADSHEET_ID = 'up_spreadsheetId';
// ScriptProperties: 旧バインド/旧デプロイ互換のグローバル DB ID
const SP_KEY_LEGACY_SPREADSHEET_ID = 'SPREADSHEET_ID';
// ScriptProperties: 配布元が設定する DB テンプレートのファイル ID
const SP_KEY_DB_TEMPLATE_ID = 'sp_dbTemplateId';

// ------------------------------------------------------------
// UserProperties への読み書き
// ------------------------------------------------------------

function getUserSpreadsheetId_() {
  try {
    return PropertiesService.getUserProperties().getProperty(UP_KEY_SPREADSHEET_ID) || '';
  } catch (e) {
    return '';
  }
}

function setUserSpreadsheetId_(id) {
  PropertiesService.getUserProperties().setProperty(UP_KEY_SPREADSHEET_ID, id);
}

function clearUserSpreadsheetId_() {
  PropertiesService.getUserProperties().deleteProperty(UP_KEY_SPREADSHEET_ID);
}

/**
 * DB スプレッドシート ID の解決。優先順位: ユーザー個別 → 旧グローバル。
 * どちらも無ければ空文字。
 */
function resolveSpreadsheetId_() {
  const userId = getUserSpreadsheetId_();
  if (userId) return userId;
  try {
    return PropertiesService.getScriptProperties().getProperty(SP_KEY_LEGACY_SPREADSHEET_ID) || '';
  } catch (e) {
    return '';
  }
}

/**
 * スプレッドシートの URL / 生ID のどちらを渡されても ID を抽出する。
 */
function extractSpreadsheetId_(input) {
  if (!input) return '';
  const str = String(input).trim();
  const m = str.match(/\/spreadsheets\/d\/([a-zA-Z0-9\-_]+)/);
  if (m) return m[1];
  if (/^[a-zA-Z0-9\-_]{20,}$/.test(str)) return str;
  return '';
}

// ------------------------------------------------------------
// ユーザー別設定アクセサ
//   読み取り: UserProperties 優先 → ScriptProperties フォールバック
//   書き込み: 常に UserProperties
// 旧環境（ScriptProperties に設定を保存していた環境）から移行しても設定が
// 失われず、保存し直すと自然にユーザー別へ移る。
// ------------------------------------------------------------

function tGetProp_(key) {
  try {
    const v = PropertiesService.getUserProperties().getProperty(key);
    if (v !== null && v !== '') return v;
  } catch (e) {}
  try {
    return PropertiesService.getScriptProperties().getProperty(key);
  } catch (e) {}
  return null;
}

function tSetProp_(key, value) {
  PropertiesService.getUserProperties().setProperty(key, value);
}

function tDeleteProp_(key) {
  PropertiesService.getUserProperties().deleteProperty(key);
}

// ------------------------------------------------------------
// スプレッドシート取得の一本化
// ------------------------------------------------------------

/**
 * 1. バインドされたスプレッドシートがあればそれ（旧バインド互換）
 * 2. Webアプリ文脈では resolveSpreadsheetId_() → openById
 * 3. どちらも無ければ throw（フロントはオンボーディングへ誘導）
 */
function getSs_() {
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (e) {}
  const id = resolveSpreadsheetId_();
  if (id) {
    try {
      return SpreadsheetApp.openById(id);
    } catch (e) {}
  }
  throw new Error('DB未設定。初期設定でデータベースを作成/紐付けしてください');
}

// ------------------------------------------------------------
// オンボーディング用 Web API（google.script.run から呼び出す）
// ------------------------------------------------------------

/**
 * テナント（DB 接続）状態を返す。
 * linked: false のときフロントはオンボーディング画面を表示する。
 */
function getTenantStatus() {
  try {
    let email = '';
    try {
      email = Session.getActiveUser().getEmail() || '';
    } catch (e) {}
    let templateConfigured = false;
    try {
      templateConfigured = !!PropertiesService.getScriptProperties().getProperty(SP_KEY_DB_TEMPLATE_ID);
    } catch (e) {}

    const base = {
      success: true,
      linked: false,
      spreadsheetId: '',
      spreadsheetName: '',
      email: email,
      canCreate: true,
      templateConfigured: templateConfigured
    };

    const id = resolveSpreadsheetId_();
    if (!id) return base;

    try {
      const ss = SpreadsheetApp.openById(id);
      base.linked = true;
      base.spreadsheetId = id;
      base.spreadsheetName = ss.getName();
      return base;
    } catch (e) {
      // ID はあるが開けない（削除済み・権限なし）→ 未接続としてオンボーディングへ
      base.spreadsheetId = id;
      return base;
    }
  } catch (e) {
    return { success: false, error: '状態の取得に失敗しました: ' + e.message };
  }
}

/**
 * 既存スプレッドシート（URL または ID）を自分の DB として紐付ける。
 * 必須シートが欠けていても紐付けは許可し warning を返す。
 */
function linkMyDatabase(input) {
  try {
    const id = extractSpreadsheetId_(input);
    if (!id) {
      return { success: false, error: 'スプレッドシートのURLまたはIDを正しく入力してください。' };
    }

    let ss;
    try {
      ss = SpreadsheetApp.openById(id);
      ss.getName(); // アクセスできるかの実チェック
    } catch (e) {
      return { success: false, error: 'スプレッドシートを開けませんでした。URL/IDが正しいか、あなたのアカウントに閲覧・編集権限があるか確認してください。' };
    }

    const required = [ROSTER_SHEET_NAME, JOURNAL_SHEET_NAME, THEME_SHEET_NAME];
    const missing = required.filter(function (n) { return !ss.getSheetByName(n); });

    setUserSpreadsheetId_(id);

    // 名簿に担任が1人もいない DB を紐付けた場合は、本人をこのテナントの管理者にする
    try {
      const rosterSheet = ss.getSheetByName(ROSTER_SHEET_NAME);
      let hasTeacher = false;
      if (rosterSheet && rosterSheet.getLastRow() >= 2) {
        const roles = rosterSheet.getRange(2, 1, rosterSheet.getLastRow() - 1, 1).getValues();
        hasTeacher = roles.some(function (r) { return r[0] === '担任'; });
      }
      if (!hasTeacher) {
        const email = Session.getActiveUser().getEmail();
        if (email) tSetProp_(PROP_ADMIN_EMAIL, email);
      }
    } catch (e) {}

    const result = { success: true, spreadsheetId: id, spreadsheetName: ss.getName() };
    if (missing.length > 0) {
      result.warning = '必須シート（' + missing.join('、') + '）が見つかりません。紐付けは完了しましたが、正しいDBスプレッドシートか確認してください。';
    }
    return result;
  } catch (e) {
    return { success: false, error: '紐付け中にエラーが発生しました: ' + e.message };
  }
}

/**
 * 自分の Drive に新しい DB スプレッドシートを作成して紐付ける。
 * 「アクセスユーザーとして実行」構成では本人が実行するため、作成物は本人所有になる。
 */
function createMyDatabase(name) {
  const lock = LockService.getUserLock();
  try {
    lock.waitLock(30000);

    const email = Session.getActiveUser().getEmail();

    // 多重実行対策: すでに有効な DB が紐付いていればそれを返す
    const existingId = getUserSpreadsheetId_();
    if (existingId) {
      try {
        const existing = SpreadsheetApp.openById(existingId);
        return { success: true, spreadsheetId: existingId, url: existing.getUrl(), method: 'existing' };
      } catch (e) {
        // 開けない ID が残っているだけなら作り直す
      }
    }

    const title = (name && String(name).trim()) || 'ふりかえりジャーナル_DB';
    let templateId = '';
    try {
      templateId = PropertiesService.getScriptProperties().getProperty(SP_KEY_DB_TEMPLATE_ID) || '';
    } catch (e) {}

    let ss;
    let method;
    if (templateId) {
      try {
        const copied = DriveApp.getFileById(templateId).makeCopy(title);
        ss = SpreadsheetApp.openById(copied.getId());
        method = 'template';
      } catch (e) {
        return {
          success: false,
          error: 'テンプレートの複製に失敗しました。テンプレートが「リンクを知っている全員が閲覧可」で共有されているか、配布元に確認してください。（詳細: ' + e.message + '）'
        };
      }
      // テンプレート由来でも必須シートが揃っているか保証する
      initializeNewDatabase_(ss);
    } else {
      ss = SpreadsheetApp.create(title);
      initializeNewDatabase_(ss);
      method = 'blank';
    }

    // 作成者をこのテナントの管理者（担任）として登録
    try {
      if (email) {
        tSetProp_(PROP_ADMIN_EMAIL, email);
        const rosterSheet = ss.getSheetByName(ROSTER_SHEET_NAME);
        if (rosterSheet) {
          let exists = false;
          if (rosterSheet.getLastRow() >= 2) {
            const emails = rosterSheet.getRange(2, 3, rosterSheet.getLastRow() - 1, 1).getValues();
            exists = emails.some(function (r) { return r[0] === email; });
          }
          if (!exists) rosterSheet.appendRow(['担任', '管理者', email]);
        }
      }
    } catch (e) {}

    setUserSpreadsheetId_(ss.getId());
    return { success: true, spreadsheetId: ss.getId(), url: ss.getUrl(), method: method };
  } catch (e) {
    return { success: false, error: 'データベースの作成中にエラーが発生しました: ' + e.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**
 * 紐付けのみ解除する（スプレッドシート自体は削除しない）。
 */
function unlinkMyDatabase() {
  try {
    clearUserSpreadsheetId_();
    return { success: true, message: 'データベースの紐付けを解除しました。スプレッドシート自体は削除されていません。' };
  } catch (e) {
    return { success: false, error: '解除中にエラーが発生しました: ' + e.message };
  }
}

/**
 * 新規スプレッドシートに必須シート（児童名簿 / ジャーナルデータ / テーマ設定）と
 * ヘッダーをプログラムで構築する。既存シートには手を付けない。
 */
function initializeNewDatabase_(ss) {
  createSheetIfNotExists_(ss, ROSTER_SHEET_NAME, ROSTER_HEADERS);
  createSheetIfNotExists_(ss, JOURNAL_SHEET_NAME, JOURNAL_HEADERS);
  createSheetIfNotExists_(ss, THEME_SHEET_NAME, THEME_HEADERS);
  const defaultSheet = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) ss.deleteSheet(defaultSheet);
  return ss;
}
