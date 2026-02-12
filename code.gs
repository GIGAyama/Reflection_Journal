/**
 * ============================================================
 * ふり返りジャーナル — サーバーサイド スクリプト（スタンドアロン版）
 * ============================================================
 * スタンドアロン GAS プロジェクトとして動作し、初回デプロイ時に
 * スプレッドシート・Drive フォルダを自動作成する Zero-Config 設計。
 *
 * ファイル構成:
 *   コード.gs  — 本ファイル（サーバーサイドロジック）
 *   index.html — SPA 構成の HTML テンプレート
 *   css.html   — スタイル定義
 *   js.html    — クライアントサイド JavaScript
 * ============================================================
 */

// =============================================================
// ■ 1. 定数定義
// =============================================================

/** シート名 */
const ROSTER_SHEET_NAME  = '児童名簿';
const JOURNAL_SHEET_NAME = 'ジャーナルデータ';
const THEME_SHEET_NAME   = 'テーマ設定';

/** Gemini API エンドポイント（モデル名を定数化） */
const GEMINI_MODEL = 'gemini-2.5-flash';
const API_ENDPOINT_V1      = `https://generativelanguage.googleapis.com/v1/models/${GEMINI_MODEL}:generateContent?key=`;
const API_ENDPOINT_V1_BETA = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=`;

/** スクリプトプロパティのキー名 */
const PROP_SPREADSHEET_ID  = 'SPREADSHEET_ID';
const PROP_IMAGE_FOLDER_ID = 'IMAGE_FOLDER_ID';
const PROP_ADMIN_EMAIL     = 'ADMIN_EMAIL';
const PROP_INITIALIZED     = 'APP_INITIALIZED';

/** ジャーナルデータシートのヘッダー定義 */
const JOURNAL_HEADERS = [
  'journalId', 'timestamp', 'email', 'theme', 'content',
  'imageFileId', 'emotion', 'teacherComment', 'highlights',
  'teacherStamp', 'status', 'pastComment', 'deletedAt'
];

/** 児童名簿シートのヘッダー定義 */
const ROSTER_HEADERS = ['役割', '氏名', 'メールアドレス'];

/** テーマ設定シートのヘッダー定義 */
const THEME_HEADERS = ['日付', 'テーマ'];


// =============================================================
// ■ 2. HTML インクルード & Webアプリ基本処理
// =============================================================

/**
 * HTML テンプレート内で CSS / JS ファイルを読み込むための関数
 * @param {string} filename - 読み込むファイル名（拡張子なし）
 * @returns {string} ファイルの内容
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Web アプリにアクセスがあったときのエントリーポイント
 * 初回アクセス時に自動セットアップを実行する（Zero-Config）
 * @param {object} e - イベントオブジェクト
 * @returns {HtmlOutput} Web ページ
 */
function doGet(e) {
  // ── Zero-Config: 初期化チェック ──
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty(PROP_INITIALIZED)) {
    initializeApp_(props);
  }

  // ── ユーザー認証 ──
  const userEmail = Session.getActiveUser().getEmail();

  // 管理者メールアドレスをチェック（管理者は常にアクセス可能）
  const adminEmail = props.getProperty(PROP_ADMIN_EMAIL);
  const isAdmin = (userEmail === adminEmail);

  const userData = getUserData_(userEmail, isAdmin);

  if (!userData) {
    return HtmlService.createHtmlOutput(
      '<div style="text-align:center;margin-top:80px;font-family:sans-serif;">' +
      '<h1>⚠️ アクセス権がありません</h1>' +
      '<p>名簿に登録されているアカウントでログインしてください。</p></div>'
    );
  }

  // ── HTML テンプレート生成 ──
  const tmpl = HtmlService.createTemplateFromFile('index');
  tmpl.user       = userData;
  tmpl.todayTheme = getTodayTheme();
  tmpl.isAdmin    = isAdmin;

  if (userData.role === '担任') {
    tmpl.classRoster = getClassRoster();
    tmpl.journals    = getJournalsForTeacher();
  } else {
    tmpl.classRoster = [];
    tmpl.journals    = getJournalsForStudent(userEmail);
  }

  return tmpl.evaluate()
    .setTitle('ふり返りジャーナル')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl('https://drive.google.com/uc?id=1VVrYSjBa-i_X7SjetvoxgX_ZPAU385rq&.png');
}

/**
 * Web アプリ自身の URL を取得する
 * @returns {string} URL
 */
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}


// =============================================================
// ■ 3. スプレッドシート取得ヘルパー（スタンドアロン対応）
// =============================================================

/**
 * スクリプトプロパティに保存された ID からスプレッドシートを取得する
 * スタンドアロンスクリプトでは getActiveSpreadsheet() が使えないため、
 * この関数を経由して全てのシート操作を行う。
 *
 * @returns {Spreadsheet} スプレッドシートオブジェクト
 * @private
 */
function getSpreadsheet_() {
  const ssId = PropertiesService.getScriptProperties().getProperty(PROP_SPREADSHEET_ID);
  if (!ssId) {
    throw new Error('データベースが初期化されていません。管理者に連絡してください。');
  }
  return SpreadsheetApp.openById(ssId);
}


// =============================================================
// ■ 4. Zero-Config 自動セットアップ
// =============================================================

/**
 * 初回起動時にスプレッドシート・Drive フォルダを新規作成する
 * スタンドアロンスクリプトのため、SpreadsheetApp.create() で
 * 新しいスプレッドシートを自動生成し、IDをプロパティに保存する。
 *
 * @param {PropertiesService.Properties} props - スクリプトプロパティ
 * @private
 */
function initializeApp_(props) {
  // ── 管理者メールアドレスを保存（初回実行者＝管理者） ──
  const adminEmail = Session.getActiveUser().getEmail();
  props.setProperty(PROP_ADMIN_EMAIL, adminEmail);

  // ── スプレッドシートを新規作成 ──
  const ss = SpreadsheetApp.create('ふり返りジャーナル_DB');
  props.setProperty(PROP_SPREADSHEET_ID, ss.getId());
  Logger.log('📊 スプレッドシートを作成しました: ' + ss.getUrl());

  // ── 各シートを作成 ──
  createSheetIfNotExists_(ss, ROSTER_SHEET_NAME, ROSTER_HEADERS);
  createSheetIfNotExists_(ss, JOURNAL_SHEET_NAME, JOURNAL_HEADERS);
  createSheetIfNotExists_(ss, THEME_SHEET_NAME, THEME_HEADERS);

  // ── デフォルトの「シート1」を削除 ──
  const defaultSheet = ss.getSheetByName('シート1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  // ── 管理者を名簿に自動登録 ──
  const rosterSheet = ss.getSheetByName(ROSTER_SHEET_NAME);
  rosterSheet.appendRow(['担任', '管理者', adminEmail]);

  // ── 画像保存用フォルダを作成 ──
  const folder = DriveApp.createFolder('ふり返りジャーナル_画像');
  props.setProperty(PROP_IMAGE_FOLDER_ID, folder.getId());

  // ── 初期化完了フラグ ──
  props.setProperty(PROP_INITIALIZED, 'true');

  Logger.log('✅ 初期セットアップが完了しました（スタンドアロン）');
  Logger.log('📁 スプレッドシートURL: ' + ss.getUrl());
}

/**
 * 指定された名前のシートが存在しなければ作成し、整形する
 * @param {Spreadsheet} ss - スプレッドシート
 * @param {string} sheetName - シート名
 * @param {string[]} headers - ヘッダー行
 * @private
 */
function createSheetIfNotExists_(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) return; // 既に存在する場合はスキップ

  sheet = ss.insertSheet(sheetName);

  // ── ヘッダー行を書き込み ──
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ── ヘッダー行のスタイル設定 ──
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285F4');
  headerRange.setFontColor('#FFFFFF');

  // ── ヘッダー行を固定（フリーズ） ──
  sheet.setFrozenRows(1);

  // ── 交互の背景色を設定 ──
  sheet.getRange(1, 1, sheet.getMaxRows(), headers.length)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);

  // ── 列幅を自動調整 ──
  headers.forEach((_, i) => sheet.autoResizeColumn(i + 1));

  Logger.log(`📄 シート「${sheetName}」を作成しました`);
}


// =============================================================
// ■ 5. ユーザー認証と名簿データ処理
// =============================================================

/**
 * メールアドレスを元にユーザー情報を検索する
 * @param {string} email - 検索するメールアドレス
 * @param {boolean} isAdmin - 管理者かどうか
 * @returns {object|null} ユーザー情報
 * @private
 */
function getUserData_(email, isAdmin) {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(ROSTER_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    // 管理者で名簿が空なら、管理者情報を直接返す
    if (isAdmin) {
      return { role: '担任', name: '管理者', email: email };
    }
    return null;
  }

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();

  for (let i = 0; i < data.length; i++) {
    if (data[i][2] === email) {
      const role = (data[i][0] === '担任') ? '担任' : '児童';
      return {
        role: role,
        name: data[i][1],
        email: data[i][2]
      };
    }
  }

  // 名簿に未登録でも管理者ならアクセス許可
  if (isAdmin) {
    return { role: '担任', name: '管理者', email: email };
  }

  return null;
}

/**
 * クライアントから呼ばれるユーザーデータ取得（公開用）
 * @param {string} email
 * @returns {object|null}
 */
function getUserData(email) {
  const adminEmail = PropertiesService.getScriptProperties().getProperty(PROP_ADMIN_EMAIL);
  return getUserData_(email, email === adminEmail);
}

/**
 * クラスの児童一覧を取得する（担任行を除く）
 * @returns {object[]} 児童リスト
 */
function getClassRoster() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(ROSTER_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  return data
    .filter(row => row[0] !== '担任' && row[2])
    .map(row => ({ name: row[1], email: row[2] }));
}


// =============================================================
// ■ 6. ジャーナルデータの読み書き（CRUD）
// =============================================================

/**
 * ジャーナルを保存する（LockService で排他制御）
 * @param {object} journalData - ジャーナルデータ
 * @returns {object} 処理結果
 */
function saveJournal(journalData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(JOURNAL_SHEET_NAME);
    const newRow = [
      Utilities.getUuid(),               // journalId
      new Date(),                        // timestamp
      journalData.email,                 // email
      journalData.theme,                 // theme
      journalData.content,               // content
      journalData.imageFileId || '',     // imageFileId
      journalData.emotion || '',         // emotion
      '',                                // teacherComment
      '[]',                              // highlights
      '',                                // teacherStamp
      '未返却',                           // status
      '',                                // pastComment
      ''                                 // deletedAt（論理削除用）
    ];
    sheet.appendRow(newRow);

    return { success: true, message: 'ジャーナルを提出しました！' };
  } catch (e) {
    Logger.log('saveJournal エラー: ' + e.message);
    return { success: false, message: '保存中にエラーが発生しました：' + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 特定の児童のジャーナルを取得する（論理削除を除外）
 * @param {string} email - 児童のメールアドレス
 * @returns {object[]} ジャーナルリスト
 */
function getJournalsForStudent(email) {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(JOURNAL_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const emailIdx     = headers.indexOf('email');
  const deletedAtIdx = headers.indexOf('deletedAt');

  return data
    .filter(row => row[emailIdx] === email && !row[deletedAtIdx])
    .map(row => rowToJournalObject_(headers, row));
}

/**
 * 教員用に全児童のジャーナルを取得する（論理削除を除外）
 * @returns {object[]} ジャーナルリスト
 */
function getJournalsForTeacher() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(JOURNAL_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const deletedAtIdx = headers.indexOf('deletedAt');
  const emailIdx     = headers.indexOf('email');

  // 名簿から名前マップを作成
  const roster  = getClassRoster();
  const nameMap = roster.reduce((map, u) => {
    map[u.email] = u.name;
    return map;
  }, {});

  return data
    .filter(row => !row[deletedAtIdx])
    .map(row => {
      const journal = rowToJournalObject_(headers, row);
      journal.studentName = nameMap[row[emailIdx]] || '不明';
      return journal;
    });
}

/**
 * スプレッドシートの1行をジャーナルオブジェクトに変換する
 * @param {string[]} headers - ヘッダー配列
 * @param {Array} row - 行データ
 * @returns {object} ジャーナルオブジェクト
 * @private
 */
function rowToJournalObject_(headers, row) {
  const journal = {};
  headers.forEach((h, i) => { journal[h] = row[i]; });

  // 日付フォーマット
  if (journal.timestamp instanceof Date) {
    journal.date = Utilities.formatDate(journal.timestamp, 'JST', 'yyyy/MM/dd');
  }

  // 画像URL生成
  journal.imageUrl = journal.imageFileId
    ? `https://lh3.googleusercontent.com/d/${journal.imageFileId}`
    : '';

  return journal;
}

/**
 * フィードバック（コメント・ハイライト）を保存し、ステータスを返却済みにする
 * @param {object} feedbackData - フィードバックデータ
 * @returns {object} 処理結果
 */
function saveFeedback(feedbackData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(JOURNAL_SHEET_NAME);
    const data    = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIdx   = headers.indexOf('journalId');

    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] === feedbackData.journalId) {
        const rowNum = i + 1;
        sheet.getRange(rowNum, headers.indexOf('teacherComment') + 1).setValue(feedbackData.comment);
        sheet.getRange(rowNum, headers.indexOf('highlights') + 1).setValue(feedbackData.highlights || '[]');
        sheet.getRange(rowNum, headers.indexOf('status') + 1).setValue('返却済み');
        return { success: true, message: 'フィードバックを保存しました！' };
      }
    }
    return { success: false, message: '該当するジャーナルが見つかりませんでした。' };
  } catch (e) {
    Logger.log('saveFeedback エラー: ' + e.message);
    return { success: false, message: 'エラーが発生しました：' + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 返却済みジャーナルのステータスを「未返却」に戻す
 * @param {string} journalId
 * @returns {object} 処理結果
 */
function revertJournalStatus(journalId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(JOURNAL_SHEET_NAME);
    const data    = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIdx     = headers.indexOf('journalId');
    const statusIdx = headers.indexOf('status');

    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] === journalId) {
        sheet.getRange(i + 1, statusIdx + 1).setValue('未返却');
        return { success: true };
      }
    }
    return { success: false, message: '該当するジャーナルが見つかりませんでした。' };
  } catch (e) {
    Logger.log('revertJournalStatus エラー: ' + e.message);
    return { success: false, message: 'ステータスの更新中にエラーが発生しました：' + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 過去の自分へのコメントを保存する
 * @param {string} journalId
 * @param {string} comment
 * @returns {object} 処理結果
 */
function addPastComment(journalId, comment) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(JOURNAL_SHEET_NAME);
    const data    = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIdx      = headers.indexOf('journalId');
    const commentIdx = headers.indexOf('pastComment');

    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] === journalId) {
        sheet.getRange(i + 1, commentIdx + 1).setValue(comment);
        return { success: true, message: '過去の自分へのコメントを保存しました！' };
      }
    }
    return { success: false, message: '該当するジャーナルが見つかりませんでした。' };
  } catch (e) {
    Logger.log('addPastComment エラー: ' + e.message);
    return { success: false, message: 'エラーが発生しました：' + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ジャーナルを論理削除する（deletedAt にタイムスタンプを記録）
 * @param {string} journalId
 * @returns {object} 処理結果
 */
function deleteJournal(journalId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(JOURNAL_SHEET_NAME);
    const data    = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIdx        = headers.indexOf('journalId');
    const deletedAtIdx = headers.indexOf('deletedAt');

    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] === journalId) {
        sheet.getRange(i + 1, deletedAtIdx + 1).setValue(new Date());
        return { success: true, message: 'ジャーナルを削除しました。' };
      }
    }
    return { success: false, message: '該当するジャーナルが見つかりませんでした。' };
  } catch (e) {
    Logger.log('deleteJournal エラー: ' + e.message);
    return { success: false, message: '削除中にエラーが発生しました：' + e.message };
  } finally {
    lock.releaseLock();
  }
}


// =============================================================
// ■ 7. Google Drive 連携（画像アップロード）
// =============================================================

/**
 * 画像を Google Drive にアップロードする
 * @param {object} fileData - { fileName, mimeType, data (base64) }
 * @returns {object} 処理結果
 */
function uploadImage(fileData) {
  try {
    const folderId = PropertiesService.getScriptProperties().getProperty(PROP_IMAGE_FOLDER_ID);
    if (!folderId) {
      return { success: false, message: '画像保存用フォルダが設定されていません。' };
    }

    const folder  = DriveApp.getFolderById(folderId);
    const decoded = Utilities.base64Decode(fileData.data, Utilities.Charset.UTF_8);
    const blob    = Utilities.newBlob(decoded, fileData.mimeType, fileData.fileName);
    const file    = folder.createFile(blob);

    // リンクを知っている人全員が閲覧可能に設定
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return { success: true, fileId: file.getId(), fileName: file.getName() };
  } catch (e) {
    Logger.log('uploadImage エラー: ' + e.message);
    return { success: false, message: '画像のアップロードに失敗しました：' + e.message };
  }
}


// =============================================================
// ■ 8. 「今日のテーマ」管理
// =============================================================

/**
 * 今日のテーマを取得する
 * 週間スケジュールが設定されている場合はそちらを優先。
 * 個別設定 > 曜日スケジュール > デフォルトメッセージの順。
 * @returns {string} テーマ文字列
 */
function getTodayTheme() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(THEME_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    // テーマシートがなくても曜日スケジュールを確認
    return getWeeklyThemeForToday_() || 'Today, let\'s reflect on our day.';
  }

  const todayStr = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd');
  const data = sheet.getDataRange().getValues();

  // 下から探索（新しい日付が優先）
  for (let i = data.length - 1; i > 0; i--) {
    const rowDate = data[i][0];
    if (rowDate instanceof Date && Utilities.formatDate(rowDate, 'JST', 'yyyy/MM/dd') === todayStr) {
      return data[i][1];
    }
  }

  // 個別設定がなければ曜日スケジュールを確認
  return getWeeklyThemeForToday_() || '今日のふり返りをしましょう';
}

/**
 * 今日のテーマを設定する
 * @param {string} theme
 * @returns {object} 処理結果
 */
function setTodayTheme(theme) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(THEME_SHEET_NAME);
    sheet.appendRow([new Date(), theme]);
    return { success: true, message: '今日のテーマを設定しました！' };
  } catch (e) {
    Logger.log('setTodayTheme エラー: ' + e.message);
    return { success: false, message: 'Error: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 曜日別テーマスケジュールを取得する
 * @returns {object} { success, data: { mon, tue, wed, thu, fri } }
 */
function getWeeklyThemes() {
  try {
    const props = PropertiesService.getScriptProperties();
    const json = props.getProperty('WEEKLY_THEMES');
    const themes = json ? JSON.parse(json) : { mon: '', tue: '', wed: '', thu: '', fri: '' };
    return { success: true, data: themes };
  } catch (e) {
    Logger.log('getWeeklyThemes エラー: ' + e.message);
    return { success: false, message: 'Error: ' + e.message };
  }
}

/**
 * 曜日別テーマスケジュールを保存する
 * @param {object} themes - { mon, tue, wed, thu, fri }
 * @returns {object} 処理結果
 */
function saveWeeklyThemes(themes) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('WEEKLY_THEMES', JSON.stringify(themes));
    return { success: true, message: '週間テーマを保存しました！' };
  } catch (e) {
    Logger.log('saveWeeklyThemes エラー: ' + e.message);
    return { success: false, message: 'Error: ' + e.message };
  }
}

/**
 * 今日の曜日に対応するスケジュールテーマを返す（内部ヘルパー）
 * @returns {string|null} テーマ文字列またはnull
 * @private
 */
function getWeeklyThemeForToday_() {
  const props = PropertiesService.getScriptProperties();
  const json = props.getProperty('WEEKLY_THEMES');
  if (!json) return null;

  const themes = JSON.parse(json);
  const dayKeys = ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat'];
  const todayKey = dayKeys[new Date().getDay()];
  return themes[todayKey] || null;
}

/**
 * スタンプのみでジャーナルを即時返却する（クイック返却）
 * @param {string} journalId
 * @param {string} stamp - 絵文字スタンプ
 * @returns {object} 処理結果
 */
function quickReturn(journalId, stamp) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(JOURNAL_SHEET_NAME);
    const data    = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIdx   = headers.indexOf('journalId');

    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] === journalId) {
        const rowNum = i + 1;
        sheet.getRange(rowNum, headers.indexOf('teacherStamp') + 1).setValue(stamp);
        sheet.getRange(rowNum, headers.indexOf('status') + 1).setValue('返却済み');
        return { success: true, message: 'スタンプで返却しました！' };
      }
    }
    return { success: false, message: '該当するジャーナルが見つかりませんでした。' };
  } catch (e) {
    Logger.log('quickReturn エラー: ' + e.message);
    return { success: false, message: 'Error: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 未返却ジャーナル（コメント付き）を一括返却する
 * @returns {object} 処理結果（返却件数）
 */
function batchReturnAll() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(JOURNAL_SHEET_NAME);
    const data    = sheet.getDataRange().getValues();
    const headers = data[0];
    const statusIdx    = headers.indexOf('status');
    const commentIdx   = headers.indexOf('teacherComment');
    const deletedIdx   = headers.indexOf('deletedAt');

    let count = 0;
    for (let i = 1; i < data.length; i++) {
      // 未返却 かつ コメントあり かつ 論理削除されていない
      if (data[i][statusIdx] === '未返却' && data[i][commentIdx] && !data[i][deletedIdx]) {
        sheet.getRange(i + 1, statusIdx + 1).setValue('返却済み');
        count++;
      }
    }

    return {
      success: true,
      count: count,
      message: count > 0
        ? `${count}件のジャーナルを一括返却しました！`
        : '返却対象のジャーナルがありません（コメント付きの未返却がありません）。'
    };
  } catch (e) {
    Logger.log('batchReturnAll エラー: ' + e.message);
    return { success: false, message: 'Error: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}


// =============================================================
// ■ 9. AI（Gemini API）連携
// =============================================================

/**
 * 未返却ジャーナルにシンプルなAIコメント案を一括生成する
 * @returns {object} 処理結果
 */
function generateAiSimpleCommentsForAll() {
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(JOURNAL_SHEET_NAME);
    const data    = sheet.getDataRange().getValues();
    const headers = data.shift();

    const contentIdx = headers.indexOf('content');
    const statusIdx  = headers.indexOf('status');
    const commentIdx = headers.indexOf('teacherComment');
    const deletedIdx = headers.indexOf('deletedAt');

    let successCount = 0;

    data.forEach((row, index) => {
      if (row[statusIdx] === '未返却' && row[contentIdx] && !row[deletedIdx]) {
        const comment = callGeminiApiForSimpleComment_(row[contentIdx]);
        sheet.getRange(index + 2, commentIdx + 1).setValue(comment);
        successCount++;
      }
    });

    return { success: true, message: `AIコメント案（シンプル）の生成が完了しました。（${successCount}件）` };
  } catch (e) {
    Logger.log('generateAiSimpleCommentsForAll エラー: ' + e.message);
    return { success: false, message: 'AI処理中にエラーが発生しました：' + e.message };
  }
}

/**
 * 未返却ジャーナルに高度なAIフィードバック案を一括生成する
 * @returns {object} 処理結果
 */
function generateAiFullFeedbackForAll() {
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(JOURNAL_SHEET_NAME);
    const data    = sheet.getDataRange().getValues();
    const headers = data.shift();

    const idIdx        = headers.indexOf('journalId');
    const contentIdx   = headers.indexOf('content');
    const statusIdx    = headers.indexOf('status');
    const commentIdx   = headers.indexOf('teacherComment');
    const highlightIdx = headers.indexOf('highlights');
    const deletedIdx   = headers.indexOf('deletedAt');

    let successCount = 0;
    let errorCount   = 0;

    data.forEach((row, index) => {
      if (row[statusIdx] === '未返却' && row[contentIdx] && !row[deletedIdx]) {
        const aiResponse = callGeminiApiForFullFeedback_(row[contentIdx]);

        if (aiResponse && aiResponse.success) {
          try {
            const feedback = JSON.parse(aiResponse.jsonText);

            if (feedback.comment) {
              sheet.getRange(index + 2, commentIdx + 1).setValue(feedback.comment);
            }

            if (feedback.highlights && feedback.highlights.length > 0) {
              const content = row[contentIdx];
              const hlToSave = [];

              feedback.highlights.forEach(h => {
                const startIndex = content.indexOf(h.textToHighlight);
                if (startIndex !== -1) {
                  hlToSave.push({
                    id: 'hl-' + Date.now() + Math.random(),
                    text: h.textToHighlight,
                    comment: h.suggestedComment || '',
                    stamp: h.suggestedStamp || '',
                    startOffset: startIndex,
                    endOffset: startIndex + h.textToHighlight.length
                  });
                }
              });

              if (hlToSave.length > 0) {
                sheet.getRange(index + 2, highlightIdx + 1).setValue(JSON.stringify(hlToSave));
              }
            }
            successCount++;
          } catch (parseErr) {
            Logger.log(`AI応答の解析エラー: JournalID: ${row[idIdx]}, Error: ${parseErr.message}`);
            errorCount++;
          }
        } else {
          Logger.log(`AI呼び出し失敗: JournalID: ${row[idIdx]}`);
          errorCount++;
        }
      }
    });

    return {
      success: true,
      message: `AIフィードバック案（高度）の生成完了。（成功: ${successCount}件, 失敗: ${errorCount}件）`
    };
  } catch (e) {
    Logger.log('generateAiFullFeedbackForAll エラー: ' + e.message);
    return { success: false, message: 'AI処理中にエラーが発生しました：' + e.message };
  }
}

/**
 * Gemini API を呼び出してシンプルなコメントを生成する
 * @param {string} journalContent - ジャーナル本文
 * @returns {string} 生成されたコメント
 * @private
 */
function callGeminiApiForSimpleComment_(journalContent) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) return 'エラー: APIキーが設定されていません。';

  const prompt = 'あなたは児童の小さな頑張りやユニークな視点を見つけて具体的に褒めるのが得意な、経験豊富な小学校の先生です。' +
    '以下の記述を読み、児童が「努力した点」「工夫した点」などを引用しつつ、' +
    '自己肯定感を育む温かい賞賛のコメントを100字程度で作成してください。見出しや解説は不要です。';

  const payload = {
    contents: [{ parts: [{ text: prompt + '\n\n---\n\n' + journalContent }] }]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(API_ENDPOINT_V1 + apiKey, options);
    if (response.getResponseCode() === 200) {
      const json = JSON.parse(response.getContentText());
      return json.candidates[0].content.parts[0].text.trim();
    }
    Logger.log('API Error (Simple): ' + response.getResponseCode());
    return 'コメントの生成に失敗しました。';
  } catch (e) {
    Logger.log('Fetch Error (Simple): ' + e.message);
    return 'コメントの生成中にエラーが発生しました。';
  }
}

/**
 * Gemini API を呼び出して高度なフィードバックを JSON で取得する
 * @param {string} journalContent - ジャーナル本文
 * @returns {object} { success, jsonText, message }
 * @private
 */
function callGeminiApiForFullFeedback_(journalContent) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    return { success: false, jsonText: null, message: 'APIキーが設定されていません。' };
  }

  const prompt = `あなたは、児童一人ひとりの学びと感情に寄り添う、経験豊富な小学校の先生です。以下の児童のジャーナルを読み、フィードバックを作成してください。
# あなたの目的
児童の記述の中から「学びの芽生え」「感情の動き」「素晴らしい視点」を見つけ出し、それを価値付けることで、児童の自己肯定感とメタ認知能力を育むことです。
# 出力形式の厳密なルール
あなたの回答は、必ず以下の構造を持つJSONオブジェクト"のみ"でなければなりません。解説や前置き、\`\`\`json や \`\`\` といったマークダウンは一切含めないでください。
{"comment":"（ここに、児童全体への温かいコメントを100字以内で記述）","highlights":[{"textToHighlight":"（ここに、ジャーナル本文から「完全一致」で引用した、ハイライトすべき最も重要な部分を記述）","suggestedComment":"（ここに、ハイライト箇所への、ごく短い共感や疑問や賞賛のコメントを記述）","suggestedStamp":"（ここに、そのハイライトに最もふさわしい絵文字スタンプを1つだけ記述）"}]}
# ハイライト選定の最重要ルール
- ジャーナルの中から、あなたの目的に最も合致する箇所を「1つだけ」選び、正確に引用してください。
---
# 児童のジャーナル
${journalContent}`;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { responseMimeType: 'application/json' }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(API_ENDPOINT_V1_BETA + apiKey, options);
    if (response.getResponseCode() === 200) {
      const json = JSON.parse(response.getContentText());
      const text = json.candidates[0].content.parts[0].text;
      return { success: true, jsonText: text, message: '成功' };
    }
    Logger.log('API Error (Full): ' + response.getResponseCode());
    return { success: false, jsonText: null, message: 'APIからエラーが返されました。' };
  } catch (e) {
    Logger.log('Fetch Error (Full): ' + e.message);
    return { success: false, jsonText: null, message: 'APIへの接続中にエラーが発生しました。' };
  }
}


// =============================================================
// ■ 10. 管理画面用 API
// =============================================================

/**
 * 名簿データを全件取得する（担任行も含む）
 * @returns {object} { success, data: [{ role, name, email }] }
 */
function getRosterAll() {
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(ROSTER_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, data: [] };

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    const roster = data
      .filter(row => row[2]) // メールアドレスが空でない行のみ
      .map(row => ({ role: row[0], name: row[1], email: row[2] }));
    return { success: true, data: roster };
  } catch (e) {
    Logger.log('getRosterAll エラー: ' + e.message);
    return { success: false, message: 'エラーが発生しました：' + e.message };
  }
}

/**
 * 名簿を一括上書き保存する
 * テキストエリアからの入力を受け取り、名簿シートを再構築する。
 * @param {object[]} rows - [{ role, name, email }] の配列
 * @returns {object} 処理結果
 */
function saveRosterAll(rows) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(ROSTER_SHEET_NAME);

    // ヘッダー行以降を全削除
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).clearContent();
    }

    // 新しいデータを書き込み
    if (rows.length > 0) {
      const values = rows.map(r => [r.role || '児童', r.name, r.email]);
      sheet.getRange(2, 1, values.length, 3).setValues(values);
    }

    return { success: true, message: `名簿を保存しました（${rows.length}件）` };
  } catch (e) {
    Logger.log('saveRosterAll エラー: ' + e.message);
    return { success: false, message: '名簿の保存中にエラーが発生しました：' + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 管理設定を取得する
 * APIキーはマスクして返す。
 * @returns {object} 設定情報
 */
function getAdminSettings() {
  try {
    const props = PropertiesService.getScriptProperties();
    const apiKey = props.getProperty('GEMINI_API_KEY') || '';
    const ssId = props.getProperty(PROP_SPREADSHEET_ID) || '';
    const folderId = props.getProperty(PROP_IMAGE_FOLDER_ID) || '';

    // APIキーをマスク（先頭6文字 + ... + 末尾4文字）
    let maskedKey = '';
    if (apiKey.length > 10) {
      maskedKey = apiKey.substring(0, 6) + '****' + apiKey.substring(apiKey.length - 4);
    } else if (apiKey) {
      maskedKey = '設定済み';
    }

    return {
      success: true,
      apiKeyMasked: maskedKey,
      hasApiKey: !!apiKey,
      spreadsheetUrl: ssId ? `https://docs.google.com/spreadsheets/d/${ssId}` : '',
      imageFolderUrl: folderId ? `https://drive.google.com/drive/folders/${folderId}` : ''
    };
  } catch (e) {
    Logger.log('getAdminSettings エラー: ' + e.message);
    return { success: false, message: 'エラーが発生しました：' + e.message };
  }
}

/**
 * 管理設定を保存する
 * @param {object} settings - { apiKey }
 * @returns {object} 処理結果
 */
function saveAdminSettings(settings) {
  try {
    const props = PropertiesService.getScriptProperties();
    if (settings.apiKey !== undefined && settings.apiKey !== '') {
      props.setProperty('GEMINI_API_KEY', settings.apiKey);
    }
    return { success: true, message: '設定を保存しました！' };
  } catch (e) {
    Logger.log('saveAdminSettings エラー: ' + e.message);
    return { success: false, message: '設定の保存中にエラーが発生しました：' + e.message };
  }
}

/**
 * Gemini APIキーの疎通テストを行う
 * @param {string} apiKey - テストするAPIキー
 * @returns {object} テスト結果
 */
function testGeminiApiKey(apiKey) {
  if (!apiKey) {
    return { success: false, message: 'APIキーを入力してください。' };
  }

  const payload = {
    contents: [{ parts: [{ text: 'テスト。「OK」とだけ返してください。' }] }]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const url = `https://generativelanguage.googleapis.com/v1/models/${GEMINI_MODEL}:generateContent?key=` + apiKey;
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();

    if (code === 200) {
      return { success: true, message: '✅ APIキーは有効です！接続テスト成功。' };
    } else if (code === 400) {
      return { success: false, message: '❌ APIキーが無効です。正しいキーを入力してください。' };
    } else if (code === 403) {
      return { success: false, message: '❌ APIキーに権限がありません。Gemini APIが有効化されているか確認してください。' };
    } else {
      return { success: false, message: `❌ エラーが発生しました (HTTP ${code})` };
    }
  } catch (e) {
    Logger.log('testGeminiApiKey エラー: ' + e.message);
    return { success: false, message: '❌ 接続エラー: ' + e.message };
  }
}

/**
 * ジャーナルデータを全削除する（シートのデータ行をクリア）
 * @returns {object} 処理結果
 */
function resetAllData() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(JOURNAL_SHEET_NAME);

    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, JOURNAL_HEADERS.length).clearContent();
      // 空行を削除
      if (sheet.getLastRow() > 1) {
        sheet.deleteRows(2, sheet.getLastRow() - 1);
      }
    }

    return { success: true, message: 'ジャーナルデータを全て削除しました。' };
  } catch (e) {
    Logger.log('resetAllData エラー: ' + e.message);
    return { success: false, message: 'データの削除中にエラーが発生しました：' + e.message };
  } finally {
    lock.releaseLock();
  }
}


// =============================================================
// ■ 11. 提出状況 & エクスポート機能
// =============================================================

/**
 * 本日の提出状況を名簿と突合して返す
 * @returns {object} { success, data: { total, submitted, students: [{ name, email, submitted }] } }
 */
function getSubmissionStatus() {
  try {
    const ss = getSpreadsheet_();
    const roster = getClassRoster(); // 児童のみ（担任除外）
    const today = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd');

    // ジャーナルシートから本日提出分のメールを抽出
    const jSheet = ss.getSheetByName(JOURNAL_SHEET_NAME);
    const submittedEmails = new Set();

    if (jSheet && jSheet.getLastRow() >= 2) {
      const data = jSheet.getDataRange().getValues();
      const headers = data.shift();
      const emailIdx = headers.indexOf('email');
      const tsIdx = headers.indexOf('timestamp');
      const deletedIdx = headers.indexOf('deletedAt');

      data.forEach(row => {
        if (row[deletedIdx]) return; // 論理削除済み
        if (row[tsIdx] instanceof Date) {
          const d = Utilities.formatDate(row[tsIdx], 'JST', 'yyyy/MM/dd');
          if (d === today) {
            submittedEmails.add(row[emailIdx]);
          }
        }
      });
    }

    const students = roster.map(s => ({
      name: s.name,
      email: s.email,
      submitted: submittedEmails.has(s.email)
    }));

    return {
      success: true,
      data: {
        total: students.length,
        submitted: students.filter(s => s.submitted).length,
        students: students
      }
    };
  } catch (e) {
    Logger.log('getSubmissionStatus エラー: ' + e.message);
    return { success: false, message: 'エラーが発生しました：' + e.message };
  }
}

/**
 * ジャーナルデータをCSVとしてエクスポートする
 * @param {object} params - { startDate, endDate, email ('all' or 特定メール) }
 * @returns {object} { success, fileUrl, fileName }
 */
function exportJournalsCsv(params) {
  try {
    const journals = getFilteredJournals_(params);
    if (journals.length === 0) {
      return { success: false, message: '該当するジャーナルが見つかりませんでした。' };
    }

    // CSV ヘッダー
    const csvHeaders = ['日付', '氏名', 'テーマ', '本文', '気持ち', '先生のコメント', 'ステータス'];
    const csvRows = journals.map(j => [
      j.date || '',
      j.studentName || '',
      j.theme || '',
      (j.content || '').replace(/"/g, '""'),
      j.emotion || '',
      (j.teacherComment || '').replace(/"/g, '""'),
      j.status || ''
    ]);

    // CSV文字列を生成（BOM付きUTF-8でExcel対応）
    const csvContent = [csvHeaders, ...csvRows]
      .map(row => row.map(cell => `"${cell}"`).join(','))
      .join('\r\n');
    const bom = '\uFEFF';

    // ファイル名を生成
    const dateStr = Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss');
    const fileName = `ジャーナルデータ_${dateStr}.csv`;

    // Drive に保存
    const folder = getExportFolder_();
    const blob = Utilities.newBlob(bom + csvContent, 'text/csv', fileName);
    const file = folder.createFile(blob);

    return {
      success: true,
      fileUrl: file.getUrl(),
      fileName: fileName,
      message: `${journals.length}件のデータをCSVで出力しました。`
    };
  } catch (e) {
    Logger.log('exportJournalsCsv エラー: ' + e.message);
    return { success: false, message: 'CSV出力中にエラーが発生しました：' + e.message };
  }
}

/**
 * ジャーナルデータを書式付きPDFとしてエクスポートする
 * @param {object} params - { startDate, endDate, email ('all' or 特定メール) }
 * @returns {object} { success, fileUrl, fileName }
 */
function exportJournalsPdf(params) {
  try {
    const journals = getFilteredJournals_(params);
    if (journals.length === 0) {
      return { success: false, message: '該当するジャーナルが見つかりませんでした。' };
    }

    // 児童ごとにグループ化
    const grouped = {};
    journals.forEach(j => {
      const key = j.studentName || '不明';
      if (!grouped[key]) grouped[key] = [];
      grouped[key].push(j);
    });

    // 期間表示用テキスト
    const periodText = formatPeriodText_(params);

    // Google Doc を一時作成
    const dateStr = Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss');
    const fileName = `ふり返りジャーナル_${dateStr}`;
    const doc = DocumentApp.create(fileName);
    const body = doc.getBody();

    // デフォルトスタイル
    body.setAttributes({
      [DocumentApp.Attribute.FONT_FAMILY]: 'BIZ UDPGothic',
      [DocumentApp.Attribute.FONT_SIZE]: 11
    });

    const studentNames = Object.keys(grouped).sort();
    studentNames.forEach((name, sIdx) => {
      const studentJournals = grouped[name];

      // ── 表紙ヘッダー ──
      const title = body.appendParagraph('📔 ふり返りジャーナル');
      title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
      title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      title.setAttributes({
        [DocumentApp.Attribute.FONT_SIZE]: 20,
        [DocumentApp.Attribute.BOLD]: true,
        [DocumentApp.Attribute.FOREGROUND_COLOR]: '#1565C0'
      });

      const nameP = body.appendParagraph(name);
      nameP.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      nameP.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      nameP.setAttributes({
        [DocumentApp.Attribute.FONT_SIZE]: 16,
        [DocumentApp.Attribute.BOLD]: true
      });

      const periodP = body.appendParagraph(periodText);
      periodP.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      periodP.setAttributes({
        [DocumentApp.Attribute.FONT_SIZE]: 11,
        [DocumentApp.Attribute.FOREGROUND_COLOR]: '#666666'
      });

      body.appendParagraph(''); // 空行

      // ── 各ジャーナル ──
      studentJournals.forEach((j, jIdx) => {
        // 日付とテーマ
        const dateHeading = body.appendParagraph(`📅 ${j.date || '日付不明'}`);
        dateHeading.setHeading(DocumentApp.ParagraphHeading.HEADING3);
        dateHeading.setAttributes({
          [DocumentApp.Attribute.FONT_SIZE]: 13,
          [DocumentApp.Attribute.BOLD]: true,
          [DocumentApp.Attribute.FOREGROUND_COLOR]: '#1565C0'
        });

        if (j.theme) {
          const themeP = body.appendParagraph(`テーマ: ${j.theme}`);
          themeP.setAttributes({
            [DocumentApp.Attribute.ITALIC]: true,
            [DocumentApp.Attribute.FOREGROUND_COLOR]: '#888888',
            [DocumentApp.Attribute.FONT_SIZE]: 10
          });
        }

        // 本文
        if (j.content) {
          body.appendParagraph(''); // 空行
          const contentP = body.appendParagraph(j.content);
          contentP.setAttributes({
            [DocumentApp.Attribute.FONT_SIZE]: 11,
            [DocumentApp.Attribute.LINE_SPACING]: 1.6
          });
        }

        // 気持ちスタンプ
        if (j.emotion) {
          body.appendParagraph('');
          const emotionP = body.appendParagraph(`気持ち: ${j.emotion}`);
          emotionP.setAttributes({
            [DocumentApp.Attribute.FONT_SIZE]: 10,
            [DocumentApp.Attribute.FOREGROUND_COLOR]: '#E65100'
          });
        }

        // 先生のコメント
        if (j.teacherComment) {
          body.appendParagraph('');
          const commentLabel = body.appendParagraph('💬 先生より');
          commentLabel.setAttributes({
            [DocumentApp.Attribute.FONT_SIZE]: 10,
            [DocumentApp.Attribute.BOLD]: true,
            [DocumentApp.Attribute.FOREGROUND_COLOR]: '#1565C0'
          });
          const commentP = body.appendParagraph(j.teacherComment);
          commentP.setAttributes({
            [DocumentApp.Attribute.FONT_SIZE]: 10,
            [DocumentApp.Attribute.FOREGROUND_COLOR]: '#333333'
          });
        }

        // 区切り線（最後のジャーナル以外）
        if (jIdx < studentJournals.length - 1) {
          body.appendParagraph('');
          const hr = body.appendParagraph('─'.repeat(40));
          hr.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
          hr.setAttributes({
            [DocumentApp.Attribute.FOREGROUND_COLOR]: '#cccccc',
            [DocumentApp.Attribute.FONT_SIZE]: 8
          });
          body.appendParagraph('');
        }
      });

      // 児童間のページ区切り
      if (sIdx < studentNames.length - 1) {
        body.appendPageBreak();
      }
    });

    // Doc を保存して PDF に変換
    doc.saveAndClose();
    const docFile = DriveApp.getFileById(doc.getId());
    const pdfBlob = docFile.getAs('application/pdf');
    pdfBlob.setName(fileName + '.pdf');

    // PDF を Drive に保存
    const folder = getExportFolder_();
    const pdfFile = folder.createFile(pdfBlob);

    // 一時 Doc を削除
    docFile.setTrashed(true);

    return {
      success: true,
      fileUrl: pdfFile.getUrl(),
      fileName: fileName + '.pdf',
      message: `${journals.length}件のデータをPDFで出力しました。`
    };
  } catch (e) {
    Logger.log('exportJournalsPdf エラー: ' + e.message);
    return { success: false, message: 'PDF出力中にエラーが発生しました：' + e.message };
  }
}

/**
 * フィルタ条件に合致するジャーナルを取得する（内部ヘルパー）
 * @param {object} params - { startDate, endDate, email }
 * @returns {object[]} ジャーナルリスト
 * @private
 */
function getFilteredJournals_(params) {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(JOURNAL_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const deletedAtIdx = headers.indexOf('deletedAt');
  const emailIdx = headers.indexOf('email');
  const tsIdx = headers.indexOf('timestamp');

  // 名前マップ
  const roster = getClassRoster();
  const nameMap = roster.reduce((map, u) => {
    map[u.email] = u.name;
    return map;
  }, {});

  // 日付フィルタ用
  const startDate = params.startDate ? new Date(params.startDate + 'T00:00:00+09:00') : null;
  const endDate = params.endDate ? new Date(params.endDate + 'T23:59:59+09:00') : null;
  // メールフィルタ用（大文字小文字・前後スペースを正規化）
  const filterEmail = (params.email && params.email !== 'all')
    ? params.email.trim().toLowerCase() : null;

  return data
    .filter(row => {
      if (row[deletedAtIdx]) return false; // 論理削除
      // メールフィルタ
      if (filterEmail && String(row[emailIdx]).trim().toLowerCase() !== filterEmail) return false;
      // 日付フィルタ
      if (row[tsIdx] instanceof Date) {
        if (startDate && row[tsIdx] < startDate) return false;
        if (endDate && row[tsIdx] > endDate) return false;
      }
      return true;
    })
    .map(row => {
      const journal = rowToJournalObject_(headers, row);
      journal.studentName = nameMap[row[emailIdx]] || '不明';
      return journal;
    })
    .sort((a, b) => {
      // 児童名順 → 日付順
      const nameComp = (a.studentName || '').localeCompare(b.studentName || '');
      if (nameComp !== 0) return nameComp;
      return (a.timestamp || 0) - (b.timestamp || 0);
    });
}

/**
 * エクスポート用フォルダを取得（なければ作成）
 * @returns {GoogleAppsScript.Drive.Folder} フォルダ
 * @private
 */
function getExportFolder_() {
  const props = PropertiesService.getScriptProperties();
  const parentFolderId = props.getProperty(PROP_IMAGE_FOLDER_ID);

  let parentFolder;
  if (parentFolderId) {
    parentFolder = DriveApp.getFolderById(parentFolderId).getParents().next();
  } else {
    parentFolder = DriveApp.getRootFolder();
  }

  // 「エクスポート」サブフォルダを探す or 作成
  const folderName = 'ふり返りジャーナル_エクスポート';
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(folderName);
}

/**
 * 期間テキストを生成する（内部ヘルパー）
 * @param {object} params - { startDate, endDate }
 * @returns {string} 期間表示テキスト
 * @private
 */
function formatPeriodText_(params) {
  const format = (dateStr) => {
    if (!dateStr) return '';
    const d = new Date(dateStr);
    return `${d.getFullYear()}年${d.getMonth() + 1}月${d.getDate()}日`;
  };

  if (params.startDate && params.endDate) {
    return `${format(params.startDate)} 〜 ${format(params.endDate)}`;
  } else if (params.startDate) {
    return `${format(params.startDate)} 〜`;
  } else if (params.endDate) {
    return `〜 ${format(params.endDate)}`;
  }
  return '全期間';
}
