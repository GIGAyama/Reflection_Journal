/**
 * ============================================================
 * ふりかえりジャーナル — サーバーサイド スクリプト (エラー対策版)
 * ============================================================
 */

const ROSTER_SHEET_NAME  = '児童名簿';
const JOURNAL_SHEET_NAME = 'ジャーナルデータ';
const THEME_SHEET_NAME   = 'テーマ設定';

const GEMINI_MODEL = 'gemini-2.5-flash';
const API_ENDPOINT_V1      = 'https://generativelanguage.googleapis.com/v1/models/' + GEMINI_MODEL + ':generateContent?key=';
const API_ENDPOINT_V1_BETA = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent?key=';

const PROP_SPREADSHEET_ID  = 'SPREADSHEET_ID';
const PROP_IMAGE_FOLDER_ID = 'IMAGE_FOLDER_ID';
const PROP_ADMIN_EMAIL     = 'ADMIN_EMAIL';
const PROP_INITIALIZED     = 'APP_INITIALIZED';

const JOURNAL_HEADERS = [
  'journalId', 'timestamp', 'email', 'theme', 'content',
  'imageFileId', 'emotion', 'teacherComment', 'highlights',
  'teacherStamp', 'status', 'pastComment', 'deletedAt'
];
const ROSTER_HEADERS = ['役割', '氏名', 'メールアドレス'];
const THEME_HEADERS = ['日付', 'テーマ'];

function doGet(e) {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty(PROP_INITIALIZED)) {
    initializeApp_(props);
  }

  const userEmail = Session.getActiveUser().getEmail();
  const adminEmail = props.getProperty(PROP_ADMIN_EMAIL);
  const isAdmin = (userEmail === adminEmail);
  const userData = getUserData_(userEmail, isAdmin);

  const tmpl = HtmlService.createTemplateFromFile('index');
  
  let journals = [];
  let classRoster = [];
  
  if (userData) {
      if (userData.role === '担任') {
          classRoster = getClassRoster();
          journals = getJournalsForTeacher();
      } else {
          journals = getJournalsForStudent(userEmail);
      }
  }

  const dataObj = {
    user: userData,
    isAdmin: isAdmin,
    todayTheme: getTodayTheme(),
    journals: journals,
    classRoster: classRoster
  };

  // エラー対策: JSON化してHTML崩れを防ぐ
  tmpl.initialData = JSON.stringify(dataObj).replace(/</g, '\\u003c');

  return tmpl.evaluate()
    .setTitle('ふりかえりジャーナル')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl('https://drive.google.com/uc?id=1rJjk2hoVW64rVz0kb-fdARn7g02Q5rjI&.png');
}

function getSpreadsheet_() {
  const ssId = PropertiesService.getScriptProperties().getProperty(PROP_SPREADSHEET_ID);
  if (!ssId) throw new Error('データベースが初期化されていません。');
  return SpreadsheetApp.openById(ssId);
}

function initializeApp_(props) {
  const adminEmail = Session.getActiveUser().getEmail();
  props.setProperty(PROP_ADMIN_EMAIL, adminEmail);

  const ss = SpreadsheetApp.create('ふりかえりジャーナル_DB');
  props.setProperty(PROP_SPREADSHEET_ID, ss.getId());

  createSheetIfNotExists_(ss, ROSTER_SHEET_NAME, ROSTER_HEADERS);
  createSheetIfNotExists_(ss, JOURNAL_SHEET_NAME, JOURNAL_HEADERS);
  createSheetIfNotExists_(ss, THEME_SHEET_NAME, THEME_HEADERS);

  const defaultSheet = ss.getSheetByName('シート1');
  if (defaultSheet && ss.getSheets().length > 1) ss.deleteSheet(defaultSheet);

  const rosterSheet = ss.getSheetByName(ROSTER_SHEET_NAME);
  rosterSheet.appendRow(['担任', '管理者', adminEmail]);

  const folder = DriveApp.createFolder('ふりかえりジャーナル_画像');
  props.setProperty(PROP_IMAGE_FOLDER_ID, folder.getId());

  props.setProperty(PROP_INITIALIZED, 'true');
}

function createSheetIfNotExists_(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) return;
  sheet = ss.insertSheet(sheetName);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, sheet.getMaxRows(), headers.length).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  headers.forEach(function(_, i) { sheet.autoResizeColumn(i + 1); });
}

function getUserData_(email, isAdmin) {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(ROSTER_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return isAdmin ? { role: '担任', name: '管理者', email: email } : null;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][2] === email) return { role: data[i][0] === '担任' ? '担任' : '児童', name: data[i][1], email: data[i][2] };
  }
  return isAdmin ? { role: '担任', name: '管理者', email: email } : null;
}

function getUserData(email) {
  const adminEmail = PropertiesService.getScriptProperties().getProperty(PROP_ADMIN_EMAIL);
  return getUserData_(email, email === adminEmail);
}

function getClassRoster() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(ROSTER_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  return data.filter(function(row) { return row[0] !== '担任' && row[2]; }).map(function(row) { return { name: row[1], email: row[2] }; });
}

function saveJournal(journalData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(JOURNAL_SHEET_NAME);
    const newRow = [
      Utilities.getUuid(), new Date(), journalData.email, journalData.theme, journalData.content,
      journalData.imageFileId || '', journalData.emotion || '', '', '[]', '', '未返却', '', ''
    ];
    sheet.appendRow(newRow);
    return { success: true, message: 'ジャーナルを提出しました！' };
  } catch (e) {
    return { success: false, message: '保存中にエラーが発生しました：' + e.message };
  } finally {
    lock.releaseLock();
  }
}

function getJournalsForStudent(email) {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(JOURNAL_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const emailIdx = headers.indexOf('email');
  const deletedAtIdx = headers.indexOf('deletedAt');
  return data.filter(function(row) { return row[emailIdx] === email && !row[deletedAtIdx]; }).map(function(row) { return rowToJournalObject_(headers, row); });
}

function getJournalsForTeacher() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(JOURNAL_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const deletedAtIdx = headers.indexOf('deletedAt');
  const emailIdx = headers.indexOf('email');
  const nameMap = getClassRoster().reduce(function(map, u) { map[u.email] = u.name; return map; }, {});

  return data.filter(function(row) { return !row[deletedAtIdx]; }).map(function(row) {
    const journal = rowToJournalObject_(headers, row);
    journal.studentName = nameMap[row[emailIdx]] || '不明';
    return journal;
  });
}

function rowToJournalObject_(headers, row) {
  const journal = {};
  headers.forEach(function(h, i) { journal[h] = row[i]; });
  if (journal.timestamp instanceof Date) {
    journal.date = Utilities.formatDate(journal.timestamp, 'JST', 'yyyy/MM/dd HH:mm');
  }
  journal.imageUrl = journal.imageFileId ? 'https://lh3.googleusercontent.com/d/' + journal.imageFileId : '';
  return journal;
}

function saveFeedback(feedbackData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(JOURNAL_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIdx = headers.indexOf('journalId');

    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] === feedbackData.journalId) {
        const rowNum = i + 1;
        sheet.getRange(rowNum, headers.indexOf('teacherComment') + 1).setValue(feedbackData.comment);
        sheet.getRange(rowNum, headers.indexOf('highlights') + 1).setValue(feedbackData.highlights || '[]');
        sheet.getRange(rowNum, headers.indexOf('status') + 1).setValue('返却済み');
        return { success: true, message: 'フィードバックを保存しました！' };
      }
    }
    return { success: false, message: '該当するジャーナルが見つかりません。' };
  } catch (e) {
    return { success: false, message: 'エラー：' + e.message };
  } finally {
    lock.releaseLock();
  }
}

function revertJournalStatus(journalId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(JOURNAL_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    for (let i = 1; i < data.length; i++) {
      if (data[i][headers.indexOf('journalId')] === journalId) {
        sheet.getRange(i + 1, headers.indexOf('status') + 1).setValue('未返却');
        return { success: true };
      }
    }
    return { success: false, message: '見つかりませんでした。' };
  } finally { lock.releaseLock(); }
}

function addPastComment(journalId, comment) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(JOURNAL_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    for (let i = 1; i < data.length; i++) {
      if (data[i][headers.indexOf('journalId')] === journalId) {
        sheet.getRange(i + 1, headers.indexOf('pastComment') + 1).setValue(comment);
        return { success: true, message: '保存しました！' };
      }
    }
    return { success: false, message: '見つかりませんでした。' };
  } finally { lock.releaseLock(); }
}

function deleteJournal(journalId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(JOURNAL_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    for (let i = 1; i < data.length; i++) {
      if (data[i][headers.indexOf('journalId')] === journalId) {
        sheet.getRange(i + 1, headers.indexOf('deletedAt') + 1).setValue(new Date());
        return { success: true, message: '削除しました。' };
      }
    }
    return { success: false, message: '見つかりませんでした。' };
  } finally { lock.releaseLock(); }
}

function uploadImage(fileData) {
  try {
    const folderId = PropertiesService.getScriptProperties().getProperty(PROP_IMAGE_FOLDER_ID);
    if (!folderId) return { success: false, message: 'フォルダがありません。' };
    const folder = DriveApp.getFolderById(folderId);
    const decoded = Utilities.base64Decode(fileData.data, Utilities.Charset.UTF_8);
    const blob = Utilities.newBlob(decoded, fileData.mimeType, fileData.fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { success: true, fileId: file.getId(), fileName: file.getName() };
  } catch (e) {
    return { success: false, message: 'エラー：' + e.message };
  }
}

function getTodayTheme() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(THEME_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return getWeeklyThemeForToday_() || '今日の学びをふり返ろう';
  
  const todayStr = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd');
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i > 0; i--) {
    if (data[i][0] instanceof Date && Utilities.formatDate(data[i][0], 'JST', 'yyyy/MM/dd') === todayStr) {
      return data[i][1];
    }
  }
  return getWeeklyThemeForToday_() || '今日の学びをふり返ろう';
}

function setTodayTheme(theme) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    getSpreadsheet_().getSheetByName(THEME_SHEET_NAME).appendRow([new Date(), theme]);
    return { success: true, message: 'テーマを設定しました！' };
  } finally { lock.releaseLock(); }
}

function getWeeklyThemes() {
  try {
    const json = PropertiesService.getScriptProperties().getProperty('WEEKLY_THEMES');
    return { success: true, data: json ? JSON.parse(json) : { mon: '', tue: '', wed: '', thu: '', fri: '' } };
  } catch (e) { return { success: false, message: 'Error: ' + e.message }; }
}

function saveWeeklyThemes(themes) {
  try {
    PropertiesService.getScriptProperties().setProperty('WEEKLY_THEMES', JSON.stringify(themes));
    return { success: true, message: '保存しました！' };
  } catch (e) { return { success: false, message: 'Error: ' + e.message }; }
}

function getWeeklyThemeForToday_() {
  const json = PropertiesService.getScriptProperties().getProperty('WEEKLY_THEMES');
  if (!json) return null;
  const themes = JSON.parse(json);
  const dayKeys = ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat'];
  return themes[dayKeys[new Date().getDay()]] || null;
}

function quickReturn(journalId, stamp) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(JOURNAL_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    for (let i = 1; i < data.length; i++) {
      if (data[i][headers.indexOf('journalId')] === journalId) {
        sheet.getRange(i + 1, headers.indexOf('teacherStamp') + 1).setValue(stamp);
        sheet.getRange(i + 1, headers.indexOf('status') + 1).setValue('返却済み');
        return { success: true, message: 'スタンプで返却しました！' };
      }
    }
    return { success: false, message: '見つかりませんでした。' };
  } finally { lock.releaseLock(); }
}

function batchReturnAll() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(JOURNAL_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    let count = 0;
    for (let i = 1; i < data.length; i++) {
      if (data[i][headers.indexOf('status')] === '未返却' && data[i][headers.indexOf('teacherComment')] && !data[i][headers.indexOf('deletedAt')]) {
        sheet.getRange(i + 1, headers.indexOf('status') + 1).setValue('返却済み');
        count++;
      }
    }
    return { success: true, count: count, message: count > 0 ? count + '件を一括返却しました！' : '返却対象がありません。' };
  } finally { lock.releaseLock(); }
}

function generateAiSimpleCommentsForAll() {
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(JOURNAL_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    let successCount = 0;

    data.forEach(function(row, index) {
      if (row[headers.indexOf('status')] === '未返却' && row[headers.indexOf('content')] && !row[headers.indexOf('deletedAt')]) {
        const comment = callGeminiApiForSimpleComment_(row[headers.indexOf('content')]);
        sheet.getRange(index + 2, headers.indexOf('teacherComment') + 1).setValue(comment);
        successCount++;
      }
    });
    return { success: true, message: 'AIコメント案を作成しました。（' + successCount + '件）' };
  } catch (e) {
    return { success: false, message: 'AI処理エラー：' + e.message };
  }
}

function generateAiFullFeedbackForAll() {
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(JOURNAL_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    let successCount = 0;
    let errorCount = 0;

    data.forEach(function(row, index) {
      if (row[headers.indexOf('status')] === '未返却' && row[headers.indexOf('content')] && !row[headers.indexOf('deletedAt')]) {
        const aiResponse = callGeminiApiForFullFeedback_(row[headers.indexOf('content')]);
        if (aiResponse && aiResponse.success) {
          try {
            const feedback = JSON.parse(aiResponse.jsonText);
            if (feedback.comment) {
              sheet.getRange(index + 2, headers.indexOf('teacherComment') + 1).setValue(feedback.comment);
            }
            if (feedback.highlights && feedback.highlights.length > 0) {
              const content = row[headers.indexOf('content')];
              const hlToSave = [];
              feedback.highlights.forEach(function(h) {
                const startIndex = content.indexOf(h.textToHighlight);
                if (startIndex !== -1) {
                  hlToSave.push({
                    id: 'hl-' + Date.now() + Math.random(), text: h.textToHighlight, comment: h.suggestedComment || '', stamp: h.suggestedStamp || '',
                    startOffset: startIndex, endOffset: startIndex + h.textToHighlight.length
                  });
                }
              });
              if (hlToSave.length > 0) sheet.getRange(index + 2, headers.indexOf('highlights') + 1).setValue(JSON.stringify(hlToSave));
            }
            successCount++;
          } catch (e) { errorCount++; }
        } else { errorCount++; }
      }
    });
    return { success: true, message: 'AI高度分析完了。（成功: ' + successCount + '件, 失敗: ' + errorCount + '件）' };
  } catch (e) { return { success: false, message: 'AIエラー：' + e.message }; }
}

function callGeminiApiForSimpleComment_(journalContent) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) return 'エラー: APIキーが設定されていません。';
  
  const promptText = 'あなたは児童の小さな頑張りやユニークな視点を見つけて具体的に褒めるのが得意な、経験豊富な小学校の先生です。以下の記述を読み、児童が努力した点などを引用しつつ、自己肯定感を育む温かい賞賛のコメントを100字程度で作成してください。見出しや解説は不要です。\n\n' + journalContent;
  const payload = { contents: [{ parts: [{ text: promptText }] }] };
  
  try {
    const res = UrlFetchApp.fetch(API_ENDPOINT_V1 + apiKey, { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true });
    if (res.getResponseCode() === 200) return JSON.parse(res.getContentText()).candidates[0].content.parts[0].text.trim();
    return 'コメント生成失敗';
  } catch (e) { return 'エラー発生'; }
}

function callGeminiApiForFullFeedback_(journalContent) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) return { success: false, jsonText: null, message: 'APIキー未設定' };
  
  const promptText = 'あなたは経験豊富な小学校の先生です。以下の児童のジャーナルを読み、フィードバックを作成してください。\n' +
    '# 出力形式の厳密なルール\n' +
    '必ず以下の構造を持つJSONのみを出力してください。\n' +
    '{"comment":"（全体への温かいコメント100字以内）","highlights":[{"textToHighlight":"（本文から完全一致で引用）","suggestedComment":"（ハイライト箇所へのコメント）","suggestedStamp":"（絵文字1つ）"}]}\n' +
    '---\n' + journalContent;
    
  const payload = { contents: [{ parts: [{ text: promptText }] }], generationConfig: { responseMimeType: 'application/json' } };
  
  try {
    const res = UrlFetchApp.fetch(API_ENDPOINT_V1_BETA + apiKey, { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true });
    if (res.getResponseCode() === 200) return { success: true, jsonText: JSON.parse(res.getContentText()).candidates[0].content.parts[0].text, message: '成功' };
    return { success: false, jsonText: null, message: 'エラー' };
  } catch (e) { return { success: false, jsonText: null, message: 'エラー' }; }
}

function getRosterAll() {
  try {
    const sheet = getSpreadsheet_().getSheetByName(ROSTER_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, data: [] };
    const roster = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues().filter(function(row) { return row[2]; }).map(function(row) { return { role: row[0], name: row[1], email: row[2] }; });
    return { success: true, data: roster };
  } catch (e) { return { success: false, message: e.message }; }
}

function saveRosterAll(rows) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSpreadsheet_().getSheetByName(ROSTER_SHEET_NAME);
    if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).clearContent();
    if (rows.length > 0) {
      const values = rows.map(function(r) { return [r.role || '児童', r.name, r.email]; });
      sheet.getRange(2, 1, values.length, 3).setValues(values);
    }
    return { success: true, message: '名簿を保存しました' };
  } finally { lock.releaseLock(); }
}

function getAdminSettings() {
  try {
    const props = PropertiesService.getScriptProperties();
    const apiKey = props.getProperty('GEMINI_API_KEY') || '';
    const ssId = props.getProperty(PROP_SPREADSHEET_ID) || '';
    const folderId = props.getProperty(PROP_IMAGE_FOLDER_ID) || '';
    let maskedKey = apiKey.length > 10 ? apiKey.substring(0, 6) + '****' + apiKey.substring(apiKey.length - 4) : (apiKey ? '設定済み' : '');
    return { success: true, apiKeyMasked: maskedKey, hasApiKey: !!apiKey, spreadsheetUrl: ssId ? 'https://docs.google.com/spreadsheets/d/' + ssId : '', imageFolderUrl: folderId ? 'https://drive.google.com/drive/folders/' + folderId : '' };
  } catch (e) { return { success: false, message: e.message }; }
}

function saveAdminSettings(settings) {
  try {
    if (settings.apiKey !== undefined && settings.apiKey !== '') PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', settings.apiKey);
    return { success: true, message: '設定を保存しました！' };
  } catch (e) { return { success: false, message: e.message }; }
}

function testGeminiApiKey(apiKey) {
  if (!apiKey) return { success: false, message: 'APIキーを入力してください。' };
  try {
    const res = UrlFetchApp.fetch(API_ENDPOINT_V1 + apiKey, {
      method: 'post', contentType: 'application/json', payload: JSON.stringify({ contents: [{ parts: [{ text: 'テスト。OKと返して' }] }] }), muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    if (code === 200) return { success: true, message: 'APIキーは有効です！' };
    return { success: false, message: 'エラー (HTTP ' + code + ')' };
  } catch (e) { return { success: false, message: '接続エラー: ' + e.message }; }
}

function resetAllData() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSpreadsheet_().getSheetByName(JOURNAL_SHEET_NAME);
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, JOURNAL_HEADERS.length).clearContent();
      if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
    }
    return { success: true, message: '全データを削除しました。' };
  } finally { lock.releaseLock(); }
}

function getSubmissionStatus() {
  try {
    const ss = getSpreadsheet_();
    const roster = getClassRoster();
    const today = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd');
    const jSheet = ss.getSheetByName(JOURNAL_SHEET_NAME);
    const submittedEmails = new Set();

    if (jSheet && jSheet.getLastRow() >= 2) {
      const data = jSheet.getDataRange().getValues();
      const headers = data.shift();
      data.forEach(function(row) {
        if (row[headers.indexOf('deletedAt')]) return;
        if (row[headers.indexOf('timestamp')] instanceof Date && Utilities.formatDate(row[headers.indexOf('timestamp')], 'JST', 'yyyy/MM/dd') === today) {
          submittedEmails.add(row[headers.indexOf('email')]);
        }
      });
    }

    const students = roster.map(function(s) { return { name: s.name, email: s.email, submitted: submittedEmails.has(s.email) }; });
    return { success: true, data: { total: students.length, submitted: students.filter(function(s) { return s.submitted; }).length, students: students } };
  } catch (e) { return { success: false, message: e.message }; }
}

function exportJournalsCsv(params) {
  try {
    const journals = getFilteredJournals_(params);
    if (journals.length === 0) return { success: false, message: 'データがありません。' };
    
    const csvHeaders = ['日付', '氏名', 'テーマ', '本文', '気持ち', '先生のコメント', 'ステータス'];
    const csvRows = journals.map(function(j) { return [j.date||'', j.studentName||'', j.theme||'', (j.content||'').replace(/"/g, '""'), j.emotion||'', (j.teacherComment||'').replace(/"/g, '""'), j.status||'']; });
    const csvContent = [csvHeaders].concat(csvRows).map(function(row) { return row.map(function(cell) { return '"' + cell + '"'; }).join(','); }).join('\r\n');
    
    const fileName = 'ジャーナルデータ_' + Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss') + '.csv';
    const file = getExportFolder_().createFile(Utilities.newBlob('\uFEFF' + csvContent, 'text/csv', fileName));
    return { success: true, fileUrl: file.getUrl(), fileName: fileName, message: 'CSV出力完了' };
  } catch (e) { return { success: false, message: e.message }; }
}

function exportJournalsPdf(params) {
  try {
    const journals = getFilteredJournals_(params);
    if (journals.length === 0) return { success: false, message: 'データがありません。' };

    const grouped = {};
    journals.forEach(function(j) { const key = j.studentName || '不明'; if (!grouped[key]) grouped[key] = []; grouped[key].push(j); });

    const fileName = 'ふりかえりジャーナル_' + Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss');
    const doc = DocumentApp.create(fileName);
    const body = doc.getBody();
    
    const attrs = {};
    attrs[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
    attrs[DocumentApp.Attribute.FONT_SIZE] = 11;
    body.setAttributes(attrs);

    const keys = Object.keys(grouped).sort();
    keys.forEach(function(name, sIdx) {
      body.appendParagraph(name).setHeading(DocumentApp.ParagraphHeading.HEADING2).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      grouped[name].forEach(function(j) {
        body.appendParagraph('📅 ' + (j.date || '')).setHeading(DocumentApp.ParagraphHeading.HEADING3);
        if (j.theme) body.appendParagraph('テーマ: ' + j.theme);
        if (j.content) body.appendParagraph('\n' + j.content);
        if (j.teacherComment) body.appendParagraph('\n💬 先生より\n' + j.teacherComment);
        body.appendParagraph('\n----------------------------------------\n');
      });
      if (sIdx < keys.length - 1) body.appendPageBreak();
    });

    doc.saveAndClose();
    const pdfFile = getExportFolder_().createFile(DriveApp.getFileById(doc.getId()).getAs('application/pdf').setName(fileName + '.pdf'));
    DriveApp.getFileById(doc.getId()).setTrashed(true);

    return { success: true, fileUrl: pdfFile.getUrl(), fileName: fileName + '.pdf', message: 'PDF出力完了' };
  } catch (e) { return { success: false, message: e.message }; }
}

function getFilteredJournals_(params) {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(JOURNAL_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const nameMap = getClassRoster().reduce(function(map, u) { map[u.email] = u.name; return map; }, {});
  const startDate = params.startDate ? new Date(params.startDate + 'T00:00:00+09:00') : null;
  const endDate = params.endDate ? new Date(params.endDate + 'T23:59:59+09:00') : null;
  const filterEmail = (params.email && params.email !== 'all') ? params.email.trim().toLowerCase() : null;

  return data.filter(function(row) {
    if (row[headers.indexOf('deletedAt')]) return false;
    if (filterEmail && String(row[headers.indexOf('email')]).trim().toLowerCase() !== filterEmail) return false;
    if (row[headers.indexOf('timestamp')] instanceof Date) {
      if (startDate && row[headers.indexOf('timestamp')] < startDate) return false;
      if (endDate && row[headers.indexOf('timestamp')] > endDate) return false;
    }
    return true;
  }).map(function(row) {
    const j = rowToJournalObject_(headers, row);
    j.studentName = nameMap[row[headers.indexOf('email')]] || '不明';
    return j;
  }).sort(function(a, b) { return (a.studentName||'').localeCompare(b.studentName||'') || (a.timestamp||0) - (b.timestamp||0); });
}

function getExportFolder_() {
  const pFolderId = PropertiesService.getScriptProperties().getProperty(PROP_IMAGE_FOLDER_ID);
  const parent = pFolderId ? DriveApp.getFolderById(pFolderId).getParents().next() : DriveApp.getRootFolder();
  const folders = parent.getFoldersByName('エクスポート');
  return folders.hasNext() ? folders.next() : parent.createFolder('エクスポート');
}
