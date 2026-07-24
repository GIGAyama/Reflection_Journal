/**
 * ============================================================
 * ふりかえりジャーナル — サーバーサイド スクリプト (自動リカバリ対応版)
 * ============================================================
 */

const ROSTER_SHEET_NAME  = '児童名簿';
const JOURNAL_SHEET_NAME = 'ジャーナルデータ';
const THEME_SHEET_NAME   = 'テーマ設定';

const GEMINI_MODEL = 'gemini-2.5-flash';
const API_ENDPOINT_V1      = 'https://generativelanguage.googleapis.com/v1/models/' + GEMINI_MODEL + ':generateContent?key=';
const API_ENDPOINT_V1_BETA = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent?key=';

// DB スプレッドシート ID の解決は Tenant.gs（resolveSpreadsheetId_ / getSs_）に一本化。
// 以下は個人ごとの設定キー（tGetProp_/tSetProp_ 経由で UserProperties に保存される）
const PROP_IMAGE_FOLDER_ID = 'IMAGE_FOLDER_ID';
const PROP_ADMIN_EMAIL     = 'ADMIN_EMAIL';

const JOURNAL_HEADERS = [
  'journalId', 'timestamp', 'email', 'theme', 'content',
  'imageFileId', 'emotion', 'teacherComment', 'highlights',
  'teacherStamp', 'status', 'pastComment', 'deletedAt'
];
const ROSTER_HEADERS = ['役割', '氏名', 'メールアドレス'];
const THEME_HEADERS = ['日付', 'テーマ'];

function doGet(e) {
  // 旧バインド環境からの移行ブリッジ: バインド先があれば個別紐付けへ引き継ぐ
  try {
    const bound = SpreadsheetApp.getActiveSpreadsheet();
    if (bound) {
      PropertiesService.getScriptProperties().setProperty(SP_KEY_LEGACY_SPREADSHEET_ID, bound.getId());
      try { if (!getUserSpreadsheetId_()) setUserSpreadsheetId_(bound.getId()); } catch (e2) {}
    }
  } catch (err) {}

  let userEmail = '';
  try { userEmail = Session.getActiveUser().getEmail() || ''; } catch (err) {}

  const tenant = getTenantStatus();

  // DB 未接続時はオンボーディング用の最小データのみ渡す
  let dataObj = {
    user: null,
    isAdmin: false,
    todayTheme: '',
    journals: [],
    classRoster: [],
    tenant: tenant.success ? tenant : { success: true, linked: false, email: userEmail, canCreate: true, templateConfigured: false }
  };

  if (tenant.success && tenant.linked) {
    try {
      const adminEmail = tGetProp_(PROP_ADMIN_EMAIL);
      let isAdmin = !!adminEmail && userEmail === adminEmail;
      const userData = getUserData_(userEmail, isAdmin);
      // 名簿で担任になっているユーザーは自分のテナントの管理者として扱う
      if (!isAdmin && userData && userData.role === '担任') {
        isAdmin = true;
        try { if (!adminEmail && userEmail) tSetProp_(PROP_ADMIN_EMAIL, userEmail); } catch (e3) {}
      }

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

      dataObj = {
        user: userData,
        isAdmin: isAdmin,
        todayTheme: getTodayTheme(),
        journals: journals,
        classRoster: classRoster,
        tenant: tenant
      };
    } catch (err) {
      // DB は紐付いているが読めない → オンボーディングに戻して再設定を促す
      dataObj.tenant = { success: true, linked: false, email: userEmail, canCreate: true, templateConfigured: tenant.templateConfigured, error: err.message };
    }
  }

  const tmpl = HtmlService.createTemplateFromFile('index');
  // エラー対策: JSON化してHTML崩れを防ぐ
  tmpl.initialData = JSON.stringify(dataObj).replace(/</g, '\\u003c');

  return tmpl.evaluate()
    .setTitle('ふりかえりジャーナル')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    // GitHub Pages シェルが iframe 埋め込みするため必須。
    // GAS は特定オリジン限定ができないので ALLOWALL 一択（リスクは README に明記）
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl('https://drive.google.com/uc?id=1rJjk2hoVW64rVz0kb-fdARn7g02Q5rjI&.png');
}

// 🌟 画像フォルダ取得（ユーザーごとに UserProperties で管理・自動リカバリ対応）
function getImageFolder_() {
  const folderId = tGetProp_(PROP_IMAGE_FOLDER_ID);

  if (folderId) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch (e) {
      // フォルダが削除されている場合はリカバリ処理へ進む
      console.warn('Image folder not found. Recovering...');
    }
  }

  // フォルダ再作成（リカバリ）: 実行ユーザー本人の Drive に作られる
  const folder = DriveApp.createFolder('ふりかえりジャーナル_画像');
  tSetProp_(PROP_IMAGE_FOLDER_ID, folder.getId());
  return folder;
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
  const ss = getSs_();
  const sheet = ss.getSheetByName(ROSTER_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return isAdmin ? { role: '担任', name: '管理者', email: email } : null;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][2] === email) return { role: data[i][0] === '担任' ? '担任' : '児童', name: data[i][1], email: data[i][2] };
  }
  return isAdmin ? { role: '担任', name: '管理者', email: email } : null;
}

function getUserData(email) {
  const adminEmail = tGetProp_(PROP_ADMIN_EMAIL);
  return getUserData_(email, !!adminEmail && email === adminEmail);
}

function getClassRoster() {
  const ss = getSs_();
  const sheet = ss.getSheetByName(ROSTER_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  return data.filter(function(row) { return row[0] !== '担任' && row[2]; }).map(function(row) { return { name: row[1], email: row[2] }; });
}

// 担任権限ガード: 担任以外が呼ぶとエラーオブジェクトを返す（担任なら null）。
// メールアドレスが取得できない旧デプロイ構成では判定不能のため従来動作を維持する。
function teacherGuard_() {
  try {
    let email = '';
    try { email = Session.getActiveUser().getEmail() || ''; } catch (e) {}
    if (!email) return null;
    const adminEmail = tGetProp_(PROP_ADMIN_EMAIL);
    const userData = getUserData_(email, !!adminEmail && email === adminEmail);
    if (userData && userData.role === '担任') return null;
    return { success: false, message: 'この操作は担任（管理者）のみ実行できます。' };
  } catch (e) {
    return null;
  }
}

// journalId から行番号を高速検索（TextFinder は全行走査より大幅に速い）
function findJournalRowById_(sheet, journalId) {
  if (!journalId || sheet.getLastRow() < 2) return -1;
  const cell = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1)
    .createTextFinder(String(journalId)).matchEntireCell(true).findNext();
  return cell ? cell.getRow() : -1;
}

function getJournalHeaders_(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function saveJournal(journalData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (!journalData || !String(journalData.content || '').trim()) {
      return { success: false, message: '本文が空です。ひとことでも書いてから提出してね。' };
    }
    // なりすまし防止: メールアドレスはサーバー側で取得（取得不能な旧構成のみクライアント値を利用）
    let email = '';
    try { email = Session.getActiveUser().getEmail() || ''; } catch (e) {}
    if (!email) email = journalData.email || '';

    const ss = getSs_();
    const sheet = ss.getSheetByName(JOURNAL_SHEET_NAME);
    if (!sheet) return { success: false, message: 'ジャーナルデータのシートが見つかりません。DBの設定を確認してください。' };
    const newRow = [
      Utilities.getUuid(), new Date(), email, journalData.theme, journalData.content,
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
  const ss = getSs_();
  const sheet = ss.getSheetByName(JOURNAL_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const emailIdx = headers.indexOf('email');
  const deletedAtIdx = headers.indexOf('deletedAt');
  return data.filter(function(row) { return row[emailIdx] === email && !row[deletedAtIdx]; }).map(function(row) { return rowToJournalObject_(headers, row); });
}

function getJournalsForTeacher() {
  const ss = getSs_();
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
  const guard = teacherGuard_(); if (guard) return guard;
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSs_().getSheetByName(JOURNAL_SHEET_NAME);
    const headers = getJournalHeaders_(sheet);
    const rowNum = findJournalRowById_(sheet, feedbackData.journalId);
    if (rowNum < 0) return { success: false, message: '該当するジャーナルが見つかりません。' };
    sheet.getRange(rowNum, headers.indexOf('teacherComment') + 1).setValue(feedbackData.comment);
    sheet.getRange(rowNum, headers.indexOf('highlights') + 1).setValue(feedbackData.highlights || '[]');
    sheet.getRange(rowNum, headers.indexOf('status') + 1).setValue('返却済み');
    return { success: true, message: 'フィードバックを保存しました！' };
  } catch (e) {
    return { success: false, message: 'エラー：' + e.message };
  } finally {
    lock.releaseLock();
  }
}

function revertJournalStatus(journalId) {
  const guard = teacherGuard_(); if (guard) return guard;
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSs_().getSheetByName(JOURNAL_SHEET_NAME);
    const rowNum = findJournalRowById_(sheet, journalId);
    if (rowNum < 0) return { success: false, message: '見つかりませんでした。' };
    sheet.getRange(rowNum, getJournalHeaders_(sheet).indexOf('status') + 1).setValue('未返却');
    return { success: true };
  } catch (e) {
    return { success: false, message: 'エラー：' + e.message };
  } finally { lock.releaseLock(); }
}

function addPastComment(journalId, comment) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSs_().getSheetByName(JOURNAL_SHEET_NAME);
    const rowNum = findJournalRowById_(sheet, journalId);
    if (rowNum < 0) return { success: false, message: '見つかりませんでした。' };
    sheet.getRange(rowNum, getJournalHeaders_(sheet).indexOf('pastComment') + 1).setValue(comment);
    return { success: true, message: '保存しました！' };
  } catch (e) {
    return { success: false, message: 'エラー：' + e.message };
  } finally { lock.releaseLock(); }
}

function deleteJournal(journalId) {
  const guard = teacherGuard_(); if (guard) return guard;
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSs_().getSheetByName(JOURNAL_SHEET_NAME);
    const rowNum = findJournalRowById_(sheet, journalId);
    if (rowNum < 0) return { success: false, message: '見つかりませんでした。' };
    sheet.getRange(rowNum, getJournalHeaders_(sheet).indexOf('deletedAt') + 1).setValue(new Date());
    return { success: true, message: '削除しました。' };
  } catch (e) {
    return { success: false, message: 'エラー：' + e.message };
  } finally { lock.releaseLock(); }
}

// 🌟 自動リカバリ対応
function uploadImage(fileData) {
  try {
    if (!fileData || !fileData.data) return { success: false, message: '画像データがありません。' };
    // 目安 10MB 超は拒否（base64 は約 4/3 倍になるため 14MB 相当で判定）
    if (String(fileData.data).length > 14 * 1024 * 1024) {
      return { success: false, message: '画像が大きすぎます（10MBまで）。小さい画像でためしてね。' };
    }
    const folder = getImageFolder_(); // リカバリ対応の取得関数を使用
    // 注意: 第2引数に Charset を渡すとバイナリが壊れるため、素の base64Decode を使う
    const decoded = Utilities.base64Decode(fileData.data);
    const blob = Utilities.newBlob(decoded, fileData.mimeType, fileData.fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { success: true, fileId: file.getId(), fileName: file.getName() };
  } catch (e) {
    return { success: false, message: 'エラー：' + e.message };
  }
}

function getTodayTheme() {
  const ss = getSs_();
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
  const guard = teacherGuard_(); if (guard) return guard;
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSs_().getSheetByName(THEME_SHEET_NAME);
    if (!sheet) return { success: false, message: 'テーマ設定シートが見つかりません。' };
    sheet.appendRow([new Date(), theme]);
    return { success: true, message: 'テーマを設定しました！' };
  } catch (e) {
    return { success: false, message: 'エラー：' + e.message };
  } finally { lock.releaseLock(); }
}

function getWeeklyThemes() {
  try {
    const json = tGetProp_('WEEKLY_THEMES');
    return { success: true, data: json ? JSON.parse(json) : { mon: '', tue: '', wed: '', thu: '', fri: '' } };
  } catch (e) { return { success: false, message: 'Error: ' + e.message }; }
}

function saveWeeklyThemes(themes) {
  try {
    // 個人ごとの設定なので UserProperties に保存（全ユーザー共有を防ぐ）
    tSetProp_('WEEKLY_THEMES', JSON.stringify(themes));
    return { success: true, message: '保存しました！' };
  } catch (e) { return { success: false, message: 'Error: ' + e.message }; }
}

function getWeeklyThemeForToday_() {
  try {
    const json = tGetProp_('WEEKLY_THEMES');
    if (!json) return null;
    const themes = JSON.parse(json);
    const dayKeys = ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat'];
    return themes[dayKeys[new Date().getDay()]] || null;
  } catch (e) {
    return null; // 設定が壊れていてもアプリ全体は止めない
  }
}

function quickReturn(journalId, stamp) {
  const guard = teacherGuard_(); if (guard) return guard;
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSs_().getSheetByName(JOURNAL_SHEET_NAME);
    const headers = getJournalHeaders_(sheet);
    const rowNum = findJournalRowById_(sheet, journalId);
    if (rowNum < 0) return { success: false, message: '見つかりませんでした。' };
    sheet.getRange(rowNum, headers.indexOf('teacherStamp') + 1).setValue(stamp);
    sheet.getRange(rowNum, headers.indexOf('status') + 1).setValue('返却済み');
    return { success: true, message: 'スタンプで返却しました！' };
  } catch (e) {
    return { success: false, message: 'エラー：' + e.message };
  } finally { lock.releaseLock(); }
}

function batchReturnAll() {
  const guard = teacherGuard_(); if (guard) return guard;
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const sheet = getSs_().getSheetByName(JOURNAL_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const statusIdx = headers.indexOf('status');
    const commentIdx = headers.indexOf('teacherComment');
    const deletedIdx = headers.indexOf('deletedAt');
    let count = 0;
    for (let i = 1; i < data.length; i++) {
      if (data[i][statusIdx] === '未返却' && data[i][commentIdx] && !data[i][deletedIdx]) {
        sheet.getRange(i + 1, statusIdx + 1).setValue('返却済み');
        count++;
      }
    }
    return { success: true, count: count, message: count > 0 ? count + '件を一括返却しました！' : '返却対象がありません。' };
  } catch (e) {
    return { success: false, message: 'エラー：' + e.message };
  } finally { lock.releaseLock(); }
}

// AI処理対象（未返却×本文あり×未削除）の抽出
function collectAiTargets_(sheet) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idx = {
    status: headers.indexOf('status'),
    content: headers.indexOf('content'),
    deletedAt: headers.indexOf('deletedAt'),
    teacherComment: headers.indexOf('teacherComment'),
    highlights: headers.indexOf('highlights')
  };
  const targets = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][idx.status] === '未返却' && data[i][idx.content] && !data[i][idx.deletedAt]) {
      targets.push({ rowNum: i + 1, content: String(data[i][idx.content]) });
    }
  }
  return { targets: targets, idx: idx };
}

// fetchAll をチャンク実行（直列比で大幅に高速。レート制限も考慮して小分けにする）
function fetchAllInChunks_(requests, chunkSize) {
  const responses = [];
  for (let i = 0; i < requests.length; i += chunkSize) {
    const chunk = UrlFetchApp.fetchAll(requests.slice(i, i + chunkSize));
    for (let j = 0; j < chunk.length; j++) responses.push(chunk[j]);
  }
  return responses;
}

const AI_SIMPLE_PROMPT = 'あなたは児童の小さな頑張りやユニークな視点を見つけて具体的に褒めるのが得意な、経験豊富な小学校の先生です。以下の記述を読み、児童が努力した点などを引用しつつ、自己肯定感を育む温かい賞賛のコメントを100字程度で作成してください。見出しや解説は不要です。\n\n';

const AI_FULL_PROMPT = 'あなたは経験豊富な小学校の先生です。以下の児童のジャーナルを読み、フィードバックを作成してください。\n' +
  '# 出力形式の厳密なルール\n' +
  '必ず以下の構造を持つJSONのみを出力してください。\n' +
  '{"comment":"（全体への温かいコメント100字以内）","highlights":[{"textToHighlight":"（本文から完全一致で引用）","suggestedComment":"（ハイライト箇所へのコメント）","suggestedStamp":"（絵文字1つ）"}]}\n' +
  '---\n';

function generateAiSimpleCommentsForAll() {
  const guard = teacherGuard_(); if (guard) return guard;
  try {
    const apiKey = tGetProp_('GEMINI_API_KEY');
    if (!apiKey) return { success: false, message: 'Gemini APIキーが設定されていません。' };
    const sheet = getSs_().getSheetByName(JOURNAL_SHEET_NAME);
    const collected = collectAiTargets_(sheet);
    if (collected.targets.length === 0) return { success: true, message: '対象（未返却×本文あり）のジャーナルがありません。' };

    const requests = collected.targets.map(function (t) {
      return {
        url: API_ENDPOINT_V1 + apiKey, method: 'post', contentType: 'application/json',
        payload: JSON.stringify({ contents: [{ parts: [{ text: AI_SIMPLE_PROMPT + t.content }] }] }),
        muteHttpExceptions: true
      };
    });
    const responses = fetchAllInChunks_(requests, 8);

    let ok = 0, ng = 0;
    responses.forEach(function (res, i) {
      try {
        if (res.getResponseCode() === 200) {
          const comment = JSON.parse(res.getContentText()).candidates[0].content.parts[0].text.trim();
          sheet.getRange(collected.targets[i].rowNum, collected.idx.teacherComment + 1).setValue(comment);
          ok++;
        } else { ng++; }
      } catch (e) { ng++; }
    });
    return { success: true, message: 'AIコメント案を作成しました。（成功: ' + ok + '件' + (ng ? '、失敗: ' + ng + '件' : '') + '）' };
  } catch (e) {
    return { success: false, message: 'AI処理エラー：' + e.message };
  }
}

function generateAiFullFeedbackForAll() {
  const guard = teacherGuard_(); if (guard) return guard;
  try {
    const apiKey = tGetProp_('GEMINI_API_KEY');
    if (!apiKey) return { success: false, message: 'Gemini APIキーが設定されていません。' };
    const sheet = getSs_().getSheetByName(JOURNAL_SHEET_NAME);
    const collected = collectAiTargets_(sheet);
    if (collected.targets.length === 0) return { success: true, message: '対象（未返却×本文あり）のジャーナルがありません。' };

    const requests = collected.targets.map(function (t) {
      return {
        url: API_ENDPOINT_V1_BETA + apiKey, method: 'post', contentType: 'application/json',
        payload: JSON.stringify({ contents: [{ parts: [{ text: AI_FULL_PROMPT + t.content }] }], generationConfig: { responseMimeType: 'application/json' } }),
        muteHttpExceptions: true
      };
    });
    const responses = fetchAllInChunks_(requests, 8);

    let successCount = 0, errorCount = 0;
    responses.forEach(function (res, i) {
      const target = collected.targets[i];
      try {
        if (res.getResponseCode() !== 200) { errorCount++; return; }
        const jsonText = JSON.parse(res.getContentText()).candidates[0].content.parts[0].text;
        const feedback = JSON.parse(jsonText);
        if (feedback.comment) {
          sheet.getRange(target.rowNum, collected.idx.teacherComment + 1).setValue(feedback.comment);
        }
        if (feedback.highlights && feedback.highlights.length > 0) {
          const content = target.content;
          const hlToSave = [];
          feedback.highlights.forEach(function (h) {
            const startIndex = content.indexOf(h.textToHighlight);
            if (startIndex !== -1) {
              hlToSave.push({
                id: 'hl-' + Date.now() + Math.random(), textToHighlight: h.textToHighlight,
                suggestedComment: h.suggestedComment || '', suggestedStamp: h.suggestedStamp || '',
                startOffset: startIndex, endOffset: startIndex + h.textToHighlight.length
              });
            }
          });
          if (hlToSave.length > 0) sheet.getRange(target.rowNum, collected.idx.highlights + 1).setValue(JSON.stringify(hlToSave));
        }
        successCount++;
      } catch (e) { errorCount++; }
    });
    return { success: true, message: 'AI高度分析完了。（成功: ' + successCount + '件, 失敗: ' + errorCount + '件）' };
  } catch (e) { return { success: false, message: 'AIエラー：' + e.message }; }
}

function getRosterAll() {
  try {
    const sheet = getSs_().getSheetByName(ROSTER_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, data: [] };
    const roster = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues().filter(function(row) { return row[2]; }).map(function(row) { return { role: row[0], name: row[1], email: row[2] }; });
    return { success: true, data: roster };
  } catch (e) { return { success: false, message: e.message }; }
}

function saveRosterAll(rows) {
  const guard = teacherGuard_(); if (guard) return guard;
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSs_().getSheetByName(ROSTER_SHEET_NAME);
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
    const apiKey = tGetProp_('GEMINI_API_KEY') || '';
    const ssId = resolveSpreadsheetId_();
    const folderId = tGetProp_(PROP_IMAGE_FOLDER_ID) || '';
    let maskedKey = apiKey.length > 10 ? apiKey.substring(0, 6) + '****' + apiKey.substring(apiKey.length - 4) : (apiKey ? '設定済み' : '');
    return { success: true, apiKeyMasked: maskedKey, hasApiKey: !!apiKey, spreadsheetUrl: ssId ? 'https://docs.google.com/spreadsheets/d/' + ssId : '', imageFolderUrl: folderId ? 'https://drive.google.com/drive/folders/' + folderId : '' };
  } catch (e) { return { success: false, message: e.message }; }
}

function saveAdminSettings(settings) {
  try {
    // APIキーは個人の秘密情報なので UserProperties に保存する
    if (settings.apiKey !== undefined && settings.apiKey !== '') tSetProp_('GEMINI_API_KEY', settings.apiKey);
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
  const guard = teacherGuard_(); if (guard) return guard;
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSs_().getSheetByName(JOURNAL_SHEET_NAME);
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
    return { success: true, message: '全データを削除しました。' };
  } catch (e) {
    return { success: false, message: 'エラー：' + e.message };
  } finally { lock.releaseLock(); }
}

function getSubmissionStatus() {
  try {
    const ss = getSs_();
    const roster = getClassRoster();
    const today = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd');
    const jSheet = ss.getSheetByName(JOURNAL_SHEET_NAME);
    const submittedEmails = new Set();

    if (jSheet && jSheet.getLastRow() >= 2) {
      const data = jSheet.getDataRange().getValues();
      const headers = data.shift();
      const deletedIdx = headers.indexOf('deletedAt');
      const tsIdx = headers.indexOf('timestamp');
      const emailIdx = headers.indexOf('email');
      data.forEach(function(row) {
        if (row[deletedIdx]) return;
        if (row[tsIdx] instanceof Date && Utilities.formatDate(row[tsIdx], 'JST', 'yyyy/MM/dd') === today) {
          submittedEmails.add(row[emailIdx]);
        }
      });
    }

    const students = roster.map(function(s) { return { name: s.name, email: s.email, submitted: submittedEmails.has(s.email) }; });
    return { success: true, data: { total: students.length, submitted: students.filter(function(s) { return s.submitted; }).length, students: students } };
  } catch (e) { return { success: false, message: e.message }; }
}

// CSVインジェクション対策: 数式として解釈されうる先頭文字はシングルクォートで無害化
function csvSafe_(v) {
  const s = String(v == null ? '' : v);
  return /^[=+\-@\t]/.test(s) ? "'" + s : s;
}

function exportJournalsCsv(params) {
  const guard = teacherGuard_(); if (guard) return guard;
  try {
    const journals = getFilteredJournals_(params);
    if (journals.length === 0) return { success: false, message: 'データがありません。' };

    const csvHeaders = ['日付', '氏名', 'テーマ', '本文', '気持ち', '先生のコメント', 'ステータス'];
    const csvRows = journals.map(function(j) { return [j.date||'', j.studentName||'', j.theme||'', j.content||'', j.emotion||'', j.teacherComment||'', j.status||'']; });
    const csvContent = [csvHeaders].concat(csvRows).map(function(row) { return row.map(function(cell) { return '"' + csvSafe_(cell).replace(/"/g, '""') + '"'; }).join(','); }).join('\r\n');
    
    const fileName = 'ジャーナルデータ_' + Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss') + '.csv';
    const file = getExportFolder_().createFile(Utilities.newBlob('\uFEFF' + csvContent, 'text/csv', fileName));
    return { success: true, fileUrl: file.getUrl(), fileName: fileName, message: 'CSV出力完了' };
  } catch (e) { return { success: false, message: e.message }; }
}

function exportJournalsPdf(params) {
  const guard = teacherGuard_(); if (guard) return guard;
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
  const ss = getSs_();
  const sheet = ss.getSheetByName(JOURNAL_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const deletedIdx = headers.indexOf('deletedAt');
  const emailIdx = headers.indexOf('email');
  const tsIdx = headers.indexOf('timestamp');
  const nameMap = getClassRoster().reduce(function(map, u) { map[u.email] = u.name; return map; }, {});
  const startDate = params.startDate ? new Date(params.startDate + 'T00:00:00+09:00') : null;
  const endDate = params.endDate ? new Date(params.endDate + 'T23:59:59+09:00') : null;
  const filterEmail = (params.email && params.email !== 'all') ? params.email.trim().toLowerCase() : null;

  return data.filter(function(row) {
    if (row[deletedIdx]) return false;
    if (filterEmail && String(row[emailIdx]).trim().toLowerCase() !== filterEmail) return false;
    if (row[tsIdx] instanceof Date) {
      if (startDate && row[tsIdx] < startDate) return false;
      if (endDate && row[tsIdx] > endDate) return false;
    }
    return true;
  }).map(function(row) {
    const j = rowToJournalObject_(headers, row);
    j.studentName = nameMap[row[emailIdx]] || '不明';
    return j;
  }).sort(function(a, b) { return (a.studentName||'').localeCompare(b.studentName||'') || (a.timestamp||0) - (b.timestamp||0); });
}

// 🌟 自動リカバリ対応: エクスポートフォルダ
function getExportFolder_() {
  const pFolder = getImageFolder_(); // リカバリ対応の取得関数を使用
  const parent = pFolder.getParents().hasNext() ? pFolder.getParents().next() : DriveApp.getRootFolder();
  const folders = parent.getFoldersByName('エクスポート');
  return folders.hasNext() ? folders.next() : parent.createFolder('エクスポート');
}
