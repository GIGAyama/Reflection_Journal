/**
 * OwnerApi.gs — デプロイ A（先生ポータル）用 API
 *
 * A は「アクセスしているユーザーとして実行」なので、本人特定は Session.getActiveUser()。
 * すべての API は { success:true, ... } / { success:false, code, error } の JSON 文字列を返す。
 * クラスを操作する API は冒頭で「そのクラスの ownerEmail が自分か」を必ず検証する。
 */

/** クラス所有者チェック（先生 API 共通ガード） */
function assertOwner_(tenantCode) {
  const email = ownerEmail_();
  const code = normalizeTenantCode_(tenantCode);
  const rec = getTenantRecord_(code);
  if (!rec || rec.revoked) {
    throw new Error('TENANT_NOT_FOUND: このクラスは存在しないか、すでに閉じられています');
  }
  if (String(rec.ownerEmail).toLowerCase() !== email) {
    throw new Error('FORBIDDEN: このクラスを管理できるのは作成した先生本人だけです');
  }
  return { email: email, code: code, rec: rec };
}

function memberUrlFor_(code) {
  const shell = getShellUrl_();
  return shell ? shell + '?t=' + code : '';
}

/** スプレッドシートの URL / 生ID のどちらを渡されても ID を抽出する */
function extractSpreadsheetId_(input) {
  if (!input) return '';
  const str = String(input).trim();
  const m = str.match(/\/spreadsheets\/d\/([a-zA-Z0-9\-_]+)/);
  if (m) return m[1];
  if (/^[a-zA-Z0-9\-_]{20,}$/.test(str)) return str;
  return '';
}

/** シート生成後の共通登録処理（コード発行 → レジストリ → メタ → 名簿） */
function registerTenant_(ss, email, name) {
  let code = null;
  withScriptLock_(function () {
    code = generateTenantCode_();
    putTenantRecord_(code, {
      spreadsheetId: ss.getId(),
      ownerEmail: email,
      tenantName: name,
      createdAt: new Date().toISOString(),
      joinOpen: true,
      requireApproval: true,   // コードは宛先であって認証ではない。既定は承認制
      revoked: false,
      memberCount: 1
    });
  });
  addOwnedCode_(email, code);
  writeMeta_(ss, {
    schemaVersion: CONFIG.SCHEMA_VERSION,
    tenantCode: code,
    tenantName: name,
    ownerEmail: email,
    createdAt: new Date().toISOString()
  });
  upsertMember_(ss, { email: email, name: '先生', role: '担任', status: 'active' });
  return code;
}

/**
 * ★ この構成の心臓部 ★
 * クラス作成。先生本人の実行なのでシートは最初から先生所有になる。
 *
 * DriveApp.makeCopy は使わない（フル Drive スコープ回避）。テンプレートは
 * SpreadsheetApp.openById(templateId).copy() で複製する（spreadsheets スコープで動く）。
 */
function opCreateTenant(tenantName) {
  const userLock = LockService.getUserLock();
  try {
    const email = ownerEmail_();
    const name = vStr_(tenantName, 50, 'クラス名').trim();
    if (!name) throw new Error('BAD_INPUT: クラス名を入力してください');
    const appAccount = String(getSetting_(PROP_KEYS.APP_ACCOUNT, true)).toLowerCase();

    if (!userLock.tryLock(CONFIG.LOCK_TIMEOUT_MS)) {
      throw new Error('LOCK_BUSY: 処理が混み合っています。数秒待ってもう一度お試しください');
    }

    // 1. シート生成（テンプレートがあれば複製、無ければ新規作成 + 初期化）
    const templateId = getSetting_(PROP_KEYS.TEMPLATE, false);
    const title = CONFIG.APP_NAME + '_' + name;
    let ss;
    if (templateId) {
      ss = SpreadsheetApp.openById(templateId).copy(title);
      ensureTenantSheets_(ss);
    } else {
      ss = SpreadsheetApp.create(title);
      initializeNewDatabase_(ss);
    }

    // 2. アプリアカウントを編集者に自動追加（この構成の心臓部）。
    //    これにより児童用デプロイ B（アプリアカウント実行）がこのシートに読み書きできる。
    //    児童自身には権限を一切与えない。
    try {
      ss.addEditor(appAccount);
    } catch (shareErr) {
      // 巻き戻し: Drive スコープを持たないためゴミ箱移動はできない。
      // 名前で失敗が分かるようにし、レジストリには登録しない（クラスとして成立させない）。
      try { ss.rename('【作成失敗・削除してください】' + title); } catch (e2) {}
      throw new Error('SHARE_FAILED: シートは作成できましたが、アプリへの共有に失敗しました。' +
        'Google Workspace で外部共有が制限されている可能性があります。組織の管理者に「' +
        appAccount + ' への共有許可」を確認してください。' +
        '（ドライブに残った「【作成失敗・削除してください】」のシートは削除して構いません）');
    }

    // 3. コード発行 → レジストリ登録 → 名簿に自分を担任/active で登録
    const code = registerTenant_(ss, email, name);

    return jsonOk_({
      tenantCode: code,
      tenantName: name,
      memberUrl: memberUrlFor_(code),
      spreadsheetUrl: ss.getUrl()
    });
  } catch (e) {
    return jsonErr_(e);
  } finally {
    try { userLock.releaseLock(); } catch (e) {}
  }
}

/**
 * 既存スプレッドシート（旧バージョンの DB 等）をクラスとして取り込む。
 * 先生本人が編集権限を持つシートであることが前提（アクセスユーザー実行なので
 * 開けない場合はここで失敗する）。必須シート・状態列は自動で補完する。
 */
function opAdoptTenant(input, tenantName) {
  const userLock = LockService.getUserLock();
  try {
    const email = ownerEmail_();
    const name = vStr_(tenantName, 50, 'クラス名').trim();
    if (!name) throw new Error('BAD_INPUT: クラス名を入力してください');
    const id = extractSpreadsheetId_(input);
    if (!id) throw new Error('BAD_INPUT: スプレッドシートの URL または ID を正しく入力してください');
    const appAccount = String(getSetting_(PROP_KEYS.APP_ACCOUNT, true)).toLowerCase();

    if (!userLock.tryLock(CONFIG.LOCK_TIMEOUT_MS)) {
      throw new Error('LOCK_BUSY: 処理が混み合っています。数秒待ってもう一度お試しください');
    }

    let ss;
    try {
      ss = SpreadsheetApp.openById(id);
      ss.getName(); // アクセスできるかの実チェック
    } catch (e) {
      throw new Error('BAD_INPUT: スプレッドシートを開けませんでした。URL/ID が正しいか、あなたのアカウントに編集権限があるか確認してください');
    }

    ensureTenantSheets_(ss);
    try {
      ss.addEditor(appAccount);
    } catch (shareErr) {
      throw new Error('SHARE_FAILED: アプリ（' + appAccount + '）への共有に失敗しました。' +
        '組織の外部共有制限を確認してください');
    }

    const code = registerTenant_(ss, email, name);
    return jsonOk_({
      tenantCode: code,
      tenantName: name,
      memberUrl: memberUrlFor_(code),
      spreadsheetUrl: ss.getUrl()
    });
  } catch (e) {
    return jsonErr_(e);
  } finally {
    try { userLock.releaseLock(); } catch (e) {}
  }
}

/** 自分のクラス一覧。memberUrl は共通 URL（GitHub Pages）+ ?t=コード */
function opListTenants() {
  try {
    const email = ownerEmail_();
    const tenants = listOwnedCodes_(email).map(function (code) {
      const rec = getTenantRecord_(code);
      if (!rec || rec.revoked) return null;
      return {
        tenantCode: code,
        tenantName: rec.tenantName,
        memberUrl: memberUrlFor_(code),
        createdAt: rec.createdAt,
        joinOpen: rec.joinOpen,
        requireApproval: rec.requireApproval,
        memberCount: rec.memberCount || 0
      };
    }).filter(Boolean);
    return jsonOk_({ tenants: tenants, ownerEmail: email, shellUrl: getShellUrl_() });
  } catch (e) {
    return jsonErr_(e);
  }
}

/** クラスコードの再発行。旧コードの URL は無効になる */
function opRegenerateCode(tenantCode) {
  try {
    const ctx = assertOwner_(tenantCode);
    let newCode = null;
    withScriptLock_(function () {
      newCode = generateTenantCode_();
      putTenantRecord_(newCode, ctx.rec);
      deleteTenantRecord_(ctx.code);
    });
    removeOwnedCode_(ctx.email, ctx.code);
    addOwnedCode_(ctx.email, newCode);
    try { writeMeta_(openTenantSs_(newCode), { tenantCode: newCode }); } catch (e) {}
    return jsonOk_({ tenantCode: newCode, memberUrl: memberUrlFor_(newCode) });
  } catch (e) {
    return jsonErr_(e);
  }
}

/** 参加受付の開閉・承認制の切り替え */
function opUpdateJoinPolicy(tenantCode, joinOpen, requireApproval) {
  try {
    const ctx = assertOwner_(tenantCode);
    withScriptLock_(function () {
      updateTenantRecord_(ctx.code, {
        joinOpen: !!joinOpen,
        requireApproval: !!requireApproval
      });
    });
    return jsonOk_({});
  } catch (e) {
    return jsonErr_(e);
  }
}

/** クラスを閉じる（レジストリから外す。シートは先生の手元に残る） */
function opRevokeTenant(tenantCode) {
  try {
    const ctx = assertOwner_(tenantCode);
    withScriptLock_(function () {
      updateTenantRecord_(ctx.code, { revoked: true });
    });
    removeOwnedCode_(ctx.email, ctx.code);
    return jsonOk_({ message: 'クラスを閉じました。スプレッドシートはあなたのドライブに残っています。' });
  } catch (e) {
    return jsonErr_(e);
  }
}

function refreshMemberCount_(code, ss) {
  withScriptLock_(function () {
    updateTenantRecord_(code, {
      memberCount: getRosterRows_(ss).filter(function (m) { return m.status === 'active'; }).length
    });
  });
}

/** 参加申請の承認 */
function opApproveMember(tenantCode, memberEmail) {
  try {
    const ctx = assertOwner_(tenantCode);
    const ss = openTenantSs_(ctx.code);
    const target = String(memberEmail || '').toLowerCase();
    const row = getMemberRow_(ss, target);
    if (!row) throw new Error('NOT_MEMBER: その申請は見つかりません');
    upsertMember_(ss, { email: target, name: row.name, role: row.role, status: 'active' });
    refreshMemberCount_(ctx.code, ss);
    return jsonOk_({});
  } catch (e) {
    return jsonErr_(e);
  }
}

/** 参加申請の却下 / 名簿からの削除 */
function opRejectMember(tenantCode, memberEmail) {
  try {
    const ctx = assertOwner_(tenantCode);
    const ss = openTenantSs_(ctx.code);
    removeMember_(ss, String(memberEmail || '').toLowerCase());
    refreshMemberCount_(ctx.code, ss);
    return jsonOk_({});
  } catch (e) {
    return jsonErr_(e);
  }
}

// ────────────────────────────────────────────────────────────────
// クラス運用 API（ジャーナル・名簿・テーマ・設定）
// ────────────────────────────────────────────────────────────────

/** ポータルの初期同期。選択中クラスの全データを返す（先生には email を含めて返してよい） */
function opGetTenantData(tenantCode) {
  try {
    const ctx = assertOwner_(tenantCode);
    const ss = openTenantSs_(ctx.code);
    const roster = getRosterRows_(ss);
    const nameMap = {};
    roster.forEach(function (m) { nameMap[m.email] = m.name; });
    const journals = getJournalsAll_(ss).map(function (j) {
      j.studentName = nameMap[String(j.email).toLowerCase()] || '不明';
      return j;
    });
    const apiKey = getTenantSetting_(ss, 'GEMINI_API_KEY');
    return jsonOk_({
      tenant: {
        tenantCode: ctx.code,
        tenantName: ctx.rec.tenantName,
        memberUrl: memberUrlFor_(ctx.code),
        joinOpen: ctx.rec.joinOpen,
        requireApproval: ctx.rec.requireApproval
      },
      journals: journals,
      classRoster: roster.filter(function (m) { return m.role !== '担任' && m.status === 'active'; }),
      pendingMembers: roster.filter(function (m) { return m.status === 'pending'; }),
      rosterAll: roster,
      todayTheme: getTodayTheme_(ss),
      weeklyThemes: getWeeklyThemes_(ss),
      settings: {
        hasApiKey: !!apiKey,
        apiKeyMasked: apiKey.length > 10
          ? apiKey.substring(0, 6) + '****' + apiKey.substring(apiKey.length - 4)
          : (apiKey ? '設定済み' : '')
      },
      spreadsheetUrl: 'https://docs.google.com/spreadsheets/d/' + ctx.rec.spreadsheetId
    });
  } catch (e) {
    return jsonErr_(e);
  }
}

/** 名簿の一括保存。payload はホワイトリストしたキーのみ書き込む */
function opSaveRoster(tenantCode, rows) {
  try {
    const ctx = assertOwner_(tenantCode);
    const ss = openTenantSs_(ctx.code);
    const clean = (rows || []).map(function (r) {
      return {
        role: r && r.role === '担任' ? '担任' : '児童',
        name: vStr_(r && r.name, 50, '氏名').trim(),
        email: vStr_(r && r.email, 100, 'メールアドレス').trim().toLowerCase(),
        status: r && r.status === 'pending' ? 'pending' : 'active'
      };
    }).filter(function (r) { return r.name && r.email; });
    // 自分（担任）が名簿から消えないように保証する
    if (!clean.some(function (r) { return r.email === ctx.email; })) {
      clean.unshift({ role: '担任', name: '先生', email: ctx.email, status: 'active' });
    }
    withScriptLock_(function () { saveRosterRows_(ss, clean); });
    refreshMemberCount_(ctx.code, ss);
    return jsonOk_({ message: '名簿を保存しました' });
  } catch (e) {
    return jsonErr_(e);
  }
}

function opSetTodayTheme(tenantCode, theme) {
  try {
    const ctx = assertOwner_(tenantCode);
    const ss = openTenantSs_(ctx.code);
    const t = vStr_(theme, 200, 'テーマ').trim();
    if (!t) throw new Error('BAD_INPUT: テーマを入力してください');
    const sheet = createSheetIfNotExists_(ss, THEME_SHEET_NAME, THEME_HEADERS);
    withScriptLock_(function () { sheet.appendRow([new Date(), t]); });
    return jsonOk_({ message: 'テーマを設定しました！' });
  } catch (e) {
    return jsonErr_(e);
  }
}

function opGetWeeklyThemes(tenantCode) {
  try {
    const ctx = assertOwner_(tenantCode);
    return jsonOk_({ data: getWeeklyThemes_(openTenantSs_(ctx.code)) });
  } catch (e) {
    return jsonErr_(e);
  }
}

function opSaveWeeklyThemes(tenantCode, themes) {
  try {
    const ctx = assertOwner_(tenantCode);
    const ss = openTenantSs_(ctx.code);
    const clean = {};
    ['mon', 'tue', 'wed', 'thu', 'fri'].forEach(function (d) {
      clean[d] = vStr_(themes && themes[d], 200, 'テーマ');
    });
    setTenantSetting_(ss, 'WEEKLY_THEMES', JSON.stringify(clean));
    return jsonOk_({ message: '保存しました！' });
  } catch (e) {
    return jsonErr_(e);
  }
}

/** highlights は JSON 配列文字列。ホワイトリストしたキーのみ通す */
function cleanHighlights_(raw) {
  let arr = [];
  try { arr = JSON.parse(String(raw || '[]')); } catch (e) { arr = []; }
  if (!Array.isArray(arr)) arr = [];
  return JSON.stringify(arr.slice(0, 30).map(function (h) {
    return {
      id: vStr_(h && h.id, 60, 'ID'),
      textToHighlight: vStr_(h && h.textToHighlight, 500, 'ハイライト'),
      suggestedComment: vStr_(h && h.suggestedComment, 500, 'コメント'),
      suggestedStamp: vStr_(h && h.suggestedStamp, 10, 'スタンプ'),
      startOffset: h && isFinite(h.startOffset) ? Number(h.startOffset) : 0,
      endOffset: h && isFinite(h.endOffset) ? Number(h.endOffset) : 0
    };
  }));
}

function opSaveFeedback(tenantCode, feedbackData) {
  try {
    const ctx = assertOwner_(tenantCode);
    const ss = openTenantSs_(ctx.code);
    const sheet = getJournalSheet_(ss);
    const d = feedbackData || {};
    return withScriptLock_(function () {
      const headers = getJournalHeaders_(sheet);
      const rowNum = findJournalRowById_(sheet, vJournalId_(d.journalId));
      if (rowNum < 0) throw new Error('NOT_FOUND: 該当するジャーナルが見つかりません');
      sheet.getRange(rowNum, headers.indexOf('teacherComment') + 1).setValue(vStr_(d.comment, 2000, 'コメント'));
      sheet.getRange(rowNum, headers.indexOf('highlights') + 1).setValue(cleanHighlights_(d.highlights));
      sheet.getRange(rowNum, headers.indexOf('status') + 1).setValue('返却済み');
      return jsonOk_({ message: 'フィードバックを保存しました！' });
    });
  } catch (e) {
    return jsonErr_(e);
  }
}

function opQuickReturn(tenantCode, journalId, stamp) {
  try {
    const ctx = assertOwner_(tenantCode);
    const ss = openTenantSs_(ctx.code);
    const sheet = getJournalSheet_(ss);
    return withScriptLock_(function () {
      const headers = getJournalHeaders_(sheet);
      const rowNum = findJournalRowById_(sheet, vJournalId_(journalId));
      if (rowNum < 0) throw new Error('NOT_FOUND: 見つかりませんでした');
      sheet.getRange(rowNum, headers.indexOf('teacherStamp') + 1).setValue(vStr_(stamp, 10, 'スタンプ'));
      sheet.getRange(rowNum, headers.indexOf('status') + 1).setValue('返却済み');
      return jsonOk_({ message: 'スタンプで返却しました！' });
    });
  } catch (e) {
    return jsonErr_(e);
  }
}

function opRevertStatus(tenantCode, journalId) {
  try {
    const ctx = assertOwner_(tenantCode);
    const ss = openTenantSs_(ctx.code);
    const sheet = getJournalSheet_(ss);
    return withScriptLock_(function () {
      const rowNum = findJournalRowById_(sheet, vJournalId_(journalId));
      if (rowNum < 0) throw new Error('NOT_FOUND: 見つかりませんでした');
      sheet.getRange(rowNum, getJournalHeaders_(sheet).indexOf('status') + 1).setValue('未返却');
      return jsonOk_({});
    });
  } catch (e) {
    return jsonErr_(e);
  }
}

function opDeleteJournal(tenantCode, journalId) {
  try {
    const ctx = assertOwner_(tenantCode);
    const ss = openTenantSs_(ctx.code);
    const sheet = getJournalSheet_(ss);
    return withScriptLock_(function () {
      const rowNum = findJournalRowById_(sheet, vJournalId_(journalId));
      if (rowNum < 0) throw new Error('NOT_FOUND: 見つかりませんでした');
      sheet.getRange(rowNum, getJournalHeaders_(sheet).indexOf('deletedAt') + 1).setValue(new Date());
      return jsonOk_({ message: '削除しました。' });
    });
  } catch (e) {
    return jsonErr_(e);
  }
}

function opBatchReturnAll(tenantCode) {
  try {
    const ctx = assertOwner_(tenantCode);
    const ss = openTenantSs_(ctx.code);
    const sheet = getJournalSheet_(ss);
    return withScriptLock_(function () {
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
      return jsonOk_({ count: count, message: count > 0 ? count + '件を一括返却しました！' : '返却対象がありません。' });
    });
  } catch (e) {
    return jsonErr_(e);
  }
}

function opResetData(tenantCode) {
  try {
    const ctx = assertOwner_(tenantCode);
    const ss = openTenantSs_(ctx.code);
    return withScriptLock_(function () {
      const sheet = getJournalSheet_(ss);
      if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
      const imgSheet = ss.getSheetByName(IMAGE_SHEET_NAME);
      if (imgSheet && imgSheet.getLastRow() > 1) imgSheet.deleteRows(2, imgSheet.getLastRow() - 1);
      return jsonOk_({ message: '全データを削除しました。' });
    });
  } catch (e) {
    return jsonErr_(e);
  }
}

/** 先生用の画像取得（担任は全児童の画像を閲覧できる） */
function opGetImage(tenantCode, imageId) {
  try {
    const ctx = assertOwner_(tenantCode);
    const img = loadImage_(openTenantSs_(ctx.code), imageId);
    if (!img) throw new Error('NOT_FOUND: 画像が見つかりません');
    return jsonOk_({ dataUrl: img.dataUrl });
  } catch (e) {
    return jsonErr_(e);
  }
}

// ────────────────────────────────────────────────────────────────
// 設定（Gemini API キー）— クラスの設定シートに保存。児童 API からは一切返さない
// ────────────────────────────────────────────────────────────────

function opSaveSettings(tenantCode, settings) {
  try {
    const ctx = assertOwner_(tenantCode);
    const ss = openTenantSs_(ctx.code);
    if (settings && settings.apiKey !== undefined && settings.apiKey !== '') {
      setTenantSetting_(ss, 'GEMINI_API_KEY', vStr_(settings.apiKey, 200, 'APIキー').trim());
    }
    return jsonOk_({ message: '設定を保存しました！' });
  } catch (e) {
    return jsonErr_(e);
  }
}

function opTestGeminiKey(apiKey) {
  try {
    ownerEmail_();   // ポータル利用者のみ
    const key = vStr_(apiKey, 200, 'APIキー').trim();
    if (!key) throw new Error('BAD_INPUT: APIキーを入力してください');
    const res = UrlFetchApp.fetch(API_ENDPOINT_V1 + encodeURIComponent(key), {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ contents: [{ parts: [{ text: 'テスト。OKと返して' }] }] }),
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    if (code === 200) return jsonOk_({ message: 'APIキーは有効です！' });
    throw new Error('BAD_INPUT: エラー (HTTP ' + code + ')');
  } catch (e) {
    return jsonErr_(e);
  }
}

// ────────────────────────────────────────────────────────────────
// AI フィードバック支援（Gemini）— 先生本人の操作でのみ実行
// ────────────────────────────────────────────────────────────────

const GEMINI_MODEL = 'gemini-2.5-flash';
const API_ENDPOINT_V1      = 'https://generativelanguage.googleapis.com/v1/models/' + GEMINI_MODEL + ':generateContent?key=';
const API_ENDPOINT_V1_BETA = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent?key=';

const AI_SIMPLE_PROMPT = 'あなたは児童の小さな頑張りやユニークな視点を見つけて具体的に褒めるのが得意な、経験豊富な小学校の先生です。以下の記述を読み、児童が努力した点などを引用しつつ、自己肯定感を育む温かい賞賛のコメントを100字程度で作成してください。見出しや解説は不要です。\n\n';

const AI_FULL_PROMPT = 'あなたは経験豊富な小学校の先生です。以下の児童のジャーナルを読み、フィードバックを作成してください。\n' +
  '# 出力形式の厳密なルール\n' +
  '必ず以下の構造を持つJSONのみを出力してください。\n' +
  '{"comment":"（全体への温かいコメント100字以内）","highlights":[{"textToHighlight":"（本文から完全一致で引用）","suggestedComment":"（ハイライト箇所へのコメント）","suggestedStamp":"（絵文字1つ）"}]}\n' +
  '---\n';

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

function requireGeminiKey_(ss) {
  const apiKey = getTenantSetting_(ss, 'GEMINI_API_KEY');
  if (!apiKey) throw new Error('BAD_INPUT: Gemini APIキーが設定されていません。設定タブから保存してください');
  return apiKey;
}

function opAiSimple(tenantCode) {
  try {
    const ctx = assertOwner_(tenantCode);
    const ss = openTenantSs_(ctx.code);
    const apiKey = requireGeminiKey_(ss);
    const sheet = getJournalSheet_(ss);
    const collected = collectAiTargets_(sheet);
    if (collected.targets.length === 0) return jsonOk_({ message: '対象（未返却×本文あり）のジャーナルがありません。' });

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
    return jsonOk_({ message: 'AIコメント案を作成しました。（成功: ' + ok + '件' + (ng ? '、失敗: ' + ng + '件' : '') + '）' });
  } catch (e) {
    return jsonErr_(e);
  }
}

function opAiFull(tenantCode) {
  try {
    const ctx = assertOwner_(tenantCode);
    const ss = openTenantSs_(ctx.code);
    const apiKey = requireGeminiKey_(ss);
    const sheet = getJournalSheet_(ss);
    const collected = collectAiTargets_(sheet);
    if (collected.targets.length === 0) return jsonOk_({ message: '対象（未返却×本文あり）のジャーナルがありません。' });

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
    return jsonOk_({ message: 'AI高度分析完了。（成功: ' + successCount + '件, 失敗: ' + errorCount + '件）' });
  } catch (e) {
    return jsonErr_(e);
  }
}

// ────────────────────────────────────────────────────────────────
// エクスポート — DriveApp を使わず CSV 文字列を返し、ブラウザ側でダウンロードさせる
// ────────────────────────────────────────────────────────────────

// CSVインジェクション対策: 数式として解釈されうる先頭文字はシングルクォートで無害化
function csvSafe_(v) {
  const s = String(v == null ? '' : v);
  return /^[=+\-@\t]/.test(s) ? "'" + s : s;
}

function opExportCsv(tenantCode, params) {
  try {
    const ctx = assertOwner_(tenantCode);
    const ss = openTenantSs_(ctx.code);
    const p = params || {};
    const roster = getRosterRows_(ss);
    const nameMap = {};
    roster.forEach(function (m) { nameMap[m.email] = m.name; });

    const startDate = p.startDate ? new Date(p.startDate + 'T00:00:00+09:00') : null;
    const endDate = p.endDate ? new Date(p.endDate + 'T23:59:59+09:00') : null;
    const filterEmail = (p.email && p.email !== 'all') ? String(p.email).trim().toLowerCase() : null;

    const journals = getJournalsAll_(ss).filter(function (j) {
      if (filterEmail && String(j.email).toLowerCase() !== filterEmail) return false;
      if (j.timestamp instanceof Date) {
        if (startDate && j.timestamp < startDate) return false;
        if (endDate && j.timestamp > endDate) return false;
      }
      return true;
    }).map(function (j) {
      j.studentName = nameMap[String(j.email).toLowerCase()] || '不明';
      return j;
    }).sort(function (a, b) {
      return (a.studentName || '').localeCompare(b.studentName || '') || (a.timestamp || 0) - (b.timestamp || 0);
    });

    if (journals.length === 0) throw new Error('NOT_FOUND: データがありません');

    const csvHeaders = ['日付', '氏名', 'テーマ', '本文', '気持ち', '先生のコメント', 'ステータス'];
    const csvRows = journals.map(function (j) {
      return [j.date || '', j.studentName || '', j.theme || '', j.content || '', j.emotion || '', j.teacherComment || '', j.status || ''];
    });
    const csvContent = [csvHeaders].concat(csvRows).map(function (row) {
      return row.map(function (cell) { return '"' + csvSafe_(cell).replace(/"/g, '""') + '"'; }).join(',');
    }).join('\r\n');

    return jsonOk_({
      csv: '\uFEFF' + csvContent,
      fileName: 'ジャーナルデータ_' + Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss') + '.csv'
    });
  } catch (e) {
    return jsonErr_(e);
  }
}
