/**
 * OwnerApi.gs — 先生の画面から呼ぶ API
 *
 * 1 ファイル = 1 クラスなので、どのクラスを触るかを引数で受け取らない。
 * 対象は必ず「このスクリプトが束ねられているスプレッドシート」である。
 *
 * すべての公開関数は先頭で assertOwner_() を通す。`google.script.run` は末尾 `_` の
 * 無い関数を誰でも呼べるので、1 つ抜けるとそこから学級全員の記録が読み書きできる。
 * 返り値は { success:true, ... } / { success:false, code, error } の JSON 文字列。
 */

// ────────────────────────────────────────────────────────────────
// クラスの状態
// ────────────────────────────────────────────────────────────────

/** 先生画面の初期同期。先生には氏名とメールアドレスを含めて返してよい */
function opGetClassData() {
  try {
    const ctx = assertOwner_();
    const ss = ctx.ss;
    const roster = getRosterRows_(ss);
    const nameMap = {};
    roster.forEach(function (m) { nameMap[m.email] = m.name; });
    const journals = getJournalsAll_(ss).map(function (j) {
      j.studentName = nameMap[String(j.email).toLowerCase()] || '不明';
      return j;
    });
    const apiKey = getTenantSetting_(ss, 'GEMINI_API_KEY');
    return jsonOk_({
      klass: {
        className: classNameOf_(ss),
        // 児童に配る URL。このデプロイの /exec そのもの
        memberUrl: ScriptApp.getService().getUrl() || '',
        joinOpen: getTenantSetting_(ss, 'JOIN_CLOSED') !== '1',
        autoApprove: getTenantSetting_(ss, 'AUTO_APPROVE') === '1',
        ownerEmail: ownerEmailOf_(ss)
      },
      journals: journals,
      classRoster: roster.filter(function (m) { return m.role !== '担任' && m.status === 'active'; }),
      pendingMembers: roster.filter(function (m) { return m.status === 'pending'; }),
      rosterAll: roster,
      todayTheme: getTodayTheme_(ss),
      weeklyThemes: getWeeklyThemes_(ss),
      // シートの作りがずれていたら、先生の画面にも出す。
      // スプレッドシートのメニューを開かない先生にも気づいてもらうため。
      sheetIssues: inspectSheets_(ss),
      settings: {
        hasApiKey: !!apiKey,
        apiKeyMasked: apiKey.length > 10
          ? apiKey.substring(0, 6) + '****' + apiKey.substring(apiKey.length - 4)
          : (apiKey ? '設定済み' : '')
      },
      spreadsheetUrl: ss.getUrl()
    });
  } catch (e) {
    return jsonErr_(e);
  }
}

/**
 * 「準備の状態」。何がまだ済んでいないかを 1 か所で返す。
 *
 * 設定作業そのものは要らなくなったが、**要らなくなったことが分かる**必要がある。
 * 「押すものが無い」と「押し忘れている」は、先生から見ると区別がつかないためである。
 */
function opGetSetupStatus() {
  try {
    const ctx = assertOwner_();
    return jsonOk_({ status: buildSetupStatus_(ctx.ss) });
  } catch (e) {
    return jsonErr_(e);
  }
}

/**
 * コピー元から引き継がれた名簿・提出・画像を消す。
 *
 * **自動では消さない。** コピーには「前の学級のものを片づけたい」場合と
 * 「去年の記録を持ち越したい」場合があり、取り違えると取り返しがつかない。
 * だから先生が「準備の状態」で押したときだけ動く。
 * 設定とテーマには触らない（テーマの案は使い回せる）。
 */
function opClearInheritedData() {
  try {
    const ctx = assertOwner_();
    return withScriptLock_(function () {
      const cleared = clearInheritedRows_(ctx.ss);
      writeMeta_(ctx.ss, { copiedAt: '' });
      return jsonOk_({
        cleared: cleared,
        message: 'コピー元から引き継がれた記録を消しました。'
      });
    });
  } catch (e) {
    return jsonErr_(e);
  }
}

/** クラス名の変更（_meta に控える） */
function opSetClassName(className) {
  try {
    const ctx = assertOwner_();
    const name = vStr_(className, 50, 'クラス名').trim();
    if (!name) throw new Error('BAD_INPUT: クラス名を入力してください');
    writeMeta_(ctx.ss, { className: name });
    return jsonOk_({ className: name });
  } catch (e) {
    return jsonErr_(e);
  }
}

/** 参加受付の開閉・自動承認の切り替え */
function opUpdateJoinPolicy(joinOpen, autoApprove) {
  try {
    const ctx = assertOwner_();
    setTenantSetting_(ctx.ss, 'JOIN_CLOSED', joinOpen ? '' : '1');
    setTenantSetting_(ctx.ss, 'AUTO_APPROVE', autoApprove ? '1' : '');
    return jsonOk_({ joinOpen: !!joinOpen, autoApprove: !!autoApprove });
  } catch (e) {
    return jsonErr_(e);
  }
}

// ────────────────────────────────────────────────────────────────
// シートの点検・修整（先生の画面からも押せるようにする）
// ────────────────────────────────────────────────────────────────

/** 点検だけ。1 セルも書き換えない */
function opInspectSheets() {
  try {
    const ctx = assertOwner_();
    return jsonOk_({ issues: inspectSheets_(ctx.ss) });
  } catch (e) {
    return jsonErr_(e);
  }
}

/**
 * 直せる範囲だけ直す。
 * 直したことと、人にしか直せなかったことの両方を返す（黙って握りつぶさない）。
 */
function opRepairSheets() {
  try {
    const ctx = assertOwner_();
    const result = repairSheets_(ctx.ss);
    return jsonOk_({
      fixed: result.fixed,
      left: result.left,
      message: result.fixed.length
        ? result.fixed.length + ' か所を直しました。'
        : '自動で直せるところはありませんでした。'
    });
  } catch (e) {
    return jsonErr_(e);
  }
}

// ────────────────────────────────────────────────────────────────
// 名簿
// ────────────────────────────────────────────────────────────────

/** 参加申請の承認 */
function opApproveMember(memberEmail) {
  try {
    const ctx = assertOwner_();
    const target = String(memberEmail || '').trim().toLowerCase();
    const row = getMemberRow_(ctx.ss, target);
    if (!row) throw new Error('NOT_MEMBER: その申請は見つかりません');
    withScriptLock_(function () {
      upsertMember_(ctx.ss, { email: target, name: row.name, role: row.role, status: 'active' });
    });
    return jsonOk_({});
  } catch (e) {
    return jsonErr_(e);
  }
}

/** 参加申請の却下 / 名簿からの削除 */
function opRejectMember(memberEmail) {
  try {
    const ctx = assertOwner_();
    const target = String(memberEmail || '').trim().toLowerCase();
    if (target === ctx.email) throw new Error('BAD_INPUT: 自分自身は名簿から外せません');
    withScriptLock_(function () { removeMember_(ctx.ss, target); });
    return jsonOk_({});
  } catch (e) {
    return jsonErr_(e);
  }
}

/** 名簿の一括保存。決めたキーだけ書き込む */
function opSaveRoster(rows) {
  try {
    const ctx = assertOwner_();
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
    withScriptLock_(function () { saveRosterRows_(ctx.ss, clean); });
    return jsonOk_({ message: '名簿を保存しました' });
  } catch (e) {
    return jsonErr_(e);
  }
}

// ────────────────────────────────────────────────────────────────
// テーマ
// ────────────────────────────────────────────────────────────────

function opSetTodayTheme(theme) {
  try {
    const ctx = assertOwner_();
    const t = vStr_(theme, 200, 'テーマ').trim();
    if (!t) throw new Error('BAD_INPUT: テーマを入力してください');
    const sheet = ctx.ss.getSheetByName(THEME_SHEET_NAME);
    if (!sheet) throw new Error('SHEET_BROKEN: 「' + THEME_SHEET_NAME + '」シートがありません');
    const map = headerMap_(sheet);
    const cDate = colOf_(map, '日付', THEME_SHEET_NAME);
    const cTheme = colOf_(map, 'テーマ', THEME_SHEET_NAME);
    withScriptLock_(function () {
      const width = Math.max(sheet.getLastColumn(), 1);
      const row = new Array(width).fill('');
      row[cDate - 1] = new Date();
      row[cTheme - 1] = t;
      sheet.appendRow(row);
    });
    return jsonOk_({ message: 'テーマを設定しました！' });
  } catch (e) {
    return jsonErr_(e);
  }
}

function opGetWeeklyThemes() {
  try {
    const ctx = assertOwner_();
    return jsonOk_({ data: getWeeklyThemes_(ctx.ss) });
  } catch (e) {
    return jsonErr_(e);
  }
}

function opSaveWeeklyThemes(themes) {
  try {
    const ctx = assertOwner_();
    const clean = {};
    ['mon', 'tue', 'wed', 'thu', 'fri'].forEach(function (d) {
      clean[d] = vStr_(themes && themes[d], 200, 'テーマ');
    });
    setTenantSetting_(ctx.ss, 'WEEKLY_THEMES', JSON.stringify(clean));
    return jsonOk_({ message: '保存しました！' });
  } catch (e) {
    return jsonErr_(e);
  }
}

// ────────────────────────────────────────────────────────────────
// 返却
// ────────────────────────────────────────────────────────────────

/** highlights は JSON 配列の文字列。決めたキーだけ通す */
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

function opSaveFeedback(feedbackData) {
  try {
    const ctx = assertOwner_();
    const sheet = getJournalSheet_(ctx.ss);
    const d = feedbackData || {};
    return withScriptLock_(function () {
      const cols = journalCols_(sheet);
      const rowNum = findJournalRowById_(sheet, vJournalId_(d.journalId));
      if (rowNum < 0) throw new Error('NOT_FOUND: 該当するジャーナルが見つかりません');
      sheet.getRange(rowNum, cols.teacherComment).setValue(vStr_(d.comment, 2000, 'コメント'));
      sheet.getRange(rowNum, cols.highlights).setValue(cleanHighlights_(d.highlights));
      sheet.getRange(rowNum, cols.status).setValue('返却済み');
      return jsonOk_({ message: 'フィードバックを保存しました！' });
    });
  } catch (e) {
    return jsonErr_(e);
  }
}

function opQuickReturn(journalId, stamp) {
  try {
    const ctx = assertOwner_();
    const sheet = getJournalSheet_(ctx.ss);
    return withScriptLock_(function () {
      const cols = journalCols_(sheet);
      const rowNum = findJournalRowById_(sheet, vJournalId_(journalId));
      if (rowNum < 0) throw new Error('NOT_FOUND: 見つかりませんでした');
      sheet.getRange(rowNum, cols.teacherStamp).setValue(vStr_(stamp, 10, 'スタンプ'));
      sheet.getRange(rowNum, cols.status).setValue('返却済み');
      return jsonOk_({ message: 'スタンプで返却しました！' });
    });
  } catch (e) {
    return jsonErr_(e);
  }
}

function opRevertStatus(journalId) {
  try {
    const ctx = assertOwner_();
    const sheet = getJournalSheet_(ctx.ss);
    return withScriptLock_(function () {
      const rowNum = findJournalRowById_(sheet, vJournalId_(journalId));
      if (rowNum < 0) throw new Error('NOT_FOUND: 見つかりませんでした');
      sheet.getRange(rowNum, journalCols_(sheet).status).setValue('未返却');
      return jsonOk_({});
    });
  } catch (e) {
    return jsonErr_(e);
  }
}

function opDeleteJournal(journalId) {
  try {
    const ctx = assertOwner_();
    const sheet = getJournalSheet_(ctx.ss);
    return withScriptLock_(function () {
      const rowNum = findJournalRowById_(sheet, vJournalId_(journalId));
      if (rowNum < 0) throw new Error('NOT_FOUND: 見つかりませんでした');
      sheet.getRange(rowNum, journalCols_(sheet).deletedAt).setValue(new Date());
      return jsonOk_({ message: '削除しました。' });
    });
  } catch (e) {
    return jsonErr_(e);
  }
}

function opBatchReturnAll() {
  try {
    const ctx = assertOwner_();
    const sheet = getJournalSheet_(ctx.ss);
    return withScriptLock_(function () {
      const cols = journalCols_(sheet);
      const data = sheet.getDataRange().getValues();
      let count = 0;
      for (let i = 1; i < data.length; i++) {
        if (data[i][cols.status - 1] === '未返却'
            && data[i][cols.teacherComment - 1]
            && !data[i][cols.deletedAt - 1]) {
          sheet.getRange(i + 1, cols.status).setValue('返却済み');
          count++;
        }
      }
      return jsonOk_({ count: count, message: count > 0 ? count + '件を一括返却しました！' : '返却対象がありません。' });
    });
  } catch (e) {
    return jsonErr_(e);
  }
}

function opResetData() {
  try {
    const ctx = assertOwner_();
    return withScriptLock_(function () {
      const sheet = getJournalSheet_(ctx.ss);
      if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
      const imgSheet = ctx.ss.getSheetByName(IMAGE_SHEET_NAME);
      if (imgSheet && imgSheet.getLastRow() > 1) imgSheet.deleteRows(2, imgSheet.getLastRow() - 1);
      return jsonOk_({ message: '全データを削除しました。' });
    });
  } catch (e) {
    return jsonErr_(e);
  }
}

/** 先生用の画像取得（担任は学級全員の画像を見られる） */
function opGetImage(imageId) {
  try {
    const ctx = assertOwner_();
    const img = loadImage_(ctx.ss, imageId);
    if (!img) throw new Error('NOT_FOUND: 画像が見つかりません');
    return jsonOk_({ dataUrl: img.dataUrl });
  } catch (e) {
    return jsonErr_(e);
  }
}

// ────────────────────────────────────────────────────────────────
// 設定（Gemini API キー）— 設定シートに保存。児童 API からは一切返さない
// ────────────────────────────────────────────────────────────────

function opSaveSettings(settings) {
  try {
    const ctx = assertOwner_();
    if (settings && settings.apiKey !== undefined && settings.apiKey !== '') {
      setTenantSetting_(ctx.ss, 'GEMINI_API_KEY', vStr_(settings.apiKey, 200, 'APIキー').trim());
    }
    return jsonOk_({ message: '設定を保存しました！' });
  } catch (e) {
    return jsonErr_(e);
  }
}

function opTestGeminiKey(apiKey) {
  try {
    assertOwner_();   // 先生だけ
    const key = vStr_(apiKey, 200, 'APIキー').trim();
    if (!key) throw new Error('BAD_INPUT: APIキーを入力してください');
    // 正本 Gemini.gs に投げる。混み合っているだけ（429）のときは正本が
    // 再試行し、それでも駄目なら「AIが混み合っています」と返る。
    // 以前は HTTP 番号だけを見せていたので、混雑と「キーが無効」を
    // 先生が区別できなかった。
    GigaGemini.callRaw({
      apiKey: key, prompt: 'テスト。OKと返して',
      model: GEMINI_MODEL, apiVersion: 'v1'
    });
    return jsonOk_({ message: 'APIキーは有効です！' });
  } catch (e) {
    return jsonErr_(e);
  }
}

// ────────────────────────────────────────────────────────────────
// AI フィードバック支援（Gemini）— 先生本人の操作でのみ実行
// ────────────────────────────────────────────────────────────────

const GEMINI_MODEL = 'gemini-2.5-flash';
// 呼び出し先の URL は正本 Gemini.gs が model と apiVersion から組み立てる。
// API キーは正本側で x-goog-api-key ヘッダに載る（URL クエリには入れない。
// アクセスログやプロキシに残るため）。

const AI_SIMPLE_PROMPT = 'あなたは児童の小さな頑張りやユニークな視点を見つけて具体的に褒めるのが得意な、経験豊富な小学校の先生です。以下の記述を読み、児童が努力した点などを引用しつつ、自己肯定感を育む温かい賞賛のコメントを100字程度で作成してください。見出しや解説は不要です。\n\n';

const AI_FULL_PROMPT = 'あなたは経験豊富な小学校の先生です。以下の児童のジャーナルを読み、フィードバックを作成してください。\n' +
  '# 出力形式の厳密なルール\n' +
  '必ず以下の構造を持つJSONのみを出力してください。\n' +
  '{"comment":"（全体への温かいコメント100字以内）","highlights":[{"textToHighlight":"（本文から完全一致で引用）","suggestedComment":"（ハイライト箇所へのコメント）","suggestedStamp":"（絵文字1つ）"}]}\n' +
  '---\n';

/** AI にかける対象（未返却 × 本文あり × 未削除）を集める */
function collectAiTargets_(sheet) {
  const cols = journalCols_(sheet);
  const data = sheet.getDataRange().getValues();
  const targets = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][cols.status - 1] === '未返却'
        && data[i][cols.content - 1]
        && !data[i][cols.deletedAt - 1]) {
      targets.push({ rowNum: i + 1, content: String(data[i][cols.content - 1]) });
    }
  }
  return { targets: targets, cols: cols };
}

function requireGeminiKey_(ss) {
  const apiKey = getTenantSetting_(ss, 'GEMINI_API_KEY');
  if (!apiKey) throw new Error('BAD_INPUT: Gemini APIキーが設定されていません。設定タブから保存してください');
  return apiKey;
}

function opAiSimple() {
  try {
    const ctx = assertOwner_();
    const apiKey = requireGeminiKey_(ctx.ss);
    const sheet = getJournalSheet_(ctx.ss);
    const collected = collectAiTargets_(sheet);
    if (collected.targets.length === 0) return jsonOk_({ message: '対象（未返却×本文あり）のジャーナルがありません。' });

    // 正本 Gemini.gs の callAll に任せる。8件ずつの小分けは同じだが、
    // 一時エラー（429/5xx）だったものだけを小分け単位で再試行するので、
    // 混み合う時間帯に「40人中10人だけ失敗」が起きにくい。
    const results = GigaGemini.callAll(collected.targets.map(function (t) {
      return {
        apiKey: apiKey, prompt: AI_SIMPLE_PROMPT + t.content,
        model: GEMINI_MODEL, apiVersion: 'v1'
      };
    }));

    let ok = 0, ng = 0;
    results.forEach(function (res, i) {
      if (!res.ok) { ng++; return; }
      sheet.getRange(collected.targets[i].rowNum, collected.cols.teacherComment).setValue(res.text);
      ok++;
    });
    return jsonOk_({ message: 'AIコメント案を作成しました。（成功: ' + ok + '件' + (ng ? '、失敗: ' + ng + '件' : '') + '）' });
  } catch (e) {
    return jsonErr_(e);
  }
}

function opAiFull() {
  try {
    const ctx = assertOwner_();
    const apiKey = requireGeminiKey_(ctx.ss);
    const sheet = getJournalSheet_(ctx.ss);
    const collected = collectAiTargets_(sheet);
    if (collected.targets.length === 0) return jsonOk_({ message: '対象（未返却×本文あり）のジャーナルがありません。' });

    const results = GigaGemini.callAll(collected.targets.map(function (t) {
      return {
        apiKey: apiKey, prompt: AI_FULL_PROMPT + t.content,
        model: GEMINI_MODEL, apiVersion: 'v1beta',
        generationConfig: { responseMimeType: 'application/json' }
      };
    }));

    let successCount = 0, errorCount = 0;
    results.forEach(function (res, i) {
      const target = collected.targets[i];
      try {
        if (!res.ok) { errorCount++; return; }
        // JSON を頼んでも前置きや ```json が混ざることがある。正本の
        // parseJsonText がそれを剥がしてから読む。
        const feedback = GigaGemini.parseJsonText(res.text);
        if (feedback.comment) {
          sheet.getRange(target.rowNum, collected.cols.teacherComment).setValue(feedback.comment);
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
          if (hlToSave.length > 0) {
            sheet.getRange(target.rowNum, collected.cols.highlights).setValue(JSON.stringify(hlToSave));
          }
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
// 書き出し — CSV 文字列を返し、ブラウザ側でダウンロードさせる
// ────────────────────────────────────────────────────────────────

/**
 * 表計算ソフトが数式として読む先頭文字を無害化する。
 * 児童が本文の先頭に =IMPORTXML("http://…") と書くと、先生が開いた瞬間に
 * 学級のデータが外部へ流れる。CSV もシートも、人の入力は必ずここを通す。
 */
function csvSafe_(v) {
  const s = String(v == null ? '' : v);
  return /^[=+\-@\t]/.test(s) ? "'" + s : s;
}

function opExportCsv(params) {
  try {
    const ctx = assertOwner_();
    const ss = ctx.ss;
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
