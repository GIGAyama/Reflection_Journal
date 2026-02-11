/**
 * ふり返りジャーナル アプリケーション
 * サーバーサイド スクリプト (Google Apps Script)
 * * 役割：
 * - Webページのリクエストを受け取り、ユーザー認証を行う
 * - GoogleスプレッドシートやGoogle Driveとのデータ送受信（読み取り、書き込み）
 * - Gemini APIと連携し、AIによるコメント生成を行う
 */

//================================================================
// ■ 1. 全体設定（定数）
//================================================================

// 各シートの名前
const ROSTER_SHEET_NAME = "児童名簿";
const JOURNAL_SHEET_NAME = "ジャーナルデータ";
const THEME_SHEET_NAME = "テーマ設定";

// ▼▼▼【改善点】AI機能の分離とAPIエンドポイントの更新 ▼▼▼
// 安定版API（コメントのみ生成用）
const API_ENDPOINT_V1 = "https://generativelanguage.googleapis.com/v1/models/gemini-2.5-flash:generateContent?key=";
// 最新機能が使えるベータ版API（JSON形式でのフィードバック生成用）
const API_ENDPOINT_V1_BETA = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=";
// ▲▲▲ ここまで ▲▲▲


//================================================================
// ■ 2. Webアプリケーションの基本処理
//================================================================

/**
 * Webアプリにアクセスがあったときに最初に実行される関数
 * @param {object} e - Webアクセス時のイベント情報
 * @returns {HtmlOutput} 生成されたWebページ
 */
function doGet(e) {
  // アクセスしてきたユーザーのメールアドレスを取得
  const userEmail = Session.getActiveUser().getEmail();
  // 名簿シートと照合し、ユーザー情報を取得
  const userData = getUserData(userEmail);

  // もし名簿に登録されていないユーザーであれば、アクセス拒否画面を表示
  if (!userData) {
    return HtmlService.createHtmlOutput('<h1>アクセス権がありません</h1><p>名簿に登録されているアカウントでログインしてください。</p>');
  }

  // Webページの元となるHTMLファイル（index.html）を読み込む
  const htmlTemplate = HtmlService.createTemplateFromFile('index');
  
  // HTML側にユーザー情報を渡す
  htmlTemplate.user = userData;
  htmlTemplate.todayTheme = getTodayTheme();
  
  // ユーザーの役割（担任か児童か）に応じて、渡すデータを変える
  if (userData.role === '担任') {
    // 先生の場合：クラス全員の名簿と、全員のジャーナルデータを渡す
    htmlTemplate.classRoster = getClassRoster();
    htmlTemplate.journals = getJournalsForTeacher();
  } else {
    // 児童の場合：自分のジャーナルデータのみを渡す
    htmlTemplate.journals = getJournalsForStudent(userEmail);
  }
  
  // 準備したデータを元に、最終的なWebページを生成して表示
  return htmlTemplate.evaluate()
    .setTitle('ふり返りジャーナル')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * Webアプリ自身のURLを取得するための関数（ページの再読み込みなどで使用）
 * @returns {string} WebアプリのURL
 */
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}


//================================================================
// ■ 3. ユーザー認証と名簿データの処理
//================================================================

/**
 * メールアドレスを元に「児童名簿」シートからユーザー情報を検索する
 * @param {string} email - 検索するユーザーのメールアドレス
 * @returns {object | null} - 見つかったユーザーの情報、見つからなければnull
 */
function getUserData(email) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ROSTER_SHEET_NAME);
  // 2行目から最終行までのデータを一括で取得
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  
  // 1行ずつループして、メールアドレスが一致する行を探す
  for (let i = 0; i < data.length; i++) {
    // C列（配列のインデックスは2）のメールアドレスをチェック
    if (data[i][2] === email) {
      // A列（インデックス0）が「担任」かどうかで役割を決定
      const role = (data[i][0] === '担任') ? '担任' : '児童';
      // 見つかったユーザーの情報をオブジェクトとして返す
      return {
        role: role,
        name: data[i][1], // B列（氏名）
        email: data[i][2] // C列（メールアドレス）
      };
    }
  }
  // 最後まで見つからなければnullを返す
  return null;
}

/**
 * 先生用に、クラスの児童一覧を取得する
 * @returns {Array} - 児童の名前とメールアドレスのリスト
 */
function getClassRoster() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ROSTER_SHEET_NAME);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  
  const roster = data
    // 「担任」の行を除外し、メールアドレスが入力されている行のみを対象にする
    .filter(row => row[0] !== '担任' && row[2])
    // 必要な情報（氏名とメールアドレス）だけを抜き出して新しい配列を作成
    .map(row => ({ name: row[1], email: row[2] }));
    
  return roster;
}


//================================================================
// ■ 4. ジャーナルデータの読み書き (CRUD処理)
//================================================================

/**
 * 児童が提出したジャーナルをスプレッドシートに保存する
 * @param {object} journalData - Webページから送られてきたジャーナルのデータ
 * @returns {object} - 保存結果（成功したか、メッセージ）
 */
function saveJournal(journalData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(JOURNAL_SHEET_NAME);
    // スプレッドシートの列の順番に合わせてデータを配列に格納
    const newRow = [
      Utilities.getUuid(),         // A: journalId (ユニークなIDを自動生成)
      new Date(),                  // B: timestamp (現在日時)
      journalData.email,           // C: email
      journalData.theme,           // D: theme
      journalData.content,         // E: content
      journalData.imageFileId || '', // F: imageFileId (なければ空文字)
      journalData.emotion || '',     // G: emotion (なければ空文字)
      '',                          // H: teacherComment (最初は空)
      '[]',                        // I: highlights (最初は空の配列)
      '',                          // J: teacherStamp (現在は未使用)
      '未返却',                      // K: status
      ''                           // L: pastComment (最初は空)
    ];
    // シートの最終行に新しいデータを追加
    sheet.appendRow(newRow);
    return { success: true, message: 'ジャーナルを提出しました。' };
  } catch (e) {
    Logger.log(e); // エラーが発生した場合はログに記録
    return { success: false, message: 'エラーが発生しました：' + e.message };
  }
}

/**
 * 特定の児童のジャーナルデータをすべて取得する
 * @param {string} email - 児童のメールアドレス
 * @returns {Array} - その児童のジャーナルデータのリスト
 */
function getJournalsForStudent(email) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(JOURNAL_SHEET_NAME);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data.shift(); // 1行目を見出しとして取得
  const emailColIndex = headers.indexOf('email');
  
  return data
    // メールアドレスが一致する行だけを絞り込む
    .filter(row => row[emailColIndex] === email)
    // 扱いやすいように、配列からオブジェクトに変換
    .map(row => {
      let journal = {};
      headers.forEach((header, i) => journal[header] = row[i]);
      if (journal.timestamp instanceof Date) {
         journal.date = Utilities.formatDate(journal.timestamp, "JST", "yyyy/MM/dd");
      }
      
      // 画像表示用のURLを生成する処理
      if (journal.imageFileId) {
        journal.imageUrl = `https://lh3.googleusercontent.com/d/${journal.imageFileId}`;
      } else {
        journal.imageUrl = ''; // 画像がない場合は空文字を設定
      }
      
      return journal;
    });
}

/**
 * 先生用に、クラス全員のジャーナルデータを取得する
 * @returns {Array} - 全員のジャーナルデータのリスト
 */
function getJournalsForTeacher() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(JOURNAL_SHEET_NAME);
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const roster = getClassRoster(); // 児童名簿を取得
    // メールアドレスと名前を紐付けるための対応表を作成
    const nameMap = roster.reduce((map, user) => {
        map[user.email] = user.name;
        return map;
    }, {});
    const emailColIndex = headers.indexOf('email');

    return data.map(row => {
        let journal = {};
        headers.forEach((header, i) => journal[header] = row[i]);
        if (journal.timestamp instanceof Date) {
            journal.date = Utilities.formatDate(journal.timestamp, "JST", "yyyy/MM/dd");
        }
        // 対応表を使って、メールアドレスから児童の名前を追加
        journal.studentName = nameMap[row[emailColIndex]] || '不明';

        // 画像表示用のURLを生成する処理
        if (journal.imageFileId) {
            journal.imageUrl = `https://lh3.googleusercontent.com/d/${journal.imageFileId}`;
        } else {
            journal.imageUrl = ''; // 画像がない場合は空文字を設定
        }

        return journal;
    });
}

/**
 * 先生からのフィードバック（コメントやハイライト）を保存する
 * @param {object} feedbackData - Webページから送られてきたフィードバック情報
 * @returns {object} - 保存結果
 */
function saveFeedback(feedbackData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(JOURNAL_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idColIndex = headers.indexOf('journalId');

    // 1行ずつ確認し、該当するジャーナルIDの行を探す
    for (let i = 1; i < data.length; i++) {
      if (data[i][idColIndex] === feedbackData.journalId) {
        const rowNum = i + 1; // 配列のインデックスとシートの行番号のズレを調整
        // 対応する列にデータを書き込む
        sheet.getRange(rowNum, headers.indexOf('teacherComment') + 1).setValue(feedbackData.comment);
        sheet.getRange(rowNum, headers.indexOf('highlights') + 1).setValue(feedbackData.highlights || '[]');
        sheet.getRange(rowNum, headers.indexOf('status') + 1).setValue('返却済み');
        return { success: true, message: 'フィードバックを保存しました。' };
      }
    }
    return { success: false, message: '該当するジャーナルが見つかりませんでした。' };
  } catch (e) {
    Logger.log(e);
    return { success: false, message: 'エラーが発生しました：' + e.message };
  }
}

/**
 * 返却済みのジャーナルを「未返却」に戻す
 * @param {string} journalId - 対象のジャーナルID
 * @returns {object} - 処理結果
 */
function revertJournalStatus(journalId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(JOURNAL_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idColIndex = headers.indexOf('journalId');
    const statusColIndex = headers.indexOf('status');

    for (let i = 1; i < data.length; i++) {
      if (data[i][idColIndex] === journalId) {
        const rowNum = i + 1;
        sheet.getRange(rowNum, statusColIndex + 1).setValue('未返却');
        return { success: true };
      }
    }
    return { success: false, message: '該当するジャーナルが見つかりませんでした。' };
  } catch (e) {
    Logger.log(e);
    return { success: false, message: 'ステータスの更新中にエラーが発生しました：' + e.message };
  }
}

/**
 * 過去のジャーナルに、現在の自分からのコメントを追記する
 * @param {string} journalId - 対象のジャーナルID
 * @param {string} comment - 追記するコメント
 * @returns {object} - 保存結果
 */
function addPastComment(journalId, comment) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(JOURNAL_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idColIndex = headers.indexOf('journalId');
    const commentColIndex = headers.indexOf('pastComment');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][idColIndex] === journalId) {
        const rowNum = i + 1;
        sheet.getRange(rowNum, commentColIndex + 1).setValue(comment);
        return { success: true, message: '過去の自分へのコメントを保存しました。' };
      }
    }
    return { success: false, message: '該当するジャーナルが見つかりませんでした。' };
  } catch (e) {
     return { success: false, message: 'エラーが発生しました：' + e.message };
  }
}

//================================================================
// ■ 5. Google Drive 連携（画像アップロード）
//================================================================

/**
 * 児童がアップロードした画像をGoogle Driveに保存する
 * @param {object} fileData - base64形式にエンコードされた画像データ
 * @returns {object} - アップロード結果（成功したか、ファイルIDなど）
 */
function uploadImage(fileData) {
  try {
    // スクリプトプロパティからフォルダIDを読み込む
    const folderId = PropertiesService.getScriptProperties().getProperty('IMAGE_FOLDER_ID');
    if (!folderId) {
      return { success: false, message: "管理エラー: 画像保存用のフォルダIDが設定されていません。" };
    }
    const folder = DriveApp.getFolderById(folderId);

    const decoded = Utilities.base64Decode(fileData.data, Utilities.Charset.UTF_8);
    const blob = Utilities.newBlob(decoded, fileData.mimeType, fileData.fileName);
    const file = folder.createFile(blob);
    
    // ★重要★ ファイルの共有設定を変更し、Webアプリから閲覧できるようにする
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return { success: true, fileId: file.getId(), fileName: file.getName() };
  } catch(e) {
    Logger.log(e);
    return { success: false, message: "画像のアップロードに失敗しました：" + e.message };
  }
}


//================================================================
// ■ 6. 「今日のテーマ」の管理
//================================================================

/**
 * 「テーマ設定」シートから、今日の日付に対応するテーマを取得する
 * @returns {string} - 今日のテーマ
 */
function getTodayTheme() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(THEME_SHEET_NAME);
    if (!sheet) return "今日のふり返りをしましょう";
    const todayStr = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd");
    const data = sheet.getDataRange().getValues();
    
    // 下の行から（＝新しい日付から）探す
    for (let i = data.length - 1; i > 0; i--) {
        const rowDate = data[i][0];
        if (rowDate instanceof Date && Utilities.formatDate(rowDate, "JST", "yyyy/MM/dd") === todayStr) {
            return data[i][1];
        }
    }
    // 見つからなければデフォルトのテーマを返す
    return "今日のふり返りをしましょう";
}

/**
 * 先生が設定した「今日のテーマ」をシートに保存する
 * @param {string} theme - 新しいテーマ
 * @returns {object} - 保存結果
 */
function setTodayTheme(theme) {
    try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(THEME_SHEET_NAME);
        const today = new Date();
        sheet.appendRow([today, theme]);
        return { success: true, message: '今日のテーマを設定しました。' };
    } catch (e) {
        return { success: false, message: 'エラーが発生しました：' + e.message };
    }
}


//================================================================
// ■ 7. AI (Gemini API) との連携
//================================================================

// ▼▼▼【改善点】AI機能を2種類に分離 ▼▼▼

/**
 *【機能1：シンプル】未返却のジャーナルに「コメント案のみ」を生成する
 */
function generateAiSimpleCommentsForAll() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(JOURNAL_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    
    const contentColIndex = headers.indexOf('content');
    const statusColIndex = headers.indexOf('status');
    const teacherCommentColIndex = headers.indexOf('teacherComment');
    
    let successCount = 0;

    data.forEach((row, index) => {
      // ステータスが「未返却」で、本文が空でないジャーナルを対象にする
      if (row[statusColIndex] === '未返却' && row[contentColIndex]) {
        const content = row[contentColIndex];
        // シンプルなコメント生成APIを呼び出す
        const generatedComment = callGeminiApiForSimpleComment(content);
        // コメント欄に結果を書き込む
        sheet.getRange(index + 2, teacherCommentColIndex + 1).setValue(generatedComment);
        successCount++;
      }
    });

    return { success: true, message: `AIコメント案（シンプル）の生成が完了しました。（${successCount}件）` };
  } catch(e) {
    Logger.log(e);
    return { success: false, message: 'AI処理中に予期せぬエラーが発生しました：' + e.message };
  }
}

/**
 *【機能2：高度】未返却のジャーナルに「コメント＋ハイライト＋スタンプ案」を生成する
 */
function generateAiFullFeedbackForAll() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(JOURNAL_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    
    const idColIndex = headers.indexOf('journalId');
    const contentColIndex = headers.indexOf('content');
    const statusColIndex = headers.indexOf('status');
    const teacherCommentColIndex = headers.indexOf('teacherComment');
    const highlightsColIndex = headers.indexOf('highlights');
    
    let successCount = 0;
    let errorCount = 0;

    data.forEach((row, index) => {
      if (row[statusColIndex] === '未返却' && row[contentColIndex]) {
        const content = row[contentColIndex];
        const aiResponse = callGeminiApiForFullFeedback(content);

        if (aiResponse && aiResponse.success) {
          try {
            const feedback = JSON.parse(aiResponse.jsonText);

            if (feedback.comment) {
              sheet.getRange(index + 2, teacherCommentColIndex + 1).setValue(feedback.comment);
            }

            if (feedback.highlights && feedback.highlights.length > 0) {
              const highlightsToSave = [];
              feedback.highlights.forEach(h => {
                const startIndex = content.indexOf(h.textToHighlight);
                if (startIndex !== -1) {
                  highlightsToSave.push({
                    id: 'hl-' + Date.now() + Math.random(),
                    text: h.textToHighlight,
                    comment: h.suggestedComment || '', // AIが提案した短いコメントを保存
                    stamp: h.suggestedStamp || '',
                    startOffset: startIndex,
                    endOffset: startIndex + h.textToHighlight.length
                  });
                } else {
                   Logger.log(`AI提案のハイライト「${h.textToHighlight}」は本文中に見つかりませんでした。`);
                }
              });
              if(highlightsToSave.length > 0){
                sheet.getRange(index + 2, highlightsColIndex + 1).setValue(JSON.stringify(highlightsToSave));
              }
            }
            successCount++;
          } catch (e) {
            Logger.log(`AI応答の解析エラー: JournalID: ${row[idColIndex]}, Error: ${e.message}`);
            errorCount++;
          }
        } else {
          Logger.log(`AI呼び出し失敗: JournalID: ${row[idColIndex]}, Message: ${aiResponse.message}`);
          errorCount++;
        }
      }
    });

    return { success: true, message: `AIフィードバック案（高度）の生成完了。（成功: ${successCount}件, 失敗: ${errorCount}件）` };
  } catch(e) {
    Logger.log(e);
    return { success: false, message: 'AI処理中に予期せぬエラーが発生しました：' + e.message };
  }
}


/**
 *【API呼出1：シンプル】Gemini APIを呼び出し、コメント文章のみを生成する
 */
function callGeminiApiForSimpleComment(journalContent) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) return "エラー: APIキーが設定されていません。";

  const prompt = `あなたは児童の小さな頑張りやユニークな視点を見つけて具体的に褒めるのが得意な、経験豊富な小学校の先生です。以下の記述を読み、児童が「努力した点」「工夫した点」などを引用しつつ、自己肯定感を育む温かい賞賛のコメントを100字程度で作成してください。見出しや解説は不要です。`;
  
  const fullPrompt = `${prompt}\n\n---\n\n${journalContent}`;
  const payload = { "contents": [{ "parts": [{ "text": fullPrompt }] }] };
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };
  
  try {
    const response = UrlFetchApp.fetch(API_ENDPOINT_V1 + apiKey, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const jsonResponse = JSON.parse(responseBody);
      return jsonResponse.candidates[0].content.parts[0].text.trim();
    } else {
      Logger.log(`API Error (Simple): ${responseCode} - ${responseBody}`);
      return "コメントの生成に失敗しました。";
    }
  } catch(e) {
      Logger.log(`Fetch Error (Simple): ${e.message}`);
      return "コメントの生成中にエラーが発生しました。";
  }
}

/**
 *【API呼出2：高度】Gemini APIを呼び出し、フィードバック案をJSON形式で受け取る
 */
function callGeminiApiForFullFeedback(journalContent) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    return { success: false, jsonText: null, message: "エラー: APIキーが設定されていません。" };
  }

  const prompt = `あなたは、児童一人ひとりの学びと感情に寄り添う、経験豊富な小学校の先生です。以下の児童のジャーナルを読み、フィードバックを作成してください。# あなたの目的: あなたの目的は、児童の記述の中から「学びの芽生え」「感情の動き」「素晴らしい視点」を見つけ出し、それを価値付けることで、児童の自己肯定感とメタ認知能力を育むことです。# 出力形式の厳密なルール: あなたの回答は、必ず以下の構造を持つJSONオブジェクト"のみ"でなければなりません。解説や前置き、\`\`\`json や \`\`\` といったマークダウンは一切含めないでください。{"comment": "（ここに、児童全体への温かいコメントを100字以内で記述）","highlights": [{"textToHighlight": "（ここに、ジャーナル本文から「完全一致」で引用した、ハイライトすべき最も重要な部分を記述）","suggestedComment": "（ここに、ハイライト箇所への、ごく短い共感や疑問や賞賛のコメントを記述。例：「すごい発見だね！」「その気持ち、わかるよ」「すごい！」「えらい！」「やったね！」「どうして？」「何をしたの？」「もっとくわしく教えてね」など）","suggestedStamp": "（ここに、そのハイライトに最もふさわしい感情や評価を表す絵文字スタンプを1つだけ記述）"}]}# ハイライト選定の最重要ルール: - ジャーナルの中から、あなたの目的（学びや感情の価値付け）に最も合致する箇所を「1つだけ」選び、正確に引用してください。- もし同じ言葉が複数ある場合は、文脈上、最も児童の気づきや感情の中心となっている箇所を選んでください。---# 児童のジャーナル\n${journalContent}`;

  const payload = {
    "contents": [{ "parts": [{ "text": prompt }] }],
    "generationConfig": { "responseMimeType": "application/json" }
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };
  
  try {
    // 最新機能用のv1betaエンドポイントを使用
    const response = UrlFetchApp.fetch(API_ENDPOINT_V1_BETA + apiKey, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const jsonResponse = JSON.parse(responseBody);
      const text = jsonResponse.candidates[0].content.parts[0].text;
      return { success: true, jsonText: text, message: "成功" };
    } else {
      const errorMsg = `API Error (Full): ${responseCode} - ${responseBody}`;
      Logger.log(errorMsg);
      return { success: false, jsonText: null, message: `APIからエラーが返されました。` };
    }
  } catch(e) {
      const errorMsg = `Fetch Error (Full): ${e.message}`;
      Logger.log(errorMsg);
      return { success: false, jsonText: null, message: `APIへの接続中にエラーが発生しました。` };
  }
}

