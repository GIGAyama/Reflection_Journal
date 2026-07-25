/**
 * Db.gs — テナント DB（クラスごとのスプレッドシート）のスキーマ定義と I/O
 *
 * シート構成（1 クラス = 1 スプレッドシート・先生所有）:
 *   児童名簿       … 役割 / 氏名 / メールアドレス / 状態(active|pending)
 *   ジャーナルデータ … journalId, timestamp, email, theme, content, imageFileId,
 *                     emotion, teacherComment, highlights, teacherStamp,
 *                     status, pastComment, deletedAt
 *   テーマ設定     … 日付 / テーマ
 *   設定           … key / value（GEMINI_API_KEY・WEEKLY_THEMES など先生の設定）
 *   画像データ     … imageId / chunkIndex / data / email / createdAt
 *   _meta          … key / value（schemaVersion・tenantCode 等の識別情報）
 *
 * ⚠️ DriveApp は使わない（フル Drive スコープ回避）。画像は Data URL を
 *    40,000 文字ずつチャンクして画像データシートに保存する（1 セル 50,000 文字上限）。
 */

const ROSTER_SHEET_NAME  = '児童名簿';
const JOURNAL_SHEET_NAME = 'ジャーナルデータ';
const THEME_SHEET_NAME   = 'テーマ設定';
const SETTINGS_SHEET_NAME = '設定';
const IMAGE_SHEET_NAME   = '画像データ';
const META_SHEET_NAME    = '_meta';

const JOURNAL_HEADERS = [
  'journalId', 'timestamp', 'email', 'theme', 'content',
  'imageFileId', 'emotion', 'teacherComment', 'highlights',
  'teacherStamp', 'status', 'pastComment', 'deletedAt'
];
const ROSTER_HEADERS = ['役割', '氏名', 'メールアドレス', '状態'];
const THEME_HEADERS = ['日付', 'テーマ'];
const SETTINGS_HEADERS = ['key', 'value'];
const IMAGE_HEADERS = ['imageId', 'chunkIndex', 'data', 'email', 'createdAt'];
const META_HEADERS = ['key', 'value'];

const IMAGE_CHUNK_SIZE = 40000;                 // 1 セル 50,000 文字上限に対する安全マージン
const IMAGE_MAX_DATAURL_CHARS = 800000;         // Data URL 上限（約 600KB 相当の画像）

// ────────────────────────────────────────────────────────────────
// シート構築
// ────────────────────────────────────────────────────────────────

function createSheetIfNotExists_(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) return sheet;
  sheet = ss.insertSheet(sheetName);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
  return sheet;
}

/** 新規スプレッドシートに必須シートとヘッダーを構築する。既存シートには手を付けない */
function initializeNewDatabase_(ss) {
  createSheetIfNotExists_(ss, ROSTER_SHEET_NAME, ROSTER_HEADERS);
  createSheetIfNotExists_(ss, JOURNAL_SHEET_NAME, JOURNAL_HEADERS);
  createSheetIfNotExists_(ss, THEME_SHEET_NAME, THEME_HEADERS);
  createSheetIfNotExists_(ss, SETTINGS_SHEET_NAME, SETTINGS_HEADERS);
  createSheetIfNotExists_(ss, IMAGE_SHEET_NAME, IMAGE_HEADERS);
  createSheetIfNotExists_(ss, META_SHEET_NAME, META_HEADERS);
  const defaultSheet = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) ss.deleteSheet(defaultSheet);
  return ss;
}

/**
 * 既存シート（テンプレート複製・旧バージョンからの取り込み）でも
 * 必須シート・必須列が揃っていることを保証する。
 * 旧スキーマ（状態列なし）の名簿は、既存行を active として扱えるよう列を追加する。
 */
function ensureTenantSheets_(ss) {
  initializeNewDatabase_(ss);
  const roster = ss.getSheetByName(ROSTER_SHEET_NAME);
  const headerRange = roster.getRange(1, 1, 1, Math.max(roster.getLastColumn(), 1));
  const headers = headerRange.getValues()[0];
  if (headers.indexOf('状態') === -1) {
    const col = headers.filter(String).length + 1;
    roster.getRange(1, col).setValue('状態');
    if (roster.getLastRow() > 1) {
      const rows = roster.getLastRow() - 1;
      const values = [];
      for (let i = 0; i < rows; i++) values.push(['active']);
      roster.getRange(2, col, rows, 1).setValues(values);
    }
  }
  return ss;
}

/** _meta シートへ識別情報を書き込む（デバッグ・復旧用。認可には使わない） */
function writeMeta_(ss, meta) {
  const sheet = createSheetIfNotExists_(ss, META_SHEET_NAME, META_HEADERS);
  Object.keys(meta).forEach(function (k) {
    upsertKeyValue_(sheet, k, String(meta[k]));
  });
}

/** key/value シート共通の upsert */
function upsertKeyValue_(sheet, key, value) {
  if (sheet.getLastRow() >= 2) {
    const cell = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1)
      .createTextFinder(String(key)).matchEntireCell(true).findNext();
    if (cell) { sheet.getRange(cell.getRow(), 2).setValue(value); return; }
  }
  sheet.appendRow([key, value]);
}

function readKeyValue_(sheet, key) {
  if (!sheet || sheet.getLastRow() < 2) return '';
  const cell = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1)
    .createTextFinder(String(key)).matchEntireCell(true).findNext();
  return cell ? String(sheet.getRange(cell.getRow(), 2).getValue()) : '';
}

// ────────────────────────────────────────────────────────────────
// クラス内設定（設定シート）— 先生の Gemini API キー・週間テーマ等
// ────────────────────────────────────────────────────────────────

function getTenantSetting_(ss, key) {
  return readKeyValue_(ss.getSheetByName(SETTINGS_SHEET_NAME), key);
}

function setTenantSetting_(ss, key, value) {
  const sheet = createSheetIfNotExists_(ss, SETTINGS_SHEET_NAME, SETTINGS_HEADERS);
  upsertKeyValue_(sheet, key, value);
}

// ────────────────────────────────────────────────────────────────
// 名簿 I/O
// ────────────────────────────────────────────────────────────────

/** 名簿全行を { role, name, email, status } で返す（email 空行は除外） */
function getRosterRows_(ss) {
  const sheet = ss.getSheetByName(ROSTER_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  return data.filter(function (r) { return r[2]; }).map(function (r) {
    return {
      role: r[0] === '担任' ? '担任' : '児童',
      name: String(r[1] || ''),
      email: String(r[2]).toLowerCase(),
      status: r[3] === 'pending' ? 'pending' : 'active'
    };
  });
}

/** 名簿を丸ごと書き換える（担任 API 用）。呼び出し側でロックすること */
function saveRosterRows_(ss, rows) {
  const sheet = createSheetIfNotExists_(ss, ROSTER_SHEET_NAME, ROSTER_HEADERS);
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).clearContent();
  if (rows.length > 0) {
    const values = rows.map(function (r) {
      return [r.role === '担任' ? '担任' : '児童', r.name, String(r.email).toLowerCase(),
              r.status === 'pending' ? 'pending' : 'active'];
    });
    sheet.getRange(2, 1, values.length, 4).setValues(values);
  }
}

/** 1 名分を追加/更新する。email をキーに既存行を上書きする */
function upsertMember_(ss, member) {
  const sheet = createSheetIfNotExists_(ss, ROSTER_SHEET_NAME, ROSTER_HEADERS);
  const email = String(member.email).toLowerCase();
  const row = [member.role === '担任' ? '担任' : '児童', member.name || '', email,
               member.status === 'pending' ? 'pending' : 'active'];
  if (sheet.getLastRow() >= 2) {
    const cell = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1)
      .createTextFinder(email).matchEntireCell(true).findNext();
    if (cell) { sheet.getRange(cell.getRow(), 1, 1, 4).setValues([row]); return; }
  }
  sheet.appendRow(row);
}

/** 名簿から 1 名を削除する（参加申請の却下・除名） */
function removeMember_(ss, email) {
  const sheet = ss.getSheetByName(ROSTER_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return false;
  const cell = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1)
    .createTextFinder(String(email).toLowerCase()).matchEntireCell(true).findNext();
  if (!cell) return false;
  sheet.deleteRow(cell.getRow());
  return true;
}

// ────────────────────────────────────────────────────────────────
// ジャーナル I/O
// ────────────────────────────────────────────────────────────────

function getJournalSheet_(ss) {
  const sheet = ss.getSheetByName(JOURNAL_SHEET_NAME);
  if (!sheet) throw new Error('SERVER_ERROR: ジャーナルデータのシートが見つかりません');
  return sheet;
}

function getJournalHeaders_(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

/** journalId から行番号を高速検索（TextFinder は全行走査より大幅に速い） */
function findJournalRowById_(sheet, journalId) {
  if (!journalId || sheet.getLastRow() < 2) return -1;
  const cell = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1)
    .createTextFinder(String(journalId)).matchEntireCell(true).findNext();
  return cell ? cell.getRow() : -1;
}

function rowToJournalObject_(headers, row) {
  const journal = {};
  headers.forEach(function (h, i) { journal[h] = row[i]; });
  if (journal.timestamp instanceof Date) {
    journal.date = Utilities.formatDate(journal.timestamp, 'JST', 'yyyy/MM/dd HH:mm');
  }
  // 画像はシート内チャンク保存。ID だけ返し、本体は getImage API で遅延取得する
  journal.imageId = journal.imageFileId || '';
  delete journal.imageFileId;
  return journal;
}

/** 未削除の全ジャーナルをオブジェクト配列で返す */
function getJournalsAll_(ss) {
  const sheet = ss.getSheetByName(JOURNAL_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const deletedAtIdx = headers.indexOf('deletedAt');
  return data.filter(function (row) { return !row[deletedAtIdx]; })
    .map(function (row) { return rowToJournalObject_(headers, row); });
}

function getJournalsForEmail_(ss, email) {
  const target = String(email).toLowerCase();
  return getJournalsAll_(ss).filter(function (j) {
    return String(j.email).toLowerCase() === target;
  });
}

// ────────────────────────────────────────────────────────────────
// テーマ I/O
// ────────────────────────────────────────────────────────────────

function getTodayTheme_(ss) {
  const sheet = ss.getSheetByName(THEME_SHEET_NAME);
  const todayStr = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd');
  if (sheet && sheet.getLastRow() >= 2) {
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i > 0; i--) {
      if (data[i][0] instanceof Date &&
          Utilities.formatDate(data[i][0], 'JST', 'yyyy/MM/dd') === todayStr) {
        return String(data[i][1]);
      }
    }
  }
  return getWeeklyThemeForToday_(ss) || '今日の学びをふり返ろう';
}

function getWeeklyThemes_(ss) {
  try {
    const json = getTenantSetting_(ss, 'WEEKLY_THEMES');
    return json ? JSON.parse(json) : { mon: '', tue: '', wed: '', thu: '', fri: '' };
  } catch (e) {
    return { mon: '', tue: '', wed: '', thu: '', fri: '' };
  }
}

function getWeeklyThemeForToday_(ss) {
  try {
    const themes = getWeeklyThemes_(ss);
    const dayKeys = ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat'];
    return themes[dayKeys[new Date().getDay()]] || null;
  } catch (e) {
    return null; // 設定が壊れていてもアプリ全体は止めない
  }
}

// ────────────────────────────────────────────────────────────────
// 画像 I/O（Data URL をチャンクしてシートに保存）
// ────────────────────────────────────────────────────────────────

/**
 * Data URL を検証してチャンク保存し imageId を返す。呼び出し側でロックすること。
 * email はアップロード者（検証済み）。取得時の所有者チェックに使う。
 */
function saveImageChunks_(ss, dataUrl, email) {
  const s = String(dataUrl || '');
  if (!/^data:image\/(png|jpeg|jpg|webp|gif);base64,[A-Za-z0-9+\/=]+$/.test(s)) {
    throw new Error('BAD_INPUT: 画像データの形式が正しくありません');
  }
  if (s.length > IMAGE_MAX_DATAURL_CHARS) {
    throw new Error('BAD_INPUT: 画像が大きすぎます。小さい画像でためしてね');
  }
  const sheet = createSheetIfNotExists_(ss, IMAGE_SHEET_NAME, IMAGE_HEADERS);
  const imageId = 'img_' + Utilities.getUuid();
  const now = new Date();
  const rows = [];
  for (let i = 0, idx = 0; i < s.length; i += IMAGE_CHUNK_SIZE, idx++) {
    rows.push([imageId, idx, s.substring(i, i + IMAGE_CHUNK_SIZE), String(email).toLowerCase(), now]);
  }
  const start = sheet.getLastRow() + 1;
  const shortage = start + rows.length - 1 - sheet.getMaxRows();
  if (shortage > 0) sheet.insertRowsAfter(sheet.getMaxRows(), shortage);
  sheet.getRange(start, 1, rows.length, 5).setValues(rows);
  return imageId;
}

/** imageId から Data URL を復元して { dataUrl, email } を返す。無ければ null */
function loadImage_(ss, imageId) {
  const sheet = ss.getSheetByName(IMAGE_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return null;
  if (!/^img_[A-Za-z0-9\-]{8,50}$/.test(String(imageId || ''))) return null;
  // 画像シートは大きくなりやすいため全読み込みせず、TextFinder で該当行のみ読む
  const cells = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1)
    .createTextFinder(String(imageId)).matchEntireCell(true).findAll();
  if (cells.length === 0) return null;
  const chunks = [];
  let email = '';
  cells.forEach(function (cell) {
    const row = sheet.getRange(cell.getRow(), 2, 1, 3).getValues()[0];
    chunks.push({ idx: Number(row[0]), data: String(row[1]) });
    email = String(row[2]).toLowerCase();
  });
  chunks.sort(function (a, b) { return a.idx - b.idx; });
  return { dataUrl: chunks.map(function (c) { return c.data; }).join(''), email: email };
}
