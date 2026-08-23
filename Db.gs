/**
 * Db.gs — このスクリプトが束ねられているスプレッドシート（＝クラスの記録）の
 *         作りの定義と、読み書き。
 *
 * ── この配り方について ────────────────────────────────────────────
 * スプレッドシートのコピーを配り、そのファイルにこのスクリプトが束ねられている
 * （コンテナバインド）。1 ファイル = 1 クラスで、先生ご自身の Google ドライブに
 * 置かれる。中央のレジストリもクラスコードも無い。開いているそのファイルが中身である。
 *
 * ── 列は「番号」ではなく「見出しの名前」で引く ──────────────────────
 * 先生が誤字を直しにシートを開くのは想定内の操作で、その途中で列が 1 本挿さることが
 * ある。列番号で読み書きしていると、そのとき **画面には何も出ないまま**
 * 「返却が別の列に入る」「自分の記録が出ない」が起きる。
 * このファイルの I/O はすべて headerMap_() を通し、見出しの名前で列を決める。
 * だから列を入れ替えても動く。無くなった列だけが問題になり、それは点検で見つかる。
 */

const ROSTER_SHEET_NAME   = '児童名簿';
const JOURNAL_SHEET_NAME  = 'ジャーナルデータ';
const THEME_SHEET_NAME    = 'テーマ設定';
const SETTINGS_SHEET_NAME = '設定';
const IMAGE_SHEET_NAME    = '画像データ';
const META_SHEET_NAME     = '_meta';

/**
 * このアプリが要るシートと見出し。ここに足すと、点検・修整・新規作成の
 * すべてが同じ定義を見る（3 か所に書くと必ずずれる）。
 *
 * defaults … あとから列を足したときに、既存の行へ入れておく値。
 *            空のままだと「状態が空の児童」が全員はじかれるので、名簿だけ要る。
 */
const SHEET_SPECS = [
  {
    name: ROSTER_SHEET_NAME,
    headers: ['役割', '氏名', 'メールアドレス', '状態'],
    defaults: { '状態': 'active' },
    about: '誰がこのクラスの児童か。役割は 担任 / 児童、状態は active / pending。'
  },
  {
    name: JOURNAL_SHEET_NAME,
    headers: ['journalId', 'timestamp', 'email', 'theme', 'content',
              'imageFileId', 'emotion', 'teacherComment', 'highlights',
              'teacherStamp', 'status', 'pastComment', 'deletedAt'],
    about: '児童の提出と、先生のおへんじ。1 行 1 提出。'
  },
  {
    name: THEME_SHEET_NAME,
    headers: ['日付', 'テーマ'],
    about: '日付を決めて配るテーマ。曜日別テーマは「設定」シートに入る。'
  },
  {
    name: SETTINGS_SHEET_NAME,
    headers: ['key', 'value'],
    about: 'クラスの設定。Gemini の API キーや曜日別テーマ。'
  },
  {
    name: IMAGE_SHEET_NAME,
    headers: ['imageId', 'chunkIndex', 'data', 'email', 'createdAt'],
    about: '添付画像。1 セル 50,000 文字の上限があるので、切って複数行に入れる。'
  },
  {
    name: META_SHEET_NAME,
    headers: ['key', 'value'],
    about: 'このファイルがいつ・どの版で作られたかの控え。認可には使わない。'
  }
];

const IMAGE_CHUNK_SIZE = 40000;                 // 1 セル 50,000 文字上限に対する安全マージン
const IMAGE_MAX_DATAURL_CHARS = 800000;         // Data URL 上限（約 600KB 相当の画像）

// ────────────────────────────────────────────────────────────────
// 見出しの読み取り（すべての I/O がここを通る）
// ────────────────────────────────────────────────────────────────

/** 1 行目を読んで { 見出しの名前: 1 始まりの列番号 } を返す。空の見出しは入れない */
function headerMap_(sheet) {
  const width = sheet.getLastColumn();
  if (width < 1) return {};
  const row = sheet.getRange(1, 1, 1, width).getValues()[0];
  const map = {};
  for (let i = 0; i < row.length; i++) {
    const name = String(row[i] === null || row[i] === undefined ? '' : row[i]).trim();
    if (!name) continue;
    if (map[name] === undefined) map[name] = i + 1;   // 重複したら左を採る（点検で報告する）
  }
  return map;
}

/**
 * 見出しの名前から列番号を取る。無ければ「どのシートの何が無いか」を出して止める。
 * 黙って 0 や -1 を返すと、-1 が別の列に化けて他人の行を書き換えることがある。
 */
function colOf_(map, name, sheetName) {
  const col = map[name];
  if (!col) {
    throw new Error('SHEET_BROKEN: 「' + sheetName + '」シートに「' + name + '」の列がありません。' +
      'スプレッドシートのメニュー「ふりかえりジャーナル」＞「シートを点検する」で確かめてください');
  }
  return col;
}

// ────────────────────────────────────────────────────────────────
// 足りないものを作る（作るだけ。既にあるものには触らない）
// ────────────────────────────────────────────────────────────────

function writeHeaderRow_(sheet, spec) {
  sheet.getRange(1, 1, 1, spec.headers.length).setValues([spec.headers]);
  sheet.getRange(1, 1, 1, spec.headers.length)
    .setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
}

/**
 * 足りないシートだけを作る。
 *
 * ふつうは 1 枚も足りないので、その場合はロックを取らずに帰る（40 台が一斉に開く朝に、
 * 全員がロック待ちに並ぶのを避けるため）。先生が誤って 1 枚消したときも、次に開いた人が
 * 作り直す。消えた中身は戻らないが、画面が真っ白になることは無くなる。
 *
 * ⚠️ ここでは見出しの直しをしない。既にあるシートの 1 行目を起動のたびに書き換えると、
 *    列がずれているだけの状態に正しいラベルを貼ってしまい、事故が見えなくなる。
 *    見出しの手当ては repairSheets_()（人が押したときだけ）で行う。
 */
function ensureSheets_(ss) {
  const missing = SHEET_SPECS.filter(function (spec) {
    const sheet = ss.getSheetByName(spec.name);
    return !sheet || sheet.getLastRow() === 0;
  });
  if (!missing.length) return ss;

  withScriptLock_(function () {
    SHEET_SPECS.forEach(function (spec) {
      let sheet = ss.getSheetByName(spec.name);
      if (sheet && sheet.getLastRow() > 0) return;   // ロック待ちの間に誰かが作っていた
      if (!sheet) sheet = ss.insertSheet(spec.name);
      writeHeaderRow_(sheet, spec);
    });
    // コピー元に残っていた「シート1」は、ほかが揃ってから消す
    const blank = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
    if (blank && ss.getSheets().length > SHEET_SPECS.length) {
      try { ss.deleteSheet(blank); } catch (e) {}
    }
  });
  return ss;
}

// ────────────────────────────────────────────────────────────────
// 点検（読むだけ。1 セルも書き換えない）
// ────────────────────────────────────────────────────────────────

/**
 * シートの作りが SHEET_SPECS のとおりかを見る。**何も書き換えない。**
 *
 * 返すのは所見の配列。それぞれに fixable を付けてあり、これが true のものだけを
 * repairSheets_() が直す。false のものは「人が中身を見ないと直せない」ものである。
 *
 * @return {{sheet:string, kind:string, detail:string, fixable:boolean, action:string}[]}
 */
function inspectSheets_(ss) {
  const found = [];
  SHEET_SPECS.forEach(function (spec) {
    const sheet = ss.getSheetByName(spec.name);

    if (!sheet) {
      found.push({
        sheet: spec.name, kind: 'シートが無い',
        detail: spec.about,
        fixable: true, action: '見出しだけのシートを作ります（消えた中身は戻りません）'
      });
      return;
    }
    if (sheet.getLastRow() === 0) {
      found.push({
        sheet: spec.name, kind: '見出しが無い',
        detail: '1 行目が空です',
        fixable: true, action: '1 行目に見出しを書きます'
      });
      return;
    }

    const width = Math.max(sheet.getLastColumn(), 1);
    const actual = sheet.getRange(1, 1, 1, width).getValues()[0].map(function (v) {
      return String(v === null || v === undefined ? '' : v).trim();
    });
    const hasRows = sheet.getLastRow() > 1;

    // 同じ見出しが 2 つあると、どちらを読んでいるか人には見えない。直せない。
    const seen = {};
    const dup = [];
    actual.forEach(function (name) {
      if (!name) return;
      if (seen[name]) { if (dup.indexOf(name) === -1) dup.push(name); }
      seen[name] = true;
    });
    if (dup.length) {
      found.push({
        sheet: spec.name, kind: '見出しが重複している',
        detail: '「' + dup.join('」「') + '」が 2 つ以上あります。左側だけが読まれます',
        fixable: false, action: 'どちらが本物かは中身を見ないと決められません。人が消してください'
      });
    }

    const missing = spec.headers.filter(function (h) { return actual.indexOf(h) === -1; });
    const unknown = actual.filter(function (h) { return h && spec.headers.indexOf(h) === -1; });

    if (missing.length && unknown.length) {
      // 見出しの付け替え（改名）の疑い。足すと、元の列の中身が読まれないまま
      // 空の新しい列ができる。**直さない。**
      found.push({
        sheet: spec.name, kind: '見出しが入れ替わっている疑い',
        detail: '「' + missing.join('」「') + '」が無く、代わりに「' + unknown.join('」「') + '」があります',
        fixable: false,
        action: '名前を変えただけなら、見出しを元の名前に戻してください。' +
                '別の用途で足した列なら、そのままで構いません（点検には毎回出ます）'
      });
    } else if (missing.length) {
      found.push({
        sheet: spec.name, kind: '列が足りない',
        detail: '「' + missing.join('」「') + '」の列がありません',
        fixable: true,
        action: hasRows
          ? '右端に足します（既にある列は動かしません。足した列の中身は空です）'
          : '見出しを書き直します（中身がまだ 1 行もないため）'
      });
    } else if (unknown.length) {
      found.push({
        sheet: spec.name, kind: '知らない列がある',
        detail: '「' + unknown.join('」「') + '」があります',
        fixable: false,
        action: '先生が足した列なら、そのままで大丈夫です。アプリは見出しの名前で列を探すので、じゃまになりません'
      });
    }

    // 並びが違うのは害が無い（見出しの名前で引いているため）。気づけるように出すだけ。
    if (!missing.length) {
      const order = spec.headers.map(function (h) { return actual.indexOf(h); });
      let sorted = true;
      for (let i = 1; i < order.length; i++) if (order[i] < order[i - 1]) sorted = false;
      if (!sorted) {
        found.push({
          sheet: spec.name, kind: '列の並びが違う',
          detail: 'アプリは見出しの名前で列を探すので、このままで動きます',
          fixable: false, action: '直さなくて構いません'
        });
      }
    }

    if (sheet.getFrozenRows() < 1) {
      found.push({
        sheet: spec.name, kind: '見出しが固定されていない',
        detail: '下へたどると見出しが見えなくなります',
        fixable: true, action: '1 行目を固定します'
      });
    }
  });
  return found;
}

// ────────────────────────────────────────────────────────────────
// 修整（人が押したときだけ動く。直せるものだけ直す）
// ────────────────────────────────────────────────────────────────

/**
 * 点検で fixable が true のものだけを直す。
 *
 * ここで守っていること:
 *   - **既にある列を動かさない・名前を変えない。** 足すのは必ず右端。
 *   - **1 セルも消さない。** 余分な列も、知らないシートも残す。
 *   - 直せなかったものは left に入れて返す。黙って握りつぶさない。
 *
 * @return {{fixed:string[], left:{sheet:string,kind:string,detail:string,action:string}[]}}
 */
function repairSheets_(ss) {
  const fixed = [];
  const before = inspectSheets_(ss);
  if (!before.length) return { fixed: fixed, left: [] };

  withScriptLock_(function () {
    SHEET_SPECS.forEach(function (spec) {
      let sheet = ss.getSheetByName(spec.name);

      if (!sheet) {
        sheet = ss.insertSheet(spec.name);
        writeHeaderRow_(sheet, spec);
        fixed.push('「' + spec.name + '」シートを作りました');
        return;
      }
      if (sheet.getLastRow() === 0) {
        writeHeaderRow_(sheet, spec);
        fixed.push('「' + spec.name + '」に見出しを書きました');
        return;
      }

      const width = Math.max(sheet.getLastColumn(), 1);
      const actual = sheet.getRange(1, 1, 1, width).getValues()[0].map(function (v) {
        return String(v === null || v === undefined ? '' : v).trim();
      });
      const missing = spec.headers.filter(function (h) { return actual.indexOf(h) === -1; });
      const unknown = actual.filter(function (h) { return h && spec.headers.indexOf(h) === -1; });

      // 入れ替えの疑いがあるときは触らない（inspectSheets_ と同じ判断）
      if (missing.length && !unknown.length) {
        if (sheet.getLastRow() <= 1) {
          // 中身がまだ無いので、見出しを丸ごと正す
          sheet.getRange(1, 1, 1, Math.max(width, spec.headers.length)).clearContent();
          writeHeaderRow_(sheet, spec);
          fixed.push('「' + spec.name + '」の見出しを書き直しました（中身はまだありませんでした）');
        } else {
          // 右端に足す。既にある列は 1 つも動かさない
          let col = width + 1;
          const rows = sheet.getLastRow() - 1;
          missing.forEach(function (name) {
            if (col > sheet.getMaxColumns()) sheet.insertColumnsAfter(sheet.getMaxColumns(), 1);
            sheet.getRange(1, col).setValue(name)
              .setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
            const fill = spec.defaults && spec.defaults[name];
            if (fill && rows > 0) {
              const values = [];
              for (let i = 0; i < rows; i++) values.push([fill]);
              sheet.getRange(2, col, rows, 1).setValues(values);
            }
            col++;
          });
          fixed.push('「' + spec.name + '」の右端に「' + missing.join('」「') + '」を足しました'
            + (spec.defaults ? '（既にある行には既定の値を入れました）' : ''));
        }
      }

      if (sheet.getFrozenRows() < 1) {
        sheet.setFrozenRows(1);
        fixed.push('「' + spec.name + '」の 1 行目を固定しました');
      }
    });
  });

  return { fixed: fixed, left: inspectSheets_(ss).filter(function (f) { return !f.fixable; }) };
}

// ────────────────────────────────────────────────────────────────
// _meta / key-value シート
// ────────────────────────────────────────────────────────────────

/** _meta シートへ控えを書く（デバッグ・復旧用。認可には使わない） */
function writeMeta_(ss, meta) {
  const sheet = ss.getSheetByName(META_SHEET_NAME);
  if (!sheet) return;
  Object.keys(meta).forEach(function (k) { upsertKeyValue_(sheet, k, String(meta[k])); });
}

function readMeta_(ss, key) {
  return readKeyValue_(ss.getSheetByName(META_SHEET_NAME), key);
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
// クラス内設定（設定シート）— 先生の Gemini API キー・曜日別テーマ等
// ────────────────────────────────────────────────────────────────

function getTenantSetting_(ss, key) {
  return readKeyValue_(ss.getSheetByName(SETTINGS_SHEET_NAME), key);
}

function setTenantSetting_(ss, key, value) {
  const sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!sheet) throw new Error('SHEET_BROKEN: 「' + SETTINGS_SHEET_NAME + '」シートがありません');
  upsertKeyValue_(sheet, key, value);
}

// ────────────────────────────────────────────────────────────────
// 名簿 I/O（すべて見出しの名前で列を引く）
// ────────────────────────────────────────────────────────────────

function getRosterSheet_(ss) {
  const sheet = ss.getSheetByName(ROSTER_SHEET_NAME);
  if (!sheet) throw new Error('SHEET_BROKEN: 「' + ROSTER_SHEET_NAME + '」シートがありません');
  return sheet;
}

/** 名簿全行を { role, name, email, status } で返す（メールアドレスが空の行は除外） */
function getRosterRows_(ss) {
  const sheet = ss.getSheetByName(ROSTER_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const map = headerMap_(sheet);
  const cRole = colOf_(map, '役割', ROSTER_SHEET_NAME);
  const cName = colOf_(map, '氏名', ROSTER_SHEET_NAME);
  const cMail = colOf_(map, 'メールアドレス', ROSTER_SHEET_NAME);
  const cStat = colOf_(map, '状態', ROSTER_SHEET_NAME);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  return data.filter(function (r) { return r[cMail - 1]; }).map(function (r) {
    return {
      role: r[cRole - 1] === '担任' ? '担任' : '児童',
      name: String(r[cName - 1] || ''),
      email: String(r[cMail - 1]).trim().toLowerCase(),
      status: r[cStat - 1] === 'pending' ? 'pending' : 'active'
    };
  });
}

/** 名簿を丸ごと書き換える（担任 API 用）。呼び出し側でロックすること */
function saveRosterRows_(ss, rows) {
  const sheet = getRosterSheet_(ss);
  const map = headerMap_(sheet);
  const cols = ['役割', '氏名', 'メールアドレス', '状態'].map(function (n) {
    return colOf_(map, n, ROSTER_SHEET_NAME);
  });
  if (sheet.getLastRow() > 1) {
    cols.forEach(function (c) { sheet.getRange(2, c, sheet.getLastRow() - 1, 1).clearContent(); });
  }
  rows.forEach(function (r, i) {
    const values = [
      r.role === '担任' ? '担任' : '児童',
      r.name || '',
      String(r.email).trim().toLowerCase(),
      r.status === 'pending' ? 'pending' : 'active'
    ];
    cols.forEach(function (c, k) { sheet.getRange(2 + i, c).setValue(values[k]); });
  });
}

/** 1 名分を追加/更新する。メールアドレスをキーに既存行を上書きする */
function upsertMember_(ss, member) {
  const sheet = getRosterSheet_(ss);
  const map = headerMap_(sheet);
  const cRole = colOf_(map, '役割', ROSTER_SHEET_NAME);
  const cName = colOf_(map, '氏名', ROSTER_SHEET_NAME);
  const cMail = colOf_(map, 'メールアドレス', ROSTER_SHEET_NAME);
  const cStat = colOf_(map, '状態', ROSTER_SHEET_NAME);
  const email = String(member.email).trim().toLowerCase();
  const role = member.role === '担任' ? '担任' : '児童';
  const status = member.status === 'pending' ? 'pending' : 'active';

  let row = 0;
  if (sheet.getLastRow() >= 2) {
    const cell = sheet.getRange(2, cMail, sheet.getLastRow() - 1, 1)
      .createTextFinder(email).matchEntireCell(true).findNext();
    if (cell) row = cell.getRow();
  }
  if (!row) row = sheet.getLastRow() + 1;
  sheet.getRange(row, cRole).setValue(role);
  sheet.getRange(row, cName).setValue(member.name || '');
  sheet.getRange(row, cMail).setValue(email);
  sheet.getRange(row, cStat).setValue(status);
}

/** 名簿から 1 名を削除する（参加申請の却下・除名） */
function removeMember_(ss, email) {
  const sheet = ss.getSheetByName(ROSTER_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return false;
  const cMail = colOf_(headerMap_(sheet), 'メールアドレス', ROSTER_SHEET_NAME);
  const cell = sheet.getRange(2, cMail, sheet.getLastRow() - 1, 1)
    .createTextFinder(String(email).trim().toLowerCase()).matchEntireCell(true).findNext();
  if (!cell) return false;
  sheet.deleteRow(cell.getRow());
  return true;
}

// ────────────────────────────────────────────────────────────────
// ジャーナル I/O
// ────────────────────────────────────────────────────────────────

function getJournalSheet_(ss) {
  const sheet = ss.getSheetByName(JOURNAL_SHEET_NAME);
  if (!sheet) {
    throw new Error('SHEET_BROKEN: 「' + JOURNAL_SHEET_NAME + '」シートがありません。' +
      'スプレッドシートのメニュー「ふりかえりジャーナル」＞「シートを点検する」で確かめてください');
  }
  return sheet;
}

/** ジャーナルシートの { 見出し: 列番号 }。書き込み前に必ずこれで列を決める */
function journalCols_(sheet) {
  const map = headerMap_(sheet);
  const need = SHEET_SPECS.filter(function (s) { return s.name === JOURNAL_SHEET_NAME; })[0].headers;
  const cols = {};
  need.forEach(function (n) { cols[n] = colOf_(map, n, JOURNAL_SHEET_NAME); });
  return cols;
}

/** journalId から行番号を探す（TextFinder は全行走査より大幅に速い） */
function findJournalRowById_(sheet, journalId) {
  if (!journalId || sheet.getLastRow() < 2) return -1;
  const cId = colOf_(headerMap_(sheet), 'journalId', JOURNAL_SHEET_NAME);
  const cell = sheet.getRange(2, cId, sheet.getLastRow() - 1, 1)
    .createTextFinder(String(journalId)).matchEntireCell(true).findNext();
  return cell ? cell.getRow() : -1;
}

function rowToJournalObject_(headers, row) {
  const journal = {};
  headers.forEach(function (h, i) { if (h) journal[h] = row[i]; });
  if (journal.timestamp instanceof Date) {
    journal.date = Utilities.formatDate(journal.timestamp, 'JST', 'yyyy/MM/dd HH:mm');
  }
  // 画像はシート内に切って入れてある。ID だけ返し、本体は getImage で遅らせて取る
  journal.imageId = journal.imageFileId || '';
  delete journal.imageFileId;
  return journal;
}

/** 未削除の全ジャーナルをオブジェクト配列で返す */
function getJournalsAll_(ss) {
  const sheet = ss.getSheetByName(JOURNAL_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data.shift().map(function (v) { return String(v || '').trim(); });
  const deletedAtIdx = headers.indexOf('deletedAt');
  return data.filter(function (row) { return deletedAtIdx < 0 || !row[deletedAtIdx]; })
    .map(function (row) { return rowToJournalObject_(headers, row); });
}

function getJournalsForEmail_(ss, email) {
  const target = String(email).toLowerCase();
  return getJournalsAll_(ss).filter(function (j) {
    return String(j.email).toLowerCase() === target;
  });
}

/** 1 行分を、見出しの名前どおりの位置へ書く（列が入れ替わっていても正しい列に入る） */
function appendJournalRow_(sheet, values) {
  const cols = journalCols_(sheet);
  const width = Math.max(sheet.getLastColumn(), 1);
  const row = new Array(width).fill('');
  Object.keys(values).forEach(function (key) {
    const col = cols[key];
    if (col) row[col - 1] = values[key];
  });
  sheet.appendRow(row);
}

// ────────────────────────────────────────────────────────────────
// テーマ I/O
// ────────────────────────────────────────────────────────────────

function getTodayTheme_(ss) {
  const sheet = ss.getSheetByName(THEME_SHEET_NAME);
  const todayStr = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd');
  if (sheet && sheet.getLastRow() >= 2) {
    const map = headerMap_(sheet);
    const cDate = map['日付'];
    const cTheme = map['テーマ'];
    if (cDate && cTheme) {
      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
      for (let i = data.length - 1; i >= 0; i--) {
        const d = data[i][cDate - 1];
        if (d instanceof Date && Utilities.formatDate(d, 'JST', 'yyyy/MM/dd') === todayStr) {
          return String(data[i][cTheme - 1]);
        }
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
// 画像 I/O（Data URL を切ってシートに入れる）
// ────────────────────────────────────────────────────────────────

/**
 * Data URL を確かめてから切って入れ、imageId を返す。呼び出し側でロックすること。
 * email はアップロードした人（確認済み）。取り出すときの持ち主の照合に使う。
 */
function saveImageChunks_(ss, dataUrl, email) {
  const s = String(dataUrl || '');
  if (!/^data:image\/(png|jpeg|jpg|webp|gif);base64,[A-Za-z0-9+\/=]+$/.test(s)) {
    throw new Error('BAD_INPUT: 画像データの形式が正しくありません');
  }
  if (s.length > IMAGE_MAX_DATAURL_CHARS) {
    throw new Error('BAD_INPUT: 画像が大きすぎます。小さい画像でためしてね');
  }
  const sheet = ss.getSheetByName(IMAGE_SHEET_NAME);
  if (!sheet) throw new Error('SHEET_BROKEN: 「' + IMAGE_SHEET_NAME + '」シートがありません');
  const map = headerMap_(sheet);
  const cId = colOf_(map, 'imageId', IMAGE_SHEET_NAME);
  const cIdx = colOf_(map, 'chunkIndex', IMAGE_SHEET_NAME);
  const cData = colOf_(map, 'data', IMAGE_SHEET_NAME);
  const cMail = colOf_(map, 'email', IMAGE_SHEET_NAME);
  const cAt = colOf_(map, 'createdAt', IMAGE_SHEET_NAME);

  const imageId = 'img_' + Utilities.getUuid();
  const now = new Date();
  const width = Math.max(sheet.getLastColumn(), 1);
  const rows = [];
  for (let i = 0, idx = 0; i < s.length; i += IMAGE_CHUNK_SIZE, idx++) {
    const row = new Array(width).fill('');
    row[cId - 1] = imageId;
    row[cIdx - 1] = idx;
    row[cData - 1] = s.substring(i, i + IMAGE_CHUNK_SIZE);
    row[cMail - 1] = String(email).toLowerCase();
    row[cAt - 1] = now;
    rows.push(row);
  }
  const start = sheet.getLastRow() + 1;
  const shortage = start + rows.length - 1 - sheet.getMaxRows();
  if (shortage > 0) sheet.insertRowsAfter(sheet.getMaxRows(), shortage);
  sheet.getRange(start, 1, rows.length, width).setValues(rows);
  return imageId;
}

/** imageId から Data URL を戻して { dataUrl, email } を返す。無ければ null */
function loadImage_(ss, imageId) {
  const sheet = ss.getSheetByName(IMAGE_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return null;
  if (!/^img_[A-Za-z0-9\-]{8,50}$/.test(String(imageId || ''))) return null;
  const map = headerMap_(sheet);
  const cId = colOf_(map, 'imageId', IMAGE_SHEET_NAME);
  const cIdx = colOf_(map, 'chunkIndex', IMAGE_SHEET_NAME);
  const cData = colOf_(map, 'data', IMAGE_SHEET_NAME);
  const cMail = colOf_(map, 'email', IMAGE_SHEET_NAME);
  // 画像シートは大きくなりやすいので全部は読まず、TextFinder で当たった行だけ読む
  const cells = sheet.getRange(2, cId, sheet.getLastRow() - 1, 1)
    .createTextFinder(String(imageId)).matchEntireCell(true).findAll();
  if (cells.length === 0) return null;
  const width = sheet.getLastColumn();
  const chunks = [];
  let email = '';
  cells.forEach(function (cell) {
    const row = sheet.getRange(cell.getRow(), 1, 1, width).getValues()[0];
    chunks.push({ idx: Number(row[cIdx - 1]), data: String(row[cData - 1]) });
    email = String(row[cMail - 1]).toLowerCase();
  });
  chunks.sort(function (a, b) { return a.idx - b.idx; });
  return { dataUrl: chunks.map(function (c) { return c.data; }).join(''), email: email };
}
