/**
 * コンテナバインド版（.gs）のテスト。
 *
 * GAS のランタイムは手元で動かせないので、SpreadsheetApp / Session / LockService /
 * Utilities を偽物に差しかえ、.gs のソースをそのまま実行して中の判断を確かめる。
 * 関数を正規表現で切り出す方式は使わない（書き方を少し変えただけで
 * 「読み取れませんでした」と落ち、検査していないことに気づけなくなる）。
 *
 * ここで見ているのは、直したつもりで壊れる場所:
 *   1. シートの点検が、ずれの種類を正しく見分けるか（勝手に直してよいものと、そうでないもの）
 *   2. 修整が、既にある列を動かさず・消さないか
 *   3. 列が入れ替わっていても、見出しの名前で正しい列に読み書きするか
 *   4. 児童が他人の記録に手を出せないか
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const GS = ['Main.gs', 'Db.gs', 'Access.gs', 'MemberApi.gs', 'OwnerApi.gs']
  .map((f) => fs.readFileSync(path.join(ROOT, f), 'utf8'));

// ────────────────────────────────────────────────────────────────
// 偽物のスプレッドシート
// ────────────────────────────────────────────────────────────────

/** 二次元配列 1 枚を Sheet に見立てる。使っている API だけを実装する */
function makeSheet(name, rows) {
  const grid = rows.map((r) => r.slice());
  let frozen = 0;
  const width = () => grid.reduce((m, r) => Math.max(m, r.length), 0);
  const padTo = (w) => grid.forEach((r) => { while (r.length < w) r.push(''); });

  const sheet = {
    _name: name,
    _grid: grid,
    getName: () => name,
    setName: (n) => { sheet._name = n; return sheet; },
    getLastRow: () => {
      for (let i = grid.length - 1; i >= 0; i--) {
        if (grid[i].some((c) => c !== '' && c !== null && c !== undefined)) return i + 1;
      }
      return 0;
    },
    getLastColumn: () => {
      let last = 0;
      grid.forEach((r) => r.forEach((c, i) => { if (c !== '' && c !== null && c !== undefined) last = Math.max(last, i + 1); }));
      return last;
    },
    getMaxRows: () => Math.max(grid.length, 1),
    getMaxColumns: () => Math.max(width(), 1),
    getFrozenRows: () => frozen,
    setFrozenRows: (n) => { frozen = n; return sheet; },
    insertRowsAfter: (after, howMany) => {
      for (let i = 0; i < howMany; i++) grid.push(new Array(width()).fill(''));
      return sheet;
    },
    insertColumnsAfter: (after, howMany) => { padTo(width() + howMany); return sheet; },
    appendRow: (values) => {
      const at = sheet.getLastRow();
      while (grid.length <= at) grid.push(new Array(width()).fill(''));
      grid[at] = values.slice();
      padTo(Math.max(width(), values.length));
      return sheet;
    },
    deleteRow: (row) => { grid.splice(row - 1, 1); return sheet; },
    deleteRows: (row, howMany) => { grid.splice(row - 1, howMany); return sheet; },
    getDataRange: () => sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), Math.max(sheet.getLastColumn(), 1)),
    getRange: (row, col, numRows = 1, numCols = 1) => {
      const range = {
        getRow: () => row,
        getValues: () => {
          const out = [];
          for (let r = 0; r < numRows; r++) {
            const line = [];
            for (let c = 0; c < numCols; c++) {
              const v = (grid[row - 1 + r] || [])[col - 1 + c];
              line.push(v === undefined ? '' : v);
            }
            out.push(line);
          }
          return out;
        },
        getValue: () => range.getValues()[0][0],
        setValues: (values) => {
          values.forEach((line, r) => {
            const target = row - 1 + r;
            while (grid.length <= target) grid.push([]);
            line.forEach((v, c) => { grid[target][col - 1 + c] = v; });
          });
          padTo(width());
          return range;
        },
        setValue: (v) => range.setValues([[v]]),
        clearContent: () => {
          for (let r = 0; r < numRows; r++) {
            for (let c = 0; c < numCols; c++) {
              if (grid[row - 1 + r]) grid[row - 1 + r][col - 1 + c] = '';
            }
          }
          return range;
        },
        setFontWeight: () => range, setBackground: () => range, setFontColor: () => range,
        createTextFinder: (needle) => {
          const hits = [];
          for (let r = 0; r < numRows; r++) {
            for (let c = 0; c < numCols; c++) {
              const v = (grid[row - 1 + r] || [])[col - 1 + c];
              if (String(v === undefined ? '' : v) === String(needle)) {
                hits.push({ getRow: () => row + r, getColumn: () => col + c });
              }
            }
          }
          return {
            matchEntireCell: () => ({ findNext: () => hits[0] || null, findAll: () => hits }),
            findNext: () => hits[0] || null,
            findAll: () => hits
          };
        }
      };
      return range;
    }
  };
  return sheet;
}

function makeSpreadsheet(sheetDefs) {
  const sheets = sheetDefs.map(([name, rows]) => makeSheet(name, rows));
  return {
    _sheets: sheets,
    getId: () => 'ss-test',
    getUrl: () => 'https://docs.google.com/spreadsheets/d/ss-test/edit',
    getSheets: () => sheets,
    getSheetByName: (n) => sheets.find((s) => s._name === n) || null,
    insertSheet: (n) => { const s = makeSheet(n, [[]]); sheets.push(s); return s; },
    deleteSheet: (s) => { const i = sheets.indexOf(s); if (i >= 0) sheets.splice(i, 1); }
  };
}

/** 見出しだけがそろった、正常なファイルを作る */
function healthySheets() {
  return [
    ['児童名簿', [['役割', '氏名', 'メールアドレス', '状態']]],
    ['ジャーナルデータ', [['journalId', 'timestamp', 'email', 'theme', 'content', 'imageFileId',
                          'emotion', 'teacherComment', 'highlights', 'teacherStamp', 'status',
                          'pastComment', 'deletedAt']]],
    ['テーマ設定', [['日付', 'テーマ']]],
    ['設定', [['key', 'value']]],
    ['画像データ', [['imageId', 'chunkIndex', 'data', 'email', 'createdAt']]],
    ['_meta', [['key', 'value']]]
  ];
}

/**
 * GAS の HtmlService を、ここで事故になった点だけ忠実に真似る。
 *
 * ★ なぜ忠実さが要るか ★
 * 2026-08-23、doGet が `HtmlService.XFrameOptionsMode.SAMEORIGIN` を渡していた。
 * この列挙子は **存在しない**（ALLOWALL と DEFAULT の 2 つだけ）ので undefined になり、
 * GAS は「引数は null にできません: mode」で落ちた。**画面が一切開かない**状態である。
 * 当時の偽物は `evaluate: () => ({})` で、setXFrameOptionsMode を呼びもしなかったため
 * テストは 23 件すべて緑のまま通した。**偽物が本物より寛容だと、テストは嘘をつく。**
 */
function makeHtmlService(rendered) {
  const requireArg = (name, value) => {
    // GAS と同じ形で落とす。undefined を黙って受けると、この事故を再び見逃す。
    if (value === undefined || value === null) {
      throw new Error('引数は null にできません: ' + name);
    }
    return value;
  };
  const makeOutput = (template) => {
    const out = {
      _template: template,
      setTitle: (t) => { rendered.title = requireArg('title', t); return out; },
      setXFrameOptionsMode: (mode) => { rendered.xFrameOptionsMode = requireArg('mode', mode); return out; },
      addMetaTag: (name, content) => {
        requireArg('name', name);
        (rendered.metaTags = rendered.metaTags || {})[name] = requireArg('content', content);
        return out;
      },
      getContent: () => ''
    };
    return out;
  };
  return {
    // ★ 実在する 2 つだけ。SAMEORIGIN を書けば undefined になり、上で落ちる。
    XFrameOptionsMode: Object.freeze({ ALLOWALL: 'ALLOWALL', DEFAULT: 'DEFAULT' }),
    SandboxMode: Object.freeze({ IFRAME: 'IFRAME' }),
    createTemplateFromFile: (name) => {
      const file = requireArg('file', name);
      const template = {
        _file: file,
        // 本物と同じく、ファイルの中身をそのまま返す（解釈を挟まない）
        getRawContent: () => fs.readFileSync(path.join(ROOT, file + '.html'), 'utf8')
      };
      template.evaluate = () => { rendered.template = { ...template }; return makeOutput(template); };
      return template;
    },
    // ★ わざと落とす。本物はここで中身を「HTML として読み直して組み立て直す」。
    //   その読み直しが app.html の中の JavaScript を壊し、2026-08-24 に
    //   画面が 1 つも出なくなった。手元にはその読み直しが無いので、
    //   偽物を素通りさせると「手元だけ緑」が再び起きる。使わせない。
    createHtmlOutputFromFile: () => {
      throw new Error(
        'createHtmlOutputFromFile は使わないこと（中身が HTML として読み直され、' +
        'app.html の中の JavaScript が壊れる）。createTemplateFromFile(...).getRawContent() を使う。');
    }
  };
}

let uuidSeq = 0;

/** .gs を読み込んで、中の関数を取り出す */
function load({ sheetDefs = healthySheets(), activeEmail = 'sensei@school.example',
                effectiveEmail = 'sensei@school.example' } = {}) {
  uuidSeq = 0;
  const ss = makeSpreadsheet(sheetDefs);
  const rendered = {};
  const sandbox = {
    console,
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ss,
      getUi: () => { throw new Error('画面がありません'); }   // ウェブアプリ文脈と同じ
    },
    Session: {
      getActiveUser: () => ({ getEmail: () => activeEmail }),
      getEffectiveUser: () => ({ getEmail: () => effectiveEmail })
    },
    LockService: {
      getScriptLock: () => ({ tryLock: () => true, waitLock: () => true, releaseLock: () => {} })
    },
    ScriptApp: { getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/AAA/exec' }) },
    Utilities: {
      getUuid: () => 'uuid-' + (++uuidSeq).toString().padStart(4, '0') + '-aaaa-bbbb',
      formatDate: (d, tz, fmt) => new Date(d).toISOString().slice(0, 10).replace(/-/g, '/'),
      computeDigest: (alg, s) => Array.from(String(s)).map((c) => c.charCodeAt(0) % 256),
      DigestAlgorithm: { SHA_256: 'SHA_256' },
      Charset: { UTF_8: 'UTF_8' }
    },
    HtmlService: makeHtmlService(rendered),
    GigaGemini: { callAll: () => [], callRaw: () => ({ ok: true, text: 'OK' }), parseJsonText: () => ({}) }
  };
  vm.createContext(sandbox);
  for (const src of GS) vm.runInContext(src, sandbox);
  return { sandbox, ss, rendered };
}

const call = (sandbox, name, ...args) => JSON.parse(vm.runInContext(
  `${name}(${args.map((a) => JSON.stringify(a)).join(',')})`, sandbox));

/** すでに 1 回開かれ、先生が控えられている状態にする */
function setUp(sandbox, ss) {
  const meta = ss.getSheetByName('_meta');
  meta.appendRow(['ownerEmail', 'sensei@school.example']);
  meta.appendRow(['className', '3年2組']);
  meta.appendRow(['fileId', 'ss-test']);
  const roster = ss.getSheetByName('児童名簿');
  roster.appendRow(['担任', '先生', 'sensei@school.example', 'active']);
}

// ════════════════════════════════════════════════════════════════
// 1. シートの点検
// ════════════════════════════════════════════════════════════════

test('そろっているファイルでは、点検が何も言わない', () => {
  const { sandbox, ss } = load();
  ss.getSheets().forEach((s) => s.setFrozenRows(1));
  // vm の中で作られた配列は別 realm のプロトタイプを持つので、
  // strict な deepEqual は「構造は同じだが参照が違う」で落ちる。中身だけ見る。
  const found = vm.runInContext('inspectSheets_(getDb_())', sandbox);
  assert.equal(found.length, 0, [...found].map((f) => `${f.sheet}:${f.kind}`).join(' / '));
});

test('シートが1枚消えていたら、自動で直せるものとして報告する', () => {
  const defs = healthySheets().filter(([n]) => n !== 'テーマ設定');
  const { sandbox, ss } = load({ sheetDefs: defs });
  ss.getSheets().forEach((s) => s.setFrozenRows(1));
  // getDb_() を通さず、作り直す前の状態を見る
  const found = vm.runInContext('inspectSheets_(SpreadsheetApp.getActiveSpreadsheet())', sandbox);
  const hit = found.find((f) => f.sheet === 'テーマ設定');
  assert.equal(hit.kind, 'シートが無い');
  assert.equal(hit.fixable, true);
});

test('列が足りないだけなら自動で直せる。名前が変わっていたら人に回す', () => {
  // ① 「状態」の列だけが無い（旧スキーマ）→ 足せる
  let defs = healthySheets();
  defs[0][1] = [['役割', '氏名', 'メールアドレス']];
  let { sandbox, ss } = load({ sheetDefs: defs });
  ss.getSheets().forEach((s) => s.setFrozenRows(1));
  let hit = vm.runInContext('inspectSheets_(getDb_())', sandbox).find((f) => f.sheet === '児童名簿');
  assert.equal(hit.kind, '列が足りない');
  assert.equal(hit.fixable, true);

  // ② 「状態」が「ステータス」に変えられている → 足すと元の列が読まれなくなるので直さない
  defs = healthySheets();
  defs[0][1] = [['役割', '氏名', 'メールアドレス', 'ステータス']];
  ({ sandbox, ss } = load({ sheetDefs: defs }));
  ss.getSheets().forEach((s) => s.setFrozenRows(1));
  hit = vm.runInContext('inspectSheets_(getDb_())', sandbox).find((f) => f.sheet === '児童名簿');
  assert.equal(hit.kind, '見出しが入れ替わっている疑い');
  assert.equal(hit.fixable, false);
});

test('列を並べ替えただけなら「直さなくてよい」と言う', () => {
  const defs = healthySheets();
  defs[0][1] = [['メールアドレス', '氏名', '状態', '役割']];
  const { sandbox, ss } = load({ sheetDefs: defs });
  ss.getSheets().forEach((s) => s.setFrozenRows(1));
  const hit = vm.runInContext('inspectSheets_(getDb_())', sandbox).find((f) => f.sheet === '児童名簿');
  assert.equal(hit.kind, '列の並びが違う');
  assert.equal(hit.fixable, false);
  assert.match(hit.action, /直さなくて/);
});

test('先生が足したメモ列は、消さずに「そのままで大丈夫」と言う', () => {
  const defs = healthySheets();
  defs[0][1] = [['役割', '氏名', 'メールアドレス', '状態', 'メモ']];
  const { sandbox, ss } = load({ sheetDefs: defs });
  ss.getSheets().forEach((s) => s.setFrozenRows(1));
  const hit = vm.runInContext('inspectSheets_(getDb_())', sandbox).find((f) => f.sheet === '児童名簿');
  assert.equal(hit.kind, '知らない列がある');
  assert.equal(hit.fixable, false);
});

test('同じ見出しが2つあったら、自動では直さない', () => {
  const defs = healthySheets();
  defs[0][1] = [['役割', '氏名', 'メールアドレス', '状態', '氏名']];
  const { sandbox, ss } = load({ sheetDefs: defs });
  ss.getSheets().forEach((s) => s.setFrozenRows(1));
  const hit = vm.runInContext('inspectSheets_(getDb_())', sandbox)
    .find((f) => f.kind === '見出しが重複している');
  assert.ok(hit, '重複を見つけられていません');
  assert.equal(hit.fixable, false);
});

test('点検は1セルも書き換えない', () => {
  const defs = healthySheets();
  defs[0][1] = [['役割', '氏名', 'メールアドレス']];      // わざとずらす
  const { sandbox, ss } = load({ sheetDefs: defs });
  const before = JSON.stringify(ss.getSheetByName('児童名簿')._grid);
  vm.runInContext('inspectSheets_(SpreadsheetApp.getActiveSpreadsheet())', sandbox);
  assert.equal(JSON.stringify(ss.getSheetByName('児童名簿')._grid), before);
});

// ════════════════════════════════════════════════════════════════
// 2. シートの修整
// ════════════════════════════════════════════════════════════════

test('足りない列は右端に足し、既にある列も中身も動かさない', () => {
  const defs = healthySheets();
  defs[0][1] = [
    ['役割', '氏名', 'メールアドレス'],
    ['児童', 'ゆうと', 'yuto@school.example'],
    ['児童', 'あおい', 'aoi@school.example']
  ];
  const { sandbox, ss } = load({ sheetDefs: defs });
  const result = vm.runInContext('repairSheets_(getDb_())', sandbox);
  const grid = ss.getSheetByName('児童名簿')._grid;

  assert.deepEqual(grid[0].slice(0, 4), ['役割', '氏名', 'メールアドレス', '状態']);
  // 既にある行の中身は 1 つも動いていない
  assert.deepEqual(grid[1].slice(0, 3), ['児童', 'ゆうと', 'yuto@school.example']);
  assert.deepEqual(grid[2].slice(0, 3), ['児童', 'あおい', 'aoi@school.example']);
  // 足した列には既定値（空だと全員がはじかれる）
  assert.equal(grid[1][3], 'active');
  assert.equal(grid[2][3], 'active');
  assert.ok(result.fixed.some((s) => s.includes('状態')), result.fixed.join(' / '));
});

test('見出しの名前が変えられている疑いがあるときは、修整も手を出さない', () => {
  const defs = healthySheets();
  defs[0][1] = [
    ['役割', '氏名', 'メールアドレス', 'ステータス'],
    ['児童', 'ゆうと', 'yuto@school.example', 'active']
  ];
  const { sandbox, ss } = load({ sheetDefs: defs });
  const result = vm.runInContext('repairSheets_(getDb_())', sandbox);
  const grid = ss.getSheetByName('児童名簿')._grid;

  assert.deepEqual(grid[0].slice(0, 4), ['役割', '氏名', 'メールアドレス', 'ステータス']);
  assert.ok(result.left.some((f) => f.kind === '見出しが入れ替わっている疑い'),
    '直せなかったことを返していません');
});

test('修整はシートも行も消さない', () => {
  const defs = healthySheets();
  defs[0][1] = [['役割', '氏名', 'メールアドレス', '状態', 'メモ'],
                ['児童', 'ゆうと', 'yuto@school.example', 'active', '転入']];
  defs.push(['先生のメモ', [['自由に使う欄']]]);
  const { sandbox, ss } = load({ sheetDefs: defs });
  vm.runInContext('repairSheets_(getDb_())', sandbox);

  assert.ok(ss.getSheetByName('先生のメモ'), '知らないシートを消しています');
  const grid = ss.getSheetByName('児童名簿')._grid;
  assert.equal(grid[0][4], 'メモ');
  assert.equal(grid[1][4], '転入');
});

// ════════════════════════════════════════════════════════════════
// 3. 列が入れ替わっていても、正しい列に読み書きする
// ════════════════════════════════════════════════════════════════

test('ジャーナルの列を入れ替えても、提出は正しい列に入る', () => {
  const defs = healthySheets();
  // 先生が journalId の前に 1 列挿し、いくつか並べ替えた状態
  defs[1][1] = [['メモ', 'email', 'content', 'journalId', 'timestamp', 'theme', 'imageFileId',
                 'emotion', 'teacherComment', 'highlights', 'teacherStamp', 'status',
                 'pastComment', 'deletedAt']];
  const { sandbox, ss } = load({ sheetDefs: defs, activeEmail: 'yuto@school.example' });
  setUp(sandbox, ss);
  ss.getSheetByName('児童名簿').appendRow(['児童', 'ゆうと', 'yuto@school.example', 'active']);

  const res = call(sandbox, 'mbSaveJournal', { content: 'きょうは分数がわかった', emotion: '😊' });
  assert.equal(res.success, true, res.error);

  const grid = ss.getSheetByName('ジャーナルデータ')._grid;
  const head = grid[0];
  const row = grid[1];
  assert.equal(row[head.indexOf('email')], 'yuto@school.example');
  assert.equal(row[head.indexOf('content')], 'きょうは分数がわかった');
  assert.equal(row[head.indexOf('status')], '未返却');
  assert.equal(row[head.indexOf('メモ')], '', '知らない列に書き込んでいます');
  assert.match(String(row[head.indexOf('journalId')]), /^uuid-/);
});

test('名簿の列を入れ替えても、役割と状態を取り違えない', () => {
  const defs = healthySheets();
  defs[0][1] = [['状態', 'メールアドレス', '役割', '氏名'],
                ['pending', 'aoi@school.example', '児童', 'あおい']];
  const { sandbox, ss } = load({ sheetDefs: defs });
  const rows = vm.runInContext('getRosterRows_(getDb_())', sandbox);
  assert.equal(rows.length, 1);
  assert.equal(rows[0].email, 'aoi@school.example');
  assert.equal(rows[0].role, '児童');
  assert.equal(rows[0].status, 'pending');
  assert.equal(rows[0].name, 'あおい');
});

test('必要な列が無いまま読み書きすると、何が無いかを名指しで止める', () => {
  const defs = healthySheets();
  defs[1][1] = [['journalId', 'timestamp', 'email']];   // content 以降が無い
  const { sandbox, ss } = load({ sheetDefs: defs, activeEmail: 'yuto@school.example' });
  setUp(sandbox, ss);
  ss.getSheetByName('児童名簿').appendRow(['児童', 'ゆうと', 'yuto@school.example', 'active']);
  const res = call(sandbox, 'mbSaveJournal', { content: 'てすと', emotion: '😊' });
  assert.equal(res.success, false);
  assert.equal(res.code, 'SHEET_BROKEN');
  assert.match(res.error, /ジャーナルデータ/);
  assert.match(res.error, /theme|content/);   // 足りない列を名指しする
});

// ════════════════════════════════════════════════════════════════
// 4. 認可
// ════════════════════════════════════════════════════════════════

test('ログインが確かめられないときは、何も書かずに理由を返す', () => {
  const { sandbox, ss } = load({ activeEmail: '' });
  setUp(sandbox, ss);
  const res = call(sandbox, 'mbSaveJournal', { content: 'なりすまし', emotion: '😊' });
  assert.equal(res.success, false);
  assert.equal(res.code, 'NO_IDENTITY');
  assert.equal(ss.getSheetByName('ジャーナルデータ').getLastRow(), 1, '行が増えています');
});

test('名簿にいない児童は提出できない', () => {
  const { sandbox, ss } = load({ activeEmail: 'yoso@other.example' });
  setUp(sandbox, ss);
  const res = call(sandbox, 'mbSaveJournal', { content: 'よその子', emotion: '😊' });
  assert.equal(res.success, false);
  assert.equal(res.code, 'NOT_MEMBER');
});

test('承認待ちの児童は提出できない', () => {
  const { sandbox, ss } = load({ activeEmail: 'aoi@school.example' });
  setUp(sandbox, ss);
  ss.getSheetByName('児童名簿').appendRow(['児童', 'あおい', 'aoi@school.example', 'pending']);
  const res = call(sandbox, 'mbSaveJournal', { content: 'まだ承認前', emotion: '😊' });
  assert.equal(res.success, false);
  assert.equal(res.code, 'MEMBER_PENDING');
});

test('参加申請は、既定で承認待ちにしかならない', () => {
  const { sandbox, ss } = load({ activeEmail: 'aoi@school.example' });
  setUp(sandbox, ss);
  const res = call(sandbox, 'mbRequestJoin', 'あおい');
  assert.equal(res.success, true, res.error);
  assert.equal(res.status, 'pending');
});

test('配ったばかりのファイルでも、児童は先生になれない', () => {
  // _meta は空。以前はここで「押すまで待つ」形にしていたが、いまは
  // resolveOwner_ がデプロイした本人を控えるので、待たせずに、かつ児童は入れない。
  const { sandbox, ss } = load({
    activeEmail: 'yuto@school.example',            // 開いたのは児童
    effectiveEmail: 'sensei@school.example'        // 動かしているのは先生
  });
  const res = call(sandbox, 'opGetClassData');
  assert.equal(res.success, false);
  assert.equal(res.code, 'FORBIDDEN', res.error);
  // 控えられたのは先生であって、呼んだ児童ではない
  const meta = ss.getSheetByName('_meta')._grid.find((r) => String(r[0]) === 'ownerEmail');
  assert.equal(String(meta[1]), 'sensei@school.example');
});

test('実行ユーザーの設定ミスは、API でも止まる（画面だけの防御にしない）', () => {
  // 画面が出ない状態でも google.script.run は呼べる。doGet だけで見ていると素通りする。
  const { sandbox, ss } = load({
    activeEmail: 'yuto@school.example',
    effectiveEmail: 'yuto@school.example'          // 児童の権限で動いている
  });
  setUp(sandbox, ss);
  ss.getSheetByName('児童名簿').appendRow(['児童', 'ゆうと', 'yuto@school.example', 'active']);
  for (const fn of ['opGetClassData', 'opResetData', 'opGetSetupStatus', 'mbGetStatus', 'mbSync']) {
    const res = call(sandbox, fn);
    assert.equal(res.success, false, `${fn} が設定ミスのまま通っています`);
    assert.equal(res.code, 'OWNER_MISMATCH', `${fn}: ${res.error}`);
  }
});

test('児童は先生の API を呼べない', () => {
  const { sandbox, ss } = load({ activeEmail: 'yuto@school.example' });
  setUp(sandbox, ss);
  ss.getSheetByName('児童名簿').appendRow(['児童', 'ゆうと', 'yuto@school.example', 'active']);
  for (const fn of ['opGetClassData', 'opBatchReturnAll', 'opResetData', 'opInspectSheets', 'opRepairSheets']) {
    const res = call(sandbox, fn);
    assert.equal(res.success, false, `${fn} が児童に通っています`);
    assert.equal(res.code, 'FORBIDDEN', `${fn}: ${res.error}`);
  }
});

test('児童は他人の行に「過去の自分へのメッセージ」を書けない', () => {
  const { sandbox, ss } = load({ activeEmail: 'yuto@school.example' });
  setUp(sandbox, ss);
  const roster = ss.getSheetByName('児童名簿');
  roster.appendRow(['児童', 'ゆうと', 'yuto@school.example', 'active']);
  roster.appendRow(['児童', 'あおい', 'aoi@school.example', 'active']);
  const j = ss.getSheetByName('ジャーナルデータ');
  j.appendRow(['jjjjjjjj-1111', new Date(), 'aoi@school.example', 'テーマ', 'あおいの本文',
               '', '😊', '', '[]', '', '未返却', '', '']);

  const res = call(sandbox, 'mbAddPastComment', 'jjjjjjjj-1111', 'よこどり');
  assert.equal(res.success, false);
  assert.equal(res.code, 'FORBIDDEN');
  assert.equal(j._grid[1][11], '', '他人の行が書き換えられています');
});

test('児童へ返す一覧に、ほかの児童のメールアドレスが出ない', () => {
  const { sandbox, ss } = load({ activeEmail: 'yuto@school.example' });
  setUp(sandbox, ss);
  ss.getSheetByName('児童名簿').appendRow(['児童', 'ゆうと', 'yuto@school.example', 'active']);
  ss.getSheetByName('ジャーナルデータ').appendRow(
    ['jjjjjjjj-2222', new Date(), 'yuto@school.example', 'テーマ', '本文',
     '', '😊', '', '[]', '', '未返却', '', '']);
  const res = call(sandbox, 'mbSync');
  assert.equal(res.success, true, res.error);
  assert.equal(res.journals.length, 1);
  assert.doesNotMatch(JSON.stringify(res.journals), /@school\.example/);
  assert.match(res.journals[0].email, /^u[0-9a-f]+$/);
});

test('先生用メニューは、画面が無い文脈（ウェブアプリ）では1セルも書かない', () => {
  const { sandbox, ss } = load({ activeEmail: 'yuto@school.example' });
  const before = JSON.stringify(ss._sheets.map((s) => s._grid));
  for (const fn of ['showSetupStatus', 'showSheetCheck', 'showSheetRepair']) {
    assert.throws(() => vm.runInContext(`${fn}()`, sandbox), /画面がありません/, fn);
  }
  assert.equal(JSON.stringify(ss._sheets.map((s) => s._grid)), before);
});

// ════════════════════════════════════════════════════════════════
// 5. 表計算ソフトに数式として食わせない
// ════════════════════════════════════════════════════════════════

test('先頭が = + - @ の入力は、CSV で数式にならないよう無害化する', () => {
  const { sandbox } = load();
  for (const [input, expected] of [
    ['=IMPORTXML("http://evil.example")', "'=IMPORTXML(\"http://evil.example\")"],
    ['+1+1', "'+1+1"],
    ['-5', "'-5"],
    ['@here', "'@here"],
    ['ふつうの文', 'ふつうの文']
  ]) {
    assert.equal(vm.runInContext(`csvSafe_(${JSON.stringify(input)})`, sandbox), expected);
  }
});

// ════════════════════════════════════════════════════════════════
// 6. doGet — ここが落ちると、画面が 1 つも開かない
// ════════════════════════════════════════════════════════════════

test('差し込みは、ファイルの中身をそのまま返す（読み直さない）', () => {
  const { sandbox } = load();
  for (const name of ['app', 'vendor', 'css', 'qr']) {
    const got = vm.runInContext(`include_(${JSON.stringify(name)})`, sandbox);
    const want = fs.readFileSync(path.join(ROOT, name + '.html'), 'utf8');
    assert.equal(got, want, `${name}.html が 1 バイトでも変わって届いている`);
  }
});

test('差し込んだ app.html が、そのままで JavaScript として読める', () => {
  // ブラウザが受け取るのは <script> 1 個ぶんの中身である。
  // 2026-08-24 の事故では、ここが壊れた状態で届き、
  // 1315 行目（QR の保存名）を構文エラーとして指された。
  const { sandbox } = load();
  const html = vm.runInContext("include_('app')", sandbox);
  const open = html.indexOf('>', html.indexOf('<script'));
  const close = html.lastIndexOf('</script');
  assert.ok(open > 0 && close > open, 'app.html が <script> 1 個の形になっていない');
  const js = html.slice(open + 1, close);
  assert.doesNotThrow(() => new Function(js), 'app.html の中の JavaScript が読めない');
});

test('中身を読み直す差し込み方（createHtmlOutputFromFile）は使えない', () => {
  const { sandbox } = load();
  assert.throws(
    () => vm.runInContext("HtmlService.createHtmlOutputFromFile('app').getContent()", sandbox),
    /createHtmlOutputFromFile は使わないこと/);
});

test('doGet が落ちずに画面を返す（先生）', () => {
  const { sandbox, ss, rendered } = load({ activeEmail: 'sensei@school.example' });
  setUp(sandbox, ss);
  vm.runInContext('doGet({ parameter: {} })', sandbox);

  assert.equal(rendered.title, 'ふりかえりジャーナル');
  // 実在しない列挙子を書くと undefined が渡り、GAS は
  // 「引数は null にできません: mode」で落ちる。名前まで見る。
  assert.equal(rendered.xFrameOptionsMode, 'DEFAULT');
  assert.match(rendered.metaTags.viewport, /viewport-fit=cover/);
  assert.doesNotMatch(rendered.metaTags.viewport, /user-scalable\s*=\s*no/);

  const boot = JSON.parse(rendered.template.bootJson);
  assert.equal(boot.mode, 'owner');
  assert.equal(boot.signedIn, true);
  assert.equal(boot.className, '3年2組');
  assert.equal(rendered.template.bootMode, 'owner');
});

test('doGet が落ちずに画面を返す（児童）', () => {
  const { sandbox, ss, rendered } = load({ activeEmail: 'yuto@school.example' });
  setUp(sandbox, ss);
  ss.getSheetByName('児童名簿').appendRow(['児童', 'ゆうと', 'yuto@school.example', 'active']);
  vm.runInContext('doGet({ parameter: {} })', sandbox);

  const boot = JSON.parse(rendered.template.bootJson);
  assert.equal(boot.mode, 'member');
  assert.equal(rendered.template.bootMode, 'member');
});

test('ログインが確かめられなくても doGet は画面を返す（理由を渡す）', () => {
  const { sandbox, ss, rendered } = load({ activeEmail: '' });
  setUp(sandbox, ss);
  vm.runInContext('doGet({ parameter: {} })', sandbox);
  const boot = JSON.parse(rendered.template.bootJson);
  assert.equal(boot.signedIn, false);
  assert.equal(boot.mode, 'member');
});

test('束ねられたファイルが無くても doGet は落ちず、理由を画面へ渡す', () => {
  const { sandbox, rendered } = load();
  // 独立スクリプトとして貼られた状態
  vm.runInContext('SpreadsheetApp.getActiveSpreadsheet = () => null;', sandbox);
  vm.runInContext('doGet({ parameter: {} })', sandbox);
  const boot = JSON.parse(rendered.template.bootJson);
  assert.match(boot.bootError, /束ねられていません/);
});

// ════════════════════════════════════════════════════════════════
// 7. 先生を「デプロイした本人」から導く（設定作業を無くしたぶん、ここが要）
// ════════════════════════════════════════════════════════════════

const metaOf = (ss, key) => {
  const grid = ss.getSheetByName('_meta')._grid;
  const row = grid.find((r) => String(r[0]) === key);
  return row ? String(row[1]) : '';
};

test('初回は、デプロイした本人が自動で先生になる（押すものは無い）', () => {
  // setUp を呼ばない ＝ _meta が空。配ったばかりのコピーと同じ状態
  const { sandbox, ss, rendered } = load({ activeEmail: 'sensei@school.example' });
  vm.runInContext('doGet({ parameter: {} })', sandbox);

  const boot = JSON.parse(rendered.template.bootJson);
  assert.equal(boot.mode, 'owner');
  assert.equal(metaOf(ss, 'ownerEmail'), 'sensei@school.example');
  assert.equal(metaOf(ss, 'fileId'), 'ss-test', 'コピーを見分けるための ID が控えられていない');
});

test('児童が先生より先に開いても、先生になるのはデプロイした人', () => {
  // これが getActiveUser を控える形との決定的な違い。
  // 旧構成では、先生より先に開いた児童が恒久的に先生になった。
  const { sandbox, ss, rendered } = load({
    activeEmail: 'yuto@school.example',            // 開いたのは児童
    effectiveEmail: 'sensei@school.example'        // 動かしているのは先生（実行:自分）
  });
  vm.runInContext('doGet({ parameter: {} })', sandbox);

  assert.equal(metaOf(ss, 'ownerEmail'), 'sensei@school.example', '児童が先生になってしまった');
  assert.equal(JSON.parse(rendered.template.bootJson).mode, 'member');
});

test('クラス名は、控えが無ければスプレッドシートの名前を使う（「のコピー」は落とす）', () => {
  const { sandbox, ss, rendered } = load();
  ss.getName = () => 'ふりかえりジャーナル のコピー';
  vm.runInContext('doGet({ parameter: {} })', sandbox);
  assert.equal(JSON.parse(rendered.template.bootJson).className, 'ふりかえりジャーナル');
});

test('コピーされたら、前の持ち主の控えを自動で貼り替える', () => {
  // 配布テンプレートに前の持ち主の _meta が残っていても、コピーした先生が入れる。
  const { sandbox, ss, rendered } = load({
    activeEmail: 'newteacher@school.example',
    effectiveEmail: 'newteacher@school.example'
  });
  const meta = ss.getSheetByName('_meta');
  meta.appendRow(['ownerEmail', 'oldteacher@school.example']);
  meta.appendRow(['fileId', 'the-template-file']);      // コピー元の ID
  vm.runInContext('doGet({ parameter: {} })', sandbox);

  const boot = JSON.parse(rendered.template.bootJson);
  assert.equal(boot.mode, 'owner', 'コピーした先生が先生になれていない');
  assert.equal(boot.copied, true);
  assert.equal(metaOf(ss, 'ownerEmail'), 'newteacher@school.example');
  assert.equal(metaOf(ss, 'fileId'), 'ss-test');
  assert.ok(metaOf(ss, 'copiedAt'), 'コピーの目印が残っていない');
});

test('実行ユーザーが「自分」でないと、画面を出さずに理由を出す', () => {
  // 「実行するユーザー: アプリケーションにアクセスしているユーザー」にすると、
  // getEffectiveUser が開いた本人を返す。その形では全員が自分を先生と名乗れる。
  const { sandbox, ss, rendered } = load({
    activeEmail: 'yuto@school.example',
    effectiveEmail: 'yuto@school.example'      // 児童の権限で動いている
  });
  setUp(sandbox, ss);                          // 先生は sensei で控えられている。ファイルは同じ
  vm.runInContext('doGet({ parameter: {} })', sandbox);

  const boot = JSON.parse(rendered.template.bootJson);
  assert.match(boot.bootError, /実行するユーザー/);
  assert.notEqual(boot.mode, 'owner', '設定ミスのまま先生画面を出している');
  assert.equal(metaOf(ss, 'ownerEmail'), 'sensei@school.example', '先生が書き換えられている');
});

test('動かしているアカウントが取れなくても、控えがあれば先生は変わらない', () => {
  const { sandbox, ss, rendered } = load({ activeEmail: 'sensei@school.example', effectiveEmail: '' });
  setUp(sandbox, ss);
  vm.runInContext('doGet({ parameter: {} })', sandbox);
  assert.equal(JSON.parse(rendered.template.bootJson).mode, 'owner');
  assert.equal(metaOf(ss, 'ownerEmail'), 'sensei@school.example');
});

// ════════════════════════════════════════════════════════════════
// 8. 準備の状態と、コピー元から引き継いだ記録
// ════════════════════════════════════════════════════════════════

test('準備の状態は、済んでいないものを名指しする', () => {
  const { sandbox, ss } = load();
  setUp(sandbox, ss);
  const res = call(sandbox, 'opGetSetupStatus');
  assert.equal(res.success, true, res.error);
  const keys = res.status.items.map((it) => it.key);
  for (const want of ['owner', 'className', 'executeAs', 'access', 'sheets', 'gemini']) {
    assert.ok(keys.includes(want), want + ' が準備の状態に出ていない');
  }
  assert.equal(res.status.ok, true);
  assert.equal(res.status.ownerEmail, 'sensei@school.example');
});

test('準備の状態は、児童からは読めない', () => {
  const { sandbox, ss } = load({ activeEmail: 'yuto@school.example' });
  setUp(sandbox, ss);
  ss.getSheetByName('児童名簿').appendRow(['児童', 'ゆうと', 'yuto@school.example', 'active']);
  const res = call(sandbox, 'opGetSetupStatus');
  assert.equal(res.success, false);
  assert.equal(res.code, 'FORBIDDEN');
});

test('コピーで引き継いだ記録は、数えるだけで勝手に消さない', () => {
  const { sandbox, ss, rendered } = load({
    activeEmail: 'newteacher@school.example', effectiveEmail: 'newteacher@school.example'
  });
  const meta = ss.getSheetByName('_meta');
  meta.appendRow(['ownerEmail', 'oldteacher@school.example']);
  meta.appendRow(['fileId', 'the-template-file']);
  ss.getSheetByName('児童名簿').appendRow(['児童', 'まえのこ', 'old@school.example', 'active']);
  ss.getSheetByName('ジャーナルデータ').appendRow(['j1', new Date(), 'old@school.example', 'テーマ', '本文']);

  vm.runInContext('doGet({ parameter: {} })', sandbox);
  assert.equal(JSON.parse(rendered.template.bootJson).copied, true);

  // doGet を通しただけでは 1 行も消えていない
  assert.equal(ss.getSheetByName('ジャーナルデータ')._grid.length - 1, 1, '勝手に消している');

  const st = call(sandbox, 'opGetSetupStatus');
  const inherited = st.status.items.find((it) => it.key === 'inherited');
  assert.ok(inherited, '引き継ぎが準備の状態に出ていない');
  assert.equal(inherited.journals, 1);

  // 先生が押したときだけ消える
  const res = call(sandbox, 'opClearInheritedData');
  assert.equal(res.success, true, res.error);
  assert.equal(ss.getSheetByName('ジャーナルデータ')._grid.length - 1, 0);
  assert.equal(ss.getSheetByName('児童名簿')._grid.length - 1, 0);
  // 見出し行は残っている
  assert.equal(String(ss.getSheetByName('ジャーナルデータ')._grid[0][0]), 'journalId');
});

test('引き継いだ記録の掃除は、児童からは呼べない', () => {
  const { sandbox, ss } = load({ activeEmail: 'yuto@school.example' });
  setUp(sandbox, ss);
  ss.getSheetByName('児童名簿').appendRow(['児童', 'ゆうと', 'yuto@school.example', 'active']);
  const before = ss.getSheetByName('児童名簿')._grid.length;
  const res = call(sandbox, 'opClearInheritedData');
  assert.equal(res.success, false);
  assert.equal(res.code, 'FORBIDDEN');
  assert.equal(ss.getSheetByName('児童名簿')._grid.length, before, '児童の呼び出しで消えている');
});

test('boot の JSON は < を潰してある（クラス名で画面が切れない）', () => {
  const { sandbox, ss, rendered } = load();
  setUp(sandbox, ss);
  // 先生が「</script>」を含む名前を付けても、テンプレートが途中で切れない
  vm.runInContext(`upsertKeyValue_(getDb_().getSheetByName('_meta'), 'className', '3年2組</script><b>');`, sandbox);
  vm.runInContext('doGet({ parameter: {} })', sandbox);
  assert.doesNotMatch(rendered.template.bootJson, /<\/script/i);
  assert.equal(JSON.parse(rendered.template.bootJson).className, '3年2組</script><b>');
});
