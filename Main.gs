/**
 * ============================================================
 * ふりかえりジャーナル — Main.gs
 * 設定・スプレッドシートの取得・メニュー・doGet ルーティング・共通ユーティリティ
 * ============================================================
 *
 * ── 配り方（コンテナバインド・先生ごとにデプロイ） ──────────────────
 * スプレッドシートのコピーを配り、そのファイルにこのスクリプトが束ねられている。
 * 1 ファイル = 1 クラスで、先生ご自身の Google ドライブに置かれる。
 * 中央のレジストリもクラスコードも運営者のアカウントも無い。
 *
 *   先生: コピー → メニュー「ふりかえりジャーナル」＞「はじめの設定」→ デプロイ → URL を配る
 *   児童: 配られた URL を開く。学校の Google アカウントでログインしていれば、それが本人確認になる
 *
 * ── 本人確認 ────────────────────────────────────────────────────
 * デプロイは「実行するユーザー: 自分」「アクセスできるユーザー: 同じ組織内の全員」を前提にする。
 * この形なら、児童が開いたときに Session.getActiveUser().getEmail() が本人のアドレスを返す。
 * 「全員（匿名ユーザーを含む）」を選ぶと空になるので、そのときは画面で理由を出して止める
 * （名前を騙れる状態で学習記録を書かせない）。
 *
 * ── 誰が先生か ──────────────────────────────────────────────────
 * **ウェブアプリからは、誰も先生になれない。** 先生の登録は、スプレッドシートを開ける人
 * （＝ファイルの持ち主）がメニューから 1 回行う（setupAsTeacher）。児童にはスプレッドシート
 * 自体を渡さないので、この経路には入れない。
 * 「未設定のときは最初に開いた人を先生にする」形にすると、先生が開く前に URL を配った学級で
 * 最初の児童が恒久的に先生になる。その事故は取り消しがきかないので、初回でも通さない。
 */

const CONFIG = {
  APP_NAME: 'ふりかえりジャーナル',
  SCHEMA_VERSION: 3,
  LOCK_TIMEOUT_MS: 10000
};

/** _meta シートに置く鍵。ScriptProperties は使わない（コピーには付いてこないため） */
const META_KEYS = {
  OWNER_EMAIL: 'ownerEmail',      // このクラスの先生
  CLASS_NAME: 'className',
  SCHEMA: 'schemaVersion',
  SETUP_AT: 'setupAt'
};

// ────────────────────────────────────────────────────────────────
// スプレッドシート（このスクリプトが束ねられているファイル）
// ────────────────────────────────────────────────────────────────

/**
 * このクラスの記録が入っているスプレッドシート。
 *
 * コンテナバインドなので、開いているそのファイルがそのまま本体である。
 * ID の設定も、自動生成も、探しに行く処理も要らない。
 *
 * ⚠️ 独立スクリプトとして貼り付けた場合はここが null になる。そのときは
 *    「作り直して探しに行く」ことをしない。児童一人ひとりの権限で走る構成に
 *    なっていると、権限の無い子が 1 回開くだけで、学級全員の記録が入ったファイルから
 *    その子の空ファイルへ静かに差し替わる。エラーは出ず「記録が消えた」ようにしか見えない。
 */
function getDb_() {
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { ss = null; }
  if (!ss) {
    throw new Error('NOT_BOUND: このスクリプトはスプレッドシートに束ねられていません。' +
      '配布用のスプレッドシートをコピーして、そのコピーの「拡張機能 ＞ Apps Script」から開いてください');
  }
  return ensureSheets_(ss);
}

// ────────────────────────────────────────────────────────────────
// 本人確認
// ────────────────────────────────────────────────────────────────

/**
 * いま画面を開いている人のメールアドレス。取れなければ空文字。
 *
 * 取れないのは、デプロイの「アクセスできるユーザー」が「全員（匿名ユーザーを含む）」に
 * なっているか、児童が学校とは別のドメインのアカウントで開いている場合。
 */
function activeEmail_() {
  try { return String(Session.getActiveUser().getEmail() || '').trim().toLowerCase(); }
  catch (e) { return ''; }
}

/** 本人が取れなければ、次に何をすればよいかまで書いて止める */
function requireEmail_() {
  const email = activeEmail_();
  if (!email) {
    throw new Error('NO_IDENTITY: 誰が使っているかを確かめられませんでした。' +
      '学校の Google アカウントでログインしてから開いてください。' +
      '（先生へ: デプロイの「アクセスできるユーザー」を「同じ組織内の全員」にしてください。' +
      '「全員（匿名ユーザーを含む）」だと、誰が書いたかを確かめられません）');
  }
  return email;
}

/** このクラスの先生のアドレス（_meta）。未設定なら空文字 */
function ownerEmailOf_(ss) {
  return String(readMeta_(ss, META_KEYS.OWNER_EMAIL) || '').trim().toLowerCase();
}

/**
 * 先生かどうか。_meta の先生本人か、名簿で役割が「担任」かつ状態が active の人。
 * どちらも、スプレッドシートを開ける人しか書き込めない場所である。
 */
function isOwnerEmail_(ss, email) {
  if (!email) return false;
  const owner = ownerEmailOf_(ss);
  if (owner && owner === email) return true;
  const m = getMemberRow_(ss, email);
  return !!(m && m.role === '担任' && m.status === 'active');
}

// ────────────────────────────────────────────────────────────────
// スプレッドシートのメニュー（先生だけが通れる経路）
// ────────────────────────────────────────────────────────────────

/**
 * スプレッドシートを開いたときのメニュー。
 * ウェブアプリとして動いているときは画面が無いので、何も起きない。
 */
function onOpen(e) {
  try {
    SpreadsheetApp.getUi()
      .createMenu(CONFIG.APP_NAME)
      .addItem('はじめの設定（最初に1回）', 'setupAsTeacher')
      .addSeparator()
      .addItem('シートを点検する', 'showSheetCheck')
      .addItem('シートを直せる範囲で直す', 'showSheetRepair')
      .addToUi();
  } catch (err) {
    // ウェブアプリ文脈では画面が無い。ここで止まってよい。
  }
}

/**
 * 「はじめの設定」。シートを作り、押した人をこのクラスの先生として控える。
 *
 * ⚠️ google.script.run は末尾 `_` の無い関数を誰でも呼べるので、この関数も児童から
 *    呼べてしまう。**先に getUi() を取る**のはそのため。画面が無い文脈（ウェブアプリ）では
 *    ここで例外になり、1 セルも書かずに終わる。
 *    加えて、すでに先生が決まっている場合は、たとえ画面があっても上書きしない。
 */
function setupAsTeacher() {
  const ui = SpreadsheetApp.getUi();          // 画面が無ければ、ここで止まる
  const ss = getDb_();
  const me = activeEmail_();
  if (!me) {
    ui.alert(CONFIG.APP_NAME, 'Google アカウントを確かめられませんでした。'
      + 'スプレッドシートを開き直してからもう一度お試しください。', ui.ButtonSet.OK);
    return;
  }

  const current = ownerEmailOf_(ss);
  if (current && current !== me) {
    ui.alert(CONFIG.APP_NAME,
      'このクラスの先生はすでに ' + current + ' で登録されています。\n\n'
      + '付け替えるときは「_meta」シートの ' + META_KEYS.OWNER_EMAIL + ' の行を書き直してください。',
      ui.ButtonSet.OK);
    return;
  }

  const answer = ui.prompt(CONFIG.APP_NAME,
    'クラスの名前を入れてください（例: 3年2組）。\nあとから「_meta」シートで変えられます。',
    ui.ButtonSet.OK_CANCEL);
  if (answer.getSelectedButton() !== ui.Button.OK) return;
  const className = String(answer.getResponseText() || '').trim().slice(0, 50) || 'わたしたちのクラス';

  writeMeta_(ss, {
    ownerEmail: me,
    className: className,
    schemaVersion: CONFIG.SCHEMA_VERSION,
    setupAt: new Date().toISOString()
  });
  upsertMember_(ss, { email: me, name: '先生', role: '担任', status: 'active' });

  ui.alert(CONFIG.APP_NAME,
    '「' + className + '」として設定しました。\n\n'
    + '次に「拡張機能 ＞ Apps Script」を開き、右上の「デプロイ ＞ 新しいデプロイ」から\n'
    + '　種類: ウェブアプリ\n'
    + '　実行するユーザー: 自分\n'
    + '　アクセスできるユーザー: 同じ組織内の全員\n'
    + 'で公開し、出てきた URL を児童に配ってください。',
    ui.ButtonSet.OK);
}

/** 点検の結果を見せるだけ。1 セルも書き換えない */
function showSheetCheck() {
  const ui = SpreadsheetApp.getUi();          // 画面が無ければ、ここで止まる
  const found = inspectSheets_(getDb_());
  ui.alert('シートの点検', sheetReportText_(found), ui.ButtonSet.OK);
}

/** 直せる範囲だけ直す。何をするかを先に見せて、押してもらってから動く */
function showSheetRepair() {
  const ui = SpreadsheetApp.getUi();          // 画面が無ければ、ここで止まる
  const ss = getDb_();
  const found = inspectSheets_(ss);
  const fixable = found.filter(function (f) { return f.fixable; });

  if (!fixable.length) {
    ui.alert('シートを直す',
      found.length
        ? '自動で直せるものはありませんでした。\n\n' + sheetReportText_(found)
        : 'シートの作りは想定どおりです。直すところはありません。',
      ui.ButtonSet.OK);
    return;
  }

  const plan = fixable.map(function (f) {
    return '・「' + f.sheet + '」' + f.kind + '\n　→ ' + f.action;
  }).join('\n');
  const answer = ui.alert('シートを直す',
    '次のことをします。既にある列は動かさず、名前も変えません。消すものはありません。\n\n'
    + plan + '\n\n実行してよろしいですか。', ui.ButtonSet.OK_CANCEL);
  if (answer !== ui.Button.OK) return;

  const result = repairSheets_(ss);
  let text = result.fixed.length
    ? '直しました。\n\n' + result.fixed.map(function (s) { return '・' + s; }).join('\n')
    : '直せるものはありませんでした。';
  if (result.left.length) {
    text += '\n\n── 人が見ないと直せないもの ──\n'
      + result.left.map(function (f) {
          return '・「' + f.sheet + '」' + f.kind + '：' + f.detail + '\n　→ ' + f.action;
        }).join('\n');
  }
  ui.alert('シートを直す', text, ui.ButtonSet.OK);
}

/** 点検の所見を、そのまま読める日本語にする（メニューと先生画面の両方で使う） */
function sheetReportText_(found) {
  if (!found.length) return 'シートの作りは想定どおりです。';
  return '次のところが、アプリの想定と違います。\n\n'
    + found.map(function (f) {
        return '・「' + f.sheet + '」' + f.kind + '：' + f.detail
          + '\n　→ ' + f.action + (f.fixable ? '（自動で直せます）' : '（自動では直しません）');
      }).join('\n')
    + '\n\n列は見出しの名前で探しているので、並べ替えは害になりません。'
    + '見出しの名前を変えたり列を消したりすると、その列の読み書きが止まります。';
}

// ────────────────────────────────────────────────────────────────
// 共通ユーティリティ
// ────────────────────────────────────────────────────────────────

function sha256Hex_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(function (b) { return ((b + 256) % 256).toString(16).padStart(2, '0'); })
    .join('');
}

/** 書き込みの直列化。保持区間は最小に（appendRow 数回程度）保つこと */
function withScriptLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(CONFIG.LOCK_TIMEOUT_MS)) {
    throw new Error('LOCK_BUSY: 処理が混み合っています。数秒待ってもう一度お試しください');
  }
  try { return fn(); } finally { try { lock.releaseLock(); } catch (e) {} }
}

/** API レスポンスは必ずこの 2 つで返す。エラーは "CODE: 日本語メッセージ" 形式で throw する */
function jsonOk_(obj) {
  const out = obj || {};
  out.success = true;
  return JSON.stringify(out);
}

function jsonErr_(e) {
  const msg = (e && e.message) ? e.message : String(e);
  const m = msg.match(/^([A-Z_]+):\s*(.*)$/);
  return JSON.stringify({
    success: false,
    code: m ? m[1] : '',
    error: m ? (m[2] || msg) : msg
  });
}

// ────────────────────────────────────────────────────────────────
// ルーティング
// ────────────────────────────────────────────────────────────────

/**
 * 入口は 1 つ。誰が開いたかで先生画面と児童画面を切り替える。
 *
 * URL パラメータでは切り替えない。?portal=1 のような形にすると、児童が
 * その 1 文字を足すだけで先生画面の見た目に入れてしまう（サーバー側の認可は
 * 別に効くが、押せないボタンが並ぶ画面を児童に見せる意味は無い）。
 */
function doGet(e) {
  let boot;
  try {
    const ss = getDb_();
    const email = activeEmail_();
    const owner = ownerEmailOf_(ss);
    boot = {
      mode: isOwnerEmail_(ss, email) ? 'owner' : 'member',
      className: String(readMeta_(ss, META_KEYS.CLASS_NAME) || CONFIG.APP_NAME),
      // 空なら「ログインが確かめられない」ことが確定する。画面で理由を出す
      signedIn: !!email,
      // 先生がまだ「はじめの設定」を押していない
      setupDone: !!owner,
      webAppUrl: ScriptApp.getService().getUrl() || ''
    };
  } catch (err) {
    boot = {
      mode: 'member', className: CONFIG.APP_NAME, signedIn: false, setupDone: false,
      webAppUrl: '', bootError: (err && err.message) ? err.message : String(err)
    };
  }

  const t = HtmlService.createTemplateFromFile('index');
  // <?!= ?> は素通しなので、閉じタグに化ける < を必ず潰してから渡す
  t.bootJson = JSON.stringify(boot).replace(/</g, '\\u003c');
  t.bootMode = boot.mode;
  return t.evaluate()
    .setTitle(CONFIG.APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.SAMEORIGIN)
    // ⚠️ viewport は index.html の <meta> にもある。addMetaTag はサーバー側の処理なので、
    //    index.html だけ直しても反映されない。必ず両方を同じ値にすること。
    //    拡大は禁止しない（誤ズーム防止より、見えづらい子が拡大できない害のほうが大きい）。
    //    viewport-fit=cover は GAS が画面を iframe で包むため、両方に要る。
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, viewport-fit=cover');
}

/**
 * HTML ファイルを本体へ差し込む。GAS には .gs と .html しか置けないため、
 * CSS も JavaScript もライブラリも .html に包んで持つ。
 *
 * vendor / css / app は npm run build が作る生成物。手で編集しないこと
 * （原本は src/app.jsx・tools/extra.css・tailwind.config.js）。
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
