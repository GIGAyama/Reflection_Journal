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
 *   先生: コピー → デプロイ → URL を配る（スプレッドシート側で押すものは 1 つも無い）
 *   児童: 配られた URL を開く。学校の Google アカウントでログインしていれば、それが本人確認になる
 *
 * ── 本人確認 ────────────────────────────────────────────────────
 * デプロイは「実行するユーザー: 自分」「アクセスできるユーザー: 同じ組織内の全員」を前提にする。
 * この形なら、児童が開いたときに Session.getActiveUser().getEmail() が本人のアドレスを返す。
 * 「全員（匿名ユーザーを含む）」を選ぶと空になるので、そのときは画面で理由を出して止める
 * （名前を騙れる状態で学習記録を書かせない）。
 *
 * ── 誰が先生か（登録しない。デプロイした本人がそのまま先生） ──────────
 * executeAs は USER_DEPLOYING なので、**誰が開いても** Session.getEffectiveUser() は
 * 「デプロイした人」を返す。児童が開いたときも返るのは児童ではなく先生である。
 * だから先生は「登録する」ものではなく、**デプロイの時点で決まっている**。
 * resolveOwner_() は、_meta が空なら getEffectiveUser() をそのまま控える。
 *
 * 「最初に開いた人を先生にする」形（＝ getActiveUser を控える形）は採らない。それだと
 * 先生より先に URL を開いた児童が恒久的に先生になる。getEffectiveUser はその競争が
 * そもそも起きない。**誰が最初に開いても、控えられるのは同じ 1 人**だからである。
 *
 * ── 実行ユーザーの設定ミスを、ここで止める ────────────────────────
 * デプロイを「実行するユーザー: アプリケーションにアクセスしているユーザー」にすると、
 * getEffectiveUser() は開いた本人を返すようになる。その形では全員が自分を先生として
 * 名乗れてしまう。控えてある先生と effective が食い違ったら、それが起きた証拠なので、
 * 画面を出さずに理由を出して止める（OWNER_MISMATCH）。
 *
 * ── コピーを見分ける ────────────────────────────────────────────
 * 配布用テンプレートに前の持ち主の _meta が残っていると、コピーした先生が
 * OWNER_MISMATCH で入れなくなる。そこで控えるときに fileId も一緒に書いておき、
 * **いま開いているファイルの ID と違えば「これはコピーだ」と判る**。
 * コピーと判ったら先生とクラス名を貼り替える。記録は消さない（消すのは先生が押したときだけ）。
 */

const CONFIG = {
  APP_NAME: 'ふりかえりジャーナル',
  SCHEMA_VERSION: 3,
  LOCK_TIMEOUT_MS: 10000
};

/** _meta シートに置く鍵。ScriptProperties は使わない（コピーには付いてこないため） */
const META_KEYS = {
  OWNER_EMAIL: 'ownerEmail',      // このクラスの先生（デプロイした本人。自動で入る）
  CLASS_NAME: 'className',
  SCHEMA: 'schemaVersion',
  SETUP_AT: 'setupAt',
  FILE_ID: 'fileId',              // 控えた時点のファイル ID。コピーを見分けるために使う
  COPIED_AT: 'copiedAt'           // コピーと判って貼り替えた時刻（引き継ぎの掃除の目印）
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

/**
 * このスクリプトを動かしているアカウント。
 *
 * executeAs = USER_DEPLOYING なので、**誰が開いてもデプロイした人**が返る。
 * 児童が開いたときも返るのは児童ではない。だからこれが「先生は誰か」の根拠になる。
 */
function effectiveEmail_() {
  try { return String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase(); }
  catch (e) { return ''; }
}

/** _meta に控えてある先生のアドレス。まだ何も控えていなければ空文字 */
function storedOwnerOf_(ss) {
  return String(readMeta_(ss, META_KEYS.OWNER_EMAIL) || '').trim().toLowerCase();
}

/**
 * このクラスの先生を決める。**設定作業は無い。**
 *
 * 1. まだ何も控えていない → いま動かしているアカウントを控える（初回）
 * 2. 控えたときと**ファイル ID が違う** → これはコピー。先生を貼り替える
 * 3. ファイルは同じなのに動かしているアカウントが違う → デプロイの実行ユーザーの
 *    設定ミス。ここで止める（黙って通すと全員が先生になれる）
 *
 * 戻り値 { email, className, copied, inheritedFrom }。
 * 読み取りだけの経路からも呼ばれるので、書き込みは「必要になったときだけ」行う。
 */
function resolveOwner_(ss) {
  const effective = effectiveEmail_();
  const stored = storedOwnerOf_(ss);
  const storedFileId = String(readMeta_(ss, META_KEYS.FILE_ID) || '').trim();
  let fileId = '';
  try { fileId = String(ss.getId() || ''); } catch (e) { fileId = ''; }

  // 動かしているアカウントが取れない。ここで控えると空文字が先生になる
  if (!effective) {
    if (stored) return { email: stored, className: classNameOf_(ss), copied: false, inheritedFrom: '' };
    throw new Error('NO_EFFECTIVE_USER: このアプリを動かしているアカウントを確かめられませんでした。' +
      '「拡張機能 ＞ Apps Script ＞ デプロイ」を開き、「実行するユーザー」を「自分」にして デプロイし直してください');
  }

  // 初回。デプロイした本人をそのまま控える
  if (!stored) {
    writeMeta_(ss, {
      ownerEmail: effective,
      className: classNameOf_(ss),
      schemaVersion: CONFIG.SCHEMA_VERSION,
      setupAt: new Date().toISOString(),
      fileId: fileId
    });
    return { email: effective, className: classNameOf_(ss), copied: false, inheritedFrom: '' };
  }

  // コピーされた。前の持ち主の控えが残っているので貼り替える。
  // 記録は消さない（消すのは先生が「準備の状態」で押したときだけ）
  if (storedFileId && fileId && storedFileId !== fileId) {
    writeMeta_(ss, {
      ownerEmail: effective,
      schemaVersion: CONFIG.SCHEMA_VERSION,
      fileId: fileId,
      copiedAt: new Date().toISOString()
    });
    return { email: effective, className: classNameOf_(ss), copied: true, inheritedFrom: stored };
  }

  // 同じファイルなのに動かしているアカウントが違う。実行ユーザーの設定ミスである
  if (stored !== effective) {
    throw new Error('OWNER_MISMATCH: デプロイの「実行するユーザー」が「自分」になっていません。' +
      'いまの設定では、開いた人それぞれの権限で動くので、誰が先生かを決められません。' +
      '「拡張機能 ＞ Apps Script ＞ デプロイ ＞ デプロイを管理」から、' +
      '「実行するユーザー: 自分」「アクセスできるユーザー: 同じ組織内の全員」にしてください');
  }

  // 控えはあるがファイル ID が無い（この仕組みより前に作られたファイル）。いま書き足す
  if (!storedFileId && fileId) writeMeta_(ss, { fileId: fileId });

  return { email: stored, className: classNameOf_(ss), copied: false, inheritedFrom: '' };
}

/**
 * クラス名。控えが無ければ**スプレッドシートの名前をそのまま使う**。
 * 「クラス名を入れてください」と尋ねる画面を無くすため。先生はあとから画面で変えられる。
 */
function classNameOf_(ss) {
  const stored = String(readMeta_(ss, META_KEYS.CLASS_NAME) || '').trim();
  if (stored) return stored;
  let name = '';
  try { name = String(ss.getName() || '').trim(); } catch (e) { name = ''; }
  // 「〜 のコピー」はそのまま出すと恥ずかしいので落とす
  name = name.replace(/\s*のコピー$/, '').trim();
  return name || CONFIG.APP_NAME;
}

/** このクラスの先生のアドレス（_meta）。未設定なら空文字 */
function ownerEmailOf_(ss) {
  return storedOwnerOf_(ss);
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
      .addItem('準備の状態を見る', 'showSetupStatus')
      .addSeparator()
      .addItem('シートを点検する', 'showSheetCheck')
      .addItem('シートを直せる範囲で直す', 'showSheetRepair')
      .addToUi();
  } catch (err) {
    // ウェブアプリ文脈では画面が無い。ここで止まってよい。
  }
}

/**
 * 「準備の状態を見る」。**1 セルも書き換えない**（resolveOwner_ の初回の控えを除く）。
 *
 * スプレッドシート側で押すものはもう無いが、シートで作業する先生のために、
 * 同じ内容をここからも読めるようにしておく。中身はアプリの「準備の状態」と同じ。
 */
function showSetupStatus() {
  const ui = SpreadsheetApp.getUi();          // 画面が無ければ、ここで止まる
  let status;
  try {
    status = buildSetupStatus_(getDb_());
  } catch (err) {
    ui.alert(CONFIG.APP_NAME, (err && err.message) ? err.message : String(err), ui.ButtonSet.OK);
    return;
  }
  ui.alert(CONFIG.APP_NAME, setupStatusText_(status), ui.ButtonSet.OK);
}

/**
 * 「準備の状態」。先生が「何がまだ済んでいないか」を 1 か所で見られるようにする。
 *
 * 設定作業そのものは要らなくなったが、**要らなくなったことが分かる**必要がある。
 * 「押すものが無い」と「押し忘れている」は、先生から見ると区別がつかないためである。
 *
 * level は 'ok' / 'warn' / 'ng'。ng が 1 つでもあれば、児童はまだ使えない。
 */
function buildSetupStatus_(ss) {
  const items = [];
  const owner = resolveOwner_(ss);          // ここで OWNER_MISMATCH なら呼び元へ投げる
  const me = activeEmail_();

  items.push({
    key: 'owner', level: 'ok', title: 'このクラスの先生',
    detail: owner.email + '（デプロイしたアカウントです。設定作業はありません）'
  });

  items.push({
    key: 'className', level: 'ok', title: 'クラスの名前',
    detail: owner.className + '（この画面で変えられます）'
  });

  // 実行ユーザー。ここまで来ている時点で resolveOwner_ を通っているので「自分」で動いている
  items.push({
    key: 'executeAs', level: 'ok', title: 'デプロイの「実行するユーザー」',
    detail: '「自分」で動いています'
  });

  // アクセス範囲。匿名で開けると誰が書いたか確かめられない
  items.push(me
    ? { key: 'access', level: 'ok', title: 'デプロイの「アクセスできるユーザー」',
        detail: 'ログインした人として開けています（' + me + '）' }
    : { key: 'access', level: 'ng', title: 'デプロイの「アクセスできるユーザー」',
        detail: '誰が使っているかを確かめられません。「同じ組織内の全員」にしてください' });

  // シートの作り
  const issues = inspectSheets_(ss);
  items.push(issues.length === 0
    ? { key: 'sheets', level: 'ok', title: 'シートの作り', detail: '6 枚とも想定どおりです' }
    : { key: 'sheets', level: issues.some(function (f) { return !f.fixable; }) ? 'ng' : 'warn',
        title: 'シートの作り',
        detail: issues.length + ' 件、想定と違うところがあります（この画面で直せます）',
        issues: issues });

  // コピー元から引き継いだ記録
  const copiedAt = String(readMeta_(ss, META_KEYS.COPIED_AT) || '').trim();
  const inherited = copiedAt ? countInheritedRows_(ss) : { members: 0, journals: 0 };
  const inheritedTotal = inherited.members + inherited.journals;
  if (copiedAt) {
    items.push(inheritedTotal === 0
      ? { key: 'inherited', level: 'ok', title: 'コピー元から引き継いだ記録', detail: '残っていません' }
      : { key: 'inherited', level: 'warn', title: 'コピー元から引き継いだ記録',
          detail: '名簿 ' + inherited.members + ' 人ぶん、提出 ' + inherited.journals + ' 件が残っています。'
            + '前の学級のものなら、この画面で消せます',
          members: inherited.members, journals: inherited.journals });
  }

  // 承認待ち
  const pending = getMembers_(ss).filter(function (m) { return m.status === 'pending'; }).length;
  if (pending) {
    items.push({ key: 'pending', level: 'warn', title: '参加の承認待ち',
      detail: pending + ' 人が待っています', count: pending });
  }

  // Gemini（任意）
  const key = String(getTenantSetting_(ss, 'GEMINI_API_KEY') || '').trim();
  items.push({ key: 'gemini', level: 'ok', title: 'AI のおへんじ（任意）',
    detail: key ? '使う設定になっています' : '使わない設定です（無くても全部の機能が動きます）' });

  return {
    ok: !items.some(function (it) { return it.level === 'ng'; }),
    ownerEmail: owner.email,
    className: owner.className,
    copiedAt: copiedAt,
    items: items
  };
}

/** 「準備の状態」を、そのまま読める日本語にする */
function setupStatusText_(status) {
  const mark = function (level) {
    return level === 'ok' ? '✅' : (level === 'warn' ? '⚠️' : '❌');
  };
  return status.items.map(function (it) {
    return mark(it.level) + ' ' + it.title + '\n　' + it.detail;
  }).join('\n\n');
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
    // 先生は「登録されている人」ではなく「デプロイした人」。ここで決まる。
    // 実行ユーザーの設定ミス（OWNER_MISMATCH）はここで例外になり、下の catch が理由を出す。
    const owner = resolveOwner_(ss);
    boot = {
      mode: (email && email === owner.email) || isOwnerEmail_(ss, email) ? 'owner' : 'member',
      className: owner.className,
      // 空なら「ログインが確かめられない」ことが確定する。画面で理由を出す
      signedIn: !!email,
      // このファイルが別のファイルのコピーで、前の学級の記録が残っているかもしれない
      copied: !!owner.copied,
      webAppUrl: ScriptApp.getService().getUrl() || ''
    };
  } catch (err) {
    boot = {
      mode: 'member', className: CONFIG.APP_NAME, signedIn: false, copied: false,
      webAppUrl: '', bootError: (err && err.message) ? err.message : String(err)
    };
  }

  const t = HtmlService.createTemplateFromFile('index');
  // <?!= ?> は素通しなので、閉じタグに化ける < を必ず潰してから渡す
  t.bootJson = JSON.stringify(boot).replace(/</g, '\\u003c');
  t.bootMode = boot.mode;
  return t.evaluate()
    .setTitle(CONFIG.APP_NAME)
    // ⚠️ XFrameOptionsMode に SAMEORIGIN は無い。ALLOWALL と DEFAULT の 2 つだけで、
    //    無い名前を書くと undefined が渡り、doGet が
    //    「引数は null にできません: mode」で落ちて画面が一切開かなくなる。
    //    旧構成では GitHub Pages のシェルが iframe で包むため ALLOWALL が要ったが、
    //    いまは自分の URL を直接開くので DEFAULT（＝他サイトからの埋め込みを許さない）でよい。
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
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
 *
 * ⚠️ createHtmlOutputFromFile(...).getContent() を使わないこと。
 *    あれは中身を「HTML として読み直して組み立て直す」。app.html の中身は
 *    <script> 1 個ぶんの JavaScript で、その中には HTML の断片を組み立てる
 *    文字列がある。読み直されると、その断片が本物のタグとして扱われ、
 *    バッククォートの対応が崩れた状態で返ってくる。
 *    getRawContent() はファイルの中身をそのまま返す（解釈を挟まない）。
 *    G11 が見ている。
 */
function include_(filename) {
  return HtmlService.createTemplateFromFile(filename).getRawContent();
}
