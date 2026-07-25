/**
 * MemberApi.gs — デプロイ B（児童アプリ）用 API
 *
 * ★ 鉄則 ★
 *   1. すべての関数の第 1 引数は idToken。フロントから email を受け取らない
 *   2. 書き込む email は必ず guardMember_() が返した検証済みのもの
 *   3. 返す前に必ず sanitize（email → uid、スプレッドシート ID を含めない）
 *   4. update 系は行所有者チェック（assertRowOwner_）を通す
 *
 * B は「自分（アプリアカウント）として実行」なので、ここでガードを抜くと
 * 誰でも全クラスのデータを読み書きできてしまう。フロントの出し分けは防御ではない。
 */

/**
 * 参加状態の確認（名簿にいない状態でも呼べる）。
 * フロントはこの結果で「参加申請フォーム / 承認待ち画面 / 本体」を出し分ける。
 */
function mbGetStatus(idToken, tenantCode) {
  try {
    const user = verifyIdToken_(idToken);                 // ①
    const code = normalizeTenantCode_(tenantCode);
    const rec = requireTenantRecord_(code);               // ②
    const ss = openTenantSs_(code);
    const row = getMemberRow_(ss, user.email);            // ③（照合のみ・拒否はしない）
    return jsonOk_({
      tenantName: rec.tenantName,
      joinOpen: !!rec.joinOpen,
      status: row ? row.status : 'unregistered',
      name: row ? row.name : ''
    });
  } catch (e) {
    return jsonErr_(e);
  }
}

/**
 * 参加申請（まだ名簿にいない状態で呼ばれる唯一の書き込み API）。
 * guardMember_ は名簿照合まで行うためここでは使わず、①②のみ手動で行う。
 */
function mbRequestJoin(idToken, tenantCode, displayName) {
  try {
    const user = verifyIdToken_(idToken);                 // ①
    const code = normalizeTenantCode_(tenantCode);
    const rec = requireTenantRecord_(code);               // ②
    if (!rec.joinOpen) throw new Error('JOIN_CLOSED: いまは参加の受付が止まっています。先生に確認してください');
    const ss = openTenantSs_(code);
    const name = vStr_(displayName, 30, '名前').trim();
    if (!name) throw new Error('BAD_INPUT: 名前を入力してください');

    const existing = getMemberRow_(ss, user.email);
    if (existing && existing.status === 'active') {
      return jsonOk_({ status: 'active' });
    }
    // クラスコードは宛先であって認証ではない。既定は承認制で pending 止まりにする
    const status = rec.requireApproval ? 'pending' : 'active';
    withScriptLock_(function () {
      upsertMember_(ss, {
        email: user.email,          // ★ 検証済み email のみ
        name: name,
        role: '児童',
        status: status
      });
    });
    return jsonOk_({ status: status });
  } catch (e) {
    return jsonErr_(e);
  }
}

/** 初期同期。自分の情報・今日のテーマ・自分のジャーナルのみ返す（すべてサニタイズ済み） */
function mbSync(idToken, tenantCode) {
  try {
    const g = guardMember_(idToken, tenantCode);          // ①②③
    return jsonOk_({
      tenantName: g.rec.tenantName,                       // ★ spreadsheetId は返さない
      me: { uid: uidOf_(g.user.email), name: g.member.name, role: g.member.role === '担任' ? '担任' : '児童' },
      todayTheme: getTodayTheme_(g.ss),
      journals: sanitizeJournals_(getJournalsForEmail_(g.ss, g.user.email))
    });
  } catch (e) {
    return jsonErr_(e);
  }
}

/**
 * ジャーナル提出。payload はホワイトリストしたキーのみ書き込む。
 * 画像は Data URL で受け取り、シート内にチャンク保存する（Drive 権限は一切使わない）。
 */
function mbSaveJournal(idToken, tenantCode, payload) {
  try {
    const g = guardMember_(idToken, tenantCode);          // ①②③
    const d = payload || {};
    const content = vStr_(d.content, 5000, '本文');
    if (!content.trim()) throw new Error('BAD_INPUT: 本文が空です。ひとことでも書いてから提出してね');
    const emotion = vStr_(d.emotion, 20, '気持ち');
    const theme = getTodayTheme_(g.ss);                   // テーマはサーバー側で決める

    let imageId = '';
    const sheet = getJournalSheet_(g.ss);
    withScriptLock_(function () {                         // 保持区間は最小に
      if (d.imageDataUrl) {
        imageId = saveImageChunks_(g.ss, d.imageDataUrl, g.user.email);
      }
      sheet.appendRow([
        Utilities.getUuid(), new Date(), g.user.email,    // ★ 検証済み email のみ
        theme, content, imageId, emotion, '', '[]', '', '未返却', '', ''
      ]);
    });
    return jsonOk_({ message: 'ジャーナルを提出しました！' });
  } catch (e) {
    return jsonErr_(e);
  }
}

/** 過去の自分へのメッセージ。⑤ 行所有者チェックを必ず通す */
function mbAddPastComment(idToken, tenantCode, journalId, comment) {
  try {
    const g = guardMember_(idToken, tenantCode);          // ①②③
    const sheet = getJournalSheet_(g.ss);
    const text = vStr_(comment, 1000, 'メッセージ');
    return withScriptLock_(function () {
      const headers = getJournalHeaders_(sheet);
      const rowNum = findJournalRowById_(sheet, vJournalId_(journalId));
      if (rowNum < 0) throw new Error('NOT_FOUND: 見つかりませんでした');
      const rowEmail = sheet.getRange(rowNum, headers.indexOf('email') + 1).getValue();
      assertRowOwner_(rowEmail, g.user.email, g.member, false);   // ⑤
      sheet.getRange(rowNum, headers.indexOf('pastComment') + 1).setValue(text);
      return jsonOk_({ message: '保存しました！' });
    });
  } catch (e) {
    return jsonErr_(e);
  }
}

/** 画像取得。自分がアップロードした画像のみ返す（⑤ 行所有者チェック相当） */
function mbGetImage(idToken, tenantCode, imageId) {
  try {
    const g = guardMember_(idToken, tenantCode);          // ①②③
    const img = loadImage_(g.ss, imageId);
    if (!img) throw new Error('NOT_FOUND: 画像が見つかりません');
    if (String(img.email).toLowerCase() !== g.user.email) {
      throw new Error('FORBIDDEN: 自分のデータ以外は見られません');
    }
    return jsonOk_({ dataUrl: img.dataUrl });
  } catch (e) {
    return jsonErr_(e);
  }
}
