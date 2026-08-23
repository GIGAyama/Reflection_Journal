/**
 * MemberApi.gs — 児童の画面から呼ぶ API
 *
 * ★ 鉄則 ★
 *   1. 書き込む email は必ず guardMember_() が返した確認済みのものだけ。
 *      画面から email を受け取らない（受け取れば、誰にでもなりすませる）
 *   2. 返す前に必ず sanitize（ほかの児童のメールアドレスを出さない）
 *   3. update 系は行の持ち主チェック（assertRowOwner_）を通す
 *   4. 一斉に叩かれる書き込みは LockService で囲む。囲む幅は appendRow 1 回分に絞る
 *
 * デプロイは「実行するユーザー: 自分（先生）」なので、ここでガードを抜くと
 * 児童が学級全員の記録を読み書きできてしまう。画面の出し分けは防御ではない。
 */

/**
 * 参加状態の確認（名簿にいない状態でも呼べる）。
 * 画面はこの結果で「参加申請フォーム / 承認待ち画面 / 本体」を出し分ける。
 */
function mbGetStatus() {
  try {
    const email = requireEmail_();                        // ①
    const ss = getDb_();
    if (!ownerEmailOf_(ss)) {
      throw new Error('SETUP_REQUIRED: このクラスはまだ準備中です。先生に確認してください');
    }
    const row = getMemberRow_(ss, email);                 // ②（照合のみ・拒否はしない）
    return jsonOk_({
      className: String(readMeta_(ss, META_KEYS.CLASS_NAME) || CONFIG.APP_NAME),
      joinOpen: getTenantSetting_(ss, 'JOIN_CLOSED') !== '1',
      status: row ? row.status : 'unregistered',
      name: row ? row.name : ''
    });
  } catch (e) {
    return jsonErr_(e);
  }
}

/**
 * 参加申請（まだ名簿にいない状態で呼ばれる唯一の書き込み API）。
 * guardMember_ は名簿照合まで行うためここでは使わず、①のみ手動で行う。
 *
 * 既定は承認制。URL を知っているだけでは名簿に active では入れない。
 */
function mbRequestJoin(displayName) {
  try {
    const email = requireEmail_();                        // ①
    const ss = getDb_();
    if (!ownerEmailOf_(ss)) {
      throw new Error('SETUP_REQUIRED: このクラスはまだ準備中です。先生に確認してください');
    }
    if (getTenantSetting_(ss, 'JOIN_CLOSED') === '1') {
      throw new Error('JOIN_CLOSED: いまは参加の受付が止まっています。先生に確認してください');
    }
    const name = vStr_(displayName, 30, '名前').trim();
    if (!name) throw new Error('BAD_INPUT: 名前を入力してください');

    const existing = getMemberRow_(ss, email);
    if (existing && existing.status === 'active') {
      return jsonOk_({ status: 'active' });
    }
    const status = getTenantSetting_(ss, 'AUTO_APPROVE') === '1' ? 'active' : 'pending';
    withScriptLock_(function () {
      upsertMember_(ss, {
        email: email,          // ★ 確認済みのアドレスだけ
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

/** 初期同期。自分の情報・今日のテーマ・自分のジャーナルだけを返す */
function mbSync() {
  try {
    const g = guardMember_();                             // ①②
    return jsonOk_({
      className: String(readMeta_(g.ss, META_KEYS.CLASS_NAME) || CONFIG.APP_NAME),
      me: {
        uid: uidOf_(g.email),
        name: g.member.name,
        role: g.member.role === '担任' ? '担任' : '児童'
      },
      todayTheme: getTodayTheme_(g.ss),
      journals: sanitizeJournals_(getJournalsForEmail_(g.ss, g.email))
    });
  } catch (e) {
    return jsonErr_(e);
  }
}

/**
 * ジャーナル提出。payload は決めたキーだけ書き込む。
 * 画像は Data URL で受け取り、シートの中に切って入れる（Drive 権限は一切使わない）。
 */
function mbSaveJournal(payload) {
  try {
    const g = guardMember_();                             // ①②
    const d = payload || {};
    const content = vStr_(d.content, 5000, '本文');
    if (!content.trim()) throw new Error('BAD_INPUT: 本文が空です。ひとことでも書いてから提出してね');
    const emotion = vStr_(d.emotion, 20, '気持ち');
    const theme = getTodayTheme_(g.ss);                   // テーマはサーバー側で決める

    const sheet = getJournalSheet_(g.ss);
    withScriptLock_(function () {                         // 保持区間は最小に
      const imageId = d.imageDataUrl ? saveImageChunks_(g.ss, d.imageDataUrl, g.email) : '';
      appendJournalRow_(sheet, {
        journalId: Utilities.getUuid(),
        timestamp: new Date(),
        email: g.email,                                   // ★ 確認済みのアドレスだけ
        theme: theme,
        content: content,
        imageFileId: imageId,
        emotion: emotion,
        teacherComment: '',
        highlights: '[]',
        teacherStamp: '',
        status: '未返却',
        pastComment: '',
        deletedAt: ''
      });
    });
    return jsonOk_({ message: 'ジャーナルを提出しました！' });
  } catch (e) {
    return jsonErr_(e);
  }
}

/** 過去の自分へのメッセージ。④ 行の持ち主チェックを必ず通す */
function mbAddPastComment(journalId, comment) {
  try {
    const g = guardMember_();                             // ①②
    const sheet = getJournalSheet_(g.ss);
    const text = vStr_(comment, 1000, 'メッセージ');
    return withScriptLock_(function () {
      const cols = journalCols_(sheet);
      const rowNum = findJournalRowById_(sheet, vJournalId_(journalId));
      if (rowNum < 0) throw new Error('NOT_FOUND: 見つかりませんでした');
      const rowEmail = sheet.getRange(rowNum, cols.email).getValue();
      assertRowOwner_(rowEmail, g.email, g.member, false);   // ④
      sheet.getRange(rowNum, cols.pastComment).setValue(text);
      return jsonOk_({ message: '保存しました！' });
    });
  } catch (e) {
    return jsonErr_(e);
  }
}

/** 画像取得。自分がアップロードした画像だけを返す（④ 相当） */
function mbGetImage(imageId) {
  try {
    const g = guardMember_();                             // ①②
    const img = loadImage_(g.ss, imageId);
    if (!img) throw new Error('NOT_FOUND: 画像が見つかりません');
    if (String(img.email).toLowerCase() !== g.email) {
      throw new Error('FORBIDDEN: 自分のデータ以外は見られません');
    }
    return jsonOk_({ dataUrl: img.dataUrl });
  } catch (e) {
    return jsonErr_(e);
  }
}
