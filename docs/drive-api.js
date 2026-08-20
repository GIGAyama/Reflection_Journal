// ============================================================
// ふりかえりジャーナル — Drive呼び出し
// ------------------------------------------------------------
// 通信・競合検知・エラー文言の共通部分は docs/kit/drive-client.js にある。
// ここには「このアプリがDriveへ置くファイルの名前と種別」だけを書く。
// ============================================================

import { RJ } from './drive-core.js';
import { DriveApiError, DriveConflictError, KitDriveClient } from './kit/drive-client.js';

export { DriveApiError, DriveConflictError };

const TYPE = { class: 'class', portfolio: 'portfolio', channel: 'channel', image: 'journal-image' };

export class DriveClient extends KitDriveClient {
  constructor(accessToken, fetchImpl) {
    super({ accessToken, namespace: RJ, fetchImpl });
  }

  listClasses() {
    // 先生が開くのは自分が作ったクラスだけ。共有された同種のファイルまで拾わないよう所有者で絞る。
    return this.listByType({ type: TYPE.class, owner: 'me' });
  }

  listOwnPortfolios(classId = '') {
    return this.listByType({ type: TYPE.portfolio, owner: 'me', tenantId: classId });
  }

  listSharedPortfolios(classId) {
    return this.listByType({ type: TYPE.portfolio, owner: 'shared', tenantId: classId });
  }

  listSharedChannels(classId = '') {
    return this.listByType({ type: TYPE.channel, owner: 'shared', tenantId: classId });
  }

  listOwnChannels(classId = '') {
    return this.listByType({ type: TYPE.channel, owner: 'me', tenantId: classId });
  }

  createClass(record) {
    return this.createJson(
      `ふりかえりジャーナル_クラス_${record.className}.rj-class.json`,
      record,
      RJ.appProperties(record.classId, TYPE.class)
    );
  }

  createPortfolio(record) {
    return this.createJson(
      `ふりかえりジャーナル_${record.class.name}_${record.student.name}.rj.json`,
      record,
      RJ.appProperties(record.class.id, TYPE.portfolio)
    );
  }

  createChannel(record, studentHash) {
    return this.createJson(
      `ふりかえりジャーナル_おへんじ_${record.class.name}_${record.student.name}.rj-channel.json`,
      record,
      RJ.appProperties(record.class.id, TYPE.channel, { [RJ.properties.member]: studentHash })
    );
  }

  createJournalImage({ classId, journalId, studentName, file }) {
    const extension = String(file.name || '').match(/\.[a-z0-9]+$/i)?.[0] || '.jpg';
    return this.createFile({
      name: `ふりかえり_${studentName}_${journalId}${extension}`,
      mimeType: file.type || 'image/jpeg',
      appProperties: { ...RJ.appProperties(classId, TYPE.image), rjJournalId: journalId }
    }, file);
  }
}
