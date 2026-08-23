# Driveネイティブ・キット

GitHub Pagesの共通URLだけで動く「分散ポートフォリオ」を、アプリ間で使い回すための一式。
運営者のサーバーやDBへ利用者データを集めず、ログイン中の**本人の権限で**Drive APIを呼ぶ。

依存パッケージは無い。ブラウザのESモジュールとしてそのまま配信でき、Node.jsの試験からも直接読める。

## 入っているもの

| ファイル | 役割 |
| --- | --- |
| `namespace.js` | アプリ固有の印を1か所で決める。テナントID、メンバー鍵、`appProperties`、Drive検索式 |
| `invite.js` | 署名付き招待（ECDSA P-256）。発行・世代交代・復号・照合 |
| `drive-client.js` | Drive REST の呼び出し。一覧・作成・更新（版照合つき）・共有・フォルダ |
| `records.js` | 共有記録の検証、版キャッシュ、名簿統合 |
| `session.js` | 公開元・ドメインの許可、トークン保存方針、スコープの段階要求 |
| `index.js` | 上記をアプリ設定1つに束ねる `createDriveNativeApp()` |

キットは**画面文言を持たない**（Driveエラーの既定文言だけは学校向けに用意し、差し替え可能にしてある）。
招待の失敗は `InviteError.code`（`unreadable` / `invalid` / `expired` / `unsigned`）、Driveの失敗は
`DriveApiError.code`（`unauthorized` / `forbidden` / `conflict` / `unknown`）で返るので、アプリ側で自分の言葉に直す。

## 使い方

```js
import { createDriveNativeApp } from './kit/index.js';

const APP = createDriveNativeApp({
  appId: 'kanji-drill',                       // 他アプリと必ず異なる値
  propertyPrefix: 'kd',                       // appProperties の接頭辞
  schemaVersion: 1,
  terms: { tenant: 'Room', member: 'Learner' },
  clientId: window.APP_CONFIG.googleClientId,
  allowedOrigins: window.APP_CONFIG.allowedOrigins,
  entryUrl: window.APP_CONFIG.publicEntryUrl
});

const drive = APP.client(accessToken);
await drive.listByType({ type: 'record', owner: 'shared', tenantId });
await drive.createJson('学習記録.json', record, APP.namespace.appProperties(tenantId, 'record'));
await drive.updateJson(file.id, next, { expectedVersion: file.version });   // 競合したら書かない
await drive.shareWithUser(file.id, hostEmail, 'reader');                     // 通知メールは送らない
```

アプリ側が書くのは、**記録の形（JSONの中身）・画面・取り込み時の追加条件**だけ。
実装例はこのリポジトリ本体（`docs/drive-core.js` / `docs/drive-api.js` / `docs/drive-app.js`）。

## 決めごと（外さないこと）

1. **書く人ごとにファイルを分ける。** 同じJSONを2人が更新する形にしない。相手には reader だけを渡す。
2. **自己申告を信じない。** 共有記録は種別・テナント・Driveの所有者・宛先の4点を `validateSharedRecord` で見る。
3. **参加資格は署名付き招待で証明する。** コードを知っているだけでは名簿へ入れない。
4. **権限は段階的に。** 最初は `drive.file`。共有記録の同期を始める直前に `drive.readonly` を説明してから求める。
5. **共有オリジンに秘密を置かない。** トークンも未送信の下書きも、既定でメモリだけに持つ。

## 移し替えの手順

GASでマルチテナント化しているアプリからの移行手順・対応表・公開前チェックリストは
[`../PORTING_FROM_GAS.md`](../PORTING_FROM_GAS.md) にある。

## 試験

`npm test` に含まれる `scripts/test-drive-kit.mjs` が、キット単体の動作に加えて
**キットの計算結果が本体アプリの値と一致し続けること**を確認する。
キットを直して本体のIDや招待が変わると、この試験が落ちる（＝既存利用者のファイルが読めなくなる変更に気づける）。
