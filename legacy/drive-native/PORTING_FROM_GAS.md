# GASマルチテナント版から、Driveネイティブ分散ポートフォリオへ移す手順

このリポジトリは、同じアプリを2つの方式で作った記録が残っている。

- **旧: GASマルチテナント**（ルート直下の `*.gs`）— 運営者のGASプロジェクトが全学級のスプレッドシートを開き、中央レジストリで学級を引く。
- **新: Driveネイティブ**（`docs/`）— GitHub Pagesの共通URLだけを配り、ログイン中の本人の権限でDrive APIを呼ぶ。児童データは運営者のサーバーにもDBにも入らない。

新方式の共通部分は `docs/kit/` に切り出してある。**他のGASマルチテナント・アプリへ横展開するときは、`docs/kit/` をそのままコピーし、アプリ固有の部分だけを書く。**

---

## 1. どちらの方式を選ぶか

| 判断材料 | GASマルチテナント | Driveネイティブ |
| --- | --- | --- |
| 利用者データの置き場所 | 運営者が用意したスプレッドシート | 利用者本人のDrive |
| 学校ごとの作業 | デプロイまたは共有設定 | 不要（共通URLのみ） |
| 学校ドメインでの導入 | 外部スクリプトの実行許可が要る | 管理者がOAuthクライアントを許可する |
| 実行回数・容量の上限 | GASの割当てを全校で分け合う | 利用者ごとのDrive割当て |
| 必要なOAuthスコープ | GAS実行のスコープ | `drive.file` ＋ 共有同期時のみ `drive.readonly`（制限付き） |
| 主催者が全データを一括操作 | できる | できない（本人所有のため） |
| 利用者間の共有が禁止された環境 | 動く | **動かない**（管理者の共有許可が前提） |

**Driveネイティブが向くのは、「本人の成果物を本人が持ち、相手へ渡す」形のアプリ**（ふりかえり、作品、日誌、記録カード）。
全員分を集計して1枚の台帳に書き続けるアプリや、利用者間の共有を禁止している環境では、GAS方式のままにする。

---

## 2. 置き換えの対応表

| GAS側（このリポジトリの旧実装） | Driveネイティブ側 | 要点 |
| --- | --- | --- |
| `Registry.gs`（ScriptPropertiesの中央レジストリ `tn_<CODE>`） | 無し。テナントIDを `namespace.tenantId(主催者メール, コード)` で**決定的に計算** | 採番簿が要らない。コード単独での検索はできない（＝総当りで他学級を引けない） |
| `Registry.gs` の `generateTenantCode_`（衝突検査つき採番） | `namespace.randomCode()` ＋ 署名付き招待 | 中央で衝突を見られない代わりに、コードだけでは参加できない設計にする |
| `Auth.gs`（IDトークンをサーバーで検証） | `kit/session.js` の `fetchUserInfo` ＋ Drive API自身の認可 | 「誰か」を決めるのはGoogle。アプリは検証済みメール以外を書き込み者にしない |
| `Tenant.gs` の `guardMember_`（トークン→テナント→名簿の多段ガード） | `kit/records.js` の `validateSharedRecord` ＋ `kit/invite.js` の `matchesIssuedKey` | サーバーが無いので、**取り込む側の端末で**同じ4点（種別・テナント・Drive所有者・宛先）を必ず見る |
| `Tenant.gs` の `assertRowOwner_`（行の所有者チェック） | ファイルを分けて所有権で分離する | 主催者の返却と参加者の本文を**別ファイル**にする。相手の本文はそもそも書けない |
| `Db.gs`（スプレッドシートの行操作、`LockService`） | JSONファイル1本＋Driveの `version` 照合 | 書く前に版を読み、画面が読んだ版と違えば上書きしない |
| `OwnerApi.gs` / `MemberApi.gs`（役割ごとのサーバーAPI） | 役割ごとの画面（同じJSを共通URLで配る） | 「フロントの出し分けは防御ではない」は同じ。防御はDriveの共有権限そのもの |
| スプレッドシートIDの秘匿 | Drive `appProperties` の印（`<prefix>Type` / `<prefix><Tenant>Id`） | IDは秘密ではない。共有されていないファイルは検索結果に出ない |
| 学級ごとのスプレッドシート共有設定 | 参加時の `shareWithUser`（通知メールなし） | 利用者に共有設定をさせない |

---

## 3. 移し替えの手順

### 手順1 — アプリの名前空間を決める

`docs/kit/` をコピーし、名前空間を1か所で定義する。**`appId` と `propertyPrefix` は既存アプリと必ず変える。** 同じ値を使うと、同じDrive上で別アプリのファイルを拾う。

```js
import { createDriveNativeApp } from './kit/index.js';

export const APP = createDriveNativeApp({
  appId: 'kanji-drill',            // ID計算の接頭辞。他アプリと重複させない
  propertyPrefix: 'kd',            // appProperties の接頭辞（kdType / kdRoomId …）
  schemaVersion: 1,                // 保存形式の版。上げると過去の招待・IDは無効になる
  terms: { tenant: 'Room', member: 'Learner' },
  clientId: window.APP_CONFIG.googleClientId,
  allowedOrigins: window.APP_CONFIG.allowedOrigins,
  allowedDomains: window.APP_CONFIG.allowedWorkspaceDomains,
  persistSessionToken: false,      // 共有オリジンでは既定のまま false
  entryUrl: window.APP_CONFIG.publicEntryUrl
});
```

### 手順2 — ファイルの割り方を決める（ここが設計の中心）

**同じJSONを2人が更新する形にしない。** 書く人ごとにファイルを分け、相手には読み取りだけを渡す。

| 種別 | 所有者 | 共有先 | 中身 |
| --- | --- | --- | --- |
| tenant（設定） | 主催者 | 共有しない | 名簿、参加状態、招待署名鍵（秘密鍵） |
| record（成果物） | 参加者 | 主催者へ reader | 本人が書いたもの |
| channel（返却） | 主催者 | 参加者へ reader | テーマ、コメント、承認状態 |

ふりかえりジャーナルでは class / portfolio / channel / journal-image がこれにあたる（`docs/DRIVE_NATIVE_ARCHITECTURE.md`）。

### 手順3 — 参加の証明を招待署名にする

GAS版の「クラスコードを知っていれば参加申請できる」は、中央で名簿照合ができたから成り立っていた。中央が無い新方式では、**署名付き招待URL**が参加資格の証明になる。

```js
// 主催者側: テナント作成時に鍵を作り、非共有ファイルへ入れる
const security = await APP.invites.createKey({ validityDays: 35 });
const token = await APP.invites.sign(
  { code, name: roomName, hostEmail: user.email, hostName: user.name }, security);
const url = APP.invites.url('', token);          // #join=… （フラグメントなのでログに残らない）

// 参加者側: 復号して、テナントIDを得る
const invite = await APP.invites.decode(token);  // 署名・期限・必須項目を確認する

// 主催者側: 届いた記録を名簿へ入れる前に、自分が発行した鍵と一致するか見る
const issued = APP.invites.matchesIssuedKey(invite, security, { tenantId });
```

秘密鍵は主催者所有の**非共有**ファイルにだけ置く。鍵を入れ替える（`generation` を上げる）と、配り終えた古いURLは失効する。

### 手順4 — 取り込み時の検証を書く

自己申告のJSONだけを信じない。Driveのメタデータ側と必ず突き合わせる。

```js
import { validateSharedRecord } from './kit/records.js';

const base = validateSharedRecord(file, record, {
  kind: 'kanji-drill-record',
  tenantId,
  tenantIdOf: (value) => value?.tenant?.id,
  ownerMustBe: (value) => value?.learner?.email   // 本人所有でなければ通さない
});
```

アプリ固有の条件（招待の世代一致、有効期限内に作られたか等）は、この後ろに足す。実例は `docs/drive-core.js` の `validatePortfolioForClass`。

### 手順5 — 権限を段階的に求める

- 最初は `APP.scopes.base`（`openid email profile drive.file`）だけを求める。自分のファイルの作成・更新・共有はこれで足りる。
- **共有された記録の同期を始める直前に**、`APP.scopes.sharedRead`（`drive.readonly`）の用途を画面で説明してから求める。これはGoogleの制限付きスコープで、許可画面には「Drive全体の閲覧」と出る。実装が読むのは `sharedWithMe` かつ自分の印が付いたファイルだけだ、と画面に明記する。
- `drive.file` はユーザーごと・ファイルごとの認可なので、**Driveで共有しただけでは相手側アプリの認可にならない**。共有記録の自動同期を `drive.file` だけで作ろうとしないこと（ここで何度も詰まる）。

### 手順6 — 競合と保存領域

- 更新前に `file.version` を読み、画面が読んだ版と違えば書かずに再読込みを求める（`updateJson(..., { expectedVersion })`）。
- 共有オリジン（`*.github.io`）では、アクセストークンを `sessionStorage` / `localStorage` / Cookie に置かない。既定 `persistSessionToken: false` のままにする。
- 未送信の下書きも共有オリジンの永続領域へ書かない。メモリに置く。

### 手順7 — 運用の準備

- Google Cloud のOAuthクライアントに、本番オリジンを「承認済みJavaScript生成元」として登録する。
- 制限付きスコープ（`drive.readonly`）を使うので、外部公開前にOAuth審査を通す。判定によってはセキュリティ評価が要る。
- Workspace管理者に、共通OAuthクライアントIDと利用スコープの許可、および利用者間のDrive共有の許可を依頼する。**共有が禁止されている環境では動かない。** 回避しない。

---

## 4. データ移行

旧GAS版のスプレッドシートから新方式へ持ち込む場合、**運営者が一括で移すことはできない**（新方式では本人しかファイルを作れない）。現実的な進め方は次の2つ。

1. **年度・学期の区切りで切り替える。** 旧データはGAS版を読み取り専用で残し、新方式は新規から始める。いちばん安全で、実際にこのアプリはこの形をとった。
2. **本人がエクスポート→インポートする。** 旧版でCSV/JSONを本人に出させ、新方式の初回ログイン時に取り込む画面を作る。移行UIの手間に見合うかを先に見積もること。

---

## 5. 横展開のチェックリスト

移し替えたアプリを公開する前に、次を確認する。

- [ ] `appId` と `propertyPrefix` を既存アプリと変えた（同一Drive上での取り違え防止）
- [ ] 参加者が書くファイルと主催者が書くファイルを分けた（同一JSONの同時更新をしていない）
- [ ] 招待は署名付きで発行し、取り込み時に鍵ID・世代・有効期限を照合している
- [ ] 共有記録の取り込みで、Driveの所有者（`owners[0].emailAddress`）を必ず見ている
- [ ] 最初の認可は `drive.file` までで、`drive.readonly` は用途説明のあとに求めている
- [ ] 更新前に `version` を照合し、競合時は上書きせず再読込みを求めている
- [ ] アクセストークンと未送信の下書きを、共有オリジンの保存領域へ書いていない
- [ ] 403（学校ポリシーによる拒否）を、利用者が次にやることが分かる文言で表示している
- [ ] 実アカウント2つ（別ドメインを含む）で、作成→参加→承認→提出→返却を通した
- [ ] Service Worker のキャッシュ接頭辞を自アプリ専用にした（同一ドメインの他アプリを巻き添えにしない）

---

## 6. 参考

- 設計の全体像: `docs/DRIVE_NATIVE_ARCHITECTURE.md`
- キットの詳細: `docs/kit/README.md`
- OAuth設定の手順: `docs/OAUTH_SHARED_RECORDS_SETUP.md`
- 公開前の運用確認: `docs/SECURITY_DEPLOYMENT.md` / `ROLLOUT.md`
