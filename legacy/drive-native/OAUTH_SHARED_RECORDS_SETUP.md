# 複数アカウント同期のOAuth設定

この文書は、GitHub Pagesの共通URLからGoogle Driveへ直接接続する現在版について、アプリ運営者がGoogle Cloud／Google Workspaceで行う共通設定を説明します。先生・児童や学校ごとのGASデプロイ手順ではありません。

## なぜ2つのDriveスコープが必要か

`drive.file` のファイル認可は利用者・ファイルごとです。児童がこのアプリでファイルを作成して先生へDrive共有しても、その共有だけで先生側のブラウザアプリに読取認可が付くわけではありません。

そこで、用途を次のように分けます。

| スコープ | 実装上の用途 |
| --- | --- |
| `https://www.googleapis.com/auth/drive.file` | ログイン中の利用者が、このアプリで作成したJSON／画像を作成・更新し、相手を閲覧者として共有する |
| `https://www.googleapis.com/auth/drive.readonly` | 別アカウントから共有された提出、テーマ、おへんじ、画像を検索・読込みする |

`drive.readonly` はDrive全体の閲覧を許可する制限付きスコープです。アプリの検索は `sharedWithMe` と `rjType`／`rjClassId` などのアプリ固有 `appProperties` に限定し、書込みには使いません。書込みは引き続き `drive.file` の専用ファイルだけです。

Google公式資料：

- [Google Drive APIのスコープを選択する](https://developers.google.com/workspace/drive/api/guides/api-specific-auth)
- [Google WorkspaceでのOAuthに関する追加事項](https://developers.google.com/identity/protocols/oauth2/production-readiness/google-workspace)

## 実際の認可タイミング

現在の `docs/drive-app.js` は、最初の［Googleアカウントで続ける］で次を要求します。

```text
openid
email
profile
https://www.googleapis.com/auth/drive.file
```

`drive.readonly` は、**相手の共有記録を実際に読む直前**に追加要求します。要求の起点は次の2か所だけです。

| 画面 | 実装 | 読むもの |
| --- | --- | --- |
| 先生がクラスを開く | `openTeacherClass` | 児童から共有された提出 |
| 児童がクラスを開く | `openPortfolio` | 先生から共有されたテーマとおへんじ |

クラス一覧、クラス作成、参加申請の画面は自分のDriveのファイルしか読まないため、`drive.file` だけで動きます。役割を選んだ時点では要求しません。

要求の前に、アプリ内へ［共有された記録の同期を許可する］画面を表示します。`drive.readonly` がDrive全体の閲覧を許可する制限付きスコープであること、実装はアプリ印の一致する共有ファイルに限定することを説明し、利用者の操作で追加要求します。断った場合はクラス一覧へ戻ります。

初回要求には `include_granted_scopes: true` を付けています（Googleが推奨する段階的認可の作法）。このため、一度 `drive.readonly` を許可した利用者は、次回以降のログインで受け取るトークンに最初からこのスコープが含まれ、アプリ内の用途説明画面は再表示されません。毎回の説明を必須にしたい場合は `include_granted_scopes` を `false` にしますが、その場合はセッションごとに説明画面と操作が1回ずつ増えます。

初回要求には `enable_granular_consent: true` も付けています。利用者が `drive.file` だけを外した場合は、最初のDrive操作で失敗させず、ログイン画面で理由を伝えて止めます。

## Google Cloudで行う共通設定

1. 共通OAuthクライアントを置くGoogle Cloudプロジェクトを選びます。
2. ［APIとサービス］でGoogle Drive APIを有効化します。
3. ［Google Auth Platform］でブランド名、サポートメール、対象ユーザーを設定します。
4. ［データアクセス］の［スコープを追加または削除］で次を追加します。
   - `https://www.googleapis.com/auth/drive.file`
   - `https://www.googleapis.com/auth/drive.readonly`
5. 検証中は［対象］のテストユーザーへ、確認に使う先生・児童アカウントを登録します。
6. OAuthクライアントを「ウェブアプリケーション」として作成します。
7. 承認済みJavaScript生成元へ、実際の本番オリジン（現在は `https://reflection-journal.giga-school.com`）を登録します。
   独自ドメインへ移行済みで、他アプリとは Web Storage も Service Worker もオリジンが分かれています。
   なお `docs/config.js` の `allowedOrigins` にも同じオリジンを入れる必要があります。片方だけ直すと、アプリは開けるのにログインだけができない状態になります。
8. クライアントIDを `docs/config.js` の `googleClientId` へ設定します。
9. `docs/config.js` の `publicEntryUrl`、`allowedOrigins`、組織限定時の `allowedWorkspaceDomains` を設定します。
10. GitHub Pagesの公開元を `main` ブランチの `/docs` にします。

ブラウザ用OAuthクライアントIDは公開識別子です。クライアントシークレット、サービスアカウント鍵、API秘密鍵をGitHub Pagesやリポジトリへ配置しないでください。

## 外部公開とOAuth審査

`drive.readonly` は制限付きスコープです。外部の任意のGoogleアカウントへ本番公開する前に、ブランド確認と制限付きスコープのOAuth審査を申請します。Googleの判定やデータの取扱い方によってはセキュリティ評価が必要です。

審査未完了の外部アプリでは、次のような制約が起こり得ます。

- 「未確認のアプリ」の警告
- テストユーザー以外の利用拒否
- テストユーザー数や認可の有効期間に関する制約
- 制限付きスコープの審査要求

## Google Workspace管理者へ伝える内容

学校アカウントで確認する場合は、学校のWorkspace管理者へ次を伝え、組織の規則に沿って信頼済みアプリとして許可してもらいます。

- アプリ名と共通URL
- Google CloudのOAuthクライアントID
- `drive.file`
- `drive.readonly`
- 児童→先生、先生→児童のユーザー指定Drive共有が必要であること
- データ所有者は各先生・児童で、運営者アカウントを共有先にしないこと

OAuthクライアントを許可しても、組織のDrive共有ポリシーが別途禁止していれば提出・返却は同期できません。OAuth許可とDrive共有許可の両方を確認します。アプリは学校の組織ポリシーを回避しません。

## 審査用のスコープ説明例

> 本アプリは学校向けの振り返り・教師フィードバックアプリです。児童と先生はそれぞれ自分のGoogle Driveで記録を所有し、相手へ閲覧者として共有します。drive.file はアプリ作成ファイルの作成・更新・共有に使用します。drive.readonly は、別の利用者から共有されたアプリ専用JSONと添付画像を自動検出・表示するために必要です。クライアントは sharedWithMe およびアプリ固有の appProperties で検索対象を限定します。既定でOAuthアクセストークンはブラウザのJavaScriptメモリにだけ保持し、sessionStorage、localStorage、Cookie、外部サーバーへ保存しません。

審査動画では、次を一続きで示します。

1. 共通URLとアプリのプライバシー説明を開く。
2. Googleログインで `drive.file` を認可し、クラスを開く時に出る用途説明から `drive.readonly` を追加認可する。
3. 先生がクラスを作成して専用URLを発行する。
4. 別アカウントの児童が参加し、本文と画像を提出する。
5. 先生が共有提出を同期して閲覧し、おへんじを返却する。
6. 児童が共有チャンネルからテーマとおへんじを受け取る。
7. アプリが専用ファイルだけを画面へ表示することを説明する。

## 実アカウント受け入れテスト

設定後は、コードの自動テストだけで完了とせず、次を確認します。

1. 個人Googleアカウント2つを先生・児童に分け、作成→招待→参加→承認→提出→画像→返却を確認する。
2. 同一学校ドメインの先生・児童アカウントで同じ流れを確認する。
3. 最初の認可では `drive.file` だけが要求され、クラス一覧・クラス作成の画面までは `drive.readonly` を求めないことを確認する。クラスを開く操作で用途説明と追加要求が出ることを確認する。
4. `drive.readonly` を許可しない場合、共有記録の同期へ進まないことを確認する。
5. 先生が［↻ 最新に更新］を押すと、既存の共有提出と新規提出の両方が表示されることを確認する。
6. 先生の返却後、児童がクラスを開き直すとおへんじが表示されることを確認する。
7. 再読込み後に再ログインが求められ、sessionStorage、localStorage、Cookieにアクセストークンが残らないことを確認する。
8. 共有を禁止したテスト環境で、403の利用者向け説明と［先生へもう一度届ける］を確認する。

既存ファイルを作り直す必要はありません。Drive共有が成功済みなら、必要な読取許可を得た後の同期で相手側へ表示されます。
