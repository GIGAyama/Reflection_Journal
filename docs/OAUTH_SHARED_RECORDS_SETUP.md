# 複数アカウント同期のOAuth設定

## なぜ追加設定が必要か

`drive.file` のファイル認可はユーザーごとです。児童が同じアプリでファイルを作成して教師へDrive共有しても、その共有だけで教師側のアプリに読取認可は付きません。

そこで、書込みは従来どおり `drive.file` に限定し、別アカウントから共有された提出・おへんじを読む時だけ `drive.readonly` を段階的に要求します。アプリの検索条件は `sharedWithMe` と `rjType` / `rjClassId` に限定されています。

Google公式資料：

- [Google Drive APIのスコープを選択する](https://developers.google.com/workspace/drive/api/guides/api-specific-auth)
- [Google WorkspaceでのOAuthに関する追加事項](https://developers.google.com/identity/protocols/oauth2/production-readiness/google-workspace)

## Cloud Consoleで先に行う作業

コードを公開する前に、共通OAuthクライアントを持つGoogle Cloudプロジェクトで行います。

1. Google Cloud Consoleの［Google Auth Platform］を開きます。
2. ［データアクセス］で［スコープを追加または削除］を開きます。
3. 次の2つを追加して保存します。
   - `https://www.googleapis.com/auth/drive.file`
   - `https://www.googleapis.com/auth/drive.readonly`
4. 検証中は［対象］のテストユーザーへ、確認に使う教師・児童の個人アカウントを登録します。
5. 学校アカウントで確認する場合は、学校のWorkspace管理者へOAuthクライアントIDと上記スコープを伝え、組織の規則に沿って信頼済みアプリとして許可してもらいます。

`drive.readonly` は制限付きスコープです。外部の任意のGoogleアカウントへ公開する前に、ブランド確認と制限付きスコープのOAuth審査を申請してください。Googleの判定やデータの取扱い方によってはセキュリティ評価が必要です。

## 今回の安全な反映順序

1. Cloud Consoleへ `drive.readonly` とテストユーザーを追加する。
2. 修正PRをマージする。
3. GitHub Pagesの反映を待ち、既存タブを再読込する。PWAでは［さいしんにする］を押す。
4. 教師・児童の両方で、初回に表示される説明を読み［共有された記録の同期を許可する］を押す。
5. 教師画面で対象クラスを開き［最新に更新］を押す。
6. 既存の未反映分と、新しく提出した1件の両方が表示されることを確認する。
7. 教師が返却し、児童側で［共有された記録の同期］後におへんじが表示されることを確認する。

既存の児童ポートフォリオは作り直しません。Drive共有が成功済みなら、追加許可後の同期で教師画面へ取り込まれます。

## 審査用のスコープ説明例

> 本アプリは学校向けの振り返り・教師フィードバックアプリです。児童と教師はそれぞれ自分のGoogle Driveで記録を所有し、相手へreader共有します。`drive.file` はアプリ作成ファイルの作成・更新に使用します。`drive.readonly` は、別の利用者から共有されたアプリ専用JSONと添付画像を自動検出・表示するために必要です。クライアントは `sharedWithMe` およびアプリ固有の `appProperties` で検索対象を限定します。OAuthアクセストークンはブラウザメモリ内だけに保持し、サーバーやブラウザストレージへ保存しません。

審査申請では、ログイン、追加許可の説明画面、児童の提出、教師での同期、教師の返却、児童でのおへんじ表示までを動画で示します。
