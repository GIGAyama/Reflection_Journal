// ============================================================
// ふりかえりジャーナル — シェル設定
// ------------------------------------------------------------
// 運営者が一度だけ設定する共通OAuthクライアント。学校ごとのGASデプロイは不要。
// ============================================================
window.APP_CONFIG = {
  // QRコード・専用URLに使うGitHub Pagesの共通入口。
  publicEntryUrl: "https://gigayama.github.io/Reflection_Journal/",
  // 本番で実行を許可するオリジン。専用カスタムドメインへ移行する場合は
  // publicEntryUrlとこの値、Google OAuthの承認済みJavaScript生成元を同時に変更する。
  allowedOrigins: ["https://gigayama.github.io"],
  // 学校で利用を限定するときは ["school.example.jp"] のように設定する。
  // 空配列はドメイン制限なし（公開デモ用）。
  allowedWorkspaceDomains: [],
  // 共有GitHub Pagesオリジンでは他アプリから分離できないため、OAuthトークンを
  // sessionStorageに残さない。リロード後は再ログインする。
  persistSessionToken: false,
  // 児童本文も共有オリジンの永続領域へ残さず、開いているページのメモリでだけ下書きを保持する。
  persistLocalDrafts: false,
  // GIS OAuthウェブクライアントID。Drive APIを有効化し、
  // 承認済みJavaScript生成元へ https://gigayama.github.io を登録する。
  googleClientId: "1074253864383-7v1d1gvtsm0ga9jonndslkesp030nd09.apps.googleusercontent.com"
};
