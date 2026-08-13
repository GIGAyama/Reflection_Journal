// ============================================================
// ふりかえりジャーナル — シェル設定
// ------------------------------------------------------------
// 運営者が一度だけ設定する共通OAuthクライアント。学校ごとのGASデプロイは不要。
// ============================================================
window.APP_CONFIG = {
  // QRコード・専用URLに使うGitHub Pagesの共通入口。
  publicEntryUrl: "https://gigayama.github.io/Reflection_Journal/",
  // GIS OAuthウェブクライアントID。Drive APIを有効化し、
  // 承認済みJavaScript生成元へ https://gigayama.github.io を登録する。
  googleClientId: "1074253864383-7v1d1gvtsm0ga9jonndslkesp030nd09.apps.googleusercontent.com"
};
