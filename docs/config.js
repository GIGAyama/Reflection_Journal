// ============================================================
// ふりかえりジャーナル — シェル設定
// ------------------------------------------------------------
// 運営者が一度だけ設定する共通OAuthクライアント。学校ごとのGASデプロイは不要。
// ============================================================
window.APP_CONFIG = {
  // QRコード・専用URLに使う共通入口。
  publicEntryUrl: "https://reflection-journal.giga-school.com/",
  // 本番で実行を許可するオリジン。
  // ⚠️ ここに載っていないオリジンでは、資産があってもログインできない
  //    （kit/session.js の originAllowed）。ドメインを変えるときは
  //    publicEntryUrl とこの値、Google OAuth の承認済みJavaScript生成元の
  //    3つを必ず同時に変えること。1つでも旧オリジンのまま残ると、
  //    アプリは開けるのにログインだけができない状態になる。
  allowedOrigins: ["https://reflection-journal.giga-school.com"],
  // 学校で利用を限定するときは ["school.example.jp"] のように設定する。
  // 空配列はドメイン制限なし（公開デモ用）。
  allowedWorkspaceDomains: [],
  // 独自ドメインへ移り、このアプリは reflection-journal.giga-school.com を
  // 単独で使うようになった。つまり「他のアプリと保存領域を共有しているから
  // 残せない」という以前の理由は、もう当てはまらない。
  //
  // それでも false のままにしてある。教室のChromebookは共用端末で、
  // 次に使う児童が前の児童のトークンや下書きを引き継いでしまうほうが困るため。
  // 残す判断をする場合は、端末の使われ方とあわせて決めること。
  persistSessionToken: false,
  persistLocalDrafts: false,
  // GIS OAuthウェブクライアントID。Drive APIを有効化し、
  // 承認済みJavaScript生成元へ https://reflection-journal.giga-school.com を登録する。
  googleClientId: "1074253864383-7v1d1gvtsm0ga9jonndslkesp030nd09.apps.googleusercontent.com"
};
