// ============================================================
// ふりかえりジャーナル — シェル設定
// ------------------------------------------------------------
// デプロイ後にこの 3 つを埋める（README のセットアップ手順参照）。
//
// exec URL をハードコードせずここから読む形を維持すること。
// 将来テナント数が増えて水平分割（クラス群ごとにデプロイBを別アプリアカウントで
// 発行）する際、ここを差し替えるだけで済むようにするため。
// ============================================================
window.APP_CONFIG = {
  // デプロイ A（先生ポータル / 実行: ウェブアプリケーションにアクセスしているユーザー）の exec URL
  execUrlA: "https://script.google.com/macros/s/AKfycbz5_km-JRG0kNz8sSUkKnqRhV07lbV_OG3-S-FujUnHsdEiAoqUr77ihhuH3abZrS0z/exec",
  // デプロイ B（児童アプリ / 実行: 自分＝アプリアカウント / アクセス: 全員）の exec URL
  execUrlB: "https://script.google.com/macros/s/AKfycbzM-RJd-ubOwJUiIgt9X2YJoCjHz7A2_QinxyHDWvHPqw8sAud_v2lTpQEyWCa3khdL/exec",
  // GIS 用 OAuth クライアント ID（ウェブアプリケーション種別）。
  // GAS 側 ScriptProperties の sp_googleClientId と同じ値にすること（aud 検証に使う）
  googleClientId: "1074253864383-7v1d1gvtsm0ga9jonndslkesp030nd09.apps.googleusercontent.com"
};
