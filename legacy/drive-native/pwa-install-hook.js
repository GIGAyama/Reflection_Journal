/*
 * インストールの合図を「いちばん先に」受け取るためのファイル。
 *
 * Chrome は条件が揃うと即座に beforeinstallprompt を出す。アプリ本体の
 * 読み込みより後で待ち構えていると、通信が遅い端末ではすでに合図が
 * 飛んだ後になってしまい、「アプリを入れる」ボタンが出なくなる。
 * だから <head> のいちばん上で、この小さなファイルだけを先に読み込む。
 *
 * インラインの <script> ではなく外部ファイルにしているのは、
 * 将来 CSP を script-src 'self' で締められるようにするため。
 */
(function () {
  window.__deferredInstallPrompt = null;

  window.addEventListener('beforeinstallprompt', function (e) {
    e.preventDefault();
    window.__deferredInstallPrompt = e;
    window.dispatchEvent(new Event('pwa-installable'));
  });

  window.addEventListener('appinstalled', function () {
    window.__deferredInstallPrompt = null;
    window.dispatchEvent(new Event('pwa-installed'));
  });
})();
