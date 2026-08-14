// ============================================================
// ふりかえりジャーナル — PWAシェル Service Worker
// ------------------------------------------------------------
// シェル資産（index.html / config.js / manifest / アイコン）のみをキャッシュする
// 最小構成。Google Identity / Drive APIなどの外部通信は
// 認証・データがあるため【絶対にキャッシュしない】。
// ============================================================

/* 【重要】キャッシュの掃除は、かならず自アプリのぶんだけに限る。
 *
 * gigayama.github.io は数十本の学習アプリが同じドメインを共有している。
 * ブラウザのキャッシュはドメイン単位なので、caches.keys() はこのアプリのものだけでなく、
 * 同居する全アプリのキャッシュを返す。
 *
 * これまでは「CACHE_NAME 以外ぜんぶ」を消していたため、ふりかえりジャーナルを開いて
 * 新しい Service Worker が有効になった瞬間、その端末に入っていた
 * 児童むけアプリ（Qalc・KANJI_Town など）のオフライン用データまで消えていた。
 * 児童がオフラインで開いても起動せず、しかも原因がそのアプリ側に見えないため
 * 「たまに開かなくなる」という再現しにくい不具合になっていた。
 *
 * Service Worker は localStorage を一切操作しない。 */
const CACHE_PREFIX = 'rj-shell-';
const APP_VERSION = 'v16';  // アプリ内の戻る操作・児童向けふりがな・ノート表示修正を配信
const CACHE_NAME = CACHE_PREFIX + APP_VERSION;
const SHELL_ASSETS = [
  './',
  './index.html',
  './diag.html',
  './drive.css',
  './drive-core.js',
  './drive-api.js',
  './drive-app.js',
  './qrcode.js',
  './config.js',
  './manifest.webmanifest',
  './offline.html',
  './icon-180.png',
  './icon-192.png',
  './icon-512.png',
  './icon-maskable-192.png',
  './icon-maskable-512.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE_NAME);
    // 1本でも失敗すると addAll 全体が落ちる。個別に入れて、取れなかったものは
    // 飛ばす（校内Wi-Fiが混んでいても導入できるようにするため）。
    await Promise.all(SHELL_ASSETS.map((u) =>
      cache.add(new Request(u, { cache: 'reload' }))
        .catch((err) => console.warn('[sw] precache skipped', u, err))));
    // ここでは skipWaiting しない。児童が書いている最中に突然切り替わらないよう、
    // 画面側で「さいしんに する」を押してもらってから切り替える（下の message）。
  })());
});

self.addEventListener('activate', (event) => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys
      .filter((k) => k.startsWith(CACHE_PREFIX) && k !== CACHE_NAME)
      .map((k) => caches.delete(k)));      // ← 自アプリ分だけ削除
    await self.clients.claim();
  })());
});

// 画面側で「さいしんに する」が押されたときだけ切り替える
self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'SKIP_WAITING') self.skipWaiting();
});

self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);

  // 同一オリジン（= GitHub Pages のシェル資産）以外は一切触らない。
  // Google APIへのリクエストは常にネットワーク直行。
  if (url.origin !== self.location.origin) return;
  if (event.request.method !== 'GET') return;

  // 画面遷移は network-first。更新をすぐ届け、圏外ならキャッシュ済みの
  // シェルを返し、それも無ければ offline.html を出す。
  if (event.request.mode === 'navigate') {
    event.respondWith((async () => {
      try {
        return await fetch(event.request);
      } catch (e) {
        return (await caches.match('./index.html'))
          || (await caches.match('./offline.html'))
          || Response.error();
      }
    })());
    return;
  }

  // シェル資産: キャッシュ優先 + バックグラウンド更新（stale-while-revalidate）
  event.respondWith(
    caches.match(event.request).then((cached) => {
      const network = fetch(event.request)
        .then((res) => {
          if (res && res.ok) {
            const clone = res.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
          }
          return res;
        })
        .catch(() => cached);
      return cached || network;
    })
  );
});
