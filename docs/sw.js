// ============================================================
// ふりかえりジャーナル — PWAシェル Service Worker
// ------------------------------------------------------------
// シェル資産（index.html / config.js / manifest / アイコン）のみをキャッシュする
// 最小構成。GAS 側（script.google.com / googleusercontent.com）は
// 認証・データがあるため【絶対にキャッシュしない】。
// ============================================================

const CACHE_NAME = 'rj-shell-v3';
const SHELL_ASSETS = [
  './',
  './index.html',
  './diag.html',
  './config.js',
  './manifest.webmanifest',
  './icon-180.png',
  './icon-192.png',
  './icon-512.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(SHELL_ASSETS)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);

  // 同一オリジン（= GitHub Pages のシェル資産）以外は一切触らない。
  // GAS へのリクエストは常にネットワーク直行。
  if (url.origin !== self.location.origin) return;
  if (event.request.method !== 'GET') return;

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
