// ============================================================
// ふりかえりジャーナル — 導入案内ページの Service Worker
// ------------------------------------------------------------
// このドメインが配るのは、アプリ本体ではなく「入れかたの案内」です。
// アプリ本体は、先生ごとに公開された script.google.com のウェブアプリで動きます。
//
// ここでキャッシュするのは案内ページ一式（HTML / CSS / manifest / アイコン）だけ。
// 外部への通信は一切キャッシュしません。
//
// ⚠️ 以前このドメインは Drive ネイティブ版のアプリ本体を配っていました。
//    その版を「アプリを入れる」でホーム画面に置いた端末は、次に開いたときに
//    この Service Worker へ入れ替わり、下の activate が古いキャッシュ
//    （rj-shell- で始まるもの）を消して、この案内ページに切り替わります。
// ============================================================

/* 【重要】キャッシュの掃除は、かならず自アプリのぶんだけに限る。
 *
 * 旧配信元の gigayama.github.io は数十本の学習アプリが同じドメインを共有していた。
 * いまは reflection-journal.giga-school.com を単独で使っているが、
 * 接頭辞で絞る形は変えない（同居する配置に戻したときに他アプリを巻き込むため）。
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
// ⚠️ この行は手で直さない。tools/build-sw.mjs が SHELL_ASSETS の中身から書き換える。
//    手書きだったころは上げるのが人の仕事で、2026-08-21 に12リポジトリで同時に
//    上げ忘れる事故が起きた。上げ忘れると古いシェルのキャッシュが掃除されず、
//    直した画面が児童の端末に届かない。
const APP_VERSION = 'v8d85144c'; /* __APP_VERSION__ */
const CACHE_NAME = CACHE_PREFIX + APP_VERSION;
const SHELL_ASSETS = [
  './',
  './index.html',
  './style.css',
  './offline.html',
  './privacy.html',
  './terms.html',
  './manifest.webmanifest',
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
