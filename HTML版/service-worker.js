/**
 * Service Worker - 自動でつくる君 PWA
 *
 * キャッシュ戦略:
 * - install時: コアファイル（HTML/CSS/JS/フォント/テンプレート）をプリキャッシュ
 * - fetch時: Cache First → Network Fallback（CDNライブラリも含む）
 * - update時: 新バージョンをバックグラウンドインストール → 次回起動で反映
 */

const CACHE_VERSION = 'tsukurukun-v1';

// プリキャッシュ対象（アプリの核となるファイル）
const PRECACHE_URLS = [
  './',
  './index.html',
  './app.js',
  './style.css',
  './manifest.json',
  './icons/icon.svg',
  './fonts/NotoSerifJP.ttf',
  './template/文書送付書.doc.docx',
];

// CDNライブラリ（ランタイムキャッシュ）
const CDN_URLS = [
  'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js',
  'https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js',
  'https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/dist/pdf-lib.min.js',
  'https://cdn.jsdelivr.net/npm/@pdf-lib/fontkit@1.1.1/dist/fontkit.umd.min.js',
  'https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js',
];

// ===== Install: プリキャッシュ =====
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_VERSION).then(cache => {
      console.log('[SW] Pre-caching core assets');
      // コアファイルを先にキャッシュ（失敗してもインストールは続行）
      return cache.addAll(PRECACHE_URLS).then(() => {
        // CDNライブラリも可能な限りキャッシュ
        return Promise.allSettled(
          CDN_URLS.map(url => cache.add(url).catch(e => console.warn('[SW] CDN cache skip:', url, e)))
        );
      });
    }).then(() => self.skipWaiting())
  );
});

// ===== Activate: 古いキャッシュ削除 =====
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys.filter(k => k !== CACHE_VERSION).map(k => {
          console.log('[SW] Deleting old cache:', k);
          return caches.delete(k);
        })
      )
    ).then(() => self.clients.claim())
  );
});

// ===== Fetch: Cache First → Network Fallback =====
self.addEventListener('fetch', event => {
  const { request } = event;

  // POST等はスキップ
  if (request.method !== 'GET') return;

  // Tesseract workerやOCRデータはネットワーク優先（大きいので初回のみ）
  const url = new URL(request.url);
  if (url.pathname.includes('tesseract') && url.pathname.includes('worker')) {
    event.respondWith(networkFirstStrategy(request));
    return;
  }

  // それ以外は Cache First
  event.respondWith(cacheFirstStrategy(request));
});

async function cacheFirstStrategy(request) {
  const cached = await caches.match(request);
  if (cached) return cached;

  try {
    const response = await fetch(request);
    if (response.ok) {
      const cache = await caches.open(CACHE_VERSION);
      cache.put(request, response.clone());
    }
    return response;
  } catch (e) {
    // オフラインでキャッシュもない場合
    return new Response('オフラインです。ネットワーク接続を確認してください。', {
      status: 503,
      headers: { 'Content-Type': 'text/plain; charset=UTF-8' },
    });
  }
}

async function networkFirstStrategy(request) {
  try {
    const response = await fetch(request);
    if (response.ok) {
      const cache = await caches.open(CACHE_VERSION);
      cache.put(request, response.clone());
    }
    return response;
  } catch (e) {
    const cached = await caches.match(request);
    return cached || new Response('', { status: 503 });
  }
}
