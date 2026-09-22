// 列印工作台 service worker
// 改版時把 VERSION 加一，舊快取會在 activate 時清掉。
const VERSION = 'v3';
const CACHE = 'printlog-' + VERSION;

// 同源的核心檔案，裝不起來就沒有離線可言，所以是必要項目
const SHELL = [
  './',
  './index.html',
  './support.js',
  './manifest.json',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/icon-maskable-192.png',
  './icons/icon-maskable-512.png',
  './icons/apple-touch-icon.png',
  './icons/favicon-32.png'
];

// support.js 會自己去 unpkg 抓 React，離線時也要有。
// 這兩個網址寫死在 support.js 裡（REACT_URL / REACT_DOM_URL），改版要跟著對。
const CDN = [
  'https://unpkg.com/react@18.3.1/umd/react.production.min.js',
  'https://unpkg.com/react-dom@18.3.1/umd/react-dom.production.min.js'
];

self.addEventListener('install', event => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE);
    await cache.addAll(SHELL);
    // CDN 抓不到也不該擋住安裝，之後 fetch 時還會再補快取
    await Promise.all(CDN.map(u => cache.add(u).catch(() => {})));
    await self.skipWaiting();
  })());
});

self.addEventListener('activate', event => {
  event.waitUntil((async () => {
    const names = await caches.keys();
    await Promise.all(names.filter(n => n.startsWith('printlog-') && n !== CACHE).map(n => caches.delete(n)));
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET') return;

  const url = new URL(req.url);
  const sameOrigin = url.origin === self.location.origin;

  if (sameOrigin) {
    // 自己的檔案走 network-first：有網路就拿最新的，沒網路才回快取。
    // cache:'reload' 是必要的：GitHub Pages 回 max-age=600，不繞過瀏覽器的
    // HTTP 快取的話，network-first 會被架空，改版後最久要等 10 分鐘才看得到。
    event.respondWith((async () => {
      try {
        const res = await fetch(req, { cache: 'reload' });
        if (res && res.ok) {
          const cache = await caches.open(CACHE);
          cache.put(req, res.clone());
        }
        return res;
      } catch (e) {
        const hit = await caches.match(req);
        if (hit) return hit;
        // 換頁請求離線時退回首頁
        if (req.mode === 'navigate') {
          const shell = await caches.match('./index.html');
          if (shell) return shell;
        }
        throw e;
      }
    })());
    return;
  }

  // 外部資源（unpkg 的 React、Google Fonts、OCR 用的 tesseract.js 與語言模型）
  // 網址都帶版本，走 cache-first；第一次辨識抓過之後就能離線用
  event.respondWith((async () => {
    const hit = await caches.match(req);
    if (hit) return hit;
    const res = await fetch(req);
    if (res && (res.ok || res.type === 'opaque')) {
      const cache = await caches.open(CACHE);
      cache.put(req, res.clone());
    }
    return res;
  })());
});
