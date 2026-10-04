/* 客戶管理 service worker：只快取程式本身，不快取任何客戶資料 */
const CACHE = 'crm-app-v1';
const CORE = ['./', 'index.html', 'crm.css', 'crm.js', 'config.js', 'vendor/msal-browser.min.js', 'manifest.webmanifest', 'icons/icon-192.png', 'icons/icon-180.png'];
self.addEventListener('install', (e) => e.waitUntil(caches.open(CACHE).then((c) => c.addAll(CORE)).then(() => self.skipWaiting())));
self.addEventListener('activate', (e) => e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k.startsWith('crm-app') && k !== CACHE).map((k) => caches.delete(k)))).then(() => self.clients.claim())));
self.addEventListener('fetch', (e) => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin) return; // OneDrive／登入請求不經過快取
  e.respondWith(fetch(e.request).then((r) => { if (r.ok) { const c = r.clone(); caches.open(CACHE).then((x) => x.put(e.request, c)); } return r; })
    .catch(() => caches.match(e.request, { ignoreSearch: true }).then((r) => r || caches.match('index.html'))));
});
