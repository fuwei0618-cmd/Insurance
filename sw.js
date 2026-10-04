/* 保險手冊 service worker：離線可用 */
const VERSION = '20261004200126';
const CORE = ['./', 'index.html', 'app.css', 'app.js', 'data/handbook.js', 'data/catalog.json', 'manifest.webmanifest', 'icons/icon-192.png', 'icons/icon-512.png', 'icons/icon-180.png'];
const IMAGES = ["img/0c35995183b6b6114f3520350d101a43.webp", "img/11ec8d486651ea1dc197185ee50e2172.webp", "img/20aa11ccb9ef9cf51abacfc1832bd97f.webp", "img/37b5ee2ed59740d656ca4117ccdddea4.webp", "img/69f67d9d0a2eb587e2b901a1cf198abe.webp", "img/7dde07323df64e40a8b55a1359c187cb.webp", "img/831123d32f9ffe2e94a97785e857c22b.webp", "img/833fda75e8c4a2d525f89bd19eb1d655.webp", "img/925140ea2425194f15b958720fec2a1f.webp", "img/a7e97a4288e127c3b9846dfa74917256.webp", "img/a8ea54a4973e00bc2fd00338228344bf.webp", "img/b479fbbc209af17bc98ac5d1387b6e85.webp", "img/b8ce3fdd8ce39b35e96dfb41e05b8085.webp", "img/d46635fa12024a55b9b5e698ceb5a157.webp", "img/d62b2dbdeea56d47ffc00343c04bfcc7.webp", "img/f5211cadb9749afa0d17f6f4c9cad0ac.webp", "img/f82c1b5f3bba79a8e91646856dccdd37.webp", "img/fa4d16fac2bc2648b11d76400dbdb3d5.webp", "img/fc6b093a5a67b0d832a8610d497dc6a8.webp"];
const CACHE = 'handbook-' + VERSION;
self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll([...CORE, ...IMAGES])).then(() => self.skipWaiting()));
});
self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k.startsWith('handbook-') && k !== CACHE).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  // 字型與 html2canvas：先用快取，沒有再上網並存起來
  if (/fonts\.(googleapis|gstatic)\.com|cdnjs\.cloudflare\.com/.test(url.host)) {
    e.respondWith(caches.open('handbook-ext').then(async (c) => {
      const hit = await c.match(req); if (hit) return hit;
      const res = await fetch(req); if (res.ok || res.type === 'opaque') c.put(req, res.clone()); return res;
    }));
    return;
  }
  if (url.origin !== location.origin) return;
  // 手冊內容：網路優先（拿到最新），離線時用快取
  e.respondWith(fetch(req).then((res) => {
    if (res.ok) { const copy = res.clone(); caches.open(CACHE).then((c) => c.put(req, copy)); }
    return res;
  }).catch(() => caches.match(req, { ignoreSearch: true }).then((r) => r || caches.match('index.html'))));
});
