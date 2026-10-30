/* 保險手冊 service worker：網路優先，離線時用看過的內容 */
const CACHE='insurance-20261030133141';
const CORE=['./','index.html','manifest.webmanifest','icons/icon-192.png','icons/icon-512.png','icons/icon-180.png'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(CACHE).then(c=>c.addAll(CORE)).then(()=>self.skipWaiting()))});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>(k.startsWith('insurance-')||k.startsWith('handbook-'))&&k!==CACHE&&k!=='handbook-ext').map(k=>caches.delete(k)))).then(()=>self.clients.claim()))});
self.addEventListener('fetch',e=>{const r=e.request;if(r.method!=='GET')return;const u=new URL(r.url);
 if(/fonts\.(googleapis|gstatic)\.com/.test(u.host)){e.respondWith(caches.open('handbook-ext').then(async c=>{const h=await c.match(r);if(h)return h;const res=await fetch(r);if(res.ok||res.type==='opaque')c.put(r,res.clone());return res;}));return;}
 if(u.origin!==location.origin)return;
 e.respondWith(fetch(r).then(res=>{if(res.ok){const cp=res.clone();caches.open(CACHE).then(c=>c.put(r,cp));}return res;}).catch(()=>caches.match(r,{ignoreSearch:true}).then(x=>x||caches.match('index.html'))));});
