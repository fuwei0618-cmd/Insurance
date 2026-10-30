/* 工作手冊：案件、掃描、遮罩、理賠參考、SOP、快查、範本
 * 資料加密後存在使用者自己的 OneDrive（應用程式 › Work › Cathay），裝置上只暫存加密檔。
 * 原圖永遠保留；遮罩是另外存的框，分享時才畫進新圖。
 */
(() => {
'use strict';
const CFG = window.CRM_CONFIG || {};
const FOLDER = (CFG.FOLDER || 'Cathay').replace(/^\/+|\/+$/g, '');
const $ = (id) => document.getElementById(id);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
const MOCK = (() => { try { if (/[?&]test\b/.test(location.search)) localStorage.setItem('crm:mock', '1'); if (/[?&]real\b/.test(location.search)) localStorage.removeItem('crm:mock'); return localStorage.getItem('crm:mock') === '1'; } catch (e) { return false; } })();
let toastT;
function toast(m, ms = 2600) { const t = $('toast'); t.textContent = m; t.hidden = false; clearTimeout(toastT); toastT = setTimeout(() => (t.hidden = true), ms); }

/* ================= 加密 ================= */
const enc = new TextEncoder(), dec = new TextDecoder();
const b64 = (u8) => { let s = ''; const a = new Uint8Array(u8); for (let i = 0; i < a.length; i += 0x8000) s += String.fromCharCode.apply(null, a.subarray(i, i + 0x8000)); return btoa(s); };
const unb64 = (s) => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));
async function deriveKey(pw, salt) {
  const base = await crypto.subtle.importKey('raw', enc.encode(pw), 'PBKDF2', false, ['deriveKey']);
  return crypto.subtle.deriveKey({ name: 'PBKDF2', salt, iterations: 310000, hash: 'SHA-256' }, base, { name: 'AES-GCM', length: 256 }, false, ['encrypt', 'decrypt']);
}
async function seal(key, bytes) { const iv = crypto.getRandomValues(new Uint8Array(12)); const ct = new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, key, bytes)); const out = new Uint8Array(12 + ct.length); out.set(iv); out.set(ct, 12); return out; }
async function unseal(key, bytes) { const u = new Uint8Array(bytes); return new Uint8Array(await crypto.subtle.decrypt({ name: 'AES-GCM', iv: u.slice(0, 12) }, key, u.slice(12))); }

/* ================= 本機暫存（IndexedDB，只放加密檔） ================= */
const Cache = {
  db: null,
  async init() {
    if (this.db) return this.db;
    this.db = await new Promise((ok, bad) => { const r = indexedDB.open('work-cache', 1); r.onupgradeneeded = () => r.result.createObjectStore('kv'); r.onsuccess = () => ok(r.result); r.onerror = () => bad(r.error); });
    return this.db;
  },
  async get(k) { try { const db = await this.init(); return await new Promise((ok) => { const r = db.transaction('kv').objectStore('kv').get(k); r.onsuccess = () => ok(r.result); r.onerror = () => ok(undefined); }); } catch (e) { return undefined; } },
  async set(k, v) { try { const db = await this.init(); await new Promise((ok) => { const t = db.transaction('kv', 'readwrite'); t.objectStore('kv').put(v, k); t.oncomplete = ok; t.onerror = ok; }); } catch (e) { } },
  async del(k) { try { const db = await this.init(); await new Promise((ok) => { const t = db.transaction('kv', 'readwrite'); t.objectStore('kv').delete(k); t.oncomplete = ok; t.onerror = ok; }); } catch (e) { } },
  async keys(prefix) { try { const db = await this.init(); const ks = await new Promise((ok) => { const r = db.transaction('kv').objectStore('kv').getAllKeys(); r.onsuccess = () => ok(r.result || []); r.onerror = () => ok([]); }); return ks.filter((k) => typeof k === 'string' && k.startsWith(prefix)); } catch (e) { return []; } },
};

/* ================= OneDrive（Microsoft Graph，應用程式資料夾） ================= */
const SCOPES = ['Files.ReadWrite.AppFolder', 'User.Read'];
const Drive = {
  msal: null, account: null,
  async init() {
    if (MOCK) { this.account = { username: '測試模式（資料只在這台裝置）' }; return; }
    this.msal = new msal.PublicClientApplication({
      auth: { clientId: CFG.CLIENT_ID, authority: 'https://login.microsoftonline.com/consumers', redirectUri: location.origin + location.pathname.replace(/index\.html$/, '') },
      cache: { cacheLocation: 'localStorage' },
    });
    await this.msal.initialize();
    const r = await this.msal.handleRedirectPromise().catch((e) => { console.warn(e); return null; });
    this.account = (r && r.account) || this.msal.getAllAccounts()[0] || null;
  },
  signIn() { return this.msal.loginRedirect({ scopes: SCOPES, prompt: 'select_account' }); },
  async signOut() { if (MOCK) { try { localStorage.removeItem('crm:mock'); } catch (e) { } location.reload(); return; } const a = this.account; this.account = null; await this.msal.logoutRedirect({ account: a }); },
  async token() {
    try { return (await this.msal.acquireTokenSilent({ scopes: SCOPES, account: this.account })).accessToken; }
    catch (e) { if (e instanceof msal.InteractionRequiredAuthError) { await this.msal.acquireTokenRedirect({ scopes: SCOPES, account: this.account }); } throw e; }
  },
  async api(path, opts = {}) { const t = await this.token(); return fetch('https://graph.microsoft.com/v1.0' + path, { ...opts, headers: { Authorization: 'Bearer ' + t, ...(opts.headers || {}) } }); },
  item: (p) => '/me/drive/special/approot:/' + (FOLDER + '/' + p).split('/').map(encodeURIComponent).join('/') + ':',
  async getMeta(p) {
    if (MOCK) return Mock.getMeta(p);
    const r = await this.api(this.item(p) + '?select=id,eTag,size,@microsoft.graph.downloadUrl');
    if (r.status === 404) return null;
    if (!r.ok) throw new Error('OneDrive ' + r.status);
    const j = await r.json(); return { eTag: j.eTag, size: j.size, url: j['@microsoft.graph.downloadUrl'] };
  },
  async download(p) {
    if (MOCK) return Mock.download(p);
    const m = await this.getMeta(p); if (!m) return null;
    const r = await fetch(m.url); if (!r.ok) throw new Error('下載失敗 ' + r.status);
    return { bytes: new Uint8Array(await r.arrayBuffer()), eTag: m.eTag };
  },
  async upload(p, bytes, eTag) { // eTag：只有在雲端版本沒被別台改過時才覆蓋
    if (MOCK) return Mock.upload(p, bytes, eTag);
    const headers = { 'Content-Type': 'application/octet-stream' };
    if (eTag) headers['If-Match'] = eTag;
    if (bytes.length <= 3.8 * 1024 * 1024) {
      const r = await this.api(this.item(p) + '/content', { method: 'PUT', body: bytes, headers });
      if (r.status === 412) { const e = new Error('conflict'); e.conflict = true; throw e; }
      if (!r.ok) throw new Error('上傳失敗 ' + r.status);
      return (await r.json()).eTag;
    }
    const s = await this.api(this.item(p) + '/createUploadSession', { method: 'POST', headers: { 'Content-Type': 'application/json', ...(eTag ? { 'If-Match': eTag } : {}) }, body: JSON.stringify({ item: { '@microsoft.graph.conflictBehavior': 'replace' } }) });
    if (s.status === 412) { const e = new Error('conflict'); e.conflict = true; throw e; }
    if (!s.ok) throw new Error('上傳失敗 ' + s.status);
    const url = (await s.json()).uploadUrl, CH = 3276800; let last;
    for (let i = 0; i < bytes.length; i += CH) {
      const part = bytes.subarray(i, Math.min(i + CH, bytes.length));
      last = await fetch(url, { method: 'PUT', body: part, headers: { 'Content-Range': `bytes ${i}-${i + part.length - 1}/${bytes.length}` } });
      if (!last.ok && last.status !== 202) throw new Error('上傳失敗 ' + last.status);
    }
    return (await last.json()).eTag;
  },
  async remove(p) { if (MOCK) return Mock.remove(p); const r = await this.api(this.item(p), { method: 'DELETE' }); return r.ok || r.status === 404; },
  async list(dir) { // 資料夾裡的檔名
    if (MOCK) return Mock.list(dir);
    const out = []; let url = this.item(dir) + '/children?select=name,size&$top=200';
    while (url) { const r = url.startsWith('http') ? await fetch(url, { headers: { Authorization: 'Bearer ' + (await this.token()) } }) : await this.api(url); if (r.status === 404) return out; if (!r.ok) throw new Error('OneDrive ' + r.status); const j = await r.json(); (j.value || []).forEach((x) => out.push(x.name)); url = j['@odata.nextLink'] || null; }
    return out;
  },
};
/* 測試模式：假雲端，資料只存在這台裝置的瀏覽器（網址加 ?test 開啟，?real 關閉） */
const Mock = {
  async getMeta(p) { const v = await Cache.get('mock:' + p); return v ? { eTag: v.eTag, size: v.bytes.length } : null; },
  async download(p) { const v = await Cache.get('mock:' + p); return v ? { bytes: v.bytes, eTag: v.eTag } : null; },
  async upload(p, bytes, eTag) { const v = await Cache.get('mock:' + p); if (eTag && v && v.eTag !== eTag) { const e = new Error('conflict'); e.conflict = true; throw e; } const t = uid(); await Cache.set('mock:' + p, { bytes, eTag: t }); return t; },
  async remove(p) { await Cache.del('mock:' + p); return true; },
  async list(dir) { const pre = 'mock:' + dir + '/'; return (await Cache.keys(pre)).map((k) => k.slice(pre.length)).filter((n) => !n.includes('/')); },
};

/* ================= 資料庫（記憶體中明文，存檔時加密） ================= */
const COLLS = ['cases', 'notes', 'files', 'clients', 'settings', 'calls', 'rels', 'gevents', 'gifts'];
const emptyDb = () => Object.fromEntries([['v', 1], ...COLLS.map((c) => [c, {}])]);
const PILL_LABEL = { synced: '✓ 已儲存', saving: '儲存中…', pending: '儲存中…', offline: '⚠ 離線，已存在這台裝置', editing: '編輯中…（離開欄位就會儲存）' };
function savePill(s) {
  let el = document.getElementById('savePill');
  if (!el) { el = document.createElement('div'); el.id = 'savePill'; el.setAttribute('role', 'status'); el.setAttribute('aria-live', 'polite'); document.body.appendChild(el); }
  if (!Vault.key) { el.hidden = true; return; }
  el.hidden = false; el.dataset.s = s; el.textContent = PILL_LABEL[s] || '';
  clearTimeout(savePill.t); if (s === 'synced') savePill.t = setTimeout(() => { el.dataset.s = 'idle'; }, 2500);
}
const SYNC_LABEL = { synced: '已同步 OneDrive', saving: '同步中…', pending: '待同步…', offline: '離線中，連上網路會自動同步', idle: '—' };
const Vault = {
  key: null, salt: null, db: null, eTag: null, dirty: false, saving: false, state: 'idle',
  async load(pw) { // 'ok'，密碼錯誤時丟出錯誤
    let raw = null, fromCache = false;
    try { raw = await Drive.download('data.enc'); } catch (e) { const c = await Cache.get('data'); if (c) { raw = c; fromCache = true; } else throw e; }
    if (!raw) { const c = await Cache.get('data'); if (c) { raw = c; fromCache = true; } else return 'new'; }
    const env = JSON.parse(dec.decode(raw.bytes));
    const salt = unb64(env.salt), key = await deriveKey(pw, salt);
    let plain;
    try { plain = await unseal(key, unb64(env.data)); } catch (e) { const err = new Error('密碼不正確'); err.badPassword = true; throw err; }
    this.key = key; this.salt = salt; this.db = { ...emptyDb(), ...JSON.parse(dec.decode(plain)) }; this.eTag = raw.eTag;
    if (!fromCache) { await Cache.set('data', raw); this.raw = raw; } // 留著做今天的備份
    const pend = await Cache.get('pending');
    if (pend) { try { const local = JSON.parse(dec.decode(await unseal(key, pend))); this.db = merge(this.db, local); this.markDirty(); } catch (e) { } }
    this.setState(fromCache ? 'offline' : 'synced');
    return 'ok';
  },
  async create(pw) {
    this.salt = crypto.getRandomValues(new Uint8Array(16)); this.key = await deriveKey(pw, this.salt); this.db = emptyDb(); this.eTag = null;
    this.dirty = true; await this.save(true);
  },
  async envelope() { const data = await seal(this.key, enc.encode(JSON.stringify(this.db))); return enc.encode(JSON.stringify({ v: 1, salt: b64(this.salt), data: b64(data), at: new Date().toISOString() })); },
  markDirty() { this.dirty = true; this.setState('pending'); clearTimeout(this.t); this.t = setTimeout(() => this.save(), 1200); const k = this.key; seal(k, enc.encode(JSON.stringify(this.db))).then((b) => { if (this.key === k) Cache.set('pending', b); }); },
  async save(force) {
    if (!this.key || (this.saving && !force)) return;
    if (!this.dirty && !force) return;
    this.saving = true; this.setState('saving');
    try {
      for (let tries = 0; tries < 4; tries++) {
        const bytes = await this.envelope();
        try {
          this.eTag = await Drive.upload('data.enc', bytes, this.eTag);
          this.dirty = false; await Cache.set('data', { bytes, eTag: this.eTag }); await Cache.del('pending');
          this.setState('synced'); break;
        } catch (e) {
          if (!e.conflict) throw e;
          await this.pull(true); // 別台裝置改過：合併後再存
        }
      }
    } catch (e) { console.warn(e); this.setState('offline'); clearTimeout(this.retry); this.retry = setTimeout(() => this.save(), 30000); }
    finally { this.saving = false; if (this.dirty && this.state === 'synced') this.save(); }
  },
  async pull(force) { // 讀回別台裝置的更新
    if (!this.key) return;
    try {
      if (!force) { const m = await Drive.getMeta('data.enc'); if (!m || m.eTag === this.eTag) return; }
      const remote = await Drive.download('data.enc'); if (!remote) return;
      const env = JSON.parse(dec.decode(remote.bytes));
      const other = JSON.parse(dec.decode(await unseal(this.key, unb64(env.data))));
      this.db = merge(other, this.db); this.eTag = remote.eTag;
      if (!this.dirty) { await Cache.set('data', remote); this.setState('synced'); }
      App.render();
    } catch (e) { if (force) throw e; }
  },
  setState(s) { this.state = s; document.querySelectorAll('.sync').forEach((el) => { el.textContent = SYNC_LABEL[s] || s; el.dataset.s = s; }); savePill(s); },
  lock() { this.key = null; this.db = null; this.salt = null; this.eTag = null; this.dirty = false; },
  put(coll, rec) { const now = new Date().toISOString(); if (!rec.id) rec.id = uid(); if (!rec.createdAt) rec.createdAt = now; rec.updatedAt = now; this.db[coll][rec.id] = rec; this.markDirty(); return rec; },
  del(coll, id) { if (!this.db[coll][id]) return; this.db[coll][id] = { id, deleted: true, updatedAt: new Date().toISOString() }; this.markDirty(); },
  all(coll) { return Object.values(this.db?.[coll] || {}).filter((r) => !r.deleted); },
  get(coll, id) { const r = this.db?.[coll]?.[id]; return r && !r.deleted ? r : null; },
};
function merge(a, b) { // 依每筆資料的 updatedAt 取較新的
  const out = emptyDb();
  COLLS.forEach((c) => { const A = a[c] || {}, B = b[c] || {}; new Set([...Object.keys(A), ...Object.keys(B)]).forEach((id) => { const x = A[id], y = B[id]; out[c][id] = !x ? y : !y ? x : ((x.updatedAt || '') >= (y.updatedAt || '') ? x : y); }); });
  return out;
}

/* ================= 圖片（加密後存 OneDrive files/） ================= */
const fpath = (id) => 'files/' + id + '.enc';
const Files = {
  mem: new Map(),
  async add(blob, name) {
    const id = uid(), bytes = new Uint8Array(await blob.arrayBuffer());
    const sealed = await seal(Vault.key, bytes);
    await Cache.set('file:' + id, sealed);
    Vault.put('files', { id, type: blob.type || 'image/jpeg', size: bytes.length, name });
    this.mem.set(id, URL.createObjectURL(blob));
    Drive.upload(fpath(id), sealed).then(() => Cache.set('file-up:' + id, 1)).catch(() => toast('圖片已存在這台裝置，連上網路後會再上傳'));
    return id;
  },
  async url(id) {
    if (this.mem.has(id)) return this.mem.get(id);
    let sealed = await Cache.get('file:' + id);
    const meta = Vault.get('files', id) || { id };
    if (!sealed) { const r = await Drive.download(fpath(id)); if (!r) return null; sealed = r.bytes; Cache.set('file:' + id, sealed); Cache.set('file-up:' + id, 1); }
    const u = URL.createObjectURL(new Blob([await unseal(Vault.key, sealed)], { type: meta.type || 'image/jpeg' }));
    this.mem.set(id, u); return u;
  },
  // 刪除＝先放「最近刪除」30 天，之後才真的從 OneDrive 移除；ctx 記住原本在哪，才放得回去
  remove(id, ctx) { const r = Vault.get('files', id); if (!r) return; r.trashedAt = new Date().toISOString(); if (ctx) r.trashCtx = ctx; Vault.put('files', r); },
  trashed() { return Vault.all('files').filter((f) => f.trashedAt).sort((a, b) => b.trashedAt.localeCompare(a.trashedAt)); },
  restore(id) {
    const r = Vault.get('files', id); if (!r) return false; const x = r.trashCtx; const target = x && Vault.get(x.coll, x.recId);
    if (!target) return false;
    target.docs = target.docs || []; target.docs.push(x.doc); if (x.doc.folder && target.folders && !target.folders.includes(x.doc.folder)) target.folders.push(x.doc.folder);
    delete r.trashedAt; delete r.trashCtx; Vault.put('files', r); Vault.put(x.coll, target); return true;
  },
  async purge() { // 超過 30 天的才真的刪
    const cut = new Date(Date.now() - 30 * 864e5).toISOString();
    for (const f of this.trashed().filter((f) => f.trashedAt < cut)) { try { if (await Drive.remove(fpath(f.id))) { Vault.del('files', f.id); await Cache.del('file:' + f.id); } } catch (e) { return; } }
  },
  async retryUploads() {
    for (const f of Vault.all('files')) {
      if (await Cache.get('file-up:' + f.id)) continue;
      const s = await Cache.get('file:' + f.id); if (!s) continue;
      try { const m = await Drive.getMeta(fpath(f.id)); if (!m) await Drive.upload(fpath(f.id), s); await Cache.set('file-up:' + f.id, 1); } catch (e) { return; }
    }
  },
  clear() { this.mem.forEach((u) => URL.revokeObjectURL(u)); this.mem.clear(); },
};
const imgCache = new Map();
function loadImg(src) { return new Promise((r) => { const i = new Image(); i.onload = () => r(i); i.onerror = () => r(null); i.src = src; }); }
function getImg(fileId) { if (!imgCache.has(fileId)) imgCache.set(fileId, Files.url(fileId).then((u) => (u ? loadImg(u) : null)).catch(() => null)); return imgCache.get(fileId); }

/* ================= 選項 ================= */
const ICON = {
  gift: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="8" width="18" height="4" rx="1"/><path d="M5 12v8h14v-8M12 8v12M12 8c-1.5-3-5-4-5-1.5S9.5 8 12 8zm0 0c1.5-3 5-4 5-1.5S14.5 8 12 8z"/></svg>',
  scan: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4 8V6a2 2 0 0 1 2-2h2M16 4h2a2 2 0 0 1 2 2v2M20 16v2a2 2 0 0 1-2 2h-2M8 20H6a2 2 0 0 1-2-2v-2M4 12h16"/></svg>',
  bolt: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M13 3 4 14h7l-1 7 9-11h-7l1-7z"/></svg>',
  folder: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4 6a2 2 0 0 1 2-2h4l2 2h6a2 2 0 0 1 2 2v9a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2z"/></svg>',
  book: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4 5a2 2 0 0 1 2-2h13v16H6a2 2 0 0 0-2 2zM4 21V5M8 7h7"/></svg>',
  list: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M9 6h11M9 12h11M9 18h11M4 6l1 1 2-2M4 12l1 1 2-2M4 18l1 1 2-2"/></svg>',
  phone: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M5 4h4l2 5-2.5 1.5a11 11 0 0 0 5 5L15 13l5 2v4a2 2 0 0 1-2 2A16 16 0 0 1 3 6a2 2 0 0 1 2-2"/></svg>',
  user: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="8" r="4"/><path d="M4 21a8 8 0 0 1 16 0"/></svg>',
  msg: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4 5h16v11H9l-5 4z"/></svg>',
};
const NAV = [['quick', '快查', 'bolt'], ['clients', '客戶', 'user'], ['cases', '案件', 'folder'], ['calls', '服務件', 'phone'], ['refs', '理賠參考', 'book'], ['sop', 'SOP', 'list'], ['tpl', '範本', 'msg'], ['gifts', '送禮', 'gift']];
const DEFAULT_CATS = ['理賠', '建議書', '新契約', '保全', '保費', '保校', '服務', '其他'];
const CATS = [...DEFAULT_CATS]; // 可以在 App 裡自己改，存在設定裡
// 分類分兩層：大類（照公司拜訪事件）› 小類（案件、SOP、範本共用）
const DEFAULT_GROUPS = [['關係／觀念建立', []], ['需求分析', ['保單健檢']], ['建議書／促約', ['建議書', '新契約']], ['客戶服務', ['理賠', '保全', '保費', '保校', '服務']], ['其他', ['其他']]];
let GROUPS = DEFAULT_GROUPS.map(([name, cats]) => ({ name, cats: [...cats] }));
function syncCats() {
  const st = Vault.get('settings', 'catgroups')?.list;
  let groups;
  if (st?.length) groups = st.map((g) => ({ name: g.name, cats: [...(g.cats || [])] }));
  else { // 還沒設定過兩層：用舊的類別清單套進預設大類
    const old = Vault.get('settings', 'cats')?.list, flat = old?.length ? old : DEFAULT_CATS;
    groups = DEFAULT_GROUPS.map(([name, cats]) => ({ name, cats: cats.filter((c) => flat.includes(c) || !DEFAULT_CATS.includes(c)) }));
    const placed = new Set(groups.flatMap((g) => g.cats));
    flat.filter((c) => !placed.has(c)).forEach((c) => groups[groups.length - 1].cats.push(c));
  }
  GROUPS = groups; const next = groups.flatMap((g) => g.cats);
  if (next.join('|') !== CATS.join('|')) CATS.splice(0, CATS.length, ...next);
}
const groupOf = (cat) => GROUPS.find((g) => g.cats.includes(cat))?.name || '';
const catOptions = (sel, withNone) => (withNone ? `<option value=""${!sel ? ' selected' : ''}>${esc(withNone)}</option>` : '') + GROUPS.filter((g) => g.cats.length).map((g) => `<optgroup label="${esc(g.name)}">${g.cats.map((k) => `<option${k === sel ? ' selected' : ''}>${esc(k)}</option>`).join('')}</optgroup>`).join('');
function openCatEditor() {
  let gs = GROUPS.map((g) => ({ old: g.name, name: g.name, cats: g.cats.map((c) => ({ old: c, name: c })) }));
  const used = (n) => cases().filter((c) => c.cat === n).length + notes().filter((x) => (x.kind === 'sop' || x.kind === 'tpl') && catsOf(x).includes(n)).length;
  const draw = () => {
    $('ceList').innerHTML = gs.map((g, gi) => `<div class="cegroup"><div class="cerow cehead"><input data-gn="${gi}" value="${esc(g.name)}" aria-label="大類名稱"><span class="sub">大類</span><button class="btn" data-gu="${gi}" aria-label="上移" ${gi ? '' : 'disabled'}>↑</button><button class="btn" data-gd="${gi}" aria-label="下移" ${gi < gs.length - 1 ? '' : 'disabled'}>↓</button><button class="btn" data-gx="${gi}" aria-label="刪除大類">✕</button></div>
      ${g.cats.map((r, i) => `<div class="cerow cesub"><input data-cn="${gi}:${i}" value="${esc(r.name)}" aria-label="小類名稱"><span class="sub">${r.old ? used(r.old) + ' 筆' : '新的'}</span><select data-cm="${gi}:${i}" aria-label="移到大類">${gs.map((x, xi) => `<option value="${xi}"${xi === gi ? ' selected' : ''}>${esc(x.name || '（未命名）')}</option>`).join('')}</select><button class="btn" data-cu="${gi}:${i}" aria-label="上移" ${i ? '' : 'disabled'}>↑</button><button class="btn" data-cd="${gi}:${i}" aria-label="下移" ${i < g.cats.length - 1 ? '' : 'disabled'}>↓</button><button class="btn" data-cx="${gi}:${i}" aria-label="刪除">✕</button></div>`).join('')}
      <div class="cerow cesub"><input data-cadd="${gi}" placeholder="＋ 在「${esc(g.name)}」新增小類，打完按 Enter" aria-label="新增小類"></div></div>`).join('');
  };
  ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>編輯分類</h2><button class="btn" id="ceX">取消</button><button class="btn pri" id="ceSave">儲存</button></div>
  <div class="ovbody" style="flex-direction:column;max-width:640px;margin:0 auto;width:100%">
   <div class="note">大類照公司的拜訪事件分，小類是案件、SOP、範本共用的分類。改小類名字，用到它的案件、SOP、範本會一起改過去；有東西在用的小類不能刪。<br>「理賠」「新契約」「建議書」有專屬功能，建議保留原名。</div>
   <div id="ceList"></div>
   <div class="actions"><input class="search" id="ceNewG" placeholder="新大類名稱" style="flex:1"><button class="btn" id="ceAddG">＋ 新增大類</button></div></div></div>`;
  draw();
  const at = (k) => k.split(':').map(Number);
  $('ceList').oninput = (e) => { const d = e.target.dataset; if (d.gn != null) gs[+d.gn].name = e.target.value; if (d.cn) { const [g, i] = at(d.cn); gs[g].cats[i].name = e.target.value; } };
  $('ceList').onchange = (e) => { const d = e.target.dataset; if (d.cm) { const [g, i] = at(d.cm); const [r] = gs[g].cats.splice(i, 1); gs[+e.target.value].cats.push(r); draw(); } };
  $('ceList').onkeydown = (e) => {
    const d = e.target.dataset; if (d.cadd == null || e.key !== 'Enter') return; e.preventDefault();
    const v = e.target.value.trim(); if (!v) return; if (gs.some((g) => g.cats.some((r) => r.name.trim() === v))) { toast('已經有這個小類'); return; }
    gs[+d.cadd].cats.push({ old: '', name: v }); draw(); $('ceList').querySelector(`[data-cadd="${d.cadd}"]`)?.focus();
  };
  $('ceList').onclick = (e) => {
    const b = e.target.closest('button'); if (!b) return; const d = b.dataset;
    if (d.gu) { const i = +d.gu; [gs[i - 1], gs[i]] = [gs[i], gs[i - 1]]; }
    if (d.gd) { const i = +d.gd; [gs[i + 1], gs[i]] = [gs[i], gs[i + 1]]; }
    if (d.gx) { const g = gs[+d.gx]; if (g.cats.length) { toast(`「${g.name}」裡還有小類，先把它們移到別的大類或刪掉`); return; } gs.splice(+d.gx, 1); }
    if (d.cu) { const [g, i] = at(d.cu); const a = gs[g].cats; [a[i - 1], a[i]] = [a[i], a[i - 1]]; }
    if (d.cd) { const [g, i] = at(d.cd); const a = gs[g].cats; [a[i + 1], a[i]] = [a[i], a[i + 1]]; }
    if (d.cx) { const [g, i] = at(d.cx); const r = gs[g].cats[i]; if (r.old && used(r.old)) { toast(`「${r.old}」還有 ${used(r.old)} 筆在用，先把它們改到別的小類`); return; } gs[g].cats.splice(i, 1); }
    draw();
  };
  const addG = () => { const v = $('ceNewG').value.trim(); if (!v) return; if (gs.some((g) => g.name.trim() === v)) { toast('已經有這個大類'); return; } gs.push({ old: '', name: v, cats: [] }); $('ceNewG').value = ''; draw(); };
  $('ceAddG').onclick = addG; $('ceNewG').onkeydown = (e) => { if (e.key === 'Enter') addG(); };
  $('ceX').onclick = () => (ov.innerHTML = '');
  $('ceSave').onclick = () => {
    const gnames = gs.map((g) => g.name.trim()), all = gs.flatMap((g) => g.cats), names = all.map((r) => r.name.trim());
    if (gnames.some((n) => !n) || names.some((n) => !n)) { toast('名稱不能空白'); return; }
    if (new Set(gnames).size !== gnames.length) { toast('有重複的大類名稱'); return; }
    if (new Set(names).size !== names.length) { toast('有重複的小類名稱'); return; }
    if (!names.length) { toast('至少要留一個小類'); return; }
    const m = new Map(all.filter((r) => r.old && r.old !== r.name.trim()).map((r) => [r.old, r.name.trim()]));
    if (m.size) {
      cases().forEach((c) => { if (m.has(c.cat)) { c.cat = m.get(c.cat); saveCase(c); } });
      notes().forEach((n) => { if (n.kind !== 'sop' && n.kind !== 'tpl') return; const k = catOf(n), al = n.alsoCats || []; if (m.has(k) || al.some((x) => m.has(x))) { if (m.has(k)) n.cat = m.get(k); n.alsoCats = al.map((x) => m.get(x) || x); if (n.kind === 'sop') n.v2 = true; Vault.put('notes', n); } });
      const st = Vault.get('settings', 'steps'); if (st?.map) { m.forEach((b, a) => { if (st.map[a]) { st.map[b] = st.map[a]; delete st.map[a]; } else if (STEPS[a] && !st.map[b]) st.map[b] = STEPS[a]; }); Vault.put('settings', st); }
      if (m.has(S.cat)) S.cat = m.get(S.cat); if (m.has(S.sc)) S.sc = m.get(S.sc);
    }
    Vault.put('settings', { id: 'catgroups', list: gs.map((g) => ({ name: g.name.trim(), cats: g.cats.map((r) => r.name.trim()) })) });
    Vault.put('settings', { id: 'cats', list: names }); // 給還沒更新的裝置用
    syncCats(); if (S.sg !== '全部' && S.sg !== '通用' && !GROUPS.some((g) => g.name === S.sg)) S.sg = '全部';
    ov.innerHTML = ''; App.render(); toast('分類已更新');
  };
}
// SOP、範本的分類與標籤（舊的 SOP 還沒重新存過時，從原本的「壽險·保費」這類分類推出來，不改動原資料）
const LEGACY_T = '商品·T字障';
function mapGroup(g) { if (!g) return ''; if (g === LEGACY_T) return CATS.includes('建議書') ? '建議書' : ''; const k = String(g).replace(/^壽險·/, ''); return CATS.includes(k) ? k : ''; }
const catOf = (n) => (n.kind === 'sop' && !n.v2 ? mapGroup(n.group) : n.cat || '');
const catsOf = (n) => [...new Set([catOf(n), ...(n.alsoCats || [])].filter(Boolean))]; // SOP、範本可以同時放在好幾個小類
const catsLabel = (n) => catsOf(n).join('、') || '通用';
const kindsOf = (n) => n.kinds || [];
const toolsOf = (n) => n.tools || (n.kind === 'sop' && !n.v2 && n.group === LEGACY_T ? ['T字障'] : []);
const TAGS = { kinds: { label: '壽險險種', def: ['壽險', '醫療', '儲蓄', '意外', '年金', '投資型', '旅平'], of: kindsOf }, pkinds: { label: '產險險種', def: ['車險', '住火', '旅綜', '寵物險'], of: (n) => n.pkinds || [] }, mkinds: { label: '整合行銷項目', def: ['房貸', '證券', '信用卡', '開戶', '推薦連結'], of: (n) => n.mkinds || [] }, tools: { label: '講解工具', def: ['T字障', '四象限', '帆船理論', '金三角', '28000天'], of: toolsOf } };
const sopsOf = (t) => t.sopIds || (t.sopId ? [t.sopId] : []); // 一則範本可以掛在好幾個 SOP 底下
const chOf = (n) => n.channels || (n.group ? [n.group] : []); // 範本可以同時用在 LINE、簡訊、信件…
const chPills = (n) => chOf(n).map((g) => `<span class="pill p-go">${esc(g)}</span>`).join('');
const tagList = (key) => { const l = Vault.get('settings', 'tag-' + key)?.list; return l?.length ? l : TAGS[key].def; };
function openTagEditor(key) {
  const T = TAGS[key]; let rows = tagList(key).map((n) => ({ old: n, name: n }));
  const used = (t) => notes().filter((n) => T.of(n).includes(t)).length;
  const draw = () => { $('ceList').innerHTML = rows.map((r, i) => `<div class="cerow"><input data-ci="${i}" value="${esc(r.name)}" aria-label="名稱"><span class="sub">${r.old ? used(r.old) + ' 筆' : '新的'}</span><button class="btn" data-cu="${i}" aria-label="上移" ${i ? '' : 'disabled'}>↑</button><button class="btn" data-cd="${i}" aria-label="下移" ${i < rows.length - 1 ? '' : 'disabled'}>↓</button><button class="btn" data-cx="${i}" aria-label="刪除">✕</button></div>`).join(''); };
  ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>編輯${T.label}</h2><button class="btn" id="ceX">取消</button><button class="btn pri" id="ceSave">儲存</button></div>
  <div class="ovbody" style="flex-direction:column;max-width:560px;margin:0 auto;width:100%"><div class="note">改名字，貼了這個標籤的筆記會一起改過去；還有筆記在用的不能刪。</div><div id="ceList"></div>
   <div class="actions"><input class="search" id="ceNew" placeholder="新增${T.label}" style="flex:1"><button class="btn" id="ceAdd">＋ 新增</button></div></div></div>`;
  draw();
  $('ceList').oninput = (e) => { const i = e.target.dataset.ci; if (i != null) rows[+i].name = e.target.value; };
  $('ceList').onclick = (e) => {
    const b = e.target.closest('button'); if (!b) return; const d = b.dataset;
    if (d.cu) { const i = +d.cu; [rows[i - 1], rows[i]] = [rows[i], rows[i - 1]]; }
    if (d.cd) { const i = +d.cd; [rows[i + 1], rows[i]] = [rows[i], rows[i + 1]]; }
    if (d.cx) { const r = rows[+d.cx]; if (r.old && used(r.old)) { toast(`還有 ${used(r.old)} 筆貼了「${r.old}」，先把標籤拿掉`); return; } rows.splice(+d.cx, 1); }
    draw();
  };
  const add = () => { const v = $('ceNew').value.trim(); if (!v) return; if (rows.some((r) => r.name.trim() === v)) { toast('已經有了'); return; } rows.push({ old: '', name: v }); $('ceNew').value = ''; draw(); };
  $('ceAdd').onclick = add; $('ceNew').onkeydown = (e) => { if (e.key === 'Enter') add(); };
  $('ceX').onclick = () => (ov.innerHTML = '');
  $('ceSave').onclick = () => {
    const names = rows.map((r) => r.name.trim());
    if (names.some((n) => !n)) { toast('名稱不能空白'); return; }
    if (new Set(names).size !== names.length) { toast('有重複的名稱'); return; }
    const m = new Map(rows.filter((r) => r.old && r.old !== r.name.trim()).map((r) => [r.old, r.name.trim()]));
    if (m.size) notes().forEach((n) => { const cur = T.of(n); if (cur.some((t) => m.has(t))) { n[key] = cur.map((t) => m.get(t) || t); Vault.put('notes', n); } });
    const fk = { kinds: 'sk', pkinds: 'spk', mkinds: 'smk', tools: 'st' }[key]; if (m.has(S[fk])) S[fk] = m.get(S[fk]);
    Vault.put('settings', { id: 'tag-' + key, list: names }); ov.innerHTML = ''; App.render(); toast(T.label + '已更新');
  };
}
const STATUS = ['進行中', '待回覆', '已結案'];
const DEFAULT_ITEMS = { 壽險: ['住院醫療', '實支實付', '住院日額', '手術', '意外醫療', '意外失能', '骨折', '重大傷病', '癌症', '長照失能', '門診'], 產險: ['車險', '住火', '旅綜', '寵物險'] , 整合行銷: ['房貸', '證券', '信用卡', '開戶', '推薦連結'] };
const LINES = ['壽險', '產險', '整合行銷'];
const MK_WORDS = ['房貸', '證券', '信用卡', '開戶', '推薦連結', '銀行'];
const P_WORDS = ['車險', '住火', '旅綜', '旅平', '寵物', '頭家壓箱寶', '產險'];
const isPItem = (x) => P_WORDS.some((w) => String(x).includes(w));
const DEFAULT_CAUSE = { '疾病': ['感染症（病毒、細菌）', '消化系統', '呼吸系統', '心血管', '腦血管／神經', '骨科／肌肉骨骼', '泌尿／腎臟', '婦科／產科', '內分泌／代謝', '腫瘤', '眼耳鼻喉', '皮膚', '牙科', '身心科', '其他疾病'], '意外': ['交通事故', '跌倒／滑倒', '運動傷害', '職業傷害', '燒燙傷', '割傷／刺傷', '撞擊／夾傷', '動物咬抓傷', '其他意外'] };
const ITEMS = { 壽險: [...DEFAULT_ITEMS.壽險], 產險: [...DEFAULT_ITEMS.產險], 整合行銷: [...DEFAULT_ITEMS.整合行銷] }, CAUSE = Object.fromEntries(Object.entries(DEFAULT_CAUSE).map(([k, v]) => [k, [...v]]));
function syncRefLists() { // 險種項目、原因可以自己改，存在設定裡
  const set = (arr, l) => { const next = l?.length ? l : null; if (next && next.join('|') !== arr.join('|')) arr.splice(0, arr.length, ...next); };
  const old = Vault.get('settings', 'items')?.list; // 舊版只有一份：產險字樣的歸產險，其他歸壽險
  LINES.forEach((ln) => { let l = Vault.get('settings', 'items-' + ln)?.list; if (!l && old && ln !== '整合行銷') l = ln === '產險' ? [...new Set([...old.filter(isPItem).map((x) => x.replace(/^產險-/, '')), ...DEFAULT_ITEMS.產險])] : old.filter((x) => !isPItem(x)); set(ITEMS[ln], l || DEFAULT_ITEMS[ln]); });
  Object.keys(CAUSE).forEach((k) => set(CAUSE[k], Vault.get('settings', 'cause-' + k)?.list || DEFAULT_CAUSE[k]));
}
// 通用的清單編輯器（改名會一起改到用到它的資料；有在用的不能刪）
function openListEditor({ title, note, list, used, rename, save }) {
  let rows = list.map((n) => ({ old: n, name: n }));
  const draw = () => { $('ceList').innerHTML = rows.map((r, i) => `<div class="cerow"><input data-ci="${i}" value="${esc(r.name)}" aria-label="名稱"><span class="sub">${r.old ? used(r.old) + ' 筆' : '新的'}</span><button class="btn" data-cu="${i}" aria-label="上移" ${i ? '' : 'disabled'}>↑</button><button class="btn" data-cd="${i}" aria-label="下移" ${i < rows.length - 1 ? '' : 'disabled'}>↓</button><button class="btn" data-cx="${i}" aria-label="刪除">✕</button></div>`).join(''); };
  ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>${esc(title)}</h2><button class="btn" id="ceX">取消</button><button class="btn pri" id="ceSave">儲存</button></div>
  <div class="ovbody" style="flex-direction:column;max-width:560px;margin:0 auto;width:100%"><div class="note">${note || '改名字，用到它的資料會一起改過去；還有資料在用的不能刪。'}</div><div id="ceList"></div>
   <div class="actions"><input class="search" id="ceNew" placeholder="新增一項，打完按 Enter" style="flex:1"><button class="btn" id="ceAdd">＋ 新增</button></div></div></div>`;
  draw();
  $('ceList').oninput = (e) => { const i = e.target.dataset.ci; if (i != null) rows[+i].name = e.target.value; };
  $('ceList').onclick = (e) => {
    const b = e.target.closest('button'); if (!b) return; const d = b.dataset;
    if (d.cu) { const i = +d.cu; [rows[i - 1], rows[i]] = [rows[i], rows[i - 1]]; }
    if (d.cd) { const i = +d.cd; [rows[i + 1], rows[i]] = [rows[i], rows[i + 1]]; }
    if (d.cx) { const r = rows[+d.cx]; if (r.old && used(r.old)) { toast(`還有 ${used(r.old)} 筆在用「${r.old}」，先把它們改掉`); return; } rows.splice(+d.cx, 1); }
    draw();
  };
  const add = () => { const v = $('ceNew').value.trim(); if (!v) return; if (rows.some((r) => r.name.trim() === v)) { toast('已經有了'); return; } rows.push({ old: '', name: v }); $('ceNew').value = ''; draw(); };
  $('ceAdd').onclick = add; $('ceNew').onkeydown = (e) => { if (e.key === 'Enter') add(); };
  $('ceX').onclick = () => (ov.innerHTML = '');
  $('ceSave').onclick = () => {
    const names = rows.map((r) => r.name.trim());
    if (!names.length) { toast('至少要留一項'); return; }
    if (names.some((n) => !n)) { toast('名稱不能空白'); return; }
    if (new Set(names).size !== names.length) { toast('有重複的名稱'); return; }
    const m = new Map(rows.filter((r) => r.old && r.old !== r.name.trim()).map((r) => [r.old, r.name.trim()]));
    if (m.size) rename(m);
    save(names); ov.innerHTML = ''; App.render(); toast(title.replace(/^編輯/, '') + '已更新');
  };
}
// 壽險／產險是每筆資料的屬性（不是分類）：案件分壽險、產險；SOP、範本多一個「通用」
const NLINES = ['壽險', '產險', '整合行銷', '通用'];
const lineN = (n) => n.line || ((n.pkinds || []).length && !(n.kinds || []).length ? '產險' : (n.mkinds || []).length && !(n.kinds || []).length ? '整合行銷' : '通用');
const TAG_LINE = { kinds: '壽險', tools: '壽險', pkinds: '產險', mkinds: '整合行銷' }; // 每排標籤屬於哪一邊
const tagShown = (key, ln) => !ln || ln === '通用' || TAG_LINE[key] === ln;
const lineOk = (n, want) => !want || lineN(n) === want || lineN(n) === '通用';
const linePill = (ln) => (ln && ln !== '通用' ? `<span class="pill ${ln === '產險' ? 'p-wait' : ln === '整合行銷' ? 'p-mk' : 'p-go'}">${ln}</span>` : '');
const lineSeg = (cur, attr, opts) => `<div class="stseg lineseg" role="group" aria-label="壽險或產險">${opts.map((ln) => `<button type="button" class="${cur === ln ? 'on st-go' : ''}" data-${attr}="${ln}">${ln}</button>`).join('')}</div>`;
const lineOf = (c) => c.line || ((c.items || []).some((x) => ITEMS.產險.includes(x) || isPItem(x)) ? '產險' : (c.items || []).some((x) => ITEMS.整合行銷.includes(x)) ? '整合行銷' : '壽險');
function editItems(ln) {
  openListEditor({ title: `編輯${ln}的險種項目`, list: ITEMS[ln], used: (n) => cases().filter((c) => (c.items || []).includes(n)).length,
    rename: (m) => cases().forEach((c) => { if ((c.items || []).some((x) => m.has(x))) { c.items = c.items.map((x) => m.get(x) || x); saveCase(c); } }),
    save: (names) => { Vault.put('settings', { id: 'items-' + ln, list: names }); syncRefLists(); } });
}
function editCause(type) {
  openListEditor({ title: `編輯「${type}」的原因`, list: CAUSE[type], used: (n) => cases().filter((c) => c.causeType === type && c.causeDetail === n).length,
    rename: (m) => cases().forEach((c) => { if (c.causeType === type && m.has(c.causeDetail)) { c.causeDetail = m.get(c.causeDetail); saveCase(c); } }),
    save: (names) => { Vault.put('settings', { id: 'cause-' + type, list: names }); syncRefLists(); } });
}
const TPLGROUPS = ['LINE', '簡訊', '信件', '通話', '其他'];
const linkify = (html) => html.replace(/(https?:\/\/[^\s<]+)/g, (u) => `<a href="${u}" target="_blank" rel="noopener">${u.length > 60 ? u.slice(0, 57) + '…' : u}</a>`);

/* ================= 狀態與畫面 ================= */
const S = { railMini: (() => { try { return localStorage.getItem('crm-railmini') === '1'; } catch (e) { return false; } })(), ginvOpen: false, gview: 'table', ktable: false, kprog: '全部', kmark: '全部', ksort: 'no', kview: 'prog', kedit: false, kfromTable: false, kar: '全部', klv: '全部', gev: null, gtype: '全部', gst: '全部', gedit: false, view: 'cases', showHealth: false, tplGroup: '全部', tplCat: '全部', ksel: null, kst: '全部', kres: '', script: null, ccat: '全部', csel: null, showIds: false, sel: null, refSel: null, cat: '全部', q: '', refItem: '全部', refCause: '全部', cg: '全部', tg: '全部', sg: '全部', sc: '全部', sk: '', st: '', showDetail: false };
const cases = () => Vault.all('cases');
const localDay = (d = new Date()) => { const t = new Date(d); return `${t.getFullYear()}-${String(t.getMonth() + 1).padStart(2, '0')}-${String(t.getDate()).padStart(2, '0')}`; };
const caseDate = (c) => c.date || (c.createdAt ? localDay(c.createdAt) : '');
const byDate = (a, b) => caseDate(b).localeCompare(caseDate(a)) || (b.updatedAt || '').localeCompare(a.updatedAt || '');
const fmtDay = (d) => (d ? d.replace(/-/g, '/') : '');
const notes = () => Vault.all('notes');
const cur = () => Vault.get('cases', S.sel);
const clients = () => Vault.all('clients');
const curClient = () => Vault.get('clients', S.csel);
const saveCase = (c) => Vault.put('cases', c);
const app = $('app');
const pillCls = (s) => (s === '已結案' ? 'st-done' : s === '待回覆' ? 'st-wait' : 'st-go');
const itemCls = (s) => (s === '已結案' ? ' is-done' : s === '待回覆' ? ' is-wait' : ' is-go');
const syncHTML = (cls = '') => `<span class="sync ${cls}" data-s="${Vault.state}">${SYNC_LABEL[Vault.state] || ''}</span>`;

const App = {
  render() {
    if (App._defer) { clearTimeout(App._dt); App._dt = setTimeout(() => App.render(), 0); return; } // 改完一格跳到下一格時，等游標到新格子再重畫
    if (!Vault.db) return;
    syncCats(); syncRefLists(); if (S.cat !== '全部' && !CATS.includes(S.cat)) S.cat = '全部'; if (S.cg !== '全部' && !GROUPS.some((g) => g.name === S.cg)) S.cg = '全部';
    const v = S.view;
    app.className = 'app' + ((v === 'cases' || v === 'refs' || v === 'clients' || (v === 'calls' && !S.ktable)) ? '' : ' single') + (S.showDetail ? ' showdetail' : '') + (MOCK ? ' test' : '') + (S.railMini ? ' railmini' : '');
    let h = MOCK ? '<div class="testbar">測試模式：資料只存在這台裝置，不會上傳 OneDrive</div>' : '';
    h += railHTML();
    if (v === 'clients') h += clientListHTML() + '<section class="detail split">' + clientDetailHTML() + '</section>';
    else if (v === 'cases') h += caseListHTML() + '<section class="detail split">' + caseDetailHTML() + '</section>';
    else if (v === 'calls' && S.ktable) h += '<section class="wide">' + callTableHTML() + '</section>';
    else if (v === 'calls') h += callListHTML() + '<section class="detail split">' + callDetailHTML() + '</section>';
    else if (v === 'refs') h += refListHTML() + '<section class="detail split">' + refDetailHTML() + '</section>';
    else if (v === 'sop') h += '<section class="wide">' + notesHTML('sop') + '</section>';
    else if (v === 'quick') h += '<section class="wide">' + quickHTML() + '</section>';
    else if (v === 'tpl') h += '<section class="wide">' + notesHTML('tpl') + '</section>';
    else if (v === 'gifts') h += '<section class="wide">' + giftsHTML() + '</section>';
    h += tabsHTML();
    const ae = document.activeElement, keep = ae?.id, pos = ae?.selectionStart, pos2 = ae?.selectionEnd; // 重畫後把游標放回原本那格，鍵盤才不會收起來
    const typing = ae && app.contains(ae) && (ae.tagName === 'TEXTAREA' || (ae.tagName === 'INPUT' && !['checkbox', 'radio', 'file', 'button'].includes(ae.type)));
    const keepSel = typing ? (keep ? '#' + CSS.escape(keep) : [...ae.attributes].filter((x) => x.name.startsWith('data-')).map((x) => `[${x.name}="${CSS.escape(x.value)}"]`).join('')) : '';
    const sc = [...app.querySelectorAll('section, .list')].map((el) => [el.className, el.scrollTop]), wy = window.scrollY; // 重畫後回到原本的捲動位置
    app.innerHTML = h;
    const pk = [S.view, S.sel, S.csel, S.ksel, S.rsel, S.showDetail].join('|'), same = App._pk === pk; App._pk = pk;
    if (same) { const els = [...app.querySelectorAll('section, .list')]; sc.forEach(([cn, t]) => { const el = els.find((e) => e.className === cn); if (el && t) el.scrollTop = t; }); if (wy) window.scrollTo(0, wy); }
    if (keepSel && (same || keep === 'q')) { const el = app.querySelector(keepSel); if (el) { el.focus({ preventScroll: true }); try { el.setSelectionRange(pos, pos2 ?? pos); } catch (e) { } } }
    drawThumbs(); fitNotes();
    if (v === 'gifts' && S.gedit) giftSortWire();
    if (v === 'clients' && curClient()) drawEgo(curClient());
  },
};
function railHTML() {
  return `<nav class="rail" aria-label="主選單"><div class="brand"><i></i><span class="bt">工作手冊</span><button class="railtog" data-act="railtog" title="${S.railMini ? '展開選單' : '收合選單'}" aria-label="${S.railMini ? '展開選單' : '收合選單'}">${S.railMini ? '»' : '«'}</button></div>
  <button class="scanbtn" data-act="scan" title="掃描">${ICON.scan}<span class="nl">掃描</span></button>
  ${NAV.map(([k, l, i]) => `<button class="nav${S.view === k ? ' on' : ''}" data-nav="${k}" title="${l}">${ICON[i]}<span class="nl">${l}</span></button>`).join('')}
  <div class="foot">${syncHTML()}<div style="margin-top:8px">${esc(Drive.account?.username || '')}</div></div>
  <button class="nav" data-act="lock">鎖定</button></nav>`;
}
function tabsHTML() {
  const t = [NAV[0], NAV[1], NAV[2], ['scan', '掃描', 'scan'], NAV[3], ['refs', '參考', 'book'], NAV[5]];
  return `<nav class="tabs" aria-label="分頁">${t.map(([k, l, i]) => k === 'scan' ? `<button class="tab scan" data-act="scan"><span class="c">${ICON.scan}</span>${l}</button>` : `<button class="tab${S.view === k ? ' on' : ''}" data-nav="${k}">${ICON[i]}${l}</button>`).join('')}</nav>`;
}
function caseSummary(c) { const a = [c.cat]; if ((c.people || []).length > 1) a.push(c.people.length + ' 人'); if (c.items?.length) a.push(c.items.slice(0, 2).join('、')); if (c.causeType) a.push(c.causeType + (c.causeDetail ? '›' + c.causeDetail : '')); return a.join(' · '); }
function filteredCases() {
  const q = S.q.trim();
  const g = S.cg !== '全部' && GROUPS.find((x) => x.name === S.cg);
  return cases().filter((c) => (!S.cline || lineOf(c) === S.cline) && (S.cat !== '全部' ? c.cat === S.cat : !g || g.cats.includes(c.cat)) && (!q || [c.client, c.title, c.notes, c.diag, c.next, ...(c.people || []).map((p) => Vault.get('clients', p.clientId)?.name || '')].join(' ').includes(q))).sort(byDate);
}
const lineFilter = (key) => `<div class="chips linefilter">${['全部', ...LINES].map((x) => `<button class="chip${(S[key] || '全部') === x ? ' on' : ''}" data-lf="${key}:${x}">${x}</button>`).join('')}</div>`;
function caseChipsHTML() { // 先選大類，再出現它的小類；只列有案件的
  const all = cases().filter((c) => !S.cline || lineOf(c) === S.cline), n = (cats) => all.filter((c) => cats.includes(c.cat)).length;
  const gs = GROUPS.filter((g) => n(g.cats)), g = GROUPS.find((x) => x.name === S.cg);
  const subs = g ? g.cats.filter((k) => n([k])) : [];
  return `${lineFilter('cline')}<div class="chips">${[['全部', all.length], ...gs.map((x) => [x.name, n(x.cats)])].map(([x, k]) => `<button class="chip${S.cg === x ? ' on' : ''}" data-cg="${esc(x)}">${esc(x)} <span class="cnt">${k}</span></button>`).join('')}<button class="chip chipedit" data-act="editcats" aria-label="編輯分類">⚙</button></div>
  ${subs.length > 1 ? `<div class="chips subchips">${[['全部', n(g.cats)], ...subs.map((k) => [k, n([k])])].map(([x, k]) => `<button class="chip${S.cat === x ? ' on' : ''}" data-cat="${esc(x)}">${esc(x)} <span class="cnt">${k}</span></button>`).join('')}</div>` : ''}`;
}
function caseListHTML() {
  const list = filteredCases();
  return `<section class="col"><div class="colhead"><h1>案件 <button class="btn" data-act="newcase">＋ 新增</button></h1>
  <div class="mobsync">${syncHTML()}</div>
  <input class="search" id="q" placeholder="搜尋客戶、案件、備註" value="${esc(S.q)}" autocomplete="off">
  ${caseChipsHTML()}</div>
  <div class="list">${list.length ? list.map((c) => `<button class="item${c.id === S.sel ? ' on' : ''}${itemCls(c.status)}" data-sel="${c.id}" title="${esc((c.client || '') + ' · ' + (c.title || ''))}"><span class="t"><span>${esc(c.client || '未填客戶')} · ${esc(c.title || '未命名')}</span><span class="pill ${pillCls(c.status)}">${c.status}</span></span><span class="s">${fmtDay(caseDate(c))} · ${esc(caseSummary(c))}${latestStep(c) ? ' · 最新：' + esc(latestStep(c).title) : ''}</span></button>`).join('') : `<div class="empty">${cases().length ? '沒有符合的案件。' : '還沒有案件。按「＋ 新增」，或按「掃描」直接拍文件建案。'}</div>`}</div></section>`;
}
// 某些類別的案件一打開就有固定資料夾（只在還沒有任何資料夾時加，刪掉了不會再自己長回來）
const DEFAULT_FOLDERS = { 理賠: ['理賠文件', '理賠明細'], 新契約: ['準備資料'], 建議書: ['準備資料'] };
function ensureFolders(c) {
  if (c.foldersCat && c.foldersCat !== c.cat) { // 換了類別：原本的預設資料夾還是空的，就換成新類別的預設
    const old = DEFAULT_FOLDERS[c.foldersCat] || [], fs = c.folders || [];
    if (fs.length && fs.every((f) => old.includes(f) && !(c.docs || []).some((d) => d.folder === f))) { c.folders = [...(DEFAULT_FOLDERS[c.cat] || [])]; }
    c.foldersCat = c.cat; Vault.put('cases', c);
  }
  const def = DEFAULT_FOLDERS[c.cat]; if (!def) return;
  if (c.foldersInit && !c.foldersV2) { // 上一版不管有沒有自己設過都加了：已經有自己資料夾的，把加進去、還是空的預設資料夾拿掉
    c.foldersV2 = true; const own = (c.folders || []).filter((f) => !def.includes(f));
    if (own.length) c.folders = c.folders.filter((f) => !def.includes(f) || c.docs.some((d) => d.folder === f));
    Vault.put('cases', c); return;
  }
  if (c.foldersInit) return;
  c.foldersInit = c.foldersV2 = true; c.foldersCat = c.cat;
  if (!(c.folders || []).length) c.folders = [...def]; // 只有完全沒設過資料夾的才自動建
  Vault.put('cases', c);
}
function caseDetailHTML() {
  const c = cur();
  if (!c) return '<div class="empty">從左邊選一個案件，或按「掃描」直接拍文件建案。</div>';
  ensureFolders(c);
  const isClaim = c.cat === '理賠';
  const cost = +c.cost || 0, ratio = +c.ratio || 0, auto = cost && ratio ? Math.round(cost * ratio / 100) : 0;
  return `<button class="back" data-act="back">‹ 案件列表</button>
  <div class="dhead"><input class="title" id="f_title" data-f="title" value="${esc(c.title)}" placeholder="案件名稱" aria-label="案件名稱"><span class="pill ${pillCls(c.status)}">${c.status}</span></div>
  ${c.convertedAt ? `<div class="note">由建議書轉成新契約 · ${esc(c.convertedAt)}</div>` : ''}
  ${c.cat === '建議書' && c.outcome !== '成交' ? `<div class="sec banner"><div><div style="font-weight:500">成交了嗎？</div><div class="sub">會另外開一件新契約，帶入客戶和要保人／被保險人。建議書的準備筆記留在這裡，不會混進新契約。</div></div><button class="btn pri" data-act="convert">成交，建立新契約</button></div>` : ''}
  ${c.cat === '建議書' && c.outcome === '成交' ? `<div class="note">已成交，轉成新契約${(c.links || []).map((id) => Vault.get('cases', id)).filter((x) => x && x.fromProposal === c.id).map((x) => ` <button class="linkbtn" data-gocase="${x.id}">${esc(x.title || '新契約')} ›</button>`).join('')}</div>` : ''}
  ${c.cat === '新契約' && c.status !== '已結案' && c.clientId && !casePeople(c).some((p) => (Vault.get('clients', p.clientId)?.disclosures || []).some((r) => r.caseId === c.id)) ? `<div class="sec banner warnbanner"><div><div style="font-weight:500">記得記錄告知資料</div><div class="sub">送件時填的財務、身高體重存起來，下次就不用再問。</div></div><button class="btn" data-act="casedisc">記錄告知</button></div>` : ''}
  ${c.fromProposal && Vault.get('cases', c.fromProposal) ? `<div class="note">來自建議書 <button class="linkbtn" data-gocase="${c.fromProposal}">${esc(Vault.get('cases', c.fromProposal).title || '建議書')} ›</button></div>` : ''}
  <div class="sec"><h2>基本資料</h2><div class="grid2">
   <label class="f">客戶${c.clientId && Vault.get('clients', c.clientId) ? ` <button class="clientchip" data-goclient="${c.clientId}">👤 看客戶資料 ›</button>` : ''}<input id="f_client" data-f="client" list="clientNames" autocomplete="off" value="${esc(c.client)}" placeholder="打名字，會自動帶出或建立客戶"><datalist id="clientNames">${clients().map((x) => `<option value="${esc(x.name)}">`).join('')}</datalist></label>
   <label class="f">日期<input type="date" id="f_date" data-f="date" value="${esc(caseDate(c))}"></label>
   <div class="f">類別<div class="catpick">${(() => { const g = GROUPS.find((x) => x.cats.includes(c.cat)); const gs = GROUPS.filter((x) => x.cats.length); return `<select data-catgroup aria-label="大類">${g ? '' : '<option selected>（請選大類）</option>'}${gs.map((x) => `<option${x === g ? ' selected' : ''}>${esc(x.name)}</option>`).join('')}</select><select id="f_cat" data-f="cat" aria-label="小類">${g ? '' : `<option selected>${esc(c.cat)}</option>`}${(g ? g.cats : []).map((k) => `<option${k === c.cat ? ' selected' : ''}>${esc(k)}</option>`).join('')}</select>`; })()}</div></div>
   <div class="f">業務別${lineSeg(lineOf(c), 'line', LINES)}</div>
   <div class="f">狀態<div class="stseg" role="group" aria-label="狀態">${STATUS.map((x) => `<button class="${x === c.status ? 'on ' + pillCls(x) : ''}" data-setst="${x}" aria-pressed="${x === c.status}">${x}</button>`).join('')}</div></div>
   <label class="f">下一步<input id="f_next" data-f="next" value="${esc(c.next)}" placeholder="例如：補診斷書正本"></label></div></div>
  ${lineOf(c) === '整合行銷' ? `<div class="sec"><h2>整合行銷項目</h2><div class="tags">${[...ITEMS.整合行銷, ...(c.items || []).filter((x) => !ITEMS.整合行銷.includes(x))].map((x) => `<button class="tag${(c.items || []).includes(x) ? ' on' : ''}" data-item="${esc(x)}">${esc(x)}</button>`).join('')}<button class="tag tagedit" data-edititems="整合行銷">⚙ 編輯</button></div></div>` : ''}
  ${peopleHTML(c)}
  ${timelineHTML(c)}
  ${isClaim ? `<div class="sec"><h2>理賠分類</h2>
   <div class="sub">${lineOf(c)}險種項目（可複選，上面「壽險／產險」切換）</div><div class="tags">${(() => { const L = ITEMS[lineOf(c)]; return [...L, ...c.items.filter((x) => !L.includes(x))]; })().map((x) => `<button class="tag${c.items.includes(x) ? ' on' : ''}" data-item="${esc(x)}">${esc(x)}</button>`).join('')}<button class="tag tagedit" data-edititems="${lineOf(c)}">⚙ 編輯</button></div>
   ${lineOf(c) === '產險' ? '' : `<div class="sub">原因</div><div class="tags">${Object.keys(CAUSE).map((x) => `<button class="tag${c.causeType === x ? ' on' : ''}" data-ctype="${x}">${x}</button>`).join('')}</div>
   ${c.causeType && CAUSE[c.causeType] ? `<div class="tags">${[...CAUSE[c.causeType], ...(c.causeDetail && !CAUSE[c.causeType].includes(c.causeDetail) ? [c.causeDetail] : [])].map((x) => `<button class="tag${c.causeDetail === x ? ' on' : ''}" data-cdet="${esc(x)}">${esc(x)}</button>`).join('')}<button class="tag tagedit" data-editcause="${esc(c.causeType)}">⚙ 編輯</button></div>` : ''}`}
   <label class="f">診斷／受傷情況<input id="f_diag" data-f="diag" value="${esc(c.diag)}" placeholder="例如：腳踝扭傷併骨裂，打石膏"></label></div>
  <div class="sec"><h2>理賠計算</h2><div class="calc">
   <label class="f">實際花費<input id="f_cost" inputmode="numeric" data-f="cost" value="${esc(c.cost)}" placeholder="76600"></label><span class="op">×</span>
   <label class="f">給付比例 %<input id="f_ratio" inputmode="numeric" data-f="ratio" value="${esc(c.ratio)}" placeholder="65"></label><span class="op">=</span>
   <span class="res">${auto ? auto.toLocaleString() : '—'}</span></div>
   <label class="f">實際理賠金額（有不同再填）<input id="f_payout" inputmode="numeric" data-f="payout" value="${esc(c.payout)}"></label></div>` : ''}
  ${caseDocsHTML(c)}
  ${linksHTML(c)}
  ${caseTplHTML(c)}
  ${['新契約', '建議書'].includes(c.cat) ? ncCaseHTML(c) : ''}
  <div class="sec"><h2>備註</h2><textarea id="f_notes" class="grow" data-f="notes" rows="4" aria-label="備註">${esc(c.notes)}</textarea></div>
  <div class="actions"><button class="btn pri savebtn" data-act="savenow">儲存</button>${isClaim ? `<button class="btn${c.isRef ? ' on' : ''}" data-act="ref">${c.isRef ? '✓ 已在理賠參考' : '做成理賠參考'}</button>` : ''}${c.clientId && ['新契約', '建議書', '保全', '產險'].includes(c.cat) ? '<button class="btn" data-act="casedisc">記錄財務／體況告知</button>' : ''}<button class="btn danger" data-act="delcase">刪除案件</button></div>`;
}
/* ================= 電訪 ================= */
const calls = () => Vault.all('calls');
const curCall = () => Vault.get('calls', S.ksel);
const CALL_ST = ['待撥', '再聯絡', '已約訪', '結案'];
const RESULTS = [['未接', '再聯絡'], ['已接・聊過', '再聯絡'], ['約訪成功', '已約訪'], ['請我再打', '再聯絡'], ['拒絕', '結案'], ['空號／錯號', '結案']];
const noCmp = (a, b) => { const x = parseInt(a.no, 10), y = parseInt(b.no, 10); return (isNaN(x) ? 1e9 : x) - (isNaN(y) ? 1e9 : y) || (a.name || '').localeCompare(b.name || '', 'zh-Hant'); };
const NO_ANSWER = ['未接', '空號／錯號'];
const PROGS = ['未打', '未接', '已接通', '已排訪', '已見面', '不經營'];
function callProg(k) { // 這個人走到哪一關
  const logs = k.logs || [];
  if (k.met) return '已見面';
  if (k.visitDate || k.status === '已約訪') return '已排訪';
  if (k.status === '結案') return '不經營';
  if (logs.some((l) => !NO_ANSWER.includes(l.result))) return '已接通';
  return logs.length ? '未接' : '未打';
}
function callSlots(k) { // 三個時段：上午、下午、晚上，各打過沒、接了沒
  const m = { 上午: '', 下午: '', 晚上: '' };
  (k.logs || []).filter((l) => l.at).forEach((l) => { const h = new Date(l.at).getHours(), s = h < 12 ? '上午' : h < 18 ? '下午' : '晚上'; if (!NO_ANSWER.includes(l.result)) m[s] = 'ok'; else if (!m[s]) m[s] = 'no'; });
  return m;
}
function callGift(k) { const gs = Vault.all('gifts').filter((g) => g.who?.type === 'call' && g.who?.id === k.id); if (gs.length) return gs.every((g) => g.status === '已送') ? '已送' : '預備'; return k.giftSkip ? '不送' : ''; }
const CALL_FLOW = [
  ['準備（做一次）', ['匯入名單：選 PDF，檢查前八位欄位對不對，按匯入。', '送禮專區建好「年曆」活動和品項。']],
  ['每天打電話', ['開「總表」，進度點「未打」，再點一個地區，一次打完一區。', '撥號前看一眼「打電話前的功課」：住址觀察、保單概況。', '照話術六步打：開場、年曆、家戶、職業、聊天、收尾。']],
  ['掛電話後三十秒內記完', ['記錄結果：未接、已接・聊過、約訪成功…', '分級：A、B、C-1、C-2、C-3。', '勾確認：服務人員、家戶、職業類別。', '年曆：要的按「＋ 預備贈送」選品項和送法；不要的按「不送年曆」。', '有加 LINE 就勾；約好了就填約訪日期。']],
  ['沒接的人', ['總表點「未接」，看三時段哪個還是灰的，換那個時段再打。', '三個時段都紅了：傳簡訊或 LINE，年曆改成寄送。']],
  ['排約訪、送年曆', ['年曆到貨後，總表點「已排訪」加地區，同一區排同一天，半天三到四家。', '見到面：「已見面」打勾，送禮那筆按成「已送」。', 'A、B 的人開一個建議書案件，接上準備清單和流程提醒。', '有人跟你買了，再按「轉成客戶」，才進客戶區。']],
  ['每週看一次', ['總表最上面的進度條：打過、接通、要年曆、加 LINE、排訪、見面、年曆已送，看這週跑到哪。']],
];
function openCallFlow() {
  ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>服務件流程</h2><button class="btn pri" id="cfX">關閉</button></div>
  <div class="ovbody" style="flex-direction:column;max-width:760px;margin:0 auto;width:100%">
   <div class="note">分級：<b>A</b> VIP 高度優先・<b>B</b> 有預算有促約機會・<b>C-1</b> 有關鍵人物預算不夠・<b>C-2</b> 關係好觀念還沒到・<b>C-3</b> 排斥忙沒需求，用寄的</div>
   ${CALL_FLOW.map(([h, l], i) => `<div class="sec"><h2>${i + 1}. ${h}</h2><ol class="cflow">${l.map((x) => `<li>${esc(x)}</li>`).join('')}</ol></div>`).join('')}</div></div>`;
  $('cfX').onclick = () => (ov.innerHTML = '');
}
const normAddr = (x) => (x || '').replace(/\s+/g, '').replace(/台/g, '臺').replace(/[０-９]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));
const hhKey = (k) => (k.hh ? 'h:' + k.hh : normAddr(k.addr) ? 'a:' + normAddr(k.addr) : 'n:' + k.id);
function callHouseholds(all) { const m = new Map(); all.forEach((k) => { const key = hhKey(k); m.has(key) ? m.get(key).push(k) : m.set(key, [k]); }); return m; }
const hhPending = (g) => g.length > 1 && !g.every((k) => k.hhOk);
const addrParts = (k) => { const m = normAddr(k.addr).match(/^(\d{3,6})?(.*)$/); return [m[1] ? parseInt(m[1].slice(0, 3), 10) : 1e6, m[2] || '']; };
const addrCmp = (a, b) => { const [za, ra] = addrParts(a), [zb, rb] = addrParts(b); return (!ra) - (!rb) || za - zb || ra.localeCompare(rb, 'zh-Hant', { numeric: true }) || noCmp(a, b); };
function linkHH(a, b) { // 把 b 拉進 a 那一戶，不用決定戶長
  const all = calls(), nm = a.hh || b.hh || '戶' + (a.no || a.name || '');
  [a.hh, b.hh].filter((x) => x && x !== nm).forEach((old) => all.filter((k) => k.hh === old).forEach((k) => { k.hh = nm; Vault.put('calls', k); }));
  [a, b].forEach((k) => { k.hh = nm; k.hhOk = true; Vault.put('calls', k); }); return nm;
}
const phoneDigits = (x) => (x || '').replace(/\D/g, '');
function hhCandidates(k) { const ad = normAddr(k.addr), hp = phoneDigits(k.home), [z] = addrParts(k), sn = (k.name || '').slice(0, 1);
  return calls().filter((o) => o.id !== k.id && !(k.hh && o.hh === k.hh)).map((o) => [o, ad && normAddr(o.addr) === ad ? '同地址' : hp.length >= 7 && phoneDigits(o.home) === hp ? '同住家電話' : sn && o.name?.[0] === sn && z < 1e6 && addrParts(o)[0] === z ? '同姓・同郵遞區號' : '']).filter(([, w]) => w); }
function callHHSectionHTML(k) {
  const mem = k.hh ? calls().filter((o) => o.hh === k.hh && o.id !== k.id).sort(noCmp) : [], cand = hhCandidates(k);
  const row = (o, btn) => `<div class="khrow"><span><span class="kno">${esc(o.no || '')}</span><b>${esc(o.name || '未填')}</b> <span class="sub">${esc(o.addr || '')}</span></span>${btn}</div>`;
  return `<div class="sec"><h2>家戶 <span class="sub">${k.hh ? esc(k.hh) + `・${mem.length + 1} 人` : '還沒拉家戶'}</span></h2>
   ${mem.map((o) => row(o, `<button class="btn sm" data-khout="${o.id}">移出</button>`)).join('')}
   ${cand.length ? `<div class="sub" style="margin:6px 0 4px">可能同一戶，電話中問到就按「＋ 同一戶」：</div>${cand.map(([o, w]) => row(o, `<span class="pill">${w}</span><button class="btn sm" data-khadd="${o.id}">＋ 同一戶</button>`)).join('')}` : mem.length ? '' : '<div class="sub">沒有找到同地址、同住家電話的人。</div>'}
   ${k.hh ? `<div class="actions" style="margin:6px 0 0"><button class="linkbtn" data-khout="${k.id}">把自己移出這一戶</button></div>` : ''}</div>`;
}
const KSKIP = ['同事的件', '保經', '其他'];
const skipWhy = (k) => k.otherWhy || '同事的件';
const KMARKS = ['全部', '已加 LINE', '未加 LINE', '退件', '不追蹤', '家戶待確認'];
function callGiftSheetHTML(list, ev, evs) { // 服務件專區裡的送禮表：一列一位，一欄一個品項，填數字就自動記到送禮專區
  if (!ev) return '<div class="empty">還沒有送禮活動。先到「送禮」按「＋ 新活動」建一個，例如「2027 年曆」，再回來這裡填。</div>';
  const items = giftItemsSorted(ev.items || []), cats = []; items.forEach((it) => { const c = it.cat || '未分類'; const l = cats[cats.length - 1]; l && l.c === c ? l.n++ : cats.push({ c, n: 1 }); });
  const gmap = new Map(giftsOf(ev.id).filter((g) => g.who?.type === 'call').map((g) => [g.who.id, g])), tot = (id) => list.reduce((a, k) => a + (gmap.get(k.id)?.items?.[id] || 0), 0);
  const sent = list.filter((k) => gmap.get(k.id)?.status === '已送').length, want = list.filter((k) => gmap.has(k.id)).length;
  return `${evs.length > 1 ? `<div class="chips"><span class="sub" style="align-self:center">活動</span>${evs.map((x) => `<button class="chip${x.id === ev.id ? ' on' : ''}" data-kgev="${x.id}">${esc(x.name || '未命名')}</button>`).join('')}</div>` : ''}
  <div class="note">「${esc(ev.name || '')}」：在品項格子填數量就會自動記到送禮專區；不要的勾「不送」。上面的篩選、排序一樣能用。</div>
  <div class="ktablewrap"><table class="ktable gsheet"><thead>
   <tr><th rowspan="2" class="gstick">序號・姓名</th><th rowspan="2">地區</th><th rowspan="2">分級</th>${cats.map((c) => `<th colspan="${c.n}" class="gcath">${esc(c.c)}</th>`).join('')}<th rowspan="2">送法</th><th rowspan="2">已送</th><th rowspan="2">不送</th></tr>
   <tr>${items.map((it) => `<th class="gith">${giftThumb(it)}<div>${esc(it.name || it.code || '品項')}</div></th>`).join('')}</tr></thead><tbody>
   ${list.map((k) => { const g = gmap.get(k.id); return `<tr class="${g?.status === '已送' ? 'gdone' : ''}${k.other || k.giftSkip ? ' kother' : ''}"><td class="gstick"><span class="kno">${esc(k.no || '')}</span><button class="linkbtn" data-ksel="${k.id}">${esc(k.name || '未填')}</button>${k.other ? ` <span class="pill p-off">${esc(skipWhy(k))}</span>` : ''}${k.mailBack ? ' <span class="pill p-off">退件</span>' : ''}</td><td>${esc(k.area || '')}</td><td>${k.level ? `<b class="lvl lv-${esc(k.level)}">${esc(k.level)}</b>` : ''}</td>
     ${items.map((it) => `<td><input class="gcell gn" inputmode="numeric" data-kgq="${k.id}:${it.id}" value="${g?.items?.[it.id] || ''}"></td>`).join('')}
     <td><select class="gcell" data-kgm="${k.id}">${['', '當面', '寄送', '放管理室'].map((m) => `<option${(g?.method || '') === m ? ' selected' : ''}>${m}</option>`).join('')}</select></td>
     <td style="text-align:center"><input type="checkbox" class="gchk" data-kgd="${k.id}"${g?.status === '已送' ? ' checked' : ''}></td>
     <td style="text-align:center"><input type="checkbox" data-kte="${k.id}" data-kfield="giftSkip"${k.giftSkip ? ' checked' : ''}></td></tr>`; }).join('')}
   <tr class="gtot"><td class="gstick">合計（${want} 位要送）</td><td></td><td></td>${items.map((it) => `<td>${tot(it.id) || ''}</td>`).join('')}<td></td><td>${sent}/${want}</td><td>${list.filter((k) => k.giftSkip).length}</td></tr>
  </tbody></table>${list.length ? '' : '<div class="empty">沒有符合的對象。</div>'}</div>`;
}
function kgRec(callId) { // 找這位在目前活動的送禮紀錄，沒有就建一筆
  const k = Vault.get('calls', callId); let g = giftsOf(S.gev).find((x) => x.who?.type === 'call' && x.who.id === callId);
  if (!g && k) g = Vault.put('gifts', { eventId: S.gev, who: { type: 'call', id: callId, name: k.name || '' }, group: '', items: {}, method: '', status: '預備', note: '' });
  return g;
}
function kgTidy(g) { if (g && !Object.keys(g.items || {}).length && !g.method && g.status !== '已送' && !g.note) Vault.del('gifts', g.id); else if (g) Vault.put('gifts', g); }
function callTableHTML() {
  const all = calls(), q = S.q.trim(), hhs = callHouseholds(all), grpOf = (k) => hhs.get(hhKey(k));
  const areas = [...new Set(all.map((k) => k.area).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'zh-Hant')); if (S.kar !== '全部' && !areas.includes(S.kar)) S.kar = '全部';
  const markOk = (k) => ({ 全部: true, '已加 LINE': !!k.line, '未加 LINE': !k.line, 退件: !!k.mailBack, 不追蹤: !!k.other, 家戶待確認: hhPending(grpOf(k)) })[S.kmark];
  let list = all.filter((k) => (S.kprog === '全部' || callProg(k) === S.kprog) && (S.kar === '全部' || k.area === S.kar) && (S.klv === '全部' || (S.klv === '未分級' ? !k.level : k.level === S.klv)) && markOk(k) && (!q || [k.no, k.name, k.phone, k.home, k.addr, k.area, k.note, k.hh, k.otherWho].join(' ').includes(q))).sort(noCmp);
  if (S.ksort === 'addr') list.sort(addrCmp);
  if (S.ksort === 'co') list.sort((a, b) => (a.ord2 ?? 1e9) - (b.ord2 ?? 1e9) || noCmp(a, b));
  if (S.ksort === 'hh') { const first = new Map(); list.forEach((k, i) => { const key = hhKey(k); if (!first.has(key)) first.set(key, i); }); list = list.map((k, i) => [k, i]).sort((x, y) => first.get(hhKey(x[0])) - first.get(hhKey(y[0])) || x[1] - y[1]).map(([k]) => k); }
  const mine = all.filter((k) => !k.other), n = (f) => mine.filter(f).length, na = (f) => all.filter(f).length, prog = (k) => callProg(k);
  const funnel = [['名單', mine.length], ['打過', n((k) => (k.logs || []).length)], ['接通', n((k) => ['已接通', '已排訪', '已見面'].includes(prog(k)))], ['要年曆', n((k) => ['預備', '已送'].includes(callGift(k)))], ['加 LINE', n((k) => k.line)], ['排訪', n((k) => ['已排訪', '已見面'].includes(prog(k)))], ['見面', n((k) => k.met)], ['年曆已送', n((k) => callGift(k) === '已送')]];
  const dot = (v, l) => `<span class="kdot ${v}" title="${l}${v === 'ok' ? '接了' : v === 'no' ? '沒接' : '還沒打'}">${l[0]}</span>`;
  const chip = (key, x, c) => `<button class="chip${S[key] === x ? ' on' : ''}" data-${{ kprog: 'kprog', klv: 'klvf', kar: 'kar', kmark: 'kmark' }[key]}="${esc(x)}">${esc(x)} <span class="cnt">${c}</span></button>`;
  const KG = S.kview === 'gift', gevs = KG ? gevents() : []; if (KG && !gevs.some((x) => x.id === S.gev)) S.gev = gevs[0]?.id || null; const gev = KG ? Vault.get('gevents', S.gev) : null;
  const E = S.kedit && !KG, ti = (k, f, w, ph) => `<input class="nctxt" data-kte="${k.id}" data-kfield="${f}" value="${esc(k[f])}" style="width:${w}px"${ph ? ` placeholder="${ph}"` : ''}>`, cb = (k, f) => `<input type="checkbox" data-kte="${k.id}" data-kfield="${f}"${k[f] ? ' checked' : ''}>`;
  const cols = E ? ['序號', '姓名', '地區', '分級', '家戶', 'LINE', '退件', '不追蹤', '排訪日期', '見面', '重要事項', '']
    : ['序號', '姓名', '地區', '分級', '家戶', '三時段', '進度', '最近結果', '年曆', 'LINE', '排訪', '見面', '重要事項'];
  let prevKey = null, maxNo = -1, jumps = 0;
  const rows = list.map((k, ri) => { const nn = parseInt(k.no, 10), jump = (S.ksort === 'addr' || S.ksort === 'co') && !isNaN(nn) && nn < maxNo; if (jump) jumps++; if (!isNaN(nn)) maxNo = Math.max(maxNo, nn);
    const prevK = list[ri - 1], linked = prevK && k.hh && prevK.hh === k.hh, lk = S.ksort !== 'no' && prevK ? `<button class="klink${linked ? ' on' : ''}" data-klink="${k.id}" data-kprev="${prevK.id}" title="${linked ? '已跟上一位同一戶，再按一下拆開' : '跟上一位同一戶'}">${linked ? '🔗' : '＋🔗'}</button>` : ''; const sl = callSlots(k), l = logsOf(k)[0], gf = callGift(k), pg = prog(k), g = grpOf(k), key = hhKey(k);
    let head = '';
    if (S.ksort === 'hh' && key !== prevKey && g.length > 1) { const pend = hhPending(g); head = `<tr class="khh"><td colspan="${cols.length}"><b>家戶・${esc(g.find((x) => x.hh)?.hh || '同地址')}</b> <span class="sub">${g.length} 人・${esc(g[0].addr || '')}</span> ${pend ? `<span class="pill p-wait">待電話確認</span> <button class="btn sm" data-khhok="${esc(key)}">✓ 確認同一戶</button>` : '<span class="pill p-done">已確認</span>'}</td></tr>`; }
    prevKey = key;
    const hhCell = k.hh ? esc(k.hh) : g.length > 1 ? `<span class="sub">同地址 ${g.length} 人</span>` : '';
    const badges = `${k.mailBack ? '<span class="pill p-off">退件</span>' : ''}${k.other ? `<span class="pill p-off">${esc(skipWhy(k))}${k.otherWho ? '・' + esc(k.otherWho) : ''}</span>` : ''}`;
    if (E) return head + `<tr class="kedit${k.other ? ' kother' : ''}${linked ? ' klinked' : ''}"><td class="sub">${lk}${esc(k.no || '')}${jump ? '<span class="kjump" title="序號比上一位小">↰</span>' : ''}</td><td>${ti(k, 'name', 90)} <button class="linkbtn" data-ksel="${k.id}">開</button></td><td>${ti(k, 'area', 70)}</td>
      <td><select data-kte="${k.id}" data-kfield="level"><option value=""></option>${LEVELS.map((x) => `<option${k.level === x ? ' selected' : ''}>${x}</option>`).join('')}</select></td><td>${ti(k, 'hh', 80, g.length > 1 ? '同地址' : '')}</td>
      <td>${cb(k, 'line')}</td><td>${cb(k, 'mailBack')}</td><td>${cb(k, 'other')} ${k.other ? `<select data-kte="${k.id}" data-kfield="otherWhy">${KSKIP.map((x) => `<option${skipWhy(k) === x ? ' selected' : ''}>${x}</option>`).join('')}</select> ${ti(k, 'otherWho', 70, skipWhy(k) === '同事的件' ? '哪位同事' : '補充')}` : ''}</td>
      <td><input type="date" data-kte="${k.id}" data-kfield="visitDate" value="${esc(k.visitDate)}"></td><td>${cb(k, 'met')}</td><td>${ti(k, 'note', 180)}</td><td>${badges}</td></tr>`;
    return head + `<tr data-ksel="${k.id}" class="kp-${pg}${k.other ? ' kother' : ''}${linked ? ' klinked' : ''}">
     <td class="sub">${lk}${esc(k.no || '')}${jump ? '<span class="kjump" title="序號比上一位小">↰</span>' : ''}</td><td><b>${esc(k.name || '未填')}</b> ${badges}</td><td>${esc(k.area || '')}</td><td>${k.level ? `<b class="lvl lv-${esc(k.level)}">${esc(k.level)}</b>` : ''}</td><td>${hhCell}</td>
     <td><span class="kdots">${dot(sl['上午'], '上午')}${dot(sl['下午'], '下午')}${dot(sl['晚上'], '晚上')}</span></td><td><span class="kprog">${pg}</span></td>
     <td>${l ? `${esc(l.result)}<span class="sub"> ${fmtDay(l.date).slice(5)}</span>` : ''}</td><td>${gf ? `<span class="kg kg-${gf}">${gf}</span>` : ''}</td><td>${k.line ? '✓' : ''}</td>
     <td>${k.visitDate ? fmtDay(k.visitDate).slice(5) : pg === '已排訪' ? '✓' : ''}</td><td>${k.met ? '✓' : ''}</td><td class="knote">${esc((k.note || l?.note || '').slice(0, 40))}</td></tr>`; }).join('');
  const tableHTML = KG ? callGiftSheetHTML(list, gev, gevs) : `<div class="ktablewrap"><table class="ktable${E ? ' editing' : ''}"><thead><tr>${cols.map((c) => `<th>${c}</th>`).join('')}</tr></thead><tbody>${rows}</tbody></table>${list.length ? '' : '<div class="empty">沒有符合的對象。</div>'}</div>`;
  return `<div class="colhead"><h1>服務件總表 <span style="display:flex;gap:6px;flex-wrap:wrap"><span class="seg2"><button class="btn sm${KG ? '' : ' pri'}" data-kview="prog">進度</button><button class="btn sm${KG ? ' pri' : ''}" data-kview="gift">🎁 送禮</button></span><span class="seg2">${[['no', '序號'], ['addr', '地址'], ...(all.some((k) => k.ord2 != null) ? [['co', '公司順序']] : []), ['hh', '家戶']].map(([v, l]) => `<button class="btn sm${S.ksort === v ? ' pri' : ''}" data-ksort="${v}">${l}</button>`).join('')}</span>${KG ? '' : `<button class="btn${E ? ' pri' : ''}" data-act="kedit">${E ? '✓ 改好了' : '✎ 直接修改'}</button>`}<button class="btn" data-act="cflow">流程</button><button class="btn" data-act="klist">‹ 回到名單</button><button class="btn" data-act="importcalls">匯入名單</button><button class="btn" data-act="korder">匯入另一種排序</button></span></h1>
   <div class="kfunnel">${funnel.map(([l, v]) => `<div><span>${l}</span><b>${v}</b></div>`).join('<i>›</i>')}</div>
   <input class="search" id="q" placeholder="搜尋序號、姓名、電話、地址、家戶" value="${esc(S.q)}" autocomplete="off">
   <div class="chips"><span class="sub" style="align-self:center">進度</span>${['全部', ...PROGS].map((x) => chip('kprog', x, x === '全部' ? all.length : na((k) => prog(k) === x))).join('')}</div>
   <div class="chips"><span class="sub" style="align-self:center">分級</span>${['全部', ...LEVELS, '未分級'].map((x) => chip('klv', x, x === '全部' ? all.length : x === '未分級' ? na((k) => !k.level) : na((k) => k.level === x))).join('')}</div>
   <div class="chips"><span class="sub" style="align-self:center">標記</span>${KMARKS.map((x) => chip('kmark', x, all.filter((k) => { const o = S.kmark; S.kmark = x; const r = markOk(k); S.kmark = o; return r; }).length)).join('')}</div>
   ${areas.length ? `<div class="chips"><span class="sub" style="align-self:center">地區</span>${['全部', ...areas].map((x) => chip('kar', x, x === '全部' ? all.length : na((k) => k.area === x))).join('')}</div>` : ''}</div>
  ${S.ksort === 'co' ? `<div class="note">照你匯入的公司排序。${jumps ? `有 <b>${jumps}</b> 位的序號比上一位小（標 ↰），就是公司排得跟序號不一樣的地方。` : '跟序號順序完全一樣。'}${all.some((k) => k.ord2 == null) ? `另外有 ${all.filter((k) => k.ord2 == null).length} 位不在那份排序裡，放在最後。` : ''} <button class="linkbtn" data-act="korder">重新匯入排序</button></div>` : ''}
  ${S.ksort === 'addr' ? `<div class="note">照郵遞區號、地址排。${jumps ? `有 <b>${jumps}</b> 位的序號比上一位小（標 ↰），那些就是跟你原本順序不一樣的地方。` : '跟你原本的序號順序完全一樣。'}地址相鄰、看起來同一戶的，按序號前面的「＋🔗」就跟上一位拉成同一戶，再按一下拆開。</div>` : S.ksort === 'hh' ? '<div class="note">同一戶的排在一起。沒拉過的，地址一樣也會先放一起，標「待電話確認」。</div>' : ''}
  ${E ? '<div class="note">直接修改模式：改完一格點別的地方就存好。序號不能改，維持名單原本的編號。家戶那格打一樣的名字（例如「王家」），地址不同也會排在一起。</div>' : ''}
  ${tableHTML}
  <div class="sub" style="margin-top:8px">三時段：灰色＝還沒打、紅色＝打了沒接、綠色＝接了。點任何一列打開這個人。不追蹤的（同事的件、保經…）不算進上面的進度條。換排序只是換顯示順序，序號不會變。</div>`;
}
const LEVELS = ['A', 'B', 'C-1', 'C-2', 'C-3'];
const LEVEL_TIP = { A: 'VIP，高度優先', B: '有預算、有促約機會', 'C-1': '有關鍵人物，預算不夠', 'C-2': '關係好，觀念還沒到', 'C-3': '排斥、忙、沒需求，用寄的' };
const CALL_CHECKS = [['chkAgent', '服務人員確認'], ['chkHouse', '家戶確認'], ['chkJob', '職業類別確認']];
const WEEK = '日一二三四五六';
const slotOf = (iso) => { const t = new Date(iso), h = t.getHours(), wk = t.getDay(); return (wk === 0 || wk === 6 ? '假日' : '平日') + (h < 12 ? '上午' : h < 18 ? '下午' : '晚上'); };
const callWhen = (iso) => { const t = new Date(iso); return `週${WEEK[t.getDay()]} ${String(t.getHours()).padStart(2, '0')}:${String(t.getMinutes()).padStart(2, '0')}`; };
function answerStatsHTML(k) { // 每通電話自動記下撥打時間，算出什麼時段比較會接
  const logs = (k.logs || []).filter((l) => l.at); if (!logs.length) return '';
  const m = {}; logs.forEach((l) => { const s = slotOf(l.at); m[s] = m[s] || { ok: 0, no: 0 }; ['未接', '空號／錯號'].includes(l.result) ? m[s].no++ : m[s].ok++; });
  return `<div class="sec"><h2>接聽時段 <span class="sub">照每通電話的撥打時間自動統計</span></h2><div class="tags">${Object.entries(m).sort((a, b) => b[1].ok - a[1].ok).map(([s, v]) => `<span class="tag${v.ok ? ' on' : ''}" style="cursor:default">${s}：接 ${v.ok}・未接 ${v.no}</span>`).join('')}</div></div>`;
}
const scripts = () => notes().filter((n) => n.kind === 'script').sort((a, b) => (a.createdAt || '').localeCompare(b.createdAt || ''));
const logsOf = (k) => (k.logs || []).slice().sort((a, b) => (b.date || '').localeCompare(a.date || '') || (b.at || '').localeCompare(a.at || ''));
const stCls = (st) => (st === '已約訪' ? 'p-done' : st === '結案' ? 'p-off' : st === '再聯絡' ? 'p-wait' : 'p-go');
const SAMPLE_SCRIPT = '1. 開場、確認服務人員\n您好，請問是○○先生／小姐嗎？我是國泰人壽的服務人員馥瑋，之後由我負責您的保單服務。請問現在講電話方便嗎？大概三分鐘。\n\n2. 年曆\n年底我準備了年曆想送您，方便我親自送過去，還是寄給您比較方便？（確認地址）\n\n3. 確認家戶\n保單上的家人還住一起嗎？要不要也幫他們準備一份？\n\n4. 確認職業\n順便幫您確認一下，工作有沒有換？職業類別不一樣，意外險的理賠可能會受影響。\n\n5. 聊天（生活／工作／理財）\n最近工作忙嗎？家裡大家都好嗎？\n→ 邊聊邊判斷分級：A／B／C-1／C-2／C-3\n\n6. 收尾\n那我們約○月○日送過去，到時候再跟您聯絡。有任何保單的問題都可以直接找我。';
function callListHTML() {
  const q = S.q.trim(), all = calls(), today = localDay();
  const rank = (k) => ({ 再聯絡: 0, 待撥: 1, 已約訪: 2, 結案: 3 }[k.status] ?? 1);
  const areas = [...new Set(all.map((k) => k.area).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'zh-Hant'));
  if (S.kar !== '全部' && !areas.includes(S.kar)) S.kar = '全部';
  const list = all.filter((k) => (S.kst === '全部' || k.status === S.kst) && (S.kar === '全部' || k.area === S.kar) && (S.klv === '全部' || (S.klv === '未分級' ? !k.level : k.level === S.klv)) && (!q || [k.name, k.phone, k.source, k.note, k.addr, k.area].join(' ').includes(q)))
    .sort((a, b) => rank(a) - rank(b) || (a.nextDate || '9999').localeCompare(b.nextDate || '9999') || noCmp(a, b));
  return `<section class="col"><div class="colhead"><h1>服務件 <span style="display:flex;gap:6px"><button class="btn" data-act="cflow">流程</button><button class="btn pri" data-act="ktable">總表</button><button class="btn" data-act="importcalls">匯入名單</button><button class="btn" data-act="kphoto">📷 拍照新增</button><button class="btn" data-act="newcall">＋ 新增</button></span></h1>
  <div class="mobsync">${syncHTML()}</div>
  <input class="search" id="q" placeholder="搜尋姓名、電話、地址、來源" value="${esc(S.q)}" autocomplete="off">
  <div class="chips">${['全部', ...CALL_ST].map((x) => `<button class="chip${S.kst === x ? ' on' : ''}" data-kst="${x}">${x} <span class="cnt">${x === '全部' ? all.length : all.filter((k) => k.status === x).length}</span></button>`).join('')}</div>
  <div class="chips subchips"><span class="sub" style="align-self:center">分級</span>${['全部', ...LEVELS, '未分級'].map((x) => `<button class="chip${S.klv === x ? ' on' : ''}" data-klvf="${x}">${x} <span class="cnt">${x === '全部' ? all.length : x === '未分級' ? all.filter((k) => !k.level).length : all.filter((k) => k.level === x).length}</span></button>`).join('')}</div>
  ${areas.length ? `<div class="chips subchips"><span class="sub" style="align-self:center">地區</span>${['全部', ...areas].map((x) => `<button class="chip${S.kar === x ? ' on' : ''}" data-kar="${esc(x)}">${esc(x)} <span class="cnt">${x === '全部' ? all.length : all.filter((k) => k.area === x).length}</span></button>`).join('')}</div>` : ''}</div>
  <div class="list">${list.length ? list.map((k) => { const l = logsOf(k)[0], due = k.nextDate && k.status !== '結案' && k.nextDate <= today; return `<button class="item${k.id === S.ksel ? ' on' : ''}" data-ksel="${k.id}"><span class="t"><span>${k.no ? `<span class="kno">${esc(k.no)}</span>` : ''}${esc(k.name || '未填姓名')}</span><span class="pill ${stCls(k.status)}">${esc(k.status)}</span></span><span class="s">${k.level ? `<b class="lvl lv-${esc(k.level)}">${esc(k.level)}</b> ` : ''}${k.area ? esc(k.area) + ' · ' : ''}${l ? `上次 ${fmtDay(l.date).slice(5)} ${esc(l.result)}` : '還沒打過'}${k.nextDate && k.status !== '結案' ? ` · <span class="${due ? 'due' : ''}">下次 ${fmtDay(k.nextDate).slice(5)}</span>` : ''}${k.source ? ' · ' + esc(k.source) : ''}</span></button>`; }).join('') : `<div class="empty">${all.length ? '沒有符合的對象。' : '還沒有服務件名單。按「匯入名單」一次貼上很多人，或「＋ 新增」一個一個加。'}</div>`}</div></section>`;
}
function callDetailHTML() {
  const k = curCall();
  if (!k) return '<div class="empty">從左邊選一位要撥打的對象。</div>';
  const scs = scripts(); if (!scs.some((x) => x.id === S.script)) S.script = scs[0]?.id || null; const sc = Vault.get('notes', S.script);
  const cl = k.clientId && Vault.get('clients', k.clientId), logs = logsOf(k);
  return `<button class="back" data-act="back">‹ 服務件名單</button>${S.kfromTable ? '<button class="btn sm" data-act="ktable" style="margin-bottom:8px">‹ 回總表</button>' : ''}
  <div class="dhead"><input class="title" id="kf_name" data-kf="name" value="${esc(k.name)}" placeholder="姓名" aria-label="姓名"><button class="btn sm" data-act="kphotoin">📷 拍照帶入</button><span class="pill ${stCls(k.status)}">${esc(k.status)}</span></div>
  <div class="sec"><h2>聯絡資料</h2><div class="grid2">
   <label class="f"><span style="display:flex;justify-content:space-between">電話${k.phone ? `<a class="linkbtn" href="tel:${esc(k.phone.replace(/[^0-9+#*]/g, ''))}">撥打</a>` : ''}</span><input id="kf_phone" data-kf="phone" inputmode="tel" value="${esc(k.phone)}"></label>
   <label class="f">狀態<select id="kf_status" data-kf="status">${CALL_ST.map((x) => `<option${x === k.status ? ' selected' : ''}>${x}</option>`).join('')}</select></label>
   <label class="f">下次聯絡<input type="date" id="kf_nextDate" data-kf="nextDate" value="${esc(k.nextDate)}"></label>
   <label class="f">來源／名單<input id="kf_source" data-kf="source" value="${esc(k.source)}" placeholder="例如：10 月服務件"></label>
   <label class="f">住家電話${k.home ? ` <a class="linkbtn" href="tel:${esc(k.home.replace(/[^0-9+#*]/g, ''))}">撥打</a>` : ''}<input id="kf_home" data-kf="home" inputmode="tel" value="${esc(k.home)}"></label>
   <label class="f">序號<input id="kf_no" data-kf="no" value="${esc(k.no)}" placeholder="名單上的編號"></label>
   <label class="f">行政區<input id="kf_area" data-kf="area" value="${esc(k.area)}" placeholder="例如：桃園區"></label>
   <label class="f">方便聯絡時段<input id="kf_slot" data-kf="slot" value="${esc(k.slot)}" placeholder="例如：平日晚上"></label></div>
   <label class="f">地址<input id="kf_addr" data-kf="addr" value="${esc(k.addr)}"></label>
   <label class="f">家戶<input id="kf_hh" data-kf="hh" value="${esc(k.hh)}" placeholder="同一戶的人打一樣的名字，例如：王家"></label>
   <div class="tags"><button class="tag${k.line ? ' on' : ''}" data-kchk="line">已加 LINE</button><button class="tag${k.hhOk ? ' on' : ''}" data-kchk="hhOk">${k.hhOk ? '✓ ' : ''}家戶電話確認過</button><button class="tag${k.mailBack ? ' on' : ''}" data-kchk="mailBack">寄信退件</button><button class="tag${k.other ? ' on' : ''}" data-kchk="other">${k.other ? '✓ ' : ''}不追蹤</button></div>
   ${k.other ? `<div class="grid2"><label class="f">原因<select id="kf_otherWhy" data-kf="otherWhy">${KSKIP.map((x) => `<option${skipWhy(k) === x ? ' selected' : ''}>${x}</option>`).join('')}</select></label><label class="f">${skipWhy(k) === '同事的件' ? '哪位同事' : '補充'}<input id="kf_otherWho" data-kf="otherWho" value="${esc(k.otherWho)}"></label></div>` : ''}
   <label class="f">備註<textarea id="kf_note" class="grow" data-kf="note" rows="2" placeholder="家庭、保單、個性…">${esc(k.note)}</textarea></label></div>
  <div class="sec"><h2>分級 <span class="sub">打完電話再分，之後可以隨時改</span></h2>
   <div class="tags">${LEVELS.map((x) => `<button class="tag lvtag${k.level === x ? ' on' : ''}" data-klv="${x}"><b>${x}</b> ${LEVEL_TIP[x]}</button>`).join('')}</div></div>
  <div class="sec"><h2>電話裡要確認的</h2>
   <div class="tags">${CALL_CHECKS.map(([key, l]) => `<button class="tag${k[key] ? ' on' : ''}" data-kchk="${key}">${k[key] ? '✓ ' : ''}${l}</button>`).join('')}</div>
   <label class="f">職業類別／工作<input id="kf_job" data-kf="job" value="${esc(k.job)}" placeholder="例如：第一類、辦公室行政"></label>
   <label class="f">家戶（同住家人、保單上的人）<input id="kf_house" data-kf="house" value="${esc(k.house)}"></label></div>
  <details class="sec"${k.addrNote || k.policy ? ' open' : ''}><summary><h2 style="display:inline">打電話前、出發前的功課</h2></summary>
   <label class="f">住址觀察<input id="kf_addrNote" data-kf="addrNote" value="${esc(k.addrNote)}" placeholder="例如：電梯大樓、透天、社區"></label>
   <label class="f">保單概況<textarea id="kf_policy" class="grow" data-kf="policy" rows="2" placeholder="有哪些保單、繳費狀況…">${esc(k.policy)}</textarea></label></details>
  <div class="sec"><h2>約訪</h2><div class="grid2"><label class="f">約訪日期<input type="date" id="kf_visitDate" data-kf="visitDate" value="${esc(k.visitDate)}"></label>
   <div class="f">結果<div class="tags"><button class="tag${k.met ? ' on' : ''}" data-kchk="met">${k.met ? '✓ ' : ''}已見面</button><button class="tag${k.giftSkip ? ' on' : ''}" data-kchk="giftSkip">不送年曆</button></div></div></div></div>
  ${callHHSectionHTML(k)}
  ${giftSectionHTML('call', k.id, k.name)}
  ${answerStatsHTML(k)}
  <details class="sec script" open><summary><h2 style="display:inline">話術 SOP</h2></summary>
   ${scs.length ? `<div class="actions" style="align-items:center"><select id="kScript" aria-label="選擇話術">${scs.map((x) => `<option value="${x.id}"${x.id === S.script ? ' selected' : ''}>${esc(x.title)}</option>`).join('')}</select><button class="linkbtn" data-act="editscript">編輯</button><button class="linkbtn" data-act="newscript">＋ 新話術</button></div><pre class="scripttext">${esc(sc?.body || '')}</pre>${nimgsHTML(sc || {})}`
    : `<div class="sub">還沒有話術。可以先用範例改成你自己的說法。</div><div class="actions"><button class="btn pri" data-act="samplescript">用範例開始</button><button class="btn" data-act="newscript">自己寫</button></div>`}</details>
  <div class="sec"><h2>記錄這通</h2>
   <div class="tags">${RESULTS.map(([r]) => `<button class="tag${S.kres === r ? ' on' : ''}" data-kres="${r}">${r}</button>`).join('')}</div>
   <label class="f">聊了什麼<textarea id="kNote" rows="3" placeholder="重點、對方的需求、約了什麼…"></textarea></label>
   <div class="actions" style="align-items:flex-end"><label class="f" style="width:170px">下次聯絡（可不填）<input type="date" id="kNext"></label><button class="btn pri" data-act="logcall">記錄</button></div></div>
  <div class="sec"><h2>通話紀錄 <span class="sub">${logs.length ? logs.length + ' 通，新的在上面' : ''}</span></h2>
   ${logs.length ? `<div class="tl">${logs.map((l) => `<div class="tli"><span class="tldot"></span><div class="tlrow"><span class="tldate" style="padding:3px 6px">${fmtDay(l.date)}${l.at ? ' ' + callWhen(l.at) : ''}</span><span class="tltitle" style="padding:3px 6px">${esc(l.result)}</span><button class="tldel" data-kdel="${l.id}" aria-label="刪除這通">✕</button></div>${l.note ? `<div class="lognote">${esc(l.note)}</div>` : ''}</div>`).join('')}</div>` : '<div class="sub">還沒有紀錄。</div>'}</div>
  <div class="actions"><button class="btn pri savebtn" data-act="savenow">儲存</button>${cl ? `<button class="btn" data-goclient="${cl.id}">看客戶資料 ›</button>` : '<button class="btn" data-act="tocliient">轉成客戶</button>'}<button class="btn danger" data-act="delcall">刪除</button></div>`;
}
const areaOf = (addr) => { const m = String(addr || '').replace(/\s+/g, '').match(/^\d{0,5}(?:[^市縣]{1,3}[市縣])?([^市縣區鄉鎮]{1,3}[區鄉鎮市])/); return m ? m[1] : ''; };
function parseCallRow(line) { // 一行一個人：序號、姓名、行動電話、住家電話、地址、備註（從 Excel 貼上或 PDF 讀進來都可以）
  if (line.includes('\t')) { const c = line.split('\t').map((x) => x.trim().replace(/^[-－—]+$/, '')); if (c.length >= 5 && /^\d{0,5}$/.test(c[0]) && !/\d/.test(c[1] || '')) { const r = { no: c[0], name: c[1], phone: c[2].replace(/\s/g, ''), home: c[3].replace(/\s/g, ''), addr: c[4], note: c.slice(5).join(' ').trim() }; r.area = areaOf(r.addr); return r; } }
  const cells = line.includes('\t') ? line.split('\t').map((x) => x.trim()) : line.split(/[,，]|\s+/).map((x) => x.trim());
  const r = { no: '', name: '', phone: '', home: '', addr: '', note: '' }; const rest = [];
  for (let i = 0; i < cells.length; i++) {
    let t = cells[i]; if (!t || /^[-－—]+$/.test(t)) continue;
    if (!r.no && !r.name && /^\d{1,4}$/.test(t) && !(cells[i + 1] && /[市縣]/.test(cells[i + 1]) && t.length === 3 && r.name)) { r.no = t; continue; }
    if (/^\d{3,5}$/.test(t) && cells[i + 1] && /[市縣區鄉鎮]/.test(cells[i + 1])) { t = t + ' ' + cells[++i]; }
    if (!r.phone && /^09\d{2}[-\s]?\d{3}[-\s]?\d{3}$/.test(t)) { r.phone = t; continue; }
    if (!r.home && /^0\d{1,2}[-\s]?\d{6,8}(#\d+)?$/.test(t)) { r.home = t; continue; }
    if (!r.addr && /[市縣].*[路街巷弄號段村里鄰]|^\d{3,5}\s*\S+[市縣]/.test(t)) { r.addr = t; continue; }
    if (!r.name && !/\d/.test(t) && t.length <= 12) { r.name = t; continue; }
    rest.push(t);
  }
  r.note = rest.join(' '); r.area = areaOf(r.addr); return r;
}
async function pdfToRows(file) { // 讀名單 PDF：先把同一列、靠在一起的字黏回來，再照表頭（序號、姓名、電話、地址…）的位置切成欄位
  const lib = await loadPdfjs(); const pdf = await lib.getDocument({ data: new Uint8Array(await file.arrayBuffer()), cMapUrl: 'vendor/pdfjs/cmaps/', cMapPacked: true }).promise; const lines = [];
  const HEADS = [['序號', '編號'], ['姓名', '客戶姓名', '名稱'], ['行動電話', '手機', '行動'], ['住家電話', '市話', '電話'], ['地址', '住址', '通訊地址'], ['備註']];
  let cols = null; // 表頭位置，下一頁沒有表頭就沿用
  for (let n = 1; n <= pdf.numPages; n++) {
    const tc = await (await pdf.getPage(n)).getTextContent();
    const items = tc.items.filter((x) => x.str.trim()).map((x) => ({ s: x.str.trim(), x: x.transform[4], y: x.transform[5], w: x.width || 0, h: Math.abs(x.transform[3]) || 10 }));
    const rows = []; items.sort((a, b) => b.y - a.y || a.x - b.x).forEach((it) => { const r = rows.find((r) => Math.abs(r.y - it.y) < Math.max(3, it.h * 0.4)); r ? r.items.push(it) : rows.push({ y: it.y, items: [it] }); });
    rows.forEach((r) => r.items.sort((a, b) => a.x - b.x));
    const head = rows.find((r) => { const t = r.items.map((i) => i.s).join('').replace(/\s/g, ''); return /姓名|名稱/.test(t) && /地址|住址|電話/.test(t); });
    if (head) { // 每個字的位置攤開，找表頭每個標題的起點和終點
      const ch = []; head.items.forEach((i) => { const t = i.s.replace(/\s/g, ''), cw = t.length ? i.w / t.length : 0; [...t].forEach((c, k) => ch.push({ c, x0: i.x + k * cw, x1: i.x + (k + 1) * cw })); });
      const txt = ch.map((c) => c.c).join(''); const found = [];
      HEADS.forEach((alts, ci) => { for (const a of alts) { const at = txt.indexOf(a); if (at >= 0 && !found.some((f) => at < f.end && at + a.length > f.at)) { found.push({ ci, at, end: at + a.length, mid: (ch[at].x0 + ch[at + a.length - 1].x1) / 2 }); break; } } });
      if (found.length >= 2) cols = found.sort((a, b) => a.mid - b.mid);
    }
    rows.filter((r) => r !== head && (!head || r.y < head.y)).forEach((r) => {
      if (!cols) { // 沒有表頭：靠在一起的字黏起來，中間有空隙就當成不同欄
        const parts = []; r.items.forEach((i) => { const last = parts[parts.length - 1]; if (last && i.x - last.end < i.h * 0.8) { last.s += i.s; last.end = i.x + i.w; } else parts.push({ s: i.s, end: i.x + i.w }); });
        lines.push(parts.map((x) => x.s).join('\t')); return;
      }
      const chunks = []; r.items.forEach((i) => { const last = chunks[chunks.length - 1]; if (last && i.x - last.end < i.h * 1.2) { last.s += (i.x - last.end > i.h * 0.3 ? ' ' : '') + i.s; last.end = i.x + i.w; } else chunks.push({ s: i.s, x: i.x, end: i.x + i.w }); });
      const cells = Array.from({ length: 6 }, () => []);
      chunks.forEach((i) => { const mid = (i.x + i.end) / 2; let best = cols[0], bd = 1e9; cols.forEach((c) => { const d = Math.abs(c.mid - mid); if (d < bd) { bd = d; best = c; } }); cells[best.ci].push(i.s); });
      lines.push(cells.map((x) => x.join(' ')).join('\t'));
    });
  }
  return lines.filter((l) => l.replace(/\t/g, '').trim());
}
function openCallImport() {
  ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>匯入服務件名單</h2><button class="btn" id="ciCancel">取消</button><button class="btn pri" id="ciGo">匯入</button></div>
  <div class="ovbody" style="flex-direction:column;max-width:760px;margin:0 auto;width:100%">
   <div class="note">一行一個人，順序是：序號、姓名、行動電話、住家電話、地址、備註（沒有的可以空著）。可以從 Excel 整列複製貼上，或直接選公司給的 PDF 名單，讀完會先列在下面讓你檢查。行政區會從地址自動抓。</div>
   <div class="actions" style="margin:0"><label class="btn">選 PDF 名單<input type="file" id="ciPdf" accept="application/pdf,.pdf" hidden></label><span class="sub" id="ciPdfMsg"></span></div>
   <label class="f">來源／名單名稱<input id="ciSrc" placeholder="例如：10 月服務件" value="${new Date().getMonth() + 1} 月服務件"></label>
   <label class="f">名單<textarea id="ciText" rows="10" placeholder="001\t王小明\t0912345678\t03-1234567\t330 桃園市桃園區某某路1號"></textarea></label>
   <div class="sub" id="ciCount"></div><div id="ciPrev"></div></div></div>`;
  const parse = () => $('ciText').value.split('\n').map((l) => l.replace(/\s+$/, '')).filter((l) => l.trim()).map(parseCallRow).filter((x) => x.name || x.phone || x.home);
  const exist = calls(); const dup = (r) => exist.some((k) => k.name === r.name && (r.no ? k.no === r.no : (k.phone || '') === (r.phone || '')));
  const show = () => { const rows = parse(), d = rows.filter(dup).length; $('ciCount').textContent = rows.length ? `讀到 ${rows.length} 位${d ? `，其中 ${d} 位已經在名單裡，會略過` : ''}。下面是前 8 位，請檢查欄位有沒有對：` : '';
    $('ciPrev').innerHTML = rows.length ? `<table class="gtable"><thead><tr><th>序號</th><th>姓名</th><th>行動</th><th>住家</th><th>行政區</th><th>地址</th></tr></thead><tbody>${rows.slice(0, 8).map((r) => `<tr><td>${esc(r.no)}</td><td>${esc(r.name)}</td><td>${esc(r.phone)}</td><td>${esc(r.home)}</td><td>${esc(r.area)}</td><td>${esc(r.addr)}</td></tr>`).join('')}</tbody></table>` : ''; };
  $('ciText').oninput = show;
  $('ciPdf').onchange = async (e) => { const f = e.target.files[0]; e.target.value = ''; if (!f) return; $('ciPdfMsg').textContent = '讀取中…'; try { const ls = await pdfToRows(f); $('ciText').value = ls.join('\n'); $('ciPdfMsg').textContent = `讀到 ${ls.length} 列`; show(); } catch (err) { console.warn(err); $('ciPdfMsg').textContent = '讀不了這個 PDF，請改用 Excel 複製貼上'; } };
  $('ciCancel').onclick = () => (ov.innerHTML = '');
  $('ciGo').onclick = () => { const rows = parse().filter((r) => !dup(r)); if (!rows.length) { toast('沒有新的人可以匯入'); return; } const src = $('ciSrc').value.trim(); rows.forEach((r) => Vault.put('calls', { ...r, source: src, status: '待撥', nextDate: '', logs: [] })); ov.innerHTML = ''; S.kst = '全部'; S.ktable = true; App.render(); toast(`已匯入 ${rows.length} 位`); };
  setTimeout(() => $('ciText')?.focus(), 50);
}
function openCallOrderImport() { // 同一份名單、不同排列：只照序號對回去，記下順序，不新增人、不改資料
  ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>匯入另一種排序</h2><button class="btn" id="coX">取消</button><button class="btn pri" id="coGo">套用排序</button></div>
  <div class="ovbody" style="flex-direction:column;max-width:760px;margin:0 auto;width:100%">
   <div class="note">選公司排好的那份 PDF（序號一樣、排列不一樣的）。這裡只會照<b>序號</b>把人對回去、記下順序，不會新增人，也不會改任何資料或序號。套用後總表排序多一個「公司順序」。</div>
   <div class="actions" style="margin:0"><label class="btn">選 PDF<input type="file" id="coPdf" accept="application/pdf,.pdf" hidden></label><span class="sub" id="coMsg"></span></div>
   <label class="f">或貼上（一行一個人，第一欄是序號）<textarea id="coText" rows="8"></textarea></label><div class="sub" id="coRes"></div></div></div>`;
  const match = () => { const rows = $('coText').value.split('\n').filter((l) => l.trim()).map(parseCallRow).filter((r) => r.no), all = calls(), byNo = new Map(all.map((k) => [String(parseInt(k.no, 10)), k]));
    const hit = [], miss = [], diff = []; rows.forEach((r, i) => { const k = byNo.get(String(parseInt(r.no, 10))); if (!k) { miss.push(r.no); return; } if (r.name && k.name && r.name !== k.name) { diff.push(r.no); return; } if (!hit.some(([x]) => x === k)) hit.push([k, i]); });
    return { rows, hit, miss, diff, rest: all.length - hit.length }; };
  const show = () => { const m = match(); $('coRes').innerHTML = m.rows.length ? `讀到 ${m.rows.length} 列：對到 <b>${m.hit.length}</b> 位${m.miss.length ? `・序號找不到 ${m.miss.length} 位（${esc(m.miss.slice(0, 10).join('、'))}${m.miss.length > 10 ? '…' : ''}）` : ''}${m.diff.length ? `・序號一樣但名字不同 ${m.diff.length} 位（${esc(m.diff.slice(0, 10).join('、'))}），先不套用` : ''}${m.rest ? `・你名單裡有 ${m.rest} 位不在這份，會排在最後` : ''}` : ''; };
  $('coText').oninput = show;
  $('coPdf').onchange = async (e) => { const f = e.target.files[0]; e.target.value = ''; if (!f) return; $('coMsg').textContent = '讀取中…'; try { const ls = await pdfToRows(f); $('coText').value = ls.join('\n'); $('coMsg').textContent = `讀到 ${ls.length} 列`; show(); } catch (err) { console.warn(err); $('coMsg').textContent = '讀不了這個 PDF'; } };
  $('coX').onclick = () => (ov.innerHTML = '');
  $('coGo').onclick = () => { const m = match(); if (!m.hit.length) { toast('沒有對到任何人'); return; } calls().forEach((k) => { if (k.ord2 != null) { delete k.ord2; Vault.put('calls', k); } }); m.hit.forEach(([k, i]) => { k.ord2 = i; Vault.put('calls', k); }); ov.innerHTML = ''; S.ksort = 'co'; S.ktable = true; App.render(); toast(`已套用 ${m.hit.length} 位的排序`); };
}
/* ================= 送禮專區：每個活動（年曆、中秋、出國伴手禮…）各有品項庫存和送禮名單 ================= */
const GIFT_TYPES = [['client', '現有客戶'], ['call', '服務件'], ['other', '親友']];
const gevents = () => Vault.all('gevents').sort((a, b) => (b.createdAt || '').localeCompare(a.createdAt || ''));
const giftsOf = (evId) => Vault.all('gifts').filter((g) => g.eventId === evId);
const giftWho = (g) => (g.who?.type === 'client' ? Vault.get('clients', g.who.id)?.name : g.who?.type === 'call' ? Vault.get('calls', g.who.id)?.name : '') || g.who?.name || '未填';
function giftTableHTML(all, ev) { // 像試算表：一列一個人，一欄一個品項，直接改數量；服務件合成一列只看總數
  const fold = S.gtype === '全部', svc = fold ? all.filter((g) => g.who?.type === 'call') : [], list = fold ? all.filter((g) => g.who?.type !== 'call') : all;
  const items = giftItemsSorted(ev.items || []); const cats = []; items.forEach((it) => { const c = it.cat || '未分類'; const last = cats[cats.length - 1]; last && last.c === c ? last.n++ : cats.push({ c, n: 1 }); });
  const tot = (id) => all.reduce((a, g) => a + (g.items?.[id] || 0), 0), stot = (id) => svc.reduce((a, g) => a + (g.items?.[id] || 0), 0);
  return `<div class="ktablewrap"><table class="ktable gsheet"><thead>
   <tr><th rowspan="2" class="gstick">名字</th><th rowspan="2">類型</th><th rowspan="2">分類</th>${cats.map((c) => `<th colspan="${c.n}" class="gcath">${esc(c.c)}</th>`).join('')}<th rowspan="2">送法</th><th rowspan="2">已送</th><th rowspan="2">備註</th></tr>
   <tr>${items.map((it) => `<th class="gith">${giftThumb(it)}<div>${esc(it.name || it.code || '品項')}</div></th>`).join('')}</tr></thead><tbody>
   ${list.map((g) => `<tr class="${g.status === '已送' ? 'gdone' : ''}"><td class="gstick"><button class="linkbtn" data-gopen="${g.id}">${esc(giftWho(g))}</button></td><td><select class="gcell gtype" data-gwho="${g.id}">${GIFT_TYPES.map(([t, l]) => `<option value="${t}"${(g.who?.type || 'other') === t ? ' selected' : ''}>${l}</option>`).join('')}</select></td>
     <td><select class="gcell gw" data-gg="${g.id}:group">${['', ...giftGroups(), ...(g.group && !giftGroups().includes(g.group) ? [g.group] : [])].map((x) => `<option${(g.group || '') === x ? ' selected' : ''}>${esc(x)}</option>`).join('')}<option value="__edit">✎ 編修分類…</option></select></td>
     ${items.map((it) => `<td><input class="gcell gn" inputmode="numeric" data-gq2="${g.id}:${it.id}" value="${g.items?.[it.id] || ''}"></td>`).join('')}
     <td><select class="gcell" data-gg="${g.id}:method">${['', '當面', '寄送', '放管理室'].map((m) => `<option${(g.method || '') === m ? ' selected' : ''}>${m}</option>`).join('')}</select></td>
     <td style="text-align:center"><input type="checkbox" class="gchk" data-gdone="${g.id}"${g.status === '已送' ? ' checked' : ''}></td>
     <td><input class="gcell gw2" data-gg="${g.id}:note" value="${esc(g.note)}"></td></tr>`).join('')}
   ${fold && calls().length ? `<tr class="gsvc"><td class="gstick"><button class="linkbtn" data-act="kgiftgo">服務件（${svc.length} 位）›</button></td><td class="sub">服務件</td><td></td>${items.map((it) => `<td style="text-align:center">${stot(it.id) || ''}</td>`).join('')}<td></td><td style="text-align:center">${svc.filter((g) => g.status === '已送').length}/${svc.length}</td><td class="sub">自動加總，到服務件專區的「🎁 送禮」表填</td></tr>` : ''}
   <tr class="gtot"><td class="gstick">合計</td><td></td><td></td>${items.map((it) => `<td>${tot(it.id) || ''}</td>`).join('')}<td></td><td>${all.filter((g) => g.status === '已送').length}/${all.length}</td><td></td></tr>
  </tbody></table></div><div class="sub" style="margin-top:6px">格子可以直接改，改完點別的地方就會存。點名字可以打開完整資料。</div>`;
}
function giftThumb(it, edit) { // 品項小圖；編修時點一下可以換圖
  const img = it.img ? `<img class="gthumbimg" data-fid="${it.img}" alt="">` : '';
  if (edit) return `<button type="button" class="gthumb${it.img ? '' : ' empty'}" data-gimg="${it.id}" title="放圖片">${img || '＋圖'}</button>`;
  return it.img ? `<button type="button" class="gthumb" data-gpic="${it.img}" title="看大圖">${img}</button>` : '<span class="gthumb none"></span>';
}
let jszipReady = null;
const loadJSZip = () => window.JSZip ? Promise.resolve(window.JSZip) : (jszipReady ||= new Promise((ok, no) => { const sc = document.createElement('script'); sc.src = 'vendor/jszip.min.js'; sc.onload = () => ok(window.JSZip); sc.onerror = () => { jszipReady = null; no(new Error('load')); }; document.head.appendChild(sc); }));
async function giftBulkImages(files) { // 一次放很多張：照檔名裡的代號（或品名）對到品項；可以選壓縮檔
  const ev = Vault.get('gevents', S.gev); if (!ev || !(ev.items || []).length) { toast('先加品項'); return; }
  toast('讀取中…'); const pics = [];
  for (const f of files) {
    if (/\.zip$/i.test(f.name || '') || /zip/.test(f.type)) { try { const Z = await loadJSZip(); const z = await Z.loadAsync(f); for (const e of Object.values(z.files)) { if (!e.dir && /\.(jpe?g|png|webp|gif)$/i.test(e.name) && !/__MACOSX/.test(e.name)) pics.push({ name: e.name.split('/').pop(), blob: await e.async('blob') }); } } catch (err) { toast('壓縮檔打不開'); } }
    else if (/^image\//.test(f.type)) pics.push({ name: f.name || '', blob: f });
  }
  const norm = (t) => String(t || '').toUpperCase().replace(/[\s_\-－]/g, '');
  const items = [...ev.items].sort((a, b) => norm(b.code).length - norm(a.code).length); let n = 0; const used = new Set();
  for (const p of pics) {
    const nm = norm(p.name); const it = items.find((x) => !used.has(x.id) && ((x.code && nm.includes(norm(x.code))) || (!x.code && x.name && nm.includes(norm(x.name)))));
    if (!it) continue; used.add(it.id);
    const b = await shrink(p.blob, 800, 0.85); if (!b) continue; const old = it.img; it.img = await Files.add(b, it.name || it.code || '品項'); if (old) Files.remove(old); n++;
  }
  if (n) { Vault.put('gevents', ev); App.render(); }
  const miss = ev.items.filter((x) => !x.img).map((x) => x.name || x.code);
  toast(n ? `已放上 ${n} 張圖${miss.length ? '，還沒有圖：' + miss.slice(0, 4).join('、') + (miss.length > 4 ? '…' : '') : ''}` : '沒有對到任何品項：檔名要有品項代號，例如 A03、CAWP-014');
}
function giftItemImage(itemId) { // 選照片：可以直接用，或拉四個角裁出一個品項
  const ev = Vault.get('gevents', S.gev); const it = ev?.items.find((x) => x.id === itemId); if (!it) return;
  const save = async (blob) => { const old = it.img; it.img = await Files.add(blob, it.name || it.code || '品項'); if (old) Files.remove(old); Vault.put('gevents', ev); App.render(); toast('已放上圖片'); };
  ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>${esc(it.name || it.code || '品項')} 的圖片</h2><button class="btn" id="giX">取消</button></div>
  <div class="ovbody" style="flex-direction:column;max-width:560px;margin:0 auto;width:100%">
   <div class="note">選一張照片或截圖。如果是整張型錄，選「裁切」拉四個角框出這一個品項。</div>
   <div class="actions"><label class="btn pri">選照片（直接用）<input type="file" id="giPick" accept="image/*" hidden></label><label class="btn">選照片再裁切<input type="file" id="giCrop" accept="image/*" hidden></label><button class="btn" id="giPaste">📋 貼上</button>${it.img ? '<button class="btn danger" id="giDel">拿掉圖片</button>' : ''}</div></div></div>`;
  $('giX').onclick = () => (ov.innerHTML = '');
  $('giPick').onchange = async (e) => { const f = e.target.files[0]; if (!f) return; ov.innerHTML = ''; const b = await shrink(f, 800, 0.85); if (b) save(b); };
  $('giPaste').onclick = async () => { const fs = await clipFiles(); if (!fs.length) return; ov.innerHTML = ''; const b = await shrink(fs[0], 800, 0.85); if (b) save(b); };
  $('giCrop').onchange = (e) => { const f = e.target.files[0]; if (!f) return; openScanner({ docs: [] }, { filter: 'color', names: [it.name || '品項'], onDone: async (pages) => { if (pages[0]) save(pages[0].blob); else App.render(); }, onCancel: () => App.render() }); onFiles([f]); };
  if ($('giDel')) $('giDel').onclick = () => { Files.remove(it.img); delete it.img; Vault.put('gevents', ev); ov.innerHTML = ''; App.render(); };
}
function giftGroups() { // 分類清單：設定裡的順序在前，用過但沒在清單裡的接在後面
  const set = Vault.get('settings', 'giftGroups')?.list || []; const used = [...new Set(Vault.all('gifts').map((g) => g.group).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'zh-Hant'));
  return [...set, ...used.filter((x) => !set.includes(x))];
}
function openGiftGroupEditor() { // 編修分類：改名（已經用的人一起改）、刪除、新增、拖曳排序
  let rows = giftGroups().map((g) => ({ id: uid(), old: g, name: g })); const v = document.createElement('div'); document.body.appendChild(v);
  const render = () => {
    v.innerHTML = `<div class="ov" style="z-index:30"><div class="ovhead"><h2>編修分類</h2><button class="btn" id="ggX">取消</button><button class="btn pri" id="ggOk">儲存</button></div>
    <div class="ovbody" style="flex-direction:column;max-width:520px;margin:0 auto;width:100%">
     <div class="note">改名字的話，已經選了這個分類的人會一起改。刪掉的分類，用過的人會變成空白。按住 ⠿ 拖曳排順序。</div>
     <div data-sec="gg"><div class="ncgrp" data-gi="0"><div class="ncbody">${rows.map((r) => `<div class="ncrow" data-iid="${r.id}"><span class="ncdrag">⠿</span><input class="nctxt" data-ggn="${r.id}" value="${esc(r.name)}"><button class="btn sm ncdel" data-ggd="${r.id}">✕</button></div>`).join('')}</div></div></div>
     <button class="btn sm" id="ggAdd" style="align-self:flex-start">＋ 新增分類</button></div></div>`;
    v.querySelector('#ggX').onclick = () => v.remove();
    v.querySelectorAll('[data-ggn]').forEach((x) => (x.oninput = () => { rows.find((r) => r.id === x.dataset.ggn).name = x.value.trim(); }));
    v.querySelectorAll('[data-ggd]').forEach((b) => (b.onclick = () => { if (!b.dataset.sure) { b.dataset.sure = 1; b.textContent = '刪？'; return; } rows = rows.filter((r) => r.id !== b.dataset.ggd); render(); }));
    v.querySelector('#ggAdd').onclick = () => { const r = { id: uid(), old: '', name: '' }; rows.push(r); render(); v.querySelector(`[data-ggn="${r.id}"]`)?.focus(); };
    makeSortable(v, () => { const m = Object.fromEntries(rows.map((r) => [r.id, r])); rows = [...v.querySelectorAll('.ncrow')].map((el) => m[el.dataset.iid]).filter(Boolean); render(); });
    v.querySelector('#ggOk').onclick = () => {
      const keep = rows.filter((r) => r.name); const map = {}; keep.forEach((r) => { if (r.old) map[r.old] = r.name; });
      const gone = giftGroups().filter((g) => !keep.some((r) => r.old === g));
      Vault.all('gifts').forEach((g) => { if (!g.group) return; if (map[g.group] && map[g.group] !== g.group) { g.group = map[g.group]; Vault.put('gifts', g); } else if (gone.includes(g.group)) { g.group = ''; Vault.put('gifts', g); } });
      Vault.put('settings', { id: 'giftGroups', list: [...new Set(keep.map((r) => r.name))] }); v.remove(); App.render(); toast('分類已儲存');
    };
  };
  render();
}
const giftItemsSorted = (items) => { const order = []; items.forEach((it) => { const c = it.cat || ''; if (!order.includes(c)) order.push(c); }); return order.flatMap((c) => items.filter((it) => (it.cat || '') === c)); }; // 照自己排的順序，同分類放一起
const giftCatsSorted = (items) => { const m = new Map(); giftItemsSorted(items).forEach((it) => { const c = it.cat || ''; m.has(c) ? m.get(c).push(it) : m.set(c, [it]); }); return [...m]; };
function giftCatRow(cat, arr, cnt) { const its = arr.filter((x) => (x.cat || '') === cat), sum = (f) => its.reduce((a, it) => a + f(it), 0); const q = sum((it) => it.qty || 0), pl = sum((it) => cnt[it.id]?.plan || 0), dn = sum((it) => cnt[it.id]?.done || 0);
  return `<tr class="gcat"><td colspan="3">${esc(cat || '未分類')}</td><td>${q}</td><td>${pl}</td><td>${dn}</td><td class="${q - pl - dn < 0 ? 'gneg' : ''}">${q - pl - dn}</td></tr>`; }
const giftItemsText = (g, ev) => Object.entries(g.items || {}).filter(([, n]) => n > 0).map(([id, n]) => { const it = ev?.items.find((x) => x.id === id); return `${it ? it.name || it.code || '品項' : '已刪品項'}×${n}`; }).join('、');
function giftCounts(ev) {
  const out = {}; (ev.items || []).forEach((it) => (out[it.id] = { plan: 0, done: 0 }));
  giftsOf(ev.id).forEach((g) => Object.entries(g.items || {}).forEach(([id, n]) => { if (!out[id]) return; out[id][g.status === '已送' ? 'done' : 'plan'] += n; }));
  return out;
}
function giftsHTML() {
  const evs = gevents(); if (!evs.some((x) => x.id === S.gev)) S.gev = evs[0]?.id || null; const ev = Vault.get('gevents', S.gev);
  const head = `<div class="colhead"><h1>送禮專區 <span style="display:flex;gap:6px"><button class="btn" data-act="newgevent">＋ 新活動</button></span></h1>
   <div class="chips">${evs.map((x) => `<button class="chip${x.id === S.gev ? ' on' : ''}" data-gev="${x.id}">${esc(x.name || '未命名')}</button>`).join('')}</div></div>`;
  if (!ev) return head + `<div class="empty">還沒有活動。按「＋ 新活動」，例如「2027 年曆」。年曆、中秋、拜訪、出國伴手禮都可以用，下一次可以從舊活動複製品項。</div>`;
  const cnt = giftCounts(ev), all = giftsOf(ev.id);
  const list = all.filter((g) => (S.gtype === '全部' || g.who?.type === S.gtype) && (S.gst === '全部' || (g.status || '預備') === S.gst))
    .sort((a, b) => ((a.status === '已送') - (b.status === '已送')) || GIFT_TYPES.findIndex(([t]) => t === a.who?.type) - GIFT_TYPES.findIndex(([t]) => t === b.who?.type) || (a.group || '').localeCompare(b.group || '', 'zh-Hant') || giftWho(a).localeCompare(giftWho(b), 'zh-Hant'));
  const items = ev.items || [];
  const tq = items.reduce((a, it) => a + (it.qty || 0), 0), tu = items.reduce((a, it) => a + (cnt[it.id]?.plan || 0) + (cnt[it.id]?.done || 0), 0), invSum = `${items.length} 項・共 ${tq} 份・剩 ${tq - tu}`;
  const inv = `<details class="sec ginv"${S.ginvOpen || S.gedit ? ' open' : ''}><summary><h2 style="display:inline">品項與庫存 <span class="sub">${invSum}</span></div></summary><div class="ginvbar">${S.gedit ? `<input class="nctxt" data-gevf="name" value="${esc(ev.name)}" style="max-width:260px" aria-label="活動名稱">` : ''} <span style="display:flex;gap:6px"><label class="btn sm">一次放圖片<input type="file" id="gImgs" accept="image/*,.zip,application/zip" multiple hidden></label><button class="btn sm${S.gedit ? ' pri' : ''}" data-act="geditems">${S.gedit ? '✓ 改好了' : '✎ 編修品項'}</button></span></h2>
   ${S.gedit ? `<div class="note">改品項名稱、代號、數量。按住左邊 ⠿ 拖曳排順序：拖分類那一列可以整個分類一起移動，把品項拖到別的分類底下就會換分類。刪除要按兩次。</div>` : ''}
   ${items.length ? `<table class="gtable"><thead><tr><th></th>${S.gedit ? '<th>分類</th>' : ''}<th>品項</th><th>代號</th><th>數量</th><th>預備</th><th>已送</th><th>剩下</th>${S.gedit ? '<th></th>' : ''}</tr></thead>
    ${S.gedit ? giftCatsSorted(items).map(([cat, its]) => `<tbody class="ncgrp ncbody" data-gcat="${esc(cat)}"><tr class="gcat"><td colspan="9"><span class="ncdrag ncgdrag" title="按住拖曳整個分類">⠿</span> ${esc(cat || '未分類')} <span class="sub">${its.length} 項</span></td></tr>${its.map((it) => { const c = cnt[it.id] || { plan: 0, done: 0 }, left = (it.qty || 0) - c.plan - c.done; return `<tr class="ncrow" data-iid="${it.id}"><td><span class="ncdrag" title="按住拖曳">⠿</span>${giftThumb(it, true)}</td><td><input class="nctxt" data-gitem="${it.id}" data-gf="cat" value="${esc(it.cat)}" placeholder="例如：彩豐月曆" style="width:120px"></td><td><input class="nctxt" data-gitem="${it.id}" data-gf="name" value="${esc(it.name)}" placeholder="例如：小熊維尼"></td><td><input class="nctxt" data-gitem="${it.id}" data-gf="code" value="${esc(it.code)}" style="width:110px"></td><td><input class="nctxt" data-gitem="${it.id}" data-gf="qty" inputmode="numeric" value="${it.qty || 0}" style="width:64px"></td><td>${c.plan}</td><td>${c.done}</td><td>${left}</td><td><button class="btn sm" data-gidel="${it.id}">✕</button></td></tr>`; }).join('')}</tbody>`).join('') + '<tbody>' : '<tbody>' + giftItemsSorted(items).map((it, ix, arr) => { const c = cnt[it.id] || { plan: 0, done: 0 }, left = (it.qty || 0) - c.plan - c.done; const catRow = !S.gedit && (ix === 0 || (arr[ix - 1].cat || '') !== (it.cat || '')) ? giftCatRow(it.cat || '', arr, cnt) : ''; return S.gedit
      ? `<tr><td>${giftThumb(it, true)}</td><td><input class="nctxt" data-gitem="${it.id}" data-gf="cat" value="${esc(it.cat)}" placeholder="例如：彩豐月曆" style="width:120px"></td><td><input class="nctxt" data-gitem="${it.id}" data-gf="name" value="${esc(it.name)}" placeholder="例如：小熊維尼"></td><td><input class="nctxt" data-gitem="${it.id}" data-gf="code" value="${esc(it.code)}" style="width:110px"></td><td><input class="nctxt" data-gitem="${it.id}" data-gf="qty" inputmode="numeric" value="${it.qty || 0}" style="width:64px"></td><td>${c.plan}</td><td>${c.done}</td><td>${left}</td><td><button class="btn sm" data-gidel="${it.id}">✕</button></td></tr>`
      : catRow + `<tr><td>${giftThumb(it)}</td><td>${esc(it.name || it.code || '（未命名）')}</td><td class="sub">${esc(it.code)}</td><td>${it.qty || 0}</td><td>${c.plan}</td><td>${c.done}</td><td class="${left < 0 ? 'gneg' : ''}">${left}</td></tr>`; }).join('')}</tbody></table>` : '<div class="sub">還沒有品項。按「編修品項」新增。</div>'}
   ${S.gedit ? `<div class="actions" style="margin:0"><button class="btn sm" data-act="gaddi">＋ 新增品項</button><button class="btn sm" data-act="gbulki">貼上一批品項</button><button class="btn sm danger" data-act="gevdel">刪除這個活動</button></div>` : ''}
   ${ev.note || S.gedit ? `<label class="f">備註${S.gedit ? `<input class="nctxt" data-gevf="note" value="${esc(ev.note)}">` : `<span>${esc(ev.note)}</span>`}</label>` : ''}</details>`;
  return head + `  <div class="sec"><h2>送禮名單 <span class="sub">${all.length} 位・已送 ${all.filter((g) => g.status === '已送').length}</span><span style="display:flex;gap:6px;flex-wrap:wrap"><span class="seg2"><button class="btn sm${S.gview === 'table' ? ' pri' : ''}" data-gview="table">表格</button><button class="btn sm${S.gview === 'list' ? ' pri' : ''}" data-gview="list">清單</button></span><button class="btn sm" data-act="gbulkw">貼上一批名單</button><button class="btn sm" data-act="gpickw">從客戶、服務件勾選</button><button class="btn sm pri" data-act="gaddwho">＋ 加一位</button></span></h2>
   <div class="chips">${[['全部', '全部'], ...GIFT_TYPES].map(([k, l]) => `<button class="chip${S.gtype === k ? ' on' : ''}" data-gtype="${k}">${l} <span class="cnt">${k === '全部' ? all.length : all.filter((g) => g.who?.type === k).length}</span></button>`).join('')}</div>
   <div class="chips">${['全部', '預備', '已送'].map((k) => `<button class="chip${S.gst === k ? ' on' : ''}" data-gst="${k}">${k} <span class="cnt">${k === '全部' ? all.length : all.filter((g) => (g.status || '預備') === k).length}</span></button>`).join('')}</div>
   ${list.length && S.gview === 'table' ? giftTableHTML(list, ev) : list.length ? `<div class="glist">${list.map((g) => `<div class="grow2${g.status === '已送' ? ' done' : ''}"><button class="gmain" data-gopen="${g.id}"><span class="t"><b>${esc(giftWho(g))}</b><span class="pill p-go">${GIFT_TYPES.find(([t]) => t === g.who?.type)?.[1] || '親友'}</span>${g.group ? `<span class="pill">${esc(g.group)}</span>` : ''}${g.method ? `<span class="sub">${esc(g.method)}</span>` : ''}</span><span class="s">${esc(giftItemsText(g, ev)) || '還沒選品項'}${g.note ? ' · ' + esc(g.note) : ''}</span></button>
     <button class="btn sm${g.status === '已送' ? ' pri' : ''}" data-gtog="${g.id}">${g.status === '已送' ? '✓ 已送' : '預備中'}</button></div>`).join('')}</div>` : '<div class="sub">這裡還沒有人。</div>'}</div>
  ${inv}`;
}
function giftSortWire() { // 編修品項時拖曳排順序
  const tb = app.querySelector('.ginv .gtable'); if (!tb) return; tb.setAttribute('data-sec', 'gi');
  makeSortable(tb, () => { const ev = Vault.get('gevents', S.gev); if (!ev) return; const m = Object.fromEntries(ev.items.map((it) => [it.id, it])), out = [];
    tb.querySelectorAll('tbody.ncgrp').forEach((bd) => bd.querySelectorAll('.ncrow').forEach((r) => { const it = m[r.dataset.iid]; if (it) { it.cat = bd.dataset.gcat; out.push(it); delete m[it.id]; } }));
    ev.items = [...out, ...Object.values(m)]; Vault.put('gevents', ev); App.render(); });
}
function giftSectionHTML(type, id) { // 服務件、客戶頁上的送禮紀錄，跟送禮專區是同一份資料
  const mine = Vault.all('gifts').filter((g) => g.who?.type === type && g.who?.id === id);
  return `<div class="sec"><h2>送禮 <span class="sub">跟送禮專區同步</span><button class="btn sm" data-gnew="${type}:${id}">＋ 預備贈送</button></h2>
   ${mine.length ? mine.map((g) => { const ev = Vault.get('gevents', g.eventId); return `<div class="grow2${g.status === '已送' ? ' done' : ''}"><button class="gmain" data-gopen="${g.id}"><span class="t"><b>${esc(ev?.name || '已刪活動')}</b>${g.method ? `<span class="sub">${esc(g.method)}</span>` : ''}</span><span class="s">${esc(giftItemsText(g, ev)) || '還沒選品項'}</span></button><button class="btn sm${g.status === '已送' ? ' pri' : ''}" data-gtog="${g.id}">${g.status === '已送' ? '✓ 已送' : '預備中'}</button></div>`; }).join('') : '<div class="sub">還沒有送禮紀錄。</div>'}</div>`;
}
function openGiftEditor(g, preset) { // 新增或修改一筆送禮
  const evs = gevents(); if (!evs.length) { toast('先到送禮專區建一個活動'); S.view = 'gifts'; App.render(); return; }
  const rec = g ? JSON.parse(JSON.stringify(g)) : { eventId: S.gev && Vault.get('gevents', S.gev) ? S.gev : evs[0].id, who: preset || { type: 'call', id: '', name: '' }, group: '', items: {}, method: '', status: '預備', note: '' };
  const fixedWho = !!(preset && preset.id) || !!g;
  const draw = () => {
    const ev = Vault.get('gevents', rec.eventId), cnt = ev ? giftCounts(ev) : {}, t = rec.who.type, q = ($('gwQ')?.value || '').trim();
    const pool = t === 'client' ? clients() : t === 'call' ? calls() : [];
    const found = q ? pool.filter((x) => (x.name || '').includes(q)).slice(0, 12) : [];
    ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>${g ? '修改送禮' : '預備贈送'}</h2>${g ? '<button class="btn danger" id="gwDel">刪除</button>' : ''}<button class="btn" id="gwX">取消</button><button class="btn pri" id="gwOk">儲存</button></div>
    <div class="ovbody" style="flex-direction:column;max-width:680px;margin:0 auto;width:100%">
     <label class="f">活動<select id="gwEv">${evs.map((x) => `<option value="${x.id}"${x.id === rec.eventId ? ' selected' : ''}>${esc(x.name || '未命名')}</option>`).join('')}</select></label>
     ${fixedWho ? `<div class="f">對象<div><b>${esc(giftWho(rec))}</b> <span class="pill p-go">${GIFT_TYPES.find(([k]) => k === t)?.[1]}</span></div></div>` : `
     <div class="f">對象是<div class="tags">${GIFT_TYPES.map(([k, l]) => `<button class="tag${t === k ? ' on' : ''}" data-gwt="${k}">${l}</button>`).join('')}</div></div>
     ${t === 'other' ? `<label class="f">名字<input id="gwName" value="${esc(rec.who.name)}" placeholder="例如：叔叔、年底聚餐"></label>`
      : `<label class="f">搜尋${t === 'client' ? '客戶' : '服務件'}<input id="gwQ" value="${esc(q)}" placeholder="打名字" autocomplete="off"></label>
      ${rec.who.id ? `<div class="sub">已選：<b>${esc(giftWho(rec))}</b></div>` : ''}<div class="list" style="padding:0">${found.map((x) => `<button class="item" data-gwp="${x.id}"><span class="t">${esc(x.name || '未填姓名')}</span></button>`).join('')}</div>`}`}
     <label class="f">分類（可不填）<input id="gwGroup" list="gwGroups" value="${esc(rec.group)}" placeholder="例如：朋友、張家、石家"><datalist id="gwGroups">${giftGroups().map((x) => `<option value="${esc(x)}">`).join('')}</datalist></label>
     <div class="f">品項和數量<div class="gqty">${giftItemsSorted(ev?.items || []).map((it, ix, arr) => (ix === 0 || (arr[ix - 1].cat || '') !== (it.cat || '') ? `<div class="gqcat">${esc(it.cat || '未分類')}</div>` : '') + (() => { const c = cnt[it.id] || { plan: 0, done: 0 }; const left = (it.qty || 0) - c.plan - c.done + ((g && g.eventId === ev.id && g.items?.[it.id]) || 0); return `<div class="gqrow"><span class="gqname">${giftThumb(it)}<span>${esc(it.name || it.code || '品項')} <span class="sub">剩 ${left}</span></span></span><span class="gstep"><button class="btn sm" data-gq="${it.id}:-1">－</button><b>${rec.items[it.id] || 0}</b><button class="btn sm" data-gq="${it.id}:1">＋</button></span></div>`; })()).join('') || '<span class="sub">這個活動還沒有品項，先到送禮專區按「編修品項」。</span>'}</div></div>
     <div class="f">怎麼送<div class="tags">${['當面', '寄送', '放管理室'].map((m) => `<button class="tag${rec.method === m ? ' on' : ''}" data-gwm="${m}">${m}</button>`).join('')}</div></div>
     <div class="f">狀態<div class="tags">${['預備', '已送'].map((m) => `<button class="tag${(rec.status || '預備') === m ? ' on' : ''}" data-gws="${m}">${m}</button>`).join('')}</div></div>
     <label class="f">備註<input id="gwNote" value="${esc(rec.note)}"></label></div></div>`;
    ov.querySelectorAll('img[data-fid]').forEach(async (im) => { const u = await Files.url(im.dataset.fid).catch(() => null); if (u && im.isConnected) im.src = u; });
    const keep = () => { rec.group = $('gwGroup').value.trim(); rec.note = $('gwNote').value.trim(); if ($('gwName')) rec.who.name = $('gwName').value.trim(); };
    $('gwX').onclick = () => { ov.innerHTML = ''; };
    $('gwEv').onchange = (e) => { keep(); if (e.target.value !== rec.eventId) { rec.eventId = e.target.value; rec.items = {}; } draw(); };
    if ($('gwQ')) { $('gwQ').oninput = () => { keep(); const v = $('gwQ').value; draw(); const el = $('gwQ'); el.focus(); el.setSelectionRange(v.length, v.length); }; }
    ov.querySelectorAll('[data-gwt]').forEach((b) => (b.onclick = () => { keep(); rec.who = { type: b.dataset.gwt, id: '', name: '' }; draw(); }));
    ov.querySelectorAll('[data-gwp]').forEach((b) => (b.onclick = () => { keep(); const x = Vault.get(t === 'client' ? 'clients' : 'calls', b.dataset.gwp); rec.who = { type: t, id: x.id, name: x.name || '' }; $('gwQ').value = ''; draw(); }));
    ov.querySelectorAll('[data-gq]').forEach((b) => (b.onclick = () => { keep(); const [id, dlt] = b.dataset.gq.split(':'); rec.items[id] = Math.max(0, (rec.items[id] || 0) + +dlt); draw(); }));
    ov.querySelectorAll('[data-gwm]').forEach((b) => (b.onclick = () => { keep(); rec.method = rec.method === b.dataset.gwm ? '' : b.dataset.gwm; draw(); }));
    ov.querySelectorAll('[data-gws]').forEach((b) => (b.onclick = () => { keep(); rec.status = b.dataset.gws; draw(); }));
    if ($('gwDel')) $('gwDel').onclick = (e) => { if (!e.currentTarget.dataset.sure) { e.currentTarget.dataset.sure = 1; e.currentTarget.textContent = '再按一次刪除'; return; } Vault.del('gifts', g.id); ov.innerHTML = ''; App.render(); };
    $('gwOk').onclick = () => {
      keep(); if (rec.who.type === 'other' ? !rec.who.name : !rec.who.id) { toast(rec.who.type === 'other' ? '先填名字' : '先搜尋並選一位'); return; }
      Object.keys(rec.items).forEach((k) => { if (!rec.items[k]) delete rec.items[k]; });
      if (rec.status === '已送' && !rec.doneDate) rec.doneDate = localDay(); if (rec.status !== '已送') rec.doneDate = '';
      Vault.put('gifts', g ? Object.assign(g, rec) : rec); ov.innerHTML = ''; App.render(); toast('已儲存');
    };
  };
  draw();
}
function openGiftBulkItems() { // 一行一個品項：分類／名稱／代號／數量
  const ev = Vault.get('gevents', S.gev); if (!ev) return;
  ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>貼上一批品項</h2><button class="btn" id="biX">取消</button><button class="btn pri" id="biOk">加入</button></div>
  <div class="ovbody" style="flex-direction:column;max-width:680px;margin:0 auto;width:100%">
   <div class="note">一行一個品項，用斜線隔開：<b>分類／名稱／代號／數量</b>。沒有名稱可以空著，例如「彩豐月曆／／A09／20」。</div>
   <textarea id="biT" rows="14" placeholder="達標桌曆／小熊維尼／CAWP-014／5"></textarea><div class="sub" id="biN"></div></div></div>`;
  const parse = () => $('biT').value.split('\n').map((l) => l.trim()).filter(Boolean).map((l) => { const c = l.split(/[／/|｜]/).map((x) => x.trim()); return { id: uid(), cat: c[0] || '', name: c[1] || '', code: c[2] || '', qty: parseInt(c[3], 10) || 0 }; }).filter((x) => x.name || x.code);
  $('biT').oninput = () => { const n = parse().length; $('biN').textContent = n ? `會加入 ${n} 個品項` : ''; };
  $('biX').onclick = () => (ov.innerHTML = '');
  $('biOk').onclick = () => { const its = parse(); if (!its.length) { toast('先貼上品項'); return; } ev.items = [...(ev.items || []), ...its]; Vault.put('gevents', ev); S.gedit = false; ov.innerHTML = ''; App.render(); toast(`已加入 ${its.length} 個品項`); };
}
function giftWhoMatch(name) { // 名字跟客戶或服務件一樣，就自動連過去
  const cl = clients().find((x) => x.name === name); if (cl) return { type: 'client', id: cl.id, name };
  const k = calls().find((x) => x.name === name); if (k) return { type: 'call', id: k.id, name };
  return { type: 'other', id: '', name };
}
const giftHas = (evId, who) => giftsOf(evId).some((g) => g.who?.type === who.type && (who.id ? g.who.id === who.id : g.who.name === who.name));
function openGiftBulkWho() { // 一行一位：名字／分類／備註／已送
  const ev = Vault.get('gevents', S.gev); if (!ev) return;
  ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>貼上一批名單</h2><button class="btn" id="bwX">取消</button><button class="btn pri" id="bwOk">加入</button></div>
  <div class="ovbody" style="flex-direction:column;max-width:680px;margin:0 auto;width:100%">
   <div class="note">一行一位，用斜線隔開：<b>名字／分類／備註／已送</b>，後面三個可以不寫。名字跟客戶或服務件一模一樣的，會自動連過去；其他的放在「親友」。品項之後再點進去選。</div>
   <textarea id="bwT" rows="14" placeholder="叔叔／石家／去年 6 套"></textarea><div id="bwP"></div></div></div>`;
  const parse = () => $('bwT').value.split('\n').map((l) => l.trim()).filter(Boolean).map((l) => { const c = l.split(/[／/|｜]/).map((x) => x.trim()); const done = c.slice(1).some((x) => x === '已送'); const rest = c.slice(1).filter((x) => x !== '已送'); return { who: giftWhoMatch(c[0]), group: rest[0] || '', note: rest[1] || '', status: done ? '已送' : '預備' }; }).filter((x) => x.who.name);
  const lab = (t) => GIFT_TYPES.find(([k]) => k === t)?.[1];
  $('bwT').oninput = () => { const rs = parse(); $('bwP').innerHTML = rs.length ? `<div class="sub">會加入 ${rs.filter((r) => !giftHas(ev.id, r.who)).length} 位${rs.some((r) => giftHas(ev.id, r.who)) ? '（已經在名單裡的會略過）' : ''}：</div><div class="tags">${rs.map((r) => `<span class="tag" style="cursor:default">${esc(r.who.name)} <span class="sub">${lab(r.who.type)}${r.group ? '・' + esc(r.group) : ''}${r.status === '已送' ? '・已送' : ''}</span></span>`).join('')}</div>` : ''; };
  $('bwX').onclick = () => (ov.innerHTML = '');
  $('bwOk').onclick = () => { const rs = parse().filter((r) => !giftHas(ev.id, r.who)); if (!rs.length) { toast('沒有新的人可以加'); return; } rs.forEach((r) => Vault.put('gifts', { eventId: ev.id, who: r.who, group: r.group, items: {}, method: '', status: r.status, doneDate: r.status === '已送' ? localDay() : '', note: r.note })); ov.innerHTML = ''; App.render(); toast(`已加入 ${rs.length} 位`); };
}
function openGiftPickWho() { // 從客戶或服務件直接勾，一次加很多位
  const ev = Vault.get('gevents', S.gev); if (!ev) return; let type = 'call', lv = '全部', ar = '全部'; const pick = new Set();
  const draw = () => {
    const q = ($('pwQ')?.value || '').trim(); const pool = (type === 'client' ? clients() : calls()).filter((x) => !giftHas(ev.id, { type, id: x.id }));
    const areas = type === 'call' ? [...new Set(calls().map((k) => k.area).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'zh-Hant')) : [];
    const list = pool.filter((x) => (!q || [x.name, x.no].join(' ').includes(q)) && (type !== 'call' || ((lv === '全部' || x.level === lv) && (ar === '全部' || x.area === ar)))).sort(type === 'call' ? noCmp : (a, b) => (a.name || '').localeCompare(b.name || '', 'zh-Hant'));
    ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>勾選要送的人</h2><span class="sub">已勾 ${pick.size} 位</span><button class="btn" id="pwX">取消</button><button class="btn pri" id="pwOk">加入 ${pick.size} 位</button></div>
    <div class="ovbody" style="flex-direction:column;max-width:720px;margin:0 auto;width:100%">
     <div class="tags">${[['call', '服務件'], ['client', '現有客戶']].map(([k, l]) => `<button class="tag${type === k ? ' on' : ''}" data-pwt="${k}">${l}</button>`).join('')}</div>
     ${type === 'call' ? `<div class="chips">${['全部', ...LEVELS].map((x) => `<button class="chip${lv === x ? ' on' : ''}" data-pwl="${x}">${x}</button>`).join('')}</div>${areas.length ? `<div class="chips">${['全部', ...areas].map((x) => `<button class="chip${ar === x ? ' on' : ''}" data-pwa="${esc(x)}">${esc(x)}</button>`).join('')}</div>` : ''}` : ''}
     <input class="search" id="pwQ" value="${esc(q)}" placeholder="搜尋名字${type === 'call' ? '或序號' : ''}" autocomplete="off">
     <div class="actions" style="margin:0"><button class="btn sm" id="pwAll">這些全選</button><button class="btn sm" id="pwNone">全部不選</button></div>
     <div class="pwlist">${list.map((x) => `<label class="ncit"><input type="checkbox" data-pw="${x.id}"${pick.has(x.id) ? ' checked' : ''}><span>${x.no ? `<span class="kno">${esc(x.no)}</span>` : ''}${esc(x.name || '未填姓名')}${x.level ? ` <b class="lvl lv-${esc(x.level)}">${esc(x.level)}</b>` : ''}${x.area ? ` <span class="sub">${esc(x.area)}</span>` : ''}</span></label>`).join('') || '<div class="sub">沒有可以加的人（已經在名單裡的不會出現）。</div>'}</div></div></div>`;
    $('pwX').onclick = () => (ov.innerHTML = '');
    $('pwQ').oninput = () => { const v = $('pwQ').value; draw(); const el = $('pwQ'); el.focus(); el.setSelectionRange(v.length, v.length); };
    ov.querySelectorAll('[data-pwt]').forEach((b) => (b.onclick = () => { type = b.dataset.pwt; pick.clear(); draw(); }));
    ov.querySelectorAll('[data-pwl]').forEach((b) => (b.onclick = () => { lv = b.dataset.pwl; draw(); }));
    ov.querySelectorAll('[data-pwa]').forEach((b) => (b.onclick = () => { ar = b.dataset.pwa; draw(); }));
    ov.querySelectorAll('[data-pw]').forEach((c) => (c.onchange = () => { c.checked ? pick.add(c.dataset.pw) : pick.delete(c.dataset.pw); ov.querySelector('.ovhead .sub').textContent = `已勾 ${pick.size} 位`; $('pwOk').textContent = `加入 ${pick.size} 位`; }));
    $('pwAll').onclick = () => { list.forEach((x) => pick.add(x.id)); draw(); };
    $('pwNone').onclick = () => { pick.clear(); draw(); };
    $('pwOk').onclick = () => { if (!pick.size) { toast('先勾要送的人'); return; } [...pick].forEach((id) => { const x = Vault.get(type === 'client' ? 'clients' : 'calls', id); if (x) Vault.put('gifts', { eventId: ev.id, who: { type, id, name: x.name || '' }, group: '', items: {}, method: '', status: '預備', note: '' }); }); ov.innerHTML = ''; App.render(); toast(`已加入 ${pick.size} 位`); };
  };
  draw();
}
function openNewGiftEvent() { // 新活動：可以從舊活動複製品項（和名單）
  const evs = gevents();
  ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>新活動</h2><button class="btn" id="neX">取消</button><button class="btn pri" id="neOk">建立</button></div>
  <div class="ovbody" style="flex-direction:column;max-width:600px;margin:0 auto;width:100%">
   <label class="f">活動名稱<input id="neName" placeholder="例如：2027 年曆、中秋、日本伴手禮"></label>
   ${evs.length ? `<label class="f">從舊活動複製（可不選）<select id="neFrom"><option value="">不用，從空白開始</option>${evs.map((x) => `<option value="${x.id}">${esc(x.name || '未命名')}</option>`).join('')}</select></label>
   <label class="ncit"><input type="checkbox" id="neItems" checked> 複製品項（數量歸零，再自己填）</label>
   <label class="ncit"><input type="checkbox" id="neWho"> 名單也一起複製（全部改成預備、不帶品項）</label>` : ''}</div></div>`;
  $('neX').onclick = () => (ov.innerHTML = '');
  $('neOk').onclick = () => {
    const name = $('neName').value.trim(); if (!name) { toast('先取個名字'); return; }
    const from = $('neFrom')?.value && Vault.get('gevents', $('neFrom').value);
    const ev = Vault.put('gevents', { name, note: '', items: from && $('neItems').checked ? from.items.map((it) => ({ id: uid(), cat: it.cat || '', name: it.name, code: it.code, qty: 0 })) : [] });
    if (from && $('neWho').checked) giftsOf(from.id).forEach((g) => Vault.put('gifts', { eventId: ev.id, who: { ...g.who }, group: g.group || '', items: {}, method: g.method || '', status: '預備', note: '' }));
    S.gev = ev.id; S.gedit = !ev.items.length; ov.innerHTML = ''; App.render(); toast('已建立「' + name + '」');
  };
  setTimeout(() => $('neName')?.focus(), 50);
}
/* 一個案件多位要保人／被保險人 */
const ROLES = ['要保人', '被保險人', '受益人', '授權人', '其他'];
function casePeople(c) { // 包含主要客戶
  const ppl = (c.people || []).filter((p) => Vault.get('clients', p.clientId));
  if (c.clientId && Vault.get('clients', c.clientId) && !ppl.some((p) => p.clientId === c.clientId)) ppl.unshift({ clientId: c.clientId, roles: [], primary: true });
  return ppl;
}
function peopleHTML(c) {
  if (!['新契約', '建議書', '保全', '產險'].includes(c.cat) && !(c.people || []).length) return '';
  const ppl = c.people || [];
  const main = c.clientId && !ppl.some((p) => p.clientId === c.clientId) ? Vault.get('clients', c.clientId) : null;
  return `<div class="sec"><h2>要保人／被保險人 <button class="btn" data-act="addperson" style="padding:4px 10px;font-size:13px">＋ 加人</button></h2>
   ${main ? `<div class="sub">主要關係人「${esc(main.name)}」也可以按「＋ 加人」加進來標角色。</div>` : ''}
   ${ppl.length ? ppl.map((p, i) => { const x = Vault.get('clients', p.clientId); if (!x) return ''; return `<div class="prow"><button class="linkbtn relname" data-goclient="${x.id}">${esc(x.name || '未填姓名')}</button><span class="tags">${ROLES.map((r) => `<button class="tag sm${p.roles.includes(r) ? ' on' : ''}" data-prole="${i}:${r}">${r}</button>`).join('')}</span><input class="pprod" data-pprod="${i}" value="${esc(p.product)}" placeholder="險種／商品"><button class="tldel" data-pdel="${i}" aria-label="移除">✕</button></div>`; }).join('')
    : '<div class="sub">一個家庭一起投保時，把每個人加進來，標出誰是要保人、誰是被保險人。案件會出現在每個人的客戶頁。</div>'}</div>`;
}
function openPersonPicker(c) {
  const fam = new Set(); const main = c.clientId;
  if (main) rels().forEach((r) => { if (FAMILY.includes(r.type)) { if (r.a === main) fam.add(r.b); if (r.b === main) fam.add(r.a); } });
  const inCase = new Set((c.people || []).map((p) => p.clientId));
  const draw = () => {
    const q = ($('ppQ')?.value || '').trim();
    const pool = clients().filter((x) => !inCase.has(x.id) && (!q || (x.name || '').includes(q))).sort((a, b) => (b.id === main) - (a.id === main) || fam.has(b.id) - fam.has(a.id) || (a.name || '').localeCompare(b.name || '', 'zh-Hant')).slice(0, 40);
    $('ppList').innerHTML = (q && !clients().some((x) => x.name === q) ? `<button class="item" data-newp="1"><span class="t">＋ 新增「${esc(q)}」</span></button>` : '') + pool.map((x) => `<button class="item" data-pp="${x.id}"><span class="t"><span>${esc(x.name || '未填姓名')}</span>${x.id === main ? '<span class="pill p-go">主要關係人</span>' : fam.has(x.id) ? '<span class="pill p-go">家人</span>' : ''}</span></button>`).join('');
  };
  ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>加入要保人／被保險人</h2><button class="btn pri" id="ppX">完成</button></div>
  <div class="ovbody" style="flex-direction:column;max-width:640px;margin:0 auto;width:100%">
   <div class="sub">先選角色，再點人（可以連續加好幾位）</div>
   <div class="tags" id="ppRoles">${ROLES.map((r, i) => `<button class="tag${i === 1 ? ' on' : ''}" data-r="${r}">${r}</button>`).join('')}</div>
   <input class="search" id="ppQ" placeholder="搜尋或輸入名字" autocomplete="off"><div class="list" id="ppList" style="padding:0"></div></div></div>`;
  let roles = ['被保險人'];
  $('ppRoles').onclick = (e) => { const b = e.target.closest('[data-r]'); if (!b) return; b.classList.toggle('on'); roles = [...$('ppRoles').querySelectorAll('.tag.on')].map((x) => x.dataset.r); };
  $('ppX').onclick = () => { ov.innerHTML = ''; App.render(); };
  $('ppQ').oninput = draw;
  $('ppList').onclick = (e) => {
    const b = e.target.closest('button'); if (!b) return;
    let x = b.dataset.pp ? Vault.get('clients', b.dataset.pp) : null;
    if (b.dataset.newp) x = Vault.put('clients', { name: $('ppQ').value.trim(), docs: [] });
    if (!x) return;
    c.people = [...(c.people || []), { clientId: x.id, roles: [...roles], product: '' }]; inCase.add(x.id); saveCase(c);
    $('ppQ').value = ''; draw(); toast(`已加入 ${x.name}（${roles.join('、') || '未標角色'}）`);
  };
  draw();
}
function openWhoPicker(c, ppl) {
  ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>存誰的告知資料？</h2><button class="btn" id="wpX">取消</button></div>
  <div class="ovbody" style="flex-direction:column;max-width:520px;margin:0 auto;width:100%"><div class="list" style="padding:0">${ppl.map((p) => { const x = Vault.get('clients', p.clientId); return `<button class="item" data-who="${x.id}"><span class="t"><span>${esc(x.name || '未填姓名')}</span><span class="sub">${esc((p.roles || []).join('、'))}</span></span></button>`; }).join('')}</div></div></div>`;
  $('wpX').onclick = () => (ov.innerHTML = '');
  ov.querySelectorAll('[data-who]').forEach((b) => (b.onclick = () => openDisclosure(Vault.get('clients', b.dataset.who), null, c.id)));
}
// 客戶頁的案件列：列出這件案子裡的每個人和角色，看得出是誰的案件
function casePeopleLine(c, meId) {
  const ppl = (c.people || []).filter((p) => Vault.get('clients', p.clientId));
  const list = [];
  if (c.clientId && !ppl.some((p) => p.clientId === c.clientId)) list.push({ id: c.clientId, name: c.client, roles: ['主要關係人'] });
  ppl.forEach((p) => list.push({ id: p.clientId, name: Vault.get('clients', p.clientId).name || '未填姓名', roles: p.roles.length ? p.roles : (p.clientId === c.clientId ? ['主要關係人'] : []) }));
  if (!list.length || (list.length === 1 && list[0].id === meId)) return '';
  return `<span class="ppline">${list.map((p) => `<span class="${p.id === meId ? 'me' : ''}">${esc(p.name)}${p.roles.length ? `<small>${esc(p.roles.join('、'))}</small>` : ''}</span>`).join('')}</span>`;
}
function clientCases(cl) { return cases().filter((c) => c.clientId === cl.id || (c.people || []).some((p) => p.clientId === cl.id)).sort(byDate); }
/* ================= 從 CRM 帶入（截圖在這台裝置上辨識，不會傳出去） ================= */
function parseCRM(raw) {
  let t = String(raw || '').replace(/\r/g, '')
    .replace(/[！-～]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xFEE0)).replace(/　/g, ' ');
  const tight = (s) => s.replace(/\s+/g, '');
  const LABELS = '姓名|性別|手機|生日|出生|住家電話|公司電話|LINE ?ID|現居地址|公司地址|收費地址|通訊地址|戶籍地址|Email|E-mail|ID|身分證|年齡|編輯';
  const lbl = (s) => s.split('').join('\\s*');
  const LBL_RE = new RegExp('(' + ['姓名', '性別', '手機', '生日', '出生日期', '住家電話', '公司電話', '現居地址', '公司地址', '收費地址', '通訊地址', '戶籍地址', '年齡', '編輯資料'].map(lbl).join('|') + '|LINE\\s*ID|E-?mail)', 'gi');
  t = t.replace(LBL_RE, '\n$1'); // 每個欄位名稱前斷行，避免黏到上一個值
  const out = {};
  const nm = t.match(new RegExp(lbl('姓名') + '[\\s:：]*((?:[\\u4e00-\\u9fff][ \\t]?){2,4})'));
  if (nm) out.name = tight(nm[1]).replace(new RegExp('(' + LABELS + ').*$'), '').slice(0, 4);
  const g = t.match(new RegExp(lbl('性別') + '[\\s:：]*([男女])')); if (g) out.gender = g[1];
  const id = t.match(/\b([A-Z][12]\d{8})\b/); if (id) out.idNo = id[1];
  const bd = t.match(new RegExp('(?:' + lbl('生日') + '|' + lbl('出生日期') + ')[\\s:：]*(\\d{2,4})\\s*[\\/.\\-年]\\s*(\\d{1,2})\\s*[\\/.\\-月]\\s*(\\d{1,2})'));
  if (bd) { let y = +bd[1]; if (y < 1911) y += 1911; out.birthday = `${y}-${String(bd[2]).padStart(2, '0')}-${String(bd[3]).padStart(2, '0')}`; }
  const ph = t.match(/(09\d{2})[\s-]?(\d{3})[\s-]?(\d{3})/); if (ph) out.phone = ph[1] + ph[2] + ph[3];
  const tel = (label) => { const m = t.match(new RegExp(lbl(label) + '[\\s:：]*(\\(?0\\d{1,2}\\)?[\\s-]?\\d{3,4}[\\s-]?\\d{3,4}(?:\\s*#\\s*\\d+)?)')); return m ? m[1].replace(/\s/g, '') : ''; };
  out.homeTel = tel('住家電話'); out.workTel = tel('公司電話');
  const em = t.match(/([A-Za-z0-9._%+-]+)\s*@\s*([A-Za-z0-9.-]+)/);
  if (em) { let dom = em[2].replace(/(com)(tw)$/i, '$1.$2'); if (!/\./.test(dom)) dom = dom.replace(/[Ll1I]?(com|net|org)$/, '.$1'); out.email = em[1] + '@' + dom.replace(/\.$/, ''); }
  const line = t.match(/LINE\s*ID[\s:：]*([A-Za-z0-9._\-]{3,})/i); if (line) out.line = line[1];
  const addr = (label) => { const m = t.match(new RegExp(lbl(label) + '[\\s:：]*(\\d{3,6}\\s*[\\u4e00-\\u9fff][^\\n]*)')); if (!m) return ''; let a = tight(m[1]); const end = a.search(/(號|樓|室)(?!.*(號|樓|室))/); if (end >= 0) a = a.slice(0, end + 1).replace(/(之\d+)?$/, (x) => x); return a; };
  out.address = addr('現居地址') || addr('通訊地址') || addr('收費地址') || addr('戶籍地址');
  out.billAddress = addr('收費地址'); out.workAddress = addr('公司地址');
  Object.keys(out).forEach((k) => { if (!out[k]) delete out[k]; });
  return out;
}

const CRM_FIELDS = [['name', '姓名'], ['gender', '性別'], ['birthday', '生日'], ['idNo', '身分證字號'], ['phone', '手機'], ['homeTel', '住家電話'], ['workTel', '公司電話'], ['line', 'LINE ID'], ['email', 'Email'], ['address', '地址'], ['billAddress', '收費地址'], ['workAddress', '公司地址']];
let OCR = null, CRMI = null;
function loadScript(src) { return new Promise((ok, bad) => { const s = document.createElement('script'); s.src = src; s.onload = ok; s.onerror = () => bad(new Error('載入失敗')); document.head.appendChild(s); }); }
async function ocrWorker() {
  if (OCR) return OCR;
  const abs = (p) => new URL(p, location.href).href;
  if (!window.Tesseract) await loadScript('vendor/ocr/tesseract.min.js');
  OCR = await Tesseract.createWorker(['chi_tra', 'eng'], 1, { workerPath: abs('vendor/ocr/worker.min.js'), corePath: abs('vendor/ocr'), langPath: abs('vendor/ocr/lang'), gzip: true, workerBlobURL: false });
  return OCR;
}
async function ocrImage(file) {
  const bmp = await createImageBitmap(file);
  const k = Math.min(2, 3600 / Math.max(bmp.width, bmp.height));
  const c = document.createElement('canvas'); c.width = Math.round(bmp.width * k); c.height = Math.round(bmp.height * k);
  const x = c.getContext('2d'); x.imageSmoothingQuality = 'high'; x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height); x.filter = 'grayscale(1)'; x.drawImage(bmp, 0, 0, c.width, c.height);
  const w = await ocrWorker(); const { data } = await w.recognize(c); return data.text;
}
function openCrmImport(target) {
  CRMI = { target, shot: null };
  ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>從 CRM 帶入${target ? '「' + esc(target.name || '這位客戶') + '」' : ''}</h2><button class="btn" id="ciX">取消</button><button class="btn pri" id="ciRead">讀取</button></div>
  <div class="ovbody" style="flex-direction:column;max-width:760px;margin:0 auto;width:100%">
   <div class="note">兩種都可以：<br>・在公司 CRM 截圖後，<b>在這裡按 Ctrl+V</b>（或「選擇截圖」「📋 貼上」）。會在這台裝置上辨識，不會傳到外面。<br>・或在 CRM 頁面把那一區的文字選起來複製，貼到下面，最準。</div>
   <div class="actions"><label class="btn">選擇截圖<input type="file" id="ciFile" accept="image/*"></label><button class="btn" id="ciPaste">📋 貼上</button><span class="sub" id="ciStatus"></span></div>
   <img id="ciShot" alt="" hidden style="max-width:100%;border:1px solid var(--line);border-radius:6px">
   <label class="f">文字（辨識結果會出現在這裡，可以直接修改）<textarea id="ciText" rows="7" placeholder="貼上文字或截圖"></textarea></label>
   <div id="ciForm"></div></div></div>`;
  $('ciX').onclick = closeCrmImport;
  $('ciRead').onclick = () => showCrmForm(parseCRM($('ciText').value));
  $('ciFile').onchange = (e) => { const f = e.target.files[0]; e.target.value = ''; if (f) crmFromImage(f); };
  $('ciPaste').onclick = async () => { try { const items = await navigator.clipboard.read(); for (const it of items) { const t = it.types.find((y) => y.startsWith('image/')); if (t) { crmFromImage(new File([await it.getType(t)], 'shot.png', { type: t })); return; } if (it.types.includes('text/plain')) { $('ciText').value = await (await it.getType('text/plain')).text(); showCrmForm(parseCRM($('ciText').value)); return; } } toast('剪貼簿是空的'); } catch (err) { toast('這裡不能讀剪貼簿，請直接按 Ctrl+V'); } };
  $('ciText').addEventListener('paste', () => setTimeout(() => showCrmForm(parseCRM($('ciText').value)), 30));
  setTimeout(() => $('ciText')?.focus(), 50);
}
let KPH = null; // 服務件拍照帶入
const nextCallNo = () => { const ns = calls().map((k) => parseInt(k.no, 10)).filter((n) => !isNaN(n)); return String(ns.length ? Math.max(...ns) + 1 : 1); };
const KPH_F = [['no', '序號'], ['name', '姓名'], ['phone', '手機'], ['home', '住家電話'], ['addr', '地址']];
function openCallPhoto(target) {
  KPH = { target };
  ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>${target ? '拍照帶入「' + esc(target.name || '這位') + '」' : '拍照新增服務件'}</h2><button class="btn" id="kpX">取消</button></div>
  <div class="ovbody" style="flex-direction:column;max-width:760px;margin:0 auto;width:100%">
   <div class="note">拍公司系統的「客戶資訊」畫面，或截圖後按「選照片」、Ctrl+V 貼上。文字在這台裝置上辨識，照片不會傳出去，也不會存起來。只帶入姓名、手機、住家電話、地址，身分證字號不帶。</div>
   <div class="actions"><label class="btn pri">📷 拍照／選照片<input type="file" id="kpFile" accept="image/*" hidden></label><span class="sub" id="kpSt"></span></div>
   <div id="kpForm"></div></div></div>`;
  $('kpX').onclick = () => { KPH = null; ov.innerHTML = ''; };
  $('kpFile').onchange = (e) => { const f = e.target.files[0]; e.target.value = ''; if (f) callPhotoFrom(f); };
  if (target) callPhotoForm({});
}
async function callPhotoFrom(f) {
  if (!KPH) return; $('kpSt').textContent = OCR ? '辨識中…' : '第一次使用要先下載辨識元件（約 13MB），請稍等…';
  try { const d = parseCRM(await ocrImage(f)); if (!KPH) return; $('kpSt').textContent = d.name || d.phone ? '辨識完成，請檢查下面的欄位' : '沒讀到資料，可以重拍，或直接在下面打'; callPhotoForm({ name: d.name, phone: d.phone, home: d.homeTel, addr: d.address || d.billAddress }); }
  catch (e) { console.warn(e); $('kpSt').textContent = '辨識失敗：' + (e.message || e); }
}
function callPhotoForm(d) {
  if (!KPH) return; const t = KPH.target;
  const same = !t && calls().find((k) => (d.phone && k.phone && k.phone.replace(/\D/g, '') === d.phone) || (d.name && k.name === d.name));
  const v = { ...d, no: t ? t.no : same ? same.no : nextCallNo() };
  $('kpForm').innerHTML = `<div class="sec"><h2>確認資料 <span class="sub">${t ? '空的欄位才會填進去，有值的會顯示「目前」讓你選' : same ? `名單裡已經有「${esc(same.name || '')}」（序號 ${esc(same.no || '—')}） <button class="linkbtn" id="kpUse">改成更新這位</button>` : '會新增一位'}</span></h2>
   ${KPH_F.map(([k, l]) => `<div class="cirow"><label class="cichk" style="min-width:80px">${l}</label><input class="ciin" id="kp_${k}" value="${esc(v[k] || (t && !d[k] ? t[k] : '') || '')}"${k === 'no' && t ? ' disabled' : ''}>${k === 'no' && !t ? '<span class="sub">接在名單最後，可以改</span>' : t && d[k] && t[k] && t[k] !== d[k] ? `<span class="sub">目前：${esc(t[k])}</span>` : ''}</div>`).join('')}
   <div class="note">地址、姓名請看一下，辨識偶爾會錯一兩個字。</div>
   <div class="actions"><button class="btn pri" id="kpGo">${t ? '帶入' : '新增'}</button></div></div>`;
  if ($('kpUse')) $('kpUse').onclick = () => { KPH.target = same; callPhotoForm(d); };
  $('kpGo').onclick = () => { const r = {}; KPH_F.forEach(([k]) => (r[k] = $('kp_' + k).value.trim())); if (!r.name && !r.phone) { toast('至少要有姓名或手機'); return; }
    let k = KPH.target; if (k) { k = Vault.get('calls', k.id); KPH_F.forEach(([f]) => { if (f !== 'no' && r[f]) k[f] = r[f]; }); } else k = { ...r, source: '拍照新增', status: '待撥', nextDate: '', logs: [] };
    if (r.addr) k.area = areaOf(r.addr) || k.area; k = Vault.put('calls', k); KPH = null; ov.innerHTML = ''; S.view = 'calls'; S.ktable = false; S.ksel = k.id; S.showDetail = true; App.render(); toast('已存好'); };
}
function closeCrmImport() { CRMI = null; ov.innerHTML = ''; }
async function crmFromImage(f) {
  if (!CRMI) return; CRMI.shot = f;
  const im = $('ciShot'); im.src = URL.createObjectURL(f); im.hidden = false;
  $('ciStatus').textContent = OCR ? '辨識中…' : '第一次使用要先下載辨識元件（約 13MB），請稍等…';
  try { const txt = await ocrImage(f); if (!CRMI) return; $('ciText').value = txt; $('ciStatus').textContent = '辨識完成，請檢查下面的欄位'; showCrmForm(parseCRM(txt)); }
  catch (e) { console.warn(e); $('ciStatus').textContent = '辨識失敗：' + (e.message || e); }
}
function showCrmForm(data) {
  if (!CRMI) return;
  // 同一個身分證字號＝同一人；只有名字一樣時，身分證不能互相矛盾。也可以手動改成建立新客戶
  let tgt = CRMI.target; if (!tgt && !CRMI.forceNew) tgt = clients().find((c) => (data.idNo && c.idNo === data.idNo) || (data.name && c.name === data.name && !(data.idNo && c.idNo && c.idNo !== data.idNo))) || null;
  CRMI.data = data;
  CRMI.resolved = tgt;
  const keys = CRM_FIELDS.filter(([k]) => data[k] || (k === 'name' && !tgt)); // 沒讀到姓名也留一格讓你填
  if (!keys.some(([k]) => data[k])) { $('ciForm').innerHTML = '<div class="note" style="color:var(--warn);background:var(--warn-soft)">沒有讀到資料。截圖請包含「姓名」「ID」「手機」那一區，或改用複製文字。</div>'; return; }
  $('ciForm').innerHTML = `<div class="sec"><h2>確認要帶入的資料 <span class="sub">${tgt ? `會更新現有客戶「${esc(tgt.name)}」${CRMI.target ? '' : ' <button class="linkbtn" id="ciNew">不是同一人，改成建立新客戶</button>'}` : '會建立新客戶'}</span></h2>
   ${keys.map(([k, l]) => { const old = tgt?.[k]; return `<div class="cirow"><label class="cichk"><input type="checkbox" data-cik="${k}" ${!old || old !== data[k] ? 'checked' : ''}> ${l}</label><input class="ciin" id="ci_${k}" value="${esc(data[k])}">${old && old !== data[k] ? `<span class="sub">目前：${esc(k === 'idNo' ? old.slice(0, 3) + '…' : old)}</span>` : old ? '<span class="sub">相同</span>' : ''}</div>`; }).join('')}
   ${CRMI.shot ? '<label class="cichk" style="margin-top:6px"><input type="checkbox" id="ciKeep"> 把這張截圖也存進「證件與文件」</label>' : ''}
   <div class="note">地址、姓名請特別看一下，辨識偶爾會錯一兩個字。</div>
   <div class="actions"><button class="btn pri" id="ciApply">${tgt ? '更新資料' : '建立客戶'}</button></div></div>`;
  if ($('ciNew')) $('ciNew').onclick = () => { CRMI.forceNew = true; showCrmForm(CRMI.data); };
  $('ciApply').onclick = async () => {
    const isNew = !CRMI.resolved, rec = CRMI.resolved || { name: '', docs: [] };
    ov.querySelectorAll('[data-cik]').forEach((cb) => { if (cb.checked) rec[cb.dataset.cik] = $('ci_' + cb.dataset.cik).value.trim(); });
    if (isNew && !rec.name) { toast('沒有姓名，請在「姓名」欄填上並勾選'); return; }
    const saved = Vault.put('clients', rec); const keep = $('ciKeep')?.checked, shot = CRMI.shot;
    closeCrmImport(); S.view = 'clients'; S.csel = saved.id; S.q = ''; S.ccat = '全部'; S.showDetail = true; S.showIds = false; App.render(); // 清掉搜尋，新客戶才不會被篩掉
    document.querySelector(`[data-csel="${saved.id}"]`)?.scrollIntoView({ block: 'nearest' });
    if (keep && shot) addImagesDirect(saved, 'clients', [shot], '');
    toast(`已${isNew ? '建立' : '更新'}「${saved.name || '客戶'}」`);
  };
}
/* ================= 告知資料（財務、體況） ================= */
const ASSETS = [['job', '職業／職稱', ''], ['company', '公司', ''], ['deposit', '存款', '萬'], ['securities', '股票・基金', '萬'], ['realEstate', '不動產', '萬'], ['policyValue', '保單價值', '萬'], ['otherAssets', '其他資產', '萬'], ['mortgage', '房貸', '萬'], ['otherDebt', '信貸・其他負債', '萬']];
/* 財產明細：照國泰「財務狀況告知書」的順序。舊資料（職業、公司、存款、股票、不動產、房貸、信貸）自動搬到新格子 */
const A_PURPOSE = ['保障', '子女教育經費', '退休規劃', '房屋貸款', '其他'];
const A_PREMSRC = ['薪資', '退休金', '閒置資金', '投資收益', '租金收入', '貸款（含保單借款）', '保險單解約金、部分提領、縮小保額', '其他'];
const A_REL = ['配偶', '父母', '子女', '祖孫', '其他'];
const A_RETYPE = ['無', '土地', '大樓', '公寓', '透天', '其他'];
const A_READDR = ['同要保人住所', '同被保人住所', '其他'];
const A_LOANS = ['房屋貸款', '保險單借款', '其他貸款'];
function assetsMigrate(a0) {
  const a = { ...(a0 || {}) };
  if (a.job && !a.title) a.title = a.job; delete a.job;
  if (a.securities && !a.stock) { a.stock = '是'; a.stockAmt = a.securities; } delete a.securities;
  if (a.realEstate && !a.re1value) a.re1value = a.realEstate; delete a.realEstate;
  const loans = new Set(a.loans || []); if (a.mortgage) loans.add('房屋貸款'); if (a.otherDebt) loans.add('其他貸款'); if (loans.size) a.loans = [...loans];
  Object.keys(a).forEach((k) => { if (a[k] === '' || (Array.isArray(a[k]) && !a[k].length)) delete a[k]; });
  return a;
}
function assetsFormHTML(a, inp) {
  const checks = (key, opts) => `<div class="tags achk" data-achk="${key}">${opts.map((o) => `<button type="button" class="tag${(a[key] || []).includes(o) ? ' on' : ''}" data-t="${esc(o)}">${esc(o)}</button>`).join('')}</div>`;
  const radio = (key, opts) => `<div class="tags arad" data-arad="${key}">${opts.map((o) => `<button type="button" class="tag${a[key] === o ? ' on' : ''}" data-t="${esc(o)}">${esc(o)}</button>`).join('')}</div>`;
  const re = (n) => `<div class="are"><div class="f">不動產 ${n}${radio('re' + n + 'type', A_RETYPE)}</div>
    <div class="f">地址${radio('re' + n + 'addrType', A_READDR)}</div>${inp('a_re' + n + 'addr', a['re' + n + 'addr'], '地址（選其他時填）', '')}
    <div class="dgrid">${inp('a_re' + n + 'value', a['re' + n + 'value'], '市值', '萬', 'inputmode="decimal"')}${inp('a_re' + n + 'size', a['re' + n + 'size'], '坪數', '坪', 'inputmode="decimal"')}${inp('a_re' + n + 'age', a['re' + n + 'age'], '屋齡', '年', 'inputmode="decimal"')}</div></div>`;
  return `<div class="f">投保目的（可複選）${checks('purpose', A_PURPOSE)}</div>${inp('a_purposeOther', a.purposeOther, '投保目的：其他', '')}
   <div class="f">保費來源（可複選）<span class="sub" style="color:var(--warn)">生調表要跟這裡一致才能上傳</span>${checks('premSrc', A_PREMSRC)}</div>${inp('a_premSrcOther', a.premSrcOther, '保費來源：其他', '')}
   <h3 class="ah">身份角色別</h3>
   <div class="f">身份${radio('role', ['要保人', '主要經濟來源者'])}<span class="sub">要保人未成年或沒有經濟能力時，填主要經濟來源者的工作和財務</span></div>
   <div class="f">主要經濟者與被保人關係${radio('earnerRel', A_REL)}</div>
   <div class="dgrid">${inp('a_earnerRelOther', a.earnerRelOther, '關係：其他', '')}${inp('a_earnerName', a.earnerName, '主要經濟者姓名', '')}${inp('a_earnerId', a.earnerId, '主要經濟者身分證字號', '')}</div>
   <h3 class="ah">工作狀況</h3>
   <div class="dgrid">${inp('a_company', a.company, '公司名稱', '')}${inp('a_business', a.business, '營業項目', '')}${inp('a_title', a.title, '職稱', '')}${inp('a_years', a.years, '年資', '年', 'inputmode="decimal"')}</div>
   ${inp('a_jobDesc', a.jobDesc, '工作內容', '')}<div class="f">公司股東${radio('shareholder', ['否', '是'])}</div>
   <h3 class="ah">財務狀況</h3>
   <div class="sub">工作年收入、其他年收入、家庭收入用上面「基本告知」的數字，不用再填一次。家庭收入選要保人時，生調表要一致。</div>
   <div class="dgrid">${inp('a_bank1', a.bank1, '往來銀行一', '')}${inp('a_bank2', a.bank2, '往來銀行二', '')}${inp('a_deposit', a.deposit, '存款總額', '萬', 'inputmode="decimal"')}</div>
   <div class="f">股票、基金、債券${radio('stock', ['否', '是'])}</div>${inp('a_stockAmt', a.stockAmt, '股票、基金、債券金額（自己參考，可不填）', '萬', 'inputmode="decimal"')}
   ${re(1)}${re(2)}
   <div class="f">借款（可複選）${checks('loans', A_LOANS)}</div>
   <div class="dgrid">${inp('a_mortgage', a.mortgage, '房貸金額（自己參考）', '萬', 'inputmode="decimal"')}${inp('a_otherDebt', a.otherDebt, '其他借款金額（自己參考）', '萬', 'inputmode="decimal"')}</div>
   <h3 class="ah">自己參考（公司表單沒有）</h3>
   <div class="dgrid">${inp('a_policyValue', a.policyValue, '保單價值', '萬', 'inputmode="decimal"')}${inp('a_otherAssets', a.otherAssets, '其他資產', '萬', 'inputmode="decimal"')}</div>
   ${inp('a_explain', a.explain, '資金來源／說明', '')}`;
}
const A_TEXT = ['purposeOther', 'premSrcOther', 'earnerRelOther', 'earnerName', 'earnerId', 'company', 'business', 'title', 'years', 'jobDesc', 'bank1', 'bank2', 'deposit', 'stockAmt', 're1addr', 're1value', 're1size', 're1age', 're2addr', 're2value', 're2size', 're2age', 'mortgage', 'otherDebt', 'policyValue', 'otherAssets', 'explain'];
function assetsRead(g) {
  const a = {}; A_TEXT.forEach((k) => { const v = g('a_' + k); if (v !== '') a[k] = v; });
  ov.querySelectorAll('[data-achk]').forEach((box) => { const v = [...box.querySelectorAll('.tag.on')].map((x) => x.dataset.t); if (v.length) a[box.dataset.achk] = v; });
  ov.querySelectorAll('[data-arad]').forEach((box) => { const v = box.querySelector('.tag.on')?.dataset.t; if (v) a[box.dataset.arad] = v; });
  return a;
}
function assetsSummaryHTML(a0) {
  const a = assetsMigrate(a0); if (!Object.keys(a).length) return '';
  const kv = (l, v, u = '') => (v ? `<div class="kv"><span>${l}</span><b>${esc(Array.isArray(v) ? v.join('、') : v)}${u ? ' ' + u : ''}</b></div>` : '');
  const re = (n) => a['re' + n + 'type'] && a['re' + n + 'type'] !== '無' ? kv('不動產 ' + n, [a['re' + n + 'type'], a['re' + n + 'addr'] || (a['re' + n + 'addrType'] !== '其他' ? a['re' + n + 'addrType'] : ''), a['re' + n + 'value'] ? '市值 ' + a['re' + n + 'value'] + ' 萬' : '', a['re' + n + 'size'] ? a['re' + n + 'size'] + ' 坪' : '', a['re' + n + 'age'] ? '屋齡 ' + a['re' + n + 'age'] + ' 年' : ''].filter(Boolean).join('・')) : kv('不動產 ' + n, a['re' + n + 'value'] ? '市值 ' + a['re' + n + 'value'] + ' 萬' : '');
  return `<div class="sub">財產明細</div><div class="kvs">${kv('投保目的', [...(a.purpose || []).filter((x) => x !== '其他'), a.purposeOther].filter(Boolean))}${kv('保費來源', [...(a.premSrc || []).filter((x) => x !== '其他'), a.premSrcOther].filter(Boolean))}
   ${kv('身份', a.role)}${kv('主要經濟者', [a.earnerName, a.earnerRel === '其他' ? a.earnerRelOther : a.earnerRel].filter(Boolean).join('・'))}
   ${kv('公司', a.company)}${kv('營業項目', a.business)}${kv('職稱', a.title)}${kv('年資', a.years, '年')}${kv('工作內容', a.jobDesc)}${kv('公司股東', a.shareholder)}
   ${kv('往來銀行', [a.bank1, a.bank2].filter(Boolean).join('、'))}${kv('存款總額', a.deposit, '萬')}${kv('股票基金債券', a.stock ? a.stock + (a.stockAmt ? '（' + a.stockAmt + ' 萬）' : '') : a.stockAmt ? a.stockAmt + ' 萬' : '')}
   ${re(1)}${re(2)}${kv('借款', a.loans)}${kv('房貸', a.mortgage, '萬')}${kv('其他借款', a.otherDebt, '萬')}${kv('保單價值', a.policyValue, '萬')}${kv('其他資產', a.otherAssets, '萬')}</div>${a.explain ? `<div class="sub">說明：${esc(a.explain)}</div>` : ''}`;
}
const FIN = [['workIncome', '工作年收入', '萬'], ['otherIncome', '其他年收入', '萬'], ['personalAssets', '個人資產', '萬'], ['familyIncome', '家庭年收入', '萬'], ['familyAssets', '家庭資產', '萬']];
function parseFin(raw) {
  const t = String(raw || '').replace(/[\uFF01-\uFF5E]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));
  const lbl = (s) => s.split('').join('\\s*'); const out = {};
  FIN.forEach(([k, l]) => { const m = t.match(new RegExp(lbl(l) + '[^0-9\\n]{0,8}([0-9][0-9,.]*)')); if (m) out[k] = m[1].replace(/,/g, ''); });
  const rel = t.match(/關\s*係\s*為[\s:：]*\d?\s*[:：]?\s*([\u4e00-\u9fff]{1,4})/); if (rel) { out.mainEarner = '否'; out.relation = rel[1]; }
  return out;
}
const bmi = (h, w) => { h = +h; w = +w; return h > 50 && w > 10 ? (w / (h / 100) ** 2).toFixed(1) : ''; };
function discSummary(r) {
  const f = r.fin || {}, h = r.health || {};
  const shown = discSlotsShown(r), have = shown.filter(([k]) => (r.files || []).some((x) => x.slot === k)), miss = shown.filter((x) => !have.includes(x));
  const shotLine = `<div class="sub" style="${miss.length ? 'color:#A32D2D' : ''}">截圖 ${have.length}/${shown.length}${miss.length ? '・還沒拍：' + miss.map(([, t]) => t).join('、') : '・齊全'}</div>`;
  const fin = FIN.filter(([k]) => f[k] !== undefined && f[k] !== '').map(([k, l, u]) => `<div class="kv"><span>${l}</span><b>${esc(f[k])} ${u}</b></div>`).join('');
  const b = bmi(h.height, h.weight);
  return shotLine + `${fin || f.mainEarner || f.incomeSrc ? `<div class="sub">財務</div><div class="kvs">${fin}${f.incomeSrc ? `<div class="kv"><span>收入來源</span><b>${esc(f.incomeSrc)}</b></div>` : ''}${f.mainEarner ? `<div class="kv"><span>主要經濟來源</span><b>${esc(f.mainEarner)}${f.mainEarner === '否' && f.relation ? '（主要經濟者是' + esc(f.relation) + '）' : ''}</b></div>` : ''}</div>${f.reason ? `<div class="sub">收入不一致原因：${esc(f.reason)}</div>` : ''}` : ''}
   ${h.height || h.weight || h.smoke || h.drink || h.items ? `<div class="sub">體況</div><div class="kvs">${h.height ? `<div class="kv"><span>身高</span><b>${esc(h.height)} cm</b></div>` : ''}${h.weight ? `<div class="kv"><span>體重</span><b>${esc(h.weight)} kg${b ? '（BMI ' + b + '）' : ''}</b></div>` : ''}${h.smoke ? `<div class="kv"><span>吸菸</span><b>${esc(h.smoke)}</b></div>` : ''}${h.drink ? `<div class="kv"><span>飲酒</span><b>${esc(h.drink)}</b></div>` : ''}</div>${h.items ? `<pre class="discpre">${esc(h.items)}</pre>` : ''}` : ''}
   ${assetsSummaryHTML(r.assets)}
   ${r.note ? `<pre class="discpre">${esc(r.note)}</pre>` : ''}
   ${(r.files || []).length ? `<div class="nimgs">${r.files.map((f, i) => f.type === 'pdf' ? `<button class="nimg pdfmini" data-discfile="${r.id}:${i}">PDF</button>` : `<button class="nimg" data-discfile="${r.id}:${i}"><img data-fid="${f.fileId}" alt=""></button>`).join('')}</div>` : ''}`;
}
function disclosureHTML(x) {
  const list = (x.disclosures || []).slice().sort((a, b) => (b.date || '').localeCompare(a.date || ''));
  const head = `<h2>告知資料 <span style="display:flex;gap:6px">${list.length ? `<button class="btn" data-act="togglehealth" style="padding:4px 10px;font-size:13px">${S.showHealth ? '隱藏' : '顯示'}</button>` : ''}<button class="btn" data-act="adddisc" style="padding:4px 10px;font-size:13px">＋ 新增告知</button></span></h2>`;
  if (!list.length) return `<div class="sec">${head}<div class="sub">送件後把財務告知、體況告知存在這裡，下次不用再問一次。</div></div>`;
  if (!S.showHealth) return `<div class="sec">${head}<div class="veilbox">已存 ${list.length} 筆告知，最新是 ${fmtDay(list[0].date)}。按「顯示」查看。</div></div>`;
  return `<div class="sec">${head}${list.map((r, i) => { const cs = r.caseId && Vault.get('cases', r.caseId); const top = `<span class="discdate">${fmtDay(r.date)}</span>${cs ? `<button class="linkbtn" data-gocase="${cs.id}">${esc(cs.cat)}・${esc(cs.title || '未命名')}</button>` : ''}${i === 0 ? '<span class="pill p-go">最新</span>' : ''}<span style="flex:1"></span><button class="linkbtn" data-editdisc="${r.id}">編輯</button><button class="tldel" data-deldisc="${r.id}">刪除</button>`;
    return i === 0 ? `<div class="disc"><div class="dischead">${top}</div>${discSummary(r)}</div>` : `<details class="disc"><summary class="dischead">${top}</summary>${discSummary(r)}</details>`; }).join('')}</div>`;
}
let DSC = null;
function openDisclosure(cl, id, caseId) {
  const old = id ? (cl.disclosures || []).find((x) => x.id === id) : null;
  let base = old ? JSON.parse(JSON.stringify(old)) : null;
  if (!base) { const prev = (cl.disclosures || []).slice().sort((a, b) => (b.date || '').localeCompare(a.date || ''))[0]; base = { id: uid(), date: localDay(), caseId: caseId || '', fin: { ...(prev?.fin || {}) }, health: { ...(prev?.health || {}) }, assets: { ...(prev?.assets || {}) }, note: '', files: [] }; }
  DSC = { cl, r: base, isNew: !old };
  const f = base.fin, h = base.health, a = base.assets || {};
  const hasAssets = Object.keys(assetsMigrate(a)).length > 0, hasMore = h.smoke || h.drink || h.items || f.reason || base.note;
  const inp = (k, v, l, u, extra = '') => `<label class="f">${l}<span class="unit"><input id="d_${k}" value="${esc(v ?? '')}" ${extra}>${u ? `<i>${u}</i>` : ''}</span></label>`;
  const cases2 = clientCases(cl);
  ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>${old ? '編輯' : '新增'}告知資料 · ${esc(cl.name || '')}</h2><button class="btn" id="dX">取消</button><button class="btn pri" id="dSave">儲存</button></div>
  <div class="ovbody" style="flex-direction:column;max-width:780px;margin:0 auto;width:100%">
   ${!old && (cl.disclosures || []).length ? '<div class="note">已先帶入上一次的內容，只要改有變動的地方。</div>' : ''}
   <div class="grid2"><label class="f">日期<input type="date" id="d_date" value="${esc(base.date)}"></label><label class="f">來自案件<select id="d_case"><option value="">（不指定）</option>${cases2.map((c) => `<option value="${c.id}"${c.id === base.caseId ? ' selected' : ''}>${esc(fmtDay(caseDate(c)) + ' ' + c.cat + '・' + (c.title || '未命名'))}</option>`).join('')}</select></label></div>
   <div class="sec"><h2>基本告知 <span class="sub">熟人填這些就夠了 · 金額單位：萬</span></h2>
    <div class="actions"><label class="btn">截圖帶入<input type="file" id="dShot" accept="image/*"></label><button class="btn" id="dPasteShot">📋 貼上截圖</button><span class="sub" id="dOcr">也可以在這個畫面按 Ctrl+V</span></div>
    <div class="dgrid">${inp('height', h.height, '身高', 'cm', 'inputmode="decimal"')}${inp('weight', h.weight, '體重', 'kg', 'inputmode="decimal"')}
     ${FIN.map(([k, l, u]) => inp(k, f[k], l, u, 'inputmode="decimal"')).join('')}
     <label class="f">主要經濟來源<select id="d_mainEarner">${['', '是', '否'].map((o) => `<option${o === (f.mainEarner || '') ? ' selected' : ''}>${o}</option>`).join('')}</select></label>
     ${inp('relation', f.relation, '主要經濟者是', '', 'placeholder="例：配偶"')}
     <label class="f">收入來源<select id="d_incomeSrc">${['', '無', '薪津', '營業收入', '投資', '其他'].map((o) => `<option${o === (f.incomeSrc || '') ? ' selected' : ''}>${o}</option>`).join('')}</select></label></div>
    ${inp('reason', f.reason, '收入不一致原因', '')}</div>
   <div class="sec"><h2>截圖 <span class="sub">照公司畫面一格一格放，紅色的就是還沒拍</span></h2>
    <div class="tags"><button type="button" class="tag${DSC.r.needAssets !== false ? ' on' : ''}" id="dNeedA">這次有填財務狀況告知書（財產版）</button></div>
    <div class="dslots" id="dSlots"></div><input type="file" id="dSlotIn" accept="image/*,application/pdf,.pdf" hidden></div>
   <details class="sec fold"${hasAssets ? ' open' : ''}><summary><h2>財產明細手打 <span class="sub">可以不填，有截圖就夠了</span></h2></summary>
    ${assetsFormHTML(assetsMigrate(a), inp)}</details>
   <details class="sec fold"${hasMore ? ' open' : ''}><summary><h2>體況詳細・其他 <span class="sub">不熟的客戶再填</span></h2></summary>
    <div class="dgrid"><label class="f">吸菸<select id="d_smoke">${['', '無', '有', '已戒'].map((o) => `<option${o === (h.smoke || '') ? ' selected' : ''}>${o}</option>`).join('')}</select></label>
     <label class="f">飲酒<select id="d_drink">${['', '無', '偶爾', '經常'].map((o) => `<option${o === (h.drink || '') ? ' selected' : ''}>${o}</option>`).join('')}</select></label></div>
    <label class="f">告知事項（病史、住院、手術、用藥、體檢異常…）<textarea id="d_items" class="grow" rows="3">${esc(h.items)}</textarea></label>
    <label class="f">其他備註<textarea id="d_note" class="grow" rows="2">${esc(base.note)}</textarea></label></details>
   <div class="sec"><h2>其他附件 <span class="sub">要保書、告知書等等</span></h2>
    <div class="nimgs edit" id="dFiles"></div>
    <div class="actions"><label class="btn">＋ 加照片或 PDF<input type="file" id="dAdd" accept="image/*,application/pdf,.pdf" multiple></label><button class="btn" id="dPaste">📋 貼上</button></div></div>
  </div></div>`;
  drawDiscFiles(); fitNotes();
  ov.querySelectorAll('[data-achk]').forEach((box) => (box.onclick = (e) => { const t = e.target.closest('[data-t]'); if (t) t.classList.toggle('on'); }));
  ov.querySelectorAll('[data-arad]').forEach((box) => (box.onclick = (e) => { const t = e.target.closest('[data-t]'); if (!t) return; const on = t.classList.contains('on'); box.querySelectorAll('.tag').forEach((x) => x.classList.remove('on')); if (!on) t.classList.add('on'); }));
  $('dX').onclick = () => { if (DSC.isNew) (DSC.r.files || []).forEach((x) => Files.remove(x.fileId)); DSC = null; ov.innerHTML = ''; };
  $('dSave').onclick = () => {
    const r = DSC.r, g = (k) => $('d_' + k)?.value.trim() ?? '';
    r.date = g('date') || localDay(); r.caseId = $('d_case').value; r.note = g('note');
    r.fin = Object.fromEntries([...FIN.map(([k]) => [k, g(k)]), ['mainEarner', g('mainEarner')], ['relation', g('relation')], ['incomeSrc', g('incomeSrc')], ['reason', g('reason')]].filter(([, v]) => v !== ''));
    r.health = Object.fromEntries([['height', g('height')], ['weight', g('weight')], ['smoke', g('smoke')], ['drink', g('drink')], ['items', $('d_items').value.trim()]].filter(([, v]) => v !== ''));
    r.assets = assetsRead(g);
    const c2 = DSC.cl; c2.disclosures = [...(c2.disclosures || []).filter((x) => x.id !== r.id), r]; Vault.put('clients', c2);
    DSC = null; ov.innerHTML = ''; S.view = 'clients'; S.csel = c2.id; S.showDetail = true; S.showHealth = true; App.render(); toast('告知資料已儲存');
  };
  $('dAdd').onchange = (e) => { const fs = [...e.target.files]; e.target.value = ''; discAddFiles(fs, false, null); };
  $('dPaste').onclick = async () => { const fs = await clipFiles(); if (fs.length) discAddFiles(fs, false, null); };
  $('dShot').onchange = (e) => { const fl = e.target.files[0]; e.target.value = ''; if (fl) discAddFiles([fl], true, 'gen'); };
  $('dPasteShot').onclick = async () => { const fs = await clipFiles(); if (fs.length) discAddFiles(fs, true, 'gen'); };
  $('dNeedA').onclick = (e) => { DSC.r.needAssets = DSC.r.needAssets === false; e.currentTarget.classList.toggle('on', DSC.r.needAssets !== false); drawDiscFiles(); };
  $('dSlotIn').onchange = (e) => { const fs = [...e.target.files]; e.target.value = ''; if (fs.length && DSC.slotPick) discAddFiles(fs, ['gen', 'p3'].includes(DSC.slotPick), DSC.slotPick); };
}
async function clipFiles() {
  try { const out = []; for (const it of await navigator.clipboard.read()) { const t = it.types.find((y) => y.startsWith('image/')); if (t) out.push(new File([await it.getType(t)], 'shot.png', { type: t })); } if (!out.length) toast('剪貼簿裡沒有圖片'); return out; }
  catch (e) { toast('這裡不能讀剪貼簿，請直接按 Ctrl+V'); return []; }
}
function drawDiscFiles() {
  const box = $('dFiles'); if (!box || !DSC) return;
  const sbox = $('dSlots');
  if (sbox) {
    const slot = (k, t, d) => { const i = (DSC.r.files || []).findIndex((f) => f.slot === k), f = DSC.r.files?.[i];
      return `<div class="dslot${f ? ' ok' : ''}"><div class="dslh"><b>${t}</b><span>${d}</span></div>
       ${f ? `<div class="nimgwrap">${f.type === 'pdf' ? `<button class="nimg pdfmini" data-dv="${i}">PDF</button>` : `<button class="nimg" data-dv="${i}"><img data-fid="${f.fileId}" alt=""></button>`}<button class="nimgdel" data-dd="${i}" aria-label="移除">✕</button></div>` : '<div class="dsmiss">還沒拍</div>'}
       <div class="dsbtn"><button type="button" class="btn sm" data-dsp="${k}">選照片</button><button type="button" class="btn sm" data-dsc="${k}">📋 貼上</button></div></div>`; };
    const ps = DISC_SLOTS.filter(([k]) => k !== 'gen'), pDone = ps.filter(([k]) => (DSC.r.files || []).some((f) => f.slot === k)).length;
    const wasOpen = sbox.querySelector('details')?.open;
    sbox.innerHTML = slot(...DISC_SLOTS[0]) + (DSC.r.needAssets !== false ? `<details class="dsgroup"${(wasOpen ?? pDone < ps.length) ? ' open' : ''}><summary><b>財產版</b> <span class="${pDone < ps.length ? 'dsbad' : 'dsgood'}">${pDone}/${ps.length}${pDone < ps.length ? ' 還沒拍齊' : ' 齊全'}</span></summary><div class="dscol">${ps.map((x) => slot(...x)).join('')}</div></details>` : '');
    sbox.querySelectorAll('[data-dsp]').forEach((b) => (b.onclick = () => { DSC.slotPick = b.dataset.dsp; $('dSlotIn').click(); }));
    sbox.querySelectorAll('[data-dsc]').forEach((b) => (b.onclick = async () => { const fs = await clipFiles(); if (fs.length) discAddFiles(fs.slice(0, 1), ['gen', 'p3'].includes(b.dataset.dsc), b.dataset.dsc); }));
  }
  box.innerHTML = (DSC.r.files || []).map((f, i) => (f.slot ? '' : `<div class="nimgwrap">${f.type === 'pdf' ? `<button class="nimg pdfmini" data-dv="${i}">PDF</button>` : `<button class="nimg" data-dv="${i}"><img data-fid="${f.fileId}" alt=""></button>`}<button class="nimgdel" data-dd="${i}" aria-label="移除">✕</button></div>`)).join('') || '<span class="sub">還沒有附件</span>';
  const both = [box, sbox].filter(Boolean), all = (sel) => both.flatMap((bx) => [...bx.querySelectorAll(sel)]);
  all('img[data-fid]').forEach(async (im) => { const u = await Files.url(im.dataset.fid).catch(() => null); if (u && im.isConnected) im.src = u; });
  all('[data-dv]').forEach((b) => (b.onclick = () => { const f = DSC.r.files[+b.dataset.dv]; Files.url(f.fileId).then((u) => { if (!u) return; const v = document.createElement('div'); v.innerHTML = `<div class="ov" style="z-index:30"><div class="ovhead"><h2>${esc(f.name)}</h2><button class="btn pri">關閉</button></div>${f.type === 'pdf' ? `<iframe src="${u}" style="flex:1;width:100%;border:0;background:#fff"></iframe>` : `<div class="ovbody" style="justify-content:center"><img src="${u}" alt="" style="max-width:100%;height:auto;align-self:flex-start"></div>`}</div>`; v.querySelector('button').onclick = () => v.remove(); document.body.appendChild(v); }); }));
  all('[data-dd]').forEach((b) => (b.onclick = () => { const [f] = DSC.r.files.splice(+b.dataset.dd, 1); if (f) Files.remove(f.fileId); drawDiscFiles(); }));
}
const DISC_SLOTS = [['gen', '一般版', '生調表的收入狀況'], ['p1', '財產版 1', '投保目的、保費來源、身份角色別、主要經濟者'], ['p2', '財產版 2', '工作狀況：公司名稱、營業項目、職稱、年資、工作內容、公司股東'], ['p3', '財產版 3', '收入、往來銀行、存款、股票'], ['p4', '財產版 4', '不動產、借款']];
const discSlotsShown = (r) => DISC_SLOTS.filter(([k]) => k === 'gen' || r.needAssets !== false);
async function discAddFiles(files, readFin, slot) {
  if (!DSC) return; const r = DSC.r; r.files = r.files || [];
  const nextEmpty = () => discSlotsShown(r).map(([k]) => k).find((k) => !r.files.some((f) => f.slot === k)) || null;
  for (const f of files) {
    const sl = slot === undefined ? nextEmpty() : slot; // 直接貼上（Ctrl+V）或拖進來：放進第一個還沒拍的格子
    const pdf = f.type === 'application/pdf' || /\.pdf$/i.test(f.name || '');
    if (!pdf && !/^image\//.test(f.type)) continue;
    const blob = pdf ? new Blob([await f.arrayBuffer()], { type: 'application/pdf' }) : await shrink(f); if (!blob) continue;
    const fileId = await Files.add(blob, f.name || '告知');
    if (sl) r.files = r.files.filter((x) => { if (x.slot === sl) { Files.remove(x.fileId); return false; } return true; }); // 同一格放新的就換掉舊的
    r.files.push({ fileId, name: sl ? DISC_SLOTS.find(([k]) => k === sl)[1] : pdf ? (f.name || '').replace(/\.pdf$/i, '') || 'PDF' : '告知截圖', type: pdf ? 'pdf' : 'img', date: localDay(), ...(sl ? { slot: sl } : {}) });
    readFin = readFin && ['gen', 'p3'].includes(sl);
    if (readFin && !pdf) {
      $('dOcr').textContent = OCR ? '辨識中…' : '第一次使用要先下載辨識元件，請稍等…';
      try { const fin = parseFin(await ocrImage(f)); const ks = Object.keys(fin); ks.forEach((k) => { const el = $('d_' + k); if (el) el.value = fin[k]; }); if ($('dOcr')) $('dOcr').textContent = ks.length ? `已帶入 ${ks.length} 個欄位，請核對數字` : '沒有讀到數字，請手動填'; }
      catch (e) { if ($('dOcr')) $('dOcr').textContent = '辨識失敗，請手動填'; }
    }
  }
  drawDiscFiles();
}
/* ================= 關係圖 ================= */
const rels = () => Vault.all('rels').filter((r) => Vault.get('clients', r.a) && Vault.get('clients', r.b));
const FAMILY = ['parent', 'spouse', 'sibling', 'relative'];
const REL_PICK = [ // 從「這位客戶」的角度選：對方是他的…
  ['父母', 'parent', 'in'], ['子女', 'parent', 'out'], ['配偶', 'spouse'], ['兄弟姊妹', 'sibling'], ['其他親戚', 'relative'],
  ['朋友', 'friend'], ['同事', 'colleague'], ['介紹人', 'referral', 'in'], ['他介紹的人', 'referral', 'out'],
];
function relLabel(r, me) { // 對方對「me」來說是什麼
  const other = Vault.get('clients', r.a === me ? r.b : r.a), g = other?.gender, out = r.a === me;
  switch (r.type) {
    case 'parent': return out ? (g === '男' ? '兒子' : g === '女' ? '女兒' : '子女') : (g === '男' ? '爸爸' : g === '女' ? '媽媽' : '父母');
    case 'spouse': return g === '男' ? '先生' : g === '女' ? '太太' : '配偶';
    case 'sibling': return g === '男' ? '兄弟' : g === '女' ? '姊妹' : '兄弟姊妹';
    case 'relative': return r.note || '親戚';
    case 'friend': return '朋友';
    case 'colleague': return '同事';
    case 'referral': return out ? '他介紹的人' : '介紹人';
  }
  return '';
}
function relsHTML(x) {
  const mine = rels().filter((r) => r.a === x.id || r.b === x.id);
  const fam = mine.filter((r) => FAMILY.includes(r.type)), soc = mine.filter((r) => !FAMILY.includes(r.type));
  const row = (r) => { const o = Vault.get('clients', r.a === x.id ? r.b : r.a); return `<div class="relrow"><span class="reltag ${FAMILY.includes(r.type) ? 'fam' : 'soc'}">${esc(relLabel(r, x.id))}</span><button class="linkbtn relname" data-goclient="${o.id}">${esc(o.name || '未填姓名')}</button><button class="tldel" data-delrel="${r.id}" aria-label="移除關係">✕</button></div>`; };
  return `<div class="sec"><h2>關係 <span style="display:flex;gap:6px"><button class="btn" data-act="relgraph" style="padding:4px 10px;font-size:13px">全部關係圖</button><button class="btn" data-act="addrel" style="padding:4px 10px;font-size:13px">＋ 新增關係</button></span></h2>
   ${mine.length ? `<div class="egograph" id="egoGraph"></div><div class="grid2">${fam.length ? `<div><div class="sub" style="margin-bottom:4px">家人</div>${fam.map(row).join('')}</div>` : ''}${soc.length ? `<div><div class="sub" style="margin-bottom:4px">朋友・其他</div>${soc.map(row).join('')}</div>` : ''}</div>`
    : '<div class="sub">還沒有關係。可以加家人（不受國泰戶長的限制）、朋友、同事、介紹人。</div>'}</div>`;
}
function openRelPicker(x) {
  let pick = null;
  const draw = () => {
    const q = ($('rpQ')?.value || '').trim();
    const linked = new Set(rels().filter((r) => r.a === x.id || r.b === x.id).map((r) => (r.a === x.id ? r.b : r.a)));
    const pool = clients().filter((c) => c.id !== x.id && (!q || (c.name || '').includes(q))).sort((a, b) => linked.has(a.id) - linked.has(b.id) || (a.name || '').localeCompare(b.name || '', 'zh-Hant')).slice(0, 40);
    $('rpList').innerHTML = (q && !clients().some((c) => c.name === q) ? `<button class="item" data-newperson="1"><span class="t">＋ 新增「${esc(q)}」</span><span class="s">會建立一位只有名字的客戶，之後可以補資料</span></button>` : '') + pool.map((c) => `<button class="item" data-person="${c.id}"><span class="t"><span>${esc(c.name || '未填姓名')}</span>${linked.has(c.id) ? '<span class="pill p-off">已有關係</span>' : ''}</span></button>`).join('');
  };
  ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>新增「${esc(x.name || '這位客戶')}」的關係</h2><button class="btn" id="rpClose">取消</button></div>
  <div class="ovbody" style="flex-direction:column;max-width:640px;margin:0 auto;width:100%">
   <div class="sub">對方是${esc(x.name || '他')}的…</div>
   <div class="tags" id="rpTypes">${REL_PICK.map(([l], i) => `<button class="tag" data-ti="${i}">${l}</button>`).join('')}</div>
   <label class="f" id="rpNoteWrap" hidden>怎麼稱呼（可不填）<input id="rpNote" placeholder="例如：阿姨、表哥、外婆"></label>
   <input class="search" id="rpQ" placeholder="搜尋或輸入名字" autocomplete="off">
   <div class="list" id="rpList" style="padding:0"></div></div></div>`;
  $('rpClose').onclick = () => (ov.innerHTML = '');
  $('rpTypes').onclick = (e) => { const b = e.target.closest('[data-ti]'); if (!b) return; pick = REL_PICK[+b.dataset.ti]; $('rpTypes').querySelectorAll('.tag').forEach((t) => t.classList.toggle('on', t === b)); $('rpNoteWrap').hidden = pick[1] !== 'relative'; $('rpQ').focus(); };
  $('rpQ').oninput = draw;
  $('rpList').onclick = (e) => {
    const b = e.target.closest('button'); if (!b) return;
    if (!pick) { toast('先選關係，例如「媽媽」選「父母」'); return; }
    let other = b.dataset.person ? Vault.get('clients', b.dataset.person) : null;
    if (b.dataset.newperson) other = Vault.put('clients', { name: $('rpQ').value.trim(), docs: [] });
    if (!other) return;
    const [, type, dir] = pick; const [a, bb] = dir === 'in' ? [other.id, x.id] : [x.id, other.id];
    Vault.put('rels', { a, b: bb, type, note: type === 'relative' ? $('rpNote').value.trim() : '' });
    ov.innerHTML = ''; App.render(); toast(`已加上：${other.name} 是 ${x.name || '他'} 的${pick[0]}`);
  };
  draw();
}
function layoutGraph(nodes, edges, W, H) { // 簡單的力導向排版
  const P = new Map(nodes.map((n, i) => [n.id, { x: W / 2 + Math.cos(i * 2.4) * (40 + i * 6), y: H / 2 + Math.sin(i * 2.4) * (40 + i * 6), vx: 0, vy: 0, fixed: n.fixed }]));
  for (let it = 0; it < 320; it++) {
    const arr = [...P.values()], ids = [...P.keys()];
    for (let i = 0; i < arr.length; i++) for (let j = i + 1; j < arr.length; j++) {
      const p = arr[i], q = arr[j]; let dx = p.x - q.x, dy = p.y - q.y, d2 = dx * dx + dy * dy || 0.01; const f = 7000 / d2, d = Math.sqrt(d2);
      p.vx += dx / d * f; p.vy += dy / d * f; q.vx -= dx / d * f; q.vy -= dy / d * f;
    }
    edges.forEach((e) => { const p = P.get(e.a), q = P.get(e.b); if (!p || !q) return; const L = FAMILY.includes(e.type) ? 115 : 175; const dx = q.x - p.x, dy = q.y - p.y, d = Math.hypot(dx, dy) || 1, f = (d - L) * 0.04; p.vx += dx / d * f; p.vy += dy / d * f; q.vx -= dx / d * f; q.vy -= dy / d * f; });
    arr.forEach((p) => { p.vx += (W / 2 - p.x) * 0.004; p.vy += (H / 2 - p.y) * 0.004; if (p.fixed) { p.x = W / 2; p.y = H / 2; p.vx = p.vy = 0; return; } p.x += Math.max(-20, Math.min(20, p.vx)) ; p.y += Math.max(-20, Math.min(20, p.vy)); p.vx *= 0.55; p.vy *= 0.55; });
    void ids;
  }
  return P;
}
function graphSVG(nodes, edges, H, focusId) {
  const W = 760, P = layoutGraph(nodes, edges, W, H);
  const xs = [...P.values()].map((p) => p.x), ys = [...P.values()].map((p) => p.y), pad = 50;
  let minX = Math.min(...xs) - pad, minY = Math.min(...ys) - pad, vw = Math.max(...xs) - minX + pad, vh = Math.max(...ys) - minY + pad;
  if (vw < 520) { minX -= (520 - vw) / 2; vw = 520; } if (vh < H * 0.6) { minY -= (H * 0.6 - vh) / 2; vh = H * 0.6; } // 人少時不要放太大
  const name = (id) => Vault.get('clients', id)?.name || '未填姓名';
  const merged = new Map(); // 同兩個人之間的多種關係合成一條線
  edges.forEach((e) => { const k = [e.a, e.b].sort().join('|'); if (!merged.has(k)) merged.set(k, { ...e, extra: [] }); else merged.get(k).extra.push(e); });
  edges = [...merged.values()];
  const lines = edges.map((e) => { const p = P.get(e.a), q = P.get(e.b); if (!p || !q) return ''; const fam = FAMILY.includes(e.type); const lbl = e.type === 'parent' ? '親子' : e.type === 'spouse' ? '夫妻' : e.type === 'sibling' ? '手足' : e.type === 'relative' ? (e.note || '親戚') : e.type === 'friend' ? '朋友' : e.type === 'colleague' ? '同事' : '介紹';
    const lab = [lbl, ...e.extra.map((x) => ({ parent: '親子', spouse: '夫妻', sibling: '手足', relative: x.note || '親戚', friend: '朋友', colleague: '同事', referral: '介紹' }[x.type]))].join('・');
    return `<line x1="${p.x}" y1="${p.y}" x2="${q.x}" y2="${q.y}" class="ge ${fam ? 'fam' : e.type}"/><text x="${(p.x + q.x) / 2}" y="${(p.y + q.y) / 2 - 4}" class="gel">${esc(lab)}</text>`; }).join('');
  const dots = nodes.map((n) => { const p = P.get(n.id); const nm = name(n.id); return `<g class="gn${n.id === focusId ? ' focus' : ''}" data-gonode="${n.id}" tabindex="0" role="button" aria-label="${esc(nm)}"><circle cx="${p.x}" cy="${p.y}" r="${n.id === focusId ? 24 : 19}"/><text x="${p.x}" y="${p.y + 4}" class="gnt">${esc(nm.slice(-2))}</text><text x="${p.x}" y="${p.y + (n.id === focusId ? 40 : 35)}" class="gnl">${esc(nm)}</text></g>`; }).join('');
  return `<svg viewBox="${minX} ${minY} ${vw} ${vh}" class="graph" style="max-height:${H}px">${lines}${dots}</svg>`;
}
function drawEgo(x) {
  const box = $('egoGraph'); if (!box) return;
  const all = rels(), ids = new Set([x.id]);
  all.forEach((r) => { if (r.a === x.id) ids.add(r.b); if (r.b === x.id) ids.add(r.a); });
  const first = new Set(ids); all.forEach((r) => { if (first.has(r.a) && FAMILY.includes(r.type)) ids.add(r.b); if (first.has(r.b) && FAMILY.includes(r.type)) ids.add(r.a); }); // 家人的家人也帶進來
  const edges = all.filter((r) => ids.has(r.a) && ids.has(r.b));
  box.innerHTML = graphSVG([...ids].map((id) => ({ id, fixed: id === x.id })), edges, 300, x.id);
}
function openRelGraph() {
  const all = rels(), ids = new Set(); all.forEach((r) => { ids.add(r.a); ids.add(r.b); });
  ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>關係圖</h2><span class="sub"><span class="lg fam"></span>家人　<span class="lg friend"></span>朋友　<span class="lg colleague"></span>同事　<span class="lg referral"></span>介紹</span><button class="btn pri" id="rgClose">關閉</button></div>
   <div class="ovbody" style="justify-content:center">${ids.size ? `<div style="width:100%">${graphSVG([...ids].map((id) => ({ id })), all, Math.max(520, window.innerHeight - 140), S.csel)}</div>` : '<div class="empty">還沒有任何關係。到客戶頁按「＋ 新增關係」。</div>'}</div></div>`;
  $('rgClose').onclick = () => (ov.innerHTML = '');
  ov.querySelectorAll('[data-gonode]').forEach((g) => (g.onclick = () => { ov.innerHTML = ''; S.view = 'clients'; S.csel = g.dataset.gonode; S.ccat = '全部'; S.showDetail = true; S.showIds = false; App.render(); }));
}
function clientListHTML() {
  const q = S.q.trim();
  const list = clients().filter((x) => !q || [x.name, x.phone, x.line, x.note].join(' ').includes(q)).sort((a, b) => (a.name || '').localeCompare(b.name || '', 'zh-Hant'));
  return `<section class="col"><div class="colhead"><h1>客戶 <span style="display:flex;gap:6px;flex-wrap:wrap;justify-content:flex-end"><button class="btn" data-act="relgraph">關係圖</button><button class="btn" data-act="crmimportnew">從 CRM 帶入</button><button class="btn" data-act="newclient">＋ 新增</button></span></h1>
  <div class="mobsync">${syncHTML()}</div>
  <input class="search" id="q" placeholder="搜尋姓名、電話、備註" value="${esc(S.q)}" autocomplete="off"></div>
  <div class="list">${list.length ? list.map((x) => { const n = clientCases(x).length, open = clientCases(x).filter((c) => c.status !== '已結案').length; return `<button class="item${x.id === S.csel ? ' on' : ''}" data-csel="${x.id}"><span class="t"><span>${esc(x.name || '未填姓名')}</span>${open ? `<span class="pill p-wait">${open} 件進行中</span>` : ''}</span><span class="s">${n} 個案件 · 證件 ${(x.docs || []).length} 張</span></button>`; }).join('') : `<div class="empty">${clients().length ? '沒有符合的客戶。' : '還沒有客戶。按「＋ 新增」，或在案件裡填客戶名字就會自動建立。'}</div>`}</div></section>`;
}
function clientDetailHTML() {
  const x = curClient();
  if (!x) return '<div class="empty">從左邊選一位客戶，或按「＋ 新增」。</div>';
  const f = (k, label, extra = '') => `<label class="f">${label}<input id="cf_${k}" data-cf="${k}" value="${esc(x[k])}" autocomplete="off" ${extra}></label>`;
  const docs = x.docs || [], cs = clientCases(x);
  if (S.ccat !== '全部' && !cs.some((c) => c.cat === S.ccat)) S.ccat = '全部';
  return `<button class="back" data-act="back">‹ 客戶列表</button>
  <div class="dhead"><input class="title" id="cf_name" data-cf="name" value="${esc(x.name)}" placeholder="客戶姓名" aria-label="客戶姓名"></div>
  <div class="sec"><h2>基本資料 <button class="btn" data-act="crmimport" style="padding:4px 10px;font-size:13px">從 CRM 帶入</button></h2><div class="grid2">
   <label class="f">性別<select id="cf_gender" data-cf="gender">${['', '女', '男'].map((g) => `<option${g === (x.gender || '') ? ' selected' : ''}>${g}</option>`).join('')}</select></label>
   ${f('birthday', '生日', 'type="date"')}
   <label class="f"><span style="display:flex;justify-content:space-between">身分證字號<button class="linkbtn" data-reveal="cf_idNo">顯示</button></span><input id="cf_idNo" data-cf="idNo" class="secret" value="${esc(x.idNo)}" autocomplete="off" autocapitalize="characters" spellcheck="false"></label>
   ${f('phone', '手機', 'inputmode="tel"')}
   ${f('homeTel', '住家電話', 'inputmode="tel"')}
   ${f('workTel', '公司電話', 'inputmode="tel"')}
   ${f('line', 'LINE ID')}
   ${f('email', 'Email', 'inputmode="email"')}
   ${f('job', '職業')}
   ${f('source', '來源／介紹人')}</div>
   ${f('address', '地址')}
   ${x.billAddress || x.workAddress ? `<div class="grid2">${f('billAddress', '收費地址')}${f('workAddress', '公司地址')}</div>` : ''}
   <label class="f">備註<textarea id="cf_note" class="grow" data-cf="note" rows="3" placeholder="家庭狀況、保單重點、喜好…">${esc(x.note)}</textarea></label></div>
  ${disclosureHTML(x)}
  ${relsHTML(x)}
  <div class="sec"><h2>證件與文件 <button class="btn" data-act="toggleids" style="padding:4px 10px;font-size:13px">${S.showIds ? '隱藏證件' : '顯示證件'}</button>${x.docManual ? '<button class="linkbtn" data-docbydate="c">改回依日期排序</button>' : ''}</h2>
   <div class="docs${S.showIds ? '' : ' veiled'}">${docIdx(x).map((i) => docTile(docs[i], i, 'c', null)).join('')}
   <button class="adddoc" data-act="cscan">${ICON.scan}拍身分證、存摺<span style="font-size:11px">也可以拖曳圖片、PDF 進來</span></button><button class="adddoc" data-act="pasteimg">📋 貼上圖片</button><label class="adddoc">＋ 上傳 PDF 或圖片<input type="file" id="docUpC" accept="image/*,application/pdf,.pdf" multiple></label></div>
   ${docs.some((d) => !isPdf(d)) ? '<div class="actions"><button class="btn" data-act="cpdf">選頁做成 PDF／列印</button><button class="btn" data-act="cshare">匯出遮罩版</button></div>' : ''}</div>
  ${giftSectionHTML('client', x.id)}
  <div class="sec"><h2>案件 <button class="btn" data-act="ccase" style="padding:4px 10px;font-size:13px">＋ 新增${S.ccat !== '全部' ? esc(S.ccat) : '案件'}</button></h2>
   ${cs.length ? `<div class="chips">${['全部', ...CATS.filter((k) => cs.some((c) => c.cat === k))].map((k) => `<button class="chip${S.ccat === k ? ' on' : ''}" data-ccat="${k}">${k} <span class="cnt">${k === '全部' ? cs.length : cs.filter((c) => c.cat === k).length}</span></button>`).join('')}</div>` : ''}
   ${cs.length ? cs.filter((c) => S.ccat === '全部' || c.cat === S.ccat).map((c) => `<button class="item${itemCls(c.status)}" data-gocase="${c.id}"><span class="t"><span>${esc(c.title || '未命名')}</span><span class="pill ${pillCls(c.status)}">${c.status}</span></span>${casePeopleLine(c, x.id)}<span class="s">${fmtDay(caseDate(c))} · ${esc(caseSummary(c))}${c.next ? ' · 下一步：' + esc(c.next) : ''}</span></button>`).join('') : '<div class="sub">還沒有案件。</div>'}</div>
  <div class="actions"><button class="btn pri savebtn" data-act="savenow">儲存</button><button class="btn danger" data-act="delclient">刪除客戶</button></div>`;
}
const isPdf = (d) => d.type === 'pdf';
function fillTpl(body, c) { return String(body || '').replace(/\{日期\}/g, fmtDay(localDay())).replace(/\{客戶\}/g, (c?.client || '').trim() || '{客戶}'); }
function caseTplHTML(c) {
  const sops = notes().filter((n) => n.kind === 'sop' && catsOf(n).includes(c.cat) && lineOk(n, lineOf(c)));
  const ts = notes().filter((n) => n.kind === 'tpl' && catsOf(n).includes(c.cat) && lineOk(n, lineOf(c)) && !sopsOf(n).some((id) => sops.some((x) => x.id === id))).sort((a, b) => TPLGROUPS.indexOf(a.group) - TPLGROUPS.indexOf(b.group));
  return `${sops.length ? `<div class="sec"><h2>「${esc(c.cat)}」SOP <button class="linkbtn" data-act="sopfor">管理 SOP</button></h2>${sops.map((n) => `<details class="tplrow sopcase"><summary><span class="tpltitle">${esc(n.title)}</span>${tagPills(n)}</summary><pre>${linkify(esc(n.body))}</pre>${nimgsHTML(n)}${sopTplsHTML(n, c)}</details>`).join('')}</div>` : ''}
  <div class="sec"><h2>「${esc(c.cat)}」範本 <button class="linkbtn" data-act="tplfor">${ts.length ? '管理範本' : '＋ 新增範本'}</button></h2>
   ${ts.length ? ts.map((n) => `<details class="tplrow"><summary>${chPills(n)}<span class="tpltitle">${esc(n.title)}</span><button class="btn" data-copy="${n.id}" data-forcase="1" style="padding:4px 10px;font-size:13px">複製</button></summary><pre>${esc(fillTpl(n.body, c))}</pre></details>`).join('') : `<div class="sub">還沒有和「${esc(c.cat)}」相關的 LINE、簡訊範本。</div>`}</div>`;
}
function docTile(d, i, p, folders) { // p：'' 案件、'c' 客戶
  const media = isPdf(d) ? `<div class="pdftile">PDF<small>${esc(d.name)}</small></div>` : `<canvas data-${p}thumb="${i}" width="240" height="320"></canvas>`;
  const mv = folders ? `<select class="mv" data-movedoc="${i}" aria-label="移到資料夾"><option value="">未分類</option>${folders.map((f) => `<option${f === d.folder ? ' selected' : ''}>${esc(f)}</option>`).join('')}</select>` : '';
  const dt = `<input type="date" class="ddate" data-${p}docdate="${i}" value="${esc(d.date || '')}" aria-label="文件日期">`;
  return `<div class="doc" data-di="${i}" data-dp="${p}">${media}${dt}${d.note ? `<div class="docnote" title="${esc(d.note)}">📝 ${esc(d.note)}</div>` : ''}<div class="bar"><button class="docname" data-${p}rename="${i}" title="改名">${esc(d.name)}${d.masks?.length ? ' · 已遮 ' + d.masks.length : ''}</button><span><button data-${p}view="${i}" hidden>看</button>${isPdf(d) ? `<button data-${p}pdfimg="${i}" title="轉成圖片頁之後就能標註、蓋章">轉成圖片頁</button>` : `<button data-${p}rot="${i}" aria-label="向右轉 90 度" title="轉方向">↻</button><button data-${p}mask="${i}">標註</button>`}<button data-${p}deldoc="${i}" aria-label="刪除文件">刪</button></span></div>${mv}</div>`;
}
const STEPS = {
  新契約: ['建議書', 'keyman 再解說', '簽約', '送件', '文件補全', '核保通過', '簽收回條', '紙本保單'],
  建議書: ['建議書', 'keyman 再解說', '簽約'],
  理賠: ['收件', '送件', '補件通知', '補件完成', '給付'],
  保全: ['收件', '送件', '補件', '完成'],
  產險: ['報價', '投保', '送件', '保單'],
};
const DEFAULT_STEPS = ['聯絡', '送件', '補件', '完成'];
function getSteps(cat) { const m = Vault.get('settings', 'steps')?.map || {}; return m[cat] || STEPS[cat] || DEFAULT_STEPS; }
function setSteps(cat, arr) { const st = Vault.get('settings', 'steps') || { id: 'steps', map: {} }; st.map = { ...(st.map || {}), [cat]: arr }; Vault.put('settings', st); }
function openStepEditor(cat) {
  const arr = getSteps(cat);
  ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>「${esc(cat)}」的常用步驟</h2><button class="btn" id="stCancel">取消</button><button class="btn pri" id="stSave">儲存</button></div>
  <div class="ovbody" style="flex-direction:column;max-width:560px;margin:0 auto;width:100%">
   <div class="note">一行一個步驟，照你想要的順序排。不常用的不用放，記錄時按「＋ 其他」打字就好。</div>
   <label class="f">常用步驟<textarea id="stText" rows="12">${esc(arr.join('\n'))}</textarea></label>
   <div class="actions"><button class="btn" id="stReset">恢復預設</button></div></div></div>`;
  $('stCancel').onclick = () => (ov.innerHTML = '');
  $('stReset').onclick = () => { $('stText').value = (STEPS[cat] || DEFAULT_STEPS).join('\n'); };
  $('stSave').onclick = () => { const list = [...new Set($('stText').value.split('\n').map((x) => x.trim()).filter(Boolean))]; setSteps(cat, list); ov.innerHTML = ''; App.render(); toast('常用步驟已更新'); };
  setTimeout(() => $('stText')?.focus(), 50);
}
function fitNote(el) { if (el.dataset.manual) return; el.style.height = 'auto'; el.style.height = el.scrollHeight + 2 + 'px'; }
function fitNotes() { document.querySelectorAll('textarea.tlnote, textarea.grow').forEach(fitNote); }
const tlSorted = (c) => (c.timeline || []).slice().sort((a, b) => (b.date || '').localeCompare(a.date || '') || (b.at || '').localeCompare(a.at || ''));
const latestStep = (c) => tlSorted(c)[0];
function timelineHTML(c) {
  const tl = tlSorted(c), done = new Set(tl.map((e) => e.title)), steps = getSteps(c.cat);
  return `<div class="sec"><h2>進度 <span class="sub">${tl.length ? tl.length + ' 筆，新的在上面' : '每個時間點記一筆'}</span></h2>
   <div class="tags">${steps.map((x) => `<button class="tag${done.has(x) ? ' did' : ''}" data-addstep="${esc(x)}">${done.has(x) ? '✓' : '＋'} ${esc(x)}</button>`).join('')}<button class="tag" data-addstep="">＋ 其他</button><button class="tag tagedit" data-act="editsteps">⚙ 編輯常用步驟</button></div>
   ${tl.length ? `<div class="tl">${tl.map((e) => `<div class="tli"><span class="tldot"></span><div class="tlrow"><input type="date" class="tldate" data-tldate="${e.id}" value="${esc(e.date)}" aria-label="日期"><input class="tltitle" data-tltitle="${e.id}" id="tlt_${e.id}" value="${esc(e.title)}" placeholder="做了什麼" aria-label="項目">${e.title && !steps.includes(e.title) ? `<button class="linkbtn" data-favstep="${esc(e.title)}" style="flex:none">＋ 常用</button>` : ''}<button class="tldel" data-tldel="${e.id}" aria-label="刪除這筆">✕</button></div><textarea class="tlnote" data-tlnote="${e.id}" id="tln_${e.id}" rows="1" placeholder="備註（可不填，Enter 換行）" aria-label="備註">${esc(e.note)}</textarea></div>`).join('')}</div>` : ''}</div>`;
}
function linksHTML(c) {
  const linked = (c.links || []).map((id) => Vault.get('cases', id)).filter(Boolean).sort(byDate);
  return `<div class="sec"><h2>相關案件 <button class="btn" data-act="linkcase" style="padding:4px 10px;font-size:13px">＋ 連結案件</button></h2>
   ${linked.length ? linked.map((x) => `<div class="linkrow"><button class="item" data-gocase="${x.id}"><span class="t"><span>${esc(x.client || '未填客戶')} · ${esc(x.title || '未命名')}</span><span class="pill ${pillCls(x.status)}">${x.status}</span></span><span class="s">${fmtDay(caseDate(x))} · ${esc(caseSummary(x))}</span></button><button class="linkbtn" data-unlink="${x.id}">取消連結</button></div>`).join('') : '<div class="sub">隔很久的同一件事，可以拆成新案件再連起來，例如去年的建議書和今年的新契約。</div>'}</div>`;
}
function openLinkPicker(c) {
  const draw = (q = '') => {
    const pool = cases().filter((x) => x.id !== c.id && !(c.links || []).includes(x.id) && (!q || [x.client, x.title, x.notes].join(' ').includes(q)))
      .sort((a, b) => ((b.clientId && b.clientId === c.clientId) - (a.clientId && a.clientId === c.clientId)) || byDate(a, b));
    $('lkList').innerHTML = pool.length ? pool.slice(0, 60).map((x) => `<button class="item" data-pick="${x.id}"><span class="t"><span>${esc(x.client || '未填客戶')} · ${esc(x.title || '未命名')}</span>${x.clientId && x.clientId === c.clientId ? '<span class="pill p-go">同一位客戶</span>' : ''}</span><span class="s">${fmtDay(caseDate(x))} · ${esc(caseSummary(x))}</span></button>`).join('') : '<div class="empty">沒有其他案件可以連結。</div>';
    $('lkList').querySelectorAll('[data-pick]').forEach((b) => (b.onclick = () => {
      const o = Vault.get('cases', b.dataset.pick);
      c.links = [...new Set([...(c.links || []), o.id])]; o.links = [...new Set([...(o.links || []), c.id])];
      saveCase(o); saveCase(c); ov.innerHTML = ''; App.render(); toast('已連結');
    }));
  };
  ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>連結案件</h2><button class="btn pri" id="lkClose">關閉</button></div><div class="ovbody" style="flex-direction:column;max-width:640px;margin:0 auto;width:100%"><input class="search" id="lkQ" placeholder="搜尋客戶、案件名稱" autocomplete="off"><div class="list" id="lkList" style="padding:0"></div></div></div>`;
  $('lkClose').onclick = () => (ov.innerHTML = '');
  $('lkQ').oninput = (e) => draw(e.target.value.trim());
  draw();
}
// 文件顯示順序：手動排過就照陣列順序，否則新到舊
function docIdx(rec) { const docs = rec?.docs || []; const ix = docs.map((_, i) => i); return rec?.docManual ? ix : ix.sort((a, b) => (docs[b].date || '').localeCompare(docs[a].date || '')); }
function curFolder(c) { return (c?.folders || []).includes(S.folder) ? S.folder : ''; }
function caseDocsHTML(c) {
  const folders = c.folders || [];
  if (!['全部', '未分類', ...folders].includes(S.folder)) S.folder = '全部';
  const cnt = (f) => c.docs.filter((d) => (f === '全部' ? true : f === '未分類' ? !d.folder : d.folder === f)).length;
  const shown = docIdx(c).map((i) => [c.docs[i], i]).filter(([d]) => S.folder === '全部' || (S.folder === '未分類' ? !d.folder : d.folder === S.folder));
  const hasImg = c.docs.some((d) => !isPdf(d));
  return `<div class="sec"><h2>文件 <span class="sub">原圖永遠保留，遮罩可以隨時改${c.docs.length > 1 ? '・按住文件拖曳可排順序' + (folders.length ? '、拖到上面資料夾' : '') : ''}</span>${c.docManual ? '<button class="linkbtn" data-docbydate="">改回依日期排序</button>' : ''}</h2>
   <div class="chips">${['全部', ...(folders.length ? ['未分類'] : []), ...folders].map((f) => `<button class="chip${S.folder === f ? ' on' : ''}" data-folder="${esc(f)}">${esc(f)} <span class="cnt">${cnt(f)}</span></button>`).join('')}<button class="chip" data-act="newfolder">＋ 資料夾</button></div>
   ${folders.includes(S.folder) ? `<div class="actions"><button class="linkbtn" data-act="renfolder">重新命名資料夾</button><button class="linkbtn" data-act="delfolder">刪除資料夾</button><span class="sub">新加的文件會放進「${esc(S.folder)}」</span></div>` : ''}
   <div class="docs">${shown.map(([d, i]) => docTile(d, i, '', folders.length ? folders : null)).join('')}
   <button class="adddoc" data-act="scan">${ICON.scan}拍照／掃描<span style="font-size:11px">也可以拖曳圖片、PDF 進來</span></button>
   <button class="adddoc" data-act="pasteimg">📋 貼上圖片<span style="font-size:11px">電腦可直接按 Ctrl+V</span></button>
   <label class="adddoc">＋ 上傳 PDF 或圖片<input type="file" id="docUp" accept="image/*,application/pdf,.pdf" multiple></label></div>
   ${hasImg ? `<div class="actions"><button class="btn" data-act="pdf">選頁做成 PDF／列印</button><button class="btn" data-act="share">匯出分享版</button>${c.isRef ? `<button class="btn" data-act="pickref">選理賠參考要放的文件（已選 ${refDocs(c).length} 張）</button>` : ''}</div>` : ''}</div>`;
}
async function addPdfs(rec, coll, files, folder) {
  let n = 0;
  for (const f of files) {
    if (f.size > 60 * 1024 * 1024) { toast(`「${f.name}」超過 60MB，太大了`); continue; }
    const fileId = await Files.add(new Blob([await f.arrayBuffer()], { type: 'application/pdf' }), f.name);
    rec.docs = rec.docs || []; rec.docs.push({ id: uid(), name: f.name.replace(/\.pdf$/i, ''), fileId, masks: [], type: 'pdf', folder: folder || '', date: localDay() }); n++;
  }
  if (n) { Vault.put(coll, rec); App.render(); toast(`已加入 ${n} 個 PDF`); }
}
function refCases() { return cases().filter((c) => c.isRef && (!S.refLine || lineOf(c) === S.refLine) && (S.refItem === '全部' || c.items.includes(S.refItem)) && (S.refCause === '全部' || c.causeType === S.refCause) && (!S.refDet || c.causeDetail === S.refDet)).sort(byDate); }
const refTitle = (c) => c.diag || c.causeDetail || c.causeType || '未分類';
function refListHTML() {
  const list = refCases(); const used = [...new Set(cases().filter((c) => c.isRef && (!S.refLine || lineOf(c) === S.refLine)).flatMap((c) => c.items))];
  return `<section class="col"><div class="colhead"><h1>理賠參考</h1>
  <div class="chips">${['全部', ...LINES].map((x) => `<button class="chip${(S.refLine || '全部') === x ? ' on' : ''}" data-rline="${x}">${x}</button>`).join('')}</div>
  <div class="sub">險種</div><div class="chips">${['全部', ...used].map((x) => `<button class="chip${S.refItem === x ? ' on' : ''}" data-ritem="${x}">${x}</button>`).join('')}</div>
  <div class="sub">原因</div><div class="chips">${['全部', ...Object.keys(CAUSE)].map((x) => `<button class="chip${S.refCause === x ? ' on' : ''}" data-rcause="${x}">${x}</button>`).join('')}</div>
  ${S.refCause !== '全部' ? (() => { const used = [...new Set(cases().filter((c) => c.isRef && c.causeType === S.refCause && c.causeDetail).map((c) => c.causeDetail))]; return used.length ? `<div class="chips subchips">${used.map((x) => `<button class="chip${S.refDet === x ? ' on' : ''}" data-rdet="${esc(x)}">${esc(x)}</button>`).join('')}</div>` : ''; })() : ''}</div>
  <div class="list">${list.length ? list.map((c) => `<button class="item${c.id === S.refSel ? ' on' : ''}" data-rsel="${c.id}"><span class="t"><span>${esc(refTitle(c))}</span></span><span class="s">${esc(c.items.join('、') || '未選險種')} · ${esc(c.causeType)}${c.causeDetail ? '›' + esc(c.causeDetail) : ''}</span></button>`).join('') : '<div class="empty">還沒有理賠參考。在理賠案件裡按「做成理賠參考」就會出現在這裡。</div>'}</div></section>`;
}
// 理賠參考要放哪些文件：可以選資料夾；沒選過的話，有「理賠」字樣的資料夾就只放那些，否則全部
function refFolderSel(c) { if (c.refFolders === 'all') return null; if (c.refFolders) return c.refFolders; const fs = (c.folders || []).filter((f) => f.includes('理賠') && c.docs.some((d) => d.folder === f)); return fs.length ? fs : null; }
const dkey = (d) => d.id || d.fileId;
const refDocs = (c) => { const sel = refFolderSel(c), pick = c.refDocIds; return c.docs.map((d, i) => [d, i]).filter(([d]) => !isPdf(d) && (pick ? pick.includes(dkey(d)) : !sel || sel.includes(d.folder || ''))); };
function refDetailHTML() {
  const c = cases().find((x) => x.id === S.refSel && x.isRef);
  if (!c) return '<div class="empty">選一則理賠參考。這裡不會顯示客戶姓名。</div>';
  const cost = +c.cost || 0, ratio = +c.ratio || 0, auto = cost && ratio ? Math.round(cost * ratio / 100) : 0;
  return `<button class="back" data-act="back">‹ 理賠參考</button>
  <div class="dhead"><h1 style="margin:0;font-size:20px">${esc(refTitle(c))}</h1></div>
  <div class="sec"><h2>分類</h2><div class="reftop">${c.items.map((x) => `<span class="pill p-go">${esc(x)}</span>`).join('')}${c.causeType ? `<span class="pill p-wait">${esc(c.causeType)}${c.causeDetail ? ' › ' + esc(c.causeDetail) : ''}</span>` : ''}</div></div>
  ${cost ? `<div class="sec"><h2>理賠金額</h2><div class="calcline">花費 ${cost.toLocaleString()} × ${ratio || '—'}% = ${auto ? auto.toLocaleString() : '—'}${c.payout ? `　實際理賠 ${(+c.payout).toLocaleString()}` : ''}</div></div>` : ''}
  ${c.docs.some((d) => !isPdf(d)) ? (() => { const rd = refDocs(c), on = new Set(rd.map(([d]) => dkey(d))), all = c.docs.map((d, i) => [d, i]).filter(([d]) => !isPdf(d));
   if (!S.refPick) return `<div class="sec"><h2>文件（遮罩版）<span class="sub">${rd.length} 張・點圖片可以放大、左右滑</span><button class="linkbtn" data-act="refpickmode">選擇要放的文件</button></h2>
     <div class="docs">${rd.map(([d, i], k) => `<div class="doc refview" data-refview="${k}"><canvas data-rthumb="${i}" width="240" height="320"></canvas>${d.note ? `<div class="docnote">${esc(d.note)}</div>` : ''}</div>`).join('') || '<div class="sub">還沒有選文件，按右上「選擇要放的文件」。</div>'}</div>${rd.length ? '<div class="actions"><button class="btn pri" data-act="rshare">匯出分享版</button></div>' : ''}</div>`;
   return `<div class="sec"><h2>選擇要放的文件<span class="sub">選了 ${on.size} / ${all.length} 張</span><button class="btn pri" data-act="refpickmode" style="padding:4px 12px;font-size:13px">完成</button></h2>
   <div class="actions" style="margin:0 0 8px"><span class="sub">點文件選或取消</span><button class="linkbtn" data-refpick="all">全選</button><button class="linkbtn" data-refpick="none">全不選</button>${(c.folders || []).some((f) => f.includes('理賠')) ? '<button class="linkbtn" data-refpick="claim">只選理賠資料夾</button>' : ''}</div>
   <div class="docs">${all.map(([d, i]) => `<div class="doc refpick${on.has(dkey(d)) ? ' on' : ''}" data-refdoc="${esc(dkey(d))}"><canvas data-rthumb="${i}" width="240" height="320"></canvas><span class="sdcheck">✓</span>${d.note ? `<div class="docnote">${esc(d.note)}</div>` : ''}</div>`).join('')}</div></div>`; })() : ''}
  <div class="note">這頁只顯示分類、計算和遮罩後的文件，不顯示客戶姓名。要修改內容請回到「案件」。</div>
  <div class="actions"><button class="btn" data-act="gocase">打開原案件</button></div>`;
}
function sopMatch(n) {
  const ks = catsOf(n); if (!lineOk(n, S.nline)) return false;
  if (S.sg === '通用') { if (ks.length) return false; } else if (S.sg !== '全部') { const g = GROUPS.find((x) => x.name === S.sg); if (!g || !ks.some((k) => g.cats.includes(k))) return false; if (S.sc !== '全部' && !ks.includes(S.sc)) return false; }
  if (S.sk && !kindsOf(n).includes(S.sk)) return false;
  if (S.spk && !(n.pkinds || []).includes(S.spk)) return false;
  if (S.smk && !(n.mkinds || []).includes(S.smk)) return false;
  if (S.st && !toolsOf(n).includes(S.st)) return false;
  return true;
}
const tagPills = (n) => [...kindsOf(n), ...(n.pkinds || []), ...(n.mkinds || []), ...toolsOf(n)].map((t) => `<span class="tagpill">${esc(t)}</span>`).join('');
function sopTplsHTML(sop, c) { // 掛在這個 SOP 底下、要傳給客戶的範本
  const ts = notes().filter((t) => t.kind === 'tpl' && sopsOf(t).includes(sop.id));
  return `<div class="soptpls">${ts.map((t) => `<details class="tplrow"><summary>${chPills(t)}<span class="tpltitle">${esc(t.title)}</span><button class="btn" data-copy="${t.id}"${c ? ' data-forcase="1"' : ''} style="padding:4px 10px;font-size:13px">複製</button></summary><pre>${esc(fillTpl(t.body, c))}</pre></details>`).join('')}<button class="linkbtn" data-newtplfor="${sop.id}">＋ 給客戶的範本</button></div>`;
}
function sopHTML() {
  const q = S.q.trim(), list = notes().filter((n) => n.kind === 'sop' && sopMatch(n) && (!q || (n.title + n.body).includes(q)));
  const g = GROUPS.find((x) => x.name === S.sg);
  const tagRow = (key, cur, attr) => `<div class="chips"><span class="sub" style="align-self:center">${TAGS[key].label}</span>${tagList(key).map((t) => `<button class="chip${cur === t ? ' on' : ''}" data-${attr}="${esc(t)}">${esc(t)}</button>`).join('')}<button class="chip chipedit" data-act="edittags-${key}">⚙</button></div>`;
  return `<h1 style="display:flex;justify-content:space-between;align-items:center;gap:8px">SOP 作業流程<span class="actions" style="margin:0"><button class="btn" data-act="import">批次匯入</button><button class="btn pri" data-act="newsop">＋ 新增 SOP</button></span></h1>
  <input class="search" id="q" placeholder="搜尋" value="${esc(S.q)}" style="max-width:420px" autocomplete="off">
  ${lineFilter('nline')}<div class="chips">${['全部', ...GROUPS.map((x) => x.name), '通用'].map((x) => `<button class="chip${S.sg === x ? ' on' : ''}" data-sg="${esc(x)}">${esc(x)}</button>`).join('')}<button class="chip chipedit" data-act="editcats">⚙ 編輯分類</button></div>
  ${g ? `<div class="chips subchips">${['全部', ...g.cats].map((x) => `<button class="chip${S.sc === x ? ' on' : ''}" data-sc="${esc(x)}">${esc(x)}</button>`).join('')}</div>` : ''}
  ${tagShown('kinds', S.nline) ? tagRow('kinds', S.sk, 'sk') : ''}${tagShown('tools', S.nline) ? tagRow('tools', S.st, 'st') : ''}${tagShown('pkinds', S.nline) ? tagRow('pkinds', S.spk, 'spk') : ''}${tagShown('mkinds', S.nline) ? tagRow('mkinds', S.smk, 'smk') : ''}
  <div class="cards">${list.map((n) => { const open = S.sopOpen === n.id, long = (n.body || '').split('\n').length > 8 || (n.body || '').length > 260 || (n.images || []).length > 4; return `<div class="card sopcard${open ? ' open' : ''}"><h3><span>${esc(n.title)}</span><span class="sub">${linePill(lineN(n))} ${esc(catsLabel(n))}</span></h3>${tagPills(n) ? `<div class="tagpills">${tagPills(n)}</div>` : ''}<pre>${linkify(esc(n.body))}</pre>${nimgsHTML(n, open ? 0 : 4)}${long ? `<button class="linkbtn" data-sopopen="${n.id}">${open ? '收起' : '展開全部'}</button>` : ''}${sopTplsHTML(n, null)}<div class="mini">${/https?:\/\//.test(n.body) ? `<button data-copy="${n.id}">複製</button>` : ''}<button data-editnote="${n.id}">編輯</button><button data-delnote="${n.id}">刪除</button></div></div>`; }).join('') || `<div class="empty">${notes().some((n) => n.kind === 'sop') ? '這個篩選沒有符合的 SOP。' : '還沒有流程。按右上角「＋ 新增 SOP」。'}</div>`}</div>
  `;
}
function notesHTML(kind) {
  if (kind === 'sop') return sopHTML();
  const tg = GROUPS.find((x) => x.name === S.tg);
  const shown = notes().filter((n) => {
    if (n.kind !== 'tpl' || !lineOk(n, S.nline)) return false;
    const ks = catsOf(n);
    if (S.tg === '通用') { if (ks.length) return false; } else if (tg) { if (!ks.some((k) => tg.cats.includes(k))) return false; if (S.tplCat !== '全部' && !ks.includes(S.tplCat)) return false; }
    if (S.tplGroup !== '全部' && !chOf(n).includes(S.tplGroup)) return false;
    const q = S.q.trim(); return !q || (n.title + n.body).includes(q);
  });
  return `<h1 style="display:flex;justify-content:space-between;align-items:center;gap:8px">話術範本<span class="actions" style="margin:0"><button class="btn" data-act="import">批次匯入</button><button class="btn pri" data-act="newtpl">＋ 新增範本</button></span></h1>
  <input class="search" id="q" placeholder="搜尋" value="${esc(S.q)}" style="max-width:420px" autocomplete="off">
  ${lineFilter('nline')}<div class="chips">${['全部', ...GROUPS.map((x) => x.name), '通用'].map((x) => `<button class="chip${S.tg === x ? ' on' : ''}" data-tgg="${esc(x)}">${esc(x)}</button>`).join('')}<button class="chip chipedit" data-act="editcats">⚙ 編輯分類</button></div>
  ${tg ? `<div class="chips subchips">${['全部', ...tg.cats].map((x) => `<button class="chip${S.tplCat === x ? ' on' : ''}" data-tc="${esc(x)}">${esc(x)}</button>`).join('')}</div>` : ''}
  <div class="chips"><span class="sub" style="align-self:center">傳送方式</span>${TPLGROUPS.map((g) => `<button class="chip${S.tplGroup === g ? ' on' : ''}" data-tg="${esc(g)}">${esc(g)}</button>`).join('')}</div>
  <div class="cards">${shown.map((n) => `<div class="card"><h3><span>${esc(n.title)}</span><span class="sub">${linePill(lineN(n))} ${esc(catsLabel(n))}</span></h3><div class="tagpills">${chPills(n)}</div>${sopsOf(n).map((id) => Vault.get('notes', id)).filter(Boolean).map((x) => `<div class="sub">↳ SOP：${esc(x.title)}</div>`).join('')}<pre>${linkify(esc(n.body))}</pre>${nimgsHTML(n)}<div class="mini">${kind === 'tpl' || /https?:\/\//.test(n.body) ? `<button data-copy="${n.id}">複製</button>` : ''}<button data-editnote="${n.id}">編輯</button><button data-delnote="${n.id}">刪除</button></div></div>`).join('') || `<div class="empty">${kind === 'sop' ? '還沒有流程。用下面的表單把 Notion、LINE 記事本裡的步驟加進來。' : '還沒有範本。'}</div>`}</div>
  `;
}
function quickHTML() {
  const phones = notes().filter((n) => n.kind === 'quick');
  const waiting = cases().filter((c) => c.status !== '已結案').sort((a, b) => (b.updatedAt || '').localeCompare(a.updatedAt || ''));
  return `<h1>快查</h1>
  <input class="search" id="qq" placeholder="搜尋全部 SOP、案件、範本" style="max-width:520px" autocomplete="off">
  <div id="qres"></div>
  <div class="cards">
  <div class="card"><h3>常用電話</h3>${phones.map((p) => `<div style="display:flex;justify-content:space-between;gap:8px;align-items:center"><span>${esc(p.title)}</span><span style="font-family:var(--f-num)">${esc(p.body)} <button class="btn" style="padding:2px 8px;font-size:12px" data-editnote="${p.id}">改</button></span></div>`).join('') || '<div class="sub">還沒有電話。</div>'}
   <div class="mini"><button data-act="addphone">＋ 新增電話</button></div></div>
  ${(() => { const t = localDay(), due = calls().filter((k) => k.nextDate && k.nextDate <= t && k.status !== '結案').sort((a, b) => a.nextDate.localeCompare(b.nextDate)); return due.length ? `<div class="card"><h3>今天要回撥 <span class="sub">${due.length} 位</span></h3>${due.slice(0, 6).map((k) => `<button class="item" data-gocall="${k.id}"><span class="t"><span>${esc(k.name)}</span><span class="pill ${stCls(k.status)}">${esc(k.status)}</span></span><span class="s">${k.nextDate < t ? '<span class="due">已過 ' + fmtDay(k.nextDate).slice(5) + '</span>' : '今天'}${k.source ? ' · ' + esc(k.source) : ''}</span></button>`).join('')}</div>` : ''; })()}
  <div class="card"><h3>待辦案件 <span class="sub">${waiting.length} 件</span></h3>${waiting.slice(0, 6).map((c) => `<button class="item${itemCls(c.status)}" data-gocase="${c.id}"><span class="t"><span>${esc(c.client)} · ${esc(c.title)}</span><span class="pill ${pillCls(c.status)}">${c.status}</span></span><span class="s">${esc(c.next || '未填下一步')}</span></button>`).join('') || '<div class="sub">沒有待辦。</div>'}</div>
  <div class="card"><h3>同步</h3>${syncHTML()}<div class="sub">資料加密後存在 OneDrive › 應用程式 › Work › ${esc(FOLDER)}。iPad 和手機登入同一個帳號就會自動同步。每天第一次打開會自動備份一份，保留 30 天。</div>
   <div class="actions"><button class="btn" data-act="backups">備份與還原</button><button class="btn" data-act="trash">最近刪除${Files.trashed().filter((f) => f.trashCtx).length ? ' ' + Files.trashed().filter((f) => f.trashCtx).length : ''}</button></div>
   <div class="actions"><button class="btn" data-act="import">批次匯入</button><button class="btn" data-act="exportnotes">複製新契約備註（去掉姓名）</button><button class="btn" data-act="lock">鎖定</button><button class="btn" data-act="signout">登出</button></div></div>
  </div>`;
}

/* ---------- 縮圖 ---------- */
// 文件上用筆畫的記號（畫筆、圈選），座標都是 0~1 的比例，跟原圖分開存
const MARK_COLORS = ['#E53935', '#1E6FD9', '#2E9E4F', '#F2A900', '#111111'];
/* ---------- 印章：文字章、標籤長條、圓章、色塊、大括號、底線（座標 0~1，跟原圖分開存） ---------- */
const STAMP_FONT = '"PingFang TC","Noto Sans TC","Microsoft JhengHei",sans-serif';
const STAMP_COLORS = ['#E24B4A', '#185FA5', '#0F6E56', '#D85A30', '#BA7517', '#993556', '#2B2724', '#FAC775', '#9FE1CB', '#F4C0D1', '#C0DD97', '#B5D4F4'];
const STAMP_KINDS = [['text', '文字'], ['uline', '文字＋底線'], ['tag', '標籤長條'], ['dot', '圓章'], ['hl', '色塊'], ['brace', '大括號'], ['line', '底線']];
const STAMP_DEFAULT = [
  ['險種', [['text', '壽險', '#0F6E56'], ['text', '意外', '#E24B4A'], ['text', '疾病', '#185FA5'], ['text', '實支', '#185FA5'], ['text', '住院', '#185FA5'], ['text', '手術', '#185FA5'], ['text', '骨折', '#E24B4A'], ['text', '無呼吸', '#185FA5'], ['text', '癌／重大', '#D85A30'], ['text', '長照', '#0F6E56']]],
  ['保障', [['uline', '意外', '#E24B4A'], ['uline', '疾病', '#185FA5']]],
  ['期間', [['text', '定期', '#185FA5'], ['text', '終身', '#185FA5'], ['dot', '定', '#185FA5'], ['dot', '終', '#1D9E75'], ['text', '繳費期間', '#D85A30']]],
  ['標籤', [['tag', '主約', '#FAC775'], ['tag', '附約', '#C0DD97'], ['text', '見卡即賠', '#E24B4A']]],
  ['色塊、括號', [['hl', '', '#FAC775'], ['hl', '', '#9FE1CB'], ['hl', '', '#F4C0D1'], ['brace', '', '#E24B4A'], ['line', '', '#E24B4A']]],
];
const stampAlpha = (k) => (k === 'hl' ? 0.45 : k === 'tag' ? 0.9 : 1);
function stampLib() { const st = Vault.get('settings', 'stamps'); return st?.groups ? st : stampDefaultLib(); }
function stampDefaultLib() {
  let i = 0; return { id: 'stamps', groups: STAMP_DEFAULT.map(([g, l]) => ({ g, items: l.map(([k, t, c]) => ({ id: 'st' + (++i), k, s: t, c, a: stampAlpha(k) })) })), recent: [] };
}
const tagInk = (c) => { const n = parseInt(c.slice(1), 16), r = n >> 16, g = (n >> 8) & 255, b = n & 255; return 0.3 * r + 0.59 * g + 0.11 * b > 150 ? '#3b2a08' : '#ffffff'; };
function drawStamp(x, m, ox, oy, w, h) {
  const X = ox + m.x * w, Y = oy + m.y * h, BW = m.w * w, BH = m.h * h; if (BH <= 0) return;
  x.save(); x.globalAlpha = m.a ?? 1; x.fillStyle = m.c; x.strokeStyle = m.c; x.textBaseline = 'middle';
  const font = (px) => `bold ${px}px ${STAMP_FONT}`;
  if (m.k === 'hl') { x.fillRect(X, Y, BW, BH); }
  else if (m.k === 'line') { x.lineWidth = Math.max(1.5, BH); x.lineCap = 'round'; x.beginPath(); x.moveTo(X, Y + BH / 2); x.lineTo(X + BW, Y + BH / 2); x.stroke(); }
  else if (m.k === 'brace') {
    x.lineWidth = Math.max(1.5, BW * 0.22); x.lineCap = 'round'; x.lineJoin = 'round'; const r = BW, my = Y + BH / 2, L = X + BW * 0.15, Rr = X + BW;
    x.beginPath(); x.moveTo(Rr, Y); x.quadraticCurveTo(L + r * 0.35, Y, L + r * 0.35, Y + Math.min(r, BH * 0.15)); x.lineTo(L + r * 0.35, my - Math.min(r, BH * 0.12)); x.quadraticCurveTo(L + r * 0.35, my, L - r * 0.05, my);
    x.quadraticCurveTo(L + r * 0.35, my, L + r * 0.35, my + Math.min(r, BH * 0.12)); x.lineTo(L + r * 0.35, Y + BH - Math.min(r, BH * 0.15)); x.quadraticCurveTo(L + r * 0.35, Y + BH, Rr, Y + BH); x.stroke();
  }
  else if (m.k === 'dot') { const d = Math.min(BW, BH); x.beginPath(); x.arc(X + d / 2, Y + d / 2, d / 2, 0, Math.PI * 2); x.fill(); x.globalAlpha = 1; x.fillStyle = '#fff'; x.font = font(d * 0.6); x.textAlign = 'center'; x.fillText(m.s || '', X + d / 2, Y + d / 2 + d * 0.03); }
  else if (m.k === 'tag') { const r = BH * 0.22; x.beginPath(); x.roundRect ? x.roundRect(X, Y, BW, BH, r) : x.rect(X, Y, BW, BH); x.fill(); x.globalAlpha = 1; x.fillStyle = tagInk(m.c); x.font = font(BH * 0.68); x.fillText(m.s || '', X + BH * 0.3, Y + BH / 2 + BH * 0.03); }
  else { x.font = font(BH * 0.86); x.fillText(m.s || '', X, Y + BH / 2 + BH * 0.03); if (m.k === 'uline') { x.lineWidth = Math.max(1.5, BH * 0.07); x.beginPath(); x.moveTo(X, Y + BH * 0.98); x.lineTo(X + BW, Y + BH * 0.98); x.stroke(); } }
  x.restore();
}
function stampFit(m, W, H) { // 文字章依文字長度決定寬度
  if (!['text', 'uline', 'tag', 'dot'].includes(m.k)) return m;
  const c = document.createElement('canvas').getContext('2d'); const px = m.h * H;
  if (m.k === 'dot') { m.w = px / W; return m; }
  c.font = `bold ${m.k === 'tag' ? px * 0.68 : px * 0.86}px ${STAMP_FONT}`; const tw = c.measureText(m.s || '').width;
  const need = (m.k === 'tag' ? tw + px * 0.6 : tw) / W; if (m.k !== 'tag' || m.w < need) m.w = need; return m;
}
function stampChip(it) { // 側欄上的小樣
  const t = esc(it.s || ''), c = it.c;
  if (it.k === 'hl') return `<span class="stc" style="width:44px;height:14px;background:${c};opacity:.75;border-radius:3px"></span>`;
  if (it.k === 'brace') return `<span class="stc" style="color:${c};font-size:22px;line-height:18px;font-weight:300">{</span>`;
  if (it.k === 'line') return `<span class="stc" style="width:40px;height:0;border-top:3px solid ${c}"></span>`;
  if (it.k === 'dot') return `<span class="stc stdot" style="background:${c}">${t}</span>`;
  if (it.k === 'tag') return `<span class="stc sttag" style="background:${c};color:${tagInk(c)}">${t}</span>`;
  return `<span class="stc sttx" style="color:${c}${it.k === 'uline' ? ';text-decoration:underline' : ''}">${t}</span>`;
}
function drawMarks(x, marks, ox, oy, w, h) {
  (marks || []).forEach((m) => {
    x.save(); x.strokeStyle = m.c; x.lineWidth = Math.max(2, Math.max(w, h) * 0.004); x.lineCap = 'round'; x.lineJoin = 'round';
    if (m.t === 'pen' && m.pts.length) { x.beginPath(); m.pts.forEach(([px, py], i) => (i ? x.lineTo(ox + px * w, oy + py * h) : x.moveTo(ox + px * w, oy + py * h))); if (m.pts.length === 1) x.lineTo(ox + m.pts[0][0] * w + 0.5, oy + m.pts[0][1] * h); x.stroke(); }
    if (m.t === 'stamp') drawStamp(x, m, ox, oy, w, h);
    if (m.t === 'circle') { x.beginPath(); x.ellipse(ox + (m.x + m.w / 2) * w, oy + (m.y + m.h / 2) * h, Math.max(1, m.w * w / 2), Math.max(1, m.h * h / 2), 0, 0, Math.PI * 2); x.stroke(); }
    x.restore();
  });
}
/* SOP／範本附圖的標註（畫筆、圈選、遮罩、筆記），存在 n.imgMarks[圖片編號]，原圖不動 */
function noteMark(fid) { for (const n of notes()) { const m = n.imgMarks?.[fid]; if (m) return m; } return null; }
const hasMark = (m) => m && ((m.masks || []).length || (m.marks || []).length);
async function noteImgURL(fid) { // 有標註就回傳畫好標註的版本
  const m = noteMark(fid); if (!hasMark(m)) return await Files.url(fid);
  const cv = await renderDoc({ fileId: fid, masks: m.masks || [], marks: m.marks || [] }, true); return cv.toDataURL('image/jpeg', 0.85);
}
function rotateMarks(m) { // 跟著圖片向右轉 90 度
  if (!m) return m; const box = (r) => ({ ...r, x: 1 - (r.y + r.h), y: r.x, w: r.h, h: r.w });
  return { ...m, masks: (m.masks || []).map(box), marks: (m.marks || []).map((k) => (k.t === 'pen' ? { ...k, pts: k.pts.map(([a, b]) => [1 - b, a]) } : k.t === 'stamp' ? { ...k, x: 1 - (k.y + k.h / 2) - k.w / 2, y: k.x + k.w / 2 - k.h / 2 } : box(k))) };
}
async function drawThumbTo(cv, d, solid) {
  if (isPdf(d)) return;
  const im = await getImg(d.fileId); if (!im || !cv.isConnected) return;
  const x = cv.getContext('2d'); const s = Math.max(cv.width / im.width, cv.height / im.height); const w = im.width * s, h = im.height * s;
  x.drawImage(im, (cv.width - w) / 2, (cv.height - h) / 2, w, h);
  x.fillStyle = solid ? '#111' : 'rgba(17,17,17,.82)';
  d.masks.forEach((m) => x.fillRect((cv.width - w) / 2 + m.x * w, (cv.height - h) / 2 + m.y * h, m.w * w, m.h * h));
  drawMarks(x, d.marks, (cv.width - w) / 2, (cv.height - h) / 2, w, h);
}
function nimgsHTML(n, max) { const all = n.images || [], cut = max && all.length > max ? max - 1 : all.length; return all.length ? `<div class="nimgs">${all.slice(0, cut).map((f) => `<button class="nimg" data-nimg="${f}" data-ntitle="${esc(n.title)}" aria-label="放大看圖"><img data-fid="${f}" alt=""></button>`).join('')}${cut < all.length ? `<button class="nimg more" data-sopopen="${n.id}">＋${all.length - cut}</button>` : ''}</div>` : ''; }
function drawThumbs() {
  const ccl = curClient(); document.querySelectorAll('[data-cthumb]').forEach((cv) => { const d = ccl?.docs[+cv.dataset.cthumb]; if (d) drawThumbTo(cv, d, false); });
  document.querySelectorAll('img[data-fid]').forEach(async (im) => { const u = await noteImgURL(im.dataset.fid).catch(() => null); if (u && im.isConnected) im.src = u; });
  const c = cur(); document.querySelectorAll('[data-thumb]').forEach((cv) => { const d = c?.docs[+cv.dataset.thumb]; if (d) drawThumbTo(cv, d, false); });
  const r = Vault.get('cases', S.refSel); document.querySelectorAll('[data-rthumb]').forEach((cv) => { const d = r?.docs[+cv.dataset.rthumb]; if (d) drawThumbTo(cv, d, true); });
}

/* ---------- 事件 ---------- */
function newCase(cat) { const c = saveCase({ date: localDay(), client: '', title: '', cat, status: '進行中', items: [], causeType: '', causeDetail: '', diag: '', next: '', cost: '', ratio: '', payout: '', notes: '', docs: [], isRef: false }); S.sel = c.id; S.showDetail = true; return c; }
function armConfirm(b, label) { if (b.dataset.confirm) return true; b.dataset.confirm = 1; b.textContent = label; setTimeout(() => { if (b.isConnected) { delete b.dataset.confirm; App.render(); } }, 4000); return false; }
app.addEventListener('click', async (e) => {
  const rview = e.target.closest('[data-refview]'); // 理賠參考：點圖片看遮罩版，可以左右翻
  if (rview) { const rc = Vault.get('cases', S.refSel); if (!rc) return; const list = refDocs(rc).map(([x]) => x); showGallery({ list, i: +rview.dataset.refview, masked: true }); return; }
  const rdoc = e.target.closest('[data-refdoc]'); // 理賠參考：點文件選／取消
  if (rdoc) { const rc = Vault.get('cases', S.refSel); if (!rc) return; const cur = rc.refDocIds || refDocs(rc).map(([x]) => dkey(x)), k = rdoc.dataset.refdoc; rc.refDocIds = cur.includes(k) ? cur.filter((x) => x !== k) : [...cur, k]; saveCase(rc); App.render(); return; }
  const gn = e.target.closest('[data-gonode]'); if (gn) { S.view = 'clients'; S.csel = gn.dataset.gonode; S.ccat = '全部'; S.showDetail = true; S.showIds = false; App.render(); return; }
  const tileImg = e.target.closest('.doc[data-di] canvas, .doc[data-di] .pdftile'); // 點圖片＝看
  if (tileImg) { tileImg.closest('.doc').querySelector('[data-view],[data-cview]')?.click(); return; }
  const kl = e.target.closest('[data-klink]'); if (kl) { const a = Vault.get('calls', kl.dataset.kprev), b = Vault.get('calls', kl.dataset.klink); if (a && b) { if (a.hh && a.hh === b.hh) { b.hh = ''; b.hhOk = false; Vault.put('calls', b); if (!calls().some((x) => x.id !== a.id && x.hh === a.hh)) { a.hh = ''; a.hhOk = false; Vault.put('calls', a); } toast('拆開了'); } else toast('同一戶：' + linkHH(a, b)); App.render(); } return; }
  const kr = e.target.closest('tr[data-ksel]'); if (kr) { S.ktable = false; S.kfromTable = true; S.ksel = kr.dataset.ksel; S.kres = ''; S.showDetail = true; App.render(); window.scrollTo(0, 0); return; }
  const b = e.target.closest('button'); if (!b) return; const d = b.dataset; const c = cur();
  if (d.nav) { S.view = d.nav; S.showDetail = false; S.q = ''; App.render(); return; }
  if (d.cat) { S.cat = d.cat; App.render(); return; }
  if (d.cg) { S.cg = d.cg; S.cat = '全部'; const g = GROUPS.find((x) => x.name === d.cg); if (g) { const has = g.cats.filter((k) => cases().some((c) => c.cat === k)); if (has.length === 1) S.cat = has[0]; } App.render(); return; }
  if (d.sel) { S.folder = '全部'; S.sel = d.sel; S.showDetail = true; App.render(); return; }
  if (d.rsel) { S.refPick = false; }
  if (d.rsel) { S.refSel = d.rsel; S.showDetail = true; App.render(); return; }
  if (d.rline) { S.refLine = d.rline === '全部' ? '' : d.rline; S.refItem = '全部'; App.render(); return; }
  if (d.ritem) { S.refItem = d.ritem; App.render(); return; }
  if (d.rcause) { S.refCause = d.rcause; S.refDet = ''; App.render(); return; }
  if (d.rdet) { S.refDet = S.refDet === d.rdet ? '' : d.rdet; App.render(); return; }
  if (d.sg) { S.sg = d.sg; S.sc = '全部'; if (d.sg === '全部') { S.sk = ''; S.spk = ''; S.smk = ''; S.st = ''; } App.render(); return; } // 點「全部」＝清掉所有篩選
  if (d.sc) { S.sc = d.sc; App.render(); return; }
  if (d.refpick) { const rc = Vault.get('cases', S.refSel); if (!rc) return; if (d.refpick === 'all') rc.refDocIds = rc.docs.filter((x) => !isPdf(x)).map(dkey); else if (d.refpick === 'none') rc.refDocIds = []; else { delete rc.refDocIds; delete rc.refFolders; } saveCase(rc); App.render(); return; }
  if (d.sopopen) { S.sopOpen = S.sopOpen === d.sopopen ? null : d.sopopen; App.render(); return; }
  if (d.sk != null && S.view === 'sop') { S.sk = S.sk === d.sk ? '' : d.sk; App.render(); return; }
  if (d.smk != null && S.view === 'sop') { S.smk = S.smk === d.smk ? '' : d.smk; App.render(); return; }
  if (d.spk != null && S.view === 'sop') { S.spk = S.spk === d.spk ? '' : d.spk; App.render(); return; }
  if (d.st != null && S.view === 'sop') { S.st = S.st === d.st ? '' : d.st; App.render(); return; }
  if (d.newtplfor) { const sop = Vault.get('notes', d.newtplfor); if (!sop) return; const t = Vault.put('notes', { kind: 'tpl', line: lineN(sop), group: 'LINE', channels: ['LINE'], cat: catOf(sop), sopIds: [sop.id], title: sop.title, body: '', images: [] }); editNote(t.id, true); return; }
  if (d.editdisc) { const cl0 = curClient(); if (cl0) openDisclosure(cl0, d.editdisc, ''); return; }
  if (d.deldisc) { const ccl = curClient(); if (!ccl) return; if (armConfirm(b, '再按一次刪除')) { const r = (ccl.disclosures || []).find((x) => x.id === d.deldisc); (r?.files || []).forEach((f) => Files.remove(f.fileId)); ccl.disclosures = (ccl.disclosures || []).filter((x) => x.id !== d.deldisc); Vault.put('clients', ccl); App.render(); } return; }
  if (d.discfile) { const [rid, i] = d.discfile.split(':'); const r = (curClient()?.disclosures || []).find((x) => x.id === rid); const f = r?.files?.[+i]; if (f) openViewer({ ...f, masks: [] }); return; }
  if (d.prole && c) { const [i, r] = d.prole.split(':'); const p = c.people[+i]; p.roles = p.roles.includes(r) ? p.roles.filter((x) => x !== r) : [...p.roles, r]; saveCase(c); App.render(); return; }
  if (d.pdel && c) { if (armConfirm(b, '移除？')) { c.people.splice(+d.pdel, 1); saveCase(c); App.render(); } return; }
  if (d.delrel) { if (armConfirm(b, '移除？')) { Vault.del('rels', d.delrel); App.render(); } return; }
  if (d.rename != null || d.crename != null) {
    const isC = d.crename != null, rec = isC ? curClient() : c, doc = rec?.docs[+(isC ? d.crename : d.rename)]; if (!doc) return;
    const inp = document.createElement('input'); inp.className = 'docnamein'; inp.value = doc.name; inp.setAttribute('aria-label', '文件名稱');
    b.replaceWith(inp); inp.focus(); inp.select();
    let done = false; const finish = (save) => { if (done) return; done = true; const v = inp.value.trim(); if (save && v && v !== doc.name) { doc.name = v; Vault.put(isC ? 'clients' : 'cases', rec); } App.render(); };
    inp.addEventListener('keydown', (ev) => { if (ev.key === 'Enter') { ev.preventDefault(); finish(true); } if (ev.key === 'Escape') finish(false); });
    inp.addEventListener('blur', () => finish(true)); return;
  }
  if (d.docbydate != null) { const rec = d.docbydate === 'c' ? curClient() : cur(); if (rec) { delete rec.docManual; Vault.put(d.docbydate === 'c' ? 'clients' : 'cases', rec); App.render(); toast('已改回依日期排序'); } return; }
  if (d.folder != null) { S.folder = d.folder; App.render(); return; }
  if (d.goclient) { S.ccat = '全部'; S.view = 'clients'; S.csel = d.goclient; S.showDetail = true; S.showIds = false; App.render(); return; }
  if (d.ncflag) { const c2 = cur(); if (!c2) return; c2.nc = c2.nc || {}; c2.nc[d.ncflag] = !c2.nc[d.ncflag]; saveCase(c2); App.render(); return; }
  if (d.klvf) { S.klv = d.klvf; App.render(); return; }
  if (d.kprog) { S.kprog = d.kprog; App.render(); return; }
  if (d.kmark) { S.kmark = d.kmark; App.render(); return; }
  if (d.ksort) { S.ksort = d.ksort; App.render(); return; }
  if (d.kview) { S.kview = d.kview; App.render(); return; }
  if (d.kgev) { S.gev = d.kgev; App.render(); return; }
  if (d.khadd) { const k = curCall(), o = Vault.get('calls', d.khadd); if (k && o) { toast('同一戶：' + linkHH(k, o)); App.render(); } return; }
  if (d.khout) { const o = Vault.get('calls', d.khout); if (o) { const old = o.hh; o.hh = ''; o.hhOk = false; Vault.put('calls', o); const left = calls().filter((x) => x.hh === old); if (left.length === 1) { left[0].hh = ''; left[0].hhOk = false; Vault.put('calls', left[0]); } App.render(); } return; }
  if (d.khhok) { const g = callHouseholds(calls()).get(d.khhok) || []; const nm = g.find((x) => x.hh)?.hh || ((g[0]?.name || '').slice(0, 1) + '家'); g.forEach((k) => { k.hh = nm; k.hhOk = true; Vault.put('calls', k); }); toast(`已確認：${nm}`); App.render(); return; }
  if (d.kar) { S.kar = d.kar; App.render(); return; }
  if (d.klv) { const k = curCall(); if (!k) return; k.level = k.level === d.klv ? '' : d.klv; Vault.put('calls', k); App.render(); return; }
  if (d.kchk) { const k = curCall(); if (!k) return; k[d.kchk] = !k[d.kchk]; Vault.put('calls', k); App.render(); return; }
  if (d.gev) { S.gev = d.gev; S.gedit = false; App.render(); return; }
  if (d.gtype) { S.gtype = d.gtype; App.render(); return; }
  if (d.gpic) { openViewer({ fileId: d.gpic, name: '品項' }, true); return; }
  if (d.gview) { S.gview = d.gview; App.render(); return; }
  if (d.gimg) { giftItemImage(d.gimg); return; }
  if (d.gst) { S.gst = d.gst; App.render(); return; }
  if (d.gtog) { const g = Vault.get('gifts', d.gtog); if (!g) return; g.status = g.status === '已送' ? '預備' : '已送'; g.doneDate = g.status === '已送' ? localDay() : ''; Vault.put('gifts', g); App.render(); toast(g.status === '已送' ? '已標成已送' : '改回預備'); return; }
  if (d.gopen) { openGiftEditor(Vault.get('gifts', d.gopen)); return; }
  if (d.gnew) { const [type, id] = d.gnew.split(':'); const nm = type === 'call' ? Vault.get('calls', id)?.name : Vault.get('clients', id)?.name; openGiftEditor(null, { type, id, name: nm || '' }); return; }
  if (d.gidel) { const ev = Vault.get('gevents', S.gev); if (!ev) return; if (!b.dataset.sure) { b.dataset.sure = 1; b.textContent = '刪？'; return; } ev.items = ev.items.filter((x) => x.id !== d.gidel); Vault.put('gevents', ev); App.render(); return; }
  if (d.ksel) { if (S.ktable) { S.ktable = false; S.kfromTable = true; } S.ksel = d.ksel; S.kres = ''; S.showDetail = true; App.render(); return; }
  if (d.gocall) { S.view = 'calls'; S.ksel = d.gocall; S.kres = ''; S.showDetail = true; App.render(); return; }
  if (d.kst) { S.kst = d.kst; App.render(); return; }
  if (d.kres) { const note = $('kNote')?.value || '', next = $('kNext')?.value || ''; S.kres = S.kres === d.kres ? '' : d.kres; App.render(); $('kNote').value = note; $('kNext').value = next; return; }
  if (d.kdel) { const k = curCall(); if (k && armConfirm(b, '刪？')) { k.logs = (k.logs || []).filter((x) => x.id !== d.kdel); Vault.put('calls', k); App.render(); } return; }
  if (d.ccat) { S.ccat = d.ccat; App.render(); return; }
  if (d.csel) { S.ccat = '全部'; S.csel = d.csel; S.showDetail = true; S.showIds = false; App.render(); return; }
  if (d.reveal) { const el = $(d.reveal); if (el) { const on = el.classList.toggle('shown'); b.textContent = on ? '隱藏' : '顯示'; } return; }
  const ccl = curClient();
  if (d.cview && ccl) { openGallery(b, ccl.docs, 'cview'); return; }
  if (d.cpdfimg != null && ccl) { pdfToImages(ccl, 'clients', +d.cpdfimg, b); return; }
  if (d.crot != null && ccl) { rotateDoc(ccl, 'clients', +d.crot, b); return; }
  if (d.cmask && ccl) { openMask(ccl, +d.cmask, 'clients'); return; }
  if (d.cdeldoc && ccl) { if (armConfirm(b, '確定？')) { const [doc] = ccl.docs.splice(+d.cdeldoc, 1); Vault.put('clients', ccl); if (doc) { Files.remove(doc.fileId, { coll: 'clients', recId: ccl.id, doc }); toast('已移到「最近刪除」，30 天內可以救回'); } App.render(); } return; }
  if (d.addstep != null && c) {
    const e = { id: uid(), date: localDay(), title: d.addstep, note: '', at: new Date().toISOString() };
    c.timeline = [...(c.timeline || []), e]; saveCase(c); App.render();
    setTimeout(() => (d.addstep ? $('tln_' + e.id) : $('tlt_' + e.id))?.focus(), 30); return;
  }
  if (d.favstep && c) { setSteps(c.cat, [...getSteps(c.cat), d.favstep]); App.render(); toast(`「${d.favstep}」已加到常用`); return; }
  if (d.setst && c) { if (c.status !== d.setst) { c.status = d.setst; saveCase(c); App.render(); } return; }
  if (d.tldel && c) { if (armConfirm(b, '刪？')) { c.timeline = (c.timeline || []).filter((x) => x.id !== d.tldel); saveCase(c); App.render(); } return; }
  if (d.unlink && c) { const o = Vault.get('cases', d.unlink); c.links = (c.links || []).filter((x) => x !== d.unlink); saveCase(c); if (o) { o.links = (o.links || []).filter((x) => x !== c.id); saveCase(o); } App.render(); toast('已取消連結'); return; }
  if (d.gocase) { S.folder = '全部'; S.view = 'cases'; S.sel = d.gocase; S.showDetail = true; App.render(); return; }
  if (d.edititems) { editItems(d.edititems); return; }
  if (d.lf) { const [k, v] = d.lf.split(':'); S[k] = v === '全部' ? '' : v; if (k === 'nline') { [['kinds', 'sk'], ['tools', 'st'], ['pkinds', 'spk'], ['mkinds', 'smk']].forEach(([t, f]) => { if (!tagShown(t, S.nline)) S[f] = ''; }); } App.render(); return; }
  if (d.line && c) { c.line = d.line; saveCase(c); App.render(); return; }
  if (d.item && c) { const i = c.items.indexOf(d.item); i < 0 ? c.items.push(d.item) : c.items.splice(i, 1); saveCase(c); App.render(); return; }
  if (d.ctype && c) { c.causeType = c.causeType === d.ctype ? '' : d.ctype; c.causeDetail = ''; saveCase(c); App.render(); return; }
  if (d.editcause) { editCause(d.editcause); return; }
  if (d.cdet && c) { c.causeDetail = c.causeDetail === d.cdet ? '' : d.cdet; saveCase(c); App.render(); return; }
  if (d.pdfimg != null && c) { pdfToImages(c, 'cases', +d.pdfimg, b); return; }
  if (d.rot != null && c) { rotateDoc(c, 'cases', +d.rot, b); return; }
  if (d.mask && c) { openMask(c, +d.mask); return; }
  if (d.view && c) { openGallery(b, c.docs, 'view'); return; }
  if (d.deldoc && c) { if (armConfirm(b, '確定？')) { const [doc] = c.docs.splice(+d.deldoc, 1); saveCase(c); if (doc) { Files.remove(doc.fileId, { coll: 'cases', recId: c.id, doc }); toast('已移到「最近刪除」，30 天內可以救回'); } App.render(); } return; }
  if (d.copy && Vault.get('notes', d.copy)) { e.preventDefault(); openSend(Vault.get('notes', d.copy), d.forcase ? c : null); return; }
  if (d.copy) { e.preventDefault(); const n = Vault.get('notes', d.copy); const txt = fillTpl(n.body, d.forcase ? c : null); try { await navigator.clipboard.writeText(txt); toast('已複製' + (d.forcase && c?.client ? `（已帶入 ${c.client}）` : '')); } catch (err) { toast('無法自動複製，請長按文字選取'); } return; }
  if (d.tg) { S.tplGroup = S.tplGroup === d.tg ? '全部' : d.tg; App.render(); return; }
  if (d.tgg) { S.tg = d.tgg; S.tplCat = '全部'; if (d.tgg === '全部') S.tplGroup = '全部'; App.render(); return; }
  if (d.tc) { S.tplCat = d.tc; App.render(); return; }
  if (d.nimg) { openViewer({ fileId: d.nimg, name: d.ntitle || '附圖' }); return; }
  if (d.delnote) { if (armConfirm(b, '確定刪除？')) { (Vault.get('notes', d.delnote)?.images || []).forEach((f) => Files.remove(f)); Vault.del('notes', d.delnote); App.render(); } return; }
  if (d.editnote) { editNote(d.editnote); return; }
  if (d.addnote) {
    const t = $('nTitle').value.trim(), bd = $('nBody').value.trim(), imgs = [...($('nImgs')?.files || [])];
    if (!t && !bd && !imgs.length) { toast('先填標題、內容或附圖'); return; }
    const n = Vault.put('notes', d.addnote === 'sop' ? { kind: 'sop', v2: true, cat: $('nCat').value, kinds: S.sk ? [S.sk] : [], tools: S.st ? [S.st] : [], title: t || '未命名', body: bd, images: [] } : { kind: d.addnote, group: $('nGroup').value, cat: $('nCat')?.value || '', title: t || '未命名', body: bd, images: [] });
    if (imgs.length) { b.disabled = true; await addNoteImages(n, imgs); }
    App.render(); toast('已新增'); return;
  }
  switch (d.act) {
    case 'newclient': { const n = Vault.put('clients', { name: '', docs: [] }); S.csel = n.id; S.q = ''; S.showDetail = true; S.showIds = false; S.view = 'clients'; App.render(); setTimeout(() => $('cf_name')?.focus(), 50); break; }
    case 'cscan': scanForClient(ccl); break;
    case 'toggleids': S.showIds = !S.showIds; App.render(); break;
    case 'cpdf': openPrintPick(ccl.docs, (ccl.name || '客戶') + '-證件'); break;
    case 'cshare': openShare(ccl, (ccl.name || '客戶') + '-證件'); break;
    case 'ccase': { const c = newCase(S.ccat !== '全部' ? S.ccat : '理賠'); c.client = ccl.name; c.clientId = ccl.id; saveCase(c); S.view = 'cases'; App.render(); break; }
    case 'delclient': if (armConfirm(b, '再按一次確定刪除（證件照片也會刪掉）')) { (ccl.docs || []).forEach((doc) => Files.remove(doc.fileId)); cases().filter((c) => c.clientId === ccl.id).forEach((c) => { c.clientId = ''; saveCase(c); }); Vault.all('rels').filter((r) => r.a === ccl.id || r.b === ccl.id).forEach((r) => Vault.del('rels', r.id)); Vault.del('clients', ccl.id); S.csel = null; S.showDetail = false; App.render(); } break;
    case 'newcall': { const k = Vault.put('calls', { name: '', phone: '', source: '', status: '待撥', nextDate: '', logs: [] }); S.ksel = k.id; S.kres = ''; S.showDetail = true; S.view = 'calls'; App.render(); setTimeout(() => $('kf_name')?.focus(), 50); break; }
    case 'importcalls': openCallImport(); break;
    case 'newgevent': openNewGiftEvent(); break;
    case 'nctpl': openNcTplEditor(); break;
    case 'kedit': S.kedit = !S.kedit; App.render(); break;
    case 'railtog': S.railMini = !S.railMini; try { localStorage.setItem('crm-railmini', S.railMini ? '1' : '0'); } catch (e) { } App.render(); break;
    case 'korder': openCallOrderImport(); break;
    case 'kphoto': openCallPhoto(null); break;
    case 'kgiftgo': S.view = 'calls'; S.ktable = true; S.kview = 'gift'; S.showDetail = false; S.q = ''; App.render(); break;
    case 'kphotoin': openCallPhoto(curCall()); break;
    case 'ktable': S.ktable = true; S.kfromTable = false; S.showDetail = false; App.render(); break;
    case 'klist': S.ktable = false; App.render(); break;
    case 'cflow': openCallFlow(); break;
    case 'gaddwho': openGiftEditor(null, null); break;
    case 'gbulki': openGiftBulkItems(); break;
    case 'gbulkw': openGiftBulkWho(); break;
    case 'gpickw': openGiftPickWho(); break;
    case 'geditems': S.gedit = !S.gedit; App.render(); break;
    case 'gaddi': { const ev = Vault.get('gevents', S.gev); if (!ev) break; ev.items = [...(ev.items || []), { id: uid(), name: '', code: '', qty: 0 }]; Vault.put('gevents', ev); App.render(); break; }
    case 'gevdel': { const ev = Vault.get('gevents', S.gev); if (!ev) break; if (!armConfirm(b, '連名單一起刪除，再按一次')) break; Vault.all('gifts').filter((g) => g.eventId === ev.id).forEach((g) => Vault.del('gifts', g.id)); Vault.del('gevents', ev.id); S.gev = null; App.render(); break; }
    case 'samplescript': { const n = Vault.put('notes', { kind: 'script', group: '電訪', title: '服務件電訪', body: SAMPLE_SCRIPT, images: [] }); S.script = n.id; editNote(n.id); break; }
    case 'newscript': { const n = Vault.put('notes', { kind: 'script', group: '電訪', title: '', body: '', images: [] }); S.script = n.id; editNote(n.id, true); break; }
    case 'editscript': if (S.script) editNote(S.script); break;
    case 'logcall': {
      const k = curCall(); if (!k) break; const note = $('kNote').value.trim(), next = $('kNext').value;
      if (!S.kres && !note) { toast('先選結果或寫一點內容'); break; }
      const res = RESULTS.find(([r]) => r === S.kres);
      k.logs = [...(k.logs || []), { id: uid(), date: localDay(), at: new Date().toISOString(), result: S.kres || '通話', note }];
      if (res) k.status = res[1]; if (next) k.nextDate = next; else if (res && res[1] === '結案') k.nextDate = '';
      Vault.put('calls', k); S.kres = ''; App.render(); toast('已記錄'); break;
    }
    case 'tocliient': { const k = curCall(); if (!k) break; let cl = clients().find((x) => x.name === k.name && k.name); if (!cl) cl = Vault.put('clients', { name: k.name, phone: k.phone, note: [k.source, k.note].filter(Boolean).join('\n'), docs: [] }); k.clientId = cl.id; Vault.put('calls', k); App.render(); toast(`已轉成客戶「${cl.name || '未填姓名'}」`); break; }
    case 'delcall': if (armConfirm(b, '再按一次確定刪除')) { Vault.del('calls', S.ksel); S.ksel = null; S.showDetail = false; App.render(); } break;
    case 'newtpl': { const ch = S.tplGroup !== '全部' ? S.tplGroup : 'LINE', tg = GROUPS.find((x) => x.name === S.tg); const n = Vault.put('notes', { kind: 'tpl', line: S.nline || '通用', group: ch, channels: [ch], cat: S.tplCat !== '全部' ? S.tplCat : tg?.cats[0] || '', title: '', body: '', images: [] }); editNote(n.id, true); break; }
    case 'refpickmode': S.refPick = !S.refPick; App.render(); break;
    case 'pickref': S.refPick = true; S.view = 'refs'; S.refSel = c.id; S.refItem = '全部'; S.refCause = '全部'; S.refDet = ''; S.showDetail = true; App.render(); break;
    case 'newsop': { const g = GROUPS.find((x) => x.name === S.sg); const n = Vault.put('notes', { kind: 'sop', v2: true, line: S.nline || '通用', cat: S.sc !== '全部' ? S.sc : g?.cats[0] || '', kinds: S.sk ? [S.sk] : [], pkinds: S.spk ? [S.spk] : [], mkinds: S.smk ? [S.smk] : [], tools: S.st ? [S.st] : [], title: '', body: '', images: [] }); editNote(n.id, true); break; }
    case 'sopfor': S.view = 'sop'; S.sg = groupOf(c.cat) || '全部'; S.sc = groupOf(c.cat) ? c.cat : '全部'; S.sk = ''; S.spk = ''; S.smk = ''; S.st = ''; S.q = ''; S.showDetail = false; App.render(); break;
    case 'tplfor': S.view = 'tpl'; S.tg = groupOf(c.cat) || '全部'; S.tplCat = groupOf(c.cat) ? c.cat : '全部'; S.tplGroup = '全部'; S.q = ''; S.showDetail = false; App.render(); break;
    case 'adddisc': if (ccl) openDisclosure(ccl, null, ''); break;
    case 'togglehealth': S.showHealth = !S.showHealth; App.render(); break;
    case 'casedisc': { const ppl = casePeople(c); if (ppl.length > 1) openWhoPicker(c, ppl); else { const cl2 = Vault.get('clients', c.clientId); if (cl2) openDisclosure(cl2, null, c.id); } break; }
    case 'addperson': openPersonPicker(c); break;
    case 'crmimport': if (ccl) openCrmImport(ccl); break;
    case 'crmimportnew': openCrmImport(null); break;
    case 'addrel': if (ccl) openRelPicker(ccl); break;
    case 'relgraph': openRelGraph(); break;
    case 'backups': openBackups(); break;
    case 'trash': openTrash(); break;
    case 'nctest': openNcTest(); break;
    case 'exportnotes': {
      const names = clients().map((x) => x.name).filter((n) => n && n.length >= 2).sort((a2, b2) => b2.length - a2.length);
      const clean = (t) => { let o = String(t || ''); names.forEach((n) => { o = o.split(n).join('○○'); }); return o.replace(/[A-Z][12]\d{8}/g, '［身分證］').replace(/09\d{2}[-\s]?\d{3}[-\s]?\d{3}/g, '［電話］').replace(/[\w.+-]+@[\w.-]+/g, '［Email］'); };
      const list = cases().filter((x) => ['新契約', '建議書'].includes(x.cat) && ((x.notes || '').trim() || ncLogText(x) || (x.timeline || []).some((e) => e.note))).sort(byDate);
      const txt = list.map((x, i) => `#${i + 1}【${x.cat}】${clean(x.title)}\n${ncLogText(x) ? clean(ncLogText(x)) + '\n' : ''}${clean(x.notes)}${(x.timeline || []).filter((e) => e.note).map((e) => `\n・${e.title}：${clean(e.note)}`).join('')}`).join('\n\n---\n\n');
      if (!txt) { toast('還沒有新契約或建議書的備註'); break; }
      try { await navigator.clipboard.writeText(txt); toast(`已複製 ${list.length} 件的備註，客戶姓名已換成 ○○`); } catch (err) { toast('無法複製，請在電腦上試試'); }
      break;
    }
    case 'savenow': {
      const a = document.activeElement; if (a && a !== document.body && a.blur) a.blur(); // 先讓正在打的欄位存進去
      b.textContent = '儲存中…'; b.disabled = true;
      setTimeout(async () => { try { await Vault.save(true); } catch (err) { } const ok = Vault.state === 'synced'; toast(ok ? '✓ 已儲存' : '已存在這台裝置，連上網路會自動上傳'); App.render(); }, 60);
      break;
    }
    case 'editcats': openCatEditor(); break;
    case 'edittags-kinds': openTagEditor('kinds'); break;
    case 'edittags-tools': openTagEditor('tools'); break;
    case 'edittags-pkinds': openTagEditor('pkinds'); break;
    case 'edittags-mkinds': openTagEditor('mkinds'); break;
    case 'pasteimg': pasteFromClipboard(); break;
    case 'editsteps': openStepEditor(c.cat); break;
    case 'linkcase': openLinkPicker(c); break;
    case 'convert': {
      const nc = saveCase({ date: localDay(), client: c.client, clientId: c.clientId || '', title: c.title, cat: '新契約', status: '進行中', items: [], causeType: '', causeDetail: '', diag: '', next: '送件', cost: '', ratio: '', payout: '', notes: '', docs: [], isRef: false,
        people: JSON.parse(JSON.stringify(c.people || [])), links: [c.id], fromProposal: c.id, nc: c.nc ? JSON.parse(JSON.stringify(c.nc)) : undefined,
        timeline: [{ id: uid(), date: localDay(), title: '簽約', note: '', at: new Date().toISOString() }] });
      c.links = [...new Set([...(c.links || []), nc.id])]; c.outcome = '成交'; c.status = '已結案'; saveCase(c);
      S.sel = nc.id; S.folder = '全部'; if (S.cat === '建議書') { S.cat = '新契約'; S.cg = groupOf('新契約') || '全部'; } App.render(); toast('已建立新契約，記得記錄告知資料'); break;
    }
    case 'newfolder': { const name = (prompt('資料夾名稱（例如：體檢報告、基本資料、收據）') || '').trim(); if (!name) break; c.folders = c.folders || []; if (!c.folders.includes(name)) c.folders.push(name); saveCase(c); S.folder = name; App.render(); break; }
    case 'renfolder': { const old = S.folder; const name = (prompt('新的資料夾名稱', old) || '').trim(); if (!name || name === old) break; c.folders = c.folders.map((f) => (f === old ? name : f)); c.docs.forEach((doc) => { if (doc.folder === old) doc.folder = name; }); saveCase(c); S.folder = name; App.render(); break; }
    case 'delfolder': if (armConfirm(b, '確定刪除？裡面的文件會移到未分類')) { const old = S.folder; c.folders = c.folders.filter((f) => f !== old); c.docs.forEach((doc) => { if (doc.folder === old) doc.folder = ''; }); saveCase(c); S.folder = '全部'; App.render(); } break;
    case 'scan': { if (S.view === 'clients' && ccl) { scanForClient(ccl); break; } let t = c; if (!t || S.view !== 'cases') { t = newCase('理賠'); S.view = 'cases'; } openScanner(t, { folder: curFolder(t) }); break; }
    case 'newcase': newCase(S.cat !== '全部' ? S.cat : GROUPS.find((x) => x.name === S.cg)?.cats[0] || '理賠'); App.render(); break;
    case 'back': S.showDetail = false; App.render(); break;
    case 'ref': c.isRef = !c.isRef; saveCase(c); App.render(); toast(c.isRef ? '已加入理賠參考（不顯示客戶姓名）' : '已從理賠參考移除'); break;
    case 'delcase': if (armConfirm(b, '再按一次確定刪除')) { c.docs.forEach((doc) => Files.remove(doc.fileId)); (c.links || []).forEach((id) => { const o = Vault.get('cases', id); if (o) { o.links = (o.links || []).filter((x) => x !== c.id); saveCase(o); } }); Vault.del('cases', c.id); S.sel = null; S.showDetail = false; App.render(); } break;
    case 'pdf': openPrintPick(c.docs, (c.client || '案件') + '-' + (c.title || '文件')); break;
    case 'share': openShare(c); break;
    case 'rshare': { const rc = Vault.get('cases', S.refSel); if (rc) openShare(rc, null, refDocs(rc).map(([d]) => d)); break; }
    case 'gocase': S.view = 'cases'; S.sel = S.refSel; S.showDetail = true; App.render(); break;
    case 'addphone': { const n = Vault.put('notes', { kind: 'quick', group: '電話', title: '', body: '' }); editNote(n.id, true); break; }
    case 'import': openImport(); break;
    case 'lock': lock(); break;
    case 'signout': await Vault.save(); Drive.signOut(); break;
  }
});
function linkClient(c) { // 案件填了客戶名字：找到同名客戶就連起來，沒有就自動建立
  const name = (c.client || '').trim(); c.client = name;
  if (!name) { c.clientId = ''; return null; }
  let cl = clients().find((x) => x.name === name), made = false;
  if (!cl) { cl = Vault.put('clients', { name, docs: [] }); made = true; }
  c.clientId = cl.id; return made ? cl : null;
}
app.addEventListener('change', () => { App._defer = true; setTimeout(() => (App._defer = false), 0); }, true);
app.addEventListener('change', (e) => {
  if (e.target.id === 'kScript') { S.script = e.target.value; App.render(); return; }
  const kgq = e.target.dataset?.kgq; if (kgq) { const [cid, iid] = kgq.split(':'), n = Math.max(0, parseInt(e.target.value, 10) || 0), g = kgRec(cid); if (!g) return; g.items = { ...(g.items || {}) }; if (n) g.items[iid] = n; else delete g.items[iid]; kgTidy(g); if (n) { const k = Vault.get('calls', cid); if (k?.giftSkip) { k.giftSkip = false; Vault.put('calls', k); } } App.render(); return; }
  const kgm = e.target.dataset?.kgm; if (kgm) { const g = kgRec(kgm); if (!g) return; g.method = e.target.value; kgTidy(g); App.render(); return; }
  const kgd = e.target.dataset?.kgd; if (kgd) { const g = kgRec(kgd); if (!g) return; g.status = e.target.checked ? '已送' : '預備'; g.doneDate = e.target.checked ? localDay() : ''; kgTidy(g); App.render(); return; }
  const kte = e.target.dataset?.kte; if (kte) { const k = Vault.get('calls', kte); if (!k) return; const f = e.target.dataset.kfield; k[f] = e.target.type === 'checkbox' ? e.target.checked : f === 'name' || f === 'hh' ? e.target.value.trim() : e.target.value; if (f === 'hh') k.hhOk = !!k.hh; Vault.put('calls', k); if (e.target.type === 'checkbox' || e.target.tagName === 'SELECT' || f === 'visitDate') App.render(); return; }
  const kf = e.target.dataset?.kf; if (kf) { const k = curCall(); if (!k) return; k[kf] = kf === 'name' ? e.target.value.trim() : e.target.value; Vault.put('calls', k); if (['status', 'nextDate', 'visitDate', 'otherWhy'].includes(kf)) App.render(); return; } // 打字的欄位不重畫，鍵盤才不會收起來
  const cf = e.target.dataset?.cf; const cl = curClient();
  if (cf && cl) {
    cl[cf] = e.target.value; if (cf === 'name') cl.name = cl.name.trim();
    Vault.put('clients', cl);
    if (cf === 'name') { cases().filter((c) => c.clientId === cl.id).forEach((c) => { c.client = cl.name; saveCase(c); }); App.render(); }
    return;
  }
});
app.addEventListener('change', (e) => {
  const tk = ['tldate', 'tltitle', 'tlnote'].find((k) => e.target.dataset?.[k] != null);
  if (tk) { const c = cur(); const it = c?.timeline?.find((x) => x.id === e.target.dataset[tk]); if (!it) return; it[{ tldate: 'date', tltitle: 'title', tlnote: 'note' }[tk]] = e.target.value; saveCase(c); if (tk !== 'tlnote') App.render(); return; }
  if (e.target.dataset?.pprod != null) { const c = cur(); if (!c) return; c.people[+e.target.dataset.pprod].product = e.target.value; saveCase(c); return; }
  if (e.target.dataset?.docdate != null) { const c = cur(); if (!c) return; c.docs[+e.target.dataset.docdate].date = e.target.value; saveCase(c); App.render(); return; }
  if (e.target.dataset?.cdocdate != null) { const cl = curClient(); if (!cl) return; cl.docs[+e.target.dataset.cdocdate].date = e.target.value; Vault.put('clients', cl); App.render(); return; }
  if (e.target.dataset?.movedoc != null) { const c = cur(); if (!c) return; c.docs[+e.target.dataset.movedoc].folder = e.target.value; saveCase(c); App.render(); return; }
  if (e.target.id === 'docUp' || e.target.id === 'docUpC') { const fs = [...e.target.files]; e.target.value = ''; takeImages(fs); }
});
app.addEventListener('change', (e) => { if (e.target.dataset?.catgroup != null) { const c = cur(); const g = GROUPS.find((x) => x.name === e.target.value); if (c && g?.cats.length) { c.cat = g.cats[0]; saveCase(c); App.render(); setTimeout(() => $('f_cat')?.focus(), 30); } return; } if (e.target.dataset?.f === 'client') { const c = cur(); if (!c) return; c.client = e.target.value; const made = linkClient(c); saveCase(c); App.render(); if (made) toast(`已建立客戶「${made.name}」`); return; } if (e.target.id === 'nImgs') { const k = e.target.files.length; $('nImgsN').textContent = k ? `已選 ${k} 張` : ''; return; } const f = e.target.dataset?.f; const c = cur(); if (!f || !c) return; c[f] = e.target.value; saveCase(c); if (['cat', 'status', 'cost', 'ratio', 'client', 'title', 'date'].includes(f)) App.render(); });
app.addEventListener('input', (e) => {
  const ds = e.target.dataset || {}; if (ds.f || ds.cf || ds.kf || ds.tltitle || ds.tlnote || ds.pprod) savePill('editing');
  if (e.target.classList?.contains('grow')) fitNote(e.target);
  if (e.target.classList?.contains('tlnote')) { fitNote(e.target); return; }
  if (e.target.id === 'q') { S.q = e.target.value; if (!e.isComposing) App.render(); } // 注音選字中不要重畫，否則會打斷輸入法
  if (e.target.id === 'qq') {
    const q = e.target.value.trim(); const box = $('qres'); if (!q) { box.innerHTML = ''; return; }
    const cls = clients().filter((x) => [x.name, x.phone, x.line, x.note].join(' ').includes(q)).slice(0, 6);
    const ks = calls().filter((x) => [x.name, x.phone, x.source].join(' ').includes(q)).slice(0, 6);
    const ns = notes().filter((n) => (n.title + n.body).includes(q)).slice(0, 6), cs = cases().filter((c) => [c.client, c.title, c.notes, c.next].join(' ').includes(q)).slice(0, 6);
    box.innerHTML = `<div class="cards" style="margin-bottom:6px">${ks.map((x) => `<button class="item" data-gocall="${x.id}"><span class="t">${esc(x.name || '未填姓名')}</span><span class="s">電訪 · ${esc(x.status)}</span></button>`).join('')}${cls.map((x) => `<button class="item" data-goclient="${x.id}"><span class="t">${esc(x.name || '未填姓名')}</span><span class="s">客戶</span></button>`).join('')}${ns.map((n) => `<div class="card"><h3><span>${esc(n.title)}</span><span class="sub">${esc(n.group)}</span></h3><pre>${esc(n.body)}</pre></div>`).join('')}${cs.map((c) => `<button class="item${itemCls(c.status)}" data-gocase="${c.id}"><span class="t">${esc(c.client)} · ${esc(c.title)}</span><span class="s">${esc(caseSummary(c))}</span></button>`).join('')}</div>${!ks.length && !cls.length && !ns.length && !cs.length ? '<div class="sub">找不到。</div>' : ''}`;
  }
});
app.addEventListener('change', (e) => { // 新契約清單：打勾、過程紀錄
  const ds = e.target.dataset || {}; if (!ds.ncc && !ds.ncv) return; const c = cur(); if (!c) return; c.nc = c.nc || {};
  if (ds.ncc) { c.nc.chk = { ...(c.nc.chk || {}), [ds.ncc]: e.target.checked }; saveCase(c); App.render(); }
  else { c.nc.val = { ...(c.nc.val || {}), [ds.ncv]: e.target.value }; saveCase(c); }
});
app.addEventListener('toggle', (e) => { if (e.target.classList?.contains('ginv')) S.ginvOpen = e.target.open; }, true);
app.addEventListener('change', (e) => { // 送禮表格：直接改數量、分類、送法、已送、備註
  if (e.target.id === 'gImgs') { const fs = [...e.target.files]; e.target.value = ''; if (fs.length) giftBulkImages(fs); return; }
  const ds = e.target.dataset || {};
  if (ds.gq2) { const [gid, iid] = ds.gq2.split(':'); const g = Vault.get('gifts', gid); if (!g) return; const n = Math.max(0, parseInt(e.target.value, 10) || 0); g.items = { ...(g.items || {}) }; if (n) g.items[iid] = n; else delete g.items[iid]; Vault.put('gifts', g); App.render(); return; }
  if (ds.gg && e.target.value === '__edit') { App.render(); openGiftGroupEditor(); return; }
  if (ds.gg) { const [gid, f] = ds.gg.split(':'); const g = Vault.get('gifts', gid); if (!g) return; g[f] = e.target.value.trim(); Vault.put('gifts', g); if (f === 'group') App.render(); return; }
  if (ds.gwho) { // 改類型：名字跟客戶或服務件一樣就自動連過去，不一樣就只標類型
    const g = Vault.get('gifts', ds.gwho); if (!g) return; const t = e.target.value, name = giftWho(g);
    const hit = t === 'client' ? clients().find((x) => x.name === name) : t === 'call' ? calls().find((x) => x.name === name) : null;
    g.who = { type: t, id: hit?.id || '', name }; Vault.put('gifts', g); App.render();
    toast(hit ? `已連到${t === 'client' ? '客戶' : '服務件'}「${name}」` : t === 'other' ? '已改成親友' : `已標成${t === 'client' ? '客戶' : '服務件'}（找不到同名的，先只標類型）`); return;
  }
  if (ds.gdone) { const g = Vault.get('gifts', ds.gdone); if (!g) return; g.status = e.target.checked ? '已送' : '預備'; g.doneDate = e.target.checked ? localDay() : ''; Vault.put('gifts', g); App.render(); }
});
app.addEventListener('change', (e) => { // 送禮：品項和活動名稱
  const gi = e.target.dataset?.gitem; if (gi) { const ev = Vault.get('gevents', S.gev); const it = ev?.items.find((x) => x.id === gi); if (!it) return; const f = e.target.dataset.gf; it[f] = f === 'qty' ? Math.max(0, parseInt(e.target.value, 10) || 0) : e.target.value.trim(); Vault.put('gevents', ev); App.render(); return; }
  const gef = e.target.dataset?.gevf; if (gef) { const ev = Vault.get('gevents', S.gev); if (!ev) return; ev[gef] = e.target.value.trim(); Vault.put('gevents', ev); App.render(); }
});
app.addEventListener('compositionend', (e) => { if (e.target.id === 'q') { S.q = e.target.value; App.render(); } });
const ov = $('ovRoot');
let NE = null; // 正在編輯的筆記（拖曳／貼上圖片會加到這裡）
async function shrink(file, max = 2400, q = 0.85) { // 截圖縮到合理大小，文字仍清楚
  const bmp = await createImageBitmap(file).catch(() => null); if (!bmp) return null;
  const r = Math.min(1, max / Math.max(bmp.width, bmp.height));
  const c = document.createElement('canvas'); c.width = Math.round(bmp.width * r); c.height = Math.round(bmp.height * r);
  const x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height); x.drawImage(bmp, 0, 0, c.width, c.height);
  return await new Promise((ok) => c.toBlob((b) => ok(b), 'image/jpeg', q));
}
async function addNoteImages(n, files) {
  const imgs = [...files].filter((f) => /^image\//.test(f.type)); let added = 0;
  for (const f of imgs) { const blob = await shrink(f); if (!blob) continue; const id = await Files.add(blob, n.title || '附圖'); n.images = [...(n.images || []), id]; added++; }
  if (added) Vault.put('notes', n);
  if (imgs.length > added) toast('有圖片打不開，換一張試試');
  return added;
}
const tplRowHTML = (t) => `<div class="etpl" data-tid="${t?.id || ''}"><div class="etplhead"><span class="tags etch">${[...new Set([...TPLGROUPS, ...(t ? chOf(t) : [])])].map((g) => `<button type="button" class="tag sm${(t ? chOf(t) : ['LINE']).includes(g) ? ' on' : ''}" data-t="${esc(g)}">${esc(g)}</button>`).join('')}</span><input class="ett" value="${esc(t?.title || '')}" placeholder="範本標題（可不填，用 SOP 標題）" aria-label="範本標題"><button type="button" class="tldel etx" aria-label="取消連結" title="取消連結（範本會留在範本頁）">✕</button></div><textarea class="etb" rows="5" placeholder="{客戶}您好，…" aria-label="範本內容">${esc(t?.body || '')}</textarea></div>`;
function editNote(id, isNew) {
  const n = Vault.get('notes', id);
  const alsoHTML = () => `<div class="f" style="font-size:13px;color:var(--muted)">也放在（可不選）<div class="tags" id="eAlso">${(n.alsoCats || []).map((k) => `<button type="button" class="tag sm on" data-also="${esc(k)}">${esc(k)} ✕</button>`).join('')}<select id="eAlsoSel" class="linktpl" style="padding:3px 8px;font-size:13px"><option value="">＋ 也放在其他分類…</option>${catOptions('', '')}</select></div></div>`; const withImgs = n.kind !== 'quick';
  NE = withImgs ? n : null;
  const imgsBlock = () => `<div class="sub">附圖 · 可以把圖片拖進來，或按 Ctrl+V 貼上截圖</div><div class="nimgs edit">${(n.images || []).map((f, i) => `<div class="nimgwrap"><button class="nimg" data-vimg="${f}" aria-label="放大看圖"><img data-fid="${f}" alt=""></button><button class="nimgmk" data-mkimg="${i}" aria-label="標註重點" title="標註重點">✏️</button>${n.imgMarks?.[f]?.note || hasMark(n.imgMarks?.[f]) ? '<span class="nimgdot" title="有標註"></span>' : ''}<button class="nimgrot" data-rotimg="${i}" aria-label="向右轉 90 度" title="轉方向">↻</button><button class="nimgdel" data-delimg="${i}" aria-label="移除這張圖">✕</button></div>`).join('')}<button class="adddoc nimgadd" id="eCam">${ICON.scan}拍照／掃描</button><button class="adddoc nimgadd" id="eAddImg">＋ 選圖片</button></div><input type="file" id="eImgs" accept="image/*" multiple>`;
  ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>${n.kind === 'quick' ? '電話' : '編輯'}</h2><button class="btn" id="eCancel">取消</button><button class="btn pri" id="eSave">儲存</button></div><div class="ovbody" style="flex-direction:column;max-width:720px;margin:0 auto;width:100%">${n.kind === 'sop' || n.kind === 'tpl' ? `<div class="f">業務別${lineSeg(lineN(n), 'eline', NLINES)}</div>` : ''}${n.kind === 'sop' ? `<label class="f">分類<select id="eC">${catOptions(catOf(n), '通用（不屬於特定分類）')}</select></label>${alsoHTML()}${Object.keys(TAGS).map((key) => `<div class="f" data-tagrow="${key}" style="font-size:13px;color:var(--muted)${tagShown(key, lineN(n)) ? '' : ';display:none'}">${TAGS[key].label}<div class="tags" data-etag="${key}">${[...new Set([...tagList(key), ...TAGS[key].of(n)])].map((t) => `<button type="button" class="tag sm${TAGS[key].of(n).includes(t) ? ' on' : ''}" data-t="${esc(t)}">${esc(t)}</button>`).join('')}</div></div>`).join('')}` : ''}${n.kind === 'tpl' ? `<label class="f">分類<select id="eC">${catOptions(n.cat, '通用')}</select></label>${alsoHTML()}<div class="f" style="font-size:13px;color:var(--muted)">用在哪裡（可複選）<div class="tags" data-etag="channels">${[...new Set([...TPLGROUPS, ...chOf(n)])].map((g) => `<button type="button" class="tag sm${chOf(n).includes(g) ? ' on' : ''}" data-t="${esc(g)}">${esc(g)}</button>`).join('')}</div></div><div class="f">掛在哪些 SOP 底下<div class="tags" id="eSops">${sopsOf(n).map((id) => Vault.get('notes', id)).filter(Boolean).map((x) => `<button type="button" class="tag sm on" data-sopid="${x.id}">${esc(x.title)} ✕</button>`).join('')}</div><select id="eS"><option value="">＋ 連結到 SOP…</option>${notes().filter((x) => x.kind === 'sop').sort((a, b) => (catOf(b) === n.cat) - (catOf(a) === n.cat) || (a.title || '').localeCompare(b.title || '')).map((x) => `<option value="${x.id}">${esc(x.title || '未命名')}${catOf(x) ? '（' + esc(catOf(x)) + '）' : ''}</option>`).join('')}</select></div><div class="sub">可以寫 {客戶}、{日期}，從案件複製時會自動帶入。</div>` : ''}<label class="f">${n.kind === 'quick' ? '名稱' : '標題'}<input id="eT" value="${esc(n.title)}"></label>${n.kind === 'quick' ? `<label class="f">電話<input id="eB" inputmode="tel" value="${esc(n.body)}"></label>` : `<label class="f">${n.kind === 'sop' ? '作業步驟' : '內容'}<textarea id="eB" rows="${n.kind === 'sop' ? 8 : 10}">${esc(n.body)}</textarea></label><div id="eImgBox">${imgsBlock()}</div>`}${n.kind === 'sop' ? `<div class="f etplbox"><b style="font-size:14px;color:var(--ink)">給客戶的範本</b><span class="sub">這個作業要傳給客戶的 LINE／簡訊，可以寫 {客戶}、{日期}。會同時出現在「範本」頁和同類案件裡。</span><div id="eTpls">${notes().filter((t) => t.kind === 'tpl' && sopsOf(t).includes(n.id)).map(tplRowHTML).join('')}</div><div class="actions" style="margin:0"><button type="button" class="btn" id="eAddTpl">＋ 寫一則新範本</button><select id="eLinkTpl" class="linktpl"><option value="">🔗 連結已有的範本…</option>${notes().filter((t) => t.kind === 'tpl' && !sopsOf(t).includes(n.id)).sort((a, b) => (b.cat === catOf(n)) - (a.cat === catOf(n)) || (a.title || '').localeCompare(b.title || '')).map((t) => `<option value="${t.id}">${esc(t.title || '未命名')}${t.cat ? '（' + esc(t.cat) + '）' : ''}</option>`).join('')}</select></div></div>` : ''}${n.kind === 'quick' && !isNew ? '<div class="actions"><button class="btn danger" id="eDel">刪除這支電話</button></div>' : ''}</div></div>`;
  const close = () => { NE = null; ov.innerHTML = ''; App.render(); };
  if ($('eS')) {
    $('eS').onchange = (e) => { const id = e.target.value; e.target.value = ''; if (!id || $('eSops').querySelector(`[data-sopid="${id}"]`)) return; const x = Vault.get('notes', id); $('eSops').insertAdjacentHTML('beforeend', `<button type="button" class="tag sm on" data-sopid="${id}">${esc(x.title || '未命名')} ✕</button>`); };
    $('eSops').onclick = (e) => { const b = e.target.closest('[data-sopid]'); if (b) b.remove(); };
  }
  if ($('eAlsoSel')) {
    $('eAlsoSel').value = '';
    $('eAlsoSel').onchange = (e) => { const k = e.target.value; e.target.value = ''; if (!k || $('eAlso').querySelector(`[data-also="${CSS.escape(k)}"]`)) return; e.target.insertAdjacentHTML('beforebegin', `<button type="button" class="tag sm on" data-also="${esc(k)}">${esc(k)} ✕</button>`); };
    $('eAlso').onclick = (e) => { const b = e.target.closest('[data-also]'); if (b) b.remove(); };
  }
  if ($('eLinkTpl')) $('eLinkTpl').onchange = (e) => { const t = Vault.get('notes', e.target.value); e.target.value = ''; if (!t || $('eTpls').querySelector(`[data-tid="${t.id}"]`)) return; $('eTpls').insertAdjacentHTML('beforeend', tplRowHTML(t)); toast(`已連結「${t.title}」，按儲存才會生效`); };
  if ($('eAddTpl')) {
    $('eAddTpl').onclick = () => { $('eTpls').insertAdjacentHTML('beforeend', tplRowHTML(null)); $('eTpls').lastElementChild.querySelector('.etb').focus(); };
    $('eTpls').onclick = (e) => { const ch = e.target.closest('.etch [data-t]'); if (ch) { ch.classList.toggle('on'); return; } const x = e.target.closest('.etx'); if (!x) return; const row = x.closest('.etpl'); row.hidden = true; row.dataset.del = 1; if (row.dataset.tid) toast('已取消連結，範本還會留在「範本」頁'); };
  }
  const saveTpls = () => { // SOP 底下的範本跟著 SOP 一起存
    ov.querySelectorAll('.etpl').forEach((row) => {
      const tid = row.dataset.tid, old = tid && Vault.get('notes', tid);
      if (row.dataset.del) { if (old) { old.sopIds = sopsOf(old).filter((x) => x !== n.id); delete old.sopId; Vault.put('notes', old); } row.remove(); return; }
      const body = row.querySelector('.etb').value, title = row.querySelector('.ett').value.trim(), channels = [...row.querySelectorAll('.etch .tag.on')].map((x) => x.dataset.t), group = channels[0] || '其他';
      if (!old && !body.trim() && !title) return;
      const t = old || { kind: 'tpl', images: [], cat: n.cat || '', line: n.line || lineN(n) }; // 已有的範本只更新內容和連結，不改它自己的分類
      Object.assign(t, { group, channels, title: title || n.title || '未命名', body, sopIds: [...new Set([...sopsOf(t), n.id])] }); delete t.sopId;
      Vault.put('notes', t); row.dataset.tid = t.id;
    });
  };
  ov.querySelectorAll('[data-eline]').forEach((b) => (b.onclick = () => {
    ov.querySelectorAll('[data-eline]').forEach((x) => x.classList.remove('on', 'st-go')); b.classList.add('on', 'st-go');
    const ln = b.dataset.eline; ov.querySelectorAll('[data-tagrow]').forEach((r) => { r.style.display = tagShown(r.dataset.tagrow, ln) ? '' : 'none'; });
  }));
  ov.querySelectorAll('[data-etag]').forEach((box) => (box.onclick = (e) => { const t = e.target.closest('[data-t]'); if (t) t.classList.toggle('on'); }));
  const collect = () => {
    n.title = $('eT').value.trim() || '未命名'; n.body = $('eB').value;
    if ($('eG')) n.group = $('eG').value; if ($('eC')) n.cat = $('eC').value; if ($('eSops')) { n.sopIds = [...$('eSops').querySelectorAll('[data-sopid]')].map((b) => b.dataset.sopid); delete n.sopId; }
    const ln = ov.querySelector('[data-eline].on'); if (ln) n.line = ln.dataset.eline;
    if ($('eAlso')) n.alsoCats = [...$('eAlso').querySelectorAll('[data-also]')].map((b) => b.dataset.also).filter((k) => k !== n.cat);
    if (n.kind === 'tpl') { const box = ov.querySelector('[data-etag="channels"]'); if (box) { n.channels = [...box.querySelectorAll('.tag.on')].map((x) => x.dataset.t); n.group = n.channels[0] || '其他'; } }
    if (n.kind === 'sop') { n.v2 = true; ov.querySelectorAll('[data-etag]').forEach((box) => { n[box.dataset.etag] = [...box.querySelectorAll('.tag.on')].map((x) => x.dataset.t); }); saveTpls(); }
  };
  const wireImgs = () => {
    if (!withImgs) return;
    $('eImgBox').innerHTML = imgsBlock(); drawThumbs();
    $('eAddImg').onclick = () => $('eImgs').click();
    $('eCam').onclick = () => { // 先把正在打的字存起來，掃描完再回到編輯
      collect(); delete n.refresh; Vault.put('notes', n);
      NE = null;
      openScanner(n, {
        onDone: async (pages) => { for (const p of pages) { n.images = [...(n.images || []), await Files.add(p.blob, n.title || p.name)]; URL.revokeObjectURL(p.url); } if (pages.length) Vault.put('notes', n); editNote(id, isNew); },
        onCancel: () => editNote(id, isNew),
      });
    };
    $('eImgs').onchange = async (e) => { const fs = [...e.target.files]; e.target.value = ''; toast('加入中…'); await addNoteImages(n, fs); wireImgs(); };
    $('eImgBox').querySelectorAll('[data-vimg]').forEach((b) => (b.onclick = () => openViewer({ fileId: b.dataset.vimg, name: n.title || '附圖' }, true)));
    $('eImgBox').querySelectorAll('[data-mkimg]').forEach((b) => (b.onclick = () => { // 標註：先存好正在打的字，標完回到編輯
      const f = n.images[+b.dataset.mkimg], m = n.imgMarks?.[f] || {};
      collect(); delete n.refresh; Vault.put('notes', n); NE = null;
      openMask({ docs: [{ fileId: f, name: n.title || '附圖', masks: m.masks || [], marks: m.marks || [], note: m.note || '' }] }, 0, 'notes', {
        title: '標註', tool: 'circle',
        onSave: (r) => { n.imgMarks = { ...(n.imgMarks || {}) }; if (hasMark(r) || r.note) n.imgMarks[f] = r; else delete n.imgMarks[f]; Vault.put('notes', n); editNote(id, isNew); toast('已存標註，原圖保留'); },
        onCancel: () => editNote(id, isNew),
      });
    }));
    $('eImgBox').querySelectorAll('[data-rotimg]').forEach((b) => (b.onclick = async () => { // 轉方向：存成新的一張，換掉舊的
      if (b.disabled) return; b.disabled = true; const i = +b.dataset.rotimg, old = n.images[i];
      try { const blob = await rotateBlob(old); const nid = await Files.add(blob, n.title || '附圖'); n.images[i] = nid; if (n.imgMarks?.[old]) { n.imgMarks[nid] = rotateMarks(n.imgMarks[old]); delete n.imgMarks[old]; } Files.remove(old); Vault.put('notes', n); wireImgs(); }
      catch (e) { toast('轉不了，請確認網路後再試'); b.disabled = false; }
    }));
    $('eImgBox').querySelectorAll('[data-delimg]').forEach((b) => (b.onclick = () => { const [f] = n.images.splice(+b.dataset.delimg, 1); Files.remove(f); if (n.imgMarks) delete n.imgMarks[f]; Vault.put('notes', n); wireImgs(); }));
  };
  NE && (NE.refresh = wireImgs);
  wireImgs();
  $('eCancel').onclick = () => { if (isNew) { (n.images || []).forEach((f) => Files.remove(f)); notes().filter((t) => t.kind === 'tpl' && sopsOf(t).includes(id)).forEach((t) => { if (sopsOf(t).length === 1) Vault.del('notes', t.id); else { t.sopIds = sopsOf(t).filter((x) => x !== id); Vault.put('notes', t); } }); Vault.del('notes', id); } close(); };
  $('eSave').onclick = () => { collect(); delete n.refresh; Vault.put('notes', n); close(); };
  if ($('eDel')) $('eDel').onclick = () => { Vault.del('notes', id); close(); };
}

/* ================= 批次匯入（Claude 整理好的匯入碼） ================= */
const KIND_LABEL = { sop: 'SOP', quick: '電話', tpl: '範本' };
function parseImport(text) {
  const t = text.trim().replace(/^```[a-z]*\s*|```\s*$/g, '');
  let arr; try { arr = JSON.parse(t); } catch (e) { return { err: '匯入碼格式不對，請整段完整複製貼上' }; }
  if (!Array.isArray(arr)) arr = [arr];
  const items = arr.filter((x) => x && typeof x === 'object' && KIND_LABEL[x.kind] && (x.title || x.body)).map((x) => ({
    kind: x.kind,
    group: x.kind === 'sop' ? String(x.group || '') : x.kind === 'tpl' ? (TPLGROUPS.includes(x.group) ? x.group : '其他') : '電話',
    title: String(x.title || '未命名').slice(0, 200), body: String(x.body || ''), cat: CATS.includes(x.cat) ? x.cat : x.kind === 'sop' ? mapGroup(x.group) : '',
    ...(x.kind === 'tpl' ? { channels: Array.isArray(x.channels) ? x.channels.map(String) : [TPLGROUPS.includes(x.group) ? x.group : '其他'] } : {}),
    ...(x.kind === 'sop' ? { v2: true, kinds: Array.isArray(x.kinds) ? x.kinds.map(String) : [], tools: Array.isArray(x.tools) ? x.tools.map(String) : (x.group === LEGACY_T ? ['T字障'] : []) } : {}),
  }));
  if (!items.length) return { err: '裡面沒有可以匯入的項目' };
  return { items };
}
function openImport() {
  ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>批次匯入</h2><button class="btn" id="imCancel">取消</button><button class="btn pri" id="imGo">檢查</button></div>
  <div class="ovbody" style="flex-direction:column;max-width:760px;margin:0 auto;width:100%">
   <div class="note">把 Claude 給你的匯入碼整段貼進來，按「檢查」，確認內容後再匯入。之後每一張都可以再點「編輯」修改。</div>
   <label class="f">匯入碼<textarea id="imText" rows="12" style="font-family:var(--f-num);font-size:13px"></textarea></label>
   <div class="err" id="imErr" style="color:var(--danger);font-size:13px"></div><div id="imPrev"></div></div></div>`;
  $('imCancel').onclick = () => (ov.innerHTML = '');
  $('imText').oninput = () => { $('imErr').textContent = ''; $('imPrev').innerHTML = ''; $('imGo').textContent = '檢查'; $('imGo').onclick = check; };
  const check = () => {
    const r = parseImport($('imText').value);
    if (r.err) { $('imErr').textContent = r.err; return; }
    const dup = new Set(notes().map((n) => n.kind + '|' + n.title));
    $('imPrev').innerHTML = `<div class="sub" style="margin:4px 0 8px">共 ${r.items.length} 項${r.items.some((x) => dup.has(x.kind + '|' + x.title)) ? '（標「已有」的是同名項目，會另外新增一張，不會蓋掉）' : ''}</div><div class="cards">${r.items.map((x) => `<div class="card"><h3><span>${esc(x.title)}</span><span class="sub">${KIND_LABEL[x.kind]} · ${esc(x.group)}${dup.has(x.kind + '|' + x.title) ? ' · 已有' : ''}</span></h3><pre>${esc(x.body)}</pre></div>`).join('')}</div>`;
    $('imGo').textContent = `匯入 ${r.items.length} 項`;
    $('imGo').onclick = () => { r.items.forEach((x) => Vault.put('notes', { ...x })); ov.innerHTML = ''; App.render(); toast(`已匯入 ${r.items.length} 項`); };
  };
  $('imGo').onclick = check;
  setTimeout(() => $('imText')?.focus(), 50);
}

/* ================= 掃描 ================= */
const fileIn = $('fileIn'), fileLib = $('fileLib');
let SC = null;
const CASE_DOC_NAMES = ['醫療收據', '診斷證明', '申請書', '其他'];
const CLIENT_DOC_NAMES = ['身分證正面', '身分證反面', '存摺封面', '健保卡', '其他'];
function openScanner(c, opts = {}) { SC = { c, folder: opts.folder || '', coll: opts.coll || 'cases', names: opts.names || CASE_DOC_NAMES, pages: [], img: null, corners: null, filter: opts.filter || 'gray', level: 50, drag: -1, queue: [], onDone: opts.onDone, onCancel: opts.onCancel }; renderScanner(); }
function scanForClient(cl) { if (!cl) { toast('先選一位客戶'); return; } openScanner(cl, { coll: 'clients', names: CLIENT_DOC_NAMES, filter: 'color' }); }
function renderScanner() {
  ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>掃描文件</h2><span class="sub">${SC.pages.length} 頁</span><button class="btn" id="scCancel">取消</button><button class="btn pri" id="scDone">完成（${SC.pages.length + (SC.img ? 1 : 0)}）</button></div>
  <div class="ovbody">${SC.img ? `<div class="stage"><canvas id="scCv"></canvas></div>
  <div class="side"><div class="note">拖曳四個角對齊文件邊緣；拖四邊中間的點可以整條邊一起移。對好按「加入這頁」，最後一張直接按「完成」也會收進去。</div>
   <div><div class="sub" style="margin-bottom:4px">效果</div><div class="seg">${[['color', '原色'], ['gray', '灰階增強'], ['bw', '黑白']].map(([k, l]) => `<button data-flt="${k}" class="${SC.filter === k ? 'on' : ''}">${l}</button>`).join('')}</div></div>
   ${SC.filter !== 'color' ? `<label class="f">${SC.filter === 'bw' ? '黑白程度' : '對比程度'}<span class="sub">（往右字越黑、底越白；字太淡就往右，糊成一團就往左）</span><input type="range" id="scLv" min="0" max="100" value="${SC.level}"></label>` : ''}
   <div class="scprev"><div class="sub" style="margin-bottom:4px">完成後的樣子</div><canvas id="scPv"></canvas></div>
   <label class="f">文件名稱<input id="scName" value="${SC.names[Math.min(SC.pages.length, SC.names.length - 1)]}"></label>
   <div class="actions"><button class="btn pri" id="scAdd">加入這頁</button><button class="btn" id="scRetake">重拍</button><button class="btn" id="scFull">用整張圖</button><button class="btn" id="scRot">↻ 轉方向</button></div>
   ${SC.pages.length ? `<div class="pages">${SC.pages.map((p) => `<img src="${p.url}" alt="${esc(p.name)}">`).join('')}</div>` : ''}</div>`
    : `<div class="pick">${ICON.scan}<div>拍一張文件、從相簿選照片，<br>或把圖片拖進來、按 Ctrl+V 貼上截圖</div><div class="actions" style="justify-content:center"><button class="btn pri" id="scCam">開相機</button><button class="btn" id="scLib">從相簿選</button></div>${SC.pages.length ? `<div class="pages">${SC.pages.map((p) => `<img src="${p.url}" alt="">`).join('')}</div>` : ''}</div>`}
  </div></div>`;
  $('scCancel').onclick = () => { const s = SC; s.pages.forEach((p) => URL.revokeObjectURL(p.url)); ov.innerHTML = ''; SC = null; if (s.onCancel) s.onCancel(); else App.render(); };
  $('scDone').onclick = async () => {
    if (SC.img) { // 還沒按「加入這頁」就按完成：把眼前這張也收進去，不要弄丟
      $('scDone').disabled = true; toast('整理中…');
      SC.pages.push({ name: $('scName').value || '文件', blob: await warp(SC.img, SC.corners, SC.filter, SC.level), url: '' });
      for (const f of SC.queue.splice(0)) { const cv = await fileToCanvas(f); if (!cv) continue; const w = cv.width, h = cv.height; SC.pages.push({ name: SC.names[Math.min(SC.pages.length, SC.names.length - 1)], blob: await warp(cv, [[0, 0], [w, 0], [w, h], [0, h]], SC.filter, SC.level), url: '' }); }
      SC.img = null;
    }
    const s = SC; ov.innerHTML = ''; SC = null;
    if (s.onDone) { if (s.pages.length) toast('加入中…'); await s.onDone(s.pages); return; }
    if (s.pages.length) {
      for (const p of s.pages) { const fileId = await Files.add(p.blob, p.name); s.c.docs.push({ id: uid(), name: p.name, fileId, masks: [], folder: s.folder || '', date: localDay() }); }
      Vault.put(s.coll, s.c); toast(`已加入 ${s.pages.length} 頁`);
    }
    App.render();
  };
  if (!SC.img) { $('scCam').onclick = () => fileIn.click(); $('scLib').onclick = () => fileLib.click(); return; }
  ov.querySelectorAll('[data-flt]').forEach((b) => (b.onclick = () => { SC.filter = b.dataset.flt; renderScanner(); }));
  if ($('scLv')) $('scLv').oninput = (e) => { SC.level = +e.target.value; scPreview(); };
  $('scRetake').onclick = () => { SC.img = null; nextQueued(); };
  $('scRot').onclick = () => { const h = SC.img.height; SC.img = rotateCanvas(SC.img); SC.corners = SC.corners.map(([x, y]) => [h - y, x]); SC.corners = [3, 0, 1, 2].map((i) => SC.corners[i]); renderScanner(); };
  $('scFull').onclick = () => { const w = SC.img.width, h = SC.img.height; SC.corners = [[0, 0], [w, 0], [w, h], [0, h]]; renderScanner(); };
  $('scAdd').onclick = async () => { const blob = await warp(SC.img, SC.corners, SC.filter, SC.level); SC.pages.push({ name: $('scName').value || '文件', blob, url: URL.createObjectURL(blob) }); SC.img = null; nextQueued(); };
  setupCornerCanvas(); scPreview();
}
/* 拖曳／貼上圖片：在任何畫面都可以，會加進目前的案件（沒有就新建一個） */
let addingImgs = Promise.resolve();
function addImagesDirect(rec, coll, files, folder) { // 貼上、拖曳、上傳的圖片直接加入，不經過拉四個角
  addingImgs = addingImgs.then(async () => {
    toast(files.length > 1 ? `加入 ${files.length} 張中…` : '加入中…');
    let n = 0;
    for (const f of files) { const blob = await shrink(f); if (!blob) continue; const fileId = await Files.add(blob, '圖片'); rec.docs = rec.docs || []; rec.docs.push({ id: uid(), name: (f.name && !/^image\.(png|jpe?g)$/i.test(f.name) ? f.name.replace(/\.[a-z]+$/i, '') : '圖片'), fileId, masks: [], folder: folder || '', date: localDay() }); n++; }
    if (n) { Vault.put(coll, rec); App.render(); toast(`已加入 ${n} 張`); } else toast('圖片打不開，換一張試試');
  });
}
async function pasteFromClipboard() { // iPad：在 LINE 長按圖片 › 拷貝，再按「貼上圖片」
  try {
    const items = await navigator.clipboard.read(); const files = [];
    for (const it of items) { const t = it.types.find((x) => x.startsWith('image/')); if (t) { const b = await it.getType(t); files.push(new File([b], 'image.' + (t.split('/')[1] || 'png'), { type: t })); } }
    if (!files.length) { toast('剪貼簿裡沒有圖片。先在 LINE 長按圖片選「拷貝」'); return; }
    takeImages(files);
  } catch (e) { toast('這裡不能讀剪貼簿。電腦請直接按 Ctrl+V'); }
}
function takeImages(files) {
  const imgs = [...files].filter((f) => /^image\//.test(f.type));
  const pdfs = [...files].filter((f) => f.type === 'application/pdf' || /\.pdf$/i.test(f.name || ''));
  if (pdfs.length) {
    if (NE || SC || S.view === 'sop' || S.view === 'tpl') toast('PDF 只能放進案件或客戶');
    else if (S.view === 'clients') { const cl = curClient(); if (cl) addPdfs(cl, 'clients', pdfs); else toast('先選一位客戶'); }
    else { let t = S.view === 'cases' ? cur() : null; if (!t) { t = newCase('理賠'); S.view = 'cases'; } addPdfs(t, 'cases', pdfs, curFolder(t)); }
  }
  if (!imgs.length) { if (files.length && !pdfs.length) toast('只能放圖片或 PDF'); return; }
  if (NE) { toast('加入中…'); addNoteImages(NE, imgs).then(() => NE?.refresh?.()); return; }
  if (!SC && S.view === 'clients') { const cl = curClient(); if (!cl) { toast('先選一位客戶，再把證件拖進來'); return; } addImagesDirect(cl, 'clients', imgs, ''); return; }
  if (!SC && (S.view === 'sop' || S.view === 'tpl')) { const n = Vault.put('notes', { kind: S.view, group: S.view === 'sop' ? '' : '其他', ...(S.view === 'sop' ? { v2: true, cat: S.sc !== '全部' ? S.sc : '', kinds: S.sk ? [S.sk] : [], tools: S.st ? [S.st] : [] } : {}), title: '', body: '', images: [] }); editNote(n.id, true); toast('加入中…'); addNoteImages(n, imgs).then(() => NE?.refresh?.()); return; }
  if (!SC) { let t = S.view === 'cases' ? cur() : null; if (!t) { t = newCase('理賠'); S.view = 'cases'; } addImagesDirect(t, 'cases', imgs, curFolder(t)); return; }
  if (SC.img) SC.queue.push(...imgs); else onFiles(imgs);
  if (imgs.length > 1) toast(`收到 ${imgs.length} 張，一張一張對齊後按「加入這頁」`);
}
// ── 文件拖曳：排順序、拖進資料夾（滑鼠直接拖；觸控長按後拖）──
let DD = null;
function ddRec(p) { return p === 'c' ? [curClient(), 'clients'] : [cur(), 'cases']; }
function ddScroller(el) { for (let n = el; n && n !== document.body; n = n.parentElement) { const o = getComputedStyle(n).overflowY; if ((o === 'auto' || o === 'scroll') && n.scrollHeight > n.clientHeight) return n; } return document.scrollingElement; }
function ddClear() { document.querySelectorAll('.ins-before,.ins-after,.droptgt').forEach((n) => n.classList.remove('ins-before', 'ins-after', 'droptgt')); }
function ddBegin() {
  const t = DD.tile, r = t.getBoundingClientRect(), g = t.cloneNode(true);
  t.querySelectorAll('canvas').forEach((cv, k) => { const gc = g.querySelectorAll('canvas')[k]; gc.getContext('2d').drawImage(cv, 0, 0); });
  g.classList.add('ddghost'); g.style.width = r.width + 'px'; g.style.left = r.left + 'px'; g.style.top = r.top + 'px';
  DD.dx = DD.x0 - r.left; DD.dy = DD.y0 - r.top; DD.ghost = g; DD.on = true; DD.scroller = ddScroller(t);
  document.body.appendChild(g); t.classList.add('ddsrc'); document.body.classList.add('dragging-doc');
  navigator.vibrate?.(15);
  DD.tick = setInterval(() => { if (!DD?.on) return; const h = window.innerHeight, v = DD.y < 70 ? -12 : DD.y > h - 70 ? 12 : 0; if (v) { DD.scroller.scrollBy(0, v); ddMove(DD.x, DD.y); } }, 16);
}
function ddMove(x, y) {
  DD.x = x; DD.y = y; if (!DD.on) return;
  DD.ghost.style.transform = `translate(${x - DD.dx - parseFloat(DD.ghost.style.left)}px,${y - DD.dy - parseFloat(DD.ghost.style.top)}px) scale(.9) rotate(-2deg)`;
  ddClear(); DD.target = null;
  const el = document.elementFromPoint(x, y); if (!el) return;
  const chip = DD.p === '' && el.closest('.chip[data-folder]');
  if (chip && chip.dataset.folder !== '全部') { chip.classList.add('droptgt'); DD.target = { folder: chip.dataset.folder }; return; }
  const t = el.closest('.doc[data-di]');
  if (t && t !== DD.tile && t.parentElement === DD.tile.parentElement) { const r = t.getBoundingClientRect(), before = x < r.left + r.width / 2; t.classList.add(before ? 'ins-before' : 'ins-after'); DD.target = { di: +t.dataset.di, before }; }
}
function ddEnd(cancel) {
  const d = DD; DD = null; if (!d) return; clearTimeout(d.timer); clearInterval(d.tick);
  if (!d.on) return;
  d.ghost.remove(); d.tile.classList.remove('ddsrc'); document.body.classList.remove('dragging-doc'); ddClear();
  window.addEventListener('click', (e) => { e.stopPropagation(); e.preventDefault(); }, { capture: true, once: true }); setTimeout(() => window.dispatchEvent(new Event('click')), 0);
  if (cancel || !d.target) return;
  const [rec, coll] = ddRec(d.p); if (!rec) return; const src = d.di;
  if (d.target.folder != null) {
    const f = d.target.folder === '未分類' ? '' : d.target.folder;
    if ((rec.docs[src].folder || '') === f) return;
    rec.docs[src].folder = f; Vault.put(coll, rec); App.render(); toast(`已移到「${d.target.folder}」`); return;
  }
  const order = docIdx(rec).filter((i) => i !== src); let pos = order.indexOf(d.target.di); if (pos < 0) return;
  order.splice(d.target.before ? pos : pos + 1, 0, src);
  rec.docs = order.map((i) => rec.docs[i]); rec.docManual = true; Vault.put(coll, rec); App.render();
}
function ddDown(e, x, y, touch) {
  if (!Vault.key || DD) return;
  const tile = e.target.closest?.('.doc[data-di]'); if (!tile || e.target.closest('button,input,select,label,a,textarea')) return;
  DD = { tile, p: tile.dataset.dp || '', di: +tile.dataset.di, x0: x, y0: y, x, y, touch, on: false };
  if (touch) { const me = DD; DD.timer = setTimeout(() => { if (DD === me) ddBegin(); }, 380); }
}
document.addEventListener('mousedown', (e) => { if (e.button === 0) ddDown(e, e.clientX, e.clientY, false); });
document.addEventListener('mousemove', (e) => { if (!DD || DD.touch) return; if (!DD.on && Math.hypot(e.clientX - DD.x0, e.clientY - DD.y0) > 6) ddBegin(); if (DD.on) { e.preventDefault(); ddMove(e.clientX, e.clientY); } });
document.addEventListener('mouseup', () => { if (DD && !DD.touch) ddEnd(false); });
document.addEventListener('touchstart', (e) => { if (e.touches.length === 1) ddDown(e, e.touches[0].clientX, e.touches[0].clientY, true); else ddEnd(true); }, { passive: true });
document.addEventListener('touchmove', (e) => { if (!DD || !DD.touch) return; const t = e.touches[0]; if (!DD.on) { if (Math.hypot(t.clientX - DD.x0, t.clientY - DD.y0) > 10) { clearTimeout(DD.timer); DD = null; } return; } e.preventDefault(); ddMove(t.clientX, t.clientY); }, { passive: false });
document.addEventListener('touchend', () => { if (DD?.touch) ddEnd(false); });
document.addEventListener('touchcancel', () => ddEnd(true));
document.addEventListener('contextmenu', (e) => { if (DD || e.target.closest?.('.doc[data-di] canvas')) e.preventDefault(); });
document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && DD) ddEnd(true); });
const hasFiles = (e) => [...(e.dataTransfer?.types || [])].includes('Files');
let dragDepth = 0;
document.addEventListener('dragenter', (e) => { if (!Vault.key || !hasFiles(e)) return; e.preventDefault(); dragDepth++; document.body.classList.add('dropping'); });
document.addEventListener('dragover', (e) => { if (!Vault.key || !hasFiles(e)) return; e.preventDefault(); e.dataTransfer.dropEffect = 'copy'; });
document.addEventListener('dragleave', () => { if (--dragDepth <= 0) { dragDepth = 0; document.body.classList.remove('dropping'); } });
document.addEventListener('drop', (e) => { if (!Vault.key || !hasFiles(e)) return; e.preventDefault(); dragDepth = 0; document.body.classList.remove('dropping'); if (MK) { toast('請先關閉遮罩畫面'); return; } if (DSC) { discAddFiles([...e.dataTransfer.files], false); return; } takeImages(e.dataTransfer.files); });
document.addEventListener('paste', (e) => {
  if (!Vault.key || MK) return;
  const files = [...(e.clipboardData?.items || [])].filter((i) => i.kind === 'file' && /^image\//.test(i.type)).map((i) => i.getAsFile()).filter(Boolean);
  if (!files.length) return; // 一般文字照常貼上
  e.preventDefault(); if (DSC) { discAddFiles(files, true); return; } if (CRMI) { crmFromImage(files[0]); return; } if (KPH) { callPhotoFrom(files[0]); return; } takeImages(files);
});
function nextQueued() { if (SC.queue.length) loadFile(SC.queue.shift()); else renderScanner(); }
function onFiles(files) { if (!SC || !files.length) return; SC.queue.push(...[...files].slice(1)); loadFile(files[0]); }
fileIn.onchange = () => { onFiles(fileIn.files); fileIn.value = ''; };
fileLib.onchange = () => { onFiles(fileLib.files); fileLib.value = ''; };
async function fileToCanvas(f) {
  const bmp = await createImageBitmap(f).catch(() => null); if (!bmp) return null;
  const s = Math.min(1, 3600 / Math.max(bmp.width, bmp.height)); // 原圖保留夠大，掃描後字才清楚
  const cv = document.createElement('canvas'); cv.width = Math.round(bmp.width * s); cv.height = Math.round(bmp.height * s); cv.getContext('2d').drawImage(bmp, 0, 0, cv.width, cv.height);
  return cv;
}
async function loadFile(f) {
  const cv = await fileToCanvas(f);
  if (!cv) { toast('這張圖片打不開，換一張試試'); return; }
  SC.img = cv; const w = cv.width, h = cv.height, m = 0.06; SC.corners = [[w * m, h * m], [w * (1 - m), h * m], [w * (1 - m), h * (1 - m)], [w * m, h * (1 - m)]];
  renderScanner();
}
let scPvT = null;
function scPreview() { // 側邊的小預覽，用小尺寸算，拖曳時不卡
  clearTimeout(scPvT); scPvT = setTimeout(async () => {
    const pv = $('scPv'); if (!pv || !SC?.img) return;
    const out = await warp(SC.img, SC.corners, SC.filter, SC.level, 520, true);
    pv.width = out.width; pv.height = out.height; pv.getContext('2d').drawImage(out, 0, 0);
  }, 60);
}
function setupCornerCanvas() {
  const cv = $('scCv'); cv.width = SC.img.width; cv.height = SC.img.height; const x = cv.getContext('2d'); const R = Math.max(cv.width, cv.height) / 40;
  const draw = () => {
    x.drawImage(SC.img, 0, 0); const k = SC.corners;
    x.fillStyle = 'rgba(20,22,16,.45)'; x.beginPath(); x.rect(0, 0, cv.width, cv.height); x.moveTo(...k[0]); x.lineTo(...k[3]); x.lineTo(...k[2]); x.lineTo(...k[1]); x.closePath(); x.fill('evenodd');
    x.strokeStyle = '#9DBA88'; x.lineWidth = R / 4; x.beginPath(); k.forEach((p, i) => (i ? x.lineTo(...p) : x.moveTo(...p))); x.closePath(); x.stroke();
    k.forEach((p) => { x.beginPath(); x.arc(p[0], p[1], R, 0, 7); x.fillStyle = 'rgba(157,186,136,.35)'; x.fill(); x.lineWidth = R / 5; x.stroke(); });
    mids().forEach((p) => { x.beginPath(); x.arc(p[0], p[1], R * 0.62, 0, 7); x.fillStyle = 'rgba(255,255,255,.85)'; x.fill(); x.lineWidth = R / 5; x.stroke(); });
  };
  // 四邊中間的點：拖它會讓這條邊的兩個角一起移動（0-1 上、1-2 右、2-3 下、3-0 左）
  const EDGE = [[0, 1], [1, 2], [2, 3], [3, 0]];
  const mids = () => EDGE.map(([a, b]) => [(SC.corners[a][0] + SC.corners[b][0]) / 2, (SC.corners[a][1] + SC.corners[b][1]) / 2]);
  const clamp = (p) => [Math.max(0, Math.min(cv.width, p[0])), Math.max(0, Math.min(cv.height, p[1]))];
  draw();
  const pt = (e) => { const r = cv.getBoundingClientRect(); return [(e.clientX - r.left) * cv.width / r.width, (e.clientY - r.top) * cv.height / r.height]; };
  let start = null;
  cv.onpointerdown = (e) => {
    const p = pt(e); let best = -1, bd = R * 3.5;
    [...SC.corners, ...mids()].forEach((c, i) => { const d = Math.hypot(c[0] - p[0], c[1] - p[1]) + (i >= 4 ? R * 0.3 : 0); if (d < bd) { bd = d; best = i; } }); // 角和中點很近時，優先抓角
    SC.drag = best; start = best >= 4 ? { p, k: SC.corners.map((c) => [...c]) } : null;
    if (best >= 0) cv.setPointerCapture(e.pointerId);
  };
  cv.onpointermove = (e) => {
    if (SC.drag < 0) return; const p = pt(e);
    if (SC.drag < 4) SC.corners[SC.drag] = clamp(p);
    else { const [a, b] = EDGE[SC.drag - 4], dx = p[0] - start.p[0], dy = p[1] - start.p[1]; SC.corners[a] = clamp([start.k[a][0] + dx, start.k[a][1] + dy]); SC.corners[b] = clamp([start.k[b][0] + dx, start.k[b][1] + dy]); }
    draw();
  };
  cv.onpointerup = cv.onpointercancel = () => { if (SC.drag >= 0) scPreview(); SC.drag = -1; start = null; };
}
async function rotateBlob(fileId) { // 向右轉 90 度
  const u = await Files.url(fileId); if (!u) throw new Error('讀不到圖片');
  const bmp = await createImageBitmap(await (await fetch(u)).blob());
  const cv = document.createElement('canvas'); cv.width = bmp.height; cv.height = bmp.width; const x = cv.getContext('2d');
  x.translate(cv.width, 0); x.rotate(Math.PI / 2); x.drawImage(bmp, 0, 0);
  return await new Promise((ok) => cv.toBlob(ok, 'image/jpeg', 0.9));
}
let pdfjsReady = null;
function loadPdfjs() { // 讀 PDF 的元件，第一次用到才載入
  if (window.pdfjsLib) return Promise.resolve(window.pdfjsLib);
  return (pdfjsReady ||= new Promise((ok, no) => { const sc = document.createElement('script'); sc.src = 'vendor/pdfjs/pdf.min.js'; sc.onload = () => { window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'vendor/pdfjs/pdf.worker.min.js'; ok(window.pdfjsLib); }; sc.onerror = () => { pdfjsReady = null; no(new Error('load')); }; document.head.appendChild(sc); }));
}
async function pdfToImages(rec, coll, i, btn) { // PDF 每一頁轉成清楚的圖片，放在 PDF 後面；原本的 PDF 保留
  const d = rec.docs[i]; if (!d || btn.disabled) return; btn.disabled = true; const label = btn.textContent;
  try {
    const lib = await loadPdfjs(); const u = await Files.url(d.fileId); if (!u) throw new Error('file');
    const pdf = await lib.getDocument({ data: new Uint8Array(await (await fetch(u)).arrayBuffer()), cMapUrl: 'vendor/pdfjs/cmaps/', cMapPacked: true }).promise;
    const out = [];
    for (let n = 1; n <= pdf.numPages; n++) {
      btn.textContent = `轉換中 ${n}/${pdf.numPages}`;
      const page = await pdf.getPage(n); const v0 = page.getViewport({ scale: 1 }); const sc = Math.min(4, 2400 / Math.max(v0.width, v0.height));
      const vp = page.getViewport({ scale: sc }); const cv = document.createElement('canvas'); cv.width = Math.round(vp.width); cv.height = Math.round(vp.height);
      const x = cv.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, cv.width, cv.height);
      await page.render({ canvasContext: x, viewport: vp }).promise;
      const blob = await new Promise((ok) => cv.toBlob(ok, 'image/jpeg', 0.92)); const name = pdf.numPages > 1 ? `${d.name} 第${n}頁` : d.name;
      out.push({ id: uid(), name, fileId: await Files.add(blob, name), masks: [], folder: d.folder || '', date: localDay() });
    }
    rec.docs.splice(i + 1, 0, ...out); coll === 'cases' ? saveCase(rec) : Vault.put(coll, rec); App.render();
    toast(`已轉成 ${out.length} 張圖片，原本的 PDF 還在`);
  } catch (e) { console.warn(e); toast('轉不了，請確認網路後再試'); btn.disabled = false; btn.textContent = label; }
}
async function rotateDoc(rec, coll, i, btn) { // 案件、客戶文件轉方向：遮罩和記號跟著轉
  const d = rec.docs[i]; if (!d || btn.disabled) return; btn.disabled = true;
  try {
    const old = d.fileId, blob = await rotateBlob(old); d.fileId = await Files.add(blob, d.name);
    const r = rotateMarks({ masks: d.masks || [], marks: d.marks || [] }); d.masks = r.masks; d.marks = r.marks;
    Files.remove(old); coll === 'cases' ? saveCase(rec) : Vault.put(coll, rec); App.render();
  } catch (e) { toast('轉不了，請確認網路後再試'); btn.disabled = false; }
}
function rotateCanvas(src) { const cv = document.createElement('canvas'); cv.width = src.height; cv.height = src.width; const x = cv.getContext('2d'); x.translate(cv.width, 0); x.rotate(Math.PI / 2); x.drawImage(src, 0, 0); return cv; }
function solveH(dst, src) {
  const A = [];
  for (let i = 0; i < 4; i++) { const [u, v] = dst[i], [x, y] = src[i]; A.push([u, v, 1, 0, 0, 0, -u * x, -v * x, x]); A.push([0, 0, 0, u, v, 1, -u * y, -v * y, y]); }
  for (let c = 0; c < 8; c++) { let p = c; for (let r = c + 1; r < 8; r++) if (Math.abs(A[r][c]) > Math.abs(A[p][c])) p = r; [A[c], A[p]] = [A[p], A[c]]; for (let r = 0; r < 8; r++) { if (r === c) continue; const f = A[r][c] / A[c][c]; for (let k = c; k < 9; k++) A[r][k] -= f * A[c][k]; } }
  return A.map((r, i) => r[8] / r[i]);
}
function warp(img, k, filter, level = 50, max = 2800, raw = false) { // level 0–100：越大越黑白分明
  const d = (a, b) => Math.hypot(a[0] - b[0], a[1] - b[1]);
  let W = Math.max(d(k[0], k[1]), d(k[3], k[2])), H = Math.max(d(k[0], k[3]), d(k[1], k[2])); const s = Math.min(1, max / Math.max(W, H)); W = Math.round(W * s); H = Math.round(H * s);
  const h = solveH([[0, 0], [W, 0], [W, H], [0, H]], k); const sd = img.getContext('2d').getImageData(0, 0, img.width, img.height).data, iw = img.width, ih = img.height;
  const out = document.createElement('canvas'); out.width = W; out.height = H; const ox = out.getContext('2d'); const od = ox.createImageData(W, H), o = od.data;
  const gy = new Float32Array(W * H);
  for (let v = 0; v < H; v++) for (let u = 0; u < W; u++) { // 雙線性取樣，字的邊緣比較平順不會鋸齒
    const z = h[6] * u + h[7] * v + 1; let x = (h[0] * u + h[1] * v + h[2]) / z, y = (h[3] * u + h[4] * v + h[5]) / z;
    x = Math.min(iw - 1.001, Math.max(0, x)); y = Math.min(ih - 1.001, Math.max(0, y)); const x0 = x | 0, y0 = y | 0, fx = x - x0, fy = y - y0;
    const i00 = (y0 * iw + x0) * 4, i10 = i00 + 4, i01 = i00 + iw * 4, i11 = i01 + 4, oi = (v * W + u) * 4;
    for (let c = 0; c < 3; c++) o[oi + c] = (sd[i00 + c] * (1 - fx) + sd[i10 + c] * fx) * (1 - fy) + (sd[i01 + c] * (1 - fx) + sd[i11 + c] * fx) * fy;
    o[oi + 3] = 255; gy[v * W + u] = 0.3 * o[oi] + 0.59 * o[oi + 1] + 0.11 * o[oi + 2];
  }
  if (filter !== 'color') { // 用附近的紙色當背景來比，陰影、角落偏暗也能變白
    const r = Math.max(8, Math.round(Math.max(W, H) / 24)), I = new Float64Array((W + 1) * (H + 1));
    for (let v = 0; v < H; v++) { let row = 0; for (let u = 0; u < W; u++) { row += gy[v * W + u]; I[(v + 1) * (W + 1) + u + 1] = I[v * (W + 1) + u + 1] + row; } }
    const L = Math.max(0, Math.min(100, level)) / 100, t = 0.78 + L * 0.14, soft = 0.28 - L * 0.16, lo = 0.25 + L * 0.35;
    for (let v = 0; v < H; v++) { const y1 = Math.max(0, v - r), y2 = Math.min(H, v + r + 1);
      for (let u = 0; u < W; u++) { const x1 = Math.max(0, u - r), x2 = Math.min(W, u + r + 1);
        const bg = (I[y2 * (W + 1) + x2] - I[y1 * (W + 1) + x2] - I[y2 * (W + 1) + x1] + I[y1 * (W + 1) + x1]) / ((x2 - x1) * (y2 - y1));
        const ratio = gy[v * W + u] / (bg * 1.06 + 1);
        let g = filter === 'gray' ? 255 * (ratio - lo) / (0.9 - lo) : 255 * (ratio - (t - soft)) / soft;
        g = g < 0 ? 0 : g > 255 ? 255 : g; const oi = (v * W + u) * 4; o[oi] = o[oi + 1] = o[oi + 2] = g; } }
  }
  ox.putImageData(od, 0, 0);
  if (raw) return out;
  return new Promise((ok) => out.toBlob((b) => ok(b), 'image/jpeg', 0.9));
}

/* ================= 看原圖 ================= */
// 依畫面上的排列順序翻頁看文件
function openGallery(btn, docs, key) {
  const tiles = [...btn.closest('.docs').querySelectorAll('.doc[data-di]')];
  const list = tiles.map((t) => docs[+t.dataset.di]).filter(Boolean);
  const i = list.indexOf(docs[+btn.dataset[key]]);
  openViewer(list[i], false, { list, i });
}
async function openViewer(d, overEdit, nav) {
  if (nav && (nav.list.length > 1 || nav.masked)) return showGallery(nav);
  const url = await Files.url(d.fileId); if (!url) { toast('檔案讀取失敗，請確認網路'); return; }
  if (isPdf(d)) {
    ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>${esc(d.name)}</h2><button class="btn" id="vShare">分享／儲存</button><button class="btn pri" id="vClose">關閉</button></div><div class="sub" style="padding:6px 16px">如果只看到第一頁，按「分享／儲存」用「檔案」或 PDF 閱讀器打開。</div><iframe src="${url}" title="${esc(d.name)}" style="flex:1;width:100%;border:0;background:#fff"></iframe></div>`;
    $('vClose').onclick = () => (ov.innerHTML = '');
    $('vShare').onclick = async () => { const blob = await (await fetch(url)).blob(); shareOrSave([new File([blob], d.name.replace(/[\\/:*?"<>|]/g, '_') + '.pdf', { type: 'application/pdf' })]); };
    return;
  }
  const nm = !d.masks && noteMark(d.fileId), vurl = hasMark(nm) ? await noteImgURL(d.fileId) : url, vnote = nm?.note ? `<div class="note" style="white-space:pre-wrap;margin:8px 16px 0">${esc(nm.note)}</div>` : '';
  if (overEdit) { const v = document.createElement('div'); v.innerHTML = `<div class="ov" style="z-index:30"><div class="ovhead"><h2>${esc(d.name)}</h2><button class="btn pri">關閉</button></div>${vnote}<div class="ovbody" style="justify-content:center"><img src="${vurl}" alt="" style="max-width:100%;height:auto;align-self:flex-start;border-radius:6px"></div></div>`; v.querySelector('button').onclick = () => v.remove(); document.body.appendChild(v); return; }
  ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>${esc(d.name)}${hasMark(nm) ? '' : '（原圖）'}</h2><button class="btn pri" id="vClose">關閉</button></div>${vnote}<div class="ovbody" style="justify-content:center"><img src="${vurl}" alt="${esc(d.name)}" style="max-width:100%;height:auto;align-self:flex-start;border-radius:6px"></div></div>`;
  $('vClose').onclick = () => (ov.innerHTML = '');
}

/* 傳送：文字＋圖片（剪貼簿一次只能放一樣，所以分開複製，或用分享一次傳） */
async function pngBlob(fileId) {
  const u = await noteImgURL(fileId); if (!u) throw new Error('讀不到圖片');
  const bmp = await createImageBitmap(await (await fetch(u)).blob());
  const cv = document.createElement('canvas'); cv.width = bmp.width; cv.height = bmp.height; cv.getContext('2d').drawImage(bmp, 0, 0);
  return await new Promise((ok) => cv.toBlob(ok, 'image/png'));
}
async function openSend(n, c) { // 傳送前可以微調文字、勾選要附的圖片
  const txt = fillTpl(n.body, c), imgs = n.images || [];
  const canShare = !!navigator.canShare;
  ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>傳送 · ${esc(n.title)}</h2><button class="btn pri" id="sdX">關閉</button></div>
  <div class="ovbody" style="flex-direction:column;max-width:680px;margin:0 auto;width:100%">
   <div class="sendbox"><div class="sub">可以直接在這裡改，改的只用在這次，不會動到範本${c?.client ? `（已帶入 ${esc(c.client)}）` : ''}</div>
    <textarea id="sdTxt" class="grow" rows="6">${esc(txt)}</textarea>
    <div class="actions" style="margin:0"><button class="btn pri" id="sdText">複製文字</button><button class="linkbtn" id="sdReset">恢復原文</button><button class="linkbtn" id="sdSave">把修改存回範本</button></div></div>
   ${imgs.length ? `<div class="sub">圖片 · 點圖片可以取消勾選，只傳要的</div>
   <div class="sendimgs">${imgs.map((f, i) => `<div class="sendimg on" data-sdi="${i}"><img data-sdimg="${f}" alt=""><span class="sdcheck">✓</span><button class="btn" data-sdcopy="${i}">複製這張</button></div>`).join('')}</div>` : ''}
   ${canShare ? `<button class="btn pri" id="sdShare" style="align-self:flex-start">分享（文字${imgs.length ? '＋勾選的圖片' : ''}）</button><div class="sub">會跳出分享選單，選 LINE 或要傳的人。LINE 有時只收圖片、沒帶到文字，那就再按「複製文字」貼上。</div>` : imgs.length ? '<div class="sub">電腦上剪貼簿一次只能放一樣：先複製文字貼上，再一張一張複製圖片貼上。</div>' : ''}
  </div></div>`;
  fitNotes();
  ov.querySelectorAll('[data-sdimg]').forEach((im) => Files.url(im.dataset.sdimg).then((u) => { if (u) im.src = u; }));
  ov.querySelectorAll('.sendimg').forEach((el) => (el.onclick = (e) => { if (e.target.closest('button')) return; el.classList.toggle('on'); }));
  const picked = () => [...ov.querySelectorAll('.sendimg.on')].map((el) => imgs[+el.dataset.sdi]);
  $('sdTxt').oninput = (e) => fitNote(e.target);
  $('sdX').onclick = () => (ov.innerHTML = '');
  $('sdReset').onclick = () => { $('sdTxt').value = txt; fitNote($('sdTxt')); };
  $('sdSave').onclick = () => {
    let body = $('sdTxt').value; if (c?.client) body = body.split(c.client.trim()).join('{客戶}'); body = body.split(fmtDay(localDay())).join('{日期}');
    n.body = body; Vault.put('notes', n); App.render(); toast('已存回範本');
  };
  $('sdText').onclick = async () => { try { await navigator.clipboard.writeText($('sdTxt').value); toast('文字已複製'); } catch (e) { toast('無法自動複製，請長按文字選取'); } };
  ov.querySelectorAll('[data-sdcopy]').forEach((b) => (b.onclick = async () => {
    try { await navigator.clipboard.write([new ClipboardItem({ 'image/png': pngBlob(imgs[+b.dataset.sdcopy]) })]); toast('圖片已複製，到 LINE 貼上'); }
    catch (e) { toast('這個瀏覽器不能複製圖片，請用「分享」或長按圖片儲存'); }
  }));
  if ($('sdShare')) $('sdShare').onclick = async () => {
    try {
      const sel = picked();
      const files = await Promise.all(sel.map(async (f, i) => new File([await pngBlob(f)], `${(n.title || '圖片').replace(/[\\/:*?"<>|]/g, '_')}-${i + 1}.png`, { type: 'image/png' })));
      const data = files.length ? { text: $('sdTxt').value, files } : { text: $('sdTxt').value };
      if (!navigator.canShare(data)) { toast('這台裝置不能一次分享圖片，請一張一張複製'); return; }
      await navigator.share(data);
    } catch (e) { if (e.name !== 'AbortError') toast('分享失敗，請改用複製'); }
  };
}
let GV = null;
async function showGallery(nav) {
  GV = nav; const d = nav.list[nav.i], n = nav.list.length;
  let url = null; try { url = nav.masked && !isPdf(d) ? (await renderDoc(d, true)).toDataURL('image/jpeg', 0.9) : await Files.url(d.fileId); } catch (e) { } // 理賠參考：看的是遮罩版
  if (GV !== nav) return; if (!url) { toast('檔案讀取失敗，請確認網路'); return; }
  const body = isPdf(d) ? `<iframe src="${url}" title="${esc(d.name)}" style="flex:1;width:100%;border:0;background:#fff;min-height:70vh"></iframe>` : `<img src="${url}" alt="${esc(d.name)}" class="gvimg">`;
  ov.innerHTML = `<div class="ov gv"><div class="ovhead"><h2>${esc(d.name)}</h2><span class="gvcnt">${nav.i + 1} / ${n}</span><button class="btn" id="gvPrev" ${nav.i ? '' : 'disabled'} aria-label="上一張">‹ 上一張</button><button class="btn" id="gvNext" ${nav.i < n - 1 ? '' : 'disabled'} aria-label="下一張">下一張 ›</button>${isPdf(d) ? '<button class="btn" id="vShare">分享／儲存</button>' : ''}<button class="btn pri" id="vClose">關閉</button></div>
   <div class="ovbody gvbody">${body}<button class="gvside l" id="gvL" ${nav.i ? '' : 'hidden'} aria-label="上一張">‹</button><button class="gvside r" id="gvR" ${nav.i < n - 1 ? '' : 'hidden'} aria-label="下一張">›</button></div></div>`;
  const go = (k) => { const j = nav.i + k; if (j < 0 || j >= n) return; showGallery({ ...nav, i: j }); };
  $('gvPrev').onclick = $('gvL').onclick = () => go(-1);
  $('gvNext').onclick = $('gvR').onclick = () => go(1);
  $('vClose').onclick = () => { ov.innerHTML = ''; GV = null; };
  if ($('vShare')) $('vShare').onclick = async () => { const blob = await (await fetch(url)).blob(); shareOrSave([new File([blob], d.name.replace(/[\\/:*?"<>|]/g, '_') + '.pdf', { type: 'application/pdf' })]); };
  const img = ov.querySelector('.gvimg');
  if (img) {
    let sx = null, sy = 0;
    img.addEventListener('touchstart', (e) => { if (e.touches.length === 1) { sx = e.touches[0].clientX; sy = e.touches[0].clientY; } else sx = null; }, { passive: true });
    img.addEventListener('touchend', (e) => { if (sx == null) return; const t = e.changedTouches[0], dx = t.clientX - sx, dy = t.clientY - sy; sx = null; if (Math.abs(dx) > 50 && Math.abs(dx) > Math.abs(dy)) { e.preventDefault(); go(dx < 0 ? 1 : -1); } });
    img.onclick = (e) => { const r = img.getBoundingClientRect(); go(e.clientX < r.left + r.width / 3 ? -1 : 1); }; // 點右邊／中間下一張，點左邊三分之一上一張
  }
}
document.addEventListener('keydown', (e) => { if (!GV || !ov.querySelector('.gv')) return; if (e.key === 'ArrowRight') $('gvNext')?.click(); else if (e.key === 'ArrowLeft') $('gvPrev')?.click(); else if (e.key === 'Escape') $('vClose')?.click(); });

/* ================= 遮罩 ================= */
/* 新契約清單：測試頁（不存任何資料，關掉就清空） */
const NC_PREP = [
  ['要問客戶的', ['身分證件', '身高體重', '工作', '法定繼承人（受益人）', '方便電話聯絡的時間']],
  ['我要準備', ['金三角', '建議書', '商品簡介、重點條款、DM', '保單健檢（保險存摺）', 'T字障', '好解說']],
  ['我要帶的', ['紙本或電子投保', '繳費：自動轉帳、信用卡、讀卡機、印泥', '文具：自動筆、原子筆、立可帶、釘書機', '其他文件：開通行動服務、A式變更（受益人）、產險投保', 'APP 更新、網路']],
];
const NC_FLOW = [
  ['個人', ['基本資料（★工作）', '投保紀錄、保險存摺、保單健檢']],
  ['生調表', ['★婚姻、服役、學歷', '菸酒檳榔安眠藥、身體健康', '★財務狀況', '貸款、解約']],
  ['健康告知', ['兩個月、一年、兩年、五年的健檢', '看診用藥、血壓血糖血脂、身高體重']],
  ['投資評估', ['外幣收付', '投資型風險適性問卷'], 'inv'],
  ['高齡', ['高齡量表：婚姻、學歷、健康、居住、投保經驗、投資經驗、可支應年數、生活變化、認知評估', '錄音：要保人、被保人、繳款人誰超過 65 歲就錄誰（傳統、外幣、投資型）'], 'old'],
  ['電訪提醒', ['投資型保費來源、有無解約']],
  ['收尾', ['Fitback', '電子投保確認書、保費授權、上傳成功', '開通行動服務、國泰人壽 APP、證券戶', '截圖存檔：身分證、存摺封面']],
];
const NC_LOG = [['認識', ['名單來源', '背景', '關係建立']], ['需求', ['缺口與觀念', '預算', '體況與核保']], ['方案', ['規劃', '促約關鍵', '結果']], ['後續', ['待辦', '下次', '一句話心得（早會用）']]];
/* 新契約、建議書案件裡的三塊：準備清單、流程提醒、過程紀錄。項目存在設定（所有案件共用），每件的打勾和文字存在案件 c.nc */
function ncDefault() { let i = 0; const mk = (arr) => arr.map(([g, l, k]) => ({ g, k: k || '', items: l.map((t) => ({ id: 'd' + (++i), t, on: true })) })); return { id: 'nctpl', prep: mk(NC_PREP), flow: mk(NC_FLOW), log: mk(NC_LOG) }; }
function ncTpl() { const t = Vault.get('settings', 'nctpl'); return t?.prep ? t : ncDefault(); }
const NC_SECS = [['prep', '一、準備清單', '見客戶前要準備、要帶、要問到的'], ['flow', '二、流程提醒', '簽約當下照順序走'], ['log', '三、過程紀錄', '早會分享、自己的業務紀錄']];
function ncCaseHTML(c) {
  const t = ncTpl(), st = c.nc || {}, chk = st.chk || {}, val = st.val || {};
  const prepDocs = (c.docs || []).filter((d) => d.folder === '準備資料').length;
  const total = (sk) => t[sk].reduce((a, g) => a + ((!g.k || st[g.k]) ? g.items.length : 0), 0), done = (sk) => t[sk].reduce((a, g) => a + ((!g.k || st[g.k]) ? g.items.filter((it) => chk[it.id]).length : 0), 0);
  let n = 0; const box = (it) => `<label class="ncit"><input type="checkbox" data-ncc="${it.id}"${chk[it.id] ? ' checked' : ''}><span class="ncno">${++n}</span><span>${esc(it.t)}</span></label>`;
  const grp = (g) => `<div class="ncgrp"${g.k && !st[g.k] ? ' hidden' : ''}><div class="ncg">${esc(g.g)}${g.k ? '<span class="ncif">' + (g.k === 'inv' ? '外幣或投資型才出現' : '有人超過 65 歲才出現') + '</span>' : ''}</div>${g.items.map(box).join('')}</div>`;
  return `<div class="sec nccase"><h2>這件有沒有： <button class="btn sm" data-act="nctpl">✎ 編修清單</button></h2><div class="tags"><button class="tag${st.inv ? ' on' : ''}" data-ncflag="inv">外幣或投資型</button><button class="tag${st.old ? ' on' : ''}" data-ncflag="old">有人超過 65 歲</button></div></div>
  ${NC_SECS.map(([sk, title, sub]) => `<details class="sec" ${sk === 'log' || done(sk) < total(sk) ? 'open' : ''}><summary><h2 style="display:inline">${title} <span class="sub">${sk === 'log' ? sub : `${sub}・完成 ${done(sk)}/${total(sk)}`}</span></h2></summary>
   ${sk === 'prep' ? `<div class="sub">建議書、保費明細表、商品簡介放進文件的「準備資料」資料夾${prepDocs ? `，目前有 ${prepDocs} 份 📎` : '，就不會忘記準備過什麼'}。</div>` : ''}
   ${sk === 'log' ? `<div class="ncgrid">${t.log.map((g) => `<div class="nclg"><div class="ncg">${esc(g.g)}</div>${g.items.filter((it) => it.on).map((it) => `<label class="f">${esc(it.t)}<textarea class="grow" rows="2" data-ncv="${it.id}">${esc(val[it.id] || '')}</textarea></label>`).join('')}</div>`).join('')}</div>`
    : `<div class="ncgrid${sk === 'prep' ? ' three' : ''}">${t[sk].map(grp).join('')}</div>`}</details>`).join('')}`;
}
function ncLogText(c) { // 匯出備註時，過程紀錄一起帶出來
  const t = ncTpl(), val = c.nc?.val || {}; const out = [];
  t.log.forEach((g) => g.items.forEach((it) => { const v = (val[it.id] || '').trim(); if (v) out.push(`【${it.t}】${v}`); }));
  return out.join('\n');
}
function openNcTplEditor() { // 編修清單：改字、刪除、新增、拖曳排序（改的是所有新契約、建議書共用的清單）
  const tpl = JSON.parse(JSON.stringify(ncTpl())); let nid = Date.now(); const v = document.createElement('div'); document.body.appendChild(v);
  const render = () => {
    v.innerHTML = `<div class="ov" style="z-index:30"><div class="ovhead"><h2>編修清單</h2><button class="btn" id="ntReset">恢復預設</button><button class="btn" id="ntX">取消</button><button class="btn pri" id="ntOk">儲存</button></div>
    <div class="ovbody" style="flex-direction:column;max-width:1100px;margin:0 auto;width:100%">
     <div class="note">這份清單所有新契約、建議書共用。可以改字、刪除（按兩次）、新增，按住 ⠿ 拖曳排序（也能拖到別的分類）。過程紀錄勾掉的項目不會出現。已經寫過的內容不會因為改名而消失。</div>
     ${NC_SECS.map(([sk, title]) => `<div class="sec"><h2>${title}</h2><div class="ncgrid${sk === 'prep' ? ' three' : ''}" data-sec="${sk}">${tpl[sk].map((g, gi) => `<div class="ncgrp" data-gi="${gi}">
       <div class="ncg"><span class="ncdrag ncgdrag">⠿</span><input class="nctxt ncgname" data-ntg="${sk}:${gi}" value="${esc(g.g)}"></div>
       <div class="ncbody">${g.items.map((it) => `<div class="ncrow${it.on ? '' : ' off'}" data-iid="${it.id}"><span class="ncdrag">⠿</span>${sk === 'log' ? `<input type="checkbox" class="ncon" data-nton="${it.id}"${it.on ? ' checked' : ''} aria-label="使用">` : ''}<input class="nctxt" data-ntt="${it.id}" value="${esc(it.t)}"><button class="btn sm ncdel" data-ntd="${it.id}">✕</button></div>`).join('')}</div>
       <button class="btn sm ncadd" data-nta="${sk}:${gi}">＋ 新增一項</button></div>`).join('')}</div></div>`).join('')}</div></div>`;
    const find = (id) => { for (const sk of ['prep', 'flow', 'log']) for (const g of tpl[sk]) { const it = g.items.find((x) => x.id === id); if (it) return [g, it]; } return []; };
    v.querySelector('#ntX').onclick = () => v.remove();
    v.querySelector('#ntOk').onclick = () => { Vault.put('settings', tpl); v.remove(); App.render(); toast('清單已儲存'); };
    v.querySelector('#ntReset').onclick = (e) => { const b = e.currentTarget; if (!b.dataset.sure) { b.dataset.sure = 1; b.textContent = '確定恢復預設？'; return; } const d = ncDefault(); tpl.prep = d.prep; tpl.flow = d.flow; tpl.log = d.log; render(); };
    v.querySelectorAll('[data-ntg]').forEach((x) => (x.oninput = () => { const [sk, gi] = x.dataset.ntg.split(':'); tpl[sk][+gi].g = x.value; }));
    v.querySelectorAll('[data-ntt]').forEach((x) => (x.oninput = () => { const [, it] = find(x.dataset.ntt); if (it) it.t = x.value; }));
    v.querySelectorAll('[data-nton]').forEach((x) => (x.onchange = () => { const [, it] = find(x.dataset.nton); if (it) it.on = x.checked; x.closest('.ncrow').classList.toggle('off', !x.checked); }));
    v.querySelectorAll('[data-ntd]').forEach((b) => (b.onclick = () => { if (!b.dataset.sure) { b.dataset.sure = 1; b.textContent = '刪？'; return; } const [g] = find(b.dataset.ntd); if (g) g.items = g.items.filter((x) => x.id !== b.dataset.ntd); render(); }));
    v.querySelectorAll('[data-nta]').forEach((b) => (b.onclick = () => { const [sk, gi] = b.dataset.nta.split(':'); const it = { id: 'u' + (++nid), t: '', on: true }; tpl[sk][+gi].items.push(it); render(); v.querySelector(`[data-ntt="${it.id}"]`)?.focus(); }));
    makeSortable(v, () => { for (const sk of ['prep', 'flow', 'log']) { const all = {}; tpl[sk].forEach((g) => g.items.forEach((it) => (all[it.id] = it))); const old = tpl[sk]; const box = v.querySelector(`[data-sec="${sk}"]`); tpl[sk] = [...box.querySelectorAll(':scope > .ncgrp')].map((ge) => ({ ...old[+ge.dataset.gi], items: [...ge.querySelectorAll('.ncrow')].map((r) => all[r.dataset.iid]).filter(Boolean) })); }
      // 項目可能跨區拖不過去（只限同一區），所以每區各自重建即可
      render(); });
  };
  render();
}
let NCT = null;
function ncInit() {
  let id = 0; const mk = (arr) => arr.map(([g, l, k]) => ({ g, k, items: l.map((t) => ({ id: 'n' + (++id), t, on: true })) }));
  return { inv: false, old: false, edit: false, chk: {}, val: {}, sec: { prep: mk(NC_PREP), flow: mk(NC_FLOW), log: mk(NC_LOG) }, nid: id };
}
function openNcTest() {
  NCT = NCT || ncInit();
  const keepScroll = ov.querySelector('.ovbody')?.scrollTop || 0;
  const ed = NCT.edit; let n = 0;
  const cond = (k) => k ? '<span class="ncif">' + (k === 'inv' ? '勾了外幣或投資型才出現' : '勾了超過 65 歲才出現') + '</span>' : '';
  const row = (sk, it) => {
    if (ed) return `<div class="ncrow${it.on ? '' : ' off'}" data-iid="${it.id}"><span class="ncdrag" aria-label="拖曳排序">⠿</span>${sk === 'log' ? `<input type="checkbox" class="ncon" data-on="${it.id}"${it.on ? ' checked' : ''} aria-label="使用">` : ''}<input class="nctxt" data-txt="${it.id}" value="${esc(it.t)}"><button class="btn sm ncdel" data-del="${it.id}" aria-label="刪除">✕</button></div>`;
    if (sk === 'log') return it.on ? `<label class="f">${esc(it.t)}<textarea rows="2" data-nval="${it.id}">${esc(NCT.val[it.id] || '')}</textarea></label>` : '';
    return `<label class="ncit"><input type="checkbox" data-nck2="${it.id}"${NCT.chk[it.id] ? ' checked' : ''}><span class="ncno">${++n}</span><span>${esc(it.t)}</span></label>`;
  };
  const grp = (sk, g, gi) => `<div class="ncgrp" data-gi="${gi}"${!ed && g.k ? ` data-nck="${g.k}"${NCT[g.k] ? '' : ' hidden'}` : ''}>
    <div class="ncg">${ed ? `<span class="ncdrag ncgdrag" aria-label="拖曳分類">⠿</span><input class="nctxt ncgname" data-gname="${sk}:${gi}" value="${esc(g.g)}">` : esc(g.g)}${cond(g.k)}</div>
    <div class="ncbody" data-sk="${sk}" data-gi="${gi}">${g.items.map((it) => row(sk, it)).join('')}</div>
    ${ed ? `<button class="btn sm ncadd" data-add="${sk}:${gi}">＋ 新增一項</button>` : ''}</div>`;
  const sec = (sk, title, sub) => `<div class="sec"><h2>${title} <span class="sub">${sub}</span></h2><div class="ncgrid${sk === 'prep' ? ' three' : ''}" data-sec="${sk}">${NCT.sec[sk].map((g, gi) => grp(sk, g, gi)).join('')}</div></div>`;
  ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>新契約清單 · 測試頁</h2><button class="btn${ed ? ' pri' : ''}" id="ncEd">${ed ? '✓ 改好了' : '✎ 編修模式'}</button><button class="btn" id="ncX">關閉</button></div>
  <div class="ovbody" style="flex-direction:column;max-width:1180px;margin:0 auto;width:100%">
   <div class="note">${ed ? '編修模式：可以改字、刪除（✕）、新增，按住 ⠿ 拖曳排順序（也可以拖到別的分類）。過程紀錄勾掉的項目不會出現。改好按右上角「改好了」。' : '這是測試頁，可以打勾、打字試用，<b>不會存任何資料</b>，關掉就清空。要改項目，按右上角「編修模式」。'}</div>
   ${ed ? '' : `<div class="sec"><h2>這件有沒有：</h2><div class="tags"><button class="tag${NCT.inv ? ' on' : ''}" data-nct="inv">外幣或投資型</button><button class="tag${NCT.old ? ' on' : ''}" data-nct="old">有人超過 65 歲</button></div></div>`}
   ${sec('prep', '一、準備清單', '見客戶前要準備、要帶、要問到的')}
   ${sec('flow', '二、流程提醒', '簽約當下照順序走')}
   ${sec('log', '三、過程紀錄', '早會分享、自己的業務紀錄')}
  </div></div>`;
  ov.querySelector('.ovbody').scrollTop = keepScroll;
  $('ncX').onclick = () => { NCT = null; ov.innerHTML = ''; };
  $('ncEd').onclick = () => { NCT.edit = !NCT.edit; openNcTest(); };
  const find = (id) => { for (const sk in NCT.sec) for (const g of NCT.sec[sk]) { const i = g.items.findIndex((x) => x.id === id); if (i >= 0) return [g, i]; } return [null, -1]; };
  ov.querySelectorAll('[data-nct]').forEach((b) => (b.onclick = () => { const k = b.dataset.nct; NCT[k] = !NCT[k]; b.classList.toggle('on', NCT[k]); ov.querySelectorAll(`[data-nck="${k}"]`).forEach((x) => (x.hidden = !NCT[k])); }));
  ov.querySelectorAll('[data-nck2]').forEach((c) => (c.onchange = () => { NCT.chk[c.dataset.nck2] = c.checked; }));
  ov.querySelectorAll('[data-nval]').forEach((t) => (t.oninput = () => { NCT.val[t.dataset.nval] = t.value; }));
  if (!ed) return;
  ov.querySelectorAll('[data-txt]').forEach((t) => (t.oninput = () => { const [g, i] = find(t.dataset.txt); if (g) g.items[i].t = t.value; }));
  ov.querySelectorAll('[data-gname]').forEach((t) => (t.oninput = () => { const [sk, gi] = t.dataset.gname.split(':'); NCT.sec[sk][+gi].g = t.value; }));
  ov.querySelectorAll('[data-on]').forEach((c) => (c.onchange = () => { const [g, i] = find(c.dataset.on); if (g) g.items[i].on = c.checked; c.closest('.ncrow').classList.toggle('off', !c.checked); }));
  ov.querySelectorAll('[data-del]').forEach((b) => (b.onclick = () => { if (!b.dataset.sure) { b.dataset.sure = 1; b.textContent = '刪？'; setTimeout(() => { if (b.isConnected) { delete b.dataset.sure; b.textContent = '✕'; } }, 3000); return; } const [g, i] = find(b.dataset.del); if (g) { g.items.splice(i, 1); openNcTest(); } }));
  ov.querySelectorAll('[data-add]').forEach((b) => (b.onclick = () => { const [sk, gi] = b.dataset.add.split(':'); const it = { id: 'n' + (++NCT.nid), t: '', on: true }; NCT.sec[sk][+gi].items.push(it); openNcTest(); ov.querySelector(`[data-txt="${it.id}"]`)?.focus(); }));
  // 拖曳排序：按住 ⠿ 拖；項目可以拖到同一區的別的分類，分類也可以拖
  const syncFromDom = () => {
    const all = {}; for (const sk in NCT.sec) NCT.sec[sk].forEach((g) => g.items.forEach((it) => (all[it.id] = it)));
    for (const sk in NCT.sec) {
      const old = NCT.sec[sk];
      NCT.sec[sk] = [...ov.querySelectorAll(`[data-sec="${sk}"] > .ncgrp`)].map((ge) => ({ ...old[+ge.dataset.gi], items: [...ge.querySelectorAll('.ncrow')].map((r) => all[r.dataset.iid]) }));
    }
  };
  ov.querySelectorAll('.ncdrag').forEach((h) => (h.onpointerdown = (e) => {
    e.preventDefault(); const isG = h.classList.contains('ncgdrag'); const el = h.closest(isG ? '.ncgrp' : '.ncrow'); const home = el.closest('[data-sec]');
    el.classList.add('ncdragging');
    const mv = (ev) => {
      const under = document.elementFromPoint(ev.clientX, ev.clientY); if (!under) return;
      if (isG) { const t = under.closest('.ncgrp'); if (!t || t === el || t.parentElement !== home) return; const r = t.getBoundingClientRect(); const after = ev.clientY > r.top + r.height / 2 || (ev.clientX > r.left + r.width / 2 && ev.clientY > r.top); t.parentElement.insertBefore(el, after ? t.nextSibling : t); return; }
      const t = under.closest('.ncrow'), body = under.closest('.ncbody');
      if (t && t !== el && t.closest('[data-sec]') === home) { const r = t.getBoundingClientRect(); t.parentElement.insertBefore(el, ev.clientY > r.top + r.height / 2 ? t.nextSibling : t); }
      else if (!t && body && body.closest('[data-sec]') === home && !body.contains(el)) body.appendChild(el);
    };
    const up = () => { window.removeEventListener('pointermove', mv); window.removeEventListener('pointerup', up); window.removeEventListener('pointercancel', up); el.classList.remove('ncdragging'); syncFromDom(); openNcTest(); };
    window.addEventListener('pointermove', mv); window.addEventListener('pointerup', up); window.addEventListener('pointercancel', up);
  }));
}
let MK = null;
async function openMask(c, i, coll = 'cases', opts = null) { const d = c.docs[i]; const im = await getImg(d.fileId); if (!im) { toast('圖片讀取失敗，請確認網路'); return; } MK = { c, coll, d, im, masks: d.masks.map((m) => ({ ...m })), marks: (d.marks || []).map((m) => ({ ...m, pts: m.pts ? m.pts.map((p) => [...p]) : undefined })), tool: opts?.tool || 'mask', color: MARK_COLORS[0], hist: [], sel: -1, mode: null, start: null, preview: false, opts }; renderMask(); }
// 遮罩畫面：復原（Ctrl+Z）、刪除鍵、Ctrl+滾輪／雙指縮放
const mkState = () => JSON.stringify({ masks: MK.masks, marks: MK.marks });
function mkCommit(pre) { if (pre && pre !== mkState()) { MK.hist.push(pre); if (MK.hist.length > 80) MK.hist.shift(); } }
function mkUndo() { const h = MK?.hist.pop(); if (!h) { toast('沒有可以復原的'); return; } const o = JSON.parse(h); MK.masks = o.masks; MK.marks = o.marks; MK.sel = -1; renderMask(); }
function mkDelSel() { if (MK.sel < 0) { toast('先點選一個框'); return; } const pre = mkState(); MK.masks.splice(MK.sel, 1); MK.sel = -1; mkCommit(pre); renderMask(); }
document.addEventListener('keydown', (e) => {
  if (!MK || !$('mkCv') || /^(TEXTAREA|INPUT)$/.test(e.target.tagName)) return;
  if ((e.ctrlKey || e.metaKey) && !e.shiftKey && e.key.toLowerCase() === 'z') { e.preventDefault(); mkUndo(); }
  else if ((e.key === 'Delete' || e.key === 'Backspace') && MK.sel >= 0) { e.preventDefault(); mkDelSel(); }
});
let mkWheelT = null;
document.addEventListener('wheel', (e) => {
  if (!MK || !e.ctrlKey || !e.target.closest?.('#ovRoot .stage')) return;
  e.preventDefault();
  const cv = $('mkCv'); MK.zoom = Math.max(1, Math.min(6, MK.zoom * Math.exp(-e.deltaY * 0.002)));
  if (MK.fit) cv.style.cssText += `;max-width:none;max-height:none;width:${MK.fit * MK.zoom}px`;
  clearTimeout(mkWheelT); mkWheelT = setTimeout(() => { if (MK) renderMask(); }, 250);
}, { passive: false });
function renderMask() {
  const st0 = document.querySelector('#ovRoot .stage'), keepScroll = st0 ? [st0.scrollLeft, st0.scrollTop] : null; // 重畫後留在原本放大的位置
  if (MK.note == null) MK.note = MK.d.note || ''; MK.zoom = MK.zoom || 1;
  ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>${MK.opts?.title || '標註'} · ${esc(MK.d.name)}</h2><button class="btn" id="mkCancel">取消</button><button class="btn pri" id="mkSave">儲存</button></div>
  <div class="ovbody"><div class="stage${MK.zoom > 1 ? ' zoomed' : ''}${MK.pan ? ' panning' : ''}"><canvas id="mkCv"></canvas></div>
  <div class="side"><div class="actions zoombar" style="margin:0"><button class="btn" id="mkZo" aria-label="縮小" ${MK.zoom > 1 ? '' : 'disabled'}>－</button><span class="sub">${Math.round(MK.zoom * 100)}%</span><button class="btn" id="mkZi" aria-label="放大" ${MK.zoom < 4 ? '' : 'disabled'}>＋</button>${MK.zoom > 1 ? `<button class="btn${MK.pan ? ' on' : ''}" id="mkPan">${MK.pan ? '✋ 移動中' : '✋ 移動畫面'}</button>` : ''}</div>
  <div class="mktools">${[['mask', '▭ 遮罩'], ['pen', '✏️ 畫筆'], ['circle', '◯ 圈選'], ['stamp', '🔖 印章'], ['erase', '⌫ 橡皮擦']].map(([k, l]) => `<button class="btn${MK.tool === k && !MK.pan ? ' on' : ''}" data-mktool="${k}">${l}</button>`).join('')}</div>
  ${MK.tool === 'pen' || MK.tool === 'circle' ? `<div class="mkcolors">${MARK_COLORS.map((c) => `<button class="mkcolor${MK.color === c ? ' on' : ''}" data-mkcolor="${c}" style="background:${c}" aria-label="顏色"></button>`).join('')}</div>` : ''}
  ${MK.tool === 'stamp' ? stampPanelHTML() : ''}<div class="note">${MK.tool === 'stamp' ? '點一個章，再點頁面要蓋的地方，可以連續蓋。點已經蓋好的章可以選取：拖曳移動、拉右下角改大小。' : MK.tool === 'mask' ? '在要遮的地方拖曳畫框。點一下框可以選取、拖曳移動。' : MK.tool === 'pen' ? '直接在圖上畫。' : MK.tool === 'circle' ? '拖曳畫出圈圈，把重點圈起來。' : '在畫過的記號上點或拖，就會擦掉。'}原圖不會被改動。${MK.zoom > 1 ? '放大後要移動畫面，按「✋ 移動畫面」再拖。' : ''}</div>
  <div class="actions" style="margin:0;align-items:center"><button class="btn" id="mkUndo" ${MK.hist.length ? '' : 'disabled'}>↶ 復原</button><span class="sub">Ctrl+Z 也可以${MK.marks.length ? `・記號 ${MK.marks.length} 個` : ''}</span></div>
  <label class="f">筆記<textarea id="mkNote" class="grow" rows="3" placeholder="這份文件的重點，例如：急診 2 天、住院日額可申請">${esc(MK.note)}</textarea></label>
  <div class="actions"><button class="btn${MK.preview ? ' on' : ''}" id="mkPrev">${MK.preview ? '回到編輯' : '看分享版效果'}</button></div>
  <div class="actions"><button class="btn danger" id="mkDel">刪除選取的框</button><button class="btn" id="mkClear">全部清除</button></div>
  <div class="sub">目前 ${MK.masks.length} 個框</div></div></div></div>`;
  $('mkCancel').onclick = () => { const o = MK.opts; ov.innerHTML = ''; MK = null; o?.onCancel?.(); };
  $('mkNote').oninput = (e) => { MK.note = e.target.value; fitNote(e.target); }; fitNote($('mkNote'));
  $('mkZi').onclick = () => { MK.zoom = Math.min(4, MK.zoom * 1.5); renderMask(); };
  $('mkZo').onclick = () => { MK.zoom = Math.max(1, MK.zoom / 1.5); if (MK.zoom < 1.05) { MK.zoom = 1; MK.pan = false; } renderMask(); };
  if ($('mkPan')) $('mkPan').onclick = () => { MK.pan = !MK.pan; renderMask(); };
  ov.querySelectorAll('[data-mktool]').forEach((b) => (b.onclick = () => { MK.tool = b.dataset.mktool; MK.pan = false; MK.sel = -1; MK.ssel = -1; renderMask(); }));
  ov.querySelectorAll('[data-mkcolor]').forEach((b) => (b.onclick = () => { MK.color = b.dataset.mkcolor; renderMask(); }));
  $('mkUndo').onclick = mkUndo;
  $('mkSave').onclick = () => { stampSaveRecent(); if (MK.opts?.onSave) { const o = MK.opts, r = { masks: MK.masks, marks: MK.marks, note: MK.note.trim() }; ov.innerHTML = ''; MK = null; o.onSave(r); return; } MK.d.masks = MK.masks; MK.d.marks = MK.marks; MK.d.note = MK.note.trim(); Vault.put(MK.coll, MK.c); ov.innerHTML = ''; MK = null; App.render(); toast('已儲存，原圖保留'); };
  $('mkPrev').onclick = () => { MK.preview = !MK.preview; renderMask(); };
  $('mkDel').onclick = mkDelSel;
  $('mkClear').onclick = () => { const pre = mkState(); MK.masks = []; MK.sel = -1; mkCommit(pre); renderMask(); };
  const cv = $('mkCv'); cv.width = MK.im.width; cv.height = MK.im.height; const x = cv.getContext('2d'); const W = cv.width, H = cv.height;
  const draw = (tmp) => { x.drawImage(MK.im, 0, 0); const all = tmp ? [...MK.masks, tmp] : MK.masks; all.forEach((m, i) => { if (MK.preview) { x.fillStyle = '#111'; x.fillRect(m.x * W, m.y * H, m.w * W, m.h * H); } else { x.fillStyle = 'rgba(17,17,17,.6)'; x.fillRect(m.x * W, m.y * H, m.w * W, m.h * H); x.lineWidth = Math.max(2, W / 300); x.strokeStyle = i === MK.sel ? '#E2B56A' : '#9DBA88'; x.strokeRect(m.x * W, m.y * H, m.w * W, m.h * H); } }); drawMarks(x, MK.cur ? [...MK.marks, MK.cur] : MK.marks, 0, 0, W, H); const sm = MK.marks[MK.ssel]; if (MK.tool === 'stamp' && sm?.t === 'stamp' && !MK.preview) { x.save(); x.setLineDash([W / 120, W / 160]); x.lineWidth = Math.max(2, W / 400); x.strokeStyle = '#1E6FD9'; const pd = W / 200; x.strokeRect(sm.x * W - pd, sm.y * H - pd, sm.w * W + pd * 2, sm.h * H + pd * 2); x.setLineDash([]); x.fillStyle = '#1E6FD9'; const hs = Math.max(10, W / 70); x.fillRect((sm.x + sm.w) * W + pd - hs / 2, (sm.y + sm.h) * H + pd - hs / 2, hs, hs); x.restore(); } }; MK.redraw = draw;
  draw();
  MK.fit = cv.getBoundingClientRect().width;
  if (MK.zoom > 1) { cv.style.maxWidth = 'none'; cv.style.maxHeight = 'none'; cv.style.width = MK.fit * MK.zoom + 'px'; }
  // 雙指縮放（iPad）：第二根手指放上來就取消正在畫的那筆，改成縮放
  const ptrs = new Map(); let pinch = null;
  const dist = () => { const [a, b] = [...ptrs.values()]; return Math.hypot(a[0] - b[0], a[1] - b[1]); };
  cv.addEventListener('pointerdown', (e) => {
    ptrs.set(e.pointerId, [e.clientX, e.clientY]);
    if (ptrs.size === 2) { e.stopImmediatePropagation(); if (MK.pre) { const o = JSON.parse(MK.pre); MK.masks = o.masks; MK.marks = o.marks; } MK.mode = null; MK.cur = null; MK.pre = null; pinch = { d0: dist(), z0: MK.zoom }; draw(); }
  }, true);
  cv.addEventListener('pointermove', (e) => {
    if (!ptrs.has(e.pointerId)) return; ptrs.set(e.pointerId, [e.clientX, e.clientY]);
    if (pinch && ptrs.size >= 2) { e.stopImmediatePropagation(); MK.zoom = Math.max(1, Math.min(6, pinch.z0 * dist() / pinch.d0)); cv.style.maxWidth = 'none'; cv.style.maxHeight = 'none'; cv.style.width = MK.fit * MK.zoom + 'px'; }
  }, true);
  const up = (e) => { ptrs.delete(e.pointerId); if (pinch) { e.stopImmediatePropagation(); if (!ptrs.size) { pinch = null; if (MK.zoom < 1.05) MK.zoom = 1; renderMask(); } } };
  cv.addEventListener('pointerup', up, true); cv.addEventListener('pointercancel', up, true);
  if (keepScroll) { const st = ov.querySelector('.stage'); st.scrollLeft = keepScroll[0]; st.scrollTop = keepScroll[1]; }
  if (MK.preview || MK.pan) return;
  const pt = (e) => { const r = cv.getBoundingClientRect(); return [(e.clientX - r.left) / r.width, (e.clientY - r.top) / r.height]; };
  let tmp = null;
  if (MK.tool === 'stamp') { wireStamps(cv, pt, W, H, draw); return; }
  if (MK.tool !== 'mask') {
    const eraseAt = (px, py) => { const n = MK.marks.length; MK.marks = MK.marks.filter((m) => { if (m.t === 'stamp') return !(px >= m.x - 0.01 && px <= m.x + m.w + 0.01 && py >= m.y - 0.01 && py <= m.y + m.h + 0.01); if (m.t === 'pen') return !m.pts.some(([a, b]) => Math.hypot(a - px, b - py) < 0.025); const rx = m.w / 2, ry = m.h / 2, v = ((px - m.x - rx) / (rx || 1e-3)) ** 2 + ((py - m.y - ry) / (ry || 1e-3)) ** 2; return !(v > 0.55 && v < 1.6); }); const k = MK.masks.length; MK.masks = MK.masks.filter((m) => !(px >= m.x && px <= m.x + m.w && py >= m.y && py <= m.y + m.h)); if (MK.marks.length !== n || MK.masks.length !== k) { MK.sel = -1; draw(); } };
    cv.onpointerdown = (e) => { const [px, py] = pt(e); cv.setPointerCapture(e.pointerId); MK.pre = mkState(); MK.mode = MK.tool; if (MK.tool === 'pen') MK.cur = { t: 'pen', c: MK.color, pts: [[px, py]] }; else if (MK.tool === 'circle') { MK.cur = { t: 'circle', c: MK.color, x: px, y: py, w: 0, h: 0 }; MK.start = [px, py]; } else eraseAt(px, py); draw(); };
    cv.onpointermove = (e) => { if (!MK.mode) return; const [px, py] = pt(e); if (MK.mode === 'pen') { MK.cur.pts.push([+px.toFixed(4), +py.toFixed(4)]); draw(); } else if (MK.mode === 'circle') { const [sx, sy] = MK.start; Object.assign(MK.cur, { x: Math.min(sx, px), y: Math.min(sy, py), w: Math.abs(px - sx), h: Math.abs(py - sy) }); draw(); } else eraseAt(px, py); };
    cv.onpointerup = cv.onpointercancel = () => { if (!MK.mode) return; if (MK.cur && (MK.cur.t === 'pen' || MK.cur.w > 0.01 || MK.cur.h > 0.01)) MK.marks.push(MK.cur); MK.cur = null; MK.mode = null; mkCommit(MK.pre); MK.pre = null; renderMask(); };
    return;
  }
  cv.onpointerdown = (e) => { const [px, py] = pt(e); cv.setPointerCapture(e.pointerId); MK.pre = mkState(); const hit = MK.masks.findIndex((m) => px >= m.x && px <= m.x + m.w && py >= m.y && py <= m.y + m.h); if (hit >= 0) { MK.sel = hit; MK.mode = 'move'; MK.start = [px, py, { ...MK.masks[hit] }]; } else { MK.sel = -1; MK.mode = 'draw'; MK.start = [px, py]; } draw(); };
  cv.onpointermove = (e) => { if (!MK.mode) return; const [px, py] = pt(e); if (MK.mode === 'move') { const [sx, sy, o] = MK.start; const m = MK.masks[MK.sel]; m.x = Math.max(0, Math.min(1 - m.w, o.x + px - sx)); m.y = Math.max(0, Math.min(1 - m.h, o.y + py - sy)); draw(); } else { const [sx, sy] = MK.start; tmp = { x: Math.min(sx, px), y: Math.min(sy, py), w: Math.abs(px - sx), h: Math.abs(py - sy) }; draw(tmp); } };
  cv.onpointerup = cv.onpointercancel = () => { if (!MK.mode) return; if (MK.mode === 'draw' && tmp && tmp.w > 0.01 && tmp.h > 0.008) { MK.masks.push(tmp); MK.sel = MK.masks.length - 1; } MK.mode = null; tmp = null; mkCommit(MK.pre); MK.pre = null; renderMask(); };
}

function stampPanelHTML() {
  const lib = stampLib(); MK.lib = lib; const cur = MK.stampPick;
  const chip = (it, key) => `<button class="stbtn${cur && cur.key === key ? ' on' : ''}" data-stk="${key}">${stampChip(it)}</button>`;
  const recent = [...(MK.recent || []), ...(lib.recent || [])].filter((it, i, a) => a.findIndex((y) => y.k === it.k && y.s === it.s && y.c === it.c) === i).slice(0, 8);
  MK.recentShown = recent;
  const sm = MK.marks[MK.ssel]; const sel = sm?.t === 'stamp' ? sm : null;
  return `<div class="stpal">${recent.length ? `<div class="stcat">最近用過</div><div class="stchips">${recent.map((it, i) => chip(it, 'r' + i)).join('')}</div>` : ''}
   ${lib.groups.map((g, gi) => `<div class="stcat">${esc(g.g)}</div><div class="stchips">${g.items.map((it, ii) => chip(it, gi + ':' + ii)).join('')}</div>`).join('')}
   <button class="btn sm" id="stEdit" style="margin-top:8px">✎ 編修印章</button></div>
   ${sel ? `<div class="stprop"><div class="stcat">選到的印章</div>
    <div class="stchips">${STAMP_COLORS.map((c) => `<button class="mkcolor${sel.c === c ? ' on' : ''}" data-stc="${c}" style="background:${c};width:26px;height:26px" aria-label="顏色"></button>`).join('')}</div>
    <label class="f">大小<input type="range" id="stSize" min="20" max="300" value="100"></label>
    <label class="f">透明度<input type="range" id="stAlpha" min="10" max="100" value="${Math.round((sel.a ?? 1) * 100)}"></label>
    <div class="actions" style="margin:0"><button class="btn sm" id="stDup">複製一個</button><button class="btn sm danger" id="stDel">刪除這個章</button></div></div>` : ''}`;
}
function stampSaveRecent() {
  if (!MK?.recent?.length) return; const lib = stampLib();
  lib.recent = [...MK.recent, ...(lib.recent || [])].filter((it, i, a) => a.findIndex((y) => y.k === it.k && y.s === it.s && y.c === it.c) === i).slice(0, 8);
  Vault.put('settings', lib); MK.recent = [];
}
function wireStamps(cv, pt, W, H, draw) {
  const lib = MK.lib || stampLib();
  const pick = (key) => key[0] === 'r' ? MK.recentShown[+key.slice(1)] : (() => { const [gi, ii] = key.split(':').map(Number); return lib.groups[gi].items[ii]; })();
  ov.querySelectorAll('[data-stk]').forEach((b) => (b.onclick = () => { const key = b.dataset.stk; MK.stampPick = MK.stampPick?.key === key ? null : { key, ...pick(key) }; MK.ssel = -1; renderMask(); }));
  if ($('stEdit')) $('stEdit').onclick = () => openStampEditor(() => { MK.lib = null; MK.stampPick = null; renderMask(); });
  const sel = () => { const m = MK.marks[MK.ssel]; return m?.t === 'stamp' ? m : null; };
  ov.querySelectorAll('[data-stc]').forEach((b) => (b.onclick = () => { const m = sel(); if (!m) return; const pre = mkState(); m.c = b.dataset.stc; mkCommit(pre); renderMask(); }));
  if ($('stSize')) { let base = null, pre = null; $('stSize').oninput = (e) => { const m = sel(); if (!m) return; if (!base) { base = { w: m.w, h: m.h }; pre = mkState(); } const f = +e.target.value / 100; m.h = base.h * f; m.w = base.w * f; if (['text', 'uline', 'dot'].includes(m.k)) stampFit(m, W, H); draw(); }; $('stSize').onchange = () => { mkCommit(pre); base = null; renderMask(); }; }
  if ($('stAlpha')) { let pre = null; $('stAlpha').oninput = (e) => { const m = sel(); if (!m) return; pre = pre || mkState(); m.a = +e.target.value / 100; draw(); }; $('stAlpha').onchange = () => { mkCommit(pre); pre = null; }; }
  if ($('stDel')) $('stDel').onclick = () => { const pre = mkState(); MK.marks.splice(MK.ssel, 1); MK.ssel = -1; mkCommit(pre); renderMask(); };
  if ($('stDup')) $('stDup').onclick = () => { const m = sel(); if (!m) return; const pre = mkState(); MK.marks.push({ ...m, x: Math.min(1 - m.w, m.x + 0.02), y: Math.min(1 - m.h, m.y + 0.02) }); MK.ssel = MK.marks.length - 1; mkCommit(pre); renderMask(); };
  const hit = (px, py) => { for (let i = MK.marks.length - 1; i >= 0; i--) { const m = MK.marks[i]; if (m.t !== 'stamp') continue; const pd = 0.012; if (px >= m.x - pd && px <= m.x + m.w + pd && py >= m.y - pd && py <= m.y + m.h + pd) return i; } return -1; };
  let mode = null, start = null;
  cv.onpointerdown = (e) => {
    const [px, py] = pt(e); cv.setPointerCapture(e.pointerId); MK.pre = mkState();
    const s = sel(); const hr = 0.03;
    if (s && Math.abs(px - (s.x + s.w)) < hr && Math.abs(py - (s.y + s.h)) < hr) { mode = 'size'; start = { px, py, w: s.w, h: s.h }; return; }
    const i = hit(px, py);
    if (i >= 0) { MK.ssel = i; const m = MK.marks[i]; mode = 'move'; start = { px, py, x: m.x, y: m.y }; renderMaskSide(); draw(); return; }
    if (MK.stampPick) {
      const it = MK.stampPick; const k = it.k;
      const m = { t: 'stamp', k, s: it.s, c: it.c, a: it.a ?? stampAlpha(k), x: px, y: py, w: 0.3, h: 0.032 };
      if (k === 'hl') { m.w = 0.3; m.h = 0.028; } else if (k === 'brace') { m.w = 0.03; m.h = 0.14; } else if (k === 'line') { m.w = 0.2; m.h = 0.006; } else if (k === 'tag') { m.w = 0; m.h = 0.032; }
      stampFit(m, W, H); m.x = Math.max(0, Math.min(1 - m.w, px - (k === 'hl' || k === 'line' ? 0 : m.w / 2))); m.y = Math.max(0, Math.min(1 - m.h, py - m.h / 2));
      MK.marks.push(m); MK.ssel = MK.marks.length - 1; MK.recent = [{ k: it.k, s: it.s, c: it.c, a: it.a }, ...(MK.recent || [])].slice(0, 8);
      mode = 'move'; start = { px, py, x: m.x, y: m.y }; draw(); return;
    }
    MK.ssel = -1; mode = null; renderMaskSide(); draw();
  };
  cv.onpointermove = (e) => {
    if (!mode) return; const [px, py] = pt(e); const m = sel(); if (!m) return;
    if (mode === 'move') { m.x = Math.max(-m.w * 0.5, Math.min(1 - m.w * 0.5, start.x + px - start.px)); m.y = Math.max(-m.h * 0.5, Math.min(1 - m.h * 0.5, start.y + py - start.py)); }
    else { const nw = Math.max(0.01, start.w + px - start.px), nh = Math.max(0.004, start.h + py - start.py);
      if (['text', 'uline', 'dot'].includes(m.k)) { m.h = nh; stampFit(m, W, H); } else if (m.k === 'tag') { m.w = nw; m.h = nh; stampFit(m, W, H); } else { m.w = nw; m.h = nh; } }
    draw();
  };
  cv.onpointerup = cv.onpointercancel = () => { if (!mode) return; mode = null; mkCommit(MK.pre); MK.pre = null; renderMaskSide(); };
}
function renderMaskSide() { renderMask(); } // 側欄跟著選取更新
function openStampEditor(done) { // 編修印章：改字、換顏色、換樣式、刪除、新增、拖曳排序
  const lib = JSON.parse(JSON.stringify(stampLib())); let nid = Date.now();
  const v = document.createElement('div'); document.body.appendChild(v);
  const render = () => {
    v.innerHTML = `<div class="ov" style="z-index:30"><div class="ovhead"><h2>編修印章</h2><button class="btn" id="seReset">恢復預設</button><button class="btn" id="seX">取消</button><button class="btn pri" id="seOk">儲存</button></div>
    <div class="ovbody" style="flex-direction:column;max-width:860px;margin:0 auto;width:100%">
     <div class="note">按住 ⠿ 拖曳排順序（也可以拖到別的分類）。刪除要按兩次。</div>
     <div data-sec="st">${lib.groups.map((g, gi) => `<div class="ncgrp sec" data-gi="${gi}" style="margin-bottom:12px">
       <div class="ncg"><span class="ncdrag ncgdrag">⠿</span><input class="nctxt ncgname" data-seg="${gi}" value="${esc(g.g)}"><button class="btn sm ncdel" data-segdel="${gi}">刪除分類</button></div>
       <div class="ncbody" data-gi="${gi}">${g.items.map((it) => `<div class="ncrow serow" data-iid="${it.id}"><span class="ncdrag">⠿</span><span class="seprev">${stampChip(it)}</span>
         <select data-sek="${it.id}">${STAMP_KINDS.map(([k, l]) => `<option value="${k}"${it.k === k ? ' selected' : ''}>${l}</option>`).join('')}</select>
         <input class="nctxt" data-ses="${it.id}" value="${esc(it.s)}" placeholder="${['hl', 'brace', 'line'].includes(it.k) ? '不用寫字' : '印章文字'}"${['hl', 'brace', 'line'].includes(it.k) ? ' disabled' : ''}>
         <span class="secols">${STAMP_COLORS.map((c) => `<button class="mkcolor${it.c === c ? ' on' : ''}" data-sec2="${it.id}" data-c="${c}" style="background:${c};width:22px;height:22px;border-width:2px" aria-label="顏色"></button>`).join('')}</span>
         <button class="btn sm ncdel" data-sedel="${it.id}">✕</button></div>`).join('')}</div>
       <button class="btn sm ncadd" data-seadd="${gi}">＋ 新增印章</button></div>`).join('')}</div>
     <button class="btn" id="seAddG">＋ 新增分類</button></div></div>`;
    const find = (id) => { for (const g of lib.groups) { const it = g.items.find((x) => x.id === id); if (it) return it; } };
    v.querySelector('#seX').onclick = () => v.remove();
    v.querySelector('#seOk').onclick = () => { Vault.put('settings', lib); v.remove(); toast('印章已儲存'); done?.(); };
    v.querySelector('#seReset').onclick = (e) => { const b = e.currentTarget; if (!b.dataset.sure) { b.dataset.sure = 1; b.textContent = '確定恢復？'; return; } lib.groups = stampDefaultLib().groups; render(); };
    v.querySelector('#seAddG').onclick = () => { lib.groups.push({ g: '新分類', items: [] }); render(); };
    v.querySelectorAll('[data-seg]').forEach((t) => (t.oninput = () => (lib.groups[+t.dataset.seg].g = t.value)));
    v.querySelectorAll('[data-segdel]').forEach((b) => (b.onclick = () => { if (!b.dataset.sure) { b.dataset.sure = 1; b.textContent = '確定刪除整個分類？'; return; } lib.groups.splice(+b.dataset.segdel, 1); render(); }));
    v.querySelectorAll('[data-sek]').forEach((t) => (t.onchange = () => { const it = find(t.dataset.sek); it.k = t.value; it.a = stampAlpha(it.k); render(); }));
    v.querySelectorAll('[data-ses]').forEach((t) => (t.oninput = () => { const it = find(t.dataset.ses); it.s = t.value; t.closest('.serow').querySelector('.seprev').innerHTML = stampChip(it); }));
    v.querySelectorAll('[data-sec2]').forEach((b) => (b.onclick = () => { find(b.dataset.sec2).c = b.dataset.c; render(); }));
    v.querySelectorAll('[data-sedel]').forEach((b) => (b.onclick = () => { if (!b.dataset.sure) { b.dataset.sure = 1; b.textContent = '刪？'; return; } for (const g of lib.groups) g.items = g.items.filter((x) => x.id !== b.dataset.sedel); render(); }));
    v.querySelectorAll('[data-seadd]').forEach((b) => (b.onclick = () => { lib.groups[+b.dataset.seadd].items.push({ id: 'st' + (++nid), k: 'text', s: '新印章', c: '#185FA5', a: 1 }); render(); }));
    makeSortable(v, () => { const all = {}; lib.groups.forEach((g) => g.items.forEach((it) => (all[it.id] = it))); lib.groups = [...v.querySelectorAll('[data-sec="st"] > .ncgrp')].map((ge) => ({ ...lib.groups[+ge.dataset.gi], items: [...ge.querySelectorAll('.ncrow')].map((r) => all[r.dataset.iid]) })); render(); });
  };
  render();
}
function makeSortable(root, onDone) { // 按住 ⠿ 拖曳：項目可換分類，分類也可以拖
  root.querySelectorAll('.ncdrag').forEach((h) => (h.onpointerdown = (e) => {
    e.preventDefault(); const isG = h.classList.contains('ncgdrag'); const el = h.closest(isG ? '.ncgrp' : '.ncrow'); const home = el.closest('[data-sec]');
    el.classList.add('ncdragging');
    const mv = (ev) => {
      const under = document.elementFromPoint(ev.clientX, ev.clientY); if (!under) return;
      if (isG) { const t = under.closest('.ncgrp'); if (!t || t === el || t.parentElement !== el.parentElement) return; const r = t.getBoundingClientRect(); t.parentElement.insertBefore(el, ev.clientY > r.top + r.height / 2 ? t.nextSibling : t); return; }
      const t = under.closest('.ncrow'), body = under.closest('.ncbody');
      if (t && t !== el && t.closest('[data-sec]') === home) { const r = t.getBoundingClientRect(); t.parentElement.insertBefore(el, ev.clientY > r.top + r.height / 2 ? t.nextSibling : t); }
      else if (!t && body && body.closest('[data-sec]') === home && !body.contains(el)) body.appendChild(el);
    };
    const up = () => { window.removeEventListener('pointermove', mv); window.removeEventListener('pointerup', up); window.removeEventListener('pointercancel', up); el.classList.remove('ncdragging'); onDone(); };
    window.addEventListener('pointermove', mv); window.addEventListener('pointerup', up); window.addEventListener('pointercancel', up);
  }));
}

/* ---------- 選頁做成 PDF／列印 ---------- */
function openPrintPick(docs, name) {
  const imgs = (docs || []).map((d, i) => [d, i]).filter(([d]) => !isPdf(d)); if (!imgs.length) { toast('沒有可以合成的圖片'); return; }
  const pick = new Set(imgs.map(([, i]) => i)); let mask = false;
  const render = () => {
    ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>選要印的頁</h2><span class="sub">已選 ${pick.size} 頁</span><button class="btn" id="ppAll">${pick.size === imgs.length ? '全部不選' : '全選'}</button><button class="btn" id="ppX">取消</button><button class="btn pri" id="ppGo">做成 PDF（可列印）</button></div>
    <div class="ovbody" style="flex-direction:column"><div class="note">點頁面勾選或取消。印章、圈選、畫筆都會一起印出來。做好後會跳出分享選單，選「列印」就能印，也可以傳 LINE。</div>
     <label class="ncit" style="font-size:14px"><input type="checkbox" id="ppMask"${mask ? ' checked' : ''}> 遮罩也印上去（給客戶以外的人看時才需要）</label>
     <div class="pppages">${imgs.map(([d, i]) => `<button class="pppage${pick.has(i) ? ' on' : ''}" data-pp="${i}"><canvas data-ppc="${i}" width="220" height="300"></canvas><span>${pick.has(i) ? '✓ ' : ''}${esc(d.name)}</span></button>`).join('')}</div></div></div>`;
    ov.querySelectorAll('[data-ppc]').forEach((cv) => drawThumbTo(cv, docs[+cv.dataset.ppc], false).catch(() => {}));
    ov.querySelectorAll('[data-pp]').forEach((b) => (b.onclick = () => { const i = +b.dataset.pp; pick.has(i) ? pick.delete(i) : pick.add(i); render(); }));
    $('ppAll').onclick = () => { if (pick.size === imgs.length) pick.clear(); else imgs.forEach(([, i]) => pick.add(i)); render(); };
    $('ppMask').onchange = (e) => (mask = e.target.checked);
    $('ppX').onclick = () => { ov.innerHTML = ''; };
    $('ppGo').onclick = async () => { if (!pick.size) { toast('先勾要印的頁'); return; } const sel = imgs.filter(([, i]) => pick.has(i)).map(([d]) => d); ov.innerHTML = ''; await makePDF(sel, mask, name); };
  };
  render();
}

/* ================= 分享版 / PDF ================= */
async function renderDoc(d, mask) { // 回傳 canvas
  const im = await getImg(d.fileId); if (!im) throw new Error('圖片讀取失敗');
  const cv = document.createElement('canvas'); cv.width = im.width; cv.height = im.height; const x = cv.getContext('2d'); x.drawImage(im, 0, 0);
  if (mask) { x.fillStyle = '#111'; d.masks.forEach((m) => x.fillRect(m.x * cv.width, m.y * cv.height, m.w * cv.width, m.h * cv.height)); }
  drawMarks(x, d.marks, 0, 0, cv.width, cv.height);
  return cv;
}
const toBlob = (cv, type = 'image/jpeg') => new Promise((ok) => cv.toBlob((b) => ok(b), type, 0.85));
async function shareOrSave(files) { // 有分享功能就跳分享（可以直接傳 LINE），沒有就下載
  try { if (navigator.canShare && navigator.canShare({ files })) { await navigator.share({ files }); return; } } catch (e) { if (e.name === 'AbortError') return; }
  files.forEach((f) => { const a = document.createElement('a'); a.href = URL.createObjectURL(f); a.download = f.name; document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 60000); });
}
async function openShare(c, pdfName, only) {
  if (!c) return;
  const docs = (only || c.docs).filter((d) => !isPdf(d)); if (!docs.length) { toast('PDF 檔請在文件上按「看」再分享'); return; }
  let canvases; try { canvases = await Promise.all(docs.map((d) => renderDoc(d, true))); } catch (e) { toast('圖片讀取失敗，請確認網路'); return; }
  const unmasked = docs.filter((d) => !d.masks.length).length;
  const urls = canvases.map((cv) => cv.toDataURL('image/jpeg', 0.85));
  ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>分享版</h2><button class="btn" id="shPdf">分享 PDF</button><button class="btn" id="shAll">分享全部圖片</button><button class="btn pri" id="shClose">關閉</button></div>
  <div class="ovbody" style="flex-direction:column;align-items:center">${unmasked ? `<div class="note" style="color:var(--warn);background:var(--warn-soft)">有 ${unmasked} 張文件還沒有遮罩，傳出去前請先檢查。</div>` : ''}
  <div class="note">遮罩已經直接畫進這些圖片裡，按「分享」可以直接傳到 LINE。案件裡的原圖不受影響。</div>
  <div class="shareimgs">${urls.map((s) => `<img src="${s}" alt="遮罩後的文件">`).join('')}</div></div></div>`;
  $('shClose').onclick = () => (ov.innerHTML = '');
  $('shAll').onclick = async () => { const files = await Promise.all(canvases.map(async (cv, i) => new File([await toBlob(cv)], `分享版-${docs[i].name}-${i + 1}.jpg`, { type: 'image/jpeg' }))); shareOrSave(files); };
  $('shPdf').onclick = () => makePDF(docs, true, pdfName || '理賠參考-' + refTitle(c));
}
async function makePDF(docs, mask, name) {
  if (!window.jspdf) { toast('PDF 元件沒有載入，請重新整理'); return; }
  toast('正在產生 PDF…');
  const { jsPDF } = window.jspdf; let pdf = null;
  try {
    for (const d of docs.filter((x) => !isPdf(x))) {
      const cv = await renderDoc(d, mask); const o = cv.width > cv.height ? 'l' : 'p'; const pw = o === 'p' ? 210 : 297, ph = o === 'p' ? 297 : 210;
      if (!pdf) pdf = new jsPDF({ orientation: o, unit: 'mm', format: 'a4' }); else pdf.addPage('a4', o);
      const s = Math.min(pw / cv.width, ph / cv.height); const w = cv.width * s, h = cv.height * s;
      pdf.addImage(cv.toDataURL('image/jpeg', 0.85), 'JPEG', (pw - w) / 2, (ph - h) / 2, w, h);
    }
  } catch (e) { toast('圖片讀取失敗，請確認網路'); return; }
  if (!pdf) return;
  shareOrSave([new File([pdf.output('blob')], name.replace(/[\\/:*?"<>|]/g, '_') + '.pdf', { type: 'application/pdf' })]);
}

/* ================= 登入／密碼 ================= */
const gate = $('gate');
function showGate(html) { app.hidden = true; gate.hidden = false; gate.innerHTML = `<div class="gate-card"><h1><i></i>工作手冊</h1>${html}</div>`; }
function gateLoading(msg) { showGate(`<p>${esc(msg)}</p>`); }
/* ================= 自動備份（每天一份 data.enc，保留 30 天） ================= */
const KEEP_DAYS = 30;
const Backup = {
  async daily() {
    const raw = Vault.raw; Vault.raw = null; if (!raw) return; // 只用剛從雲端讀到的版本（離線開啟就不備份）
    try {
      const name = 'backup/' + localDay() + '.enc';
      if (await Drive.getMeta(name)) return;
      await Drive.upload(name, raw.bytes);
      await this.prune();
    } catch (e) { console.warn('備份失敗', e); }
  },
  async list() { return (await Drive.list('backup')).filter((n) => /^\d{4}-\d{2}-\d{2}.*\.enc$/.test(n)).sort().reverse(); },
  async prune() {
    const all = await this.list(), daily = all.filter((n) => /^\d{4}-\d{2}-\d{2}\.enc$/.test(n)), before = all.filter((n) => !/^\d{4}-\d{2}-\d{2}\.enc$/.test(n));
    for (const n of [...daily.slice(KEEP_DAYS), ...before.slice(5)]) await Drive.remove('backup/' + n).catch(() => { });
  },
  async read(name) {
    const r = await Drive.download('backup/' + name); if (!r) throw new Error('找不到這份備份');
    const env = JSON.parse(dec.decode(r.bytes));
    if (env.salt !== b64(Vault.salt)) throw new Error('這份備份用的是不同的密碼，無法直接還原');
    return JSON.parse(dec.decode(await unseal(Vault.key, unb64(env.data))));
  },
  async restore(name) {
    const old = { ...emptyDb(), ...(await this.read(name)) };
    // 還原前先把現在的資料另存一份，還原錯了也能再回來
    const t = new Date(), hm = String(t.getHours()).padStart(2, '0') + String(t.getMinutes()).padStart(2, '0');
    await Drive.upload(`backup/${localDay()}-還原前-${hm}.enc`, await Vault.envelope());
    const now = t.toISOString(), db = emptyDb();
    COLLS.forEach((c) => {
      const cur = Vault.db[c] || {}, o = old[c] || {};
      Object.keys(cur).forEach((id) => { if (!o[id]) db[c][id] = { id, deleted: true, updatedAt: now }; });
      Object.entries(o).forEach(([id, r]) => { db[c][id] = { ...r, updatedAt: now }; });
    });
    Vault.db = db; Vault.markDirty(); await Vault.save(true);
    await this.prune();
  },
};
const bkLabel = (n) => { const m = n.match(/^(\d{4})-(\d{2})-(\d{2})(?:-還原前-(\d{2})(\d{2}))?\.enc$/); if (!m) return n; const day = `${m[1]}/${m[2]}/${m[3]}`; return m[4] ? `${day} ${m[4]}:${m[5]} 還原前的資料` : day + (n.startsWith(localDay()) ? '（今天開啟時）' : ''); };
async function openBackups() {
  ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>備份與還原</h2><button class="btn pri" id="bkX">關閉</button></div>
  <div class="ovbody" style="flex-direction:column;max-width:600px;margin:0 auto;width:100%">
   <div class="note">每天第一次打開 App，會把前一次的資料另存一份到 OneDrive，保留最近 ${KEEP_DAYS} 天。只備份文字資料（案件、客戶、進度、備註…），照片本來就不會被改動，刪掉的照片在「最近刪除」放 30 天。</div>
   <div class="warnbox">還原會把<b>全部資料</b>變回那一天的樣子，之後新增或修改的會不見。所以還原前會自動把現在的資料再存一份，還原錯了可以再還原回來。</div>
   <div class="list" id="bkList" style="padding:0"><div class="sub">讀取中…</div></div></div></div>`;
  $('bkX').onclick = () => (ov.innerHTML = '');
  let names; try { names = await Backup.list(); } catch (e) { $('bkList').innerHTML = '<div class="sub">連不上 OneDrive，請確認網路。</div>'; return; }
  if (!$('bkList')) return;
  $('bkList').innerHTML = names.length ? names.map((n) => `<div class="linkrow"><span class="item" style="cursor:default"><span class="t">${esc(bkLabel(n))}</span></span><button class="btn" data-bk="${esc(n)}">還原到這份</button></div>`).join('') : '<div class="sub">還沒有備份。明天第一次打開 App 時會自動建立第一份。</div>';
  $('bkList').querySelectorAll('[data-bk]').forEach((b) => (b.onclick = async () => {
    if (!b.dataset.sure) { b.dataset.sure = 1; b.textContent = '確定？再按一次'; b.classList.add('danger'); setTimeout(() => { if (b.isConnected) { delete b.dataset.sure; b.textContent = '還原到這份'; b.classList.remove('danger'); } }, 4000); return; }
    b.disabled = true; b.textContent = '還原中…';
    try { await Backup.restore(b.dataset.bk); ov.innerHTML = ''; S.sel = null; S.csel = null; S.showDetail = false; App.render(); toast('已還原。還原前的資料也另存了一份'); }
    catch (e) { b.disabled = false; b.textContent = '還原到這份'; toast('還原失敗：' + (e.message || e)); }
  }));
}
function openTrash() {
  const draw = () => {
    const list = Files.trashed().filter((f) => f.trashCtx);
    const where = (x) => { const r = Vault.get(x.coll, x.recId); return r ? (x.coll === 'clients' ? '客戶 ' + (r.name || '') : '案件 ' + [r.client, r.title].filter(Boolean).join(' · ')) : '（原本的案件／客戶已刪除）'; };
    const left = (f) => Math.max(0, 30 - Math.floor((Date.now() - new Date(f.trashedAt)) / 864e5));
    $('trList').innerHTML = list.length ? `<div class="docs">${list.map((f) => `<div class="doc">${f.trashCtx.doc.type === 'pdf' ? `<div class="pdftile">PDF<small>${esc(f.trashCtx.doc.name)}</small></div>` : `<img data-trimg="${f.id}" alt="">`}<div class="bar" style="flex-direction:column;align-items:stretch;gap:2px"><b style="font-weight:500;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(f.trashCtx.doc.name)}</b><span class="sub" style="font-size:11px">${esc(where(f.trashCtx))}</span><span class="sub" style="font-size:11px">還有 ${left(f)} 天</span><button class="btn" data-trback="${f.id}" style="padding:3px 8px;font-size:12px;margin-top:4px">放回去</button></div></div>`).join('')}</div>` : '<div class="sub">沒有最近刪除的文件。</div>';
    $('trList').querySelectorAll('[data-trimg]').forEach((im) => Files.url(im.dataset.trimg).then((u) => { if (u) im.src = u; }));
    $('trList').querySelectorAll('[data-trback]').forEach((b) => (b.onclick = () => { if (Files.restore(b.dataset.trback)) { toast('已放回原本的地方'); draw(); App.render(); } else toast('原本的案件或客戶已經刪除，放不回去'); }));
  };
  ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>最近刪除</h2><button class="btn pri" id="trX">關閉</button></div>
  <div class="ovbody" style="flex-direction:column;max-width:900px;margin:0 auto;width:100%"><div class="note">刪掉的文件會在這裡放 30 天，之後才真的刪除。按「放回去」會回到原本的案件或客戶。</div><div id="trList"></div></div></div>`;
  $('trX').onclick = () => (ov.innerHTML = '');
  draw();
}

async function boot() {
  gateLoading('載入中…');
  if (!CFG.CLIENT_ID && !MOCK) { showGate(`<p>還差一步：把 Azure「Work」應用程式的用戶端識別碼填進 <b>crm/config.js</b>。</p><p>想先試用，可以開測試模式，資料只會存在這台裝置。</p><a class="btn" href="?test">先用測試模式</a>`); return; }
  try { await Drive.init(); } catch (e) { console.warn(e); showGate(`<p>登入元件載入失敗，請重新整理。</p><div class="err">${esc(e.message || e)}</div>`); return; }
  if (!Drive.account) {
    showGate(`<p>用你的 Microsoft 帳號登入。資料會加密存在你自己的 OneDrive，只有你看得到。</p><button class="btn pri" id="gSign">用 Microsoft 帳號登入</button>`);
    $('gSign').onclick = () => Drive.signIn(); return;
  }
  gateLoading('正在連線 OneDrive…');
  let exists = false, offline = false;
  try { exists = !!(await Drive.getMeta('data.enc')); } catch (e) { offline = true; exists = !!(await Cache.get('data')); }
  if (!exists && offline) { showGate(`<p>連不上 OneDrive，請確認網路後再試一次。</p><button class="btn pri" id="gRetry">重試</button>`); $('gRetry').onclick = boot; return; }
  exists ? askUnlock() : askCreate();
}
function askCreate() {
  showGate(`<p>第一次使用，請設定一組密碼。所有資料（包含照片）都會用這組密碼加密後才存到 OneDrive。</p>
  <div class="warnbox">忘記密碼，資料就救不回來。請把它記在安全的地方。</div>
  <label class="f">密碼（至少 6 個字）<input type="password" id="gP1" autocomplete="new-password"></label>
  <label class="f">再輸入一次<input type="password" id="gP2" autocomplete="new-password"></label>
  <div class="err" id="gErr"></div><button class="btn pri" id="gGo">建立</button>
  <p style="font-size:12px">登入帳號：${esc(Drive.account.username)}</p>`);
  $('gGo').onclick = async () => {
    const a = $('gP1').value, b = $('gP2').value;
    if (a.length < 6) { $('gErr').textContent = '密碼至少要 6 個字'; return; }
    if (a !== b) { $('gErr').textContent = '兩次輸入的密碼不一樣'; return; }
    $('gGo').disabled = true; $('gErr').textContent = '';
    try { await Vault.create(a); enter(); } catch (e) { $('gGo').disabled = false; $('gErr').textContent = '建立失敗：' + (e.message || e); }
  };
}
function askUnlock() {
  showGate(`<p>輸入密碼解鎖。</p>
  <label class="f">密碼<input type="password" id="gP" autocomplete="current-password"></label>
  <div class="err" id="gErr"></div><button class="btn pri" id="gGo">解鎖</button>
  <p style="font-size:12px">登入帳號：${esc(Drive.account.username)}　<button class="btn" id="gOut" style="padding:2px 8px;font-size:12px">換帳號</button></p>`);
  const go = async () => {
    $('gGo').disabled = true; $('gErr').textContent = '';
    try { const r = await Vault.load($('gP').value); if (r === 'new') { askCreate(); return; } enter(); }
    catch (e) { $('gGo').disabled = false; $('gErr').textContent = e.badPassword ? '密碼不正確' : '讀取失敗：' + (e.message || e); }
  };
  $('gGo').onclick = go; $('gP').onkeydown = (e) => { if (e.key === 'Enter') go(); };
  $('gOut').onclick = () => Drive.signOut();
  setTimeout(() => $('gP')?.focus(), 50);
}
let lastAct = Date.now();
if (MOCK) window.__crm = { get Vault() { return Vault; }, migrateLine: () => migrateLine(), migrateMarketing: () => migrateMarketing(), App }; // 只有測試模式才開，方便測試
function migrateLine() { // 一次性：「產險」大類拿掉，裡面的東西改成「產險」屬性
  if (Vault.get('settings', 'migr-line')) return;
  syncCats(); syncRefLists();
  const g = GROUPS.find((x) => x.name === '產險'), pcats = [...new Set([...(g ? g.cats : []), ...(CATS.includes('產險') ? ['產險'] : [])])];
  if (pcats.length) {
    const fb = CATS.includes('其他') && !pcats.includes('其他') ? '其他' : '';
    const kinds = pcats.filter((k) => k !== '產險');
    cases().forEach((c) => { if (!pcats.includes(c.cat)) return; const k = c.cat; c.line = '產險'; c.cat = fb || (GROUPS.find((x) => x !== g && x.cats.length)?.cats[0] || '其他'); if (k !== '產險' && !(c.items || []).includes(k)) c.items = [...(c.items || []), k]; saveCase(c); });
    notes().forEach((n) => { if ((n.kind !== 'sop' && n.kind !== 'tpl') || !pcats.includes(catOf(n))) return; const k = catOf(n); n.line = '產險'; n.cat = fb; if (n.kind === 'sop') n.v2 = true; if (k !== '產險') n.pkinds = [...new Set([...(n.pkinds || []), k])]; Vault.put('notes', n); });
    Vault.put('settings', { id: 'items-產險', list: [...new Set([...ITEMS.產險, ...kinds])] });
    Vault.put('settings', { id: 'tag-pkinds', list: [...new Set([...tagList('pkinds'), ...kinds])] });
    const groups = GROUPS.filter((x) => x !== g).map((x) => ({ name: x.name, cats: x.cats.filter((k) => !pcats.includes(k)) }));
    Vault.put('settings', { id: 'catgroups', list: groups }); Vault.put('settings', { id: 'cats', list: groups.flatMap((x) => x.cats) });
  }
  Vault.put('settings', { id: 'migr-line', done: true }); syncCats(); syncRefLists();
}
function migrateMarketing() { // 一次性：「整合行銷」大類拿掉，改成跟壽險、產險同一層
  if (Vault.get('settings', 'migr-mk')) return;
  syncCats(); syncRefLists();
  const g = GROUPS.find((x) => x.name === '整合行銷'), mcats = [...new Set([...(g ? g.cats : []), ...CATS.filter((k) => MK_WORDS.includes(k))])];
  if (mcats.length) {
    const fb = CATS.includes('其他') && !mcats.includes('其他') ? '其他' : '';
    cases().forEach((c) => { if (!mcats.includes(c.cat)) return; const k = c.cat; c.line = '整合行銷'; c.cat = fb || '其他'; if (!(c.items || []).includes(k)) c.items = [...(c.items || []), k]; saveCase(c); });
    notes().forEach((n) => { if ((n.kind !== 'sop' && n.kind !== 'tpl') || !mcats.includes(catOf(n))) return; const k = catOf(n); n.line = '整合行銷'; n.cat = fb; if (n.kind === 'sop') n.v2 = true; n.mkinds = [...new Set([...(n.mkinds || []), k])]; Vault.put('notes', n); });
    Vault.put('settings', { id: 'items-整合行銷', list: [...new Set([...ITEMS.整合行銷, ...mcats])] });
    Vault.put('settings', { id: 'tag-mkinds', list: [...new Set([...tagList('mkinds'), ...mcats])] });
    const groups = GROUPS.filter((x) => x !== g).map((x) => ({ name: x.name, cats: x.cats.filter((k) => !mcats.includes(k)) }));
    Vault.put('settings', { id: 'catgroups', list: groups }); Vault.put('settings', { id: 'cats', list: groups.flatMap((x) => x.cats) });
  }
  Vault.put('settings', { id: 'migr-mk', done: true }); syncCats(); syncRefLists();
}
function enter() {
  gate.hidden = true; app.hidden = false;
  syncCats(); syncRefLists(); migrateLine(); migrateMarketing();
  cases().filter((c) => c.client && !c.clientId).forEach((c) => { linkClient(c); saveCase(c); }); // 舊案件補連客戶
  cases().filter((c) => !c.date || c.docs.some((d) => !d.date)).forEach((c) => { c.date = caseDate(c); c.docs.forEach((d) => { if (!d.date) d.date = c.date; }); saveCase(c); }); // 舊資料補日期
  const first = cases().slice().sort((a, b) => (b.updatedAt || '').localeCompare(a.updatedAt || ''))[0];
  S.sel = first?.id || null; S.refSel = cases().find((c) => c.isRef)?.id || null; S.showDetail = false;
  lastAct = Date.now(); App.render(); Files.retryUploads();
  Backup.daily().then(() => Files.purge());
}
async function lock() {
  if (Vault.dirty) await Vault.save();
  S.showIds = false; S.showHealth = false; Vault.lock(); Files.clear(); imgCache.clear(); ov.innerHTML = ''; app.innerHTML = '';
  askUnlock();
}
['pointerdown', 'keydown'].forEach((t) => document.addEventListener(t, () => (lastAct = Date.now()), true));
setInterval(() => { if (Vault.key && Date.now() - lastAct > (CFG.AUTO_LOCK_MINUTES || 30) * 60000) lock(); }, 30000);
document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible' && Vault.key) { if (Date.now() - lastAct > (CFG.AUTO_LOCK_MINUTES || 30) * 60000) lock(); else Vault.pull(); } else if (Vault.dirty) Vault.save(); });
window.addEventListener('online', () => { if (Vault.key) { Vault.save(); Files.retryUploads(); } });
window.addEventListener('beforeunload', (e) => { if (Vault.dirty) { Vault.save(); e.preventDefault(); e.returnValue = ''; } });

/* ================= 好打的日期欄位 =================
 * 原生日期欄位打一位數就跳格，改成一般文字欄位：
 * 8/21、0821（今年）、115/8/21、1150821（民國）、2026/8/21、20260821 都可以，離開欄位時自動整理成 2026-08-21。 */
function parseDay(raw, fallbackYear) {
  let t = String(raw || '').trim().replace(/[\uFF10-\uFF19]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xFEE0)).replace(/\s+/g, '');
  if (!t) return '';
  t = t.replace(/[年月.\-]/g, '/').replace(/日$/, '').replace(/^民國/, '');
  let y, m, d; const now = new Date(), Y = fallbackYear || now.getFullYear();
  if (/^\d+$/.test(t)) {
    if (t.length === 8) [y, m, d] = [t.slice(0, 4), t.slice(4, 6), t.slice(6)];
    else if (t.length === 7) [y, m, d] = [t.slice(0, 3), t.slice(3, 5), t.slice(5)];
    else if (t.length === 6) [y, m, d] = [t.slice(0, 2), t.slice(2, 4), t.slice(4)];
    else if (t.length === 4) [y, m, d] = [Y, t.slice(0, 2), t.slice(2)];
    else if (t.length === 3) [y, m, d] = [Y, t.slice(0, 1), t.slice(1)];
    else return null;
  } else {
    const p = t.split('/').filter(Boolean);
    if (p.length === 3) [y, m, d] = p; else if (p.length === 2) [y, m, d] = [Y, ...p]; else return null;
  }
  y = +y; m = +m; d = +d;
  if (y < 1911) y += 1911;
  const dt = new Date(y, m - 1, d);
  if (!(m >= 1 && m <= 12 && d >= 1 && d <= 31) || dt.getMonth() !== m - 1 || y < 1900 || y > 2200) return null;
  return `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
}
function upgradeDate(el) {
  if (el.dataset.dt) return; el.dataset.dt = '1';
  el.type = 'text'; el.setAttribute('inputmode', 'decimal'); el.setAttribute('autocomplete', 'off');
  if (!el.placeholder) el.placeholder = '例：8/21、115/8/21';
  el.classList.add('dtin');
  const wrap = document.createElement('span'); wrap.className = 'dtwrap'; el.parentNode.insertBefore(wrap, el); wrap.appendChild(el);
  const btn = document.createElement('button'); btn.type = 'button'; btn.className = 'dtpick'; btn.setAttribute('aria-label', '選日期'); btn.textContent = '📅'; wrap.appendChild(btn);
  btn.addEventListener('click', (e) => {
    e.preventDefault(); e.stopPropagation();
    const h = document.createElement('input'); h.type = 'date'; h.className = 'dthidden'; h.value = parseDay(el.value) || localDay();
    wrap.appendChild(h);
    h.addEventListener('change', () => { el.value = h.value; h.remove(); el.dispatchEvent(new Event('change', { bubbles: true })); });
    h.addEventListener('blur', () => setTimeout(() => h.remove(), 300));
    try { h.showPicker(); } catch (err) { h.focus(); h.click(); }
  });
}
new MutationObserver((ms) => { for (const m of ms) for (const n of m.addedNodes) { if (n.nodeType !== 1) continue; if (n.matches?.('input[type=date]')) upgradeDate(n); n.querySelectorAll?.('input[type=date]').forEach(upgradeDate); } }).observe(document.body, { childList: true, subtree: true });
document.addEventListener('focusin', (e) => { if (e.target.classList?.contains('dtin')) e.target.dataset.prev = e.target.value; }, true);
document.addEventListener('change', (e) => { // 先整理日期，再交給原本的處理
  const el = e.target; if (!el.classList?.contains('dtin')) return;
  const v = parseDay(el.value);
  if (v === null) { toast('看不懂這個日期，可以打 8/21、115/8/21 或 2026/8/21'); el.value = el.dataset.prev || ''; e.stopImmediatePropagation(); return; }
  el.value = v; el.dataset.prev = v;
}, true);
document.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.classList?.contains('dtin')) { e.preventDefault(); e.target.blur(); } }, true);

/* 有新版本時提示更新（不會自動重新整理，以免打到一半的字不見） */
if ('serviceWorker' in navigator) {
  const hadController = !!navigator.serviceWorker.controller;
  navigator.serviceWorker.register('sw.js').then((reg) => {
    document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') reg.update().catch(() => { }); });
  }).catch(() => { });
  navigator.serviceWorker.addEventListener('controllerchange', () => {
    if (!hadController || document.querySelector('.updbar')) return;
    const bar = document.createElement('button'); bar.className = 'updbar'; bar.textContent = '有新版本，點這裡更新';
    bar.onclick = async () => { bar.textContent = '更新中…'; try { if (Vault.dirty) await Vault.save(); } catch (e) { } location.reload(); };
    document.body.appendChild(bar);
  });
}
boot();
})();
