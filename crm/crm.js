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
};
/* 測試模式：假雲端，資料只存在這台裝置的瀏覽器（網址加 ?test 開啟，?real 關閉） */
const Mock = {
  async getMeta(p) { const v = await Cache.get('mock:' + p); return v ? { eTag: v.eTag, size: v.bytes.length } : null; },
  async download(p) { const v = await Cache.get('mock:' + p); return v ? { bytes: v.bytes, eTag: v.eTag } : null; },
  async upload(p, bytes, eTag) { const v = await Cache.get('mock:' + p); if (eTag && v && v.eTag !== eTag) { const e = new Error('conflict'); e.conflict = true; throw e; } const t = uid(); await Cache.set('mock:' + p, { bytes, eTag: t }); return t; },
  async remove(p) { await Cache.del('mock:' + p); return true; },
};

/* ================= 資料庫（記憶體中明文，存檔時加密） ================= */
const COLLS = ['cases', 'notes', 'files', 'clients', 'settings', 'calls', 'rels'];
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
    if (!fromCache) await Cache.set('data', raw);
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
  async remove(id) { Vault.del('files', id); await Cache.del('file:' + id); Drive.remove(fpath(id)).catch(() => { }); },
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
  scan: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4 8V6a2 2 0 0 1 2-2h2M16 4h2a2 2 0 0 1 2 2v2M20 16v2a2 2 0 0 1-2 2h-2M8 20H6a2 2 0 0 1-2-2v-2M4 12h16"/></svg>',
  bolt: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M13 3 4 14h7l-1 7 9-11h-7l1-7z"/></svg>',
  folder: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4 6a2 2 0 0 1 2-2h4l2 2h6a2 2 0 0 1 2 2v9a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2z"/></svg>',
  book: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4 5a2 2 0 0 1 2-2h13v16H6a2 2 0 0 0-2 2zM4 21V5M8 7h7"/></svg>',
  list: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M9 6h11M9 12h11M9 18h11M4 6l1 1 2-2M4 12l1 1 2-2M4 18l1 1 2-2"/></svg>',
  phone: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M5 4h4l2 5-2.5 1.5a11 11 0 0 0 5 5L15 13l5 2v4a2 2 0 0 1-2 2A16 16 0 0 1 3 6a2 2 0 0 1 2-2"/></svg>',
  user: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="8" r="4"/><path d="M4 21a8 8 0 0 1 16 0"/></svg>',
  msg: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4 5h16v11H9l-5 4z"/></svg>',
};
const NAV = [['quick', '快查', 'bolt'], ['clients', '客戶', 'user'], ['cases', '案件', 'folder'], ['calls', '電訪', 'phone'], ['refs', '理賠參考', 'book'], ['sop', 'SOP', 'list'], ['tpl', '範本', 'msg']];
const DEFAULT_CATS = ['理賠', '建議書', '新契約', '保全', '保費', '保校', '房貸', '服務', '其他', '產險'];
const CATS = [...DEFAULT_CATS]; // 可以在 App 裡自己改，存在設定裡
function syncCats() { const list = Vault.get('settings', 'cats')?.list; const next = list?.length ? list : DEFAULT_CATS; if (next.join('|') !== CATS.join('|')) CATS.splice(0, CATS.length, ...next); }
function openCatEditor() {
  let rows = CATS.map((n) => ({ old: n, name: n }));
  const used = (n) => cases().filter((c) => c.cat === n).length;
  const draw = () => {
    $('ceList').innerHTML = rows.map((r, i) => `<div class="cerow"><input data-ci="${i}" value="${esc(r.name)}" aria-label="類別名稱"><span class="sub">${r.old ? used(r.old) + ' 件' : '新的'}</span><button class="btn" data-cu="${i}" aria-label="上移" ${i ? '' : 'disabled'}>↑</button><button class="btn" data-cd="${i}" aria-label="下移" ${i < rows.length - 1 ? '' : 'disabled'}>↓</button><button class="tldel" data-cx="${i}" aria-label="刪除">✕</button></div>`).join('');
  };
  ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>編輯案件類別</h2><button class="btn" id="ceX">取消</button><button class="btn pri" id="ceSave">儲存</button></div>
  <div class="ovbody" style="flex-direction:column;max-width:560px;margin:0 auto;width:100%">
   <div class="note">改名字，舊案件會一起改過去。有案件在用的類別不能刪，要先把那些案件改到別的類別。<br>「理賠」有理賠分類和理賠參考，「新契約」「建議書」有轉換和要保人功能，這幾個建議保留原名。</div>
   <div id="ceList"></div>
   <div class="actions"><input class="search" id="ceNew" placeholder="新類別名稱，例如：轉介紹" style="flex:1"><button class="btn" id="ceAdd">＋ 新增</button><button class="btn" id="ceReset">恢復預設</button></div></div></div>`;
  draw();
  $('ceList').oninput = (e) => { const i = e.target.dataset.ci; if (i != null) rows[+i].name = e.target.value; };
  $('ceList').onclick = (e) => {
    const b = e.target.closest('button'); if (!b) return; const d = b.dataset;
    if (d.cu) { const i = +d.cu; [rows[i - 1], rows[i]] = [rows[i], rows[i - 1]]; }
    if (d.cd) { const i = +d.cd; [rows[i + 1], rows[i]] = [rows[i], rows[i + 1]]; }
    if (d.cx) { const r = rows[+d.cx]; if (r.old && used(r.old)) { toast(`「${r.old}」還有 ${used(r.old)} 件案件，先把它們改到別的類別`); return; } rows.splice(+d.cx, 1); }
    draw();
  };
  const add = () => { const v = $('ceNew').value.trim(); if (!v) return; if (rows.some((r) => r.name.trim() === v)) { toast('已經有這個類別'); return; } rows.push({ old: '', name: v }); $('ceNew').value = ''; draw(); };
  $('ceAdd').onclick = add; $('ceNew').onkeydown = (e) => { if (e.key === 'Enter') add(); };
  $('ceReset').onclick = () => { const keep = rows.filter((r) => r.old && used(r.old) && !DEFAULT_CATS.includes(r.name.trim())); rows = [...DEFAULT_CATS.map((n) => ({ old: rows.find((r) => r.old === n)?.old || '', name: n })), ...keep]; draw(); };
  $('ceX').onclick = () => (ov.innerHTML = '');
  $('ceSave').onclick = () => {
    const names = rows.map((r) => r.name.trim());
    if (names.some((n) => !n)) { toast('類別名稱不能空白'); return; }
    if (new Set(names).size !== names.length) { toast('有重複的類別名稱'); return; }
    const ren = rows.filter((r) => r.old && r.old !== r.name.trim()).map((r) => [r.old, r.name.trim()]);
    if (ren.length) {
      const m = new Map(ren);
      cases().forEach((c) => { if (m.has(c.cat)) { c.cat = m.get(c.cat); saveCase(c); } });
      notes().forEach((n) => { if (n.cat && m.has(n.cat)) { n.cat = m.get(n.cat); Vault.put('notes', n); } });
      const st = Vault.get('settings', 'steps'); if (st?.map) { ren.forEach(([a, b]) => { if (st.map[a]) { st.map[b] = st.map[a]; delete st.map[a]; } else if (STEPS[a] && !st.map[b]) st.map[b] = STEPS[a]; }); Vault.put('settings', st); }
      if (m.has(S.cat)) S.cat = m.get(S.cat);
    }
    Vault.put('settings', { id: 'cats', list: names }); syncCats(); ov.innerHTML = ''; App.render(); toast('案件類別已更新');
  };
}
const STATUS = ['進行中', '待回覆', '已結案'];
const ITEMS = ['住院醫療', '實支實付', '住院日額', '手術', '意外醫療', '意外失能', '骨折', '重大傷病', '癌症', '長照失能', '門診', '產險-車險', '產險-住火', '產險-旅平'];
const CAUSE = { '疾病': ['消化系統', '呼吸系統', '心血管', '骨科', '婦科', '腫瘤', '眼耳鼻喉', '其他疾病'], '意外': ['交通事故', '跌倒', '運動傷害', '職業傷害', '燒燙傷', '其他意外'] };
const TPLGROUPS = ['LINE', '簡訊', '信件', '其他'];
const SOPGROUPS = ['壽險·新契約', '壽險·保全', '壽險·保費', '壽險·保校', '壽險·理賠', '壽險·房貸', '壽險·服務', '壽險·其他', '產險', '推薦連結'];
const linkify = (html) => html.replace(/(https?:\/\/[^\s<]+)/g, (u) => `<a href="${u}" target="_blank" rel="noopener">${u.length > 60 ? u.slice(0, 57) + '…' : u}</a>`);

/* ================= 狀態與畫面 ================= */
const S = { view: 'cases', showHealth: false, tplGroup: '全部', tplCat: '全部', ksel: null, kst: '全部', kres: '', script: null, ccat: '全部', csel: null, showIds: false, sel: null, refSel: null, cat: '全部', q: '', refItem: '全部', refCause: '全部', sopGroup: '全部', showDetail: false };
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
    if (!Vault.db) return;
    syncCats(); if (S.cat !== '全部' && !CATS.includes(S.cat)) S.cat = '全部';
    const v = S.view;
    app.className = 'app' + ((v === 'cases' || v === 'refs' || v === 'clients' || v === 'calls') ? '' : ' single') + (S.showDetail ? ' showdetail' : '') + (MOCK ? ' test' : '');
    let h = MOCK ? '<div class="testbar">測試模式：資料只存在這台裝置，不會上傳 OneDrive</div>' : '';
    h += railHTML();
    if (v === 'clients') h += clientListHTML() + '<section class="detail split">' + clientDetailHTML() + '</section>';
    else if (v === 'cases') h += caseListHTML() + '<section class="detail split">' + caseDetailHTML() + '</section>';
    else if (v === 'calls') h += callListHTML() + '<section class="detail split">' + callDetailHTML() + '</section>';
    else if (v === 'refs') h += refListHTML() + '<section class="detail split">' + refDetailHTML() + '</section>';
    else if (v === 'sop') h += '<section class="wide">' + notesHTML('sop') + '</section>';
    else if (v === 'quick') h += '<section class="wide">' + quickHTML() + '</section>';
    else if (v === 'tpl') h += '<section class="wide">' + notesHTML('tpl') + '</section>';
    h += tabsHTML();
    const keep = document.activeElement?.id, pos = document.activeElement?.selectionStart;
    const sc = [...app.querySelectorAll('section, .list')].map((el) => [el.className, el.scrollTop]), wy = window.scrollY; // 重畫後回到原本的捲動位置
    app.innerHTML = h;
    const pk = [S.view, S.sel, S.csel, S.ksel, S.rsel, S.showDetail].join('|'), same = App._pk === pk; App._pk = pk;
    if (same) { const els = [...app.querySelectorAll('section, .list')]; sc.forEach(([cn, t]) => { const el = els.find((e) => e.className === cn); if (el && t) el.scrollTop = t; }); if (wy) window.scrollTo(0, wy); }
    if (keep && (keep === 'q')) { const el = $(keep); if (el) { el.focus(); try { el.setSelectionRange(pos, pos); } catch (e) { } } }
    drawThumbs(); fitNotes();
    if (v === 'clients' && curClient()) drawEgo(curClient());
  },
};
function railHTML() {
  return `<nav class="rail" aria-label="主選單"><div class="brand"><i></i>工作手冊</div>
  <button class="scanbtn" data-act="scan">${ICON.scan}掃描</button>
  ${NAV.map(([k, l, i]) => `<button class="nav${S.view === k ? ' on' : ''}" data-nav="${k}">${ICON[i]}${l}</button>`).join('')}
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
  return cases().filter((c) => (S.cat === '全部' || c.cat === S.cat) && (!q || [c.client, c.title, c.notes, c.diag, c.next, ...(c.people || []).map((p) => Vault.get('clients', p.clientId)?.name || '')].join(' ').includes(q))).sort(byDate);
}
function caseListHTML() {
  const list = filteredCases();
  return `<section class="col"><div class="colhead"><h1>案件 <button class="btn" data-act="newcase">＋ 新增</button></h1>
  <div class="mobsync">${syncHTML()}</div>
  <input class="search" id="q" placeholder="搜尋客戶、案件、備註" value="${esc(S.q)}" autocomplete="off">
  <div class="chips">${['全部', ...CATS].map((c) => `<button class="chip${S.cat === c ? ' on' : ''}" data-cat="${esc(c)}">${esc(c)}</button>`).join('')}<button class="chip chipedit" data-act="editcats" aria-label="編輯案件類別">⚙ 編輯</button></div></div>
  <div class="list">${list.length ? list.map((c) => `<button class="item${c.id === S.sel ? ' on' : ''}${itemCls(c.status)}" data-sel="${c.id}" title="${esc((c.client || '') + ' · ' + (c.title || ''))}"><span class="t"><span>${esc(c.client || '未填客戶')} · ${esc(c.title || '未命名')}</span><span class="pill ${pillCls(c.status)}">${c.status}</span></span><span class="s">${fmtDay(caseDate(c))} · ${esc(caseSummary(c))}${latestStep(c) ? ' · 最新：' + esc(latestStep(c).title) : ''}</span></button>`).join('') : `<div class="empty">${cases().length ? '沒有符合的案件。' : '還沒有案件。按「＋ 新增」，或按「掃描」直接拍文件建案。'}</div>`}</div></section>`;
}
function caseDetailHTML() {
  const c = cur();
  if (!c) return '<div class="empty">從左邊選一個案件，或按「掃描」直接拍文件建案。</div>';
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
   <label class="f">類別<select id="f_cat" data-f="cat">${(CATS.includes(c.cat) ? CATS : [c.cat, ...CATS]).map((x) => `<option${x === c.cat ? ' selected' : ''}>${x}</option>`).join('')}</select></label>
   <div class="f">狀態<div class="stseg" role="group" aria-label="狀態">${STATUS.map((x) => `<button class="${x === c.status ? 'on ' + pillCls(x) : ''}" data-setst="${x}" aria-pressed="${x === c.status}">${x}</button>`).join('')}</div></div>
   <label class="f">下一步<input id="f_next" data-f="next" value="${esc(c.next)}" placeholder="例如：補診斷書正本"></label></div></div>
  ${peopleHTML(c)}
  ${timelineHTML(c)}
  ${isClaim ? `<div class="sec"><h2>理賠分類</h2>
   <div class="sub">險種項目（可複選）</div><div class="tags">${ITEMS.map((x) => `<button class="tag${c.items.includes(x) ? ' on' : ''}" data-item="${x}">${x}</button>`).join('')}</div>
   <div class="sub">原因</div><div class="tags">${Object.keys(CAUSE).map((x) => `<button class="tag${c.causeType === x ? ' on' : ''}" data-ctype="${x}">${x}</button>`).join('')}</div>
   ${c.causeType ? `<div class="tags">${CAUSE[c.causeType].map((x) => `<button class="tag${c.causeDetail === x ? ' on' : ''}" data-cdet="${x}">${x}</button>`).join('')}</div>` : ''}
   <label class="f">診斷／受傷情況<input id="f_diag" data-f="diag" value="${esc(c.diag)}" placeholder="例如：腳踝扭傷併骨裂，打石膏"></label></div>
  <div class="sec"><h2>理賠計算</h2><div class="calc">
   <label class="f">實際花費<input id="f_cost" inputmode="numeric" data-f="cost" value="${esc(c.cost)}" placeholder="76600"></label><span class="op">×</span>
   <label class="f">給付比例 %<input id="f_ratio" inputmode="numeric" data-f="ratio" value="${esc(c.ratio)}" placeholder="65"></label><span class="op">=</span>
   <span class="res">${auto ? auto.toLocaleString() : '—'}</span></div>
   <label class="f">實際理賠金額（有不同再填）<input id="f_payout" inputmode="numeric" data-f="payout" value="${esc(c.payout)}"></label></div>` : ''}
  ${caseDocsHTML(c)}
  ${linksHTML(c)}
  ${caseTplHTML(c)}
  <div class="sec"><h2>備註</h2><textarea id="f_notes" class="grow" data-f="notes" rows="4" aria-label="備註">${esc(c.notes)}</textarea></div>
  <div class="actions"><button class="btn pri savebtn" data-act="savenow">儲存</button>${isClaim ? `<button class="btn${c.isRef ? ' on' : ''}" data-act="ref">${c.isRef ? '✓ 已在理賠參考' : '做成理賠參考'}</button>` : ''}${c.clientId && ['新契約', '建議書', '保全', '產險'].includes(c.cat) ? '<button class="btn" data-act="casedisc">記錄財務／體況告知</button>' : ''}<button class="btn danger" data-act="delcase">刪除案件</button></div>`;
}
/* ================= 電訪 ================= */
const calls = () => Vault.all('calls');
const curCall = () => Vault.get('calls', S.ksel);
const CALL_ST = ['待撥', '再聯絡', '已約訪', '結案'];
const RESULTS = [['未接', '再聯絡'], ['已接・聊過', '再聯絡'], ['約訪成功', '已約訪'], ['請我再打', '再聯絡'], ['拒絕', '結案'], ['空號／錯號', '結案']];
const scripts = () => notes().filter((n) => n.kind === 'script').sort((a, b) => (a.createdAt || '').localeCompare(b.createdAt || ''));
const logsOf = (k) => (k.logs || []).slice().sort((a, b) => (b.date || '').localeCompare(a.date || '') || (b.at || '').localeCompare(a.at || ''));
const stCls = (st) => (st === '已約訪' ? 'p-done' : st === '結案' ? 'p-off' : st === '再聯絡' ? 'p-wait' : 'p-go');
const SAMPLE_SCRIPT = '1. 開場\n您好，請問是○○先生／小姐嗎？我是國泰人壽的服務人員馥瑋，之後由我負責您的保單服務。\n\n2. 確認時間\n請問現在講電話方便嗎？大概佔用您三分鐘。\n\n3. 說明來意\n這次打來主要是幫您確認聯絡資料，順便看看保單有沒有需要協助的地方。\n\n4. 了解近況\n最近家裡、工作有沒有什麼變化？之前有沒有理賠或保單上的問題？\n\n5. 約訪\n我可以帶一份保單整理給您看，您比較方便平日還是週末？\n\n6. 結尾\n謝謝您！我先加您 LINE，有任何保單問題都可以直接找我。';
function callListHTML() {
  const q = S.q.trim(), all = calls(), today = localDay();
  const rank = (k) => ({ 再聯絡: 0, 待撥: 1, 已約訪: 2, 結案: 3 }[k.status] ?? 1);
  const list = all.filter((k) => (S.kst === '全部' || k.status === S.kst) && (!q || [k.name, k.phone, k.source, k.note].join(' ').includes(q)))
    .sort((a, b) => rank(a) - rank(b) || (a.nextDate || '9999').localeCompare(b.nextDate || '9999') || (a.name || '').localeCompare(b.name || '', 'zh-Hant'));
  return `<section class="col"><div class="colhead"><h1>電訪 <span style="display:flex;gap:6px"><button class="btn" data-act="importcalls">匯入名單</button><button class="btn" data-act="newcall">＋ 新增</button></span></h1>
  <div class="mobsync">${syncHTML()}</div>
  <input class="search" id="q" placeholder="搜尋姓名、電話、來源" value="${esc(S.q)}" autocomplete="off">
  <div class="chips">${['全部', ...CALL_ST].map((x) => `<button class="chip${S.kst === x ? ' on' : ''}" data-kst="${x}">${x} <span class="cnt">${x === '全部' ? all.length : all.filter((k) => k.status === x).length}</span></button>`).join('')}</div></div>
  <div class="list">${list.length ? list.map((k) => { const l = logsOf(k)[0], due = k.nextDate && k.status !== '結案' && k.nextDate <= today; return `<button class="item${k.id === S.ksel ? ' on' : ''}" data-ksel="${k.id}"><span class="t"><span>${esc(k.name || '未填姓名')}</span><span class="pill ${stCls(k.status)}">${esc(k.status)}</span></span><span class="s">${l ? `上次 ${fmtDay(l.date).slice(5)} ${esc(l.result)}` : '還沒打過'}${k.nextDate && k.status !== '結案' ? ` · <span class="${due ? 'due' : ''}">下次 ${fmtDay(k.nextDate).slice(5)}</span>` : ''}${k.source ? ' · ' + esc(k.source) : ''}</span></button>`; }).join('') : `<div class="empty">${all.length ? '沒有符合的對象。' : '還沒有電訪名單。按「匯入名單」一次貼上很多人，或「＋ 新增」一個一個加。'}</div>`}</div></section>`;
}
function callDetailHTML() {
  const k = curCall();
  if (!k) return '<div class="empty">從左邊選一位要撥打的對象。</div>';
  const scs = scripts(); if (!scs.some((x) => x.id === S.script)) S.script = scs[0]?.id || null; const sc = Vault.get('notes', S.script);
  const cl = k.clientId && Vault.get('clients', k.clientId), logs = logsOf(k);
  return `<button class="back" data-act="back">‹ 電訪名單</button>
  <div class="dhead"><input class="title" id="kf_name" data-kf="name" value="${esc(k.name)}" placeholder="姓名" aria-label="姓名"><span class="pill ${stCls(k.status)}">${esc(k.status)}</span></div>
  <div class="sec"><h2>聯絡資料</h2><div class="grid2">
   <label class="f"><span style="display:flex;justify-content:space-between">電話${k.phone ? `<a class="linkbtn" href="tel:${esc(k.phone.replace(/[^0-9+#*]/g, ''))}">撥打</a>` : ''}</span><input id="kf_phone" data-kf="phone" inputmode="tel" value="${esc(k.phone)}"></label>
   <label class="f">狀態<select id="kf_status" data-kf="status">${CALL_ST.map((x) => `<option${x === k.status ? ' selected' : ''}>${x}</option>`).join('')}</select></label>
   <label class="f">下次聯絡<input type="date" id="kf_nextDate" data-kf="nextDate" value="${esc(k.nextDate)}"></label>
   <label class="f">來源／名單<input id="kf_source" data-kf="source" value="${esc(k.source)}" placeholder="例如：10 月服務件"></label></div>
   <label class="f">備註<textarea id="kf_note" class="grow" data-kf="note" rows="2" placeholder="家庭、保單、個性…">${esc(k.note)}</textarea></label></div>
  <details class="sec script" open><summary><h2 style="display:inline">話術 SOP</h2></summary>
   ${scs.length ? `<div class="actions" style="align-items:center"><select id="kScript" aria-label="選擇話術">${scs.map((x) => `<option value="${x.id}"${x.id === S.script ? ' selected' : ''}>${esc(x.title)}</option>`).join('')}</select><button class="linkbtn" data-act="editscript">編輯</button><button class="linkbtn" data-act="newscript">＋ 新話術</button></div><pre class="scripttext">${esc(sc?.body || '')}</pre>${nimgsHTML(sc || {})}`
    : `<div class="sub">還沒有話術。可以先用範例改成你自己的說法。</div><div class="actions"><button class="btn pri" data-act="samplescript">用範例開始</button><button class="btn" data-act="newscript">自己寫</button></div>`}</details>
  <div class="sec"><h2>記錄這通</h2>
   <div class="tags">${RESULTS.map(([r]) => `<button class="tag${S.kres === r ? ' on' : ''}" data-kres="${r}">${r}</button>`).join('')}</div>
   <label class="f">聊了什麼<textarea id="kNote" rows="3" placeholder="重點、對方的需求、約了什麼…"></textarea></label>
   <div class="actions" style="align-items:flex-end"><label class="f" style="width:170px">下次聯絡（可不填）<input type="date" id="kNext"></label><button class="btn pri" data-act="logcall">記錄</button></div></div>
  <div class="sec"><h2>通話紀錄 <span class="sub">${logs.length ? logs.length + ' 通，新的在上面' : ''}</span></h2>
   ${logs.length ? `<div class="tl">${logs.map((l) => `<div class="tli"><span class="tldot"></span><div class="tlrow"><span class="tldate" style="padding:3px 6px">${fmtDay(l.date)}</span><span class="tltitle" style="padding:3px 6px">${esc(l.result)}</span><button class="tldel" data-kdel="${l.id}" aria-label="刪除這通">✕</button></div>${l.note ? `<div class="lognote">${esc(l.note)}</div>` : ''}</div>`).join('')}</div>` : '<div class="sub">還沒有紀錄。</div>'}</div>
  <div class="actions"><button class="btn pri savebtn" data-act="savenow">儲存</button>${cl ? `<button class="btn" data-goclient="${cl.id}">看客戶資料 ›</button>` : '<button class="btn" data-act="tocliient">轉成客戶</button>'}<button class="btn danger" data-act="delcall">刪除</button></div>`;
}
function openCallImport() {
  ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>匯入電訪名單</h2><button class="btn" id="ciCancel">取消</button><button class="btn pri" id="ciGo">匯入</button></div>
  <div class="ovbody" style="flex-direction:column;max-width:640px;margin:0 auto;width:100%">
   <div class="note">一行一個人：姓名、電話，中間用空白或逗號隔開，後面可以接備註。可以直接從 Excel 複製兩欄貼上。</div>
   <label class="f">來源／名單名稱<input id="ciSrc" placeholder="例如：10 月服務件" value="${new Date().getMonth() + 1} 月服務件"></label>
   <label class="f">名單<textarea id="ciText" rows="12" placeholder="王小明 0912345678\n陳美麗, 0987654321, 有小孩"></textarea></label>
   <div class="sub" id="ciCount"></div></div></div>`;
  const parse = () => $('ciText').value.split('\n').map((l) => l.trim()).filter(Boolean).map((l) => { const parts = l.split(/[\t,，]+|\s+/).filter(Boolean); const pi = parts.findIndex((x) => /^[0-9+()#\-]{7,}$/.test(x)); const phone = pi >= 0 ? parts[pi] : ''; const rest = parts.filter((_, i) => i !== pi); return { name: rest[0] || '', phone, note: rest.slice(1).join(' ') }; }).filter((x) => x.name || x.phone);
  $('ciText').oninput = () => { const n = parse().length; $('ciCount').textContent = n ? `會新增 ${n} 位` : ''; };
  $('ciCancel').onclick = () => (ov.innerHTML = '');
  $('ciGo').onclick = () => { const rows = parse(); if (!rows.length) { toast('先貼上名單'); return; } const src = $('ciSrc').value.trim(); rows.forEach((r) => Vault.put('calls', { ...r, source: src, status: '待撥', nextDate: '', logs: [] })); ov.innerHTML = ''; S.kst = '全部'; App.render(); toast(`已匯入 ${rows.length} 位`); };
  setTimeout(() => $('ciText')?.focus(), 50);
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
  let tgt = CRMI.target; if (!tgt) tgt = clients().find((c) => (data.idNo && c.idNo === data.idNo) || (data.name && c.name === data.name)) || null;
  CRMI.resolved = tgt;
  const keys = CRM_FIELDS.filter(([k]) => data[k]);
  if (!keys.length) { $('ciForm').innerHTML = '<div class="note" style="color:var(--warn);background:var(--warn-soft)">沒有讀到資料。截圖請包含「姓名」「ID」「手機」那一區，或改用複製文字。</div>'; return; }
  $('ciForm').innerHTML = `<div class="sec"><h2>確認要帶入的資料 <span class="sub">${tgt ? `會更新「${esc(tgt.name)}」` : '會建立新客戶'}</span></h2>
   ${keys.map(([k, l]) => { const old = tgt?.[k]; return `<div class="cirow"><label class="cichk"><input type="checkbox" data-cik="${k}" ${!old || old !== data[k] ? 'checked' : ''}> ${l}</label><input class="ciin" id="ci_${k}" value="${esc(data[k])}">${old && old !== data[k] ? `<span class="sub">目前：${esc(k === 'idNo' ? old.slice(0, 3) + '…' : old)}</span>` : old ? '<span class="sub">相同</span>' : ''}</div>`; }).join('')}
   ${CRMI.shot ? '<label class="cichk" style="margin-top:6px"><input type="checkbox" id="ciKeep"> 把這張截圖也存進「證件與文件」</label>' : ''}
   <div class="note">地址、姓名請特別看一下，辨識偶爾會錯一兩個字。</div>
   <div class="actions"><button class="btn pri" id="ciApply">${tgt ? '更新資料' : '建立客戶'}</button></div></div>`;
  $('ciApply').onclick = async () => {
    const isNew = !CRMI.resolved, rec = CRMI.resolved || { name: '', docs: [] };
    ov.querySelectorAll('[data-cik]').forEach((cb) => { if (cb.checked) rec[cb.dataset.cik] = $('ci_' + cb.dataset.cik).value.trim(); });
    const saved = Vault.put('clients', rec); const keep = $('ciKeep')?.checked, shot = CRMI.shot;
    closeCrmImport(); S.view = 'clients'; S.csel = saved.id; S.showDetail = true; S.showIds = false; App.render();
    if (keep && shot) addImagesDirect(saved, 'clients', [shot], '');
    toast(`已${isNew ? '建立' : '更新'}「${saved.name || '客戶'}」`);
  };
}
/* ================= 告知資料（財務、體況） ================= */
const ASSETS = [['job', '職業／職稱', ''], ['company', '公司', ''], ['deposit', '存款', '萬'], ['securities', '股票・基金', '萬'], ['realEstate', '不動產', '萬'], ['policyValue', '保單價值', '萬'], ['otherAssets', '其他資產', '萬'], ['mortgage', '房貸', '萬'], ['otherDebt', '信貸・其他負債', '萬']];
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
  const fin = FIN.filter(([k]) => f[k] !== undefined && f[k] !== '').map(([k, l, u]) => `<div class="kv"><span>${l}</span><b>${esc(f[k])} ${u}</b></div>`).join('');
  const b = bmi(h.height, h.weight);
  return `${fin || f.mainEarner ? `<div class="sub">財務</div><div class="kvs">${fin}${f.mainEarner ? `<div class="kv"><span>主要經濟來源</span><b>${esc(f.mainEarner)}${f.mainEarner === '否' && f.relation ? '（主要經濟者是' + esc(f.relation) + '）' : ''}</b></div>` : ''}</div>${f.reason ? `<div class="sub">收入不一致原因：${esc(f.reason)}</div>` : ''}` : ''}
   ${h.height || h.weight || h.smoke || h.drink || h.items ? `<div class="sub">體況</div><div class="kvs">${h.height ? `<div class="kv"><span>身高</span><b>${esc(h.height)} cm</b></div>` : ''}${h.weight ? `<div class="kv"><span>體重</span><b>${esc(h.weight)} kg${b ? '（BMI ' + b + '）' : ''}</b></div>` : ''}${h.smoke ? `<div class="kv"><span>吸菸</span><b>${esc(h.smoke)}</b></div>` : ''}${h.drink ? `<div class="kv"><span>飲酒</span><b>${esc(h.drink)}</b></div>` : ''}</div>${h.items ? `<pre class="discpre">${esc(h.items)}</pre>` : ''}` : ''}
   ${(() => { const a = r.assets || {}; const kv = ASSETS.filter(([k]) => a[k]).map(([k, l, u]) => `<div class="kv"><span>${l}</span><b>${esc(a[k])}${u ? ' ' + u : ''}</b></div>`).join(''); return kv || a.explain ? `<div class="sub">財產明細</div><div class="kvs">${kv}</div>${a.explain ? `<div class="sub">說明：${esc(a.explain)}</div>` : ''}` : ''; })()}
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
  const hasAssets = ASSETS.some(([k]) => a[k]), hasMore = h.smoke || h.drink || h.items || f.reason || base.note;
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
     ${inp('relation', f.relation, '主要經濟者是', '', 'placeholder="例：配偶"')}</div></div>
   <details class="sec fold"${hasAssets ? ' open' : ''}><summary><h2>財產明細 <span class="sub">需要交代財產時（財務狀況告知書）</span></h2></summary>
    <div class="dgrid">${ASSETS.map(([k, l, u]) => inp('a_' + k, a[k], l, u, u ? 'inputmode="decimal"' : '')).join('')}</div>
    ${inp('a_explain', a.explain, '資金來源／說明', '')}</details>
   <details class="sec fold"${hasMore ? ' open' : ''}><summary><h2>體況詳細・其他 <span class="sub">不熟的客戶再填</span></h2></summary>
    <div class="dgrid"><label class="f">吸菸<select id="d_smoke">${['', '無', '有', '已戒'].map((o) => `<option${o === (h.smoke || '') ? ' selected' : ''}>${o}</option>`).join('')}</select></label>
     <label class="f">飲酒<select id="d_drink">${['', '無', '偶爾', '經常'].map((o) => `<option${o === (h.drink || '') ? ' selected' : ''}>${o}</option>`).join('')}</select></label></div>
    <label class="f">告知事項（病史、住院、手術、用藥、體檢異常…）<textarea id="d_items" class="grow" rows="3">${esc(h.items)}</textarea></label>
    ${inp('reason', f.reason, '收入不一致原因', '')}
    <label class="f">其他備註<textarea id="d_note" class="grow" rows="2">${esc(base.note)}</textarea></label></details>
   <div class="sec"><h2>附件 <span class="sub">要保書、告知書、財務狀況告知書的照片或 PDF</span></h2>
    <div class="nimgs edit" id="dFiles"></div>
    <div class="actions"><label class="btn">＋ 加照片或 PDF<input type="file" id="dAdd" accept="image/*,application/pdf,.pdf" multiple></label><button class="btn" id="dPaste">📋 貼上</button></div></div>
  </div></div>`;
  drawDiscFiles(); fitNotes();
  $('dX').onclick = () => { if (DSC.isNew) (DSC.r.files || []).forEach((x) => Files.remove(x.fileId)); DSC = null; ov.innerHTML = ''; };
  $('dSave').onclick = () => {
    const r = DSC.r, g = (k) => $('d_' + k)?.value.trim() ?? '';
    r.date = g('date') || localDay(); r.caseId = $('d_case').value; r.note = g('note');
    r.fin = Object.fromEntries([...FIN.map(([k]) => [k, g(k)]), ['mainEarner', g('mainEarner')], ['relation', g('relation')], ['reason', g('reason')]].filter(([, v]) => v !== ''));
    r.health = Object.fromEntries([['height', g('height')], ['weight', g('weight')], ['smoke', g('smoke')], ['drink', g('drink')], ['items', $('d_items').value.trim()]].filter(([, v]) => v !== ''));
    r.assets = Object.fromEntries([...ASSETS.map(([k]) => [k, g('a_' + k)]), ['explain', g('a_explain')]].filter(([, v]) => v !== ''));
    const c2 = DSC.cl; c2.disclosures = [...(c2.disclosures || []).filter((x) => x.id !== r.id), r]; Vault.put('clients', c2);
    DSC = null; ov.innerHTML = ''; S.view = 'clients'; S.csel = c2.id; S.showDetail = true; S.showHealth = true; App.render(); toast('告知資料已儲存');
  };
  $('dAdd').onchange = (e) => { const fs = [...e.target.files]; e.target.value = ''; discAddFiles(fs, false); };
  $('dPaste').onclick = async () => { const fs = await clipFiles(); if (fs.length) discAddFiles(fs, false); };
  $('dShot').onchange = (e) => { const fl = e.target.files[0]; e.target.value = ''; if (fl) discAddFiles([fl], true); };
  $('dPasteShot').onclick = async () => { const fs = await clipFiles(); if (fs.length) discAddFiles(fs, true); };
}
async function clipFiles() {
  try { const out = []; for (const it of await navigator.clipboard.read()) { const t = it.types.find((y) => y.startsWith('image/')); if (t) out.push(new File([await it.getType(t)], 'shot.png', { type: t })); } if (!out.length) toast('剪貼簿裡沒有圖片'); return out; }
  catch (e) { toast('這裡不能讀剪貼簿，請直接按 Ctrl+V'); return []; }
}
function drawDiscFiles() {
  const box = $('dFiles'); if (!box || !DSC) return;
  box.innerHTML = (DSC.r.files || []).map((f, i) => `<div class="nimgwrap">${f.type === 'pdf' ? `<button class="nimg pdfmini" data-dv="${i}">PDF</button>` : `<button class="nimg" data-dv="${i}"><img data-fid="${f.fileId}" alt=""></button>`}<button class="nimgdel" data-dd="${i}" aria-label="移除">✕</button></div>`).join('') || '<span class="sub">還沒有附件</span>';
  box.querySelectorAll('img[data-fid]').forEach(async (im) => { const u = await Files.url(im.dataset.fid).catch(() => null); if (u && im.isConnected) im.src = u; });
  box.querySelectorAll('[data-dv]').forEach((b) => (b.onclick = () => { const f = DSC.r.files[+b.dataset.dv]; Files.url(f.fileId).then((u) => { if (!u) return; const v = document.createElement('div'); v.innerHTML = `<div class="ov" style="z-index:30"><div class="ovhead"><h2>${esc(f.name)}</h2><button class="btn pri">關閉</button></div>${f.type === 'pdf' ? `<iframe src="${u}" style="flex:1;width:100%;border:0;background:#fff"></iframe>` : `<div class="ovbody" style="justify-content:center"><img src="${u}" alt="" style="max-width:100%;height:auto;align-self:flex-start"></div>`}</div>`; v.querySelector('button').onclick = () => v.remove(); document.body.appendChild(v); }); }));
  box.querySelectorAll('[data-dd]').forEach((b) => (b.onclick = () => { const [f] = DSC.r.files.splice(+b.dataset.dd, 1); if (f) Files.remove(f.fileId); drawDiscFiles(); }));
}
async function discAddFiles(files, readFin) {
  if (!DSC) return; const r = DSC.r; r.files = r.files || [];
  for (const f of files) {
    const pdf = f.type === 'application/pdf' || /\.pdf$/i.test(f.name || '');
    if (!pdf && !/^image\//.test(f.type)) continue;
    const blob = pdf ? new Blob([await f.arrayBuffer()], { type: 'application/pdf' }) : await shrink(f); if (!blob) continue;
    const fileId = await Files.add(blob, f.name || '告知');
    r.files.push({ fileId, name: pdf ? (f.name || '').replace(/\.pdf$/i, '') || 'PDF' : '告知截圖', type: pdf ? 'pdf' : 'img', date: localDay() });
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
   ${docs.some((d) => !isPdf(d)) ? '<div class="actions"><button class="btn" data-act="cpdf">圖片合成 PDF</button><button class="btn" data-act="cshare">匯出遮罩版</button></div>' : ''}</div>
  <div class="sec"><h2>案件 <button class="btn" data-act="ccase" style="padding:4px 10px;font-size:13px">＋ 新增${S.ccat !== '全部' ? esc(S.ccat) : '案件'}</button></h2>
   ${cs.length ? `<div class="chips">${['全部', ...CATS.filter((k) => cs.some((c) => c.cat === k))].map((k) => `<button class="chip${S.ccat === k ? ' on' : ''}" data-ccat="${k}">${k} <span class="cnt">${k === '全部' ? cs.length : cs.filter((c) => c.cat === k).length}</span></button>`).join('')}</div>` : ''}
   ${cs.length ? cs.filter((c) => S.ccat === '全部' || c.cat === S.ccat).map((c) => `<button class="item${itemCls(c.status)}" data-gocase="${c.id}"><span class="t"><span>${esc(c.title || '未命名')}</span><span class="pill ${pillCls(c.status)}">${c.status}</span></span>${casePeopleLine(c, x.id)}<span class="s">${fmtDay(caseDate(c))} · ${esc(caseSummary(c))}${c.next ? ' · 下一步：' + esc(c.next) : ''}</span></button>`).join('') : '<div class="sub">還沒有案件。</div>'}</div>
  <div class="actions"><button class="btn pri savebtn" data-act="savenow">儲存</button><button class="btn danger" data-act="delclient">刪除客戶</button></div>`;
}
const isPdf = (d) => d.type === 'pdf';
function fillTpl(body, c) { return String(body || '').replace(/\{日期\}/g, fmtDay(localDay())).replace(/\{客戶\}/g, (c?.client || '').trim() || '{客戶}'); }
function caseTplHTML(c) {
  const ts = notes().filter((n) => n.kind === 'tpl' && n.cat === c.cat).sort((a, b) => TPLGROUPS.indexOf(a.group) - TPLGROUPS.indexOf(b.group));
  return `<div class="sec"><h2>「${esc(c.cat)}」範本 <button class="linkbtn" data-act="tplfor">${ts.length ? '管理範本' : '＋ 新增範本'}</button></h2>
   ${ts.length ? ts.map((n) => `<details class="tplrow"><summary><span class="pill p-go">${esc(n.group)}</span><span class="tpltitle">${esc(n.title)}</span><button class="btn" data-copy="${n.id}" data-forcase="1" style="padding:4px 10px;font-size:13px">複製</button></summary><pre>${esc(fillTpl(n.body, c))}</pre></details>`).join('') : `<div class="sub">還沒有和「${esc(c.cat)}」相關的 LINE、簡訊範本。</div>`}</div>`;
}
function docTile(d, i, p, folders) { // p：'' 案件、'c' 客戶
  const media = isPdf(d) ? `<div class="pdftile">PDF<small>${esc(d.name)}</small></div>` : `<canvas data-${p}thumb="${i}" width="240" height="320"></canvas>`;
  const mv = folders ? `<select class="mv" data-movedoc="${i}" aria-label="移到資料夾"><option value="">未分類</option>${folders.map((f) => `<option${f === d.folder ? ' selected' : ''}>${esc(f)}</option>`).join('')}</select>` : '';
  const dt = `<input type="date" class="ddate" data-${p}docdate="${i}" value="${esc(d.date || '')}" aria-label="文件日期">`;
  return `<div class="doc" data-di="${i}" data-dp="${p}">${media}${dt}<div class="bar"><button class="docname" data-${p}rename="${i}" title="改名">${esc(d.name)}${d.masks?.length ? ' · 已遮 ' + d.masks.length : ''}</button><span><button data-${p}view="${i}" hidden>看</button>${isPdf(d) ? '' : `<button data-${p}mask="${i}">遮罩</button>`}<button data-${p}deldoc="${i}" aria-label="刪除文件">刪</button></span></div>${mv}</div>`;
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
   ${hasImg ? '<div class="actions"><button class="btn" data-act="pdf">圖片合成 PDF</button><button class="btn" data-act="share">匯出分享版</button></div>' : ''}</div>`;
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
function refCases() { return cases().filter((c) => c.isRef && (S.refItem === '全部' || c.items.includes(S.refItem)) && (S.refCause === '全部' || c.causeType === S.refCause)).sort(byDate); }
const refTitle = (c) => c.diag || c.causeDetail || c.causeType || '未分類';
function refListHTML() {
  const list = refCases(); const used = [...new Set(cases().filter((c) => c.isRef).flatMap((c) => c.items))];
  return `<section class="col"><div class="colhead"><h1>理賠參考</h1>
  <div class="sub">險種</div><div class="chips">${['全部', ...used].map((x) => `<button class="chip${S.refItem === x ? ' on' : ''}" data-ritem="${x}">${x}</button>`).join('')}</div>
  <div class="sub">原因</div><div class="chips">${['全部', '疾病', '意外'].map((x) => `<button class="chip${S.refCause === x ? ' on' : ''}" data-rcause="${x}">${x}</button>`).join('')}</div></div>
  <div class="list">${list.length ? list.map((c) => `<button class="item${c.id === S.refSel ? ' on' : ''}" data-rsel="${c.id}"><span class="t"><span>${esc(refTitle(c))}</span></span><span class="s">${esc(c.items.join('、') || '未選險種')} · ${esc(c.causeType)}${c.causeDetail ? '›' + esc(c.causeDetail) : ''}</span></button>`).join('') : '<div class="empty">還沒有理賠參考。在理賠案件裡按「做成理賠參考」就會出現在這裡。</div>'}</div></section>`;
}
function refDetailHTML() {
  const c = cases().find((x) => x.id === S.refSel && x.isRef);
  if (!c) return '<div class="empty">選一則理賠參考。這裡不會顯示客戶姓名。</div>';
  const cost = +c.cost || 0, ratio = +c.ratio || 0, auto = cost && ratio ? Math.round(cost * ratio / 100) : 0;
  return `<button class="back" data-act="back">‹ 理賠參考</button>
  <div class="dhead"><h1 style="margin:0;font-size:20px">${esc(refTitle(c))}</h1></div>
  <div class="sec"><h2>分類</h2><div class="reftop">${c.items.map((x) => `<span class="pill p-go">${esc(x)}</span>`).join('')}${c.causeType ? `<span class="pill p-wait">${esc(c.causeType)}${c.causeDetail ? ' › ' + esc(c.causeDetail) : ''}</span>` : ''}</div></div>
  ${cost ? `<div class="sec"><h2>理賠金額</h2><div class="calcline">花費 ${cost.toLocaleString()} × ${ratio || '—'}% = ${auto ? auto.toLocaleString() : '—'}${c.payout ? `　實際理賠 ${(+c.payout).toLocaleString()}` : ''}</div></div>` : ''}
  ${c.docs.length ? `<div class="sec"><h2>文件（遮罩版）</h2><div class="docs">${c.docs.map((d, i) => (isPdf(d) ? '' : `<div class="doc"><canvas data-rthumb="${i}" width="240" height="320"></canvas></div>`)).join('')}</div><div class="actions"><button class="btn pri" data-act="rshare">匯出分享版</button></div></div>` : ''}
  <div class="note">這頁只顯示分類、計算和遮罩後的文件，不顯示客戶姓名。要修改內容請回到「案件」。</div>
  <div class="actions"><button class="btn" data-act="gocase">打開原案件</button></div>`;
}
function notesHTML(kind) {
  const groups = kind === 'sop' ? SOPGROUPS : TPLGROUPS;
  const list = notes().filter((n) => n.kind === kind && (kind !== 'sop' || S.sopGroup === '全部' || n.group === S.sopGroup) && (kind !== 'tpl' || ((S.tplGroup === '全部' || n.group === S.tplGroup) && (S.tplCat === '全部' || (S.tplCat === '通用' ? !n.cat : n.cat === S.tplCat)))));
  const allT = notes().filter((n) => n.kind === 'tpl'), usedCats = CATS.filter((k) => allT.some((n) => n.cat === k));
  const q = S.q.trim(); const shown = list.filter((n) => !q || (n.title + n.body).includes(q));
  return `<h1 style="display:flex;justify-content:space-between;align-items:center;gap:8px">${kind === 'sop' ? 'SOP 作業流程' : '話術範本'}<button class="btn" data-act="import">批次匯入</button></h1>
  <input class="search" id="q" placeholder="搜尋" value="${esc(S.q)}" style="max-width:420px" autocomplete="off">
  ${kind === 'sop' ? `<div class="chips">${['全部', ...groups].map((g) => `<button class="chip${S.sopGroup === g ? ' on' : ''}" data-sg="${g}">${g}</button>`).join('')}</div>` : ''}
  ${kind === 'tpl' ? `<div class="chips">${['全部', ...groups].map((g) => `<button class="chip${S.tplGroup === g ? ' on' : ''}" data-tg="${g}">${g}</button>`).join('')}</div>
   <div class="chips"><span class="sub" style="align-self:center">相關案件</span>${['全部', '通用', ...usedCats].map((g) => `<button class="chip${S.tplCat === g ? ' on' : ''}" data-tc="${g}">${g}</button>`).join('')}</div>` : ''}
  <div class="cards">${shown.map((n) => `<div class="card"><h3><span>${esc(n.title)}</span><span class="sub">${esc(n.group)}${n.cat ? ' · ' + esc(n.cat) : ''}</span></h3><pre>${linkify(esc(n.body))}</pre>${nimgsHTML(n)}<div class="mini">${kind === 'tpl' || /https?:\/\//.test(n.body) ? `<button data-copy="${n.id}">複製</button>` : ''}<button data-editnote="${n.id}">編輯</button><button data-delnote="${n.id}">刪除</button></div></div>`).join('') || `<div class="empty">${kind === 'sop' ? '還沒有流程。用下面的表單把 Notion、LINE 記事本裡的步驟加進來。' : '還沒有範本。'}</div>`}</div>
  <div class="sec" style="max-width:620px"><h2>新增${kind === 'sop' ? '流程' : '範本'}</h2>
   <div class="grid2"><label class="f">分類<select id="nGroup">${groups.map((g) => `<option${(kind === 'sop' ? S.sopGroup : S.tplGroup) === g ? ' selected' : ''}>${g}</option>`).join('')}</select></label>${kind === 'tpl' ? `<label class="f">相關案件<select id="nCat"><option value="">通用</option>${CATS.map((k) => `<option${S.tplCat === k ? ' selected' : ''}>${k}</option>`).join('')}</select></label>` : ''}<label class="f">標題<input id="nTitle"></label></div>
   ${kind === 'tpl' ? '<div class="sub">內容裡可以寫 {客戶}、{日期}，從案件複製時會自動換成客戶名字和今天日期。</div>' : ''}
   <label class="f">內容<textarea id="nBody" rows="4"></textarea></label>
   <div class="f" style="font-size:13px;color:var(--muted)">附圖（可多張；要拍照拉正，新增後按「編輯 › 拍照／掃描」）<div class="actions" style="align-items:center"><label class="btn">＋ 選圖片<input type="file" id="nImgs" accept="image/*" multiple></label><span class="sub" id="nImgsN"></span></div></div>
   <div class="actions"><button class="btn pri" data-addnote="${kind}">新增</button></div></div>`;
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
  <div class="card"><h3>同步</h3>${syncHTML()}<div class="sub">資料加密後存在 OneDrive › 應用程式 › Work › ${esc(FOLDER)}。iPad 和手機登入同一個帳號就會自動同步，不需要備份。</div>
   <div class="actions"><button class="btn" data-act="import">批次匯入</button><button class="btn" data-act="exportnotes">複製新契約備註（去掉姓名）</button><button class="btn" data-act="lock">鎖定</button><button class="btn" data-act="signout">登出</button></div></div>
  </div>`;
}

/* ---------- 縮圖 ---------- */
async function drawThumbTo(cv, d, solid) {
  if (isPdf(d)) return;
  const im = await getImg(d.fileId); if (!im || !cv.isConnected) return;
  const x = cv.getContext('2d'); const s = Math.max(cv.width / im.width, cv.height / im.height); const w = im.width * s, h = im.height * s;
  x.drawImage(im, (cv.width - w) / 2, (cv.height - h) / 2, w, h);
  x.fillStyle = solid ? '#111' : 'rgba(17,17,17,.82)';
  d.masks.forEach((m) => x.fillRect((cv.width - w) / 2 + m.x * w, (cv.height - h) / 2 + m.y * h, m.w * w, m.h * h));
}
function nimgsHTML(n) { return n.images?.length ? `<div class="nimgs">${n.images.map((f) => `<button class="nimg" data-nimg="${f}" data-ntitle="${esc(n.title)}" aria-label="放大看圖"><img data-fid="${f}" alt=""></button>`).join('')}</div>` : ''; }
function drawThumbs() {
  const ccl = curClient(); document.querySelectorAll('[data-cthumb]').forEach((cv) => { const d = ccl?.docs[+cv.dataset.cthumb]; if (d) drawThumbTo(cv, d, false); });
  document.querySelectorAll('img[data-fid]').forEach(async (im) => { const u = await Files.url(im.dataset.fid).catch(() => null); if (u && im.isConnected) im.src = u; });
  const c = cur(); document.querySelectorAll('[data-thumb]').forEach((cv) => { const d = c?.docs[+cv.dataset.thumb]; if (d) drawThumbTo(cv, d, false); });
  const r = Vault.get('cases', S.refSel); document.querySelectorAll('[data-rthumb]').forEach((cv) => { const d = r?.docs[+cv.dataset.rthumb]; if (d) drawThumbTo(cv, d, true); });
}

/* ---------- 事件 ---------- */
function newCase(cat) { const c = saveCase({ date: localDay(), client: '', title: '', cat, status: '進行中', items: [], causeType: '', causeDetail: '', diag: '', next: '', cost: '', ratio: '', payout: '', notes: '', docs: [], isRef: false }); S.sel = c.id; S.showDetail = true; return c; }
function armConfirm(b, label) { if (b.dataset.confirm) return true; b.dataset.confirm = 1; b.textContent = label; setTimeout(() => { if (b.isConnected) { delete b.dataset.confirm; App.render(); } }, 4000); return false; }
app.addEventListener('click', async (e) => {
  const gn = e.target.closest('[data-gonode]'); if (gn) { S.view = 'clients'; S.csel = gn.dataset.gonode; S.ccat = '全部'; S.showDetail = true; S.showIds = false; App.render(); return; }
  const tileImg = e.target.closest('.doc[data-di] canvas, .doc[data-di] .pdftile'); // 點圖片＝看
  if (tileImg) { tileImg.closest('.doc').querySelector('[data-view],[data-cview]')?.click(); return; }
  const b = e.target.closest('button'); if (!b) return; const d = b.dataset; const c = cur();
  if (d.nav) { S.view = d.nav; S.showDetail = false; S.q = ''; App.render(); return; }
  if (d.cat) { S.cat = d.cat; App.render(); return; }
  if (d.sel) { S.folder = '全部'; S.sel = d.sel; S.showDetail = true; App.render(); return; }
  if (d.rsel) { S.refSel = d.rsel; S.showDetail = true; App.render(); return; }
  if (d.ritem) { S.refItem = d.ritem; App.render(); return; }
  if (d.rcause) { S.refCause = d.rcause; App.render(); return; }
  if (d.sg) { S.sopGroup = d.sg; App.render(); return; }
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
  if (d.ksel) { S.ksel = d.ksel; S.kres = ''; S.showDetail = true; App.render(); return; }
  if (d.gocall) { S.view = 'calls'; S.ksel = d.gocall; S.kres = ''; S.showDetail = true; App.render(); return; }
  if (d.kst) { S.kst = d.kst; App.render(); return; }
  if (d.kres) { const note = $('kNote')?.value || '', next = $('kNext')?.value || ''; S.kres = S.kres === d.kres ? '' : d.kres; App.render(); $('kNote').value = note; $('kNext').value = next; return; }
  if (d.kdel) { const k = curCall(); if (k && armConfirm(b, '刪？')) { k.logs = (k.logs || []).filter((x) => x.id !== d.kdel); Vault.put('calls', k); App.render(); } return; }
  if (d.ccat) { S.ccat = d.ccat; App.render(); return; }
  if (d.csel) { S.ccat = '全部'; S.csel = d.csel; S.showDetail = true; S.showIds = false; App.render(); return; }
  if (d.reveal) { const el = $(d.reveal); if (el) { const on = el.classList.toggle('shown'); b.textContent = on ? '隱藏' : '顯示'; } return; }
  const ccl = curClient();
  if (d.cview && ccl) { openViewer(ccl.docs[+d.cview]); return; }
  if (d.cmask && ccl) { openMask(ccl, +d.cmask, 'clients'); return; }
  if (d.cdeldoc && ccl) { if (armConfirm(b, '確定？')) { const [doc] = ccl.docs.splice(+d.cdeldoc, 1); Vault.put('clients', ccl); if (doc) Files.remove(doc.fileId); App.render(); } return; }
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
  if (d.item && c) { const i = c.items.indexOf(d.item); i < 0 ? c.items.push(d.item) : c.items.splice(i, 1); saveCase(c); App.render(); return; }
  if (d.ctype && c) { c.causeType = c.causeType === d.ctype ? '' : d.ctype; c.causeDetail = ''; saveCase(c); App.render(); return; }
  if (d.cdet && c) { c.causeDetail = c.causeDetail === d.cdet ? '' : d.cdet; saveCase(c); App.render(); return; }
  if (d.mask && c) { openMask(c, +d.mask); return; }
  if (d.view && c) { openViewer(c.docs[+d.view]); return; }
  if (d.deldoc && c) { if (armConfirm(b, '確定？')) { const [doc] = c.docs.splice(+d.deldoc, 1); saveCase(c); if (doc) Files.remove(doc.fileId); App.render(); } return; }
  if (d.copy) { e.preventDefault(); const n = Vault.get('notes', d.copy); const txt = fillTpl(n.body, d.forcase ? c : null); try { await navigator.clipboard.writeText(txt); toast('已複製' + (d.forcase && c?.client ? `（已帶入 ${c.client}）` : '')); } catch (err) { toast('無法自動複製，請長按文字選取'); } return; }
  if (d.tg) { S.tplGroup = d.tg; App.render(); return; }
  if (d.tc) { S.tplCat = d.tc; App.render(); return; }
  if (d.nimg) { openViewer({ fileId: d.nimg, name: d.ntitle || '附圖' }); return; }
  if (d.delnote) { if (armConfirm(b, '確定刪除？')) { (Vault.get('notes', d.delnote)?.images || []).forEach((f) => Files.remove(f)); Vault.del('notes', d.delnote); App.render(); } return; }
  if (d.editnote) { editNote(d.editnote); return; }
  if (d.addnote) {
    const t = $('nTitle').value.trim(), bd = $('nBody').value.trim(), imgs = [...($('nImgs')?.files || [])];
    if (!t && !bd && !imgs.length) { toast('先填標題、內容或附圖'); return; }
    const n = Vault.put('notes', { kind: d.addnote, group: $('nGroup').value, cat: $('nCat')?.value || '', title: t || '未命名', body: bd, images: [] });
    if (imgs.length) { b.disabled = true; await addNoteImages(n, imgs); }
    App.render(); toast('已新增'); return;
  }
  switch (d.act) {
    case 'newclient': { const n = Vault.put('clients', { name: '', docs: [] }); S.csel = n.id; S.showDetail = true; S.showIds = false; S.view = 'clients'; App.render(); setTimeout(() => $('cf_name')?.focus(), 50); break; }
    case 'cscan': scanForClient(ccl); break;
    case 'toggleids': S.showIds = !S.showIds; App.render(); break;
    case 'cpdf': makePDF(ccl.docs, false, (ccl.name || '客戶') + '-證件'); break;
    case 'cshare': openShare(ccl, (ccl.name || '客戶') + '-證件'); break;
    case 'ccase': { const c = newCase(S.ccat !== '全部' ? S.ccat : '理賠'); c.client = ccl.name; c.clientId = ccl.id; saveCase(c); S.view = 'cases'; App.render(); break; }
    case 'delclient': if (armConfirm(b, '再按一次確定刪除（證件照片也會刪掉）')) { (ccl.docs || []).forEach((doc) => Files.remove(doc.fileId)); cases().filter((c) => c.clientId === ccl.id).forEach((c) => { c.clientId = ''; saveCase(c); }); Vault.all('rels').filter((r) => r.a === ccl.id || r.b === ccl.id).forEach((r) => Vault.del('rels', r.id)); Vault.del('clients', ccl.id); S.csel = null; S.showDetail = false; App.render(); } break;
    case 'newcall': { const k = Vault.put('calls', { name: '', phone: '', source: '', status: '待撥', nextDate: '', logs: [] }); S.ksel = k.id; S.kres = ''; S.showDetail = true; S.view = 'calls'; App.render(); setTimeout(() => $('kf_name')?.focus(), 50); break; }
    case 'importcalls': openCallImport(); break;
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
    case 'tplfor': S.view = 'tpl'; S.tplCat = c.cat; S.tplGroup = '全部'; S.q = ''; S.showDetail = false; App.render(); break;
    case 'adddisc': if (ccl) openDisclosure(ccl, null, ''); break;
    case 'togglehealth': S.showHealth = !S.showHealth; App.render(); break;
    case 'casedisc': { const ppl = casePeople(c); if (ppl.length > 1) openWhoPicker(c, ppl); else { const cl2 = Vault.get('clients', c.clientId); if (cl2) openDisclosure(cl2, null, c.id); } break; }
    case 'addperson': openPersonPicker(c); break;
    case 'crmimport': if (ccl) openCrmImport(ccl); break;
    case 'crmimportnew': openCrmImport(null); break;
    case 'addrel': if (ccl) openRelPicker(ccl); break;
    case 'relgraph': openRelGraph(); break;
    case 'exportnotes': {
      const names = clients().map((x) => x.name).filter((n) => n && n.length >= 2).sort((a2, b2) => b2.length - a2.length);
      const clean = (t) => { let o = String(t || ''); names.forEach((n) => { o = o.split(n).join('○○'); }); return o.replace(/[A-Z][12]\d{8}/g, '［身分證］').replace(/09\d{2}[-\s]?\d{3}[-\s]?\d{3}/g, '［電話］').replace(/[\w.+-]+@[\w.-]+/g, '［Email］'); };
      const list = cases().filter((x) => ['新契約', '建議書'].includes(x.cat) && ((x.notes || '').trim() || (x.timeline || []).some((e) => e.note))).sort(byDate);
      const txt = list.map((x, i) => `#${i + 1}【${x.cat}】${clean(x.title)}\n${clean(x.notes)}${(x.timeline || []).filter((e) => e.note).map((e) => `\n・${e.title}：${clean(e.note)}`).join('')}`).join('\n\n---\n\n');
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
    case 'pasteimg': pasteFromClipboard(); break;
    case 'editsteps': openStepEditor(c.cat); break;
    case 'linkcase': openLinkPicker(c); break;
    case 'convert': {
      const nc = saveCase({ date: localDay(), client: c.client, clientId: c.clientId || '', title: c.title, cat: '新契約', status: '進行中', items: [], causeType: '', causeDetail: '', diag: '', next: '送件', cost: '', ratio: '', payout: '', notes: '', docs: [], isRef: false,
        people: JSON.parse(JSON.stringify(c.people || [])), links: [c.id], fromProposal: c.id,
        timeline: [{ id: uid(), date: localDay(), title: '簽約', note: '', at: new Date().toISOString() }] });
      c.links = [...new Set([...(c.links || []), nc.id])]; c.outcome = '成交'; c.status = '已結案'; saveCase(c);
      S.sel = nc.id; S.folder = '全部'; if (S.cat === '建議書') S.cat = '新契約'; App.render(); toast('已建立新契約，記得記錄告知資料'); break;
    }
    case 'newfolder': { const name = (prompt('資料夾名稱（例如：體檢報告、基本資料、收據）') || '').trim(); if (!name) break; c.folders = c.folders || []; if (!c.folders.includes(name)) c.folders.push(name); saveCase(c); S.folder = name; App.render(); break; }
    case 'renfolder': { const old = S.folder; const name = (prompt('新的資料夾名稱', old) || '').trim(); if (!name || name === old) break; c.folders = c.folders.map((f) => (f === old ? name : f)); c.docs.forEach((doc) => { if (doc.folder === old) doc.folder = name; }); saveCase(c); S.folder = name; App.render(); break; }
    case 'delfolder': if (armConfirm(b, '確定刪除？裡面的文件會移到未分類')) { const old = S.folder; c.folders = c.folders.filter((f) => f !== old); c.docs.forEach((doc) => { if (doc.folder === old) doc.folder = ''; }); saveCase(c); S.folder = '全部'; App.render(); } break;
    case 'scan': { if (S.view === 'clients' && ccl) { scanForClient(ccl); break; } let t = c; if (!t || S.view !== 'cases') { t = newCase('理賠'); S.view = 'cases'; } openScanner(t, { folder: curFolder(t) }); break; }
    case 'newcase': newCase(S.cat === '全部' ? '理賠' : S.cat); App.render(); break;
    case 'back': S.showDetail = false; App.render(); break;
    case 'ref': c.isRef = !c.isRef; saveCase(c); App.render(); toast(c.isRef ? '已加入理賠參考（不顯示客戶姓名）' : '已從理賠參考移除'); break;
    case 'delcase': if (armConfirm(b, '再按一次確定刪除')) { c.docs.forEach((doc) => Files.remove(doc.fileId)); (c.links || []).forEach((id) => { const o = Vault.get('cases', id); if (o) { o.links = (o.links || []).filter((x) => x !== c.id); saveCase(o); } }); Vault.del('cases', c.id); S.sel = null; S.showDetail = false; App.render(); } break;
    case 'pdf': makePDF(c.docs, false, (c.client || '案件') + '-' + (c.title || '文件')); break;
    case 'share': openShare(c); break;
    case 'rshare': openShare(Vault.get('cases', S.refSel)); break;
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
app.addEventListener('change', (e) => {
  if (e.target.id === 'kScript') { S.script = e.target.value; App.render(); return; }
  const kf = e.target.dataset?.kf; if (kf) { const k = curCall(); if (!k) return; k[kf] = kf === 'name' ? e.target.value.trim() : e.target.value; Vault.put('calls', k); if (kf !== 'note') App.render(); return; }
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
app.addEventListener('change', (e) => { if (e.target.dataset?.f === 'client') { const c = cur(); if (!c) return; c.client = e.target.value; const made = linkClient(c); saveCase(c); App.render(); if (made) toast(`已建立客戶「${made.name}」`); return; } if (e.target.id === 'nImgs') { const k = e.target.files.length; $('nImgsN').textContent = k ? `已選 ${k} 張` : ''; return; } const f = e.target.dataset?.f; const c = cur(); if (!f || !c) return; c[f] = e.target.value; saveCase(c); if (['cat', 'status', 'cost', 'ratio', 'client', 'title', 'date'].includes(f)) App.render(); });
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
function editNote(id, isNew) {
  const n = Vault.get('notes', id); const withImgs = n.kind !== 'quick';
  NE = withImgs ? n : null;
  const imgsBlock = () => `<div class="sub">附圖 · 可以把圖片拖進來，或按 Ctrl+V 貼上截圖</div><div class="nimgs edit">${(n.images || []).map((f, i) => `<div class="nimgwrap"><button class="nimg" data-vimg="${f}" aria-label="放大看圖"><img data-fid="${f}" alt=""></button><button class="nimgdel" data-delimg="${i}" aria-label="移除這張圖">✕</button></div>`).join('')}<button class="adddoc nimgadd" id="eCam">${ICON.scan}拍照／掃描</button><button class="adddoc nimgadd" id="eAddImg">＋ 選圖片</button></div><input type="file" id="eImgs" accept="image/*" multiple>`;
  ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>${n.kind === 'quick' ? '電話' : '編輯'}</h2><button class="btn" id="eCancel">取消</button><button class="btn pri" id="eSave">儲存</button></div><div class="ovbody" style="flex-direction:column;max-width:720px;margin:0 auto;width:100%">${n.kind === 'sop' || n.kind === 'tpl' ? `<label class="f">分類<select id="eG">${(n.kind === 'sop' ? SOPGROUPS : TPLGROUPS).map((g) => `<option${g === n.group ? ' selected' : ''}>${g}</option>`).join('')}</select></label>` : ''}${n.kind === 'tpl' ? `<label class="f">相關案件<select id="eC"><option value="">通用</option>${CATS.map((k) => `<option${k === n.cat ? ' selected' : ''}>${k}</option>`).join('')}</select></label><div class="sub">可以寫 {客戶}、{日期}，從案件複製時會自動帶入。</div>` : ''}<label class="f">${n.kind === 'quick' ? '名稱' : '標題'}<input id="eT" value="${esc(n.title)}"></label>${n.kind === 'quick' ? `<label class="f">電話<input id="eB" inputmode="tel" value="${esc(n.body)}"></label>` : `<label class="f">內容<textarea id="eB" rows="10">${esc(n.body)}</textarea></label><div id="eImgBox">${imgsBlock()}</div>`}${n.kind === 'quick' && !isNew ? '<div class="actions"><button class="btn danger" id="eDel">刪除這支電話</button></div>' : ''}</div></div>`;
  const close = () => { NE = null; ov.innerHTML = ''; App.render(); };
  const wireImgs = () => {
    if (!withImgs) return;
    $('eImgBox').innerHTML = imgsBlock(); drawThumbs();
    $('eAddImg').onclick = () => $('eImgs').click();
    $('eCam').onclick = () => { // 先把正在打的字存起來，掃描完再回到編輯
      n.title = $('eT').value.trim(); n.body = $('eB').value; if ($('eG')) n.group = $('eG').value; if ($('eC')) n.cat = $('eC').value; delete n.refresh; Vault.put('notes', n);
      NE = null;
      openScanner(n, {
        onDone: async (pages) => { for (const p of pages) { n.images = [...(n.images || []), await Files.add(p.blob, n.title || p.name)]; URL.revokeObjectURL(p.url); } if (pages.length) Vault.put('notes', n); editNote(id, isNew); },
        onCancel: () => editNote(id, isNew),
      });
    };
    $('eImgs').onchange = async (e) => { const fs = [...e.target.files]; e.target.value = ''; toast('加入中…'); await addNoteImages(n, fs); wireImgs(); };
    $('eImgBox').querySelectorAll('[data-vimg]').forEach((b) => (b.onclick = () => openViewer({ fileId: b.dataset.vimg, name: n.title || '附圖' }, true)));
    $('eImgBox').querySelectorAll('[data-delimg]').forEach((b) => (b.onclick = () => { const [f] = n.images.splice(+b.dataset.delimg, 1); Files.remove(f); Vault.put('notes', n); wireImgs(); }));
  };
  NE && (NE.refresh = wireImgs);
  wireImgs();
  $('eCancel').onclick = () => { if (isNew) { (n.images || []).forEach((f) => Files.remove(f)); Vault.del('notes', id); } close(); };
  $('eSave').onclick = () => { n.title = $('eT').value.trim() || '未命名'; n.body = $('eB').value; if ($('eG')) n.group = $('eG').value; if ($('eC')) n.cat = $('eC').value; delete n.refresh; Vault.put('notes', n); close(); };
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
    group: x.kind === 'sop' ? (SOPGROUPS.includes(x.group) ? x.group : '壽險·其他') : x.kind === 'tpl' ? (TPLGROUPS.includes(x.group) ? x.group : '其他') : '電話',
    title: String(x.title || '未命名').slice(0, 200), body: String(x.body || ''), cat: x.kind === 'tpl' && CATS.includes(x.cat) ? x.cat : '',
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
function openScanner(c, opts = {}) { SC = { c, folder: opts.folder || '', coll: opts.coll || 'cases', names: opts.names || CASE_DOC_NAMES, pages: [], img: null, corners: null, filter: opts.filter || 'gray', drag: -1, queue: [], onDone: opts.onDone, onCancel: opts.onCancel }; renderScanner(); }
function scanForClient(cl) { if (!cl) { toast('先選一位客戶'); return; } openScanner(cl, { coll: 'clients', names: CLIENT_DOC_NAMES, filter: 'color' }); }
function renderScanner() {
  ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>掃描文件</h2><span class="sub">${SC.pages.length} 頁</span><button class="btn" id="scCancel">取消</button><button class="btn pri" id="scDone">完成（${SC.pages.length}）</button></div>
  <div class="ovbody">${SC.img ? `<div class="stage"><canvas id="scCv"></canvas></div>
  <div class="side"><div class="note">拖曳四個角對齊文件邊緣，再按「加入這頁」。</div>
   <div><div class="sub" style="margin-bottom:4px">效果</div><div class="seg">${[['color', '原色'], ['gray', '灰階增強'], ['bw', '黑白']].map(([k, l]) => `<button data-flt="${k}" class="${SC.filter === k ? 'on' : ''}">${l}</button>`).join('')}</div></div>
   <label class="f">文件名稱<input id="scName" value="${SC.names[Math.min(SC.pages.length, SC.names.length - 1)]}"></label>
   <div class="actions"><button class="btn pri" id="scAdd">加入這頁</button><button class="btn" id="scRetake">重拍</button><button class="btn" id="scFull">用整張圖</button></div>
   ${SC.pages.length ? `<div class="pages">${SC.pages.map((p) => `<img src="${p.url}" alt="${esc(p.name)}">`).join('')}</div>` : ''}</div>`
    : `<div class="pick">${ICON.scan}<div>拍一張文件、從相簿選照片，<br>或把圖片拖進來、按 Ctrl+V 貼上截圖</div><div class="actions" style="justify-content:center"><button class="btn pri" id="scCam">開相機</button><button class="btn" id="scLib">從相簿選</button></div>${SC.pages.length ? `<div class="pages">${SC.pages.map((p) => `<img src="${p.url}" alt="">`).join('')}</div>` : ''}</div>`}
  </div></div>`;
  $('scCancel').onclick = () => { const s = SC; s.pages.forEach((p) => URL.revokeObjectURL(p.url)); ov.innerHTML = ''; SC = null; if (s.onCancel) s.onCancel(); else App.render(); };
  $('scDone').onclick = async () => {
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
  $('scRetake').onclick = () => { SC.img = null; nextQueued(); };
  $('scFull').onclick = () => { const w = SC.img.width, h = SC.img.height; SC.corners = [[0, 0], [w, 0], [w, h], [0, h]]; renderScanner(); };
  $('scAdd').onclick = async () => { const blob = await warp(SC.img, SC.corners, SC.filter); SC.pages.push({ name: $('scName').value || '文件', blob, url: URL.createObjectURL(blob) }); SC.img = null; nextQueued(); };
  setupCornerCanvas();
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
  if (!SC && (S.view === 'sop' || S.view === 'tpl')) { const n = Vault.put('notes', { kind: S.view, group: S.view === 'sop' ? (S.sopGroup === '全部' ? '壽險·其他' : S.sopGroup) : '其他', title: '', body: '', images: [] }); editNote(n.id, true); toast('加入中…'); addNoteImages(n, imgs).then(() => NE?.refresh?.()); return; }
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
  e.preventDefault(); if (DSC) { discAddFiles(files, true); return; } if (CRMI) { crmFromImage(files[0]); return; } takeImages(files);
});
function nextQueued() { if (SC.queue.length) loadFile(SC.queue.shift()); else renderScanner(); }
function onFiles(files) { if (!SC || !files.length) return; SC.queue.push(...[...files].slice(1)); loadFile(files[0]); }
fileIn.onchange = () => { onFiles(fileIn.files); fileIn.value = ''; };
fileLib.onchange = () => { onFiles(fileLib.files); fileLib.value = ''; };
async function loadFile(f) {
  const bmp = await createImageBitmap(f).catch(() => null);
  if (!bmp) { toast('這張圖片打不開，換一張試試'); return; }
  const s = Math.min(1, 2000 / Math.max(bmp.width, bmp.height));
  const cv = document.createElement('canvas'); cv.width = Math.round(bmp.width * s); cv.height = Math.round(bmp.height * s); cv.getContext('2d').drawImage(bmp, 0, 0, cv.width, cv.height);
  SC.img = cv; const w = cv.width, h = cv.height, m = 0.06; SC.corners = [[w * m, h * m], [w * (1 - m), h * m], [w * (1 - m), h * (1 - m)], [w * m, h * (1 - m)]];
  renderScanner();
}
function setupCornerCanvas() {
  const cv = $('scCv'); cv.width = SC.img.width; cv.height = SC.img.height; const x = cv.getContext('2d'); const R = Math.max(cv.width, cv.height) / 40;
  const draw = () => {
    x.drawImage(SC.img, 0, 0); const k = SC.corners;
    x.fillStyle = 'rgba(20,22,16,.45)'; x.beginPath(); x.rect(0, 0, cv.width, cv.height); x.moveTo(...k[0]); x.lineTo(...k[3]); x.lineTo(...k[2]); x.lineTo(...k[1]); x.closePath(); x.fill('evenodd');
    x.strokeStyle = '#9DBA88'; x.lineWidth = R / 4; x.beginPath(); k.forEach((p, i) => (i ? x.lineTo(...p) : x.moveTo(...p))); x.closePath(); x.stroke();
    k.forEach((p) => { x.beginPath(); x.arc(p[0], p[1], R, 0, 7); x.fillStyle = 'rgba(157,186,136,.35)'; x.fill(); x.lineWidth = R / 5; x.stroke(); });
  };
  draw();
  const pt = (e) => { const r = cv.getBoundingClientRect(); return [(e.clientX - r.left) * cv.width / r.width, (e.clientY - r.top) * cv.height / r.height]; };
  cv.onpointerdown = (e) => { const p = pt(e); let best = -1, bd = R * 3.5; SC.corners.forEach((c, i) => { const d = Math.hypot(c[0] - p[0], c[1] - p[1]); if (d < bd) { bd = d; best = i; } }); SC.drag = best; if (best >= 0) cv.setPointerCapture(e.pointerId); };
  cv.onpointermove = (e) => { if (SC.drag < 0) return; const p = pt(e); SC.corners[SC.drag] = [Math.max(0, Math.min(cv.width, p[0])), Math.max(0, Math.min(cv.height, p[1]))]; draw(); };
  cv.onpointerup = cv.onpointercancel = () => (SC.drag = -1);
}
function solveH(dst, src) {
  const A = [];
  for (let i = 0; i < 4; i++) { const [u, v] = dst[i], [x, y] = src[i]; A.push([u, v, 1, 0, 0, 0, -u * x, -v * x, x]); A.push([0, 0, 0, u, v, 1, -u * y, -v * y, y]); }
  for (let c = 0; c < 8; c++) { let p = c; for (let r = c + 1; r < 8; r++) if (Math.abs(A[r][c]) > Math.abs(A[p][c])) p = r; [A[c], A[p]] = [A[p], A[c]]; for (let r = 0; r < 8; r++) { if (r === c) continue; const f = A[r][c] / A[c][c]; for (let k = c; k < 9; k++) A[r][k] -= f * A[c][k]; } }
  return A.map((r, i) => r[8] / r[i]);
}
function warp(img, k, filter) {
  const d = (a, b) => Math.hypot(a[0] - b[0], a[1] - b[1]);
  let W = Math.max(d(k[0], k[1]), d(k[3], k[2])), H = Math.max(d(k[0], k[3]), d(k[1], k[2])); const s = Math.min(1, 1800 / Math.max(W, H)); W = Math.round(W * s); H = Math.round(H * s);
  const h = solveH([[0, 0], [W, 0], [W, H], [0, H]], k); const sd = img.getContext('2d').getImageData(0, 0, img.width, img.height).data, iw = img.width, ih = img.height;
  const out = document.createElement('canvas'); out.width = W; out.height = H; const ox = out.getContext('2d'); const od = ox.createImageData(W, H), o = od.data;
  let sum = 0;
  for (let v = 0; v < H; v++) for (let u = 0; u < W; u++) {
    const z = h[6] * u + h[7] * v + 1; let x = (h[0] * u + h[1] * v + h[2]) / z, y = (h[3] * u + h[4] * v + h[5]) / z;
    x = Math.min(iw - 1, Math.max(0, x | 0)); y = Math.min(ih - 1, Math.max(0, y | 0));
    const si = (y * iw + x) * 4, oi = (v * W + u) * 4; o[oi] = sd[si]; o[oi + 1] = sd[si + 1]; o[oi + 2] = sd[si + 2]; o[oi + 3] = 255; sum += 0.3 * sd[si] + 0.59 * sd[si + 1] + 0.11 * sd[si + 2];
  }
  if (filter !== 'color') {
    const mean = sum / (W * H);
    for (let i = 0; i < o.length; i += 4) {
      let g = 0.3 * o[i] + 0.59 * o[i + 1] + 0.11 * o[i + 2];
      if (filter === 'gray') { g = (g - mean * 0.55) * (255 / (mean * 0.55 || 1)); g = Math.max(0, Math.min(255, g)); } else { g = g > mean * 0.82 ? 255 : Math.max(0, g * 0.35); }
      o[i] = o[i + 1] = o[i + 2] = g;
    }
  }
  ox.putImageData(od, 0, 0);
  return new Promise((ok) => out.toBlob((b) => ok(b), 'image/jpeg', 0.82));
}

/* ================= 看原圖 ================= */
async function openViewer(d, overEdit) {
  const url = await Files.url(d.fileId); if (!url) { toast('檔案讀取失敗，請確認網路'); return; }
  if (isPdf(d)) {
    ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>${esc(d.name)}</h2><button class="btn" id="vShare">分享／儲存</button><button class="btn pri" id="vClose">關閉</button></div><div class="sub" style="padding:6px 16px">如果只看到第一頁，按「分享／儲存」用「檔案」或 PDF 閱讀器打開。</div><iframe src="${url}" title="${esc(d.name)}" style="flex:1;width:100%;border:0;background:#fff"></iframe></div>`;
    $('vClose').onclick = () => (ov.innerHTML = '');
    $('vShare').onclick = async () => { const blob = await (await fetch(url)).blob(); shareOrSave([new File([blob], d.name.replace(/[\\/:*?"<>|]/g, '_') + '.pdf', { type: 'application/pdf' })]); };
    return;
  }
  if (overEdit) { const v = document.createElement('div'); v.innerHTML = `<div class="ov" style="z-index:30"><div class="ovhead"><h2>${esc(d.name)}</h2><button class="btn pri">關閉</button></div><div class="ovbody" style="justify-content:center"><img src="${url}" alt="" style="max-width:100%;height:auto;align-self:flex-start;border-radius:6px"></div></div>`; v.querySelector('button').onclick = () => v.remove(); document.body.appendChild(v); return; }
  ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>${esc(d.name)}（原圖）</h2><button class="btn pri" id="vClose">關閉</button></div><div class="ovbody" style="justify-content:center"><img src="${url}" alt="${esc(d.name)}" style="max-width:100%;height:auto;align-self:flex-start;border-radius:6px"></div></div>`;
  $('vClose').onclick = () => (ov.innerHTML = '');
}

/* ================= 遮罩 ================= */
let MK = null;
async function openMask(c, i, coll = 'cases') { const d = c.docs[i]; const im = await getImg(d.fileId); if (!im) { toast('圖片讀取失敗，請確認網路'); return; } MK = { c, coll, d, im, masks: d.masks.map((m) => ({ ...m })), sel: -1, mode: null, start: null, preview: false }; renderMask(); }
function renderMask() {
  ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>遮罩 · ${esc(MK.d.name)}</h2><button class="btn" id="mkCancel">取消</button><button class="btn pri" id="mkSave">儲存遮罩</button></div>
  <div class="ovbody"><div class="stage"><canvas id="mkCv"></canvas></div>
  <div class="side"><div class="note">在要遮的地方拖曳畫框。點一下框可以選取、拖曳移動。原圖不會被改動。</div>
  <div class="actions"><button class="btn${MK.preview ? ' on' : ''}" id="mkPrev">${MK.preview ? '回到編輯' : '看分享版效果'}</button></div>
  <div class="actions"><button class="btn danger" id="mkDel">刪除選取的框</button><button class="btn" id="mkClear">全部清除</button></div>
  <div class="sub">目前 ${MK.masks.length} 個框</div></div></div></div>`;
  $('mkCancel').onclick = () => { ov.innerHTML = ''; MK = null; };
  $('mkSave').onclick = () => { MK.d.masks = MK.masks; Vault.put(MK.coll, MK.c); ov.innerHTML = ''; MK = null; App.render(); toast('遮罩已儲存，原圖保留'); };
  $('mkPrev').onclick = () => { MK.preview = !MK.preview; renderMask(); };
  $('mkDel').onclick = () => { if (MK.sel >= 0) { MK.masks.splice(MK.sel, 1); MK.sel = -1; renderMask(); } else toast('先點選一個框'); };
  $('mkClear').onclick = () => { MK.masks = []; MK.sel = -1; renderMask(); };
  const cv = $('mkCv'); cv.width = MK.im.width; cv.height = MK.im.height; const x = cv.getContext('2d'); const W = cv.width, H = cv.height;
  const draw = (tmp) => { x.drawImage(MK.im, 0, 0); const all = tmp ? [...MK.masks, tmp] : MK.masks; all.forEach((m, i) => { if (MK.preview) { x.fillStyle = '#111'; x.fillRect(m.x * W, m.y * H, m.w * W, m.h * H); } else { x.fillStyle = 'rgba(17,17,17,.6)'; x.fillRect(m.x * W, m.y * H, m.w * W, m.h * H); x.lineWidth = Math.max(2, W / 300); x.strokeStyle = i === MK.sel ? '#E2B56A' : '#9DBA88'; x.strokeRect(m.x * W, m.y * H, m.w * W, m.h * H); } }); };
  draw(); if (MK.preview) return;
  const pt = (e) => { const r = cv.getBoundingClientRect(); return [(e.clientX - r.left) / r.width, (e.clientY - r.top) / r.height]; };
  let tmp = null;
  cv.onpointerdown = (e) => { const [px, py] = pt(e); cv.setPointerCapture(e.pointerId); const hit = MK.masks.findIndex((m) => px >= m.x && px <= m.x + m.w && py >= m.y && py <= m.y + m.h); if (hit >= 0) { MK.sel = hit; MK.mode = 'move'; MK.start = [px, py, { ...MK.masks[hit] }]; } else { MK.sel = -1; MK.mode = 'draw'; MK.start = [px, py]; } draw(); };
  cv.onpointermove = (e) => { if (!MK.mode) return; const [px, py] = pt(e); if (MK.mode === 'move') { const [sx, sy, o] = MK.start; const m = MK.masks[MK.sel]; m.x = Math.max(0, Math.min(1 - m.w, o.x + px - sx)); m.y = Math.max(0, Math.min(1 - m.h, o.y + py - sy)); draw(); } else { const [sx, sy] = MK.start; tmp = { x: Math.min(sx, px), y: Math.min(sy, py), w: Math.abs(px - sx), h: Math.abs(py - sy) }; draw(tmp); } };
  cv.onpointerup = cv.onpointercancel = () => { if (MK.mode === 'draw' && tmp && tmp.w > 0.01 && tmp.h > 0.008) { MK.masks.push(tmp); MK.sel = MK.masks.length - 1; } MK.mode = null; tmp = null; renderMask(); };
}

/* ================= 分享版 / PDF ================= */
async function renderDoc(d, mask) { // 回傳 canvas
  const im = await getImg(d.fileId); if (!im) throw new Error('圖片讀取失敗');
  const cv = document.createElement('canvas'); cv.width = im.width; cv.height = im.height; const x = cv.getContext('2d'); x.drawImage(im, 0, 0);
  if (mask) { x.fillStyle = '#111'; d.masks.forEach((m) => x.fillRect(m.x * cv.width, m.y * cv.height, m.w * cv.width, m.h * cv.height)); }
  return cv;
}
const toBlob = (cv, type = 'image/jpeg') => new Promise((ok) => cv.toBlob((b) => ok(b), type, 0.85));
async function shareOrSave(files) { // 有分享功能就跳分享（可以直接傳 LINE），沒有就下載
  try { if (navigator.canShare && navigator.canShare({ files })) { await navigator.share({ files }); return; } } catch (e) { if (e.name === 'AbortError') return; }
  files.forEach((f) => { const a = document.createElement('a'); a.href = URL.createObjectURL(f); a.download = f.name; document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 60000); });
}
async function openShare(c, pdfName) {
  if (!c) return;
  const docs = c.docs.filter((d) => !isPdf(d)); if (!docs.length) { toast('PDF 檔請在文件上按「看」再分享'); return; }
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
function enter() {
  gate.hidden = true; app.hidden = false;
  cases().filter((c) => c.client && !c.clientId).forEach((c) => { linkClient(c); saveCase(c); }); // 舊案件補連客戶
  cases().filter((c) => !c.date || c.docs.some((d) => !d.date)).forEach((c) => { c.date = caseDate(c); c.docs.forEach((d) => { if (!d.date) d.date = c.date; }); saveCase(c); }); // 舊資料補日期
  const first = cases().slice().sort((a, b) => (b.updatedAt || '').localeCompare(a.updatedAt || ''))[0];
  S.sel = first?.id || null; S.refSel = cases().find((c) => c.isRef)?.id || null; S.showDetail = false;
  lastAct = Date.now(); App.render(); Files.retryUploads();
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
