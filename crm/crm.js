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
const COLLS = ['cases', 'notes', 'files', 'clients', 'settings'];
const emptyDb = () => Object.fromEntries([['v', 1], ...COLLS.map((c) => [c, {}])]);
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
  setState(s) { this.state = s; document.querySelectorAll('.sync').forEach((el) => { el.textContent = SYNC_LABEL[s] || s; el.dataset.s = s; }); },
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
  user: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="8" r="4"/><path d="M4 21a8 8 0 0 1 16 0"/></svg>',
  msg: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4 5h16v11H9l-5 4z"/></svg>',
};
const NAV = [['quick', '快查', 'bolt'], ['clients', '客戶', 'user'], ['cases', '案件', 'folder'], ['refs', '理賠參考', 'book'], ['sop', 'SOP', 'list'], ['tpl', '範本', 'msg']];
const CATS = ['理賠', '建議書', '新契約', '保全', '保費', '保校', '房貸', '服務', '其他', '產險'];
const STATUS = ['進行中', '待回覆', '已結案'];
const ITEMS = ['住院醫療', '實支實付', '住院日額', '手術', '意外醫療', '意外失能', '骨折', '重大傷病', '癌症', '長照失能', '門診', '產險-車險', '產險-住火', '產險-旅平'];
const CAUSE = { '疾病': ['消化系統', '呼吸系統', '心血管', '骨科', '婦科', '腫瘤', '眼耳鼻喉', '其他疾病'], '意外': ['交通事故', '跌倒', '運動傷害', '職業傷害', '燒燙傷', '其他意外'] };
const SOPGROUPS = ['壽險·新契約', '壽險·保全', '壽險·保費', '壽險·保校', '壽險·理賠', '壽險·房貸', '壽險·服務', '壽險·其他', '產險'];

/* ================= 狀態與畫面 ================= */
const S = { view: 'cases', ccat: '全部', csel: null, showIds: false, sel: null, refSel: null, cat: '全部', q: '', refItem: '全部', refCause: '全部', sopGroup: '全部', showDetail: false };
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
const pillCls = (s) => (s === '已結案' ? 'p-done' : s === '待回覆' ? 'p-wait' : 'p-go');
const syncHTML = (cls = '') => `<span class="sync ${cls}" data-s="${Vault.state}">${SYNC_LABEL[Vault.state] || ''}</span>`;

const App = {
  render() {
    if (!Vault.db) return;
    const v = S.view;
    app.className = 'app' + ((v === 'cases' || v === 'refs' || v === 'clients') ? '' : ' single') + (S.showDetail ? ' showdetail' : '') + (MOCK ? ' test' : '');
    let h = MOCK ? '<div class="testbar">測試模式：資料只存在這台裝置，不會上傳 OneDrive</div>' : '';
    h += railHTML();
    if (v === 'clients') h += clientListHTML() + '<section class="detail split">' + clientDetailHTML() + '</section>';
    else if (v === 'cases') h += caseListHTML() + '<section class="detail split">' + caseDetailHTML() + '</section>';
    else if (v === 'refs') h += refListHTML() + '<section class="detail split">' + refDetailHTML() + '</section>';
    else if (v === 'sop') h += '<section class="wide">' + notesHTML('sop') + '</section>';
    else if (v === 'quick') h += '<section class="wide">' + quickHTML() + '</section>';
    else if (v === 'tpl') h += '<section class="wide">' + notesHTML('tpl') + '</section>';
    h += tabsHTML();
    const keep = document.activeElement?.id, pos = document.activeElement?.selectionStart;
    app.innerHTML = h;
    if (keep && (keep === 'q')) { const el = $(keep); if (el) { el.focus(); try { el.setSelectionRange(pos, pos); } catch (e) { } } }
    drawThumbs();
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
  const t = [NAV[0], NAV[1], ['scan', '掃描', 'scan'], NAV[2], NAV[3], NAV[4]];
  return `<nav class="tabs" aria-label="分頁">${t.map(([k, l, i]) => k === 'scan' ? `<button class="tab scan" data-act="scan"><span class="c">${ICON.scan}</span>${l}</button>` : `<button class="tab${S.view === k ? ' on' : ''}" data-nav="${k}">${ICON[i]}${l}</button>`).join('')}</nav>`;
}
function caseSummary(c) { const a = [c.cat]; if (c.items?.length) a.push(c.items.slice(0, 2).join('、')); if (c.causeType) a.push(c.causeType + (c.causeDetail ? '›' + c.causeDetail : '')); return a.join(' · '); }
function filteredCases() {
  const q = S.q.trim();
  return cases().filter((c) => (S.cat === '全部' || c.cat === S.cat) && (!q || [c.client, c.title, c.notes, c.diag, c.next].join(' ').includes(q))).sort(byDate);
}
function caseListHTML() {
  const list = filteredCases();
  return `<section class="col"><div class="colhead"><h1>案件 <button class="btn" data-act="newcase">＋ 新增</button></h1>
  <div class="mobsync">${syncHTML()}</div>
  <input class="search" id="q" placeholder="搜尋客戶、案件、備註" value="${esc(S.q)}" autocomplete="off">
  <div class="chips">${['全部', ...CATS].map((c) => `<button class="chip${S.cat === c ? ' on' : ''}" data-cat="${c}">${c}</button>`).join('')}</div></div>
  <div class="list">${list.length ? list.map((c) => `<button class="item${c.id === S.sel ? ' on' : ''}" data-sel="${c.id}"><span class="t"><span>${esc(c.client || '未填客戶')} · ${esc(c.title || '未命名')}</span><span class="pill ${pillCls(c.status)}">${c.status}</span></span><span class="s">${fmtDay(caseDate(c))} · ${esc(caseSummary(c))}${latestStep(c) ? ' · 最新：' + esc(latestStep(c).title) : ''}</span></button>`).join('') : `<div class="empty">${cases().length ? '沒有符合的案件。' : '還沒有案件。按「＋ 新增」，或按「掃描」直接拍文件建案。'}</div>`}</div></section>`;
}
function caseDetailHTML() {
  const c = cur();
  if (!c) return '<div class="empty">從左邊選一個案件，或按「掃描」直接拍文件建案。</div>';
  const isClaim = c.cat === '理賠';
  const cost = +c.cost || 0, ratio = +c.ratio || 0, auto = cost && ratio ? Math.round(cost * ratio / 100) : 0;
  return `<button class="back" data-act="back">‹ 案件列表</button>
  <div class="dhead"><input class="title" id="f_title" data-f="title" value="${esc(c.title)}" placeholder="案件名稱" aria-label="案件名稱"><span class="pill ${pillCls(c.status)}">${c.status}</span></div>
  ${c.convertedAt ? `<div class="note">由建議書轉成新契約 · ${esc(c.convertedAt)}</div>` : ''}
  ${c.cat === '建議書' ? `<div class="sec" style="flex-direction:row;align-items:center;justify-content:space-between;flex-wrap:wrap"><div><div style="font-weight:500">促約成功了嗎？</div><div class="sub">轉成新契約後，文件、備註都會保留，客戶頁的「新契約」就會出現這一件。</div></div><button class="btn pri" data-act="convert">轉成新契約</button></div>` : ''}
  <div class="sec"><h2>基本資料</h2><div class="grid2">
   <label class="f">客戶${c.clientId && Vault.get('clients', c.clientId) ? ` <button class="linkbtn" data-goclient="${c.clientId}">看客戶資料 ›</button>` : ''}<input id="f_client" data-f="client" list="clientNames" autocomplete="off" value="${esc(c.client)}" placeholder="打名字，會自動帶出或建立客戶"><datalist id="clientNames">${clients().map((x) => `<option value="${esc(x.name)}">`).join('')}</datalist></label>
   <label class="f">日期<input type="date" id="f_date" data-f="date" value="${esc(caseDate(c))}"></label>
   <label class="f">類別<select id="f_cat" data-f="cat">${CATS.map((x) => `<option${x === c.cat ? ' selected' : ''}>${x}</option>`).join('')}</select></label>
   <label class="f">狀態<select id="f_status" data-f="status">${STATUS.map((x) => `<option${x === c.status ? ' selected' : ''}>${x}</option>`).join('')}</select></label>
   <label class="f">下一步<input id="f_next" data-f="next" value="${esc(c.next)}" placeholder="例如：補診斷書正本"></label></div></div>
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
  <div class="sec"><h2>備註</h2><textarea id="f_notes" data-f="notes" rows="4" aria-label="備註">${esc(c.notes)}</textarea></div>
  <div class="actions">${isClaim ? `<button class="btn${c.isRef ? ' on' : ''}" data-act="ref">${c.isRef ? '✓ 已在理賠參考' : '做成理賠參考'}</button>` : ''}<button class="btn danger" data-act="delcase">刪除案件</button></div>`;
}
function clientCases(cl) { return cases().filter((c) => c.clientId === cl.id).sort(byDate); }
function clientListHTML() {
  const q = S.q.trim();
  const list = clients().filter((x) => !q || [x.name, x.phone, x.line, x.note].join(' ').includes(q)).sort((a, b) => (a.name || '').localeCompare(b.name || '', 'zh-Hant'));
  return `<section class="col"><div class="colhead"><h1>客戶 <button class="btn" data-act="newclient">＋ 新增</button></h1>
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
  <div class="sec"><h2>基本資料</h2><div class="grid2">
   <label class="f">性別<select id="cf_gender" data-cf="gender">${['', '女', '男'].map((g) => `<option${g === (x.gender || '') ? ' selected' : ''}>${g}</option>`).join('')}</select></label>
   ${f('birthday', '生日', 'type="date"')}
   <label class="f"><span style="display:flex;justify-content:space-between">身分證字號<button class="linkbtn" data-reveal="cf_idNo">顯示</button></span><input id="cf_idNo" data-cf="idNo" class="secret" value="${esc(x.idNo)}" autocomplete="off" autocapitalize="characters" spellcheck="false"></label>
   ${f('phone', '手機', 'inputmode="tel"')}
   ${f('line', 'LINE ID')}
   ${f('email', 'Email', 'inputmode="email"')}
   ${f('job', '職業')}
   ${f('source', '來源／介紹人')}</div>
   ${f('address', '地址')}
   <label class="f">備註<textarea id="cf_note" data-cf="note" rows="3" placeholder="家庭狀況、保單重點、喜好…">${esc(x.note)}</textarea></label></div>
  <div class="sec"><h2>證件與文件 <button class="btn" data-act="toggleids" style="padding:4px 10px;font-size:13px">${S.showIds ? '隱藏證件' : '顯示證件'}</button></h2>
   <div class="docs${S.showIds ? '' : ' veiled'}">${docs.map((d, i) => [d, i]).sort(([a], [b]) => (b.date || '').localeCompare(a.date || '')).map(([d, i]) => docTile(d, i, 'c', null)).join('')}
   <button class="adddoc" data-act="cscan">${ICON.scan}拍身分證、存摺<span style="font-size:11px">也可以拖曳圖片、PDF 進來</span></button><label class="adddoc">＋ 上傳 PDF 或圖片<input type="file" id="docUpC" accept="image/*,application/pdf,.pdf" multiple></label></div>
   ${docs.some((d) => !isPdf(d)) ? '<div class="actions"><button class="btn" data-act="cpdf">圖片合成 PDF</button><button class="btn" data-act="cshare">匯出遮罩版</button></div>' : ''}</div>
  <div class="sec"><h2>案件 <button class="btn" data-act="ccase" style="padding:4px 10px;font-size:13px">＋ 新增${S.ccat !== '全部' ? esc(S.ccat) : '案件'}</button></h2>
   ${cs.length ? `<div class="chips">${['全部', ...CATS.filter((k) => cs.some((c) => c.cat === k))].map((k) => `<button class="chip${S.ccat === k ? ' on' : ''}" data-ccat="${k}">${k} <span class="cnt">${k === '全部' ? cs.length : cs.filter((c) => c.cat === k).length}</span></button>`).join('')}</div>` : ''}
   ${cs.length ? cs.filter((c) => S.ccat === '全部' || c.cat === S.ccat).map((c) => `<button class="item" data-gocase="${c.id}"><span class="t"><span>${esc(c.title || '未命名')}</span><span class="pill ${pillCls(c.status)}">${c.status}</span></span><span class="s">${fmtDay(caseDate(c))} · ${esc(caseSummary(c))}${c.next ? ' · 下一步：' + esc(c.next) : ''}</span></button>`).join('') : '<div class="sub">還沒有案件。</div>'}</div>
  <div class="actions"><button class="btn danger" data-act="delclient">刪除客戶</button></div>`;
}
const isPdf = (d) => d.type === 'pdf';
function docTile(d, i, p, folders) { // p：'' 案件、'c' 客戶
  const media = isPdf(d) ? `<div class="pdftile">PDF<small>${esc(d.name)}</small></div>` : `<canvas data-${p}thumb="${i}" width="240" height="320"></canvas>`;
  const mv = folders ? `<select class="mv" data-movedoc="${i}" aria-label="移到資料夾"><option value="">未分類</option>${folders.map((f) => `<option${f === d.folder ? ' selected' : ''}>${esc(f)}</option>`).join('')}</select>` : '';
  const dt = `<input type="date" class="ddate" data-${p}docdate="${i}" value="${esc(d.date || '')}" aria-label="文件日期">`;
  return `<div class="doc">${media}${dt}<div class="bar"><span>${esc(d.name)}${d.masks?.length ? ' · 已遮 ' + d.masks.length : ''}</span><span><button data-${p}view="${i}">看</button>${isPdf(d) ? '' : `<button data-${p}mask="${i}">遮罩</button>`}<button data-${p}deldoc="${i}" aria-label="刪除文件">刪</button></span></div>${mv}</div>`;
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
const tlSorted = (c) => (c.timeline || []).slice().sort((a, b) => (b.date || '').localeCompare(a.date || '') || (b.at || '').localeCompare(a.at || ''));
const latestStep = (c) => tlSorted(c)[0];
function timelineHTML(c) {
  const tl = tlSorted(c), done = new Set(tl.map((e) => e.title)), steps = getSteps(c.cat);
  return `<div class="sec"><h2><span>進度 <span class="sub" style="letter-spacing:0">${tl.length ? tl.length + ' 筆，新的在上面' : '每個時間點記一筆'}</span></span><button class="linkbtn" data-act="editsteps">編輯常用步驟</button></h2>
   <div class="tags">${steps.map((x) => `<button class="tag${done.has(x) ? ' did' : ''}" data-addstep="${esc(x)}">${done.has(x) ? '✓' : '＋'} ${esc(x)}</button>`).join('')}<button class="tag" data-addstep="">＋ 其他</button></div>
   ${tl.length ? `<div class="tl">${tl.map((e) => `<div class="tli"><span class="tldot"></span><div class="tlrow"><input type="date" class="tldate" data-tldate="${e.id}" value="${esc(e.date)}" aria-label="日期"><input class="tltitle" data-tltitle="${e.id}" id="tlt_${e.id}" value="${esc(e.title)}" placeholder="做了什麼" aria-label="項目">${e.title && !steps.includes(e.title) ? `<button class="linkbtn" data-favstep="${esc(e.title)}" style="flex:none">＋ 常用</button>` : ''}<button class="tldel" data-tldel="${e.id}" aria-label="刪除這筆">✕</button></div><input class="tlnote" data-tlnote="${e.id}" id="tln_${e.id}" value="${esc(e.note)}" placeholder="備註（可不填）" aria-label="備註"></div>`).join('')}</div>` : ''}</div>`;
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
function curFolder(c) { return (c?.folders || []).includes(S.folder) ? S.folder : ''; }
function caseDocsHTML(c) {
  const folders = c.folders || [];
  if (!['全部', '未分類', ...folders].includes(S.folder)) S.folder = '全部';
  const cnt = (f) => c.docs.filter((d) => (f === '全部' ? true : f === '未分類' ? !d.folder : d.folder === f)).length;
  const shown = c.docs.map((d, i) => [d, i]).filter(([d]) => S.folder === '全部' || (S.folder === '未分類' ? !d.folder : d.folder === S.folder)).sort(([a], [b]) => (b.date || '').localeCompare(a.date || ''));
  const hasImg = c.docs.some((d) => !isPdf(d));
  return `<div class="sec"><h2>文件 <span class="sub">原圖永遠保留，遮罩可以隨時改</span></h2>
   <div class="chips">${['全部', ...(c.docs.some((d) => !d.folder) && folders.length ? ['未分類'] : []), ...folders].map((f) => `<button class="chip${S.folder === f ? ' on' : ''}" data-folder="${esc(f)}">${esc(f)} <span class="cnt">${cnt(f)}</span></button>`).join('')}<button class="chip" data-act="newfolder">＋ 資料夾</button></div>
   ${folders.includes(S.folder) ? `<div class="actions"><button class="linkbtn" data-act="renfolder">重新命名資料夾</button><button class="linkbtn" data-act="delfolder">刪除資料夾</button><span class="sub">新加的文件會放進「${esc(S.folder)}」</span></div>` : ''}
   <div class="docs">${shown.map(([d, i]) => docTile(d, i, '', folders.length ? folders : null)).join('')}
   <button class="adddoc" data-act="scan">${ICON.scan}拍照／掃描<span style="font-size:11px">也可以拖曳圖片、PDF 進來</span></button>
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
  const groups = kind === 'sop' ? SOPGROUPS : ['簡訊', '信件', '其他'];
  const list = notes().filter((n) => n.kind === kind && (kind !== 'sop' || S.sopGroup === '全部' || n.group === S.sopGroup));
  const q = S.q.trim(); const shown = list.filter((n) => !q || (n.title + n.body).includes(q));
  return `<h1 style="display:flex;justify-content:space-between;align-items:center;gap:8px">${kind === 'sop' ? 'SOP 作業流程' : '話術範本'}<button class="btn" data-act="import">批次匯入</button></h1>
  <input class="search" id="q" placeholder="搜尋" value="${esc(S.q)}" style="max-width:420px" autocomplete="off">
  ${kind === 'sop' ? `<div class="chips">${['全部', ...groups].map((g) => `<button class="chip${S.sopGroup === g ? ' on' : ''}" data-sg="${g}">${g}</button>`).join('')}</div>` : ''}
  <div class="cards">${shown.map((n) => `<div class="card"><h3><span>${esc(n.title)}</span><span class="sub">${esc(n.group)}</span></h3><pre>${esc(n.body)}</pre>${nimgsHTML(n)}<div class="mini">${kind === 'tpl' ? `<button data-copy="${n.id}">複製</button>` : ''}<button data-editnote="${n.id}">編輯</button><button data-delnote="${n.id}">刪除</button></div></div>`).join('') || `<div class="empty">${kind === 'sop' ? '還沒有流程。用下面的表單把 Notion、LINE 記事本裡的步驟加進來。' : '還沒有範本。'}</div>`}</div>
  <div class="sec" style="max-width:620px"><h2>新增${kind === 'sop' ? '流程' : '範本'}</h2>
   <div class="grid2"><label class="f">分類<select id="nGroup">${groups.map((g) => `<option${S.sopGroup === g ? ' selected' : ''}>${g}</option>`).join('')}</select></label><label class="f">標題<input id="nTitle"></label></div>
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
  <div class="card"><h3>待辦案件 <span class="sub">${waiting.length} 件</span></h3>${waiting.slice(0, 6).map((c) => `<button class="item" data-gocase="${c.id}"><span class="t"><span>${esc(c.client)} · ${esc(c.title)}</span><span class="pill ${pillCls(c.status)}">${c.status}</span></span><span class="s">${esc(c.next || '未填下一步')}</span></button>`).join('') || '<div class="sub">沒有待辦。</div>'}</div>
  <div class="card"><h3>同步</h3>${syncHTML()}<div class="sub">資料加密後存在 OneDrive › 應用程式 › Work › ${esc(FOLDER)}。iPad 和手機登入同一個帳號就會自動同步，不需要備份。</div>
   <div class="actions"><button class="btn" data-act="import">批次匯入</button><button class="btn" data-act="lock">鎖定</button><button class="btn" data-act="signout">登出</button></div></div>
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
  const b = e.target.closest('button'); if (!b) return; const d = b.dataset; const c = cur();
  if (d.nav) { S.view = d.nav; S.showDetail = false; S.q = ''; App.render(); return; }
  if (d.cat) { S.cat = d.cat; App.render(); return; }
  if (d.sel) { S.folder = '全部'; S.sel = d.sel; S.showDetail = true; App.render(); return; }
  if (d.rsel) { S.refSel = d.rsel; S.showDetail = true; App.render(); return; }
  if (d.ritem) { S.refItem = d.ritem; App.render(); return; }
  if (d.rcause) { S.refCause = d.rcause; App.render(); return; }
  if (d.sg) { S.sopGroup = d.sg; App.render(); return; }
  if (d.folder != null) { S.folder = d.folder; App.render(); return; }
  if (d.goclient) { S.ccat = '全部'; S.view = 'clients'; S.csel = d.goclient; S.showDetail = true; S.showIds = false; App.render(); return; }
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
  if (d.tldel && c) { if (armConfirm(b, '刪？')) { c.timeline = (c.timeline || []).filter((x) => x.id !== d.tldel); saveCase(c); App.render(); } return; }
  if (d.unlink && c) { const o = Vault.get('cases', d.unlink); c.links = (c.links || []).filter((x) => x !== d.unlink); saveCase(c); if (o) { o.links = (o.links || []).filter((x) => x !== c.id); saveCase(o); } App.render(); toast('已取消連結'); return; }
  if (d.gocase) { S.folder = '全部'; S.view = 'cases'; S.sel = d.gocase; S.showDetail = true; App.render(); return; }
  if (d.item && c) { const i = c.items.indexOf(d.item); i < 0 ? c.items.push(d.item) : c.items.splice(i, 1); saveCase(c); App.render(); return; }
  if (d.ctype && c) { c.causeType = c.causeType === d.ctype ? '' : d.ctype; c.causeDetail = ''; saveCase(c); App.render(); return; }
  if (d.cdet && c) { c.causeDetail = c.causeDetail === d.cdet ? '' : d.cdet; saveCase(c); App.render(); return; }
  if (d.mask && c) { openMask(c, +d.mask); return; }
  if (d.view && c) { openViewer(c.docs[+d.view]); return; }
  if (d.deldoc && c) { if (armConfirm(b, '確定？')) { const [doc] = c.docs.splice(+d.deldoc, 1); saveCase(c); if (doc) Files.remove(doc.fileId); App.render(); } return; }
  if (d.copy) { const n = Vault.get('notes', d.copy); try { await navigator.clipboard.writeText(n.body); toast('已複製'); } catch (err) { toast('無法自動複製，請長按文字選取'); } return; }
  if (d.nimg) { openViewer({ fileId: d.nimg, name: d.ntitle || '附圖' }); return; }
  if (d.delnote) { if (armConfirm(b, '確定刪除？')) { (Vault.get('notes', d.delnote)?.images || []).forEach((f) => Files.remove(f)); Vault.del('notes', d.delnote); App.render(); } return; }
  if (d.editnote) { editNote(d.editnote); return; }
  if (d.addnote) {
    const t = $('nTitle').value.trim(), bd = $('nBody').value.trim(), imgs = [...($('nImgs')?.files || [])];
    if (!t && !bd && !imgs.length) { toast('先填標題、內容或附圖'); return; }
    const n = Vault.put('notes', { kind: d.addnote, group: $('nGroup').value, title: t || '未命名', body: bd, images: [] });
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
    case 'delclient': if (armConfirm(b, '再按一次確定刪除（證件照片也會刪掉）')) { (ccl.docs || []).forEach((doc) => Files.remove(doc.fileId)); cases().filter((c) => c.clientId === ccl.id).forEach((c) => { c.clientId = ''; saveCase(c); }); Vault.del('clients', ccl.id); S.csel = null; S.showDetail = false; App.render(); } break;
    case 'editsteps': openStepEditor(c.cat); break;
    case 'linkcase': openLinkPicker(c); break;
    case 'convert': { const t = new Date(); const day = `${t.getFullYear()}-${String(t.getMonth() + 1).padStart(2, '0')}-${String(t.getDate()).padStart(2, '0')}`; c.cat = '新契約'; c.status = '進行中'; c.convertedAt = day; saveCase(c); S.cat = S.cat === '建議書' ? '新契約' : S.cat; App.render(); toast('已轉成新契約'); break; }
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
  if (e.target.dataset?.docdate != null) { const c = cur(); if (!c) return; c.docs[+e.target.dataset.docdate].date = e.target.value; saveCase(c); App.render(); return; }
  if (e.target.dataset?.cdocdate != null) { const cl = curClient(); if (!cl) return; cl.docs[+e.target.dataset.cdocdate].date = e.target.value; Vault.put('clients', cl); App.render(); return; }
  if (e.target.dataset?.movedoc != null) { const c = cur(); if (!c) return; c.docs[+e.target.dataset.movedoc].folder = e.target.value; saveCase(c); App.render(); return; }
  if (e.target.id === 'docUp' || e.target.id === 'docUpC') { const fs = [...e.target.files]; e.target.value = ''; takeImages(fs); }
});
app.addEventListener('change', (e) => { if (e.target.dataset?.f === 'client') { const c = cur(); if (!c) return; c.client = e.target.value; const made = linkClient(c); saveCase(c); App.render(); if (made) toast(`已建立客戶「${made.name}」`); return; } if (e.target.id === 'nImgs') { const k = e.target.files.length; $('nImgsN').textContent = k ? `已選 ${k} 張` : ''; return; } const f = e.target.dataset?.f; const c = cur(); if (!f || !c) return; c[f] = e.target.value; saveCase(c); if (['cat', 'status', 'cost', 'ratio', 'client', 'title', 'date'].includes(f)) App.render(); });
app.addEventListener('input', (e) => {
  if (e.target.id === 'q') { S.q = e.target.value; App.render(); }
  if (e.target.id === 'qq') {
    const q = e.target.value.trim(); const box = $('qres'); if (!q) { box.innerHTML = ''; return; }
    const cls = clients().filter((x) => [x.name, x.phone, x.line, x.note].join(' ').includes(q)).slice(0, 6);
    const ns = notes().filter((n) => (n.title + n.body).includes(q)).slice(0, 6), cs = cases().filter((c) => [c.client, c.title, c.notes, c.next].join(' ').includes(q)).slice(0, 6);
    box.innerHTML = `<div class="cards" style="margin-bottom:6px">${cls.map((x) => `<button class="item" data-goclient="${x.id}"><span class="t">${esc(x.name || '未填姓名')}</span><span class="s">客戶</span></button>`).join('')}${ns.map((n) => `<div class="card"><h3><span>${esc(n.title)}</span><span class="sub">${esc(n.group)}</span></h3><pre>${esc(n.body)}</pre></div>`).join('')}${cs.map((c) => `<button class="item" data-gocase="${c.id}"><span class="t">${esc(c.client)} · ${esc(c.title)}</span><span class="s">${esc(caseSummary(c))}</span></button>`).join('')}</div>${!cls.length && !ns.length && !cs.length ? '<div class="sub">找不到。</div>' : ''}`;
  }
});
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
  ov.innerHTML = `<div class="ov"><div class="ovhead"><h2>${n.kind === 'quick' ? '電話' : '編輯'}</h2><button class="btn" id="eCancel">取消</button><button class="btn pri" id="eSave">儲存</button></div><div class="ovbody" style="flex-direction:column;max-width:720px;margin:0 auto;width:100%">${n.kind === 'sop' || n.kind === 'tpl' ? `<label class="f">分類<select id="eG">${(n.kind === 'sop' ? SOPGROUPS : ['簡訊', '信件', '其他']).map((g) => `<option${g === n.group ? ' selected' : ''}>${g}</option>`).join('')}</select></label>` : ''}<label class="f">${n.kind === 'quick' ? '名稱' : '標題'}<input id="eT" value="${esc(n.title)}"></label>${n.kind === 'quick' ? `<label class="f">電話<input id="eB" inputmode="tel" value="${esc(n.body)}"></label>` : `<label class="f">內容<textarea id="eB" rows="10">${esc(n.body)}</textarea></label><div id="eImgBox">${imgsBlock()}</div>`}${n.kind === 'quick' && !isNew ? '<div class="actions"><button class="btn danger" id="eDel">刪除這支電話</button></div>' : ''}</div></div>`;
  const close = () => { NE = null; ov.innerHTML = ''; App.render(); };
  const wireImgs = () => {
    if (!withImgs) return;
    $('eImgBox').innerHTML = imgsBlock(); drawThumbs();
    $('eAddImg').onclick = () => $('eImgs').click();
    $('eCam').onclick = () => { // 先把正在打的字存起來，掃描完再回到編輯
      n.title = $('eT').value.trim(); n.body = $('eB').value; if ($('eG')) n.group = $('eG').value; delete n.refresh; Vault.put('notes', n);
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
  $('eSave').onclick = () => { n.title = $('eT').value.trim() || '未命名'; n.body = $('eB').value; if ($('eG')) n.group = $('eG').value; delete n.refresh; Vault.put('notes', n); close(); };
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
    group: x.kind === 'sop' ? (SOPGROUPS.includes(x.group) ? x.group : '壽險·其他') : x.kind === 'tpl' ? (['簡訊', '信件', '其他'].includes(x.group) ? x.group : '其他') : '電話',
    title: String(x.title || '未命名').slice(0, 200), body: String(x.body || ''),
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
  if (!SC && S.view === 'clients') { const cl = curClient(); if (!cl) { toast('先選一位客戶，再把證件拖進來'); return; } scanForClient(cl); }
  if (!SC && (S.view === 'sop' || S.view === 'tpl')) { const n = Vault.put('notes', { kind: S.view, group: S.view === 'sop' ? (S.sopGroup === '全部' ? '壽險·其他' : S.sopGroup) : '其他', title: '', body: '', images: [] }); editNote(n.id, true); toast('加入中…'); addNoteImages(n, imgs).then(() => NE?.refresh?.()); return; }
  if (!SC) { let t = S.view === 'cases' ? cur() : null; if (!t) { t = newCase('理賠'); S.view = 'cases'; } openScanner(t, { folder: curFolder(t) }); }
  if (SC.img) SC.queue.push(...imgs); else onFiles(imgs);
  if (imgs.length > 1) toast(`收到 ${imgs.length} 張，一張一張對齊後按「加入這頁」`);
}
const hasFiles = (e) => [...(e.dataTransfer?.types || [])].includes('Files');
let dragDepth = 0;
document.addEventListener('dragenter', (e) => { if (!Vault.key || !hasFiles(e)) return; e.preventDefault(); dragDepth++; document.body.classList.add('dropping'); });
document.addEventListener('dragover', (e) => { if (!Vault.key || !hasFiles(e)) return; e.preventDefault(); e.dataTransfer.dropEffect = 'copy'; });
document.addEventListener('dragleave', () => { if (--dragDepth <= 0) { dragDepth = 0; document.body.classList.remove('dropping'); } });
document.addEventListener('drop', (e) => { if (!Vault.key || !hasFiles(e)) return; e.preventDefault(); dragDepth = 0; document.body.classList.remove('dropping'); if (MK) { toast('請先關閉遮罩畫面'); return; } takeImages(e.dataTransfer.files); });
document.addEventListener('paste', (e) => {
  if (!Vault.key || MK) return;
  const files = [...(e.clipboardData?.items || [])].filter((i) => i.kind === 'file' && /^image\//.test(i.type)).map((i) => i.getAsFile()).filter(Boolean);
  if (!files.length) return; // 一般文字照常貼上
  e.preventDefault(); takeImages(files);
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
  S.showIds = false; Vault.lock(); Files.clear(); imgCache.clear(); ov.innerHTML = ''; app.innerHTML = '';
  askUnlock();
}
['pointerdown', 'keydown'].forEach((t) => document.addEventListener(t, () => (lastAct = Date.now()), true));
setInterval(() => { if (Vault.key && Date.now() - lastAct > (CFG.AUTO_LOCK_MINUTES || 30) * 60000) lock(); }, 30000);
document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible' && Vault.key) { if (Date.now() - lastAct > (CFG.AUTO_LOCK_MINUTES || 30) * 60000) lock(); else Vault.pull(); } else if (Vault.dirty) Vault.save(); });
window.addEventListener('online', () => { if (Vault.key) { Vault.save(); Files.retryUploads(); } });
window.addEventListener('beforeunload', (e) => { if (Vault.dirty) { Vault.save(); e.preventDefault(); e.returnValue = ''; } });

if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => { });
boot();
})();
