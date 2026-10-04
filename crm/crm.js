/* 客戶管理 CRM：資料加密後存在使用者自己的 OneDrive（應用程式資料夾），iPad 上只暫存加密檔 */
(() => {
  'use strict';
  const CFG = window.CRM_CONFIG || {};
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
  const today = () => { const d = new Date(); d.setMinutes(d.getMinutes() - d.getTimezoneOffset()); return d.toISOString().slice(0, 10); };
  const addDays = (iso, n) => { const d = new Date(iso + 'T00:00:00'); d.setDate(d.getDate() + n); d.setMinutes(d.getMinutes() - d.getTimezoneOffset()); return d.toISOString().slice(0, 10); };
  const fmtDate = (iso) => (iso ? iso.replace(/^(\d{4})-(\d{2})-(\d{2}).*/, '$1/$2/$3') : '');
  const MOCK = (() => { try { return localStorage.getItem('crm:mock') === '1'; } catch (e) { return false; } })();

  function toast(msg, ms = 2000) { const t = $('toast'); t.textContent = msg; t.classList.remove('hidden'); clearTimeout(toast.t); toast.t = setTimeout(() => t.classList.add('hidden'), ms); }

  /* ================= 加密 ================= */
  const enc = new TextEncoder(), dec = new TextDecoder();
  const b64 = (u8) => { let s = ''; const a = new Uint8Array(u8); for (let i = 0; i < a.length; i += 0x8000) s += String.fromCharCode.apply(null, a.subarray(i, i + 0x8000)); return btoa(s); };
  const unb64 = (s) => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));
  async function deriveKey(pw, salt) {
    const base = await crypto.subtle.importKey('raw', enc.encode(pw), 'PBKDF2', false, ['deriveKey']);
    return crypto.subtle.deriveKey({ name: 'PBKDF2', salt, iterations: 310000, hash: 'SHA-256' }, base, { name: 'AES-GCM', length: 256 }, false, ['encrypt', 'decrypt']);
  }
  async function seal(key, bytes) { const iv = crypto.getRandomValues(new Uint8Array(12)); const ct = new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, key, bytes)); const out = new Uint8Array(12 + ct.length); out.set(iv); out.set(ct, 12); return out; }
  async function open(key, bytes) { const u = new Uint8Array(bytes); return new Uint8Array(await crypto.subtle.decrypt({ name: 'AES-GCM', iv: u.slice(0, 12) }, key, u.slice(12))); }

  /* ================= 本機暫存（IndexedDB，只放加密檔） ================= */
  const Cache = {
    db: null,
    async init() {
      if (this.db) return this.db;
      this.db = await new Promise((ok, bad) => { const r = indexedDB.open('crm-cache', 1); r.onupgradeneeded = () => r.result.createObjectStore('kv'); r.onsuccess = () => ok(r.result); r.onerror = () => bad(r.error); });
      return this.db;
    },
    async get(k) { try { const db = await this.init(); return await new Promise((ok) => { const r = db.transaction('kv').objectStore('kv').get(k); r.onsuccess = () => ok(r.result); r.onerror = () => ok(undefined); }); } catch (e) { return undefined; } },
    async set(k, v) { try { const db = await this.init(); await new Promise((ok) => { const t = db.transaction('kv', 'readwrite'); t.objectStore('kv').put(v, k); t.oncomplete = ok; t.onerror = ok; }); } catch (e) { } },
    async del(k) { try { const db = await this.init(); await new Promise((ok) => { const t = db.transaction('kv', 'readwrite'); t.objectStore('kv').delete(k); t.oncomplete = ok; t.onerror = ok; }); } catch (e) { } },
    async clear() { try { const db = await this.init(); await new Promise((ok) => { const t = db.transaction('kv', 'readwrite'); t.objectStore('kv').clear(); t.oncomplete = ok; t.onerror = ok; }); } catch (e) { } },
  };

  /* ================= OneDrive（Microsoft Graph，應用程式資料夾） ================= */
  const SCOPES = ['Files.ReadWrite.AppFolder', 'User.Read'];
  const Drive = {
    msal: null, account: null,
    async init() {
      if (MOCK) { this.account = { username: '測試帳號' }; return; }
      this.msal = new msal.PublicClientApplication({
        auth: { clientId: CFG.CLIENT_ID, authority: 'https://login.microsoftonline.com/consumers', redirectUri: location.origin + location.pathname.replace(/index\.html$/, '') },
        cache: { cacheLocation: 'localStorage' },
      });
      await this.msal.initialize();
      const r = await this.msal.handleRedirectPromise().catch((e) => { console.warn(e); return null; });
      this.account = (r && r.account) || this.msal.getAllAccounts()[0] || null;
    },
    signIn() { return this.msal.loginRedirect({ scopes: SCOPES, prompt: 'select_account' }); },
    async signOut() { if (MOCK) return; const a = this.account; this.account = null; await this.msal.logoutRedirect({ account: a }); },
    async token() {
      try { return (await this.msal.acquireTokenSilent({ scopes: SCOPES, account: this.account })).accessToken; }
      catch (e) { if (e instanceof msal.InteractionRequiredAuthError) { await this.msal.acquireTokenRedirect({ scopes: SCOPES, account: this.account }); } throw e; }
    },
    async api(path, opts = {}) {
      const t = await this.token();
      const r = await fetch('https://graph.microsoft.com/v1.0' + path, { ...opts, headers: { Authorization: 'Bearer ' + t, ...(opts.headers || {}) } });
      return r;
    },
    item: (p) => '/me/drive/special/approot:/' + p.split('/').map(encodeURIComponent).join('/') + ':',
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
  /* 測試用假雲端（localStorage crm:mock=1 時啟用，資料存在瀏覽器裡） */
  const Mock = {
    async getMeta(p) { const v = await Cache.get('mock:' + p); return v ? { eTag: v.eTag, size: v.bytes.length } : null; },
    async download(p) { const v = await Cache.get('mock:' + p); return v ? { bytes: v.bytes, eTag: v.eTag } : null; },
    async upload(p, bytes, eTag) { const v = await Cache.get('mock:' + p); if (eTag && v && v.eTag !== eTag) { const e = new Error('conflict'); e.conflict = true; throw e; } const t = uid(); await Cache.set('mock:' + p, { bytes, eTag: t }); return t; },
    async remove(p) { await Cache.del('mock:' + p); return true; },
  };

  /* ================= 資料庫（記憶體中明文，存檔時加密） ================= */
  const COLLS = ['clients', 'policies', 'deals', 'services', 'claims', 'tasks', 'logs', 'files'];
  const emptyDb = () => Object.fromEntries([['v', 1], ...COLLS.map((c) => [c, {}])]);
  const Vault = {
    key: null, salt: null, db: null, eTag: null, dirty: false, saving: false, state: 'idle',
    async load(pw) { // 回傳 'new' | 'ok'，密碼錯誤時丟出錯誤
      let raw = null, fromCache = false;
      try { raw = await Drive.download('data.enc'); } catch (e) { const c = await Cache.get('data'); if (c) { raw = c; fromCache = true; } else throw e; }
      if (!raw) return 'new';
      const env = JSON.parse(dec.decode(raw.bytes));
      const salt = unb64(env.salt), key = await deriveKey(pw, salt);
      let plain;
      try { plain = await open(key, unb64(env.data)); } catch (e) { const err = new Error('密碼不正確'); err.badPassword = true; throw err; }
      this.key = key; this.salt = salt; this.db = { ...emptyDb(), ...JSON.parse(dec.decode(plain)) }; this.eTag = raw.eTag;
      if (!fromCache) await Cache.set('data', raw);
      const pend = await Cache.get('pending');
      if (pend) { try { const local = JSON.parse(dec.decode(await open(key, pend))); this.db = merge(this.db, local); this.markDirty(); } catch (e) { } }
      this.setState(fromCache ? 'offline' : 'synced');
      return 'ok';
    },
    async create(pw) {
      this.salt = crypto.getRandomValues(new Uint8Array(16)); this.key = await deriveKey(pw, this.salt); this.db = emptyDb(); this.eTag = null;
      await this.save(true);
    },
    async envelope() { const data = await seal(this.key, enc.encode(JSON.stringify(this.db))); return enc.encode(JSON.stringify({ v: 1, salt: b64(this.salt), data: b64(data), at: new Date().toISOString() })); },
    markDirty() { this.dirty = true; this.setState('pending'); clearTimeout(this.t); this.t = setTimeout(() => this.save(), 1200); seal(this.key, enc.encode(JSON.stringify(this.db))).then((b) => Cache.set('pending', b)); },
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
            const remote = await Drive.download('data.enc'); // 別台裝置改過：合併後再存
            const env = JSON.parse(dec.decode(remote.bytes));
            const other = JSON.parse(dec.decode(await open(this.key, unb64(env.data))));
            this.db = merge(other, this.db); this.eTag = remote.eTag; App.refresh();
          }
        }
      } catch (e) { console.warn(e); this.setState('offline'); clearTimeout(this.retry); this.retry = setTimeout(() => this.save(), 30000); }
      finally { this.saving = false; if (this.dirty && this.state === 'synced') this.save(); }
    },
    setState(s) {
      this.state = s;
      const L = { synced: '✓ 已同步 OneDrive', saving: '同步中…', pending: '待同步…', offline: '⚠ 離線，連上網路會自動同步', idle: '—' };
      const el = $('sync'); if (el) { el.textContent = L[s] || s; el.dataset.s = s; }
    },
    lock() { this.key = null; this.db = null; this.salt = null; },
    put(coll, rec) { const now = new Date().toISOString(); if (!rec.id) { rec.id = uid(); rec.createdAt = now; } rec.updatedAt = now; this.db[coll][rec.id] = rec; this.markDirty(); return rec; },
    del(coll, id) { const r = this.db[coll][id]; if (!r) return; this.db[coll][id] = { id, deleted: true, updatedAt: new Date().toISOString() }; this.markDirty(); },
    all(coll) { return Object.values(this.db[coll] || {}).filter((r) => !r.deleted); },
    get(coll, id) { const r = this.db[coll][id]; return r && !r.deleted ? r : null; },
  };
  function merge(a, b) { // 依每筆資料的 updatedAt 取較新的
    const out = emptyDb();
    COLLS.forEach((c) => { const A = a[c] || {}, B = b[c] || {}; new Set([...Object.keys(A), ...Object.keys(B)]).forEach((id) => { const x = A[id], y = B[id]; out[c][id] = !x ? y : !y ? x : ((x.updatedAt || '') >= (y.updatedAt || '') ? x : y); }); });
    return out;
  }

  /* ================= 照片（加密後存 OneDrive files/） ================= */
  const fpath = (f) => 'files/' + f.id + (f.kv ? '.k' + f.kv : '') + '.enc';
  const Files = {
    mem: new Map(),
    async add(file, ref) {
      const blob = await compress(file);
      const id = uid(), bytes = new Uint8Array(await blob.arrayBuffer());
      const sealed = await seal(Vault.key, bytes);
      await Cache.set('file:' + id, sealed);
      Vault.put('files', { id, type: blob.type, size: bytes.length, ref, name: file.name || '照片' });
      this.mem.set(id, URL.createObjectURL(blob));
      Drive.upload(fpath({ id }), sealed).then(() => Cache.set('file-up:' + id, 1)).catch(() => toast('照片已存在 iPad，連上網路後會再上傳'));
      return id;
    },
    async url(id) {
      if (this.mem.has(id)) return this.mem.get(id);
      let sealed = await Cache.get('file:' + id);
      const meta = Vault.get('files', id) || { id };
      if (!sealed) { const r = await Drive.download(fpath(meta)); if (!r) return null; sealed = r.bytes; Cache.set('file:' + id, sealed); }
      const u = URL.createObjectURL(new Blob([await open(Vault.key, sealed)], { type: meta.type || 'image/jpeg' }));
      this.mem.set(id, u); return u;
    },
    async remove(id) { const f = Vault.get('files', id) || { id }; Vault.del('files', id); await Cache.del('file:' + id); Drive.remove(fpath(f)).catch(() => { }); },
    async retryUploads() {
      for (const f of Vault.all('files')) {
        if (await Cache.get('file-up:' + f.id)) continue;
        const s = await Cache.get('file:' + f.id); if (!s) continue;
        try { const m = await Drive.getMeta(fpath(f)); if (!m) await Drive.upload(fpath(f), s); await Cache.set('file-up:' + f.id, 1); } catch (e) { return; }
      }
    },
    forRef(coll, id) { return Vault.all('files').filter((f) => f.ref && f.ref.coll === coll && f.ref.id === id).sort((a, b) => (a.createdAt < b.createdAt ? -1 : 1)); },
  };
  async function compress(file, max = 1800, q = 0.82) {
    if (!/^image\//.test(file.type)) return file;
    const bmp = await createImageBitmap(file).catch(() => null); if (!bmp) return file;
    const r = Math.min(1, max / Math.max(bmp.width, bmp.height));
    const c = document.createElement('canvas'); c.width = Math.round(bmp.width * r); c.height = Math.round(bmp.height * r);
    c.getContext('2d').drawImage(bmp, 0, 0, c.width, c.height);
    return await new Promise((ok) => c.toBlob((b) => ok(b || file), 'image/jpeg', q));
  }

  /* ================= 個資遮罩 ================= */
  const Mask = {
    all: false, timer: null,
    name: (s) => { s = String(s || ''); if (s.length <= 1) return s; if (s.length === 2) return s[0] + '○'; return s[0] + '○'.repeat(s.length - 2) + s[s.length - 1]; },
    id: (s) => { s = String(s || ''); return s.length < 5 ? '＊＊＊' : s.slice(0, 3) + '＊'.repeat(Math.max(1, s.length - 5)) + s.slice(-2); },
    phone: (s) => { const d = String(s || ''); return d.length < 7 ? '＊＊＊' : d.slice(0, 4) + '-＊＊＊-' + d.slice(-3); },
    addr: (s) => { s = String(s || ''); return s.length > 6 ? s.slice(0, 6) + '…' : s; },
    email: (s) => { s = String(s || ''); const [u, d] = s.split('@'); return d ? u[0] + '＊＊＊@' + d : '＊＊＊'; },
    date: (s) => (s ? '＊＊＊＊/' + s.slice(5).replace('-', '/') : ''),
    num: (s) => { s = String(s || ''); return s.length > 4 ? '＊'.repeat(s.length - 4) + s.slice(-4) : '＊＊＊＊'; },
    text: () => '（已遮蔽，按一下顯示）',
  };
  // pii(值, 遮罩方式)：第一層自動遮蔽；按一下（或右上「顯示個資」）才顯示完整內容
  function pii(v, kind) {
    if (v == null || v === '') return '';
    const full = String(v), masked = (Mask[kind] || Mask.text)(full);
    if (Mask.all) return `<span class="pii open">${esc(full)}</span>`;
    return `<span class="pii" data-full="${esc(full)}" data-masked="${esc(masked)}" title="按一下顯示">${esc(masked)}</span>`;
  }
  document.addEventListener('click', (e) => {
    const p = e.target.closest('.pii:not(.open)'); if (!p) return;
    e.preventDefault(); e.stopPropagation();
    p.textContent = p.dataset.full; p.classList.add('open');
    setTimeout(() => { if (!Mask.all && p.isConnected) { p.textContent = p.dataset.masked; p.classList.remove('open'); } }, 30000);
  }, true);
  function setRevealAll(on) {
    Mask.all = on; clearTimeout(Mask.timer);
    $('btnReveal').setAttribute('aria-pressed', on); $('btnReveal').textContent = on ? '隱藏個資' : '顯示個資';
    document.body.classList.toggle('reveal-all', on);
    if (on) Mask.timer = setTimeout(() => setRevealAll(false), 120000);
    App.refresh();
  }

  /* ================= 選項 ================= */
  const STAGES = ['名單', '約訪', '需求分析', '建議書', '送件', '核保', '成交', '暫緩'];
  const SERVICE_TYPES = ['受益人變更', '地址／聯絡資料變更', '繳費方式變更', '保額／保障變更', '保單借款／還款', '停效／復效', '契約轉換', '解約／部分解約', '要保人變更', '其他'];
  const SERVICE_STATUS = ['待辦', '準備文件', '已送件', '補件中', '完成'];
  const CLAIM_STATUS = ['準備文件', '已送件', '補件中', '已給付', '拒賠／結案'];
  const CLAIM_DOCS = ['理賠申請書', '診斷證明書', '醫療收據', '住院證明', '手術證明', '身分證影本', '存摺影本', '事故證明（意外）', '病歷摘要', '重大傷病證明'];
  const TASK_CATS = ['行政', '業績', '學習', '活動', '其他'];
  const LOG_TYPES = ['電話', '面談', 'LINE', '拜訪', '送禮', '其他'];

  /* ================= 表單 ================= */
  // 欄位：{k,label,type:text|date|number|select|textarea|client|checks|tags, opts, full:佔整行, ph}
  function formHtml(fields, rec) {
    return '<form class="form" autocomplete="off">' + fields.map((f) => {
      const v = rec[f.k] ?? (f.def ?? '');
      const id = 'f_' + f.k, cls = 'fld' + (f.full ? ' full' : '');
      let input;
      if (f.type === 'select') input = `<select id="${id}" name="${f.k}">${f.opts.map((o) => `<option ${o === v ? 'selected' : ''}>${esc(o)}</option>`).join('')}</select>`;
      else if (f.type === 'textarea') input = `<textarea id="${id}" name="${f.k}" rows="3" placeholder="${esc(f.ph || '')}">${esc(v)}</textarea>`;
      else if (f.type === 'client') input = clientPicker(id, f.k, v);
      else if (f.type === 'checks') { const set = new Set(v || []); input = `<div class="checks">${f.opts.map((o) => `<label><input type="checkbox" name="${f.k}" value="${esc(o)}" ${set.has(o) ? 'checked' : ''}>${esc(o)}</label>`).join('')}</div>`; }
      else input = `<input id="${id}" name="${f.k}" type="${f.type || 'text'}" value="${esc(v)}" placeholder="${esc(f.ph || '')}" ${f.type === 'number' ? 'inputmode="decimal"' : ''}>`;
      return `<div class="${cls}"><label for="${id}">${esc(f.label)}</label>${input}</div>`;
    }).join('') + '</form>';
  }
  function clientPicker(id, name, v) {
    const cs = Vault.all('clients').sort((a, b) => (a.name || '').localeCompare(b.name || '', 'zh-Hant'));
    return `<div class="picker"><input class="pick-q" placeholder="搜尋姓名或電話"><select id="${id}" name="${name}"><option value="">（未指定）</option>${cs.map((c) => `<option value="${c.id}" ${c.id === v ? 'selected' : ''} data-s="${esc((c.name || '') + (c.phone || ''))}">${esc(Mask.all ? c.name : Mask.name(c.name))}</option>`).join('')}</select></div>`;
  }
  function readForm(form, fields, rec) {
    const fd = new FormData(form);
    fields.forEach((f) => {
      if (f.type === 'checks') rec[f.k] = fd.getAll(f.k);
      else if (f.type === 'number') { const n = fd.get(f.k); rec[f.k] = n === '' || n == null ? '' : Number(n); }
      else rec[f.k] = (fd.get(f.k) ?? '').toString().trim();
    });
    return rec;
  }
  function wirePickers(root) {
    root.querySelectorAll('.picker').forEach((p) => {
      const q = p.querySelector('.pick-q'), sel = p.querySelector('select');
      q.addEventListener('input', () => { const t = q.value.trim(); [...sel.options].forEach((o, i) => { if (i) o.hidden = t && !(o.dataset.s || '').includes(t); }); const first = [...sel.options].find((o, i) => i && !o.hidden); if (t && first) sel.value = first.value; });
    });
  }

  /* ================= 彈出面板 ================= */
  const Sheet = {
    open(title, html, opts = {}) {
      $('sheetTitle').textContent = title; const b = $('sheetBody'); b.innerHTML = html;
      $('sheet').classList.remove('hidden'); $('sheet').classList.toggle('wide', !!opts.wide); b.scrollTop = 0; wirePickers(b); return b;
    },
    close() { $('sheet').classList.add('hidden'); },
  };
  $('sheetClose').onclick = () => Sheet.close();
  $('sheet').addEventListener('click', (e) => { if (e.target.id === 'sheet') Sheet.close(); });
  function confirmBox(msg, ok, okText = '確定', danger = true) {
    const b = Sheet.open('請確認', `<p class="lead">${esc(msg)}</p><div class="row"><button class="${danger ? 'danger' : 'primary'}" id="cOk">${esc(okText)}</button><button class="ghostbtn" id="cNo">取消</button></div>`);
    b.querySelector('#cOk').onclick = () => { Sheet.close(); ok(); }; b.querySelector('#cNo').onclick = () => Sheet.close();
  }
  function editor(title, fields, rec, onSave, extra = {}) {
    const b = Sheet.open(title, formHtml(fields, rec) + (extra.html || '') + `<div class="row foot">${extra.onDelete ? '<button class="danger" id="eDel">刪除</button>' : ''}<div class="spacer"></div><button class="ghostbtn" id="eCancel">取消</button><button class="primary" id="eSave">儲存</button></div>`, { wide: true });
    const form = b.querySelector('form');
    if (extra.wire) extra.wire(b);
    b.querySelector('#eCancel').onclick = () => Sheet.close();
    b.querySelector('#eSave').onclick = () => {
      const r = readForm(form, fields, { ...rec });
      const bad = fields.find((f) => f.req && !r[f.k]); if (bad) { toast('請填寫「' + bad.label + '」'); return; }
      onSave(r); Sheet.close(); App.refresh(); toast('已儲存');
    };
    if (extra.onDelete) b.querySelector('#eDel').onclick = () => confirmBox('確定要刪除嗎？刪除後無法復原。', () => { extra.onDelete(); App.refresh(); toast('已刪除'); }, '刪除');
  }

  /* ================= 欄位定義 ================= */
  const F = {
    client: [
      { k: 'name', label: '姓名', req: 1 }, { k: 'gender', label: '性別', type: 'select', opts: ['', '女', '男'] },
      { k: 'birthday', label: '生日', type: 'date' }, { k: 'idNo', label: '身分證字號' },
      { k: 'phone', label: '手機', type: 'tel' }, { k: 'line', label: 'LINE ID' }, { k: 'email', label: 'Email', type: 'email' },
      { k: 'job', label: '職業' }, { k: 'address', label: '地址', full: 1 },
      { k: 'source', label: '來源／介紹人' }, { k: 'tags', label: '標籤', ph: '例：家庭支柱、VIP、轉介紹（用逗號分隔）' },
      { k: 'family', label: '家庭狀況', type: 'textarea', full: 1, ph: '配偶、子女、父母、家庭責任…' },
      { k: 'note', label: '備註', type: 'textarea', full: 1 },
    ],
    policy: [
      { k: 'product', label: '商品名稱', req: 1 }, { k: 'policyNo', label: '保單號碼' },
      { k: 'company', label: '保險公司', def: '國泰人壽' }, { k: 'status', label: '狀態', type: 'select', opts: ['有效', '繳清', '停效', '已終止'] },
      { k: 'holder', label: '要保人' }, { k: 'insured', label: '被保險人' }, { k: 'beneficiary', label: '受益人', full: 1 },
      { k: 'amount', label: '保額' }, { k: 'premium', label: '年繳保費', type: 'number' },
      { k: 'payMode', label: '繳別', type: 'select', opts: ['年繳', '半年繳', '季繳', '月繳', '躉繳'] }, { k: 'startDate', label: '生效日', type: 'date' },
      { k: 'payYears', label: '繳費年期' }, { k: 'note', label: '備註', type: 'textarea', full: 1 },
    ],
    deal: [
      { k: 'clientId', label: '客戶', type: 'client', full: 1 }, { k: 'title', label: '需求／商品', req: 1, ph: '例：補重大傷病、小孩醫療' },
      { k: 'stage', label: '階段', type: 'select', opts: STAGES }, { k: 'premium', label: '預估年繳保費', type: 'number' },
      { k: 'nextAction', label: '下一步', ph: '例：送建議書、約第二次面談' }, { k: 'nextDate', label: '下一步日期', type: 'date' },
      { k: 'note', label: '紀錄', type: 'textarea', full: 1 },
    ],
    service: [
      { k: 'clientId', label: '客戶', type: 'client', full: 1 }, { k: 'type', label: '保全項目', type: 'select', opts: SERVICE_TYPES },
      { k: 'status', label: '進度', type: 'select', opts: SERVICE_STATUS }, { k: 'policyNo', label: '保單號碼' },
      { k: 'submitDate', label: '送件日', type: 'date' }, { k: 'dueDate', label: '追蹤日', type: 'date' },
      { k: 'note', label: '內容／紀錄', type: 'textarea', full: 1 },
    ],
    claim: [
      { k: 'clientId', label: '客戶', type: 'client', full: 1 }, { k: 'insured', label: '被保險人' },
      { k: 'kind', label: '類型', type: 'select', opts: ['疾病', '意外', '癌症', '重大傷病', '長照／失能', '身故', '其他'] },
      { k: 'eventDate', label: '事故日／確診日', type: 'date' }, { k: 'hospital', label: '醫院' },
      { k: 'diagnosis', label: '診斷／事故經過', type: 'textarea', full: 1 },
      { k: 'docs', label: '文件清單（已備齊打勾）', type: 'checks', opts: CLAIM_DOCS, full: 1 },
      { k: 'status', label: '進度', type: 'select', opts: CLAIM_STATUS }, { k: 'submitDate', label: '送件日', type: 'date' },
      { k: 'followDate', label: '追蹤日', type: 'date' }, { k: 'paidAmount', label: '給付金額', type: 'number' },
      { k: 'paidDate', label: '給付日', type: 'date' }, { k: 'note', label: '備註', type: 'textarea', full: 1 },
    ],
    task: [
      { k: 'title', label: '事項', req: 1, full: 1 }, { k: 'cat', label: '分類', type: 'select', opts: TASK_CATS },
      { k: 'due', label: '期限', type: 'date' }, { k: 'clientId', label: '相關客戶（可不填）', type: 'client', full: 1 },
      { k: 'note', label: '備註', type: 'textarea', full: 1 },
    ],
    log: [
      { k: 'date', label: '日期', type: 'date' }, { k: 'type', label: '方式', type: 'select', opts: LOG_TYPES },
      { k: 'content', label: '內容', type: 'textarea', full: 1, req: 1 },
    ],
  };

  /* ================= 照片區塊 ================= */
  function photoBlock(coll, id) {
    if (!id) return '<p class="hint">儲存後即可加入照片。</p>';
    const fs = Files.forRef(coll, id);
    return `<div class="photos" data-coll="${coll}" data-id="${id}">${fs.map((f) => `<button class="ph" data-f="${f.id}"><span>載入中…</span></button>`).join('')}<label class="ph add"><input type="file" accept="image/*" multiple hidden>＋ 拍照／選照片</label></div>`;
  }
  function wirePhotos(root) {
    root.querySelectorAll('.photos').forEach((box) => {
      box.querySelectorAll('.ph[data-f]').forEach(async (b) => {
        try { const u = await Files.url(b.dataset.f); if (u) b.innerHTML = `<img src="${u}" alt="照片">`; else b.innerHTML = '<span>找不到</span>'; } catch (e) { b.innerHTML = '<span>離線中</span>'; }
        b.onclick = () => Viewer.open(b.dataset.f, () => App.refresh());
      });
      const inp = box.querySelector('input[type=file]');
      inp.onchange = async () => {
        const list = [...inp.files]; if (!list.length) return;
        toast('加密並上傳中…', 4000);
        for (const f of list) await Files.add(f, { coll: box.dataset.coll, id: box.dataset.id });
        toast('已加入 ' + list.length + ' 張照片'); App.refresh();
        const fresh = document.querySelector(`.photos[data-coll="${box.dataset.coll}"][data-id="${box.dataset.id}"]`);
        if (fresh) { fresh.outerHTML = photoBlock(box.dataset.coll, box.dataset.id); } else { box.outerHTML = photoBlock(box.dataset.coll, box.dataset.id); }
        wirePhotos(document);
      };
    });
  }
  const Viewer = {
    id: null,
    async open(id, after) {
      this.id = id; this.after = after; $('viewerImg').src = await Files.url(id); $('viewer').classList.remove('hidden');
    },
    close() { $('viewer').classList.add('hidden'); $('viewerImg').removeAttribute('src'); },
  };
  $('viewerClose').onclick = () => Viewer.close();
  $('viewerShare').onclick = async () => {
    const u = $('viewerImg').src, blob = await (await fetch(u)).blob(), file = new File([blob], '照片.jpg', { type: blob.type });
    if (navigator.canShare && navigator.canShare({ files: [file] })) navigator.share({ files: [file] }).catch(() => { });
    else { const a = document.createElement('a'); a.href = u; a.download = '照片.jpg'; a.click(); }
  };
  $('viewerDel').onclick = () => confirmBox('確定要刪除這張照片嗎？', async () => { await Files.remove(Viewer.id); Viewer.close(); App.refresh(); toast('照片已刪除'); }, '刪除');

  /* ================= 小元件 ================= */
  const clientName = (id) => { const c = id && Vault.get('clients', id); return c ? pii(c.name, 'name') : '<span class="muted">（未指定客戶）</span>'; };
  const money = (n) => (n === '' || n == null ? '' : Number(n).toLocaleString('zh-TW'));
  const dueTag = (d) => { if (!d) return ''; const t = today(); const cls = d < t ? 'late' : d === t ? 'now' : d <= addDays(t, 7) ? 'soon' : ''; const lab = d < t ? '已過期 ' : d === t ? '今天 ' : ''; return `<span class="due ${cls}">${lab}${fmtDate(d)}</span>`; };
  const pill = (s, map) => `<span class="pill s-${(map || {})[s] || 'x'}">${esc(s)}</span>`;
  const STAT = { '完成': 'ok', '已給付': 'ok', '成交': 'ok', '拒賠／結案': 'off', '暫緩': 'off', '已送件': 'go', '核保': 'go', '送件': 'go', '補件中': 'warn' };
  function empty(msg, btn) { return `<div class="empty"><p>${esc(msg)}</p>${btn ? `<button class="primary" data-add>${esc(btn)}</button>` : ''}</div>`; }

  /* ================= 各頁面 ================= */
  const Views = {
    today() {
      const t = today(), wk = addDays(t, 7);
      const items = [];
      Vault.all('deals').filter((d) => d.nextDate && d.nextDate <= wk && !['成交', '暫緩'].includes(d.stage)).forEach((d) => items.push({ d: d.nextDate, kind: '促約', html: `${clientName(d.clientId)}・${esc(d.title)}<small>${esc(d.nextAction || d.stage)}</small>`, open: () => Edit.deal(d) }));
      Vault.all('services').filter((s) => s.dueDate && s.dueDate <= wk && s.status !== '完成').forEach((s) => items.push({ d: s.dueDate, kind: '保全', html: `${clientName(s.clientId)}・${esc(s.type)}<small>${esc(s.status)}</small>`, open: () => Edit.service(s) }));
      Vault.all('claims').filter((c) => c.followDate && c.followDate <= wk && !['已給付', '拒賠／結案'].includes(c.status)).forEach((c) => items.push({ d: c.followDate, kind: '理賠', html: `${clientName(c.clientId)}・${esc(c.kind)}<small>${esc(c.status)}</small>`, open: () => Edit.claim(c) }));
      Vault.all('tasks').filter((x) => !x.done && x.due && x.due <= wk).forEach((x) => items.push({ d: x.due, kind: '行政', html: `${esc(x.title)}<small>${esc(x.cat || '')}</small>`, open: () => Edit.task(x) }));
      items.sort((a, b) => (a.d < b.d ? -1 : 1));
      const m = t.slice(5, 7);
      const bdays = Vault.all('clients').filter((c) => c.birthday && c.birthday.slice(5, 7) === m).sort((a, b) => (a.birthday.slice(8) < b.birthday.slice(8) ? -1 : 1));
      const annis = [];
      Vault.all('policies').filter((p) => p.startDate && p.status !== '已終止').forEach((p) => {
        const y = +t.slice(0, 4); let a = y + p.startDate.slice(4); if (a < t) a = (y + 1) + p.startDate.slice(4);
        if (a <= addDays(t, 30) && a.slice(0, 4) !== p.startDate.slice(0, 4)) annis.push({ a, p });
      });
      annis.sort((x, y) => (x.a < y.a ? -1 : 1));
      const openClaims = Vault.all('claims').filter((c) => !['已給付', '拒賠／結案'].includes(c.status)).length;
      const openDeals = Vault.all('deals').filter((d) => !['成交', '暫緩'].includes(d.stage));
      const pipe = openDeals.reduce((s, d) => s + (+d.premium || 0), 0);
      this._items = items;
      return `<div class="kpis"><div><b>${items.filter((i) => i.d <= t).length}</b><span>今天要處理</span></div><div><b>${openDeals.length}</b><span>進行中促約</span></div><div><b>${money(pipe) || 0}</b><span>預估年繳保費</span></div><div><b>${openClaims}</b><span>處理中理賠</span></div></div>
      <div class="cols2">
        <div class="panel"><h2>未來 7 天待辦</h2>${items.length ? `<ul class="list">${items.map((i, n) => `<li class="click" data-i="${n}"><span class="k k-${i.kind}">${i.kind}</span><div class="main">${i.html}</div>${dueTag(i.d)}</li>`).join('')}</ul>` : '<p class="muted pad">沒有到期事項。</p>'}</div>
        <div class="panel"><h2>本月生日（${+m} 月）</h2>${bdays.length ? `<ul class="list">${bdays.map((c) => `<li class="click" data-c="${c.id}"><div class="main">${pii(c.name, 'name')}</div><span class="due">${+c.birthday.slice(5, 7)}/${+c.birthday.slice(8)}</span></li>`).join('')}</ul>` : '<p class="muted pad">這個月沒有客戶生日。</p>'}
        <h2>30 天內保單週年</h2>${annis.length ? `<ul class="list">${annis.map(({ a, p }) => `<li class="click" data-c="${p.clientId}"><div class="main">${clientName(p.clientId)}・${esc(p.product)}</div>${dueTag(a)}</li>`).join('')}</ul>` : '<p class="muted pad">沒有即將到來的保單週年。</p>'}</div>
      </div>`;
    },
    clients() {
      const q = (App.q || '').trim();
      let cs = Vault.all('clients');
      if (q) cs = cs.filter((c) => [c.name, c.phone, c.tags, c.source, c.job].join(' ').includes(q));
      cs.sort((a, b) => (b.updatedAt || '').localeCompare(a.updatedAt || ''));
      const counts = (id, coll) => Vault.all(coll).filter((x) => x.clientId === id).length;
      return `<div class="toolbar"><input id="q" class="search" placeholder="搜尋姓名、電話、標籤…" value="${esc(q)}"><span class="muted">${cs.length} 位客戶</span></div>` +
        (cs.length ? `<div class="grid">${cs.map((c) => `<button class="ccard" data-c="${c.id}"><div class="cn">${pii(c.name, 'name')}${c.gender ? `<small>${esc(c.gender)}</small>` : ''}</div><div class="cp">${pii(c.phone, 'phone')}</div><div class="ct">${(c.tags || '').split(/[,，、]/).filter(Boolean).map((t) => `<span class="tag">${esc(t.trim())}</span>`).join('')}</div><div class="cm">保單 ${counts(c.id, 'policies')}・促約 ${counts(c.id, 'deals')}・理賠 ${counts(c.id, 'claims')}</div></button>`).join('')}</div>` : empty(q ? '找不到符合的客戶。' : '還沒有客戶。從右上「＋ 新增」開始建立第一位客戶。', q ? '' : '＋ 新增客戶'));
    },
    client(id) {
      const c = Vault.get('clients', id); if (!c) return empty('找不到這位客戶。');
      const pol = Vault.all('policies').filter((p) => p.clientId === id);
      const logs = Vault.all('logs').filter((l) => l.clientId === id).sort((a, b) => (b.date || '').localeCompare(a.date || ''));
      const deals = Vault.all('deals').filter((d) => d.clientId === id), svcs = Vault.all('services').filter((s) => s.clientId === id), cls = Vault.all('claims').filter((x) => x.clientId === id);
      const row = (l, v) => (v ? `<div class="kv"><span>${l}</span><b>${v}</b></div>` : '');
      const age = c.birthday ? Math.floor((Date.now() - new Date(c.birthday)) / 31557600000) : null;
      return `<div class="cols2 detail">
        <div class="panel"><div class="ph-head"><h2>${pii(c.name, 'name')} ${age != null ? `<small>${age} 歲</small>` : ''}</h2><button class="ghostbtn" data-edit-client>編輯</button></div>
          ${row('生日', pii(c.birthday, 'date'))}${row('身分證', pii(c.idNo, 'id'))}${row('手機', pii(c.phone, 'phone'))}${row('LINE', pii(c.line, 'text'))}${row('Email', pii(c.email, 'email'))}${row('地址', pii(c.address, 'addr'))}${row('職業', esc(c.job))}${row('來源', esc(c.source))}${row('標籤', esc(c.tags))}
          ${c.family ? `<h3>家庭狀況</h3><p class="pre">${pii(c.family, 'text')}</p>` : ''}${c.note ? `<h3>備註</h3><p class="pre">${pii(c.note, 'text')}</p>` : ''}
          <h3>照片／文件</h3>${photoBlock('clients', c.id)}
        </div>
        <div>
          <div class="panel"><div class="ph-head"><h2>保單（${pol.length}）</h2><button class="ghostbtn" data-add-policy>＋ 保單</button></div>
            ${pol.length ? `<ul class="list">${pol.map((p) => `<li class="click" data-p="${p.id}"><div class="main">${esc(p.product)}<small>${p.policyNo ? '號碼 ' + pii(p.policyNo, 'num') + '・' : ''}${esc(p.payMode || '')} ${money(p.premium)}${p.startDate ? '・生效 ' + fmtDate(p.startDate) : ''}</small></div>${pill(p.status || '有效', { '有效': 'ok', '繳清': 'go', '停效': 'warn', '已終止': 'off' })}</li>`).join('')}</ul>` : '<p class="muted pad">尚未登錄保單。</p>'}
          </div>
          <div class="panel"><div class="ph-head"><h2>往來紀錄</h2><button class="ghostbtn" data-add-log>＋ 紀錄</button></div>
            ${logs.length ? `<ul class="list">${logs.map((l) => `<li class="click" data-l="${l.id}"><span class="k">${esc(l.type)}</span><div class="main">${pii(l.content, 'text')}</div><span class="due">${fmtDate(l.date)}</span></li>`).join('')}</ul>` : '<p class="muted pad">還沒有紀錄。</p>'}
          </div>
          <div class="panel"><div class="ph-head"><h2>相關案件</h2><div class="row"><button class="ghostbtn" data-new="deal">＋ 促約</button><button class="ghostbtn" data-new="service">＋ 保全</button><button class="ghostbtn" data-new="claim">＋ 理賠</button></div></div>
            <ul class="list">${deals.map((d) => `<li class="click" data-deal="${d.id}"><span class="k k-促約">促約</span><div class="main">${esc(d.title)}<small>${esc(d.nextAction || '')}</small></div>${pill(d.stage, STAT)}</li>`).join('')}${svcs.map((s) => `<li class="click" data-svc="${s.id}"><span class="k k-保全">保全</span><div class="main">${esc(s.type)}</div>${pill(s.status, STAT)}</li>`).join('')}${cls.map((x) => `<li class="click" data-claim="${x.id}"><span class="k k-理賠">理賠</span><div class="main">${esc(x.kind)}・${fmtDate(x.eventDate)}</div>${pill(x.status, STAT)}</li>`).join('')}</ul>
            ${deals.length + svcs.length + cls.length ? '' : '<p class="muted pad">沒有相關案件。</p>'}
          </div>
        </div></div>`;
    },
    deals() {
      const ds = Vault.all('deals');
      if (!ds.length) return empty('還沒有促約案件。從名單開始，記下每位潛在客戶的下一步。', '＋ 新增促約');
      return `<div class="kanban">${STAGES.map((s) => { const list = ds.filter((d) => (d.stage || '名單') === s).sort((a, b) => (a.nextDate || '9') < (b.nextDate || '9') ? -1 : 1); const sum = list.reduce((x, d) => x + (+d.premium || 0), 0);
        return `<div class="lane"><div class="lane-h"><b>${s}</b><span>${list.length}${sum ? '・' + money(sum) : ''}</span></div>${list.map((d) => `<div class="kcard" data-deal="${d.id}"><div class="kc-n">${clientName(d.clientId)}</div><div class="kc-t">${esc(d.title)}</div>${d.nextAction ? `<div class="kc-a">→ ${esc(d.nextAction)}</div>` : ''}<div class="kc-f">${dueTag(d.nextDate)}${d.premium ? `<span class="muted">${money(d.premium)}</span>` : ''}</div><div class="kc-mv"><button data-mv="-1" aria-label="上一階段">‹</button><button data-mv="1" aria-label="下一階段">›</button></div></div>`).join('')}</div>`; }).join('')}</div>`;
    },
    services() { return listCases('services', SERVICE_STATUS, (s) => `${esc(s.type)}${s.policyNo ? `<small>保單 ${pii(s.policyNo, 'num')}</small>` : ''}`, (s) => s.dueDate, '還沒有保全案件。', '＋ 新增保全'); },
    claims() { return listCases('claims', CLAIM_STATUS, (c) => { const n = (c.docs || []).length; return `${esc(c.kind)}${c.insured ? '・被保人 ' + pii(c.insured, 'name') : ''}<small>事故 ${fmtDate(c.eventDate) || '—'}・文件 ${n}/${CLAIM_DOCS.length}${c.paidAmount ? '・給付 ' + money(c.paidAmount) : ''}・照片 ${Files.forRef('claims', c.id).length}</small>`; }, (c) => c.followDate, '還沒有理賠案件。', '＋ 新增理賠'); },
    tasks() {
      const ts = Vault.all('tasks').sort((a, b) => (a.done - b.done) || ((a.due || '9') < (b.due || '9') ? -1 : 1));
      if (!ts.length) return empty('沒有行政待辦。', '＋ 新增待辦');
      const open = ts.filter((x) => !x.done), done = ts.filter((x) => x.done).slice(0, 30);
      const li = (x) => `<li class="task ${x.done ? 'done' : ''}"><input type="checkbox" data-done="${x.id}" ${x.done ? 'checked' : ''} aria-label="完成"><div class="main click" data-task="${x.id}">${esc(x.title)}<small>${esc(x.cat || '')}${x.clientId ? '・' + clientName(x.clientId) : ''}</small></div>${dueTag(x.due)}</li>`;
      return `<div class="panel"><h2>待辦（${open.length}）</h2><ul class="list">${open.map(li).join('') || '<p class="muted pad">全部完成了。</p>'}</ul></div>${done.length ? `<div class="panel"><h2>最近完成</h2><ul class="list">${done.map(li).join('')}</ul></div>` : ''}`;
    },
  };
  function listCases(coll, statuses, body, dateOf, emptyMsg, btn) {
    const xs = Vault.all(coll); if (!xs.length) return empty(emptyMsg, btn);
    const key = coll === 'services' ? 'svc' : 'claim';
    return statuses.map((st) => { const list = xs.filter((x) => (x.status || statuses[0]) === st).sort((a, b) => ((dateOf(a) || '9') < (dateOf(b) || '9') ? -1 : 1)); if (!list.length) return '';
      return `<div class="panel"><h2>${pill(st, STAT)} <span class="muted">${list.length}</span></h2><ul class="list">${list.map((x) => `<li class="click" data-${key}="${x.id}"><div class="main"><b>${clientName(x.clientId)}</b>・${body(x)}</div>${dueTag(dateOf(x))}</li>`).join('')}</ul></div>`; }).join('');
  }

  /* ================= 新增／編輯 ================= */
  const Edit = {
    client(c = {}) { editor(c.id ? '編輯客戶' : '新增客戶', F.client, c, (r) => { Vault.put('clients', r); if (!c.id) location.hash = '#client/' + r.id; }, { onDelete: c.id ? () => Vault.del('clients', c.id) : null }); },
    policy(p) { editor(p.id ? '編輯保單' : '新增保單', F.policy, p, (r) => Vault.put('policies', r), { onDelete: p.id ? () => Vault.del('policies', p.id) : null }); },
    log(l) { editor(l.id ? '編輯紀錄' : '新增往來紀錄', F.log, { date: today(), type: 'LINE', ...l }, (r) => Vault.put('logs', r), { onDelete: l.id ? () => Vault.del('logs', l.id) : null }); },
    deal(d = {}) { editor(d.id ? '促約案件' : '新增促約', F.deal, { stage: '名單', ...d }, (r) => Vault.put('deals', r), { onDelete: d.id ? () => Vault.del('deals', d.id) : null }); },
    service(s = {}) { editor(s.id ? '保全案件' : '新增保全', F.service, { status: '待辦', ...s }, (r) => Vault.put('services', r), { onDelete: s.id ? () => Vault.del('services', s.id) : null, html: s.id ? '<h3>照片／文件</h3>' + photoBlock('services', s.id) : '<p class="hint">儲存後可加入照片。</p>', wire: wirePhotos }); },
    claim(c = {}) { editor(c.id ? '理賠案件' : '新增理賠', F.claim, { status: '準備文件', docs: [], ...c }, (r) => Vault.put('claims', r), { onDelete: c.id ? () => Vault.del('claims', c.id) : null, html: c.id ? '<h3>照片／文件（收據、診斷書…）</h3>' + photoBlock('claims', c.id) : '<p class="hint">先儲存，再打開這筆理賠就能拍照加入收據和診斷書。</p>', wire: wirePhotos }); },
    task(t = {}) { editor(t.id ? '編輯待辦' : '新增待辦', F.task, { cat: '行政', ...t }, (r) => Vault.put('tasks', r), { onDelete: t.id ? () => Vault.del('tasks', t.id) : null }); },
  };

  /* ================= App ================= */
  const TITLES = { today: '今日', clients: '客戶', deals: '新契約促約', services: '保全服務', claims: '理賠', tasks: '行政待辦', client: '客戶資料' };
  const App = {
    route: 'today', arg: null, q: '',
    go() {
      const [r, a] = (location.hash.slice(1) || 'today').split('/');
      this.route = Views[r] ? r : 'today'; this.arg = a || null;
      document.querySelectorAll('#nav a').forEach((x) => x.classList.toggle('on', x.dataset.r === (this.route === 'client' ? 'clients' : this.route)));
      this.refresh(true);
    },
    refresh(scrollTop) {
      if (!Vault.db || $('app').classList.contains('hidden')) return;
      $('title').textContent = TITLES[this.route];
      $('btnBack').classList.toggle('hidden', this.route !== 'client');
      $('btnAdd').classList.toggle('hidden', this.route === 'client');
      const v = $('view'), y = v.scrollTop;
      v.innerHTML = Views[this.route](this.arg);
      if (this.route === 'client') wirePhotos(v);
      const q = $('q'); if (q) { q.oninput = () => { this.q = q.value; const pos = q.selectionStart; this.refresh(); const n = $('q'); n.focus(); n.setSelectionRange(pos, pos); }; }
      v.scrollTop = scrollTop === true ? 0 : y;
    },
  };
  window.addEventListener('hashchange', () => App.go());
  $('btnBack').onclick = () => history.length > 1 ? history.back() : (location.hash = '#clients');
  $('btnAdd').onclick = () => ({ today: () => quickAdd(), clients: () => Edit.client(), deals: () => Edit.deal(), services: () => Edit.service(), claims: () => Edit.claim(), tasks: () => Edit.task() }[App.route] || quickAdd)();
  function quickAdd() {
    const b = Sheet.open('新增', `<div class="qa">${[['client', '客戶'], ['deal', '促約'], ['service', '保全'], ['claim', '理賠'], ['task', '行政待辦']].map(([k, t]) => `<button class="primary" data-k="${k}">＋ ${t}</button>`).join('')}</div>`);
    b.onclick = (e) => { const x = e.target.closest('[data-k]'); if (x) { Sheet.close(); setTimeout(() => Edit[x.dataset.k]({}), 50); } };
  }
  $('view').addEventListener('click', (e) => {
    const t = e.target.closest('[data-c],[data-deal],[data-svc],[data-claim],[data-task],[data-p],[data-l],[data-i],[data-add],[data-edit-client],[data-add-policy],[data-add-log],[data-new],[data-mv]');
    if (!t) return;
    const d = t.dataset;
    if (d.mv) { e.stopPropagation(); const card = t.closest('[data-deal]'); const r = Vault.get('deals', card.dataset.deal); const i = Math.max(0, Math.min(STAGES.length - 1, STAGES.indexOf(r.stage || '名單') + +d.mv)); Vault.put('deals', { ...r, stage: STAGES[i] }); App.refresh(); return; }
    if (d.add !== undefined) return $('btnAdd').onclick();
    if (d.i !== undefined) return Views._items && Views._items[+d.i].open();
    if (d.c) { location.hash = '#client/' + d.c; return; }
    if (d.deal) return Edit.deal(Vault.get('deals', d.deal));
    if (d.svc) return Edit.service(Vault.get('services', d.svc));
    if (d.claim) return Edit.claim(Vault.get('claims', d.claim));
    if (d.task) return Edit.task(Vault.get('tasks', d.task));
    if (d.p) return Edit.policy(Vault.get('policies', d.p));
    if (d.l) return Edit.log(Vault.get('logs', d.l));
    if (d.editClient !== undefined) return Edit.client(Vault.get('clients', App.arg));
    if (d.addPolicy !== undefined) return Edit.policy({ clientId: App.arg, company: '國泰人壽', status: '有效', payMode: '年繳' });
    if (d.addLog !== undefined) return Edit.log({ clientId: App.arg });
    if (d.new) return Edit[d.new]({ clientId: App.arg });
  });
  $('view').addEventListener('change', (e) => { const x = e.target.closest('[data-done]'); if (x) { const r = Vault.get('tasks', x.dataset.done); Vault.put('tasks', { ...r, done: x.checked, doneAt: x.checked ? today() : '' }); App.refresh(); } });
  $('btnReveal').onclick = () => setRevealAll(!Mask.all);

  /* ================= 解鎖畫面 ================= */
  const Gate = {
    show(html) { $('app').classList.add('hidden'); const g = $('gate'); g.innerHTML = `<div class="gate-card"><div class="gate-logo">客戶管理</div>${html}</div>`; g.classList.remove('hidden'); return g; },
    hide() { $('gate').classList.add('hidden'); $('gate').innerHTML = ''; $('app').classList.remove('hidden'); },
    async start() {
      if (!CFG.CLIENT_ID && !MOCK) return this.show(`<p class="lead">還差一步設定：請把 Azure「保險 CRM」的應用程式識別碼交給 Claude，設定完成後這裡就會出現登入按鈕。</p>`);
      try { await Drive.init(); } catch (e) { return this.show(`<p class="lead">無法載入 Microsoft 登入，請確認網路後重新開啟。</p><p class="err">${esc(e.message)}</p>`); }
      if (!Drive.account) {
        const g = this.show(`<p class="lead">客戶資料會加密後存在<b>你自己的 OneDrive</b>（應用程式 → 保險 CRM），iPad 清空也不會不見。</p><button class="primary big" id="gSign">用 Microsoft 帳號登入 OneDrive</button>`);
        g.querySelector('#gSign').onclick = () => Drive.signIn();
        return;
      }
      let exists = null;
      try { exists = !!(await Drive.getMeta('data.enc')); } catch (e) { exists = !!(await Cache.get('data')); }
      if (exists) this.unlock(); else this.setup();
    },
    unlock() {
      const g = this.show(`<p class="lead">輸入 CRM 密碼解鎖</p><form id="gForm"><input id="gPw" type="password" class="big" autocomplete="current-password" placeholder="密碼"><p class="err" id="gErr"></p><button class="primary big" type="submit">解鎖</button></form><p class="muted small">已登入：${esc(Drive.account.username || '')}　<a href="#" id="gOut">換帳號</a></p>`);
      const pw = g.querySelector('#gPw'); setTimeout(() => pw.focus(), 50);
      g.querySelector('#gOut').onclick = (e) => { e.preventDefault(); Cache.clear(); Drive.signOut(); };
      g.querySelector('#gForm').onsubmit = async (e) => {
        e.preventDefault(); const err = g.querySelector('#gErr'); err.textContent = '解鎖中…';
        try { const r = await Vault.load(pw.value); if (r === 'new') return this.setup(); this.enter(); }
        catch (x) { err.textContent = x.badPassword ? '密碼不正確，請再試一次。' : '讀取 OneDrive 失敗：' + x.message; }
      };
    },
    setup() {
      const g = this.show(`<p class="lead">第一次使用：設定 CRM 密碼</p><p class="warn">所有客戶資料都用這組密碼加密。<b>忘記密碼就無法打開資料，沒有人能幫你找回</b>，請記在安全的地方。</p><form id="gForm"><input id="gPw" type="password" class="big" autocomplete="new-password" placeholder="設定密碼（至少 8 個字）"><input id="gPw2" type="password" class="big" autocomplete="new-password" placeholder="再輸入一次"><label class="chk"><input type="checkbox" id="gOk"> 我已經把密碼記下來了</label><p class="err" id="gErr"></p><button class="primary big" type="submit">建立加密資料庫</button></form>`);
      g.querySelector('#gForm').onsubmit = async (e) => {
        e.preventDefault(); const a = g.querySelector('#gPw').value, b = g.querySelector('#gPw2').value, err = g.querySelector('#gErr');
        if (a.length < 8) return (err.textContent = '密碼至少 8 個字。');
        if (a !== b) return (err.textContent = '兩次輸入的密碼不一樣。');
        if (!g.querySelector('#gOk').checked) return (err.textContent = '請先確認已記下密碼。');
        err.textContent = '建立中…';
        try { await Vault.create(a); this.enter(); } catch (x) { err.textContent = '建立失敗：' + x.message; }
      };
    },
    enter() { this.hide(); setRevealAll(false); App.go(); Idle.reset(); Files.retryUploads(); try { navigator.storage && navigator.storage.persist && navigator.storage.persist(); } catch (e) { } },
  };
  function lockNow() { Vault.save(); Vault.lock(); Files.mem.forEach((u) => URL.revokeObjectURL(u)); Files.mem.clear(); Sheet.close(); Viewer.close(); $('view').innerHTML = ''; Gate.unlock(); }
  $('btnLock').onclick = lockNow;
  const Idle = { reset() { clearTimeout(this.t); this.t = setTimeout(() => { if (Vault.key) { lockNow(); } }, (CFG.AUTO_LOCK_MINUTES || 10) * 60000); } };
  ['pointerdown', 'keydown'].forEach((ev) => document.addEventListener(ev, () => Vault.key && Idle.reset(), { passive: true }));
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'hidden' && Vault.dirty) Vault.save(); });
  window.addEventListener('online', () => { if (Vault.key) { Vault.save(true); Files.retryUploads(); } });

  /* ================= 設定 ================= */
  $('btnSettings').onclick = () => {
    const n = (c) => Vault.all(c).length;
    const b = Sheet.open('設定', `<div class="settings">
      <p class="muted">已登入：${esc((Drive.account || {}).username || '')}<br>資料位置：OneDrive → 應用程式 → 保險 CRM（加密檔）<br>客戶 ${n('clients')}・保單 ${n('policies')}・促約 ${n('deals')}・保全 ${n('services')}・理賠 ${n('claims')}・照片 ${n('files')}</p>
      <button class="primary" id="sSync">立即同步</button>
      <button class="ghostbtn" id="sCsv">匯出客戶清單（Excel 可開）</button>
      <button class="ghostbtn" id="sPw">更改 CRM 密碼</button>
      <button class="ghostbtn" id="sClear">清除這台 iPad 上的暫存</button>
      <button class="danger" id="sOut">登出 OneDrive</button>
      <p class="muted small">這台 iPad 只暫存加密檔；資料以 OneDrive 為準。閒置 ${CFG.AUTO_LOCK_MINUTES || 10} 分鐘會自動鎖定。</p></div>`);
    b.querySelector('#sSync').onclick = async () => { await Vault.save(true); Files.retryUploads(); toast(Vault.state === 'synced' ? '已同步' : '目前離線，稍後自動同步'); };
    b.querySelector('#sCsv').onclick = () => reauth('匯出客戶清單', exportCsv);
    b.querySelector('#sPw').onclick = () => changePw();
    b.querySelector('#sClear').onclick = () => confirmBox('清除這台 iPad 上的暫存？雲端資料不受影響，下次開啟會從 OneDrive 重新下載。', async () => { await Vault.save(true); await Cache.clear(); toast('已清除暫存'); }, '清除', false);
    b.querySelector('#sOut').onclick = () => confirmBox('登出後這台 iPad 的暫存也會清除，資料仍保留在 OneDrive。', async () => { await Vault.save(true); await Cache.clear(); Vault.lock(); Drive.signOut(); }, '登出');
  };
  function reauth(title, then) {
    const b = Sheet.open(title, `<p class="lead">請再輸入一次 CRM 密碼</p><form id="rf"><input id="rp" type="password" class="big" autocomplete="current-password"><p class="err" id="re"></p><button class="primary" type="submit">確認</button></form>`);
    const p = b.querySelector('#rp'); setTimeout(() => p.focus(), 50);
    b.querySelector('#rf').onsubmit = async (e) => { e.preventDefault(); try { const k = await deriveKey(p.value, Vault.salt); const probe = await seal(Vault.key, enc.encode('x')); await open(k, probe); Sheet.close(); then(p.value); } catch (x) { b.querySelector('#re').textContent = '密碼不正確。'; } };
  }
  function exportCsv() {
    const cols = [['name', '姓名'], ['gender', '性別'], ['birthday', '生日'], ['idNo', '身分證'], ['phone', '手機'], ['line', 'LINE'], ['email', 'Email'], ['address', '地址'], ['job', '職業'], ['source', '來源'], ['tags', '標籤'], ['note', '備註']];
    const q = (s) => '"' + String(s ?? '').replace(/"/g, '""') + '"';
    const lines = [cols.map((c) => c[1]).join(',')].concat(Vault.all('clients').map((c) => cols.map(([k]) => q(c[k])).join(',')));
    const blob = new Blob(['﻿' + lines.join('\r\n')], { type: 'text/csv' }), name = `客戶清單_${today()}.csv`, file = new File([blob], name, { type: 'text/csv' });
    if (navigator.canShare && navigator.canShare({ files: [file] })) navigator.share({ files: [file] }).catch(() => { });
    else { const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; a.click(); }
    toast('檔案含完整個資，用完請刪除');
  }
  function changePw() {
    const b = Sheet.open('更改 CRM 密碼', `<form id="cp"><input id="c0" type="password" class="big" placeholder="目前密碼" autocomplete="current-password"><input id="c1" type="password" class="big" placeholder="新密碼（至少 8 個字）" autocomplete="new-password"><input id="c2" type="password" class="big" placeholder="再輸入一次新密碼" autocomplete="new-password"><p class="err" id="ce"></p><button class="primary" type="submit">更改密碼</button></form><p class="muted small">更改時會把所有照片重新加密上傳，照片多時需要一點時間，請保持網路連線、不要關閉 App。</p>`);
    b.querySelector('#cp').onsubmit = async (e) => {
      e.preventDefault(); const er = b.querySelector('#ce');
      const [o, n1, n2] = ['c0', 'c1', 'c2'].map((i) => b.querySelector('#' + i).value);
      if (n1.length < 8) return (er.textContent = '新密碼至少 8 個字。'); if (n1 !== n2) return (er.textContent = '兩次新密碼不一樣。');
      try { const k = await deriveKey(o, Vault.salt); await open(k, await seal(Vault.key, enc.encode('x'))); } catch (x) { return (er.textContent = '目前密碼不正確。'); }
      // 註：AES-GCM 金鑰無法直接比對，上面用「舊密碼重新推導的金鑰」解目前金鑰封裝的資料來驗證
      er.textContent = '重新加密中…';
      try {
        // 先把照片用新密碼另存新檔，全部成功後才切換資料庫密碼，最後刪除舊檔；中途失敗時舊密碼仍可用
        const oldKey = Vault.key, oldSalt = Vault.salt, salt = crypto.getRandomValues(new Uint8Array(16)), key = await deriveKey(n1, salt);
        const kv = Date.now().toString(36), fs = Vault.all('files'), done = []; let i = 0;
        for (const f of fs) {
          er.textContent = `重新加密照片 ${++i}/${fs.length}…`;
          let s = await Cache.get('file:' + f.id); if (!s) { const r = await Drive.download(fpath(f)); if (!r) continue; s = r.bytes; }
          const ns = await seal(key, await open(oldKey, s)); await Drive.upload(fpath({ id: f.id, kv }), ns); done.push({ f, ns });
        }
        const oldDb = JSON.parse(JSON.stringify(Vault.db));
        done.forEach(({ f }) => { Vault.db.files[f.id] = { ...f, kv, updatedAt: new Date().toISOString() }; });
        Vault.key = key; Vault.salt = salt; Vault.dirty = true; await Vault.save(true);
        if (Vault.state !== 'synced') { Vault.key = oldKey; Vault.salt = oldSalt; Vault.db = oldDb; throw new Error('無法存到 OneDrive'); }
        for (const { f, ns } of done) { await Cache.set('file:' + f.id, ns); Drive.remove(fpath(f)).catch(() => { }); }
        await Cache.del('pending');
        Sheet.close(); toast('密碼已更改，請記好新密碼', 3500);
      } catch (x) { er.textContent = '更改失敗（' + x.message + '），請在網路穩定時再試；舊密碼仍然有效。'; }
    };
  }

  if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => { });
  Gate.start();
})();
