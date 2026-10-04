/* 保險手冊 PWA */
(() => {
  'use strict';
  const H = window.HANDBOOK;
  const $ = (id) => document.getElementById(id);
  const LS = {
    get(k, d) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (e) { return d; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { } },
    del(k) { try { localStorage.removeItem(k); } catch (e) { } },
  };
  const CH = H.chapters;
  const NUM = { '一': 1, '二': 2, '三': 3 };
  const app = $('app'), vp = $('viewport'), stage = $('stage'), host = $('slideHost');
  let cur = { c: 0, s: 0 };
  const backStack = [];
  let scale = 1;

  /* ---------- 小工具 ---------- */
  function toast(msg, ms = 1800) {
    const t = $('toast'); t.textContent = msg; t.classList.remove('hidden');
    clearTimeout(toast._t); toast._t = setTimeout(() => t.classList.add('hidden'), ms);
  }
  function esc(s) { return String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
  const keyOf = (c, s) => CH[c].n + '/' + CH[c].slides[s].id;
  function locate(target) { // "2/hub" → {c,s}
    const [n, id] = String(target).split('/');
    const c = CH.findIndex((x) => String(x.n) === n);
    if (c < 0) return null;
    const s = id ? CH[c].slides.findIndex((x) => x.id === id) : 0;
    return s < 0 ? null : { c, s };
  }
  function pageTarget(n, page) { // 第 n 章第 page 頁 → "n/id"
    const ch = CH.find((x) => x.n === n);
    const sl = ch && ch.slides[page - 1];
    return sl ? n + '/' + sl.id : null;
  }

  /* ---------- 版面縮放 ---------- */
  function fit() {
    const w = vp.clientWidth, h = vp.clientHeight;
    scale = Math.min(w / 1920, h / 1080);
    const x = (w - 1920 * scale) / 2, y = (h - 1080 * scale) / 2;
    stage.style.transform = `translate(${x}px,${y}px) scale(${scale})`;
    Ink.resize();
    if (!$('board').classList.contains('hidden')) Board.resize();
  }
  window.addEventListener('resize', fit);
  window.addEventListener('orientationchange', () => setTimeout(fit, 250));

  /* ---------- 投影片強化：跳頁連結 ---------- */
  const PAINTED = /background|border/;
  function paintedAncestor(el, sec) {
    for (let e = el.parentElement; e && e !== sec; e = e.parentElement) {
      if (e.tagName === 'DIV' && PAINTED.test(e.getAttribute('style') || '')) return e;
    }
    return null;
  }
  function chapterHint(el) {
    const blk = el.closest('p,li,td,th,h1,h2,h3') || el;
    const m = (blk.textContent || '').match(/第([一二三])章/);
    return m ? NUM[m[1]] : null;
  }
  function enhance(sec, c) {
    const chN = CH[c].n, selfId = CH[c].slides[cur.s].id;
    // 表格：一個儲存格設了 padding，就套用到整張表
    sec.querySelectorAll('table').forEach((t) => {
      const cell = [...t.querySelectorAll('td,th')].find((x) => x.style.padding);
      if (cell) t.querySelectorAll('td,th').forEach((x) => { if (!x.style.padding) x.style.padding = cell.style.padding; });
    });
    // 連結
    sec.querySelectorAll('a[href]').forEach((a) => {
      const h = a.getAttribute('href');
      if (h.startsWith('#catalog/')) { a.addEventListener('click', (e) => { e.preventDefault(); Catalog.open(h.slice(9)); }); }
      else if (/^https?:/.test(h)) { a.target = '_blank'; a.rel = 'noopener'; }
    });
    // P# → 可點
    const walker = document.createTreeWalker(sec, NodeFilter.SHOW_TEXT);
    const nodes = [];
    while (walker.nextNode()) if (/P\d/.test(walker.currentNode.nodeValue)) nodes.push(walker.currentNode);
    nodes.forEach((node) => {
      const target = chapterHint(node.parentElement) || chN;
      const parts = node.nodeValue.split(/(P\d{1,2}(?:[–-]P?\d{1,2})?)/);
      if (parts.length < 2) return;
      const frag = document.createDocumentFragment();
      parts.forEach((part) => {
        const m = part.match(/^P(\d{1,2})/);
        const to = m && pageTarget(target, +m[1]);
        if (to && to !== chN + '/' + selfId) {
          const sp = document.createElement('span');
          sp.className = 'jump'; sp.dataset.to = to; sp.textContent = part;
          frag.appendChild(sp);
        } else frag.appendChild(document.createTextNode(part));
      });
      node.parentNode.replaceChild(frag, node);
    });
    // 整張卡片可點：卡片裡只有一個跳頁目標、沒有其他連結
    sec.querySelectorAll('.jump').forEach((j) => {
      const card = paintedAncestor(j, sec);
      if (!card || card.dataset.go || card.querySelector('a')) return;
      const tos = new Set([...card.querySelectorAll('.jump')].map((x) => x.dataset.to));
      if (tos.size === 1) card.dataset.go = j.dataset.to;
    });
    // 章節與 5W2H 卡片
    const W = { 'WHY': '1/why', 'WHAT': '1/what', 'HOW': '1/how', 'HOW MUCH': '1/howmuch', 'WHO': '3/who', 'WHEN': '3/when', 'WHERE': '3/where' };
    sec.querySelectorAll('p,h3').forEach((p) => {
      const t = (p.textContent || '').trim();
      let to = null;
      if (selfId === 'overview' && W[t]) to = W[t];
      else if (/四大金律/.test(t)) to = '1/how';
      else if (/^第([一二三])章/.test(t)) to = NUM[t[1]] + '/';
      else if (selfId === 'toc' && /^0[123]$/.test(t)) to = (+t) + '/';
      if (!to) return;
      to = to.endsWith('/') ? to + CH.find((x) => x.n === +to[0]).slides[0].id : to;
      if (to === chN + '/' + selfId) return;
      const card = paintedAncestor(p, sec);
      const el = card && !card.dataset.go && !card.querySelector('.jump,a') ? card : p;
      if (!el.dataset.go) el.dataset.go = to;
    });
  }
  host.addEventListener('click', (e) => {
    if (Ink.justDrew()) return;
    if (app.classList.contains('edit-hs')) return;
    const j = e.target.closest('.jump,[data-go]');
    if (!j) return;
    e.preventDefault();
    go(j.dataset.to || j.dataset.go, true);
  });

  /* ---------- 導覽 ---------- */
  function show(c, s, opts = {}) {
    if (opts.push) backStack.push({ ...cur });
    cur = { c, s };
    const sl = CH[c].slides[s];
    host.innerHTML = sl.html;
    const sec = host.firstElementChild;
    sec.classList.add('slide');
    if (!opts.noAnim) { sec.classList.add('fade-in'); }
    enhance(sec, c);
    Hotspots.render();
    Ink.load(keyOf(c, s));
    $('whereCh').textContent = CH[c].short + '｜' + CH[c].title;
    $('whereTitle').textContent = sl.title;
    $('pageNo').textContent = `${CH[c].short}　P${s + 1} / ${CH[c].slides.length}`;
    $('btnBack').classList.toggle('hidden', backStack.length === 0);
    Notes.update();
    try { history.replaceState(null, '', '#' + keyOf(c, s)); } catch (e) { }
  }
  function go(target, push) {
    const p = locate(target);
    if (p) show(p.c, p.s, { push });
  }
  function step(d) {
    let { c, s } = cur;
    s += d;
    if (s >= CH[c].slides.length) { if (c < CH.length - 1) { c++; s = 0; } else return; }
    if (s < 0) { if (c > 0) { c--; s = CH[c].slides.length - 1; } else return; }
    show(c, s);
  }
  $('btnPrev').onclick = () => step(-1);
  $('btnNext').onclick = () => step(1);
  $('btnHome').onclick = () => show(0, 0, { push: true });
  $('btnBack').onclick = () => { const p = backStack.pop(); if (p) show(p.c, p.s); };
  document.addEventListener('keydown', (e) => {
    if (e.target.matches('input,textarea')) return;
    if (!$('sheet').classList.contains('hidden')) { if (e.key === 'Escape') Sheet.close(); return; }
    if (['ArrowRight', 'PageDown', ' '].includes(e.key)) { e.preventDefault(); step(1); }
    if (['ArrowLeft', 'PageUp'].includes(e.key)) { e.preventDefault(); step(-1); }
  });

  /* ---------- 彈出面板 ---------- */
  const Sheet = {
    open(title, body) {
      $('sheetTitle').textContent = title;
      const b = $('sheetBody'); b.innerHTML = ''; if (typeof body === 'string') b.innerHTML = body; else b.appendChild(body);
      $('sheet').classList.remove('hidden'); b.scrollTop = 0; return b;
    },
    close() { $('sheet').classList.add('hidden'); },
  };
  $('sheetClose').onclick = Sheet.close;
  $('sheet').addEventListener('click', (e) => { if (e.target.id === 'sheet') Sheet.close(); });
  function confirmBox(title, msg, okText, onOk, danger) {
    const b = Sheet.open(title, `<p class="note" style="font-size:16px">${esc(msg)}</p><div class="row" style="margin-top:14px"><button class="btn ${danger ? 'danger' : ''}" id="cfOk">${esc(okText)}</button><button class="btn ghost" id="cfNo">取消</button></div>`);
    b.querySelector('#cfOk').onclick = () => { Sheet.close(); onOk(); };
    b.querySelector('#cfNo').onclick = Sheet.close;
  }

  /* ---------- 目錄 ---------- */
  function slidePicker(onPick, highlight) {
    const wrap = document.createElement('div');
    CH.forEach((ch, c) => {
      const sec = document.createElement('div'); sec.className = 'toc-ch';
      sec.innerHTML = `<h3>${esc(ch.short)}｜${esc(ch.title)}</h3>`;
      const list = document.createElement('div'); list.className = 'toc-list';
      ch.slides.forEach((sl, s) => {
        const b = document.createElement('button'); b.className = 'toc-item' + (highlight && c === cur.c && s === cur.s ? ' cur' : '');
        b.innerHTML = `<b>P${s + 1}</b><span>${esc(sl.title || sl.id)}</span>`;
        b.onclick = () => onPick(c, s);
        list.appendChild(b);
      });
      sec.appendChild(list); wrap.appendChild(sec);
    });
    return wrap;
  }
  $('btnToc').onclick = () => {
    Sheet.open('目錄', slidePicker((c, s) => { Sheet.close(); show(c, s, { push: true }); }, true));
    const el = $('sheetBody').querySelector('.toc-item.cur'); if (el) el.scrollIntoView({ block: 'center' });
  };

  /* ---------- 畫筆 ---------- */
  const Tool = { tool: 'pen', color: '#C9962E', finger: LS.get('fingerDraw', false) };
  function setTool(btn) {
    document.querySelectorAll('#penbar .pen').forEach((b) => b.classList.toggle('on', b === btn));
    Tool.tool = btn.dataset.tool; Tool.color = btn.dataset.color || Tool.color;
  }
  document.querySelectorAll('#penbar .pen').forEach((b) => b.onclick = () => setTool(b));
  setTool(document.querySelector('#penbar .pen'));
  $('btnFinger').classList.toggle('on', Tool.finger);
  $('btnFinger').onclick = () => {
    Tool.finger = !Tool.finger; LS.set('fingerDraw', Tool.finger);
    $('btnFinger').classList.toggle('on', Tool.finger);
    toast(Tool.finger ? '手指也能畫（滑動換頁暫停）' : '手指滑動換頁，Apple Pencil 畫圖');
  };
  function strokeStyle(ctx, st) {
    ctx.lineCap = 'round'; ctx.lineJoin = 'round';
    if (st.t === 'hl') { ctx.globalAlpha = 0.38; ctx.strokeStyle = st.c; ctx.lineWidth = st.w; ctx.lineCap = 'butt'; }
    else { ctx.globalAlpha = 1; ctx.strokeStyle = st.c; ctx.lineWidth = st.w; }
  }
  function drawStroke(ctx, st) {
    const p = st.p; if (p.length < 2) return;
    strokeStyle(ctx, st); ctx.beginPath(); ctx.moveTo(p[0], p[1]);
    if (p.length === 2) { ctx.lineTo(p[0] + 0.1, p[1] + 0.1); }
    for (let i = 2; i < p.length - 2; i += 2) {
      const mx = (p[i] + p[i + 2]) / 2, my = (p[i + 1] + p[i + 3]) / 2;
      ctx.quadraticCurveTo(p[i], p[i + 1], mx, my);
    }
    if (p.length > 2) ctx.lineTo(p[p.length - 2], p[p.length - 1]);
    ctx.stroke(); ctx.globalAlpha = 1;
  }
  function hitErase(strokes, x, y, r) {
    return strokes.filter((st) => { for (let i = 0; i < st.p.length; i += 2) { if (Math.hypot(st.p[i] - x, st.p[i + 1] - y) < r + st.w / 2) return false; } return true; });
  }
  function makeSurface(canvas, opts) {
    // opts: width(), height(), resolution(), storeKey(), toLocal(e)
    const S = { strokes: [], undo: [], key: null, live: null };
    const ctx = canvas.getContext('2d');
    S.resize = () => {
      const k = opts.resolution();
      canvas.width = Math.round(opts.width() * k); canvas.height = Math.round(opts.height() * k);
      S.k = k; S.redraw();
    };
    S.redraw = () => {
      ctx.setTransform(1, 0, 0, 1, 0, 0); ctx.clearRect(0, 0, canvas.width, canvas.height);
      ctx.setTransform(S.k, 0, 0, S.k, 0, 0);
      S.strokes.forEach((st) => drawStroke(ctx, st));
      if (S.live) drawStroke(ctx, S.live);
    };
    S.load = (key) => { S.key = key; S.strokes = LS.get('ink:' + key, []); S.undo = []; S.redraw(); };
    S.save = () => { if (S.strokes.length) LS.set('ink:' + S.key, S.strokes); else LS.del('ink:' + S.key); };
    S.snapshot = () => { S.undo.push(JSON.stringify(S.strokes)); if (S.undo.length > 40) S.undo.shift(); };
    S.doUndo = () => { const u = S.undo.pop(); if (u == null) { toast('沒有可以復原的筆畫'); return; } S.strokes = JSON.parse(u); S.save(); S.redraw(); };
    S.clear = () => { if (!S.strokes.length) return; S.snapshot(); S.strokes = []; S.save(); S.redraw(); };
    S.begin = (x, y, pressure) => {
      S.snapshot();
      if (Tool.tool === 'eraser') { S.erasing = true; S.erase(x, y); return; }
      const w = (Tool.tool === 'hl' ? 34 : 6) * (opts.unit ? opts.unit() : 1);
      S.live = { t: Tool.tool, c: Tool.color, w, p: [Math.round(x), Math.round(y)] };
      S.redraw();
    };
    S.move = (x, y) => {
      if (S.erasing) { S.erase(x, y); return; }
      if (!S.live) return;
      const p = S.live.p, lx = p[p.length - 2], ly = p[p.length - 1];
      if (Math.hypot(x - lx, y - ly) < 1.5) return;
      p.push(Math.round(x), Math.round(y));
      ctx.setTransform(S.k, 0, 0, S.k, 0, 0);
      if (S.live.t === 'hl') S.redraw();
      else { strokeStyle(ctx, S.live); ctx.beginPath(); ctx.moveTo(lx, ly); ctx.lineTo(x, y); ctx.stroke(); }
    };
    S.end = () => {
      if (S.erasing) { S.erasing = false; S.save(); return; }
      if (S.live) { S.strokes.push(S.live); S.live = null; S.save(); S.redraw(); }
    };
    S.cancel = () => { S.live = null; S.erasing = false; const u = S.undo.pop(); if (u != null) S.strokes = JSON.parse(u); S.redraw(); };
    S.erase = (x, y) => { const n = hitErase(S.strokes, x, y, 18 * (opts.unit ? opts.unit() : 1)); if (n.length !== S.strokes.length) { S.strokes = n; S.redraw(); } };
    return S;
  }

  /* 投影片上的筆畫（座標 = 投影片 1920×1080） */
  const inkCanvas = $('ink');
  const Ink = makeSurface(inkCanvas, {
    width: () => 1920, height: () => 1080,
    resolution: () => Math.min(2, Math.max(1, scale * (window.devicePixelRatio || 1))),
  });
  inkCanvas.style.pointerEvents = 'none';
  let drewAt = 0;
  Ink.justDrew = () => Date.now() - drewAt < 350;
  function toSlide(e) {
    const r = stage.getBoundingClientRect();
    return { x: (e.clientX - r.left) / scale, y: (e.clientY - r.top) / scale };
  }

  // 指標事件：Apple Pencil 畫圖、手指滑動換頁（或手指也畫）
  let gest = null;
  vp.addEventListener('pointerdown', (e) => {
    if (e.button > 0) return;
    if (app.classList.contains('edit-hs')) { Hotspots.down(e); return; }
    const draw = e.pointerType === 'pen' || Tool.finger;
    gest = { id: e.pointerId, type: e.pointerType, x0: e.clientX, y0: e.clientY, t0: Date.now(), draw, started: false };
    if (draw) {
      try { vp.setPointerCapture(e.pointerId); } catch (err) { }
      const p = toSlide(e); gest.p0 = p;
    }
  });
  vp.addEventListener('pointermove', (e) => {
    if (app.classList.contains('edit-hs')) { Hotspots.move(e); return; }
    if (!gest || e.pointerId !== gest.id || !gest.draw) return;
    const evs = e.getCoalescedEvents ? e.getCoalescedEvents() : [e];
    if (!gest.started) {
      if (Math.hypot(e.clientX - gest.x0, e.clientY - gest.y0) < 4) return;
      gest.started = true; Ink.begin(gest.p0.x, gest.p0.y);
    }
    evs.forEach((ev) => { const p = toSlide(ev); Ink.move(p.x, p.y); });
    e.preventDefault();
  });
  function endGesture(e, cancelled) {
    if (app.classList.contains('edit-hs')) { Hotspots.up(e); return; }
    if (!gest || e.pointerId !== gest.id) return;
    const g = gest; gest = null;
    if (g.draw) {
      if (g.started) { if (cancelled) Ink.cancel(); else Ink.end(); drewAt = Date.now(); }
      else if (Date.now() - g.t0 > 300 && Tool.tool !== 'eraser') { // 長按一點：畫一個點
        Ink.begin(g.p0.x, g.p0.y); Ink.move(g.p0.x + 0.5, g.p0.y + 0.5); Ink.end(); drewAt = Date.now();
      }
      return;
    }
    const dx = e.clientX - g.x0, dy = e.clientY - g.y0;
    if (!cancelled && Math.abs(dx) > 60 && Math.abs(dx) > Math.abs(dy) * 1.4 && Date.now() - g.t0 < 900) {
      drewAt = Date.now(); step(dx < 0 ? 1 : -1);
    }
  }
  vp.addEventListener('pointerup', (e) => endGesture(e, false));
  vp.addEventListener('pointercancel', (e) => endGesture(e, true));
  vp.addEventListener('touchmove', (e) => e.preventDefault(), { passive: false });

  const activeSurface = () => (isBoard() ? Board : (Doc.isOpen() ? Doc.active() : Ink));
  $('btnUndo').onclick = () => { const S = activeSurface(); if (S) S.doUndo(); else toast('這一頁還沒有筆畫'); };
  $('btnClearPage').onclick = () => {
    const B = isBoard(), S = activeSurface();
    if (!S) { toast('這一頁還沒有筆畫'); return; }
    confirmBox(B ? '清空白板' : '清除本頁畫記', B ? '白板上的筆畫會全部清除。' : '這一頁的筆畫會全部清除，其他頁不受影響。', '清除', () => S.clear(), true);
  };

  /* ---------- 存成圖片 ---------- */
  function loadScript(src) {
    return new Promise((ok, bad) => { const s = document.createElement('script'); s.src = src; s.onload = ok; s.onerror = bad; document.head.appendChild(s); });
  }
  async function showImage(canvas, name) {
    const url = canvas.toDataURL('image/png');
    const blob = await (await fetch(url)).blob();
    const file = new File([blob], name + '.png', { type: 'image/png' });
    const b = Sheet.open('存成圖片', `<img src="${url}" alt="${esc(name)}" style="width:100%;border-radius:12px;border:1px solid #D8D3C4"><p class="note">在 iPad 上長按圖片，選「加入照片」即可儲存；也可以用「分享」傳給客戶。</p><div class="row"><button class="btn" id="imgShare">分享</button></div>`);
    const sh = b.querySelector('#imgShare');
    if (navigator.canShare && navigator.canShare({ files: [file] })) sh.onclick = () => navigator.share({ files: [file], title: name }).catch(() => { });
    else sh.classList.add('hidden');
  }
  async function saveSlideImage() {
    toast('製作圖片中…', 4000);
    try {
      if (!window.html2canvas) await loadScript('https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js');
      const sec = host.firstElementChild;
      const prev = stage.style.transform; stage.style.transform = 'none';
      const shot = await window.html2canvas(sec, { scale: 1, width: 1920, height: 1080, backgroundColor: null, useCORS: true, logging: false });
      stage.style.transform = prev;
      const out = document.createElement('canvas'); out.width = 1920; out.height = 1080;
      const ctx = out.getContext('2d'); ctx.drawImage(shot, 0, 0, 1920, 1080);
      ctx.drawImage(inkCanvas, 0, 0, 1920, 1080);
      $('toast').classList.add('hidden');
      showImage(out, CH[cur.c].short + '_P' + (cur.s + 1));
    } catch (err) { fit(); toast('需要連上網路一次才能製作圖片'); }
  }

  /* ---------- 白板 ---------- */
  const boardCanvas = $('boardInk');
  const Board = makeSurface(boardCanvas, {
    width: () => boardCanvas.clientWidth, height: () => boardCanvas.clientHeight,
    resolution: () => window.devicePixelRatio || 1, unit: () => 0.6,
  });
  const isBoard = () => !$('board').classList.contains('hidden');
  $('btnBoard').onclick = () => {
    $('board').classList.remove('hidden'); $('btnBoard').classList.add('on');
    Board.resize(); Board.load('board');
  };
  $('boardClose').onclick = () => { $('board').classList.add('hidden'); $('btnBoard').classList.remove('on'); };
  $('boardUndo').onclick = () => Board.doUndo();
  $('boardClear').onclick = () => confirmBox('清空白板', '白板上的筆畫會全部清除。', '清空', () => Board.clear(), true);
  $('boardSave').onclick = () => {
    const out = document.createElement('canvas'); out.width = boardCanvas.width; out.height = boardCanvas.height;
    const ctx = out.getContext('2d'); ctx.fillStyle = '#FBFAF6'; ctx.fillRect(0, 0, out.width, out.height); ctx.drawImage(boardCanvas, 0, 0);
    showImage(out, '白板');
  };
  let bg = null;
  boardCanvas.addEventListener('pointerdown', (e) => {
    if (e.button > 0) return; bg = e.pointerId;
    try { boardCanvas.setPointerCapture(e.pointerId); } catch (err) { }
    const r = boardCanvas.getBoundingClientRect(); Board.begin(e.clientX - r.left, e.clientY - r.top);
  });
  boardCanvas.addEventListener('pointermove', (e) => {
    if (e.pointerId !== bg) return;
    const r = boardCanvas.getBoundingClientRect();
    (e.getCoalescedEvents ? e.getCoalescedEvents() : [e]).forEach((ev) => Board.move(ev.clientX - r.left, ev.clientY - r.top));
  });
  boardCanvas.addEventListener('pointerup', (e) => { if (e.pointerId === bg) { bg = null; Board.end(); } });
  boardCanvas.addEventListener('pointercancel', (e) => { if (e.pointerId === bg) { bg = null; Board.cancel(); } });
  boardCanvas.addEventListener('touchmove', (e) => e.preventDefault(), { passive: false });

  /* ---------- 熱區（可自由框選，點了跳頁） ---------- */
  const Hotspots = {
    all() { const base = H.hotspots || {}; const mine = LS.get('hotspots', {}); const k = keyOf(cur.c, cur.s); return { base: base[k] || [], mine: mine[k] || [] }; },
    render() {
      const box = $('hotspots'); box.innerHTML = '';
      const { base, mine } = this.all();
      [...base.map((h) => ({ ...h, baked: true })), ...mine.map((h, i) => ({ ...h, i }))].forEach((h) => {
        const d = document.createElement('div'); d.className = 'hs';
        Object.assign(d.style, { left: h.x + 'px', top: h.y + 'px', width: h.w + 'px', height: h.h + 'px' });
        const p = locate(h.to); const lab = p ? `${CH[p.c].short} P${p.s + 1}` : h.to;
        d.innerHTML = `<span>→ ${esc(lab)}</span>` + (h.baked ? '' : '<b>×</b>');
        d.addEventListener('click', (e) => {
          if (app.classList.contains('edit-hs')) {
            if (!h.baked && e.target.tagName === 'B') { const all = LS.get('hotspots', {}); const k = keyOf(cur.c, cur.s); all[k].splice(h.i, 1); LS.set('hotspots', all); this.render(); }
            return;
          }
          if (Ink.justDrew()) return;
          go(h.to, true);
        });
        box.appendChild(d);
      });
    },
    down(e) {
      if (e.target.closest('.hs')) return;
      const p = toSlide(e); this.d = { x0: p.x, y0: p.y, id: e.pointerId };
      const el = document.createElement('div'); el.id = 'hsDraft'; stage.appendChild(el); this.el = el;
      try { vp.setPointerCapture(e.pointerId); } catch (err) { }
    },
    move(e) {
      if (!this.d || e.pointerId !== this.d.id) return;
      const p = toSlide(e), d = this.d;
      Object.assign(this.el.style, { left: Math.min(d.x0, p.x) + 'px', top: Math.min(d.y0, p.y) + 'px', width: Math.abs(p.x - d.x0) + 'px', height: Math.abs(p.y - d.y0) + 'px' });
    },
    up(e) {
      if (!this.d || e.pointerId !== this.d.id) return;
      const p = toSlide(e), d = this.d; this.d = null; this.el.remove();
      const r = { x: Math.round(Math.min(d.x0, p.x)), y: Math.round(Math.min(d.y0, p.y)), w: Math.round(Math.abs(p.x - d.x0)), h: Math.round(Math.abs(p.y - d.y0)) };
      if (r.w < 30 || r.h < 30) return;
      const k = keyOf(cur.c, cur.s);
      const pick = slidePicker((c, s) => {
        Sheet.close();
        const all = LS.get('hotspots', {}); (all[k] = all[k] || []).push({ ...r, to: keyOf(c, s) }); LS.set('hotspots', all);
        this.render(); toast('熱區已新增');
      });
      Sheet.open('這個框要跳到哪一頁？', pick);
    },
  };
  function toggleEditHotspots(on) {
    app.classList.toggle('edit-hs', on);
    toast(on ? '編輯熱區：在投影片上拖曳框出範圍，再選要跳到哪一頁。按 × 刪除。' : '已結束編輯熱區', on ? 4200 : 1500);
  }

  /* ---------- 講稿（密碼解密） ---------- */
  const Notes = {
    plain: null, open: false,
    async unlock(pw) {
      const n = H.notes, d = (s) => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));
      const base = await crypto.subtle.importKey('raw', new TextEncoder().encode(pw), 'PBKDF2', false, ['deriveKey']);
      const key = await crypto.subtle.deriveKey({ name: 'PBKDF2', salt: d(n.salt), iterations: n.iter, hash: 'SHA-256' }, base, { name: 'AES-GCM', length: 256 }, false, ['decrypt']);
      const pt = await crypto.subtle.decrypt({ name: 'AES-GCM', iv: d(n.iv) }, key, d(n.data));
      this.plain = JSON.parse(new TextDecoder().decode(pt));
    },
    update() {
      if (!this.open || !this.plain) return;
      const k = keyOf(cur.c, cur.s);
      $('npTitle').textContent = `講稿｜${CH[cur.c].short} P${cur.s + 1}`;
      $('npBody').textContent = this.plain[k] || '（這一頁沒有講稿）';
      $('npBody').scrollTop = 0;
    },
    show() { this.open = true; $('notesPanel').classList.remove('hidden'); $('btnNotes').classList.add('on'); this.update(); },
    hide() { this.open = false; $('notesPanel').classList.add('hidden'); $('btnNotes').classList.remove('on'); },
  };
  async function tryRemembered() {
    const pw = LS.get('npw', null) || (sessionStorage && sessionStorage.getItem('npw'));
    if (!pw) return false;
    try { await Notes.unlock(pw); return true; } catch (e) { LS.del('npw'); return false; }
  }
  $('btnNotes').onclick = async () => {
    if (Notes.open) { Notes.hide(); return; }
    if (Notes.plain || await tryRemembered()) { Notes.show(); return; }
    const b = Sheet.open('講者模式', `<p class="note" style="font-size:15px">輸入密碼後才會顯示講稿。</p><form id="pwForm"><input class="field" id="pwIn" type="password" autocomplete="current-password" placeholder="密碼"><label class="row note" style="margin-top:10px"><input type="checkbox" id="pwRemember"> 在這台裝置記住密碼</label><div class="msg" id="pwMsg"></div><button class="btn" type="submit">解鎖講稿</button></form>`);
    const inp = b.querySelector('#pwIn'); setTimeout(() => inp.focus(), 50);
    b.querySelector('#pwForm').onsubmit = async (e) => {
      e.preventDefault(); b.querySelector('#pwMsg').textContent = '解鎖中…';
      try {
        await Notes.unlock(inp.value);
        try { sessionStorage.setItem('npw', inp.value); } catch (err) { }
        if (b.querySelector('#pwRemember').checked) LS.set('npw', inp.value);
        Sheet.close(); Notes.show();
      } catch (err) { b.querySelector('#pwMsg').textContent = '密碼不正確，請再試一次。'; }
    };
  };
  $('npClose').onclick = () => Notes.hide();
  let npSize = LS.get('npSize', 18);
  const setNp = () => { $('npBody').style.setProperty('--np', npSize + 'px'); LS.set('npSize', npSize); };
  setNp();
  $('npSmaller').onclick = () => { npSize = Math.max(14, npSize - 2); setNp(); };
  $('npBigger').onclick = () => { npSize = Math.min(30, npSize + 2); setNp(); };

  /* ---------- 商品庫 ---------- */
  const Catalog = {
    data: null, group: 'all', mineOnly: true, cur: 'all',
    GROUPS: [
      { k: 'all', t: '全部' },
      { k: 'accident', t: '意外', cats: ['意外傷害'] },
      { k: 'hosp', t: '住院・手術・特定處置', cats: ['住院手術'] },
      { k: 'medexp', t: '實支實付', cats: ['實支實付'] },
      { k: 'ci', t: '癌症・重疾', cats: ['重大疾病/傷病'] },
      { k: 'ltc', t: '長照・失能', cats: ['長期照顧'] },
      { k: 'life', t: '壽險', cats: ['壽險'], fx: 1 },
      { k: 'invest', t: '投資・還本年金', cats: ['投資型・變額年金', '投資型・變額壽險', '投資型・變額萬能壽險', '投資型・連結型', '還本/年金', '還本/年金・利變年金', '還本/年金・遞延年金'], fx: 1 },
      { k: 'other', t: '旅平・OIU', cats: ['旅行平安', '國際保險(OIU)'], fx: 1 },
    ],
    async open(group) {
      if (group) this.group = group;
      if (!this.data) {
        try { this.data = await (await fetch('data/catalog.json')).json(); }
        catch (e) { Sheet.open('商品庫', '<p class="note">商品資料載入失敗，請連上網路後再試一次。</p>'); return; }
      }
      const b = Sheet.open('國泰商品庫', '<div id="catUi"></div>');
      this.render(b.querySelector('#catUi'));
    },
    render(el) {
      const sel = new Set(H.selection || []);
      const G = this.GROUPS, g = G.find((x) => x.k === this.group) || G[0];
      const showFx = g.k === 'all' || g.fx;
      let list = this.data.products.filter((p) => p.status !== '已下架' && !p.channel && !p.tele);
      if (this.mineOnly) list = list.filter((p) => sel.has(p.code));
      if (g.cats) list = list.filter((p) => g.cats.includes(p.cat));
      if (showFx && this.cur !== 'all') list = list.filter((p) => ((p.currency || '台幣') !== '台幣') === (this.cur === 'fx'));
      const chip = (on, attr, t) => `<button class="chip ${on ? 'on' : ''}" ${attr}>${esc(t)}</button>`;
      let h = `<div class="chips">${chip(this.mineOnly, 'data-mine="1"', '我的商品')}${chip(!this.mineOnly, 'data-mine="0"', '官網全部在售')}</div>`;
      h += `<div class="chips">${G.map((x) => chip(x.k === g.k, `data-g="${x.k}"`, x.t)).join('')}</div>`;
      if (showFx) h += `<div class="chips">${[['all', '全部幣別'], ['twd', '台幣'], ['fx', '外幣']].map(([k, t]) => chip(k === this.cur, `data-c="${k}"`, t)).join('')}</div>`;
      if (!list.length) h += `<p class="note">${this.mineOnly ? '這一類還沒有勾選商品。可切到「官網全部在售」查看，或到 Claude 商品庫勾選。' : '沒有符合條件的商品。'}</p>`;
      const by = {};
      list.forEach((p) => { const gg = G.find((x) => x.cats && x.cats.includes(p.cat)) || G[G.length - 1]; (by[gg.t] = by[gg.t] || []).push(p); });
      Object.entries(by).forEach(([t, arr]) => {
        h += `<div class="grp">${esc(t)}（${arr.length}）</div>`;
        arr.forEach((p) => {
          h += `<div class="prod"><div><span class="nm">${esc(p.name)}</span><span class="cd">${esc(p.code)}</span>${sel.has(p.code) ? '<span class="tg sel">我的商品</span>' : ''}${p.type ? `<span class="tg">${esc(p.type)}</span>` : ''}${p.currency && p.currency !== '台幣' ? `<span class="tg">${esc(p.currency)}</span>` : ''}</div>
          <div class="ds">${esc(p.desc)}</div>
          <div class="mt">${p.age ? '承保年齡 ' + esc(p.age) : ''}${p.limit ? '　保額 ' + esc(p.limit) : ''}</div>
          ${Doc.has(p.code) ? `<div class="docbtns">${Doc.has(p.code, 'dm') ? `<button class="btn sm" data-doc="${esc(p.code)}" data-kind="dm">看 DM</button>` : ''}${Doc.has(p.code, 'terms') ? `<button class="btn sm" data-doc="${esc(p.code)}" data-kind="sum">條款重點</button><button class="btn sm ghost" data-doc="${esc(p.code)}" data-kind="terms">條款全文</button>` : ''}</div>` : ''}
          <div class="ln">${p.dm ? `<a href="${esc(p.dm)}" target="_blank" rel="noopener">官網 DM</a>` : ''}${p.terms ? `<a href="${esc(p.terms)}" target="_blank" rel="noopener">官網條款</a>` : ''}<a href="${esc(p.page)}" target="_blank" rel="noopener">官網頁面</a></div></div>`;
        });
      });
      h += `<p class="note" style="margin-top:14px">資料來源：國泰人壽官網，最後同步 ${esc(this.data.updated)}。保障內容以保單條款為準。</p>`;
      el.innerHTML = h;
      el.onclick = (e) => {
        const b = e.target.closest('button'); if (!b) return;
        if (b.dataset.doc) { Sheet.close(); Doc.open(b.dataset.doc, b.dataset.kind); return; }
        if (b.dataset.mine) this.mineOnly = b.dataset.mine === '1';
        if (b.dataset.g) { this.group = b.dataset.g; this.cur = 'all'; }
        if (b.dataset.c) this.cur = b.dataset.c;
        this.render(el);
      };
    },
  };
  $('btnCatalog').onclick = () => Catalog.open();


  /* ---------- DM／條款閱讀器（可畫） ---------- */
  const Doc = {
    data: (window.HANDBOOK && H.docs) || { files: {}, summaries: {} },
    code: null, kind: null, pages: [], cur: null, zoom: LS.get('docZoom', 1),
    has(code, kind) { const f = this.data.files[code]; return !!(f && (!kind || f[kind])); },
    isOpen() { return !$('doc').classList.contains('hidden'); },
    active() { return this.cur; },
    name(code) { const s = this.data.summaries[code]; return (s && s.name) || code; },
    open(code, kind, page) {
      this.code = code; this.kind = kind;
      $('doc').classList.remove('hidden');
      $('docName').textContent = this.name(code).replace(/^國泰人壽/, '');
      const hasTerms = this.has(code, 'terms'), hasDm = this.has(code, 'dm');
      $('docTabs').innerHTML = [hasDm && ['dm', 'DM'], hasTerms && ['sum', '條款重點'], hasTerms && ['terms', '條款全文']].filter(Boolean)
        .map(([k, t]) => `<button class="chip ${k === kind ? 'on' : ''}" data-k="${k}">${t}</button>`).join('');
      if (kind === 'sum') this.renderSummary(); else this.renderPages(page);
    },
    close() { $('doc').classList.add('hidden'); $('docBody').innerHTML = ''; this.pages = []; this.cur = null; if (this.io) this.io.disconnect(); },
    renderSummary() {
      const s = this.data.summaries[this.code], body = $('docBody');
      $('docZoomBox').classList.add('hidden'); $('docSave').classList.add('hidden'); $('docPage').textContent = '';
      this.pages = []; this.cur = null;
      let h = `<div class="sum"><p class="sum-lead">${esc(s.summary)}</p>`;
      s.sections.forEach((sec) => {
        h += `<h3>${esc(sec.h)}</h3><ul>` + sec.items.map((it) => `<li><span>${esc(it.t)}</span><button class="pg" data-p="${it.p}">條款第 ${it.p} 頁 →</button></li>`).join('') + '</ul>';
      });
      h += '<p class="note">本頁為條款重點整理，實際保障內容以保單條款原文為準。點右側頁碼可直接看原文。</p></div>';
      body.innerHTML = h; body.scrollTop = 0;
    },
    renderPages(page) {
      const f = this.data.files[this.code][this.kind], body = $('docBody');
      $('docZoomBox').classList.remove('hidden'); $('docSave').classList.remove('hidden');
      body.innerHTML = ''; this.pages = []; this.cur = null;
      const wrap = document.createElement('div'); wrap.className = 'pages'; wrap.style.setProperty('--z', this.zoom);
      for (let n = 1; n <= f.pages; n++) {
        const pg = document.createElement('div'); pg.className = 'pg-wrap'; pg.dataset.n = n;
        const [pw, ph] = (f.sizes && f.sizes[n - 1]) || [f.w, f.h];
        pg.style.aspectRatio = `${pw} / ${ph}`; if (pw > ph) pg.classList.add('wide');
        pg.innerHTML = `<img alt="第 ${n} 頁" loading="${n <= 2 ? 'eager' : 'lazy'}" decoding="async" src="docs/${this.code}/${this.kind}/${String(n).padStart(2, '0')}.webp"><span class="pg-no">${n} / ${f.pages}</span>`;
        wrap.appendChild(pg);
        this.pages.push({ n, el: pg, surface: null, w: pw, h: ph });
      }
      body.appendChild(wrap);
      if (this.io) this.io.disconnect();
      this.io = new IntersectionObserver((ents) => ents.forEach((en) => {
        const P = this.pages[+en.target.dataset.n - 1];
        if (en.isIntersecting) this.mount(P); else this.unmount(P);
      }), { root: body, rootMargin: '600px 0px' });
      this.pages.forEach((P) => this.io.observe(P.el));
      body.scrollTop = 0;
      if (page) requestAnimationFrame(() => { const P = this.pages[page - 1]; if (P) body.scrollTop = P.el.offsetTop - 8; });
      this.updatePageNo();
    },
    key(P) { return `doc:${this.code}:${this.kind}:${P.n}`; },
    mount(P) {
      if (P.surface) return;
      const c = document.createElement('canvas'); c.className = 'pg-ink'; P.el.appendChild(c);
      const self = this;
      P.surface = makeSurface(c, { width: () => P.w, height: () => P.h, resolution: () => Math.min(2, (P.el.clientWidth / P.w) * (window.devicePixelRatio || 1)), unit: () => P.w / 1500 });
      P.surface.resize(); P.surface.load(self.key(P)); P.canvas = c;
    },
    unmount(P) { if (!P.surface) return; P.canvas.remove(); P.surface = null; P.canvas = null; },
    pageAt(e) { const el = e.target.closest('.pg-wrap'); return el ? this.pages[+el.dataset.n - 1] : null; },
    updatePageNo() {
      const body = $('docBody'), mid = body.scrollTop + body.clientHeight / 3;
      const P = this.pages.find((x) => x.el.offsetTop + x.el.offsetHeight > mid) || this.pages[0];
      if (P) { $('docPage').textContent = `第 ${P.n} / ${this.pages.length} 頁`; this.visible = P; if (!this.cur || !this.cur.surface) this.cur = P.surface; }
    },
    setZoom(z) {
      this.zoom = Math.max(1, Math.min(2.5, z)); LS.set('docZoom', this.zoom);
      const body = $('docBody'), r = body.scrollTop / (body.scrollHeight || 1);
      body.querySelector('.pages').style.setProperty('--z', this.zoom);
      requestAnimationFrame(() => { body.scrollTop = r * body.scrollHeight; this.pages.forEach((P) => P.surface && P.surface.resize()); });
    },
    saveImage() {
      const P = this.visible; if (!P) return;
      const img = P.el.querySelector('img'), out = document.createElement('canvas'); out.width = P.w; out.height = P.h;
      const ctx = out.getContext('2d'); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, P.w, P.h);
      ctx.drawImage(img, 0, 0, P.w, P.h); if (P.canvas) ctx.drawImage(P.canvas, 0, 0, P.w, P.h);
      showImage(out, `${this.name(this.code).replace(/^國泰人壽/, '')}_${this.kind === 'dm' ? 'DM' : '條款'}_第${P.n}頁`);
    },
  };
  $('docTabs').onclick = (e) => { const b = e.target.closest('button'); if (b) Doc.open(Doc.code, b.dataset.k); };
  $('docClose').onclick = () => Doc.close();
  $('docZoomIn').onclick = () => Doc.setZoom(Doc.zoom + 0.25);
  $('docZoomOut').onclick = () => Doc.setZoom(Doc.zoom - 0.25);
  $('docSave').onclick = () => Doc.saveImage();
  $('docBody').addEventListener('scroll', () => Doc.updatePageNo(), { passive: true });
  $('docBody').addEventListener('click', (e) => { const b = e.target.closest('.pg'); if (b) Doc.open(Doc.code, 'terms', +b.dataset.p); });
  // 畫筆：Apple Pencil 在頁面上畫；手指捲動（開啟「手指也能畫」時改成畫）
  let dg = null;
  const docBody = $('docBody');
  const pagePt = (P, e) => { const r = P.el.getBoundingClientRect(); return { x: (e.clientX - r.left) * P.w / r.width, y: (e.clientY - r.top) * P.h / r.height }; };
  docBody.addEventListener('pointerdown', (e) => {
    if (Doc.kind === 'sum') return;
    const P = Doc.pageAt(e); if (!P || !P.surface) return;
    if (!(e.pointerType === 'pen' || Tool.finger || e.pointerType === 'mouse')) return;
    dg = { id: e.pointerId, P }; Doc.cur = P.surface;
    try { docBody.setPointerCapture(e.pointerId); } catch (err) { }
    const p = pagePt(P, e); P.surface.begin(p.x, p.y); e.preventDefault();
  });
  docBody.addEventListener('pointermove', (e) => {
    if (!dg || e.pointerId !== dg.id) return;
    (e.getCoalescedEvents ? e.getCoalescedEvents() : [e]).forEach((ev) => { const p = pagePt(dg.P, ev); dg.P.surface && dg.P.surface.move(p.x, p.y); });
    e.preventDefault();
  });
  const dEnd = (e, cancel) => { if (!dg || e.pointerId !== dg.id) return; const S = dg.P.surface; dg = null; if (S) (cancel ? S.cancel() : S.end()); };
  docBody.addEventListener('pointerup', (e) => dEnd(e, false));
  docBody.addEventListener('pointercancel', (e) => dEnd(e, true));
  docBody.addEventListener('touchmove', (e) => { if (dg || Tool.finger) e.preventDefault(); }, { passive: false });
  window.addEventListener('resize', () => { if (Doc.isOpen()) Doc.pages.forEach((P) => P.surface && P.surface.resize()); });

  /* ---------- 更多 ---------- */
  $('btnMore').onclick = () => {
    const editing = app.classList.contains('edit-hs');
    const b = Sheet.open('更多', `<div style="display:flex;flex-direction:column;gap:10px">
      <button class="btn" id="mSave">把這一頁（含畫記）存成圖片</button>
      <button class="btn ghost" id="mHs">${editing ? '結束編輯熱區' : '編輯熱區（自由框選、點了跳頁）'}</button>
      <button class="btn ghost" id="mHsExport">匯出我的熱區（交給 Claude 寫進手冊）</button>
      <button class="btn ghost" id="mClearAll">清除所有頁面的畫記（換下一位客戶前）</button>
      <button class="btn ghost" id="mLock">鎖上講稿（忘記這台裝置的密碼）</button>
      <button class="btn ghost" id="mUpdate">檢查更新</button>
      <p class="note">版本 ${esc(H.version)}。點投影片上的「P 數字」或卡片可直接跳頁，看完按左下「返回」。Apple Pencil 直接畫，手指左右滑動換頁。</p></div>`);
    b.querySelector('#mSave').onclick = () => { Sheet.close(); saveSlideImage(); };
    b.querySelector('#mHs').onclick = () => { Sheet.close(); toggleEditHotspots(!editing); Hotspots.render(); };
    b.querySelector('#mHsExport').onclick = async () => {
      const txt = JSON.stringify(LS.get('hotspots', {}));
      try { await navigator.clipboard.writeText(txt); toast('已複製熱區資料，貼給 Claude 即可'); }
      catch (e) { Sheet.open('我的熱區', `<textarea class="field" rows="8" readonly>${esc(txt)}</textarea><p class="note">全選後複製，貼給 Claude。</p>`); }
    };
    b.querySelector('#mClearAll').onclick = () => confirmBox('清除所有畫記', '所有頁面和白板的筆畫都會清除，無法復原。', '全部清除', () => {
      try { Object.keys(localStorage).filter((k) => k.startsWith('ink:') || k.startsWith('ink:doc:')).forEach((k) => localStorage.removeItem(k)); } catch (e) { }
      Ink.load(keyOf(cur.c, cur.s)); toast('已清除所有畫記');
    }, true);
    b.querySelector('#mLock').onclick = () => { LS.del('npw'); try { sessionStorage.removeItem('npw'); } catch (e) { } Notes.plain = null; Notes.hide(); Sheet.close(); toast('講稿已鎖上'); };
    b.querySelector('#mUpdate').onclick = async () => {
      Sheet.close();
      if (!('serviceWorker' in navigator)) { location.reload(); return; }
      const reg = await navigator.serviceWorker.getRegistration();
      if (reg) { await reg.update(); toast('已檢查更新，有新版本會自動套用'); setTimeout(() => location.reload(), 1500); }
    };
  };

  /* ---------- 啟動 ---------- */
  function start() {
    fit();
    const h = decodeURIComponent(location.hash.slice(1));
    if (h.startsWith('catalog/')) { show(0, 0, { noAnim: true }); Catalog.open(h.slice(8)); }
    else { const p = locate(h); p ? show(p.c, p.s, { noAnim: true }) : show(0, 0, { noAnim: true }); }
    if (document.fonts && document.fonts.ready) document.fonts.ready.then(() => Ink.redraw());
  }
  start();
  window.addEventListener('hashchange', () => { const h = decodeURIComponent(location.hash.slice(1)); if (h.startsWith('catalog/')) { Catalog.open(h.slice(8)); return; } const p = locate(h); if (p && (p.c !== cur.c || p.s !== cur.s)) show(p.c, p.s); });
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('sw.js').then((reg) => {
      reg.addEventListener('updatefound', () => {
        const w = reg.installing; if (!w) return;
        w.addEventListener('statechange', () => { if (w.state === 'installed' && navigator.serviceWorker.controller) toast('手冊已更新，重新開啟即可看到新版', 4000); });
      });
    }).catch(() => { });
  }
})();
