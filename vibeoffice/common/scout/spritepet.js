/* VibeOffice — a drawn assistant played from a Codex pet sheet (v1: 8×9 cells, 192×208).
 * Nikki's sheet is nikki.webp, beside this file. assistant.js builds her instead of Scout when she is chosen
 * (Help ▸ Select Assistant). The calls are Scout's, and the action name stays the one assistant.js asked for;
 * the clip underneath is the nearest row: idle, running-right, running-left, waving, jumping, failed, waiting,
 * running, review. The balloon is Scout's (scout.css).
 */
(function (root) {
  'use strict';
  const L = root.L || (root.L = {});
  if (typeof document === 'undefined') return;

  const COLS = 8, CELL_W = 192, CELL_H = 208, FRAME_MS = 375, FPS = 1000 / FRAME_MS;
  const SHEET = new URL('nikki.webp', document.currentScript.src).href;
  const ROW = { idle: 0, 'running-right': 1, 'running-left': 2, waving: 3, jumping: 4, failed: 5, waiting: 6, running: 7, review: 8 };
  /* what assistant.js plays, and the row that stands in for it */
  const CLIP = {
    Show: 'jumping', Greeting: 'waving', Hide: 'jumping', GoodBye: 'waving',
    Print: 'running', Fetch: 'running-right', Searching: 'waiting', CheckingSomething: 'review',
    SendMail: 'running-right', Writing: 'review', EmptyTrash: 'failed', GetArtsy: 'review',
    GetTechy: 'waiting', GetWizardy: 'jumping', Processing: 'running', Save: 'running', Alert: 'failed',
    GestureUp: 'waving', GestureDown: 'waving', GestureLeft: 'running-left', GestureRight: 'running-right',
    Idle1_1: 'idle',
  };
  const HIDES = { Hide: true, GoodBye: true };
  const clamp = (v, a, b) => (v < a ? a : v > b ? b : v);
  const reduce = () => matchMedia('(prefers-reduced-motion: reduce)').matches;

  let counts = null;
  const sheet = new Image();
  sheet.onload = () => { counts = framesOf(sheet); if (shown) shown._paint(); };
  sheet.src = SHEET;
  let shown = null;

  /** How many leading cells of each row are drawn; a clear cell ends the row. */
  function framesOf(img) {
    const c = document.createElement('canvas');
    c.width = CELL_W; c.height = CELL_H;
    const g = c.getContext('2d', { willReadFrequently: true });
    const out = [];
    for (let r = 0; r < 9; r++) {
      let n = 0;
      for (let col = 0; col < COLS; col++) {
        g.clearRect(0, 0, CELL_W, CELL_H);
        g.drawImage(img, col * CELL_W, r * CELL_H, CELL_W, CELL_H, 0, 0, CELL_W, CELL_H);
        const d = g.getImageData(0, 0, CELL_W, CELL_H).data;
        let ink = false;
        for (let i = 3; i < d.length; i += 16) if (d[i] > 12) { ink = true; break; }
        if (!ink) break;
        n++;
      }
      out.push(n || 1);
    }
    return out;
  }

  class SpritePet {
    constructor(opts = {}) {
      this.o = Object.assign({ size: 192, h: Math.round(192 * CELL_H / CELL_W), idle: true, label: 'Nikki, the assistant' }, opts);
      this._q = []; this._running = false; this._cur = null; this._nx = null;
      this._hidden = true; this._frozen = false; this._timer = 0; this._quiet = false; this._replacing = false;
      this._pos = [0, 0]; this._frame = -1; this._row = 0;
      const el = this.el = document.createElement('div');
      el.className = 'scout-agent is-hidden';
      el.style.setProperty('--scout-w', this.o.size + 'px');
      el.style.height = this.o.h + 'px';
      el.tabIndex = 0;
      el.setAttribute('role', 'img');
      el.setAttribute('aria-label', this.o.label);
      const canvas = this.canvas = document.createElement('canvas');
      canvas.width = CELL_W; canvas.height = CELL_H;
      canvas.setAttribute('aria-hidden', 'true');
      el.appendChild(canvas);
      this.g = canvas.getContext('2d');
      const b = this.balloon = document.createElement('div');
      b.className = 'scout-balloon';
      b.setAttribute('role', 'status');
      b.innerHTML = '<div class="scout-text"></div><div class="scout-choices" role="group" hidden></div>';
      this._bText = b.firstChild; this._bChoices = b.lastChild;
      b.addEventListener('click', (e) => {
        const bt = e.target.closest('.scout-choices button');
        if (bt) { if (this._bPick) this._bPick(+bt.dataset.i); return; }
        if (!this._bPick) this._closeBalloon();
      });
      b.addEventListener('keydown', (e) => { if (e.key === 'Escape') this._closeBalloon(); });
      el.appendChild(b);
      (this.o.parent || document.body).appendChild(el);
      this._bind();
      shown = this;
      if (counts) this._paint();
    }

    play(name, opts = {}) {
      if (opts.interrupt) { this.stop(true); return this._playNow(name, opts); }
      if (opts.next) return this._playNext(name, opts);
      return this._enqueue(() => this._playNow(name, opts));
    }
    speak(text, opts = {}) { return this._enqueue(() => this._speakNow(String(text), opts)); }
    ask(text, choices, opts = {}) { return this.speak(text, Object.assign({}, opts, { choices })); }
    delay(ms) { return this._enqueue(() => new Promise((r) => setTimeout(() => r(true), ms))); }
    gestureAt(x) {
      const r = this.el.getBoundingClientRect();
      return this.play(x >= r.left + r.width / 2 ? 'GestureRight' : 'GestureLeft');
    }
    animate() {
      const pool = ['Greeting', 'GetArtsy', 'GetWizardy', 'Alert', 'Searching', 'Print'];
      return this.play(pool[Math.floor(Math.random() * pool.length)]);
    }
    /** quiet: do not drop back into the idle clip (a play with interrupt stops, then starts its own). */
    stop(quiet) {
      this._quiet = true;
      while (this._q.length) this._q.shift().res(false);
      if (this._nx) { this._nx.res(false); this._nx = null; }
      if (this._cur) this._end(false);
      this._closeBalloon();
      this._quiet = false;
      if (!quiet && !this._hidden && !this._frozen) this._playNow('Idle1_1', { idle: true, loop: true });
      return this;
    }
    hide(fast) {
      return this._enqueue(() => {
        if (this._hidden) return true;
        if (fast) {
          this._quiet = true;
          if (this._cur) this._end(false);
          this._quiet = false;
          this._conceal();
          return true;
        }
        return this._playNow('Hide');
      });
    }
    userActive() { return this; }
    freeze(on) {
      this._frozen = !!on;
      if (this._frozen) this._halt();
      else if (this._cur) this._arm();
      return this;
    }
    get frozen() { return !!this._frozen; }
    get hidden() { return this._hidden; }
    get current() {
      const c = this._cur;
      return c ? { name: c.name, t: c.t, dur: c.dur, loop: !!c.loop, idle: !!c.idle } : null;
    }
    closeBalloon() { this._closeBalloon(); return this; }
    destroy() {
      this.stop(true);
      this._halt();
      if (shown === this) shown = null;
      this.el.remove();
    }
    placeAt(x, y) {
      const W = this.o.size, H = this.o.h;
      this._pos = [clamp(x, -W * .2, innerWidth - W * .8), clamp(y, -H * .2, innerHeight - H * .7)];
      this.el.style.transform = 'translate(' + Math.round(this._pos[0]) + 'px,' + Math.round(this._pos[1]) + 'px)';
      return this;
    }

    _enqueue(fn) {
      return new Promise((res) => { this._q.push({ fn, res }); this._pump(); });
    }
    _pump() {
      if (this._running || !this._q.length) return;
      this._running = true;
      const job = this._q.shift();
      Promise.resolve().then(job.fn).then((v) => job.res(v), () => job.res(false)).then(() => {
        this._running = false;
        this._pump();
      });
    }
    _playNext(name, opts) {
      if (this._nx) { this._nx.res(false); this._nx = null; }
      if (!this._cur || this._cur.idle) return this._playNow(name, opts);
      return new Promise((res) => { this._nx = { name, opts, res }; });
    }
    _playNow(name, opts) {
      this._replacing = true;
      if (this._cur) this._end(false);
      this._replacing = false;
      const clip = CLIP[name] || 'idle';
      const row = ROW[clip];
      const frames = (counts && counts[row]) || 1;
      const cycle = frames / FPS;
      const loop = !!(opts.loop || opts.until || opts.idle);
      if (this._hidden && name !== 'Hide') this._reveal();
      return new Promise((res) => {
        const c = this._cur = {
          name, row, frames, cycle, loop, idle: !!opts.idle, hides: !!HIDES[name],
          dur: loop ? cycle : Math.max(cycle, 6 * FRAME_MS / 1000), t: 0, res,
        };
        if (opts.until && typeof opts.until.then === 'function') {
          c.loop = true;
          opts.until.then(() => { if (this._cur === c) { c.loop = false; if (reduce()) this._end(true); } }, () => { if (this._cur === c) c.loop = false; });
        }
        this._frame = -1;
        this._paint();
        if (reduce() && !c.loop) this._timer = setTimeout(() => { this._timer = 0; if (this._cur === c) this._end(true); }, c.dur * 1000);
        else this._arm();
      });
    }
    _end(done) {
      const c = this._cur;
      if (!c) return;
      this._cur = null;
      this._halt();
      if (done && c.hides) this._conceal();
      c.res(done);
      if (this._replacing) return;
      if (this._nx) { const nx = this._nx; this._nx = null; this._playNow(nx.name, nx.opts).then(nx.res); return; }
      if (!this._quiet && !this._hidden && !this._frozen && this.o.idle !== false) this._playNow('Idle1_1', { idle: true, loop: true });
    }
    _arm() {
      if (this._timer || this._frozen || !this._cur || reduce()) return;
      this._last = performance.now();
      const step = () => {
        this._timer = 0;
        const c = this._cur;
        if (!c || this._frozen) return;
        const now = performance.now();
        c.t += Math.min(1, (now - this._last) / 1000);
        this._last = now;
        if (c.loop) { if (c.t >= c.cycle) c.t %= c.cycle; }
        else if (c.t >= c.dur) { this._end(true); return; }
        this._paint();
        this._timer = setTimeout(step, 1000 / FPS);
      };
      this._timer = setTimeout(step, 1000 / FPS);
    }
    _halt() { if (this._timer) clearTimeout(this._timer); this._timer = 0; }
    _paint() {
      const c = this._cur;
      if (!c || !sheet.complete || !sheet.naturalWidth) return;
      const frame = c.frames < 2 ? 0 : Math.floor(c.t / c.cycle * c.frames) % c.frames;
      if (frame === this._frame && c.row === this._row) return;
      this._frame = frame; this._row = c.row;
      this.g.clearRect(0, 0, CELL_W, CELL_H);
      this.g.drawImage(sheet, frame * CELL_W, c.row * CELL_H, CELL_W, CELL_H, 0, 0, CELL_W, CELL_H);
    }
    _reveal() { this._hidden = false; this.el.classList.remove('is-hidden'); }
    _conceal() { this._hidden = true; this.el.classList.add('is-hidden'); this._halt(); this._closeBalloon(); }

    _speakNow(text, opts) {
      if (this._hidden) this._reveal();
      const b = this.balloon, tx = this._bText, ch = this._bChoices;
      clearTimeout(this._bTimer); clearInterval(this._bType);
      const choices = (opts.choices || []).map((c) => typeof c === 'string' ? { label: c, value: c } : { label: String(c.label), value: c.value !== undefined ? c.value : c.label });
      ch.innerHTML = '';
      choices.forEach((c, i) => { const bt = document.createElement('button'); bt.type = 'button'; bt.textContent = c.label; bt.dataset.i = String(i); ch.appendChild(bt); });
      b.classList.toggle('has-choices', choices.length > 0);
      this._placeBalloon(text, choices.length > 0);
      tx.textContent = ''; ch.hidden = true;
      b.classList.add('is-on');
      const words = text.split(/(\s+)/);
      let i = 0;
      return new Promise((res) => {
        this._bDone = (val) => {
          clearTimeout(this._bTimer); clearInterval(this._bType);
          this._bDone = null; this._bPick = null;
          b.classList.remove('is-on'); res(val);
        };
        this._bPick = choices.length ? ((n) => { if (this._bDone) this._bDone(choices[n].value); }) : null;
        const finish = () => {
          clearInterval(this._bType); tx.textContent = text;
          if (choices.length) { ch.hidden = false; return; }
          if (opts.hold) return;
          this._bTimer = setTimeout(() => this._bDone && this._bDone(true), opts.timeout || 1800 + text.length * 38);
        };
        if (reduce()) { finish(); return; }
        this._bType = setInterval(() => { tx.textContent += words[i++] || ''; if (i >= words.length) finish(); }, 70);
      });
    }
    _closeBalloon() { if (this._bDone) this._bDone(this._bPick ? null : false); }
    _placeBalloon(text, withChoices) {
      const b = this.balloon, W = this.el.offsetWidth || this.o.size, H = this.el.offsetHeight || this.o.h;
      this._bText.textContent = text; this._bChoices.hidden = !withChoices;
      b.style.width = ''; b.style.left = '0px'; b.style.top = '0px'; b.style.bottom = 'auto'; b.classList.remove('is-below');
      const bw = Math.ceil(b.getBoundingClientRect().width), bh = b.offsetHeight;
      b.style.width = bw + 'px';
      let left = W * 0.5 - bw * 0.5;
      if (this._pos[0] + left < 8) left += 8 - (this._pos[0] + left);
      if (this._pos[0] + left + bw > innerWidth - 8) left -= this._pos[0] + left + bw - (innerWidth - 8);
      const above = this._pos[1] + 4 - bh >= 8;
      b.style.left = left + 'px';
      if (above) { b.style.top = 'auto'; b.style.bottom = (H + 10) + 'px'; }
      else { b.style.top = (H - 4) + 'px'; b.classList.add('is-below'); }
      b.style.setProperty('--scout-tip', clamp(W * 0.5 - left, 16, Math.max(16, bw - 16)) + 'px');
    }
    _bind() {
      const el = this.el, o = this.o;
      let start = null;
      el.addEventListener('pointerdown', (e) => {
        if (e.button !== 0 || this._hidden || e.target.closest('.scout-balloon')) return;
        start = { x: e.clientX, y: e.clientY, pos: this._pos.slice(), moved: false };
        el.setPointerCapture(e.pointerId);
      });
      el.addEventListener('mousedown', (e) => { e.preventDefault(); });
      el.addEventListener('pointermove', (e) => {
        if (!start) return;
        const dx = e.clientX - start.x, dy = e.clientY - start.y;
        if (!start.moved && Math.hypot(dx, dy) < 4) return;
        start.moved = true; el.classList.add('is-dragging');
        this.placeAt(start.pos[0] + dx, start.pos[1] + dy);
      });
      const end = (e) => {
        if (!start) return;
        const s = start; start = null;
        el.classList.remove('is-dragging');
        if (e.type !== 'pointerup') return;
        if (s.moved) { if (o.onmoved) o.onmoved(this._pos[0], this._pos[1]); }
        else if (o.onclick) o.onclick();
      };
      el.addEventListener('pointerup', end);
      el.addEventListener('pointercancel', end);
      el.addEventListener('keydown', (e) => {
        if (e.target.closest('.scout-balloon')) return;
        if ((e.key === 'Enter' || e.key === ' ') && o.onclick) { e.preventDefault(); o.onclick(); }
        else if ((e.key === 'ContextMenu' || (e.key === 'F10' && e.shiftKey)) && o.onmenu) {
          e.preventDefault();
          const r = this.canvas.getBoundingClientRect();
          o.onmenu(r.left + r.width / 2, r.top + r.height / 2);
        } else if (e.key === 'Escape') this.stop();
      });
      el.addEventListener('contextmenu', (e) => {
        e.preventDefault();
        if (o.onmenu && !e.target.closest('.scout-balloon')) o.onmenu(e.clientX, e.clientY);
      });
    }
  }

  /** What assistant.js asks for. null means Scout, and only an explicit stored id of scout asks for him.
   *  With nothing stored, Nikki is the assistant. */
  function current() {
    let id = 'nikki';
    try { id = (JSON.parse(localStorage.getItem('vibeoffice.assistant')) || {}).id || 'nikki'; } catch (e) { /* Nikki */ }
    if (id === 'scout') return null;
    const size = 192; // 128 × 1.5: half again as tall, the drawing's own proportions kept
    return {
      id: 'nikki', name: 'Nikki', ctor: SpritePet, size, h: Math.round(size * CELL_H / CELL_W),
      label: 'Nikki, the assistant. Click her for help; right-click for more.',
      line: "I'm Nikki. I'll keep you company while you work.",
      tip: 'Nikki, the assistant',
    };
  }

  L.SpritePet = SpritePet;
  L.assistantPet = { current };
})(typeof window !== 'undefined' ? window : globalThis);
