/* Quire DOM Console. Help opens a modeless window; the script runs in this page.
   The object model is js/dom.js. xterm is public/common/vendor/xterm/. */
(function (window) {
  'use strict';
  const L = window.L = window.L || {};

  function identChar(c) { return !!c && /[A-Za-z0-9_$]/.test(c); }
  function skipWs(src, i) {
    while (i < src.length && ' \t\n\r\f'.includes(src[i])) i++;
    return i;
  }
  function readIdent(src, i) {
    if (!src[i] || !/[A-Za-z_$]/.test(src[i])) return '';
    let j = i + 1;
    while (j < src.length && identChar(src[j])) j++;
    return src.slice(i, j);
  }

  /** a named function or async function at i, or null */
  function matchFunction(src, i) {
    let j = i;
    if (src.startsWith('async', j) && !identChar(src[j + 5])) {
      const k = skipWs(src, j + 5);
      if (src.startsWith('function', k) && !identChar(src[k + 8])) j = k;
      else return null;
    }
    if (!src.startsWith('function', j) || identChar(src[j + 8])) return null;
    j = skipWs(src, j + 8);
    const name = readIdent(src, j);
    if (!name) return null;
    j = skipWs(src, j + name.length);
    if (src[j] !== '(') return null;
    return { name: name, paren: j };
  }

  /** index just after the function body that starts at the parameter "(" */
  function functionEnd(src, parenAt) {
    let i = parenAt, paren = 0, brace = 0, bracket = 0, state = 'code', body = false;
    const tmpl = [];
    while (i < src.length) {
      const c = src[i], d = src[i + 1];
      if (state === 'line') { if (c === '\n') state = 'code'; i++; continue; }
      if (state === 'block') { if (c === '*' && d === '/') { state = 'code'; i += 2; continue; } i++; continue; }
      if (state === 'sq' || state === 'dq') {
        if (c === '\\') { i += 2; continue; }
        if (c === (state === 'sq' ? "'" : '"')) state = 'code';
        i++; continue;
      }
      if (state === 'tmpl') {
        if (c === '\\') { i += 2; continue; }
        if (c === '`') { state = 'code'; i++; continue; }
        if (c === '$' && d === '{') { tmpl.push(brace); state = 'code'; brace++; i += 2; continue; }
        i++; continue;
      }
      if (c === '/' && d === '/') { state = 'line'; i += 2; continue; }
      if (c === '/' && d === '*') { state = 'block'; i += 2; continue; }
      if (c === "'") { state = 'sq'; i++; continue; }
      if (c === '"') { state = 'dq'; i++; continue; }
      if (c === '`') { state = 'tmpl'; i++; continue; }
      if (c === '(') paren++;
      else if (c === ')') paren--;
      else if (c === '[') bracket++;
      else if (c === ']') bracket--;
      else if (c === '{') { brace++; if (paren === 0 && !body) body = true; }
      else if (c === '}') {
        brace--;
        if (tmpl.length && brace === tmpl[tmpl.length - 1]) { tmpl.pop(); state = 'tmpl'; }
        else if (body && brace === 0 && paren === 0 && bracket === 0) return i + 1;
      }
      i++;
    }
    return -1;
  }

  /** true while a bracket, string, template or block comment is still open */
  function needsMore(src) {
    let i = 0, brace = 0, paren = 0, bracket = 0, state = 'code';
    const tmpl = [];
    while (i < src.length) {
      const c = src[i], d = src[i + 1];
      if (state === 'line') { if (c === '\n') state = 'code'; i++; continue; }
      if (state === 'block') { if (c === '*' && d === '/') { state = 'code'; i += 2; continue; } i++; continue; }
      if (state === 'sq' || state === 'dq') {
        if (c === '\\') { i += 2; continue; }
        if (c === (state === 'sq' ? "'" : '"')) state = 'code';
        i++; continue;
      }
      if (state === 'tmpl') {
        if (c === '\\') { i += 2; continue; }
        if (c === '`') { state = 'code'; i++; continue; }
        if (c === '$' && d === '{') { tmpl.push(brace); state = 'code'; brace++; i += 2; continue; }
        i++; continue;
      }
      if (c === '/' && d === '/') { state = 'line'; i += 2; continue; }
      if (c === '/' && d === '*') { state = 'block'; i += 2; continue; }
      if (c === "'") { state = 'sq'; i++; continue; }
      if (c === '"') { state = 'dq'; i++; continue; }
      if (c === '`') { state = 'tmpl'; i++; continue; }
      if (c === '(') paren++;
      else if (c === ')') paren--;
      else if (c === '[') bracket++;
      else if (c === ']') bracket--;
      else if (c === '{') brace++;
      else if (c === '}') {
        brace--;
        if (tmpl.length && brace === tmpl[tmpl.length - 1]) { tmpl.pop(); state = 'tmpl'; }
      }
      i++;
    }
    return (state !== 'code' && state !== 'line') || brace > 0 || paren > 0 || bracket > 0 || tmpl.length > 0;
  }

  /** top-level function and async function declarations become assignments on globalThis */
  function hoistFunctions(src) {
    const edits = [];
    let i = 0, brace = 0, paren = 0, bracket = 0, state = 'code', boundary = true;
    const tmpl = [];
    while (i < src.length) {
      const c = src[i], d = src[i + 1];
      if (state === 'line') { if (c === '\n') state = 'code'; i++; continue; }
      if (state === 'block') { if (c === '*' && d === '/') { state = 'code'; i += 2; continue; } i++; continue; }
      if (state === 'sq' || state === 'dq') {
        if (c === '\\') { i += 2; continue; }
        if (c === (state === 'sq' ? "'" : '"')) state = 'code';
        i++; continue;
      }
      if (state === 'tmpl') {
        if (c === '\\') { i += 2; continue; }
        if (c === '`') { state = 'code'; i++; continue; }
        if (c === '$' && d === '{') { tmpl.push(brace); state = 'code'; brace++; boundary = false; i += 2; continue; }
        i++; continue;
      }
      if (c === '/' && d === '/') { state = 'line'; i += 2; continue; }
      if (c === '/' && d === '*') { state = 'block'; i += 2; continue; }
      if (c === "'") { state = 'sq'; boundary = false; i++; continue; }
      if (c === '"') { state = 'dq'; boundary = false; i++; continue; }
      if (c === '`') { state = 'tmpl'; boundary = false; i++; continue; }
      if (c === ' ' || c === '\t' || c === '\n' || c === '\r' || c === '\f') { i++; continue; }
      if (boundary && brace === 0 && paren === 0 && bracket === 0) {
        const m = matchFunction(src, i);
        if (m) {
          const end = functionEnd(src, m.paren);
          if (end < 0) break;
          edits.push({ start: i, end: end, name: m.name });
          i = end;
          boundary = true;
          continue;
        }
      }
      if (c === '(') paren++;
      else if (c === ')') paren--;
      else if (c === '[') bracket++;
      else if (c === ']') bracket--;
      else if (c === '{') brace++;
      else if (c === '}') {
        brace--;
        if (tmpl.length && brace === tmpl[tmpl.length - 1]) { tmpl.pop(); state = 'tmpl'; }
      }
      boundary = c === ';' || c === '{' || c === '}';
      i++;
    }
    let out = src;
    for (let k = edits.length - 1; k >= 0; k--) {
      const e = edits[k];
      out = out.slice(0, e.start) + 'globalThis.' + e.name + ' = ' + out.slice(e.start, e.end) + out.slice(e.end);
    }
    return out;
  }

  function startsWithFunction(code) {
    const stripped = String(code).replace(/^\s*((\/\/[^\n]*\n)|(\/\*[\s\S]*?\*\/)\s*)*/, '');
    return /^(async\s+)?function\b/.test(stripped);
  }

  const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor;

  async function evaluate(code) {
    const prepared = hoistFunctions(code);
    const run = (body) => (new AsyncFunction(body)).call(globalThis);
    try {
      return await run('return (\n' + prepared + '\n);');
    } catch (e) {
      if (!(e instanceof SyntaxError)) throw e;
    }
    return await run(prepared);
  }

  function formatValue(v) {
    if (v === null) return 'null';
    if (v === undefined) return 'undefined';
    if (typeof v === 'string') return v;
    if (typeof v === 'number' || typeof v === 'boolean' || typeof v === 'bigint') return String(v);
    if (typeof v === 'function') return v.name ? '[Function ' + v.name + ']' : '[Function]';
    if (v instanceof Error) {
      const code = v.code ? ' (' + v.code + ')' : '';
      return (v.stack || (v.name + ': ' + v.message)) + code;
    }
    try {
      const s = JSON.stringify(v);
      if (typeof s === 'string') return s;
    } catch (e) { /* a cycle, or a host object */ }
    return String(v);
  }

  function cells(s) {
    let n = 0;
    for (let i = 0; i < s.length; i++) {
      const c = s.charCodeAt(i);
      if (c >= 0xD800 && c <= 0xDBFF) { n += 2; i++; continue; }
      n += (c >= 0x1100 && (c <= 0x115F || c === 0x2329 || c === 0x232A ||
        (c >= 0x2E80 && c <= 0xA4CF && c !== 0x303F) || (c >= 0xAC00 && c <= 0xD7A3) ||
        (c >= 0xF900 && c <= 0xFAFF) || (c >= 0xFE10 && c <= 0xFE19) ||
        (c >= 0xFE30 && c <= 0xFE6F) || (c >= 0xFF00 && c <= 0xFF60) ||
        (c >= 0xFFE0 && c <= 0xFFE6))) ? 2 : 1;
    }
    return n;
  }

  let session = null;

  function refreshDocument(term) {
    try { if (L.layout && L.layout.render) L.layout.render(); } catch (e) { /* the next line prints it */ }
    try { if (L.ed && L.ed.restoreDom) L.ed.restoreDom(true); } catch (e) { /* the page still has the model */ }
    const back = () => { try { term.focus(); } catch (e) { /* the window is gone */ } };
    back();
    setTimeout(back, 0);
  }

  function attach(host, dlg, bar) {
    const term = new window.Terminal({
      cursorBlink: true,
      fontFamily: 'monospace',
      fontSize: 14,
      lineHeight: 1.2,
      scrollback: 2000,
      cols: 80,
      rows: 22,
      theme: { background: '#0c0c0c', foreground: '#cccccc', cursor: '#cccccc', selectionBackground: '#264f78' },
    });
    term.open(host);
    const box = { dlg: dlg, term: term };
    let lines = [''], row = 0, col = 0, history = [], hist = -1, draft = null, busy = false, pending = '', ignoreCR = false;

    const promptOf = (i) => (i ? '... ' : '> ');
    function writeOut(s) {
      term.write(String(s).replace(/\r\n/g, '\n').replace(/\n/g, '\r\n') + '\r\n');
    }
    function paint() {
      const line = lines[row];
      const after = line.slice(col);
      term.write('\r\x1b[K' + promptOf(row) + line);
      const back = cells(after);
      if (back) term.write('\x1b[' + back + 'D');
    }
    function paintAll() {
      lines.forEach((line, i) => {
        if (i) term.write('\r\n');
        term.write(promptOf(i) + line);
      });
      row = lines.length - 1;
      col = lines[row].length;
    }
    function eraseAll() {
      term.write('\r\x1b[K');
      while (row > 0) { term.write('\x1b[1A\r\x1b[K'); row--; }
    }
    function place(r, c) {
      const up = (lines.length - 1) - r;
      if (up > 0) term.write('\x1b[' + up + 'A');
      row = r;
      col = Math.max(0, Math.min(c, lines[row].length));
      const back = cells(lines[row].slice(col));
      if (back) term.write('\x1b[' + back + 'D');
    }
    function hookConsole() {
      const names = ['log', 'info', 'warn', 'error'];
      const prev = {};
      names.forEach((name) => {
        prev[name] = console[name];
        console[name] = function () {
          writeOut(Array.prototype.map.call(arguments, formatValue).join(' '));
          if (typeof prev[name] === 'function') prev[name].apply(console, arguments);
        };
      });
      return () => names.forEach((name) => { console[name] = prev[name]; });
    }
    async function runUser(code) {
      const seq = L.D && L.D.hist ? L.D.hist.seq : null;
      const doc = L.D ? L.D.doc : null;
      const restore = hookConsole();
      try {
        const value = await evaluate(code);
        if (value !== undefined && !startsWithFunction(code)) writeOut(formatValue(value));
      } catch (e) {
        writeOut(formatValue(e));
      } finally {
        restore();
        if (L.D && (L.D.doc !== doc || (seq !== null && L.D.hist && L.D.hist.seq !== seq))) refreshDocument(term);
      }
    }
    async function submit() {
      if (busy) return;
      busy = true;
      const code = lines.join('\n').trimEnd();
      term.write('\r\n');
      lines = ['']; row = 0; col = 0; hist = -1; draft = null;
      try {
        if (code.trim()) {
          if (history[history.length - 1] !== code) history.push(code);
          if (history.length > 100) history.shift();
          await runUser(code);
        }
      } finally {
        busy = false;
        if (term.element && term.element.isConnected) { term.write('> '); term.focus(); }
      }
    }
    function insert(text) {
      const chunk = String(text).replace(/\r\n/g, '\n').replace(/\r/g, '\n');
      if (!chunk) return;
      if (!chunk.includes('\n') && row === lines.length - 1 && col === lines[row].length) {
        lines[row] += chunk;
        col += chunk.length;
        term.write(chunk);
        return;
      }
      const startRow = row, startCol = col;
      const cur = lines[row];
      const merged = cur.slice(0, col) + chunk + cur.slice(col);
      let parts = merged.split('\n');
      const head = lines.slice(0, row);
      const tail = lines.slice(row + 1);
      let run = false;
      if (chunk.endsWith('\n')) {
        const candidate = head.concat(parts.slice(0, -1), tail).join('\n');
        if (!needsMore(candidate)) { parts = parts.slice(0, -1); run = true; }
      }
      if (!parts.length) parts = [''];
      const added = chunk.split('\n');
      eraseAll();
      lines = head.concat(parts, tail);
      paintAll();
      if (run) { submit(); return; }
      const destRow = Math.min(lines.length - 1, startRow + added.length - 1);
      const destCol = added.length === 1 ? startCol + added[0].length : added[added.length - 1].length;
      place(destRow, destCol);
    }
    function splitLine() {
      const at = row;
      const left = lines[at].slice(0, col), right = lines[at].slice(col);
      eraseAll();
      lines[at] = left;
      lines.splice(at + 1, 0, right);
      paintAll();
      place(at + 1, 0);
    }
    function enter(force) {
      if (force || needsMore(lines.join('\n'))) splitLine();
      else submit();
    }
    function backspace() {
      if (col > 0) {
        lines[row] = lines[row].slice(0, col - 1) + lines[row].slice(col);
        col--;
        paint();
        return;
      }
      if (row === 0) return;
      const extra = lines[row], prev = row - 1, at = lines[prev].length;
      eraseAll();
      lines[prev] += extra;
      lines.splice(prev + 1, 1);
      paintAll();
      place(prev, at);
    }
    function del() {
      if (col < lines[row].length) {
        lines[row] = lines[row].slice(0, col) + lines[row].slice(col + 1);
        paint();
        return;
      }
      if (row >= lines.length - 1) return;
      const extra = lines[row + 1], r = row, c = col;
      eraseAll();
      lines[r] += extra;
      lines.splice(r + 1, 1);
      paintAll();
      place(r, c);
    }
    function show(text) {
      eraseAll();
      lines = String(text).split('\n');
      if (!lines.length) lines = [''];
      paintAll();
    }
    function up() {
      if (row > 0) { row--; col = Math.min(col, lines[row].length); term.write('\x1b[1A'); paint(); return; }
      if (!history.length) return;
      if (hist < 0) { draft = lines.join('\n'); hist = history.length; }
      if (hist === 0) return;
      hist--;
      show(history[hist]);
    }
    function down() {
      if (row < lines.length - 1) { row++; col = Math.min(col, lines[row].length); term.write('\x1b[1B'); paint(); return; }
      if (hist < 0) return;
      hist++;
      if (hist >= history.length) { hist = -1; show(draft || ''); }
      else show(history[hist]);
    }
    function left() {
      if (col > 0) { col--; term.write('\x1b[' + cells(lines[row].slice(col, col + 1)) + 'D'); return; }
      if (row === 0) return;
      row--; col = lines[row].length; term.write('\x1b[1A'); paint();
    }
    function right() {
      if (col < lines[row].length) { const ch = lines[row].slice(col, col + 1); col++; term.write('\x1b[' + cells(ch) + 'C'); return; }
      if (row >= lines.length - 1) return;
      row++; col = 0; term.write('\x1b[1B'); paint();
    }
    function cancel() {
      term.write('^C\r\n');
      lines = ['']; row = 0; col = 0; hist = -1; draft = null;
      term.write('> ');
    }
    function key(seq) {
      if (seq === '\r' || seq === '\n') { if (ignoreCR) { ignoreCR = false; return; } enter(false); return; }
      if (seq === '\x7f' || seq === '\b') { backspace(); return; }
      if (seq === '\x03') { cancel(); return; }
      if (seq === '\t') { insert('  '); return; }
      if (seq === '\x1b[A' || seq === '\x1bOA') { up(); return; }
      if (seq === '\x1b[B' || seq === '\x1bOB') { down(); return; }
      if (seq === '\x1b[D' || seq === '\x1bOD') { left(); return; }
      if (seq === '\x1b[C' || seq === '\x1bOC') { right(); return; }
      if (seq === '\x1b[H' || seq === '\x1b[1~' || seq === '\x1bOH') { col = 0; paint(); return; }
      if (seq === '\x1b[F' || seq === '\x1b[4~' || seq === '\x1bOF') { col = lines[row].length; paint(); return; }
      if (seq === '\x1b[3~') del();
    }
    term.attachCustomKeyEventHandler((ev) => {
      if (ev.type !== 'keydown' || busy) return true;
      if (ev.key === 'Enter' && ev.shiftKey) {
        ignoreCR = true;
        enter(true);
        Promise.resolve().then(() => { ignoreCR = false; });
        return false;
      }
      if (ev.key === 'Tab') { insert('  '); return false; }
      return true;
    });
    term.onData((data) => {
      if (busy) return;
      pending += data;
      while (pending) {
        if (pending[0] === '\x1b') {
          const m = /^\x1b(?:\[[0-9;?]*[A-Za-z~]|O[A-Z])/.exec(pending);
          if (!m) { if (pending.length > 16) pending = pending.slice(1); else return; continue; }
          key(m[0]);
          pending = pending.slice(m[0].length);
          continue;
        }
        const ch = pending.charCodeAt(0);
        if (ch < 32 || ch === 127) { key(pending[0]); pending = pending.slice(1); continue; }
        const cut = pending.search(/[\u0000-\u001f\u007f]/);
        const take = cut < 0 ? pending : pending.slice(0, cut);
        insert(take);
        pending = cut < 0 ? '' : pending.slice(cut);
      }
    });
    function writeClipboard(text) {
      const value = String(text);
      if (navigator.clipboard && navigator.clipboard.writeText) {
        return navigator.clipboard.writeText(value).catch(() => writeClipboardFallback(value));
      }
      writeClipboardFallback(value);
      return Promise.resolve();
    }
    function writeClipboardFallback(value) {
      const ta = document.createElement('textarea');
      ta.value = value;
      ta.setAttribute('readonly', '');
      ta.style.position = 'fixed';
      ta.style.left = '-9999px';
      document.body.appendChild(ta);
      ta.select();
      try { document.execCommand('copy'); } catch (e) { /* the browser refused */ }
      ta.remove();
    }
    function copyOut() {
      let text = '';
      try { if (term.hasSelection()) text = term.getSelection() || ''; } catch (e) { /* no selection */ }
      if (!text) text = lines.join('\n');
      writeClipboard(text).then(() => { if (term.element && term.element.isConnected) term.focus(); });
    }
    function askPaste() {
      const note = L.h('div', { class: 'dom-console-paste-note', text: 'Paste the text here, then choose OK.' });
      const area = L.h('textarea', { class: 'dom-console-paste', rows: '12' });
      L.ui.dialog({
        title: 'Paste',
        body: L.h('div', {}, note, area),
        width: 520,
        buttons: [
          { label: 'OK', primary: true, onClick: () => { if (area.value) insert(area.value); } },
          { label: 'Cancel' },
        ],
        onClose: () => setTimeout(() => { if (term.element && term.element.isConnected) term.focus(); }, 0),
      });
    }
    function pasteIn() {
      if (busy) return;
      const read = navigator.clipboard && navigator.clipboard.readText
        ? navigator.clipboard.readText()
        : Promise.reject(new Error('clipboard'));
      read.then((text) => {
        if (text) { insert(text); if (term.element && term.element.isConnected) term.focus(); return; }
        askPaste();
      }).catch(() => askPaste());
    }
    function clearScreen() {
      if (busy) return;
      lines = ['']; row = 0; col = 0; hist = -1; draft = null; pending = '';
      term.write('\x1b[2J\x1b[3J\x1b[H> ');
      term.focus();
    }
    if (bar) {
      const toolButton = (o, fn) => {
        const b = L.h('button', { class: 'tb-btn with-label', type: 'button', 'aria-label': o.label, 'data-tip': o.label });
        const icon = o.icon && L.icons ? L.icons.get(o.icon) : '';
        const text = L.ui.ampHTML ? L.ui.ampHTML(o.label) : o.label;
        b.innerHTML = icon + '<span class="tb-lbl">' + text + '</span>';
        b.addEventListener('pointerdown', (e) => e.preventDefault());
        b.addEventListener('click', (e) => { e.preventDefault(); fn(); });
        return b;
      };
      bar.append(
        toolButton({ label: 'Copy', icon: 'copy' }, copyOut),
        toolButton({ label: 'Paste', icon: 'paste' }, pasteIn),
        toolButton({ label: 'Clear', icon: 'eraser' }, clearScreen),
        L.h('span', { class: 'tb-sep' }),
        toolButton({ label: 'DOM API References', icon: 'read' }, () => { /* the reference is not written yet */ }),
      );
    }
    term.write([
      L.t('Welcome to the DOM Console.'),
      L.t('Tips:'),
      L.t('1. If the right-click menu cannot paste, use the Paste button.'),
      L.t('2. Outside a function, do not declare variables with let, var, or const. Each entry runs on its own, so those declarations do not carry to the next line. Assign directly, as in x = 1. The name is then on the global object, and later entries can use it.'),
    ].join('\r\n') + '\r\n');
    term.write('> ');
    term.focus();
    return box;
  }

  function open() {
    if (typeof document === 'undefined' || !L.ui || typeof L.ui.dialog !== 'function') return;
    if (typeof window.Terminal !== 'function') {
      if (L.ui.msg) L.ui.msg('The terminal library did not load.');
      return;
    }
    if (session && session.dlg.isConnected) {
      try { session.dlg.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true })); } catch (e) { /* focus is enough */ }
      session.term.focus();
      return;
    }
    const bar = L.h('div', { class: 'dom-console-bar' });
    const host = L.h('div', { class: 'dom-console-host' });
    host.addEventListener('keydown', (e) => e.stopPropagation());
    L.ui.dialog({
      title: 'Quire DOM Console',
      body: L.h('div', {}, bar, host),
      width: 720,
      buttons: [],
      modeless: true,
      onOpen: (el) => {
        el.classList.add('dom-console');
        el.addEventListener('keydown', (e) => {
          if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); }
        }, true);
        session = attach(host, el, bar);
      },
      onClose: () => {
        if (session) { try { session.term.dispose(); } catch (e) { /* already gone */ } session = null; }
      },
    });
  }

  L.domConsole = { open: open, evaluate: evaluate, needsMore: needsMore, hoistFunctions: hoistFunctions };
})(window);
