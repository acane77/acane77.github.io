/* Ledger DOM Console. Help opens a modeless window; the script runs in this page.
   Visual Basic 6 uses evalVB6. Run Script uses runVB6.
   An empty Run Script buffer opens on a sample Sub Main that writes a small sheet.
   Office vbaProject.bin is not listed or run. A macro opened here keeps its own source.
   The object model is js/dom.js. xterm is public/common/vendor/xterm/.
   The script editor is public/common/vendor/codemirror/. */
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

  /** VB source with comments removed. Strings stay, so a quote inside one is not a comment. */
  function vbCode(line) {
    let out = '';
    let q = false;
    for (let i = 0; i < line.length; i++) {
      const c = line[i];
      if (q) {
        out += c;
        if (c === '"') {
          if (line[i + 1] === '"') { out += '"'; i++; }
          else q = false;
        }
        continue;
      }
      if (c === '"') { q = true; out += c; continue; }
      if (c === "'") break;
      out += c;
    }
    return out.trim();
  }

  /** Join lines that end in the VB continuation character. */
  function joinVb(src) {
    const lines = String(src).replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
    const out = [];
    let buf = null;
    for (let i = 0; i < lines.length; i++) {
      const raw = lines[i];
      const code = vbCode(raw);
      const continues = /_\s*$/.test(code);
      const piece = continues ? raw.replace(/_\s*$/, '').replace(/\s+$/, '') : raw;
      buf = buf == null ? piece : buf + ' ' + piece.trim();
      if (continues) continue;
      out.push(buf);
      buf = null;
    }
    if (buf != null) out.push(buf);
    return out.join('\n');
  }

  function vbParts(stmt) {
    const parts = [];
    let start = 0;
    let q = false;
    for (let i = 0; i < stmt.length; i++) {
      const c = stmt[i];
      if (q) {
        if (c === '"') { if (stmt[i + 1] === '"') i++; else q = false; }
        continue;
      }
      if (c === '"') { q = true; continue; }
      if (c === ':') { parts.push(stmt.slice(start, i)); start = i + 1; }
    }
    parts.push(stmt.slice(start));
    return parts;
  }

  /** A line ending in _, or an unfinished Sub, Function, If or With, is not ready to run. */
  function vbNeedsMore(src) {
    const rawLines = String(src).replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
    let openCont = false;
    for (let i = 0; i < rawLines.length; i++) openCont = /_\s*$/.test(vbCode(rawLines[i]));
    if (openCont) return true;
    let sub = 0, fn = 0, iff = 0, wit = 0;
    const logical = joinVb(src).split('\n');
    for (let i = 0; i < logical.length; i++) {
      const parts = vbParts(vbCode(logical[i]));
      for (let p = 0; p < parts.length; p++) {
        const s = parts[p].trim();
        if (!s || /^rem\b/i.test(s)) continue;
        if (/^end\s+sub\b/i.test(s)) sub = Math.max(0, sub - 1);
        else if (/^(?:(?:public|private|friend|static)\s+)*sub\b/i.test(s)) sub++;
        else if (/^end\s+function\b/i.test(s)) fn = Math.max(0, fn - 1);
        else if (/^(?:(?:public|private|friend|static)\s+)*function\b/i.test(s)) fn++;
        else if (/^end\s+with\b/i.test(s)) wit = Math.max(0, wit - 1);
        else if (/^with\b/i.test(s)) wit++;
        else if (/^end\s+if\b/i.test(s)) iff = Math.max(0, iff - 1);
        else if (/^if\b/i.test(s)) {
          const m = /\bthen\b/i.exec(s);
          const after = m ? s.slice(m.index + 4).trim() : '';
          if (!m || after === '') iff++;
        }
      }
    }
    return sub > 0 || fn > 0 || iff > 0 || wit > 0;
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
  let scriptBox = null;
  const LANG_KEY = 'vibeoffice.ledger.domConsole.lang';
  const SCRIPT_KEY = 'vibeoffice.ledger.vbScript';
  const SCRIPT_LOG_KEY = 'vibeoffice.ledger.vbScriptLog';
  /* Shown when Run Script has never been edited, or the buffer was empty when it last closed.
     The sample writes a few cells on the active sheet. The same text is tests/ledger/vb-playground.bas. */
  const DEFAULT_SCRIPT =
    "' Welcome to Visual Basic Playground\n" +
    "' Write your own Visual Basic code in this editor,\n" +
    "' then click Run to run and test it.\n" +
    "'\n" +
    "' See also: Office VBA Documentation:\n" +
    "'    https://learn.microsoft.com/en-us/office/vba/api/overview/excel\n" +
    "\n" +
    "' Your code should contain a Main procedure\n" +
    'Sub Main()\n' +
    "  ' Write a short list and a total on the active sheet.\n" +
    '  ActiveSheet.Range("A1").Value = "Item"\n' +
    '  ActiveSheet.Range("B1").Value = "Amount"\n' +
    '  ActiveSheet.Range("A2").Value = "Pens"\n' +
    '  ActiveSheet.Range("B2").Value = 12.5\n' +
    '  ActiveSheet.Range("A3").Value = "Paper"\n' +
    '  ActiveSheet.Range("B3").Value = 7.5\n' +
    '  ActiveSheet.Range("B4").Formula = "=SUM(B2:B3)"\n' +
    '  ActiveSheet.Range("A1:B1").Font.Bold = True\n' +
    '  ActiveSheet.Range("B2:B4").NumberFormat = "0.00"\n' +
    '  ActiveWorkbook.Names.Add "Total", "=Sheet1!$B$4"\n' +
    '  ActiveSheet.Range("A6").Value = WorksheetFunction.Sum(ActiveSheet.Range("B2:B3"))\n' +
    'End Sub\n';

  function scriptSource(src) {
    const text = src == null ? '' : String(src);
    return text.trim() ? text : DEFAULT_SCRIPT;
  }

  function storedLang() {
    try {
      return localStorage.getItem(LANG_KEY) === 'js' ? 'js' : 'vb';
    } catch (e) {
      return 'vb';
    }
  }

  function storeLang(value) {
    try { localStorage.setItem(LANG_KEY, value); } catch (e) { /* private mode */ }
  }

  function refreshDocument(focus) {
    try { if (L.grid && L.grid.paint) L.grid.paint(); } catch (e) { /* the page still has the model */ }
    const back = () => { try { if (focus) focus(); } catch (e) { /* the window is gone */ } };
    back();
    setTimeout(back, 0);
  }

  function attach(host, dlg, bar) {
    let lang = storedLang();
    const more = (src) => (lang === 'vb' ? vbNeedsMore(src) : needsMore(src));
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
      const restore = lang === 'vb' ? null : hookConsole();
      try {
        if (lang === 'vb') await runImmediate(code);
        else {
          const value = await evaluate(code);
          if (value !== undefined && !startsWithFunction(code)) writeOut(formatValue(value));
        }
      } catch (e) {
        writeOut(formatValue(e));
      } finally {
        if (restore) restore();
        refreshDocument(() => term.focus());
      }
    }
    async function runImmediate(code) {
      const host = await L.vb.session();
      const stopPrints = followPrints(host, (line) => writeOut(line));
      try {
        await evalVB6(joinVb(code));
      } finally {
        stopPrints();
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
        if (!more(candidate)) { parts = parts.slice(0, -1); run = true; }
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
      if (force || more(lines.join('\n'))) splitLine();
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
    function welcome() {
      const lines = [
        L.t('Welcome to the DOM Console.'),
        L.t('Tips:'),
        L.t('1. If the right-click menu cannot paste, use the Paste button.'),
      ];
      if (lang === 'vb') {
        lines.push(L.t('2. A line is a Visual Basic statement. Assign directly, as in x = 1. The name stays, and a later line can use it. Read a value with ?x.'));
        lines.push(L.t('3. A line ending in _ continues on the next line. An unfinished Sub, Function, If or With stays open until it ends. A whole script belongs in Run Script.'));
      } else {
        lines.push(L.t('2. Outside a function, do not declare variables with let, var, or const. Each entry runs on its own, so those declarations do not carry to the next line. Assign directly, as in x = 1. The name is then on the global object, and later entries can use it.'));
        lines.push(L.t('3. JavaScript runs here as well, but the F12 console is the better place to run a JavaScript script.'));
      }
      return lines.join('\r\n') + '\r\n';
    }
    function clearScreen(reprint) {
      if (busy) return;
      lines = ['']; row = 0; col = 0; hist = -1; draft = null; pending = '';
      term.write('\x1b[2J\x1b[3J\x1b[H');
      if (reprint) term.write(welcome());
      term.write('> ');
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
      const langSel = L.ui.select([
        ['vb', 'Visual Basic 6'],
        ['js', 'JavaScript'],
      ], lang, (value) => {
        if (busy || value === lang) { langSel.value = lang; return; }
        lang = value === 'js' ? 'js' : 'vb';
        storeLang(lang);
        if (runBtn) runBtn.hidden = lang !== 'vb';
        clearScreen(true);
      }, { id: 'dom-console-lang' });
      const runBtn = toolButton({ label: 'Run Script' }, () => openScript());
      runBtn.hidden = lang !== 'vb';
      bar.append(
        toolButton({ label: 'Copy', icon: 'copy' }, copyOut),
        toolButton({ label: 'Paste', icon: 'paste' }, pasteIn),
        toolButton({ label: 'Clear', icon: 'eraser' }, clearScreen),
        L.h('span', { class: 'tb-sep' }),
        toolButton({ label: 'DOM API References', icon: 'read' }, () => { if (L.vbDomRef) L.vbDomRef.show(); }),
        L.h('span', { class: 'tb-sep' }),
        L.h('span', { class: 'dom-console-lang' },
          L.h('label', { for: 'dom-console-lang', html: L.ui.ampHTML('&Language:') }),
          langSel),
        runBtn,
      );
    }
    term.write(welcome());
    term.write('> ');
    term.focus();
    return box;
  }

  /* Prints land in host.output as they happen. The Output box used to read that
     list only after the script returned, so a MsgBox hid every Debug.Print until No. */
  function followPrints(host, write) {
    let seen = host.output.length;
    function flush() {
      if (!host.output || seen >= host.output.length) return;
      const next = host.output.slice(seen);
      seen = host.output.length;
      for (let i = 0; i < next.length; i++) write(next[i]);
    }
    const prev = host.onPrint;
    function onPrint() { flush(); }
    host.onPrint = onPrint;
    return function stop() {
      if (host.onPrint === onPrint) host.onPrint = prev;
      flush();
    };
  }

  function rememberScript(text, log) {
    try {
      localStorage.setItem(SCRIPT_KEY, text);
      localStorage.setItem(SCRIPT_LOG_KEY, log);
    } catch (e) { /* private mode */ }
  }

  function openScript(opts) {
    opts = opts || {};
    if (typeof document === 'undefined' || !L.ui || typeof L.ui.dialog !== 'function') return;
    if (scriptBox && scriptBox.dlg.isConnected) {
      try { scriptBox.dlg.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true })); } catch (e) { /* focus is enough */ }
      if (scriptBox.show) scriptBox.show(opts.macro || null);
      else if (scriptBox.editor) scriptBox.editor.focus();
      return;
    }
    if (typeof window.CodeMirror !== 'function') {
      if (L.ui.msg) L.ui.msg('The script editor did not load.');
      return;
    }
    const nameInput = L.h('input', { type: 'text', id: 'vb-script-macro-name', class: 'vb-script-name-input', spellcheck: 'false', autocomplete: 'off', translate: 'no' });
    nameInput.readOnly = true;
    const nameRow = L.h('div', { class: 'vb-script-name', hidden: true },
      L.h('label', { for: 'vb-script-macro-name', html: L.ui.ampHTML('Macro na&me:') }),
      nameInput);
    const area = L.h('textarea', { class: 'vb-script-src' });
    const log = L.h('textarea', { class: 'vb-script-out', rows: '6', readonly: 'readonly', spellcheck: 'false' });
    const editorHost = L.h('div', { class: 'vb-script-editor' });
    let editor = null;
    let saveBtn = null;
    let box = null;
    let macro = null;
    let savedSource = '';
    let running = false;
    let asking = false;
    let rejecting = false;
    function text() { return editor ? editor.getValue() : area.value; }
    function setText(value) {
      if (editor) editor.setValue(value);
      else area.value = value;
    }
    function scratchText() {
      try { return { src: localStorage.getItem(SCRIPT_KEY) || '', log: localStorage.getItem(SCRIPT_LOG_KEY) || '' }; }
      catch (e) { return { src: '', log: '' }; }
    }
    function persist() { if (macro) return; rememberScript(text(), log.value); }
    function append(line) {
      const chunk = String(line);
      if (!chunk) return;
      log.value += (log.value && !log.value.endsWith('\n') ? '\n' : '') + chunk + (chunk.endsWith('\n') ? '' : '\n');
      log.scrollTop = log.scrollHeight;
      persist();
    }
    function save() {
      if (!macro || !L.macros || !L.macros.commit) return false;
      const ok = L.macros.commit(macro.name, text(), { created: macro.created, author: macro.author, store: macro.store });
      if (ok) { savedSource = text(); macro.source = savedSource; }
      else if (macro.store === 'all' && L.ui && L.ui.msg) L.ui.msg('The macros stored on this computer could not be saved.', { icon: 'error', title: 'Macros' });
      return !!ok;
    }
    function apply(next) {
      macro = next || null;
      nameRow.hidden = !macro;
      if (saveBtn) saveBtn.hidden = !macro;
      if (macro) {
        nameInput.value = macro.name || '';
        nameInput.readOnly = true;
        savedSource = macro.source || '';
        log.value = '';
        setText(savedSource);
      } else {
        const stored = scratchText();
        savedSource = '';
        log.value = stored.log;
        setText(scriptSource(stored.src));
      }
      if (editor) editor.focus();
    }
    async function askClose() {
      if (!macro || text() === savedSource) { if (box) box.close(); return; }
      if (asking) return;
      asking = true;
      try {
        const i = await L.ui.msg('Save the changes to this macro?', { icon: 'question', title: 'Run Visual Basic Script', buttons: [L.t('Save'), L.t("Don't Save"), L.t('Cancel')] });
        if (i === 0) { if (await saveChecked() && box) box.close(); }
        else if (i === 1 && box) box.close();
      } finally { asking = false; }
    }
    function legalName(next) {
      return !next || !!(L.macros && L.macros.vbIdentifier && L.macros.vbIdentifier(next.name));
    }
    async function rejectName(next) {
      if (legalName(next) || rejecting) return false;
      rejecting = true;
      try {
        if (L.ui && L.ui.msg) await L.ui.msg('The procedure name is not a legal Visual Basic identifier.', { icon: 'warn', title: 'Run Visual Basic Script' });
      } finally {
        rejecting = false;
        if (box && !legalName(macro)) box.close();
      }
      return true;
    }
    async function show(next) {
      const home = (m) => (m && m.store === 'all' ? 'all' : 'document');
      const same = !!(macro && next && next.name === macro.name && home(next) === home(macro) && text() === (next.source || ''));
      if (same) { if (editor) editor.focus(); }
      else {
        if (macro && text() !== savedSource) {
          if (asking) return;
          asking = true;
          let i;
          try {
            i = await L.ui.msg('Save the changes to this macro?', { icon: 'question', title: 'Run Visual Basic Script', buttons: [L.t('Save'), L.t("Don't Save"), L.t('Cancel')] });
          } finally { asking = false; }
          if (i === 0) { if (!(await saveChecked())) return; }
          else if (i !== 1) return;
        }
        apply(next);
      }
      await rejectName(next);
    }
    async function compiles() {
      const code = text();
      const editing = macro;
      let compiled;
      try {
        if (!L.vb) throw new Error(L.t('The VB runtime is not running'));
        if (editing) {
          if (!L.macros || !L.macros.draftModules || !L.vb.compileDraft) throw new Error(L.t('The VB runtime is not running'));
          compiled = await L.vb.compileDraft(L.macros.draftModules(editing.name, code, editing.store));
        } else {
          if (!L.vb.compileScript) throw new Error(L.t('The VB runtime is not running'));
          compiled = await L.vb.compileScript(code);
        }
      } catch (e) {
        if (L.ui && L.ui.msg) await L.ui.msg(e && e.message ? e.message : String(e), { icon: 'error', title: 'Run Visual Basic Script' });
        return false;
      }
      if (text() !== code || macro !== editing) return false;
      if (compiled && compiled.valid) return true;
      if (L.vbIde && L.vbIde.reportDiagnostics) await L.vbIde.reportDiagnostics(compiled && compiled.diagnostics, 'Run Visual Basic Script');
      return false;
    }
    async function saveChecked() {
      if (!macro || !L.macros || !L.macros.commit) return false;
      if (!(await compiles())) return false;
      return save();
    }
    async function run() {
      if (running) return;
      if (!(await compiles())) return;
      running = true;
      const code = text();
      const editing = macro;
      /* One run, one result. The saved log is the last run, not a transcript of every earlier one. */
      log.value = '';
      if (!editing) persist();
      let stopPrints = null;
      try {
        const host = await L.vb.session();
        stopPrints = followPrints(host, append);
        const value = editing && L.macros ? await L.macros.runDraft(editing.name, code, editing.store) : await L.runVB6(code);
        stopPrints();
        stopPrints = null;
        if (value !== undefined) append(editing ? String(value) : formatValue(value));
      } catch (e) {
        if (stopPrints) { stopPrints(); stopPrints = null; }
        if (!(e && e.quiet)) append(editing && e && e.message ? e.message : formatValue(e));
      } finally {
        if (stopPrints) stopPrints();
        running = false;
        refreshDocument(() => { if (editor) editor.focus(); });
      }
    }
    box = L.ui.dialog({
      title: 'Run Visual Basic Script',
      body: L.h('div', { class: 'vb-script-body' },
        editorHost,
        nameRow,
        L.h('div', { class: 'vb-script-out-label', text: 'Output' }),
        log),
      width: 720,
      buttons: [
        { label: 'Save', onClick: () => { saveChecked(); return false; } },
        { label: 'Run', primary: true, onClick: () => { run(); return false; } },
        { label: 'Close', onClick: () => { askClose(); return false; } },
      ],
      modeless: true,
      onOpen: (el) => {
        el.classList.add('vb-script');
        el.addEventListener('click', (e) => {
          if (!e.target.closest || !e.target.closest('.dlg-x') || e.target.closest('.dlg-help')) return;
          if (!macro || text() === savedSource) return;
          e.preventDefault();
          e.stopImmediatePropagation();
          askClose();
        }, true);
        saveBtn = el.querySelector('.dlg-foot .btn');
        editorHost.appendChild(area);
        editor = window.CodeMirror.fromTextArea(area, {
          mode: 'text/x-vb',
          lineNumbers: true,
          indentUnit: 2,
          tabSize: 2,
        });
        editor.on('change', () => { if (!macro) persist(); });
        scriptBox = { dlg: el, editor: editor, show: show };
        apply(opts.macro || null);
        setTimeout(() => { if (editor) editor.refresh(); }, 0);
      },
      onClose: () => { if (!macro) persist(); scriptBox = null; },
    });
    rejectName(opts.macro || null);
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
      title: 'Ledger DOM Console',
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

  L.domConsole = {
    open: open,
    openScript: openScript,
    evaluate: evaluate,
    needsMore: needsMore,
    vbNeedsMore: vbNeedsMore,
    joinVb: joinVb,
    hoistFunctions: hoistFunctions,
    scriptSource: scriptSource,
    followPrints: followPrints,
  };
})(window);
