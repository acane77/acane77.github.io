/* VibeOffice — core utilities, shared by Quire, Ledger and Lectern.
 * Shared namespace, DOM helpers, units, colors, events, file I/O, media store, PDF writer.
 * Every other script attaches itself to window.L. Each page names its app before loading this file:
 *   <script>window.L = { APP_ID: 'quire', APP_NAME: 'Quire', APP: 'Quire' };</script>
 * APP_ID keeps each app's saved settings apart (localStorage 'quire.…'); APP_NAME and APP appear in messages.
 */
(function (root) {
  'use strict';
  const L = root.L || (root.L = {});
  L.VERSION = root.VO?.VERSION || 'dev';   // one version for the suite: common/suite.js
  L.APP_ID = L.APP_ID || 'vibeoffice';
  L.APP_NAME = L.APP_NAME || 'VibeOffice';
  L.APP = L.APP || L.APP_NAME;

  /* ---------- DOM ---------- */
  /* the text h() puts in the UI is shown in the UI's language (common/i18n.js): text, these attributes and string
     children whose whole text has a translation */
  const SHOWN = { title: 1, placeholder: 1, 'aria-label': 1, 'data-tip': 1 };
  /* Office writes a CJK access key after the words, 文件(&F). Ribbon Golden Gate does not show it. */
  L.dropCjkAccess = (s) => (typeof s === 'string' && typeof document !== 'undefined' && document.documentElement && document.documentElement.dataset.look === 'goldengate' ? s.replace(/\s*[（(]\s*&[^&\s]\s*[）)]/g, '') : s);
  const tr = (v, ...args) => L.dropCjkAccess(typeof v !== 'string' ? v : L.t ? L.t(v, ...args) : v.replace(/(\S)\|[A-Za-z][A-Za-z -]*$/, '$1').replace(/\{(\d+)\}/g, (m, i) => (args[i] != null ? args[i] : m)));
  L.h = function h(tag, attrs, ...kids) {
    const el = document.createElement(tag);
    /* translate: 'no' marks data (a font's name, a slide's title): shown as it is */
    const data = attrs && attrs.translate === 'no';
    if (attrs) {
      for (const k in attrs) {
        let v = attrs[k];
        if (v == null || v === false) continue;
        if (SHOWN[k] && !data) v = tr(v);
        if (k === 'class') el.className = v;
        else if (k === 'style' && typeof v === 'object') Object.assign(el.style, v);
        else if (k === 'style') el.style.cssText = v;
        else if (k === 'html') el.innerHTML = v;
        else if (k === 'text') el.textContent = data ? v : tr(v);
        else if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2), v);
        else if (k === 'dataset') Object.assign(el.dataset, v);
        else if (v === true) el.setAttribute(k, '');
        else el.setAttribute(k, v);
      }
    }
    appendKids(el, kids, data);
    return el;
  };
  function appendKids(el, kids, data) {
    for (const k of kids) {
      if (k == null || k === false) continue;
      if (Array.isArray(k)) appendKids(el, k, data);
      else if (k instanceof Node) el.appendChild(k);
      else el.appendChild(document.createTextNode(typeof k === 'string' && !data ? tr(k) : String(k)));
    }
  }
  L.$ = (s, r) => (r || document).querySelector(s);
  L.$$ = (s, r) => Array.from((r || document).querySelectorAll(s));
  L.svgNS = 'http://www.w3.org/2000/svg';
  L.s = function s(tag, attrs, ...kids) {
    const el = document.createElementNS(L.svgNS, tag);
    if (attrs) for (const k in attrs) {
      const v = attrs[k];
      if (v == null || v === false) continue;
      if (k === 'style' && typeof v === 'object') Object.assign(el.style, v);
      else el.setAttribute(k, v);
    }
    for (const k of kids.flat()) if (k != null) el.appendChild(k instanceof Node ? k : document.createTextNode(String(k)));
    return el;
  };
  L.clear = (el) => { while (el.firstChild) el.removeChild(el.firstChild); return el; };

  /* ---------- misc ---------- */
  let idc = 0;
  L.uid = (p) => (p || 'i') + Date.now().toString(36).slice(-4) + (idc++).toString(36) + Math.random().toString(36).slice(2, 6);
  L.clone = (o) => (o == null ? o : JSON.parse(JSON.stringify(o)));
  L.clamp = (v, a, b) => Math.max(a, Math.min(b, v));
  L.round = (v, d) => { const m = Math.pow(10, d == null ? 2 : d); return Math.round(v * m) / m; };
  L.debounce = (fn, ms) => { let t; return function (...a) { clearTimeout(t); t = setTimeout(() => fn.apply(this, a), ms); }; };
  L.rafThrottle = (fn) => { let q = false, args; return function (...a) { args = a; if (q) return; q = true; requestAnimationFrame(() => { q = false; fn.apply(this, args); }); }; };
  L.esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  L.xesc = (s) => String(s == null ? '' : s)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, '')
    .replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  L.isMac = typeof navigator !== 'undefined' && /Mac|iPhone|iPad/.test(navigator.platform || navigator.userAgent);
  /** iOS or Android (an iPad says it is a Mac, but has a touch screen) */
  L.isMobile = typeof navigator !== 'undefined' && (/Android|iPhone|iPad|iPod/.test(navigator.userAgent || '') || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1));
  L.deepMerge = function deepMerge(a, b) {
    if (b == null) return a;
    if (a == null || typeof a !== 'object' || typeof b !== 'object' || Array.isArray(b)) return L.clone(b);
    const out = Object.assign({}, a);
    for (const k in b) {
      if (b[k] === undefined) continue;
      const bv = b[k];
      if (bv && typeof bv === 'object' && !Array.isArray(bv) && out[k] && typeof out[k] === 'object' && !Array.isArray(out[k])) out[k] = deepMerge(out[k], bv);
      else out[k] = bv && typeof bv === 'object' ? L.clone(bv) : bv;
    }
    return out;
  };
  L.equal = (a, b) => JSON.stringify(a) === JSON.stringify(b);

  /* ---------- units ----------
   * The model stores geometry and font sizes in points.
   * 1 pt = 12700 EMU, 72 pt = 1 inch, 96 CSS px = 1 inch.
   */
  L.EMU = 12700;
  L.pt2emu = (pt) => Math.round((pt || 0) * 12700);
  L.emu2pt = (e) => (Number(e) || 0) / 12700;
  L.PX_PER_PT = 96 / 72;
  L.fmtIn = (pt) => (L.round(pt / 72, 2)) + '"';

  /* ---------- event bus ---------- */
  const handlers = {};
  L.bus = {
    on(ev, fn) { (handlers[ev] = handlers[ev] || []).push(fn); return () => this.off(ev, fn); },
    off(ev, fn) { const a = handlers[ev]; if (a) { const i = a.indexOf(fn); if (i >= 0) a.splice(i, 1); } },
    emit(ev, ...args) { const a = handlers[ev]; if (a) for (const f of a.slice()) { try { f(...args); } catch (e) { console.error(e); } } },
  };

  /* ---------- colors ---------- */
  const C = (L.color = {});
  C.hexToRgb = (hex) => {
    let h = String(hex || '#000').replace('#', '').trim();
    if (h.length === 3) h = h.split('').map((c) => c + c).join('');
    const n = parseInt(h.slice(0, 6), 16) || 0;
    return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
  };
  C.rgbToHex = (r, g, b) => '#' + [r, g, b].map((v) => L.clamp(Math.round(v), 0, 255).toString(16).padStart(2, '0')).join('').toUpperCase();
  /** RGB (0-255) to HSL (0-1): lightness the middle of the largest and smallest channel, saturation their difference
      over their sum (over 2 − their sum above half lightness), hue from the largest channel's sextant (Foley and
      van Dam, "RGB to HLS") */
  C.rgbToHsl = (r, g, b) => {
    const R = r / 255, G = g / 255, B = b / 255;
    const top = Math.max(R, G, B), bottom = Math.min(R, G, B), l = (top + bottom) / 2;
    if (top === bottom) return [0, 0, l];
    const span = top - bottom;
    const s = l > 0.5 ? span / (2 - top - bottom) : span / (top + bottom);
    let h;
    if (top === R) h = (G - B) / span + (G < B ? 6 : 0);
    else if (top === G) h = 2 + (B - R) / span;
    else h = 4 + (R - G) / span;
    return [h / 6, s, l];
  };
  /** HSL (0-1) to RGB (0-255): the algorithm of CSS Color Level 3, 4.2.4 ("hsl.to.rgb" and "hue.to.rgb") */
  C.hslToRgb = (h, s, l) => {
    if (s === 0) return [l * 255, l * 255, l * 255];
    const m2 = l < 0.5 ? l * (s + 1) : l + s - l * s, m1 = l * 2 - m2;
    const hueToRgb = (x) => {
      if (x < 0) x += 1;
      if (x > 1) x -= 1;
      if (x < 1 / 6) return m1 + (m2 - m1) * 6 * x;
      if (x < 1 / 2) return m2;
      if (x < 2 / 3) return m1 + (m2 - m1) * (2 / 3 - x) * 6;
      return m1;
    };
    return [hueToRgb(h + 1 / 3) * 255, hueToRgb(h) * 255, hueToRgb(h - 1 / 3) * 255];
  };
  C.mix = (a, b, t) => { const x = C.hexToRgb(a), y = C.hexToRgb(b); return C.rgbToHex(x[0] + (y[0] - x[0]) * t, x[1] + (y[1] - x[1]) * t, x[2] + (y[2] - x[2]) * t); };
  C.lighten = (hex, t) => C.mix(hex, '#FFFFFF', t);
  C.darken = (hex, t) => C.mix(hex, '#000000', t);
  C.luma = (hex) => { const [r, g, b] = C.hexToRgb(hex); return (0.299 * r + 0.587 * g + 0.114 * b) / 255; };
  C.rgba = (hex, a) => { const [r, g, b] = C.hexToRgb(hex); return a == null || a >= 1 ? C.rgbToHex(r, g, b) : `rgba(${r},${g},${b},${L.round(a, 3)})`; };
  C.gray = (hex) => { const v = Math.round(C.luma(hex) * 255); return C.rgbToHex(v, v, v); };
  /* OOXML colour modifiers (values in 1/1000 %) */
  const toLin = (c) => { c /= 255; return c <= 0.04045 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4); };
  const toS = (l) => { l = Math.min(1, Math.max(0, l)); return 255 * (l <= 0.0031308 ? l * 12.92 : 1.055 * Math.pow(l, 1 / 2.4) - 0.055); };
  C.applyMods = (hex, mods) => {
    if (!mods || !mods.length) return hex;
    let [r, g, b] = C.hexToRgb(hex);
    for (const m of mods) {
      const v = m.val / 100000;
      let [h, s, l] = C.rgbToHsl(r, g, b);
      switch (m.name) {
        case 'lumMod': l = l * v; [r, g, b] = C.hslToRgb(h, s, L.clamp(l, 0, 1)); break;
        case 'lumOff': l = l + v; [r, g, b] = C.hslToRgb(h, s, L.clamp(l, 0, 1)); break;
        case 'lum': [r, g, b] = C.hslToRgb(h, s, L.clamp(v, 0, 1)); break;
        case 'satMod': [r, g, b] = C.hslToRgb(h, L.clamp(s * v, 0, 1), l); break;
        case 'satOff': [r, g, b] = C.hslToRgb(h, L.clamp(s + v, 0, 1), l); break;
        case 'sat': [r, g, b] = C.hslToRgb(h, L.clamp(v, 0, 1), l); break;
        case 'hueMod': [r, g, b] = C.hslToRgb((h * v) % 1, s, l); break;
        case 'hueOff': [r, g, b] = C.hslToRgb((h + m.val / 21600000 + 1) % 1, s, l); break;
        case 'hue': [r, g, b] = C.hslToRgb((m.val / 21600000) % 1, s, l); break;
        /* DrawingML tint/shade work in linear light (scRGB), not on gamma-encoded values */
        case 'tint': [r, g, b] = [r, g, b].map((c) => toS(toLin(c) * v + (1 - v))); break;
        case 'shade': [r, g, b] = [r, g, b].map((c) => toS(toLin(c) * v)); break;
        case 'comp': [r, g, b] = C.hslToRgb((h + 0.5) % 1, s, l); break;
        case 'inv': r = 255 - r; g = 255 - g; b = 255 - b; break;
        case 'gray': { const y = 0.299 * r + 0.587 * g + 0.114 * b; r = g = b = y; break; }
        default: break;
      }
    }
    return C.rgbToHex(r, g, b);
  };
  C.PRESET = {
    black: '#000000', white: '#FFFFFF', red: '#FF0000', green: '#008000', blue: '#0000FF', yellow: '#FFFF00', cyan: '#00FFFF', magenta: '#FF00FF',
    gray: '#808080', grey: '#808080', silver: '#C0C0C0', maroon: '#800000', navy: '#000080', olive: '#808000', purple: '#800080', teal: '#008080', lime: '#00FF00',
    orange: '#FFA500', darkBlue: '#00008B', darkRed: '#8B0000', darkGreen: '#006400', ltGray: '#D3D3D3', dkGray: '#A9A9A9', lightGray: '#D3D3D3', darkGray: '#A9A9A9',
    gold: '#FFD700', pink: '#FFC0CB', brown: '#A52A2A', indigo: '#4B0082', violet: '#EE82EE', tan: '#D2B48C', beige: '#F5F5DC', lavender: '#E6E6FA', coral: '#FF7F50',
  };
  /* Office 2003 standard colour grid (the 40-swatch palette) */
  C.STANDARD = [
    '#000000', '#993300', '#333300', '#003300', '#003366', '#000080', '#333399', '#333333',
    '#800000', '#FF6600', '#808000', '#008000', '#008080', '#0000FF', '#666699', '#808080',
    '#FF0000', '#FF9900', '#99CC00', '#339966', '#33CCCC', '#3366FF', '#800080', '#969696',
    '#FF00FF', '#FFCC00', '#FFFF00', '#00FF00', '#00FFFF', '#00CCFF', '#993366', '#C0C0C0',
    '#FF99CC', '#FFCC99', '#FFFF99', '#CCFFCC', '#CCFFFF', '#99CCFF', '#CC99FF', '#FFFFFF',
  ];

  /* ---------- storage (never relied on) ---------- */
  L.store = {
    get(k, d) { try { const v = localStorage.getItem(L.APP_ID + '.' + k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } },
    set(k, v) { try { localStorage.setItem(L.APP_ID + '.' + k, JSON.stringify(v)); return true; } catch (e) { return false; } },
    del(k) { try { localStorage.removeItem(L.APP_ID + '.' + k); } catch (e) { /* ignore */ } },
  };

  /* ---------- the look (common/looks.css) for chrome drawn on canvases ---------- */
  let lookCache = {};
  /** a look setting by name ('ruler-off' reads --ruler-off), or d where there is no page (Node) */
  L.lookColor = (name, d) => {
    if (!(name in lookCache)) lookCache[name] = typeof document === 'undefined' ? '' : getComputedStyle(document.documentElement).getPropertyValue('--' + name).trim();
    return lookCache[name] || d;
  };
  /** call fn after the look changes (VO.look.set, here or in another window), once the new values apply */
  L.onLook = (fn) => { if (typeof window !== 'undefined' && window.addEventListener) window.addEventListener('vo-look', () => { lookCache = {}; fn(); }); };
  if (typeof window !== 'undefined' && window.addEventListener) window.addEventListener('vo-look', () => { lookCache = {}; });

  /* ---------- file I/O ---------- */
  /**
   * Minimal PDF writer: one full-page JPEG per page.
   * pages: [{ jpeg: Uint8Array, iw, ih (pixels), pw, ph (points) }]
   */
  L.buildPDF = function (pages, info) {
    const enc = new TextEncoder();
    const chunks = [];
    let len = 0;
    const offs = [];
    const push = (x) => { const b = typeof x === 'string' ? enc.encode(x) : x; chunks.push(b); len += b.length; };
    const pdfStr = (t) => { let hex = 'FEFF'; for (const ch of String(t || '')) { const c = ch.codePointAt(0); if (c > 0xffff) { const v = c - 0x10000; hex += (0xd800 + (v >> 10)).toString(16).padStart(4, '0') + (0xdc00 + (v & 0x3ff)).toString(16).padStart(4, '0'); } else hex += c.toString(16).padStart(4, '0'); } return '<' + hex.toUpperCase() + '>'; };
    const num = (v) => String(Math.round(v * 100) / 100);
    push('%PDF-1.4\n');
    push(new Uint8Array([0x25, 0xe2, 0xe3, 0xcf, 0xd3, 0x0a]));
    const obj = (id, body) => { offs[id] = len; push(id + ' 0 obj\n'); if (typeof body === 'string') push(body); else body(); push('\nendobj\n'); };
    const n = pages.length;
    const infoId = 3 + n * 3;
    obj(1, '<< /Type /Catalog /Pages 2 0 R >>');
    obj(2, `<< /Type /Pages /Kids [${pages.map((_, i) => (3 + i * 3) + ' 0 R').join(' ')}] /Count ${n} >>`);
    pages.forEach((p, i) => {
      const pid = 3 + i * 3, cid = pid + 1, iid = pid + 2;
      obj(pid, `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${num(p.pw)} ${num(p.ph)}] /Resources << /XObject << /Im${i} ${iid} 0 R >> >> /Contents ${cid} 0 R >>`);
      const cs = `q ${num(p.pw)} 0 0 ${num(p.ph)} 0 0 cm /Im${i} Do Q`;
      obj(cid, `<< /Length ${cs.length} >>\nstream\n${cs}\nendstream`);
      obj(iid, () => { push(`<< /Type /XObject /Subtype /Image /Width ${p.iw} /Height ${p.ih} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${p.jpeg.length} >>\nstream\n`); push(p.jpeg); push('\nendstream'); });
    });
    const d = new Date();
    const pad = (v) => String(v).padStart(2, '0');
    obj(infoId, `<< /Title ${pdfStr(info && info.title)} /Author ${pdfStr(info && info.author)} /Producer ${pdfStr(L.APP)} /CreationDate (D:${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}) >>`);
    const xref = len;
    const total = infoId + 1;
    push(`xref\n0 ${total}\n0000000000 65535 f \n`);
    for (let i = 1; i < total; i++) push(String(offs[i]).padStart(10, '0') + ' 00000 n \n');
    push(`trailer\n<< /Size ${total} /Root 1 0 R /Info ${infoId} 0 R >>\nstartxref\n${xref}\n%%EOF\n`);
    return new Blob(chunks, { type: 'application/pdf' });
  };
  /** Offer a generated file to the user as a download. Returns 'saved', and tells vo-saved { name } on window once
   *  the file is handed over (the assistant acts out Save and Save As then, common/scout/assistant.js).
   *  In the VS Code editor the bytes go to the workspace file instead (VO.host.writeFile): 'saved' when that
   *  file was written, 'exported' when a different name or folder was written. opts.folder is the directory
   *  Save As chose. */
  L.saveFile = async function (name, blob, opts) {
    if (window.VO && VO.host && VO.host.kind === 'vscode' && VO.host.writeFile) {
      /* an in-place save of the open file is one editor save (suite.js); a copy is written here */
      let r;
      if (VO.host.saveThroughEditor) r = await VO.host.saveThroughEditor(name, blob, opts);
      if (r === undefined) r = await VO.host.writeFile(name, blob, opts);
      if (r === 'saved' || r === 'exported') window.dispatchEvent(new CustomEvent('vo-saved', { detail: { name } }));
      return r;
    }
    const url = URL.createObjectURL(blob);
    const a = L.h('a', { href: url, download: name, style: 'display:none' });
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 5000);
    window.dispatchEvent(new CustomEvent('vo-saved', { detail: { name } }));
    return 'saved';
  };
  L.pickFiles = function (accept, multiple) {
    return new Promise((resolve) => {
      const inp = L.h('input', { type: 'file', accept: accept || '', style: 'position:fixed;left:-9999px;top:0' });
      if (multiple) inp.multiple = true;
      document.body.appendChild(inp);
      let done = false;
      inp.addEventListener('change', () => { done = true; resolve(Array.from(inp.files || [])); inp.remove(); });
      window.addEventListener('focus', function f() {
        window.removeEventListener('focus', f);
        setTimeout(() => { if (!done) { resolve([]); inp.remove(); } }, 1500);
      });
      inp.click();
    });
  };
  L.readAsArrayBuffer = (blob) => blob.arrayBuffer ? blob.arrayBuffer() : new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsArrayBuffer(blob); });
  L.readAsText = (blob) => blob.text ? blob.text() : new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsText(blob); });
  L.blobToDataURL = (blob) => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsDataURL(blob); });
  L.loadImage = (src) => new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = () => rej(new Error('image load failed')); i.src = src; });
  L.extToMime = (ext) => ({
    png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', jpe: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp',
    tif: 'image/tiff', tiff: 'image/tiff', emf: 'image/x-emf', wmf: 'image/x-wmf', ico: 'image/x-icon',
  }[String(ext || '').toLowerCase()] || 'application/octet-stream');
  L.mimeToExt = (m) => ({ 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/gif': 'gif', 'image/bmp': 'bmp', 'image/svg+xml': 'svg', 'image/webp': 'webp', 'image/tiff': 'tiff', 'image/x-emf': 'emf', 'image/x-wmf': 'wmf' }[m] || 'png');

  /** pixel size and resolution of a PNG / JPEG / GIF / BMP from its header: {w, h, dpiX, dpiY} or null */
  L.imageSize = function (b) {
    try {
      const u32 = (o) => ((b[o] << 24) | (b[o + 1] << 16) | (b[o + 2] << 8) | b[o + 3]) >>> 0;
      const u16 = (o) => (b[o] << 8) | b[o + 1];
      const le16 = (o) => b[o] | (b[o + 1] << 8), le32 = (o) => (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0;
      if (b[0] === 0x89 && b[1] === 0x50) {
        const r = { w: u32(16), h: u32(20), dpiX: 96, dpiY: 96 };
        for (let o = 8; o + 12 <= b.length && o < 4096;) {
          const len = u32(o), type = String.fromCharCode(b[o + 4], b[o + 5], b[o + 6], b[o + 7]);
          if (type === 'pHYs' && b[o + 16] === 1) { r.dpiX = Math.round(u32(o + 8) * 0.0254) || 96; r.dpiY = Math.round(u32(o + 12) * 0.0254) || 96; }
          if (type === 'IDAT') break;
          o += 12 + len;
        }
        return r;
      }
      if (b[0] === 0xff && b[1] === 0xd8) {
        let dpiX = 96, dpiY = 96;
        for (let o = 2; o + 9 < b.length;) {
          if (b[o] !== 0xff) { o++; continue; }
          const m = b[o + 1], len = u16(o + 2);
          if (m === 0xe0 && b[o + 4] === 0x4a && b[o + 11] >= 1) { const k = b[o + 11] === 2 ? 2.54 : 1; dpiX = Math.round(u16(o + 12) * k) || 96; dpiY = Math.round(u16(o + 14) * k) || 96; }
          if (m >= 0xc0 && m <= 0xcf && m !== 0xc4 && m !== 0xc8 && m !== 0xcc) return { w: u16(o + 7), h: u16(o + 5), dpiX, dpiY };
          o += 2 + len;
        }
        return null;
      }
      if (b[0] === 0x47 && b[1] === 0x49) return { w: le16(6), h: le16(8), dpiX: 96, dpiY: 96 };
      if (b[0] === 0x42 && b[1] === 0x4d) { const px = le32(38), py = le32(42); return { w: le32(18), h: Math.abs(le32(22) | 0), dpiX: px ? Math.round(px * 0.0254) : 96, dpiY: py ? Math.round(py * 0.0254) : 96 }; }
    } catch (e) { /* unknown */ }
    return null;
  };

  /* ---------- media store (binary assets live outside undo snapshots) ---------- */
  const media = new Map();
  L.media = {
    /** blob: the original bytes (what gets saved); view: an optional browser-displayable rendition (e.g. PNG of a WMF) */
    add(blob, name, view) {
      const id = L.uid('m');
      media.set(id, { blob, view: view || blob, url: URL.createObjectURL(view || blob), type: blob.type || 'image/png', name: name || id });
      return id;
    },
    get: (id) => media.get(id),
    url: (id) => (media.get(id) ? media.get(id).url : ''),
    has: (id) => media.has(id),
    all: () => media,
    clear() { for (const m of media.values()) URL.revokeObjectURL(m.url); media.clear(); },
  };

  /** a copy of a picture with one colour made transparent (Set Transparent Color, a:clrChange); cached per colour */
  const clearCache = new Map();
  L.media.withTransparent = function (id, hex, tol) {
    const key = id + '|' + hex + '|' + (tol || 0);
    if (!clearCache.has(key)) clearCache.set(key, (async () => {
      try {
        const im = await L.loadImage(L.media.url(id));
        const c = document.createElement('canvas');
        c.width = im.naturalWidth || 1; c.height = im.naturalHeight || 1;
        const g = c.getContext('2d');
        g.drawImage(im, 0, 0);
        const data = g.getImageData(0, 0, c.width, c.height);
        const [tr, tg, tb] = L.color.hexToRgb(hex);
        const t = tol || 0, d = data.data;
        for (let i = 0; i < d.length; i += 4) if (Math.abs(d[i] - tr) <= t && Math.abs(d[i + 1] - tg) <= t && Math.abs(d[i + 2] - tb) <= t) d[i + 3] = 0;
        g.putImageData(data, 0, 0);
        const blob = await new Promise((res) => c.toBlob(res, 'image/png'));
        if (!blob) return null;
        const nid = L.media.add(blob, 'transparent.png');
        const o = L.media.get(id);
        if (o && o.size) L.media.get(nid).size = o.size;
        return nid;
      } catch (e) { return null; }
    })());
    return clearCache.get(key);
  };

  /* ---------- text helpers ---------- */
  L.titleCase = (s) => s.replace(/\w\S*/g, (w) => w[0].toUpperCase() + w.slice(1).toLowerCase());
  L.sentenceCase = (s) => s.toLowerCase().replace(/(^\s*\w|[.!?]\s+\w)/g, (c) => c.toUpperCase());
  L.toggleCase = (s) => s.split('').map((c) => (c === c.toUpperCase() ? c.toLowerCase() : c.toUpperCase())).join('');
  L.romanize = (n) => {
    if (n <= 0 || n > 3999) return String(n);
    const t = [['M', 1000], ['CM', 900], ['D', 500], ['CD', 400], ['C', 100], ['XC', 90], ['L', 50], ['XL', 40], ['X', 10], ['IX', 9], ['V', 5], ['IV', 4], ['I', 1]];
    let s = ''; for (const [r, v] of t) while (n >= v) { s += r; n -= v; } return s;
  };
  L.alpha = (n) => { let s = ''; n = Math.max(1, n); while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); } return s; };
  L.fmtDate = (d, fmt) => {
    const M = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
    const D = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    const p = (n) => String(n).padStart(2, '0');
    const h12 = d.getHours() % 12 || 12, ap = d.getHours() < 12 ? 'AM' : 'PM';
    switch (fmt) {
      case 'datetime2': return `${D[d.getDay()]}, ${M[d.getMonth()]} ${d.getDate()}, ${d.getFullYear()}`;
      case 'datetime3': return `${d.getDate()} ${M[d.getMonth()]} ${d.getFullYear()}`;
      case 'datetime4': return `${M[d.getMonth()]} ${d.getDate()}, ${d.getFullYear()}`;
      case 'datetime5': return `${p(d.getDate())}-${M[d.getMonth()].slice(0, 3)}-${String(d.getFullYear()).slice(2)}`;
      case 'datetime6': return `${M[d.getMonth()]} ${String(d.getFullYear()).slice(2)}`;
      case 'datetime7': return `${M[d.getMonth()].slice(0, 3)}-${String(d.getFullYear()).slice(2)}`;
      case 'datetime8': return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()} ${h12}:${p(d.getMinutes())} ${ap}`;
      case 'datetime9': return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()} ${h12}:${p(d.getMinutes())}:${p(d.getSeconds())} ${ap}`;
      case 'datetime10': return `${d.getHours()}:${p(d.getMinutes())}`;
      case 'datetime11': return `${d.getHours()}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
      case 'datetime12': return `${h12}:${p(d.getMinutes())} ${ap}`;
      case 'datetime13': return `${h12}:${p(d.getMinutes())}:${p(d.getSeconds())} ${ap}`;
      default: return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`;
    }
  };
  L.DATE_FORMATS = ['datetime1', 'datetime2', 'datetime3', 'datetime4', 'datetime5', 'datetime6', 'datetime7', 'datetime8', 'datetime9', 'datetime10', 'datetime11', 'datetime12', 'datetime13'];

  /* ---------- fonts ---------- */
  const FONT_FALLBACK = {
    'arial': '"Arimo","Liberation Sans",Helvetica,sans-serif',
    'arial black': '"Arial Black","Arimo",sans-serif',
    'arial narrow': '"Arial Narrow","Arimo",sans-serif',
    'helvetica': '"Helvetica Neue","TeX Gyre Heros","Nimbus Sans",Arial,"Arimo",sans-serif',
    'times new roman': '"Tinos","Liberation Serif",Times,serif',
    'times': '"Times New Roman","Tinos","TeX Gyre Termes",serif',
    'courier new': '"Cousine","Liberation Mono",Courier,monospace',
    'courier': '"Courier New","Cousine",monospace',
    'calibri': '"Carlito",Candara,"Segoe UI",sans-serif',
    'calibri light': '"Carlito",sans-serif',
    'aptos': 'Calibri,"Carlito",sans-serif',
    'aptos display': '"Calibri Light",Calibri,"Carlito",sans-serif',
    'aptos narrow': '"Arial Narrow",Calibri,"Carlito",sans-serif',
    'liberation sans': 'Arial,"Arimo",sans-serif',
    'liberation serif': '"Times New Roman","Tinos",serif',
    'liberation mono': '"Courier New","Cousine",monospace',
    'cambria': '"Caladea",Georgia,serif',
    'georgia': '"Gelasio","DejaVu Serif","Times New Roman","Tinos",serif',
    'garamond': '"EB Garamond","Times New Roman","Tinos",serif',
    'book antiqua': '"Palatino Linotype",Palatino,"TeX Gyre Pagella","P052","URW Palladio L","Times New Roman","Tinos",serif',
    'palatino linotype': 'Palatino,"Book Antiqua","TeX Gyre Pagella","P052","URW Palladio L","Times New Roman","Tinos",serif',
    'palatino': '"Palatino Linotype","Book Antiqua","TeX Gyre Pagella","P052","Times New Roman","Tinos",serif',
    'century schoolbook': '"New Century Schoolbook","TeX Gyre Schola","C059","Times New Roman","Tinos",serif',
    'century': '"Century Schoolbook","TeX Gyre Schola","C059","Times New Roman","Tinos",serif',
    'consolas': '"Inconsolata","DejaVu Sans Mono","Courier New","Cousine",monospace',
    'bookman old style': '"Bookman","TeX Gyre Bonum","URW Bookman","URW Bookman L","Times New Roman","Tinos",serif',
    'century gothic': '"TeX Gyre Adventor","URW Gothic","URW Gothic L","Futura",Arial,"Arimo",sans-serif',
    'tahoma': '"DejaVu Sans Condensed",Verdana,Arial,"Arimo",sans-serif',
    'verdana': '"DejaVu Sans",Tahoma,Arial,"Arimo",sans-serif',
    'trebuchet ms': '"Segoe UI",Arial,"Arimo",sans-serif',
    'comic sans ms': '"Comic Neue","Chalkboard SE",cursive',
    'impact': 'Haettenschweiler,"Arial Black",sans-serif',
    'lucida console': 'Monaco,"DejaVu Sans Mono","Courier New","Cousine",monospace',
    'segoe ui': 'Arial,"Arimo",sans-serif',
    'franklin gothic medium': '"Libre Franklin",Arial,"Arimo",sans-serif',
    'wingdings': '"Segoe UI Symbol",sans-serif',
    'symbol': '"Segoe UI Symbol",serif',
  };
  /* East Asian fonts (docs/i18n.md, Fonts): Office's names in English and as Japanese, Korean and Chinese Office
     write them, and the common open-source Chinese families, each with the fonts a Mac, Windows or Linux computer has
     in their place. None is vendored (several megabytes each, in every offline install). [English name, stack, other names] */
  const CJK_STACKS = {
    jpSerif: '"Yu Mincho","YuMincho","Hiragino Mincho ProN","Hiragino Mincho Pro","Noto Serif CJK JP","Noto Serif JP","IPAexMincho","IPAMincho",serif',
    jpSans: '"Yu Gothic","YuGothic","Hiragino Sans","Hiragino Kaku Gothic ProN","Meiryo","Noto Sans CJK JP","Noto Sans JP","IPAexGothic","IPAGothic",sans-serif',
    krSerif: '"Batang","AppleMyungjo","Noto Serif CJK KR","Noto Serif KR","UnBatang",serif',
    krSans: '"Malgun Gothic","Apple SD Gothic Neo","Noto Sans CJK KR","Noto Sans KR","UnDotum",sans-serif',
    scSerif: '"SimSun","Songti SC","STSong","Noto Serif CJK SC","Noto Serif SC","Source Han Serif SC","AR PL UMing CN",serif',
    scSans: '"Microsoft YaHei","DengXian","PingFang SC","Hiragino Sans GB","Noto Sans CJK SC","Noto Sans SC","Source Han Sans SC","WenQuanYi Micro Hei",sans-serif',
    tcSerif: '"PMingLiU","Songti TC","LiSong Pro","Noto Serif CJK TC","Noto Serif TC","AR PL UMing TW",serif',
    tcSans: '"Microsoft JhengHei","PingFang TC","Noto Sans CJK TC","Noto Sans TC",sans-serif',
    /* Source Han and Noto CJK are the same designs: each stands in for the other first */
    scOpenSerif: '"Source Han Serif SC","Noto Serif CJK SC","Noto Serif SC","SimSun","Songti SC","STSong","AR PL UMing CN",serif',
    scOpenSans: '"Source Han Sans SC","Noto Sans CJK SC","Noto Sans SC","Microsoft YaHei","DengXian","PingFang SC","Hiragino Sans GB","WenQuanYi Micro Hei",sans-serif',
  };
  const CJK = [['MS Mincho', 'jpSerif', 'ＭＳ 明朝'], ['MS PMincho', 'jpSerif', 'ＭＳ Ｐ明朝'], ['MS Gothic', 'jpSans', 'ＭＳ ゴシック'], ['MS PGothic', 'jpSans', 'ＭＳ Ｐゴシック'],
    ['MS UI Gothic', 'jpSans'], ['Yu Mincho', 'jpSerif', '游明朝'], ['Yu Gothic', 'jpSans', '游ゴシック'], ['Yu Gothic UI', 'jpSans'], ['Yu Gothic Light', 'jpSans', '游ゴシック Light'],
    ['Meiryo', 'jpSans', 'メイリオ'], ['Meiryo UI', 'jpSans'], ['HG Mincho E', 'jpSerif', 'HG明朝E'], ['HG Gothic E', 'jpSans', 'HGゴシックE'],
    ['Malgun Gothic', 'krSans', '맑은 고딕'], ['Batang', 'krSerif', '바탕'], ['BatangChe', 'krSerif', '바탕체'], ['Gulim', 'krSans', '굴림'], ['GulimChe', 'krSans', '굴림체'],
    ['Dotum', 'krSans', '돋움'], ['DotumChe', 'krSans', '돋움체'], ['Gungsuh', 'krSerif', '궁서'], ['GungsuhChe', 'krSerif', '궁서체'],
    ['SimSun', 'scSerif', '宋体'], ['NSimSun', 'scSerif', '新宋体'], ['SimHei', 'scSans', '黑体'], ['DengXian', 'scSans', '等线'], ['DengXian Light', 'scSans', '等线 Light'],
    ['Microsoft YaHei', 'scSans', '微软雅黑'], ['KaiTi', 'scSerif', '楷体'], ['FangSong', 'scSerif', '仿宋'],
    ['WenQuanYi Micro Hei', 'scSans', '文泉驿微米黑', '文泉驛微米黑'],
    ['Source Han Sans SC', 'scOpenSans', '思源黑体'], ['Source Han Serif SC', 'scOpenSerif', '思源宋体'],
    ['Noto Sans CJK SC', 'scOpenSans'], ['Noto Sans SC', 'scOpenSans'], ['Noto Serif CJK SC', 'scOpenSerif'], ['Noto Serif SC', 'scOpenSerif'],
    ['PMingLiU', 'tcSerif', '新細明體'], ['MingLiU', 'tcSerif', '細明體'], ['Microsoft JhengHei', 'tcSans', '微軟正黑體'], ['DFKai-SB', 'tcSerif', '標楷體']];
  const CJK_BY = new Map(), ALIAS = new Map(), NAMES_BY = new Map();
  for (const [en, stack, ...names] of CJK) {
    const key = en.toLowerCase();
    CJK_BY.set(key, { en, stack });
    NAMES_BY.set(key, [en, ...names]);
    for (const n of names) ALIAS.set(n.toLowerCase(), key);
  }
  const rawKey = (name) => { const n = String(name || '').replace(/"/g, '').trim().toLowerCase(); return ALIAS.get(n) || n; };
  /* Word writes several families in one name, "宋体;SimSun" or "楷体_GB2312;楷体". LibreOffice
     (GetNextFontToken, PhysicalFontCollection::FindFontFamily) splits on ';' and ',' and draws
     with the first family that is installed, else the first it can stand in for. The file keeps
     the whole list; only drawing, measuring and the Font boxes use the chosen family. */
  function probeInstalled(name) {
    try { return !!(name && L.fontInstalled(name)); } catch (e) { return false; }
  }
  function faceOf(token) {
    if (probeInstalled(token)) return token;
    const k = rawKey(token);
    for (const n of NAMES_BY.get(k) || []) if (n !== token && probeInstalled(n)) return n;
    const listed = L.localFonts().find((f) => rawKey(f) === k);
    if (listed) return listed;
    if (CJK_BY.has(k) || FONT_FALLBACK[k]) return token;
    return '';
  }
  /** the family to draw `name` with. One name is itself. A ';' or ',' list is the first family
      installed here (or the same family's other name, when that one is), else the first family
      the suite knows a stand-in for, else the first name in the list. */
  L.fontFace = function (name) {
    const s = String(name == null ? '' : name).replace(/"/g, '').trim();
    if (!s || !/[;,]/.test(s)) return s;
    const tokens = [];
    for (const part of s.split(/[;,]/)) { const t = part.trim(); if (t) tokens.push(t); }
    if (tokens.length <= 1) return tokens[0] || s;
    for (const t of tokens) { const face = faceOf(t); if (face) return face; }
    return tokens[0];
  };
  /** the key every font table uses for a font: its English name, lower case (ＭＳ 明朝 → ms mincho).
      A list of names uses the family L.fontFace picks. */
  L.fontKey = (name) => rawKey(L.fontFace(name));
  /* a family named in Chinese, Japanese or Korean is an East Asian font even when it is not one of Office's.
     Full-width Latin (Ａ) does not count. A listed local font with a Latin name is measured below. */
  const EA_NAME = /[\u1100-\u11ff\u2e80-\u9fff\uac00-\ud7af\uf900-\ufaff]/;
  let probe = null;
  const eaSeen = new Map();
  function drawsHan(family) {
    if (typeof document === 'undefined' || !document.createElement) return false;
    try {
      if (!probe) probe = document.createElement('canvas').getContext('2d');
      if (!probe) return false;
      const q = '"' + String(family).replace(/"/g, '') + '"';
      const w = (font) => { probe.font = '72px ' + font; return probe.measureText('永').width; };
      /* its own Han glyph differs from every generic fallback; a Latin font falls through and matches one */
      return ['monospace', 'serif', 'sans-serif'].every((g) => w(q + ',' + g) !== w(g));
    } catch (e) { return false; }
  }
  /** whether a font is an East Asian family: Office's list, a name written in those scripts, or a font
      this computer listed that draws Han itself */
  L.fontIsEA = (name) => {
    const k = L.fontKey(name);
    if (!k) return false;
    if (CJK_BY.has(k)) return true;
    if (EA_NAME.test(String(name || ''))) return true;
    if (eaSeen.has(k)) return eaSeen.get(k);
    const listed = L.localFonts && L.localFonts().some((f) => L.fontKey(f) === k);
    if (!listed) return false;
    const yes = drawsHan(name);
    eaSeen.set(k, yes);
    return yes;
  };
  const EA_RE = /[\u1100-\u11ff\u2e80-\u9fff\uac00-\ud7af\uf900-\ufaff\ufe30-\ufe4f\uff00-\uffef]/;
  /* Unicode East Asian Width Ambiguous (UAX #11), letters and marks left out: in a Chinese document
     Word, WPS and PowerPoint draw every one of these with the East Asian font, beside digits or not.
     The ellipsis, the middle dot and the multiplication sign are in this set; so are the degree sign,
     the dashes, the quotation marks, the roman numerals and the other ambiguous symbols. Accented
     Latin, Greek and Cyrillic stay with their own font. Private use is not included. */
  const ZH_AMBIG = [0xA1,0xA1,0xA4,0xA4,0xA7,0xA8,0xAD,0xAE,0xB0,0xB4,0xB6,0xB9,0xBC,0xBF,0xD7,0xD7,0xF7,0xF7,0x2C4,0x2C4,0x2D8,0x2DB,0x2DD,0x2DD,0x2DF,0x2DF,0x2010,0x2010,0x2013,0x2016,0x2018,0x2019,0x201C,0x201D,0x2020,0x2022,0x2024,0x2027,0x2030,0x2030,0x2032,0x2033,0x2035,0x2035,0x203B,0x203B,0x203E,0x203E,0x2074,0x2074,0x2081,0x2084,0x20AC,0x20AC,0x2103,0x2103,0x2105,0x2105,0x2109,0x2109,0x2116,0x2116,0x2121,0x2122,0x2153,0x2154,0x215B,0x215E,0x2160,0x216B,0x2170,0x2179,0x2189,0x2189,0x2190,0x2199,0x21B8,0x21B9,0x21D2,0x21D2,0x21D4,0x21D4,0x21E7,0x21E7,0x2200,0x2200,0x2202,0x2203,0x2207,0x2208,0x220B,0x220B,0x220F,0x220F,0x2211,0x2211,0x2215,0x2215,0x221A,0x221A,0x221D,0x2220,0x2223,0x2223,0x2225,0x2225,0x2227,0x222C,0x222E,0x222E,0x2234,0x2237,0x223C,0x223D,0x2248,0x2248,0x224C,0x224C,0x2252,0x2252,0x2260,0x2261,0x2264,0x2267,0x226A,0x226B,0x226E,0x226F,0x2282,0x2283,0x2286,0x2287,0x2295,0x2295,0x2299,0x2299,0x22A5,0x22A5,0x22BF,0x22BF,0x2312,0x2312,0x2460,0x24E9,0x24EB,0x254B,0x2550,0x2573,0x2580,0x258F,0x2592,0x2595,0x25A0,0x25A1,0x25A3,0x25A9,0x25B2,0x25B3,0x25B6,0x25B7,0x25BC,0x25BD,0x25C0,0x25C1,0x25C6,0x25C8,0x25CB,0x25CB,0x25CE,0x25D1,0x25E2,0x25E5,0x25EF,0x25EF,0x2605,0x2606,0x2609,0x2609,0x260E,0x260F,0x261C,0x261C,0x261E,0x261E,0x2640,0x2640,0x2642,0x2642,0x2660,0x2661,0x2663,0x2665,0x2667,0x266A,0x266C,0x266D,0x266F,0x266F,0x269E,0x269F,0x26BF,0x26BF,0x26C6,0x26CD,0x26CF,0x26D3,0x26D5,0x26E1,0x26E3,0x26E3,0x26E8,0x26E9,0x26EB,0x26F1,0x26F4,0x26F4,0x26F6,0x26F9,0x26FB,0x26FC,0x26FE,0x26FF,0x273D,0x273D,0x2776,0x277F,0x2B56,0x2B59,0x3248,0x324F,0x1F100,0x1F10A,0x1F110,0x1F12D,0x1F130,0x1F169,0x1F170,0x1F18D,0x1F18F,0x1F190,0x1F19B,0x1F1AC];
  const zhAmbig = (cp) => {
    let lo = 0, hi = ZH_AMBIG.length / 2 - 1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1, a = ZH_AMBIG[mid * 2], b = ZH_AMBIG[mid * 2 + 1];
      if (cp < a) hi = mid - 1;
      else if (cp > b) lo = mid + 1;
      else return true;
    }
    return false;
  };
  /* Common-script symbols with no width of their own take the script of the characters beside them.
     A Chinese document draws a run of only those symbols with the East Asian font. LibreOffice calls
     this class weak (i18nutil script class). */
  const weakSymbol = (cp) => (cp >= 0x2000 && cp <= 0x27FF) || (cp >= 0x2900 && cp <= 0x2BFF) || (cp >= 0x2E00 && cp <= 0x2E7F);
  const charClass = (ch, zh) => {
    if (EA_RE.test(ch)) return 'ea';
    const cp = ch.codePointAt(0);
    if (zh && zhAmbig(cp)) return 'ea';
    if (weakSymbol(cp)) return 'weak';
    return 'other';
  };
  /** true where each character of s is drawn with the East Asian font. A weak symbol uses the strong
      character before it, or the first strong character when it leads the text. */
  function eaMask(s, zh) {
    const chars = [...String(s || '')];
    const kinds = chars.map((ch) => charClass(ch, zh));
    let first = null;
    for (const k of kinds) if (k !== 'weak') { first = k; break; }
    /* no Han or Latin letter to follow: a Chinese document uses the East Asian font */
    if (first == null) first = zh && kinds.some((k) => k === 'weak') ? 'ea' : 'other';
    let prev = first;
    return kinds.map((k) => {
      if (k === 'weak') return prev === 'ea';
      prev = k;
      return k === 'ea';
    });
  }
  /** zh, zh-CN, zh-Hans, zh-Hant, zh-TW and the other Chinese tags */
  L.chineseLang = (lang) => /^zh(?:-|$)/i.test(String(lang || ''));
  /** a Chinese document: Word's East Asian language (theme or the default run), or a presentation's text language */
  L.chineseDoc = function (doc) {
    if (!doc) return false;
    const theme = doc.settings && doc.settings.themeLang;
    if (theme && L.chineseLang(theme.eastAsia)) return true;
    const def = doc.defaults && doc.defaults.rPr;
    if (def && L.chineseLang(def.langEA)) return true;
    return L.chineseLang(doc.lang);
  };
  /** whether text has characters Word draws with a run's East Asian font. Pass true for a Chinese document:
      every East Asian Width Ambiguous symbol is one of them, as is punctuation that sits with East Asian characters. */
  L.hasEAText = (s, zh) => eaMask(s, zh).some(Boolean);
  /** the Latin slot names the same East Asian family as the East Asian slot: Word's Font box shows
      "(Use Asian text font)" rather than that family again */
  L.fontUsesEA = (font, ea) => !!(font && ea && L.fontIsEA(ea) && L.fontKey(font) === L.fontKey(ea));
  /** props the toolbar Font box applies to one run (resolved fonts in, direct props out).
      An East Asian family is written to both slots: it draws Chinese and, like 宋体, Latin too.
      A Latin family is the Latin slot only, and the East Asian face stays when that face was
      the Latin slot or the same family. */
  L.fontPick = function (name, resolved) {
    resolved = resolved || {};
    if (!name) return {};
    if (L.fontIsEA(name)) return { font: name, fontEA: name };
    const out = { font: name };
    const ea = resolved.fontEA, latin = resolved.font || '';
    const separate = !!(ea && L.fontIsEA(ea) && L.fontKey(ea) !== L.fontKey(latin));
    if (!separate) {
      const face = ea && L.fontIsEA(ea) ? ea : (L.fontIsEA(latin) ? latin : '');
      if (face && L.fontKey(face) !== L.fontKey(name)) out.fontEA = face;
    }
    return out;
  };
  /** which face a cell uses for Latin and for East Asian text. styleName is the resolved cell font;
      runs are rich text ({t, font:{name}}) or absent, when the cell font draws both. */
  L.scriptSlots = function (styleName, runs) {
    let latin = styleName || '';
    let ea = '';
    if (runs && runs.length) {
      for (const run of runs) {
        const name = (run.font && run.font.name) || latin;
        if (L.hasEAText(run.t)) { if (L.fontIsEA(name)) ea = name; else if (!ea && L.fontIsEA(latin)) ea = latin; }
        else if (run.t) latin = name;
      }
    }
    if (!ea && L.fontIsEA(latin)) ea = latin;
    return { font: latin, fontEA: ea };
  };
  /** rich-text runs after a toolbar font pick, or null when the cell font alone is enough.
      styleFont supplies size and weight; pick is L.fontPick's result. A single East Asian run is kept
      when its face is not the Latin face, so Chinese-only text does not follow a Latin font. */
  L.scriptFontRuns = function (text, runs, styleFont, pick) {
    text = text == null ? '' : String(text);
    const latin = pick && pick.font;
    if (!text || !latin) return null;
    const ea = pick.fontEA || '';
    const base = styleFont || {};
    const joined = runs && runs.length ? runs.map((r) => r.t || '').join('') : null;
    const src = joined === text ? runs : [{ t: text, font: base }];
    const copy = (from, name) => {
      const srcF = Object.assign({}, base, from || {});
      const f = { name };
      for (const k of ['b', 'i', 'u', 'strike', 'sz', 'color', 'vert']) if (srcF[k] != null) f[k] = srcF[k];
      return f;
    };
    const out = [];
    for (const run of src) {
      for (const piece of L.splitEA(run.t || '')) {
        const name = piece.ea ? (ea || (run.font && run.font.name) || latin) : latin;
        const font = copy(run.font, name);
        const prev = out[out.length - 1];
        if (prev && JSON.stringify(prev.font) === JSON.stringify(font)) prev.t += piece.text;
        else out.push({ t: piece.text, font });
      }
    }
    if (!out.length) return null;
    if (out.every((r) => L.fontKey(r.font.name) === L.fontKey(latin))) return null;
    return out;
  };
  /** text in order, each piece { ea, text }: East Asian characters, then the rest, as Word picks a run's font.
      zh: a Chinese document, where every East Asian Width Ambiguous symbol is East Asian. Other
      general punctuation takes the script of the characters beside it. */
  L.splitEA = function (s, zh) {
    const chars = [...String(s || '')];
    const mask = eaMask(chars.join(''), zh);
    const out = [];
    let buf = '', ea = false, on = false;
    for (let i = 0; i < chars.length; i++) {
      const e = mask[i];
      if (!on) { on = true; ea = e; buf = chars[i]; }
      else if (e === ea) buf += chars[i];
      else { out.push({ ea, text: buf }); ea = e; buf = chars[i]; }
    }
    if (on) out.push({ ea, text: buf });
    return out;
  };
  /** a CSS font-family list for a font and what stands in for it; with ea, an East Asian font, its letters come
      from the first and its East Asian characters from the second (as Word draws a run with both fonts) */
  L.fontStack = function (name, ea) {
    if (!name) return ea ? L.fontStack(ea) : 'Arial,"Arimo",sans-serif';
    name = L.fontFace(name);
    if (ea) ea = L.fontFace(ea);
    const n = String(name).replace(/"/g, ''), k = L.fontKey(n), cjk = CJK_BY.get(k);
    const stack = cjk ? `"${n}",` + (cjk.en !== n ? `"${cjk.en}",` : '') + CJK_STACKS[cjk.stack] : `"${n}",${FONT_FALLBACK[k] || 'Arial,"Arimo",sans-serif'}`;
    if (!ea || L.fontKey(ea) === k || !L.fontIsEA(ea)) return stack;
    return stack.replace(/,(serif|sans-serif|monospace|cursive)$/, '') + ',' + L.fontStack(ea);
  };
  /* Which family in a stack the browser will draw with: an installed font (detected by measuring
     text against the generic fallbacks), or one of the suite's own stand-ins in public/fonts/. */
  const installed = new Map();
  L.fontInstalled = function (family) {
    if (installed.has(family)) return installed.get(family);
    if (typeof document === 'undefined' || !document.createElement) return false;
    if (!probe) probe = document.createElement('canvas').getContext('2d');
    if (!probe) return false;
    const text = 'mmmmmmmmmmlli1WQ@ ', width = (font) => { probe.font = '72px ' + font; return probe.measureText(text).width; };
    const yes = ['monospace', 'serif', 'sans-serif'].some((g) => width(`"${family}",${g}`) !== width(g));
    installed.set(family, yes);
    return yes;
  };
  const standIns = () => new Set(Array.from(document.fonts || [], (f) => f.family.replace(/["']/g, '')));
  /** the family that draws `name` (an Office font name): itself when installed, else its stand-in */
  L.fontUsed = function (name) {
    const own = standIns();
    for (const f of L.fontStack(name).split(',').map((s) => s.trim().replace(/^"|"$/g, ''))) {
      if (['serif', 'sans-serif', 'monospace', 'cursive', 'fantasy'].includes(f)) return f;
      if (own.has(f) || L.fontInstalled(f)) return f;
    }
    return 'sans-serif';
  };
  const GENERIC = ['serif', 'sans-serif', 'monospace', 'cursive', 'fantasy'];
  /** what draws `name` on this computer when it isn't installed: a family, a generic one (serif…), or '' */
  L.fontStandIn = function (name) {
    name = L.fontFace(name);
    if (L.localFonts().includes(name)) return '';   // listed by the browser as installed here
    const used = L.fontUsed(name);
    return used.toLowerCase() === String(name).toLowerCase() ? '' : used;
  };
  /** the Font boxes' grey note beside a font name, and its tooltip: what draws it on this computer, if not itself
      ({ text, tip }, or '' when it is installed) */
  L.fontNote = function (name) {
    const used = L.fontStandIn(name);
    if (!used) return '';
    const generic = GENERIC.includes(used), t = tr;   // t: UI text (tools/i18n.mjs)
    return {
      text: generic ? t('Not installed') : t('Not installed; uses {0}', used),
      tip: generic ? t('{0} is not installed on this computer, so its text is shown in a default font. The document keeps the font name {0} and shows in that font on computers that have it.', name)
        : t('{0} is not installed on this computer, so its text is shown in {1}. The document keeps the font name {0} and shows in that font on computers that have it.', name, used),
    };
  };
  /** @font-face rules for a standalone export (an HTML page): the stand-ins from public/fonts/ that
   * `markup` names, as data: URLs, with bold/italic faces only if the markup uses them. */
  L.embedFonts = async function (markup) {
    const base = (root.VO?.base || new URL('..', location.href).href) + 'fonts/';
    const css = await (await fetch(base + 'fonts.css')).text();
    const bold = /bold|font-weight[:=]\s*"?[6-9]00/i.test(markup), italic = /italic/i.test(markup);
    const files = new Map();   // one rule per file: a variable font's regular and bold share it
    for (const rule of css.match(/@font-face\s*\{[^}]*\}/g) || []) {
      const family = /font-family:\s*'([^']+)'/.exec(rule)[1], file = /url\('([^']+)'\)/.exec(rule)[1];
      const style = /font-style:\s*(\w+)/.exec(rule)[1], weight = +/font-weight:\s*(\d+)/.exec(rule)[1];
      if (!markup.includes(family) || (!bold && weight > 400) || (!italic && style === 'italic')) continue;
      if (!files.has(file)) files.set(file, { family, style, weights: [] });
      files.get(file).weights.push(weight);
    }
    const rules = [];
    for (const [file, f] of files) {
      const bytes = new Uint8Array(await (await fetch(base + file)).arrayBuffer());
      let bin = ''; for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
      const weight = f.weights.length > 1 ? Math.min(...f.weights) + ' ' + Math.max(...f.weights) : f.weights[0];
      rules.push(`@font-face{font-family:'${f.family}';font-style:${f.style};font-weight:${weight};src:url(data:font/woff2;base64,${btoa(bin)}) format('woff2')}`);
    }
    return rules.join('\n');
  };
  /* the Font boxes: the fonts Office users pick, by Office's names, in every UI language (documents add their own).
     Japanese, Korean and Chinese families are always listed, including the open-source Chinese ones under their
     installed names (sources and English aliases: docs/i18n.md, Fonts). */
  L.FONT_LIST = Array.from(new Set(['Aptos', 'Aptos Display', 'Aptos Narrow', 'Arial', 'Arial Black', 'Calibri', 'Calibri Light', 'Cambria', 'Century Gothic', 'Comic Sans MS', 'Consolas',
    'Courier New', 'Garamond', 'Georgia', 'Impact', 'Palatino Linotype', 'Segoe UI', 'Symbol', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana', 'Wingdings',
    '游明朝', '游ゴシック', 'ＭＳ 明朝', 'ＭＳ ゴシック', 'ＭＳ Ｐ明朝', 'ＭＳ Ｐゴシック', 'メイリオ',
    '맑은 고딕', '바탕', '바탕체', '굴림', '굴림체', '돋움', '돋움체', '궁서',
    '宋体', '新宋体', '黑体', '楷体', '仿宋', '等线', '微软雅黑', '文泉驿微米黑', '思源黑体', '思源宋体',
    'Noto Sans CJK SC', 'Noto Sans SC', 'Noto Serif CJK SC', 'Noto Serif SC',
    '新細明體', '細明體', '標楷體', '微軟正黑體']));
  /** the Font box list in Office's sections: Theme Fonts (the document's heading and body fonts),
   * Recently Used Fonts, All Fonts (L.FONT_LIST plus fonts the document uses), then the computer's
   * fonts once loaded (Chrome and Edge: the first item loads them). Items are strings,
   * { heading } rows or { value, label }. */
  L.fontSections = function ({ major, minor, extra = [] } = {}) {
    const face = (n) => L.fontFace(n);
    const out = [], recent = L.store.get('recentFonts', []).slice(0, 5).map(face).filter(Boolean), local = L.localFonts();
    // Chrome and Edge can list the computer's fonts (Local Font Access, after the user allows it)
    if (typeof root.queryLocalFonts === 'function') out.push({ label: local.length ? 'Reload fonts on your computer...' : 'Load all fonts on your computer...', run: L.loadLocalFonts }, { sep: true });
    if (major || minor) {
      out.push({ heading: 'Theme Fonts' });
      if (major) { const f = face(major); out.push({ value: f, label: tr('{0} (Headings)', f) }); }
      if (minor) { const f = face(minor); out.push({ value: f, label: tr('{0} (Body)', f) }); }
    }
    if (recent.length) out.push({ heading: 'Recently Used Fonts' }, ...recent);
    const all = Array.from(new Set(L.FONT_LIST.concat(extra.filter(Boolean).map(face)))).sort((x, y) => x.localeCompare(y));
    out.push({ heading: 'All Fonts' }, ...all);
    const more = local.filter((f) => !all.includes(f));
    if (more.length) out.push({ heading: 'Fonts on This Computer' }, ...more);
    return out;
  };
  /* the families queryLocalFonts() found, kept for every app of the suite (VO.localFonts, common/suite.js;
     the Start Center's Settings can load or forget them) */
  L.localFonts = () => (root.VO ? root.VO.localFonts.list() : []);
  L.loadLocalFonts = async function () {
    try {
      return await root.VO.localFonts.load();
    } catch (e) {
      L.ui?.msg(e.name === 'NotAllowedError' || e.name === 'SecurityError' ? 'The fonts on this computer could not be listed: permission was not given. You can allow it in the site settings of your browser.' : L.t('The fonts on this computer could not be listed: {0}', e.message || e), { icon: 'warn' });
      return 0;
    }
  };
  /** remember a font picked in a Font box (for Recently Used Fonts) */
  L.addRecentFont = function (name) {
    if (name) L.store.set('recentFonts', [name].concat(L.store.get('recentFonts', []).filter((f) => f !== name)).slice(0, 10));
  };
  /** the option a typed prefix completes to: the first name that starts with it, after `order` (if given) sorts the list */
  L.completeMatch = function (names, prefix, order) {
    const p = String(prefix || '');
    if (!p) return '';
    const list = typeof order === 'function' ? (order(names.slice(), p) || names) : names;
    const low = p.toLowerCase();
    for (const n of list) if (String(n).toLowerCase().startsWith(low)) return String(n);
    return '';
  };
  /** Font-box completion order: a recently used family before the others, then alphabetical */
  L.fontComplete = function (names) {
    const recent = L.store.get('recentFonts', []);
    const rank = (n) => { const k = L.fontKey(n); const i = recent.findIndex((f) => L.fontKey(f) === k); return i < 0 ? recent.length + 1 : i; };
    return names.slice().sort((a, b) => rank(a) - rank(b) || String(a).localeCompare(String(b)));
  };
  /* the Font Size box of Word and Excel; Lectern sets PowerPoint's list in its app.js */
  L.SIZE_LIST = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

  /* Wingdings / Symbol bullet glyphs mapped to Unicode so imported documents look right */
  L.mapSymbolChar = function (ch, font) {
    const f = String(font || '').toLowerCase();
    const code = ch ? ch.charCodeAt(0) & 0xff : 0;
    /* Wingdings 2 and 3 carry the round, square and triangular bullets of many Office themes */
    if (/^wingdings\s*2/.test(f)) {
      if (code >= 0x96 && code <= 0x9d) return '●';
      if (code >= 0x9e && code <= 0xa5) return '■';
      if (code >= 0xa6 && code <= 0xad) return '◆';
      return { 0x50: '✓', 0x4f: '✗', 0x51: '☒', 0x52: '☑', 0xe9: '★', 0xea: '★' }[code] || '•';
    }
    if (/^wingdings\s*3/.test(f)) return { 0x7d: '►', 0x84: '►', 0x75: '►', 0x71: '◄', 0x70: '▲', 0x72: '▼', 0xc6: '➔', 0xc7: '⬅' }[code] || '▸';
    if (f.startsWith('wingdings')) {
      const m = { 'l': '●', 'n': '■', 'q': '❑', 'u': '◆', 'v': '❖', 'Ø': '➢', 'ü': '✓', '§': '▪', 'o': '❍', 'p': '□', 'w': '⬥', 'Ü': '➢', 'à': '➔', 'è': '➜', 'Þ': '➝', 'ð': '⇨', 'ª': '✦', '¨': '◻', '\uF06E': '■', '\uF06C': '●', '\uF075': '◆', '\uF0FC': '✓', '\uF0D8': '➢', '\uF0A7': '▪' };
      return m[ch] || '•';
    }
    if (f === 'symbol') { const m = { '·': '•', '\uF0B7': '•', 'Þ': '⇒', '®': '→', '¾': '—' }; return m[ch] || ch; }
    if (ch === '' || ch === '') return '•';
    if (ch && ch.charCodeAt(0) >= 0xF000 && ch.charCodeAt(0) <= 0xF0FF) return '•';
    return ch;
  };

  /* ---------- geometry helpers ---------- */
  L.rotPt = (x, y, cx, cy, deg) => {
    const a = (deg * Math.PI) / 180, c = Math.cos(a), s = Math.sin(a);
    const dx = x - cx, dy = y - cy;
    return [cx + dx * c - dy * s, cy + dx * s + dy * c];
  };
  /* Axis-aligned bounds of a rotated rectangle */
  L.rotBounds = (sh) => {
    if (!sh.rot) return { x: sh.x, y: sh.y, w: sh.w, h: sh.h };
    const cx = sh.x + sh.w / 2, cy = sh.y + sh.h / 2;
    const pts = [[sh.x, sh.y], [sh.x + sh.w, sh.y], [sh.x + sh.w, sh.y + sh.h], [sh.x, sh.y + sh.h]].map(([x, y]) => L.rotPt(x, y, cx, cy, sh.rot));
    const xs = pts.map((p) => p[0]), ys = pts.map((p) => p[1]);
    const x0 = Math.min(...xs), y0 = Math.min(...ys);
    return { x: x0, y: y0, w: Math.max(...xs) - x0, h: Math.max(...ys) - y0 };
  };
  L.unionBounds = (list) => {
    if (!list.length) return { x: 0, y: 0, w: 0, h: 0 };
    let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
    for (const b of list) { x0 = Math.min(x0, b.x); y0 = Math.min(y0, b.y); x1 = Math.max(x1, b.x + b.w); y1 = Math.max(y1, b.y + b.h); }
    return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
  };
})(typeof window !== 'undefined' ? window : globalThis);
