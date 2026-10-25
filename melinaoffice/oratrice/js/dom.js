/* Oratrice — PowerPoint object model on the open presentation.
   Inventory: docs/oratrice_dom_api_specifications.md */
(function () {
  'use strict';
  const L = window.L;
  const M = L.model;
  const T = L.txt;
  const dom = (L.dom = L.dom || {});

  function gap(code, member) {
    const e = new Error(member ? member + ' is not available (' + code + ')' : code);
    e.code = code;
    throw e;
  }

  function missing() {
    const e = new Error('The requested member of the collection does not exist.');
    e.code = 'missing';
    throw e;
  }

  function on(v) { return v === true || v === -1; }
  function tri(v) { return v ? -1 : 0; }
  function edit(label, fn) {
    if (L.hist && L.hist.push) L.hist.push(label);
    const out = fn();
    try { if (L.ed && L.ed.render) L.ed.render(); } catch (e) { /* the page is not mounted */ }
    return out;
  }

  const Application = {};
  function stamp(obj, parent) {
    obj.Application = Application;
    if (typeof parent === 'function') {
      Object.defineProperty(obj, 'Parent', { enumerable: true, configurable: true, get: parent });
    } else {
      obj.Parent = parent == null ? Application : parent;
    }
    Object.defineProperty(obj, 'Creator', {
      enumerable: true, configurable: true,
      get() { gap('G-ENGINE', 'Creator'); },
    });
    return obj;
  }

  function denyGet(obj, names, code) {
    for (const n of names) {
      Object.defineProperty(obj, n, {
        enumerable: true, configurable: true,
        get() { gap(code, n); },
        set() { gap(code, n); },
      });
    }
  }
  function denyFn(obj, names, code) {
    for (const n of names) obj[n] = function () { gap(code, n); };
  }

  function bgrToHex(n) {
    n = n >>> 0;
    const r = n & 255, g = (n >> 8) & 255, b = (n >> 16) & 255;
    return '#' + [r, g, b].map((x) => x.toString(16).padStart(2, '0')).join('').toUpperCase();
  }
  function hexToBgr(hex) {
    const h = String(hex).replace('#', '');
    if (h.length < 6) return 0;
    return parseInt(h.slice(0, 2), 16) | (parseInt(h.slice(2, 4), 16) << 8) | (parseInt(h.slice(4, 6), 16) << 16);
  }

  const THEME = {
    dk1: 1, lt1: 2, dk2: 3, lt2: 4,
    accent1: 5, accent2: 6, accent3: 7, accent4: 8, accent5: 9, accent6: 10,
    hlink: 11, folHlink: 12, tx1: 13, bg1: 14, tx2: 15, bg2: 16,
  };
  const THEME_NAME = {};
  for (const k in THEME) THEME_NAME[THEME[k]] = k;
  const SCHEME = { lt1: 1, dk1: 2, lt2: 3, dk2: 4, accent1: 5, accent2: 6, hlink: 7, folHlink: 8 };
  const SCHEME_NAME = {};
  for (const k in SCHEME) SCHEME_NAME[SCHEME[k]] = k;

  const GEOM = {
    1: 'rect', 4: 'diamond', 5: 'roundRect', 9: 'ellipse', 10: 'hexagon', 11: 'plus',
    13: 'can', 14: 'cube', 18: 'donut', 20: 'blockArc', 21: 'heart', 22: 'lightningBolt',
    23: 'sun', 24: 'moon', 25: 'arc',
  };
  const GEOM_CODE = {};
  for (const k in GEOM) GEOM_CODE[GEOM[k]] = +k;

  const LAYOUT = {
    1: 'title', 2: 'text', 3: 'twoText', 4: 'table', 8: 'chart', 11: 'titleOnly', 12: 'blank',
    13: 'textContent', 14: 'contentText', 15: 'contentOnly', 16: 'content',
    19: 'contentOverText', 20: 'textOverContent', 21: 'textTwoContent', 22: 'twoContentText',
    24: 'fourContent', 25: 'vertText', 27: 'vertTitleText', 29: 'twoContent',
    30: 'contentTwoContent', 31: 'twoContentContent',
  };
  const LAYOUT_CODE = {};
  for (const k in LAYOUT) LAYOUT_CODE[LAYOUT[k]] = +k;

  const PH = { title: 1, body: 2, ctrTitle: 3, subTitle: 4, obj: 7, chart: 8, tbl: 12, sldNum: 13, ftr: 15, dt: 16 };
  const ALIGN_TO = { l: 1, left: 1, ctr: 2, center: 2, r: 3, right: 3, just: 4, justify: 4, dist: 5, distribute: 5 };
  const ALIGN_FROM = { 1: 'l', 2: 'ctr', 3: 'r', 4: 'just', 5: 'dist' };
  const VANCH_TO = { t: 1, ctr: 3, b: 4 };
  const VANCH_FROM = { 1: 't', 3: 'ctr', 4: 'b' };

  const EFFECTS = [
    [0, { type: 'none' }],
    [257, { type: 'cut' }],
    [258, { type: 'cut', thruBlk: true }],
    [1537, { type: 'dissolve' }],
    [2817, { type: 'wipe', dir: 'l' }],
    [2818, { type: 'wipe', dir: 'u' }],
    [2819, { type: 'wipe', dir: 'r' }],
    [2820, { type: 'wipe', dir: 'd' }],
    [3849, { type: 'fade' }],
    [3852, { type: 'push', dir: 'd' }],
    [3853, { type: 'push', dir: 'l' }],
    [3854, { type: 'push', dir: 'r' }],
    [3855, { type: 'push', dir: 'u' }],
  ];
  const SPEED_TO = { slow: 1, med: 2, fast: 3 };
  const SPEED_FROM = { 1: 'slow', 2: 'med', 3: 'fast' };

  let wraps = [];
  let multi = false;
  const state = { slide: 0, shapes: [], text: null, slideSel: null };
  let clip = null;

  function fileName() {
    return (L.app && L.app.fileName) || 'Presentation1';
  }

  function ensure() {
    if (!L.pres) L.pres = M.newPresentation();
    let w = wraps.find((x) => x.pres === L.pres);
    if (!w) {
      if (wraps.length === 1 && !multi) {
        wraps[0].pres = L.pres;
        wraps[0].name = fileName();
        w = wraps[0];
      } else {
        w = { pres: L.pres, name: fileName() };
        wraps.push(w);
      }
    }
    return w;
  }

  function pres() { return ensure().pres; }

  function designOf(slide) { return M.design(pres(), slide); }

  function clampSlide() {
    const n = pres().slides.length;
    if (!n) return 0;
    if (L.ed && L.ed.view !== 'master' && L.ed.view !== 'sorter' && typeof L.ed.idx === 'number') state.slide = L.ed.idx;
    if (state.slide < 0 || state.slide >= n) state.slide = Math.max(0, Math.min(state.slide, n - 1));
    return state.slide;
  }

  function slideByIndex(index) {
    const i = (index | 0) - 1;
    const slides = pres().slides;
    if (i < 0 || i >= slides.length) missing();
    return slides[i];
  }

  function focusSlide(i) {
    state.slide = i;
    state.shapes = [];
    state.text = null;
    state.slideSel = [i];
    if (L.ed) {
      L.ed.idx = i;
      if (typeof L.ed.select === 'function') L.ed.select([]);
    }
    if (L.app && L.app.slideSel) L.app.slideSel = new Set([i]);
  }

  function selectedShapeIds() {
    if (state.text) return [state.text.shapeId];
    if (L.ed && Array.isArray(L.ed.sel) && L.ed.view !== 'sorter' && L.app && L.app.focusArea !== 'slides') return L.ed.sel.filter(Boolean);
    return state.shapes.slice();
  }

  function selectedSlideIndexes() {
    if (L.app && (L.app.view === 'sorter' || L.app.focusArea === 'slides') && L.app.slideSel && L.app.slideSel.size) {
      return Array.from(L.app.slideSel).filter((i) => i >= 0 && i < pres().slides.length).sort((a, b) => a - b);
    }
    if (state.slideSel && state.slideSel.length) return state.slideSel.filter((i) => i >= 0 && i < pres().slides.length);
    return [];
  }

  function shapeById(slide, id) {
    const f = M.find(slide, id);
    return f ? f.shape : null;
  }

  function crText(s) { return String(s == null ? '' : s).replace(/\n/g, '\r'); }
  function normText(s) { return String(s == null ? '' : s).replace(/\r\n/g, '\n').replace(/\r/g, '\n'); }

  function whole(tx) {
    const last = Math.max(0, tx.ps.length - 1);
    return { p0: 0, o0: 0, p1: last, o1: T.paraText(tx.ps[last] || { rs: [] }).length };
  }

  function spanText(tx, a) {
    if (!tx || !tx.ps || !tx.ps.length) return '';
    if (a.p0 === a.p1) return T.paraText(tx.ps[a.p0] || { rs: [] }).slice(a.o0, a.o1);
    const parts = [T.paraText(tx.ps[a.p0] || { rs: [] }).slice(a.o0)];
    for (let i = a.p0 + 1; i < a.p1; i++) parts.push(T.paraText(tx.ps[i] || { rs: [] }));
    parts.push(T.paraText(tx.ps[a.p1] || { rs: [] }).slice(0, a.o1));
    return parts.join('\r');
  }

  function charCount(tx, a) {
    return spanText(tx, a).replace(/\r/g, '').length;
  }

  function writeSpan(tx, a, raw) {
    const all = whole(tx);
    const full = a.p0 === all.p0 && a.o0 === all.o0 && a.p1 === all.p1 && a.o1 === all.o1;
    if (full) T.setPlain(tx, normText(raw));
    else {
      const pos = T.deleteRange(tx, a);
      T.insertAt(tx, pos, normText(raw));
    }
  }

  function parasTouched(tx, a) {
    const out = [];
    for (let i = a.p0; i <= a.p1 && tx.ps[i]; i++) out.push(tx.ps[i]);
    return out;
  }

  function runsTouched(tx, a) {
    const out = [];
    parasTouched(tx, a).forEach((p, k) => {
      const pi = a.p0 + k;
      const len = T.paraText(p).length;
      const from = pi === a.p0 ? a.o0 : 0;
      const to = pi === a.p1 ? a.o1 : len;
      if (!p.rs.length) { if (p.end) out.push(p.end); return; }
      let o = 0;
      for (const r of p.rs) {
        const end = o + r.t.length;
        if (end > from && o < to) out.push(r);
        o = end;
      }
    });
    return out;
  }

  function effectivePara(shape, p) {
    try { return L.style.para(shape, p, designOf(slideHolding(shape))); }
    catch (e) { return (p && p.pp) || {}; }
  }

  function slideHolding(shape) {
    const slides = pres().slides;
    for (const s of slides) {
      if (M.find(s, shape.id)) return s;
    }
    return slides[clampSlide()] || slides[0];
  }

  function effectiveRun(shape, p, r) {
    const ps = effectivePara(shape, p);
    const run = r || (p.rs && p.rs[0]) || p.end || {};
    return L.style.run(ps, run);
  }

  function colorFormat(parent, get, set) {
    const cf = stamp({}, parent);
    function stored() { return get() == null ? '' : String(get()); }
    Object.defineProperty(cf, 'RGB', {
      enumerable: true,
      get() {
        const c = stored();
        const hex = c.charAt(0) === '#' ? c : M.resolveColor(c || 'tx1', designOf(pres().slides[clampSlide()] || { design: Object.keys(pres().designs)[0] }));
        return hexToBgr(hex);
      },
      set(v) { set(bgrToHex(v)); },
    });
    Object.defineProperty(cf, 'ObjectThemeColor', {
      enumerable: true,
      get() { return THEME[stored()] || 0; },
      set(v) {
        const name = THEME_NAME[v | 0];
        if (!name) gap('G-ENGINE', 'ObjectThemeColor');
        set(name);
      },
    });
    Object.defineProperty(cf, 'SchemeColor', {
      enumerable: true,
      get() { return SCHEME[stored()] || 0; },
      set(v) {
        if ((v | 0) === 0) return;
        const name = SCHEME_NAME[v | 0];
        if (!name) gap('G-ENGINE', 'SchemeColor');
        set(name);
      },
    });
    denyGet(cf, ['TintAndShade', 'Brightness'], 'G-ENGINE');
    return cf;
  }

  function collection(parent, getList) {
    const col = function (index) { return col.Item(index); };
    stamp(col, parent);
    col.Item = function (index) {
      const list = getList();
      if (typeof index === 'string' && index !== '' && !/^\d+$/.test(index)) {
        const hit = list.find((x) => x && x._name && x._name() === index);
        if (!hit) missing();
        return hit;
      }
      const i = index | 0;
      if (i < 1 || i > list.length) missing();
      return list[i - 1];
    };
    Object.defineProperty(col, 'Count', { enumerable: true, configurable: true, get() { return getList().length; } });
    Object.defineProperty(col, 'First', { enumerable: true, configurable: true, get() { const list = getList(); return list.length ? list[0] : null; } });
    Object.defineProperty(col, 'Last', { enumerable: true, configurable: true, get() { const list = getList(); return list.length ? list[list.length - 1] : null; } });
    return col;
  }

  function charactersOf(parent, get) {
    const col = collection(parent, () => []);
    col.Item = function () { gap('G-TEXT', 'Characters'); };
    Object.defineProperty(col, 'Count', { enumerable: true, get: get });
    Object.defineProperty(col, 'First', { enumerable: true, get() { gap('G-TEXT', 'Characters'); } });
    Object.defineProperty(col, 'Last', { enumerable: true, get() { gap('G-TEXT', 'Characters'); } });
    return col;
  }

  function makeFont(parent, shape, tx, anchor) {
    const font = stamp({}, parent);
    function runs() { return runsTouched(tx, anchor()); }
    function read(key, fallback) {
      const list = runs();
      if (!list.length) {
        const p = tx.ps[anchor().p0];
        const eff = p ? effectiveRun(shape, p, p.end) : {};
        return eff[key] != null ? eff[key] : fallback;
      }
      const vals = list.map((r) => {
        const p = tx.ps.find((para) => para.rs.includes(r) || para.end === r) || tx.ps[anchor().p0];
        return effectiveRun(shape, p, r)[key];
      });
      return vals[0] != null ? vals[0] : fallback;
    }
    function write(key, value) {
      edit('Font', () => {
        const a = anchor();
        const empty = parasTouched(tx, a).every((p) => !T.paraText(p).length);
        if (empty) {
          parasTouched(tx, a).forEach((p) => { p.end = Object.assign({}, p.end, { [key]: value }); });
          return;
        }
        T.mapRange(tx, a, (r) => Object.assign({}, r, { [key]: value }));
      });
    }
    Object.defineProperty(font, 'Name', {
      enumerable: true,
      get() { return L.style.font(read('font', '+mn'), designOf(slideHolding(shape))); },
      set(v) { write('font', String(v)); },
    });
    Object.defineProperty(font, 'NameFarEast', {
      enumerable: true,
      get() { return L.style.fontEA(read('fontEA', ''), designOf(slideHolding(shape))) || ''; },
      set(v) { write('fontEA', String(v)); },
    });
    Object.defineProperty(font, 'Size', {
      enumerable: true,
      get() { return read('sz', 18) || 18; },
      set(v) { write('sz', +v); },
    });
    function flag(key) {
      return {
        enumerable: true,
        get() {
          const list = runs();
          if (!list.length) return tri(read(key, false));
          const vals = list.map((r) => !!r[key]);
          if (vals.every(Boolean)) return -1;
          if (vals.every((x) => !x)) return 0;
          return -2;
        },
        set(v) { write(key, on(v)); },
      };
    }
    Object.defineProperty(font, 'Bold', flag('b'));
    Object.defineProperty(font, 'Italic', flag('i'));
    Object.defineProperty(font, 'Underline', {
      enumerable: true,
      get() {
        const list = runs();
        const u = (r) => r.u && r.u !== 'none';
        if (!list.length) return tri(u({ u: read('u', false) }));
        const vals = list.map(u);
        if (vals.every(Boolean)) return -1;
        if (vals.every((x) => !x)) return 0;
        return -2;
      },
      set(v) { write('u', on(v) ? 'sng' : 'none'); },
    });
    Object.defineProperty(font, 'Color', {
      enumerable: true,
      get() {
        return colorFormat(font, () => read('color', 'tx1'), (hex) => write('color', hex));
      },
    });
    denyGet(font, ['NameAscii', 'NameComplexScript', 'SizeBi', 'BoldBi', 'ItalicBi'], 'G-ENGINE');
    return font;
  }

  function makeParagraphFormat(parent, shape, tx, anchor) {
    const fmt = stamp({}, parent);
    function paras() { return parasTouched(tx, anchor()); }
    Object.defineProperty(fmt, 'Alignment', {
      enumerable: true,
      get() {
        const p = paras()[0];
        if (!p) return 1;
        return ALIGN_TO[effectivePara(shape, p).algn] || 1;
      },
      set(v) {
        const algn = ALIGN_FROM[v | 0];
        if (!algn) gap('G-ENGINE', 'Alignment');
        edit('Alignment', () => { paras().forEach((p) => { p.pp = Object.assign({}, p.pp, { algn }); }); });
      },
    });
    function space(key) {
      return {
        enumerable: true,
        get() {
          const p = paras()[0];
          const v = p && effectivePara(shape, p)[key];
          if (v == null) return 0;
          if (typeof v === 'number') return v;
          if (v.pt != null) return v.pt;
          gap('G-ENGINE', key);
        },
        set(n) {
          edit(key, () => { paras().forEach((p) => { p.pp = Object.assign({}, p.pp, { [key]: { pt: +n } }); }); });
        },
      };
    }
    Object.defineProperty(fmt, 'SpaceBefore', space('spcBef'));
    Object.defineProperty(fmt, 'SpaceAfter', space('spcAft'));
    Object.defineProperty(fmt, 'LineRuleWithin', {
      enumerable: true,
      get() {
        const p = paras()[0];
        const v = p && effectivePara(shape, p).lnSpc;
        if (v && v.pt != null) return 0;
        return -1;
      },
      set(v) { fmt._rule = on(v); },
    });
    Object.defineProperty(fmt, 'SpaceWithin', {
      enumerable: true,
      get() {
        const p = paras()[0];
        const v = p && effectivePara(shape, p).lnSpc;
        if (!v) return 1;
        if (v.pt != null) return v.pt;
        if (v.pct != null) return v.pct / 100;
        return 1;
      },
      set(n) {
        edit('Line spacing', () => {
          const rule = fmt._rule !== false;
          paras().forEach((p) => {
            p.pp = Object.assign({}, p.pp, { lnSpc: rule ? { pct: Math.round(+n * 100) } : { pt: +n } });
          });
        });
      },
    });
    const bullet = stamp({}, fmt);
    Object.defineProperty(bullet, 'Type', {
      enumerable: true,
      get() {
        const p = paras()[0];
        const bu = p && effectivePara(shape, p).bu;
        if (!bu || bu.t === 'none') return 0;
        if (bu.t === 'auto' || bu.t === 'num') return 2;
        return 1;
      },
      set(v) {
        const t = (v | 0) === 0 ? 'none' : (v | 0) === 2 ? 'auto' : 'char';
        if ((v | 0) > 2) gap('G-ENGINE', 'Bullet');
        edit('Bullet', () => {
          paras().forEach((p) => {
            const bu = t === 'char' ? { t: 'char', ch: '•' } : { t: t };
            p.pp = Object.assign({}, p.pp, { bu: bu });
          });
        });
      },
    });
    Object.defineProperty(bullet, 'Visible', {
      enumerable: true,
      get() { return tri(bullet.Type !== 0); },
      set(v) { bullet.Type = on(v) ? 1 : 0; },
    });
    fmt.Bullet = bullet;
    return fmt;
  }

  function makeTextRange(parent, shape, tx, anchor) {
    const range = stamp({}, parent);
    function a() { return anchor(); }
    Object.defineProperty(range, 'Text', {
      enumerable: true,
      get() { return spanText(tx, a()); },
      set(v) { edit('Text', () => writeSpan(tx, a(), v)); },
    });
    range.InsertBefore = function (text) {
      edit('Insert', () => { const s = a(); T.insertAt(tx, { p: s.p0, o: s.o0 }, normText(text)); });
    };
    range.InsertAfter = function (text) {
      edit('Insert', () => { const s = a(); T.insertAt(tx, { p: s.p1, o: s.o1 }, normText(text)); });
    };
    range.Delete = function () { edit('Delete', () => { T.deleteRange(tx, a()); }); };
    range.Select = function () {
      const s = a();
      state.text = { shapeId: shape.id, p0: s.p0, o0: s.o0, p1: s.p1, o1: s.o1 };
      state.shapes = [shape.id];
      state.slideSel = null;
      const sl = slideHolding(shape);
      const idx = pres().slides.indexOf(sl);
      if (idx >= 0) state.slide = idx;
      if (L.ed && typeof L.ed.select === 'function') L.ed.select([shape.id]);
    };
    Object.defineProperty(range, 'Font', { enumerable: true, get() { return makeFont(range, shape, tx, a); } });
    Object.defineProperty(range, 'ParagraphFormat', { enumerable: true, get() { return makeParagraphFormat(range, shape, tx, a); } });
    Object.defineProperty(range, 'IndentLevel', {
      enumerable: true,
      get() { const p = parasTouched(tx, a())[0]; return p ? (p.lvl || 0) + 1 : 1; },
      set(v) {
        const n = (v | 0) - 1;
        if (n < 0 || n > 8) gap('G-ENGINE', 'IndentLevel');
        edit('Indent', () => { parasTouched(tx, a()).forEach((p) => { p.lvl = n; }); });
      },
    });
    Object.defineProperty(range, 'Paragraphs', {
      enumerable: true,
      get() {
        return collection(range, () => parasTouched(tx, a()).map((p, i) => {
          const pi = a().p0 + i;
          const len = T.paraText(p).length;
          const para = makeParagraph(range, shape, tx, () => ({ p0: pi, o0: 0, p1: pi, o1: len }));
          para._name = () => String(i + 1);
          return para;
        }));
      },
    });
    Object.defineProperty(range, 'Characters', {
      enumerable: true,
      get() { return charactersOf(range, () => charCount(tx, a())); },
    });
    range.InsertSymbol = function (characterNumber, fontName) {
      const ch = String.fromCodePoint(characterNumber | 0);
      edit('Symbol', () => {
        const s = a();
        const pos = T.insertAt(tx, { p: s.p1, o: s.o1 }, ch, fontName ? { font: String(fontName) } : undefined);
        void pos;
      });
    };
    range.InsertDateTime = function () {
      const text = L.fmtDate ? L.fmtDate(new Date(), 'datetime1') : new Date().toDateString();
      range.InsertAfter(text);
    };
    range.InsertSlideNumber = function () {
      edit('Slide Number', () => {
        const s = a();
        const p = tx.ps[s.p1];
        if (!p) return;
        const i = T.splitAt(p, s.o1);
        p.rs.splice(i, 0, { t: '', fld: 'slidenum' });
        T.normalize(p);
      });
    };
    denyGet(range, ['Start', 'Length', 'BoundLeft', 'BoundTop', 'BoundWidth', 'BoundHeight'], 'G-TEXT');
    denyFn(range, ['Find', 'Replace'], 'G-TEXT');
    Object.defineProperty(range, 'Words', { enumerable: true, get() { gap('G-TEXT', 'Words'); } });
    Object.defineProperty(range, 'Sentences', { enumerable: true, get() { gap('G-TEXT', 'Sentences'); } });
    Object.defineProperty(range, 'Lines', { enumerable: true, get() { gap('G-TEXT', 'Lines'); } });
    Object.defineProperty(range, 'Runs', { enumerable: true, get() { gap('G-TEXT', 'Runs'); } });
    range.Hyperlinks = hyperlinksOf(range, () => tx, () => a());
    return range;
  }

  function makeParagraph(parent, shape, tx, anchor) {
    const para = stamp({}, parent);
    Object.defineProperty(para, 'Range', { enumerable: true, get() { return makeTextRange(para, shape, tx, anchor); } });
    Object.defineProperty(para, 'Alignment', {
      enumerable: true,
      get() { return para.Range.ParagraphFormat.Alignment; },
      set(v) { para.Range.ParagraphFormat.Alignment = v; },
    });
    Object.defineProperty(para, 'Font', { enumerable: true, get() { return para.Range.Font; } });
    Object.defineProperty(para, 'ParagraphFormat', { enumerable: true, get() { return para.Range.ParagraphFormat; } });
    Object.defineProperty(para, 'IndentLevel', {
      enumerable: true,
      get() { return para.Range.IndentLevel; },
      set(v) { para.Range.IndentLevel = v; },
    });
    return para;
  }

  function hyperlinksOf(parent, getTx, getAnchor) {
    function hits() {
      const tx = getTx();
      const a = getAnchor ? getAnchor() : whole(tx);
      const out = [];
      if (!tx || !tx.ps) return out;
      tx.ps.forEach((p, pi) => {
        if (pi < a.p0 || pi > a.p1) return;
        let o = 0;
        for (const r of p.rs) {
          const end = o + r.t.length;
          const from = pi === a.p0 ? a.o0 : 0;
          const to = pi === a.p1 ? a.o1 : T.paraText(p).length;
          if (r.link && end > from && o < to) out.push({ tx: tx, pi: pi, o0: o, o1: end, link: r.link });
          o = end;
        }
      });
      return out;
    }
    const col = collection(parent, () => hits().map((h, i) => {
      const link = makeHyperlink(col, h);
      link._name = () => String(i + 1);
      return link;
    }));
    col.Add = function (anchorRange, address, subAddress, screenTip, textToDisplay) {
      const tx = getTx();
      edit('Hyperlink', () => {
        let a = getAnchor ? getAnchor() : whole(tx);
        if (textToDisplay != null && textToDisplay !== '') writeSpan(tx, a, textToDisplay);
        a = getAnchor ? getAnchor() : whole(tx);
        const link = { url: address == null ? '' : String(address) };
        if (subAddress) link.sub = String(subAddress);
        if (screenTip) link.tip = String(screenTip);
        T.mapRange(tx, a, (r) => Object.assign({}, r, { link: L.clone(link) }));
        void anchorRange;
      });
      const list = hits();
      return list.length ? makeHyperlink(col, list[list.length - 1]) : null;
    };
    return col;
  }

  function makeHyperlink(parent, hit) {
    const link = stamp({}, parent);
    function run() {
      const p = hit.tx.ps[hit.pi];
      let o = 0;
      for (const r of p.rs) {
        if (o === hit.o0 && r.link) return r;
        o += r.t.length;
      }
      return null;
    }
    Object.defineProperty(link, 'Address', {
      enumerable: true,
      get() { const r = run(); return (r && r.link && (r.link.url || r.link.address)) || ''; },
      set(v) { edit('Hyperlink', () => { const r = run(); if (r && r.link) r.link.url = String(v); }); },
    });
    Object.defineProperty(link, 'SubAddress', {
      enumerable: true,
      get() { const r = run(); return (r && r.link && (r.link.sub || '')) || ''; },
      set(v) { edit('Hyperlink', () => { const r = run(); if (r && r.link) r.link.sub = String(v); }); },
    });
    Object.defineProperty(link, 'ScreenTip', {
      enumerable: true,
      get() { const r = run(); return (r && r.link && r.link.tip) || ''; },
      set(v) { edit('Hyperlink', () => { const r = run(); if (r && r.link) r.link.tip = String(v); }); },
    });
    Object.defineProperty(link, 'TextToDisplay', {
      enumerable: true,
      get() { const r = run(); return r ? r.t : ''; },
      set(v) { edit('Hyperlink', () => { const r = run(); if (r) r.t = String(v); }); },
    });
    Object.defineProperty(link, 'Range', {
      enumerable: true,
      get() {
        const shape = shapeForTx(hit.tx);
        return makeTextRange(link, shape, hit.tx, () => ({ p0: hit.pi, o0: hit.o0, p1: hit.pi, o1: hit.o1 }));
      },
    });
    link.Delete = function () {
      edit('Hyperlink', () => { const r = run(); if (r) delete r.link; });
    };
    link.Follow = function () { gap('G-HOST', 'Follow'); };
    denyGet(link, ['Type'], 'G-ENGINE');
    return link;
  }

  function shapeForTx(tx) {
    let found = null;
    M.walk(pres().slides.flatMap((s) => s.shapes), (sh) => {
      if (sh.tx === tx) { found = sh; return false; }
      if (sh.tbl) for (const row of sh.tbl.rows) for (const c of row.cells) if (c.tx === tx) { found = sh; return false; }
      return true;
    });
    return found || { id: '', tx: tx };
  }

  function makeTextFrame(parent, shape, tx) {
    const frame = stamp({}, parent);
    if (!tx) {
      denyGet(frame, ['TextRange', 'WordWrap', 'AutoSize', 'MarginLeft', 'MarginRight', 'MarginTop', 'MarginBottom', 'VerticalAnchor'], 'G-ENGINE');
      return frame;
    }
    function body() { return shape.tx || tx; }
    Object.defineProperty(frame, 'TextRange', {
      enumerable: true,
      get() { const b = body(); return makeTextRange(frame, shape, b, () => whole(b)); },
    });
    Object.defineProperty(frame, 'HasText', { enumerable: true, get() { return tri(!T.isEmpty(body())); } });
    Object.defineProperty(frame, 'WordWrap', {
      enumerable: true,
      get() { return tri(body().wrap !== false); },
      set(v) { edit('Word wrap', () => { body().wrap = on(v); }); },
    });
    Object.defineProperty(frame, 'AutoSize', {
      enumerable: true,
      get() { return body().autofit === 'shape' ? 1 : body().autofit === 'norm' ? 2 : 0; },
      set(v) {
        const n = v | 0;
        if (n !== 0 && n !== 1 && n !== 2) gap('G-ENGINE', 'AutoSize');
        edit('AutoSize', () => { body().autofit = n === 1 ? 'shape' : n === 2 ? 'norm' : 'none'; });
      },
    });
    function margin(i, name) {
      return {
        enumerable: true,
        get() { const ins = body().ins || [7.2, 3.6, 7.2, 3.6]; return ins[i]; },
        set(v) {
          edit(name, () => {
            const ins = (body().ins || [7.2, 3.6, 7.2, 3.6]).slice();
            ins[i] = +v;
            body().ins = ins;
          });
        },
      };
    }
    Object.defineProperty(frame, 'MarginLeft', margin(0, 'Margin'));
    Object.defineProperty(frame, 'MarginTop', margin(1, 'Margin'));
    Object.defineProperty(frame, 'MarginRight', margin(2, 'Margin'));
    Object.defineProperty(frame, 'MarginBottom', margin(3, 'Margin'));
    Object.defineProperty(frame, 'VerticalAnchor', {
      enumerable: true,
      get() { return VANCH_TO[body().anchor] || 1; },
      set(v) {
        const a = VANCH_FROM[v | 0];
        if (!a) gap('G-ENGINE', 'VerticalAnchor');
        edit('Anchor', () => { body().anchor = a; });
      },
    });
    denyGet(frame, ['Orientation'], 'G-ENGINE');
    return frame;
  }

  function makeFill(parent, getFill, setFill) {
    const fill = stamp({}, parent);
    Object.defineProperty(fill, 'Type', {
      enumerable: true,
      get() {
        const f = getFill();
        if (!f || f.t === 'none') return 0;
        if (f.t === 'solid') return 1;
        if (f.t === 'grad') return 3;
        return 0;
      },
      set(v) {
        if ((v | 0) !== 1) gap('G-ENGINE', 'Fill');
        edit('Fill', () => {
          const f = getFill() || { t: 'solid', c: 'accent1', a: 1 };
          setFill({ t: 'solid', c: f.c || 'accent1', a: f.a == null ? 1 : f.a });
        });
      },
    });
    Object.defineProperty(fill, 'ForeColor', {
      enumerable: true,
      get() {
        return colorFormat(fill, () => {
          const f = getFill();
          return (f && f.c) || 'accent1';
        }, (hex) => {
          edit('Fill', () => {
            const f = Object.assign({ t: 'solid', a: 1 }, getFill(), { c: hex });
            if (f.t === 'none') f.t = 'solid';
            setFill(f);
          });
        });
      },
    });
    Object.defineProperty(fill, 'Transparency', {
      enumerable: true,
      get() { const f = getFill(); const a = f && f.a != null ? f.a : 1; return 1 - a; },
      set(v) {
        edit('Fill', () => {
          const f = Object.assign({ t: 'solid', c: 'accent1', a: 1 }, getFill());
          f.a = 1 - (+v);
          setFill(f);
        });
      },
    });
    Object.defineProperty(fill, 'Visible', {
      enumerable: true,
      get() { const f = getFill(); return tri(f && f.t !== 'none'); },
      set(v) {
        edit('Fill', () => {
          if (!on(v)) setFill({ t: 'none' });
          else {
            const f = Object.assign({ t: 'solid', c: 'accent1', a: 1 }, getFill());
            if (f.t === 'none') f.t = 'solid';
            setFill(f);
          }
        });
      },
    });
    return fill;
  }

  function makeLine(parent, getLine, setLine) {
    const line = stamp({}, parent);
    Object.defineProperty(line, 'Weight', {
      enumerable: true,
      get() { const ln = getLine() || {}; return ln.w == null ? 0.75 : ln.w; },
      set(v) { edit('Line', () => { const ln = Object.assign({ c: 'tx1', dash: 'solid' }, getLine(), { w: +v }); if (ln.t === 'none') delete ln.t; setLine(ln); }); },
    });
    Object.defineProperty(line, 'ForeColor', {
      enumerable: true,
      get() {
        return colorFormat(line, () => ((getLine() && getLine().c) || 'tx1'), (hex) => {
          edit('Line', () => { const ln = Object.assign({ w: 0.75, dash: 'solid' }, getLine(), { c: hex }); if (ln.t === 'none') delete ln.t; setLine(ln); });
        });
      },
    });
    Object.defineProperty(line, 'Visible', {
      enumerable: true,
      get() { const ln = getLine(); return tri(!ln || ln.t !== 'none'); },
      set(v) {
        edit('Line', () => {
          if (!on(v)) setLine(Object.assign({}, getLine(), { t: 'none' }));
          else { const ln = Object.assign({ c: 'tx1', w: 0.75, dash: 'solid' }, getLine()); delete ln.t; setLine(ln); }
        });
      },
    });
    denyGet(line, ['DashStyle'], 'G-ENGINE');
    return line;
  }

  function shapeType(sh) {
    if (sh.ph) return 14;
    if (sh.type === 'text') return 17;
    if (sh.type === 'image') return 13;
    if (sh.type === 'table') return 19;
    if (sh.type === 'chart') return 3;
    if (sh.type === 'group') return 6;
    if (sh.type === 'line') return 9;
    if (sh.type === 'wordart') return 15;
    return 1;
  }

  function makeShape(parent, slide, sh) {
    const id = sh.id;
    function place() {
      const slides = pres().slides;
      for (const s of slides) {
        const f = M.find(s, id);
        if (f) return f;
      }
      if (slide && M.find(slide, id)) return M.find(slide, id);
      missing();
    }
    function shNow() { return place().shape; }
    function slideNow() {
      for (const s of pres().slides) if (M.find(s, id)) return s;
      return slide;
    }
    const shape = stamp({}, parent);
    shape._name = () => shNow().name || '';
    function box(key) {
      return {
        enumerable: true,
        get() { return shNow()[key] || 0; },
        set(v) { edit(key, () => { shNow()[key] = +v; }); },
      };
    }
    Object.defineProperty(shape, 'Left', box('x'));
    Object.defineProperty(shape, 'Top', box('y'));
    Object.defineProperty(shape, 'Width', box('w'));
    Object.defineProperty(shape, 'Height', box('h'));
    Object.defineProperty(shape, 'Rotation', {
      enumerable: true,
      get() { return shNow().rot || 0; },
      set(v) { edit('Rotation', () => { shNow().rot = +v; }); },
    });
    Object.defineProperty(shape, 'Name', {
      enumerable: true,
      get() { return shNow().name || ''; },
      set(v) { edit('Name', () => { shNow().name = String(v); }); },
    });
    Object.defineProperty(shape, 'Id', { enumerable: true, get() { return id; } });
    Object.defineProperty(shape, 'Type', { enumerable: true, get() { return shapeType(shNow()); } });
    Object.defineProperty(shape, 'HasTextFrame', { enumerable: true, get() { return tri(!!shNow().tx); } });
    Object.defineProperty(shape, 'AutoShapeType', {
      enumerable: true,
      get() {
        const code = GEOM_CODE[shNow().geom];
        if (!code) gap('G-ENGINE', 'AutoShapeType');
        return code;
      },
      set(v) {
        const geom = GEOM[v | 0];
        if (!geom) gap('G-ENGINE', 'AutoShapeType');
        edit('AutoShape', () => { shNow().geom = geom; });
      },
    });
    Object.defineProperty(shape, 'TextFrame', { enumerable: true, get() { const s = shNow(); return makeTextFrame(shape, s, s.tx); } });
    Object.defineProperty(shape, 'TextFrame2', { enumerable: true, get() { return shape.TextFrame; } });
    Object.defineProperty(shape, 'Fill', {
      enumerable: true,
      get() { return makeFill(shape, () => shNow().fill, (f) => { shNow().fill = f; }); },
    });
    Object.defineProperty(shape, 'Line', {
      enumerable: true,
      get() { return makeLine(shape, () => shNow().line, (ln) => { shNow().line = ln; }); },
    });
    Object.defineProperty(shape, 'PlaceholderFormat', {
      enumerable: true,
      get() {
        const cur = shNow();
        if (!cur.ph) gap('G-ENGINE', 'PlaceholderFormat');
        const ph = stamp({}, shape);
        Object.defineProperty(ph, 'Type', {
          enumerable: true,
          get() {
            const s = shNow();
            if (s.tx && s.tx.vert === 'vert' && s.ph.type === 'title') return 5;
            if (s.tx && s.tx.vert === 'vert' && s.ph.type === 'body') return 6;
            return PH[s.ph.type] || 0;
          },
        });
        Object.defineProperty(ph, 'ContainedType', { enumerable: true, get() { return ph.Type; } });
        return ph;
      },
    });
    Object.defineProperty(shape, 'Table', {
      enumerable: true,
      get() {
        const cur = shNow();
        if (cur.type !== 'table' || !cur.tbl) gap('G-ENGINE', 'Table');
        return makeTable(shape, slideNow(), cur);
      },
    });
    Object.defineProperty(shape, 'HasTable', { enumerable: true, get() { return tri(shNow().type === 'table'); } });
    Object.defineProperty(shape, 'GroupItems', {
      enumerable: true,
      get() {
        const cur = shNow();
        if (cur.type !== 'group' || !cur.kids) gap('G-ENGINE', 'GroupItems');
        return makeShapes(shape, slideNow(), () => shNow().kids, false);
      },
    });
    shape.Delete = function () {
      edit('Delete', () => {
        const f = place();
        const i = f.list.indexOf(f.shape);
        if (i >= 0) f.list.splice(i, 1);
        state.shapes = state.shapes.filter((sid) => sid !== id);
      });
    };
    shape.Duplicate = function () {
      let copy = null;
      edit('Duplicate', () => {
        const cur = shNow();
        copy = JSON.parse(JSON.stringify(cur));
        copy.id = L.uid('s');
        if (copy.ph && !['dt', 'ftr', 'sldNum'].includes(copy.ph.type)) delete copy.ph;
        const re = (x) => { if (x !== copy) x.id = L.uid('s'); if (x.kids) x.kids.forEach(re); };
        if (copy.kids) copy.kids.forEach(re);
        const f = place();
        const i = f.list.indexOf(f.shape);
        f.list.splice(i + 1, 0, copy);
      });
      return makeShape(parent, slideNow(), copy);
    };
    shape.Select = function () {
      const host = slideNow();
      const idx = pres().slides.indexOf(host);
      if (idx >= 0) state.slide = idx;
      state.shapes = [id];
      state.text = null;
      state.slideSel = null;
      if (L.ed) {
        L.ed.idx = state.slide;
        if (typeof L.ed.select === 'function') L.ed.select([id]);
      }
    };
    shape.ZOrder = function (cmd) {
      edit('Order', () => {
        const host = slideNow();
        const cur = shNow();
        const list = host.shapes;
        const i = list.indexOf(cur);
        if (i < 0) return;
        const n = cmd | 0;
        list.splice(i, 1);
        if (n === 0) list.push(cur);
        else if (n === 1) list.unshift(cur);
        else if (n === 2) list.splice(Math.min(list.length, i + 1), 0, cur);
        else if (n === 3) list.splice(Math.max(0, i - 1), 0, cur);
        else { list.splice(i, 0, cur); gap('G-ENGINE', 'ZOrder'); }
      });
    };
    Object.defineProperty(shape, 'ZOrderPosition', {
      enumerable: true,
      get() { const i = slideNow().shapes.indexOf(shNow()); return i < 0 ? 0 : i + 1; },
    });
    Object.defineProperty(shape, 'Hyperlinks', {
      enumerable: true,
      get() {
        const cur = shNow();
        if (!cur.tx) gap('G-ENGINE', 'Hyperlinks');
        return hyperlinksOf(shape, () => shNow().tx, () => whole(shNow().tx));
      },
    });
    denyGet(shape, ['Chart', 'OLEFormat', 'MediaFormat', 'ActionSettings', 'AnimationSettings', 'Nodes', 'Tags'], 'G-ENGINE');
    return shape;
  }

  function makeShapes(parent, slide, getList, canAdd) {
    const col = collection(parent, () => getList().map((sh) => {
      const wrap = makeShape(col, slide, sh);
      return wrap;
    }));
    Object.defineProperty(col, 'Title', {
      enumerable: true,
      get() {
        const sh = getList().find((s) => s.ph && (s.ph.type === 'title' || s.ph.type === 'ctrTitle'));
        if (!sh) missing();
        return makeShape(col, slide, sh);
      },
    });
    Object.defineProperty(col, 'Placeholders', {
      enumerable: true,
      get() {
        return collection(col, () => getList().filter((s) => s.ph).map((sh) => makeShape(col, slide, sh)));
      },
    });
    if (canAdd) {
      col.AddShape = function (type, left, top, width, height) {
        const geom = GEOM[type | 0];
        if (!geom || !L.geom || !L.geom.get) gap('G-ENGINE', 'AddShape');
        let sh = null;
        edit('Shape', () => {
          sh = M.newShape(pres(), slide, geom, +left, +top, +width, +height);
          slide.shapes.push(sh);
        });
        const wrap = makeShape(col, slide, sh);
        wrap.Select();
        return wrap;
      };
      col.AddTextbox = function (orientation, left, top, width, height) {
        if (orientation !== undefined && (orientation | 0) !== 1) gap('G-ENGINE', 'AddTextbox');
        let sh = null;
        edit('Text Box', () => {
          sh = M.newTextBox(pres(), slide, +left, +top, +width, +height, true);
          slide.shapes.push(sh);
        });
        const wrap = makeShape(col, slide, sh);
        wrap.Select();
        return wrap;
      };
      col.AddTable = function (numRows, numColumns, left, top, width) {
        const rows = numRows | 0, cols = numColumns | 0;
        if (rows < 1 || cols < 1) missing();
        let sh = null;
        edit('Table', () => {
          sh = M.newTable(pres(), slide, rows, cols, +left, +top, +(width == null ? 400 : width));
          slide.shapes.push(sh);
        });
        return makeShape(col, slide, sh);
      };
      col.AddPicture = function (fileName, linkToFile) {
        if (linkToFile === true || linkToFile === -1) {
          const e = new Error('The parameter LinkToFile is not available.');
          e.code = 'error';
          throw e;
        }
        if (!L.hostAudit || !L.hostAudit.has('Picture')) gap('G-HOST', 'AddPicture');
        L.hostAudit.rejectExtra(arguments, 2, ['FileName', 'LinkToFile', 'SaveWithDocument', 'Left', 'Top', 'Width', 'Height']);
        return L.hostAudit.readBytes(fileName, 'Picture', 'RequestApproval').then((bytes) => {
          const file = new File([bytes], String(fileName).split(/[/\\]/).pop());
          return L.app.insertPictureFiles([file]);
        });
      };
      denyFn(col, ['AddChart', 'AddSmartArt', 'AddDiagram', 'AddMediaObject', 'AddOLEObject', 'AddTextEffect'], 'G-ENGINE');
    }
    return col;
  }

  function makeTable(parent, slide, sh) {
    const tbl = sh.tbl;
    const table = stamp({}, parent);
    function cellAt(r, c) {
      const row = tbl.rows[(r | 0) - 1];
      const cell = row && row.cells[(c | 0) - 1];
      if (!cell) missing();
      return cell;
    }
    table.Cell = function (row, column) { return makeCell(table, sh, (row | 0) - 1, (column | 0) - 1); };
    Object.defineProperty(table, 'Rows', {
      enumerable: true,
      get() {
        const rows = collection(table, () => tbl.rows.map((row, i) => makeRow(table, sh, i)));
        rows.Add = function (beforeRow) {
          let at = tbl.rows.length;
          edit('Row', () => {
            if (beforeRow && beforeRow._index != null) at = beforeRow._index;
            const src = tbl.rows[Math.max(0, at - 1)] || tbl.rows[0];
            const cells = src.cells.map((c) => ({ tx: T.body([T.para('')], L.clone(c.tx)), fill: L.clone(c.fill), bd: L.clone(c.bd) }));
            tbl.rows.splice(at, 0, { h: src.h || 28.8, cells: cells });
            sh.h = tbl.rows.reduce((n, r) => n + (r.h || 0), 0);
          });
          return makeRow(table, sh, at === tbl.rows.length ? at - 1 : at);
        };
        return rows;
      },
    });
    Object.defineProperty(table, 'Columns', {
      enumerable: true,
      get() {
        const cols = collection(table, () => tbl.cols.map((w, i) => makeColumn(table, sh, i)));
        cols.Add = function (before) {
          let at = tbl.cols.length;
          edit('Column', () => {
            if (before && before._index != null) at = before._index;
            const w = tbl.cols[Math.max(0, at - 1)] || 72;
            tbl.cols.splice(at, 0, w);
            for (const row of tbl.rows) {
              const src = row.cells[Math.max(0, at - 1)] || row.cells[0];
              row.cells.splice(at, 0, { tx: T.body([T.para('')]), fill: L.clone(src.fill), bd: L.clone(src.bd) });
            }
            sh.w = tbl.cols.reduce((n, x) => n + x, 0);
          });
          return makeColumn(table, sh, at === tbl.cols.length ? at - 1 : at);
        };
        return cols;
      },
    });
    table.Delete = function () { parent.Delete ? parent.Delete() : gap('G-ENGINE', 'Delete'); };
    denyFn(table, ['Merge', 'ApplyStyle'], 'G-ENGINE');
    void cellAt;
    return table;
  }

  function makeCell(parent, sh, ri, ci) {
    const cell = stamp({}, parent);
    cell._index = ri;
    function model() {
      const row = sh.tbl.rows[ri];
      const c = row && row.cells[ci];
      if (!c) missing();
      return c;
    }
    const fake = { id: sh.id + ':' + ri + ':' + ci, tx: model().tx, type: 'text' };
    Object.defineProperty(cell, 'Shape', {
      enumerable: true,
      get() {
        fake.tx = model().tx;
        const shape = stamp({}, cell);
        Object.defineProperty(shape, 'TextFrame', { enumerable: true, get() { return makeTextFrame(shape, fake, fake.tx); } });
        Object.defineProperty(shape, 'TextFrame2', { enumerable: true, get() { return shape.TextFrame; } });
        Object.defineProperty(shape, 'Fill', {
          enumerable: true,
          get() { return makeFill(shape, () => model().fill, (f) => { model().fill = f; }); },
        });
        return shape;
      },
    });
    return cell;
  }

  function makeRow(parent, sh, index) {
    const row = stamp({}, parent);
    row._index = index;
    Object.defineProperty(row, 'Height', {
      enumerable: true,
      get() { return sh.tbl.rows[index].h; },
      set(v) { edit('Row', () => { sh.tbl.rows[index].h = +v; sh.h = sh.tbl.rows.reduce((n, r) => n + (r.h || 0), 0); }); },
    });
    Object.defineProperty(row, 'Cells', {
      enumerable: true,
      get() { return collection(row, () => sh.tbl.rows[index].cells.map((c, ci) => makeCell(row, sh, index, ci))); },
    });
    row.Delete = function () {
      if (sh.tbl.rows.length < 2) gap('G-ENGINE', 'Delete');
      edit('Delete Row', () => {
        sh.tbl.rows.splice(index, 1);
        sh.h = sh.tbl.rows.reduce((n, r) => n + (r.h || 0), 0);
      });
    };
    return row;
  }

  function makeColumn(parent, sh, index) {
    const col = stamp({}, parent);
    col._index = index;
    Object.defineProperty(col, 'Width', {
      enumerable: true,
      get() { return sh.tbl.cols[index]; },
      set(v) { edit('Column', () => { sh.tbl.cols[index] = +v; sh.w = sh.tbl.cols.reduce((n, x) => n + x, 0); }); },
    });
    col.Delete = function () {
      if (sh.tbl.cols.length < 2) gap('G-ENGINE', 'Delete');
      edit('Delete Column', () => {
        sh.tbl.cols.splice(index, 1);
        for (const row of sh.tbl.rows) row.cells.splice(index, 1);
        sh.w = sh.tbl.cols.reduce((n, x) => n + x, 0);
      });
    };
    return col;
  }

  function makeSlide(parent, slide, index) {
    const obj = stamp({}, parent);
    obj._index = index;
    obj._name = () => slide.name || ('Slide' + (index + 1));
    Object.defineProperty(obj, 'Name', {
      enumerable: true,
      get() { return slide.name || ('Slide' + (index + 1)); },
      set(v) { edit('Slide name', () => { slide.name = String(v); }); },
    });
    Object.defineProperty(obj, 'SlideIndex', { enumerable: true, get() { return pres().slides.indexOf(slide) + 1; } });
    Object.defineProperty(obj, 'SlideNumber', {
      enumerable: true,
      get() { return pres().slides.indexOf(slide) + (pres().firstNum || 1); },
    });
    Object.defineProperty(obj, 'SlideID', { enumerable: true, get() { gap('G-ENGINE', 'SlideID'); } });
    Object.defineProperty(obj, 'Layout', {
      enumerable: true,
      get() { return LAYOUT_CODE[slide.layout] || 32; },
      set(v) {
        const key = LAYOUT[v | 0];
        if (!key) gap('G-ENGINE', 'Layout');
        edit('Layout', () => { M.applyLayout(pres(), slide, key); });
      },
    });
    Object.defineProperty(obj, 'Shapes', { enumerable: true, get() { return makeShapes(obj, slide, () => slide.shapes, true); } });
    Object.defineProperty(obj, 'Design', {
      enumerable: true,
      get() { return makeDesign(obj, designOf(slide)); },
      set(v) {
        const name = v && v.Name ? v.Name : String(v);
        const hit = (M.DESIGNS || []).find((d) => d.name === name || d.key === name);
        if (!hit) gap('G-ENGINE', 'Design');
        edit('Design', () => { M.applyDesign(pres(), hit.key, [slide]); });
      },
    });
    Object.defineProperty(obj, 'HeadersFooters', { enumerable: true, get() { return makeHeaders(obj); } });
    Object.defineProperty(obj, 'SlideShowTransition', { enumerable: true, get() { return makeTransition(obj, slide); } });
    Object.defineProperty(obj, 'Hyperlinks', {
      enumerable: true,
      get() {
        const tx = { ps: [] };
        M.walk(slide.shapes, (s) => { if (s.tx) tx.ps = tx.ps.concat(s.tx.ps); return true; });
        return hyperlinksOf(obj, () => tx);
      },
    });
    obj.Select = function () { focusSlide(pres().slides.indexOf(slide)); };
    obj.Delete = function () {
      edit('Delete Slide', () => {
        const i = pres().slides.indexOf(slide);
        if (i < 0) return;
        pres().slides.splice(i, 1);
        if (state.slide >= pres().slides.length) state.slide = Math.max(0, pres().slides.length - 1);
      });
    };
    obj.Duplicate = function () {
      let copy = null;
      edit('Duplicate Slide', () => {
        copy = JSON.parse(JSON.stringify(slide));
        copy.id = L.uid('sl');
        const re = (x) => { x.id = L.uid('s'); if (x.kids) x.kids.forEach(re); };
        copy.shapes.forEach(re);
        const i = pres().slides.indexOf(slide);
        M.insertSlides(pres(), i + 1, [copy]);
      });
      return makeSlide(parent, copy, pres().slides.indexOf(copy));
    };
    obj.MoveTo = function (toPos) {
      const to = (toPos | 0) - 1;
      const slides = pres().slides;
      if (to < 0 || to >= slides.length) missing();
      edit('Move Slide', () => {
        const i = slides.indexOf(slide);
        slides.splice(i, 1);
        slides.splice(to, 0, slide);
        state.slide = to;
      });
    };
    denyGet(obj, ['NotesPage', 'Comments', 'TimeLine', 'Tags', 'CustomerData', 'Scripts'], 'G-ENGINE');
    return obj;
  }

  function makeTransition(parent, slide) {
    const tr = stamp({}, parent);
    function t() { return slide.trans || (slide.trans = { type: 'none', spd: 'fast', click: true, after: null }); }
    function matchEffect() {
      const cur = t();
      const hit = EFFECTS.find((e) => e[1].type === (cur.type || 'none') && (e[1].dir || cur.dir) === (cur.dir || e[1].dir) && !!e[1].thruBlk === !!cur.thruBlk);
      return hit;
    }
    Object.defineProperty(tr, 'EntryEffect', {
      enumerable: true,
      get() {
        const hit = matchEffect();
        if (!hit) gap('G-ENGINE', 'EntryEffect');
        return hit[0];
      },
      set(v) {
        const hit = EFFECTS.find((e) => e[0] === (v | 0));
        if (!hit) gap('G-ENGINE', 'EntryEffect');
        edit('Transition', () => {
          const keep = { spd: t().spd, click: t().click, after: t().after };
          slide.trans = Object.assign({ type: 'none' }, hit[1], keep);
          if (!hit[1].thruBlk) delete slide.trans.thruBlk;
          if (!hit[1].dir) delete slide.trans.dir;
        });
      },
    });
    Object.defineProperty(tr, 'Speed', {
      enumerable: true,
      get() { return SPEED_TO[t().spd] || 3; },
      set(v) {
        const spd = SPEED_FROM[v | 0];
        if (!spd) gap('G-ENGINE', 'Speed');
        edit('Transition', () => { t().spd = spd; delete t().dur; });
      },
    });
    Object.defineProperty(tr, 'AdvanceOnClick', {
      enumerable: true,
      get() { return tri(t().click !== false); },
      set(v) { edit('Transition', () => { t().click = on(v); }); },
    });
    Object.defineProperty(tr, 'AdvanceOnTime', {
      enumerable: true,
      get() { return tri(t().after != null); },
      set(v) { edit('Transition', () => { t().after = on(v) ? (t().after != null ? t().after : 0) : null; }); },
    });
    Object.defineProperty(tr, 'AdvanceTime', {
      enumerable: true,
      get() { return t().after == null ? 0 : t().after / 1000; },
      set(n) { edit('Transition', () => { t().after = Math.round(+n * 1000); }); },
    });
    Object.defineProperty(tr, 'Hidden', {
      enumerable: true,
      get() { return tri(!!slide.hidden); },
      set(v) { edit('Hide Slide', () => { slide.hidden = on(v); }); },
    });
    denyGet(tr, ['Duration', 'SoundEffect'], 'G-ENGINE');
    return tr;
  }

  function makeHeaders(parent) {
    const hf = stamp({}, parent);
    function bag() { return pres().hf || (pres().hf = { dt: false, dtAuto: true, dtText: '', num: false, ftr: false, ftrText: '', notOnTitle: false }); }
    function item(key, textKey) {
      const it = stamp({}, hf);
      Object.defineProperty(it, 'Visible', {
        enumerable: true,
        get() { return tri(!!bag()[key]); },
        set(v) { edit('Header and Footer', () => { bag()[key] = on(v); }); },
      });
      if (textKey) {
        Object.defineProperty(it, 'Text', {
          enumerable: true,
          get() { return bag()[textKey] || ''; },
          set(v) { edit('Header and Footer', () => { bag()[textKey] = String(v); bag()[key] = true; }); },
        });
      }
      denyGet(it, ['Format', 'UseFormat'], 'G-ENGINE');
      return it;
    }
    Object.defineProperty(hf, 'DateAndTime', { enumerable: true, get() { return item('dt', 'dtText'); } });
    Object.defineProperty(hf, 'Footer', { enumerable: true, get() { return item('ftr', 'ftrText'); } });
    Object.defineProperty(hf, 'SlideNumber', { enumerable: true, get() { return item('num'); } });
    Object.defineProperty(hf, 'DisplayOnTitleSlide', {
      enumerable: true,
      get() { return tri(!bag().notOnTitle); },
      set(v) { edit('Header and Footer', () => { bag().notOnTitle = !on(v); }); },
    });
    return hf;
  }

  function makeDesign(parent, design) {
    const obj = stamp({}, parent);
    Object.defineProperty(obj, 'Name', { enumerable: true, get() { return (design && (design.name || design.key)) || ''; } });
    obj._name = () => obj.Name;
    return obj;
  }

  function makePageSetup(parent) {
    const ps = stamp({}, parent);
    Object.defineProperty(ps, 'SlideWidth', {
      enumerable: true,
      get() { return pres().W; },
      set(v) { edit('Page Setup', () => { pres().W = +v; }); },
    });
    Object.defineProperty(ps, 'SlideHeight', {
      enumerable: true,
      get() { return pres().H; },
      set(v) { edit('Page Setup', () => { pres().H = +v; }); },
    });
    Object.defineProperty(ps, 'FirstSlideNumber', {
      enumerable: true,
      get() { return pres().firstNum || 1; },
      set(v) { edit('Page Setup', () => { pres().firstNum = v | 0; }); },
    });
    denyGet(ps, ['SlideSize', 'NotesOrientation', 'SlideOrientation'], 'G-ENGINE');
    return ps;
  }

  function showBag() {
    return pres().show || (pres().show = { loop: false, noAnim: false, useTimings: true, from: 0, to: 0, penColor: '#FF0000', kiosk: false });
  }

  function showOrder() {
    const sh = showBag();
    const n = pres().slides.length;
    let from = 0, to = n - 1;
    if (sh.from && sh.to) { from = Math.max(0, sh.from - 1); to = Math.min(n - 1, sh.to - 1); }
    const out = [];
    for (let i = from; i <= to; i++) if (pres().slides[i] && !pres().slides[i].hidden) out.push(i);
    return out;
  }

  function makeSlideShowSettings(parent) {
    const ss = stamp({}, parent);
    Object.defineProperty(ss, 'StartingSlide', {
      enumerable: true,
      get() { return showBag().from || 1; },
      set(v) { edit('Slide Show', () => { showBag().from = v | 0; if (!showBag().to) showBag().to = pres().slides.length; }); },
    });
    Object.defineProperty(ss, 'EndingSlide', {
      enumerable: true,
      get() { return showBag().to || pres().slides.length; },
      set(v) { edit('Slide Show', () => { showBag().to = v | 0; if (!showBag().from) showBag().from = 1; }); },
    });
    Object.defineProperty(ss, 'RangeType', {
      enumerable: true,
      get() { return showBag().from && showBag().to ? 2 : 1; },
      set(v) {
        const n = v | 0;
        if (n === 3) gap('G-ENGINE', 'RangeType');
        if (n !== 1 && n !== 2) gap('G-ENGINE', 'RangeType');
        edit('Slide Show', () => {
          if (n === 1) { showBag().from = 0; showBag().to = 0; }
          else if (!showBag().from || !showBag().to) { showBag().from = 1; showBag().to = pres().slides.length; }
        });
      },
    });
    Object.defineProperty(ss, 'LoopUntilStopped', {
      enumerable: true,
      get() { return tri(!!showBag().loop); },
      set(v) { edit('Slide Show', () => { showBag().loop = on(v); }); },
    });
    Object.defineProperty(ss, 'AdvanceMode', {
      enumerable: true,
      get() { return showBag().useTimings === false ? 1 : 2; },
      set(v) {
        const n = v | 0;
        if (n !== 1 && n !== 2) gap('G-ENGINE', 'AdvanceMode');
        edit('Slide Show', () => { showBag().useTimings = n === 2; });
      },
    });
    Object.defineProperty(ss, 'ShowType', {
      enumerable: true,
      get() { return showBag().kiosk ? 3 : 1; },
      set(v) {
        const n = v | 0;
        if (n === 2) gap('G-HOST', 'ShowType');
        if (n !== 1 && n !== 3) gap('G-ENGINE', 'ShowType');
        edit('Slide Show', () => { showBag().kiosk = n === 3; });
      },
    });
    ss.Run = function () {
      const order = showOrder();
      if (!order.length) gap('G-ENGINE', 'Run');
      showBag()._dom = { on: true, pos: 0 };
      if (L.show && typeof L.show.start === 'function' && typeof document !== 'undefined' && document.body) {
        try { L.show.start({ from: order[0] }); } catch (e) { /* headless */ }
      }
    };
    denyGet(ss, ['NamedSlideShows', 'PointerColor'], 'G-ENGINE');
    return ss;
  }

  function makeShowView(parent) {
    const view = stamp({}, parent);
    function dom() { return showBag()._dom; }
    function running() { return !!(dom() && dom().on) || (L.show && L.show.active); }
    view.Next = function () {
      if (!running()) gap('G-ENGINE', 'Next');
      const order = showOrder();
      const d = dom() || (showBag()._dom = { on: true, pos: 0 });
      if (d.pos < order.length - 1) d.pos += 1;
      else if (showBag().loop) d.pos = 0;
      if (L.show && L.show.active && L.show.next) L.show.next();
    };
    view.Previous = function () {
      if (!running()) gap('G-ENGINE', 'Previous');
      const d = dom();
      if (d && d.pos > 0) d.pos -= 1;
      if (L.show && L.show.active && L.show.prev) L.show.prev();
    };
    view.GotoSlide = function (index) {
      if (!running()) gap('G-ENGINE', 'GotoSlide');
      const order = showOrder();
      const at = order.indexOf((index | 0) - 1);
      if (at < 0) missing();
      dom().pos = at;
      if (L.show && L.show.active && L.show.goto) L.show.goto((index | 0) - 1);
    };
    view.Exit = function () {
      if (showBag()._dom) showBag()._dom.on = false;
      if (L.show && L.show.active && L.show.end) {
        try { L.show.end(); } catch (e) { /* already closed */ }
      }
    };
    Object.defineProperty(view, 'CurrentShowPosition', {
      enumerable: true,
      get() { const d = dom(); return d && d.on ? d.pos + 1 : 0; },
    });
    denyGet(view, ['State', 'PointerType', 'AcceleratorsEnabled'], 'G-HOST');
    return view;
  }

  function makeShowWindow(parent) {
    const win = stamp({}, parent);
    Object.defineProperty(win, 'View', { enumerable: true, get() { return makeShowView(win); } });
    return win;
  }

  function makeShowWindows(parent) {
    const col = collection(parent, () => {
      const on = (showBag()._dom && showBag()._dom.on) || (L.show && L.show.active);
      return on ? [makeShowWindow(col)] : [];
    });
    return col;
  }

  function propsBag(custom) {
    const p = pres();
    if (custom) return p.customProps || (p.customProps = {});
    return p.props || (p.props = {});
  }

  const BUILTIN = { Title: 'title', Subject: 'subject', Author: 'author', Keywords: 'keywords', Comments: 'comments', Category: 'category', Company: 'company' };

  function makeProperty(parent, name, custom) {
    const prop = stamp({}, parent);
    Object.defineProperty(prop, 'Name', { enumerable: true, get() { return name; } });
    Object.defineProperty(prop, 'Value', {
      enumerable: true,
      get() {
        if (!custom) {
          const key = BUILTIN[name];
          if (!key) gap('G-ENGINE', name);
          return propsBag(false)[key] || '';
        }
        if (!Object.prototype.hasOwnProperty.call(propsBag(true), name)) missing();
        return propsBag(true)[name];
      },
      set(v) {
        edit('Property', () => {
          if (!custom) {
            const key = BUILTIN[name];
            if (!key) gap('G-ENGINE', name);
            propsBag(false)[key] = String(v);
          } else propsBag(true)[name] = String(v);
        });
      },
    });
    return prop;
  }

  function makeProperties(parent, custom) {
    const names = () => custom ? Object.keys(propsBag(true)) : Object.keys(BUILTIN);
    const col = collection(parent, () => names().map((n) => {
      const prop = makeProperty(col, n, custom);
      prop._name = () => n;
      return prop;
    }));
    const item = col.Item;
    col.Item = function (index) {
      if (typeof index === 'string' && !/^\d+$/.test(index)) {
        if (!custom && !BUILTIN[index]) gap('G-ENGINE', index);
        if (custom && !Object.prototype.hasOwnProperty.call(propsBag(true), index)) {
          propsBag(true)[index] = '';
        }
        return makeProperty(col, index, custom);
      }
      return item(index);
    };
    if (custom) {
      col.Add = function (name, linkToContent, type, value) {
        if (linkToContent) gap('G-ENGINE', 'Add');
        if (type !== undefined && (type | 0) !== 4) gap('G-ENGINE', 'Add');
        edit('Property', () => { propsBag(true)[String(name)] = value == null ? '' : String(value); });
        return makeProperty(col, String(name), true);
      };
    }
    return col;
  }

  function makeSlideMaster(parent) {
    const master = stamp({}, parent);
    Object.defineProperty(master, 'Name', { enumerable: true, get() { return designOf(pres().slides[clampSlide()] || {}).name || ''; } });
    Object.defineProperty(master, 'HeadersFooters', { enumerable: true, get() { return makeHeaders(master); } });
    denyGet(master, ['Shapes', 'CustomLayouts', 'Design', 'Background'], 'G-ENGINE');
    return master;
  }

  function makePresentation(wrap) {
    if (wrap.api) return wrap.api;
    const doc = stamp({}, () => Presentations);
    wrap.api = doc;
    Object.defineProperty(doc, 'Name', { enumerable: true, get() { return wrap.pres === L.pres && L.app && L.app.fileName ? L.app.fileName : wrap.name; } });
    Object.defineProperty(doc, 'FullName', {
      enumerable: true,
      get() {
        if (!L.hostAudit || !L.hostAudit.has('Path')) gap('G-HOST', 'FullName');
        return L.hostAudit.documentPath();
      },
    });
    Object.defineProperty(doc, 'Path', {
      enumerable: true,
      get() {
        if (!L.hostAudit || !L.hostAudit.has('Path')) gap('G-HOST', 'Path');
        return L.hostAudit.documentDir();
      },
    });
    Object.defineProperty(doc, 'Saved', {
      enumerable: true,
      get() { return !L.hist.dirty; },
      set(v) {
        const save = !(v === false || v === 0);
        L.hist.dirty = !save;
        if (L.app) L.app.saved = save;
      },
    });
    Object.defineProperty(doc, 'Slides', {
      enumerable: true,
      get() {
        const slides = collection(doc, () => wrap.pres.slides.map((s, i) => makeSlide(slides, s, i)));
        slides.Add = function (index, layout) {
          const key = LAYOUT[layout | 0];
          if (!key) gap('G-ENGINE', 'Add');
          const at = (index | 0) - 1;
          if (at < 0 || at > wrap.pres.slides.length) missing();
          let slide = null;
          edit('New Slide', () => {
            slide = M.newSlide(wrap.pres, key);
            M.insertSlides(wrap.pres, at, [slide]);
          });
          focusSlide(at);
          return makeSlide(slides, slide, at);
        };
        slides.Range = function (index) {
          const ids = Array.isArray(index) ? index : [index];
          const list = ids.map((n) => slideByIndex(n));
          return collection(slides, () => list.map((s) => makeSlide(slides, s, wrap.pres.slides.indexOf(s))));
        };
        return slides;
      },
    });
    Object.defineProperty(doc, 'PageSetup', { enumerable: true, get() { return makePageSetup(doc); } });
    Object.defineProperty(doc, 'SlideShowSettings', { enumerable: true, get() { return makeSlideShowSettings(doc); } });
    Object.defineProperty(doc, 'SlideShowWindows', { enumerable: true, get() { return makeShowWindows(doc); } });
    Object.defineProperty(doc, 'Designs', {
      enumerable: true,
      get() {
        return collection(doc, () => (M.DESIGNS || []).map((d) => {
          const obj = makeDesign(doc, d);
          obj._name = () => d.name;
          return obj;
        }));
      },
    });
    Object.defineProperty(doc, 'SlideMaster', { enumerable: true, get() { return makeSlideMaster(doc); } });
    Object.defineProperty(doc, 'BuiltInDocumentProperties', { enumerable: true, get() { return makeProperties(doc, false); } });
    Object.defineProperty(doc, 'CustomDocumentProperties', { enumerable: true, get() { return makeProperties(doc, true); } });
    doc.Undo = function (times) {
      const n = times == null ? 1 : (times | 0);
      let ok = false;
      for (let i = 0; i < n; i++) ok = !!(L.hist && L.hist.doUndo && L.hist.doUndo());
      return tri(ok);
    };
    doc.Redo = function (times) {
      const n = times == null ? 1 : (times | 0);
      let ok = false;
      for (let i = 0; i < n; i++) ok = !!(L.hist && L.hist.doRedo && L.hist.doRedo());
      return tri(ok);
    };
    doc.Save = function () {
      if (L.app && typeof L.app.save === 'function') {
        if (!L.app.saved) gap('G-HOST', 'Save');
        const run = () => L.app.save();
        if (!L.hostAudit) return run();
        return L.hostAudit.must('RequestApproval', 'Save', 'save the presentation').then(run);
      }
      L.hist.dirty = false;
    };
    function deckBytes(fileFormat) {
      const type = fileFormat == null || fileFormat === '' || fileFormat === 'pptx' ? (L.app && L.app.fileType) || 'pptx' : fileFormat;
      if (type === 'pdf') return L.app.buildPDF('slides');
      if (!/^(pptx|pptm|ppsx|ppsm|potx|potm)$/.test(String(type))) {
        const e = new Error('The parameter FileFormat is not available.');
        e.code = 'error';
        throw e;
      }
      return L.pptx.write(L.pres, { format: type });
    }
    doc.SaveAs = function (fileName, fileFormat) {
      if (!L.hostAudit || !L.hostAudit.has('SaveAs')) gap('G-HOST', 'SaveAs');
      L.hostAudit.rejectExtra(arguments, 2, ['FileName', 'FileFormat', 'EmbedTrueTypeFonts']);
      return L.hostAudit.saveCopy(fileName, () => deckBytes(fileFormat));
    };
    doc.SaveAs2 = doc.SaveAs;
    doc.PrintOut = function () {
      if (!L.hostAudit || !L.hostAudit.has('Print')) gap('G-HOST', 'PrintOut');
      return L.hostAudit.printPdf(() => L.app.buildPDF('slides'));
    };
    doc.ExportAsFixedFormat = function (path, format) {
      if (!L.hostAudit || !L.hostAudit.has('FixedFormat')) gap('G-HOST', 'ExportAsFixedFormat');
      return L.hostAudit.exportPdf(format, path, () => L.app.buildPDF('slides'));
    };
    doc.Export = async function (path, filter, scaleWidth, scaleHeight) {
      if (!L.hostAudit || !L.hostAudit.has('Export')) gap('G-HOST', 'Export');
      const ext = String(filter || 'png').toLowerCase();
      if (ext !== 'png' && ext !== 'jpg' && ext !== 'jpeg') {
        const e = new Error('The parameter FilterName is not available.');
        e.code = 'error';
        throw e;
      }
      const pres = L.pres;
      let width = scaleWidth == null || scaleWidth === '' ? null : +scaleWidth;
      let height = scaleHeight == null || scaleHeight === '' ? null : +scaleHeight;
      if (width == null && height == null) width = Math.round(pres.W * 2.6667);
      else if (width == null) width = Math.round(height * pres.W / pres.H);
      await L.hostAudit.must('RequestApproval', 'Export', 'export slides to ' + path, path);
      const mime = ext === 'png' ? 'image/png' : 'image/jpeg';
      const suffix = ext === 'jpeg' ? 'jpg' : ext;
      let n = 0;
      for (const slide of pres.slides) {
        n++;
        let canvas = await L.app.slideCanvas(slide, width);
        if (scaleWidth != null && scaleWidth !== '' && scaleHeight != null && scaleHeight !== '') {
          const out = document.createElement('canvas');
          out.width = +scaleWidth;
          out.height = +scaleHeight;
          out.getContext('2d').drawImage(canvas, 0, 0, out.width, out.height);
          canvas = out;
        }
        const blob = await new Promise((res, rej) => canvas.toBlob((b) => (b ? res(b) : rej(new Error('The picture could not be created.'))), mime, 0.92));
        await L.hostAudit.writeBlob(L.hostAudit.joinPath(path, 'Slide' + n + '.' + suffix), blob, 'Export', false);
      }
    };
    doc.FollowHyperlink = function (address) {
      if (!L.hostAudit || !L.hostAudit.has('Hyperlink')) gap('G-HOST', 'FollowHyperlink');
      L.hostAudit.rejectExtra(arguments, 1, ['Address', 'SubAddress', 'NewWindow', 'AddHistory', 'ExtraInfo', 'Method', 'HeaderInfo']);
      return L.hostAudit.follow(address);
    };
    doc.Close = function () {
      const i = wraps.indexOf(wrap);
      if (i < 0) return;
      wraps.splice(i, 1);
      delete wrap.api;
      if (wrap.pres === L.pres) {
        if (wraps.length) {
          const next = wraps[Math.min(i, wraps.length - 1)];
          L.pres = next.pres;
          L.hist.clear();
          state.slide = 0;
          state.shapes = [];
          state.text = null;
        } else {
          multi = false;
          L.pres = M.newPresentation();
          L.hist.clear();
          wraps.push({ pres: L.pres, name: 'Presentation1' });
          state.slide = 0;
        }
      }
    };
    denyFn(doc, ['NewWindow'], 'G-HOST');
    denyFn(doc, ['CreateVideo', 'Merge', 'PublishSlides'], 'G-ENGINE');
    return doc;
  }

  const Presentations = function (index) { return Presentations.Item(index); };
  stamp(Presentations, Application);
  Object.defineProperty(Presentations, 'Count', { enumerable: true, get() { ensure(); return wraps.length; } });
  Presentations.Item = function (index) {
    ensure();
    const i = index | 0;
    if (i < 1 || i > wraps.length) missing();
    return makePresentation(wraps[i - 1]);
  };
  Object.defineProperty(Presentations, 'First', { enumerable: true, get() { return Presentations.Count ? Presentations.Item(1) : null; } });
  Object.defineProperty(Presentations, 'Last', { enumerable: true, get() { return Presentations.Count ? Presentations.Item(Presentations.Count) : null; } });
  Presentations.Add = function (withWindow) {
    if (withWindow !== undefined) gap('G-HOST', 'Add');
    const created = M.newPresentation();
    const wrap = { pres: created, name: 'Presentation' + (wraps.length + 1) };
    wraps.push(wrap);
    multi = true;
    L.pres = created;
    if (L.hist && L.hist.clear) L.hist.clear();
    state.slide = 0;
    state.shapes = [];
    state.text = null;
    if (L.app && typeof L.app.loadPres === 'function') {
      try { L.app.loadPres(created, wrap.name); } catch (e) { L.pres = created; }
    }
    return makePresentation(wrap);
  };
  Presentations.Open = function (fileName) {
    if (!L.hostAudit || !L.hostAudit.has('Open')) gap('G-HOST', 'Open');
    L.hostAudit.rejectExtra(arguments, 1, ['FileName', 'ReadOnly', 'Untitled', 'WithWindow']);
    return L.hostAudit.openPath(fileName);
  };

  function shapeRangeOf(ids) {
    const slide = pres().slides[clampSlide()];
    const list = ids.map((id) => shapeById(slide, id)).filter(Boolean);
    const col = collection(Selection, () => list.map((sh) => makeShape(col, slide, sh)));
    Object.defineProperty(col, 'Left', {
      enumerable: true,
      get() { return list[0] ? list[0].x : 0; },
      set(v) { edit('Left', () => { list.forEach((sh) => { sh.x = +v; }); }); },
    });
    Object.defineProperty(col, 'Top', {
      enumerable: true,
      get() { return list[0] ? list[0].y : 0; },
      set(v) { edit('Top', () => { list.forEach((sh) => { sh.y = +v; }); }); },
    });
    col.Delete = function () { list.forEach((sh) => makeShape(col, slide, sh).Delete()); state.shapes = []; };
    return col;
  }

  const Selection = stamp({}, () => ActiveWindow);
  Object.defineProperty(Selection, 'Type', {
    enumerable: true,
    get() {
      if (state.text || (L.te && L.te.active && L.te.active())) return 3;
      if (selectedShapeIds().length) return 2;
      if (selectedSlideIndexes().length) return 1;
      return 0;
    },
  });
  Object.defineProperty(Selection, 'ShapeRange', {
    enumerable: true,
    get() {
      const ids = selectedShapeIds();
      if (!ids.length) gap('G-ENGINE', 'ShapeRange');
      return shapeRangeOf(ids);
    },
  });
  Object.defineProperty(Selection, 'SlideRange', {
    enumerable: true,
    get() {
      let idxs = selectedSlideIndexes();
      if (!idxs.length && pres().slides.length) idxs = [clampSlide()];
      if (!idxs.length) gap('G-ENGINE', 'SlideRange');
      return collection(Selection, () => idxs.map((i) => makeSlide(Selection, pres().slides[i], i)));
    },
  });
  Object.defineProperty(Selection, 'TextRange', {
    enumerable: true,
    get() {
      if (!state.text) gap('G-ENGINE', 'TextRange');
      const slide = pres().slides[clampSlide()];
      const sh = shapeById(slide, state.text.shapeId);
      if (!sh || !sh.tx) gap('G-ENGINE', 'TextRange');
      const t = state.text;
      return makeTextRange(Selection, sh, sh.tx, () => ({ p0: t.p0, o0: t.o0, p1: t.p1, o1: t.o1 }));
    },
  });
  Selection.Unselect = function () {
    state.shapes = [];
    state.text = null;
    state.slideSel = null;
    if (L.ed && L.ed.select) L.ed.select([]);
  };
  Selection.Delete = function () {
    if (state.text) { Selection.TextRange.Delete(); state.text = null; return; }
    const ids = selectedShapeIds();
    if (ids.length) { shapeRangeOf(ids).Delete(); return; }
    const idxs = selectedSlideIndexes();
    if (idxs.length) {
      edit('Delete Slide', () => {
        idxs.slice().reverse().forEach((i) => pres().slides.splice(i, 1));
        state.slideSel = null;
      });
    }
  };
  function rememberClip(c) {
    clip = c;
    if (L.app) L.app.clip = c;
  }
  Selection.Copy = function () {
    const slide = pres().slides[clampSlide()];
    const ids = selectedShapeIds();
    if (state.text) {
      rememberClip({ kind: 'text', text: Selection.TextRange.Text });
      return;
    }
    if (!ids.length) return;
    const shapes = ids.map((id) => shapeById(slide, id)).filter(Boolean);
    rememberClip({ kind: 'shapes', shapes: JSON.parse(JSON.stringify(shapes)), label: shapes[0] && shapes[0].name });
  };
  Selection.Cut = function () { Selection.Copy(); Selection.Delete(); };
  Selection.Paste = function () {
    const c = (L.app && L.app.clip) || clip;
    if (!c) return;
    const slide = pres().slides[clampSlide()];
    if (!slide) return;
    if (c.kind === 'text') {
      if (state.text) Selection.TextRange.InsertAfter(c.text);
      return;
    }
    if (c.kind === 'shapes') {
      edit('Paste', () => {
        const copies = JSON.parse(JSON.stringify(c.shapes));
        const re = (x) => { x.id = L.uid('s'); if (x.kids) x.kids.forEach(re); };
        copies.forEach((s) => { re(s); slide.shapes.push(s); });
        state.shapes = copies.map((s) => s.id);
        state.text = null;
      });
    }
  };
  denyGet(Selection, ['ChildShapeRange'], 'G-ENGINE');

  const ActiveWindow = stamp({}, () => Application.ActivePresentation);
  Object.defineProperty(ActiveWindow, 'Selection', { enumerable: true, get() { return Selection; } });
  Object.defineProperty(ActiveWindow, 'Presentation', { enumerable: true, get() { return Application.ActivePresentation; } });
  denyGet(ActiveWindow, ['View', 'ViewType', 'Caption', 'ActivePane', 'Panes', 'Split'], 'G-HOST');

  stamp(Application, null);
  Application.Parent = Application;
  Application.Application = Application;
  Object.defineProperty(Application, 'ActivePresentation', {
    enumerable: true, configurable: true,
    get() { return makePresentation(ensure()); },
  });
  Object.defineProperty(Application, 'Presentations', { enumerable: true, get() { ensure(); return Presentations; } });
  Object.defineProperty(Application, 'ActiveWindow', { enumerable: true, get() { return ActiveWindow; } });
  Application.FileDialog = function (type) {
    if (!L.hostAudit) gap('G-HOST', 'FileDialog');
    return L.hostAudit.fileDialog(type);
  };
  Application.HasHostAbility = function (name) { return L.hostAudit && L.hostAudit.has(name) ? -1 : 0; };
  Application.MsgBox = function (prompt, buttons, title) {
    if (!L.hostAudit) gap('G-HOST', 'MsgBox');
    return L.hostAudit.msgBox(prompt, buttons, title);
  };
  Application.InputBox = function (prompt, title, def) {
    if (!L.hostAudit) gap('G-HOST', 'InputBox');
    return L.hostAudit.inputBox(prompt, title, def);
  };
  Application.HostDir = function (pattern) {
    if (!L.hostAudit || !L.hostAudit.has('Dir')) gap('G-HOST', 'Dir');
    return L.hostAudit.dir(pattern);
  };
  Application.HostShell = function () {
    if (!L.hostAudit || !L.hostAudit.has('Shell')) gap('G-HOST', 'Shell');
    return L.hostAudit.shell();
  };
  Application.GetOpenFilename = function (filter, filterIndex, title, buttonText, multi) {
    void filter; void filterIndex; void title; void buttonText;
    if (!L.hostAudit) gap('G-HOST', 'GetOpenFilename');
    return L.hostAudit.getOpenFilename(multi);
  };
  Application.GetSaveAsFilename = function () {
    if (!L.hostAudit) gap('G-HOST', 'GetSaveAsFilename');
    return L.hostAudit.getSaveAsFilename();
  };
  Object.defineProperty(Application, 'SlideShowWindows', { enumerable: true, get() { return makeShowWindows(Application); } });

  dom.Application = Application;
  dom.gap = gap;
  dom.reset = function () {
    wraps = [];
    multi = false;
    state.slide = 0;
    state.shapes = [];
    state.text = null;
    state.slideSel = null;
    clip = null;
    if (!L.pres) L.pres = M.newPresentation();
    if (L.hist && L.hist.clear) L.hist.clear();
    wraps.push({ pres: L.pres, name: fileName() });
  };

  const CONSTS = {
    ppSelectionNone: 0, ppSelectionSlides: 1, ppSelectionShapes: 2, ppSelectionText: 3,
    ppAlignLeft: 1, ppAlignCenter: 2, ppAlignRight: 3, ppAlignJustify: 4, ppAlignDistribute: 5,
    ppLayoutTitle: 1, ppLayoutText: 2, ppLayoutTwoColumnText: 3, ppLayoutTable: 4, ppLayoutChart: 8,
    ppLayoutTitleOnly: 11, ppLayoutBlank: 12, ppLayoutTextAndObject: 13, ppLayoutObjectAndText: 14,
    ppLayoutLargeObject: 15, ppLayoutObject: 16, ppLayoutObjectOverText: 19, ppLayoutTextOverObject: 20,
    ppLayoutTextAndTwoObjects: 21, ppLayoutTwoObjectsAndText: 22, ppLayoutFourObjects: 24,
    ppLayoutVerticalText: 25, ppLayoutVerticalTitleAndText: 27, ppLayoutTwoObjects: 29,
    ppLayoutObjectAndTwoObjects: 30, ppLayoutTwoObjectsAndObject: 31, ppLayoutCustom: 32,
    ppPlaceholderTitle: 1, ppPlaceholderBody: 2, ppPlaceholderCenterTitle: 3, ppPlaceholderSubtitle: 4,
    ppPlaceholderVerticalTitle: 5, ppPlaceholderVerticalBody: 6, ppPlaceholderObject: 7, ppPlaceholderChart: 8,
    ppPlaceholderTable: 12, ppPlaceholderSlideNumber: 13, ppPlaceholderFooter: 15, ppPlaceholderDate: 16,
    ppBulletNone: 0, ppBulletUnnumbered: 1, ppBulletNumbered: 2,
    ppAutoSizeNone: 0, ppAutoSizeShapeToFitText: 1,
    msoAutoSizeNone: 0, msoAutoSizeShapeToFitText: 1, msoAutoSizeTextToFitShape: 2,
    ppAnchorTop: 1, ppAnchorMiddle: 3, ppAnchorBottom: 4,
    ppTextOrientationHorizontal: 1,
    msoTrue: -1, msoFalse: 0, msoTriStateMixed: -2,
    msoFileDialogOpen: 1, msoFileDialogSaveAs: 2, msoFileDialogFilePicker: 3, msoFileDialogFolderPicker: 4,
    ppFixedFormatTypePDF: 2, ppFixedFormatTypeXPS: 1,
    msoBringToFront: 0, msoSendToBack: 1, msoBringForward: 2, msoSendBackward: 3,
    msoShapeRectangle: 1, msoShapeDiamond: 4, msoShapeRoundedRectangle: 5, msoShapeOval: 9,
    msoShapeHexagon: 10, msoShapeCross: 11, msoShapeCan: 13, msoShapeCube: 14, msoShapeDonut: 18,
    msoShapeBlockArc: 20, msoShapeHeart: 21, msoShapeLightningBolt: 22, msoShapeSun: 23, msoShapeMoon: 24, msoShapeArc: 25,
    msoAutoShape: 1, msoChart: 3, msoGroup: 6, msoLine: 9, msoPicture: 13, msoPlaceholder: 14, msoTextEffect: 15, msoTextBox: 17, msoTable: 19,
    msoFillSolid: 1, msoFillGradient: 3,
    msoThemeColorDark1: 1, msoThemeColorLight1: 2, msoThemeColorDark2: 3, msoThemeColorLight2: 4,
    msoThemeColorAccent1: 5, msoThemeColorAccent2: 6, msoThemeColorAccent3: 7, msoThemeColorAccent4: 8,
    msoThemeColorAccent5: 9, msoThemeColorAccent6: 10, msoThemeColorHyperlink: 11, msoThemeColorFollowedHyperlink: 12,
    msoThemeColorText1: 13, msoThemeColorBackground1: 14, msoThemeColorText2: 15, msoThemeColorBackground2: 16,
    ppBackground: 1, ppForeground: 2, ppShadow: 3, ppTitle: 4, ppFill: 5, ppAccent: 6, ppAccentAndHyperlink: 7, ppAccentAndFollowedHyperlink: 8,
    ppEffectNone: 0, ppEffectCut: 257, ppEffectCutThroughBlack: 258, ppEffectDissolve: 1537,
    ppEffectWipeLeft: 2817, ppEffectWipeUp: 2818, ppEffectWipeRight: 2819, ppEffectWipeDown: 2820,
    ppEffectFadeSmoothly: 3849, ppEffectPushDown: 3852, ppEffectPushLeft: 3853, ppEffectPushRight: 3854, ppEffectPushUp: 3855,
    ppTransitionSpeedSlow: 1, ppTransitionSpeedMedium: 2, ppTransitionSpeedFast: 3,
    ppShowAll: 1, ppShowSlideRange: 2, ppShowNamedSlideShow: 3,
    ppShowTypeSpeaker: 1, ppShowTypeWindow: 2, ppShowTypeKiosk: 3,
    ppSlideShowManualAdvance: 1, ppSlideShowUseSlideTimings: 2,
  };

  for (const k in CONSTS) {
    dom[k] = CONSTS[k];
    try {
      const d = Object.getOwnPropertyDescriptor(window, k);
      if (!d || d.configurable) Object.defineProperty(window, k, { configurable: true, enumerable: true, writable: true, value: CONSTS[k] });
    } catch (e) { /* leave a non-configurable global alone */ }
  }

  function defineGetter(obj, name, get) {
    try {
      const d = Object.getOwnPropertyDescriptor(obj, name);
      if (d && !d.configurable) return;
      Object.defineProperty(obj, name, { configurable: true, enumerable: true, get });
    } catch (e) { /* ignore */ }
  }
  defineGetter(window, 'Application', () => Application);
  defineGetter(window, 'ActivePresentation', () => Application.ActivePresentation);
  defineGetter(window, 'Presentations', () => Presentations);
  defineGetter(window, 'ActiveWindow', () => ActiveWindow);

  dom.reset();
})();
