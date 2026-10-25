/* Quire — Word object model on the open document.
   Inventory: docs/quire_dom_api_specifications.md */
(function () {
  'use strict';
  const L = window.L;
  const D = L.D;
  const O = L.O;
  const dom = (L.dom = L.dom || {});

  function gap(code, member) {
    const e = new Error(member ? member + ' is not available (' + code + ')' : code);
    e.code = code;
    throw e;
  }

  const Application = {};

  function stamp(obj, parent) {
    obj.Application = Application;
    obj.Parent = parent == null ? Application : parent;
    Object.defineProperty(obj, 'Creator', {
      enumerable: true, configurable: true,
      get() { gap('G-ENGINE', 'Creator'); },
    });
    return obj;
  }

  function defineGetter(obj, name, get) {
    try {
      const d = Object.getOwnPropertyDescriptor(obj, name);
      if (d && !d.configurable) return false;
      Object.defineProperty(obj, name, { configurable: true, enumerable: true, get });
      return true;
    } catch (e) { return false; }
  }

  function storyGets(obj, names) {
    for (const n of names) Object.defineProperty(obj, n, { enumerable: true, configurable: true, get() { gap('G-STORY', n); } });
  }

  function storyFns(obj, names) {
    for (const n of names) obj[n] = function () { gap('G-STORY', n); };
  }

  const RANGE_FNS = ['Move', 'MoveStart', 'MoveEnd', 'MoveWhile', 'MoveUntil', 'MoveStartWhile', 'MoveStartUntil', 'MoveEndWhile', 'MoveEndUntil', 'StartOf', 'EndOf', 'Expand', 'Information', 'GoTo', 'GoToNext', 'GoToPrevious', 'Next', 'Previous'];

  function charactersOf(parent, getEnds) {
    const col = function (index) { return col.Item(index); };
    stamp(col, parent);
    col.Item = function (index) {
      const e = getEnds ? getEnds() : null;
      if (!e) missing();
      const n = index | 0;
      if (n < 1) missing();
      let left = n;
      const paras = D.parasBetween(D.doc, e[0].p, e[1].p);
      for (let i = 0; i < paras.length; i++) {
        const p = paras[i];
        const from = i === 0 ? e[0].o : 0;
        const to = i === paras.length - 1 ? e[1].o : D.plen(p);
        const len = Math.max(0, to - from);
        if (left <= len) {
          const o = from + left - 1;
          return makeRange(parent, { kind: 'span', a: D.pos(p, o), f: D.pos(p, Math.min(o + 1, D.plen(p))) });
        }
        left -= len;
      }
      missing();
    };
    Object.defineProperty(col, 'Count', {
      enumerable: true,
      get() {
        const e = getEnds ? getEnds() : null;
        return e ? spanLen(e[0], e[1]) : 0;
      },
    });
    return col;
  }

  function spanLen(a, b) {
    [a, b] = D.order(D.doc, a, b);
    if (a.p === b.p) return Math.max(0, b.o - a.o);
    let n = D.plen(a.p) - a.o;
    const paras = D.parasBetween(D.doc, a.p, b.p);
    for (let i = 1; i < paras.length - 1; i++) n += D.plen(paras[i]);
    return n + b.o;
  }

  function missing() {
    const e = new Error('The requested member of the collection does not exist.');
    e.code = 'missing';
    throw e;
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

  function selFrom(res) {
    if (!res) return null;
    if (res.a && res.f) return { a: res.a, f: res.f };
    if (res.p && res.o != null) return { a: res, f: res };
    return null;
  }

  function edit(label, fn) {
    let moved = null;
    const res = D.tx(label, () => {
      const out = fn();
      const sel = selFrom(out);
      if (sel) {
        caret = sel;
        if (L.ed) L.ed.sel = { a: sel.a, f: sel.f };
        moved = sel;
      }
      return out;
    });
    if (moved && L.ed && typeof L.ed.setSel === 'function') L.ed.setSel(moved.a, moved.f, { noScroll: true });
    return res;
  }

  function stripOneBreak(s) {
    s = String(s == null ? '' : s);
    if (s.endsWith('\r') || s.endsWith('\n')) return s.slice(0, -1);
    return s;
  }

  function writeParagraphBody(p, raw) {
    const s = stripOneBreak(raw);
    edit('Text', () => {
      D.touch(p);
      p.runs = [];
      let pos = D.pos(p, 0);
      const parts = s.split(/\r\n|\r|\n/);
      parts.forEach((part, k) => {
        if (k) pos = O.insertItem(pos, D.item('br', { type: 'line' }));
        if (part) pos = O.insertText(pos, part);
      });
      return holdCaret();
    });
  }

  const ALIGN_JC = ['left', 'center', 'right', 'both', 'distribute'];

  function alignOf(p) {
    const jc = p.pPr && p.pPr.jc;
    if (jc === 'center') return 1;
    if (jc === 'right') return 2;
    if (jc === 'both' || jc === 'justify') return 3;
    if (jc === 'distribute') return 4;
    return 0;
  }

  function writeAlign(p, v) {
    const jc = ALIGN_JC[v | 0];
    if (!jc) gap('G-ENGINE', 'Alignment');
    edit('Alignment', () => {
      O.setParaProps([p], (pPr) => Object.assign({}, pPr, { jc }));
      return holdCaret();
    });
  }

  function styleName(model, p) {
    const id = (p.pPr && p.pPr.style) || 'Normal';
    const st = model.styles[id];
    return st ? D.styleDisplayName(st) : id;
  }

  function writeStyle(p, value) {
    const model = D.doc;
    let id = null;
    if (value && typeof value === 'object') value = value.NameLocal || value.Name || value._id;
    if (value != null) {
      const found = D.findStyleByName(model, String(value));
      if (found) id = found.id;
      else if (model.styles[value]) id = value;
    }
    if (!id) {
      const e = new Error('The style does not exist');
      e.code = 'missing';
      throw e;
    }
    edit('Style', () => {
      O.applyStyle(D.pos(p, 0), D.pos(p, D.plen(p)), id);
      return holdCaret();
    });
  }

  const WD_UNDEFINED = 9999999;
  let pendingBox = { pending: null };

  function pendHost() { return L.ed || pendingBox; }

  function clearPending() { pendHost().pending = null; }

  function ensureCaret() {
    if (caret && caret.a && caret.a.p && D.info(D.doc, caret.a.p)) return caret;
    if (L.ed && L.ed.sel && L.ed.sel.a && D.info(D.doc, L.ed.sel.a.p)) {
      caret = { a: L.ed.sel.a, f: L.ed.sel.f };
      return caret;
    }
    const p = D.firstPara(D.doc.main.blocks);
    caret = { a: D.pos(p, 0), f: D.pos(p, 0) };
    return caret;
  }

  function endsOf(anchor) {
    if (!anchor) return null;
    if (anchor.kind === 'paragraph') {
      const p = anchor.block;
      return [D.pos(p, 0), D.pos(p, D.plen(p))];
    }
    if (anchor.kind === 'cell') {
      const paras = (anchor.cell.blocks || []).filter((b) => b.t === 'p');
      if (!paras.length) return null;
      const last = paras[paras.length - 1];
      return [D.pos(paras[0], 0), D.pos(last, D.plen(last))];
    }
    if (anchor.kind === 'span') return D.order(D.doc, anchor.a, anchor.f);
    if (anchor.kind === 'selection') {
      const s = ensureCaret();
      return D.order(D.doc, s.a, s.f);
    }
    if (anchor.kind === 'story' || anchor.kind === 'section') {
      const paras = parasOfAnchor(anchor);
      if (!paras.length) return null;
      const last = paras[paras.length - 1];
      return [D.pos(paras[0], 0), D.pos(last, D.plen(last))];
    }
    if (anchor.kind === 'table') {
      const table = anchor.table;
      const first = table.rows[0] && table.rows[0].cells[0] && D.firstPara(table.rows[0].cells[0].blocks);
      if (!first) return null;
      const after = paragraphAfterTable(table);
      if (after) return [D.pos(first, 0), D.pos(after, 0)];
      const lastRow = table.rows[table.rows.length - 1];
      const lastCell = lastRow.cells[lastRow.cells.length - 1];
      const last = (lastCell.blocks || []).filter((b) => b.t === 'p').pop();
      return [D.pos(first, 0), D.pos(last || first, D.plen(last || first))];
    }
    return null;
  }

  /* The story index counts each paragraph's text plus its paragraph mark, in
     reading order. It is Quire's own order: a table's end is the following
     paragraph, not Word's cell-mark count. */
  function paragraphAfterTable(table) {
    const info = D.tblInfo(D.doc, table);
    const blocks = info && info.cont && info.cont.blocks;
    if (!blocks) return null;
    const i = blocks.indexOf(table);
    if (i < 0) return null;
    const next = blocks[i + 1];
    if (!next) return null;
    if (next.t === 'p') return next;
    if (next.t === 'tbl') return D.firstPara(next.rows[0].cells[0].blocks);
    return null;
  }

  function storyPair(anchor) {
    const e = endsOf(anchor);
    if (!e || !D.doc) return null;
    const st = storyTextOf(D.storyOf(D.doc, e[0].p));
    return { st, a: offInStory(st, e[0]), b: offInStory(st, e[1]) };
  }

  function retarget(anchor, a, b) {
    if (!anchor) return;
    if (anchor.kind === 'selection') { placeSpan(a, b); return; }
    anchor.kind = 'span';
    anchor.a = a;
    anchor.f = b;
  }

  function layoutUnit(unit) { return unit === 5 || unit === 7 || unit === 9 || unit === 11; }

  function stepOff(st, off, unit, dir) {
    const len = st.text.length;
    if (unit === 6) { const n = dir < 0 ? 0 : len; return n === off ? null : n; }
    if (unit === 1) { const n = off + dir; return n < 0 || n > len ? null : n; }
    if (unit === 4) {
      const pos = posInStory(st, len ? Math.min(Math.max(off, 0), len - 1) : 0);
      const i = st.paras.indexOf(pos.p);
      if (i < 0) return null;
      if (dir > 0) return i + 1 < st.starts.length ? st.starts[i + 1] : (off < len ? len : null);
      const start = st.starts[i];
      if (off > start) return start;
      return i > 0 ? st.starts[i - 1] : (off === 0 ? null : 0);
    }
    if (unit === 2 || unit === 3) {
      const extFn = unit === 2 ? O.wordExtent : O.sentenceExtent;
      if (typeof extFn !== 'function') gap('G-ENGINE', 'Move');
      const pos = posInStory(st, len ? Math.min(Math.max(off, 0), len - 1) : 0);
      const ext = extFn(pos);
      if (!ext) return null;
      const s = offInStory(st, ext[0]);
      const e = offInStory(st, ext[1]);
      if (dir > 0) {
        if (off < e) return e;
        if (e >= len) return null;
        const n = extFn(posInStory(st, Math.min(e, len - 1)));
        const ne = offInStory(st, n[1]);
        return ne === off ? null : ne;
      }
      if (off > s) return s;
      if (s <= 0) return null;
      return offInStory(st, extFn(posInStory(st, s - 1))[0]);
    }
    gap('G-STORY', 'Move');
  }

  function unitEdge(st, off, unit, end) {
    const len = st.text.length;
    if (unit === 6) return end ? len : 0;
    if (unit === 4) {
      const pos = posInStory(st, len ? Math.min(Math.max(off, 0), len - 1) : 0);
      const i = st.paras.indexOf(pos.p);
      if (i < 0) return off;
      return end ? (i + 1 < st.starts.length ? st.starts[i + 1] : len) : st.starts[i];
    }
    if (unit === 1) return off;
    if (unit === 2 || unit === 3) {
      const extFn = unit === 2 ? O.wordExtent : O.sentenceExtent;
      const pos = posInStory(st, len ? Math.min(Math.max(off, 0), len - 1) : 0);
      const ext = extFn(pos);
      return offInStory(st, end ? ext[1] : ext[0]);
    }
    gap('G-STORY', 'Move');
  }

  function attachStory(obj, getAnchor) {
    function pair() { return storyPair(getAnchor()); }
    function writeOff(start, end) {
      const anchor = getAnchor();
      const p = pair();
      const st = p ? p.st : (D.doc ? storyTextOf(D.doc.main) : null);
      if (!st || !anchor) return;
      retarget(anchor, posInStory(st, start | 0), posInStory(st, end == null ? (start | 0) : (end | 0)));
    }
    Object.defineProperty(obj, 'Start', {
      enumerable: true, configurable: true,
      get() { const p = pair(); return p ? p.a : 0; },
      set(v) { const p = pair(); if (p) writeOff(v | 0, p.b); },
    });
    Object.defineProperty(obj, 'End', {
      enumerable: true, configurable: true,
      get() { const p = pair(); return p ? p.b : 0; },
      set(v) { const p = pair(); if (p) writeOff(p.a, v | 0); },
    });
    obj.SetRange = function (start, end) { writeOff(start | 0, end == null ? (start | 0) : (end | 0)); };
    function shift(which, unit, count) {
      unit = unit == null ? 1 : (unit | 0);
      count = count == null ? 1 : (count | 0);
      if (layoutUnit(unit)) gap('G-STORY', 'Move');
      const p = pair();
      if (!p) return 0;
      const dir = count < 0 ? -1 : 1;
      const n = Math.abs(count);
      let a = p.a, b = p.b, moved = 0;
      for (let i = 0; i < n; i++) {
        const at = which === 'start' ? a : which === 'end' ? b : (dir < 0 ? a : b);
        const next = stepOff(p.st, at, unit, dir);
        if (next == null) break;
        if (which === 'start') a = next;
        else if (which === 'end') b = next;
        else a = b = next;
        moved++;
      }
      writeOff(Math.min(a, b), Math.max(a, b));
      return moved * dir;
    }
    obj.Move = function (unit, count) { return shift('both', unit, count); };
    obj.MoveStart = function (unit, count) { return shift('start', unit, count); };
    obj.MoveEnd = function (unit, count) { return shift('end', unit, count); };
    obj.StartOf = function (unit, extend) {
      unit = unit == null ? 6 : (unit | 0);
      if (layoutUnit(unit)) gap('G-STORY', 'StartOf');
      const p = pair();
      if (!p) return 0;
      const edge = unitEdge(p.st, p.a, unit, false);
      writeOff(edge, extend ? p.b : edge);
      return edge === p.a ? 0 : 1;
    };
    obj.EndOf = function (unit, extend) {
      unit = unit == null ? 6 : (unit | 0);
      if (layoutUnit(unit)) gap('G-STORY', 'EndOf');
      const p = pair();
      if (!p) return 0;
      const edge = unitEdge(p.st, p.b, unit, true);
      writeOff(extend ? p.a : edge, edge);
      return edge === p.b ? 0 : 1;
    };
    obj.Expand = function (unit) {
      unit = unit == null ? 1 : (unit | 0);
      if (layoutUnit(unit)) gap('G-STORY', 'Expand');
      const p = pair();
      if (!p) return 0;
      writeOff(unitEdge(p.st, p.a, unit, false), unitEdge(p.st, Math.max(p.a, p.b - 1), unit, true));
      return 1;
    };
    storyFns(obj, ['MoveWhile', 'MoveUntil', 'MoveStartWhile', 'MoveStartUntil', 'MoveEndWhile', 'MoveEndUntil', 'Information', 'GoTo', 'GoToNext', 'GoToPrevious', 'Next', 'Previous']);
  }

  function moveSelection(dir, unit, count, extend) {
    unit = unit == null ? 1 : (unit | 0);
    count = count == null ? 1 : (count | 0);
    if (layoutUnit(unit)) gap('G-STORY', 'Move');
    const p = storyPair({ kind: 'selection' });
    if (!p) return 0;
    const stepDir = dir * (count < 0 ? -1 : 1);
    const n = Math.abs(count);
    let a = p.a, b = p.b, moved = 0;
    if (!extend && a !== b) { if (stepDir < 0) b = a; else a = b; }
    for (let i = 0; i < n; i++) {
      if (extend) {
        const at = stepDir < 0 ? a : b;
        const next = stepOff(p.st, at, unit, stepDir);
        if (next == null) break;
        if (stepDir < 0) a = next; else b = next;
      } else {
        const next = stepOff(p.st, a, unit, stepDir);
        if (next == null) break;
        a = b = next;
      }
      moved++;
    }
    const lo = Math.min(a, b), hi = Math.max(a, b);
    placeSpan(posInStory(p.st, lo), posInStory(p.st, hi));
    return moved;
  }

  const LINE_TO = { nil: 0, none: 0, single: 1, dotted: 2, dashSmallGap: 3, dashed: 4, dotDash: 5, dotDotDash: 6, double: 7, triple: 8, thinThickSmallGap: 9, thickThinSmallGap: 10, thinThickThinSmallGap: 11, wave: 18, doubleWave: 19, dashDotStroked: 20, threeDEmboss: 21, threeDEngrave: 22, outset: 23, inset: 24 };
  const LINE_FROM = { 0: 'nil', 1: 'single', 2: 'dotted', 3: 'dashSmallGap', 4: 'dashed', 5: 'dotDash', 6: 'dotDotDash', 7: 'double', 8: 'triple', 9: 'thinThickSmallGap', 10: 'thickThinSmallGap', 11: 'thinThickThinSmallGap', 18: 'wave', 19: 'doubleWave', 20: 'dashDotStroked', 21: 'threeDEmboss', 22: 'threeDEngrave', 23: 'outset', 24: 'inset' };
  const CI_HEX = ['', '000000', '0000FF', '00FFFF', '00FF00', 'FF00FF', 'FF0000', 'FFFF00', 'FFFFFF', '000080', '008080', '008000', '800080', '800000', '808000', '808080', 'C0C0C0'];
  const TEX_TO = { clear: 0, nil: 0, none: 0, solid: 1000, pct5: 50, pct10: 100, pct12: 125, pct15: 150, pct20: 200, pct25: 250, pct30: 300, pct35: 350, pct37: 375, pct40: 400, pct45: 450, pct50: 500, pct55: 550, pct60: 600, pct62: 625, pct65: 650, pct70: 700, pct75: 750, pct80: 800, pct85: 850, pct87: 875, pct90: 900, pct95: 950, horzStripe: -1, vertStripe: -2, reverseDiagStripe: -3, diagStripe: -4, horzCross: -5, diagCross: -6, thinHorzStripe: -7, thinVertStripe: -8, thinReverseDiagStripe: -9, thinDiagStripe: -10, thinHorzCross: -11, thinDiagCross: -12 };
  const TEX_FROM = {};
  for (const texName in TEX_TO) if (TEX_FROM[TEX_TO[texName]] == null) TEX_FROM[TEX_TO[texName]] = texName === 'nil' || texName === 'none' ? 'clear' : texName;

  function lineStyleOf(edge) {
    if (!edge || !edge.val || edge.val === 'nil' || edge.val === 'none') return 0;
    return LINE_TO[edge.val] != null ? LINE_TO[edge.val] : 1;
  }
  function edgeVisible(edge) { return !!(edge && edge.val && edge.val !== 'nil' && edge.val !== 'none'); }
  function colorIndexOf(hex) {
    if (!hex || hex === 'auto') return 0;
    const h = String(hex).replace('#', '').toUpperCase();
    const i = CI_HEX.indexOf(h);
    return i > 0 ? i : 0;
  }

  function bordersOf(parent, spec) {
    const borders = function (i) { return borders.Item(i); };
    stamp(borders, parent);
    const byIndex = {};
    for (const edge of spec.edges) byIndex[edge.index] = edge.key;
    function read() { return spec.getBag() || {}; }
    function write(bag, reason) { spec.setBag(bag, reason); }
    function paint(keys, edge) {
      const bag = Object.assign({}, read());
      for (const key of keys) bag[key] = edge;
      write(bag);
    }
    borders.Item = function (index) {
      const key = byIndex[index | 0];
      if (!key) gap('G-ENGINE', 'Borders');
      const side = stamp({}, borders);
      function current() { return read()[key] || null; }
      Object.defineProperty(side, 'LineStyle', {
        enumerable: true, configurable: true,
        get() { return lineStyleOf(current()); },
        set(v) {
          const n = v | 0;
          const prev = current() || {};
          paint([key], n ? { val: LINE_FROM[n] || 'single', sz: prev.sz || 0.5, color: prev.color || 'auto', space: prev.space, shadow: prev.shadow } : { val: 'nil' });
        },
      });
      Object.defineProperty(side, 'LineWidth', {
        enumerable: true, configurable: true,
        get() { const e = current(); return e && e.sz ? Math.round(e.sz * 8) : 0; },
        set(v) { const prev = current() || { val: 'single', color: 'auto' }; paint([key], Object.assign({}, prev, { val: prev.val && prev.val !== 'nil' ? prev.val : 'single', sz: (Number(v) || 0) / 8 })); },
      });
      Object.defineProperty(side, 'Color', {
        enumerable: true, configurable: true,
        get() { return colorToLong(current() && current().color); },
        set(v) { const prev = current() || { val: 'single', sz: 0.5 }; paint([key], Object.assign({}, prev, { color: longToColor(v) || 'auto' })); },
      });
      Object.defineProperty(side, 'ColorIndex', {
        enumerable: true, configurable: true,
        get() { return colorIndexOf(current() && current().color); },
        set(v) { const hex = CI_HEX[v | 0]; const prev = current() || { val: 'single', sz: 0.5 }; paint([key], Object.assign({}, prev, { color: hex || 'auto' })); },
      });
      Object.defineProperty(side, 'Visible', {
        enumerable: true,
        get() { return edgeVisible(current()) ? -1 : 0; },
        set(v) { side.LineStyle = triWrite(v) ? 1 : 0; },
      });
      Object.defineProperty(side, 'Inside', { enumerable: true, get() { return spec.inside && spec.inside.indexOf(key) >= 0 ? -1 : 0; } });
      return side;
    };
    const enableKeys = spec.enable || spec.edges.map((e) => e.key);
    Object.defineProperty(borders, 'Enable', {
      enumerable: true, configurable: true,
      get() {
        const bag = read();
        const bits = enableKeys.map((key) => edgeVisible(bag[key]));
        if (!bits.length || bits.every((v) => !v)) return 0;
        if (bits.every(Boolean)) return -1;
        return WD_UNDEFINED;
      },
      set(v) {
        const on = triWrite(v);
        if (on === 'skip') return;
        const bag = Object.assign({}, read());
        const edge = on ? { val: 'single', sz: 0.5, color: 'auto' } : { val: 'nil' };
        for (const key of enableKeys) bag[key] = Object.assign({}, edge);
        write(bag, 'enable');
      },
    });
    function groupProp(name, keys, field, from, to) {
      Object.defineProperty(borders, name, {
        enumerable: true, configurable: true,
        get() {
          if (!keys.length) return 0;
          const bag = read();
          const vals = keys.map((key) => from(bag[key]));
          return vals.every((v) => v === vals[0]) ? vals[0] : WD_UNDEFINED;
        },
        set(v) { if (v === WD_UNDEFINED) return; paint(keys, to(v, read()[keys[0]])); },
      });
    }
    const outside = spec.outside || [];
    const inside = spec.inside || [];
    groupProp('OutsideLineStyle', outside, 'style', lineStyleOf, (v, prev) => ((v | 0) ? { val: LINE_FROM[v | 0] || 'single', sz: (prev && prev.sz) || 0.5, color: (prev && prev.color) || 'auto' } : { val: 'nil' }));
    groupProp('InsideLineStyle', inside, 'style', lineStyleOf, (v, prev) => ((v | 0) ? { val: LINE_FROM[v | 0] || 'single', sz: (prev && prev.sz) || 0.5, color: (prev && prev.color) || 'auto' } : { val: 'nil' }));
    groupProp('OutsideLineWidth', outside, 'w', (e) => (e && e.sz ? Math.round(e.sz * 8) : 0), (v, prev) => Object.assign({ val: 'single', color: 'auto' }, prev, { sz: (Number(v) || 0) / 8, val: (prev && prev.val && prev.val !== 'nil') ? prev.val : 'single' }));
    groupProp('InsideLineWidth', inside, 'w', (e) => (e && e.sz ? Math.round(e.sz * 8) : 0), (v, prev) => Object.assign({ val: 'single', color: 'auto' }, prev, { sz: (Number(v) || 0) / 8, val: (prev && prev.val && prev.val !== 'nil') ? prev.val : 'single' }));
    groupProp('OutsideColor', outside, 'c', (e) => colorToLong(e && e.color), (v, prev) => Object.assign({ val: 'single', sz: 0.5 }, prev, { color: longToColor(v) || 'auto' }));
    groupProp('InsideColor', inside, 'c', (e) => colorToLong(e && e.color), (v, prev) => Object.assign({ val: 'single', sz: 0.5 }, prev, { color: longToColor(v) || 'auto' }));
    Object.defineProperty(borders, 'Shadow', {
      enumerable: true, configurable: true,
      get() {
        const bag = read();
        const bits = spec.edges.map((e) => (bag[e.key] && bag[e.key].shadow ? -1 : 0));
        return triRead(bits.map((b) => b === -1));
      },
      set(v) {
        const t = triWrite(v);
        if (t === 'skip') return;
        const bag = Object.assign({}, read());
        for (const edge of spec.edges) bag[edge.key] = Object.assign({ val: 'single', sz: 0.5, color: 'auto' }, bag[edge.key], { shadow: !!t });
        write(bag);
      },
    });
    function spaceProp(name, key) {
      Object.defineProperty(borders, name, {
        enumerable: true, configurable: true,
        get() {
          const bag = read();
          const vals = spec.edges.map((e) => (bag[e.key] && bag[e.key].space) || 0);
          return vals.every((n) => n === vals[0]) ? vals[0] : WD_UNDEFINED;
        },
        set(v) {
          const bag = Object.assign({}, read());
          for (const edge of spec.edges) if (!key || edge.key === key) bag[edge.key] = Object.assign({ val: 'single', sz: 0.5, color: 'auto' }, bag[edge.key], { space: Number(v) || 0 });
          write(bag);
        },
      });
    }
    const spaceKey = { DistanceFromTop: 'top', DistanceFromLeft: 'left', DistanceFromBottom: 'bottom', DistanceFromRight: 'right' };
    for (const name in spaceKey) if (byIndex[-1] && (name === 'DistanceFromTop' ? byIndex[-1] : true)) spaceProp(name, spaceKey[name] && byIndex[{ DistanceFromTop: -1, DistanceFromLeft: -2, DistanceFromBottom: -3, DistanceFromRight: -4 }[name]] ? spaceKey[name] : null);
    Object.defineProperty(borders, 'Count', { enumerable: true, get() { return spec.edges.length; } });
    Object.defineProperty(borders, 'HasHorizontal', { enumerable: true, get() { const bag = read(); return (bag.insideH || bag.between || bag.top) && edgeVisible(bag.insideH || bag.between || bag.top) ? -1 : 0; } });
    Object.defineProperty(borders, 'HasVertical', { enumerable: true, get() { const bag = read(); return edgeVisible(bag.insideV || bag.left) ? -1 : 0; } });
    if (spec.page) {
      Object.defineProperty(borders, 'DistanceFrom', {
        enumerable: true, configurable: true,
        get() { return read().offsetFrom === 'page' ? 0 : 1; },
        set(v) { const bag = Object.assign({}, read()); bag.offsetFrom = (v | 0) === 0 ? 'page' : 'text'; write(bag); },
      });
      Object.defineProperty(borders, 'AlwaysInFront', {
        enumerable: true, configurable: true,
        get() { return read().zOrder === 'back' ? 0 : -1; },
        set(v) { const bag = Object.assign({}, read()); bag.zOrder = triWrite(v) ? 'front' : 'back'; write(bag); },
      });
      function pageFlag(name, token) {
        Object.defineProperty(borders, name, {
          enumerable: true, configurable: true,
          get() { const d = read().display || 'allPages'; return d === 'allPages' || d === token ? -1 : 0; },
          set(v) {
            const bag = Object.assign({}, read());
            const on = !!triWrite(v);
            const other = token === 'firstPage' ? 'notFirstPage' : 'firstPage';
            const cur = bag.display || 'allPages';
            if (on && (cur === 'allPages' || cur === token)) bag.display = 'allPages';
            else if (on) bag.display = token;
            else if (cur === 'allPages') bag.display = other;
            else bag.display = token === 'firstPage' ? 'notFirstPage' : 'firstPage';
            write(bag);
          },
        });
      }
      pageFlag('EnableFirstPageInSection', 'firstPage');
      pageFlag('EnableOtherPagesInSection', 'notFirstPage');
    } else {
      denyGet(borders, ['DistanceFrom', 'AlwaysInFront', 'EnableFirstPageInSection', 'EnableOtherPagesInSection'], 'G-ENGINE');
    }
    denyGet(borders, ['ApplyPageBordersToAllSections', 'JoinBorders', 'SurroundHeader', 'SurroundFooter', 'ArtStyle', 'ArtWidth'], 'G-ENGINE');
    return borders;
  }

  function shadingOf(parent, getShd, setShd) {
    const sh = stamp({}, parent);
    function cur() { return getShd() || {}; }
    Object.defineProperty(sh, 'BackgroundPatternColor', {
      enumerable: true, configurable: true,
      get() { return colorToLong(cur().fill); },
      set(v) { const prev = cur(); setShd(Object.assign({}, prev, { val: prev.val || 'clear', fill: longToColor(v) || 'auto' })); },
    });
    Object.defineProperty(sh, 'ForegroundPatternColor', {
      enumerable: true, configurable: true,
      get() { return colorToLong(cur().color); },
      set(v) { const prev = cur(); setShd(Object.assign({}, prev, { val: prev.val || 'clear', color: longToColor(v) || 'auto' })); },
    });
    Object.defineProperty(sh, 'BackgroundPatternColorIndex', {
      enumerable: true, configurable: true,
      get() { return colorIndexOf(cur().fill); },
      set(v) { const prev = cur(); setShd(Object.assign({}, prev, { val: prev.val || 'clear', fill: CI_HEX[v | 0] || 'auto' })); },
    });
    Object.defineProperty(sh, 'ForegroundPatternColorIndex', {
      enumerable: true, configurable: true,
      get() { return colorIndexOf(cur().color); },
      set(v) { const prev = cur(); setShd(Object.assign({}, prev, { val: prev.val || 'clear', color: CI_HEX[v | 0] || 'auto' })); },
    });
    Object.defineProperty(sh, 'Texture', {
      enumerable: true, configurable: true,
      get() { const val = cur().val || 'clear'; return TEX_TO[val] != null ? TEX_TO[val] : 0; },
      set(v) {
        const prev = cur();
        const name = typeof v === 'string' ? v : (TEX_FROM[v | 0] || 'clear');
        setShd(Object.assign({}, prev, { val: name }));
      },
    });
    return sh;
  }

  function triWrite(v) {
    if (v === WD_UNDEFINED) return 'skip';
    return v === true || v === -1 || v === 1;
  }

  function triRead(vals) {
    if (!vals.length) return 0;
    const bits = vals.map((v) => (v === true || v === -1 || v === 1 ? -1 : 0));
    return bits.every((b) => b === bits[0]) ? bits[0] : WD_UNDEFINED;
  }

  function runValues(a, b, key) {
    if (D.eqPos(a, b)) {
      let r = O.rPrAt(a) || {};
      const pend = pendHost().pending;
      if (pend) r = D.mergeRPr(r, pend);
      return [r[key]];
    }
    const vals = [];
    for (const p of D.parasBetween(D.doc, a.p, b.p)) {
      const s = p === a.p ? a.o : 0;
      const e = p === b.p ? b.o : D.plen(p);
      let pos = 0;
      for (const it of p.runs) {
        const n = D.ilen(it);
        const s0 = pos;
        pos += n;
        if (!n || pos <= s || s0 >= e) continue;
        if (it.rPr && it.rPr.del) continue;
        vals.push(D.rProps(D.doc, p, it, null)[key]);
      }
    }
    return vals;
  }

  function applyRun(a, b, props) {
    if (D.eqPos(a, b)) {
      const host = pendHost();
      host.pending = Object.assign({}, host.pending || {});
      for (const k in props) {
        if (props[k] === undefined || props[k] === null) delete host.pending[k];
        else host.pending[k] = props[k];
      }
      return;
    }
    edit('Font', () => {
      O.setRunProps(a, b, props);
      return caret;
    });
  }

  function placeCaret(pos) {
    caret = { a: pos, f: pos };
    if (L.ed) {
      L.ed.sel = caret;
      if (typeof L.ed.setSel === 'function') L.ed.setSel(pos, pos, { noScroll: true });
    }
  }

  /* A Range edit must not relocate the Selection. Clamp an offset the edit shortened. */
  function holdCaret() {
    if (!caret || !caret.a || !caret.a.p || !D.info(D.doc, caret.a.p)) return null;
    function clamp(pos) {
      if (!pos || !pos.p || !D.info(D.doc, pos.p)) return null;
      const n = D.plen(pos.p);
      let o = pos.o | 0;
      if (o < 0) o = 0;
      if (o > n) o = n;
      return o === pos.o ? pos : D.pos(pos.p, o);
    }
    const a = clamp(caret.a);
    const f = clamp(caret.f);
    if (!a || !f) return null;
    return a === caret.a && f === caret.f ? caret : { a: a, f: f };
  }

  function longToColor(n) {
    if (n === -16777216 || n == null) return undefined;
    const r = n & 255, g = (n >> 8) & 255, b = (n >> 16) & 255;
    return [r, g, b].map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase();
  }

  function colorToLong(c) {
    if (!c || c === 'auto') return -16777216;
    const h = String(c).replace('#', '');
    if (h.length < 6) return -16777216;
    return parseInt(h.slice(0, 2), 16) | (parseInt(h.slice(2, 4), 16) << 8) | (parseInt(h.slice(4, 6), 16) << 16);
  }

  const U_FROM = { 0: false, 1: 'single', 2: 'words', 3: 'double', 4: 'dotted', 6: 'thick', 7: 'dash', 9: 'dotDash', 10: 'dotDotDash', 11: 'wave' };
  const U_TO = { single: 1, words: 2, double: 3, dotted: 4, thick: 6, dash: 7, dotDash: 9, dotDotDash: 10, wave: 11, wavy: 11 };
  const HL_FROM = { 0: false, 1: 'black', 2: 'blue', 3: 'cyan', 4: 'green', 5: 'magenta', 6: 'red', 7: 'yellow', 8: 'white', 9: 'darkBlue', 10: 'darkCyan', 11: 'darkGreen', 12: 'darkMagenta', 13: 'darkRed', 14: 'darkYellow', 15: 'darkGray', 16: 'lightGray' };
  const HL_TO = { black: 1, blue: 2, cyan: 3, green: 4, magenta: 5, red: 6, yellow: 7, white: 8, darkBlue: 9, darkCyan: 10, darkGreen: 11, darkMagenta: 12, darkRed: 13, darkYellow: 14, darkGray: 15, lightGray: 16 };
  const TRI_KEY = { Bold: 'b', Italic: 'i', BoldBi: 'bCs', ItalicBi: 'iCs', StrikeThrough: 'strike', DoubleStrikeThrough: 'dstrike', AllCaps: 'caps', SmallCaps: 'smallCaps', Hidden: 'hidden', Outline: 'outline', Shadow: 'shadow', Emboss: 'emboss', Engrave: 'imprint' };

  function makeFont(parent, getEnds, styleBox) {
    const font = stamp({}, parent);
    function readKey(key) {
      if (styleBox) return [styleBox.style.rPr ? styleBox.style.rPr[key] : undefined];
      const e = getEnds();
      return e ? runValues(e[0], e[1], key) : [];
    }
    function writeKey(key, val) {
      if (styleBox) {
        edit('Style', () => {
          D.touchKey(styleBox.model, 'styles');
          const r = Object.assign({}, styleBox.style.rPr || {});
          if (val === false || val == null) delete r[key];
          else r[key] = val;
          styleBox.style.rPr = r;
          D.stylesChanged();
          return caret;
        });
        return;
      }
      const e = getEnds();
      if (e) applyRun(e[0], e[1], { [key]: val });
    }
    for (const name in TRI_KEY) {
      const key = TRI_KEY[name];
      Object.defineProperty(font, name, {
        enumerable: true, configurable: true,
        get() { return triRead(readKey(key)); },
        set(v) { const t = triWrite(v); if (t === 'skip') return; writeKey(key, t ? true : false); },
      });
    }
    function vertProp(which) {
      Object.defineProperty(font, which === 'superscript' ? 'Superscript' : 'Subscript', {
        enumerable: true, configurable: true,
        get() { return triRead(readKey('vert').map((v) => v === which)); },
        set(v) { const t = triWrite(v); if (t === 'skip') return; writeKey('vert', t ? which : false); },
      });
    }
    vertProp('superscript');
    vertProp('subscript');
    function numProp(name, key) {
      Object.defineProperty(font, name, {
        enumerable: true, configurable: true,
        get() {
          const vals = readKey(key);
          if (!vals.length) return 0;
          return vals.every((v) => v === vals[0]) ? (vals[0] == null ? 0 : vals[0]) : WD_UNDEFINED;
        },
        set(v) { if (v === WD_UNDEFINED) return; writeKey(key, v); },
      });
    }
    numProp('Size', 'sz');
    numProp('SizeBi', 'szCs');
    numProp('Spacing', 'spacing');
    numProp('Position', 'pos');
    numProp('Scaling', 'w');
    numProp('Kerning', 'kern');
    function nameProp(name, key) {
      Object.defineProperty(font, name, {
        enumerable: true, configurable: true,
        get() {
          const vals = readKey(key).map((v) => v || '');
          if (!vals.length) return '';
          return vals.every((v) => v === vals[0]) ? vals[0] : WD_UNDEFINED;
        },
        set(v) { if (v === WD_UNDEFINED) return; writeKey(key, v ? String(v) : false); },
      });
    }
    nameProp('Name', 'font');
    nameProp('NameAscii', 'font');
    nameProp('NameFarEast', 'fontEA');
    nameProp('NameBi', 'fontCS');
    Object.defineProperty(font, 'Underline', {
      enumerable: true, configurable: true,
      get() {
        const vals = readKey('u').map((u) => (u && U_TO[u] != null ? U_TO[u] : 0));
        if (!vals.length) return 0;
        return vals.every((v) => v === vals[0]) ? vals[0] : WD_UNDEFINED;
      },
      set(v) { if (v === WD_UNDEFINED) return; writeKey('u', U_FROM[v | 0] == null ? false : U_FROM[v | 0]); },
    });
    Object.defineProperty(font, 'Color', {
      enumerable: true, configurable: true,
      get() {
        const vals = readKey('color').map((c) => colorToLong(c));
        if (!vals.length) return -16777216;
        return vals.every((v) => v === vals[0]) ? vals[0] : WD_UNDEFINED;
      },
      set(v) { if (v === WD_UNDEFINED) return; writeKey('color', longToColor(v) || false); },
    });
    Object.defineProperty(font, 'UnderlineColor', {
      enumerable: true, configurable: true,
      get() {
        const vals = readKey('uColor').map((c) => colorToLong(c));
        if (!vals.length) return -16777216;
        return vals.every((v) => v === vals[0]) ? vals[0] : WD_UNDEFINED;
      },
      set(v) { if (v === WD_UNDEFINED) return; writeKey('uColor', longToColor(v) || false); },
    });
    Object.defineProperty(font, 'HighlightColorIndex', {
      enumerable: true, configurable: true,
      get() {
        const vals = readKey('hl').map((h) => (h && HL_TO[h] != null ? HL_TO[h] : 0));
        if (!vals.length) return 0;
        return vals.every((v) => v === vals[0]) ? vals[0] : WD_UNDEFINED;
      },
      set(v) { if (v === WD_UNDEFINED) return; writeKey('hl', HL_FROM[v | 0] == null ? false : HL_FROM[v | 0]); },
    });
    Object.defineProperty(font, 'EmphasisMark', {
      enumerable: true, configurable: true,
      get() {
        const vals = readKey('em').map((v) => v || 0);
        if (!vals.length) return 0;
        return vals.every((v) => v === vals[0]) ? vals[0] : WD_UNDEFINED;
      },
      set(v) { if (v === WD_UNDEFINED) return; writeKey('em', v || false); },
    });
    Object.defineProperty(font, 'Borders', {
      enumerable: true, configurable: true,
      get() {
        return bordersOf(font, {
          edges: [{ key: 'border', index: -1 }],
          enable: ['border'],
          outside: ['border'],
          getBag() { const v = readKey('border'); return { border: v.find((x) => x) || null }; },
          setBag(bag) { writeKey('border', bag.border && bag.border.val !== 'nil' ? bag.border : false); },
        });
      },
    });
    Object.defineProperty(font, 'Shading', {
      enumerable: true, configurable: true,
      get() {
        return shadingOf(font,
          () => { const v = readKey('shd'); return v.find((x) => x) || {}; },
          (shd) => writeKey('shd', shd));
      },
    });
    denyGet(font, ['Animation', 'DiacriticColor', 'StylisticSet', 'NumberSpacing', 'NumberForm', 'Ligatures', 'ContextualAlternates'], 'G-ENGINE');
    return font;
  }

  function replaceSpan(a, b, raw) {
    const s = String(raw == null ? '' : raw);
    edit('Text', () => {
      let pos = D.eqPos(a, b) ? a : O.deleteRange(a, b);
      if (s) pos = O.insertPlain(pos, s);
      return pos;
    });
  }

  function insertBeside(a, b, where, text) {
    edit('Insert', () => {
      const at = where === 'before' ? D.pos(a.p, a.o) : D.pos(b.p, b.o);
      const end = O.insertPlain(at, String(text == null ? '' : text));
      return { a: at, f: end };
    });
  }

  function typeText(text) {
    const sel = ensureCaret();
    const typed = String(text == null ? '' : text);
    edit('Type', () => {
      let pos = D.eqPos(sel.a, sel.f) ? sel.f : O.deleteRange(sel.a, sel.f);
      const pend = pendHost().pending;
      const rPr = pend ? D.mergeRPr(O.inheritRPr(pos), pend) : null;
      if (typed) pos = O.insertPlain(pos, typed, rPr);
      return pos;
    });
  }

  function typeParagraph() {
    const sel = ensureCaret();
    edit('Paragraph', () => {
      const pos = D.eqPos(sel.a, sel.f) ? sel.f : O.deleteRange(sel.a, sel.f);
      return O.splitPara(pos);
    });
  }

  function typeBackspace() {
    const sel = ensureCaret();
    edit('Delete', () => {
      if (!D.eqPos(sel.a, sel.f)) return O.deleteRange(sel.a, sel.f);
      const pos = sel.f;
      if (pos.o > 0) return O.deleteRange(D.pos(pos.p, pos.o - 1), pos);
      const prev = D.prevPara(D.doc, pos.p);
      if (!prev || O.containerOf(prev) !== O.containerOf(pos.p)) return pos;
      return O.joinNext(prev) || pos;
    });
  }

  function deleteForward(unit, count) {
    if (unit !== undefined || count !== undefined) gap('G-STORY', 'Delete');
    const sel = ensureCaret();
    edit('Delete', () => {
      if (!D.eqPos(sel.a, sel.f)) return O.deleteRange(sel.a, sel.f);
      const pos = sel.f;
      if (pos.o < D.plen(pos.p)) return O.deleteRange(pos, D.pos(pos.p, pos.o + 1));
      const next = D.nextPara(D.doc, pos.p);
      if (!next || O.containerOf(next) !== O.containerOf(pos.p)) return pos;
      return O.joinNext(pos.p) || pos;
    });
  }

  function collapseSel(dir) {
    const [a, b] = endsOf({ kind: 'selection' });
    placeCaret(dir === 1 ? a : b);
  }

  const PARA_DENY = ['CharacterUnitLeftIndent', 'CharacterUnitRightIndent', 'CharacterUnitFirstLineIndent', 'LineUnitBefore', 'LineUnitAfter', 'AutoAdjustRightIndent', 'BaselineAlignment', 'FarEastLineBreakControl', 'HangingPunctuation', 'HalfWidthPunctuationOnTopOfLine', 'AddSpaceBetweenFarEastAndAlpha', 'AddSpaceBetweenFarEastAndDigit', 'TextboxTightWrap', 'Collapsed', 'WordWrap'];
  const TAB_AL_FROM = { 0: 'left', 1: 'center', 2: 'right', 3: 'decimal' };
  const TAB_AL_TO = { left: 0, center: 1, right: 2, decimal: 3, num: 0 };
  const TAB_LD_FROM = { 0: 'none', 1: 'dot', 2: 'hyphen', 3: 'underscore' };
  const TAB_LD_TO = { none: 0, dot: 1, hyphen: 2, underscore: 3, space: 0 };

  function parasOfAnchor(anchor) {
    if (!anchor) return [];
    if (anchor.kind === 'paragraph') return anchor.block ? [anchor.block] : [];
    if (anchor.kind === 'cell') return (anchor.cell.blocks || []).filter((b) => b.t === 'p');
    if (anchor.kind === 'story') {
      const out = [];
      if (anchor.story && anchor.story.blocks) D.walk(anchor.story, (b) => { if (b.t === 'p') out.push(b); });
      return out;
    }
    if (anchor.kind === 'section') {
      const bl = (D.doc.main && D.doc.main.blocks) || [];
      const out = [];
      const sec = anchor.sec;
      for (let i = sec.from; i < sec.to; i++) {
        const b = bl[i];
        if (!b) continue;
        if (b.t === 'p') out.push(b);
        else if (b.t === 'tbl') for (const row of b.rows) for (const c of row.cells) D.walk(c, (x) => { if (x.t === 'p') out.push(x); });
      }
      return out;
    }
    const e = endsOf(anchor);
    if (!e) return [];
    return D.parasBetween(D.doc, e[0].p, e[1].p);
  }

  function siblingPara(block, dir) {
    const info = D.info(D.doc, block);
    if (!info || !info.cont) return null;
    const bl = info.cont.blocks;
    let i = bl.indexOf(block) + dir;
    while (i >= 0 && i < bl.length) {
      if (bl[i].t === 'p') return bl[i];
      i += dir;
    }
    return null;
  }

  function writeBlocks(blocks, fn) {
    if (!blocks.length) return;
    edit('Paragraph', () => {
      for (const b of blocks) {
        if (b._style) {
          D.touchKey(b._model, 'styles');
          b._style.pPr = fn(b._style.pPr || {});
          b.pPr = b._style.pPr;
          D.stylesChanged();
        } else O.setParaProps([b], (cur) => fn(cur || {}));
      }
      return caret;
    });
  }

  function ruleOf(pp) {
    const sp = pp.sp || {};
    const rule = sp.rule || 'auto';
    const line = sp.line == null ? 1 : sp.line;
    if (rule === 'atLeast') return 3;
    if (rule === 'exact') return 4;
    if (line === 1.5) return 1;
    if (line === 2) return 2;
    if (line === 1) return 0;
    return 5;
  }

  function spacingPoints(pp) {
    const sp = pp.sp || {};
    const line = sp.line == null ? 1 : sp.line;
    if ((sp.rule || 'auto') === 'atLeast' || sp.rule === 'exact') return line;
    return line * 12;
  }

  function makeTabStops(parent, getBlocks) {
    const col = function (i) { return col.Item(i); };
    stamp(col, parent);
    function tabsOf(b) { return (b.pPr && b.pPr.tabs) || []; }
    Object.defineProperty(col, 'Count', { enumerable: true, get() { const b = getBlocks()[0]; return b ? tabsOf(b).length : 0; } });
    col.Item = function (index) {
      const b = getBlocks()[0];
      const t = b && tabsOf(b)[(index | 0) - 1];
      if (!t) missing();
      return makeTabStop(col, getBlocks, (index | 0) - 1);
    };
    col.Add = function (position, alignment, leader) {
      const al = TAB_AL_FROM[alignment | 0] || 'left';
      const ld = TAB_LD_FROM[leader | 0] || 'none';
      const pos = Number(position) || 0;
      writeBlocks(getBlocks(), (pp) => {
        const list = (pp.tabs || []).filter((t) => t.pos !== pos);
        list.push({ pos, al, leader: ld });
        list.sort((a, c) => a.pos - c.pos);
        return Object.assign({}, pp, { tabs: list });
      });
      const b = getBlocks()[0];
      const i = b ? tabsOf(b).findIndex((t) => t.pos === pos) : 0;
      return makeTabStop(col, getBlocks, Math.max(0, i));
    };
    col.ClearAll = function () {
      writeBlocks(getBlocks(), (pp) => { const n = Object.assign({}, pp); delete n.tabs; return n; });
    };
    return col;
  }

  function makeTabStop(parent, getBlocks, index) {
    const obj = stamp({}, parent);
    function tab() { const b = getBlocks()[0]; return b && b.pPr && b.pPr.tabs && b.pPr.tabs[index]; }
    function rewrite(fn) {
      writeBlocks(getBlocks(), (pp) => {
        const list = (pp.tabs || []).map((t) => Object.assign({}, t));
        if (!list[index]) return pp;
        fn(list[index]);
        list.sort((a, c) => a.pos - c.pos);
        return Object.assign({}, pp, { tabs: list });
      });
    }
    Object.defineProperty(obj, 'Position', {
      enumerable: true,
      get() { const t = tab(); return t ? t.pos : 0; },
      set(v) { rewrite((t) => { t.pos = Number(v) || 0; }); },
    });
    Object.defineProperty(obj, 'Alignment', {
      enumerable: true,
      get() { const t = tab(); return t && TAB_AL_TO[t.al] != null ? TAB_AL_TO[t.al] : 0; },
      set(v) { rewrite((t) => { t.al = TAB_AL_FROM[v | 0] || 'left'; }); },
    });
    Object.defineProperty(obj, 'Leader', {
      enumerable: true,
      get() { const t = tab(); return t && TAB_LD_TO[t.leader] != null ? TAB_LD_TO[t.leader] : 0; },
      set(v) { rewrite((t) => { t.leader = TAB_LD_FROM[v | 0] || 'none'; }); },
    });
    obj.Clear = function () {
      writeBlocks(getBlocks(), (pp) => {
        const list = (pp.tabs || []).slice();
        list.splice(index, 1);
        const n = Object.assign({}, pp);
        if (list.length) n.tabs = list; else delete n.tabs;
        return n;
      });
    };
    return obj;
  }

  function makeParaFormat(parent, getBlocks) {
    const fmt = stamp({}, parent);
    function blocks() { return getBlocks().filter((b) => b && (b.t === 'p' || b._style)); }
    function read(fn, empty) {
      const bs = blocks();
      if (!bs.length) return empty;
      const vals = bs.map((b) => fn(b.pPr || {}));
      return vals.every((v) => v === vals[0]) ? vals[0] : WD_UNDEFINED;
    }
    function write(fn) { writeBlocks(blocks(), fn); }
    Object.defineProperty(fmt, 'Alignment', {
      enumerable: true, configurable: true,
      get() { return read((pp) => alignOf({ pPr: pp }), 0); },
      set(v) {
        const jc = ALIGN_JC[v | 0];
        if (!jc) gap('G-ENGINE', 'Alignment');
        write((pp) => Object.assign({}, pp, { jc }));
      },
    });
    function indentProp(name, key) {
      Object.defineProperty(fmt, name, {
        enumerable: true, configurable: true,
        get() { return read((pp) => (pp.ind && pp.ind[key]) || 0, 0); },
        set(v) {
          if (v === WD_UNDEFINED) return;
          write((pp) => Object.assign({}, pp, { ind: Object.assign({}, pp.ind || {}, { [key]: Number(v) || 0 }) }));
        },
      });
    }
    indentProp('LeftIndent', 'l');
    indentProp('RightIndent', 'r');
    indentProp('FirstLineIndent', 'fl');
    function spaceProp(name, key) {
      Object.defineProperty(fmt, name, {
        enumerable: true, configurable: true,
        get() { return read((pp) => (pp.sp && pp.sp[key]) || 0, 0); },
        set(v) {
          if (v === WD_UNDEFINED) return;
          write((pp) => Object.assign({}, pp, { sp: Object.assign({}, pp.sp || {}, { [key]: Number(v) || 0 }) }));
        },
      });
    }
    spaceProp('SpaceBefore', 'b');
    spaceProp('SpaceAfter', 'a');
    function autoProp(name, key) {
      Object.defineProperty(fmt, name, {
        enumerable: true, configurable: true,
        get() { return read((pp) => (pp.sp && pp.sp[key] ? -1 : 0), 0); },
        set(v) {
          const t = triWrite(v);
          if (t === 'skip') return;
          write((pp) => Object.assign({}, pp, { sp: Object.assign({}, pp.sp || {}, { [key]: !!t }) }));
        },
      });
    }
    autoProp('SpaceBeforeAuto', 'bAuto');
    autoProp('SpaceAfterAuto', 'aAuto');
    Object.defineProperty(fmt, 'LineSpacingRule', {
      enumerable: true, configurable: true,
      get() { return read(ruleOf, 0); },
      set(v) {
        const n = v | 0;
        if (n < 0 || n > 5) gap('G-ENGINE', 'LineSpacingRule');
        write((pp) => {
          const sp = Object.assign({}, pp.sp || {});
          if (n === 0) { sp.rule = 'auto'; sp.line = 1; }
          else if (n === 1) { sp.rule = 'auto'; sp.line = 1.5; }
          else if (n === 2) { sp.rule = 'auto'; sp.line = 2; }
          else if (n === 3) { sp.rule = 'atLeast'; if (sp.line == null || sp.line <= 3) sp.line = 12; }
          else if (n === 4) { sp.rule = 'exact'; if (sp.line == null || sp.line <= 3) sp.line = 12; }
          else sp.rule = 'auto';
          return Object.assign({}, pp, { sp });
        });
      },
    });
    Object.defineProperty(fmt, 'LineSpacing', {
      enumerable: true, configurable: true,
      get() { return read(spacingPoints, 12); },
      set(v) {
        if (v === WD_UNDEFINED) return;
        write((pp) => {
          const sp = Object.assign({}, pp.sp || {});
          const rule = sp.rule || 'auto';
          sp.line = (rule === 'atLeast' || rule === 'exact') ? (Number(v) || 0) : (Number(v) || 12) / 12;
          return Object.assign({}, pp, { sp });
        });
      },
    });
    Object.defineProperty(fmt, 'OutlineLevel', {
      enumerable: true, configurable: true,
      get() { return read((pp) => (pp.outline == null ? 10 : pp.outline + 1), 10); },
      set(v) {
        const n = v | 0;
        write((pp) => {
          const next = Object.assign({}, pp);
          if (n >= 10) delete next.outline;
          else next.outline = Math.max(0, n - 1);
          return next;
        });
      },
    });
    function boolProp(name, key) {
      Object.defineProperty(fmt, name, {
        enumerable: true, configurable: true,
        get() { return read((pp) => (pp[key] ? -1 : 0), 0); },
        set(v) {
          const t = triWrite(v);
          if (t === 'skip') return;
          write((pp) => Object.assign({}, pp, { [key]: !!t }));
        },
      });
    }
    boolProp('KeepTogether', 'keepLines');
    boolProp('KeepWithNext', 'keepNext');
    boolProp('PageBreakBefore', 'pageBreakBefore');
    boolProp('DisableLineHeightGrid', 'noSnap');
    Object.defineProperty(fmt, 'WidowControl', {
      enumerable: true, configurable: true,
      get() { return read((pp) => (pp.widow === false ? 0 : -1), -1); },
      set(v) {
        const t = triWrite(v);
        if (t === 'skip') return;
        write((pp) => Object.assign({}, pp, { widow: !!t }));
      },
    });
    Object.defineProperty(fmt, 'Hyphenation', {
      enumerable: true, configurable: true,
      get() { return read((pp) => (pp.noHyphen ? 0 : -1), -1); },
      set(v) {
        const t = triWrite(v);
        if (t === 'skip') return;
        write((pp) => Object.assign({}, pp, { noHyphen: !t }));
      },
    });
    Object.defineProperty(fmt, 'ReadingOrder', {
      enumerable: true, configurable: true,
      get() { return read((pp) => (pp.bidi ? 1 : 0), 0); },
      set(v) { write((pp) => Object.assign({}, pp, { bidi: (v | 0) === 1 })); },
    });
    Object.defineProperty(fmt, 'TabStops', {
      enumerable: true, configurable: true,
      get() { return makeTabStops(fmt, blocks); },
    });
    const paraEdges = [
      { key: 'top', index: -1 }, { key: 'left', index: -2 }, { key: 'bottom', index: -3 },
      { key: 'right', index: -4 }, { key: 'between', index: -5 },
    ];
    Object.defineProperty(fmt, 'Borders', {
      enumerable: true, configurable: true,
      get() {
        return bordersOf(fmt, {
          edges: paraEdges,
          enable: ['top', 'left', 'bottom', 'right', 'between'],
          outside: ['top', 'left', 'bottom', 'right'],
          inside: ['between'],
          getBag() { const p = blocks()[0]; return (p && p.pPr && p.pPr.borders) || {}; },
          setBag(bag) { write((pp) => Object.assign({}, pp, { borders: bag })); },
        });
      },
    });
    Object.defineProperty(fmt, 'Shading', {
      enumerable: true, configurable: true,
      get() {
        return shadingOf(fmt,
          () => { const p = blocks()[0]; return (p && p.pPr && p.pPr.shd) || {}; },
          (shd) => write((pp) => Object.assign({}, pp, { shd: shd })));
      },
    });
    boolProp('MirrorIndents', 'mirrorInd');
    denyGet(fmt, PARA_DENY, 'G-ENGINE');
    return fmt;
  }

  function makeParagraphFormat(parent, block) {
    return makeParaFormat(parent, () => [block]);
  }

  function makeParagraph(doc, block) {
    const para = stamp({}, doc);
    para._block = block;
    para.Range = makeRange(para, { kind: 'paragraph', block, doc });
    Object.defineProperty(para, 'Alignment', {
      enumerable: true, configurable: true,
      get() { return alignOf(block); },
      set(v) { writeAlign(block, v); },
    });
    Object.defineProperty(para, 'Style', {
      enumerable: true, configurable: true,
      get() { return styleName(doc.model, block); },
      set(v) { writeStyle(block, v); },
    });
    Object.defineProperty(para, 'Format', {
      enumerable: true, configurable: true,
      get() { return makeParagraphFormat(para, block); },
    });
    for (const n of ['LeftIndent', 'RightIndent', 'FirstLineIndent', 'SpaceBefore', 'SpaceAfter', 'SpaceBeforeAuto', 'SpaceAfterAuto', 'LineSpacing', 'LineSpacingRule', 'OutlineLevel', 'KeepTogether', 'KeepWithNext', 'PageBreakBefore', 'WidowControl', 'Hyphenation', 'DisableLineHeightGrid', 'ReadingOrder']) {
      Object.defineProperty(para, n, {
        enumerable: true, configurable: true,
        get() { return para.Format[n]; },
        set(v) { para.Format[n] = v; },
      });
    }
    Object.defineProperty(para, 'TabStops', { enumerable: true, get() { return para.Format.TabStops; } });
    Object.defineProperty(para, 'ListFormat', { enumerable: true, get() { return listFormatOf(para, () => [block]); } });
    para.Space1 = function () { para.LineSpacingRule = 0; };
    para.Space15 = function () { para.LineSpacingRule = 1; };
    para.Space2 = function () { para.LineSpacingRule = 2; };
    para.Indent = function () { para.LeftIndent = (Number(para.LeftIndent) || 0) + 36; };
    para.IndentCharWidth = function () { para.Indent(); };
    para.Outdent = function () { para.LeftIndent = Math.max(0, (Number(para.LeftIndent) || 0) - 36); };
    para.Reset = function () {
      edit('Paragraph', () => {
        O.setParaProps([block], (pp) => {
          const keep = {};
          if (pp.style) keep.style = pp.style;
          if (pp.num) keep.num = pp.num;
          return keep;
        });
        return caret;
      });
    };
    para.Next = function (count) {
      let n = block;
      const c = count == null ? 1 : (count | 0);
      const dir = c < 0 ? -1 : 1;
      for (let i = 0; i < Math.abs(c || 1); i++) {
        const s = siblingPara(n, dir);
        if (!s) return null;
        n = s;
      }
      return makeParagraph(doc, n);
    };
    para.Previous = function (count) { return para.Next(count == null ? -1 : -(count | 0)); };
    Object.defineProperty(para, 'DropCap', {
      enumerable: true,
      get() {
        const dc = stamp({}, para);
        function bag() { return (block.pPr && block.pPr.dropCap) || block.dropCap || {}; }
        function writeDc(patch) {
          edit('Drop Cap', () => {
            D.touch(block);
            const next = Object.assign({ type: 'none', lines: 3, dist: 0 }, bag(), patch);
            if (!next.type || next.type === 'none') {
              if (block.pPr) delete block.pPr.dropCap;
              delete block.dropCap;
            } else {
              block.pPr = Object.assign({}, block.pPr, { dropCap: next });
              block.dropCap = next;
            }
            return caret;
          });
        }
        Object.defineProperty(dc, 'Position', {
          enumerable: true, configurable: true,
          get() { const t = bag().type; return t === 'margin' ? 2 : (t && t !== 'none' ? 1 : 0); },
          set(v) { writeDc({ type: (v | 0) === 2 ? 'margin' : (v | 0) === 1 ? 'drop' : 'none' }); },
        });
        Object.defineProperty(dc, 'Lines', {
          enumerable: true, configurable: true,
          get() { return bag().lines || 0; },
          set(v) { writeDc({ lines: v | 0, type: bag().type && bag().type !== 'none' ? bag().type : 'drop' }); },
        });
        Object.defineProperty(dc, 'DistanceFromText', {
          enumerable: true, configurable: true,
          get() { return bag().dist || 0; },
          set(v) { writeDc({ dist: Number(v) || 0, type: bag().type && bag().type !== 'none' ? bag().type : 'drop' }); },
        });
        return dc;
      },
    });
    denyFn(para, ['OutlineDemote', 'OutlinePromote', 'OutlineDemoteToBody'], 'G-ENGINE');
    return para;
  }

  function makeRange(parent, anchor) {
    const r = stamp({}, parent);
    r._anchor = anchor || null;
    attachStory(r, () => anchor);
    function rangeEnds() { return endsOf(anchor); }
    r.Characters = charactersOf(r, rangeEnds);
    Object.defineProperty(r, 'Text', {
      enumerable: true, configurable: true,
      get() {
        if (!anchor) return '';
        if (anchor.kind === 'paragraph') return D.ptext(anchor.block) + '\r';
        if (anchor.kind === 'cell') return cellText(anchor.cell);
        if (anchor.kind === 'story' || anchor.kind === 'section') {
          const paras = parasOfAnchor(anchor);
          return paras.length ? paras.map((p) => D.ptext(p)).join('\r') + '\r' : '';
        }
        const e = rangeEnds();
        return e ? O.textRange(e[0], e[1], '\r') : '';
      },
      set(v) {
        if (!anchor) return;
        if (anchor.kind === 'paragraph') { writeParagraphBody(anchor.block, v); return; }
        if (anchor.kind === 'cell') { writeCellText(anchor.cell, v); return; }
        if (anchor.kind === 'story') { writeStory(anchor.story, v); return; }
        const e = rangeEnds();
        if (e) replaceSpan(e[0], e[1], v);
      },
    });
    Object.defineProperty(r, 'Style', {
      enumerable: true, configurable: true,
      get() {
        if (anchor && anchor.kind === 'paragraph') return styleName(anchor.doc.model, anchor.block);
        const e = rangeEnds();
        if (!e) return '';
        const names = D.parasBetween(D.doc, e[0].p, e[1].p).map((p) => styleName(D.doc, p));
        return names.every((n) => n === names[0]) ? names[0] : WD_UNDEFINED;
      },
      set(v) {
        if (anchor && anchor.kind === 'paragraph') { writeStyle(anchor.block, v); return; }
        const e = rangeEnds();
        if (!e) return;
        for (const p of D.parasBetween(D.doc, e[0].p, e[1].p)) writeStyle(p, v);
      },
    });
    Object.defineProperty(r, 'ParagraphFormat', {
      enumerable: true, configurable: true,
      get() {
        if (anchor && anchor.kind === 'paragraph') return makeParagraphFormat(r, anchor.block);
        return makeSpanFormat(r, rangeEnds);
      },
    });
    Object.defineProperty(r, 'Borders', { enumerable: true, get() { return r.ParagraphFormat.Borders; } });
    Object.defineProperty(r, 'Shading', { enumerable: true, get() { return r.ParagraphFormat.Shading; } });
    Object.defineProperty(r, 'Font', {
      enumerable: true, configurable: true,
      get() { return makeFont(r, rangeEnds); },
    });
    r.Bold = undefined;
    Object.defineProperty(r, 'Bold', { enumerable: true, configurable: true, get() { return r.Font.Bold; }, set(v) { r.Font.Bold = v; } });
    Object.defineProperty(r, 'Italic', { enumerable: true, configurable: true, get() { return r.Font.Italic; }, set(v) { r.Font.Italic = v; } });
    Object.defineProperty(r, 'Underline', { enumerable: true, configurable: true, get() { return r.Font.Underline; }, set(v) { r.Font.Underline = v; } });
    r.InsertBefore = function (text) { const e = rangeEnds(); if (e) insertBeside(e[0], e[1], 'before', text); };
    r.InsertAfter = function (text) { const e = rangeEnds(); if (e) insertBeside(e[0], e[1], 'after', text); };
    r.InsertParagraph = function () {
      const e = rangeEnds();
      if (!e) return;
      edit('Paragraph', () => O.splitPara(D.eqPos(e[0], e[1]) ? e[0] : O.deleteRange(e[0], e[1])));
    };
    r.InsertParagraphBefore = function () { const e = rangeEnds(); if (e) edit('Paragraph', () => O.splitPara(e[0])); };
    r.InsertParagraphAfter = function () { const e = rangeEnds(); if (e) edit('Paragraph', () => O.splitPara(e[1])); };
    r.Delete = function (unit, count) {
      if (unit !== undefined || count !== undefined) gap('G-STORY', 'Delete');
      const e = rangeEnds();
      if (!e || D.eqPos(e[0], e[1])) return;
      edit('Delete', () => O.deleteRange(e[0], e[1]));
    };
    Object.defineProperty(r, 'Find', {
      enumerable: true, configurable: true,
      get() { if (!r._find) r._find = makeFind(r, anchor); return r._find; },
    });
    r.Collapse = function (dir) {
      if (!anchor) return;
      if (anchor.kind === 'selection') { collapseSel(dir); return; }
      const e = rangeEnds();
      if (!e) return;
      const pos = dir === 1 ? e[0] : e[1];
      anchor.kind = 'span';
      anchor.a = pos;
      anchor.f = pos;
    };
    attachRangeExtras(r, anchor, rangeEnds);
    return r;
  }

  function makeSpanFormat(parent, getEnds) {
    return makeParaFormat(parent, () => {
      const e = getEnds();
      return e ? D.parasBetween(D.doc, e[0].p, e[1].p) : [];
    });
  }

  function placeSpan(a, b) {
    caret = { a: a, f: b };
    if (L.ed) {
      L.ed.sel = caret;
      if (typeof L.ed.setSel === 'function') L.ed.setSel(a, b, { noScroll: true });
    }
  }

  function cloneAnchor(a) {
    if (!a) return null;
    const c = Object.assign({}, a);
    if (a.a) c.a = D.pos(a.a.p, a.a.o);
    if (a.f) c.f = D.pos(a.f.p, a.f.o);
    return c;
  }

  function clipBlocks() {
    if (L.ed && L.ed.clip && L.ed.clip.blocks) return L.ed.clip.blocks;
    return dom.clip || null;
  }

  function setClip(blocks) {
    dom.clip = blocks;
    if (L.ed) L.ed.clip = { blocks: blocks };
  }

  function copySpan(a, b, cut) {
    if (D.eqPos(a, b)) return;
    edit(cut ? 'Cut' : 'Copy', () => {
      setClip(O.copyRange(a, b));
      if (cut) return O.deleteRange(a, b);
      return caret;
    });
  }

  function collectExtents(paras, kind) {
    const out = [];
    const fn = kind === 'word' ? O.wordExtent : O.sentenceExtent;
    for (const p of paras) {
      const n = D.plen(p);
      if (!n) continue;
      let o = 0;
      let guard = 0;
      while (o < n && guard++ < n + 2) {
        const ext = fn(D.pos(p, o));
        if (!ext) break;
        const b = ext[1];
        if (!b || b.p !== p || b.o <= o) { o++; continue; }
        const text = D.ptext(p, ext[0].o, b.o);
        if (kind === 'word' && !/[\p{L}\p{N}]/u.test(text)) { o = b.o; continue; }
        if (kind === 'sentence' && !String(text).trim()) { o = b.o; continue; }
        out.push({ a: D.pos(p, ext[0].o), b: D.pos(p, b.o) });
        o = b.o;
      }
    }
    return out;
  }

  function extentCollection(parent, getParas, kind) {
    const col = function (i) { return col.Item(i); };
    stamp(col, parent);
    function items() { return collectExtents(getParas(), kind); }
    Object.defineProperty(col, 'Count', { enumerable: true, get() { return items().length; } });
    col.Item = function (index) {
      const it = items()[(index | 0) - 1];
      if (!it) missing();
      return makeRange(col, { kind: 'span', a: it.a, f: it.b });
    };
    Object.defineProperty(col, 'First', { enumerable: true, get() { return col.Count ? col.Item(1) : null; } });
    Object.defineProperty(col, 'Last', { enumerable: true, get() { return col.Count ? col.Item(col.Count) : null; } });
    return col;
  }

  function listFormatOf(parent, getParas) {
    const obj = stamp({}, parent);
    function paras() { return getParas().filter((p) => p && p.t === 'p'); }
    Object.defineProperty(obj, 'ListLevelNumber', {
      enumerable: true, configurable: true,
      get() {
        const lv = paras().map((p) => (p.pPr && p.pPr.num && p.pPr.num.id && p.pPr.num.id !== '0' ? (p.pPr.num.lvl | 0) + 1 : 1));
        if (!lv.length) return 1;
        return lv.every((v) => v === lv[0]) ? lv[0] : WD_UNDEFINED;
      },
      set(v) {
        const lvl = Math.max(0, (v | 0) - 1);
        writeBlocks(paras(), (pp) => Object.assign({}, pp, { num: Object.assign({ id: '1', lvl: 0 }, pp.num || {}, { lvl }) }));
      },
    });
    function apply(id) {
      writeBlocks(paras(), (pp) => Object.assign({}, pp, { num: { id: id, lvl: (pp.num && pp.num.lvl) || 0 } }));
    }
    obj.ApplyBulletDefault = function () { apply('1'); };
    obj.ApplyNumberDefault = function () { apply('2'); };
    obj.RemoveNumbers = function () { apply('0'); };
    obj.ListIndent = function () {
      const n = obj.ListLevelNumber === WD_UNDEFINED ? 1 : obj.ListLevelNumber;
      obj.ListLevelNumber = n + 1;
    };
    obj.ListOutdent = function () {
      const n = obj.ListLevelNumber === WD_UNDEFINED ? 1 : obj.ListLevelNumber;
      obj.ListLevelNumber = Math.max(1, n - 1);
    };
    obj.ApplyListTemplate = function () { gap('G-ENGINE', 'ApplyListTemplate'); };
    Object.defineProperty(obj, 'ListType', {
      enumerable: true,
      get() {
        const labels = typeof D.computeLists === 'function' ? D.computeLists(D.doc) : null;
        const types = paras().map((p) => {
          const lab = labels && labels.get ? labels.get(p.id) : null;
          if (!lab) return 0;
          return lab.lv && lab.lv.fmt === 'bullet' ? 2 : 3;
        });
        if (!types.length || types.every((t) => t === 0)) return 0;
        return types.every((t) => t === types[0]) ? types[0] : 5;
      },
    });
    denyGet(obj, ['CanContinuePreviousList'], 'G-ENGINE');
    return obj;
  }

  const LCID = { 1033: 'en-US', 2052: 'zh-CN', 1041: 'ja-JP', 1036: 'fr-FR', 1031: 'de-DE', 1042: 'ko-KR', 2057: 'en-GB' };

  function langTag(v) {
    if (typeof v === 'number') {
      if (!LCID[v]) gap('G-ENGINE', 'LanguageID');
      return LCID[v];
    }
    return String(v);
  }

  function langRead(vals) {
    const s = vals.map((v) => v || '');
    if (!s.length || s.every((x) => !x)) return WD_UNDEFINED;
    if (!s.every((x) => x === s[0])) return WD_UNDEFINED;
    for (const n in LCID) if (LCID[n] === s[0]) return Number(n);
    return s[0];
  }

  function insertBreak(type) {
    const t = type == null ? 7 : (type | 0);
    const c = ensureCaret();
    if (t === 6 || t === 11) {
      edit('Break', () => O.insertItem(c.f, D.item('br', { type: t === 11 ? 'wrap' : 'line' })));
      return;
    }
    if (t === 7) { edit('Break', () => O.insertItem(c.f, D.item('br', { type: 'page' }))); return; }
    if (t === 8) { edit('Break', () => O.insertItem(c.f, D.item('br', { type: 'column' }))); return; }
    if (t >= 2 && t <= 5) {
      const types = { 2: 'nextPage', 3: 'continuous', 4: 'evenPage', 5: 'oddPage' };
      edit('Break', () => {
        const pos = O.splitPara(c.f);
        const info = D.info(D.doc, pos.p);
        const bl = info && info.cont ? info.cont.blocks : [];
        const prev = bl[bl.indexOf(pos.p) - 1];
        if (prev && prev.t === 'p') {
          D.touch(prev);
          prev.sect = Object.assign(L.clone(D.sectFor(D.doc, prev) || {}), { type: types[t] });
        }
        return pos;
      });
      return;
    }
    gap('G-ENGINE', 'InsertBreak');
  }

  function insertSymbol(n, font) {
    const ch = String.fromCodePoint(n | 0);
    const rPr = font ? { font: String(font) } : null;
    edit('Symbol', () => {
      const c = ensureCaret();
      let pos = D.eqPos(c.a, c.f) ? c.f : O.deleteRange(c.a, c.f);
      return O.insertText(pos, ch, rPr);
    });
  }

  function formatDate(fmt) {
    const d = new Date();
    const pad = (n) => String(n).padStart(2, '0');
    const map = { yyyy: String(d.getFullYear()), yy: pad(d.getFullYear() % 100), MM: pad(d.getMonth() + 1), dd: pad(d.getDate()), HH: pad(d.getHours()), mm: pad(d.getMinutes()), ss: pad(d.getSeconds()) };
    if (!fmt) return d.toDateString();
    return String(fmt).replace(/yyyy|yy|MM|dd|HH|mm|ss/g, (k) => map[k]);
  }

  function insertDateTime(fmt, asField) {
    const text = formatDate(fmt);
    if (asField === true || asField === -1) {
      edit('Date', () => {
        const c = ensureCaret();
        const pos = D.eqPos(c.a, c.f) ? c.f : O.deleteRange(c.a, c.f);
        return O.insertField(pos, 'DATE', text);
      });
      return;
    }
    typeText(text);
  }

  function writeStory(story, raw) {
    if (!story || story._synthetic) gap('G-ENGINE', 'StoryRanges');
    const parts = stripOneBreak(raw).split(/\r\n|\r|\n/);
    edit('Text', () => {
      D.touchList(story);
      story.blocks = parts.map((part) => {
        const p = D.para();
        if (part) p.runs = [D.text(part)];
        return p;
      });
      if (!story.blocks.length) story.blocks.push(D.para());
      D.doc._idxDirty = true;
      return D.pos(story.blocks[0], D.plen(story.blocks[0]));
    });
  }

  function pictureAtOf(anchor) {
    return function () {
      const e = endsOf(anchor);
      if (!e || !D.doc) return null;
      const story = D.storyOf(D.doc, e[0].p);
      if (!story) return null;
      return { story, off: offInStory(storyTextOf(story), e[0]) };
    };
  }

  function attachRangeExtras(r, anchor, rangeEnds) {
    Object.defineProperty(r, '_pictureAt', { value: pictureAtOf(anchor) });
    r.Select = function () {
      const e = rangeEnds();
      if (e) placeSpan(e[0], e[1]);
    };
    r.Duplicate = function () { return makeRange(r.Parent, cloneAnchor(anchor)); };
    r.Copy = function () { const e = rangeEnds(); if (e) copySpan(e[0], e[1], false); };
    r.Cut = function () { const e = rangeEnds(); if (e) copySpan(e[0], e[1], true); };
    r.Paste = function () { const e = rangeEnds(); if (e) pasteBlocks(e[0], e[1], true); };
    Object.defineProperty(r, 'Case', {
      enumerable: true, configurable: true,
      get() { return WD_UNDEFINED; },
      set(v) {
        const mode = { 0: 'lower', 1: 'upper', 2: 'title', 4: 'sentence', 5: 'toggle' }[v | 0];
        if (!mode) gap('G-ENGINE', 'Case');
        const e = rangeEnds();
        if (!e || D.eqPos(e[0], e[1])) return;
        edit('Case', () => { O.changeCase(e[0], e[1], mode); return caret; });
      },
    });
    function langProp(name, key) {
      Object.defineProperty(r, name, {
        enumerable: true, configurable: true,
        get() {
          const e = rangeEnds();
          return e ? langRead(runValues(e[0], e[1], key)) : WD_UNDEFINED;
        },
        set(v) {
          const e = rangeEnds();
          if (e) applyRun(e[0], e[1], { [key]: langTag(v) });
        },
      });
    }
    langProp('LanguageID', 'lang');
    langProp('LanguageIDFarEast', 'langEA');
    r.InsertBreak = function (type) { const e = rangeEnds(); if (e) placeCaret(e[1]); insertBreak(type); };
    r.InsertSymbol = function (n, font) { const e = rangeEnds(); if (e) placeCaret(e[1]); insertSymbol(n, font); };
    r.InsertDateTime = function (fmt, asField, fullWidth, dateLanguage, calendarType) {
      if (fullWidth || calendarType) gap('G-ENGINE', 'InsertDateTime');
      const e = rangeEnds();
      if (e) placeCaret(e[1]);
      insertDateTime(fmt, asField);
    };
    Object.defineProperty(r, 'Sentences', { enumerable: true, get() { return extentCollection(r, () => parasOfAnchor(anchor), 'sentence'); } });
    Object.defineProperty(r, 'Words', { enumerable: true, get() { return extentCollection(r, () => parasOfAnchor(anchor), 'word'); } });
    Object.defineProperty(r, 'ListFormat', { enumerable: true, get() { return listFormatOf(r, () => parasOfAnchor(anchor)); } });
    Object.defineProperty(r, 'Hyperlinks', { enumerable: true, get() { return hyperlinksOf(r, anchor); } });
    Object.defineProperty(r, 'Paragraphs', { enumerable: true, get() { return rangeParagraphs(r, anchor); } });
    Object.defineProperty(r, 'Tables', { enumerable: true, get() { return rangeTables(r, anchor); } });
    Object.defineProperty(r, 'Bookmarks', { enumerable: true, get() { return rangeBookmarks(r, anchor); } });
    Object.defineProperty(r, 'Fields', { enumerable: true, get() { return rangeFields(r, anchor); } });
    Object.defineProperty(r, 'Comments', { enumerable: true, get() { return rangeComments(r, anchor); } });
    Object.defineProperty(r, 'InlineShapes', { enumerable: true, get() { return rangeShapes(r, anchor, false); } });
    Object.defineProperty(r, 'Shapes', { enumerable: true, get() { return rangeShapes(r, anchor, true); } });
    denyGet(r, ['FormattedText'], 'G-ENGINE');
    r.InsertFile = function (fileName) {
      if (!L.hostAudit || !L.hostAudit.has('InsertFile')) gap('G-HOST', 'InsertFile');
      L.hostAudit.rejectExtra(arguments, 1, ['FileName', 'Range', 'ConfirmConversions', 'Link', 'Attachment']);
      return L.hostAudit.readBytes(fileName, 'InsertFile', 'AlwaysAllow').then((bytes) => {
        return L.app.insertFileBytes(String(fileName).split(/[/\\]/).pop(), bytes);
      });
    };
    denyFn(r, ['InsertCaption', 'PasteAndFormat', 'PasteSpecial', 'PasteExcelTable'], 'G-ENGINE');
  }

  function rangeParagraphs(parent, anchor) {
    const col = function (i) { return col.Item(i); };
    stamp(col, parent);
    function paras() { return parasOfAnchor(anchor); }
    Object.defineProperty(col, 'Count', { enumerable: true, get() { return paras().length; } });
    col.Item = function (index) {
      const p = paras()[(index | 0) - 1];
      if (!p) missing();
      return makeParagraph(activeDocument(), p);
    };
    Object.defineProperty(col, 'First', { enumerable: true, get() { return col.Count ? col.Item(1) : null; } });
    Object.defineProperty(col, 'Last', { enumerable: true, get() { return col.Count ? col.Item(col.Count) : null; } });
    return col;
  }

  function rangeTables(parent, anchor) {
    const col = function (i) { return col.Item(i); };
    stamp(col, parent);
    function tables() {
      const out = [];
      for (const p of parasOfAnchor(anchor)) {
        const info = D.info(D.doc, p);
        if (!info || !info.parents) continue;
        for (const up of info.parents) if (up.tbl && out.indexOf(up.tbl) < 0) out.push(up.tbl);
      }
      return out;
    }
    Object.defineProperty(col, 'Count', { enumerable: true, get() { return tables().length; } });
    col.Item = function (index) {
      const t = tables()[(index | 0) - 1];
      if (!t) missing();
      return makeTable(activeDocument(), t);
    };
    col.Add = function (range, numRows, numColumns, defaultTableBehavior, autoFitBehavior) {
      const at = range && range._anchor ? range : makeRange(parent, anchor);
      return addTables(at, numRows, numColumns, defaultTableBehavior, autoFitBehavior);
    };
    Object.defineProperty(col, 'First', { enumerable: true, get() { return col.Count ? col.Item(1) : null; } });
    Object.defineProperty(col, 'Last', { enumerable: true, get() { return col.Count ? col.Item(col.Count) : null; } });
    return col;
  }

  function rangeBookmarks(parent, anchor) {
    const col = function (i) { return col.Item(i); };
    stamp(col, parent);
    function list() {
      const paras = new Set(parasOfAnchor(anchor));
      return listBookmarks(D.doc).filter((b) => paras.has(b.p) || paras.has(b.endP));
    }
    Object.defineProperty(col, 'Count', { enumerable: true, get() { return list().length; } });
    col.Item = function (index) {
      const all = list();
      if (typeof index === 'string') {
        const b = all.find((x) => x.name.toLowerCase() === index.toLowerCase());
        if (!b) missing();
        return makeBookmark(activeDocument(), b);
      }
      const b = all[(index | 0) - 1];
      if (!b) missing();
      return makeBookmark(activeDocument(), b);
    };
    col.Exists = function (name) {
      return list().some((b) => b.name.toLowerCase() === String(name).toLowerCase());
    };
    col.Add = function (name, range) {
      return bookmarksOf(activeDocument()).Add(name, range && range._anchor ? range : makeRange(parent, anchor));
    };
    Object.defineProperty(col, 'First', { enumerable: true, get() { return col.Count ? col.Item(1) : null; } });
    Object.defineProperty(col, 'Last', { enumerable: true, get() { return col.Count ? col.Item(col.Count) : null; } });
    return col;
  }

  function rangeFields(parent, anchor) {
    const holder = { blocks: parasOfAnchor(anchor).filter((p) => p.t === 'p') };
    const wrapped = { blocks: holder.blocks.map((p) => p) };
    return fieldsOf(activeDocument(), { blocks: wrapped.blocks });
  }

  function rangeComments(parent, anchor) {
    const col = function (i) { return col.Item(i); };
    stamp(col, parent);
    function ids() {
      const paras = new Set(parasOfAnchor(anchor));
      const found = [];
      for (const p of paras) {
        for (const it of p.runs) if (it.t === 'cs' && found.indexOf(String(it.id)) < 0) found.push(String(it.id));
      }
      return found;
    }
    Object.defineProperty(col, 'Count', { enumerable: true, get() { return ids().length; } });
    col.Item = function (index) {
      const id = ids()[(index | 0) - 1];
      if (id == null) missing();
      return makeComment(activeDocument(), id);
    };
    col.Add = function (range, text) {
      return commentsOf(activeDocument()).Add(range && range._anchor ? range : makeRange(parent, anchor), text);
    };
    Object.defineProperty(col, 'First', { enumerable: true, get() { return col.Count ? col.Item(1) : null; } });
    Object.defineProperty(col, 'Last', { enumerable: true, get() { return col.Count ? col.Item(col.Count) : null; } });
    return col;
  }

  function rangeShapes(parent, anchor, floating) {
    const col = function (i) { return col.Item(i); };
    stamp(col, parent);
    function list() {
      const paras = new Set(parasOfAnchor(anchor));
      return shapeItems(!!floating).filter((hit) => paras.has(hit.p));
    }
    Object.defineProperty(col, 'Count', { enumerable: true, get() { return list().length; } });
    col.Item = function (index) {
      const hit = list()[(index | 0) - 1];
      if (!hit) missing();
      return makeShape(activeDocument(), hit.p, hit.it);
    };
    col.AddPicture = function () {
      const args = Array.prototype.slice.call(arguments);
      const slot = floating ? 7 : 3;
      while (args.length <= slot) args.push(undefined);
      if (args[slot] == null) args[slot] = { _pictureAt: pictureAtOf(anchor) };
      return shapesOf(activeDocument(), !!floating).AddPicture.apply(null, args);
    };
    if (floating) col.AddPicture2 = function () { return col.AddPicture.apply(col, arguments); };
    col.AddChart = function () { gap('G-ENGINE', 'AddChart'); };
    Object.defineProperty(col, 'First', { enumerable: true, get() { return col.Count ? col.Item(1) : null; } });
    Object.defineProperty(col, 'Last', { enumerable: true, get() { return col.Count ? col.Item(col.Count) : null; } });
    return col;
  }

  function pasteBlocks(a, b, replace) {
    const blocks = clipBlocks();
    if (!blocks || !blocks.length) return;
    edit('Paste', () => {
      const pos = replace && !D.eqPos(a, b) ? O.deleteRange(a, b) : a;
      return O.insertBlocks(pos, blocks);
    });
  }

  function mainParas(model) {
    const out = [];
    if (model && model.main) D.walk(model.main, (b) => { if (b.t === 'p') out.push(b); });
    return out;
  }

  function requireActive(wrap) {
    if (!wrap || wrap !== activeWrap || wrap.model !== D.doc) {
      const e = new Error('Activate the document first');
      e.code = 'inactive';
      throw e;
    }
  }

  const wraps = [];
  let activeWrap = null;
  let caret = null;
  let untitled = 0;

  function adopt(model, name) {
    const found = wraps.find((w) => w.model === model);
    if (found) return found;
    untitled++;
    const wrap = makeDocument(model, name || ('Document' + untitled));
    wraps.push(wrap);
    return wrap;
  }

  function activeDocument() {
    if (!D.doc) return null;
    const found = wraps.find((w) => w.model === D.doc);
    if (found) { activeWrap = found; return found; }
    wraps.length = 0;
    caret = null;
    activeWrap = adopt(D.doc);
    clearPending();
    return activeWrap;
  }

  function fileTitle(name) {
    const base = String(name || '').split(/[/\\]/).pop();
    return base || 'Document';
  }

  function makeDocument(model, name) {
    const doc = stamp({}, null);
    doc.model = model;
    doc._name = fileTitle(name);

    Object.defineProperty(doc, 'Name', {
      enumerable: true, configurable: true,
      get() { return doc._name; },
    });

    doc.Activate = function () {
      D.doc = model;
      activeWrap = doc;
      return doc;
    };

    doc.Save = function () {
      requireActive(doc);
      const run = () => (L.app && typeof L.app.save === 'function' ? L.app.save() : undefined);
      if (!L.hostAudit) return run();
      return L.hostAudit.must('RequestApproval', 'Save', 'save the open file').then(run);
    };

    function quireSaveBytes(fileFormat) {
      if (fileFormat === 17 || fileFormat === 'pdf') return L.app.buildPDF();
      if (fileFormat != null && fileFormat !== '' && fileFormat !== 12 && fileFormat !== 16 && fileFormat !== 'docx') {
        const e = new Error('The parameter FileFormat is not available.');
        e.code = 'error';
        throw e;
      }
      return L.docx.write(model, { format: 'docx' });
    }
    doc.SaveAs = function (fileName, fileFormat) {
      if (!L.hostAudit || !L.hostAudit.has('SaveAs')) gap('G-HOST', 'SaveAs');
      L.hostAudit.rejectExtra(arguments, 2, ['FileName', 'FileFormat', 'LockComments', 'Password']);
      return L.hostAudit.saveCopy(fileName, () => quireSaveBytes(fileFormat));
    };
    doc.SaveAs2 = doc.SaveAs;

    function times(n) { return n == null ? 1 : (n | 0); }
    doc.Undo = function (n) {
      requireActive(doc);
      let any = false;
      for (let i = 0; i < times(n); i++) { if (!D.undo()) break; any = true; }
      return any;
    };
    doc.Redo = function (n) {
      requireActive(doc);
      let any = false;
      for (let i = 0; i < times(n); i++) { if (!D.redo()) break; any = true; }
      return any;
    };
    doc.Range = function (start, end) {
      requireActive(doc);
      const st = storyTextOf(doc.model.main);
      const a = posInStory(st, start == null ? 0 : (start | 0));
      const b = posInStory(st, end == null ? st.text.length : (end | 0));
      return makeRange(doc, { kind: 'span', a: a, f: b });
    };
    doc.GoTo = function () { requireActive(doc); gap('G-STORY', 'GoTo'); };

    Object.defineProperty(doc, 'Characters', {
      enumerable: true, configurable: true,
      get() {
        requireActive(doc);
        return charactersOf(doc, () => endsOf({ kind: 'story', story: doc.model.main }));
      },
    });

    Object.defineProperty(doc, 'Paragraphs', {
      enumerable: true, configurable: true,
      get() {
        requireActive(doc);
        const col = function (index) { return col.Item(index); };
        stamp(col, doc);
        Object.defineProperty(col, 'Count', {
          enumerable: true,
          get() { requireActive(doc); return mainParas(model).length; },
        });
        col.Item = function (index) {
          requireActive(doc);
          const paras = mainParas(model);
          const i = index | 0;
          if (i < 1 || i > paras.length) {
            const e = new Error('The requested member of the collection does not exist.');
            e.code = 'missing';
            throw e;
          }
          return makeParagraph(doc, paras[i - 1]);
        };
        Object.defineProperty(col, 'First', { enumerable: true, get() { return col.Count ? col.Item(1) : null; } });
        Object.defineProperty(col, 'Last', { enumerable: true, get() { return col.Count ? col.Item(col.Count) : null; } });
        return col;
      },
    });

    Object.defineProperty(doc, 'Tables', {
      enumerable: true, configurable: true,
      get() {
        requireActive(doc);
        const col = function (index) { return col.Item(index); };
        stamp(col, doc);
        Object.defineProperty(col, 'Count', {
          enumerable: true,
          get() { requireActive(doc); return mainTables(model).length; },
        });
        col.Item = function (index) {
          requireActive(doc);
          const tables = mainTables(model);
          const i = index | 0;
          if (i < 1 || i > tables.length) {
            const e = new Error('The requested member of the collection does not exist.');
            e.code = 'missing';
            throw e;
          }
          return makeTable(doc, tables[i - 1]);
        };
        col.Add = function (range, numRows, numColumns, defaultTableBehavior, autoFitBehavior) {
          requireActive(doc);
          return addTables(range, numRows, numColumns, defaultTableBehavior, autoFitBehavior);
        };
        Object.defineProperty(col, 'First', { enumerable: true, get() { return col.Count ? col.Item(1) : null; } });
        Object.defineProperty(col, 'Last', { enumerable: true, get() { return col.Count ? col.Item(col.Count) : null; } });
        return col;
      },
    });

    Object.defineProperty(doc, 'Bookmarks', {
      enumerable: true, configurable: true,
      get() { requireActive(doc); return bookmarksOf(doc); },
    });

    attachDocument(doc, model);
    doc.Parent = Documents;
    return doc;
  }

  const BUILTIN_PROPS = { Title: 'title', Subject: 'subject', Author: 'creator', Keywords: 'keywords', Comments: 'description', Category: 'category', Manager: 'manager', Company: 'company' };
  const STYLE_TYPE = { paragraph: 1, character: 2, table: 3, list: 4 };
  const STYLE_KIND = { 1: 'paragraph', 2: 'character', 3: 'table', 4: 'list' };

  function makeStyle(doc, id) {
    const model = doc.model;
    const obj = stamp({}, doc);
    obj._id = id;
    function st() { return model.styles[id]; }
    Object.defineProperty(obj, 'NameLocal', {
      enumerable: true,
      get() { const s = st(); return s ? D.styleDisplayName(s) : ''; },
      set(v) {
        edit('Style', () => { D.touchKey(model, 'styles'); if (st()) st().name = String(v); D.stylesChanged(); return caret; });
      },
    });
    Object.defineProperty(obj, 'Name', { enumerable: true, get() { return obj.NameLocal; }, set(v) { obj.NameLocal = v; } });
    Object.defineProperty(obj, 'Type', { enumerable: true, get() { const s = st(); return s ? (STYLE_TYPE[s.type] || 1) : 1; } });
    Object.defineProperty(obj, 'InUse', { enumerable: true, get() { return D.usedStyles(model).has(id) ? -1 : 0; } });
    Object.defineProperty(obj, 'Font', { enumerable: true, get() { return makeFont(obj, () => null, { model, style: st() }); } });
    Object.defineProperty(obj, 'ParagraphFormat', {
      enumerable: true,
      get() {
        const s = st();
        const fake = { t: 'p', pPr: (s && s.pPr) || {}, _style: s, _model: model };
        return makeParagraphFormat(obj, fake);
      },
    });
    function styleRef(key) {
      return {
        enumerable: true, configurable: true,
        get() {
          const s = st();
          const ref = s && s[key];
          if (!ref || !model.styles[ref]) return '';
          return D.styleDisplayName(model.styles[ref]);
        },
        set(v) {
          const found = D.findStyleByName(model, String(v));
          if (!found) missing();
          edit('Style', () => { D.touchKey(model, 'styles'); if (st()) st()[key] = found.id; D.stylesChanged(); return caret; });
        },
      };
    }
    Object.defineProperty(obj, 'BaseStyle', styleRef('basedOn'));
    Object.defineProperty(obj, 'NextParagraphStyle', styleRef('next'));
    obj.Delete = function () {
      const s = st();
      if (!s || s.builtin) gap('G-ENGINE', 'Style.Delete');
      edit('Style', () => {
        D.touchKey(model, 'styles');
        D.walk(model.main, (b) => {
          if (b.t === 'p' && b.pPr && b.pPr.style === id) { D.touch(b); b.pPr = Object.assign({}, b.pPr, { style: 'Normal' }); }
        });
        delete model.styles[id];
        D.stylesChanged();
        return caret;
      });
    };
    denyGet(obj, ['UnhideWhenUsed', 'Locked'], 'G-ENGINE');
    Object.defineProperty(obj, 'LinkStyle', styleRef('link'));
    Object.defineProperty(obj, 'QuickStyle', {
      enumerable: true, configurable: true,
      get() { return st() && st().q ? -1 : 0; },
      set(v) { const t = triWrite(v); if (t === 'skip') return; edit('Style', () => { D.touchKey(model, 'styles'); if (st()) st().q = !!t; D.stylesChanged(); return caret; }); },
    });
    Object.defineProperty(obj, 'Priority', {
      enumerable: true, configurable: true,
      get() { return (st() && st().prio) || 0; },
      set(v) { edit('Style', () => { D.touchKey(model, 'styles'); if (st()) st().prio = v | 0; D.stylesChanged(); return caret; }); },
    });
    Object.defineProperty(obj, 'Visibility', {
      enumerable: true, configurable: true,
      get() { const s = st(); return s && (s.hidden || s.semiHidden) ? -1 : 0; },
      set(v) { const t = triWrite(v); if (t === 'skip') return; edit('Style', () => { D.touchKey(model, 'styles'); if (st()) { st().semiHidden = !!t; st().hidden = !!t; } D.stylesChanged(); return caret; }); },
    });
    return obj;
  }

  function stylesOf(doc) {
    const model = doc.model;
    const col = function (i) { return col.Item(i); };
    stamp(col, doc);
    function ids() { return Object.keys(model.styles || {}); }
    Object.defineProperty(col, 'Count', { enumerable: true, get() { return ids().length; } });
    col.Item = function (index) {
      if (typeof index === 'number') {
        const id = ids()[(index | 0) - 1];
        if (!id) missing();
        return makeStyle(doc, id);
      }
      const found = D.findStyleByName(model, String(index));
      if (!found) missing();
      return makeStyle(doc, found.id);
    };
    col.Add = function (name, type) {
      const kind = STYLE_KIND[type == null ? 1 : (type | 0)];
      if (!kind) gap('G-ENGINE', 'Styles.Add');
      let id = null;
      edit('Style', () => {
        D.touchKey(model, 'styles');
        id = D.uniqueStyleId(model, name);
        model.styles[id] = { id, name: String(name), type: kind, basedOn: kind === 'paragraph' ? 'Normal' : null, next: kind === 'paragraph' ? 'Normal' : null, pPr: {}, rPr: {}, builtin: false };
        D.stylesChanged();
        return caret;
      });
      return makeStyle(doc, id);
    };
    return col;
  }

  function writeSect(sec, fn) {
    edit('Page Setup', () => {
      if (sec.endPara) {
        D.touch(sec.endPara);
        sec.endPara.sect = fn(Object.assign({}, sec.endPara.sect || {}));
        sec.sect = sec.endPara.sect;
      } else {
        D.touchKey(D.doc, 'sect');
        const next = fn(Object.assign({}, sec.sect));
        Object.assign(sec.sect, next);
      }
      return caret;
    });
  }

  function makePageSetup(parent, sec) {
    const ps = stamp({}, parent);
    function num(name, key) {
      Object.defineProperty(ps, name, {
        enumerable: true, configurable: true,
        get() { return sec.sect[key] || 0; },
        set(v) { writeSect(sec, (s) => Object.assign({}, s, { [key]: Number(v) || 0 })); },
      });
    }
    num('PageWidth', 'pgW');
    num('PageHeight', 'pgH');
    num('TopMargin', 'mt');
    num('BottomMargin', 'mb');
    num('LeftMargin', 'ml');
    num('RightMargin', 'mr');
    num('Gutter', 'gutter');
    num('HeaderDistance', 'hdr');
    num('FooterDistance', 'ftr');
    Object.defineProperty(ps, 'Orientation', {
      enumerable: true, configurable: true,
      get() { return sec.sect.orient === 'landscape' ? 1 : 0; },
      set(v) {
        const land = (v | 0) === 1;
        writeSect(sec, (s) => {
          const next = Object.assign({}, s);
          const want = land ? 'landscape' : 'portrait';
          if ((next.orient || 'portrait') !== want) {
            const w = next.pgW;
            next.pgW = next.pgH;
            next.pgH = w;
            next.orient = want;
          }
          return next;
        });
      },
    });
    Object.defineProperty(ps, 'DifferentFirstPageHeaderFooter', {
      enumerable: true, configurable: true,
      get() { return sec.sect.titlePg ? -1 : 0; },
      set(v) {
        const t = triWrite(v);
        if (t === 'skip') return;
        writeSect(sec, (s) => Object.assign({}, s, { titlePg: !!t }));
      },
    });
    Object.defineProperty(ps, 'TextColumns', {
      enumerable: true,
      get() {
        const tc = stamp({}, ps);
        function cols() { return Object.assign({ n: 1, space: 36, eq: true }, sec.sect.cols || {}); }
        function writeCols(patch) {
          writeSect(sec, (s) => Object.assign({}, s, { cols: Object.assign(cols(), patch) }));
        }
        Object.defineProperty(tc, 'Count', { enumerable: true, get() { return cols().n || 1; }, set(v) { writeCols({ n: v | 0 }); } });
        Object.defineProperty(tc, 'Spacing', { enumerable: true, get() { return cols().space || 0; }, set(v) { writeCols({ space: Number(v) || 0 }); } });
        Object.defineProperty(tc, 'EvenlySpaced', { enumerable: true, get() { return cols().eq === false ? 0 : -1; }, set(v) { const t = triWrite(v); if (t !== 'skip') writeCols({ eq: !!t }); } });
        return tc;
      },
    });
    const PAPER_WD = { Letter: 2, Legal: 4, Executive: 5, A3: 6, A4: 7, A5: 9, 'B5 (JIS)': 11, Tabloid: 23, Statement: 22, 'Envelope #10': 25, 'Envelope DL': 37, 'Envelope C5': 34, 'Envelope Monarch': 39 };
    const PAPER_CODE = { Letter: 1, Tabloid: 3, Legal: 5, Statement: 6, Executive: 7, A3: 8, A4: 9, A5: 11, 'B5 (JIS)': 13, 'Envelope #10': 20, 'Envelope DL': 27, 'Envelope C5': 28, 'Envelope Monarch': 37 };
    const WD_PAPER = {};
    for (const paperName in PAPER_WD) WD_PAPER[PAPER_WD[paperName]] = paperName;
    Object.defineProperty(ps, 'PaperSize', {
      enumerable: true, configurable: true,
      get() { return PAPER_WD[D.paperName(sec.sect)] || 41; },
      set(v) {
        const n = v | 0;
        if (n === 41) return;
        const name = WD_PAPER[n];
        if (!name) gap('G-ENGINE', 'PaperSize');
        const spec = D.PAPER.find((p) => p[0] === name);
        if (!spec) return;
        writeSect(sec, (s) => {
          const land = s.orient === 'landscape';
          return Object.assign({}, s, { pgW: land ? spec[2] : spec[1], pgH: land ? spec[1] : spec[2], paperCode: PAPER_CODE[name] || s.paperCode });
        });
      },
    });
    const SECT_VA = { top: 0, center: 1, both: 2, justify: 2, bottom: 3 };
    const SECT_VA_FROM = { 0: 'top', 1: 'center', 2: 'both', 3: 'bottom' };
    Object.defineProperty(ps, 'VerticalAlignment', {
      enumerable: true, configurable: true,
      get() { return SECT_VA[sec.sect.vAlign || 'top'] || 0; },
      set(v) { writeSect(sec, (s) => Object.assign({}, s, { vAlign: SECT_VA_FROM[v | 0] || 'top' })); },
    });
    return ps;
  }

  function headerStory(model, sect, index, isHeader) {
    const bag = sect.refs && sect.refs[isHeader ? 'hdr' : 'ftr'];
    if (!bag) return null;
    const type = index === 2 ? 'first' : index === 3 ? 'even' : 'default';
    const id = bag[type];
    return id && model.hf ? model.hf[id] : null;
  }

  function makeHeaderFooter(doc, sec, index, isHeader) {
    const obj = stamp({}, doc);
    const story = headerStory(doc.model, sec.sect, index | 0, isHeader);
    Object.defineProperty(obj, 'Exists', { enumerable: true, get() { return story ? -1 : 0; } });
    Object.defineProperty(obj, 'Range', {
      enumerable: true,
      get() {
        if (!story) {
          const empty = stamp({}, obj);
          Object.defineProperty(empty, 'Text', { enumerable: true, get() { return ''; }, set() { gap('G-ENGINE', 'HeaderFooter'); } });
          return empty;
        }
        return makeRange(obj, { kind: 'story', story });
      },
    });
    denyGet(obj, ['LinkToPrevious'], 'G-ENGINE');
    return obj;
  }

  function makeSection(doc, sec) {
    const obj = stamp({}, doc);
    obj.Range = makeRange(obj, { kind: 'section', sec });
    Object.defineProperty(obj, 'PageSetup', { enumerable: true, get() { return makePageSetup(obj, sec); } });
    obj.Headers = function (index) { return makeHeaderFooter(doc, sec, index == null ? 1 : index, true); };
    obj.Footers = function (index) { return makeHeaderFooter(doc, sec, index == null ? 1 : index, false); };
    const pageEdges = [
      { key: 'top', index: -1 }, { key: 'left', index: -2 }, { key: 'bottom', index: -3 }, { key: 'right', index: -4 },
    ];
    Object.defineProperty(obj, 'Borders', {
      enumerable: true,
      get() {
        return bordersOf(obj, {
          page: true,
          edges: pageEdges,
          enable: ['top', 'left', 'bottom', 'right'],
          outside: ['top', 'left', 'bottom', 'right'],
          inside: [],
          getBag() { return sec.sect.borders || {}; },
          setBag(bag) { writeSect(sec, (s) => Object.assign({}, s, { borders: bag })); },
        });
      },
    });
    Object.defineProperty(obj, 'ProtectedForForms', {
      enumerable: true, configurable: true,
      get() { return sec.sect.formProt ? -1 : 0; },
      set(v) { const t = triWrite(v); if (t === 'skip') return; writeSect(sec, (s) => Object.assign({}, s, { formProt: !!t })); },
    });
    return obj;
  }

  function sectionsOf(doc) {
    const col = function (i) { return col.Item(i); };
    stamp(col, doc);
    function list() { return D.sections(doc.model); }
    Object.defineProperty(col, 'Count', { enumerable: true, get() { return list().length; } });
    col.Item = function (index) {
      const s = list()[(index | 0) - 1];
      if (!s) missing();
      return makeSection(doc, s);
    };
    return col;
  }

  function makeDocProp(doc, name, custom) {
    const obj = stamp({}, doc);
    Object.defineProperty(obj, 'Name', { enumerable: true, get() { return String(name); } });
    Object.defineProperty(obj, 'Value', {
      enumerable: true, configurable: true,
      get() {
        const model = doc.model;
        if (custom) return model.custom && model.custom[name] != null ? model.custom[name] : '';
        const key = BUILTIN_PROPS[name];
        if (!key) gap('G-ENGINE', 'BuiltInDocumentProperties');
        return (model.props && model.props[key]) || '';
      },
      set(v) {
        if (custom && typeof v !== 'string') gap('G-ENGINE', 'CustomDocumentProperties');
        edit('Properties', () => {
          const model = doc.model;
          if (custom) {
            D.touchKey(model, 'custom');
            model.custom = model.custom || {};
            model.custom[name] = String(v);
          } else {
            const key = BUILTIN_PROPS[name];
            if (!key) gap('G-ENGINE', 'BuiltInDocumentProperties');
            D.touchKey(model, 'props');
            model.props = model.props || {};
            model.props[key] = String(v);
          }
          return caret;
        });
      },
    });
    return obj;
  }

  function builtinProps(doc) {
    const col = function (name) { return col.Item(name); };
    stamp(col, doc);
    col.Item = function (name) { return makeDocProp(doc, name, false); };
    return col;
  }

  function customProps(doc) {
    const col = function (name) { return col.Item(name); };
    stamp(col, doc);
    function names() { return Object.keys(doc.model.custom || {}); }
    Object.defineProperty(col, 'Count', { enumerable: true, get() { return names().length; } });
    col.Item = function (name) {
      if (typeof name === 'number') {
        const k = names()[(name | 0) - 1];
        if (k == null) missing();
        return makeDocProp(doc, k, true);
      }
      return makeDocProp(doc, name, true);
    };
    col.Add = function (name, linkToContent, type, value) {
      if (linkToContent) gap('G-ENGINE', 'CustomDocumentProperties');
      if (type != null && type !== 4) gap('G-ENGINE', 'CustomDocumentProperties');
      if (typeof value !== 'string') gap('G-ENGINE', 'CustomDocumentProperties');
      const prop = makeDocProp(doc, name, true);
      prop.Value = value;
      return prop;
    };
    return col;
  }

  function storyByType(model, t) {
    if (t === 1) return model.main;
    function bag(obj) {
      const blocks = [];
      for (const k in obj || {}) for (const b of (obj[k].blocks || [])) blocks.push(b);
      return blocks.length ? { blocks, _synthetic: true } : null;
    }
    if (t === 2) return bag(model.fn);
    if (t === 3) return bag(model.en);
    if (t === 4) return bag(model.comments);
    if (t === 7 || t === 10 || t === 6) {
      for (const id in model.hf || {}) if (model.hf[id] && model.hf[id].kind === 'hdr') return model.hf[id];
      return null;
    }
    if (t === 9 || t === 11 || t === 8) {
      for (const id in model.hf || {}) if (model.hf[id] && model.hf[id].kind === 'ftr') return model.hf[id];
      return null;
    }
    return null;
  }

  function finishClose(doc) {
    const i = wraps.indexOf(doc);
    if (i >= 0) wraps.splice(i, 1);
    if (activeWrap === doc) {
      activeWrap = wraps[wraps.length - 1] || null;
      if (activeWrap) D.doc = activeWrap.model;
      else {
        const model = D.newDoc();
        D.doc = model;
        D.resetHistory();
        activeWrap = adopt(model);
      }
      caret = null;
      clearPending();
    }
  }

  function attachDocument(doc, model) {
    Object.defineProperty(doc, 'Saved', {
      enumerable: true, configurable: true,
      get() { return !model.dirty; },
      set(v) { model.dirty = !(v === true || v === -1); },
    });
    Object.defineProperty(doc, 'Content', {
      enumerable: true,
      get() { requireActive(doc); return makeRange(doc, { kind: 'story', story: model.main }); },
    });
    Object.defineProperty(doc, 'Sentences', {
      enumerable: true,
      get() { requireActive(doc); return extentCollection(doc, () => mainParas(model), 'sentence'); },
    });
    Object.defineProperty(doc, 'Words', {
      enumerable: true,
      get() { requireActive(doc); return extentCollection(doc, () => mainParas(model), 'word'); },
    });
    Object.defineProperty(doc, 'Styles', { enumerable: true, get() { requireActive(doc); return stylesOf(doc); } });
    Object.defineProperty(doc, 'Sections', { enumerable: true, get() { requireActive(doc); return sectionsOf(doc); } });
    Object.defineProperty(doc, 'Fields', { enumerable: true, get() { requireActive(doc); return fieldsOf(doc, model.main); } });
    Object.defineProperty(doc, 'Hyperlinks', { enumerable: true, get() { requireActive(doc); return hyperlinksOf(doc, doc); } });
    Object.defineProperty(doc, 'Comments', { enumerable: true, get() { requireActive(doc); return commentsOf(doc); } });
    Object.defineProperty(doc, 'Footnotes', { enumerable: true, get() { requireActive(doc); return notesOf(doc, 'fn'); } });
    Object.defineProperty(doc, 'Endnotes', { enumerable: true, get() { requireActive(doc); return notesOf(doc, 'en'); } });
    Object.defineProperty(doc, 'InlineShapes', { enumerable: true, get() { requireActive(doc); return shapesOf(doc, false); } });
    Object.defineProperty(doc, 'Shapes', { enumerable: true, get() { requireActive(doc); return shapesOf(doc, true); } });
    Object.defineProperty(doc, 'Lists', { enumerable: true, get() { requireActive(doc); return listsOf(doc); } });
    Object.defineProperty(doc, 'ListParagraphs', { enumerable: true, get() { requireActive(doc); return listParagraphsOf(doc); } });
    Object.defineProperty(doc, 'Revisions', { enumerable: true, get() { requireActive(doc); return revisionsOf(doc); } });
    Object.defineProperty(doc, 'BuiltInDocumentProperties', { enumerable: true, get() { requireActive(doc); return builtinProps(doc); } });
    Object.defineProperty(doc, 'CustomDocumentProperties', { enumerable: true, get() { requireActive(doc); return customProps(doc); } });
    Object.defineProperty(doc, 'StoryRanges', {
      enumerable: true,
      get() {
        requireActive(doc);
        const col = function (t) { return col.Item(t); };
        stamp(col, doc);
        col.Item = function (type) {
          const story = storyByType(model, type | 0);
          if (!story) gap('G-ENGINE', 'StoryRanges');
          return makeRange(col, { kind: 'story', story });
        };
        return col;
      },
    });
    Object.defineProperty(doc, 'TrackRevisions', {
      enumerable: true, configurable: true,
      get() { return model.settings && model.settings.track ? -1 : 0; },
      set(v) {
        const t = triWrite(v);
        if (t === 'skip') return;
        edit('Track Changes', () => {
          D.touchKey(model, 'settings');
          model.settings = model.settings || {};
          model.settings.track = !!t;
          return caret;
        });
      },
    });
    doc.ComputeStatistics = function (stat) {
      requireActive(doc);
      const n = stat | 0;
      if (n === 1) gap('G-ENGINE', 'ComputeStatistics');
      if (n === 2) {
        if (L.layout && Array.isArray(L.layout.pages) && L.layout.pages.length) return L.layout.pages.length;
        gap('G-ENGINE', 'ComputeStatistics');
      }
      const paras = mainParas(model);
      if (n === 4) return paras.length;
      let chars = 0, withSpaces = 0, words = 0;
      for (const p of paras) {
        const t = D.plainText(p);
        withSpaces += t.length + 1;
        chars += t.replace(/\s/g, '').length;
        const m = t.match(/[\p{L}\p{N}_'’]+/gu);
        if (m) words += m.length;
      }
      if (n === 0) return words;
      if (n === 3) return chars;
      if (n === 5) return withSpaces;
      gap('G-ENGINE', 'ComputeStatistics');
    };
    doc.Select = function () {
      requireActive(doc);
      const paras = mainParas(model);
      if (!paras.length) return;
      placeSpan(D.pos(paras[0], 0), D.pos(paras[paras.length - 1], D.plen(paras[paras.length - 1])));
    };
    doc.Close = function (saveChanges) {
      requireActive(doc);
      if (arguments.length > 1) gap('G-HOST', 'Close');
      const mode = saveChanges == null ? 0 : (saveChanges | 0);
      if (mode === -2) gap('G-HOST', 'Close');
      if (mode === -1) {
        const saved = doc.Save();
        if (saved && typeof saved.then === 'function') return saved.then(() => finishClose(doc));
      } else if (mode !== 0) gap('G-ENGINE', 'Close');
      finishClose(doc);
    };
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
    doc.PrintOut = function () {
      if (!L.hostAudit || !L.hostAudit.has('Print')) gap('G-HOST', 'PrintOut');
      return L.hostAudit.printPdf(() => L.app.buildPDF());
    };
    doc.ExportAsFixedFormat = function (format, fileName) {
      if (!L.hostAudit || !L.hostAudit.has('FixedFormat')) gap('G-HOST', 'ExportAsFixedFormat');
      return L.hostAudit.exportPdf(format, fileName, () => L.app.buildPDF());
    };
    doc.FollowHyperlink = function (address) {
      if (!L.hostAudit || !L.hostAudit.has('Hyperlink')) gap('G-HOST', 'FollowHyperlink');
      L.hostAudit.rejectExtra(arguments, 1, ['Address', 'SubAddress', 'NewWindow', 'AddHistory', 'ExtraInfo', 'Method', 'HeaderInfo']);
      return L.hostAudit.follow(address);
    };
    denyFn(doc, ['UndoClear', 'Protect', 'Unprotect'], 'G-ENGINE');
    denyFn(doc, ['PrintPreview'], 'G-HOST');
    doc.Compare = function (name) {
      if (!L.hostAudit || typeof L.hostAudit.readBytes !== 'function') gap('G-HOST', 'Compare');
      if (!L.hostAudit.has('Open') && !L.hostAudit.has('InsertFile')) gap('G-HOST', 'Compare');
      if (!L.docx || typeof L.docx.read !== 'function' || !L.review || typeof L.review.compareWith !== 'function') gap('G-ENGINE', 'Compare');
      const ability = L.hostAudit.has('Open') ? 'Open' : 'InsertFile';
      return L.hostAudit.readBytes(name, ability, 'AlwaysAllow').then((bytes) => {
        return Promise.resolve(L.docx.read(bytes)).then((res) => {
          const other = res.doc || res;
          return L.review.compareWith(other, (other.props && (other.props.lastModifiedBy || other.props.creator)) || String(name || ''));
        });
      });
    };
  }

  function listsOf(doc) {
    const col = function (i) { return col.Item(i); };
    stamp(col, doc);
    function ids() {
      const out = [];
      for (const p of mainParas(doc.model)) {
        const num = p.pPr && p.pPr.num;
        if (num && num.id && num.id !== '0' && out.indexOf(String(num.id)) < 0) out.push(String(num.id));
      }
      return out;
    }
    Object.defineProperty(col, 'Count', { enumerable: true, get() { return ids().length; } });
    col.Item = function (index) {
      const id = ids()[(index | 0) - 1];
      if (!id) missing();
      const obj = stamp({}, col);
      Object.defineProperty(obj, 'ListParagraphs', { enumerable: true, get() { return listParagraphsOf(doc, id); } });
      Object.defineProperty(obj, 'Range', {
        enumerable: true,
        get() {
          const ps = mainParas(doc.model).filter((p) => p.pPr && p.pPr.num && String(p.pPr.num.id) === String(id));
          if (!ps.length) return makeRange(obj, null);
          const last = ps[ps.length - 1];
          return makeRange(obj, { kind: 'span', a: D.pos(ps[0], 0), f: D.pos(last, D.plen(last)) });
        },
      });
      obj.CountNumberedItems = function () { return listParagraphsOf(doc, id).Count; };
      return obj;
    };
    denyGet(col, ['ListTemplates'], 'G-ENGINE');
    return col;
  }

  function listParagraphsOf(doc, onlyId) {
    const col = function (i) { return col.Item(i); };
    stamp(col, doc);
    function paras() {
      return mainParas(doc.model).filter((p) => {
        const num = p.pPr && p.pPr.num;
        if (!num || !num.id || num.id === '0') return false;
        return onlyId == null || String(num.id) === String(onlyId);
      });
    }
    Object.defineProperty(col, 'Count', { enumerable: true, get() { return paras().length; } });
    col.Item = function (index) {
      const p = paras()[(index | 0) - 1];
      if (!p) missing();
      return makeParagraph(doc, p);
    };
    return col;
  }

  function collectFields(container) {
    const out = [];
    const open = new Map();
    const paras = [];
    if (container && container.blocks) D.walk(container, (b) => { if (b.t === 'p') paras.push(b); });
    for (const p of paras) {
      let o = 0;
      for (const it of p.runs) {
        if (it.t === 'fb') open.set(it.fid, { fid: it.fid, instr: it.instr || '', result: '', closed: false, p, o, lock: it.lock });
        else if (it.t === 'text') { for (const f of open.values()) if (!f.closed) f.result += it.text; }
        else if (it.t === 'fe' && open.has(it.fid)) {
          const f = open.get(it.fid);
          f.closed = true;
          f.endP = p;
          f.endO = o;
          out.push(f);
          open.delete(it.fid);
        }
        o += D.ilen(it);
      }
    }
    return out;
  }

  function fieldById(model, fid) {
    let hit = null;
    D.walk(model.main, (p) => {
      if (p.t !== 'p' || hit) return;
      for (const it of p.runs) if (it.t === 'fb' && it.fid === fid) hit = { p, it };
    });
    return hit;
  }

  function makeField(doc, rec) {
    const obj = stamp({}, doc);
    obj._fid = rec.fid;
    Object.defineProperty(obj, 'Type', { enumerable: true, get() { return String(rec.instr || '').trim().split(/\s+/)[0] || ''; } });
    Object.defineProperty(obj, 'Code', {
      enumerable: true,
      get() {
        const code = stamp({}, obj);
        Object.defineProperty(code, 'Text', {
          enumerable: true,
          get() { return String(rec.instr || '').trim(); },
          set(v) {
            edit('Field', () => {
              const hit = fieldById(doc.model, rec.fid);
              if (hit) { D.touch(hit.p); hit.it.instr = ' ' + String(v).trim() + ' '; rec.instr = hit.it.instr; }
              return caret;
            });
          },
        });
        function fieldStory() {
          const again = collectFields(doc.model.main).find((f) => f.fid === rec.fid) || rec;
          if (!again.p) return null;
          return storyTextOf(D.storyOf(doc.model, again.p));
        }
        Object.defineProperty(code, 'Start', {
          enumerable: true,
          get() {
            const again = collectFields(doc.model.main).find((f) => f.fid === rec.fid) || rec;
            const st = fieldStory();
            return again.p && st ? offInStory(st, D.pos(again.p, again.o || 0)) : 0;
          },
        });
        Object.defineProperty(code, 'End', {
          enumerable: true,
          get() {
            const again = collectFields(doc.model.main).find((f) => f.fid === rec.fid) || rec;
            const st = fieldStory();
            if (!again.p || !st) return 0;
            const end = again.endP ? D.pos(again.endP, again.endO || 0) : D.pos(again.p, again.o || 0);
            return offInStory(st, end);
          },
        });
        return code;
      },
    });
    Object.defineProperty(obj, 'Result', {
      enumerable: true,
      get() {
        const res = stamp({}, obj);
        Object.defineProperty(res, 'Text', {
          enumerable: true,
          get() {
            const again = collectFields(doc.model.main).find((f) => f.fid === rec.fid);
            return again ? again.result : rec.result;
          },
          set(v) {
            edit('Field', () => {
              D.walk(doc.model.main, (p) => {
                if (p.t !== 'p') return;
                const runs = p.runs;
                let sep = -1, end = -1;
                for (let i = 0; i < runs.length; i++) {
                  if (runs[i].fid !== rec.fid) continue;
                  if (runs[i].t === 'fs') sep = i;
                  if (runs[i].t === 'fe') end = i;
                }
                if (sep >= 0 && end > sep) {
                  D.touch(p);
                  const text = String(v || '');
                  const mid = text ? [D.text(text)] : [];
                  runs.splice(sep + 1, end - sep - 1, ...mid);
                }
              });
              return caret;
            });
          },
        });
        return res;
      },
    });
    obj.Update = function () {
      if (L.fields && typeof L.fields.updateField === 'function') {
        edit('Update Field', () => { L.fields.updateField({ fid: rec.fid }); return caret; });
      }
    };
    obj.Delete = function () {
      edit('Field', () => {
        D.walk(doc.model.main, (p) => {
          if (p.t !== 'p') return;
          if (p.runs.some((it) => it.fid === rec.fid)) {
            D.touch(p);
            p.runs = p.runs.filter((it) => it.fid !== rec.fid);
          }
        });
        return caret;
      });
    };
    obj.Copy = function () { gap('G-ENGINE', 'Field.Copy'); };
    Object.defineProperty(obj, 'Locked', {
      enumerable: true, configurable: true,
      get() {
        const hit = fieldById(doc.model, rec.fid);
        return hit && hit.it.lock ? -1 : 0;
      },
      set(v) {
        const t = triWrite(v);
        if (t === 'skip') return;
        edit('Field', () => {
          const hit = fieldById(doc.model, rec.fid);
          if (hit) { D.touch(hit.p); hit.it.lock = t ? true : undefined; }
          return caret;
        });
      },
    });
    obj.Unlink = function () {
      edit('Unlink Fields', () => {
        D.walk(doc.model.main, (p) => {
          if (p.t !== 'p') return;
          if (!p.runs.some((it) => it.fid === rec.fid && (it.t === 'fb' || it.t === 'fs' || it.t === 'fe'))) return;
          D.touch(p);
          p.runs = p.runs.filter((it) => !(it.fid === rec.fid && (it.t === 'fb' || it.t === 'fs' || it.t === 'fe')));
        });
        return caret;
      });
    };
    obj.Select = function () {
      const again = collectFields(doc.model.main).find((f) => f.fid === rec.fid);
      if (!again || !again.p) return;
      const end = again.endP ? D.pos(again.endP, (again.endO || 0) + 1) : D.pos(again.p, (again.o || 0) + 1);
      placeSpan(D.pos(again.p, again.o || 0), end);
    };
    obj.DoClick = function () {
      const instr = String(rec.instr || '');
      if (!/^HYPERLINK\b/i.test(instr.trim())) gap('G-ENGINE', 'Field.DoClick');
      const m = instr.match(/"([^"]+)"/);
      const address = m ? m[1] : '';
      if (!L.hostAudit || !L.hostAudit.has('Hyperlink')) gap('G-HOST', 'Field.DoClick');
      return L.hostAudit.follow(address);
    };
    return obj;
  }

  function fieldsOf(doc, container) {
    const col = function (i) { return col.Item(i); };
    stamp(col, doc);
    function list() { return collectFields(container); }
    Object.defineProperty(col, 'Count', { enumerable: true, get() { return list().length; } });
    col.Item = function (index) {
      const f = list()[(index | 0) - 1];
      if (!f) missing();
      return makeField(doc, f);
    };
    col.Add = function (range, type, text) {
      const an = range && range._anchor;
      const e = an ? endsOf(an) : endsOf({ kind: 'selection' });
      if (!e) gap('G-ENGINE', 'Fields.Add');
      const instr = text != null ? String(text) : (type != null ? String(type) : 'QUOTE');
      edit('Field', () => {
        const pos = D.eqPos(e[0], e[1]) ? e[0] : O.deleteRange(e[0], e[1]);
        return O.insertField(pos, instr, '');
      });
      const all = list();
      return all.length ? makeField(doc, all[all.length - 1]) : null;
    };
    col.Update = function () {
      if (L.fields && typeof L.fields.updateAll === 'function') L.fields.updateAll();
    };
    col.ToggleShowCodes = function () { gap('G-ENGINE', 'ToggleShowCodes'); };
    return col;
  }

  function linkSpans(scope) {
    const paras = scope && scope.model && scope.model.main ? mainParas(scope.model) : parasOfAnchor(scope);
    const out = [];
    for (const p of paras) {
      let o = 0;
      for (const it of p.runs) {
        const n = D.ilen(it);
        if (n && it.rPr && it.rPr.link) {
          const last = out[out.length - 1];
          const link = it.rPr.link;
          if (last && last.p === p && last.end === o && last.link.url === link.url && last.link.anchor === link.anchor) last.end = o + n;
          else out.push({ p, start: o, end: o + n, link });
        }
        o += n;
      }
    }
    return out;
  }

  function makeHyperlink(parent, hit) {
    const obj = stamp({}, parent);
    function span() { return [D.pos(hit.p, hit.start), D.pos(hit.p, hit.end)]; }
    Object.defineProperty(obj, 'Address', {
      enumerable: true,
      get() { return hit.link.url || ''; },
      set(v) { edit('Hyperlink', () => { O.setRunProps(span()[0], span()[1], { link: Object.assign({}, hit.link, { url: String(v) }) }); hit.link.url = String(v); return caret; }); },
    });
    Object.defineProperty(obj, 'SubAddress', {
      enumerable: true,
      get() { return hit.link.anchor || ''; },
      set(v) { edit('Hyperlink', () => { O.setRunProps(span()[0], span()[1], { link: Object.assign({}, hit.link, { anchor: String(v) }) }); hit.link.anchor = String(v); return caret; }); },
    });
    Object.defineProperty(obj, 'ScreenTip', {
      enumerable: true,
      get() { return hit.link.tip || ''; },
      set(v) { edit('Hyperlink', () => { O.setRunProps(span()[0], span()[1], { link: Object.assign({}, hit.link, { tip: String(v) }) }); hit.link.tip = String(v); return caret; }); },
    });
    Object.defineProperty(obj, 'TextToDisplay', {
      enumerable: true,
      get() { return D.ptext(hit.p, hit.start, hit.end); },
      set(v) {
        edit('Hyperlink', () => {
          const end = O.insertText(O.deleteRange(span()[0], span()[1]), String(v), { link: L.clone(hit.link), style: 'Hyperlink' });
          hit.end = hit.start + String(v).length;
          return end;
        });
      },
    });
    Object.defineProperty(obj, 'Range', { enumerable: true, get() { return makeRange(obj, { kind: 'span', a: span()[0], f: span()[1] }); } });
    obj.Delete = function () {
      edit('Hyperlink', () => { O.setRunProps(span()[0], span()[1], { link: null }); return caret; });
    };
    obj.Follow = function () {
      if (!L.hostAudit || !L.hostAudit.has('Hyperlink')) gap('G-HOST', 'Hyperlink.Follow');
      return L.hostAudit.follow(hit.link.url || '');
    };
    denyGet(obj, ['Type', 'Target'], 'G-ENGINE');
    return obj;
  }

  function hyperlinksOf(parent, scope) {
    const col = function (i) { return col.Item(i); };
    stamp(col, parent);
    function list() { return linkSpans(scope); }
    Object.defineProperty(col, 'Count', { enumerable: true, get() { return list().length; } });
    col.Item = function (index) {
      const hit = list()[(index | 0) - 1];
      if (!hit) missing();
      return makeHyperlink(col, hit);
    };
    col.Add = function (range, address, subAddress, screenTip, textToDisplay, target) {
      if (target) gap('G-ENGINE', 'Hyperlinks.Add');
      const an = range && range._anchor;
      const e = an ? endsOf(an) : null;
      if (!e) gap('G-ENGINE', 'Hyperlinks.Add');
      const link = { url: address == null ? '' : String(address), anchor: subAddress == null ? '' : String(subAddress), tip: screenTip == null ? '' : String(screenTip) };
      edit('Hyperlink', () => {
        if (textToDisplay != null && String(textToDisplay) !== '') {
          const pos = D.eqPos(e[0], e[1]) ? e[0] : O.deleteRange(e[0], e[1]);
          const end = O.insertText(pos, String(textToDisplay), { link, style: 'Hyperlink' });
          return { a: pos, f: end };
        }
        if (!D.eqPos(e[0], e[1])) O.setRunProps(e[0], e[1], { link, style: 'Hyperlink' });
        return caret;
      });
      const all = list();
      return all.length ? makeHyperlink(col, all[all.length - 1]) : null;
    };
    return col;
  }

  function makeComment(doc, id) {
    const obj = stamp({}, doc);
    function rec() { return doc.model.comments[id]; }
    Object.defineProperty(obj, 'Author', { enumerable: true, get() { return (rec() && rec().author) || ''; } });
    Object.defineProperty(obj, 'Initials', {
      enumerable: true, configurable: true,
      get() { return (rec() && rec().initials) || ''; },
      set(v) { edit('Comment', () => { D.touchKey(doc.model, 'comments'); if (rec()) rec().initials = String(v || ''); return caret; }); },
    });
    Object.defineProperty(obj, 'Initial', {
      enumerable: true, configurable: true,
      get() { return obj.Initials; },
      set(v) { obj.Initials = v; },
    });
    Object.defineProperty(obj, 'Date', { enumerable: true, get() { return (rec() && rec().date) || ''; } });
    Object.defineProperty(obj, 'Range', {
      enumerable: true,
      get() { return rec() ? makeRange(obj, { kind: 'story', story: rec() }) : makeRange(obj, null); },
    });
    obj.Delete = function () {
      edit('Comment', () => {
        D.touchKey(doc.model, 'comments');
        const drop = new Set([String(id)]);
        for (const cid of Object.keys(doc.model.comments)) {
          if (doc.model.comments[cid] && drop.has(String(doc.model.comments[cid].parent))) drop.add(String(cid));
        }
        for (const cid of drop) delete doc.model.comments[cid];
        D.walk(doc.model.main, (p) => {
          if (p.t !== 'p') return;
          if (p.runs.some((it) => (it.t === 'cs' || it.t === 'ce') && drop.has(String(it.id)))) {
            D.touch(p);
            p.runs = p.runs.filter((it) => !((it.t === 'cs' || it.t === 'ce') && drop.has(String(it.id))));
          }
        });
        doc.model._idxDirty = true;
        return caret;
      });
    };
    obj.Reply = function (text) {
      const d = doc.model;
      let n = 0;
      while (d.comments[n] != null) n++;
      const nid = String(n);
      edit('Reply', () => {
        D.touchKey(d, 'comments');
        d.comments[nid] = { kind: 'cmt', id: nid, author: (typeof O.author === 'function' ? O.author() : 'Quire'), initials: 'Q', date: new Date().toISOString().replace(/\.\d+Z$/, 'Z'), blocks: [D.para([D.text(String(text || ''))], { style: 'CommentText' })], parent: id };
        D.walk(d.main, (p) => {
          if (p.t !== 'p') return;
          let changed = false;
          const runs = [];
          for (const it of p.runs) {
            runs.push(it);
            if ((it.t === 'cs' || it.t === 'ce') && String(it.id) === String(id)) { runs.push(D.item(it.t, { id: nid })); changed = true; }
          }
          if (changed) { D.touch(p); p.runs = runs; }
        });
        d._idxDirty = true;
        return caret;
      });
      return makeComment(doc, nid);
    };
    obj.AddReply = function (text) { return obj.Reply(text); };
    Object.defineProperty(obj, 'Scope', {
      enumerable: true,
      get() {
        let a = null, f = null;
        D.walk(doc.model.main, (p) => {
          if (p.t !== 'p') return;
          let o = 0;
          for (const it of p.runs) {
            if (String(it.id) === String(id) && (it.t === 'cs' || it.t === 'ce')) {
              if (it.t === 'cs' && !a) a = D.pos(p, o);
              if (it.t === 'ce') f = D.pos(p, o);
            }
            o += D.ilen(it);
          }
        });
        if (!a) return makeRange(obj, null);
        return makeRange(obj, { kind: 'span', a, f: f || a });
      },
    });
    obj.Edit = function () { gap('G-ENGINE', 'Comment.Edit'); };
    return obj;
  }

  function commentsOf(doc) {
    const col = function (i) { return col.Item(i); };
    stamp(col, doc);
    function ids() { return Object.keys(doc.model.comments || {}).sort((a, b) => Number(a) - Number(b)); }
    Object.defineProperty(col, 'Count', { enumerable: true, get() { return ids().length; } });
    col.Item = function (index) {
      const id = ids()[(index | 0) - 1];
      if (id == null) missing();
      return makeComment(doc, id);
    };
    col.Add = function (range, text) {
      const an = range && range._anchor;
      const e = an ? endsOf(an) : endsOf({ kind: 'selection' });
      if (!e) gap('G-ENGINE', 'Comments.Add');
      const d = doc.model;
      let n = 0;
      while (d.comments[n] != null) n++;
      const id = String(n);
      edit('Comment', () => {
        D.touchKey(d, 'comments');
        d.comments[id] = { kind: 'cmt', id, author: (typeof O.author === 'function' ? O.author() : 'Quire'), initials: (typeof O.initials === 'function' ? O.initials() : 'Q'), date: new Date().toISOString().replace(/\.\d+Z$/, 'Z'), blocks: [D.para([D.text(String(text || ''))], { style: 'CommentText' })] };
        D.touch(e[1].p);
        e[1].p.runs.splice(D.splitAt(e[1].p, e[1].o), 0, D.item('ce', { id }));
        D.touch(e[0].p);
        e[0].p.runs.splice(D.splitAt(e[0].p, e[0].o), 0, D.item('cs', { id }));
        d._idxDirty = true;
        return caret;
      });
      return makeComment(doc, id);
    };
    col.AddReply = function (comment, text) {
      const target = comment && typeof comment.Reply === 'function' ? comment : (comment && comment._id != null ? makeComment(doc, comment._id) : null);
      if (!target) gap('G-ENGINE', 'Comments.AddReply');
      return target.Reply(text);
    };
    return col;
  }

  function makeNote(doc, kind, id) {
    const obj = stamp({}, doc);
    function rec() { return doc.model[kind][id]; }
    Object.defineProperty(obj, 'Index', { enumerable: true, get() { return Number(id) || 0; } });
    Object.defineProperty(obj, 'Range', {
      enumerable: true,
      get() { return rec() ? makeRange(obj, { kind: 'story', story: rec() }) : makeRange(obj, null); },
    });
    obj.Delete = function () {
      edit(kind === 'en' ? 'Endnote' : 'Footnote', () => {
        D.touchKey(doc.model, kind);
        delete doc.model[kind][id];
        D.walk(doc.model.main, (p) => {
          if (p.t !== 'p') return;
          if (p.runs.some((it) => it.t === kind && String(it.id) === String(id) && !it.self)) {
            D.touch(p);
            p.runs = p.runs.filter((it) => !(it.t === kind && String(it.id) === String(id)));
          }
        });
        doc.model._idxDirty = true;
        return caret;
      });
    };
    Object.defineProperty(obj, 'Reference', {
      enumerable: true,
      get() {
        let hit = null;
        D.walk(doc.model.main, (p) => {
          if (p.t !== 'p' || hit) return;
          let o = 0;
          for (const it of p.runs) {
            if (it.t === kind && String(it.id) === String(id) && !it.self) { hit = D.pos(p, o); return; }
            o += D.ilen(it);
          }
        });
        if (!hit) return makeRange(obj, null);
        return makeRange(obj, { kind: 'span', a: hit, f: D.pos(hit.p, Math.min(hit.o + D.ilen({ t: kind }), D.plen(hit.p))) });
      },
    });
    return obj;
  }

  function notesOf(doc, kind) {
    const col = function (i) { return col.Item(i); };
    stamp(col, doc);
    function ids() { return Object.keys(doc.model[kind] || {}).sort((a, b) => Number(a) - Number(b)); }
    Object.defineProperty(col, 'Count', { enumerable: true, get() { return ids().length; } });
    col.Item = function (index) {
      const id = ids()[(index | 0) - 1];
      if (id == null) missing();
      return makeNote(doc, kind, id);
    };
    col.Add = function (range, reference, text) {
      const body = text != null ? text : reference;
      const d = doc.model;
      let n = 1;
      while (d[kind][n] != null || d[kind][String(n)] != null) n++;
      const id = String(n);
      const an = range && range._anchor;
      const e = an ? endsOf(an) : endsOf({ kind: 'selection' });
      if (!e) gap('G-ENGINE', kind === 'en' ? 'Endnotes.Add' : 'Footnotes.Add');
      const refStyle = kind === 'en' ? 'EndnoteReference' : 'FootnoteReference';
      const textStyle = kind === 'en' ? 'EndnoteText' : 'FootnoteText';
      edit(kind === 'en' ? 'Endnote' : 'Footnote', () => {
        D.touchKey(d, kind);
        d[kind][id] = { kind, id, blocks: [D.para([D.item(kind, { id, self: true }, { style: refStyle }), D.text(' ' + String(body || ''))], { style: textStyle })] };
        O.insertItem(e[1], D.item(kind, { id }, { style: refStyle }));
        d._idxDirty = true;
        return caret;
      });
      return makeNote(doc, kind, id);
    };
    col.Convert = function () {
      const from = kind;
      const to = from === 'fn' ? 'en' : 'fn';
      edit('Convert Notes', () => {
        const d = doc.model;
        D.touchKey(d, 'fn');
        D.touchKey(d, 'en');
        const map = {};
        let next = 1;
        for (const nid of Object.keys(d[from] || {})) {
          while (d[to][next] != null) next++;
          const id = String(next++);
          const st = d[from][nid];
          st.kind = to;
          st.id = id;
          D.walk(st, (b) => {
            if (b.t !== 'p') return;
            for (const it of b.runs) if (it.t === from && it.self) {
              it.t = to;
              it.id = id;
              it.rPr = Object.assign({}, it.rPr, { style: to === 'fn' ? 'FootnoteReference' : 'EndnoteReference' });
            }
            if (b.pPr && b.pPr.style === (from === 'fn' ? 'FootnoteText' : 'EndnoteText')) b.pPr.style = to === 'fn' ? 'FootnoteText' : 'EndnoteText';
          });
          d[to][id] = st;
          map[nid] = id;
        }
        d[from] = {};
        D.walk(d.main, (b) => {
          if (b.t !== 'p') return;
          let changed = false;
          for (const it of b.runs) if (it.t === from && !it.self && map[it.id]) {
            if (!changed) { D.touch(b); changed = true; }
            it.t = to;
            it.id = map[it.id];
            it.rPr = Object.assign({}, it.rPr, { style: to === 'fn' ? 'FootnoteReference' : 'EndnoteReference' });
          }
        });
        d._idxDirty = true;
        return caret;
      });
    };
    return col;
  }

  function collectRevs(model) {
    const out = [];
    D.walk(model.main, (p) => {
      if (p.t !== 'p') return;
      let o = 0;
      for (const it of p.runs) {
        const n = D.ilen(it);
        if (n && it.rPr && (it.rPr.ins || it.rPr.del)) out.push({ p, start: o, end: o + n, kind: it.rPr.del ? 'del' : 'ins' });
        o += n;
      }
    });
    return out;
  }

  function applyRev(rec, accept) {
    edit(accept ? 'Accept' : 'Reject', () => {
      const a = D.pos(rec.p, rec.start), b = D.pos(rec.p, rec.end);
      const drop = (accept && rec.kind === 'del') || (!accept && rec.kind === 'ins');
      if (drop) return O.deleteRange(a, b);
      O.setRunProps(a, b, rec.kind === 'del' ? { del: null } : { ins: null });
      return caret;
    });
  }

  function revisionsOf(doc) {
    const col = function (i) { return col.Item(i); };
    stamp(col, doc);
    function list() { return collectRevs(doc.model); }
    Object.defineProperty(col, 'Count', { enumerable: true, get() { return list().length; } });
    col.Item = function (index) {
      const rec = list()[(index | 0) - 1];
      if (!rec) missing();
      const obj = stamp({}, col);
      Object.defineProperty(obj, 'Type', { enumerable: true, get() { return rec.kind === 'del' ? 2 : 1; } });
      obj.Accept = function () { applyRev(rec, true); };
      obj.Reject = function () { applyRev(rec, false); };
      return obj;
    };
    col.AcceptAll = function () { const all = list(); for (let i = all.length - 1; i >= 0; i--) applyRev(all[i], true); };
    col.RejectAll = function () { const all = list(); for (let i = all.length - 1; i >= 0; i--) applyRev(all[i], false); };
    return col;
  }

  function shapeItems(floating) {
    const out = [];
    if (!D.doc || !D.doc.main) return out;
    D.walk(D.doc.main, (p) => {
      if (p.t !== 'p') return;
      for (const it of p.runs) if ((it.t === 'shape' || it.t === 'img') && (!!it.float) === !!floating) out.push({ p, it });
    });
    return out;
  }

  function makeShape(doc, host, it) {
    const obj = stamp({}, doc);
    function touch() { if (host) D.touch(host); }
    Object.defineProperty(obj, 'Width', {
      enumerable: true,
      get() { return it.w || 0; },
      set(v) { edit('Shape', () => { touch(); it.w = Number(v) || 0; return caret; }); },
    });
    Object.defineProperty(obj, 'Height', {
      enumerable: true,
      get() { return it.h || 0; },
      set(v) { edit('Shape', () => { touch(); it.h = Number(v) || 0; return caret; }); },
    });
    Object.defineProperty(obj, 'Fill', {
      enumerable: true,
      get() {
        const fill = stamp({}, obj);
        Object.defineProperty(fill, 'ForeColor', {
          enumerable: true,
          get() {
            const c = stamp({}, fill);
            Object.defineProperty(c, 'RGB', {
              enumerable: true,
              get() { return colorToLong(it.fill && it.fill.c); },
              set(v) {
                const hex = longToColor(v);
                edit('Shape', () => { touch(); it.fill = { t: hex ? 'solid' : 'none', c: hex ? '#' + hex : undefined, a: 1 }; return caret; });
              },
            });
            return c;
          },
        });
        return fill;
      },
    });
    Object.defineProperty(obj, 'Line', {
      enumerable: true,
      get() {
        const line = stamp({}, obj);
        Object.defineProperty(line, 'ForeColor', {
          enumerable: true,
          get() {
            const c = stamp({}, line);
            Object.defineProperty(c, 'RGB', {
              enumerable: true,
              get() { return colorToLong(it.line && it.line.c); },
              set(v) {
                const hex = longToColor(v);
                edit('Shape', () => { touch(); it.line = Object.assign({ w: 0.75, dash: 'solid' }, it.line, { c: hex ? '#' + hex : '#000000' }); return caret; });
              },
            });
            return c;
          },
        });
        return line;
      },
    });
    Object.defineProperty(obj, 'TextFrame', {
      enumerable: true,
      get() {
        const frame = stamp({}, obj);
        Object.defineProperty(frame, 'TextRange', {
          enumerable: true,
          get() {
            if (!it.tb) gap('G-ENGINE', 'TextFrame');
            return makeRange(frame, { kind: 'story', story: it.tb });
          },
        });
        return frame;
      },
    });
    obj.Delete = function () {
      edit('Shape', () => {
        if (!host) return caret;
        D.touch(host);
        host.runs = host.runs.filter((run) => run !== it);
        return caret;
      });
    };
    denyGet(obj, ['OLEFormat', 'Chart', 'GroupItems', 'Nodes'], 'G-ENGINE');
    function shapeAt() {
      if (!host || !host.runs) return null;
      let o = 0;
      for (const run of host.runs) {
        if (run === it) return D.pos(host, o);
        o += D.ilen(run);
      }
      return null;
    }
    const SHAPE_TYPE = { img: 3, chart: 12, diagram: 13 };
    Object.defineProperty(obj, 'Type', { enumerable: true, get() { return SHAPE_TYPE[it.t] || (it.t === 'group' ? 6 : 1); } });
    Object.defineProperty(obj, 'AlternativeText', {
      enumerable: true, configurable: true,
      get() { return it.alt || ''; },
      set(v) { edit('Shape', () => { touch(); it.alt = String(v || ''); return caret; }); },
    });
    Object.defineProperty(obj, 'Title', {
      enumerable: true, configurable: true,
      get() { return it.title || ''; },
      set(v) { edit('Shape', () => { touch(); it.title = String(v || ''); return caret; }); },
    });
    Object.defineProperty(obj, 'ScaleWidth', {
      enumerable: true, configurable: true,
      get() { return it.natW ? Math.round((it.w || 0) / it.natW * 100) : 100; },
      set(v) { edit('Shape', () => { touch(); it.w = (it.natW || it.w || 72) * (Number(v) || 0) / 100; return caret; }); },
    });
    Object.defineProperty(obj, 'ScaleHeight', {
      enumerable: true, configurable: true,
      get() { return it.natH ? Math.round((it.h || 0) / it.natH * 100) : 100; },
      set(v) { edit('Shape', () => { touch(); it.h = (it.natH || it.h || 72) * (Number(v) || 0) / 100; return caret; }); },
    });
    Object.defineProperty(obj, 'Left', {
      enumerable: true, configurable: true,
      get() { return (it.float && it.float.posH && it.float.posH.off) || 0; },
      set(v) {
        edit('Shape', () => {
          touch();
          it.float = it.float || { wrap: 'square', posH: {}, posV: {} };
          it.float.posH = Object.assign({}, it.float.posH, { off: Number(v) || 0 });
          return caret;
        });
      },
    });
    Object.defineProperty(obj, 'Top', {
      enumerable: true, configurable: true,
      get() { return (it.float && it.float.posV && it.float.posV.off) || 0; },
      set(v) {
        edit('Shape', () => {
          touch();
          it.float = it.float || { wrap: 'square', posH: {}, posV: {} };
          it.float.posV = Object.assign({}, it.float.posV, { off: Number(v) || 0 });
          return caret;
        });
      },
    });
    const WRAP_TO = { square: 0, tight: 1, through: 2, none: 3, topBottom: 4, inline: 7 };
    const WRAP_FROM = { 0: 'square', 1: 'tight', 2: 'through', 3: 'none', 4: 'topBottom', 7: 'inline' };
    Object.defineProperty(obj, 'WrapFormat', {
      enumerable: true,
      get() {
        const wrap = stamp({}, obj);
        Object.defineProperty(wrap, 'Type', {
          enumerable: true, configurable: true,
          get() { return it.float ? (WRAP_TO[it.float.wrap] != null ? WRAP_TO[it.float.wrap] : 0) : 7; },
          set(v) {
            const name = WRAP_FROM[v | 0];
            if (!name) gap('G-ENGINE', 'WrapFormat');
            edit('Shape', () => {
              touch();
              if (name === 'inline') { delete it.float; return caret; }
              it.float = Object.assign({ posH: { off: 0 }, posV: { off: 0 } }, it.float, { wrap: name });
              return caret;
            });
          },
        });
        return wrap;
      },
    });
    Object.defineProperty(obj, 'Range', {
      enumerable: true,
      get() {
        const at = shapeAt();
        if (!at) return makeRange(obj, null);
        return makeRange(obj, { kind: 'span', a: at, f: D.pos(at.p, Math.min(at.o + 1, D.plen(at.p))) });
      },
    });
    obj.Select = function () {
      const at = shapeAt();
      if (at) placeSpan(at, D.pos(at.p, Math.min(at.o + 1, D.plen(at.p))));
    };
    Object.defineProperty(obj, 'Hyperlink', {
      enumerable: true,
      get() {
        const link = stamp({}, obj);
        Object.defineProperty(link, 'Address', {
          enumerable: true, configurable: true,
          get() { return (it.link && it.link.url) || ''; },
          set(v) { edit('Shape', () => { touch(); it.link = Object.assign({}, it.link, { url: String(v || '') }); return caret; }); },
        });
        link.Follow = function () {
          if (!L.hostAudit || !L.hostAudit.has('Hyperlink')) gap('G-HOST', 'Hyperlink.Follow');
          return L.hostAudit.follow((it.link && it.link.url) || '');
        };
        return link;
      },
    });
    Object.defineProperty(obj, 'PictureFormat', {
      enumerable: true,
      get() {
        const pic = stamp({}, obj);
        function numProp(name, key) {
          Object.defineProperty(pic, name, {
            enumerable: true, configurable: true,
            get() { return (it.crop && it.crop[key]) || (key === 'bright' || key === 'contrast' ? it[key] : 0) || 0; },
            set(v) {
              edit('Shape', () => {
                touch();
                if (key === 'bright' || key === 'contrast') it[key] = Number(v) || 0;
                else it.crop = Object.assign({}, it.crop, { [key]: Number(v) || 0 });
                return caret;
              });
            },
          });
        }
        numProp('CropLeft', 'l');
        numProp('CropTop', 't');
        numProp('CropRight', 'r');
        numProp('CropBottom', 'b');
        numProp('Brightness', 'bright');
        numProp('Contrast', 'contrast');
        return pic;
      },
    });
    obj.ConvertToShape = function () {
      edit('Shape', () => {
        touch();
        if (!it.float) {
          if (L.drawing && typeof L.drawing.setWrap === 'function') L.drawing.setWrap(it, 'square', { x: 0, y: 0 });
          else it.float = { wrap: 'square', posH: { rel: 'column', off: 0 }, posV: { rel: 'paragraph', off: 0 } };
        }
        return caret;
      });
      return obj;
    };
    obj.ConvertToInlineShape = function () {
      edit('Shape', () => { touch(); delete it.float; return caret; });
      return obj;
    };
    return obj;
  }

  function shapesOf(doc, floating) {
    const col = function (i) { return col.Item(i); };
    stamp(col, doc);
    function list() { return shapeItems(floating); }
    Object.defineProperty(col, 'Count', { enumerable: true, get() { return list().length; } });
    col.Item = function (index) {
      const hit = list()[(index | 0) - 1];
      if (!hit) missing();
      return makeShape(doc, hit.p, hit.it);
    };
    col.AddShape = function (type, left, top, width, height) {
      if (!L.drawing || typeof L.drawing.newShape !== 'function') gap('G-ENGINE', 'AddShape');
      const geom = typeof type === 'string' ? type : ({ 1: 'rect', 5: 'roundRect', 9: 'ellipse' }[type | 0] || 'rect');
      const it = L.drawing.newShape(geom, width || 72, height || 72);
      if (floating && typeof L.drawing.setWrap === 'function') L.drawing.setWrap(it, 'square', { x: left || 0, y: top || 0 });
      let host = null;
      edit('Shape', () => {
        const c = ensureCaret();
        host = c.f.p;
        return O.insertItem(c.f, it);
      });
      return makeShape(doc, host, it);
    };
    col.AddTextbox = function (orientation, left, top, width, height) {
      if (!L.drawing || typeof L.drawing.newShape !== 'function') gap('G-ENGINE', 'AddTextbox');
      const it = L.drawing.newShape('rect', width || 72, height || 36, { textbox: true });
      if (floating && typeof L.drawing.setWrap === 'function') L.drawing.setWrap(it, 'square', { x: left || 0, y: top || 0 });
      let host = null;
      edit('Text Box', () => {
        const c = ensureCaret();
        host = c.f.p;
        return O.insertItem(c.f, it);
      });
      return makeShape(doc, host, it);
    };
    function picturePos(target) {
      if (target == null || target === false || target === 0 || target === '') {
        const c = ensureCaret();
        return c && c.a;
      }
      const point = typeof target._pictureAt === 'function' ? target._pictureAt() : null;
      const story = point && point.story;
      const off = point ? point.off : (typeof target.Start === 'number' ? target.Start : null);
      const st = story ? storyTextOf(story) : (off != null && D.doc ? storyTextOf(D.doc.main) : null);
      if (!st || off == null || !st.paras.length) {
        const e = new Error('The parameter ' + (floating ? 'Anchor' : 'Range') + ' is not available.');
        e.code = 'error';
        throw e;
      }
      return posInStory(st, off);
    }
    function floatPicture(it, left, top, width, height) {
      if (L.drawing && typeof L.drawing.setWrap === 'function') L.drawing.setWrap(it, 'square', { x: left, y: top });
      else it.float = { wrap: 'square', posH: { rel: 'column', off: left }, posV: { rel: 'paragraph', off: top } };
      if (width != null && width !== false) it.w = Number(width);
      if (height != null && height !== false) it.h = Number(height);
    }
    col.AddPicture = function (fileName, linkToFile, saveWithDocument, fourth, fifth, sixth, seventh, eighth) {
      if (linkToFile) {
        const e = new Error('The parameter LinkToFile is not available.');
        e.code = 'error';
        throw e;
      }
      if (saveWithDocument === false || saveWithDocument === 0) {
        const e = new Error('The parameter SaveWithDocument is not available.');
        e.code = 'error';
        throw e;
      }
      if (!L.hostAudit || !L.hostAudit.has('Picture')) gap('G-HOST', 'AddPicture');
      const target = floating ? eighth : fourth;
      return L.hostAudit.readBytes(fileName, 'Picture', 'RequestApproval').then((bytes) => {
        if (!L.app || typeof L.app.insertPictureFiles !== 'function') gap('G-ENGINE', 'AddPicture');
        const pos = picturePos(target);
        const file = new File([bytes], String(fileName).split(/[/\\]/).pop());
        const opts = { pos };
        if (floating) opts.prepare = (it) => floatPicture(it, Number(fourth) || 0, Number(fifth) || 0, sixth, seventh);
        return Promise.resolve(L.app.insertPictureFiles([file], opts)).then((items) => {
          const it = items && items[0];
          if (!it) return null;
          let host = null;
          for (const st of D.stories(D.doc)) {
            D.walk(st, (p) => { if (!host && p.t === 'p' && p.runs && p.runs.indexOf(it) >= 0) host = p; });
            if (host) break;
          }
          return makeShape(doc, host, it);
        });
      });
    };
    if (floating) col.AddPicture2 = function () { return col.AddPicture.apply(col, arguments); };
    col.AddChart = function () { gap('G-ENGINE', 'AddChart'); };
    return col;
  }

  function gridRegular(table) {
    const n = table.grid ? table.grid.length : 0;
    if (!n) return false;
    return table.rows.every((r) => r.cells.length === n && r.cells.every((c) => !c.tcPr || ((!c.tcPr.span || c.tcPr.span === 1) && c.tcPr.vMerge !== 'continue')));
  }

  function addRow(table, before) {
    edit('Insert Row', () => {
      D.touchTbl(table);
      const at = before ? Math.max(0, Math.min((before | 0) - 1, table.rows.length)) : table.rows.length;
      const src = table.rows[Math.min(at, table.rows.length - 1)] || table.rows[0];
      const cells = src.cells.map((c) => D.cell([D.para()], L.clone(c.tcPr || {})));
      const row = D.row(cells, {});
      table.rows.splice(at, 0, row);
      D.doc._idxDirty = true;
      return caret;
    });
  }

  function deleteRowAt(table, index) {
    edit('Delete Row', () => {
      D.touchTbl(table);
      table.rows.splice(index, 1);
      D.doc._idxDirty = true;
      if (!table.rows.length) return O.removeBlock(table);
      return caret;
    });
  }

  function addColumn(table, before) {
    if (!gridRegular(table)) gap('G-ENGINE', 'Columns.Add');
    edit('Insert Column', () => {
      D.touchTbl(table);
      const at = before ? Math.max(0, Math.min((before | 0) - 1, table.grid.length)) : table.grid.length;
      const w = table.grid[Math.min(at, table.grid.length - 1)] || 72;
      table.grid.splice(at, 0, w);
      for (const row of table.rows) row.cells.splice(at, 0, D.cell([D.para()], { w: w }));
      D.doc._idxDirty = true;
      return caret;
    });
  }

  function deleteColumnAt(table, index) {
    if (!gridRegular(table)) gap('G-ENGINE', 'Columns.Delete');
    edit('Delete Column', () => {
      D.touchTbl(table);
      table.grid.splice(index, 1);
      for (const row of table.rows) row.cells.splice(index, 1);
      D.doc._idxDirty = true;
      if (!table.grid.length) return O.removeBlock(table);
      return caret;
    });
  }

  function caretCell() {
    const c = ensureCaret();
    const info = D.info(D.doc, c.f.p);
    if (!info || !info.parents || !info.parents.length) return null;
    const top = info.parents[info.parents.length - 1];
    if (!top || !top.tbl) return null;
    return { table: top.tbl, row: top.row, cell: top.cell, rowIndex: top.tbl.rows.indexOf(top.row) + 1 };
  }

  function insertRowsAtCaret(n) {
    const hit = caretCell();
    if (!hit) gap('G-ENGINE', 'InsertRows');
    const count = n == null ? 1 : (n | 0);
    for (let i = 0; i < count; i++) addRow(hit.table, hit.rowIndex);
  }

  function insertColsAtCaret(n) {
    const hit = caretCell();
    if (!hit) gap('G-ENGINE', 'InsertColumns');
    const col = hit.row.cells.indexOf(hit.cell) + 1;
    const count = n == null ? 1 : (n | 0);
    for (let i = 0; i < count; i++) addColumn(hit.table, col);
  }

  function convertTableToText(table, sep) {
    const mode = sep == null ? 1 : (sep | 0);
    edit('Convert To Text', () => {
      const info = D.tblInfo(D.doc, table);
      if (!info) return caret;
      const paras = [];
      if (mode === 0) {
        for (const row of table.rows) for (const c of row.cells) for (const b of c.blocks) if (b.t === 'p') paras.push(b);
      } else {
        const ch = mode === 2 ? ', ' : '\t';
        for (const row of table.rows) {
          const line = row.cells.map((c) => stripOneBreak(cellText(c))).join(ch);
          paras.push(D.para(line ? [D.text(line)] : []));
        }
      }
      D.touchList(info.cont);
      const i = info.cont.blocks.indexOf(table);
      const made = paras.length ? paras : [D.para()];
      info.cont.blocks.splice(i, 1, ...made);
      D.doc._idxDirty = true;
      const p = made[0].t === 'p' ? made[0] : D.firstPara(made);
      return D.pos(p, 0);
    });
  }

  function convertSelectionToTable(separator, numRows, numColumns) {
    const e = endsOf({ kind: 'selection' });
    if (!e || D.eqPos(e[0], e[1])) gap('G-ENGINE', 'ConvertToTable');
    const text = O.textRange(e[0], e[1], '\n');
    const mode = separator == null ? 1 : (separator | 0);
    let grid;
    if (mode === 0) grid = text.split('\n').filter((s) => s.length).map((line) => [line]);
    else grid = text.split('\n').filter((s) => s.length).map((line) => line.split(mode === 2 ? ',' : '\t'));
    if (!grid.length) gap('G-ENGINE', 'ConvertToTable');
    const nc = (numColumns | 0) || Math.max.apply(null, grid.map((r) => r.length));
    const nr = (numRows | 0) || grid.length;
    while (grid.length < nr) grid.push([]);
    let table = null;
    edit('Convert To Table', () => {
      const pos = O.deleteRange(e[0], e[1]);
      const info = D.info(D.doc, pos.p);
      const width = Math.max(72, (D.doc.sect.pgW || 612) - (D.doc.sect.ml || 0) - (D.doc.sect.mr || 0));
      table = D.simpleTable(nr, nc, width);
      for (let r = 0; r < nr; r++) for (let c = 0; c < nc; c++) {
        const val = (grid[r] && grid[r][c]) || '';
        const p = table.rows[r].cells[c].blocks[0];
        if (val) p.runs = [D.text(String(val).trim())];
      }
      const cont = info.cont;
      const i = cont.blocks.indexOf(pos.p);
      D.touchList(cont);
      if (!D.plen(pos.p)) cont.blocks.splice(i, 1, table);
      else cont.blocks.splice(i + 1, 0, table);
      D.doc._idxDirty = true;
      return D.pos(table.rows[0].cells[0].blocks[0], 0);
    });
    return makeTable(activeDocument(), table);
  }

  function cellPos(table, cell) {
    for (let r = 0; r < table.rows.length; r++) {
      let c0 = (table.rows[r].trPr && table.rows[r].trPr.gridBefore) || 0;
      for (const c of table.rows[r].cells) {
        if (c === cell) return { row: r + 1, column: c0 + 1 };
        c0 += (c.tcPr && c.tcPr.span) || 1;
      }
    }
    return { row: 0, column: 0 };
  }

  function cellCollection(doc, table, cells) {
    const col = function (i) { return col.Item(i); };
    stamp(col, doc);
    Object.defineProperty(col, 'Count', { enumerable: true, get() { return cells.length; } });
    col.Item = function (index) {
      const cell = cells[(index | 0) - 1];
      if (!cell) missing();
      const wrapped = makeCell(doc, table, cell);
      wrapped.Parent = col;
      return wrapped;
    };
    col.Merge = function () {
      if (cells.length < 2) return;
      const first = makeCell(doc, table, cells[0]);
      for (let i = 1; i < cells.length; i++) first.Merge(makeCell(doc, table, cells[i]));
    };
    col.DistributeWidth = function () {
      edit('Distribute Width', () => {
        D.touchTbl(table);
        const total = table.grid.reduce((a, b) => a + b, 0) || table.grid.length || 1;
        const each = total / (table.grid.length || 1);
        table.grid = table.grid.map(() => each);
        for (const row of table.rows) for (const c of row.cells) c.tcPr = Object.assign({}, c.tcPr, { w: each });
        return caret;
      });
    };
    col.DistributeHeight = function () {
      edit('Distribute Height', () => {
        D.touchTbl(table);
        const h = table.rows.reduce((m, row) => Math.max(m, (row.trPr && row.trPr.h) || 0), 0);
        if (h) for (const row of table.rows) row.trPr = Object.assign({}, row.trPr, { h: h, hRule: row.trPr && row.trPr.hRule || 'atLeast' });
        return caret;
      });
    };
    return col;
  }

  function attachTable(obj, doc, table) {
    Object.defineProperty(obj, 'Rows', {
      enumerable: true, configurable: true,
      get() {
        const rows = function (i) { return rows.Item(i); };
        stamp(rows, obj);
        Object.defineProperty(rows, 'Count', { enumerable: true, get() { return table.rows.length; } });
        rows.Item = function (index) {
          const row = table.rows[(index | 0) - 1];
          if (!row) missing();
          const wrap = stamp({}, rows);
          Object.defineProperty(wrap, 'HeadingFormat', {
            enumerable: true,
            get() { return row.trPr && row.trPr.header ? -1 : 0; },
            set(v) {
              const t = triWrite(v);
              if (t === 'skip') return;
              edit('Row', () => { D.touchTbl(table); row.trPr = Object.assign({}, row.trPr, { header: !!t }); return caret; });
            },
          });
          const ROW_H = { auto: 0, atLeast: 1, exact: 2 };
          const ROW_H_FROM = { 0: 'auto', 1: 'atLeast', 2: 'exact' };
          Object.defineProperty(wrap, 'Height', {
            enumerable: true,
            get() { return (row.trPr && row.trPr.h) || 0; },
            set(v) { edit('Row', () => { D.touchTbl(table); row.trPr = Object.assign({}, row.trPr, { h: Number(v) || 0, hRule: (row.trPr && row.trPr.hRule) || 'atLeast' }); return caret; }); },
          });
          Object.defineProperty(wrap, 'HeightRule', {
            enumerable: true, configurable: true,
            get() { return ROW_H[(row.trPr && row.trPr.hRule) || 'auto'] || 0; },
            set(v) { edit('Row', () => { D.touchTbl(table); row.trPr = Object.assign({}, row.trPr, { hRule: ROW_H_FROM[v | 0] || 'atLeast' }); return caret; }); },
          });
          Object.defineProperty(wrap, 'AllowBreakAcrossPages', {
            enumerable: true, configurable: true,
            get() { return row.trPr && row.trPr.cantSplit ? 0 : -1; },
            set(v) { const t = triWrite(v); if (t === 'skip') return; edit('Row', () => { D.touchTbl(table); row.trPr = Object.assign({}, row.trPr, { cantSplit: !t }); return caret; }); },
          });
          Object.defineProperty(wrap, 'Alignment', {
            enumerable: true, configurable: true,
            get() { return { left: 0, center: 1, right: 2, both: 3 }[(row.trPr && row.trPr.jc) || 'left'] || 0; },
            set(v) { const jc = { 0: 'left', 1: 'center', 2: 'right', 3: 'both' }[v | 0]; if (!jc) return; edit('Row', () => { D.touchTbl(table); row.trPr = Object.assign({}, row.trPr, { jc }); return caret; }); },
          });
          wrap.Delete = function () { deleteRowAt(table, (index | 0) - 1); };
          Object.defineProperty(wrap, 'Cells', { enumerable: true, get() { return cellCollection(doc, table, row.cells); } });
          Object.defineProperty(wrap, 'Range', {
            enumerable: true,
            get() {
              const first = D.firstPara(row.cells[0].blocks);
              const lastCell = row.cells[row.cells.length - 1];
              const last = (lastCell.blocks || []).filter((b) => b.t === 'p').pop();
              if (!first || !last) return makeRange(wrap, null);
              return makeRange(wrap, { kind: 'span', a: D.pos(first, 0), f: D.pos(last, D.plen(last)) });
            },
          });
          wrap.Select = function () {
            const p = D.firstPara(row.cells[0].blocks);
            const last = row.cells[row.cells.length - 1];
            const endP = last.blocks.filter((b) => b.t === 'p').pop();
            if (p && endP) placeSpan(D.pos(p, 0), D.pos(endP, D.plen(endP)));
          };
          return wrap;
        };
        rows.Add = function (before) { addRow(table, before); return rows.Item(before ? (before | 0) : table.rows.length); };
        rows.Delete = function () {
          for (let i = table.rows.length - 1; i >= 0; i--) deleteRowAt(table, i);
        };
        return rows;
      },
    });
    Object.defineProperty(obj, 'Columns', {
      enumerable: true, configurable: true,
      get() {
        const cols = function (i) { return cols.Item(i); };
        stamp(cols, obj);
        const n = () => (table.grid && table.grid.length) || 0;
        Object.defineProperty(cols, 'Count', { enumerable: true, get() { return n(); } });
        cols.Item = function (index) {
          if ((index | 0) < 1 || (index | 0) > n()) missing();
          const wrap = stamp({}, cols);
          const at = (index | 0) - 1;
          Object.defineProperty(wrap, 'Width', {
            enumerable: true,
            get() { return table.grid[at] || 0; },
            set(v) {
              edit('Column', () => {
                D.touchTbl(table);
                table.grid[at] = Number(v) || 0;
                for (const row of table.rows) if (row.cells[at]) row.cells[at].tcPr = Object.assign({}, row.cells[at].tcPr, { w: Number(v) || 0 });
                return caret;
              });
            },
          });
          wrap.Delete = function () { deleteColumnAt(table, at); };
          wrap.Select = function () {
            const cell = table.rows[0] && table.rows[0].cells[at];
            const p = cell && D.firstPara(cell.blocks);
            if (p) placeCaret(D.pos(p, 0));
          };
          Object.defineProperty(wrap, 'Cells', {
            enumerable: true,
            get() { return cellCollection(doc, table, table.rows.map((row) => row.cells[at]).filter(Boolean)); },
          });
          return wrap;
        };
        cols.Add = function (before) { addColumn(table, before); return cols.Item(before ? (before | 0) : table.grid.length); };
        return cols;
      },
    });
    Object.defineProperty(obj, 'Style', {
      enumerable: true, configurable: true,
      get() {
        const id = table.tblPr && table.tblPr.style;
        const st = id && doc.model.styles[id];
        return st ? D.styleDisplayName(st) : (id || '');
      },
      set(v) {
        const found = D.findStyleByName(doc.model, String(v));
        edit('Table Style', () => {
          D.touchTbl(table);
          table.tblPr = Object.assign({}, table.tblPr, { style: found ? found.id : String(v) });
          return caret;
        });
      },
    });
    function lookProp(name, key, invert) {
      Object.defineProperty(obj, name, {
        enumerable: true, configurable: true,
        get() {
          const look = (table.tblPr && table.tblPr.look) || {};
          return invert ? (look[key] ? 0 : -1) : (look[key] ? -1 : 0);
        },
        set(v) {
          const t = triWrite(v);
          if (t === 'skip') return;
          edit('Table', () => {
            D.touchTbl(table);
            const look = Object.assign({}, (table.tblPr && table.tblPr.look) || {});
            look[key] = invert ? !t : !!t;
            table.tblPr = Object.assign({}, table.tblPr, { look });
            return caret;
          });
        },
      });
    }
    lookProp('ApplyStyleHeadingRows', 'firstRow');
    lookProp('ApplyStyleLastRow', 'lastRow');
    lookProp('ApplyStyleFirstColumn', 'firstCol');
    lookProp('ApplyStyleLastColumn', 'lastCol');
    lookProp('ApplyStyleRowBands', 'noHBand', true);
    lookProp('ApplyStyleColumnBands', 'noVBand', true);
    obj.Delete = function () { edit('Delete Table', () => O.removeBlock(table)); };
    obj.Select = function () {
      const p = D.firstPara(table.rows[0].cells[0].blocks);
      if (p) placeCaret(D.pos(p, 0));
    };
    obj.Split = function (beforeRow) {
      const at = (beforeRow | 0) - 1;
      if (at < 1 || at >= table.rows.length) gap('G-ENGINE', 'Table.Split');
      edit('Split Table', () => {
        const info = D.tblInfo(D.doc, table);
        if (!info) return caret;
        D.touchTbl(table);
        const moved = table.rows.splice(at);
        const copy = D.table(moved, table.grid.slice(), L.clone(table.tblPr || {}));
        D.touchList(info.cont);
        const i = info.cont.blocks.indexOf(table);
        info.cont.blocks.splice(i + 1, 0, D.para(), copy);
        D.doc._idxDirty = true;
        return caret;
      });
    };
    obj.Sort = function (excludeHeader, fieldNumber, sortFieldType, sortOrder) {
      if (sortFieldType != null && sortFieldType !== 0 && sortFieldType !== 1) gap('G-ENGINE', 'Sort');
      const desc = (sortOrder | 0) === 1;
      const col = Math.max(0, (fieldNumber | 0) - 1);
      edit('Sort', () => {
        D.touchTbl(table);
        const head = excludeHeader ? table.rows.slice(0, 1) : [];
        const body = excludeHeader ? table.rows.slice(1) : table.rows.slice();
        body.sort((a, b) => {
          const at = stripOneBreak(cellText(a.cells[col] || a.cells[0]));
          const bt = stripOneBreak(cellText(b.cells[col] || b.cells[0]));
          const cmp = sortFieldType === 1 ? (Number(at) || 0) - (Number(bt) || 0) : at.localeCompare(bt);
          return desc ? -cmp : cmp;
        });
        table.rows = head.concat(body);
        D.doc._idxDirty = true;
        return caret;
      });
    };
    obj.AutoFitBehavior = function (behavior) {
      const b = behavior | 0;
      edit('AutoFit', () => {
        D.touchTbl(table);
        table.tblPr = table.tblPr || {};
        if (b === 0) table.tblPr.layout = 'fixed';
        else if (b === 1) table.tblPr.layout = 'autofit';
        else if (b === 2) {
          table.tblPr.layout = 'autofit';
          table.tblPr.w = { type: 'pct', v: 5000 };
          const sect = D.doc.sect || {};
          const outer = Math.max(36, (sect.pgW || 612) - (sect.ml || 0) - (sect.mr || 0));
          const total = table.grid.reduce((a, c) => a + c, 0) || 1;
          table.grid = table.grid.map((g) => g * outer / total);
        } else gap('G-ENGINE', 'AutoFitBehavior');
        return caret;
      });
    };
    obj.ConvertToText = function (separator) { convertTableToText(table, separator); };
    Object.defineProperty(obj, 'Cells', {
      enumerable: true,
      get() { return cellCollection(doc, table, table.rows.reduce((all, row) => all.concat(row.cells), [])); },
    });
    Object.defineProperty(obj, 'Range', {
      enumerable: true,
      get() {
        const first = table.rows[0] && table.rows[0].cells[0] && D.firstPara(table.rows[0].cells[0].blocks);
        if (!first) return makeRange(obj, null);
        return makeRange(obj, { kind: 'table', table: table });
      },
    });
    const tblEdges = [
      { key: 'top', index: -1 }, { key: 'left', index: -2 }, { key: 'bottom', index: -3 },
      { key: 'right', index: -4 }, { key: 'insideH', index: -5 }, { key: 'insideV', index: -6 },
    ];
    Object.defineProperty(obj, 'Borders', {
      enumerable: true,
      get() {
        return bordersOf(obj, {
          edges: tblEdges,
          enable: ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'],
          outside: ['top', 'left', 'bottom', 'right'],
          inside: ['insideH', 'insideV'],
          getBag() { return (table.tblPr && table.tblPr.borders) || {}; },
          setBag(bag, reason) {
            edit('Borders', () => {
              D.touchTbl(table);
              table.tblPr = Object.assign({}, table.tblPr, { borders: bag });
              if (reason === 'enable') {
                for (const row of table.rows) for (const cell of row.cells) {
                  if (cell.tcPr && cell.tcPr.borders) cell.tcPr = Object.assign({}, cell.tcPr, { borders: undefined });
                }
              }
              return caret;
            });
          },
        });
      },
    });
    Object.defineProperty(obj, 'Shading', {
      enumerable: true,
      get() {
        return shadingOf(obj,
          () => (table.tblPr && table.tblPr.shd) || {},
          (shd) => edit('Shading', () => { D.touchTbl(table); table.tblPr = Object.assign({}, table.tblPr, { shd }); return caret; }));
      },
    });
    function tblEdit(name, patch) {
      edit(name, () => { D.touchTbl(table); table.tblPr = Object.assign({}, table.tblPr, patch(table.tblPr || {})); return caret; });
    }
    Object.defineProperty(obj, 'Spacing', {
      enumerable: true, configurable: true,
      get() { return (table.tblPr && table.tblPr.cellSpacing) || 0; },
      set(v) { tblEdit('Table', () => ({ cellSpacing: Number(v) || 0 })); },
    });
    Object.defineProperty(obj, 'Uniform', {
      enumerable: true,
      get() {
        const g = table.grid || [];
        if (g.length < 2) return -1;
        return g.every((w) => Math.abs(w - g[0]) < 0.05) ? -1 : 0;
      },
    });
    Object.defineProperty(obj, 'AllowAutoFit', {
      enumerable: true, configurable: true,
      get() { return table.tblPr && table.tblPr.layout === 'fixed' ? 0 : -1; },
      set(v) { const t = triWrite(v); if (t === 'skip') return; tblEdit('Table', () => ({ layout: t ? 'autofit' : 'fixed' })); },
    });
    Object.defineProperty(obj, 'PreferredWidth', {
      enumerable: true, configurable: true,
      get() {
        const w = table.tblPr && table.tblPr.w;
        if (!w || w.type === 'auto') return 0;
        if (w.type === 'pct') return w.v > 100 ? w.v / 50 : w.v;
        return w.v || 0;
      },
      set(v) {
        tblEdit('Table', (pr) => {
          const prev = pr.w || { type: 'dxa', v: 0 };
          const type = prev.type === 'pct' ? 'pct' : 'dxa';
          return { w: { type, v: type === 'pct' ? (Number(v) || 0) * 50 : (Number(v) || 0) } };
        });
      },
    });
    Object.defineProperty(obj, 'PreferredWidthType', {
      enumerable: true, configurable: true,
      get() { const w = table.tblPr && table.tblPr.w; return !w || w.type === 'auto' ? 1 : w.type === 'pct' ? 2 : 3; },
      set(v) {
        const type = { 1: 'auto', 2: 'pct', 3: 'dxa' }[v | 0];
        if (!type) return;
        tblEdit('Table', (pr) => ({ w: { type, v: type === 'pct' ? ((pr.w && pr.w.v) || 5000) : ((pr.w && pr.w.v) || 0) } }));
      },
    });
    function pad(name, key) {
      Object.defineProperty(obj, name, {
        enumerable: true, configurable: true,
        get() { return (table.tblPr && table.tblPr.cellMar && table.tblPr.cellMar[key]) || 0; },
        set(v) { tblEdit('Table', (pr) => ({ cellMar: Object.assign({}, pr.cellMar, { [key]: Number(v) || 0 }) })); },
      });
    }
    pad('TopPadding', 't'); pad('LeftPadding', 'l'); pad('BottomPadding', 'b'); pad('RightPadding', 'r');
    Object.defineProperty(obj, 'Alignment', {
      enumerable: true, configurable: true,
      get() { return { left: 0, center: 1, right: 2, both: 3 }[(table.tblPr && table.tblPr.jc) || 'left'] || 0; },
      set(v) { const jc = { 0: 'left', 1: 'center', 2: 'right', 3: 'both' }[v | 0]; if (jc) tblEdit('Table', () => ({ jc })); },
    });
    Object.defineProperty(obj, 'Indent', { enumerable: true, get() { return (table.tblPr && table.tblPr.ind) || 0; } });
    Object.defineProperty(obj, 'TableDirection', {
      enumerable: true, configurable: true,
      get() { return table.tblPr && table.tblPr.bidi ? 1 : 0; },
      set(v) { tblEdit('Table', () => ({ bidi: (v | 0) === 1 })); },
    });
    Object.defineProperty(obj, 'Title', {
      enumerable: true, configurable: true,
      get() { return (table.tblPr && table.tblPr.caption) || ''; },
      set(v) { tblEdit('Table', () => ({ caption: String(v || '') })); },
    });
    Object.defineProperty(obj, 'Descr', {
      enumerable: true, configurable: true,
      get() { return (table.tblPr && table.tblPr.desc) || ''; },
      set(v) { tblEdit('Table', () => ({ desc: String(v || '') })); },
    });
    Object.defineProperty(obj, 'NestingLevel', {
      enumerable: true,
      get() { const info = D.tblInfo(D.doc, table); return info ? (info.parents ? info.parents.length : 0) + 1 : 1; },
    });
  }

  function mainTables(model) {
    const out = [];
    if (model && model.main) D.walk(model.main, (b) => { if (b.t === 'tbl') out.push(b); });
    return out;
  }

  function cellText(cell) {
    const paras = (cell.blocks || []).filter((b) => b.t === 'p');
    if (!paras.length) return '\r';
    return paras.map((p) => D.ptext(p)).join('\r') + '\r';
  }

  function writeCellText(cell, raw) {
    const parts = stripOneBreak(raw).split(/\r\n|\r|\n/);
    edit('Text', () => {
      D.touchList(cell);
      let first = (cell.blocks || []).find((b) => b.t === 'p');
      if (!first) first = D.para();
      const made = parts.map((part, i) => {
        const p = i === 0 ? first : D.para([], L.clone(first.pPr || {}), L.clone(first.rPr || {}));
        D.touch(p);
        p.runs = [];
        if (part) O.insertText(D.pos(p, 0), part);
        return p;
      });
      cell.blocks = made.length ? made : [first];
      const last = cell.blocks[cell.blocks.length - 1];
      return D.pos(last, D.plen(last));
    });
  }

  function cellAt(table, row, column) {
    const rows = table.rows.map((r) => {
      let c0 = (r.trPr && r.trPr.gridBefore) || 0;
      return r.cells.map((cell) => {
        const span = (cell.tcPr && cell.tcPr.span) || 1;
        const o = { cell, c0, span };
        c0 += span;
        return o;
      });
    });
    const nCols = table.grid && table.grid.length ? table.grid.length : rows.reduce((m, r) => Math.max(m, r.reduce((s, x) => Math.max(s, x.c0 + x.span), 0)), 0);
    const grid = rows.map(() => new Array(nCols));
    rows.forEach((r, ri) => {
      for (const m of r) {
        let owner = m.cell;
        if (m.cell.tcPr && m.cell.tcPr.vMerge === 'continue') {
          for (let k = ri - 1; k >= 0; k--) {
            const above = rows[k].find((x) => x.c0 === m.c0);
            if (!above) break;
            owner = above.cell;
            if (!above.cell.tcPr || above.cell.tcPr.vMerge !== 'continue') break;
          }
        }
        for (let k = 0; k < m.span; k++) if (m.c0 + k < nCols) grid[ri][m.c0 + k] = owner;
      }
    });
    const cell = grid[row - 1] && grid[row - 1][column - 1];
    if (!cell) {
      const e = new Error('The requested member of the collection does not exist.');
      e.code = 'missing';
      throw e;
    }
    return cell;
  }

  function makeCell(doc, table, cell) {
    const obj = stamp({}, null);
    obj._cell = cell;
    obj.Range = makeRange(obj, { kind: 'cell', cell, table, doc });
    obj.Parent = null;
    const pos = () => cellPos(table, cell);
    Object.defineProperty(obj, 'RowIndex', { enumerable: true, get() { return pos().row; } });
    Object.defineProperty(obj, 'ColumnIndex', { enumerable: true, get() { return pos().column; } });
    Object.defineProperty(obj, 'Width', {
      enumerable: true, configurable: true,
      get() { return (cell.tcPr && cell.tcPr.w) || 0; },
      set(v) { edit('Cell', () => { D.touchTbl(table); cell.tcPr = Object.assign({}, cell.tcPr, { w: Number(v) || 0 }); return caret; }); },
    });
    Object.defineProperty(obj, 'Height', {
      enumerable: true, configurable: true,
      get() {
        for (const row of table.rows) if (row.cells.indexOf(cell) >= 0) return (row.trPr && row.trPr.h) || 0;
        return 0;
      },
      set(v) {
        edit('Cell', () => {
          D.touchTbl(table);
          for (const row of table.rows) if (row.cells.indexOf(cell) >= 0) row.trPr = Object.assign({}, row.trPr, { h: Number(v) || 0, hRule: 'atLeast' });
          return caret;
        });
      },
    });
    const cellEdges = [
      { key: 'top', index: -1 }, { key: 'left', index: -2 }, { key: 'bottom', index: -3 },
      { key: 'right', index: -4 }, { key: 'insideH', index: -5 }, { key: 'insideV', index: -6 },
      { key: 'tl2br', index: -7 }, { key: 'tr2bl', index: -8 },
    ];
    function cellBorders() { return (cell.tcPr && cell.tcPr.borders) || {}; }
    Object.defineProperty(obj, 'Shading', {
      enumerable: true,
      get() {
        return shadingOf(obj,
          () => (cell.tcPr && cell.tcPr.shd) || {},
          (shd) => edit('Shading', () => {
            D.touchTbl(table);
            cell.tcPr = Object.assign({}, cell.tcPr, { shd: shd });
            return caret;
          }));
      },
    });
    Object.defineProperty(obj, 'Borders', {
      enumerable: true,
      get() {
        return bordersOf(obj, {
          edges: cellEdges,
          enable: ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'],
          outside: ['top', 'left', 'bottom', 'right'],
          inside: ['insideH', 'insideV'],
          getBag: cellBorders,
          setBag(bag) {
            edit('Borders', () => {
              D.touchTbl(table);
              cell.tcPr = Object.assign({}, cell.tcPr, { borders: bag });
              return caret;
            });
          },
        });
      },
    });
    const CELL_VA = { top: 0, center: 1, bottom: 3 };
    const CELL_VA_FROM = { 0: 'top', 1: 'center', 3: 'bottom' };
    Object.defineProperty(obj, 'VerticalAlignment', {
      enumerable: true, configurable: true,
      get() { return CELL_VA[(cell.tcPr && cell.tcPr.vAlign) || 'top'] || 0; },
      set(v) { edit('Cell', () => { D.touchTbl(table); cell.tcPr = Object.assign({}, cell.tcPr, { vAlign: CELL_VA_FROM[v | 0] || 'top' }); return caret; }); },
    });
    function cellPad(name, key) {
      Object.defineProperty(obj, name, {
        enumerable: true, configurable: true,
        get() { return (cell.tcPr && cell.tcPr.mar && cell.tcPr.mar[key]) || 0; },
        set(v) { edit('Cell', () => { D.touchTbl(table); cell.tcPr = Object.assign({}, cell.tcPr, { mar: Object.assign({}, cell.tcPr && cell.tcPr.mar, { [key]: Number(v) || 0 }) }); return caret; }); },
      });
    }
    cellPad('TopPadding', 't'); cellPad('LeftPadding', 'l'); cellPad('BottomPadding', 'b'); cellPad('RightPadding', 'r');
    Object.defineProperty(obj, 'FitText', {
      enumerable: true, configurable: true,
      get() { return cell.tcPr && cell.tcPr.fitText ? -1 : 0; },
      set(v) { const t = triWrite(v); if (t === 'skip') return; edit('Cell', () => { D.touchTbl(table); cell.tcPr = Object.assign({}, cell.tcPr, { fitText: !!t }); return caret; }); },
    });
    Object.defineProperty(obj, 'WordWrap', {
      enumerable: true, configurable: true,
      get() { return cell.tcPr && cell.tcPr.noWrap ? 0 : -1; },
      set(v) { const t = triWrite(v); if (t === 'skip') return; edit('Cell', () => { D.touchTbl(table); cell.tcPr = Object.assign({}, cell.tcPr, { noWrap: !t }); return caret; }); },
    });
    obj.Merge = function (other) {
      const dropCell = other && other._cell;
      if (!dropCell || dropCell === cell) return;
      let row = null;
      for (const r of table.rows) if (r.cells.indexOf(cell) >= 0 && r.cells.indexOf(dropCell) >= 0) row = r;
      if (!row) gap('G-ENGINE', 'Merge');
      edit('Merge', () => {
        D.touchTbl(table);
        const ia = row.cells.indexOf(cell);
        const ib = row.cells.indexOf(dropCell);
        const keep = ia < ib ? cell : dropCell;
        const drop = keep === cell ? dropCell : cell;
        keep.tcPr = Object.assign({}, keep.tcPr, { span: ((keep.tcPr && keep.tcPr.span) || 1) + ((drop.tcPr && drop.tcPr.span) || 1) });
        const extra = (drop.blocks || []).filter((b) => b.t !== 'p' || D.plen(b));
        if (extra.length) keep.blocks = keep.blocks.concat(extra);
        row.cells.splice(row.cells.indexOf(drop), 1);
        D.doc._idxDirty = true;
        return caret;
      });
    };
    obj.Split = function (numRows, numColumns) {
      const nRows = numRows == null ? 1 : (numRows | 0);
      const nCols = numColumns == null ? 1 : (numColumns | 0);
      if (nRows !== 1 || nCols < 2 || !gridRegular(table)) gap('G-ENGINE', 'Split');
      const col = cellPos(table, cell).column - 1;
      edit('Split', () => {
        D.touchTbl(table);
        const each = (table.grid[col] || 72) / nCols;
        table.grid[col] = each;
        for (let k = 1; k < nCols; k++) table.grid.splice(col + k, 0, each);
        for (const row of table.rows) {
          const host = row.cells[col];
          if (!host) continue;
          host.tcPr = Object.assign({}, host.tcPr, { w: each });
          for (let k = 1; k < nCols; k++) row.cells.splice(col + k, 0, D.cell([D.para()], { w: each }));
        }
        D.doc._idxDirty = true;
        return caret;
      });
    };
    return obj;
  }

  function makeTable(doc, table) {
    const obj = stamp({}, doc);
    obj._table = table;
    obj.Parent = doc.Tables;
    Object.defineProperty(obj, 'Rows', {
      enumerable: true, configurable: true,
      get() {
        const rows = stamp({}, obj);
        Object.defineProperty(rows, 'Count', { enumerable: true, get() { return table.rows.length; } });
        return rows;
      },
    });
    Object.defineProperty(obj, 'Columns', {
      enumerable: true, configurable: true,
      get() {
        const cols = stamp({}, obj);
        const n = table.grid && table.grid.length ? table.grid.length : (table.rows[0] ? table.rows[0].cells.length : 0);
        Object.defineProperty(cols, 'Count', { enumerable: true, get() { return n; } });
        return cols;
      },
    });
    obj.Cell = function (row, column) {
      const cell = cellAt(table, row | 0, column | 0);
      const wrapped = makeCell(doc, table, cell);
      wrapped.Parent = obj;
      return wrapped;
    };
    attachTable(obj, doc, table);
    return obj;
  }

  function listBookmarks(model) {
    D.reindex(model);
    const out = [];
    const open = new Map();
    for (const st of D.stories(model)) for (const p of st.paras || []) {
      let o = 0;
      for (const it of p.runs) {
        if (it.t === 'bs' && it.name) {
          const b = { name: it.name, id: it.id, p, o, endP: p, endO: o };
          out.push(b);
          open.set(it.id, b);
        }
        if (it.t === 'be' && open.has(it.id)) {
          const b = open.get(it.id);
          b.endP = p;
          b.endO = o;
          open.delete(it.id);
        }
        o += D.ilen(it);
      }
    }
    return out;
  }

  function removeBookmark(b) {
    if (L.fields && typeof L.fields.removeBookmarkItems === 'function') {
      L.fields.removeBookmarkItems(b);
      return;
    }
    const seen = new Set();
    for (const p of [b.p, b.endP]) {
      if (!p || seen.has(p)) continue;
      seen.add(p);
      D.touch(p);
      p.runs = p.runs.filter((it) => !((it.t === 'bs' || it.t === 'be') && it.id === b.id));
    }
  }

  function insertBookmark(name, a, b) {
    if (L.fields && typeof L.fields.addBookmark === 'function') {
      L.fields.addBookmark(name, a, b);
      return;
    }
    const d = D.doc;
    [a, b] = D.order(d, a, b);
    const old = listBookmarks(d).find((x) => x.name.toLowerCase() === name.toLowerCase());
    if (old) removeBookmark(old);
    const id = String(D.nid());
    D.touch(a.p);
    const ia = D.splitAt(a.p, a.o);
    a.p.runs.splice(ia, 0, D.item('bs', { id, name }));
    D.touch(b.p);
    let ib = D.splitAt(b.p, b.o);
    if (a.p === b.p && ib <= ia) ib = ia + 1;
    b.p.runs.splice(ib, 0, D.item('be', { id }));
  }

  function makeBookmark(doc, b) {
    const obj = stamp({}, doc);
    obj._bm = b;
    Object.defineProperty(obj, 'Name', { enumerable: true, get() { return b.name; } });
    Object.defineProperty(obj, 'Empty', { enumerable: true, get() { return b.p === b.endP && b.o === b.endO; } });
    Object.defineProperty(obj, 'Range', {
      enumerable: true,
      get() { return makeRange(obj, { kind: 'span', a: D.pos(b.p, b.o), f: D.pos(b.endP, b.endO), doc }); },
    });
    function bmOff(end) {
      const st = storyTextOf(D.storyOf(D.doc, b.p));
      return offInStory(st, end ? D.pos(b.endP, b.endO) : D.pos(b.p, b.o));
    }
    function bmSet(end, v) {
      edit('Bookmark', () => {
        const st = storyTextOf(D.storyOf(D.doc, b.p));
        const pos = posInStory(st, v | 0);
        if (end) { b.endP = pos.p; b.endO = pos.o; }
        else { b.p = pos.p; b.o = pos.o; }
        return caret;
      });
    }
    Object.defineProperty(obj, 'Start', { enumerable: true, configurable: true, get() { return bmOff(false); }, set(v) { bmSet(false, v); } });
    Object.defineProperty(obj, 'End', { enumerable: true, configurable: true, get() { return bmOff(true); }, set(v) { bmSet(true, v); } });
    obj.Delete = function () { edit('Bookmark', () => { removeBookmark(b); return caret; }); };
    obj.Select = function () {
      caret = { a: D.pos(b.p, b.o), f: D.pos(b.endP, b.endO) };
      if (L.ed) {
        L.ed.sel = caret;
        if (typeof L.ed.setSel === 'function') L.ed.setSel(caret.a, caret.f, { noScroll: true });
      }
    };
    obj.Copy = function () { gap('G-ENGINE', 'Bookmark.Copy'); };
    return obj;
  }

  function bookmarksOf(doc) {
    const col = function (index) { return col.Item(index); };
    stamp(col, doc);
    const list = () => listBookmarks(doc.model);
    const byName = (name) => list().find((b) => b.name.toLowerCase() === String(name).toLowerCase());
    Object.defineProperty(col, 'Count', { enumerable: true, get() { return list().length; } });
    col.Item = function (index) {
      if (typeof index === 'string') {
        const b = byName(index);
        if (!b) { const e = new Error('The requested member of the collection does not exist.'); e.code = 'missing'; throw e; }
        return makeBookmark(doc, b);
      }
      const all = list();
      const i = index | 0;
      if (i < 1 || i > all.length) { const e = new Error('The requested member of the collection does not exist.'); e.code = 'missing'; throw e; }
      return makeBookmark(doc, all[i - 1]);
    };
    col.Exists = function (name) { return !!byName(name); };
    col.Add = function (name, range) {
      const anchor = range && range._anchor ? range._anchor : { kind: 'selection' };
      const ends = endsOf(anchor);
      if (!ends) { const e = new Error('Bookmark needs a range'); e.code = 'missing'; throw e; }
      edit('Bookmark', () => { insertBookmark(String(name), ends[0], ends[1]); return caret; });
      const b = byName(name);
      return b ? makeBookmark(doc, b) : null;
    };
    Object.defineProperty(col, 'First', { enumerable: true, get() { return col.Count ? col.Item(1) : null; } });
    Object.defineProperty(col, 'Last', { enumerable: true, get() { return col.Count ? col.Item(col.Count) : null; } });
    return col;
  }

  function storyTextOf(story) {
    D.reindex(D.doc);
    const paras = story.paras || [];
    const starts = [];
    let s = '';
    for (const p of paras) { starts.push(s.length); s += D.ptext(p) + '\r'; }
    return { text: s, paras, starts };
  }

  function posInStory(st, off) {
    let lo = 0, hi = st.starts.length - 1;
    while (lo < hi) { const mid = (lo + hi + 1) >> 1; if (st.starts[mid] <= off) lo = mid; else hi = mid - 1; }
    const p = st.paras[lo];
    return D.pos(p, Math.min(Math.max(0, off - st.starts[lo]), D.plen(p)));
  }

  function offInStory(st, pos) {
    const i = st.paras.indexOf(pos.p);
    return i < 0 ? 0 : st.starts[i] + pos.o;
  }

  function literalMatches(text, needle, matchCase, whole) {
    if (!needle) return [];
    const hay = matchCase ? text : text.toLowerCase();
    const nd = matchCase ? needle : needle.toLowerCase();
    const word = (c) => /[\p{L}\p{N}_]/u.test(c || '');
    const out = [];
    let i = 0;
    while (i <= hay.length - nd.length) {
      const j = hay.indexOf(nd, i);
      if (j < 0) break;
      const e = j + nd.length;
      if (!whole || (!word(text[j - 1]) && !word(text[e]))) out.push([j, e]);
      i = j + Math.max(1, nd.length);
    }
    return out;
  }

  function makeFind(parent, anchor) {
    const find = stamp({}, parent);
    find.Text = '';
    find.MatchCase = false;
    find.MatchWholeWord = false;
    find.MatchWildcards = false;
    find.MatchSoundsLike = false;
    find.MatchAllWordForms = false;
    find.Forward = true;
    find.Wrap = 1;
    find.Format = false;
    find.Found = false;
    find.ClearFormatting = function () { find.Format = false; };
    find._repl = '';
    const repl = stamp({}, find);
    Object.defineProperty(repl, 'Text', {
      enumerable: true, configurable: true,
      get() { return find._repl; },
      set(v) { find._repl = String(v == null ? '' : v); },
    });
    denyGet(repl, ['Font', 'ParagraphFormat'], 'G-ENGINE');
    find.Replacement = repl;
    denyFn(find, ['HitHighlight', 'ClearHitHighlight'], 'G-ENGINE');
    denyGet(find, ['MatchPrefix', 'MatchSuffix', 'MatchPhrase', 'MatchByte', 'IgnoreSpace', 'IgnorePunct'], 'G-ENGINE');
    find.Execute = function (findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replace, matchKashida, matchDiacritics, matchAlefHamza, matchControl) {
      if (matchKashida || matchDiacritics || matchAlefHamza || matchControl) gap('G-ENGINE', 'Find');
      if (findText !== undefined) find.Text = String(findText);
      if (matchCase !== undefined) find.MatchCase = matchCase === true || matchCase === -1;
      if (matchWholeWord !== undefined) find.MatchWholeWord = matchWholeWord === true || matchWholeWord === -1;
      if (matchWildcards !== undefined) find.MatchWildcards = matchWildcards === true || matchWildcards === -1;
      if (matchSoundsLike !== undefined) find.MatchSoundsLike = matchSoundsLike === true || matchSoundsLike === -1;
      if (matchAllWordForms !== undefined) find.MatchAllWordForms = matchAllWordForms === true || matchAllWordForms === -1;
      if (forward !== undefined) find.Forward = !(forward === false || forward === 0);
      if (wrap !== undefined) find.Wrap = wrap;
      if (format !== undefined) find.Format = format === true || format === -1;
      if (replaceWith !== undefined) find._repl = String(replaceWith);
      if (find.Wrap === 2) gap('G-HOST', 'Find');
      const story = (() => { const e = endsOf(anchor); return e ? D.storyOf(D.doc, e[0].p) : D.doc.main; })();
      const up = !find.Forward;
      const origin = anchor && anchor.kind === 'selection'
        ? (up ? ensureCaret().a : ensureCaret().f)
        : ((() => { const e = endsOf(anchor); return e ? (up ? e[1] : e[0]) : ensureCaret().f; })());
      let matches = [];
      if (L.find && typeof L.find.all === 'function') {
        if (find.Text === '' && !find.Format) { find.Found = false; return false; }
        const o = { text: find.Text, matchCase: !!find.MatchCase, wholeWord: !!find.MatchWholeWord, wildcards: !!find.MatchWildcards, soundsLike: !!find.MatchSoundsLike, allForms: !!find.MatchAllWordForms, dir: up ? 'up' : 'down' };
        try { matches = (L.find.all(o, story).matches || []).map((m) => ({ a: m.a, b: m.b })); }
        catch (e) { if (e && e.message === 'empty') { find.Found = false; return false; } throw e; }
      } else {
        if (find.MatchWildcards || find.MatchSoundsLike || find.MatchAllWordForms) gap('G-ENGINE', 'Find');
        const st = storyTextOf(story);
        const from = offInStory(st, origin);
        for (const pair of literalMatches(st.text, find.Text, !!find.MatchCase, !!find.MatchWholeWord)) {
          matches.push({ a: posInStory(st, pair[0]), b: posInStory(st, pair[1]), s: pair[0], e: pair[1], from });
        }
      }
      const d = D.doc;
      const before = (m) => (m.e != null ? m.e <= offInStory(storyTextOf(story), origin) : D.cmp(d, m.b, origin) <= 0);
      const after = (m) => (m.s != null ? m.s >= offInStory(storyTextOf(story), origin) : D.cmp(d, m.a, origin) >= 0);
      let hit = null;
      if (up) {
        const prev = matches.filter(before);
        hit = prev.length ? prev[prev.length - 1] : (find.Wrap === 1 && matches.length ? matches[matches.length - 1] : null);
      } else {
        hit = matches.find(after) || (find.Wrap === 1 && matches.length ? matches[0] : null);
      }
      if (!hit) { find.Found = false; return false; }
      const replText = find._repl;
      const how = replace == null ? 0 : (replace | 0);
      if (how === 2) {
        const ordered = matches.slice().reverse();
        edit('Replace', () => {
          let pos = origin;
          for (const m of ordered) {
            pos = O.deleteRange(m.a, m.b);
            if (replText) pos = O.insertPlain(pos, replText);
          }
          return pos;
        });
      } else if (how === 1) {
        edit('Replace', () => {
          let pos = O.deleteRange(hit.a, hit.b);
          if (replText) pos = O.insertPlain(pos, replText);
          return pos;
        });
      } else {
        caret = { a: hit.a, f: hit.b };
        if (L.ed) {
          L.ed.sel = caret;
          if (typeof L.ed.setSel === 'function') L.ed.setSel(hit.a, hit.b, { noScroll: true });
        }
      }
      find.Found = true;
      return true;
    };
    return find;
  }

  function addTables(range, numRows, numColumns, defaultTableBehavior, autoFitBehavior) {
    if (defaultTableBehavior !== undefined && (defaultTableBehavior | 0) !== 0 && (defaultTableBehavior | 0) !== 1) gap('G-ENGINE', 'Tables.Add');
    const table = addTable(activeDocument(), range, numRows, numColumns);
    if (autoFitBehavior !== undefined) {
      const fit = autoFitBehavior | 0;
      if (fit !== 0 && fit !== 1 && fit !== 2) gap('G-ENGINE', 'Tables.Add');
      table.AutoFitBehavior(fit);
    }
    return table;
  }

  function addTable(doc, range, numRows, numColumns) {
    const nr = numRows | 0;
    const nc = numColumns | 0;
    if (nr < 1 || nc < 1) {
      const e = new Error('The number of rows and columns must be at least 1');
      e.code = 'missing';
      throw e;
    }
    const anchor = range && range._anchor;
    const ends = anchor ? endsOf(anchor) : null;
    const pos = ends ? ends[0] : ensureCaret().a;
    const sect = doc.model.sect || {};
    const width = Math.max(72, (sect.pgW || 612) - (sect.ml || 90) - (sect.mr || 90));
    let table = null;
    edit('Insert Table', () => {
      const host = pos.p;
      const num = host && host.pPr && host.pPr.num;
      const carry = num && num.id && num.id !== '0' ? num : null;
      table = D.simpleTable(nr, nc, width);
      if (carry) {
        for (const row of table.rows) for (const cell of row.cells) {
          for (const b of cell.blocks) if (b.t === 'p') b.pPr = Object.assign({}, b.pPr, { num: L.clone(carry) });
        }
      }
      O.insertBlockAt(pos, table);
      const p = D.firstPara(table.rows[0].cells[0].blocks);
      return D.pos(p, 0);
    });
    return makeTable(doc, table);
  }

  const Documents = function (index) { return Documents.Item(index); };
  stamp(Documents, Application);
  Object.defineProperty(Documents, 'Count', {
    enumerable: true,
    get() { activeDocument(); return wraps.length; },
  });
  Documents.Item = function (index) {
    activeDocument();
    const i = index | 0;
    if (i < 1 || i > wraps.length) {
      const e = new Error('The requested member of the collection does not exist.');
      e.code = 'missing';
      throw e;
    }
    return wraps[i - 1];
  };
  Object.defineProperty(Documents, 'First', { enumerable: true, get() { return Documents.Count ? Documents.Item(1) : null; } });
  Object.defineProperty(Documents, 'Last', { enumerable: true, get() { return Documents.Count ? Documents.Item(Documents.Count) : null; } });

  Documents.Add = function (template, newTemplate, documentType, visible) {
    if (template !== undefined || newTemplate !== undefined || documentType !== undefined || visible !== undefined) gap('G-ENGINE', 'Documents.Add');
    if (L.app && typeof L.app.newDocument === 'function') return L.app.newDocument();
    const model = D.newDoc();
    const wrap = adopt(model);
    D.doc = model;
    D.resetHistory();
    activeWrap = wrap;
    caret = null;
    clearPending();
    return wrap;
  };
  Documents.Open = function (fileName) {
    if (!L.hostAudit || !L.hostAudit.has('Open')) gap('G-HOST', 'Open');
    L.hostAudit.rejectExtra(arguments, 1, ['FileName', 'ConfirmConversions', 'ReadOnly', 'AddToRecentFiles', 'PasswordDocument']);
    return L.hostAudit.openPath(fileName);
  };
  Documents.Close = function () { gap('G-HOST', 'Documents.Close'); };
  Documents.Save = function () { gap('G-HOST', 'Documents.Save'); };

  function makeSelection() {
    const sel = stamp({}, Application);
    const selAnchor = { kind: 'selection' };
    attachStory(sel, () => selAnchor);
    storyFns(sel, ['GoTo', 'GoToNext', 'GoToPrevious']);
    sel.MoveLeft = function (unit, count, extend) { return moveSelection(-1, unit == null ? 1 : unit, count, extend); };
    sel.MoveRight = function (unit, count, extend) { return moveSelection(1, unit == null ? 1 : unit, count, extend); };
    sel.MoveUp = function (unit, count, extend) {
      const u = unit == null ? 5 : (unit | 0);
      if (layoutUnit(u)) gap('G-STORY', 'MoveUp');
      return moveSelection(-1, u, count, extend);
    };
    sel.MoveDown = function (unit, count, extend) {
      const u = unit == null ? 5 : (unit | 0);
      if (layoutUnit(u)) gap('G-STORY', 'MoveDown');
      return moveSelection(1, u, count, extend);
    };
    function keyEdge(unit, extend, end) {
      const u = unit == null ? 5 : (unit | 0);
      if (layoutUnit(u)) gap('G-STORY', end ? 'EndKey' : 'HomeKey');
      const p = storyPair(selAnchor);
      if (!p) return 0;
      const edge = unitEdge(p.st, end ? p.b : p.a, u, end);
      const a = end ? (extend ? p.a : edge) : edge;
      const b = end ? edge : (extend ? p.b : edge);
      placeSpan(posInStory(p.st, Math.min(a, b)), posInStory(p.st, Math.max(a, b)));
      return 1;
    }
    sel.HomeKey = function (unit, extend) { return keyEdge(unit, extend, false); };
    sel.EndKey = function (unit, extend) { return keyEdge(unit, extend, true); };
    sel.Characters = charactersOf(sel, () => endsOf({ kind: 'selection' }));
    Object.defineProperty(sel, 'Range', {
      enumerable: true,
      get() { return makeRange(sel, selAnchor); },
    });
    Object.defineProperty(sel, 'Text', {
      enumerable: true, configurable: true,
      get() { const e = endsOf(selAnchor); return e ? O.textRange(e[0], e[1], '\r') : ''; },
      set(v) { const e = endsOf(selAnchor); if (e) replaceSpan(e[0], e[1], v); },
    });
    Object.defineProperty(sel, 'Font', {
      enumerable: true, configurable: true,
      get() { return makeFont(sel, () => endsOf(selAnchor)); },
    });
    Object.defineProperty(sel, 'ParagraphFormat', {
      enumerable: true, configurable: true,
      get() { return makeSpanFormat(sel, () => endsOf(selAnchor)); },
    });
    Object.defineProperty(sel, 'Style', {
      enumerable: true, configurable: true,
      get() { return sel.Range.Style; },
      set(v) { sel.Range.Style = v; },
    });
    sel.TypeText = function (text) { typeText(text); };
    sel.TypeParagraph = function () { typeParagraph(); };
    sel.TypeBackspace = function () { typeBackspace(); };
    sel.Delete = function (unit, count) { deleteForward(unit, count); };
    sel.InsertBefore = function (text) { const e = endsOf(selAnchor); if (e) insertBeside(e[0], e[1], 'before', text); };
    sel.InsertAfter = function (text) { const e = endsOf(selAnchor); if (e) insertBeside(e[0], e[1], 'after', text); };
    sel.Collapse = function (dir) { collapseSel(dir); };
    sel.WholeStory = function () {
      const paras = mainParas(D.doc);
      if (!paras.length) return;
      placeSpan(D.pos(paras[0], 0), D.pos(paras[paras.length - 1], D.plen(paras[paras.length - 1])));
    };
    sel.Select = function () { sel.Range.Select(); };
    sel.Copy = function () { sel.Range.Copy(); };
    sel.Cut = function () { sel.Range.Cut(); };
    sel.Paste = function () {
      const e = endsOf({ kind: 'selection' });
      if (!e) return;
      const blocks = clipBlocks();
      if (!blocks || !blocks.length) return;
      edit('Paste', () => {
        const pos = D.eqPos(e[0], e[1]) ? e[0] : O.deleteRange(e[0], e[1]);
        return O.insertBlocks(pos, blocks);
      });
    };
    sel.InsertBreak = function (type) { insertBreak(type); };
    sel.InsertSymbol = function (characterNumber, font) { insertSymbol(characterNumber, font); };
    sel.InsertDateTime = function (fmt, asField, fullWidth, dateLanguage, calendarType) {
      if (fullWidth || calendarType) gap('G-ENGINE', 'InsertDateTime');
      insertDateTime(fmt, asField);
    };
    sel.InsertRows = function (n) { insertRowsAtCaret(n); };
    sel.InsertColumns = function (n) { insertColsAtCaret(n); };
    sel.InsertCells = function () { insertRowsAtCaret(1); };
    sel.ConvertToTable = function (separator, numRows, numColumns) {
      const rest = Array.prototype.slice.call(arguments, 3);
      if (rest.some((v) => v === true || v === -1)) gap('G-ENGINE', 'ConvertToTable');
      return convertSelectionToTable(separator, numRows, numColumns);
    };
    Object.defineProperty(sel, 'Type', {
      enumerable: true,
      get() {
        if (!D.doc) return 0;
        const e = endsOf({ kind: 'selection' });
        if (!e) return 0;
        return D.eqPos(e[0], e[1]) ? 1 : 2;
      },
    });
    sel.InsertFile = function (fileName) {
      if (!L.hostAudit || !L.hostAudit.has('InsertFile')) gap('G-HOST', 'InsertFile');
      L.hostAudit.rejectExtra(arguments, 1, ['FileName', 'Range', 'ConfirmConversions', 'Link', 'Attachment']);
      return L.hostAudit.readBytes(fileName, 'InsertFile', 'AlwaysAllow').then((bytes) => {
        return L.app.insertFileBytes(String(fileName).split(/[/\\]/).pop(), bytes);
      });
    };
    denyFn(sel, ['InsertCaption', 'PasteAndFormat', 'PasteSpecial', 'PasteExcelTable'], 'G-ENGINE');
    for (const name of ['Paragraphs', 'Sentences', 'Words', 'Tables', 'ListFormat', 'Hyperlinks', 'Bookmarks', 'Fields', 'Comments', 'InlineShapes', 'FormattedText']) {
      Object.defineProperty(sel, name, {
        enumerable: true, configurable: true,
        get() { return sel.Range[name]; },
      });
    }
    for (const name of ['Bold', 'Italic', 'Underline', 'Case', 'LanguageID', 'LanguageIDFarEast']) {
      Object.defineProperty(sel, name, {
        enumerable: true, configurable: true,
        get() { return sel.Range[name]; },
        set(v) { sel.Range[name] = v; },
      });
    }
    for (const name of ['InsertParagraph', 'InsertParagraphBefore', 'InsertParagraphAfter']) {
      sel[name] = function () { return sel.Range[name].apply(sel.Range, arguments); };
    }
    Object.defineProperty(sel, 'Document', {
      enumerable: true, configurable: true,
      get() { return activeDocument(); },
    });
    sel._find = makeFind(sel, selAnchor);
    Object.defineProperty(sel, 'Find', { enumerable: true, configurable: true, get() { return sel._find; } });
    Object.defineProperty(sel, 'Borders', { enumerable: true, get() { return sel.ParagraphFormat.Borders; } });
    Object.defineProperty(sel, 'Shading', { enumerable: true, get() { return sel.ParagraphFormat.Shading; } });
    Object.defineProperty(sel, 'Shapes', {
      enumerable: true,
      get() { return rangeShapes(sel, { kind: 'selection' }, true); },
    });
    Object.defineProperty(sel, 'ShapeRange', { enumerable: true, get() { return sel.Shapes; } });
    Object.defineProperty(sel, 'Cells', {
      enumerable: true,
      get() {
        const hit = caretCell();
        if (!hit) return cellCollection(activeDocument(), { rows: [], grid: [] }, []);
        return cellCollection(activeDocument(), hit.table, [hit.cell]);
      },
    });
    function emptyRows(parent) {
      const rows = function () { missing(); };
      stamp(rows, parent);
      Object.defineProperty(rows, 'Count', { enumerable: true, get() { return 0; } });
      rows.Item = function () { missing(); };
      rows.Delete = function () {};
      return rows;
    }
    Object.defineProperty(sel, 'Rows', {
      enumerable: true,
      get() {
        const hit = caretCell();
        if (!hit) return emptyRows(sel);
        return makeTable(activeDocument(), hit.table).Rows;
      },
    });
    Object.defineProperty(sel, 'Columns', {
      enumerable: true,
      get() {
        const hit = caretCell();
        if (!hit) return emptyRows(sel);
        return makeTable(activeDocument(), hit.table).Columns;
      },
    });
    Object.defineProperty(sel, 'Sections', { enumerable: true, get() { return activeDocument().Sections; } });
    Object.defineProperty(sel, 'PageSetup', { enumerable: true, get() { const s = activeDocument().Sections; return s.Count ? s.Item(1).PageSetup : null; } });
    Object.defineProperty(sel, 'StoryType', {
      enumerable: true,
      get() {
        const e = endsOf(selAnchor);
        if (!e || !D.doc) return 1;
        const story = D.storyOf(D.doc, e[0].p);
        return story === D.doc.main ? 1 : 5;
      },
    });
    Object.defineProperty(sel, 'StoryLength', {
      enumerable: true,
      get() {
        const e = endsOf(selAnchor);
        const story = e && D.doc ? D.storyOf(D.doc, e[0].p) : (D.doc && D.doc.main);
        return story ? storyTextOf(story).text.length : 0;
      },
    });
    sel.InStory = function (range) { return range ? -1 : 0; };
    sel.InRange = function (range) {
      if (!range || !range.Start && range.Start !== 0) return 0;
      const a = sel.Start, b = sel.End;
      return a >= range.Start && b <= range.End ? -1 : 0;
    };
    Object.defineProperty(sel, 'StartIsActive', { enumerable: true, get() { return 0; } });
    Object.defineProperty(sel, 'BookmarkID', { enumerable: true, get() { return 0; } });
    Object.defineProperty(sel, 'NoProofing', {
      enumerable: true, configurable: true,
      get() {
        const e = endsOf(selAnchor);
        if (!e) return 0;
        const vals = runValues(e[0], e[1], 'noProof');
        return triRead(vals);
      },
      set(v) {
        const t = triWrite(v);
        if (t === 'skip') return;
        const e = endsOf(selAnchor);
        if (e) applyRun(e[0], e[1], { noProof: !!t });
      },
    });
    sel.SelectCell = function () { const hit = caretCell(); if (hit) makeCell(activeDocument(), hit.table, hit.cell).Range.Select(); };
    sel.SelectRow = function () {
      const hit = caretCell();
      if (!hit) return;
      const row = makeTable(activeDocument(), hit.table).Rows.Item(hit.rowIndex);
      if (row) row.Select();
    };
    sel.SelectColumn = function () {
      const hit = caretCell();
      if (!hit) return;
      const col = hit.row.cells.indexOf(hit.cell) + 1;
      makeTable(activeDocument(), hit.table).Columns.Item(col).Select();
    };
    Object.defineProperty(sel, '_pictureAt', { value: pictureAtOf({ kind: 'selection' }) });
    sel.InsertRowsAbove = function (n) { insertRowsAtCaret(n); };
    sel.InsertRowsBelow = function (n) { insertRowsAtCaret(n); };
    sel.InsertColumnsRight = function (n) { insertColsAtCaret(n); };
    const storyGap = ['MoveWhile', 'MoveUntil', 'MoveStartWhile', 'MoveEndWhile', 'MoveStartUntil', 'MoveEndUntil', 'IPAtEndOfLine', 'IsEndOfRowMark'];
    storyGets(sel, storyGap.filter((n) => n === 'IPAtEndOfLine' || n === 'IsEndOfRowMark'));
    denyGet(sel, ['ChildShapeRange', 'HasChildShapeRange', 'Editors', 'FormFields', 'Frames', 'HTMLDivisions', 'WordOpenXML', 'XML', 'EndnoteOptions', 'FootnoteOptions', 'FitTextWidth', 'Flags', 'ExtendMode', 'ColumnSelectMode', 'EnhMetaFileBits', 'LanguageIDOther', 'HeaderFooter', 'TopLevelTables'], 'G-ENGINE');
    denyFn(sel, ['InsertFormula', 'InsertCrossReference', 'InsertXML', 'InsertStyleSeparator', 'PasteAsNestedTable', 'PasteAppendTable', 'PasteFormat', 'CopyAsPicture', 'ToggleCharacterCode', 'DetectLanguage', 'Calculate', 'Shrink', 'EscapeKey', 'CreateAutoTextEntry', 'NextRevision', 'BoldRun', 'ItalicRun', 'ClearFormatting', 'CreateTextbox', 'NextField', 'PreviousField', 'InsertNewPage', 'SortAscending', 'SortDescending', 'SplitTable'], 'G-ENGINE');
    return sel;
  }
  const Selection = makeSelection();

  stamp(Application, null);
  Application.Parent = Application;
  Application.Application = Application;

  Object.defineProperty(Application, 'ActiveDocument', { enumerable: true, configurable: true, get: activeDocument });
  Object.defineProperty(Application, 'Documents', { enumerable: true, configurable: true, get() { return Documents; } });
  Object.defineProperty(Application, 'Selection', { enumerable: true, configurable: true, get() { return Selection; } });
  Object.defineProperty(Application, 'Dialogs', {
    enumerable: true,
    get() { return L.hostAudit ? L.hostAudit.dialogsCollection() : gap('G-HOST', 'Dialogs'); },
  });
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

  dom.Application = Application;
  dom.gap = gap;

  const CONSTS = {
    wdAlignParagraphLeft: 0, wdAlignParagraphCenter: 1, wdAlignParagraphRight: 2, wdAlignParagraphJustify: 3, wdAlignParagraphDistribute: 4,
    wdCollapseEnd: 0, wdCollapseStart: 1,
    wdDoNotSaveChanges: 0, wdSaveChanges: -1, wdPromptToSaveChanges: -2,
    wdFindStop: 0, wdFindContinue: 1, wdFindAsk: 2, wdReplaceNone: 0, wdReplaceOne: 1, wdReplaceAll: 2,
    wdUnderlineNone: 0, wdUnderlineSingle: 1, wdUnderlineWords: 2, wdUnderlineDouble: 3, wdUnderlineDotted: 4, wdUnderlineThick: 6, wdUnderlineDash: 7, wdUnderlineDotDash: 9, wdUnderlineDotDotDash: 10, wdUnderlineWavy: 11,
    wdLineSpaceSingle: 0, wdLineSpace1pt5: 1, wdLineSpaceDouble: 2, wdLineSpaceAtLeast: 3, wdLineSpaceExactly: 4, wdLineSpaceMultiple: 5,
    wdMainTextStory: 1, wdFootnotesStory: 2, wdEndnotesStory: 3, wdCommentsStory: 4, wdTextFrameStory: 5, wdEvenPagesHeaderStory: 6, wdPrimaryHeaderStory: 7, wdEvenPagesFooterStory: 8, wdPrimaryFooterStory: 9, wdFirstPageHeaderStory: 10, wdFirstPageFooterStory: 11,
    wdHeaderFooterPrimary: 1, wdHeaderFooterFirstPage: 2, wdHeaderFooterEvenPages: 3,
    wdOrientPortrait: 0, wdOrientLandscape: 1,
    wdStatisticWords: 0, wdStatisticLines: 1, wdStatisticPages: 2, wdStatisticCharacters: 3, wdStatisticParagraphs: 4, wdStatisticCharactersWithSpaces: 5,
    wdNoSelection: 0, wdSelectionIP: 1, wdSelectionNormal: 2,
    wdColorAutomatic: -16777216,
    wdStyleTypeParagraph: 1, wdStyleTypeCharacter: 2, wdStyleTypeTable: 3, wdStyleTypeList: 4,
    wdReadingOrderLtr: 0, wdReadingOrderRtl: 1,
    wdLowerCase: 0, wdUpperCase: 1, wdTitleWord: 2, wdTitleSentence: 4, wdToggleCase: 5,
    wdLineBreak: 6, wdPageBreak: 7, wdColumnBreak: 8, wdSectionBreakNextPage: 2, wdSectionBreakContinuous: 3,
    wdAutoFitFixed: 0, wdAutoFitContent: 1, wdAutoFitWindow: 2,
    wdSeparateByParagraphs: 0, wdSeparateByTabs: 1, wdSeparateByCommas: 2,
    wdSortOrderAscending: 0, wdSortOrderDescending: 1,
    wdBorderTop: -1, wdBorderLeft: -2, wdBorderBottom: -3, wdBorderRight: -4, wdBorderHorizontal: -5, wdBorderVertical: -6, wdBorderDiagonalDown: -7, wdBorderDiagonalUp: -8,
    wdLineStyleNone: 0, wdLineStyleSingle: 1, wdLineStyleDot: 2, wdLineStyleDashSmallGap: 3, wdLineStyleDashLargeGap: 4, wdLineStyleDashDot: 5, wdLineStyleDashDotDot: 6, wdLineStyleDouble: 7, wdLineStyleTriple: 8, wdLineStyleThinThickSmallGap: 9, wdLineStyleThickThinSmallGap: 10, wdLineStyleThinThickThinSmallGap: 11, wdLineStyleWave: 18, wdLineStyleDoubleWave: 19, wdLineStyleDashDotStroked: 20, wdLineStyleEmboss3D: 21, wdLineStyleEngrave3D: 22, wdLineStyleOutset: 23, wdLineStyleInset: 24,
    wdLineWidth025pt: 2, wdLineWidth050pt: 4, wdLineWidth075pt: 6, wdLineWidth100pt: 8, wdLineWidth150pt: 12, wdLineWidth225pt: 18, wdLineWidth300pt: 24, wdLineWidth450pt: 36, wdLineWidth600pt: 48,
    wdTextureNone: 0, wdTextureSolid: 1000, wdTexture5Percent: 50, wdTexture10Percent: 100, wdTexture25Percent: 250, wdTexture50Percent: 500, wdTexture75Percent: 750,
    wdCharacter: 1, wdWord: 2, wdSentence: 3, wdParagraph: 4, wdLine: 5, wdStory: 6, wdScreen: 7, wdWindow: 11,
    wdPaperLetter: 2, wdPaperLegal: 4, wdPaperExecutive: 5, wdPaperA3: 6, wdPaperA4: 7, wdPaperA5: 9, wdPaperB5: 11, wdPaperTabloid: 23, wdPaperStatement: 22, wdPaperEnvelope10: 25, wdPaperEnvelopeDL: 37, wdPaperEnvelopeC5: 34, wdPaperEnvelopeMonarch: 39, wdPaperCustom: 41,
    wdRowHeightAuto: 0, wdRowHeightAtLeast: 1, wdRowHeightExactly: 2,
    wdDropNone: 0, wdDropNormal: 1, wdDropMargin: 2,
    wdListNoNumbering: 0, wdListBullet: 2, wdListSimpleNumbering: 3, wdListMixedNumbering: 5,
    wdPreferredWidthAuto: 1, wdPreferredWidthPercent: 2, wdPreferredWidthPoints: 3,
    wdAlignVerticalTop: 0, wdAlignVerticalCenter: 1, wdAlignVerticalJustify: 2, wdAlignVerticalBottom: 3,
    wdWrapSquare: 0, wdWrapTight: 1, wdWrapThrough: 2, wdWrapNone: 3, wdWrapTopBottom: 4, wdWrapInline: 7,
    wdInlineShapePicture: 3, wdInlineShapeChart: 12, wdInlineShapeSmartArt: 13,
    wdPropertyTypeString: 4, wdEnglishUS: 1033,
    wdDialogFileOpen: 80, wdDialogFileSaveAs: 84, wdDialogFilePrint: 88, wdDialogFileSummaryInfo: 86,
    wdDialogFormatFont: 174, wdDialogFormatParagraph: 175, wdDialogFormatTabs: 179, wdDialogFormatBordersAndShading: 189,
    wdDialogFormatBulletsAndNumbering: 824, wdDialogFormatColumns: 177, wdDialogFormatDropCap: 488,
    wdDialogFormatChangeCase: 322, wdDialogFormatStyle: 180, wdDialogFormatDrawingObject: 960,
    wdDialogFilePageSetup: 178, wdDialogFileDocumentLayout: 178, wdDialogInsertPageNumbers: 294,
    wdDialogInsertBreak: 159, wdDialogInsertDateTime: 165, wdDialogInsertField: 166, wdDialogInsertSymbol: 162,
    wdDialogInsertFootnote: 370, wdDialogInsertCaption: 357, wdDialogInsertCrossReference: 367,
    wdDialogInsertBookmark: 168, wdDialogInsertHyperlink: 925, wdDialogInsertObject: 172,
    wdDialogTableInsertTable: 129, wdDialogTableInsertCells: 130, wdDialogTableDeleteCells: 133,
    wdDialogTableSplitCells: 137, wdDialogTableAutoFormat: 563, wdDialogTableSort: 199, wdDialogTableFormula: 348,
    wdDialogTextToTable: 127, wdDialogTableToText: 128, wdDialogToolsWordCount: 228, wdDialogToolsHyphenation: 195,
    wdDialogToolsEnvelopesAndLabels: 607, wdDialogLetterWizard: 821, wdDialogToolsTemplates: 87,
    wdDialogOrganizer: 222, wdDialogToolsAutoCorrect: 378, wdDialogToolsOptions: 974, wdDialogToolsLanguage: 188,
    wdDialogEditPasteSpecial: 111, wdExportFormatPDF: 17, wdExportFormatXPS: 18,
    msoFileDialogOpen: 1, msoFileDialogSaveAs: 2, msoFileDialogFilePicker: 3, msoFileDialogFolderPicker: 4,
  };

  function hexToBgr(hex) {
    const h = String(hex).replace('#', '');
    return parseInt(h.slice(0, 2), 16) | (parseInt(h.slice(2, 4), 16) << 8) | (parseInt(h.slice(4, 6), 16) << 16);
  }
  const PALETTE = [
    ['Black', '#000000'], ['Brown', '#993300'], ['OliveGreen', '#333300'], ['DarkGreen', '#003300'], ['DarkTeal', '#003366'], ['DarkBlue', '#000080'], ['Indigo', '#333399'], ['Gray80', '#333333'],
    ['DarkRed', '#800000'], ['Orange', '#FF6600'], ['DarkYellow', '#808000'], ['Green', '#008000'], ['Teal', '#008080'], ['Blue', '#0000FF'], ['BlueGray', '#666699'], ['Gray50', '#808080'],
    ['Red', '#FF0000'], ['LightOrange', '#FF9900'], ['Lime', '#99CC00'], ['SeaGreen', '#339966'], ['Aqua', '#33CCCC'], ['LightBlue', '#3366FF'], ['Violet', '#800080'], ['Gray40', '#969696'],
    ['Pink', '#FF00FF'], ['Gold', '#FFCC00'], ['Yellow', '#FFFF00'], ['BrightGreen', '#00FF00'], ['Turquoise', '#00FFFF'], ['SkyBlue', '#00CCFF'], ['Plum', '#993366'], ['Gray25', '#C0C0C0'],
    ['Rose', '#FF99CC'], ['Tan', '#FFCC99'], ['LightYellow', '#FFFF99'], ['LightGreen', '#CCFFCC'], ['LightTurquoise', '#CCFFFF'], ['PaleBlue', '#99CCFF'], ['Lavender', '#CC99FF'], ['White', '#FFFFFF'],
  ];
  for (const pair of PALETTE) CONSTS['wdColor' + pair[0]] = hexToBgr(pair[1]);

  for (const k in CONSTS) {
    dom[k] = CONSTS[k];
    try {
      const d = Object.getOwnPropertyDescriptor(window, k);
      if (!d || d.configurable) Object.defineProperty(window, k, { configurable: true, enumerable: true, writable: true, value: CONSTS[k] });
    } catch (e) { /* leave a non-configurable global alone */ }
  }

  defineGetter(window, 'Application', () => Application);
  defineGetter(window, 'ActiveDocument', () => Application.ActiveDocument);
  defineGetter(window, 'Documents', () => Documents);
  defineGetter(window, 'Selection', () => Selection);
})();
