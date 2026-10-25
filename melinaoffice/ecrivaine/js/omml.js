/* Ecrivaine — Office Math (OMML) to MathML, so equations from Word render as real two-dimensional math.
 * Covers the structures Word's equation tools produce: fractions, scripts, radicals, delimiters,
 * n-ary operators, functions, accents, bars, group characters, limits, matrices, equation arrays,
 * boxes, phantoms and pre-scripts. Rendering relies on the browser's MathML Core support. */
(function () {
  'use strict';
  const L = window.L;
  const MNS = 'http://www.w3.org/1998/Math/MathML';
  const M = (L.omml = {});

  const lname = (n) => n.localName || n.nodeName.replace(/^.*:/, '');
  const kids = (n, name) => Array.from(n.childNodes).filter((c) => c.nodeType === 1 && (!name || lname(c) === name));
  const kid = (n, name) => (n ? kids(n, name)[0] || null : null);
  const val = (n) => {
    if (!n) return null;
    for (const a of Array.from(n.attributes || [])) if (a.localName === 'val') return a.value;
    return null;
  };
  const prop = (n, pr, name) => val(kid(kid(n, pr), name));
  const onOff = (v) => v != null && v !== '0' && v !== 'false' && v !== 'off';

  function mk(tag, text, attrs) {
    const e = document.createElementNS(MNS, tag);
    if (text != null) e.textContent = text;
    if (attrs) for (const k in attrs) if (attrs[k] != null) e.setAttribute(k, attrs[k]);
    return e;
  }
  function row(nodes) {
    const list = nodes.filter(Boolean);
    if (list.length === 1) return list[0];
    const r = mk('mrow');
    for (const n of list) r.appendChild(n);
    return r;
  }

  /* ---------- styled alphabets (m:scr) ---------- */
  const HOLES = {
    'double-struck': { C: 0x2102, H: 0x210D, N: 0x2115, P: 0x2119, Q: 0x211A, R: 0x211D, Z: 0x2124 },
    script: { B: 0x212C, E: 0x2130, F: 0x2131, H: 0x210B, I: 0x2110, L: 0x2112, M: 0x2133, R: 0x211B, e: 0x212F, g: 0x210A, o: 0x2134 },
    fraktur: { C: 0x212D, H: 0x210C, I: 0x2111, R: 0x211C, Z: 0x2128 },
  };
  const BASE = { 'double-struck': 0x1D538, script: 0x1D49C, fraktur: 0x1D504, 'sans-serif': 0x1D5A0, monospace: 0x1D670 };
  function styled(ch, scr) {
    const b = BASE[scr];
    if (!b) return ch;
    const h = HOLES[scr] && HOLES[scr][ch];
    if (h) return String.fromCodePoint(h);
    const c = ch.charCodeAt(0);
    if (c >= 65 && c <= 90) return String.fromCodePoint(b + c - 65);
    if (c >= 97 && c <= 122) return String.fromCodePoint(b + 26 + c - 97);
    if (c >= 48 && c <= 57 && (scr === 'double-struck' || scr === 'sans-serif' || scr === 'monospace')) {
      return String.fromCodePoint({ 'double-struck': 0x1D7D8, 'sans-serif': 0x1D7E2, monospace: 0x1D7F6 }[scr] + c - 48);
    }
    return ch;
  }

  /* ---------- runs ---------- */
  const LETTER = /\p{L}|\p{M}/u;
  const BRACKETS = /[()[\]{}⟨⟩|‖⌈⌉⌊⌋〈〉⟦⟧]/;
  function runNodes(r) {
    const rp = kid(r, 'rPr');
    const sty = prop(r, 'rPr', 'sty');
    const nor = rp && kid(rp, 'nor') && onOff(val(kid(rp, 'nor')) == null ? '1' : val(kid(rp, 'nor')));
    const scr = prop(r, 'rPr', 'scr');
    let text = '';
    for (const c of kids(r)) {
      const n = lname(c);
      if (n === 't') text += c.textContent;
      else if (n === 'br') text += ' ';
    }
    if (!text) return [];
    const bold = sty === 'b' || sty === 'bi';
    const upright = sty === 'p' || sty === 'b';
    if (nor) {
      const t = mk('mtext', text);
      if (bold) t.setAttribute('style', 'font-weight:bold');
      return [t];
    }
    const out = [];
    const chars = Array.from(text);
    let i = 0;
    while (i < chars.length) {
      const ch = chars[i];
      if (/\s/.test(ch)) { out.push(mk('mspace', null, { width: '0.25em' })); i++; continue; }
      if (/[0-9]/.test(ch)) {
        let s = ch; i++;
        while (i < chars.length && (/[0-9]/.test(chars[i]) || (chars[i] === '.' && /[0-9]/.test(chars[i + 1] || '')))) s += chars[i++];
        out.push(mk('mn', scr ? Array.from(s).map((c) => styled(c, scr)).join('') : s));
        continue;
      }
      if (LETTER.test(ch) || ch === '∞' || ch === '∂' || ch === '∅' || ch === 'ℏ') {
        let s = ch; i++;
        while (i < chars.length && (LETTER.test(chars[i]))) s += chars[i++];
        if (scr) s = Array.from(s).map((c) => styled(c, scr)).join('');
        if (upright || scr) out.push(mk('mi', s, { mathvariant: 'normal' }));
        else for (const c of Array.from(s)) out.push(mk('mi', c));
        continue;
      }
      const mo = mk('mo', ch === '-' && !upright ? '\u2212' : ch === '*' ? '\u2217' : ch);
      /* brackets typed as plain characters take no operator spacing (their form is not inferable here) */
      if (BRACKETS.test(ch)) { mo.setAttribute('lspace', '0'); mo.setAttribute('rspace', '0'); mo.setAttribute('stretchy', 'false'); }
      else if (ch === ',' || ch === ';') { mo.setAttribute('lspace', '0'); mo.setAttribute('rspace', '0.1667em'); }
      out.push(mo);
      i++;
    }
    if (bold) for (const n of out) n.setAttribute('style', 'font-weight:bold');
    return out;
  }

  /* ---------- structures ---------- */
  const ACCENT = { '̀': '`', '́': '´', '̂': 'ˆ', '̃': '˜', '̄': '¯', '̅': '¯', '̆': '˘', '̇': '˙', '̈': '¨', '̌': 'ˇ', '⃗': '→', '⃖': '←', '⃡': '↔', '⃐': '↼', '⃑': '⇀', '̑': '⌢', '̿': '‗' };
  const INTEGRALS = /[∫∬∭∮∯∰∱∲∳⨌]/;

  /** convert the children of an argument element (m:e, m:num, …) */
  function arg(n) { return n ? row(seq(n)) : mk('mrow'); }
  function seq(n) {
    const out = [];
    for (const c of kids(n)) {
      const x = conv(c);
      if (Array.isArray(x)) out.push(...x); else if (x) out.push(x);
    }
    return out;
  }
  function conv(c) {
    switch (lname(c)) {
      case 'r': return runNodes(c);
      case 'f': {
        const type = prop(c, 'fPr', 'type') || 'bar';
        const num = arg(kid(c, 'num')), den = arg(kid(c, 'den'));
        if (type === 'lin' || type === 'skw') return row([num, mk('mo', '/'), den]);
        const f = mk('mfrac');
        if (type === 'noBar') f.setAttribute('linethickness', '0');
        f.append(num, den);
        return f;
      }
      case 'sSup': { const e = mk('msup'); e.append(arg(kid(c, 'e')), arg(kid(c, 'sup'))); return e; }
      case 'sSub': { const e = mk('msub'); e.append(arg(kid(c, 'e')), arg(kid(c, 'sub'))); return e; }
      case 'sSubSup': { const e = mk('msubsup'); e.append(arg(kid(c, 'e')), arg(kid(c, 'sub')), arg(kid(c, 'sup'))); return e; }
      case 'sPre': {
        const e = mk('mmultiscripts');
        e.append(arg(kid(c, 'e')), mk('mprescripts'), arg(kid(c, 'sub')), arg(kid(c, 'sup')));
        return e;
      }
      case 'rad': {
        const hide = onOff(prop(c, 'radPr', 'degHide'));
        const deg = kid(c, 'deg');
        const e = arg(kid(c, 'e'));
        if (hide || !deg || !deg.textContent.trim()) { const s = mk('msqrt'); s.appendChild(e); return s; }
        const r = mk('mroot'); r.append(e, arg(deg)); return r;
      }
      case 'd': {
        const pr = kid(c, 'dPr');
        const g = (name, dflt) => { const v = val(kid(pr, name)); return v == null ? dflt : v; };
        const beg = g('begChr', '('), end = g('endChr', ')'), sep = g('sepChr', '|');
        const parts = [];
        const grow = val(kid(pr, 'grow'));
        const args = kids(c, 'e').map(arg);
        /* grow only around tall content; a stretched variant of a plain bracket is too wide */
        const tall = args.some((a) => a.querySelector && (/^(mfrac|munderover|munder|mover|mtable)$/.test(a.localName) || a.querySelector('mfrac, munderover, munder, mtable')));
        const st = (grow != null && !onOff(grow)) || !tall ? 'false' : 'true';
        if (beg) parts.push(mk('mo', beg, { fence: 'true', form: 'prefix', stretchy: st, lspace: '0', rspace: '0', symmetric: 'true' }));
        args.forEach((a, i) => { if (i && sep) parts.push(mk('mo', sep, { separator: 'true', stretchy: st, lspace: '0', rspace: '0.1667em' })); parts.push(a); });
        if (end) parts.push(mk('mo', end, { fence: 'true', form: 'postfix', stretchy: st, lspace: '0', rspace: '0', symmetric: 'true' }));
        const r = mk('mrow'); for (const p of parts) r.appendChild(p); return r;
      }
      case 'nary': {
        const pr = kid(c, 'naryPr');
        const chr = val(kid(pr, 'chr')) || '∫';
        const subHide = onOff(val(kid(pr, 'subHide'))), supHide = onOff(val(kid(pr, 'supHide')));
        const loc = val(kid(pr, 'limLoc')) || (INTEGRALS.test(chr) ? 'subSup' : 'undOvr');
        const op = mk('mo', chr, { largeop: 'true', movablelimits: 'false' });
        const sub = subHide ? null : kid(c, 'sub'), sup = supHide ? null : kid(c, 'sup');
        const filled = (n) => !!n && kids(n).some((k) => lname(k) !== 'ctrlPr' && lname(k) !== 'argPr');
        const hasSub = filled(sub), hasSup = filled(sup);
        let base = op;
        const under = loc === 'undOvr';
        if (hasSub && hasSup) { base = mk(under ? 'munderover' : 'msubsup'); base.append(op, arg(sub), arg(sup)); }
        else if (hasSub) { base = mk(under ? 'munder' : 'msub'); base.append(op, arg(sub)); }
        else if (hasSup) { base = mk(under ? 'mover' : 'msup'); base.append(op, arg(sup)); }
        return row([base, arg(kid(c, 'e'))]);
      }
      case 'func': {
        const name = kid(c, 'fName');
        const fn = name ? row(seq(name).map((n) => { if (n.localName === 'mi' && n.textContent.length === 1) n.setAttribute('mathvariant', 'normal'); return n; })) : null;
        /* function names are upright: merge consecutive single letters into one identifier */
        if (fn && fn.localName === 'mrow') {
          const txt = Array.from(fn.childNodes).every((n) => n.localName === 'mi') ? fn.textContent : null;
          if (txt) return row([mk('mi', txt), mk('mo', '⁡'), arg(kid(c, 'e'))]);
        }
        return row([fn, mk('mo', '⁡'), arg(kid(c, 'e'))]);
      }
      case 'acc': {
        const ch = prop(c, 'accPr', 'chr') || '̂';
        const e = mk('mover', null, { accent: 'true' });
        e.append(arg(kid(c, 'e')), mk('mo', ACCENT[ch] || ch, { stretchy: /[→←↔¯]/.test(ACCENT[ch] || ch) ? 'true' : 'false' }));
        return e;
      }
      case 'bar': {
        const top = prop(c, 'barPr', 'pos') === 'top';
        const e = mk(top ? 'mover' : 'munder', null, top ? { accent: 'true' } : { accentunder: 'true' });
        e.append(arg(kid(c, 'e')), mk('mo', top ? '¯' : '_', { stretchy: 'true' }));
        return e;
      }
      case 'groupChr': {
        const pr = kid(c, 'groupChrPr');
        const chr = val(kid(pr, 'chr')) || '⏟';
        const top = val(kid(pr, 'pos')) === 'top';
        const e = mk(top ? 'mover' : 'munder');
        e.append(arg(kid(c, 'e')), mk('mo', chr, { stretchy: 'true' }));
        return e;
      }
      case 'limLow': { const e = mk('munder'); e.append(arg(kid(c, 'e')), arg(kid(c, 'lim'))); return e; }
      case 'limUpp': { const e = mk('mover'); e.append(arg(kid(c, 'e')), arg(kid(c, 'lim'))); return e; }
      case 'm': {
        const t = mk('mtable');
        for (const mr of kids(c, 'mr')) {
          const tr = mk('mtr');
          for (const e of kids(mr, 'e')) { const td = mk('mtd'); td.appendChild(arg(e)); tr.appendChild(td); }
          t.appendChild(tr);
        }
        return t;
      }
      case 'eqArr': {
        const t = mk('mtable', null, { columnalign: 'left' });
        for (const e of kids(c, 'e')) { const tr = mk('mtr'); const td = mk('mtd'); td.appendChild(arg(e)); tr.appendChild(td); t.appendChild(tr); }
        return t;
      }
      case 'box': return arg(kid(c, 'e'));
      case 'borderBox': { const r = mk('mrow', null, { style: 'border:1px solid currentColor;padding:1px 2px' }); r.appendChild(arg(kid(c, 'e'))); return r; }
      case 'phant': {
        const show = prop(c, 'phantPr', 'show');
        if (show != null && !onOff(show)) { const p = mk('mphantom'); p.appendChild(arg(kid(c, 'e'))); return p; }
        return arg(kid(c, 'e'));
      }
      case 'oMath': return seq(c);
      /* wrappers that carry runs: revisions, content controls, smart tags */
      case 'ins': case 'sdt': case 'sdtContent': case 'customXml': case 'smartTag': case 'moveTo': return seq(c);
      case 'del': case 'moveFrom': return null;
      default: return null;
    }
  }

  /** first explicit font size (pt) in the equation, from w:rPr/w:sz inside its runs */
  M.fontSize = function (node) {
    const all = node.getElementsByTagNameNS ? node.getElementsByTagNameNS('*', 'sz') : [];
    for (const s of Array.from(all)) {
      const p = s.parentNode;
      if (p && lname(p) === 'rPr' && p.parentNode && lname(p.parentNode) === 'r') { const v = +val(s); if (v) return v / 2; }
    }
    return null;
  };

  /**
   * Build a <math> element from an m:oMath / m:oMathPara element or its serialized XML.
   * Returns {math, display: bool, jc} or null when the input cannot be parsed.
   */
  M.toMathML = function (src) {
    let node = src;
    if (typeof src === 'string') {
      const xml = /xmlns:m=/.test(src) ? src : src.replace(/^<m:(oMath\w*)/, '<m:$1 xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"');
      const doc = new DOMParser().parseFromString(xml, 'application/xml');
      if (doc.getElementsByTagName('parsererror').length) return null;
      node = doc.documentElement;
    }
    if (!node) return null;
    const isPara = lname(node) === 'oMathPara';
    const math = mk('math', null, isPara ? { display: 'block' } : null);
    if (isPara) {
      const maths = kids(node, 'oMath');
      if (maths.length > 1) {
        /* several equations in one paragraph: one per line */
        const t = mk('mtable');
        for (const m of maths) { const tr = mk('mtr'); const td = mk('mtd'); td.appendChild(row(seq(m))); tr.appendChild(td); t.appendChild(tr); }
        math.appendChild(t);
      } else if (maths[0]) for (const n of seq(maths[0])) math.appendChild(n);
    } else for (const n of seq(node)) math.appendChild(n);
    const jc = isPara ? prop(node, 'oMathParaPr', 'jc') || 'centerGroup' : null;
    return { math, display: isPara, jc };
  };

  /* Office's LaTeX export dialect. The next open derives it again, so braces may differ from what was typed. */
  const FUN = new Set(['sin', 'cos', 'tan', 'cot', 'sec', 'csc', 'log', 'ln', 'lim', 'max', 'min', 'sup', 'inf', 'det', 'gcd', 'exp', 'arcsin', 'arccos', 'arctan', 'sinh', 'cosh', 'tanh']);
  const NARY_CMD = { '∑': '\\sum', '∏': '\\prod', '∫': '\\int', '∮': '\\oint', '⋃': '\\bigcup', '⋂': '\\bigcap' };
  const ACC_CMD = { '\u0302': '\\hat', '\u02C6': '\\hat', '\u0303': '\\tilde', '\u02DC': '\\tilde', '\u0304': '\\bar', '\u00AF': '\\bar', '\u0305': '\\overline', '\u0307': '\\dot', '\u20D7': '\\vec', '\u2192': '\\vec' };
  const CHAR_CMD = {
    '\u2026': '\\ldots', '\u22EF': '\\cdots', '\u22EE': '\\vdots', '\u22F1': '\\ddots', '\u221E': '\\infty', '\u2212': '-',
    '\u00B1': '\\pm', '\u2213': '\\mp', '\u00D7': '\\times', '\u00F7': '\\div', '\u22C5': '\\cdot',
    '\u2264': '\\leq', '\u2265': '\\geq', '\u2260': '\\neq', '\u2248': '\\approx', '\u2261': '\\equiv',
    '\u2192': '\\to', '\u2190': '\\leftarrow', '\u21D2': '\\Rightarrow', '\u21D0': '\\Leftarrow',
    '\u2208': '\\in', '\u2209': '\\notin', '\u2202': '\\partial', '\u2207': '\\nabla',
  };
  const joinLatex = (parts) => {
    let s = '';
    for (const p of parts) {
      if (!p) continue;
      if (s && /\\[a-zA-Z]+$/.test(s) && /^[A-Za-z0-9]/.test(p)) s += ' ';
      s += p;
    }
    return s;
  };
  const braceAtom = (s) => {
    if (!s) return '{}';
    if (/^[A-Za-z0-9]$/.test(s) || /^\\[a-zA-Z]+$/.test(s)) return s;
    if (/^\\(left|frac|sqrt|sum|prod|int|oint|mathrm|operatorname|hat|bar|vec|dot|tilde|overline|underline)/.test(s)) return s;
    return '{' + s + '}';
  };
  const scripted = (base, sub, sup) => {
    let s = braceAtom(base);
    if (sub && sub.trim()) s += '_{' + sub + '}';
    if (sup && sup.trim()) s += '^{' + sup + '}';
    return s;
  };
  const fence = (ch, side) => {
    const c = ch == null ? (side === 'l' ? '(' : ')') : ch;
    const cmd = side === 'l' ? '\\left' : '\\right';
    if (!c) return cmd + '.';
    if (c === '{' || c === '}') return cmd + '\\' + c;
    if (c === '\u2016' || c === '\u2225') return cmd + '\\|';
    if (c === '\u27E8') return cmd + '\\langle';
    if (c === '\u27E9') return cmd + '\\rangle';
    return cmd + c;
  };
  function charsToLatex(text) {
    let s = '', buf = '';
    const flush = () => { s += buf; buf = ''; };
    for (const ch of text) {
      const cmd = CHAR_CMD[ch];
      if (cmd) {
        flush();
        if (cmd.startsWith('\\') && s && /[a-zA-Z]$/.test(s)) s += ' ';
        s += cmd;
        continue;
      }
      if ('\\{}^_&#%'.includes(ch)) {
        flush();
        s += ch === '\\' ? '\\backslash ' : '\\mathrm{' + ch + '}';
        continue;
      }
      buf += ch;
    }
    flush();
    return s;
  }
  function runLatex(r) {
    let text = '';
    for (const c of kids(r)) if (lname(c) === 't') text += c.textContent;
    const sty = prop(r, 'rPr', 'sty');
    if ((sty === 'p' || sty === 'b') && FUN.has(text)) return '\\' + text;
    if ((sty === 'p' || sty === 'b') && text && !/[\\{}^_&#%]/.test(text)) return '\\mathrm{' + text + '}';
    return charsToLatex(text);
  }
  function argLatex(n) { return n ? joinLatex(kids(n).map(latexOf)) : ''; }
  function latexOf(c) {
    switch (lname(c)) {
      case 'r': return runLatex(c);
      case 'f': return '\\frac{' + argLatex(kid(c, 'num')) + '}{' + argLatex(kid(c, 'den')) + '}';
      case 'sSup': return scripted(argLatex(kid(c, 'e')), '', argLatex(kid(c, 'sup')));
      case 'sSub': return scripted(argLatex(kid(c, 'e')), argLatex(kid(c, 'sub')), '');
      case 'sSubSup': return scripted(argLatex(kid(c, 'e')), argLatex(kid(c, 'sub')), argLatex(kid(c, 'sup')));
      case 'sPre': return '{}_{' + argLatex(kid(c, 'sub')) + '}^{' + argLatex(kid(c, 'sup')) + '}' + braceAtom(argLatex(kid(c, 'e')));
      case 'd': {
        const parts = kids(c, 'e').map(argLatex);
        const sep = prop(c, 'dPr', 'sepChr');
        return fence(prop(c, 'dPr', 'begChr'), 'l') + parts.join(parts.length > 1 ? (sep == null ? '|' : sep) : '') + fence(prop(c, 'dPr', 'endChr'), 'r');
      }
      case 'rad': {
        const hide = onOff(prop(c, 'radPr', 'degHide'));
        const deg = argLatex(kid(c, 'deg')).trim();
        const e = argLatex(kid(c, 'e'));
        if (hide || !deg) return '\\sqrt{' + e + '}';
        return '\\sqrt[' + deg + ']{' + e + '}';
      }
      case 'nary': {
        const chr = prop(c, 'naryPr', 'chr') || '∫';
        const cmd = NARY_CMD[chr] || '\\operatorname{' + chr + '}';
        const sub = onOff(prop(c, 'naryPr', 'subHide')) ? '' : argLatex(kid(c, 'sub')).trim();
        const sup = onOff(prop(c, 'naryPr', 'supHide')) ? '' : argLatex(kid(c, 'sup')).trim();
        const e = argLatex(kid(c, 'e')).trim();
        let s = cmd;
        if (sub) s += '_{' + sub + '}';
        if (sup) s += '^{' + sup + '}';
        if (e) s += ' ' + e;
        return s;
      }
      case 'func': {
        const name = argLatex(kid(c, 'fName')).trim();
        const e = argLatex(kid(c, 'e')).trim();
        let s = /[\\^_]/.test(name) ? name : FUN.has(name) ? '\\' + name : '\\operatorname{' + name + '}';
        if (e) s += ' ' + e;
        return s;
      }
      case 'acc': {
        const ch = prop(c, 'accPr', 'chr') || '\u0302';
        return (ACC_CMD[ch] || '\\hat') + '{' + argLatex(kid(c, 'e')) + '}';
      }
      case 'bar': {
        const top = prop(c, 'barPr', 'pos') !== 'bot';
        return (top ? '\\overline{' : '\\underline{') + argLatex(kid(c, 'e')) + '}';
      }
      case 'm': return '\\matrix{' + kids(c, 'mr').map((mr) => kids(mr, 'e').map(argLatex).join('&')).join('\\\\') + '}';
      case 'eqArr': return kids(c, 'e').map(argLatex).join('\\\\');
      case 'limLow': return scripted(argLatex(kid(c, 'e')), argLatex(kid(c, 'lim')), '');
      case 'limUpp': return scripted(argLatex(kid(c, 'e')), '', argLatex(kid(c, 'lim')));
      case 'groupChr': case 'box': case 'borderBox': case 'phant': return argLatex(kid(c, 'e'));
      case 'oMath': case 'ins': case 'sdt': case 'sdtContent': case 'customXml': case 'smartTag': case 'moveTo': return argLatex(c);
      case 'del': case 'moveFrom': case 'oMathParaPr': case 'ctrlPr': return '';
      default: return '';
    }
  }
  function runsOnly(n) {
    for (const c of kids(n)) {
      const name = lname(c);
      if (name === 'r') continue;
      if ((name === 'ins' || name === 'sdt' || name === 'sdtContent' || name === 'customXml' || name === 'smartTag' || name === 'moveTo' || name === 'oMath') && runsOnly(c)) continue;
      return false;
    }
    return true;
  }
  function zoneText(n) {
    let s = '';
    for (const c of kids(n)) {
      if (lname(c) === 'r') { for (const t of kids(c)) if (lname(t) === 't') s += t.textContent; }
      else s += zoneText(c);
    }
    return s;
  }
  function looksLikeLatex(s) {
    if (/\\[a-zA-Z]/.test(s)) return true;
    if (/[∫∑∏√]/.test(s)) return false;
    return /[\^_]/.test(s);
  }
  function unicodeToLatex(s) {
    let t = s.trim();
    t = t.replace(/√\(([^()]*)\)/g, '\\sqrt{$1}').replace(/√\s*([A-Za-z0-9]+)/g, '\\sqrt{$1}');
    t = t.replace(/∑/g, '\\sum ').replace(/∏/g, '\\prod ').replace(/∫/g, '\\int ').replace(/∮/g, '\\oint ');
    t = t.replace(/\u2026/g, '\\ldots').replace(/\u22EF/g, '\\cdots').replace(/\u221E/g, '\\infty');
    const atom = '(?:\\\\(?:sqrt|sum|prod|int|oint|ldots|cdots|infty)(?:\\{[^{}]*\\})*|\\{[^{}]*\\}|\\([^()]*\\)|[A-Za-z0-9]+)';
    for (let i = 0; i < 6; i++) {
      const next = t.replace(new RegExp('(' + atom + ')\\s*/\\s*(' + atom + ')'), '\\frac{$1}{$2}');
      if (next === t) break;
      t = next;
    }
    return t.replace(/[ \t]+/g, ' ').trim();
  }
  function oneLatex(om) {
    if (runsOnly(om)) {
      const text = zoneText(om).trim();
      if (looksLikeLatex(text)) return text;
      return unicodeToLatex(text);
    }
    return argLatex(om);
  }

  /** LaTeX for an m:oMath / m:oMathPara element or its serialized XML. Empty when it cannot be read. */
  M.toLatex = function (src) {
    let node = src;
    if (typeof src === 'string') {
      if (!src) return '';
      const xml = /xmlns:m=/.test(src) ? src : src.replace(/^<m:(oMath\w*)/, '<m:$1 xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"');
      const doc = new DOMParser().parseFromString(xml, 'application/xml');
      if (doc.getElementsByTagName('parsererror').length) return '';
      node = doc.documentElement;
    }
    if (!node) return '';
    if (lname(node) === 'oMathPara') {
      const maths = kids(node, 'oMath');
      return maths.map(oneLatex).filter(Boolean).join('\\\\');
    }
    return oneLatex(node);
  };

  /**
   * Justification of a display equation: m:oMathParaPr/m:jc. Omitted means Centered as Group,
   * which is what Word uses when the element is absent.
   */
  M.jcOf = function (xml) {
    const pr = /<m:oMathParaPr\b[^>]*>[\s\S]*?<\/m:oMathParaPr>/.exec(String(xml || ''));
    if (!pr) return 'centerGroup';
    const m = /<m:jc\b[^>]*\bm:val="(left|right|center|centerGroup)"/.exec(pr[0]);
    return m ? m[1] : 'centerGroup';
  };

  /** Set m:jc on a display equation. The equation body is kept. An inline equation is unchanged. */
  M.setJc = function (xml, jc) {
    if (jc !== 'left' && jc !== 'right' && jc !== 'center' && jc !== 'centerGroup') return xml;
    const s = String(xml || '').trim();
    if (!/^<m:oMathPara[\s>]/.test(s)) return s;
    const tag = `<m:jc m:val="${jc}"/>`;
    const empty = /<m:oMathParaPr\b[^>]*\/>/.exec(s);
    if (empty) return s.slice(0, empty.index) + '<m:oMathParaPr>' + tag + '</m:oMathParaPr>' + s.slice(empty.index + empty[0].length);
    const pr = /<m:oMathParaPr\b[^>]*>[\s\S]*?<\/m:oMathParaPr>/.exec(s);
    if (!pr) return s.replace(/^(<m:oMathPara\b[^>]*>)/, '$1<m:oMathParaPr>' + tag + '</m:oMathParaPr>');
    let body = pr[0];
    if (/<m:jc\b/.test(body)) body = body.replace(/<m:jc\b[^>]*\/>/, tag).replace(/<m:jc\b[^>]*>[\s\S]*?<\/m:jc>/, tag);
    else body = body.replace(/(<m:oMathParaPr\b[^>]*>)/, '$1' + tag);
    return s.slice(0, pr.index) + body + s.slice(pr.index + pr[0].length);
  };

  /**
   * Inline m:oMath wrapped as a display m:oMathPara. Namespaces move to the outer element.
   * An equation that is already display is returned unchanged.
   */
  M.asDisplay = function (xml) {
    const s = String(xml || '').trim();
    if (/^<m:oMathPara[\s>]/.test(s)) return s;
    const m = /^<m:oMath\b([^>]*)>/.exec(s);
    if (!m) return '<m:oMathPara>' + s + '</m:oMathPara>';
    return '<m:oMathPara' + m[1] + '><m:oMath>' + s.slice(m[0].length) + '</m:oMathPara>';
  };

  /**
   * Display m:oMathPara unwrapped to its m:oMath. Drops m:oMathParaPr.
   * Namespaces move onto the m:oMath. Several equations become one.
   */
  M.asInline = function (xml) {
    const s = String(xml || '').trim();
    const m = /^<m:oMathPara\b([^>]*)>([\s\S]*)<\/m:oMathPara>$/.exec(s);
    if (!m) return s;
    let inner = m[2].replace(/^\s*<m:oMathParaPr\b[\s\S]*?<\/m:oMathParaPr>\s*/, '');
    const parts = [...inner.matchAll(/<m:oMath\b[^>]*>[\s\S]*?<\/m:oMath>/g)];
    if (parts.length > 1) {
      const bodies = parts.map((p) => p[0].replace(/^<m:oMath\b[^>]*>/, '').replace(/<\/m:oMath>$/, ''));
      inner = '<m:oMath>' + bodies.join('') + '</m:oMath>';
    }
    if (m[1] && /xmlns/.test(m[1]) && /^<m:oMath\b(?![^>]*\bxmlns\b)/.test(inner)) inner = inner.replace(/^<m:oMath\b/, '<m:oMath' + m[1]);
    return inner.trim();
  };
})();
