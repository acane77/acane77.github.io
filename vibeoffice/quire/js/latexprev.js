/* Quire — LaTeX for the equation editor. One parser builds a small tree; the preview draws MathML
 * from it and Update writes OMML from it. The two are not converted through each other.
 * An unknown command or a broken group is a syntax error. */
(function () {
  'use strict';
  const L = window.L;
  const P = (L.latex = {});
  const NS = 'http://www.w3.org/1998/Math/MathML';
  const CTRL = '<m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr></m:ctrlPr>';
  const mk = (tag, text, attrs) => {
    const e = document.createElementNS(NS, tag);
    if (text != null) e.textContent = text;
    if (attrs) for (const k in attrs) e.setAttribute(k, attrs[k]);
    return e;
  };
  const row = (nodes) => {
    if (nodes.length === 1) return nodes[0];
    const r = mk('mrow');
    for (const n of nodes) r.appendChild(n);
    return r;
  };
  const xmlEsc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const SYM = {
    alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ε', varepsilon: 'ε', zeta: 'ζ', eta: 'η', theta: 'θ', vartheta: 'ϑ',
    iota: 'ι', kappa: 'κ', lambda: 'λ', mu: 'μ', nu: 'ν', xi: 'ξ', pi: 'π', varpi: 'ϖ', rho: 'ρ', sigma: 'σ', tau: 'τ',
    upsilon: 'υ', phi: 'φ', varphi: 'ϕ', chi: 'χ', psi: 'ψ', omega: 'ω',
    Gamma: 'Γ', Delta: 'Δ', Theta: 'Θ', Lambda: 'Λ', Xi: 'Ξ', Pi: 'Π', Sigma: 'Σ', Phi: 'Φ', Psi: 'Ψ', Omega: 'Ω',
    infty: '∞', partial: '∂', nabla: '∇', pm: '±', mp: '∓', times: '×', div: '÷', cdot: '⋅', circ: '∘', bullet: '•',
    leq: '≤', geq: '≥', neq: '≠', approx: '≈', equiv: '≡', sim: '∼', simeq: '≃', ll: '≪', gg: '≫',
    to: '→', rightarrow: '→', leftarrow: '←', Rightarrow: '⇒', Leftarrow: '⇐', leftrightarrow: '↔',
    ldots: '…', cdots: '⋯', vdots: '⋮', ddots: '⋱', forall: '∀', exists: '∃', in: '∈', notin: '∉',
    subset: '⊂', supset: '⊃', cup: '∪', cap: '∩', emptyset: '∅',
    lbrace: '{', rbrace: '}', langle: '⟨', rangle: '⟩', lvert: '|', rvert: '|', Vert: '‖',
  };
  const NARY = { sum: '∑', prod: '∏', int: '∫', oint: '∮', bigcup: '⋃', bigcap: '⋂' };
  const FUN = new Set(['sin', 'cos', 'tan', 'cot', 'sec', 'csc', 'log', 'ln', 'lim', 'max', 'min', 'sup', 'inf', 'det', 'gcd', 'exp', 'arcsin', 'arccos', 'arctan', 'sinh', 'cosh', 'tanh']);
  const ACC = { hat: 'ˆ', bar: '¯', vec: '→', dot: '˙', tilde: '˜' };
  const ACC_OMML = { hat: '\u0302', bar: '\u0304', vec: '\u20D7', dot: '\u0307', tilde: '\u0303' };
  const REL = new Set(['leq', 'geq', 'neq', 'approx', 'equiv', 'sim', 'simeq', 'to', 'rightarrow', 'leftarrow', 'Rightarrow', 'Leftarrow', 'leftrightarrow', 'in', 'notin', 'subset', 'supset', 'right']);

  function Parser(s) { this.s = s; this.i = 0; }
  Parser.prototype.skip = function () { while (this.i < this.s.length && /\s/.test(this.s[this.i])) this.i++; };
  Parser.prototype.done = function () { return this.i >= this.s.length; };
  Parser.prototype.stop = function (extra) {
    if (this.done() || this.s[this.i] === '}' || this.s.startsWith('\\right', this.i) || this.s.startsWith('\\\\', this.i)) return true;
    return !!(extra && extra.includes(this.s[this.i]));
  };
  Parser.prototype.peekCmd = function () {
    if (this.s[this.i] !== '\\') return '';
    let j = this.i + 1;
    if (!/[a-zA-Z]/.test(this.s[j] || '')) return this.s[j] || '';
    let name = '';
    while (/[a-zA-Z]/.test(this.s[j] || '')) name += this.s[j++];
    return name;
  };
  Parser.prototype.command = function () {
    this.i++;
    const c = this.s[this.i];
    if (c === ' ' || c === ',' || c === ';' || c === '!' || c === ':') { this.i++; return c; }
    let name = '';
    if (c && /[a-zA-Z]/.test(c)) { while (this.i < this.s.length && /[a-zA-Z]/.test(this.s[this.i])) name += this.s[this.i++]; return name; }
    if (this.i < this.s.length) name = this.s[this.i++];
    return name;
  };
  Parser.prototype.group = function () {
    this.skip();
    if (this.s[this.i] !== '{') throw new Error('{');
    this.i++;
    const items = this.seq();
    this.skip();
    if (this.s[this.i] !== '}') throw new Error('}');
    this.i++;
    return { k: 'row', items };
  };
  Parser.prototype.textGroup = function () {
    this.skip();
    if (this.s[this.i] !== '{') throw new Error('{');
    this.i++;
    let s = '', depth = 1;
    while (this.i < this.s.length && depth) {
      const c = this.s[this.i++];
      if (c === '{') { depth++; s += c; }
      else if (c === '}') { depth--; if (depth) s += c; }
      else s += c;
    }
    if (depth) throw new Error('}');
    return s;
  };
  Parser.prototype.delim = function () {
    this.skip();
    if (this.s[this.i] === '\\') {
      const name = this.command();
      const d = { '{': '{', '}': '}', '|': '‖', lbrace: '{', rbrace: '}', langle: '⟨', rangle: '⟩', lvert: '|', rvert: '|', Vert: '‖', '.': '' }[name];
      if (d == null) throw new Error('delim');
      return d;
    }
    const c = this.s[this.i++];
    if (c === '.') return '';
    if ('()[]|'.includes(c)) return c;
    throw new Error('delim');
  };
  Parser.prototype.scriptArg = function () {
    this.skip();
    if (this.s[this.i] === '{') return this.group();
    const a = this.atom();
    if (!a) throw new Error('script');
    return a;
  };
  Parser.prototype.limits = function (node) {
    for (;;) {
      this.skip();
      const c = this.s[this.i];
      if (c !== '^' && c !== '_') break;
      this.i++;
      const arg = this.scriptArg();
      if (c === '^') { if (node.sup) throw new Error('^'); node.sup = arg; }
      else { if (node.sub) throw new Error('_'); node.sub = arg; }
    }
  };
  Parser.prototype.takesOperand = function () {
    this.skip();
    if (this.stop()) return false;
    const c = this.s[this.i];
    if ('+-=<>&'.includes(c)) return false;
    return !REL.has(this.peekCmd());
  };
  Parser.prototype.matrix = function () {
    this.skip();
    if (this.s[this.i] !== '{') throw new Error('{');
    this.i++;
    const rows = [];
    let line = [];
    for (;;) {
      line.push(this.seq('&'));
      this.skip();
      if (this.s[this.i] === '&') { this.i++; continue; }
      if (this.s.startsWith('\\\\', this.i)) { this.i += 2; rows.push(line); line = []; continue; }
      if (this.s[this.i] === '}') { this.i++; rows.push(line); return rows; }
      throw new Error('matrix');
    }
  };
  Parser.prototype.atom = function () {
    this.skip();
    if (this.stop()) return null;
    const c = this.s[this.i];
    if (c === '{') return this.group();
    if (c === '^' || c === '_' || c === '&') throw new Error(c);
    if (c === '\\') {
      const name = this.command();
      if (name === 'frac' || name === 'dfrac' || name === 'tfrac' || name === 'cfrac') return { k: 'frac', num: this.group(), den: this.group() };
      if (name === 'sqrt') {
        this.skip();
        let deg = null;
        if (this.s[this.i] === '[') {
          this.i++;
          const parts = [];
          while (!this.done() && this.s[this.i] !== ']') {
            this.skip();
            if (this.s[this.i] === ']') break;
            const f = this.factor();
            if (!f) throw new Error('sqrt');
            parts.push(f);
          }
          if (this.s[this.i] !== ']') throw new Error('sqrt');
          this.i++;
          deg = { k: 'row', items: parts };
        }
        return { k: 'rad', e: this.group(), deg };
      }
      if (name === 'text' || name === 'mathrm' || name === 'mathbf' || name === 'operatorname') return { k: 'up', s: this.textGroup() };
      if (name === 'left') {
        const beg = this.delim();
        const items = this.seq();
        this.skip();
        if (!this.s.startsWith('\\right', this.i)) throw new Error('right');
        this.i += 6;
        return { k: 'd', beg, end: this.delim(), items };
      }
      if (name === 'matrix' || name === 'pmatrix' || name === 'bmatrix') {
        const mat = { k: 'mat', rows: this.matrix() };
        if (name === 'matrix') return mat;
        return { k: 'd', beg: name === 'pmatrix' ? '(' : '[', end: name === 'pmatrix' ? ')' : ']', items: [mat] };
      }
      if (name === 'overline') return { k: 'bar', top: true, e: this.group() };
      if (name === 'underline') return { k: 'bar', top: false, e: this.group() };
      if (ACC[name]) return { k: 'acc', chr: ACC[name], omml: ACC_OMML[name], stretch: name === 'vec' || name === 'bar', e: this.group() };
      if (NARY[name]) {
        const node = { k: 'nary', chr: NARY[name], sub: null, sup: null, e: null };
        this.limits(node);
        if (this.takesOperand()) {
          const e = this.factor();
          if (!e) throw new Error('nary');
          node.e = e;
        }
        return node;
      }
      if (name === ' ' || name === ',' || name === ';' || name === ':' || name === '!' || name === 'quad' || name === 'qquad') return { k: 'sp' };
      if (name === '{' || name === '}') return { k: 't', s: name };
      if (name === '|') return { k: 't', s: '‖' };
      if (name === 'backslash') return { k: 't', s: '\\' };
      if (FUN.has(name)) return { k: 'up', s: name };
      if (SYM[name]) return { k: 't', s: SYM[name] };
      throw new Error(name || 'cmd');
    }
    if (/[0-9]/.test(c)) {
      let n = '';
      while (this.i < this.s.length && /[0-9.]/.test(this.s[this.i])) n += this.s[this.i++];
      return { k: 't', s: n };
    }
    if (/[A-Za-z]/.test(c)) return { k: 't', s: this.s[this.i++] };
    this.i++;
    return { k: 't', s: c };
  };
  Parser.prototype.scripts = function (base) {
    let sup = null, sub = null;
    for (;;) {
      this.skip();
      const c = this.s[this.i];
      if (c !== '^' && c !== '_') break;
      this.i++;
      const arg = this.scriptArg();
      if (c === '^') { if (sup) throw new Error('^'); sup = arg; }
      else { if (sub) throw new Error('_'); sub = arg; }
    }
    if (sup && sub) return { k: 'ss', b: base, sub, sup };
    if (sup) return { k: 'sup', b: base, sup };
    if (sub) return { k: 'sub', b: base, sub };
    return base;
  };
  Parser.prototype.factor = function () {
    const a = this.atom();
    if (!a) return null;
    if (a.k === 'nary') return a;
    return this.scripts(a);
  };
  Parser.prototype.seq = function (extra) {
    const out = [];
    for (;;) {
      this.skip();
      if (this.stop(extra)) break;
      const f = this.factor();
      if (!f) throw new Error('parse');
      out.push(f);
    }
    return out;
  };

  function parseLines(src) {
    const text = String(src || '');
    if (!text.trim()) return null;
    const p = new Parser(text);
    const lines = [];
    for (;;) {
      lines.push(p.seq());
      p.skip();
      if (!p.s.startsWith('\\\\', p.i)) break;
      p.i += 2;
    }
    p.skip();
    if (!p.done()) return null;
    if (!lines.some((ln) => ln.length)) return null;
    return lines;
  }

  function mathText(s) {
    const nodes = [];
    const chars = Array.from(s);
    let i = 0;
    while (i < chars.length) {
      const ch = chars[i];
      if (ch === ' ') { nodes.push(mk('mspace', null, { width: '0.2em' })); i++; continue; }
      if (/[0-9]/.test(ch)) {
        let n = ch; i++;
        while (i < chars.length && /[0-9.]/.test(chars[i])) n += chars[i++];
        nodes.push(mk('mn', n));
        continue;
      }
      if (/[A-Za-z]/.test(ch) || /[Α-Ωα-ω]/.test(ch)) { nodes.push(mk('mi', ch)); i++; continue; }
      nodes.push(mk('mo', ch === '-' ? '\u2212' : ch, '()[]{}'.includes(ch) ? { fence: 'true', stretchy: 'false' } : null));
      i++;
    }
    return row(nodes.length ? nodes : [mk('mrow')]);
  }
  function mathOf(n) {
    if (!n) return mk('mrow');
    switch (n.k) {
      case 't': return mathText(n.s);
      case 'sp': return mk('mspace', null, { width: '0.2em' });
      case 'up': return mk('mi', n.s, { mathvariant: 'normal' });
      case 'row': return row(n.items.length ? n.items.map(mathOf) : [mk('mrow')]);
      case 'frac': { const f = mk('mfrac'); f.append(mathOf(n.num), mathOf(n.den)); return f; }
      case 'rad': {
        if (n.deg) { const r = mk('mroot'); r.append(mathOf(n.e), mathOf(n.deg)); return r; }
        const r = mk('msqrt'); r.appendChild(mathOf(n.e)); return r;
      }
      case 'sup': { const e = mk('msup'); e.append(mathOf(n.b), mathOf(n.sup)); return e; }
      case 'sub': { const e = mk('msub'); e.append(mathOf(n.b), mathOf(n.sub)); return e; }
      case 'ss': { const e = mk('msubsup'); e.append(mathOf(n.b), mathOf(n.sub), mathOf(n.sup)); return e; }
      case 'd': {
        const r = mk('mrow');
        if (n.beg) r.appendChild(mk('mo', n.beg, { fence: 'true', stretchy: 'true' }));
        for (const it of n.items) r.appendChild(mathOf(it));
        if (n.end) r.appendChild(mk('mo', n.end, { fence: 'true', stretchy: 'true' }));
        return r;
      }
      case 'acc': {
        const e = mk('mover', null, { accent: 'true' });
        e.append(mathOf(n.e), mk('mo', n.chr, { stretchy: n.stretch ? 'true' : 'false' }));
        return e;
      }
      case 'bar': {
        const e = mk(n.top ? 'mover' : 'munder', null, n.top ? { accent: 'true' } : { accentunder: 'true' });
        e.append(mathOf(n.e), mk('mo', n.top ? '\u00AF' : '_', { stretchy: 'true' }));
        return e;
      }
      case 'nary': {
        const op = mk('mo', n.chr, { largeop: 'true' });
        let base = op;
        if (n.sub && n.sup) { base = mk('msubsup'); base.append(op, mathOf(n.sub), mathOf(n.sup)); }
        else if (n.sub) { base = mk('msub'); base.append(op, mathOf(n.sub)); }
        else if (n.sup) { base = mk('msup'); base.append(op, mathOf(n.sup)); }
        return n.e ? row([base, mathOf(n.e)]) : base;
      }
      case 'mat': {
        const t = mk('mtable');
        for (const rw of n.rows) {
          const tr = mk('mtr');
          for (const cell of rw) {
            const td = mk('mtd');
            td.appendChild(row(cell.length ? cell.map(mathOf) : [mk('mrow')]));
            tr.appendChild(td);
          }
          t.appendChild(tr);
        }
        return t;
      }
      default: return mk('mrow');
    }
  }
  function mathLines(lines) {
    const math = mk('math');
    math.setAttribute('displaystyle', 'true');
    if (lines.length === 1) { for (const n of lines[0]) math.appendChild(mathOf(n)); return math; }
    const table = mk('mtable', null, { columnalign: 'left' });
    for (const ln of lines) {
      const tr = mk('mtr'), td = mk('mtd');
      td.appendChild(row(ln.length ? ln.map(mathOf) : [mk('mrow')]));
      tr.appendChild(td);
      table.appendChild(tr);
    }
    math.appendChild(table);
    return math;
  }

  function run(text, upright) {
    if (!text) return '';
    const space = /^\s|\s$/.test(text) ? ' xml:space="preserve"' : '';
    const sty = upright ? '<m:rPr><m:sty m:val="p"/></m:rPr>' : '';
    return `<m:r>${sty}<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t${space}>${xmlEsc(text)}</m:t></m:r>`;
  }
  function emitItems(items) {
    let out = '', buf = '', up = false;
    const flush = () => { if (buf) { out += run(buf, up); buf = ''; } };
    for (const n of items) {
      if (n.k === 't' || n.k === 'sp') {
        if (up) flush();
        up = false;
        buf += n.k === 'sp' ? ' ' : n.s;
      } else if (n.k === 'up') {
        if (!up) flush();
        up = true;
        buf += n.s;
      } else { flush(); up = false; out += emitNode(n); }
    }
    flush();
    return out;
  }
  function emitBase(n) {
    if (!n) return '';
    if (n.k === 'row') return emitItems(n.items);
    if (n.k === 't' || n.k === 'sp' || n.k === 'up') return emitItems([n]);
    return emitNode(n);
  }
  function emitNode(n) {
    switch (n.k) {
      case 'row': return emitItems(n.items);
      case 't': case 'sp': case 'up': return emitItems([n]);
      case 'frac': return `<m:f><m:fPr>${CTRL}</m:fPr><m:num>${emitBase(n.num)}</m:num><m:den>${emitBase(n.den)}</m:den></m:f>`;
      case 'sup': return `<m:sSup><m:sSupPr>${CTRL}</m:sSupPr><m:e>${emitBase(n.b)}</m:e><m:sup>${emitBase(n.sup)}</m:sup></m:sSup>`;
      case 'sub': return `<m:sSub><m:sSubPr>${CTRL}</m:sSubPr><m:e>${emitBase(n.b)}</m:e><m:sub>${emitBase(n.sub)}</m:sub></m:sSub>`;
      case 'ss': return `<m:sSubSup><m:sSubSupPr>${CTRL}</m:sSubSupPr><m:e>${emitBase(n.b)}</m:e><m:sub>${emitBase(n.sub)}</m:sub><m:sup>${emitBase(n.sup)}</m:sup></m:sSubSup>`;
      case 'rad': {
        const hide = n.deg ? '' : '<m:degHide m:val="1"/>';
        return `<m:rad><m:radPr>${hide}${CTRL}</m:radPr><m:deg>${n.deg ? emitBase(n.deg) : ''}</m:deg><m:e>${emitBase(n.e)}</m:e></m:rad>`;
      }
      case 'd': {
        let pr = '';
        if (n.beg === '') pr += '<m:begChr m:val=""/>';
        else if (n.beg && n.beg !== '(') pr += `<m:begChr m:val="${xmlEsc(n.beg)}"/>`;
        if (n.end === '') pr += '<m:endChr m:val=""/>';
        else if (n.end && n.end !== ')') pr += `<m:endChr m:val="${xmlEsc(n.end)}"/>`;
        return `<m:d><m:dPr>${pr}${CTRL}</m:dPr><m:e>${emitItems(n.items)}</m:e></m:d>`;
      }
      case 'nary': {
        const hs = n.sub ? '' : '<m:subHide m:val="1"/>';
        const hp = n.sup ? '' : '<m:supHide m:val="1"/>';
        return `<m:nary><m:naryPr><m:chr m:val="${xmlEsc(n.chr)}"/>${hs}${hp}${CTRL}</m:naryPr><m:sub>${n.sub ? emitBase(n.sub) : ''}</m:sub><m:sup>${n.sup ? emitBase(n.sup) : ''}</m:sup><m:e>${n.e ? emitBase(n.e) : ''}</m:e></m:nary>`;
      }
      case 'acc': return `<m:acc><m:accPr><m:chr m:val="${xmlEsc(n.omml || n.chr)}"/>${CTRL}</m:accPr><m:e>${emitBase(n.e)}</m:e></m:acc>`;
      case 'bar': {
        const pos = n.top ? '' : '<m:pos m:val="bot"/>';
        return `<m:bar><m:barPr>${pos}${CTRL}</m:barPr><m:e>${emitBase(n.e)}</m:e></m:bar>`;
      }
      case 'mat': return `<m:m><m:mPr>${CTRL}</m:mPr>${n.rows.map((rw) => `<m:mr>${rw.map((cell) => `<m:e>${emitItems(cell)}</m:e>`).join('')}</m:mr>`).join('')}</m:m>`;
      default: return '';
    }
  }

  /** MathML for a LaTeX string, or null when the string has a syntax error. */
  P.toMathML = function (src) {
    try {
      const lines = parseLines(src);
      return lines ? mathLines(lines) : null;
    } catch (e) { return null; }
  };

  /**
   * OMML for a LaTeX string, or null when it does not parse.
   * opts.para wraps m:oMathPara (and keeps opts.paraPr, the original m:oMathParaPr, when there is one).
   */
  P.toOMML = function (src, opts) {
    let lines;
    try { lines = parseLines(src); } catch (e) { return null; }
    if (!lines) return null;
    const inner = lines.length > 1
      ? `<m:eqArr><m:eqArrPr>${CTRL}</m:eqArrPr>${lines.map((ln) => `<m:e>${emitItems(ln)}</m:e>`).join('')}</m:eqArr>`
      : emitItems(lines[0]);
    const om = `<m:oMath>${inner}</m:oMath>`;
    return opts && opts.para ? `<m:oMathPara>${opts.paraPr || ''}${om}</m:oMathPara>` : om;
  };
})();
