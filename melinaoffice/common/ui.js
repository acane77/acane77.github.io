/* Melina Office — UI toolkit shared by Ecrivaine, Matheuse and Oratrice: commands, menus, toolbars, dialogs, pickers.
 * Styled after the Office 2003 "Luna" command bars.
 */
(function (root) {
  'use strict';
  const L = root.L || (root.L = {});
  const { h } = L;
  const ui = (L.ui = {});

  /* What the app's editor does around commands, toolbar boxes and dialogs; each app fills these in (app.js).
   * beforeExec: bring the selection model up to date; refocus: give the keyboard back to the document;
   * commitEdit: finish an edit in progress before a toolbar box takes the focus; ownsText(el): el is the document's
   * own editable text, which refocus gives back with the app's selection (other editable text gets its own back). */
  ui.hooks = { beforeExec() {}, refocus() {}, commitEdit() {}, ownsText() { return false; } };
  /* Theme colours for the colour menus (Oratrice): schemeColors(design) -> [{ ref, hex, tip }], currentDesign() */
  ui.schemeColors = null;
  ui.currentDesign = null;

  /* ---------- command registry ---------- */
  /* what the user is doing, for whoever follows along (the assistant, common/scout/assistant.js), as events on
   * window: vo-command { id, result } after a command ran, vo-busy { on, text } when work that keeps the user
   * waiting starts or ends, vo-message { icon } when a message box opens */
  const tell = (type, detail) => { if (typeof window !== 'undefined' && window.dispatchEvent) window.dispatchEvent(new CustomEvent(type, { detail })); };

  ui.cmds = {};
  ui.cmd = (id, def) => { const o = { id }; Object.defineProperties(o, Object.getOwnPropertyDescriptors(def)); ui.cmds[id] = o; return o; };
  ui.exec = (id, arg, ev) => {
    const c = ui.cmds[id];
    if (!c) { console.warn('unknown command', id); return; }
    if (c.enabled && !c.enabled()) return;
    ui.hooks.beforeExec();
    let r;
    try { r = c.run(arg, ev); } catch (e) { console.error(e); ui.msg(t('That action could not be completed: {0}', e && e.message ? e.message : e), { icon: 'warn' }); }
    ui.refresh();
    tell('vo-command', { id, result: r });
    return r;
  };
  ui.enabled = (id) => { const c = ui.cmds[id]; return !!c && (!c.enabled || !!c.enabled()); };
  ui.checked = (id) => { const c = ui.cmds[id]; return !!c && !!c.checked && !!c.checked(); };
  /* UI text is English in the code and shown in the UI's language (common/i18n.js; docs/i18n.md): labels through
     stripAmp / ampHTML, the rest through t */
  const t = (s, ...args) => (L.t ? L.t(s, ...args) : s);
  /* "&x" marks the access key; "&&" and "& " stay a literal ampersand; "|chart" after the text is its context
     (common/i18n.js), not shown */
  const plainAmp = (s) => (root.VO && root.VO.i18n ? root.VO.i18n.plain(String(s || '')) : String(s || '')).replace(/&&/g, '\u0000').replace(/&(\S)/g, '$1').replace(/\u0000/g, '&');
  /** a label as shown, without its access key mark */
  const stripAmp = (s) => plainAmp(L.dropCjkAccess(t(s)));
  /** text as HTML, its access key underlined */
  const ampMark = (s) => L.esc(s || '').replace(/&amp;&amp;/g, '\u0000').replace(/&amp;(\S)/g, '<u>$1</u>').replace(/\u0000/g, '&amp;');
  /** a label as shown, its access key underlined */
  const ampHTML = (s) => ampMark(L.dropCjkAccess(t(s)));
  const keyOf = (s) => { const m = /&([^&\s])/.exec(s || ''); return m ? m[1].toLowerCase() : null; };
  /** the access key of a label as shown */
  const accessKey = (s) => keyOf(t(s));
  /** a shortcut as shown: its key names in the UI's language (Ctrl+Shift+F12 is Strg+Umschalt+F12 in German), each
      in the context "key" (a key's Home, not a Home button's). On a Mac, where ⌘ does what Ctrl does, the modifiers
      are Apple's symbols in Apple's order, without "+": Ctrl+Shift+S is ⇧⌘S */
  const MAC_MODS = [['Alt', '⌥'], ['Shift', '⇧'], ['Ctrl', '⌘']];
  const keyName = (p) => ({ Ctrl: t('Ctrl|key'), Shift: t('Shift|key'), Alt: t('Alt|key'), Del: t('Del|key'), Ins: t('Ins|key'), Esc: t('Esc|key'), Home: t('Home|key'), End: t('End|key'), Enter: t('Enter|key'), Tab: t('Tab|key'), Space: t('Space|key'), Backspace: t('Backspace|key'), 'Page Up': t('Page Up|key'), 'Page Down': t('Page Down|key'), PgUp: t('PgUp|key'), PgDn: t('PgDn|key') })[p] || p;
  const keyText = (k) => {
    if (!k) return '';
    const parts = String(k).split('+');
    if (!L.isMac) return parts.map(keyName).join('+');
    return MAC_MODS.filter(([m]) => parts.includes(m)).map(([, s]) => s).join('') + parts.filter((p) => !MAC_MODS.some(([m]) => m === p)).map(keyName).join('+');
  };
  ui.stripAmp = stripAmp;
  ui.plainAmp = plainAmp;
  ui.ampHTML = ampHTML;
  ui.accessKey = accessKey;
  ui.keyText = keyText;

  /* ---------- tooltips ---------- */
  let tipEl = null, tipTimer = null, tipTarget = null;
  function showTip(target, x, y) {
    let t = target.getAttribute('data-tip');
    if (!t) return;
    if (!target.closest('[translate="no"]') && L.t) t = L.t(t);   // data (a font's name) stays as it is
    /* Ribbon Golden Gate shows no shortcut keys: a tip's trailing "(Ctrl+S)" or "(⌘S)" comes off */
    if (document.documentElement.dataset.look === 'goldengate' || document.documentElement.dataset.look === 'goldengatedark' || document.documentElement.dataset.look === 'vscodegg') t = t.replace(/\s*\([^)]*\)\s*$/, '');
    if (!t) return;
    if (!tipEl) { tipEl = h('div', { class: 'tooltip', role: 'tooltip' }); document.body.appendChild(tipEl); }
    tipEl.textContent = t;
    tipEl.hidden = false;
    const r = tipEl.getBoundingClientRect();
    let left = x + 2, top = y + 20;
    if (left + r.width > window.innerWidth - 4) left = window.innerWidth - r.width - 4;
    if (top + r.height > window.innerHeight - 4) top = y - r.height - 6;
    tipEl.style.left = Math.max(2, left) + 'px';
    tipEl.style.top = Math.max(2, top) + 'px';
  }
  function hideTip() { clearTimeout(tipTimer); tipTarget = null; if (tipEl) tipEl.hidden = true; }
  if (typeof document !== 'undefined') {
    document.addEventListener('pointerover', (e) => {
      const t = e.target.closest && e.target.closest('[data-tip]');
      if (t === tipTarget) return;
      hideTip();
      if (!t || e.pointerType === 'touch') return;
      tipTarget = t;
      const x = e.clientX, y = e.clientY;
      tipTimer = setTimeout(() => showTip(t, x, y), 650);
    });
    document.addEventListener('pointerdown', hideTip, true);
    document.addEventListener('keydown', hideTip, true);
  }
  ui.hideTip = hideTip;

  /* ---------- menus ---------- */
  const openStack = []; /* stack of open menu panels */
  let menuBarState = null; /* {bar, index} */
  function resolveItem(it) {
    if (typeof it === 'string') {
      if (it === '-' || it === '|') return { sep: true };
      const c = ui.cmds[it];
      if (!c) return { label: it, disabled: true };
      return { cmd: it, label: c.menuLabel || c.label, icon: c.icon, key: c.key, enabled: c.enabled, checked: c.checked, radio: c.radio };
    }
    if (it.cmd && ui.cmds[it.cmd]) { const c = ui.cmds[it.cmd]; return Object.assign({ label: c.menuLabel || c.label, icon: c.icon, key: c.key, enabled: c.enabled, checked: c.checked, radio: c.radio }, it); }
    return it;
  }
  ui.closeMenus = (keepDepth) => {
    keepDepth = keepDepth || 0;
    while (openStack.length > keepDepth) {
      const m = openStack.pop();
      (m.shell || m.el).remove();
      if (m.opts && m.opts.onClose) m.opts.onClose();
    }
    if (!keepDepth && menuBarState) {
      L.$$('.mb-item.open', menuBarState.bar).forEach((x) => x.classList.remove('open'));
      menuBarState = null;
    }
  };
  /**
   * Open a dropdown menu. items: array of item specs or a function returning them.
   * at: {x,y} or DOMRect (opens below) ; opts: {depth, onClose, fromBar, side}
   */
  ui.openMenu = function (items, at, opts) {
    opts = opts || {};
    const depth = opts.depth || 0;
    ui.closeMenus(depth);
    const list = (typeof items === 'function' ? items() : items).map(resolveItem).filter((x) => x && !(typeof x.hidden === 'function' ? x.hidden() : x.hidden));
    /* popup menus and right-click menus stay above every window, modal or not, and above the formatting bar */
    let menuZ = 6000;
    const above = (n) => { if (n >= menuZ) menuZ = n + 20; };
    for (const w of winStack) above(parseInt(w.overlay.style.zIndex, 10));
    const bar = document.querySelector('.minibar');
    if (bar && !bar.hidden) above(parseInt(bar.style.zIndex, 10));
    const el = h('div', { class: 'menu' + (opts.cls ? ' ' + opts.cls : '') + (winStack.length ? ' over-dlg' : ''), role: 'menu', tabindex: '-1' });
    /* the glass sits on a shell. The menu itself scrolls, and overflow on the blurred element paints a solid fill.
       position:relative drops the viewport as the containing block, so the percentage max-height never applies. */
    const shell = document.documentElement.dataset.look === 'goldengate' || document.documentElement.dataset.look === 'goldengatedark' || document.documentElement.dataset.look === 'vscodegg' ? h('div', { class: 'gg-menu-glass' }) : null;
    const host = shell || el;
    if (shell) {
      shell.appendChild(el);
      el.style.position = 'relative';
      el.style.maxHeight = Math.max(80, window.innerHeight - 8) + 'px';
      el.style.overflowY = 'auto';
    }
    host.style.zIndex = String(menuZ);
    const rows = [];
    list.forEach((it) => {
      if (it.sep) { el.appendChild(h('div', { class: 'm-sep' })); return; }
      if (it.heading) { el.appendChild(h('div', { class: 'm-head', text: it.heading })); return; }
      if (it.custom) { const c = it.custom(() => ui.closeMenus()); el.appendChild(c); return; }
      const disabled = it.disabled || (it.enabled && !it.enabled());
      const checked = it.checked && it.checked();
      const row = h('div', { class: 'm-item' + (disabled ? ' dis' : '') + (checked ? ' chk' : '') + (it.sub ? ' has-sub' : '') + (it.bold ? ' bold' : ''), role: 'menuitem', 'aria-disabled': disabled ? 'true' : null, title: it.tip || null });
      const icoHTML = it.html ? it.html : it.icon ? L.icons.get(it.icon) : checked ? (it.radio ? '<span class="m-radio">●</span>' : '<span class="m-check">✓</span>') : '';
      row.appendChild(h('span', { class: 'm-ico' + (checked && it.icon ? ' on' : ''), html: icoHTML }));
      /* it.raw: data shown as it is (a font's name in a font list, a file's in Recent Files), not UI text; its
         access key, if any, is the label's own */
      row.appendChild(h('span', { class: 'm-lbl', html: it.raw ? ampMark(it.label) : ampHTML(it.label), translate: it.raw ? 'no' : null }));
      row.appendChild(h('span', { class: 'm-key', text: it.raw ? it.key || '' : keyText(it.key), translate: it.raw ? 'no' : null }));
      row.appendChild(h('span', { class: 'm-sub', html: it.sub ? '▶' : '' }));
      row._it = it;
      row._mn = it.raw ? keyOf(it.label) : accessKey(it.label);
      row.addEventListener('pointerenter', () => { setActive(row); if (it.sub && !disabled) openSub(row); else ui.closeMenus(depth + 1); });
      row.addEventListener('click', (e) => { e.stopPropagation(); activate(row); });
      rows.push(row);
      el.appendChild(row);
    });
    let active = null;
    function setActive(r) { if (active) active.classList.remove('act'); active = r; if (r) r.classList.add('act'); }
    function openSub(row) {
      const r = row.getBoundingClientRect();
      ui.openMenu(row._it.sub, { x: r.right - 3, y: r.top - 3, side: r.left }, { depth: depth + 1 });
    }
    function activate(row) {
      const it = row._it;
      if (it.disabled || (it.enabled && !it.enabled())) return;
      if (it.sub) { openSub(row); return; }
      ui.closeMenus();
      if (it.run) it.run();
      else if (it.cmd) ui.exec(it.cmd, it.arg);
      if (!ui.dialogOpen() && keyboardLost()) ui.hooks.refocus();
    }
    el._keys = (e) => {
      const idx = rows.indexOf(active);
      const usable = rows.filter((r) => !r.classList.contains('dis'));
      if (e.key === 'ArrowDown') { const n = usable[(usable.indexOf(active) + 1) % usable.length]; setActive(n); return true; }
      if (e.key === 'ArrowUp') { const i = usable.indexOf(active); setActive(usable[(i - 1 + usable.length) % usable.length]); return true; }
      if (e.key === 'ArrowRight' && active && active._it.sub) { openSub(active); const sub = openStack[openStack.length - 1]; if (sub && sub.first) sub.first(); return true; }
      if (e.key === 'Enter' || e.key === ' ') { if (active) activate(active); return true; }
      if (e.key.length === 1) {
        const k = e.key.toLowerCase();
        const hit = usable.find((r) => r._mn === k);
        if (hit) { setActive(hit); activate(hit); return true; }
      }
      void idx;
      return false;
    };
    document.body.appendChild(host);
    /* position */
    const r = el.getBoundingClientRect();
    let x, y;
    if (at instanceof DOMRect || (at && at.bottom != null && at.left != null && at.x == null)) { x = at.left; y = at.bottom; }
    else if (at && at.bottom != null) { x = at.left; y = at.bottom; }
    else { x = at.x; y = at.y; }
    if (x + r.width > window.innerWidth - 2) x = at && at.side != null ? at.side - r.width + 3 : window.innerWidth - r.width - 2;
    if (y + r.height > window.innerHeight - 2) y = Math.max(2, window.innerHeight - r.height - 2);
    host.style.left = Math.max(2, x) + 'px';
    host.style.top = Math.max(2, y) + 'px';
    const entry = { el, shell, opts, first: () => setActive(rows.find((r2) => !r2.classList.contains('dis'))) };
    openStack.push(entry);
    if (opts.selectFirst) entry.first();
    return el;
  };
  ui.contextMenu = (e, items) => { e.preventDefault(); ui.openMenu(items, { x: e.clientX, y: e.clientY }); };
  if (typeof document !== 'undefined') {
    document.addEventListener('pointerdown', (e) => {
      if (!openStack.length) return;
      if (e.target.closest('.menu') || e.target.closest('.mb-item')) return;
      ui.closeMenus();
    }, true);
    window.addEventListener('blur', () => ui.closeMenus());
    window.addEventListener('resize', () => ui.closeMenus());
    /* The keyboard stays with the document when chrome is pressed: a toolbar's, the ribbon's or a pane's background, a
       grip, the status bar, a tab strip, a menu, the menu bar or a button outside a dialog. The browser would give it
       to <body> or to what was pressed, and the arrows, Delete and typing would stop working on the document until it
       is clicked. A control of its own (a box, a list, the document itself, anything in a dialog) still takes it. */
    document.addEventListener('mousedown', (e) => {
      const t = e.target;
      if (!t || !t.closest || t.closest('.dlg')) return;
      const f = t.closest('input, select, textarea, [contenteditable=""], [contenteditable="true"], a[href], label, [tabindex], button');
      if (f && !f.matches('button, .menu, .mb-item')) return;
      e.preventDefault();
    }, true);
  }
  /** after chrome that took the keyboard is done with it (a menu command, a dialog), the document gets it back */
  const keyboardLost = () => !document.activeElement || document.activeElement === document.body || !document.activeElement.isConnected;
  ui.menuOpen = () => openStack.length > 0;
  /* menu keyboard routing; returns true if handled */
  ui.menuKey = (e) => {
    if (!openStack.length) return false;
    const top = openStack[openStack.length - 1];
    if (e.key === 'Escape') {
      if (openStack.length > 1) ui.closeMenus(openStack.length - 1); else ui.closeMenus();
      return true;
    }
    if (e.key === 'ArrowLeft') {
      if (openStack.length > 1) { ui.closeMenus(openStack.length - 1); return true; }
      if (menuBarState) { menuBarState.step(-1); return true; }
    }
    if (e.key === 'ArrowRight' && menuBarState && !(top.el._keys && top.el.querySelector('.m-item.act.has-sub'))) { menuBarState.step(1); return true; }
    if (top.el._keys && top.el._keys(e)) return true;
    return true; /* swallow keys while a menu is open */
  };

  /* the keyboard of another command-bar layout (the ribbon, common/ribbon.js) while it is shown instead of the menus:
     { active(), byKey(ch) → handled, index(i) } */
  ui.barKeys = null;
  /** Menu bar: menus = [{label, items}] */
  ui.menuBar = function (container, menus) {
    const bar = h('div', { class: 'menubar', role: 'menubar' });
    const els = [];
    menus.forEach((m, i) => {
      const it = h('div', { class: 'mb-item', role: 'menuitem', tabindex: '-1', html: ampHTML(m.label) });
      it.addEventListener('pointerdown', (e) => { e.preventDefault(); if (it.classList.contains('open')) ui.closeMenus(); else open(i); });
      it.addEventListener('pointerenter', () => { if (menuBarState && menuBarState.index !== i) open(i); });
      els.push(it);
      bar.appendChild(it);
    });
    function open(i, selectFirst) {
      ui.closeMenus();
      const it = els[i];
      it.classList.add('open');
      const r = it.getBoundingClientRect();
      ui.openMenu(menus[i].items, { left: r.left, bottom: r.bottom - 1 }, { cls: 'from-bar', selectFirst });
      menuBarState = { bar, index: i, step: (d) => open((i + d + els.length) % els.length, true) };
    }
    bar.openByKey = (ch) => {
      if (ui.barKeys && ui.barKeys.active()) return ui.barKeys.byKey(ch);
      const i = menus.findIndex((m) => accessKey(m.label) === ch.toLowerCase());
      if (i >= 0) { open(i, true); return true; }
      return false;
    };
    bar.openIndex = (i) => { if (ui.barKeys && ui.barKeys.active()) ui.barKeys.index(i); else open(i, true); };
    bar.menus = ui.menus = menus;   // also read by Search commands (common/search.js)
    container.appendChild(bar);
    return bar;
  };

  /* ---------- toolbars ---------- */
  const tbButtons = [];
  ui.tbButton = function (id, opts) {
    opts = opts || {};
    const c = ui.cmds[id] || {};
    const label = opts.label != null ? opts.label : c.tbLabel;
    const b = h('button', { class: 'tb-btn' + (label ? ' with-label' : ''), type: 'button', 'data-cmd': id, 'aria-label': stripAmp(c.label || id), 'data-tip': stripAmp(c.tip || c.label || id) + (c.key ? ` (${keyText(c.key)})` : '') });
    b.innerHTML = (c.icon || opts.icon ? L.icons.get(opts.icon || c.icon) : '') + (label ? `<span class="tb-lbl">${ampHTML(label)}</span>` : '');
    b.addEventListener('pointerdown', (e) => e.preventDefault()); /* keep the focus and selection where they are */
    b.addEventListener('click', (e) => ui.exec(id, opts.arg, e));
    tbButtons.push(b);
    return b;
  };
  /** split button: main action + dropdown arrow */
  ui.tbSplit = function (id, menuFn, opts) {
    opts = opts || {};
    const wrap = h('div', { class: 'tb-split' });
    const main = ui.tbButton(id, opts);
    const dd = h('button', { class: 'tb-dd', type: 'button', 'aria-label': 'More options', 'data-tip': (ui.cmds[id] && stripAmp(ui.cmds[id].label)) || '' });
    dd.innerHTML = '<span class="dd-arrow"></span>';
    dd.addEventListener('pointerdown', (e) => e.preventDefault());
    dd.addEventListener('click', () => { const r = wrap.getBoundingClientRect(); menuFn({ left: r.left, bottom: r.bottom }, wrap); });
    wrap.append(main, dd);
    wrap.dataset.cmd = id;
    return wrap;
  };
  /** dropdown button (icon/label + arrow), opens a menu */
  ui.tbDrop = function (label, icon, menuFn, tip) {
    const b = h('button', { class: 'tb-btn tb-drop' + (label ? ' with-label' : ''), type: 'button', 'data-tip': tip || stripAmp(label), 'aria-label': tip || stripAmp(label) });
    b.innerHTML = (icon ? L.icons.get(icon) : '') + (label ? `<span class="tb-lbl">${ampHTML(label)}</span>` : '') + '<span class="dd-arrow"></span>';
    b.addEventListener('pointerdown', (e) => e.preventDefault());
    b.addEventListener('click', () => { const r = b.getBoundingClientRect(); menuFn({ left: r.left, bottom: r.bottom }, b); });
    return b;
  };
  /** type-ahead for an editable box: the typed prefix stays, the rest of the first matching option is selected.
   * o.values(): option strings. o.order(names, prefix): sorts them, preferred first (L.fontComplete for fonts).
   * o.onText(text): the name now shown (the match, or the prefix when nothing matches). */
  ui.bindComplete = function (inp, o) {
    let composing = false;
    const values = () => {
      const out = [], seen = new Set();
      for (const v of (o.values() || [])) {
        if (v == null || v === '') continue;
        const s = String(v), k = s.toLowerCase();
        if (seen.has(k)) continue;
        seen.add(k);
        out.push(s);
      }
      return out;
    };
    function fill() {
      if (composing) return;
      const caret = inp.selectionStart;
      if (caret == null || (inp.selectionEnd !== inp.value.length && caret !== inp.selectionEnd)) return;
      const prefix = inp.value.slice(0, caret);
      if (!prefix) { if (o.onText) o.onText(''); return; }
      const match = L.completeMatch(values(), prefix, o.order);
      if (match && match !== prefix) {
        const typed = match.toLowerCase() === prefix.toLowerCase();
        inp.value = match;
        inp.setSelectionRange(typed ? match.length : prefix.length, match.length);
        if (o.onText) o.onText(match);
      } else if (o.onText) o.onText(match || prefix);
    }
    inp.addEventListener('compositionstart', () => { composing = true; });
    inp.addEventListener('compositionend', () => { composing = false; fill(); });
    inp.addEventListener('input', fill);
    inp.addEventListener('keydown', (e) => {
      if (composing || inp.selectionEnd !== inp.value.length || inp.selectionStart === inp.selectionEnd) return;
      if (e.key === 'Backspace' && inp.selectionStart > 0) {
        e.preventDefault();
        const prefix = inp.value.slice(0, inp.selectionStart - 1);
        inp.value = prefix;
        inp.setSelectionRange(prefix.length, prefix.length);
        fill();
      } else if (e.key === 'Delete') {
        e.preventDefault();
        const prefix = inp.value.slice(0, inp.selectionStart);
        inp.value = prefix;
        inp.setSelectionRange(prefix.length, prefix.length);
        if (o.onText) o.onText(prefix);
      } else if (e.key === 'ArrowRight' || e.key === 'End') {
        e.preventDefault();
        const n = inp.value.length;
        inp.setSelectionRange(n, n);
      }
    });
  };
  /** editable combo box (font name / size / zoom); o.options() may include { heading } rows, { sep }, { value, label }
   * and { label, run } actions (the list reopens after the action);
   * o.note(v): grey text beside an option in the list, or { text, tip } with a tooltip for its row.
   * o.complete: type-ahead. true keeps the option order; a function(names, prefix) sorts them first (preferred first). */
  ui.combo = function (o) {
    const wrap = h('div', { class: 'combo', style: `width:${o.width || 120}px`, 'data-tip': o.tip || '' });
    const inp = h('input', { type: 'text', class: 'combo-in', id: o.id, 'aria-label': o.tip || o.id, autocomplete: 'off', spellcheck: 'false' });
    const btn = h('button', { class: 'combo-btn', type: 'button', tabindex: '-1', 'aria-label': 'Open list', html: '<span class="dd-arrow"></span>' });
    wrap.append(inp, btn);
    const commit = () => { o.onChange(inp.value); inp.blur(); };
    inp.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); commit(); ui.hooks.refocus(); }
      if (e.key === 'Escape') { e.preventDefault(); inp.value = o.value(); inp.blur(); ui.hooks.refocus(); }
      e.stopPropagation();
    });
    inp.addEventListener('focus', () => { inp.select(); wrap.classList.add('focus'); ui.hooks.commitEdit(); });
    inp.addEventListener('blur', () => wrap.classList.remove('focus'));
    btn.addEventListener('pointerdown', (e) => e.preventDefault());
    btn.addEventListener('click', () => {
      ui.hooks.commitEdit();
      const r = wrap.getBoundingClientRect();
      // options: strings, { heading } section rows, or { value, label }
      ui.openMenu(o.options().map((x) => {
        if (x && x.heading) return { heading: x.heading };
        if (x && x.sep) return { sep: true };
        if (x && x.run) return { label: x.label, run: async () => { await x.run(); btn.click(); } };   // an action, then the list again
        const v = x && typeof x === 'object' ? x.value : x;
        const note = o.note ? o.note(v) : '';
        return {
          label: String(x && x.label || v).replace(/&/g, '&&'), value: v, raw: true,
          html: '', run: () => { inp.value = v; o.onChange(String(v)); ui.hooks.refocus(); },
          checked: () => String(v) === String(o.value()),
          key: note && typeof note === 'object' ? note.text : note, tip: note && note.tip,
        };
      }), { left: r.left, bottom: r.bottom }, { cls: 'combo-list' + (o.listCls ? ' ' + o.listCls : '') });
      if (o.preview) {
        const menu = document.querySelector('.menu.combo-list');
        // each row in its own font once scrolled into view: a long list (the computer's fonts) opens at once
        const show = (row) => { const st = o.preview(row._it.value); if (st) L.$('.m-lbl', row).style.cssText = st; };
        const io = menu && 'IntersectionObserver' in window ? new IntersectionObserver((es) => es.forEach((e) => { if (e.isIntersecting) { show(e.target); io.unobserve(e.target); } }), { root: menu }) : null;
        if (menu) L.$$('.m-item', menu).forEach((row) => row._it.value == null ? null : io ? io.observe(row) : show(row));
      }
      const m = document.querySelector('.menu.combo-list .m-item.chk');
      if (m) m.scrollIntoView({ block: 'nearest' });
    });
    if (o.complete) {
      const names = () => {
        const out = [];
        for (const x of o.options()) {
          if (!x || (typeof x === 'object' && (x.heading || x.sep || x.run))) continue;
          out.push(x && typeof x === 'object' ? x.value : x);
        }
        return out;
      };
      ui.bindComplete(inp, { values: names, order: typeof o.complete === 'function' ? o.complete : null });
    }
    wrap._refresh = () => { if (document.activeElement !== inp) inp.value = o.value(); const en = o.enabled ? o.enabled() : true; inp.disabled = btn.disabled = !en; wrap.classList.toggle('dis', !en); };
    tbButtons.push(wrap);
    return wrap;
  };
  /** Build a toolbar row */
  ui.toolbar = function (id, title, items) {
    const tb = h('div', { class: 'toolbar', id: 'tb-' + id, role: 'toolbar', 'aria-label': title });
    tb.appendChild(h('span', { class: 'tb-grip', 'aria-hidden': 'true' }));
    for (const it of items) {
      if (it === '|') tb.appendChild(h('span', { class: 'tb-sep' }));
      else if (typeof it === 'string') tb.appendChild(ui.tbButton(it));
      else if (it instanceof Node) tb.appendChild(it);
    }
    tb.appendChild(h('span', { class: 'tb-end' }));
    return tb;
  };
  ui.refresh = L.rafThrottle(() => {
    for (const b of tbButtons) {
      if (!b.isConnected) continue;
      if (b._refresh) { b._refresh(); continue; }
      const id = b.dataset.cmd;
      const c = ui.cmds[id];
      if (!c) continue;
      const en = !c.enabled || c.enabled();
      b.disabled = !en;
      b.classList.toggle('dis', !en);
      b.classList.toggle('on', !!(c.checked && c.checked()));
      b.setAttribute('aria-pressed', c.checked ? String(!!c.checked()) : null);
      if (c.swatch) { const sw = c.swatch(); if (sw) b.style.setProperty('--icon-swatch', sw); }
    }
    L.bus.emit('ui-refresh');
  });

  /* ---------- dragging toolbars (Classic, Windows 7, Windows XP) ---------- */
  /* A grip (or a floating window's title) pulls a toolbar off its row. Near an edge it docks; anywhere else it
   * floats in the ordinary modeless window (isToolBoxWindow). The place is kept per app. Closing the window, or
   * turning the toolbar off, forgets it, so the next time it is shown it is back where the app first built it. */
  ui.tbDock = (function () {
    const DOCKS = { classic: 1, win7: 1, xp: 1 };
    const home = new Map();
    const defaults = {};
    const bars = new Map();
    let captured = false, dragOn = false, chromeSig = '';

    function on() {
      const html = document.documentElement;
      const look = html.dataset.look || 'classic';
      return html.dataset.layout !== 'ribbon' && !!DOCKS[look];
    }
    function idOf(tb) { return tb && tb.id && tb.id.indexOf('tb-') === 0 ? tb.id.slice(3) : ''; }
    function load() {
      const v = L.store.get('tbDock', {});
      return v && typeof v === 'object' ? v : {};
    }
    function save(map) { L.store.set('tbDock', map); }
    function spotOf(id) {
      const raw = load()[id];
      const d = defaults[id];
      if (!raw || !raw.dock) return Object.assign({}, d);
      if (raw.dock === 'float') return { dock: 'float', x: Number.isFinite(+raw.x) ? +raw.x : 8, y: Number.isFinite(+raw.y) ? +raw.y : 8, back: raw.back || { dock: d.dock, row: d.row, order: d.order } };
      const dock = raw.dock === 'bottom' || raw.dock === 'left' || raw.dock === 'right' ? raw.dock : 'top';
      return { dock, row: Math.max(0, +raw.row || 0), order: Math.max(0, +raw.order || 0) };
    }
    function commit(id, spot) {
      const map = load();
      map[id] = spot;
      save(map);
      sync();
    }
    /** forget a toolbar's place. The next sync puts it back on the row it was built on. */
    function reset(id) {
      const map = load();
      if (!map[id]) return;
      delete map[id];
      save(map);
    }
    function capture() {
      const top = document.getElementById('toolbars');
      if (!top) return;
      top.querySelectorAll(':scope > .tb-row').forEach((row, ri) => {
        [...row.children].forEach((tb, oi) => {
          if (!tb.classList || !tb.classList.contains('toolbar')) return;
          const id = idOf(tb);
          if (!id || defaults[id]) return;
          home.set(tb, { parent: row, index: oi });
          defaults[id] = { dock: 'top', row: ri, order: oi };
          bars.set(id, tb);
          tb.classList.add('tb-dockable');
        });
      });
      const draw = document.getElementById('drawrow');
      if (draw) [...draw.children].forEach((tb, oi) => {
        if (!tb.classList || !tb.classList.contains('toolbar')) return;
        const id = idOf(tb);
        if (!id || defaults[id]) return;
        home.set(tb, { parent: draw, index: oi });
        defaults[id] = { dock: 'bottom', row: 0, order: oi };
        bars.set(id, tb);
        tb.classList.add('tb-dockable');
      });
      captured = bars.size > 0;
    }
    function hostOf(dock) {
      if (dock === 'top') return document.getElementById('toolbars');
      if (dock === 'bottom') return document.getElementById('drawrow');
      return document.getElementById(dock === 'left' ? 'tb-dock-left' : 'tb-dock-right');
    }
    function ensureSides() {
      const main = document.querySelector('.app > .main');
      if (!main) return;
      if (!document.getElementById('tb-dock-left')) {
        const left = h('div', { id: 'tb-dock-left', class: 'tb-dock tb-dock-side tb-dock-left', hidden: true });
        main.insertBefore(left, main.firstChild);
      }
      if (!document.getElementById('tb-dock-right')) main.appendChild(h('div', { id: 'tb-dock-right', class: 'tb-dock tb-dock-side tb-dock-right', hidden: true }));
    }
    function bandsOf(host) {
      return [...host.children].filter((c) => c.classList.contains('tb-row') || c.classList.contains('tb-band'));
    }
    function bandAt(host, row) {
      let bands = bandsOf(host);
      while (bands.length <= row) {
        const b = h('div', { class: host.id === 'tb-dock-left' || host.id === 'tb-dock-right' ? 'tb-band' : 'tb-row tb-band', dataset: { extra: '1' } });
        host.appendChild(b);
        bands = bandsOf(host);
      }
      return bands[row];
    }
    function unwrap(tb) {
      const w = tb._toolwin;
      if (!w) return;
      tb._floatClosing = 1;
      tb._toolwin = null;
      if (w.el && w.el.isConnected) w.close();
      tb._floatClosing = 0;
    }
    function hideBar(id) {
      reset(id);
      const app = L.app;
      if (!app || !app.opts || !app.opts.toolbars) return;
      app.opts.toolbars[id] = false;
      if (app.saveOpts) app.saveOpts();
      if (app.updateToolbars) app.updateToolbars();
    }
    function quietApp() {
      const app = document.querySelector('.app');
      return !!(app && (app.classList.contains('reading') || app.classList.contains('preview') || app.classList.contains('fullscreen')));
    }
    function floatWin(tb) {
      const cur = tb._toolwin;
      if (cur && cur.el && cur.el.isConnected) return cur;
      const id = idOf(tb);
      const dlg = ui.dialog({
        title: tb.getAttribute('aria-label') || id,
        body: tb,
        modeless: true,
        isToolBoxWindow: 1,
        width: Math.min(Math.max((tb.offsetWidth || 240) + 28, 160), window.innerWidth - 16),
        onClose() {
          tb._toolwin = null;
          if (tb._floatClosing) return;
          hideBar(id);
        },
      });
      tb._toolwin = dlg;
      return dlg;
    }
    function placeFloat(tb, spot) {
      if (tb.hidden || quietApp()) {
        if (tb._toolwin && tb._toolwin.overlay) tb._toolwin.overlay.hidden = true;
        return;
      }
      const win = floatWin(tb);
      if (tb.parentNode !== win.body) win.body.appendChild(tb);
      const dlg = win.el;
      const x = Math.max(0, Math.min(spot.x, Math.max(0, window.innerWidth - 80)));
      const y = Math.max(0, Math.min(spot.y, Math.max(0, window.innerHeight - 40)));
      if (dlg.style.left !== x + 'px') dlg.style.left = x + 'px';
      if (dlg.style.top !== y + 'px') dlg.style.top = y + 'px';
      const w = Math.min(Math.max((tb.scrollWidth || tb.offsetWidth || 160) + 28, 160), window.innerWidth - 16);
      if (dlg.style.width !== w + 'px') dlg.style.width = w + 'px';
      win.overlay.hidden = false;
    }
    function restore() {
      const groups = new Map();
      home.forEach((h0, tb) => {
        unwrap(tb);
        if (!groups.has(h0.parent)) groups.set(h0.parent, []);
        groups.get(h0.parent).push({ tb, index: h0.index });
      });
      groups.forEach((list, parent) => {
        if (!parent.isConnected) return;
        list.sort((a, b) => a.index - b.index);
        if (list.some((it, i) => parent.children[i] !== it.tb)) list.forEach((it) => parent.appendChild(it.tb));
      });
      document.querySelectorAll('.tb-band[data-extra]').forEach((b) => b.remove());
      for (const id of ['tb-dock-left', 'tb-dock-right']) {
        const el = document.getElementById(id);
        if (el) el.hidden = true;
      }
    }
    function layout() {
      ensureSides();
      const groups = new Map();
      bars.forEach((tb, id) => {
        const spot = spotOf(id);
        if (spot.dock === 'float') { placeFloat(tb, spot); return; }
        if (tb._toolwin) unwrap(tb);
        const dock = spot.dock === 'left' || spot.dock === 'right' ? spot.dock : spot.dock === 'bottom' ? 'bottom' : 'top';
        const row = dock === 'left' || dock === 'right' ? 0 : spot.row;
        const key = dock + ':' + row;
        if (!groups.has(key)) groups.set(key, []);
        groups.get(key).push({ tb, order: spot.order });
      });
      groups.forEach((list, key) => {
        const dock = key.slice(0, key.indexOf(':'));
        const row = +key.slice(key.indexOf(':') + 1);
        const host = hostOf(dock);
        if (!host) return;
        list.sort((a, b) => a.order - b.order);
        const band = bandAt(host, row);
        const same = band.children.length === list.length && list.every((it, i) => band.children[i] === it.tb);
        if (!same) list.forEach((it) => band.appendChild(it.tb));
      });
      ['top', 'bottom', 'left', 'right'].forEach((dock) => {
        const host = hostOf(dock);
        if (!host) return;
        let any = false;
        bandsOf(host).forEach((b) => {
          const vis = [...b.children].some((c) => c.classList.contains('toolbar') && !c.hidden);
          b.hidden = !vis;
          if (vis) any = true;
          if (!b.children.length && b.dataset.extra) b.remove();
        });
        if (dock !== 'top') host.hidden = !any;
      });
      bars.forEach((tb) => {
        const w = tb._toolwin;
        if (w && w.overlay) w.overlay.hidden = !!tb.hidden || quietApp();
      });
      const sig = ['top', 'bottom', 'left', 'right'].map((dock) => {
        const host = hostOf(dock);
        if (!host || host.hidden) return '0';
        return String(bandsOf(host).filter((b) => !b.hidden).length);
      }).join('');
      if (sig !== chromeSig) {
        chromeSig = sig;
        if (L.layout && L.layout.fitSizer) L.layout.fitSizer();
        else if (L.ed && L.ed.layout) L.ed.layout();
      }
    }
    function sync() {
      if (dragOn || !document.getElementById('toolbars')) return;
      if (!captured) capture();
      if (!captured) return;
      const use = on();
      document.documentElement.classList.toggle('tb-dock', use);
      if (!use) { restore(); return; }
      layout();
    }
    function zoneAt(x, y) {
      const pad = 18;
      const box = (el) => {
        if (!el || el.hidden) return null;
        const r = el.getBoundingClientRect();
        return r.width > 2 && r.height > 2 ? r : null;
      };
      const inside = (r, q) => r && x >= r.left - q && x <= r.right + q && y >= r.top - q && y <= r.bottom + q;
      const top = box(document.getElementById('toolbars'));
      if (inside(top, 8) && y <= top.bottom + 8) return 'top';
      const bot = box(document.getElementById('drawrow'));
      if (inside(bot, 8)) return 'bottom';
      const left = box(document.getElementById('tb-dock-left'));
      if (left && x <= left.right + 8 && y >= left.top && y <= left.bottom) return 'left';
      const right = box(document.getElementById('tb-dock-right'));
      if (right && x >= right.left - 8 && y >= right.top && y <= right.bottom) return 'right';
      const main = document.querySelector('.main');
      if (!main) return null;
      const m = main.getBoundingClientRect();
      const cmd = document.querySelector('.cmdbars');
      if (cmd) {
        const c = cmd.getBoundingClientRect();
        if (y >= c.top && y <= c.bottom + pad && x >= c.left && x <= c.right) return 'top';
      }
      const sb = document.getElementById('statusbar');
      const edge = sb && !sb.hidden ? sb.getBoundingClientRect().top : m.bottom;
      if (y >= edge - pad && y <= edge + 6 && x >= m.left && x <= m.right) return 'bottom';
      if (x <= m.left + pad && y >= m.top && y <= m.bottom) return 'left';
      if (x >= m.right - pad && y >= m.top && y <= m.bottom) return 'right';
      return null;
    }
    function slotAt(dock, x, y, tb) {
      const host = hostOf(dock);
      const side = dock === 'left' || dock === 'right';
      const all = host ? bandsOf(host) : [];
      const vis = all.map((b, i) => ({ b, i })).filter((x) => !x.b.hidden && x.b.offsetWidth);
      let row = 0, order = 0;
      if (!side && vis.length) {
        row = vis[vis.length - 1].i + 1;
        for (const x of vis) {
          const r = x.b.getBoundingClientRect();
          if (y < r.top + r.height * 0.75) { row = x.i; break; }
        }
      }
      const band = all[row];
      if (band) {
        const kids = [...band.children].filter((c) => c.classList.contains('toolbar') && c !== tb && !c.hidden);
        order = kids.length;
        for (let i = 0; i < kids.length; i++) {
          const r = kids[i].getBoundingClientRect();
          const mid = side ? r.top + r.height / 2 : r.left + r.width / 2;
          if ((side ? y : x) < mid) { order = i; break; }
        }
      }
      return { dock: side ? dock : dock, row: side ? 0 : row, order };
    }
    function toggleFloat(tb) {
      const id = idOf(tb);
      const cur = spotOf(id);
      if (cur.dock === 'float') commit(id, { dock: (cur.back && cur.back.dock) || defaults[id].dock, row: cur.back ? cur.back.row || 0 : defaults[id].row, order: cur.back ? cur.back.order || 0 : defaults[id].order });
      else {
        const r = tb.getBoundingClientRect();
        commit(id, { dock: 'float', x: Math.max(8, r.left), y: Math.max(8, r.bottom + 2), back: { dock: cur.dock, row: cur.row || 0, order: cur.order || 0 } });
      }
    }
    let ghost = null;
    function clearGhost() { if (ghost) { ghost.remove(); ghost = null; } }
    function ghostBox(left, top, w, ht, dock) {
      if (!ghost) { ghost = h('div', { class: 'tb-ghost' }); document.body.appendChild(ghost); }
      ghost.classList.toggle('dock', !!dock);
      ghost.style.left = left + 'px';
      ghost.style.top = top + 'px';
      ghost.style.width = Math.max(11, Math.round(w)) + 'px';
      ghost.style.height = Math.max(11, Math.round(ht)) + 'px';
    }
    /* the dotted frame sits on the window. Near an edge it becomes the solid dock frame; the window stays on screen. */
    function showGhost(tb, x, y) {
      const dock = zoneAt(x, y);
      const win = tb._toolwin && tb._toolwin.el && tb._toolwin.el.isConnected ? tb._toolwin.el : null;
      if (!dock) {
        if (!win) return;
        const r = win.getBoundingClientRect();
        ghostBox(r.left, r.top, r.width, r.height, false);
        return;
      }
      const horizontal = dock === 'top' || dock === 'bottom';
      const wr = win ? win.getBoundingClientRect() : null;
      const w = horizontal ? (wr ? wr.width : Math.min(tb.offsetWidth || 160, 520)) : 168;
      const ht = horizontal ? 30 : Math.min(200, Math.max(88, Math.round((tb.offsetWidth || 240) / 4)));
      const host = hostOf(dock);
      const r = host && !host.hidden ? host.getBoundingClientRect() : null;
      const main = document.querySelector('.main');
      const mr = main ? main.getBoundingClientRect() : { left: 0, right: window.innerWidth, top: 0, bottom: window.innerHeight };
      let left = x, top = y;
      if (dock === 'top') {
        const cmd = document.querySelector('.cmdbars').getBoundingClientRect();
        top = r ? r.top : cmd.bottom;
        left = Math.max(4, Math.min(x - w / 2, (r ? r.right : cmd.right) - w));
      } else if (dock === 'bottom') {
        const sb = document.getElementById('statusbar');
        const edge = sb ? sb.getBoundingClientRect().top : mr.bottom;
        top = r ? r.top : edge - ht;
        left = Math.max(mr.left, Math.min(x - w / 2, mr.right - w));
      } else if (dock === 'left') {
        left = r ? r.left : mr.left;
        top = Math.max(mr.top, Math.min(y - 12, mr.bottom - ht));
      } else {
        left = r ? r.right - w : mr.right - w;
        top = Math.max(mr.top, Math.min(y - 12, mr.bottom - ht));
      }
      ghostBox(left, top, w, ht, true);
    }
    function drop(tb, x, y, ox, oy) {
      const id = idOf(tb);
      const dock = zoneAt(x, y);
      if (!dock) {
        const cur = spotOf(id);
        const back = cur.dock === 'float' ? cur.back : { dock: cur.dock, row: cur.row || 0, order: cur.order || 0 };
        commit(id, { dock: 'float', x: x + ox, y: y + oy, back });
        return;
      }
      commit(id, slotAt(dock, x, y, tb));
    }
    if (typeof document !== 'undefined') {
      document.addEventListener('pointerdown', (e) => {
        if (!on() || e.button !== 0 || dragOn) return;
        const grip = e.target.closest && e.target.closest('.tb-grip');
        const title = !grip && e.target.closest && e.target.closest('.dlg.toolbox > .dlg-title');
        if (title && e.target.closest('button')) return;
        const tb = grip ? grip.closest('.toolbar') : title ? title.parentNode.querySelector('.toolbar') : null;
        if (!tb || !defaults[idOf(tb)] && !captured) return;
        if (!defaults[idOf(tb)]) return;
        if (e.detail >= 2) { e.preventDefault(); toggleFloat(tb); return; }
        const shell = title ? title.parentNode : tb;
        const rect = shell.getBoundingClientRect();
        let ox = rect.left - e.clientX, oy = rect.top - e.clientY;
        e.preventDefault();
        const follow = (x, y) => {
          let win = tb._toolwin;
          if (!win || !win.el || !win.el.isConnected) {
            placeFloat(tb, { x: x + ox, y: y + oy });
            win = tb._toolwin;
            if (!win || !win.el) return;
            if (!title) {
              const th = win.el.querySelector('.dlg-title');
              if (th) win.el.style.top = (parseFloat(win.el.style.top) - th.offsetHeight) + 'px';
            }
            const r = win.el.getBoundingClientRect();
            ox = r.left - x;
            oy = r.top - y;
            return;
          }
          win.el.style.left = (x + ox) + 'px';
          win.el.style.top = (y + oy) + 'px';
          if (win.overlay) win.overlay.hidden = false;
        };
        ui.drag(e, {
          cursor: 'default',
          capture: document.documentElement,
          move(ev) { dragOn = true; follow(ev.clientX, ev.clientY); showGhost(tb, ev.clientX, ev.clientY); },
          up(ev, d) { dragOn = false; clearGhost(); if (d.moved) drop(tb, ev.clientX, ev.clientY, ox, oy); else sync(); },
          cancel() { dragOn = false; clearGhost(); sync(); },
        });
      }, true);
      window.addEventListener('resize', () => { if (on() && captured) layout(); });
      if (L.onLook) L.onLook(() => sync());
    }
    return { sync, reset };
  })();

  /* ---------- dialogs ---------- */
  /* dlgStack: modal dialogs only, topmost last (keyboard). winStack: every open window, frontmost last. */
  const dlgStack = [];
  const winStack = [];
  let dlgBackdrop = null;
  ui.dialogOpen = () => dlgStack.length > 0;
  /** Windows 7, Ribbon 2010, Ribbon Golden Gate, Paper and Paper 2016 window show and close. Other looks, and reduced motion, stay instant. */
  function win7WindowMotion() {
    const look = document.documentElement && document.documentElement.dataset.look;
    if (look !== 'win7' && look !== 'ribbon2010' && look !== 'goldengate' && look !== 'goldengatedark' && look !== 'vscodegg' && look !== 'paper' && look !== 'paper2016') return false;
    try { return !window.matchMedia('(prefers-reduced-motion: reduce)').matches; } catch (e) { return true; }
  }
  const windowMotionDone = (e, name) => e.animationName === 'win7-' + name || e.animationName === 'goldengate-' + name || e.animationName === 'goldengatedark-' + name || e.animationName === 'vscodegg-' + name;
  /** drop windows a test removed from the page, then stack the rest and show the veil behind them when a modal is open */
  function layoutWindows() {
    for (let i = winStack.length - 1; i >= 0; i--) if (!winStack[i].overlay.isConnected) {
      const gone = winStack.splice(i, 1)[0];
      const k = dlgStack.indexOf(gone);
      if (k >= 0) dlgStack.splice(k, 1);
    }
    let topModal = -1;
    for (let i = winStack.length - 1; i >= 0; i--) if (!winStack[i].modeless) { topModal = i; break; }
    if (topModal >= 0) {
      if (!dlgBackdrop) dlgBackdrop = h('div', { class: 'dlg-overlay dlg-backdrop' });
      const first = winStack[0].overlay;
      if (dlgBackdrop.nextSibling !== first) document.body.insertBefore(dlgBackdrop, first);
    } else if (dlgBackdrop) { dlgBackdrop.remove(); dlgBackdrop = null; }
    winStack.forEach((w, i) => {
      w.overlay.style.zIndex = String(2000 + i);
      /* a modal in front keeps the windows behind it from taking clicks; they stay on screen */
      w.dlg.style.pointerEvents = topModal >= 0 && i < topModal ? 'none' : '';
    });
    if (flashing && flashing !== (topModal >= 0 ? winStack[topModal] : null)) stopFlash();
    syncActive();
    ui.placeMinibar();
  }
  /** the formatting bar sits one above the front window. With no window open that is 2000, under the menu floor.
   *  A modal hides the bar (it is not a layer above the dialog). */
  ui.placeMinibar = function () {
    const bar = document.querySelector('.minibar');
    if (!bar || bar.hidden) return;
    if (dlgStack.length) { if (L.minibar) L.minibar.hide(); return; }
    let z = 1999;
    for (const w of winStack) {
      const n = parseInt(w.overlay.style.zIndex, 10);
      if (n > z) z = n;
    }
    bar.style.zIndex = String(z + 1);
  };
  /** the front modal is active; otherwise the modeless window that holds the keyboard. prefer: a window just clicked or opened.
   *  a modal that is flashing owns its own active/inactive class until the flash ends */
  function syncActive(prefer) {
    let topModal = null;
    for (let i = winStack.length - 1; i >= 0; i--) if (!winStack[i].modeless) { topModal = winStack[i]; break; }
    let active = topModal;
    if (!active) {
      const ae = document.activeElement;
      if (prefer && winStack.indexOf(prefer) >= 0 && (!ae || prefer.dlg.contains(ae) || ae === document.body || ae === document.documentElement)) active = prefer;
      else for (let i = winStack.length - 1; i >= 0; i--) if (ae && winStack[i].dlg.contains(ae)) { active = winStack[i]; break; }
    }
    for (const w of winStack) if (w !== flashing) w.dlg.classList.toggle('inactive', w !== active);
  }
  /* a click outside the front modal flashes it between the active and inactive looks, then leaves it active */
  let flashing = null;
  let flashTimer = 0;
  function stopFlash() {
    if (flashTimer) clearInterval(flashTimer);
    flashTimer = 0;
    const was = flashing;
    flashing = null;
    if (was && was.dlg.isConnected) was.dlg.classList.remove('inactive');
  }
  function flashModal(entry) {
    if (!entry || !entry.dlg.isConnected) return;
    stopFlash();
    flashing = entry;
    let step = 0;
    const tick = () => {
      if (!entry.dlg.isConnected) { stopFlash(); syncActive(); return; }
      step++;
      entry.dlg.classList.toggle('inactive', step % 2 === 1);
      if (step >= 6) { stopFlash(); syncActive(); }
    };
    tick();
    flashTimer = setInterval(tick, 80);
  }
  /** bring a window to the front. a modal in front of it stays in front */
  function raiseWindow(entry) {
    const i = winStack.indexOf(entry);
    if (i < 0) return;
    for (let k = i + 1; k < winStack.length; k++) if (!winStack[k].modeless) return;
    if (i !== winStack.length - 1) { winStack.splice(i, 1); winStack.push(entry); document.body.appendChild(entry.overlay); }
    layoutWindows();
  }
  function dialogPlaceOn(opts) {
    const v = opts && opts.restoreWindowSizeAndPosition;
    return v === 1 || v === true || v === '1';
  }

  function dialogPlaceName(opts) {
    if (opts && opts.windowName != null && String(opts.windowName) !== '') return String(opts.windowName);
    return opts && opts.title ? String(opts.title) : '';
  }

  function loadDialogPlace(name) {
    if (!name || !L.store) return null;
    const all = L.store.get('dlgPlace', {});
    const raw = all && all[name];
    if (!raw || typeof raw !== 'object') return null;
    const left = +raw.left, top = +raw.top, width = +raw.width, height = +raw.height;
    if (![left, top, width, height].every(Number.isFinite) || width < 40 || height < 40) return null;
    return { left: left, top: top, width: width, height: height };
  }

  function saveDialogPlace(name, dlg) {
    if (!name || !L.store || !dlg) return;
    const all = L.store.get('dlgPlace', {});
    const map = all && typeof all === 'object' ? all : {};
    map[name] = { left: dlg.offsetLeft, top: dlg.offsetTop, width: dlg.offsetWidth, height: dlg.offsetHeight };
    L.store.set('dlgPlace', map);
  }

  function applyDialogPlace(dlg, opts) {
    const place = loadDialogPlace(dialogPlaceName(opts));
    if (!place) return false;
    const edge = 4;
    const minW = Number(opts.minWidth) > 0 ? Number(opts.minWidth) : 160;
    const minH = Number(opts.minHeight) > 0 ? Number(opts.minHeight) : 80;
    const maxW = Math.max(minW, window.innerWidth - edge * 2);
    const maxH = Math.max(minH, window.innerHeight - edge * 2);
    const width = L.clamp(place.width, Math.min(minW, maxW), maxW);
    const height = L.clamp(place.height, Math.min(minH, maxH), maxH);
    dlg.style.width = width + 'px';
    dlg.style.height = height + 'px';
    dlg.style.left = L.clamp(place.left, -width + 60, Math.max(-width + 60, window.innerWidth - 60)) + 'px';
    dlg.style.top = L.clamp(place.top, 0, Math.max(0, window.innerHeight - 30)) + 'px';
    return true;
  }

  /* Dragging an edge or a corner. The title stays the title: no extra button is added for this. */
  function attachDialogResize(dlg, opts) {
    const minW = Number(opts.minWidth) > 0 ? Number(opts.minWidth) : 160;
    const minH = Number(opts.minHeight) > 0 ? Number(opts.minHeight) : 80;
    let frame = 0;
    const fire = () => {
      frame = 0;
      if (typeof opts.onResize === 'function') opts.onResize(dlg);
    };
    const edges = ['n', 's', 'e', 'w', 'ne', 'nw', 'se', 'sw'];
    for (let i = 0; i < edges.length; i++) {
      const edge = edges[i];
      const handle = h('div', { class: 'dlg-resize dlg-resize-' + edge });
      handle.addEventListener('pointerdown', (e) => {
        if (e.button !== 0) return;
        e.preventDefault();
        const left0 = dlg.offsetLeft, top0 = dlg.offsetTop, w0 = dlg.offsetWidth, h0 = dlg.offsetHeight;
        const x0 = e.clientX, y0 = e.clientY;
        const west = edge.indexOf('w') >= 0, east = edge.indexOf('e') >= 0;
        const north = edge.indexOf('n') >= 0, south = edge.indexOf('s') >= 0;
        const move = (ev) => {
          const edgeX = 4, edgeY = 4;
          const limitW = Math.max(120, window.innerWidth - edgeX * 2);
          const limitH = Math.max(64, window.innerHeight - edgeY * 2);
          const floorW = Math.min(minW, limitW), floorH = Math.min(minH, limitH);
          const right = left0 + w0, bottom = top0 + h0;
          let w = w0, h = h0;
          if (east) w = w0 + ev.clientX - x0;
          if (west) w = w0 - (ev.clientX - x0);
          if (south) h = h0 + ev.clientY - y0;
          if (north) h = h0 - (ev.clientY - y0);
          w = L.clamp(w, floorW, limitW);
          h = L.clamp(h, floorH, limitH);
          let left = west ? right - w : left0;
          let top = north ? bottom - h : top0;
          if (left < edgeX) { left = edgeX; if (west) w = L.clamp(right - left, floorW, limitW); }
          if (top < edgeY) { top = edgeY; if (north) h = L.clamp(bottom - top, floorH, limitH); }
          if (left + w > window.innerWidth - edgeX) w = Math.max(floorW, window.innerWidth - edgeX - left);
          if (top + h > window.innerHeight - edgeY) h = Math.max(floorH, window.innerHeight - edgeY - top);
          left = L.clamp(left, edgeX, Math.max(edgeX, window.innerWidth - w - edgeX));
          top = L.clamp(top, edgeY, Math.max(edgeY, window.innerHeight - h - edgeY));
          dlg.style.width = w + 'px';
          dlg.style.height = h + 'px';
          dlg.style.left = left + 'px';
          dlg.style.top = top + 'px';
          if (!frame) frame = requestAnimationFrame(fire);
        };
        const up = () => {
          handle.removeEventListener('pointermove', move);
          handle.removeEventListener('pointerup', up);
          handle.removeEventListener('pointercancel', up);
          if (frame) { cancelAnimationFrame(frame); frame = 0; }
          fire();
        };
        handle.setPointerCapture(e.pointerId);
        handle.addEventListener('pointermove', move);
        handle.addEventListener('pointerup', up);
        handle.addEventListener('pointercancel', up);
      });
      dlg.appendChild(handle);
    }
  }

  /**
   * opts: {title, body, width, buttons:[{label, primary, cancel, onClick(close, button)→false to keep open}], onClose, help, modeless, canClose, brand, isToolBoxWindow, resizeable, minWidth, minHeight, onResize, windowName, restoreWindowSizeAndPosition}
   * a new window opens in front of the ones already open and leaves them where they are.
   * canClose false hides the title's close button, and Escape does not close the window (it does not press Cancel either).
   * resizeable defaults to false. True lets an edge or a corner change the size. The title bar does not gain a button for it.
   * restoreWindowSizeAndPosition defaults to 0. 1 keeps this window's left, top, width and height in this app, and the next open uses them. windowName is the stored key; without it the title is the key.
   * brand sets data-brand to the app, so the body and foot use that app's --brand-bg. The frame is not restyled.
   * isToolBoxWindow defaults to 0. 1 is a modeless window whose close button is square; the frame is otherwise the same window, inactive the same way.
   */
  ui.dialog = function (opts) {
    opts.isToolBoxWindow = opts.isToolBoxWindow ? 1 : 0;
    if (opts.isToolBoxWindow) opts.modeless = true;
    ui.closeMenus();
    hideTip();
    const prevFocus = document.activeElement;
    /* the selection in text being edited, for text the app doesn't give back itself (Oratrice's speaker notes) */
    const s0 = window.getSelection();
    const prevRange = prevFocus && prevFocus.isContentEditable && s0.rangeCount && prevFocus.contains(s0.anchorNode) ? s0.getRangeAt(0).cloneRange() : null;
    const overlay = h('div', { class: 'dlg-overlay' + (opts.modeless ? ' modeless' : '') + (opts.isToolBoxWindow ? ' toolbox' : '') });
    const dlg = h('div', { class: 'dlg' + (opts.brand ? ' brand' : '') + (opts.isToolBoxWindow ? ' toolbox' : '') + (opts.resizeable ? ' dlg-resizeable' : ''), role: 'dialog', tabindex: '-1', 'aria-modal': opts.modeless ? 'false' : 'true', 'aria-label': opts.title, dataset: opts.brand && L.APP_ID ? { brand: L.APP_ID } : null });
    dlg.style.width = (opts.isToolBoxWindow ? Math.min(opts.width || 420, window.innerWidth - 16) : Math.min(opts.width || 380, window.innerWidth - 16)) + 'px';
    const titleBar = h('div', { class: 'dlg-title' }, h('span', { class: 'dlg-ttl', text: opts.title }),
      opts.help ? h('button', { class: 'dlg-x dlg-help', type: 'button', 'aria-label': 'Help', text: '?', onclick: opts.help }) : null,
      opts.canClose === false ? null : h('button', { class: 'dlg-x', type: 'button', 'aria-label': 'Close', html: '<svg width="9" height="9" viewBox="0 0 9 9"><path d="M1 1l7 7M8 1L1 8" stroke="currentColor" stroke-width="1.5"/></svg>', onclick: () => close(null) }));
    const body = h('div', { class: 'dlg-body' });
    if (opts.body) body.appendChild(opts.body);
    const foot = h('div', { class: 'dlg-foot' });
    let result;
    const done = new Promise((res) => { result = res; });
    const buttonList = opts.isToolBoxWindow ? [] : (opts.buttons || [{ label: 'OK', primary: true }]);
    const btns = buttonList.map((b, i) => {
      const el = h('button', { class: 'btn' + (b.primary ? ' primary' : ''), type: 'button', html: ampHTML(b.label) });
      el.addEventListener('click', () => {
        if (b.onClick) { const r = b.onClick(close, el); if (r === false) return; }
        close(i);
      });
      foot.appendChild(el);
      return el;
    });
    dlg.append(titleBar, body, foot);
    overlay.appendChild(dlg);
    document.body.appendChild(overlay);
    /* center; each further window steps down so the one behind it stays visible */
    const r = dlg.getBoundingClientRect();
    const step = Math.min(winStack.length, 8) * 28;
    dlg.style.left = L.clamp((window.innerWidth - r.width) / 2 + step, 4, Math.max(4, window.innerWidth - r.width - 4)) + 'px';
    dlg.style.top = L.clamp((window.innerHeight - r.height) / 2.4 + step, 4, Math.max(4, window.innerHeight - 40)) + 'px';
    /* measured before the class, so the opening scale does not shift the centred box */
    if (win7WindowMotion()) {
      dlg.classList.add('is-opening');
      const endOpen = (e) => { if (e.target !== dlg || !windowMotionDone(e, 'open')) return; dlg.classList.remove('is-opening'); dlg.removeEventListener('animationend', endOpen); };
      dlg.addEventListener('animationend', endOpen);
      setTimeout(() => dlg.classList.remove('is-opening'), 400);
    }
    const entry = { overlay, dlg, modeless: !!opts.modeless, isToolBoxWindow: opts.isToolBoxWindow, keys: null };
    dlg.addEventListener('pointerdown', (e) => {
      raiseWindow(entry);
      /* a press on the frame, not a control, still counts as the window taking the keyboard */
      if (!e.target.closest('input, select, textarea, button, a[href], [contenteditable="true"]')) { try { dlg.focus({ preventScroll: true }); } catch (err) { /* ignore */ } }
      syncActive(entry);
    });
    /* drag */
    titleBar.addEventListener('pointerdown', (e) => {
      if (e.target.closest('button') || opts.isToolBoxWindow) return;
      const sx = e.clientX, sy = e.clientY, ox = dlg.offsetLeft, oy = dlg.offsetTop;
      titleBar.setPointerCapture(e.pointerId);
      const mv = (ev) => { const w = dlg.offsetWidth; dlg.style.left = L.clamp(ox + ev.clientX - sx, -w + 60, window.innerWidth - 60) + 'px'; dlg.style.top = L.clamp(oy + ev.clientY - sy, 0, window.innerHeight - 30) + 'px'; };
      const up = () => { titleBar.removeEventListener('pointermove', mv); titleBar.removeEventListener('pointerup', up); };
      titleBar.addEventListener('pointermove', mv);
      titleBar.addEventListener('pointerup', up);
    });
    if (opts.resizeable) attachDialogResize(dlg, opts);
    let rememberPlace = null;
    if (dialogPlaceOn(opts)) {
      rememberPlace = () => saveDialogPlace(dialogPlaceName(opts), dlg);
      window.addEventListener('beforeunload', rememberPlace);
    }
    winStack.push(entry);
    if (opts.modeless) dlg.addEventListener('keydown', (e) => { entry.keys(e); e.stopPropagation(); });   // keys stay in a modeless dialog
    else dlgStack.push(entry);
    layoutWindows();
    syncActive(opts.isToolBoxWindow ? null : entry);
    entry.keys = (e) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        /* canClose false: stay open. Escape must not press Cancel, which would close the window too. */
        if (opts.canClose === false) return true;
        const ci = buttonList.findIndex((b) => b.cancel || /^cancel$|^close$/i.test(plainAmp(b.label)));
        if (ci >= 0) btns[ci].click(); else close(null);
        return true;
      }
        if (e.key === 'Enter' && !e.shiftKey && !(e.target && (e.target.tagName === 'TEXTAREA' || e.target.tagName === 'BUTTON' || e.target.isContentEditable || e.target.tagName === 'SELECT'))) {
        const pi = buttonList.findIndex((b) => b.primary);
        if (pi >= 0) { e.preventDefault(); btns[pi].click(); return true; }
      }
      if (e.key === 'Tab') {
        const f = L.$$('input,select,textarea,button,[tabindex="0"]', dlg).filter((x) => !x.disabled && x.offsetParent !== null);
        if (!f.length) return false;
        const i = f.indexOf(document.activeElement);
        if (e.shiftKey && (i <= 0)) { e.preventDefault(); f[f.length - 1].focus(); return true; }
        if (!e.shiftKey && i === f.length - 1) { e.preventDefault(); f[0].focus(); return true; }
      }
      return false;
    };
    function close(i) {
      if (entry.closed) return;
      if (rememberPlace) {
        rememberPlace();
        window.removeEventListener('beforeunload', rememberPlace);
        rememberPlace = null;
      }
      entry.closed = true;
      const k = dlgStack.indexOf(entry);
      if (k >= 0) dlgStack.splice(k, 1);
      const w = winStack.indexOf(entry);
      if (w >= 0) winStack.splice(w, 1);
      const drop = () => { if (overlay.isConnected) overlay.remove(); };
      if (win7WindowMotion()) {
        dlg.classList.remove('is-opening');
        dlg.classList.add('is-closing');
        let dropped = false;
        const finish = () => { if (dropped) return; dropped = true; drop(); };
        dlg.addEventListener('animationend', (e) => { if (e.target === dlg && windowMotionDone(e, 'close')) finish(); });
        setTimeout(finish, 400);
      } else drop();
      layoutWindows();
      if (opts.onClose) opts.onClose(i);
      result(i);
      /* the keyboard goes back to what had it: a box, a pane's list, the dialog under this one, the grid, text being
         edited with the selection it had (the page's own moved into the dialog); the document's own text the app gives
         back itself, with its selection (ui.hooks.ownsText); nothing (<body>, a menu gone) goes to the document */
      const live = prevFocus && prevFocus.isConnected && prevFocus !== document.body;
      if (live && !(prevFocus.isContentEditable && ui.hooks.ownsText(prevFocus))) {
        try { prevFocus.focus({ preventScroll: true }); } catch (e) { /* ignore */ }
        if (prevRange && prevRange.startContainer.isConnected && prevRange.endContainer.isConnected) { const s = window.getSelection(); s.removeAllRanges(); s.addRange(prevRange); }
      } else ui.hooks.refocus();
      syncActive();
    }
    setTimeout(() => {
      if (opts.isToolBoxWindow) return;
      const first = opts.focus ? (typeof opts.focus === 'string' ? L.$(opts.focus, dlg) : opts.focus) : L.$('input:not([type=checkbox]):not([type=radio]),select,textarea', body) || btns.find((b, i) => buttonList[i] && buttonList[i].primary) || btns[0];
      if (first) { first.focus(); if (first.select && first.tagName === 'INPUT') first.select(); }
    }, 0);
    if (opts.onOpen) opts.onOpen(dlg);
    if (dialogPlaceOn(opts) && applyDialogPlace(dlg, opts) && typeof opts.onResize === 'function') opts.onResize(dlg);
    return { el: dlg, body, overlay, close, done, buttons: btns };
  };
  ui.dialogKey = (e) => { const top = dlgStack[dlgStack.length - 1]; return top ? top.keys(e) || true : false; };
  if (typeof document !== 'undefined') {
    document.addEventListener('focusin', () => syncActive());
    document.addEventListener('pointerdown', (e) => {
      if (!dlgBackdrop || !dlgBackdrop.isConnected || e.target.closest('.menu')) return;
      let top = null;
      for (let i = winStack.length - 1; i >= 0; i--) if (!winStack[i].modeless) { top = winStack[i]; break; }
      if (!top) return;
      if (top.dlg.contains(e.target)) { if (flashing) { stopFlash(); syncActive(); } return; }
      flashModal(top);
    }, true);
  }

  /** Message box. Returns Promise<buttonIndex|null> */
  ui.msg = function (text, o) {
    o = o || {};
    const icons = {
      info: '<svg width="32" height="32" viewBox="0 0 32 32"><circle cx="16" cy="16" r="14" fill="#2c64c8" stroke="#163b84"/><text x="12.5" y="24" font-family="Georgia,serif" font-size="19" font-weight="bold" fill="#fff">i</text></svg>',
      warn: '<svg width="32" height="32" viewBox="0 0 32 32"><path d="M16 2l14.5 26h-29z" fill="#f6cf3b" stroke="#a07c08"/><text x="13" y="25" font-family="Arial" font-size="17" font-weight="bold" fill="#222">!</text></svg>',
      question: '<svg width="32" height="32" viewBox="0 0 32 32"><circle cx="16" cy="16" r="14" fill="#2c64c8" stroke="#163b84"/><text x="10.5" y="23.5" font-family="Arial" font-size="18" font-weight="bold" fill="#fff">?</text></svg>',
      error: '<svg width="32" height="32" viewBox="0 0 32 32"><circle cx="16" cy="16" r="14" fill="#d33a2c" stroke="#8a1a10"/><path d="M10 10l12 12M22 10L10 22" stroke="#fff" stroke-width="3.4"/></svg>',
    };
    const body = h('div', { class: 'msgbox' }, h('div', { class: 'msg-ico', html: icons[o.icon || 'info'] }), h('div', { class: 'msg-text' }));
    const t = body.querySelector('.msg-text');
    String(L.t ? L.t(text) : text).split('\n').forEach((line, i) => { if (i) t.appendChild(h('br')); t.appendChild(document.createTextNode(line)); });
    const labels = o.buttons || ['OK'];
    tell('vo-message', { icon: o.icon || 'info' });
    const d = ui.dialog({ title: o.title || L.APP, body, width: o.width || 380, buttons: labels.map((l, i) => ({ label: l, primary: i === (o.def || 0), cancel: /cancel|^no$/i.test(plainAmp(l)) })) });
    return d.done;
  };
  /** The about copyright rolls up once, from below the area until it has left the top. Reduced motion shows it in place. */
  function rollCredits(view, track) {
    let anim = null;
    let done = false;
    const stopAnim = () => { if (anim) { anim.cancel(); anim = null; } view.removeEventListener('wheel', onWheel); };
    const hold = (ty) => {
      if (done) return;
      done = true;
      stopAnim();
      track.style.transform = 'translateY(' + ty + 'px)';
    };
    const onWheel = (e) => {
      let ty = 0;
      try {
        const m = getComputedStyle(track).transform;
        if (m && m !== 'none') ty = new DOMMatrix(m).m42;
      } catch (err) { /* stop where it is */ }
      if (done) return;
      done = true;
      stopAnim();
      track.style.transform = '';
      view.classList.add('is-done');
      view.scrollTop = Math.max(0, Math.round(-ty + (e.deltaY || 0)));
    };
    const reduce = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
    const viewH = view.clientHeight;
    const trackH = track.scrollHeight;
    if (reduce || !viewH || !trackH || typeof track.animate !== 'function') {
      done = true;
      view.classList.add('is-done');
      return { stop() {} };
    }
    const start = viewH;
    const end = -trackH;
    const duration = Math.max(6000, Math.round((start - end) / 40 * 1000));
    anim = track.animate(
      [{ transform: 'translateY(' + start + 'px)' }, { transform: 'translateY(' + end + 'px)' }],
      { duration, easing: 'linear', fill: 'forwards' });
    anim.onfinish = () => hold(end);
    view.addEventListener('wheel', onWheel, { passive: true });
    return { stop() { if (!done) hold(end); } };
  }
  /** Help ▸ About, the same in every app. The frame stays. Inside: the name, then the version, the app's
   *  paragraphs, the open-source licenses, the license and font notes, and the trademark note, rolling up once.
   *  The button is OK. o: { name, text: [paragraphs], office } */
  ui.about = function (o) {
    const credits = h('div', { class: 'brand-marquee' },
      ...o.text.map((para) => h('p', { text: para })),
      h('p', null, ...t('Made possible by open source software: {0}.').split('{0}').flatMap((part, i) => (i ? [h('a', { href: root.VO ? root.VO.LICENSES : '../oss-licenses.html', target: '_blank', rel: 'noopener',
        class: 'about-oss', text: 'Open source licenses', onclick: (e) => { if (root.VO) { e.preventDefault(); root.VO.showLicenses(); } } }), part] : [part]))),
      h('p', null, ...t('This project is a derivative of {0}, which is licensed under the Apache License 2.0.').split('{0}').flatMap((part, i) => (i ? [h('a', { href: 'https://github.com/tcchen2026/VibeOffice', target: '_blank', rel: 'noopener',
        class: 'about-oss', text: 'VibeOffice' }), part] : [part]))),
      h('p', { text: 'Open source under the Apache License 2.0, provided as is, without warranty. Melina Office is an independent project. Product names in its documentation are trademarks of their owners; Melina Office is not affiliated with them.' }),
      h('p', null, ...t('Nikki (暖暖), an optional assistant, is shown from {0}. Her name and likeness are the copyright of Papergames (叠纸), also known as Infold. Melina Office is not affiliated with them. That sprite is not covered by the Apache License and is not licensed for commercial use.').split('{0}').flatMap((part, i) => (i ? [h('a', { href: 'https://codex-pets.net/#/pets/nikki-bloom-poetry', target: '_blank', rel: 'noopener', class: 'about-oss', text: 'a sprite published at codex-pets.net' }), part] : [part]))),
      h('p', { text: 'Font names are used solely to identify the fonts specified in documents and to maintain interoperability with Office Open XML files and the applications that create them; such names may be trademarks of their respective owners. Melina Office renders text with the fonts installed on your computer. Where a specified font is not installed, Melina Office substitutes an open-source font supplied with it or another font available on your computer, and the document may therefore appear differently than in other applications. Melina Office does not include or distribute any commercial font software.' }),
      h('p', { text: t('Microsoft, {0} and Office are trademarks of Microsoft Corporation. {1} is an independent work and is not affiliated with Microsoft.', o.office, o.name) }));
    const view = h('div', { class: 'brand-credits-view' },
      h('div', { class: 'brand-fade brand-fade-top' }),
      h('div', { class: 'brand-fade brand-fade-bottom' }),
      credits);
    /* Not translated, half-width colons. public/version.json is written by tools/package_vsix.sh. */
    const versionLine = h('p', { class: 'brand-version', translate: 'no', text: ['Version', ': ', L.VERSION || 'dev'].join('') });
    try {
      const url = new URL('../version.json', document.baseURI);
      fetch(url.href).then((r) => (r.ok ? r.json() : null)).then((info) => {
        if (!info || !versionLine.isConnected || info.version == null) return;
        versionLine.textContent = 'Version: v' + info.version + ' (' + (info.commit_id || '') + '), Build Date: ' + (info.build_time || '');
      }).catch(() => {});
    } catch (e) { /* opened without a base URL */ }
    const body = h('div', { class: 'brand-about' },
      h('p', { class: 'brand-name', translate: 'no', text: o.name }),
      versionLine,
      view);
    let roll = null;
    return ui.dialog({
      title: t('About {0}', o.name), width: 500, brand: true, body,
      buttons: [{ label: 'OK', primary: true }],
      onOpen() { requestAnimationFrame(() => { if (view.isConnected) roll = rollCredits(view, credits); }); },
      onClose() { if (roll) roll.stop(); },
    });
  };
  ui.prompt = function (label, value, title) {
    const inp = h('input', { type: 'text', id: 'prompt-in', value: value || '', style: 'width:100%' });
    let val = null;
    const d = ui.dialog({ title: title || L.APP, body: h('div', { class: 'col' }, h('label', { for: 'prompt-in', text: label }), inp), buttons: [{ label: 'OK', primary: true, onClick: () => { val = inp.value; } }, { label: 'Cancel' }] });
    return d.done.then(() => val);
  };

  /** Check a prepared Office save before handing it to the download host.
   * Returns entries to acknowledge after a successful save, or null on cancel.
   * Snapshot writers never call this UI or acknowledge a loss.
   */
  ui.compatibility = async function (doc, entries) {
    const pending = L.opc.noticeLosses(doc, entries);
    if (!pending.length) return [];
    // One line per message, with every place it applies to
    const groups = new Map();
    for (const e of pending) {
      if (!groups.has(e.what)) groups.set(e.what, []);
      if (e.place && !groups.get(e.what).includes(e.place)) groups.get(e.what).push(e.place);
    }
    const list = h('ul', { style: 'max-height:240px;overflow:auto;margin:8px 0;padding-left:22px' });
    for (const [what, places] of groups) list.appendChild(h('li', { style: 'margin-bottom:8px' },
      h('div', { text: what }), places.length ? h('div', { class: 'hint', text: places.length > 4 ? t('{0} and {1} more', places.slice(0, 4).join(', '), places.length - 4) : places.join(', ') }) : null));
    const body = h('div', { class: 'col' },
      h('div', { text: 'Some of this file can\'t be saved exactly as it was:' }), list,
      h('div', { text: 'Click Continue to save anyway, or Cancel to go back to your document without saving.' }));
    const result = await ui.dialog({ title: 'Compatibility Checker', width: 540, body,
      buttons: [{ label: '&Continue', primary: true }, { label: 'Cancel', cancel: true }] }).done;
    return result === 0 ? pending : null;
  };

  /** the "Password" box for opening a protected file; resolves to the password or null. wrong: after a wrong one */
  ui.password = function (fileName, wrong) {
    const inp = h('input', { type: 'password', id: 'pw-in', value: '', style: 'width:100%', autocomplete: 'off' });
    let val = null;
    const d = ui.dialog({ title: 'Password', width: 340, body: h('div', { class: 'col' }, h('div', { text: wrong ? 'The password is incorrect. Try again.' : t('\'{0}\' is protected.', fileName) }), h('label', { for: 'pw-in', text: 'Enter password to open file' }), inp), buttons: [{ label: 'OK', primary: true, onClick: () => { val = inp.value; } }, { label: 'Cancel' }] });
    setTimeout(() => inp.focus(), 0);
    return d.done.then(() => val);
  };

  /* ---------- form controls ---------- */
  let ctlId = 0;
  ui.button = (label, onClick, o) => { const b = h('button', Object.assign({ class: 'btn', type: 'button', html: ampHTML(label) }, o || {})); b.addEventListener('click', onClick); return b; };
  ui.check = (label, checked, onChange, o) => {
    const id = (o && o.id) || 'ck' + (++ctlId);
    const inp = h('input', { type: 'checkbox', id });
    inp.checked = !!checked;
    inp.addEventListener('change', () => onChange && onChange(inp.checked));
    const wrap = h('label', { class: 'ck', for: id }, inp, h('span', { html: ampHTML(label) }));
    wrap.input = inp;
    return wrap;
  };
  ui.radio = (name, label, checked, onChange, o) => {
    const id = (o && o.id) || 'rd' + (++ctlId);
    const inp = h('input', { type: 'radio', id, name });
    inp.checked = !!checked;
    inp.addEventListener('change', () => { if (inp.checked && onChange) onChange(); });
    const wrap = h('label', { class: 'ck', for: id }, inp, h('span', { html: ampHTML(label) }));
    wrap.input = inp;
    return wrap;
  };
  /** options: [value, label] pairs, whose labels are UI text, or plain values shown as they are (font names, sizes).
      The string '-' is a separator, not a choice. o.translate: 'no' shows every label as it is */
  ui.select = (options, value, onChange, o) => {
    const sel = h('select', Object.assign({ id: (o && o.id) || 'sel' + (++ctlId) }, o || {}));
    let group = null;
    for (const op of options) {
      if (op === '-') { group = h('optgroup', { label: '\u2500'.repeat(8) }); sel.appendChild(group); continue; }
      const [v, l] = Array.isArray(op) ? op : [op, op];
      const opt = h('option', { value: v, text: l, translate: !Array.isArray(op) || (o && o.translate === 'no') ? 'no' : null });
      if (String(v) === String(value)) opt.selected = true;
      (group || sel).appendChild(opt);
    }
    sel.addEventListener('change', () => onChange && onChange(sel.value));
    return sel;
  };
  /** numeric spinner with unit, e.g. {value: 1.5, unit:'"', step:0.1, min, max} */
  ui.spin = (o) => {
    const id = o.id || 'sp' + (++ctlId);
    const fmt = (v) => (o.fmt ? o.fmt(v) : L.round(v, o.dec == null ? 2 : o.dec) + (o.unit || ''));
    const inp = h('input', { type: 'text', id, class: 'spin-in', value: fmt(o.value), inputmode: 'decimal', autocomplete: 'off' });
    const parse = () => { const v = parseFloat(String(inp.value).replace(/[^0-9.\-]/g, '')); return isNaN(v) ? o.value : v; };
    const set = (v) => { v = L.clamp(v, o.min == null ? -1e9 : o.min, o.max == null ? 1e9 : o.max); o.value = v; inp.value = fmt(v); if (o.onChange) o.onChange(v); };
    const up = h('button', { type: 'button', class: 'spin-b up', tabindex: '-1', 'aria-label': 'Increase', html: '<span></span>' });
    const dn = h('button', { type: 'button', class: 'spin-b dn', tabindex: '-1', 'aria-label': 'Decrease', html: '<span></span>' });
    up.addEventListener('click', () => set(L.round(parse() + (o.step || 1), 4)));
    dn.addEventListener('click', () => set(L.round(parse() - (o.step || 1), 4)));
    inp.addEventListener('change', () => set(parse()));
    inp.addEventListener('keydown', (e) => { if (e.key === 'ArrowUp') { e.preventDefault(); up.click(); } if (e.key === 'ArrowDown') { e.preventDefault(); dn.click(); } });
    const wrap = h('span', { class: 'spin' }, inp, h('span', { class: 'spin-bs' }, up, dn));
    wrap.input = inp;
    wrap.get = () => parse();
    wrap.set = (v) => { o.value = v; inp.value = fmt(v); };
    wrap.setDisabled = (d) => { inp.disabled = d; up.disabled = d; dn.disabled = d; };
    return wrap;
  };
  /** Tabs: [{label, body}] */
  ui.tabs = (tabs, active) => {
    const head = h('div', { class: 'tabs-head', role: 'tablist' });
    const bodies = h('div', { class: 'tabs-body' });
    const btns = tabs.map((t, i) => {
      const b = h('button', { type: 'button', class: 'tab', role: 'tab', html: ampHTML(t.label) });
      b.addEventListener('click', () => sel(i));
      head.appendChild(b);
      t.body.classList.add('tab-pane');
      bodies.appendChild(t.body);
      return b;
    });
    function sel(i) { btns.forEach((b, k) => { b.classList.toggle('on', k === i); b.setAttribute('aria-selected', String(k === i)); tabs[k].body.hidden = k !== i; }); if (tabs[i].onShow) tabs[i].onShow(); }
    sel(active || 0);
    const el = h('div', { class: 'tabs' }, head, bodies);
    el.select = sel;
    return el;
  };
  ui.field = (label, ctl, o) => {
    const id = ctl.id || (ctl.input && ctl.input.id) || (ctl.querySelector && ctl.querySelector('input,select') && ctl.querySelector('input,select').id);
    return h('div', { class: 'field' + (o && o.cls ? ' ' + o.cls : '') }, h('label', { for: id || null, html: ampHTML(label) }), ctl);
  };
  /** File ▸ Save As in the VS Code editor: a Location row under the file name. It starts as the open file's
   *  folder; Browse opens the editor's folder dialog. null in the browser and the desktop app, which keep
   *  their own note about where a download goes. { el, uri() }. */
  ui.saveLocation = function () {
    const host = root.VO && root.VO.host;
    if (!host || host.kind !== 'vscode' || !host.pickFolder) return null;
    let uri = '';
    let chosen = false;
    const input = h('input', { type: 'text', id: 'sa-location', readOnly: true, spellcheck: 'false', autocomplete: 'off', translate: 'no' });
    const show = (f) => {
      if (!f || chosen) return;
      uri = f.uri || '';
      input.value = f.label || '';
      input.title = input.value;
    };
    show(host.documentFolder && host.documentFolder());
    if (host.askFolder) host.askFolder().then(show).catch((e) => console.error('folder:', e));
    let busy = false;
    const browse = ui.button('&Browse...', () => {
      if (busy) return;
      busy = true;
      browse.disabled = true;
      host.pickFolder(uri).then((f) => {
        if (!f) return;
        chosen = true;
        uri = f.uri || '';
        input.value = f.label || '';
        input.title = input.value;
      }).catch((e) => console.error('folder:', e)).finally(() => { busy = false; browse.disabled = false; });
    });
    const row = h('div', { class: 'sa-loc' }, input, browse);
    return { el: ui.field('&Location:', row, { cls: 'wide' }), uri: () => uri };
  };
  ui.group = (title, ...kids) => h('fieldset', { class: 'grp' }, h('legend', { html: ampHTML(title) }), ...kids);
  /** Tools ▸ Options ▸ General ▸ Theme: the suite's look (common/looks.css). { el, apply() }; apply on OK
   *  changes every open window of the suite. null outside the suite (no common/suite.js). */
  /** Help ▸ Theme: the same looks as the Start Center, the current one marked. Golden Gate is the stored default
   *  outside a VS Code webview; there it is Golden Gate (Follow VS Code Color Theme).
   *  A look marked sep has a separator before it. */
  ui.themeMenu = () => ({
    label: '&Theme',
    sub: () => {
      if (!window.VO || !VO.look) return [];
      const items = [];
      for (const l of VO.look.list()) {
        if (l.sep) items.push('-');
        items.push({ label: l.name, radio: true, checked: () => VO.look.get() === l.id, run: () => VO.look.set(l.id) });
      }
      return items;
    },
  });
  /** Remember a UI language. In VS Code the extension keeps it and reloads this editor when the file
   *  has no unsaved changes. The language is part of the page HTML: assigning the same HTML does not
   *  load the editor again. A choice already stored, while this page still shows another language, reloads too. */
  ui.applyLanguage = (id) => {
    if (!window.VO || !VO.lang) return;
    const changed = VO.lang.get() !== id;
    if (changed) VO.lang.set(id);
    const inVs = VO.host && VO.host.kind === 'vscode' && VO.host.setLanguage;
    if (!inVs) { if (changed) ui.toast('The new language is used in the windows you open from now on.'); return; }
    const painted = document.documentElement.lang === 'qps-ploc' ? 'qps' : document.documentElement.lang;
    const want = VO.lang.resolve(id, navigator.languages || [navigator.language]);
    if (!changed && painted === want) return;
    const dirty = !!(VO.unsaved && VO.unsaved());
    VO.host.setLanguage(id, !dirty);
    if (dirty) ui.toast('The new language is used in the windows you open from now on.');
  };
  /** Help ▸ Language, and the same list on the ribbon's File menu: the suite's UI language, only in the
   *  VS Code editor (common/lang.js). Each language is named in itself. */
  ui.languageMenu = () => ({
    label: '&Language',
    hidden: () => !window.VO || !VO.host || VO.host.kind !== 'vscode' || !VO.lang,
    sub: () => (!window.VO || !VO.lang ? [] : VO.lang.list().map((l) => ({
      label: l.id === 'auto' ? 'Auto (as the browser)' : l.name,
      raw: l.id !== 'auto',
      radio: true,
      checked: () => VO.lang.get() === l.id,
      run: () => ui.applyLanguage(l.id),
    }))),
  });
  /** the suite's own choices for Tools ▸ Options: the look (Theme) and the UI's language (common/i18n.js) */
  ui.lookField = () => {
    if (!window.VO || !VO.look) return null;
    let pick = VO.look.get();
    const lookOpts = [];
    for (const l of VO.look.list()) { if (l.sep) lookOpts.push('-'); lookOpts.push([l.id, l.name]); }
    const sel = ui.select(lookOpts, pick, (v) => { pick = v; }, { id: 'op-look' });
    const el = h('div', { class: 'col' }, ui.field('Theme:', sel));
    let lang = VO.lang && VO.lang.get();
    /* each language named in itself */
    if (VO.lang) el.appendChild(ui.field('Language:', ui.select(VO.lang.list().map((l) => [l.id, l.id === 'auto' ? t('Auto (as the browser)') : l.name]), lang, (v) => { lang = v; }, { id: 'op-lang', translate: 'no' })));
    return {
      el,
      apply: () => {
        if (pick !== VO.look.get()) VO.look.set(pick);
        if (VO.lang && lang !== VO.lang.get()) ui.applyLanguage(lang);
      },
    };
  };

  /* ---------- colour pickers ---------- */
  const recent = [];
  ui.addRecentColor = (c) => { if (!c || c[0] !== '#') return; const i = recent.indexOf(c); if (i >= 0) recent.splice(i, 1); recent.unshift(c); recent.length = Math.min(recent.length, 8); };
  /**
   * Colour drop-down. o: {mode:'fill'|'line'|'font'|'plain', design, noneLabel, effects:bool, automatic:bool,
   *   scheme:bool (theme colours; on when the app provides ui.schemeColors), grid:bool (the standard colours too)}
   * onPick(value) where value is a colour ('#hex' or scheme ref), {none:true}, {auto:true}, {more:true}, {effects:true}
   */
  ui.colorMenu = function (anchor, o, onPick) {
    o = o || {};
    const design = o.design || (ui.currentDesign ? ui.currentDesign() : null);
    const scheme = o.scheme != null ? !!o.scheme : !!ui.schemeColors;
    const r = anchor.getBoundingClientRect ? anchor.getBoundingClientRect() : anchor;
    const panel = (close) => {
      const el = h('div', { class: 'color-panel' });
      if (o.mode === 'fill' || o.mode === 'line') el.appendChild(h('button', { class: 'cp-wide', type: 'button', text: o.mode === 'fill' ? 'No Fill' : 'No Line', onclick: () => { close(); onPick({ none: true }); } }));
      if (o.mode === 'font' || o.automatic) el.appendChild(h('button', { class: 'cp-wide', type: 'button', text: 'Automatic', onclick: () => { close(); onPick({ auto: true }); } }));
      el.appendChild(h('div', { class: 'cp-sep' }));
      if (scheme && ui.schemeColors) {
        const row = h('div', { class: 'cp-row' });
        for (const sc of ui.schemeColors(design)) {
          const b = h('button', { class: 'cp-sw', type: 'button', style: `background:${sc.hex}`, 'data-tip': sc.tip || null, 'aria-label': sc.tip || sc.ref });
          b.addEventListener('click', () => { close(); onPick(sc.ref); });
          row.appendChild(b);
        }
        el.appendChild(row);
      }
      if (recent.length) {
        el.appendChild(h('div', { class: 'cp-sep' }));
        const rr = h('div', { class: 'cp-row' });
        for (const c of recent) rr.appendChild(h('button', { class: 'cp-sw', type: 'button', style: `background:${c}`, 'aria-label': c, onclick: () => { close(); onPick(c); } }));
        el.appendChild(rr);
      }
      if (o.mode === 'plain' || o.grid || !scheme) {
        if (scheme) el.appendChild(h('div', { class: 'cp-sep' }));
        const g = h('div', { class: 'cp-grid' });
        for (const c of L.color.STANDARD) g.appendChild(h('button', { class: 'cp-sw', type: 'button', style: `background:${c}`, 'aria-label': c, 'data-tip': ui.colorName(c), onclick: () => { close(); ui.addRecentColor(c); onPick(c); } }));
        el.appendChild(g);
      }
      el.appendChild(h('div', { class: 'cp-sep' }));
      el.appendChild(h('button', { class: 'cp-wide', type: 'button', text: o.mode === 'fill' ? 'More Fill Colors...' : o.mode === 'line' ? 'More Line Colors...' : 'More Colors...', onclick: () => { close(); ui.moreColors(null, (c) => { if (c) { ui.addRecentColor(c); onPick(c); } }); } }));
      if (o.effects) el.appendChild(h('button', { class: 'cp-wide', type: 'button', text: o.mode === 'line' ? 'Patterned Lines...' : 'Fill Effects...', onclick: () => { close(); onPick({ effects: true }); } }));
      return el;
    };
    ui.openMenu([{ custom: panel }], { left: r.left, bottom: r.bottom }, { cls: 'color-menu' });
  };
  /* the colour names Office shows, in the UI's language where they are shown */
  const N = (x) => x;
  const CNAMES = { '#000000': N('Black'), '#993300': N('Brown'), '#333300': N('Olive Green'), '#003300': N('Dark Green'), '#003366': N('Dark Teal'), '#000080': N('Dark Blue'), '#333399': N('Indigo'), '#333333': N('Gray-80%'), '#800000': N('Dark Red'), '#FF6600': N('Orange'), '#808000': N('Dark Yellow'), '#008000': N('Green'), '#008080': N('Teal'), '#0000FF': N('Blue'), '#666699': N('Blue-Gray'), '#808080': N('Gray-50%'), '#FF0000': N('Red'), '#FF9900': N('Light Orange'), '#99CC00': N('Lime'), '#339966': N('Sea Green'), '#33CCCC': N('Aqua'), '#3366FF': N('Light Blue'), '#800080': N('Violet'), '#969696': N('Gray-40%'), '#FF00FF': N('Pink'), '#FFCC00': N('Gold'), '#FFFF00': N('Yellow'), '#00FF00': N('Bright Green'), '#00FFFF': N('Turquoise'), '#00CCFF': N('Sky Blue'), '#993366': N('Plum'), '#C0C0C0': N('Gray-25%'), '#FF99CC': N('Rose'), '#FFCC99': N('Tan'), '#FFFF99': N('Light Yellow'), '#CCFFCC': N('Light Green'), '#CCFFFF': N('Light Turquoise'), '#99CCFF': N('Pale Blue'), '#CC99FF': N('Lavender'), '#FFFFFF': N('White') };
  ui.colorName = (c) => t(CNAMES[String(c).toUpperCase()]) || String(c).toUpperCase();
  /** "Colors" dialog with Standard and Custom tabs. cb(hex|null) */
  ui.moreColors = function (initial, cb) {
    let cur = initial && initial[0] === '#' ? initial : '#3366FF';
    const prevNew = h('div', { class: 'cc-new', style: `background:${cur}` });
    const prevOld = h('div', { class: 'cc-cur', style: `background:${cur}` });
    const set = (c) => { cur = c.toUpperCase(); prevNew.style.background = cur; hex.value = cur; const [r2, g2, b2] = L.color.hexToRgb(cur); R.set(r2); Gs.set(g2); Bs.set(b2); };
    /* standard: hexagon-ish honeycomb approximated by staggered rows */
    const std = h('div', { class: 'cc-honey' });
    const rows = 13;
    for (let r2 = 0; r2 < rows; r2++) {
      const n = r2 < 7 ? 7 + r2 : 19 - r2;
      const row = h('div', { class: 'cc-hrow' });
      for (let i = 0; i < n; i++) {
        const ang = ((i / n) * 360 + r2 * 23) % 360;
        const dist = Math.abs(r2 - 6) / 6, light = 0.5 + (r2 < 6 ? 0.35 : -0.25) * dist;
        const [cr, cg, cb] = L.color.hslToRgb(ang / 360, 1 - dist * 0.4, L.clamp(light, 0.1, 0.92));
        const c = L.color.rgbToHex(cr, cg, cb);
        row.appendChild(h('button', { type: 'button', class: 'cc-hex', style: `background:${c}`, 'aria-label': c, onclick: () => set(c) }));
      }
      std.appendChild(row);
    }
    const grays = h('div', { class: 'cc-grays' });
    for (let i = 0; i <= 15; i++) { const v = Math.round((255 * i) / 15); const c = L.color.rgbToHex(v, v, v); grays.appendChild(h('button', { type: 'button', class: 'cc-hex', style: `background:${c}`, 'aria-label': c, onclick: () => set(c) })); }
    const nat = h('input', { type: 'color', id: 'cc-native', value: cur.toLowerCase() });
    nat.addEventListener('input', () => set(nat.value));
    const R = ui.spin({ value: 0, min: 0, max: 255, step: 1, dec: 0, onChange: () => upd() });
    const Gs = ui.spin({ value: 0, min: 0, max: 255, step: 1, dec: 0, onChange: () => upd() });
    const Bs = ui.spin({ value: 0, min: 0, max: 255, step: 1, dec: 0, onChange: () => upd() });
    const hex = h('input', { type: 'text', id: 'cc-hex', value: cur, maxlength: 7, style: 'width:80px' });
    hex.addEventListener('change', () => { if (/^#?[0-9a-f]{6}$/i.test(hex.value)) set(hex.value[0] === '#' ? hex.value : '#' + hex.value); });
    function upd() { cur = L.color.rgbToHex(R.get(), Gs.get(), Bs.get()); prevNew.style.background = cur; hex.value = cur; nat.value = cur.toLowerCase(); }
    const custom = h('div', { class: 'cc-custom' },
      h('div', { class: 'field' }, h('label', { for: 'cc-native', text: 'Pick:' }), nat),
      ui.field('&Red:', R), ui.field('&Green:', Gs), ui.field('&Blue:', Bs), ui.field('He&x:', hex));
    const tabs = ui.tabs([{ label: 'Standard', body: h('div', null, h('div', { class: 'cc-cap', text: 'Colors:' }), std, grays) }, { label: 'Custom', body: custom }]);
    const side = h('div', { class: 'cc-side' }, h('div', { class: 'cc-cap', text: 'New|color' }), prevNew, prevOld, h('div', { class: 'cc-cap', text: 'Current' }));
    set(cur);
    let ok = false;
    ui.dialog({ title: 'Colors', body: h('div', { class: 'cc-wrap' }, tabs, side), width: 440, buttons: [{ label: 'OK', primary: true, onClick: () => { ok = true; } }, { label: 'Cancel' }] })
      .done.then(() => cb(ok ? cur : null));
  };

  /* ---------- brushes ---------- */
  /* A brush (Format Painter, highlighter) is a mode that acts on what the user marks with the mouse, once,
   * when the mouse button is let go: never while a selection is still being dragged out. A click on its
   * button arms it for one use, or ends it; a double-click keeps it on until the button again or Esc. One
   * brush is on at a time. The app gives
   *   el()           the element that shows the brush cursor (class 'painting')
   *   pick()         what the brush carries, e.g. the formatting at the selection; null: nothing to carry
   *   apply(v, ...)  paint what the user marked with v; true when something was painted
   *   hint, status   optional: status(hint) while it is on, status('') after
   * and calls brush.release(...) when the user lets go of the mouse over what they marked. */
  let liveBrush = null;
  ui.brush = function (o) {
    const b = {
      value: null,
      sticky: false,
      on: () => liveBrush === b,
      /** arm it (value: carry this instead of pick()) */
      start(sticky, value) {
        const v = value !== undefined ? value : o.pick();
        if (v == null) return false;
        if (liveBrush && liveBrush !== b) liveBrush.stop();
        liveBrush = b;
        b.value = v;
        b.sticky = !!sticky;
        o.el().classList.add('painting');
        if (o.status) o.status(o.hint || '');
        ui.refresh();
        return true;
      },
      stop() {
        if (liveBrush !== b) return;
        liveBrush = null;
        b.value = null;
        b.sticky = false;
        o.el().classList.remove('painting');
        if (o.status) o.status('');
        ui.refresh();
      },
      /** its button: the second click of a double-click keeps it on */
      click(ev) {
        if (ev && ev.detail === 2) { if (b.on()) { b.sticky = true; if (o.status) o.status(o.hint || ''); } else b.start(true); }
        else if (b.on()) b.stop();
        else b.start(false);
      },
      release(...args) {
        if (!b.on()) return false;
        const done = !!o.apply(b.value, ...args);
        if (done && !b.sticky) b.stop();
        return done;
      },
    };
    return b;
  };
  ui.brushOn = () => liveBrush;
  if (typeof document !== 'undefined') document.addEventListener('keydown', (e) => {
    if (e.key !== 'Escape' || !liveBrush || ui.dialogOpen() || ui.menuOpen()) return;
    e.preventDefault();
    e.stopPropagation();
    liveBrush.stop();
  }, true);

  /* ---------- drags ---------- */
  /* A drag is what a press with the primary button starts and its release ends: moving a shape, a ruler marker or a
   * cell range, drawing a shape. One at a time, the same in every app (LibreOffice's vcl tracking does the same):
   *   ui.drag(e, { move(ev, d), up(ev, d), cancel(d), threshold = 4, capture })
   * on the press event e (a pointerdown, or a mousedown: no pointer capture then, the window listeners follow) returns
   * the drag d ({ x0, y0, x, y, dx, dy, moved, data: {} for the app }); null for a press of another button.
   * - move runs only once the pointer is `threshold` px from the press (Windows' drag rectangle), so a click that
   *   never moved changes nothing: up gets d.moved false.
   * - up runs on the release of the primary button. The pointer is held by `capture` (by default e.target; false: not
   *   held, the window listeners follow), so the click and dblclick that follow go to that element; if it is drawn
   *   again in the middle of the drag, the drag goes on. Don't hold it on a CSS-scaled ancestor of text the browser
   *   selects: its selection then stops short of the pointer.
   * - cancel(d, why) runs instead, and nothing may commit (the app puts back what it moved), on: Esc ('esc', which
   *   goes no further), a second button ('button', its context menu doesn't open), a move that reports the button no
   *   longer held ('lost': it was let go where the page couldn't see), a lost capture or pointercancel ('lost'), the
   *   window losing focus or the page hidden ('blur'), Ctrl+Z or Ctrl+Y ('undo': the drag is what gets undone), a new
   *   drag ('new'), ui.dragCancel() ('api': a tool or view switch). A tool stays armed unless why is 'esc'; Esc also
   *   puts a brush down (ui.brush). Other keys go nowhere while the button is held, except the modifiers.
   * - While it moves, the cursor shown at the press stays wherever the pointer goes (`cursor` sets another one;
   *   false: none is held).
   * A text selection is the browser's own drag; an app may watch it with { capture: false, cursor: false } (Ecrivaine
   * does, so its release, or how it ended otherwise, is known the same way). */
  let liveDrag = null, swallow = null, cursorStyle = null;
  /** the events of the button that cancelled a drag go no further: its press, release, auxclick and context menu */
  const swallowButton = (button) => { swallow = { button, until: Date.now() + 1500, seen: new Set() }; };
  const holdCursor = (c) => {
    if (!cursorStyle) { cursorStyle = document.createElement('style'); document.head.appendChild(cursorStyle); }
    cursorStyle.textContent = c ? `*, *::before, *::after { cursor: ${c} !important; }` : '';
  };
  ui.drag = function (e, o) {
    if (e.button != null && e.button !== 0) return null;
    if (liveDrag) liveDrag.end(false, null, 'new');
    const d = { x0: e.clientX, y0: e.clientY, x: e.clientX, y: e.clientY, dx: 0, dy: 0, moved: false, data: {} };
    const id = e.pointerId, threshold = o.threshold == null ? 4 : o.threshold;
    const cap = o.capture === false ? null : o.capture || e.target;
    let cursor = o.cursor || null;
    if (!cursor && o.cursor !== false) { try { const c = getComputedStyle(e.target).cursor; if (c && c !== 'auto') cursor = c; } catch (err) { /* no cursor kept */ } }
    try { if (cap && id != null && cap.setPointerCapture) cap.setPointerCapture(id); } catch (err) { /* window listeners still follow */ }
    const on = (t, type, fn, capture) => { t.addEventListener(type, fn, capture); return () => t.removeEventListener(type, fn, capture); };
    const second = (button) => { swallowButton(button); end(false, null, 'button'); };
    const offs = [
      on(window, 'pointermove', (ev) => {
        if (id != null && ev.pointerId !== id) return;
        if (ev.buttons & 6) { second(ev.buttons & 2 ? 2 : 1); return; }
        if (!(ev.buttons & 1)) { end(false, null, 'lost'); return; }
        d.x = ev.clientX; d.y = ev.clientY; d.dx = d.x - d.x0; d.dy = d.y - d.y0;
        if (!d.moved && Math.hypot(d.dx, d.dy) < threshold) return;
        if (!d.moved && cursor) holdCursor(cursor);
        d.moved = true;
        if (o.move) o.move(ev, d);
      }, true),
      on(window, 'pointerup', (ev) => { if ((id == null || ev.pointerId === id) && ev.button === 0) end(true, ev); }, true),
      on(window, 'keydown', (ev) => {
        /* the modifiers shape the drag (Shift constrains it, Ctrl copies, Alt doesn't snap) */
        if (/^(Shift|Control|Alt|Meta|AltGraph|CapsLock)$/.test(ev.key)) return;
        ev.preventDefault(); ev.stopImmediatePropagation();
        /* Ctrl+Z (or Y) in the middle of a drag undoes the drag, not what came before it */
        if ((ev.ctrlKey || ev.metaKey) && /^[zy]$/i.test(ev.key)) { end(false, null, 'undo'); return; }
        if (ev.key === 'Escape') { end(false, null, 'esc'); if (liveBrush) liveBrush.stop(); }
        /* any other key waits for the release: an edit in the middle of a drag would come between its start and the
           step it records */
      }, true),
      on(window, 'pointercancel', () => end(false, null, 'lost'), true),
      on(window, 'blur', () => end(false, null, 'blur')),
      on(document, 'visibilitychange', () => { if (document.hidden) end(false, null, 'blur'); }),
    ];
    if (cap) offs.push(on(cap, 'lostpointercapture', () => { if (cap.isConnected) end(false, null, 'lost'); }));
    function end(ok, ev, why) {
      if (liveDrag !== live) return;
      liveDrag = null;
      offs.forEach((off) => off());
      if (cursorStyle) holdCursor(null);
      try { if (cap && id != null && cap.hasPointerCapture && cap.hasPointerCapture(id)) cap.releasePointerCapture(id); } catch (err) { /* ignore */ }
      if (ok) { if (o.up) o.up(ev, d); } else if (o.cancel) o.cancel(d, why);
    }
    const live = { end, d };
    liveDrag = live;
    return d;
  };
  ui.dragging = () => !!liveDrag;
  ui.dragCancel = () => { if (liveDrag) liveDrag.end(false, null, 'api'); };
  /* a second button cancels the drag; its own events then go no further (the pointer events report it as a move,
     before the mouse events come): its press, its release, and its auxclick or context menu, whichever ends that
     click, in whatever order the system sends them (or for 1.5 s at most) */
  if (typeof window !== 'undefined') for (const type of ['mousedown', 'mouseup', 'auxclick', 'contextmenu']) window.addEventListener(type, (e) => {
    if (e.button === 0) return;
    if (liveDrag) { swallowButton(e.button); liveDrag.end(false, null, 'button'); }
    if (!swallow || e.button !== swallow.button) return;
    if (Date.now() > swallow.until) { swallow = null; return; }
    e.preventDefault(); e.stopImmediatePropagation();
    swallow.seen.add(type);
    if (swallow.seen.has('mouseup') && swallow.seen.has(swallow.button === 2 ? 'contextmenu' : 'auxclick')) swallow = null;
  }, true);

  /* ---------- zoom ---------- */
  /* A notch of a mouse wheel (about 100px, or 3 lines) changes the zoom by roughly a tenth. A trackpad
     pinch arrives as many small deltaY values and stays smooth. */
  ui.wheelZoomFactor = function (deltaY, deltaMode) {
    let dy = deltaY || 0;
    if (deltaMode === 1) dy *= 32;
    else if (deltaMode === 2) dy *= 800;
    if (!dy) return 1;
    return Math.exp(-dy * 0.0015);
  };
  /** two fingers: scale is the change in the span between them, dx/dy how their midpoint moved */
  ui.pinchStep = function (a0, b0, a1, b1) {
    const dist = (p, q) => Math.hypot(p.x - q.x, p.y - q.y);
    const d0 = dist(a0, b0);
    return {
      scale: d0 > 0.5 ? dist(a1, b1) / d0 : 1,
      dx: (a1.x + b1.x) / 2 - (a0.x + b0.x) / 2,
      dy: (a1.y + b1.y) / 2 - (a0.y + b0.y) / 2,
    };
  };
  /** A two-finger tap is a right-click. g is the gesture so far, or null. ev is
      { type: 'start'|'move'|'end'|'cancel', t, touches: [{ id, x, y }] } with the fingers still down.
      Both fingers up within 500ms, the first finger within 14px of where it was and the span between
      them within 18px (or a fifth of that span), is a tap: menu is that first finger's point.
      cancelDrag is set while two fingers are down, so the press does not also drag. */
  ui.twoFingerTap = function (g, ev) {
    const touches = ev.touches || [];
    const at = (id) => touches.find((t) => t.id === id);
    if (ev.type === 'cancel' || (ev.type === 'start' && touches.length > 2)) return { g: null };
    if (ev.type === 'start' && touches.length === 2 && !g) {
      const first = touches[0].id < touches[1].id ? touches[0] : touches[1];
      const other = first === touches[0] ? touches[1] : touches[0];
      return { g: { id: first.id, oid: other.id, x: first.x, y: first.y, t: ev.t, span: Math.hypot(other.x - first.x, other.y - first.y), native: false }, cancelDrag: true };
    }
    if (!g) return { g: null };
    if (ev.type === 'move') {
      const a = at(g.id);
      if (!a || Math.hypot(a.x - g.x, a.y - g.y) > 14) return { g: null, cancelDrag: touches.length > 1 };
      const b = at(g.oid);
      if (b && Math.abs(Math.hypot(b.x - a.x, b.y - a.y) - g.span) > Math.max(18, g.span * 0.2)) return { g: null, cancelDrag: true };
      return { g, cancelDrag: touches.length > 1 };
    }
    if (ev.type === 'end') {
      if (touches.length > 0) return { g, cancelDrag: touches.length > 1 };
      if (g.native || ev.t - g.t > 500) return { g: null };
      return { g: null, menu: { x: g.x, y: g.y }, cancelDrag: true };
    }
    return { g };
  };
  /** scroll position that keeps the point under `from` under `to` after oldZ becomes newZ.
      `from` and `to` are viewport coordinates; `scroll` is the element's scrollLeft/scrollTop. */
  ui.zoomScroll = function (scroll, oldZ, newZ, from, to) {
    const ratio = oldZ ? newZ / oldZ : 1;
    return {
      left: (from.x + scroll.left) * ratio - to.x,
      top: (from.y + scroll.top) * ratio - to.y,
    };
  };
  /** Where a zoom percent sits on the status-bar slider: 0 at the minimum, 1 at the maximum.
      100% is the middle, so the left half runs from the minimum to 100% and the right half
      from 100% to the maximum. A range that does not contain 100% stays linear. */
  ui.zoomTrack = function (pct, min, max) {
    min = Number(min); max = Number(max);
    pct = Math.min(max, Math.max(min, Number(pct)));
    if (!(min < 100 && max > 100)) return max === min ? 0 : (pct - min) / (max - min);
    if (pct <= 100) return (pct - min) / (100 - min) * 0.5;
    return 0.5 + (pct - 100) / (max - 100) * 0.5;
  };
  /** Zoom percent at track position t (0..1). The inverse of ui.zoomTrack. */
  ui.zoomFromTrack = function (t, min, max) {
    min = Number(min); max = Number(max);
    t = Math.min(1, Math.max(0, Number(t) || 0));
    if (!(min < 100 && max > 100)) return min + t * (max - min);
    if (t <= 0.5) return min + (t / 0.5) * (100 - min);
    return 100 + ((t - 0.5) / 0.5) * (max - 100);
  };
  /** status-bar zoom: −, a slider, +, and a percent that can be typed. o.min/o.max are percents.
      o.value() is the current percent; o.onChange(pct) runs as the slider moves and as a typed
      number settles. The slider's middle is 100% (ui.zoomTrack). Listens for 'zoom' (and
      'layout-done') to catch up. */
  ui.zoomBar = function (o) {
    const min = o.min, max = o.max, step = o.step || 10, scale = 1000;
    const minus = h('button', { type: 'button', class: 'sb-zoom-btn', tabindex: '-1', 'aria-label': 'Zoom Out', text: '−' });
    const range = h('input', { type: 'range', class: 'sb-zoom-range', min: '0', max: String(scale), step: '1', value: String(scale / 2), 'aria-label': 'Zoom' });
    const plus = h('button', { type: 'button', class: 'sb-zoom-btn', tabindex: '-1', 'aria-label': 'Zoom In', text: '+' });
    const pct = h('input', { type: 'text', class: 'sb-zoom-pct', value: '100%', 'aria-label': 'Zoom', inputmode: 'numeric', spellcheck: 'false', autocomplete: 'off' });
    const box = h('div', { class: 'sb-zoom' }, minus, range, plus, pct);
    let typing = false, revert = false, typeGen = 0, sliding = false;
    const read = () => {
      const n = Math.round(Number(o.value()));
      return n >= min && n <= max ? n : Math.round(L.clamp(Number(o.value()) || 100, min, max));
    };
    const show = (n) => {
      /* while the pointer is down the thumb follows it; snapping onto the percent would fight the drag */
      if (!sliding) {
        const pos = String(Math.round(ui.zoomTrack(n, min, max) * scale));
        if (range.value !== pos) range.value = pos;
      }
      range.setAttribute('aria-valuemin', String(min));
      range.setAttribute('aria-valuemax', String(max));
      range.setAttribute('aria-valuenow', String(n));
      range.setAttribute('aria-valuetext', n + '%');
      if (document.activeElement !== pct) pct.value = n + '%';
      /* the accent runs to the thumb's centre, half a thumb in from each end */
      const lo = Number(range.min) || 0, hi = Number(range.max);
      const ratio = hi === lo ? 0 : (Number(range.value) - lo) / (hi - lo);
      range.style.setProperty('--gg-zoom-ratio', String(ratio));
    };
    const apply = (n) => { n = Math.round(L.clamp(n, min, max)); o.onChange(n); show(n); };
    const sync = () => { if (!typing) show(read()); };
    const stepTo = (dir) => {
      const n = read();
      apply(dir < 0 ? Math.floor((n - 1) / step) * step : Math.ceil((n + 1) / step) * step);
    };
    const fromSlider = () => Math.round(L.clamp(ui.zoomFromTrack(Number(range.value) / scale, min, max), min, max));
    const liveType = L.debounce((gen) => {
      if (gen !== typeGen || !typing) return;
      const m = /^(\d{1,4})\s*%?$/.exec(pct.value.trim());
      if (!m) return;
      const n = Number(m[1]);
      if (n >= min && n <= max) { o.onChange(n); show(n); }
    }, 250);
    minus.addEventListener('click', () => stepTo(-1));
    plus.addEventListener('click', () => stepTo(1));
    range.addEventListener('pointerdown', () => { sliding = true; });
    range.addEventListener('input', () => {
      const n = fromSlider();
      if (n !== read()) o.onChange(n);
      show(read());
    });
    const endSlide = () => { if (!sliding) return; sliding = false; show(read()); };
    range.addEventListener('pointerup', () => {
      endSlide();
      range.blur();
      setTimeout(() => { if (!document.activeElement || document.activeElement === document.body) ui.hooks.refocus(); }, 0);
    });
    range.addEventListener('pointercancel', endSlide);
    range.addEventListener('blur', endSlide);
    range.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') { e.preventDefault(); range.blur(); ui.hooks.refocus(); return; }
      /* the track's own steps are positions, not percents, so the keys move the zoom itself */
      let n = null;
      if (e.key === 'ArrowLeft' || e.key === 'ArrowDown') n = read() - 1;
      else if (e.key === 'ArrowRight' || e.key === 'ArrowUp') n = read() + 1;
      else if (e.key === 'PageDown') n = read() - step;
      else if (e.key === 'PageUp') n = read() + step;
      else if (e.key === 'Home') n = min;
      else if (e.key === 'End') n = max;
      if (n === null) return;
      e.preventDefault();
      apply(n);
    });
    pct.addEventListener('focus', () => { typing = true; pct.value = String(read()); pct.select(); });
    pct.addEventListener('input', () => liveType(++typeGen));
    pct.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); pct.blur(); }
      else if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); revert = true; pct.blur(); }
      else if (e.key === 'ArrowUp' || e.key === 'ArrowDown') { e.preventDefault(); stepTo(e.key === 'ArrowUp' ? 1 : -1); pct.value = String(read()); }
    });
    pct.addEventListener('blur', () => {
      typing = false;
      typeGen++;
      if (!revert) {
        const m = /(\d+)/.exec(pct.value);
        apply(m ? Number(m[1]) : read());
      }
      revert = false;
      show(read());
      setTimeout(() => { if (!document.activeElement || document.activeElement === document.body) ui.hooks.refocus(); }, 0);
    });
    if (L.bus) { L.bus.on('zoom', sync); L.bus.on('layout-done', sync); }
    sync();
    queueMicrotask(() => installDocStatus(box));
    return box;
  };
  /* on a phone the status cells are one button; the zoom control stays on the right. The same cells move into the window, so they keep updating. */
  const narrowStatus = () => typeof matchMedia === 'function' && matchMedia('(max-width: 960px)').matches;
  function installDocStatus(zoom) {
    const sb = zoom.parentElement;
    if (!sb || sb.id !== 'statusbar' || sb.querySelector('.sb-docstatus')) return;
    const btn = h('button', { type: 'button', class: 'sb-docstatus', text: 'Document Status' });
    sb.insertBefore(btn, zoom);
    let session = null;
    const restore = (cells) => {
      const z = sb.querySelector('.sb-zoom');
      for (const c of cells) sb.insertBefore(c, z);
      session = null;
    };
    btn.addEventListener('click', () => {
      if (session) { session.close(null); return; }
      const cells = [...sb.querySelectorAll(':scope > .sb-cell')];
      const list = h('div', { class: 'sb-status-list' });
      for (const c of cells) list.append(c);
      session = ui.dialog({
        title: '',
        body: list,
        width: 100,
        modeless: true,
        buttons: [],
        onOpen(dlg) {
          dlg.classList.add('sb-status-dlg');
          const place = () => {
            const b = btn.getBoundingClientRect();
            const w = dlg.offsetWidth || 100;
            const top = b.top - dlg.offsetHeight - 6;
            dlg.style.left = Math.max(4, Math.min(b.left, window.innerWidth - w - 4)) + 'px';
            dlg.style.top = Math.max(4, top) + 'px';
          };
          place();
          requestAnimationFrame(place);
        },
        onClose() { restore(cells); },
      });
    });
    window.addEventListener('resize', () => { if (session && !narrowStatus()) session.close(null); });
  }
  /** Pinch and two-finger pan on an editing surface. o.zoom() is a fraction (1 = 100%), o.setZoom(z)
      applies one, o.min/o.max bound it. o.about(oldZ, newZ, x0, y0, x1, y1) replaces the default,
      which scrolls `el`. o.wheelScroll(e) handles a two-finger swipe that is not a pinch; without it
      the browser scrolls `el`. o.ignore(e) leaves the event alone. o.metaZoom also treats ⌘+scroll
      as a pinch. A pinch anywhere else is cancelled so the browser does not zoom the page. */
  ui.watchDocGestures = function (el, o) {
    if (!el) return;
    el.setAttribute('data-doc-zoom', '');
    let gesture = null;
    const applyZoom = (oldZ, newZ, x0, y0, x1, y1) => {
      newZ = L.clamp(newZ, o.min, o.max);
      if (o.about) { o.about(oldZ, newZ, x0, y0, x1, y1); return; }
      const rect = el.getBoundingClientRect();
      const scroll = { left: el.scrollLeft, top: el.scrollTop };
      if (Math.abs(newZ - oldZ) > 1e-4) o.setZoom(newZ);
      const s = ui.zoomScroll(scroll, oldZ, newZ, { x: x0 - rect.left, y: y0 - rect.top }, { x: x1 - rect.left, y: y1 - rect.top });
      el.scrollLeft = s.left;
      el.scrollTop = s.top;
    };
    el.addEventListener('wheel', (e) => {
      if (o.ignore && o.ignore(e)) return;
      if (e.ctrlKey || (o.metaZoom && e.metaKey)) {
        e.preventDefault();
        if (gesture || pinch) return;
        const oldZ = o.zoom();
        applyZoom(oldZ, oldZ * ui.wheelZoomFactor(e.deltaY, e.deltaMode), e.clientX, e.clientY, e.clientX, e.clientY);
        return;
      }
      if (o.wheelScroll) o.wheelScroll(e);
    }, { passive: false, capture: true });
    const fingers = (e) => [{ x: e.touches[0].clientX, y: e.touches[0].clientY }, { x: e.touches[1].clientX, y: e.touches[1].clientY }];
    let pinch = null;
    el.addEventListener('touchstart', (e) => {
      if (e.touches.length !== 2 || (o.ignore && o.ignore(e))) return;
      const [a, b] = fingers(e);
      pinch = { a, b, z: o.zoom() };
    }, { passive: false });
    el.addEventListener('touchmove', (e) => {
      if (!pinch || e.touches.length !== 2) return;
      e.preventDefault();
      const [a, b] = fingers(e);
      const step = ui.pinchStep(pinch.a, pinch.b, a, b);
      const newZ = L.clamp(pinch.z * step.scale, o.min, o.max);
      const mid0 = { x: (pinch.a.x + pinch.b.x) / 2, y: (pinch.a.y + pinch.b.y) / 2 };
      const mid1 = { x: (a.x + b.x) / 2, y: (a.y + b.y) / 2 };
      applyZoom(pinch.z, newZ, mid0.x, mid0.y, mid1.x, mid1.y);
      pinch.a = a; pinch.b = b; pinch.z = newZ;
    }, { passive: false });
    const endPinch = () => { pinch = null; };
    el.addEventListener('touchend', endPinch);
    el.addEventListener('touchcancel', endPinch);
    el.addEventListener('gesturestart', (e) => { e.preventDefault(); gesture = { z: o.zoom(), x: e.clientX, y: e.clientY }; }, { passive: false });
    el.addEventListener('gesturechange', (e) => {
      e.preventDefault();
      if (!gesture || pinch) return;
      const newZ = L.clamp(gesture.z * (e.scale || 1), o.min, o.max);
      applyZoom(o.zoom(), newZ, gesture.x, gesture.y, e.clientX, e.clientY);
      gesture.x = e.clientX; gesture.y = e.clientY;
    }, { passive: false });
    el.addEventListener('gestureend', () => { gesture = null; });
  };
  if (typeof window !== 'undefined') {
    window.addEventListener('wheel', (e) => { if (e.ctrlKey) e.preventDefault(); }, { passive: false, capture: true });
    const stopGesture = (e) => e.preventDefault();
    window.addEventListener('gesturestart', stopGesture, { passive: false, capture: true });
    window.addEventListener('gesturechange', stopGesture, { passive: false, capture: true });
    /* a two-finger tap opens the right-click menu at the finger that landed first */
    let fingers = null, eatClick = 0, opened = 0;
    const touchesOf = (e) => [...e.touches].map((t) => ({ id: t.identifier, x: t.clientX, y: t.clientY }));
    const take = (e, type) => {
      const r = ui.twoFingerTap(fingers, { type, t: Date.now(), touches: touchesOf(e) });
      fingers = r.g;
      if (r.cancelDrag) ui.dragCancel();
      if (!r.menu) return;
      const el = document.elementFromPoint(r.menu.x, r.menu.y);
      if (!el) return;
      const at = { bubbles: true, cancelable: true, view: window, clientX: r.menu.x, clientY: r.menu.y, screenX: r.menu.x, screenY: r.menu.y, button: 2, buttons: 2 };
      try { el.dispatchEvent(new PointerEvent('pointerdown', Object.assign({ pointerId: 1, pointerType: 'touch', isPrimary: true }, at))); } catch (err) { /* the menu event below still opens it */ }
      opened = Date.now();
      if (el.dispatchEvent(new MouseEvent('contextmenu', at))) { opened = 0; return; }
      e.preventDefault();
      eatClick = Date.now() + 500;
    };
    window.addEventListener('touchstart', (e) => take(e, 'start'), { capture: true, passive: true });
    window.addEventListener('touchmove', (e) => take(e, 'move'), { capture: true, passive: true });
    window.addEventListener('touchend', (e) => take(e, 'end'), { capture: true, passive: false });
    window.addEventListener('touchcancel', (e) => take(e, 'cancel'), { capture: true, passive: true });
    window.addEventListener('contextmenu', (e) => {
      if (!e.isTrusted) return;
      if (fingers) fingers.native = true;
      /* the tap already opened the menu; the system's own right-click for the same tap does not open another */
      if (opened && Date.now() - opened < 400) { e.preventDefault(); e.stopImmediatePropagation(); }
    }, true);
    window.addEventListener('click', (e) => {
      if (!eatClick || Date.now() > eatClick) return;
      eatClick = 0;
      e.preventDefault();
      e.stopPropagation();
    }, true);
  }

  /* ---------- status bar & misc ---------- */
  /* one notice. A later call replaces the words and restarts the wait. Hiding fades it out, then sets hidden. */
  function toastEl() {
    let t = L.$('.toast');
    if (t) return t;
    t = h('div', { class: 'toast', role: 'status', hidden: true });
    t.appendChild(h('span', { class: 'toast-in' }));
    document.body.appendChild(t);
    return t;
  }
  function showToast(t) {
    const gen = (t._gen || 0) + 1;
    t._gen = gen;
    clearTimeout(t._hide);
    const replay = t.hidden || t.classList.contains('is-out') || !t.classList.contains('is-in');
    t.classList.remove('is-out');
    t.hidden = false;
    if (!replay) return;
    t.classList.remove('is-in');
    void t.offsetWidth;
    if (t._gen !== gen) return;
    t.classList.add('is-in');
  }
  function hideToast(t) {
    if (!t || t.hidden || t.classList.contains('is-out')) return;
    const gen = (t._gen || 0) + 1;
    t._gen = gen;
    clearTimeout(t._t);
    t.classList.remove('is-in');
    t.classList.add('is-out');
    const finish = () => {
      if (t._gen !== gen) return;
      t.hidden = true;
      t.classList.remove('is-out');
    };
    const onEnd = (e) => {
      if (e.target !== t || e.animationName !== 'toast-out') return;
      t.removeEventListener('animationend', onEnd);
      finish();
    };
    t.addEventListener('animationend', onEnd);
    t._hide = setTimeout(() => { t.removeEventListener('animationend', onEnd); finish(); }, 400);
  }
  /** a short notice at the bottom; action: { label, run } adds a button to it (e.g. Browse) */
  ui.toast = (text, ms, action) => {
    const t = toastEl();
    const inn = t.querySelector('.toast-in');
    inn.textContent = L.t ? L.t(text) : text;
    if (action) inn.append(' ', h('button', { class: 'toast-act', type: 'button', text: action.label, onclick: () => { hideToast(t); action.run(); } }));
    showToast(t);
    clearTimeout(t._t);
    t._t = setTimeout(() => hideToast(t), ms || 2600);
  };
  /* a modal while work keeps the user waiting. The title is the app's name. No close button and no buttons, so
     Escape does not dismiss it. A window already open stays where it is and this one covers it. Another call
     while it is up only changes the message. */
  let busyDlg = null;
  ui.busy = (on, text) => {
    const open = !!(busyDlg && busyDlg.el.isConnected);
    if (!!on !== open) tell('vo-busy', { on: !!on, text: on ? text || 'Working...' : '' });
    if (!on) { if (open) { const d = busyDlg; busyDlg = null; d.close(null); } return; }
    const msg = t(text || 'Working...');
    if (open) { busyDlg.text.textContent = msg; return; }
    const label = h('span', { class: 'busy-t' });
    label.textContent = msg;
    busyDlg = ui.dialog({
      title: L.APP || L.APP_NAME || 'Melina Office',
      width: 360,
      canClose: false,
      buttons: [],
      body: h('div', { class: 'busy-box' }, h('span', { class: 'busy-spin' }), label),
      focus: label,
    });
    busyDlg.text = label;
    busyDlg.el.classList.add('busy');
  };
  /* a host that set VO.openingDialog (the VS Code editor) shows it once dialogs exist, before the app paints */
  if (typeof window !== 'undefined' && window.VO && typeof VO.openingDialog === 'function') {
    const run = () => VO.openingDialog();
    if (VO.i18n) VO.i18n.whenReady(run); else run();
  }
})(typeof window !== 'undefined' ? window : globalThis);
