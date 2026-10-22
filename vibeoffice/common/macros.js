/* Tools ▸ Macros. Quire, Lectern and Ledger fill L.macroHost and open this window after their notices.
   Step Into and Organizer stay disabled: stepping waits for the Visual Basic editor. */
(function (root) {
  'use strict';
  const L = root.L || (root.L = {});

  /** Which of the right-hand buttons do anything. Names are matched without regard to case. */
  L.macroButtonState = function (name, names, enabled) {
    const typed = String(name || '');
    const hit = !!(enabled && typed !== '' && (names || []).some((n) => String(n).toLowerCase() === typed.toLowerCase()));
    const fresh = !!(enabled && typed !== '' && !hit);
    return {
      run: hit,
      step: false,
      edit: hit,
      create: fresh,
      del: hit,
      organizer: false,
      cancel: true,
      primary: hit ? 'run' : fresh ? 'create' : null,
    };
  };

  L.openMacroDialog = function () {
    if (!L.ui || typeof L.ui.dialog !== 'function') return;
    const ui = L.ui;
    const host = L.macroHost;
    const enabled = !!(host && host.enabled);
    const nameInput = L.h('input', { type: 'text', id: 'macro-name', style: 'width:100%', spellcheck: 'false', autocomplete: 'off' });
    const list = L.h('div', { class: 'listbox macro-dlg-list', role: 'listbox', translate: 'no' });
    const where = L.h('select', { id: 'macro-in', disabled: !enabled },
      L.h('option', { value: 'document', text: 'Current Document' }),
      L.h('option', { value: 'all', text: 'All templates and documents' }));
    const desc = L.h('textarea', { id: 'macro-desc', class: 'macro-dlg-desc', readonly: 'readonly', spellcheck: 'false' });
    if (!enabled) { nameInput.disabled = true; desc.disabled = true; }
    const specs = [
      { id: 'run', label: '&Run' },
      { id: 'step', label: '&Step Into' },
      { id: 'edit', label: '&Edit' },
      { id: 'create', label: '&Create' },
      { id: 'del', label: '&Delete' },
      { id: 'organizer', label: 'Or&ganizer...' },
      { id: 'cancel', label: 'Cancel' },
    ];
    const buttons = {};
    for (const { id, label } of specs) {
      const el = L.h('button', { type: 'button', class: 'btn' + (id === 'cancel' ? ' macro-cancel' : ''), id: id === 'cancel' ? 'macro-cancel' : null, html: ui.ampHTML(label) });
      el._macroKey = label;
      buttons[id] = el;
    }
    let names = [];
    let closeDlg = function () {};
    function typed() { return nameInput.value; }
    function loc() { return where.value === 'all' ? 'all' : 'document'; }
    function exact() {
      const t = typed().toLowerCase();
      return names.find((n) => n.toLowerCase() === t) || null;
    }
    function paint() {
      const state = L.macroButtonState(typed(), names, enabled);
      const hit = exact();
      L.clear(list);
      for (const n of names) {
        const row = L.h('div', { class: 'li' + (hit === n ? ' on' : ''), role: 'option', translate: 'no' }, document.createTextNode(n));
        row.addEventListener('mousedown', (e) => {
          e.preventDefault();
          nameInput.value = n;
          paint();
        });
        row.addEventListener('dblclick', () => { if (L.macroButtonState(n, names, enabled).run) press('run'); });
        list.appendChild(row);
      }
      desc.value = hit && host && host.describe ? host.describe(hit, loc()) : '';
      for (const id of Object.keys(buttons)) {
        const on = !!state[id];
        buttons[id].disabled = !on;
        buttons[id].classList.toggle('primary', state.primary === id);
      }
    }
    function reload() {
      const listed = enabled && host && host.names ? host.names(loc()) : [];
      names = (listed || []).map((n) => (typeof n === 'string' ? n : n.name));
      names.sort((a, b) => a.localeCompare(b));
      paint();
    }
    function fail(e) {
      const msg = e && e.message ? e.message : String(e || '');
      ui.msg(msg || 'The macro could not run.', { icon: 'error', title: 'Macros' });
    }
    function press(id) {
      const state = L.macroButtonState(typed(), names, enabled);
      if (!state[id]) return;
      const name = exact() || typed();
      if (id === 'cancel') { closeDlg(); return; }
      if (id === 'del') {
        const ok = !host.remove || host.remove(name, loc());
        if (!ok) {
          if (loc() === 'all') ui.msg('The macros stored on this computer could not be saved.', { icon: 'error', title: 'Macros' });
          return;
        }
        nameInput.value = '';
        reload();
        return;
      }
      if (id === 'run') {
        closeDlg();
        Promise.resolve(host.run(name, loc())).catch(fail);
        return;
      }
      if (id === 'edit') { closeDlg(); if (host.edit) host.edit(name, loc()); return; }
      if (id === 'create') { closeDlg(); if (host.create) host.create(typed(), loc()); }
    }
    for (const id of Object.keys(buttons)) buttons[id].addEventListener('click', () => press(id));
    nameInput.addEventListener('input', paint);
    where.addEventListener('change', () => { nameInput.value = ''; reload(); });
    const body = L.h('div', { class: 'macro-dlg-body' },
      L.h('div', { class: 'macro-dlg-main' },
        L.h('label', { for: 'macro-name', html: ui.ampHTML('Macro na&me:') }),
        nameInput,
        list,
        L.h('label', { for: 'macro-in', html: ui.ampHTML('M&acros in:') }),
        where,
        L.h('label', { for: 'macro-desc', html: ui.ampHTML('Descr&iption:') }),
        desc),
      L.h('div', { class: 'macro-dlg-side' }, buttons.run, buttons.step, buttons.edit, buttons.create, buttons.del, buttons.organizer, buttons.cancel));
    const box = ui.dialog({
      title: 'Macros',
      width: 520,
      body: body,
      buttons: [],
      focus: enabled ? '#macro-name' : '#macro-cancel',
      onOpen: (el) => {
        el.classList.add('macro-dlg');
        el.addEventListener('keydown', (e) => {
          if (e.altKey && !e.ctrlKey && !e.metaKey && e.key.length === 1) {
            const ch = e.key.toLowerCase();
            if (ch === ui.accessKey('Macro na&me:') && !nameInput.disabled) { e.preventDefault(); nameInput.focus(); nameInput.select(); return; }
            if (ch === ui.accessKey('Descr&iption:') && !desc.disabled) { e.preventDefault(); desc.focus(); return; }
            if (ch === ui.accessKey('M&acros in:') && !where.disabled) { e.preventDefault(); where.focus(); return; }
            for (const id of Object.keys(buttons)) {
              if (buttons[id].disabled) continue;
              if (ui.accessKey(buttons[id]._macroKey) === ch) { e.preventDefault(); buttons[id].click(); return; }
            }
          }
          if (e.key === 'Enter' && !e.shiftKey && e.target && e.target.tagName !== 'TEXTAREA' && e.target.tagName !== 'BUTTON' && e.target.tagName !== 'SELECT') {
            const state = L.macroButtonState(typed(), names, enabled);
            if (state.primary) { e.preventDefault(); press(state.primary); }
          }
        });
      },
    });
    closeDlg = () => { if (box && box.close) box.close(); };
    reload();
    return box;
  };
})(typeof window !== 'undefined' ? window : globalThis);
