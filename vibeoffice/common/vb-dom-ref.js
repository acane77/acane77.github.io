/* Visual Basic DOM API Reference, shared by Quire, Ledger and Lectern.
   Each app registers its own catalog with mount(). The language section is the same.
   A change to an app's object model updates docs/<app>_dom_api_specifications.md
   and public/<app>/js/vb-dom-ref.js in the same change. */
(function () {
  'use strict';
  const L = window.L;
  const { h } = L;
  const N = L.N;
  const KIND = { get: N('Read-only'), set: N('Write'), rw: N('Read/write'), call: N('Call') };

  /* the same limits in every app: VB6 on the open document, not VBScript and not the full VBA host */
  const LANGUAGE = [
    [N('What this runs'), [
      N('This is Visual Basic 6 on the open document. It is not VBScript, and it is not the full Microsoft VBA host.'),
      N('Only members listed on this page run. Any other name fails.'),
      N('Application and Parent are on every object, for navigation. Creator fails on every object: this app does not report the Microsoft Office creator code.'),
    ]],
    [N('The console'), [
      N('A line is one statement. Assign with x = 1. The name stays, and a later line can use it. ?x prints the value.'),
      N('A line ending in _ continues on the next line. An unfinished Sub, Function, If or With stays open until it ends.'),
      N('A procedure typed in the console does not become a macro. A whole script belongs in Run Script.'),
    ]],
    [N('Run Script and macros'), [
      N('Run Script compiles one standard module and runs Sub Main. It does not see macros saved in the document.'),
      N('Document macros are Visual Basic 6 source stored with the file, or in this browser for all templates and documents. A Microsoft VBA project in the file is kept and is not run.'),
    ]],
    [N('How a call differs'), [
      N('Collections are 1-based. Item(1) is the first. Count, First and Last are properties.'),
      N('True is -1 and False is 0. Arguments stay in VBA order. Omit arguments from the right. There is no options object.'),
    ]],
    [N('Not in this host'), [
      N('There is no class module, no UserForm, no Declare and no Win32 call.'),
      N('A path, a printer, a file picker or a modal VBA dialog fails. Those calls need a host this page does not have.'),
    ]],
  ];

  function show() {
    if (L.panes && L.panes.task) L.panes.task.show('vb-dom');
  }

  function para(text) {
    return h('p', { class: 'vbref-p', 'data-hay': String(text || '').toLowerCase(), text: text });
  }

  function member(m) {
    const kind = KIND[m.kind] || '';
    const hay = [m.name, kind, m.gap || '', m.note || ''].join(' ').toLowerCase();
    return h('div', { class: 'vbref-mem', 'data-hay': hay },
      h('div', { class: 'vbref-mem-line' },
        h('code', { class: 'vbref-name', translate: 'no', text: m.name }),
        kind ? h('span', { class: 'vbref-kind', text: kind }) : null,
        m.gap ? h('span', { class: 'vbref-gap', translate: 'no', text: m.gap }) : null),
      m.note ? h('div', { class: 'vbref-note', text: m.note }) : null);
  }

  function section(title, open, nodes, ident) {
    const body = h('div', { class: 'vbref-body' });
    if (!open) body.hidden = true;
    const btn = h('button', { class: 'vbref-h', type: 'button', 'aria-expanded': open ? 'true' : 'false' }, h('span', ident ? { translate: 'no', text: title } : { text: title }));
    btn.addEventListener('click', () => {
      const on = btn.getAttribute('aria-expanded') !== 'true';
      btn.setAttribute('aria-expanded', on ? 'true' : 'false');
      body.hidden = !on;
    });
    for (const n of nodes) if (n) body.append(n);
    return h('div', { class: 'vbref-sec' }, btn, body);
  }

  function blocks(groups) {
    const out = [];
    for (const [title, lines] of groups) {
      out.push(h('div', { class: 'tp-h', text: title }));
      for (const line of lines) out.push(para(line));
    }
    return out;
  }

  function render(b, catalog) {
    const q = h('input', { type: 'text', class: 'vbref-find', placeholder: 'Search the reference', 'aria-label': 'Search the reference' });
    const empty = h('div', { class: 'tp-note', hidden: true, text: 'No topics match your search.' });
    const root = h('div');
    const parts = [];
    parts.push(section(N('Language'), true, blocks(LANGUAGE)));
    parts.push(section(N('This application'), true, (catalog.rules || []).map(para).concat((catalog.gaps || []).map((g) =>
      h('div', { class: 'vbref-mem', 'data-hay': (g.code + ' ' + g.note).toLowerCase() },
        h('div', { class: 'vbref-mem-line' }, h('code', { class: 'vbref-name', translate: 'no', text: g.code })),
        h('div', { class: 'vbref-note', text: g.note }))))));
    if (catalog.unavailable && catalog.unavailable.length) {
      parts.push(section(N('Not available'), false, catalog.unavailable.map((o) =>
        member({ name: o.name, kind: '', gap: o.code, note: o.note }))));
    }
    for (const obj of catalog.objects || []) {
      const nodes = [];
      if (obj.note) nodes.push(para(obj.note));
      for (const mem of obj.members || []) nodes.push(member(mem));
      parts.push(section(obj.name, false, nodes, true));
    }
    if (catalog.constantGroups && catalog.constantGroups.length) {
      const nodes = [];
      for (const g of catalog.constantGroups) {
        nodes.push(h('div', { class: 'tp-h', text: g.title }));
        for (const [name, value] of g.items) {
          const line = name + '  ' + value;
          nodes.push(h('div', { class: 'vbref-const', translate: 'no', 'data-hay': line.toLowerCase(), text: line }));
        }
      }
      parts.push(section(N('Constants'), false, nodes));
    }
    if (catalog.example) {
      const nodes = [];
      if (catalog.exampleNote) nodes.push(para(catalog.exampleNote));
      nodes.push(h('pre', { class: 'vbref-ex', translate: 'no', 'data-hay': catalog.example.toLowerCase(), text: catalog.example }));
      parts.push(section(N('Example'), true, nodes));
    }
    root.append(...parts);
    const apply = () => {
      const needle = q.value.trim().toLowerCase();
      let shown = 0;
      for (const sec of root.querySelectorAll('.vbref-sec')) {
        const title = (sec.querySelector('.vbref-h span') || {}).textContent || '';
        const titleHit = !!needle && title.toLowerCase().includes(needle);
        let any = titleHit;
        for (const row of sec.querySelectorAll('[data-hay]')) {
          const hit = !needle || titleHit || row.dataset.hay.includes(needle);
          row.hidden = !hit;
          if (hit && needle) any = true;
        }
        sec.hidden = !!needle && !any;
        if (needle && any) {
          const btn = sec.querySelector('.vbref-h');
          const body = sec.querySelector('.vbref-body');
          btn.setAttribute('aria-expanded', 'true');
          body.hidden = false;
        }
        if (!sec.hidden) shown++;
      }
      empty.hidden = !needle || shown > 0;
    };
    q.addEventListener('input', apply);
    q.addEventListener('keydown', (e) => e.stopPropagation());
    b.append(
      q,
      h('p', { class: 'vbref-lead', text: 'This page lists what this app runs. It is not the full Microsoft VBA reference.' }),
      root,
      empty,
      h('p', { class: 'tp-note', text: L.t('When this API changes, update {0} and this reference together.', catalog.spec) }));
  }

  let listed = null;

  /* The note beside a member, as written in the reference. Not UI text. */
  function purpose(typeName, memberName) {
    const typeKey = String(typeName || '').toLowerCase();
    const memberKey = String(memberName || '').toLowerCase();
    if (!listed || !typeKey || !memberKey) return '';
    const objects = listed.objects || [];
    for (let i = 0; i < objects.length; i++) {
      if (String(objects[i].name || '').toLowerCase() !== typeKey) continue;
      const members = objects[i].members || [];
      for (let j = 0; j < members.length; j++) {
        const names = String(members[j].name || '').split(',');
        for (let k = 0; k < names.length; k++) {
          if (names[k].trim().toLowerCase() === memberKey) return members[j].note || '';
        }
      }
    }
    return '';
  }

  function mount(catalog) {
    if (catalog) listed = catalog;
    const task = L.panes && L.panes.task;
    if (!task || !listed) return;
    task.register({
      id: 'vb-dom',
      title: 'Visual Basic DOM API Reference',
      render(body) { render(body, listed); },
    });
  }

  L.vbDomRef = { show, mount, purpose };
})();
