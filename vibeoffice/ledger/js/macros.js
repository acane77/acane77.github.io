/* Ledger macros: VB6 source stored in the workbook, compiled together so they can call each other.
   Office VBA (vbaProject.bin) is a different thing and stays unread.
   Current Document is this workbook. All templates and documents is this browser, not the file. */
(function () {
  'use strict';
  const L = window.L;
  const NS = 'urn:vibeoffice:macros';
  const ITEM_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml';
  const PROPS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps';
  const PROPS_CT = 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml';
  const MAIN = 'xl/workbook.xml';
  const esc = (s) => (L.xmlTree ? L.xmlTree.esc(s) : String(s));
  const escText = (s) => (L.xmlTree ? L.xmlTree.escText(s) : String(s));
  const ALL_KEY = 'vibeoffice.ledger.allMacros';
  const N = (s) => (L.N ? L.N(s) : s);

  let loadedBook = null;
  let loadedSig = '';
  let loadError = null;
  let ticket = 0;

  function book() {
    const app = L.app;
    if (app && app.books && app.books.length) {
      const b = app.books[app.cur] || app.books[0];
      return b && b.wb;
    }
    return null;
  }
  function list(d) { return (d && d.macros) || []; }
  function signature(d) { return list(d).map((m) => m.name + '\n' + m.source).join('\n\0'); }
  function author() { return (L.store && L.store.get('userName', 'Ledger User')) || 'Ledger User'; }
  function today() {
    const d = new Date();
    const p = (n) => String(n).padStart(2, '0');
    return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
  }
  function titleWord(word) { return word.charAt(0).toUpperCase() + word.slice(1); }
  function stub(name, created, who) {
    return titleWord('public') + ' ' + titleWord('sub') + ' ' + name + '()\n' +
      "'\n' " + name + ' ' + titleWord('macro') + '\n' +
      "' " + titleWord('macro') + ' created ' + created + ' by ' + who + '\n' +
      "'\n" + titleWord('end') + ' ' + titleWord('sub') + '\n';
  }
  const isDigit = (c) => c >= 48 && c <= 57;
  const isIdentifierStart = (c) => c === 95 || c >= 65 && c <= 90 || c >= 97 && c <= 122 || c >= 128 && c <= 65535;
  const isIdentifierPart = (c) => isIdentifierStart(c) || isDigit(c);
  const isSpace = (c) => c === 32 || c >= 9 && c <= 13 || c > 127 && /\s/u.test(String.fromCharCode(c));
  const isTypeSuffix = (c) => c === '$' || c === '%' || c === '&' || c === '!' || c === '#' || c === '@';
  function identifierEnd(text, start) {
    let end = start + 1;
    while (isIdentifierPart(text.charCodeAt(end))) end++;
    if (isTypeSuffix(text[end]) && !(text[end] === '!' && (isIdentifierStart(text.charCodeAt(end + 1)) || text[end + 1] === '['))) end++;
    return end;
  }
  function vbIdentifier(text) {
    const s = String(text == null ? '' : text);
    if (!s) return false;
    if (s[0] === '[') {
      const end = s.indexOf(']');
      if (end < 0 || !s.slice(1, end)) return false;
      return end + 1 === s.length;
    }
    const c0 = s.charCodeAt(0);
    if (isSpace(c0) || !isIdentifierStart(c0)) return false;
    return identifierEnd(s, 0) === s.length;
  }
  function find(d, name) {
    const t = String(name || '').toLowerCase();
    return list(d).find((m) => m.name.toLowerCase() === t) || null;
  }
  /* Macros for every workbook opened in this copy of Ledger. Not written into the file. */
  function allList() {
    try {
      if (typeof localStorage === 'undefined') return [];
      const raw = localStorage.getItem(ALL_KEY);
      const parsed = raw ? JSON.parse(raw) : [];
      if (!Array.isArray(parsed)) return [];
      return parsed.filter((m) => m && typeof m.name === 'string' && m.name).map((m) => ({
        name: m.name,
        source: typeof m.source === 'string' ? m.source : '',
        created: typeof m.created === 'string' ? m.created : '',
        author: typeof m.author === 'string' ? m.author : '',
      }));
    } catch (e) { return []; }
  }
  function allWrite(macros) {
    try {
      if (typeof localStorage === 'undefined') return false;
      localStorage.setItem(ALL_KEY, JSON.stringify(macros));
      return true;
    } catch (e) { return false; }
  }
  function findAll(name) {
    const t = String(name || '').toLowerCase();
    return allList().find((m) => m.name.toLowerCase() === t) || null;
  }
  function describe(name, location) {
    const m = location === 'all' ? findAll(name) : find(book(), name);
    if (!m) return '';
    return L.t('Macro created {0} by {1}', m.created || '', m.author || '');
  }

  function serialize(macros) {
    const body = macros.map((m) =>
      '<macro name="' + esc(m.name) + '" created="' + esc(m.created || '') + '" author="' + esc(m.author || '') + '">' +
      '<source>' + escText(m.source || '') + '</source></macro>').join('');
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<macros xmlns="' + NS + '">' + body + '</macros>';
  }
  function parse(text) {
    if (!text || !L.xmlTree) return null;
    let root;
    try { root = L.xmlTree.parse(text); } catch (e) { return null; }
    if (!root || root.localName !== 'macros' || root.namespaceURI !== NS) return null;
    const out = [];
    for (const el of root.children || []) {
      if (el.localName !== 'macro') continue;
      const name = el.getAttribute('name') || '';
      if (!name) continue;
      const source = (el.children || []).find((c) => c.localName === 'source');
      out.push({ name: name, created: el.getAttribute('created') || '', author: el.getAttribute('author') || '', source: source ? source.textContent : '' });
    }
    return out;
  }
  function relKind(type, kind) { return new RegExp('/' + kind + '$').test(type || ''); }
  function keptOf(d) { return d && d.extra && d.extra.vibeMacros; }

  function adopt(d) {
    if (!d.macros) d.macros = [];
    const pkg = d.pkg;
    if (!pkg || !pkg.main) return;
    for (const rel of pkg.rels(pkg.main)) {
      if (rel.external || !relKind(rel.type, 'customXml')) continue;
      const text = pkg.text(rel.part);
      const macros = text && parse(text);
      if (!macros) continue;
      d.macros = macros;
      const propsRel = pkg.rels(rel.part).find((r) => !r.external && relKind(r.type, 'customXmlProps'));
      let itemID = '';
      if (propsRel && pkg.text(propsRel.part)) {
        try {
          const tree = L.xmlTree.parse(pkg.text(propsRel.part));
          const id = (tree.attributes || []).find((a) => a.localName === 'itemID');
          itemID = id ? id.value : '';
        } catch (e) { /* the item still loads; a new id is written on save */ }
      }
      d.extra = d.extra || {};
      d.extra.vibeMacros = { item: rel.part, props: propsRel ? propsRel.part : '', itemID: itemID };
      return;
    }
  }
  function ownsRel(d, rel) {
    const item = keptOf(d) && keptOf(d).item;
    return !!(item && rel && rel.part === item);
  }
  function guid() {
    const id = (typeof crypto !== 'undefined' && crypto.randomUUID ? crypto.randomUUID() : '00000000-0000-4000-8000-000000000000').toUpperCase();
    return '{' + id + '}';
  }
  function propsXml(itemID) {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<ds:datastoreItem ds:itemID="' + esc(itemID) + '" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">' +
      '<ds:schemaRefs><ds:schemaRef ds:uri="' + NS + '"/></ds:schemaRefs></ds:datastoreItem>';
  }
  function freeName(writer, fallback) {
    if (!writer.names.has(fallback.toLowerCase())) { writer.names.add(fallback.toLowerCase()); return fallback; }
    const slash = fallback.lastIndexOf('/'), dot = fallback.lastIndexOf('.');
    return writer.name(fallback.slice(0, Math.max(0, slash)), fallback.slice(slash + 1, dot).replace(/\d+$/, ''), fallback.slice(dot + 1));
  }
  /* A new workbook has no package. claim() needs the original part. */
  function inPackage(writer, part) {
    return !!(part && writer.pkg && writer.pkg.has && writer.pkg.has(part));
  }
  function finish(d, pack) {
    const writer = pack.writer;
    const macros = list(d);
    const kept = keptOf(d);
    const owner = pack.part ? pack.part(MAIN) : MAIN;
    if (!macros.length) {
      if (kept && inPackage(writer, kept.item)) {
        if (inPackage(writer, kept.props)) writer.claim(kept.props, 'regenerated', kept.props);
        writer.omit(kept.item, L.t('The macros were removed from this document.'), { notify: false });
      }
      return;
    }
    let item, props, itemID;
    if (kept && inPackage(writer, kept.item)) {
      item = writer.claim(kept.item, 'regenerated', kept.item);
      props = inPackage(writer, kept.props) ? writer.claim(kept.props, 'regenerated', kept.props) : freeName(writer, 'customXml/itemPropsVibeMacros.xml');
      itemID = kept.itemID || guid();
    } else {
      item = freeName(writer, 'customXml/itemVibeMacros.xml');
      props = freeName(writer, 'customXml/itemPropsVibeMacros.xml');
      itemID = (kept && kept.itemID) || guid();
      d.extra = d.extra || {};
      d.extra.vibeMacros = { item: item, props: props, itemID: itemID };
    }
    writer.put(item, serialize(macros), 'application/xml');
    writer.put(props, propsXml(itemID), PROPS_CT);
    writer.rels(item).add(PROPS_REL, L.opc.relative(item, props));
    writer.rels(owner).add(ITEM_REL, L.opc.relative(owner, item));
  }

  function copies(macros) { return macros.map((m) => ({ name: m.name, source: m.source, created: m.created, author: m.author })); }
  function change(label, fn) {
    const d = book();
    if (!d || !L.ops) return false;
    if (!d.macros) d.macros = [];
    const before = copies(list(d));
    L.ops.tx(d, label, () => {
      const next = copies(before);
      fn(next);
      d.macros = next;
      const saved = copies(next);
      d.undo.op(
        () => { d.macros = copies(before); sync(); },
        () => { d.macros = copies(saved); sync(); });
    });
    sync();
    return true;
  }
  function allCommit(name, source, meta) {
    const macros = allList();
    const key = String(name).toLowerCase();
    const i = macros.findIndex((m) => m.name.toLowerCase() === key);
    const existing = i >= 0 ? macros[i] : null;
    const rec = {
      name: String(name),
      source: String(source),
      created: (meta && meta.created) || (existing && existing.created) || today(),
      author: (meta && meta.author) || (existing && existing.author) || author(),
    };
    if (i >= 0) macros[i] = rec; else macros.push(rec);
    return allWrite(macros);
  }
  function allRemove(name) {
    const macros = allList();
    const i = macros.findIndex((m) => m.name.toLowerCase() === String(name).toLowerCase());
    if (i < 0) return false;
    macros.splice(i, 1);
    return allWrite(macros);
  }
  function commit(name, source, meta) {
    if (meta && meta.store === 'all') return allCommit(name, source, meta);
    const d = book();
    if (!d) return false;
    const existing = find(d, name);
    const created = (meta && meta.created) || (existing && existing.created) || today();
    const who = (meta && meta.author) || (existing && existing.author) || author();
    return change(N('Save Macro'), (next) => {
      const i = next.findIndex((m) => m.name.toLowerCase() === String(name).toLowerCase());
      const rec = { name: name, source: source, created: created, author: who };
      if (i >= 0) next[i] = rec; else next.push(rec);
    });
  }
  function remove(name) {
    return change(N('Delete Macro'), (next) => {
      const i = next.findIndex((m) => m.name.toLowerCase() === String(name).toLowerCase());
      if (i >= 0) next.splice(i, 1);
    });
  }

  async function sync() {
    const d = book();
    const macros = list(d);
    const sig = d ? signature(d) : '';
    if (d === loadedBook && sig === loadedSig) return;
    const n = ++ticket;
    if (!macros.length) {
      const host = L.vb && L.vb.current && L.vb.current();
      if (host) host.unloadMacros();
      if (n !== ticket) return;
      loadedBook = d; loadedSig = sig; loadError = null;
      return;
    }
    if (!L.vb || !L.vb.session) {
      loadError = new Error(L.t('The VB runtime is not running'));
      loadedBook = d; loadedSig = sig;
      return;
    }
    try {
      const host = await L.vb.session();
      if (n !== ticket) return;
      await host.loadMacros(macros);
      if (n !== ticket) return;
      loadedBook = d; loadedSig = sig; loadError = null;
    } catch (e) {
      if (n !== ticket || (e && e.code === 'vb-busy')) return;
      const host = L.vb.current && L.vb.current();
      if (host) host.unloadMacros();
      loadedBook = d; loadedSig = sig;
      loadError = new Error(L.t('The macros in this document could not be compiled. {0}', (e && e.message) || e));
    }
  }
  async function run(name, location) {
    if (location === 'all') {
      if (!L.vb || !L.vb.runDraft) throw new Error(L.t('The VB runtime is not running'));
      return L.vb.runDraft(allList(), name);
    }
    await sync();
    if (loadError) throw loadError;
    if (!L.vb || !L.vb.runSaved) throw new Error(L.t('The VB runtime is not running'));
    return L.vb.runSaved(name);
  }
  function draftModules(name, source, store) {
    const pool = store === 'all' ? allList() : list(book());
    const others = pool.filter((m) => m.name.toLowerCase() !== String(name).toLowerCase());
    others.push({ name: name, source: source });
    return others;
  }
  async function runDraft(name, source, store) {
    if (!L.vb || !L.vb.runDraft) throw new Error(L.t('The VB runtime is not running'));
    return L.vb.runDraft(draftModules(name, source, store), name);
  }
  function openEditor(macro) {
    if (L.domConsole && L.domConsole.openScript) L.domConsole.openScript({ macro: macro });
  }

  /* The notice comes first. OK opens the Macros window. The title's close button does not.
     Never show this message again is unchecked until the user checks it and clicks OK.
     The flag is ledger.macroNotice, so Quire's and Lectern's notices stay their own. */
  async function open() {
    if (L.store && !L.store.get('macroNotice', false) && L.ui && L.h) {
      let hide = false;
      const check = L.ui.check('Never show this message again', false, (v) => { hide = v; });
      check.style.marginTop = '8px';
      const i = await L.ui.dialog({
        title: 'Notice on Macros in Ledger',
        width: 480,
        body: L.h('div', {},
          L.h('p', { text: 'Macros created by Microsoft Excel are not executed in Ledger. A VBA project in a .xlsm file is kept in the package and is not run.' }),
          L.h('p', { text: 'Additionally, macros created in Ledger cannot be loaded or executed by Microsoft Excel, nor can macros created in Excel be loaded or executed by Ledger.' }),
          L.h('p', { text: 'The purpose of introducing Visual Basic-based macros in Ledger is to automate workbook editing using existing VB6 code and to provide MCP (Model Context Protocol) capabilities for AI agents, rather than to achieve compatibility with Microsoft Excel\'s .xlsm format or VBA projects.' }),
          check),
        buttons: [{ label: 'OK', primary: true }],
      }).done;
      if (i !== 0) return;
      if (hide) L.store.set('macroNotice', true);
    }
    if (L.openMacroDialog) L.openMacroDialog();
  }

  L.macros = {
    adopt: adopt,
    ownsRel: ownsRel,
    finish: finish,
    parse: parse,
    serialize: serialize,
    sync: sync,
    commit: commit,
    remove: remove,
    run: run,
    runDraft: runDraft,
    draftModules: draftModules,
    describe: describe,
    stub: stub,
    vbIdentifier: vbIdentifier,
    open: open,
  };
  L.macroHost = {
    enabled: true,
    names: (location) => (location === 'all' ? allList() : list(book())).map((m) => m.name),
    describe: describe,
    run: run,
    edit: (name, location) => {
      if (location === 'all') {
        const m = findAll(name);
        if (m) openEditor({ name: m.name, source: m.source, created: m.created, author: m.author, store: 'all' });
        return;
      }
      const m = find(book(), name);
      if (m) openEditor(m);
    },
    create: (name, location) => {
      const created = today(), who = author();
      const macro = { name: name, source: stub(name, created, who), created: created, author: who, isNew: true };
      if (location === 'all') macro.store = 'all';
      openEditor(macro);
    },
    remove: (name, location) => (location === 'all' ? allRemove(name) : remove(name)),
  };
  if (L.bus) {
    L.bus.on('book', () => { sync(); });
    L.bus.on('changed', () => { sync(); });
  }
  if (book()) sync();
})();
