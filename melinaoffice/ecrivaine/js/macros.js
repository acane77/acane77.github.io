/* Ecrivaine macros: VB6 source stored in the document, compiled together so they can call each other.
   Office VBA (vbaProject.bin) is a different thing and stays unread. */
(function () {
  'use strict';
  const L = window.L;
  const NS = 'https://acane.net/melinaoffice/macros';
  const ITEM_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml';
  const PROPS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps';
  const PROPS_CT = 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml';
  const esc = (s) => (L.xmlTree ? L.xmlTree.esc(s) : String(s));
  const escText = (s) => (L.xmlTree ? L.xmlTree.escText(s) : String(s));
  const ALL_KEY = 'melinaoffice.ecrivaine.allMacros';

  let loadedDoc = null;
  let loadedSig = '';
  let loadError = null;
  let ticket = 0;

  function doc() { return L.D && L.D.doc; }
  function list(d) { return (d && d.macros) || []; }
  function signature(d) { return list(d).map((m) => m.name + '\n' + m.source).join('\n\0'); }
  function author() { return L.store.get('userName', 'Ecrivaine User'); }
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
  /* One identifier, as the Visual Basic lexer reads one
     (third_party/VB6/src/language/source-scanner.js identifierEnd, and the '[' form in lexer.js).
     A letter, underscore or code point U+0080–U+FFFF starts it; those or an ASCII digit continue it;
     one type suffix may end it, except ! before another name or '['. A bracketed name is the other form.
     The whole string is that one token. */
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
  /* Macros for every document opened in this copy of Ecrivaine. Not written into the file. */
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
    const m = location === 'all' ? findAll(name) : find(doc(), name);
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
      d.keep = d.keep || {};
      d.keep.vibeMacros = { item: rel.part, props: propsRel ? propsRel.part : '', itemID: itemID };
      return;
    }
  }
  function ownsRel(d, rel) {
    const item = d && d.keep && d.keep.vibeMacros && d.keep.vibeMacros.item;
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
  function finish(d, pack) {
    const writer = pack.writer;
    const macros = list(d);
    const kept = d.keep && d.keep.vibeMacros;
    if (!macros.length) {
      if (kept && kept.item) {
        if (kept.props) writer.claim(kept.props, 'regenerated', kept.props);
        writer.omit(kept.item, L.t('The macros were removed from this document.'), { notify: false });
      }
      return;
    }
    let item, props, itemID;
    if (kept && kept.item) {
      item = writer.claim(kept.item, 'regenerated', kept.item);
      props = kept.props ? writer.claim(kept.props, 'regenerated', kept.props) : freeName(writer, 'customXml/itemPropsMacros.xml');
      itemID = kept.itemID || guid();
    } else {
      item = freeName(writer, 'customXml/itemMacros.xml');
      props = freeName(writer, 'customXml/itemPropsMacros.xml');
      itemID = guid();
      d.keep = d.keep || {};
      d.keep.vibeMacros = { item: item, props: props, itemID: itemID };
    }
    writer.put(item, serialize(macros), 'application/xml');
    writer.put(props, propsXml(itemID), PROPS_CT);
    writer.rels(item).add(PROPS_REL, L.opc.relative(item, props));
    writer.rels('word/document.xml').add(ITEM_REL, L.opc.relative('word/document.xml', item));
  }

  function change(label, fn) {
    const d = doc();
    if (!d || !L.D) return false;
    L.D.tx(label, () => {
      L.D.touchKey(d, 'macros');
      const next = list(d).slice();
      fn(next);
      d.macros = next;
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
    const d = doc();
    if (!d) return false;
    const existing = find(d, name);
    const created = (meta && meta.created) || (existing && existing.created) || today();
    const who = (meta && meta.author) || (existing && existing.author) || author();
    return change('Save Macro', (next) => {
      const i = next.findIndex((m) => m.name.toLowerCase() === String(name).toLowerCase());
      const rec = { name: name, source: source, created: created, author: who };
      if (i >= 0) next[i] = rec; else next.push(rec);
    });
  }
  function remove(name) {
    return change('Delete Macro', (next) => {
      const i = next.findIndex((m) => m.name.toLowerCase() === String(name).toLowerCase());
      if (i >= 0) next.splice(i, 1);
    });
  }

  async function sync() {
    const d = doc();
    const macros = list(d);
    const sig = d ? signature(d) : '';
    if (d === loadedDoc && sig === loadedSig) return;
    const n = ++ticket;
    if (!macros.length) {
      const host = L.vb && L.vb.current && L.vb.current();
      if (host) host.unloadMacros();
      if (n !== ticket) return;
      loadedDoc = d; loadedSig = sig; loadError = null;
      return;
    }
    if (!L.vb || !L.vb.session) {
      loadError = new Error(L.t('The VB runtime is not running'));
      loadedDoc = d; loadedSig = sig;
      return;
    }
    try {
      const host = await L.vb.session();
      if (n !== ticket) return;
      await host.loadMacros(macros);
      if (n !== ticket) return;
      loadedDoc = d; loadedSig = sig; loadError = null;
    } catch (e) {
      if (n !== ticket || (e && e.code === 'vb-busy')) return;
      const host = L.vb.current && L.vb.current();
      if (host) host.unloadMacros();
      loadedDoc = d; loadedSig = sig;
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
    const pool = store === 'all' ? allList() : list(doc());
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
  };
  L.macroHost = {
    enabled: true,
    names: (location) => (location === 'all' ? allList() : list(doc())).map((m) => m.name),
    describe: describe,
    run: run,
    edit: (name, location) => {
      if (location === 'all') {
        const m = findAll(name);
        if (m) openEditor({ name: m.name, source: m.source, created: m.created, author: m.author, store: 'all' });
        return;
      }
      const m = find(doc(), name);
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
    L.bus.on('doc-loaded', () => { sync(); });
    L.bus.on('history', () => { sync(); });
  }
  if (L.D && L.D.doc) sync();
})();
