/* Host abilities and the three audit tiers. A member runs only when this host implements
 * that ability. The tier says whether an agent script must be allowed first.
 * AlwaysAllow needs no VibeOffice confirmation. RequestApproval asks when an agent started
 * the run. AlwaysDeny refuses. A file outside the workspace is RequestApproval for that
 * call. Docs: docs/dom_api_security.md. */
(function (root) {
  'use strict';
  const L = root.L = root.L || {};
  const VO = root.VO = root.VO || {};
  const COMMON = ['Path', 'Save', 'SaveAs', 'Open', 'Print', 'FixedFormat', 'Hyperlink', 'Picture', 'FileDialog', 'Dir', 'Shell'];
  const EXTRA = { quire: ['InsertFile'], ledger: ['SheetSaveAs'], lectern: ['Export'] };
  const granted = new Set();
  const dirState = { list: null, i: 0 };

  const QUIRE_OFFICE = {
    174: 'font', 175: 'paragraph', 179: 'tabs', 189: 'borders', 824: 'bullets', 177: 'columns',
    488: 'dropCap', 322: 'changeCase', 180: 'style', 960: 'formatObject', 178: 'pageSetup',
    88: 'print', 294: 'pageNumbers', 159: 'breakDlg', 165: 'dateTime', 166: 'field', 162: 'symbol',
    370: 'footnote', 357: 'caption', 367: 'crossRef', 168: 'bookmark', 925: 'hyperlink',
    172: 'insertObject', 129: 'insertTable', 130: 'insertCells', 133: 'deleteCells', 137: 'splitCells',
    563: 'tableAutoFormat', 199: 'sort', 348: 'formula', 127: 'textToTable', 128: 'tableToText',
    228: 'wordCount', 195: 'hyphenation', 607: 'envelopes', 821: 'letterWizard', 87: 'templates',
    222: 'organizer', 378: 'autocorrect', 974: 'options', 188: 'language', 86: 'properties', 111: 'pasteSpecial',
  };
  const LEDGER_OFFICE = {
    150: 'formatCells', 127: 'rowHeight', 47: 'colWidth', 256: 'zoom', 7: 'pageSetup', 8: 'print',
    64: 'find', 130: 'findReplace', 63: 'goTo', 132: 'goToSpecial', 53: 'pasteSpecial', 55: 'insertCells',
    54: 'deleteCells', 39: 'sort', 40: 'series', 450: 'insertFunction', 61: 'defineName', 62: 'createNames',
    133: 'applyNames', 58: 'pasteName', 398: 'subtotals', 422: 'textToColumns', 370: 'advancedFilter',
    198: 'goalSeek', 305: 'scenarios', 307: 'scenarios', 308: 'scenarios', 311: 'scenarios',
    41: 'dataTable', 191: 'consolidate', 312: 'pivotTable', 474: 'properties', 356: 'options',
    485: 'autoCorrect', 596: 'hyperlink', 154: 'comment', 417: 'protectWorkbook', 28: 'protectSheet',
  };

  function host() { return VO.host || {}; }
  function vscode() { return host().kind === 'vscode'; }
  function appId() { return String(L.APP_ID || '').toLowerCase(); }

  function documentPath() {
    const h = host();
    return typeof h.documentPath === 'function' ? (h.documentPath() || '') : (h.documentPath || '');
  }
  function documentDir() {
    const h = host();
    if (typeof h.documentDir === 'function') return h.documentDir() || '';
    return h.documentDir || '';
  }

  function abilityHit(name) {
    const want = String(name || '').toLowerCase();
    return COMMON.concat(EXTRA[appId()] || []).find(function (item) { return item.toLowerCase() === want; }) || '';
  }

  function has(name) {
    if (!vscode()) return false;
    const hit = abilityHit(name);
    if (!hit) return false;
    if (hit === 'Path' || hit === 'Save' || hit === 'Print' || hit === 'FixedFormat') return !!documentPath();
    return true;
  }

  /* Why a host member cannot run here. Empty when it can. */
  function why(name) {
    if (has(name)) return '';
    if (!vscode()) return 'platform';
    const hit = abilityHit(name);
    if (!hit) return 'app';
    if (hit === 'Path' || hit === 'Save' || hit === 'Print' || hit === 'FixedFormat') return 'path';
    return 'platform';
  }

  function beginScript() {
    granted.clear();
    dirState.list = null;
    dirState.i = 0;
  }

  function origin() {
    if (VO.audit && VO.audit.origin) return VO.audit.origin;
    if (host().vbsAgent) return 'agent';
    return 'user';
  }

  function denied(code, message) {
    return { ok: false, code: code, message: message };
  }

  async function outside(path) {
    if (!path) return false;
    const h = host();
    if (typeof h.pathOutside !== 'function') return false;
    try { return !!(await h.pathOutside(path)); }
    catch (e) { return true; }
  }

  async function ask(opts) {
    if (VO.audit && typeof VO.audit.decide === 'function') return VO.audit.decide(opts);
    if (typeof host().approve === 'function') return host().approve(opts);
    if (origin() === 'agent') return denied('E-APPROVAL', 'The editor did not allow this operation.');
    return { ok: true };
  }

  async function gate(opts) {
    opts = opts || {};
    const tier0 = opts.tier || 'AlwaysAllow';
    const ability = opts.ability || opts.member || '';
    if (tier0 === 'AlwaysDeny') return denied('E-DENIED', 'The editor refused this operation.');
    if (opts.fileDialog) return { ok: true };
    let tier = tier0;
    if (opts.path && await outside(opts.path)) tier = 'RequestApproval';
    if (tier !== 'RequestApproval') return { ok: true };
    if (granted.has(ability)) return { ok: true };
    if (origin() !== 'agent') {
      granted.add(ability);
      return { ok: true };
    }
    const answer = await ask(opts);
    if (answer && answer.ok) {
      granted.add(ability);
      return { ok: true };
    }
    return denied('E-APPROVAL', (answer && answer.message) || 'The editor did not allow this operation.');
  }

  function must(tier, ability, detail, path) {
    return gate({ tier: tier, ability: ability, detail: detail, path: path }).then(function (r) {
      if (r.ok) return;
      const e = new Error(r.message);
      e.code = r.code;
      throw e;
    });
  }

  function rejectExtra(args, from, names) {
    for (let i = from; i < args.length; i++) {
      const v = args[i];
      if (v == null || v === false || v === 0 || v === '') continue;
      const e = new Error('The parameter ' + (names[i] || ('argument ' + (i + 1))) + ' is not available.');
      e.code = 'error';
      throw e;
    }
  }

  function pdfBeside() {
    const p = documentPath();
    if (!p) return '';
    const slash = Math.max(p.lastIndexOf('/'), p.lastIndexOf('\\'));
    const base = slash >= 0 ? p.slice(slash + 1) : p;
    const dot = base.lastIndexOf('.');
    const stem = dot > 0 ? base.slice(0, dot) : base;
    const dir = slash >= 0 ? p.slice(0, slash) : '';
    const sep = p.indexOf('\\') >= 0 && p.indexOf('/') < 0 ? '\\' : '/';
    return (dir ? dir + sep : '') + stem + '.pdf';
  }

  function joinPath(dir, name) {
    const sep = String(dir).indexOf('\\') >= 0 && String(dir).indexOf('/') < 0 ? '\\' : '/';
    return String(dir).replace(/[\\/]+$/, '') + sep + name;
  }

  function parentOf(path) {
    const slash = Math.max(String(path).lastIndexOf('/'), String(path).lastIndexOf('\\'));
    return slash >= 0 ? String(path).slice(0, slash) : String(path);
  }

  function fromB64(s) {
    const bin = atob(s || '');
    const out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }

  async function writeBlob(path, blob, ability, openAfter) {
    const h = host();
    if (!h || typeof h.writeFile !== 'function') {
      const e = new Error('This session does not have that host ability (G-HOST)');
      e.code = 'G-HOST';
      throw e;
    }
    const name = String(path).split(/[/\\]/).pop();
    const body = blob instanceof Blob ? blob : new Blob([blob]);
    await h.writeFile(name, body, { purpose: 'host', path: path, ability: ability, open: !!openAfter });
  }

  async function readBytes(filePath, ability, tier) {
    await must(tier || 'RequestApproval', ability, 'read ' + filePath, filePath);
    const h = host();
    if (!h || typeof h.request !== 'function') {
      const e = new Error('This session does not have that host ability (G-HOST)');
      e.code = 'G-HOST';
      throw e;
    }
    const r = await h.request({ type: 'host-read', path: String(filePath), ability: ability });
    if (!r || r.ok === false) {
      const e = new Error((r && (r.error || r.message)) || 'The file could not be read.');
      e.code = (r && r.code) || 'error';
      throw e;
    }
    return fromB64(r.b64);
  }

  async function pick(kind, multi) {
    const h = host();
    if (!h || typeof h.request !== 'function') {
      const e = new Error('FileDialog is not available (G-HOST)');
      e.code = 'G-HOST';
      throw e;
    }
    const r = await h.request({ type: 'host-pick', kind: kind || 'open', multi: !!multi });
    if (!r || r.ok === false) {
      const e = new Error((r && (r.error || r.message)) || 'The file could not be chosen.');
      e.code = (r && r.code) || 'error';
      throw e;
    }
    return r.paths || [];
  }

  async function openPath(filePath) {
    await must('RequestApproval', 'Open', 'open ' + filePath, filePath);
    const h = host();
    const r = await h.request({ type: 'host-open', path: String(filePath) });
    if (!r || r.ok === false) {
      const e = new Error((r && (r.error || r.message)) || 'The file could not be opened.');
      e.code = (r && r.code) || 'error';
      throw e;
    }
  }

  async function saveCopy(fileName, makeBlob, ability) {
    if (!fileName) {
      const e = new Error('The parameter FileName is not available.');
      e.code = 'error';
      throw e;
    }
    const token = ability || 'SaveAs';
    await must('RequestApproval', token, 'save as ' + fileName, fileName);
    const blob = await makeBlob();
    await writeBlob(String(fileName), blob, token, true);
  }

  async function printPdf(makeBlob) {
    const path = pdfBeside();
    await must('AlwaysAllow', 'Print', 'write a PDF beside the file', path);
    const blob = await makeBlob();
    if (!blob) {
      const e = new Error('There is nothing to print.');
      e.code = 'error';
      throw e;
    }
    await writeBlob(path, blob, 'Print', false);
  }

  function isXps(type) {
    if (type == null || type === '') return false;
    if (typeof type === 'string' && /xps/i.test(type)) return true;
    const n = type | 0;
    const app = appId();
    if (app === 'quire') return n === 18;
    return n === 1;
  }
  function isPdf(type) {
    if (type == null || type === '') return true;
    if (typeof type === 'string' && /pdf/i.test(type)) return true;
    const n = type | 0;
    const app = appId();
    if (app === 'quire') return n === 17;
    if (app === 'lectern') return n === 2;
    return n === 0;
  }

  async function exportPdf(type, fileName, makeBlob) {
    if (isXps(type)) {
      const e = new Error('XPS is not available.');
      e.code = 'error';
      throw e;
    }
    if (!isPdf(type)) {
      const e = new Error('The parameter ExportFormat is not available.');
      e.code = 'error';
      throw e;
    }
    const path = fileName || pdfBeside();
    await must('RequestApproval', 'FixedFormat', 'export a PDF to ' + path, path);
    const blob = await makeBlob();
    await writeBlob(path, blob, 'FixedFormat', false);
  }

  async function follow(address) {
    const s = String(address || '');
    if (!/^(https?:|mailto:)/i.test(s)) {
      const e = new Error('That address is not available.');
      e.code = 'error';
      throw e;
    }
    await must('AlwaysAllow', 'Hyperlink', 'open ' + s);
    if (typeof host().openUrl === 'function') await host().openUrl(s);
  }

  async function dir(pattern) {
    if (pattern == null || pattern === '') {
      if (!dirState.list) return '';
      dirState.i++;
      return dirState.list[dirState.i] || '';
    }
    const text = String(pattern);
    if (VO.audit && typeof VO.audit.virtualDir === 'function') {
      const virt = VO.audit.virtualDir(text);
      if (virt && virt.length) {
        dirState.list = virt;
        dirState.i = 0;
        return virt[0] || '';
      }
    }
    await must('RequestApproval', 'Dir', 'list ' + text, parentOf(text));
    const r = await host().request({ type: 'host-dir', pattern: text });
    if (!r || r.ok === false) {
      const e = new Error((r && (r.error || r.message)) || 'The folder could not be listed.');
      e.code = (r && r.code) || 'error';
      throw e;
    }
    dirState.list = r.names || [];
    dirState.i = 0;
    return dirState.list[0] || '';
  }

  function shell() {
    const e = new Error('The editor refused this operation.');
    e.code = 'E-DENIED';
    throw e;
  }

  function buttonValue(style, index) {
    const table = { ok: [1], okcancel: [1, 2], yn: [6, 7], ync: [6, 7, 2] };
    const row = table[style] || [1];
    return row[index] == null ? 2 : row[index];
  }

  function msgBox(prompt, buttons, title) {
    if (!L.ui || typeof L.ui.msg !== 'function') return Promise.resolve(1);
    const b = Number(buttons) || 0;
    const style = (b & 7) === 3 ? 'ync' : (b & 7) === 4 ? 'yn' : (b & 7) === 1 ? 'okcancel' : 'ok';
    const labels = { ok: ['OK'], okcancel: ['OK', 'Cancel'], yn: ['Yes', 'No'], ync: ['Yes', 'No', 'Cancel'] };
    return L.ui.msg(String(prompt == null ? '' : prompt), { title: title || (L.APP || 'VibeOffice'), buttons: labels[style], icon: style === 'ok' ? 'info' : 'question' }).then(function (index) {
      return buttonValue(style, index);
    });
  }

  function inputBox(prompt, title, def) {
    if (!L.ui || typeof L.ui.prompt !== 'function') return Promise.resolve('');
    return L.ui.prompt(String(prompt || ''), def || '', title || (L.APP || 'VibeOffice')).then(function (v) {
      return v == null ? '' : String(v);
    });
  }

  function showOffice(open) {
    if (typeof document === 'undefined' || !document.body) {
      try { if (typeof open === 'function') open(); } catch (e) { /* the page has no dialog yet */ }
      return Promise.resolve(-1);
    }
    const before = document.querySelectorAll('.dlg-overlay').length;
    let ret;
    try { ret = open(); }
    catch (e) { return Promise.reject(e); }
    if (ret && typeof ret.then === 'function') {
      return Promise.resolve(ret).then(function (v) { return v == null ? -1 : v; });
    }
    const overlays = document.querySelectorAll('.dlg-overlay');
    if (overlays.length <= before) return Promise.resolve(-1);
    const overlay = overlays[overlays.length - 1];
    return new Promise(function (resolve) {
      const obs = new MutationObserver(function () {
        if (!overlay.isConnected) { obs.disconnect(); resolve(-1); }
      });
      obs.observe(document.body, { childList: true, subtree: true });
    });
  }

  function officeMap() {
    if (appId() === 'ledger') return { fileOpen: 1, fileSave: 5, office: LEDGER_OFFICE };
    if (appId() === 'quire') return { fileOpen: 80, fileSave: 84, office: QUIRE_OFFICE };
    return null;
  }

  function callDlg(name) {
    if (name === 'findReplace' && L.dlg && L.dlg.find) return L.dlg.find('replace');
    if (name === 'find' && L.dlg && L.dlg.find) return L.dlg.find('find');
    if (name === 'protectSheet' && L.dlg && L.dlg.protectSheet) return L.dlg.protectSheet();
    if (name === 'protectWorkbook' && L.dlg && L.dlg.protectWorkbook) return L.dlg.protectWorkbook();
    if (!L.dlg || typeof L.dlg[name] !== 'function') {
      const e = new Error('That dialog is not available (G-HOST)');
      e.code = 'G-HOST';
      throw e;
    }
    return L.dlg[name]();
  }

  async function showDialog(id) {
    const map = officeMap();
    const n = id | 0;
    if (!map) {
      const e = new Error('Dialogs is not available (G-HOST)');
      e.code = 'G-HOST';
      throw e;
    }
    if (n === map.fileOpen || n === map.fileSave) {
      if (!has('FileDialog')) {
        const e = new Error('Dialogs is not available (G-HOST)');
        e.code = 'G-HOST';
        throw e;
      }
      const paths = await pick(n === map.fileSave ? 'save' : 'open', false);
      return paths.length ? -1 : 0;
    }
    const name = map.office[n];
    if (!name) {
      const e = new Error('That dialog is not available (G-HOST)');
      e.code = 'G-HOST';
      throw e;
    }
    return showOffice(function () { return callDlg(name); });
  }

  function dialogsCollection() {
    const col = function (i) { return col.Item(i); };
    col.Item = function (id) {
      return { Show: function () { return showDialog(id); } };
    };
    return col;
  }

  function fileDialog(type) {
    const obj = {
      SelectedItems: [],
      Show: async function () {
        if (!has('FileDialog')) {
          const e = new Error('FileDialog is not available (G-HOST)');
          e.code = 'G-HOST';
          throw e;
        }
        const n = type | 0;
        const kind = n === 4 ? 'folder' : n === 2 ? 'save' : 'open';
        const paths = await pick(kind, false);
        obj.SelectedItems = paths;
        return paths.length ? -1 : 0;
      },
    };
    return obj;
  }

  async function getOpenFilename(multi) {
    if (!has('FileDialog')) {
      const e = new Error('GetOpenFilename is not available (G-HOST)');
      e.code = 'G-HOST';
      throw e;
    }
    const paths = await pick('open', multi === true || multi === -1);
    if (!paths.length) return '';
    if (multi === true || multi === -1) return paths;
    return paths[0];
  }

  async function getSaveAsFilename() {
    if (!has('FileDialog')) {
      const e = new Error('GetSaveAsFilename is not available (G-HOST)');
      e.code = 'G-HOST';
      throw e;
    }
    const paths = await pick('save', false);
    return paths[0] || '';
  }

  function moduleLines(app) {
    return [
      'Public Const vbOK As Long = 1',
      'Public Const vbCancel As Long = 2',
      'Public Const vbYes As Long = 6',
      'Public Const vbNo As Long = 7',
      'Public Const vbOKOnly As Long = 0',
      'Public Const vbOKCancel As Long = 1',
      'Public Const vbYesNoCancel As Long = 3',
      'Public Const vbYesNo As Long = 4',
      'Public Function HasHostAbility(Ability)',
      '    HasHostAbility = ' + app + '.HasHostAbility(Ability)',
      'End Function',
      'Public Function MsgBox(Prompt, Optional Buttons, Optional Title)',
      '    MsgBox = ' + app + '.MsgBox(Prompt, Buttons, Title)',
      'End Function',
      'Public Function InputBox(Prompt, Optional Title, Optional Default)',
      '    InputBox = ' + app + '.InputBox(Prompt, Title, Default)',
      'End Function',
      'Public Function HostDir(Optional PathName)',
      '    HostDir = ' + app + '.HostDir(PathName)',
      'End Function',
      'Public Function HostShell(PathName, Optional WindowStyle)',
      '    HostShell = ' + app + '.HostShell(PathName)',
      'End Function',
      'Public Function Dir(Optional PathName, Optional Attributes)',
      '    Dir = ' + app + '.HostDir(PathName)',
      'End Function',
      'Public Function Shell(PathName, Optional WindowStyle)',
      '    Shell = ' + app + '.HostShell(PathName)',
      'End Function',
    ];
  }

  L.hostAudit = {
    has: has,
    why: why,
    beginScript: beginScript,
    gate: gate,
    must: must,
    rejectExtra: rejectExtra,
    documentPath: documentPath,
    documentDir: documentDir,
    pdfBeside: pdfBeside,
    joinPath: joinPath,
    writeBlob: writeBlob,
    readBytes: readBytes,
    pick: pick,
    openPath: openPath,
    saveCopy: saveCopy,
    printPdf: printPdf,
    exportPdf: exportPdf,
    follow: follow,
    dir: dir,
    shell: shell,
    msgBox: msgBox,
    inputBox: inputBox,
    showDialog: showDialog,
    dialogsCollection: dialogsCollection,
    fileDialog: fileDialog,
    getOpenFilename: getOpenFilename,
    getSaveAsFilename: getSaveAsFilename,
    moduleLines: moduleLines,
    isXps: isXps,
    isPdf: isPdf,
  };
})(typeof window !== 'undefined' ? window : globalThis);
