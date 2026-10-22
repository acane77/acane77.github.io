/* VibeOffice suite services, shared by the start page and the apps (classic script, window.VO).
 *
 * - Which app opens which file (VO.APPS, VO.appFor).
 * - The host (VO.host): browser or desktop app, the files the system asks us to open.
 * - The service worker (public/sw.js), registered from every page.
 * - Files the start page hands to an app, and the Recent Files list. Both live in IndexedDB
 *   ('vibeoffice', stores 'handoff' and 'recent') in this browser only; nothing is uploaded.
 * - Unsaved versions: while a document has unsaved changes the app's own format is written (autosave)
 *   into its Recent Files entry as a draft; saving clears it. The Start Center, the apps' File menus and
 *   task panes and the Document Recovery pane all show the one list.
 *
 * An app page is opened as  <app>/             a blank document
 *                           <app>/?new           a blank document
 *                           <app>/?template=<id>  a new document from one of the app's templates
 *                           <app>/?open=<id>      a file handed over by VO.handoff (read once, then deleted)
 *                           <app>/?recent=<key>   a Recent Files entry (&draft: its unsaved version)
 * The app reads VO.launch while it starts, then calls VO.attach(app, { open, thumb }), and reports
 * VO.opened(file) after it opens a file and VO.saved(name, blob) after it saves in its own format.
 */
(function () {
  'use strict';
  const VO = (window.VO = window.VO || {});   // keep the language choice made before the first paint
  VO.VERSION = 'dev';   // the published copy says v<yyyy>.<mm>.<dd>-<6 hex> (tools/current-tree-version.sh, stamped by the sync script)
  if (window.__VO_HOST__ && window.__VO_HOST__.appVersion) VO.VERSION = window.__VO_HOST__.appVersion;   // the desktop app's build
  /** a version's date as a number (20261011), or null for a version without one (dev); copies before
   *  v<yyyy>.<mm>.<dd>-<hex> said <yyyy>.<mm>.<hex>, day 0 */
  const versionDay = (v) => {
    const m = /^v?(\d{4})\.(\d{2})(?:\.(\d{2})-[0-9a-f]{6})?/i.exec(v || '');
    return m ? +(m[1] + m[2] + (m[3] || '00')) : null;
  };
  /** compare two versions by their dates alone: < 0 if a is older than b, 0 for the same day, > 0 if newer;
   *  a version without a date (dev) is newer than any dated one */
  VO.compareVersions = (a, b) => {
    const x = versionDay(a), y = versionDay(b);
    return x === y ? 0 : x === null ? 1 : y === null ? -1 : Math.sign(x - y);
  };
  const BASE = new URL('..', document.currentScript.src).href;   // public/
  VO.base = BASE;

  /* ------------------------------------------------------------ apps */
  VO.APPS = {
    quire: { name: 'Quire', kind: 'Document', ext: ['docx', 'docm', 'dotx', 'dotm', 'doc', 'rtf', 'htm', 'html', 'txt', 'md', 'markdown', 'zip'] },
    ledger: { name: 'Ledger', kind: 'Spreadsheet', ext: ['xlsx', 'xlsm', 'xltx', 'xltm', 'xls', 'csv', 'tsv', 'prn', 'xml'] },
    lectern: { name: 'Lectern', kind: 'Presentation', ext: ['pptx', 'pptm', 'ppsx', 'ppsm', 'potx', 'potm', 'ppt'] },
  };
  VO.ext = (name) => (/\.([^./\\]+)$/.exec(name || '') || ['', ''])[1].toLowerCase();
  /** the app for a file name, or null */
  VO.appFor = (name) => { const e = VO.ext(name); return Object.keys(VO.APPS).find((a) => VO.APPS[a].ext.includes(e)) || null; };
  VO.accept = () => Object.values(VO.APPS).flatMap((a) => a.ext.map((e) => '.' + e)).join(',');
  /** VO.url('quire', { new: true }) -> .../quire/?new */
  VO.url = (app, params) => {
    const qs = Object.keys(params || {}).map((k) => (params[k] === true ? k : k + '=' + encodeURIComponent(params[k]))).join('&');
    return new URL(app + '/', BASE).href + (qs ? '?' + qs : '');
  };

  /** the page's icon (browser tab) from an inline <svg> */
  VO.setIcon = (svg) => {
    if (!/xmlns=/.test(svg)) svg = svg.replace('<svg', '<svg xmlns="http://www.w3.org/2000/svg"');
    let link = document.querySelector('link[rel="icon"]');
    if (!link) link = document.head.appendChild(Object.assign(document.createElement('link'), { rel: 'icon' }));
    link.type = 'image/svg+xml';
    link.href = 'data:image/svg+xml,' + encodeURIComponent(svg);
  };

  /* ------------------------------------------------------------ the host */
  /** VO.host: the one interface between the pages and what runs them (docs/desktop.md, docs/vscode.md).
   *  kind 'web': a browser tab or the installed PWA; 'desktop': the native app (desktop/), which sets
   *  window.__VO_HOST__ and answers the host_* commands (desktop/src/host.rs); 'vscode': a VS Code or
   *  code-server custom editor (vscode/extension.js), which sets window.__VO_HOST__ = { kind: 'vscode' }
   *  and window.__VO_VSCODE__ from acquireVsCodeApi(). Pages use VO.host, never Tauri's, VS Code's or the
   *  system's file APIs directly. */
  const announced = window.__VO_HOST__ || null;
  const native = announced && window.__TAURI_INTERNALS__ ? announced : null;
  const inVs = !native && announced && announced.kind === 'vscode' ? announced : null;
  /* The webview host listens on this window for Ctrl+A and answers with document.execCommand('selectAll')
     (VS Code's webview page, the execCommand message). That selects the whole webview, chrome included.
     Quire then reads a selection it cannot map and snaps the caret, so the selection shows and vanishes.
     Edit ▸ Select All never presses the key, which is why the menu works. The browser has no host listener.
     Stop the key short of the host, and drop a selectAll that still arrives on the heels of that key.
     A selectAll with no key (the editor's own Edit menu) is the app's command; a text field keeps the
     browser's. */
  if (inVs) {
    let selectAllKeyAt = 0;
    document.addEventListener('keydown', (e) => {
      if (e.isComposing || e.altKey || e.shiftKey || !(e.ctrlKey || e.metaKey)) return;
      if (e.key !== 'a' && e.key !== 'A') return;
      selectAllKeyAt = Date.now();
      e.stopPropagation();
    });
    const nativeExec = document.execCommand.bind(document);
    document.execCommand = function (commandId, showUI, value) {
      if (String(commandId).toLowerCase() !== 'selectall') return nativeExec(commandId, showUI, value);
      if (Date.now() - selectAllKeyAt < 500) return true;
      const el = document.activeElement;
      if (el && (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA')) return nativeExec(commandId, showUI, value);
      const ui = window.L && window.L.ui;
      if (ui && ui.cmds && ui.cmds.selectAll) { ui.exec('selectAll'); return true; }
      return nativeExec(commandId, showUI, value);
    };
  }
  const call = (cmd, args) => window.__TAURI_INTERNALS__.invoke(cmd, args);
  const vsPost = (msg) => { const api = window.__VO_VSCODE__; if (api) api.postMessage(msg); };
  /** one round trip to the extension. The reply arrives as { type: 'reply', id }. Dialogs stay open as long as the user needs. */
  const vsRequest = (msg) => {
    const api = window.__VO_VSCODE__;
    if (!api) return Promise.reject(new Error('This page is not inside the editor.'));
    const id = Math.random().toString(36).slice(2);
    return new Promise((resolve) => {
      function on(ev) {
        const m = ev.data;
        if (!m || m.type !== 'reply' || m.id !== id) return;
        window.removeEventListener('message', on);
        resolve(m);
      }
      window.addEventListener('message', on);
      api.postMessage(Object.assign({ id }, msg));
    });
  };
  const stampFile = (file, m) => {
    if (!file || !m) return file;
    if (m.uri) { file.voUri = m.uri; documentUri = m.uri; }
    if (m.name || file.name) documentName = m.name || file.name;
    if (m.dir) openFolder = { uri: m.dir, label: m.dirLabel || '' };
    if (m.fsPath) fileFsPath = m.fsPath;
    if (m.dirLabel) fileDir = m.dirLabel;
    if (m.draft && m.draftKey) file.voDraft = m.draftKey;
    return file;
  };
  let hostLive = false;   // vscode: the workspace file finished opening, so Save may write it
  let documentUri = '';   // vscode: the workspace file this editor tab is bound to
  let documentName = '';  // vscode: that file's name, for telling an in-place save from a copy
  let openFolder = null;  // vscode: { uri, label } of that file's folder, for Save As
  let fileFsPath = '';    // vscode: the open file's path on the extension host
  let fileDir = '';       // vscode: that file's folder, as a path
  const b64 = (bytes) => {
    let s = '';
    for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, Math.min(bytes.length, i + 0x8000)));
    return btoa(s);
  };
  const fromB64 = (s) => {
    const bin = atob(s || '');
    const out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  };
  VO.host = {
    kind: native ? 'desktop' : inVs ? 'vscode' : 'web',
    platform: (native || inVs) ? (native || inVs).platform : 'web',
    version: (native || inVs) ? ((native || inVs).version || 0) : 0,
    /** whether this edition looks for its updates (updater/current.json): not the Mac App Store's, which the store updates */
    updates: !native || native.edition !== 'appstore',
    /** open a web or mail address in the system's browser: a new tab on the web, the default browser in the
     *  desktop app (which also sends every link that leaves the app there), the editor's browser in VS Code */
    documentPath() { return fileFsPath; },
    documentDir() { return fileDir; },
    vbsAgent: false,
    request(msg) { return vsRequest(msg); },
    approve(opts) {
      return vsRequest({
        type: 'host-approval',
        ability: (opts && opts.ability) || '',
        tier: (opts && opts.tier) || '',
        detail: (opts && opts.detail) || '',
        path: (opts && opts.path) || '',
      }).then((m) => ({ ok: !!(m && m.ok), code: m && m.code, message: (m && (m.error || m.message)) || '' }));
    },
    pathOutside(path) {
      return vsRequest({ type: 'host-outside', path: String(path || '') }).then((m) => !!(m && m.outside));
    },
    async openUrl(url) {
      if (native) return call('host_open_url', { url: String(url) });
      if (inVs) { vsPost({ type: 'open-url', url: String(url) }); return; }
      window.open(url, '_blank', 'noopener');
    },
    /** the whole UI's scale (1 = 100 %), or null where the page cannot set it (the web and VS Code: the
     *  browser's own zoom, Ctrl + and Ctrl −, does it); the desktop app keeps it for the next start */
    async uiScale() { return native ? call('host_ui_scale') : null; },
    async setUiScale(scale) { if (native) return call('host_set_ui_scale', { scale: +scale }); },
    /** open one of the suite's pages in a window of its own: the desktop app makes a full window (host and
     *  all); the web opens a window or tab, false if the browser blocked it; VS Code keeps this editor */
    async newWindow(url) {
      if (native) { await call('host_new_window', { url: new URL(url, location.href).href }); return true; }
      if (inVs) return false;
      return !!window.open(url, '_blank');
    },
    /** close this window if it is an app window the suite opened; false for the Start Center's (and on the web) */
    async closeWindow() { return native ? call('host_close_window') : false; },
    /** VS Code: remember the UI language and, when reload is true, load this editor again so it shows.
     *  With unsaved changes the page asks to remember only; the next editor uses it. */
    setLanguage(id, reload) { if (inVs) vsPost({ type: 'lang', lang: String(id || ''), reload: !!reload }); },
    /** Browse after a save: the Downloads folder, with this window's last saved file selected (desktop) */
    async showDownloads() { if (native) return call('host_show_downloads'); },
    /** fn(files) for the files the system asks us to open (Open with, a double-click on an associated
     *  file, the installed PWA's file handlers); the page that calls this takes them */
    onFiles(fn) {
      if (native) {
        // at start, and again when another start of the app hands this one its files (vo-files: Windows)
        const take = () => call('host_launch_files').then(async (list) => {
          if (!list.length) return;
          fn(await Promise.all(list.map(async (f) => new File([await call('host_read', { handle: f.handle })], f.name))));
        }).catch((e) => console.error('host:', e));
        take();
        window.addEventListener('vo-files', take);
      } else if (!inVs && 'launchQueue' in window) {
        window.launchQueue.setConsumer(async (params) => {
          if (params.files && params.files.length) fn(await Promise.all(params.files.map((h) => h.getFile())));
        });
      }
    },
    /** the workspace file this editor was opened on (vscode). One File, once the extension has sent it. */
    takeDocument() {
      const api = window.__VO_VSCODE__;
      if (!api) return Promise.reject(new Error('This page is not inside the editor.'));
      return new Promise((resolve, reject) => {
        let done = false, got = 0, expect = 0;
        const parts = [];
        const timer = setTimeout(() => finish(new Error('The editor did not send the file.')), 180000);
        function finish(err, file) {
          if (done) return;
          done = true;
          clearTimeout(timer);
          window.removeEventListener('message', on);
          if (err) reject(err); else resolve(file);
        }
        async function on(ev) {
          const m = ev.data;
          if (!m || done) return;
          if (m.type === 'open') {
            try {
              const res = await fetch(m.url);
              if (!res.ok) throw new Error('HTTP ' + res.status);
              finish(null, stampFile(new File([await res.blob()], m.name), m));
            } catch (e) {
              console.warn('open:', e);
              api.postMessage({ type: 'read-fallback' });
            }
          } else if (m.type === 'open-chunk') {
            if (parts[m.i] == null) { parts[m.i] = fromB64(m.b64); got++; }
            expect = m.n || 1;
            if (got >= expect) {
              let size = 0;
              for (let i = 0; i < expect; i++) size += parts[i] ? parts[i].length : 0;
              const out = new Uint8Array(size);
              let o = 0;
              for (let i = 0; i < expect; i++) if (parts[i]) { out.set(parts[i], o); o += parts[i].length; }
              finish(null, stampFile(new File([out], m.name), m));
            }
          } else if (m.type === 'blank') {
            /* New File…: keep the blank document this page already started, and write it into the new file */
            const file = stampFile(new File([], m.name || ''), m);
            if (file) file.voBlank = true;
            finish(null, file);
          } else if (m.type === 'open-error') finish(new Error(m.error || 'The file could not be read.'));
        }
        window.addEventListener('message', on);
        api.postMessage({ type: 'ready' });
      });
    },
    /** write bytes back through the editor (vscode). 'saved' is the open file; 'exported' is another file,
     *  written in the folder Save As chose or beside the open file. purpose 'backup' answers a hot-exit
     *  request and does not touch the open file. 'create' writes a new file from File ▸ New (or a Recent
     *  Files copy with no path) and opens it. 'reopen' carries unsaved bytes into a new tab for a file
     *  this editor already knows. opts.folder is a directory URI the extension already offered or the user picked. */
    writeFile(name, blob, opts) {
      const purpose = (opts && opts.purpose) || 'save';
      if (purpose === 'save' && !hostLive) return Promise.reject(new Error('The file is not open in the editor yet.'));
      const api = window.__VO_VSCODE__;
      if (!api) return Promise.reject(new Error('This page is not inside the editor.'));
      const id = (opts && opts.id) || Math.random().toString(36).slice(2);
      const uri = (opts && opts.uri) || '';
      const folder = (opts && opts.folder) || '';
      const draftKey = (opts && opts.draftKey) || '';
      const hostPath = (opts && opts.path) || '';
      const ability = (opts && opts.ability) || '';
      const openAfter = !!(opts && opts.open);
      return blob.arrayBuffer().then((buf) => new Promise((resolve, reject) => {
        const bytes = new Uint8Array(buf);
        const CHUNK = 256 * 1024;
        const n = Math.max(1, Math.ceil(bytes.length / CHUNK));
        const timer = setTimeout(() => { window.removeEventListener('message', on); reject(new Error('Saving timed out.')); }, 300000);
        function on(ev) {
          const m = ev.data;
          if (!m || m.type !== 'wrote' || m.id !== id) return;
          clearTimeout(timer);
          window.removeEventListener('message', on);
          if (!m.ok) { reject(new Error(m.error || 'The file could not be saved.')); return; }
          if (purpose === 'host') { resolve('exported'); return; }
          if (purpose !== 'backup' && purpose !== 'create' && purpose !== 'reopen' && !m.inplace && window.L && L.ui && L.ui.toast) L.ui.toast(VO.t('Saved {0}', m.name || name), 6000);
          resolve(purpose === 'backup' ? m : (m.inplace ? 'saved' : 'exported'));
        }
        window.addEventListener('message', on);
        for (let i = 0; i < n; i++) {
          api.postMessage({ type: 'chunk', id, i, n, name: String(name || ''), purpose, uri, folder, draftKey, path: hostPath, ability, open: openAfter, b64: b64(bytes.subarray(i * CHUNK, Math.min(bytes.length, (i + 1) * CHUNK))) });
        }
      }));
    },
    /** File ▸ Open in VS Code: the editor's own dialog, then a new tab for each file */
    openWorkspace(accept) {
      return vsRequest({ type: 'pick-open', accept: accept || '' });
    },
    /** open a workspace file this editor remembers, in its own tab */
    openUri(uri) {
      return vsRequest({ type: 'open-uri', uri: String(uri || '') });
    },
    /** the folder of the file this editor tab has open: { uri, label }, or null before that file arrives */
    documentFolder() {
      return openFolder ? { uri: openFolder.uri, label: openFolder.label } : null;
    },
    /** ask the extension again for that folder. The label is the path the editor shows. */
    async askFolder() {
      if (!inVs) return null;
      const r = await vsRequest({ type: 'document-folder' });
      if (!r || !r.ok) return openFolder ? { uri: openFolder.uri, label: openFolder.label } : null;
      openFolder = { uri: r.uri || '', label: r.label || '' };
      return { uri: openFolder.uri, label: openFolder.label };
    },
    /** File ▸ Save As ▸ Browse: VS Code's folder dialog. { uri, label }, or null if the user cancels. */
    async pickFolder(uri) {
      if (!inVs) return null;
      const r = await vsRequest({ type: 'pick-folder', uri: String(uri || '') });
      if (!r || !r.ok || r.cancelled) return null;
      return { uri: r.uri || '', label: r.label || '' };
    },
    /** File ▸ New in VS Code: ask for a path, then let makeBlob build the document, then open it in a new tab.
     *  The document in this tab stays as it is. ext, when set, is the suffix the saved name must end with. */
    async createDocument(makeBlob, name, ext) {
      const picked = await vsRequest({ type: 'pick-new', name: name || '', ext: ext || '' });
      if (!picked || !picked.ok || picked.alreadyOpen) return picked && picked.alreadyOpen ? 'open' : 'cancel';
      const ui = window.L && L.ui;
      if (ui && ui.busy) ui.busy(true);
      try {
        let blob;
        try { blob = await makeBlob(); }
        catch (e) {
          console.error(e);
          if (ui && ui.msg) ui.msg(String(e && e.message || e), { icon: 'error' });
          return 'cancel';
        }
        if (!blob) return 'cancel';
        try { await VO.host.writeFile(picked.name, blob, { purpose: 'create', uri: picked.uri }); }
        catch (e) { console.error(e); return 'cancel'; }   // the extension already reported this
        return 'created';
      } finally {
        if (ui && ui.busy) ui.busy(false);
      }
    },
  };

  /* ------------------------------------------------------------ updates */
  /** updater/current.json on the published site says, per policy and platform, the version an edition must
   *  be at least: older than "show-update-dialog" asks with a dialog, older than "show-update-button" shows an
   *  Update button. Versions compare by date (VO.compareVersions); a dev build never asks. The desktop app
   *  checks when the Start Center opens; nothing about the user or their files is sent. */
  VO.UPDATES = (native && native.updates) || 'https://vibeoffice.work/updater/current.json';
  /** updater/whatsnew.json: { "<version>": "<plain text>" }, one list for every platform */
  VO.WHATSNEW = (native && native.whatsnew) || 'https://vibeoffice.work/updater/whatsnew.json';
  /** the page the Update button opens (in the system's browser): it gets this edition's version and platform */
  VO.updateUrl = () => 'https://vibeoffice.work/updater/?from_version=' + encodeURIComponent(VO.VERSION)
    + '&current_platform=' + encodeURIComponent(VO.host.platform);
  /** { ask: 'dialog', 'button' or null (up to date), news: [{ version, text }] }: when an update is due, the
   *  whatsnew.json entries newer than this version, newest first (none if that file cannot be read).
   *  Throws when current.json cannot be read (offline). */
  VO.checkUpdate = async () => {
    const res = await fetch(VO.UPDATES, { cache: 'no-store' });
    if (!res.ok) throw new Error('updates: HTTP ' + res.status);
    const policy = (await res.json()) || {};
    const older = (k) => policy[k] && typeof policy[k][VO.host.platform] === 'string'
      && VO.compareVersions(VO.VERSION, policy[k][VO.host.platform]) < 0;
    const ask = older('show-update-dialog') ? 'dialog' : older('show-update-button') ? 'button' : null;
    let all = {};
    if (ask) try { const r = await fetch(VO.WHATSNEW, { cache: 'no-store' }); if (r.ok) all = (await r.json()) || {}; } catch (e) { /* the dialog without notes */ }
    const news = Object.entries(all)
      .filter(([v, text]) => typeof text === 'string' && versionDay(v) !== null && VO.compareVersions(VO.VERSION, v) < 0)
      .sort(([a], [b]) => VO.compareVersions(b, a))
      .map(([version, text]) => ({ version, text }));
    return { ask, news };
  };

  /* ------------------------------------------------------------ open source licenses */
  /** About ▸ Open source licenses: the credits page (made by tools/oss-licenses.py) in a window or tab of its own */
  VO.LICENSES = BASE + 'oss-licenses.html';
  VO.showLicenses = () => VO.host.newWindow(VO.LICENSES).catch((e) => console.error('licenses:', e));

  /* ------------------------------------------------------------ fonts on this computer */
  /** the font families Local Font Access listed (Chrome and Edge, after the user allows it), kept in this
   *  browser for every app of the suite; the Font boxes show them (L.localFonts in core.js) */
  const LOCAL_FONTS = 'vibeoffice.localFonts';
  let localFonts = null;
  window.addEventListener('storage', (e) => { if (e.key === LOCAL_FONTS || e.key === null) localFonts = null; });
  VO.localFonts = {
    supported: () => typeof window.queryLocalFonts === 'function',
    list() {
      if (!localFonts) { try { localFonts = JSON.parse(localStorage.getItem(LOCAL_FONTS)) || []; } catch (e) { localFonts = []; } }
      return localFonts;
    },
    /** ask the browser for the list (it asks the user the first time) and keep it; the number of families */
    async load() {
      const families = Array.from(new Set((await window.queryLocalFonts()).map((f) => f.family))).sort((x, y) => x.localeCompare(y));
      localFonts = families;
      try { localStorage.setItem(LOCAL_FONTS, JSON.stringify(families)); } catch (e) { /* storage off: kept for this page only */ }
      return families.length;
    },
    forget() { localFonts = []; try { localStorage.removeItem(LOCAL_FONTS); } catch (e) { /* storage off */ } },
  };

  /* Saving, for now: in the desktop app a save is a download into the Downloads folder; the native side says
     which file it wrote (the name only), and the app shows it with Browse (docs/desktop.md, Saving) */
  if (native) window.addEventListener('vo-download', (e) => {
    const ui = window.L && window.L.ui, d = e.detail || {};
    if (!ui) return;
    if (d.ok) ui.toast(VO.t('Saved to Downloads folder: {0}', d.name), 8000, { label: 'Browse', run: () => VO.host.showDownloads().catch((err) => console.error('browse:', err)) });
    else ui.toast('The file could not be saved to the Downloads folder.', 8000);
  });

  /* ------------------------------------------------------------ service worker */
  /* The published site, https://vibeoffice.work/, is cached for offline use, one version per page (sw.js); the
     worker is registered there and on 127.0.0.1 (any port: a published copy tried locally; the dev server's own
     files are unstamped, so the worker caches nothing there), and anywhere else (preview deployments) any earlier
     registration is removed. While the worker caches a version it reports to every page: 'vo-cache' on window, detail { state:
     'progress', 'done' or 'failed', update (an earlier version was cached: this is an update), progress (0 to 1),
     version }. updateViaCache 'none': the browser always asks the server for sw.js, so a new version is found on
     the next visit. */
  if (VO.host.kind === 'web' && 'serviceWorker' in navigator) {
    if (location.origin === 'https://vibeoffice.work' || location.hostname === '127.0.0.1') {
      navigator.serviceWorker.register(BASE + 'sw.js', { updateViaCache: 'none' }).catch((e) => console.warn('service worker:', e.message));
      navigator.serviceWorker.addEventListener('message', (e) => {
        if (e.data && e.data.type === 'vo-cache') window.dispatchEvent(new CustomEvent('vo-cache', { detail: e.data }));
      });
      navigator.serviceWorker.startMessages();
    } else {
      navigator.serviceWorker.getRegistrations().then((all) => all.forEach((r) => r.unregister())).catch(() => {});
    }
  }

  /* ------------------------------------------------------------ IndexedDB */
  let dbp = null;
  function db() {
    return dbp || (dbp = new Promise((resolve, reject) => {
      const r = indexedDB.open('vibeoffice', 1);
      r.onupgradeneeded = () => {
        r.result.createObjectStore('recent', { keyPath: 'key' });
        r.result.createObjectStore('handoff', { keyPath: 'id' });
      };
      r.onsuccess = () => resolve(r.result);
      r.onerror = () => reject(r.error);
    }));
  }
  /** run fn(store) in one transaction; resolves with the result of the request fn returns, once committed */
  async function tx(name, mode, fn) {
    const t = (await db()).transaction(name, mode);
    return new Promise((resolve, reject) => {
      let out;
      const req = fn(t.objectStore(name));
      if (req) req.onsuccess = () => { out = req.result; };
      t.oncomplete = () => resolve(out);
      t.onerror = t.onabort = () => reject(t.error);
    });
  }
  const chan = 'BroadcastChannel' in window ? new BroadcastChannel('vibeoffice') : null;
  const listeners = new Set();
  const notify = () => { VO.recent.refresh(); for (const fn of listeners) if (!fn.cacheOnly) fn(); };   // cache users redraw after the refresh
  const changed = () => { if (chan) chan.postMessage('recent'); notify(); };
  if (chan) chan.addEventListener('message', (e) => { if (e.data === 'recent') notify(); else if (e.data === 'look') applyLook(storedLook()); });
  /** call fn when the Recent Files list changes (here or in another page); returns the function that stops it */
  VO.onRecentChange = (fn) => { listeners.add(fn); return () => listeners.delete(fn); };

  /* ------------------------------------------------------------ Looks */
  /* the chrome's colours (common/looks.css) and, for a look with layout 'ribbon', the ribbon instead of menus and
     toolbars (common/ribbon.js); one choice for the whole suite: <html data-look data-layout> is set before
     the first paint by each page's head script from localStorage 'vibeoffice.look'; changing it here
     repaints every open window of the suite. Canvas-drawn chrome (rulers, sheet headers) redraws on 'vo-look'. */
  const N = (s) => s;   // UI text, translated where it is shown (VO.N, common/i18n.js)
  N('Blank Document'); N('Blank Workbook'); N('Blank Presentation');   // templates.json's Blank entries (tools/templates.mjs)
  const LOOKS = [{ id: 'win7', name: N('Windows 7'), note: 'Glass buttons and dialogs. The menus stay the blue of Office 2003.' },
    { id: 'xp', name: N('Windows XP'), note: 'Luna windows and menus, on the classic blue toolbars.' },
    { id: 'classic', name: N('Classic'), note: 'The classic blue toolbars and menus.' },
    { id: 'paper', name: N('Paper'), note: 'Light and calm: warm paper tones, one teal accent, rounded shapes.' },
    { id: 'paper2016', name: N('Paper 2016'), layout: 'ribbon', note: 'Paper with tabs of commands, grouped by task, instead of menus and toolbars.' },
    { id: 'ribbon2010', name: N('Ribbon 2010'), layout: 'ribbon', note: 'The Office 2010 ribbon. Windows, buttons, scroll bars and the type stay as in Windows 7.' },
    { id: 'goldengate', name: N('Ribbon Golden Gate'), layout: 'ribbon', note: 'A macOS ribbon: a white toolbar, liquid-glass tabs and menus, and traffic-light windows.' },
    { id: 'goldengatedark', name: N('Ribbon Golden Gate (Dark)'), layout: 'ribbon', note: 'Ribbon Golden Gate in Puppertino dark colours. Pages, the sheet and slides stay light.' },
    { id: 'vscodegg', name: N('VS Code Golden Gate'), layout: 'ribbon', note: 'Golden Gate in the colours of the current VS Code theme. Glass follows light or dark.' }];
  const storedLook = () => {
    try {
      const id = localStorage.getItem('vibeoffice.look') || 'ribbon2010';
      /* this look exists only in a VS Code or code-server webview; a browser keeps Ribbon 2010 */
      if (id === 'vscodegg' && !(window.__VO_HOST__ && window.__VO_HOST__.kind === 'vscode')) return 'ribbon2010';
      return id;
    } catch (e) { return 'ribbon2010'; }
  };
  const applyLook = (id) => {
    const root = document.documentElement;
    if (!LOOKS.some((l) => l.id === id)) return;
    const changed = root.dataset.look !== id;
    root.dataset.look = id;
    root.dataset.layout = LOOKS.find((l) => l.id === id).layout || 'bars';
    const meta = document.querySelector('meta[name="theme-color"]');
    const c = getComputedStyle(root).getPropertyValue('--theme-color').trim();
    if (meta && c) meta.content = c;
    if (changed) window.dispatchEvent(new Event('vo-look'));
  };
  VO.look = {
    /** [{ id, name, note, layout }] */
    list: () => LOOKS.filter((l) => l.id !== 'vscodegg' || VO.host.kind === 'vscode').map((l) => ({ ...l })),
    get: () => document.documentElement.dataset.look || 'ribbon2010',
    /** 'bars' (menus and toolbars) or 'ribbon' */
    layout: () => document.documentElement.dataset.layout || 'bars',
    set(id) {
      if (!LOOKS.some((l) => l.id === id)) return;
      if (id === 'vscodegg' && VO.host.kind !== 'vscode') return;
      try { localStorage.setItem('vibeoffice.look', id); } catch (e) { /* this window still changes */ }
      applyLook(id);
      if (chan) chan.postMessage('look');
    },
  };
  applyLook(storedLook());   // a page without the head script, or a stored look that changed meanwhile

  /* ------------------------------------------------------------ Recent Files */
  const MAX_RECENT = 30;
  const MAX_BYTES = 50 * 1024 * 1024;   // bigger files are not kept (open them from disk again)
  const prune = async () => {
    /* entries with unsaved changes are never dropped */
    const all = (await VO.recent.list()).filter((e) => !e.draft);
    if (all.length > MAX_RECENT) await tx('recent', 'readwrite', (s) => { for (const e of all.slice(MAX_RECENT)) s.delete(e.key); });
  };
  /* Safari clears a site's data after about 7 days without a visit, unsaved versions included; once a
     document has unsaved changes, ask the browser to keep the data (Safari and Chrome decide without
     asking the user; Firefox may ask) */
  let persistAsked = false;
  const keepData = () => {
    if (persistAsked || !navigator.storage || !navigator.storage.persist) return;
    persistAsked = true;
    navigator.storage.persisted().then((yes) => yes || navigator.storage.persist()).catch(() => {});
  };
  VO.recent = {
    /** newest first: { key, app, name, size, time, file: Blob|null (opened or saved), thumb: Blob|null,
     *  draft: { file: Blob, name, time } | null (unsaved changes) }; a never-saved document has only a draft */
    async list() {
      const all = (await tx('recent', 'readonly', (s) => s.getAll())) || [];
      return all.sort((a, b) => b.time - a.time);
    },
    get: (key) => tx('recent', 'readonly', (s) => s.get(key)),
    /** the list as last read (for menus, which are built synchronously) */
    cache: [],
    refresh: (() => { let t = 0; return () => { clearTimeout(t); t = setTimeout(() => VO.recent.list().then((l) => { VO.recent.cache = l; for (const fn of listeners) if (fn.cacheOnly) fn(); }, () => {}), 30); }; })(),
    /** this app's entries, newest first, from the cache */
    forApp: (app) => VO.recent.cache.filter((e) => e.app === (app || (current && current.app))),
    /** add or refresh the entry for this file name (a newer copy replaces the older one, its thumbnail stays);
     *  keepDraft: opening a file keeps unsaved changes made to it, saving it clears them.
     *  extra.uri, when passed, is the workspace file this editor tab is bound to (vscode); pass '' to clear it */
    async put(app, name, blob, keepDraft, extra) {
      if (!app || !blob || blob.size > MAX_BYTES) return null;
      const key = app + ':' + name;
      const old = await VO.recent.get(key);
      const uri = extra && 'uri' in extra ? (extra.uri || '') : ((old && old.uri) || '');
      const rec = { key, app, name, size: blob.size, time: Date.now(), file: blob, thumb: old ? old.thumb : null, draft: keepDraft && old ? old.draft || null : null };
      if (uri) rec.uri = uri;
      await tx('recent', 'readwrite', (s) => s.put(rec));
      await prune();
      changed();
      return key;
    },
    /** store the unsaved version of a document */
    async setDraft(key, app, name, draftName, blob, lossState) {
      const e = (await VO.recent.get(key)) || { key, app, name, size: 0, file: null, thumb: null };
      if (!e.file) e.name = name;
      e.draft = { file: blob, name: draftName || name, time: Date.now(), lossState };
      e.time = e.draft.time;
      await tx('recent', 'readwrite', (s) => s.put(e));
      changed();
      keepData();
    },
    /** throw the unsaved version away; a document that was never saved leaves the list */
    async discardDraft(key) {
      const e = await VO.recent.get(key);
      if (!e) return;
      if (!e.file) await tx('recent', 'readwrite', (s) => s.delete(key));
      else { e.draft = null; await tx('recent', 'readwrite', (s) => s.put(e)); }
      changed();
    },
    async setThumb(key, thumb) {
      const e = await VO.recent.get(key);
      if (!e) return;
      e.thumb = thumb;
      await tx('recent', 'readwrite', (s) => s.put(e));
      changed();
    },
    async remove(key) { await tx('recent', 'readwrite', (s) => s.delete(key)); changed(); },
    async clear() { await tx('recent', 'readwrite', (s) => s.clear()); changed(); },
  };

  /* ------------------------------------------------------------ handing files to apps */
  /** where an app page opens: a new window in the desktop app or when installed (the start page stays, as in
   *  LibreOffice), else this tab */
  VO.go = async (url) => {
    if (VO.host.kind === 'vscode') return;   // this editor stays on the workspace file it was opened with
    const standalone = matchMedia('(display-mode: standalone)').matches || matchMedia('(display-mode: window-controls-overlay)').matches;
    try { if ((VO.host.kind === 'desktop' || standalone) && await VO.host.newWindow(url)) return; } catch (e) { console.error('new window:', e); }
    location.href = url;
  };
  /** File ▸ Exit, after the app has closed its documents: an app window the Start Center opened
   * (installed) closes and shows the Start Center behind it; a tab goes to the Start Center.
   * In VS Code the editor tab closes. */
  VO.exit = async () => {
    if (VO.host.kind === 'vscode') { vsPost({ type: 'close' }); return; }
    if (await VO.host.closeWindow().catch(() => false)) return;   // the desktop app's app windows
    if (window.opener && !window.opener.closed) { window.close(); if (window.closed) return; }
    location.href = BASE;
  };
  /** open a file in the app made for it; false if no app opens this kind of file */
  VO.handoff = async (file) => {
    const app = VO.appFor(file.name);
    if (!app) return false;
    if (VO.host.kind === 'vscode') {
      const ui = window.L && window.L.ui;
      const text = VO.t('Open “{0}” from the Explorer.', file.name);
      if (ui && ui.msg) ui.msg(text, { icon: 'info' });
      else alert(text);
      return true;
    }
    const id = Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
    const day = Date.now() - 86400000;
    await tx('handoff', 'readwrite', (s) => {
      s.put({ id, name: file.name, file, time: Date.now() });
      s.openCursor().onsuccess = (e) => { const c = e.target.result; if (!c) return; if (c.value.time < day) c.delete(); c.continue(); };
    });
    VO.go(VO.url(app, { open: id }));
    return true;
  };

  /* ------------------------------------------------------------ the app side */
  const q = new URLSearchParams(location.search);
  /** how this app page was opened: null (plain: a blank document), { new }, { template: id }, { open: id }, { recent: key }
   *  or { vscode } (the custom editor: a blank document, then the workspace file) */
  VO.launch = inVs ? { vscode: true }
    : q.has('open') ? { open: q.get('open') } : q.has('recent') ? { recent: q.get('recent'), draft: q.has('draft') }
    : q.has('template') ? { template: q.get('template') } : q.has('new') ? { new: true } : null;

  let current = null;   // { app, hooks }
  /* VS Code has two ways to press Ctrl+S: the page's command, and the workbench. Both must be one
     write. The page asks the extension to save (request-save). saveCustomDocument answers do-save,
     which is the only call that runs the app's save or sends a blob the app already built. */
  let inHostSave = false;
  let hostSavePromise = null;
  let hostSaveRun = null;
  let pendingWrite = null;
  let heldDirty = false;
  let origExec = null;
  const fileTitle = (name) => {
    const s = String(name || '').replace(/\\/g, '/');
    const i = s.lastIndexOf('/');
    return i >= 0 ? s.slice(i + 1) : s;
  };
  const isThisFile = (name, opts) => {
    if (!hostLive || !documentName) return false;
    const folder = opts && opts.folder;
    if (folder && (!openFolder || folder !== openFolder.uri)) return false;
    const a = fileTitle(name), b = fileTitle(documentName);
    return !!a && a.toLowerCase() === b.toLowerCase();
  };
  const requestHostSave = () => {
    if (!hostSavePromise) {
      hostSavePromise = vsRequest({ type: 'request-save' }).then((r) => {
        if (!r || !r.ok) return false;
        return 'saved';
      }).finally(() => { hostSavePromise = null; });
    }
    return hostSavePromise;
  };
  const stillDirty = () => {
    const H = current && current.hooks;
    if (!H || !H.isDirty) return true;
    let d = null;
    try { d = H.current && H.current(); } catch (e) { return true; }
    if (!d) return false;
    try { return !!H.isDirty(d); } catch (e) { return true; }
  };
  const installSaveHook = () => {
    if (!inVs) return;
    const ui = window.L && L.ui;
    if (!ui || ui.exec._voHostSave) return;
    origExec = ui.exec;
    function wrapped(id, arg, ev) {
      if (id === 'save') return requestHostSave();
      return origExec.call(ui, id, arg, ev);
    }
    wrapped._voHostSave = true;
    ui.exec = wrapped;
  };
  if (inVs) VO.host.saveThroughEditor = (name, blob, opts) => {
    if (inHostSave || !isThisFile(name, opts)) return undefined;
    const ticket = { name, blob, opts: opts || null };
    pendingWrite = ticket;
    return requestHostSave().then((r) => {
      if (pendingWrite !== ticket) return r;
      pendingWrite = null;
      if (r !== 'saved') return r;
      return VO.host.saveThroughEditor(name, blob, opts);
    });
  };
  async function performHostSave() {
    inHostSave = true;
    heldDirty = false;
    let stagedSave = false;
    try {
      installSaveHook();
      let r;
      const staged = pendingWrite;
      pendingWrite = null;
      if (staged) {
        /* the app already built this blob and is waiting above saveFile; do not build it again */
        stagedSave = true;
        r = await VO.host.writeFile(staged.name, staged.blob, staged.opts || undefined);
      } else if (origExec) r = await origExec.call(window.L.ui, 'save');
      else r = await window.L.ui.exec('save');
      const ok = r === 'saved' || r === true;
      return { ok, error: ok ? '' : (r === 'exported' ? 'Saved a copy next to this file. This file still has unsaved changes.' : 'Save was cancelled.') };
    } catch (e) {
      return { ok: false, error: String(e && e.message || e) };
    } finally {
      inHostSave = false;
      /* staged: the app clears its dirty flag after this returns. exec: it already did,
         so a change that is still dirty happened during the save and the tab must keep it */
      if (heldDirty && !stagedSave && stillDirty()) vsPost({ type: 'dirty', dirty: true });
      heldDirty = false;
    }
  }
  /* Keep in step with plainValue in vscode/documents.js. A VB object must not cross postMessage. */
  function plainVbValue(value, depth) {
    if (value == null) return null;
    if (typeof value === 'string' || typeof value === 'boolean') return value;
    if (typeof value === 'number') return Number.isFinite(value) ? value : null;
    const level = depth || 0;
    if (!Array.isArray(value) || level > 8) return undefined;
    const out = [];
    for (let i = 0; i < value.length; i++) {
      const item = plainVbValue(value[i], level + 1);
      if (item === undefined) return undefined;
      out.push(item);
    }
    return out;
  }
  function vbFailure(e) {
    const error = { message: e && e.message ? String(e.message) : String(e) };
    if (e && typeof e.number === 'number') error.number = e.number;
    if (e && e.code != null && e.code !== '') error.code = String(e.code);
    return error;
  }
  /** Run one agent script on the document this tab has open. Prints are the new Debug.Print lines. */
  async function runHostVbs(m) {
    const app = current && current.app || '';
    const fileName = documentName || '';
    const base = { app: app, fileName: fileName, prints: [], value: null };
    if (!hostLive || !window.L) {
      return Object.assign(base, { ok: false, error: { message: 'The file is not open in the editor yet.' } });
    }
    const api = window.L;
    if (!api.vb || typeof api.vb.session !== 'function' || typeof api.runVB6 !== 'function' || typeof api.evalVB6 !== 'function') {
      return Object.assign(base, { ok: false, error: { message: 'Visual Basic is not available in this editor.' } });
    }
    let host;
    try { host = await api.vb.session(); }
    catch (e) { return Object.assign(base, { ok: false, error: vbFailure(e) }); }
    if (window.L && L.hostAudit && L.hostAudit.beginScript) L.hostAudit.beginScript();
    const wasAgent = VO.host.vbsAgent;
    VO.host.vbsAgent = !!m.agent;
    const mark = host && host.output ? host.output.length : 0;
    const printed = () => (host && host.output ? host.output.slice(mark).map((line) => String(line)) : []);
    try {
      let value;
      if (m.mode === 'line') value = await api.evalVB6(String(m.source || ''));
      else value = await api.runVB6(String(m.source || ''), m.entry ? String(m.entry) : undefined, Array.isArray(m.args) ? m.args : undefined);
      const plain = plainVbValue(value);
      const dropped = value != null && plain === undefined;
      const out = { ok: true, app: app, fileName: fileName, prints: printed(), value: plain === undefined ? null : plain, error: null };
      if (dropped) out.note = 'The result was a VB object. Use Debug.Print to return a string, number, or Boolean.';
      return out;
    } catch (e) {
      return { ok: false, app: app, fileName: fileName, prints: printed(), value: null, error: vbFailure(e) };
    } finally {
      VO.host.vbsAgent = wasAgent;
    }
  }
  if (inVs) window.addEventListener('message', async (ev) => {
    const m = ev.data;
    if (!m || typeof m !== 'object') return;
    if (m.type === 'do-save') {
      if (!hostLive || !window.L || !L.ui) {
        vsPost({ type: 'save-result', id: m.id, ok: false, error: 'The file is not open in the editor yet.' });
        return;
      }
      if (!hostSaveRun) hostSaveRun = performHostSave().finally(() => { hostSaveRun = null; });
      const result = await hostSaveRun;
      vsPost({ type: 'save-result', id: m.id, ok: result.ok, error: result.error || '' });
    } else if (m.type === 'run-vbs') {
      let result;
      try { result = await runHostVbs(m); }
      catch (e) {
        result = { ok: false, app: current && current.app || '', fileName: documentName || '', prints: [], value: null, error: { message: String(e && e.message || e) } };
      }
      vsPost(Object.assign({ type: 'vbs-result', id: m.id }, result));
    } else if (m.type === 'backup') {
      try {
        const H = current && current.hooks;
        const doc = H && H.current && H.current();
        if (!H || !H.snapshot || !doc) throw new Error('not ready');
        await VO.host.writeFile('backup', await H.snapshot(doc), { purpose: 'backup', id: m.id });
      } catch (e) {
        vsPost({ type: 'wrote', id: m.id, ok: false, error: String(e && e.message || e) });
      }
    }
  });
  /* VS Code: this dialog is up from the moment dialogs exist until the workspace file is open, so the blank
     document the page builds first is covered. Cancel closes the editor tab. The title has no close button,
     and Escape does not dismiss it (ui.dialog, canClose: false). */
  let openingDlg = null;
  let openingCancelled = false;
  const endOpening = () => {
    const d = openingDlg;
    openingDlg = null;
    if (d) d.close(null);
  };
  /** this editor was opened on a file that will not be shown: the password was cancelled, it was wrong,
   *  or the encryption cannot be read. Leave the opening dialog up and close the tab, so the blank
   *  document underneath never appears. No effect outside VS Code. */
  VO.abandonEditor = () => {
    if (!inVs) return;
    openingCancelled = true;
    vsPost({ type: 'close' });
  };
  if (inVs) VO.openingDialog = () => {
    if (openingDlg || openingCancelled || !window.L || !L.ui || !L.ui.dialog || !L.h) return;
    const body = L.h('div', { class: 'brand-splash' },
      L.h('p', { class: 'brand-name', translate: 'no', text: L.APP || '' }),
      L.h('p', { class: 'brand-wait', text: 'Opening the file. Please wait...' }));
    openingDlg = L.ui.dialog({
      title: VO.t('Starting {0}', L.APP || ''),
      width: 500,
      brand: true,
      canClose: false,
      body,
      focus: body,
      buttons: [{ label: 'Cancel', cancel: true, onClick() { openingCancelled = true; vsPost({ type: 'close' }); return false; } }],
    });
  };
  /** an app is ready: open the file it was launched with, if any, then show the Document Recovery pane when
   *  other documents of this app have unsaved changes.
   *  hooks: open(file, { draft, name, saved }), adoptBlank(name) keeps the page's own blank document as that
   *  new file (VS Code's New File…), thumb() -> a canvas (or image) of the first page, message(text),
   *  templates() -> [{ id, name }] (read by tools/templates.mjs); for autosave: snapshot(doc) -> Blob in the
   *  app's format, docInfo(doc) -> { name, draftName }, isDirty(doc), current() -> doc, autosave() -> false
   *  when turned off; showRecovery() opens the Document Recovery pane */
  VO.attach = async (app, hooks) => {
    installSaveHook();
    current = { app, hooks: hooks || {} };
    VO.recent.refresh();
    const l = VO.launch;
    if (l && !l.new) {
      if (!inVs && history.replaceState) history.replaceState(null, '', location.pathname);   // a reload starts clean
      if (l.vscode) {
        try {
          const file = await VO.host.takeDocument();
          if (openingCancelled) return;
          if (file && file.voBlank) {
            /* the page already built a blank document; name it as this file and write that document in.
               A failed write is reported by the extension; the document stays on screen so Save can try again. */
            try {
              if (current.hooks.adoptBlank) current.hooks.adoptBlank(file.name);
              hostLive = true;
              const doc = current.hooks.current && current.hooks.current();
              const blob = doc && current.hooks.snapshot ? await current.hooks.snapshot(doc) : null;
              if (openingCancelled) return;
              if (!blob) throw new Error('The editor did not build a document.');
              await VO.host.writeFile(file.name, blob);
              if (openingCancelled) return;
              VO.saved(file.name, blob, doc);
            } catch (e) { console.error(e); }
          } else {
            const draft = !!(file && file.voDraft);
            let lossState;
            if (draft) {
              try { const rec = await VO.recent.get(file.voDraft); lossState = rec && rec.draft && rec.draft.lossState; }
              catch (err) { /* the unsaved copy's notes stay with the tab that wrote them */ }
            }
            /* voLaunch: this open is the file the tab was opened on. A refused password closes the tab. */
            if (file) file.voLaunch = true;
            if (file && current.hooks.open) await current.hooks.open(file, draft ? { draft: true, name: file.name, saved: true, lossState } : {});
            if (openingCancelled) return;
            if (!(file && current.hooks.open)) (current.hooks.message || alert)('That file is no longer available here. Open it again from the Explorer.');
            if (file) hostLive = true;
            /* unsaved bytes opened into this tab still belong to the workspace file, so the tab shows them as changes */
            if (draft) vsPost({ type: 'dirty', dirty: true });
          }
        } catch (e) {
          if (openingCancelled) return;
          console.error(e);
          (current.hooks.message || alert)('That file could not be opened. ' + (e.message || e));
        } finally {
          if (!openingCancelled) endOpening();
        }
      } else if (l.open) {
        let rec = null;
        try {
          rec = await tx('handoff', 'readonly', (s) => s.get(l.open));
          if (rec) await tx('handoff', 'readwrite', (s) => s.delete(l.open));
        } catch (e) { console.error(e); }
        if (rec && rec.file) await current.hooks.open(asFile(rec.file, rec.name), {});
        else (current.hooks.message || alert)(VO.t('That file is no longer available here. Open it again from your computer.'));
      } else if (l.recent) await VO.openRecent(l.recent, { draft: l.draft });
    }
    setTimeout(async () => {
      const others = (await VO.recent.list()).filter((e) => e.app === app && e.draft && !openKeys.has(e.key));
      if (others.length && current.hooks.showRecovery) current.hooks.showRecovery();
    }, 700);
  };
  const asFile = (blob, name) => (blob instanceof File && blob.name === name ? blob : new File([blob], name, { type: blob.type }));
  /** open a Recent Files entry in this app: its unsaved version when draft is true, else the saved file.
   *  In VS Code the entry opens in its own tab. A remembered workspace path is that file; without one, the
   *  copy kept here is written to a path the user chooses. */
  VO.openRecent = async (key, o) => {
    let e = null;
    try { e = await VO.recent.get(key); } catch (err) { console.error(err); }
    const draft = !!(o && o.draft && e && e.draft);
    const blob = e && (draft ? e.draft.file : e.file);
    if (!blob) { (current.hooks.message || alert)(VO.t('That file is no longer available here. Open it again from your computer.')); return; }
    if (inVs && VO.host.createDocument) {
      if (e.uri) {
        try {
          const res = draft
            ? await VO.host.writeFile(e.draft.name || e.name, blob, { purpose: 'reopen', uri: e.uri, draftKey: key })
            : await VO.host.openUri(e.uri);
          if (!draft && res && res.ok !== false) return;
          if (draft) return;
        } catch (err) { console.warn('recent file:', err); }
      }
      const name = draft ? (e.draft.name || e.name) : e.name;
      const ext = (String(name).match(/\.([^.]+)$/) || [])[1] || '';
      await VO.host.createDocument(async () => blob, name, ext.toLowerCase());
      return;
    }
    const f = asFile(blob, draft ? e.draft.name : e.name);
    if (draft) f.voDraft = key;
    await current.hooks.open(f, { draft, name: e.name, saved: !!e.file, lossState: draft ? e.draft.lossState : undefined });
  };
  /** File ▸ New… inside an app: Blank and the app's templates with their previews (the Start Center's
   *  Create File gallery, common/templates.json), in the app's own dialog. create(id): id '' is Blank. */
  VO.newDialog = async (app, create) => {
    const ui = window.L && window.L.ui;
    if (!ui || !ui.dialog) return create('');
    let list = null;
    try { list = (await (await fetch(BASE + 'common/templates.json')).json())[app]; } catch (e) { console.warn('templates:', e); }
    if (!list || !list.length) list = [{ id: '', name: 'Blank', img: null }];
    if (!document.getElementById('vo-gallery-css')) {
      const st = document.createElement('style');
      st.id = 'vo-gallery-css';
      st.textContent = '.vo-gal{display:grid;grid-template-columns:repeat(auto-fill,minmax(132px,1fr));gap:8px;max-height:min(62vh,540px);overflow:auto;padding:6px;background:var(--surface);border:1px solid var(--field-border)}'
        + '.vo-tpl{display:flex;flex-direction:column;gap:4px;padding:5px;border:1px solid transparent;border-radius:2px;background:none;text-align:left;cursor:default;font:inherit;color:var(--ink)}'
        + '.vo-tpl:hover,.vo-tpl:focus-visible{background:var(--menu-hot);border-color:var(--card-hot-border);outline:none}'
        + '.vo-tpl span{display:flex;align-items:center;justify-content:center;height:112px;background:var(--card-thumb-bg);border:1px solid var(--card-thumb-line);overflow:hidden}'
        + '.vo-tpl img{max-width:100%;max-height:100%;box-shadow:0 1px 2px var(--tip-shadow-c)}'
        + '.vo-tpl b{font-weight:normal;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}'
        + '.vo-gal-note{margin:0 0 6px}';
      document.head.appendChild(st);
    }
    const gal = document.createElement('div');
    gal.className = 'vo-gal';
    let d = null;
    for (const t of list) {
      const b = document.createElement('button');
      b.type = 'button'; b.className = 'vo-tpl'; b.title = VO.t(t.name);
      const pic = document.createElement('span');
      if (t.img) { const img = document.createElement('img'); img.src = BASE + t.img; img.alt = ''; pic.append(img); }
      const name = document.createElement('b');
      name.textContent = VO.t(t.name);
      b.append(pic, name);
      b.addEventListener('click', () => { if (d) d.close(-1); create(t.id); });
      gal.append(b);
    }
    const body = document.createElement('div');
    const note = document.createElement('p');
    note.className = 'vo-gal-note';
    note.textContent = VO.t('Choose Blank or a template:');
    body.append(note, gal);
    d = ui.dialog({ title: VO.t({ quire: 'New Document', ledger: 'New Spreadsheet', lectern: 'New Presentation' }[app]), width: 700, body, buttons: [{ label: 'Cancel', cancel: true }], focus: gal.firstChild });
  };
  /** the open app's templates: [{ id, name }] */
  VO.templates = () => (current && current.hooks.templates ? current.hooks.templates() : []);
  /** a JPEG (at most 320 px) of the open document's first page, or null */
  VO.thumbnail = async () => {
    const make = current && current.hooks.thumb;
    const src = make && (await make());
    if (!src) return null;
    const k = Math.min(1, 320 / src.width, 320 / src.height);
    const c = document.createElement('canvas');
    c.width = Math.max(1, Math.round(src.width * k)); c.height = Math.max(1, Math.round(src.height * k));
    const g = c.getContext('2d');
    g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height);
    g.drawImage(src, 0, 0, c.width, c.height);
    return new Promise((r) => c.toBlob(r, 'image/jpeg', 0.88));
  };
  /** the thumbnail for a Recent Files entry, a little after the document has been drawn */
  function thumbLater(key) {
    if (!key) return;
    setTimeout(async () => {
      try { const blob = await VO.thumbnail(); if (blob) await VO.recent.setThumb(key, blob); }
      catch (e) { console.warn('thumbnail:', e); }
    }, 600);
  }
  /* ------------------------------------------------------------ documents, autosave */
  const docKeys = new WeakMap();   // the app's document object -> its Recent Files key
  const openKeys = new Set();      // keys of the documents open in this page (left out of Document Recovery)
  const timers = new WeakMap(), slow = new WeakMap();
  const DRAFT_DELAY = 15000;
  const bind = (doc, key) => { if (doc) docKeys.set(doc, key); openKeys.add(key); };
  /** the app opened a file (from any source) as doc */
  VO.opened = (file, doc) => {
    if (!current) return;
    if (file.voDraft) { bind(doc, file.voDraft); return; }   // an unsaved version: keep its entry as it is
    const extra = inVs ? { uri: (file && file.voUri) || '' } : undefined;
    VO.recent.put(current.app, file.name, file, true, extra).then((key) => { if (key) bind(doc, key); thumbLater(key); }, (e) => console.warn('recent files:', e));
  };
  const unsavedDocs = new Set();
  /** true when an open document still has unsaved changes (Help ▸ Language will not reload over them) */
  VO.unsaved = () => {
    const H = current && current.hooks;
    if (H && H.isDirty) {
      if (H.current) {
        let d = null;
        try { d = H.current(); } catch (e) { return true; }
        if (d) { if (H.isDirty(d)) unsavedDocs.add(d); else unsavedDocs.delete(d); }
      }
      for (const d of [...unsavedDocs]) {
        let dirty = true;
        try { dirty = !!H.isDirty(d); } catch (e) { dirty = true; }
        if (!dirty) unsavedDocs.delete(d);
      }
    }
    return unsavedDocs.size > 0;
  };
  /** the app saved doc in its own format as name: that is its entry now, without unsaved changes */
  VO.saved = (name, blob, doc) => {
    if (doc) unsavedDocs.delete(doc);
    /* The tab dot is cleared by the editor save that wrote the file, not by this notice.
       Posting it in the middle of that save used to start a second save and leave the dot on. */
    if (inVs && hostLive && !inHostSave) vsPost({ type: 'dirty', dirty: false });
    if (!current) return;
    const old = doc && docKeys.get(doc);
    if (doc) clearTimeout(timers.get(doc));
    VO.recent.put(current.app, name, blob, false, inVs && documentUri ? { uri: documentUri } : undefined).then(async (key) => {
      if (!key) return;
      if (old && old !== key) { await VO.recent.discardDraft(old); openKeys.delete(old); }
      bind(doc, key);
      thumbLater(key);
    }, (e) => console.warn('recent files:', e));
  };
  /** doc has new unsaved changes: its unsaved version is written a little later */
  VO.changed = (doc) => {
    const H = current && current.hooks;
    if (doc && H && H.isDirty && H.isDirty(doc)) unsavedDocs.add(doc);
    if (inVs && hostLive && H && H.isDirty && H.isDirty(doc)) {
      if (inHostSave) heldDirty = true;
      else vsPost({ type: 'dirty', dirty: true });
    }
    if (!current || !doc || !current.hooks.snapshot) return;
    clearTimeout(timers.get(doc));
    timers.set(doc, setTimeout(() => saveDraft(doc), Math.max(DRAFT_DELAY, slow.get(doc) || 0)));
  };
  async function saveDraft(doc) {
    const H = current.hooks;
    if ((H.isDirty && !H.isDirty(doc)) || (H.autosave && H.autosave() === false)) return;
    const t0 = performance.now();
    let blob = null;
    try { blob = await H.snapshot(doc); } catch (e) { console.warn('autosave:', e); return; }
    slow.set(doc, (performance.now() - t0) * 30);   // a document that takes 1 s to write is saved every 30 s at most
    if (!blob || blob.size > MAX_BYTES || (H.isDirty && !H.isDirty(doc))) return;
    const info = H.docInfo(doc);
    let key = docKeys.get(doc);
    if (!key) { key = current.app + ':~' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6); bind(doc, key); }
    try { await VO.recent.setDraft(key, current.app, info.name, info.draftName, blob, info.lossState); } catch (e) { console.warn('autosave:', e); return; }
    if (!H.current || H.current() === doc) thumbLater(key);
  }
  /** the user chose not to keep doc's changes (No in "Do you want to save the changes…?") */
  VO.discard = (doc) => {
    if (!doc) return;
    unsavedDocs.delete(doc);
    clearTimeout(timers.get(doc));
    const key = docKeys.get(doc);
    if (!key) return;
    docKeys.delete(doc);
    openKeys.delete(key);
    VO.recent.discardDraft(key).catch(() => {});
  };

  /* ------------------------------------------------------------ the list inside the apps */
  const when = (t) => {
    const d = new Date(t), today = new Date().toDateString() === d.toDateString();
    const loc = VO.i18n ? VO.i18n.dates : [];
    return today ? d.toLocaleTimeString(loc, { hour: '2-digit', minute: '2-digit' }) : d.toLocaleDateString(loc, { month: 'short', day: 'numeric' });
  };
  /* its text in the UI's language (common/i18n.js) */
  const el = (tag, cls, text) => { const e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = VO.t ? VO.t(text) : text; return e; };
  const icon = (name) => (window.L && L.icons ? L.icons.get(name) : '');
  /** File menu: the app's four most recent entries ({ label, run } items for L.ui menus) */
  VO.recentMenuItems = (max) => VO.recent.forApp().slice(0, max || 4).map((e, i) => ({
    label: '&' + (i + 1) + ' ' + (e.draft ? VO.t('{0} (unsaved changes)', e.name) : e.name).replace(/&/g, '&&'), raw: true,
    run: () => VO.openRecent(e.key, { draft: !!e.draft }),
  }));
  /** task pane: the app's recent entries as links (Getting Started ▸ Open); keeps itself up to date */
  VO.recentLinks = (max) => {
    const box = el('div', 'vo-recent');
    const draw = () => {
      if (!box.isConnected && box.dataset.drawn) { stop(); return; }
      box.dataset.drawn = '1';
      box.replaceChildren(...VO.recent.forApp().slice(0, max || 4).map((e) => {
        const b = el('button', 'tp-link');
        b.type = 'button';
        b.title = e.draft ? VO.t('{0} — unsaved changes from {1}', e.name, new Date(e.draft.time).toLocaleString(VO.i18n ? VO.i18n.dates : [])) : e.name;
        b.translate = false;
        const ic = el('span'); ic.innerHTML = icon(e.draft ? 'save' : 'open');
        const name = el('span');
        name.textContent = e.draft ? VO.t('{0} (unsaved)', e.name) : e.name;
        b.append(ic, name);
        b.addEventListener('click', () => VO.openRecent(e.key, { draft: !!e.draft }));
        return b;
      }));
    };
    draw.cacheOnly = true;
    const stop = VO.onRecentChange(draw);
    draw();
    return box;
  };
  /** the Document Recovery task pane: this app's unsaved versions, each to open or to discard */
  VO.recoveryPane = (b) => {
    const list = el('div', 'vo-recovery');
    const draw = () => {
      if (!list.isConnected && list.dataset.drawn) { stop(); return; }
      list.dataset.drawn = '1';
      const drafts = VO.recent.forApp().filter((e) => e.draft);
      list.replaceChildren(...(drafts.length ? drafts.map((e) => {
        const item = el('div', 'tp-sec');
        const head = el('div', 'tp-h', e.name);
        const from = e.file ? VO.t('Unsaved changes from {0}', when(e.draft.time)) : VO.t('Unsaved changes from {0} (never saved)', when(e.draft.time));
        const note = el('div', 'tp-note', openKeys.has(e.key) ? VO.t('{0} — open in this window', from) : from);
        const open = el('button', 'tp-link'); open.type = 'button';
        const oi = el('span'); oi.innerHTML = icon('open'); open.append(oi, el('span', null, 'Open the unsaved version'));
        open.addEventListener('click', () => VO.openRecent(e.key, { draft: true }));
        const drop = el('button', 'tp-link'); drop.type = 'button';
        const di = el('span'); di.innerHTML = icon('delete') || icon('cut'); drop.append(di, el('span', null, 'Discard the changes'));
        drop.addEventListener('click', () => VO.recent.discardDraft(e.key));
        item.append(head, note, open, drop);
        return item;
      }) : [el('div', 'tp-note', 'No unsaved versions are waiting.')]));
    };
    draw.cacheOnly = true;
    const stop = VO.onRecentChange(draw);
    b.append(el('div', 'tp-note', VO.t('{0} kept the unsaved changes of these documents. Open the ones you want to keep and save them; discard the rest.', current ? VO.APPS[current.app].name : 'VibeOffice')), list);
    VO.recent.list().then((l) => { VO.recent.cache = l; draw(); }, () => draw());
  };
})();
