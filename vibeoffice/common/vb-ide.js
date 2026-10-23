/* Visual Basic IDE. One standard module, debugged on the app host's debug VM.
   Opened from View. The window is the ProjectEditor layout: editor, output,
   expression, breakpoints, watches, call stack, object values. */
(function () {
  'use strict';
  const L = window.L;
  const MODULE = 'module1';

  let code = '';
  let breakpoints = [];
  let watches = [];
  let expression = '';
  let expressionResult = '';
  let projectName = 'Project.json';
  let editor = null;
  let openDlg = null;
  let output = null;
  let exprInput = null;
  let exprOut = null;
  let lineInput = null;
  let condInput = null;
  let watchInput = null;
  let bpBody = null;
  let watchBody = null;
  let stackList = null;
  let localList = null;
  let editorHost = null;
  let lowerPane = null;
  let editorShare = 1.15;
  let startBtn = null;
  let contBtn = null;
  let intoBtn = null;
  let overBtn = null;
  let stopBtn = null;
  let phase = 'idle';
  let generation = 0;
  let watchGen = 0;
  let suppressOutput = false;
  let stopRequested = false;
  let run = null;
  let selectedLine = 0;
  let selectedWatch = -1;
  let selectedFrame = null;
  let currentMark = null;
  let conditionPanel = null;
  let panelCloser = null;

  function sameModule(name) {
    return String(name || '').toLowerCase() === MODULE;
  }

  function appendOutput(text) {
    if (!output) return;
    const line = text == null ? '' : String(text);
    if (!line) return;
    output.value = output.value ? output.value + '\n' + line : line;
    output.scrollTop = output.scrollHeight;
  }

  function setPhase(next) {
    phase = next;
    const idle = next === 'idle';
    const paused = next === 'paused';
    if (startBtn) startBtn.disabled = !idle;
    if (contBtn) contBtn.disabled = !paused;
    if (intoBtn) intoBtn.disabled = !paused;
    if (overBtn) overBtn.disabled = !paused;
    if (stopBtn) stopBtn.disabled = idle;
  }

  function sourceLine(moduleName, line) {
    if (sameModule(moduleName) && editor) {
      const row = editor.getLine(line - 1);
      return row == null ? '' : row;
    }
    const mod = run && run.vm.program.modules.get(String(moduleName || '').toLowerCase());
    if (!mod || mod.source == null) return '';
    const rows = String(mod.source).split('\n');
    const row = rows[line - 1];
    return row == null ? '' : String(row).replace(/\r$/, '');
  }

  function stackLabel(line, procedure) {
    return L.t('Line {0}| {1}', line, procedure == null ? '' : procedure);
  }

  function objectValue(entry) {
    const type = entry.type || '';
    if (entry.expandable || type === 'Object' || type === 'Collection' || type === 'Dictionary') return L.t('[Object]');
    return entry.value == null ? '' : String(entry.value);
  }

  function shownValue(value) {
    if (run && run.host && run.host.engine && run.host.engine.describe) return run.host.engine.describe(value);
    return String(value);
  }

  function breakpointRecord(item) {
    if (typeof item === 'number') {
      if (!Number.isInteger(item) || item < 1) return null;
      return { line: item, condition: '', enabled: true };
    }
    if (!item || typeof item !== 'object') return null;
    const line = Number(item.line);
    if (!Number.isInteger(line) || line < 1) return null;
    return { line: line, condition: item.condition == null ? '' : String(item.condition), enabled: item.enabled !== false };
  }

  function watchRecord(item) {
    if (!item || typeof item !== 'object') return null;
    const expressionText = item.expression == null ? '' : String(item.expression);
    if (!expressionText.trim()) return null;
    return {
      expression: expressionText,
      enabled: item.enabled !== false,
      value: item.value == null ? '' : String(item.value),
      lastValue: item.lastValue == null ? '' : String(item.lastValue),
      evaluated: item.value != null && String(item.value) !== '',
      changed: false,
    };
  }

  function bpAt(line) {
    for (let i = 0; i < breakpoints.length; i++) if (breakpoints[i].line === line) return breakpoints[i];
    return null;
  }

  function jumpToLine(line) {
    if (!editor) return;
    if (!Number.isInteger(line) || line < 1 || line > editor.lineCount()) return;
    editor.setCursor({ line: line - 1, ch: 0 });
    editor.scrollIntoView({ line: line - 1, ch: 0 }, 80);
    editor.focus();
  }

  function clearCurrent() {
    if (currentMark && editor) editor.removeLineClass(currentMark, 'background', 'vb-ide-current');
    currentMark = null;
  }

  function markCurrent(line) {
    clearCurrent();
    if (!editor || !Number.isInteger(line) || line < 1 || line > editor.lineCount()) return;
    currentMark = editor.addLineClass(line - 1, 'background', 'vb-ide-current');
    editor.scrollIntoView({ line: line - 1, ch: 0 }, 60);
  }

  function closeConditionPanel() {
    if (panelCloser) { const close = panelCloser; panelCloser = null; close(); }
    if (conditionPanel) { conditionPanel.remove(); conditionPanel = null; }
  }

  function syncBreakpoint(bp) {
    if (!run || run.ignore || !bp) return;
    try {
      if (bp.enabled === false) run.vm.removeBreakpoint(run.moduleName, bp.line);
      else run.vm.setBreakpoint(run.moduleName, bp.line, bp.condition || '', true);
    } catch (e) {
      try { run.vm.removeBreakpoint(run.moduleName, bp.line); } catch (gone) { /* not installed */ }
      appendOutput(L.t('Line {0}: {1}', bp.line, e && e.message ? e.message : e));
    }
  }

  function paintDots() {
    if (!editor) return;
    editor.eachLine((handle) => editor.setGutterMarker(handle, 'vb-ide-bp-gutter', null));
    for (let i = 0; i < breakpoints.length; i++) {
      const bp = breakpoints[i];
      if (bp.line < 1 || bp.line > editor.lineCount()) continue;
      const dot = document.createElement('span');
      dot.className = 'vb-ide-dot' + (bp.enabled === false ? ' off' : '');
      editor.setGutterMarker(bp.line - 1, 'vb-ide-bp-gutter', dot);
    }
  }

  function cell(text, data) {
    return L.h('td', { translate: data ? 'no' : null, text: text == null ? '' : String(text) });
  }

  function paintBpList() {
    if (!bpBody) return;
    bpBody.replaceChildren();
    for (let i = 0; i < breakpoints.length; i++) {
      const bp = breakpoints[i];
      const tr = L.h('tr', { class: bp.line === selectedLine ? 'on' : '' });
      const box = L.h('input', { type: 'checkbox' });
      box.checked = bp.enabled !== false;
      box.addEventListener('click', (e) => e.stopPropagation());
      box.addEventListener('change', () => {
        bp.enabled = box.checked;
        selectedLine = bp.line;
        syncBreakpoint(bp);
        paintDots();
        paintBpList();
      });
      tr.appendChild(L.h('td', {}, box));
      tr.appendChild(cell(bp.line, true));
      tr.appendChild(cell(bp.condition, true));
      tr.appendChild(cell(sourceLine(MODULE, bp.line), true));
      tr.addEventListener('click', () => {
        selectedLine = bp.line;
        paintBpList();
        jumpToLine(bp.line);
      });
      bpBody.appendChild(tr);
    }
  }

  function paintWatches() {
    if (!watchBody) return;
    watchBody.replaceChildren();
    for (let i = 0; i < watches.length; i++) {
      const watch = watches[i];
      const tr = L.h('tr', { class: (i === selectedWatch ? 'on ' : '') + (watch.changed ? 'vb-ide-watch-changed' : '') });
      const box = L.h('input', { type: 'checkbox' });
      box.checked = watch.enabled !== false;
      box.addEventListener('click', (e) => e.stopPropagation());
      box.addEventListener('change', () => {
        watch.enabled = box.checked;
        selectedWatch = i;
        paintWatches();
      });
      tr.appendChild(L.h('td', {}, box));
      tr.appendChild(cell(watch.expression, true));
      tr.appendChild(cell(watch.value, true));
      tr.appendChild(cell(watch.lastValue, true));
      tr.addEventListener('click', () => {
        selectedWatch = i;
        paintWatches();
      });
      watchBody.appendChild(tr);
    }
  }

  function rememberBreakpoint(bp) {
    const existing = bpAt(bp.line);
    if (existing) {
      existing.condition = bp.condition;
      return existing;
    }
    breakpoints.push(bp);
    breakpoints.sort((a, b) => a.line - b.line);
    return bp;
  }

  function toggleBreakpoint(line) {
    if (!Number.isInteger(line) || line < 1) return;
    const at = breakpoints.findIndex((bp) => bp.line === line);
    if (at >= 0) {
      breakpoints.splice(at, 1);
      if (selectedLine === line) selectedLine = 0;
      if (run && !run.ignore) {
        try { run.vm.removeBreakpoint(run.moduleName, line); } catch (e) { /* already gone */ }
      }
    } else {
      const bp = rememberBreakpoint({ line: line, condition: '', enabled: true });
      selectedLine = line;
      syncBreakpoint(bp);
    }
    paintDots();
    paintBpList();
  }

  function installBreakpoints() {
    if (!run || run.ignore) return;
    for (let i = 0; i < breakpoints.length; i++) syncBreakpoint(breakpoints[i]);
  }

  function openConditionPanel(line, x, y) {
    const bp = bpAt(line);
    if (!bp) return;
    closeConditionPanel();
    const input = L.h('input', { type: 'text', translate: 'no' });
    input.value = bp.condition || '';
    const ok = L.h('button', { class: 'btn', type: 'button', text: 'OK' });
    const cancel = L.h('button', { class: 'btn', type: 'button', text: 'Cancel' });
    const panel = L.h('div', { class: 'vb-ide-cond-panel' },
      L.h('div', { class: 'vb-ide-cond-row' }, L.h('span', { text: 'Line Number:' }), L.h('span', { translate: 'no', text: String(line) })),
      L.h('div', { class: 'vb-ide-cond-row' }, L.h('label', { text: 'Condition:' }), input),
      L.h('div', { class: 'vb-ide-cond-buttons' }, ok, cancel));
    panel.style.left = Math.max(8, x) + 'px';
    panel.style.top = Math.max(8, y) + 'px';
    document.body.appendChild(panel);
    const rect = panel.getBoundingClientRect();
    if (rect.right > window.innerWidth - 8) panel.style.left = Math.max(8, window.innerWidth - rect.width - 8) + 'px';
    if (rect.bottom > window.innerHeight - 8) panel.style.top = Math.max(8, window.innerHeight - rect.height - 8) + 'px';
    conditionPanel = panel;
    const commit = () => {
      bp.condition = input.value.trim();
      syncBreakpoint(bp);
      paintDots();
      paintBpList();
      closeConditionPanel();
    };
    ok.addEventListener('click', (e) => { e.stopPropagation(); commit(); });
    cancel.addEventListener('click', (e) => { e.stopPropagation(); closeConditionPanel(); });
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); commit(); }
      if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); closeConditionPanel(); }
    });
    const away = (e) => { if (panel.contains(e.target)) return; closeConditionPanel(); };
    const onKey = (e) => {
      if (e.key !== 'Escape') return;
      e.preventDefault();
      e.stopPropagation();
      closeConditionPanel();
    };
    panelCloser = () => {
      document.removeEventListener('pointerdown', away, true);
      document.removeEventListener('keydown', onKey, true);
    };
    setTimeout(() => {
      if (conditionPanel !== panel) return;
      document.addEventListener('pointerdown', away, true);
      document.addEventListener('keydown', onKey, true);
      input.focus();
      input.select();
    }, 0);
  }

  function paintStack() {
    if (!stackList) return;
    stackList.replaceChildren();
    if (!run || phase !== 'paused') return;
    let frames = [];
    try { frames = run.vm.debugStack(); } catch (e) { frames = []; }
    for (let i = frames.length - 1; i >= 0; i--) {
      const frame = frames[i];
      const li = L.h('li', {
        class: frame.index === selectedFrame ? 'on' : '',
        translate: 'no',
        text: stackLabel(frame.line, frame.procedure),
      });
      li.addEventListener('click', () => {
        selectedFrame = frame.index;
        paintStack();
        paintLocals(frame.index);
        refreshWatches();
        if (sameModule(frame.module)) jumpToLine(frame.line);
      });
      stackList.appendChild(li);
    }
  }

  function paintLocals(frameIndex) {
    if (!localList) return;
    localList.replaceChildren();
    if (!run || phase !== 'paused') return;
    let rows = [];
    try { rows = run.vm.debugLocals({ frameIndex: frameIndex == null ? null : frameIndex, includeFields: true }); }
    catch (e) { rows = []; }
    for (let i = 0; i < rows.length; i++) {
      const entry = rows[i];
      localList.appendChild(L.h('li', {
        translate: 'no',
        text: entry.name + ' = ' + objectValue(entry),
      }));
    }
  }

  function clearRuntimeLists() {
    selectedFrame = null;
    watchGen++;
    suppressOutput = false;
    for (let i = 0; i < watches.length; i++) watches[i].changed = false;
    if (stackList) stackList.replaceChildren();
    if (localList) localList.replaceChildren();
    paintWatches();
  }

  async function refreshWatches() {
    const token = ++watchGen;
    const frame = selectedFrame;
    if (!run || phase !== 'paused') {
      for (let i = 0; i < watches.length; i++) watches[i].changed = false;
      paintWatches();
      return;
    }
    suppressOutput = true;
    try {
      for (let i = 0; i < watches.length; i++) {
        if (token !== watchGen) return;
        const watch = watches[i];
        if (watch.enabled === false) { watch.changed = false; continue; }
        const expr = String(watch.expression || '').trim().replace(/^\?/, '');
        if (!expr) { watch.changed = false; continue; }
        const previous = watch.value;
        const had = watch.evaluated === true;
        try {
          const value = await run.vm.immediate('?' + expr, { frameIndex: frame });
          if (token !== watchGen || phase !== 'paused') return;
          const shown = shownValue(value);
          watch.lastValue = previous;
          watch.value = shown;
          watch.changed = had && shown !== previous;
          watch.evaluated = true;
        } catch (e) {
          if (token !== watchGen || phase !== 'paused') return;
          const message = e && e.message ? e.message : String(e);
          watch.lastValue = previous;
          watch.value = message;
          watch.changed = had && message !== previous;
          watch.evaluated = true;
        }
      }
    } finally {
      if (token === watchGen) {
        suppressOutput = false;
        paintWatches();
      }
    }
  }

  function onPause(e) {
    setPhase('paused');
    const ins = e.instruction || {};
    selectedFrame = null;
    if (sameModule(ins.source)) markCurrent(ins.line);
    else clearCurrent();
    paintStack();
    paintLocals(null);
    if (e.reason === 'breakpoint-condition' && e.conditionError) appendOutput(e.conditionError);
    else if (e.reason === 'error' && e.error && e.error.message) appendOutput(e.error.message);
    refreshWatches();
  }

  function go(mode) {
    if (phase !== 'paused' || !run) return;
    watchGen++;
    suppressOutput = false;
    setPhase('running');
    try { run.vm.resume(mode); }
    catch (e) {
      appendOutput(e && e.message ? e.message : e);
      setPhase('paused');
    }
  }

  async function start(ignoreBreakpoints) {
    if (phase !== 'idle' || !editor) return;
    setPhase('running');
    stopRequested = false;
    clearCurrent();
    clearRuntimeLists();
    if (output) output.value = '';
    const token = ++generation;
    let host = null;
    let begun = null;
    try {
      if (!L.vb || !L.vb.session) throw new Error(L.t('The VB runtime is not running'));
      host = await L.vb.session();
      if (token !== generation || stopRequested) return;
      begun = await host.debugBegin(editor.getValue());
      if (token !== generation || stopRequested) return;
      run = { vm: begun.vm, host: host, moduleName: begun.moduleName, ignore: !!ignoreBreakpoints };
      begun.vm.on('output', (ev) => { if (token === generation && !suppressOutput) appendOutput(ev.text); });
      begun.vm.on('pause', (ev) => { if (token === generation) onPause(ev); });
      if (!ignoreBreakpoints) installBreakpoints();
      if (stopRequested) return;
      await begun.vm.start();
    } catch (e) {
      if (token === generation) appendOutput(e && e.message ? e.message : String(e));
    } finally {
      if (begun && host) host.debugEnd();
      if (token === generation) {
        run = null;
        clearCurrent();
        clearRuntimeLists();
        setPhase('idle');
      }
    }
  }

  function stopDebugging() {
    stopRequested = true;
    if (run) {
      try { run.vm.stop(); } catch (e) { /* already stopped */ }
    }
  }

  async function evaluate() {
    expression = exprInput ? exprInput.value : expression;
    if (!run || phase !== 'paused') {
      const message = L.t('The program is not paused.');
      expressionResult = message;
      if (exprOut) exprOut.value = message;
      return;
    }
    try {
      const value = await run.vm.immediate(expression, { frameIndex: selectedFrame });
      const shown = shownValue(value);
      expressionResult = shown;
      if (exprOut) exprOut.value = shown;
      paintLocals(selectedFrame);
    } catch (e) {
      const message = e && e.message ? e.message : String(e);
      expressionResult = message;
      if (exprOut) exprOut.value = message;
    }
  }

  function projectData() {
    return {
      schema: 1,
      kind: 'vibeoffice.vb6.ide',
      module: 'Module1',
      code: editor ? editor.getValue() : code,
      breakpoints: breakpoints.map((bp) => ({ line: bp.line, condition: bp.condition || '', enabled: bp.enabled !== false })),
      watches: watches.map((watch) => ({
        expression: watch.expression,
        enabled: watch.enabled !== false,
        value: watch.value || '',
        lastValue: watch.lastValue || '',
      })),
      expression: exprInput ? exprInput.value : expression,
      expressionResult: exprOut ? exprOut.value : expressionResult,
    };
  }

  function applyProject(data) {
    code = String(data.code || '');
    breakpoints = (Array.isArray(data.breakpoints) ? data.breakpoints : []).map(breakpointRecord).filter(Boolean);
    watches = (Array.isArray(data.watches) ? data.watches : []).map(watchRecord).filter(Boolean);
    expression = String(data.expression || '');
    expressionResult = String(data.expressionResult || '');
    selectedLine = 0;
    selectedWatch = -1;
    if (editor) editor.setValue(code);
    if (exprInput) exprInput.value = expression;
    if (exprOut) exprOut.value = expressionResult;
    paintDots();
    paintBpList();
    paintWatches();
  }

  async function openProject() {
    if (phase !== 'idle') return;
    const files = await L.pickFiles('.json,application/json', false);
    if (!files || !files.length) return;
    let data;
    try { data = JSON.parse(await L.readAsText(files[0])); }
    catch (e) {
      appendOutput(L.t('This file is not a Visual Basic project.'));
      return;
    }
    if (!data || data.kind !== 'vibeoffice.vb6.ide' || data.schema !== 1) {
      appendOutput(L.t('This file is not a Visual Basic project.'));
      return;
    }
    projectName = files[0].name || 'Project.json';
    applyProject(data);
  }

  async function saveProject() {
    const name = projectName || 'Project.json';
    const blob = new Blob([JSON.stringify(projectData(), null, 2)], { type: 'application/json' });
    await L.saveFile(name, blob);
  }

  async function openVisualBasicFile() {
    const files = await L.pickFiles('.bas,.cls,.frm,.vbs,.txt,text/plain', false);
    if (!files || !files.length) return;
    code = await L.readAsText(files[0]);
    if (editor) editor.setValue(code);
    paintDots();
    paintBpList();
  }

  function exitIde() {
    stopDebugging();
    closeConditionPanel();
    if (openDlg && openDlg.isConnected) {
      const closer = openDlg.querySelector('.dlg-x');
      if (closer) closer.click();
    }
  }

  function addBreakpoint() {
    const line = parseInt(lineInput && lineInput.value, 10);
    if (!Number.isInteger(line) || line < 1) return;
    const condition = condInput ? condInput.value.trim() : '';
    const bp = rememberBreakpoint({ line: line, condition: condition, enabled: true });
    if (bpAt(line) === bp && bp.condition !== condition) bp.condition = condition;
    bp.condition = condition;
    selectedLine = line;
    syncBreakpoint(bp);
    paintDots();
    paintBpList();
    jumpToLine(line);
  }

  function removeBreakpoint() {
    if (!selectedLine) return;
    const line = selectedLine;
    if (!bpAt(line)) return;
    toggleBreakpoint(line);
  }

  function removeAllBreakpoints() {
    const lines = breakpoints.map((bp) => bp.line);
    breakpoints = [];
    selectedLine = 0;
    if (run && !run.ignore) {
      for (let i = 0; i < lines.length; i++) {
        try { run.vm.removeBreakpoint(run.moduleName, lines[i]); } catch (e) { /* already gone */ }
      }
    }
    paintDots();
    paintBpList();
  }

  function addWatch() {
    const text = watchInput ? watchInput.value.trim() : '';
    if (!text) return;
    watches.push({ expression: text, enabled: true, value: '', lastValue: '', evaluated: false, changed: false });
    selectedWatch = watches.length - 1;
    paintWatches();
    if (phase === 'paused') refreshWatches();
  }

  function removeWatch() {
    if (selectedWatch < 0 || selectedWatch >= watches.length) return;
    watches.splice(selectedWatch, 1);
    selectedWatch = -1;
    paintWatches();
  }

  function removeAllWatches() {
    watches = [];
    selectedWatch = -1;
    paintWatches();
  }

  function menuButton(label, items) {
    const button = L.h('button', { class: 'vb-ide-menu-btn', type: 'button', text: label });
    button.addEventListener('click', () => {
      const rect = button.getBoundingClientRect();
      L.ui.openMenu(items, { left: rect.left, bottom: rect.bottom });
    });
    return button;
  }

  function showPage(pages, buttons, name) {
    for (let i = 0; i < pages.length; i++) pages[i].el.hidden = pages[i].name !== name;
    for (let i = 0; i < buttons.length; i++) buttons[i].classList.toggle('on', buttons[i].dataset.page === name);
  }

  function applyEditorShare() {
    if (editorHost) editorHost.style.flexGrow = String(editorShare);
    if (lowerPane) lowerPane.style.flexGrow = '1';
  }

  function grid(headers) {
    const body = L.h('tbody');
    const head = L.h('tr');
    for (let i = 0; i < headers.length; i++) head.appendChild(L.h('th', { text: headers[i] }));
    const table = L.h('table', { class: 'vb-ide-grid' }, L.h('thead', {}, head), body);
    return { wrap: L.h('div', { class: 'vb-ide-grid-wrap' }, table), body: body };
  }

  function buildBody() {
    const menu = L.h('div', { class: 'vb-ide-menu' },
      menuButton('File', () => [
        { label: '&Open Project...', enabled: () => phase === 'idle', run: openProject },
        { label: '&Save Project...', run: saveProject },
        '-',
        { label: '&Open Visual Basic File...', run: openVisualBasicFile },
        '-',
        { label: '&Exit', run: exitIde },
      ]),
      menuButton('Debug', () => [
        { label: 'Start &Debug', enabled: () => phase === 'idle', run: () => start(false) },
        { label: 'Start &Run', enabled: () => phase === 'idle', run: () => start(true) },
        '-',
        { label: 'Step &Over', enabled: () => phase === 'paused', run: () => go('over') },
        { label: 'Step &Into', enabled: () => phase === 'paused', run: () => go('into') },
        { label: '&Continue', enabled: () => phase === 'paused', run: () => go('continue') },
        { label: '&Stop Debugging', enabled: () => phase !== 'idle', run: stopDebugging },
      ]));

    startBtn = L.h('button', { class: 'btn', type: 'button', text: 'Start Debug(F5)' });
    contBtn = L.h('button', { class: 'btn', type: 'button', text: 'Continue' });
    intoBtn = L.h('button', { class: 'btn', type: 'button', text: 'Step In' });
    overBtn = L.h('button', { class: 'btn', type: 'button', text: 'Step Over' });
    stopBtn = L.h('button', { class: 'btn', type: 'button', text: 'Stop Debugging' });
    startBtn.addEventListener('click', () => start(false));
    contBtn.addEventListener('click', () => go('continue'));
    intoBtn.addEventListener('click', () => go('into'));
    overBtn.addEventListener('click', () => go('over'));
    stopBtn.addEventListener('click', stopDebugging);
    const bar = L.h('div', { class: 'vb-ide-bar' }, startBtn, L.h('span', { class: 'vb-ide-sep' }), contBtn, intoBtn, overBtn, stopBtn);

    editorHost = L.h('div', { class: 'vb-ide-editor' });
    const area = L.h('textarea', { translate: 'no' });

    output = L.h('textarea', { class: 'vb-ide-output', readonly: true, translate: 'no' });
    const outputPage = L.h('div', { class: 'vb-ide-page' }, output);

    exprInput = L.h('input', { class: 'vb-ide-expr', type: 'text', translate: 'no' });
    exprInput.value = expression;
    exprInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); evaluate(); }
    });
    const evalBtn = L.h('button', { class: 'btn', type: 'button', text: 'Evaluate' });
    evalBtn.addEventListener('click', evaluate);
    exprOut = L.h('textarea', { class: 'vb-ide-expr-out', readonly: true, translate: 'no' });
    exprOut.value = expressionResult;
    const expressionPage = L.h('div', { class: 'vb-ide-page', hidden: true },
      L.h('div', { class: 'vb-ide-row' },
        L.h('label', { text: 'Expression' }),
        exprInput,
        evalBtn),
      exprOut);

    lineInput = L.h('input', { class: 'vb-ide-line', type: 'text', translate: 'no' });
    condInput = L.h('input', { class: 'vb-ide-cond', type: 'text', translate: 'no', placeholder: L.t('(VB bool expression. Leave empty to set unconditional breakpoint)') });
    const addBtn = L.h('button', { class: 'btn', type: 'button', text: 'Add' });
    addBtn.addEventListener('click', addBreakpoint);
    const bpGrid = grid(['Enabled', 'Line Number', 'Condition Expression', 'Code at Line']);
    bpBody = bpGrid.body;
    const removeBtn = L.h('button', { class: 'btn', type: 'button', text: 'Remove Breakpoint' });
    const removeAllBtn = L.h('button', { class: 'btn', type: 'button', text: 'Remove All Breakpoints' });
    removeBtn.addEventListener('click', removeBreakpoint);
    removeAllBtn.addEventListener('click', removeAllBreakpoints);
    const breakPage = L.h('div', { class: 'vb-ide-page', hidden: true },
      L.h('div', { class: 'vb-ide-row' },
        L.h('label', { text: 'Line Number' }),
        lineInput,
        addBtn),
      L.h('div', { class: 'vb-ide-row' },
        L.h('label', { text: 'Condition Expression' }),
        condInput),
      bpGrid.wrap,
      L.h('div', { class: 'vb-ide-row' }, removeBtn, removeAllBtn));

    watchInput = L.h('input', { class: 'vb-ide-expr', type: 'text', translate: 'no' });
    watchInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); addWatch(); }
    });
    const watchAdd = L.h('button', { class: 'btn', type: 'button', text: 'Add' });
    watchAdd.addEventListener('click', addWatch);
    const watchGrid = grid(['Enabled', 'Expression', 'Value', 'Last Value']);
    watchBody = watchGrid.body;
    const watchRemove = L.h('button', { class: 'btn', type: 'button', text: 'Remove' });
    const watchRemoveAll = L.h('button', { class: 'btn', type: 'button', text: 'Remove All' });
    watchRemove.addEventListener('click', removeWatch);
    watchRemoveAll.addEventListener('click', removeAllWatches);
    const watchPage = L.h('div', { class: 'vb-ide-page', hidden: true },
      L.h('div', { class: 'vb-ide-row' },
        L.h('label', { text: 'Expression' }),
        watchInput,
        watchAdd),
      watchGrid.wrap,
      L.h('div', { class: 'vb-ide-row' }, watchRemove, watchRemoveAll));

    const pages = [
      { name: 'output', el: outputPage },
      { name: 'expression', el: expressionPage },
      { name: 'breakpoints', el: breakPage },
      { name: 'watch', el: watchPage },
    ];
    const tabButtons = [
      ['Output', 'output'],
      ['Expression', 'expression'],
      ['Breakpoints', 'breakpoints'],
      ['Watch', 'watch'],
    ].map((pair) => {
      const button = L.h('button', { class: 'vb-ide-tab', type: 'button', text: pair[0], dataset: { page: pair[1] } });
      button.addEventListener('click', () => showPage(pages, tabButtons, pair[1]));
      return button;
    });
    showPage(pages, tabButtons, 'output');

    stackList = L.h('ul', { class: 'vb-ide-list' });
    localList = L.h('ul', { class: 'vb-ide-list' });
    lowerPane = L.h('div', { class: 'vb-ide-lower' },
      L.h('div', { class: 'vb-ide-tabs' }, L.h('div', { class: 'vb-ide-tabbar' }, tabButtons), outputPage, expressionPage, breakPage, watchPage),
      L.h('div', { class: 'vb-ide-col' }, L.h('div', { class: 'vb-ide-cap', text: 'Call Stack' }), stackList),
      L.h('div', { class: 'vb-ide-col' }, L.h('div', { class: 'vb-ide-cap', text: 'Object Values' }), localList));

    const split = L.h('div', { class: 'splitter h vb-ide-split', role: 'separator' });
    split.addEventListener('pointerdown', (e) => {
      if (e.button !== 0 || !editorHost || !lowerPane) return;
      e.preventDefault();
      L.ui.drag(e, {
        threshold: 0,
        move(ev) {
          const top = editorHost.getBoundingClientRect().top;
          const bottom = lowerPane.getBoundingClientRect().bottom;
          const y = L.clamp(ev.clientY, top + 80, Math.max(top + 80, bottom - 100));
          editorShare = (y - top) / Math.max(1, bottom - y);
          applyEditorShare();
          if (editor) editor.refresh();
        },
      });
    });
    applyEditorShare();

    return { root: L.h('div', { class: 'vb-ide' }, menu, bar, editorHost, split, lowerPane), editorHost: editorHost, area: area };
  }

  function open() {
    if (typeof document === 'undefined' || !L.ui || typeof L.ui.dialog !== 'function') return;
    if (openDlg && openDlg.isConnected) {
      try { openDlg.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true })); } catch (e) { /* raise is enough */ }
      if (editor) editor.focus();
      return;
    }
    if (!window.CodeMirror) {
      L.ui.msg('The script editor did not load.');
      return;
    }
    const built = buildBody();
    L.ui.dialog({
      title: 'VibeOffice Visual Basic',
      body: built.root,
      width: 820,
      buttons: [],
      modeless: true,
      resizeable: true,
      minWidth: 520,
      minHeight: 360,
      onResize: () => { if (editor) editor.refresh(); },
      onOpen: (el) => {
        el.classList.add('vb-ide-dlg');
        el.style.height = Math.min(560, window.innerHeight - 24) + 'px';
        el.addEventListener('keydown', (e) => {
          if (e.key !== 'F5' || e.ctrlKey || e.altKey || e.shiftKey || e.metaKey) return;
          e.preventDefault();
          e.stopPropagation();
          start(false);
        }, true);
        openDlg = el;
        built.editorHost.appendChild(built.area);
        editor = window.CodeMirror.fromTextArea(built.area, {
          mode: 'text/x-vb',
          lineNumbers: true,
          indentUnit: 2,
          tabSize: 2,
          gutters: [{ className: 'vb-ide-bp-gutter', style: 'width: 14px' }, 'CodeMirror-linenumbers'],
        });
        editor.on('gutterClick', (cm, n, gutter, event) => {
          if (event && event.button === 2) return;
          if (gutter !== 'vb-ide-bp-gutter' && gutter !== 'CodeMirror-linenumbers') return;
          toggleBreakpoint(n + 1);
        });
        editor.on('gutterContextMenu', (cm, n, gutter, event) => {
          if (gutter !== 'vb-ide-bp-gutter' || !bpAt(n + 1)) return;
          if (event) { event.preventDefault(); event.stopPropagation(); }
          openConditionPanel(n + 1, event ? event.clientX : 0, event ? event.clientY : 0);
        });
        let paintQueued = false;
        editor.on('change', () => {
          code = editor.getValue();
          if (paintQueued) return;
          paintQueued = true;
          setTimeout(() => {
            paintQueued = false;
            if (!editor) return;
            paintDots();
            paintBpList();
          }, 0);
        });
        editor.setValue(code);
        setPhase(phase);
        paintDots();
        paintBpList();
        paintWatches();
        setTimeout(() => { if (editor) editor.refresh(); }, 0);
      },
      onClose: () => {
        closeConditionPanel();
        if (editor) code = editor.getValue();
        if (exprInput) expression = exprInput.value;
        if (exprOut) expressionResult = exprOut.value;
        editor = null;
        openDlg = null;
        output = null;
        editorHost = null;
        lowerPane = null;
        bpBody = null;
        watchBody = null;
        stopDebugging();
      },
    });
  }

  L.vbIde = { open: open };
})();
