/* Visual Basic IDE. One standard module, debugged on the app host's debug VM.
   Opened from View. The window is the ProjectEditor layout: editor, output,
   expression, breakpoints, watches, call stack, object values. */
(function () {
  'use strict';
  const L = window.L;
  const MODULE = 'module1';
  const SESSION = 'vbIde';
  const EDITOR_SHARE = 'vbIdeEditor';
  const BLANK_CODE = "Sub Main\n    ' Insert your code here\n    \nEnd Sub";

  let code = '';
  let outputText = '';
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
  let runBtn = null;
  let contBtn = null;
  let intoBtn = null;
  let overBtn = null;
  let stopBtn = null;
  let phase = 'idle';
  let pauseReason = '';
  let statusEl = null;
  let generation = 0;
  let watchGen = 0;
  let suppressOutput = false;
  let stopRequested = false;
  let errorBody = null;
  let showErrors = null;
  let diagnostics = [];
  let selectedError = -1;
  let squiggles = [];
  let checkTimer = null;
  let checkGen = 0;
  let starting = false;
  let run = null;
  let selectedLine = 0;
  let selectedWatch = -1;
  let selectedFrame = null;
  let currentMark = null;
  let conditionPanel = null;
  let panelCloser = null;
  let listEl = null;
  let sigEl = null;
  let listItems = [];
  let listIndex = 0;
  let listSpan = null;
  let listTarget = null;
  let listGen = 0;
  let accepting = false;
  let hoverTip = null;
  let hoverTimer = null;
  let hoverGen = 0;
  let hoverKey = '';
  let hoverPoint = null;
  let linkMark = null;
  let linkOff = null;

  function sameModule(name) {
    return String(name || '').toLowerCase() === MODULE;
  }

  function appendOutput(text) {
    if (!output) return;
    const line = text == null ? '' : String(text);
    if (!line) return;
    output.value = output.value ? output.value + '\n' + line : line;
    output.scrollTop = output.scrollHeight;
    outputText = output.value;
  }

  function statusText() {
    if (phase === 'running') return L.t('Running');
    if (phase === 'paused') {
      if (pauseReason === 'breakpoint' || pauseReason === 'breakpoint-condition' || pauseReason === 'run-to-cursor') return L.t('Triggered Breakpoint');
      if (pauseReason === 'error') return L.t('Triggered Error Interrupt');
      if (pauseReason.indexOf('watch:') === 0) return L.t('Triggered Watch');
      return L.t('Break');
    }
    return L.t('Ready');
  }

  function paintStatus() {
    if (statusEl) statusEl.textContent = statusText();
  }

  function setPhase(next) {
    phase = next;
    if (next !== 'paused') pauseReason = '';
    const idle = next === 'idle';
    const paused = next === 'paused';
    if (startBtn) startBtn.disabled = !idle;
    if (runBtn) runBtn.disabled = !idle;
    if (contBtn) contBtn.disabled = !paused;
    if (intoBtn) intoBtn.disabled = !paused;
    if (overBtn) overBtn.disabled = !paused;
    if (stopBtn) stopBtn.disabled = idle;
    if (next !== 'paused' && listTarget && listTarget.kind === 'input') closeList();
    paintStatus();
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

  function jumpToLine(line, column) {
    if (!editor) return;
    if (!Number.isInteger(line) || line < 1 || line > editor.lineCount()) return;
    const ch = Math.max(0, (Number(column) || 1) - 1);
    editor.setCursor({ line: line - 1, ch: ch });
    editor.scrollIntoView({ line: line - 1, ch: ch }, 80);
    editor.focus();
  }

  function clearSquiggles() {
    for (let i = 0; i < squiggles.length; i++) {
      try { squiggles[i].clear(); } catch (e) { /* the editor already dropped it */ }
    }
    squiggles = [];
  }

  function paintSquiggles() {
    clearSquiggles();
    if (!editor) return;
    for (let i = 0; i < diagnostics.length; i++) {
      const d = diagnostics[i];
      if (String(d.source || '').toLowerCase() !== MODULE) continue;
      const line = Number(d.line) || 1;
      if (!Number.isInteger(line) || line < 1 || line > editor.lineCount()) continue;
      const text = editor.getLine(line - 1);
      if (!text.length) {
        const widget = document.createElement('span');
        widget.className = 'vb-ide-squiggle vb-ide-squiggle-empty';
        widget.textContent = ' ';
        squiggles.push(editor.setBookmark({ line: line - 1, ch: 0 }, { widget: widget }));
        continue;
      }
      let ch = Math.max(0, (Number(d.column) || 1) - 1);
      if (ch >= text.length) ch = text.length - 1;
      squiggles.push(editor.markText({ line: line - 1, ch: ch }, { line: line - 1, ch: text.length }, {
        className: 'vb-ide-squiggle',
      }));
    }
  }

  function paintErrors() {
    if (!errorBody) return;
    errorBody.replaceChildren();
    for (let i = 0; i < diagnostics.length; i++) {
      const d = diagnostics[i];
      const tr = L.h('tr', { class: i === selectedError ? 'on' : '' });
      tr.appendChild(cell(d.line || '', true));
      tr.appendChild(cell(d.column || '', true));
      tr.appendChild(cell(d.number || '', true));
      tr.appendChild(cell(d.source || '', true));
      tr.appendChild(cell(d.message || '', true));
      tr.addEventListener('click', () => {
        selectedError = i;
        paintErrors();
        if (String(d.source || '').toLowerCase() === MODULE) jumpToLine(Number(d.line) || 1, d.column);
      });
      errorBody.appendChild(tr);
    }
  }

  function applyDiagnostics(list) {
    diagnostics = Array.isArray(list) ? list : [];
    selectedError = -1;
    paintSquiggles();
    paintErrors();
  }

  function diagnosticText(list) {
    const lines = [];
    for (let i = 0; i < list.length; i++) {
      const d = list[i];
      lines.push(L.t('{0}: line {1}, column {2}: {3} ({4})', d.source || 'Module1', d.line || 1, d.column || 1, d.message || '', d.number == null ? '' : d.number));
    }
    return lines.length ? lines.join('\n') : L.t('The code could not be compiled.');
  }

  async function reportDiagnostics(list, title) {
    if (L.ui && L.ui.msg) await L.ui.msg(diagnosticText(list), { icon: 'error', title: title || 'Melina Office Visual Basic', width: 520 });
  }

  function scheduleCheck() {
    if (checkTimer) clearTimeout(checkTimer);
    const gen = ++checkGen;
    checkTimer = setTimeout(() => {
      checkTimer = null;
      runCheck(gen);
    }, 300);
  }

  async function runCheck(gen) {
    if (!editor || !L.vb || !L.vb.compileIde) return;
    const source = editor.getValue();
    let compiled;
    try { compiled = await L.vb.compileIde(source); }
    catch (e) { return; }
    if (gen !== checkGen || !editor || editor.getValue() !== source) return;
    applyDiagnostics(compiled.diagnostics);
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
    editor.clearGutter('vb-ide-bp-gutter');
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
    const enabledBox = L.h('input', { type: 'checkbox' });
    enabledBox.checked = bp.enabled !== false;
    const ok = L.h('button', { class: 'btn', type: 'button', text: 'OK' });
    const cancel = L.h('button', { class: 'btn', type: 'button', text: 'Cancel' });
    const panel = L.h('div', { class: 'vb-ide-cond-panel' },
      L.h('div', { class: 'vb-ide-cond-row' }, L.h('span', { text: 'Line Number:' }), L.h('span', { translate: 'no', text: String(line) })),
      L.h('div', { class: 'vb-ide-cond-row' }, L.h('label', { text: 'Condition:' }), input),
      L.h('label', { class: 'vb-ide-cond-enable' }, enabledBox, 'Enable this breakpoint'),
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
      bp.enabled = enabledBox.checked;
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
    pauseReason = e && e.reason || '';
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
    if (phase !== 'idle' || !editor || starting) return;
    const source = editor.getValue();
    starting = true;
    try {
      let compiled = null;
      try {
        if (!L.vb || !L.vb.compileIde) throw new Error(L.t('The VB runtime is not running'));
        compiled = await L.vb.compileIde(source);
      } catch (e) {
        appendOutput(e && e.message ? e.message : String(e));
        return;
      }
      if (phase !== 'idle' || !editor || editor.getValue() !== source) return;
      if (!compiled.valid) {
        applyDiagnostics(compiled.diagnostics);
        if (showErrors) showErrors();
        await reportDiagnostics(compiled.diagnostics);
        return;
      }
    } finally {
      starting = false;
    }
    setPhase('running');
    stopRequested = false;
    clearCurrent();
    clearRuntimeLists();
    if (output) output.value = '';
    outputText = '';
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
      if (token === generation && !(e && e.quiet)) appendOutput(e && e.message ? e.message : String(e));
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
      kind: 'melinaoffice.vb6.ide',
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

  function sessionData() {
    return {
      schema: 1,
      kind: 'melinaoffice.vb6.ide',
      code: editor ? editor.getValue() : code,
      breakpoints: breakpoints.map((bp) => ({ line: bp.line, condition: bp.condition || '', enabled: bp.enabled !== false })),
      output: output ? output.value : outputText,
    };
  }

  function loadSaved() {
    const data = L.store ? L.store.get(SESSION, null) : null;
    if (!data || data.kind !== 'melinaoffice.vb6.ide' || data.schema !== 1) {
      code = BLANK_CODE;
      breakpoints = [];
      outputText = '';
      return;
    }
    code = typeof data.code === 'string' ? data.code : BLANK_CODE;
    breakpoints = (Array.isArray(data.breakpoints) ? data.breakpoints : []).map(breakpointRecord).filter(Boolean);
    outputText = data.output == null ? '' : String(data.output);
  }

  function saveProject() {
    if (!L.store) return;
    const data = sessionData();
    code = data.code;
    outputText = data.output;
    L.store.set(SESSION, data);
  }

  function newBlankProject() {
    if (phase !== 'idle') return;
    code = BLANK_CODE;
    breakpoints = [];
    outputText = '';
    watches = [];
    expression = '';
    expressionResult = '';
    selectedLine = 0;
    selectedWatch = -1;
    projectName = 'Project.json';
    if (editor) editor.setValue(code);
    if (output) output.value = '';
    if (exprInput) exprInput.value = '';
    if (exprOut) exprOut.value = '';
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
    if (!data || data.kind !== 'melinaoffice.vb6.ide' || data.schema !== 1) {
      appendOutput(L.t('This file is not a Visual Basic project.'));
      return;
    }
    projectName = files[0].name || 'Project.json';
    applyProject(data);
  }

  async function saveProjectAs() {
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

  function loadEditorShare() {
    if (!L.store) return;
    const n = Number(L.store.get(EDITOR_SHARE, NaN));
    if (Number.isFinite(n) && n > 0) editorShare = n;
  }

  function saveEditorShare() {
    if (!L.store || !(editorShare > 0)) return;
    L.store.set(EDITOR_SHARE, editorShare);
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
        { label: '&New Blank Project', enabled: () => phase === 'idle', run: newBlankProject },
        { label: '&Open Project...', enabled: () => phase === 'idle', run: openProject },
        { label: '&Save Project', key: 'Ctrl+S', run: saveProject },
        { label: 'Save Project &As...', run: saveProjectAs },
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

    startBtn = L.h('button', { class: 'btn', type: 'button', text: 'Start Debug (F5)' });
    runBtn = L.h('button', { class: 'btn', type: 'button', text: 'Start Run' });
    contBtn = L.h('button', { class: 'btn', type: 'button', text: 'Continue' });
    intoBtn = L.h('button', { class: 'btn', type: 'button', text: 'Step Into' });
    overBtn = L.h('button', { class: 'btn', type: 'button', text: 'Step Over' });
    stopBtn = L.h('button', { class: 'btn', type: 'button', text: 'Stop Debugging' });
    startBtn.addEventListener('click', () => start(false));
    runBtn.addEventListener('click', () => start(true));
    contBtn.addEventListener('click', () => go('continue'));
    intoBtn.addEventListener('click', () => go('into'));
    overBtn.addEventListener('click', () => go('over'));
    stopBtn.addEventListener('click', stopDebugging);
    const bar = L.h('div', { class: 'vb-ide-bar' }, startBtn, runBtn, L.h('span', { class: 'vb-ide-sep' }), contBtn, intoBtn, overBtn, stopBtn);

    editorHost = L.h('div', { class: 'vb-ide-editor' });
    const area = L.h('textarea', { translate: 'no' });

    output = L.h('textarea', { class: 'vb-ide-output', readonly: true, translate: 'no' });
    output.value = outputText;
    const outputPage = L.h('div', { class: 'vb-ide-page' }, output);

    const errorGrid = grid([L.t('Line#'), L.t('Column#'), L.t('Error Number'), L.t('Source'), L.t('Error Message')]);
    errorBody = errorGrid.body;
    const errorsPage = L.h('div', { class: 'vb-ide-page', hidden: true }, errorGrid.wrap);

    exprInput = L.h('input', { class: 'vb-ide-expr', type: 'text', translate: 'no' });
    exprInput.value = expression;
    exprInput.addEventListener('keydown', (e) => {
      if (listEl && listTarget && listTarget.input === exprInput && handleListKey(e)) {
        e.preventDefault();
        e.stopPropagation();
        return;
      }
      if (isCompleteKey(e)) {
        if (phase !== 'paused') return;
        e.preventDefault();
        e.stopPropagation();
        openCompletion({ kind: 'input', input: exprInput });
        return;
      }
      if (e.key === 'Enter') { e.preventDefault(); evaluate(); }
    });
    exprInput.addEventListener('input', () => refreshInput(exprInput));
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
      if (listEl && listTarget && listTarget.input === watchInput && handleListKey(e)) {
        e.preventDefault();
        e.stopPropagation();
        return;
      }
      if (isCompleteKey(e)) {
        if (phase !== 'paused') return;
        e.preventDefault();
        e.stopPropagation();
        openCompletion({ kind: 'input', input: watchInput });
        return;
      }
      if (e.key === 'Enter') { e.preventDefault(); addWatch(); }
    });
    watchInput.addEventListener('input', () => refreshInput(watchInput));
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
      { name: 'errors', el: errorsPage },
      { name: 'expression', el: expressionPage },
      { name: 'breakpoints', el: breakPage },
      { name: 'watch', el: watchPage },
    ];
    const tabButtons = [
      ['Output', 'output'],
      ['Errors', 'errors'],
      ['Expression', 'expression'],
      ['Breakpoints', 'breakpoints'],
      ['Watch', 'watch'],
    ].map((pair) => {
      const button = L.h('button', { class: 'vb-ide-tab', type: 'button', text: pair[1] === 'errors' ? L.t('Errors') : pair[0], dataset: { page: pair[1] } });
      button.addEventListener('click', () => showPage(pages, tabButtons, pair[1]));
      return button;
    });
    showPage(pages, tabButtons, 'output');
    showErrors = () => showPage(pages, tabButtons, 'errors');
    paintErrors();

    stackList = L.h('ul', { class: 'vb-ide-list' });
    localList = L.h('ul', { class: 'vb-ide-list' });
    lowerPane = L.h('div', { class: 'vb-ide-lower' },
      L.h('div', { class: 'vb-ide-tabs' }, L.h('div', { class: 'vb-ide-tabbar' }, tabButtons), outputPage, errorsPage, expressionPage, breakPage, watchPage),
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
        up() { saveEditorShare(); },
      });
    });
    applyEditorShare();

    statusEl = L.h('div', { class: 'vb-ide-status', text: 'Ready' });
    paintStatus();

    return { root: L.h('div', { class: 'vb-ide' }, menu, bar, editorHost, split, lowerPane, statusEl), editorHost: editorHost, area: area };
  }

  function removeList() {
    if (listEl) listEl.remove();
    if (sigEl) sigEl.remove();
    listEl = null;
    sigEl = null;
    listItems = [];
    listTarget = null;
    listSpan = null;
  }

  function closeList() {
    listGen++;
    removeList();
  }

  function ensureEngine() {
    if (window.VB6Engine && window.VB6Engine.scanDeclarations) return Promise.resolve(window.VB6Engine);
    if (L.vb && L.vb.engine) return L.vb.engine();
    return Promise.reject(new Error(L.t('The VB runtime is not running')));
  }

  function completionFrame() {
    if (phase !== 'paused' || !run) return null;
    let frames = [];
    try { frames = run.vm.debugStack(); } catch (e) { frames = []; }
    const frame = selectedFrame == null ? frames[frames.length - 1] : frames.find((item) => item.index === selectedFrame);
    if (!frame) return null;
    let locals = [];
    try { locals = run.vm.debugLocals({ frameIndex: selectedFrame, includeFields: false }); }
    catch (e) { locals = []; }
    return {
      procedure: frame.procedure,
      module: frame.module,
      locals: locals.map((entry) => ({ name: entry.name, type: entry.type || '' })),
    };
  }

  function completionRequest(text, offset, frame) {
    const host = L.vb && L.vb.completionHost ? L.vb.completionHost() : { name: '', code: '', types: {}, files: false };
    const modules = [];
    if (host.name) modules.push({ name: host.name, code: host.code || '', host: true });
    return {
      engine: window.VB6Engine,
      code: text,
      offset: offset,
      moduleName: 'Module1',
      source: editor ? editor.getValue() : code,
      modules: modules,
      types: host.types || {},
      member: L.domIr && L.domIr.member,
      domTypes: L.domIr && L.domIr.types,
      purpose: L.vbDomRef && L.vbDomRef.purpose,
      intrinsics: window.VB6Engine.intrinsicCatalog(),
      files: host.files === true,
      frame: frame,
    };
  }

  function clearLink() {
    if (linkMark) {
      try { linkMark.clear(); } catch (e) { /* the editor already dropped it */ }
      linkMark = null;
    }
  }

  function nameChars(ch) {
    return /[A-Za-z_\u0080-\uffff0-9$%&!#@\[\]]/.test(ch || '');
  }

  function identifierRange(text, index) {
    let start = index;
    let end = index;
    if (start < text.length && nameChars(text[start])) {
      while (start > 0 && nameChars(text[start - 1])) start--;
      end = start;
      while (end < text.length && nameChars(text[end])) end++;
    } else if (start > 0 && nameChars(text[start - 1])) {
      end = start;
      while (start > 0 && nameChars(text[start - 1])) start--;
    } else return null;
    if (!/[A-Za-z_\u0080-\uffff\[]/.test(text.slice(start, end))) return null;
    return { start: start, end: end };
  }

  function paintLink(x, y) {
    if (!editor) return;
    const pos = editor.coordsChar({ left: x, top: y });
    if (!pos || pos.outside) { clearLink(); return; }
    const text = editor.getValue();
    const range = identifierRange(text, editor.indexFromPos(pos));
    if (!range) { clearLink(); return; }
    const from = editor.posFromIndex(range.start);
    const to = editor.posFromIndex(range.end);
    const found = linkMark && linkMark.find();
    if (found && found.from.line === from.line && found.from.ch === from.ch && found.to.line === to.line && found.to.ch === to.ch) return;
    clearLink();
    linkMark = editor.markText(from, to, { className: 'vb-ide-link' });
  }

  function hideHover() {
    if (hoverTimer) { clearTimeout(hoverTimer); hoverTimer = null; }
    if (hoverTip) hoverTip.hidden = true;
  }

  function tipBox(className) {
    if (hoverTip && hoverTip.className !== className) {
      hoverTip.remove();
      hoverTip = null;
    }
    if (!hoverTip) {
      hoverTip = L.h('div', { class: className, translate: 'no' });
      document.body.appendChild(hoverTip);
    }
    return hoverTip;
  }

  function positionHover(x, y) {
    hoverTip.hidden = false;
    hoverTip.style.left = (x + 12) + 'px';
    hoverTip.style.top = (y + 16) + 'px';
  }

  function placeHover(text, x, y) {
    if (!text) { hideHover(); return; }
    const tip = tipBox('tooltip vb-ide-tip');
    tip.textContent = text;
    positionHover(x, y);
  }

  /* Purpose stays in the reference's own words. Signature, Returns and Type are chrome. */
  function appendPurpose(parent, item) {
    if (item.purpose) parent.appendChild(L.h('div', { class: 'vb-ide-complete-sig-purpose', translate: 'no', text: item.purpose }));
  }

  function appendProcedureDetail(parent, item) {
    if (item.params) {
      parent.appendChild(L.h('div', { class: 'vb-ide-complete-sig-label', text: 'Signature' }));
      parent.appendChild(L.h('div', { class: 'vb-ide-complete-sig-params', translate: 'no', text: '(' + item.params + ')' }));
    }
    if (item.type && item.type !== 'Variant') {
      parent.appendChild(L.h('div', { class: 'vb-ide-complete-sig-label', text: 'Returns' }));
      parent.appendChild(L.h('div', { class: 'vb-ide-complete-sig-params', translate: 'no', text: item.type }));
    }
  }

  function appendWhy(parent, item) {
    let text = '';
    if (item.why === 'path') text = 'Not available on this platform because the open file has no path.';
    else if (item.why === 'app') text = 'Not available on this platform because this application does not implement it.';
    else if (item.why === 'platform') text = 'Not available on this platform because this session has no file host.';
    if (text) parent.appendChild(L.h('div', { class: 'vb-ide-complete-sig-why', text: text }));
  }

  function appendType(parent, item) {
    if (item.kind === 'procedure' || item.kind === 'keyword' || item.kind === 'module' || item.kind === 'type') return;
    if (!item.type || item.type === 'Variant') return;
    parent.appendChild(L.h('div', { class: 'vb-ide-complete-sig-label', text: 'Type' }));
    parent.appendChild(L.h('div', { class: 'vb-ide-complete-sig-params', translate: 'no', text: item.type }));
  }

  function fillSig(parent, item) {
    const head = L.h('div', { class: 'vb-ide-complete-sig-name' });
    const icon = L.icons && L.icons.get(COMPLETE_ICONS[item.kind] || 'field', 16);
    if (icon) head.appendChild(L.h('span', { class: 'vb-ide-complete-ico', html: icon }));
    head.appendChild(L.h('span', { translate: 'no', text: item.name }));
    parent.appendChild(head);
    appendPurpose(parent, item);
    if (item.kind === 'procedure') appendProcedureDetail(parent, item);
    else appendType(parent, item);
    appendWhy(parent, item);
  }

  function showSymbolTip(item, x, y) {
    if (!item || !item.name || item.kind === 'keyword' || item.kind === 'module') { hideHover(); return; }
    const tip = tipBox('vb-ide-complete-sig vb-ide-hover');
    tip.replaceChildren();
    fillSig(tip, item);
    positionHover(x, y);
  }

  function diagnosticAt(line, ch) {
    let best = null;
    for (let i = 0; i < diagnostics.length; i++) {
      const d = diagnostics[i];
      if (String(d.source || '').toLowerCase() !== MODULE) continue;
      if ((Number(d.line) || 1) !== line + 1) continue;
      const start = Math.max(0, (Number(d.column) || 1) - 1);
      if (ch < start) continue;
      if (!best || start >= best.start) best = { start: start, message: d.message || '' };
    }
    return best ? best.message : '';
  }

  async function inspectAt(index) {
    await ensureEngine();
    if (!editor || !L.vbInspect) return null;
    return L.vbInspect(completionRequest(editor.getValue(), index, completionFrame()));
  }

  function bindHover() {
    const wrap = editor.getWrapperElement();
    wrap.addEventListener('mousemove', (e) => {
      hoverPoint = { x: e.clientX, y: e.clientY };
      if ((e.ctrlKey || e.metaKey) && !e.altKey && !e.shiftKey && !e.buttons) paintLink(e.clientX, e.clientY);
      else clearLink();
      if (!editor || e.buttons) { hoverGen++; hideHover(); return; }
      const pos = editor.coordsChar({ left: e.clientX, top: e.clientY });
      if (pos.outside) { hoverGen++; hideHover(); return; }
      const key = pos.line + ':' + pos.ch;
      if (key === hoverKey && hoverTimer) return;
      hoverKey = key;
      if (hoverTimer) clearTimeout(hoverTimer);
      const gen = ++hoverGen;
      const point = { x: e.clientX, y: e.clientY, line: pos.line, ch: pos.ch };
      hoverTimer = setTimeout(async () => {
        hoverTimer = null;
        if (gen !== hoverGen || !editor) return;
        const message = diagnosticAt(point.line, point.ch);
        if (message) {         placeHover(message, point.x, point.y); return; }
        let item = null;
        try { item = await inspectAt(editor.indexFromPos({ line: point.line, ch: point.ch })); }
        catch (err) { item = null; }
        if (gen !== hoverGen || !editor) return;
        showSymbolTip(item, point.x, point.y);
      }, 300);
    });
    wrap.addEventListener('mouseleave', () => { hoverGen++; hoverKey = ''; hoverPoint = null; clearLink(); hideHover(); });
    editor.on('scroll', () => { hoverGen++; hoverKey = ''; clearLink(); hideHover(); });
    editor.on('mousedown', (cm, e) => {
      if (!(e.ctrlKey || e.metaKey) || e.altKey || e.shiftKey || e.button !== 0) return;
      const pos = cm.coordsChar({ left: e.clientX, top: e.clientY });
      if (!pos || pos.outside) return;
      const index = cm.indexFromPos(pos);
      const text = cm.getValue();
      const at = text[index] || '';
      const before = index > 0 ? text[index - 1] : '';
      if (!/[A-Za-z_\u0080-\uffff0-9$%&!#@\]]/.test(at) && !/[A-Za-z_\u0080-\uffff0-9$%&!#@\]]/.test(before)) return;
      e.preventDefault();
      hoverGen++;
      clearLink();
      hideHover();
      followName(index);
    });
    const onKey = (e) => {
      if (!editor) return;
      if (e.key !== 'Control' && e.key !== 'Meta') {
        if (!e.ctrlKey && !e.metaKey) clearLink();
        return;
      }
      if (e.type === 'keyup' || e.altKey || e.shiftKey || !hoverPoint) clearLink();
      else paintLink(hoverPoint.x, hoverPoint.y);
    };
    window.addEventListener('keydown', onKey, true);
    window.addEventListener('keyup', onKey, true);
    window.addEventListener('blur', clearLink);
    linkOff = () => {
      window.removeEventListener('keydown', onKey, true);
      window.removeEventListener('keyup', onKey, true);
      window.removeEventListener('blur', clearLink);
      clearLink();
      linkOff = null;
    };
  }

  async function followName(index) {
    let item = null;
    try { item = await inspectAt(index); }
    catch (e) { return; }
    if (!item || !item.name || !editor) return;
    if (item.local && item.line) jumpToLine(item.line);
  }

  function placeList(anchor) {
    listEl.style.left = '0px';
    listEl.style.top = '0px';
    const width = listEl.offsetWidth || 360;
    const height = listEl.offsetHeight || 180;
    const left = Math.max(2, Math.min(window.innerWidth - width - 4, anchor.left));
    let top = anchor.bottom + 2;
    if (top + height > window.innerHeight) top = Math.max(2, anchor.top - height - 2);
    listEl.style.left = left + 'px';
    listEl.style.top = top + 'px';
  }

  function placeSig() {
    if (!sigEl || !listEl) return;
    const rect = listEl.getBoundingClientRect();
    const gap = 2;
    sigEl.style.width = rect.width + 'px';
    sigEl.style.top = rect.top + 'px';
    let left = rect.right + gap;
    if (left + rect.width > window.innerWidth) left = rect.left - gap - rect.width;
    sigEl.style.left = left + 'px';
  }

  /* Beside the list: every name, with its own icon. A note when the reference has one, Type for a variable or member, and why a gray member cannot run. */
  function paintSig() {
    if (sigEl) { sigEl.remove(); sigEl = null; }
    const item = listItems[listIndex];
    if (!item) return;
    sigEl = L.h('div', { class: 'vb-ide-complete-sig' });
    fillSig(sigEl, item);
    document.body.appendChild(sigEl);
  }

  const COMPLETE_ICONS = {
    procedure: 'fx',
    member: 'properties',
    keyword: 'currency',
    variable: 'field',
    constant: 'symbol',
    module: 'onePage',
    type: 'object',
  };

  function hintsOf(item) {
    const hints = [];
    if (item.kind !== 'procedure' && item.params) hints.push('(' + item.params + ')');
    if (item.kind !== 'type' && item.type && item.type !== 'Variant') hints.push('(As ' + item.type + ')');
    if (item.kind === 'keyword') hints.push('(' + L.t('Keyword') + ')');
    return hints.join(' ');
  }

  /* A name or a member dot. A quote, apostrophe or date mark does not ask for a list; the scan then drops anything still inside a string, comment or date. */
  function headWantsList(head) {
    return /[A-Za-z_\u0080-\uFFFF$%&!#@]$/.test(head) || /\.\s*$/.test(head);
  }

  function paintList(result, anchor) {
    removeList();
    if (!result.items.length) return;
    listItems = result.items;
    listSpan = { start: result.start, end: result.end };
    listIndex = 0;
    listEl = L.h('div', { class: 'vb-ide-complete', role: 'listbox' });
    listEl.setAttribute('aria-label', L.t('Completions'));
    for (let i = 0; i < listItems.length; i++) {
      const item = listItems[i];
      const row = L.h('div', {
        class: 'vb-ide-complete-item' + (i === 0 ? ' on' : '') + (item.unavailable ? ' dis' : ''),
        role: 'option',
        translate: 'no',
      });
      row.dataset.name = item.name;
      const icon = L.icons && L.icons.get(COMPLETE_ICONS[item.kind] || 'field', 16);
      if (icon) row.appendChild(L.h('span', { class: 'vb-ide-complete-ico', html: icon }));
      row.appendChild(L.h('span', { class: 'vb-ide-complete-name', translate: 'no', text: item.name }));
      const detail = hintsOf(item);
      if (detail) row.appendChild(L.h('span', { class: 'vb-ide-complete-detail', translate: 'no', text: detail }));
      row.addEventListener('pointerdown', (e) => {
        e.preventDefault();
        e.stopPropagation();
        listIndex = i;
        acceptList();
      });
      listEl.appendChild(row);
    }
    document.body.appendChild(listEl);
    placeList(anchor);
    paintSig();
    placeSig();
    const selected = listEl.querySelector('.on');
    if (selected) selected.scrollIntoView({ block: 'nearest' });
  }

  function moveList(delta) {
    if (!listEl || !listItems.length) return;
    const count = listItems.length;
    listIndex = delta === 'home' ? 0 : delta === 'end' ? count - 1 : Math.max(0, Math.min(count - 1, listIndex + delta));
    const rows = listEl.querySelectorAll('.vb-ide-complete-item');
    for (let i = 0; i < rows.length; i++) rows[i].classList.toggle('on', i === listIndex);
    paintSig();
    placeSig();
    if (rows[listIndex]) rows[listIndex].scrollIntoView({ block: 'nearest' });
  }

  function acceptList() {
    const item = listItems[listIndex];
    const target = listTarget;
    const span = listSpan;
    if (!item || !target || !span) { closeList(); return; }
    closeList();
    if (target.kind === 'editor' && editor) {
      editor.replaceRange(item.name, editor.posFromIndex(span.start), editor.posFromIndex(span.end), '+complete');
      editor.focus();
      return;
    }
    if (target.input) {
      const input = target.input;
      const value = input.value;
      accepting = true;
      input.value = value.slice(0, span.start) + item.name + value.slice(span.end);
      const caret = span.start + item.name.length;
      input.setSelectionRange(caret, caret);
      accepting = false;
      input.focus();
    }
  }

  function handleListKey(e) {
    if (!listEl) return false;
    if (e.key === 'ArrowUp' || e.key === 'ArrowDown' || e.key === 'PageUp' || e.key === 'PageDown' || e.key === 'Home' || e.key === 'End') {
      moveList({ ArrowUp: -1, ArrowDown: 1, PageUp: -8, PageDown: 8, Home: 'home', End: 'end' }[e.key]);
      return true;
    }
    if ((e.key === 'Enter' || e.key === 'Tab') && !e.shiftKey && !e.ctrlKey && !e.metaKey && !e.altKey) {
      acceptList();
      return true;
    }
    if (e.key === 'Escape') { closeList(); return true; }
    if (e.key === 'ArrowLeft' || e.key === 'ArrowRight') { closeList(); return false; }
    return false;
  }

  function isCompleteKey(e) {
    return (e.ctrlKey || e.metaKey) && !e.altKey && !e.shiftKey && (e.key === ' ' || e.code === 'Space');
  }

  async function openCompletion(target) {
    const gen = ++listGen;
    let text = '';
    let offset = 0;
    let frame = null;
    let anchor = null;
    if (target.kind === 'editor') {
      if (!editor) return;
      text = editor.getValue();
      offset = editor.indexFromPos(editor.getCursor());
      const coords = editor.cursorCoords(null, 'window');
      anchor = { left: coords.left, top: coords.top, bottom: coords.bottom };
    } else {
      frame = completionFrame();
      if (!frame || !target.input) { if (gen === listGen) removeList(); return; }
      text = target.input.value;
      offset = target.input.selectionStart || 0;
      const rect = target.input.getBoundingClientRect();
      anchor = { left: rect.left, top: rect.top, bottom: rect.bottom };
    }
    try { await ensureEngine(); }
    catch (e) { if (gen === listGen) removeList(); return; }
    if (gen !== listGen || !L.vbComplete) return;
    let result = { items: [] };
    try { result = L.vbComplete(completionRequest(text, offset, frame)); }
    catch (e) { result = { items: [] }; }
    if (gen !== listGen) return;
    if (!result.items.length) { removeList(); return; }
    paintList(result, anchor);
    listTarget = target;
  }

  function refreshInput(input) {
    if (accepting) return;
    if (phase !== 'paused') {
      if (listTarget && listTarget.input === input) closeList();
      return;
    }
    const head = input.value.slice(0, input.selectionStart || 0);
    if (headWantsList(head)) openCompletion({ kind: 'input', input: input });
    else if (listEl && listTarget && listTarget.input === input) closeList();
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
    loadSaved();
    loadEditorShare();
    const built = buildBody();
    L.ui.dialog({
      title: 'Melina Office Visual Basic',
      body: built.root,
      width: 820,
      buttons: [],
      modeless: true,
      resizeable: true,
      windowName: 'vb_ide',
      restoreWindowSizeAndPosition: 1,
      minWidth: 520,
      minHeight: 360,
      onResize: () => { if (editor) editor.refresh(); },
      onOpen: (el) => {
        el.classList.add('vb-ide-dlg');
        el.style.height = Math.min(560, window.innerHeight - 24) + 'px';
        el.addEventListener('keydown', (e) => {
          if ((e.ctrlKey || e.metaKey) && !e.altKey && !e.shiftKey && (e.key === 's' || e.key === 'S')) {
            e.preventDefault();
            e.stopPropagation();
            saveProject();
            return;
          }
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
          indentUnit: 4,
          tabSize: 4,
          indentWithTabs: false,
          gutters: [{ className: 'vb-ide-bp-gutter', style: 'width: 14px' }, 'CodeMirror-linenumbers'],
        });
        editor.on('gutterClick', (cm, n, gutter, event) => {
          if (event && event.button === 2) return;
          if (gutter !== 'vb-ide-bp-gutter' && gutter !== 'CodeMirror-linenumbers') return;
          const line = n + 1;
          /* the second press of a double-click: keep the breakpoint and open its condition */
          if (event && event.detail > 1) {
            if (event.detail === 2) {
              if (!bpAt(line)) toggleBreakpoint(line);
              if (bpAt(line)) openConditionPanel(line, event.clientX, event.clientY);
            }
            return;
          }
          toggleBreakpoint(line);
        });
        editor.on('gutterContextMenu', (cm, n, gutter, event) => {
          if (gutter !== 'vb-ide-bp-gutter' || !bpAt(n + 1)) return;
          if (event) { event.preventDefault(); event.stopPropagation(); }
          openConditionPanel(n + 1, event ? event.clientX : 0, event ? event.clientY : 0);
        });
        editor.on('keydown', (cm, e) => {
          hoverGen++;
          hideHover();
          if (listEl && listTarget && listTarget.kind === 'editor' && handleListKey(e)) {
            e.preventDefault();
            e.stopPropagation();
            return;
          }
          if (!isCompleteKey(e)) return;
          e.preventDefault();
          e.stopPropagation();
          openCompletion({ kind: 'editor' });
        });
        bindHover();
        let paintQueued = false;
        editor.on('change', (cm, change) => {
          hoverGen++;
          hideHover();
          code = editor.getValue();
          const origin = change && change.origin;
          if (origin !== '+complete' && origin !== 'setValue') {
            const cur = cm.getCursor();
            const head = cm.getLine(cur.line).slice(0, cur.ch);
            if (headWantsList(head)) openCompletion({ kind: 'editor' });
            else if (listEl && listTarget && listTarget.kind === 'editor') closeList();
          }
          if (paintQueued) return;
          paintQueued = true;
          setTimeout(() => {
            paintQueued = false;
            if (!editor) return;
            paintDots();
            paintBpList();
          }, 0);
          scheduleCheck();
        });
        editor.setValue(code);
        scheduleCheck();
        setPhase(phase);
        paintDots();
        paintBpList();
        paintWatches();
        setTimeout(() => { if (editor) editor.refresh(); }, 0);
      },
      onClose: () => {
        saveEditorShare();
        closeList();
        closeConditionPanel();
        if (checkTimer) { clearTimeout(checkTimer); checkTimer = null; }
        checkGen++;
        clearSquiggles();
        hoverGen++;
        hideHover();
        if (linkOff) linkOff();
        else clearLink();
        if (hoverTip) { hoverTip.remove(); hoverTip = null; }
        if (editor) code = editor.getValue();
        if (exprInput) expression = exprInput.value;
        if (exprOut) expressionResult = exprOut.value;
        editor = null;
        openDlg = null;
        output = null;
        statusEl = null;
        editorHost = null;
        lowerPane = null;
        bpBody = null;
        watchBody = null;
        errorBody = null;
        showErrors = null;
        stopDebugging();
      },
    });
  }

  if (typeof document !== 'undefined') {
    document.addEventListener('pointerdown', (e) => {
      if (!listEl || listEl.contains(e.target) || (sigEl && sigEl.contains(e.target))) return;
      closeList();
    });
  }

  L.vbIde = { open: open, reportDiagnostics: reportDiagnostics };
})();
