/* Visual Basic IDE. One standard module, debugged on the app host's debug VM.
   Opened from View. The window is the ProjectEditor layout: editor, output,
   expression, breakpoints, call stack, object values. */
(function () {
  'use strict';
  const L = window.L;
  const MODULE = 'module1';

  let code = '';
  let breakpoints = [];
  let expression = '';
  let expressionResult = '';
  let projectName = 'Project.json';
  let editor = null;
  let openDlg = null;
  let output = null;
  let exprInput = null;
  let exprOut = null;
  let lineInput = null;
  let bpList = null;
  let stackList = null;
  let localList = null;
  let startBtn = null;
  let runBtn = null;
  let contBtn = null;
  let intoBtn = null;
  let overBtn = null;
  let stopBtn = null;
  let phase = 'idle';
  let generation = 0;
  let stopRequested = false;
  let run = null;
  let selectedLine = 0;
  let selectedFrame = null;
  let currentMark = null;

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
    if (runBtn) runBtn.disabled = !idle;
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

  function lineLabel(line, text) {
    return L.t('Line {0} - {1}', line, text == null ? '' : text);
  }

  function stackLabel(line, procedure) {
    return L.t('Line {0}| {1}', line, procedure == null ? '' : procedure);
  }

  function objectValue(entry) {
    const type = entry.type || '';
    if (entry.expandable || type === 'Object' || type === 'Collection' || type === 'Dictionary') return L.t('[Object]');
    return entry.value == null ? '' : String(entry.value);
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

  function paintDots() {
    if (!editor) return;
    editor.eachLine((handle) => editor.setGutterMarker(handle, 'vb-ide-bp-gutter', null));
    for (let i = 0; i < breakpoints.length; i++) {
      const line = breakpoints[i];
      if (line < 1 || line > editor.lineCount()) continue;
      const dot = document.createElement('span');
      dot.className = 'vb-ide-dot';
      editor.setGutterMarker(line - 1, 'vb-ide-bp-gutter', dot);
    }
  }

  function paintBpList() {
    if (!bpList) return;
    bpList.replaceChildren();
    for (let i = 0; i < breakpoints.length; i++) {
      const line = breakpoints[i];
      const li = L.h('li', {
        class: line === selectedLine ? 'on' : '',
        translate: 'no',
        text: lineLabel(line, editor ? sourceLine(MODULE, line) : ''),
      });
      li.addEventListener('click', () => {
        selectedLine = line;
        paintBpList();
        jumpToLine(line);
      });
      bpList.appendChild(li);
    }
  }

  function rememberBreakpoint(line) {
    if (breakpoints.indexOf(line) >= 0) return;
    breakpoints.push(line);
    breakpoints.sort((a, b) => a - b);
  }

  function toggleBreakpoint(line) {
    if (!Number.isInteger(line) || line < 1) return;
    const at = breakpoints.indexOf(line);
    if (at >= 0) {
      breakpoints.splice(at, 1);
      if (selectedLine === line) selectedLine = 0;
      if (run && !run.ignore) {
        try { run.vm.removeBreakpoint(run.moduleName, line); } catch (e) { /* already gone */ }
      }
    } else {
      rememberBreakpoint(line);
      selectedLine = line;
      if (run && !run.ignore) {
        try { run.vm.setBreakpoint(run.moduleName, line); }
        catch (e) { appendOutput(L.t('Line {0}: {1}', line, e && e.message ? e.message : e)); }
      }
    }
    paintDots();
    paintBpList();
  }

  function installBreakpoints() {
    if (!run) return;
    for (let i = 0; i < breakpoints.length; i++) {
      const line = breakpoints[i];
      try { run.vm.setBreakpoint(run.moduleName, line); }
      catch (e) { appendOutput(L.t('Line {0}: {1}', line, e && e.message ? e.message : e)); }
    }
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
    if (stackList) stackList.replaceChildren();
    if (localList) localList.replaceChildren();
  }

  function onPause(e) {
    setPhase('paused');
    const ins = e.instruction || {};
    selectedFrame = null;
    if (sameModule(ins.source)) markCurrent(ins.line);
    else clearCurrent();
    paintStack();
    paintLocals(null);
    if (e.reason === 'error' && e.error && e.error.message) appendOutput(e.error.message);
  }

  function go(mode) {
    if (phase !== 'paused' || !run) return;
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
      begun.vm.on('output', (ev) => { if (token === generation) appendOutput(ev.text); });
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
      const shown = run.host.engine && run.host.engine.describe ? run.host.engine.describe(value) : String(value);
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
      breakpoints: breakpoints.slice(),
      expression: exprInput ? exprInput.value : expression,
      expressionResult: exprOut ? exprOut.value : expressionResult,
    };
  }

  function applyProject(data) {
    code = String(data.code || '');
    breakpoints = (Array.isArray(data.breakpoints) ? data.breakpoints : []).map(Number).filter((n) => Number.isInteger(n) && n > 0);
    expression = String(data.expression || '');
    expressionResult = String(data.expressionResult || '');
    selectedLine = 0;
    if (editor) editor.setValue(code);
    if (exprInput) exprInput.value = expression;
    if (exprOut) exprOut.value = expressionResult;
    paintDots();
    paintBpList();
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
    if (openDlg && openDlg.isConnected) {
      const closer = openDlg.querySelector('.dlg-x');
      if (closer) closer.click();
    }
  }

  function addBreakpoint() {
    const line = parseInt(lineInput && lineInput.value, 10);
    if (!Number.isInteger(line) || line < 1) return;
    if (breakpoints.indexOf(line) >= 0) {
      selectedLine = line;
      paintBpList();
      jumpToLine(line);
      return;
    }
    toggleBreakpoint(line);
    jumpToLine(line);
  }

  function removeBreakpoint() {
    if (!selectedLine) return;
    const line = selectedLine;
    if (breakpoints.indexOf(line) < 0) return;
    toggleBreakpoint(line);
  }

  function removeAllBreakpoints() {
    const lines = breakpoints.slice();
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

    const editorHost = L.h('div', { class: 'vb-ide-editor' });
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
    const addBtn = L.h('button', { class: 'btn', type: 'button', text: 'Add' });
    addBtn.addEventListener('click', addBreakpoint);
    bpList = L.h('ul', { class: 'vb-ide-list' });
    const removeBtn = L.h('button', { class: 'btn', type: 'button', text: 'Remove Breakpoint' });
    const removeAllBtn = L.h('button', { class: 'btn', type: 'button', text: 'Remove All Breakpoints' });
    removeBtn.addEventListener('click', removeBreakpoint);
    removeAllBtn.addEventListener('click', removeAllBreakpoints);
    const breakPage = L.h('div', { class: 'vb-ide-page', hidden: true },
      L.h('div', { class: 'vb-ide-row' },
        L.h('label', { text: 'Line Number' }),
        lineInput,
        addBtn),
      bpList,
      L.h('div', { class: 'vb-ide-row' }, removeBtn, removeAllBtn));

    const pages = [
      { name: 'output', el: outputPage },
      { name: 'expression', el: expressionPage },
      { name: 'breakpoints', el: breakPage },
    ];
    const tabButtons = ['Output', 'Expression', 'Breakpoints'].map((label) => {
      const button = L.h('button', { class: 'vb-ide-tab', type: 'button', text: label, dataset: { page: label.toLowerCase() } });
      button.addEventListener('click', () => showPage(pages, tabButtons, label.toLowerCase()));
      return button;
    });
    showPage(pages, tabButtons, 'output');

    stackList = L.h('ul', { class: 'vb-ide-list' });
    localList = L.h('ul', { class: 'vb-ide-list' });
    const lower = L.h('div', { class: 'vb-ide-lower' },
      L.h('div', { class: 'vb-ide-tabs' }, L.h('div', { class: 'vb-ide-tabbar' }, tabButtons), outputPage, expressionPage, breakPage),
      L.h('div', { class: 'vb-ide-col' }, L.h('div', { class: 'vb-ide-cap', text: 'Call Stack' }), stackList),
      L.h('div', { class: 'vb-ide-col' }, L.h('div', { class: 'vb-ide-cap', text: 'Object Values' }), localList));

    return { root: L.h('div', { class: 'vb-ide' }, menu, bar, editorHost, lower), editorHost: editorHost, area: area };
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
      title: 'Visual Basic IDE',
      body: built.root,
      width: 820,
      buttons: [],
      modeless: true,
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
        editor.on('gutterClick', (cm, n, gutter) => {
          if (gutter !== 'vb-ide-bp-gutter' && gutter !== 'CodeMirror-linenumbers') return;
          toggleBreakpoint(n + 1);
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
        setTimeout(() => { if (editor) editor.refresh(); }, 0);
      },
      onClose: () => {
        if (editor) code = editor.getValue();
        if (exprInput) expression = exprInput.value;
        if (exprOut) expressionResult = exprOut.value;
        editor = null;
        openDlg = null;
        output = null;
        stopDebugging();
      },
    });
  }

  L.vbIde = { open: open };
})();
