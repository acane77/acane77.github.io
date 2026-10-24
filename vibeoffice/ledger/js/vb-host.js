/* Ledger VB host. Immediate VM for evalVB6, a second VM for runVB6,
   a workbook-macro VM that stays loaded so macros can call each other,
   and a debug VM for the Visual Basic IDE (debugBegin / debugEnd).
   vbaProject.bin is not run. Forms are not started.
   Build the engine first: node tools/vb6/build-engine.mjs */
(function () {
  'use strict';
  const L = window.L;
  const SLOTS = ['createForm', 'beginModal', 'msgBox', 'inputBox', 'openFile', 'saveFile', 'loadProject'];
  const CALL_PARAMS = [];
  for (let i = 1; i <= 16; i++) CALL_PARAMS.push({ name: 'Arg' + i, optional: true });

  function engineReady() {
    const engine = window.VB6Engine;
    return engine && engine.VirtualMachine ? engine : null;
  }

  function loadEngine() {
    const ready = engineReady();
    if (ready) return Promise.resolve(ready);
    if (typeof document === 'undefined' || !document.body) {
      return Promise.reject(new Error('VB6 engine is not loaded. Run node tools/vb6/build-engine.mjs'));
    }
    return new Promise(function (resolve, reject) {
      const script = document.createElement('script');
      script.src = 'vendor/vb6/vb6-engine.js';
      script.onload = function () {
        const engine = engineReady();
        if (engine) resolve(engine);
        else reject(new Error('vendor/vb6/vb6-engine.js did not define VB6Engine'));
      };
      script.onerror = function () {
        reject(new Error('VB6 engine is not built. Run node tools/vb6/build-engine.mjs'));
      };
      document.body.appendChild(script);
    });
  }

  function callParams(desc) {
    if (desc && Array.isArray(desc.params)) {
      return desc.params.map(function (p) { return { name: p.name, optional: p.optional !== false }; });
    }
    return CALL_PARAMS;
  }

  function memberMeta(name, desc) {
    const params = callParams(desc);
    if (desc.kind === 'gap') {
      const as = desc.as || 'any';
      if (as === 'call') return { name: name, modes: [1], params: params };
      if (as === 'get') return { name: name, modes: [2], params: [] };
      if (as === 'rw') return { name: name, modes: [2, 4, 8], params: [] };
      return { name: name, modes: [1, 2, 4, 8], params: params };
    }
    if (desc.kind === 'call' || desc.kind === 'both') return { name: name, modes: [1, 2], params: params };
    if (desc.kind === 'rw') return { name: name, modes: [2, 4, 8], params: [] };
    return { name: name, modes: [2], params: [] };
  }

  function moduleCode() {
    const lines = [
      'Public Ledger As Object',
      '',
      'Public Property Get Application() As Object',
      '    Set Application = Ledger',
      'End Property',
      '',
      'Public Property Get ActiveWorkbook() As Object',
      '    Set ActiveWorkbook = Ledger.ActiveWorkbook',
      'End Property',
      '',
      'Public Property Get ThisWorkbook() As Object',
      '    Set ThisWorkbook = Ledger.ThisWorkbook',
      'End Property',
      '',
      'Public Property Get Workbooks() As Object',
      '    Set Workbooks = Ledger.Workbooks',
      'End Property',
      '',
      'Public Property Get ActiveSheet() As Object',
      '    Set ActiveSheet = Ledger.ActiveSheet',
      'End Property',
      '',
      'Public Property Get ActiveCell() As Object',
      '    Set ActiveCell = Ledger.ActiveCell',
      'End Property',
      '',
      'Public Property Get Selection() As Object',
      '    Set Selection = Ledger.Selection',
      'End Property',
      '',
      'Public Property Get WorksheetFunction() As Object',
      '    Set WorksheetFunction = Ledger.WorksheetFunction',
      'End Property',
      '',
    ];
    const names = Object.keys(L.dom).filter(function (key) {
      const value = L.dom[key];
      return /^(xl|mso)[A-Za-z][A-Za-z0-9]*$/.test(key) && typeof value === 'number' && Number.isInteger(value) && value >= -2147483648 && value <= 2147483647;
    }).sort();
    for (let i = 0; i < names.length; i++) lines.push('Public Const ' + names[i] + ' As Long = ' + L.dom[names[i]]);
    const gaps = gapNames();
    if (gaps.length) lines.push('');
    for (let i = 0; i < gaps.length; i++) lines.push('Public ' + gaps[i] + ' As Object');
    if (L.hostAudit && L.hostAudit.moduleLines) {
      lines.push('');
      const extra = L.hostAudit.moduleLines('Ledger');
      for (let i = 0; i < extra.length; i++) lines.push(extra[i]);
    }
    return lines.join('\n');
  }

  function gapNames() {
    const types = L.domIr && L.domIr.types;
    if (!types) return [];
    return Object.keys(types).filter(function (name) {
      return types[name].gap && /^[A-Za-z][A-Za-z0-9]*$/.test(name);
    }).sort();
  }

  function LedgerVbHost() {
    for (let i = 0; i < SLOTS.length; i++) this[SLOTS[i]] = null;
    this.output = [];
    this.errors = [];
    this.vm = null;
    this.ir = null;
    this.engine = null;
    this.auto = null;
    this.byHandle = new Map();
    this.byObject = new WeakMap();
    this.meta = new Map();
  }

  LedgerVbHost.prototype.metadata = function (type) {
    const hit = this.meta.get(type);
    if (hit) return hit;
    const spec = L.domIr.types[type];
    const names = ['Application', 'Parent', 'Creator'];
    if (spec) {
      for (const name in spec.members) if (names.indexOf(name) < 0) names.push(name);
    }
    const members = [];
    for (let i = 0; i < names.length; i++) {
      const desc = L.domIr.member(type, names[i]);
      if (desc) members.push(memberMeta(names[i], desc));
    }
    if (members.length > 1024) throw new Error(type + ' exceeds the Automation member limit');
    const meta = { members: members };
    if ((spec && spec.gap) || type === 'WorksheetFunction') meta.acceptAnyMember = true;
    if (type === 'Range' || (spec && spec.collection && members.some(function (m) { return m.name === 'Item'; }))) meta.defaultMember = 'Item';
    this.meta.set(type, meta);
    return meta;
  };

  LedgerVbHost.prototype.adapter = function (handle, type) {
    const host = this;
    return {
      metadata: host.metadata(type),
      invoke: function (name, mode, values) { return host.dispatch(handle, name, mode, values); },
      release: function () {},
    };
  };

  LedgerVbHost.prototype.remember = function (handle, object) {
    this.byHandle.set(handle, object);
    this.byObject.set(object, handle);
    return object;
  };

  LedgerVbHost.prototype.objectFor = function (handle, type) {
    const hit = this.byHandle.get(handle);
    if (hit) return hit;
    if (!this.auto) throw new Error('Automation session is not open');
    return this.remember(handle, this.auto.adopt(this.adapter(handle, type || 'Object')));
  };

  function isVbArray(value) {
    return !!(value && Array.isArray(value.bounds) && Array.isArray(value.data) && typeof value.get === 'function');
  }

  LedgerVbHost.prototype.fromArray = function (value) {
    if (!isVbArray(value)) return value;
    const bounds = value.bounds;
    if (bounds.length === 1) {
      const out = [];
      for (let i = bounds[0][0]; i <= bounds[0][1]; i++) out.push(this.fromArray(value.get(i)));
      return out;
    }
    if (bounds.length === 2) {
      const out = [];
      for (let r = bounds[0][0]; r <= bounds[0][1]; r++) {
        const row = [];
        for (let c = bounds[1][0]; c <= bounds[1][1]; c++) row.push(this.fromArray(value.get(r, c)));
        out.push(row);
      }
      return out;
    }
    return value;
  };

  LedgerVbHost.prototype.toArray = function (value) {
    const VBArray = this.engine && this.engine.VBArray;
    if (!VBArray || !Array.isArray(value)) return value;
    if (!value.length) return new VBArray([[1, 1], [1, 1]]);
    const rows = value.length;
    const cols = Array.isArray(value[0]) ? value[0].length : 0;
    if (!cols) {
      const a = new VBArray([[1, rows]]);
      for (let i = 0; i < rows; i++) a.set([i + 1], value[i] == null ? undefined : value[i]);
      return a;
    }
    const a = new VBArray([[1, rows], [1, cols]]);
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        const cell = value[r][c];
        a.set([r + 1, c + 1], cell == null ? undefined : cell);
      }
    }
    return a;
  };

  LedgerVbHost.prototype.toIr = function (value) {
    if (value && value.__missing === true) return L.domIr.Missing;
    if (isVbArray(value)) return this.fromArray(value);
    if (value && typeof value === 'object') {
      const handle = this.byObject.get(value);
      if (handle != null) return { handle: handle };
    }
    return value;
  };

  LedgerVbHost.prototype.dispatch = async function (handle, name, mode, values) {
    const args = [];
    for (let i = 0; i < values.length; i++) args.push(this.toIr(values[i]));
    let req;
    if (mode === 4 || mode === 8) {
      req = { op: 'set', handle: handle, member: name, value: args.pop() };
    } else if (mode === 1 && String(name).toLowerCase() === 'item' && args.length <= 1) {
      req = { op: 'item', handle: handle, member: 'Item', args: args };
    } else if (mode === 1 || (mode === 2 && args.length)) {
      req = { op: 'call', handle: handle, member: name, args: args };
    } else {
      req = { op: 'get', handle: handle, member: name };
    }
    const result = await Promise.resolve(this.ir.exec(req));
    if (!result || !result.ok) {
      const err = new this.engine.VBError((result && result.message) || 'IR call failed', 438);
      if (result && result.code) err.code = result.code;
      throw err;
    }
    if (result.kind === 'handle') return { value: this.objectFor(result.handle, result.type || 'Object') };
    if (result.kind === 'nothing') return { value: this.engine.NOTHING };
    if (result.kind === 'array') return { value: this.toArray(result.value) };
    return { value: result.kind === 'empty' ? undefined : result.value };
  };

  function projectOf(id, modules) {
    return { schema: 1, id: id, name: 'Ledger', startup: '', modules: modules, settings: {}, references: [], assets: {} };
  }

  function domModule() {
    return { id: 'ledger-dom', name: 'LedgerDOM', kind: 'module', code: moduleCode() };
  }

  function procedure(machine, name) {
    const key = String(name).toLowerCase();
    let found = null;
    for (const instance of machine.instances.values()) {
      const mod = instance.module;
      if (!mod || !mod.procedures) continue;
      const proc = mod.procedures.get(key);
      if (!proc || proc.kind !== 'sub' || proc.scope === 'private') continue;
      const modName = String(mod.name || '').toLowerCase();
      if (modName === key) return { instance: instance, proc: proc };
      if (!found && modName !== 'ledgerdom') found = { instance: instance, proc: proc };
    }
    return found;
  }

  /* The engine treats a standard module's name as an object, and that object hides a procedure of the
     same name. The module is therefore not named after the macro, so another macro can call the public Sub. */
  function macroModules(macros) {
    const modules = [domModule()];
    const used = new Set(['ledgerdom']);
    for (const macro of macros) {
      const called = String((macro && macro.name) || '').toLowerCase();
      if (called) used.add(called);
    }
    for (let i = 0; i < macros.length; i++) {
      let name = 'VibeMacro' + (i + 1);
      while (used.has(name.toLowerCase())) name += 'X';
      used.add(name.toLowerCase());
      modules.push({ id: 'm' + (i + 1), name: name, kind: 'module', code: String((macros[i] && macros[i].source) || '') });
    }
    return modules;
  }

  LedgerVbHost.prototype.machineHost = function (registry) {
    const host = this;
    const vmHost = {
      automation: registry,
      /* onPrint runs in this turn, so the Output box can show the line before MsgBox waits. */
      print: function (text) {
        host.output.push(text == null ? '' : String(text));
        if (typeof host.onPrint === 'function') host.onPrint();
      },
      error: function (err) { host.errors.push(err); },
    };
    for (let i = 0; i < SLOTS.length; i++) {
      (function (name) {
        Object.defineProperty(vmHost, name, { enumerable: true, get: function () { return host[name]; } });
      })(SLOTS[i]);
    }
    return vmHost;
  };

  LedgerVbHost.prototype.registry = function () {
    const host = this;
    const registry = new this.engine.AutomationRegistry();
    registry.register('Ledger.Application', function (session) {
      host.auto = session;
      return host.adapter(host.ir.root.handle, 'Application');
    });
    return registry;
  };

  LedgerVbHost.prototype.bindLedger = async function (machine) {
    const app = await machine.automation.create('Ledger.Application');
    this.remember(this.ir.root.handle, app);
    const inst = machine.instances.get('ledgerdom');
    const cell = inst && inst.fields.get('ledger');
    if (!cell) throw new Error('The injected module has no Ledger object');
    cell.set(app);
    const gaps = gapNames();
    for (let i = 0; i < gaps.length; i++) {
      const named = this.ir.named(gaps[i]);
      if (!named.ok) continue;
      const field = inst.fields.get(gaps[i].toLowerCase());
      if (field) field.set(this.objectFor(named.handle, named.type));
    }
  };

  LedgerVbHost.prototype.start = async function () {
    if (this.vm) return this;
    const engine = await loadEngine();
    this.engine = engine;
    this.ir = L.domIr.open();
    const project = projectOf('ledger-vb', [domModule()]);
    try {
      this.vm = new engine.VirtualMachine(project, this.machineHost(this.registry()), { debuggerEnabled: true });
      await this.vm.prepareImmediateContext();
      await this.bindLedger(this.vm);
    } catch (error) {
      this.close();
      throw error;
    }
    return this;
  };

  LedgerVbHost.prototype.run = async function (source, entry, args) {
    if (this.script) throw new Error('A VB script is running');
    if (this.vm && this.vm.debugEvaluation) throw new Error('A VB statement is running');
    if (!this.vm || this.vm.state === 'stopped' || !this.ir || !this.engine) throw new Error('The VB runtime is not running');
    if (args != null && !Array.isArray(args)) throw new Error('Procedure arguments must be an array');
    const saved = { byHandle: this.byHandle, byObject: this.byObject, auto: this.auto };
    const named = entry != null && entry !== '';
    let machine = null;
    this.script = true;
    try {
      this.byHandle = new Map();
      this.byObject = new WeakMap();
      this.auto = null;
      const project = projectOf('ledger-vb-script', [
        domModule(),
        { id: 'script', name: 'Script', kind: 'module', code: String(source) },
      ]);
      machine = new this.engine.VirtualMachine(project, this.machineHost(this.registry()), { debuggerEnabled: false });
      this.script = machine;
      await machine.initialize();
      await this.bindLedger(machine);
      machine.setState('running');
      if (!named) {
        let main = null;
        let owner = null;
        for (const instance of machine.instances.values()) {
          const proc = instance.module.procedures.get('main');
          if (proc && proc.kind === 'sub' && proc.scope !== 'private') { main = proc; owner = instance; break; }
        }
        if (!main) throw new this.engine.VBError('Startup Sub Main or startup form was not found', 35);
        await machine.callProcedure(owner, main, []);
        return undefined;
      }
      const inst = machine.instances.get('script');
      const key = String(entry).toLowerCase();
      const proc = inst && (inst.module.procedures.get(key) || inst.module.procedures.get(key + ':get'));
      if (!proc || proc.scope === 'private') {
        const msg = L.t ? L.t('Procedure not found: {0}', entry) : String(entry);
        throw new this.engine.VBError(msg, 35);
      }
      const list = args || [];
      if (proc.kind === 'sub') { await machine.callProcedure(inst, proc, list); return undefined; }
      return this.engine.unbox(await machine.callProcedure(inst, proc, list));
    } finally {
      try { if (machine && machine.state !== 'stopped') machine.stop(); } catch (e) { /* already stopped */ }
      this.script = null;
      this.byHandle = saved.byHandle;
      this.byObject = saved.byObject;
      this.auto = saved.auto;
    }
  };

  /* Workbook macros. One project, one module each, kept after the call so they can call each other. */
  LedgerVbHost.prototype.unloadMacros = function () {
    const machine = this.macroMachine;
    this.macroMachine = null;
    this.macroMaps = null;
    if (this.script === machine) this.script = null;
    if (machine && machine.stop && machine.state !== 'stopped') {
      try { machine.stop(); } catch (e) { /* already stopped */ }
    }
  };

  LedgerVbHost.prototype.loadMacros = async function (macros) {
    if (this.script || (this.vm && this.vm.debugEvaluation)) {
      const err = new Error(this.script ? 'A VB script is running' : 'A VB statement is running');
      err.code = 'vb-busy';
      throw err;
    }
    if (!this.vm || this.vm.state === 'stopped' || !this.ir || !this.engine) throw new Error('The VB runtime is not running');
    this.unloadMacros();
    const saved = { byHandle: this.byHandle, byObject: this.byObject, auto: this.auto };
    this.byHandle = new Map();
    this.byObject = new WeakMap();
    this.auto = null;
    let machine = null;
    try {
      machine = new this.engine.VirtualMachine(projectOf('ledger-vb-macros', macroModules(macros)), this.machineHost(this.registry()), { debuggerEnabled: false });
      await machine.initialize();
      await this.bindLedger(machine);
      machine.setState('running');
      this.macroMachine = machine;
      this.macroMaps = { byHandle: this.byHandle, byObject: this.byObject, auto: this.auto };
    } catch (e) {
      if (machine && machine.stop) { try { machine.stop(); } catch (err) { /* already stopped */ } }
      throw e;
    } finally {
      this.byHandle = saved.byHandle;
      this.byObject = saved.byObject;
      this.auto = saved.auto;
    }
  };

  LedgerVbHost.prototype.runSaved = async function (name) {
    if (this.script) throw new Error('A VB script is running');
    if (this.vm && this.vm.debugEvaluation) throw new Error('A VB statement is running');
    if (!this.macroMachine || !this.macroMaps) throw new Error('The document macros are not loaded');
    const saved = { byHandle: this.byHandle, byObject: this.byObject, auto: this.auto };
    this.byHandle = this.macroMaps.byHandle;
    this.byObject = this.macroMaps.byObject;
    this.auto = this.macroMaps.auto;
    this.script = this.macroMachine;
    try {
      if (this.macroMachine.state !== 'running') this.macroMachine.setState('running');
      const found = procedure(this.macroMachine, name);
      if (!found) {
        const msg = L.t ? L.t('Procedure not found: {0}', name) : String(name);
        throw new this.engine.VBError(msg, 35);
      }
      await this.macroMachine.callProcedure(found.instance, found.proc, []);
    } finally {
      this.macroMaps = { byHandle: this.byHandle, byObject: this.byObject, auto: this.auto };
      this.script = null;
      this.byHandle = saved.byHandle;
      this.byObject = saved.byObject;
      this.auto = saved.auto;
    }
  };

  /* The editor buffer plus the other saved macros, on a machine that is thrown away. */
  LedgerVbHost.prototype.runDraft = async function (macros, entry) {
    if (this.script) throw new Error('A VB script is running');
    if (this.vm && this.vm.debugEvaluation) throw new Error('A VB statement is running');
    if (!this.vm || this.vm.state === 'stopped' || !this.ir || !this.engine) throw new Error('The VB runtime is not running');
    const saved = { byHandle: this.byHandle, byObject: this.byObject, auto: this.auto };
    this.byHandle = new Map();
    this.byObject = new WeakMap();
    this.auto = null;
    let machine = null;
    this.script = true;
    try {
      machine = new this.engine.VirtualMachine(projectOf('ledger-vb-draft', macroModules(macros)), this.machineHost(this.registry()), { debuggerEnabled: false });
      this.script = machine;
      await machine.initialize();
      await this.bindLedger(machine);
      machine.setState('running');
      const found = procedure(machine, entry);
      if (!found) {
        const msg = L.t ? L.t('Procedure not found: {0}', entry) : String(entry);
        throw new this.engine.VBError(msg, 35);
      }
      await machine.callProcedure(found.instance, found.proc, []);
    } finally {
      try { if (machine && machine.state !== 'stopped') machine.stop(); } catch (e) { /* already stopped */ }
      this.script = null;
      this.byHandle = saved.byHandle;
      this.byObject = saved.byObject;
      this.auto = saved.auto;
    }
  };

  function compileResult(engine, project) {
    const program = engine.compileProject(project);
    return { valid: !!program.valid, diagnostics: program.diagnostics || [] };
  }

  /* Syntax only. Same projects as debugBegin, run and runDraft. No machine is started. */
  LedgerVbHost.prototype.compileIde = function (source) {
    if (!this.engine) throw new Error('The VB runtime is not running');
    return compileResult(this.engine, projectOf('ledger-vb-ide', [
      domModule(),
      { id: 'module1', name: 'Module1', kind: 'module', code: String(source) },
    ]));
  };

  LedgerVbHost.prototype.compileScript = function (source) {
    if (!this.engine) throw new Error('The VB runtime is not running');
    return compileResult(this.engine, projectOf('ledger-vb-script', [
      domModule(),
      { id: 'script', name: 'Script', kind: 'module', code: String(source) },
    ]));
  };

  LedgerVbHost.prototype.compileDraft = function (macros) {
    if (!this.engine) throw new Error('The VB runtime is not running');
    return compileResult(this.engine, projectOf('ledger-vb-draft', macroModules(macros)));
  };

  /* Debug VM for the Visual Basic IDE. The Immediate VM stays. Maps stay swapped until debugEnd. */
  LedgerVbHost.prototype.debugBegin = async function (source) {
    if (this.script) throw new Error('A VB script is running');
    if (this.vm && this.vm.debugEvaluation) throw new Error('A VB statement is running');
    if (!this.vm || this.vm.state === 'stopped' || !this.ir || !this.engine) {
      throw new Error('The VB runtime is not running');
    }
    const project = projectOf('ledger-vb-ide', [
      domModule(),
      { id: 'module1', name: 'Module1', kind: 'module', code: String(source) },
    ]);
    const program = this.engine.compileProject(project);
    if (!program.valid) {
      const err = new Error(program.diagnostics.map(function (d) {
        return d.source + ':' + d.line + ': ' + d.message;
      }).join('\n'));
      err.diagnostics = program.diagnostics;
      err.number = 1002;
      throw err;
    }
    const saved = { byHandle: this.byHandle, byObject: this.byObject, auto: this.auto };
    this.byHandle = new Map();
    this.byObject = new WeakMap();
    this.auto = null;
    let machine = null;
    try {
      const vmHost = this.machineHost(this.registry());
      vmHost.print = function () {};
      vmHost.error = function () {};
      machine = new this.engine.VirtualMachine(program, vmHost, {
        debuggerEnabled: true,
        debugStatements: true,
        errorTrapping: 'unhandled',
      });
      const host = this;
      const initialize = machine.initialize.bind(machine);
      machine.initialize = async function () {
        await initialize();
        await host.bindLedger(machine);
      };
      this.script = machine;
      this.debugSaved = saved;
      return { vm: machine, moduleName: 'Module1' };
    } catch (error) {
      if (machine && machine.stop && machine.state !== 'stopped') {
        try { machine.stop(); } catch (e) { /* already stopped */ }
      }
      this.script = null;
      this.debugSaved = null;
      this.byHandle = saved.byHandle;
      this.byObject = saved.byObject;
      this.auto = saved.auto;
      throw error;
    }
  };

  LedgerVbHost.prototype.debugEnd = function () {
    if (!this.debugSaved) return;
    const saved = this.debugSaved;
    this.debugSaved = null;
    const machine = this.script;
    this.script = null;
    if (machine && machine.stop && machine.state !== 'stopped') {
      try { machine.stop(); } catch (e) { /* already stopped */ }
    }
    this.byHandle = saved.byHandle;
    this.byObject = saved.byObject;
    this.auto = saved.auto;
  };

  LedgerVbHost.prototype.eval = function (text) {
    if (this.script) return Promise.reject(new Error('A VB script is running'));
    if (!this.vm || this.vm.state === 'stopped') return Promise.reject(new Error('The VB runtime is not running'));
    return this.vm.immediate(String(text));
  };

  LedgerVbHost.prototype.close = function () {
    this.debugEnd();
    if (this.unloadMacros) this.unloadMacros();
    const script = this.script;
    this.script = null;
    if (script && script.stop && script.state !== 'stopped') {
      try { script.stop(); } catch (e) { /* already stopped */ }
    }
    const vm = this.vm;
    this.vm = null;
    if (vm && vm.state !== 'stopped') {
      try { vm.stop(); } catch (e) { /* already stopped */ }
    }
    if (this.ir) {
      try { this.ir.release(); } catch (e) { /* already released */ }
      this.ir = null;
    }
    this.auto = null;
  };

  function completionHost() {
    return {
      name: 'LedgerDOM',
      code: moduleCode(),
      files: false,
      types: {
        Ledger: 'Application',
        Application: 'Application',
        ActiveWorkbook: 'Workbook',
        ThisWorkbook: 'Workbook',
        Workbooks: 'Workbooks',
        ActiveSheet: 'Worksheet',
        ActiveCell: 'Range',
        Selection: 'Range',
        WorksheetFunction: 'WorksheetFunction',
      },
    };
  }

  let shared = null;
  let starting = null;

  function alive(host) {
    return !!(host && host.vm && host.vm.state !== 'stopped');
  }

  function session() {
    if (alive(shared)) return Promise.resolve(shared);
    if (starting) return starting;
    starting = open().then(function (host) {
      shared = host;
      starting = null;
      return host;
    }, function (error) {
      starting = null;
      throw error;
    });
    return starting;
  }

  function expressionResult(text) {
    return /^\?|^(?:Debug\.)?Print\b/i.test(String(text).trim());
  }

  async function evalVB6(text) {
    const host = await session();
    const value = await host.eval(text);
    return expressionResult(text) ? value : undefined;
  }

  async function runVB6(source, entry, args) {
    const host = await session();
    return host.run(source, entry, args);
  }

  async function loadMacros(macros) {
    const host = await session();
    return host.loadMacros(macros);
  }

  async function runSaved(name) {
    const host = await session();
    return host.runSaved(name);
  }

  async function runDraft(macros, entry) {
    const host = await session();
    return host.runDraft(macros, entry);
  }

  async function compileIde(source) {
    const host = await session();
    return host.compileIde(source);
  }

  async function compileScript(source) {
    const host = await session();
    return host.compileScript(source);
  }

  async function compileDraft(macros) {
    const host = await session();
    return host.compileDraft(macros);
  }

  async function open() {
    const host = new LedgerVbHost();
    await host.start();
    return host;
  }

  window.evalVB6 = evalVB6;
  window.runVB6 = runVB6;
  L.evalVB6 = evalVB6;
  L.runVB6 = runVB6;
  L.vb = {
    LedgerVbHost: LedgerVbHost,
    open: open,
    session: session,
    current: function () { return alive(shared) ? shared : null; },
    engine: loadEngine,
    completionHost: completionHost,
    loadMacros: loadMacros,
    runSaved: runSaved,
    runDraft: runDraft,
    compileIde: compileIde,
    compileScript: compileScript,
    compileDraft: compileDraft,
  };
})();
