/* Quire VB host. Immediate VM for evalVB6, a second VM for runVB6.
   Form slots stay empty until a form project is loaded.
   Build the engine first: make */
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
      return Promise.reject(new Error('VB6 engine is not loaded. Run make'));
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
        reject(new Error('VB6 engine is not built. Run make'));
      };
      document.body.appendChild(script);
    });
  }

  function memberMeta(name, desc) {
    if (desc.kind === 'gap') {
      const as = desc.as || 'any';
      if (as === 'call') return { name: name, modes: [1], params: CALL_PARAMS };
      if (as === 'get') return { name: name, modes: [2], params: [] };
      if (as === 'rw') return { name: name, modes: [2, 4, 8], params: [] };
      return { name: name, modes: [1, 2, 4, 8], params: CALL_PARAMS };
    }
    if (desc.kind === 'call') return { name: name, modes: [1], params: CALL_PARAMS };
    if (desc.kind === 'rw') return { name: name, modes: [2, 4, 8], params: [] };
    return { name: name, modes: [2], params: [] };
  }

  function moduleCode() {
    const lines = [
      'Public Quire As Object',
      '',
      'Public Property Get Application() As Object',
      '    Set Application = Quire',
      'End Property',
      '',
      'Public Property Get ActiveDocument() As Object',
      '    Set ActiveDocument = Quire.ActiveDocument',
      'End Property',
      '',
      'Public Property Get Documents() As Object',
      '    Set Documents = Quire.Documents',
      'End Property',
      '',
      'Public Property Get Selection() As Object',
      '    Set Selection = Quire.Selection',
      'End Property',
      '',
    ];
    const names = Object.keys(L.dom).filter(function (key) {
      const value = L.dom[key];
      return /^wd[A-Za-z][A-Za-z0-9]*$/.test(key) && typeof value === 'number' && Number.isInteger(value) && value >= -2147483648 && value <= 2147483647;
    }).sort();
    for (let i = 0; i < names.length; i++) lines.push('Public Const ' + names[i] + ' As Long = ' + L.dom[names[i]]);
    const gaps = gapNames();
    if (gaps.length) lines.push('');
    for (let i = 0; i < gaps.length; i++) lines.push('Public ' + gaps[i] + ' As Object');
    return lines.join('\n');
  }

  function gapNames() {
    const types = L.domIr && L.domIr.types;
    if (!types) return [];
    return Object.keys(types).filter(function (name) {
      return types[name].gap && /^[A-Za-z][A-Za-z0-9]*$/.test(name);
    }).sort();
  }

  function QuireVbHost() {
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

  QuireVbHost.prototype.metadata = function (type) {
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
    if (spec && spec.gap) meta.acceptAnyMember = true;
    if (spec && spec.collection && members.some(function (m) { return m.name === 'Item'; })) meta.defaultMember = 'Item';
    this.meta.set(type, meta);
    return meta;
  };

  QuireVbHost.prototype.adapter = function (handle, type) {
    const host = this;
    return {
      metadata: host.metadata(type),
      invoke: function (name, mode, values) { return host.dispatch(handle, name, mode, values); },
      release: function () {},
    };
  };

  QuireVbHost.prototype.remember = function (handle, object) {
    this.byHandle.set(handle, object);
    this.byObject.set(object, handle);
    return object;
  };

  QuireVbHost.prototype.objectFor = function (handle, type) {
    const hit = this.byHandle.get(handle);
    if (hit) return hit;
    if (!this.auto) throw new Error('Automation session is not open');
    return this.remember(handle, this.auto.adopt(this.adapter(handle, type || 'Object')));
  };

  QuireVbHost.prototype.toIr = function (value) {
    if (value && value.__missing === true) return L.domIr.Missing;
    if (value && typeof value === 'object') {
      const handle = this.byObject.get(value);
      if (handle != null) return { handle: handle };
    }
    return value;
  };

  QuireVbHost.prototype.dispatch = async function (handle, name, mode, values) {
    const args = [];
    for (let i = 0; i < values.length; i++) args.push(this.toIr(values[i]));
    let req;
    if (mode === 4 || mode === 8) {
      req = { op: 'set', handle: handle, member: name, value: args.pop() };
    } else if (mode === 1 && String(name).toLowerCase() === 'item') {
      req = { op: 'item', handle: handle, member: 'Item', args: args };
    } else if (mode === 1) {
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
    return { value: result.kind === 'empty' ? undefined : result.value };
  };

  function projectOf(id, modules) {
    return {
      schema: 1,
      id: id,
      name: 'Quire',
      startup: '',
      modules: modules,
      settings: {},
      references: [],
      assets: {},
    };
  }

  QuireVbHost.prototype.machineHost = function (registry) {
    const host = this;
    const vmHost = {
      automation: registry,
      print: function (text) { host.output.push(text == null ? '' : String(text)); },
      error: function (err) { host.errors.push(err); },
    };
    for (let i = 0; i < SLOTS.length; i++) {
      (function (name) {
        Object.defineProperty(vmHost, name, { enumerable: true, get: function () { return host[name]; } });
      })(SLOTS[i]);
    }
    return vmHost;
  };

  QuireVbHost.prototype.registry = function () {
    const host = this;
    const registry = new this.engine.AutomationRegistry();
    registry.register('Quire.Application', function (session) {
      host.auto = session;
      return host.adapter(host.ir.root.handle, 'Application');
    });
    return registry;
  };

  QuireVbHost.prototype.bindQuire = async function (machine) {
    const app = await machine.automation.create('Quire.Application');
    this.remember(this.ir.root.handle, app);
    const inst = machine.instances.get('quiredom');
    const cell = inst && inst.fields.get('quire');
    if (!cell) throw new Error('The injected module has no Quire object');
    cell.set(app);
    const gaps = gapNames();
    for (let i = 0; i < gaps.length; i++) {
      const named = this.ir.named(gaps[i]);
      if (!named.ok) continue;
      const field = inst.fields.get(gaps[i].toLowerCase());
      if (field) field.set(this.objectFor(named.handle, named.type));
    }
  };

  QuireVbHost.prototype.start = async function () {
    if (this.vm) return this;
    const engine = await loadEngine();
    this.engine = engine;
    this.ir = L.domIr.open();
    const project = projectOf('quire-vb', [
      { id: 'quire-dom', name: 'QuireDOM', kind: 'module', code: moduleCode() },
    ]);
    try {
      this.vm = new engine.VirtualMachine(project, this.machineHost(this.registry()), { debuggerEnabled: true });
      await this.vm.prepareImmediateContext();
      await this.bindQuire(this.vm);
    } catch (error) {
      this.close();
      throw error;
    }
    return this;
  };

  /* Script VM. Its Automation maps replace the Immediate maps only while it runs. */
  QuireVbHost.prototype.run = async function (source, entry, args) {
    if (this.script) throw new Error('A VB script is running');
    if (this.vm && this.vm.debugEvaluation) throw new Error('A VB statement is running');
    if (!this.vm || this.vm.state === 'stopped' || !this.ir || !this.engine) {
      throw new Error('The VB runtime is not running');
    }
    if (args != null && !Array.isArray(args)) throw new Error('Procedure arguments must be an array');
    const saved = { byHandle: this.byHandle, byObject: this.byObject, auto: this.auto };
    const named = entry != null && entry !== '';
    let machine = null;
    this.script = true;
    try {
      this.byHandle = new Map();
      this.byObject = new WeakMap();
      this.auto = null;
      const project = projectOf('quire-vb-script', [
        { id: 'quire-dom', name: 'QuireDOM', kind: 'module', code: moduleCode() },
        { id: 'script', name: 'Script', kind: 'module', code: String(source) },
      ]);
      machine = new this.engine.VirtualMachine(project, this.machineHost(this.registry()), { debuggerEnabled: false });
      this.script = machine;
      await machine.initialize();
      await this.bindQuire(machine);
      machine.setState('running');
      if (!named) {
        let main = null;
        let owner = null;
        for (const instance of machine.instances.values()) {
          const proc = instance.module.procedures.get('main');
          if (proc && proc.kind === 'sub' && proc.scope !== 'private') {
            main = proc;
            owner = instance;
            break;
          }
        }
        if (!main) throw new this.engine.VBError('Startup Sub Main or startup form was not found', 35);
        await machine.callProcedure(owner, main, []);
        return undefined;
      }
      const inst = machine.instances.get('script');
      const key = String(entry).toLowerCase();
      const proc = inst && (inst.module.procedures.get(key) || inst.module.procedures.get(key + ':get'));
      if (!proc || proc.scope === 'private') throw new this.engine.VBError('Procedure not found: ' + entry, 35);
      const list = args || [];
      if (proc.kind === 'sub') {
        await machine.callProcedure(inst, proc, list);
        return undefined;
      }
      return this.engine.unbox(await machine.callProcedure(inst, proc, list));
    } finally {
      try {
        if (machine && machine.state !== 'stopped') machine.stop();
      } catch (e) { /* already stopped */ }
      this.script = null;
      this.byHandle = saved.byHandle;
      this.byObject = saved.byObject;
      this.auto = saved.auto;
    }
  };

  QuireVbHost.prototype.eval = function (text) {
    if (this.script) return Promise.reject(new Error('A VB script is running'));
    if (!this.vm || this.vm.state === 'stopped') return Promise.reject(new Error('The VB runtime is not running'));
    return this.vm.immediate(String(text));
  };

  QuireVbHost.prototype.close = function () {
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

  /* ?expr and Print are expressions. Any other Immediate line is a statement. */
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

  async function open() {
    const host = new QuireVbHost();
    await host.start();
    return host;
  }

  window.evalVB6 = evalVB6;
  window.runVB6 = runVB6;
  L.evalVB6 = evalVB6;
  L.runVB6 = runVB6;
  L.vb = {
    QuireVbHost: QuireVbHost,
    open: open,
    session: session,
  };
})();
