/* VB6 forms on the suite's own windows. The virtual machine calls createForm,
   Show, Hide and the control properties. Pixels are a ui.dialog in the current
   look. This is not ApplicationHost and not BrowserForm. */
(function () {
  'use strict';
  const root = typeof window !== 'undefined' ? window : globalThis;
  const L = root.L || (root.L = {});
  const VO = root.VO || (root.VO = {});

  const DRAWN = { Label: 1, TextBox: 1, CommandButton: 1, CheckBox: 1, OptionButton: 1, ListBox: 1, ComboBox: 1, Frame: 1, HScrollBar: 1, VScrollBar: 1 };
  const TEXT_TYPE = { TextBox: 1, ComboBox: 1, ListBox: 1 };
  const VALUE_TYPE = { CheckBox: 1, OptionButton: 1, HScrollBar: 1, VScrollBar: 1 };
  const nodes = new WeakMap();
  const PROP_NAMES = ['Caption', 'Text', 'Value', 'Enabled', 'Visible', 'Left', 'Top', 'Width', 'Height', 'ClientWidth', 'ClientHeight', 'TabIndex', 'TabStop', 'ToolTipText', 'Tag', 'FontName', 'FontSize', 'FontBold', 'FontItalic', 'Alignment', 'BackColor', 'ForeColor', 'MultiLine', 'Locked', 'MaxLength', 'PasswordChar', 'Min', 'Max', 'SmallChange', 'LargeChange', 'Interval', 'Style', 'Sorted', 'ListIndex', 'Default', 'Cancel', 'Name'];

  function fail(message, number) {
    const err = new Error(message);
    err.name = 'VBError';
    err.number = number;
    throw err;
  }

  function on(value) {
    return !(value === 0 || value === false || value === null || value === undefined);
  }

  function px(twips) {
    const n = Number(twips);
    return (Number.isFinite(n) ? n : 0) / 15;
  }

  function oleColor(value) {
    const n = Number(value);
    if (!Number.isFinite(n)) return '';
    if (n < 0) {
      const index = n + 2147483648;
      if (index === 5) return 'var(--surface)';
      if (index === 8 || index === 18) return 'var(--dlg-ink)';
      return 'var(--dlg)';
    }
    return 'rgb(' + (n & 255) + ',' + ((n >> 8) & 255) + ',' + ((n >> 16) & 255) + ')';
  }

  function stripAmp(text) {
    return String(text == null ? '' : text).replace(/&&/g, '\0').replace(/&/g, '').replace(/\0/g, '&');
  }

  function el(tag, attrs) {
    if (typeof document !== 'undefined' && L.h) return L.h(tag, attrs || {});
    const node = {
      tagName: String(tag || 'div').toUpperCase(),
      style: {},
      children: [],
      hidden: false,
      disabled: false,
      value: '',
      checked: false,
      readOnly: false,
      selectedIndex: -1,
      className: attrs && attrs.class ? attrs.class : '',
      classList: { add() {}, remove() {}, toggle() {} },
      _on: {},
      textContent: '',
      appendChild(child) { if (child) this.children.push(child); return child; },
      append() { for (let i = 0; i < arguments.length; i++) if (arguments[i]) this.appendChild(arguments[i]); },
      setAttribute(name, value) { this[name] = value; },
      getAttribute(name) { return this[name] == null ? null : this[name]; },
      removeAttribute() {},
      addEventListener(type, fn) { (this._on[type] = this._on[type] || []).push(fn); },
      removeEventListener() {},
      focus() { this.focused = true; },
      select() {},
      click() {
        const list = this._on.click || [];
        for (let i = 0; i < list.length; i++) list[i]({ target: this, preventDefault() {}, stopPropagation() {} });
      },
      querySelector() { return null; },
    };
    if (attrs && attrs.type) node.type = attrs.type;
    if (attrs && attrs.text) node.textContent = attrs.text;
    return node;
  }

  function Control(model, form) {
    const props = {
      Left: 0, Top: 0, Width: 1200, Height: 300, Visible: -1, Enabled: -1, TabIndex: 0, TabStop: -1,
      ToolTipText: '', Tag: '', FontName: 'MS Sans Serif', FontSize: 8.25, FontBold: 0, FontItalic: 0,
      Alignment: 0, BackColor: -2147483633, ForeColor: -2147483640, Caption: '', Text: '', Value: 0,
      MultiLine: 0, Locked: 0, MaxLength: 0, PasswordChar: '', Min: 0, Max: 32767, SmallChange: 1, LargeChange: 1,
      Interval: 0, Style: 0, Sorted: 0, ListIndex: -1, Default: 0, Cancel: 0,
    };
    const given = model.properties || {};
    for (const key in given) if (Object.prototype.hasOwnProperty.call(given, key)) props[key] = given[key];
    props.Name = model.name;
    this.__control = true;
    this.__type = model.type;
    this.type = model.type;
    this.model = model;
    this.form = form;
    this._props = props;
    this._items = Array.isArray(props.List) ? props.List.map(String) : [];
    this._itemData = [];
    this._selected = new Set();
    this._timer = null;
    const self = this;
    for (let i = 0; i < PROP_NAMES.length; i++) {
      (function (key) {
        Object.defineProperty(self, key, {
          enumerable: true,
          configurable: true,
          get: function () { return self._get(key); },
          set: function (value) { self._set(key, value, false); },
        });
      })(PROP_NAMES[i]);
    }
    if (given.Index !== undefined) {
      Object.defineProperty(this, 'Index', {
        enumerable: true,
        get: function () { return self._props.Index; },
      });
      this._props.Index = Number(given.Index);
    }
    Object.defineProperty(this, 'ListCount', { enumerable: true, get: function () { return self._items.length; } });
    Object.defineProperty(this, 'SelCount', { enumerable: true, get: function () { return self._selected.size; } });
    this._build();
    this._paint();
  }

  Control.prototype._get = function (key) {
    if (key === 'Text' && (this.type === 'ListBox' || this.type === 'ComboBox') && Number(this._props.ListIndex) >= 0) {
      return this._items[Number(this._props.ListIndex)] || '';
    }
    if (key === 'Value' && this.type === 'OptionButton') return on(this._props.Value) ? -1 : 0;
    const value = this._props[key];
    return value == null ? (key === 'Text' || key === 'Caption' || key === 'Tag' || key === 'ToolTipText' || key === 'PasswordChar' ? '' : 0) : value;
  };

  Control.prototype._set = function (key, value, fromDom) {
    if (key === 'ListIndex') {
      value = Number(value);
      if (value < -1 || value >= this._items.length) fail('Invalid property value: ListIndex', 380);
    }
    if (key === 'Value' && (this.type === 'HScrollBar' || this.type === 'VScrollBar')) {
      value = Number(value);
      const lo = Math.min(Number(this._props.Min), Number(this._props.Max));
      const hi = Math.max(Number(this._props.Min), Number(this._props.Max));
      if (!Number.isFinite(value) || value < lo || value > hi) fail('Invalid property value: Value', 380);
    }
    if (key === 'Interval') {
      value = Number(value);
      if (value < 0 || value > 65535) fail('Invalid Timer interval', 380);
    }
    const previous = this._props[key];
    if (key === 'Text' || key === 'Caption' || key === 'Tag' || key === 'ToolTipText' || key === 'PasswordChar' || key === 'FontName') value = value == null ? '' : String(value);
    this._props[key] = value;
    if (key === 'Value' && this.type === 'OptionButton' && on(value)) {
      const parent = this.model.parent || null;
      const controls = this.form._controls;
      for (let i = 0; i < controls.length; i++) {
        const other = controls[i];
        if (other !== this && other.type === 'OptionButton' && (other.model.parent || null) === parent) other._props.Value = 0;
      }
    }
    if (key === 'Text' && this.type === 'ComboBox') {
      const at = this._items.indexOf(String(value));
      this._props.ListIndex = at;
    }
    this._paint();
    if (this.type === 'Timer' && (key === 'Interval' || key === 'Enabled')) this._armTimer();
    if (fromDom || previous === value) return;
    if (key === 'Text') this._emit('Change');
    if (key === 'Value' && (this.type === 'HScrollBar' || this.type === 'VScrollBar')) this._emit('Change');
    if (key === 'ListIndex') this._emit('Click');
  };

  Control.prototype._eventName = function (name) {
    return this.model.name + '_' + name;
  };

  Control.prototype._emit = function (name, args) {
    const all = this._props.Index !== undefined ? [Number(this._props.Index)].concat(args || []) : (args || []);
    const vm = this.form && this.form.vm;
    const pending = vm && vm.dispatch ? vm.dispatch(this.form.instance, this._eventName(name), all) : Promise.resolve();
    this._lastEvent = pending;
    return pending;
  };

  Control.prototype._build = function () {
    const type = this.type;
    if (!DRAWN[type]) {
      nodes.set(this, null);
      return;
    }
    const node = el('div', { class: 'vb-ctl vb-' + type.toLowerCase() });
    nodes.set(this, node);
    const self = this;
    if (type === 'CommandButton') {
      this._face = el('button', { class: 'btn', type: 'button' });
      this._face.addEventListener('click', function () { self._emit('Click'); });
      node.appendChild(this._face);
    } else if (type === 'TextBox') {
      this._face = el(on(this._props.MultiLine) ? 'textarea' : 'input', { type: 'text' });
      this._face.addEventListener('input', function () { self._set('Text', self._face.value, true); });
      node.appendChild(this._face);
    } else if (type === 'CheckBox' || type === 'OptionButton') {
      this._face = el('input', { type: type === 'CheckBox' ? 'checkbox' : 'radio' });
      this._cap = el('span');
      this._face.addEventListener('change', function () {
        const next = self._face.checked ? (type === 'CheckBox' ? 1 : -1) : 0;
        self._set('Value', next, true);
        self._emit('Click');
      });
      node.appendChild(this._face);
      node.appendChild(this._cap);
    } else if (type === 'ListBox' || (type === 'ComboBox' && Number(this._props.Style) === 2)) {
      this._face = el('select');
      if (type === 'ListBox') this._face.size = 4;
      this._face.addEventListener('change', function () {
        self._props.ListIndex = self._face.selectedIndex;
        if (type === 'ComboBox') self._props.Text = self._items[self._face.selectedIndex] || '';
        self._emit('Click');
      });
      node.appendChild(this._face);
    } else if (type === 'ComboBox') {
      this._face = el('input', { type: 'text' });
      this._face.addEventListener('input', function () { self._set('Text', self._face.value, true); });
      node.appendChild(this._face);
    } else if (type === 'Frame') {
      this._cap = el('div', { class: 'vb-frame-cap' });
      this._client = el('div', { class: 'vb-frame-client' });
      node.appendChild(this._cap);
      node.appendChild(this._client);
    } else if (type === 'HScrollBar' || type === 'VScrollBar') {
      this._face = el('input', { type: 'range' });
      this._face.addEventListener('input', function () {
        self._props.Value = Number(self._face.value);
        self._emit('Scroll');
      });
      this._face.addEventListener('change', function () { self._set('Value', Number(self._face.value), true); self._emit('Change'); });
      node.appendChild(this._face);
    } else if (type === 'Label') {
      this._cap = el('span');
      node.appendChild(this._cap);
    }
    node.addEventListener('focusin', function () { self._emit('GotFocus'); });
    node.addEventListener('focusout', function () { self._emit('LostFocus'); });
  };

  Control.prototype._paint = function () {
    const node = nodes.get(this);
    if (!node) return;
    const p = this._props;
    node.style.left = px(p.Left) + 'px';
    node.style.top = px(p.Top) + 'px';
    node.style.width = Math.max(1, px(p.Width)) + 'px';
    node.style.height = Math.max(1, px(p.Height)) + 'px';
    node.style.fontSize = (Number(p.FontSize) || 8.25) * 96 / 72 + 'px';
    node.style.fontWeight = on(p.FontBold) ? 'bold' : 'normal';
    node.style.fontStyle = on(p.FontItalic) ? 'italic' : 'normal';
    node.style.color = oleColor(p.ForeColor);
    node.hidden = !on(p.Visible);
    node.title = String(p.ToolTipText || '');
    if (this._face) {
      this._face.disabled = !on(p.Enabled);
      if (this.type === 'TextBox' || (this.type === 'ComboBox' && Number(p.Style) !== 2)) {
        if (this._face.value !== String(p.Text == null ? '' : p.Text)) this._face.value = String(p.Text == null ? '' : p.Text);
        this._face.style.textAlign = Number(p.Alignment) === 1 ? 'right' : Number(p.Alignment) === 2 ? 'center' : 'left';
        this._face.readOnly = on(p.Locked);
        if (p.PasswordChar && this._face.type !== undefined) this._face.type = 'password';
      }
      if (this.type === 'CheckBox' || this.type === 'OptionButton') this._face.checked = on(p.Value);
      if (this.type === 'CommandButton') this._face.textContent = stripAmp(p.Caption);
      if (this.type === 'HScrollBar' || this.type === 'VScrollBar') {
        this._face.min = p.Min;
        this._face.max = p.Max;
        this._face.value = p.Value;
      }
      if ((this.type === 'ListBox' || (this.type === 'ComboBox' && Number(p.Style) === 2)) && this._face.options) {
        const labels = this._items.map(String);
        if (typeof document !== 'undefined' && L.h) {
          this._face.replaceChildren.apply(this._face, labels.map(function (text, index) { return L.h('option', { value: String(index), text: text }); }));
          this._face.selectedIndex = Number(p.ListIndex);
        }
      }
    }
    if (this._cap && this.type !== 'CommandButton') this._cap.textContent = stripAmp(p.Caption);
    if (this.type === 'Frame' && this._client) {
      this._client.style.position = 'relative';
      this._client.style.width = '100%';
      this._client.style.height = '100%';
    }
    if (this.type === 'Label' || this.type === 'Frame') node.style.background = Number(p.BackColor) === 0 ? 'transparent' : oleColor(p.BackColor);
  };

  Control.prototype._armTimer = function () {
    const self = this;
    if (this._timer) { clearInterval(this._timer); this._timer = null; }
    if (this.type !== 'Timer' || !this.form || !this.form.shown) return;
    const interval = Number(this._props.Interval) || 0;
    if (!on(this._props.Enabled) || interval <= 0) return;
    this._timer = setInterval(function () { self._emit('Timer'); }, interval);
  };

  Control.prototype.defaultValue = function () {
    if (TEXT_TYPE[this.type]) return this.Text;
    if (VALUE_TYPE[this.type]) return this.Value;
    return this.Caption;
  };

  Control.prototype.defaultRef = function () {
    const key = TEXT_TYPE[this.type] ? 'Text' : VALUE_TYPE[this.type] ? 'Value' : 'Caption';
    const self = this;
    return { get: function () { return self[key]; }, set: function (value) { self[key] = value; }, getScalar: function () { return self[key]; } };
  };

  Control.prototype.AddItem = function (text, index) {
    const at = index === undefined || index === null ? this._items.length : Number(index);
    if (at < 0 || at > this._items.length) fail('Invalid property value', 380);
    this._items.splice(at, 0, String(text));
    this._itemData.splice(at, 0, 0);
    if (on(this._props.Sorted)) this._items.sort(function (a, b) { return String(a).localeCompare(String(b)); });
    this._paint();
  };

  Control.prototype.RemoveItem = function (index) {
    index = Number(index);
    if (index < 0 || index >= this._items.length) fail('Invalid property array index', 381);
    this._items.splice(index, 1);
    this._itemData.splice(index, 1);
    this._props.ListIndex = -1;
    this._paint();
  };

  Control.prototype.Clear = function () {
    this._items = [];
    this._itemData = [];
    this._selected.clear();
    this._props.ListIndex = -1;
    this._paint();
  };

  Control.prototype.List = function (index) {
    index = Number(index);
    if (index < 0 || index >= this._items.length) fail('Invalid property array index', 381);
    return this._items[index];
  };

  Control.prototype.ItemData = function (index) {
    index = Number(index);
    if (index < 0 || index >= this._items.length) fail('Invalid property array index', 381);
    return this._itemData[index] || 0;
  };

  Control.prototype.Selected = function (index) {
    return this._selected.has(Number(index)) ? -1 : 0;
  };

  Control.prototype.setIndexed = function (name, args, value) {
    const key = String(name || '').toLowerCase();
    const index = Number(args && args[0]);
    if (key === 'list') { this._items[index] = String(value); this._paint(); return; }
    if (key === 'itemdata') { this._itemData[index] = Number(value) || 0; return; }
    if (key === 'selected') {
      if (on(value)) this._selected.add(index); else this._selected.delete(index);
      this._paint();
    }
  };

  Control.prototype.SetFocus = function () {
    const node = nodes.get(this);
    if (node && node.focus) node.focus();
    else if (this._face && this._face.focus) this._face.focus();
  };

  Control.prototype.Move = function (left, top, width, height) {
    if (arguments.length > 0 && left != null) this.Left = Number(left);
    if (arguments.length > 1 && top != null) this.Top = Number(top);
    if (arguments.length > 2 && width != null) this.Width = Number(width);
    if (arguments.length > 3 && height != null) this.Height = Number(height);
  };

  Control.prototype.Refresh = function () { this._paint(); };

  function ControlArray(template, form) {
    this.items = new Map();
    this.__type = 'ControlArray';
    this._template = template;
    this._form = form;
    const self = this;
    this.Load = function (index) { return self._load(index); };
    this.Unload = function (index) { return self._unload(index); };
    this.Item = function (index) {
      const item = self.items.get(Number(index));
      if (!item) fail('Control array element does not exist', 340);
      return item;
    };
    Object.defineProperty(this, 'Count', { enumerable: true, get: function () { return self.items.size; } });
  }

  ControlArray.prototype._bounds = function (last) {
    if (!this.items.size) fail('Control array has no elements', 9);
    let result = last ? -Infinity : Infinity;
    this.items.forEach(function (item, index) { result = last ? Math.max(result, index) : Math.min(result, index); });
    return result;
  };

  ControlArray.prototype._load = function (index) {
    index = Number(index);
    if (!Number.isInteger(index) || index < 0 || index > 32767) fail('Invalid control array index', 380);
    if (this.items.has(index)) fail('Object already loaded', 360);
    const model = JSON.parse(JSON.stringify(this._template));
    model.properties = model.properties || {};
    model.properties.Index = index;
    model.properties.Visible = 0;
    const control = new Control(model, this._form);
    this.items.set(index, control);
    this._form._controls.push(control);
    this._form._mount(control);
    return control;
  };

  ControlArray.prototype._unload = function (index) {
    index = Number(index);
    const control = this.items.get(index);
    if (!control) fail('Control array element does not exist', 340);
    this.items.delete(index);
    const node = nodes.get(control);
    if (node && node.remove) node.remove();
    if (control._timer) clearInterval(control._timer);
  };

  Object.defineProperty(ControlArray.prototype, 'LBound', { get: function () { return this._bounds(false); } });
  Object.defineProperty(ControlArray.prototype, 'UBound', { get: function () { return this._bounds(true); } });

  function MenuItem(form, model) {
    const props = { Caption: model.name, Enabled: -1, Visible: -1, Checked: 0 };
    const given = model.properties || {};
    for (const key in given) if (Object.prototype.hasOwnProperty.call(given, key)) props[key] = given[key];
    this.Name = model.name;
    this._props = props;
    this._form = form;
    const self = this;
    ['Caption', 'Enabled', 'Visible', 'Checked'].forEach(function (key) {
      Object.defineProperty(self, key, {
        enumerable: true,
        get: function () { return self._props[key]; },
        set: function (value) { self._props[key] = value; if (self._form._paintMenus) self._form._paintMenus(); },
      });
    });
  }

  function Form(model, instance, vm) {
    Control.call(this, Object.assign({}, model, { type: model.type || 'Form' }), null);
    this.form = this;
    this.vm = vm;
    this.instance = instance;
    this.shown = false;
    this.controls = [];
    this._controls = [];
    this.controlMap = new Map();
    this._shell = null;
    this._client = el('div', { class: 'vb-form-client' });
    this._menuBar = el('div', { class: 'vb-form-menu' });
    const width = this._props.ClientWidth != null ? this._props.ClientWidth : (this._props.Width || 4800);
    const height = this._props.ClientHeight != null ? this._props.ClientHeight : (this._props.Height || 3600);
    this._props.ClientWidth = width;
    this._props.ClientHeight = height;
    this._client.style.width = px(width) + 'px';
    this._client.style.height = px(height) + 'px';
    this._client.style.position = 'relative';
    const self = this;
    this.Controls = {
      get Count() { return self._controls.length; },
      Item: function (key) {
        if (typeof key === 'number') {
          const control = self._controls[key];
          if (!control) fail('Control not found: ' + key, 35601);
          return control;
        }
        const found = self.controlMap.get(String(key).toLowerCase());
        if (!found) fail('Control not found: ' + key, 35601);
        return found;
      },
    };
    this._mountAll(model.controls || []);
    this._mountMenus(model.menus || []);
    nodes.set(this, this._client);
  }

  Form.prototype = Object.create(Control.prototype);
  Form.prototype.constructor = Form;

  Form.prototype._eventName = function (name) { return 'Form_' + name; };

  Form.prototype._mountAll = function (models) {
    const arrays = new Map();
    for (let i = 0; i < models.length; i++) {
      const model = models[i];
      const control = new Control(model, this);
      this._controls.push(control);
      const key = String(model.name).toLowerCase();
      if (model.properties && model.properties.Index !== undefined) {
        let array = arrays.get(key);
        if (!array) {
          array = new ControlArray(model, this);
          arrays.set(key, array);
          this.controlMap.set(key, array);
        }
        array.items.set(Number(model.properties.Index), control);
      } else this.controlMap.set(key, control);
    }
    for (let i = 0; i < this._controls.length; i++) this._mount(this._controls[i]);
  };

  Form.prototype._mount = function (control) {
    const node = nodes.get(control);
    if (!node) return;
    let host = this._client;
    const parentName = control.model.parent;
    if (parentName) {
      const parent = this.controlMap.get(String(parentName).toLowerCase());
      const parentControl = parent && parent.__type === 'ControlArray' ? null : parent;
      if (parentControl && parentControl._client) host = parentControl._client;
    }
    host.appendChild(node);
  };

  Form.prototype._mountMenus = function (menus) {
    for (let i = 0; i < menus.length; i++) {
      const menu = new MenuItem(this, menus[i]);
      this.controlMap.set(String(menus[i].name).toLowerCase(), menu);
    }
    this._menus = menus;
    this._paintMenus();
  };

  Form.prototype._paintMenus = function () {
    const bar = this._menuBar;
    if (!bar) return;
    const roots = (this._menus || []).filter(function (menu) { return !menu.parent; });
    bar.hidden = !roots.length;
    const self = this;
    const buttons = roots.map(function (menu) {
      const object = self.controlMap.get(String(menu.name).toLowerCase());
      if (object && !on(object.Visible)) return null;
      const button = el('button', { class: 'btn', type: 'button' });
      button.textContent = stripAmp(object ? object.Caption : menu.name);
      button.disabled = object ? !on(object.Enabled) : false;
      button.addEventListener('click', function () {
        if (self.vm && self.vm.dispatch) self.vm.dispatch(self.instance, menu.name + '_Click', []);
      });
      return button;
    }).filter(Boolean);
    if (bar.replaceChildren) bar.replaceChildren.apply(bar, buttons);
    else { bar.children = buttons; buttons.forEach(function (button) { bar.appendChild(button); }); }
  };

  Form.prototype._requestClose = function () {
    if (this.vm && this.vm.requestUnload) this.vm.requestUnload(this.instance);
    else this.Hide();
  };

  Form.prototype.Show = function (modal) {
    this.shown = true;
    this._open(on(modal));
    this._paint();
    for (let i = 0; i < this._controls.length; i++) {
      this._controls[i]._paint();
      this._controls[i]._armTimer();
    }
    const self = this;
    const first = this._controls.filter(function (control) {
      return DRAWN[control.type] && on(control.TabStop) && on(control.Visible) && on(control.Enabled);
    }).sort(function (a, b) { return Number(a.TabIndex) - Number(b.TabIndex); })[0];
    if (first) first.SetFocus();
    else if (this._shell && this._shell.focus) this._shell.focus();
    return self;
  };

  Form.prototype._paint = function () {
    if (!this._client) return;
    this._client.style.width = px(this._props.ClientWidth) + 'px';
    this._client.style.height = px(this._props.ClientHeight) + 'px';
    const title = this._shell && this._shell.querySelector && this._shell.querySelector('.dlg-ttl');
    if (title) title.textContent = stripAmp(this.Caption || this.model.name || '');
  };

  Form.prototype._open = function (modal) {
    if (this._shell) return;
    const ui = L.ui;
    const self = this;
    if (ui && ui.dialog && typeof document !== 'undefined') {
      const wrap = L.h('div', null, this._menuBar, this._client);
      const dialog = ui.dialog({
        title: stripAmp(this.Caption || this.model.name || ''),
        width: Math.max(180, Math.round(px(this._props.ClientWidth) + 28)),
        modeless: !modal,
        canClose: false,
        buttons: [],
        body: wrap,
      });
      dialog.el.classList.add('vb-form-dlg');
      const bar = dialog.el.querySelector('.dlg-title');
      if (bar) {
        const close = L.h('button', { class: 'dlg-x', type: 'button', 'aria-label': 'Close', text: '\u00d7' });
        close.addEventListener('click', function (event) { event.preventDefault(); event.stopPropagation(); self._requestClose(); });
        bar.appendChild(close);
      }
      this._dialog = dialog;
      this._shell = dialog.el;
      dialog.el.addEventListener('keydown', function (event) { self._keys(event); });
      return;
    }
    this._shell = this._client;
  };

  Form.prototype._keys = function (event) {
    if (!event || event.defaultPrevented) return;
    if (event.key === 'Enter') {
      const button = this._controls.filter(function (control) { return control.type === 'CommandButton' && on(control.Default) && on(control.Enabled); })[0];
      if (button) { event.preventDefault(); button._emit('Click'); }
    } else if (event.key === 'Escape') {
      const button = this._controls.filter(function (control) { return control.type === 'CommandButton' && on(control.Cancel) && on(control.Enabled); })[0];
      event.preventDefault();
      if (button) button._emit('Click');
      else this._requestClose();
    }
  };

  Form.prototype.Hide = function () {
    this.shown = false;
    for (let i = 0; i < this._controls.length; i++) {
      if (this._controls[i]._timer) { clearInterval(this._controls[i]._timer); this._controls[i]._timer = null; }
    }
    if (this._dialog) { this._dialog.close(null); this._dialog = null; }
    this._shell = null;
  };

  Form.prototype.setInert = function (value) {
    if (this._shell) this._shell.inert = !!value;
  };

  Form.prototype.initializeDataBindings = function () {};

  function buttonSet(style) {
    const which = Number(style) & 15;
    const sets = {
      0: [['OK', 1]],
      1: [['OK', 1], ['Cancel', 2]],
      2: [['Abort', 3], ['Retry', 4], ['Ignore', 5]],
      3: [['Yes', 6], ['No', 7], ['Cancel', 2]],
      4: [['Yes', 6], ['No', 7]],
      5: [['Retry', 4], ['Cancel', 2]],
    };
    return (sets[which] || sets[0]).map(function (pair) { return { label: pair[0], value: pair[1] }; });
  }

  function messageIcon(style) {
    const which = Number(style) & 240;
    if (which === 16) return 'error';
    if (which === 32) return 'question';
    if (which === 48) return 'warn';
    if (which === 64) return 'info';
    return 'info';
  }

  function install(host) {
    host._vbFormList = host._vbFormList || [];
    host.createForm = function (model, instance, vm) {
      const form = new Form(model, instance, vm);
      host._vbFormList.push(form);
      return form;
    };
    host.beginModal = function (form) {
      host._vbModal = host._vbModal || [];
      host._vbModal.push(form);
      const update = function () {
        const top = host._vbModal[host._vbModal.length - 1];
        for (let i = 0; i < host._vbFormList.length; i++) {
          const item = host._vbFormList[i];
          if (item.setInert) item.setInert(!!top && item !== top);
        }
      };
      update();
      return function () {
        host._vbModal = host._vbModal.filter(function (item) { return item !== form; });
        update();
      };
    };
    host.msgBox = function (text, style, title) {
      const buttons = buttonSet(style);
      const def = Math.min(buttons.length - 1, (Number(style) >>> 8) & 3);
      const cancel = buttons.filter(function (button) { return button.value === 2; })[0] || ((Number(style) & 15) === 0 ? buttons[0] : null);
      const ask = L.ui && L.ui.msg;
      if (!ask) return Promise.resolve(buttons[def].value);
      return Promise.resolve(ask(String(text == null ? '' : text), {
        title: title || (L.APP || 'VibeOffice'),
        icon: messageIcon(style),
        buttons: buttons.map(function (button) { return button.label; }),
        def: def,
      })).then(function (index) {
        if (index == null) return cancel ? cancel.value : buttons[0].value;
        const button = buttons[index];
        return button ? button.value : buttons[0].value;
      });
    };
    host.inputBox = function (text, title, def) {
      const ask = L.ui && L.ui.prompt;
      if (!ask) return Promise.resolve(def == null ? '' : String(def));
      return Promise.resolve(ask(String(text == null ? '' : text), def == null ? '' : String(def), title || (L.APP || 'VibeOffice'))).then(function (value) {
        return value == null ? '' : String(value);
      });
    };
  }

  function remember(host) {
    return { byHandle: host.byHandle, byObject: host.byObject, auto: host.auto };
  }

  function useMaps(host, maps) {
    host.byHandle = maps.byHandle;
    host.byObject = maps.byObject;
    host.auto = maps.auto;
  }

  function running(host) {
    const form = host.formMachine;
    return !!(form && form.state !== 'stopped' && form.state !== 'error');
  }

  function runningNow() {
    const host = L.vb && L.vb.current && L.vb.current();
    return !!(host && running(host));
  }

  function loadedCount(machine) {
    let n = 0;
    if (machine && machine.formInstances) machine.formInstances.forEach(function (inst) { if (inst.loaded) n++; });
    return n;
  }

  function refreshChrome() {
    if (typeof document === 'undefined') return;
    try {
      if (L.app && L.app.updateToolbars) L.app.updateToolbars();
      if (L.ui && L.ui.refresh) L.ui.refresh();
    } catch (e) { /* the page is not built */ }
  }

  /* Unload of the last loaded form ends the form machine, the way closing the
     startup form ends a VB6 program. Hide alone does not. */
  function watchUnload(host, machine) {
    const unload = machine.unloadForm.bind(machine);
    machine.unloadForm = async function (instance, mode) {
      const unloaded = await unload(instance, mode);
      if (unloaded && host.formMachine === machine && loadedCount(machine) === 0) stop(host);
      return unloaded;
    };
  }

  function stop(host) {
    const machine = host.formMachine;
    host.formMachine = null;
    const list = host._vbFormList || [];
    host._vbFormList = [];
    for (let i = 0; i < list.length; i++) { try { list[i].Hide(); } catch (e) { /* already hidden */ } }
    if (machine && machine.stop && machine.state !== 'stopped') {
      try { machine.stop(); } catch (e) { /* already stopped */ }
    }
    if (host.immediateMaps) {
      useMaps(host, host.immediateMaps);
      host.immediateMaps = null;
      host.formMaps = null;
    }
    host.createForm = null;
    host.beginModal = null;
    host.msgBox = null;
    host.inputBox = null;
    refreshChrome();
  }

  function stopCurrent() {
    const host = L.vb && L.vb.current && L.vb.current();
    if (host && running(host)) stop(host);
  }

  /* A script, a macro or the Visual Basic IDE. true: the form stays, do not continue. */
  async function block(host) {
    if (!running(host)) return false;
    const ask = L.ui && L.ui.msg;
    if (!ask) throw new Error('A VB form is running');
    const index = await Promise.resolve(ask('A VB form is running, do you want to stop it?', {
      icon: 'question',
      buttons: ['&Yes', '&No'],
    }));
    if (index !== 0) return true;
    stop(host);
    return false;
  }

  function evalForm(host, text) {
    const form = host.formMachine;
    if (!form || form.state === 'stopped') return null;
    if (form.state === 'error') { stop(host); return null; }
    if (form.debugEvaluation || (form.stack && form.stack.length && form.state !== 'paused')) {
      return Promise.reject(new Error('A VB script is running'));
    }
    return form.immediate(String(text));
  }

  function parseValue(raw) {
    const text = String(raw == null ? '' : raw).trim();
    if (text.charAt(0) === '"') {
      let out = '';
      for (let i = 1; i < text.length; i++) {
        if (text.charAt(i) === '"') {
          if (text.charAt(i + 1) === '"') { out += '"'; i++; continue; }
          break;
        }
        out += text.charAt(i);
      }
      return out;
    }
    if (/^true$/i.test(text)) return -1;
    if (/^false$/i.test(text)) return 0;
    const hex = text.match(/^&H([0-9A-F]+)(&?)$/i);
    if (hex) {
      let n = parseInt(hex[1], 16);
      if (n > 0x7FFFFFFF) n -= 0x100000000;
      return n;
    }
    if (/^-?\d+(\.\d+)?$/.test(text)) return Number(text);
    return text;
  }

  function parseBegin(line) {
    const property = line.match(/^\s*BeginProperty\s+(\S+)/i);
    if (property) return { kind: 'property', name: property[1] };
    const begin = line.match(/^\s*Begin\s+(\S+)\s+(\S+)/);
    if (!begin) return null;
    const type = begin[1].split('.').pop();
    if (/^property$/i.test(type)) return { kind: 'property', name: begin[2] };
    return { kind: 'control', type: type, name: begin[2] };
  }

  function applyFont(parent, block) {
    if (!parent || !parent.properties) return;
    const props = block.props || {};
    if (props.Name != null) parent.properties.FontName = props.Name;
    if (props.Size != null) parent.properties.FontSize = props.Size;
    if (props.Weight != null) parent.properties.FontBold = Number(props.Weight) >= 700 ? -1 : 0;
    if (props.Italic != null) parent.properties.FontItalic = on(props.Italic) ? -1 : 0;
    if (props.Underline != null) parent.properties.FontUnderline = on(props.Underline) ? -1 : 0;
  }

  function flatten(node, parentName, out) {
    out = out || [];
    const children = node.controls || [];
    for (let i = 0; i < children.length; i++) {
      const child = children[i];
      out.push({ name: child.name, type: child.type, parent: parentName || null, properties: child.properties });
      flatten(child, child.name, out);
    }
    return out;
  }

  function flattenMenus(node, parentName, out) {
    out = out || [];
    const children = node.menus || [];
    for (let i = 0; i < children.length; i++) {
      const child = children[i];
      out.push({ name: child.name, type: 'Menu', parent: parentName || null, properties: child.properties });
      flattenMenus(child, child.name, out);
    }
    return out;
  }

  function parseFrm(text) {
    const lines = String(text || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
    const stack = [];
    let root = null;
    let i = 0;
    for (; i < lines.length; i++) {
      const line = lines[i];
      const begin = parseBegin(line);
      if (begin) {
        if (begin.kind === 'property') { stack.push({ kind: 'property', name: begin.name, props: {} }); continue; }
        const node = { kind: begin.type === 'Menu' ? 'menu' : 'control', type: begin.type, name: begin.name, properties: {}, controls: [], menus: [] };
        stack.push(node);
        if (!root && (begin.type === 'Form' || begin.type === 'MDIForm')) root = node;
        continue;
      }
      if (/^\s*End\s*$/i.test(line) || /^\s*EndProperty\b/i.test(line)) {
        const node = stack.pop();
        if (!node) break;
        if (node.kind === 'property') { applyFont(stack[stack.length - 1], node); continue; }
        if (node === root) { i++; break; }
        const parent = stack[stack.length - 1];
        if (!parent || parent.kind === 'property') continue;
        if (node.kind === 'menu') parent.menus.push(node);
        else parent.controls.push(node);
        continue;
      }
      const assign = line.match(/^\s*([A-Za-z_][\w]*)\s*=\s*(.*)$/);
      if (assign && stack.length) {
        const top = stack[stack.length - 1];
        const bag = top.kind === 'property' ? top.props : top.properties;
        if (bag) bag[assign[1]] = parseValue(assign[2]);
      }
    }
    if (!root) fail('This file has no Visual Basic form', 1002);
    const rest = lines.slice(i);
    let name = root.name;
    let codeAt = 0;
    while (codeAt < rest.length && (/^\s*$/.test(rest[codeAt]) || /^\s*Attribute\s/i.test(rest[codeAt]))) {
      const named = rest[codeAt].match(/^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"/i);
      if (named) name = named[1];
      codeAt++;
    }
    return {
      name: name || 'Form1',
      code: rest.slice(codeAt).join('\n'),
      form: { type: root.type, name: name || root.name, properties: root.properties, controls: flatten(root), menus: flattenMenus(root) },
    };
  }

  function parseBas(text, filename) {
    const lines = String(text || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
    let name = String(filename || 'Module1').replace(/^.*[\\/]/, '').replace(/\.[^.]+$/, '').replace(/[^\w]/g, '') || 'Module1';
    let i = 0;
    while (i < lines.length && (/^\s*$/.test(lines[i]) || /^\s*Attribute\s/i.test(lines[i]))) {
      const named = lines[i].match(/^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"/i);
      if (named) name = named[1];
      i++;
    }
    if (!/^[A-Za-z_][\w]*$/.test(name)) name = 'Module1';
    return { name: name, code: lines.slice(i).join('\n') };
  }

  function filesToProject(files) {
    const modules = [];
    const losses = [];
    let startup = '';
    const list = files || [];
    for (let i = 0; i < list.length; i++) {
      const file = list[i] || {};
      const filename = String(file.name || '');
      const ext = (filename.split('.').pop() || '').toLowerCase();
      const text = String(file.text || '');
      if (ext === 'frx') {
        losses.push({ id: 'vb-form:frx:' + filename, what: 'The form picture was left out.', where: filename, action: 'drop', notify: false });
        continue;
      }
      if (ext === 'vbp') {
        const named = text.match(/^\s*Startup\s*=\s*"([^"]+)"/im);
        if (named) startup = named[1];
        continue;
      }
      if (ext === 'frm') {
        const parsed = parseFrm(text);
        modules.push({ id: parsed.name, name: parsed.name, kind: 'form', code: parsed.code, form: parsed.form });
        continue;
      }
      if (ext === 'bas') {
        const parsed = parseBas(text, filename);
        modules.push({ id: parsed.name, name: parsed.name, kind: 'module', code: parsed.code });
      }
    }
    if (!startup) {
      const form = modules.filter(function (module) { return module.kind === 'form'; })[0];
      startup = form ? form.name : '';
    }
    return {
      losses: losses,
      project: { schema: 1, id: 'vb-form', name: startup || 'Form', startup: startup, modules: modules, settings: {} },
    };
  }

  function formDomModule(module) {
    const copy = Object.assign({}, module);
    copy.code = String(module.code || '').replace(/Public Function (?:MsgBox|InputBox)\([\s\S]*?\nEnd Function\r?\n?/g, '');
    return copy;
  }

  function asProject(input, losses) {
    if (input && Array.isArray(input.files)) {
      const built = filesToProject(input.files);
      for (let i = 0; i < built.losses.length; i++) losses.push(built.losses[i]);
      return built.project;
    }
    if (input && Array.isArray(input.modules)) {
      const settings = Object.assign({}, input.settings || {});
      delete settings.anchoring;
      delete settings.xaml;
      return {
        schema: 1,
        id: input.id || 'vb-form',
        name: input.name || 'Form',
        startup: input.startup || '',
        modules: input.modules,
        settings: settings,
      };
    }
    throw new Error('A form project needs modules or files');
  }

  async function load(host, project, hooks) {
    stop(host);
    const losses = [];
    const spec = asProject(project, losses);
    host.formLosses = (host.formLosses || []).concat(losses);
    /* The document module publishes MsgBox and InputBox for macros. A form's
       MsgBox is the virtual machine's own, which calls host.msgBox. */
    const modules = [formDomModule(hooks.domModule())];
    const names = new Set([String(modules[0].name).toLowerCase()]);
    for (let i = 0; i < spec.modules.length; i++) {
      const module = spec.modules[i];
      const key = String(module.name || '').toLowerCase();
      if (!key || names.has(key)) throw new Error('Duplicate module: ' + module.name);
      names.add(key);
      modules.push(module);
    }
    const engineProject = {
      schema: 1,
      id: hooks.id,
      name: hooks.name,
      startup: spec.startup || '',
      modules: modules,
      settings: spec.settings || {},
    };
    const program = host.engine.compileProject(engineProject);
    if (!program.valid) {
      const err = new Error((program.diagnostics || []).map(function (item) {
        return item.source + ':' + item.line + ': ' + item.message;
      }).join('\n') || 'The form project did not compile');
      err.diagnostics = program.diagnostics;
      err.number = 1002;
      throw err;
    }
    install(host);
    const saved = remember(host);
    host.immediateMaps = saved;
    host.byHandle = new Map();
    host.byObject = new WeakMap();
    host.auto = null;
    host.formMaps = remember(host);
    let machine = null;
    try {
      machine = new host.engine.VirtualMachine(program, host.machineHost(host.registry()), { debuggerEnabled: false, instructionLimit: 5000000 });
      watchUnload(host, machine);
      const initialize = machine.initialize.bind(machine);
      machine.initialize = async function () {
        await initialize();
        await hooks.bind(machine);
      };
      host.formMachine = machine;
      await machine.start();
      refreshChrome();
      return machine;
    } catch (error) {
      stop(host);
      throw error;
    }
  }

  function pick(session) {
    if (typeof document === 'undefined') return Promise.reject(new Error('Run Form needs a page'));
    const input = document.createElement('input');
    input.type = 'file';
    input.multiple = true;
    input.accept = '.frm,.bas,.frx,.vbp,text/plain';
    const chosen = new Promise(function (resolve) {
      input.addEventListener('change', function () {
        const files = input.files ? Array.prototype.slice.call(input.files) : [];
        Promise.all(files.map(function (file) {
          return file.text().then(function (text) { return { name: file.name, text: text }; });
        })).then(resolve, function () { resolve([]); });
      });
      input.addEventListener('cancel', function () { resolve([]); });
    });
    input.click();
    return chosen.then(async function (files) {
      if (!files.length) return;
      if (!files.some(function (file) { return /\.frm$/i.test(file.name); })) {
        const say = L.t ? L.t('Choose a Visual Basic form (.frm).') : 'Choose a Visual Basic form (.frm).';
        if (L.ui && L.ui.msg) await L.ui.msg(say, { icon: 'info' });
        return;
      }
      const host = await session();
      try {
        await host.loadProject({ files: files });
      } catch (error) {
        const say = L.t ? L.t('That form could not be run: {0}', error && error.message ? error.message : error) : String(error && error.message ? error.message : error);
        if (L.ui && L.ui.msg) await L.ui.msg(say, { icon: 'warn' });
        else throw error;
      }
    });
  }

  VO.vbForms = {
    install: install,
    load: load,
    stop: stop,
    stopCurrent: stopCurrent,
    block: block,
    eval: evalForm,
    running: running,
    runningNow: runningNow,
    pick: pick,
    parseFrm: parseFrm,
    parseBas: parseBas,
    filesToProject: filesToProject,
    element: function (control) { return control && (control._face || nodes.get(control)) || null; },
    create: function (model, instance, vm) { return new Form(model, instance, vm); },
  };
})();
