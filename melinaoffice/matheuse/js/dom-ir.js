/* Matheuse DOM IR. Inventory: docs/matheuse_dom_api_specifications.md
   Protocol: docs/matheuse_dom_ir_specifications.md */
(function () {
  'use strict';
  const L = window.L;
  const dom = L.dom;
  const ir = (L.domIr = L.domIr || {});
  const Missing = Object.freeze({ missing: true });
  ir.Missing = Missing;

  const CATALOG = `
# Application
ActiveWorkbook get Workbook
ThisWorkbook get Workbook
ActiveSheet get Worksheet
ActiveCell get Range
Selection get Range
Workbooks get Workbooks
Worksheets get Worksheets
Sheets get Sheets
Names get Names
WorksheetFunction get WorksheetFunction
ScreenUpdating rw
DisplayAlerts rw
Calculation rw
ReferenceStyle rw
Calculate call
Undo call
Redo call
Windows gap G-HOST get
AddIns gap G-HOST get
Dialogs get BuiltinDialogs
FileDialog call BuiltinFileDialog ([Type])
CommandBars gap G-HOST get
Run gap G-HOST call ([Macro], [Arg1], [Arg2], [Arg3], [Arg4], [Arg5], [Arg6], [Arg7], [Arg8], [Arg9], [Arg10], [Arg11], [Arg12], [Arg13], [Arg14], [Arg15], [Arg16], [Arg17], [Arg18], [Arg19], [Arg20], [Arg21], [Arg22], [Arg23], [Arg24], [Arg25], [Arg26], [Arg27], [Arg28], [Arg29], [Arg30])
SendKeys gap G-HOST call ([Keys], [Wait])
Wait gap G-HOST call ([Time])
OnTime gap G-HOST call ([EarliestTime], [Procedure], [LatestTime], [Schedule])
GetOpenFilename call RequestApproval FileDialog dialog ([FileFilter], [FilterIndex], [Title], [ButtonText], [MultiSelect])
GetSaveAsFilename call RequestApproval FileDialog dialog ([InitialFilename], [FileFilter], [FilterIndex], [Title], [ButtonText])
HasHostAbility call ([Ability])
MsgBox call AlwaysAllow ([Prompt], [Buttons], [Title])
InputBox call AlwaysAllow ([Prompt], [Title], [Default])
HostShell call AlwaysDeny Shell ([PathName])
HostDir call RequestApproval Dir ([PathName])
EnableEvents gap G-ENGINE get
# Workbooks collection
Item call Workbook
Add call Workbook ([Template])
Open call RequestApproval Open ([Filename], [UpdateLinks], [ReadOnly], [Format], [Password])
Close gap G-HOST call
# Workbook
Name get
FullName get AlwaysAllow Path
Path get AlwaysAllow Path
Saved rw
Worksheets get Worksheets
Sheets get Sheets
Charts get Sheets
ActiveSheet get Worksheet
Names get Names
Styles get Styles
BuiltInDocumentProperties get DocumentProperties
CustomDocumentProperties get DocumentProperties
Activate call
Save call RequestApproval Save
Close call ([SaveChanges], [Filename], [RouteWorkbook])
Protect call ([Password])
Unprotect call ([Password])
SaveAs call RequestApproval SaveAs ([Filename], [FileFormat], [Password], [WriteResPassword])
SaveAs2 gap G-HOST call ([Filename], [FileFormat], [Password], [WriteResPassword])
PrintOut call AlwaysAllow Print ([From], [To], [Copies], [Preview], [ActivePrinter], [PrintToFile], [Collate], [PrToFileName], [IgnorePrintAreas])
ExportAsFixedFormat call RequestApproval FixedFormat ([Type], [Filename])
FollowHyperlink call AlwaysAllow Hyperlink ([Address], [SubAddress], [NewWindow], [AddHistory], [ExtraInfo], [Method], [HeaderInfo])
SendMail gap G-HOST call ([Recipients], [Subject], [ReturnReceipt])
Windows gap G-HOST get
# DocumentProperties collection
Item call DocumentProperty
Add call DocumentProperty ([Name], [LinkToContent], [Type], [Value])
# DocumentProperty
Name get
Value rw
# Names collection
Item call Name
Add call Name ([Name], [RefersTo])
# Name
Name rw
RefersTo rw
RefersToR1C1 get
Delete call
# Styles collection
Item call Style
# Style
Name get
# Sheets collection
Item call Worksheet
Add call Worksheet ([Before], [After], [Count])
# Worksheets collection
Item call Worksheet
Add call Worksheet ([Before], [After], [Count])
# Worksheet
Name rw
Index get
Visible rw
UsedRange get Range
Range call Range ([Cell1], [Cell2])
Cells both Range ([RowIndex], [ColumnIndex])
Rows get Range
Columns get Range
Comments get Comments
Hyperlinks get Hyperlinks
Names get Names
ListObjects get ListObjects
ChartObjects get ChartObjects
Shapes get Shapes
PivotTables get PivotTables
PageSetup get PageSetup
HPageBreaks get PageBreaks
VPageBreaks get PageBreaks
Outline get Outline
ProtectionMode get
AutoFilter get AutoFilter
Scenarios get Scenarios
Activate call
Select call
Delete call
Copy call Worksheet ([Before], [After])
Move call ([Before], [After])
Calculate call
Protect call ([Password])
Unprotect call ([Password])
QueryTables gap G-ENGINE get
XmlMaps gap G-ENGINE get
OLEObjects gap G-ENGINE get
CommentsThreaded gap G-ENGINE get
PrintOut call AlwaysAllow Print ([From], [To], [Copies], [Preview], [ActivePrinter], [PrintToFile], [Collate], [PrToFileName], [IgnorePrintAreas])
ExportAsFixedFormat call RequestApproval FixedFormat ([Type], [Filename])
SaveAs call RequestApproval SheetSaveAs ([Filename], [FileFormat], [Password], [WriteResPassword], [ReadOnlyRecommended], [CreateBackup])
# Range
Count get
Row get
Column get
Rows get Range
Columns get Range
Worksheet get Worksheet
Address get
Formula rw
FormulaR1C1 rw
Value rw
Value2 rw
Text get
NumberFormat rw
Font get Font
Interior get Interior
Borders get Borders
ColumnWidth rw
RowHeight rw
Hidden rw
MergeCells get
Item call Range ([RowIndex], [ColumnIndex])
Cells both Range ([RowIndex], [ColumnIndex])
Resize call Range ([RowSize], [ColumnSize])
Offset call Range ([RowOffset], [ColumnOffset])
EntireRow get Range
EntireColumn get Range
End call Range ([Direction])
CurrentRegion get Range
SpecialCells call Range ([Type], [Value])
Clear call
ClearContents call
ClearFormats call
ClearComments call
Merge call
UnMerge call
Select call
Activate call
Copy call ([Destination])
PasteSpecial call ([Paste])
Insert call ([Shift])
Delete call ([Shift])
Find call Range ([What], [After], [LookIn], [LookAt], [SearchOrder], [SearchDirection], [MatchCase])
Replace call ([What], [Replacement], [LookAt], [SearchOrder], [MatchCase])
Sort call ([Key1], [Order1], [Key2], [Type], [Order2], [Key3], [Order3], [Header])
AutoFilter call ([Field], [Criteria1])
AddComment call Comment ([Text])
Comment get Comment
Hyperlinks get Hyperlinks
Validation get Validation
FormatConditions get FormatConditions
Areas get Areas
SortFields gap G-ENGINE get
# Areas collection
Item call Range
# Font
Name rw
Size rw
Bold rw
Italic rw
Underline rw
Color rw
ColorIndex rw
ThemeColor rw
TintAndShade rw
# Interior
Color rw
Pattern rw
ColorIndex rw
ThemeColor rw
TintAndShade rw
PatternColor rw
# Borders
Item call Border ([Index])
LineStyle rw
# Border
LineStyle rw
Weight rw
Color rw
# Comments collection
Item call Comment
Add call Comment ([Range], [Text])
# Comment
Text rw
Author get
Delete call
# Hyperlinks collection
Item call Hyperlink
Add call Hyperlink ([Anchor], [Address], [SubAddress])
Delete call
# Hyperlink
Address rw
SubAddress rw
Range get Range
TextToDisplay rw
Delete call
Follow gap G-HOST call
# Validation
Add call ([Type], [AlertStyle], [Operator], [Formula1])
Delete call
Modify gap G-ENGINE call ([Type], [AlertStyle], [Operator], [Formula1], [Formula2])
# FormatConditions collection
Item call FormatCondition
Add call FormatCondition ([Type], [Operator], [Formula1], [Formula2])
Delete call
# FormatCondition
Type get
Formula1 get
Delete call
# ListObjects collection
Item call ListObject
Add call ListObject ([SourceType], [Source], [LinkSource], [XlListObjectHasHeaders], [Destination])
# ListObject
Name rw
Range get Range
ListColumns get ListColumns
Delete call
Resize call ([Range])
Refresh gap G-ENGINE call
# ListColumns collection
Item call ListColumn
# ListColumn
Name rw
# ChartObjects collection
Item call ChartObject
Add call ChartObject ([Left], [Top], [Width], [Height])
# ChartObject
Name rw
Left rw
Top rw
Width rw
Height rw
Chart get Chart
Delete call
# Chart
ChartType rw
Name get
SeriesCollection gap G-ENGINE get
Axes gap G-ENGINE get
SetSourceData gap G-ENGINE call ([Source], [PlotBy])
# Shapes collection
Item call Shape
AddTextbox call Shape ([Orientation], [Left], [Top], [Width], [Height])
AddPicture call RequestApproval Picture ([Filename], [LinkToFile], [SaveWithDocument], [Left], [Top], [Width], [Height])
# Shape
Name rw
Left rw
Top rw
Width rw
Height rw
TextFrame get TextFrame
Delete call
OLEFormat gap G-ENGINE get
# TextFrame
TextRange get TextRange
# TextRange
Text rw
# PivotTables collection
Item call PivotTable
# PivotTable
Name get
Refresh gap G-ENGINE call
PivotFields gap G-ENGINE get
# PageSetup
Orientation rw
Zoom rw
FitToPagesWide rw
FitToPagesTall rw
LeftMargin rw
RightMargin rw
TopMargin rw
BottomMargin rw
PrintArea rw
CenterHorizontally rw
CenterVertically rw
PrintOut call AlwaysAllow Print ([From], [To], [Copies], [Preview], [ActivePrinter], [PrintToFile], [Collate], [PrToFileName], [IgnorePrintAreas])
# PageBreaks collection
Item call PageBreak
Add call ([Before])
# PageBreak
Location get Range
Delete call
# Outline
SummaryRow get
ShowLevels call ([RowLevels], [ColumnLevels])
# AutoFilter
Range get Range
ShowAllData call
# Scenarios collection
Item call Scenario
Add call Scenario ([Name], [ChangingCells], [Values])
# Scenario
Name rw
Comment rw
Show call
Delete call
ChangeScenario gap G-ENGINE call ([ChangingCells], [Values])
# WorksheetFunction
# QueryTables gap G-ENGINE
# QueryTable gap G-ENGINE
# Connections gap G-ENGINE
# WorkbookConnection gap G-ENGINE
# XmlMaps gap G-ENGINE
# XmlMap gap G-ENGINE
# Slicers gap G-ENGINE
# SlicerCaches gap G-ENGINE
# Slicer gap G-ENGINE
# SparklineGroups gap G-ENGINE
# SparklineGroup gap G-ENGINE
# OLEObjects gap G-ENGINE
# OLEObject gap G-ENGINE
# CommentsThreaded gap G-ENGINE
# CommentThreaded gap G-ENGINE
# Phonetics gap G-ENGINE
# Phonetic gap G-ENGINE
# Model gap G-ENGINE
# Queries gap G-ENGINE
# Windows gap G-HOST
# Window gap G-HOST
# AddIns gap G-HOST
# AddIn gap G-HOST
# Dialogs gap G-HOST
# BuiltinDialogs
Item call BuiltinDialog ([Index])
# BuiltinDialog
Show call AlwaysAllow
# BuiltinFileDialog
Show call RequestApproval FileDialog dialog
SelectedItems get
# Dialog gap G-HOST
# FileDialog gap G-HOST
# CommandBars gap G-HOST
`;

  function takeAudit(parts) {
    const AUDIT = { AlwaysAllow: 1, RequestApproval: 1, AlwaysDeny: 1 };
    const out = parts.slice();
    let fileDialog = false;
    if (out[out.length - 1] === 'dialog') { fileDialog = true; out.pop(); }
    let audit = '';
    let ability = '';
    const last = out[out.length - 1];
    const prev = out[out.length - 2];
    if (prev && AUDIT[prev]) { audit = prev; ability = last; out.splice(out.length - 2, 2); }
    else if (last && AUDIT[last]) { audit = last; out.pop(); }
    return { parts: out, audit: audit, ability: ability, fileDialog: fileDialog };
  }

  function takeSignature(line) {
    const open = line.indexOf('(');
    if (open < 0) return { line: line, params: null };
    const close = line.indexOf(')', open);
    if (close < 0) throw new Error('Unclosed signature: ' + line);
    const body = line.slice(open + 1, close).trim();
    const params = [];
    if (body) {
      for (const bit of body.split(',')) {
        const tok = bit.trim();
        const opt = /^\[([A-Za-z][A-Za-z0-9]*)\]$/.exec(tok);
        const bare = /^[A-Za-z][A-Za-z0-9]*$/.exec(tok);
        const name = opt ? opt[1] : (bare ? bare[0] : '');
        if (!name) throw new Error('Bad signature name: ' + line);
        params.push({ name: name, optional: !!opt });
      }
    }
    return { line: (line.slice(0, open) + line.slice(close + 1)).replace(/\s+/g, ' ').trim(), params: params };
  }

  function parseCatalog(text) {
    const types = {};
    let cur = null;
    for (const raw of text.split('\n')) {
      const line = raw.trim();
      if (!line || line.charAt(0) === '/') continue;
      if (line.charAt(0) === '#') {
        const bits = line.slice(1).trim().split(/\s+/);
        cur = { name: bits[0], members: Object.create(null), collection: bits.indexOf('collection') >= 0, gap: null };
        const gi = bits.indexOf('gap');
        if (gi >= 0) cur.gap = bits[gi + 1];
        types[cur.name] = cur;
        continue;
      }
      const signed = takeSignature(line);
      const taken = takeAudit(signed.line.split(/\s+/));
      const parts = taken.parts;
      const rec = { kind: parts[1], audit: taken.audit, ability: taken.ability, fileDialog: taken.fileDialog };
      if (signed.params) rec.params = signed.params;
      if (parts[1] === 'gap') {
        rec.code = parts[2];
        rec.as = parts[3] || 'any';
      } else {
        rec.ret = parts[2] && parts[2] !== 'live' ? parts[2] : (parts[1] === 'call' || parts[1] === 'both' ? 'empty' : 'scalar');
        const live = parts.indexOf('live');
        if (live >= 0) rec.live = parts[live + 1];
      }
      cur.members[parts[0]] = rec;
    }
    for (const name in types) {
      const t = types[name];
      if (!t.collection) continue;
      const item = t.members.Item;
      const ret = item && item.ret && item.ret !== 'empty' ? item.ret : 'Object';
      if (!t.members.Count) t.members.Count = { kind: 'get', ret: 'scalar' };
      if (!t.members.First) t.members.First = { kind: 'get', ret: ret };
      if (!t.members.Last) t.members.Last = { kind: 'get', ret: ret };
      if (!t.members.Item) t.members.Item = { kind: 'call', ret: ret };
      if (t.members.Item.kind === 'call' && !t.members.Item.params) t.members.Item.params = [{ name: 'Index', optional: true }];
    }
    return types;
  }

  const types = parseCatalog(CATALOG);

  function fnKnown(member) {
    if (!L.fninfo || !L.fninfo.get) return false;
    if (L.fninfo.get(member)) return true;
    return !!L.fninfo.get(String(member).replace(/_/g, '.'));
  }

  function describe(type, member) {
    if (member === 'Creator') return { kind: 'gap', code: 'G-ENGINE', as: 'get' };
    if (member === 'Application') return { kind: 'get', ret: 'Application' };
    if (member === 'Parent') return { kind: 'get', ret: 'parent' };
    if (type === 'WorksheetFunction' && member !== 'Application' && member !== 'Parent' && member !== 'Creator') {
      if (fnKnown(member)) return { kind: 'call', ret: 'scalar' };
      return null;
    }
    const t = types[type];
    if (!t) return null;
    if (t.members[member]) return t.members[member];
    if (t.gap && member !== 'Application' && member !== 'Parent' && member !== 'Creator') return { kind: 'gap', code: t.gap, as: 'any' };
    return null;
  }

  ir.types = types;
  ir.member = describe;

  function fail(code, member, message) {
    return { ok: false, code: code, member: member || '', message: message || (member ? member + ' is not available (' + code + ')' : code) };
  }

  function auditCall(desc, member, produce) {
    if (desc && desc.audit === 'AlwaysDeny') {
      if (!L.hostAudit || !L.hostAudit.has(desc.ability || '')) return fail('G-HOST', member, member + ' is not available (G-HOST)');
      return fail('E-DENIED', member, 'The editor refused this operation.');
    }
    if (!desc || desc.audit !== 'RequestApproval' || desc.fileDialog || member === 'ExportAsFixedFormat' || !L.hostAudit) return produce();
    return Promise.resolve(L.hostAudit.gate({
      tier: 'RequestApproval', ability: desc.ability || member, member: member,
    })).then(function (r) {
      if (!r || r.ok === false) return fail((r && r.code) || 'E-APPROVAL', member, (r && r.message) || 'The editor did not allow this operation.');
      try { return produce(); }
      catch (e) { return fail((e && e.code) || 'error', member, e && e.message); }
    });
  }

  function hasMember(obj, name) {
    if (obj == null || (typeof obj !== 'object' && typeof obj !== 'function')) return false;
    if (obj._wf && fnKnown(name)) return true;
    return Object.prototype.hasOwnProperty.call(obj, name) || Object.getOwnPropertyDescriptor(obj, name) != null;
  }

  function unwrap(session, value) {
    if (value === Missing) return undefined;
    if (Array.isArray(value)) return value.map((v) => unwrap(session, v));
    if (value && typeof value === 'object' && Object.prototype.hasOwnProperty.call(value, 'handle')) {
      const rec = session.byId.get(value.handle);
      if (!rec) {
        const err = new Error('Unknown handle');
        err.code = 'handle';
        throw err;
      }
      return rec.obj;
    }
    return value;
  }

  function bindArgs(session, args) {
    const src = args || [];
    let end = src.length;
    while (end > 0 && src[end - 1] === Missing) end--;
    const out = [];
    for (let i = 0; i < end; i++) out.push(unwrap(session, src[i]));
    return out;
  }

  function settle(session, value, ret) {
    if (value && typeof value.then === 'function') {
      return Promise.resolve(value).then(
        (v) => present(session, v, ret),
        (e) => fail((e && e.code) || 'error', '', e && e.message)
      );
    }
    return present(session, value, ret);
  }

  function present(session, value, ret) {
    if (ret === 'empty') {
      if (value == null || typeof value === 'object' || typeof value === 'function') {
        if (Array.isArray(value)) return { ok: true, kind: 'array', value: value };
        return { ok: true, kind: 'empty' };
      }
      return { ok: true, kind: 'scalar', value: value };
    }
    if (Array.isArray(value)) return { ok: true, kind: 'array', value: value };
    if (ret === 'scalar') return { ok: true, kind: 'scalar', value: value };
    if (value == null) return { ok: true, kind: 'nothing' };
    if (typeof value !== 'object' && typeof value !== 'function') return { ok: true, kind: 'scalar', value: value };
    const rec = session.intern(value, ret === 'parent' ? null : ret);
    return { ok: true, kind: 'handle', handle: rec.handle, type: rec.type };
  }

  function liveName(desc, member) {
    return desc.live || member;
  }

  function gapCovers(desc, op) {
    if (!desc || desc.kind !== 'gap') return false;
    if (desc.as === 'any') return true;
    if (op === 'set') return desc.as === 'rw';
    if (op === 'call' || op === 'item') return desc.as === 'call' || desc.as === 'any';
    return desc.as === 'get' || desc.as === 'rw';
  }

  ir.open = function open() {
    const byId = new Map();
    const byObj = new WeakMap();
    let next = 0;
    let released = false;
    const session = { byId: byId };
    session.intern = function (obj, type) {
      const hit = byObj.get(obj);
      if (hit) return hit;
      const rec = { handle: ++next, type: type || 'Object', obj: obj };
      byId.set(rec.handle, rec);
      byObj.set(obj, rec);
      return rec;
    };
    const root = session.intern(dom.Application, 'Application');
    session.root = { handle: root.handle, type: 'Application' };
    session.named = function (name) {
      if (released) return fail('released', name);
      if (name === 'Application') return { ok: true, kind: 'handle', handle: root.handle, type: 'Application' };
      const t = types[name];
      if (!t || !t.gap) return fail('absent', name, name + ' is not a registered gap object');
      const obj = {};
      obj.Application = dom.Application;
      obj.Parent = dom.Application;
      const rec = session.intern(obj, name);
      return { ok: true, kind: 'handle', handle: rec.handle, type: name };
    };
    session.exec = function (req) {
      if (released) return fail('released', req && req.member);
      const op = req && req.op;
      const rec = byId.get(req && req.handle);
      if (!rec) return fail('handle', req && req.member, 'Unknown handle');
      try {
        if (op === 'item') {
          const desc = describe(rec.type, 'Item');
          if (!desc) return fail('absent', 'Item');
          if (gapCovers(desc, 'call')) return fail(desc.code, 'Item');
          const index = unwrap(session, req.args && req.args[0]);
          if (typeof rec.obj.Item !== 'function') return fail('missing-live', 'Item');
          return settle(session, rec.obj.Item(index), desc.ret || 'Object');
        }
        if (op !== 'get' && op !== 'set' && op !== 'call') return fail('op', '', 'Unknown operation');
        const member = req.member;
        const desc = describe(rec.type, member);
        if (!desc) return fail('absent', member);
        if (gapCovers(desc, op)) return fail(desc.code, member);
        const name = liveName(desc, member);
        if (rec.type === 'WorksheetFunction' && op === 'call' && rec.obj.invoke) {
          return settle(session, rec.obj.invoke(member, bindArgs(session, req.args)), 'scalar');
        }
        if (op === 'get') {
          if (desc.kind === 'call') return fail('absent', member);
          if (!hasMember(rec.obj, name)) return fail('missing-live', member);
          if (desc.kind === 'both' && typeof rec.obj[name] === 'function') return settle(session, rec.obj[name](), desc.ret);
          return settle(session, rec.obj[name], desc.ret);
        }
        if (op === 'set') {
          if (desc.kind !== 'rw') return fail('readonly', member);
          if (!hasMember(rec.obj, name)) return fail('missing-live', member);
          rec.obj[name] = unwrap(session, req.value);
          return { ok: true, kind: 'empty' };
        }
        const args = bindArgs(session, req.args);
        if (typeof rec.obj[name] === 'function') return auditCall(desc, member, function () {
          return settle(session, rec.obj[name].apply(rec.obj, args), desc.ret || 'empty');
        });
        if (!args.length && hasMember(rec.obj, name)) return settle(session, rec.obj[name], desc.kind === 'call' ? (desc.ret === 'empty' ? 'scalar' : desc.ret) : desc.ret);
        return fail('missing-live', member);
      } catch (e) {
        if (e && e.code === 'handle') return fail('handle', req.member, e.message);
        return fail((e && e.code) || 'error', req.member, e && e.message);
      }
    };
    session.release = function () {
      released = true;
      byId.clear();
    };
    return session;
  };
})();
