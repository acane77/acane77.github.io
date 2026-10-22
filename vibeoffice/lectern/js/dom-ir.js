/* Lectern DOM IR. Inventory: docs/lectern_dom_api_specifications.md
   Protocol: docs/lectern_dom_ir_specifications.md */
(function () {
  'use strict';
  const L = window.L;
  const dom = L.dom;
  const ir = (L.domIr = L.domIr || {});
  const Missing = Object.freeze({ missing: true });
  ir.Missing = Missing;

  const CATALOG = `
# Application
ActivePresentation get Presentation
Presentations get Presentations
ActiveWindow get DocumentWindow
SlideShowWindows get SlideShowWindows
HasHostAbility call
MsgBox call AlwaysAllow
InputBox call AlwaysAllow
HostShell call AlwaysDeny Shell
HostDir call RequestApproval Dir
FileDialog call BuiltinFileDialog
GetOpenFilename call RequestApproval FileDialog dialog
GetSaveAsFilename call RequestApproval FileDialog dialog
# Presentations collection
Item call Presentation
Add call Presentation
Open call RequestApproval Open
# Presentation
Name get
FullName get AlwaysAllow Path
Path get AlwaysAllow Path
Saved rw
Slides get Slides
PageSetup get PageSetup
SlideShowSettings get SlideShowSettings
SlideShowWindows get SlideShowWindows
Designs get Designs
SlideMaster get SlideMaster
BuiltInDocumentProperties get DocumentProperties
CustomDocumentProperties get DocumentProperties
Undo call
Redo call
Save call RequestApproval Save
Close call
SaveAs call RequestApproval SaveAs
SaveAs2 call RequestApproval SaveAs
Export call RequestApproval Export
ExportAsFixedFormat call RequestApproval FixedFormat
PrintOut call AlwaysAllow Print
FollowHyperlink call AlwaysAllow Hyperlink
NewWindow gap G-HOST call
CreateVideo gap G-ENGINE call
Merge gap G-ENGINE call
PublishSlides gap G-ENGINE call
# DocumentWindow
Selection get Selection
Presentation get Presentation
View gap G-HOST get
ViewType gap G-HOST get
Caption gap G-HOST get
ActivePane gap G-HOST get
Panes gap G-HOST get
Split gap G-HOST get
# Selection
Type get
ShapeRange get ShapeRange
SlideRange get SlideRange
TextRange get TextRange
Unselect call
Delete call
Copy call
Cut call
Paste call
ChildShapeRange gap G-ENGINE get
# Slides collection
Item call Slide
Add call Slide
Range call SlideRange
# SlideRange collection
Item call Slide
# Slide
Name rw
SlideIndex get
SlideNumber get
SlideID gap G-ENGINE get
Layout rw
Shapes get Shapes
Design rw Design
HeadersFooters get HeadersFooters
SlideShowTransition get SlideShowTransition
Hyperlinks get Hyperlinks
Select call
Delete call
Duplicate call Slide
MoveTo call
NotesPage gap G-ENGINE get
Comments gap G-ENGINE get
TimeLine gap G-ENGINE get
Tags gap G-ENGINE get
CustomerData gap G-ENGINE get
# Shapes collection
Item call Shape
Title get Shape
Placeholders get Shapes
AddShape call Shape
AddTextbox call Shape
AddTable call Shape
AddPicture call RequestApproval Picture
AddChart gap G-ENGINE call
AddSmartArt gap G-ENGINE call
AddDiagram gap G-ENGINE call
AddMediaObject gap G-ENGINE call
AddOLEObject gap G-ENGINE call
AddTextEffect gap G-ENGINE call
# ShapeRange collection
Item call Shape
Left rw
Top rw
Delete call
# Shape
Left rw
Top rw
Width rw
Height rw
Rotation rw
Name rw
Id get
Type get
HasTextFrame get
AutoShapeType rw
TextFrame get TextFrame
TextFrame2 get TextFrame
Fill get FillFormat
Line get LineFormat
PlaceholderFormat get PlaceholderFormat
Table get Table
HasTable get
GroupItems get Shapes
Delete call
Duplicate call Shape
Select call
ZOrder call
ZOrderPosition get
Hyperlinks get Hyperlinks
Chart gap G-ENGINE get
OLEFormat gap G-ENGINE get
MediaFormat gap G-ENGINE get
ActionSettings gap G-ENGINE get
AnimationSettings gap G-ENGINE get
Nodes gap G-ENGINE get
Tags gap G-ENGINE get
# TextFrame
TextRange get TextRange
HasText get
WordWrap rw
AutoSize rw
MarginLeft rw
MarginTop rw
MarginRight rw
MarginBottom rw
VerticalAnchor rw
Orientation gap G-ENGINE get
# TextRange
Text rw
InsertBefore call
InsertAfter call
Delete call
Select call
Font get Font
ParagraphFormat get ParagraphFormat
IndentLevel rw
Paragraphs get Paragraphs
Characters get Characters
InsertSymbol call
InsertDateTime call
InsertSlideNumber call
Hyperlinks get Hyperlinks
Start gap G-TEXT get
Length gap G-TEXT get
BoundLeft gap G-TEXT get
BoundTop gap G-TEXT get
BoundWidth gap G-TEXT get
BoundHeight gap G-TEXT get
Find gap G-TEXT call
Replace gap G-TEXT call
Words gap G-TEXT get
Sentences gap G-TEXT get
Lines gap G-TEXT get
Runs gap G-TEXT get
# Paragraphs collection
Item call Paragraph
# Paragraph
Range get TextRange
Alignment rw
Font get Font
ParagraphFormat get ParagraphFormat
IndentLevel rw
# Font
Name rw
NameFarEast rw
Size rw
Bold rw
Italic rw
Underline rw
Color get ColorFormat
NameAscii gap G-ENGINE get
NameComplexScript gap G-ENGINE get
SizeBi gap G-ENGINE get
BoldBi gap G-ENGINE get
ItalicBi gap G-ENGINE get
# ParagraphFormat
Alignment rw
SpaceBefore rw
SpaceAfter rw
LineRuleWithin rw
SpaceWithin rw
Bullet get BulletFormat
# BulletFormat
Type rw
Visible rw
# ColorFormat
RGB rw
ObjectThemeColor rw
SchemeColor rw
TintAndShade gap G-ENGINE get
Brightness gap G-ENGINE get
# FillFormat
Type rw
ForeColor get ColorFormat
Transparency rw
Visible rw
# LineFormat
Weight rw
ForeColor get ColorFormat
Visible rw
DashStyle gap G-ENGINE get
# PlaceholderFormat
Type get
ContainedType get
# Table
Cell call Cell
Rows get Rows
Columns get Columns
Delete call
Merge gap G-ENGINE call
ApplyStyle gap G-ENGINE call
# Rows collection
Item call Row
Add call Row
# Row
Height rw
Cells get Cells
Delete call
# Cells collection
Item call Cell
# Columns collection
Item call Column
Add call Column
# Column
Width rw
Delete call
# Cell
Shape get Shape
# Hyperlinks collection
Item call Hyperlink
Add call Hyperlink
# Hyperlink
Address rw
SubAddress rw
ScreenTip rw
TextToDisplay rw
Range get TextRange
Delete call
Follow gap G-HOST call
Type gap G-ENGINE get
# HeadersFooters
DateAndTime get HeaderFooter
Footer get HeaderFooter
SlideNumber get HeaderFooter
DisplayOnTitleSlide rw
# HeaderFooter
Visible rw
Text rw
Format gap G-ENGINE get
UseFormat gap G-ENGINE get
# PageSetup
SlideWidth rw
SlideHeight rw
FirstSlideNumber rw
SlideSize gap G-ENGINE get
NotesOrientation gap G-ENGINE get
SlideOrientation gap G-ENGINE get
# Designs collection
Item call Design
# Design
Name get
# DocumentProperties collection
Item call DocumentProperty
Add call DocumentProperty
# DocumentProperty
Name get
Value rw
# SlideShowSettings
StartingSlide rw
EndingSlide rw
RangeType rw
LoopUntilStopped rw
AdvanceMode rw
ShowType rw
Run call
NamedSlideShows gap G-ENGINE get
PointerColor gap G-ENGINE get
# SlideShowWindows collection
Item call SlideShowWindow
# SlideShowWindow
View get SlideShowView
# SlideShowView
Next call
Previous call
GotoSlide call
Exit call
CurrentShowPosition get
State gap G-HOST get
PointerType gap G-HOST get
AcceleratorsEnabled gap G-HOST get
# SlideShowTransition
EntryEffect rw
Speed rw
AdvanceOnClick rw
AdvanceOnTime rw
AdvanceTime rw
Hidden rw
Duration gap G-ENGINE get
SoundEffect gap G-ENGINE get
# SlideMaster
Name get
HeadersFooters get HeadersFooters
Shapes gap G-ENGINE get
CustomLayouts gap G-ENGINE get
Design gap G-ENGINE get
Background gap G-ENGINE get
# Characters collection
Item gap G-TEXT call
Count get
# TimeLine gap G-ENGINE
# Sequence gap G-ENGINE
# Effect gap G-ENGINE
# Chart gap G-ENGINE
# ChartData gap G-ENGINE
# Diagram gap G-ENGINE
# DiagramNode gap G-ENGINE
# SmartArt gap G-ENGINE
# OLEFormat gap G-ENGINE
# MediaFormat gap G-ENGINE
# CustomerData gap G-ENGINE
# NotesPage gap G-ENGINE
# Comments gap G-ENGINE
# Comment gap G-ENGINE
# Tags gap G-ENGINE
# PrintOptions gap G-ENGINE
# Panes gap G-HOST
# Pane gap G-HOST
# View gap G-HOST
# CommandBars gap G-HOST
# FileDialog gap G-HOST
# BuiltinFileDialog
Show call RequestApproval FileDialog dialog
SelectedItems get
# Dialogs gap G-HOST
# ActionSettings gap G-ENGINE
# NamedSlideShows gap G-ENGINE
# SectionProperties gap G-ENGINE
# ColorSchemes gap G-ENGINE
# AnimationSettings gap G-ENGINE
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
      const taken = takeAudit(line.split(/\s+/));
      const parts = taken.parts;
      const rec = { kind: parts[1], audit: taken.audit, ability: taken.ability, fileDialog: taken.fileDialog };
      if (parts[1] === 'gap') {
        rec.code = parts[2];
        rec.as = parts[3] || 'any';
      } else {
        rec.ret = parts[2] && parts[2] !== 'live' ? parts[2] : (parts[1] === 'call' ? 'empty' : 'scalar');
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
    }
    return types;
  }

  const types = parseCatalog(CATALOG);

  function describe(type, member) {
    if (member === 'Creator') return { kind: 'gap', code: 'G-ENGINE', as: 'get' };
    if (member === 'Application') return { kind: 'get', ret: 'Application' };
    if (member === 'Parent') return { kind: 'get', ret: 'parent' };
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
    return Object.prototype.hasOwnProperty.call(obj, name) || Object.getOwnPropertyDescriptor(obj, name) != null;
  }

  function unwrap(session, value) {
    if (value === Missing) return undefined;
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
      if (value == null || typeof value === 'object' || typeof value === 'function') return { ok: true, kind: 'empty' };
      return { ok: true, kind: 'scalar', value: value };
    }
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
        if (op === 'get') {
          if (desc.kind === 'call') return fail('absent', member);
          if (!hasMember(rec.obj, name)) return fail('missing-live', member);
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
