/* Ecrivaine DOM IR. Inventory: docs/ecrivaine_dom_api_specifications.md
   Protocol: docs/ecrivaine_dom_ir_specifications.md */
(function () {
  'use strict';
  const L = window.L;
  const dom = L.dom;
  const ir = (L.domIr = L.domIr || {});
  const Missing = Object.freeze({ missing: true });
  ir.Missing = Missing;

  const CATALOG = `
# Application
ActiveDocument get Document
Documents get Documents
Selection get Selection
HasHostAbility call ([Ability])
MsgBox call AlwaysAllow ([Prompt], [Buttons], [Title])
InputBox call AlwaysAllow ([Prompt], [Title], [Default])
HostShell call AlwaysDeny Shell ([PathName])
HostDir call RequestApproval Dir ([PathName])
Dialogs get BuiltinDialogs
FileDialog call BuiltinFileDialog ([Type])
GetOpenFilename call RequestApproval FileDialog dialog ([FileFilter], [FilterIndex], [Title], [ButtonText], [MultiSelect])
GetSaveAsFilename call RequestApproval FileDialog dialog ([InitialFilename], [FileFilter], [FilterIndex], [Title], [ButtonText])
# Documents collection
Item call Document
Add call Document ([Template], [NewTemplate], [DocumentType], [Visible])
Open call RequestApproval Open ([FileName], [ConfirmConversions], [ReadOnly], [AddToRecentFiles], [PasswordDocument])
Close gap G-HOST call ([SaveChanges], [OriginalFormat], [RouteDocument])
Save gap G-HOST call ([NoPrompt], [OriginalFormat])
# Document
Name get
FullName get AlwaysAllow Path
Path get AlwaysAllow Path
Saved rw
Content get Range
Paragraphs get Paragraphs
Sentences get Sentences
Words get Words
Characters get Characters
Tables get Tables
Sections get Sections
StoryRanges get StoryRanges
Styles get Styles
Bookmarks get Bookmarks
Fields get Fields
Hyperlinks get Hyperlinks
Comments get Comments
Footnotes get Footnotes
Endnotes get Endnotes
InlineShapes get InlineShapes
Shapes get Shapes
Lists get Lists
ListParagraphs get ListParagraphs
Revisions get Revisions
BuiltInDocumentProperties get DocumentProperties
CustomDocumentProperties get DocumentProperties
Undo call ([Times])
Redo call ([Times])
UndoClear gap G-ENGINE call
Range call Range ([Start], [End])
Select call
Activate call
Close call ([SaveChanges], [OriginalFormat], [RouteDocument])
Save call RequestApproval Save
SaveAs call RequestApproval SaveAs ([FileName], [FileFormat], [LockComments], [Password])
SaveAs2 call RequestApproval SaveAs ([FileName], [FileFormat], [LockComments], [Password])
ComputeStatistics call ([Statistic], [IncludeFootnotesAndEndnotes])
PrintOut call AlwaysAllow Print ([Background], [Append], [Range], [OutputFileName], [From], [To], [Item], [Copies], [Pages], [PageType], [PrintToFile], [Collate], [FileName], [ActivePrinterMacGX], [ManualDuplexPrint], [PrintZoomColumn], [PrintZoomRow], [PrintZoomPaperWidth], [PrintZoomPaperHeight])
PrintPreview gap G-HOST call
Protect gap G-ENGINE call ([Type], [NoReset], [Password], [UseIRM], [EnforceStyleLock])
Unprotect gap G-ENGINE call ([Password])
Compare call RequestApproval Open ([Name], [AuthorName], [CompareTarget], [DetectFormatChanges], [IgnoreAllComparisonWarnings], [AddToRecentFiles], [RemovePersonalInformation], [RemoveDateAndTime])
FollowHyperlink call AlwaysAllow Hyperlink ([Address], [SubAddress], [NewWindow], [AddHistory], [ExtraInfo], [Method], [HeaderInfo])
GoTo gap G-STORY call ([What], [Which], [Count], [Name])
ExportAsFixedFormat call RequestApproval FixedFormat ([OutputFileFormat], [OutputFileName])
TrackRevisions rw
# Selection
Text rw
TypeText call ([Text])
TypeParagraph call
TypeBackspace call
Delete call ([Unit], [Count])
InsertBefore call ([Text])
InsertAfter call ([Text])
Font get Font
ParagraphFormat get ParagraphFormat
Style rw
Range get Range
Collapse call ([Direction])
Select call
WholeStory call
Copy call
Cut call
Paste call
Find get Find
Bold rw
Italic rw
Underline rw
Paragraphs get Paragraphs
Sentences get Sentences
Words get Words
Tables get Tables
Bookmarks get Bookmarks
Fields get Fields
Comments get Comments
Hyperlinks get Hyperlinks
InlineShapes get InlineShapes
ListFormat get ListFormat
Case rw
LanguageID rw
LanguageIDFarEast rw
FormattedText gap G-ENGINE get
Document get Document
InsertParagraph call
InsertParagraphBefore call
InsertParagraphAfter call
Type get
Start rw
End rw
SetRange call ([Start], [End])
Move call ([Unit], [Count])
MoveStart call ([Unit], [Count])
MoveEnd call ([Unit], [Count])
MoveLeft call ([Unit], [Count], [Extend])
MoveRight call ([Unit], [Count], [Extend])
MoveUp call ([Unit], [Count], [Extend])
MoveDown call ([Unit], [Count], [Extend])
HomeKey call ([Unit], [Extend])
EndKey call ([Unit], [Extend])
StartOf call ([Unit], [Extend])
EndOf call ([Unit], [Extend])
Expand call ([Unit])
Borders get Borders
Shading get Shading
Shapes get Shapes
ShapeRange get Shapes
Cells get Cells
Rows get Rows
Columns get Columns
Sections get Sections
PageSetup get PageSetup
StoryType get
StoryLength get
InRange call ([Range])
InStory call ([Range])
BookmarkID get
StartIsActive get
NoProofing rw
SelectCell call
SelectRow call
SelectColumn call
InsertRowsAbove call ([NumRows])
InsertRowsBelow call ([NumRows])
InsertColumnsRight call ([NumColumns])
MoveWhile gap G-STORY call ([Cset], [Count])
MoveUntil gap G-STORY call ([Cset], [Count])
IPAtEndOfLine gap G-STORY get
IsEndOfRowMark gap G-STORY get
ChildShapeRange gap G-ENGINE get
HasChildShapeRange gap G-ENGINE get
Editors gap G-ENGINE get
FormFields gap G-ENGINE get
Frames gap G-ENGINE get
HTMLDivisions gap G-ENGINE get
WordOpenXML gap G-ENGINE get
XML gap G-ENGINE get
EndnoteOptions gap G-ENGINE get
FootnoteOptions gap G-ENGINE get
FitTextWidth gap G-ENGINE get
Flags gap G-ENGINE get
ExtendMode gap G-ENGINE get
ColumnSelectMode gap G-ENGINE get
EnhMetaFileBits gap G-ENGINE get
HeaderFooter gap G-ENGINE get
TopLevelTables gap G-ENGINE get
LanguageIDOther gap G-ENGINE get
InsertFormula gap G-ENGINE call
InsertCrossReference gap G-ENGINE call ([ReferenceType], [ReferenceKind], [ReferenceItem], [InsertAsHyperlink], [IncludePosition], [SeparateNumbers], [SeparatorString])
InsertXML gap G-ENGINE call
InsertStyleSeparator gap G-ENGINE call
PasteAsNestedTable gap G-ENGINE call
PasteAppendTable gap G-ENGINE call
PasteFormat gap G-ENGINE call
CopyAsPicture gap G-ENGINE call
ToggleCharacterCode gap G-ENGINE call
DetectLanguage gap G-ENGINE call
Calculate gap G-ENGINE call
Shrink gap G-ENGINE call
EscapeKey gap G-ENGINE call
CreateAutoTextEntry gap G-ENGINE call
NextRevision gap G-ENGINE call
NextField gap G-ENGINE call
PreviousField gap G-ENGINE call
InsertNewPage gap G-ENGINE call
SortAscending gap G-ENGINE call
SortDescending gap G-ENGINE call
SplitTable gap G-ENGINE call
BoldRun gap G-ENGINE call
ItalicRun gap G-ENGINE call
ClearFormatting gap G-ENGINE call
CreateTextbox gap G-ENGINE call
GoTo gap G-STORY call ([What], [Which], [Count], [Name])
GoToNext gap G-STORY call ([What])
GoToPrevious gap G-STORY call ([What])
Information gap G-STORY call ([Type])
InsertBreak call ([Type])
InsertSymbol call ([CharacterNumber], [Font], [Unicode], [Bias])
InsertDateTime call ([DateTimeFormat], [InsertAsField], [InsertAsFullWidth], [DateLanguage], [CalendarType])
InsertFile call AlwaysAllow InsertFile ([FileName], [Range], [ConfirmConversions], [Link], [Attachment])
InsertCaption gap G-ENGINE call ([Label], [Title], [TitleAutoText], [Position], [ExcludeLabel])
InsertCells call
InsertRows call ([NumRows])
InsertColumns call ([NumColumns])
ConvertToTable call Table ([Separator], [NumRows], [NumColumns], [InitialColumnWidth], [Format], [ApplyBorders], [ApplyShading], [ApplyFont], [ApplyColor], [ApplyHeadingRows], [ApplyLastRow], [ApplyFirstColumn], [ApplyLastColumn], [AutoFit], [AutoFitBehavior], [DefaultTableBehavior])
PasteAndFormat gap G-ENGINE call ([Type])
PasteSpecial gap G-ENGINE call ([IconIndex], [Link], [Placement], [DisplayAsIcon], [DataType], [IconFileName], [IconLabel])
PasteExcelTable gap G-ENGINE call ([LinkedToExcel], [WordFormatting], [RTF])
Characters get Characters
# Range
Text rw
InsertBefore call ([Text])
InsertAfter call ([Text])
InsertParagraph call
InsertParagraphBefore call
InsertParagraphAfter call
Delete call ([Unit], [Count])
Select call
Collapse call ([Direction])
Duplicate call Range
Copy call
Cut call
Paste call
InsertFile call AlwaysAllow InsertFile ([FileName], [Range], [ConfirmConversions], [Link], [Attachment])
Font get Font
ParagraphFormat get ParagraphFormat
Style rw
Bold rw
Italic rw
Underline rw
Paragraphs get Paragraphs
Sentences get Sentences
Words get Words
Tables get Tables
Find get Find
Characters get Characters
Case rw
ListFormat get ListFormat
Hyperlinks get Hyperlinks
Bookmarks get Bookmarks
Fields get Fields
Comments get Comments
InlineShapes get InlineShapes
Shapes get Shapes
FormattedText gap G-ENGINE get
LanguageID rw
LanguageIDFarEast rw
Start rw
End rw
SetRange call ([Start], [End])
Borders get Borders
Shading get Shading
Move call ([Unit], [Count])
MoveStart call ([Unit], [Count])
MoveEnd call ([Unit], [Count])
MoveWhile gap G-STORY call ([Cset], [Count])
MoveUntil gap G-STORY call ([Cset], [Count])
MoveStartWhile gap G-STORY call ([Cset], [Count])
MoveStartUntil gap G-STORY call ([Cset], [Count])
MoveEndWhile gap G-STORY call ([Cset], [Count])
MoveEndUntil gap G-STORY call ([Cset], [Count])
StartOf call ([Unit], [Extend])
EndOf call ([Unit], [Extend])
Expand call ([Unit])
Information gap G-STORY call ([Type])
GoTo gap G-STORY call ([What], [Which], [Count], [Name])
GoToNext gap G-STORY call ([What])
GoToPrevious gap G-STORY call ([What])
Next gap G-STORY call ([Unit], [Count])
Previous gap G-STORY call ([Unit], [Count])
# Paragraphs collection
Item call Paragraph
# Paragraph
Range get Range
Style rw
Format get ParagraphFormat
LeftIndent rw
RightIndent rw
FirstLineIndent rw
SpaceBefore rw
SpaceAfter rw
LineSpacing rw
Alignment rw
OutlineLevel rw
KeepTogether rw
KeepWithNext rw
PageBreakBefore rw
WidowControl rw
TabStops get TabStops
Next call Paragraph ([Count])
Previous call Paragraph ([Count])
IndentCharWidth call
Space1 call
Space15 call
Space2 call
OutlineDemote gap G-ENGINE call
OutlinePromote gap G-ENGINE call
OutlineDemoteToBody gap G-ENGINE call
DropCap get DropCap
# DropCap
Position rw
Lines rw
DistanceFromText rw
# Sentences collection
Item call Range
# Words collection
Item call Range
# Tables collection
Item call Table
Add call Table ([Range], [NumRows], [NumColumns], [DefaultTableBehavior], [AutoFitBehavior])
# Table
Cell call Cell ([Row], [Column])
Rows get Rows
Columns get Columns
Range get Range
Style rw
Select call
Delete call
Split call ([BeforeRow])
AutoFitBehavior call ([Behavior])
Sort call ([ExcludeHeader], [FieldNumber], [SortFieldType], [SortOrder])
ConvertToText call ([Separator], [NestedTables])
ApplyStyleHeadingRows rw
ApplyStyleLastRow rw
ApplyStyleFirstColumn rw
ApplyStyleLastColumn rw
ApplyStyleRowBands rw
ApplyStyleColumnBands rw
Borders get Borders
Shading get Shading
Uniform get
Spacing rw
AllowAutoFit rw
PreferredWidth rw
PreferredWidthType rw
TopPadding rw
BottomPadding rw
LeftPadding rw
RightPadding rw
Alignment rw
TableDirection rw
Title rw
Descr rw
NestingLevel get
Indent get
# Rows collection
Item call Row
Add call Row ([BeforeRow])
Delete call
# Row
Delete call
Select call
Range get Range
Cells get Cells
HeadingFormat rw
Height rw
HeightRule rw
AllowBreakAcrossPages rw
Alignment rw
# Columns collection
Item call Column
Add call Column ([BeforeColumn])
# Column
Delete call
Select call
Width rw
Cells get Cells
# Cells collection
Item call Cell
Merge call
Split call ([NumRows], [NumColumns], [MergeBeforeSplit])
DistributeWidth call
DistributeHeight call
# Cell
Range get Range
RowIndex get
ColumnIndex get
Width rw
Height rw
Borders get Borders
Shading get Shading
VerticalAlignment rw
TopPadding rw
BottomPadding rw
LeftPadding rw
RightPadding rw
FitText rw
WordWrap rw
# Borders collection
Item call Border
Enable rw
OutsideLineStyle rw
OutsideLineWidth rw
OutsideColor rw
InsideLineStyle rw
InsideLineWidth rw
InsideColor rw
Shadow rw
DistanceFromTop rw
DistanceFromLeft rw
DistanceFromBottom rw
DistanceFromRight rw
DistanceFrom rw
AlwaysInFront rw
EnableFirstPageInSection rw
EnableOtherPagesInSection rw
HasHorizontal get
HasVertical get
Count get
ApplyPageBordersToAllSections gap G-ENGINE call
JoinBorders gap G-ENGINE get
SurroundHeader gap G-ENGINE get
SurroundFooter gap G-ENGINE get
ArtStyle gap G-ENGINE get
ArtWidth gap G-ENGINE get
# Border
LineStyle rw
LineWidth rw
Color rw
ColorIndex rw
Visible rw
Inside get
# Shading
BackgroundPatternColor rw
ForegroundPatternColor rw
BackgroundPatternColorIndex rw
ForegroundPatternColorIndex rw
Texture rw
# Sections collection
Item call Section
# Section
Range get Range
PageSetup get PageSetup
Headers call HeaderFooter ([Index])
Footers call HeaderFooter ([Index])
Borders get Borders
ProtectedForForms rw
# PageSetup
PageWidth rw
PageHeight rw
Orientation rw
TopMargin rw
BottomMargin rw
LeftMargin rw
RightMargin rw
Gutter rw
HeaderDistance rw
FooterDistance rw
TextColumns get TextColumns
PaperSize rw
VerticalAlignment rw
# TextColumns
Count rw
Spacing rw
EvenlySpaced rw
# HeaderFooter
Range get Range
Exists get
LinkToPrevious gap G-ENGINE get
# StoryRanges collection
Item call Range
# Styles collection
Item call Style
Add call Style ([Name], [Type])
# Style
NameLocal rw
Type get
Font get Font
ParagraphFormat get ParagraphFormat
Delete call
InUse get
BaseStyle rw
NextParagraphStyle rw
LinkStyle rw
Visibility rw
UnhideWhenUsed gap G-ENGINE get
QuickStyle rw
Priority rw
Locked gap G-ENGINE get
# Font
Name rw
NameAscii rw
NameFarEast rw
NameBi rw
Size rw
SizeBi rw
Bold rw
Italic rw
BoldBi rw
ItalicBi rw
Underline rw
UnderlineColor rw
StrikeThrough rw
DoubleStrikeThrough rw
AllCaps rw
SmallCaps rw
Hidden rw
Superscript rw
Subscript rw
Color rw
HighlightColorIndex rw
Spacing rw
Position rw
Scaling rw
Kerning rw
Outline rw
Shadow rw
Emboss rw
Engrave rw
EmphasisMark rw
Borders get Borders
Shading get Shading
Animation gap G-ENGINE get
DiacriticColor gap G-ENGINE get
StylisticSet gap G-ENGINE get
NumberSpacing gap G-ENGINE get
NumberForm gap G-ENGINE get
Ligatures gap G-ENGINE get
ContextualAlternates gap G-ENGINE get
# ParagraphFormat
Alignment rw
LeftIndent rw
RightIndent rw
FirstLineIndent rw
CharacterUnitLeftIndent gap G-ENGINE rw
CharacterUnitRightIndent gap G-ENGINE rw
CharacterUnitFirstLineIndent gap G-ENGINE rw
SpaceBefore rw
SpaceAfter rw
LineSpacing rw
LineSpacingRule rw
SpaceBeforeAuto rw
SpaceAfterAuto rw
KeepTogether rw
KeepWithNext rw
PageBreakBefore rw
WidowControl rw
OutlineLevel rw
Hyphenation rw
TabStops get TabStops
Borders get Borders
Shading get Shading
ReadingOrder rw
DisableLineHeightGrid rw
MirrorIndents rw
AutoAdjustRightIndent gap G-ENGINE get
FarEastLineBreakControl gap G-ENGINE get
WordWrap gap G-ENGINE get
HangingPunctuation gap G-ENGINE get
HalfWidthPunctuationOnTopOfLine gap G-ENGINE get
AddSpaceBetweenFarEastAndAlpha gap G-ENGINE get
AddSpaceBetweenFarEastAndDigit gap G-ENGINE get
# TabStops collection
Item call TabStop
Add call TabStop ([Position], [Alignment], [Leader])
ClearAll call
# TabStop
Position rw
Alignment rw
Leader rw
Clear call
# ListFormat
ListType get
ListLevelNumber rw
ApplyBulletDefault call
ApplyNumberDefault call
ApplyListTemplate gap G-ENGINE call ([ListTemplate], [ContinuePreviousList], [ApplyTo], [DefaultListBehavior])
RemoveNumbers call ([NumberType])
ListIndent call
ListOutdent call
CanContinuePreviousList gap G-ENGINE get
# Lists collection
Item call List
ListTemplates gap G-ENGINE get
ListTemplate gap G-ENGINE get
ListLevels gap G-ENGINE get
ListLevel gap G-ENGINE get
# List
Range get Range
ListParagraphs get ListParagraphs
CountNumberedItems call ([NumberType], [Level])
# ListParagraphs collection
Item call Paragraph
# Find
Execute call ([FindText], [MatchCase], [MatchWholeWord], [MatchWildcards], [MatchSoundsLike], [MatchAllWordForms], [Forward], [Wrap], [Format], [ReplaceWith], [Replace], [MatchKashida], [MatchDiacritics], [MatchAlefHamza], [MatchControl])
Found get
Text rw
MatchCase rw
MatchWholeWord rw
MatchWildcards rw
MatchSoundsLike rw
MatchAllWordForms rw
Forward rw
Wrap rw
Format rw
Replacement get Replacement
ClearFormatting call
HitHighlight gap G-ENGINE call ([FindText], [HighlightColor], [TextColor], [MatchCase], [MatchWholeWord], [MatchWildcards], [MatchSoundsLike], [MatchAllWordForms], [Forward], [Wrap], [Format], [MatchKashida], [MatchDiacritics], [MatchAlefHamza], [MatchControl], [IgnoreSpace], [IgnorePunct], [HighlightColorIndex])
ClearHitHighlight gap G-ENGINE call
MatchPrefix gap G-ENGINE get
MatchSuffix gap G-ENGINE get
MatchPhrase gap G-ENGINE get
MatchByte gap G-ENGINE get
IgnoreSpace gap G-ENGINE get
IgnorePunct gap G-ENGINE get
# Replacement
Text rw
Font gap G-ENGINE get
ParagraphFormat gap G-ENGINE get
# Bookmarks collection
Item call Bookmark
Exists call ([Name])
Add call Bookmark ([Name], [Range])
# Bookmark
Name get
Range get Range
Delete call
Select call
Empty get
Start rw
End rw
Copy gap G-ENGINE call ([Name])
# Fields collection
Item call Field
Add call Field ([Range], [Type], [Text], [PreserveFormatting])
Update call
ToggleShowCodes gap G-ENGINE call
# Field
Result get Range
Code get Range
Update call
Unlink call
Delete call
Locked rw
Type get
Select call
Copy gap G-ENGINE call
DoClick call AlwaysAllow Hyperlink
# Hyperlinks collection
Item call Hyperlink
Add call Hyperlink ([Anchor], [Address], [SubAddress], [ScreenTip], [TextToDisplay], [Target])
# Hyperlink
Address rw
SubAddress rw
TextToDisplay rw
ScreenTip rw
Range get Range
Delete call
Follow call AlwaysAllow Hyperlink ([NewWindow], [AddHistory], [ExtraInfo], [Method], [HeaderInfo])
Type gap G-ENGINE get
Target gap G-ENGINE get
# Comments collection
Item call Comment
Add call Comment ([Range], [Text])
AddReply call ([Comment], [Text])
# Comment
Range get Range
Scope get Range
Author rw
Initial rw live Initials
Delete call
Edit gap G-ENGINE call
# Footnotes collection
Item call Footnote
Add call Footnote ([Range], [Reference], [Text])
Convert call
# Endnotes collection
Item call Endnote
Add call Endnote ([Range], [Reference], [Text])
Convert call
# Footnote
Range get Range
Index get
Reference get Range
# Endnote
Range get Range
Index get
Reference get Range
# InlineShapes collection
Item call InlineShape
AddPicture call InlineShape RequestApproval Picture ([FileName], [LinkToFile], [SaveWithDocument], [Range])
# InlineShape
Width rw
Height rw
ScaleWidth rw
ScaleHeight rw
Type get
Select call
Delete call
Range get Range
AlternativeText rw
Title rw
Left rw
Top rw
WrapFormat get WrapFormat
ConvertToShape call Shape
Hyperlink get Hyperlink
PictureFormat get PictureFormat
# Shapes collection
Item call Shape
AddShape call Shape ([Type], [Left], [Top], [Width], [Height], [Anchor])
AddPicture call Shape RequestApproval Picture ([FileName], [LinkToFile], [SaveWithDocument], [Left], [Top], [Width], [Height], [Anchor])
AddPicture2 call Shape RequestApproval Picture ([FileName], [LinkToFile], [SaveWithDocument], [Left], [Top], [Width], [Height], [Anchor], [Compress])
AddTextbox call Shape ([Orientation], [Left], [Top], [Width], [Height], [Anchor])
AddChart gap G-ENGINE call ([Type], [Left], [Top], [Width], [Height], [Anchor])
# Shape
Left rw
Top rw
Width rw
Height rw
WrapFormat get WrapFormat
ConvertToInlineShape call InlineShape
TextFrame get TextFrame
Select call
Delete call
Fill get FillFormat
Line get LineFormat
OLEFormat gap G-ENGINE get
Chart gap G-ENGINE get
GroupItems gap G-ENGINE get
Nodes gap G-ENGINE get
# WrapFormat
Type rw
# TextFrame
TextRange get Range
# FillFormat
ForeColor get ColorFormat
# LineFormat
ForeColor get ColorFormat
Weight rw
# ColorFormat
RGB rw
# PictureFormat
CropLeft rw
CropTop rw
CropRight rw
CropBottom rw
Brightness rw
Contrast rw
# Revisions collection
Item call Revision
AcceptAll call
RejectAll call
# Revision
Type get
Author get
Date get
Range get Range
Accept call
Reject call
# DocumentProperties collection
Item call DocumentProperty
Add call DocumentProperty ([Name], [LinkToContent], [Type], [Value])
# DocumentProperty
Name get
Value rw
# Characters collection
Item call Range ([Index])
# ListTemplates gap G-ENGINE
# ListTemplate gap G-ENGINE
# ListLevels gap G-ENGINE
# ListLevel gap G-ENGINE
# ContentControls gap G-ENGINE
# ContentControl gap G-ENGINE
# XMLNodes gap G-ENGINE
# XMLNode gap G-ENGINE
# CustomXMLParts gap G-ENGINE
# CustomXMLPart gap G-ENGINE
# FormFields gap G-ENGINE
# FormField gap G-ENGINE
# CheckBox gap G-ENGINE
# DropDown gap G-ENGINE
# TextInput gap G-ENGINE
# Frames gap G-ENGINE
# Frame gap G-ENGINE
# Subdocuments gap G-ENGINE
# MailMerge gap G-ENGINE
# MailMergeFields gap G-ENGINE
# MailMergeDataSource gap G-ENGINE
# OMaths gap G-ENGINE
# OMath gap G-ENGINE
# BuildingBlocks gap G-ENGINE
# BuildingBlock gap G-ENGINE
# Variables gap G-ENGINE
# Variable gap G-ENGINE
# Chart gap G-ENGINE
# ChartData gap G-ENGINE
# Diagram gap G-ENGINE
# DiagramNode gap G-ENGINE
# OLEFormat gap G-ENGINE
# OLEControl gap G-ENGINE
# TablesOfContents gap G-ENGINE
# Indexes gap G-ENGINE
# TablesOfFigures gap G-ENGINE
# Windows gap G-HOST
# Window gap G-HOST
# Panes gap G-HOST
# Pane gap G-HOST
# View gap G-HOST
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
# System gap G-HOST
# Tasks gap G-HOST
# CommandBars gap G-HOST
# KeyBindings gap G-HOST
# Languages gap G-HOST
# AutoCorrect gap G-HOST
# Options gap G-HOST
# DefaultWebOptions gap G-HOST
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
      if (t.members.Item.kind === 'call' && !t.members.Item.params) t.members.Item.params = [{ name: 'Index', optional: true }];
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
