/* Ledger Visual Basic DOM API reference.
   The inventory is docs/ledger_dom_api_specifications.md.
   When a member or a gap changes, update that specification and this file in the same change. */
(function () {
  'use strict';
  const L = window.L;
  const N = L.N;
  const m = (name, kind, note, gap) => {
    const o = { name: name, kind: kind };
    if (note) o.note = note;
    if (gap) o.gap = gap;
    return o;
  };
  const u = (name, code, note) => ({ name: name, code: code, note: note });

  L.vbDomRef.mount({
    spec: 'docs/ledger_dom_api_specifications.md',
    rules: [
      N('The entry is Application. ThisWorkbook in the console is the active workbook.'),
      N('A single cell Value, Value2 or Formula is one value. A block is a rectangle. One cell is never a 1 by 1 array. In VB a block is a 1-based VBArray. A range larger than 1,000,000 cells rejects Value and Formula.'),
      N('A date is an Excel serial number, not a VBA Date. Value and Value2 both return the serial. An error is a string such as #N/A. An empty cell is Empty. A string Value that starts with an equals sign is stored as a formula. A formula written onto a block shifts relative references from the top-left cell.'),
      N('Colour is an Office BGR long. ThemeColor uses the Excel theme enum: 1 dk1, 2 lt1, 3 dk2, 4 lt2, 5 through 10 accent 1 through 6, 11 hyperlink, 12 followed hyperlink.'),
      N('WorksheetFunction calls a name in the Ledger function catalog. A dotted name such as BINOM.DIST is also BINOM_DIST. A name that is not in the catalog is absent.'),
      N('ScreenUpdating and DisplayAlerts are host flags. Calculation maps automatic and manual. Semiautomatic fails. ReferenceStyle maps xlA1 and xlR1C1.'),
      N('Classic FormatConditions, cell is and expression, are implemented. Data bars, colour scales and icon sets stay display-only. Sort is the three-key sort. SortFields fails.'),
      N('Copy and PasteSpecial use an in-app snapshot, not the operating-system clipboard. Save works when the workbook already has a destination. Closing the last workbook opens a blank one. With DisplayAlerts and a dirty book, Close with no SaveChanges fails. SaveChanges True also fails.'),
      N('Macro sheets and dialog sheets appear in Sheets. Editing their cells fails. A chart sheet has no Range.'),
      N('The object model does not list the macro store. Application.Run fails.'),
    ],
    gaps: [
      { code: 'G-ENGINE', note: N('There is no document operation for this member. A cell is a row and a column, so there is no story-index gap and no text-index gap.') },
      { code: 'G-HOST', note: N('The member needs a path, a window, a printer, a modal or a file picker.') },
    ],
    unavailable: [
      u('QueryTables, QueryTable, Connections, WorkbookConnection, XmlMaps, XmlMap', 'G-ENGINE', N('Connections and maps are kept, not executed.')),
      u('Slicers, SlicerCaches, Slicer, SparklineGroups, SparklineGroup', 'G-ENGINE', N('Display and package parts, with no editor.')),
      u('OLEObjects, OLEObject', 'G-ENGINE', N('Embedded objects are kept.')),
      u('CommentsThreaded, CommentThreaded, Phonetics, Phonetic', 'G-ENGINE', N('Not this note model.')),
      u('Model, Queries', 'G-ENGINE', N('Power Pivot and Power Query are not run.')),
      u('Windows, Window, AddIns, AddIn, Dialogs, Dialog, FileDialog, CommandBars', 'G-HOST', N('Windows, add-ins and file pickers.')),
    ],
    objects: [
      {
        name: 'Application',
        members: [
          m('ActiveWorkbook, ThisWorkbook', 'get', N('The console workbook.')),
          m('Workbooks', 'get', N('Open workbooks.')),
          m('ActiveSheet', 'get'),
          m('ActiveCell, Selection', 'get', N('The selection, when the grid workbook matches. Otherwise the sheet stored selection.')),
          m('Names', 'get', N('Workbook names.')),
          m('WorksheetFunction', 'get', N('Catalog functions. Sum, If and VLookup are in the sample set.')),
          m('ScreenUpdating, DisplayAlerts', 'rw', N('Host flags.')),
          m('Calculation', 'rw', N('Automatic or manual. Semiautomatic is G-ENGINE.')),
          m('ReferenceStyle', 'rw', N('xlA1 or xlR1C1.')),
          m('Calculate, Undo, Redo', 'call', N('Recalculate, or one undo or redo step. Returns -1 if a step ran, else 0.')),
          m('Windows, AddIns, Dialogs, FileDialog, CommandBars', 'get', N('Windows, add-ins and file pickers.'), 'G-HOST'),
          m('Run, SendKeys, Wait, OnTime', 'call', N('Those host calls are not on this page.'), 'G-HOST'),
          m('GetOpenFilename, GetSaveAsFilename', 'call', N('On the VS Code host, the editor file dialog. Cancel returns an empty string.')),
          m('EnableEvents', 'get', N('Events are not this flag.'), 'G-ENGINE'),
        ],
      },
      {
        name: 'Workbooks',
        members: [
          m('Item', 'call', N('By name or index.')),
          m('Add', 'call', N('Opens a blank book.')),
          m('Open, Close', 'call', N('Opening and closing the collection need a host.'), 'G-HOST'),
        ],
      },
      {
        name: 'Workbook',
        members: [
          m('Name', 'get'),
          m('FullName, Path', 'get', N('On the VS Code host, the file path and its folder. Elsewhere G-HOST.')),
          m('Saved', 'rw', N('False after an edit until a successful save.')),
          m('Worksheets, Sheets, Charts, ActiveSheet, Names, Styles', 'get', N('Charts is the chart sheets inside Sheets.')),
          m('BuiltInDocumentProperties, CustomDocumentProperties', 'get', N('Title, subject, creator, keywords, description, category, manager, company. Add with LinkToContent, or a type other than string, fails.')),
          m('Activate', 'call', N('Makes the book current.')),
          m('Save', 'call', N('Clears dirty when a destination exists. Otherwise G-HOST.')),
          m('Close', 'call', N('See the rules for the last book and for SaveChanges.')),
          m('Protect, Unprotect', 'call', N('Workbook structure. The password is hashed when the dialog helper is loaded.')),
          m('SaveAs, SaveAs2', 'call', N('On the VS Code host, writes that path. Elsewhere G-HOST.')),
          m('PrintOut', 'call', N('On the VS Code host, a PDF of the visible sheets beside the file.')),
          m('ExportAsFixedFormat', 'call', N('On the VS Code host, a PDF. XPS is an error.')),
          m('FollowHyperlink', 'call', N('On the VS Code host, opens http, https, or mailto.')),
          m('SendMail, Windows', 'call', N('Those calls are not on this page.'), 'G-HOST'),
        ],
      },
      {
        name: 'Names',
        members: [
          m('Item, Add', 'call', N('Add(Name, RefersTo).')),
          m('Name, RefersTo', 'rw'),
          m('RefersToR1C1', 'get'),
          m('Delete', 'call'),
        ],
      },
      {
        name: 'Styles',
        members: [
          m('Item', 'call', N('An existing cell style.')),
          m('Name', 'get'),
        ],
      },
      {
        name: 'Sheets',
        note: N('Sheets, Worksheets and Worksheet share these members. Macro sheets and dialog sheets are listed. Editing their cells fails.'),
        members: [
          m('Item, Add', 'call', N('Add(Before, After, Count) inserts at the Before sheet, or after After.')),
        ],
      },
      {
        name: 'Worksheets',
        members: [
          m('Item, Add', 'call', N('Same Add as Sheets, for worksheets.')),
        ],
      },
      {
        name: 'Worksheet',
        members: [
          m('Name', 'rw'),
          m('Index', 'get', N('1-based.')),
          m('Visible', 'rw', N('xlSheetVisible -1, xlSheetHidden 0, xlSheetVeryHidden 2.')),
          m('UsedRange, Range, Cells, Rows, Columns', 'get', N('Cells with no arguments is the used sheet. Range takes an address or two corners.')),
          m('Comments, Hyperlinks, Names, ListObjects, ChartObjects, Shapes', 'get'),
          m('PivotTables, PageSetup, HPageBreaks, VPageBreaks, Outline, AutoFilter, Scenarios', 'get'),
          m('ProtectionMode', 'get'),
          m('Activate, Select, Delete, Move, Calculate, Protect, Unprotect', 'call'),
          m('Copy', 'call', N('A shallow sheet copy.')),
          m('QueryTables, XmlMaps, OLEObjects, CommentsThreaded', 'get', N('Those collections have no editor.'), 'G-ENGINE'),
          m('PrintOut', 'call', N('On the VS Code host, a PDF of this sheet.')),
          m('ExportAsFixedFormat', 'call', N('On the VS Code host, a PDF of this sheet. XPS is an error.')),
          m('SaveAs', 'call', N('On the VS Code host, this sheet as CSV, text, or a one-sheet workbook.')),
        ],
      },
      {
        name: 'Range',
        members: [
          m('Count, Row, Column, Address, Text, MergeCells', 'get', N('Text is the number format applied to the value.')),
          m('Rows, Columns, Worksheet, Font, Interior, Borders', 'get'),
          m('Formula, FormulaR1C1, Value, Value2, NumberFormat', 'rw'),
          m('ColumnWidth, RowHeight, Hidden', 'rw', N('Hidden applies to a whole row or a whole column. Width 0 hides a column. A hidden row reports height 0.')),
          m('Item, Cells, Resize, Offset, End, SpecialCells', 'call', N('SpecialCells covers constants, formulas, blanks, the last cell, visible cells and comments, plus the value-type bits. Other types fail. More than 64 hits returns the first cell.')),
          m('EntireRow, EntireColumn, CurrentRegion', 'get'),
          m('Clear, ClearContents, ClearFormats, ClearComments, Merge, UnMerge', 'call'),
          m('Select, Activate, Copy, PasteSpecial, Insert, Delete, Sort, AutoFilter', 'call'),
          m('Find', 'call', N('Nothing when there is no hit.')),
          m('Replace', 'call', N('Returns 0 or -1.')),
          m('AddComment, Comment, Hyperlinks, Validation, FormatConditions, Areas', 'get'),
          m('SortFields', 'get', N('Sort fields are not this object. Use Sort.'), 'G-ENGINE'),
        ],
      },
      {
        name: 'Font',
        members: [
          m('Name, Size, Bold, Italic, Underline, Color, ColorIndex, ThemeColor, TintAndShade', 'rw'),
        ],
      },
      {
        name: 'Interior',
        members: [
          m('Color, Pattern, ColorIndex, ThemeColor, TintAndShade, PatternColor', 'rw'),
        ],
      },
      {
        name: 'Borders',
        note: N('An edge is xlEdgeLeft 7, top 8, bottom 9, right 10, inside vertical 11, inside horizontal 12.'),
        members: [
          m('Item, LineStyle', 'rw'),
          m('Border.LineStyle, Weight, Color', 'rw', N('xlNone clears. Continuous, dash, dot, double, medium and thick map onto the stored line.')),
        ],
      },
      {
        name: 'Comments',
        members: [
          m('Item, Add', 'call', N('A note with author, text and visible.')),
          m('Text, Author, Delete', 'rw'),
        ],
      },
      {
        name: 'Hyperlinks',
        members: [
          m('Item, Add, Delete', 'call'),
          m('Address, SubAddress, Range, TextToDisplay, Delete', 'rw'),
          m('Follow', 'call', N('Opening an address is not this call.'), 'G-HOST'),
        ],
      },
      {
        name: 'Validation',
        members: [
          m('Add, Delete', 'call'),
          m('Modify', 'call', N('Change the rule by deleting it and adding another.'), 'G-ENGINE'),
        ],
      },
      {
        name: 'FormatConditions',
        members: [
          m('Item, Add, Delete', 'call', N('Classic cell-is and expression rules.')),
          m('Type, Formula1, Delete', 'rw'),
        ],
      },
      {
        name: 'ListObjects',
        note: N('Lists on a worksheet.'),
        members: [
          m('Item, Add', 'call', N('Add uses the destination first row as column names and ignores the source type.')),
          m('Name, Range, ListColumns, Delete, Resize', 'rw'),
          m('Refresh', 'call', N('A list is not refreshed from a connection.'), 'G-ENGINE'),
          m('ListColumn.Name', 'rw'),
        ],
      },
      {
        name: 'ChartObjects',
        members: [
          m('Item, Add', 'call', N('Add creates a column chart drawing.')),
          m('Name, Left, Top, Width, Height, Chart, Delete', 'rw', N('Positions are points.')),
        ],
      },
      {
        name: 'Chart',
        members: [
          m('ChartType, Name', 'rw', N('Known types: column clustered 51, bar clustered 57, line 4, pie 5, scatter -4169, area 1, doughnut -4120. Any other type fails.')),
          m('SeriesCollection, Axes, SetSourceData', 'get', N('The series and axes are not this object.'), 'G-ENGINE'),
        ],
      },
      {
        name: 'Shapes',
        members: [
          m('Item, AddTextbox', 'call', N('AddTextbox ignores orientation.')),
          m('AddPicture', 'call', N('On the VS Code host, embeds the picture.')),
          m('Name, Left, Top, Width, Height, TextFrame, Delete', 'rw'),
          m('OLEFormat', 'get', N('Embedded objects are kept, not edited.'), 'G-ENGINE'),
          m('TextFrame.TextRange, TextRange.Text', 'rw'),
        ],
      },
      {
        name: 'PageSetup',
        members: [
          m('Orientation, Zoom, FitToPages, margins, PrintArea, CenterHorizontally, CenterVertically', 'rw', N('Margins are inches. PrintOut is G-HOST.')),
        ],
      },
      {
        name: 'PageBreaks',
        note: N('HPageBreaks and VPageBreaks.'),
        members: [
          m('Item, Add', 'call', N('Add stores the 1-based row or column.')),
          m('Location, Delete', 'rw'),
        ],
      },
      {
        name: 'Outline',
        members: [
          m('SummaryRow, ShowLevels', 'rw'),
        ],
      },
      {
        name: 'AutoFilter',
        members: [
          m('Range', 'get'),
          m('ShowAllData', 'call'),
        ],
      },
      {
        name: 'Scenarios',
        members: [
          m('Item, Add', 'call'),
          m('Name, Comment, Show, Delete', 'rw'),
          m('ChangeScenario', 'call', N('Change a scenario by deleting it and adding another.'), 'G-ENGINE'),
        ],
      },
      {
        name: 'PivotTables',
        members: [
          m('Item', 'call', N('Existing tables only.')),
          m('Name', 'get'),
          m('Refresh, PivotFields', 'call', N('A pivot is listed, not rebuilt.'), 'G-ENGINE'),
        ],
      },
    ],
    constantGroups: [
      { title: N('Sheet and calculation'), items: [['xlSheetVisible', '-1'], ['xlSheetHidden', '0'], ['xlSheetVeryHidden', '2'], ['xlCalculationAutomatic', '-4105'], ['xlCalculationManual', '-4135'], ['xlCalculationSemiautomatic', '2'], ['xlA1', '1'], ['xlR1C1', '-4150']] },
      { title: N('Direction and search'), items: [['xlUp', '-4162'], ['xlDown', '-4121'], ['xlToLeft', '-4159'], ['xlToRight', '-4161'], ['xlWhole', '1'], ['xlPart', '2'], ['xlByRows', '1'], ['xlByColumns', '2'], ['xlAscending', '1'], ['xlDescending', '2'], ['xlGuess', '0'], ['xlYes', '1'], ['xlNo', '2']] },
      { title: N('Special cells'), items: [['xlCellTypeConstants', '2'], ['xlCellTypeFormulas', '-4123'], ['xlCellTypeBlanks', '4'], ['xlCellTypeLastCell', '11'], ['xlCellTypeVisible', '12'], ['xlCellTypeComments', '-4144'], ['xlNumbers', '1'], ['xlTextValues', '2'], ['xlLogical', '4'], ['xlErrors', '16']] },
      { title: N('Paste and shift'), items: [['xlPasteAll', '-4104'], ['xlPasteValues', '-4163'], ['xlPasteFormulas', '-4123'], ['xlPasteFormats', '-4122'], ['xlPasteComments', '-4144'], ['xlShiftDown', '-4121'], ['xlShiftToRight', '-4161'], ['xlShiftUp', '-4162'], ['xlShiftToLeft', '-4159']] },
      { title: N('Lines, pattern, page'), items: [['xlContinuous', '1'], ['xlDash', '-4115'], ['xlDot', '-4118'], ['xlDouble', '-4119'], ['xlNone', '-4142'], ['xlHairline', '1'], ['xlThin', '2'], ['xlMedium', '-4138'], ['xlThick', '4'], ['xlUnderlineStyleNone', '-4142'], ['xlUnderlineStyleSingle', '2'], ['xlPatternNone', '-4142'], ['xlPatternSolid', '1'], ['xlPortrait', '1'], ['xlLandscape', '2'], ['xlColorIndexAutomatic', '-4105'], ['xlColorIndexNone', '-4142'], ['msoTrue', '-1'], ['msoFalse', '0'], ['msoTextOrientationHorizontal', '1']] },
      { title: N('Conditions and validation'), items: [['xlCellValue', '1'], ['xlExpression', '2'], ['xlEqual', '3'], ['xlNotEqual', '4'], ['xlGreater', '5'], ['xlLess', '6'], ['xlGreaterEqual', '7'], ['xlLessEqual', '8'], ['xlBetween', '1'], ['xlNotBetween', '2'], ['xlValidateList', '3'], ['xlValidateWholeNumber', '1'], ['xlValidateDecimal', '2'], ['xlValidateDate', '4'], ['xlValidateTime', '5'], ['xlValidateTextLength', '6'], ['xlValidateCustom', '7']] },
    ],
    exampleNote: N('ActiveWorkbook.FullName is G-HOST without the VS Code host. ActiveWorkbook.Creator is G-ENGINE.'),
    example: 'ActiveSheet.Range("A1").Value = "Item"\nActiveSheet.Range("B2").Value = 12.5\nActiveSheet.Range("B4").Formula = "=SUM(B2:B3)"\nblock = ActiveSheet.Range("B2:B3").Value\nActiveSheet.Range("A1:B1").Font.Bold = True\nActiveSheet.Range("B2:B4").NumberFormat = "0.00"\nActiveWorkbook.Names.Add "Total", "=Sheet1!$B$4"\nActiveSheet.Range("A6").Value = WorksheetFunction.Sum(ActiveSheet.Range("B2:B3"))',
  });
})();
