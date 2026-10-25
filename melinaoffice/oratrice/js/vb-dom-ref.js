/* Oratrice Visual Basic DOM API reference.
   The inventory is docs/oratrice_dom_api_specifications.md.
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
    spec: 'docs/oratrice_dom_api_specifications.md',
    rules: [
      N('The entry is Application.ActivePresentation and Presentations. Selection is ActiveWindow.Selection.'),
      N('Colour is ColorFormat. RGB is an Office BGR long. ObjectThemeColor is a theme slot. SchemeColor is one of the eight 2003 slots. Writing tint, shade or brightness fails.'),
      N('TextFrame2 is the same object as TextFrame. Members that exist only on TextFrame2 fail or are absent.'),
      N('Shape Left, Top, Width and Height are points.'),
      N('A text range is an anchor inside one shape. Text joins paragraphs with a carriage return. Characters.Count counts stored characters and does not include that return.'),
      N('A held shape is read again by id, so undo does not leave Left on the shape undo replaced. Read a text range again after undo.'),
      N('ActiveWindow is thin. Selection and Presentation work. View, ViewType, Caption, ActivePane, Panes and Split fail.'),
      N('SlideShowWindow has Next, Previous, GotoSlide and Exit. Pointer, state and accelerators fail.'),
      N('Closing the last presentation opens a blank one. The editor always has a presentation.'),
      N('Document macros are Visual Basic 6 source in the file. The object model does not list that store.'),
    ],
    gaps: [
      { code: 'G-TEXT', note: N('The member needs a character index inside one shape. PowerPoint counts the paragraph mark. Oratrice stores a paragraph index and an offset. Characters(start, length), Words, Sentences, Lines, Runs, Find, Replace and the Bound members fail. Characters.Count works. Item fails.') },
      { code: 'G-ENGINE', note: N('There is no document operation for this member. The object model does not lay out the slide.') },
      { code: 'G-HOST', note: N('The member needs a path, a window, a printer or a modal.') },
    ],
    unavailable: [
      u('TimeLine, Sequence, Effect', 'G-ENGINE', N('A slide animation record is not a VBA timeline.')),
      u('Chart, ChartData', 'G-ENGINE', N('A chart shape is not a Chart object.')),
      u('Diagram, DiagramNode, SmartArt', 'G-ENGINE', N('IntelliArt nodes are not this model.')),
      u('OLEFormat, MediaFormat, CustomerData', 'G-ENGINE', N('No operation on those parts.')),
      u('NotesPage', 'G-ENGINE', N('Notes are a string on the slide, not a shape page.')),
      u('Comments, Comment', 'G-ENGINE', N('Comment parts are kept. There is no comment editor.')),
      u('Tags', 'G-ENGINE', N('There is no tag bag.')),
      u('PrintOptions', 'G-ENGINE', N('Printing is the print dialog.')),
      u('Panes, Pane, View, CommandBars, FileDialog, Dialogs', 'G-HOST', N('Windows, panes and the printer.')),
      u('ActionSettings, NamedSlideShows, SectionProperties, ColorSchemes, AnimationSettings', 'G-ENGINE', N('No matching operation.')),
    ],
    objects: [
      {
        name: 'Application',
        members: [
          m('ActivePresentation', 'get', N('The open presentation.')),
          m('Presentations', 'get'),
          m('ActiveWindow', 'get', N('The document window.')),
          m('SlideShowWindows', 'get', N('The running show, or empty.')),
        ],
      },
      {
        name: 'Presentations',
        members: [
          m('Count, Item, First, Last', 'get', N('1-based. Presentations(1) is Item.')),
          m('Add', 'call', N('No arguments. Returns the new presentation and makes it active. WithWindow fails.'), ),
          m('Open', 'call', N('Opening is a file picker, not a path.'), 'G-HOST'),
        ],
      },
      {
        name: 'Presentation',
        members: [
          m('Name', 'get', N('The file name, or Presentation1.')),
          m('FullName, Path', 'get', N('On the VS Code host, the file path and its folder. Elsewhere G-HOST.')),
          m('Saved', 'rw', N('True when the deck is not dirty. A new unedited deck reads True.')),
          m('Slides', 'get'),
          m('PageSetup', 'get'),
          m('SlideShowSettings', 'get'),
          m('SlideShowWindows', 'get'),
          m('Designs', 'get', N('The built-in design list.')),
          m('SlideMaster', 'get', N('Name and headers and footers. Shapes, layouts, design and background fail.')),
          m('BuiltInDocumentProperties', 'get', N('Title, Subject, Author, Keywords, Comments, Category, Company.')),
          m('CustomDocumentProperties', 'get', N('String values. Add with LinkToContent, or a type other than 4, fails.')),
          m('Undo, Redo', 'call', N('Returns -1 when a step ran.')),
          m('Save', 'call', N('Calls the app save when a file is already saved. Otherwise, in the browser, this fails. With no save function, it clears the dirty flag.')),
          m('Close', 'call', N('Drops this presentation. The last one is replaced by a blank deck.')),
          m('SaveAs, SaveAs2', 'call', N('On the VS Code host, writes that path. Elsewhere G-HOST.')),
          m('Export', 'call', N('On the VS Code host, PNG or JPEG of every slide.')),
          m('ExportAsFixedFormat', 'call', N('On the VS Code host, a PDF. XPS is an error.')),
          m('PrintOut', 'call', N('On the VS Code host, a PDF of the slides beside the file.')),
          m('FollowHyperlink', 'call', N('On the VS Code host, opens http, https, or mailto.')),
          m('NewWindow', 'call', N('A new window is not this call.'), 'G-HOST'),
          m('CreateVideo, Merge, PublishSlides', 'call', N('Those operations are not on the deck.'), 'G-ENGINE'),
        ],
      },
      {
        name: 'DocumentWindow',
        members: [
          m('Selection', 'get'),
          m('Presentation', 'get', N('The active presentation.')),
          m('View, ViewType, Caption, ActivePane, Panes, Split', 'get', N('The window model is not on this page.'), 'G-HOST'),
        ],
      },
      {
        name: 'Selection',
        members: [
          m('Type', 'get', N('ppSelectionNone 0, ppSelectionSlides 1, ppSelectionShapes 2, ppSelectionText 3.')),
          m('ShapeRange', 'get', N('Left, Top, Delete, and the shapes. Fails when nothing is selected.')),
          m('SlideRange', 'get', N('The selected slides, or the current slide.')),
          m('TextRange', 'get', N('The selected text anchor. Fails when the selection is not text.')),
          m('Unselect, Delete, Copy, Cut, Paste', 'call', N('Shape and text clipboard. A slide-only paste is not this call.')),
          m('ChildShapeRange', 'get', N('Child shapes are not a separate range.'), 'G-ENGINE'),
        ],
      },
      {
        name: 'Slides',
        members: [
          m('Add', 'call', N('Add(index, layout). Index is 1-based. An unmapped layout code fails.')),
          m('Range', 'call'),
        ],
      },
      {
        name: 'Slide',
        members: [
          m('Name', 'rw'),
          m('SlideIndex, SlideNumber', 'get', N('SlideNumber adds FirstSlideNumber.')),
          m('SlideID', 'get', N('Slide ids are strings, not this long.'), 'G-ENGINE'),
          m('Layout', 'rw', N('An unmapped stored layout reads ppLayoutCustom (32). An unmapped write fails.')),
          m('Shapes', 'get'),
          m('Design', 'rw', N('A design name or key. An unknown name fails.')),
          m('HeadersFooters', 'get', N('The presentation header and footer record.')),
          m('SlideShowTransition', 'get'),
          m('Select, Delete, Duplicate, MoveTo', 'call'),
          m('Hyperlinks', 'get', N('Links on the slide text. Add writes onto those paragraphs.')),
          m('NotesPage, Comments, TimeLine, Tags, CustomerData', 'get', N('The same gaps as the objects in Not available.'), 'G-ENGINE'),
        ],
      },
      {
        name: 'Shapes',
        note: N('Shapes on a slide can add. GroupItems cannot.'),
        members: [
          m('Item, Count, Title, Placeholders', 'get', N('Title is the title placeholder.')),
          m('AddShape', 'call', N('AddShape(type, left, top, width, height). Types are in Constants. Any other auto-shape type fails.')),
          m('AddTextbox', 'call', N('Horizontal only.')),
          m('AddTable', 'call', N('AddTable(rows, cols, left, top, width).')),
          m('AddPicture', 'call', N('On the VS Code host, embeds the picture.')),
          m('AddChart, AddSmartArt, AddDiagram, AddMediaObject, AddOLEObject, AddTextEffect', 'call', N('Those inserts are not this call.'), 'G-ENGINE'),
        ],
      },
      {
        name: 'Shape',
        members: [
          m('Left, Top, Width, Height, Rotation', 'rw', N('Points. Rotation is degrees.')),
          m('Name', 'rw'),
          m('Id, Type, HasTextFrame, HasTable, ZOrderPosition', 'get', N('Type is auto shape, chart, group, line, picture, placeholder, text effect, text box or table.')),
          m('AutoShapeType', 'rw', N('The mapped presets only.')),
          m('TextFrame, TextFrame2', 'get', N('The same object.')),
          m('Fill, Line', 'get', N('Solid fill and a solid line. A gradient reads as a gradient. Writing a type other than solid fails. Line.DashStyle fails.')),
          m('PlaceholderFormat', 'get', N('Type and ContainedType. Vertical title and body are types 5 and 6. Fails when the shape is not a placeholder.')),
          m('Table', 'get', N('Fails when the shape is not a table.')),
          m('GroupItems', 'get', N('Fails when the shape is not a group.')),
          m('Delete, Duplicate, Select, ZOrder', 'call', N('ZOrder is bring to front 0, send to back 1, bring forward 2, send backward 3. Z-order index 0 is the back, so ZOrderPosition is the index plus one.')),
          m('Hyperlinks', 'get', N('Links stored on the shape text.')),
          m('Chart, OLEFormat, MediaFormat, ActionSettings, AnimationSettings, Nodes, Tags', 'get', N('Those objects are not on the shape.'), 'G-ENGINE'),
        ],
      },
      {
        name: 'TextFrame',
        members: [
          m('TextRange', 'get', N('The whole body.')),
          m('HasText', 'get'),
          m('WordWrap, AutoSize', 'rw', N('ppAutoSizeNone 0, ppAutoSizeShapeToFitText 1. msoAutoSizeTextToFitShape 2 is autofit norm.')),
          m('MarginLeft, MarginTop, MarginRight, MarginBottom', 'rw', N('Points.')),
          m('VerticalAnchor', 'rw', N('ppAnchorTop 1, ppAnchorMiddle 3, ppAnchorBottom 4.')),
          m('Orientation', 'get', N('Text orientation is not this property.'), 'G-ENGINE'),
        ],
      },
      {
        name: 'TextRange',
        members: [
          m('Text', 'rw', N('Paragraphs joined by a carriage return. A whole-body write replaces the body. A partial write deletes that span and inserts.')),
          m('InsertBefore, InsertAfter, Delete, Select', 'call'),
          m('Font, ParagraphFormat', 'get'),
          m('IndentLevel', 'rw', N('1-based. Levels 1 through 9. Stored 0-based.')),
          m('Paragraphs', 'get'),
          m('Characters', 'get', N('Count only.')),
          m('InsertSymbol, InsertDateTime, InsertSlideNumber', 'call', N('The date uses the datetime1 format when the page has it.')),
          m('Start, Length, BoundLeft, BoundTop, BoundWidth, BoundHeight', 'get', N('A character index inside the shape.'), 'G-TEXT'),
          m('Find, Replace', 'call', N('Find and replace need that index.'), 'G-TEXT'),
          m('Words, Sentences, Lines, Runs', 'get', N('Those ranges need that index.'), 'G-TEXT'),
        ],
      },
      {
        name: 'Paragraph',
        members: [
          m('Alignment', 'rw', N('ppAlignLeft 1 through ppAlignDistribute 5.')),
        ],
      },
      {
        name: 'Font',
        members: [
          m('Name, NameFarEast, Size, Bold, Italic, Underline', 'rw', N('Name writes the Latin face. Underline msoTrue is a single line. NameAscii, NameComplexScript, SizeBi, BoldBi and ItalicBi fail.')),
          m('Color', 'get', N('ColorFormat.')),
        ],
      },
      {
        name: 'ParagraphFormat',
        members: [
          m('Alignment, SpaceBefore, SpaceAfter, SpaceWithin, LineRuleWithin', 'rw', N('Point spacing stores points. A percent spacing value fails to read. LineRuleWithin is -1 for a multiple and 0 for points.')),
          m('Bullet.Type, Bullet.Visible', 'rw', N('ppBulletNone 0, ppBulletUnnumbered 1, ppBulletNumbered 2.')),
        ],
      },
      {
        name: 'Table',
        members: [
          m('Cell, Rows, Columns', 'call', N('Cell(row, column). Rows.Add and Columns.Add append or insert.')),
          m('Row.Height, Column.Width, Cell.Shape', 'rw'),
          m('Merge, ApplyStyle', 'call', N('Merge and applying a table style are not this call.'), 'G-ENGINE'),
        ],
      },
      {
        name: 'Fill',
        members: [
          m('Type', 'rw', N('0 none, 1 solid. A gradient reads 3. Writing another type fails.')),
          m('ForeColor, Transparency, Visible', 'rw', N('Transparency is 1 minus alpha.')),
        ],
      },
      {
        name: 'Line',
        members: [
          m('Weight, ForeColor, Visible', 'rw', N('Visible False sets the line to none.')),
          m('DashStyle', 'get', N('Dash style is not stored on the line.'), 'G-ENGINE'),
        ],
      },
      {
        name: 'ColorFormat',
        members: [
          m('RGB', 'rw', N('BGR long.')),
          m('ObjectThemeColor, SchemeColor', 'rw', N('See Constants.')),
          m('TintAndShade, Brightness', 'get', N('Tint, shade and brightness are not written.'), 'G-ENGINE'),
        ],
      },
      {
        name: 'Hyperlink',
        members: [
          m('Add', 'call', N('Add(anchor, address, subAddress, screenTip, textToDisplay). Writes the link onto the anchor runs.')),
          m('Address, SubAddress, ScreenTip, TextToDisplay', 'rw'),
          m('Follow', 'call', N('Opening an address is not this call.'), 'G-HOST'),
          m('Type', 'get', N('The link type is not published.'), 'G-ENGINE'),
        ],
      },
      {
        name: 'PageSetup',
        members: [
          m('SlideWidth, SlideHeight, FirstSlideNumber', 'rw', N('Points, and the first number.')),
          m('SlideSize, NotesOrientation, SlideOrientation', 'get', N('Set width and height directly.'), 'G-ENGINE'),
        ],
      },
      {
        name: 'Design',
        members: [
          m('Designs, Name', 'get', N('Built-in designs.')),
        ],
      },
      {
        name: 'HeadersFooters',
        members: [
          m('DateAndTime, Footer, SlideNumber', 'get', N('Visible and Text.')),
          m('DisplayOnTitleSlide', 'rw', N('The opposite of not on title.')),
          m('Format, UseFormat', 'get', N('The header format is not this object.'), 'G-ENGINE'),
        ],
      },
      {
        name: 'SlideShowTransition',
        members: [
          m('EntryEffect', 'rw', N('The effects in Constants. A stored effect outside that list fails to read.')),
          m('Speed', 'rw', N('ppTransitionSpeedSlow 1, medium 2, fast 3.')),
          m('AdvanceOnClick, AdvanceOnTime, AdvanceTime, Hidden', 'rw', N('AdvanceTime is seconds. The deck stores milliseconds.')),
          m('Duration, SoundEffect', 'get', N('Duration and sound are not this object.'), 'G-ENGINE'),
        ],
      },
      {
        name: 'SlideShowSettings',
        members: [
          m('StartingSlide, EndingSlide', 'rw', N('1-based. Both empty means the whole deck.')),
          m('RangeType', 'rw', N('ppShowAll 1, ppShowSlideRange 2. ppShowNamedSlideShow 3 fails.')),
          m('LoopUntilStopped, AdvanceMode, ShowType', 'rw', N('ppShowTypeWindow 2 is G-HOST. Kiosk is 3.')),
          m('Run', 'call', N('Starts the headless show position. When the page has a running show, it starts that too.')),
          m('NamedSlideShows, PointerColor', 'get', N('Named shows and the pointer colour are not this object.'), 'G-ENGINE'),
        ],
      },
      {
        name: 'SlideShowWindow',
        note: N('Hidden slides are skipped in the show order.'),
        members: [
          m('View.Next, View.Previous, View.GotoSlide, View.Exit', 'call', N('Next before Run fails.')),
          m('View.CurrentShowPosition', 'get', N('1-based while the show is on, otherwise 0.')),
          m('View.State, View.PointerType, View.AcceleratorsEnabled', 'get', N('Pointer, state and accelerators are the host.'), 'G-HOST'),
        ],
      },
    ],
    constantGroups: [
      { title: N('Layouts'), items: [['ppLayoutTitle', '1'], ['ppLayoutText', '2'], ['ppLayoutTwoColumnText', '3'], ['ppLayoutTable', '4'], ['ppLayoutChart', '8'], ['ppLayoutTitleOnly', '11'], ['ppLayoutBlank', '12'], ['ppLayoutTextAndObject', '13'], ['ppLayoutObjectAndText', '14'], ['ppLayoutLargeObject', '15'], ['ppLayoutObject', '16'], ['ppLayoutObjectOverText', '19'], ['ppLayoutTextOverObject', '20'], ['ppLayoutTextAndTwoObjects', '21'], ['ppLayoutTwoObjectsAndText', '22'], ['ppLayoutFourObjects', '24'], ['ppLayoutVerticalText', '25'], ['ppLayoutVerticalTitleAndText', '27'], ['ppLayoutTwoObjects', '29'], ['ppLayoutObjectAndTwoObjects', '30'], ['ppLayoutTwoObjectsAndObject', '31'], ['ppLayoutCustom', '32']] },
      { title: N('Placeholders'), items: [['ppPlaceholderTitle', '1'], ['ppPlaceholderBody', '2'], ['ppPlaceholderCenterTitle', '3'], ['ppPlaceholderSubtitle', '4'], ['ppPlaceholderVerticalTitle', '5'], ['ppPlaceholderVerticalBody', '6'], ['ppPlaceholderObject', '7'], ['ppPlaceholderChart', '8'], ['ppPlaceholderTable', '12'], ['ppPlaceholderSlideNumber', '13'], ['ppPlaceholderFooter', '15'], ['ppPlaceholderDate', '16']] },
      { title: N('AutoShapes'), items: [['msoShapeRectangle', '1'], ['msoShapeDiamond', '4'], ['msoShapeRoundedRectangle', '5'], ['msoShapeOval', '9'], ['msoShapeHexagon', '10'], ['msoShapeCross', '11'], ['msoShapeCan', '13'], ['msoShapeCube', '14'], ['msoShapeDonut', '18'], ['msoShapeBlockArc', '20'], ['msoShapeHeart', '21'], ['msoShapeLightningBolt', '22'], ['msoShapeSun', '23'], ['msoShapeMoon', '24'], ['msoShapeArc', '25']] },
      { title: N('Shape types'), items: [['msoAutoShape', '1'], ['msoChart', '3'], ['msoGroup', '6'], ['msoLine', '9'], ['msoPicture', '13'], ['msoPlaceholder', '14'], ['msoTextEffect', '15'], ['msoTextBox', '17'], ['msoTable', '19']] },
      { title: N('Scheme colours'), items: [['ppBackground', '1'], ['ppForeground', '2'], ['ppShadow', '3'], ['ppTitle', '4'], ['ppFill', '5'], ['ppAccent', '6'], ['ppAccentAndHyperlink', '7'], ['ppAccentAndFollowedHyperlink', '8']] },
      { title: N('Transitions'), items: [['ppEffectNone', '0'], ['ppEffectCut', '257'], ['ppEffectCutThroughBlack', '258'], ['ppEffectDissolve', '1537'], ['ppEffectWipeLeft', '2817'], ['ppEffectWipeUp', '2818'], ['ppEffectWipeRight', '2819'], ['ppEffectWipeDown', '2820'], ['ppEffectFadeSmoothly', '3849'], ['ppEffectPushDown', '3852'], ['ppEffectPushLeft', '3853'], ['ppEffectPushRight', '3854'], ['ppEffectPushUp', '3855']] },
      { title: N('Show'), items: [['ppShowAll', '1'], ['ppShowSlideRange', '2'], ['ppShowNamedSlideShow', '3'], ['ppShowTypeSpeaker', '1'], ['ppShowTypeWindow', '2'], ['ppShowTypeKiosk', '3'], ['ppSlideShowManualAdvance', '1'], ['ppSlideShowUseSlideTimings', '2']] },
    ],
    exampleNote: N('TextRange.Characters(1, 5) is G-TEXT. Characters.Count on Hello is 5. Content layouts 13 through 16, 19 through 22, 24, 25, 27 and 29 through 31 are also layout codes. Theme colours run from Dark1 1 through Background2 16.'),
    example: 'ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Text = "Hello"\nActivePresentation.Slides.Add 2, ppLayoutText\nActiveWindow.Selection.ShapeRange.Left = 72',
  });
})();
