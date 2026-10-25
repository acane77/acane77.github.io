/* Matheuse — Excel object model on the open workbook.
   Inventory: docs/matheuse_dom_api_specifications.md
   Rows and columns here are 1-based. The workbook model stays 0-based. */
(function () {
  'use strict';
  const L = window.L;
  const M = L.model;
  const F = L.formula;
  const O = L.ops;
  const NF = L.numfmt;
  const dom = (L.dom = L.dom || {});
  const MAXR = M.MAXR, MAXC = M.MAXC;
  const BIG = 1000000;

  const XL = {
    xlSheetVisible: -1, xlSheetHidden: 0, xlSheetVeryHidden: 2,
    xlCalculationAutomatic: -4105, xlCalculationManual: -4135, xlCalculationSemiautomatic: 2,
    xlA1: 1, xlR1C1: -4150,
    xlUp: -4162, xlDown: -4121, xlToLeft: -4159, xlToRight: -4161,
    xlWhole: 1, xlPart: 2, xlByRows: 1, xlByColumns: 2,
    xlAscending: 1, xlDescending: 2, xlGuess: 0, xlYes: 1, xlNo: 2,
    xlCellTypeConstants: 2, xlCellTypeFormulas: -4123, xlCellTypeBlanks: 4,
    xlCellTypeLastCell: 11, xlCellTypeVisible: 12, xlCellTypeComments: -4144,
    xlNumbers: 1, xlTextValues: 2, xlLogical: 4, xlErrors: 16,
    xlPasteAll: -4104, xlPasteValues: -4163, xlPasteFormulas: -4123, xlPasteFormats: -4122, xlPasteComments: -4144,
    xlShiftDown: -4121, xlShiftToRight: -4161, xlShiftUp: -4162, xlShiftToLeft: -4159,
    xlContinuous: 1, xlDash: -4115, xlDot: -4118, xlDouble: -4119, xlNone: -4142, xlLineStyleNone: -4142,
    xlHairline: 1, xlThin: 2, xlMedium: -4138, xlThick: 4,
    xlEdgeLeft: 7, xlEdgeTop: 8, xlEdgeBottom: 9, xlEdgeRight: 10, xlInsideVertical: 11, xlInsideHorizontal: 12,
    xlUnderlineStyleNone: -4142, xlUnderlineStyleSingle: 2,
    xlPatternNone: -4142, xlPatternSolid: 1,
    xlPortrait: 1, xlLandscape: 2,
    xlCellValue: 1, xlExpression: 2,
    xlEqual: 3, xlNotEqual: 4, xlGreater: 5, xlLess: 6, xlGreaterEqual: 7, xlLessEqual: 8, xlBetween: 1, xlNotBetween: 2,
    xlValidateList: 3, xlValidateWholeNumber: 1, xlValidateDecimal: 2, xlValidateDate: 4, xlValidateTime: 5,
    xlValidateTextLength: 6, xlValidateCustom: 7,
    xlColumnClustered: 51, xlBarClustered: 57, xlLine: 4, xlPie: 5, xlXYScatter: -4169, xlArea: 1, xlDoughnut: -4120,
    xlThemeColorDark1: 1, xlThemeColorLight1: 2, xlThemeColorDark2: 3, xlThemeColorLight2: 4,
    xlThemeColorAccent1: 5, xlThemeColorAccent2: 6, xlThemeColorAccent3: 7, xlThemeColorAccent4: 8,
    xlThemeColorAccent5: 9, xlThemeColorAccent6: 10, xlThemeColorHyperlink: 11, xlThemeColorFollowedHyperlink: 12,
    xlColorIndexAutomatic: -4105, xlColorIndexNone: -4142,
    msoTrue: -1, msoFalse: 0, msoTextOrientationHorizontal: 1,
    xlDialogOpen: 1, xlDialogSaveAs: 5, xlDialogPageSetup: 7, xlDialogPrint: 8,
    xlDialogFormatFont: 150, xlDialogRowHeight: 127, xlDialogColumnWidth: 47, xlDialogZoom: 256,
    xlDialogFormulaFind: 64, xlDialogFormulaReplace: 130, xlDialogFormulaGoto: 63, xlDialogSelectSpecial: 132,
    xlDialogPasteSpecial: 53, xlDialogInsert: 55, xlDialogEditDelete: 54, xlDialogSort: 39, xlDialogDataSeries: 40,
    xlDialogFunctionWizard: 450, xlDialogDefineName: 61, xlDialogCreateNames: 62, xlDialogApplyNames: 133,
    xlDialogPasteNames: 58, xlDialogSubtotalCreate: 398, xlDialogTextToColumns: 422, xlDialogFilterAdvanced: 370,
    xlDialogGoalSeek: 198, xlDialogScenarioCells: 305, xlDialogTable: 41, xlDialogConsolidate: 191,
    xlDialogPivotTableWizard: 312, xlDialogSummaryInfo: 474, xlDialogOptionsGeneral: 356, xlDialogAutoCorrect: 485,
    xlDialogInsertHyperlink: 596, xlDialogNote: 154, xlDialogWorkbookProtect: 417, xlDialogProtectDocument: 28,
    xlCSV: 6, xlTextWindows: -4158, xlOpenXMLWorkbook: 51, xlOpenXMLWorkbookMacroEnabled: 52,
    msoFileDialogOpen: 1, msoFileDialogSaveAs: 2, msoFileDialogFilePicker: 3, msoFileDialogFolderPicker: 4,
    xlTypePDF: 0, xlTypeXPS: 1,
  };
  Object.assign(dom, XL);

  function gap(code, member) {
    const e = new Error(member ? member + ' is not available (' + code + ')' : code);
    e.code = code;
    throw e;
  }
  function missing() {
    const e = new Error('The requested member of the collection does not exist.');
    e.code = 'missing';
    throw e;
  }
  function on(v) { return v === true || v === -1; }
  function tri(v) { return v ? -1 : 0; }
  function bgrToHex(n) {
    n = n >>> 0;
    const r = n & 255, g = (n >> 8) & 255, b = (n >> 16) & 255;
    return '#' + [r, g, b].map((x) => x.toString(16).padStart(2, '0')).join('').toUpperCase();
  }
  function hexToBgr(hex) {
    const h = String(hex || '').replace('#', '');
    if (h.length < 6) return 0;
    return parseInt(h.slice(0, 2), 16) | (parseInt(h.slice(2, 4), 16) << 8) | (parseInt(h.slice(4, 6), 16) << 16);
  }
  /* Excel theme-color enum → SpreadsheetML theme index (0 lt1, 1 dk1, …). */
  const THEME_TO = { 1: 1, 2: 0, 3: 3, 4: 2, 5: 4, 6: 5, 7: 6, 8: 7, 9: 8, 10: 9, 11: 10, 12: 11 };
  const THEME_FROM = {};
  for (const k in THEME_TO) THEME_FROM[THEME_TO[k]] = +k;

  const flags = { screen: true, alerts: true };
  const hold = { books: [], cur: 0 };
  let clip = null;

  function paint() {
    if (!flags.screen) return;
    try { if (L.grid && L.grid.paint) L.grid.paint(); } catch (e) { /* the page is not mounted */ }
  }
  function books() {
    if (L.app && Array.isArray(L.app.books) && L.app.books.length) return L.app.books;
    if (!hold.books.length) hold.books.push(blankRecord('Book1'));
    return hold.books;
  }
  function curIndex() {
    if (L.app && Array.isArray(L.app.books) && L.app.books.length) return L.app.cur || 0;
    return hold.cur || 0;
  }
  function setCur(i) {
    if (L.app && L.app.books && L.app.books.length) {
      if (L.app.switchBook) L.app.switchBook(i);
      else L.app.cur = i;
    } else hold.cur = i;
    paint();
  }
  function blankWb() {
    const w = L.app && L.app.blankWorkbook ? L.app.blankWorkbook() : new M.Workbook();
    if (!w.sheets.length) { w.addSheet('Sheet1'); w.addSheet('Sheet2'); w.addSheet('Sheet3'); }
    return w;
  }
  function blankRecord(name) {
    const w = blankWb();
    w.fileName = name;
    return { wb: w, name: name, dirty: false, saved: true, untitled: true, type: 'xlsx' };
  }
  function bookOf(wb) { return books().find((b) => b.wb === wb) || null; }
  function curBook() { return books()[curIndex()]; }
  function mark(wb) {
    const b = bookOf(wb);
    if (b) b.dirty = true;
  }
  function edit(wb, label, fn) {
    const out = O.tx(wb, label, fn);
    mark(wb);
    paint();
    return out;
  }
  function gridSheet(sh) {
    if (sh.kind === 'macrosheet' || sh.kind === 'dialogsheet') gap('G-ENGINE', 'Range');
    if (sh.kind === 'chart') gap('G-ENGINE', 'Range');
  }

  const Application = {};
  function stamp(obj, parent) {
    obj.Application = Application;
    if (typeof parent === 'function') Object.defineProperty(obj, 'Parent', { enumerable: true, configurable: true, get: parent });
    else obj.Parent = parent == null ? Application : parent;
    Object.defineProperty(obj, 'Creator', { enumerable: true, configurable: true, get() { gap('G-ENGINE', 'Creator'); } });
    return obj;
  }
  function denyGet(obj, names, code) {
    for (const n of names) {
      Object.defineProperty(obj, n, { enumerable: true, configurable: true, get() { gap(code, n); }, set() { gap(code, n); } });
    }
  }
  function denyFn(obj, names, code) {
    for (const n of names) obj[n] = function () { gap(code, n); };
  }
  function gapBag(code) {
    const o = stamp({}, Application);
    return new Proxy(o, {
      get(t, p) {
        if (p === 'Application' || p === 'Parent' || p === 'Creator' || typeof p === 'symbol') return t[p];
        gap(code, String(p));
      },
    });
  }

  function collection(parent, listFn, make, nameOf) {
    const o = stamp({}, parent);
    Object.defineProperty(o, 'Count', { enumerable: true, get() { return listFn().length; } });
    o.Item = function (index) {
      const list = listFn();
      if (typeof index === 'string') {
        const key = index.toLowerCase();
        const hit = list.find((x) => String(nameOf(x)).toLowerCase() === key);
        if (!hit) missing();
        return make(hit);
      }
      const i = (+index) - 1;
      if (!(i >= 0) || i >= list.length || i !== (i | 0)) missing();
      return make(list[i]);
    };
    Object.defineProperty(o, 'First', { enumerable: true, get() { const list = listFn(); if (!list.length) missing(); return make(list[0]); } });
    Object.defineProperty(o, 'Last', { enumerable: true, get() { const list = listFn(); if (!list.length) missing(); return make(list[list.length - 1]); } });
    return o;
  }

  function cellOut(v) {
    if (v == null) return null;
    if (M.isErr && M.isErr(v)) return String(v.e || v);
    if (typeof v === 'boolean') return v ? -1 : 0;
    return v;
  }
  function cellIn(v) {
    if (v == null || v === '') return null;
    if (v === true || v === -1) return true;
    if (v === false || v === 0 && typeof v === 'boolean') return false;
    if (typeof v === 'string' && M.ERR && M.ERR[v]) return M.ERR[v];
    return v;
  }
  function styleAt(sh, r, c) {
    const cell = sh.get(r, c);
    if (L.layout && L.layout.styleOf) return L.layout.styleOf(sh, r, c, cell);
    return sh.wb.styles.get(cell && cell.s || 0);
  }
  function areasOf(rg) { return rg._areas; }
  function tooBig(areas) {
    let n = 0;
    for (const a of areas) {
      n += (a.r2 - a.r1 + 1) * (a.c2 - a.c1 + 1);
      if (n > BIG) return true;
    }
    return false;
  }
  function eachPos(areas, fn) {
    if (tooBig(areas)) gap('G-ENGINE', 'Value');
    for (const a of areas) {
      for (let r = a.r1; r <= a.r2; r++) for (let c = a.c1; c <= a.c2; c++) fn(r, c);
    }
  }
  function topLeft(areas) { return areas[0]; }
  function oneCell(areas) {
    return areas.length === 1 && areas[0].r1 === areas[0].r2 && areas[0].c1 === areas[0].c2;
  }
  function readValue(sh, r, c) {
    const cell = sh.get(r, c);
    if (!cell) return null;
    let v = cell.v;
    if (cell.f != null && cell.dirty && L.calc && L.calc.evalCell) {
      try { L.calc.evalCell(sh, r, c, cell); v = cell.v; } catch (e) { v = cell.v; }
    }
    return cellOut(v);
  }
  function writeScalar(sh, r, c, v) {
    if (typeof v === 'string' && v.charAt(0) === '=') {
      const f = F.toStore(v.slice(1));
      O.setFormula(sh, r, c, f);
      return;
    }
    if (v == null || v === '') { O.setValue(sh, r, c, null); return; }
    if (v === true || v === -1) { O.setValue(sh, r, c, true); return; }
    if (v === false) { O.setValue(sh, r, c, false); return; }
    if (typeof v === 'string' && M.ERR && M.ERR[v]) { O.setValue(sh, r, c, M.ERR[v]); return; }
    O.setValue(sh, r, c, v);
  }
  function writeFormula(sh, r, c, text, r1c1) {
    let body = String(text == null ? '' : text);
    if (r1c1) {
      if (!L.xmlss || !L.xmlss.fromR1C1) gap('G-ENGINE', 'FormulaR1C1');
      body = L.xmlss.fromR1C1(body, r, c);
    } else body = body.replace(/^=/, '');
    const stored = F.toStore(body);
    O.setFormula(sh, r, c, stored);
  }
  function shiftFormula(base, dr, dc) {
    if (!dr && !dc) return base;
    try { return F.toText(F.shift(F.parse(base), dr, dc), { store: true }); } catch (e) { return base; }
  }

  function parseAreas(text) {
    const parts = String(text).split(',');
    const areas = [];
    for (const part of parts) {
      const rg = F.parseRange(part.trim());
      if (!rg) {
        const e = new Error('The range address is not valid.');
        e.code = 'missing';
        throw e;
      }
      areas.push(rg);
    }
    if (!areas.length) missing();
    return areas;
  }
  function normalizeArea(a) {
    let r1 = Math.min(a.r1, a.r2), r2 = Math.max(a.r1, a.r2), c1 = Math.min(a.c1, a.c2), c2 = Math.max(a.c1, a.c2);
    r1 = Math.max(0, Math.min(MAXR - 1, r1));
    r2 = Math.max(0, Math.min(MAXR - 1, r2));
    c1 = Math.max(0, Math.min(MAXC - 1, c1));
    c2 = Math.max(0, Math.min(MAXC - 1, c2));
    return { r1: r1, c1: c1, r2: r2, c2: c2 };
  }

  function makeRange(sh, areas) {
    gridSheet(sh);
    areas = areas.map(normalizeArea);
    const rg = stamp({}, () => sheetObj(sh));
    rg._areas = areas;
    rg._sheet = sh;
    rg._matheuseRange = true;
    Object.defineProperty(rg, 'Count', { enumerable: true, get() { return areas.reduce((n, a) => n + (a.r2 - a.r1 + 1) * (a.c2 - a.c1 + 1), 0); } });
    Object.defineProperty(rg, 'Row', { enumerable: true, get() { return areas[0].r1 + 1; } });
    Object.defineProperty(rg, 'Column', { enumerable: true, get() { return areas[0].c1 + 1; } });
    Object.defineProperty(rg, 'Rows', { enumerable: true, get() { return rowColCollection(sh, areas, 'r'); } });
    Object.defineProperty(rg, 'Columns', { enumerable: true, get() { return rowColCollection(sh, areas, 'c'); } });
    Object.defineProperty(rg, 'Worksheet', { enumerable: true, get() { return sheetObj(sh); } });
    Object.defineProperty(rg, 'Address', {
      enumerable: true,
      get() { return areas.map((a) => F.rangeName(a)).join(','); },
    });
    Object.defineProperty(rg, 'Formula', {
      enumerable: true,
      get() { return readFormula(sh, areas, false); },
      set(v) { putFormula(sh, areas, v, false); },
    });
    Object.defineProperty(rg, 'FormulaR1C1', {
      enumerable: true,
      get() { return readFormula(sh, areas, true); },
      set(v) { putFormula(sh, areas, v, true); },
    });
    Object.defineProperty(rg, 'Value', {
      enumerable: true,
      get() { return readBlock(sh, areas); },
      set(v) { putBlock(sh, areas, v, false); },
    });
    Object.defineProperty(rg, 'Value2', {
      enumerable: true,
      get() { return rg.Value; },
      set(v) { rg.Value = v; },
    });
    Object.defineProperty(rg, 'Text', {
      enumerable: true,
      get() {
        const a = topLeft(areas);
        const v = readValue(sh, a.r1, a.c1);
        if (v == null) return '';
        if (typeof v === 'string' && v.charAt(0) === '#') return v;
        const st = styleAt(sh, a.r1, a.c1);
        const raw = v === -1 ? true : v === 0 && typeof readRaw(sh, a.r1, a.c1) === 'boolean' ? false : v;
        try { return NF.text(st.nf || 'General', raw === -1 ? true : raw); } catch (e) { return String(v); }
      },
    });
    Object.defineProperty(rg, 'NumberFormat', {
      enumerable: true,
      get() { return styleAt(sh, areas[0].r1, areas[0].c1).nf || 'General'; },
      set(v) { O.format(sh, areas, { nf: v == null ? 'General' : String(v) }, 'Number Format'); mark(sh.wb); paint(); },
    });
    Object.defineProperty(rg, 'Font', { enumerable: true, get() { return fontOf(sh, areas); } });
    Object.defineProperty(rg, 'Interior', { enumerable: true, get() { return interiorOf(sh, areas); } });
    Object.defineProperty(rg, 'Borders', { enumerable: true, get() { return bordersOf(sh, areas); } });
    Object.defineProperty(rg, 'ColumnWidth', {
      enumerable: true,
      get() {
        const c = areas[0].c1;
        const o = sh.cols[c];
        if (o && o.hidden) return 0;
        if (o && o.w != null) return o.w;
        return sh.baseColW == null ? 8 : sh.baseColW;
      },
      set(v) {
        const cols = [];
        for (const a of areas) for (let c = a.c1; c <= a.c2; c++) cols.push(c);
        O.setColWidth(sh, cols, +v, 'Column Width');
        mark(sh.wb); paint();
      },
    });
    Object.defineProperty(rg, 'RowHeight', {
      enumerable: true,
      get() {
        const r = areas[0].r1;
        const row = sh.rows[r];
        if (row && row.hidden) return 0;
        if (row && row.ht != null) return row.ht;
        return M.defaultRowPt(sh);
      },
      set(v) {
        const rows = [];
        for (const a of areas) for (let r = a.r1; r <= a.r2; r++) rows.push(r);
        O.setRowHeight(sh, rows, +v, 'Row Height');
        mark(sh.wb); paint();
      },
    });
    Object.defineProperty(rg, 'Hidden', {
      enumerable: true,
      get() {
        const wholeRow = areas[0].c1 === 0 && areas[0].c2 === MAXC - 1;
        const wholeCol = areas[0].r1 === 0 && areas[0].r2 === MAXR - 1;
        if (wholeCol) return tri(sh.cols[areas[0].c1] && sh.cols[areas[0].c1].hidden);
        if (wholeRow) return tri(sh.rows[areas[0].r1] && sh.rows[areas[0].r1].hidden);
        return 0;
      },
      set(v) { setHidden(sh, areas, on(v)); },
    });
    Object.defineProperty(rg, 'MergeCells', {
      enumerable: true,
      get() { return tri(sh.merges.some((m) => areas.some((a) => M.rangesOverlap(m, a)))); },
    });
    rg.Item = function (a, b) {
      if (b == null) {
        const n = (+a) - 1;
        let i = 0;
        for (const area of areas) {
          for (let r = area.r1; r <= area.r2; r++) for (let c = area.c1; c <= area.c2; c++) {
            if (i === n) return makeRange(sh, [{ r1: r, c1: c, r2: r, c2: c }]);
            i++;
          }
        }
        missing();
      }
      const r = areas[0].r1 + (+a) - 1, c = areas[0].c1 + (+b) - 1;
      if (r < areas[0].r1 || r > areas[0].r2 || c < areas[0].c1 || c > areas[0].c2) missing();
      return makeRange(sh, [{ r1: r, c1: c, r2: r, c2: c }]);
    };
    rg.Cells = function (a, b) {
      if (a == null) return rg;
      return rg.Item(a, b);
    };
    rg.Resize = function (rows, cols) {
      const a = areas[0];
      const r2 = a.r1 + (rows == null ? (a.r2 - a.r1) : (+rows)) - (rows == null ? 0 : 1);
      const c2 = a.c1 + (cols == null ? (a.c2 - a.c1) : (+cols)) - (cols == null ? 0 : 1);
      return makeRange(sh, [{ r1: a.r1, c1: a.c1, r2: r2, c2: c2 }]);
    };
    rg.Offset = function (rowOff, colOff) {
      const dr = rowOff == null ? 0 : +rowOff, dc = colOff == null ? 0 : +colOff;
      return makeRange(sh, areas.map((a) => ({ r1: a.r1 + dr, c1: a.c1 + dc, r2: a.r2 + dr, c2: a.c2 + dc })));
    };
    Object.defineProperty(rg, 'EntireRow', { enumerable: true, get() { return makeRange(sh, areas.map((a) => ({ r1: a.r1, c1: 0, r2: a.r2, c2: MAXC - 1 }))); } });
    Object.defineProperty(rg, 'EntireColumn', { enumerable: true, get() { return makeRange(sh, areas.map((a) => ({ r1: 0, c1: a.c1, r2: MAXR - 1, c2: a.c2 }))); } });
    rg.End = function (dir) { return makeRange(sh, [endCell(sh, areas[0], dir)]); };
    Object.defineProperty(rg, 'CurrentRegion', { enumerable: true, get() { return makeRange(sh, [currentRegion(sh, areas[0].r1, areas[0].c1)]); } });
    rg.SpecialCells = function (type, value) { return specialCells(sh, areas, type, value); };
    rg.Clear = function () { O.clear(sh, areas, 'all'); mark(sh.wb); paint(); };
    rg.ClearContents = function () { O.clear(sh, areas, 'contents'); mark(sh.wb); paint(); };
    rg.ClearFormats = function () { O.clear(sh, areas, 'formats'); mark(sh.wb); paint(); };
    rg.ClearComments = function () { clearNotes(sh, areas); };
    rg.Merge = function () { O.merge(sh, areas); mark(sh.wb); paint(); };
    rg.UnMerge = function () { O.unmerge(sh, areas); mark(sh.wb); paint(); };
    rg.Select = function () { selectAreas(sh, areas); };
    rg.Activate = function () { selectAreas(sh, [areas[0] && { r1: areas[0].r1, c1: areas[0].c1, r2: areas[0].r1, c2: areas[0].c1 }]); };
    rg.Copy = function (dest) {
      if (dest && dest._matheuseRange) { pasteOnto(dest._sheet, dest._areas[0], snapshot(sh, areas), XL.xlPasteAll); return; }
      clip = snapshot(sh, areas);
    };
    rg.PasteSpecial = function (paste) { if (!clip) gap('G-ENGINE', 'PasteSpecial'); pasteOnto(sh, areas[0], clip, paste == null ? XL.xlPasteAll : paste); };
    rg.Insert = function (shift) { shiftCells(sh, areas[0], shift == null ? XL.xlShiftDown : shift, true); };
    rg.Delete = function (shift) {
      const a = areas[0];
      if (a.c1 === 0 && a.c2 >= MAXC - 1) { O.insertLines(sh, 'r', a.r1, -(a.r2 - a.r1 + 1)); mark(sh.wb); paint(); return; }
      if (a.r1 === 0 && a.r2 >= MAXR - 1) { O.insertLines(sh, 'c', a.c1, -(a.c2 - a.c1 + 1)); mark(sh.wb); paint(); return; }
      shiftCells(sh, a, shift == null ? XL.xlShiftUp : shift, false);
    };
    rg.Find = function (what, after, lookIn, lookAt, order, dir, matchCase) {
      return findIn(sh, areas, what, { after: after, lookIn: lookIn, lookAt: lookAt, matchCase: on(matchCase) });
    };
    rg.Replace = function (what, repl, lookAt, order, matchCase) {
      return replaceIn(sh, areas, what, repl, { lookAt: lookAt, matchCase: on(matchCase) });
    };
    rg.Sort = function (key1, order1, key2, type, order2, key3, order3, header) {
      sortRange(sh, areas[0], key1, order1, key2, order2, key3, order3, header);
    };
    rg.AutoFilter = function (field, criteria1) { return autoFilter(sh, areas[0], field, criteria1); };
    rg.AddComment = function (text) { return setNote(sh, areas[0].r1, areas[0].c1, text == null ? '' : String(text)); };
    Object.defineProperty(rg, 'Comment', { enumerable: true, get() { return noteAt(sh, areas[0].r1, areas[0].c1); } });
    Object.defineProperty(rg, 'Hyperlinks', { enumerable: true, get() { return linksOf(sh, areas); } });
    Object.defineProperty(rg, 'Validation', { enumerable: true, get() { return validationOf(sh, areas); } });
    Object.defineProperty(rg, 'FormatConditions', { enumerable: true, get() { return conditionsOf(sh, areas); } });
    Object.defineProperty(rg, 'Areas', {
      enumerable: true,
      get() {
        return collection(() => rg, () => areas, (a) => makeRange(sh, [a]), () => '');
      },
    });
    denyGet(rg, ['SortFields'], 'G-ENGINE');
    return rg;
  }

  function readRaw(sh, r, c) {
    const cell = sh.get(r, c);
    return cell ? cell.v : null;
  }
  function readBlock(sh, areas) {
    if (oneCell(areas)) return readValue(sh, areas[0].r1, areas[0].c1);
    if (tooBig(areas)) gap('G-ENGINE', 'Value');
    const out = [];
    for (const a of areas) {
      for (let r = a.r1; r <= a.r2; r++) {
        const row = [];
        for (let c = a.c1; c <= a.c2; c++) row.push(readValue(sh, r, c));
        out.push(row);
      }
    }
    return out;
  }
  function putBlock(sh, areas, v, asFormula) {
    const a0 = areas[0];
    if (Array.isArray(v)) {
      const rows = v.length;
      const cols = Array.isArray(v[0]) ? v[0].length : 1;
      const h = a0.r2 - a0.r1 + 1, w = a0.c2 - a0.c1 + 1;
      const single = oneCell(areas);
      if (!single && (rows !== h || cols !== w)) {
        const e = new Error('The array and the range are different sizes.');
        e.code = 'missing';
        throw e;
      }
      edit(sh.wb, 'Value', () => {
        for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) {
          const cell = Array.isArray(v[r]) ? v[r][c] : v[r];
          writeScalar(sh, a0.r1 + r, a0.c1 + c, cell);
        }
      });
      return;
    }
    edit(sh.wb, asFormula ? 'Formula' : 'Value', () => {
      eachPos(areas, (r, c) => { if (asFormula) writeFormula(sh, r, c, v, false); else writeScalar(sh, r, c, v); });
    });
  }
  function readFormula(sh, areas, r1c1) {
    const one = (r, c) => {
      const cell = sh.get(r, c);
      if (!cell || cell.f == null) {
        const v = readValue(sh, r, c);
        return v == null ? '' : v;
      }
      if (r1c1) {
        if (!L.xmlss || !L.xmlss.toR1C1) gap('G-ENGINE', 'FormulaR1C1');
        return L.xmlss.toR1C1(cell.f, r, c, sh.wb);
      }
      return '=' + cell.f;
    };
    if (oneCell(areas)) return one(areas[0].r1, areas[0].c1);
    if (tooBig(areas)) gap('G-ENGINE', 'Formula');
    const out = [];
    for (const a of areas) {
      for (let r = a.r1; r <= a.r2; r++) {
        const row = [];
        for (let c = a.c1; c <= a.c2; c++) row.push(one(r, c));
        out.push(row);
      }
    }
    return out;
  }
  function putFormula(sh, areas, v, r1c1) {
    if (Array.isArray(v)) { putBlock(sh, areas, v, false); return; }
    const a0 = areas[0];
    let base = String(v == null ? '' : v);
    if (!r1c1) base = F.toStore(base.replace(/^=/, ''));
    edit(sh.wb, 'Formula', () => {
      eachPos(areas, (r, c) => {
        if (r1c1) writeFormula(sh, r, c, base, true);
        else O.setFormula(sh, r, c, shiftFormula(base, r - a0.r1, c - a0.c1));
      });
    });
  }
  function endCell(sh, a, dir) {
    let r = a.r1, c = a.c1;
    const step = dir === XL.xlUp ? [-1, 0] : dir === XL.xlDown ? [1, 0] : dir === XL.xlToLeft ? [0, -1] : [0, 1];
    const blank = (rr, cc) => {
      const cell = sh.get(rr, cc);
      return !cell || (cell.v == null && cell.f == null);
    };
    if (blank(r, c)) {
      while (r + step[0] >= 0 && c + step[1] >= 0 && r + step[0] < MAXR && c + step[1] < MAXC && blank(r + step[0], c + step[1])) {
        r += step[0]; c += step[1];
        if (step[0] > 0 && r > sh.maxR) break;
        if (step[1] > 0 && c > sh.maxC) break;
      }
      if (r + step[0] >= 0 && c + step[1] >= 0 && r + step[0] < MAXR && c + step[1] < MAXC && !blank(r + step[0], c + step[1])) { r += step[0]; c += step[1]; }
    } else {
      while (r + step[0] >= 0 && c + step[1] >= 0 && r + step[0] < MAXR && c + step[1] < MAXC && !blank(r + step[0], c + step[1])) { r += step[0]; c += step[1]; }
    }
    return { r1: r, c1: c, r2: r, c2: c };
  }
  function currentRegion(sh, r, c) {
    let r1 = r, c1 = c, r2 = r, c2 = c, grew = true;
    const filled = (rr, cc) => { const cell = sh.get(rr, cc); return cell && (cell.v != null || cell.f != null); };
    while (grew) {
      grew = false;
      for (let rr = r1; rr <= r2; rr++) for (let cc = c1; cc <= c2; cc++) {
        for (let dr = -1; dr <= 1; dr++) for (let dc = -1; dc <= 1; dc++) {
          const nr = rr + dr, nc = cc + dc;
          if (nr < 0 || nc < 0 || nr >= MAXR || nc >= MAXC || !filled(nr, nc)) continue;
          if (nr < r1) { r1 = nr; grew = true; }
          if (nc < c1) { c1 = nc; grew = true; }
          if (nr > r2) { r2 = nr; grew = true; }
          if (nc > c2) { c2 = nc; grew = true; }
        }
      }
    }
    return { r1: r1, c1: c1, r2: r2, c2: c2 };
  }
  function specialCells(sh, areas, type, value) {
    const hits = [];
    const used = { r1: 0, c1: 0, r2: Math.max(0, sh.maxR), c2: Math.max(0, sh.maxC) };
    const want = (cell, r, c) => {
      const blank = !cell || (cell.v == null && cell.f == null);
      if (type === XL.xlCellTypeBlanks) return blank;
      if (type === XL.xlCellTypeFormulas) return !!(cell && cell.f != null);
      if (type === XL.xlCellTypeComments) return sh.comments && sh.comments.has(M.key(r, c));
      if (type === XL.xlCellTypeLastCell) return false;
      if (type === XL.xlCellTypeVisible) {
        const row = sh.rows[r], col = sh.cols[c];
        return !(row && row.hidden) && !(col && col.hidden);
      }
      if (type === XL.xlCellTypeConstants) {
        if (!cell || cell.f != null || cell.v == null) return false;
        if (value == null) return true;
        const bits = +value;
        const v = cell.v;
        if ((bits & XL.xlNumbers) && typeof v === 'number') return true;
        if ((bits & XL.xlTextValues) && typeof v === 'string' && !(M.isErr && M.isErr(v))) return true;
        if ((bits & XL.xlLogical) && typeof v === 'boolean') return true;
        if ((bits & XL.xlErrors) && M.isErr && M.isErr(v)) return true;
        return false;
      }
      gap('G-ENGINE', 'SpecialCells');
    };
    if (type === XL.xlCellTypeLastCell) {
      const end = sh.contentEnd ? sh.contentEnd() : { r: Math.max(0, sh.maxR), c: Math.max(0, sh.maxC) };
      return makeRange(sh, [{ r1: end.r, c1: end.c, r2: end.r, c2: end.c }]);
    }
    for (const a of areas) {
      const r2 = Math.min(a.r2, used.r2), c2 = Math.min(a.c2, used.c2);
      for (let r = a.r1; r <= r2; r++) for (let c = a.c1; c <= c2; c++) {
        if (want(sh.get(r, c), r, c)) hits.push({ r1: r, c1: c, r2: r, c2: c });
      }
    }
    if (!hits.length) missing();
    return makeRange(sh, hits.length > 64 ? [hits[0]] : hits);
  }
  function selectAreas(sh, areas) {
    const a = areas[0];
    sh.wb.active = sh.index;
    sh.view.sel = { r: a.r1, c: a.c1, ranges: areas.map((x) => Object.assign({}, x)) };
    if (sh._vs && sh._vs.sel) {
      sh._vs.sel.r = a.r1; sh._vs.sel.c = a.c1;
      sh._vs.sel.ranges = areas.map((x) => Object.assign({}, x));
      sh._vs.sel.active = 0;
      sh._vs.sel.anchor = { r: a.r1, c: a.c1 };
    }
    if (L.grid && L.grid.wb === sh.wb && L.grid.select) {
      L.grid.select(a.r1, a.c1, { noScroll: false });
      if (L.grid.sel && (a.r2 !== a.r1 || a.c2 !== a.c1 || areas.length > 1)) {
        const s = L.grid.sel();
        s.ranges = areas.map((x) => Object.assign({}, x));
        s.r = a.r1; s.c = a.c1; s.active = 0;
      }
    }
    paint();
  }
  function selectionOf(sh) {
    if (L.grid && L.grid.wb === sh.wb && L.grid.sheet && L.grid.sheet() === sh && L.grid.ranges) {
      const list = L.grid.ranges();
      if (list && list.length) return makeRange(sh, list.map((x) => Object.assign({}, x)));
    }
    const s = (sh._vs && sh._vs.sel) || sh.view.sel;
    if (s && s.ranges && s.ranges.length) return makeRange(sh, s.ranges.map((x) => Object.assign({}, x)));
    return makeRange(sh, [{ r1: 0, c1: 0, r2: 0, c2: 0 }]);
  }
  function setHidden(sh, areas, hide) {
    edit(sh.wb, hide ? 'Hide' : 'Unhide', () => {
      for (const a of areas) {
        if (a.c1 === 0 && a.c2 >= MAXC - 1) {
          for (let r = a.r1; r <= a.r2; r++) { const row = sh.rowObj(r); row.hidden = hide; }
        } else if (a.r1 === 0 && a.r2 >= MAXR - 1) {
          for (let c = a.c1; c <= a.c2; c++) { const col = sh.colObj(c); col.hidden = hide; if (hide) col.w = 0; }
        }
      }
      if (L.layout && L.layout.invalidate) L.layout.invalidate(sh);
    });
  }
  function shiftCells(sh, a, shift, insert) {
    const down = shift === XL.xlShiftDown || shift === XL.xlDown;
    const right = shift === XL.xlShiftToRight || shift === XL.xlToRight;
    const up = shift === XL.xlShiftUp || shift === XL.xlUp;
    const left = shift === XL.xlShiftToLeft || shift === XL.xlToLeft;
    const dir = down ? 'down' : right ? 'right' : up ? 'up' : left ? 'left' : (insert ? 'down' : 'up');
    O.shiftCells(sh, a, dir, insert);
    mark(sh.wb); paint();
  }
  function snapshot(sh, areas) {
    const a = areas[0];
    const rows = [];
    for (let r = a.r1; r <= a.r2; r++) {
      const row = [];
      for (let c = a.c1; c <= a.c2; c++) {
        const cell = sh.get(r, c);
        row.push(cell ? { v: cell.v, f: cell.f, s: cell.s } : null);
      }
      rows.push(row);
    }
    const notes = [];
    if (sh.comments) sh.comments.forEach((cm) => {
      if (cm.r >= a.r1 && cm.r <= a.r2 && cm.c >= a.c1 && cm.c <= a.c2) notes.push(Object.assign({}, cm));
    });
    return { rows: rows, notes: notes, h: a.r2 - a.r1 + 1, w: a.c2 - a.c1 + 1 };
  }
  function pasteOnto(sh, dest, snap, paste) {
    edit(sh.wb, 'Paste', () => {
      for (let r = 0; r < snap.h; r++) for (let c = 0; c < snap.w; c++) {
        const src = snap.rows[r][c];
        const rr = dest.r1 + r, cc = dest.c1 + c;
        if (paste === XL.xlPasteFormats) {
          if (src && src.s) O.format(sh, [{ r1: rr, c1: cc, r2: rr, c2: cc }], () => null);
          if (src && src.s != null) {
            const cell = sh.get(rr, cc);
            const n = cell ? Object.assign({}, cell) : { v: null };
            n.s = src.s;
            O.put(sh, rr, cc, n);
          }
          continue;
        }
        if (paste === XL.xlPasteComments) continue;
        if (paste === XL.xlPasteFormulas && src && src.f != null) O.setFormula(sh, rr, cc, src.f);
        else if (paste === XL.xlPasteValues || !src || src.f == null) O.setValue(sh, rr, cc, src ? src.v : null, true);
        else O.setFormula(sh, rr, cc, src.f);
        if (paste === XL.xlPasteAll && src && src.s != null) {
          const cell = sh.get(rr, cc) || { v: null };
          cell.s = src.s;
          O.put(sh, rr, cc, cell);
        }
      }
      if ((paste === XL.xlPasteAll || paste === XL.xlPasteComments) && snap.notes) {
        for (const cm of snap.notes) {
          const r = dest.r1 + (cm.r - (snap.rows.length ? 0 : 0));
          void r;
        }
      }
    });
  }
  function findIn(sh, areas, what, opt) {
    const needle = String(what == null ? '' : what);
    const whole = opt.lookAt === XL.xlWhole;
    const formulas = opt.lookIn === XL.xlPasteFormulas || opt.lookIn === XL.xlCellTypeFormulas;
    const fold = (s) => (opt.matchCase ? String(s) : String(s).toLowerCase());
    const n = fold(needle);
    for (const a of areas) {
      for (let r = a.r1; r <= Math.min(a.r2, Math.max(a.r1, sh.maxR)); r++) {
        for (let c = a.c1; c <= Math.min(a.c2, Math.max(a.c1, sh.maxC)); c++) {
          const cell = sh.get(r, c);
          if (!cell) continue;
          const hay = formulas ? (cell.f || '') : (cell.v == null ? '' : cell.v);
          const h = fold(hay);
          const hit = whole ? h === n : String(h).indexOf(n) >= 0;
          if (hit && n !== '') return makeRange(sh, [{ r1: r, c1: c, r2: r, c2: c }]);
        }
      }
    }
    return null;
  }
  function replaceIn(sh, areas, what, repl, opt) {
    const needle = String(what == null ? '' : what);
    if (!needle) return 0;
    let n = 0;
    edit(sh.wb, 'Replace', () => {
      for (const a of areas) {
        sh.each(a.r1, a.c1, a.r2, a.c2, (cell, r, c) => {
          const src = cell.f != null ? cell.f : (cell.v == null ? '' : String(cell.v));
          const flags = opt.matchCase ? 'g' : 'gi';
          const next = opt.lookAt === XL.xlWhole
            ? (opt.matchCase ? src === needle : String(src).toLowerCase() === needle.toLowerCase()) ? String(repl) : src
            : String(src).replace(new RegExp(needle.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), flags), String(repl));
          if (next !== src) {
            if (cell.f != null) O.setFormula(sh, r, c, next);
            else O.setValue(sh, r, c, next);
            n++;
          }
        });
      }
    });
    return n === 0 ? 0 : -1;
  }
  function sortRange(sh, a, key1, order1, key2, order2, key3, order3, header) {
    const keys = [];
    const add = (key, order) => {
      if (!key || !key._matheuseRange) return;
      const col = key._areas[0].c1 - a.c1;
      if (col < 0) return;
      keys.push({ index: col, desc: order === XL.xlDescending });
    };
    add(key1, order1); add(key2, order2); add(key3, order3);
    if (!keys.length) keys.push({ index: 0, desc: false });
    O.sort(sh, a, keys, { header: header === XL.xlYes, byRows: true });
    mark(sh.wb); paint();
  }
  function autoFilter(sh, a, field, criteria1) {
    if (field == null) {
      const onNow = sh.autoFilter && sh.autoFilter.ref;
      edit(sh.wb, 'AutoFilter', () => { O.setAutoFilter(sh, onNow ? null : { ref: a, cols: [] }); });
      return tri(!onNow);
    }
    const cols = (sh.autoFilter && sh.autoFilter.cols ? sh.autoFilter.cols.slice() : []);
    cols[(+field) - 1] = { values: [String(criteria1)] };
    edit(sh.wb, 'AutoFilter', () => {
      O.setAutoFilter(sh, { ref: (sh.autoFilter && sh.autoFilter.ref) || a, cols: cols });
      if (L.filter && L.filter.applyRows) L.filter.applyRows(sh, a, sh.autoFilter);
    });
    return true;
  }

  function fontOf(sh, areas) {
    const o = stamp({}, () => makeRange(sh, areas));
    const st = () => styleAt(sh, areas[0].r1, areas[0].c1);
    const setFont = (delta) => { O.format(sh, areas, { font: delta }, 'Font'); mark(sh.wb); paint(); };
    defFont(o, 'Name', () => (st().font && st().font.name) || '', (v) => setFont({ name: String(v) }));
    defFont(o, 'Size', () => (st().font && st().font.sz) || 10, (v) => setFont({ sz: +v }));
    defFont(o, 'Bold', () => tri(st().font && st().font.b), (v) => setFont({ b: on(v) || undefined }));
    defFont(o, 'Italic', () => tri(st().font && st().font.i), (v) => setFont({ i: on(v) || undefined }));
    defFont(o, 'Underline', () => (st().font && st().font.u ? XL.xlUnderlineStyleSingle : XL.xlUnderlineStyleNone), (v) => setFont({ u: (v === XL.xlUnderlineStyleNone || v === 0 || v === false) ? undefined : 'single' }));
    defFont(o, 'Color', () => hexToBgr(colorHex(sh.wb, st().font && st().font.color) || '#000000'), (v) => setFont({ color: { rgb: 'FF' + bgrToHex(v).slice(1) } }));
    defFont(o, 'ColorIndex', () => colorIndex(sh.wb, st().font && st().font.color), (v) => setFont({ color: indexColor(sh.wb, v) }));
    defFont(o, 'ThemeColor', () => themeEnum(st().font && st().font.color), (v) => setFont({ color: { theme: THEME_TO[+v] } }));
    defFont(o, 'TintAndShade', () => (st().font && st().font.color && st().font.color.tint) || 0, (v) => {
      const cur = (st().font && st().font.color) || { theme: 1 };
      setFont({ color: Object.assign({}, cur, { tint: +v }) });
    });
    return o;
  }
  function defFont(o, name, get, set) {
    Object.defineProperty(o, name, { enumerable: true, configurable: true, get: get, set: set });
  }
  function colorHex(wb, c) { return M.colorHex(wb, c, '#000000'); }
  function colorIndex(wb, c) {
    if (!c) return XL.xlColorIndexAutomatic;
    if (c.indexed != null) return c.indexed;
    return XL.xlColorIndexNone;
  }
  function indexColor(wb, v) {
    if (v === XL.xlColorIndexNone || v === XL.xlColorIndexAutomatic) return null;
    return { indexed: +v };
  }
  function themeEnum(c) {
    if (!c || c.theme == null) return 0;
    return THEME_FROM[c.theme] || 0;
  }
  function interiorOf(sh, areas) {
    const o = stamp({}, () => makeRange(sh, areas));
    const fill = () => styleAt(sh, areas[0].r1, areas[0].c1).fill;
    const setFill = (delta) => { O.format(sh, areas, { fill: delta }, 'Fill'); mark(sh.wb); paint(); };
    defFont(o, 'Color', () => {
      const f = fill();
      if (!f || !f.fg) return hexToBgr('#FFFFFF');
      return hexToBgr(colorHex(sh.wb, f.fg) || '#FFFFFF');
    }, (v) => setFill({ pattern: 'solid', fg: { rgb: 'FF' + bgrToHex(v).slice(1) } }));
    defFont(o, 'Pattern', () => (fill() && fill().pattern && fill().pattern !== 'none' ? XL.xlPatternSolid : XL.xlPatternNone), (v) => {
      if (v === XL.xlPatternNone || v === XL.xlNone) setFill(null);
      else setFill(Object.assign({ pattern: 'solid' }, fill() || {}));
    });
    defFont(o, 'ColorIndex', () => colorIndex(sh.wb, fill() && fill().fg), (v) => setFill({ pattern: 'solid', fg: indexColor(sh.wb, v) }));
    defFont(o, 'ThemeColor', () => themeEnum(fill() && fill().fg), (v) => setFill({ pattern: 'solid', fg: { theme: THEME_TO[+v] } }));
    defFont(o, 'TintAndShade', () => (fill() && fill().fg && fill().fg.tint) || 0, (v) => {
      const fg = Object.assign({}, (fill() && fill().fg) || { theme: 0 }, { tint: +v });
      setFill({ pattern: 'solid', fg: fg });
    });
    defFont(o, 'PatternColor', () => hexToBgr(colorHex(sh.wb, fill() && fill().bg) || '#FFFFFF'), (v) => setFill(Object.assign({}, fill() || { pattern: 'solid' }, { bg: { rgb: 'FF' + bgrToHex(v).slice(1) } })));
    return o;
  }
  const EDGE = { 7: 'l', 8: 't', 9: 'b', 10: 'r', 11: 'insideV', 12: 'insideH' };
  function bordersOf(sh, areas) {
    const o = stamp({}, () => makeRange(sh, areas));
    o.Item = function (index) { return borderEdge(sh, areas, EDGE[index] || 'l'); };
    o.LineStyle = XL.xlNone;
    Object.defineProperty(o, 'LineStyle', {
      enumerable: true,
      get() { return borderEdge(sh, areas, 'l').LineStyle; },
      set(v) {
        if (v === XL.xlNone || v === XL.xlLineStyleNone) O.borders(sh, areas, { none: true });
        else O.borders(sh, areas, { all: true, line: lineOf(v, null) });
        mark(sh.wb); paint();
      },
    });
    return o;
  }
  function lineOf(style, color) {
    const name = style === XL.xlDouble ? 'double' : style === XL.xlDash ? 'dashed' : style === XL.xlDot ? 'dotted' : style === XL.xlMedium ? 'medium' : style === XL.xlThick ? 'thick' : 'thin';
    const line = { style: name };
    if (color != null) line.color = { rgb: 'FF' + bgrToHex(color).slice(1) };
    return line;
  }
  function borderEdge(sh, areas, edge) {
    const o = stamp({}, () => bordersOf(sh, areas));
    const key = edge === 'insideV' || edge === 'insideH' ? null : edge;
    Object.defineProperty(o, 'LineStyle', {
      enumerable: true,
      get() {
        if (!key) return XL.xlNone;
        const b = styleAt(sh, areas[0].r1, areas[0].c1).border;
        return b && b[key] ? XL.xlContinuous : XL.xlNone;
      },
      set(v) {
        const spec = {};
        const line = (v === XL.xlNone || v === XL.xlLineStyleNone) ? null : lineOf(v, null);
        if (edge === 'insideH') spec.insideH = line;
        else if (edge === 'insideV') spec.insideV = line;
        else spec[edge === 'l' ? 'left' : edge === 'r' ? 'right' : edge === 't' ? 'top' : 'bottom'] = line;
        O.borders(sh, areas, spec);
        mark(sh.wb); paint();
      },
    });
    Object.defineProperty(o, 'Weight', {
      enumerable: true,
      get() { return XL.xlThin; },
      set(v) { o.LineStyle = v === XL.xlMedium || v === XL.xlThick ? v : XL.xlContinuous; },
    });
    Object.defineProperty(o, 'Color', {
      enumerable: true,
      get() {
        const b = styleAt(sh, areas[0].r1, areas[0].c1).border;
        const line = key && b && b[key];
        return hexToBgr(colorHex(sh.wb, line && line.color) || '#000000');
      },
      set(v) {
        const spec = {};
        const line = lineOf(XL.xlContinuous, v);
        if (edge === 'l' || edge === 'r' || edge === 't' || edge === 'b') spec[edge === 'l' ? 'left' : edge === 'r' ? 'right' : edge === 't' ? 'top' : 'bottom'] = line;
        O.borders(sh, areas, spec);
        mark(sh.wb); paint();
      },
    });
    return o;
  }
  function rowColCollection(sh, areas, axis) {
    const list = [];
    for (const a of areas) {
      if (axis === 'r') for (let r = a.r1; r <= a.r2; r++) list.push(r);
      else for (let c = a.c1; c <= a.c2; c++) list.push(c);
    }
    const uniq = [];
    const seen = new Set();
    for (const n of list) if (!seen.has(n)) { seen.add(n); uniq.push(n); }
    return collection(() => makeRange(sh, areas), () => uniq, (n) => {
      if (axis === 'r') return makeRange(sh, [{ r1: n, c1: 0, r2: n, c2: MAXC - 1 }]);
      return makeRange(sh, [{ r1: 0, c1: n, r2: MAXR - 1, c2: n }]);
    }, (n) => String(n + 1));
  }

  function noteAt(sh, r, c) {
    const cm = sh.comments && sh.comments.get(M.key(r, c));
    if (!cm) return null;
    const o = stamp({}, () => makeRange(sh, [{ r1: r, c1: c, r2: r, c2: c }]));
    Object.defineProperty(o, 'Text', {
      enumerable: true,
      get() { return cm.text || ''; },
      set(v) { setNote(sh, r, c, String(v), cm.author); },
    });
    Object.defineProperty(o, 'Author', { enumerable: true, get() { return cm.author || ''; } });
    o.Delete = function () { clearOneNote(sh, r, c); };
    denyFn(o, ['Next', 'Previous'], 'G-ENGINE');
    return o;
  }
  function setNote(sh, r, c, text, author) {
    const k = M.key(r, c);
    edit(sh.wb, 'Comment', () => {
      const old = sh.comments.get(k) || null;
      const next = { r: r, c: c, author: author || (old && old.author) || '', text: text, visible: false };
      const apply = (v) => { if (v) sh.comments.set(k, v); else sh.comments.delete(k); };
      apply(next);
      sh.wb.undo.op(() => apply(old), () => apply(next));
    });
    return noteAt(sh, r, c);
  }
  function clearOneNote(sh, r, c) {
    const k = M.key(r, c);
    edit(sh.wb, 'Comment', () => {
      const old = sh.comments.get(k) || null;
      sh.comments.delete(k);
      sh.wb.undo.op(() => { if (old) sh.comments.set(k, old); }, () => sh.comments.delete(k));
    });
  }
  function clearNotes(sh, areas) {
    edit(sh.wb, 'Comment', () => {
      const gone = [];
      sh.comments.forEach((cm, k) => {
        if (areas.some((a) => cm.r >= a.r1 && cm.r <= a.r2 && cm.c >= a.c1 && cm.c <= a.c2)) gone.push([k, cm]);
      });
      for (const [k] of gone) sh.comments.delete(k);
      sh.wb.undo.op(() => { for (const [k, cm] of gone) sh.comments.set(k, cm); }, () => { for (const [k] of gone) sh.comments.delete(k); });
    });
  }
  function commentsOf(sh) {
    const list = () => {
      const out = [];
      if (sh.comments) sh.comments.forEach((cm) => out.push(cm));
      out.sort((a, b) => a.r - b.r || a.c - b.c);
      return out;
    };
    const col = collection(() => sheetObj(sh), list, (cm) => noteAt(sh, cm.r, cm.c), (cm) => F.cellName(cm.r, cm.c));
    col.Add = function (range, text) {
      const rg = range && range._matheuseRange ? range : null;
      if (!rg) missing();
      return setNote(rg._sheet, rg._areas[0].r1, rg._areas[0].c1, text == null ? '' : String(text));
    };
    return col;
  }
  function linksOf(sh, areas) {
    const list = () => (sh.links || []).filter((h) => !areas || areas.some((a) => M.rangesOverlap(h.ref, a)));
    const col = collection(() => sheetObj(sh), list, (h) => hyperlink(sh, h), (h) => h.target || h.location || '');
    col.Add = function (anchor, address, subAddress) {
      const rg = anchor && anchor._matheuseRange ? anchor : makeRange(sh, areas || [{ r1: 0, c1: 0, r2: 0, c2: 0 }]);
      const ref = Object.assign({}, rg._areas[0]);
      const item = { ref: ref, target: address || undefined, location: subAddress || undefined };
      edit(sh.wb, 'Hyperlink', () => { O.setLinks(sh, (sh.links || []).concat([item])); });
      return hyperlink(sh, item);
    };
    col.Delete = function () {
      edit(sh.wb, 'Hyperlink', () => {
        O.setLinks(sh, (sh.links || []).filter((h) => !list().includes(h)));
      });
    };
    return col;
  }
  function hyperlink(sh, h) {
    const o = stamp({}, () => sheetObj(sh));
    Object.defineProperty(o, 'Address', { enumerable: true, get() { return h.target || ''; }, set(v) { h.target = String(v); mark(sh.wb); } });
    Object.defineProperty(o, 'SubAddress', { enumerable: true, get() { return h.location || ''; }, set(v) { h.location = String(v); mark(sh.wb); } });
    Object.defineProperty(o, 'Range', { enumerable: true, get() { return makeRange(sh, [h.ref]); } });
    Object.defineProperty(o, 'TextToDisplay', { enumerable: true, get() { return h.display || ''; }, set(v) { h.display = String(v); mark(sh.wb); } });
    o.Delete = function () {
      edit(sh.wb, 'Hyperlink', () => { O.setLinks(sh, (sh.links || []).filter((x) => x !== h)); });
    };
    denyFn(o, ['Follow', 'CreateNewDocument'], 'G-HOST');
    return o;
  }
  function validationOf(sh, areas) {
    const o = stamp({}, () => makeRange(sh, areas));
    o.Add = function (type, alertStyle, operator, formula1) {
      const kind = type === XL.xlValidateList ? 'list' : type === XL.xlValidateCustom ? 'custom' : type === XL.xlValidateDate ? 'date' : type === XL.xlValidateWholeNumber ? 'whole' : type === XL.xlValidateDecimal ? 'decimal' : type === XL.xlValidateTextLength ? 'textLength' : 'custom';
      const rule = { ranges: areas.map((a) => Object.assign({}, a)), type: kind, formula1: formula1 == null ? '' : String(formula1), operator: operator || undefined };
      edit(sh.wb, 'Validation', () => { O.setDV(sh, (sh.dv || []).concat([rule])); });
    };
    o.Delete = function () {
      edit(sh.wb, 'Validation', () => {
        O.setDV(sh, (sh.dv || []).filter((d) => !(d.ranges || []).some((r) => areas.some((a) => M.rangesOverlap(r, a)))));
      });
    };
    denyFn(o, ['Modify'], 'G-ENGINE');
    return o;
  }
  const OP = { 1: 'between', 2: 'notBetween', 3: 'equal', 4: 'notEqual', 5: 'greaterThan', 6: 'lessThan', 7: 'greaterThanOrEqual', 8: 'lessThanOrEqual' };
  function conditionsOf(sh, areas) {
    const rules = () => {
      const out = [];
      for (const cf of sh.cf || []) {
        if (!(cf.ranges || []).some((r) => areas.some((a) => M.rangesOverlap(r, a)))) continue;
        for (const rule of cf.rules || []) {
          if (rule.type === 'cellIs' || rule.type === 'expression') out.push({ cf: cf, rule: rule });
        }
      }
      return out;
    };
    const col = collection(() => makeRange(sh, areas), rules, (x) => condition(sh, x), () => '');
    col.Add = function (type, operator, formula1, formula2) {
      const dxf = sh.wb.dxfs.length;
      sh.wb.dxfs.push({ font: { b: true } });
      const rule = type === XL.xlExpression
        ? { type: 'expression', f: [String(formula1 == null ? '' : formula1)], dxf: dxf }
        : { type: 'cellIs', op: OP[operator] || 'greaterThan', f: [String(formula1 == null ? '' : formula1), formula2 == null ? undefined : String(formula2)].filter((x) => x != null), dxf: dxf };
      const block = { ranges: areas.map((a) => Object.assign({}, a)), rules: [rule] };
      edit(sh.wb, 'Conditional Formatting', () => { O.setCF(sh, (sh.cf || []).concat([block])); });
      return condition(sh, { cf: block, rule: rule });
    };
    col.Delete = function () {
      edit(sh.wb, 'Conditional Formatting', () => {
        O.setCF(sh, (sh.cf || []).filter((cf) => !(cf.ranges || []).some((r) => areas.some((a) => M.rangesOverlap(r, a))) || (cf.rules || []).some((rule) => rule.type !== 'cellIs' && rule.type !== 'expression')));
      });
    };
    return col;
  }
  function condition(sh, x) {
    const o = stamp({}, () => sheetObj(sh));
    Object.defineProperty(o, 'Type', { enumerable: true, get() { return x.rule.type === 'expression' ? XL.xlExpression : XL.xlCellValue; } });
    Object.defineProperty(o, 'Formula1', { enumerable: true, get() { return (x.rule.f && x.rule.f[0]) || ''; } });
    o.Delete = function () {
      edit(sh.wb, 'Conditional Formatting', () => {
        x.cf.rules = (x.cf.rules || []).filter((r) => r !== x.rule);
        O.setCF(sh, (sh.cf || []).filter((cf) => cf !== x.cf || (cf.rules && cf.rules.length)));
      });
    };
    denyGet(o, ['AppliesTo'], 'G-ENGINE');
    return o;
  }

  const sheets = new Map();
  function sheetObj(sh) {
    let o = sheets.get(sh);
    if (o) return o;
    o = stamp({}, () => bookObj(sh.wb));
    o._sheet = sh;
    sheets.set(sh, o);
    Object.defineProperty(o, 'Name', {
      enumerable: true,
      get() { return sh.name; },
      set(v) { O.renameSheet(sh.wb, sh, String(v)); mark(sh.wb); paint(); },
    });
    Object.defineProperty(o, 'Index', { enumerable: true, get() { return sh.index + 1; } });
    Object.defineProperty(o, 'Visible', {
      enumerable: true,
      get() { return sh.state === 'veryHidden' ? XL.xlSheetVeryHidden : sh.state === 'hidden' ? XL.xlSheetHidden : XL.xlSheetVisible; },
      set(v) {
        const state = v === XL.xlSheetVeryHidden ? 'veryHidden' : (v === XL.xlSheetHidden || v === 0 || v === false) ? 'hidden' : 'visible';
        edit(sh.wb, 'Sheet', () => { sh.state = state; });
      },
    });
    Object.defineProperty(o, 'UsedRange', {
      enumerable: true,
      get() {
        if (sh.maxR < 0) return makeRange(sh, [{ r1: 0, c1: 0, r2: 0, c2: 0 }]);
        return makeRange(sh, [{ r1: 0, c1: 0, r2: sh.maxR, c2: sh.maxC }]);
      },
    });
    o.Range = function (a, b) {
      if (a && a._matheuseRange) {
        const a1 = a._areas[0], b1 = b && b._matheuseRange ? b._areas[0] : a1;
        return makeRange(sh, [{ r1: Math.min(a1.r1, b1.r1), c1: Math.min(a1.c1, b1.c1), r2: Math.max(a1.r2, b1.r2), c2: Math.max(a1.c2, b1.c2) }]);
      }
      if (b != null && typeof a === 'string') return makeRange(sh, [unionAddr(a, b)]);
      if (typeof a === 'string') return makeRange(sh, parseAreas(a));
      missing();
    };
    o.Cells = function (r, c) {
      if (r == null) return makeRange(sh, [{ r1: 0, c1: 0, r2: MAXR - 1, c2: MAXC - 1 }]);
      if (c == null) {
        const n = (+r) - 1;
        const rr = Math.floor(n / MAXC), cc = n % MAXC;
        return makeRange(sh, [{ r1: rr, c1: cc, r2: rr, c2: cc }]);
      }
      if (typeof c === 'string') c = F.colIndex(c) + 1;
      return makeRange(sh, [{ r1: (+r) - 1, c1: (+c) - 1, r2: (+r) - 1, c2: (+c) - 1 }]);
    };
    Object.defineProperty(o, 'Rows', { enumerable: true, get() { return o.Cells().Rows; } });
    Object.defineProperty(o, 'Columns', { enumerable: true, get() { return o.Cells().Columns; } });
    Object.defineProperty(o, 'Comments', { enumerable: true, get() { return commentsOf(sh); } });
    Object.defineProperty(o, 'Hyperlinks', { enumerable: true, get() { return linksOf(sh, null); } });
    Object.defineProperty(o, 'Names', { enumerable: true, get() { return namesOf(sh.wb, sh.index); } });
    Object.defineProperty(o, 'ListObjects', { enumerable: true, get() { return listObjects(sh); } });
    Object.defineProperty(o, 'ChartObjects', { enumerable: true, get() { return chartObjects(sh); } });
    Object.defineProperty(o, 'Shapes', { enumerable: true, get() { return shapesOf(sh); } });
    Object.defineProperty(o, 'PivotTables', { enumerable: true, get() { return pivotTables(sh); } });
    Object.defineProperty(o, 'PageSetup', { enumerable: true, get() { return pageSetup(sh); } });
    Object.defineProperty(o, 'HPageBreaks', { enumerable: true, get() { return pageBreaks(sh, 'rowBreaks'); } });
    Object.defineProperty(o, 'VPageBreaks', { enumerable: true, get() { return pageBreaks(sh, 'colBreaks'); } });
    Object.defineProperty(o, 'Outline', { enumerable: true, get() { return outlineOf(sh); } });
    Object.defineProperty(o, 'ProtectionMode', { enumerable: true, get() { return tri(!!sh.protection); } });
    Object.defineProperty(o, 'AutoFilter', { enumerable: true, get() { return sh.autoFilter ? autoFilterObj(sh) : null; } });
    Object.defineProperty(o, 'Scenarios', { enumerable: true, get() { return scenariosOf(sh); } });
    o.Activate = function () { sh.wb.active = sh.index; paint(); };
    o.Select = function () { o.Activate(); };
    o.Delete = function () { O.deleteSheet(sh.wb, sh); sheets.delete(sh); mark(sh.wb); paint(); };
    o.Copy = function (before, after) { copySheet(sh, before, after); };
    o.Move = function (before, after) { moveSheet(sh, before, after); };
    o.Calculate = function () { if (L.calc && L.calc.recalcAll) L.calc.recalcAll(sh.wb); paint(); };
    o.Protect = function (password) { protectSheet(sh, password); };
    o.Unprotect = function (password) { unprotectSheet(sh, password); };
    denyGet(o, ['QueryTables', 'XmlMaps', 'OLEObjects', 'CommentsThreaded'], 'G-ENGINE');
    function sheetBytes(fileFormat) {
      if (fileFormat === 6 || fileFormat === 'csv') return L.csv.write(sh.wb, sh, ',');
      if (fileFormat === -4158 || fileFormat === 'txt') return L.csv.write(sh.wb, sh, '\t');
      if (fileFormat == null || fileFormat === '' || fileFormat === 51 || fileFormat === 'xlsx') {
        return L.xlsxWrite.write(Object.assign({}, sh.wb, { sheets: [sh], active: 0 }), { type: 'xlsx' });
      }
      const e = new Error('The parameter FileFormat is not available.');
      e.code = 'error';
      throw e;
    }
    o.PrintOut = function () {
      if (!L.hostAudit || !L.hostAudit.has('Print')) gap('G-HOST', 'PrintOut');
      return L.hostAudit.printPdf(() => L.print.pdfBytes({ sheet: sh }));
    };
    o.ExportAsFixedFormat = function (format, fileName) {
      if (!L.hostAudit || !L.hostAudit.has('FixedFormat')) gap('G-HOST', 'ExportAsFixedFormat');
      return L.hostAudit.exportPdf(format, fileName, () => L.print.pdfBytes({ sheet: sh }));
    };
    o.SaveAs = function (fileName, fileFormat) {
      if (!L.hostAudit || !L.hostAudit.has('SheetSaveAs')) gap('G-HOST', 'SaveAs');
      L.hostAudit.rejectExtra(arguments, 2, ['Filename', 'FileFormat', 'Password', 'WriteResPassword', 'ReadOnlyRecommended', 'CreateBackup']);
      return L.hostAudit.saveCopy(fileName, () => sheetBytes(fileFormat), 'SheetSaveAs');
    };
    return o;
  }
  function unionAddr(a, b) {
    const pa = F.parseRange(a), pb = F.parseRange(b);
    if (!pa || !pb) missing();
    return { r1: Math.min(pa.r1, pb.r1), c1: Math.min(pa.c1, pb.c1), r2: Math.max(pa.r2, pb.r2), c2: Math.max(pa.c2, pb.c2) };
  }
  function sheetList(wb, kind) {
    return wb.sheets.filter((s) => {
      if (kind === 'worksheet') return s.kind === 'worksheet' || !s.kind;
      if (kind === 'chart') return s.kind === 'chart';
      return true;
    });
  }
  function sheetsCol(wb, kind) {
    const col = collection(() => bookObj(wb), () => sheetList(wb, kind), (s) => sheetObj(s), (s) => s.name);
    col.Add = function (before, after, count) {
      const n = count == null ? 1 : +count;
      let at = wb.sheets.length;
      const anchor = (before && before._sheet) || (after && after._sheet);
      if (before && before._sheet) at = before._sheet.index;
      else if (after && after._sheet) at = after._sheet.index + 1;
      void anchor;
      let made = null;
      for (let i = 0; i < n; i++) made = O.addSheet(wb, null, at);
      mark(wb); paint();
      return sheetObj(made);
    };
    return col;
  }
  function copySheet(sh, before, after) {
    const at = before && before._sheet ? before._sheet.index : after && after._sheet ? after._sheet.index + 1 : sh.wb.sheets.length;
    let copy;
    O.tx(sh.wb, 'Copy Sheet', () => {
      copy = sh.wb.addSheet(sh.wb.nextSheetName(sh.name), at, sh.kind);
      copy.rows = sh.rows.map((row) => row && { cells: row.cells.slice(), ht: row.ht, hidden: row.hidden, s: row.s, level: row.level });
      copy.cols = sh.cols.slice();
      copy.maxR = sh.maxR; copy.maxC = sh.maxC;
    });
    mark(sh.wb); paint();
    return sheetObj(copy);
  }
  function moveSheet(sh, before, after) {
    let to = before && before._sheet ? before._sheet.index : after && after._sheet ? after._sheet.index + 1 : 0;
    O.moveSheet(sh.wb, sh, to);
    mark(sh.wb); paint();
  }
  function hashPw(pw) {
    if (L.dlg && L.dlg.legacyHash) return L.dlg.legacyHash(pw);
    return String(pw || '');
  }
  function protectSheet(sh, password) {
    const attrs = { sheet: '1' };
    if (password) attrs.password = hashPw(password);
    edit(sh.wb, 'Protect', () => { O.setProtection(sh, { attrs: attrs }); });
  }
  function unprotectSheet(sh, password) {
    if (!sh.protection) return;
    const attrs = sh.protection.attrs || {};
    if (attrs.password && hashPw(password || '') !== attrs.password && String(password || '') !== attrs.password) {
      const e = new Error('The password is not correct.');
      e.code = 'missing';
      throw e;
    }
    edit(sh.wb, 'Protect', () => { O.setProtection(sh, null); });
  }
  function pageSetup(sh) {
    const o = stamp({}, () => sheetObj(sh));
    const p = () => sh.print;
    const set = (fn) => { edit(sh.wb, 'Page Setup', () => { const next = JSON.parse(JSON.stringify(p())); fn(next); O.setPrint(sh, next); }); };
    defFont(o, 'Orientation', () => (p().orientation === 'landscape' ? XL.xlLandscape : XL.xlPortrait), (v) => set((n) => { n.orientation = v === XL.xlLandscape ? 'landscape' : 'portrait'; }));
    defFont(o, 'Zoom', () => (p().fit ? false : p().scale || 100), (v) => set((n) => { n.scale = +v; n.fit = false; }));
    defFont(o, 'FitToPagesWide', () => p().fitW, (v) => set((n) => { n.fitW = +v; n.fit = true; }));
    defFont(o, 'FitToPagesTall', () => p().fitH, (v) => set((n) => { n.fitH = +v; n.fit = true; }));
    defFont(o, 'LeftMargin', () => p().margins.l, (v) => set((n) => { n.margins.l = +v; }));
    defFont(o, 'RightMargin', () => p().margins.r, (v) => set((n) => { n.margins.r = +v; }));
    defFont(o, 'TopMargin', () => p().margins.t, (v) => set((n) => { n.margins.t = +v; }));
    defFont(o, 'BottomMargin', () => p().margins.b, (v) => set((n) => { n.margins.b = +v; }));
    defFont(o, 'PrintArea', () => (p().area ? F.rangeName(p().area) : ''), (v) => set((n) => { n.area = v ? F.parseRange(String(v)) : null; }));
    defFont(o, 'CenterHorizontally', () => tri(p().hCenter), (v) => set((n) => { n.hCenter = on(v); }));
    defFont(o, 'CenterVertically', () => tri(p().vCenter), (v) => set((n) => { n.vCenter = on(v); }));
    o.PrintOut = function () {
      if (!L.hostAudit || !L.hostAudit.has('Print')) gap('G-HOST', 'PrintOut');
      return L.hostAudit.printPdf(() => L.print.pdfBytes({ sheet: sh }));
    };
    return o;
  }
  function pageBreaks(sh, key) {
    const list = () => (sh.print && sh.print[key]) || [];
    const col = collection(() => sheetObj(sh), list, (id) => {
      const o = stamp({}, () => sheetObj(sh));
      Object.defineProperty(o, 'Location', { enumerable: true, get() { return makeRange(sh, [{ r1: key === 'rowBreaks' ? id - 1 : 0, c1: key === 'colBreaks' ? id - 1 : 0, r2: key === 'rowBreaks' ? id - 1 : 0, c2: key === 'colBreaks' ? id - 1 : 0 }]); } });
      o.Delete = function () {
        edit(sh.wb, 'Page Break', () => {
          const next = JSON.parse(JSON.stringify(sh.print));
          next[key] = (next[key] || []).filter((x) => x !== id);
          O.setPrint(sh, next);
        });
      };
      return o;
    }, (id) => String(id));
    col.Add = function (before) {
      const rg = before && before._matheuseRange ? before : null;
      if (!rg) missing();
      const id = key === 'rowBreaks' ? rg._areas[0].r1 + 1 : rg._areas[0].c1 + 1;
      edit(sh.wb, 'Page Break', () => {
        const next = JSON.parse(JSON.stringify(sh.print));
        next[key] = (next[key] || []).concat([id]);
        O.setPrint(sh, next);
      });
    };
    return col;
  }
  function outlineOf(sh) {
    const o = stamp({}, () => sheetObj(sh));
    Object.defineProperty(o, 'SummaryRow', { enumerable: true, get() { return sh.outline && sh.outline.below === false ? 0 : 1; } });
    o.ShowLevels = function (rowLevels) {
      edit(sh.wb, 'Outline', () => {
        const level = rowLevels == null ? 8 : +rowLevels;
        sh.rows.forEach((row) => { if (row && row.level) row.hidden = row.level >= level; });
        if (L.layout && L.layout.invalidate) L.layout.invalidate(sh);
      });
    };
    return o;
  }
  function autoFilterObj(sh) {
    const o = stamp({}, () => sheetObj(sh));
    Object.defineProperty(o, 'Range', { enumerable: true, get() { return sh.autoFilter && sh.autoFilter.ref ? makeRange(sh, [sh.autoFilter.ref]) : null; } });
    o.ShowAllData = function () {
      edit(sh.wb, 'AutoFilter', () => {
        if (sh.autoFilter) O.setAutoFilter(sh, { ref: sh.autoFilter.ref, cols: [] });
        sh.rows.forEach((row) => { if (row && row.filtered) { row.hidden = false; delete row.filtered; } });
      });
    };
    return o;
  }
  function scenariosOf(sh) {
    const list = () => sh.scenarios || [];
    const col = collection(() => sheetObj(sh), list, (s) => scenario(sh, s), (s) => s.name);
    col.Add = function (name, changing, values) {
      const cells = [];
      const rg = changing && changing._matheuseRange ? changing : null;
      if (rg) {
        const flat = Array.isArray(values) ? values.flat ? values.flat() : [].concat.apply([], values) : [values];
        let i = 0;
        eachPos(rg._areas, (r, c) => { cells.push({ r: r, c: c, v: flat[i++] }); });
      }
      const item = { name: String(name), comment: '', cells: cells };
      edit(sh.wb, 'Scenario', () => { sh.scenarios = (sh.scenarios || []).concat([item]); });
      return scenario(sh, item);
    };
    return col;
  }
  function scenario(sh, s) {
    const o = stamp({}, () => sheetObj(sh));
    Object.defineProperty(o, 'Name', { enumerable: true, get() { return s.name; }, set(v) { s.name = String(v); mark(sh.wb); } });
    Object.defineProperty(o, 'Comment', { enumerable: true, get() { return s.comment || ''; }, set(v) { s.comment = String(v); mark(sh.wb); } });
    o.Show = function () {
      edit(sh.wb, 'Scenario', () => { for (const cell of s.cells || []) O.setValue(sh, cell.r, cell.c, cell.v); });
    };
    o.Delete = function () {
      edit(sh.wb, 'Scenario', () => { sh.scenarios = (sh.scenarios || []).filter((x) => x !== s); });
    };
    denyFn(o, ['ChangeScenario'], 'G-ENGINE');
    return o;
  }
  function listObjects(sh) {
    const list = () => (sh.wb.tables || []).filter((t) => t.sheet === sh);
    const col = collection(() => sheetObj(sh), list, (t) => listObject(sh, t), (t) => t.name);
    col.Add = function (sourceType, source, link, xlList, dest) {
      void sourceType; void source; void link; void xlList;
      const rg = dest && dest._matheuseRange ? dest : null;
      if (!rg) missing();
      const ref = Object.assign({}, rg._areas[0]);
      const cols = [];
      for (let c = ref.c1; c <= ref.c2; c++) cols.push({ name: String(readValue(sh, ref.r1, c) || ('Column' + (c - ref.c1 + 1))), id: cols.length + 1 });
      const t = { name: 'Table' + ((sh.wb.tables || []).length + 1), sheet: sh, ref: ref, header: true, totals: false, columns: cols };
      edit(sh.wb, 'List', () => { sh.wb.tables = (sh.wb.tables || []).concat([t]); });
      return listObject(sh, t);
    };
    return col;
  }
  function listObject(sh, t) {
    const o = stamp({}, () => sheetObj(sh));
    Object.defineProperty(o, 'Name', { enumerable: true, get() { return t.name; }, set(v) { t.name = String(v); mark(sh.wb); } });
    Object.defineProperty(o, 'Range', { enumerable: true, get() { return makeRange(sh, [t.ref]); } });
    Object.defineProperty(o, 'ListColumns', {
      enumerable: true,
      get() { return collection(() => o, () => t.columns || [], (c) => { const x = stamp({}, () => o); Object.defineProperty(x, 'Name', { enumerable: true, get() { return c.name; }, set(v) { c.name = String(v); mark(sh.wb); } }); return x; }, (c) => c.name); },
    });
    o.Delete = function () { edit(sh.wb, 'List', () => { sh.wb.tables = (sh.wb.tables || []).filter((x) => x !== t); }); };
    o.Resize = function (range) {
      if (range && range._matheuseRange) { t.ref = Object.assign({}, range._areas[0]); mark(sh.wb); }
    };
    denyFn(o, ['Refresh', 'Publish'], 'G-ENGINE');
    return o;
  }
  const CHART = { 51: 'column', 57: 'bar', 4: 'line', 5: 'pie', [-4169]: 'xy', 1: 'area', [-4120]: 'doughnut' };
  const CHART_BACK = { column: 51, bar: 57, line: 4, pie: 5, xy: -4169, area: 1, doughnut: -4120 };
  function chartObjects(sh) {
    const list = () => (sh.drawings || []).filter((d) => d.kind === 'chart');
    const col = collection(() => sheetObj(sh), list, (d) => chartObject(sh, d), (d) => d.name || '');
    col.Add = function (left, top, width, height) {
      const d = { kind: 'chart', name: 'Chart ' + (list().length + 1), id: drawingId(sh), anchor: { type: 'abs', x: +left || 0, y: +top || 0, w: +width || 360, h: +height || 240 }, chart: { domType: 'column' } };
      edit(sh.wb, 'Chart', () => { O.setDrawings(sh, (sh.drawings || []).concat([d])); });
      return chartObject(sh, d);
    };
    return col;
  }
  function drawingId(sh) { return Math.max(1, ...(sh.drawings || []).map((d) => d.id || 0)) + 1; }
  function anchorBox(d) {
    const a = d.anchor || {};
    if (a.type === 'abs') return { x: a.x || 0, y: a.y || 0, w: a.w || 0, h: a.h || 0 };
    return { x: 0, y: 0, w: 0, h: 0 };
  }
  function setBox(sh, d, box) {
    edit(sh.wb, 'Drawing', () => {
      d.anchor = { type: 'abs', x: box.x, y: box.y, w: box.w, h: box.h };
      O.setDrawings(sh, sh.drawings.slice());
    });
  }
  function chartObject(sh, d) {
    const o = stamp({}, () => sheetObj(sh));
    Object.defineProperty(o, 'Name', { enumerable: true, get() { return d.name || ''; }, set(v) { d.name = String(v); mark(sh.wb); } });
    Object.defineProperty(o, 'Left', { enumerable: true, get() { return anchorBox(d).x; }, set(v) { const b = anchorBox(d); b.x = +v; setBox(sh, d, b); } });
    Object.defineProperty(o, 'Top', { enumerable: true, get() { return anchorBox(d).y; }, set(v) { const b = anchorBox(d); b.y = +v; setBox(sh, d, b); } });
    Object.defineProperty(o, 'Width', { enumerable: true, get() { return anchorBox(d).w; }, set(v) { const b = anchorBox(d); b.w = +v; setBox(sh, d, b); } });
    Object.defineProperty(o, 'Height', { enumerable: true, get() { return anchorBox(d).h; }, set(v) { const b = anchorBox(d); b.h = +v; setBox(sh, d, b); } });
    Object.defineProperty(o, 'Chart', { enumerable: true, get() { return chartOf(sh, d); } });
    o.Delete = function () { edit(sh.wb, 'Chart', () => { O.setDrawings(sh, (sh.drawings || []).filter((x) => x !== d)); }); };
    return o;
  }
  function chartOf(sh, d) {
    const o = stamp({}, () => chartObject(sh, d));
    Object.defineProperty(o, 'ChartType', {
      enumerable: true,
      get() {
        const t = (d.chart && (d.chart.domType || d.chart.type)) || 'column';
        return CHART_BACK[t] != null ? CHART_BACK[t] : XL.xlColumnClustered;
      },
      set(v) {
        const name = CHART[v];
        if (!name) gap('G-ENGINE', 'ChartType');
        d.chart = Object.assign({}, d.chart || {}, { domType: name });
        mark(sh.wb);
      },
    });
    Object.defineProperty(o, 'Name', { enumerable: true, get() { return d.name || ''; } });
    denyGet(o, ['SeriesCollection', 'Axes', 'ChartData', 'PivotLayout'], 'G-ENGINE');
    denyFn(o, ['SetSourceData', 'ApplyLayout'], 'G-ENGINE');
    return o;
  }
  function shapesOf(sh) {
    const list = () => (sh.drawings || []).filter((d) => d.kind === 'shape' || d.kind === 'image');
    const col = collection(() => sheetObj(sh), list, (d) => shapeOf(sh, d), (d) => d.name || '');
    col.AddTextbox = function (orient, left, top, width, height) {
      void orient;
      const d = { kind: 'shape', txBox: true, name: 'TextBox ' + (list().length + 1), id: drawingId(sh), text: '', anchor: { type: 'abs', x: +left || 0, y: +top || 0, w: +width || 120, h: +height || 40 } };
      edit(sh.wb, 'Text Box', () => { O.setDrawings(sh, (sh.drawings || []).concat([d])); });
      return shapeOf(sh, d);
    };
    col.AddPicture = function (fileName, linkToFile) {
      if (linkToFile === true || linkToFile === -1) {
        const e = new Error('The parameter LinkToFile is not available.');
        e.code = 'error';
        throw e;
      }
      if (!L.hostAudit || !L.hostAudit.has('Picture')) gap('G-HOST', 'AddPicture');
      L.hostAudit.rejectExtra(arguments, 2, ['Filename', 'LinkToFile', 'SaveWithDocument', 'Left', 'Top', 'Width', 'Height']);
      return L.hostAudit.readBytes(fileName, 'Picture', 'RequestApproval').then((bytes) => {
        const file = new File([bytes], String(fileName).split(/[/\\]/).pop());
        return L.app.insertPictureFile(file);
      });
    };
    denyFn(col, ['AddChart', 'AddOLEObject'], 'G-HOST');
    return col;
  }
  function shapeOf(sh, d) {
    const o = stamp({}, () => sheetObj(sh));
    Object.defineProperty(o, 'Name', { enumerable: true, get() { return d.name || ''; }, set(v) { d.name = String(v); mark(sh.wb); } });
    Object.defineProperty(o, 'Left', { enumerable: true, get() { return anchorBox(d).x; }, set(v) { const b = anchorBox(d); b.x = +v; setBox(sh, d, b); } });
    Object.defineProperty(o, 'Top', { enumerable: true, get() { return anchorBox(d).y; }, set(v) { const b = anchorBox(d); b.y = +v; setBox(sh, d, b); } });
    Object.defineProperty(o, 'Width', { enumerable: true, get() { return anchorBox(d).w; }, set(v) { const b = anchorBox(d); b.w = +v; setBox(sh, d, b); } });
    Object.defineProperty(o, 'Height', { enumerable: true, get() { return anchorBox(d).h; }, set(v) { const b = anchorBox(d); b.h = +v; setBox(sh, d, b); } });
    Object.defineProperty(o, 'TextFrame', { enumerable: true, get() { return textFrame(sh, d); } });
    o.Delete = function () { edit(sh.wb, 'Shape', () => { O.setDrawings(sh, (sh.drawings || []).filter((x) => x !== d)); }); };
    denyGet(o, ['OLEFormat', 'Chart', 'ActionSettings'], 'G-ENGINE');
    return o;
  }
  function textFrame(sh, d) {
    const o = stamp({}, () => shapeOf(sh, d));
    Object.defineProperty(o, 'TextRange', { enumerable: true, get() { return textRange(sh, d); } });
    return o;
  }
  function textRange(sh, d) {
    const o = stamp({}, () => textFrame(sh, d));
    Object.defineProperty(o, 'Text', {
      enumerable: true,
      get() { return d.text || ''; },
      set(v) { d.text = String(v == null ? '' : v); mark(sh.wb); paint(); },
    });
    return o;
  }
  function pivotTables(sh) {
    const list = () => {
      const bag = sh.wb.extra && sh.wb.extra.pivots && sh.wb.extra.pivots.tables;
      if (!bag) return [];
      return bag.filter((t) => t.sheet === sh.id || t.sheetName === sh.name);
    };
    return collection(() => sheetObj(sh), list, (t) => {
      const o = stamp({}, () => sheetObj(sh));
      Object.defineProperty(o, 'Name', { enumerable: true, get() { return t.name || ''; } });
      denyFn(o, ['Refresh', 'Update'], 'G-ENGINE');
      denyGet(o, ['PivotFields', 'PivotCache', 'DataBodyRange'], 'G-ENGINE');
      return o;
    }, (t) => t.name || '');
  }

  const BUILTIN = { title: 'Title', subject: 'Subject', creator: 'Author', keywords: 'Keywords', description: 'Comments', category: 'Category', manager: 'Manager', company: 'Company' };
  const BUILTIN_KEY = {};
  for (const k in BUILTIN) BUILTIN_KEY[BUILTIN[k].toLowerCase()] = k;
  function propsOf(wb, custom) {
    const list = () => {
      if (!custom) return Object.keys(BUILTIN).map((k) => ({ name: BUILTIN[k], key: k, value: (wb.props && wb.props[k]) || '' }));
      return (wb.props && wb.props.custom) || [];
    };
    const col = collection(() => bookObj(wb), list, (p) => propItem(wb, p, custom), (p) => p.name);
    if (custom) col.Add = function (name, link, type, value) {
      if (link || (type != null && type !== 4)) gap('G-ENGINE', 'Add');
      const item = { name: String(name), value: value == null ? '' : String(value), type: 'lpwstr' };
      if (!wb.props.custom) wb.props.custom = [];
      wb.props.custom.push(item);
      mark(wb);
      return propItem(wb, item, true);
    };
    return col;
  }
  function propItem(wb, p, custom) {
    const o = stamp({}, () => bookObj(wb));
    Object.defineProperty(o, 'Name', { enumerable: true, get() { return p.name; } });
    Object.defineProperty(o, 'Value', {
      enumerable: true,
      get() { return custom ? (p.value || '') : (wb.props[p.key] || ''); },
      set(v) {
        if (custom) p.value = String(v);
        else wb.props[p.key] = String(v);
        mark(wb);
      },
    });
    return o;
  }
  function namesOf(wb, scope) {
    const list = () => wb.names.filter((n) => (scope == null ? n.scope == null : n.scope === scope));
    const col = collection(() => scope == null ? bookObj(wb) : sheetObj(wb.sheets[scope]), list, (n) => nameObj(wb, n), (n) => n.name);
    col.Add = function (name, refersTo) {
      const item = { name: String(name), ref: String(refersTo || '').replace(/^=/, ''), scope: scope == null ? null : scope };
      edit(wb, 'Name', () => { O.setNames(wb, wb.names.concat([item])); });
      return nameObj(wb, item);
    };
    return col;
  }
  function nameObj(wb, n) {
    const o = stamp({}, () => bookObj(wb));
    Object.defineProperty(o, 'Name', { enumerable: true, get() { return n.name; }, set(v) { n.name = String(v); mark(wb); } });
    Object.defineProperty(o, 'RefersTo', {
      enumerable: true,
      get() { return '=' + String(n.ref || '').replace(/^=/, ''); },
      set(v) { n.ref = String(v || '').replace(/^=/, ''); n._ast = undefined; mark(wb); },
    });
    Object.defineProperty(o, 'RefersToR1C1', {
      enumerable: true,
      get() {
        if (!L.xmlss) return o.RefersTo;
        return L.xmlss.toR1C1(String(n.ref || '').replace(/^=/, ''), 0, 0, wb);
      },
    });
    o.Delete = function () { edit(wb, 'Name', () => { O.setNames(wb, wb.names.filter((x) => x !== n)); }); };
    return o;
  }
  function stylesOf(wb) {
    const list = () => wb.cellStyles || [];
    return collection(() => bookObj(wb), list, (s) => {
      const o = stamp({}, () => bookObj(wb));
      Object.defineProperty(o, 'Name', { enumerable: true, get() { return s.name; } });
      Object.defineProperty(o, 'Font', { enumerable: true, get() { return fontOf(wb.sheets[wb.active], [{ r1: 0, c1: 0, r2: 0, c2: 0 }]); } });
      return o;
    }, (s) => s.name);
  }

  const books_ = new Map();
  function bookObj(wb) {
    let o = books_.get(wb);
    if (o) return o;
    o = stamp({}, () => workbooks());
    o._wb = wb;
    books_.set(wb, o);
    Object.defineProperty(o, 'Name', { enumerable: true, get() { const b = bookOf(wb); return (b && b.name) || wb.fileName || 'Book1'; } });
    Object.defineProperty(o, 'Saved', {
      enumerable: true,
      get() { const b = bookOf(wb); return b ? !b.dirty : true; },
      set(v) { const b = bookOf(wb); if (b && on(v)) b.dirty = false; },
    });
    Object.defineProperty(o, 'Worksheets', { enumerable: true, get() { return sheetsCol(wb, 'worksheet'); } });
    Object.defineProperty(o, 'Sheets', { enumerable: true, get() { return sheetsCol(wb, 'all'); } });
    Object.defineProperty(o, 'Charts', { enumerable: true, get() { return sheetsCol(wb, 'chart'); } });
    Object.defineProperty(o, 'ActiveSheet', { enumerable: true, get() { return sheetObj(wb.sheets[wb.active] || wb.sheets[0]); } });
    Object.defineProperty(o, 'Names', { enumerable: true, get() { return namesOf(wb, null); } });
    Object.defineProperty(o, 'Styles', { enumerable: true, get() { return stylesOf(wb); } });
    Object.defineProperty(o, 'BuiltInDocumentProperties', { enumerable: true, get() { return propsOf(wb, false); } });
    Object.defineProperty(o, 'CustomDocumentProperties', { enumerable: true, get() { return propsOf(wb, true); } });
    o.Activate = function () {
      const i = books().findIndex((b) => b.wb === wb);
      if (i >= 0) setCur(i);
    };
    o.Save = function () {
      const b = bookOf(wb);
      if (!b || b.untitled || !b.saved) gap('G-HOST', 'Save');
      const run = () => {
        if (L.app && L.app.save) return L.app.save();
        b.dirty = false;
      };
      if (!L.hostAudit) return run();
      return L.hostAudit.must('RequestApproval', 'Save', 'save the workbook').then(run);
    };
    function bookBytes(fileFormat) {
      if (fileFormat == null || fileFormat === '' || fileFormat === 51 || fileFormat === 'xlsx') return L.xlsxWrite.write(wb, { type: 'xlsx' });
      if (fileFormat === 52 || fileFormat === 'xlsm') return L.xlsxWrite.write(wb, { type: 'xlsm' });
      if (fileFormat === 6 || fileFormat === 'csv') return L.csv.write(wb, wb.sheets[wb.active] || wb.sheets[0], ',');
      if (fileFormat === -4158 || fileFormat === 'txt') return L.csv.write(wb, wb.sheets[wb.active] || wb.sheets[0], '\t');
      const e = new Error('The parameter FileFormat is not available.');
      e.code = 'error';
      throw e;
    }
    o.SaveAs = function (fileName, fileFormat) {
      if (!L.hostAudit || !L.hostAudit.has('SaveAs')) gap('G-HOST', 'SaveAs');
      L.hostAudit.rejectExtra(arguments, 2, ['Filename', 'FileFormat', 'Password', 'WriteResPassword']);
      return L.hostAudit.saveCopy(fileName, () => bookBytes(fileFormat));
    };
    o.SaveAs2 = o.SaveAs;
    o.PrintOut = function () {
      if (!L.hostAudit || !L.hostAudit.has('Print')) gap('G-HOST', 'PrintOut');
      return L.hostAudit.printPdf(() => L.print.pdfBytes({ what: 'workbook' }));
    };
    o.ExportAsFixedFormat = function (format, fileName) {
      if (!L.hostAudit || !L.hostAudit.has('FixedFormat')) gap('G-HOST', 'ExportAsFixedFormat');
      return L.hostAudit.exportPdf(format, fileName, () => L.print.pdfBytes({ what: 'workbook' }));
    };
    o.FollowHyperlink = function (address) {
      if (!L.hostAudit || !L.hostAudit.has('Hyperlink')) gap('G-HOST', 'FollowHyperlink');
      L.hostAudit.rejectExtra(arguments, 1, ['Address', 'SubAddress', 'NewWindow', 'AddHistory', 'ExtraInfo', 'Method', 'HeaderInfo']);
      return L.hostAudit.follow(address);
    };
    o.Close = function (saveChanges) {
      const b = bookOf(wb);
      if (b && b.dirty && saveChanges == null && flags.alerts) gap('G-HOST', 'Close');
      if (b && b.dirty && on(saveChanges)) gap('G-HOST', 'Close');
      const list = books();
      const i = list.indexOf(b);
      if (L.app && L.app.books && b) {
        if (L.app.closeWorkbook && L.app.books[L.app.cur] === b) L.app.closeWorkbook();
        else {
          L.app.books.splice(i, 1);
          if (!L.app.books.length) L.app.newWorkbook && L.app.newWorkbook();
          else L.app.cur = Math.min(L.app.cur, L.app.books.length - 1);
        }
      } else if (i >= 0) {
        hold.books.splice(i, 1);
        if (!hold.books.length) hold.books.push(blankRecord('Book1'));
        hold.cur = Math.min(hold.cur, hold.books.length - 1);
      }
      books_.delete(wb);
      paint();
    };
    o.Protect = function (password) {
      edit(wb, 'Protect', () => { wb.protection = { workbook: true, password: password ? hashPw(password) : undefined }; });
    };
    o.Unprotect = function (password) {
      if (!wb.protection) return;
      if (wb.protection.password && hashPw(password || '') !== wb.protection.password) {
        const e = new Error('The password is not correct.');
        e.code = 'missing';
        throw e;
      }
      edit(wb, 'Protect', () => { wb.protection = null; });
    };
    Object.defineProperty(o, 'FullName', {
      enumerable: true,
      get() {
        if (!L.hostAudit || !L.hostAudit.has('Path')) gap('G-HOST', 'FullName');
        return L.hostAudit.documentPath();
      },
    });
    Object.defineProperty(o, 'Path', {
      enumerable: true,
      get() {
        if (!L.hostAudit || !L.hostAudit.has('Path')) gap('G-HOST', 'Path');
        return L.hostAudit.documentDir();
      },
    });
    denyGet(o, ['Windows'], 'G-HOST');
    denyFn(o, ['SendMail'], 'G-HOST');
    denyFn(o, ['Run', 'Merge'], 'G-ENGINE');
    return o;
  }
  function workbooks() {
    const col = collection(() => Application, books, (b) => bookObj(b.wb), (b) => b.name);
    col.Add = function () {
      const rec = blankRecord(nextBookName());
      if (L.app && L.app.addBook) L.app.addBook(rec.wb, rec.name, { type: 'xlsx', untitled: true });
      else { hold.books.push(rec); hold.cur = hold.books.length - 1; }
      paint();
      return bookObj(rec.wb);
    };
    col.Open = function (fileName) {
      if (!L.hostAudit || !L.hostAudit.has('Open')) gap('G-HOST', 'Open');
      L.hostAudit.rejectExtra(arguments, 1, ['Filename', 'UpdateLinks', 'ReadOnly', 'Format', 'Password']);
      return L.hostAudit.openPath(fileName);
    };
    col.Close = function () { gap('G-HOST', 'Close'); };
    return col;
  }
  function nextBookName() {
    const used = new Set(books().map((b) => String(b.name).toLowerCase()));
    let i = books().length + 1;
    while (used.has('book' + i)) i++;
    return 'Book' + i;
  }
  function activeSheet() {
    const wb = curBook().wb;
    return wb.sheets[wb.active] || wb.sheets[0];
  }
  function argFormula(a) {
    if (a && a._matheuseRange) {
      const sh = a._sheet;
      const name = F.quoteSheet(sh.name);
      return name + '!' + F.rangeName(a._areas[0]);
    }
    if (Array.isArray(a)) {
      const rows = a.map((row) => (Array.isArray(row) ? row : [row]).map((x) => argFormula(x)).join(','));
      return '{' + rows.join(';') + '}';
    }
    if (a == null) return '0';
    if (typeof a === 'string') return '"' + a.replace(/"/g, '""') + '"';
    if (a === true || a === -1) return 'TRUE';
    if (a === false) return 'FALSE';
    return String(a);
  }
  function callFn(name, args) {
    const info = L.fninfo && L.fninfo.get(String(name).replace(/_/g, '.'));
    const fn = info ? info.name : String(name);
    if (!L.fninfo || !L.fninfo.get(fn)) {
      const e = new Error(name + ' is not available');
      e.code = 'absent';
      throw e;
    }
    const sh = activeSheet();
    const text = fn + '(' + Array.prototype.map.call(args, argFormula).join(',') + ')';
    let v;
    try { v = L.calc.evalScalar(sh.wb, sh, 0, 0, text); } catch (e) { gap('G-ENGINE', fn); }
    if (v && typeof v === 'object' && !M.isErr(v) && Array.isArray(v.rows)) {
      return v.rows;
    }
    return cellOut(v);
  }
  function worksheetFunction() {
    const o = stamp({}, Application);
    o._wf = true;
    o.invoke = function (name, args) { return callFn(name, args || []); };
    const fns = Object.create(null);
    if (L.fninfo && L.fninfo.all) {
      for (const f of L.fninfo.all()) {
        const call = function () { return callFn(f.name, arguments); };
        fns[f.name.toUpperCase()] = call;
        fns[f.name.replace(/\./g, '_').toUpperCase()] = call;
      }
    }
    return new Proxy(o, {
      get(t, p) {
        if (typeof p !== 'string') return t[p];
        if (Object.prototype.hasOwnProperty.call(t, p) || p === 'Application' || p === 'Parent' || p === 'Creator') return t[p];
        return fns[p.toUpperCase()] || t[p];
      },
    });
  }

  stamp(Application, Application);
  Object.defineProperty(Application, 'ActiveWorkbook', { enumerable: true, get() { return bookObj(curBook().wb); } });
  Object.defineProperty(Application, 'ThisWorkbook', { enumerable: true, get() { return Application.ActiveWorkbook; } });
  Object.defineProperty(Application, 'ActiveSheet', { enumerable: true, get() { return sheetObj(activeSheet()); } });
  Object.defineProperty(Application, 'ActiveCell', { enumerable: true, get() { return selectionOf(activeSheet()).Item(1); } });
  Object.defineProperty(Application, 'Selection', { enumerable: true, get() { return selectionOf(activeSheet()); } });
  Object.defineProperty(Application, 'Workbooks', { enumerable: true, get() { return workbooks(); } });
  Object.defineProperty(Application, 'Worksheets', { enumerable: true, get() { return sheetsCol(curBook().wb, 'worksheet'); } });
  Object.defineProperty(Application, 'Sheets', { enumerable: true, get() { return sheetsCol(curBook().wb, 'all'); } });
  Object.defineProperty(Application, 'Names', { enumerable: true, get() { return namesOf(curBook().wb, null); } });
  Object.defineProperty(Application, 'WorksheetFunction', { enumerable: true, get() { return worksheetFunction(); } });
  Object.defineProperty(Application, 'ScreenUpdating', { enumerable: true, get() { return tri(flags.screen); }, set(v) { flags.screen = on(v) || v == null; if (flags.screen) paint(); } });
  Object.defineProperty(Application, 'DisplayAlerts', { enumerable: true, get() { return tri(flags.alerts); }, set(v) { flags.alerts = on(v); } });
  Object.defineProperty(Application, 'Calculation', {
    enumerable: true,
    get() { return curBook().wb.calcPr.mode === 'manual' ? XL.xlCalculationManual : XL.xlCalculationAutomatic; },
    set(v) {
      if (v === XL.xlCalculationSemiautomatic) gap('G-ENGINE', 'Calculation');
      curBook().wb.calcPr.mode = v === XL.xlCalculationManual ? 'manual' : 'auto';
    },
  });
  Object.defineProperty(Application, 'ReferenceStyle', {
    enumerable: true,
    get() { return curBook().wb.r1c1 ? XL.xlR1C1 : XL.xlA1; },
    set(v) { curBook().wb.r1c1 = v === XL.xlR1C1; paint(); },
  });
  Application.Calculate = function () {
    const wb = curBook().wb;
    if (L.calc && L.calc.recalcAll) L.calc.recalcAll(wb);
    paint();
  };
  Application.Undo = function () {
    const t = curBook().wb.undo.undo();
    paint();
    return t ? -1 : 0;
  };
  Application.Redo = function () {
    const t = curBook().wb.undo.redo();
    paint();
    return t ? -1 : 0;
  };
  denyGet(Application, ['Windows', 'AddIns', 'CommandBars'], 'G-HOST');
  denyFn(Application, ['Run', 'SendKeys', 'Wait', 'OnTime'], 'G-HOST');
  Object.defineProperty(Application, 'Dialogs', {
    enumerable: true,
    get() { return L.hostAudit ? L.hostAudit.dialogsCollection() : gap('G-HOST', 'Dialogs'); },
  });
  Application.FileDialog = function (type) {
    if (!L.hostAudit) gap('G-HOST', 'FileDialog');
    return L.hostAudit.fileDialog(type);
  };
  Application.HasHostAbility = function (name) { return L.hostAudit && L.hostAudit.has(name) ? -1 : 0; };
  Application.MsgBox = function (prompt, buttons, title) {
    if (!L.hostAudit) gap('G-HOST', 'MsgBox');
    return L.hostAudit.msgBox(prompt, buttons, title);
  };
  Application.InputBox = function (prompt, title, def) {
    if (!L.hostAudit) gap('G-HOST', 'InputBox');
    return L.hostAudit.inputBox(prompt, title, def);
  };
  Application.HostDir = function (pattern) {
    if (!L.hostAudit || !L.hostAudit.has('Dir')) gap('G-HOST', 'Dir');
    return L.hostAudit.dir(pattern);
  };
  Application.HostShell = function () {
    if (!L.hostAudit || !L.hostAudit.has('Shell')) gap('G-HOST', 'Shell');
    return L.hostAudit.shell();
  };
  Application.GetOpenFilename = function (filter, filterIndex, title, buttonText, multi) {
    void filter; void filterIndex; void title; void buttonText;
    if (!L.hostAudit) gap('G-HOST', 'GetOpenFilename');
    return L.hostAudit.getOpenFilename(multi);
  };
  Application.GetSaveAsFilename = function () {
    if (!L.hostAudit) gap('G-HOST', 'GetSaveAsFilename');
    return L.hostAudit.getSaveAsFilename();
  };
  denyGet(Application, ['EnableEvents'], 'G-ENGINE');

  dom.Application = Application;
  dom.reset = function () {
    hold.books = [];
    hold.cur = 0;
    flags.screen = true;
    flags.alerts = true;
    clip = null;
    sheets.clear();
    books_.clear();
  };
  dom.flags = flags;
})();
