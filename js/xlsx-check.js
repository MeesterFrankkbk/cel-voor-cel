/* Cel voor Cel – leest een .xlsx-bestand en geeft waarden én opmaak terug.
   Werkt in de browser (window.XlsxCheck) en in Node (module.exports).
   Afhankelijk van JSZip en een DOMParser (worden meegegeven). */
(function (root) {
  'use strict';

  const F = (typeof module !== 'undefined' && module.exports) ? require('./formula.js') : root.CvcFormula;

  // ---------- adressen ----------
  function colToNum(col) {
    let n = 0;
    for (const ch of col.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n;
  }
  function numToCol(n) {
    let s = '';
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }
  function parseAddr(a) {
    const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(String(a).trim());
    if (!m) return null;
    return { c: colToNum(m[1]), r: parseInt(m[2], 10) };
  }
  function addr(c, r) { return numToCol(c) + r; }
  function parseRange(ref) {
    const [a, b] = String(ref).split(':');
    const p = parseAddr(a), q = parseAddr(b || a);
    return { c1: Math.min(p.c, q.c), r1: Math.min(p.r, q.r), c2: Math.max(p.c, q.c), r2: Math.max(p.r, q.r) };
  }
  function eachCell(ref, fn) {
    const g = parseRange(ref);
    for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) fn(addr(c, r), c, r);
  }

  // ---------- XML-hulp (los van namespaces/prefixen) ----------
  function kids(el, name) {
    const out = [];
    if (!el) return out;
    for (let n = el.firstChild; n; n = n.nextSibling) {
      if (n.nodeType === 1 && (n.localName || n.nodeName.split(':').pop()) === name) out.push(n);
    }
    return out;
  }
  function kid(el, name) { return kids(el, name)[0] || null; }
  function deep(el, name) {
    const out = [];
    (function walk(e) {
      for (let n = e.firstChild; n; n = n.nextSibling) {
        if (n.nodeType !== 1) continue;
        if ((n.localName || n.nodeName.split(':').pop()) === name) out.push(n);
        walk(n);
      }
    })(el);
    return out;
  }
  function attr(el, name) {
    if (!el) return null;
    if (el.hasAttribute && el.hasAttribute(name)) return el.getAttribute(name);
    // attributen met prefix (r:id)
    const at = el.attributes;
    for (let i = 0; at && i < at.length; i++) {
      const a = at[i];
      if ((a.localName || a.name.split(':').pop()) === name) return a.value;
    }
    return null;
  }
  function textOf(el) { return el ? (el.textContent || '') : ''; }

  // ---------- kleuren ----------
  const INDEXED = ['000000','FFFFFF','FF0000','00FF00','0000FF','FFFF00','FF00FF','00FFFF','000000','FFFFFF','FF0000','00FF00','0000FF','FFFF00','FF00FF','00FFFF','800000','008000','000080','808000','800080','008080','C0C0C0','808080','9999FF','993366','FFFFCC','CCFFFF','660066','FF8080','0066CC','CCCCFF','000080','FF00FF','FFFF00','00FFFF','800080','800000','008080','0000FF','00CCFF','CCFFFF','CCFFCC','FFFF99','99CCFF','FF99CC','CC99FF','FFCC99','3366FF','33CCCC','99CC00','FFCC00','FF9900','FF6600','666699','969696','003366','339966','003300','333300','993300','993366','333399','333333'];

  function rgbToHsl(hex) {
    const r = parseInt(hex.slice(0, 2), 16) / 255, g = parseInt(hex.slice(2, 4), 16) / 255, b = parseInt(hex.slice(4, 6), 16) / 255;
    const max = Math.max(r, g, b), min = Math.min(r, g, b);
    let h = 0, s = 0; const l = (max + min) / 2;
    if (max !== min) {
      const d = max - min;
      s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
      if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
      else if (max === g) h = (b - r) / d + 2;
      else h = (r - g) / d + 4;
      h *= 60;
    }
    return { h, s, l };
  }
  function hslToRgb(h, s, l) {
    function f(n) { const k = (n + h / 30) % 12; const a = s * Math.min(l, 1 - l); return l - a * Math.max(-1, Math.min(k - 3, 9 - k, 1)); }
    const to = v => Math.round(v * 255).toString(16).padStart(2, '0').toUpperCase();
    return to(f(0)) + to(f(8)) + to(f(4));
  }
  function applyTint(hex, tint) {
    if (!tint) return hex;
    const { h, s, l } = rgbToHsl(hex);
    const nl = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
    return hslToRgb(h, s, Math.max(0, Math.min(1, nl)));
  }
  function resolveColor(el, theme) {
    if (!el) return null;
    let hex = null;
    const rgb = attr(el, 'rgb'), th = attr(el, 'theme'), ix = attr(el, 'indexed'), auto = attr(el, 'auto');
    if (rgb) hex = rgb.length === 8 ? rgb.slice(2) : rgb;
    else if (th !== null && theme[+th]) hex = theme[+th];
    else if (ix !== null) { const i = +ix; if (i === 64 || i === 65) return null; hex = INDEXED[i] || null; }
    else if (auto) return null;
    if (!hex) return null;
    const tint = parseFloat(attr(el, 'tint') || '0');
    return applyTint(hex.toUpperCase(), tint);
  }
  /* Geeft de kleurnamen die bij een kleur passen (ruim: een kleur kan bij 2 namen horen). */
  function colorNames(hex) {
    if (!hex) return [];
    const { h, s, l } = rgbToHsl(hex);
    if (l >= 0.96) return ['wit'];
    if (l <= 0.12) return ['zwart'];
    if (s < 0.18) return ['grijs'];
    const names = [];
    const inR = (a, b) => (a <= b ? h >= a && h < b : h >= a || h < b);
    if (inR(340, 18)) names.push('rood');
    if (inR(300, 350) && l > 0.55) names.push('roze');
    if (inR(12, 50)) names.push('oranje');
    if (inR(38, 75)) names.push('geel');
    if (inR(70, 170)) names.push('groen');
    if (inR(165, 260)) names.push('blauw');
    if (inR(255, 320)) names.push('paars');
    if (inR(15, 45) && l < 0.35) names.push('bruin');
    return names;
  }

  // ---------- werkmap lezen ----------
  async function readWorkbook(data, deps) {
    const JSZip = deps.JSZip, DP = deps.DOMParser;
    const zip = await JSZip.loadAsync(data);
    const parse = async path => {
      const f = zip.file(path);
      if (!f) return null;
      const txt = await f.async('string');
      return new DP().parseFromString(txt, 'application/xml').documentElement;
    };

    const wbXml = await parse('xl/workbook.xml');
    if (!wbXml) throw new Error('Dit is geen Excel-bestand (.xlsx).');
    const relsXml = await parse('xl/_rels/workbook.xml.rels');
    const rels = {};
    kids(relsXml, 'Relationship').forEach(r => { rels[attr(r, 'Id')] = attr(r, 'Target'); });

    // thema
    const themeTarget = Object.values(rels).find(t => /theme/i.test(t));
    const theme = [];
    if (themeTarget) {
      const themeXml = await parse('xl/' + themeTarget.replace(/^\/?xl\//, '').replace(/^\//, ''));
      const scheme = themeXml && deep(themeXml, 'clrScheme')[0];
      if (scheme) {
        const get = n => {
          const e = kid(scheme, n); if (!e) return null;
          const c = e.firstElementChild || kids(e, 'srgbClr')[0] || kids(e, 'sysClr')[0];
          if (!c) return null;
          return (attr(c, 'val') && (c.localName || c.nodeName.split(':').pop()) === 'srgbClr') ? attr(c, 'val') : (attr(c, 'lastClr') || attr(c, 'val'));
        };
        // Excel wisselt 0<->1 en 2<->3
        ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'].forEach(n => theme.push((get(n) || '').toUpperCase() || null));
        if (theme[0] && !/^[0-9A-F]{6}$/.test(theme[0])) theme[0] = 'FFFFFF';
        if (theme[1] && !/^[0-9A-F]{6}$/.test(theme[1])) theme[1] = '000000';
      }
    }
    if (!theme.length) theme.push('FFFFFF', '000000', 'E8E8E8', '0E2841', '156082', 'E97132', '196B24', '0F9ED5', 'A02B93', '4EA72E');

    // gedeelde teksten
    const sst = [];
    const sstXml = await parse('xl/sharedStrings.xml');
    if (sstXml) kids(sstXml, 'si').forEach(si => {
      let t = '';
      (function walk(e) {
        for (let n = e.firstChild; n; n = n.nextSibling) {
          if (n.nodeType !== 1) continue;
          const ln = n.localName || n.nodeName.split(':').pop();
          if (ln === 'rPh' || ln === 'phoneticPr') continue;
          if (ln === 't') t += textOf(n); else walk(n);
        }
      })(si);
      sst.push(cleanText(t));
    });

    // stijlen
    const styles = await readStyles(await parse('xl/styles.xml'), theme);

    // bladen
    const sheets = [];
    for (const s of deep(wbXml, 'sheet')) {
      const target = rels[attr(s, 'id')];
      if (!target) continue;
      const path = target.startsWith('/') ? target.slice(1) : 'xl/' + target;
      const sx = await parse(path);
      if (!sx) continue;
      sheets.push(readSheet(sx, attr(s, 'name'), attr(s, 'state') || 'visible', sst, styles));
    }

    // laatst bewaard
    let modified = null, lastModifiedBy = null, application = null;
    const core = await parse('docProps/core.xml');
    if (core) {
      const m = deep(core, 'modified')[0]; if (m) modified = new Date(textOf(m));
      const lm = deep(core, 'lastModifiedBy')[0]; if (lm) lastModifiedBy = textOf(lm);
    }
    const app = await parse('docProps/app.xml');
    if (app) { const a = deep(app, 'Application')[0]; if (a) application = textOf(a); }

    return new Workbook(sheets, styles, { modified, lastModifiedBy, application });
  }

  function cleanText(t) {
    return t.replace(/_x000D_/g, '').replace(/_x000A_/g, '\n').replace(/\r\n?/g, '\n');
  }

  async function readStyles(xml, theme) {
    const res = { fonts: [], fills: [], borders: [], xfs: [], numFmts: {}, dxfs: [] };
    if (!xml) { res.fonts.push({}); res.fills.push(null); res.borders.push({}); res.xfs.push({ fontId: 0, fillId: 0, borderId: 0, align: {} }); return res; }
    const nf = kid(xml, 'numFmts');
    kids(nf, 'numFmt').forEach(n => { res.numFmts[attr(n, 'numFmtId')] = attr(n, 'formatCode'); });
    const readFont = f => {
      const has = n => { const e = kid(f, n); return !!e && attr(e, 'val') !== '0' && attr(e, 'val') !== 'false' && attr(e, 'val') !== 'none'; };
      const sz = kid(f, 'sz'), nm = kid(f, 'name'), u = kid(f, 'u');
      return {
        bold: has('b'), italic: has('i'), strike: has('strike'),
        underline: u ? (attr(u, 'val') || 'single') !== 'none' : false,
        size: sz ? parseFloat(attr(sz, 'val')) : null,
        name: nm ? attr(nm, 'val') : null,
        color: resolveColor(kid(f, 'color'), theme)
      };
    };
    const readFill = f => {
      const p = kid(f, 'patternFill');
      if (p) {
        const type = attr(p, 'patternType');
        if (!type || type === 'none' || type === 'gray125') {
          // in dxf-stijlen staat soms geen patternType maar wel een bgColor
          const bg = resolveColor(kid(p, 'bgColor'), theme) || resolveColor(kid(p, 'fgColor'), theme);
          return (!type && bg) ? bg : null;
        }
        return resolveColor(kid(p, 'fgColor'), theme) || resolveColor(kid(p, 'bgColor'), theme);
      }
      const g = kid(f, 'gradientFill');
      if (g) { const st = deep(g, 'color')[0]; return resolveColor(st, theme); }
      return null;
    };
    const readBorder = b => {
      const side = n => { const e = kid(b, n); const st = e && attr(e, 'style'); return st && st !== 'none' ? st : null; };
      return { left: side('left') || side('start'), right: side('right') || side('end'), top: side('top'), bottom: side('bottom') };
    };
    kids(kid(xml, 'fonts'), 'font').forEach(f => res.fonts.push(readFont(f)));
    kids(kid(xml, 'fills'), 'fill').forEach(f => res.fills.push(readFill(f)));
    kids(kid(xml, 'borders'), 'border').forEach(b => res.borders.push(readBorder(b)));
    kids(kid(xml, 'cellXfs'), 'xf').forEach(x => {
      const al = kid(x, 'alignment');
      res.xfs.push({
        fontId: +(attr(x, 'fontId') || 0), fillId: +(attr(x, 'fillId') || 0), borderId: +(attr(x, 'borderId') || 0),
        numFmtId: +(attr(x, 'numFmtId') || 0),
        align: al ? { h: attr(al, 'horizontal'), v: attr(al, 'vertical'), wrap: attr(al, 'wrapText') === '1' || attr(al, 'wrapText') === 'true', rotate: +(attr(al, 'textRotation') || 0) } : {},
        locked: (() => { const p = kid(x, 'protection'); return p ? attr(p, 'locked') !== '0' : true; })()
      });
    });
    kids(kid(xml, 'dxfs'), 'dxf').forEach(d => {
      const f = kid(d, 'font'), fi = kid(d, 'fill');
      res.dxfs.push({ font: f ? readFont(f) : null, fill: fi ? readFill(fi) : null });
    });
    res.readDxf = d => { const f = kid(d, 'font'), fi = kid(d, 'fill'); return { font: f ? readFont(f) : null, fill: fi ? readFill(fi) : null }; };
    if (!res.fonts.length) res.fonts.push({});
    return res;
  }

  function readSheet(sx, name, state, sst, styles) {
    const cells = {};
    const rowsInfo = {};
    const sd = kid(sx, 'sheetData');
    let rowNo = 0;
    kids(sd, 'row').forEach(row => {
      rowNo = attr(row, 'r') ? +attr(row, 'r') : rowNo + 1;
      if (attr(row, 'ht')) rowsInfo[rowNo] = { ht: parseFloat(attr(row, 'ht')), custom: attr(row, 'customHeight') === '1', hidden: attr(row, 'hidden') === '1' };
      else if (attr(row, 'hidden') === '1') rowsInfo[rowNo] = { hidden: true };
      let colNo = 0;
      kids(row, 'c').forEach(c => {
        const p = attr(c, 'r') ? parseAddr(attr(c, 'r')) : { c: colNo + 1, r: rowNo };
        colNo = p.c;
        const t = attr(c, 't') || 'n';
        const vEl = kid(c, 'v'), fEl = kid(c, 'f');
        let v = null;
        if (t === 's') v = vEl ? sst[+textOf(vEl)] : null;
        else if (t === 'inlineStr') { const is = kid(c, 'is'); v = is ? cleanText(deep(is, 't').map(textOf).join('')) : ''; }
        else if (t === 'str') v = vEl ? cleanText(textOf(vEl)) : '';
        else if (t === 'b') v = vEl ? textOf(vEl) === '1' : null;
        else if (t === 'e') v = vEl ? textOf(vEl) : null;
        else v = vEl && textOf(vEl) !== '' ? parseFloat(textOf(vEl)) : null;
        cells[addr(p.c, p.r)] = { v, t, f: fEl ? textOf(fEl) : null, ft: fEl ? attr(fEl, 't') : null, si: fEl ? attr(fEl, 'si') : null, s: +(attr(c, 's') || 0) };
      });
    });
    // gedeelde formules: Excel bewaart de formule één keer en verwijst ernaar vanuit de andere cellen
    const masters = {};
    Object.entries(cells).forEach(([a, c]) => { if (c.ft === 'shared' && c.f && c.si != null) masters[c.si] = { a, f: c.f }; });
    Object.entries(cells).forEach(([a, c]) => {
      if (c.ft === 'shared' && !c.f && masters[c.si] && F) {
        const m = parseAddr(masters[c.si].a), q = parseAddr(a);
        c.f = F.shift(masters[c.si].f, q.r - m.r, q.c - m.c);
      }
    });
    const cols = [];
    kids(kid(sx, 'cols'), 'col').forEach(c => cols.push({
      min: +attr(c, 'min'), max: +attr(c, 'max'), width: attr(c, 'width') ? parseFloat(attr(c, 'width')) : null,
      custom: attr(c, 'customWidth') === '1', bestFit: attr(c, 'bestFit') === '1', hidden: attr(c, 'hidden') === '1', s: +(attr(c, 'style') || 0)
    }));
    const fp = kid(sx, 'sheetFormatPr');
    const defaultColWidth = fp && attr(fp, 'defaultColWidth') ? parseFloat(attr(fp, 'defaultColWidth')) : (fp && attr(fp, 'baseColWidth') ? parseFloat(attr(fp, 'baseColWidth')) + 0.71 : 8.43);
    const defaultRowHeight = fp && attr(fp, 'defaultRowHeight') ? parseFloat(attr(fp, 'defaultRowHeight')) : 15;
    const merges = kids(kid(sx, 'mergeCells'), 'mergeCell').map(m => attr(m, 'ref'));
    const prot = kid(sx, 'sheetProtection');
    const isProtected = !!prot && attr(prot, 'sheet') !== '0' && attr(prot, 'sheet') !== 'false' && attr(prot, 'sheet') !== null;
    const dataValidations = deep(sx, 'dataValidation').map(d => ({
      sqref: (attr(d, 'sqref') || textOf(kid(d, 'sqref'))).trim(), type: attr(d, 'type') || 'any', operator: attr(d, 'operator') || 'between',
      formula1: textOf(kid(d, 'formula1')).trim(), formula2: textOf(kid(d, 'formula2')).trim(), errorTitle: attr(d, 'errorTitle'), errorStyle: attr(d, 'errorStyle') || 'stop',
      showErrorMessage: ['1', 'true'].includes(attr(d, 'showErrorMessage')), error: attr(d, 'error'), prompt: attr(d, 'prompt')
    }));
    const conditionalFormats = [];
    deep(sx, 'conditionalFormatting').forEach(cf => {
      const sq = attr(cf, 'sqref') || textOf(kid(cf, 'sqref'));
      kids(cf, 'cfRule').forEach(r => {
        const fs = kids(r, 'formula').map(textOf);
        conditionalFormats.push({
          sqref: (sq || '').trim(), type: attr(r, 'type'), operator: attr(r, 'operator'), text: attr(r, 'text'), priority: +(attr(r, 'priority') || 99),
          formulas: fs.length ? fs : kids(r, 'f').map(textOf),
          dxf: attr(r, 'dxfId') !== null ? styles.dxfs[+attr(r, 'dxfId')] : (kid(r, 'dxf') && styles.readDxf ? styles.readDxf(kid(r, 'dxf')) : null)
        });
      });
    });
    const sv = deep(sx, 'sheetView')[0];
    const showGrid = sv ? attr(sv, 'showGridLines') !== '0' && attr(sv, 'showGridLines') !== 'false' : true;
    const drawings = kids(sx, 'drawing').length;
    return new Sheet({ name, state, cells, rowsInfo, cols, defaultColWidth, defaultRowHeight, merges, isProtected, dataValidations, conditionalFormats, showGrid, drawings }, styles);
  }

  // ---------- Werkmap / blad ----------
  function Workbook(sheets, styles, meta) { this.sheets = sheets; this.styles = styles; this.meta = meta; sheets.forEach(s => { s.wb = this; }); }
  Workbook.prototype.sheet = function (name) {
    if (name == null) return this.sheets.find(s => s.state === 'visible') || this.sheets[0];
    const n = String(name).trim().toLowerCase();
    return this.sheets.find(s => s.name.trim().toLowerCase() === n) || null;
  };
  Workbook.prototype.sheetNames = function () { return this.sheets.map(s => s.name); };

  function Sheet(o, styles) { Object.assign(this, o); this.styles = styles; }
  Sheet.prototype.cell = function (a) { return this.cells[String(a).toUpperCase()] || null; };
  Sheet.prototype.value = function (a) { const c = this.cell(a); return c ? c.v : null; };
  Sheet.prototype.text = function (a) { const v = this.value(a); return v == null ? '' : String(v); };
  Sheet.prototype.isEmpty = function (a) { const v = this.value(a); return v == null || String(v).trim() === ''; };
  Sheet.prototype.filled = function () {
    return Object.keys(this.cells).filter(a => !this.isEmpty(a));
  };
  Sheet.prototype.find = function (pred) { return this.filled().filter(a => pred(this.value(a), a)); };
  Sheet.prototype.findText = function (txt) {
    const t = norm(txt);
    return this.filled().filter(a => norm(this.text(a)) === t);
  };
  Sheet.prototype.style = function (a) {
    const c = this.cell(a);
    let s = c ? c.s : 0;
    if (!c) { // kolomstijl
      const p = parseAddr(a); const col = this.cols.find(k => p.c >= k.min && p.c <= k.max); if (col) s = col.s;
    }
    const xf = this.styles.xfs[s] || this.styles.xfs[0] || { align: {} };
    const font = this.styles.fonts[xf.fontId] || {};
    const fill = this.styles.fills[xf.fillId] || null;
    const border = this.styles.borders[xf.borderId] || {};
    return {
      bold: !!font.bold, italic: !!font.italic, underline: !!font.underline, strike: !!font.strike,
      size: font.size, fontName: font.name, fontColor: font.color || null, fill,
      h: xf.align.h || null, v: xf.align.v || null, wrap: !!xf.align.wrap, rotate: xf.align.rotate || 0,
      border, numFmtId: xf.numFmtId, numFmt: this.styles.numFmts[xf.numFmtId] || BUILTIN_FMT[xf.numFmtId] || 'General',
      locked: xf.locked !== false
    };
  };
  /* De formule in een cel (zonder =), of null. */
  Sheet.prototype.formula = function (a) { const c = this.cell(a); return c && c.f ? c.f : null; };
  /* Rekent een cel zelf uit (ook als Excel geen uitkomst bewaarde). Fout -> { error: '#DIV/0!' }. */
  Sheet.prototype.calc = function (a, seen) {
    a = String(a).toUpperCase();
    const c = this.cell(a);
    if (!c) return null;
    if (!c.f || !F) { if (c.t === 'e' && F) return new F.Err(c.v); return c.v; }
    this._memo = this._memo || {};
    if (a in this._memo) return this._memo[a];
    seen = seen || new Set();
    if (seen.has(this.name + '!' + a)) return c.v;
    seen.add(this.name + '!' + a);
    let v;
    try {
      v = F.evaluate(c.f, { locale: 'en', get: (ref, sheet) => {
        const sh = sheet ? (this.wb && this.wb.sheet(sheet)) : this;
        if (!sh) return new F.Err('#REF!');
        return sh.calc(ref, seen);
      } });
    } catch (e) { v = c.v; }
    if (F.isErr(v) && v.error === '#NAME?' && c.v != null && c.t !== 'e') v = c.v; // onbekende functie: vertrouw de uitkomst van Excel
    this._memo[a] = v;
    return v;
  };
  /* Ligt cel a in een sqref ("A1:B3 D5")? */
  function inSqref(sqref, a) {
    const p = parseAddr(a);
    return String(sqref || '').split(/\s+/).filter(Boolean).some(part => { const g = parseRange(part); return p.c >= g.c1 && p.c <= g.c2 && p.r >= g.r1 && p.r <= g.r2; });
  }
  Sheet.prototype.dvAt = function (a) { return this.dataValidations.filter(d => inSqref(d.sqref, a)); };
  Sheet.prototype.cfAt = function (a) { return this.conditionalFormats.filter(c => inSqref(c.sqref, a)).sort((x, y) => x.priority - y.priority); };
  /* Welke opmaak krijgt cel a volgens de voorwaardelijke opmaak, als er 'value' in zou staan? */
  Sheet.prototype.cfStyle = function (a, value) {
    const out = { fill: null, font: null, rules: [] };
    for (const r of this.cfAt(a)) {
      if (!cfMatch(this, r, a, value)) continue;
      out.rules.push(r);
      if (r.dxf) { if (!out.fill && r.dxf.fill) out.fill = r.dxf.fill; if (!out.font && r.dxf.font && r.dxf.font.color) out.font = r.dxf.font.color; }
    }
    return out;
  };
  function cfMatch(sheet, r, a, value) {
    if (!F) return false;
    const ev = f => F.evaluate(f, { locale: 'en', get: (ref, sh) => sh ? (sheet.wb && sheet.wb.sheet(sh) ? sheet.wb.sheet(sh).calc(ref) : null) : (ref === a ? value : sheet.calc(ref)) });
    const txt = v => v == null ? '' : String(v).toLowerCase();
    switch (r.type) {
      case 'cellIs': {
        if (F.isErr(value)) return false;
        const x = ev(r.formulas[0] || '0'), y = r.formulas[1] != null ? ev(r.formulas[1]) : null;
        const c = (p, q) => F.compare(p, q);
        switch (r.operator) {
          case 'equal': return c(value, x) === 0;
          case 'notEqual': return c(value, x) !== 0;
          case 'greaterThan': return c(value, x) > 0;
          case 'lessThan': return c(value, x) < 0;
          case 'greaterThanOrEqual': return c(value, x) >= 0;
          case 'lessThanOrEqual': return c(value, x) <= 0;
          case 'between': { const lo = c(x, y) <= 0 ? x : y, hi = lo === x ? y : x; return c(value, lo) >= 0 && c(value, hi) <= 0; }
          case 'notBetween': { const lo = c(x, y) <= 0 ? x : y, hi = lo === x ? y : x; return !(c(value, lo) >= 0 && c(value, hi) <= 0); }
        }
        return false;
      }
      case 'containsText': return txt(value).includes(txt(r.text));
      case 'notContainsText': return !txt(value).includes(txt(r.text));
      case 'beginsWith': return txt(value).startsWith(txt(r.text));
      case 'endsWith': return txt(value).endsWith(txt(r.text));
      case 'containsBlanks': return value == null || String(value).trim() === '';
      case 'notContainsBlanks': return !(value == null || String(value).trim() === '');
      case 'expression': {
        const first = String(r.sqref).split(/\s+/)[0].split(':')[0];
        const p = parseAddr(first), q = parseAddr(a);
        const v = ev(F.shift(r.formulas[0] || 'FALSE', q.r - p.r, q.c - p.c));
        return v === true || (typeof v === 'number' && v !== 0);
      }
    }
    return false;
  }
  /* De keuzes van een keuzelijst (gegevensvalidatie 'list'). */
  Sheet.prototype.dvListItems = function (dv) {
    let f = String(dv.formula1 || '').trim();
    if (/^".*"$/.test(f)) return f.slice(1, -1).split(/[,;]/).map(x => x.trim()).filter(Boolean);
    f = f.replace(/^=/, '');
    const m = /^(?:'?([^'!]+)'?!)?\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$/i.exec(f);
    if (!m) return null;
    const sh = m[1] ? (this.wb && this.wb.sheet(m[1])) : this;
    if (!sh) return null;
    const out = [];
    eachCell(m[2] + m[3] + ':' + (m[4] || m[2]) + (m[5] || m[3]), x => { const v = sh.calc(x); if (v != null && String(v).trim() !== '') out.push(String(v)); });
    return out;
  };
  Sheet.prototype.dvIsRange = function (dv) { return !/^".*"$/.test(String(dv.formula1 || '').trim()); };
  /* Wat staat er in de formule van deze cel? */
  Sheet.prototype.formulaInfo = function (a) { const f = this.formula(a); return f && F ? F.info(f, 'en') : null; };
  Sheet.prototype.numFmt = function (a) {
    const st = this.style(a);
    return numFmtInfo(st.numFmtId, this.styles.numFmts[st.numFmtId]);
  };
  Sheet.prototype.colWidth = function (colNum) {
    const c = this.cols.find(k => colNum >= k.min && colNum <= k.max && k.width != null);
    return c ? c.width : this.defaultColWidth;
  };
  Sheet.prototype.rowHeight = function (r) {
    return this.rowsInfo[r] && this.rowsInfo[r].ht ? this.rowsInfo[r].ht : this.defaultRowHeight;
  };
  Sheet.prototype.mergeAt = function (a) {
    const p = parseAddr(a);
    return this.merges.find(m => { const g = parseRange(m); return p.c >= g.c1 && p.c <= g.c2 && p.r >= g.r1 && p.r <= g.r2; }) || null;
  };
  /* Heeft deze cel een rand aan deze kant? (ook als de buurcel die rand tekent) */
  Sheet.prototype.hasEdge = function (a, side) {
    const p = parseAddr(a);
    if (this.style(a).border[side]) return true;
    const opp = { left: 'right', right: 'left', top: 'bottom', bottom: 'top' }[side];
    if ((side === 'left' && p.c === 1) || (side === 'top' && p.r === 1)) return false;
    const n = side === 'left' ? addr(p.c - 1, p.r) : side === 'right' ? addr(p.c + 1, p.r) : side === 'top' ? addr(p.c, p.r - 1) : addr(p.c, p.r + 1);
    return !!this.style(n).border[opp];
  };

  const BUILTIN_FMT = { 0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00', 14: 'd/m/yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 22: 'd/m/yyyy h:mm', 49: '@' };

  /* Wat voor getalnotatie is dit? { kind: general|number|percent|currency|date|text, dec } */
  function numFmtInfo(id, code) {
    id = +id || 0;
    if (id === 0 && (!code || /^general$/i.test(code))) return { kind: 'general', dec: null, code: 'General' };
    if (id === 9) return { kind: 'percent', dec: 0, code };
    if (id === 10) return { kind: 'percent', dec: 2, code };
    if ((id >= 14 && id <= 22) || (id >= 27 && id <= 36) || (id >= 45 && id <= 47) || (id >= 50 && id <= 58)) return { kind: 'date', dec: null, code };
    if ([5, 6, 7, 8, 42, 44].includes(id) && !/[0#]/.test(code || '')) return { kind: 'currency', dec: [5, 6, 42].includes(id) ? 0 : 2, code };
    if (id === 49 || code === '@') return { kind: 'text', dec: null, code };
    code = code || BUILTIN_FMT[id] || 'General';
    const first = code.split(';')[0];
    const euro = /€|EUR/i.test(first);
    const bare = first.replace(/"[^"]*"/g, '').replace(/\[^]/g, '').replace(/\[[^\]]*\]/g, '').replace(/_./g, '').replace(/\*./g, '');
    const decM = /0\.(0+)/.exec(bare);
    const dec = decM ? decM[1].length : 0;
    if (/%/.test(bare)) return { kind: 'percent', dec, code };
    if (euro) return { kind: 'currency', dec, code };
    if (/[dyj]/i.test(bare) || /m{3,}/i.test(bare) || /(^|[\/\-. ])m{1,2}([\/\-. ]|$)/i.test(bare)) return { kind: 'date', dec: null, code };
    if (/^general$/i.test(bare.trim())) return { kind: 'general', dec: null, code };
    if (/[0#]/.test(bare)) return { kind: 'number', dec, code };
    if (/@/.test(bare)) return { kind: 'text', dec: null, code };
    return { kind: 'general', dec: null, code };
  }
  /* Datum -> Excel-dagnummer (1 = 1 januari 1900). */
  function excelSerial(y, m, d) { return Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000); }
  function serialToDate(n) { const t = new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 86400000); return { y: t.getUTCFullYear(), m: t.getUTCMonth() + 1, d: t.getUTCDate() }; }

  function norm(s) { return String(s == null ? '' : s).replace(/\s+/g, ' ').trim().toLowerCase(); }

  /* Schatting van de breedte (in Excel-tekens) die nodig is voor een tekst in Calibri 11.
     Pixelbreedtes per letter; Excel rekent in "breedte van een cijfer" (7 px) plus wat marge. */
  const PX = { a:7,b:7,c:6,d:7,e:7,f:4,g:7,h:7,i:3,j:3,k:6,l:3,m:11,n:7,o:7,p:7,q:7,r:5,s:5,t:4,u:7,v:6,w:9,x:6,y:6,z:5,' ':3,'.':3,',':3,"'":2,'-':4,'_':7 };
  function textWidth(txt) {
    let px = 0;
    for (const ch of String(txt)) {
      if (PX[ch] != null) px += PX[ch];
      else if (ch >= '0' && ch <= '9') px += 7;
      else if (ch >= 'A' && ch <= 'Z') px += ch === 'M' || ch === 'W' ? 11 : ch === 'I' ? 3 : 8;
      else px += 7;
    }
    return (px + 7) / 7;
  }
  /* Is de kolom (ongeveer) passend gemaakt voor deze tekst? */
  function fitCheck(sheet, colNum, longest) {
    const w = sheet.colWidth(colNum);
    const need = textWidth(longest);
    if (w < need * 0.8) return { ok: false, w, need, why: 'smal' };
    if (w > need * 1.5 + 3) return { ok: false, w, need, why: 'breed' };
    return { ok: true, w, need };
  }

  const api = { numFmtInfo, excelSerial, serialToDate, readWorkbook, colToNum, numToCol, parseAddr, parseRange, addr, eachCell, colorNames, norm, textWidth, fitCheck, applyTint };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.XlsxCheck = api;
})(typeof window !== 'undefined' ? window : globalThis);
