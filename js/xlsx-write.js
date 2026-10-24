/* Cel voor Cel – maakt eenvoudige startbestanden (.xlsx) in de browser.
   spec = { sheetName, cells, colWidths }  of  { sheets: [ { name, cells, colWidths } ] }
   cel: 'tekst' | 12 | { v, f (formule zonder =, Engels), bold, size, fmt: '0%' | '0.00' } */
(function (root) {
  'use strict';
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const colToNum = c => [...c.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  const BUILTIN = { 'General': 0, '0': 1, '0.00': 2, '0%': 9, '0.00%': 10 };

  async function buildXlsx(spec, JSZip, outType) {
    const sheets = spec.sheets || [{ name: spec.sheetName || 'Blad1', cells: spec.cells, colWidths: spec.colWidths }];
    // stijlen: één combinatie (vet, grootte, notatie) = één xf
    const fonts = [{ b: false, sz: 11 }], fontKeys = ['0|11'];
    const xfs = [{ font: 0, fmt: 0 }], xfKeys = ['0|0'];
    const custom = {}; let nextFmt = 164;
    const fmtId = f => { if (!f) return 0; if (f in BUILTIN) return BUILTIN[f]; if (!custom[f]) custom[f] = nextFmt++; return custom[f]; };
    const xfFor = c => {
      const fk = (c.bold ? 1 : 0) + '|' + (c.size || 11);
      let fi = fontKeys.indexOf(fk);
      if (fi < 0) { fontKeys.push(fk); fonts.push({ b: !!c.bold, sz: c.size || 11 }); fi = fonts.length - 1; }
      const nf = fmtId(c.fmt), xk = fi + '|' + nf;
      let xi = xfKeys.indexOf(xk);
      if (xi < 0) { xfKeys.push(xk); xfs.push({ font: fi, fmt: nf }); xi = xfs.length - 1; }
      return xi;
    };
    const strings = [], sIndex = {};
    const sheetXml = sheets.map(sh => {
      const rows = {};
      Object.entries(sh.cells || {}).forEach(([a, c]) => {
        const m = /^([A-Z]+)(\d+)$/.exec(a);
        (rows[+m[2]] = rows[+m[2]] || []).push({ a, col: colToNum(m[1]), c: typeof c === 'object' && c !== null ? c : { v: c } });
      });
      let data = '';
      Object.keys(rows).map(Number).sort((x, y) => x - y).forEach(r => {
        data += `<row r="${r}">`;
        rows[r].sort((x, y) => x.col - y.col).forEach(({ a, c }) => {
          const s = xfFor(c);
          if (c.f) data += `<c r="${a}" s="${s}"><f>${esc(c.f)}</f></c>`;
          else if (c.v == null) data += `<c r="${a}" s="${s}"/>`;
          else if (typeof c.v === 'number') data += `<c r="${a}" s="${s}"><v>${c.v}</v></c>`;
          else {
            if (!(c.v in sIndex)) { sIndex[c.v] = strings.length; strings.push(c.v); }
            data += `<c r="${a}" s="${s}" t="s"><v>${sIndex[c.v]}</v></c>`;
          }
        });
        data += '</row>';
      });
      const cols = Object.entries(sh.colWidths || {}).map(([c, w]) => { const n = colToNum(c); return `<col min="${n}" max="${n}" width="${w}" customWidth="1"/>`; }).join('');
      return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>' +
        (cols ? `<cols>${cols}</cols>` : '') + `<sheetData>${data}</sheetData>` +
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>';
    });

    const zip = new JSZip();
    const n = sheets.length;
    zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('') +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
      '</Types>');
    zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
      '</Relationships>');
    const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    zip.file('docProps/core.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      '<dc:creator>Cel voor Cel</dc:creator>' +
      `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>` +
      '</cp:coreProperties>');
    zip.file('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets>' + sheets.map((sh, i) => `<sheet name="${esc(sh.name || 'Blad' + (i + 1))}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('') + '</sheets>' +
      '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>');
    zip.file('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('') +
      `<Relationship Id="rId${n + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
      `<Relationship Id="rId${n + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>` +
      '</Relationships>');
    sheetXml.forEach((x, i) => zip.file(`xl/worksheets/sheet${i + 1}.xml`, x));
    zip.file('xl/sharedStrings.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${strings.length}" uniqueCount="${strings.length}">` +
      strings.map(s => `<si><t xml:space="preserve">${esc(s)}</t></si>`).join('') + '</sst>');
    const numFmts = Object.entries(custom);
    zip.file('xl/styles.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      (numFmts.length ? `<numFmts count="${numFmts.length}">` + numFmts.map(([code, id]) => `<numFmt numFmtId="${id}" formatCode="${esc(code)}"/>`).join('') + '</numFmts>' : '') +
      `<fonts count="${fonts.length}">` + fonts.map(f => `<font>${f.b ? '<b/>' : ''}<sz val="${f.sz}"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>`).join('') + '</fonts>' +
      '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
      '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      `<cellXfs count="${xfs.length}">` + xfs.map((x, i) => `<xf numFmtId="${x.fmt}" fontId="${x.font}" fillId="0" borderId="0" xfId="0"${x.font ? ' applyFont="1"' : ''}${x.fmt ? ' applyNumberFormat="1"' : ''}/>`).join('') + '</cellXfs>' +
      '<cellStyles count="1"><cellStyle name="Standaard" xfId="0" builtinId="0"/></cellStyles>' +
      '</styleSheet>');
    return zip.generateAsync({ type: outType || 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  const api = { buildXlsx };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.XlsxWrite = api;
})(typeof window !== 'undefined' ? window : globalThis);
