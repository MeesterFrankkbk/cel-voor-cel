/* Cel voor Cel – maakt eenvoudige startbestanden (.xlsx) in de browser.
   spec = { sheetName, cells: { 'A1': { v: 'tekst' | 12, bold, size } }, colWidths: { A: 18 } } */
(function (root) {
  'use strict';
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const colToNum = c => [...c.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);

  async function buildXlsx(spec, JSZip, outType) {
    const fontKeys = ['0|11'];               // bold|size
    const fonts = [{ b: false, sz: 11 }];
    const xfFor = c => {
      const key = (c.bold ? 1 : 0) + '|' + (c.size || 11);
      let i = fontKeys.indexOf(key);
      if (i < 0) { fontKeys.push(key); fonts.push({ b: !!c.bold, sz: c.size || 11 }); i = fonts.length - 1; }
      return i; // één xf per lettertype
    };
    const strings = [], sIndex = {};
    const rows = {};
    Object.entries(spec.cells || {}).forEach(([a, c]) => {
      const m = /^([A-Z]+)(\d+)$/.exec(a);
      const r = +m[2];
      (rows[r] = rows[r] || []).push({ a, col: colToNum(m[1]), c: typeof c === 'object' && c !== null ? c : { v: c } });
    });
    let sheetData = '';
    Object.keys(rows).map(Number).sort((x, y) => x - y).forEach(r => {
      sheetData += `<row r="${r}">`;
      rows[r].sort((x, y) => x.col - y.col).forEach(({ a, c }) => {
        const s = xfFor(c);
        if (typeof c.v === 'number') sheetData += `<c r="${a}" s="${s}"><v>${c.v}</v></c>`;
        else {
          if (!(c.v in sIndex)) { sIndex[c.v] = strings.length; strings.push(c.v); }
          sheetData += `<c r="${a}" s="${s}" t="s"><v>${sIndex[c.v]}</v></c>`;
        }
      });
      sheetData += '</row>';
    });
    const cols = Object.entries(spec.colWidths || {}).map(([c, w]) => { const n = colToNum(c); return `<col min="${n}" max="${n}" width="${w}" customWidth="1"/>`; }).join('');

    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
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
      `<sheets><sheet name="${esc(spec.sheetName || 'Blad1')}" sheetId="1" r:id="rId1"/></sheets></workbook>`);
    zip.file('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' +
      '</Relationships>');
    zip.file('xl/worksheets/sheet1.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>' +
      (cols ? `<cols>${cols}</cols>` : '') + `<sheetData>${sheetData}</sheetData>` +
      '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>');
    zip.file('xl/sharedStrings.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${strings.length}" uniqueCount="${strings.length}">` +
      strings.map(s => `<si><t xml:space="preserve">${esc(s)}</t></si>`).join('') + '</sst>');
    zip.file('xl/styles.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      `<fonts count="${fonts.length}">` + fonts.map(f => `<font>${f.b ? '<b/>' : ''}<sz val="${f.sz}"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>`).join('') + '</fonts>' +
      '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
      '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      `<cellXfs count="${fonts.length}">` + fonts.map((f, i) => `<xf numFmtId="0" fontId="${i}" fillId="0" borderId="0" xfId="0"${i ? ' applyFont="1"' : ''}/>`).join('') + '</cellXfs>' +
      '<cellStyles count="1"><cellStyle name="Standaard" xfId="0" builtinId="0"/></cellStyles>' +
      '</styleSheet>');
    return zip.generateAsync({ type: outType || 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  const api = { buildXlsx };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.XlsxWrite = api;
})(typeof window !== 'undefined' ? window : globalThis);
