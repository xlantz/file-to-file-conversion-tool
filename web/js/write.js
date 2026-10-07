// Writers: build the output files (as Blobs) from extracted text.
import { OUTPUT_FORMATS, escapeXml, textToRows, toCsv } from './util.js';

// --- Word (.docx), assembled by hand with JSZip ---

const DOCX_CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;

const PACKAGE_RELS = (target) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="${target}"/>
</Relationships>`;

const DOCX_DOC_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

const DOCX_STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="264" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="160"/><w:outlineLvl w:val="0"/></w:pPr>
<w:rPr><w:b/><w:color w:val="1F3864"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>
</w:styles>`;

function docxRuns(text) {
  return text.split('\t')
    .map((part) => `<w:t xml:space="preserve">${escapeXml(part)}</w:t>`)
    .join('<w:tab/>');
}

function docxParagraph(text, style) {
  const pPr = style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : '';
  return `<w:p>${pPr}<w:r>${docxRuns(text)}</w:r></w:p>`;
}

const DOCX_PAGE_BREAK = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';

// sections: [{ heading?: string, pages: string[] }] — each section starts on a new page
async function buildDocx(sections) {
  const body = [];
  sections.forEach((section, si) => {
    if (si > 0) body.push(DOCX_PAGE_BREAK);
    if (section.heading) body.push(docxParagraph(section.heading, 'Heading1'));
    section.pages.forEach((page, pi) => {
      if (pi > 0) body.push(DOCX_PAGE_BREAK);
      for (const line of page.split(/\r?\n/)) body.push(docxParagraph(line));
    });
  });

  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body.join('')}
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>
</w:body></w:document>`;

  const zip = new window.JSZip();
  zip.file('[Content_Types].xml', DOCX_CONTENT_TYPES);
  zip.file('_rels/.rels', PACKAGE_RELS('word/document.xml'));
  zip.file('word/_rels/document.xml.rels', DOCX_DOC_RELS);
  zip.file('word/styles.xml', DOCX_STYLES);
  zip.file('word/document.xml', documentXml);
  return zip.generateAsync({ type: 'blob', mimeType: OUTPUT_FORMATS.word.mime, compression: 'DEFLATE' });
}

// --- Excel (.xlsx), assembled by hand with JSZip ---

function colName(i) {
  let s = '';
  for (i += 1; i > 0; i = Math.floor((i - 1) / 26)) s = String.fromCharCode(65 + ((i - 1) % 26)) + s;
  return s;
}

function sheetXml(rows) {
  const out = rows.map((row, r) => {
    const cells = row.map((value, c) => {
      if (value === '') return '';
      const ref = `${colName(c)}${r + 1}`;
      // Store plain numbers as numbers (but keep things like ZIP codes "00501" as text)
      if (/^-?(0|[1-9]\d{0,14})(\.\d+)?$/.test(value)) return `<c r="${ref}"><v>${value}</v></c>`;
      return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(value.slice(0, 32767))}</t></is></c>`;
    }).join('');
    return `<row r="${r + 1}">${cells}</row>`;
  }).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${out}</sheetData></worksheet>`;
}

// Excel sheet names: max 31 chars, no []:*?/\ , unique (case-insensitive)
function uniqueSheetName(name, used) {
  const base = name.replace(/[[\]:*?/\\]/g, '_').replace(/^'+|'+$/g, '').slice(0, 31) || 'Sheet';
  let candidate = base;
  for (let n = 2; used.has(candidate.toLowerCase()); n++) {
    const suffix = `_${n}`;
    candidate = base.slice(0, 31 - suffix.length) + suffix;
  }
  used.add(candidate.toLowerCase());
  return candidate;
}

// sheets: [{ name, rows: string[][] }]
async function buildXlsx(sheets) {
  const used = new Set();
  const named = sheets.map((s) => ({ ...s, name: uniqueSheetName(s.name, used) }));
  const zip = new window.JSZip();
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
${named.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('\n')}
</Types>`);
  zip.file('_rels/.rels', PACKAGE_RELS('xl/workbook.xml'));
  zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>
${named.map((s, i) => `<sheet name="${escapeXml(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}
</sheets></workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${named.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('\n')}
</Relationships>`);
  named.forEach((s, i) => zip.file(`xl/worksheets/sheet${i + 1}.xml`, sheetXml(s.rows)));
  return zip.generateAsync({ type: 'blob', mimeType: OUTPUT_FORMATS.excel.mime, compression: 'DEFLATE' });
}

// --- PDF, via jsPDF ---

// jsPDF's built-in fonts only cover Windows-1252; replace anything else so the PDF stays readable
const CP1252_EXTRAS = '€‚ƒ„…†‡ˆ‰Š‹ŒŽ‘’“”•–—˜™š›œžŸ';
function pdfSafe(text) {
  return text.replace(/\t/g, '    ').replace(/[^\n\x20-\x7E\xA0-\xFF]/g, (ch) => (CP1252_EXTRAS.includes(ch) ? ch : '?'));
}

function buildPdf(sections) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: 'pt', format: 'letter' });
  const margin = 54;
  const width = doc.internal.pageSize.getWidth() - margin * 2;
  const bottom = doc.internal.pageSize.getHeight() - margin;
  let y = margin;
  let first = true;

  const newPage = () => {
    if (!first) doc.addPage();
    first = false;
    y = margin;
  };
  const write = (text, size, bold, lineHeight) => {
    doc.setFont('helvetica', bold ? 'bold' : 'normal');
    doc.setFontSize(size);
    for (const line of doc.splitTextToSize(pdfSafe(text), width)) {
      if (y + lineHeight > bottom) { doc.addPage(); y = margin; }
      doc.text(line, margin, y + size);
      y += lineHeight;
    }
  };

  for (const section of sections) {
    newPage();
    if (section.heading) {
      write(section.heading, 15, true, 20);
      y += 8;
    }
    section.pages.forEach((page, pi) => {
      if (pi > 0) newPage();
      for (const line of page.split(/\r?\n/)) write(line || ' ', 10.5, false, 14);
    });
  }
  if (first) newPage();
  return doc.output('blob');
}

// --- Public API ---

// One output file from one input (single-file and "separate files" modes)
export async function writeSingle(format, pages, sourceFile) {
  const text = pages.join('\n\n');
  switch (format) {
    case 'txt':
      return new Blob([text], { type: 'text/plain;charset=utf-8' });
    case 'word':
      return buildDocx([{ pages }]);
    case 'csv':
      // BOM so Excel opens UTF-8 CSVs with the right characters
      return new Blob(['﻿', toCsv(textToRows(text))], { type: 'text/csv;charset=utf-8' });
    case 'excel':
      return buildXlsx([{ name: 'Sheet1', rows: textToRows(text) }]);
    case 'pdf':
      // A PDF converted to PDF is just copied, as in the Python version
      if (sourceFile && /\.pdf$/i.test(sourceFile.name)) return new Blob([await sourceFile.arrayBuffer()], { type: 'application/pdf' });
      return buildPdf([{ pages }]);
    default:
      throw new Error(`Unknown format ${format}`);
  }
}

// One output file containing every input, each labelled with its source file
// results: [{ label, pages }]
export async function writeCombined(format, results) {
  switch (format) {
    case 'txt': {
      const text = results.map((r) => `===== ${r.label} =====\n\n${r.pages.join('\n\n')}`).join('\n\n');
      return new Blob([text], { type: 'text/plain;charset=utf-8' });
    }
    case 'word':
      return buildDocx(results.map((r) => ({ heading: r.label, pages: r.pages })));
    case 'pdf':
      return buildPdf(results.map((r) => ({ heading: r.label, pages: r.pages })));
    case 'csv': {
      // First column says which file each row came from
      const rows = [];
      for (const r of results) {
        for (const row of textToRows(r.pages.join('\n\n'))) rows.push([r.label, ...row]);
      }
      return new Blob(['﻿', toCsv(rows)], { type: 'text/csv;charset=utf-8' });
    }
    case 'excel':
      // One worksheet per source file
      return buildXlsx(results.map((r) => ({
        name: r.label.split('/').pop().replace(/\.[^.]+$/, ''),
        rows: textToRows(r.pages.join('\n\n')),
      })));
    default:
      throw new Error(`Unknown format ${format}`);
  }
}

export async function buildZip(entries) {
  const zip = new window.JSZip();
  for (const { path, blob } of entries) zip.file(path, blob);
  return zip.generateAsync({ type: 'blob', compression: 'DEFLATE' });
}
