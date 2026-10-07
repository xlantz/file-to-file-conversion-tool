// Small helpers shared by the readers and writers.

export const OUTPUT_FORMATS = {
  word: { ext: 'docx', label: 'Word', mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
  excel: { ext: 'xlsx', label: 'Excel', mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
  csv: { ext: 'csv', label: 'CSV', mime: 'text/csv' },
  txt: { ext: 'txt', label: 'Text', mime: 'text/plain' },
  pdf: { ext: 'pdf', label: 'PDF', mime: 'application/pdf' },
};

export function splitExt(name) {
  const i = name.lastIndexOf('.');
  return i > 0 ? [name.slice(0, i), name.slice(i + 1).toLowerCase()] : [name, ''];
}

export function sanitizeFilename(name) {
  return name.replace(/[\\/*?:"<>|]/g, '').trim();
}

// Characters XML 1.0 can't contain; OCR and PDF text occasionally produce them
const INVALID_XML = /[\x00-\x08\x0B\x0C\x0E-\x1F￾￿]/g;

export function escapeXml(text) {
  return String(text)
    .replace(INVALID_XML, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

export function escapeHtml(text) {
  return escapeXml(text).replace(/'/g, '&#39;');
}

// Splits extracted text into table rows. Tab-separated lines (spreadsheets, Word tables)
// split on tabs so empty cells survive; other lines split on runs of 2+ spaces.
export function textToRows(text) {
  const rows = [];
  for (const line of text.split(/\r?\n/)) {
    if (!line.trim()) continue;
    let cells = line.includes('\t')
      ? line.split('\t').map((c) => c.trim())
      : line.trim().split(/\s{2,}/);
    while (cells.length > 1 && cells[cells.length - 1] === '') cells.pop();
    rows.push(cells);
  }
  return rows;
}

// Minimal RFC 4180 CSV parser (handles quoted fields, escaped quotes and newlines in quotes)
export function parseCsv(text, delimiter = ',') {
  const rows = [];
  let row = [];
  let field = '';
  let quoted = false;
  text = text.replace(/^﻿/, '');
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"') {
        if (text[i + 1] === '"') { field += '"'; i++; } else quoted = false;
      } else field += ch;
    } else if (ch === '"') quoted = true;
    else if (ch === delimiter) { row.push(field); field = ''; }
    else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      row.push(field); rows.push(row); row = []; field = '';
    } else field += ch;
  }
  if (field !== '' || row.length) { row.push(field); rows.push(row); }
  return rows;
}

export function toCsv(rows) {
  const width = Math.max(0, ...rows.map((r) => r.length));
  const esc = (v) => (/[",\r\n]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v);
  // Pad rows so every line has the same column count, like pandas does
  return rows.map((r) => Array.from({ length: width }, (_, i) => esc(r[i] ?? '')).join(',')).join('\r\n');
}

// Zip-safe unique names: report.pdf + report.docx -> report.txt + report_docx.txt
export function makeUniqueNamer() {
  const used = new Set();
  return (dir, stem, srcExt, outExt) => {
    let name = stem;
    let n = 2;
    const key = () => `${dir}/${name}.${outExt}`.toLowerCase();
    if (used.has(key())) name = `${stem}_${srcExt}`;
    while (used.has(key())) name = `${stem}_${srcExt}_${n++}`;
    used.add(key());
    return dir ? `${dir}/${name}.${outExt}` : `${name}.${outExt}`;
  };
}

export function formatBytes(n) {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / 1024 / 1024).toFixed(1)} MB`;
}

export function loadScript(src) {
  return new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = src;
    s.onload = resolve;
    s.onerror = () => reject(new Error(`Failed to load ${src}`));
    document.head.appendChild(s);
  });
}

// Children of an XML element by local name (ignores namespace prefixes)
export function kids(el, name) {
  return Array.from(el.children).filter((c) => c.localName === name);
}

export function parseXml(text) {
  const doc = new DOMParser().parseFromString(text, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) throw new Error('File contains malformed XML');
  return doc;
}
