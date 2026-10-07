// Readers: turn each supported input file into a list of "pages" of plain text.
// Mirrors the extractors in file_to_file_conversion_tool.py, but runs entirely in the browser.
import { kids, loadScript, parseCsv, parseXml, splitExt } from './util.js';

const vendor = (p) => new URL(`../vendor/${p}`, import.meta.url).href;

export const INPUT_TYPES = {
  pdf: 'PDF', docx: 'Word', pptx: 'PowerPoint', xlsx: 'Excel', xlsm: 'Excel',
  csv: 'CSV', tsv: 'TSV', txt: 'Text', md: 'Text', log: 'Text', json: 'Text', xml: 'Text',
  png: 'Image', jpg: 'Image', jpeg: 'Image', bmp: 'Image', webp: 'Image', gif: 'Image',
};
const IMAGE_EXTS = new Set(['png', 'jpg', 'jpeg', 'bmp', 'webp', 'gif']);
const TEXT_EXTS = new Set(['txt', 'md', 'log', 'json', 'xml']);
// Old binary Office formats need a server-side converter; tell the user how to get around it
export const LEGACY_HINTS = { doc: 'docx', ppt: 'pptx', xls: 'xlsx' };

export function isSupported(name) {
  return splitExt(name)[1] in INPUT_TYPES;
}

// --- OCR (Tesseract.js), loaded on first use and shared across the whole batch ---

let ocrWorkerPromise = null;
let ocrProgress = () => {};

function getOcrWorker() {
  if (!ocrWorkerPromise) {
    ocrWorkerPromise = (async () => {
      await loadScript(vendor('tesseract/tesseract.min.js'));
      return window.Tesseract.createWorker('eng', 1, {
        workerPath: vendor('tesseract/worker.min.js'),
        corePath: vendor('tesseract/core'),
        langPath: vendor('tesseract/lang'),
        logger: (m) => ocrProgress(m),
      });
    })();
    ocrWorkerPromise.catch(() => { ocrWorkerPromise = null; });
  }
  return ocrWorkerPromise;
}

async function ocr(image, onStatus, what) {
  ocrProgress = (m) => {
    if (m.status === 'recognizing text') onStatus(`Reading text from ${what}… ${Math.round(m.progress * 100)}%`);
    else if (m.status) onStatus(`Preparing text recognition (${m.status})…`);
  };
  const worker = await getOcrWorker();
  const { data } = await worker.recognize(image);
  return data.text.trim();
}

export async function terminateOcr() {
  if (ocrWorkerPromise) {
    const w = await ocrWorkerPromise.catch(() => null);
    ocrWorkerPromise = null;
    if (w) await w.terminate();
  }
}

// --- PDF ---

let pdfjsPromise = null;
function getPdfjs() {
  if (!pdfjsPromise) {
    pdfjsPromise = import(vendor('pdfjs/pdf.min.mjs')).then((lib) => {
      lib.GlobalWorkerOptions.workerSrc = vendor('pdfjs/pdf.worker.min.mjs');
      return lib;
    });
  }
  return pdfjsPromise;
}

async function extractPdf(file, onStatus) {
  const pdfjs = await getPdfjs();
  const task = pdfjs.getDocument({
    data: new Uint8Array(await file.arrayBuffer()),
    cMapUrl: vendor('pdfjs/cmaps/'),
    standardFontDataUrl: vendor('pdfjs/standard_fonts/'),
    wasmUrl: vendor('pdfjs/wasm/'),
  });
  const pdf = await task.promise;

  const pages = [];
  try {
    for (let i = 1; i <= pdf.numPages; i++) {
      onStatus(`Reading page ${i} of ${pdf.numPages}…`);
      const page = await pdf.getPage(i);
      const content = await page.getTextContent();
      let text = '';
      for (const item of content.items) {
        if (typeof item.str !== 'string') continue;
        text += item.str;
        if (item.hasEOL) text += '\n';
      }
      // Same rule as the Python version: almost no text layer means a scanned page, so OCR it
      if (text.trim().length < 5) {
        const viewport = page.getViewport({ scale: 2 });
        const canvas = document.createElement('canvas');
        canvas.width = Math.ceil(viewport.width);
        canvas.height = Math.ceil(viewport.height);
        await page.render({ canvas, canvasContext: canvas.getContext('2d'), viewport }).promise;
        try {
          text = await ocr(canvas, onStatus, `scanned page ${i} of ${pdf.numPages}`);
        } catch (err) {
          console.warn(`OCR failed on page ${i}`, err);
        }
        canvas.width = canvas.height = 0;
      }
      pages.push(text.trim());
      page.cleanup();
    }
  } finally {
    await task.destroy();
  }
  return pages;
}

// --- Office Open XML helpers ---

async function openZip(file) {
  return window.JSZip.loadAsync(await file.arrayBuffer());
}

async function readPart(zip, path) {
  const entry = zip.file(path);
  if (!entry) throw new Error(`Missing ${path} — the file may be damaged or not a real Office file`);
  return parseXml(await entry.async('string'));
}

// Resolves a relationship Target against the folder of the part that owns it
function resolveTarget(baseDir, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = `${baseDir}/${target}`.split('/');
  const out = [];
  for (const p of parts) {
    if (p === '..') out.pop();
    else if (p && p !== '.') out.push(p);
  }
  return out.join('/');
}

async function readRels(zip, relsPath, baseDir) {
  const map = {};
  const entry = zip.file(relsPath);
  if (!entry) return map;
  const doc = parseXml(await entry.async('string'));
  for (const rel of doc.getElementsByTagName('Relationship')) {
    map[rel.getAttribute('Id')] = resolveTarget(baseDir, rel.getAttribute('Target'));
  }
  return map;
}

const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

// --- Word (.docx) ---

// Text of a Word paragraph: w:t runs, with tabs and line breaks kept
function wordRunText(el) {
  let text = '';
  for (const child of el.children) {
    switch (child.localName) {
      case 't': text += child.textContent; break;
      case 'tab': text += '\t'; break;
      case 'br': case 'cr': text += '\n'; break;
      case 'delText': case 'instrText': case 'rPr': case 'pPr': break;
      default: text += wordRunText(child);
    }
  }
  return text;
}

function wordBlocks(container, lines) {
  for (const el of container.children) {
    if (el.localName === 'p') lines.push(wordRunText(el));
    else if (el.localName === 'tbl') {
      for (const tr of kids(el, 'tr')) {
        const cells = kids(tr, 'tc').map((tc) => {
          const inner = [];
          wordBlocks(tc, inner);
          return inner.join(' ').replace(/\s+/g, ' ').trim();
        });
        lines.push(cells.join('\t'));
      }
    } else if (el.localName === 'sdt') {
      for (const content of kids(el, 'sdtContent')) wordBlocks(content, lines);
    }
  }
}

async function extractDocx(file) {
  const zip = await openZip(file);
  const doc = await readPart(zip, 'word/document.xml');
  const body = doc.getElementsByTagNameNS('*', 'body')[0];
  const lines = [];
  if (body) wordBlocks(body, lines);
  return [lines.join('\n')];
}

// --- PowerPoint (.pptx) ---

async function extractPptx(file) {
  const zip = await openZip(file);
  const pres = await readPart(zip, 'ppt/presentation.xml');
  const rels = await readRels(zip, 'ppt/_rels/presentation.xml.rels', 'ppt');
  let slidePaths = Array.from(pres.getElementsByTagNameNS('*', 'sldId'))
    .map((s) => rels[s.getAttributeNS(REL_NS, 'id')])
    .filter(Boolean);
  if (!slidePaths.length) {
    // Fall back to numeric order if the presentation's slide list can't be resolved
    slidePaths = Object.keys(zip.files)
      .filter((p) => /^ppt\/slides\/slide\d+\.xml$/.test(p))
      .sort((a, b) => parseInt(a.match(/\d+/)[0], 10) - parseInt(b.match(/\d+/)[0], 10));
  }

  const slides = [];
  for (const [i, path] of slidePaths.entries()) {
    const slide = await readPart(zip, path);
    const paras = Array.from(slide.getElementsByTagNameNS('*', 'p'))
      .filter((p) => p.namespaceURI.includes('drawingml'))
      .map((p) => Array.from(p.getElementsByTagNameNS('*', 't')).map((t) => t.textContent).join(''))
      .filter((t) => t.trim());
    slides.push(`--- SLIDE ${i + 1} ---\n${paras.join('\n')}`);
  }
  return slides;
}

// --- Excel (.xlsx) ---

const BUILTIN_DATE_FORMATS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58]);

function isDateFormat(code) {
  const stripped = code.replace(/"[^"]*"|\[[^\]]*\]|\\./g, '');
  return /[dmyhs]/i.test(stripped) && !/^general$/i.test(stripped);
}

function excelSerialToString(serial, date1904) {
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
  const d = new Date(epoch + Math.round(serial * 86400) * 1000);
  const iso = d.toISOString();
  return Number.isInteger(serial) ? iso.slice(0, 10) : `${iso.slice(0, 10)} ${iso.slice(11, 19)}`;
}

function colIndex(ref) {
  let n = 0;
  for (const ch of ref.replace(/\d+$/, '')) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

async function extractXlsx(file) {
  const zip = await openZip(file);
  const wb = await readPart(zip, 'xl/workbook.xml');
  const rels = await readRels(zip, 'xl/_rels/workbook.xml.rels', 'xl');
  const date1904 = ['1', 'true'].includes(wb.getElementsByTagNameNS('*', 'workbookPr')[0]?.getAttribute('date1904'));

  const shared = [];
  if (zip.file('xl/sharedStrings.xml')) {
    const sst = await readPart(zip, 'xl/sharedStrings.xml');
    for (const si of sst.getElementsByTagNameNS('*', 'si')) {
      // Skip phonetic (rPh) runs, which also contain <t> elements
      shared.push(Array.from(si.getElementsByTagNameNS('*', 't'))
        .filter((t) => t.parentElement.localName !== 'rPh')
        .map((t) => t.textContent).join(''));
    }
  }

  // Which cell styles are dates, so serial numbers can be shown as dates (as pandas does)
  const dateStyles = new Set();
  if (zip.file('xl/styles.xml')) {
    const styles = await readPart(zip, 'xl/styles.xml');
    const custom = {};
    for (const f of styles.getElementsByTagNameNS('*', 'numFmt')) {
      custom[f.getAttribute('numFmtId')] = f.getAttribute('formatCode') || '';
    }
    const cellXfs = styles.getElementsByTagNameNS('*', 'cellXfs')[0];
    if (cellXfs) {
      kids(cellXfs, 'xf').forEach((xf, i) => {
        const id = parseInt(xf.getAttribute('numFmtId') || '0', 10);
        if (BUILTIN_DATE_FORMATS.has(id) || (custom[id] && isDateFormat(custom[id]))) dateStyles.add(i);
      });
    }
  }

  const sheets = Array.from(wb.getElementsByTagNameNS('*', 'sheet'));
  const pages = [];
  for (const sheet of sheets) {
    const path = rels[sheet.getAttributeNS(REL_NS, 'id')];
    if (!path || !zip.file(path)) continue;
    const ws = await readPart(zip, path);
    const rows = [];
    for (const row of ws.getElementsByTagNameNS('*', 'row')) {
      const cells = [];
      let next = 0;
      for (const c of kids(row, 'c')) {
        const ref = c.getAttribute('r');
        const idx = ref ? colIndex(ref) : next;
        next = idx + 1;
        const type = c.getAttribute('t') || 'n';
        const v = kids(c, 'v')[0]?.textContent ?? '';
        let value;
        if (type === 's') value = shared[parseInt(v, 10)] ?? '';
        else if (type === 'inlineStr') value = Array.from(c.getElementsByTagNameNS('*', 't')).map((t) => t.textContent).join('');
        else if (type === 'b') value = v === '1' ? 'TRUE' : 'FALSE';
        else if (type === 'n' && v !== '' && dateStyles.has(parseInt(c.getAttribute('s') || '0', 10))) value = excelSerialToString(parseFloat(v), date1904);
        else value = v;
        cells[idx] = value.replace(/[\t\r\n]+/g, ' ');
      }
      rows.push(Array.from(cells, (x) => x ?? '').join('\t'));
    }
    while (rows.length && !rows[rows.length - 1].trim()) rows.pop();
    const body = rows.join('\n');
    pages.push(sheets.length > 1 ? `--- SHEET ${sheet.getAttribute('name')} ---\n${body}` : body);
  }
  return pages;
}

// --- Plain text, CSV and images ---

async function extractDelimited(file, delimiter) {
  const rows = parseCsv(await file.text(), delimiter);
  return [rows.map((r) => r.map((c) => c.replace(/[\t\r\n]+/g, ' ')).join('\t')).join('\n')];
}

async function extractImage(file, onStatus) {
  return [await ocr(file, onStatus, 'image')];
}

export async function extractText(file, onStatus = () => {}) {
  const ext = splitExt(file.name)[1];
  if (ext === 'pdf') return extractPdf(file, onStatus);
  if (ext === 'docx') return extractDocx(file);
  if (ext === 'pptx') return extractPptx(file);
  if (ext === 'xlsx' || ext === 'xlsm') return extractXlsx(file);
  if (ext === 'csv') return extractDelimited(file, ',');
  if (ext === 'tsv') return extractDelimited(file, '\t');
  if (IMAGE_EXTS.has(ext)) return extractImage(file, onStatus);
  if (TEXT_EXTS.has(ext)) return [await file.text()];
  throw new Error(`Unsupported file type .${ext}`);
}
