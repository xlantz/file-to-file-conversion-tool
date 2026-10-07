// Copies the browser libraries the static site needs from node_modules into web/vendor.
// Run by `npm run build` (Vercel / Cloudflare Pages run it automatically on deploy).
import { cpSync, mkdirSync, rmSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const nm = (...p) => join(root, 'node_modules', ...p);
const out = join(root, 'web', 'vendor');

const copies = [
  // pdf.js (legacy build = wider browser support): text extraction + rendering scanned pages for OCR
  [nm('pdfjs-dist/legacy/build/pdf.min.mjs'), 'pdfjs/pdf.min.mjs'],
  [nm('pdfjs-dist/legacy/build/pdf.worker.min.mjs'), 'pdfjs/pdf.worker.min.mjs'],
  [nm('pdfjs-dist/cmaps'), 'pdfjs/cmaps'],
  [nm('pdfjs-dist/standard_fonts'), 'pdfjs/standard_fonts'],
  [nm('pdfjs-dist/wasm'), 'pdfjs/wasm'],
  [nm('pdfjs-dist/LICENSE'), 'pdfjs/LICENSE'],
  // JSZip: reading .docx/.pptx/.xlsx and writing .docx/.xlsx/.zip
  [nm('jszip/dist/jszip.min.js'), 'jszip/jszip.min.js'],
  [nm('jszip/LICENSE.markdown'), 'jszip/LICENSE.markdown'],
  // jsPDF: writing .pdf
  [nm('jspdf/dist/jspdf.umd.min.js'), 'jspdf/jspdf.umd.min.js'],
  [nm('jspdf/LICENSE'), 'jspdf/LICENSE'],
  // Tesseract.js: OCR for images and scanned PDF pages (LSTM-only builds are all we load)
  [nm('tesseract.js/dist/tesseract.min.js'), 'tesseract/tesseract.min.js'],
  [nm('tesseract.js/dist/worker.min.js'), 'tesseract/worker.min.js'],
  [nm('tesseract.js/LICENSE.md'), 'tesseract/LICENSE.md'],
  [nm('tesseract.js-core/tesseract-core-lstm.wasm.js'), 'tesseract/core/tesseract-core-lstm.wasm.js'],
  [nm('tesseract.js-core/tesseract-core-simd-lstm.wasm.js'), 'tesseract/core/tesseract-core-simd-lstm.wasm.js'],
  [nm('tesseract.js-core/tesseract-core-relaxedsimd-lstm.wasm.js'), 'tesseract/core/tesseract-core-relaxedsimd-lstm.wasm.js'],
  [nm('@tesseract.js-data/eng/4.0.0_best_int/eng.traineddata.gz'), 'tesseract/lang/eng.traineddata.gz'],
];

rmSync(out, { recursive: true, force: true });
for (const [src, dest] of copies) {
  const target = join(out, dest);
  mkdirSync(dirname(target), { recursive: true });
  cpSync(src, target, { recursive: true });
}
console.log(`Copied ${copies.length} vendor assets to web/vendor`);
