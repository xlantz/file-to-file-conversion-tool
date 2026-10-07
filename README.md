# file-to-file-conversion-tool

Convert a file — or a whole folder of files — to **Word, Excel, CSV, text or PDF**.
A folder can be exported as **one combined file** or as **separate files** (downloaded as a .zip).

There are two ways to use it:

| | Web app (`web/`) | Python script (`file_to_file_conversion_tool.py`) |
|---|---|---|
| Runs | In the browser — nothing is uploaded | On your computer, from a terminal |
| Hosting cost | **$0** (static site) | n/a |
| Setup | None for users — just open the URL | Python + pip packages + Poppler |

## Web app

Everything happens in the visitor's browser: pdf.js reads PDFs, Tesseract.js does OCR on
images and scanned PDF pages, and the output files are generated locally. There is no server
or API, so hosting is free, there are no upload size limits, and files stay private.

**Reads:** PDF (including scanned pages), .docx, .pptx, .xlsx, CSV/TSV, text, and images
(PNG, JPG, BMP, WebP, GIF). Old binary .doc/.xls/.ppt files must be re-saved in the newer format.

**Combined output** labels each source file: a heading per file in Word/PDF, a
`===== file =====` divider in text, a source-file column in CSV, one worksheet per file in Excel.

### Deploy for free

The site is static. The build step only copies the browser libraries out of `node_modules`
into `web/vendor/` (about 21 MB, mostly the OCR engine), so nothing depends on a third-party CDN.

**Vercel (Hobby plan, free — personal / non-commercial use)**
1. Go to vercel.com → *Add New… → Project* → import this GitHub repo.
2. Leave everything as detected — `vercel.json` already sets the build command
   (`npm run build`) and output directory (`web`). Click *Deploy*.

**Cloudflare Pages (free, commercial use allowed, unlimited bandwidth)**
1. dash.cloudflare.com → *Workers & Pages → Create → Pages → Connect to Git* → pick this repo.
2. Build command: `npm run build` · Build output directory: `web` · Deploy.

Both redeploy automatically on every push.

### Run locally

```bash
npm install
npm run build
python3 -m http.server 8000 --directory web   # then open http://localhost:8000
```

(Opening `web/index.html` directly from disk won't work — browsers block module scripts on `file://`.)

## Python script

```bash
pip install pdfplumber pandas openpyxl python-docx python-pptx pillow pdf2image easyocr
python file_to_file_conversion_tool.py
```

Enter a file or folder path when prompted; for a folder, choose whether to include
subfolders and whether to export one combined file or separate files.
