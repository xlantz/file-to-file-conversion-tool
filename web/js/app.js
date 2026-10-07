import { INPUT_TYPES, LEGACY_HINTS, extractText, isSupported, terminateOcr } from './extract.js';
import { buildZip, writeCombined, writeSingle } from './write.js';
import { OUTPUT_FORMATS, escapeHtml, formatBytes, makeUniqueNamer, sanitizeFilename, splitExt } from './util.js';

const $ = (id) => document.getElementById(id);
const els = {
  dropzone: $('dropzone'), folderInput: $('folder-input'), fileInput: $('file-input'),
  selection: $('selection'), summary: $('selection-summary'), list: $('file-list'), skipped: $('skipped'),
  subfoldersWrap: $('subfolders-wrap'), subfolders: $('subfolders'),
  modeField: $('mode-field'), formatHint: $('format-hint'),
  outName: $('out-name'), outNameLabel: $('out-name-label'), outExt: $('out-ext'),
  convert: $('convert'), resultCard: $('result-card'), resultTitle: $('result-title'),
  progress: $('progress'), progressBar: $('progress-bar'), status: $('status'),
  download: $('download'), failures: $('failures'),
};

const state = {
  entries: [], // { file, path } — path is relative, starting with the folder name for folder picks
  nameEdited: false,
  busy: false,
  downloadUrl: null,
};

const getMode = () => document.querySelector('input[name=mode]:checked').value;
const getFormat = () => document.querySelector('input[name=format]:checked').value;
const hasFolders = () => state.entries.some((e) => e.path.includes('/'));

// Office lock files (~$report.docx) and hidden files (.DS_Store) are never wanted
const isJunk = (name) => name.startsWith('~$') || name.startsWith('.');

function partition() {
  const included = [];
  const unsupported = [];
  for (const entry of state.entries) {
    const name = entry.file.name;
    if (isJunk(name)) continue;
    // Without subfolders, keep only files directly inside the chosen folder ("Folder/file.pdf")
    if (!els.subfolders.checked && entry.path.split('/').length > 2) continue;
    (isSupported(name) ? included : unsupported).push(entry);
  }
  included.sort((a, b) => a.path.localeCompare(b.path, undefined, { numeric: true }));
  return { included, unsupported };
}

// Display labels: paths relative to the chosen folder
function labelsFor(entries) {
  const firstSegs = new Set(entries.map((e) => (e.path.includes('/') ? e.path.split('/')[0] : '')));
  const strip = firstSegs.size === 1 && !firstSegs.has('');
  return entries.map((e) => (strip ? e.path.slice(e.path.indexOf('/') + 1) : e.path));
}

function defaultName(included) {
  if (included.length === 1) return splitExt(included[0].file.name)[0];
  const roots = new Set(included.map((e) => (e.path.includes('/') ? e.path.split('/')[0] : '')));
  if (roots.size === 1 && !roots.has('')) return `${[...roots][0]} converted`;
  return 'converted';
}

function renderList() {
  const { included, unsupported } = partition();
  const labels = labelsFor(included);

  els.selection.hidden = state.entries.length === 0;
  els.subfoldersWrap.hidden = !state.entries.some((e) => e.path.split('/').length > 2);
  els.summary.textContent = included.length
    ? `${included.length} file${included.length === 1 ? '' : 's'} ready to convert`
    : 'No supported files selected';

  els.list.innerHTML = included.map((e, i) => {
    const ext = splitExt(e.file.name)[1];
    return `<li data-index="${i}"><span class="file-name">${escapeHtml(labels[i])}</span>
      <span class="file-meta">${INPUT_TYPES[ext]} · ${formatBytes(e.file.size)}</span>
      <span class="file-status" hidden></span></li>`;
  }).join('');

  if (unsupported.length) {
    const legacy = unsupported.filter((e) => splitExt(e.file.name)[1] in LEGACY_HINTS);
    let msg = `Skipping ${unsupported.length} unsupported file${unsupported.length === 1 ? '' : 's'}: `
      + unsupported.slice(0, 8).map((e) => e.file.name).join(', ')
      + (unsupported.length > 8 ? `, and ${unsupported.length - 8} more` : '') + '.';
    if (legacy.length) msg += ' Older .doc/.xls/.ppt files can be converted after re-saving them as .docx/.xlsx/.pptx.';
    els.skipped.textContent = msg;
    els.skipped.hidden = false;
  } else {
    els.skipped.hidden = true;
  }
}

function renderOptions() {
  const { included } = partition();
  const multi = included.length > 1;
  els.modeField.hidden = !multi;
  const separate = multi && getMode() === 'separate';
  const format = getFormat();
  els.outNameLabel.textContent = separate ? 'Zip file name' : 'File name';
  els.outExt.textContent = separate ? '.zip' : `.${OUTPUT_FORMATS[format].ext}`;
  if (!state.nameEdited) els.outName.value = included.length ? defaultName(included) : '';

  const hints = {
    word: multi && !separate ? 'Each file gets its own heading and starts on a new page.' : '',
    excel: multi && !separate ? 'Each file goes on its own worksheet.' : 'Columns are detected from tabs and runs of spaces.',
    csv: multi && !separate ? 'The first column shows which file each row came from.' : 'Columns are detected from tabs and runs of spaces.',
    txt: multi && !separate ? 'Each file is separated by a ===== file name ===== line.' : '',
    pdf: 'PDF inputs are copied as-is in separate mode. Text uses standard fonts, so non-Western characters show as “?”.',
  };
  els.formatHint.textContent = hints[format];
  els.formatHint.hidden = !hints[format];

  els.convert.disabled = state.busy || included.length === 0;
  els.convert.textContent = included.length > 1 ? `Convert ${included.length} files` : 'Convert';
}

function render() {
  renderList();
  renderOptions();
}

function setEntries(entries) {
  if (state.busy) return;
  state.entries = entries;
  state.nameEdited = false;
  els.resultCard.hidden = true;
  render();
}

// --- Picking files ---

els.dropzone.querySelector('#pick-folder').addEventListener('click', () => els.folderInput.click());
els.dropzone.querySelector('#pick-files').addEventListener('click', () => els.fileInput.click());

for (const input of [els.folderInput, els.fileInput]) {
  input.addEventListener('change', () => {
    setEntries(Array.from(input.files, (file) => ({ file, path: file.webkitRelativePath || file.name })));
    input.value = '';
  });
}

// Reads a dropped folder recursively (readEntries returns results in batches)
async function walkEntry(entry, prefix, out) {
  if (entry.isFile) {
    const file = await new Promise((res, rej) => entry.file(res, rej));
    out.push({ file, path: prefix + file.name });
  } else if (entry.isDirectory) {
    const reader = entry.createReader();
    let batch;
    do {
      batch = await new Promise((res, rej) => reader.readEntries(res, rej));
      for (const child of batch) await walkEntry(child, `${prefix}${entry.name}/`, out);
    } while (batch.length);
  }
}

['dragenter', 'dragover'].forEach((type) => els.dropzone.addEventListener(type, (e) => {
  e.preventDefault();
  if (!state.busy) els.dropzone.classList.add('over');
}));
['dragleave', 'drop'].forEach((type) => els.dropzone.addEventListener(type, (e) => {
  if (type === 'dragleave' && els.dropzone.contains(e.relatedTarget)) return;
  els.dropzone.classList.remove('over');
}));
els.dropzone.addEventListener('drop', async (e) => {
  e.preventDefault();
  if (state.busy) return;
  const items = Array.from(e.dataTransfer.items || []);
  const roots = items.map((it) => it.webkitGetAsEntry?.()).filter(Boolean);
  if (!roots.length) {
    setEntries(Array.from(e.dataTransfer.files, (file) => ({ file, path: file.name })));
    return;
  }
  const out = [];
  for (const root of roots) await walkEntry(root, '', out);
  setEntries(out);
});
// Stop a stray drop outside the box from navigating away to the file
window.addEventListener('dragover', (e) => e.preventDefault());
window.addEventListener('drop', (e) => e.preventDefault());

$('clear').addEventListener('click', () => setEntries([]));
els.subfolders.addEventListener('change', render);
document.querySelectorAll('input[name=mode], input[name=format]').forEach((r) => r.addEventListener('change', renderOptions));
els.outName.addEventListener('input', () => { state.nameEdited = true; });

// --- Converting ---

function setRowStatus(index, cls, text) {
  const li = els.list.querySelector(`li[data-index="${index}"]`);
  if (!li) return;
  li.className = cls;
  const status = li.querySelector('.file-status');
  status.hidden = !text;
  status.textContent = text;
  if (cls === 'working') li.scrollIntoView({ block: 'nearest' });
}

function setProgress(fraction, text) {
  const pct = Math.round(fraction * 100);
  els.progressBar.style.width = `${pct}%`;
  els.progress.setAttribute('aria-valuenow', String(pct));
  if (text) els.status.textContent = text;
}

function offerDownload(blob, filename) {
  if (state.downloadUrl) URL.revokeObjectURL(state.downloadUrl);
  state.downloadUrl = URL.createObjectURL(blob);
  els.download.href = state.downloadUrl;
  els.download.download = filename;
  els.download.textContent = `Download ${filename} (${formatBytes(blob.size)})`;
  els.download.hidden = false;
  els.download.click();
}

els.convert.addEventListener('click', async () => {
  const { included } = partition();
  if (!included.length || state.busy) return;
  const labels = labelsFor(included);
  const format = getFormat();
  const ext = OUTPUT_FORMATS[format].ext;
  const mode = included.length > 1 ? getMode() : 'single';
  const baseName = sanitizeFilename(els.outName.value) || defaultName(included);

  state.busy = true;
  document.body.classList.add('busy');
  els.convert.disabled = true;
  els.resultCard.hidden = false;
  els.resultTitle.textContent = 'Converting…';
  els.download.hidden = true;
  els.failures.hidden = true;
  setProgress(0, 'Starting…');
  els.resultCard.scrollIntoView({ behavior: 'smooth', block: 'nearest' });

  const results = [];
  const zipEntries = [];
  const failed = [];
  const uniqueName = makeUniqueNamer();

  try {
    for (const [i, entry] of included.entries()) {
      const label = labels[i];
      const prefix = included.length > 1 ? `File ${i + 1} of ${included.length}: ${label}` : label;
      setRowStatus(i, 'working', 'Converting…');
      setProgress(i / included.length, prefix);
      try {
        const pages = await extractText(entry.file, (msg) => {
          setRowStatus(i, 'working', msg);
          els.status.textContent = `${prefix} — ${msg}`;
        });
        if (mode === 'separate') {
          const [stem, srcExt] = splitExt(entry.file.name);
          const dir = label.includes('/') ? label.slice(0, label.lastIndexOf('/')) : '';
          zipEntries.push({ path: uniqueName(dir, stem, srcExt, ext), blob: await writeSingle(format, pages, entry.file) });
        } else {
          results.push({ label, pages, file: entry.file });
        }
        const empty = pages.every((p) => !p.trim());
        setRowStatus(i, 'done', empty ? 'Done — no text found' : 'Done');
      } catch (err) {
        console.error(err);
        failed.push({ label, message: err.message || String(err) });
        setRowStatus(i, 'failed', `Couldn't convert: ${err.message || err}`);
      }
    }

    setProgress(1, 'Saving…');
    if (mode === 'single' && results.length) {
      offerDownload(await writeSingle(format, results[0].pages, results[0].file), `${baseName}.${ext}`);
    } else if (mode === 'combined' && results.length) {
      offerDownload(await writeCombined(format, results), `${baseName}.${ext}`);
    } else if (mode === 'separate' && zipEntries.length) {
      offerDownload(await buildZip(zipEntries), `${baseName}.zip`);
    }

    const ok = included.length - failed.length;
    if (ok === 0) {
      els.resultTitle.textContent = 'Nothing could be converted';
      els.status.textContent = 'None of the files could be read, so nothing was saved.';
    } else {
      els.resultTitle.textContent = 'Done';
      els.status.textContent = mode === 'combined'
        ? `Combined ${ok} file${ok === 1 ? '' : 's'} into one ${OUTPUT_FORMATS[format].label} file.`
        : mode === 'separate'
          ? `Converted ${ok} file${ok === 1 ? '' : 's'} to ${OUTPUT_FORMATS[format].label}, bundled as a .zip.`
          : `Converted to ${OUTPUT_FORMATS[format].label}.`;
    }
    if (failed.length) {
      els.failures.innerHTML = `${failed.length} file${failed.length === 1 ? '' : 's'} couldn't be converted:<br>`
        + failed.map((f) => `<strong>${escapeHtml(f.label)}</strong>: ${escapeHtml(f.message)}`).join('<br>');
      els.failures.hidden = false;
    }
  } catch (err) {
    console.error(err);
    els.resultTitle.textContent = 'Something went wrong';
    els.status.textContent = err.message || String(err);
  } finally {
    await terminateOcr();
    state.busy = false;
    document.body.classList.remove('busy');
    renderOptions();
  }
});

render();
