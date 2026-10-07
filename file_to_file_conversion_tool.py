import os
import re
import shutil
import pdfplumber
import pandas as pd
from docx import Document
from pptx import Presentation
from PIL import Image
from pdf2image import convert_from_path
import easyocr  # Swapped pytesseract for easyocr

SUPPORTED_EXTENSIONS = {
    '.pdf', '.docx', '.doc', '.pptx', '.ppt', '.csv', '.xlsx', '.xls',
    '.png', '.jpg', '.jpeg', '.tiff', '.bmp', '.txt',
}
EXT_MAP = {'word': 'docx', 'excel': 'xlsx', 'csv': 'csv', 'txt': 'txt', 'pdf': 'pdf'}

# EasyOCR model loading is slow, so share one reader across every file in a batch
_ocr_reader = None

def get_ocr_reader():
    global _ocr_reader
    if _ocr_reader is None:
        _ocr_reader = easyocr.Reader(['en'], gpu=False)
    return _ocr_reader

# --- USER PROMPTS ---

def ask_choice(prompt, options):
    """Prompts until the user picks one of the given options (matched case-insensitively)."""
    while True:
        answer = input(prompt).strip().lower()
        if answer in options:
            return answer
        print(f"❌ Please enter one of: {', '.join(options)}")

def get_input_files():
    """Prompts for a file or folder path and returns the list of files to convert."""
    while True:
        path = input("Enter the absolute path of the file or folder you want to convert: ").strip().strip('"')
        if os.path.isfile(path):
            _, ext = os.path.splitext(path.lower())
            print(f"🔍 Auto-detected file type: {ext.upper()}")
            return [path], False
        if os.path.isdir(path):
            files = collect_folder_files(path)
            if files:
                return files, True
            print("❌ No supported files were found in that folder.")
            continue
        print("❌ Invalid path. Please ensure the file or folder exists.")

def collect_folder_files(folder):
    """Lists the supported files in a folder, optionally including subfolders."""
    recursive = ask_choice("Include files in subfolders too? (y/n): ", ['y', 'n']) == 'y'
    files = []
    if recursive:
        for root, dirs, names in os.walk(folder):
            dirs.sort()
            files.extend(os.path.join(root, n) for n in sorted(names))
    else:
        files = [os.path.join(folder, n) for n in sorted(os.listdir(folder))
                 if os.path.isfile(os.path.join(folder, n))]

    # Skip hidden files and Office lock files (e.g. "~$report.docx")
    supported, skipped = [], []
    for f in files:
        name = os.path.basename(f)
        if name.startswith('.') or name.startswith('~$'):
            continue
        if os.path.splitext(name.lower())[1] in SUPPORTED_EXTENSIONS:
            supported.append(f)
        else:
            skipped.append(name)

    print(f"\n📂 Found {len(supported)} supported file(s):")
    for f in supported:
        print(f"   • {os.path.relpath(f, folder)}")
    if skipped:
        print(f"⏭️  Skipping {len(skipped)} unsupported file(s): {', '.join(skipped)}")
    return supported

def get_output_format():
    print(f"\nSupported destination formats: {', '.join(EXT_MAP)}")
    return ask_choice("What file type do you want it saved as? ", list(EXT_MAP))

def get_output_dir():
    while True:
        out_dir = input("Enter the destination folder path (where to save the file(s)): ").strip().strip('"')
        if os.path.isdir(out_dir):
            return out_dir
        print("❌ Destination directory does not exist. Please enter a valid folder path.")

def get_filename():
    while True:
        filename = input("Enter a name for your new file (do not include the extension): ").strip()
        filename = re.sub(r'[\\/*?:"<>|]', "", filename)
        if filename:
            return filename
        print("❌ Filename cannot be blank or contain illegal characters.")

# --- EXTRACTORS FOR DIFFERENT INPUT TYPES ---

def extract_from_pdf(pdf_path):
    print("⏳ Processing PDF (Checking for text layers and image layouts)...")
    pages = []

    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text(x_tolerance=3, y_tolerance=3)
            if not text or len(text.strip()) < 5:
                print(f"🖼️ Page {i} appears to be an image. Running EasyOCR...")
                try:
                    images = convert_from_path(pdf_path, first_page=i, last_page=i)
                    if images:
                        reader = get_ocr_reader()

                        # Save image temporarily as EasyOCR works best with file structures or arrays
                        temp_img_path = f"temp_page_{i}.png"
                        images[0].save(temp_img_path)

                        results = reader.readtext(temp_img_path, detail=0)
                        text = "\n".join(results)

                        # Cleanup temp file
                        if os.path.exists(temp_img_path):
                            os.remove(temp_img_path)
                except Exception as ocr_err:
                    print(f"⚠️ OCR failed on page {i}: {ocr_err}. Skipping image scan.")
            pages.append(text if text else "")
    return pages

def extract_from_word(docx_path):
    print("⏳ Processing Word Document...")
    doc = Document(docx_path)
    text = "\n".join([p.text for p in doc.paragraphs])
    for table in doc.tables:
        for row in table.rows:
            row_text = "\t".join([cell.text.strip() for cell in row.cells])
            text += f"\n{row_text}"
    return [text]

def extract_from_powerpoint(pptx_path):
    print("⏳ Processing PowerPoint Presentation Layouts...")
    prs = Presentation(pptx_path)
    slide_texts = []
    for i, slide in enumerate(prs.slides, start=1):
        slide_text = f"--- SLIDE {i} ---\n"
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                slide_text += f"{shape.text}\n"
        slide_texts.append(slide_text)
    return slide_texts

def extract_from_spreadsheet(sheet_path, is_csv=False):
    print("⏳ Processing Spreadsheet Grid Data...")
    df = pd.read_csv(sheet_path) if is_csv else pd.read_excel(sheet_path)
    text_representation = df.to_string(index=False)
    return [text_representation]

def extract_from_image(img_path):
    print("📷 Image detected. Executing EasyOCR...")
    try:
        reader = get_ocr_reader()
        results = reader.readtext(img_path, detail=0)
        text = "\n".join(results)
        return [text]
    except Exception as e:
        print(f"❌ EasyOCR reading failed: {e}")
        return [""]

def extract_text(in_path):
    """Dispatches to the right extractor based on the file extension. Returns a list of page texts."""
    _, in_ext = os.path.splitext(in_path.lower())
    if in_ext == '.pdf':
        return extract_from_pdf(in_path)
    elif in_ext in ['.docx', '.doc']:
        return extract_from_word(in_path)
    elif in_ext in ['.pptx', '.ppt']:
        return extract_from_powerpoint(in_path)
    elif in_ext == '.csv':
        return extract_from_spreadsheet(in_path, is_csv=True)
    elif in_ext in ['.xlsx', '.xls']:
        return extract_from_spreadsheet(in_path, is_csv=False)
    elif in_ext in ['.png', '.jpg', '.jpeg', '.tiff', '.bmp']:
        return extract_from_image(in_path)
    elif in_ext == '.txt':
        with open(in_path, 'r', encoding='utf-8', errors='replace') as f:
            return [f.read()]
    else:
        print(f"⚠️ Unrecognized file type {in_ext}. Attempting generic text read...")
        with open(in_path, 'r', encoding='utf-8', errors='replace') as f:
            return [f.read()]

# --- WRITING OUT THE TARGET FORMAT ---

def text_to_rows(text):
    rows = []
    for line in text.split('\n'):
        if line.strip():
            columns = re.split(r'\s{2,}|\t', line.strip())
            rows.append(columns)
    return rows

def save_docx_fallback(doc, out_path):
    temp_docx = out_path.replace('.pdf', '.docx')
    doc.save(temp_docx)
    print(f"⚠️ Saved layout file to text structural payload: {temp_docx}")

def write_output(fmt, text_pages, out_path, in_path):
    """Writes one converted file (used for single-file and separate-files modes)."""
    combined_text = "\n\n".join(text_pages)

    if fmt == 'txt':
        with open(out_path, 'w', encoding='utf-8', errors='replace') as f:
            f.write(combined_text)

    elif fmt == 'word':
        doc = Document()
        for i, page in enumerate(text_pages, start=1):
            if i > 1:
                doc.add_page_break()
            doc.add_paragraph(page)
        doc.save(out_path)

    elif fmt in ['csv', 'excel']:
        df = pd.DataFrame(text_to_rows(combined_text))
        if fmt == 'csv':
            df.to_csv(out_path, index=False, header=False, encoding='utf-8')
        else:
            df.to_excel(out_path, index=False, header=False)

    elif fmt == 'pdf':
        if in_path.lower().endswith('.pdf'):
            shutil.copyfile(in_path, out_path)
        else:
            doc = Document()
            doc.add_paragraph(combined_text)
            save_docx_fallback(doc, out_path)

def unique_sheet_name(name, used):
    """Excel sheet names max out at 31 chars, can't contain []:*?/\\ and must be unique."""
    base = re.sub(r'[\[\]:*?/\\]', '_', name)[:31] or "Sheet"
    candidate, n = base, 2
    while candidate.lower() in used:
        suffix = f"_{n}"
        candidate = base[:31 - len(suffix)] + suffix
        n += 1
    used.add(candidate.lower())
    return candidate

def write_combined_output(fmt, results, out_path):
    """Writes every file's text into a single output file, labelled by source file name.
    `results` is a list of (label, text_pages) tuples."""

    if fmt == 'txt':
        with open(out_path, 'w', encoding='utf-8', errors='replace') as f:
            for i, (label, pages) in enumerate(results):
                if i > 0:
                    f.write("\n\n")
                f.write(f"===== {label} =====\n\n")
                f.write("\n\n".join(pages))

    elif fmt in ['word', 'pdf']:
        doc = Document()
        for i, (label, pages) in enumerate(results):
            if i > 0:
                doc.add_page_break()
            doc.add_heading(label, level=1)
            for j, page in enumerate(pages):
                if j > 0:
                    doc.add_page_break()
                doc.add_paragraph(page)
        if fmt == 'word':
            doc.save(out_path)
        else:
            save_docx_fallback(doc, out_path)

    elif fmt == 'csv':
        # Prefix every row with the file it came from so the rows stay traceable
        rows = []
        for label, pages in results:
            rows.extend([label] + r for r in text_to_rows("\n\n".join(pages)))
        pd.DataFrame(rows).to_csv(out_path, index=False, header=False, encoding='utf-8')

    elif fmt == 'excel':
        # One worksheet per source file
        used = set()
        with pd.ExcelWriter(out_path) as writer:
            for label, pages in results:
                df = pd.DataFrame(text_to_rows("\n\n".join(pages)))
                sheet = unique_sheet_name(os.path.splitext(os.path.basename(label))[0], used)
                df.to_excel(writer, sheet_name=sheet, index=False, header=False)

def separate_output_path(in_path, out_dir, fmt, used):
    """Names each output after its source file; adds the source extension if two names collide
    (e.g. report.pdf and report.docx -> report.txt and report_docx.txt)."""
    stem, ext = os.path.splitext(os.path.basename(in_path))
    name = stem
    if name.lower() in used:
        name = f"{stem}_{ext.lstrip('.')}"
    n = 2
    while name.lower() in used:
        name = f"{stem}_{ext.lstrip('.')}_{n}"
        n += 1
    used.add(name.lower())
    return os.path.join(out_dir, f"{name}.{EXT_MAP[fmt]}")

# --- MAIN CONTROLLER ---

def convert_single(in_path):
    fmt = get_output_format()
    out_dir = get_output_dir()
    out_path = os.path.join(out_dir, f"{get_filename()}.{EXT_MAP[fmt]}")

    try:
        text_pages = extract_text(in_path)
    except Exception as e:
        print(f"❌ Error reading source file: {e}")
        return

    write_output(fmt, text_pages, out_path, in_path)
    print(f"\n🎉 Success! Process complete. Output file saved to:\n👉 {out_path}")

def convert_folder(files):
    print("\nHow would you like the files exported?")
    print("  1) One combined file containing every file's contents")
    print("  2) Separate files, one per input file")
    mode = ask_choice("Choose 1 or 2: ", ['1', '2'])

    fmt = get_output_format()
    out_dir = get_output_dir()
    if mode == '1':
        out_path = os.path.join(out_dir, f"{get_filename()}.{EXT_MAP[fmt]}")
    else:
        print("ℹ️ Each output file will be named after its source file.")

    # Labels shown in combined output: the path relative to the common input folder
    base = os.path.commonpath([os.path.dirname(f) for f in files])
    results, failed, saved = [], [], []
    used_names = set()

    for idx, in_path in enumerate(files, start=1):
        label = os.path.relpath(in_path, base)
        print(f"\n[{idx}/{len(files)}] {label}")
        try:
            text_pages = extract_text(in_path)
        except Exception as e:
            print(f"❌ Error reading source file: {e}")
            failed.append(label)
            continue

        if mode == '1':
            results.append((label, text_pages))
            continue

        file_out = separate_output_path(in_path, out_dir, fmt, used_names)
        try:
            write_output(fmt, text_pages, file_out, in_path)
            saved.append(file_out)
            print(f"✅ Saved: {file_out}")
        except Exception as e:
            print(f"❌ Error writing output for {label}: {e}")
            failed.append(label)

    if mode == '1':
        if not results:
            print("\n❌ None of the files could be read, so nothing was saved.")
            return
        write_combined_output(fmt, results, out_path)
        print(f"\n🎉 Success! Combined {len(results)} file(s) into:\n👉 {out_path}")
    else:
        print(f"\n🎉 Done! Saved {len(saved)} file(s) to:\n👉 {out_dir}")

    if failed:
        print(f"⚠️ {len(failed)} file(s) could not be converted: {', '.join(failed)}")

def main():
    print("=== Universal Auto-Detecting File Converter ===")
    files, is_folder = get_input_files()
    if is_folder:
        convert_folder(files)
    else:
        convert_single(files[0])

if __name__ == "__main__":
    main()
