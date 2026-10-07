import os
import re
import pdfplumber
import pandas as pd
from docx import Document
from pptx import Presentation
from PIL import Image
from pdf2image import convert_from_path
import easyocr  # Swapped pytesseract for easyocr

def get_input_file():
    """Prompts for a file path, sanitizes it, and auto-detects its extension."""
    while True:
        path = input("Enter the absolute path of the file you want to convert: ").strip().strip('"')
        if os.path.exists(path):
            _, ext = os.path.splitext(path.lower())
            print(f"🔍 Auto-detected file type: {ext.upper()}")
            return path, ext
        print("❌ Invalid path. Please ensure the file exists.")

def get_output_settings():
    valid_formats = ['word', 'excel', 'csv', 'txt', 'pdf']
    print(f"\nSupported destination formats: {', '.join(valid_formats)}")
    
    while True:
        fmt = input("What file type do you want it saved as? ").strip().lower()
        if fmt in valid_formats:
            break
        print("❌ Format not recognized. Choose from: word, excel, csv, txt, pdf")
        
    while True:
        out_dir = input("Enter the destination folder path (where to save the file): ").strip().strip('"')
        if os.path.isdir(out_dir):
            break
        print("❌ Destination directory does not exist. Please enter a valid folder path.")

    while True:
        filename = input("Enter a name for your new file (do not include the extension): ").strip()
        filename = re.sub(r'[\\/*?:"<>|]', "", filename)
        if filename:
            break
        print("❌ Filename cannot be blank or contain illegal characters.")

    ext_map = {'word': 'docx', 'excel': 'xlsx', 'csv': 'csv', 'txt': 'txt', 'pdf': 'pdf'}
    out_path = os.path.join(out_dir, f"{filename}.{ext_map[fmt]}")
    return fmt, out_path

# --- EXTRACTORS FOR DIFFERENT INPUT TYPES ---

def extract_from_pdf(pdf_path):
    print("⏳ Processing PDF (Checking for text layers and image layouts)...")
    pages = []
    # Initialize EasyOCR reader inside the function so it only runs if needed
    reader = None 
    
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text(x_tolerance=3, y_tolerance=3)
            if not text or len(text.strip()) < 5:
                print(f"🖼️ Page {i} appears to be an image. Running EasyOCR...")
                try:
                    images = convert_from_path(pdf_path, first_page=i, last_page=i)
                    if images:
                        if reader is None:
                            reader = easyocr.Reader(['en'], gpu=False)
                        
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
        reader = easyocr.Reader(['en'], gpu=False)
        results = reader.readtext(img_path, detail=0)
        text = "\n".join(results)
        return [text]
    except Exception as e:
        print(f"❌ EasyOCR reading failed: {e}")
        return [""]

# --- MAIN CONTROLLER ---

def main():
    print("=== Universal Auto-Detecting File Converter ===")
    in_path, in_ext = get_input_file()
    fmt, out_path = get_output_settings()
    
    try:
        if in_ext == '.pdf':
            text_pages = extract_from_pdf(in_path)
        elif in_ext in ['.docx', '.doc']:
            text_pages = extract_from_word(in_path)
        elif in_ext in ['.pptx', '.ppt']:
            text_pages = extract_from_powerpoint(in_path)
        elif in_ext == '.csv':
            text_pages = extract_from_spreadsheet(in_path, is_csv=True)
        elif in_ext in ['.xlsx', '.xls']:
            text_pages = extract_from_spreadsheet(in_path, is_csv=False)
        elif in_ext in ['.png', '.jpg', '.jpeg', '.tiff', '.bmp']:
            text_pages = extract_from_image(in_path)
        elif in_ext == '.txt':
            with open(in_path, 'r', encoding='utf-8', errors='replace') as f:
                text_pages = [f.read()]
        else:
            print(f"⚠️ Unrecognized file type {in_ext}. Attempting generic text read...")
            with open(in_path, 'r', encoding='utf-8', errors='replace') as f:
                text_pages = [f.read()]
    except Exception as e:
        print(f"❌ Error reading source file: {e}")
        return

    # --- WRITING OUT THE TARGET FORMAT ---
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
        rows = []
        for line in combined_text.split('\n'):
            if line.strip():
                columns = re.split(r'\s{2,}|\t', line.strip())
                rows.append(columns)
        df = pd.DataFrame(rows)
        if fmt == 'csv':
            df.to_csv(out_path, index=False, header=False, encoding='utf-8')
        else:
            df.to_excel(out_path, index=False, header=False)
            
    elif fmt == 'pdf':
        if in_ext == '.pdf':
            import shutil
            shutil.copyfile(in_path, out_path)
        else:
            doc = Document()
            doc.add_paragraph(combined_text)
            temp_docx = out_path.replace('.pdf', '.docx')
            doc.save(temp_docx)
            print(f"⚠️ Saved layout file to text structural payload: {temp_docx}")

    print(f"\n🎉 Success! Process complete. Output file saved to:\n👉 {out_path}")

if __name__ == "__main__":
    main()
