import os
import zipfile
import fitz  # PyMuPDF
from flask import Flask, request, send_file
from flask_cors import CORS
import pdfplumber
import pandas as pd
from pypdf import PdfReader, PdfWriter
import urllib.request
import json

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})

UPLOAD_FOLDER = os.path.abspath('temp_files')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/', methods=['GET'])
def home():
    return "Webo1 Tools - Stable Backend is 100% Live!"

def simple_translate(text, target_lang='hi'):
    try:
        if not text or not text.strip():
            return ""
        text = text[:1500]
        url = f"https://translate.googleapis.com/translate_a/single?client=gtx&sl=auto&tl={target_lang}&dt=t&q={urllib.parse.quote(text)}"
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as response:
            res = json.loads(response.read().decode('utf-8'))
            return "".join([item[0] for item in res[0] if item[0]])
    except Exception as e:
        print(f"Translation error: {e}")
        return text

@app.route('/convert', methods=['POST'])
def convert_file():
    tool_name = request.form.get('toolName')
    output_path = None

    try:
        if tool_name == 'TRANSLATE PDF':
            if 'file' not in request.files:
                return "No file uploaded", 400
            
            file = request.files['file']
            filename = file.filename.replace(" ", "_")
            input_path = os.path.join(UPLOAD_FOLDER, filename)
            file.save(input_path)
            
            target_lang = request.form.get('targetLanguage', 'hi')
            output_path = os.path.join(UPLOAD_FOLDER, "translated_" + filename)
            
            doc = fitz.open(input_path)
            translated_doc = fitz.open()

            for page in doc:
                # Extract text using both block and standard method for better accuracy
                text = page.get_text("text")
                if not text.strip():
                    # Fallback to pdfplumber extraction if fitz is empty
                    with pdfplumber.open(input_path) as p_pdf:
                        p_page = p_pdf.pages[page.number]
                        text = p_page.extract_text() or ""

                translated_text = simple_translate(text, target_lang) if text.strip() else "[No extractable text found on this page]"
                
                new_page = translated_doc.new_page(width=page.rect.width, height=page.rect.height)
                
                # Draw translated text neatly onto the new PDF page
                rect = fitz.Rect(50, 50, page.rect.width - 50, page.rect.height - 50)
                new_page.insert_textbox(rect, translated_text, fontsize=10, fontname="helv")

            translated_doc.save(output_path)
            translated_doc.close()
            doc.close()

            if os.path.exists(input_path):
                try: os.remove(input_path)
                except: pass

            return send_file(output_path, as_attachment=True)

        elif tool_name == 'MERGE PDF':
            files = request.files.getlist('file')
            if not files or len(files) < 2:
                return "Please select at least 2 PDF files to merge", 400
            
            writer = PdfWriter()
            saved_paths = []
            for file in files:
                if file.filename:
                    fname = file.filename.replace(" ", "_")
                    fpath = os.path.join(UPLOAD_FOLDER, fname)
                    file.save(fpath)
                    saved_paths.append(fpath)
                    reader = PdfReader(fpath)
                    for page in reader.pages:
                        writer.add_page(page)
            
            output_path = os.path.join(UPLOAD_FOLDER, "merged.pdf")
            with open(output_path, "wb") as output_file:
                writer.write(output_file)

            for p in saved_paths:
                try: os.remove(p)
                except: pass

        elif tool_name == 'JPG to PDF':
            files = request.files.getlist('file')
            if not files:
                return "Please select image files", 400
            
            img_paths = []
            for file in files:
                if file.filename:
                    fname = file.filename.replace(" ", "_")
                    fpath = os.path.join(UPLOAD_FOLDER, fname)
                    file.save(fpath)
                    img_paths.append(fpath)
            
            output_path = os.path.join(UPLOAD_FOLDER, "images.pdf")
            doc = fitz.open()
            for img_path in img_paths:
                try:
                    img_doc = fitz.open(img_path)
                    pdfbytes = img_doc.convert_to_pdf()
                    img_pdf = fitz.open("pdf", pdfbytes)
                    doc.insert_pdf(img_pdf)
                    img_doc.close()
                    img_pdf.close()
                except Exception as e:
                    print(f"Skipping image error: {e}")
            doc.save(output_path)
            doc.close()
            for p in img_paths:
                try: os.remove(p)
                except: pass

        else:
            if 'file' not in request.files:
                return "No file uploaded", 400
            
            file = request.files['file']
            filename = file.filename.replace(" ", "_")
            input_path = os.path.join(UPLOAD_FOLDER, filename)
            file.save(input_path)

            if tool_name == 'SPLIT PDF':
                reader = PdfReader(input_path)
                writer = PdfWriter()
                if len(reader.pages) > 0:
                    writer.add_page(reader.pages[0])
                output_path = input_path.rsplit('.', 1)[0] + '_split.pdf'
                with open(output_path, "wb") as output_file:
                    writer.write(output_file)

            elif tool_name == 'COMPRESS PDF':
                output_path = input_path.rsplit('.', 1)[0] + '_compressed.pdf'
                doc = fitz.open(input_path)
                doc.save(output_path, garbage=4, deflate=True, clean=True)
                doc.close()

            elif tool_name == 'PDF to EXCEL':
                output_path = input_path.rsplit('.', 1)[0] + '.xlsx'
                all_rows = []
                with pdfplumber.open(input_path) as pdf:
                    for page in pdf.pages:
                        tables = page.extract_tables()
                        if tables:
                            for table in tables:
                                for row in table:
                                    all_rows.append(row)
                        else:
                            text = page.extract_text()
                            if text:
                                for line in text.split('\n'):
                                    all_rows.append([line])
                
                if all_rows:
                    df = pd.DataFrame(all_rows)
                    df.to_excel(output_path, index=False, header=False)
                else:
                    pd.DataFrame({"Message": ["No data found in PDF"]}).to_excel(output_path, index=False)

            elif tool_name == 'PDF to JPG':
                doc = fitz.open(input_path)
                image_paths = []
                for i, page in enumerate(doc):
                    pix = page.get_pixmap(dpi=120)
                    img_path = os.path.join(UPLOAD_FOLDER, f"page_{i+1}.jpg")
                    pix.save(img_path)
                    image_paths.append(img_path)
                doc.close()
                
                if len(image_paths) == 1:
                    output_path = image_paths[0]
                else:
                    output_path = os.path.join(UPLOAD_FOLDER, "images.zip")
                    with zipfile.ZipFile(output_path, 'w') as zipf:
                        for img in image_paths:
                            zipf.write(img, os.path.basename(img))
                    for img in image_paths:
                        try: os.remove(img)
                        except: pass

            if os.path.exists(input_path):
                try: os.remove(input_path)
                except: pass

        if output_path and os.path.exists(output_path):
            return send_file(output_path, as_attachment=True)
        else:
            return "Conversion failed on server", 500

    except Exception as e:
        print(f"Error: {e}")
        return str(e), 500

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 10000))
    app.run(host='0.0.0.0', port=port)
