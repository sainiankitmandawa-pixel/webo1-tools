import os
import subprocess
import zipfile
import fitz  # PyMuPDF for PDF to Image conversion
from flask import Flask, request, send_file
from flask_cors import CORS
from pdf2docx import Converter
import pdfplumber
import pandas as pd
from pptx import Presentation

app = Flask(__name__)
# Allow CORS for your live domain
CORS(app, resources={r"/*": {"origins": "*"}})

UPLOAD_FOLDER = os.path.abspath('temp_files')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/', methods=['GET'])
def home():
    return "Webo1 Tools - Backend is 100% Live!"

def convert_to_pdf_linux(input_path, output_dir):
    # Linux (Cloud) par LibreOffice ka use karke convert karein
    subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', input_path, '--outdir', output_dir])

@app.route('/convert', methods=['POST'])
def convert_file():
    if 'file' not in request.files:
        return "No file uploaded", 400
        
    file = request.files['file']
    tool_name = request.form.get('toolName')
    
    filename = file.filename.replace(" ", "_")
    input_path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(input_path)
    
    output_path = None

    try:
        # 1. PDF to WORD
        if tool_name == 'PDF to WORD':
            output_path = input_path.rsplit('.', 1)[0] + '.docx'
            cv = Converter(input_path)
            cv.convert(output_path)
            cv.close()

        # 2. WORD, PPT, EXCEL to PDF
        elif tool_name in ['WORD to PDF', 'POWERPOINT to PDF', 'EXCEL to PDF']:
            convert_to_pdf_linux(input_path, UPLOAD_FOLDER)
            base_name = os.path.splitext(filename)[0]
            output_path = os.path.join(UPLOAD_FOLDER, base_name + '.pdf')

        # 3. PDF to EXCEL (Smart Fallback for Leads & Tables)
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
                        # Agar strict table border na ho, toh text lines extract kar lo taaki leads miss na ho
                        text = page.extract_text()
                        if text:
                            for line in text.split('\n'):
                                all_rows.append([line])
            
            if all_rows:
                df = pd.DataFrame(all_rows)
                df.to_excel(output_path, index=False, header=False)
            else:
                pd.DataFrame({"Message": ["No data or leads found in PDF"]}).to_excel(output_path, index=False)

        # 4. PDF to POWERPOINT
        elif tool_name == 'PDF to POWERPOINT':
            output_path = input_path.rsplit('.', 1)[0] + '.pptx'
            prs = Presentation()
            with pdfplumber.open(input_path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    slide = prs.slides.add_slide(prs.slide_layouts[1])
                    title = slide.shapes.title
                    content = slide.placeholders[1]
                    title.text = "Webo1 Tools - Converted Page"
                    content.text = text if text else "No text found on this page"
            prs.save(output_path)

        # 5. PDF to JPG / IMAGES (Multi-page support with ZIP)
        elif tool_name in ['PDF to JPG', 'PDF to PNG', 'PDF to IMAGE']:
            doc = fitz.open(input_path)
            image_paths = []
            for i, page in enumerate(doc):
                pix = page.get_pixmap(dpi=150)
                img_path = os.path.join(UPLOAD_FOLDER, f"page_{i+1}.jpg")
                pix.save(img_path)
                image_paths.append(img_path)
            
            if len(image_paths) == 1:
                output_path = image_paths[0]
            else:
                output_path = os.path.join(UPLOAD_FOLDER, "converted_images.zip")
                with zipfile.ZipFile(output_path, 'w') as zipf:
                    for img in image_paths:
                        zipf.write(img, os.path.basename(img))

        else:
            return f"Invalid or unsupported tool name: {tool_name}", 400

        if output_path and os.path.exists(output_path):
            return send_file(output_path, as_attachment=True)
        else:
            return "Conversion failed to generate output file", 500

    except Exception as e:
        print(f"Error: {e}")
        return str(e), 500
        
    finally:
        if os.path.exists(input_path):
            try: os.remove(input_path)
            except: pass

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
