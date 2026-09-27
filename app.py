import os
import subprocess
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
    # Linux (Cloud) par MS Office ki jagah LibreOffice ka use
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
    
    try:
        # PDF to Word
        if tool_name == 'PDF to WORD':
            output_path = input_path.replace('.pdf', '.docx')
            cv = Converter(input_path)
            cv.convert(output_path)
            cv.close()

        # Word, PPT, Excel to PDF (Handled by LibreOffice on Cloud)
        elif tool_name in ['WORD to PDF', 'POWERPOINT to PDF', 'EXCEL to PDF']:
            convert_to_pdf_linux(input_path, UPLOAD_FOLDER)
            base_name = os.path.splitext(filename)[0]
            output_path = os.path.join(UPLOAD_FOLDER, base_name + '.pdf')

        # PDF to Excel
        elif tool_name == 'PDF to EXCEL':
            output_path = input_path.replace('.pdf', '.xlsx')
            tables = []
            with pdfplumber.open(input_path) as pdf:
                for page in pdf.pages:
                    table = page.extract_table()
                    if table:
                        df = pd.DataFrame(table[1:], columns=table[0])
                        tables.append(df)
            if tables:
                final_df = pd.concat(tables, ignore_index=True)
                final_df.to_excel(output_path, index=False)
            else:
                pd.DataFrame({"Message": ["No tables found"]}).to_excel(output_path, index=False)

        # PDF to PPT
        elif tool_name == 'PDF to POWERPOINT':
            output_path = input_path.replace('.pdf', '.pptx')
            prs = Presentation()
            with pdfplumber.open(input_path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    slide = prs.slides.add_slide(prs.slide_layouts[1])
                    title = slide.shapes.title
                    content = slide.placeholders[1]
                    title.text = "Webo1 Tools - Converted Page"
                    content.text = text if text else "No text"
            prs.save(output_path)

        return send_file(output_path, as_attachment=True)

    except Exception as e:
        print(f"Error: {e}")
        return str(e), 500
        
    finally:
        if os.path.exists(input_path):
            try: os.remove(input_path)
            except: pass

if __name__ == '__main__':
    app.run(port=5000)