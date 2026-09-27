import os
import subprocess
import zipfile
import fitz  # PyMuPDF
from flask import Flask, request, send_file
from flask_cors import CORS
from pdf2docx import Converter
import pdfplumber
import pandas as pd
from pptx import Presentation
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})

UPLOAD_FOLDER = os.path.abspath('temp_files')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/', methods=['GET'])
def home():
    return "Webo1 Tools - Backend is 100% Live & Updated!"

def convert_to_pdf_linux(input_path, output_dir):
    subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', input_path, '--outdir', output_dir])

@app.route('/convert', methods=['POST'])
def convert_file():
    if 'file' not in request.files and 'files' not in request.files:
        return "No file uploaded", 400
        
    tool_name = request.form.get('toolName')
    output_path = None

    try:
        # Multi-file handling for JPG to PDF
        if tool_name in ['JPG to PDF', 'IMAGE to PDF', 'PNG to PDF']:
            files = request.files.getlist('file') or request.files.getlist('files')
            if not files:
                return "No images uploaded", 400
            
            img_paths = []
            for file in files:
                if file.filename:
                    fname = file.filename.replace(" ", "_")
                    fpath = os.path.join(UPLOAD_FOLDER, fname)
                    file.save(fpath)
                    img_paths.append(fpath)
            
            output_path = os.path.join(UPLOAD_FOLDER, "converted_images.pdf")
            
            # Convert multiple images into a single clean PDF using fitz (PyMuPDF)
            doc = fitz.open()
            for img_path in img_paths:
                img_doc = fitz.open(img_path)
                pdfbytes = img_doc.convert_to_pdf()
                img_pdf = fitz.open("pdf", pdfbytes)
                doc.insert_pdf(img_pdf)
            doc.save(output_path)
            doc.close()

        else:
            file = request.files['file']
            filename = file.filename.replace(" ", "_")
            input_path = os.path.join(UPLOAD_FOLDER, filename)
            file.save(input_path)

            # 1. PDF to WORD
            if tool_name == 'PDF to WORD':
                output_path = input_path.rsplit('.', 1)[0] + '.docx'
                cv = Converter(input_path)
                cv.convert(output_path)
                cv.close()

            # 2. WORD, PPT to PDF (LibreOffice)
            elif tool_name in ['WORD to PDF', 'POWERPOINT to PDF']:
                convert_to_pdf_linux(input_path, UPLOAD_FOLDER)
                base_name = os.path.splitext(filename)[0]
                output_path = os.path.join(UPLOAD_FOLDER, base_name + '.pdf')

            # 3. EXCEL to PDF (Professional Table Layout + Webo1 Branding)
            elif tool_name == 'EXCEL to PDF':
                output_path = input_path.rsplit('.', 1)[0] + '.pdf'
                
                # Read Excel or CSV
                if input_path.endswith('.csv'):
                    df = pd.read_csv(input_path)
                else:
                    df = pd.read_excel(input_path)
                
                # Clean columns and data
                df = df.fillna("")
                data = [df.columns.tolist()] + df.values.tolist()

                # Generate professional PDF using ReportLab
                doc = SimpleDocTemplate(output_path, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=40, bottomMargin=40)
                elements = []
                
                styles = getSampleStyleSheet()
                title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=16, textColor=colors.HexColor('#1a365d'), spaceAfter=10, alignment=1)
                footer_style = ParagraphStyle('FooterStyle', parent=styles['Normal'], fontSize=9, textColor=colors.HexColor('#718096'), alignment=1)

                # Title & Webo1 Branding Header
                elements.append(Paragraph("<b>Kanha Computers & Webo1 Data Report</b>", title_style))
                elements.append(Paragraph("Powered by Webo1 (webo1.com) — Digital Academy & IT Services, Jaipur", ParagraphStyle('Sub', parent=styles['Normal'], fontSize=10, textColor=colors.HexColor('#2b6cb0'), alignment=1, spaceAfter=15)))

                # Format table for clean readability
                table_data = []
                for row in data:
                    table_data.append([Paragraph(str(cell), styles['Normal']) for cell in row])

                t = Table(table_data, repeatRows=1)
                t.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2b6cb0')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 10),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#f7fafc')),
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#cbd5e0')),
                    ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#edf2f7')]),
                ]))
                
                elements.append(t)
                elements.append(Spacer(1, 20))
                elements.append(Paragraph("© Webo1 Enterprise | Bani Park, Jaipur | webo1.com", footer_style))
                
                doc.build(elements)

            # 4. PDF to EXCEL
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
                    pd.DataFrame({"Message": ["No data or leads found in PDF"]}).to_excel(output_path, index=False)

            # 5. PDF to POWERPOINT
            elif tool_name == 'PDF to POWERPOINT':
                output_path = input_path.rsplit('.', 1)[0] + '.pptx'
                prs = Presentation()
                with pdfplumber.open(input_path) as pdf:
                    for page in pdf.pages:
                        text = page.extract_text()
                        slide = prs.slides.add_slide(prs.slide_layouts[1])
                        slide.shapes.title.text = "Kanha Computers - Converted Page"
                        slide.placeholders[1].text = text if text else "No text found"
                prs.save(output_path)

            # 6. PDF to JPG / IMAGES (Multi-page ZIP support)
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
                    output_path = os.path.join(UPLOAD_FOLDER, "webo1_converted_images.zip")
                    with zipfile.ZipFile(output_path, 'w') as zipf:
                        for img in image_paths:
                            zipf.write(img, os.path.basename(img))

            if os.path.exists(input_path):
                try: os.remove(input_path)
                except: pass

        if output_path and os.path.exists(output_path):
            return send_file(output_path, as_attachment=True)
        else:
            return "Conversion failed", 500

    except Exception as e:
        print(f"Error: {e}")
        return str(e), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
