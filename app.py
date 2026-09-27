import os
import subprocess
import zipfile
import fitz  # PyMuPDF
from flask import Flask, request, send_file
from flask_cors import CORS
import pdfplumber
import pandas as pd
from pypdf import PdfMerger, PdfReader, PdfWriter

try:
    from pdf2docx import Converter
except:
    Converter = None

try:
    from pptx import Presentation
except:
    Presentation = None

try:
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
except:
    SimpleDocTemplate = None

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})

UPLOAD_FOLDER = os.path.abspath('temp_files')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/', methods=['GET'])
def home():
    return "Webo1 Tools - Stable Backend is 100% Live!"

def convert_to_pdf_linux(input_path, output_dir):
    subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', input_path, '--outdir', output_dir])

@app.route('/convert', methods=['POST'])
def convert_file():
    tool_name = request.form.get('toolName')
    output_path = None

    try:
        if tool_name == 'MERGE PDF':
            files = request.files.getlist('file')
            if not files or len(files) < 2:
                return "Please select at least 2 PDF files to merge", 400
            
            merger = PdfMerger()
            saved_paths = []
            for file in files:
                if file.filename:
                    fname = file.filename.replace(" ", "_")
                    fpath = os.path.join(UPLOAD_FOLDER, fname)
                    file.save(fpath)
                    saved_paths.append(fpath)
                    merger.append(fpath)
            
            output_path = os.path.join(UPLOAD_FOLDER, "webo1_merged.pdf")
            merger.write(output_path)
            merger.close()
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
            
            output_path = os.path.join(UPLOAD_FOLDER, "webo1_images.pdf")
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

            elif tool_name == 'PDF to WORD':
                if not Converter:
                    return "pdf2docx is not available on server", 500
                output_path = input_path.rsplit('.', 1)[0] + '.docx'
                cv = Converter(input_path)
                cv.convert(output_path)
                cv.close()

            elif tool_name in ['WORD to PDF', 'POWERPOINT to PDF', 'HTML to PDF']:
                convert_to_pdf_linux(input_path, UPLOAD_FOLDER)
                base_name = os.path.splitext(filename)[0]
                output_path = os.path.join(UPLOAD_FOLDER, base_name + '.pdf')

            elif tool_name == 'EXCEL to PDF':
                if not SimpleDocTemplate:
                    return "ReportLab is not available on server", 500
                output_path = input_path.rsplit('.', 1)[0] + '.pdf'
                if input_path.endswith('.csv'):
                    df = pd.read_csv(input_path)
                else:
                    df = pd.read_excel(input_path)
                
                df = df.fillna("")
                data = [df.columns.tolist()] + df.values.tolist()

                doc = SimpleDocTemplate(output_path, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=40, bottomMargin=40)
                elements = []
                styles = getSampleStyleSheet()
                
                title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=16, textColor=colors.HexColor('#e53e3e'), spaceAfter=6, alignment=1)
                footer_style = ParagraphStyle('FooterStyle', parent=styles['Normal'], fontSize=9, textColor=colors.HexColor('#718096'), alignment=1)

                elements.append(Paragraph("<b>Webo1 Data Report</b>", title_style))
                elements.append(Paragraph("Powered by Webo1 (webo1.com)", ParagraphStyle('Sub', parent=styles['Normal'], fontSize=10, textColor=colors.HexColor('#2b6cb0'), alignment=1, spaceAfter=15)))

                table_data = [[Paragraph(str(cell), styles['Normal']) for cell in row] for row in data]
                t = Table(table_data, repeatRows=1)
                t.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#e53e3e')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 10),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#f7fafc')),
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#cbd5e0')),
                ]))
                elements.append(t)
                elements.append(Spacer(1, 20))
                elements.append(Paragraph("© Webo1 Enterprise | webo1.com", footer_style))
                doc.build(elements)

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

            elif tool_name == 'PDF to POWERPOINT':
                if not Presentation:
                    return "python-pptx is not available on server", 500
                output_path = input_path.rsplit('.', 1)[0] + '.pptx'
                prs = Presentation()
                with pdfplumber.open(input_path) as pdf:
                    for page in pdf.pages:
                        text = page.extract_text()
                        slide = prs.slides.add_slide(prs.slide_layouts[1])
                        slide.shapes.title.text = "Webo1 Converted Page"
                        slide.placeholders[1].text = text if text else "No text found"
                prs.save(output_path)

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
                    output_path = os.path.join(UPLOAD_FOLDER, "webo1_images.zip")
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
