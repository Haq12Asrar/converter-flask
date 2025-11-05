from flask import Flask, render_template, request, send_file
import os
from werkzeug.utils import secure_filename
import PyPDF2
from PIL import Image
from pdf2docx import Converter
from docx2pdf import convert
from docx import Document
import pandas as pd
from pptx import Presentation
import csv
import io

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # 32MB max file size

# Create uploads folder if it doesn't exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {
    'pdf': ['txt', 'docx'],  # pdf can be converted to txt and docx
    'docx': ['pdf', 'txt'],  # docx can be converted to pdf and txt
    'jpg': ['png', 'pdf', 'webp'],  # jpg can be converted to png, pdf and webp
    'png': ['jpg', 'pdf', 'webp'],  # png can be converted to jpg, pdf and webp
    'jpeg': ['png', 'pdf', 'webp'],  # jpeg can be converted to png, pdf and webp
    'webp': ['jpg', 'png', 'pdf'],  # webp can be converted to jpg, png and pdf
    'xlsx': ['csv', 'pdf', 'txt'],  # excel can be converted to csv, pdf and txt
    'xls': ['xlsx', 'csv', 'pdf', 'txt'],  # old excel can be converted to new excel, csv, pdf and txt
    'csv': ['xlsx', 'txt', 'pdf'],  # csv can be converted to excel, txt and pdf
    'pptx': ['pdf', 'txt'],  # powerpoint can be converted to pdf and txt
    'txt': ['pdf', 'docx'],  # txt can be converted to pdf and docx
}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def get_possible_conversions(filename):
    if '.' not in filename:
        return []
    ext = filename.rsplit('.', 1)[1].lower()
    return ALLOWED_EXTENSIONS.get(ext, [])

@app.route('/')
def index():
    return render_template('index.html')

def check_file_permissions(file):
    try:
        # Check if file size is within limits
        file.seek(0, 2)  # Go to end of file
        size = file.tell()
        file.seek(0)  # Reset file pointer
        
        if size > app.config['MAX_CONTENT_LENGTH']:
            return False, "File size exceeds maximum limit (32MB)"
        
        # Check file extension
        if not allowed_file(file.filename):
            return False, "File type not supported"
        
        return True, "OK"
    except Exception as e:
        return False, str(e)

@app.route('/convert', methods=['POST'])
def convert_file():
    if 'file' not in request.files:
        return {'error': 'No file uploaded'}, 400
    
    file = request.files['file']
    if file.filename == '':
        return {'error': 'No selected file'}, 400

    # Check file permissions and validity
    is_valid, message = check_file_permissions(file)
    if not is_valid:
        return {'error': message}, 400

    target_format = request.form.get('target_format')
    if not target_format:
        return 'No target format specified', 400

    filename = secure_filename(file.filename)
    input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(input_path)

    try:
        output_filename = f"{filename.rsplit('.', 1)[0]}.{target_format}"
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)

        # Perform conversion based on input and target format
        if filename.endswith('.pdf'):
            if target_format == 'txt':
                # Convert PDF to text
                with open(input_path, 'rb') as pdf_file:
                    pdf_reader = PyPDF2.PdfReader(pdf_file)
                    with open(output_path, 'w', encoding='utf-8') as txt_file:
                        for page in pdf_reader.pages:
                            txt_file.write(page.extract_text())
            elif target_format == 'docx':
                # Convert PDF to DOCX
                cv = Converter(input_path)
                cv.convert(output_path)
                cv.close()

        elif filename.endswith('.docx'):
            if target_format == 'pdf':
                # Convert DOCX to PDF
                convert(input_path, output_path)
            elif target_format == 'txt':
                # Convert DOCX to TXT
                doc = Document(input_path)
                with open(output_path, 'w', encoding='utf-8') as txt_file:
                    for para in doc.paragraphs:
                        txt_file.write(para.text + '\n')

        elif filename.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(input_path)
            if target_format == 'csv':
                df.to_csv(output_path, index=False)
            elif target_format == 'txt':
                df.to_string(output_path, index=False)
            elif target_format == 'pdf':
                # Convert to PDF using HTML as intermediate
                html = df.to_html()
                pdf = PyPDF2.PdfWriter()
                pdf.add_page()
                pdf.write(output_path)
            elif target_format == 'xlsx' and filename.endswith('.xls'):
                df.to_excel(output_path, index=False)

        elif filename.endswith('.csv'):
            df = pd.read_csv(input_path)
            if target_format == 'xlsx':
                df.to_excel(output_path, index=False)
            elif target_format == 'txt':
                df.to_string(output_path, index=False)
            elif target_format == 'pdf':
                # Convert to PDF using HTML as intermediate
                html = df.to_html()
                pdf = PyPDF2.PdfWriter()
                pdf.add_page()
                pdf.write(output_path)

        elif filename.endswith('.pptx'):
            prs = Presentation(input_path)
            if target_format == 'txt':
                with open(output_path, 'w', encoding='utf-8') as txt_file:
                    for slide in prs.slides:
                        for shape in slide.shapes:
                            if hasattr(shape, "text"):
                                txt_file.write(shape.text + '\n')
            elif target_format == 'pdf':
                # Save as PDF (Note: This is a basic conversion)
                pdf = PyPDF2.PdfWriter()
                for slide in prs.slides:
                    pdf.add_page()
                pdf.write(output_path)

        elif filename.endswith('.txt'):
            with open(input_path, 'r', encoding='utf-8') as txt_file:
                content = txt_file.read()
                if target_format == 'pdf':
                    pdf = PyPDF2.PdfWriter()
                    pdf.add_page()
                    pdf.write(output_path)
                elif target_format == 'docx':
                    doc = Document()
                    doc.add_paragraph(content)
                    doc.save(output_path)

        elif filename.endswith(('.jpg', '.jpeg', '.png', '.webp')) and target_format in ['jpg', 'png', 'pdf', 'webp']:
            img = Image.open(input_path)
            if target_format == 'pdf':
                # Convert to RGB if necessary
                if img.mode in ('RGBA', 'LA'):
                    bg = Image.new('RGB', img.size, (255, 255, 255))
                    bg.paste(img, mask=img.split()[-1])
                    img = bg
                img.save(output_path, 'PDF', resolution=100.0)
            else:
                img.save(output_path)

        # Clean up input file
        os.remove(input_path)
        
        # Send the converted file
        return send_file(output_path, as_attachment=True, download_name=output_filename)

    except PermissionError:
        return {'error': 'Permission denied while accessing the file'}, 403
    except OSError as e:
        return {'error': f'System error: {str(e)}'}, 500
    except Exception as e:
        return {'error': f'Conversion error: {str(e)}'}, 500
    finally:
        # Clean up files
        try:
            if os.path.exists(input_path):
                os.remove(input_path)
            if os.path.exists(output_path):
                os.remove(output_path)
        except:
            pass  # Ignore cleanup errors

if __name__ == '__main__':
    app.run(debug=True)