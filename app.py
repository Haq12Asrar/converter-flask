from flask import Flask, render_template, request, send_file
import os
import pypandoc
from pdf2docx import Converter
from pptx import Presentation
from fpdf import FPDF
from PIL import Image
import fitz  # PyMuPDF for PDF handling

# 🧩 Auto-install Pandoc on Render (since it's not preinstalled)
if not os.path.exists("/usr/bin/pandoc"):
    os.system("apt-get update && apt-get install -y pandoc")

app = Flask(__name__)

# Use /tmp folder for Render (writable directory)
UPLOAD_FOLDER = "/tmp"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# ---------- Helper Functions ---------- #

def pdf_to_docx(pdf_path, output_path):
    """Convert PDF → DOCX using pdf2docx with fallback to pypandoc."""
    try:
        cv = Converter(pdf_path)
        cv.convert(output_path, start=0, end=None)
        cv.close()
        return "Converted successfully (pdf2docx)"
    except Exception as e:
        try:
            pypandoc.convert_file(pdf_path, 'docx', outputfile=output_path)
            return "Converted successfully (pypandoc fallback)"
        except Exception as ex:
            return f"Conversion error: {str(ex)}"

def docx_to_pdf(input_path, output_path):
    """Convert DOCX → PDF using Pandoc (works on Linux)."""
    try:
        pypandoc.convert_file(input_path, 'pdf', outputfile=output_path)
        return "Converted successfully"
    except Exception as e:
        return f"Conversion error: {str(e)}"

def pptx_to_pdf(pptx_path, output_pdf):
    """Convert PPTX → PDF by creating a blank image per slide."""
    try:
        prs = Presentation(pptx_path)
        pdf = FPDF()
        for i, slide in enumerate(prs.slides):
            img_path = f"/tmp/slide_{i}.png"
            img = Image.new("RGB", (1280, 720), "white")
            img.save(img_path)
            pdf.add_page()
            pdf.image(img_path, 0, 0, 210, 148)  # Fit to A4 width
        pdf.output(output_pdf)
        return "Converted successfully"
    except Exception as e:
        return f"Conversion error: {str(e)}"

def pdf_to_ppt(pdf_path, output_pptx):
    """Convert PDF → PPTX by creating slides from PDF pages."""
    try:
        doc = fitz.open(pdf_path)
        prs = Presentation()
        blank_layout = prs.slide_layouts[6]  # blank slide
        for page in doc:
            pix = page.get_pixmap()
            img_path = f"/tmp/page_{page.number}.png"
            pix.save(img_path)
            slide = prs.slides.add_slide(blank_layout)
            slide.shapes.add_picture(img_path, 0, 0, prs.slide_width, prs.slide_height)
        prs.save(output_pptx)
        return "Converted successfully"
    except Exception as e:
        return f"Conversion error: {str(e)}"

# ---------- Flask Routes ---------- #

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert_file():
    file = request.files.get('file')
    convert_type = request.form.get('target_format')

    if not file:
        return "No file uploaded!"

    input_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(input_path)

    name, _ = os.path.splitext(file.filename)
    output_path = os.path.join(UPLOAD_FOLDER, f"{name}_converted")

    if convert_type == 'pdf_to_docx':
        output_path += ".docx"
        msg = pdf_to_docx(input_path, output_path)

    elif convert_type == 'docx_to_pdf':
        output_path += ".pdf"
        msg = docx_to_pdf(input_path, output_path)

    elif convert_type == 'pptx_to_pdf':
        output_path += ".pdf"
        msg = pptx_to_pdf(input_path, output_path)

    elif convert_type == 'pdf_to_pptx':
        output_path += ".pptx"
        msg = pdf_to_ppt(input_path, output_path)

    else:
        msg = "Unsupported conversion type"

    if "Converted successfully" in msg:
        return send_file(output_path, as_attachment=True)
    else:
        return f"<h3>{msg}</h3><a href='/'>Go Back</a>"

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000)
