from flask import Flask, render_template, request, send_file
import os
import tempfile  # Used for safer temporary file handling
from pdf2docx import Converter
from pptx import Presentation
from PIL import Image
import fitz  # PyMuPDF for PDF rendering

app = Flask(__name__)

# --- CONFIGURATION ---
# Use the 'uploads' folder in your project directory
# This matches the folder you created in your project
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
# Set a max file size (e.g., 32MB) to match your frontend note
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024


# ---------- Helper Functions ---------- #

def pdf_to_docx(pdf_path, output_path):
    """Convert PDF -> DOCX with layout retention."""
    try:
        cv = Converter(pdf_path)
        cv.convert(output_path, start=0, end=None)
        cv.close()
        return "Converted successfully"
    except Exception as e:
        return f"Conversion error: {str(e)}"


def pdf_to_ppt(pdf_path, output_pptx):
    """Convert PDF -> PPTX by adding each page as an image."""
    try:
        doc = fitz.open(pdf_path)
        prs = Presentation()
        blank_slide = prs.slide_layouts[6]

        # Use a secure temporary directory for page images
        with tempfile.TemporaryDirectory() as temp_dir:
            for page in doc:
                pix = page.get_pixmap()
                img_path = os.path.join(temp_dir, f"page_{page.number}.png")
                pix.save(img_path)
                
                slide = prs.slides.add_slide(blank_slide)
                # Add the image to the slide
                slide.shapes.add_picture(img_path, 0, 0, prs.slide_width, prs.slide_height)
                
        prs.save(output_pptx)
        return "Converted successfully"
    except Exception as e:
        return f"Conversion error: {str(e)}"

# --- NEW FUNCTION ---
def image_to_pdf(input_path, output_path):
    """Convert JPG, PNG, etc. -> PDF."""
    try:
        # Open the image file
        img = Image.open(input_path)
        # Convert to 'RGB' mode, which is required for saving as PDF
        img_converted = img.convert('RGB')
        # Save the converted image as a PDF
        img_converted.save(output_path)
        return "Converted successfully"
    except Exception as e:
        return f"Conversion error: {str(e)}"

#
# --- NOTE ---
# I have removed your original 'docx_to_pdf' and 'pptx_to_pdf' functions
# because they were not working correctly.
# 'pptx_to_pdf' was creating blank pages.
# 'docx_to_pdf' was only saving plain text, not the real layout.
# It is better to remove a broken feature than to have it fail for users.
#


# ---------- Flask Routes ---------- #

@app.route('/')
def index():
    """Serves the main index.html page."""
    return render_template('index.html')


@app.route('/convert', methods=['POST'])
def convert_file():
    """Handles the file upload and conversion logic."""
    
    # --- 1. GET FILE AND FORMAT ---
    file = request.files.get('file')
    # This is the target format, e.g., "docx"
    target_format = request.form.get('target_format')

    if not file or not target_format:
        return "<h3>Missing file or target format!</h3><a href='/'>Go Back</a>", 400

    # --- 2. GET SOURCE FILE EXTENSION ---
    original_name, original_ext = os.path.splitext(file.filename)
    # This is the source format, e.g., "pdf"
    source_ext = original_ext.lower().replace('.', '')

    # --- 3. SAVE UPLOADED FILE ---
    input_filename = f"{original_name}{original_ext}"
    input_path = os.path.join(app.config['UPLOAD_FOLDER'], input_filename)
    file.save(input_path)

    # --- 4. DEFINE OUTPUT FILE ---
    output_filename = f"{original_name}_converted.{target_format}"
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
    
    msg = ""

    # --- 5. RUN CONVERSION LOGIC ---
    # This is the main fix. We check BOTH the source and target format.
    try:
        if source_ext == 'pdf' and target_format == 'docx':
            msg = pdf_to_docx(input_path, output_path)
        
        elif source_ext == 'pdf' and target_format == 'pptx':
            msg = pdf_to_ppt(input_path, output_path)

        # This case was missing before
        elif source_ext in ['jpg', 'jpeg', 'png'] and target_format == 'pdf':
            msg = image_to_pdf(input_path, output_path)
        
        else:
            # This is the error you were seeing
            msg = f"Unsupported conversion: From {source_ext} to {target_format}"
    
    except Exception as e:
        msg = f"A server error occurred: {str(e)}"

    # --- 6. CLEAN UP INPUT FILE ---
    # Always delete the uploaded file after conversion
    if os.path.exists(input_path):
        os.remove(input_path)

    # --- 7. SEND FILE OR ERROR ---
    if "Converted successfully" in msg and os.path.exists(output_path):
        
        # This is a special function to clean up the *output* file
        # after it has been sent to the user.
        @app.after_request
        def cleanup(response):
            if os.path.exists(output_path):
                os.remove(output_path)
            return response

        # Send the file to the user for download
        return send_file(output_path, as_attachment=True, download_name=output_filename)
    
    else:
        # If conversion failed, delete any broken output file
        if os.path.exists(output_path):
            os.remove(output_path)
        # Send an error message back to the user
        return f"<h3>Conversion Failed: {msg}</h3><a href='/'>Go Back</a>", 500


if __name__ == "__main__":
    # This 'app.run()' is only for testing on your local computer.
    # Render will use a different command (like Gunicorn) to run your app.
    app.run(host='0.0.0.0', port=5000, debug=True)