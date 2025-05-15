
from flask import Flask, render_template, request, send_file, redirect, url_for, flash
import os
import fitz  # PyMuPDF
from pdf2docx import Converter
from PIL import Image
import tempfile
from werkzeug.utils import secure_filename

app = Flask(__name__, static_folder='../static', template_folder='../templates')
app.secret_key = 'your_secret_key'
UPLOAD_FOLDER = '../uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert_pdf_to_word', methods=['POST'])
def convert_pdf_to_word():
    if 'pdf_file' not in request.files:
        flash('No PDF file uploaded.')
        return redirect(url_for('index'))

    pdf_file = request.files['pdf_file']
    if pdf_file.filename == '':
        flash('No selected PDF file.')
        return redirect(url_for('index'))

    filename = secure_filename(pdf_file.filename)
    input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    pdf_file.save(input_path)

    try:
        temp_fd, repaired_path = tempfile.mkstemp(suffix=".pdf")
        os.close(temp_fd)
        doc = fitz.open(input_path)
        doc.save(repaired_path)
        doc.close()

        output_path = input_path.replace(".pdf", "_converted.docx")
        cv = Converter(repaired_path)
        cv.convert(output_path, start=0, end=None)
        cv.close()

        os.remove(repaired_path)
        return send_file(output_path, as_attachment=True)
    except Exception as e:
        flash(f"Conversion failed: {e}")
        return redirect(url_for('index'))

@app.route('/convert_jpg_to_pdf', methods=['POST'])
def convert_jpg_to_pdf():
    files = request.files.getlist('jpg_files')
    if not files:
        flash('No JPG files uploaded.')
        return redirect(url_for('index'))

    try:
        image_paths = []
        for file in files:
            filename = secure_filename(file.filename)
            path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(path)
            image_paths.append(path)

        if len(image_paths) % 2 != 0:
            image_paths.append(None)

        a4_width, a4_height = 595, 842
        img_width, img_height = 249, 331
        top_margin = 50
        space_between = 30
        left_margin = (a4_width - img_width) // 2

        images = []
        for i in range(0, len(image_paths), 2):
            img1 = Image.open(image_paths[i]).convert("RGB").resize((img_width, img_height))
            if image_paths[i + 1]:
                img2 = Image.open(image_paths[i + 1]).convert("RGB").resize((img_width, img_height))
            else:
                img2 = Image.new("RGB", (img_width, img_height), "white")

            page = Image.new("RGB", (a4_width, a4_height), "white")
            page.paste(img1, (left_margin, top_margin))
            page.paste(img2, (left_margin, top_margin + img_height + space_between))
            images.append(page)

        output_pdf = os.path.join(app.config['UPLOAD_FOLDER'], "converted_output.pdf")
        images[0].save(output_pdf, save_all=True, append_images=images[1:], resolution=100.0)

        return send_file(output_pdf, as_attachment=True)

    except Exception as e:
        flash(f"JPG to PDF failed: {e}")
        return redirect(url_for('index'))

handler = app
