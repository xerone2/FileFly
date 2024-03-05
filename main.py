from flask import Flask, render_template, request, send_file
from pdf2docx import Converter
import os
import win32com
from win32com import client

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
CONVERTED_FOLDER = 'converted'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['CONVERTED_FOLDER'] = CONVERTED_FOLDER


def convert_pdf_to_docx(pdf_path, docx_path):
    try:
        cv = Converter(pdf_path)
        cv.convert(docx_path, start=0, end=None)
        cv.close()
        return True, f"Conversion Successful"
    except Exception as e:
        return False, f"Error converting PDF to DOCX: {e}"


def convert_docx_to_pdf(docx_path, pdf_path):
    try:
        word = win32com.client.Dispatch("Word.Application")
        doc = word.Documents.Open(str(docx_path))
        doc.SaveAs(str(pdf_path), FileFormat=17)  # 17 represents PDF format
        doc.Close()
        word.Quit()
        return True, f"Conversion successful: {pdf_path}"
    except Exception as e:
        return False, f"Error converting DOCX to PDF: {e}"


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/convert-to-document', methods=['POST'])
def convert_doc():
    if 'file' not in request.files:
        return 'No file part'

    file = request.files['file']

    if file.filename == '':
        return 'No selected file'

    if file:
        filename = file.filename
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        docx_dir = app.config['CONVERTED_FOLDER']
        docx_path = os.path.join(docx_dir, filename + '.docx')

        # Create the directory if it doesn't exist
        os.makedirs(docx_dir, exist_ok=True)

        file.save(pdf_path)

        convert_pdf_to_docx(pdf_path, docx_path)

        return send_file(docx_path, as_attachment=True)


@app.route('/convert-to-pdf', methods=['POST'])
def convert_pdf():
    if 'file' not in request.files:
        return 'No file part'

    file = request.files['file']

    if file.filename == '':
        return 'No selected file'

    if file:
        filename = file.filename
        docx_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        pdf_dir = app.config['CONVERTED_FOLDER']
        pdf_path = os.path.join(pdf_dir, filename + '.pdf')

        # Create the directory if it doesn't exist
        os.makedirs(pdf_dir, exist_ok=True)

        file.save(pdf_path)

        convert_docx_to_pdf(docx_path, pdf_path)

        return send_file(pdf_path, as_attachment=True)

