import io
import pythoncom
from flask import Flask, render_template, request, send_file, make_response
from pdf2docx import Converter
import os
import win32com.client
import getpass

app = Flask(__name__)

# Get the username and Downloads folder path
username = getpass.getuser()
downloads_path = os.path.join("C:\\Users", username, "Downloads")


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
        # Save the uploaded file
        filename = file.filename
        file_path = os.path.join(downloads_path, filename)
        file.save(file_path)

        # Generate the destination PDF file path
        destination_extension = '.docx'
        destination_file_name = filename.replace('.pdf', destination_extension)
        docx_path = os.path.join(downloads_path, destination_file_name)

        def convert_pdf_to_docx(pdf_path, doc_path):
            try:
                cv = Converter(pdf_path)
                cv.convert(doc_path, start=0, end=None)
                cv.close()
                return True, f"Conversion Successful"
            except Exception as e:
                return False, f"Error converting PDF to DOCX: {e}"

        convert_pdf_to_docx(file_path, docx_path)

        return send_file(docx_path, as_attachment=True)


@app.route('/convert-to-pdf', methods=['POST'])
def convert_docx_to_pdf():
    if 'file' not in request.files:
        return 'No file part'

    file = request.files['file']

    if file.filename == '':
        return 'No selected file'

    try:

        # Save the uploaded file
        filename = file.filename
        file_path = os.path.join(downloads_path, filename)
        file.save(file_path)

        # Generate the destination PDF file path
        destination_extension = '.pdf'
        destination_file_name = filename.replace('.docx', destination_extension)
        pdf_path = os.path.join(downloads_path, destination_file_name)

        # Convert DOCX to PDF using Word COM object
        word = win32com.client.Dispatch("Word.Application", pythoncom.CoInitialize())
        doc = word.Documents.Open(file_path)
        pdf_data = io.BytesIO()
        doc.SaveAs(pdf_path, FileFormat=17)  # 17 represents PDF format
        doc.Close()
        word.Quit()

        print(f"Conversion successful: {pdf_path}")
        pdf_data.seek(0)

        response = make_response(pdf_data.getvalue())

        response.headers['Content-Type'] = 'application/pdf'
        response.headers['Content-Disposition'] = f'attachment; filename={destination_file_name}'
        # Return the converted PDF file
        return response
    except Exception as e:
        # Handle any errors that occur during conversion
        print(f"Error converting DOCX to PDF: {e}")
        return "Error converting file"
