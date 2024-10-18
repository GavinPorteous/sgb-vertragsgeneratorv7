import os
import pandas as pd
from flask import Flask, request, render_template, send_file
from pdfrw import PdfReader, PdfWriter, PdfDict
import zipfile

app = Flask(__name__)

# Directory and file paths
TEMPLATE_DIR = "templates/"
OUTPUT_DIR = "output/"
PDF_TEMPLATE_PATH = os.path.join(TEMPLATE_DIR, "contract_template.pdf")
EXCEL_TEMPLATE_PATH = os.path.join(TEMPLATE_DIR, "data_template.xlsx")

# Helper function to create output directory
def create_output_directory():
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

# Function to read PDF fields
def get_pdf_fields(pdf_path):
    template_pdf = PdfReader(pdf_path)
    fields = {}
    for page in template_pdf.pages:
        annotations = page.Annots
        if annotations:
            for annotation in annotations:
                field_name = annotation.T
                if field_name:
                    fields[field_name[1:-1]] = None  # Remove brackets
    return fields

# Function to create contracts from Excel
def create_contracts_from_excel(excel_path, pdf_template_path):
    create_output_directory()  # Ensure output directory exists
    df = pd.read_excel(excel_path)
    pdf_files = []  # To keep track of generated PDFs

    pdf_fields = get_pdf_fields(pdf_template_path)
    
    for index, row in df.iterrows():
        if row.isnull().all():
            continue  # Skip empty rows

        filled_fields = {pdf_field: str(row[pdf_field]) for pdf_field in pdf_fields if pdf_field in row}
        company_name = row["###company###"]
        output_pdf_name = f"{company_name} Antrag SGB Portfolio.pdf"
        output_pdf_path = os.path.join(OUTPUT_DIR, output_pdf_name)

        # Generate PDF
        generate_pdf(filled_fields, pdf_template_path, output_pdf_path)
        pdf_files.append(output_pdf_name)

    return pdf_files

# Function to generate PDF
def generate_pdf(field_data, template_path, output_pdf_path):
    reader = PdfReader(template_path)
    writer = PdfWriter()

    for page in reader.pages:
        annotations = page.Annots
        if annotations:
            for annotation in annotations:
                field_name = annotation.T
                if field_name:
                    field_name_str = field_name[1:-1]
                    if field_name_str in field_data:
                        annotation.update(PdfDict(V='{}'.format(field_data[field_name_str])))

        writer.addpage(page)

    with open(output_pdf_path, "wb") as output_pdf_file:
        writer.write(output_pdf_file)

# Function to create a zip file
def create_zip(pdf_files):
    zip_filename = os.path.join(OUTPUT_DIR, "contracts.zip")
    with zipfile.ZipFile(zip_filename, 'w') as zip_file:
        for pdf_file in pdf_files:
            zip_file.write(os.path.join(OUTPUT_DIR, pdf_file), pdf_file)
    return zip_filename

# Function to handle file uploads
@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "No file part", 400
    file = request.files['file']
    if file.filename == '':
        return "No selected file", 400

    # Save the uploaded file
    excel_path = os.path.join(TEMPLATE_DIR, file.filename)
    file.save(excel_path)

    # Generate PDFs from the uploaded Excel file
    create_output_directory()
    pdf_files = create_contracts_from_excel(excel_path, PDF_TEMPLATE_PATH)
    zip_filename = create_zip(pdf_files)

    return render_template('success.html', pdf_files=pdf_files, zip_file=zip_filename)

# Route for the index page
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        return upload_file()  # Delegate to the upload_file function
    return render_template('index.html')

@app.route('/download/<filename>', methods=['GET'])
def download_pdf(filename):
    return send_file(os.path.join(OUTPUT_DIR, filename), as_attachment=True)

@app.route('/download_zip', methods=['GET'])
def download_zip():
    return send_file(os.path.join(OUTPUT_DIR, "contracts.zip"), as_attachment=True)

@app.route('/download_template', methods=['GET'])  # New route for the Excel template
def download_template():
    return send_file(EXCEL_TEMPLATE_PATH, as_attachment=True)

@app.route('/hello')
def hello():
    return "Hello, AWS Lambda!"  # This can be removed if unnecessary

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=int(os.environ.get("PORT", 5000)), debug=True)  # Allow external connections

