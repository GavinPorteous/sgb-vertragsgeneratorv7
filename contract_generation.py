import os
import pandas as pd
from pdfrw import PdfReader, PdfWriter, PdfDict

# Directory and file paths
TEMPLATE_DIR = "templates/"
OUTPUT_DIR = "output/"
PDF_TEMPLATE_PATH = os.path.join(TEMPLATE_DIR, "contract_template.pdf")
EXCEL_TEMPLATE_PATH = os.path.join(TEMPLATE_DIR, "data_template.xlsx")

# Helper function to create output directory
def create_output_directory():
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

# Function to read PDF fields with error handling
def get_pdf_fields(pdf_path):
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"Template PDF not found at {pdf_path}")

    try:
        template_pdf = PdfReader(pdf_path)
    except Exception as e:
        raise pdfrw.errors.PdfParseError(f"Could not read PDF file {pdf_path}: {str(e)}")

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
                    field_name_str = field_name[1:-1]  # Remove brackets
                    if field_name_str in field_data:
                        annotation.update(PdfDict(V='{}'.format(field_data[field_name_str])))

        writer.addpage(page)

    with open(output_pdf_path, "wb") as output_pdf_file:
        writer.write(output_pdf_file)

if __name__ == "__main__":
    create_output_directory()
    create_contracts_from_excel(EXCEL_TEMPLATE_PATH, PDF_TEMPLATE_PATH)


