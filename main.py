import openpyxl
from docxtpl import DocxTemplate
from docx2pdf import convert
import os
from datetime import date

# Function to load Excel data
def load_excel_data(excel_file):
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook.active
    return list(sheet.values)

# Function to set up directories for output files
def setup_output_directories():
    word_dir = os.path.join("Associate_Degree_word_files")
    pdf_dir = os.path.join("Associate_Degree_pdf_files")
    os.makedirs(word_dir, exist_ok=True)
    os.makedirs(pdf_dir, exist_ok=True)
    return word_dir, pdf_dir

# Function to prepare the context for the template (mapping Excel data to the template)
def prepare_context(template_keys, row_data):
    if len(row_data) < len(template_keys):
        row_data = row_data + ("",) * (len(template_keys) - len(row_data))
    context = {template_keys[i]: row_data[i] for i in range(len(template_keys))}
    context["cur_date"] = date.today().strftime("%Y-%m-%d")

    # Convert fields to Khmer numerals (example: id_kh)
    if "id_kh" in context:
        context["id_kh"] = convert_to_khmer_number(context["id_kh"])
    
    return context

# Function to generate the Word document from the template
def generate_word_certificate(template, data, word_dir):
    safe_name = str(data["name_e"]).strip().replace("/", "_").replace("\\", "_")
    word_path = os.path.join(word_dir, f"{safe_name}.docx")
    template.render(data)
    template.save(word_path)
    print(f"Created Word document: {word_path}")
    return word_path

# Function to convert the Word document to PDF
def convert_to_pdf(word_path, pdf_dir):
    safe_name = os.path.splitext(os.path.basename(word_path))[0]
    pdf_path = os.path.join(pdf_dir, f"{safe_name}.pdf")
    convert(word_path, pdf_path)
    print(f"Created PDF document: {pdf_path}")

# Function to convert Arabic numbers to Khmer numbers
def convert_to_khmer_number(number):
    arabic_to_khmer = {
        '0': '០', '1': '១', '2': '២', '3': '៣', '4': '៤',
        '5': '៥', '6': '៦', '7': '៧', '8': '៨', '9': '៩'
    }
    
    khmer_number = ''.join(arabic_to_khmer.get(char, char) for char in str(number))
    return khmer_number

# Function to generate certificates for all rows in the Excel file
def generate_certificates(excel_file, template_file):
    if not os.path.exists(excel_file):
        print(f"Excel file not found: {excel_file}")
        return

    if not os.path.exists(template_file):
        print(f"Template file not found: {template_file}")
        return

    try:
        data = load_excel_data(excel_file)
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        return

    word_dir, pdf_dir = setup_output_directories()
    template = DocxTemplate(template_file)
    template_keys = ["name_kh", "name_e", "g1", "g2", "id_kh", "id_e", "dob_kh", "dob_e", "pro_kh", "pro_e", "ed_kh", "ed_e"]

    for row in data[1:]:
        try:
            if len(row) < 2 or not row[1]:  # Skip rows with missing names
                print(f"Skipping invalid or incomplete row: {row}")
                continue

            # Prepare the context and convert necessary fields
            template_data = prepare_context(template_keys, row)

            # Generate Word document and convert to PDF
            word_path = generate_word_certificate(template, template_data, word_dir)
            convert_to_pdf(word_path, pdf_dir)
        except Exception as e:
            print(f"Error processing row {row}: {e}")
            continue

    print("All certificates have been successfully generated!")

# Usage
generate_certificates("Data.xlsx", "Certificate_degree.docx")