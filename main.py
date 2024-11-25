import openpyxl
from docxtpl import DocxTemplate
import os
from docx2pdf import convert
from datetime import date

# Load data from Excel
def load_excel_data(filename):
    workbook = openpyxl.load_workbook(filename)
    sheet = workbook.active
    return list(sheet.values)


# Prepare context for template rendering
def prepare_context(template_keys, row_data):
    if len(row_data) < len(template_keys):
        row_data = row_data + ("",) * (len(template_keys) - len(row_data))
        # Add the current date to the context with the key "cur_date"
    context = {template_keys[i]: row_data[i] for i in range(len(template_keys))}
    context["cur_date"] = date.today().strftime("%d-%m-%y")  # Use desired date format (e.g., YYYY-MM-DD)
    return context


# Render a Word document
def render_document(template_path, context, output_path):
    doc = DocxTemplate(template_path)
    doc.render(context)
    doc.save(output_path)
    print(f"Document saved: {output_path}")


# Generate Word documents
def generate_documents(excel_file, word_template, output_dir):
    # Define the template variable keys
    template_keys = [
        "student_id", "first_name", "last_name", "logic", "l_g", "bcum", "bc_g", "design", 
        "d_g", "p1", "p1_g", "e1", "e1_g", "wd", "wd_g", "algo", "al_g", "p2", "p2_g", "e2", 
        "e2_g", "sd", "sd_g", "js", "js_g", "php", "ph_g", "db", "db_g", "vc1", "v1_g", "node", 
        "no_g", "e3", "e3_g", "p3", "p3_g", "oop", "op_g", "lar", "lar_g", "vue", "vu_g", "vc2", 
        "v2_g", "e4", "e4_g", "p4", "p4_g", "int", "in_g"
    ]

    # Load data from the Excel file
    data = load_excel_data(excel_file)

    # Skip the header row and process each student
    for row in data[1:]:
        # Prepare context for the current row
        context = prepare_context(template_keys, row)

        # Generate output file name
        output_name = f"{context['first_name']}_{context['last_name']}.docx"
        output_path = os.path.join(output_dir, output_name)

        # Render and save the Word document
        render_document(word_template, context, output_path)


# Convert Word documents to PDF in a new folder
def convert_docx_to_pdf(input_dir, output_dir):
    # Create output directory for PDFs if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)

    # Loop through all .docx files in the input directory
    for file in os.listdir(input_dir):
        if file.endswith(".docx"):
            docx_path = os.path.join(input_dir, file)
            pdf_path = os.path.join(output_dir, file.replace(".docx", ".pdf"))
            # Convert each .docx file to a PDF
            convert(docx_path, pdf_path)
            print(f"Converted to PDF: {pdf_path}")


# Main function
def main():
    # Input and output file paths
    excel_file = "data_student.xlsx"
    word_template = "template-pnc.docx"
    output_dir = "Academic_Transcripts"
    pdf_output_dir = "Academic_Transcript_PDF"

    # Create output directory for Word documents if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)

    # Generate Word documents
    generate_documents(excel_file, word_template, output_dir)

    # Convert generated Word documents to PDF in a separate folder
    convert_docx_to_pdf(output_dir, pdf_output_dir)


if __name__ == "__main__":
    main()