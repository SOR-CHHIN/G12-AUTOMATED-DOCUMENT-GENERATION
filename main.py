import pandas as pd
import os
from PIL import Image, ImageDraw, ImageFont

# Function to create the output folder if it doesn't exist
def create_output_folder(output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        print(f"Output folder '{output_folder}' created.")
    else:
        print(f"Output folder '{output_folder}' already exists.")

# Function to generate a certificate for a single student
def generate_certificate_for_student(name, template_file, output_folder, font_name):
    certificate = Image.open(template_file)
    draw = ImageDraw.Draw(certificate)

    # Calculate text bounding box (width and height)
    bbox = draw.textbbox((0, 0), name, font=font_name)
    text_width = bbox[2] - bbox[0]  # Width of the text
    text_height = bbox[3] - bbox[1]  # Height of the text
    
    # Center the text horizontally and set the Y position
    certificate_width, certificate_height = certificate.size
    name_position = ((certificate_width - text_width) // 2, 620)

    # Draw the name on the certificate
    draw.text(name_position, name, fill="orange", font=font_name)

    # Output path using concatenation
    output_path = os.path.join(output_folder, f"template_{name}.png")

    # Save the certificate
    certificate.save(output_path)

    print(f"Certificate generated for {name} and saved to {output_path}")

# Function to generate certificates for all students
def generate_certificates(excel_file, template_file, output_folder):
    # Load the Excel file
    data = pd.read_excel(excel_file)

    # Create the output folder if it doesn't exist
    create_output_folder(output_folder)

    # Font settings
    bold_font = "arialbd.ttf"
    font_name = ImageFont.truetype(bold_font, 90)

    # Generate certificates for each student
    for index, row in data.iterrows():
        name = row["student_name"]
        generate_certificate_for_student(name, template_file, output_folder, font_name)

    print("All certificates have been generated!")

# Example of how to call the functions
excel_file = "Certificate_Data.xlsx"
template_file = "template.png"
output_folder = "Generate_certificate"

generate_certificates(excel_file, template_file, output_folder)

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
