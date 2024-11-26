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

