import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl
from docxtpl import DocxTemplate
import os
from docx2pdf import convert
from datetime import date
import pandas as pd
from PIL import Image, ImageDraw, ImageFont


def load_excel_data(filename):
    workbook = openpyxl.load_workbook(filename)
    sheet = workbook.active
    data = list(sheet.values)
    print("Loaded Excel Data:", data)  # Debugging
    return data


# Prepare context for template rendering
def prepare_context(template_keys, row_data):
    if len(row_data) < len(template_keys):
        row_data = row_data + ("",) * (len(template_keys) - len(row_data))
    context = {template_keys[i]: row_data[i] for i in range(len(template_keys))}
    context["cur_date"] = date.today().strftime("%d-%m-%Y")
    return context

# Render a Word document
def render_document(template_path, context, output_path):
    doc = DocxTemplate(template_path)
    doc.render(context)
    doc.save(output_path)
    print(f"Document saved: {output_path}")

# Generate Word documents for transcripts
def generate_transcripts(excel_file, word_template, output_dir):
    template_keys = [
        "student_id", "first_name", "last_name", "logic", "l_g", "bcum", "bc_g", "design", 
        "d_g", "p1", "p1_g", "e1", "e1_g", "wd", "wd_g", "algo", "al_g", "p2", "p2_g", "e2", 
        "e2_g", "sd", "sd_g", "js", "js_g", "php", "ph_g", "db", "db_g", "vc1", "v1_g", "node", 
        "no_g", "e3", "e3_g", "p3", "p3_g", "oop", "op_g", "lar", "lar_g", "vue", "vu_g", "vc2", 
        "v2_g", "e4", "e4_g", "p4", "p4_g", "int", "in_g"
    ]
    data = load_excel_data(excel_file)
    os.makedirs(output_dir, exist_ok=True)
    for row in data[1:]:
        context = prepare_context(template_keys, row)
        output_name = f"{context['first_name']}_{context['last_name']}.docx"
        output_path = os.path.join(output_dir, output_name)
        render_document(word_template, context, output_path)
    messagebox.showinfo("Success", "Word transcripts generated successfully!")


def generate_degrees(excel_file, word_template, output_dir):
    # Define the keys that map to the associate degree template
    template_keys = [
        "name_kh", "name_e", "g1", "g2", "id_kh", "id_e", "dob_kh", "dob_e", "pro_kh", "pro_e", "ed_kh", "ed_e"
    ]

    # Load Excel data
    try:
        data = load_excel_data(excel_file)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load Excel file: {e}")
        return

    # Create output directory
    os.makedirs(output_dir, exist_ok=True)

    for row in data[1:]:
        try:
            context = prepare_context(template_keys, row)
            print("Generated Context:", context)  # Debugging
            # Use a meaningful name based on context
            first_name = context.get("name_e", "Unknown")
            last_name = context.get("id_e", "Unknown")
            output_name = f"{first_name}_{last_name}_degree.docx"
            output_path = os.path.join(output_dir, output_name)
            render_document(word_template, context, output_path)
        except Exception as e:
            print(f"Error processing row {row}: {e}")
    messagebox.showinfo("Success", "Associate Degree generated successfully!")


def convert_docx_to_pdf(input_dir, output_dir):
    # Ensure the input directory exists
    if not os.path.exists(input_dir):
        print(f"Input directory does not exist: {input_dir}")
        return  # or create the directory
    os.makedirs(output_dir, exist_ok=True)
    for file in os.listdir(input_dir):
        if file.endswith(".docx"):
            docx_path = os.path.join(input_dir, file)
            pdf_path = os.path.join(output_dir, file.replace(".docx", ".pdf"))
            convert(docx_path, pdf_path)
            print(f"Converted to PDF: {pdf_path}")
    messagebox.showinfo("Success", "Generated as PDFs successfully!")

# Generate certificates as images
def generate_certificates_as_images(excel_file, output_folder):
    # Ask user to select an image template
    template_file = filedialog.askopenfilename(
        title="Select Certificate Template",
        filetypes=[("Image Files", "*.png"), ("JPEG Files", "*.jpg")]
    )

    if not template_file:
        messagebox.showwarning("No Template Selected", "Please select an image template!")
        return

    # Process the certificates
    data = pd.read_excel(excel_file)
    os.makedirs(output_folder, exist_ok=True)
    bold_font = "arialbd.ttf"
    font_name = ImageFont.truetype(bold_font, 90)
    for index, row in data.iterrows():
        name = row["student_name"]
        certificate = Image.open(template_file)
        draw = ImageDraw.Draw(certificate)
        bbox = draw.textbbox((0, 0), name, font=font_name)
        text_width = bbox[2] - bbox[0]
        certificate_width, certificate_height = certificate.size
        name_position = ((certificate_width - text_width) // 2, 620)
        draw.text(name_position, name, fill="orange", font=font_name)
        output_path = os.path.join(output_folder, f"certificate_{name}.png")
        certificate.save(output_path)
        print(f"Certificate generated for {name} and saved to {output_path}")
    messagebox.showinfo("Success", "Certificates generated successfully!")

# Tkinter GUI
def main():
    def select_excel_file():
        filepath = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel Files", "*.xlsx")])
        if filepath:
            excel_entry.delete(0, tk.END)
            excel_entry.insert(0, filepath)

    def select_word_template():
        filepath = filedialog.askopenfilename(title="Select Word Template", filetypes=[("Word Files", "*.docx")])
        if filepath:
            word_entry.delete(0, tk.END)
            word_entry.insert(0, filepath)

    def generate_transcript_word():
        excel_file = excel_entry.get()
        word_template = word_entry.get()
        output_dir = "Transcripts_Word"
        generate_transcripts(excel_file, word_template, output_dir)

    def generate_transcript_pdf():
        input_dir = "Transcripts_Word"
        output_dir = "Transcripts_PDF"
        convert_docx_to_pdf(input_dir, output_dir)

    def generate_degree_word():
        excel_file = excel_entry.get()
        word_template = word_entry.get()
        output_dir = "Degrees_Word"
        os.makedirs(output_dir, exist_ok=True)
        generate_degrees(excel_file, word_template, output_dir)

    def generate_degree_pdf():
        input_dir = "Degrees_Word"
        output_dir = "Degrees_PDF"
        convert_docx_to_pdf(input_dir, output_dir)

    def generate_certificate_images():
        excel_file = excel_entry.get()
        output_folder = "Certificates_Images"
        generate_certificates_as_images(excel_file, output_folder)

    window = tk.Tk()
    window.title("Document Generator")
    window.geometry("600x500")

    # Welcome paragraph
    welcome_label = tk.Label(
        window,
        text="Welcome to Automated Documents Generation",
        font=("Arial", 18),  # Adjust the font and size as needed
        fg="black"  # Set the text color to black
    )
    welcome_label.pack(pady=20)  # Add some spacing around the label

    # Input fields
    tk.Label(window, text="Excel File:").pack(pady=5)
    excel_entry = tk.Entry(window, width=50)
    excel_entry.pack(pady=5)
    tk.Button(window, text="Browse Excel File", command=select_excel_file).pack(pady=5)

    tk.Label(window, text="Word Template:").pack(pady=5)
    word_entry = tk.Entry(window, width=50)
    word_entry.pack(pady=5)
    tk.Button(window, text="Browse Word Template", command=select_word_template).pack(pady=5)

    # Buttons
    tk.Button(window, text="Generate Transcript as Word", command=generate_transcript_word, bg="green", fg="white").pack(pady=5)
    tk.Button(window, text="Generate Transcript as PDF", command=generate_transcript_pdf, bg="green", fg="white").pack(pady=5)
    tk.Button(window, text="Generate Associate Degree as Word", command=generate_degree_word, bg="orange", fg="white").pack(pady=5)
    tk.Button(window, text="Generate Associate Degree as PDF", command=generate_degree_pdf, bg="orange", fg="white").pack(pady=5)
    tk.Button(window, text="Generate Certificate as Images", command=generate_certificate_images, bg="teal", fg="white").pack(pady=5)

    window.mainloop()


if __name__ == "__main__":
    main()