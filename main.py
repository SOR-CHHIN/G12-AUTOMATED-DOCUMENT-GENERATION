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

# Function to convert Arabic numerals to Khmer numerals
def convert_to_khmer_number(arabic_number):
    arabic_to_khmer = {
        '0': '០', '1': '១', '2': '២', '3': '៣', '4': '៤',
        '5': '៥', '6': '៦', '7': '៧', '8': '៨', '9': '៩'
    }
    return ''.join(arabic_to_khmer.get(char, char) for char in str(arabic_number))


# Prepare context for template rendering
def prepare_context(template_keys, row_data, khmer_fields=None):
    if len(row_data) < len(template_keys):
        row_data = row_data + ("",) * (len(template_keys) - len(row_data))
    context = {template_keys[i]: row_data[i] for i in range(len(template_keys))}

    # Convert specific fields to Khmer numerals if specified
    if khmer_fields:
        for field in khmer_fields:
            if field in context and isinstance(context[field], (int, str)):
                context[field] = convert_to_khmer_number(context[field])

    context["cur_date"] = date.today().strftime("%d %b %Y")
    return context

# Render a Word document
def render_document(template_path, context, output_path):
    doc = DocxTemplate(template_path)
    doc.render(context)
    doc.save(output_path)
    print(f"Document saved: {output_path}")

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

    # Fields that require Khmer numeral conversion
    khmer_fields = ["id_kh", "dob_kh", "g1", "g2"]

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
            context = prepare_context(template_keys, row, khmer_fields=khmer_fields)
            print("Generated Context:", context)  # Debugging
            # Use a meaningful name based on context
            first_name = context.get("name_e", "Unknown")
            last_name = context.get("id_e", "Unknown")
            output_name = f"{first_name}.docx"
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
    generate_degrees(excel_file, word_template, output_dir)

def generate_degree_pdf():
    input_dir = "Degrees_Word"
    output_dir = "Degrees_PDF"
    convert_docx_to_pdf(input_dir, output_dir)

def generate_certificate_images():
    excel_file = excel_entry.get()
    output_folder = "Certificates_Images"
    generate_certificates_as_images(excel_file, output_folder)




# Create main window
window = tk.Tk()
window.title("Document Generator")
window.geometry("700x500")
window.config(bg="#f0f0f0")  # Light background for better contrast

# Header
header = tk.Label(window, text="Automated Document Generation", font=("Arial", 20, "bold"), bg="#f0f0f0", fg="#333")
header.pack(pady=10)

# Input Section
input_frame = tk.Frame(window, bg="#f0f0f0")
input_frame.pack(pady=20)

# Excel File Selection
tk.Label(input_frame, text="Data File:", font=("Arial", 10), bg="#f0f0f0").grid(row=0, column=0, sticky="w", padx=10, pady=5)
excel_entry = tk.Entry(input_frame, width=40)
excel_entry.grid(row=0, column=1, padx=10, pady=5)
tk.Button(input_frame, text="Browse", command=select_excel_file, bg="#007BFF", fg="white").grid(row=0, column=2, padx=10, pady=5)

# Word Template Selection
tk.Label(input_frame, text="Templates:", font=("Arial", 10), bg="#f0f0f0").grid(row=1, column=0, sticky="w", padx=10, pady=5)
word_entry = tk.Entry(input_frame, width=40)
word_entry.grid(row=1, column=1, padx=10, pady=5)
tk.Button(input_frame, text="Browse", command=select_word_template, bg="#007BFF", fg="white").grid(row=1, column=2, padx=10, pady=5)

# Button Section
button_frame = tk.Frame(window, bg="#f0f0f0")
button_frame.pack(pady=20)

# Transcript Buttons
tk.Label(button_frame, text="Transcript Options:", font=("Arial", 12, "bold"), bg="#f0f0f0").grid(row=0, columnspan=2, pady=10)
tk.Button(button_frame, text="Generate as Word", command=generate_transcript_word, bg="green", fg="white", width=25).grid(row=1, column=0, padx=10, pady=5)
tk.Button(button_frame, text="Generate as PDF", command=generate_transcript_pdf, bg="green", fg="white", width=25).grid(row=1, column=1, padx=10, pady=5)

# Degree Buttons
tk.Label(button_frame, text="Associate Degree Options:", font=("Arial", 12, "bold"), bg="#f0f0f0").grid(row=2, columnspan=2, pady=10)
tk.Button(button_frame, text="Generate as Word", command=generate_degree_word, bg="orange", fg="white", width=25).grid(row=3, column=0, padx=10, pady=5)
tk.Button(button_frame, text="Generate as PDF", command=generate_degree_pdf, bg="orange", fg="white", width=25).grid(row=3, column=1, padx=10, pady=5)

# Certificate Button
tk.Label(button_frame, text="Certificate Options:", font=("Arial", 12, "bold"), bg="#f0f0f0").grid(row=4, columnspan=2, pady=10)
tk.Button(button_frame, text="Generate as Images", command=generate_certificate_images, bg="teal", fg="white", width=25).grid(row=5, columnspan=2, pady=10)


# Run the application
window.mainloop()
