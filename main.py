import os
import mimetypes
import pandas as pd
import pdfplumber
from docx import Document

def convert_excel_to_markdown(excel_file_path, output_file_path):
    try:
        df = pd.read_excel(excel_file_path, engine='openpyxl')
        markdown = df.to_markdown(index=False)
        with open(output_file_path, "w", encoding="utf-8") as file:
            file.write(markdown)
    except Exception as e:
        print(f"Error converting {excel_file_path} to Markdown:", e)

def convert_pdf_to_markdown(pdf_file_path, output_file_path):
    try:
        with pdfplumber.open(pdf_file_path) as pdf:
            markdown = ""
            for page in pdf.pages:
                text = page.extract_text()
                markdown += text + "\n\n"
        with open(output_file_path, "w", encoding="utf-8") as file:
            file.write(markdown)
    except Exception as e:
        print(f"Error converting {pdf_file_path} to Markdown:", e)

def convert_word_to_markdown(word_file_path, output_file_path):
    try:
        doc = Document(word_file_path)
        markdown = "\n\n".join([paragraph.text for paragraph in doc.paragraphs])
        with open(output_file_path, "w", encoding="utf-8") as file:
            file.write(markdown)
    except Exception as e:
        print(f"Error converting {word_file_path} to Markdown:", e)

def convert_to_markdown(file_path, output_folder):
    file_name = os.path.basename(file_path)
    output_file_path = os.path.join(output_folder, f"{os.path.splitext(file_name)[0]}.md")
    
    file_type, _ = mimetypes.guess_type(file_path)
    
    if file_type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
        convert_excel_to_markdown(file_path, output_file_path)
    elif file_type == "application/pdf":
        convert_pdf_to_markdown(file_path, output_file_path)
    elif file_type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        convert_word_to_markdown(file_path, output_file_path)

    return output_file_path

# Get the current directory path
current_folder = os.getcwd()
output_folder = os.path.join(current_folder, "output")
os.makedirs(output_folder, exist_ok=True)

# Supported file types
supported_file_types = ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "application/pdf", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"]

# Iterate over all files in the current directory
for file_name in os.listdir(current_folder):
    file_path = os.path.join(current_folder, file_name)
    
    # Skip if it's a directory or a temp file from office
    if os.path.isdir(file_path) or file_name.startswith('~$'):
        continue
    
    # Get the file type
    file_type, _ = mimetypes.guess_type(file_path)
    
    if file_type in supported_file_types:
        convert_to_markdown(file_path, output_folder)
