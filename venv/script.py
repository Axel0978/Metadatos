import os
from docx import Document
from openpyxl import load_workbook
from PyPDF2 import PdfReader

def extract_docx_metadata(file_path):
    doc = Document(file_path)
    properties = doc.core_properties
    return {
        'title': properties.title,
        'author': properties.author,
        'creation_date': properties.created,
        'modification_date': properties.modified
    }

def extract_xlsx_metadata(file_path):
    wb = load_workbook(filename=file_path)
    props = wb.properties
    return {
        'title': props.title,
        'author': props.creator,
        'creation_date': props.created,
        'modification_date': props.modified
    }

def extract_pdf_metadata(file_path):
    with open(file_path, 'rb') as f:
        pdf = PdfReader(f)
        info = pdf.metadata
        return {
            'title': info.get('/Title', 'Unknown'),
            'author': info.get('/Author', 'Unknown'),
            'creation_date': info.get('/CreationDate', 'Unknown'),
            'modification_date': info.get('/ModDate', 'Unknown')
        }

def extract_metadata(directory):
    for filename in os.listdir(directory):
        if filename.endswith('.docx'):
            metadata = extract_docx_metadata(os.path.join(directory, filename))
            print(f"Metadata for {filename}: {metadata}")
        elif filename.endswith('.xlsx'):
            metadata = extract_xlsx_metadata(os.path.join(directory, filename))
            print(f"Metadata for {filename}: {metadata}")
        elif filename.endswith('.pdf'):
            metadata = extract_pdf_metadata(os.path.join(directory, filename))
            print(f"Metadata for {filename}: {metadata}")

if __name__ == "__main__":
    directory = input("Enter the directory path: ")
    extract_metadata(directory)

