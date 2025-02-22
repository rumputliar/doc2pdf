import os
import win32com.client as win32
from docx2pdf import convert
from pathlib import Path

def create_pdf_folder():
    pdf_folder = Path('pdf')
    if not pdf_folder.exists():
        pdf_folder.mkdir()
    return pdf_folder

def convert_doc_to_pdf(doc_path, pdf_folder):
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    doc = word.Documents.Open(doc_path)
    pdf_path = pdf_folder / (Path(doc_path).stem + '.pdf')
    doc.SaveAs(str(pdf_path), FileFormat=17)  # FileFormat=17 is for PDF
    doc.Close()
    word.Quit()
    return pdf_path

def convert_docx_to_pdf(docx_path, pdf_folder):
    pdf_path = pdf_folder / (Path(docx_path).stem + '.pdf')
    convert(docx_path, str(pdf_path))
    return pdf_path

def main():
    pdf_folder = create_pdf_folder()
    for root, _, files in os.walk('.'):
        for file in files:
            if file.endswith('.doc'):
                doc_path = os.path.join(root, file)
                pdf_path = convert_doc_to_pdf(doc_path, pdf_folder)
                print(f"Converted {doc_path} to {pdf_path}")
            elif file.endswith('.docx'):
                docx_path = os.path.join(root, file)
                pdf_path = convert_docx_to_pdf(docx_path, pdf_folder)
                print(f"Converted {docx_path} to {pdf_path}")

if __name__ == "__main__":
    main()