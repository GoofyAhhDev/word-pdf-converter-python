from docx2pdf import convert
import os

def convert_docx_to_pdf(docx_path, pdf_path):
    if not os.path.isfile(docx_path):
        print(f"The file {docx_path} does not exist.")
        return
    
    try:
        convert(docx_path, pdf_path)
        print(f"Converted {docx_path} to {pdf_path}")
    except Exception as e:
        print(f"An error occurred: {e}")

def main():
    try:
        docx_path = input("Enter the path to the DOCX file (e.g., C:\\path\\to\\file.docx): ")
        pdf_path = input("Enter the path to save the PDF file (including filename, e.g., C:\\path\\to\\save\\file.pdf): ")
        convert_docx_to_pdf(docx_path, pdf_path)
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        input("Press Enter to exit...")

if __name__ == "__main__":
    main()
