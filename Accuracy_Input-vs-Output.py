import os
import fitz  # PyMuPDF library
import docx
from difflib import SequenceMatcher

def extract_pdf_text(pdf_file):
    pdf_document = fitz.open(pdf_file)
    text = ""
    for page_num in range(pdf_document.page_count):
        page = pdf_document[page_num]
        text += page.get_text()
    pdf_document.close()
    return text

def extract_docx_text(docx_file):
    doc = docx.Document(docx_file)
    text = ""
    for para in doc.paragraphs:
        text += para.text + "\n"
    return text

def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()

def compare_accuracy(input_folder, output_folder):
    total_accuracy_pdf = 0.0
    total_accuracy_docx = 0.0
    total_files_pdf = 0
    total_files_docx = 0

    for filename in os.listdir(input_folder):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(input_folder, filename)
            docx_file = filename.replace(".pdf", ".docx")  # Corresponding DOCX file
            docx_path = os.path.join(output_folder, docx_file)

            if os.path.exists(docx_path):
                pdf_text = extract_pdf_text(pdf_path)
                docx_text = extract_docx_text(docx_path)

                accuracy_pdf = similar(pdf_text, docx_text) * 100
                total_accuracy_pdf += accuracy_pdf
                total_files_pdf += 1

                print(f"PDF Accuracy for {filename}: {accuracy_pdf:.2f}%")

        elif filename.endswith(".docx"):
            input_docx_path = os.path.join(input_folder, filename)
            output_docx_path = os.path.join(output_folder, filename)

            if os.path.exists(output_docx_path):
                input_docx_text = extract_docx_text(input_docx_path)
                output_docx_text = extract_docx_text(output_docx_path)

                accuracy_docx = similar(input_docx_text, output_docx_text) * 100
                total_accuracy_docx += accuracy_docx
                total_files_docx += 1

                print(f"DOCX Accuracy for {filename}: {accuracy_docx:.2f}%")

    if total_files_pdf > 0:
        average_accuracy_pdf = total_accuracy_pdf / total_files_pdf
        print(f"Average PDF accuracy: {average_accuracy_pdf:.2f}%")

    if total_files_docx > 0:
        average_accuracy_docx = total_accuracy_docx / total_files_docx
        print(f"Average DOCX accuracy: {average_accuracy_docx:.2f}%")

    overall_total_files = total_files_pdf + total_files_docx
    if overall_total_files > 0:
        overall_average_accuracy = (total_accuracy_pdf + total_accuracy_docx) / overall_total_files
        print(f"Overall Average accuracy: {overall_average_accuracy:.2f}%")

def main():
    input_folder = "C://Users//DELL//Downloads//input_resume_folder"
    output_folder = "C://Users//DELL//Downloads//test_resume//new_docx_files"

    compare_accuracy(input_folder, output_folder)

if __name__ == "__main__":
    main()
 