from pdf2docx import Converter
from docx2pdf import convert
from docx import Document
from docx.shared import Inches
import os

def pdf_to_word():
    pdf_folder = r"C:\\FIM\\PDF"
    word_folder = r"C:\\FIM\\WORD"
    image_path = r"C:\\FIM\\assGrottone.bmp"
    for pdf_file in os.listdir(pdf_folder):
        if pdf_file.endswith(".pdf"):
            pdf_path = os.path.join(pdf_folder, pdf_file)
            word_file = pdf_file.replace(".pdf", ".docx")
            word_path = os.path.join(word_folder, word_file)
            cv = Converter(pdf_path)
            cv.convert(word_path, start=0, end=None)
            cv.close()

              # Abra o documento do Word e insira a imagem padrão
            doc = Document(word_path)
            doc.add_picture(image_path, width=Inches(2))  
            doc.save(word_path)
            print(f"{pdf_file} convertido para Word com sucesso!")

def word_to_pdf():
    word_folder = "C:\\FIM\\WORD"
    pdf_folder = "C:\\FIM\\SAIDAS"
    for word_file in os.listdir(word_folder):
        if word_file.endswith(".docx"):
            word_path = os.path.join(word_folder, word_file)
            pdf_file = word_file.replace(".docx", ".pdf")
            pdf_path = os.path.join(pdf_folder, pdf_file)
            convert(word_path, pdf_path)
            print(f"{word_file} convertido para PDF com sucesso!")

def main():
    choice = input("Escolha a conversão (1:pdf_to_word or 2:word_to_pdf): ").strip().lower()
    if choice == "1":
        pdf_to_word()
    elif choice == "2":
        word_to_pdf()
    else:
        print("Escolha inválida. Por favor, escolha 'pdf_to_word' or 'word_to_pdf'.")

if __name__ == "__main__":
    main()
