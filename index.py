from pdf2docx import Converter
from docx2pdf import convert
from docx import Document
from docx.shared import Inches
import os

def pdf_to_word():
    pdf_folder = r"C:\\FIM\\PDF"
    word_folder = r"C:\\FIM\\WORD"
    image_path = r"C:\\FIM\\assGrottone.bmp"
    
    
    if not os.path.exists(word_folder):
        os.makedirs(word_folder)
    
    for pdf_file in os.listdir(pdf_folder):
        if pdf_file.endswith(".pdf"):
            pdf_path = os.path.join(pdf_folder, pdf_file)
            word_file = pdf_file.replace(".pdf", ".docx")
            word_path = os.path.join(word_folder, word_file)
            
            try:
                # Verifica se o arquivo PDF existe e não está vazio
                if os.path.getsize(pdf_path) > 0:
                    cv = Converter(pdf_path)
                    cv.convert(word_path, start=0, end=None)
                    cv.close()

               
                    doc = Document(word_path)
                    doc.add_picture(image_path, width=Inches(2))  
                    doc.save(word_path)
                    
                    print(f"{pdf_file} convertido para Word com sucesso!")
                else:
                    print(f"O arquivo PDF {pdf_file} está vazio ou não pode ser lido.")
            except Exception as e:
                print(f"Erro ao processar o arquivo {pdf_file}: {e}")

def word_to_pdf():
    word_folder = "C:\\FIM\\WORD"
    pdf_folder = "C:\\FIM\\SAIDAS"
    
  
    if not os.path.exists(pdf_folder):
        os.makedirs(pdf_folder)
    
    for word_file in os.listdir(word_folder):
        if word_file.endswith(".docx"):
            word_path = os.path.join(word_folder, word_file)
            pdf_file = word_file.replace(".docx", ".pdf")
            pdf_path = os.path.join(pdf_folder, pdf_file)
            
            try:
                convert(word_path, pdf_path)
                print(f"{word_file} convertido para PDF com sucesso!")
            except Exception as e:
                print(f"Erro ao processar o arquivo {word_file}: {e}")

def main():
    choice = input("Escolha a conversão (1: pdf_to_word or 2: word_to_pdf): ").strip().lower()
    if choice == "1":
        pdf_to_word()
    elif choice == "2":
        word_to_pdf()
    else:
        print("Escolha inválida. Por favor, escolha 'pdf_to_word' or 'word_to_pdf'.")

if __name__ == "__main__":
    main()
