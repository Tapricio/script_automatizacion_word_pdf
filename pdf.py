import os
import win32com.client


print("\nConvirtiendo archivos Word a PDF...")

#print("WORD FOLDER ORIGINAL", word_folder)
#word_folder = r"C:\PatricioTorres\script_automatizacion_word_pdf\07-05-2025 carta devolución jubilados (2)\07-05-2025 word"
#print("WORD FOLDER FORMATEADO", word_folder)
#pdf_folder = r"C:\PatricioTorres\script_automatizacion_word_pdf\07-05-2025 carta devolución jubilados (2)\07-05-2025 pdf"

# Crear objeto de Word
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word_folder="./output/cartas reconocer/word"
pdf_folder = "./output/cartas reconocer/pdf"
# Convertir todos los .docx a .pdf
for filename in os.listdir(word_folder):
    if filename.endswith(".docx") and not filename.startswith("~$"):  # Evitar archivos temporales
        #docx_path = os.path.join(word_folder, filename)
        docx_path = os.path.abspath(os.path.join(word_folder, filename))
        pdf_name = os.path.splitext(filename)[0] + ".pdf"
        pdf_path = os.path.abspath(os.path.join(pdf_folder, pdf_name))

        try:
            doc = word.Documents.Open(os.path.abspath(docx_path))
            doc.SaveAs(pdf_path, FileFormat=17)  # 17 = PDF
            doc.Close()
            print(f"Convertido: {filename}")
        except Exception as e:
            print(f"Error al convertir {filename}: {e}")

# Cerrar Word
word.Quit()


print(f"\n✅ Archivos PDF guardados en: {os.path.abspath(pdf_folder)}")





