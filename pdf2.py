import os
import win32com.client  # Asegúrate de tener 'pywin32' instalado

# Rutas de entrada y salida
base_folder = r".\output\cartas reconocer"
word_folder = os.path.join(base_folder, "word")
pdf_folder = os.path.join(base_folder, "pdf")

# Crear carpeta PDF si no existe
os.makedirs(pdf_folder, exist_ok=True)

# Crear objeto de Word
word = win32com.client.Dispatch("Word.Application")
word.Visible = False

# Convertir todos los .docx a .pdf
for filename in os.listdir(word_folder):
    if filename.endswith(".docx") and not filename.startswith("~$"):  # Evita temporales
        docx_path = os.path.join(word_folder, filename)
        pdf_name = os.path.splitext(filename)[0] + ".pdf"
        pdf_path = os.path.join(pdf_folder, pdf_name)

        try:
            doc = word.Documents.Open(os.path.abspath(docx_path))
            doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)  # 17 = PDF
            doc.Close()
            print(f"✅ Convertido: {filename}")
        except Exception as e:
            print(f"❌ Error al convertir {filename}: {e}")

# Cerrar Word
word.Quit()

print(f"\nTodos los PDF se guardaron en: {os.path.abspath(pdf_folder)}")
