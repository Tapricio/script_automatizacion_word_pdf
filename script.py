import os
import re
import pandas as pd
from datetime import datetime
from tkinter import Tk, filedialog
from docx import Document
import win32com.client


# Obtener la fecha actual en formato DD-MM-YYYY
fecha_hoy = datetime.today().strftime('%d-%m-%Y')
idTemplate = 10001

# Crear las carpetas necesarias sin sobrescribir
#base_folder_name = rf"G:\Unidades compartidas\Salud\Cartas devolución jubilados\{fecha_hoy} carta devolución jubilados"
base_folder_name = rf".\output\cartas reconocer"
#base_folder = base_folder_name
#counter = 1


# Oculta la ventana principal de Tkinter
root = Tk()
root.withdraw()

# Selección de archivo (descomentar para usar ventana de diálogo)

file_path = filedialog.askopenfilename(
    title="Selecciona un archivo Excel",
    filetypes=[("Excel files", "*.xlsx *.xls")]
)

fechasInvalidas = []
fechasValidas = []
rutInvalidos = []
edadInvalidas = []
casosInvalidos = []
idInvalidos = []

val = 0

def limpiar_nombre_archivo(nombre):
    return re.sub(r'[\/:*?"<>|]', '', nombre)

def docx_replace_regex(doc_obj, regex , replace):

    for p in doc_obj.paragraphs:
        if regex.search(p.text):
            inline = p.runs
            # Loop added to work with runs (strings with same style)
            for i in range(len(inline)):
                if regex.search(inline[i].text):
                    text = regex.sub(replace, inline[i].text)
                    inline[i].text = text

    for table in doc_obj.tables:
        for row in table.rows:
            for cell in row.cells:
                docx_replace_regex(cell, regex , replace)




def docx_replace_multiple_regex(doc_obj, replacements):
    # Compila todos los regex
    compiled_replacements = {re.compile(k): v for k, v in replacements.items()}

    def replace_in_paragraphs(paragraphs):
        for p in paragraphs:
            for regex, replace in compiled_replacements.items():
                if regex.search(p.text):
                    inline = p.runs
                    for i in range(len(inline)):
                        if regex.search(inline[i].text):
                            inline[i].text = regex.sub(replace, inline[i].text)

    def replace_in_tables(tables):
        for table in tables:
            for row in table.rows:
                for cell in row.cells:
                    replace_in_paragraphs(cell.paragraphs)

    replace_in_paragraphs(doc_obj.paragraphs)
    replace_in_tables(doc_obj.tables)



if file_path:
    df = pd.read_excel(file_path, header=1)
    df.iloc[:, 4] = pd.to_datetime(df.iloc[:, 4], errors='coerce')  # Columna fecha nacimiento
    count=0
    for index, row in df.iterrows():
        if pd.notna(df.iloc[index, 2]):  # Validar que hay paciente
            print(df.iloc[index, 2])

