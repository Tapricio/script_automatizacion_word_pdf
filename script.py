import os
import re
import pandas as pd
from datetime import datetime
from tkinter import Tk, filedialog
from docx import Document
import win32com.client
import comtypes.client


# Obtener la fecha actual en formato DD-MM-YYYY
fecha_hoy = datetime.today().strftime('%d-%m-%Y')
idTemplate = 10318

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
edadErroneas = []
casosInvalidos = []
idInvalidos = []

id = 10213

def limpiar_nombre_archivo(nombre):
    return re.sub(r'[\\:*?"<>|]', '', nombre)

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

 

datosErroneos=[]

if file_path:
    df = pd.read_excel(file_path, header=1)
    for index, row in df.iterrows():
        if pd.notna(df.iloc[index, 2]):  # Validar que hay paciente
            
            fechaNacimiento = df.iloc[index, 4]
                #nombre
            try:
                nombre = re.sub(r'\s+', ' ', df.iloc[index, 1].strip())
                #print(nombre)
            except Exception:
                datosErroneos.append(f"index: {index}, data - Nombre: {df.iloc[index, 1]} Rut: {df.iloc[index, 2]} Edad: {df.iloc[index,5]} Fecha de nacimiento: {df.iloc[index, 4]} Precio: {df.iloc[index,25]}")
                continue
            
            #rut
            try:
                rut_response = df.iloc[index, 2].replace(" ", "").replace(".", "")
                rut_base = rut_response[:-2]
                verificador = rut_response[-1]
                rut = "{:,}".format(int(rut_base)).replace(",", ".") + "-" + verificador
                rut = rut.upper()
                #print(f"{rut} - {df.iloc[index, 2]}")
            except Exception:
                datosErroneos.append(f"index: {index}, data - Nombre: {df.iloc[index, 1]} Rut: {df.iloc[index, 2]} Edad: {df.iloc[index,5]} Fecha de nacimiento: {df.iloc[index, 4]} Precio: {df.iloc[index,25]}")
                continue

            
            
            print(f"Nombre: {df.iloc[index, 1]} Rut: {df.iloc[index, 2]} ID: {id}")
            id+=1


            #modificamos el word
            doc = Document("Carta Tratamiento dental Reconocer2.docx")
            docx_replace_regex(doc, re.compile(r"NombreTemplate"), nombre)
            docx_replace_regex(doc, re.compile(r"RutTemplate"), rut)
            docx_replace_regex(doc, re.compile(r"IdTemplate"),str(id))

            # guardar en carpeta "word"
            nombre_archivo_seguro = limpiar_nombre_archivo(nombre)
            word_folder=".\output\cartas reconocer\word"
            ruta_guardado = os.path.join(word_folder, f"{nombre_archivo_seguro} {rut}.docx")
            doc.save(ruta_guardado)     

                
       


if datosErroneos:
    print("------------------------")
    print("ERROR:")
    for error in datosErroneos:
        print(error)
    print("------------------------")



