import os
import re
from tkinter import filedialog

def mkdirFolders (path): 
    count = 1
    while os.path.exists(path):
        path = f"{path}({count})"
        count +=1

    pdf_folder = os.path.join(path, "test pdf")
    word_folder = os.path.join(path, "test word")

    os.makedirs(pdf_folder, exist_ok=True)
    os.makedirs(word_folder, exist_ok=True)

#def desdentado():

def limpiar_nombre_archivo(nombre):
    return re.sub(r'[\/:*?"<>|]', '', nombre)

          
def reemplazar_texto(doc,regex,reemplazo):
    for p in doc.paragraphs:
        if regex.search(p.text):
            runs =p.runs
            for i in range(len(runs)):
                if regex.search(runs[i].text):
                    try:
                        runs[i].text = regex.sub(reemplazo, runs[i].text)
                    except Exception as e:
                        print(e)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                reemplazar_texto(doc,regex,reemplazo)

#revisar párrafos y runs
def revisar_text_doc(doc):
    for p in doc.paragraphs:
        print("Párrafo:", p.text)
        for r in p.runs:
            print("   Run:", r.text)

def rut_formateado (data):
    rut_sin_dv = data[:-2]
    dv = data[-1]
    rut = "{:,}".format(int(rut_sin_dv)).replace(",", ".") + "-" + dv
    return rut.upper() 

def cargar_excel():
    doc = filedialog.askopenfilename(
        title="Selecciona un archivo Excel",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    return doc