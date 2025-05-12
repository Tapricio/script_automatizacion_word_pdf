from datetime import datetime, timedelta
import math
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
    return pdf_folder, word_folder
#def desdentado():

def limpiar_nombre_archivo(nombre):
    ret = re.sub(r'[\/:*?"<>|]', '', nombre) 
    return ret.strip() 

          
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
                reemplazar_texto(cell,regex,reemplazo)
    
#revisar párrafos y runs
def revisar_text_doc(doc):
    for p in doc.paragraphs:
        print("Párrafo:", p.text)
        for r in p.runs:
            print("   Run:", r.text)

def rut_formateado (data):
    data.strip().replace(" ","")
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

def convertir_fecha(valor):
    if isinstance(valor, datetime):
        #return valor
        #return datetime.strftime(valor,"%d/%m/%Y")
        return valor.strftime("%d-%m-%Y")
    elif isinstance(valor, (int, float)) and not math.isnan(valor):
        # Convertir desde fecha Excel
        fecha =datetime(1899, 12, 30) + timedelta(days=int(valor))
        return fecha.strftime("%d-%m-%Y")
    elif isinstance(valor, str):
        for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(valor.strip(), fmt).strftime("%d-%m-%Y")
            except ValueError:
                continue
        
    raise ValueError(f"Formato de fecha inválido: {valor}")

def errores_exception(atributo,errores,exception,index,data=None):
    errores.setdefault(atributo,[]).append({
        "fila_excel": index + 3,
        "error": str(exception),
        "data": data
    })
def validar_edad(edad):
    if edad<18 or isinstance(edad,int) or math.isnan(edad):
        raise ValueError("Edad incorrecta")
    return str(int(edad))
    #"fecha_de_nacimiento",errores,e,i

def fecha_actual():
    hoy = datetime.today().strftime('%d-%m-%Y')
    return hoy

def id_template(id):
    try:
        idTemporal=int(id)+1
        return str(idTemporal)
    except ValueError:
        raise ValueError("Error id template")
    