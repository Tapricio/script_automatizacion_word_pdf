import os, re, win32com.client, config, utils
import pandas as pd
from datetime import datetime
from tkinter import Tk, filedialog
from docx import Document

# Creamos las carpetas necesarias / root -> word y pdf
utils.mkdirFolders(config.CARPETA_BASE)

# Oculta la ventana principal de Tkinter
root = Tk()
root.withdraw()

# Selección de archivo
archivo_seleccionado = filedialog.askopenfilename(
    title="Selecciona un archivo Excel",
    filetypes=[("Excel files", "*.xlsx *.xls")]
)

# Información de iteraciones
errores = {
"fecha_nacimiento" : [],
"rut" : [],
"edad" : [],
"caso" : [],
"id" : [],
"nombre": []
}

paciente = {
"nombre": "",
"rut":"",
"fecha_nacimiento":"",
"edad":0,
"id": 0,
"desdentado":""
}

fechasValidas = []



if archivo_seleccionado:
    df = pd.read_excel(archivo_seleccionado, header=1)
    df.iloc[:, 4] = pd.to_datetime(df.iloc[:, 4], errors='coerce')  # Columna fecha nacimiento

    for index, row in df.iterrows():
        if pd.notna(df.iloc[index, 0]):  # Validar que haya número de paciente
            #Fecha nacimiento paciente
            try:
                if pd.notna(df.iloc[index, 4]):
                    paciente["fecha_nacimiento"] = df.iloc[index, 4].strftime('%d-%m-%Y')       
            except Exception as e:
                errores["fecha_nacimiento"].append({"fila":index,"error":str(e),"tipo": type(e).__name__})
                continue
            #Nombre paciente
            try:
                if pd.notna(df.iloc[index, 2]):
                    paciente["nombre"] = re.sub(r'\s+', ' ', df.iloc[index, 2].strip())
            except Exception as e:
                errores["nombre"].append({"fila":index,"error":str(e),"tipo": type(e).__name__})
                continue
            
            #Rut paciente
            try:
                if pd.notna(df.iloc[index, 3]):
                    rut_response = df.iloc[index, 3].replace(" ", "").replace(".", "")
                    paciente["rut"] = utils.rut_formateado(rut_response)
            except Exception as e:
                errores["rut"].append({"fila":index,"error":str(e),"tipo": type(e).__name__}) 
                continue
            #Edad paciente
            try:
                if pd.notna(df.iloc[index, 5]):
                    paciente["edad"] = int(float(df.iloc[index, 5])).strip()
            except Exception as e:
                errores["edad"].append({"fila":index,"error":str(e),"tipo": type(e).__name__}) 
                continue       
            #Desdentado paciente
            try:
                if pd.notna(df.iloc[index,1]):
                    paciente["desdentado"] = df.iloc[index,1].lower().strip().replace(" ","")
                
                doc = Document(config.DESDENTADO_SI if paciente["desdentado"].lower().strip() == config.VALIDACION_DESDENTADO else config.DESDENTADO_NO)    
                
            except Exception as e:
                errores["desdentado"].append({"fila":index,"error":str(e),"tipo": type(e).__name__}) 
                continue        
                   

            print(f"Paciente: {paciente['nombre']}, Rut: {paciente['rut']}, Fecha de Nacimiento: {paciente['fecha_nacimiento']}, Edad: {str(paciente['edad'])}, ID: {str(paciente['id'])}")
            fechasValidas.append(index)
                    
                    

                    
""" 
                     # Abrir plantilla y reemplazar
                    #doc = Document("plantilla.docx")
                    docx_replace_regex(doc, re.compile(r"FechaDeEmicionTemplate"), config.TODAY_D_M_Y)
                    docx_replace_regex(doc, re.compile(r"NombreTemplate"), nombre)
                    docx_replace_regex(doc, re.compile(r"RutTemplate"), rut)
                    docx_replace_regex(doc, re.compile(r"EdadTemplate"), str(edad))
                    docx_replace_regex(doc, re.compile(r"FechaDeNacimientoTemplate"), fechaFormateada)
                    docx_replace_regex(doc, re.compile(r"IdTemplate"),str(idTemplate))
                    

                    try:
                        idTemplate+=1
                    except Exception:
                        idInvalidos.append(index)
                        continue


                    # Guardar en carpeta "word"
                    nombre_archivo_seguro = limpiar_nombre_archivo(nombre)
                    ruta_guardado = os.path.join(word_folder, f"{nombre_archivo_seguro} {rut}.docx")
                    doc.save(ruta_guardado)
                    
                    
                   
                    

                except Exception:
                    rutInvalidos.append(index)
                    continue
            else:
                fechasInvalidas.append(index)

    print("\n--- Resumen ---")
    print("Fechas inválidas:", fechasInvalidas)
    print("RUT inválidos:", rutInvalidos)
    print("Edades inválidas:", edadInvalidas)
    print("Fechas válidas:", fechasValidas)
    print("Total pacientes válidos:", val)
    print(f"\nArchivos Word generados en: {os.path.abspath(word_folder)}")
  """