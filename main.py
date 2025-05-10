import os, re, win32com.client, config, utils
import pandas as pd
from datetime import datetime
from tkinter import Tk, filedialog
from docx import Document

# Creamos las carpetas necesarias / root -> word y pdf
#utils.mkdirFolders(config.CARPETA_BASE)

# Oculta la ventana principal de Tkinter
root = Tk()
root.withdraw()

# Selección de archivo
#archivo_excel = utils.cargar_excel()
archivo_excel = "Seguimiento Fundación Reconocer (1) (1).xlsx"
columnas_validadas = False

if archivo_excel :
    df = pd.read_excel(archivo_excel, header=1)
    #validamos las cabeceras del Excel, para ver que tenga los nombres correctos.
    try:
        df.columns = df.columns.str.strip().str.lower().str.replace(r'\s+', '_', regex=True)
    except Exception as e:
        print("Error al transformar columnas ", e)
    validacion_columnas = [col for col in df.columns if col not in config.COLUMNAS_VALIDAS]
    if validacion_columnas:
        print("columna erronea: ",validacion_columnas)
    else:
        print("Columnas validas!")
        columnas_validadas = True



if columnas_validadas:
    print("before: ",type(df.loc[0,"fecha_de_nacimiento"]))
    df["fecha_de_nacimiento"] = pd.to_datetime(df["fecha_de_nacimiento"], errors='coerce')
    #df["fecha_de_nacimiento"] = df["fecha_de_nacimiento"].dt.strftime('%d-%m-%Y')       
    fechastest=df["fecha_de_nacimiento"].isna()
    if fechastest.any():
        print("fechas invalidas")
        print(df[fechastest])


    data = df.to_dict(orient="records")
    print("after: ",type(df.loc[0,"fecha_de_nacimiento"]))
    
    for paciente in data:
        print("test data: ",paciente["fecha_de_nacimiento"])


if not columnas_validadas:
    df.loc[:,"fecha_de_nacimiento"] = pd.to_datetime(df.loc[:, "fecha_de_nacimiento"], errors='coerce')  # Columna fecha nacimiento
   
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
            #fechasValidas.append(index)       
                   

            
                    
                    

                    
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