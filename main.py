import math
import os, re, win32com.client, config, utils
import pandas as pd
from datetime import datetime, timedelta
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
errores = config.ERRORES

if archivo_excel :
    df = pd.read_excel(archivo_excel, header=1)
    #validamos las cabeceras del Excel, para ver que tenga los nombres correctos.
    try:
        df.columns = df.columns.str.strip().str.lower().str.replace(r'\s+', '_', regex=True)
        validacion_columnas = [col for col in df.columns if col not in config.PACIENTE]
    except Exception as e:
        print("Error al transformar columnas ", e)
    if validacion_columnas:
        print("columna erronea: ",validacion_columnas)
    else:
        print("Columnas validas!")
        columnas_validadas = True


if columnas_validadas:
    #print("before: ",type(df.loc[0,"fecha_de_nacimiento"]))
    data = df.to_dict(orient="records")
    paciente = config.PACIENTE
    data.pop() #eliminamos fila de total, que es la última
    for i,row in enumerate(data):
        """ #asignamos caso
        try:
            paciente["caso"] = row["caso"].strip().lower().replace(" ","")
        except Exception as e:
            utils. errores_exception("caso",errores,e,i,row)
            continue

        #asignamos nombre de paciente
        try:
            paciente["nombre_paciente"] = row["nombre_paciente"].strip()
        except Exception as e:
            utils. errores_exception("nombre_paciente",errores,e,i,row)
            continue
        
        #asignamos rut
        try:
            paciente["rut"] = utils.rut_formateado(row["rut"])
        except Exception as e:
            utils. errores_exception("rut",errores,e,i,data)
            continue
 """
        #asignamos fecha de nacimiento
        try:
            paciente["fecha_de_nacimiento"] = utils.convertir_fecha(row["fecha_de_nacimiento"])
        except Exception as e:
            utils. errores_exception("fecha_de_nacimiento",errores,e,i)
            continue
        print(paciente["fecha_de_nacimiento"], row["nombre_paciente"])
        #paciente["nombre_paciente"] = row["nombre_paciente"].strip()
        """ "n°_paciente": 0,
            "caso": "",
            "nombre_paciente": "",
            "rut":"",
            "fecha_de_nacimiento":"",
            "edad":0,
            "id": 0, """
        
        """ paciente["fecha_de_nacimiento"]=utils.convertir_fecha(paciente["fecha_de_nacimiento"])
        if pd.notna:
           correctas.append([paciente["nombre_paciente"], paciente["fecha_de_nacimiento"]])
        else:
            incorrectas.append([paciente["nombre_paciente"],paciente["fecha_de_nacimiento"]])
            continue """
       # print(paciente["fecha_de_nacimiento"], paciente["rut"])
    #print(errores["rut"])
    print(errores["fecha_de_nacimiento"])

"""  
    

if not columnas_validadas:
   
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