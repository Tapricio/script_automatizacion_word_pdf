import os, config, utils
import pandas as pd
from tkinter import Tk
from docx import Document



# Oculta la ventana principal de Tkinter
root = Tk()
root.withdraw()

# Selección de archivo
archivo_excel = utils.cargar_excel()
#archivo_excel = "Seguimiento Fundación Reconocer (1) (1).xlsx"
#archivo_excel = "test.xlsx"
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
        # Creamos las carpetas necesarias / root -> word y pdf e iniciamos el idTemplate
        try:
            pdf_folder, word_folder = utils.mkdirFolders(config.CARPETA_BASE)
            with open("idTemplate.txt", "r",encoding="utf-8") as f:
                idTemplateInicial = f.read()
            print(f"primer id: {idTemplateInicial}")
            idTemplate= idTemplateInicial    
            idTemplateFinal = ""
        except Exception:
            print(str(Exception))


if columnas_validadas:
    #print("before: ",type(df.loc[0,"fecha_de_nacimiento"]))
    data = df.to_dict(orient="records")
    paciente = config.PACIENTE
    data.pop() #eliminamos fila de total, que es la última
    for i,row in enumerate(data):
        #asignamos caso
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

        #asignamos fecha de nacimiento
        try:
            paciente["fecha_de_nacimiento"] = utils.convertir_fecha(row["fecha_de_nacimiento"])
        except Exception as e:
            utils. errores_exception("fecha_de_nacimiento",errores,e,i,data)
            continue

        #asignamos edad
        try:
            paciente["edad"] = utils.validar_edad(row["edad"])
        except Exception as e:
            utils.errores_exception("edad",errores,e,i)
            continue
        
        #asignamos id template
        try:
            paciente["id_template"]=utils.id_template(idTemplate)
            idTemplateFinal=paciente["id_template"]
            idTemplate = int(idTemplateFinal)
        except Exception as e:
            utils.errores_exception("id_template",errores,e,i,data)
            continue
      
        
        
        print(paciente["n°_paciente"],paciente["caso"],paciente["nombre_paciente"],paciente["rut"],paciente["fecha_de_nacimiento"], paciente["edad"], type(paciente["id_template"] ))

        #modificar word
        
        ruta_word =config.DESDENTADO_SI if paciente["caso"] == "desdentado" else config.DESDENTADO_NO
        word = Document(ruta_word)
        
        utils.reemplazar_texto(word,config.PALABRAS_WORD_MODIFICAR["Nombre_paciente_template"],paciente["nombre_paciente"])
        utils.reemplazar_texto(word,config.PALABRAS_WORD_MODIFICAR["fecha_emicion_template"],utils.fecha_actual())
        utils.reemplazar_texto(word,config.PALABRAS_WORD_MODIFICAR["fecha_emicion_template"],utils.fecha_actual())
        utils.reemplazar_texto(word,config.PALABRAS_WORD_MODIFICAR["rut_template"],paciente["rut"])
        utils.reemplazar_texto(word,config.PALABRAS_WORD_MODIFICAR["edad_template"],paciente["edad"])
        utils.reemplazar_texto(word,config.PALABRAS_WORD_MODIFICAR["fecha_nacimiento_template"],paciente["fecha_de_nacimiento"])
        utils.reemplazar_texto(word,config.PALABRAS_WORD_MODIFICAR["id_template"],paciente["id_template"])
        



        nombre_archivo_nuevo = f"{utils.limpiar_nombre_archivo(paciente["nombre_paciente"])} {paciente["rut"]}.docx"
        ruta_guardado_archivo_nuevo = os.path.join(word_folder,nombre_archivo_nuevo)
        word.save(ruta_guardado_archivo_nuevo)
       
    
    if errores: 
        print("--------")
        print("Errores")
        for atributo, e in errores.items():
            for error in e:
                print(f"Atributo: {atributo}- Fila excel: {error["fila_excel"]} - Error: {error["error"]}")
    with open("idTemplate.txt","w",encoding="utf-8") as f:
        f.write(idTemplateFinal)
    print("id final: ", idTemplateFinal)
    print(f"\nArchivos Word generados en: {os.path.abspath(word_folder)}")


