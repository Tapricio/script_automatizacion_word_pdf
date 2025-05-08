import re
import pandas as pd
from tkinter import Tk, filedialog

pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.max_colwidth', None)

# Oculta la ventana principal de Tkinter
root = Tk()
root.withdraw()

# Abre el cuadro de diálogo para seleccionar archivo
""" file_path = filedialog.askopenfilename(
    title="Selecciona un archivo Excel",
    filetypes=[("Excel files", "*.xlsx *.xls")]
) """
file_path = 'Seguimiento Fundación Reconocer (1).xlsx'
fechasInvalidas = []
fechasValidas = []
rutInvalidos = []
edadInvalidas = []

val=0


if file_path:
    # Lee el Excel con pandas
    df = pd.read_excel(file_path, header=1)
    df.iloc[:, 4] = pd.to_datetime(df.iloc[:, 4], errors='coerce')  # Convierte la columna de fechas

    for index, row in df.iterrows():
        if pd.notna(df.iloc[index, 0]):  # Validamos que haya número de paciente
            fechaNacimiento = df.iloc[index, 4]

            if pd.notna(fechaNacimiento):
                try:
                    # Validación y limpieza del nombre
                    nombre = re.sub(r'\s+', ' ', df.iloc[index, 2].strip())

                    # ---- RUT ----
                    rut_response = df.iloc[index, 3].replace(" ", "").replace(".", "")
                    rut_base = rut_response[:-2]
                    verificador = rut_response[-1]
                    rut = "{:,}".format(int(rut_base)).replace(",", ".") + "-" + verificador
                    rut = rut.upper()

                    # ---- Fecha de nacimiento ----
                    try:
                        fechaFormateada = fechaNacimiento.strftime('%d/%m/%Y')
                    except Exception:
                        fechasInvalidas.append(index)
                        continue  # Omitimos esta fila si la fecha da error

                    # ---- Edad ----
                    try:
                        edad_raw = df.iloc[index, 5]
                        if isinstance(edad_raw, str):
                            edad_raw = edad_raw.strip()
                        edad = int(float(edad_raw))  # Convierte "70.0" → 70
                    except Exception:
                        edadInvalidas.append(index)
                        continue  # Omitimos esta fila si la edad da error

                    # ---- Salida ----
                    print(f"Paciente: {nombre}, Rut: {rut}, Fecha de Nacimiento: {fechaFormateada}, Edad: {edad}")
                    val += 1
                    fechasValidas.append(index)

                except Exception:
                    rutInvalidos.append(index)
                    continue  # Omitimos esta fila si el RUT da error
            else:
                fechasInvalidas.append(index)


    

    


    print("fechas invalidas: ",fechasInvalidas)
    print("rut invalidos: ",rutInvalidos)
    print("edad invalidas: ",edadInvalidas)
    print(fechasValidas)
    print("val", val)

   


    """ 
        print(df.shape[0])

        if pd.isna(df.iloc[df.shape[0]-2,0]):
            print("vacio")
        else:
            print("datos") """
    



else:
    print("No se seleccionó ningún archivo.")
