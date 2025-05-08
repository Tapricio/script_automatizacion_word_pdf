from datetime import datetime
import os

URL_BASE_PROYECTO = os.path.dirname(os.path.abspath(__file__)) #carpeta donde esta el script
CARPETA_BASE = "test_folder"
HOY_D_M_Y = datetime.today().strftime('%d-%m-%Y')
DESDENTADO_SI = os.path.join(URL_BASE_PROYECTO,"templates\plantilla0.docx")
DESDENTADO_NO = os.path.join(URL_BASE_PROYECTO,"templates\plantilla9990.docx")
VALIDACION_DESDENTADO = "desdentado"

