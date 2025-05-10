from datetime import datetime
import os

URL_BASE_PROYECTO = os.path.dirname(os.path.abspath(__file__)) #carpeta donde esta el script
CARPETA_BASE = "test_folder"
HOY_D_M_Y = datetime.today().strftime('%d-%m-%Y')
DESDENTADO_SI = os.path.join(URL_BASE_PROYECTO,"templates\plantilla0.docx")
DESDENTADO_NO = os.path.join(URL_BASE_PROYECTO,"templates\plantilla9990.docx")
VALIDACION_DESDENTADO = "desdentado"

ERRORES =  {}
PACIENTE = {
"nombre": "",
"rut":"",
"fecha_nacimiento":"",
"edad":0,
"id": 0,
"desdentado":""
}
COLUMNAS_VALIDAS=['n°_paciente', 'caso', 'nombre_paciente', 'rut', 'fecha_de_nacimiento',
       'edad', 'plan_de_tratamiento_ingreso_(link)',
       'link_plan_de_tratamiento_(pdt)_(dis)', 'precio_pdt_ingreso',
       'aprobación_fundación_(pdt_diagnóstico)',
       'órden_de_atención_entregada_al_centro.1',
       'n°_orden_de_atención_(folio)_pdf',
       'disponibilidad_paciente_(diagnóstico)_opción_1',
       'disponibilidad_paciente_(diagnóstico)_opción_2',
       'profesional_que_realiza_diagnóstico',
       'cita_(fecha_inicio,_diagnóstico_+_rx+_higiene)',
       'asistencia_(diagnóstico)', 'pdt_tratamiento',
       'link_pdf_pdt_tratamiento', 'precio_pdt_tratamiento',
       'aprobación_fundación_(pdt_tratamiento)',
       'órden_de_atención_entregada_al_centro', 'profesional_tratante',
       'fecha_1ra_cita', 'asistencia_1ra_cita', 'fecha_2da_cita',
       'asistencia_2da_cita', 'fecha_3ra_cita', 'asistencia_3ra_cita',
       'fecha_4ta_cita', 'asistencia_4ta_cita', 'fecha_5ta_cita',
       'asistencia_5ta_cita', 'fecha_6ta_cita', 'asistencia_6ta_cita',
       'fecha_7ma_cita', 'asistencia_7ma_cita', 'fecha_8va_cita',
       'asistencia_8va_cita', 'fecha_9na_cita', 'asistencia_9na_cita',
       'asistencia_(fecha_término)', 'certificado_alta_(link)', 'observaciones']

