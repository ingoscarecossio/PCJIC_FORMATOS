from datetime import date
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import pandas as pd
import app

out = app.APP_DIR / 'salida_demo'
out.mkdir(exist_ok=True)

datos = {
    'programa': 'Programa de prueba',
    'asignatura': 'Hidrología Aplicada',
    'codigo': 'HID-101',
    'grupo': '01',
    'profesor': 'Oscar Cossio',
    'correo': 'docente@institucion.edu.co',
    'periodo': '2026-2',
    'creditos': '3',
    'htp': 48,
    'hti': 96,
    'fecha_inicio': date(2026, 8, 3),
    'fecha_fin': date(2026, 11, 20),
    'estado': 'Planeación',
    **app.TEXTOS_PREDEFINIDOS_GC71,
}
modulos = pd.DataFrame([
    {'Unidad': 'UNIDAD 1', 'Contenido / tema central': 'Fundamentos hidrológicos', 'Horas presenciales': 12, 'Sesiones': 6, 'Trabajo presencial': 'Clase orientadora y ejercicios aplicados.', 'Trabajo independiente': 'Lectura y taller diagnóstico.'},
    {'Unidad': 'UNIDAD 2', 'Contenido / tema central': 'Análisis de precipitación y caudales', 'Horas presenciales': 18, 'Sesiones': 9, 'Trabajo presencial': 'Taller con datos y discusión.', 'Trabajo independiente': 'Procesamiento de series.'},
    {'Unidad': 'UNIDAD 3', 'Contenido / tema central': 'Modelación y proyecto integrador', 'Horas presenciales': 18, 'Sesiones': 9, 'Trabajo presencial': 'Laboratorio y socialización.', 'Trabajo independiente': 'Proyecto final.'},
])
horarios = pd.DataFrame([
    {'Día': 'Lunes', 'Hora inicio': '18:00', 'Hora fin': '20:00', 'Lugar / ambiente': 'Aula 301'},
    {'Día': 'Miércoles', 'Hora inicio': '18:00', 'Hora fin': '20:00', 'Lugar / ambiente': 'Laboratorio SIG'},
])
fechas = app.generar_fechas_clase(datos['fecha_inicio'], datos['fecha_fin'], horarios)
sesiones = app.expandir_plan_sesiones(modulos, fechas, criterio='Horas presenciales')
evaluaciones = pd.DataFrame([
    {'Tipo de evaluación': 'Seguimiento', 'Procedimiento de evaluación': 'Talleres y actividades aplicadas', 'Valor (%)': 30, 'Fecha de realización': '2026-09-01', 'Unidad relacionada': 'UNIDAD 1', 'Corte': 'Primer corte'},
    {'Tipo de evaluación': 'Parcial', 'Procedimiento de evaluación': 'Evaluación individual teórico-práctica', 'Valor (%)': 30, 'Fecha de realización': '2026-10-01', 'Unidad relacionada': 'UNIDAD 2', 'Corte': 'Segundo corte'},
    {'Tipo de evaluación': 'Proyecto final', 'Procedimiento de evaluación': 'Entrega y sustentación', 'Valor (%)': 40, 'Fecha de realización': '2026-11-18', 'Unidad relacionada': 'UNIDAD 3', 'Corte': 'Final'},
])
estudiantes = pd.DataFrame([
    {'Nombre completo': 'Estudiante Uno', 'Documento': '1001', 'Correo': 'uno@correo.com', 'Plan': '', 'Observación': '', 'Estado': 'Activo'},
    {'Nombre completo': 'Estudiante Dos', 'Documento': '1002', 'Correo': 'dos@correo.com', 'Plan': '', 'Observación': '', 'Estado': 'Activo'},
])
reps = pd.DataFrame(columns=['Nombre de los estudiantes', 'N° de cédula o carné estudiantil', 'Firma'])

(out / 'demo_FD-GC71.docx').write_bytes(app.crear_gc71_docx(datos, sesiones, evaluaciones, reps))
(out / 'demo_Evaluacion.xlsx').write_bytes(app.crear_plantilla_evaluacion_xlsx(estudiantes, evaluaciones, datos))
(out / 'demo_Calendario.ics').write_bytes(app.build_ics_calendar(sesiones, datos))
(out / 'demo_Paquete.zip').write_bytes(app.crear_paquete_curso_zip(datos, sesiones, evaluaciones, estudiantes, reps, None))
print('OK', out)
