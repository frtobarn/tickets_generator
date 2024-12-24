import pandas as pd
import shutil
import qrcode
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, KeepTogether, Image
from reportlab.lib.units import mm
from datetime import datetime
import pytz

def formatear_fecha():
    timezone = pytz.timezone("America/Bogota")
    now = datetime.now(timezone)
    dias_semana = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]
    meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
    dia_semana = dias_semana[now.weekday()]
    dia = now.day
    mes = meses[now.month - 1]
    anio = now.year
    hora = now.strftime("%I").lstrip('0')
    minutos = now.strftime("%M")
    am_pm = now.strftime("%p").lower()
    fecha_formateada = f"{dia_semana} {dia} de {mes} de {anio} a las {hora}:{minutos} {am_pm}"
    return fecha_formateada, f"{anio}_{mes}_{dia}_{dia_semana}_{hora}_{minutos}_{am_pm}"

def cargar_datos(id_hoja, registro_inicial, registro_final, gc):
    hoja = gc.open_by_key(id_hoja).sheet1
    datos = hoja.get_all_values()
    registro_actual = registro_inicial - 1
    data = pd.DataFrame(datos[registro_actual:registro_final], columns=datos[1])
    return data

def procesar_datos(data):
    resultados = []
    for _, fila in data.iterrows():
        resultado_actual = {
            'Nombre': fila['NOMBRE DE USUARIO'],
            'Cedula': fila['N° Identificación'],
            'Direccion': fila['DIRECCIÓN'],
            'Localidad': fila['LOCALIDAD'],
            'Barrio': fila['BARRIO'],
            'Telefono': fila['N° Telefono'],
            'Biblioteca': fila['MATERIAL 9']
        }
        resultados.append(resultado_actual)
    return pd.DataFrame(resultados)

def generar_tiquet(pdf_elements, datos_tiquet, img_path, banner, fecha_formateada, nombre_biblioteca):
    img = Image(img_path)
    img.drawHeight = 4.5 * mm
    img.drawWidth = 36.2 * mm

    qr = qrcode.QRCode(version=1, error_correction=qrcode.constants.ERROR_CORRECT_L, box_size=10, border=4)
    qr.add_data(f"{datos_tiquet['Cedula']}\t{datos_tiquet['Nombre']}\t{datos_tiquet['Direccion']}")
    qr.make(fit=True)
    img_qr = qr.make_image(fill_color="black", back_color="white")
    img_qr_path = f"/content/{datos_tiquet['Cedula']}_qr.png"
    img_qr.save(img_qr_path)

    img_qr = Image(img_qr_path)
    img_qr.drawHeight = 75
    img_qr.drawWidth = 75

    data = [
        [img, "", "Servicio de préstamo a domicilio", "", ""],
        ["Datos de origen", "Fecha de alistamiento", fecha_formateada, "", img_qr],
        ["", "Biblioteca que envía", nombre_biblioteca, "", ""],
        ["", "Biblioteca que recibe", datos_tiquet['Biblioteca'], "", ""],
        ["Datos del usuario", "Nombre", datos_tiquet['Nombre'], "", ""],
        ["", "Dirección", datos_tiquet['Direccion'], "", ""],
        ["", "Barrio", datos_tiquet['Barrio'], "Localidad", datos_tiquet['Localidad']],
        ["", "Documento", datos_tiquet['Cedula'], "Teléfono", datos_tiquet['Telefono']],
        ["", banner, "", "", ""]
    ]

    table = Table(data, colWidths=[80, 120, 180, 100, 75])
    table.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('SPAN', (0, 0), (1, 0)),
        ('SPAN', (2, 0), (4, 0)),
        ('SPAN', (1, 1), (3, 1)),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ]))
    pdf_elements.append(KeepTogether(table))

def generar_pdf(df_resultados, img_path, banner, fecha_formateada, nombre_biblioteca, archivo_salida):
    doc = SimpleDocTemplate(archivo_salida, pagesize=letter, leftMargin=5, rightMargin=5, topMargin=5, bottomMargin=5)
    elements = []
    for _, datos_tiquet in df_resultados.iterrows():
        generar_tiquet(elements, datos_tiquet, img_path, banner, fecha_formateada, nombre_biblioteca)
    doc.build(elements)
