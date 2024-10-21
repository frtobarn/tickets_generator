import pandas as pd
import shutil
import gspread
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, KeepTogether, Image
from reportlab.lib.units import mm
from io import BytesIO
from google.auth import default
from datetime import datetime
import pytz
import qrcode

# Autenticación de Google
creds, _ = default()
gc = gspread.authorize(creds)

# Función para generar el tiquete
def generar_tiquet(pdf_elements, datos_tiquet, img_path, nombre_biblioteca, fecha_formateada, banner):
    print(datos_tiquet)

    # Generar el código QR
    datos = f"{datos_tiquet['N° Identificación']}\t{datos_tiquet['Nombre']}\t{datos_tiquet['Direccion']} {datos_tiquet['Localidad']} {datos_tiquet['Barrio']}\t{datos_tiquet['Telefono']}"
    qr = qrcode.QRCode(version=1, error_correction=qrcode.constants.ERROR_CORRECT_L, box_size=10, border=4)
    qr.add_data(datos)
    qr.make(fit=True)
    img_qr = qr.make_image(fill_color="black", back_color="white")

    # Guardar temporalmente la imagen del QR
    img_qr_path = f"/content/{datos_tiquet['Cedula']}_qr.png"
    img_qr.save(img_qr_path)

    # Crear un objeto Image de ReportLab para insertar en la tabla
    img_qr = Image(img_qr_path)
    img_qr.drawHeight = 75
    img_qr.drawWidth = 75

    # Preparar el logo
    img = Image(img_path)
    img.drawHeight = 4.5 * mm
    img.drawWidth = 36.2 * mm

    # Estructura del tiquete
    data = [
        [img, "", "Servicio de préstamo a domicilio", "", ""],
        ["Datos\nde\norigen", "Fecha de alistamiento", fecha_formateada, "", img_qr],
        ["", "Biblioteca que envía", nombre_biblioteca, "", ""],
        ["", "Biblioteca que recibe", datos_tiquet['Biblioteca'], "", ""],
        ["Datos del\nusuario\nsolicitante", "Nombre", datos_tiquet['Nombre'], "", ""],
        ["", "Dirección", datos_tiquet['Direccion'], "", ""],
        ["", "Barrio", datos_tiquet['Barrio'], "Localidad", datos_tiquet['Localidad']],
        ["", "No. de documento", datos_tiquet['Cedula'], "No. de teléfono", datos_tiquet['Telefono']],
        ["Total\nsolicitud", "Libros", "", "Audiovisuales", ""],
        ["", banner, "", "", ""]
    ]

    # Anchos de las columnas
    column_widths = [2.1 * 28.35, 4.1 * 28.35, 6 * 28.35, 3.0 * 28.35, 5 * 28.35]
    table = Table(data, colWidths=column_widths)

    # Estilo de la tabla
    style = TableStyle([

      #Todo el documento
      ('GRID', (0, 0), (-1, -1), 1, colors.black),
      ('WORDWRAP', (0, 0), (-1, -1)),
      ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
      ('FONT', (0, 0), (-1, -1), 'Helvetica-Bold'),
      ('FONTSIZE', (0, 0), (-1, -1), 10),
      ('TOPPADDING', (0, 0), (-1, -1), 1),  # Eliminar padding superior
      ('BOTTOMPADDING', (0, 0), (-1, -1), 1),  # Eliminar padding inferior

      # Logos
      ('SPAN', (0, 0), (1, 0)),
      ('ALIGN', (0, 0), (2, 0), 'CENTER'),

      # servicio de prestsamo a domicilio
      ('SPAN', (2, 0), (4, 0)),
      ('ALIGN', (2, 0), (4, 0), 'CENTER'),
      ('BACKGROUND', (2, 0), (4, 0), colors.grey),
      ('TEXTCOLOR', (2, 0), (4, 0), colors.white),

      #Datos de origen
      ('SPAN', (0, 1), (0, 3)),
      ('ALIGN', (0, 1), (0, 3), 'LEFT'),
      ('BACKGROUND', (0, 1), (0, 3), colors.grey),
      ('TEXTCOLOR', (0, 1), (0, 3), colors.white),

      #Qr
      ('SPAN', (4, 1), (4, 4)),
      ('ALIGN',(4, 1), (4, 4), 'CENTER'),
      ('TOPPADDING', (4, 1), (4, 3), 0),  # Eliminar padding superior
      ('BOTTOMPADDING', (4, 1), (4, 3), 0),  # Eliminar padding inferior

      # espacio para fecha
      ('SPAN', (2, 1), (3, 1)),
      ('ALIGN', (2, 1), (3, 1), 'CENTER'),

      # espacio para quien envía
      ('SPAN', (2, 2), (3, 2)),
      ('ALIGN', (2, 2), (3, 2), 'CENTER'),


      # espacio para quien recibe
      ('SPAN', (2, 3), (3, 3)),
      ('ALIGN', (2, 3), (3, 3), 'CENTER'),

      #Datos del usuario solicitante
      ('SPAN', (0, 4), (0, 7)),
      ('ALIGN', (0, 4), (0, 7), 'LEFT'),
      ('BACKGROUND', (0, 4), (0, 7), colors.grey),
      ('TEXTCOLOR', (0, 4), (0, 7), colors.white),

      # espacio para nombre
      ('SPAN', (2, 4), (3, 4)),
      ('ALIGN', (2, 4), (3, 4), 'CENTER'),

      # espacio para dirección
      ('SPAN', (2, 5), (4, 5)),
      ('ALIGN', (2, 5), (4, 5), 'CENTER'),

      # Sombreado del total
      ('BACKGROUND', (0, 8), (0, 9), colors.grey),
      ('TEXTCOLOR', (0, 8), (0, 9), colors.white),

      # novedades
      ('SPAN', (1, 9), (4, 9)),
  ])

    table.setStyle(style)
    pdf_elements.append(KeepTogether(table))
