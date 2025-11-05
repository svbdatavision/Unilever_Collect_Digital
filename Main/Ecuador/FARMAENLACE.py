import pdfplumber
import pandas as pd
import os
import numpy as np
from openpyxl import load_workbook
from utils import *
import config


encabezado_objetivo = (
    'Documento SAP',
    'Factura / Ref.',
    'Nro. Comprobante',
    'Importe',
    'Retención',
    'Retención I.V.A.'
)

customer_id = 10397487

def procesar(archivos_remittance, archivo_fbl5n):
    tablas_filtradas = []
    rutas = {
        "remittance": archivos_remittance,
        "fbl5n": archivo_fbl5n,
        "salida": os.path.join(os.path.dirname(archivos_remittance), "Farmaenlace.xlsx")
    }

    # Extraer tablas del PDF
    with pdfplumber.open(archivos_remittance) as pdf:
        for pagina in pdf.pages:
            tablas = pagina.extract_tables()
            for tabla in tablas:
                if tabla and len(tabla) > 1:
                    encabezado = tuple(tabla[0])
                    if encabezado == encabezado_objetivo:
                        df = pd.DataFrame(tabla[1:], columns=tabla[0])
                        tablas_filtradas.append(df)

    # Unir todas las tablas encontradas
    resultado = pd.concat(tablas_filtradas, ignore_index=True)

    # Filtrar filas válidas
    resultado_filtrado = resultado[
        resultado['Factura / Ref.'].notna() & (resultado['Factura / Ref.'].str.strip() != '')
    ]

    # Seleccionar columnas necesarias
    columnas_deseadas = ['Factura / Ref.', 'Importe']
    remittance = resultado_filtrado[columnas_deseadas]

    # Limpiar y convertir la columna 'Importe' a float
    remittance["Importe"] = (
        remittance["Importe"]
        .astype(str)
        .str.replace('.', '', regex=False)
        .str.replace(',', '.', regex=False)
        .str.replace(' ', '', regex=False)
        .astype(float)
    )

    # Crear columna 'Tipo de Documento' según el valor de 'Importe'
    remittance['Tipo de Documento'] = np.where(remittance['Importe'] < 0, 'Nota de crédito', 'Factura')

    # Renombrar columnas finales
    remittance = remittance.rename(columns={
        "Factura / Ref.": "Referencia / Factura",
        "Importe": "Importe de Remittance"
    })

    # Procesos adicionales
    FBL5N = procesar_cartera_cliente(rutas["fbl5n"], customer_id)
    hrc_template = merge_remittance_cartera(remittance, FBL5N)
    hrc_template = procesar_diferencias(hrc_template)
    hrc_template = procesamiento_nro(hrc_template, FBL5N)

    # Ajustar columnas finales
    hrc_template["Pago Neto"] = hrc_template["Importe de factura"]
    columnas_finales = [
        "Tipo de Documento",
        "Referencia / Factura",
        "Importe de factura",
        "Descuento",
        "Motivo del descuento",
        "Pago Neto",
        "Comentarios"
    ]
    hrc_template = hrc_template[columnas_finales]

    # Datos adicionales
    numero_orden = ""
    id_cliente = customer_id
    nombre_cliente = "FARMAENLACE CIA LTDA"

    # Guardar remittance temporal en Excel
    ruta_remittance_excel = os.path.join(os.path.dirname(archivos_remittance), "remittance_temp.xlsx")
    remittance.to_excel(ruta_remittance_excel, index=False)

    # Exportar template final
    exportar_template(
        hrc_template=hrc_template,
        suma_remittance=remittance["Importe de Remittance"].sum(),
        numero_orden=numero_orden,
        id_cliente=id_cliente,
        nombre_cliente=nombre_cliente,
        ruta_remittance=ruta_remittance_excel,
        ruta_salida=rutas["salida"]
    )

    # (Opcional) Eliminar el archivo temporal
    if os.path.exists(ruta_remittance_excel):
        os.remove(ruta_remittance_excel)

    return hrc_template

