import fitz  # PyMuPDF
import camelot
import pandas as pd
import re
import os
from utils import *
import config

customer_id = 10273158

def procesar(archivos_remittance, archivo_fbl5n):
    tablas_filtradas = []
    rutas = {
        "remittance": archivos_remittance,
        "fbl5n": archivo_fbl5n,
        "salida": os.path.join(os.path.dirname(archivos_remittance), "Difare.xlsx")
    }

    doc = fitz.open(archivos_remittance)
    total_pages = len(doc)

    tables = camelot.read_pdf(archivos_remittance, pages=f"2-{total_pages}", flavor='stream')

    rows = []

    for table in tables:
        df = table.df
        for idx in range(1, len(df)):
            fila = df.iloc[idx]
            fila_str = ' '.join([str(x) for x in fila if str(x).strip() != ''])
            tipo_doc_match = re.search(r'(Factura MM|NC Acreedores Elec)', fila_str)
            referencia_match = re.search(r'(\d{9,})', fila_str)
            valor_match = re.search(r'(-?\d+\.\d+)', fila_str)
            if tipo_doc_match and referencia_match and valor_match:
                tipo_doc_raw = tipo_doc_match.group(1)
                referencia = referencia_match.group(1)
                valor = float(valor_match.group(1))
                if tipo_doc_raw == "Factura MM":
                    tipo_doc = "Factura"
                else:
                    tipo_doc = "Nota de crédito"
                    valor = -abs(valor)
                rows.append([tipo_doc, referencia, valor])

    remittance = pd.DataFrame(rows, columns=['Tipo de Documento', 'Referencia / Factura', 'Importe de Remittance'])
    # Procesos adicionales
    FBL5N, id_cliente, nombre_cliente = procesar_cartera_cliente(rutas["fbl5n"], customer_id)
    hrc_template = merge_remittance_cartera(remittance, FBL5N)
    hrc_template = procesar_diferencias(hrc_template)
    hrc_template = procesamiento_nro(hrc_template, FBL5N)

    if "Motivo del descuento" in hrc_template.columns:
        mask = hrc_template["Motivo del descuento"].astype(str).str.strip().eq("384")
        if mask.any():
            hrc_template.loc[mask, "Motivo del descuento"] = "WOB"

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
