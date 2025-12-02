import pdfplumber
import pandas as pd
import os
import numpy as np
from openpyxl import load_workbook
from utils import *
import config

encabezado_objetivo = (
    'Nro.', 
    'Fecha\nDocumento', 
    'Tipo Documento', 
    'Número\nDocumento', 
    'Valor', 
    'Retención en la\nfuente IR', 
    'Retención IVA', 
    'Valor Pagado'
)

customer_id = 10273281

def procesar(archivos_remittance, archivo_fbl5n):
    tablas_filtradas = []
    rutas = {
        "remittance": archivos_remittance,
        "fbl5n": archivo_fbl5n,
        "salida": os.path.join(os.path.dirname(archivos_remittance), "Favorita.xlsx")
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
        resultado['Tipo Documento'].notna() & (resultado['Tipo Documento'].str.strip() != '')
    ]

    # Seleccionar columnas necesarias
    columnas_deseadas = ['Tipo Documento', 'Número\nDocumento', 'Valor Pagado']
    remittance = resultado_filtrado[columnas_deseadas]

    # Renombrar columnas
    remittance = remittance.rename(columns={
        "Tipo Documento": "Tipo de Documento",
        "Número\nDocumento": "Referencia / Factura",
        "Valor Pagado": "Importe de remittance"
    })

    # Convertir valores a numéricos y manejar paréntesis como negativos
    remittance["Importe de Remittance"] = (
        remittance["Importe de remittance"]
        .replace(r'[\$,]', '', regex=True)
        .apply(lambda x: -float(x.replace('(', '').replace(')', '')) if isinstance(x, str) and '(' in x else float(x))
    )

    # Guardar valor original antes de forzar "Factura"
    remittance["Tipo Documento Original"] = remittance["Tipo de Documento"]

    # Crear columnas adicionales si no existen
    for col in ["Descuento", "Motivo del descuento", "Comentarios"]:
        if col not in remittance.columns:
            remittance[col] = ""

    # Condiciones y motivos
    conds = [
        remittance["Tipo de Documento"].str.startswith("CR-Factura", na=False),
        remittance["Tipo de Documento"].str.startswith("RI-Factura manual", na=False),
        remittance["Tipo de Documento"].str.startswith("DV-Devolución a Proveedores", na=False),
        remittance["Tipo de Documento"].str.startswith("RT-Retenciones", na=False),
        remittance["Tipo de Documento"].str.startswith("N2-Notas de Cobro Comercia", na=False)
    ]

    descuentos = ["Descuento cliente"] * 5
    motivos = ["986", "986", "659", "WHT", "388"]

    remittance["Descuento"] = np.select(conds, descuentos, default=remittance["Descuento"])
    remittance["Motivo del descuento"] = np.select(conds, motivos, default=remittance["Motivo del descuento"])

    # Forzar tipo de documento a "Factura" después de aplicar condiciones
    remittance["Tipo de Documento"] = "Factura"

    # Crear columna Comentarios si hay motivo de descuento
    remittance.loc[
        remittance["Motivo del descuento"].notna() & (remittance["Motivo del descuento"].str.strip() != ""),
        "Comentarios"
    ] = remittance["Tipo Documento Original"].str.replace(" ", "") + " " + remittance["Referencia / Factura"].astype(str)

    # Procesos adicionales
    FBL5N = procesar_cartera_cliente(rutas["fbl5n"], customer_id)
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
    id_cliente = customer_id
    nombre_cliente = "CORPORACION FAVORITA C A"

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
