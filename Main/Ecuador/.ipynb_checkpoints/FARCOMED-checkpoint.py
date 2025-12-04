import pandas as pd
from io import StringIO
from openpyxl import load_workbook
import numpy as np
from utils import *
import os
import config


customer_id = 10299021

def procesar(archivos_remittance,archivo_fbl5n):
    rutas = {
        "remittance": archivos_remittance,
        "fbl5n": archivo_fbl5n,
        "salida": os.path.join(os.path.dirname(archivos_remittance), "Farcomed.xlsx")
    }

    # Leer todas las líneas del archivo
    with open(archivos_remittance, encoding='latin1') as f:
        lineas = f.readlines()

    # Encontrar la línea de inicio (encabezados)
    for i, line in enumerate(lineas):
        if line.strip().startswith("Cía;Cliente;Tipo Doc;Num Doc;Num Factura;Valor;Fecha Doc"):
            start_row = i
            break

    # Encontrar la línea donde aparece "DOCUMENTOS"
    for j, line in enumerate(lineas):
        if "DOCUMENTOS" in line:
            end_row = j
            break
    else:
        end_row = len(lineas)  # Si no se encuentra, leer hasta el final

    # Unir solo las líneas relevantes en un solo string
    contenido_relevante = "".join(lineas[start_row:end_row])

    # Usar StringIO para simular un archivo en memoria
    df = pd.read_csv(
        StringIO(contenido_relevante),
        delimiter=';',
        encoding='latin1',
        usecols=['Tipo Doc','Num Factura', 'Valor']
    )

    # Filtrar filas donde 'Num Factura' NO contiene la palabra 'Total'
    remittance = df[~df['Num Factura'].astype(str).str.contains('Total', case=False, na=False)]

    remittance = remittance.rename(columns={
            "Tipo Doc": "Tipo de Documento",
            "Num Factura": "Referencia / Factura",
            "Valor": "Importe de Remittance"
        })
    
    remittance["Referencia / Factura"] = (
        remittance["Referencia / Factura"]
        .astype(str)
        .str.replace("'", "")
    )

        # Crear columnas adicionales si no existen
    for col in ["Descuento", "Motivo del descuento", "Comentarios"]:
        if col not in remittance.columns:
            remittance[col] = ""

    conds = [
        remittance["Referencia / Factura"].str.startswith("NDQ", na=False),
        remittance["Referencia / Factura"].str.startswith("NDC", na=False),
        remittance["Referencia / Factura"].str.startswith("NDQAUTH#", na=False),
        remittance["Referencia / Factura"].str.startswith("NDQREF#PIO", na=False),
    ]

    descuentos = ["Descuento cliente"] * 4
    motivos = ["551", "100", "388", "388"]

    remittance["Descuento"] = np.select(conds, descuentos, default=remittance["Descuento"])
    remittance["Motivo del descuento"] = np.select(conds, motivos, default=remittance["Motivo del descuento"])

    # Forzar tipo de documento a "Factura" después de aplicar condiciones
    remittance["Tipo de Documento"] = remittance["Tipo de Documento"].apply(
        lambda x: "Factura" if str(x).strip().upper() in ["PV", "V1"]
        else "Nota de crédito" if str(x).strip().upper() == "PD"
        else "Saldo"
    )

    remittance.loc[
        remittance["Motivo del descuento"].notna() & (remittance["Motivo del descuento"].str.strip() != ""),
        "Comentarios"
    ] = remittance["Tipo de Documento"].str.replace(" ", "") + " " + remittance["Referencia / Factura"].astype(str)

    # Procesos adicionales
    FBL5N, id_cliente, nombre_cliente = procesar_cartera_cliente(rutas["fbl5n"], customer_id)
    hrc_template = merge_remittance_cartera(remittance, FBL5N)
    hrc_template = procesar_diferencias(hrc_template)
    hrc_template = procesamiento_nro(hrc_template, FBL5N)

    if "Motivo del descuento" in hrc_template.columns:
        mask = hrc_template["Motivo del descuento"].astype(str).str.strip().eq("384")
        if mask.any():
            hrc_template.loc[mask, "Motivo del descuento"] = "WOB"

    if "Motivo del descuento" in hrc_template.columns:
        mask = hrc_template["Motivo del descuento"].astype(str).str.strip().eq("987")
        if mask.any():
            hrc_template.loc[mask, "Motivo del descuento"] = "388"

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

