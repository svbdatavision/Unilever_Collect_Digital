# =====================================================
# 0. Importación de librerías y módulos utilitarios
# =====================================================
import sys  # Para detectar ejecución empaquetada (frozen) y resolver rutas
import os   # Para construir rutas relativas al proyecto
import pandas as pd  # Principal herramienta para manipulación tabular
import numpy as np   # Utilidades numéricas/condicionales (np.select, np.where)
from openpyxl import load_workbook  # Leer valores dinámicos desde archivos Excel
import re

from utils import *

# =====================================================
# 1. Localización dinámica de la carpeta raíz del proyecto
# =====================================================

# =====================================================
# 2. Función principal del proceso (procesar)
# =====================================================
def procesar(archivo_remittance,archivo_fbl5n):
    """
    Orquestador principal para OXXO:
    Flujo numerado siguiendo el estándar de procesos.
    Devuelve un diccionario con los 4 DataFrames finales listos para exportar.
    """
    # =====================================================
    # 2.1 Definición de rutas de entrada y salida
    # =====================================================
    rutas = {
        "remittance": archivo_remittance,
        "fbl5n": archivo_fbl5n,
        "salida": os.path.join(os.path.dirname(archivo_remittance))
    }
    customer_id = 10324901

    # =====================================================
    # 2.2 Lectura de Remittance
    # =====================================================
    remittance = pd.read_excel(rutas["remittance"], skiprows=11)
    if remittance.columns[0].startswith("Unnamed"):
        remittance = remittance.iloc[:, 1:]
    remittance = remittance.dropna(subset=["Unidad Operativa","Número Factura"])
    df = remittance.copy()

    # =====================================================
    # 2.3 Función para convertir filas a celdas
    # =====================================================
    def row_to_cells(row):
        vals = [str(x).strip() for x in row if str(x).strip() != ""]
        if len(vals) == 0:
            return []
        if len(vals) == 1:
            s = vals[0]
            if "\t" in s:
                parts = re.split(r'\t+', s)
            else:
                parts = re.split(r'\s{2,}', s)
            parts = [p.strip() for p in parts if p.strip() != ""]
            return parts
        else:
            return vals

    rows = [row_to_cells(r.tolist()) for _, r in df.iterrows()]

    # =====================================================
    # 2.4 Detección de encabezados
    # =====================================================
    header_positions = [i for i, cells in enumerate(rows) if any("unidad operativa" in str(c).lower() for c in cells)]
    if 0 not in header_positions:
        header_positions = [0] + header_positions
    header_positions.append(len(rows))
    default_header = ["Unidad Operativa", "Número Pago", "Número Factura", "Fecha Factura",
                      "Importe Factura", "Importe Pagado", "Impuesto", "Retención"]

    # =====================================================
    # 2.5 División en tablas
    # =====================================================
    tables = []
    for i in range(len(header_positions)-1):
        start = header_positions[i]
        if any("unidad operativa" in str(c).lower() for c in rows[start]):
            start += 1
        end = header_positions[i+1]
        block = rows[start:end]
        if len(block) == 0:
            continue
        header_row = rows[header_positions[i]] if any("unidad operativa" in str(c).lower() for c in rows[header_positions[i]]) else default_header
        cols = header_row if len(header_row) >= 3 else default_header
        norm_block = []
        for r_cells in block:
            if len(r_cells) == 0:
                continue
            if len(r_cells) >= len(cols):
                row_fixed = r_cells[:len(cols)]
            else:
                row_fixed = r_cells + [None] * (len(cols) - len(r_cells))
            norm_block.append(row_fixed)
        subdf = pd.DataFrame(norm_block, columns=cols)
        tables.append(subdf.reset_index(drop=True))

    # =====================================================
    # 2.6 Función de transformación de remittance
    # =====================================================
    def transform_remittance(df: pd.DataFrame) -> pd.DataFrame:
        cols_to_keep = ["Número Factura", "Importe Factura", "Número Pago", "Impuesto"]
        df = df[cols_to_keep].copy()
        df = df.rename(columns={
            "Número Factura": "Referencia / Factura",
            "Importe Factura": "Importe de Remittance"
        })
        df["Importe de Remittance"] = pd.to_numeric(df["Importe de Remittance"], errors="coerce").round(2)
        df["Impuesto"] = pd.to_numeric(df["Impuesto"], errors="coerce").round(2)
        df["Referencia / Factura"] = df["Referencia / Factura"].str.replace(r"[\u202A-\u202E\u200E\u200F]", "", regex=True).str.strip()
        conds = [
            df["Referencia / Factura"].str.startswith("PMP", na=False) & (df["Importe de Remittance"] > 0),
            ~df["Referencia / Factura"].str.startswith("PMP", na=False) & (df["Importe de Remittance"] > 0),
            df["Importe de Remittance"] < 0
        ]
        choices = ["Factura", "Nota Debito", "Descuentos Clientes"]
        df["Tipo de Documento"] = np.select(conds, choices, default="")
        df.reset_index(drop=True, inplace=True)
        return df

    # =====================================================
    # 2.7 Función asignar_cards
    # =====================================================
    def asignar_cards(remittance: pd.DataFrame) -> pd.DataFrame:
        if "Descuento" not in remittance.columns:
            remittance["Descuento"] = ""
        if "Motivo del descuento" not in remittance.columns:
            remittance["Motivo del descuento"] = ""
        conds_desc = [
            remittance["Referencia / Factura"].str.startswith("NCM", na=False),
            remittance["Referencia / Factura"].str.startswith("74", na=False),
            remittance["Referencia / Factura"].str.startswith("90", na=False),
            remittance["Referencia / Factura"].str.startswith("91", na=False),
            remittance["Referencia / Factura"].str.startswith("92", na=False),
            remittance["Referencia / Factura"].str.startswith("395", na=False),
            remittance["Referencia / Factura"].str.startswith("396", na=False),
            remittance["Referencia / Factura"].str.startswith("397", na=False)
        ]
        descuentos = ["CONVENIOS", "DSCT PROMOCIONAL", "DSCT PROMOCIONAL", "FACT PROVEEDOR",
                      "CONVENIOS", "DPP", "RECHAZOS", "DSCT AVERIAS"]
        motivos = ["657", "987", "987", "CSB", "657", "206", "551", "522"]
        remittance["Descuento"] = np.select(conds_desc, descuentos, default=remittance["Descuento"])
        remittance["Motivo del descuento"] = np.select(conds_desc, motivos, default=remittance["Motivo del descuento"])
        return remittance

    # =====================================================
    # 2.8 Procesar cada tabla y generar templates
    # =====================================================
    resultados = {}
    for idx, df in enumerate(tables, start=1):
        df = transform_remittance(df)
        df = asignar_cards(df)
        df["Descuento"] = df["Descuento"].fillna("").astype(str)
        df = procesar_descuentos_y_comentarios(df)
        FBL5N, id_cliente, nombre_cliente = procesar_cartera_cliente(rutas["fbl5n"], customer_id)
        hrc_template = merge_remittance_cartera(df, FBL5N)
        hrc_template["Referencia / Factura"] = hrc_template["Referencia / Factura"].str.replace(r"^NC-", "", regex=True)
        hrc_template = procesar_diferencias(hrc_template)
        hrc_template = procesamiento_nro(hrc_template, FBL5N)
        hrc_template["Pago Neto"] = hrc_template["Importe de factura"]
        hrc_template["Referencia / Factura"] = hrc_template["Referencia / Factura"].str.replace(r"^CNQ", "", regex=True)
        columnas_finales = ["Tipo de Documento", "Referencia / Factura", "Importe de factura",
                            "Descuento", "Motivo del descuento", "Pago Neto", "Comentarios"]
        hrc_template = hrc_template[columnas_finales]

        wb_rem = load_workbook(rutas["remittance"], data_only=True)
        ws_rem = wb_rem.active
        numero_orden = ws_rem["B7"].value
        
        nombre_archivo = f"Template_HRC_oxxo_{idx}.xlsx"
        ruta_salida = os.path.join(rutas["salida"], nombre_archivo)
        
        exportar_template(
            hrc_template=hrc_template,
            suma_remittance=df["Importe de Remittance"].sum(),
            numero_orden=numero_orden,
            id_cliente=id_cliente,
            nombre_cliente=nombre_cliente,
            ruta_remittance=rutas["remittance"],
            ruta_salida=ruta_salida
        )

        resultados[f"remittance_{idx}"] = hrc_template

    # =====================================================
    # 2.9 Devolución de los templates finales
    # =====================================================
    return resultados
