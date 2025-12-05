import pdfplumber
import pandas as pd
import re
import os
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side
from tkinter import messagebox
from clientes.utils import *
import sys


# ===========================================================
# 1. Ruta raíz del proyecto (NO SE TOCA)
# ===========================================================
def _project_root():
    if getattr(sys, "frozen", False):
        macos_dir = os.path.dirname(sys.executable)
        contents_dir = os.path.dirname(macos_dir)
        app_bundle = os.path.dirname(contents_dir)
        return os.path.dirname(app_bundle)

    return os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))


# ===========================================================
# 2. CUSTOMER ID del cliente en Perú
# ===========================================================
customer_id = 10449763   # <-- Aquí lo dejas fijo para Perú – NF


# ===========================================================
# 3. PROCESO PRINCIPAL (SE EJECUTA DESDE TKINTER)
# ===========================================================
def procesar():
    try:
        root = _project_root()

        # ======================================================
        # Definir todas las rutas EXACTAS dentro del proyecto
        # ======================================================
        rutas = {
            "remittance": os.path.join(root, "Archivos", "Remittance", "Peru", "Remittance_NF.pdf"),
            "fbl5n":      os.path.join(root, "Archivos", "Cartera", "Peru", "FBL5N_NF.xlsx"),
            "salida":     os.path.join(root, "Archivos", "Template", "Peru", "Template_HRC_NF.xlsx")
        }

        archivos_remittance = rutas["remittance"]
        archivo_fbl5n = rutas["fbl5n"]

        # ======================================================
        # 4. LÓGICA ORIGINAL — NO SE CAMBIA NADA
        # ======================================================
        tablas_filtradas = []
        rows = []

        with pdfplumber.open(archivos_remittance) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        rows.append(row)

        # Buscar encabezado
        header_idx = None
        for i, row in enumerate(rows):
            if row and "Fec Emision" in row:
                header_idx = i
                break

        if header_idx is None:
            raise Exception("No se encontró el encabezado 'Fec Emision' en el PDF.")

        header = rows[header_idx]
        data_rows = rows[header_idx + 1:]
        df = pd.DataFrame(data_rows, columns=header)

        # Eliminar columna
        df = df.drop(columns=["Fec Emision"])

        # Renombrar columnas
        remittance = df.rename(columns={
            "Tipo Doc": "Tipo de Documento",
            "Num Documento": "Referencia / Factura",
            "Monto": "Importe de Remittance"
        })

        # Importe numérico
        remittance["Importe de Remittance"] = (
            remittance["Importe de Remittance"]
            .replace({',': '', ' ': ''}, regex=True)
            .replace('', np.nan)
            .astype(float)
        )

        remittance.loc[
            remittance["Tipo de Documento"].str.contains("FACTURA X COBRAR|NOTA DE CREDITO", case=False, na=False),
            "Importe de Remittance"
        ] *= -1

        # Formato referencia
        remittance["Referencia / Factura"] = remittance["Referencia / Factura"].apply(
            lambda x: f"{x.split('-')[0]}-{x.split('-')[1].zfill(8)}"
            if isinstance(x, str) and '-' in x else x
        )

        # Columnas adicionales
        for col in ["Descuento", "Motivo del descuento", "Comentarios"]:
            if col not in remittance.columns:
                remittance[col] = ""

        # Descuentos
        conds = [
            remittance["Tipo de Documento"].str.startswith("FACTURA X COBRAR", na=True)
        ]
        descuentos = ["Descuento cliente"]
        motivos = ["987"]

        remittance["Descuento"] = np.select(conds, descuentos, default=remittance["Descuento"])
        remittance["Motivo del descuento"] = np.select(conds, motivos, default=remittance["Motivo del descuento"])

        # Comentarios
        remittance.loc[
            remittance["Motivo del descuento"].notna() &
            (remittance["Motivo del descuento"].str.strip() != ""),
            "Comentarios"
        ] = remittance["Referencia / Factura"].astype(str)

        # Procesos adicionales
        FBL5N, id_cliente, nombre_cliente = procesar_cartera_cliente(archivo_fbl5n, customer_id)
        hrc_template = merge_remittance_cartera(remittance, FBL5N)
        hrc_template = procesar_diferencias(hrc_template)
        hrc_template = procesamiento_nro(hrc_template, FBL5N)

        # Ajustes finales
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

        # Guardar remittance temporal
        ruta_temp = os.path.join(os.path.dirname(archivos_remittance), "remittance_temp.xlsx")
        remittance.to_excel(ruta_temp, index=False)

        # Exportar template final
        exportar_template(
            hrc_template=hrc_template,
            suma_remittance=remittance["Importe de Remittance"].sum(),
            numero_orden="",
            id_cliente=id_cliente,
            nombre_cliente=nombre_cliente,
            ruta_remittance=ruta_temp,
            ruta_salida=rutas["salida"]
        )

        if os.path.exists(ruta_temp):
            os.remove(ruta_temp)

        messagebox.showinfo("Éxito", f"Archivo exportado correctamente en:\n{rutas['salida']}")
        return hrc_template

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error:\n{e}")
        raise e
