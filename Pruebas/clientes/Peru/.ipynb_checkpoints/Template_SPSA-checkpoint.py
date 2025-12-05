import os
import sys
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side
from tkinter import messagebox
from clientes.utils import *


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
# 2. CUSTOMER ID del cliente en Perú – SPSA
# ===========================================================
customer_id = 10263165


# ===========================================================
# 3. PROCESO PRINCIPAL
# ===========================================================
def procesar():
    try:
        # ===============================================
        # RUTAS INTERNAS (INPUT + OUTPUT)
        # ===============================================
        root = _project_root()

        rutas = {
            "remittance": os.path.join(root, "Archivos", "Remittance", "Peru", "Remittance_SPSA.xlsx"),
            "fbl5n":      os.path.join(root, "Archivos", "Cartera", "Peru", "FBL5N_SPSA.xlsx"),
            "salida":     os.path.join(root, "Archivos", "Template", "Peru", "Template_HRC_SPSA.xlsx")
        }

        archivos_remittance = rutas["remittance"]
        archivo_fbl5n = rutas["fbl5n"]

        # ======================================================
        # 4. LÓGICA ORIGINAL — NO SE CAMBIA
        # ======================================================
        df = pd.read_excel(archivos_remittance, sheet_name='Pagos(detalle)', engine='openpyxl')

        columnas_deseadas = [
            'Descripción Tipo Documento',
            'Nro. Documento Proveedor',
            'Importe Documento'
        ]
        remittance = df[columnas_deseadas]

        tipo_map = {
            "Nota de Débito": "Nota de Débito",
            "Nota de Crédito": "Nota de Crédito",
            "Factura Mercadería": "Factura",
            "Facturas por Cobrar": "Fact. Convenio",
        }

        remittance = remittance.rename(columns={
            "Descripción Tipo Documento": "Tipo de Documento",
            "Nro. Documento Proveedor": "Referencia / Factura",
            "Importe Documento": "Importe de Remittance"
        })

        # Columnas adicionales
        for col in ["Descuento", "Motivo del descuento", "Comentarios"]:
            if col not in remittance.columns:
                remittance[col] = ""

        conds = [
            remittance["Tipo de Documento"].str.startswith("Facturas por Cobrar", na=False),
        ]

        descuentos = ["Descuento cliente"]
        motivos = ["657"]

        remittance["Descuento"] = np.select(conds, descuentos, default=remittance["Descuento"])
        remittance["Motivo del descuento"] = np.select(conds, motivos, default=remittance["Motivo del descuento"])

        # Comentarios
        remittance.loc[
            remittance["Motivo del descuento"].notna() &
            (remittance["Motivo del descuento"].str.strip() != ""),
            "Comentarios"
        ] = remittance["Referencia / Factura"].astype(str)

        # Tipo de doc segun lógica
        remittance["Tipo de Documento"] = np.select(
            [
                remittance["Motivo del descuento"].notna() &
                (remittance["Motivo del descuento"].str.strip() != ""),
                remittance["Importe de Remittance"] < 0,
                remittance["Importe de Remittance"] >= 0
            ],
            [
                "Descuento cliente",
                "Nota de crédito",
                "Factura"
            ],
            default="Factura"
        )

        # Procesos adicionales
        FBL5N, id_cliente, nombre_cliente = procesar_cartera_cliente(archivo_fbl5n, customer_id)
        hrc_template = merge_remittance_cartera(remittance, FBL5N)
        hrc_template = procesar_diferencias(hrc_template)
        hrc_template = procesamiento_nro(hrc_template, FBL5N)

        # Columnas finales
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

        # Remittance temporal
        ruta_remittance_excel = os.path.join(os.path.dirname(archivos_remittance), "remittance_temp.xlsx")
        remittance.to_excel(ruta_remittance_excel, index=False)

        # Export final
        exportar_template(
            hrc_template=hrc_template,
            suma_remittance=remittance["Importe de Remittance"].sum(),
            numero_orden="",
            id_cliente=id_cliente,
            nombre_cliente=nombre_cliente,
            ruta_remittance=ruta_remittance_excel,
            ruta_salida=rutas["salida"]
        )

        if os.path.exists(ruta_remittance_excel):
            os.remove(ruta_remittance_excel)

        messagebox.showinfo("Éxito", f"Archivo exportado exitosamente como:\n{rutas['salida']}")

        return hrc_template

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error:\n{e}")
        raise e
