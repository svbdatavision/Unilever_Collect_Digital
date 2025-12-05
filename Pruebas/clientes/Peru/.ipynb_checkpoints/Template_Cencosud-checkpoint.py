# =====================================================
# 0. Importación de librerías y módulos utilitarios
# =====================================================
import sys
import os
import pandas as pd
import numpy as np
from tkinter import messagebox
from openpyxl import load_workbook

from clientes.utils import *   # Importa tus funciones utilitarias internas

# =====================================================
# 1. Localización dinámica de la carpeta raíz del proyecto
# =====================================================
def _project_root():
    if getattr(sys, "frozen", False):
        macos_dir = os.path.dirname(sys.executable)
        contents_dir = os.path.dirname(macos_dir)
        app_bundle = os.path.dirname(contents_dir)
        return os.path.dirname(app_bundle)

    return os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))


# =====================================================
# 2. Función principal del proceso (procesar)
# =====================================================
def procesar():

    root = _project_root()

    rutas = {
        "remittance": os.path.join(root, "Archivos", "Remittance", "Colombia", "Remittance.xlsx"),
        "fbl5n": os.path.join(root, "Archivos", "Cartera", "FBL5N.xlsx"),
        "salida": os.path.join(root, "Archivos", "Template", "Colombia", "Cencosud.xlsx")
    }

    # Manteniendo tu Customer ID original
    customer_id = 10262842

    try:
        # =====================================================
        # 3. Lectura de Remittance
        # =====================================================
        df = pd.read_excel(
            rutas["remittance"],
            sheet_name='dataTableAbonosAnterioresDetall',
            engine='openpyxl'
        )

        columnas_deseadas = ['Nro. Documento', 'Tipo', 'Neto']
        remittance = df[columnas_deseadas]

        tipo_map = {
            "NOTA DE DEBITO PROVEEDOR": "Nota de Débito",
            "NOTA DE CREDITO PROVEEDOR": "Nota de Crédito",
            "FACTURA PROVEEDOR": "Factura"
        }

        remittance = remittance.rename(columns={
            "Tipo": "Tipo de Documento",
            "Nro. Documento": "Referencia / Factura",
            "Neto": "Importe de Remittance"
        })

        remittance['Referencia / Factura'] = (
            remittance['Referencia / Factura']
            .astype(str)
            .str.replace(r'^(01-|07-|00-)', '', regex=True)
        )

        for col in ["Descuento", "Motivo del descuento", "Comentarios"]:
            if col not in remittance.columns:
                remittance[col] = ""

        conds = [
            remittance["Referencia / Factura"].str.startswith(("FA", "FN"), na=False),
        ]
        descuentos = ["Descuento cliente"]
        motivos = ["657"]

        remittance["Descuento"] = np.select(conds, descuentos, default=remittance["Descuento"])
        remittance["Motivo del descuento"] = np.select(conds, motivos, default=remittance["Motivo del descuento"])

        remittance.loc[
            remittance["Motivo del descuento"].notna() & (remittance["Motivo del descuento"].str.strip() != ""),
            "Comentarios"
        ] = remittance["Referencia / Factura"].astype(str)

        remittance["Tipo de Documento"] = np.select(
            [
                remittance["Motivo del descuento"].notna() & (remittance["Motivo del descuento"].str.strip() != ""),
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

        # =====================================================
        # 4. Lectura y procesado de FBL5N
        # =====================================================
        FBL5N, id_cliente, nombre_cliente = procesar_cartera_cliente(rutas["fbl5n"], customer_id)

        hrc_template = merge_remittance_cartera(remittance, FBL5N)
        hrc_template = procesar_diferencias(hrc_template)
        hrc_template = procesamiento_nro(hrc_template, FBL5N)

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

        numero_orden = ""

        ruta_remittance_excel = os.path.join(root, "Archivos", "Remittance", "Colombia", "remittance_temp.xlsx")
        remittance.to_excel(ruta_remittance_excel, index=False)

        exportar_template(
            hrc_template=hrc_template,
            suma_remittance=remittance["Importe de Remittance"].sum(),
            numero_orden=numero_orden,
            id_cliente=id_cliente,
            nombre_cliente=nombre_cliente,
            ruta_remittance=ruta_remittance_excel,
            ruta_salida=rutas["salida"]
        )

        if os.path.exists(ruta_remittance_excel):
            os.remove(ruta_remittance_excel)

        messagebox.showinfo("Exitoso", f"✅ Archivo exportado exitosamente como {rutas['salida']}")
        return hrc_template

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error: {e}")
        raise e
