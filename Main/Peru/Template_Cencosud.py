import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side
from tkinter import messagebox
import numpy as np
from utils import *


customer_id = 10262842


def procesar(archivos_remittance, archivo_fbl5n):
    try: 
    
        rutas = {
            "remittance": archivos_remittance,
            "fbl5n": archivo_fbl5n,
            "salida": os.path.join(os.path.dirname(archivos_remittance), "Cencosud.xlsx")
        }

        df = pd.read_excel(archivos_remittance, sheet_name='dataTableAbonosAnterioresDetall', engine='openpyxl')

        # Seleccionar las columnas deseadas
        columnas_deseadas = ['Nro. Documento', 'Tipo', 'Neto']
        remittance = df[columnas_deseadas]

        tipo_map = {
            "NOTA DE DEBITO PROVEEDOR": "Nota de Débito",
            "NOTA DE CREDITO PROVEEDOR": "Nota de Crédito",
            "FACTURA PROVEEDOR": "Factura"
        }


        # Renombrar columnas
        remittance = remittance.rename(columns={
            "Tipo": "Tipo de Documento",
            "Nro. Documento": "Referencia / Factura",
            "Neto": "Importe de Remittance"
        })

        # Elimina los prefijos en la columna 'Comprobante proveedor'

        remittance['Referencia / Factura'] = (
            remittance['Referencia / Factura']
            .astype(str)
            .str.replace(r'^(01-|07-|00-)', '', regex=True)
        )

        for col in ["Descuento", "Motivo del descuento", "Comentarios"]:
            if col not in remittance.columns:
                remittance[col] = ""

        # Condiciones y motivos
        conds = [
            remittance["Referencia / Factura"].str.startswith(("FA", "FN"), na=False),
        ]

        descuentos = ["Descuento cliente"]
        motivos = ["657"]

        remittance["Descuento"] = np.select(conds, descuentos, default=remittance["Descuento"])
        remittance["Motivo del descuento"] = np.select(conds, motivos, default=remittance["Motivo del descuento"])

        # Crear columna Comentarios si hay motivo de descuento
        remittance.loc[
            remittance["Motivo del descuento"].notna() & (remittance["Motivo del descuento"].str.strip() != ""),
            "Comentarios"
        ] = remittance["Referencia / Factura"].astype(str)

        # Determinar tipo de documento según valor y motivo de descuento
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
            default="Factura"  # ← Asegúrate de que esto también sea str
        )

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
        nombre_cliente = "CENCOSUD RETAIL PERU S A"

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
        messagebox.showinfo("Exitoso",f"✅ Archivo exportado exitosamente como {rutas['salida']}")
        return hrc_template

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error: {e}")
        raise e