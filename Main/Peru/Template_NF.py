import pdfplumber
import pandas as pd
import re
import os
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side
from tkinter import messagebox
from utils import *
import config 


customer_id = 10449763

def procesar(archivos_remittance, archivo_fbl5n):
    try:    
        tablas_filtradas = []
        rutas = {
            "remittance": archivos_remittance,
            "fbl5n": archivo_fbl5n,
            "salida": os.path.join(os.path.dirname(archivos_remittance), "Nortfarma.xlsx")
        }
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

        # Eliminar columna 'Fec Emision'
        df = df.drop(columns=["Fec Emision"])


        # Renombrar columnas
        remittance = df.rename(columns={
            "Tipo Doc": "Tipo de Documento",
            "Num Documento": "Referencia / Factura",
            "Monto": "Importe de Remittance"
        })

        # Convertir 'Importe de Remittance' a numérico (float)
        remittance["Importe de Remittance"] = (
            remittance["Importe de Remittance"]
            .replace({',': '', ' ': ''}, regex=True)  # Elimina comas y espacios
            .replace('', np.nan)                      # Convierte cadenas vacías en NaN
            .astype(float)                            # Convierte a float
        )


        # Ajustar formato de 'Referencia / Factura' para que tenga 8 dígitos después del guion
        remittance["Referencia / Factura"] = remittance["Referencia / Factura"].apply(
            lambda x: f"{x.split('-')[0]}-{x.split('-')[1].zfill(8)}" if isinstance(x, str) and '-' in x else x
        )

        # Crear columnas adicionales si no existen
        for col in ["Descuento", "Motivo del descuento", "Comentarios"]:
            if col not in remittance.columns:
                remittance[col] = ""

        # Condiciones y motivos
        conds = [
            remittance["Tipo de Documento"].str.startswith("FACTURA X COBRAR", na=True)
        ]
        descuentos = ["Descuento cliente"]
        motivos = ["987"]

        remittance["Descuento"] = np.select(conds, descuentos, default=remittance["Descuento"])
        remittance["Motivo del descuento"] = np.select(conds, motivos, default=remittance["Motivo del descuento"])

        # Crear columna Comentarios si hay motivo de descuento
        remittance.loc[
            remittance["Motivo del descuento"].notna() & (remittance["Motivo del descuento"].str.strip() != ""),
            "Comentarios"
        ] = remittance["Tipo de Documento"].str.replace(" ", "") + " " + remittance["Referencia / Factura"].astype(str)

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
        nombre_cliente = "NORTFARMA S A C"

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

