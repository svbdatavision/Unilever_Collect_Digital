import fitz  # PyMuPDF
import pandas as pd
import re
import os
import sys
import numpy as np
from openpyxl import load_workbook
from clientes.utils import *
from tkinter import messagebox

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
        "remittance": os.path.join(root,"Archivos", "Remittance", "Peru", "Remittance_Tottus.pdf"), # Completar nombre del Remittance (pdf) a trabajar
        "fbl5n": os.path.join(root,"Archivos", "Cartera", "Peru", "FBL5N_Tottus.xlsx"), # Completar nombre de la cartera (excel) a trabajar
        "salida": os.path.join(root,"Archivos", "Template", "Peru", "Template_HRC_Tottus.xlsx") # Colocar el nombre de salida que deseen (Ej: Template_HRC_nombre_cliente.xlsx)
    }
    customer_id = 10299933 # Colocar el Customer ID del cliente

    # =====================================================
    # Lógica original del procesamiento (sin modificar)
    # =====================================================
    try:
        archivos_remittance = rutas["remittance"]
        archivo_fbl5n = rutas["fbl5n"]

        PREFIJOS_A_ELIMINAR = ("08-", "07-", "00-", "01-")
        TIPO_MAP = {
            "Fact.Elect. Af. Emi": "Factura por convenio",
            "Fac Elect Ex Emitida": "Factura por convenio",
            "Ndd Afecta Elec. Rec": "Nota de débito",
            "Fac Afecta Elec. Rec": "Factura",
            "Ncr Ex Elect. Rec": "Nota de crédito"
        }

        doc = fitz.open(archivos_remittance)
        tabla_documentos = []

        regex_fila = re.compile(
            r"(\d{2}\.\d{2}\.\d{4})\s+([A-Z0-9\-]+)\s+([\w\s\.\-]+?)\s+([\-]?\d{1,3}(?:[\.,]\d{3})*[\.,]\d{2})"
        )

        for page in doc:
            text = page.get_text()
            for match in regex_fila.findall(text):
                fecha, referencia, tipo_doc, monto_texto = match

                # Eliminar prefijos molestos
                for prefijo in PREFIJOS_A_ELIMINAR:
                    if referencia.startswith(prefijo):
                        referencia = referencia[len(prefijo):]
                        break

                descripcion = TIPO_MAP.get(tipo_doc.strip(), tipo_doc.strip())

                # Convertir monto a float (limpiando separadores)
                monto_texto = monto_texto.replace('.', '').replace(',', '.').replace(' ', '')
                monto = float(monto_texto.replace('-', ''))

                tabla_documentos.append({
                    "Tipo de Documento": descripcion,
                    "Referencia / Factura": referencia,
                    "Importe de Remittance": monto
                })

        # Crear DataFrame
        remittance = pd.DataFrame(tabla_documentos)

        remittance["Importe de Remittance"] = remittance.apply(
            lambda row: -abs(row["Importe de Remittance"]) if row["Tipo de Documento"] in ["Nota de crédito", "Factura por convenio"]
            else abs(row["Importe de Remittance"]),
            axis=1
        )

        # Columnas adicionales si no existen
        for col in ["Descuento", "Motivo del descuento", "Comentarios"]:
            if col not in remittance.columns:
                remittance[col] = ""

        # Condiciones y motivos
        conds = [
            remittance["Referencia / Factura"].str.startswith("FF", na=False)
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

        # Procesos adicionales con utils
        FBL5N, id_cliente, nombre_cliente = procesar_cartera_cliente(archivo_fbl5n, customer_id)
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
