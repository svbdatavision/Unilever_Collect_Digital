# =====================================================
# 0. Importación de librerías y módulos utilitarios
# =====================================================
import sys
import os
import pandas as pd
import numpy as np
from openpyxl import load_workbook
import re
import pdfplumber
import io

from utils import *  # Importación de funciones utilitarias del paquete interno 'clientes'

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
def procesar(archivo_remittance,archivo_fbl5n):
    root = _project_root()

    rutas = {
        "remittance": archivo_remittance,
        "fbl5n": archivo_fbl5n,
        # Si necesitas una ruta de salida, puedes definirla aquí:
        "salida": os.path.join(os.path.dirname(archivo_remittance), "Jeronimo.xlsx")
    }
    customer_id = 10851239 # Colocar el Customer ID del cliente

    # =====================================================
    # 3. Lectura de Remittance
    # =====================================================
    
    # Documento al inicio de la línea (alfanumérico largo)
    pat_doc = re.compile(r"^(?:\s*)([A-Z0-9]{6,})\b")

    # Importe al final de la línea (con puntos de miles)
    pat_amount = re.compile(r"(\d{1,3}(?:\.\d{3})+|\d+)\s*$")

    rows = []
    with pdfplumber.open(rutas["remittance"]) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for raw in text.split("\n"):
                line = raw.strip()
                if not line:
                    continue
                # Saltar encabezado
                if line.lower().startswith("n. documento"):
                    continue

                mdoc = pat_doc.match(line)
                mamt = pat_amount.search(line)
                if mdoc and mamt:
                    doc = mdoc.group(1)
                    amount = mamt.group(1)

                    # ================================
                    # Opción A: excluir documentos sin dígitos (p. ej. 'CITIBANK')
                    # ================================
                    if not any(ch.isdigit() for ch in doc):
                        continue

                    rows.append({"N. Documento": doc, "Importe de pago": amount})

    remittance = pd.DataFrame(rows)
    
   # 2.1 Guardar Remittance en buffer en memoria
    remittance_buffer = io.BytesIO()
    with pd.ExcelWriter(remittance_buffer, engine="openpyxl") as writer:
        remittance.to_excel(writer, index=False)
    remittance_buffer.seek(0)

    # =====================================================
    # 4. Limpieza de Remittance
    # =====================================================
    
    remittance["Importe de pago"] = (
            remittance["Importe de pago"]
            .astype(str)
            .str.replace(".", "", regex=False)
            .astype(np.int64)
        )

    remittance["Tipo de Documento"] = np.where(
        remittance["N. Documento"].str.upper().str.startswith(("PMP", "NCM")),
        "Factura",
        "Descuentos Clientes"
    )

    # Renombrar columnas
    remittance.rename(columns={
        "N. Documento": "Referencia / Factura",
        "Importe de pago": "Importe de Remittance"
    }, inplace=True)

    # =====================================================
    # 5. Manejo de Reglas y CARDs
    # =====================================================  
    
    # Inicializar columnas si no existen
    for col in ["Descuento", "Motivo del descuento", "Comentarios"]:
        if col not in remittance.columns:
            remittance[col] = ""
 
    # Regla única:
    # Tipo de Documento = Descuentos Clientes
    mask = remittance["Tipo de Documento"] == "Descuentos Clientes"
 
    remittance.loc[mask, "Descuento"] = remittance.loc[mask, "Referencia / Factura"]
    remittance.loc[mask, "Motivo del descuento"] = "987"
 
    # (Opcional) Comentarios
    remittance.loc[mask, "Comentarios"] = (
        "DESCUENTO CLIENTE " + remittance.loc[mask, "Referencia / Factura"]
    )
    
    
    # =====================================================
    # 6. Procesamiento de columnas 'Descuento' y 'Comentarios'
    # =====================================================
    
    
    remittance = procesar_descuentos_y_comentarios(remittance)


    # =====================================================
    # 7. Lectura de la Cartera (FBL5N) (datos desde SAP)
    # =====================================================
    # =====================================================
    # 8. Filtro de la cartera del cliente
    # =====================================================
    # =====================================================
    # 9. Renombrado y limpieza de columnas
    # =====================================================
    FBL5N, id_cliente, nombre_cliente = procesar_cartera_cliente(rutas["fbl5n"], customer_id)

    
    # =====================================================
    # 10. Merge Remittance + FBL5N por "Referencia / Factura"
    # =====================================================
    hrc_template = merge_remittance_cartera(remittance, FBL5N)
    
    # =====================================================
    # 11. Cálculo de diferencias
    # =====================================================
    hrc_template = procesar_diferencias(hrc_template)

    # =====================================================
    # 12. Agregamos registros NRO
    # =====================================================
    hrc_template = procesamiento_nro(hrc_template, FBL5N)

    # =====================================================
    # 13. Asignación de Pago Neto (Pago Neto = Importe de factura) y otros ajustes
    # =====================================================
    hrc_template["Pago Neto"] = hrc_template["Importe de factura"]
    
    # =====================================================
    # 14. Definición de columnas finales para el template
    # =====================================================
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

    # =====================================================
    # 15. Preparación de parámetros y extracción de datos dinámicos (para exportar_template)
    # =====================================================

    exportar_template(
        hrc_template=hrc_template,
        suma_remittance = remittance["Importe de Remittance"].sum(),
        numero_orden="",
        id_cliente=id_cliente,
        nombre_cliente=nombre_cliente,
        ruta_remittance=remittance_buffer,
        ruta_salida=rutas["salida"]
    )

    return hrc_template
    