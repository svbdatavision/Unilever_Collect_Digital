# =====================================================
# 0. Importación de librerías y módulos utilitarios
# =====================================================
import sys
import os
import pandas as pd
import numpy as np
from openpyxl import load_workbook
import io

import pytesseract
from PIL import Image
import re

from clientes.utils import *  # Importación de funciones utilitarias del paquete interno 'clientes'

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
        "remittance": os.path.join(root, "Archivos", "Remittance", "Colombia", "Remittance_Cruz_Verde.png"),
        "fbl5n": os.path.join(root, "Archivos", "Cartera", "FBL5N_CRUZ_VERDE.xlsx"),
        "salida": os.path.join(root, "Archivos", "Template", "Colombia", "Template_HRC_Cruz_Verde.xlsx")
    }

    # Colocar el Customer ID del cliente
    customer_id = 10554302

    # =====================================================
    # 3. Lectura de Remittance desde imagen (OCR)
    # =====================================================
    image_path = rutas["remittance"]
    img = Image.open(image_path)

    raw_text = pytesseract.image_to_string(img, lang="spa")
    lines = raw_text.split("\n")

    facturas = []
    importes = []

    # Regex para capturar número de factura
    regex_fact = r"\b(PMP\d+|NCMI\d+)\b"

    # Regex para capturar importes
    regex_importe = r"\(?\d{1,3}(?:,\d{3})+\)?"

    for line in lines:
        mf = re.search(regex_fact, line)
        mi = re.search(regex_importe, line)

        if mf and mi:
            facturas.append(mf.group(0))
            importes.append(mi.group(0))

    # ---------- Limpieza de importes ----------
    def limpiar_importe(x):
        x = x.replace(",", "")
        if x.startswith("(") and x.endswith(")"):
            return -int(x.strip("()"))
        return int(x)

    importes_limpios = [limpiar_importe(i) for i in importes]

    # ---------- DataFrame remittance formato estándar ----------
    remittance = pd.DataFrame({
        "Referencia / Factura": facturas,
        "Importe de Remittance": importes_limpios
    })
    
    # 2.1 Guardar Remittance en buffer en memoria
    remittance_buffer = io.BytesIO()
    with pd.ExcelWriter(remittance_buffer, engine="openpyxl") as writer:
        remittance.to_excel(writer, index=False)
    remittance_buffer.seek(0)

    # Asignación condicional del Tipo de Documento
    remittance["Tipo de Documento"] = np.select(
        [
            remittance["Referencia / Factura"].str.startswith("PMP", na=False),
            remittance["Referencia / Factura"].str.startswith("NCMI", na=False)
        ],
        [
            "Factura",
            "Descuentos Cliente"
        ],
        default=""
    )
    
    # Columnas requeridas
    remittance["Descuento"] = ""
    remittance["Motivo del descuento"] = ""

    # Normalización simple (igual que tu proceso original)
    remittance["Referencia / Factura"] = (
        remittance["Referencia / Factura"]
        .astype(str)
    )

    remittance["Importe de Remittance"] = remittance["Importe de Remittance"].astype(float)

    # =====================================================
    # 5. Manejo de Reglas y CARDs
    # =====================================================
    """
    conds = [
        (remittance["Referencia / Factura"].str.startswith("PMP", na=False)) & (remittance["Importe de Remittance"] < 0),
        remittance["Referencia / Factura"].str.startswith("0085", na=False),
        remittance["Referencia / Factura"].str.startswith("46", na=False),
        remittance["Referencia / Factura"].str.startswith("98", na=False),
        remittance["Referencia / Factura"].str.startswith("310", na=False)
    ]

    descuentos = ["RECHAZO", "DESCUENTO", "AVERIA", "FACT PROVEEDOR", "NOTA DEBITO"]
    motivos = ["551", "987", "522", "CSB", "987"]

    remittance["Descuento"] = np.select(conds, descuentos, default=remittance["Descuento"])
    remittance["Motivo del descuento"] = np.select(conds, motivos, default=remittance["Motivo del descuento"])
    """
    # =====================================================
    # 6. Procesamiento de columnas 'Descuento' y 'Comentarios'
    # =====================================================
    remittance = procesar_descuentos_y_comentarios(remittance)

    # =====================================================
    # 7–9. Procesamiento de Cartera FBL5N
    # =====================================================
    FBL5N, id_cliente, nombre_cliente = procesar_cartera_cliente(rutas["fbl5n"], customer_id)

    # =====================================================
    # 10. Merge Remittance + FBL5N por factura
    # =====================================================
    hrc_template = merge_remittance_cartera(remittance, FBL5N)

    # =====================================================
    # 11. Cálculo de diferencias
    # =====================================================
    hrc_template = procesar_diferencias(hrc_template)

    # =====================================================
    # 12. Agregar registros NRO
    # =====================================================
    hrc_template = procesamiento_nro(hrc_template, FBL5N)

    # =====================================================
    # 13. Pago Neto
    # =====================================================
    hrc_template["Pago Neto"] = hrc_template["Importe de factura"]

    # =====================================================
    # 14. Columnas finales
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
    # 15. Preparación para exportar (los valores se tomarán del Excel original)
    # =====================================================
    # Como ya no hay un Excel de Remittance, el sistema tomará referencias dummy
    # Ajusta esto si deseas extraer datos adicionales desde la imagen o un JSON.
    suma_remittance = remittance["Importe de Remittance"].sum()
    numero_orden = ""

    exportar_template(
        hrc_template=hrc_template,
        suma_remittance=suma_remittance,
        numero_orden=numero_orden,
        id_cliente=id_cliente,
        nombre_cliente=nombre_cliente,
        ruta_remittance=remittance_buffer,
        ruta_salida=rutas["salida"]
    )

    return hrc_template
