import sys  # Para detectar ejecución empaquetada (frozen) y resolver rutas
import os   # Para construir rutas relativas al proyecto
import re
import io
import pandas as pd
import numpy as np
import fitz  # PyMuPDF, para leer PDF
from openpyxl import load_workbook
from clientes.utils.formato_template import exportar_template
from clientes.utils.diferencias import procesar_diferencias


def _project_root():
    """
    Devuelve la carpeta raíz del proyecto:
    - Si corre dentro de un .app (PyInstaller / py2app) -> carpeta que contiene el .app
    - Si corre como script -> la carpeta del archivo actual (../)
    """
    if getattr(sys, "frozen", False):
        macos_dir = os.path.dirname(sys.executable)
        contents_dir = os.path.dirname(macos_dir)
        app_bundle = os.path.dirname(contents_dir)
        return os.path.dirname(app_bundle)
    return os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))


def procesar():
    """
    Proceso general para Remittance + FBL5N
    Adaptable a cualquier cliente.
    Incluye comentarios específicos para D1.
    """
    root = _project_root()

    # --- Rutas ---
    rutas = {
        "remittance_pdf": os.path.join(root, "Archivos", "Remittance", "Remittance_D1.pdf"),  # D1
        "fbl5n": os.path.join(root, "Archivos", "Base_de_datos", "FBL5N_d1.xlsx"),  # D1
        "salida": os.path.join(root, "Archivos", "Template", "Template_HRC_D1.xlsx")  # D1
    }

    # =====================================================
    # 1. Lectura de Remitente
    # =====================================================
    # Para D1: Remittance en PDF
    factura_pattern = re.compile(
        r"RE\s+\d+\s+PMP\d+\s+\d{1,3}(?:\.\d{3})*\s+\d+\s+\d{1,3}(?:\.\d{3})*\s+\d+\s+\d{1,3}(?:\.\d{3})*"
    )
    facturas = []
    with fitz.open(rutas["remittance_pdf"]) as doc:
        for page in doc:
            text = page.get_text()
            for match in factura_pattern.findall(text):
                parts = match.split()
                if len(parts) == 8:
                    facturas.append({
                        "TD": parts[0],
                        "Doc.Interno": parts[1],
                        "Referencia / Factura": parts[2],
                        "Valor Bruto": parts[3].replace('.', ''),
                        "Retenciones": parts[4],
                        "IVA": parts[5].replace('.', ''),
                        "Desc/Rec": parts[6],
                        "Neto Pagado": parts[7].replace('.', '')
                    })

    remittance = pd.DataFrame(facturas)
    remittance["Importe de factura"] = remittance["Neto Pagado"].astype(float)
    remittance["Tipo de Documento"] = "Factura"  # D1: todas las facturas son tipo "Factura"

    # =====================================================
    # 2. Limpieza de Remittance
    # =====================================================
    # Limpiar referencias, caracteres invisibles y espacios
    remittance["Referencia / Factura"] = remittance["Referencia / Factura"].str.replace(
        r"[\u202A-\u202E\u200E\u200F]", "", regex=True
    ).str.strip()

    # 2.1 Guardar Remittance en buffer en memoria
    # - Crear un BytesIO con la tabla remittance lista para exportar
    remittance_buffer = io.BytesIO()
    with pd.ExcelWriter(remittance_buffer, engine="openpyxl") as writer:
        remittance.to_excel(writer, index=False)
    remittance_buffer.seek(0)  # rebobinar el buffer

    # =====================================================
    # 3. Diferenciamos Facturas de diferencias (CARDs / Reglas)
    # =====================================================
    # D1: ejemplo simplificado de reglas
    conds = [
        remittance["Referencia / Factura"].str.startswith("PMP", na=False) & (remittance["Importe de factura"] < 0),
        # FALTA: otras reglas específicas de D1 si aplica
    ]
    descuentos = ["RECHAZO"]  # FALTA: ajustar según reglas D1
    motivos = ["551"]          # FALTA: ajustar según reglas D1

    remittance["Descuento"] = np.select(conds, descuentos, default="Descuento")
    remittance["Motivo del descuento"] = np.select(conds, motivos, default="")

    # =====================================================
    # 4. Lectura de FBL5N
    # =====================================================
    FBL5N = pd.read_excel(
        rutas["fbl5n"],
        usecols=["Document Type", "Reference", "Amount in local currency", "Reason code", "Document Number", "Text"]
    )

    # =====================================================
    # 5. Filtro FBL5N
    # =====================================================
    FBL5N = FBL5N[(FBL5N["Document Type"] == "RV") | (FBL5N["Reason code"] == "NRO")]

    # =====================================================
    # 6. Rename de columnas
    # =====================================================
    FBL5N = FBL5N.rename(columns={
        "Reference": "Referencia / Factura",
        "Amount in local currency": "importe_FBL5N"
    }).reset_index(drop=True)

    # =====================================================
    # 7. Trabajamos con los NRO
    # =====================================================
    # FALTA: lógica completa de NRO para D1 si aplica

    # =====================================================
    # 8. Merge por "Referencia / Factura"
    # =====================================================
    hrc_template = pd.merge(remittance, FBL5N, on="Referencia / Factura", how="left")

    # =====================================================
    # 9. Calculamos Diferencias
    # =====================================================
    hrc_template = procesar_diferencias(hrc_template)  # lógica centralizada

    # =====================================================
    # 10. Ajuste de Comentarios y Pago Neto
    # =====================================================
    hrc_template["Comentarios"] = np.where(
        hrc_template["Tipo de Documento"] == "Factura", "",
        np.where(
            hrc_template["Descuento"] == "MENORES VALORES",
            hrc_template["Descuento"],
            hrc_template["Referencia / Factura"].fillna("")
        )
    )
    hrc_template["Pago Neto"] = hrc_template["Importe de factura"]

    # =====================================================
    # 11. Definimos columnas finales
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
    # 12. Exportar a Template
    # =====================================================
    # FALTA: Extraer numero_orden, id_cliente, nombre_cliente dinámicamente
    numero_orden = ""  # FALTA
    id_cliente = ""    # FALTA
    nombre_cliente = ""  # FALTA

    exportar_template(
        hrc_template=hrc_template,
        numero_orden=numero_orden,
        id_cliente=id_cliente,
        nombre_cliente=nombre_cliente,
        ruta_remittance=remittance_buffer,  # PDF en memoria
        ruta_salida=rutas["salida"]
    )

    return hrc_template
