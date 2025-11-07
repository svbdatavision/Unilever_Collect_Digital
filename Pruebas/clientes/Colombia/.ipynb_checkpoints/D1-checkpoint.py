# =====================================================
# 0. Importación de librerías y módulos utilitarios
# =====================================================
import sys  # Para detectar ejecución empaquetada (frozen) y resolver rutas
import os   # Para construir rutas relativas al proyecto
import re
import io
import pandas as pd
import numpy as np
import fitz  # PyMuPDF, para leer PDF
from openpyxl import load_workbook
# Buscamos las funciones en la carpeta Main
# current_dir = os.path.dirname(os.path.abspath(__file__))
# project_root = os.path.abspath(os.path.join(current_dir, "..", "..", ".."))  # Sube desde /Pruebas/clientes/Colombia a /Raiz
# sys.path.append(project_root)
# from Main.utils import *  # Importación de funciones utilitarias del paquete interno 'clientes'

from clientes.utils import *  # Importación de funciones utilitarias del paquete interno 'clientes'



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
        "remittance_pdf": os.path.join(root, "Archivos", "Remittance", "Colombia", "Remittance_D1.pdf"),
#        "fbl5n": os.path.join(root, "Archivos", "Cartera", "FBL5N.xlsx"),
        "fbl5n": os.path.join(root, "Archivos", "Cartera", "FBL5N_d1.xlsx"),
        "salida": os.path.join(root, "Archivos", "Template", "Colombia", "Template_HRC_D1.xlsx")
    }

    # =====================================================
    # 1. Lectura de Remitente
    # =====================================================
    # Leemos la tabla principal del Remittance: (ajustado a PDF) 
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
    remittance["Importe de Remittance"] = remittance["Neto Pagado"].astype(float)
    remittance["Tipo de Documento"] = "Factura"  

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
    remittance_buffer.seek(0) 

    # =====================================================
    # 3. Diferenciamos Facturas de diferencias (CARDs / Reglas)
    # =====================================================
    conds = [
        remittance["Referencia / Factura"].str.startswith("PMP", na=False) & (remittance["Importe de Remittance"] < 0),
        # FALTA: otras reglas específicas de D1 si aplica
    ]
    descuentos = ["RECHAZO"]  # FALTA: ajustar según reglas D1
    motivos = ["551"]          # FALTA: ajustar según reglas D1

    remittance["Descuento"] = np.select(conds, descuentos, default="Descuento")
    remittance["Motivo del descuento"] = np.select(conds, motivos, default="")
    
    # =====================================================
    # 4. Procesamiento de columnas 'Descuento' y 'Comentarios'
    # =====================================================
    
    # No va.

    # =====================================================
    # 4. Lectura de la Cartera (FBL5N) (datos desde SAP)
    # =====================================================
    # Se leen solo las columnas necesarias, todo como texto (dtype=str) para evitar errores de tipo
    # 'engine="openpyxl"' es el más estable y rápido para archivos .xlsx
    FBL5N = pd.read_excel(
        rutas["fbl5n"],
        usecols=[
            "Document Type",
            "Reference",
            "Amount in local currency",
            "Reason code",
            "Name 1"
        ],
        dtype=str,
        engine="openpyxl"
    )
    
    # =====================================================
    # 5. Filtro de la cartera del cliente (traer solo RV / facturas relevantes)
    # =====================================================
    # Se conservan únicamente las filas donde:
    #   - "Document Type" == "RV" (facturas)
    #   - O "Reason code" == "NRO" (casos especiales)
    #   - Y además el campo "Name 1" contenga "SUPERTIENDAS Y DROGUERIAS OLIM"
    FBL5N = FBL5N[
        ((FBL5N["Document Type"] == "RV") | (FBL5N["Reason code"] == "NRO"))
        & (FBL5N["Name 1"].str.contains("D1 S A S", case=False, na=False))
    ].reset_index(drop=True)
    
    # =====================================================
    # 6. Renombrado de columnas 
    # =====================================================
    # Se renombran las columnas clave para mayor claridad y consistencia con el resto del proceso.
    FBL5N = FBL5N.rename(columns={
        "Reference": "Referencia / Factura",
        "Amount in local currency": "Importe de factura"
    })
    
    # Crear una máscara booleana que detecte valores entre paréntesis (formato contable negativo)
    mask_negativo = FBL5N["Importe de factura"].str.contains(r"\(", regex=True)

    # Limpiar y convertir los valores
    FBL5N["Importe de factura"] = (
        FBL5N["Importe de factura"]
        .str.replace(",", "", regex=False)       # eliminar separadores de miles
        .str.replace(r"[\(\)]", "", regex=True)  # eliminar paréntesis
        .astype(float)                           # convertir a float
        * mask_negativo.map(lambda x: -1 if x else 1)  # aplicar signo negativo
    )
    
    # =====================================================
    # 7. Merge entre Remittance y FBL5N por "Referencia / Factura"
    #    8.1 how="left" -> conservar todo Remittance y añadir datos de FBL5N cuando coincidan
    #    8.2 Eliminamos referencias especiales (ej: CNQ*) para evitar falsos positivos
    # =====================================================
    hrc_template = pd.merge(remittance, FBL5N, on="Referencia / Factura", how="left")
#    hrc_template = merge_remittance_cartera(remittance, FBL5N)

    # =====================================================
    # 8. Calculamos Diferencias entre Importe de factura - Importe de Remittance
    #    (logica centralizada en procesar_diferencias)
    # =====================================================
    hrc_template = procesar_diferencias(hrc_template)

    # =====================================================
    # 9. Agregamos registros NRO
    # =====================================================
    
    # =====================================================
    # 10. Dato de Pago Neto = Importe de factura y otros ajustes
    # =====================================================
    

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
    # 12. Generación de parámetros de entrada para la función exportar_template
    # =====================================================
    # FALTA: Extraer numero_orden, id_cliente, nombre_cliente dinámicamente
    numero_orden = ""
    id_cliente = ""
    nombre_cliente = ""

    exportar_template(
        hrc_template=hrc_template,
        suma_remittance = remittance["Importe de Remittance"].sum(),
        numero_orden=numero_orden,
        id_cliente=id_cliente,
        nombre_cliente=nombre_cliente,
        ruta_remittance=remittance_buffer,
        ruta_salida=rutas["salida"]
    )

    return hrc_template
