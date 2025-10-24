# =====================================================
# 0. Importación de librerías y módulos utilitarios
# =====================================================

import os       # Manejo de rutas y directorios del sistema operativo
import sys      # Detección de ejecución empaquetada y manipulación de rutas del intérprete
import re       # Expresiones regulares (búsqueda y limpieza de texto)
import io        # Manejo de flujos de datos en memoria (buffers, streams)
import warnings # Control de advertencias del sistema y librerías externas

import numpy as np  # Operaciones numéricas y lógicas (np.where, np.select, etc.)
import pandas as pd  # Manipulación y análisis de datos tabulares
import camelot   # Extracción de tablas desde archivos PDF
from PyPDF2 import PdfReader  # Lectura y procesamiento de archivos PDF
from openpyxl import load_workbook  # Lectura de archivos Excel (.xlsx)

# Buscamos las funciones en la carpeta Main
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.abspath(os.path.join(current_dir, "..", "..", ".."))  # Sube desde /Pruebas/clientes/Colombia a /Raiz
sys.path.append(project_root)
from Main.utils import *  # Importación de funciones utilitarias del paquete interno 'clientes'

# Configuración de advertencias
warnings.filterwarnings("ignore", category=UserWarning, module="camelot") # Suprime advertencias generadas por Camelot (usualmente por manejo de PDFs)

# =====================================================
# 1. Localización dinámica de la carpeta raíz del proyecto
# =====================================================
def _project_root():
    """
    Obtiene la ruta base del proyecto sin importar el entorno de ejecución.

    - Si el código se ejecuta empaquetado (por ejemplo, como .app o .exe),
      sube desde la ruta del ejecutable hasta la carpeta que contiene el proyecto.
    - Si se ejecuta como script Python normal, sube dos niveles desde
      el archivo actual (../..), asumiendo la estructura estándar del proyecto.

    Devuelve:
        str: Ruta absoluta a la carpeta raíz del proyecto.
    """
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
    """
    Orquestador principal para Cencosud
    """
    root = _project_root()

    # --- Rutas de entrada/salida ---
    rutas = {
        "pdf_remittance": os.path.join(root, "Archivos", "Remittance", "Ecuador", "Remittance_Mico.pdf"),
#        "fbl5n": os.path.join(root, "Archivos", "Cartera", "FBL5N.xlsx"),
#        "fbl5n": os.path.join(root, "Archivos", "Cartera", "FBL5N_Mico.xlsx"),
        "salida": os.path.join(root, "Archivos", "Template", "Ecuador", "Template_HRC_Mico.xlsx")
    }
    # Colocar el Customer ID del cliente
    customer_id = #FALTA

    # =====================================================
    # 1. Lectura de Remitente
    # =====================================================
    # Leemos la tabla principal del Remittance: (ajustado a PDF) - usando Camelot
    
    # --- Expresión regular robusta ---
    pattern = re.compile(
        r"([A-Z0-9\-]+)\s+"             # Doc/Factura
        r"(\d+)\s+"                     # No.Documento
        r"(\d{2}\.\d{2}\.\d{4})\s+"     # Fc.Documento
        r"(\d{2}\.\d{2}\.\d{4})\s+"     # Fc.Contabiliz.
        r"([\d.,]+-?)\s+"               # Importe (puede terminar en '-')
        r"([A-Z]{3})"                   # Moneda
    )

    rows = []

    with fitz.open(pdf_path) as doc:
        for page_num, page in enumerate(doc, start=1):
            text = page.get_text("text")
            matches = pattern.findall(text)
            print(f"📄 Página {page_num} — {len(matches)} filas detectadas")
            for m in matches:
                rows.append(m)

    # --- Crear DataFrame ---
    df = pd.DataFrame(
        rows,
        columns=[
            "Doc/Factura",
            "No.Documento",
            "Fc.Documento",
            "Fc.Contabiliz.",
            "Importe",
            "Moneda",
        ],
    )

    # --- Limpieza de importes (siempre positivos) ---
    def parse_importe(val):
        """
        Convierte valores tipo '27,024.63-' o '895.76-' en floats POSITIVOS (sin signo negativo).
        """
        if not isinstance(val, str):
            return None
        val = val.strip().replace("-", "")  # Quitamos el signo

        # Si hay ',' y '.', asumimos formato tipo '27,024.63' → quitamos ',' (miles)
        if "," in val and "." in val:
            val = val.replace(",", "")
        # Si hay solo ',', asumimos que ',' son decimales (formato europeo)
        elif "," in val and "." not in val:
            val = val.replace(",", ".")

        try:
            return float(val)
        except ValueError:
            return None

    df["Importe"] = df["Importe"].apply(parse_importe)

    return df.head(100)