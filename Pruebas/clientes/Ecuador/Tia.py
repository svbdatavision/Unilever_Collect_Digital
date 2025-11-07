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
# current_dir = os.path.dirname(os.path.abspath(__file__))
# project_root = os.path.abspath(os.path.join(current_dir, "..", "..", ".."))  # Sube desde /Pruebas/clientes/Colombia a /Raiz
# sys.path.append(project_root)
# from Main.utils import *  # Importación de funciones utilitarias del paquete interno 'clientes'
from clientes.utils import *  # Importación de funciones utilitarias del paquete interno 'clientes'

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
    root = _project_root()

    # --- Rutas de entrada/salida ---
    rutas = {
        "pdf_remittance": os.path.join(root, "Archivos", "Remittance", "Ecuador", "Remittance_TIA.pdf"),
#        "pdf_remittance": os.path.join(root, "Archivos", "Remittance", "Ecuador", "Remittance_TIA 2.pdf"),        
#        "fbl5n": os.path.join(root, "Archivos", "Cartera", "FBL5N.xlsx"),
#        "fbl5n": os.path.join(root, "Archivos", "Cartera", "FBL5N_TIA.xlsx"),
        "salida": os.path.join(root, "Archivos", "Template", "Ecuador", "Template_HRC_TIA.xlsx")
    }
    # Colocar el Customer ID del cliente
    customer_id = #FALTA
    
    # =====================================================
    # 1. Lectura de Remitente
    # =====================================================
    # Leemos la tabla principal del Remittance: (ajustado a PDF) - usando Camelot
    
    def extraer_facturas(archivo_pdf):
        filas = []

        with pdfplumber.open(rutas["pdf_remittance"]) as pdf:
            for pagina in pdf.pages:
                words = pagina.extract_words(x_tolerance=2, y_tolerance=3)
                lineas_dict = {}

                for w in words:
                    y = round(w["top"])
                    x = w.get("x", w.get("x0", 0))
                    lineas_dict.setdefault(y, []).append((x, w["text"]))

                for y, contenido in sorted(lineas_dict.items()):
                    linea = " ".join(t for _, t in sorted(contenido))
                    if re.match(r"^\d{4}-\d{2}-\d{2}\s+\d+", linea):
                        filas.append(linea)

        registros = []
        for f in filas:
            numeros = re.findall(r"-?[\d,]+\.\d{2}", f)
            if len(numeros) < 3:
                continue

            bruto, ret, neto = [n.replace(",", "") for n in numeros[-3:]]
            parte_antes = f[:f.rfind(bruto)].strip()
            tokens = parte_antes.split()

            # Fecha y código plaza
            fecha = tokens[0]
            plaza_codigo = tokens[1]

            # Buscar posición del F, C o D
            tipo_idx = None
            for i, tok in enumerate(tokens):
                if tok in ["F", "C", "D"]:
                    tipo_idx = i
                    break

            if tipo_idx is None:
                print("⚠️ No se encontró tipo en:", f)
                continue

            # Plaza completa
            plaza_pago = " ".join(tokens[2:tipo_idx])
            tipo_letra = tokens[tipo_idx]
            documento = tokens[tipo_idx + 1]
            descripcion = " ".join(tokens[tipo_idx + 2:])

            # 🔧 Documento = Tipo + número
            documento_final = f"{tipo_letra} {documento}"
            tipo_texto = descripcion.strip()

            registros.append([
                fecha,
                f"{plaza_codigo} {plaza_pago}".strip(),
                tipo_texto,
                documento_final,
                bruto,
                ret,
                neto
            ])

        df = pd.DataFrame(
            registros,
            columns=[
                "Fecha",
                "Plaza Pago",
                "Documento",
                "Tipo",
                "Bruto",
                "Retención",
                "Neto a Pagar",
            ],
        )

        # Parche
        df[["Documento", "Tipo"]] = df[["Tipo", "Documento"]].copy()

        for c in ["Bruto", "Retención", "Neto a Pagar"]:
            df[c] = pd.to_numeric(df[c], errors="coerce")

        return df


    # --- Ejecutar ---
    df = extraer_facturas(rutas["pdf_remittance"])
    
    return df.head(100)