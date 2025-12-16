# =====================================================
# 0. Importación de librerías y módulos utilitarios
# =====================================================
import sys  # Para detectar ejecución empaquetada (frozen) y resolver rutas
import os   # Para construir rutas relativas al proyecto
import pandas as pd  # Principal herramienta para manipulación tabular
import numpy as np   # Utilidades numéricas/condicionales (np.select, np.where)
from openpyxl import load_workbook  # Leer valores dinámicos desde archivos Excel

# Buscamos las funciones en la carpeta Main
# current_dir = os.path.dirname(os.path.abspath(__file__))
# project_root = os.path.abspath(os.path.join(current_dir, "..", "..", ".."))  # Sube desde /Pruebas/clientes/Colombia a /Raiz
# sys.path.append(project_root)
# from Main.utils import *  # Importación de funciones utilitarias del paquete interno 'clientes'
# from clientes.utils import *  # Importación de funciones utilitarias del paquete interno 'clientes'

# =====================================================
# 1. Localización dinámica de la carpeta raíz del proyecto
# =====================================================
def _project_root():

    # Caso 1: el programa está empaquetado (ej. PyInstaller o app bundle)
    if getattr(sys, "frozen", False):
        # sys.executable apunta al ejecutable dentro del .app (Mac) o .exe (Windows)
        macos_dir = os.path.dirname(sys.executable)
        contents_dir = os.path.dirname(macos_dir)
        app_bundle = os.path.dirname(contents_dir)
        # Devuelve la carpeta que contiene el .app (la raíz del proyecto)
        return os.path.dirname(app_bundle)
    # Caso 2: ejecución normal desde código fuente (.py)
    # Sube dos niveles desde el archivo actual para llegar a la raíz del proyecto
    return os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))

# =====================================================
# 2. Función principal del proceso (procesar)
# =====================================================

def procesar():
    """
    Orquestador principal para Farmatodo:
    Flujo numerado (1..12) siguiendo el estándar de procesos.
    Devuelve el DataFrame final listo para exportar.
    """
    # 1.1 Obtener ruta raíz del proyecto
    root = _project_root()

    # 1.2 Definición de rutas de entrada y salida
    rutas = {
        "remittance": os.path.join(root, "Archivos", "Remittance", "Colombia", "Remittance_olimpica.xlsx"),
        # "fbl5n": os.path.join(root, "Archivos", "Cartera", "FBL5N.xlsx"),
        "fbl5n": os.path.join(root, "Archivos", "Cartera", "FBL5N_olimpica.xlsx"),
        "salida": os.path.join(root, "Archivos", "Template", "Colombia", "Template_HRC_olimpica.xlsx"),
    }
    # Colocar el Customer ID del cliente
    customer_id = 10266237
    
    # =====================================================
    # 3. Lectura de Remittance
    # =====================================================
    # Leemos la tabla principal del Remittance: (ajustado a Excel)
# Leer archivo remittance
    remittance = (
    pd.read_excel(
        rutas["remittance"],
        skiprows=21,
        nrows=2000,
        usecols=["Código de Documento", "No. Doc", "Total a Pagar"],
        dtype={"Código de Documento": str, "No. Doc": str, "Total a Pagar": str}
    )
    .dropna(subset=["Código de Documento"])
    )

# Renombrar columnas
    remittance.rename(columns={
    "Código de Documento": "Tipo de Documento",
    "No. Doc": "Referencia / Factura",
    "Total a Pagar": "Importe de Remittance"
    }, inplace=True)

# Ajustar tipos y formatos
    remittance["Referencia / Factura"] = remittance["Referencia / Factura"].astype(str).str[:-3]

    remittance["Tipo de Documento"] = remittance["Tipo de Documento"].replace({
    "380": "Factura",
    "381": "Descuentos Clientes"
    })

# Limpiar y convertir el importe
    remittance["Importe de Remittance"] = (
    remittance["Importe de Remittance"]
    .str.replace(".", "", regex=False)   # quitar puntos de miles
    .str.replace(",", ".", regex=False)  # convertir coma decimal a punto
    .astype(float)
    .round(2)
    )

# Agregar columnas faltantes
    for col in ["Descuento", "Motivo del descuento"]:
        if col not in remittance.columns:
            remittance[col] = ""

    # =====================================================
    # 5. Manejo de Reglas y CARDs
    # =====================================================  
 # Definir condiciones
    conds = [
        (remittance["Referencia / Factura"].str.startswith("PMP", na=False)) & (remittance["Importe de Remittance"] < 0),
        remittance["Referencia / Factura"].str.startswith("0085", na=False),
        remittance["Referencia / Factura"].str.startswith("46", na=False),
        remittance["Referencia / Factura"].str.startswith("98", na=False),
        remittance["Referencia / Factura"].str.startswith("310", na=False)
    ]

# Valores para cada condición
    descuentos = ["RECHAZO", "DESCUENTO", "AVERIA", "FACT PROVEEDOR", "NOTA DEBITO"]
    motivos = ["551", "987", "522", "CSB", "987"]

# Aplicar reglas usando np.select
    remittance["Descuento"] = np.select(conds, descuentos, default=remittance.get("Descuento", ""))
    remittance["Motivo del descuento"] = np.select(conds, motivos, default=remittance.get("Motivo del descuento", ""))

    # print(remittance.dtypes)
    print(remittance.head())
procesar()