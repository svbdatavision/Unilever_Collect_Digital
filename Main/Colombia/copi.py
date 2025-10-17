# =====================================================
# 0. Importación de librerías y módulos utilitarios
# =====================================================
import sys
import pandas as pd
import numpy as np
from datetime import datetime
from openpyxl import load_workbook
import io
from utils import *
import os

# =====================================================
# 1. Localización dinámica de la carpeta raíz del proyecto
# =====================================================
def _project_root():
    """
    Devuelve la carpeta raíz del proyecto:
    - Si corre dentro de un .app -> la carpeta que contiene el .app
    - Si corre como script -> la carpeta del archivo actual (../)
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
def procesar(archivo_fbl5n,archivo_remmitance):

    rutas = {
    "remittance": archivo_remittance,
    "fbl5n": archivo_fbl5n,
        # Si necesitas una ruta de salida, puedes definirla aquí:
    "salida": os.path.join(os.path.dirname(archivo_remittance), "Copidrogas.xlsx")}
    #root = _project_root()
    
    # --- Rutas ---
    #rutas = {
    #    "remittance": os.path.join(root, "Archivos", "Remittance", "Colombia", "Remittance_copi.xlsx"),
#        "fbl5n": os.path.join(root, "Archivos", "Cartera", "FBL5N.xlsx"),
    #    "fbl5n": os.path.join(root, "Archivos", "Cartera", "FBL5N_copi.xlsx"),
    #    "salida": os.path.join(root, "Archivos", "Template", "Colombia", "Template_HRC_copi.xlsx")
    #}
    # Colocar el Customer ID del cliente
    customer_id = 10267298

    # =====================================================
    # 3. Lectura de Remittance
    # =====================================================
    # Leemos la tabla principal del Remittance: (ajustado a Excel)
    remittance = pd.read_excel(
        rutas["remittance"], skiprows=1, nrows=2000,
        usecols=["Referencia", "Clase", "Importe en ML","Texto"]
    )
    
    # =====================================================
    # 4. Limpieza de Remittance
    # =====================================================
    # --- Eliminar guiones de la columna 'Referencia' ---
    remittance["Referencia"] = remittance["Referencia"].astype(str).str.replace("-", "", regex=False)

    # --- Renombrar Columnas
    remittance = remittance.rename(columns={
        "Referencia": "Referencia / Factura",
        "Clase": "Tipo de Documento",
        "Importe en ML": "Importe de Remittance",
    })
    
    # Limpio de nan dataset
    remittance = remittance.dropna(subset=["Tipo de Documento"])
    # --- Intercambiar signo de los valores
    remittance["Importe de Remittance"] = -remittance["Importe de Remittance"]


    # =====================================================
    # 5. Manejo de Reglas y CARDs
    # =====================================================  
    for col in ["Descuento", "Motivo del descuento"]:
        if col not in remittance.columns:
            remittance[col] = ""

    conds = [
        remittance["Tipo de Documento"].str.startswith("Devolucion", na=False),
        remittance["Tipo de Documento"].str.startswith("Reduc Factura Compra", na=False),
        remittance["Tipo de Documento"].str.startswith("Traslado Notas  Deudor acreedor", na=False),
       
    ]
    descuentos = ["AVERIA", "DESCUENTO", "FACT PROVEEDOR"]
    motivos = ["522", "987", "CSB"]
    remittance["Descuento"] = np.select(conds, descuentos, default=remittance["Descuento"])
    remittance["Motivo del descuento"] = np.select(conds, motivos, default=remittance["Motivo del descuento"])

    # --- Condición adicional para textos que comienzan con "DCTO 2.00%" ---
    mask_dcto = remittance["Texto"].astype(str).str.startswith("DCTO 2.00%")
    remittance.loc[mask_dcto, "Descuento"] = "DPP NO PROCEDE"
    remittance.loc[mask_dcto, "Motivo del descuento"] = "667"

    # --- Condición adicional para textos que comienzan con "Dev.>" ---
    mask_dev = remittance["Texto"].astype(str).str.startswith("Dev.>")
    remittance.loc[mask_dev, "Descuento"] = "AVERIA"
    remittance.loc[mask_dev, "Motivo del descuento"] = "522"

    # --- Condición para las Notas que no estan con datos en descuento y motivo de descuento
    condicion = (remittance["Tipo de Documento"] == "Nota") & (remittance["Descuento"].astype(str).str.strip() == "")
    remittance.loc[condicion, "Descuento"] = "DESCUENTO"
    remittance.loc[condicion, "Motivo del descuento"] = 987
    

    remittance["Tipo de Documento"] = remittance["Tipo de Documento"].replace({
        "Factura Acrededor": "Factura",
        "Devolucion": "Descuentos Clientes",
        "Reduc Factura Compra": "Descuentos Clientes"
    })
    
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
    FBL5N = procesar_cartera_cliente(rutas["fbl5n"], customer_id)
    
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
    
    hrc_template.loc[hrc_template["Tipo de Documento"] == "Nota de reintegro", ["Descuento", "Comentarios"]] = ""

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
    wb_rem = load_workbook(rutas["remittance"], data_only=True)
    ws_rem = wb_rem.active
    numero_orden = ws_rem["B7"].value

    fbl5n = pd.read_excel(rutas["fbl5n"], usecols=["Customer", "Name 1"], nrows=1)
    id_cliente = fbl5n["Customer"].iloc[0]
    nombre_cliente = fbl5n["Name 1"].iloc[0]

    exportar_template(
        hrc_template=hrc_template,
        suma_remittance = remittance["Importe de Remittance"].sum(),
        numero_orden=numero_orden,
        id_cliente=id_cliente,
        nombre_cliente=nombre_cliente,
        ruta_remittance=rutas["remittance"],
        ruta_salida=rutas["salida"]
    )

    return hrc_template