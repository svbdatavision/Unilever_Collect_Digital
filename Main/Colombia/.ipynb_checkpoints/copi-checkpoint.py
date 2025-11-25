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
def procesar(archivo_remittance,archivo_fbl5n):


    rutas = {
        "remittance": archivo_remittance,
        "fbl5n": archivo_fbl5n,
        "salida": os.path.join(os.path.dirname(archivo_remittance), "Copidrogas.xlsx")
    }
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
    # Normalizar campos 
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
    # Normalización de compo "Factura" en Tipo de Documento
    remittance["Tipo de Documento"] = remittance["Tipo de Documento"].replace({
        "Factura Acrededor": "Factura"
    })
    # Asegurar columnas base
    for col in ["Descuento", "Motivo del descuento", "Comentarios"]:
        if col not in remittance.columns:
            remittance[col] = ""


    # =====================================================
    # 5. Manejo de Reglas y CARDs
    # =====================================================  
    # A) REFERENCIA VACÍA → USA "Tipo de Documento"
    mask_ref_vacia = (
        remittance["Referencia / Factura"].isna() |
        (remittance["Referencia / Factura"].astype(str).str.strip().isin(["", "nan"]))
    )

    remittance.loc[mask_ref_vacia, "Referencia / Factura"] = remittance.loc[mask_ref_vacia, "Tipo de Documento"]
    remittance.loc[mask_ref_vacia, "Comentarios"] = remittance.loc[mask_ref_vacia, "Texto"].astype(str)

    # B) DEVOLUCIÓN
    mask_dev = remittance["Tipo de Documento"] == "Devolucion"
    remittance.loc[mask_dev, ["Descuento", "Motivo del descuento"]] = ["AVERIA", "522"]
    remittance.loc[mask_dev, "Comentarios"] = (
        remittance.loc[mask_dev, "Tipo de Documento"].astype(str) + " " +
        remittance.loc[mask_dev, "Referencia / Factura"].astype(str)
    )

    # C) TRASLADO NOTAS DEUDOR ACREEDOR
    mask_tnda = remittance["Tipo de Documento"] == "Traslado Notas  Deudor acreedor"
    remittance.loc[mask_tnda, ["Descuento", "Motivo del descuento"]] = ["Factura proveedor", "CSB"]
    remittance.loc[mask_tnda, "Comentarios"] = (
        "FACT PROVEEDOR " + remittance.loc[mask_tnda, "Texto"].astype(str)
    )

    # D) REDUC FACTURA COMPRA
    mask_rfc = remittance["Tipo de Documento"] == "Reduc Factura Compra"
    remittance.loc[mask_rfc, ["Descuento", "Motivo del descuento"]] = ["Ventas", "987"]
    remittance.loc[mask_rfc, "Comentarios"] = (
        remittance.loc[mask_rfc, "Tipo de Documento"].astype(str) + " " +
        remittance.loc[mask_rfc, "Referencia / Factura"].astype(str)
    )

    # E) NOTAS
    mask_nota = remittance["Tipo de Documento"] == "Nota"
    texto = remittance["Texto"].astype(str)

    # Nota → empieza por DCTO
    mask_dcto = mask_nota & texto.str.startswith("DCTO")
    remittance.loc[mask_dcto, ["Descuento", "Motivo del descuento"]] = ["DPP", "667"]
    remittance.loc[mask_dcto, "Comentarios"] = "DPP " + remittance.loc[mask_dcto, "Texto"].astype(str)

    # Nota → empieza por Dev.>
    mask_dev2 = mask_nota & texto.str.startswith("Dev.>")
    remittance.loc[mask_dev2, ["Descuento", "Motivo del descuento"]] = ["COL", "522"]
    remittance.loc[mask_dev2, "Comentarios"] = remittance.loc[mask_dev2, "Texto"].astype(str)

    # Nota → resto de casos
    mask_resto_nota = mask_nota & ~(mask_dcto | mask_dev2)
    remittance.loc[mask_resto_nota, ["Descuento", "Motivo del descuento"]] = ["Ventas", "987"]
    remittance.loc[mask_resto_nota, "Comentarios"] = remittance.loc[mask_resto_nota, "Texto"].astype(str)
    
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
    
    # No se porque pusimos esta linea:
#    hrc_template.loc[hrc_template["Tipo de Documento"] == "Nota de reintegro", ["Descuento", "Comentarios"]] = ""

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