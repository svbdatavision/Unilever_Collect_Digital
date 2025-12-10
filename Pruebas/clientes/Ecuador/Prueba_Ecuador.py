# =====================================================
# 0. Importación de librerías y módulos utilitarios
# =====================================================
import sys
import os
import pandas as pd
import numpy as np
from openpyxl import load_workbook

from clientes.utils import *  # Importación de funciones utilitarias del paquete interno 'clientes'

# ESTOS CODIGOS NOS PERMITEN VER TODAS LAS FILAS Y LAS COLUMNAS. QUITAR AL MOMENTO DE PASAR A PRODUCCION
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

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
        "remittance": os.path.join(root,"Archivos", "Remittance", "Ecuador", "Remittance_nombre_cliente.xlsx"), # Completar nombre del Remittance (excel) a trabajar
        "fbl5n": os.path.join(root,"Archivos", "Cartera", "FBL5N_nombre_cliente.xlsx"), # Completar nombre de la cartera (excel) a trabajar
        "salida": os.path.join(root,"Archivos", "Template", "Ecuador", "Template_HRC_nombre_cliente.xlsx") # Colocar el nombre de salida que deseen (Ej: Template_HRC_nombre_cliente.xlsx)
    }
    customer_id = 1 # Colocar el Customer ID del cliente

    # =====================================================
    # 3. Lectura de Remittance
    # =====================================================
    
    # =====================================================
    # 4. Limpieza de Remittance
    # =====================================================
    
    # =====================================================
    # 5. Manejo de Reglas y CARDs
    # =====================================================  
    
    
    # UNA VEZ QUE VEIFIQUEN QUE YA LOGRARON LEER CORRECTAMENTE, DESCOMENTAR LAS FUNCIONES Y EJECUTARLO.
    """
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
    wb_rem = load_workbook(rutas["remittance"], data_only=True)
    celda = wb_rem.active["C10"].value
    numero_orden = wb_rem.active["F8"].value

    exportar_template(
        hrc_template=hrc_template,
        suma_remittance = remittance["Importe de Remittance"].sum(),
        numero_orden=numero_orden,
        id_cliente=id_cliente,
        nombre_cliente=nombre_cliente,
        ruta_remittance=rutas["remittance"],
        ruta_salida=rutas["salida"]
    )

    return hrc_template"""
    
    # SALIDAS PARA PRUEBAS
    remittance.to_excel(rutas["salida"], index=False)
    
    print("\n📌 Tipos de datos del Remittance:")
    print(remittance.dtypes)

    print("\n📌 Tabla:")
    print(remittance)

    remittance.to_excel(rutas["salida"], index=False)
    return remittance