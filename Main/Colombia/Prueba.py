# =====================================================
# 0. Importación de librerías y módulos utilitarios
# =====================================================
import sys
import os
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from utils import * # imports relativos para integrarlos en el paquete

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

def procesar(archivo_remittance,archivo_fbl5n):


    rutas = {
        "remittance": archivo_remittance,
        "fbl5n": archivo_fbl5n,
        # Si necesitas una ruta de salida, puedes definirla aquí:
        "salida": os.path.join(os.path.dirname(archivo_remittance), "Olimpica2.xlsx")
    }
   
    customer_id = 10266237

    # =====================================================
    # 3. Lectura de Remittance
    # =====================================================
    
    remittance = (
        pd.read_excel(
            rutas["remittance"], skiprows=21, nrows=2000,
            usecols=["Código de Documento", "No. Doc", "Total a Pagar"]
        )
        .dropna(subset=["Código de Documento"])
    )
    
    # =====================================================
    # 4. Limpieza de Remittance
    # ===================================================== 

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
    
    remittance = remittance[~remittance['Tipo de Documento'].str.startswith('Totales:', na=False)]
    
    # =====================================================
    # 5. Manejo de Reglas y CARDs
    # =====================================================  
    
    conds = [
        (remittance["Referencia / Factura"].str.startswith("PMP", na=False)) & (remittance["Importe de Remittance"] < 0),
        remittance["Referencia / Factura"].str.startswith("0085", na=False),
        remittance["Referencia / Factura"].str.startswith("46", na=False),
        remittance["Referencia / Factura"].str.startswith("98", na=False),
        remittance["Referencia / Factura"].str.startswith("310", na=False)
    ]
    descuentos = ["RECHAZO", "DESCUENTO", "AVERIA", "FACT PROVEEDOR", "NOTA DEBITO"]
    motivos = ["551", "987", "522", "CSB", "987"]

    # Aplicar reglas de descuento y motivo
    remittance["Descuento"] = np.select(conds, descuentos, default=remittance["Descuento"])
    remittance["Motivo del descuento"] = np.select(conds, motivos, default=remittance["Motivo del descuento"])

    # UNA VEZ QUE VEIFIQUEN QUE YA LOGRARON LEER CORRECTAMENTE, DESCOMENTAR LAS FUNCIONES Y EJECUTARLO.
    
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

    return hrc_template
