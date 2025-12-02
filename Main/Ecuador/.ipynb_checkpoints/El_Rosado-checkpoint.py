# =====================================================
# 0. Importación de librerías y módulos utilitarios
# =====================================================
 
import os
import sys
import io
import warnings
import pandas as pd
import camelot
import numpy as np
from utils import *
 
warnings.filterwarnings("ignore", category=UserWarning, module="camelot")
 
# =====================================================
# 2. Función principal del proceso (procesar)
# =====================================================
def procesar(archivo_remittance,archivo_fbl5n):
    rutas = {
        "pdf_remittance": archivo_remittance,
        "fbl5n": archivo_fbl5n,
        "salida": os.path.join(os.path.dirname(archivo_remittance), "El_Rosado_Mico.xlsx")
    }
    # Colocar el Customer ID del cliente
    customer_id = 10273190
 
    # =====================================================
    # 3. Lectura de Remitente
    # =====================================================
    # Leemos la tabla principal del Remittance: (ajustado a PDF) - usando Camelot
   
    tables = camelot.read_pdf(rutas["pdf_remittance"], pages='all', flavor='stream', strip_text='\n')
    df_all = pd.concat([t.df for t in tables], ignore_index=True)
 
    # Find header row
 
    header_idx = df_all[df_all.apply(lambda r: r.astype(str).str.contains('Doc/Factura').any(), axis=1)].index[0]
    df_all.columns = df_all.iloc[header_idx]
    df_all = df_all.drop(index=list(range(header_idx + 1))).reset_index(drop=True)
    df_all = df_all[~df_all.apply(lambda r: all(r.astype(str) == df_all.columns.astype(str)), axis=1)].reset_index(drop=True)
 
    #print("Columnas detectadas:", df_all.columns.tolist())
 
    # Extract amount and currency if combined in one column
    col_comb = next((col for col in df_all.columns if 'Importe' in str(col) and 'Moneda' in str(col)), None)
 
    if col_comb:
        df_all[['Importe', 'Moneda']] = df_all[col_comb].str.extract(r'([\d.,\-]+)\s*([A-Z]{3})')
        df_all['Importe'] = (
            df_all['Importe']
            .str.replace('-', '', regex=False)
        )

    remittance = df_all
       
    # 2.1 Guardar Remittance en buffer en memoria
    remittance_buffer = io.BytesIO()
    with pd.ExcelWriter(remittance_buffer, engine="openpyxl") as writer:
        remittance.to_excel(writer, index=False)
    remittance_buffer.seek(0)
 
    # ============================================
    # LIMPIEZA
    # ============================================
    # Limpieza - Renombrar Columnas
 
 
    remittance = remittance[["Doc/Factura", "Importe"]]
    remittance = remittance.rename(columns={
        "Doc/Factura" : "Referencia / Factura",
        "Importe": "Importe de Remittance"
    })

    remittance["Importe de Remittance"] = (
    pd.to_numeric(
        remittance["Importe de Remittance"].astype(str)
        .str.extract(r"([-\d.,]+)")[0]
        .str.replace(",", "", regex=False)
        .str.replace(".", ".", regex=False),  # Mantiene el punto como decimal
        errors="coerce"
    )
    )
 
    # 1. Convertir la columna a numérico (ya lo hicimos antes)
    remittance["Importe de Remittance"] = pd.to_numeric(remittance["Importe de Remittance"], errors="coerce")
 
    # 2. Filtrar filas donde el importe no sea NaN
    remittance = remittance.dropna(subset=["Importe de Remittance"])
 
    # 3. Limpiar la columna 'Referencia / Factura' eliminando guiones
    remittance["Referencia / Factura"] = remittance["Referencia / Factura"].astype(str).str.replace("-", "", regex=False)
 

    # 1. Convertir la columna a numérico (ya lo hicimos antes)
    remittance["Importe de Remittance"] = pd.to_numeric(remittance["Importe de Remittance"], errors="coerce")

    # 2. Filtrar filas donde el importe no sea NaN
    remittance = remittance.dropna(subset=["Importe de Remittance"])

    # 3. Limpiar la columna 'Referencia / Factura' eliminando guiones
    remittance["Referencia / Factura"] = remittance["Referencia / Factura"].astype(str).str.replace("-", "", regex=False)

    # 4. Crear la columna 'Tipo de Documento'
    remittance["Tipo de Documento"] = remittance["Importe de Remittance"].apply(
        lambda x: "Factura" if x > 0 else "Descuento"
    )
 
    # print(remittance)

    # =====================================================
    # Lectura de la Cartera (FBL5N) (datos desde SAP)
    # =====================================================
 
    FBL5N = procesar_cartera_cliente(rutas["fbl5n"], customer_id)
 
    # =====================================================
    #  Merge Remittance + FBL5N por "Referencia / Factura"
    # =====================================================
 
    hrc_template = merge_remittance_cartera(remittance, FBL5N)
 
    # =====================================================
    # 11. Cálculo de diferencias
    # =====================================================
 
    hrc_template = procesar_diferencias(hrc_template)
 
    # =====================================================
    #  Agregamos registros NRO
    # =====================================================
    hrc_template = procesamiento_nro(hrc_template, FBL5N)
 
    # =====================================================
    #  Asignación de Pago Neto (Pago Neto = Importe de factura) y otros ajustes
    # Reemplazar solo valores 'CSR' en 'Motivo del descuento' por '987' (solo cuando exista la columna)
    if "Motivo del descuento" in hrc_template.columns:
        mask = hrc_template["Motivo del descuento"].astype(str).str.strip().eq("CSR")
        if mask.any():
            hrc_template.loc[mask, "Motivo del descuento"] = "987"
    if "Motivo del descuento" in hrc_template.columns:
        mask = hrc_template["Motivo del descuento"].astype(str).str.strip().eq("384")
        if mask.any():
            hrc_template.loc[mask, "Motivo del descuento"] = "WOB"
    # =====================================================
    # Por defecto, 'Pago Neto' = 'Importe de factura'
 
    # =====================================================
    # Por defecto, 'Pago Neto' = 'Importe de factura'

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
   
    # Mantener solo las columnas relevantes en el orden esperado por el template
    hrc_template = hrc_template[columnas_finales]
 
    # =====================================================
    # 15. Preparación de parámetros y extracción de datos dinámicos (para exportar_template)
    # =====================================================
    # 15.1 Extraer datos dinámicos del Remittance (orden de pago, id/nombre de cliente)
        # wb_rem = load_workbook(rutas["remittance"], data_only=True)
        # ws_rem = wb_rem.active
        # numero_orden = ws_rem["B7"].value
 
    # 15.2 Extraer id_cliente / nombre_cliente desde FBL5N (primer registro)
    fbl5n_meta = pd.read_excel(rutas["fbl5n"], usecols=["Customer", "Name 1"], nrows=1)
    id_cliente = fbl5n_meta["Customer"].iloc[0]
    nombre_cliente = fbl5n_meta["Name 1"].iloc[0]
 
    # 15.3 Exportar (la función exportar_template aplica el formato y copia hoja Remittance)
    exportar_template(
        hrc_template=hrc_template,
        suma_remittance = remittance["Importe de Remittance"].sum(),
        numero_orden="numero_orden",
        id_cliente=id_cliente,
        nombre_cliente=nombre_cliente,
        ruta_remittance=remittance_buffer,
        ruta_salida=rutas["salida"]
        )
 
    # Devolución del template final
    return hrc_template