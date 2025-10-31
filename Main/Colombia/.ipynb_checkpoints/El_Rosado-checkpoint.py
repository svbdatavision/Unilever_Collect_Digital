# El Rosado
import os
import sys
import io
import warnings
import pandas as pd
import camelot
import numpy as np
from clientes.utils import *
 
warnings.filterwarnings("ignore", category=UserWarning, module="camelot")
 

def procesar(archivo_remittance,archivo_fbl5n):
    root = _project_root()
    rutas = {
        "pdf_remittance": archivo_remittance,,
        "fbl5n": archivo_fbl5n,
        "salida": os.path.join(os.path.dirname(archivo_remittance), "El_Rosado_Mico.xlsx")
    }
    # Colocar el Customer ID del cliente
    customer_id = 10273190
 
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
 
 
    remittance = remittance[["Doc/Factura", "No.Documento", "Importe"]]
    remittance = remittance.rename(columns={
        "Doc/Factura" : "Relación Cliente",
        "No.Documento": "Referencia / Factura",
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
    
#    remittance["Importe de Remittance"] = remittance["Importe de Remittance"] * -1

# Tipo de Documento (factura / Descuentos Cliente)
    conds = [
        remittance["Relación Cliente"].str.startswith("210", na=False) & (remittance["Importe de Remittance"] > 0),
        ~remittance["Relación Cliente"].str.startswith("210", na=False) & (remittance["Importe de Remittance"] > 0)
    ]
    choices = ["Factura", "Descuentos Cliente"]
    remittance["Tipo de Documento"] = np.select(conds, choices, default="")
 
    remittance = remittance.dropna(subset=["Importe de Remittance"])
    remittance.loc[remittance["Tipo de Documento"] == "Descuentos Cliente", "Importe de Remittance"] *= -1

# =====================================================
#     Manejo de Reglas y CARDs
# =====================================================
 
    if "Descuento" not in remittance.columns:
        remittance["Descuento"] = ""
    if "Motivo del descuento" not in remittance.columns:
        remittance["Motivo del descuento"] = ""
    if "Comentarios" not in remittance.columns:
        remittance["Comentarios"] = ""    
    conds_desc = [
        remittance["Tipo de Documento"].str.startswith("Descuentos Cliente", na=False),
     ]
    descuentos = ["DESCUENTO"]
    motivos = ["987"]
   
    remittance.loc[remittance["Tipo de Documento"] == "Descuentos Cliente", "Comentarios"] = (
    remittance["Relación Cliente"].astype(str) + " " + remittance["Referencia / Factura"].astype(str)
    )
 
    remittance["Motivo del descuento"] = np.select(conds_desc, motivos, default=remittance["Motivo del descuento"])    

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