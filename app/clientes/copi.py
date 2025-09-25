#Copi
import sys
import pandas as pd
import numpy as np
from datetime import datetime
from openpyxl import load_workbook
from .formato_template import exportar_template

import os
# Asegúrate de que esta importación sea válida en tu entorno
# from formato_template import exportar_template

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
    return os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

def procesar():
    root = _project_root()
    
    # --- Rutas ---
    rutas = {
        "remittance": os.path.join(root, "Archivos", "Remittance", "Remittance_copi.xlsx"),
        "fbl5n": os.path.join(root, "Archivos", "Base_de_datos", "FBL5N_copi.xlsx"),
        "fbl3n": os.path.join(root, "Archivos", "Base_de_datos", "FBL3N.xlsx"),
        "salida": os.path.join(root, "Archivos", "Template", "Template_HRC_copi.xlsx")
    }

    # --- Paso 1: Leer Remittance ---
    remittance = pd.read_excel(
        rutas["remittance"], skiprows=1, nrows=2000,
        usecols=["Referencia", "Clase", "Importe en ML","Texto"]
    )

    # --- Eliminar guiones de la columna 'Referencia' ---
    remittance["Referencia"] = remittance["Referencia"].astype(str).str.replace("-", "", regex=False)

    # --- Renombrar Columnas
    remittance = remittance.rename(columns={
        "Referencia": "Referencia / Factura",
        "Clase": "Tipo de Documento",
        "Importe en ML": "Importe de factura",
    })

    # --- Intercambiar signo de los valores
    remittance["Importe de factura"] = remittance["Importe de factura"] * -1

    # --- Definición de Reglas (CARDs)
    for col in ["Descuento", "Motivo del descuento"]:
        if col not in remittance.columns:
            remittance[col] = ""

    conds = [
        #(remittance["Tipo de Documento"].str.startswith("Factura Acrededor", na=False)) & (remittance["Importe de factura"] < 0),
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
    remittance.loc[mask_dcto, "Motivo del descuento"] = "677"

    # --- Condición adicional para textos que comienzan con "Dev.>" ---
    mask_dev = remittance["Texto"].astype(str).str.startswith("Dev.>")
    remittance.loc[mask_dev, "Descuento"] = "AVERIA"
    remittance.loc[mask_dev, "Motivo del descuento"] = "522"

    # --- Condición para las Notas que no estan con datos en descuento y motivo de descuento
    condicion = (remittance["Tipo de Documento"] == "Nota") & (remittance["Descuento"].astype(str).str.strip() == "")
    remittance.loc[condicion, "Descuento"] = "DESCUENTO"
    remittance.loc[condicion, "Motivo del descuento"] = 987

    # --- Paso 5: Leer FBL5N ---
    FBL5N = pd.read_excel(
        rutas["fbl5n"],
        sheet_name="Sheet1",
        usecols=["Document Type", "Reference", "Amount in local currency", "Reason code", "Document Number", "Text"]
    )
    FBL5N = FBL5N[(FBL5N["Document Type"] == "RV") | (FBL5N["Reason code"] == "NRO")]
    FBL5N = FBL5N.rename(columns={
        "Reference": "Referencia / Factura",
        "Amount in local currency": "importe_FBL5N"
    }).reset_index(drop=True)

    # --- Incluir partidas NRO de la FBL5N
    FBL5N["Referencia / Factura"] = np.where(
        FBL5N["Reason code"] == "NRO",
        FBL5N["Document Number"].astype("Int64").astype(str),
        FBL5N["Referencia / Factura"]
    )

    # ---  Merge ---
    hrc_template = pd.merge(remittance, FBL5N, on="Referencia / Factura", how="left")
    hrc_template["Referencia / Factura"] = hrc_template["Referencia / Factura"].str.replace(r"^NC-", "", regex=True)

    # ---  Diferencias ---
    hrc_template["Diferencia"] = pd.NA
    hrc_template.loc[hrc_template["Tipo de Documento"] == "Factura", "Diferencia"] = (
        hrc_template["importe_FBL5N"] - hrc_template["Importe de factura"]
    )

    diferencias = hrc_template[hrc_template["Diferencia"].notna() & (hrc_template["Diferencia"] != 0)].copy()
    registros_diferencias = pd.DataFrame({
        "Referencia / Factura": diferencias["Referencia / Factura"],
        "Texto": diferencias["Texto"],
        "Importe de factura": diferencias["Diferencia"],
        "Tipo de Documento": "Descuentos no asociados a FC",
        "Pago Neto": "",
        "Descuento": "MENORES VALORES",
        "Motivo del descuento": np.select(
            condlist=[
                (diferencias["Diferencia"] <= -20000) | (diferencias["Diferencia"] >= 20000),
                (diferencias["Diferencia"].between(-20000, 0, inclusive="neither")),
                (diferencias["Diferencia"].between(0, 20000, inclusive="left"))
            ],
            choicelist=["987", "WOB", "384"],
            default="Error (Revisar)"
        )
    })

    hrc_template = pd.concat([hrc_template, registros_diferencias], ignore_index=True)

    # --- MENORES VALORES ---
    hrc_template["Comentarios"] = np.where(
    hrc_template["Tipo de Documento"] == "Factura", "",
    np.where(
        hrc_template["Descuento"] == "MENORES VALORES",
        "MENORES VALORES",
        np.where(
            hrc_template["Tipo de Documento"] == "Nota",
            hrc_template["Texto"].fillna(""),
            hrc_template["Tipo de Documento"].fillna("") + " " + hrc_template["Referencia / Factura"].fillna("")
        )
    )
)
    
    hrc_template["Pago Neto"] = hrc_template["Importe de factura"]

    # --- NRO en templeate ---
    nota_credito = FBL5N[FBL5N["Reason code"] == "NRO"].copy()
    nota_credito["Tipo de Documento"] = "Nota de Crédito"

    hrc_template = pd.concat([hrc_template, nota_credito], ignore_index=True)
    hrc_template.loc[
        hrc_template["Tipo de Documento"] == "Nota de Crédito",
        ["Importe de factura", "Pago Neto", "Comentarios", "Motivo del descuento"]
    ] = hrc_template.loc[
        hrc_template["Tipo de Documento"] == "Nota de Crédito",
        ["importe_FBL5N", "importe_FBL5N", "Text", "Reason code"]
    ].values

    # --- Columnas finales ---
    columnas_finales = [
        "Tipo de Documento", "Referencia / Factura", "Importe de factura",
        "Descuento", "Motivo del descuento", "Pago Neto", "Comentarios"
    ]
    hrc_template = hrc_template[columnas_finales]
    
    # --- Datos dinámicos ---
    wb_rem = load_workbook(rutas["remittance"], data_only=True)
    ws_rem = wb_rem.active
    numero_orden = ws_rem["B7"].value

    fbl5n = pd.read_excel(rutas["fbl5n"], usecols=["Customer", "Name 1"], nrows=1)
    id_cliente = fbl5n["Customer"].iloc[0]
    nombre_cliente = fbl5n["Name 1"].iloc[0]

#    fecha_pago = datetime.today().strftime("%#m/%#d/%y")
#    importe_FBL3N = hrc_template["Pago Neto"].sum()  # usamos el total calculado

    # --- Exportar con formato genérico ---
    exportar_template(
        hrc_template,
        numero_orden,
#        fecha_pago,
#        importe_FBL3N,
        id_cliente,
        nombre_cliente,
        rutas["remittance"],
        rutas["salida"]
    )

    return hrc_template
    
    
    print(f"Se encontraron {len(remittance)} registros en Remittance")
    print("Mostrando los primeros 45 registros:")
    print(remittance.head(15))

# Ejecutar la función
if __name__ == "__main__":
    procesar()
