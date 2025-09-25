import sys
import pandas as pd
import numpy as np
from datetime import datetime
from openpyxl import load_workbook
import os
from .formato_template import exportar_template

def _project_root():
    """
    Devuelve la carpeta raíz del proyecto:
    - Si corre dentro de un .app -> la carpeta que contiene el .app
    - Si corre como script -> la carpeta del archivo actual (../)
    """
    if getattr(sys, "frozen", False):
        macos_dir = os.path.dirname(sys.executable)      # .../MyApp.app/Contents/MacOS
        contents_dir = os.path.dirname(macos_dir)        # .../MyApp.app/Contents
        app_bundle = os.path.dirname(contents_dir)       # .../MyApp.app
        return os.path.dirname(app_bundle)               # carpeta que contiene el .app
    return os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

def procesar():
    root = _project_root()
    # --- Rutas ---
    rutas = {
        "remittance": os.path.join("Archivos", "Remittance", "Remittance_olimpica.xlsx"),
        "fbl5n": os.path.join("Archivos", "Base_de_datos", "FBL5N_olimpica.xlsx"),
        "fbl3n": os.path.join("Archivos", "Base_de_datos", "FBL3N.xlsx"),
        "salida": os.path.join("Archivos", "Template", "Template_HRC_olimpica.xlsx")
    }

    # --- Paso 1: Leer Remittance ---
    remittance = (
        pd.read_excel(
            rutas["remittance"], skiprows=21, nrows=33,
            usecols=["Código de Documento", "No. Doc", "Total a Pagar"]
        )
        .dropna(subset=["Código de Documento"])
    )

    # eliminar filas donde aparezca la palabra "Total" (mayúsculas o minúsculas)
    remittance["Código de Documento"] = remittance["Código de Documento"].astype(str)
    remittance = remittance[~remittance["Código de Documento"].str.contains("Total", case=False, na=False)]


    remittance = remittance.rename(columns={
        "Código de Documento": "Tipo de Documento",
        "No. Doc": "Referencia / Factura",
        "Total a Pagar": "Importe de factura"
    })
    remittance["Referencia / Factura"] = remittance["Referencia / Factura"].astype(str).str[0:-3]

    remittance["Tipo de Documento"] = remittance["Tipo de Documento"].replace({
        "380": "Factura",
        "381": "Descuentos no asociados a FC"
    })

    remittance["Importe de factura"] = (
    remittance["Importe de factura"]
    .astype(str)              # convertir todo a string
    .str.replace(".", "", regex=False)  # eliminar puntos de miles
    .str.replace(",", ".", regex=False) # si existiera coma decimal, la pasamos a punto
    .astype(float)            # convertir a numérico
    .round(2)
    )

    for col in ["Descuento", "Motivo del descuento"]:
        if col not in remittance.columns:
            remittance[col] = ""

    conds = [
        (remittance["Referencia / Factura"].str.startswith("PMP", na=False)) & (remittance["Importe de factura"] < 0),
        remittance["Referencia / Factura"].str.startswith("0085489", na=False),
        remittance["Referencia / Factura"].str.startswith("463", na=False),
        remittance["Referencia / Factura"].str.startswith("9801649", na=False),
        remittance["Referencia / Factura"].str.startswith("310", na=False)
    ]
    descuentos = ["RECHAZO", "DESCUENTO", "AVERIA", "FACT PROVEEDOR", "NOTA DEBITO"]
    motivos = ["551", "987", "522", "CSB", "987"]
    remittance["Descuento"] = np.select(conds, descuentos, default=remittance["Descuento"])
    remittance["Motivo del descuento"] = np.select(conds, motivos, default=remittance["Motivo del descuento"])

    # --- Paso 2: Leer FBL5N ---
    FBL5N = pd.read_excel(rutas["fbl5n"], usecols=["Document Type", "Reference", "Amount in local currency"])
    FBL5N = FBL5N[FBL5N["Document Type"] == "RV"]
    FBL5N = FBL5N.rename(columns={
        "Reference": "Referencia / Factura",
        "Amount in local currency": "importe_FBL5N"
    }).reset_index(drop=True)

    # --- Paso 3: Merge ---
    hrc_template = pd.merge(remittance, FBL5N, on="Referencia / Factura", how="left")
    hrc_template["Diferencia"] = pd.NA
    hrc_template.loc[hrc_template["Tipo de Documento"] == "Factura", "Diferencia"] = (
        hrc_template["importe_FBL5N"] - hrc_template["Importe de factura"]
    )

    diferencias = hrc_template[hrc_template["Diferencia"].notna() & (hrc_template["Diferencia"] != 0)].copy()
    registros_diferencias = pd.DataFrame({
        "Tipo de Documento": "Descuentos no asociados a FC",
        "Referencia / Factura": diferencias["Referencia / Factura"],
        "Importe de factura": diferencias["Diferencia"],
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

    # --- Paso 4: Comentarios y campos finales ---
    hrc_template["Comentarios"] = np.where(
        hrc_template["Tipo de Documento"] == "Factura", "",
        np.where(
            hrc_template["Descuento"] == "MENORES VALORES",
            hrc_template["Descuento"],
            hrc_template["Descuento"].fillna("") + " " + hrc_template["Referencia / Factura"].fillna("")
        )
    )
    cond_1 = (hrc_template["Motivo del descuento"] == "987") & (hrc_template["Importe de factura"] < -20000)
    cond_2 = (hrc_template["Motivo del descuento"] == "987") & (hrc_template["Importe de factura"] > 20000)
    hrc_template.loc[cond_1, "Comentarios"] = "Myr Vlr Pagado " + hrc_template.loc[cond_1, "Referencia / Factura"].fillna("")
    hrc_template.loc[cond_2, "Comentarios"] = "Saldo FC " + hrc_template.loc[cond_2, "Referencia / Factura"].fillna("")
    
    hrc_template["Pago Neto"] = hrc_template["Importe de factura"]

    
    columnas_finales = ["Tipo de Documento", "Referencia / Factura", "Importe de factura",
                        "Descuento", "Motivo del descuento", "Pago Neto", "Comentarios"]
    hrc_template = hrc_template[columnas_finales]

    # --- Paso 4: Datos dinámicos ---
    wb_rem = load_workbook(rutas["remittance"], data_only=True)
    celda = wb_rem.active["C10"].value
    numero_orden = celda.split("Orden de Pago:")[1].strip() if celda and "Orden de Pago:" in celda else ""

    fbl5n = pd.read_excel(rutas["fbl5n"], usecols=["Customer", "Name 1"], nrows=1)
    id_cliente = fbl5n["Customer"].iloc[0]
    nombre_cliente = fbl5n["Name 1"].iloc[0]
    

    # --- Paso 5: Exportar con formato genérico ---
    exportar_template(hrc_template, numero_orden, 
                      id_cliente, nombre_cliente,
                      rutas["remittance"], rutas["salida"])

    return hrc_template
