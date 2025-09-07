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
    remittance = pd.read_excel(
        rutas["remittance"], skiprows=25, nrows=65,
        usecols=["DOC", "No. Doc", "Total a Pagar"]
    ).dropna(subset=["DOC"])

    remittance = remittance.rename(columns={
        "DOC": "Tipo de Documento",
        "No. Doc": "Referencia / Factura",
        "Total a Pagar": "Importe de factura"
    })
    remittance["Referencia / Factura"] = remittance["Referencia / Factura"].str[1:-3]
    remittance["Tipo de Documento"] = remittance["Tipo de Documento"].replace({
        "Factura Comercial": "Factura",
        "Nota Crédito": "Descuentos no asociados a FC"
    })
    remittance["Importe de factura"] = pd.to_numeric(remittance["Importe de factura"], errors="coerce").round(2)

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
        "Descuento": diferencias["Diferencia"].apply(lambda x: "MENORES VALORES" if -20000 < x < 20000 else "A definir"),
        "Motivo del descuento": diferencias["Diferencia"].apply(lambda x: "WOB" if x < 0 else "384")
    })
    hrc_template = pd.concat([hrc_template, registros_diferencias], ignore_index=True)

    hrc_template["Comentarios"] = np.where(
        hrc_template["Tipo de Documento"] == "Factura", "",
        np.where(
            hrc_template["Descuento"] == "MENORES VALORES",
            hrc_template["Descuento"],
            hrc_template["Descuento"].fillna("") + " " + hrc_template["Referencia / Factura"].fillna("")
        )
    )
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

    wb_fbl3n = load_workbook(rutas["fbl3n"], data_only=True)
    fecha_dt = datetime.strptime(wb_fbl3n.active["K8"].value, "%d.%m.%Y")
    fecha_pago = fecha_dt.strftime("%-m/%-d/%y")
    importe_FBL3N = abs(float(wb_fbl3n.active["N8"].value.replace(".", "").replace(",", ".")))

    # --- Paso 5: Exportar con formato genérico ---
    exportar_template(hrc_template, numero_orden, fecha_pago, importe_FBL3N,
                      id_cliente, nombre_cliente,
                      rutas["remittance"], rutas["salida"])

    return hrc_template
