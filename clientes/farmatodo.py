import pandas as pd
import numpy as np
from datetime import datetime
from openpyxl import load_workbook
import os
from .formato_template import exportar_template


def procesar():
    # --- Rutas ---
    rutas = {
        "remittance": os.path.join("Archivos", "Remittance", "Remittance_farmatodo.xlsx"),
        "fbl5n": os.path.join("Archivos", "Base_de_datos", "FBL5N_farmatodo.xlsx"),
        "salida": os.path.join("Archivos", "Template", "Template_HRC_farmatodo.xlsx"),
    }

    # --- Paso 1: Leer Remittance ---
    remittance = pd.read_excel(rutas["remittance"], skiprows=6, nrows=20, header=[0, 1])
    remittance.columns = [
        str(col[0]).strip() if "Unnamed" in str(col[1]) else str(col[1]).strip()
        for col in remittance.columns
    ]
    remittance = remittance[["Nro Factura", "Descripción", "Total"]]
    remittance = remittance.rename(columns={
        "Nro Factura": "Referencia / Factura",
        "Total": "Importe de factura"
    })
    remittance["Importe de factura"] = pd.to_numeric(remittance["Importe de factura"], errors="coerce").round(2)

    # --- Paso 2: Tipo de Documento ---
    conds = [
        remittance["Referencia / Factura"].str.startswith("PMP", na=False) & (remittance["Importe de factura"] > 0),
        ~remittance["Referencia / Factura"].str.startswith("PMP", na=False) & (remittance["Importe de factura"] > 0),
        remittance["Importe de factura"] < 0
    ]
    choices = ["Factura", "Nota Debito", "Descuentos no asociados a FC"]
    remittance["Tipo de Documento"] = np.select(conds, choices, default="")

    # --- Paso 3: Descuentos y motivos ---
    if "Descuento" not in remittance.columns:
        remittance["Descuento"] = ""
    if "Motivo del descuento" not in remittance.columns:
        remittance["Motivo del descuento"] = ""

    conds_desc = [
        remittance["Referencia / Factura"].str.startswith("NC-DC05", na=False),
        remittance["Referencia / Factura"].str.startswith("NC-DC04", na=False),
        remittance["Referencia / Factura"].str.startswith("NC-DC06", na=False),
        remittance["Referencia / Factura"].str.contains("XKZV", na=False),
        remittance["Referencia / Factura"].str.startswith("NCF-DC", na=False),
        remittance["Referencia / Factura"].str.startswith("NC-PMP", na=False),
        remittance["Referencia / Factura"].str.startswith("CNQPMP", na=False),
        remittance["Referencia / Factura"].str.startswith("NC-100", na=False),
        remittance["Referencia / Factura"].str.startswith("NC1", na=False),
    ]
    descuentos = [
        "CONVENIOS", "DSCT PROMOCIONAL", "DSCT PROMOCIONAL", "FACT PROVEEDOR",
        "CONVENIOS", "DSCT AVERIAS", "PGO PDTE SOPORTE", "DSCT AVERIAS", "DSCT AVERIAS"
    ]
    motivos = ["657", "987", "987", "CSB", "657", "206", "551", "522", "522"]

    remittance["Descuento"] = np.select(conds_desc, descuentos, default=remittance["Descuento"])
    remittance["Motivo del descuento"] = np.select(conds_desc, motivos, default=remittance["Motivo del descuento"])

    # --- Paso 4: Ordenar ---
    remittance = remittance.sort_values(by="Tipo de Documento", ascending=False).reset_index(drop=True)

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

    # Ajustar notas de crédito
    FBL5N["Referencia / Factura"] = np.where(
        FBL5N["Reason code"] == "NRO",
        FBL5N["Document Number"].astype("Int64").astype(str),
        FBL5N["Referencia / Factura"]
    )

    # --- Paso 6: Merge ---
    hrc_template = pd.merge(remittance, FBL5N, on="Referencia / Factura", how="left")

    # --- Paso 7: Diferencias ---
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
        "Motivo del descuento": diferencias["Diferencia"].apply(lambda x: "WOB" if x < 0 else "384"),
    })
    hrc_template = pd.concat([hrc_template, registros_diferencias], ignore_index=True)

    # --- Paso 8: Comentarios y Pago Neto ---
    hrc_template["Comentarios"] = np.where(
        hrc_template["Tipo de Documento"] == "Factura",
        "",
        np.where(
            hrc_template["Descuento"] == "MENORES VALORES",
            hrc_template["Descuento"],
            hrc_template["Descuento"].fillna("") + " " +
            hrc_template["Referencia / Factura"].fillna("") + " " +
            hrc_template.get("Descripción", "").fillna("")
        )
    )
    hrc_template["Pago Neto"] = hrc_template["Importe de factura"]

    # --- Paso 9: Notas de Crédito ---
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

    # --- Paso 10: Columnas finales ---
    columnas_finales = [
        "Tipo de Documento", "Referencia / Factura", "Importe de factura",
        "Descuento", "Motivo del descuento", "Pago Neto", "Comentarios"
    ]
    hrc_template = hrc_template[columnas_finales]

    # --- Paso 11: Datos dinámicos ---
    wb_rem = load_workbook(rutas["remittance"], data_only=True)
    ws_rem = wb_rem.active
    numero_orden = ws_rem["B7"].value

    fbl5n = pd.read_excel(rutas["fbl5n"], usecols=["Customer", "Name 1"], nrows=1)
    id_cliente = fbl5n["Customer"].iloc[0]
    nombre_cliente = fbl5n["Name 1"].iloc[0]

    fecha_pago = datetime.today().strftime("%-m/%-d/%y")
    importe_FBL3N = hrc_template["Pago Neto"].sum()  # usamos el total calculado

    # --- Paso 12: Exportar con formato genérico ---
    exportar_template(
        hrc_template,
        numero_orden,
        fecha_pago,
        importe_FBL3N,
        id_cliente,
        nombre_cliente,
        rutas["remittance"],
        rutas["salida"]
    )

    return hrc_template
