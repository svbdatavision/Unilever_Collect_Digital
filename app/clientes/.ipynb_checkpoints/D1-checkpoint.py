import sys
import os
import re
import pandas as pd
import numpy as np
import fitz  # PyMuPDF
import io
from .formato_template import exportar_template


def _project_root():
    if getattr(sys, "frozen", False):
        macos_dir = os.path.dirname(sys.executable)
        contents_dir = os.path.dirname(macos_dir)
        app_bundle = os.path.dirname(contents_dir)
        return os.path.dirname(app_bundle)
    return os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))


def procesar():
    root = _project_root()

    rutas = {
        "remittance_pdf": os.path.join(root, "Archivos", "Remittance", "Remittance_D1.pdf"),
        "fbl5n":          os.path.join(root, "Archivos", "Base_de_datos", "FBL5N_d1.xlsx"),
        "salida":         os.path.join(root, "Archivos", "Template", "Template_HRC_D1.xlsx"),
    }

    # --- Paso 1: Extraer facturas del PDF ---
    factura_pattern = re.compile(
        r"RE\s+\d+\s+PMP\d+\s+\d{1,3}(?:\.\d{3})*\s+\d+\s+\d{1,3}(?:\.\d{3})*\s+\d+\s+\d{1,3}(?:\.\d{3})*"
    )
    facturas = []
    with fitz.open(rutas["remittance_pdf"]) as doc:
        for page in doc:
            text = page.get_text()
            for match in factura_pattern.findall(text):
                parts = match.split()
                if len(parts) == 8:
                    facturas.append({
                        "TD": parts[0],
                        "Doc.Interno": parts[1],
                        "Referencia / Factura": parts[2],
                        "Valor Bruto": parts[3].replace('.', ''),
                        "Retenciones": parts[4],
                        "IVA": parts[5].replace('.', ''),
                        "Desc/Rec": parts[6],
                        "Neto Pagado": parts[7].replace('.', '')
                    })

    remittance = pd.DataFrame(facturas)
    remittance["Importe de factura"] = remittance["Neto Pagado"].astype(float)
    remittance["Tipo de Documento"] = "Factura"

    # --- Guardar Remittance en un buffer en memoria ---
    remittance_buffer = io.BytesIO()
    with pd.ExcelWriter(remittance_buffer, engine="openpyxl") as writer:
        remittance.to_excel(writer, index=False)
    remittance_buffer.seek(0)  # rebobinar el buffer

    # --- Paso 2: Leer FBL5N y preparar ---
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

    FBL5N["Referencia / Factura"] = np.where(
        FBL5N["Reason code"] == "NRO",
        FBL5N["Document Number"].astype("Int64").astype(str),
        FBL5N["Referencia / Factura"]
    )

    # --- Paso 3: Merge y diferencias ---
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
            hrc_template["Referencia / Factura"].fillna("")
        )
    )
    cond_1 = (hrc_template["Motivo del descuento"] == "987") & (hrc_template["Importe de factura"] < -20000)
    cond_2 = (hrc_template["Motivo del descuento"] == "987") & (hrc_template["Importe de factura"] > 20000)
    hrc_template.loc[cond_1, "Comentarios"] = "Myr Vlr Pagado " + hrc_template.loc[cond_1, "Referencia / Factura"].fillna("")
    hrc_template.loc[cond_2, "Comentarios"] = "Saldo FC " + hrc_template.loc[cond_2, "Referencia / Factura"].fillna("")

    hrc_template["Pago Neto"] = hrc_template["Importe de factura"]

    nota_credito = FBL5N[FBL5N["Reason code"] == "NRO"].copy()
    nota_credito["Tipo de Documento"] = "Nota de Crédito"
    hrc_template = pd.concat([hrc_template, nota_credito], ignore_index=True)

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

    # --- Paso 5: Exportar con formato usando formato_template ---
    exportar_template(
        hrc_template=hrc_template,
        numero_orden="",  # TODO parametrizar
        fecha_pago="",  # TODO parametrizar
        importe_FBL3N=hrc_template["Pago Neto"].sum(),
        id_cliente="",  # TODO parametrizar
        nombre_cliente="",  # TODO parametrizar
        ruta_remittance=remittance_buffer,  # <- buffer en memoria
        ruta_salida=rutas["salida"]
    )

    return hrc_template
