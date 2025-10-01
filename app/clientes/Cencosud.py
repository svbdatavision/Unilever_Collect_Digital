import os
import sys
import re
import numpy as np
import pandas as pd
import camelot
from PyPDF2 import PdfReader
from openpyxl import load_workbook
import io
from .formato_template import exportar_template
from .diferencias import procesar_diferencias

import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="camelot")


def _project_root():
    """Carpeta raíz del proyecto."""
    if getattr(sys, "frozen", False):
        macos_dir = os.path.dirname(sys.executable)
        contents_dir = os.path.dirname(macos_dir)
        app_bundle = os.path.dirname(contents_dir)
        return os.path.dirname(app_bundle)
    return os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))


def procesar():
    """
    Proceso general de Remittance + FBL5N para Cencosud
    Incluye comentarios específicos de Cencosud.
    """
    root = _project_root()

    # --- Rutas de entrada/salida ---
    rutas = {
        "pdf_remittance": os.path.join(root, "Archivos", "Remittance", "Remittance_Cenco.pdf"),  # Cencosud
        "fbl5n": os.path.join(root, "Archivos", "Base_de_datos", "FBL5N_Cenco.xlsx"),
        "salida": os.path.join(root, "Archivos", "Template", "Template_HRC_Cenco.xlsx")
    }

    # =====================================================
    # 1. Lectura de Remitente (PDF)
    # =====================================================
    # Extraemos tablas con Camelot y PyPDF2
    reader = PdfReader(rutas["pdf_remittance"])
    num_pages = len(reader.pages)
    all_pages = set(range(1, num_pages + 1))

    tables_stream = camelot.read_pdf(rutas["pdf_remittance"], pages='all', flavor='stream', strip_text='\n')
    stream_pages = set(int(t.page) for t in tables_stream)
    missing_pages = all_pages - stream_pages

    tables_lattice = []
    if missing_pages:
        missing_pages_str = ",".join(str(p) for p in missing_pages)
        tables_lattice = camelot.read_pdf(rutas["pdf_remittance"], pages=missing_pages_str, flavor='lattice', strip_text='\n')

    df_all = pd.concat([t.df for t in list(tables_stream) + list(tables_lattice)], ignore_index=True)

    # Detectar fila de encabezado real
    header_row_idx = df_all[df_all.apply(lambda r: r.astype(str).str.contains('VOUCHER').any(), axis=1)].index[0]
    df_all.columns = df_all.iloc[header_row_idx]
    df_all = df_all.drop(index=list(range(header_row_idx + 1))).reset_index(drop=True)
    df_all = df_all[~df_all.apply(lambda r: all(r.astype(str) == df_all.columns.astype(str)), axis=1)].reset_index(drop=True)

    # =====================================================
    # 2. Limpieza de Remittance
    # =====================================================
    filter_values = ['CH','DAV','DCA','DCC','DCF','DEV','DND','DPC','FPM','FS','LTG','RPL','VOUCHER']
    remittance = df_all[df_all[df_all.columns[0]].isin(filter_values)].copy()

    def sort_key(val):
        if val == "VOUCHER": return "0"
        if val == "FPM":     return "1"
        return "2" + str(val)

    remittance["sort_order"] = remittance[remittance.columns[0]].apply(sort_key)
    remittance = (remittance.sort_values(by="sort_order")
                  .drop(columns="sort_order")
                  .drop_duplicates()
                  .reset_index(drop=True))

    # 2.1 Guardar Remittance en buffer en memoria
    remittance_buffer = io.BytesIO()
    with pd.ExcelWriter(remittance_buffer, engine="openpyxl") as writer:
        remittance.to_excel(writer, index=False)
    remittance_buffer.seek(0)

    # =====================================================
    # 3. Renombrar columnas de interés
    # =====================================================
    remittance = remittance.rename(columns={
        "DESCRIPCION": "Tipo de Documento",
        "DOCUMENTO": "Referencia / Factura",
        "VALOR PAG": "Importe de factura"
    })

    # Limpieza de importes
    remittance["Importe de factura"] = (
        pd.to_numeric(
            remittance["Importe de factura"].astype(str)
            .str.replace(".", "", regex=False)
            .str.replace(",", "", regex=False)
            .str.extract(r"([-\d]+)")[0],
            errors="coerce"
        )
    )
    remittance = remittance[remittance["Importe de factura"].notna()]
    remittance["Importe de factura"] *= -1  # Ajustar signo

    # Asegurar columnas de descuento
    for col in ["Descuento","Motivo del descuento","Comentarios"]:
        if col not in remittance.columns:
            remittance[col] = ""

    # =====================================================
    # 4. Reglas de negocio por VOUCHER
    # =====================================================
    # LTG
    mask = remittance["VOUCHER"] == "LTG"
    remittance.loc[mask, "Tipo de Documento"] = "Descuentos no asociados a FC"
    remittance.loc[mask, "Descuento"] = "Rechazo"
    remittance.loc[mask, "Motivo del descuento"] = "551"
    remittance.loc[mask, "Comentarios"] = remittance.loc[mask, "Descuento"] + " " + remittance.loc[mask, "Referencia / Factura"]

    # Grupo de vouchers: DAT, DAV, DCA, DCC, DCF, DND, DPC, RPC, DCR, RPL
    grupo = ["DAT","DAV","DCA","DCC","DCF","DND","DPC","RPC","DCR", "RPL"]
    mask = remittance["VOUCHER"].isin(grupo)
    agrupados = (
        remittance.loc[mask]
        .groupby(["VOUCHER", "Tipo de Documento", "SECCION", "DOC.SOPORTE"], as_index=False)
        .agg({
            "Referencia / Factura": lambda x: " ".join(x.dropna().astype(str)),
            "Importe de factura": "sum"
        })
    )
    agrupados["Referencia / Factura"] = agrupados["Tipo de Documento"]
    agrupados["Tipo de Documento"] = "Descuentos no asociados a FC"
    agrupados["Descuento"] = agrupados["SECCION"]
    agrupados["Motivo del descuento"] = "987"
    agrupados["Comentarios"] = agrupados["VOUCHER"] + " DSCTO " + agrupados["DOC.SOPORTE"].fillna("") + " " + agrupados["Referencia / Factura"].fillna("") + " " + agrupados["SECCION"].fillna("")
    remittance = pd.concat([remittance.loc[~mask].copy(), agrupados], ignore_index=True)

    # FS
    mask = remittance["VOUCHER"] == "FS"
    remittance.loc[mask, "Tipo de Documento"] = "Descuentos no asociados a FC"
    remittance.loc[mask, "Descuento"] = "FACT PROVEED"
    remittance.loc[mask, "Motivo del descuento"] = "CSB"
    remittance.loc[mask, "Comentarios"] = remittance.loc[mask, "Referencia / Factura"]

    # DEV
    mask = remittance["VOUCHER"] == "DEV"
    remittance.loc[mask, "Referencia / Factura"]  = remittance.loc[mask, "Tipo de Documento"]
    remittance.loc[mask, "Descuento"] = remittance.loc[mask, "Tipo de Documento"]
    remittance.loc[mask, "Tipo de Documento"] = "Descuentos no asociados a FC"
    remittance.loc[mask, "Motivo del descuento"] = "522"
    remittance.loc[mask, "Comentarios"] = remittance.loc[mask, "VOUCHER"] + " " + remittance.loc[mask, "Referencia / Factura"] + " " + remittance.loc[mask, "SECCION"]

    # CH (MENORES VALORES)
    mask = remittance["VOUCHER"] == "CH"
    remittance.loc[mask, "Referencia / Factura"] = remittance.loc[mask, "Tipo de Documento"]
    remittance.loc[mask, "Descuento"] = remittance.loc[mask, "Tipo de Documento"]
    remittance.loc[mask, "Tipo de Documento"] = "Descuentos no asociados a FC"
    remittance.loc[mask, "Motivo del descuento"] = np.where(
        (remittance.loc[mask, "Importe de factura"].abs() >= 20000), "987",
        np.where(remittance.loc[mask, "Importe de factura"].between(-20000,0, inclusive="neither"), "WOB","384")
    )
    remittance.loc[mask, "Comentarios"] = "MENORES VALORES"

    # =====================================================
    # 5. Lectura de FBL5N y merge
    # =====================================================
    FBL5N = pd.read_excel(rutas["fbl5n"], usecols=["Document Type", "Reference", "Amount in local currency"])
    FBL5N = FBL5N[FBL5N["Document Type"] == "RV"]
    FBL5N = FBL5N.rename(columns={"Reference": "Referencia / Factura", "Amount in local currency": "importe_FBL5N"}).reset_index(drop=True)

    hrc_template = pd.merge(remittance, FBL5N, on="Referencia / Factura", how="left")

    # =====================================================
    # 6. Calculamos diferencias
    # =====================================================
    hrc_template = procesar_diferencias(hrc_template)
    hrc_template["Pago Neto"] = hrc_template["Importe de factura"]

    # =====================================================
    # 7. Definimos columnas finales para template
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
    # 8. Exportar a Template
    # =====================================================
    exportar_template(
        hrc_template=hrc_template,
        numero_orden="",   # FALTA: parametrizar
        id_cliente="",     # FALTA: parametrizar
        nombre_cliente="", # FALTA: parametrizar
        ruta_remittance=remittance_buffer,  # buffer en memoria
        ruta_salida=rutas["salida"]
    )

    return hrc_template
