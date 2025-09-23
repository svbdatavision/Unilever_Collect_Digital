import os
import sys
import re
import numpy as np
import pandas as pd
import camelot
from PyPDF2 import PdfReader
from datetime import datetime
from openpyxl import load_workbook
import io
from .formato_template import exportar_template


def _project_root():
    """Carpeta raíz del proyecto."""
    if getattr(sys, "frozen", False):
        macos_dir = os.path.dirname(sys.executable)
        contents_dir = os.path.dirname(macos_dir)
        app_bundle = os.path.dirname(contents_dir)
        return os.path.dirname(app_bundle)
    return os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))


def procesar():
    root = _project_root()

    # --- Rutas de entrada/salida ---
    rutas = {
        "pdf_remittance": os.path.join("Archivos", "Remittance", "Remittance_Cenco.pdf"),
        "fbl5n":          os.path.join("Archivos", "Base_de_datos", "FBL5N_Cenco.xlsx"),
        "fbl3n":          os.path.join("Archivos", "Base_de_datos", "FBL3N.xlsx"),
        "salida":         os.path.join("Archivos", "Template", "Template_HRC_Cenco.xlsx")
    }

    pdf_to_read = rutas["pdf_remittance"]

    # ------------------------------------------------------------------
    # 1) Extraer todas las tablas del PDF de forma robusta
    # ------------------------------------------------------------------
    # Número total de páginas
    reader = PdfReader(pdf_to_read)
    num_pages = len(reader.pages)
    all_pages = set(range(1, num_pages+1))

    # Intentamos 'stream' primero
    tables_stream = camelot.read_pdf(pdf_to_read, pages='all', flavor='stream', strip_text='\n')

    # Identificar páginas vacías en stream
    stream_pages = set(int(t.page) for t in tables_stream)
    missing_pages = all_pages - stream_pages

    # Si hay páginas que no se detectaron, usar 'lattice' solo allí
    tables_lattice = []
    if missing_pages:
        missing_pages_str = ",".join(str(p) for p in missing_pages)
        tables_lattice = camelot.read_pdf(pdf_to_read, pages=missing_pages_str, flavor='lattice', strip_text='\n')

    # Concatenar todas las tablas detectadas
    df_all = pd.concat([t.df for t in list(tables_stream) + list(tables_lattice)], ignore_index=True)

    # Detectar la fila de encabezado real
    header_row_idx = df_all[df_all.apply(lambda r: r.astype(str).str.contains('VOUCHER').any(), axis=1)].index[0]
    df_all.columns = df_all.iloc[header_row_idx]
    df_all = df_all.drop(index=list(range(header_row_idx + 1))).reset_index(drop=True)
    df_all = df_all[~df_all.apply(lambda r: all(r.astype(str) == df_all.columns.astype(str)), axis=1)].reset_index(drop=True)

    # Lista de columnas de referencia
    ref_columns = df_all.columns.tolist()

    # ------------------------------------------------------------------
    # 2) Filtrar y ordenar
    # ------------------------------------------------------------------
    filter_values = ['CH','DAV','DCA','DCC','DCF','DEV','DND','DPC','FPM','FS','LTG','RPL','VOUCHER']

    filtered_df = df_all[df_all[df_all.columns[0]].isin(filter_values)].copy()

    def sort_key(val):
        if val == "VOUCHER": return "0"
        if val == "FPM":     return "1"
        return "2" + str(val)

    filtered_df["sort_order"] = filtered_df[filtered_df.columns[0]].apply(sort_key)
    filtered_df = (filtered_df
                   .sort_values(by="sort_order")
                   .drop(columns="sort_order")
                   .drop_duplicates()
                   .reset_index(drop=True))

    # ------------------------------------------------------------------
    # 3) Renombrar columnas de interés para generar remittance
    # ------------------------------------------------------------------
    remittance = filtered_df.rename(columns={
        "DESCRIPCION": "Tipo de Documento",
        "DOCUMENTO": "Referencia / Factura",
        "VALOR PAG": "Importe de factura"
    })[
        ["VOUCHER", "Tipo de Documento", "Referencia / Factura", "Importe de factura", "DOC.SOPORTE", "SECCION"]
    ]

    # Limpieza de importe
    # --- Limpieza de "Importe de factura": forzar solo valores numéricos ---
    remittance["Importe de factura"] = (
        pd.to_numeric(
            remittance["Importe de factura"]
            .astype(str)
            .str.replace(".", "", regex=False)
            .str.replace(",", "", regex=False)
            .str.extract(r"([-\d]+)")[0],  # extrae solo números y signo
            errors="coerce"
        )
    )

    # eliminar filas que no tienen importe numérico
    remittance = remittance[remittance["Importe de factura"].notna()].reset_index(drop=True)

    # cambiar signo según tu lógica original
    remittance["Importe de factura"] = remittance["Importe de factura"] * -1


    # ------------------------------------------------------------------
    # 5) Reglas de negocio por VOUCHER
    # ------------------------------------------------------------------
    for col in ["Descuento","Motivo del descuento"]:
        if col not in remittance.columns:
            remittance[col] = ""

    # LTG
    mask = remittance["VOUCHER"] == "LTG"
    remittance.loc[mask, "Tipo de Documento"] = "Descuentos no asociados a FC"
    remittance.loc[mask, "Descuento"] = "Rechazo"
    remittance.loc[mask, "Motivo del descuento"] = "551"
    
    # eliminar nombres duplicados de columnas
    remittance = remittance.loc[:, ~remittance.columns.duplicated()].copy()

    remittance.loc[mask, "Comentarios"] = remittance.loc[mask, "Descuento"] + remittance.loc[mask, "Referencia / Factura"]

    # DAT, DAV, DCA, DCC, DCF, DND, DPC, RPC, DCR
    grupo = ["DAT","DAV","DCA","DCC","DCF","DND","DPC","RPC","DCR"]
    mask = remittance["VOUCHER"].isin(grupo)
    remittance.loc[mask, "Tipo de Documento"] = "Descuentos no asociados a FC"
    remittance.loc[mask, "Descuento"] = remittance.loc[mask, "Tipo de Documento"]
    remittance.loc[mask, "Motivo del descuento"] = "987"
    remittance.loc[mask, "Comentarios"] = ("DSCTO"
                                           + remittance.loc[mask, "DOC.SOPORTE"]
                                           + remittance.loc[mask, "Tipo de Documento"]
                                           + remittance.loc[mask, "SECCION"])
    # agrupar sumando Importe solo en este grupo
    remittance = pd.concat([
        remittance[~remittance["VOUCHER"].isin(grupo)],
        remittance[remittance["VOUCHER"].isin(grupo)]
            .groupby(["VOUCHER","Tipo de Documento","Descuento",
                      "Motivo del descuento","Comentarios",
                      "DOC.SOPORTE","SECCION"], as_index=False)
            .agg({"Referencia / Factura":"first","Importe de factura":"sum"})
    ], ignore_index=True)

    # FS
    mask = remittance["VOUCHER"] == "FS"
    remittance.loc[mask, "Tipo de Documento"] = "Descuentos no asociados a FC"
    remittance.loc[mask, "Descuento"] = "FACT PROVEED"
    remittance.loc[mask, "Motivo del descuento"] = "CSB"
    remittance.loc[mask, "Comentarios"] = remittance.loc[mask, "Descuento"] + remittance.loc[mask, "Referencia / Factura"]

    # DEV
    mask = remittance["VOUCHER"] == "DEV"
    remittance.loc[mask, "Tipo de Documento"] = "Descuentos no asociados a FC"
    remittance.loc[mask, "Descuento"] = remittance.loc[mask, "Tipo de Documento"]
    remittance.loc[mask, "Motivo del descuento"] = "522"
    remittance.loc[mask, "Comentarios"] = remittance.loc[mask, "Tipo de Documento"] + remittance.loc[mask, "SECCION"]

    # CH (usar norma MENORES VALORES del template original)
    mask = remittance["VOUCHER"] == "CH"
    remittance.loc[mask, "Tipo de Documento"] = "Descuentos no asociados a FC"
    remittance.loc[mask, "Descuento"] = remittance.loc[mask, "Tipo de Documento"]
    remittance.loc[mask, "Motivo del descuento"] = np.where(
        (remittance.loc[mask, "Importe de factura"].abs() >= 20000), "987",
        np.where(remittance.loc[mask, "Importe de factura"].between(-20000,0, inclusive="neither"), "WOB","384")
    )
    remittance.loc[mask, "Comentarios"] = "MENORES VALORES"
    
    # --- Guardar Remittance en un buffer en memoria ---
    remittance_buffer = io.BytesIO()
    with pd.ExcelWriter(remittance_buffer, engine="openpyxl") as writer:
        remittance.to_excel(writer, index=False)
    remittance_buffer.seek(0)  # rebobinar el buffer
    
    # ------------------------------------------------------------------
    # 6) FBL5N y merge
    # ------------------------------------------------------------------
    FBL5N = pd.read_excel(rutas["fbl5n"], usecols=["Document Type", "Reference", "Amount in local currency"])
    FBL5N = FBL5N[FBL5N["Document Type"] == "RV"]
    FBL5N = FBL5N.rename(columns={"Reference":"Referencia / Factura","Amount in local currency":"importe_FBL5N"}).reset_index(drop=True)

    hrc_template = pd.merge(remittance, FBL5N, on="Referencia / Factura", how="left")
    hrc_template["Diferencia"] = pd.NA
    hrc_template.loc[hrc_template["Tipo de Documento"] == "Factura","Diferencia"] = (
        hrc_template["importe_FBL5N"] - hrc_template["Importe de factura"]
    )

    # diferencias no 0 -> registrar como Descuentos no asociados a FC
    dif = hrc_template[hrc_template["Diferencia"].notna() & (hrc_template["Diferencia"] != 0)].copy()
    registros_dif = pd.DataFrame({
        "Tipo de Documento": "Descuentos no asociados a FC",
        "Referencia / Factura": dif["Referencia / Factura"],
        "Importe de factura": dif["Diferencia"],
        "Pago Neto": "",
        "Descuento": "MENORES VALORES",
        "Motivo del descuento": np.select(
            condlist=[
                (dif["Diferencia"] <= -20000) | (dif["Diferencia"] >= 20000),
                (dif["Diferencia"].between(-20000,0, inclusive="neither")),
                (dif["Diferencia"].between(0,20000, inclusive="left"))
            ],
            choicelist=["987","WOB","384"], default="Error (Revisar)"
        )
    })
    hrc_template = pd.concat([hrc_template, registros_dif], ignore_index=True)

    # Comentarios y campos finales
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
