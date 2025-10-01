import sys  # Para detectar ejecución empaquetada (frozen) y resolver rutas
import os   # Para construir rutas relativas al proyecto
import pandas as pd  # Principal herramienta para manipulación tabular
import numpy as np   # Utilidades numéricas/condicionales (np.select, np.where)
from openpyxl import load_workbook  # Leer valores dinámicos desde archivos Excel
from .formato_template import exportar_template  # Paso 12: exportar con formato
from .diferencias import procesar_diferencias    # Paso 7/10: lógica centralizada de diferencias


def _project_root():
    """
    Devuelve la carpeta raíz del proyecto:
    - Si corre dentro de un .app -> carpeta que contiene el .app
    - Si corre como script -> carpeta del archivo actual (../)
    """
    if getattr(sys, "frozen", False):
        macos_dir = os.path.dirname(sys.executable)
        contents_dir = os.path.dirname(macos_dir)
        app_bundle = os.path.dirname(contents_dir)
        return os.path.dirname(app_bundle)
    return os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))


def procesar():
    """
    Orquestador principal para Farmatodo:
    Flujo numerado (1..12) siguiendo el estándar de procesos.
    Devuelve el DataFrame final listo para exportar.
    """
    root = _project_root()

    # --- Paso 0: Rutas ---
    rutas = {
        "remittance": os.path.join(root, "Archivos", "Remittance", "Remittance_farmatodo.xlsx"),
        "fbl5n": os.path.join(root, "Archivos", "Base_de_datos", "FBL5N_farmatodo.xlsx"),
        "salida": os.path.join(root, "Archivos", "Template", "Template_HRC_farmatodo.xlsx"),
    }

    # =====================================================
    # 1. Lectura de Remittance
    # =====================================================
    remittance = pd.read_excel(rutas["remittance"], skiprows=6, nrows=20, header=[0, 1])
    remittance.columns = [
        str(col[0]).strip() if "Unnamed" in str(col[1]) else str(col[1]).strip()
        for col in remittance.columns
    ]
    remittance = remittance[["Nro Factura", "Descripción", "Total"]]
    remittance = remittance.dropna(subset=["Nro Factura"]).reset_index(drop=True)
    remittance = remittance.rename(columns={
        "Nro Factura": "Referencia / Factura",
        "Total": "Importe de factura"
    })
    remittance["Importe de factura"] = pd.to_numeric(remittance["Importe de factura"], errors="coerce").round(2)
    # Limpiar caracteres invisibles
    remittance["Referencia / Factura"] = remittance["Referencia / Factura"].str.replace(r"[\u202A-\u202E\u200E\u200F]", "", regex=True).str.strip()

    # =====================================================
    # 2. Tipo de Documento
    # =====================================================
    conds = [
        remittance["Referencia / Factura"].str.startswith("PMP", na=False) & (remittance["Importe de factura"] > 0),
        ~remittance["Referencia / Factura"].str.startswith("PMP", na=False) & (remittance["Importe de factura"] > 0),
        remittance["Importe de factura"] < 0
    ]
    choices = ["Factura", "Nota Debito", "Descuentos no asociados a FC"]
    remittance["Tipo de Documento"] = np.select(conds, choices, default="")

    # =====================================================
    # 3. Descuentos y motivos
    # =====================================================
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
        "CONVENIOS", "DPP", "RECHAZOS", "DSCT AVERIAS", "DSCT AVERIAS"
    ]
    motivos = ["657", "987", "987", "CSB", "657", "206", "551", "522", "522"]
    remittance["Descuento"] = np.select(conds_desc, descuentos, default=remittance["Descuento"])
    remittance["Motivo del descuento"] = np.select(conds_desc, motivos, default=remittance["Motivo del descuento"])

    # =====================================================
    # 4. Ordenar por tipo de documento
    # =====================================================
    remittance = remittance.sort_values(by="Tipo de Documento", ascending=False).reset_index(drop=True)

    # =====================================================
    # 5. Lectura de FBL5N
    # =====================================================
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

    # =====================================================
    # 6. Merge entre Remittance y FBL5N
    # =====================================================
    hrc_template = pd.merge(remittance, FBL5N, on="Referencia / Factura", how="left")
    hrc_template["Referencia / Factura"] = hrc_template["Referencia / Factura"].str.replace(r"^NC-", "", regex=True)

    # =====================================================
    # 7. Procesar diferencias
    # =====================================================
    hrc_template = procesar_diferencias(hrc_template)

    # =====================================================
    # 8. Comentarios y limpieza de CNQ*
    # =====================================================
    hrc_template["Comentarios"] = np.where(
        hrc_template["Tipo de Documento"] == "Factura", "",
        np.where(
            hrc_template["Descuento"] == "MENORES VALORES",
            hrc_template["Descuento"],
            hrc_template["Descuento"].fillna("") + " " + hrc_template["Referencia / Factura"].fillna("") + " " + hrc_template["Descripción"].fillna("")
        )
    )
    hrc_template = hrc_template[~hrc_template["Referencia / Factura"].astype(str).str.startswith("CNQ", na=False)]
    hrc_template["Pago Neto"] = hrc_template["Importe de factura"]

    # =====================================================
    # 9. Notas de Crédito
    # =====================================================
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

    # =====================================================
    # 10. Columnas finales
    # =====================================================
    columnas_finales = [
        "Tipo de Documento", "Referencia / Factura", "Importe de factura",
        "Descuento", "Motivo del descuento", "Pago Neto", "Comentarios"
    ]
    hrc_template = hrc_template[columnas_finales]

    # =====================================================
    # 11. Datos dinámicos para exportar (orden de pago, cliente)
    # =====================================================
    wb_rem = load_workbook(rutas["remittance"], data_only=True)
    ws_rem = wb_rem.active
    numero_orden = ws_rem["B7"].value

    fbl5n_meta = pd.read_excel(rutas["fbl5n"], usecols=["Customer", "Name 1"], nrows=1)
    id_cliente = fbl5n_meta["Customer"].iloc[0]
    nombre_cliente = fbl5n_meta["Name 1"].iloc[0]

    # =====================================================
    # 12. Exportar Template final con formato y hoja Remittance
    # =====================================================
    exportar_template(
        hrc_template=hrc_template,
        numero_orden=numero_orden,
        id_cliente=id_cliente,
        nombre_cliente=nombre_cliente,
        ruta_remittance=rutas["remittance"],
        ruta_salida=rutas["salida"]
    )

    return hrc_template
