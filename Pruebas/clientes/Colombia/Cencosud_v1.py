# =====================================================
# 0. Importación de librerías y módulos utilitarios
# =====================================================

import os
import sys
import re
import io
import warnings

import numpy as np
import pandas as pd
import camelot
from PyPDF2 import PdfReader
from openpyxl import load_workbook
import pdfplumber
import traceback

from clientes.utils import *  # funciones del paquete interno 'clientes'

warnings.filterwarnings("ignore", category=UserWarning, module="camelot")


# =====================================================
# 1. Localización dinámica de la carpeta raíz del proyecto
# =====================================================
def _project_root():
    if getattr(sys, "frozen", False):
        macos_dir = os.path.dirname(sys.executable)
        contents_dir = os.path.dirname(macos_dir)
        app_bundle = os.path.dirname(contents_dir)
        return os.path.dirname(app_bundle)
    return os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))


# =====================================================
# 2. Función principal del proceso (procesar)
# =====================================================
def procesar():
    root = _project_root()

    rutas = {
        "remittance": os.path.join(root, "Archivos", "Remittance", "Colombia", "Remittance_Cenco.pdf"),
        "fbl5n": os.path.join(root, "Archivos", "Cartera", "FBL5N_Cenco.xlsx"),
        "salida": os.path.join(root, "Archivos", "Template", "Colombia", "Template_HRC_Cenco.xlsx")
    }

    customer_id = 10267301

    # =====================================================
    # Helpers internos
    # =====================================================
    def detect_header_positions(pdf_path, page_number=1, headers_expected=None):
        if headers_expected is None:
            headers_expected = HEADERS_EXPECTED

        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[page_number - 1]
            words = page.extract_words(use_text_flow=True)

        words_u = [(w['text'].strip().upper(), w) for w in words if w.get('text') and w['text'].strip() != ""]
        header_positions = {}
        for header in headers_expected:
            h = header.strip().upper()
            candidates = [w for t, w in words_u if t == h]
            if not candidates:
                candidates = [w for t, w in words_u if h in t]
            if not candidates:
                first = h.split()[0]
                candidates = [w for t, w in words_u if t.startswith(first)]
            if candidates:
                chosen = sorted(candidates, key=lambda d: d['x0'])[0]
                header_positions[h] = chosen['x0']
            else:
                header_positions[h] = None
        return header_positions

    def build_cuts_from_positions(header_positions):
        x_list = [int(round(v)) for v in header_positions.values() if v is not None]
        x_list = sorted(set(x_list))
        if len(x_list) < 6:
            return None
        cuts = [max(1, x_list[0] - 20)]
        for i in range(len(x_list) - 1):
            mid = int(round((x_list[i] + x_list[i+1]) / 2))
            cuts.append(mid)
        cuts.append(int(round(x_list[-1] + 120)))
        return sorted(list(dict.fromkeys(cuts)))

    def camelot_read_with_columns(pdf_path, columns_cuts):
        """
        Lee el PDF usando Camelot y aplica los cortes de columnas solo
        si TODOS los valores son numéricos válidos.
        Esto evita errores como:
           ValueError: could not convert string to float: ''
        """

        # Filtrar valores inválidos (None, "", whitespace)
        cuts_validos = [c for c in columns_cuts if isinstance(c, (int, float))]

        # Si QUEDA ALGÚN valor inválido → NO usar columns
        if len(cuts_validos) != len(columns_cuts):
            print("⚠ Advertencia: columnas inválidas detectadas → Camelot SIN 'columns'")
            cuts_validos = []

        # Construir parámetro columns si corresponde
        cols_str = ",".join(str(x) for x in cuts_validos) if cuts_validos else None

        # Llamada a Camelot
        if cols_str:
            print("✔ Usando columnas:", cols_str)
            tables = camelot.read_pdf(
                pdf_path,
                pages="all",
                flavor="stream",
                strip_text="\n",
                columns=cols_str,
                row_tol=8,
                column_tol=6
            )
        else:
            print("✔ Leyendo SIN columnas")
            tables = camelot.read_pdf(
                pdf_path,
                pages="all",
                flavor="stream",
                strip_text="\n",
                row_tol=8,
                column_tol=6
            )

        return tables


    def ensure_column_names(df):
        """
        Normaliza nombres de columna para evitar KeyError.
        Detecta variantes y renombra a las convenciones usadas en el script.
        """
        cols_map = {}
        existing = {c.upper().replace(" ", ""): c for c in df.columns}

        # mapping heurístico
        def find_col(key_variants):
            for k in key_variants:
                ku = k.upper().replace(" ", "")
                if ku in existing:
                    return existing[ku]
            return None

        mapping_candidates = {
            "VOUCHER": ["VOUCHER", "VOUCHER"],
            "DESCRIPCION": ["DESCRIPCION", "DESCRIPTION", "DESC"],
            "DOCUMENTO": ["DOCUMENTO", "DOC", "DOCUMENT"],
            "TIENDA": ["TIENDA", "STORE"],
            "SECCION": ["SECCION", "SECTION"],
            "F. REGISTRO": ["F. REGISTRO", "F.REGISTRO", "F REGISTRO", "FECHA"],
            "VALOR PAG": ["VALOR PAG", "VALORPAG", "VALOR PAGADO", "VALOR"],
            "DOC.SOPORTE": ["DOC.SOPORTE", "DOC SOPORTE", "DOCSOPORTE", "DOC SOP"]
        }

        for canonical, variants in mapping_candidates.items():
            found = find_col(variants)
            if found:
                cols_map[found] = canonical

        if cols_map:
            df = df.rename(columns=cols_map)
        return df

    # =====================================================
    # 1. Lectura de Remittance
    # =====================================================

    HEADERS_EXPECTED = [
        "VOUCHER", "DESCRIPCION", "DOCUMENTO", "TIENDA", "F. REGISTRO", "SECCION", "VALOR",
        "FAC. IVA", "FAC. RET. FUENTE", "RET. IVA", "RET. ICA", "OTROS IMP.", "VALOR PAG", "DOC.SOPORTE"
    ]

    header_positions = detect_header_positions(rutas["remittance"], page_number=1)
    cuts = build_cuts_from_positions(header_positions)

    if cuts:
        print("\n✔️ Cortes detectados automáticamente:", cuts)
    else:
        print("\n⚠️ No se detectaron suficientes columnas — usando fallback por defecto.")
        cuts = [
            8, 46, 102, 190, 222, 271, 318, 358, 410, 527, 661, 780
        ]

    tables = camelot_read_with_columns(rutas["remittance"], cuts)

    if len(tables) == 0:
        raise ValueError("Camelot no pudo leer ninguna tabla del PDF.")

    df_all = pd.concat([t.df for t in tables], ignore_index=True)

    # Intento de normalizar nombres tempranamente para evitar KeyError posteriores
    df_all = ensure_column_names(df_all)

    # ------------------------------------------------------------
    # Detectar encabezado real
    # ------------------------------------------------------------
    header_mask = df_all.apply(lambda r: r.astype(str).str.contains("VOUCHER", case=False).any(), axis=1)
    if not header_mask.any():
        raise ValueError("No se encontró fila de encabezado (VOUCHER) en df_all. Revisa extracción Camelot.")
    header_row_idx = header_mask.idxmax()

    df_all.columns = df_all.iloc[header_row_idx].astype(str).str.replace("\n", " ").str.strip()
    df_all = df_all.drop(index=range(0, header_row_idx + 1)).reset_index(drop=True)
    df_all.columns = df_all.columns.str.replace("\n", " ").str.strip()

    # quitar columnas literales vacías
    cols_to_drop = [c for c in df_all.columns if str(c).strip().lower() in ("nan", "", "none")]
    if cols_to_drop:
        df_all = df_all.drop(columns=cols_to_drop)

    df_all = df_all.loc[~df_all.apply(lambda r: r.astype(str).str.strip().eq('').all(), axis=1)].reset_index(drop=True)

    # Asegurar nombres canónicos otra vez tras reindex
    df_all = ensure_column_names(df_all)

    # ==========================================================
    # FIX: Si Camelot pegó DOCUMENTO+TIENDA en una sóla columna, separarla
    # ==========================================================
    merged_col = None
    for c in df_all.columns:
        cu = str(c).upper().replace(" ", "")
        if "DOCUMENTO" in cu and "TIENDA" in cu:
            merged_col = c
            break

    if merged_col is not None:
        merged_series = df_all[merged_col].astype(str)

        def split_documento_tienda(s):
            s = s.replace("\n", " ").strip()
            if s == "" or s.upper() == "NAN":
                return ("", "")
            m = re.search(r'\s{2,}', s)
            if m:
                left = s[:m.start()].strip()
                right = s[m.end():].strip()
                return (left, right)
            patterns = [' VPP', ' ADM', ' JUMBO', ' PLAT', ' PLATA', ' RANCHO', ' PERFUME', ' DROGUE', ' CUCU', ' DOCKIN', ' BUCARAM']
            for p in patterns:
                if p in s.upper():
                    idx = s.upper().find(p)
                    return (s[:idx].strip(), s[idx:].strip())
            parts = s.rsplit(' ', 1)
            if len(parts) == 2 and len(parts[1]) <= 10:
                return (parts[0].strip(), parts[1].strip())
            return (s, "")

        splits = merged_series.map(split_documento_tienda)
        df_all['DOCUMENTO'] = splits.map(lambda x: x[0])
        df_all['TIENDA'] = splits.map(lambda x: x[1])
        df_all = df_all.drop(columns=[merged_col])

    else:
        cols_upper = [str(c).upper().replace(" ", "") for c in df_all.columns]
        if 'DOCUMENTOTIENDA' in cols_upper:
            idx = cols_upper.index('DOCUMENTOTIENDA')
            colname = df_all.columns[idx]
            merged_series = df_all[colname].astype(str)

            def split_documento_tienda_simple(s):
                s = s.replace("\n", " ").strip()
                m = re.search(r'\s{2,}', s)
                if m:
                    left = s[:m.start()].strip()
                    right = s[m.end():].strip()
                    return (left, right)
                parts = s.rsplit(' ', 1)
                if len(parts) == 2 and len(parts[1]) <= 10:
                    return (parts[0].strip(), parts[1].strip())
                return (s, "")

            splits = merged_series.map(split_documento_tienda_simple)
            df_all['DOCUMENTO'] = splits.map(lambda x: x[0])
            df_all['TIENDA'] = splits.map(lambda x: x[1])
            df_all = df_all.drop(columns=[colname])

    # Normalizar nombres de columnas finales
    df_all.columns = [str(c).strip() for c in df_all.columns]
    df_all = ensure_column_names(df_all)  # segunda pasada para estabilizar nombres

    # Asegurar columnas esperadas
    expected = ["VOUCHER", "DESCRIPCION", "DOCUMENTO", "TIENDA", "SECCION", "F. REGISTRO", "VALOR PAG", "DOC.SOPORTE"]
    for e in expected:
        if e not in df_all.columns:
            df_all[e] = ""

    # Reindex lógico para que DOCUMENTO y TIENDA estén cerca
    cols_order = [c for c in df_all.columns if c not in ("DOCUMENTO", "TIENDA")]
    insert_at = 2 if len(cols_order) >= 2 else len(cols_order)
    if 'DOCUMENTO' in cols_order:
        cols_order.remove('DOCUMENTO')
    if 'TIENDA' in cols_order:
        cols_order.remove('TIENDA')
    new_order = cols_order[:insert_at] + ['DOCUMENTO', 'TIENDA'] + cols_order[insert_at:]
    seen = set()
    new_order = [x for x in new_order if not (x in seen or seen.add(x))]
    df_all = df_all.reindex(columns=new_order)

    # ==========================================================
    # FIX PARA FILAS DESALINEADAS CUANDO SECCION COMIENZA CON "PLAT"
    # ==========================================================
    def corregir_filas_plat(df):
        for idx, row in df.iterrows():
            sec = str(row.get("SECCION", "")).strip().upper()
            if sec.startswith("PLAT"):
                df.at[idx, "TIENDA"] = row.get("SECCION", "")
                df.at[idx, "SECCION"] = row.get("F. REGISTRO", "")
                df.at[idx, "F. REGISTRO"] = row.get("VALOR FAC.", "")
                df.at[idx, "VALOR FAC."] = row.get("IVA FAC.", "")
                df.at[idx, "IVA FAC."] = row.get("RET. FUENTE", "")
                df.at[idx, "RET. FUENTE"] = row.get("RET. IVA", "")
                df.at[idx, "RET. IVA"] = row.get("RET. ICA", "")
                df.at[idx, "RET. ICA"] = row.get("OTROS IMP.", "")
                df.at[idx, "OTROS IMP."] = row.get("VALOR PAG", "")
                df.at[idx, "VALOR PAG"] = row.get("DOC.SOPORTE", "")
                df.at[idx, "DOC.SOPORTE"] = ""
        return df

    df_all = corregir_filas_plat(df_all)

    # =====================================================
    # FIX EXTRA — corregir columnas corridas por SECCION pegada a TIENDA
    # =====================================================
    secciones_validas = ["PERFUME", "DROGUE", "PERFUMERIA", "DROGUERIA", "MONTERIA"]

    def corregir_tienda_seccion(row):
        tienda = str(row.get("TIENDA", "")).strip()
        seccion = str(row.get("SECCION", "")).strip()
        if seccion in secciones_validas:
            return row
        for sec in secciones_validas:
            if tienda.endswith(" " + sec):
                base = tienda[: -len(sec)].strip()
                row["TIENDA"] = base
                row["SECCION"] = sec
                columnas = [
                    "F. REGISTRO", "VALOR FAC.", "IVA FAC.", "RET. FUENTE",
                    "RET. IVA", "RET. ICA", "OTROS IMP.", "VALOR PAG", "DOC.SOPORTE"
                ]
                valores = [row.get(c, "") for c in columnas]
                if row.get("DOC.SOPORTE", "") in ["", None] and str(row.get("VALOR PAG", "")).isdigit():
                    valores = valores[1:] + [""]
                for i, c in enumerate(columnas):
                    row[c] = valores[i]
                return row
        return row

    df_all = df_all.apply(corregir_tienda_seccion, axis=1)

    # =====================================================
    # 4. Limpieza de Remittance
    # =====================================================
    filter_values = [
        'DAT', 'CH', 'DAV', 'DCA', 'DCC', 'DCF', 'DEV', 'DND', 'DPC', 'FPM', 'FS', 'LTG', 'RPL', 'DEC'
    ]
    remittance = df_all[df_all[df_all.columns[0]].isin(filter_values)].copy()

    # =====================================================
    # AJUSTE: completar SECCION si viene pegada al final de DOCUMENTO
    # =====================================================
    secciones_validas = ["DROGUE", "PERFUME"]

    def corregir_seccion(row):
        seccion = str(row.get("SECCION", "")).strip()
        documento = str(row.get("DOCUMENTO", "")).strip()
        if "/" in seccion:
            for sec in secciones_validas:
                if documento.endswith(sec):
                    row["SECCION"] = sec
                    row["DOCUMENTO"] = documento[:-len(sec)].strip()
                    return row
            row["SECCION"] = ""
            return row
        if seccion in ["", None]:
            for sec in secciones_validas:
                if documento.endswith(sec):
                    row["SECCION"] = sec
                    row["DOCUMENTO"] = documento[:-len(sec)].strip()
                    return row
        return row

    remittance = remittance.apply(corregir_seccion, axis=1)

    # =====================================================
    # LIMPIEZA VALOR PAG (numerización y fixes)
    # =====================================================
    remittance["VALOR PAG"] = (
        remittance["VALOR PAG"]
            .astype(str)
            .str.replace(".", "", regex=False)
            .str.replace(",", ".", regex=False)
    )
    remittance["VALOR PAG"] = pd.to_numeric(remittance["VALOR PAG"], errors="coerce")

    # Esta línea reemplaza VALOR PAG == 0 por DOC.SOPORTE (tu petición)
    remittance.loc[remittance["VALOR PAG"].fillna(0) == 0, "VALOR PAG"] = remittance.loc[remittance["VALOR PAG"].fillna(0) == 0, "DOC.SOPORTE"]

    # FIX: mover valores muy grandes de VALOR PAG a DOC.SOPORTE
    umbral = 20000000000000  # 20 billones

    def clean_to_numeric(col):
        return pd.to_numeric(col.astype(str).str.replace(".", "", regex=False), errors="coerce")

    valor_pag_num = clean_to_numeric(remittance["VALOR PAG"])
    otros_imp_num = clean_to_numeric(remittance.get("OTROS IMP.", pd.Series(["0"] * len(remittance))))

    mask_grande = valor_pag_num > umbral

    # Inicializar DOC.SOPORTE como cadena vacía
    remittance["DOC.SOPORTE"] = remittance["DOC.SOPORTE"].astype(str).fillna("")

    # Pasar valores grandes a DOC.SOPORTE (como texto original, sin romper arrays)
    remittance.loc[mask_grande, "DOC.SOPORTE"] = remittance.loc[mask_grande, "VALOR PAG"].astype(str)

    # Mover OTROS IMP. a VALOR PAG solo si OTROS IMP. != 0
    mask_otros = otros_imp_num != 0
    remittance.loc[mask_grande & mask_otros, "VALOR PAG"] = remittance.loc[mask_grande & mask_otros, "OTROS IMP."]

    # Aseguramos que DOC.SOPORTE sea string; si el valor numérico no supera umbral dejamos vacío
    doc_sop_num = pd.to_numeric(remittance["DOC.SOPORTE"].astype(str).str.replace(".", "", regex=False), errors="coerce")
    remittance["DOC.SOPORTE"] = np.where(doc_sop_num > umbral, remittance["DOC.SOPORTE"].astype(str), "")

    # OTROS IMP. siempre = 0 (según tu regla)
    if "OTROS IMP." in remittance.columns:
        remittance["OTROS IMP."] = 0
    else:
        remittance["OTROS IMP."] = 0

    # =====================================================
    # Ordenamiento y limpieza final antes de reglas
    # =====================================================
    def sort_key(val):
        if val == "VOUCHER":
            return "0"
        if val == "FPM":
            return "1"
        return "2" + str(val)

    remittance["sort_order"] = remittance[remittance.columns[0]].apply(sort_key)
    remittance = (
        remittance.sort_values(by="sort_order")
        .drop(columns="sort_order")
        .drop_duplicates()
        .reset_index(drop=True)
    )

    # Asegurar columnas string y uppercase en campos claves
    valores_validos = ["DROGUE", "PERFUME", "PLATOS", "RANCHO"]
    for col in ["F. REGISTRO", "TIENDA", "SECCION"]:
        if col in remittance.columns:
            remittance[col] = remittance[col].fillna("").astype(str).str.upper().str.strip()
        else:
            remittance[col] = ""

    # 1) Coincidencia exacta en F. REGISTRO
    if "F. REGISTRO" in remittance.columns:
        mask_registro = remittance["F. REGISTRO"].isin(valores_validos)
        remittance.loc[mask_registro, "SECCION"] = remittance.loc[mask_registro, "F. REGISTRO"]

    # 2) Coincidencia parcial dentro de TIENDA
    if "TIENDA" in remittance.columns:
        for val in valores_validos:
            mask_tienda = remittance["TIENDA"].str.contains(val, case=False, na=False)
            remittance.loc[mask_tienda, "SECCION"] = val

    # Guardar buffer de remittance en memoria
    remittance_buffer = io.BytesIO()
    with pd.ExcelWriter(remittance_buffer, engine="openpyxl") as writer:
        remittance.to_excel(writer, index=False)
    remittance_buffer.seek(0)

    # =====================================================
    # Renombrar columnas de interés (ahora que la tabla está estable)
    # =====================================================
    remittance = remittance.rename(columns={
        "DESCRIPCION": "Tipo de Documento",
        "DOCUMENTO": "Referencia / Factura",
        "VALOR PAG": "Importe de Remittance"
    })

    # Normalizar "Tipo de Documento" especial case
    if "Tipo de Documento" in remittance.columns:
        remittance.loc[remittance["Tipo de Documento"] == "FACTURA PROVEEDOR", "Tipo de Documento"] = "Factura"
    else:
        # Si no existía, intentar mapear desde DESCRIPCION alternativa
        if "DESCRIPCION" in remittance.columns:
            remittance = remittance.rename(columns={"DESCRIPCION": "Tipo de Documento"})

    # =====================================================
    # 5. Manejo de Reglas y CARDs (ahora con columnas ya renombradas)
    # =====================================================
    # LTG
    if "VOUCHER" in remittance.columns:
        mask = remittance["VOUCHER"] == "LTG"
        remittance.loc[mask, "Tipo de Documento"] = "Descuentos Clientes"
        remittance.loc[mask, "Descuento"] = "Rechazo"
        remittance.loc[mask, "Motivo del descuento"] = "551"
        remittance.loc[mask, "Comentarios"] = remittance.loc[mask, "Descuento"].astype(str) + " " + remittance.loc[mask, "Referencia / Factura"].astype(str)

    # Grupo de vouchers
    grupo = ["DAT", "DAV", "DCA", "DCC", "DCF", "DND", "DPC", "RPC", "DCR", "RPL", "DEC"]
    if "VOUCHER" in remittance.columns:
        mask = remittance["VOUCHER"].isin(grupo)
        # Agrupar por las columnas que sí existen: si "Tipo de Documento" no existe usar "DESCRIPCION" como fallback
        group_cols = []
        if "VOUCHER" in remittance.columns:
            group_cols.append("VOUCHER")
        if "Tipo de Documento" in remittance.columns:
            group_cols.append("Tipo de Documento")
        elif "DESCRIPCION" in remittance.columns:
            group_cols.append("DESCRIPCION")
        else:
            # fallback genérico: crear columna temporal
            remittance["Tipo de Documento"] = remittance.get("DESCRIPCION", "").astype(str)
            group_cols.append("Tipo de Documento")
        for c in ["SECCION", "DOC.SOPORTE"]:
            if c in remittance.columns:
                group_cols.append(c)
            else:
                remittance[c] = ""
                group_cols.append(c)

        agrupados = (
            remittance.loc[mask]
            .groupby(group_cols, as_index=False)
            .agg({
                "Referencia / Factura": lambda x: " ".join(x.dropna().astype(str)) if "Referencia / Factura" in remittance.columns else "",
                "Importe de Remittance": "sum" if "Importe de Remittance" in remittance.columns else ("sum" if "VALOR PAG" in remittance.columns else "sum")
            })
        )

        # Normalizamos los nombres producidos por el groupby
        if "Tipo de Documento" not in agrupados.columns and "DESCRIPCION" in agrupados.columns:
            agrupados = agrupados.rename(columns={"DESCRIPCION": "Tipo de Documento"})

        agrupados["Referencia / Factura"] = agrupados.get("Tipo de Documento", agrupados.get("Referencia / Factura", ""))
        agrupados["Tipo de Documento"] = "Descuentos Clientes"
        agrupados["Descuento"] = agrupados.get("SECCION", "")
        agrupados["Motivo del descuento"] = "987"
        agrupados["Comentarios"] = agrupados.get("VOUCHER", "").astype(str) + " DSCTO " + agrupados.get("DOC.SOPORTE", "").astype(str) + " " + agrupados.get("Referencia / Factura", "").astype(str) + " " + agrupados.get("SECCION", "").astype(str)

        remittance = pd.concat([remittance.loc[~mask].copy(), agrupados], ignore_index=True)

    # FS
    if "VOUCHER" in remittance.columns:
        mask = remittance["VOUCHER"] == "FS"
        remittance.loc[mask, "Tipo de Documento"] = "Descuentos Clientes"
        remittance.loc[mask, "Descuento"] = "FACT PROVEED"
        remittance.loc[mask, "Motivo del descuento"] = "CSB"
        remittance.loc[mask, "Comentarios"] = remittance.loc[mask, "Referencia / Factura"].astype(str)

    # DEV
    if "VOUCHER" in remittance.columns:
        mask = remittance["VOUCHER"] == "DEV"
        if "Tipo de Documento" not in remittance.columns:
            remittance["Tipo de Documento"] = remittance.get("DESCRIPCION", "").astype(str)
        remittance.loc[mask, "Referencia / Factura"] = remittance.loc[mask, "Tipo de Documento"]
        remittance.loc[mask, "Descuento"] = remittance.loc[mask, "Tipo de Documento"]
        remittance.loc[mask, "Tipo de Documento"] = "Descuentos Clientes"
        remittance.loc[mask, "Motivo del descuento"] = "522"
        remittance.loc[mask, "Comentarios"] = remittance.loc[mask, "VOUCHER"].astype(str) + " " + remittance.loc[mask, "Referencia / Factura"].astype(str) + " " + remittance.loc[mask, "SECCION"].astype(str)

    # CH (MENORES VALORES)
    if "VOUCHER" in remittance.columns:
        mask = remittance["VOUCHER"] == "CH"
        remittance.loc[mask, "Referencia / Factura"] = remittance.loc[mask].get("Tipo de Documento", remittance.loc[mask].get("DESCRIPCION", "")).astype(str)
        remittance.loc[mask, "Descuento"] = remittance.loc[mask].get("Tipo de Documento", remittance.loc[mask].get("DESCRIPCION", "")).astype(str)
        remittance.loc[mask, "Tipo de Documento"] = "Descuentos Clientes"
        # Ajuste pandas between con compatibilidad de versiones
        values_series = remittance.loc[mask, "Importe de Remittance"] if "Importe de Remittance" in remittance.columns else remittance.loc[mask, "VALOR PAG"]
        remittance.loc[mask, "Motivo del descuento"] = np.where(
            (values_series.abs() >= 20000), "987",
            np.where(values_series.between(-20000, 0), "WOB", "384")
        )
        remittance.loc[mask, "Comentarios"] = "MENORES VALORES"

    # =====================================================
    # 6. Procesamiento de columnas 'Descuento' y 'Comentarios'
    # (usa función externa que importaste)
    # =====================================================
    remittance = procesar_descuentos_y_comentarios(remittance)

    # =====================================================
    # 7-9. Cartera y renombrados posteriores
    # =====================================================
    FBL5N, id_cliente, nombre_cliente = procesar_cartera_cliente(rutas["fbl5n"], customer_id)

    # =====================================================
    # 10. Merge Remittance + FBL5N por "Referencia / Factura"
    # =====================================================
    hrc_template = merge_remittance_cartera(remittance, FBL5N)

    # =====================================================
    # 11. Cálculo de diferencias
    # =====================================================
    hrc_template = procesar_diferencias(hrc_template)

    # =====================================================
    # 12. Agregamos registros NRO
    # =====================================================
    hrc_template = procesamiento_nro(hrc_template, FBL5N)

    # =====================================================
    # 13. Pago Neto y limpieza final
    # =====================================================
    hrc_template["Pago Neto"] = hrc_template["Importe de factura"]

    hrc_template["Referencia / Factura"] = (
        hrc_template["Referencia / Factura"]
            .apply(lambda x: "" if x is None else str(x))
            .str.replace(r"[\[\]\(\)\{\}]", "", regex=True)
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
    )

    hrc_template = (
        hrc_template[
            ~(
                (hrc_template["Tipo de Documento"] == "Factura") &
                (hrc_template["Referencia / Factura"].str.len() != 10)
            )
        ]
        .reset_index(drop=True)
    )

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
    # 15. Exportar template final (usa tu función exportar_template)
    # =====================================================
    exportar_template(
        hrc_template=hrc_template,
        suma_remittance=remittance["Importe de Remittance"].sum() if "Importe de Remittance" in remittance.columns else 0,
        numero_orden="",
        id_cliente=id_cliente,
        nombre_cliente=nombre_cliente,
        ruta_remittance=remittance_buffer,
        ruta_salida=rutas["salida"]
    )

    return hrc_template
