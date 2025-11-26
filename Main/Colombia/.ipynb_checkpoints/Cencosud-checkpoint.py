# =====================================================
# 0. Importación de librerías y módulos utilitarios
# =====================================================

import os       # Manejo de rutas y directorios del sistema operativo
import sys      # Detección de ejecución empaquetada y manipulación de rutas del intérprete
import re       # Expresiones regulares (búsqueda y limpieza de texto)
import io        # Manejo de flujos de datos en memoria (buffers, streams)
import warnings # Control de advertencias del sistema y librerías externas

import numpy as np  # Operaciones numéricas y lógicas (np.where, np.select, etc.)
import pandas as pd  # Manipulación y análisis de datos tabulares
import camelot   # Extracción de tablas desde archivos PDF
from PyPDF2 import PdfReader  # Lectura y procesamiento de archivos PDF
from openpyxl import load_workbook  # Lectura de archivos Excel (.xlsx)
import pdfplumber
import traceback

from utils import *  # Importación de funciones utilitarias del paquete interno 'clientes'

# Configuración de advertencias
warnings.filterwarnings("ignore", category=UserWarning, module="camelot") # Suprime advertencias generadas por Camelot (usualmente por manejo de PDFs)


# =====================================================
# 1. Localización dinámica de la carpeta raíz del proyecto
# =====================================================
def _project_root():
    """
    Obtiene la ruta base del proyecto sin importar el entorno de ejecución.

    - Si el código se ejecuta empaquetado (por ejemplo, como .app o .exe),
      sube desde la ruta del ejecutable hasta la carpeta que contiene el proyecto.
    - Si se ejecuta como script Python normal, sube dos niveles desde
      el archivo actual (../..), asumiendo la estructura estándar del proyecto.

    Devuelve:
        str: Ruta absoluta a la carpeta raíz del proyecto.
    """
    if getattr(sys, "frozen", False):
        macos_dir = os.path.dirname(sys.executable)
        contents_dir = os.path.dirname(macos_dir)
        app_bundle = os.path.dirname(contents_dir)
        return os.path.dirname(app_bundle)
    return os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))

# =====================================================
# 2. Función principal del proceso (procesar)
# =====================================================
def procesar(archivo_remittance,archivo_fbl5n):


    rutas = {
        "remittance": archivo_remittance,
        "fbl5n": archivo_fbl5n,
        # Si necesitas una ruta de salida, puedes definirla aquí:
        "salida": os.path.join(os.path.dirname(archivo_remittance), "Cencosud.xlsx")
    }
    customer_id = 10267301

    # =====================================================
    # 1. Lectura de Remitente
    # =====================================================
    # Leemos la tabla principal del Remittance: (ajustado a PDF) - usando Camelot
    # LECTURA AUTOMÁTICA DEL PDF: DETECCIÓN DE COLUMNAS + CAMELOT

    HEADERS_EXPECTED = [
        "VOUCHER","DESCRIPCION","DOCUMENTO","TIENDA","F. REGISTRO","SECCION","VALOR",
        "FAC. IVA","FAC. RET. FUENTE","RET. IVA","RET. ICA","OTROS IMP.","VALOR PAG","DOC.SOPORTE"
    ]

    def detect_header_positions(pdf_path, page_number=1, headers_expected=None):
        """Devuelve dict {HEADER: x0} usando pdfplumber."""
        if headers_expected is None:
            headers_expected = HEADERS_EXPECTED

        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[page_number - 1]
            words = page.extract_words(use_text_flow=True)

        words_u = [
            (w['text'].strip().upper(), w)
            for w in words if w.get('text') and w['text'].strip() != ""
        ]

        header_positions = {}

        for header in headers_expected:
            h = header.strip().upper()

            # 1) match exacto
            candidates = [w for t, w in words_u if t == h]

            # 2) match substring
            if not candidates:
                candidates = [w for t, w in words_u if h in t]

            # 3) match por primera palabra
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

        cuts = []
        cuts.append(max(1, x_list[0] - 20))  # left margin

        for i in range(len(x_list) - 1):
            mid = int(round((x_list[i] + x_list[i+1]) / 2))
            cuts.append(mid)

        cuts.append(int(round(x_list[-1] + 120)))  # right margin

        return sorted(list(dict.fromkeys(cuts)))


    def camelot_read_with_columns(pdf_path, columns_cuts):
        """Llama a Camelot usando columnas detectadas y genera fallback."""
        cols_str = ",".join(str(x) for x in columns_cuts) if columns_cuts else None
        error = None

        try:
            if cols_str:
                tables = camelot.read_pdf(
                    pdf_path, pages='all', flavor='stream',
                    strip_text='\n', columns=cols_str,
                    row_tol=8, column_tol=6
                )
            else:
                tables = camelot.read_pdf(
                    pdf_path, pages='all', flavor='stream',
                    strip_text='\n', row_tol=8, column_tol=6
                )
            return tables
        except Exception:
            error = traceback.format_exc()
            print("\n⚠️ Error usando columnas detectadas — intentando fallback sin columnas…")

            try:
                return camelot.read_pdf(pdf_path, pages='all', flavor='stream', strip_text='\n')
            except Exception:
                print(error)
                print("\n❌ Fallback también falló.")
                return []


    # === EJECUCIÓN REAL DE LECTURA ===

    header_positions = detect_header_positions(rutas["remittance"], page_number=1)
    cuts = build_cuts_from_positions(header_positions)

    if cuts:
        print("\n✔️ Cortes detectados automáticamente:", cuts)
    else:
        print("\n⚠️ No se detectaron suficientes columnas — usando fallback por defecto.")
        cuts = [
            8,     # VOUCHER
            46,    # DESCRIPCION
            102,   # DOCUMENTO 
            190,   # TIENDA 
            222,   # TIENDA → SECCION
            271,   # SECCION → F. REGISTRO
            318,   # F. REGISTRO → VALOR
            358,   # VALOR → FAC. IVA
            410,   # FAC. IVA → RET. FUENTE
            527,   # RET FUENTE → RET IVA
            661,   # RET IVA → RET ICA
            780    # OTROS IMP → VALOR PAG → DOC SOPORTE
        ]

    tables = camelot_read_with_columns(rutas["remittance"], cuts)

    if len(tables) == 0:
        raise ValueError("Camelot no pudo leer ninguna tabla del PDF.")

    df_all = pd.concat([t.df for t in tables], ignore_index=True)

    # ------------------------------------------------------------
    # FIX ROBUSTO: detectar encabezado, normalizar y separar DOCUMENTO/TIENDA
    # ------------------------------------------------------------

    # 1) localizar la fila que contiene el encabezado real (contiene "VOUCHER")
    header_mask = df_all.apply(lambda r: r.astype(str).str.contains("VOUCHER", case=False).any(), axis=1)
    if not header_mask.any():
        raise ValueError("No se encontró fila de encabezado (VOUCHER) en df_all. Revisa extracción Camelot.")
    header_row_idx = header_mask.idxmax()

    # 2) usar esa fila como encabezado real
    df_all.columns = df_all.iloc[header_row_idx].astype(str).str.replace("\n"," ").str.strip()

    # 3) eliminar todo lo anterior (cabeceras y basura)
    df_all = df_all.drop(index=range(0, header_row_idx + 1)).reset_index(drop=True)

    # 4) limpieza básica de nombres de columnas
    df_all.columns = df_all.columns.str.replace("\n", " ").str.strip()

    # 5) eliminar columnas vacías tipo 'nan' (nombre literal 'nan' o columnas con nombre vacío)
    cols_to_drop = [c for c in df_all.columns if str(c).strip().lower() in ("nan", "", "none")]
    if cols_to_drop:
        df_all = df_all.drop(columns=cols_to_drop)

    # 6) eliminar filas vacías iniciales si las hay
    df_all = df_all.loc[~df_all.apply(lambda r: r.astype(str).str.strip().eq('').all(), axis=1)].reset_index(drop=True)

    # 7) Si Camelot pegó DOCUMENTO+TIENDA en una sóla columna (ej. 'DOCUMENTOTIENDA'), dividirla
    #    Buscamos columnas que contengan 'DOCUMENTO' y 'TIENDA' en el nombre; si existe una que contenga ambos, la partimos.
    merged_col = None
    for c in df_all.columns:
        cu = str(c).upper().replace(" ", "")
        if "DOCUMENTO" in cu and "TIENDA" in cu:
            merged_col = c
            break

    if merged_col is not None:
        # Heurística de separación:
        #  - 1) intentar split por dos o más espacios
        #  - 2) intentar split por tokens típicos de tienda (VPP1,VPP2,ADM.,JUMBO,RANCHO,PERFUME,DROGUE, etc.)
        #  - 3) fallback: split por la última ocurrencia de espacio si el último token es corto (código tienda)
        import re
        merged_series = df_all[merged_col].astype(str)

        def split_documento_tienda(s):
            s = s.replace("\n", " ").strip()
            if s == "" or s.upper() == "NAN":
                return ("", "")
            # 1) doble espacio
            m = re.search(r'\s{2,}', s)
            if m:
                left = s[:m.start()].strip()
                right = s[m.end():].strip()
                return (left, right)
            # 2) patrones de tiendas conocidos
            patterns = [' VPP', ' ADM', ' JUMBO', ' PLAT', ' PLATA', ' RANCHO', ' PERFUME', ' DROGUE', ' CUCU', ' DOCKIN', ' BUCARAM']
            for p in patterns:
                if p in s.upper():
                    idx = s.upper().find(p)
                    # left part before token, right part from token start
                    return (s[:idx].strip(), s[idx:].strip())
            # 3) fallback por último espacio si último token corto (códigos tienda o abreviaturas)
            parts = s.rsplit(' ', 1)
            if len(parts) == 2 and len(parts[1]) <= 10:
                return (parts[0].strip(), parts[1].strip())
            # 4) si no se detecta, devolvemos original como DOCUMENTO y tienda vacía
            return (s, "")

        splits = merged_series.map(split_documento_tienda)
        df_all['DOCUMENTO'] = splits.map(lambda x: x[0])
        df_all['TIENDA'] = splits.map(lambda x: x[1])

        # borrar la columna mergeada original
        df_all = df_all.drop(columns=[merged_col])

    else:
        # 8) Si las columnas existen separadas pero con nombres pegados (ej. 'DOCUMENTOTIENDA' sin espacio),
        # intentamos renombrar por coincidencias parciales: preferimos columnas separadas si están como 2 columnas en una celda.
        # Buscamos nombres parecidos y normalizamos.
        cols_upper = [str(c).upper().replace(" ", "") for c in df_all.columns]
        # Si existe 'DOCUMENTOTIENDA' como nombre exacto
        if 'DOCUMENTOTIENDA' in cols_upper:
            idx = cols_upper.index('DOCUMENTOTIENDA')
            colname = df_all.columns[idx]
            merged_series = df_all[colname].astype(str)
            # reusar la misma heurística de split
            import re
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

    # 9) Normalizar nombres: dejar en mayúsculas sin puntos para las referencias internas
    df_all.columns = [str(c).strip() for c in df_all.columns]

    # 10) Asegurar que las columnas esperadas existan (crearlas vacías si no)
    expected = ["VOUCHER","DESCRIPCION","DOCUMENTO","TIENDA","SECCION","F. REGISTRO","VALOR PAG","DOC.SOPORTE"]
    for e in expected:
        if e not in df_all.columns:
            df_all[e] = ""

    # 11) reindex opcional (orden lógico)
    # mantenemos todas las columnas que vinieron pero aseguramos que DOCUMENTO y TIENDA existan
    cols_order = [c for c in df_all.columns if c not in ("DOCUMENTO","TIENDA")]
    # colocar DOCUMENTO y TIENDA junto a la DESCRIPCION (si existe), intentando una posición coherente
    insert_at = 2 if len(cols_order) >= 2 else len(cols_order)
    if 'DOCUMENTO' in df_all.columns:
        if 'DOCUMENTO' in cols_order:
            cols_order.remove('DOCUMENTO')
    if 'TIENDA' in df_all.columns:
        if 'TIENDA' in cols_order:
            cols_order.remove('TIENDA')
    new_order = cols_order[:insert_at] + ['DOCUMENTO','TIENDA'] + cols_order[insert_at:]
    # Asegurarnos de eliminar duplicados en order
    seen = set()
    new_order = [x for x in new_order if not (x in seen or seen.add(x))]
    df_all = df_all.reindex(columns=new_order)
    
    # ==========================================================
    # FIX PARA FILAS DESALINEADAS CUANDO SECCION COMIENZA CON "PLAT"
    # ==========================================================

    def corregir_filas_plat(df):
        """
        Detecta filas donde SECCION arranca con 'PLAT' pero el patrón
        de columnas está corrido hacia la derecha.

        Corrección:
            TIENDA        <- SECCION
            SECCION       <- F. REGISTRO
            F. REGISTRO   <- VALOR FAC.
            VALOR FAC.    <- IVA FAC.
            IVA FAC.      <- RET. FUENTE
            RET. FUENTE   <- RET. IVA
            RET. IVA      <- RET. ICA
            RET. ICA      <- OTROS IMP.
            OTROS IMP.    <- VALOR PAG
            VALOR PAG     <- DOC.SOPORTE
            DOC.SOPORTE   <- ""   (queda vacío)
        """

        cols = [
            "VOUCHER","DESCRIPCION","DOCUMENTOTIENDA",
            "SECCION","F. REGISTRO","VALOR FAC.","IVA FAC.",
            "RET. FUENTE","RET. IVA","RET. ICA","OTROS IMP.",
            "VALOR PAG","DOC.SOPORTE"
        ]

        for idx, row in df.iterrows():
            sec = str(row.get("SECCION", "")).strip().upper()

            # Detecta patrón PLAT desalineado
            if sec.startswith("PLAT"):

                # Aplicar corrimiento
                df.at[idx, "TIENDA"]        = row["SECCION"]
                df.at[idx, "SECCION"]       = row["F. REGISTRO"]
                df.at[idx, "F. REGISTRO"]   = row["VALOR FAC."]
                df.at[idx, "VALOR FAC."]    = row["IVA FAC."]
                df.at[idx, "IVA FAC."]      = row["RET. FUENTE"]
                df.at[idx, "RET. FUENTE"]   = row["RET. IVA"]
                df.at[idx, "RET. IVA"]      = row["RET. ICA"]
                df.at[idx, "RET. ICA"]      = row["OTROS IMP."]
                df.at[idx, "OTROS IMP."]    = row["VALOR PAG"]
                df.at[idx, "VALOR PAG"]     = row["DOC.SOPORTE"]
                df.at[idx, "DOC.SOPORTE"]   = ""

        return df


    # 🔧 APLICAR FIX AQUÍ:
    df_all = corregir_filas_plat(df_all)
    
    # =====================================================
    # FIX EXTRA — corregir columnas corridas por SECCION pegada a TIENDA
    # =====================================================

    secciones_validas = ["PERFUME", "DROGUE", "PERFUMERIA", "DROGUERIA", "MONTERIA"]

    def corregir_tienda_seccion(row):
        tienda = str(row.get("TIENDA", "")).strip()
        seccion = str(row.get("SECCION", "")).strip()

        # Si la SECCION ya está correctamente ubicada → no tocar
        if seccion in secciones_validas:
            return row

        # Detectar si TIENDA termina en una SECCION válida
        for sec in secciones_validas:
            if tienda.endswith(" " + sec):
                base = tienda[: -(len(sec))].strip()
                row["TIENDA"] = base
                row["SECCION"] = sec

                # DESPLAZAR TODAS LAS COLUMNAS CORRIDAS
                columnas = [
                    "F. REGISTRO", "VALOR FAC.", "IVA FAC.", "RET. FUENTE",
                    "RET. IVA", "RET. ICA", "OTROS IMP.", "VALOR PAG", "DOC.SOPORTE"
                ]

                valores = [row.get(c, "") for c in columnas]

                # Si DOC.SOPORTE vino vacío o tiene basura pero VALOR PAG está desplazado
                # detectamos desplazamiento cuando VALOR PAG es muy grande o raro
                if row.get("DOC.SOPORTE", "") in ["", None] and str(row.get("VALOR PAG", "")).isdigit():
                    # shift left
                    valores = valores[1:] + [""]

                for i, c in enumerate(columnas):
                    row[c] = valores[i]

                return row

        return row

    df_all = df_all.apply(corregir_tienda_seccion, axis=1)

    # Print de columnas
#    print("✔️ Normalización completa. Columnas resultantes:")
#    print(df_all.columns.tolist())

    # =====================================================
    # 4. Limpieza de Remittance
    # =====================================================

    # Tipos válidos de VOUCHER para Cencosud
    filter_values = [
        'DAT','CH','DAV','DCA','DCC','DCF','DEV','DND','DPC','FPM','FS','LTG','RPL','DEC'
    ]

    # Filtrar filas por la primera columna (la del tipo de voucher)
    remittance = df_all[df_all[df_all.columns[0]].isin(filter_values)].copy()
    
    # =====================================================
    # AJUSTE: Completar SECCION cuando viene vacía y viene
    # pegada al final de DOCUMENTO (DROGUE / PERFUME)
    # =====================================================

    secciones_validas = ["DROGUE", "PERFUME"]


    def corregir_seccion(row):
        seccion = str(row.get("SECCION", "")).strip()
        documento = str(row.get("DOCUMENTO", "")).strip()

        # Caso 1: SECCION contiene una fecha → detectar si contiene "/"
        if "/" in seccion:
            for sec in secciones_validas:
                if documento.endswith(sec):
                    row["SECCION"] = sec
                    row["DOCUMENTO"] = documento[: -len(sec)].strip()
                    return row
            # Si tiene fecha pero no coincide con ningún final válido, simplemente vaciar
            row["SECCION"] = ""
            return row

        # Caso 2: SECCION está vacía → la lógica original
        if seccion in ["", None]:
            for sec in secciones_validas:
                if documento.endswith(sec):
                    row["SECCION"] = sec
                    row["DOCUMENTO"] = documento[: -len(sec)].strip()
                    return row

        return row

    remittance = remittance.apply(corregir_seccion, axis=1)
    
    
    # Tranformamos a numerico el dato "VALOR PAG"
    remittance["VALOR PAG"] = (
        remittance["VALOR PAG"]
        .astype(str)
        .str.replace(".", "", regex=False)   # quitar puntos de miles
        .str.replace(",", ".", regex=False)  # convertir coma a decimal
    )
    
    remittance["VALOR PAG"] = pd.to_numeric(remittance["VALOR PAG"], errors="coerce")

    # Parche necesario, dejar. Por lectura de pdf se puede movel el valor de la columna "VALOR PAG" a "DOC.SOPORTE".
    # Esta linea reemplaza los valores de "VALOR PAG" cuando es 0 por el valor de "DOC.SOPORTE"
    remittance.loc[remittance["VALOR PAG"] == 0, "VALOR PAG"] = remittance["DOC.SOPORTE"]
    
    # FIX: mover valores muy grandes de VALOR PAG a DOC.SOPORTE
    umbral = 20000000000000  # 20 billones

    # --- Limpiar y convertir columnas a numérico para comparación ---
    def clean_to_numeric(col):
        return pd.to_numeric(
            col.astype(str).str.replace(".", "", regex=False),
            errors="coerce"
        )

    valor_pag_num = clean_to_numeric(remittance["VALOR PAG"])
    otros_imp_num = clean_to_numeric(remittance["OTROS IMP."])

    # --- Máscara de valores muy grandes ---
    mask_grande = valor_pag_num > umbral

    # Pasar valores grandes a DOC.SOPORTE
    remittance["DOC.SOPORTE"] = ""
    remittance.loc[mask_grande, "DOC.SOPORTE"] = remittance.loc[mask_grande, "VALOR PAG"]

    # Mover OTROS IMP. a VALOR PAG solo si OTROS IMP. != 0
    mask_otros = otros_imp_num != 0
    remittance.loc[mask_grande & mask_otros, "VALOR PAG"] = remittance.loc[mask_grande & mask_otros, "OTROS IMP."]

    # Aseguramos que DOC.SOPORTE sea string y vacío si <= umbral
    remittance["DOC.SOPORTE"] = remittance["DOC.SOPORTE"].apply(
        lambda x: str(x) if clean_to_numeric(pd.Series([x])).iloc[0] > umbral else ""
    )

    # OTROS IMP. siempre = 0
    remittance["OTROS IMP."] = 0

    # Ordenamiento personalizado
    def sort_key(val):
        if val == "VOUCHER": return "0"
        if val == "FPM":     return "1"
        return "2" + str(val)

    remittance["sort_order"] = remittance[remittance.columns[0]].apply(sort_key)

    remittance = (
        remittance.sort_values(by="sort_order")
        .drop(columns="sort_order")
        .drop_duplicates()
        .reset_index(drop=True)
    )


    # 2.1 Guardar Remittance en buffer en memoria
    remittance_buffer = io.BytesIO()
    with pd.ExcelWriter(remittance_buffer, engine="openpyxl") as writer:
        remittance.to_excel(writer, index=False)
    remittance_buffer.seek(0)

    
    # Renombrar columnas de interés
    remittance = remittance.rename(columns={
        "DESCRIPCION": "Tipo de Documento",
        "DOCUMENTO": "Referencia / Factura",
        "VALOR PAG": "Importe de Remittance"
    })
    
    # Renombramos columna "Referencia / Factura" "FACTURA PROVEEDOR" por "Factura"
    remittance.loc[remittance["Tipo de Documento"] == "FACTURA PROVEEDOR", "Tipo de Documento"] = "Factura"

#    # Limpieza de importes
    remittance["Importe de Remittance"] = (
        pd.to_numeric(
            remittance["Importe de Remittance"].astype(str)
            .str.replace(".", "", regex=False)
            .str.replace(",", "", regex=False)
            .str.extract(r"([-\d]+)")[0],
            errors="coerce"
        )
    )
    # Quitamos filas "Importe de Remittance" nulas
    remittance = remittance[remittance["Importe de Remittance"].notna()]
    remittance["Importe de Remittance"] *= -1  # Ajustar signo

    # Asegurar columnas de descuento
    for col in ["Descuento","Motivo del descuento","Comentarios"]:
        if col not in remittance.columns:
            remittance[col] = ""


    # =====================================================
    # 5. Manejo de Reglas y CARDs
    # =====================================================
    # Aquí se aplican todas las reglas de tipo de documento según VOUCHER
    # LTG
    mask = remittance["VOUCHER"] == "LTG"
    remittance.loc[mask, "Tipo de Documento"] = "Descuentos Clientes"
    remittance.loc[mask, "Descuento"] = "Rechazo"
    remittance.loc[mask, "Motivo del descuento"] = "551"
    remittance.loc[mask, "Comentarios"] = remittance.loc[mask, "Descuento"] + " " + remittance.loc[mask, "Referencia / Factura"]

    # Grupo de vouchers: DAT, DAV, DCA, DCC, DCF, DND, DPC, RPC, DCR, RPL
    grupo = ["DAT","DAV","DCA","DCC","DCF","DND","DPC","RPC","DCR", "RPL", "DEC"]
    mask = remittance["VOUCHER"].isin(grupo)
    agrupados = (
        remittance.loc[mask]
        .groupby(["VOUCHER", "Tipo de Documento", "SECCION", "DOC.SOPORTE"], as_index=False)
        .agg({
            "Referencia / Factura": lambda x: " ".join(x.dropna().astype(str)),
            "Importe de Remittance": "sum"
        })
    )
    agrupados["Referencia / Factura"] = agrupados["Tipo de Documento"]
    agrupados["Tipo de Documento"] = "Descuentos Clientes"
    agrupados["Descuento"] = agrupados["SECCION"]
    agrupados["Motivo del descuento"] = "987"
    agrupados["Comentarios"] = agrupados["VOUCHER"] + " DSCTO " + agrupados["DOC.SOPORTE"].fillna("") + " " + agrupados["Referencia / Factura"].fillna("") + " " + agrupados["SECCION"].fillna("")
    remittance = pd.concat([remittance.loc[~mask].copy(), agrupados], ignore_index=True)

    # FS
    mask = remittance["VOUCHER"] == "FS"
    remittance.loc[mask, "Tipo de Documento"] = "Descuentos Clientes"
    remittance.loc[mask, "Descuento"] = "FACT PROVEED"
    remittance.loc[mask, "Motivo del descuento"] = "CSB"
    remittance.loc[mask, "Comentarios"] = remittance.loc[mask, "Referencia / Factura"]

    # DEV
    mask = remittance["VOUCHER"] == "DEV"
    remittance.loc[mask, "Referencia / Factura"]  = remittance.loc[mask, "Tipo de Documento"]
    remittance.loc[mask, "Descuento"] = remittance.loc[mask, "Tipo de Documento"]
    remittance.loc[mask, "Tipo de Documento"] = "Descuentos Clientes"
    remittance.loc[mask, "Motivo del descuento"] = "522"
    remittance.loc[mask, "Comentarios"] = remittance.loc[mask, "VOUCHER"] + " " + remittance.loc[mask, "Referencia / Factura"] + " " + remittance.loc[mask, "SECCION"]

    # CH (MENORES VALORES)
    mask = remittance["VOUCHER"] == "CH"
    remittance.loc[mask, "Referencia / Factura"] = remittance.loc[mask, "Tipo de Documento"]
    remittance.loc[mask, "Descuento"] = remittance.loc[mask, "Tipo de Documento"]
    remittance.loc[mask, "Tipo de Documento"] = "Descuentos Clientes"
    remittance.loc[mask, "Motivo del descuento"] = np.where(
        (remittance.loc[mask, "Importe de Remittance"].abs() >= 20000), "987",
        np.where(remittance.loc[mask, "Importe de Remittance"].between(-20000,0, inclusive="neither"), "WOB","384")
    )
    remittance.loc[mask, "Comentarios"] = "MENORES VALORES"
    
    # =====================================================
    # 6. Procesamiento de columnas 'Descuento' y 'Comentarios'
    # =====================================================
    
    remittance = procesar_descuentos_y_comentarios(remittance)

    # =====================================================
    # 7. Lectura de la Cartera (FBL5N) (datos desde SAP)
    # =====================================================
    # =====================================================
    # 8. Filtro de la cartera del cliente
    # =====================================================
    # =====================================================
    # 9. Renombrado y limpieza de columnas
    # =====================================================
    FBL5N, id_cliente, nombre_cliente = procesar_cartera_cliente(rutas["fbl5n"], customer_id)
    
    # =====================================================
    # 10. Merge Remittance + FBL5N por "Referencia / Factura"
    # =====================================================
    # Se realiza un merge tipo "left" sobre 'Referencia / Factura' para mantener todas
    # las filas de Remittance y añadir información de FBL5N cuando exista coincidencia
    hrc_template = merge_remittance_cartera(remittance, FBL5N)
    

    # =====================================================
    # 11. Cálculo de diferencias
    # =====================================================
    # Se calcula la diferencia entre 'Importe de factura' y 'Importe de Remittance'
    # La lógica centralizada se encuentra en la función procesar_diferencias()
    hrc_template = procesar_diferencias(hrc_template)
    
    # =====================================================
    # 12. Agregamos registros NRO
    # =====================================================
    hrc_template = procesamiento_nro(hrc_template, FBL5N)
    
    # =====================================================
    # 13. Asignación de Pago Neto (Pago Neto = Importe de factura) y otros ajustes
    # =====================================================
    # Por defecto, 'Pago Neto' = 'Importe de factura'
    
    hrc_template["Pago Neto"] = hrc_template["Importe de factura"]
    
    # Elimina las facturas mal cargadas
    hrc_template = hrc_template[
        ~(
            (hrc_template["Tipo de Documento"] == "Factura") &
            (hrc_template["Referencia / Factura"].astype(str).str.len() != 10)
        )
    ].reset_index(drop=True)

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
    
    # Exportamos template final, aplicando formato y copiando hoja de Remittance
    exportar_template(
        hrc_template=hrc_template, 
        suma_remittance = remittance["Importe de Remittance"].sum(),
        numero_orden="",   # FALTA: parametrizar
        id_cliente=id_cliente,
        nombre_cliente=nombre_cliente,
        ruta_remittance=remittance_buffer,  # buffer en memoria
        ruta_salida=rutas["salida"]
    )
    
    # Devolución del template final
    return hrc_template
