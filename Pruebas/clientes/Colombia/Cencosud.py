import os
import sys
import re
import warnings
import camelot
import pdfplumber
import pandas as pd
import numpy as np
import traceback
import io
from clientes.utils import *

warnings.filterwarnings("ignore", category=UserWarning, module="camelot")

# =====================================================
# Ruta raíz del proyecto
# =====================================================
def _project_root():
    if getattr(sys, "frozen", False):
        macos_dir = os.path.dirname(sys.executable)
        contents_dir = os.path.dirname(macos_dir)
        app_bundle = os.path.dirname(contents_dir)
        return os.path.dirname(app_bundle)
    return os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))

# =====================================================
# 2. Función principal del proceso 
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
    # 1. Lectura de Remitente
    # =====================================================
    # ------------------------
    # Configuración y patrones
    # ------------------------

    # Vouchers reconocidos
    VOUCHERS = r"(FS|CH|DEC|NC|ND|LTG|FPM|DCA|DCF|DND|DAV|DCC|RPL|DAT|DEV|DPC)"
    pat_inicio = re.compile(rf"^{VOUCHERS}\b", flags=re.IGNORECASE)

    # Patrones para documentos específicos
    pat_documento_dec = re.compile(r"\b(\d{4}-\d{7,20})\b")     # DEC
    pat_documento_pmp = re.compile(r"\b(PMP\d{3,20})\b")        # PMP / LTG / FPM
    pat_documento_fs_pair = re.compile(r"\b([A-Z0-9]{2,6})\s+(\d{5,12})\b")  # FS
    pat_fecha = re.compile(r"\d{2}/\d{2}/\d{4}")                # Fecha

    # Secciones posibles
    SECCIONES = {"PERFUME", "DROGUE", "RANCHO", "PLATOS"}

    # -----------------------
    # Funciones utilitarias
    # -----------------------

    def normalize_whitespace(s: str) -> str:
        """Quita espacios extra y normaliza la cadena."""
        return re.sub(r"\s+", " ", s).strip()

    def is_start_of_record(line: str) -> bool:
        """Verifica si una línea comienza con un voucher reconocido."""
        return bool(pat_inicio.match(line.strip()))

    def merge_lines(lines):
        """
        Une todas las líneas de un mismo registro:
        - Si empieza con voucher → nuevo registro
        - Si no, se agrega al registro actual
        Esto asegura que multi-líneas en DESCRIPCION o TIENDA queden juntas.
        """
        registros = []
        buffer = ""

        for raw in lines:
            line = raw.strip()
            if not line:
                continue

            if pat_inicio.match(line):
                if buffer:
                    registros.append(normalize_whitespace(buffer))
                buffer = line
            else:
                buffer += " " + line  # unir toda línea que no empieza con voucher

        if buffer:
            registros.append(normalize_whitespace(buffer))
        return registros

    def extract_numeric_values_after_date(text: str):
        """Extrae todos los valores numéricos después de la fecha en la línea."""
        t = re.sub(r"[^\d\.,\-]+", " ", text)
        toks = [x for x in t.split() if x.strip() != ""]
        return [tk for tk in toks if re.search(r"\d", tk)]

    # ------------------------
    # Parsers por tipo de VOUCHER
    # ------------------------

    def parse_DEC(line, voucher):
        """Parsea registros tipo DEC, incluyendo especiales 'DTO POR ESCALA VOLU'."""
        line = normalize_whitespace(line)
        mfecha = pat_fecha.search(line)
        if not mfecha:
            return None
        fecha = mfecha.group(0)

        pre_date = line[len(voucher):mfecha.start()].strip()
        post_date = line[mfecha.end():].strip()

        prefix = "DTO POR ESCALA VOLU"
        if pre_date.startswith(prefix):
            descripcion = prefix
            rest = pre_date[len(prefix):].strip()
            # Capturar documento DEC seguido de todo lo que queda como tienda
            mdoc = re.match(r"(\d{4}-\d{7,20})(.*)", rest)
            if mdoc:
                documento = mdoc.group(1)
                tienda = normalize_whitespace(mdoc.group(2))
            else:
                documento = ""
                tienda = rest
        else:
            # Caso genérico DEC
            mdoc = pat_documento_dec.search(pre_date)
            if mdoc:
                documento = mdoc.group(1)
                descripcion = normalize_whitespace(pre_date[:mdoc.start()])
                tienda = normalize_whitespace(pre_date[mdoc.end():])
            else:
                documento = ""
                descripcion = pre_date
                tienda = ""

        # Extraer valores numéricos después de la fecha
        nums = extract_numeric_values_after_date(post_date)
        while len(nums) < 8:
            nums.append("0")
        valor, iva, ret_fuente, ret_iva, ret_ica, otros_imp, valor_pag = nums[:7]
        doc_soporte = nums[7]

        # Detectar sección dentro de DESCRIPCION o TIENDA
        seccion = ""
        for sec in SECCIONES:
            if sec in descripcion:
                seccion = sec
                descripcion = descripcion.replace(sec, "").strip()
                break
            elif sec in tienda:
                seccion = sec
                tienda = tienda.replace(sec, "").strip()
                break

        return {
            "VOUCHER": voucher,
            "DESCRIPCION": descripcion,
            "DOCUMENTO": documento,
            "TIENDA": tienda,
            "SECCION": seccion.upper(),
            "F. REGISTRO": fecha,
            "VALOR": valor,
            "IVA": iva,
            "RET. FUENTE": ret_fuente,
            "RET. IVA": ret_iva,
            "RET. ICA": ret_ica,
            "OTROS IMP.": otros_imp,
            "VALOR PAG": valor_pag,
            "DOC.SOPORTE": doc_soporte
        }

    def parse_LTG_FPM(line, voucher):
        """Parsea vouchers tipo LTG y FPM."""
        m = pat_documento_pmp.search(line)
        if not m:
            return None
        documento = m.group(1)
        doc_start, doc_end = m.start(), m.end()
        descripcion = normalize_whitespace(line[len(voucher):doc_start])

        tail = line[doc_end:].strip()
        tokens_tail = tail.split()
        idx_sec = None
        for i, t in enumerate(tokens_tail):
            if t.upper() in SECCIONES:
                idx_sec = i
                break

        if idx_sec is None:
            mfecha = pat_fecha.search(tail)
            if not mfecha:
                return None
            fecha = mfecha.group(0)
            before_date = tail.split(fecha, 1)[0].strip()
            btoks = before_date.split()
            if btoks and btoks[-1].upper() in SECCIONES:
                seccion = btoks[-1]
                tienda = " ".join(btoks[:-1]).strip()
            else:
                tienda = before_date
                seccion = ""
        else:
            tienda = " ".join(tokens_tail[:idx_sec]).strip()
            seccion = tokens_tail[idx_sec]

        mfecha = pat_fecha.search(tail)
        fecha = mfecha.group(0)

        nums = extract_numeric_values_after_date(tail.split(fecha,1)[1])
        while len(nums) < 8:
            nums.append("0")
        valor, iva, ret_fuente, ret_iva, ret_ica, otros_imp, valor_pag = nums[:7]
        doc_soporte = nums[7]

        return {
            "VOUCHER": voucher,
            "DESCRIPCION": descripcion,
            "DOCUMENTO": documento,
            "TIENDA": tienda,
            "SECCION": seccion,
            "F. REGISTRO": fecha,
            "VALOR": valor,
            "IVA": iva,
            "RET. FUENTE": ret_fuente,
            "RET. IVA": ret_iva,
            "RET. ICA": ret_ica,
            "OTROS IMP.": otros_imp,
            "VALOR PAG": valor_pag,
            "DOC.SOPORTE": doc_soporte
        }

    def parse_FS(line, voucher):
        """Parsea vouchers tipo FS."""
        m = pat_documento_fs_pair.search(line)
        if not m:
            return None
        documento = f"{m.group(1)} {m.group(2)}"
        doc_start = m.start()
        descripcion = normalize_whitespace(line[len(voucher):doc_start])

        tail = line[m.end():].strip()
        mfecha = pat_fecha.search(tail)
        if not mfecha:
            return None
        fecha = mfecha.group(0)
        before_date = tail.split(fecha,1)[0].strip()

        adm_idx = re.search(r"\bADM\.", before_date, flags=re.IGNORECASE)
        tienda = before_date[adm_idx.start():].strip() if adm_idx else before_date

        nums = extract_numeric_values_after_date(tail.split(fecha,1)[1])
        while len(nums) < 8:
            nums.append("0")
        valor, iva, ret_fuente, ret_iva, ret_ica, otros_imp, valor_pag = nums[:7]
        doc_soporte = nums[7]

        return {
            "VOUCHER": voucher,
            "DESCRIPCION": descripcion,
            "DOCUMENTO": documento,
            "TIENDA": tienda,
            "SECCION": "",
            "F. REGISTRO": fecha,
            "VALOR": valor,
            "IVA": iva,
            "RET. FUENTE": ret_fuente,
            "RET. IVA": ret_iva,
            "RET. ICA": ret_ica,
            "OTROS IMP.": otros_imp,
            "VALOR PAG": valor_pag,
            "DOC.SOPORTE": doc_soporte
        }

    def parse_CH(line, voucher):
        """Parsea vouchers tipo CH o registros con sección fija ("DCA", "DCF", "DND", "DAV", "DCC", "RPL")."""
        mfecha = pat_fecha.search(line)
        if not mfecha:
            return None
        fecha = mfecha.group(0)
        pre_date = line[len(voucher):mfecha.start()].strip()
        adm_idx = re.search(r"\bADM\.", pre_date, flags=re.IGNORECASE)
        if adm_idx:
            descripcion = normalize_whitespace(pre_date[:adm_idx.start()])
            tienda = pre_date[adm_idx.start():].strip()
        else:
            descripcion = pre_date
            tienda = ""

        # Extraer valores numéricos
        nums = extract_numeric_values_after_date(line[mfecha.end():])
        while len(nums) < 8:
            nums.append("0")
        valor, iva, ret_fuente, ret_iva, ret_ica, otros_imp, valor_pag = nums[:7]
        doc_soporte = nums[7]

        return {
            "VOUCHER": voucher,
            "DESCRIPCION": descripcion,
            "DOCUMENTO": "",
            "TIENDA": tienda,
            "SECCION": "",
            "F. REGISTRO": fecha,
            "VALOR": valor,
            "IVA": iva,
            "RET. FUENTE": ret_fuente,
            "RET. IVA": ret_iva,
            "RET. ICA": ret_ica,
            "OTROS IMP.": otros_imp,
            "VALOR PAG": valor_pag,
            "DOC.SOPORTE": doc_soporte
        }

    def parse_generic(line, voucher):
        """Intento de parseo genérico si no entra en los tipos anteriores."""
        for fn in (parse_DEC, parse_LTG_FPM, parse_FS, parse_CH):
            r = fn(line, voucher)
            if r:
                return r
        # Fallback
        mfecha = pat_fecha.search(line)
        if not mfecha:
            return None
        fecha = mfecha.group(0)
        descripcion = normalize_whitespace(line[len(voucher):mfecha.start()])
        nums = extract_numeric_values_after_date(line[mfecha.end():])
        while len(nums) < 8:
            nums.append("0")
        valor, iva, ret_fuente, ret_iva, ret_ica, otros_imp, valor_pag = nums[:7]
        doc_soporte = nums[7]
        return {
            "VOUCHER": voucher,
            "DESCRIPCION": descripcion,
            "DOCUMENTO": "",
            "TIENDA": "",
            "SECCION": "",
            "F. REGISTRO": fecha,
            "VALOR": valor,
            "IVA": iva,
            "RET. FUENTE": ret_fuente,
            "RET. IVA": ret_iva,
            "RET. ICA": ret_ica,
            "OTROS IMP.": otros_imp,
            "VALOR PAG": valor_pag,
            "DOC.SOPORTE": doc_soporte
        }

    # -----------------------
    # Dispatcher principal
    # -----------------------
    def parse_record(line: str):
        line = normalize_whitespace(line)
        m_start = re.match(rf"^{VOUCHERS}\b", line, flags=re.IGNORECASE)
        if not m_start:
            return None
        voucher = m_start.group(1).upper()

        if voucher in {"DCA", "DCF", "DND", "DAV", "DCC", "RPL"}:
            return parse_CH(line, voucher)  # sección fija
        if voucher == "DEC":
            return parse_DEC(line, voucher)
        if voucher in {"LTG", "FPM"}:
            return parse_LTG_FPM(line, voucher)
        if voucher == "FS":
            return parse_FS(line, voucher)
        if voucher == "CH":
            return parse_CH(line, voucher)

        return parse_generic(line, voucher)

    # ------------------------
    # Lectura y parseo PDF
    # ------------------------
    records = []
    with pdfplumber.open(rutas["remittance"]) as pdf:
        for page in pdf.pages:
            txt = page.extract_text()
            if not txt:
                continue
            lines = txt.split("\n")
            # Saltar encabezado
            for i, l in enumerate(lines):
                if l.strip().startswith("VOUCHER"):
                    lines = lines[i+1:]
                    break
            merged = merge_lines(lines)
            for rec_line in merged:
                parsed = parse_record(rec_line)
                if parsed:
                    records.append(parsed)

    # ------------------------
    # Crear DataFrame
    # ------------------------
    cols = [
        "VOUCHER","DESCRIPCION","DOCUMENTO","TIENDA","SECCION",
        "F. REGISTRO","VALOR FAC.","IVA FAC.","RET. FUENTE","RET. IVA",
        "RET. ICA","OTROS IMP.","VALOR PAG","DOC.SOPORTE"
    ]
    df = pd.DataFrame(records)
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    df = df[cols]

    # ------------------------
    # Separar DOCUMENTO + TIENDA pegados (especial DEC)
    # ------------------------
    pat_merged = re.compile(r"^(\d{4}-\d{7,20})([A-ZÁÉÍÓÚÑ0-9 .-]+)$", flags=re.IGNORECASE)
    for idx, row in df.iterrows():
        doc = str(row["DOCUMENTO"]).strip()
        tienda = str(row["TIENDA"]).strip()
        if not doc and tienda:
            m = pat_merged.match(tienda.replace(" ", ""))
            if m:
                df.at[idx, "DOCUMENTO"] = m.group(1)
                df.at[idx, "TIENDA"] = normalize_whitespace(m.group(2))

    # ------------------------
    # Extraer SECCION pegada en DESCRIPCION o TIENDA
    # ------------------------
    secciones_sorted = sorted(SECCIONES, key=lambda s: -len(s))

    def extract_and_move_section_from_text(text: str):
        if not isinstance(text, str):
            return text, ""
        txt_low = text.lower()
        for sec in secciones_sorted:
            sec_low = sec.lower()
            idx = txt_low.rfind(sec_low)
            if idx != -1:
                seccion = text[idx:idx+len(sec)]
                new_text = re.sub(r"\s+", " ", (text[:idx] + text[idx+len(sec):]).strip())
                return new_text, seccion.upper()
        return text, ""

    def fix_row_move_section(row):
        desc = str(row.get("DESCRIPCION", "")).strip()
        tienda = str(row.get("TIENDA", "")).strip()
        seccion_actual = str(row.get("SECCION", "")).strip()

        new_desc, found = extract_and_move_section_from_text(desc)
        if found:
            row["DESCRIPCION"] = new_desc
            row["SECCION"] = found
            tienda = tienda.replace(found, "").strip()
            row["TIENDA"] = re.sub(r"\s+", " ", tienda)
            return row

        new_tienda, found = extract_and_move_section_from_text(tienda)
        if found:
            row["TIENDA"] = new_tienda
            row["SECCION"] = found
            return row

        row["DESCRIPCION"] = re.sub(r"\s+", " ", desc)
        row["TIENDA"] = re.sub(r"\s+", " ", tienda)
        row["SECCION"] = seccion_actual
        return row

    df = df.apply(fix_row_move_section, axis=1)

    # Validar SECCION
    df["SECCION"] = df["SECCION"].fillna("").str.upper().str.strip()
    df.loc[~df["SECCION"].isin(SECCIONES), "SECCION"] = ""
    df["TIENDA"] = df["TIENDA"].astype(str).str.strip()

    # PARCHES
    # Parche para casos puntuales por no poder leer registros multi-línea
    for idx, row in df.iterrows():
        # DESCRIPCION específica para RPL
        if row["VOUCHER"] == "RPL" and row["DESCRIPCION"].strip().upper() == "DESCUENTO":
            df.at[idx, "DESCRIPCION"] = "DESCUENTO COMERCIAL"

        # Ajuste de SECCION
        if row["SECCION"].strip().upper() == "DROGUE":
            df.at[idx, "SECCION"] = "DROGUER"

        # Ajustes específicos de TIENDA
        tienda_val = row["TIENDA"].strip().upper()
        if tienda_val == "PLAT - CROSS":
            df.at[idx, "TIENDA"] = "PLAT - CROSS DOCKIN"
        elif tienda_val == "PLAT - PLAT":
            df.at[idx, "TIENDA"] = "PLAT - PLAT BUCARAM"
        elif tienda_val == "PLAT - CROSSD":
            df.at[idx, "TIENDA"] = "PLAT - CROSSD AVERI"
            
    # Creacion de remittance
    remittance = df
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

    # Extraer números de forma robusta
    remittance["Importe de Remittance"] = (
        remittance["Importe de Remittance"]
            .astype(str)
            .str.replace(".", "", regex=False)
            .str.replace(",", "", regex=False)
            .str.extract(r"([-\d]+)", expand=False)  # expand=False → devuelve Series, NO DataFrame
    )

    # Convertir a numérico sin romper
    remittance["Importe de Remittance"] = pd.to_numeric(
        remittance["Importe de Remittance"],
        errors="coerce"
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
    
    # Normalizar columna REFERENCIA / FACTURA
    # 1. Reemplaza valores no escalares por string seguro
    hrc_template["Referencia / Factura"] = (
        hrc_template["Referencia / Factura"]
            .apply(lambda x: "" if x is None else str(x))  # convierte todo a texto
            .str.replace(r"[\[\]\(\)\{\}]", "", regex=True)  # remueve secuencias tipo lista/array
            .str.replace(r"\s+", " ", regex=True)           # normaliza espacios
            .str.strip()
    )
    # Eliminar facturas mal cargadas
    # (luego de normalizar y asegurar scalar strings)
    hrc_template = (
        hrc_template[
            ~(
                (hrc_template["Tipo de Documento"] == "Factura") &
                (hrc_template["Referencia / Factura"].str.len() != 10)
            )
        ]
        .reset_index(drop=True)
    )

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