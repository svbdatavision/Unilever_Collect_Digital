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
        "salida": os.path.join(os.path.dirname(archivo_remittance), "Cenocsud.xlsx")
    }
    #root = _project_root()

    # --- Rutas de entrada/salida ---
    #rutas = {
     #   "pdf_remittance": os.path.join(root, "Archivos", "Remittance", "Colombia", "Remittance_Cenco.pdf"),  # Cencosud
#        "fbl5n": os.path.join(root, "Archivos", "Cartera", "FBL5N.xlsx"),
      #  "fbl5n": os.path.join(root, "Archivos", "Cartera", "FBL5N_Cenco.xlsx"),
       # "salida": os.path.join(root, "Archivos", "Template", "Colombia", "Template_HRC_Cenco.xlsx")
    #}
    # Colocar el Customer ID del cliente
    customer_id = 10267301

    # =====================================================
    # 1. Lectura de Remitente
    # =====================================================
    reader = PdfReader(rutas["remittance"])
    num_pages = len(reader.pages)
    all_pages = set(range(1, num_pages + 1))

    
    # ------------------------------
# LECTURA ROBUSTA DEL REMITTANCE
# ------------------------------
print("Intentando extracción con Camelot (stream) agresivo...")
try:
    tables_stream = camelot.read_pdf(
        rutas["remittance"],
        pages="all",
        flavor="stream",
        row_tol=15,
        edge_tol=300,
        strip_text="\n",
        split_text=True,
        flag_size=True
    )
    df_all = pd.concat([t.df for t in tables_stream], ignore_index=True)
    print("Camelot stream devolvió filas:", df_all.shape[0])
except Exception as e:
    print("Camelot stream falló:", e)
    df_all = pd.DataFrame()

# Si la extracción de Camelot no es coherente (pocas filas o muchas filas con headers repetidos),
# usamos el fallback por texto (más robusto a líneas partidas).
need_fallback = False
if df_all.empty:
    need_fallback = True
else:
    # heurística simple: si hay muchas filas que parecen encabezado repetido o VALOR PAG vacío, fallback
    n_header_like = df_all.apply(lambda r: r.astype(str).str.upper().str.contains("DESCRIPCION|VALOR PAG|DOCUMENTO").any(), axis=1).sum()
    if n_header_like / max(1, len(df_all)) > 0.6:
        need_fallback = True

if need_fallback:
    print("Usando fallback por texto (parsing heurístico).")
    reader = PdfReader(rutas["remittance"])
    pages_text = [p.extract_text() or "" for p in reader.pages]

    # Tokens VOUCHER que indican inicio de registro (ajusta si necesitás más)
    voucher_tokens = {"DAT","CH","DAV","DCA","DCC","DCF","DEV","DND","DPC","FPM","FS","LTG","RPL","VOUCHER"}

    records = []
    for pageno, text in enumerate(pages_text, start=1):
        # limpiar saltos múltiples y dividir en líneas
        lines = [ln.strip() for ln in text.splitlines() if ln.strip() != ""]
        # unir líneas que claramente continúan (heurística): si una línea NO empieza por voucher token o por fecha/valor
        buffer = []
        for ln in lines:
            # Si la línea empieza con un voucher token (p.ej. "FPM ") -> inicio de nuevo registro
            first_word = ln.split()[0] if len(ln.split())>0 else ""
            if first_word in voucher_tokens:
                # volcar buffer como registro si existe
                if buffer:
                    records.append(" ".join(buffer))
                buffer = [ln]
            else:
                # continuación de la línea previa -> anexar
                if buffer:
                    buffer.append(ln)
                else:
                    # no hay buffer: iniciar (por si el primer registro no lleva voucher en la misma línea)
                    buffer = [ln]
        if buffer:
            records.append(" ".join(buffer))

    # Ahora parseamos cada record con regex heurísticos
    parsed = []
    date_re = r"\b\d{2}/\d{2}/\d{4}\b"
    doc_re = r"\b[Pp][A-Z0-9]*\d{4,}\b"  # ej. PMP1284460 (ajusta si hay otros formatos)
    num_re = r"-?\d[\d\.,]*"  # número con puntos y comas
    for rec in records:
        # Inicializar campos
        voucher = None
        descripcion = None
        documento = None
        seccion = None
        fecha = None
        valores = []

        # voucher: primer token si está en la lista
        parts = rec.split()
        if parts and parts[0] in voucher_tokens:
            voucher = parts[0]
            rest = " ".join(parts[1:])
        else:
            # buscar voucher anywhere
            m_v = re.search(r"\b(" + "|".join(voucher_tokens) + r")\b", rec)
            voucher = m_v.group(1) if m_v else ""
            rest = rec

        # descripcion (buscar palabras clave comunes)
        m_desc = re.search(r"(FACTURA PROVEEDOR|FACTURA VENTA|COSTO DE TRANSFEREN|COSTO DE TRANSFERENCIA|DIF\. COSTO\/CANTIDAD|DESCRIPCION)", rest, flags=re.IGNORECASE)
        if m_desc:
            descripcion = m_desc.group(0).strip().upper()
        else:
            # fallback: tomar las primeras 3-4 palabras como descripción
            descripcion = " ".join(rest.split()[:3]).upper()

        # documento: buscar patrón PMP... u otro alfanumérico seguido de dígitos
        m_doc = re.search(doc_re, rec)
        if m_doc:
            documento = m_doc.group(0).strip().upper()

        # fecha
        m_fecha = re.search(date_re, rec)
        if m_fecha:
            fecha = m_fecha.group(0)

        # extraer todos los números (orden relativo a layout)
        valores = re.findall(num_re, rec)
        # limpiar valores (quitar casos de voucher token mezclado etc.)
        valores_clean = []
        for v in valores:
            v2 = v.replace(".", "").replace(",", "")
            # si v2 es sólo '-' o empty -> skip
            if re.search(r"\d", v2):
                valores_clean.append(v)

        # heurística para asignar cantidades (ajustar si necesario)
        # asumimos que los últimos 2 números suelen ser VALOR PAG y DOC.SOPORTE (o solo VALOR PAG)
        valor_fac = None
        iva_fac = None
        valor_pag = None
        doc_soporte = None
        if len(valores_clean) >= 4:
            # tomar los últimos 4 como [VALOR FAC., IVA, OTROS..., VALOR PAG, DOC.SOPORTE] -> esto depende del layout
            # estrategia simple: asignar desde el final
            doc_soporte = valores_clean[-1]
            valor_pag = valores_clean[-2]
            # opcionales:
            iva_fac = valores_clean[-3] if len(valores_clean) >= 3 else None
            valor_fac = valores_clean[-4] if len(valores_clean) >= 4 else None
        elif len(valores_clean) == 3:
            doc_soporte = valores_clean[-1]
            valor_pag = valores_clean[-2]
            iva_fac = valores_clean[-3]
        elif len(valores_clean) == 2:
            valor_pag = valores_clean[-1]
            doc_soporte = valores_clean[-1]
        elif len(valores_clean) == 1:
            valor_pag = valores_clean[0]

        parsed.append({
            "VOUCHER": voucher or "",
            "DESCRIPCION": descripcion or "",
            "DOCUMENTO": documento or "",
            "SECCION": "",      # se puede rellenar por heurística posterior
            "F. REGISTRO": fecha or "",
            "VALOR FAC.": valor_fac or "",
            "IVA FAC.": iva_fac or "",
            "OTROS IMP.": "",
            "VALOR PAG": valor_pag or "",
            "DOC.SOPORTE": doc_soporte or "",
            "PAGE": None
        })

    # crear dataframe desde parsed
    df_all = pd.DataFrame(parsed)
    # normalizar columnas a mayúsculas
    df_all.columns = [c.upper() for c in df_all.columns]
    print("Fallback creó registros:", df_all.shape[0])

# A partir de aquí df_all contiene las filas (vía Camelot o fallback)
# Normalizamos nombres de columna para el flujo que sigue
df_all = df_all.rename(columns=lambda c: c.strip().upper())
# Si quedó DOCUMENTO combinado con TIENDA, forzamos DOCUMENTO (ya lo normalizamos arriba)
df_all.columns = df_all.columns.str.replace("DOCUMENTO TIENDA", "DOCUMENTO", regex=False)

"""    # --- STREAM ---
    tables_stream = camelot.read_pdf(
        rutas["remittance"],
        pages="all",
        flavor="stream",
        row_tol=15,
        edge_tol=300,
        strip_text="\n",
        split_text=True,
        flag_size=True
    )

    stream_pages = set(int(t.page) for t in tables_stream)
    missing_pages = all_pages - stream_pages

    # --- LATTICE PARA PÁGINAS QUE FALTEN ---
    tables_lattice = []
    if missing_pages:
        missing_pages_str = ",".join(str(p) for p in missing_pages)
        tables_lattice = camelot.read_pdf(
            rutas["remittance"], pages=missing_pages_str, flavor='lattice', strip_text='\n'
        )

    # --- UNIR TODAS LAS TABLAS ---
    df_all = pd.concat(
        [t.df for t in list(tables_stream) + list(tables_lattice)],
        ignore_index=True
    )

    # DETECTAR FILA DE ENCABEZADO REAL (DOCUMENTO o VOUCHER)
    header_idx_candidates = df_all.apply(
        lambda row: row.astype(str).str.contains("DOCUMENTO", case=False).any()
                 or row.astype(str).str.contains("VOUCHER", case=False).any(),
        axis=1
    )

    if not header_idx_candidates.any():
        raise ValueError(
            f"No se detectó encabezado válido.\nPrimeras filas:\n{df_all.head()}"
        )

    header_row_idx = header_idx_candidates.idxmax()

    # CONVERTIR FILA A ENCABEZADO REAL
    df_all.columns = df_all.iloc[header_row_idx]
    df_all = df_all.drop(index=list(range(0, header_row_idx + 1))).reset_index(drop=True)

    # LIMPIAR COLUMNAS NULAS (ANTES DE NORMALIZAR)
    df_all = df_all.loc[:, ~df_all.columns.isna()]

    # NORMALIZAR TODAS LAS COLUMNAS
    df_all.columns = (
        df_all.columns.astype(str)
        .str.replace(r"[\n\r]+", " ", regex=True)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
        .str.upper()
        .str.replace("DOCUMENTOTIENDA", "DOCUMENTO", regex=False) # Estos codigos están como seguro.
        .str.replace("DOCUMENTO TIENDA", "DOCUMENTO", regex=False)
#        .str.replace("DOCUMENTO-TIENDA", "DOCUMENTO", regex=False)
            # Normalizar todas las variantes de VALOR PAG
        .str.replace("VALOR PAG.", "VALOR PAG", regex=False)
#        .str.replace("VALOR PAGO", "VALOR PAG", regex=False)
#        .str.replace("VALOR_PAG", "VALOR PAG", regex=False)
#        .str.replace("VALOR   PAG", "VALOR PAG", regex=False)
    )
    """
# =====================================================
# UNIFICAR FILAS PARTIDAS DE FACTURA PROVEEDOR (FPM)
# =====================================================

    def es_num(x):
        try:
            float(str(x).replace(".", "").replace(",", "").replace("-", ""))
            return True
        except:
            return False

    rows_to_drop = []

    for i in range(len(df_all) - 1):
        row = df_all.iloc[i]
        next_row = df_all.iloc[i + 1]

        # Detectar fila partida
        cond = (
            (str(row.get("DESCRIPCION", "")) == "FACTURA PROVEEDOR") &
            (str(row.get("VOUCHER", "")) == "FPM") &
            (str(row.get("VALOR PAG", "0")).strip() in ["0", "0.0", ""]) &
            es_num(next_row.get("DOC.SOPORTE", ""))
        )

        if cond:
            # mover VALOR PAG desde DOC.SOPORTE de la fila siguiente
            df_all.at[i, "VALOR PAG"] = next_row["DOC.SOPORTE"]

            # completar SECCION, FECHA y otros si están vacíos
            for col in ["SECCION", "F. REGISTRO"]:
                if str(row.get(col, "")).strip() == "":
                    df_all.at[i, col] = next_row.get(col, "")

            rows_to_drop.append(i + 1)

    # eliminar filas sobrantes
    df_all = df_all.drop(rows_to_drop).reset_index(drop=True)

    # Asegurarnos que ambas columnas existan
    if "VALOR PAG" in df_all.columns and "DOC.SOPORTE" in df_all.columns:

        # Convertir a numérico de forma segura (solo en este paso)
        tmp_valor = (
            df_all["VALOR PAG"]
            .astype(str)
            .str.replace(".", "", regex=False)
            .str.replace(",", "", regex=False)
#            .str.extract(r"([-\d]+)")[0]
            .astype(float)
            .fillna(0)
        )

        tmp_soporte = (
            df_all["DOC.SOPORTE"]
            .astype(str)
            .str.replace(".", "", regex=False)
            .str.replace(",", "", regex=False)
#            .str.extract(r"([-\d]+)")[0]
            .astype(float)
            .fillna(0)
        )

    # Reemplazar SOLO donde VALOR PAG es 0 y DOC.SOPORTE trae algo válido
    mask_replace = (tmp_valor == 0) & (tmp_soporte != 0)

    df_all.loc[mask_replace, "VALOR PAG"] = df_all.loc[mask_replace, "DOC.SOPORTE"]

    # RENOMBRAR LA COLUMNA DOCUMENTO AUN SI VIENE DISTINTA
    col_documento = [c for c in df_all.columns if re.search(r"\bDOCUMENTO\b", c)] # 2do seguro.

    if len(col_documento) == 0:
        raise ValueError(
            f"No se encontró columna DOCUMENTO. Columnas: {list(df_all.columns)}"
        )

    df_all = df_all.rename(columns={col_documento[0]: "DOCUMENTO"})
    

    # =====================================================
    # 4. Limpieza de Remittance
    # =====================================================
    # Definimos los tipos de VOUCHER que nos interesan para Cencosud.
    filter_values = ['DAT','CH','DAV','DCA','DCC','DCF','DEV','DND','DPC','FPM','FS','LTG','RPL','VOUCHER']
    
    # Filtramos únicamente las filas que contienen estos códigos
    remittance = df_all[df_all[df_all.columns[0]].isin(filter_values)].copy()
    

    # 4 Ordenar Remittance para asegurar consistencia
    def sort_key(val):
        if val == "VOUCHER": return "0"
        if val == "FPM":     return "1"
        return "2" + str(val)
    # Aplicamos la función de ordenamiento a cada fila, generando una nueva columna temporal
    remittance["sort_order"] = remittance[remittance.columns[0]].apply(sort_key)
    
    # Ordenamos el DataFrame según la columna temporal y limpiamos duplicados
    remittance = (remittance.sort_values(by="sort_order")
                  .drop(columns="sort_order")
                  .drop_duplicates()
                  .reset_index(drop=True))

    # 2.1 Guardar Remittance en buffer en memoria
    remittance_buffer = io.BytesIO()
    with pd.ExcelWriter(remittance_buffer, engine="openpyxl") as writer:
        remittance.to_excel(writer, index=False)
    remittance_buffer.seek(0)

    # Fuerzo que existe la columna DOCUMENTO, ya lo normalizamos arriba - Funciona como CHequeo
    if "DOCUMENTO" not in remittance.columns:
        posibles = [c for c in remittance.columns if "DOCUMENTO" in c]
        if posibles:
            remittance = remittance.rename(columns={posibles[0]: "DOCUMENTO"})
            
    # Fix robusto por si Camelot devolvió variantes de VALOR PAG
    if "VALOR PAG" not in remittance.columns:
        posibles_valor = [c for c in remittance.columns if "VALOR" in c and "PAG" in c]
        if posibles_valor:
            remittance = remittance.rename(columns={posibles_valor[0]: "VALOR PAG"})

    # Renombrar columnas de interés
    remittance = remittance.rename(columns={
        "DESCRIPCION": "Tipo de Documento",
        "DOCUMENTO": "Referencia / Factura",
        "VALOR PAG": "Importe de Remittance"
    })

    
    # Renombramos columna "Referencia / Factura" "FACTURA PROVEEDOR" por "Factura"
    remittance.loc[remittance["Tipo de Documento"] == "FACTURA PROVEEDOR", "Tipo de Documento"] = "Factura"

    # Limpieza de importes
    remittance["Importe de Remittance"] = (
        pd.to_numeric(
            remittance["Importe de Remittance"].astype(str)
            .str.replace(".", "", regex=False)
            .str.replace(",", "", regex=False)
#            .str.extract(r"([-\d]+)")[0],
            errors="coerce"
        )
    )
    # Quitamos filas "Importe de Remittance" nulas
    remittance = remittance[remittance["Importe de Remittance"].notna()]
#    remittance["Importe de Remittance"] *= -1  # Ajustar signo
    # FILTRO: eliminar descuentos con importe 0

    remittance = remittance[
        ~(
            (remittance["Importe de Remittance"] == 0)
        )
    ].copy()

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
    grupo = ["DAT","DAV","DCA","DCC","DCF","DND","DPC","RPC","DCR", "RPL"]
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
    FBL5N = procesar_cartera_cliente(rutas["fbl5n"], customer_id)
    
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
    # Extraemos id_cliente y nombre_cliente desde el primer registro de FBL5N
    fbl5n_meta = pd.read_excel(rutas["fbl5n"], usecols=["Customer", "Name 1"], nrows=1)
    id_cliente = fbl5n_meta["Customer"].iloc[0] if not fbl5n_meta.empty else ""
    nombre_cliente = fbl5n_meta["Name 1"].iloc[0] if not fbl5n_meta.empty else ""
    
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
