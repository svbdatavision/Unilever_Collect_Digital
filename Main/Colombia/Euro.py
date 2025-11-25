# =====================================================
# 0. Importación de librerías y módulos utilitarios
# =====================================================
 
import os
import sys
import re
import numpy as np
import pandas as pd
import camelot
from PyPDF2 import PdfReader
from openpyxl import load_workbook
import io
import unicodedata
import pdfplumber, re, statistics, pandas as pd

from utils import *  # Importación de funciones utilitarias del paquete interno 'clientes'

 
# Configuración de advertencias
# warnings.filterwarnings("ignore", category=UserWarning, module="camelot") # Suprime advertencias generadas por Camelot (usualmente por manejo de PDFs)
 
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
    """
    Orquestador principal para Cencosud
    """
    root = _project_root()

    # --- Rutas de entrada/salida ---
    rutas = {
        "remittance": archivo_remittance,
        "fbl5n": archivo_fbl5n,
        # Si necesitas una ruta de salida, puedes definirla aquí:
        "salida": os.path.join(os.path.dirname(archivo_remittance), "Euro.xlsx")
    }
    # Colocar el Customer ID del cliente
    customer_id = 10309611
    
    results = []
    # =====================================================
    # 1. Lectura de Remitente
    # =====================================================
    # --- Funciones de limpieza y extracción ---
    def prefer_doc(doccol, co, detalle):
        """
        Busca prefijo PMP/DEC/PM/DE + número en cualquier lugar de la fila
        """
 
        full_txt = ' '.join(filter(None, [doccol, co, detalle])).replace('\n', ' ')
        prefix_match = re.search(r'\b(PMP|DEC|PM|DE|DEV)\b', full_txt, flags=re.IGNORECASE)
        num_match    = re.search(r'\b(\d{3,10})\b', full_txt.replace(',', ''))
        if prefix_match and num_match:
            return f"{prefix_match.group(1).upper()} {num_match.group(1)}"
        if num_match:
            return num_match.group(1)
        return ''
 
    def clean_co(co_raw):
        if not co_raw: return ''
        toks = str(co_raw).split()
        for t in toks:
            if re.match(r'^[A-Za-z]{2,5}$', t):
                return t
        return toks[0] if toks else ''
 
    def clean_detalle(co_raw, doccol, detalle_raw, doc_found):
        comb = str(detalle_raw or '').replace('\n', ' ').strip()
        if doc_found:
            comb = re.sub(r'\b' + re.escape(doc_found) + r'\b', ' ', comb)
        comb = re.sub(r'\b\d{3,10}\b', ' ', comb)
        comb = re.sub(r'\s+', ' ', comb).strip()
        return comb
 
    def extract_num_columns(row, num_cols=['Descuentos','Retenciones','Valor Factura','Valor Pago']):
        """
        Extrae todos los números de la fila (incluyendo $-) y los asigna en orden a num_cols
        """
        text = ' '.join([str(row.get(col,'')).replace('\n',' ').replace('\r',' ').strip() for col in num_cols])
        # Buscar números con o sin signo, incluyendo $-
        nums = re.findall(r'\$?-?\d[\d,\.]*', text)
        nums = [n.replace('$','').replace(' ','') for n in nums]
        result = {}
        for i, col in enumerate(num_cols):
            result[col] = nums[i] if i < len(nums) else ''
        return result
 
    def merge_doc_prefix_rows(parsed, doc_field='Doc.Cruce'):
        merged = []
        i = 0
        while i < len(parsed):
            row = parsed[i].copy()
            doc = str(row.get(doc_field, '')).strip()
            if re.match(r'^[A-Za-z]{2,4}$', doc, flags=re.IGNORECASE) and i + 1 < len(parsed):
                nxt = parsed[i+1]
                nxt_doc = str(nxt.get(doc_field, '')).strip()
                if re.match(r'^\d{3,10}$', nxt_doc):
                    for k in list(row.keys()):
                        if k == 'page':
                            continue
                        row[k] = (str(row.get(k, '')).strip() + ' ' + str(nxt.get(k, '')).strip()).strip()
                    try:
                        row['raw_top'] = min(float(row.get('raw_top', 0)), float(nxt.get('raw_top', row.get('raw_top', 0))))
                    except Exception:
                        pass
                    merged.append(row)
                    i += 2
                    continue
            merged.append(row)
            i += 1
        return merged
 
    # --- Lectura del PDF ---
    with pdfplumber.open(rutas["remittance"]) as pdf:
        for pnum, page in enumerate(pdf.pages, start=1):
            words = page.extract_words()
            if not words:
                continue
 
            # Detectar encabezados
            header_cands = []
            for w in words:
                t = re.sub(r'\s+', ' ', w['text']).strip()
                tl = t.lower().replace(' ', '')
                if tl in ('reg','reg.') or 'detalle' in tl or 'doc' in tl or 'cruce' in tl or tl in ('c','c.','o','o.') \
                or tl in ('descuentos','retenciones','valorfactura','valorpago'):
                    header_cands.append({'text': t, 'x0': w['x0'], 'top': w['top']})
            if not header_cands:
                continue
 
            # Agrupar encabezados por top
            header_cands.sort(key=lambda x: x['top'])
            clusters = []
            tol = 4
            for h in header_cands:
                if not clusters or abs(h['top'] - clusters[-1]['top_mean']) > tol:
                    clusters.append({'items':[h], 'top_mean': h['top']})
                else:
                    clusters[-1]['items'].append(h)
                    clusters[-1]['top_mean'] = statistics.mean([it['top'] for it in clusters[-1]['items']])
            cluster = max(clusters, key=lambda c: c['top_mean'])
            header_items = sorted(cluster['items'], key=lambda it: it['x0'])
            header_top = cluster['top_mean']
 
            # Calcular posiciones de columnas
            positions = {}
            co_x_list, doc_x_list = [], []
            for it in header_items:
                tnorm = it['text'].strip().lower().replace(' ','')
                if 'reg' in tnorm:
                    positions['Reg.'] = it['x0']
                elif 'detalle' in tnorm:
                    positions['Detalle'] = it['x0']
                elif tnorm in ('c','c.','o','o.'):
                    co_x_list.append(it['x0'])
                elif 'doc' in tnorm or 'cruce' in tnorm:
                    doc_x_list.append(it['x0'])
                elif 'descuentos' in tnorm:
                    positions['Descuentos'] = it['x0']
                elif 'retenciones' in tnorm:
                    positions['Retenciones'] = it['x0']
                elif 'valorfactura' in tnorm:
                    positions['Valor Factura'] = it['x0']
                elif 'valorpago' in tnorm:
                    positions['Valor Pago'] = it['x0']
 
            if not (positions.get('Reg.') and positions.get('Detalle') and doc_x_list):
                continue
 
            positions['C.O.'] = statistics.mean(co_x_list) if co_x_list else (positions['Reg.'] + min(doc_x_list))/2
            positions['Doc.Cruce'] = statistics.mean(doc_x_list)
 
            centers_sorted = sorted(positions.items(), key=lambda x: x[1])
            colnames = [c[0] for c in centers_sorted]
            xs = [c[1] for c in centers_sorted]
 
            # Crear bounds de columnas
            bounds = [0.0] + [(a+b)/2.0 for a,b in zip(xs, xs[1:])] + [page.width + 1.0]
            padding = 5
            for i, col in enumerate(colnames):
                if col in ('Descuentos','Retenciones','Valor Factura','Valor Pago'):
                    bounds[i] -= padding
                    bounds[i+1] += padding
 
            # Palabras debajo del encabezado
            data_words = [w for w in words if w['top'] > header_top - 2]
            data_words.sort(key=lambda w: (w['top'], w['x0']))
 
            # Agrupar por fila
            rows = []
            current, last_top = [], None
            row_tol = 8
            for w in data_words:
                if last_top is None or abs(w['top'] - last_top) <= row_tol:
                    current.append(w)
                    last_top = w['top'] if last_top is None else (last_top + w['top'])/2.0
                else:
                    rows.append(current)
                    current = [w]
                    last_top = w['top']
            if current:
                rows.append(current)
 
            # Asignar palabras a columnas
            parsed = []
            for row_words in rows:
                cells = {name: [] for name in colnames}
                for w in row_words:
                    x = w['x0']
                    col_idx = next((i for i in range(len(bounds)-1) if x >= bounds[i] and x < bounds[i+1]), None)
                    if col_idx is None:
                        col_idx = min(range(len(bounds)-1), key=lambda i: abs((bounds[i]+bounds[i+1])/2 - x))
                    cells[colnames[col_idx]].append(w['text'])
                row_dict = {'page': pnum, 'raw_top': statistics.mean([w['top'] for w in row_words])}
                for name in colnames:
                    row_dict[name] = ' '.join(cells[name]).replace('\n',' ').replace('\r',' ').strip()
                parsed.append(row_dict)
 
            # Unir filas donde el prefijo PMP/DEC está separado
            parsed = merge_doc_prefix_rows(parsed, doc_field='Doc.Cruce')
 
            # Filtrar filas válidas
            finals = []
            for r in parsed:
                reg_val = str(r.get('Reg.', '')).strip()
                co_val = str(r.get('C.O.', '')).strip()
                doc_val = str(r.get('Doc.Cruce', '')).strip()
                if re.match(r'^\s*\d+\b', reg_val) or re.match(r'^\s*\d+\b', co_val):
                    finals.append(r)
                    continue
                if re.search(r'\b([A-Za-z]{2,4})\s*\d{3,10}\b', doc_val):
                    finals.append(r)
                    continue
                if re.search(r'\b(PMP|PM|DEC|DE)\b', doc_val, flags=re.IGNORECASE):
                    finals.append(r)
                    continue
 
            # Corregir caso C.O.
            for r in finals:
                if not re.match(r'^\d+\b', str(r.get('Reg.','')).strip()) and re.match(r'^\d+\b', str(r.get('C.O.','')).strip()):
                    m = re.match(r'^(\d+)\s*(.*)$', r['C.O.'])
                    if m:
                        r['Reg.'] = m.group(1)
                        r['C.O.'] = m.group(2).strip()
 
            results.extend(finals)
 
    # --- Paso final: crear DataFrame ---
    df = pd.DataFrame(results)
 
    # Unir filas multilínea
    grouped_rows = []
    buffer = None
    num_cols = ['Descuentos','Retenciones','Valor Factura','Valor Pago']
    for r in df.to_dict(orient='records'):
        reg_val = r.get('Reg.', '').strip()
        co_val = r.get('C.O.', '').strip()
        if reg_val or co_val:
            if buffer:
                grouped_rows.append(buffer)
            buffer = r.copy()
        else:
            if buffer is None:
                buffer = r.copy()
            else:
                for col in ['Detalle','Doc.Cruce','C.O.'] + num_cols:
                    buffer[col] = (buffer.get(col,'') + ' ' + r.get(col,'')).strip()
    if buffer:
        grouped_rows.append(buffer)
 
    # Crear DataFrame final con números correctamente extraídos
    final_rows = []
    for r in grouped_rows:
        reg = str(r.get('Reg.', '')).strip()
        co_raw = r.get('C.O.', '') or ''
        doccol = r.get('Doc.Cruce', '') or ''
        detalle_raw = r.get('Detalle', '') or ''
        doc = prefer_doc(doccol, co_raw, detalle_raw)
        co = clean_co(co_raw)
        detalle = clean_detalle(co_raw, doccol, detalle_raw, doc)
        nums = extract_num_columns(r, num_cols)
       
 
        final_rows.append({
    #        'page': r.get('page', ''),
            'Reg.': reg,
            'C.O.': co,
            'Doc.Cruce': doc,
            'Detalle': detalle,
            'Descuentos': nums['Descuentos'],
            'Retenciones': nums['Retenciones'],
            'Valor Factura': nums['Valor Factura'],
            'Valor Pago': nums['Valor Pago']
        })
       
    df_final = pd.DataFrame(final_rows)
 
    # Parche
    mask = (~df_final["Descuentos"].str.contains(',', na=False)) & (df_final["Retenciones"].str.contains(',', na=False))
    df_final.loc[mask, ["Descuentos", "Retenciones"]] = df_final.loc[mask, ["Retenciones", "Descuentos"]].values
 
    # 2.1 Guardar Remittance en buffer en memoria
    remittance_buffer = io.BytesIO()
    with pd.ExcelWriter(remittance_buffer, engine="openpyxl") as writer:
        df_final.to_excel(writer, index=False)
    remittance_buffer.seek(0)
 
    # =====================================================
    # 4. Limpieza de Remittance
    # =====================================================
   
    # Replace 'PMP ' and 'PM ' with 'PMP' in the 'Doc.Cruce' column
    df_final['Doc.Cruce'] = df_final['Doc.Cruce'].str.replace('PMP ', 'PMP', regex=False)
    df_final['Doc.Cruce'] = df_final['Doc.Cruce'].str.replace('PM ', 'PMP', regex=False)
 
    # Clean and convert the 'Valor Pago' column to float
    df_final['Valor Pago'] = (
        df_final['Valor Pago']
        .astype(str)
        .str.replace(',', '', regex=False)
        .str.replace(' ', '', regex=False)
        .str.replace('−', '-', regex=False)  # Handle special minus sign if present
        .astype(float)
    )
 
    #Create the new column 'Tipo de Documento' based on the conditions
    df_final['Tipo de Documento'] = df_final.apply(
        lambda row: 'Factura' if row['C.O.'] == 'CED' and row['Valor Pago'] > 0 else 'Descuento Cliente',
        axis=1
    )
 
    # Apply the transformation to 'Doc.Cruce' based on 'Tipo de Documento'
    df_final['Doc.Cruce'] = df_final.apply(
        lambda row: str(row['Doc.Cruce'])[:10] if row['Tipo de Documento'] == 'Factura' else row['Doc.Cruce'],
        axis=1
    )
    # Rename the column 'Valor Pago' to 'Importe de Remittance'
    df_final.rename(columns={"Valor Pago": "Importe de Remittance"}, inplace=True)
    df_final.rename(columns={"Doc.Cruce": "Referencia / Factura"}, inplace=True)
 
    # =====================================================
    # 5. Manejo de Reglas y CARDs
    # =====================================================
    # Ensure the columns exist
    if "Tipo de Documento" in df_final.columns and "C.O." in df_final.columns:
        # Create empty columns if they don't exist
        if "Descuento" not in df_final.columns:
            df_final["Descuento"] = ""
        if "Motivo del descuento" not in df_final.columns:
            df_final["Motivo del descuento"] = ""
        if "Descrición" not in df_final.columns:
            df_final["Comentarios"] = ""
 
        # Define conditions only for rows where 'Tipo de Documento' is 'Descuento Cliente'
        mask_descuento = df_final["Tipo de Documento"] == "Descuento Cliente"
        co_values = ["MAY", "BEL", "MIX", "VEG", "FLO", "CAR", "SAL", "NUM", "LOB", "FRO",
                    "TER", "SAB", "ACA", "LAU", "ITA", "GUA", "LLA", "MUR", "MON", "ROS", "CED"]
 
        # Create condition list for 'C.O.' prefixes
        conds_desc = [df_final["C.O."].str.startswith(prefix, na=False) & mask_descuento for prefix in co_values]
 
        # Define corresponding values
        descuentos = ["DESCUENTO"] * len(conds_desc)
        motivos = ["987"] * len(conds_desc)
 
        df_final["Descuento"] = np.select(conds_desc, descuentos, default = df_final["Descuento"])
        df_final["Motivo del descuento"] = np.select(conds_desc, motivos, default = df_final["Motivo del descuento"])
        mask = df_final["Tipo de Documento"] == "Descuento Cliente"
        df_final.loc[mask, "Comentarios"] = (
            df_final.loc[mask, "Descuento"].astype(str) + " " +
            df_final.loc[mask, "C.O."].astype(str) + " " +
            df_final.loc[mask, "Referencia / Factura"].astype(str)
        )
       
    remittance = df_final
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