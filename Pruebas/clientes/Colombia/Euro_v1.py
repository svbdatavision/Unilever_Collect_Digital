# =====================================================
# 0. Importación de librerías y módulos utilitarios
# =====================================================

import os
import sys
import re
import io
import statistics
import numpy as np
import pandas as pd
import pdfplumber
from openpyxl import load_workbook

# Importar utilidades locales (tu paquete)
from clientes.utils import *

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
    Orquestador principal para Cencosud (versión robusta para Remittance EURO).
    """
    root = _project_root()

    # --- Rutas de entrada/salida ---
    rutas = {
        "pdf_remittance": os.path.join(root, "Archivos", "Remittance", "Colombia", "Remittance_euro.pdf"),
        "fbl5n": os.path.join(root, "Archivos", "Cartera", "FBL5N_Euro.xlsx"),
        "salida": os.path.join(root, "Archivos", "Template", "Colombia", "Template_HRC_euro.xlsx")
    }

    customer_id = 10309611

    # -----------------------
    # Helpers de texto/header
    # -----------------------
    def normalize_header_token(t):
        # quitar guiones finales que indican palabra partida y espacios al final
        return re.sub(r'[-\u2011]\s*$', '', str(t)).strip()

    def collect_header_candidates(words):
        candidates = []
        for w in words:
            txt = w.get('text','')
            # Normalizamos para detectar tokens parcialmente partidos
            tnorm = re.sub(r'\s+','', txt).lower().replace('­','')
            # busco substrings importantes (acepta variantes)
            if any(k in tnorm for k in ['reg','detalle','doc','cruce','descuentos','retenciones','valorfactura','valorpago','c.o.','c.o','c.','o.','c']):
                candidates.append({'text': txt, 'x0': w.get('x0',0), 'x1': w.get('x1',0), 'top': w.get('top',0), 'bottom': w.get('bottom', w.get('top',0))})
        return candidates

    def cluster_by_top(header_cands, tol=6):
        header_cands = sorted(header_cands, key=lambda x: x['top'])
        clusters = []
        for h in header_cands:
            if not clusters or abs(h['top'] - clusters[-1]['top_mean']) > tol:
                clusters.append({'items':[h], 'top_mean': h['top']})
            else:
                clusters[-1]['items'].append(h)
                clusters[-1]['top_mean'] = statistics.mean([it['top'] for it in clusters[-1]['items']])
        return max(clusters, key=lambda c: len(c['items'])) if clusters else None

    def build_positions(header_items):
        """
        Devuelve un dict con centros aproximados de cada columna detectada.
        Maneja tokens partidos porque normaliza y une.
        """
        positions = {}
        co_x_list, doc_x_list = [], []
        for it in header_items:
            t_raw = it.get('text','')
            t = normalize_header_token(t_raw)
            tl = re.sub(r'\s+','', t).lower()
            center = (it.get('x0',0) + it.get('x1',0)) / 2.0
            if 'reg' in tl:
                positions['Reg.'] = center
            elif 'detalle' in tl:
                positions['Detalle'] = center
            elif tl in ('c','c.','o','o.') or tl.startswith('c.o'):
                co_x_list.append(center)
            elif 'doc' in tl or 'cruce' in tl:
                doc_x_list.append(center)
            elif 'descuentos' in tl or 'descuen' in tl:
                positions['Descuentos'] = center
            elif 'retenciones' in tl or 'retenc' in tl:
                positions['Retenciones'] = center
            elif 'valorfactura' in tl or 'valor' in tl and 'fact' in tl:
                positions['Valor Factura'] = center
            elif 'valorpago' in tl or 'pago' in tl:
                positions['Valor Pago'] = center

        if co_x_list:
            positions['C.O.'] = statistics.mean(co_x_list)
        if doc_x_list:
            positions['Doc.Cruce'] = statistics.mean(doc_x_list)

        return positions

    # ---------------------------
    # Extracción NUMÉRICA robusta
    # ---------------------------
    def extract_num_columns_from_row(row_words, positions=None, colnames_num=None):
        """
        Detecta tokens numéricos dentro de row_words,
        filtra solo números financieros reales,
        reconstruye fragmentos y asigna a:
        Descuentos, Retenciones, Valor Factura, Valor Pago
        """
        num_cols = ["Descuentos", "Retenciones", "Valor Factura", "Valor Pago"]
        if colnames_num is None:
            colnames_num = num_cols

        # 1) Extraer tokens numéricos REALES (formato valor monetario)
        tokens = []
        for w in row_words:
            txt = str(w.get("text", "")).strip()

            if not re.search(r'\d', txt):
                continue

            # Formatos válidos:
            is_monetary = (
                re.search(r'\d{1,3}(,\d{3})+', txt) or
                re.search(r'-\d', txt) or
                re.search(r'\d{5,}', txt)
            )

            if not is_monetary:
                continue

            xcenter = (w.get("x0", 0) + w.get("x1", 0)) / 2.0
            tokens.append({"text": txt, "x": xcenter, "w": w})

        if not tokens:
            return {c: "" for c in num_cols}

        # 2) Normalizar y agrupar fragmentos de un MISMO número
        norm = []
        i = 0
        while i < len(tokens):
            cur = tokens[i]
            txt = cur["text"].replace("−", "-").replace("$", "")
            txt = re.sub(r"[^\d,\.-]", "", txt)
            x = cur["x"]
            merged = txt
            j = i + 1

            while j < len(tokens):
                nxt = tokens[j]["text"].replace("−", "-").replace("$", "")
                nxt = re.sub(r"[^\d,\.-]", "", nxt)

                # Detectar número completo → no unir
                is_complete_number = (
                    re.search(r"\d{1,3}(,\d{3})+", nxt) or
                    re.search(r"-\d", nxt) or
                    re.search(r"\d{5,}", nxt)
                )
                if is_complete_number:
                    break

                # Unir solo si son fragmentos MUY cercanos
                if abs(tokens[j]["x"] - x) <= 12:
                    merged = merged + nxt
                    x = (x + tokens[j]["x"]) / 2.0
                    j += 1
                else:
                    break

            norm.append((x, merged))
            i = j

        # 3) Limpiar ensamblado
        cleaned = []
        for x, t in norm:
            t = t.replace(" ", "")
            t = re.sub(r'-{2,}', '-', t)
            t = re.sub(r'^[\.,]+', '', t)
            t = re.sub(r'[^0-9\-,\.]', '', t)
            t = t.strip()
            if t and t not in ("-", "--"):
                cleaned.append((x, t))

        if not cleaned:
            return {c: "" for c in num_cols}

        # 4) Ordenar por posición (izq → der)
        cleaned.sort(key=lambda it: it[0])

        # 5) Asignación final
        result = {c: "" for c in num_cols}

        # Caso ideal: ya hay 4 números financieros
        if len(cleaned) >= 4:
            for idx, col in enumerate(num_cols):
                result[col] = cleaned[idx][1]
            return result

        # Si hay menos de 4 → asignar por proximidad a centers detectados
        if positions:
            centers = {c: positions[c] for c in num_cols if c in positions}
            if centers:
                for x, val in cleaned:
                    best = min(centers.keys(), key=lambda c: abs(centers[c] - x))
                    if not result[best]:
                        result[best] = val
                    else:
                        # buscar siguiente columna libre
                        idx = num_cols.index(best)
                        placed = False
                        for k in range(idx + 1, len(num_cols)):
                            if not result[num_cols[k]]:
                                result[num_cols[k]] = val
                                placed = True
                                break
                        if not placed:
                            result[best] = result[best] + val  # caso extremo
                return result

        # fallback L→R si todo lo demás falla
        for i, (_, val) in enumerate(cleaned):
            if i < len(num_cols):
                result[num_cols[i]] = val

        return result





    # ---------------------------
    # Extracción de filas y textos
    # ---------------------------
    def extract_rows_from_page(page, positions, header_top):
        words = page.extract_words(use_text_flow=True)
        data_words = [w for w in words if w['top'] > header_top - 2]
        data_words.sort(key=lambda w: (w['top'], w['x0']))

        # Agrupar por fila (por top)
        rows = []
        current, last_top = [], None
        row_tol = 8
        for w in data_words:
            if last_top is None or abs(w['top'] - last_top) <= row_tol:
                current.append(w)
                last_top = w['top'] if last_top is None else (last_top + w['top']) / 2.0
            else:
                rows.append(current)
                current = [w]
                last_top = w['top']
        if current:
            rows.append(current)

        # Determinar orden de columnas por centers detectados
        centers_sorted = sorted(positions.items(), key=lambda x: x[1])
        colnames = [c[0] for c in centers_sorted]

        parsed = []
        for row_words in rows:
            # Extraer números robustamente (usa positions para fallback si faltan números)
            num_values = extract_num_columns_from_row(row_words, positions=positions)

            # Construir bounds para asignación de textos (igual que antes)
            xs = [c[1] for c in centers_sorted] if centers_sorted else []
            bounds = [0.0] + [(a+b)/2.0 for a,b in zip(xs, xs[1:])] + [page.width + 1.0]
            padding = 5
            for i, col in enumerate(colnames):
                if col in ('Descuentos','Retenciones','Valor Factura','Valor Pago'):
                    bounds[i] -= padding
                    bounds[i+1] += padding

            cells = {name: [] for name in colnames}
            for w in row_words:
                x = (w['x0'] + w['x1']) / 2.0
                col_idx = None
                if len(bounds) > 1:
                    try:
                        col_idx = next((i for i in range(len(bounds)-1) if x >= bounds[i] and x < bounds[i+1]), None)
                    except Exception:
                        col_idx = None
                if col_idx is None:
                    # fallback al centro más cercano
                    if centers_sorted:
                        col_idx = min(range(len(centers_sorted)), key=lambda i: abs(centers_sorted[i][1] - x))
                    else:
                        col_idx = 0
                cells[colnames[col_idx]].append(w['text'])

            row_dict = {'page': page.page_number, 'raw_top': statistics.mean([w['top'] for w in row_words])}
            for name in colnames:
                if name in ('Descuentos','Retenciones','Valor Factura','Valor Pago'):
                    row_dict[name] = num_values.get(name, '').strip()
                else:
                    row_dict[name] = ' '.join(cells[name]).replace('\n',' ').replace('\r',' ').strip()
            parsed.append(row_dict)

        return parsed

    # ---------------------------
    # Merge doc prefix helper (mantener)
    # ---------------------------
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

    # ---------------------------
    # Procesamiento de PDF (principal)
    # ---------------------------
    def process_pdf(path):
        results = []
        with pdfplumber.open(path) as pdf:
            for pnum, page in enumerate(pdf.pages, start=1):
                words = page.extract_words(use_text_flow=True)
                if not words:
                    continue

                # Detectar encabezados
                header_cands = collect_header_candidates(words)
                if not header_cands:
                    continue

                cluster = cluster_by_top(header_cands, tol=6)
                if not cluster:
                    continue

                header_items = sorted(cluster['items'], key=lambda it: it['x0'])
                header_top = cluster['top_mean']
                positions = build_positions(header_items)
                
                # --- Normalizar y asegurar columnas numéricas clave ---
                # (pegar justo después de positions = build_positions(header_items))

                # 1) Normalizar keys con lower/strip y unir tokens hyphenados en header_items
                def _normalize_key(k):
                    if k is None:
                        return k
                    s = str(k).lower().replace('\n', '').replace('\r','').replace(' ', '')
                    s = s.replace('-', '')  # para 'Val or Fac-' o 'Valorfac-tura' -> 'valorfactura'
                    return s

                # Re-map de posibles variants a las keys canónicas
                mapping_variants = {
                    'descuentos': 'Descuentos',
                    'retenciones': 'Retenciones',
                    'valorfactura': 'Valor Factura',
                    'valorfact': 'Valor Factura',
                    'valorpago': 'Valor Pago',
                    'pago': 'Valor Pago',
                    'detalle': 'Detalle',
                    'reg': 'Reg.',
                    'c.o.': 'C.O.',
                    'doc': 'Doc.Cruce'
                }

                # Reescalar keys detectadas (por si build_positions devolvió algo raro)
                fixed_positions = {}
                for k, v in positions.items():
                    kn = _normalize_key(k)
                    # buscar variante conocida
                    match = None
                    for var, canon in mapping_variants.items():
                        if var in kn:
                            match = canon
                            break
                    if match:
                        fixed_positions[match] = v
                    else:
                        # si no hay match, mantener el original label (capitalizado)
                        fixed_positions[k] = v

                positions = fixed_positions

                # 2) Si falta "Valor Factura", intentar inferirla:
                if "Valor Factura" not in positions:
                    # si tenemos Retenciones y Valor Pago → la ubicamos en el medio
                    if "Retenciones" in positions and "Valor Pago" in positions:
                        positions["Valor Factura"] = (positions["Retenciones"] + positions["Valor Pago"]) / 2.0
                    # si solo Retenciones → la colocamos a la derecha por ~80 px
                    elif "Retenciones" in positions:
                        positions["Valor Factura"] = positions["Retenciones"] + 80.0
                    # si solo Valor Pago → la colocamos a la izquierda por ~80 px
                    elif "Valor Pago" in positions:
                        positions["Valor Factura"] = positions["Valor Pago"] - 80.0
                    # fallback razonable: después de Descuentos
                    elif "Descuentos" in positions:
                        positions["Valor Factura"] = positions["Descuentos"] + 140.0
                    else:
                        # último recurso: una posición por defecto
                        positions["Valor Factura"] = max(positions.values()) + 80.0 if positions else 400.0

                # 3) Reordenar positions por coordenada X (garantiza L->R independientemente de cómo aparezcan los headers)
                positions = dict(sorted(positions.items(), key=lambda kv: float(kv[1])))

                # 4) Preparar lista ordenada de columnas según X (para usar después)
                col_ordered = list(positions.keys())

                # Opcional: asegurarnos que las 4 numéricas estén en col_ordered en algún orden
                NUM_COLS = ["Descuentos", "Retenciones", "Valor Factura", "Valor Pago"]
                # Si alguna no está en col_ordered, la añadimos al final
                for c in NUM_COLS:
                    if c not in col_ordered:
                        col_ordered.append(c)

                # Ahora `positions` tiene centros confiables y `col_ordered` el orden L->R.
                # Usá `positions` para asignar por proximidad y `col_ordered` para iterar en orden.

                # Reordenar columnas según X real
                positions = dict(sorted(positions.items(), key=lambda x: x[1]))

                
                # FIX — añadir columnas faltantes

                num_cols = ["Descuentos","Retenciones","Valor Factura","Valor Pago"]

                # Si el PDF no trae "Valor Factura", lo creamos
                if "Valor Factura" not in positions:
                    if "Retenciones" in positions and "Valor Pago" in positions:
                        positions["Valor Factura"] = (positions["Retenciones"] + positions["Valor Pago"]) / 2
                    elif "Retenciones" in positions:
                        positions["Valor Factura"] = positions["Retenciones"] + 60
                    else:
                        positions["Valor Factura"] = positions.get("Descuentos", 300) + 120


                # Requiere al menos Reg., Detalle y Doc/CO
                if not (positions.get('Reg.') and positions.get('Detalle') and (positions.get('Doc.Cruce') or positions.get('C.O.'))):
                    continue

                parsed = extract_rows_from_page(page, positions, header_top)
                parsed = merge_doc_prefix_rows(parsed, doc_field='Doc.Cruce')

                finals = []
                for r in parsed:
                    reg_val = str(r.get('Reg.', '')).strip()
                    co_val = str(r.get('C.O.', '')).strip()
                    doc_val = str(r.get('Doc.Cruce', '')).strip()
                    if re.match(r'^\s*\d+\b', reg_val) or re.match(r'^\s*\d+\b', co_val):
                        finals.append(r); continue
                    if re.search(r'\b([A-Za-z]{2,4})\s*\d{3,10}\b', doc_val):
                        finals.append(r); continue
                    if re.search(r'\b(PMP|PM|DEC|DE|DEV)\b', doc_val, flags=re.IGNORECASE):
                        finals.append(r); continue
                results.extend(finals)

        # DataFrame provisional (cada fila ya trae las 4 columnas numéricas posiblemente vacías)
        df = pd.DataFrame(results)

        # ---------------------------
        # Agrupar filas multilínea (mejor control sobre numéricas)
        # ---------------------------
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
                    # concatenar campos textuales
                    for col in ['Detalle','Doc.Cruce','C.O.']:
                        buffer[col] = (buffer.get(col,'') + ' ' + r.get(col,'')).strip()
                    # columnas numéricas: preferir valor no vacío de la línea de continuación
                    for col in num_cols:
                        val_next = (r.get(col,'') or '').strip()
                        if val_next:
                            if not buffer.get(col):
                                buffer[col] = val_next
                            else:
                                if buffer.get(col) in ("", "0", "0.0"):
                                    buffer[col] = val_next
                                else:
                                    # raro: mantener ambos concatenados para revisión manual si ocurre
                                    buffer[col] = (str(buffer.get(col)) + " " + val_next).strip()
        if buffer:
            grouped_rows.append(buffer)

        # ---------------------------
        # Limpiar y armar filas finales
        # ---------------------------
        final_rows = []
        for r in grouped_rows:
            reg = str(r.get('Reg.', '')).strip()
            co_raw = r.get('C.O.', '') or ''
            doccol = r.get('Doc.Cruce', '') or ''
            detalle_raw = r.get('Detalle', '') or ''

            # prefer_doc / clean_co / clean_detalle (mismo comportamiento previo)
            def prefer_doc(doccol, co, detalle):
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

            doc = prefer_doc(doccol, co_raw, detalle_raw)
            co = clean_co(co_raw)
            detalle = clean_detalle(co_raw, doccol, detalle_raw, doc)

            # recoger nums ya presentes en r (pueden venir vacíos)
            nums = {c: (r.get(c,'') or '').strip() for c in ['Descuentos','Retenciones','Valor Factura','Valor Pago']}
            # normalizar guiones repetidos
            for k,v in nums.items():
                nums[k] = re.sub(r'^-+','-', str(v)) if v else ''

            final_rows.append({
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

        # heurístico swap si algunas filas aparecen invertidas (parche existente)
        if not df_final.empty and "Descuentos" in df_final.columns and "Retenciones" in df_final.columns:
            mask = (~df_final["Descuentos"].astype(str).str.contains(',', na=False)) & (df_final["Retenciones"].astype(str).str.contains(',', na=False))
            if mask.any():
                df_final.loc[mask, ["Descuentos", "Retenciones"]] = df_final.loc[mask, ["Retenciones", "Descuentos"]].values

        # normalizar Doc.Cruce
        if 'Doc.Cruce' in df_final.columns:
            df_final['Doc.Cruce'] = df_final['Doc.Cruce'].str.replace('PMP ', 'PMP', regex=False)
            df_final['Doc.Cruce'] = df_final['Doc.Cruce'].str.replace('PM ', 'PMP', regex=False)

        # Normalizar valores numericos en columnas (quitar $ y espacios)
        for col in ['Descuentos','Retenciones','Valor Factura','Valor Pago']:
            if col in df_final.columns:
                df_final[col] = df_final[col].astype(str).str.replace('$','', regex=False).str.replace(' ','', regex=False).str.replace('−','-', regex=False)
                # remover repeticiones de '-' al inicio
                df_final[col] = df_final[col].str.replace(r'^-+','-', regex=True)

        # Intentar convertir Valor Pago a float (mantenemos la columna con ese nombre por ahora)
        if 'Valor Pago' in df_final.columns:
            def tryfloat(x):
                try:
                    if x in ('','nan','None'):
                        return 0.0
                    return float(re.sub(r'[^\d\-\.]', '', str(x)))
                except:
                    cleaned = re.sub(r'[^\d\-\.]', '', str(x))
                    try:
                        return float(cleaned) if cleaned not in ('','-') else 0.0
                    except:
                        return 0.0
            df_final['Valor Pago'] = df_final['Valor Pago'].apply(tryfloat)

        # limpiar signos en Descuentos/Retenciones como texto (si quieres convertirlos luego, adaptamos)
        if 'Descuentos' in df_final.columns:
            df_final['Descuentos'] = df_final['Descuentos'].astype(str).str.replace('$','', regex=False).str.strip()
        if 'Retenciones' in df_final.columns:
            df_final['Retenciones'] = df_final['Retenciones'].astype(str).str.replace('$','', regex=False).str.strip()

        return df_final

    # ---------------------------
    # Ejecutar procesamiento y devolver df_final
    # ---------------------------
    df_final = process_pdf(rutas["pdf_remittance"])

    # -------------------------------------------------
    # Guardar Remittance en buffer en memoria (Excel)
    # -------------------------------------------------
    remittance_buffer = io.BytesIO()
    with pd.ExcelWriter(remittance_buffer, engine="openpyxl") as writer:
        # Si df_final estuviera vacío, crear un sheet visible mínimo para evitar error openpyxl
        if df_final is None or df_final.empty:
            pd.DataFrame([{"Info":"No data extracted"}]).to_excel(writer, index=False)
        else:
            df_final.to_excel(writer, index=False)
    remittance_buffer.seek(0)

    # =====================================================
    # 4. Limpieza de Remittance (post-procesos)
    # =====================================================

    # Replace 'PMP ' and 'PM ' with 'PMP' en 'Doc.Cruce' si existe
    if 'Doc.Cruce' in df_final.columns:
        df_final['Doc.Cruce'] = df_final['Doc.Cruce'].str.replace('PMP ', 'PMP', regex=False)
        df_final['Doc.Cruce'] = df_final['Doc.Cruce'].str.replace('PM ', 'PMP', regex=False)

    # Asegurarse de que Valor Pago está como float (si no existe, crear 0)
    if 'Valor Pago' not in df_final.columns:
        df_final['Valor Pago'] = 0.0

    #Create the new column 'Tipo de Documento' basado en condiciones (mantener lógica anterior)
    if 'C.O.' in df_final.columns and 'Valor Pago' in df_final.columns:
        df_final['Tipo de Documento'] = df_final.apply(
            lambda row: 'Factura' if row['C.O.'] == 'CED' and float(row['Valor Pago']) > 0 else 'Descuento Cliente',
            axis=1
        )
    else:
        df_final['Tipo de Documento'] = 'Descuento Cliente'

    # Aplicar transformacion a Doc.Cruce si corresponde
    if 'Doc.Cruce' in df_final.columns:
        df_final['Doc.Cruce'] = df_final.apply(
            lambda row: str(row['Doc.Cruce'])[:10] if row['Tipo de Documento'] == 'Factura' else row['Doc.Cruce'],
            axis=1
        )

    # Renombrar Valor Pago -> Importe de Remittance (esto lo hacía tu pipeline original)
    if 'Valor Pago' in df_final.columns:
        df_final.rename(columns={"Valor Pago": "Importe de Remittance"}, inplace=True)

    # Renombrar Doc.Cruce -> Referencia / Factura
    if 'Doc.Cruce' in df_final.columns:
        df_final.rename(columns={"Doc.Cruce": "Referencia / Factura"}, inplace=True)

    # =====================================================
    # 5. Manejo de Reglas y CARDs (igual que antes)
    # =====================================================
    if "Tipo de Documento" in df_final.columns and "C.O." in df_final.columns:
        if "Descuento" not in df_final.columns:
            df_final["Descuento"] = ""
        if "Motivo del descuento" not in df_final.columns:
            df_final["Motivo del descuento"] = ""
        if "Comentarios" not in df_final.columns:
            df_final["Comentarios"] = ""

        mask_descuento = df_final["Tipo de Documento"] == "Descuento Cliente"
        co_values = ["MAY", "BEL", "MIX", "VEG", "FLO", "CAR", "SAL", "NUM", "LOB", "FRO",
                    "TER", "SAB", "ACA", "LAU", "ITA", "GUA", "LLA", "MUR", "MON", "ROS", "CED"]
        conds_desc = [df_final["C.O."].str.startswith(prefix, na=False) & mask_descuento for prefix in co_values]
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

    remittance = df_final.copy()

    # =====================================================
    # 6. Procesamiento de columnas 'Descuento' y 'Comentarios'
    # =====================================================
    remittance = procesar_descuentos_y_comentarios(remittance)

    # =====================================================
    # 7. Lectura de la Cartera (FBL5N) (datos desde SAP)
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
    # 13. Asignación de Pago Neto (Pago Neto = Importe de factura)
    # =====================================================
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
    hrc_template = hrc_template[columnas_finales]

    # =====================================================
    # 15. Exportar template final (usa tu función exportar_template)
    # =====================================================
    exportar_template(
        hrc_template=hrc_template,
        suma_remittance = remittance["Importe de Remittance"].sum() if "Importe de Remittance" in remittance.columns else 0.0,
        numero_orden="",
        id_cliente=id_cliente,
        nombre_cliente=nombre_cliente,
        ruta_remittance=remittance_buffer,
        ruta_salida=rutas["salida"]
    )

    return hrc_template
