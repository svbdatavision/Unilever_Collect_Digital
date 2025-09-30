#Euro
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

    root = _project_root()

    # --- Rutas de entrada/salida ---
    rutas = {
        "pdf_remittance": os.path.join("Archivos", "Remittance", "Remittance_euro.pdf"),
        "fbl5n":          os.path.join("Archivos", "Base_de_datos", "FBL5N_euro.xlsx"),
        "salida":         os.path.join("Archivos", "Template", "Template_HRC_euro.xlsx")
    }

    pdf_to_read = rutas["pdf_remittance"]
 
    # ------------------------------------------------------------------
    # 1) Extraer todas las tablas del PDF de forma robusta
    # ------------------------------------------------------------------
    reader = PdfReader(pdf_to_read)
    num_pages = len(reader.pages)
 
    all_pages = set(range(1, num_pages+1))

    tables_stream = camelot.read_pdf(pdf_to_read, pages='all', flavor='stream', strip_text='\n')

    stream_pages = set(int(t.page) for t in tables_stream)
    missing_pages = all_pages - stream_pages

    tables_lattice = []
    if missing_pages:
        missing_pages_str = ",".join(str(p) for p in missing_pages)
        tables_lattice = camelot.read_pdf(pdf_to_read, pages=missing_pages_str, flavor='lattice', strip_text='\n')

    df_all = pd.concat([t.df for t in list(tables_stream) + list(tables_lattice)], ignore_index=True)

    header_row_idx = df_all[df_all.apply(lambda r: r.astype(str).str.contains('C. O.').any(), axis=1)].index[0]
    df_all.columns = df_all.iloc[header_row_idx]
    df_all = df_all.drop(index=list(range(header_row_idx + 1))).reset_index(drop=True)
    df_all = df_all[~df_all.apply(lambda r: all(r.astype(str) == df_all.columns.astype(str)), axis=1)].reset_index(drop=True)

    # ------------------------------------------------------------------
    # 2) Filtrar y ordenar
    # ------------------------------------------------------------------
    filter_values = ['CED','MAY','BEL','MIX','VEG','FLO','CAR','SAL','NUM','LOB','FRO','TER','SAB','ACA','LAU','ITA','GUA','LLA','MUR','MON','ROS']
    filtered_df = df_all[df_all[df_all.columns[0]].isin(filter_values)].copy()

    def sort_key(val):
        if val == "C. O.": return "0"
        if val == "CED":     return "1"
        return "2" + str(val)

    filtered_df["sort_order"] = filtered_df[filtered_df.columns[0]].apply(sort_key)
    filtered_df = (filtered_df
                   .sort_values(by="sort_order")
                   .drop(columns="sort_order")
                   .drop_duplicates()
                   .reset_index(drop=True))

  
    # --- Guardar Remittance en un buffer en memoria ---

    remittance_buffer = io.BytesIO()
    with pd.ExcelWriter(remittance_buffer, engine="openpyxl") as writer:
        filtered_df.to_excel(writer, index=False)
    remittance_buffer.seek(0)


    #Guardar como CSV
    filtered_df.to_csv("prueba_euro.csv", index=False)
    print(f"✅ Archivo CSV guardado en: {output_csv_path}")


