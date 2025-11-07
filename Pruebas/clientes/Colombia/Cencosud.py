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

# Buscamos las funciones en la carpeta Main
# current_dir = os.path.dirname(os.path.abspath(__file__))
# project_root = os.path.abspath(os.path.join(current_dir, "..", "..", ".."))  # Sube desde /Pruebas/clientes/Colombia a /Raiz
# sys.path.append(project_root)
# from Main.utils import *  # Importación de funciones utilitarias del paquete interno 'clientes'

from clientes.utils import *  # Importación de funciones utilitarias del paquete interno 'clientes'

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
def procesar():
    """
    Orquestador principal para Cencosud
    """
    root = _project_root()

    # --- Rutas de entrada/salida ---
    rutas = {
        "pdf_remittance": os.path.join(root, "Archivos", "Remittance", "Colombia", "Remittance_Cenco.pdf"),  # Cencosud
#        "fbl5n": os.path.join(root, "Archivos", "Cartera", "FBL5N.xlsx"),
        "fbl5n": os.path.join(root, "Archivos", "Cartera", "FBL5N_Cenco.xlsx"),
        "salida": os.path.join(root, "Archivos", "Template", "Colombia", "Template_HRC_Cenco.xlsx")
    }
    # Colocar el Customer ID del cliente
    customer_id = 10267301

    # =====================================================
    # 1. Lectura de Remitente
    # =====================================================
    # Leemos la tabla principal del Remittance: (ajustado a PDF) - usando Camelot
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
