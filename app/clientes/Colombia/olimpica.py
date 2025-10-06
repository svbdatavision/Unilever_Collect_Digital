import sys  # Para detectar ejecución empaquetada (frozen) y resolver rutas
import os   # Para construir rutas relativas al proyecto
import pandas as pd  # Principal herramienta para manipulación tabular
import numpy as np   # Utilidades numéricas/condicionales (np.select, np.where)
from openpyxl import load_workbook  # Leer valores dinámicos desde archivos Excel
# Usamos imports relativos para integrarlos en el paquete.
from clientes.utils.formato_template import exportar_template  # Paso 11: exportar con formato
from clientes.utils.diferencias import procesar_diferencias    # Paso 9: lógica centralizada de diferencias


def _project_root():
    """
    Devuelve la carpeta raíz del proyecto:
    - Si corre dentro de un .app (PyInstaller / py2app) -> carpeta que contiene el .app
    - Si corre como script -> la carpeta del archivo actual (../)
    """
    if getattr(sys, "frozen", False):
        # Estructura típica de aplicaciones macOS empaquetadas
        macos_dir = os.path.dirname(sys.executable)      # .../MyApp.app/Contents/MacOS
        contents_dir = os.path.dirname(macos_dir)        # .../MyApp.app/Contents
        app_bundle = os.path.dirname(contents_dir)       # .../MyApp.app
        return os.path.dirname(app_bundle)               # carpeta que contiene el .app
    return os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))


def procesar():
    """
    Orquestador principal:
    Sigue el flujo numerado (1..11) documentado en el requerimiento.
    Devuelve el DataFrame final listo para exportar.
    """
    root = _project_root()

    # --- Rutas (ajustar si usas rutas absolutas o entornos distintos) ---
    rutas = {
        "remittance": os.path.join("Archivos", "Remittance", "Colombia", "Remittance_olimpica.xlsx"),
        "fbl5n": os.path.join("Archivos", "Cartera", "FBL5N_olimpica.xlsx"),
        "salida": os.path.join("Archivos", "Template", "Colombia", "Template_HRC_olimpica.xlsx")
    }

    # =====================================================
    # 1. Lectura de Remittance
    # =====================================================
    # Leemos la tabla principal del Remittance: (ajustado a tu Excel)
    remittance = (
        pd.read_excel(
            rutas["remittance"], skiprows=21, nrows=33,
            usecols=["Código de Documento", "No. Doc", "Total a Pagar"]
        )
        .dropna(subset=["Código de Documento"])
    )

    # =====================================================
    # 2. Limpieza de Remittance
    #    - Quitar filas con "Total"
    #    - Normalizar nombres de columnas
    #    - Normalizar referencias e importes
    # =====================================================
    # Garantizar tipo string para búsquedas textuales
    remittance["Código de Documento"] = remittance["Código de Documento"].astype(str)

    # Eliminar filas de totales que suelen venir en el footer del Excel
    remittance = remittance[~remittance["Código de Documento"].str.contains("Total", case=False, na=False)]

    # Renombrado estándar para trabajar con los mismos nombres en todo el flujo
    remittance = remittance.rename(columns={
        "Código de Documento": "Tipo de Documento",
        "No. Doc": "Referencia / Factura",
        "Total a Pagar": "Importe de factura"
    })

    # Normalizar referencia (ej: eliminar sufijos/últimos 3 caracteres si corresponde)
    remittance["Referencia / Factura"] = remittance["Referencia / Factura"].astype(str).str[0:-3]

    # Mapear códigos a nombres entendibles
    remittance["Tipo de Documento"] = remittance["Tipo de Documento"].replace({
        "380": "Factura",
        "381": "Descuentos no asociados a FC"
    })

    # Normalizar importes: quitar separador de miles y estandarizar decimales
    remittance["Importe de factura"] = (
        remittance["Importe de factura"]
        .astype(str)                     # aseguramos string antes de manipular
        .str.replace(".", "", regex=False)  # quitar puntos de miles
        .str.replace(",", ".", regex=False) # convertir coma decimal a punto
        .astype(float)                   # convertir a float
        .round(2)
    )

    # Asegurar la existencia de columnas que usaremos más adelante
    for col in ["Descuento", "Motivo del descuento"]:
        if col not in remittance.columns:
            remittance[col] = ""

    # =====================================================
    # 3. Diferenciamos Facturas de diferencias (Asignamos CARDs / Reglas)
    #    3.1 Referencia / Factura
    #    3.2 Descuentos (si vacío -> "Descuento")
    #    3.3 Motivo del descuento (ODIS)
    #    3.4 Comentarios (más adelante)
    # =====================================================
    # Reglas basadas en prefijos de la referencia y signo del importe
    conds = [
        (remittance["Referencia / Factura"].str.startswith("PMP", na=False)) & (remittance["Importe de factura"] < 0),
        remittance["Referencia / Factura"].str.startswith("0085489", na=False),
        remittance["Referencia / Factura"].str.startswith("463", na=False),
        remittance["Referencia / Factura"].str.startswith("9801649", na=False),
        remittance["Referencia / Factura"].str.startswith("310", na=False)
    ]
    descuentos = ["RECHAZO", "DESCUENTO", "AVERIA", "FACT PROVEEDOR", "NOTA DEBITO"]
    motivos = ["551", "987", "522", "CSB", "987"]

    # Aplicar reglas de descuento y motivo
    remittance["Descuento"] = np.select(conds, descuentos, default=remittance["Descuento"])
    remittance["Motivo del descuento"] = np.select(conds, motivos, default=remittance["Motivo del descuento"])

    # 3.2.1 Si Descuento quedó vacío -> lo completamos con la palabra "Descuento"
    # Esto facilita que en pasos posteriores los defaults no queden vacíos.
    remittance["Descuento"] = remittance["Descuento"].fillna("").astype(str)
    remittance.loc[remittance["Descuento"].str.strip() == "", "Descuento"] = "Descuento"

    # =====================================================
    # 4. Lectura de FBL5N (datos desde SAP / conciliación)
    # =====================================================
    FBL5N = pd.read_excel(rutas["fbl5n"], usecols=["Document Type", "Reference", "Amount in local currency"])

    # =====================================================
    # 5. Filtro FBL5N (traer solo RV / facturas que nos interesan)
    # =====================================================
    FBL5N = FBL5N[FBL5N["Document Type"] == "RV"]  # keep only invoices (RV)

    # =====================================================
    # 6. Rename de columnas FBL5N
    #    6.1 "Reference" -> "Referencia / Factura"
    #    6.2 "Amount in local currency" -> "importe_FBL5N"
    # =====================================================
    FBL5N = FBL5N.rename(columns={
        "Reference": "Referencia / Factura",
        "Amount in local currency": "importe_FBL5N"
    }).reset_index(drop=True)

    # =====================================================
    # 7. Trabajamos con los NRO (opcional / placeholder)
    #    7.1 Aquí podrías leer FBL3N y crear reglas específicas para NRO.
    #    (No implementado automáticamente: dejé el placeholder para que lo completes)
    # =====================================================
    # Ejemplo (comentado):
    # # fbl3n = pd.read_excel(rutas['fbl3n'], ...)
    # # aplicar lógica NRO -> marcar registros en remittance o FBL5N
    
    
    # FALTA
    
    # =====================================================
    # 8. Merge por "Referencia / Factura"
    #    8.1 how="left" -> conservar todo Remittance y añadir datos de FBL5N cuando coincidan
    #    8.2 Eliminamos referencias especiales (ej: CNQ*) para evitar falsos positivos
    # =====================================================
    hrc_template = pd.merge(remittance, FBL5N, on="Referencia / Factura", how="left")

    # 8.2 Eliminar referencias que empiezan con "CNQ" (solicitud previa).
    #      Esto evita que entradas como 'CNQ...' generen diferencias o filas incorrectas.
    hrc_template = hrc_template[~hrc_template["Referencia / Factura"].astype(str).str.startswith("CNQ", na=False)].reset_index(drop=True)

    # =====================================================
    # 9. Calculamos Diferencias entre importe_FBL5N - Importe de factura
    #    (logica centralizada en procesar_diferencias)
    # =====================================================
    # Llamamos a la función reutilizable que crea las filas de "MENORES VALORES"
    hrc_template = procesar_diferencias(hrc_template)

    # 10.1 Construcción/ajuste de Comentarios (si se requiere sobrescribir/ajustar)
    hrc_template["Comentarios"] = np.where(
        hrc_template["Tipo de Documento"] == "Factura", "",
        np.where(
            hrc_template["Descuento"] == "MENORES VALORES",
            hrc_template["Descuento"],
            hrc_template["Descuento"].fillna("") + " " + hrc_template["Referencia / Factura"].fillna("")
        )
    )
    # 10.2 Pago Neto = Importe de factura (por diseño; puede ajustarse si hace falta)
    hrc_template["Pago Neto"] = hrc_template["Importe de factura"]
    
    # =====================================================
    # 11. Definimos columnas para template (formato final)
    #    11.1 Tipo de Documento
    #    11.2 Referencia / Factura
    #    11.3 Importe de factura
    #    11.4 Descuento
    #    11.5 Motivo del descuento
    #    11.6 Pago Neto
    #    11.7 Comentarios
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
    # 12. Ajustamos a formato Template y exportamos (encabezados, colores, hoja Remittance)
    # =====================================================
    # 12.1 Extraer datos dinámicos del Remittance (orden de pago, id/nombre de cliente)
    wb_rem = load_workbook(rutas["remittance"], data_only=True)
    celda = wb_rem.active["C10"].value
    numero_orden = celda.split("Orden de Pago:")[1].strip() if celda and "Orden de Pago:" in celda else ""

    # 12.2 Extraer id_cliente / nombre_cliente desde FBL5N (primer registro)
    fbl5n_meta = pd.read_excel(rutas["fbl5n"], usecols=["Customer", "Name 1"], nrows=1)
    id_cliente = fbl5n_meta["Customer"].iloc[0] if not fbl5n_meta.empty else ""
    nombre_cliente = fbl5n_meta["Name 1"].iloc[0] if not fbl5n_meta.empty else ""

    # 12.3 Exportar (la función exportar_template aplica el formato y copia hoja Remittance)
    exportar_template(
        hrc_template=hrc_template,
        numero_orden=numero_orden,
        id_cliente=id_cliente,
        nombre_cliente=nombre_cliente,
        ruta_remittance=rutas["remittance"],
        ruta_salida=rutas["salida"]
    )

    # Devolver template final (útil para testing / inspección)
    return hrc_template
