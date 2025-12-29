# =====================================================
# 0. Importación de librerías y módulos utilitarios
# =====================================================
import sys  # Para detectar ejecución empaquetada (frozen) y resolver rutas
import os   # Para construir rutas relativas al proyecto
import pandas as pd  # Principal herramienta para manipulación tabular
import numpy as np   # Utilidades numéricas/condicionales (np.select, np.where)
from openpyxl import load_workbook  # Leer valores dinámicos desde archivos Excel

# Buscamos las funciones en la carpeta Main
# current_dir = os.path.dirname(os.path.abspath(__file__))
# project_root = os.path.abspath(os.path.join(current_dir, "..", "..", ".."))  # Sube desde /Pruebas/clientes/Colombia a /Raiz
# sys.path.append(project_root)
# from Main.utils import *  # Importación de funciones utilitarias del paquete interno 'clientes'
from clientes.utils import *  # Importación de funciones utilitarias del paquete interno 'clientes'

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
    # Caso 1: el programa está empaquetado (ej. PyInstaller o app bundle)
    if getattr(sys, "frozen", False):
        # sys.executable apunta al ejecutable dentro del .app (Mac) o .exe (Windows)
        macos_dir = os.path.dirname(sys.executable)
        contents_dir = os.path.dirname(macos_dir)
        app_bundle = os.path.dirname(contents_dir)
        # Devuelve la carpeta que contiene el .app (la raíz del proyecto)
        return os.path.dirname(app_bundle)
    # Caso 2: ejecución normal desde código fuente (.py)
    # Sube dos niveles desde el archivo actual para llegar a la raíz del proyecto
    return os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))

# =====================================================
# 2. Función principal del proceso (procesar)
# =====================================================

def procesar():
    """
    Orquestador principal para Farmatodo:
    Flujo numerado (1..12) siguiendo el estándar de procesos.
    Devuelve el DataFrame final listo para exportar.
    """
    # 1.1 Obtener ruta raíz del proyecto
    root = _project_root()

    # 1.2 Definición de rutas de entrada y salida
    rutas = {
        "remittance": os.path.join(root, "Archivos", "Remittance", "Colombia", "Remittance_farmatodo.xlsx"),
        "fbl5n": os.path.join(root, "Archivos", "Cartera", "FBL5N_farmatodo.xlsx"),
       # "fbl5n": os.path.join(root, "Archivos", "Cartera", "FBL5N_farmatodo.xlsx"),
        "salida": os.path.join(root, "Archivos", "Template", "Colombia", "Template_HRC_farmatodo.xlsx"),
    }
    # Colocar el Customer ID del cliente
    customer_id = 10324901

    # =====================================================
    # 3. Lectura de Remittance
    # =====================================================
    # Leemos la tabla principal del Remittance: (ajustado a Excel)
    remittance = pd.read_excel(rutas["remittance"], skiprows=6, nrows=2000, header=[0, 1])
    remittance.columns = [
        str(col[0]).strip() if "Unnamed" in str(col[1]) else str(col[1]).strip()
        for col in remittance.columns
    ]
    # =====================================================
    # 4. Limpieza de Remittance
    # =====================================================    
    remittance = remittance[["Nro Factura", "Descripción", "Total"]]
    remittance = remittance.dropna(subset=["Nro Factura"]).reset_index(drop=True)
    remittance = remittance.rename(columns={
        "Nro Factura": "Referencia / Factura",
        "Total": "Importe de Remittance"
    })
    
    # Esta función nos sirve para trabajar los valores numericos sin importar el formato con el cual nos los pasan.
    def normalize_amount(x):
        x = str(x)

        # Caso 1: tiene coma y punto → formato europeo (1.234.567,89)
        if "," in x and "." in x:
            x = x.replace(".", "").replace(",", ".")
        # Caso 2: solo coma → decimal con coma (1234,56)
        elif "," in x:
            x = x.replace(",", ".")
        # Caso 3: solo punto → decimal con punto (1234.56)
        # No se hace nada

        return float(x)

    remittance["Importe de Remittance"] = (
        remittance["Importe de Remittance"]
        .apply(normalize_amount)
        .round(2)
    )

    # Limpiar caracteres invisibles
    remittance["Referencia / Factura"] = remittance["Referencia / Factura"].str.replace(r"[\u202A-\u202E\u200E\u200F]", "", regex=True).str.strip()

    conds = [
        remittance["Referencia / Factura"].str.startswith("PMP", na=False) & (remittance["Importe de Remittance"] > 0),
        ~remittance["Referencia / Factura"].str.startswith("PMP", na=False) & (remittance["Importe de Remittance"] > 0),
        remittance["Importe de Remittance"] < 0
    ]
    choices = ["Factura", "Nota Debito", "Descuentos Clientes"]
    remittance["Tipo de Documento"] = np.select(conds, choices, default="")

    # =====================================================
    # 5. Manejo de Reglas y CARDs
    # =====================================================    
    #    - Referencia / Factura
    #    3.2 Descuentos (si vacío -> "Descuento")
    #    3.3 Motivo del descuento (ODIS)
    #    3.4 Comentarios (más adelante)
    # =====================================================
    # Reglas basadas en prefijos de la referencia y signo del importe
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
    # 6. Procesamiento de columnas 'Descuento' y 'Comentarios'
    # =====================================================
    
    remittance["Descuento"] = remittance["Descuento"].fillna("").astype(str)
    
    
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

    hrc_template["Referencia / Factura"] = hrc_template["Referencia / Factura"].str.replace(r"^NC-", "", regex=True)

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
    
    hrc_template["Referencia / Factura"] = hrc_template["Referencia / Factura"].str.replace(r"^CNQ", "", regex=True)


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
    # 15.1 Extraer datos dinámicos del Remittance (orden de pago, id/nombre de cliente)
    wb_rem = load_workbook(rutas["remittance"], data_only=True)
    ws_rem = wb_rem.active
    numero_orden = ws_rem["B7"].value

    # 15.3 Exportar (la función exportar_template aplica el formato y copia hoja Remittance)
    exportar_template(
        hrc_template=hrc_template,
        suma_remittance = remittance["Importe de Remittance"].sum(),
        numero_orden=numero_orden,
        id_cliente=id_cliente,
        nombre_cliente=nombre_cliente,
        ruta_remittance=rutas["remittance"],
        ruta_salida=rutas["salida"]
    )

    # Devolución del template final
    return hrc_template