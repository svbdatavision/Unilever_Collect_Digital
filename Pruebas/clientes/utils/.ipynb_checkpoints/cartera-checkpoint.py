import pandas as pd

def procesar_cartera_cliente(ruta_fbl5n, customer_id):
    """
    Lee el archivo FBL5N desde SAP, filtra solo las filas relevantes para un cliente
    y limpia/renombra columnas para integrarlas al flujo de conciliación.
    
    Parámetros:
    ----------
    ruta_fbl5n : str
        Ruta al archivo Excel de FBL5N.
    customer_id : int
        ID del cliente a filtrar en la columna 'Customer'.
    
    Retorna:
    -------
    pd.DataFrame
        DataFrame filtrado y con columnas renombradas y 'Importe de factura' como float.
    """

    # =====================================================
    # 7. Lectura de la Cartera (FBL5N) (datos desde SAP)
    # =====================================================
    # Se leen solo las columnas necesarias, todo como texto (dtype=str) para evitar errores de tipo
    # 'engine="openpyxl"' es el más estable y rápido para archivos .xlsx
    FBL5N = pd.read_excel(
        ruta_fbl5n,
        usecols=[
            "Customer",
            "Document Type",
            "Reference",
            "Amount in local currency",
            "Reason code",
            "Name 1",
            "Document Number", 
            "Text"
        ],
        dtype=str,
        engine="openpyxl"
    )

    # Convertimos Customer a numérico
    FBL5N["Customer"] = pd.to_numeric(FBL5N["Customer"], errors="coerce")

    # =====================================================
    # 8. Filtro de la cartera del cliente
    # =====================================================
    # Se conservan únicamente las filas donde:
    #   - "Document Type" == "RV" (facturas)
    #   - O "Reason code" == "NRO" (casos especiales)
    #   -  "Customer" == customer_id
    FBL5N = FBL5N[
        ((FBL5N["Document Type"] == "RV") | (FBL5N["Reason code"] == "NRO"))
        & (FBL5N["Customer"] == customer_id)
    ].reset_index(drop=True)

    # =====================================================
    # 9. Renombrado y limpieza de columnas
    # =====================================================
    # Se renombran las columnas clave para mayor claridad y consistencia con el resto del proceso.
    FBL5N = FBL5N.rename(columns={
        "Reference": "Referencia / Factura",
        "Amount in local currency": "Importe de factura"
    }).reset_index(drop=True)

    # Crear una máscara booleana que detecte valores entre paréntesis (formato contable negativo)
    mask_negativo = FBL5N["Importe de factura"].str.contains(r"\(", regex=True)

    # Limpiar y convertir a float
    FBL5N["Importe de factura"] = (
        FBL5N["Importe de factura"]
        .str.replace(",", "", regex=False)       # eliminar separadores de miles
        .str.replace(r"[\(\)]", "", regex=True)  # eliminar paréntesis
        .astype(float)                           # convertir a float
        * mask_negativo.map(lambda x: -1 if x else 1)  # aplicar signo negativo
    )

    return FBL5N
