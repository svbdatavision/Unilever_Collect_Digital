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
        DataFrame filtrado y con columnas renombradas y 'Importe de factura' como float (con decimales).
    """

    # =====================================================
    # 7. Lectura de la Cartera (FBL5N)
    # =====================================================
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

    # =====================================================
    # 8. Conversión y filtrado de registros
    # =====================================================
    # Convertir Customer a numérico
    FBL5N["Customer"] = pd.to_numeric(FBL5N["Customer"], errors="coerce")

    # Filtrar solo facturas (RV) o casos especiales (NRO) del cliente
    FBL5N = FBL5N[
        ((FBL5N["Document Type"] == "RV") | (FBL5N["Reason code"] == "NRO"))
        & (FBL5N["Customer"] == customer_id)
    ].reset_index(drop=True)

    # =====================================================
    # 9. Renombrar columnas clave
    # =====================================================
    FBL5N = FBL5N.rename(columns={
        "Reference": "Referencia / Factura",
        "Amount in local currency": "Importe de factura"
    }).reset_index(drop=True)

    # =====================================================
    # 10. Limpieza segura del campo monetario
    # =====================================================
    # El formato esperado es: -182098.71 o 18098.70
    # Eliminamos comas de miles y espacios, sin tocar los puntos decimales.
    FBL5N["Importe de factura"] = (
        FBL5N["Importe de factura"]
        .str.replace(",", "", regex=False)       # quitar comas (miles)
        .str.replace(r"\s+", "", regex=True)     # quitar espacios
    )

    # Convertir a float directamente (manteniendo signo y decimales)
    FBL5N["Importe de factura"] = pd.to_numeric(FBL5N["Importe de factura"], errors="coerce")

    # =====================================================
    # 11. Retornar el resultado final
    # =====================================================
    return FBL5N
