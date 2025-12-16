import pandas as pd

def procesar_cartera_cliente(ruta_fbl5n, customer_id):
    """
    Lee el archivo FBL5N desde SAP, filtra solo las filas relevantes para un cliente
    y limpia/renombra columnas para integrarlas al flujo de conciliación.

    Retorna:
        FBL5N (DataFrame filtrado)
        id_cliente (str)
        nombre_cliente (str)
    """

    # =====================================================
    # 1. Lectura única del archivo FBL5N
    # =====================================================
    df = pd.read_excel(
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

    # Convertir Customer a numérico para filtrado
    df["Customer"] = pd.to_numeric(df["Customer"], errors="coerce")

    # =====================================================
    # 2. Extraer meta del cliente ANTES de filtrar Document Type
    # =====================================================
    meta_cliente = df[df["Customer"] == customer_id].copy()

    id_cliente = (
        int(meta_cliente["Customer"].iloc[0])
        if not meta_cliente.empty else None
    )

    nombre_cliente = (
        meta_cliente["Name 1"].iloc[0]
        if not meta_cliente.empty else ""
    )

    # =====================================================
    # 3. Filtrar solo facturas válidas (RV y NRO) del cliente
    # =====================================================
    FBL5N = df[
        ((df["Document Type"] == "RV") | (df["Reason code"] == "NRO"))
        & (df["Customer"] == customer_id)
    ].reset_index(drop=True)

    # =====================================================
    # 4. Renombrar columnas
    # =====================================================
    FBL5N = FBL5N.rename(columns={
        "Reference": "Referencia / Factura",
        "Amount in local currency": "Importe de factura"
    }).reset_index(drop=True)

    # =====================================================
    # 5. Limpieza del campo monetario
    # =====================================================
    FBL5N["Importe de factura"] = (
        FBL5N["Importe de factura"]
        .astype(str)
        .str.replace(",", "", regex=False)     # quitar separador de miles
        .str.replace(r"\s+", "", regex=True)   # quitar espacios
    )

    FBL5N["Importe de factura"] = pd.to_numeric(
        FBL5N["Importe de factura"], errors="coerce"
    )

    # =====================================================
    # 6. Retornar resultados
    # =====================================================
    return FBL5N, id_cliente, nombre_cliente
