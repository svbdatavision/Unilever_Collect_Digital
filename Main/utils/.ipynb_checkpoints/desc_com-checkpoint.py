import numpy as np

def procesar_descuentos_y_comentarios(remittance):
    """
    =====================================================
    6. Procesamiento de columnas 'Descuento' y 'Comentarios'
    =====================================================
    - Si 'Descuento' está vacío y NO es una factura, completa con 'DESCUENTO'.
    - Si 'Tipo de Documento' = 'Factura', deja 'Comentarios' vacío.
    - En otros casos, genera 'Comentarios' combinando 'Descuento' y 'Referencia / Factura'.

    Parámetros:
        remittance (pd.DataFrame): DataFrame con la información de pagos y facturas.

    Retorna:
        pd.DataFrame: remittance actualizado con las columnas 'Descuento' y 'Comentarios' procesadas.
    """

    # =====================================================
    # 0. Asegurar existencia de columnas requeridas
    # =====================================================
    for col in ["Descuento", "Comentarios"]:
        if col not in remittance.columns:
            remittance[col] = ""

    # =====================================================
    # 1. Asegurar tipo string y limpiar nulos
    # =====================================================
    remittance["Descuento"] = remittance["Descuento"].fillna("").astype(str)

    # =====================================================
    # 2. Completar 'DESCUENTO' solo si está vacío y no es factura
    # =====================================================
    remittance.loc[
        (remittance["Descuento"].str.strip() == "") &
        (remittance["Tipo de Documento"] != "Factura"),
        "Descuento"
    ] = "DESCUENTO"

    # =====================================================
    # 3. Asignar 'Comentarios' solo si están vacíos
    # =====================================================
    mask_coment_vacios = remittance["Comentarios"].fillna("").str.strip() == ""

    remittance.loc[mask_coment_vacios, "Comentarios"] = np.where(
        remittance.loc[mask_coment_vacios, "Tipo de Documento"] == "Factura",
        "",
        (remittance.loc[mask_coment_vacios, "Descuento"].fillna("") + " " +
         remittance.loc[mask_coment_vacios, "Referencia / Factura"].fillna("")).str.strip()
    )

    return remittance
