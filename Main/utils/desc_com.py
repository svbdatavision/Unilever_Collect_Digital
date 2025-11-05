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
    # Asegurar tipo string y limpiar nulos
    remittance["Descuento"] = remittance["Descuento"].fillna("").astype(str)

    # Completar 'DESCUENTO' solo si está vacío y no es factura
    remittance.loc[
        (remittance["Descuento"].str.strip() == "") &
        (remittance["Tipo de Documento"] != "Factura"),
        "Descuento"
    ] = "DESCUENTO"

    # Asignar 'Comentarios'
    remittance["Comentarios"] = np.where(
        remittance["Tipo de Documento"] == "Factura",
        "",
        (remittance["Descuento"].fillna("") + " " + remittance["Referencia / Factura"].fillna("")).str.strip()
    )

    return remittance
