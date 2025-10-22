import pandas as pd

def merge_remittance_cartera(remittance, FBL5N):
    """
    Realiza un merge entre remittance y FBL5N solo para filas donde Tipo de Documento = "Factura".
    Conserva todas las filas de remittance y añade columnas nuevas de FBL5N solo en facturas.

    Parámetros:
        remittance (pd.DataFrame): DataFrame principal con todas las transacciones.
        FBL5N (pd.DataFrame): DataFrame con información adicional a unir por 'Referencia / Factura'.

    Retorna:
        pd.DataFrame: DataFrame resultante con todas las filas de remittance y columnas adicionales de FBL5N solo en facturas.
    """
    # Hacemos una copia para no modificar el original
    hrc_template = remittance.copy()
    
    # Filtramos solo las facturas
    facturas = remittance[remittance["Tipo de Documento"] == "Factura"]
    
    # Hacemos el merge solo con esas filas
    facturas_merged = pd.merge(
        facturas,
        FBL5N,
        on="Referencia / Factura",
        how="left",
        suffixes=("", "_FBL5N")  # evita colisiones de nombres
    )
    
    # Determinamos columnas nuevas provenientes de FBL5N
    columnas_nuevas = [col for col in facturas_merged.columns if col not in remittance.columns]
    
    # Insertamos esas columnas en hrc_template (vacías por defecto)
    for col in columnas_nuevas:
        hrc_template[col] = None
    
    # Reemplazamos los valores de las filas que son facturas
    hrc_template.loc[
        hrc_template["Tipo de Documento"] == "Factura",
        columnas_nuevas
    ] = facturas_merged[columnas_nuevas].values
    
    return hrc_template
