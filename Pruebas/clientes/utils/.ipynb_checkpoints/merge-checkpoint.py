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
    

    # Merge SOLO por referencia, sin filtrar facturas todavía
    merged = pd.merge(
        remittance[["Referencia / Factura"]],
        FBL5N,
        on="Referencia / Factura",
        how="left",
        suffixes=("", "_FBL5N")
    )
    
    # Filtramos solo las facturas
    facturas = remittance[
        (remittance["Tipo de Documento"] == "Factura")
        & (remittance["Referencia / Factura"].isin(FBL5N["Referencia / Factura"]))
    ]

    # Hacemos el merge solo con esas filas
    facturas_merged = pd.merge(
        facturas,
        FBL5N,
        on="Referencia / Factura",
        how="left",
        suffixes=("", "_FBL5N")  # evita colisiones de nombres
    )

    # Columnas nuevas (solo las del FBL5N)
    columnas_nuevas = [col for col in merged.columns if col not in remittance.columns]

    # Añadir columnas nuevas vacías
    for col in columnas_nuevas:
        hrc_template[col] = None

    # Asignar SOLO donde coincide la factura
    mask = hrc_template["Tipo de Documento"] == "Factura"
    hrc_template.loc[mask, columnas_nuevas] = merged.loc[mask, columnas_nuevas].values

    # Normalizar columna REFERENCIA / FACTURA
    # 1. Reemplaza valores no escalares por string seguro
    hrc_template["Referencia / Factura"] = (
        hrc_template["Referencia / Factura"]
            .apply(lambda x: "" if x is None else str(x))  # convierte todo a texto
            .str.replace(r"[\[\]\(\)\{\}]", "", regex=True)  # remueve secuencias tipo lista/array
            .str.replace(r"\s+", " ", regex=True)           # normaliza espacios
            .str.strip()
    )
    # Eliminar facturas mal cargadas
    # (luego de normalizar y asegurar scalar strings)
    hrc_template = (
        hrc_template[
            ~(
                (hrc_template["Tipo de Documento"] == "Factura") &
                (hrc_template["Referencia / Factura"].str.len() != 10)
            )
        ]
        .reset_index(drop=True)
    )
    
    return hrc_template
