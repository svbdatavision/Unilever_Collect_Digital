import pandas as pd
import numpy as np

def procesamiento_nro(hrc_template, FBL5N):
    """
    Procesa los registros de FBL5N con Reason code == 'NRO' (Notas de Crédito).
    - Ajusta la columna 'Referencia / Factura' usando 'Document Number' si se requiere.
    - Crea las notas de crédito como nuevas filas con 'Tipo de Documento' = 'Nota de Crédito'.
    - Completa las columnas relevantes: 'Pago Neto', 'Comentarios', 'Motivo del descuento', 'Descuento'.
    - Limpia y asigna los valores de descuento correspondientes ('Descuento' = 'NRO', 'Motivo del descuento' vacío).
    
    Parámetros:
        FBL5N (pd.DataFrame): Cartera proveniente de SAP.
        hrc_template (pd.DataFrame): DataFrame final (remittance + cartera).
    
    Retorna:
        pd.DataFrame: hrc_template actualizado con las notas de crédito.
    """
    # Filtrar solo las notas de crédito
    notas_credito = FBL5N[FBL5N["Reason code"] == "NRO"].copy()
    if notas_credito.empty:
        return hrc_template  # No hay notas de crédito, retornar tal cual

    # Marcar Tipo de Documento
    notas_credito["Tipo de Documento"] = "Nota de Crédito"

    # Completar columnas específicas directamente en notas_credito
#    notas_credito["Pago Neto"] = notas_credito["Importe de factura"]
    notas_credito["Referencia / Factura"] = notas_credito["Document Number"]
    notas_credito["Comentarios"] = notas_credito["Text"]
    notas_credito["Motivo del descuento"] = ""
    notas_credito["Descuento"] = "NRO"

    # Concatenar al template principal
    hrc_template = pd.concat([hrc_template, notas_credito], ignore_index=True)

    return hrc_template
