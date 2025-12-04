import pandas as pd
import os
from openpyxl import load_workbook
from utils import *
import config

customer_id = 10489225

def procesar(archivos_remittance, archivo_fbl5n):
    rutas = {
        "remittance": archivos_remittance,
        "fbl5n": archivo_fbl5n,
        "salida": os.path.join(os.path.dirname(archivos_remittance), "Mega Santa Maria.xlsx")
    }

    df = pd.read_excel(archivos_remittance, sheet_name='Control de pagos - Detalle', engine='openpyxl')

    # Seleccionar las columnas deseadas
    columnas_deseadas = ['Documento', 'Comprobante proveedor', 'Monto pago']
    remittance = df[columnas_deseadas]

    # Eliminar guiones y juntar los números en la columna 'Comprobante proveedor'
    remittance['Comprobante proveedor'] = remittance['Comprobante proveedor'].str.replace('-', '', regex=False)
    remittance['Comprobante proveedor'] = remittance['Comprobante proveedor'].str.replace(' ', '', regex=False)

    # Renombrar columnas
    remittance = remittance.rename(columns={
        "Documento": "Tipo de Documento",
        "Comprobante proveedor": "Referencia / Factura",
        "Monto pago": "Importe de Remittance"
    })

    remittance['Tipo de Documento'] = remittance['Tipo de Documento'].apply(
        lambda x: 'Factura'
        if isinstance(x, str) and x.strip().upper() == 'FACTURA PROVEEDOR'
        else ('Factura' if str(x).strip().upper() == 'FACTURA PROVEEDOR' else 'Nota de credito')
    )

    FBL5N, id_cliente, nombre_cliente = procesar_cartera_cliente(rutas["fbl5n"], customer_id)
    hrc_template = merge_remittance_cartera(remittance, FBL5N)
    hrc_template = procesar_diferencias(hrc_template)
    hrc_template = procesamiento_nro(hrc_template, FBL5N)

    if "Motivo del descuento" in hrc_template.columns:
        mask = hrc_template["Motivo del descuento"].astype(str).str.strip().eq("384")
        if mask.any():
            hrc_template.loc[mask, "Motivo del descuento"] = "WOB"

    if "Motivo del descuento" in hrc_template.columns:
        mask = hrc_template["Motivo del descuento"].astype(str).str.strip().eq("987")
        if mask.any():
            hrc_template.loc[mask, "Motivo del descuento"] = "388"

    # Ajustar columnas finales
    hrc_template["Pago Neto"] = hrc_template["Importe de factura"]
    columnas_finales = [
        "Tipo de Documento",
        "Referencia / Factura",
        "Importe de factura",
        "Descuento",
        "Motivo del descuento",
        "Pago Neto",
        "Comentarios"
    ]
    hrc_template = hrc_template[columnas_finales]

    # Datos adicionales
    numero_orden = ""

    # Guardar remittance temporal en Excel
    ruta_remittance_excel = os.path.join(os.path.dirname(archivos_remittance), "remittance_temp.xlsx")
    remittance.to_excel(ruta_remittance_excel, index=False)

    # Exportar template final
    exportar_template(
        hrc_template=hrc_template,
        suma_remittance=remittance["Importe de Remittance"].sum(),
        numero_orden=numero_orden,
        id_cliente=id_cliente,
        nombre_cliente=nombre_cliente,
        ruta_remittance=ruta_remittance_excel,
        ruta_salida=rutas["salida"]
    )

    # (Opcional) Eliminar el archivo temporal
    if os.path.exists(ruta_remittance_excel):
        os.remove(ruta_remittance_excel)

    return hrc_template