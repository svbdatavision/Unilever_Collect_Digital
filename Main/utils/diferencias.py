import pandas as pd
import numpy as np
import config

def procesar_diferencias(hrc_template: pd.DataFrame) -> pd.DataFrame:

    pais_actual = config.pais_actual
    if pais_actual is None:
        raise ValueError("La variable 'pais_actual' no está definida en config.")

    elif pais_actual == 'Colombia':
        limite_inferior = -20000
        limite_superior = 20000
    elif pais_actual == 'Ecuador':
        limite_inferior = -1
        limite_superior = 1
    else:
        raise ValueError(fr"País no soportado. Use 'Colombia' o 'Ecuador'. El pais actual es {pais_actual}")

    # Crear columna de diferencia solo para Facturas
    hrc_template["Diferencia"] = pd.NA
    hrc_template.loc[hrc_template["Tipo de Documento"] == "Factura", "Diferencia"] = (
        hrc_template["Importe de factura"] - hrc_template["Importe de Remittance"]
    )

    # Filtrar registros con diferencias distintas de cero
    diferencias = hrc_template[
        hrc_template["Diferencia"].notna() & (hrc_template["Diferencia"] != 0)
    ].copy()

    if len(diferencias) > 20:
        suma_total = diferencias["Diferencia"].sum()
        registros_diferencias = pd.DataFrame({
            "Tipo de Documento": ["Descuentos Clientes"],
            "Referencia / Factura": ["MENORES VALORES"],
            "Importe de Remittance": [suma_total],
            "Pago Neto": [""],
            "Descuento": ["MENORES VALORES"],
            "Motivo del descuento": ["WOB" if suma_total < 0 else "384"],
            "Comentarios": ["MENORES VALORES"]
        })
    else:
        registros_diferencias = pd.DataFrame({
            "Tipo de Documento": "Descuentos Clientes",
            "Referencia / Factura": diferencias["Referencia / Factura"],
            "Importe de Remittance": diferencias["Diferencia"],
            "Pago Neto": "",
            "Descuento": "MENORES VALORES",
            "Motivo del descuento": np.select(
                condlist=[
                    (diferencias["Diferencia"] <= limite_inferior) | (diferencias["Diferencia"] >= limite_superior),
                    (diferencias["Diferencia"].between(limite_inferior, 0, inclusive="neither")),
                    (diferencias["Diferencia"].between(0, limite_superior, inclusive="left"))
                ],
                choicelist=["987", "WOB", "384"],
                default="Error (Revisar)"
            )
        })

    registros_diferencias["Comentarios"] = np.where(
        registros_diferencias["Descuento"] == "MENORES VALORES",
        "MENORES VALORES",
        registros_diferencias.get("Comentarios", "")
    )

    cond_1 = (registros_diferencias["Motivo del descuento"] == "987") & \
             (registros_diferencias["Importe de Remittance"] < limite_inferior)
    cond_2 = (registros_diferencias["Motivo del descuento"] == "987") & \
             (registros_diferencias["Importe de Remittance"] > limite_superior)

    registros_diferencias.loc[cond_1, "Comentarios"] = (
        "Myr Vlr Pagado " + registros_diferencias.loc[cond_1, "Referencia / Factura"].fillna("")
    )
    registros_diferencias.loc[cond_2, "Comentarios"] = (
        "Saldo FC " + registros_diferencias.loc[cond_2, "Referencia / Factura"].fillna("")
    )

    hrc_template = pd.concat([hrc_template, registros_diferencias], ignore_index=True)

    hrc_template["Importe de factura"] = hrc_template["Importe de factura"].where(
        hrc_template["Importe de factura"].notna() & (hrc_template["Importe de factura"] != ""),
        hrc_template["Importe de Remittance"]
    )

    return hrc_template