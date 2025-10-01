import pandas as pd
import numpy as np


def procesar_diferencias(hrc_template: pd.DataFrame) -> pd.DataFrame:
    """
    Paso 9. Calculamos Diferencias entre valor Facturas FBL5N - valor Facturas Remittance
    """

    # 9.1 Crear columna de diferencia solo para Facturas
    hrc_template["Diferencia"] = pd.NA
    hrc_template.loc[hrc_template["Tipo de Documento"] == "Factura", "Diferencia"] = (
        hrc_template["importe_FBL5N"] - hrc_template["Importe de factura"]
    )

    # 9.2 Filtrar registros donde haya diferencias distintas de cero
    diferencias = hrc_template[
        hrc_template["Diferencia"].notna() & (hrc_template["Diferencia"] != 0)
    ].copy()

    if len(diferencias) > 20:
        # ---- 9.3 Caso de más de 20 diferencias: línea resumida ----
        suma_total = diferencias["Diferencia"].sum()

        registros_diferencias = pd.DataFrame({
            "Tipo de Documento": ["Descuentos no asociados a FC"],
            "Referencia / Factura": ["MENORES VALORES"],
            "Importe de factura": [suma_total],
            "Pago Neto": [""],
            "Descuento": ["MENORES VALORES"],
            # 9.3.1 Motivo según signo de la suma total
            "Motivo del descuento": ["WOB" if suma_total < 0 else "384"],
            "Comentarios": ["MENORES VALORES"]
        })

    else:
        # ---- 9.4 Caso de 20 o menos diferencias: detalle línea por línea ----
        registros_diferencias = pd.DataFrame({
            "Tipo de Documento": "Descuentos no asociados a FC",
            "Referencia / Factura": diferencias["Referencia / Factura"],
            "Importe de factura": diferencias["Diferencia"],
            "Pago Neto": "",
            "Descuento": "MENORES VALORES",
            # 9.4.1 Motivo de descuento (ODIS) según rangos
            "Motivo del descuento": np.select(
                condlist=[
                    # Mayor a ±20.000
                    (diferencias["Diferencia"] <= -20000) | (diferencias["Diferencia"] >= 20000),
                    # Entre -20.000 y 0
                    (diferencias["Diferencia"].between(-20000, 0, inclusive="neither")),
                    # Entre 0 y 20.000
                    (diferencias["Diferencia"].between(0, 20000, inclusive="left"))
                ],
                choicelist=["987", "WOB", "384"],
                default="Error (Revisar)"
            )
        })

        # 9.4.2 Construcción de comentarios
        comentarios = np.where(
            registros_diferencias["Tipo de Documento"] == "Factura", "",
            np.where(
                registros_diferencias["Descuento"] == "MENORES VALORES",
                "MENORES VALORES " + registros_diferencias["Referencia / Factura"].fillna(""),
                registros_diferencias["Descuento"].fillna("") + " "
                + registros_diferencias["Referencia / Factura"].fillna("")
            )
        )
        registros_diferencias["Comentarios"] = comentarios

        # 9.4.3 Ajustes adicionales para motivo 987
        cond_1 = (registros_diferencias["Motivo del descuento"] == "987") & (registros_diferencias["Importe de factura"] < -20000)
        cond_2 = (registros_diferencias["Motivo del descuento"] == "987") & (registros_diferencias["Importe de factura"] > 20000)

        registros_diferencias.loc[cond_1, "Comentarios"] = (
            "Myr Vlr Pagado " + registros_diferencias.loc[cond_1, "Referencia / Factura"].fillna("")
        )
        registros_diferencias.loc[cond_2, "Comentarios"] = (
            "Saldo FC " + registros_diferencias.loc[cond_2, "Referencia / Factura"].fillna("")
        )

    # 9.5 Concatenar diferencias al template original
    hrc_template = pd.concat([hrc_template, registros_diferencias], ignore_index=True)

    return hrc_template
