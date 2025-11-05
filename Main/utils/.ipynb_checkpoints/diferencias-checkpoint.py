import pandas as pd
import numpy as np

    # =====================================================
    # 8. Calculamos Diferencias entre Importe de factura - Importe de Remittance
    #    (logica centralizada en procesar_diferencias)
    # =====================================================
    # Llamamos a la función reutilizable que crea las filas de "MENORES VALORES"

def procesar_diferencias(hrc_template: pd.DataFrame) -> pd.DataFrame:

    # 8.1 Crear columna de diferencia solo para Facturas
    hrc_template["Diferencia"] = pd.NA
    hrc_template.loc[hrc_template["Tipo de Documento"] == "Factura", "Diferencia"] = (
        hrc_template["Importe de factura"] - hrc_template["Importe de Remittance"]
    )

    # 8.2 Filtrar registros donde haya diferencias distintas de cero
    diferencias = hrc_template[
        hrc_template["Diferencia"].notna() & (hrc_template["Diferencia"] != 0)
    ].copy()

    if len(diferencias) > 20:
        # ---- 9.3 Caso de más de 20 diferencias: línea resumida ----
        suma_total = diferencias["Diferencia"].sum()

        registros_diferencias = pd.DataFrame({
            "Tipo de Documento": ["Descuentos Clientes"],
            "Referencia / Factura": ["MENORES VALORES"],
            "Importe de Remittance": [suma_total],
            "Pago Neto": [""],
            "Descuento": ["MENORES VALORES"],
            # 9.3.1 Motivo según signo de la suma total
            "Motivo del descuento": ["WOB" if suma_total < 0 else "384"],
            "Comentarios": ["MENORES VALORES"]
        })

    else:
        # ---- 8.4 Caso de 20 o menos diferencias: detalle línea por línea ----
        registros_diferencias = pd.DataFrame({
            "Tipo de Documento": "Descuentos Clientes",
            "Referencia / Factura": diferencias["Referencia / Factura"],
            "Importe de Remittance": diferencias["Diferencia"],
            "Pago Neto": "",
            "Descuento": "MENORES VALORES",
            # 8.4.1 Motivo de descuento (ODIS) según rangos
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

    # 8.4.2 Asignar comentario solo a los registros que tienen "MENORES VALORES"
    registros_diferencias["Comentarios"] = np.where(
        registros_diferencias["Descuento"] == "MENORES VALORES",
        "MENORES VALORES",
        registros_diferencias.get("Comentarios", "")  # mantiene valores previos o vacío si no existían
    )

    # 8.4.3 Ajustes adicionales para motivo 987
    cond_1 = (registros_diferencias["Motivo del descuento"] == "987") & \
             (registros_diferencias["Importe de Remittance"] < -20000)
    cond_2 = (registros_diferencias["Motivo del descuento"] == "987") & \
             (registros_diferencias["Importe de Remittance"] > 20000)

    registros_diferencias.loc[cond_1, "Comentarios"] = (
        "Myr Vlr Pagado " + registros_diferencias.loc[cond_1, "Referencia / Factura"].fillna("")
    )
    registros_diferencias.loc[cond_2, "Comentarios"] = (
        "Saldo FC " + registros_diferencias.loc[cond_2, "Referencia / Factura"].fillna("")
    )

    # 8.5 Concatenar diferencias al template original
    hrc_template = pd.concat([hrc_template, registros_diferencias], ignore_index=True)

    # Parche Valores de descuentos
    hrc_template["Importe de factura"] = hrc_template["Importe de factura"].where(
    hrc_template["Importe de factura"].notna() & (hrc_template["Importe de factura"] != ""),
    hrc_template["Importe de Remittance"]
    )

#    hrc_template["Importe de factura"] = hrc_template["Importe de factura"].where(
#    hrc_template["Descuento"] != "Rechazo",
#    -hrc_template["Importe de factura"].astype(float)
#    )


    return hrc_template
