# =====================================================
# 11. Cálculo de diferencias
# =====================================================

import pandas as pd
import numpy as np

def procesar_diferencias(hrc_template: pd.DataFrame) -> pd.DataFrame:
    """
    Procesa las diferencias entre Importe de factura e Importe de Remittance.
    - Crea líneas de 'Descuento Cliente' cuando las diferencias superan ±20000.
    - Agrupa el resto de diferencias en una o varias líneas 'MENORES VALORES',
      garantizando que la sumatoria no supere ±20000.
    """

    # =====================================================
    # 11.1 Crear columna de diferencia solo para Facturas
    # =====================================================
    hrc_template["Diferencia"] = pd.NA
    hrc_template.loc[hrc_template["Tipo de Documento"] == "Factura", "Diferencia"] = (
        (hrc_template["Importe de factura"] - hrc_template["Importe de Remittance"]) * -1
    )

    # =====================================================
    # 11.2 Filtrar diferencias no nulas y no cero
    # =====================================================
    diferencias = hrc_template[
        hrc_template["Diferencia"].notna() & (hrc_template["Diferencia"] != 0)
    ].copy()

    # =====================================================
    # 11.3 Separar grandes diferencias (>|20000|)
    # =====================================================
    grandes_dif = diferencias[
        (diferencias["Diferencia"] > 20000) | (diferencias["Diferencia"] < -20000)
    ].copy()

    menores_dif = diferencias[
        ~((diferencias["Diferencia"] > 20000) | (diferencias["Diferencia"] < -20000))
    ].copy()

    # =====================================================
    # 11.4 Crear líneas individuales para grandes diferencias
    # =====================================================
    grandes_lineas = grandes_dif.assign(
        **{
            "Tipo de Documento": "Descuento Cliente",
            "Importe de Remittance": grandes_dif["Diferencia"],
            "Pago Neto": "",
            "Descuento": grandes_dif["Diferencia"].apply(
                lambda x: "Menor Vlr Pagado" if x < -20000 else "Myr Vlr Pagado"
            ),
            "Motivo del descuento": grandes_dif["Diferencia"].apply(
                lambda x: "CSR" if x < -20000 else "388"
            ),
            "Comentarios": grandes_dif.apply(
                lambda x: f"{'Menor Vlr Pagado' if x['Diferencia'] < -20000 else 'Myr Vlr Pagado'} {x['Referencia / Factura']}", axis=1
            )
        }
    )[
        [
            "Tipo de Documento",
            "Referencia / Factura",
            "Importe de Remittance",
            "Pago Neto",
            "Descuento",
            "Motivo del descuento",
            "Comentarios",
        ]
    ]

    # =====================================================
    # 11.5 Agrupar menores diferencias en bloques de ±20000
    # =====================================================
    registros_menores = []

    if not menores_dif.empty:
        menores_dif = menores_dif.sort_values(by="Diferencia", ascending=False)
        acumulado = 0
        bloque = []

        for _, row in menores_dif.iterrows():
            nueva_suma = acumulado + row["Diferencia"]

            # Si agregar esta línea supera el rango permitido ±20000 → cerrar bloque
            if abs(nueva_suma) > 20000 and bloque:
                registros_menores.append({
                    "Tipo de Documento": "Descuento Cliente",
                    "Referencia / Factura": "MENORES VALORES",
                    "Importe de Remittance": acumulado,
                    "Pago Neto": "",
                    "Descuento": "MENORES VALORES",
                    "Motivo del descuento": "CSR" if acumulado < 0 else "388",
                    "Comentarios": "MENORES VALORES"
                })
                # Reiniciar bloque con la nueva línea
                bloque = [row["Diferencia"]]
                acumulado = row["Diferencia"]
            else:
                bloque.append(row["Diferencia"])
                acumulado = nueva_suma

        # Cerrar el último bloque si queda algo
        if bloque:
            registros_menores.append({
                "Tipo de Documento": "Descuento Cliente",
                "Referencia / Factura": "MENORES VALORES",
                "Importe de Remittance": acumulado,
                "Pago Neto": "",
                "Descuento": "MENORES VALORES",
                "Motivo del descuento": "CSR" if acumulado < 0 else "388",
                "Comentarios": "MENORES VALORES"
            })

    # =====================================================
    # 11.6 Combinar todas las líneas generadas
    # =====================================================
    registros_finales = pd.concat(
        [grandes_lineas, pd.DataFrame(registros_menores)], ignore_index=True
    )

    # =====================================================
    # 11.7 Concatenar al template original
    # =====================================================
    hrc_template = pd.concat([hrc_template, registros_finales], ignore_index=True)

    # =====================================================
    # 11.8 Parche valores vacíos de factura
    # =====================================================
    hrc_template["Importe de factura"] = hrc_template["Importe de factura"].where(
        hrc_template["Importe de factura"].notna()
        & (hrc_template["Importe de factura"] != "")
        & (hrc_template["Importe de factura"] != 0),
        hrc_template["Importe de Remittance"]
    )

    return hrc_template