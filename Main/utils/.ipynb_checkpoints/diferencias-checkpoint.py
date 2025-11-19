# =====================================================
# 11. Cálculo de diferencias
# =====================================================

import pandas as pd
import numpy as np
import config

def procesar_diferencias(hrc_template: pd.DataFrame) -> pd.DataFrame:
    """
    Procesa las diferencias entre Importe de factura e Importe de Remittance.
    - Para diferencias mayores al umbral (según país), crea líneas individuales
      manteniendo la misma 'Referencia / Factura'.
    - Para diferencias menores, agrupa todo en una línea 'MENORES VALORES'.
    """

    # =====================================================
    # 11.1. Definir límites por país
    # =====================================================
    pais_actual = config.pais_actual
    if pais_actual is None:
        raise ValueError("La variable 'pais_actual' no está definida en config.")
    elif pais_actual == 'Colombia':
        limite_inferior, limite_superior = -20000, 20000
    elif pais_actual == 'Ecuador':
        limite_inferior, limite_superior = -1, 1
    else:
        raise ValueError(fr"País no soportado. Use 'Colombia' o 'Ecuador'. El país actual es {pais_actual}")

    # =====================================================
    # 11.2. Calcular diferencias solo para Facturas
    # =====================================================
    hrc_template["Diferencia"] = pd.NA
    hrc_template.loc[hrc_template["Tipo de Documento"] == "Factura", "Diferencia"] = (
        (hrc_template["Importe de factura"] - hrc_template["Importe de Remittance"]) * -1
    )

    # =====================================================
    # 11.3. Filtrar registros con diferencias válidas
    # =====================================================
    diferencias = hrc_template[
        hrc_template["Diferencia"].notna() & (hrc_template["Diferencia"] != 0)
    ].copy()

    if diferencias.empty:
        return hrc_template  # nada que procesar

    # =====================================================
    # 11.4. Separar grandes diferencias (>|límite|)
    # =====================================================
    grandes_dif = diferencias[
        (diferencias["Diferencia"] > limite_superior) | (diferencias["Diferencia"] < limite_inferior)
    ].copy()

    menores_dif = diferencias.drop(grandes_dif.index).copy()

    # =====================================================
    # 11.5. Crear líneas individuales para grandes diferencias
    # =====================================================
    grandes_lineas = grandes_dif.assign(
        **{
            "Tipo de Documento": "Descuento Cliente",
            "Importe de Remittance": grandes_dif["Diferencia"],
            "Pago Neto": "",
            "Descuento": grandes_dif["Diferencia"].apply(
                lambda x: "Menor Vlr Pagado" if x < limite_inferior else "Myr Vlr Pagado"
            ),
            "Motivo del descuento": grandes_dif["Diferencia"].apply(
                lambda x: "CSR" if x < limite_inferior else "388"
            ),
            "Comentarios": grandes_dif.apply(
                lambda x: f"{'Menor Vlr Pagado' if x['Diferencia'] < limite_inferior else 'Myr Vlr Pagado'} {x['Referencia / Factura']}", axis=1
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
    # 11.6. Agrupar menores diferencias si hay muchas
    # =====================================================
    registros_menores = pd.DataFrame()
    if not menores_dif.empty:
        if len(menores_dif) > 20:
            # Agrupar en una sola línea “MENORES VALORES”
            suma_total = menores_dif["Diferencia"].sum()
            registros_menores = pd.DataFrame([{
                "Tipo de Documento": "Descuento Cliente",
                "Referencia / Factura": "MENORES VALORES",
                "Importe de Remittance": suma_total,
                "Pago Neto": "",
                "Descuento": "MENORES VALORES",
                "Motivo del descuento": "WOB" if suma_total < 0 else "384",
                "Comentarios": "MENORES VALORES"
            }])
        else:
            # Crear líneas individuales para menores diferencias
            registros_menores = menores_dif.assign(
                **{
                    "Tipo de Documento": "Descuento Cliente",
                    "Importe de Remittance": menores_dif["Diferencia"],
                    "Pago Neto": "",
                    "Descuento": "MENORES VALORES",
                    "Motivo del descuento": np.where(
                        menores_dif["Diferencia"] < 0, "WOB", "384"
                    ),
                    "Comentarios": "MENORES VALORES"
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
    # 11.7. Combinar grandes y menores diferencias
    # =====================================================
    registros_finales = pd.concat(
        [grandes_lineas, registros_menores], ignore_index=True
    )

    # =====================================================
    # 11.8. Unir con el template original
    # =====================================================
    hrc_template = pd.concat([hrc_template, registros_finales], ignore_index=True)

    # =====================================================
    # 11.9. Completar valores vacíos de 'Importe de factura'
    # =====================================================
    hrc_template["Importe de factura"] = hrc_template["Importe de factura"].where(
        hrc_template["Importe de factura"].notna() & (hrc_template["Importe de factura"] != ""),
        hrc_template["Importe de Remittance"]
    )

    return hrc_template
