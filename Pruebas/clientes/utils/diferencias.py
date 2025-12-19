import pandas as pd
import numpy as np
import config


def procesar_diferencias(hrc_template: pd.DataFrame) -> pd.DataFrame:
    """
    Procesa diferencias entre Importe de factura e Importe de Remittance
    """

    # =====================================================
    # 11.1. Umbrales por país
    # =====================================================
    pais_actual = config.pais_actual
    if pais_actual is None:
        raise ValueError("La variable 'pais_actual' no está definida en config.")

    umbrales = {
        'Colombia': (-20000, 20000),
        'Ecuador': (-1, 1),
        'Peru': (-4, 4)
    }

    if pais_actual not in umbrales:
        raise ValueError(
            fr"País no soportado. Use 'Colombia', 'Ecuador' o 'Peru'. El país actual es {pais_actual}"
        )

    limite_inferior, limite_superior = umbrales[pais_actual]

    # =====================================================
    # 11.2. Calcular Diferencia (solo Facturas)
    # =====================================================
    hrc_template["Diferencia"] = pd.NA

    if pais_actual == "Colombia":
        hrc_template.loc[hrc_template["Tipo de Documento"] == "Factura", "Diferencia"] = \
            (hrc_template["Importe de factura"] - hrc_template["Importe de Remittance"]) * -1
    else:
        hrc_template.loc[hrc_template["Tipo de Documento"] == "Factura", "Diferencia"] = \
            hrc_template["Importe de Remittance"] - hrc_template["Importe de factura"]

    # =====================================================
    # 11.3. Filtrar diferencias válidas
    # =====================================================
    diferencias = hrc_template[
        hrc_template["Diferencia"].notna() & (hrc_template["Diferencia"] != 0)
    ].copy()

    # =====================================================
    # 11.3.1. Caso con muchos registros y valores bajos
    # =====================================================
    if len(diferencias) <= 20:

        registros_diferencias = pd.DataFrame({
            "Tipo de Documento": "Descuentos Clientes",
            "Referencia / Factura": diferencias["Referencia / Factura"],
            "Importe de Remittance": diferencias["Diferencia"],
            "Pago Neto": "",
            "Descuento": "MENORES VALORES",
            "Motivo del descuento": np.select(
                condlist=[
                    (diferencias["Diferencia"] <= limite_inferior) |
                    (diferencias["Diferencia"] >= limite_superior),
                    diferencias["Diferencia"].between(limite_inferior, 0, inclusive="neither"),
                    diferencias["Diferencia"].between(0, limite_superior, inclusive="left")
                ],
                choicelist=["987", "WOB", "384"],
                default="Error (Revisar)"
            ),
            "Comentarios": "MENORES VALORES"
        })

        # Ajustes especiales de comentarios
        cond_1 = (registros_diferencias["Motivo del descuento"] == "987") & \
                 (registros_diferencias["Importe de Remittance"] < limite_inferior)
        cond_2 = (registros_diferencias["Motivo del descuento"] == "987") & \
                 (registros_diferencias["Importe de Remittance"] > limite_superior)

        registros_diferencias.loc[cond_1, "Comentarios"] = \
            "Myr Vlr Pagado " + registros_diferencias["Referencia / Factura"].astype(str)

        registros_diferencias.loc[cond_2, "Comentarios"] = \
            "Saldo FC " + registros_diferencias["Referencia / Factura"].astype(str)

        hrc_template = pd.concat([hrc_template, registros_diferencias], ignore_index=True)

        # Ajustar Importes de factura faltantes
        hrc_template["Importe de factura"] = hrc_template["Importe de factura"].where(
            hrc_template["Importe de factura"].notna() & (hrc_template["Importe de factura"] != ""),
            hrc_template["Importe de Remittance"]
        )

        return hrc_template

    # =====================================================
    # 11.4. Si hay MÁS de 20 → usar la lógica nueva (grandes + menores)
    # =====================================================

    grandes_dif = diferencias[
        (diferencias["Diferencia"] > limite_superior) |
        (diferencias["Diferencia"] < limite_inferior)
    ].copy()

    menores_dif = diferencias.drop(grandes_dif.index).copy()

    # =====================================================
    # 11.5. Crear líneas para grandes diferencias
    # =====================================================
    grandes_lineas = grandes_dif.assign(
        **{
            "Tipo de Documento": "Descuentos Clientes",
            "Importe de Remittance": grandes_dif["Diferencia"],
            "Pago Neto": "",
            "Descuento": grandes_dif["Diferencia"].apply(lambda x: "Menor Vlr Pagado" if x < 0 else "Myr Vlr Pagado"),
            "Motivo del descuento": grandes_dif["Diferencia"].apply(lambda x: "CSR" if x < 0 else "388"),
            "Comentarios": grandes_dif.apply(
                lambda x: f"{'Menor Vlr Pagado' if x['Diferencia'] < 0 else 'Myr Vlr Pagado'} {x['Referencia / Factura']}",
                axis=1
            )
        }
    )[
        [
            "Tipo de Documento", "Referencia / Factura", "Importe de Remittance",
            "Pago Neto", "Descuento", "Motivo del descuento", "Comentarios"
        ]
    ]

    # =====================================================
    # 11.6. Crear líneas para menores diferencias
    # =====================================================
    registros_menores = pd.DataFrame()

    if not menores_dif.empty:

        # Caso: agrupar si hay muchísimas
        if len(menores_dif) > 20:
            suma_total = menores_dif["Diferencia"].sum()
            registros_menores = pd.DataFrame([{
                "Tipo de Documento": "Descuentos Clientes",
                "Referencia / Factura": "MENORES VALORES",
                "Importe de Remittance": suma_total,
                "Pago Neto": "",
                "Descuento": "MENORES VALORES",
                "Motivo del descuento": "WOB" if suma_total < 0 else "384",
                "Comentarios": "MENORES VALORES"
            }])

        else:
            # Líneas individuales
            condiciones = [
                (menores_dif["Diferencia"] <= limite_inferior) | 
                (menores_dif["Diferencia"] >= limite_superior),
                menores_dif["Diferencia"].between(limite_inferior, 0, inclusive="neither"),
                menores_dif["Diferencia"].between(0, limite_superior, inclusive="left"),
            ]
            elecciones = ["987", "WOB", "384"]

            registros_menores = menores_dif.assign(
                **{
                    "Tipo de Documento": "Descuentos Clientes",
                    "Importe de Remittance": menores_dif["Diferencia"],
                    "Pago Neto": "",
                    "Descuento": "MENORES VALORES",
                    "Motivo del descuento": np.select(condiciones, elecciones, default="Error (Revisar)"),
                    "Comentarios": ""
                }
            )

            # Ajustes de comentarios
            cond_1 = (registros_menores["Motivo del descuento"] == "987") & \
                     (registros_menores["Importe de Remittance"] < limite_inferior)
            cond_2 = (registros_menores["Motivo del descuento"] == "987") & \
                     (registros_menores["Importe de Remittance"] > limite_superior)

            registros_menores.loc[cond_1, "Comentarios"] = \
                "Myr Vlr Pagado " + registros_menores["Referencia / Factura"].astype(str)
            registros_menores.loc[cond_2, "Comentarios"] = \
                "Saldo FC " + registros_menores["Referencia / Factura"].astype(str)

            registros_menores["Comentarios"] = np.where(
                registros_menores["Comentarios"] == "",
                "MENORES VALORES",
                registros_menores["Comentarios"]
            )

            registros_menores = registros_menores[
                [
                    "Tipo de Documento", "Referencia / Factura", "Importe de Remittance",
                    "Pago Neto", "Descuento", "Motivo del descuento", "Comentarios"
                ]
            ]

    # =====================================================
    # 11.7. Combinar todo
    # =====================================================
    registros_finales = pd.concat([grandes_lineas, registros_menores], ignore_index=True)

    hrc_template = pd.concat([hrc_template, registros_finales], ignore_index=True)

    # =====================================================
    # 11.8. Completar importes vacíos
    # =====================================================
    hrc_template["Importe de factura"] = hrc_template["Importe de factura"].where(
        hrc_template["Importe de factura"].notna() & (hrc_template["Importe de factura"] != ""),
        hrc_template["Importe de Remittance"]
    )

    return hrc_template
