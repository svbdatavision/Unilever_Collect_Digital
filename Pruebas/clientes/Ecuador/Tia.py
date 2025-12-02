import os
import sys
import io
import warnings
import pandas as pd
import numpy as np
import camelot
import pdfplumber
import re

from clientes.utils import *

warnings.filterwarnings("ignore", category=UserWarning, module="camelot")

# =====================================================
# 1. Localización dinámica de la carpeta raíz del proyecto
# =====================================================
def _project_root():
    if getattr(sys, "frozen", False):
        macos_dir = os.path.dirname(sys.executable)
        contents_dir = os.path.dirname(macos_dir)
        app_bundle = os.path.dirname(contents_dir)
        return os.path.dirname(app_bundle)
    return os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))

# =====================================================
# 2. Función principal del proceso (procesar)
# =====================================================
def procesar():
    root = _project_root()
    
    # --- Rutas ---
    rutas = {
        "pdf_remittance": os.path.join(root, "Archivos", "Remittance", "Ecuador", "Remittance_TIA.pdf"),
        "fbl5n": os.path.join(root, "Archivos", "Cartera", "FBL5N_TIA.xlsx"),
        "salida": os.path.join(root, "Archivos", "Template", "Ecuador", "Template_HRC_TIA.xlsx")
    }
    
    customer_id = 10273302
    
    # =====================================================
    # 1. Lectura de Remitente
    # =====================================================
    def extraer_facturas(archivo_pdf):
        filas = []

        with pdfplumber.open(archivo_pdf) as pdf:
            for pagina in pdf.pages:
                words = pagina.extract_words(x_tolerance=2, y_tolerance=3)
                lineas_dict = {}

                for w in words:
                    y = round(w["top"])
                    x = w.get("x", w.get("x0", 0))
                    lineas_dict.setdefault(y, []).append((x, w["text"]))

                for y, contenido in sorted(lineas_dict.items()):
                    linea = " ".join(t for _, t in sorted(contenido))
                    if re.match(r"^\d{4}-\d{2}-\d{2}\s+\d+", linea):
                        filas.append(linea)

        registros = []
        for f in filas:
            numeros = re.findall(r"-?[\d,]+\.\d{2}", f)
            if len(numeros) < 3:
                continue

            bruto, ret, neto = [n.replace(",", "") for n in numeros[-3:]]
            parte_antes = f[:f.rfind(bruto)].strip()
            tokens = parte_antes.split()

            fecha = tokens[0]
            plaza_codigo = tokens[1]

            tipo_idx = next((i for i, tok in enumerate(tokens) if tok in ["F","C","D"]), None)
            if tipo_idx is None:
                continue

            plaza_pago = " ".join(tokens[2:tipo_idx])
            tipo_letra = tokens[tipo_idx]
            documento = tokens[tipo_idx + 1]
            descripcion = " ".join(tokens[tipo_idx + 2:])
            documento_final = f"{tipo_letra} {documento}"

            registros.append([
                fecha,
                f"{plaza_codigo} {plaza_pago}".strip(),
                tipo_letra,            # Tipo
                documento_final,       # Documento
                bruto,
                ret,
                neto
            ])

        df = pd.DataFrame(
            registros,
            columns=["Fecha Factura", "Plaza Pago", "Documento", "Tipo", "Bruto", "Retención", "Neto a Pagar"]
        )

        # Parche
        df[["Documento", "Tipo"]] = df[["Tipo", "Documento"]].copy()
        for c in ["Bruto", "Retención", "Neto a Pagar"]:
            df[c] = pd.to_numeric(df[c], errors="coerce")
        return df

    remittance = extraer_facturas(rutas["pdf_remittance"])
    
    # 2.1 Guardar Remittance en buffer en memoria
    remittance_buffer = io.BytesIO()
    with pd.ExcelWriter(remittance_buffer, engine="openpyxl") as writer:
        remittance.to_excel(writer, index=False)
    remittance_buffer.seek(0)
    # --- LIMPIEZA Y TRANSFORMACIÓN ---
    remittance = remittance.rename(columns={
        "Documento": "Referencia / Factura",
        "Tipo": "Relación Cliente",
        "Neto a Pagar": "Importe de Remittance"
    })

    # Conversión a numérico
    remittance["Importe de Remittance"] = pd.to_numeric(remittance["Importe de Remittance"], errors="coerce")
    remittance["Referencia / Factura"] = remittance["Referencia / Factura"].astype(str).str[1:]
    # Tipo de documento
    conds = [
        remittance["Relación Cliente"].str.startswith("F", na=False) & (remittance["Importe de Remittance"] > 0),
        remittance["Relación Cliente"].str.startswith("D", na=False) & (remittance["Importe de Remittance"] > 0)
    ]
    choices = ["Factura", "Descuentos Cliente"]
    remittance["Tipo de Documento"] = np.select(conds, choices, default="")

    # Ajuste de montos negativos para descuentos
    remittance.loc[remittance["Tipo de Documento"] == "Descuentos Cliente", "Importe de Remittance"] *= -1

    # Columnas para CARDs
    for col in ["Descuento", "Motivo del descuento", "Comentarios"]:
        if col not in remittance.columns:
            remittance[col] = ""

    remittance.loc[remittance["Tipo de Documento"] == "Descuentos Cliente", "Comentarios"] = (
        remittance["Relación Cliente"].astype(str) + " " + remittance["Referencia / Factura"].astype(str)
    )
    remittance["Motivo del descuento"] = np.where(
        remittance["Tipo de Documento"] == "Descuentos Cliente",
        "987",
        remittance["Motivo del descuento"]
    )

    
    # =====================================================
    # 6. Procesamiento de columnas 'Descuento' y 'Comentarios'
    # =====================================================
    
    remittance = procesar_descuentos_y_comentarios(remittance)

    # =====================================================
    # 7. Lectura de la Cartera (FBL5N) (datos desde SAP)
    # =====================================================
    # =====================================================
    # 8. Filtro de la cartera del cliente
    # =====================================================
    # =====================================================
    # 9. Renombrado y limpieza de columnas
    # =====================================================
    FBL5N, id_cliente, nombre_cliente = procesar_cartera_cliente(rutas["fbl5n"], customer_id)
    
    # =====================================================
    # 10. Merge Remittance + FBL5N por "Referencia / Factura"
    # =====================================================
    # Se realiza un merge tipo "left" sobre 'Referencia / Factura' para mantener todas
    # las filas de Remittance y añadir información de FBL5N cuando exista coincidencia
    hrc_template = merge_remittance_cartera(remittance, FBL5N)
    

    # =====================================================
    # 11. Cálculo de diferencias
    # =====================================================
    # Se calcula la diferencia entre 'Importe de factura' y 'Importe de Remittance'
    # La lógica centralizada se encuentra en la función procesar_diferencias()
    hrc_template = procesar_diferencias(hrc_template)
    
    # =====================================================
    # 12. Agregamos registros NRO
    # =====================================================
    hrc_template = procesamiento_nro(hrc_template, FBL5N)
    
    # =====================================================
    # 13. Asignación de Pago Neto (Pago Neto = Importe de factura) y otros ajustes
    # =====================================================
    # Por defecto, 'Pago Neto' = 'Importe de factura'
    # Pago Neto
    hrc_template["Pago Neto"] = hrc_template["Importe de factura"]

    # Columnas finales
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


    exportar_template(
        hrc_template=hrc_template,
        suma_remittance=remittance["Importe de Remittance"].sum(),
        numero_orden="numero_orden",
        id_cliente=id_cliente,
        nombre_cliente=nombre_cliente,
        ruta_remittance=remittance_buffer,
        ruta_salida=rutas["salida"]
    )

    return hrc_template
