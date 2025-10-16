from .desc_com import procesar_descuentos_y_comentarios
from .merge import merge_remittance_cartera
from .diferencias import procesar_diferencias
from .nro import procesamiento_nro
from .formato_template import exportar_template
from .cartera import procesar_cartera_cliente



__all__ = [
    "procesar_descuentos_y_comentarios",
    "merge_remittance_cartera",
    "procesar_diferencias", # Paso 9: lógica centralizada de diferencias
    "procesamiento_nro",
    "exportar_template", # Paso 11: exportar con formato
    "procesar_cartera_cliente"
]
