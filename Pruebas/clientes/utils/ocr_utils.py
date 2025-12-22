from PIL import Image
import pytesseract
import re
import io
import pandas as pd


def ocr_from_image_bytes(
    image_bytes: bytes,
    factura_regex: str,
    importe_regex: str,
    lang: str = "spa",
    psm: int = 6
) -> pd.DataFrame:
    """
    Ejecuta OCR sobre una imagen y extrae factura + importe según regex.
    """

    img = Image.open(io.BytesIO(image_bytes)).convert("L")

    text = pytesseract.image_to_string(
        img,
        lang=lang,
        config=f"--psm {psm}"
    )

    facturas = []
    importes = []

    for line in text.splitlines():
        mf = re.search(factura_regex, line)
        mi = re.search(importe_regex, line)

        if mf and mi:
            facturas.append(mf.group())
            importes.append(_limpiar_importe(mi.group()))

    return pd.DataFrame({
        "factura": facturas,
        "importe": importes
    })


def _limpiar_importe(x: str) -> int:
    x = x.replace(",", "")
    if x.startswith("(") and x.endswith(")"):
        return -int(x.strip("()"))
    return int(x)
