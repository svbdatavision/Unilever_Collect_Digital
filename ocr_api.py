from fastapi import FastAPI, UploadFile, File
from PIL import Image
import pytesseract
import re
import io

app = FastAPI(title="Remittance OCR")

FACT_REGEX = r"\b(PMP\d+|NCMI\d+)\b"
IMP_REGEX = r"\(?\d{1,3}(?:,\d{3})+\)?"

def limpiar_importe(x: str) -> int:
    x = x.replace(",", "")
    if x.startswith("(") and x.endswith(")"):
        return -int(x.strip("()"))
    return int(x)

@app.post("/ocr/remittance")
async def ocr_remittance(file: UploadFile = File(...)):
    image_bytes = await file.read()
    img = Image.open(io.BytesIO(image_bytes)).convert("L")

    text = pytesseract.image_to_string(
        img,
        lang="spa",
        config="--psm 6"
    )

    facturas = []
    importes = []

    for line in text.splitlines():
        mf = re.search(FACT_REGEX, line)
        mi = re.search(IMP_REGEX, line)

        if mf and mi:
            facturas.append(mf.group())
            importes.append(limpiar_importe(mi.group()))

    return {
        "total": len(facturas),
        "items": [
            {"factura": f, "importe": i}
            for f, i in zip(facturas, importes)
        ]
    }
