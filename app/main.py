"""API FastAPI: sirve la interfaz drag & drop y convierte el PDF a Excel."""

from __future__ import annotations

import os
import re
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, Response

from app.excel_writer import build_xlsx
from app.pdf_parser import parse_pdf

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"

XLSX_MEDIA_TYPE = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

app = FastAPI(title="Conversor PDF → Excel (Finiquito)")


def _safe_filename(name: str) -> str:
    stem = Path(name).stem or "finiquito"
    stem = re.sub(r"[^A-Za-z0-9_.\- ]+", "", stem).strip() or "finiquito"
    return f"{stem}.xlsx"


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/convert")
async def convert(file: UploadFile = File(...)):
    if not file.filename or not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="El archivo debe ser un PDF.")

    content = await file.read()
    if not content:
        raise HTTPException(status_code=400, detail="El archivo está vacío.")

    try:
        rows, _provider = parse_pdf(content)
    except Exception:
        raise HTTPException(
            status_code=400,
            detail="No se pudo leer el PDF. Verifica que sea un PDF válido de "
            "'CONFIRMACIÓN DE PAGOS'.",
        )

    if not rows:
        raise HTTPException(
            status_code=400,
            detail="No se encontraron filas de pagos en el PDF. "
            "¿Es el formato esperado?",
        )

    xlsx_bytes = build_xlsx(rows)
    out_name = _safe_filename(file.filename)

    return Response(
        content=xlsx_bytes,
        media_type=XLSX_MEDIA_TYPE,
        headers={
            "Content-Disposition": f'attachment; filename="{out_name}"',
            "X-Row-Count": str(len(rows)),
        },
    )


@app.get("/")
def index():
    return FileResponse(STATIC_DIR / "index.html")


if __name__ == "__main__":
    import uvicorn

    port = int(os.environ.get("PORT", "8000"))
    uvicorn.run(app, host="0.0.0.0", port=port)
