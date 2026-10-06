"""API FastAPI: sirve la interfaz drag & drop y convierte PDF(s) a Excel."""

from __future__ import annotations

import io
import os
import re
import zipfile
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
ZIP_MEDIA_TYPE = "application/zip"

app = FastAPI(title="Conversor PDF → Excel (Finiquito)")


def _safe_filename(name: str) -> str:
    stem = Path(name).stem or "finiquito"
    stem = re.sub(r"[^A-Za-z0-9_.\- ]+", "", stem).strip() or "finiquito"
    return f"{stem}.xlsx"


@app.get("/health")
def health():
    return {"status": "ok"}


def _convert_one(filename: str, content: bytes) -> tuple[bytes | None, int, str | None]:
    """Convierte un PDF. Devuelve (xlsx_bytes, filas, error)."""
    if not content:
        return None, 0, "archivo vacío"
    try:
        rows, _provider = parse_pdf(content)
    except Exception:
        return None, 0, "no se pudo leer el PDF"
    if not rows:
        return None, 0, "no se encontraron filas de pagos"
    return build_xlsx(rows), len(rows), None


@app.post("/convert")
async def convert(files: list[UploadFile] = File(...)):
    pdfs = [f for f in files if f.filename]
    if not pdfs:
        raise HTTPException(status_code=400, detail="No se recibieron archivos.")

    for f in pdfs:
        if not f.filename.lower().endswith(".pdf"):
            raise HTTPException(
                status_code=400, detail=f"'{f.filename}' no es un PDF."
            )

    results: list[tuple[str, bytes, int]] = []  # (nombre.xlsx, bytes, filas)
    errors: list[str] = []
    for f in pdfs:
        content = await f.read()
        xlsx, nrows, err = _convert_one(f.filename, content)
        if err:
            errors.append(f"{f.filename}: {err}")
        else:
            results.append((_safe_filename(f.filename), xlsx, nrows))

    if not results:
        detail = errors[0] if len(pdfs) == 1 else (
            "Ningún PDF pudo convertirse: " + "; ".join(errors)
        )
        raise HTTPException(status_code=400, detail=detail)

    total_rows = sum(r[2] for r in results)

    # Un solo PDF convertido y sin errores -> Excel directo.
    if len(results) == 1 and not errors:
        return Response(
            content=results[0][1],
            media_type=XLSX_MEDIA_TYPE,
            headers={
                "Content-Disposition": f'attachment; filename="{results[0][0]}"',
                "X-Row-Count": str(total_rows),
                "X-File-Count": "1",
                "X-Failed-Count": "0",
            },
        )

    # Varios PDFs -> ZIP con un Excel por PDF (+ _errores.txt si hubo fallos).
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as z:
        used: set[str] = set()
        for name, data, _n in results:
            entry, i = name, 2
            while entry in used:
                entry = f"{Path(name).stem}-{i}.xlsx"
                i += 1
            used.add(entry)
            z.writestr(entry, data)
        if errors:
            z.writestr("_errores.txt", "\n".join(errors))

    return Response(
        content=buffer.getvalue(),
        media_type=ZIP_MEDIA_TYPE,
        headers={
            "Content-Disposition": 'attachment; filename="finiquitos.zip"',
            "X-Row-Count": str(total_rows),
            "X-File-Count": str(len(results)),
            "X-Failed-Count": str(len(errors)),
        },
    )


@app.get("/")
def index():
    return FileResponse(STATIC_DIR / "index.html")


if __name__ == "__main__":
    import uvicorn

    port = int(os.environ.get("PORT", "8000"))
    uvicorn.run(app, host="0.0.0.0", port=port)
