"""API FastAPI: sirve la interfaz drag & drop y convierte PDF(s) a Excel.
También expone /entregas para marcar solicitudes como entregadas vía API."""

from __future__ import annotations

import io
import os
import re
import threading
import uuid
import zipfile
from io import BytesIO
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse, Response

from app import entregas
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


# ------------------------------------------------------------------
# Marcado de entregas (Excel o TXT de IDs) — jobs en segundo plano
# ------------------------------------------------------------------

JOBS: dict[str, dict] = {}
JOBS_LOCK = threading.Lock()


def _run_entrega_job(job: dict, filename: str, content: bytes, limite: int | None):
    def log(line: str):
        job["log"].append(line)

    try:
        lower = filename.lower()
        if lower.endswith(".txt"):
            ids = entregas.leer_ids_txt(content)
            if not ids:
                raise entregas.EntregaError("El .txt no contiene IDs.")
            items = [{"id": i, "factura": None, "codigo": entregas.CODIGO_PROVEEDOR_TXT}
                     for i in ids]
            job["total"] = len(items)
            objetivo = limite or len(items)
            stats = entregas.procesar(items, objetivo, log)
            job.update(stats)

            # Archivo de resultados: 'id;Estado' por línea
            lines = []
            for it in items:
                estado = {"entregada": "Entregado", "culminada": "Ya culminada"}.get(
                    it.get("categoria"), f"Error: {it.get('detalle', '')[:80]}"
                    if it.get("categoria") else "Pendiente")
                lines.append(f"{it['id']};{estado}")
            stem = Path(filename).stem or "ids"
            job["result_name"] = f"{stem}_resultado.txt"
            job["result_bytes"] = "\n".join(lines).encode("utf-8")

        elif lower.endswith(".xlsx"):
            wb, ws, filas, col_ent, col_res = entregas.leer_excel(content)
            items = [
                {"id": f["id"], "factura": f["factura"], "codigo": f["codigo"],
                 "_row": f["row"]}
                for f in filas if f["entregado"] != "SI"
            ]
            if not items:
                raise entregas.EntregaError(
                    "No hay solicitudes pendientes (todas marcadas como entregadas)."
                )
            job["total"] = len(items)
            objetivo = limite or len(items)
            stats = entregas.procesar(items, objetivo, log)
            job.update(stats)

            for it in items:
                cat = it.get("categoria")
                if cat in ("entregada", "culminada"):
                    ws.cell(row=it["_row"], column=col_ent, value="SI")
                    ws.cell(row=it["_row"], column=col_res,
                            value="Entregado" if cat == "entregada" else "Ya culminada")
                elif cat == "error":
                    ws.cell(row=it["_row"], column=col_res,
                            value=f"Error: {it.get('detalle','')[:120]}")
            buf = BytesIO()
            wb.save(buf)
            stem = Path(filename).stem or "solicitudes"
            job["result_name"] = f"{stem}_actualizado.xlsx"
            job["result_bytes"] = buf.getvalue()
        else:
            raise entregas.EntregaError("El archivo debe ser .txt o .xlsx")

        log(f"--- Fin: {job['entregadas']} entregadas, "
            f"{job['culminadas']} ya culminadas, {job['fallidas']} con error ---")
        job["state"] = "done"
    except entregas.EntregaError as e:
        job["state"] = "error"
        job["error"] = str(e)
    except Exception as e:
        job["state"] = "error"
        job["error"] = f"error inesperado: {e}"
    finally:
        JOBS_LOCK.release()


@app.post("/api/entregas")
async def start_entrega(file: UploadFile = File(...), limite: int | None = Form(None)):
    if not entregas.token_configured():
        raise HTTPException(500, "Falta configurar ENTREGA_TOKEN en el archivo .env")
    if not file.filename:
        raise HTTPException(400, "No se recibió archivo.")
    if not file.filename.lower().endswith((".txt", ".xlsx")):
        raise HTTPException(400, "El archivo debe ser .txt o .xlsx")

    content = await file.read()
    if not content:
        raise HTTPException(400, "El archivo está vacío.")

    if not JOBS_LOCK.acquire(blocking=False):
        raise HTTPException(409, "Ya hay un proceso de entregas en curso. Espera a que termine.")

    job_id = uuid.uuid4().hex[:12]
    JOBS[job_id] = {
        "state": "running", "log": [], "total": 0,
        "entregadas": 0, "culminadas": 0, "fallidas": 0, "intentos": 0,
        "result_bytes": None, "result_name": None, "error": None,
    }
    threading.Thread(
        target=_run_entrega_job,
        args=(JOBS[job_id], file.filename, content, limite),
        daemon=True,
    ).start()
    return {"job_id": job_id}


@app.post("/api/entregas/id")
async def marcar_un_id(identificador: str = Form(...)):
    """Marca un solo ID como entregado (síncrono, sin NumeroFactura)."""
    if not entregas.token_configured():
        raise HTTPException(500, "Falta configurar ENTREGA_TOKEN en el archivo .env")
    identificador = identificador.strip()
    if not identificador:
        raise HTTPException(400, "El ID no puede estar vacío.")

    categoria, detalle = entregas.enviar(
        identificador, None, entregas.CODIGO_PROVEEDOR_TXT
    )
    return {"id": identificador, "categoria": categoria, "detalle": detalle}


@app.get("/api/entregas/{job_id}")
def entrega_status(job_id: str):
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(404, "Job no encontrado.")
    return JSONResponse({
        "state": job["state"],
        "log": job["log"],
        "total": job["total"],
        "intentos": job["intentos"],
        "entregadas": job["entregadas"],
        "culminadas": job["culminadas"],
        "fallidas": job["fallidas"],
        "error": job["error"],
        "result_name": job["result_name"],
    })


@app.get("/api/entregas/{job_id}/download")
def entrega_download(job_id: str):
    job = JOBS.get(job_id)
    if not job or job["state"] != "done" or not job["result_bytes"]:
        raise HTTPException(404, "Resultado no disponible.")
    media = ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
             if job["result_name"].endswith(".xlsx") else "text/plain")
    return Response(
        content=job["result_bytes"], media_type=media,
        headers={"Content-Disposition": f'attachment; filename="{job["result_name"]}"'},
    )


@app.get("/entregas")
def entregas_page():
    return FileResponse(STATIC_DIR / "entregas.html")


if __name__ == "__main__":
    import uvicorn

    port = int(os.environ.get("PORT", "8000"))
    uvicorn.run(app, host="0.0.0.0", port=port)
