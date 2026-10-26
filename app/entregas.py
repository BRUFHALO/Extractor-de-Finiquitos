"""Lógica de marcado de entregas contra la API (usada por la interfaz web).

Configuración por variables de entorno (ver .env, ignorado por git):
  ENTREGA_TOKEN        -> token 'authtoken' (obligatorio)
  ENTREGA_API_URL      -> endpoint (tiene default)
  ENTREGA_COD_PROVEEDOR-> código de proveedor en modo .txt (default '00077')
"""

from __future__ import annotations

import os
import time
from io import BytesIO
from pathlib import Path

import requests
from openpyxl import load_workbook


def _load_env_file():
    """Carga variables de .env (raíz del proyecto) si no están definidas."""
    env_path = Path(__file__).resolve().parent.parent / ".env"
    if not env_path.exists():
        return
    for line in env_path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            k, v = line.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip())


_load_env_file()

API_URL = os.environ.get(
    "ENTREGA_API_URL",
    "https://pr6joz9r6e.execute-api.us-east-1.amazonaws.com/prod/InsertMessage",
)
AUTH_TOKEN = os.environ.get("ENTREGA_TOKEN", "")
CODIGO_PROVEEDOR_TXT = os.environ.get("ENTREGA_COD_PROVEEDOR", "00077")

UMBRAL_DELAY = 2
DELAY_SEGUNDOS = 5
MAX_ERRORES_SEGUIDOS = 5
TIMEOUT = 15


class EntregaError(Exception):
    pass


def token_configured() -> bool:
    return bool(AUTH_TOKEN)


def _resultado_api(response) -> tuple[str, str]:
    """Devuelve (categoria, detalle): 'entregada' | 'culminada' | 'error'."""
    if response.status_code != 200:
        return "error", f"HTTP {response.status_code}: {response.text[:200]}"

    try:
        body = response.json()
    except ValueError:
        return "error", f"respuesta no JSON: {response.text[:200]}"

    inner = body.get("status")
    answer = str(body.get("answer", ""))

    if str(inner) == "200":
        return "entregada", answer
    if "ya se encuentra culminada" in answer.lower():
        return "culminada", answer
    return "error", f"status {inner}: {answer}"


def enviar(identificador: str, num_factura: str | None, codigo_proveedor: str):
    """Envía el evento 'Entregado'. Si num_factura es None, el campo
    NumeroFactura NO se incluye en el body (entrega solo por ID)."""
    payload = {
        "IdentificadorEvento": str(identificador),
        "Tipo": "Entregado",
        "Mensaje": "Medicamentos Entregado",
        "Adjunto": "",
        "CodigoProveedor": str(codigo_proveedor),
    }
    if num_factura is not None:
        payload["NumeroFactura"] = str(num_factura)

    try:
        response = requests.post(
            API_URL, json=payload,
            headers={"Content-Type": "application/json", "authtoken": AUTH_TOKEN},
            timeout=TIMEOUT,
        )
        return _resultado_api(response)
    except Exception as e:
        return "error", f"conexión: {e}"


def _valor(v) -> str:
    """'225058.0' -> '225058'; '67-39526' -> '67-39526'."""
    if v is None:
        return ""
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


# ------------------------------------------------------------------
# Fuentes de datos
# ------------------------------------------------------------------

def leer_ids_txt(file_bytes: bytes) -> list[str]:
    """IDs de un .txt: uno por línea, ignora vacías y comentarios '#'."""
    ids, vistos = [], set()
    for linea in file_bytes.decode("utf-8", errors="replace").splitlines():
        s = linea.strip()
        if not s or s.startswith("#") or s in vistos:
            continue
        vistos.add(s)
        ids.append(s)
    return ids


def leer_excel(file_bytes: bytes):
    """Lee el xlsx de solicitudes. Detecta la fila de encabezado (la que
    contiene 'Identificador') y crea columnas 'Entregado'/'Resultado' si
    faltan. Devuelve (workbook, worksheet, filas, col_entregado, col_resultado)."""
    wb = load_workbook(BytesIO(file_bytes))
    ws = wb.active

    header_row = None
    for r in range(1, min(ws.max_row or 1, 15) + 1):
        for c in ws[r]:
            if str(c.value).strip().lower() == "identificador":
                header_row = r
                break
        if header_row:
            break
    if not header_row:
        raise EntregaError(
            "No se encontró la fila de encabezado (columna 'Identificador')."
        )

    headers = {}
    for i, c in enumerate(ws[header_row], start=1):
        if c.value is not None:
            headers[str(c.value).strip()] = i

    col_id = headers.get("Identificador")
    col_fac = headers.get("Número Factura")
    col_cod = headers.get("Codigo Proveedor")

    col_ent = headers.get("Entregado")
    if col_ent is None:
        col_ent = ws.max_column + 1
        ws.cell(row=header_row, column=col_ent, value="Entregado")
    col_res = headers.get("Resultado")
    if col_res is None:
        col_res = ws.max_column + 1
        ws.cell(row=header_row, column=col_res, value="Resultado")

    filas = []
    for r in range(header_row + 1, (ws.max_row or 0) + 1):
        idv = _valor(ws.cell(row=r, column=col_id).value)
        if not idv:
            continue
        raw_cod = _valor(ws.cell(row=r, column=col_cod).value) if col_cod else ""
        filas.append({
            "row": r,
            "id": idv,
            "factura": _valor(ws.cell(row=r, column=col_fac).value) if col_fac else "",
            "codigo": raw_cod[2:] if raw_cod.startswith("P-") else (raw_cod or CODIGO_PROVEEDOR_TXT),
            "entregado": _valor(ws.cell(row=r, column=col_ent).value),
        })
    return wb, ws, filas, col_ent, col_res


# ------------------------------------------------------------------
# Procesamiento (genérico, con callback de log)
# ------------------------------------------------------------------

def procesar(items, objetivo: int, log) -> dict:
    """Entrega hasta `objetivo` solicitudes nuevas.

    items: lista de dicts {'id', 'factura'|None, 'codigo'} — el callback
    debe persistir el resultado de cada uno vía marcar(item, cat, detalle).

    log(str) recibe cada línea de progreso.
    Devuelve stats: {'entregadas','culminadas','fallidas','intentos'}."""
    entregadas = culminadas = fallidas = intentos = 0
    con_delay = objetivo > UMBRAL_DELAY
    if con_delay:
        log(f"Meta: {objetivo} entregas -> delay de {DELAY_SEGUNDOS}s entre envíos.")

    errores_seguidos = 0
    for item in items:
        if entregadas >= objetivo:
            break
        if errores_seguidos >= MAX_ERRORES_SEGUIDOS:
            log(f">> {MAX_ERRORES_SEGUIDOS} errores seguidos. "
                "Se detiene el lote (posible problema con la API).")
            break
        intentos += 1
        if con_delay and intentos > 1:
            time.sleep(DELAY_SEGUNDOS)

        idv, fac = item["id"], item.get("factura")
        label = f"ID: {idv}" + (f" | Factura: {fac}" if fac else "")
        log(f"[intento {intentos} | entregadas {entregadas}/{objetivo}] {label}...")

        cat, det = enviar(idv, fac, item.get("codigo") or CODIGO_PROVEEDOR_TXT)
        item["categoria"], item["detalle"] = cat, det

        if cat == "entregada":
            log(f" -> Entregada OK  {det}")
            entregadas += 1
            errores_seguidos = 0
        elif cat == "culminada":
            log(f" -> Ya estaba culminada: {det}")
            culminadas += 1
            errores_seguidos = 0
        else:
            log(f" -> Error: {det}")
            fallidas += 1
            errores_seguidos += 1

    return {
        "entregadas": entregadas,
        "culminadas": culminadas,
        "fallidas": fallidas,
        "intentos": intentos,
    }
