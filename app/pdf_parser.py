"""Extrae las filas de pago de los PDF y las devuelve con el esquema del
Excel `modelo finiquito` (12 columnas).

Soporta dos formatos de PDF, detectados por la ESTRUCTURA del contenido
(sin depender de títulos ni textos específicos del documento):

- Formato A: filas "factura fecha clave cert nombre cipac monto monto
  fecha referencia" (montos opcionalmente con prefijo 'Bs.', un solo
  nombre).
- Formato B: tabla con las mismas 12 columnas del Excel (encabezado con
  'Fec.Recep.', dos nombres). Se parsea por posición de caracteres para
  tolerar nombres largos que se solapan con la columna siguiente.

Esquema unificado de cada fila (claves = columnas A..L del Excel):
  fact, fecha_fac, fec_recep, ci_pac, nombre_pac, ci_tit, nombre_tit,
  total_fac, monto_pagar, pagado, referencia, fecha_pago
"""

from __future__ import annotations

import io
import re
from datetime import date, datetime

import pdfplumber

# ---------------------------------------------------------------------------
# Utilidades comunes
# ---------------------------------------------------------------------------

def _to_number(raw: str) -> float | None:
    """'83.861,99' -> 83861.99 ; '946,46' -> 946.46"""
    cleaned = (raw or "").strip().replace(".", "").replace(",", ".")
    if not cleaned:
        return None
    return float(cleaned)


def _to_date(raw: str) -> date | None:
    """Acepta dd/mm/yyyy y d/m/yyyy."""
    s = (raw or "").strip()
    if not s:
        return None
    try:
        return datetime.strptime(s, "%d/%m/%Y").date()
    except ValueError:
        return None


# ---------------------------------------------------------------------------
# Formato A: CONFIRMACIÓN DE PAGOS
# ---------------------------------------------------------------------------

ROW_RE_A = re.compile(
    r"^(?P<factura>[A-Za-z0-9]+(?:-[A-Za-z0-9]+)?)\s+"
    r"(?P<focurr>\d{1,2}/\d{1,2}/\d{4})\s+"
    r"(?P<clave>\d+-\d+)\s+"
    r"(?P<cert>\S+)\s+"
    r"(?P<nombre>.+?)\s+"
    r"(?P<cipac>(?:[A-Za-z]-)?\d+|SC)\s+"
    r"(?:Bs\.\s*)?(?P<tfac>[\d.,]+)\s+"
    r"(?:Bs\.\s*)?(?P<tpag>[\d.,]+)\s+"
    r"(?P<fpago>\d{1,2}/\d{1,2}/\d{4})\s+"
    r"(?P<ref>\S+)\s*$",
    re.UNICODE,
)

NAME_CONT_RE = re.compile(r"^[A-ZÁÉÍÓÚÑÜ .,'\-]+$", re.UNICODE)

PROVIDER_RE = re.compile(r"PROVEEDOR\s+(?P<prov>.+?)\s+RIF", re.UNICODE)

NOISE_KEYWORDS = (
    "CONFIRMACIÓN DE PAGOS", "FINIQUITO DE PAGOS", "PROVEEDOR",
    "# de Factura", "Página", "FECHA",
    "Pagos en", "Desde:", "Estimados", "Atención", "Departamento", "TOTALES",
    "NOTA", "Ocurrencia", "Paciente", "MiSaludyMás", "RIF", "Certificado",
)


def _is_noise(line: str) -> bool:
    return any(kw in line for kw in NOISE_KEYWORDS)


def _parse_format_a(full_text: str) -> list[dict]:
    rows: list[dict] = []
    for raw_line in full_text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        m = ROW_RE_A.match(line)
        if m:
            pagado = _to_number(m.group("tpag"))
            rows.append(
                {
                    "fact": m.group("factura"),
                    "fecha_fac": _to_date(m.group("focurr")),
                    "fec_recep": None,
                    "ci_pac": m.group("cipac"),
                    "nombre_pac": re.sub(r"\s+", " ", m.group("nombre")).strip(),
                    "ci_tit": m.group("cert"),
                    "nombre_tit": None,
                    "total_fac": _to_number(m.group("tfac")),
                    "monto_pagar": pagado,
                    "pagado": pagado,
                    "referencia": m.group("ref"),
                    "fecha_pago": _to_date(m.group("fpago")),
                }
            )
            continue
        # Continuación del nombre del paciente en la línea siguiente.
        if rows and not _is_noise(line) and NAME_CONT_RE.match(line):
            rows[-1]["nombre_pac"] = (
                rows[-1]["nombre_pac"] + " " + line.strip()
            ).strip()
    return rows


# ---------------------------------------------------------------------------
# Formato B: export "Finiquito Beconsult" (mismas columnas del Excel)
# ---------------------------------------------------------------------------

# El encabezado puede venir con o sin espacios entre columnas (pdfplumber no
# siempre emite glifos de espacio), y la puntuación puede variar. Se detecta
# por palabras clave, no por el texto exacto: toda tabla formato B tiene
# 'Fec.Recep.' y 'Fecha'.
def _looks_like_header_b(norm: str) -> bool:
    low = norm.lower()
    return "recep" in low and "fecha" in low

# Una fila de datos contiene al menos una fecha d/m/aaaa.
DATA_DATE_RE = re.compile(r"\d{1,2}/\d{1,2}/\d{4}", re.UNICODE)


def _cluster_rows(chars, tol: float = 3.0) -> list[list]:
    """Agrupa caracteres en líneas por su coordenada vertical ('top')."""
    groups: list[tuple[float, list]] = []
    for c in sorted(chars, key=lambda z: (z["top"], z["x0"])):
        placed = False
        for g in groups:
            if abs(c["top"] - g[0]) <= tol:
                g[1].append(c)
                placed = True
                break
        if not placed:
            groups.append((c["top"], [c]))
    return [sorted(g[1], key=lambda z: z["x0"]) for g in sorted(groups, key=lambda z: z[0])]


def _detect_columns(header_chars, gap: float = 5.0) -> list[float]:
    """Devuelve el borde izquierdo (x0) de cada columna del encabezado."""
    lefts: list[float] = []
    prev_x1 = None
    for c in sorted(header_chars, key=lambda z: z["x0"]):
        if c["text"].strip() == "" and prev_x1 is not None:
            # Los espacios cuentan para el ancho pero no inician columna.
            prev_x1 = c["x1"]
            continue
        if prev_x1 is None or (c["x0"] - prev_x1) > gap:
            lefts.append(c["x0"])
        prev_x1 = c["x1"]
    return lefts


def _txt(chars) -> str:
    return "".join(c["text"] for c in sorted(chars, key=lambda z: z["x0"]))


def _split_by_gap(chars, gap: float = 8.0) -> list[list]:
    groups: list[list] = []
    prev_x1 = None
    for c in sorted(chars, key=lambda z: z["x0"]):
        if prev_x1 is None or (c["x0"] - prev_x1) > gap:
            groups.append([])
        groups[-1].append(c)
        prev_x1 = c["x1"]
    return groups


def _parse_row_b(rowchars, lefts: list[float]) -> dict | None:
    if len(lefts) < 12:
        return None
    L = lefts + [float("inf")]  # L[12] = infinito

    def region(a: float, b: float):
        return [c for c in rowchars if a <= c["x0"] < b]

    def only(chars, allowed: str) -> str:
        return "".join(ch for ch in _txt(chars) if ch in allowed or ch.isdigit())

    def date_of(chars):
        return _to_date("".join(ch for ch in _txt(chars) if ch.isdigit() or ch == "/"))

    def name_of(chars) -> str:
        raw = _txt([c for c in chars if not c["text"].isdigit()])
        return re.sub(r"\s+", " ", raw).strip(" ,.-")

    # Nombre paciente: región propia + desbordamiento (letras) sobre C.I. Tit.
    name_pac = name_of([c for c in rowchars if L[4] <= c["x0"] < L[6]])
    # Nombre titular: región propia + desbordamiento (solo letras) hacia Total Fac.
    tit_base = [c for c in rowchars if L[6] <= c["x0"] < L[7]]
    tit_over = [c for c in rowchars if c["x0"] >= L[7] and c["text"].isalpha()]
    name_tit = re.sub(r"\s+", " ", _txt(tit_base + tit_over)).strip(" ,.-")
    # quitar dígitos que pudieran haber quedado en la base
    name_tit = re.sub(r"\d+", "", name_tit).strip(" ,.-")

    fact = _txt(region(L[0], L[1])).strip()
    ci_pac = _txt(region(L[3], L[4])).strip()
    ci_tit_digits = "".join(ch for ch in _txt(region(L[5], L[6])) if ch.isdigit())
    ci_tit = ci_tit_digits or _txt(region(L[5], L[6])).strip()

    # Los tres montos: caracteres numéricos entre Total Fac. y Referencia.
    amt_chars = [
        c for c in rowchars
        if L[7] <= c["x0"] < L[10] and (c["text"].isdigit() or c["text"] in ".,")
    ]
    amounts = [_to_number(_txt(g)) for g in _split_by_gap(amt_chars)]
    amounts = [a for a in amounts if a is not None]
    if len(amounts) < 3:
        return None
    total_fac, monto_pagar, pagado = amounts[0], amounts[1], amounts[2]

    referencia = "".join(
        ch for ch in _txt(region(L[10], L[11])) if ch.isdigit() or ch == "-"
    )

    return {
        "fact": fact,
        "fecha_fac": date_of(region(L[1], L[2])),
        "fec_recep": date_of(region(L[2], L[3])),
        "ci_pac": ci_pac,
        "nombre_pac": name_pac,
        "ci_tit": ci_tit,
        "nombre_tit": name_tit,
        "total_fac": total_fac,
        "monto_pagar": monto_pagar,
        "pagado": pagado,
        "referencia": referencia,
        "fecha_pago": date_of(region(L[11], L[12])),
    }


def _parse_format_b(pages) -> list[dict]:
    rows: list[dict] = []
    lefts: list[float] | None = None
    for page in pages:
        for rowchars in _cluster_rows(page.chars):
            norm = re.sub(r"\s+", " ", _txt(rowchars)).strip()
            if not norm:
                continue
            if _looks_like_header_b(norm):
                lefts = _detect_columns(rowchars)
                continue
            if lefts is None or not DATA_DATE_RE.search(norm):
                continue
            rec = _parse_row_b(rowchars, lefts)
            if rec:
                rows.append(rec)
    return rows


def _provider_format_b(full_text: str) -> str | None:
    for line in full_text.splitlines():
        s = line.strip()
        if not s:
            continue
        if _looks_like_header_b(s):
            break
        return s
    return None


# ---------------------------------------------------------------------------
# Punto de entrada
# ---------------------------------------------------------------------------

def parse_pdf(file_bytes: bytes) -> tuple[list[dict], str | None]:
    """Devuelve (filas, nombre_proveedor).

    La detección es estructural: no depende de títulos ni textos fijos.
    Se prueba primero el formato A (filas con patrón de pago) y, si no hay
    coincidencias, el formato B (tabla de 12 columnas con 'Fec.Recep.').
    """
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        pages = pdf.pages
        full_text = "\n".join((p.extract_text() or "") for p in pages)

        # Formato A: detección por las filas de datos mismas.
        rows_a = _parse_format_a(full_text)
        if rows_a:
            m = PROVIDER_RE.search(full_text)
            provider = m.group("prov").strip() if m else None
            return rows_a, provider

        # Formato B: encabezado con 'Fec.Recep.' + 'Fecha'.
        if _looks_like_header_b(re.sub(r"\s+", " ", full_text)):
            return _parse_format_b(pages), _provider_format_b(full_text)

    return [], None
