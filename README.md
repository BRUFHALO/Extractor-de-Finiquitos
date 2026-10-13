# Conversor PDF → Excel (Finiquito)

Aplicación web con **drag & drop** que convierte PDFs de pagos en archivos
Excel con el formato `modelo finiquito`. Acepta **uno o varios PDFs a la vez**
(si se suben varios, se descarga un `.zip` con un Excel por PDF). La detección
del formato es **estructural** — no depende de títulos ni textos específicos:

- **Formato A — listado de pagos**: filas con patrón
  `factura fecha clave certificado nombre cédula monto monto fecha referencia`
  (montos con o sin prefijo `Bs.`). Las columnas `Fec.Recep.` y `Nombre`
  (titular) quedan en `0`.
- **Formato B — tabla de 12 columnas**: ya trae el formato completo (encabezado
  con `Fec.Recep.`, dos nombres, dos cédulas). Se parsea por posición de
  caracteres para reconstruir correctamente nombres largos que se solapan con
  la columna siguiente.

Las celdas sin dato se escriben como `0`.

## Tecnología
- **Backend:** FastAPI + pdfplumber + openpyxl
- **Frontend:** HTML/JS (sin dependencias)
- **Despliegue:** Docker

## Mapeo PDF → Excel
| Excel | Columna | Origen en PDF |
|---|---|---|
| A | Fact. | # de Factura |
| B | Fecha Fac | Fecha de Ocurrencia |
| C | Fec.Recep. | (0 en formato A / fecha de recepción en formato B) |
| D | C.I. Pac. | C.I del Paciente |
| E | Nombre | Nombre del Paciente |
| F | C.I. Tit. | Certificado |
| G | Nombre | (0 en formato A / nombre del titular en formato B) |
| H | Total Fac. | Total Facturado |
| I | Monto a pagar | Total Pagado |
| J | Pagado | Total Pagado |
| K | Referencia | Referencia |
| L | Fecha pago | Fecha de Pago |

Los montos `Bs. 83.861,99` se convierten a número (`83861.99`) y las fechas
`dd/mm/yyyy` se guardan como fecha real de Excel.

## Ejecutar en local (script)
```bash
./run.sh          # levanta la app (Docker si está disponible, si no, venv+uvicorn)
./run.sh down     # detiene el contenedor
./run.sh logs     # muestra los logs
```
Abre http://localhost:8000 — arrastra los PDFs y descarga el Excel/ZIP.

## Ejecutar con Docker Compose (manual)
```bash
docker compose up -d
docker compose down   # para detener
```

## Ejecutar con Docker (sin Compose)
```bash
docker build -t finiquito .
docker run --rm -p 8000:8000 finiquito
```

## Ejecutar en local (sin Docker)
```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --reload
```

## Desplegar en DigitalOcean

### Opción A — App Platform (recomendada)
1. Sube este proyecto a un repositorio (GitHub/GitLab).
2. En DigitalOcean: **Create → Apps** y conecta el repo.
3. Detecta el `Dockerfile` automáticamente. App Platform inyecta la variable
   `PORT`; el contenedor ya la respeta.
4. Deploy. Obtendrás una URL pública.

### Opción B — Droplet con Docker
```bash
# en el Droplet (con Docker instalado)
git clone <tu-repo> && cd <carpeta-del-repo>
docker build -t finiquito .
docker run -d --restart unless-stopped -p 80:8000 finiquito
```

## Endpoints
- `GET /` — interfaz web
- `POST /convert` — recibe uno o varios PDFs (`multipart/form-data`, campo
  `files` repetido). Con un solo PDF devuelve el `.xlsx` directo; con varios
  devuelve `finiquitos.zip` (un Excel por PDF + `_errores.txt` si alguno falló).
  Headers de respuesta: `X-Row-Count`, `X-File-Count`, `X-Failed-Count`.
- `GET /health` — healthcheck

## Notas
- Los PDFs y archivos Excel de ejemplo no se incluyen en el repositorio por
  contener datos sensibles. Coloca tus propios archivos de prueba en el
  directorio del proyecto para validar localmente.
- Las facturas alfanuméricas (ej. `A-68899`) se conservan como texto.
- Las cédulas `SC` se conservan como texto.
