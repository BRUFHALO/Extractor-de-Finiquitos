# Conversor PDF → Excel (Finiquito)

Aplicación web con **drag & drop** que convierte un PDF de pagos en un archivo
Excel con el formato `modelo finiquito`. Detecta automáticamente el formato del
PDF:

- **Formato A — "Confirmación de Pagos"**: montos con `Bs.`, un solo nombre
  (paciente). Las columnas `Fec.Recep.` y `Nombre` (titular) quedan vacías.
- **Formato B — export "Finiquito Beconsult"**: ya trae las 12 columnas del
  Excel (dos nombres, dos cédulas, montos sin `Bs.`). Se parsea por posición de
  caracteres para reconstruir correctamente nombres largos que se solapan con la
  columna siguiente.

## Tecnología
- **Backend:** FastAPI + pdfplumber + openpyxl
- **Frontend:** HTML/JS (sin dependencias)
- **Despliegue:** Docker

## Mapeo PDF → Excel
| Excel | Columna | Origen en PDF |
|---|---|---|
| A | Fact. | # de Factura |
| B | Fecha Fac | Fecha de Ocurrencia |
| C | Fec.Recep. | (vacía) |
| D | C.I. Pac. | C.I del Paciente |
| E | Nombre | Nombre del Paciente |
| F | C.I. Tit. | Certificado |
| G | Nombre | (vacía) |
| H | Total Fac. | Total Facturado |
| I | Monto a pagar | Total Pagado |
| J | Pagado | Total Pagado |
| K | Referencia | Referencia |
| L | Fecha pago | Fecha de Pago |

Los montos `Bs. 83.861,99` se convierten a número (`83861.99`) y las fechas
`dd/mm/yyyy` se guardan como fecha real de Excel.

## Ejecutar en local (Docker)
```bash
docker build -t finiquito .
docker run --rm -p 8000:8000 finiquito
```
Abre http://localhost:8000

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
git clone <tu-repo> && cd "Finiquito Beconsult"
docker build -t finiquito .
docker run -d --restart unless-stopped -p 80:8000 finiquito
```

## Endpoints
- `GET /` — interfaz web
- `POST /convert` — recibe un PDF (`multipart/form-data`, campo `file`) y devuelve el `.xlsx`
- `GET /health` — healthcheck
