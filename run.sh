#!/usr/bin/env bash
# Ejecuta el conversor PDF → Excel en local.
#
# Uso:
#   ./run.sh          -> levanta la app (Docker si está disponible, si no, venv+uvicorn)
#   ./run.sh down     -> detiene el contenedor Docker
#   ./run.sh logs     -> muestra los logs del contenedor
#
# URL: http://localhost:8000

set -euo pipefail
cd "$(dirname "$0")"

PORT="${PORT:-8000}"

compose_cmd() {
  if docker compose version >/dev/null 2>&1; then
    echo "docker compose"
  elif command -v podman-compose >/dev/null 2>&1; then
    echo "podman-compose"
  else
    return 1
  fi
}

case "${1:-}" in
  down)
    if CMD=$(compose_cmd); then
      $CMD down
      echo "Detenido."
    else
      echo "No hay docker/podman compose instalado."
    fi
    exit 0
    ;;
  logs)
    if CMD=$(compose_cmd); then
      $CMD logs -f
    else
      echo "No hay docker/podman compose instalado."
    fi
    exit 0
    ;;
  "") ;;
  *)
    echo "Uso: ./run.sh [down|logs]"
    exit 1
    ;;
esac

# --- Opción 1: Docker Compose ---
if CMD=$(compose_cmd); then
  echo ">> Levantando con: $CMD up -d --build"
  $CMD up -d --build

  echo ">> Esperando healthcheck..."
  for i in $(seq 1 30); do
    if curl -sf "http://localhost:${PORT}/health" >/dev/null 2>&1; then
      echo ">> Listo ✓  http://localhost:${PORT}"
      exit 0
    fi
    sleep 2
  done
  echo ">> El contenedor arrancó pero /health no respondió aún."
  echo ">> Revisa con: ./run.sh logs"
  exit 1
fi

# --- Opción 2: venv + uvicorn (sin Docker) ---
echo ">> Docker no disponible, usando entorno virtual local."
if [ ! -d .venv ]; then
  echo ">> Creando .venv..."
  python3 -m venv .venv
fi
# shellcheck disable=SC1091
source .venv/bin/activate

if ! python -c "import fastapi, uvicorn, pdfplumber, openpyxl" >/dev/null 2>&1; then
  echo ">> Instalando dependencias..."
  pip install -r requirements.txt
fi

echo ">> Listo ✓  http://localhost:${PORT}  (Ctrl+C para detener)"
exec uvicorn app.main:app --host 0.0.0.0 --port "$PORT"
