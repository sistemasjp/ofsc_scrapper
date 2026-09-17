#!/usr/bin/env bash
# Genera el reporte de consumo de materiales para una fecha, sin intervención
# manual (usa las credenciales de .env para el login automático).
#
# Uso:
#   ./reporte.sh 2026-09-10
#
# Requiere haber hecho la instalación una sola vez (ver README.md):
#   python3 -m venv venv && source venv/bin/activate
#   pip install -r requirements.txt && playwright install chromium
#   cp .env.example .env   (y completar OFSC_USERNAME / OFSC_PASSWORD)

set -euo pipefail
cd "$(dirname "$0")"

FECHA="${1:?Uso: ./reporte.sh AAAA-MM-DD}"

if [ ! -f ".env" ]; then
    echo "No existe .env. Copia .env.example a .env y completa tus credenciales antes de correr esto." >&2
    exit 1
fi

source venv/bin/activate
python ofsc_scraper.py --date "$FECHA" --auto-login --headless --output "consumo_${FECHA}.xlsx"
