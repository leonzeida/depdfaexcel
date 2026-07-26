#!/bin/bash
# Lanzador: uso  ./f41_a_excel.sh archivo1.pdf [archivo2.pdf ...] | carpeta/
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
exec "$DIR/venv/bin/python3" "$DIR/f41_a_excel.py" "$@"
