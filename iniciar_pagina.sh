#!/bin/bash
# Doble clic (o ejecutar) para abrir la página de conversión PDF -> Excel.
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

if (exec 3<>/dev/tcp/127.0.0.1/5000) 2>/dev/null; then
  exec 3>&-
  # Ya está abierta: solo abrir una pestaña nueva apuntando a la página.
  xdg-open http://127.0.0.1:5000 >/dev/null 2>&1
else
  "$DIR/venv/bin/python3" "$DIR/webapp/app.py"
fi
