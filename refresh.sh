#!/usr/bin/env bash
# Semanal: python3 build.py sobre source.xlsx. Imprime UNCHANGED / CHANGED ... / ERROR ...
# build.py lee la pestaña Viajes (no Resumen) y agrupa cada crucero como un viaje.
# Tiene que usarse junto con assets/app.js de esta misma versión: el HTML viejo
# vuelve a contar cada tramo de crucero por separado.
set -u
cd "$(dirname "$0")"
python3 build.py source.xlsx "$@"
