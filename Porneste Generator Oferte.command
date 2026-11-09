#!/bin/bash
# Pornește Generator Oferte BESS pe Mac (dublu-click). Stă în același folder cu generator_oferte_bess.py.
# Folosește Python-ul de pe python.org (cel din macOS e vechi și fără interfața grafică necesară).
cd "$(dirname "$0")"
PY=$(ls -d /Library/Frameworks/Python.framework/Versions/3.*/bin/python3 2>/dev/null | sort -V | tail -1)
exec "${PY:-python3}" generator_oferte_bess.py
