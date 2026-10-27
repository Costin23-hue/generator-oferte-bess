#!/bin/bash
# Construieste "Generator Oferte BESS.app" pe Mac.
# Pune acest fisier langa generator_oferte_bess.py si construieste_exe.py, apoi in Terminal:
#     bash construieste_app.command
cd "$(dirname "$0")"
python3 -m pip install pyinstaller "pandas==2.3.3" openpyxl xlwings python-pptx
python3 -m pip list > versiuni_folosite.txt
python3 construieste_exe.py > construire_log.txt 2>&1
grep -E "Python |module incluse|Gata|ERROR|Error" construire_log.txt
echo
echo "(Detalii complete in construire_log.txt, versiunile bibliotecilor in versiuni_folosite.txt)"
