#!/bin/bash
# Instalarea pe Mac (o singură dată), după ce Python 3.14 e instalat de pe python.org.
# Rulare în Terminal:
#   curl -fsSL https://raw.githubusercontent.com/Costin23-hue/generator-oferte-bess/main/instalare_mac.sh | bash
set -e
FOLDER="$HOME/Generator Oferte BESS"
SURSA="https://raw.githubusercontent.com/Costin23-hue/generator-oferte-bess/main"

# Python de pe python.org (nu cel din macOS, care e vechi)
PY=$(ls -d /Library/Frameworks/Python.framework/Versions/3.*/bin/python3 2>/dev/null | sort -V | tail -1)
[ -n "$PY" ] || { echo "Lipsește Python de pe python.org: instalează-l de pe https://www.python.org/downloads/"; exit 1; }
echo "Python: $("$PY" --version)"

# Python de pe python.org are nevoie de certificatele pentru conexiunile HTTPS (Microsoft, GitHub)
for c in /Applications/Python\ 3*/Install\ Certificates.command; do
  [ -f "$c" ] && bash "$c" >/dev/null 2>&1 || true
done

echo "Instalez bibliotecile (pandas, openpyxl, python-pptx)..."
"$PY" -m pip install --quiet --upgrade pip
"$PY" -m pip install --quiet "pandas==2.3.3" openpyxl python-pptx pypdf anthropic \
  || "$PY" -m pip install --user --quiet "pandas==2.3.3" openpyxl python-pptx pypdf anthropic

mkdir -p "$FOLDER"
cd "$FOLDER"
for F in generator_oferte_bess.py bess_setari.py bess_logica.py bess_excel.py bess_prezentare.py bess_drive.py \
         bess_generare.py bess_atr.py bess_motor.py bess_fereastra.py; do
  curl -fsSL -o "$F" "$SURSA/$F"
done
curl -fsSL -o "Porneste Generator Oferte.command" "$SURSA/Porneste%20Generator%20Oferte.command"
chmod +x "Porneste Generator Oferte.command"

# scurtătură pe Desktop
ln -sf "$FOLDER/Porneste Generator Oferte.command" "$HOME/Desktop/Generator Oferte BESS.command"

"$PY" -c "import pandas, openpyxl, pptx, tkinter; print('Biblioteci OK')"
echo
echo "Gata. Pornește programul cu dublu-click pe 'Generator Oferte BESS' de pe Desktop."
echo "(La prima pornire se deschide browserul pentru logarea cu contul Microsoft de firmă.)"
