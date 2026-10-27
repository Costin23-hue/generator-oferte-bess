@echo off
REM Construieste "Generator Oferte BESS.exe". Pune acest fisier langa generator_oferte_bess.py
REM si construieste_exe.py, apoi da dublu-click pe el.
REM NU actualizeaza bibliotecile: foloseste exact versiunile care merg deja (pandas 2.3.3 etc.).
cd /d "%~dp0"
set PYTHONUTF8=1
py -3.14 -m pip install pyinstaller
py -3.14 -m pip list > versiuni_folosite.txt
py -3.14 construieste_exe.py > construire_log.txt 2>&1
type construire_log.txt | findstr /C:"Python " /C:"module incluse" /C:"Gata" /C:"ERROR" /C:"Error"
echo.
echo (Detalii complete in construire_log.txt, versiunile bibliotecilor in versiuni_folosite.txt)
pause
