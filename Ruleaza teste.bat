@echo off
rem Ruleaza testele automate ale programului (date inventate, nu porneste Excel / Drive)
cd /d "%~dp0"
py -3.14 -m unittest discover -s teste -v
pause
