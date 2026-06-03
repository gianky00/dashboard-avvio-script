@echo off
setlocal
set PYTHONUTF8=1

:: Controllo se l'argomento -hidden è presente
if "%~1"=="-hidden" goto :launch

:hide
:: Crea un piccolo script VBS temporaneo per rilanciare questo batch in modalità nascosta
echo Set shell = CreateObject("WScript.Shell") > "%temp%\launch_hidden.vbs"
echo shell.Run """%~f0"" -hidden", 0, false >> "%temp%\launch_hidden.vbs"
wscript "%temp%\launch_hidden.vbs"
del "%temp%\launch_hidden.vbs"
exit /b

:launch
:: Verifica se poetry è installato
where poetry >nul 2>nul
if %errorlevel% neq 0 exit /b

:: Esegue l'applicazione usando pythonw (senza finestra console) tramite poetry
:: Usiamo il modulo invece dell'entry point per avere più controllo sull'eseguibile python
poetry run pythonw -m dashboard_app
