@echo off
title Dashboard Launcher

echo Installazione delle dipendenze in corso...
pip install -r requirements.txt

echo.
echo Avvio della dashboard...
python dashboard.py

echo.
echo Applicazione chiusa. Premi un tasto per uscire.
pause