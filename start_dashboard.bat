@echo off
REM Check if the script is running in silent mode
if "%1" == "-silent" goto :silent

REM Relaunch the script in a hidden window
> %temp%\silent.vbs echo Set WshShell = CreateObject("WScript.Shell")
>>%temp%\silent.vbs echo WshShell.Run "cmd /c """"%~f0"""" -silent""", 0, False
cscript //nologo %temp%\silent.vbs
del %temp%\silent.vbs
exit

:silent
REM This is the silent part of the script
title Dashboard Launcher (Silent)

REM Install dependencies quietly
call pip install -r requirements.txt --quiet

REM Run the dashboard application without a console window
start "Dashboard" /B pythonw.exe dashboard.py
exit