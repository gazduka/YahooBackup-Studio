@echo off
setlocal
cd /d "%~dp0"
set "PY=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
if not exist "%PY%" (
    where py >nul 2>&1
    if not errorlevel 1 (py -3 "%~dp0YahooBackupStudio.py" & goto end)
    where python >nul 2>&1
    if not errorlevel 1 (python "%~dp0YahooBackupStudio.py" & goto end)
    echo Python non trovato. Installare Python 3.13 oppure modificare il percorso in questo file.
    pause
    goto end
)
"%PY%" "%~dp0YahooBackupStudio.py"
if errorlevel 1 pause
:end
endlocal
