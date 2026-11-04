@echo off
setlocal
set "PY=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
if not exist "%PY%" (
  echo Python 3.13 non trovato nel percorso previsto.
  pause
  exit /b 1
)
"%PY%" -m pip install IMAPClient aspose-email-foss
pause
