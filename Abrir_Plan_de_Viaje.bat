@echo off
chcp 65001 >nul
title Plan de Gestion de Viaje - Streamlit
set "URL=http://localhost:8501"
REM Evitar cd con ruta entre comillas que termina en barra invertida (rompe cmd). Se usa el sufijo .
cd /d "%~dp0."

if exist ".venv\Scripts\python.exe" (
  set "PY=.venv\Scripts\python.exe"
  echo Usando entorno virtual .venv
) else (
  set "PY=python"
  echo Usando Python del sistema
)

REM Si falta Streamlit u otras librerias en este entorno, instalarlas una vez
"%PY%" -c "import streamlit" 2>nul
if errorlevel 1 (
  echo.
  echo Instalando dependencias desde requirements.txt ^(solo hace falta la primera vez o si cambian^)...
  "%PY%" -m pip install -r "%~dp0requirements.txt"
  if errorlevel 1 (
    echo Error al instalar. Prueba manualmente: pip install -r requirements.txt
    pause
    exit /b 1
  )
)

REM Si la app ya esta corriendo, solo abrir el navegador y salir.
powershell -NoProfile -Command "try { (Invoke-WebRequest -Uri '%URL%' -UseBasicParsing -TimeoutSec 2) ^| Out-Null; exit 0 } catch { exit 1 }" >nul 2>nul
if not errorlevel 1 (
  echo La app ya esta en ejecucion. Abriendo navegador...
  start "" "%URL%"
  exit /b 0
)

REM Iniciar Streamlit en una nueva ventana para no bloquear este lanzador.
start "Plan de Viaje - Streamlit" "%PY%" -m streamlit run app.py
timeout /t 3 >nul
start "" "%URL%"
exit /b 0
