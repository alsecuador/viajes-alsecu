@echo off
chcp 65001 >nul
title Plan de Gestión de Viaje — Streamlit
cd /d "%~dp0"

if exist ".venv\Scripts\python.exe" (
  set "PY=.venv\Scripts\python.exe"
  echo Usando entorno virtual .venv
) else (
  set "PY=python"
  echo Usando Python del sistema
)

REM Si falta Streamlit u otras librerías en este entorno, instalarlas una vez
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

"%PY%" -m streamlit run app.py

if errorlevel 1 (
  echo.
  echo No se pudo iniciar la app.
  pause
)
