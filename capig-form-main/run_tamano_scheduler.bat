@echo off
setlocal
rem Ejecuta el scheduler de tamaños en bucle. Ajusta el intervalo con la variable TAMANO_JOB_INTERVAL_SECONDS (min 300).
rem Coloca este .bat en el Programador de tareas o en la carpeta de Inicio para que se ejecute solo.

set "WORKDIR=%~dp0"
set "PYTHON_EXE=%WORKDIR%venv\Scripts\python.exe"

if not exist "%PYTHON_EXE%" (
  echo No se encontro el interprete en %PYTHON_EXE%.
  exit /b 1
)

cd /d "%WORKDIR%"
"%PYTHON_EXE%" -m capig_form.services.tamano_scheduler
