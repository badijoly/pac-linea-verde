@echo off
title Dashboard Evaluacion PAC 2026 - Linea Verde

echo ============================================================
echo   Iniciando Servidor del Dashboard de Evaluacion PAC 2026
echo ============================================================

set "PYTHON_EXE="

:: 1. Buscar en entornos virtuales locales conocidos
if exist "C:\Users\olfor\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" (
    set PYTHON_EXE="C:\Users\olfor\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
    goto :found_python
)

if exist "C:\Users\olfor\Documents\ComfyUI\.venv\Scripts\python.exe" (
    set PYTHON_EXE="C:\Users\olfor\Documents\ComfyUI\.venv\Scripts\python.exe"
    goto :found_python
)

:: 2. Probar lanzador oficial de Python (py.exe)
py -3 --version >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    set PYTHON_EXE=py -3
    goto :found_python
)

:: 3. Probar si python del sistema funciona
python -c "import sys; sys.exit(0)" >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    set PYTHON_EXE=python
    goto :found_python
)

echo.
echo [ERROR] No se pudo encontrar un ejecutable funcional de Python.
echo Por favor abre directamente el archivo 'index.html' en tu navegador.
echo.
pause
exit /b 1

:found_python
echo [OK] Usando Python: %PYTHON_EXE%
echo.

cd /d "%~dp0"

:: Abrir navegador automaticamente tras 1.5 segundos
start "" cmd /c "timeout /t 2 /nobreak >nul && start http://localhost:8085"

%PYTHON_EXE% server.py
pause
