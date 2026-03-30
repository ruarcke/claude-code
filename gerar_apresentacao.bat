@echo off
chcp 65001 >nul
title RCPCC - Gerador de Apresentacao
echo.
echo ========================================
echo   RCK Advogados - Gerador de Apresentacao RCPCC
echo ========================================
echo.

cd /d "%~dp0"

:: Use Python 3.12 specifically
set PYTHON="%USERPROFILE%\AppData\Local\Programs\Python\Python312\python.exe"

if not exist %PYTHON% (
    echo ERRO: Python 3.12 nao encontrado.
    echo Instale em: https://www.python.org/downloads/
    pause
    exit /b 1
)

:: If a file was dragged onto this .bat, use it as argument
if "%~1"=="" (
    %PYTHON% "%~dp0rcpcc_generator.py"
) else (
    %PYTHON% "%~dp0rcpcc_generator.py" "%~1"
)

echo.
pause
