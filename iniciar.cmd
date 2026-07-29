@echo off
setlocal EnableExtensions

set "PROJECT_ROOT=%~dp0"
set "VENV_PYTHON=%PROJECT_ROOT%.venv\Scripts\python.exe"
set "PAUSE_AT_END=0"

if "%~1"=="" set "PAUSE_AT_END=1"

if not exist "%VENV_PYTHON%" goto :setup
goto :run

:setup
echo Ambiente virtual nao encontrado. Preparando o projeto...
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%PROJECT_ROOT%setup.ps1"
set "SETUP_EXIT=%ERRORLEVEL%"

if "%SETUP_EXIT%"=="0" goto :run

echo A preparacao do ambiente falhou com o codigo %SETUP_EXIT%.
if "%PAUSE_AT_END%"=="1" pause
endlocal & exit /b %SETUP_EXIT%

:run
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%PROJECT_ROOT%run.ps1" %*
set "RUN_EXIT=%ERRORLEVEL%"

if "%PAUSE_AT_END%"=="1" echo.
if "%PAUSE_AT_END%"=="1" pause
endlocal & exit /b %RUN_EXIT%
