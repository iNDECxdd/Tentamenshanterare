@echo off
setlocal EnableExtensions

REM Kör från samma mapp som .bat ligger
pushd "%~dp0"

REM Sätt kodsida till UTF-8 så å/ä/ö funkar
chcp 65001 >nul

REM (Valfritt) Underlag från GUI skickas som första argument: en textfil.
REM Denna filväg skickas vidare till Kvitto_generator.py.
set "INPUT_FILE=%~1"

set "SCRIPT=%~dp0Kvitto_generator.py"
if not exist "%SCRIPT%" (
  echo [FEL] Hittar inte "%SCRIPT%".
  echo Flytta .bat till samma mapp som Kvitto_generator.py, eller uppdatera skriptets namn.
  goto :end
)

REM Om virtuell miljö finns, kör den först
set "VENV_PY=.venv\Scripts\python.exe"
if exist "%VENV_PY%" (
  set "PYEXE=%VENV_PY%"
) else (
  REM Försök med 'py' (Python launcher), annars 'python'
  where py >nul 2>nul && (set "PYEXE=py") || (set "PYEXE=python")
)

echo Startar: %PYEXE% "%SCRIPT%"
if defined INPUT_FILE echo Underlag: "%INPUT_FILE%"
echo.

if /i "%PYEXE%"=="py" (
  if defined INPUT_FILE (
    py "%SCRIPT%" "%INPUT_FILE%"
  ) else (
    py "%SCRIPT%"
  )
) else (
  if defined INPUT_FILE (
    "%PYEXE%" "%SCRIPT%" "%INPUT_FILE%"
  ) else (
    "%PYEXE%" "%SCRIPT%"
  )
)

set "EC=%ERRORLEVEL%"
echo.
if not "%EC%"=="0" (
  echo [FEL] Programmet avslutades med felkod %EC%.
) else (
  echo Klart utan fel.
)

:end
popd
pause
exit /b %EC%
