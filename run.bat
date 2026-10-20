@echo off
title Qiuzhao Application Tracker
cd /d "%~dp0"

echo.
echo   ==============================================================
echo    Qiuzhao Application Tracker  -  local entry service
echo   ==============================================================
echo.
echo   Looking for Python 3 ...

set "PYEXE="

for %%C in (python.exe py.exe) do (
  if not defined PYEXE (
    for /f "delims=" %%P in ('where %%C 2^>nul') do (
      if not defined PYEXE (
        "%%P" -c "import sys;raise SystemExit(0 if sys.version_info[0]>=3 else 1)" >nul 2>nul
        if not errorlevel 1 set "PYEXE=%%P"
      )
    )
  )
)

if not defined PYEXE (
  for %%D in (
    "%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
    "C:\Python313\python.exe"
    "C:\Python312\python.exe"
    "C:\Python311\python.exe"
    "D:\Python\python.exe"
    "D:\Python3\python.exe"
  ) do (
    if not defined PYEXE if exist %%D set "PYEXE=%%~D"
  )
)

if not defined PYEXE (
  echo.
  echo   [ERROR] Python 3 was not found on this computer.
  echo.
  echo   Install it from:  https://www.python.org/downloads/
  echo   During setup, tick the box "Add python.exe to PATH".
  echo   Then run this file again.
  echo.
  pause
  exit /b 1
)

echo   Python : %PYEXE%

"%PYEXE%" -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
  echo.
  echo   openpyxl is missing, installing it now ...
  "%PYEXE%" -m pip install openpyxl
  if errorlevel 1 (
    echo.
    echo   [ERROR] Could not install openpyxl automatically.
    echo   Open a terminal and run:  pip install openpyxl
    echo.
    pause
    exit /b 1
  )
)

echo.
echo   --------------------------------------------------------------
echo    Starting.  KEEP THIS WINDOW OPEN while you use the page.
echo    The browser will open automatically in a second.
echo    If it does not, open this address manually:
echo.
echo        http://127.0.0.1:8760
echo.
echo    Closing this window stops the service.
echo    Your data is already saved in the .xlsx file.
echo   --------------------------------------------------------------
echo.

"%PYEXE%" "%~dp0app.py"
set "RC=%ERRORLEVEL%"

echo.
if not "%RC%"=="0" echo   [ERROR] The service exited with code %RC%.
echo   Service stopped. Your data is already saved in the .xlsx file.
echo.
pause
