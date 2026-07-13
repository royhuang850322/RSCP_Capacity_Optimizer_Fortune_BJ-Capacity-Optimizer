@echo off
setlocal
title Fortune BJ Optimizer - License Generator
cd /d "%~dp0\.."

set "PY_CMD="
set "PYW_CMD="

call :probe_python "python"
if not defined PY_CMD call :probe_python "py -3"
if not defined PY_CMD if exist "%LocalAppData%\Programs\Python\Python313\python.exe" (
    set "PY_CMD=%LocalAppData%\Programs\Python\Python313\python.exe"
)

call :probe_pythonw "pythonw"
if not defined PYW_CMD if exist "%LocalAppData%\Programs\Python\Python313\pythonw.exe" (
    set "PYW_CMD=%LocalAppData%\Programs\Python\Python313\pythonw.exe"
)

echo.
echo Fortune BJ Optimizer - License Generator
echo -----------------------------------------
echo Project folder:
echo   %CD%
echo.

if not defined PY_CMD (
    echo Python was not found on this computer.
    echo Install Python and required packages first, then try again.
    goto :end
)

echo Checking required Python packages...
call %PY_CMD% -c "import tkinter, cryptography"
if errorlevel 1 (
    echo.
    echo Required Python packages are missing or incomplete.
    echo Install project requirements first, then try again.
    goto :end
)

if defined PYW_CMD (
    start "" %PYW_CMD% license_admin\license_tools\license_generator_ui.py
    exit /b 0
)

call %PY_CMD% license_admin\license_tools\license_generator_ui.py
if errorlevel 1 (
    echo.
    echo The license generator did not start successfully.
)

:end
echo.
pause
exit /b 0

:probe_python
%~1 -c "import sys" >nul 2>nul
if not errorlevel 1 set "PY_CMD=%~1"
exit /b 0

:probe_pythonw
%~1 -c "import sys" >nul 2>nul
if not errorlevel 1 set "PYW_CMD=%~1"
exit /b 0
