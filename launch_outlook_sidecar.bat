@echo off

setlocal enabledelayedexpansion

cd /d %~dp0

set "PYTHONPATH=%~dp0src;%PYTHONPATH%"



where py >nul 2>&1

if %errorlevel%==0 (

    set "PYTHON=py"

) else (

    set "PYTHON=python"

)



%PYTHON% -m pip install --upgrade pip

%PYTHON% -m pip install -r requirements.txt

echo.
echo [Outlook Sidecar] Checking optional PST/OST support...
%PYTHON% -m pip install libpff-python>=20231205
if %errorlevel% neq 0 (
    echo.
    echo [Warning] libpff-python could not be installed automatically.
    echo Install Microsoft Visual C++ Build Tools, then run: pip install libpff-python
)

%PYTHON% -m outlook_sidecar.gui



endlocal

pause

