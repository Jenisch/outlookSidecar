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

%PYTHON% -m outlook_sidecar.gui



endlocal

pause

