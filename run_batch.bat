@echo off
REM ==========================================================================
REM  Batch-run the SOR Polymorphizer over many swaggers from a manifest.
REM
REM  Usage:
REM    * Drag a manifest (.csv or .xlsx) onto this file, OR
REM    * Double-click it (it then looks for "manifest.csv" in this folder).
REM
REM  Change PARALLEL below to run more jobs at once.
REM ==========================================================================
setlocal
cd /d "%~dp0"

set PARALLEL=4
set MF=%~1
if "%MF%"=="" set MF=manifest.csv

if not exist "%MF%" (
    echo Manifest not found: %MF%
    echo Drag a .csv/.xlsx manifest onto this file, or put "manifest.csv" here.
    pause
    exit /b 1
)

REM make sure the engine's dependencies are present (harmless if already installed)
python -m pip install ruamel.yaml openpyxl >nul 2>&1

echo Running batch from: %MF%
python polymorphize_batch.py --manifest "%MF%" --jobs %PARALLEL%
if errorlevel 2 goto :err

echo.
echo Done. Open batch_index.html (next to your manifest) for the roll-up.
pause
goto :eof

:err
echo.
echo Batch failed to start - see messages above.
pause
exit /b 1
