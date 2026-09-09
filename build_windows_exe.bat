@echo off
REM ==========================================================================
REM  Build a standalone Windows .exe for the SOR Polymorphizer GUI.
REM
REM  Version 6.4.  Keep the VERSION line below in step with __version__ in
REM  polymorphize_gui.py: it is what the closing message tells the user to
REM  check in the title bar, and a stale number there sends people hunting a
REM  build problem that does not exist.
REM
REM  Requires Python 3.10 to 3.13 on PATH.
REM
REM  Always rebuilds from the freshest source and clears every cache, so an
REM  old build can never leave the fields pre-filled.
REM ==========================================================================
setlocal
set VERSION=6.4
cd /d "%~dp0"
echo SOR Polymorphizer %VERSION% build
echo Working folder: %CD%
echo.

REM --- newest download wins: copy any *.py.txt over the matching *.py --------
REM  Browsers and mail clients rename .py to .py.txt. Every module of the
REM  toolkit is listed, so a partial download cannot leave a stale module
REM  behind while the rest updates.
for %%M in (
    polymorphize_gui
    polymorphize_cli
    polymorphize_core
    polymorphize_workbook
    polymorphize_generate
    polymorphize_merge
    polymorphize_sor
    polymorphize_showcase
    polymorphize_validate
    polymorphize_template
    polymorphize_batch
) do if exist "%%M.py.txt" copy /Y "%%M.py.txt" "%%M.py" >nul

REM --- the modules the GUI cannot run without --------------------------------
set MISSING=
for %%M in (
    polymorphize_gui
    polymorphize_workbook
    polymorphize_generate
    polymorphize_merge
    polymorphize_sor
    polymorphize_showcase
    polymorphize_validate
    polymorphize_template
    polymorphize_core
) do if not exist "%%M.py" set MISSING=%%M.py !MISSING!
if not "%MISSING%"=="" goto :missing

echo Building from these source files (check the dates are recent):
for %%F in (polymorphize_gui.py polymorphize_workbook.py polymorphize_merge.py
            polymorphize_sor.py polymorphize_showcase.py) do echo    %%~tF   %%F
echo.

echo Cleaning previous build artifacts and caches...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist __pycache__ rmdir /s /q __pycache__
if exist tests\__pycache__ rmdir /s /q tests\__pycache__
del /q PolymorphizeSwagger.spec 2>nul

echo Checking Python...
python --version
if errorlevel 1 (
    echo.
    echo Python was not found on PATH. Re-run the Python installer, tick
    echo "Add python.exe to PATH", then try again.
    goto :err
)

echo.
echo Installing dependencies...
python -m pip install --upgrade pip
python -m pip install pyinstaller ruamel.yaml openpyxl tkinterdnd2
if errorlevel 1 goto :err

echo.
echo Building PolymorphizeSwagger.exe ...
REM  Every module is imported lazily somewhere, so each one is named as a
REM  hidden import. PyInstaller cannot see an import that happens inside a
REM  function, and the merge, SOR and showcase modules are all imported that
REM  way to keep the GUI's start-up quick.
python -m PyInstaller --onefile --windowed --noconfirm --clean --name PolymorphizeSwagger ^
    --collect-all ruamel.yaml ^
    --collect-all tkinterdnd2 ^
    --hidden-import openpyxl ^
    --hidden-import polymorphize_core ^
    --hidden-import polymorphize_workbook ^
    --hidden-import polymorphize_generate ^
    --hidden-import polymorphize_merge ^
    --hidden-import polymorphize_sor ^
    --hidden-import polymorphize_showcase ^
    --hidden-import polymorphize_validate ^
    --hidden-import polymorphize_template ^
    polymorphize_gui.py
if errorlevel 1 goto :err

echo.
echo ==========================================================
echo  Success: %CD%\dist\PolymorphizeSwagger.exe
echo  The app's title bar must read:  SOR Polymorphizer %VERSION%
echo  If it does not, you launched an OLD exe. Use the one in
echo  the dist folder just created.
echo ==========================================================
pause
goto :eof

:missing
echo.
echo These toolkit modules are missing from this folder:
echo    %MISSING%
echo.
echo Extract polymorphizer_v%VERSION%.zip here, or put the .py or .py.txt
echo files next to this .bat, then run again.
dir /b *.py
pause
exit /b 1

:err
echo.
echo BUILD FAILED. See the messages above.
pause
exit /b 1
