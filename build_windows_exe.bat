@echo off
REM ==========================================================================
REM  Build a standalone Windows .exe for the SOR Polymorphizer GUI.
REM  Requires Python 3.10-3.13 on PATH.
REM
REM  This version ALWAYS rebuilds from the freshest source and clears every
REM  cache, so an old build can never leave the fields pre-filled.
REM ==========================================================================
setlocal
cd /d "%~dp0"
echo Working folder: %CD%
echo.

REM --- newest download wins: copy any *.py.txt over the matching *.py --------
if exist "polymorphize_gui.py.txt"   copy /Y "polymorphize_gui.py.txt"   "polymorphize_gui.py"   >nul
if exist "polymorphize_core.py.txt"  copy /Y "polymorphize_core.py.txt"  "polymorphize_core.py"  >nul
if exist "polymorphize_cli.py.txt"   copy /Y "polymorphize_cli.py.txt"   "polymorphize_cli.py"   >nul
if exist "polymorphize_batch.py.txt" copy /Y "polymorphize_batch.py.txt" "polymorphize_batch.py" >nul

if not exist "polymorphize_gui.py"  goto :missing
if not exist "polymorphize_core.py" goto :missing

echo Building from these source files (check the dates are recent):
for %%F in (polymorphize_gui.py polymorphize_core.py) do echo    %%~tF   %%F
echo.

echo Cleaning previous build artifacts and caches...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist __pycache__ rmdir /s /q __pycache__
del /q PolymorphizeSwagger.spec 2>nul

echo Checking Python...
python --version
if errorlevel 1 (
    echo.
    echo Python was not found on PATH. Re-run the Python installer and tick
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
python -m PyInstaller --onefile --windowed --noconfirm --clean --name PolymorphizeSwagger ^
    --collect-all ruamel.yaml ^
    --collect-all tkinterdnd2 ^
    --hidden-import openpyxl ^
    polymorphize_gui.py
if errorlevel 1 goto :err

echo.
echo ==========================================================
echo  Success: %CD%\dist\PolymorphizeSwagger.exe
echo  The app's title bar must read:  SOR Polymorphizer  v5.1
echo  If it does not, you launched an OLD exe - use the one in
echo  the dist folder just created.
echo ==========================================================
pause
goto :eof

:missing
echo.
echo Could not find polymorphize_gui.py / polymorphize_core.py in this folder.
echo Extract polymorphize_toolkit.zip here (it has the correct names), or put
echo the .py / .py.txt files next to this .bat, then run again.
dir /b
pause
exit /b 1

:err
echo.
echo BUILD FAILED - see the messages above.
pause
exit /b 1
