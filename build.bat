@echo off
setlocal

rem ============================================================
rem build.bat - builds ProjectTrackingTool and installer
rem Run this from the project folder: build.bat
rem Requires:
rem   pip install pyinstaller
rem   Inno Setup 6 (https://jrsoftware.org/isinfo.php)
rem ============================================================

rem Read version from version.py
for /f "tokens=3 delims= " %%v in ('findstr "__version__" version.py') do set "VERSION=%%~v"

echo ============================================================
echo  Building Project Tracking Tool v%VERSION%
echo ============================================================
echo.

rem Step 1: PyInstaller
echo [1/3] Running PyInstaller...
.venv\Scripts\pyinstaller ^
    --noconfirm ^
    --onedir ^
    --windowed ^
    --icon=PTT_Normal.ico ^
    --name=ProjectTrackingTool ^
    --add-data="PTT_Transparent.png;." ^
    --add-data="PTT_Normal.ico;." ^
    --add-data="phoenix_style.qss;." ^
    --add-data="pyxlsb;pyxlsb" ^
    --hidden-import=openpyxl ^
    --hidden-import=openpyxl.cell._writer ^
    --collect-submodules=openpyxl ^
    --collect-submodules=PySide6.QtCore ^
    --collect-submodules=PySide6.QtGui ^
    --collect-submodules=PySide6.QtWidgets ^
    --hidden-import=pyxlsb ^
    project_tracker_gui.py

if errorlevel 1 (
    echo.
    echo ERROR: PyInstaller build failed.
    exit /b 1
)
echo [1/3] PyInstaller complete.
echo.

rem Step 2: Inno Setup installer
echo [2/3] Building installer with Inno Setup...

set "ISCC="
if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" set "ISCC=C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if exist "C:\Program Files\Inno Setup 6\ISCC.exe" set "ISCC=C:\Program Files\Inno Setup 6\ISCC.exe"
if exist "%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe" set "ISCC=%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"
if exist "C:\Users\justing\AppData\Local\Programs\Inno Setup 6\ISCC.exe" set "ISCC=C:\Users\justing\AppData\Local\Programs\Inno Setup 6\ISCC.exe"

if not defined ISCC (
    echo.
    echo WARNING: Inno Setup 6 not found. Skipping installer creation.
    echo          Download from: https://jrsoftware.org/isinfo.php
    echo          Then re-run build.bat.
    echo.
    goto zips
)

"%ISCC%" /DMyAppVersion=%VERSION% installer.iss
if errorlevel 1 (
    echo.
    echo ERROR: Inno Setup build failed.
    exit /b 1
)
echo [2/3] Installer created: dist\ProjectTrackingToolSetup.exe
echo.

rem Step 3: Create zips
:zips
echo [3/3] Creating zip archives...

powershell -NoProfile -ExecutionPolicy Bypass -Command "Compress-Archive -Path 'dist\ProjectTrackingTool\ProjectTrackingTool.exe' -DestinationPath 'dist\ProjectTrackingTool.zip' -Force"
if errorlevel 1 (
    echo.
    echo ERROR: Auto-updater zip creation failed.
    exit /b 1
)
echo   Created: dist\ProjectTrackingTool.zip  (auto-updater)

powershell -NoProfile -ExecutionPolicy Bypass -Command "Compress-Archive -Path 'dist\ProjectTrackingTool' -DestinationPath 'dist\ProjectTrackingTool_FullInstall.zip' -Force"
if errorlevel 1 (
    echo.
    echo ERROR: Full install zip creation failed.
    exit /b 1
)
echo   Created: dist\ProjectTrackingTool_FullInstall.zip  (manual install)

echo.
echo ============================================================
echo  Build complete - v%VERSION%
echo ============================================================
echo.
echo  dist\ProjectTrackingTool\ProjectTrackingTool.exe   ^<-- test this first
echo  dist\ProjectTrackingToolSetup.exe                  ^<-- installer
echo  dist\ProjectTrackingTool.zip                       ^<-- auto-updater zip
echo  dist\ProjectTrackingTool_FullInstall.zip           ^<-- manual install zip
echo.
echo  Upload to GitHub Release:
echo    - ProjectTrackingTool.zip          (required for auto-updater)
echo    - ProjectTrackingToolSetup.exe     (recommended for new users)
echo.

endlocal
