@echo off
setlocal

rem ============================================================
rem build.bat - builds ProjectTrackingTool and installer
rem Run this from the project folder: build.bat
rem Requires:
rem   pip install pyinstaller
rem   Inno Setup 6 (https://jrsoftware.org/isinfo.php)
rem ============================================================

rem Python 3.12 soft-warn (FROZEN_BUILD_BASELINE / ADR-014)
for /f "tokens=2" %%P in ('.venv\Scripts\python --version 2^>^&1') do set PYTHON_VERSION=%%P
echo Detected venv Python: %PYTHON_VERSION%
echo %PYTHON_VERSION% | findstr /b "3.12." >nul
if errorlevel 1 (
    echo WARNING: Canonical frozen-build venv is Python 3.12 per ADR-014 / FROZEN_BUILD_BASELINE.
    echo          Current interpreter is %PYTHON_VERSION%. Build will proceed but the
    echo          S1-safe bootloader profile is only verified on 3.12.
)

rem Commons preflight
.venv\Scripts\python -c "import phoenix_commons" 2>nul
if errorlevel 1 (
    echo ERROR: phoenix_commons not importable in this venv.
    echo        Run: git submodule update --init ^&^& pip install -e ./commons
    exit /b 1
)

rem Read version from version.py
for /f "tokens=3 delims= " %%v in ('findstr "__version__" version.py') do set "VERSION=%%~v"

if not defined VERSION (
    echo ERROR: Could not read version from version.py.
    exit /b 1
)

echo ============================================================
echo  Building Project Tracking Tool v%VERSION%
echo ============================================================
echo.

rem Step 0: sanity checks + full cleanup
echo [0/4] Running sanity checks + full cleanup...
if exist dist  rmdir /s /q dist
if exist build rmdir /s /q build
findstr /C:"Current Version: v%VERSION%" README.md >nul
if errorlevel 1 (
    echo.
    echo ERROR: README.md Current Version does not match version.py v%VERSION%.
    exit /b 1
)

.venv\Scripts\python -m py_compile version.py updater.py project_tracker_backend.py project_tracker_gui.py
if errorlevel 1 (
    echo.
    echo ERROR: Python compile check failed.
    exit /b 1
)

.venv\Scripts\python -m unittest discover -s tests
if errorlevel 1 (
    echo.
    echo ERROR: Regression tests failed.
    exit /b 1
)
echo [0/4] Sanity checks passed.
echo.

rem Step 1: PyInstaller
echo [1/4] Running PyInstaller...
.venv\Scripts\pyinstaller ^
    --noconfirm ^
    --onedir ^
    --windowed ^
    --noupx ^
    --icon=PTT_Normal.ico ^
    --name=ProjectTrackingTool ^
    --add-data="PTT_Transparent.png;." ^
    --add-data="PTT_Normal.ico;." ^
    --add-data="phoenix_style.qss;." ^
    --add-data="pyxlsb;pyxlsb" ^
    --hidden-import=openpyxl ^
    --hidden-import=openpyxl.cell._writer ^
    --collect-submodules=openpyxl ^
    --collect-all=phoenix_commons ^
    --collect-submodules=PySide6.QtCore ^
    --collect-submodules=PySide6.QtGui ^
    --collect-submodules=PySide6.QtWidgets ^
    --hidden-import=pyxlsb ^
    --exclude-module=tkinter ^
    --exclude-module=_tkinter ^
    --exclude-module=tcl ^
    --exclude-module=tk ^
    --exclude-module=lib2to3 ^
    --exclude-module=idlelib ^
    --exclude-module=turtle ^
    --exclude-module=turtledemo ^
    project_tracker_gui.py

if errorlevel 1 (
    echo.
    echo ERROR: PyInstaller build failed.
    exit /b 1
)
if not exist "dist\ProjectTrackingTool\ProjectTrackingTool.exe" (
    echo.
    echo ERROR: PyInstaller output missing ProjectTrackingTool.exe.
    exit /b 1
)
if not exist "dist\ProjectTrackingTool\_internal" (
    echo.
    echo ERROR: PyInstaller output missing _internal runtime folder.
    exit /b 1
)
echo [1/4] PyInstaller complete.
echo.

rem Step 2: Inno Setup installer
echo [2/4] Building installer with Inno Setup...

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
if not exist "dist\ProjectTrackingToolSetup.exe" (
    echo.
    echo ERROR: Installer output missing dist\ProjectTrackingToolSetup.exe.
    exit /b 1
)
echo [2/4] Installer created: dist\ProjectTrackingToolSetup.exe
echo.

rem Step 3: Create zips
:zips
echo [3/4] Creating zip archives...

powershell -NoProfile -ExecutionPolicy Bypass -Command "Compress-Archive -Path 'dist\ProjectTrackingTool\*' -DestinationPath 'dist\ProjectTrackingTool.zip' -Force"
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
echo [4/4] Verifying release artifacts...
if not exist "dist\ProjectTrackingTool.zip" (
    echo ERROR: Missing dist\ProjectTrackingTool.zip.
    exit /b 1
)
if not exist "dist\ProjectTrackingTool_FullInstall.zip" (
    echo ERROR: Missing dist\ProjectTrackingTool_FullInstall.zip.
    exit /b 1
)
powershell -NoProfile -ExecutionPolicy Bypass -Command "$z='dist\ProjectTrackingTool.zip'; Add-Type -AssemblyName System.IO.Compression.FileSystem; $zip=[System.IO.Compression.ZipFile]::OpenRead($z); try { $names=$zip.Entries.FullName | ForEach-Object { $_ -replace '\\','/' }; if ($names -notcontains 'ProjectTrackingTool.exe') { exit 2 }; if (-not ($names | Where-Object { $_ -like '_internal/*' })) { exit 3 } } finally { $zip.Dispose() }"
if errorlevel 1 (
    echo.
    echo ERROR: Auto-updater zip must contain ProjectTrackingTool.exe and _internal\*.
    exit /b 1
)
echo [4/4] Artifact verification passed.

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
