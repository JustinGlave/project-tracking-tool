@echo off
:: ============================================================
:: build.bat — builds ProjectTrackingTool
:: Run this from the project folder:  build.bat
:: Requires PyInstaller:  pip install pyinstaller
:: ============================================================

:: Read version from version.py
for /f "tokens=3 delims= " %%v in ('findstr "__version__" version.py') do set VERSION=%%~v

echo Building Project Tracking Tool v%VERSION%...

pyinstaller ^
    --onedir ^
    --windowed ^
    --icon=PTT_Normal.ico ^
    --name=ProjectTrackingTool ^
    --add-data="PTT_Transparent.png;." ^
    --add-data="PTT_Normal.ico;." ^
    --hidden-import=openpyxl ^
    --hidden-import=openpyxl.cell._writer ^
    --collect-submodules=openpyxl ^
    --collect-all=PySide6 ^
    --hidden-import=xml.etree.ElementTree ^
    --hidden-import=xml.etree.cElementTree ^
    --collect-submodules=xml ^
    project_tracker_gui.py

if errorlevel 1 (
    echo BUILD FAILED.
    pause
    exit /b 1
)

echo.
echo Build complete: dist\ProjectTrackingTool\ProjectTrackingTool.exe
echo Version: %VERSION%
echo.

:: Zip 1 - exe only for auto-updater
echo Creating update zip (exe only)...
powershell -Command "Compress-Archive -Path 'dist\ProjectTrackingTool\ProjectTrackingTool.exe' -DestinationPath 'dist\ProjectTrackingTool.zip' -Force"
echo   Created: dist\ProjectTrackingTool.zip

:: Zip 2 - full folder for fresh installs
echo Creating full install zip...
powershell -Command "Compress-Archive -Path 'dist\ProjectTrackingTool' -DestinationPath 'dist\ProjectTrackingTool_FullInstall.zip' -Force"
echo   Created: dist\ProjectTrackingTool_FullInstall.zip

echo.
echo Next steps:
echo   1. Test dist\ProjectTrackingTool\ProjectTrackingTool.exe
echo   2. Go to GitHub Releases, Draft new release
echo   3. Tag: v%VERSION%
echo   4. Upload BOTH:
echo        dist\ProjectTrackingTool.zip           (auto-updater)
echo        dist\ProjectTrackingTool_FullInstall.zip  (fresh install)
echo   5. Publish release
echo.
pause
