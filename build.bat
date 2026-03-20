@echo off
:: ============================================================
:: build.bat — builds ProjectTrackingTool.exe
:: Run this from the project folder:  build.bat
:: Requires PyInstaller:  pip install pyinstaller
:: ============================================================

:: Read version from version.py
for /f "tokens=3 delims= " %%v in ('findstr "__version__" version.py') do set VERSION=%%~v

echo Building Project Tracking Tool v%VERSION%...

pyinstaller ^
    --onefile ^
    --windowed ^
    --icon=PTT_Normal.ico ^
    --name=ProjectTrackingTool ^
    --add-data="PTT_Transparent.png;." ^
    --add-data="PTT_Normal.ico;." ^
    project_tracker_gui.py

if errorlevel 1 (
    echo BUILD FAILED.
    pause
    exit /b 1
)

echo.
echo Build complete: dist\ProjectTrackingTool.exe
echo Version: %VERSION%
echo.
echo Next steps:
echo   1. Test dist\ProjectTrackingTool.exe
echo   2. Go to GitHub ^> Releases ^> Draft new release
echo   3. Tag: v%VERSION%
echo   4. Upload: dist\ProjectTrackingTool.exe
echo   5. Publish release
echo.
pause
