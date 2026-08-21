@echo off
title Build Ultimate PDF Toolkit Pro
cd /d "%~dp0"

echo =================================================================
echo   Building Ultimate PDF Toolkit Pro Setup Wizard and Portable App
echo =================================================================
echo.

:: 1. Check Python
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed or not added to PATH.
    echo.
    echo If your friend does NOT have Python installed:
    echo   - They do NOT need to build from code!
    echo   - They can directly run the Setup Installer from Executable_Packages folder.
    echo.
    echo To install Python automatically, run in PowerShell/CMD:
    echo   winget install Python.Python.3.13
    echo.
    pause
    exit /b 1
)

:: 2. Dependencies
echo [1/3] Checking and installing Python dependencies...
pip install -r requirements.txt
pip install pyinstaller

:: 3. Build Core App into temporary build folder (Portable Version)
echo.
echo [2/3] Compiling core application logic (Portable Version)...
if not exist "temp_build" mkdir "temp_build"
pyinstaller --noconfirm --onefile --windowed --distpath "temp_build" --name "Ultimate_PDF_Toolkit_Pro" "Ultimate_PDF_Toolkit_Pro.py"

:: 4. Bundle into Standalone Setup Wizard Installer in the SAME directory
echo.
echo [3/3] Packaging into Setup Wizard Installer (.exe)...
pyinstaller --noconfirm --onefile --windowed --distpath "." --name "Ultimate_PDF_Toolkit_Pro_Setup" --add-data "temp_build\Ultimate_PDF_Toolkit_Pro.exe;." "setup_wizard.py"

:: 5. Copy built files to Executable_Packages directory
echo.
echo Distributing executables to Executable_Packages folders...
if not exist "..\Executable_Packages\Portable_Version" mkdir "..\Executable_Packages\Portable_Version"
if not exist "..\Executable_Packages\Installer_Version" mkdir "..\Executable_Packages\Installer_Version"
copy /Y "temp_build\Ultimate_PDF_Toolkit_Pro.exe" "..\Executable_Packages\Portable_Version\" >nul
copy /Y "Ultimate_PDF_Toolkit_Pro_Setup.exe" "..\Executable_Packages\Installer_Version\" >nul

:: 6. Clean up temporary builds, specs, and intermediate files
echo.
echo Cleaning up temporary build artifacts...
if exist "temp_build" rmdir /s /q "temp_build"
if exist "build" rmdir /s /q "build"
if exist "*.spec" del /f /q "*.spec"

echo.
echo =================================================================
echo   SUCCESS! 
echo   1. The Setup Installer is available here in Code_Installation:
echo      --^> Ultimate_PDF_Toolkit_Pro_Setup.exe
echo   2. Executables have also been placed in Executable_Packages folder:
echo      --^> Portable Workable .exe
echo      --^> Installer Package .exe
echo =================================================================
echo.
echo You can now use the executables directly without rebuilding!
echo.
pause
