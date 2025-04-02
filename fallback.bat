@echo off
chcp 65001 > nul
title PPT to Video Tool - Simplified Build

:: Create a bare minimum packaging process
echo [INFO] Starting simplified build process...

:: Clean previous build files
if exist dist rmdir /s /q dist
if exist build rmdir /s /q build
if exist "*.spec" del "*.spec"

:: Install bare minimum requirements
echo [INFO] Installing essential packages...
python -m pip install pyinstaller --quiet

:: Run the simplest possible PyInstaller command
echo [INFO] Running PyInstaller with minimal options...
python -m PyInstaller --clean --noconfirm --onefile --name="PPT转视频工具" app.py

:: Check result
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Build failed.
    echo [INFO] You may need to run cmd as administrator and try again.
    goto end
) else (
    echo [SUCCESS] Build completed successfully!
    echo [INFO] The executable is available in the 'dist' folder.
    
    :: Create English-only batch file to avoid encoding issues
    echo @echo off > "dist\run.bat"
    echo cd /d "%%~dp0" >> "dist\run.bat"
    echo start "" "PPT转视频工具.exe" >> "dist\run.bat"
    
    echo [INFO] Created run.bat launcher in dist folder
)

:: Copy readme if available
if exist README.txt copy README.txt dist\

:end
echo.
echo Press any key to exit...
pause > nul
