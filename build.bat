@echo off
chcp 65001 > nul
title PPT to Video Tool - Build

:: Create a simplified one-step packaging process
echo [INFO] Starting build process...

:: Create resources folder if it doesn't exist
if not exist resources mkdir resources

:: Clean previous build files
if exist dist rmdir /s /q dist
if exist build rmdir /s /q build
if exist "PPT转视频工具.spec" del "PPT转视频工具.spec"
if exist "custom_spec.spec" del "custom_spec.spec"

:: Install required packages silently
echo [INFO] Installing required packages...
python -m pip install pyinstaller pillow moviepy pyttsx3 websocket-client requests ntplib --quiet

:: Use the Python packager script instead of direct PyInstaller call
echo [INFO] Running packaging script...
python simple_packager.py

:: Check result
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Build failed using simple_packager.py.
    echo [INFO] Trying fallback method with direct PyInstaller call...
    
    python -m PyInstaller --clean --noconfirm --windowed --onefile ^
      --hidden-import=moviepy --hidden-import=moviepy.editor ^
      --hidden-import=PIL --hidden-import=numpy --hidden-import=scipy ^
      --hidden-import=pyttsx3 --hidden-import=websocket ^
      --name="PPT转视频工具" app.py
    
    if %ERRORLEVEL% NEQ 0 (
        echo [ERROR] All build methods failed.
        goto end
    )
)

:: Copy resources
echo [INFO] Copying resource files...
if exist XFYUN_TROUBLESHOOTING.md copy XFYUN_TROUBLESHOOTING.md dist\
if exist README.txt copy README.txt dist\

:: Cleanup
if exist simple_build.py del simple_build.py

echo.
echo [SUCCESS] Build completed successfully!
echo [INFO] The executable is available in the 'dist' folder.

:end
echo.
echo Press any key to exit...
pause > nul
