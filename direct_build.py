"""
A minimal script to directly build the EXE without complex spec file generation
"""
import os
import sys
import subprocess
import shutil

def main():
    print("Starting simplified packaging process...")
    
    # Clean previous builds
    for path in ["dist", "build"]:
        if os.path.exists(path):
            print(f"Cleaning {path}...")
            shutil.rmtree(path)
    
    # Remove old spec files
    for spec_file in ["PPT转视频工具.spec", "app.spec", "custom_spec.spec"]:
        if os.path.exists(spec_file):
            os.remove(spec_file)
    
    # Create resources folder
    if not os.path.exists("resources"):
        os.makedirs("resources")
    
    # Install required packages
    print("Installing required packages...")
    packages = ["pyinstaller", "pillow", "moviepy", "pyttsx3", "gtts", "websocket-client"]
    for pkg in packages:
        try:
            subprocess.run([sys.executable, "-m", "pip", "install", pkg, "--upgrade"], check=False)
        except Exception as e:
            print(f"Warning: Failed to install {pkg}: {e}")
    
    # Direct PyInstaller command with minimum complexity
    print("Running PyInstaller...")
    cmd = [
        sys.executable, 
        "-m", 
        "PyInstaller",
        "--clean",
        "--noconfirm", 
        "--onefile",
        "--windowed",
        "--name=PPT转视频工具",
        "--hidden-import=pyttsx3.drivers",
        "--hidden-import=pyttsx3.drivers.sapi5",
        "--hidden-import=moviepy.editor",
        "--hidden-import=moviepy.audio.io.readers",
        "--hidden-import=moviepy.video.io.ffmpeg_reader",
        "--hidden-import=websocket",
        "app.py"
    ]
    
    print(f"Command: {' '.join(cmd)}")
    
    try:
        subprocess.run(cmd, check=True)
        
        # Create a more robust launcher .bat file that handles Unicode paths properly
        launcher_path = os.path.join("dist", "启动程序.bat")
        with open(launcher_path, "w", encoding="utf-8") as f:
            f.write('@echo off\r\n')
            f.write('chcp 65001 >nul\r\n')  # Set console to UTF-8
            f.write('setlocal enabledelayedexpansion\r\n')
            f.write('cd /d "%~dp0"\r\n')  # Change to the batch file's directory
            f.write('echo 正在启动PPT转视频工具...\r\n')
            f.write('echo 如果程序闪退，请查看error_log.txt文件获取错误信息\r\n')
            f.write('echo.\r\n')
            f.write('echo 注意：首次运行可能会较慢，请耐心等待\r\n')
            f.write('echo.\r\n')
            f.write('if exist "PPT转视频工具.exe" (\r\n')
            f.write('  "PPT转视频工具.exe" 2> error_log.txt\r\n')
            f.write('  if !ERRORLEVEL! NEQ 0 (\r\n')
            f.write('    echo 程序运行出错，详细信息已记录到error_log.txt\r\n')
            f.write('    notepad error_log.txt\r\n')
            f.write('    pause\r\n')
            f.write('  )\r\n')
            f.write(') else (\r\n')
            f.write('  echo 错误：未找到可执行文件"PPT转视频工具.exe"\r\n')
            f.write('  echo 请确保该文件与启动程序.bat在同一目录\r\n')
            f.write('  pause\r\n')
            f.write(')\r\n')
        
        # Also create a simple runner without Chinese characters
        simple_launcher_path = os.path.join("dist", "run.bat")
        with open(simple_launcher_path, "w", encoding="utf-8") as f:
            f.write('@echo off\r\n')
            f.write('cd /d "%~dp0"\r\n')
            f.write('start "" "PPT转视频工具.exe"\r\n')

        # Copy resource files
        for file in ["XFYUN_TROUBLESHOOTING.md", "README.txt", "README.md"]:
            if os.path.exists(file):
                if os.path.exists(os.path.join("dist", file)):
                    os.remove(os.path.join("dist", file))
                shutil.copy(file, "dist")
        
        print("\nBuild completed successfully!")
        print("The executable is in the 'dist' folder.")
        print("Use '启动程序.bat' for better error handling.")
        
        return 0
        
    except subprocess.CalledProcessError as e:
        print(f"Error running PyInstaller: {e}")
        return 1

if __name__ == "__main__":
    sys.exit(main())
