"""
Simple packaging script that ensures all dependencies are properly included
"""
import os
import sys
import subprocess
import shutil
import importlib
import time

def check_dependency_importable(module_name):
    """Check if a module can be imported"""
    try:
        importlib.import_module(module_name)
        return True
    except ImportError:
        return False

def main():
    print("Starting packaging process...")
    
    # Create resources directory if it doesn't exist
    if not os.path.exists("resources"):
        os.makedirs("resources")
    
    # Clean previous build
    for path in ["dist", "build"]:
        if os.path.exists(path):
            print(f"Cleaning {path}...")
            shutil.rmtree(path)
    
    for spec_file in ["PPT转视频工具.spec", "custom_spec.spec"]:
        if os.path.exists(spec_file):
            os.remove(spec_file)
    
    # Install all required packages explicitly
    print("Installing required packages...")
    required_packages = [
        "pyinstaller",
        "pillow",
        "moviepy",
        "pyttsx3",
        "gTTS",
        "pywin32",
        "numpy", 
        "requests",
        "websocket-client",
        "ntplib",
        "dnspython",
        "cffi"
    ]
    
    for package in required_packages:
        print(f"Installing {package}...")
        try:
            subprocess.check_call([
                sys.executable, "-m", "pip", "install", package, "--upgrade"
            ])
        except subprocess.CalledProcessError as e:
            print(f"Warning: Failed to install {package}: {e}")
    
    # Verify critical dependencies
    critical_deps = ["moviepy", "pyttsx3", "PIL", "numpy"]
    missing_deps = []
    for dep in critical_deps:
        if not check_dependency_importable(dep):
            missing_deps.append(dep)
    
    if missing_deps:
        print(f"ERROR: The following critical dependencies could not be imported: {', '.join(missing_deps)}")
        print("Please install them manually before continuing.")
        return 1
    
    # Create a comprehensive spec file
    print("Creating PyInstaller spec file...")
    spec_content = """
# -*- mode: python ; coding: utf-8 -*-

import sys
from PyInstaller.utils.hooks import collect_all

block_cipher = None

# Collect all for key packages (this helps with hidden imports and data files)
moviepy_a = collect_all('moviepy')
pyttsx3_a = collect_all('pyttsx3')
gtts_a = collect_all('gtts')
pillow_a = collect_all('PIL')
numpy_a = collect_all('numpy')
websocket_a = collect_all('websocket')

# Combine all collected data
hiddenimports = []
data = []
binaries = []

for a in [moviepy_a, pyttsx3_a, gtts_a, pillow_a, numpy_a, websocket_a]:
    hiddenimports.extend(a[0])
    data.extend(a[1])
    binaries.extend(a[2])

# Additional hidden imports often missed
hiddenimports.extend([
    'moviepy.audio.fx', 
    'moviepy.audio.fx.all',
    'moviepy.audio.io',
    'moviepy.video.fx', 
    'moviepy.video.fx.all',
    'moviepy.video.io',
    'moviepy.audio.io.readers',
    'moviepy.video.io.readers',
    'scipy',
    'scipy.io',
    'scipy.io.wavfile',
    'scipy.fftpack',
    'proglog',
    'tqdm',
    'pydub',
    'pyttsx3.drivers',
    'pyttsx3.drivers.sapi5',
    'win32com.client',
    'win32com',
    'websocket',
    'ssl',
    'wave',
    'hmac',
    'hashlib',
    'urllib',
    'urllib.parse',
    'base64',
    'datetime',
    'ntplib',
    'dns',
    'dns.resolver',
])

# Main analysis
a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=binaries,
    datas=data,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='PPT转视频工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # Set to True for debugging, change to False for final release
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
"""
    
    with open("custom_spec.spec", "w", encoding="utf-8") as f:
        f.write(spec_content)
    
    # Run PyInstaller with the custom spec file
    print("Running PyInstaller with custom spec file...")
    try:
        subprocess.check_call([
            sys.executable, "-m", "PyInstaller",
            "custom_spec.spec"
        ])
        
        # Give a little time for files to be fully written
        time.sleep(2)
        
        # Copy resource files
        print("Copying resource files...")
        for file in ["XFYUN_TROUBLESHOOTING.md", "README.txt"]:
            if os.path.exists(file):
                if os.path.exists(os.path.join("dist", file)):
                    os.remove(os.path.join("dist", file))
                shutil.copy(file, "dist")
        
        print("\nPackaging completed successfully!")
        print("The executable is available in the 'dist' folder.")
        
    except subprocess.CalledProcessError as e:
        print(f"\nError: PyInstaller failed: {e}")
        print("\nTrying alternative approach...")
        
        try:
            # Try a simpler approach with direct command line arguments
            subprocess.check_call([
                sys.executable, "-m", "PyInstaller",
                "--clean",
                "--noconfirm",
                "--windowed",
                "--onefile",
                "--add-data", "README.txt;.",
                "--hidden-import", "moviepy",
                "--hidden-import", "moviepy.editor",
                "--hidden-import", "PIL",
                "--hidden-import", "PIL._imagingft",
                "--hidden-import", "numpy",
                "--hidden-import", "scipy",
                "--hidden-import", "pyttsx3",
                "--hidden-import", "gtts",
                "--hidden-import", "websocket",
                "--hidden-import", "win32com",
                "--hidden-import", "win32com.client",
                "app.py"
            ])
            
            print("\nAlternative packaging method succeeded!")
            print("The executable is available in the 'dist' folder.")
            
        except subprocess.CalledProcessError as e2:
            print(f"\nAlternative approach also failed: {e2}")
            print("Check if all dependencies are installed and try again.")
            return 1
    
    # Create a batch file to run the app with error capture
    launcher_path = os.path.join("dist", "启动程序.bat")
    with open(launcher_path, "w", encoding="utf-8") as f:
        f.write('@echo off\n')
        f.write('echo 正在启动PPT转视频工具...\n')
        f.write('echo 如果程序闪退，请查看error_log.txt文件获取错误信息\n')
        f.write('echo.\n')
        f.write('"PPT转视频工具.exe" 2> error_log.txt\n')
        f.write('if %ERRORLEVEL% NEQ 0 (\n')
        f.write('  echo 程序运行出错，详细信息已记录到error_log.txt\n')
        f.write('  notepad error_log.txt\n')
        f.write('  pause\n')
        f.write(')\n')
    
    print(f"Created launcher batch file: {launcher_path}")
    print("If the main EXE crashes, use this batch file to capture error details.")
    
    return 0

if __name__ == "__main__":
    sys.exit(main())
