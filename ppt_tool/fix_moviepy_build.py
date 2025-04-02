"""
Ultra-simplified build script with explicit focus on moviepy inclusion
"""
import os
import sys
import subprocess
import shutil
import site
import importlib.util

def main():
    print("Starting specialized packaging process for moviepy...")
    
    # Clean previous builds
    for path in ["dist", "build"]:
        if os.path.exists(path):
            print(f"Cleaning {path}...")
            shutil.rmtree(path)
    
    # Remove old spec files
    for spec_file in ["PPT转视频工具.spec", "app.spec", "custom_spec.spec"]:
        if os.path.exists(spec_file):
            os.remove(spec_file)
    
    # Install required packages with exact versions that we know work
    print("Installing specific package versions...")
    packages = [
        "pyinstaller==6.5.0",
        "moviepy==1.0.3",  # Specify exact version
        "numpy==1.24.3",   # Compatible version for moviepy
        "pillow==9.5.0",   # Compatible version
        "pyttsx3==2.90",
        "websocket-client==1.7.0"
    ]  # 移除了gTTS
    
    for pkg in packages:
        try:
            print(f"Installing {pkg}...")
            subprocess.run([sys.executable, "-m", "pip", "install", pkg, "--force-reinstall"], check=True)
        except Exception as e:
            print(f"Warning: Failed to install {pkg}: {e}")
    
    # Create a helper script to ensure moviepy is properly imported
    with open("import_helper.py", "w") as f:
        f.write("# Helper script to ensure PyInstaller includes all necessary modules\n")
        f.write("import moviepy\n")
        f.write("import moviepy.editor\n")
        f.write("from moviepy.editor import *\n")
        f.write("import moviepy.video.io.ffmpeg_reader\n")
        f.write("import moviepy.audio.io.readers\n")
        f.write("import moviepy.video.fx.all\n")
        f.write("import moviepy.audio.fx.all\n")
        f.write("import numpy\n")
        f.write("import PIL\n")
        f.write("from PIL import Image, ImageDraw, ImageFont\n")
        f.write("import pyttsx3\n")
        f.write("import pyttsx3.drivers\n")
        f.write("import pyttsx3.drivers.sapi5\n")
        f.write("import win32com.client\n")
        f.write("import websocket\n")
        f.write("import requests\n")
        f.write("print('All required modules imported successfully!')\n")
    
    # Test if our module imports work
    print("Testing module imports...")
    result = subprocess.run([sys.executable, "import_helper.py"], capture_output=True, text=True)
    if "All required modules imported successfully!" in result.stdout:
        print("Module import test passed!")
    else:
        print("WARNING: Module import test failed!")
        print(f"STDOUT: {result.stdout}")
        print(f"STDERR: {result.stderr}")
    
    # Find moviepy installation paths
    moviepy_spec = importlib.util.find_spec("moviepy")
    if moviepy_spec:
        moviepy_path = os.path.dirname(moviepy_spec.origin)
        print(f"Found moviepy at: {moviepy_path}")
    else:
        print("WARNING: Could not find moviepy installation path")
        moviepy_path = None
    
    # Create a .spec file that explicitly includes moviepy
    print("Creating custom .spec file...")
    spec_content = f"""
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# Explicitly gather moviepy modules
added_files = []
moviepy_modules = [
    'moviepy', 'moviepy.editor', 'moviepy.config', 
    'moviepy.tools', 'moviepy.audio', 'moviepy.video',
    'moviepy.audio.fx', 'moviepy.audio.fx.all', 'moviepy.audio.io',
    'moviepy.video.fx', 'moviepy.video.fx.all', 'moviepy.video.io',
    'moviepy.audio.io.readers', 'moviepy.video.io.ffmpeg_reader',
    'moviepy.video.io.html_tools', 'moviepy.video.io.ffmpeg_writer'
]

a = Analysis(
    ['app.py', 'import_helper.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=moviepy_modules + [
        'PIL', 'PIL._imagingft', 'PIL.ImageFont', 'PIL.ImageDraw',
        'numpy', 'scipy', 'scipy.io', 'scipy.signal',
        'pyttsx3', 'pyttsx3.drivers', 'pyttsx3.drivers.sapi5',
        'win32com', 'win32com.client',
        'websocket', 'ssl', 'wave', 'hmac',
        'hashlib', 'urllib', 'urllib.parse', 'base64', 'datetime'
    ],  # 移除了gtts
    hookspath=[],
    hooksconfig={{}},
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
    console=True,  # Set to True for debugging
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
"""
    
    with open("moviepy_app.spec", "w", encoding="utf-8") as f:
        f.write(spec_content)
    
    # Run PyInstaller with the custom spec file
    print("Running PyInstaller with custom spec file...")
    try:
        subprocess.run([sys.executable, "-m", "PyInstaller", "moviepy_app.spec"], check=True)
        
        # Create a simple batch launcher
        with open(os.path.join("dist", "run.bat"), "w", encoding="utf-8") as f:
            f.write("@echo off\n")
            f.write("cd /d \"%~dp0\"\n")
            f.write("start \"\" \"PPT转视频工具.exe\"\n")
        
        # Copy documentation files
        for file in ["README.txt"]:
            if os.path.exists(file):
                shutil.copy(file, "dist")
        
        print("\nBuild completed successfully!")
        print("The executable is in the dist folder.")
        print("Use run.bat to start the application.")
        
        return 0
    except subprocess.CalledProcessError as e:
        print(f"Error running PyInstaller: {e}")
        
        # Fallback to direct command line if spec file fails
        print("\nTrying direct command line approach...")
        try:
            # Prepare a list of all hidden imports
            hidden_imports = []
            for module in [
                "moviepy", "moviepy.editor", "moviepy.video.io.ffmpeg_reader",
                "moviepy.audio.io.readers", "PIL", "PIL.ImageFont", "PIL.ImageDraw",
                "numpy", "pyttsx3", "pyttsx3.drivers.sapi5", "win32com.client",
                "websocket"
            ]:
                hidden_imports.extend(["--hidden-import", module])
            
            cmd = [
                sys.executable, "-m", "PyInstaller",
                "--clean", "--noconfirm", "--onefile",
                "--name=PPT转视频工具"
            ] + hidden_imports + ["app.py"]
            
            print(f"Running command: {' '.join(cmd)}")
            subprocess.run(cmd, check=True)
            
            # Create simple launcher
            with open(os.path.join("dist", "run.bat"), "w", encoding="utf-8") as f:
                f.write("@echo off\n")
                f.write("cd /d \"%~dp0\"\n")
                f.write("start \"\" \"PPT转视频工具.exe\"\n")
            
            print("\nFallback build completed!")
            return 0
        except subprocess.CalledProcessError as e2:
            print(f"Fallback approach also failed: {e2}")
            return 1

if __name__ == "__main__":
    sys.exit(main())
