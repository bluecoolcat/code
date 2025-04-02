"""
A utility script to help identify missing Python modules in a packaged application
"""
import sys
import os
import importlib
import subprocess

def test_import(module_name):
    """Test if a module can be imported"""
    try:
        importlib.import_module(module_name)
        return True
    except ImportError as e:
        return False, str(e)

def main():
    print("=== Python Module Diagnostic Tool ===")
    print(f"Python version: {sys.version}")
    print(f"Executable path: {sys.executable}")
    print("")

    # Check key modules needed by the application
    modules_to_check = [
        "moviepy", "moviepy.editor",
        "PIL", "PIL.Image", "PIL.ImageDraw", "PIL.ImageFont",
        "numpy",
        "scipy", 
        "pyttsx3",
        "gtts",
        "requests",
        "websocket",
        "win32com", "win32com.client",
        "ssl", "wave", "hmac", "hashlib", "urllib", "base64", "datetime"
    ]

    print("Testing module imports:")
    missing_modules = []
    
    for module in modules_to_check:
        result = test_import(module)
        if result is True:
            print(f"✓ {module} - OK")
        else:
            error_msg = result[1]
            print(f"✗ {module} - FAILED: {error_msg}")
            missing_modules.append(module)
    
    if missing_modules:
        print("\nMissing modules:")
        for module in missing_modules:
            print(f"- {module}")
        
        print("\nTrying to install missing modules:")
        for module in missing_modules:
            base_module = module.split('.')[0]  # Get the base module name
            print(f"Installing {base_module}...")
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", base_module])
                print(f"Successfully installed {base_module}")
            except subprocess.CalledProcessError:
                print(f"Failed to install {base_module}")
    else:
        print("\nAll required modules are available!")
    
    print("\nPress Enter to exit...")
    input()

if __name__ == "__main__":
    main()
