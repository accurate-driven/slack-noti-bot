"""
Build script to create executable from app.py
"""
import subprocess
import sys
import os

def main():
    print("Building Windows Notification Bot executable...")
    print()
    
    # Check if PyInstaller is installed
    try:
        import PyInstaller
        print("✓ PyInstaller is installed")
    except ImportError:
        print("Installing PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("✓ PyInstaller installed")
    
    print()
    print("Building executable...")
    
    # Build command
    cmd = [
        "pyinstaller",
        "--onefile",
        "--console",
        "--name", "wppbot",
        "app.py"
    ]
    
    try:
        subprocess.check_call(cmd)
        print()
        print("✓ Build complete! Executable is in the 'dist' folder: dist\\wppbot.exe")
        print()
        print("IMPORTANT: Copy the .env file to the same folder as wppbot.exe")
    except subprocess.CalledProcessError as e:
        print(f"✗ Build failed: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()

