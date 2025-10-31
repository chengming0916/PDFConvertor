import subprocess
import sys
import os

def install_pyinstaller():
    """Install PyInstaller if not already installed"""
    try:
        import PyInstaller
    except ImportError:
        print("Installing PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "PyInstaller"])

def build_executable():
    """Build the executable using PyInstaller"""
    # Install PyInstaller if needed
    install_pyinstaller()
    
    # Define PyInstaller command
    cmd = [
        "pyinstaller",
        "--onefile",  # Create a single executable file
        "--windowed",  # No console window (for GUI apps)
        "--name=PDFConvertor",  # Name of the executable
        "--icon=NONE",  # No icon for now
        "main.py"  # Main script to package
    ]
    
    print("Building executable...")
    print(f"Running command: {' '.join(cmd)}")
    
    try:
        # Run PyInstaller
        subprocess.run(cmd, check=True)
        print("\nBuild completed successfully!")
        print("Executable location: dist/PDFConvertor.exe")
    except subprocess.CalledProcessError as e:
        print(f"\nError building executable: {e}")
        sys.exit(1)

if __name__ == "__main__":
    build_executable()
