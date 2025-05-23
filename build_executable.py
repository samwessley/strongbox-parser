#!/usr/bin/env python3
"""
Build script for creating StrongboxParser executables using PyInstaller.
This script provides better error handling and platform-specific configurations.
"""

import sys
import os
import subprocess
import platform

def main():
    """Build the executable using PyInstaller"""
    
    print(f"Building for platform: {platform.system()}")
    print(f"Python version: {sys.version}")
    
    # Define the base PyInstaller command
    cmd = [
        'pyinstaller',
        '--onefile',  # Create a single executable file
        '--windowed',  # No console window (GUI app)
        '--name', 'StrongboxParser',
        '--clean',  # Clean before building
        # Add all the hidden imports
        '--hidden-import', 'pandas',
        '--hidden-import', 'openpyxl',
        '--hidden-import', 'dateutil',
        '--hidden-import', 'dateutil.relativedelta',
        '--hidden-import', 'tkinter',
        '--hidden-import', 'tkinter.ttk',
        '--hidden-import', 'tkinter.scrolledtext',
        '--hidden-import', 'tkinter.filedialog',
        '--hidden-import', 'tkinter.messagebox',
        '--hidden-import', 'math',
        '--hidden-import', 'calendar',
        '--hidden-import', 'copy',
        '--hidden-import', 'sys',
        '--hidden-import', 'os',
        'strongbox_parser.py'
    ]
    
    # Add platform-specific options
    if platform.system() == 'Windows':
        cmd.extend(['--exclude-module', 'matplotlib'])
    
    print("Running PyInstaller with command:")
    print(' '.join(cmd))
    
    try:
        # Run PyInstaller
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print("Build successful!")
        print("Output:", result.stdout)
        
        # Check if the executable was created
        if platform.system() == 'Windows':
            exe_path = 'dist/StrongboxParser.exe'
        else:
            exe_path = 'dist/StrongboxParser'
            
        if os.path.exists(exe_path):
            print(f"Executable created at: {exe_path}")
            print(f"File size: {os.path.getsize(exe_path)} bytes")
        else:
            print(f"Warning: Expected executable not found at {exe_path}")
            
    except subprocess.CalledProcessError as e:
        print(f"Build failed with error code {e.returncode}")
        print("stdout:", e.stdout)
        print("stderr:", e.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Unexpected error: {e}")
        sys.exit(1)

if __name__ == '__main__':
    main() 