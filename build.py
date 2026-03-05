"""
Build script for packaging BluetoothCrossfadeMixer as a standalone .exe
using PyInstaller.

Usage:
    python build.py

The resulting executable will be in the dist/ folder.
"""

import subprocess
import sys

subprocess.run([
    sys.executable, "-m", "PyInstaller",
    "--onefile",
    "--console",
    "--name", "BluetoothCrossfadeMixer",
    "--add-data", "templates;templates",
    "server.py"
], check=True)
