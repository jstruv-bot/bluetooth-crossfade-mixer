"""
Build script for packaging BluetoothCrossfadeMixer as a standalone .exe
using PyInstaller.

Usage:
    python build.py

The resulting executable will be in the dist/ folder.
"""

import subprocess
import sys
import os

# Find the PyAudioWPatch portaudio DLL path
pyaudio_dll = None
try:
    import pyaudiowpatch
    pyaudio_dir = os.path.dirname(pyaudiowpatch.__file__)
    for f in os.listdir(pyaudio_dir):
        if f.lower().endswith(".dll") and "portaudio" in f.lower():
            pyaudio_dll = os.path.join(pyaudio_dir, f)
            break
except ImportError:
    pass

cmd = [
    sys.executable, "-m", "PyInstaller",
    "--onefile",
    "--console",
    "--name", "BluetoothCrossfadeMixer",
    "--add-data", "templates;templates",
]

if pyaudio_dll:
    cmd.extend(["--add-binary", f"{pyaudio_dll};."])
    print(f"Including PortAudio DLL: {pyaudio_dll}")

cmd.append("server.py")

subprocess.run(cmd, check=True)
