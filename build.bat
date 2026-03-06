@echo off
pip install -r requirements.txt
python build.py
echo.
echo Build complete! Find the .exe in the dist/ folder.
pause
