@echo off
python build_exe.py
if errorlevel 1 (
    py build_exe.py
)
pause
