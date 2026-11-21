@echo off
python gui_app.py
if errorlevel 1 (
    py gui_app.py
)
pause
