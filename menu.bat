@echo off
chcp 65001 >nul
cd /d "%~dp0"
python _scripts\menu_gui.py
pause
