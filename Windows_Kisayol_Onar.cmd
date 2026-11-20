@echo off
setlocal
cd /d "%~dp0"
py -3 hybrid_shortcut_setup.py
if errorlevel 1 pause
