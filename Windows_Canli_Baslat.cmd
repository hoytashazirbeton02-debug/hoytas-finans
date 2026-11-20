@echo off
setlocal
cd /d "%~dp0"
py -3 hybrid_start_live.py
if errorlevel 1 pause
