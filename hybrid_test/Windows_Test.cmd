@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>&1
if errorlevel 1 (
  echo Python 3 launcher required.
  pause
  exit /b 1
)
choice /c 12 /m "1 Offline queue, 2 Online send"
if errorlevel 2 (py -3 windows_probe.py) else (py -3 windows_probe.py --offline)
if errorlevel 1 echo Test failed; check connection and credentials.
pause
