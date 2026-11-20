@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>&1
if errorlevel 1 (
  echo Python 3 launcher required.
  pause
  exit /b 1
)
choice /c 123 /m "1 Offline queue, 2 Online send, 3 Render round trip"
if errorlevel 3 (
  py -3 windows_roundtrip.py
) else if errorlevel 2 (
  py -3 windows_probe.py
) else (
  py -3 windows_probe.py --offline
)
if errorlevel 1 echo Test failed; check connection and credentials.
pause
