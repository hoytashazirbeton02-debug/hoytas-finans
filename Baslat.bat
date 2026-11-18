@echo off
setlocal
cd /d "%~dp0"
if not exist "%~dp0Baslat.vbs" (
  echo HOYTAS program dosyalari bulunamadi. ZIP dosyasini tam olarak ayiklayin.
  exit /b 1
)
start "" "%SystemRoot%\System32\wscript.exe" "%~dp0Baslat.vbs"
exit /b 0