@echo off
setlocal
cd /d "%~dp0"
if not exist "%~dp0Baslat.vbs" (
  echo HOYTAS program dosyalari bulunamadi. ZIP dosyasini tam olarak ayiklayin.
  pause
  exit /b 1
)
if not exist "%~dp0HOYTAS_SABIT_BASLATICI.vbs" (
  echo Sabit baslatici dosyasi eksik.
  pause
  exit /b 1
)
echo HOYTAS masaustu kisayolu guncelleniyor...
"%SystemRoot%\System32\wscript.exe" "%~dp0MASAUSTU_KISAYOLU_OLUSTUR.vbs"
if errorlevel 1 (
  echo Kisayol olusturulamadi. Hata ayrintisi icin ekrandaki mesaji kontrol edin.
  pause
  exit /b 1
)
endlocal