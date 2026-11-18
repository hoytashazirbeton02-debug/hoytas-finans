@echo off
chcp 65001 >nul
setlocal
set "SOURCE=%~dp0guncelleme_ayarlari_HOYTAS.json"
set "TARGET_DIR=%LOCALAPPDATA%\HOYTASFinans"
set "TARGET=%TARGET_DIR%\guncelleme_ayarlari.json"
if not exist "%SOURCE%" (
  echo Ayar dosyasi bulunamadi: %SOURCE%
  pause
  exit /b 1
)
if not exist "%TARGET_DIR%" mkdir "%TARGET_DIR%"
if exist "%TARGET%" (
  copy /Y "%TARGET%" "%TARGET%.onceki_yedek" >nul
)
copy /Y "%SOURCE%" "%TARGET%" >nul
if errorlevel 1 (
  echo Ayar dosyasi kopyalanamadi.
  pause
  exit /b 1
)
echo HOYTAS otomatik guncelleme adresi kaydedildi.
echo Konum: %TARGET%
echo Programi kapatip yeniden acin ve "Guncellemeleri Denetle" dugmesine basin.
pause
endlocal