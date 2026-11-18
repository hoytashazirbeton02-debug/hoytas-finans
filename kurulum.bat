@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo HOYTAS Finans - Kurulum ve bagimlilik kontrolu
echo.
where py >nul 2>nul
if not errorlevel 1 (
  py -3 -m pip install -r requirements.txt
) else (
  where python >nul 2>nul
  if errorlevel 1 (
    echo Python bulunamadi. Python 3.11, 3.12 veya 3.13 kurun.
    pause
    exit /b 1
  )
  python -m pip install -r requirements.txt
)
if errorlevel 1 (
  echo Paket kurulumu basarisiz. Internet baglantisini ve Python kurulumunu kontrol edin.
  pause
  exit /b 1
)
echo.
echo Kurulum tamamlandi. Baslat.vbs veya Baslat.bat dosyasini calistirin.
pause