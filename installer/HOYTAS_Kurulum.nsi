Unicode True
Name "HOYTAŞ Finans"
OutFile "${OUTPUT_EXE}"
RequestExecutionLevel user
InstallDir "$LOCALAPPDATA\HOYTASFinans\uygulama\v96"
ShowInstDetails show
AutoCloseWindow true
SetCompressor /SOLID lzma
Icon "${SOURCE_DIR}/HOYTAS.ico"
Page instfiles

Var PythonDir
Var Code

Section "HOYTAŞ Finans v96"
    SetShellVarContext current
    SetRegView 64
    InitPluginsDir

    ReadRegStr $PythonDir HKCU "Software\Python\PythonCore\3.14\InstallPath" ""
    IfFileExists "$PythonDir\pythonw.exe" RuntimeFound
    ReadRegStr $PythonDir HKLM "Software\Python\PythonCore\3.14\InstallPath" ""
    IfFileExists "$PythonDir\pythonw.exe" RuntimeFound
    StrCpy $PythonDir "$LOCALAPPDATA\Programs\Python\Python314"
    IfFileExists "$PythonDir\pythonw.exe" RuntimeFound

    DetailPrint "Python çalışma ortamı kuruluyor..."
    SetOutPath "$PLUGINSDIR"
    File /oname=python-kurulum.exe "${PYTHON_INSTALLER}"
    ClearErrors
    ExecWait '"$PLUGINSDIR\python-kurulum.exe" /quiet InstallAllUsers=0 TargetDir="$PythonDir" Include_launcher=0 Include_pip=1 Include_tcltk=1 Include_test=0 Include_doc=0 Shortcuts=0 AssociateFiles=0 PrependPath=0' $Code
    IfErrors RuntimeFailed
    StrCmp $Code 0 RuntimeFound RuntimeFailed

RuntimeFound:
    IfFileExists "$PythonDir\pythonw.exe" 0 RuntimeFailed
    DetailPrint "Excel ve PDF bileşenleri kuruluyor..."
    SetOutPath "$PLUGINSDIR\wheels"
    File "${WHEEL_DIR}/*.whl"
    ClearErrors
    ExecWait '"$PythonDir\python.exe" -I -m pip --disable-pip-version-check install --no-index --no-user --upgrade --find-links "$PLUGINSDIR\wheels" --target "$LOCALAPPDATA\HOYTASFinans\Python314-Kitaplik" openpyxl reportlab' $Code
    IfErrors PackagesFailed
    StrCmp $Code 0 0 PackagesFailed
    ClearErrors
    ExecWait '"$PythonDir\pythonw.exe" -I -c "import tkinter, sqlite3"' $Code
    IfErrors RuntimeFailed
    StrCmp $Code 0 0 RuntimeFailed

    DetailPrint "HOYTAŞ masaüstü uygulaması kuruluyor..."
    SetOutPath "$INSTDIR"
    ClearErrors
    File "${SOURCE_DIR}/hoytas.py"
    File "${SOURCE_DIR}/HOYTAS.pyw"
    File "${SOURCE_DIR}/HOYTAS.ico"
    File "${SOURCE_DIR}/CANLI_GUNCELLEME.json"
    File "${SOURCE_DIR}/hybrid_*.py"
    File "${SOURCE_DIR}/HOYTAS_Dogrudan_Kisayol.vbs"
    IfErrors FilesFailed
    ClearErrors
    CreateShortcut "$DESKTOP\HOYTAŞ.lnk" "$PythonDir\pythonw.exe" '"$INSTDIR\HOYTAS.pyw"' "$INSTDIR\HOYTAS.ico" 0
    IfErrors FilesFailed
    ClearErrors
    Exec '"$PythonDir\pythonw.exe" "$INSTDIR\HOYTAS.pyw"'
    IfErrors FilesFailed
    Goto Done

RuntimeFailed:
    MessageBox MB_ICONSTOP "HOYTAŞ için Python kurulamadı veya Tkinter bileşeni eksik. Kurulum kodu: $Code"
    Abort
PackagesFailed:
    MessageBox MB_ICONSTOP "Excel ve PDF bileşenleri kurulamadı. Kurulum kodu: $Code"
    Abort
FilesFailed:
    MessageBox MB_ICONSTOP "HOYTAŞ dosyaları veya masaüstü kısayolu hazırlanamadı."
    Abort
Done:
SectionEnd
