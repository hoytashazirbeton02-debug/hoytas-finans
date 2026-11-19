Option Explicit
Dim sh, fso, root, pointer, folder, script, pyw, pyexe, f, ownFolder
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
root = sh.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\HOYTASFinans"
If Not fso.FolderExists(root) Then fso.CreateFolder(root)
pointer = fso.BuildPath(root, "aktif_program_yolu.txt")
folder = ""
ownFolder = fso.GetParentFolderName(WScript.ScriptFullName)
If fso.FileExists(fso.BuildPath(ownFolder, "hoytas.py")) Then
  folder = ownFolder
  Set f = fso.CreateTextFile(pointer, True, True)
  f.Write folder
  f.Close
ElseIf fso.FileExists(pointer) Then
  Set f = fso.OpenTextFile(pointer, 1, False, -1)
  folder = Trim(f.ReadAll)
  f.Close
End If
If folder = "" Or Not fso.FolderExists(folder) Then
  MsgBox "HOYTAS program klasoru bulunamadi. ZIP dosyasini tamamen ayiklayin ve Baslat.vbs dosyasina cift tiklayin.", vbExclamation, "HOYTAS Finans"
  WScript.Quit 1
End If
script = fso.BuildPath(folder, "hoytas.py")
If Not fso.FileExists(script) Then
  MsgBox "hoytas.py dosyasi bulunamadi: " & script, vbCritical, "HOYTAS Finans"
  WScript.Quit 1
End If
sh.CurrentDirectory = folder
pyw = FindExe("pyw.exe")
If pyw = "" Then
  pyexe = FindExe("py.exe")
  If pyexe <> "" Then
    pyw = fso.BuildPath(fso.GetParentFolderName(pyexe), "pyw.exe")
    If Not fso.FileExists(pyw) Then pyw = ""
  End If
End If
If pyw <> "" Then
  sh.Run """" & pyw & """ -3 """ & script & """", 0, False
  WScript.Quit 0
End If
pyw = FindExe("pythonw.exe")
If pyw = "" Then
  pyexe = FindExe("python.exe")
  If pyexe <> "" Then
    pyw = fso.BuildPath(fso.GetParentFolderName(pyexe), "pythonw.exe")
    If Not fso.FileExists(pyw) Then pyw = ""
  End If
End If
If pyw <> "" Then
  sh.Run """" & pyw & """ """ & script & """", 0, False
  WScript.Quit 0
End If
MsgBox "Python'un pencere gostermeyen baslaticisi (pyw.exe/pythonw.exe) bulunamadi." & vbCrLf & _
       "Python 3 kurulumunu onarin; sonra Baslat.vbs dosyasina yeniden cift tiklayin.", vbCritical, "HOYTAS Finans"
WScript.Quit 1
Function FindExe(exeName)
  Dim tempPath, fileObj, found
  tempPath = fso.BuildPath(fso.GetSpecialFolder(2), fso.GetTempName)
  sh.Run "cmd.exe /d /c where " & exeName & " > """ & tempPath & """ 2>nul", 0, True
  found = ""
  If fso.FileExists(tempPath) Then
    Set fileObj = fso.OpenTextFile(tempPath, 1, False, 0)
    If Not fileObj.AtEndOfStream Then found = Trim(fileObj.ReadLine)
    fileObj.Close
    fso.DeleteFile tempPath, True
  End If
  If found <> "" And fso.FileExists(found) Then
    FindExe = found
  Else
    FindExe = ""
  End If
End Function