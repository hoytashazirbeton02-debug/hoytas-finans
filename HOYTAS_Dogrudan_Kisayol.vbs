Option Explicit
Dim shell, files, folder, desktop, name, candidate, target, iconPath, file, oldLink, command
Set shell = CreateObject("WScript.Shell")
Set files = CreateObject("Scripting.FileSystemObject")
folder = files.GetParentFolderName(WScript.ScriptFullName)
desktop = shell.SpecialFolders("Desktop")
target = WScript.Arguments(0)
iconPath = files.BuildPath(folder, "HOYTAS.ico")
name = "HOYTA" & ChrW(350)
If Not files.FileExists(target) Or Not files.FileExists(iconPath) Then WScript.Quit 1

For Each file In files.GetFolder(desktop).Files
    If LCase(files.GetExtensionName(file.Name)) = "lnk" Then
        If InStr(1, file.Name, name, vbTextCompare) > 0 Or InStr(1, file.Name, "HOYTAS", vbTextCompare) > 0 Then
            Set oldLink = shell.CreateShortcut(file.Path)
            command = LCase(oldLink.TargetPath & " " & oldLink.Arguments)
            If InStr(command, "windows_canli_baslat.cmd") > 0 Or InStr(command, "hoytas_sabit_baslatici.vbs") > 0 _
               Or InStr(command, "hoytas.py") > 0 Or InStr(command, "baslat.vbs") > 0 Or InStr(command, "baslat.bat") > 0 Then
                SaveLink file.Path
            End If
        End If
    End If
Next

For Each candidate In Array(name & ".lnk", "HOYTAS.lnk", name & " v92.lnk", name & " Canl" & ChrW(305) & ".lnk")
    If candidate = name & ".lnk" Or files.FileExists(files.BuildPath(desktop, candidate)) Then
        SaveLink files.BuildPath(desktop, candidate)
    End If
Next

Sub SaveLink(ByVal path)
    Dim link, verified, args
    args = Chr(34) & files.BuildPath(folder, "HOYTAS.pyw") & Chr(34)
    Set link = shell.CreateShortcut(path)
    link.TargetPath = target
    link.Arguments = args
    link.IconLocation = iconPath & ",0"
    link.WorkingDirectory = folder
    link.Description = "HOYTAS Finans ve Isletme Yonetimi"
    link.Save
    Set verified = shell.CreateShortcut(path)
    If StrComp(verified.TargetPath, target, vbTextCompare) <> 0 Or verified.Arguments <> args Then WScript.Quit 2
End Sub
