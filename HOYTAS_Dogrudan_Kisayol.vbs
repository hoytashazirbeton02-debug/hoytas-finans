Option Explicit
Dim shell, files, folder, desktop, name, target, iconPath, file, selected, canonical
Set shell = CreateObject("WScript.Shell")
Set files = CreateObject("Scripting.FileSystemObject")
folder = files.GetParentFolderName(WScript.ScriptFullName)
desktop = shell.SpecialFolders("Desktop")
target = WScript.Arguments(0)
iconPath = files.BuildPath(folder, "HOYTAS.ico")
name = "HOYTA" & ChrW(350)
canonical = files.BuildPath(desktop, name & ".lnk")

If Not files.FileExists(target) Or Not files.FileExists(iconPath) Then WScript.Quit 1
If Not files.FileExists(files.BuildPath(folder, "HOYTAS.pyw")) Then WScript.Quit 1

selected = ""
For Each file In files.GetFolder(desktop).Files
    If LCase(files.GetExtensionName(file.Name)) = "lnk" Then
        If IsHoytasLink(file.Path) Then
            If selected = "" Then selected = file.Path
            If StrComp(file.Path, canonical, vbTextCompare) = 0 Then
                selected = file.Path
                Exit For
            End If
        End If
    End If
Next

If selected = "" Then
    If WScript.Arguments.Count > 1 Then
        If WScript.Arguments(1) = "existing-only" Then WScript.Quit 0
    End If
    If files.FileExists(canonical) Then WScript.Quit 3
    selected = canonical
End If
SaveLink selected

Function IsHoytasLink(path)
    Dim link, label, command
    Set link = shell.CreateShortcut(path)
    label = files.GetBaseName(path)
    command = LCase(link.TargetPath & " " & link.Arguments)
    IsHoytasLink = False

    If link.Description = "HOYTAS Finans ve Isletme Yonetimi" Then
        IsHoytasLink = True
        Exit Function
    End If
    If InStr(1, label, name, vbTextCompare) = 0 And InStr(1, label, "HOYTAS", vbTextCompare) = 0 Then Exit Function

    IsHoytasLink = InStr(command, "hoytas.py") > 0 Or _
        InStr(command, "windows_canli_baslat.cmd") > 0 Or _
        InStr(command, "hoytas_sabit_baslatici.vbs") > 0 Or _
        InStr(command, "hoytas_ac.cmd") > 0 Or _
        InStr(command, "baslat.vbs") > 0 Or InStr(command, "baslat.bat") > 0
End Function

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
