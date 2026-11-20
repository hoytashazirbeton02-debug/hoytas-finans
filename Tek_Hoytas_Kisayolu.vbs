Option Explicit
Dim shell, files, folder, desktop, launcher, shortcutPath, shortcut
Set shell = CreateObject("WScript.Shell")
Set files = CreateObject("Scripting.FileSystemObject")
folder = files.GetParentFolderName(WScript.ScriptFullName)
desktop = shell.SpecialFolders("Desktop")
launcher = files.BuildPath(folder, "Windows_Canli_Baslat.cmd")
If Not files.FileExists(launcher) Then WScript.Quit 1

shortcutPath = files.BuildPath(desktop, "HOYTA" & ChrW(350) & ".lnk")
If Not files.FileExists(shortcutPath) And files.FileExists(files.BuildPath(desktop, "HOYTAS.lnk")) Then
    shortcutPath = files.BuildPath(desktop, "HOYTAS.lnk")
End If
If files.FileExists(shortcutPath) And Not files.FileExists(shortcutPath & ".bak") Then
    files.CopyFile shortcutPath, shortcutPath & ".bak", False
End If
Set shortcut = shell.CreateShortcut(shortcutPath)
shortcut.TargetPath = shell.ExpandEnvironmentStrings("%ComSpec%")
shortcut.Arguments = "/d /c " & Chr(34) & Chr(34) & launcher & Chr(34) & Chr(34)
shortcut.WorkingDirectory = folder
shortcut.Description = "HOYTAS Finans ve Isletme Yonetimi"
shortcut.Save
