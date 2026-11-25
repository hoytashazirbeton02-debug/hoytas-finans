Option Explicit
Dim shell, files, folder, desktop, launcher, icon, shortcutPath, link, arguments
Set shell = CreateObject("WScript.Shell")
Set files = CreateObject("Scripting.FileSystemObject")
folder = files.GetParentFolderName(WScript.ScriptFullName)
desktop = shell.SpecialFolders("Desktop")
launcher = files.BuildPath(folder, "Baslat.vbs")
icon = files.BuildPath(folder, "HOYTAS.ico")
If Not files.FileExists(launcher) Or Not files.FileExists(icon) Then WScript.Quit 1

shortcutPath = files.BuildPath(desktop, "HOYTA" & ChrW(350) & ".lnk")
If Not files.FileExists(shortcutPath) And files.FileExists(files.BuildPath(desktop, "HOYTAS.lnk")) Then
    shortcutPath = files.BuildPath(desktop, "HOYTAS.lnk")
End If
Set link = shell.CreateShortcut(shortcutPath)
arguments = LCase(link.Arguments)
If InStr(arguments, "hoytas.pyw") > 0 Or InStr(arguments, "windows_canli_baslat.cmd") > 0 Then
    shortcutPath = files.BuildPath(desktop, "HOYTA" & ChrW(350) & " Yerel.lnk")
    Set link = shell.CreateShortcut(shortcutPath)
End If
link.TargetPath = shell.ExpandEnvironmentStrings("%SystemRoot%\System32\wscript.exe")
link.Arguments = Chr(34) & launcher & Chr(34)
link.WorkingDirectory = folder
link.IconLocation = icon & ",0"
link.Description = "HOYTAS yerel profilini ac"
link.Save
