Option Explicit
Dim shell, files, folder, desktop, name, link, candidate, target, iconPath
Set shell = CreateObject("WScript.Shell")
Set files = CreateObject("Scripting.FileSystemObject")
folder = files.GetParentFolderName(WScript.ScriptFullName)
desktop = shell.SpecialFolders("Desktop")
target = WScript.Arguments(0)
iconPath = files.BuildPath(folder, "HOYTAS.ico")
name = "HOYTA" & ChrW(350)
For Each candidate In Array(name & ".lnk", "HOYTAS.lnk", name & " v92.lnk", name & " Canl" & ChrW(305) & ".lnk")
    If candidate = name & ".lnk" Or files.FileExists(files.BuildPath(desktop, candidate)) Then
        Set link = shell.CreateShortcut(files.BuildPath(desktop, candidate))
        link.TargetPath = target
        link.Arguments = Chr(34) & files.BuildPath(folder, "HOYTAS.pyw") & Chr(34)
        link.IconLocation = iconPath & ",0"
        link.WorkingDirectory = folder
        link.Description = "HOYTAS Finans ve Isletme Yonetimi"
        link.Save
    End If
Next
