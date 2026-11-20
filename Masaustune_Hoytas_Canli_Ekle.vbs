Option Explicit

Dim shell, files, folder, launcher, shortcut, displayName
Set shell = CreateObject("WScript.Shell")
Set files = CreateObject("Scripting.FileSystemObject")
folder = files.GetParentFolderName(WScript.ScriptFullName)
launcher = files.BuildPath(folder, "Windows_Canli_Baslat.cmd")
displayName = "HOYTA" & ChrW(350) & " Canl" & ChrW(305)

If Not files.FileExists(launcher) Then
    MsgBox "Windows_Canli_Baslat.cmd bulunamadi. ZIP dosyasinin tamamini ayiklayin.", vbExclamation, displayName
    WScript.Quit 1
End If

Set shortcut = shell.CreateShortcut(files.BuildPath(shell.SpecialFolders("Desktop"), displayName & ".lnk"))
shortcut.TargetPath = shell.ExpandEnvironmentStrings("%ComSpec%")
shortcut.Arguments = "/d /c " & Chr(34) & Chr(34) & launcher & Chr(34) & Chr(34)
shortcut.WorkingDirectory = folder
shortcut.Description = "HOYTAS canli profilini ac"
shortcut.Save

MsgBox displayName & " kisayolu masaustune eklendi.", vbInformation, displayName
