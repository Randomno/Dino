Set objShell = CreateObject("WScript.Shell")

strStartupPath = objShell.SpecialFolders("Startup")

Set objShortcut = objShell.CreateShortcut(strStartupPath & "\BasiliskStartup.lnk")

objShortcut.TargetPath = "C:\basilisk\basilisk.exe"

objShortcut.Arguments = "C:\a.htm"

objShortcut.Save