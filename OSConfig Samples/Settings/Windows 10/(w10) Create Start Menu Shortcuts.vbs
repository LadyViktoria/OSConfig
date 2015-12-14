'David Segura
'winpeguy.com

set WshShell = WScript.CreateObject("WScript.Shell")
strStartMenu = WshShell.SpecialFolders("AllUsersStartMenu")

'Create Internet Explorer Shortcut
set oShellLink = WshShell.CreateShortcut(strStartMenu & "\Programs\Accessories\Internet Explorer.lnk")
oShellLink.TargetPath = "C:\Program Files\Internet Explorer\iexplore.exe"
oShellLink.WindowStyle = 1
oShellLink.IconLocation = "C:\Program Files\Internet Explorer\iexplore.exe, 0"
oShellLink.Description = "Finds and displays information and Web sites on the Internet."
oShellLink.WorkingDirectory = "%HOMEDRIVE%%HOMEPATH%"
oShellLink.Save

'Create Command Prompt Shortcut
set oShellLink = WshShell.CreateShortcut(strStartMenu & "\Programs\System Tools\Command Prompt.lnk")
oShellLink.TargetPath = "%windir%\system32\cmd.exe"
oShellLink.WindowStyle = 1
oShellLink.IconLocation = "%windir%\system32\cmd.exe, 0"
oShellLink.Description = "Performs text-based (command-line) functions."
oShellLink.WorkingDirectory = "%HOMEDRIVE%%HOMEPATH%"
oShellLink.Save