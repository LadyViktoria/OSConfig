@echo off
::	David Segura
::	https://winpeguy.com/
::	Adds Right Click to Take Ownership

::	Windows 7
::	Windows 8
::	Windows 8.1
::	Windows 10

reg add HKLM\Software\Classes\*\shell\runas /ve /d "Take Ownership" /f
reg add HKLM\Software\Classes\*\shell\runas /v NoWorkingDirectory /d "" /f

reg add HKLM\Software\Classes\*\shell\runas\command /ve /d "cmd.exe /c takeown /f \"%%1\" && icacls \"%%1\" /grant administrators:F" /f
reg add HKLM\Software\Classes\*\shell\runas\command /v IsolatedCommand /d "cmd.exe /c takeown /f \"%%1\" && icacls \"%%1\" /grant administrators:F" /f

reg add HKLM\Software\Classes\Directory\shell\runas /ve /d "Take Ownership" /f
reg add HKLM\Software\Classes\Directory\shell\runas /v NoWorkingDirectory /d "" /f

reg add HKLM\Software\Classes\Directory\shell\runas\command /ve /d "cmd.exe /c takeown /f \"%%1\" /r /d y && icacls \"%%1\" /grant administrators:F /t" /f
reg add HKLM\Software\Classes\Directory\shell\runas\command /v IsolatedCommand /d "cmd.exe /c takeown /f \"%%1\" /r /d y && icacls \"%%1\" /grant administrators:F /t" /f
