@echo off
::	David Segura
::	https://winpeguy.com/
::	Sets My Computer / This PC on Desktop to show This PC %ComputerName%

::	Windows XP
::	Windows 7
::	Windows 8
::	Windows 8.1
::	Windows 10

reg add HKLM\SOFTWARE\Classes\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D} /v LocalizedString /t REG_EXPAND_SZ /d "This PC %%ComputerName%%" /f