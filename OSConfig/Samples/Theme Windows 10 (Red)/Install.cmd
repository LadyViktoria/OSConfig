@echo off
::	David Segura
::	https://winpeguy.com/

:: Theme Cleanup
del "%WinDir%\Resources\Themes\theme1.theme" /f /q
del "%WinDir%\Resources\Themes\theme2.theme" /f /q

::	Wallpaper Cleanup
del %WinDir%\Web\Screen\img100.jpg /f /q
del %WinDir%\Web\Screen\img101.jpg /f /q
del %WinDir%\Web\Screen\img101.png /f /q
del %WinDir%\Web\Screen\img102.jpg /f /q
del %WinDir%\Web\Screen\img103.jpg /f /q
del %WinDir%\Web\Screen\img103.png /f /q
del %WinDir%\Web\Screen\img104.jpg /f /q
del %WinDir%\Web\Screen\img105.jpg /f /q
rd "%WinDir%\Web\Wallpaper\Theme1" /s /q
rd "%WinDir%\Web\Wallpaper\Theme2" /s /q

::	Lock Screen - Force a specific default lock screen image
::	This is set automatically with the by OSConfig if C:\Windows\OSConfig\Theme\Wallpaper\LockScreen.jpg exists
::	reg add HKLM\Software\Policies\Microsoft\Windows\Personalization /v LockScreenImage /t REG_SZ /d "%SystemRoot%\Web\Screen\LockScreen.jpg" /f


::	Lock Screen - Prevent changing lock screen image
reg add HKLM\Software\Policies\Microsoft\Windows\Personalization /v NoChangingLockScreen /t REG_DWORD /d 1 /f


::	Logon Screen - Remove the Default Hero Background
::	This is set automatically with the by OSConfig if C:\Windows\OSConfig\Theme\Wallpaper\LockScreen.jpg exists
::	reg add HKLM\Software\Policies\Microsoft\Windows\System /v DisableLogonBackgroundImage /t REG_DWORD /d 1 /f



::	OEM Logo - Set the System Properties OEM Logo
::	This is set automatically with the by OSConfig if C:\Windows\OSConfig\Theme\Logos\OEMLogo.bmp exists
::	reg add HKLM\Software\Microsoft\Windows\CurrentVersion\OEMInformation /v Logo /t REG_SZ /d "%SystemRoot%\OSConfig\Theme\Logos\OEMLogo.bmp" /f
