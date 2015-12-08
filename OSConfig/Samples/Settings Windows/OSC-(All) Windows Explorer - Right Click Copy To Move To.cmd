@echo off
::	David Segura
::	https://winpeguy.com/
::	Adds Right Click Copy To and Move To

::	Windows XP
::	Windows 7
::	Windows 8
::	Windows 8.1
::	Windows 10

reg add "HKLM\Software\Classes\AllFileSystemObjects\shellex\ContextMenuHandlers\Copy To" /ve /d {C2FBB630-2971-11D1-A18C-00C04FD75D13} /f
reg add "HKLM\Software\Classes\AllFileSystemObjects\shellex\ContextMenuHandlers\Move To" /ve /d {C2FBB631-2971-11D1-A18C-00C04FD75D13} /f
