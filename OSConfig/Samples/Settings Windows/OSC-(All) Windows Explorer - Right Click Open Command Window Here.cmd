@echo off
::	David Segura
::	https://winpeguy.com/
::	Enables Open Command Window Here on all files and directories

::	Windows XP
::	Windows 7
::	Windows 8
::	Windows 8.1
::	Windows 10

reg delete HKCR\Directory\Background\shell\cmd /v Extended /f
reg delete HKCR\Directory\shell\cmd /v Extended /f
reg delete HKCR\Drive\shell\cmd /v Extended /f
