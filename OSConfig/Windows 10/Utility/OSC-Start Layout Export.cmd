@echo off
::	David Segura
::	https://winpeguy.com/
::	https://technet.microsoft.com/en-us/library/mt592638(v=vs.85).aspx

Set Arch=x64
If /I %Processor_Architecture%==x86 Set Arch=x86

echo powershell Export-StartLayout -Path '%~dp0LayoutModification%Arch%.xml' -Verbose
powershell Export-StartLayout -Path '%~dp0LayoutModification%Arch%.xml' -Verbose
pause