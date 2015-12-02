@echo off

::	https://technet.microsoft.com/en-us/library/hh824855.aspx

echo Dism /Online /Export-DefaultAppAssociations:"%~dp0AppAssoc.xml"
Dism /Online /Export-DefaultAppAssociations:"%~dp0AppAssoc.xml"
pause