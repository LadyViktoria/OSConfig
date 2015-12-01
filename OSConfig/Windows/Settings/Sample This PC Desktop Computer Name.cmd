::	Set Computer Name
reg add HKLM\SOFTWARE\Classes\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D} /v LocalizedString /t REG_EXPAND_SZ /d "This PC %%ComputerName%%" /f