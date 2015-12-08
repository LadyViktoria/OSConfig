'============================================================================================== DISCLAIMER
' This script, macro, and other code examples for illustration only, without warranty either expressed or implied, including but not
' limited to the implied warranties of merchantability and/or fitness for a particular purpose. This script is provided 'as is' and the Author does not
' guarantee that the following script, macro, or code can be used in all situations.
'
' I work for an Enterprise Corporation, and am not a Consultant
' This script is provided free of charge without support - David
'============================================================================================== DECLARATIONS
	Option Explicit
	'On Error Resume Next
'============================================================================================== USER CONFIGURABLE CONSTANTS

	Const Author			=	"David Segura"
	Const AuthorEmail		=	"david@segura.org"
	Const Company			=	""
	Const Script			=	"OSConfig.vbs"
	Const Description		=	"Script to import OSConfig Files"
	Const Release			=	"https://winpeguy.wordpress.com/"
	Const Reference			=	"https://winpeguy.wordpress.com/"

	Const Title 			=	"OSConfig"
	Const Version 			=	20151208
	Const VersionFull 		=	20151208.1
	Dim TitleVersion		:	TitleVersion = Title & " (" & Version & ")"
	
	Const SupportContact	=	"David Segura"
	Const SupportEmail		=	"david@segura.org"
	Const SupportAction		=	"Submit an Incident for Technical Support"
	Const SupportGroup		=	"SupportGroup"
	Const SupportArea		=	"SupportArea"
	Const SupportSubject	=	"OSConfig Issue"
	Const SupportProblem	=	"Complete Description of Problem Including All Logs"
	
'============================================================================================== VERSIONS
'	20151204	Theme: AddedOEMLogo.bmp support
'				Theme: Removed Logon Background Image if Lock Screen Image exists
'	20151203	Updated sorting for Appx Package Exports
'	20151202	Added functions for AppAssoc, currently in testing
'				Corrected logging of OEM Folders
'				Renamed Registry Backup Reg files for PostOSConfig
'				Resolved issue when using LockScreen.jpg
'	20151201	Initial Release
'============================================================================================== SYSTEM CONSTANTS

	Const ForReading			=	1
	Const ForWriting			=	2
	Const ForAppending			=	8
	Const OverwriteExisting 	=	True
	
	Const TristateFalse			=	0
	Const TristateTrue			=	-1
	Const TristateUseDefault	=	-2
	

	Const HKEY_CLASSES_ROOT		= 	&H80000000
	Const HKEY_CURRENT_USER		= 	&H80000001
	Const HKEY_LOCAL_MACHINE	= 	&H80000002
	Const HKEY_USERS			= 	&H80000003
	Const HKEY_CURRENT_CONFIG	= 	&H80000005

'============================================================================================== OBJECTS

	Dim objComputer				: 	objComputer				=	"."
	Dim objShell				: 	Set objShell			=	CreateObject("Wscript.Shell")
	Dim objShellApp				: 	Set objShellApp			=	CreateObject("Shell.Application")
	Dim objFSO					: 	Set objFSO 				=	CreateObject("Scripting.FileSystemObject")
	Dim objDictionary			: 	Set objDictionary		=	CreateObject("Scripting.Dictionary")
	Dim objRegEx				:	Set objRegEx			=	CreateObject("VBScript.RegExp")
	
	'Be aware that WMI does not run in Windows Setup First Boot
	'These entries are in here for compatibility with testing on a Live OS
	Dim objWMIService			: 	Set objWMIService 		=	GetObject("winmgmts:{impersonationLevel = impersonate}!\\" & objComputer & "\root\cimv2")
	Dim objRegistry				: 	Set objRegistry 		=	GetObject("winmgmts:{impersonationLevel = impersonate}!\\" & objComputer & "\root\default:StdRegProv")

'============================================================================================== VARIABLES: SYSTEM
	
	Dim MyUserName				: 	MyUserName				= Lcase(objShell.ExpandEnvironmentStrings("%UserName%"))
	Dim MyComputerName			: 	MyComputerName			= Ucase(objShell.ExpandEnvironmentStrings("%ComputerName%"))
	Dim MyTemp					: 	MyTemp					= Lcase(objShell.ExpandEnvironmentStrings("%Temp%"))
	Dim MyWindir				: 	MyWindir				= Lcase(objShell.ExpandEnvironmentStrings("%Windir%"))
	Dim MySystemDrive			: 	MySystemDrive			= Lcase(objShell.ExpandEnvironmentStrings("%SystemDrive%"))
	Dim MyArchitecture			: 	MyArchitecture			= Lcase(objShell.ExpandEnvironmentStrings("%Processor_Architecture%"))
	If MyArchitecture		= "amd64" Then MyArchitecture = "x64"
	
	'Wscript.Echo ""
	'Wscript.Echo "<Property> MyUserName: " & MyUserName
	'Wscript.Echo "<Property> MyComputerName: " & MyComputerName
	'Wscript.Echo "<Property> MyTemp: " & MyTemp
	'Wscript.Echo "<Property> MyWindir: " & MyWindir
	'Wscript.Echo "<Property> MySystemDrive: " & MySystemDrive
	'Wscript.Echo "<Property> MyArchitecture: " & MyArchitecture
	'Wscript.Echo "<Property> MyUserName: " & MyUserName
	'Wscript.Echo ""

'============================================================================================== VARIABLES: CURRENT DIRECTORY

	Dim MyScriptFullPath		: 	MyScriptFullPath			= Wscript.ScriptFullName							'Full Path and File Name with Extension
	Dim MyScriptFileName		: 	MyScriptFileName			= objFSO.GetFileName(MyScriptFullPath)				'File Name with Extension
	Dim MyScriptBaseName		: 	MyScriptBaseName			= objFSO.GetBaseName(MyScriptFullPath)				'File Name 
	Dim MyScriptParentFolder	: 	MyScriptParentFolder		= objFSO.GetParentFolderName(MyScriptFullPath)		'Current Directory (Parent Folder)
	Dim MyScriptGParentFolder	: 	MyScriptGParentFolder		= objFSO.GetParentFolderName(MyScriptParentFolder)	'Parent of the Current Directory (Parent of the Parent Folder)
	Dim arrNames				:	arrNames					= Split(MyScriptParentFolder, "\")
	Dim intIndex				:	intIndex					= Ubound(arrNames)
	Dim MyParentFolderName		:	MyParentFolderName			= arrNames(intIndex)
	
	'Wscript.Echo ""
	'Wscript.Echo "<Property> MyScriptFullPath: " & MyScriptFullPath
	'Wscript.Echo "<Property> MyScriptFileName: " & MyScriptFileName
	'Wscript.Echo "<Property> MyScriptBaseName: " & MyScriptBaseName
	'Wscript.Echo "<Property> MyScriptParentFolder: " & MyScriptParentFolder
	'Wscript.Echo "<Property> MyScriptGParentFolder: " & MyScriptGParentFolder
	'Wscript.Echo "<Property> MyParentFolderName: " & MyParentFolderName
	'Wscript.Echo ""
	
'============================================================================================== VARIABLES: LOGGING
'Wscript.Echo "Processing Logging"
	'Only one line below must be uncommented
	Dim MyLogFile				: 	MyLogFile					= MyTemp & "\" & Title & ".log"						'Places the LOG in the Temp Directory	

	'Only one line below must be uncommented
	Dim DoLogging				: 	DoLogging					= True		'Creates a LOG
	'Dim DoLogging				: 	DoLogging					= False		'Prevents a LOG from being written

	'Only one line below must be uncommented
	'Dim TextFormat				: 	TextFormat					= True		'Results in a TEXT formatted LOG
	Dim TextFormat				: 	TextFormat					= False		'Results in a CMTRACE formatted LOG (default)
	
	LogStart					'Generate the LOG file

'==============================================================================================
'==============================================================================================
	'Gets the current date as 8 digit like 20150505
	Dim MyFullDate				:	MyFullDate					= Year(Date) & Right(String(2, "0") & Month(date), 2) & Right(String(2, "0") & Day(date), 2)
	Dim objTextFile
	Dim Return
	Dim Failed

	TraceLog "============================================================= Checking Admin Rights", 2
	'Wscript.Echo "Checking Admin Rights"
	IsAdmin						'Will return IsAdmin = True if it is running with Admin Rights

	TraceLog "============================================================= Checking System Account", 2
	'Wscript.Echo "Checking System Account"
	IsSystem					'Checks to see if this is running under the System Account

	TraceLog "============================================================= Processing Operating System", 2
	Dim MyOperatingSystem
	
	'Set MyOperatingSystem to Unknown by default
	MyOperatingSystem = "Unknown"
	
	'Get Operating System information from WMI
	If MyOperatingSystem = "Unknown" Then
		TraceLog "Getting Operating System information from WMI", 1
		GetMyOperatingSystem		'Checks the Operating System.  We can stop specific OS's in this Sub
	End If
'==============================================================================================
'==============================================================================================
	Dim sCmd




	TraceLog "============================================================= Checking Command Line Arguments", 2
	Dim sArgumentUAC
	CheckArguments
	
	TraceLog "============================================================= Processing Elevation", 2
	Elevation

	TraceLog "============================================================= Creating Local Directories", 2
	OSConfigCreateLocalDirectories
	
	TraceLog "============================================================= Creating Registry Snapshot CMD", 2
	TraceLog "Reference: https://winpeguy.wordpress.com/2015/12/04/tool-osconfig-registry-snapshot/", 1
	OSConfigCreateRegistryCMD

	TraceLog "============================================================= Merging OSConfig to Production Format", 2
	OSConfigMergeFolders
	
	TraceLog "============================================================= Removing Unnecessary Content", 2
	OSConfigCleanupContent

	TraceLog "============================================================= Processing OEM Folders", 2
	TraceLog "Reference: https://winpeguy.wordpress.com/2015/12/01/tool-osconfig-oem-folders/", 1
	OSConfigOEMFolders

	TraceLog "============================================================= Mounting Default User Hive", 2
	OSConfigMountDefaultUser
	
	TraceLog "============================================================= Mounting Administrator Hive", 2
	OSConfigMountAdministrator

	TraceLog "============================================================= Creating Registry Backup Before OSConfig", 2
	TraceLog "Reference: https://winpeguy.wordpress.com/2015/12/04/tool-osconfig-registry-backup/", 1
	OSConfigRegistryBackupPre

	TraceLog "============================================================= Backup AppxPackages Before OSConfig", 2
	TraceLog "Reference: https://winpeguy.wordpress.com/2015/12/04/tool-osconfig-logs/", 1
	OSConfigAppxPackagesPre
	
	TraceLog "============================================================= Export-DefaultAppAssociations Before OSConfig", 2
	TraceLog "Reference: https://winpeguy.wordpress.com/2015/12/04/tool-osconfig-logs/", 1
	OSConfigDefaultAppAssociationsPre
	
	TraceLog "============================================================= Windows 10 Disable Consumer Experiences", 2
	TraceLog "Reference: https://winpeguy.wordpress.com/2015/12/06/win10-start-menu-junk-and-candy-crush-soda-saga/", 1
	OSConfigConsumerExperiences

	TraceLog "============================================================= Processing Theme Files", 2
	TraceLog "Reference: https://winpeguy.wordpress.com/2015/12/04/tool-osconfig-theme-sample-windows-10/", 1
	OSConfigTheme

	TraceLog "============================================================= Processing Settings", 2
	TraceLog "Reference: https://winpeguy.wordpress.com/2015/12/04/tool-osconfig-settings/", 1
	ApplyConfigs(MyScriptParentFolder & "\Settings")
	
	TraceLog "============================================================= List AppxPackages After OSConfig", 2
	TraceLog "Reference: https://winpeguy.wordpress.com/2015/12/04/tool-osconfig-logs/", 1
	OSConfigAppxPackagesPost
	
	TraceLog "============================================================= Export-DefaultAppAssociations After OSConfig", 2
	TraceLog "Reference: https://winpeguy.wordpress.com/2015/12/04/tool-osconfig-logs/", 1
	OSConfigDefaultAppAssociationsPost
	
	TraceLog "============================================================= Creating Registry Backup After OSConfig", 2
	TraceLog "Reference: https://winpeguy.wordpress.com/2015/12/04/tool-osconfig-registry-backup/", 1
	OSConfigRegistryBackupPost
	
	TraceLog "============================================================= Unmounting Default User Hive", 2
	OSConfigUnmountDefaultUser
	
	TraceLog "============================================================= Unmounting Administrator Hive", 2
	OSConfigUnmountAdministrator
	
	TraceLog "============================================================= Copy Log and Exit", 2
	OSConfigExit
'==============================================================================================
'==============================================================================================





















'==============================================================================================
'==============================================================================================
Sub OSConfigCreateLocalDirectories
	funBuildDir(MyWindir & "\OSConfig")
	funBuildDir(MyWindir & "\OSConfig\Logs")
	funBuildDir(MyWindir & "\OSConfig\Registry")
	funBuildDir(MyWindir & "\OSConfig\Registry\Backup")
	funBuildDir(MyWindir & "\OSConfig\Settings")
	funBuildDir(MyWindir & "\OSConfig\Settings\Administrator")
	funBuildDir(MyWindir & "\OSConfig\Settings\Default User")
	funBuildDir(MyWindir & "\OSConfig\Settings\RunOnce")
	funBuildDir(MyWindir & "\OSConfig\Settings\NotApplicable")
	TraceLog "Section Complete", 1
End Sub
'==============================================================================================
'==============================================================================================
Sub OSConfigCreateRegistryCMD
	Dim sFile
	sFile = MyWindir & "\OSConfig\Registry\_Snapshot.cmd"
	
	TraceLog "Building: " & sFile, 1
	
	Dim oFile
	Set oFile = objFSO.CreateTextFile(sFile, True, False)
	oFile.WriteLine "@echo off"
	oFile.WriteLine "echo Creating Registry Backup"
	oFile.WriteLine "reg load HKLM\DefaultUser %SystemDrive%\Users\Default\NTUser.dat"
	oFile.WriteLine "reg export HKCU ""%WinDir%\OSConfig\Registry\HKEY_CURRENT_USER.reg"" /y"
	oFile.WriteLine "reg export HKLM\DefaultUser ""%WinDir%\OSConfig\Registry\DEFAULT_USER.reg"" /y"
	oFile.WriteLine "reg export HKU\.DEFAULT ""%WinDir%\OSConfig\Registry\HKEY_USERS.DEFAULT.reg"" /y"
	oFile.WriteLine "reg export HKLM\SOFTWARE\Microsoft ""%WinDir%\OSConfig\Registry\SOFTWARE.Microsoft.reg"" /y"
	oFile.WriteLine "reg export HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer ""%WinDir%\OSConfig\Registry\SOFTWARE.Microsoft.Windows.CurrentVersion.Explorer.reg"" /y"
	oFile.WriteLine "reg export HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies ""%WinDir%\OSConfig\Registry\SOFTWARE.Microsoft.Windows.CurrentVersion.Policies.reg"" /y"
	oFile.WriteLine "reg export HKLM\SOFTWARE\Policies ""%WinDir%\OSConfig\Registry\SOFTWARE.Policies.reg"" /y"
	oFile.WriteLine "reg export HKLM\SYSTEM ""%WinDir%\OSConfig\Registry\SYSTEM.reg"" /y"
	oFile.WriteLine "reg unload HKLM\DefaultUser"
	oFile.WriteLine "echo."
	
	oFile.WriteLine "echo REG files have been captured, make changes before pressing a key"
	oFile.WriteLine "pause"
	oFile.WriteLine "reg load HKLM\DefaultUser %SystemDrive%\Users\Default\NTUser.dat"
	oFile.WriteLine "reg export HKCU ""%WinDir%\OSConfig\Registry\HKEY_CURRENT_USER.Changed.reg"" /y"
	oFile.WriteLine "reg export HKLM\DefaultUser ""%WinDir%\OSConfig\Registry\DEFAULT_USER.Changed.reg"" /y"
	oFile.WriteLine "reg export HKU\.DEFAULT ""%WinDir%\OSConfig\Registry\HKEY_USERS.DEFAULT.Changed.reg"" /y"
	oFile.WriteLine "reg export HKLM\SOFTWARE\Microsoft ""%WinDir%\OSConfig\Registry\SOFTWARE.Microsoft.Changed.reg"" /y"
	oFile.WriteLine "reg export HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer ""%WinDir%\OSConfig\Registry\SOFTWARE.Microsoft.Windows.CurrentVersion.Explorer.Changed.reg"" /y"
	oFile.WriteLine "reg export HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies ""%WinDir%\OSConfig\Registry\SOFTWARE.Microsoft.Windows.CurrentVersion.Policies.Changed.reg"" /y"
	oFile.WriteLine "reg export HKLM\SOFTWARE\Policies ""%WinDir%\OSConfig\Registry\SOFTWARE.Policies.Changed.reg"" /y"
	oFile.WriteLine "reg export HKLM\SYSTEM ""%WinDir%\OSConfig\Registry\SYSTEM.Changed.reg"" /y"
	oFile.WriteLine "reg unload HKLM\DefaultUser"
	oFile.WriteLine "PING 127.0.0.1 -n 3 > NUL"
	oFile.Close
	TraceLog "Section Complete", 1
End Sub
'==============================================================================================
'==============================================================================================
Sub OSConfigMergeFolders
	If objFSO.FolderExists(MyWindir & "\OSConfig\Windows") Then
		TraceLog "Merging " & MyWindir & "\OSConfig\Windows", 1
		objFSO.CopyFolder MyWindir & "\OSConfig\Windows", MyWindir & "\OSConfig", True
	End If
	
	If objFSO.FolderExists(MyWindir & "\OSConfig\" & MyOperatingSystem) Then
		TraceLog "Merging " & MyWindir & "\OSConfig\" & MyOperatingSystem, 1
		objFSO.CopyFolder MyWindir & "\OSConfig\" & MyOperatingSystem, MyWindir & "\OSConfig", True
	End If
	
	If objFSO.FolderExists(MyWindir & "\OSConfig\$OEM$\$$ " & MyArchitecture) Then
		TraceLog "Merging " & MyWindir & "\OSConfig\$OEM$\$$ " & MyArchitecture, 1
		objFSO.CopyFolder MyWindir & "\OSConfig\$OEM$\$$ " & MyArchitecture, MyWindir & "\OSConfig\$OEM$\$$", True
	End If
	
	If objFSO.FolderExists(MyWindir & "\OSConfig\$OEM$\$1 " & MyArchitecture) Then
		TraceLog "Merging " & MyWindir & "\OSConfig\$OEM$\$1 " & MyArchitecture, 1
		objFSO.CopyFolder MyWindir & "\OSConfig\$OEM$\$1 " & MyArchitecture, MyWindir & "\OSConfig\$OEM$\$1", True
	End If
	TraceLog "Section Complete", 1
End Sub
'==============================================================================================
'==============================================================================================
Sub OSConfigCleanupContent
	If objFSO.FolderExists(MyWindir & "\OSConfig\Windows") Then objFSO.DeleteFolder MyWindir & "\OSConfig\Windows", True
	If objFSO.FolderExists(MyWindir & "\OSConfig\Windows XP") Then objFSO.DeleteFolder MyWindir & "\OSConfig\Windows XP", True
	If objFSO.FolderExists(MyWindir & "\OSConfig\Windows 7") Then objFSO.DeleteFolder MyWindir & "\OSConfig\Windows 7", True
	If objFSO.FolderExists(MyWindir & "\OSConfig\Windows 8") Then objFSO.DeleteFolder MyWindir & "\OSConfig\Windows 8", True
	If objFSO.FolderExists(MyWindir & "\OSConfig\Windows 8.1") Then objFSO.DeleteFolder MyWindir & "\OSConfig\Windows 8.1", True
	If objFSO.FolderExists(MyWindir & "\OSConfig\Windows 10") Then objFSO.DeleteFolder MyWindir & "\OSConfig\Windows 10", True
	If objFSO.FolderExists(MyWindir & "\OSConfig\$OEM$\$$ x86") Then objFSO.DeleteFolder MyWindir & "\OSConfig\$OEM$\$$ x86", True
	If objFSO.FolderExists(MyWindir & "\OSConfig\$OEM$\$$ x64") Then objFSO.DeleteFolder MyWindir & "\OSConfig\$OEM$\$$ x64", True
	If objFSO.FolderExists(MyWindir & "\OSConfig\$OEM$\$1 x86") Then objFSO.DeleteFolder MyWindir & "\OSConfig\$OEM$\$1 x86", True
	If objFSO.FolderExists(MyWindir & "\OSConfig\$OEM$\$1 x64") Then objFSO.DeleteFolder MyWindir & "\OSConfig\$OEM$\$1 x64", True
	TraceLog "Section Complete", 1
End Sub
'==============================================================================================
'==============================================================================================
Sub OSConfigOEMFolders
	If objFSO.FolderExists(MyWindir & "\OSConfig\$OEM$\$$") Then
		sCmd = "robocopy %WinDir%\OSConfig\$OEM$\$$ %WinDir% *.* /e /ndl /xj /r:0 /w:0 /xf desktop.ini /LOG+:" & MyWindir & "\OSConfig\Logs\OEMWindows.log"
		TraceLog "Running Command: " & sCmd, 1
		objShell.Run sCmd, 1, True
	Else
		TraceLog "Not Found: " & MyWindir & "\OSConfig\$OEM$\$$", 1
	End If

	If objFSO.FolderExists(MyWindir & "\OSConfig\$OEM$\$$ " & MyArchitecture) Then
		sCmd = "robocopy " & """%WinDir%\OSConfig\$OEM$\$$ " & MyArchitecture & Chr(34) & " %WinDir% *.* /e /ndl /xj /r:0 /w:0 /xf desktop.ini /LOG+:" & MyWindir & "\OSConfig\Logs\OEMWindows.log"
		TraceLog "Running Command: " & sCmd, 1
		objShell.Run sCmd, 1, True
	Else
		TraceLog "Not Found: " & MyWindir & "\OSConfig\$OEM$\$$ " & MyArchitecture, 1
	End If
	
	If objFSO.FolderExists(MyWindir & "\OSConfig\$OEM$\$1") Then
		sCmd = "robocopy %WinDir%\OSConfig\$OEM$\$1 %SystemDrive%\ *.* /e /ndl /xj /r:0 /w:0 /xf desktop.ini /LOG+:" & MyWindir & "\OSConfig\Logs\OEMSystemDrive.log"
		TraceLog "Running Command: " & sCmd, 1
		objShell.Run sCmd, 1, True
	Else
		TraceLog "Not Found: " & MyWindir & "\OSConfig\$OEM$\$1", 1
	End If
	
	If objFSO.FolderExists(MyWindir & "\OSConfig\$OEM$\$1 " & MyArchitecture) Then
		sCmd = "robocopy " & Chr(34) & "%WinDir%\OSConfig\$OEM$\$1 " & MyArchitecture & Chr(34) & " %SystemDrive% *.* /e /ndl /xj /r:0 /w:0 /xf desktop.ini /LOG+:" & MyWindir & "\OSConfig\Logs\OEMSystemDrive.log"
		TraceLog "Running Command: " & sCmd, 1
		objShell.Run sCmd, 1, True
	Else
		TraceLog "Not Found: " & MyWindir & "\OSConfig\$OEM$\$1 " & MyArchitecture, 1
	End If
	
	'Reset ProgramData to Hidden if exists
	sCmd = "attrib C:\ProgramData +H"
	objShell.Run sCmd, 7, True
	
	TraceLog "Section Complete", 1
End Sub
'==============================================================================================
'==============================================================================================
Sub OSConfigMountDefaultUser
	On Error Resume Next
	If objFSO.FileExists("C:\Users\Default\NTUser.dat") Then
		sCmd = "reg load HKLM\DefaultUser " & """C:\Users\Default\NTUser.dat"""
		TraceLog "Running Command: " & sCmd, 1
		objShell.Run sCmd, 7, True
	ElseIf objFSO.FileExists("C:\Documents and Settings\Default User\NTUser.dat") Then
		sCmd = "reg load HKLM\DefaultUser " & """C:\Documents and Settings\Default User\NTUser.dat"""
		TraceLog "Running Command: " & sCmd, 1
		objShell.Run sCmd, 7, True
	ElseIf objFSO.FileExists("D:\Documents and Settings\Default User\NTUser.dat") Then
		sCmd = "reg load HKLM\DefaultUser " & """D:\Documents and Settings\Default User\NTUser.dat"""
		TraceLog "Running Command: " & sCmd, 1
		objShell.Run sCmd, 7, True
	Else
		TraceLog "Hive was NOT located", 3
	End If
	TraceLog "Section Complete", 1
End Sub
'==============================================================================================
'==============================================================================================
Sub OSConfigMountAdministrator
	On Error Resume Next
	If objFSO.FileExists("C:\Users\Administrator\NTUser.dat") Then
		sCmd = "reg load HKLM\Administrator " & """C:\Users\Administrator\NTUser.dat"""
		TraceLog "Running Command: " & sCmd, 1
		objShell.Run sCmd, 7, True
	ElseIf objFSO.FileExists("C:\Documents and Settings\Administrator\NTUser.dat") Then
		sCmd = "reg load HKLM\Administrator " & """C:\Documents and Settings\Administrator\NTUser.dat"""
		TraceLog "Running Command: " & sCmd, 1
		objShell.Run sCmd, 7, True
	ElseIf objFSO.FileExists("D:\Documents and Settings\Administrator\NTUser.dat") Then
		sCmd = "reg load HKLM\Administrator " & """D:\Documents and Settings\Administrator\NTUser.dat"""
		TraceLog "Running Command: " & sCmd, 1
		objShell.Run sCmd, 7, True
	Else
		TraceLog "Hive was NOT located", 3
	End If
	TraceLog "Section Complete", 1
End Sub
'==============================================================================================
'==============================================================================================
Sub OSConfigRegistryBackupPre
	objShell.Run "reg export HKCU %WinDir%\OSConfig\Registry\Backup\HKEY_CURRENT_USER.reg /y", 7, True
	objShell.Run "reg export HKLM\DefaultUser %WinDir%\OSConfig\Registry\Backup\DEFAULT_USER.reg /y", 7, True
	objShell.Run "reg export HKU\.DEFAULT %WinDir%\OSConfig\Registry\Backup\HKEY_USERS.DEFAULT.reg /y", 7, True
	objShell.Run "reg export HKLM\SOFTWARE\Microsoft %WinDir%\OSConfig\Registry\Backup\HKEY_LOCAL_MACHINE.SOFTWARE.Microsoft.reg /y", 7, True
	objShell.Run "reg export HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer %WinDir%\OSConfig\Registry\Backup\HKEY_LOCAL_MACHINE.SOFTWARE.Microsoft.Windows.CurrentVersion.Explorer.reg /y", 7, True
	objShell.Run "reg export HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies %WinDir%\OSConfig\Registry\Backup\HKEY_LOCAL_MACHINE.SOFTWARE.Microsoft.Windows.CurrentVersion.Policies.reg /y", 7, True
	objShell.Run "reg export HKLM\SOFTWARE\Policies %WinDir%\OSConfig\Registry\Backup\HKEY_LOCAL_MACHINE.SOFTWARE.Policies.reg /y", 7, True
	objShell.Run "reg export HKLM\SYSTEM %WinDir%\OSConfig\Registry\Backup\HKEY_LOCAL_MACHINE.SYSTEM.reg /y", 7, True
	TraceLog "Section Complete", 1
End Sub
'==============================================================================================
'==============================================================================================
Sub OSConfigAppxPackagesPre
	If MyOperatingSystem = "Windows 8" or MyOperatingSystem = "Windows 8.1" or MyOperatingSystem = "Windows 10" Then
		TraceLog "Creating list of default AppxPackages at C:\Windows\OSConfig\Logs\MyAppxPackages.txt", 1
		sCmd = "powershell Get-AppxPackage | Sort Name | Select Name | Out-File -FilePath C:\Windows\OSConfig\Logs\MyAppxPackage.txt"
		TraceLog "Running Command: " & sCmd, 1
		objShell.Run sCmd, 7, True
		sCmd = "powershell Get-AppxPackage | Sort Name | Out-File -FilePath C:\Windows\OSConfig\Logs\MyAppxPackage.txt -Append"
		TraceLog "Running Command: " & sCmd, 1
		objShell.Run sCmd, 7, True
		
		TraceLog "Creating list of default ProvisionedAppxPackages at C:\Windows\OSConfig\Logs\MyProvisionedAppxPackage.txt", 1
		sCmd = "powershell Get-ProvisionedAppxPackage -Online | Sort DisplayName | Select DisplayName | Out-File -FilePath C:\Windows\OSConfig\Logs\MyProvisionedAppxPackage.txt"
		TraceLog "Running Command: " & sCmd, 1
		objShell.Run sCmd, 7, True
		sCmd = "powershell Get-ProvisionedAppxPackage -Online | Sort DisplayName | Out-File -FilePath C:\Windows\OSConfig\Logs\MyProvisionedAppxPackage.txt -Append"
		TraceLog "Running Command: " & sCmd, 1
		objShell.Run sCmd, 7, True
	Else
		TraceLog "AppxPackages do not apply to this OS", 1
	End If
	TraceLog "Section Complete", 1
End Sub
'==============================================================================================
'==============================================================================================
Sub OSConfigDefaultAppAssociationsPre
	If MyOperatingSystem = "Windows 8" or MyOperatingSystem = "Windows 8.1" or MyOperatingSystem = "Windows 10" Then
		TraceLog "Exporting Default App Associations at C:\Windows\OSConfig\Logs\MyAppAssoc.xml", 1
		sCmd = "Dism /Online /Export-DefaultAppAssociations:C:\Windows\OSConfig\Logs\MyAppAssoc.xml"
		TraceLog "Running Command: " & sCmd, 1
		objShell.Run sCmd, 7, True
	Else
		TraceLog "Dism Export-DefaultAppAssociations does not apply to this OS", 1
	End If
	TraceLog "Section Complete", 1
End Sub
'==============================================================================================
'==============================================================================================
Sub OSConfigConsumerExperiences
	If MyOperatingSystem = "Windows 10" Then
		TraceLog "Disabling Consumer Experiences", 1
		sCmd = "reg add HKLM\SOFTWARE\Policies\Microsoft\Windows\CloudContent /v DisableWindowsConsumerFeatures /t REG_DWORD /d 1 /f"
		TraceLog "Running Command: " & sCmd, 1
		objShell.Run sCmd, 7, True
	Else
		TraceLog "Disable Consumer Experiences does not apply to this OS", 1
	End If
	TraceLog "Section Complete", 1
End Sub
'==============================================================================================
'==============================================================================================
Sub OSConfigTheme
	If MyOperatingSystem = "Windows XP" Then Exit Sub

	Dim sFile, dFile

	'Install.cmd
	sFile = MyWindir & "\OSConfig\Theme\Install.cmd"
	TraceLog "Configuring: " & sFile, 1
	If objFSO.FileExists(sFile) Then
		sCmd = "cmd /c """ & sFile & """"
		TraceLog "Running Command: " & sCmd, 1
		objShell.Run sCmd, 1, True
	Else
		TraceLog "File was NOT located.  No actions taken.", 1
	End If
	
	'Aero.theme
	sFile = MyWindir & "\OSConfig\Theme\Themes\aero.theme"
	TraceLog "Configuring: " & sFile, 1
	dFile = MyWindir & "\Resources\Themes\aero.theme"
	TraceLog "Copy " & sFile & " to " & dFile, 1
	If objFSO.FileExists(sFile) Then
		If objFSO.FileExists(dFile) Then
			sCmd = "takeown /F """ & dFile & """"
			TraceLog "Taking Ownership Using Command: " & sCmd, 1
			objShell.Run sCmd, 7, True
			sCmd = "icacls """ & dFile & """" & " /grant administrators:F"
			TraceLog "Applying Permission Using Command: " & sCmd, 1
			objShell.Run sCmd, 7, True
		End If
		TraceLog "File was located and will be copied", 1
		objFSO.CopyFile sFile, dFile, True
	Else
		TraceLog "File was NOT located.  No actions taken.", 1
	End If

	'Basic.theme
	sFile = MyWindir & "\OSConfig\Theme\Themes\basic.theme"
	TraceLog "Configuring: " & sFile, 1
	dFile = MyWindir & "\Resources\Ease of Access Themes\basic.theme"
	TraceLog "Copy " & sFile & " to " & dFile, 1
	If objFSO.FileExists(sFile) Then
		If objFSO.FileExists(dFile) Then
			sCmd = "takeown /F """ & dFile & """"
			TraceLog "Taking Ownership Using Command: " & sCmd, 1
			objShell.Run sCmd, 7, True
			sCmd = "icacls """ & dFile & """" & " /grant administrators:F"
			TraceLog "Applying Permission Using Command: " & sCmd, 1
			objShell.Run sCmd, 7, True
		End If
		TraceLog "File was located and will be copied", 1
		objFSO.CopyFile sFile, dFile, True
	Else
		TraceLog "File was NOT located.  No actions taken.", 1
	End If
	
	'OEMLogo.bmp
	sFile = MyWindir & "\OSConfig\Theme\Logos\OEMLogo.bmp"
	TraceLog "Configuring: " & sFile, 1
	If objFSO.FileExists(sFile) Then
		TraceLog "Setting Logo in OEMInformation", 1
		sCmd = "reg add HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation /v Logo /t REG_SZ /d """"%SystemRoot%\OSConfig\Theme\Logos\OEMLogo.bmp"""" /f"
		TraceLog "Running Command: " & sCmd, 1
		objShell.Run sCmd, 7, True
	Else
		TraceLog "File was NOT located.  No actions taken.", 1
	End If
	
	'img0.jpg
	sFile = MyWindir & "\OSConfig\Theme\Wallpaper\img0.jpg"
	TraceLog "Configuring: " & sFile, 1
	dFile = MyWindir & "\Web\Wallpaper\Windows\img0.jpg"
	TraceLog "Copy " & sFile & " to " & dFile, 1
	If objFSO.FileExists(sFile) Then
		If objFSO.FileExists(dFile) Then
			sCmd = "takeown /F """ & dFile & """"
			TraceLog "Taking Ownership Using Command: " & sCmd, 1
			objShell.Run sCmd, 7, True
			sCmd = "icacls """ & dFile & """" & " /grant administrators:F"
			TraceLog "Applying Permission Using Command: " & sCmd, 1
			objShell.Run sCmd, 7, True
		End If
		TraceLog "File was located and will be copied", 1
		objFSO.CopyFile sFile, dFile, True
	Else
		TraceLog "File was NOT located.  No actions taken.", 1
	End If
	
	'LockScreen.jpg
	sFile = MyWindir & "\OSConfig\Theme\Wallpaper\LockScreen.jpg"
	TraceLog "Configuring: " & sFile, 1
	dFile = MyWindir & "\Web\Screen\LockScreen.jpg"
	TraceLog "Copy " & sFile & " to " & dFile, 1
	If objFSO.FileExists(sFile) Then
		If objFSO.FileExists(dFile) Then
			sCmd = "takeown /F """ & dFile & """"
			TraceLog "Taking Ownership Using Command: " & sCmd, 1
			objShell.Run sCmd, 7, True
			sCmd = "icacls """ & dFile & """" & " /grant administrators:F"
			TraceLog "Applying Permission Using Command: " & sCmd, 1
			objShell.Run sCmd, 7, True
		End If
		TraceLog "File was located and will be copied", 1
		objFSO.CopyFile sFile, dFile, True
		TraceLog "Setting " & dFile & " to Default", 1
		sCmd = "reg add HKLM\SOFTWARE\Policies\Microsoft\Windows\Personalization /v LockScreenImage /t REG_SZ /d """"%SystemRoot%\Web\Screen\LockScreen.jpg"""" /f"
		TraceLog "Running Command: " & sCmd, 1
		objShell.Run sCmd, 7, True
		TraceLog "Disabling Logon Background Image", 1
		sCmd = "reg add HKLM\SOFTWARE\Policies\Microsoft\Windows\System /v DisableLogonBackgroundImage /t REG_DWORD /d 1 /f"
		TraceLog "Running Command: " & sCmd, 1
		objShell.Run sCmd, 7, True
	Else
		TraceLog "File was NOT located.  No actions taken.", 1
	End If

	'Background.bmp
	sFile = MyWindir & "\OSConfig\Theme\Wallpaper\Background.bmp"
	TraceLog "Configuring: " & sFile, 1
	dFile = MyWindir & "\System32\oobe\Background.bmp"
	TraceLog "Copy " & sFile & " to " & dFile, 1
	If objFSO.FileExists(sFile) Then
		If objFSO.FileExists(dFile) Then
			sCmd = "takeown /F """ & dFile & """"
			TraceLog "Taking Ownership Using Command: " & sCmd, 1
			objShell.Run sCmd, 7, True
			sCmd = "icacls """ & dFile & """" & " /grant administrators:F"
			TraceLog "Applying Permission Using Command: " & sCmd, 1
			objShell.Run sCmd, 7, True
		End If
		TraceLog "File was located and will be copied", 1
		objFSO.CopyFile sFile, dFile, True
	Else
		TraceLog "File was NOT located.  No actions taken.", 1
	End If
	
	
	'backgroundDefault.jpg
	sFile = MyWindir & "\OSConfig\Theme\Wallpaper\backgroundDefault.jpg"
	TraceLog "Configuring: " & sFile, 1
	dFile = MyWindir & "\System32\oobe\info\backgrounds\backgroundDefault.jpg"
	TraceLog "Copy " & sFile & " to " & dFile, 1
	If objFSO.FileExists(sFile) Then
		If objFSO.FileExists(dFile) Then
			sCmd = "takeown /F """ & dFile & """"
			TraceLog "Taking Ownership Using Command: " & sCmd, 1
			objShell.Run sCmd, 7, True
			sCmd = "icacls """ & dFile & """" & " /grant administrators:F"
			TraceLog "Applying Permission Using Command: " & sCmd, 1
			objShell.Run sCmd, 7, True
		End If
		If NOT objFSO.FolderExists(MyWindir & "\System32\oobe\info") Then funBuildDir(MyWindir & "\System32\oobe\info")
		If NOT objFSO.FolderExists(MyWindir & "\System32\oobe\info\backgrounds") Then funBuildDir(MyWindir & "\System32\oobe\info\backgrounds")
		TraceLog "File was located and will be copied", 1
		objFSO.CopyFile sFile, dFile, True
	Else
		TraceLog "File was NOT located.  No actions taken.", 1
	End If
	

	'User Account Pictures
	TraceLog "Configuring: " & MyWindir & "\OSConfig\Theme\User Account Pictures", 1
	TraceLog "Copy " & MyWindir & "\OSConfig\Theme\User Account Pictures to " & MySystemDrive & "\ProgramData\Microsoft\User Account Pictures", 1
	If objFSO.FolderExists(MyWindir & "\OSConfig\Theme\User Account Pictures") Then
		TraceLog "Folder was located and will be copied", 1
		objFSO.CopyFolder MyWindir & "\OSConfig\Theme\User Account Pictures", MySystemDrive & "\ProgramData\Microsoft", True
	Else
		TraceLog "Folder was NOT located.  No actions taken.", 1
	End If
	
	'DefaultLayouts.xml
	sFile = MyWindir & "\OSConfig\Theme\Start\DefaultLayouts" & MyArchitecture & ".xml"
	TraceLog "Configuring: " & sFile, 1
	dFile = MySystemDrive & "\Users\Default\AppData\Local\Microsoft\Windows\Shell\DefaultLayouts.xml"
	TraceLog "Copy " & sFile & " to " & dFile, 1
	If objFSO.FileExists(sFile) Then
		TraceLog "File was located and will be copied", 1
		objFSO.CopyFile sFile, dFile, True
	Else
		TraceLog "File was NOT located.  No actions taken.", 1
	End If

	'LayoutModification
	TraceLog "Configuring: " & MyWindir & "\OSConfig\Theme\Start\LayoutModification" & MyArchitecture & ".xml", 1
	If objFSO.FileExists(MyWindir & "\OSConfig\Theme\Start\LayoutModification" & MyArchitecture & ".xml") Then
		sCmd = "powershell -ExecutionPolicy Bypass Import-StartLayout -LayoutPath " & MyWindir & "\OSConfig\Theme\Start\LayoutModification" & MyArchitecture & ".xml -MountPath $env:SystemDrive\"
		TraceLog "Running Command: " & sCmd, 1
		objShell.Run sCmd, 1, True
	Else
		TraceLog "File was NOT located.  No actions taken.", 1
	End If
	
	TraceLog "Section Complete", 1
End Sub
'==============================================================================================
'==============================================================================================
Function ApplyConfigs(ConfigsDir)
	Dim OSConfigFile, Extension
	Dim objFile, strText, strNewText
	Dim HasRun
	
	Dim ApplyXP
	Dim Apply7
	Dim Apply8
	Dim Apply81
	Dim Apply10
	If objFSO.FolderExists(ConfigsDir) Then
		TraceLog "Checking Directory: " & ConfigsDir, 1
		On Error Resume Next
		For Each OSConfigFile In objFSO.GetFolder(ConfigsDir).Files
			If Len(OSConfigFile.Name) > 4 Then
				HasRun = False
				TraceLog "=============================================================", 1
				TraceLog "Checking File: " & OSConfigFile.Name, 1
				
				
				If Instr(UCase(OSConfigFile.Name), "SAMPLE") Then
					TraceLog OSConfigFile.Name & " is a Sample . . . Skipping", 1

				ElseIf Instr(UCase(OSConfigFile.Name), "UNDO") Then
					TraceLog OSConfigFile.Name & " is an Undo File . . . Skipping", 1
				ElseIf Instr(UCase(OSConfigFile.Name), "TSYES") and NOT objFSO.FileExists("C:\_SMSTaskSequence\OSConfig\OSConfig.cmd") Then
					TraceLog OSConfigFile.Name & " should only be run in a Task Sequence . . . Skipping", 1
				ElseIf Instr(UCase(OSConfigFile.Name), "TSNO") and objFSO.FileExists("C:\_SMSTaskSequence\OSConfig\OSConfig.cmd") Then
					TraceLog OSConfigFile.Name & " should NOT be run in a Task Sequence . . . Skipping", 1
				ElseIf Instr(UCase(OSConfigFile.Name), "MODERN") and MyOperatingSystem = "Windows XP" Then
					TraceLog OSConfigFile.Name & " requires a Modern Operating System (Windows 7 or Newer) . . . Moving to NotApplicable", 1
					objFSO.CopyFile OSConfigFile.Path, MyWindir & "\OSConfig\Settings\NotApplicable\" & OSConfigFile.Name, True
					objFSO.DeleteFile OSConfigFile.Path
				ElseIf Instr(OSConfigFile.Name, "w7+") and MyOperatingSystem = "Windows XP" Then
					TraceLog OSConfigFile.Name & " requires a Modern Operating System (Windows 7 or Newer) . . . Moving to NotApplicable", 1
					objFSO.CopyFile OSConfigFile.Path, MyWindir & "\OSConfig\Settings\NotApplicable\" & OSConfigFile.Name, True
					objFSO.DeleteFile OSConfigFile.Path
				ElseIf Instr(LCase(OSConfigFile.Name), "x86") and MyArchitecture = "x64" Then
					TraceLog OSConfigFile.Name & " requires x86 Architecture . . . Moving to NotApplicable", 1
					objFSO.CopyFile OSConfigFile.Path, MyWindir & "\OSConfig\Settings\NotApplicable\" & OSConfigFile.Name, True
					objFSO.DeleteFile OSConfigFile.Path
				ElseIf Instr(LCase(OSConfigFile.Name), "x64") and MyArchitecture = "x86" Then
					TraceLog OSConfigFile.Name & " requires x64 Architecture . . . Moving to NotApplicable", 1
					objFSO.CopyFile OSConfigFile.Path, MyWindir & "\OSConfig\Settings\NotApplicable\" & OSConfigFile.Name, True
					objFSO.DeleteFile OSConfigFile.Path
				ElseIf Instr(UCase(OSConfigFile.Name), "APPASSOC") Then
					TraceLog "Importing Default App Associations at C:\Windows\OSConfig\Settings\" & OSConfigFile.Name, 1
					sCmd = "cmd /c Dism /Online /Import-DefaultAppAssociations:""C:\Windows\OSConfig\Settings\" & OSConfigFile.Name & """"
					'sCmd = "Dism /Image:C:\ /Import-DefaultAppAssociations:""C:\Windows\OSConfig\Settings\" & OSConfigFile.Name & """"
					TraceLog "Running Command: " & sCmd, 1
					objShell.Run sCmd, 1, True
				Else
					Extension = UCase(Right(OSConfigFile.Name, 3))
					Select Case Extension
					Case "BAT":
						sCmd = "cmd /c """ & OSConfigFile.Path & """"
						TraceLog "Running Command: " & sCmd, 1
						objShell.Run sCmd, 1, True
						HasRun = True
					Case "CMD":
						sCmd = "cmd /c """ & OSConfigFile.Path & """"
						TraceLog "Running Command: " & sCmd, 1
						objShell.Run sCmd, 1, True
						HasRun = True
					Case "EXE":
						sCmd = "cmd /c """ & OSConfigFile.Path & """"
						TraceLog "Running Command: " & sCmd, 1
						objShell.Run sCmd, 1, True
						HasRun = True
					Case "PS1":
						sCmd = "powershell -ExecutionPolicy Bypass -File """ & OSConfigFile.Path & """"
						TraceLog "Running Command: " & sCmd, 1
						objShell.Run sCmd, 1, True
						HasRun = True
					Case "REG":
						'Import the REG File as is
						sCmd = "reg import """ & OSConfigFile.Path & """"
						TraceLog "Running Command: " & sCmd, 1
						objShell.Run sCmd, 7, True
						HasRun = True
						'Reg files are UNICODE and must make some changes in logic
						'Script will try to determine the format of the REG file and change accordingly
						
						'Read the Reg File as Unicode
						TraceLog "Reading Reg file: " & OSConfigFile.Path, 1
						Set objFile = objFSO.OpenTextFile(OSConfigFile.Path, ForReading, True, TristateTrue)
						strText = objFile.ReadAll
						objFile.Close
						
						'Log Reg Contents
						TraceLog strText, 1
						
						'Check if HKEY_CURRENT_USER exists and create an Administrator and a Default User Reg
						If Instr(strText, "HKEY_CURRENT_USER") Then
							'Copy REG File to New Directory
							TraceLog "Copying Reg file to: " & MyWindir & "\OSConfig\Settings\Administrator\" & OSConfigFile.Name, 1
							objFSO.CopyFile OSConfigFile, MyWindir & "\OSConfig\Settings\Administrator\" & OSConfigFile.Name
							
							'Replace Current User with Mounted Administrator
							strText = Replace(strText,"HKEY_CURRENT_USER","HKEY_LOCAL_MACHINE\Administrator")
							
							'Open Copied Reg file as Unicode
							Set objFile = objFSO.OpenTextFile(MyWindir & "\OSConfig\Settings\Administrator" & "\" & OSConfigFile.Name, ForWriting, True, TristateTrue)
							
							'Write the new contents
							objFile.WriteLine strText
							objFile.Close
							TraceLog strText, 1
							
							'Import the new Reg File
							sCmd = "reg import """ & MyWindir & "\OSConfig\Settings\Administrator\" & OSConfigFile.Name & """"
							TraceLog "Running Command: " & sCmd, 1
							objShell.Run sCmd, 7, True
							
							
							
							'Copy REG File to New Directory
							TraceLog "Copying Reg file to: " & MyWindir & "\OSConfig\Settings\Default User\" & OSConfigFile.Name, 1
							objFSO.CopyFile OSConfigFile, MyWindir & "\OSConfig\Settings\Default User\" & OSConfigFile.Name
	
							'Replace Current User with Mounted Default User
							strText = Replace(strText,"HKEY_LOCAL_MACHINE\Administrator","HKEY_LOCAL_MACHINE\DefaultUser")
							
							'Open Copied Reg file as Unicode
							Set objFile = objFSO.OpenTextFile(MyWindir & "\OSConfig\Settings\Default User" & "\" & OSConfigFile.Name, ForWriting, True, TristateTrue)
							
							'Write the new contents
							objFile.WriteLine strText
							objFile.Close
							TraceLog strText, 1
							
							'Import the new Reg File
							sCmd = "reg import """ & MyWindir & "\OSConfig\Settings\Default User\" & OSConfigFile.Name & """"
							TraceLog "Running Command: " & sCmd, 1
							objShell.Run sCmd, 7, True
						Else
							TraceLog "HKEY_CURRENT_USER was not found in the Reg file", 1
						End If
					Case "VBS":
						sCmd = "cscript """ & OSConfigFile.Path & """"
						TraceLog "Running Command: " & sCmd, 1
						objShell.Run sCmd, 1, True
						HasRun = True
					Case Else
						TraceLog "No Actions Taken", 1
					End Select
				End If
				
				'If the file has RUNONCE in the name, then we need to move it to the RunOnce SubDirectory
				If Instr(UCase(OSConfigFile.Name), "RUNONCE") and HasRun = True Then
					TraceLog OSConfigFile.Name & " should only be run once . . . Moving to RunOnce", 1
					objFSO.CopyFile OSConfigFile.Path, MyWindir & "\OSConfig\Settings\RunOnce\" & OSConfigFile.Name, True
					objFSO.DeleteFile OSConfigFile.Path
				End If
				
			End If
		Next
	End If
	TraceLog "Section Complete", 1
End Function
'==============================================================================================
'==============================================================================================
Sub OSConfigAppxPackagesPost
	If MyOperatingSystem = "Windows 8" or MyOperatingSystem = "Windows 8.1" or MyOperatingSystem = "Windows 10" Then
		TraceLog "Creating list of default AppxPackages at C:\Windows\OSConfig\Logs\MyAppxPackage-PostOSConfig.txt", 1
		sCmd = "powershell Get-AppxPackage | Sort Name | Select Name | Out-File -FilePath C:\Windows\OSConfig\Logs\MyAppxPackage-PostOSConfig.txt"
		TraceLog "Running Command: " & sCmd, 1
		objShell.Run sCmd, 7, True
		sCmd = "powershell Get-AppxPackage | Sort Name | Out-File -FilePath C:\Windows\OSConfig\Logs\MyAppxPackage-PostOSConfig.txt -Append"
		TraceLog "Running Command: " & sCmd, 1
		objShell.Run sCmd, 7, True
		
		TraceLog "Creating list of default ProvisionedAppxPackages at C:\Windows\OSConfig\Logs\MyProvisionedAppxPackage-PostOSConfig.txt", 1
		sCmd = "powershell Get-ProvisionedAppxPackage -Online | Sort DisplayName | Select DisplayName | Out-File -FilePath C:\Windows\OSConfig\Logs\MyProvisionedAppxPackage-PostOSConfig.txt"
		TraceLog "Running Command: " & sCmd, 1
		objShell.Run sCmd, 7, True
		sCmd = "powershell Get-ProvisionedAppxPackage -Online | Sort DisplayName | Out-File -FilePath C:\Windows\OSConfig\Logs\MyProvisionedAppxPackage-PostOSConfig.txt -Append"
		TraceLog "Running Command: " & sCmd, 1
		objShell.Run sCmd, 7, True
	Else
		TraceLog "AppxPackages do not apply to this OS", 1
	End If
	TraceLog "Section Complete", 1
End Sub
'==============================================================================================
'==============================================================================================
Sub OSConfigDefaultAppAssociationsPost
	If MyOperatingSystem = "Windows 8" or MyOperatingSystem = "Windows 8.1" or MyOperatingSystem = "Windows 10" Then
		TraceLog "Exporting Default App Associations at C:\Windows\OSConfig\Logs\MyAppAssoc-PostOSConfig.xml", 1
		sCmd = "Dism /Online /Export-DefaultAppAssociations:C:\Windows\OSConfig\Logs\MyAppAssoc-PostOSConfig.xml"
		TraceLog "Running Command: " & sCmd, 1
		objShell.Run sCmd, 7, True
	Else
		TraceLog "Dism Export-DefaultAppAssociations does not apply to this OS", 1
	End If
	TraceLog "Section Complete", 1
End Sub
'==============================================================================================
'==============================================================================================
Sub OSConfigRegistryBackupPost
	objShell.Run "reg export HKCU %WinDir%\OSConfig\Registry\Backup\HKEY_CURRENT_USER-PostOSConfig.reg /y", 7, True
	objShell.Run "reg export HKLM\DefaultUser %WinDir%\OSConfig\Registry\Backup\DEFAULT_USER-PostOSConfig.reg /y", 7, True
	objShell.Run "reg export HKU\.DEFAULT %WinDir%\OSConfig\Registry\Backup\HKEY_USERS.DEFAULT-PostOSConfig.reg /y", 7, True
	objShell.Run "reg export HKLM\SOFTWARE\Microsoft %WinDir%\OSConfig\Registry\Backup\HKEY_LOCAL_MACHINE.SOFTWARE.Microsoft-PostOSConfig.reg /y", 7, True
	objShell.Run "reg export HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer %WinDir%\OSConfig\Registry\Backup\HKEY_LOCAL_MACHINE.SOFTWARE.Microsoft.Windows.CurrentVersion.Explorer-PostOSConfig.reg /y", 7, True
	objShell.Run "reg export HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies %WinDir%\OSConfig\Registry\Backup\HKEY_LOCAL_MACHINE.SOFTWARE.Microsoft.Windows.CurrentVersion.Policies-PostOSConfig.reg /y", 7, True
	objShell.Run "reg export HKLM\SOFTWARE\Policies %WinDir%\OSConfig\Registry\Backup\HKEY_LOCAL_MACHINE.SOFTWARE.Policies-PostOSConfig.reg /y", 7, True
	objShell.Run "reg export HKLM\SYSTEM %WinDir%\OSConfig\Registry\Backup\HKEY_LOCAL_MACHINE.SYSTEM-PostOSConfig.reg /y", 7, True
	TraceLog "Section Complete", 1
End Sub
'==============================================================================================
'==============================================================================================
Sub OSConfigUnmountDefaultUser
	sCmd = "reg unload HKLM\DefaultUser"
	TraceLog "Running Command: " & sCmd, 1
	objShell.Run sCmd, 7, True
	TraceLog "Section Complete", 1
End Sub
'==============================================================================================
'==============================================================================================
Sub OSConfigUnmountAdministrator
	sCmd = "reg unload HKLM\Administrator"
	TraceLog "Running Command: " & sCmd, 1
	objShell.Run sCmd, 7, True
	TraceLog "Section Complete", 1
End Sub
'==============================================================================================
'==============================================================================================
Sub OSConfigExit
	If objFSO.FileExists(MyLogFile) Then
		objFSO.CopyFile MyLogFile, MyWindir & "\OSConfig\Logs\", True
	End If
	TraceLog "Section Complete", 1
End Sub
'==============================================================================================
'==============================================================================================
Function funBuildDir(strBuildDir)
	If NOT objFSO.FolderExists(strBuildDir) Then
		TraceLog "Creating Directory: " & strBuildDir, 1
		objFSO.CreateFolder(strBuildDir)
	Else
		TraceLog "Directory Exists: " & strBuildDir, 1
	End If
End Function
'==============================================================================================
'==============================================================================================

















'==============================================================================================
'==============================================================================================

	REM Constant			Value			Description
	REM vbOKOnly				0			Display OK button only.
	REM vbOKCancel				1			Display OK and Cancel buttons.
	REM vbAbortRetryIgnore		2			Display Abort, Retry, and Ignore buttons.
	REM vbYesNoCancel			3			Display Yes, No, and Cancel buttons.
	REM vbYesNo					4			Display Yes and No buttons.
	REM vbRetryCancel			5			Display Retry and Cancel buttons.
	REM vbCritical				16			Display Critical Message icon.
	REM vbQuestion				32			Display Warning Query icon.
	REM vbExclamation			48			Display Warning Message icon.
	REM vbInformation			64			Display Information Message icon.
	REM vbDefaultButton1		0			First button is default.
	REM vbDefaultButton2		256			Second button is default.
	REM vbDefaultButton3		512			Third button is default.
	REM vbDefaultButton4		768			Fourth button is default.
	REM vbApplicationModal		0			Application modal; the user must respond to the message box before continuing work in the current application.
	REM vbSystemModal			4096		System modal; all applications are suspended until the user responds to the message box.
	REM vbMsgBoxHelpButton		16384		Adds Help button to the message box
	REM VbMsgBoxSetForeground	65536		Specifies the message box window as the foreground window
	REM vbMsgBoxRight			524288		Text is right aligned
	REM vbMsgBoxRtlReading		1048576		Specifies text should appear as right-to-left reading on Hebrew and Arabic systems
	
'==============================================================================================
'==============================================================================================
	' /////////////////////////////////////////////////////////
	' Logging Function with Trace Log
	' /////////////////////////////////////////////////////////
	Function TraceLog(LogText, LogError)
		Dim LogTemp
		Dim FileOut, MyLogFileX, TitelX, Tst
	
		If DoLogging = False Then Exit Function
		
		If TextFormat = True Then
			If LogError = 0 Then
				Set FileOut = objFSO.OpenTextFile( MyLogFile, ForWriting, True)
			Else
				Set FileOut = objFSO.OpenTextFile( MyLogFile, ForAppending, True)
			End If
			FileOut.WriteLine Now()& " - " & LogText
			FileOut.Close
			Set FileOut = Nothing
			Exit Function
		End If
	
		'***********************************************************
		' Write Trace32 / CMTrace compatible log file
		' logfile - syntax (SMS Trace)
		' <![LOG[...]LOG]!>
		' <
		'    time="04:00:54.309+-60"
		'    date="03-14-2008"
		'    component="SrcUpdateMgr"
		'    context=""
		'    type="0"
		'    thread="1812"
		'    file="productpackage.cpp:97"
		' >
		'
		'    "context="		will not display
		'    type="0"		TraceLog-procedure delete logfile an create new logfile
		'    type="1"		display as normally line
		'    type="2"		display as yellow line / warn
		'    type="3"		display as red line / error
		'    type="F"		display as red line / error

		'    "thread="		number, display as "Tread:", example "Tread: 33 (0x21)"
		'    "file="		diplay as "Source:"

		On Error Resume Next
		Tst = KeineLog
		On Error Goto 0
		If UCase( Tst ) = "JA" Then Exit Function

		On Error Resume Next
		TitelX = Titel
		' if not set 'Titel' outside procedure 'TitelX' is empty
		TitelX = title
		' if not set 'title' outside procedure 'TitelX' is empty

		If Len( TitelX ) < 2 Then TitelX = document.title
		' set title in .HTA
		If Len( TitelX ) < 2 Then TitelX = WScript.ScriptName
		' set title in .VBS
		On Error Goto 0

		On Error Resume Next
		MyLogFileX = MyLogFile
		' if not set 'MyLogFile' outside procedure, 'MyLogFileX' is empty
		If Len( MyLogFileX ) < 2    Then MyLogFileX = WScript.ScriptFullName & ".log"' .vbs
		If Len( MyLogFileX ) < 2    Then MyLogFileX = TitelX & ".log"        ' .hta
		On Error Goto 0

		' Enumerate Milliseconds
		Tst = Timer()               ' timer() in USA: 1234.22; dot separation
		Tst = Replace( Tst, "," , ".")        ' timer() in german: 23454,12; comma separation
		If InStr( Tst, "." ) = 0 Then Tst = Tst & ".000"
		Tst = Mid( Tst, InStr( Tst, "." ), 4 )
		If Len( Tst ) < 3 Then Tst = Tst & "0"

		' Enumerate Time Zone
		Dim AktDMTF : Set AktDMTF = CreateObject("WbemScripting.SWbemDateTime")
		AktDMTF.SetVarDate Now(), True : Tst = Tst & Mid( AktDMTF, 22 ) ' : MsgBox Tst, , "099 :: "
		' MsgBox "AktDMTF: '" & AktDMTF & "'", , "100 :: "
		Set AktDMTF = Nothing
		LogTemp = LogText
		LogTemp = "<![LOG[" & LogTemp & "]LOG]!>"
		LogTemp = LogTemp & "<"
		LogTemp = LogTemp & "time=""" & Hour( Time() ) & ":" & Minute( Time() ) & ":" & Second( Time() ) & Tst & """ "
		LogTemp = LogTemp & "date=""" & Month( Date() ) & "-" & Day( Date() ) & "-" & Year( Date() ) & """ "
		LogTemp = LogTemp & "component=""" & TitelX & """ "
		LogTemp = LogTemp & "context="""" "
		LogTemp = LogTemp & "type=""" & LogError & """ "
		LogTemp = LogTemp & "thread=""0"" "
		LogTemp = LogTemp & "file=""David.Segura"" "
		LogTemp = LogTemp & ">"

		Tst = 8							'ForAppending
		If LogError = 0 Then Tst = 2	'ForWriting

		Set FileOut = objFSO.OpenTextFile( MyLogFileX, Tst, True)
		If     LogTemp = vbCRLF Then FileOut.WriteLine ( LogTemp )
		If Not LogTemp = vbCRLF Then FileOut.WriteLine ( LogTemp )
		FileOut.Close
		Set FileOut	= Nothing
		'Set objFSO	= Nothing
	End Function
	' /////////////////////////////////////////////////////////
	' Trace Log Solid Line
	' /////////////////////////////////////////////////////////
	Sub LogLine
			TraceLog "=================================================================", 1
	End Sub
	' /////////////////////////////////////////////////////////
	' Trace Log Blank Space
	' /////////////////////////////////////////////////////////
	Sub LogSpace
		TraceLog "", 1
	End Sub
	' /////////////////////////////////////////////////////////
	' Trace Log Contents
	' /////////////////////////////////////////////////////////
	Sub LogStart
		'Tracelog "Start a new Log File", 0											'Clears any existing content
		'TraceLog "This is a standard line", 1										'Create an Entry
		'TraceLog "This is a warning line", 2										'Create an Entry and highlight yellow (Warning)
		'TraceLog "This is an error line", 3										'Create an Entry and highlight red (Error or Critical)
		'LogSpace																	'Create a Line without content
		'LogLine																	'Create a Line with =====================================

		If WScript.Arguments.length = 0 Then TraceLog "Starting "					& WScript.ScriptFullName, 0
		If WScript.Arguments.length <> 0 Then TraceLog "Starting "					& WScript.ScriptFullName, 2
		TraceLog "Start Date and Time is "											& Now, 1
		TraceLog "Script Last Modified: " 											& CreateObject("Scripting.FileSystemObject").GetFile(Wscript.ScriptFullName).DateLastModified, 1
		LogLine
		TraceLog "<Constant> Author: " 												& Author, 1
		TraceLog "<Constant> Author Email: " 										& AuthorEmail, 1
		TraceLog "<Constant> Company: " 											& Company, 1
		TraceLog "Do not contact the Author directly for Support", 3
		LogLine
		TraceLog "<Constant> Script: " 												& Script, 1
		TraceLog "<Constant> Description: " 										& Description, 1
		LogLine
		TraceLog "The defined Support process is to Submit an Incident", 3
		TraceLog "This script is provided for Testing Only (No Priority Support)", 3
		TraceLog "<Constant> SupportAction: " 										& SupportAction, 2
		TraceLog "<Constant> Incident Area: " 										& SupportArea, 2
		TraceLog "<Constant> Assign to Group: " 									& SupportGroup, 2
		TraceLog "<Constant> Assignee: " 											& SupportContact, 2
		TraceLog "<Constant> Subject: " 											& SupportSubject, 2
		TraceLog "<Constant> Description: " 										& SupportProblem, 2
		LogLine
		TraceLog "<Constant> Title: " 												& Title, 1
		TraceLog "<Constant> Version: " 											& Version, 2
		TraceLog "<Constant> VersionFull: " 										& VersionFull, 2
		LogLine
		TraceLog "<Variable> MyUserName: " 											& MyUserName, 1
		TraceLog "<Variable> MyComputerName: " 										& MyComputerName, 1
		TraceLog "<Variable> MyWindir: " 											& MyWindir, 1
		TraceLog "<Variable> MyTemp: " 												& MyTemp, 1
		TraceLog "<Variable> MySystemDrive: " 										& MySystemDrive, 1
		TraceLog "<Variable> MyArchitecture: " 										& MyArchitecture, 1
		LogLine
		TraceLog "<Variable> MyScriptFullPath = "									& MyScriptFullPath,			1
		TraceLog "<Variable> MyScriptFileName = "									& MyScriptFileName, 		1
		TraceLog "<Variable> MyScriptBaseName = "									& MyScriptBaseName, 		1
		TraceLog "<Variable> MyScriptParentFolder = "								& MyScriptParentFolder, 	1
		TraceLog "<Variable> MyScriptGParentFolder = "								& MyScriptGParentFolder,	1
		TraceLog "<Variable> MyParentFolderName = "									& MyParentFolderName,		1
	End Sub
'==============================================================================================
'==============================================================================================
	Function BuildLogScript1
		If WScript.Arguments.length = 0 Then TraceLog "Starting " & WScript.ScriptFullName, 0
		If WScript.Arguments.length <> 0 Then TraceLog "Starting " & WScript.ScriptFullName, 1
		TraceLog "Script Last Modified: " & CreateObject("Scripting.FileSystemObject").GetFile(MyScriptFullPath).DateLastModified, 1
		LogSpace
	End Function
'==============================================================================================
'==============================================================================================
	' /////////////////////////////////////////////////////////
	' Check if we have Admin Rights
	' /////////////////////////////////////////////////////////
	'	Usage:	If IsAdmin = False Then Wscript.Quit
	'	Result:	Script will exit
	'
	'	Usage:	If IsAdmin = False Then DoElevate
	'	Result:	Script will run the DoElevate Subroutine
	
	Function IsAdmin
		'LogLine
		'TraceLog "Function IsAdmin", 1
		
		Dim RegKey
		IsAdmin = False
		On Error Resume Next
		
		'Try to read a Registry Key that is only readable with Admin Rights
		RegKey = CreateObject("WScript.Shell").RegRead("HKEY_USERS\S-1-5-19\")
		If Err.Number = 0 Then IsAdmin = True
		
		'Log Result
		If IsAdmin = True Then TraceLog "<IsAdmin = True> User has Admin Rights", 1
		If IsAdmin = False Then TraceLog "<IsAdmin = False> User does not have Admin Rights", 1
	End Function
'==============================================================================================
'==============================================================================================
	' /////////////////////////////////////////////////////////
	' Check if we are running under SYSTEM Account
	' /////////////////////////////////////////////////////////
	Function IsSystem
		'LogLine
		'TraceLog "Function IsSystem", 1
		
		IsSystem = False
		
		'Determine if we are running this under the System Account and LOG result
		If Lcase(CreateObject("WScript.Network").UserName) = "system" Then
			IsSystem = True
			TraceLog "<IsSystem = True> Script is being run under the SYSTEM context, possibly from SCCM or as a Scheduled Task", 2
		Else
			TraceLog "<IsSystem = False> Script is NOT being run under the SYSTEM context", 1
		End If
	End Function
'==============================================================================================
'==============================================================================================
	' /////////////////////////////////////////////////////////
	' Check for Command Line Arguments and Elevate if necessary
	' /////////////////////////////////////////////////////////
	Sub CheckArguments
		'LogLine
		TraceLog "Sub CheckArguments", 1

		Dim sArgument, sArguments
		Set sArguments = Wscript.Arguments
		
		If sArguments.Count = 0 Then
			TraceLog "Arguments have NOT been passed.  Exiting Sub", 1
			Exit Sub
		Else
			TraceLog "Arguments have been passed", 1
		End If
		
		For Each sArgument in sArguments
			TraceLog "<Variable> sArgument = " & sArgument, 1
			If Lcase(sArgument) = "uac"	Then sArgumentUAC = True
		Next
		
		'Find Named Arguments
		Dim colArgs
		Set colArgs = WScript.Arguments.Named
		
		If colArgs.Exists("OS") Then
			MyOperatingSystem = colArgs.Item("OS")
		Else
			MyOperatingSystem = "Unknown"
		End If
		
		TraceLog "<Variable> MyOperatingSystem = " & MyOperatingSystem, 1
	End Sub
'==============================================================================================
'==============================================================================================
	' /////////////////////////////////////////////////////////
	' Check Operating System Properties
	' /////////////////////////////////////////////////////////
	Sub GetMyOperatingSystem
		'LogLine
		TraceLog "Sub GetMyOperatingSystem",1

		Dim objItem, colItems
		Dim Unsupported
		Dim tempMyOperatingSystem
		
		tempMyOperatingSystem = "Unknown"
		
		On Error Resume Next
		Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
		For Each objItem In colItems
			tempMyOperatingSystem = objItem.Caption
		Next
		
		If tempMyOperatingSystem = "Unknown" Then
			TraceLog "Cannot determine OS by WMI", 2
			Exit Sub
		Else
			MyOperatingSystem = tempMyOperatingSystem
		End If

		For Each objItem In colItems
			TraceLog "<Property> Caption: " & objItem.Caption,1
			TraceLog "<Property> OperatingSystemSKU: " & objItem.OperatingSystemSKU,1
			TraceLog "<Property> Organization: " & objItem.Organization,1
			TraceLog "<Property> OSArchitecture: " & objItem.OSArchitecture,1
			TraceLog "<Property> OSProductSuite: " & objItem.OSProductSuite,1
			TraceLog "<Property> OSType: " & objItem.OSType,1
			TraceLog "<Property> ProductType: " & objItem.ProductType,1
			TraceLog "<Property> RegisteredUser: " & objItem.RegisteredUser,1
			TraceLog "<Property> SerialNumber: " & objItem.SerialNumber,1
			TraceLog "<Property> Status: " & objItem.Status,1
			TraceLog "<Property> SuiteMask: " & objItem.SuiteMask,1
			TraceLog "<Property> Version: " & objItem.Version,1

			With objItem
			Select Case True	
				'Client Operating Systems
				Case Left(.Version,3) = "5.1" and .ProductType = 1
					MyOperatingSystem = "Windows XP"
				Case Left(.Version,3) = "5.2" and .ProductType = 1
					MyOperatingSystem = "Windows XP"
				Case Left(.Version,3) = "6.0" and .ProductType = 1
					MyOperatingSystem = "Windows Vista"
				Case Left(.Version,3) = "6.1" and .ProductType = 1
					MyOperatingSystem = "Windows 7"
				Case Left(.Version,3) = "6.2" and .ProductType = 1
					MyOperatingSystem = "Windows 8"
				Case Left(.Version,3) = "6.3" and .ProductType = 1
					MyOperatingSystem = "Windows 8.1"
				Case Left(.Version,3) = "10." and .ProductType = 1
					MyOperatingSystem = "Windows 10"
				'Server Operating Systems
				Case Left(.Version,3) = "5.2" and .ProductType > 1
					MyOperatingSystem = "Windows Server 2003"
				Case Left(.Version,3) = "6.0" and .ProductType > 1
					MyOperatingSystem = "Windows Server 2008"
				Case Left(.Version,3) = "6.1" and .ProductType > 1
					MyOperatingSystem = "Windows Server 2008 R2"
				Case Left(.Version,3) = "6.2" and .ProductType > 1
					MyOperatingSystem = "Windows Server 2012"
				Case Left(.Version,3) = "6.3" and .ProductType > 1
					MyOperatingSystem = "Windows Server 2012 R2"
				Case Left(.Version,3) = "10." and .ProductType > 1
					MyOperatingSystem = "Windows Server 10"
				Case Else
					MyOperatingSystem = "Unknown"
				End Select
			End With

			If MyOperatingSystem = "" or Unsupported = True Then
				MyOperatingSystem = objItem.Caption
				TraceLog "<Property> MyOperatingSystem = " & MyOperatingSystem, 3
				TraceLog MyOperatingSystem & " is not supported by this Script", 3
				Wscript.Quit
			Else
				TraceLog "<Variable> MyOperatingSystem = " & MyOperatingSystem, 2
				TraceLog "<Variable> MyArchitecture = " & MyArchitecture, 2
			End If
		Next
	End Sub
'==============================================================================================
'==============================================================================================
	' /////////////////////////////////////////////////////////
	' Relaunch Elevated
	' /////////////////////////////////////////////////////////
	Sub Elevation	
		'If UAC was in the Arguments, we have already launched a second time for Elevation
		If sArgumentUAC = True Then
			TraceLog "Script is running under UAC, no need to relaunch Elevated", 3
			Exit Sub
		End If
		
		'If running in the System Context, we do not need to Elevate
		If IsSystem = True Then
			TraceLog "Script is running under SYSTEM, no need to relaunch Elevated", 3
			Exit Sub
		End If
		
		If MyOperatingSystem = "Windows XP" Then
			TraceLog "Running Windows XP, no need to relaunch Elevated", 3
			Exit Sub
		End If
		
		'Relaunch Elevated
		TraceLog "Relaunching Elevated", 3
		LogLine
		LogLine
		LogLine
		LogLine
		LogLine
		TraceLog "Runing Command: cscript.exe " & Chr(34) & WScript.ScriptFullName & Chr(34) & " /OS:""" & MyOperatingSystem & """ uac", 1
		objShellApp.ShellExecute "cscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " /OS:""" & MyOperatingSystem & """ uac", "", "runas", 0
		WScript.Quit
	End Sub
'==============================================================================================
'==============================================================================================
Function ReadFile(filename)
	Dim bom, f, stream
	
	bom = ""
	Set f = objFSO.OpenTextFile(filename)
	Do Until f.AtEndOfStream Or bom = "ÿþ" Or bom = "þÿ" Or Len(bom) >= 3
		bom = bom & f.Read(1)
	Loop
	f.Close

	Select Case bom
		Case "ï»¿"       'UTF-8 text
			'Wscript.Echo "Reg File encoded UTF-8"
			Set stream = CreateObject("ADODB.Stream")
			stream.Open
			stream.Type = 2
			stream.Charset = "utf-8"
			stream.LoadFromFile filename
			ReadFile = stream.ReadText
			stream.Close
		Case "ÿþ", "þÿ" 'UTF-16 text
			'Wscript.Echo "Reg File encoded UTF-16"
			Set f = objFSO.OpenTextFile(filename, 1, False, -1)
			ReadFile = f.ReadAll
			f.Close
		'Case Else        'ASCII text
		'	'Wscript.Echo "Reg File encoded ASCII"
		'	Set f = objFSO.OpenTextFile(filename, 1, False, 0)
		'	ReadFile = f.ReadAll
		'	f.Close
		Case Else		'Should be Unicode, so default to Unicode anyway
			'Wscript.Echo "Reg File encoded UTF-16"
			Set f = objFSO.OpenTextFile(filename, 1, False, -1)
			ReadFile = f.ReadAll
			f.Close
	End Select
End Function
'==============================================================================================
'==============================================================================================