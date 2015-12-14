#https://winpeguy.wordpress.com/my-scripts/tool-osconfig/
#http://ccmexec.com/2015/08/removing-built-in-apps-from-windows-10-using-powershell/

$AppsList = "Microsoft.3DBuilder",
			#"Microsoft.Appconnector",
			#"Microsoft.BingFinance",
			#"Microsoft.BingNews",
			#"Microsoft.BingSports",
			#"Microsoft.BingWeather",
			#"Microsoft.CommsPhone",
			#"Microsoft.ConnectivityStore",
			#"Microsoft.Getstarted",
			#"Microsoft.Messaging",
			"Microsoft.MicrosoftOfficeHub",
			#"Microsoft.MicrosoftSolitaireCollection",
			#"Microsoft.Office.OneNote",
			#"Microsoft.Office.Sway",
			#"Microsoft.People",
			"Microsoft.SkypeApp",
			#"Microsoft.VCLibs.120.00",
			#"Microsoft.VCLibs.140.00",
			#"Microsoft.Windows.Photos",
			#"Microsoft.WindowsAlarms",
			#"Microsoft.WindowsCalculator",
			#"Microsoft.WindowsCamera",
			#"Microsoft.WindowsCommunicationsApps",
			#"Microsoft.WindowsMaps",
			#"Microsoft.WindowsPhone",
			#"Microsoft.WindowsSoundRecorder",
			#"Microsoft.WindowsStore",
			#"Microsoft.XboxApp",
			#"Microsoft.ZuneMusic",
			#"Microsoft.ZuneVideo"
						
ForEach ($App in $AppsList)
	{
		$PackageFullName = (Get-AppxPackage $App).PackageFullName
		$ProPackageFullName = (Get-AppxProvisionedPackage -online | where {$_.Displayname -eq $App}).PackageName
		write-host $PackageFullName
		Write-Host $ProPackageFullName
		if ($PackageFullName)
			{
				Write-Host "Removing Package: $App"
				Remove-AppxPackage -package $PackageFullName
			}
		else
			{
				Write-Host "Unable to find package: $App"
			}
		if ($ProPackageFullName)
			{
				Write-Host "Removing Provisioned Package: $ProPackageFullName"
				Remove-AppxProvisionedPackage -online -packagename $ProPackageFullName
			}
		else
			{
				Write-Host "Unable to find provisioned package: $App"
			}
	}