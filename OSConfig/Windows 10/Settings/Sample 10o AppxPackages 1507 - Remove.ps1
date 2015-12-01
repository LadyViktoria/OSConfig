#http://ccmexec.com/2015/08/removing-built-in-apps-from-windows-10-using-powershell/

$AppsList = "Microsoft.3DBuilder",
			"Microsoft.BingFinance",
			"Microsoft.BingNews",
			"Microsoft.BingSports",
			"Microsoft.BingWeather",
			"Microsoft.MicrosoftOfficeHub",
			"Microsoft.Office.OneNote",
			"Microsoft.MicrosoftSolitaireCollection",
			"Microsoft.People",
			"Microsoft.SkypeApp",
			"Microsoft.WindowsCommunicationsApps",
			"Microsoft.WindowsPhone",
			"Microsoft.Windows.Photos",
			"Microsoft.XboxApp",
			"Microsoft.ZuneMusic",
			"Microsoft.ZuneVideo"

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